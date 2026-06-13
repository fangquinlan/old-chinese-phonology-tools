/**
 * 文本处理模块
 * 负责从文本中提取汉字读音与字典释义
 */

const HAN_REGEX_SOURCE = '\\p{Script=Han}';
const HAN_REGEX_FALLBACK = [
    '[\\u3400-\\u4DBF\\u4E00-\\u9FFF\\uF900-\\uFAFF]',
    '[\\uD840-\\uD868][\\uDC00-\\uDFFF]',
    '\\uD869[\\uDC00-\\uDEDF]',
    '\\uD869[\\uDF00-\\uDFFF]',
    '[\\uD86A-\\uD87A][\\uDC00-\\uDFFF]',
    '\\uD87B[\\uDC00-\\uDE5F]',
    '\\uD87E[\\uDC00-\\uDE1F]',
    '[\\uD880-\\uD88C][\\uDC00-\\uDFFF]',
    '\\uD88D[\\uDC00-\\uDC7F]'
].join('|');

function createHanRegex(flags) {
    try {
        return new RegExp(HAN_REGEX_SOURCE, flags);
    } catch (error) {
        return new RegExp(HAN_REGEX_FALLBACK, flags.replace('u', ''));
    }
}

const HAN_CHAR_REGEX_GLOBAL = createHanRegex('gu');
const HAN_CHAR_REGEX_SINGLE = createHanRegex('u');
const OLD_CHINESE_WORKBOOK_FILE = '上古汉语音节表.xlsx';
const OLD_CHINESE_DICTIONARY_SHEET = '字典表';
const CHEN_INDEX_WORKBOOK_FILE = '陳靖《兩周古文字編注》索引.xlsx';
const CHEN_INDEX_MAIN_SHEET = '正表';
const PHONETIC_DOMAIN_SOURCES = {
    oldChinese: {
        value: 'old-chinese',
        label: '上古汉语音节表'
    },
    chenIndex: {
        value: 'chen-index',
        label: '陳靖《兩周古文字編注》索引'
    },
    merged: {
        value: 'merged',
        label: '两表合并'
    }
};

let textProcessResults = [];
let glossExportResults = [];
let glossDictionaryMap = null;
let glossDictionaryPromise = null;
let dictionaryRows = null;
let dictionaryRowsPromise = null;
let chenIndexRows = null;
let chenIndexRowsPromise = null;
let oldChineseWorkbookGlossMap = null;
let oldChineseWorkbookGlossMapPromise = null;
let phoneticDomainExportRows = [];
let phoneticDomainExportMeta = null;
let idsMap = null;
let idsMapPromise = null;
const MANUAL_PHONETIC_OVERRIDES_STORAGE_KEY = 'oldChineseManualPhoneticOverrides';
let manualPhoneticOverrides = loadManualPhoneticOverrides();

function escapeHtmlText(value) {
    return (value ?? '').toString()
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function normalizeCellText(value) {
    return (value ?? '').toString().replace(/\r\n?/g, '\n').trim();
}

function formatCellHtml(text, fallback = '—') {
    const normalized = normalizeCellText(text);
    return escapeHtmlText(normalized || fallback).replace(/\n/g, '<br>');
}

function getLocalDateStamp() {
    const now = new Date();
    const year = now.getFullYear();
    const month = String(now.getMonth() + 1).padStart(2, '0');
    const day = String(now.getDate()).padStart(2, '0');
    return `${year}-${month}-${day}`;
}

function loadManualPhoneticOverrides() {
    try {
        const saved = JSON.parse(localStorage.getItem(MANUAL_PHONETIC_OVERRIDES_STORAGE_KEY) || '{}');
        return new Map(Object.entries(saved).filter(([, sound]) => typeof sound === 'string' && sound.trim()));
    } catch (error) {
        console.warn('手动补音读取失败，将使用空白状态。', error);
        return new Map();
    }
}

function persistManualPhoneticOverrides() {
    try {
        if (manualPhoneticOverrides.size === 0) {
            localStorage.removeItem(MANUAL_PHONETIC_OVERRIDES_STORAGE_KEY);
            return;
        }

        localStorage.setItem(
            MANUAL_PHONETIC_OVERRIDES_STORAGE_KEY,
            JSON.stringify(Object.fromEntries(manualPhoneticOverrides))
        );
    } catch (error) {
        console.warn('手动补音保存失败。', error);
    }
}

function buildCsv(headers, rows) {
    const toCell = value => `"${(value ?? '').toString().replace(/"/g, '""')}"`;
    return [
        headers.map(toCell).join(','),
        ...rows.map(row => row.map(toCell).join(','))
    ].join('\n');
}

function downloadCsv(csv, filename) {
    const blob = new Blob([`\ufeff${csv}`], { type: 'text/csv;charset=utf-8;' });
    const link = document.createElement('a');
    link.href = URL.createObjectURL(blob);
    link.download = filename;
    link.click();
    setTimeout(() => URL.revokeObjectURL(link.href), 0);
}

function sanitizeFilenamePart(value, fallback = '未命名') {
    const text = normalizeCellText(value).replace(/[\\/:*?"<>|]/g, '_');
    return (text || fallback).slice(0, 40);
}

function sanitizeSheetName(value, fallback = '同諧聲域') {
    const text = normalizeCellText(value).replace(/[\\/?*\[\]:]/g, '_');
    return (text || fallback).slice(0, 31);
}

function downloadWorkbook(headers, rows, filename, sheetName) {
    if (!window.XLSX) {
        const csvFilename = filename.replace(/\.xlsx$/i, '.csv');
        downloadCsv(buildCsv(headers, rows), csvFilename);
        return csvFilename;
    }

    const worksheet = window.XLSX.utils.aoa_to_sheet([headers, ...rows]);
    const defaultWidths = [16, 18, 10, 32, 26, 58, 72];
    worksheet['!cols'] = headers.map((_, index) => ({ wch: defaultWidths[index] || 28 }));

    const workbook = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(workbook, worksheet, sanitizeSheetName(sheetName));
    window.XLSX.writeFile(workbook, filename);
    return filename;
}

function parseIdsLine(line) {
    const cleanedLine = line.replace(/^\uFEFF/, '').trimEnd();
    if (!cleanedLine) return null;

    const columns = cleanedLine.split('\t');
    const char = normalizeCellText(columns.shift());
    if (!char) return null;

    return {
        char,
        ids: columns.map(normalizeCellText).filter(Boolean).join('；')
    };
}

async function loadIdsMap() {
    if (idsMap) {
        return idsMap;
    }

    if (idsMapPromise) {
        return idsMapPromise;
    }

    idsMapPromise = (async () => {
        const response = await fetch('ids_lv0.txt', { cache: 'no-store' });
        if (!response.ok) {
            throw new Error(`IDS request failed with status ${response.status}.`);
        }

        const text = await response.text();
        const map = new Map();

        text.split(/\r?\n/).forEach(line => {
            const entry = parseIdsLine(line);
            if (entry) {
                map.set(entry.char, entry.ids);
            }
        });

        idsMap = map;
        return map;
    })().catch(error => {
        idsMapPromise = null;
        throw error;
    });

    return idsMapPromise;
}

function getExtractedCharStats(text) {
    const chineseChars = extractChineseChars(text);
    return {
        chineseChars,
        uniqueChars: [...new Set(chineseChars)]
    };
}

/**
 * 处理文本：去标点、去重、查询读音
 */
function processText() {
    const text = document.getElementById('textInput').value.trim();

    if (!text) {
        alert('请输入文本！');
        return;
    }

    const { chineseChars, uniqueChars } = getExtractedCharStats(text);

    if (chineseChars.length === 0) {
        alert('未找到汉字，请检查输入！');
        return;
    }

    const results = queryPronunciations(uniqueChars);
    textProcessResults = results;

    displayTextProcessResults(results, chineseChars.length, uniqueChars.length);
    document.getElementById('exportTextBtn').disabled = false;
}

/**
 * 提取汉字（去除标点和非汉字字符）
 * @param {string} text - 输入文本
 * @returns {Array} 汉字数组
 */
function extractChineseChars(text) {
    return text.match(HAN_CHAR_REGEX_GLOBAL) || [];
}

/**
 * 查询每个字的所有读音
 * @param {Array} chars - 汉字数组
 * @returns {Array} 查询结果数组
 */
function queryPronunciations(chars) {
    return chars.map(char => {
        const charData = window.app.allData.filter(d => d['字'] === char);

        if (!charData.length) {
            return {
                char,
                pinyin: 'N/A',
                pronunciations: [],
                pronunciationCount: 0,
                notFound: true
            };
        }

        const pronunciationMap = new Map();

        charData.forEach(item => {
            const sound = item['音'] || '';
            if (!sound || pronunciationMap.has(sound)) return;

            pronunciationMap.set(sound, {
                sound,
                meaning: item['釋義'] || '',
                shijing: item['見詩經韻'] === '√',
                xizhou: parseFloat(item['見西周']) > 0,
                preQinFreq: parseFloat(item['先秦字頻（歸一化）']) || 0
            });
        });

        const pronunciations = Array.from(pronunciationMap.values());

        return {
            char,
            pinyin: charData[0]['拼音'] || 'N/A',
            pronunciations,
            pronunciationCount: pronunciations.length
        };
    }).sort((a, b) => a.char.localeCompare(b.char, 'zh'));
}

/**
 * 显示文本处理结果
 * @param {Array} results - 查询结果数组
 * @param {number} originalCount - 原始字数
 * @param {number} uniqueCount - 去重后字数
 */
function displayTextProcessResults(results, originalCount, uniqueCount) {
    const container = document.getElementById('textProcessResult');
    container.style.display = 'block';

    const foundCount = results.filter(item => !item.notFound).length;
    const notFoundCount = results.filter(item => item.notFound).length;
    const totalPronunciations = results.reduce((sum, item) => sum + item.pronunciationCount, 0);

    let html = '<div class="comparison-container">';

    html += '<h3>统计信息</h3>';
    html += '<div class="text-stats">';
    html += `<div class="stat-item"><div class="label">原文字数</div><div class="value">${originalCount}</div></div>`;
    html += `<div class="stat-item"><div class="label">去重后字数</div><div class="value">${uniqueCount}</div></div>`;
    html += `<div class="stat-item"><div class="label">数据库中找到</div><div class="value">${foundCount}</div></div>`;
    html += `<div class="stat-item"><div class="label">数据库中未找到</div><div class="value status-danger">${notFoundCount}</div></div>`;
    html += `<div class="stat-item"><div class="label">读音条目总数</div><div class="value">${totalPronunciations}</div></div>`;
    html += '</div>';

    html += '<h3 class="result-subtitle">汉字读音详情</h3>';
    html += '<div class="char-list">';

    results.forEach(item => {
        if (item.notFound) {
            html += `
                <div class="char-pronunciation-card char-pronunciation-card-muted">
                    <div class="char-header">
                        <div class="char">${escapeHtmlText(item.char)}</div>
                        <div class="char-pinyin">未找到</div>
                    </div>
                    <div class="not-found-note">数据库中无此字</div>
                </div>
            `;
            return;
        }

        html += `
            <div class="char-pronunciation-card">
                <div class="char-header">
                    <div class="char">${escapeHtmlText(item.char)}</div>
                    <div class="char-pinyin">${escapeHtmlText(item.pinyin)}</div>
                </div>
                <div class="pronunciation-list">
        `;

        item.pronunciations.forEach(pron => {
            html += `
                <div class="pronunciation-item">
                    <div class="sound">[${escapeHtmlText(pron.sound)}]</div>
                    <div class="meaning">${escapeHtmlText(truncateText(pron.meaning, 30))}</div>
                    <div>
                        ${pron.shijing ? '<span class="source-badge badge-shijing-small">诗经</span>' : ''}
                        ${pron.xizhou ? '<span class="source-badge badge-xizhou-small">西周</span>' : ''}
                    </div>
                </div>
            `;
        });

        html += `
                </div>
                <div class="pronunciation-count">
                    共 ${item.pronunciationCount} 个读音
                    <div class="pronunciation-sounds">
                        ${escapeHtmlText(item.pronunciations.map(pron => pron.sound).join(', '))}
                    </div>
                </div>
            </div>
        `;
    });

    html += '</div></div>';
    container.innerHTML = html;
}

/**
 * 截断文本
 * @param {string} text - 要截断的文本
 * @param {number} maxLength - 最大长度
 * @returns {string} 截断后的文本
 */
function truncateText(text, maxLength) {
    if (!text) return '';
    return text.length > maxLength ? `${text.substring(0, maxLength)}...` : text;
}

function pruneManualPhoneticOverrides(text) {
    const activeChars = new Set(extractChineseChars(text));
    let changed = false;

    manualPhoneticOverrides.forEach((_, char) => {
        if (!activeChars.has(char)) {
            manualPhoneticOverrides.delete(char);
            changed = true;
        }
    });

    if (changed) {
        persistManualPhoneticOverrides();
    }
}

function getCharSoundEntry(char) {
    const manualSound = manualPhoneticOverrides.get(char);
    if (manualSound) {
        return {
            sound: manualSound,
            missing: false,
            manual: true
        };
    }

    const charData = window.app.allData.filter(item => item['字'] === char);
    const sounds = [...new Set(charData.map(item => item['音']).filter(Boolean))];

    if (sounds.length > 0) {
        return {
            sound: sounds.join('/'),
            missing: false,
            manual: false
        };
    }

    return {
        sound: '□',
        missing: true,
        manual: false
    };
}

function buildPhoneticPairMarkup(entries) {
    const sourceHtml = entries.map(entry => {
        if (entry.type !== 'han') {
            return `<span class="phonetic-source-mark">${escapeHtmlText(entry.text)}</span>`;
        }

        const title = entry.missing
            ? '暂无读音，可在下方手动补音并固定'
            : `对应读音：${entry.sound}`;

        return `<span class="phonetic-source-char${entry.manual ? ' phonetic-source-char-manual' : ''}" data-pair-id="${entry.pairId}" data-char="${escapeHtmlText(entry.char)}" title="${escapeHtmlText(title)}">${escapeHtmlText(entry.char)}</span>`;
    }).join('');

    const outputHtml = entries.map(entry => {
        if (entry.type !== 'han') {
            return `<span class="phonetic-punctuation">${escapeHtmlText(entry.text)}</span>`;
        }

        const statusClass = [
            entry.missing ? 'phonetic-token-missing' : '',
            entry.manual ? 'phonetic-token-manual' : ''
        ].filter(Boolean).join(' ');
        const label = entry.missing
            ? `对应汉字 ${entry.char}，暂无读音，可手动补音`
            : `对应汉字 ${entry.char}`;

        return `<span class="phonetic-token ${statusClass}" data-pair-id="${entry.pairId}" data-char="${escapeHtmlText(entry.char)}" tabindex="0" role="button" aria-label="${escapeHtmlText(label)}">${escapeHtmlText(entry.sound)}</span>`;
    }).join('');

    return { sourceHtml, outputHtml };
}

function buildManualPhoneticEditor(entries) {
    const rows = [];
    const seen = new Set();

    entries.forEach(entry => {
        if (entry.type !== 'han' || (!entry.missing && !entry.manual) || seen.has(entry.char)) {
            return;
        }

        seen.add(entry.char);
        rows.push(entry);
    });

    if (rows.length === 0) {
        return '';
    }

    return `
        <div class="manual-phonetic-editor" id="manualPhoneticEditor">
            <div class="manual-phonetic-head">
                <h4>手动补音</h4>
                <p>给显示为方块的汉字填写读音后会固定到当前文本；当这个字从输入文本中删除时，固定值会自动移除。</p>
            </div>
            <div class="manual-phonetic-list">
                ${rows.map(entry => `
                    <div class="manual-phonetic-row${entry.manual ? ' is-fixed' : ''}" data-char="${escapeHtmlText(entry.char)}">
                        <div class="manual-phonetic-char">${escapeHtmlText(entry.char)}</div>
                        <input type="text" class="manual-phonetic-input" value="${entry.manual ? escapeHtmlText(entry.sound) : ''}"
                            placeholder="输入读音，如：kjaŋ" aria-label="给 ${escapeHtmlText(entry.char)} 手动补音">
                        <button type="button" class="btn-secondary manual-phonetic-save">${entry.manual ? '更新固定' : '固定'}</button>
                    </div>
                `).join('')}
            </div>
        </div>
    `;
}

function highlightPhoneticPairs(pairIds) {
    const container = document.getElementById('textProcessResult');
    if (!container) return;

    const activeIds = new Set(pairIds.filter(Boolean));

    container.querySelectorAll('[data-pair-id]').forEach(element => {
        element.classList.toggle('is-linked-highlight', activeIds.has(element.dataset.pairId));
    });
}

function getSelectedPhoneticPairIds(output) {
    const selection = window.getSelection();
    if (!selection || selection.isCollapsed || selection.rangeCount === 0) {
        return [];
    }

    if (!output.contains(selection.anchorNode) && !output.contains(selection.focusNode)) {
        return [];
    }

    const ranges = Array.from({ length: selection.rangeCount }, (_, index) => selection.getRangeAt(index));

    return Array.from(output.querySelectorAll('.phonetic-token'))
        .filter(token => ranges.some(range => {
            try {
                return range.intersectsNode(token);
            } catch (error) {
                return selection.containsNode(token, true);
            }
        }))
        .map(token => token.dataset.pairId);
}

function syncSelectedPhoneticTokens(output) {
    const selectedIds = getSelectedPhoneticPairIds(output);
    if (selectedIds.length > 0) {
        highlightPhoneticPairs(selectedIds);
    }
}

function focusManualPhoneticInput(char) {
    const editor = document.getElementById('manualPhoneticEditor');
    if (!editor) return;

    const row = Array.from(editor.querySelectorAll('.manual-phonetic-row'))
        .find(item => item.dataset.char === char);
    const input = row?.querySelector('.manual-phonetic-input');

    if (!row || !input) return;

    row.classList.add('is-editing');
    input.focus();
    input.select();

    window.setTimeout(() => row.classList.remove('is-editing'), 1200);
}

function saveManualPhoneticOverride(row) {
    const char = row?.dataset.char;
    const input = row?.querySelector('.manual-phonetic-input');
    const sound = input?.value.trim() || '';

    if (!char || !input) return;

    if (!sound) {
        input.focus();
        alert('请输入要固定的读音。');
        return;
    }

    manualPhoneticOverrides.set(char, sound);
    persistManualPhoneticOverrides();
    convertTextToPhonetic();
    focusManualPhoneticInput(char);
}

function bindManualPhoneticEditor() {
    const editor = document.getElementById('manualPhoneticEditor');
    if (!editor) return;

    editor.addEventListener('click', event => {
        const button = event.target.closest('.manual-phonetic-save');
        if (!button) return;

        saveManualPhoneticOverride(button.closest('.manual-phonetic-row'));
    });

    editor.addEventListener('keydown', event => {
        if (event.key !== 'Enter') return;

        const input = event.target.closest('.manual-phonetic-input');
        if (!input) return;

        event.preventDefault();
        saveManualPhoneticOverride(input.closest('.manual-phonetic-row'));
    });
}

function bindPhoneticHighlighting() {
    const output = document.getElementById('phoneticOutput');
    const source = document.getElementById('phoneticSourceText');
    if (!output || !source) return;

    output.addEventListener('click', event => {
        const selectedIds = getSelectedPhoneticPairIds(output);
        if (selectedIds.length > 0) {
            highlightPhoneticPairs(selectedIds);
            return;
        }

        const token = event.target.closest('.phonetic-token');
        if (token) {
            highlightPhoneticPairs([token.dataset.pairId]);
            if (token.classList.contains('phonetic-token-missing')) {
                focusManualPhoneticInput(token.dataset.char);
            }
        }
    });

    output.addEventListener('mouseup', () => {
        window.setTimeout(() => syncSelectedPhoneticTokens(output), 0);
    });

    output.addEventListener('keyup', event => {
        const token = event.target.closest?.('.phonetic-token');
        if ((event.key === 'Enter' || event.key === ' ') && token) {
            highlightPhoneticPairs([token.dataset.pairId]);
            return;
        }

        syncSelectedPhoneticTokens(output);
    });

    output.addEventListener('focusin', event => {
        const token = event.target.closest?.('.phonetic-token');
        if (token) {
            highlightPhoneticPairs([token.dataset.pairId]);
        }
    });

    source.addEventListener('click', event => {
        const char = event.target.closest('.phonetic-source-char');
        if (char) {
            highlightPhoneticPairs([char.dataset.pairId]);
            focusManualPhoneticInput(char.dataset.char);
        }
    });
}

/**
 * 清除文本处理结果
 */
function clearTextProcessing() {
    document.getElementById('textInput').value = '';
    document.getElementById('textProcessResult').style.display = 'none';
    document.getElementById('textProcessResult').innerHTML = '';
    document.getElementById('exportTextBtn').disabled = true;
    manualPhoneticOverrides.clear();
    persistManualPhoneticOverrides();
    textProcessResults = [];
}

/**
 * 导出文本处理结果
 */
async function exportTextResults() {
    if (textProcessResults.length === 0) {
        alert('没有可导出的数据！');
        return;
    }

    let charIdsMap;
    try {
        charIdsMap = await loadIdsMap();
    } catch (error) {
        console.error('IDS 数据加载失败:', error);
        alert('读取 ids_lv0.txt 失败，请确认文件可正常访问。');
        return;
    }

    const rows = textProcessResults.map(item => [
        item.char,
        item.notFound ? '未找到' : item.pronunciations.map(pron => pron.sound).join('/'),
        charIdsMap.get(item.char) || ''
    ]);

    downloadCsv(
        buildCsv(['汉字', '上古音', '字头IDS数据'], rows),
        `汉字读音提取结果_${getLocalDateStamp()}.csv`
    );
}

async function loadWorkbookSheetRows(filename, sheetName) {
    if (!window.XLSX) {
        throw new Error('XLSX parser is unavailable.');
    }

    const url = new URL(filename, window.location.href);
    url.searchParams.set('v', Date.now().toString());

    const response = await fetch(url.href, { cache: 'no-store' });
    if (!response.ok) {
        throw new Error(`Workbook request failed with status ${response.status}.`);
    }

    const workbookBuffer = await response.arrayBuffer();
    const workbook = window.XLSX.read(workbookBuffer, { type: 'array' });
    const sheet = workbook.Sheets[sheetName];

    if (!sheet) {
        throw new Error(`Workbook sheet "${sheetName}" was not found.`);
    }

    return window.XLSX.utils.sheet_to_json(sheet, { defval: '' });
}

async function loadDictionaryRowsFromWorkbook() {
    if (dictionaryRows) {
        return dictionaryRows;
    }

    if (dictionaryRowsPromise) {
        return dictionaryRowsPromise;
    }

    dictionaryRowsPromise = (async () => {
        dictionaryRows = await loadWorkbookSheetRows(OLD_CHINESE_WORKBOOK_FILE, OLD_CHINESE_DICTIONARY_SHEET);
        return dictionaryRows;
    })().catch(error => {
        dictionaryRowsPromise = null;
        throw error;
    });

    return dictionaryRowsPromise;
}

async function loadChenIndexRowsFromWorkbook() {
    if (chenIndexRows) {
        return chenIndexRows;
    }

    if (chenIndexRowsPromise) {
        return chenIndexRowsPromise;
    }

    chenIndexRowsPromise = (async () => {
        chenIndexRows = await loadWorkbookSheetRows(CHEN_INDEX_WORKBOOK_FILE, CHEN_INDEX_MAIN_SHEET);
        return chenIndexRows;
    })().catch(error => {
        chenIndexRowsPromise = null;
        throw error;
    });

    return chenIndexRowsPromise;
}

async function loadOldChineseWorkbookGlossMap() {
    if (oldChineseWorkbookGlossMap) {
        return oldChineseWorkbookGlossMap;
    }

    if (oldChineseWorkbookGlossMapPromise) {
        return oldChineseWorkbookGlossMapPromise;
    }

    oldChineseWorkbookGlossMapPromise = (async () => {
        const rows = await loadDictionaryRowsFromWorkbook();
        oldChineseWorkbookGlossMap = buildGlossDictionaryMap(rows);
        return oldChineseWorkbookGlossMap;
    })().catch(error => {
        oldChineseWorkbookGlossMapPromise = null;
        throw error;
    });

    return oldChineseWorkbookGlossMapPromise;
}

async function loadGlossDictionaryMap() {
    if (glossDictionaryMap) {
        return glossDictionaryMap;
    }

    if (glossDictionaryPromise) {
        return glossDictionaryPromise;
    }

    glossDictionaryPromise = (async () => {
        if (Array.isArray(window.app?.allData) && window.app.allData.length > 0) {
            glossDictionaryMap = buildGlossDictionaryMap(window.app.allData);
            return glossDictionaryMap;
        }

        const rows = await loadDictionaryRowsFromWorkbook();
        glossDictionaryMap = buildGlossDictionaryMap(rows);
        return glossDictionaryMap;
    })().catch(error => {
        glossDictionaryPromise = null;
        throw error;
    });

    return glossDictionaryPromise;
}

function buildGlossDictionaryMap(rows) {
    const dictionaryMap = new Map();

    rows.forEach(row => {
        const char = normalizeCellText(row['字']);
        if (!char) return;

        const entry = dictionaryMap.get(char) || {
            char,
            meanings: [],
            notes: []
        };

        const meaning = normalizeCellText(row['釋義']);
        const note = normalizeCellText(row['注釋']);

        if (meaning && !entry.meanings.includes(meaning)) {
            entry.meanings.push(meaning);
        }

        if (note && !entry.notes.includes(note)) {
            entry.notes.push(note);
        }

        dictionaryMap.set(char, entry);
    });

    return dictionaryMap;
}

function setGlossProcessingState(isProcessing) {
    const processButton = document.getElementById('processGlossBtn');
    const exportButton = document.getElementById('exportGlossBtn');

    if (processButton) {
        processButton.disabled = isProcessing;
        processButton.textContent = isProcessing ? '正在读取字典表...' : '提取释义注释';
    }

    if (exportButton && isProcessing) {
        exportButton.disabled = true;
    }
}

function buildGlossResult(char, entry) {
    if (!entry) {
        return {
            char,
            meaning: '未找到',
            note: '',
            notFound: true
        };
    }

    return {
        char,
        meaning: entry.meanings.join('\n\n'),
        note: entry.notes.join('\n\n'),
        notFound: false
    };
}

async function processGlossText() {
    const text = document.getElementById('glossTextInput').value.trim();
    const container = document.getElementById('glossResult');

    if (!text) {
        alert('请输入文本！');
        return;
    }

    const { chineseChars, uniqueChars } = getExtractedCharStats(text);

    if (chineseChars.length === 0) {
        alert('未找到汉字，请检查输入！');
        return;
    }

    container.style.display = 'block';
    container.innerHTML = '<div class="loading">正在读取《上古汉语音节表.xlsx》的字典表...</div>';
    setGlossProcessingState(true);

    try {
        const dictionaryMap = await loadGlossDictionaryMap();
        const results = uniqueChars.map(char => buildGlossResult(char, dictionaryMap.get(char)));

        glossExportResults = results;
        displayGlossResults(results, chineseChars.length, uniqueChars.length);
        document.getElementById('exportGlossBtn').disabled = false;
    } catch (error) {
        console.error('释义注释提取失败:', error);
        glossExportResults = [];
        container.innerHTML = '<div class="no-results">读取《上古汉语音节表.xlsx》失败，请确认文件可正常访问，且浏览器能加载 XLSX 解析脚本。</div>';
        alert('读取《上古汉语音节表.xlsx》失败，请稍后重试。');
    } finally {
        setGlossProcessingState(false);
    }
}

function displayGlossResults(results, originalCount, uniqueCount) {
    const container = document.getElementById('glossResult');
    container.style.display = 'block';

    const foundCount = results.filter(item => !item.notFound).length;
    const notFoundCount = results.filter(item => item.notFound).length;
    const notedCount = results.filter(item => normalizeCellText(item.note)).length;

    let html = '<div class="comparison-container">';
    html += '<h3>统计信息</h3>';
    html += '<div class="text-stats">';
    html += `<div class="stat-item"><div class="label">原文字数</div><div class="value">${originalCount}</div></div>`;
    html += `<div class="stat-item"><div class="label">去重后字数</div><div class="value">${uniqueCount}</div></div>`;
    html += `<div class="stat-item"><div class="label">字典表中找到</div><div class="value">${foundCount}</div></div>`;
    html += `<div class="stat-item"><div class="label">未找到</div><div class="value status-danger">${notFoundCount}</div></div>`;
    html += `<div class="stat-item"><div class="label">含注释条目</div><div class="value">${notedCount}</div></div>`;
    html += '</div>';

    html += '<h3 class="result-subtitle">字典表预览</h3>';
    html += `
        <div class="table-container glossary-table-container">
            <table class="glossary-table">
                <thead>
                    <tr>
                        <th>字头</th>
                        <th>释义</th>
                        <th>注释</th>
                    </tr>
                </thead>
                <tbody>
                    ${results.map(item => `
                        <tr class="${item.notFound ? 'glossary-row-muted' : ''}">
                            <td class="char-cell">${escapeHtmlText(item.char)}</td>
                            <td class="glossary-cell">${formatCellHtml(item.meaning, item.notFound ? '未找到' : '—')}</td>
                            <td class="glossary-cell">${formatCellHtml(item.note)}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;

    html += '</div>';
    container.innerHTML = html;
}

function clearGlossProcessing() {
    document.getElementById('glossTextInput').value = '';
    document.getElementById('glossResult').style.display = 'none';
    document.getElementById('glossResult').innerHTML = '';
    document.getElementById('exportGlossBtn').disabled = true;
    setGlossProcessingState(false);
    glossExportResults = [];
}

function exportGlossResults() {
    if (glossExportResults.length === 0) {
        alert('没有可导出的数据！');
        return;
    }

    const rows = glossExportResults.map(item => [
        item.char,
        item.meaning,
        item.note
    ]);

    downloadCsv(
        buildCsv(['字头', '释义', '注释'], rows),
        `字头释义注释_${getLocalDateStamp()}.csv`
    );
}

function setPhoneticDomainState(isProcessing) {
    const button = document.getElementById('extractPhoneticDomainBtn');

    if (button) {
        button.disabled = isProcessing;
        button.textContent = isProcessing ? '正在读取表格...' : '一键提取并导出 XLSX';
    }
}

function getSelectedPhoneticDomainSource() {
    const select = document.getElementById('phoneticDomainSourceMode');
    const selectedValue = select?.value || PHONETIC_DOMAIN_SOURCES.oldChinese.value;
    const source = Object.values(PHONETIC_DOMAIN_SOURCES).find(item => item.value === selectedValue);
    return source || PHONETIC_DOMAIN_SOURCES.oldChinese;
}

function getDictionaryRowChar(row) {
    return normalizeCellText(row['字'] ?? row['字頭'] ?? row['字头']);
}

function getDictionaryRowDomain(row) {
    return normalizeCellText(row['諧聲域'] ?? row['谐声域']);
}

function getChenIndexRowChar(row) {
    return normalizeCellText(row['字頭'] ?? row['字头'] ?? row['字']);
}

function getChenIndexRowDomain(row) {
    return normalizeCellText(row['諧聲域'] ?? row['谐声域']);
}

function getPhoneticDomainSourcePlans(sourceMode) {
    const oldChinesePlan = {
        key: PHONETIC_DOMAIN_SOURCES.oldChinese.value,
        label: PHONETIC_DOMAIN_SOURCES.oldChinese.label,
        loadRows: loadDictionaryRowsFromWorkbook,
        getChar: getDictionaryRowChar,
        getDomain: getDictionaryRowDomain,
        buildEntry: buildOldChineseDomainEntry
    };
    const chenIndexPlan = {
        key: PHONETIC_DOMAIN_SOURCES.chenIndex.value,
        label: PHONETIC_DOMAIN_SOURCES.chenIndex.label,
        loadRows: loadChenIndexRowsFromWorkbook,
        getChar: getChenIndexRowChar,
        getDomain: getChenIndexRowDomain,
        buildEntry: buildChenIndexDomainEntry
    };

    if (sourceMode === PHONETIC_DOMAIN_SOURCES.chenIndex.value) {
        return [chenIndexPlan];
    }

    if (sourceMode === PHONETIC_DOMAIN_SOURCES.merged.value) {
        return [oldChinesePlan, chenIndexPlan];
    }

    return [oldChinesePlan];
}

function buildFieldNote(row, fields) {
    return fields
        .map(([label, key]) => {
            const value = normalizeCellText(row[key]);
            return value ? `${label}：${value}` : '';
        })
        .filter(Boolean)
        .join('\n');
}

function buildOldChineseDomainEntry(row, charIdsMap, plan) {
    const char = getDictionaryRowChar(row);
    return {
        source: plan.label,
        sourceKey: plan.key,
        domain: getDictionaryRowDomain(row),
        char,
        ids: charIdsMap.get(char) || '',
        sound: normalizeCellText(row['音']),
        meaning: normalizeCellText(row['釋義'] ?? row['释义']),
        note: normalizeCellText(row['注釋'] ?? row['注释'])
    };
}

function buildChenIndexDomainEntry(row, charIdsMap, plan, context = {}) {
    const char = getChenIndexRowChar(row);
    const glossEntry = context.oldChineseGlossMap?.get(char);
    const sound = [
        normalizeCellText(row['聲紐'] ?? row['声纽']),
        normalizeCellText(row['韻部'] ?? row['韵部'])
    ].filter(Boolean).join(' / ');

    return {
        source: plan.label,
        sourceKey: plan.key,
        domain: getChenIndexRowDomain(row),
        char,
        ids: charIdsMap.get(char) || '',
        sound,
        meaning: glossEntry ? glossEntry.meanings.join('\n\n') : '',
        note: buildFieldNote(row, [
            ['冊數', '册數'],
            ['頁碼', '頁碼'],
            ['聲首', '聲首'],
            ['聲符', '聲符'],
            ['非聲符', '非聲符'],
            ['備注', '備注'],
            ['備注', '備注_1']
        ])
    };
}

function buildPhoneticDomainSourceResult(sourceChar, rows, charIdsMap, plan, context = {}) {
    const sourceRows = rows.filter(row => plan.getChar(row) === sourceChar);
    const domains = [...new Set(sourceRows.map(plan.getDomain).filter(Boolean))];
    const domainSet = new Set(domains);

    const resultRows = rows
        .filter(row => domainSet.has(plan.getDomain(row)))
        .map(row => plan.buildEntry(row, charIdsMap, plan, context))
        .filter(item => item.char);

    return {
        source: plan.label,
        sourceKey: plan.key,
        sourceRowCount: sourceRows.length,
        domains,
        rows: resultRows
    };
}

function dedupeAndSortDomainRows(rows) {
    const seen = new Set();

    return rows
        .filter(item => {
            const key = [
                item.source,
                item.domain,
                item.char,
                item.ids,
                item.sound,
                item.meaning,
                item.note
            ].join('\u0001');

            if (seen.has(key)) return false;
            seen.add(key);
            return true;
        })
        .sort((a, b) => (
            a.source.localeCompare(b.source, 'zh') ||
            a.domain.localeCompare(b.domain, 'zh') ||
            a.char.localeCompare(b.char, 'zh') ||
            a.sound.localeCompare(b.sound, 'zh')
        ));
}

function buildPhoneticDomainResult(sourceChar, loadedSources, charIdsMap, sourceMode, context = {}) {
    const sourceResults = loadedSources.map(({ plan, rows }) => (
        buildPhoneticDomainSourceResult(sourceChar, rows, charIdsMap, plan, context)
    ));
    const resultRows = dedupeAndSortDomainRows(sourceResults.flatMap(result => result.rows));
    const modeInfo = Object.values(PHONETIC_DOMAIN_SOURCES)
        .find(source => source.value === sourceMode) || PHONETIC_DOMAIN_SOURCES.oldChinese;

    return {
        sourceChar,
        sourceMode,
        sourceLabel: modeInfo.label,
        sourceRowCount: sourceResults.reduce((sum, result) => sum + result.sourceRowCount, 0),
        sourceResults,
        domains: [...new Set(sourceResults.flatMap(result => result.domains))],
        rows: resultRows
    };
}

function getPhoneticDomainExportValues(rows) {
    return rows.map(item => [
        item.source,
        item.domain,
        item.char,
        item.ids,
        item.sound,
        item.meaning,
        item.note
    ]);
}

function getPhoneticDomainHeaders() {
    return ['来源', '諧聲域', '字头', 'IDS', '音', '释义', '注释'];
}

function getPhoneticDomainLabel(result) {
    return result.sourceResults
        .map(sourceResult => {
            const domainText = sourceResult.domains.length ? sourceResult.domains.join('、') : '未找到';
            return `${sourceResult.source}：${domainText}`;
        })
        .join('；');
}

function displayPhoneticDomainResults(result, filename = '') {
    const container = document.getElementById('phoneticDomainResult');
    container.style.display = 'block';

    if (result.sourceRowCount === 0) {
        container.innerHTML = `<div class="no-results">未在所选来源中找到「${escapeHtmlText(result.sourceChar)}」。</div>`;
        return;
    }

    if (result.domains.length === 0) {
        container.innerHTML = `<div class="no-results">「${escapeHtmlText(result.sourceChar)}」在所选来源中没有可用的「諧聲域」值。</div>`;
        return;
    }

    const distinctChars = new Set(result.rows.map(item => item.char)).size;
    const previewRows = result.rows.slice(0, 200);
    const domainLabel = getPhoneticDomainLabel(result);
    const matchedSourceCount = result.sourceResults.filter(item => item.domains.length > 0).length;

    let html = '<div class="comparison-container">';
    html += '<h3>同諧聲域导出结果</h3>';
    html += '<div class="text-stats">';
    html += `<div class="stat-item"><div class="label">输入字</div><div class="value">${escapeHtmlText(result.sourceChar)}</div></div>`;
    html += `<div class="stat-item"><div class="label">数据来源</div><div class="value">${escapeHtmlText(result.sourceLabel)}</div></div>`;
    html += `<div class="stat-item"><div class="label">諧聲域</div><div class="value">${escapeHtmlText(domainLabel)}</div></div>`;
    html += `<div class="stat-item"><div class="label">导出字条</div><div class="value">${result.rows.length}</div></div>`;
    html += `<div class="stat-item"><div class="label">不同字头</div><div class="value">${distinctChars}</div></div>`;
    html += `<div class="stat-item"><div class="label">匹配来源</div><div class="value">${matchedSourceCount}</div></div>`;
    html += '</div>';

    if (filename) {
        html += `<div class="field-hint">已导出：${escapeHtmlText(filename)}</div>`;
    }

    if (previewRows.length < result.rows.length) {
        html += `<div class="field-hint">下方预览前 ${previewRows.length} 条，导出文件包含全部 ${result.rows.length} 条。</div>`;
    }

    html += '<h3 class="result-subtitle">表格预览</h3>';
    html += `
        <div class="table-container glossary-table-container">
            <table class="glossary-table">
                <thead>
                    <tr>
                        <th>来源</th>
                        <th>諧聲域</th>
                        <th>字头</th>
                        <th>IDS</th>
                        <th>音</th>
                        <th>释义</th>
                        <th>注释</th>
                    </tr>
                </thead>
                <tbody>
                    ${previewRows.map(item => `
                        <tr>
                            <td>${formatCellHtml(item.source)}</td>
                            <td>${formatCellHtml(item.domain)}</td>
                            <td class="char-cell">${escapeHtmlText(item.char)}</td>
                            <td class="glossary-cell">${formatCellHtml(item.ids)}</td>
                            <td>${formatCellHtml(item.sound)}</td>
                            <td class="glossary-cell">${formatCellHtml(item.meaning)}</td>
                            <td class="glossary-cell">${formatCellHtml(item.note)}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
        </div>
    `;

    html += '</div>';
    container.innerHTML = html;
}

async function extractAndExportPhoneticDomain() {
    const input = document.getElementById('phoneticDomainCharInput').value.trim();
    const chars = extractChineseChars(input);
    const container = document.getElementById('phoneticDomainResult');
    const selectedSource = getSelectedPhoneticDomainSource();
    const sourcePlans = getPhoneticDomainSourcePlans(selectedSource.value);
    const needsOldChineseGloss = sourcePlans.some(plan => plan.key === PHONETIC_DOMAIN_SOURCES.chenIndex.value);

    if (chars.length === 0) {
        alert('请输入一个汉字！');
        return;
    }

    if (chars.length > 1) {
        alert('请只输入一个汉字。');
        return;
    }

    const sourceChar = chars[0];
    container.style.display = 'block';
    container.innerHTML = `<div class="loading">正在读取${escapeHtmlText(selectedSource.label)}与 IDS 数据...</div>`;
    setPhoneticDomainState(true);

    try {
        const [charIdsMap, oldChineseGlossMap, ...loadedRows] = await Promise.all([
            loadIdsMap(),
            needsOldChineseGloss ? loadOldChineseWorkbookGlossMap() : Promise.resolve(null)
        ].concat(sourcePlans.map(plan => plan.loadRows())));
        const loadedSources = sourcePlans.map((plan, index) => ({
            plan,
            rows: loadedRows[index]
        }));
        const result = buildPhoneticDomainResult(
            sourceChar,
            loadedSources,
            charIdsMap,
            selectedSource.value,
            { oldChineseGlossMap }
        );

        phoneticDomainExportRows = result.rows;
        phoneticDomainExportMeta = {
            sourceChar,
            sourceMode: selectedSource.value,
            sourceLabel: selectedSource.label,
            domains: result.domains,
            sourceResults: result.sourceResults.map(sourceResult => ({
                source: sourceResult.source,
                domains: sourceResult.domains
            }))
        };

        if (result.rows.length === 0) {
            displayPhoneticDomainResults(result);
            return;
        }

        const headers = getPhoneticDomainHeaders();
        const exportRows = getPhoneticDomainExportValues(result.rows);
        const domainFilenamePart = sanitizeFilenamePart(getPhoneticDomainLabel(result), '未知諧聲域');
        const filename = `同諧聲域_${sanitizeFilenamePart(sourceChar)}_${sanitizeFilenamePart(selectedSource.label)}_${domainFilenamePart}_${getLocalDateStamp()}.xlsx`;
        const exportedFilename = downloadWorkbook(headers, exportRows, filename, `同諧聲域_${sourceChar}`);

        window.__lastPhoneticDomainExport = {
            sourceChar,
            sourceMode: selectedSource.value,
            sourceLabel: selectedSource.label,
            domains: result.domains,
            sourceResults: phoneticDomainExportMeta.sourceResults,
            rowCount: result.rows.length,
            filename: exportedFilename
        };

        displayPhoneticDomainResults(result, exportedFilename);
    } catch (error) {
        console.error('同諧聲域导出失败:', error);
        phoneticDomainExportRows = [];
        phoneticDomainExportMeta = null;
        container.innerHTML = '<div class="no-results">读取表格或 IDS 数据失败，请确认文件可正常访问。</div>';
        alert('同諧聲域导出失败，请稍后重试。');
    } finally {
        setPhoneticDomainState(false);
    }
}

function clearPhoneticDomain() {
    document.getElementById('phoneticDomainCharInput').value = '';
    document.getElementById('phoneticDomainResult').style.display = 'none';
    document.getElementById('phoneticDomainResult').innerHTML = '';
    setPhoneticDomainState(false);
    phoneticDomainExportRows = [];
    phoneticDomainExportMeta = null;
}

/**
 * 转换为拟音文本
 * 逐字将字替换成上古音，保留标点
 */
function convertTextToPhonetic() {
    const text = document.getElementById('textInput').value;

    if (!text) {
        alert('请输入文本！');
        return;
    }

    pruneManualPhoneticOverrides(text);

    const container = document.getElementById('textProcessResult');
    container.style.display = 'block';

    let html = '<div class="comparison-container">';
    html += '<h3>拟音文本转换结果</h3>';

    let resultText = '';
    const entries = [];
    let pairIndex = 0;

    for (const char of text) {
        if (HAN_CHAR_REGEX_SINGLE.test(char)) {
            const soundEntry = getCharSoundEntry(char);

            resultText += `${soundEntry.sound} `;
            entries.push({
                type: 'han',
                char,
                sound: soundEntry.sound,
                missing: soundEntry.missing,
                manual: soundEntry.manual,
                pairId: `phonetic-pair-${pairIndex}`
            });
            pairIndex += 1;
        } else {
            resultText += char;
            entries.push({
                type: 'text',
                text: char
            });
        }
    }

    const { sourceHtml, outputHtml } = buildPhoneticPairMarkup(entries);
    const manualEditorHtml = buildManualPhoneticEditor(entries);

    html += `
        <div class="phonetic-pair-view">
            <div class="phonetic-source-panel">
                <div class="phonetic-panel-label">原文</div>
                <div id="phoneticSourceText" class="phonetic-source-text">${sourceHtml}</div>
            </div>
            <div class="phonetic-output-wrap">
                <div class="phonetic-panel-label">拟音文本</div>
                <div id="phoneticOutput" class="phonetic-output phonetic-output-display" tabindex="0"
                    data-plain-text="${escapeHtmlText(resultText)}">${outputHtml}</div>
            </div>
        </div>
        ${manualEditorHtml}
        <div class="phonetic-actions">
             <button type="button" class="btn-primary" onclick="window.TextProcessing.copyPhoneticText(this)">
                复制结果
            </button>
        </div>
    `;

    html += '</div>';
    container.innerHTML = html;
    bindPhoneticHighlighting();
    bindManualPhoneticEditor();
}

/**
 * 复制拟音文本
 */
async function copyPhoneticText(button) {
    const output = document.getElementById('phoneticOutput');
    if (!output) return;

    const text = output.value ?? output.dataset.plainText ?? output.textContent ?? '';

    try {
        if (navigator.clipboard?.writeText) {
            await navigator.clipboard.writeText(text);
        } else {
            const range = document.createRange();
            range.selectNodeContents(output);
            const selection = window.getSelection();
            selection.removeAllRanges();
            selection.addRange(range);
            document.execCommand('copy');
        }

        const originalText = button.innerHTML;
        button.innerHTML = '已复制';
        setTimeout(() => {
            button.innerHTML = originalText;
        }, 1800);
    } catch (error) {
        console.error('复制失败:', error);
        alert('复制失败，请手动复制。');
    }
}

window.TextProcessing = {
    processText,
    extractChineseChars,
    queryPronunciations,
    displayTextProcessResults,
    clearTextProcessing,
    exportTextResults,
    processGlossText,
    displayGlossResults,
    clearGlossProcessing,
    exportGlossResults,
    extractAndExportPhoneticDomain,
    clearPhoneticDomain,
    convertTextToPhonetic,
    copyPhoneticText
};

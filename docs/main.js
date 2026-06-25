/**
 * 主程序文件
 * 负责初始化和数据加载
 */

window.app = {
    allData: [],
    charIndex: new Map(),
    dataPromise: null,
    dataLoaded: false
};

async function loadData() {
    if (window.app.dataLoaded) {
        return window.app;
    }

    if (window.app.dataPromise) {
        return window.app.dataPromise;
    }

    window.app.dataPromise = (async () => {
    try {
        const dataResponse = await fetch('data/phon.json');
        if (!dataResponse.ok) {
            throw new Error(`Data request failed with status ${dataResponse.status}.`);
        }

        window.app.allData = await dataResponse.json();
        window.app.charIndex = window.app.allData.reduce((map, item) => {
            const char = item?.['字'];
            if (!char) return map;
            if (!map.has(char)) {
                map.set(char, []);
            }
            map.get(char).push(item);
            return map;
        }, new Map());
        window.app.dataLoaded = true;
        return window.app;
    } catch (error) {
        window.app.dataPromise = null;
        console.error('Error loading data:', error);
        alert('数据加载失败，请确认 `data/phon.json` 可正常访问。');
        throw error;
    }
    })();

    return window.app.dataPromise;
}

async function initializeApp() {
    if (window.UI?.initializeUI) {
        window.UI.initializeUI();
    }
}

window.app.ensureDataLoaded = loadData;
window.addEventListener('load', initializeApp);

/**
 * 主程序文件
 * 负责初始化和数据加载
 */

window.app = {
    allData: [],
    charIndex: new Map()
};

async function loadData() {
    try {
        const dataResponse = await fetch('上古音韵数据.json');

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
    } catch (error) {
        console.error('Error loading data:', error);
        alert('数据加载失败，请确认 `上古音韵数据.json` 可正常访问。');
    }
}

async function initializeApp() {
    if (window.UI?.initializeUI) {
        window.UI.initializeUI();
    }

    await loadData();
}

window.addEventListener('load', initializeApp);

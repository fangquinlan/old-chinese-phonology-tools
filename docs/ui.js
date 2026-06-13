/**
 * UI 初始化占位模块
 * 当前页面只保留文本处理工具，复杂字表 UI 已移除。
 */

function initializeUI() {
    if (window.__uiInitialized) return;
    window.__uiInitialized = true;
}

window.UI = {
    initializeUI
};

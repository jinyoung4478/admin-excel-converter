/**
 * 엑셀 변환 앱 - 탭 관리 및 컨버터 레지스트리
 */

import hyundaiConverter from './converters/hyundai.js?v=8';

// ========== 상수 ==========
const DOM_IDS = {
    TAB_LIST: 'tab-list',
    TAB_CONTENTS: 'tab-contents'
};

const CSS_CLASSES = {
    TAB_BTN: 'tab-btn',
    TAB_CONTENT: 'tab-content',
    ACTIVE: 'active'
};

const DATA_ATTRS = {
    TAB: 'tab',
    INITIALIZED: 'initialized'
};

// ========== 상태 ==========
const converters = [
    hyundaiConverter
    // 새 컨버터 추가시 여기에 임포트하여 추가
];

let activeTabId = null;

// ========== 탭 요소 생성 ==========

function createTabButton(converter, isActive) {
    const button = document.createElement('button');
    button.className = isActive
        ? `${CSS_CLASSES.TAB_BTN} ${CSS_CLASSES.ACTIVE}`
        : CSS_CLASSES.TAB_BTN;
    button.dataset[DATA_ATTRS.TAB] = converter.id;
    button.textContent = converter.name;
    button.addEventListener('click', () => switchTab(converter.id));
    return button;
}

function createTabContent(converter, isActive) {
    const content = document.createElement('div');
    content.className = isActive
        ? `${CSS_CLASSES.TAB_CONTENT} ${CSS_CLASSES.ACTIVE}`
        : CSS_CLASSES.TAB_CONTENT;
    content.id = buildTabContentId(converter.id);
    content.dataset[DATA_ATTRS.INITIALIZED] = 'false';
    return content;
}

function buildTabContentId(tabId) {
    return `tab-${tabId}`;
}

// ========== 탭 관리 ==========

function createTabs() {
    const tabList = document.getElementById(DOM_IDS.TAB_LIST);
    const tabContents = document.getElementById(DOM_IDS.TAB_CONTENTS);

    converters.forEach((converter, index) => {
        const isFirst = index === 0;
        tabList.appendChild(createTabButton(converter, isFirst));
        tabContents.appendChild(createTabContent(converter, isFirst));
    });

    if (converters.length > 0) {
        const firstConverter = converters[0];
        initializeTab(firstConverter.id);
        activeTabId = firstConverter.id;
    }
}

function switchTab(tabId) {
    if (activeTabId === tabId) return;

    updateTabButtonStates(tabId);
    updateTabContentVisibility(tabId);
    initializeTab(tabId);
    activeTabId = tabId;
}

function updateTabButtonStates(activeId) {
    document.querySelectorAll(`.${CSS_CLASSES.TAB_BTN}`).forEach(btn => {
        btn.classList.toggle(CSS_CLASSES.ACTIVE, btn.dataset[DATA_ATTRS.TAB] === activeId);
    });
}

function updateTabContentVisibility(activeId) {
    document.querySelectorAll(`.${CSS_CLASSES.TAB_CONTENT}`).forEach(content => {
        content.classList.toggle(CSS_CLASSES.ACTIVE, content.id === buildTabContentId(activeId));
    });
}

function initializeTab(tabId) {
    const tabContent = document.getElementById(buildTabContentId(tabId));
    if (!tabContent || tabContent.dataset[DATA_ATTRS.INITIALIZED] === 'true') return;

    const converter = converters.find(c => c.id === tabId);
    if (converter?.init) {
        converter.init(tabContent);
        tabContent.dataset[DATA_ATTRS.INITIALIZED] = 'true';
    }
}

// ========== 동적 컨버터 등록 ==========

function registerConverter(converter) {
    converters.push(converter);

    const tabList = document.getElementById(DOM_IDS.TAB_LIST);
    if (tabList) {
        appendConverterTab(converter);
    }
}

function appendConverterTab(converter) {
    const tabList = document.getElementById(DOM_IDS.TAB_LIST);
    const tabContents = document.getElementById(DOM_IDS.TAB_CONTENTS);

    tabList.appendChild(createTabButton(converter, false));
    tabContents.appendChild(createTabContent(converter, false));
}

// ========== 초기화 ==========

document.addEventListener('DOMContentLoaded', createTabs);

// ========== 외부 API ==========

window.ExcelConverterApp = {
    registerConverter,
    switchTab,
    getConverters: () => [...converters]
};

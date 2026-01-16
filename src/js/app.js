/**
 * 엑셀 변환 앱 - 탭 관리 및 컨버터 레지스트리
 */

// 컨버터 임포트
import hyundaiConverter from './converters/hyundai.js?v=7';

// 등록된 컨버터 목록
const converters = [
    hyundaiConverter,
    // 새 컨버터 추가시 여기에 임포트하여 추가
    // import newConverter from './converters/new.js';
];

// 현재 활성 탭
let activeTab = null;

// 탭 UI 생성
function createTabs() {
    const tabList = document.getElementById('tab-list');
    const tabContents = document.getElementById('tab-contents');

    converters.forEach((converter, index) => {
        // 탭 버튼 생성
        const tabBtn = document.createElement('button');
        tabBtn.className = 'tab-btn' + (index === 0 ? ' active' : '');
        tabBtn.dataset.tab = converter.id;
        tabBtn.textContent = converter.name;
        tabBtn.addEventListener('click', () => switchTab(converter.id));
        tabList.appendChild(tabBtn);

        // 탭 콘텐츠 영역 생성
        const tabContent = document.createElement('div');
        tabContent.className = 'tab-content' + (index === 0 ? ' active' : '');
        tabContent.id = `tab-${converter.id}`;
        tabContent.dataset.initialized = 'false';
        tabContents.appendChild(tabContent);
    });

    // 첫 번째 탭 초기화
    if (converters.length > 0) {
        initializeTab(converters[0].id);
        activeTab = converters[0].id;
    }
}

// 탭 전환
function switchTab(tabId) {
    if (activeTab === tabId) return;

    // 이전 탭 비활성화
    document.querySelectorAll('.tab-btn').forEach(btn => {
        btn.classList.toggle('active', btn.dataset.tab === tabId);
    });

    document.querySelectorAll('.tab-content').forEach(content => {
        content.classList.toggle('active', content.id === `tab-${tabId}`);
    });

    // 탭 초기화 (최초 접근시)
    initializeTab(tabId);
    activeTab = tabId;
}

// 탭 초기화
function initializeTab(tabId) {
    const tabContent = document.getElementById(`tab-${tabId}`);
    if (tabContent.dataset.initialized === 'true') return;

    const converter = converters.find(c => c.id === tabId);
    if (converter && converter.init) {
        converter.init(tabContent);
        tabContent.dataset.initialized = 'true';
    }
}

// 컨버터 등록 API (동적 추가용)
function registerConverter(converter) {
    converters.push(converter);
    // 이미 DOM이 로드되었으면 탭 추가
    if (document.getElementById('tab-list')) {
        addConverterTab(converter);
    }
}

// 단일 컨버터 탭 추가
function addConverterTab(converter) {
    const tabList = document.getElementById('tab-list');
    const tabContents = document.getElementById('tab-contents');

    const tabBtn = document.createElement('button');
    tabBtn.className = 'tab-btn';
    tabBtn.dataset.tab = converter.id;
    tabBtn.textContent = converter.name;
    tabBtn.addEventListener('click', () => switchTab(converter.id));
    tabList.appendChild(tabBtn);

    const tabContent = document.createElement('div');
    tabContent.className = 'tab-content';
    tabContent.id = `tab-${converter.id}`;
    tabContent.dataset.initialized = 'false';
    tabContents.appendChild(tabContent);
}

// 앱 초기화
document.addEventListener('DOMContentLoaded', () => {
    createTabs();
});

// 외부 API 노출
window.ExcelConverterApp = {
    registerConverter,
    switchTab,
    getConverters: () => [...converters]
};

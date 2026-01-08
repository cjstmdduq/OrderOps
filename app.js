// ========== ES6 모듈 Import ==========
import {
  PRODUCT_TYPE_MAP,
  PUZZLE_PRODUCT_IDS,
  ROLL_PRODUCT_IDS,
  TAPE_PRODUCT_IDS,
  PACKAGING_THRESHOLDS,
  PUZZLE_BOX_CAPACITY,
  CUTTING_KEYWORDS,
  FINISHING_KEYWORDS,
  COLUMN_TOOLTIPS
} from './js/constants.js';

import {
  parseCSVToRows,
  parseOrderData,
  parseOptionInfo,
  parseProductName,
  parseThickness,
  parseWidth,
  parseLengthToMeters,
  parsePrice,
  extractZipCode,
  detectCuttingRequest,
  detectFinishingRequest,
  getProductTypeByName
} from './js/parsers.js';

import {
  groupOrdersByRecipient,
  determineRollPackaging,
  calculatePuzzleBoxes,
  applySmartCutting,
  processPacking,
  collectWarningsFromItems,
  collectDeliveryMemos
} from './js/packing.js';

import {
  loadShippingFees as loadShippingFeesAsync,
  loadProductDb as loadProductDbAsync,
  calculateShippingFee,
  matchDesignCodeFromDb,
  getProductDbMatch,
  generateDesignCode as generateDesignCodeFromDb
} from './js/dataLoaders.js';

import {
  formatCurrency,
  splitAddress,
  normalizePhone,
  normalizeAddress,
  makeCustomerKey,
  getDimensions,
  formatLengthToMeters,
  getProductBadgeClass,
  getStatusClass,
  downloadExcel
} from './js/utils.js';

// ========== 전역 상태 ==========
let ordersData = [];
let filteredOrders = [];
let groupedOrders = {};
let packedOrders = [];
let rawOrderData = [];
let shippingFees = [];
let productDb = [];
let currentTabId = 'raw';

// ========== DOM 요소 (DOMContentLoaded에서 초기화) ==========
let uploadArea, fileInput, uploadBtn, fileInfo, fileName, clearBtn;
let summarySection, filterSection, ordersSection, ordersTableBody;
let kyungdongCount, logenCount, totalCustomers, giftOrders;
let productTypeFilter, orderStatusFilter, giftFilter, searchInput, selectAll;
let orderModal, modalClose, modalBody;
let sidebarItems, shippingFeesPage, shippingFeesTableBody;
let tabBtns, contentSections;

// ========== 헬퍼 함수 (모듈 래퍼) ==========
function generateDesignCode(order) {
  return generateDesignCodeFromDb(order, productDb);
}

// ========== 초기화 ==========
document.addEventListener('DOMContentLoaded', async () => {
  // DOM 요소 초기화
  uploadArea = document.getElementById('uploadArea');
  fileInput = document.getElementById('fileInput');
  uploadBtn = document.getElementById('uploadBtn');
  fileInfo = document.getElementById('fileInfo');
  fileName = document.getElementById('fileName');
  clearBtn = document.getElementById('clearBtn');

  summarySection = document.getElementById('summarySection');
  filterSection = document.getElementById('filterSection');
  ordersSection = document.getElementById('ordersSection');
  ordersTableBody = document.getElementById('ordersTableBody');

  kyungdongCount = document.getElementById('kyungdongCount');
  logenCount = document.getElementById('logenCount');
  totalCustomers = document.getElementById('totalCustomers');
  giftOrders = document.getElementById('giftOrders');

  productTypeFilter = document.getElementById('productTypeFilter');
  orderStatusFilter = document.getElementById('orderStatusFilter');
  giftFilter = document.getElementById('giftFilter');
  searchInput = document.getElementById('searchInput');
  selectAll = document.getElementById('selectAll');

  orderModal = document.getElementById('orderModal');
  modalClose = document.getElementById('modalClose');
  modalBody = document.getElementById('modalBody');

  sidebarItems = document.querySelectorAll('.sidebar-item');
  shippingFeesPage = document.getElementById('shippingFeesPage');
  shippingFeesTableBody = document.getElementById('shippingFeesTableBody');

  tabBtns = document.querySelectorAll('.tab-btn');
  contentSections = document.querySelectorAll('.content-section');

  initEventListeners();
  initColumnResizers();
  updateDateDisplay();

  // 데이터 로드
  shippingFees = await loadShippingFeesAsync();
  productDb = await loadProductDbAsync();

  // 배송비 테이블 렌더링
  renderShippingFeesTable();

  switchPage('orders');
});

function initEventListeners() {
  // 파일 업로드
  uploadBtn.addEventListener('click', () => fileInput.click());
  uploadArea.addEventListener('click', (e) => {
    if (e.target !== uploadBtn && !uploadBtn.contains(e.target)) {
      fileInput.click();
    }
  });
  fileInput.addEventListener('change', handleFileSelect);

  // 드래그 앤 드롭
  uploadArea.addEventListener('dragover', (e) => {
    e.preventDefault();
    uploadArea.classList.add('dragover');
  });
  uploadArea.addEventListener('dragleave', () => {
    uploadArea.classList.remove('dragover');
  });
  uploadArea.addEventListener('drop', handleDrop);

  // 필터
  if (productTypeFilter) productTypeFilter.addEventListener('change', applyFilters);
  if (orderStatusFilter) orderStatusFilter.addEventListener('change', applyFilters);
  if (giftFilter) giftFilter.addEventListener('change', applyFilters);
  if (searchInput) searchInput.addEventListener('input', applyFilters);

  // 모달
  if (modalClose) modalClose.addEventListener('click', () => orderModal.style.display = 'none');
  if (orderModal) {
    orderModal.addEventListener('click', (e) => {
      if (e.target === orderModal) orderModal.style.display = 'none';
    });
  }

  // 사이드바
  sidebarItems.forEach(item => {
    item.addEventListener('click', () => {
      const page = item.getAttribute('data-page');
      switchPage(page);
    });
  });

  // 탭
  tabBtns.forEach(btn => {
    btn.addEventListener('click', () => {
      switchTab(btn.dataset.tab);
    });
  });

  // 엑셀 내보내기 버튼
  const exportCombined1Btn = document.getElementById('exportCombined1Btn');
  const exportCombined2Btn = document.getElementById('exportCombined2Btn');
  const exportKyungdong1Btn = document.getElementById('exportKyungdong1Btn');
  const exportKyungdong2Btn = document.getElementById('exportKyungdong2Btn');
  const exportKyungdong3Btn = document.getElementById('exportKyungdong3Btn');
  const exportLogenBtn = document.getElementById('exportLogenBtn');
  const exportShippingBtn = document.getElementById('exportShippingBtn');

  if (exportCombined1Btn) exportCombined1Btn.addEventListener('click', exportToCombined1);
  if (exportCombined2Btn) exportCombined2Btn.addEventListener('click', exportToCombined2);
  if (exportKyungdong1Btn) exportKyungdong1Btn.addEventListener('click', exportToKyungdong1);
  if (exportKyungdong2Btn) exportKyungdong2Btn.addEventListener('click', exportToKyungdong2);
  if (exportKyungdong3Btn) exportKyungdong3Btn.addEventListener('click', exportToKyungdong3);
  if (exportLogenBtn) exportLogenBtn.addEventListener('click', exportToLogen);
  if (exportShippingBtn) exportShippingBtn.addEventListener('click', exportToShipping);
}

function updateDateDisplay() {
  const dateElement = document.getElementById('currentDate');
  if (dateElement) {
    const now = new Date();
    const options = { year: 'numeric', month: 'long', day: 'numeric', weekday: 'long' };
    dateElement.textContent = now.toLocaleDateString('ko-KR', options);
  }
}

// ========== 컬럼 리사이즈 기능 ==========
function initColumnResizers() {
  // 모든 테이블에 리사이즈 핸들 추가
  const tables = ['rawTable', 'combined1Table', 'combined2Table', 'kyungdong1Table', 'kyungdong2Table', 'kyungdong3Table', 'logenTable', 'shippingTable', 'shippingFeesTable'];
  
  tables.forEach(tableId => {
    const table = document.getElementById(tableId);
    if (table) {
      setupTableResizers(table);
    }
  });
}

function setupTableResizers(table) {
  const thead = table.querySelector('thead');
  if (!thead) return;

  const headerRow = thead.querySelector('tr');
  if (!headerRow) return;

  // 기존 리사이저 제거
  headerRow.querySelectorAll('.column-resizer').forEach(resizer => resizer.remove());

  // 각 헤더 셀에 리사이저 추가
  const headers = headerRow.querySelectorAll('th');
  headers.forEach((th, index) => {
    // 마지막 컬럼은 리사이저 제외
    if (index === headers.length - 1) return;

    const resizer = document.createElement('div');
    resizer.className = 'column-resizer';
    th.style.position = 'relative';
    th.appendChild(resizer);

    let startX = 0;
    let startWidth = 0;
    let isResizing = false;

    resizer.addEventListener('mousedown', (e) => {
      e.preventDefault();
      isResizing = true;
      startX = e.pageX;
      startWidth = th.offsetWidth;
      resizer.classList.add('resizing');
      document.body.style.cursor = 'col-resize';
      document.body.style.userSelect = 'none';
    });

    document.addEventListener('mousemove', (e) => {
      if (!isResizing) return;
      
      const diff = e.pageX - startX;
      const newWidth = Math.max(50, startWidth + diff); // 최소 50px
      th.style.width = newWidth + 'px';
      th.style.minWidth = newWidth + 'px';
      
      // 같은 인덱스의 모든 셀 너비도 조절
      const tbody = table.querySelector('tbody');
      if (tbody) {
        tbody.querySelectorAll('tr').forEach(tr => {
          const td = tr.querySelectorAll('td')[index];
          if (td) {
            td.style.width = newWidth + 'px';
            td.style.minWidth = newWidth + 'px';
          }
        });
      }
    });

    document.addEventListener('mouseup', () => {
      if (isResizing) {
        isResizing = false;
        resizer.classList.remove('resizing');
        document.body.style.cursor = '';
        document.body.style.userSelect = '';
      }
    });
  });
}

// 테이블 렌더링 후 리사이저 다시 설정
function refreshTableResizers(tableId) {
  const table = document.getElementById(tableId);
  if (table) {
    setupTableResizers(table);
  }
}

// ========== 파일 처리 ==========
function handleFileSelect(e) {
  e.preventDefault();
  const file = e.target.files[0];
  if (file) processFile(file);
}

function handleDrop(e) {
  e.preventDefault();
  uploadArea.classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if (file) processFile(file);
}

function processFile(file) {
  const ext = file.name.split('.').pop().toLowerCase();

  if (ext === 'csv') {
    const reader = new FileReader();
    reader.onload = (e) => {
      parseCSV(e.target.result);
    };
    reader.readAsText(file, 'UTF-8');
  } else if (ext === 'xlsx' || ext === 'xls') {
    const reader = new FileReader();
    reader.onload = (e) => {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });
      const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
      const csv = XLSX.utils.sheet_to_csv(firstSheet);
      parseCSV(csv);
    };
    reader.readAsArrayBuffer(file);
  } else {
    alert('CSV 또는 Excel 파일만 업로드 가능합니다.');
    return;
  }

  fileName.textContent = file.name;
  fileInfo.style.display = 'flex';
  if (clearBtn) clearBtn.addEventListener('click', clearFile);
}

function parseCSV(text) {
  // BOM 제거
  if (text.charCodeAt(0) === 0xFEFF) {
    text = text.slice(1);
  }

  const rows = parseCSVToRows(text);

  if (!rows || rows.length < 2 || !rows[0]) {
    console.error('CSV 파싱 실패: rows=', rows);
    alert('CSV 파일 형식이 올바르지 않습니다.');
    return;
  }

  const headers = rows[0].map(h => h ? h.trim() : '');
  ordersData = [];
  rawOrderData = [];

  for (let i = 1; i < rows.length; i++) {
    const rowValues = rows[i];
    const rowObj = {};
    headers.forEach((header, index) => {
      rowObj[header] = rowValues[index] || '';
    });

    rawOrderData.push(rowObj);

    const order = parseOrderData(rowObj, productDb);
    if (order) ordersData.push(order);
  }

  filteredOrders = [...ordersData];
  updateUI();
}

function clearFile() {
  fileInput.value = '';
  fileInfo.style.display = 'none';
  ordersData = [];
  filteredOrders = [];
  packedOrders = [];
  rawOrderData = [];
  updateUI();
}

// ========== UI 업데이트 ==========
function updateUI() {
  if (ordersData.length === 0) {
    summarySection.style.display = 'none';
    filterSection.style.display = 'none';
    ordersSection.style.display = 'none';
    uploadArea.style.display = 'flex';
    uploadArea.style.flexDirection = 'column';
    uploadArea.style.alignItems = 'center';
    uploadArea.style.justifyContent = 'center';
    return;
  }

  uploadArea.style.display = 'none';
  summarySection.style.display = 'block';
  filterSection.style.display = 'flex';
  ordersSection.style.display = 'block';

  // 패킹 처리
  groupedOrders = groupOrdersByRecipient(filteredOrders);
  packedOrders = processPacking(groupedOrders, generateDesignCode);

  updateSummary();
  renderOrders();
  switchTab(currentTabId);
}

function updateSummary() {
  const uniqueCustomers = Object.keys(groupedOrders).length;
  const giftEligible = Object.values(groupedOrders).filter(g => g.totalPrice >= 95000).length;
  const kyungdongBoxes = packedOrders.filter(box => box.courier === 'kyungdong').length;
  const logenBoxes = packedOrders.filter(box => box.courier === 'logen').length;

  kyungdongCount.textContent = kyungdongBoxes;
  logenCount.textContent = logenBoxes;
  totalCustomers.textContent = uniqueCustomers;
  giftOrders.textContent = giftEligible;
}

// ========== 탭 전환 ==========
function switchTab(tabId) {
  currentTabId = tabId;

  tabBtns.forEach(btn => {
    btn.classList.toggle('active', btn.dataset.tab === tabId);
  });

  contentSections.forEach(section => {
    section.style.display = 'none';
    section.classList.remove('active');
  });

  if (tabId === 'combined1') {
    document.getElementById('combined1Section').style.display = 'block';
    renderCombined1Table();
  } else if (tabId === 'combined2') {
    document.getElementById('combined2Section').style.display = 'block';
    renderCombined2Table();
  } else if (tabId === 'kyungdong1') {
    document.getElementById('kyungdong1Section').style.display = 'block';
    renderKyungdong1Table();
  } else if (tabId === 'kyungdong2') {
    document.getElementById('kyungdong2Section').style.display = 'block';
    renderKyungdong2Table();
  } else if (tabId === 'kyungdong3') {
    document.getElementById('kyungdong3Section').style.display = 'block';
    renderKyungdong3Table();
  } else if (tabId === 'logen') {
    document.getElementById('logenSection').style.display = 'block';
    renderLogenTable();
  } else if (tabId === 'shipping') {
    document.getElementById('shippingSection').style.display = 'block';
    renderShippingTable();
  } else if (tabId === 'raw') {
    document.getElementById('rawSection').style.display = 'block';
    renderRawData();
  }
}

// ========== 페이지 전환 ==========
function switchPage(page) {
  sidebarItems.forEach(item => {
    item.classList.toggle('active', item.getAttribute('data-page') === page);
  });

  const ordersPage = document.querySelector('.main-content');

  if (page === 'orders') {
    if (ordersPage) ordersPage.style.display = 'block';
    if (shippingFeesPage) shippingFeesPage.style.display = 'none';
  } else if (page === 'shipping-fees') {
    if (ordersPage) ordersPage.style.display = 'none';
    if (shippingFeesPage) shippingFeesPage.style.display = 'block';
    renderShippingFeesTable();
  }
}

// ========== 필터 ==========
function applyFilters() {
  filteredOrders = ordersData.filter(order => {
    if (productTypeFilter && productTypeFilter.value && order.productType !== productTypeFilter.value) {
      return false;
    }
    if (orderStatusFilter && orderStatusFilter.value && order.status !== orderStatusFilter.value) {
      return false;
    }
    if (giftFilter && giftFilter.value) {
      const hasGift = order.gift && order.gift.trim() !== '';
      if (giftFilter.value === 'Y' && !hasGift) return false;
      if (giftFilter.value === 'N' && hasGift) return false;
    }
    if (searchInput && searchInput.value) {
      const searchTerm = searchInput.value.toLowerCase();
      const searchFields = [
        order.customerName,
        order.productName,
        order.orderId,
        order.address
      ].join(' ').toLowerCase();
      if (!searchFields.includes(searchTerm)) return false;
    }
    return true;
  });

  groupedOrders = groupOrdersByRecipient(filteredOrders);
  packedOrders = processPacking(groupedOrders, generateDesignCode);
  updateSummary();
  switchTab(currentTabId);
}

// ========== 렌더링 함수들 ==========
function renderOrders() {
  renderTableView();
}

function renderTableView() {
  if (!ordersTableBody) return;
  ordersTableBody.innerHTML = '';

  filteredOrders.forEach(order => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><input type="checkbox" class="order-checkbox" data-id="${order.id}"></td>
      <td>${order.orderId}</td>
      <td>${order.customerName}</td>
      <td><span class="product-badge ${getProductBadgeClass(order.productType)}">${order.productType}</span></td>
      <td>${order.productName}</td>
      <td>${order.quantity}</td>
      <td>${formatCurrency(order.price)}</td>
      <td><span class="status-badge ${getStatusClass(order.status)}">${order.status}</span></td>
    `;
    ordersTableBody.appendChild(tr);
  });
}

function getFilteredRawOrderData() {
  if (!rawOrderData || rawOrderData.length === 0) return [];
  if (!filteredOrders || filteredOrders.length === 0) return [];

  const rawRowsSet = new Set(filteredOrders.map(o => o.rawRow).filter(Boolean));
  if (rawRowsSet.size === 0) return rawOrderData;
  return rawOrderData.filter(row => rawRowsSet.has(row));
}

function renderRawData() {
  const table = document.getElementById('rawTable');
  if (!table) return;

  const thead = table.querySelector('thead');
  const tbody = table.querySelector('tbody');

  thead.innerHTML = '';
  tbody.innerHTML = '';

  const rowsToRender = getFilteredRawOrderData();
  if (!rowsToRender || rowsToRender.length === 0) {
    tbody.innerHTML = '<tr><td colspan="20" style="text-align: center; padding: 2rem; color: var(--text-muted);">데이터가 없습니다. 파일을 업로드해주세요.</td></tr>';
    return;
  }

  const headers = Object.keys(rowsToRender[0]);
  if (headers.length === 0) {
    tbody.innerHTML = '<tr><td colspan="20" style="text-align: center; padding: 2rem; color: var(--text-muted);">유효한 데이터가 없습니다.</td></tr>';
    return;
  }

  const trHead = document.createElement('tr');
  headers.forEach(header => {
    const th = document.createElement('th');
    th.textContent = header;
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);

  let currentCustomer = '';
  let groupIndex = 0;

  rowsToRender.forEach(row => {
    const tr = document.createElement('tr');
    tr.style.cursor = 'pointer';

    const customerName = row['수취인명'] || row['구매자명'] || '';
    if (customerName && customerName !== currentCustomer) {
      currentCustomer = customerName;
      groupIndex++;
    }

    tr.classList.add(groupIndex % 2 === 0 ? 'raw-row-group-even' : 'raw-row-group-odd');

    headers.forEach(header => {
      const td = document.createElement('td');
      let cellValue = row[header] || '';
      if (typeof cellValue === 'string') {
        cellValue = cellValue.replace(/\\/g, '');
      }

      const costColumns = ['최종 상품별 총 주문금액', '상품별 총 주문금액', '주문금액', '금액', '가격', '비용'];
      if (costColumns.some(col => header.includes(col))) {
        if (cellValue) {
          const numValue = parsePrice(cellValue);
          if (numValue > 0) {
            cellValue = formatCurrency(numValue);
          }
        }
      }

      td.textContent = cellValue;
      // 미등록 텍스트를 붉은색으로 표시
      if (typeof cellValue === 'string' && cellValue.includes('[미등록]')) {
        td.classList.add('unregistered-text');
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  });

  // 리사이저 다시 설정
  setupTableResizers(table);
}

function renderShippingFeesTable() {
  if (!shippingFeesTableBody) return;
  shippingFeesTableBody.innerHTML = '';

  if (shippingFees.length === 0) {
    shippingFeesTableBody.innerHTML = '<tr><td colspan="7" style="text-align: center; padding: 2rem;">배송비 데이터를 로드 중...</td></tr>';
    return;
  }

  shippingFees.forEach(fee => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${fee.productGroup}</td>
      <td>${fee.packageType}</td>
      <td>${fee.width || '-'}</td>
      <td>${fee.thickness || '-'}</td>
      <td>${fee.lengthMin || '-'}</td>
      <td>${fee.lengthMax || '-'}</td>
      <td>${formatCurrency(fee.fee)}</td>
    `;
    shippingFeesTableBody.appendChild(tr);
  });

  // 리사이저 다시 설정
  const shippingFeesTable = document.getElementById('shippingFeesTable');
  if (shippingFeesTable) {
    setupTableResizers(shippingFeesTable);
  }
}

// ========== 발주서 테이블 렌더링 ==========
function renderCombined1Table() {
  renderCombinedTableCommon('combined1Table', 'combined1TableBody', getCombined1Data);
}

function renderCombined2Table() {
  renderCombinedTableCommon('combined2Table', 'combined2TableBody', getCombined2Data);
}

function renderCombinedTableCommon(tableId, tbodyId, getDataFn) {
  const table = document.getElementById(tableId);
  const thead = table?.querySelector('thead');
  const tbody = document.getElementById(tbodyId);
  if (!tbody || !table) return;

  thead.innerHTML = '';
  tbody.innerHTML = '';

  const data = getDataFn();
  if (data.length <= 1) {
    tbody.innerHTML = '<tr><td colspan="15" style="text-align: center; padding: 2rem;">데이터가 없습니다.</td></tr>';
    return;
  }

  // 헤더 렌더링
  if (data.length > 0 && data[0]) {
    const headerRow = document.createElement('tr');
    data[0].forEach(header => {
      const th = document.createElement('th');
      th.textContent = header || '';
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
  }

  // 데이터 렌더링 (고객별 색상 구분)
  let currentCustomer = '';
  let groupIndex = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tr = document.createElement('tr');

    // 수취인명 컬럼 (2번째 컬럼, 인덱스 1)으로 고객 구분
    const customerName = row[1] || '';
    if (customerName && customerName !== currentCustomer) {
      currentCustomer = customerName;
      groupIndex++;
    }

    // 고객별 교차 색상 적용
    tr.classList.add(groupIndex % 2 === 0 ? 'raw-row-group-even' : 'raw-row-group-odd');

    row.forEach(cell => {
      const td = document.createElement('td');
      const cellValue = cell || '';
      td.textContent = cellValue;
      // 미등록 텍스트를 붉은색으로 표시
      if (typeof cellValue === 'string' && cellValue.includes('[미등록]')) {
        td.classList.add('unregistered-text');
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }

  // 리사이저 다시 설정
  setupTableResizers(table);
}

function renderKyungdong1Table() {
  renderKyungdongTableCommon('kyungdong1Table', 'kyungdong1TableBody', getKyungdong1Data);
}

function renderKyungdong2Table() {
  renderKyungdongTableCommon('kyungdong2Table', 'kyungdong2TableBody', getKyungdong2Data);
}

function renderKyungdong3Table() {
  renderKyungdongTableCommon('kyungdong3Table', 'kyungdong3TableBody', getKyungdong3Data);
}

function renderKyungdongTableCommon(tableId, tbodyId, getDataFn) {
  const table = document.getElementById(tableId);
  const thead = table?.querySelector('thead');
  const tbody = document.getElementById(tbodyId);
  if (!tbody || !table) return;
  
  thead.innerHTML = '';
  tbody.innerHTML = '';

  const data = getDataFn();
  if (data.length <= 1) {
    tbody.innerHTML = '<tr><td colspan="20" style="text-align: center; padding: 2rem;">경동택배 데이터가 없습니다.</td></tr>';
    return;
  }

  // 헤더 렌더링
  if (data.length > 0 && data[0]) {
    const headerRow = document.createElement('tr');
    data[0].forEach(header => {
      const th = document.createElement('th');
      th.textContent = header || '';
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
  }

  // 데이터 렌더링 (고객별 색상 구분)
  let currentCustomer = '';
  let groupIndex = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tr = document.createElement('tr');

    // 받는분 컬럼 (첫 번째 컬럼, 인덱스 0)으로 고객 구분
    const customerName = row[0] || '';
    if (customerName && customerName !== currentCustomer) {
      currentCustomer = customerName;
      groupIndex++;
    }

    // 고객별 교차 색상 적용
    tr.classList.add(groupIndex % 2 === 0 ? 'raw-row-group-even' : 'raw-row-group-odd');

    row.forEach(cell => {
      const td = document.createElement('td');
      const cellValue = cell || '';
      td.textContent = cellValue;
      // 미등록 텍스트를 붉은색으로 표시
      if (typeof cellValue === 'string' && cellValue.includes('[미등록]')) {
        td.classList.add('unregistered-text');
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }

  // 리사이저 다시 설정
  setupTableResizers(table);
}

function renderLogenTable() {
  const table = document.getElementById('logenTable');
  const thead = table?.querySelector('thead');
  const tbody = document.getElementById('logenTableBody');
  if (!tbody || !table) return;
  
  thead.innerHTML = '';
  tbody.innerHTML = '';

  const data = getLogenData();
  if (data.length <= 1) {
    tbody.innerHTML = '<tr><td colspan="15" style="text-align: center; padding: 2rem;">로젠택배 데이터가 없습니다.</td></tr>';
    return;
  }

  // 헤더 렌더링
  if (data.length > 0 && data[0]) {
    const headerRow = document.createElement('tr');
    data[0].forEach(header => {
      const th = document.createElement('th');
      th.textContent = header || '';
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
  }

  // 데이터 렌더링 (고객별 색상 구분)
  let currentCustomer = '';
  let groupIndex = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tr = document.createElement('tr');

    // 받는분 컬럼 (첫 번째 컬럼, 인덱스 0)으로 고객 구분
    const customerName = row[0] || '';
    if (customerName && customerName !== currentCustomer) {
      currentCustomer = customerName;
      groupIndex++;
    }

    // 고객별 교차 색상 적용
    tr.classList.add(groupIndex % 2 === 0 ? 'raw-row-group-even' : 'raw-row-group-odd');

    row.forEach(cell => {
      const td = document.createElement('td');
      const cellValue = cell || '';
      td.textContent = cellValue;
      // 미등록 텍스트를 붉은색으로 표시
      if (typeof cellValue === 'string' && cellValue.includes('[미등록]')) {
        td.classList.add('unregistered-text');
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }

  // 리사이저 다시 설정
  setupTableResizers(table);
}

function renderShippingTable() {
  const table = document.getElementById('shippingTable');
  const thead = table?.querySelector('thead');
  const tbody = document.getElementById('shippingTableBody');
  if (!tbody || !table) return;
  
  thead.innerHTML = '';
  tbody.innerHTML = '';

  const data = getShippingData();
  if (data.length <= 1) {
    tbody.innerHTML = '<tr><td colspan="10" style="text-align: center; padding: 2rem;">발송처리 데이터가 없습니다.</td></tr>';
    return;
  }

  // 헤더 렌더링
  if (data.length > 0 && data[0]) {
    const headerRow = document.createElement('tr');
    data[0].forEach(header => {
      const th = document.createElement('th');
      th.textContent = header || '';
      headerRow.appendChild(th);
    });
    thead.appendChild(headerRow);
  }

  // 데이터 렌더링 (고객별 색상 구분)
  let currentCustomer = '';
  let groupIndex = 0;

  console.log('[발송처리] 렌더링 시작, 데이터 행 수:', data.length - 1);

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tr = document.createElement('tr');

    // 수취인명 컬럼 (두 번째 컬럼, 인덱스 1)으로 고객 구분
    const customerName = row[1] || '';
    if (customerName && customerName !== currentCustomer) {
      currentCustomer = customerName;
      groupIndex++;
      console.log(`[발송처리] 새 고객: ${customerName}, groupIndex: ${groupIndex}`);
    }

    // 고객별 교차 색상 적용
    const className = groupIndex % 2 === 0 ? 'raw-row-group-even' : 'raw-row-group-odd';
    tr.classList.add(className);

    if (i <= 3) {
      console.log(`[발송처리] 행${i}: ${customerName}, 클래스: ${className}`);
    }

    row.forEach(cell => {
      const td = document.createElement('td');
      const cellValue = cell || '';
      td.textContent = cellValue;
      // 미등록 텍스트를 붉은색으로 표시
      if (typeof cellValue === 'string' && cellValue.includes('[미등록]')) {
        td.classList.add('unregistered-text');
      }
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }

  console.log('[발송처리] 렌더링 완료');

  // 리사이저 다시 설정
  setupTableResizers(table);
}

// ========== 발주서 데이터 생성 ==========
function getCombined1Data() {
  return getCombinedDataCommon(false);
}

function getCombined2Data() {
  const headers = ['상품주문번호', '수취인명', '운임타입', '송장수량', '디자인', '길이', '수량', '디자인+수량', '배송메세지', '합포장', '비고', '주문하신분 핸드폰', '받는분 핸드폰', '우편번호', '받는분 주소', '주문자명', '실제결제금액'];
  const data = [headers];
  const MAX_DESIGN_WITH_QTY_LENGTH = 30; // 디자인+수량 컬럼 임계값 (박혜정 36자, 김태종 48자 → 분할)

  // 고객별로 박스 그룹화 및 총 결제금액 계산
  const customerGroups = new Map();
  packedOrders.forEach(box => {
    const customerKey = box.customerName || box.recipientLabel || '';
    if (!customerGroups.has(customerKey)) {
      customerGroups.set(customerKey, {
        boxes: [],
        totalPayment: 0
      });
    }
    const group = customerGroups.get(customerKey);
    group.boxes.push(box);

    // 이 박스의 결제금액 합산
    const boxPayment = box.items.reduce((sum, item) => {
      const itemRawRow = item.rawRow || {};
      const payment = parsePrice(itemRawRow['최종 상품별 총 주문금액']) || item.price || 0;
      return sum + payment;
    }, 0);
    group.totalPayment += boxPayment;
  });

  // 고객별 첫 행 추적
  const customerFirstRowFlags = new Map();

  packedOrders.forEach(box => {
    const designCode = box.designText || '';
    const quantity = box.totalQtyInBox || 1;

    // 원본 데이터에서 추가 정보 가져오기
    const firstItem = box.group?.items?.[0];
    const rawRow = firstItem?.rawRow || {};
    const buyerPhone = rawRow['구매자연락처'] || '';
    const buyerName = rawRow['구매자명'] || '';

    // 고객별 총 결제금액 가져오기
    const customerKey = box.customerName || box.recipientLabel || '';
    const customerGroup = customerGroups.get(customerKey);
    const formattedPayment = customerGroup ? customerGroup.totalPayment.toLocaleString('ko-KR') : '';

    // 이 고객의 첫 행인지 확인
    const isCustomerFirstRow = !customerFirstRowFlags.has(customerKey);
    if (isCustomerFirstRow) {
      customerFirstRowFlags.set(customerKey, true);
    }

    // 증정품 계산을 위한 고객 전체 롤매트 길이와 결제금액
    const totalPayment = customerGroup ? customerGroup.totalPayment : 0;

    // 배송메세지 생성 (파손주의 표시 추가)
    let deliveryMemo = (box.deliveryMemos || []).join(' / ');
    const hasFragileMarking = (box.recipientLabel || box.customerName || '').includes('★');
    if (hasFragileMarking) {
      deliveryMemo = deliveryMemo ? `★파손주의★ ${deliveryMemo}` : '★파손주의★';
    }

    // 증정품 계산
    const rollMatProductIds = ['6092903705', '6626596277', '4200445704'];
    const puzzleMatProductId = '5994906898';

    // 롤매트 총 길이 계산
    const totalRollLength = box.items.reduce((sum, item) => {
      const productId = item.rawRow?.['상품번호'] || '';
      if (rollMatProductIds.includes(productId)) {
        return sum + (item.lengthM || 0) * (item.quantity || 1);
      }
      return sum;
    }, 0);

    // 롤매트 또는 퍼즐매트 포함 여부 확인
    const hasRollOrPuzzle = box.items.some(item => {
      const productId = item.rawRow?.['상품번호'] || '';
      return rollMatProductIds.includes(productId) || productId === puzzleMatProductId;
    });

    const gifts = [];

    // 테이프 증정 (롤매트 10m 이상 + 195,000원 이상)
    if (totalRollLength >= 10 && totalPayment >= 195000) {
      let tapeCount;
      if (totalRollLength >= 50) {
        tapeCount = 3;
      } else if (totalRollLength >= 30) {
        tapeCount = 2;
      } else {
        tapeCount = 1;
      }
      gifts.push(`★증정★테이프20m x${tapeCount}`);
    }

    // 팻말 증정 (롤매트 or 퍼즐매트 + 200,000원 이상)
    if (hasRollOrPuzzle && totalPayment >= 200000) {
      gifts.push('★증정★팻말 x1');
    }

    const giftText = gifts.join(' / ');

    const fee = calculateShippingFee(box, shippingFees);

    // 각 아이템별로 디자인코드 + 길이 추출
    const itemsWithDesign = box.items.map(item => {
      const code = generateDesignCode(item);
      const lengthM = item.lengthM || 0;
      const qty = item.quantity || 1;
      return {
        item: item,
        designCode: code,
        lengthM: lengthM,
        qty: qty
      };
    });

    // 같은 디자인+길이 아이템 그룹핑 및 수량 합산
    const groupedItems = [];
    itemsWithDesign.forEach(itemData => {
      const key = `${itemData.designCode}_${itemData.lengthM}`;
      const existing = groupedItems.find(g => `${g.designCode}_${g.lengthM}` === key);
      if (existing) {
        existing.qty += itemData.qty;
      } else {
        groupedItems.push({ ...itemData });
      }
    });

    // 그룹핑된 아이템으로 텍스트 생성
    const itemTexts = groupedItems.map(itemData => {
      let result = itemData.designCode;
      // 테이프는 이미 디자인 코드에 길이가 포함되어 있으므로 길이 표시 생략
      const isTape = itemData.designCode.includes('테이프');
      if (itemData.lengthM > 0 && !isTape) {
        result += ` ${itemData.lengthM}m`;
      }
      result += ` x${itemData.qty}`;
      return {
        text: result,
        designCode: itemData.designCode,
        lengthM: itemData.lengthM,
        qty: itemData.qty
      };
    });

    const designWithQty = itemTexts.map(it => it.text).join(' / ');

    // 디버깅: 텍스트 길이 확인
    if (box.customerName === '장하은' || box.customerName === '김태종' || box.customerName?.includes('박혜정')) {
      console.log(`[송장화2 DEBUG] ${box.customerName}:`, {
        designWithQty: designWithQty,
        length: designWithQty.length,
        originalItemsCount: box.items.length,
        groupedItemsCount: groupedItems.length,
        groupedItems: groupedItems.map(g => ({ design: g.designCode, length: g.lengthM, qty: g.qty })),
        willSplit: designWithQty.length >= MAX_DESIGN_WITH_QTY_LENGTH && groupedItems.length > 1
      });
    }

    // 텍스트 길이 초과 여부 확인 (그룹핑된 아이템 개수로 체크)
    if (designWithQty.length >= MAX_DESIGN_WITH_QTY_LENGTH && groupedItems.length > 1) {
      console.log(`[송장화2 분할] ${box.customerName}: ${groupedItems.length}개 그룹 분할`);

      // 분할 로직: 아이템별로 행 생성
      itemTexts.forEach((itemText, index) => {
        const isFirstRow = index === 0;
        const isTape = itemText.designCode.includes('테이프');
        // 이 박스의 첫 행이면서 고객의 첫 행인 경우에만 실제결제금액 표시
        const shouldShowPayment = isFirstRow && isCustomerFirstRow;
        const row = [
          isFirstRow ? (box.group?.orderId || '') : '',
          isFirstRow ? (box.recipientLabel || box.customerName || '') : '',
          isFirstRow ? fee : '',
          isFirstRow ? 1 : '',
          itemText.designCode,
          (itemText.lengthM && !isTape) ? `${itemText.lengthM}m` : '',
          itemText.qty,
          itemText.text,
          isFirstRow ? deliveryMemo : '',
          isFirstRow ? (box.isCombined ? '합' : '') : '',
          isFirstRow ? giftText : '',
          isFirstRow ? buyerPhone : '',
          isFirstRow ? (box.phone || '') : '',
          isFirstRow ? (box.zipCode || '') : '',
          isFirstRow ? (box.address || '') : '',
          isFirstRow ? buyerName : '',
          shouldShowPayment ? formattedPayment : ''
        ];
        data.push(row);
      });
    } else {
      // 기존 로직: 단일 행 (고객의 첫 행인 경우에만 실제결제금액 표시)
      const row = [
        box.group?.orderId || '',
        box.recipientLabel || box.customerName || '',
        fee,
        1,
        designCode,
        box.items[0]?.lengthM ? `${box.items[0].lengthM}m` : '',
        quantity,
        designWithQty,
        deliveryMemo,
        box.isCombined ? '합' : '',
        giftText,
        buyerPhone,
        box.phone || '',
        box.zipCode || '',
        box.address || '',
        buyerName,
        isCustomerFirstRow ? formattedPayment : ''
      ];
      data.push(row);
    }
  });

  return data;
}

function getCombinedDataCommon(useDesignWithQty) {
  const headers = ['상품주문번호', '수취인명', '운임타입', '송장수량', '디자인', '길이', '수량', '추가상품1', '배송메세지', '합포장', '비고', '주문하신분 핸드폰', '받는분 핸드폰', '우편번호', '받는분 주소', '주문자명', '실제결제금액'];
  const data = [headers];

  packedOrders.forEach(box => {
    // 원본 데이터에서 추가 정보 가져오기
    const firstItem = box.group?.items?.[0];
    const rawRow = firstItem?.rawRow || {};
    const buyerPhone = rawRow['구매자연락처'] || '';
    const buyerName = rawRow['구매자명'] || '';

    // 각 원본 주문 라인별로 행 생성
    box.items.forEach((item, itemIndex) => {
      const designCode = box.designText || '';  // 송장화2와 동일하게 box.designText 사용
      const quantity = item.quantity || 1;
      const deliveryMemo = item.deliveryMemo || (box.deliveryMemos || []).join(' / ');

      // 원본 주문 라인의 실제 결제 금액 (최종 상품별 총 주문금액)
      const itemRawRow = item.rawRow || {};
      const actualPayment = parsePrice(itemRawRow['최종 상품별 총 주문금액']) || item.price || 0;
      // 천 단위 콤마 추가
      const formattedPayment = actualPayment.toLocaleString('ko-KR');

      const row = [
        itemRawRow['상품주문번호'] || box.group?.orderId || '',
        box.recipientLabel || box.customerName || '',
        '',  // 운임타입 공란
        '',  // 송장수량 공란
        designCode,
        item.lengthM ? `${item.lengthM}m` : '',
        quantity,
        '',
        deliveryMemo,
        box.isCombined ? 'Y' : 'N',
        box.remark || '',
        buyerPhone,
        box.phone || '',
        box.zipCode || '',
        box.address || '',
        buyerName,
        formattedPayment
      ];
      data.push(row);
    });
  });

  return data;
}

// 경동전체1 양식
function getKyungdong1Data() {
  const headers = ['받는분', '운임타입', '송장수량', '디자인', '배송메세지', '합포장', '비고', '받는분 전화번호', '받는분 전화번호', '우편번호', '받는분 주소', '', '', '포장상태'];
  const data = [headers];

  const kyungdongOrders = packedOrders.filter(box => box.courier === 'kyungdong');

  // 고객별로 박스 그룹화
  const customerGroups = new Map();
  kyungdongOrders.forEach(box => {
    const customerKey = box.customerName || box.recipientLabel || '';
    if (!customerGroups.has(customerKey)) {
      customerGroups.set(customerKey, []);
    }
    customerGroups.get(customerKey).push(box);
  });

  // 고객별로 처리
  customerGroups.forEach((boxes, customerKey) => {
    // 고객별 총 결제금액 및 롤매트 길이 계산 (증정품 계산용)
    let totalPayment = 0;
    const rollMatProductIds = ['6092903705', '6626596277', '4200445704'];
    let totalRollLength = 0;
    
    boxes.forEach(box => {
      box.items.forEach(item => {
        const itemRawRow = item.rawRow || {};
        const payment = parsePrice(itemRawRow['최종 상품별 총 주문금액']) || item.price || 0;
        totalPayment += payment;
        
        const productId = item.rawRow?.['상품번호'] || '';
        if (rollMatProductIds.includes(productId)) {
          totalRollLength += (item.lengthM || 0) * (item.quantity || 1);
        }
      });
    });

    // 증정품 계산
    const gifts = [];
    if (totalRollLength >= 10 && totalPayment >= 195000) {
      let tapeCount;
      if (totalRollLength >= 50) {
        tapeCount = 3;
      } else if (totalRollLength >= 30) {
        tapeCount = 2;
      } else {
        tapeCount = 1;
      }
      gifts.push(`★증정★테이프20mx${tapeCount}`);
    }
    const giftText = gifts.join(' / ');

    // 각 박스별로 행 생성
    boxes.forEach((box, boxIndex) => {
      const fee = calculateShippingFee(box, shippingFees);
      const deliveryMemo = (box.deliveryMemos || []).join(' / ');
      const isFirstBox = boxIndex === 0;
      const MAX_DESIGN_LENGTH = 30; // 디자인 텍스트 최대 길이

      // 각 아이템별로 디자인코드 + 길이 추출
      const itemsWithDesign = box.items.map(item => {
        const code = generateDesignCode(item);
        const lengthM = item.lengthM || 0;
        const qty = item.quantity || 1;
        return {
          item: item,
          designCode: code,
          lengthM: lengthM,
          qty: qty
        };
      });

      // 같은 디자인+길이 아이템 그룹핑 및 수량 합산
      const groupedItems = [];
      itemsWithDesign.forEach(itemData => {
        const key = `${itemData.designCode}_${itemData.lengthM}`;
        const existing = groupedItems.find(g => `${g.designCode}_${g.lengthM}` === key);
        if (existing) {
          existing.qty += itemData.qty;
        } else {
          groupedItems.push({ ...itemData });
        }
      });

      // 그룹핑된 아이템으로 텍스트 생성
      const itemTexts = groupedItems.map(itemData => {
        let result = itemData.designCode;
        // 테이프는 이미 디자인 코드에 길이가 포함되어 있으므로 길이 표시 생략
        const isTape = itemData.designCode.includes('테이프');
        if (itemData.lengthM > 0 && !isTape) {
          result += ` ${itemData.lengthM}m`;
        }
        result += ` x${itemData.qty}`;
        return {
          text: result,
          designCode: itemData.designCode,
          lengthM: itemData.lengthM,
          qty: itemData.qty
        };
      });

      const designWithQty = itemTexts.map(it => it.text).join(' / ');

      // 텍스트 길이 초과 여부 확인 (그룹핑된 아이템 개수로 체크)
      if (designWithQty.length >= MAX_DESIGN_LENGTH && groupedItems.length > 1) {
        // 분할 로직: 아이템별로 행 생성
        itemTexts.forEach((itemText, itemIndex) => {
          const isFirstItem = itemIndex === 0;
          const row = [
            isFirstItem ? (box.recipientLabel || box.customerName || '') : '',
            isFirstItem ? fee : '', // 운임타입 (각 박스별로, 합포장 포함)
            isFirstItem && isFirstBox ? 1 : '', // 첫 번째 박스의 첫 번째 아이템에만 송장수량
            itemText.text,
            isFirstItem && isFirstBox ? deliveryMemo : '', // 첫 번째 박스의 첫 번째 아이템에만 배송메세지
            isFirstItem && isFirstBox ? (boxes.length > 1 ? '합' : '') : '', // 첫 번째 박스의 첫 번째 아이템에만 합포장 표시
            isFirstItem && isFirstBox ? giftText : '', // 첫 번째 박스의 첫 번째 아이템에만 비고(증정품)
            isFirstItem ? (box.phone || '') : '',
            isFirstItem ? (box.phone || '') : '',
            isFirstItem ? (box.zipCode || '') : '',
            isFirstItem ? (box.address || '') : '',
            '', // 빈 컬럼
            '', // 빈 컬럼
            isFirstItem ? (box.packagingType === 'vinyl' ? '비닐' : '박스') : '' // 포장상태
          ];
          data.push(row);
        });
      } else {
        // 기존 로직: 단일 행
        const row = [
          box.recipientLabel || box.customerName || '',
          fee, // 운임타입 (각 박스별로, 합포장 포함)
          isFirstBox ? 1 : '', // 첫 번째 박스에만 송장수량
          designWithQty || box.designText || '',
          isFirstBox ? deliveryMemo : '', // 첫 번째 박스에만 배송메세지
          isFirstBox ? (boxes.length > 1 ? '합' : '') : '', // 첫 번째 박스에만 합포장 표시
          isFirstBox ? giftText : '', // 첫 번째 박스에만 비고(증정품)
          box.phone || '',
          box.phone || '',
          box.zipCode || '',
          box.address || '',
          '', // 빈 컬럼
          '', // 빈 컬럼
          box.packagingType === 'vinyl' ? '비닐' : '박스' // 포장상태
        ];
        data.push(row);
      }
    });
  });

  return data;
}

// 경동발주서 양식2
function getKyungdong2Data() {
  const headers = ['받는분', '주소', '상세주소', '운송장번호', '고객사주문번호', '우편번호', '도착영업소', '전화번호', '기타전화번호', '선불후불', '품목명', '수량', '포장상태', '가로', '세로', '높이', '무게', '개별단가(만원)', '운임', '기타운임', '별도운임', '할증운임', '도서운임', '메모1'];
  const data = [headers];

  const kyungdongOrders = packedOrders.filter(box => box.courier === 'kyungdong');

  kyungdongOrders.forEach(box => {
    const { base, detail } = splitAddress(box.address);
    const dims = getDimensions(box);
    const fee = calculateShippingFee(box, shippingFees);
    
    // 증정품 계산
    const customerKey = box.customerName || box.recipientLabel || '';
    const customerBoxes = packedOrders.filter(b => 
      (b.customerName || b.recipientLabel || '') === customerKey && b.courier === 'kyungdong'
    );
    
    let totalPayment = 0;
    customerBoxes.forEach(cb => {
      cb.items.forEach(item => {
        const itemRawRow = item.rawRow || {};
        const payment = parsePrice(itemRawRow['최종 상품별 총 주문금액']) || item.price || 0;
        totalPayment += payment;
      });
    });
    
    const rollMatProductIds = ['6092903705', '6626596277', '4200445704'];
    let totalRollLength = 0;
    customerBoxes.forEach(cb => {
      cb.items.forEach(item => {
        const productId = item.rawRow?.['상품번호'] || '';
        if (rollMatProductIds.includes(productId)) {
          totalRollLength += (item.lengthM || 0) * (item.quantity || 1);
        }
      });
    });

    const gifts = [];
    if (totalRollLength >= 10 && totalPayment >= 195000) {
      let tapeCount;
      if (totalRollLength >= 50) {
        tapeCount = 3;
      } else if (totalRollLength >= 30) {
        tapeCount = 2;
      } else {
        tapeCount = 1;
      }
      gifts.push(`★증정★테이프20mx${tapeCount}`);
    }
    const giftText = gifts.join(' / ');

    const row = [
      box.recipientLabel || box.customerName || '',
      base,
      (box.deliveryMemos || []).join(' / '),
      '',
      box.group?.orderId || '',
      box.zipCode || '',
      '',
      box.phone || '',
      box.phone || '',
      '선불',
      giftText || box.designText || '',
      1,
      box.packagingType === 'vinyl' ? '비닐' : '박스',
      dims.width || 1,
      dims.depth || 1,
      dims.height || 1,
      1,
      50,
      fee,
      100,
      0,
      0,
      0,
      box.designText || ''
    ];
    data.push(row);
  });

  return data;
}

// 경동샘플 양식3
function getKyungdong3Data() {
  const headers = ['받는분', '주소', '상세주소', '운송장번호', '고객사주문번호', '우편번호', '도착영업소', '전화번호', '기타전화번호', '선불후불', '품목명', '수량', '포장상태', '가로', '세로', '높이', '무게', '개별단가(만원)', '운임', '기타운임', '별도운임', '할증운임', '도서운임', '메모1', '메모2', '메모3', '메모1,메모2,메모3'];
  const data = [headers];

  const kyungdongOrders = packedOrders.filter(box => box.courier === 'kyungdong');

  kyungdongOrders.forEach(box => {
    const { base, detail } = splitAddress(box.address);
    const dims = getDimensions(box);
    const fee = calculateShippingFee(box, shippingFees);
    
    // 증정품 계산
    const customerKey = box.customerName || box.recipientLabel || '';
    const customerBoxes = packedOrders.filter(b => 
      (b.customerName || b.recipientLabel || '') === customerKey && b.courier === 'kyungdong'
    );
    
    let totalPayment = 0;
    customerBoxes.forEach(cb => {
      cb.items.forEach(item => {
        const itemRawRow = item.rawRow || {};
        const payment = parsePrice(itemRawRow['최종 상품별 총 주문금액']) || item.price || 0;
        totalPayment += payment;
      });
    });
    
    const rollMatProductIds = ['6092903705', '6626596277', '4200445704'];
    let totalRollLength = 0;
    customerBoxes.forEach(cb => {
      cb.items.forEach(item => {
        const productId = item.rawRow?.['상품번호'] || '';
        if (rollMatProductIds.includes(productId)) {
          totalRollLength += (item.lengthM || 0) * (item.quantity || 1);
        }
      });
    });

    const gifts = [];
    if (totalRollLength >= 10 && totalPayment >= 195000) {
      let tapeCount;
      if (totalRollLength >= 50) {
        tapeCount = 3;
      } else if (totalRollLength >= 30) {
        tapeCount = 2;
      } else {
        tapeCount = 1;
      }
      gifts.push(`★증정★테이프20mx${tapeCount}`);
    }
    const giftText = gifts.join(' / ');

    const memo1 = box.designText || '';
    const memo2 = '';
    const memo3 = '';
    const memoCombined = [memo1, memo2, memo3].filter(m => m).join(',');

    const row = [
      box.recipientLabel || box.customerName || '',
      base,
      (box.deliveryMemos || []).join(' / '),
      '',
      box.group?.orderId || '',
      box.zipCode || '',
      '',
      box.phone || '',
      box.phone || '',
      '선불',
      giftText || box.designText || '',
      1,
      box.packagingType === 'vinyl' ? '비닐' : '박스',
      dims.width || 1,
      dims.depth || 1,
      dims.height || 1,
      1,
      50,
      fee,
      100,
      0,
      0,
      0,
      memo1,
      memo2,
      memo3,
      memoCombined
    ];
    data.push(row);
  });

  return data;
}

function getLogenData() {
  const headers = ['받는분', '받는분 전화번호', '받는분 핸드폰', '우편번호', '받는분 주소', '운임타입', '송장수량', '디자인', '디자인+수량', '배송메세지', '합배송여부', '배송운임'];
  const data = [headers];

  const logenOrders = packedOrders.filter(box => box.courier === 'logen');

  logenOrders.forEach(box => {
    const designCode = box.designText || '';
    // 각 아이템별로 디자인코드 + 길이 + 수량 형태로 생성
    const designWithQty = box.items.map(item => {
      const code = generateDesignCode(item);
      const lengthM = item.lengthM || 0;
      const qty = item.quantity || 1;
      let result = code;
      if (lengthM > 0) {
        result += ` ${lengthM}m`;
      }
      result += ` x${qty}`;
      return result;
    }).join(' / ');
    const fee = calculateShippingFee(box, shippingFees);

    const row = [
      box.recipientLabel || box.customerName || '',
      box.phone || '',
      box.phone || '',
      box.zipCode || '',
      box.address || '',
      '선불',
      box.totalBoxes || 1,
      designCode,
      designWithQty,
      (box.deliveryMemos || []).join(' / '),
      box.totalBoxes > 1 ? 'Y' : 'N',
      fee
    ];
    data.push(row);
  });

  return data;
}

function getShippingData() {
  const headers = ['상품주문번호', '수취인명', '배송지', '우편번호', '연락처', '상품명', '수량', '배송메세지', '택배사'];
  const data = [headers];

  packedOrders.forEach(box => {
    const row = [
      box.group?.orderId || '',
      box.customerName || '',
      box.address || '',
      box.zipCode || '',
      box.phone || '',
      box.designText || '',
      box.totalQtyInBox || 1,
      (box.deliveryMemos || []).join(' / '),
      box.courier === 'kyungdong' ? '경동' : '로젠'
    ];
    data.push(row);
  });

  return data;
}

// ========== 엑셀 내보내기 ==========
function exportToKyungdong1() {
  downloadExcel(getKyungdong1Data(), '경동전체1');
}

function exportToKyungdong2() {
  downloadExcel(getKyungdong2Data(), '경동발주서');
}

function exportToKyungdong3() {
  downloadExcel(getKyungdong3Data(), '경동샘플');
}

function exportToLogen() {
  downloadExcel(getLogenData(), '로젠발주서');
}

function exportToShipping() {
  downloadExcel(getShippingData(), '발송처리');
}

function exportToCombined1() {
  downloadExcel(getCombined1Data(), '송장화1');
}

function exportToCombined2() {
  downloadExcel(getCombined2Data(), '송장화2');
}

// 전역으로 내보내기 함수 노출 (버튼 onclick용)
window.exportToKyungdong1 = exportToKyungdong1;
window.exportToKyungdong2 = exportToKyungdong2;
window.exportToKyungdong3 = exportToKyungdong3;
window.exportToLogen = exportToLogen;
window.exportToShipping = exportToShipping;
window.exportToCombined1 = exportToCombined1;
window.exportToCombined2 = exportToCombined2;

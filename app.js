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
  const exportKyungdongBtn = document.getElementById('exportKyungdongBtn');
  const exportLogenBtn = document.getElementById('exportLogenBtn');
  const exportShippingBtn = document.getElementById('exportShippingBtn');

  if (exportCombined1Btn) exportCombined1Btn.addEventListener('click', exportToCombined1);
  if (exportCombined2Btn) exportCombined2Btn.addEventListener('click', exportToCombined2);
  if (exportKyungdongBtn) exportKyungdongBtn.addEventListener('click', exportToKyungdong);
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
  const tables = ['rawTable', 'combined1Table', 'combined2Table', 'kyungdongTable', 'logenTable', 'shippingTable', 'shippingFeesTable'];
  
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
  } else if (tabId === 'kyungdong') {
    document.getElementById('kyungdongSection').style.display = 'block';
    renderKyungdongTable();
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

  // 데이터 렌더링
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tr = document.createElement('tr');
    row.forEach(cell => {
      const td = document.createElement('td');
      td.textContent = cell || '';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }

  // 리사이저 다시 설정
  setupTableResizers(table);
}

function renderKyungdongTable() {
  const table = document.getElementById('kyungdongTable');
  const thead = table?.querySelector('thead');
  const tbody = document.getElementById('kyungdongTableBody');
  if (!tbody || !table) return;
  
  thead.innerHTML = '';
  tbody.innerHTML = '';

  const data = getKyungdongData();
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

  // 데이터 렌더링
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tr = document.createElement('tr');
    row.forEach(cell => {
      const td = document.createElement('td');
      td.textContent = cell || '';
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

  // 데이터 렌더링
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tr = document.createElement('tr');
    row.forEach(cell => {
      const td = document.createElement('td');
      td.textContent = cell || '';
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

  // 데이터 렌더링
  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tr = document.createElement('tr');
    row.forEach(cell => {
      const td = document.createElement('td');
      td.textContent = cell || '';
      tr.appendChild(td);
    });
    tbody.appendChild(tr);
  }
}

// ========== 발주서 데이터 생성 ==========
function getCombined1Data() {
  return getCombinedDataCommon(false);
}

function getCombined2Data() {
  const headers = ['상품주문번호', '수취인명', '운임타입', '송장수량', '디자인', '길이', '수량', '디자인+수량', '배송메세지', '합포장', '비고'];
  const data = [headers];

  packedOrders.forEach(box => {
    const designCode = box.designText || '';
    const quantity = box.totalQtyInBox || 1;
    // 각 아이템별로 디자인코드 + 길이 + 수량 형태로 생성
    const designWithQty = box.items.map(item => {
      const code = generateDesignCode(item);
      const lengthM = item.lengthM || 0;
      const qty = item.quantity || 1;
      // 길이와 수량 모두 표시 (예: (110)12T마블아이보리 3m x1)
      let result = code;
      if (lengthM > 0) {
        result += ` ${lengthM}m`;
      }
      result += ` x${qty}`;
      return result;
    }).join(' / ');
    const deliveryMemo = (box.deliveryMemos || []).join(' / ');
    const fee = calculateShippingFee(box, shippingFees);

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
      box.isCombined ? 'Y' : 'N',
      box.remark || ''
    ];
    data.push(row);
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

function getKyungdongData() {
  const headers = ['받는분', '주소', '상세주소', '운송장번호', '고객사주문번호', '우편번호', '도착영업소', '전화번호', '기타전화번호', '선불후불', '품목명', '수량', '포장상태', '가로', '세로', '높이', '무게', '개별단가', '배송운임', '메모'];
  const data = [headers];

  const kyungdongOrders = packedOrders.filter(box => box.courier === 'kyungdong');

  kyungdongOrders.forEach(box => {
    const { base, detail } = splitAddress(box.address);
    const dims = getDimensions(box);
    const fee = calculateShippingFee(box, shippingFees);

    const row = [
      box.recipientLabel || box.customerName || '',
      base,
      (box.deliveryMemos || []).join(' / '),
      '',
      box.group?.orderId || '',
      box.zipCode || '',
      '',
      box.phone || '',
      '',
      '선불',
      box.designText || '',
      1,
      box.packagingType === 'vinyl' ? '비닐' : '박스',
      dims.width,
      dims.depth,
      dims.height,
      5,
      50,
      fee,
      box.designText || ''
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
function exportToKyungdong() {
  downloadExcel(getKyungdongData(), '경동발주서');
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
window.exportToKyungdong = exportToKyungdong;
window.exportToLogen = exportToLogen;
window.exportToShipping = exportToShipping;
window.exportToCombined1 = exportToCombined1;
window.exportToCombined2 = exportToCombined2;

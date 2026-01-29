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
  normalizeAddressForKyungdong,
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
let productDbPage, productDbTableBody;
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
  productDbPage = document.getElementById('productDbPage');
  productDbTableBody = document.getElementById('productDbTableBody');

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

  // 상품 DB 테이블 렌더링
  renderProductDbTable();

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
  const exportKyungdongFinalBtn = document.getElementById('exportKyungdongFinalBtn');
  const exportLogenBtn = document.getElementById('exportLogenBtn');
  const exportShippingBtn = document.getElementById('exportShippingBtn');

  if (exportCombined1Btn) exportCombined1Btn.addEventListener('click', exportToCombined1);
  if (exportCombined2Btn) exportCombined2Btn.addEventListener('click', exportToCombined2);
  if (exportKyungdongFinalBtn) exportKyungdongFinalBtn.addEventListener('click', exportToKyungdongFinal);
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
  const tables = ['rawTable', 'combined1Table', 'combined2Table', 'kyungdongFinalTable', 'logenTable', 'shippingTable', 'shippingFeesTable'];
  
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
      try {
        const data = new Uint8Array(e.target.result);
        const workbook = XLSX.read(data, { type: 'array' });
        if (!workbook.SheetNames || workbook.SheetNames.length === 0) {
          throw new Error('PasswordProtectedOrEmpty');
        }
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]];
        if (!firstSheet || !firstSheet['!ref']) {
          throw new Error('PasswordProtectedOrEmpty');
        }
        const csv = XLSX.utils.sheet_to_csv(firstSheet);
        parseCSV(csv);
      } catch (err) {
        const msg = '비밀번호가 걸려 있는 파일은 업로드할 수 없습니다.\n비밀번호를 제거한 뒤 다시 업로드해 주세요.\n\n【비밀번호 제거 방법】\n1. 엑셀에서 해당 파일을 연다(비밀번호 입력 후 열기)\n2. [파일] → [다른 이름으로 저장]\n3. 저장 창에서 [도구] → [일반 옵션]\n4. "열기 비밀번호" / "쓰기 비밀번호" 칸을 비운 뒤 [확인]\n5. 저장 후 새 파일을 업로드한다';
        alert(msg);
      }
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

/**
 * 네이버 주문 원본 헤더인지 판별.
 * 네이버 다운로드 파일은 1행에 '엑셀 일괄발송' 안내 문구가 올 수 있음(※ 1행 삭제 후 업로드 부탁 등).
 * 그 경우 2행이 실제 헤더(상품주문번호, 주문번호, 수취인명...)이므로 2행부터 헤더로 사용.
 */
function isNaverOrderHeader(row) {
  if (!row || !Array.isArray(row) || row.length === 0) return false;
  const joined = row.map(c => String(c || '').trim()).join(' ');
  // 1행 설명 문구에 흔히 포함되는 문구 → 헤더가 아님
  if (/엑셀\s*일괄발송|1행\s*삭제|다운로드\s*받은\s*파일로/.test(joined)) return false;
  const required = ['상품주문번호', '주문번호', '수취인명', '구매자명', '상품번호', '상품명'];
  const hasEnough = required.filter(k => joined.includes(k)).length >= 2;
  return hasEnough;
}

function parseCSV(text) {
  // BOM 제거
  if (text.charCodeAt(0) === 0xFEFF) {
    text = text.slice(1);
  }

  const rows = parseCSVToRows(text);

  if (!rows || rows.length < 2) {
    console.error('CSV 파싱 실패: rows=', rows);
    alert('CSV 파일 형식이 올바르지 않습니다.');
    return;
  }

  // 1행이 헤더가 아니면(네이버 원본의 설명 행 등) 2행부터 헤더로 사용
  let headerRowIndex = 0;
  if (!isNaverOrderHeader(rows[0]) && rows.length >= 3 && isNaverOrderHeader(rows[1])) {
    headerRowIndex = 1;
  }

  const headers = rows[headerRowIndex].map(h => h ? h.trim() : '');
  if (!headers.some(h => h)) {
    alert('CSV 파일 형식이 올바르지 않습니다. 헤더 행을 찾을 수 없습니다.');
    return;
  }

  ordersData = [];
  rawOrderData = [];

  for (let i = headerRowIndex + 1; i < rows.length; i++) {
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
  packedOrders = processPacking(groupedOrders, generateDesignCode, shippingFees);

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
  } else if (tabId === 'kyungdongFinal') {
    document.getElementById('kyungdongFinalSection').style.display = 'block';
    renderKyungdongFinalTable();
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
    if (productDbPage) productDbPage.style.display = 'none';
  } else if (page === 'shipping-fees') {
    if (ordersPage) ordersPage.style.display = 'none';
    if (shippingFeesPage) shippingFeesPage.style.display = 'block';
    if (productDbPage) productDbPage.style.display = 'none';
    renderShippingFeesTable();
  } else if (page === 'product-db') {
    if (ordersPage) ordersPage.style.display = 'none';
    if (shippingFeesPage) shippingFeesPage.style.display = 'none';
    if (productDbPage) productDbPage.style.display = 'block';
    renderProductDbTable();
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
  packedOrders = processPacking(groupedOrders, generateDesignCode, shippingFees);
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

  shippingFees.forEach((fee, index) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${index + 1}</td>
      <td>${fee.productGroup}</td>
      <td>${fee.packageType}</td>
      <td>${formatCurrency(fee.fee)}</td>
      <td>${fee.width || '-'}</td>
      <td>${fee.thickness || '-'}</td>
      <td>${fee.lengthMin || '-'}</td>
      <td>${fee.lengthMax || '-'}</td>
      <td></td>
    `;
    shippingFeesTableBody.appendChild(tr);
  });

  // 리사이저 다시 설정
  const shippingFeesTable = document.getElementById('shippingFeesTable');
  if (shippingFeesTable) {
    setupTableResizers(shippingFeesTable);
  }
}

function renderProductDbTable() {
  if (!productDbTableBody) return;
  productDbTableBody.innerHTML = '';

  if (productDb.length === 0) {
    productDbTableBody.innerHTML = '<tr><td colspan="6" style="text-align: center; padding: 2rem;">상품 DB 데이터를 로드 중...</td></tr>';
    return;
  }

  productDb.forEach((product, index) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${index + 1}</td>
      <td>${product.productId}</td>
      <td>${product.optionInfo}</td>
      <td>${product.designCode}</td>
      <td>${product.lengthNum || '-'}</td>
      <td>${product.lengthM || '-'}</td>
    `;
    productDbTableBody.appendChild(tr);
  });

  // 리사이저 다시 설정
  const productDbTable = document.getElementById('productDbTable');
  if (productDbTable) {
    setupTableResizers(productDbTable);
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

    // _customerKey 메타데이터로 고객 구분 (동명이인 구분)
    const customerKey = row._customerKey || '';
    if (customerKey && customerKey !== currentCustomer) {
      currentCustomer = customerKey;
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

function renderKyungdongFinalTable() {
  renderKyungdongTableCommon('kyungdongFinalTable', 'kyungdongFinalTableBody', getKyungdongFinalData);
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

    // _customerKey 메타데이터로 고객 구분 (동명이인 구분)
    const customerKey = row._customerKey || '';
    if (customerKey && customerKey !== currentCustomer) {
      currentCustomer = customerKey;
      groupIndex++;
      console.log(`[발송처리] 새 고객: ${customerKey}, groupIndex: ${groupIndex}`);
    }

    // 고객별 교차 색상 적용
    const className = groupIndex % 2 === 0 ? 'raw-row-group-even' : 'raw-row-group-odd';
    tr.classList.add(className);

    if (i <= 3) {
      console.log(`[발송처리] 행${i}: ${customerKey}, 클래스: ${className}`);
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

// ========== 경동 발주서 공통 헬퍼 함수 ==========
/**
 * 고객별 박스 그룹화
 */
function groupKyungdongBoxesByCustomer(kyungdongOrders) {
  const customerGroups = new Map();
  kyungdongOrders.forEach(box => {
    const customerName = box.customerName || box.recipientLabel || '';
    const customerPhone = box.phone || '';
    const customerAddress = box.address || '';

    // 동명이인 구분을 위해 이름+전화번호+주소 조합으로 키 생성
    const customerKey = `${customerName}|${customerPhone}|${customerAddress}`;
    if (!customerName) return; // 고객명이 없으면 스킵

    if (!customerGroups.has(customerKey)) {
      customerGroups.set(customerKey, []);
    }
    customerGroups.get(customerKey).push(box);
  });
  return customerGroups;
}

/**
 * 고객별 총 결제금액, 롤매트 길이, 퍼즐매트 포함 여부 계산
 */
function calculateCustomerTotals(boxes) {
  const rollMatProductIds = ['6092903705', '6626596277', '4200445704'];
  const puzzleMatProductId = '5994906898';
  
  let totalPayment = 0;
  let totalRollLength = 0;
  let hasRollOrPuzzle = false;

  boxes.forEach(box => {
    if (!box.items || !Array.isArray(box.items)) return;
    
    box.items.forEach(item => {
      const itemRawRow = item.rawRow || {};
      const payment = parsePrice(itemRawRow['최종 상품별 총 주문금액']) || item.price || 0;
      totalPayment += payment;

      const productId = item.rawRow?.['상품번호'] || '';
      if (rollMatProductIds.includes(productId)) {
        const lengthM = item.lengthM || 0;
        const qty = item.quantity || 1;
        totalRollLength += lengthM * qty;
        hasRollOrPuzzle = true;
      }
      if (productId === puzzleMatProductId) {
        hasRollOrPuzzle = true;
      }
    });
  });

  return { totalPayment, totalRollLength, hasRollOrPuzzle };
}

/**
 * 증정품 계산 (테이프, 팻말)
 */
function calculateGiftsForCustomer(totalRollLength, totalPayment, hasRollOrPuzzle) {
  const gifts = [];

  // 테이프 증정 (100,000원 이상, 결제금액 기준 수량)
  if (totalPayment >= 100000) {
    let tapeCount;
    if (totalPayment >= 500000) {
      tapeCount = 3;
    } else if (totalPayment >= 300000) {
      tapeCount = 2;
    } else {
      tapeCount = 1;
    }
    gifts.push(`★증정★테이프20m x${tapeCount}`);
  }

  // 팻말 증정 (롤매트 or 퍼즐매트 + 300,000원 이상)
  if (hasRollOrPuzzle && totalPayment >= 300000) {
    gifts.push('★팻말 증정★ x1');
  }

  return gifts.join(' / ');
}

/**
 * 박스의 아이템들을 디자인코드별로 그룹핑
 */
function groupItemsByDesign(box) {
  if (!box.items || !Array.isArray(box.items)) return [];

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

  return groupedItems;
}

/**
 * 그룹핑된 아이템으로 텍스트 생성
 */
function createItemTexts(groupedItems) {
  return groupedItems.map(itemData => {
    let result = itemData.designCode || '';
    const isTape = itemData.designCode?.includes('테이프') || false;
    if (itemData.lengthM > 0 && !isTape) {
      result += ` ${itemData.lengthM}m`;
    }
    result += ` x${itemData.qty || 1}`;
    return {
      text: result,
      designCode: itemData.designCode,
      lengthM: itemData.lengthM,
      qty: itemData.qty
    };
  });
}

function getCombined1Data() {
  return getCombinedDataCommon(false);
}

function getCombined2Data() {
  const headers = ['상품주문번호', '수취인명', '운임타입', '송장수량', '디자인', '길이', '수량', '디자인+수량', '배송메세지', '합포장', '비고', '주문하신분 핸드폰', '받는분 핸드폰', '우편번호', '받는분 주소', '주문자명', '실제결제금액'];
  const MAX_DESIGN_WITH_QTY_LENGTH = 30; // 디자인+수량 컬럼 임계값 (박혜정 36자, 김태종 48자 → 분할)

  // 고객별로 박스 그룹화 및 총 결제금액 계산
  const customerGroups = new Map();
  // 주문번호별 결제금액 추적 (중복 계산 방지)
  const orderPaymentMap = new Map();
  const rollMatProductIds = ['6092903705', '6626596277', '4200445704'];
  const puzzleMatProductId = '5994906898';
  
  packedOrders.forEach(box => {
    const customerName = box.customerName || box.recipientLabel || '';
    const customerPhone = box.phone || '';
    const customerAddress = box.address || '';
    const customerKey = `${customerName}|${customerPhone}|${customerAddress}`;

    if (!customerGroups.has(customerKey)) {
      customerGroups.set(customerKey, {
        boxes: [],
        totalPayment: 0,
        totalRollLength: 0,
        hasRollOrPuzzle: false
      });
    }
    const group = customerGroups.get(customerKey);
    group.boxes.push(box);

    // 이 박스의 결제금액 합산 (주문번호별로 중복 계산 방지)
    box.items.forEach(item => {
      const itemRawRow = item.rawRow || {};
      const orderId = itemRawRow['상품주문번호'] || item.id || '';
      const payment = parsePrice(itemRawRow['최종 상품별 총 주문금액']) || item.price || 0;
      
      // 같은 주문번호의 결제금액은 한 번만 계산
      if (orderId && !orderPaymentMap.has(orderId)) {
        orderPaymentMap.set(orderId, payment);
        group.totalPayment += payment;
      } else if (!orderId) {
        // 주문번호가 없으면 그냥 합산 (중복 가능성 있지만 어쩔 수 없음)
        group.totalPayment += payment;
      }
      
      // 고객 전체 롤매트 길이 계산
      const productId = itemRawRow['상품번호'] || '';
      if (rollMatProductIds.includes(productId)) {
        group.totalRollLength += (item.lengthM || 0) * (item.quantity || 1);
        group.hasRollOrPuzzle = true;
      }
      if (productId === puzzleMatProductId) {
        group.hasRollOrPuzzle = true;
      }
    });
  });

  // 고객별 첫 행 추적
  const customerFirstRowFlags = new Map();
  
  // 정렬을 위한 행 데이터 저장 (제품번호, 폭 정보 포함)
  const rowsWithSortInfo = [];

  packedOrders.forEach(box => {
    const designCode = box.designText || '';
    const quantity = box.totalQtyInBox || 1;

    // 원본 데이터에서 추가 정보 가져오기
    const firstItem = box.group?.items?.[0];
    const rawRow = firstItem?.rawRow || {};
    const buyerPhone = rawRow['구매자연락처'] || '';
    const buyerName = rawRow['구매자명'] || '';

    // 고객별 총 결제금액 가져오기 (동일한 키 생성 방식 사용)
    const customerName = box.customerName || box.recipientLabel || '';
    const customerPhone = box.phone || '';
    const customerAddress = box.address || '';
    const customerKey = `${customerName}|${customerPhone}|${customerAddress}`;
    const customerGroup = customerGroups.get(customerKey);
    const formattedPayment = customerGroup ? customerGroup.totalPayment.toLocaleString('ko-KR') : '';

    // 이 고객의 첫 행인지 확인
    const isCustomerFirstRow = !customerFirstRowFlags.has(customerKey);
    if (isCustomerFirstRow) {
      customerFirstRowFlags.set(customerKey, true);
    }

    // 증정품 계산을 위한 고객 전체 롤매트 길이와 결제금액
    const totalPayment = customerGroup ? customerGroup.totalPayment : 0;
    const totalRollLength = customerGroup ? customerGroup.totalRollLength : 0;
    const hasRollOrPuzzle = customerGroup ? customerGroup.hasRollOrPuzzle : false;

    // 배송메세지 생성 (파손주의 표시 추가)
    let deliveryMemo = (box.deliveryMemos || []).join(' / ');
    const hasFragileMarking = (box.recipientLabel || box.customerName || '').includes('★');
    if (hasFragileMarking) {
      deliveryMemo = deliveryMemo ? `★파손주의★ ${deliveryMemo}` : '★파손주의★';
    }

    // 증정품 계산 (고객의 첫 번째 박스에만 표시)
    const gifts = [];
    if (isCustomerFirstRow) {
    // 테이프 증정 (100,000원 이상, 결제금액 기준 수량)
    if (totalPayment >= 100000) {
        let tapeCount;
        if (totalPayment >= 500000) {
          tapeCount = 3;
        } else if (totalPayment >= 300000) {
          tapeCount = 2;
        } else {
          tapeCount = 1;
        }
        gifts.push(`★증정★테이프20m x${tapeCount}`);
      }

      // 팻말 증정 (롤매트 or 퍼즐매트 + 300,000원 이상) - 한 번만 표시
      if (hasRollOrPuzzle && totalPayment >= 300000) {
        gifts.push('★팻말 증정★ x1');
      }
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
      // 정렬을 위한 제품번호와 폭 정보 포함
      const productId = itemData.item?.rawRow?.['상품번호'] || itemData.item?.productId || '';
      const widthNum = itemData.item?.widthNum || 0;
      return {
        text: result,
        designCode: itemData.designCode,
        lengthM: itemData.lengthM,
        qty: itemData.qty,
        productId: productId,
        widthNum: widthNum,
        isTape: isTape
      };
    });

    const designWithQty = itemTexts.map(it => it.text).join(' / ');
    const hasTapeItem = itemTexts.some(it => it.isTape);

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
    // 또는 테이프가 다른 상품과 함께 있을 때는 항상 분할해서 행을 나눠줌
    if ((designWithQty.length >= MAX_DESIGN_WITH_QTY_LENGTH && groupedItems.length > 1) ||
        (hasTapeItem && groupedItems.length > 1)) {
      console.log(`[송장화2 분할] ${box.customerName}: ${groupedItems.length}개 그룹 분할`);

      // 분할 로직: 아이템별로 행 생성
      itemTexts.forEach((itemText, index) => {
        const isFirstRow = index === 0;
        const isTape = itemText.designCode.includes('테이프');
        // 이 박스의 첫 행이면서 고객의 첫 행인 경우에만 실제결제금액 표시
        const shouldShowPayment = isFirstRow && isCustomerFirstRow;
        // 정렬을 위한 제품번호와 폭 정보 추출
        const productId = itemText.productId || '';
        const widthNum = itemText.widthNum || 0;
        
        const row = [
          box.group?.orderId || '',
          box.recipientLabel || box.customerName || '',
          isFirstRow ? fee : '',
          isFirstRow ? 1 : '',
          itemText.designCode,
          (itemText.lengthM && !isTape) ? `${itemText.lengthM}m` : '',
          itemText.qty,
          itemText.text,
          deliveryMemo,
          box.isCombined ? '합' : '',
          giftText,
          buyerPhone,
          box.phone || '',
          box.zipCode || '',
          box.address || '',
          buyerName,
          shouldShowPayment ? formattedPayment : ''
        ];
        // 정렬 정보와 함께 저장 (고객 키 포함)
        rowsWithSortInfo.push({
          row: row,
          customerKey: customerKey,
          productId: productId,
          widthNum: widthNum
        });
      });
    } else {
      // 기존 로직: 단일 행 (고객의 첫 행인 경우에만 실제결제금액 표시)
      // 정렬을 위한 제품번호와 폭 정보 추출
      const firstItem = box.items[0];
      const productId = firstItem?.rawRow?.['상품번호'] || firstItem?.productId || '';
      const widthNum = firstItem?.widthNum || 0;
      
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
      // 정렬 정보와 함께 저장 (고객 키 포함)
      rowsWithSortInfo.push({
        row: row,
        customerKey: customerKey,
        productId: productId,
        widthNum: widthNum
      });
    }
  });

  // 고객별로 먼저 그룹화, 그 다음 제품번호, 폭순으로 정렬
  rowsWithSortInfo.sort((a, b) => {
    const customerKeyA = a.customerKey || '';
    const customerKeyB = b.customerKey || '';
    const productIdA = a.productId || '';
    const productIdB = b.productId || '';
    const widthA = a.widthNum || 0;
    const widthB = b.widthNum || 0;
    
    // 고객별로 먼저 그룹화
    if (customerKeyA !== customerKeyB) {
      return customerKeyA.localeCompare(customerKeyB);
    }
    // 같은 고객이면 제품번호로 정렬
    if (productIdA !== productIdB) {
      return productIdA.localeCompare(productIdB);
    }
    // 같은 제품이면 폭으로 정렬
    return widthA - widthB;
  });
  
  // 정렬된 행만 추출
  const sortedRows = rowsWithSortInfo.map(item => item.row);
  
  // 헤더와 정렬된 데이터 합치기
  return [headers, ...sortedRows];
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

    // 고객 키 생성 (동명이인 구분용)
    const customerName = box.customerName || box.recipientLabel || '';
    const customerPhone = box.phone || '';
    const customerAddress = box.address || '';
    const customerKey = `${customerName}|${customerPhone}|${customerAddress}`;

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
      // 행에 고객 키 메타데이터 추가 (렌더링 시 색상 구분용)
      row._customerKey = customerKey;
      data.push(row);
    });
  });

  return data;
}

// 경동1, 2, 3 함수 제거됨 - 경동발주서(최종)만 사용

// 경동발주서(최종) - 리팩토링된 버전
function getKyungdongFinalData() {
  const headers = ['받는분', '주소', '상세주소', '운송장번호', '고객사주문번호', '우편번호', '도착영업소', '전화번호', '기타전화번호', '선불후불', '품목명', '수량', '포장상태', '가로', '세로', '높이', '무게', '개별단가', '배송운임', '기타운임', '별도운임', '할증운임', '도서운임', '메모'];
  const data = [headers];

  const kyungdongOrders = packedOrders.filter(box => box.courier === 'kyungdong');
  if (!kyungdongOrders || kyungdongOrders.length === 0) {
  return data;
}

  // 고객별로 박스 그룹화
  const customerGroups = groupKyungdongBoxesByCustomer(kyungdongOrders);

  // 고객별로 처리
  customerGroups.forEach((boxes) => {
    if (!boxes || boxes.length === 0) return;

    // 고객별 총 결제금액 및 롤매트 길이 계산 (증정품 계산용)
    const { totalPayment, totalRollLength, hasRollOrPuzzle } = calculateCustomerTotals(boxes);

    // 증정품 계산
    const giftText = calculateGiftsForCustomer(totalRollLength, totalPayment, hasRollOrPuzzle);

    // 각 박스별로 행 생성
    boxes.forEach((box, boxIndex) => {
      if (!box) return;

      const dims = getDimensions(box) || {};
      const fee = calculateShippingFee(box, shippingFees) || '';
      const isFirstBox = boxIndex === 0;
      const MAX_DESIGN_LENGTH = 30;

      // 아이템 그룹핑 및 텍스트 생성
      const groupedItems = groupItemsByDesign(box);
      const itemTexts = createItemTexts(groupedItems);
      const designWithQty = itemTexts.map(it => it.text).join(' / ');
      
      // 메모 텍스트 생성: 증정품만
      const memoText = isFirstBox ? (giftText || '') : '';

      // 주소 처리: 주소와 상세주소를 합쳐서 주소 컬럼에 넣기 (경동 발주서용 정규화 적용)
      const fullAddress = normalizeAddressForKyungdong(box.address || '');
      const recipientName = box.recipientLabel || box.customerName || '';
      const orderId = box.group?.orderId || '';
      
      // 상세주소 컬럼: 파손주의 + 배송메모(재단요청 등)
      const hasFragileMarking = box.needsStar || (box.recipientLabel || '').includes('★');
      const deliveryMemosRaw = (box.deliveryMemos || []).join(' / ');
      const deliveryMemos = hasFragileMarking 
        ? (deliveryMemosRaw ? `★파손주의★ / ${deliveryMemosRaw}` : '★파손주의★')
        : deliveryMemosRaw;

      // 텍스트 길이 초과 여부 확인
      if (designWithQty.length >= MAX_DESIGN_LENGTH && groupedItems.length > 1) {
        // 분할 로직: 아이템별로 행 생성
        itemTexts.forEach((itemText, itemIndex) => {
          const isFirstItem = itemIndex === 0;
          const row = [
            isFirstItem ? recipientName : '', // 받는분
            isFirstItem ? fullAddress : '', // 주소 (전체 주소)
            isFirstItem ? deliveryMemos : '', // 상세주소 (배송메모: 재단요청 등)
            '', // 운송장번호
            isFirstItem ? orderId : '', // 고객사주문번호
            isFirstItem ? (box.zipCode || '') : '', // 우편번호
            '', // 도착영업소
            isFirstItem ? (box.phone || '') : '', // 전화번호
            isFirstItem ? (box.phone || '') : '', // 기타전화번호
            isFirstItem ? '선불' : '', // 선불후불
            isFirstItem ? (memoText || '따사룸') : '', // 품목명 (증정품 정보) - 첫 행만
            isFirstItem ? 1 : '', // 수량 - 첫 행만
            isFirstItem ? (box.packagingType === 'vinyl' ? '비닐' : '박스') : '', // 포장상태
            isFirstItem ? (dims.width || 1) : '', // 가로
            isFirstItem ? (dims.depth || 1) : '', // 세로
            isFirstItem ? (dims.height || 1) : '', // 높이
            isFirstItem ? 1 : '', // 무게
            isFirstItem ? 50 : '', // 개별단가
            isFirstItem ? fee : '', // 배송운임
            isFirstItem ? 100 : '', // 기타운임
            '', // 별도운임
            '', // 할증운임
            '', // 도서운임
            itemText.text || '' // 메모 (디자인코드+수량) - 모든 행에 표시
          ];
          data.push(row);
        });
      } else {
        // 기존 로직: 단일 행
        const row = [
          recipientName, // 받는분
          fullAddress, // 주소 (전체 주소)
          deliveryMemos, // 상세주소 (배송메모: 재단요청 등)
          '', // 운송장번호
          orderId, // 고객사주문번호
          box.zipCode || '', // 우편번호
          '', // 도착영업소
          box.phone || '', // 전화번호
          box.phone || '', // 기타전화번호
          '선불', // 선불후불
          memoText || '따사룸', // 품목명 (증정품 정보)
          1, // 수량
          box.packagingType === 'vinyl' ? '비닐' : '박스', // 포장상태
          dims.width || 1, // 가로
          dims.depth || 1, // 세로
          dims.height || 1, // 높이
          1, // 무게
          50, // 개별단가
          fee, // 배송운임
          100, // 기타운임
          '', // 별도운임
          '', // 할증운임
          '', // 도서운임
          designWithQty || box.designText || '따사룸' // 메모 (디자인코드+수량)
        ];
        data.push(row);
      }
    });
  });

  return data;
}


function getLogenData() {
  const headers = ['받는분', '받는분 전화번호', '받는분 핸드폰', '우편번호', '받는분 주소', '운임타입', '송장수량', '디자인', '배송메세지', '합배송여부'];
  const data = [headers];

  const logenOrders = packedOrders.filter(box => box.courier === 'logen');

  // 고객별로 박스 그룹화
  const customerGroups = new Map();
  logenOrders.forEach(box => {
    const customerName = box.customerName || box.recipientLabel || '';
    const customerPhone = box.phone || '';
    const customerAddress = box.address || '';

    // 동명이인 구분을 위해 이름+전화번호+주소 조합으로 키 생성
    const customerKey = `${customerName}|${customerPhone}|${customerAddress}`;
    if (!customerName) return;

    if (!customerGroups.has(customerKey)) {
      customerGroups.set(customerKey, []);
    }
    customerGroups.get(customerKey).push(box);
  });

  // 고객별로 처리
  customerGroups.forEach((boxes) => {
    if (!boxes || boxes.length === 0) return;

    // 고객별 총 결제금액 및 롤매트 길이 계산 (증정품 계산용)
    const rollMatProductIds = ['6092903705', '6626596277', '4200445704'];
    const puzzleMatProductId = '5994906898';

    let totalPayment = 0;
    let totalRollLength = 0;
    let hasRollOrPuzzle = false;

    boxes.forEach(box => {
      if (!box.items || !Array.isArray(box.items)) return;

      box.items.forEach(item => {
        const itemRawRow = item.rawRow || {};
        const payment = parsePrice(itemRawRow['최종 상품별 총 주문금액']) || item.price || 0;
        totalPayment += payment;

        const productId = item.rawRow?.['상품번호'] || '';
        if (rollMatProductIds.includes(productId)) {
          const lengthM = item.lengthM || 0;
          const qty = item.quantity || 1;
          totalRollLength += lengthM * qty;
          hasRollOrPuzzle = true;
        }
        if (productId === puzzleMatProductId) {
          hasRollOrPuzzle = true;
        }
      });
    });

    // 증정품 계산
    const gifts = [];
    // 테이프 증정 (100,000원 이상, 결제금액 기준 수량)
    if (totalPayment >= 100000) {
      let tapeCount;
      if (totalPayment >= 500000) {
        tapeCount = 3;
      } else if (totalPayment >= 300000) {
        tapeCount = 2;
      } else {
        tapeCount = 1;
      }
      gifts.push(`★증정★테이프20m x${tapeCount}`);
    }
    if (hasRollOrPuzzle && totalPayment >= 200000) {
      gifts.push('★팻말 증정★ x1');
    }

    // 각 박스별로 행 생성
    boxes.forEach((box, boxIndex) => {
      if (!box) return;

      const fee = calculateShippingFee(box, shippingFees);
      const isFirstBox = boxIndex === 0;
      const hasMultipleBoxes = boxes.length > 1;

      // 박스의 아이템들을 디자인코드별로 그룹핑
      const groupedItems = groupItemsByDesign(box);
      const itemTexts = createItemTexts(groupedItems);

      // 배송메모
      const deliveryMemos = (box.deliveryMemos || []).join(' / ');

      // 각 아이템별로 행 생성
      itemTexts.forEach((itemText, itemIndex) => {
        const isFirstItem = itemIndex === 0;
        const row = [
          isFirstItem ? (box.recipientLabel || box.customerName || '') : '',
          isFirstItem ? (box.phone || '') : '',
          isFirstItem ? (box.phone || '') : '',
          isFirstItem ? (box.zipCode || '') : '',
          isFirstItem ? (box.address || '') : '',
          isFirstItem ? fee : '',
          isFirstItem ? 1 : '',
          itemText.text,
          isFirstItem ? deliveryMemos : '',
          (hasMultipleBoxes && !isFirstBox) ? '합' : ''
        ];
        data.push(row);
      });

      // 증정품 행 추가 (첫 번째 박스만)
      if (isFirstBox && gifts.length > 0) {
        gifts.forEach(gift => {
          const row = [
            '',
            '',
            '',
            '',
            '',
            '',
            '',
            gift,
            deliveryMemos,
            hasMultipleBoxes ? '합' : ''
          ];
          data.push(row);
        });
      }
    });
  });

  return data;
}

function getShippingData() {
  const headers = ['상품주문번호', '수취인명', '배송지', '우편번호', '연락처', '상품명', '수량', '배송메세지', '택배사'];
  const data = [headers];

  packedOrders.forEach(box => {
    // 고객 키 생성 (동명이인 구분용)
    const customerName = box.customerName || box.recipientLabel || '';
    const customerPhone = box.phone || '';
    const customerAddress = box.address || '';
    const customerKey = `${customerName}|${customerPhone}|${customerAddress}`;

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
    // 행에 고객 키 메타데이터 추가 (렌더링 시 색상 구분용)
    row._customerKey = customerKey;
    data.push(row);
  });

  return data;
}

// ========== 엑셀 내보내기 ==========
function exportToKyungdongFinal() {
  downloadExcel(getKyungdongFinalData(), '경동발주서(최종)');
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
window.exportToKyungdongFinal = exportToKyungdongFinal;
window.exportToLogen = exportToLogen;
window.exportToShipping = exportToShipping;
window.exportToCombined1 = exportToCombined1;
window.exportToCombined2 = exportToCombined2;

// ========== 전역 상태 ==========
let ordersData = [];
let filteredOrders = [];
let groupedOrders = {}; // 고객별 그룹화된 주문
let packedOrders = []; // 패킹 처리된 결과
let rawOrderData = []; // 원본 CSV/Excel 데이터
let shippingFees = []; // 배송비 데이터
let currentTabId = 'raw';

// ========== 제품 및 발주 로직 상수 ==========
// 상품번호 → 제품타입 매핑
const PRODUCT_TYPE_MAP = {
  "6092903705": { type: "롤", name: "유아롤매트", category: "babyRoll" },
  "4200445704": { type: "롤", name: "애견롤매트", category: "petRoll" },
  "6626596277": { type: "롤", name: "롤매트", category: "roll" },
  "5994906898": { type: "퍼즐", name: "퍼즐매트", category: "puzzle" },
  "5994903887": { type: "퍼즐", name: "퍼즐매트", category: "puzzle" },
  "5569937047": { type: "TPU", name: "TPU매트", category: "tpu" },
  "11101602541": { type: "TPU", name: "TPU(B)", category: "tpu" },
  "282153807": { type: "벽지", name: "벽지", category: "wallpaper" },
  "561709916": { type: "벽지", name: "벽지", category: "wallpaper" },
  "4723369915": { type: "테이프", name: "테이프", category: "tape" },
  "6710830644": { type: "PE롤", name: "PE롤", category: "peRoll" }
};

// 퍼즐 상품번호 (비닐 포장 대상)
const PUZZLE_PRODUCT_IDS = ["5994906898", "5994903887"];

// 롤매트 상품번호
const ROLL_PRODUCT_IDS = ["6092903705", "4200445704", "6626596277", "6710830644"];

// 테이프 상품번호
const TAPE_PRODUCT_IDS = ["4723369915"];

// 두께별 포장 기준 (m 단위)
const PACKAGING_THRESHOLDS = {
  "6": { small: 8, large: 12, vinyl: 12.5 },
  "9": { small: 6, large: 10, vinyl: 10.5 },
  "10": { small: 5, large: 9, vinyl: 9.5 }, // 10T 추가
  "12": { small: 3.5, large: 8, vinyl: 8.5 },
  "15": { small: 1, large: 7, vinyl: 8 },
  "17": { small: 3, large: 7, vinyl: 8 },
  "22": { small: 1, large: 3, vinyl: 3.5 }
};

// 퍼즐매트 박스당 최대 수량
const PUZZLE_BOX_CAPACITY = {
  "25": 6,
  "40": 4   // 사용자 확인: 40mm는 4개가 맞음
};

// 재단 요청 감지 키워드
const CUTTING_KEYWORDS = [
  "재단", "커팅", "컷팅", "절단", "잘라", "짤라", "컷트", "사이즈", "규격", "크기",
  "가로", "세로", "폭", "높이", "길이", "cm", "mm", "센치", "센티미터", "메터",
  "메타", "메다", "밀리", "등분", "조각", "나누어", "나눠서", "분할", "절개", "재난",
  "짜르다", "자르다", "쟐라", "쨀라", "반씩", "반으로", "잘게", "자르기", "절단기",
  "컷팅기", "가름", "가릅", "통째", "토막"
];

// 마감재 요청 감지 키워드
const FINISHING_KEYWORDS = ["마감재", "마감제", "I자형", "L자형", "ㄱ자", "ㄴ자"];

// 색상 축약 규칙
const COLOR_ABBREVIATIONS = {
  "바닐라아이보리": "아이보리",
  "바닐라 아이보리": "아이보리",
  "라이트그레이": "그레이",
  "라이트 그레이": "그레이",
  "스카이블루": "블루",
  "스카이 블루": "블루",
  "파스텔핑크": "핑크",
  "파스텔 핑크": "핑크",
  "퓨어아이보리": "퓨어",
  "퓨어 아이보리": "퓨어",
  "마블아이보리": "마블",
  "마블 아이보리": "마블",
  "그레이캔버스": "그캔",
  "그레이 캔버스": "그캔",
  "딜라이트우드": "우드",
  "딜라이트 우드": "우드",
  "비글퍼비": "비글",
  "비글 퍼비": "비글",
  "도그프린드": "도그",
  "도그 프린드": "도그",
  "도그프렌드": "도그",
  "소프트그레이": "소그",
  "소프트 그레이": "소그",
  "테리어": "테리",
  "크림아이보리": "크림",
  "크림 아이보리": "크림",
  "소프트아이보리": "소프트",
  "소프트 아이보리": "소프트",
  "모던그레이": "모던그레",
  "스노우화이트": "스노우",
  "내추럴우드": "내추럴"
};

// ========== DOM 요소 ==========
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const uploadBtn = document.getElementById('uploadBtn');
const testDataBtn = document.getElementById('testDataBtn');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const clearBtn = document.getElementById('clearBtn');

const summarySection = document.getElementById('summarySection');
const filterSection = document.getElementById('filterSection');
const ordersSection = document.getElementById('ordersSection');
const ordersTableBody = document.getElementById('ordersTableBody');

const kyungdongCount = document.getElementById('kyungdongCount');
const lozenCount = document.getElementById('lozenCount');
const totalCustomers = document.getElementById('totalCustomers');
const giftOrders = document.getElementById('giftOrders');

const productTypeFilter = document.getElementById('productTypeFilter');
const orderStatusFilter = document.getElementById('orderStatusFilter');
const giftFilter = document.getElementById('giftFilter');
const searchInput = document.getElementById('searchInput');
const selectAll = document.getElementById('selectAll');

const orderModal = document.getElementById('orderModal');
const modalClose = document.getElementById('modalClose');
const modalBody = document.getElementById('modalBody');

const sidebarItems = document.querySelectorAll('.sidebar-item');
const shippingFeesPage = document.getElementById('shippingFeesPage');
const shippingFeesTableBody = document.getElementById('shippingFeesTableBody');

// ========== 초기화 ==========
document.addEventListener('DOMContentLoaded', () => {
  initEventListeners();
  updateDateDisplay();
  loadShippingFees(); // 배송비 데이터 로드
  switchPage('orders'); // 초기 페이지 설정
});

function initEventListeners() {
  // 파일 업로드
  uploadBtn.addEventListener('click', () => fileInput.click());
  uploadArea.addEventListener('click', (e) => {
    if (e.target !== uploadBtn) fileInput.click();
  });
  fileInput.addEventListener('change', handleFileSelect);
  clearBtn.addEventListener('click', clearFile);

  // 테스트 데이터 버튼
  if (testDataBtn) {
    testDataBtn.addEventListener('click', loadTestData);
  }

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
  productTypeFilter.addEventListener('change', applyFilters);
  orderStatusFilter.addEventListener('change', applyFilters);
  giftFilter.addEventListener('change', applyFilters);
  searchInput.addEventListener('input', applyFilters);

  // 전체 선택 (체크박스가 있는 경우에만)
  if (selectAll) {
    selectAll.addEventListener('change', handleSelectAll);
  }

  // 모달
  if (modalClose) {
    modalClose.addEventListener('click', (e) => {
      e.preventDefault();
      e.stopPropagation();
      closeModal();
    });
  }
  if (orderModal) {
    orderModal.addEventListener('click', (e) => {
      if (e.target === orderModal) closeModal();
    });
  }

  // 발주서 다운로드 버튼
  const exportCombinedBtn = document.getElementById('exportCombinedBtn');
  const exportKyungdongBtn = document.getElementById('exportKyungdongBtn');
  const exportLozenBtn = document.getElementById('exportLozenBtn');

  if (exportCombinedBtn) {
    exportCombinedBtn.addEventListener('click', () => exportToCombined());
  }
  if (exportKyungdongBtn) {
    exportKyungdongBtn.addEventListener('click', () => exportToKyungdong());
  }
  if (exportLozenBtn) {
    exportLozenBtn.addEventListener('click', () => exportToLozen());
  }

  // 사이드바 네비게이션
  sidebarItems.forEach(item => {
    item.addEventListener('click', () => {
      const page = item.getAttribute('data-page');
      switchPage(page);
    });
  });
}

function updateDateDisplay() {
  const now = new Date();
  const options = { year: 'numeric', month: 'long', day: 'numeric', weekday: 'long' };
  document.getElementById('currentDate').textContent = now.toLocaleDateString('ko-KR', options);
}

// ========== 파일 처리 ==========
function handleFileSelect(e) {
  const file = e.target.files[0];
  if (file) processFile(file);
}

function handleDrop(e) {
  e.preventDefault();
  uploadArea.classList.remove('dragover');
  const file = e.dataTransfer.files[0];
  if (file) processFile(file);
}

// 테스트용 CSV 파일 로드 (/OredrOps_Test.csv)
async function loadTestData() {
  try {
    const response = await fetch('OredrOps_Test.csv');
    if (!response.ok) {
      alert('테스트 CSV 파일을 불러오지 못했습니다.');
      return;
    }

    const text = await response.text();
    // 기존 CSV 파서 재사용
    parseCSV(text);

    // UI 업데이트: 파일명 표시
    if (fileName && fileInfo) {
      fileName.textContent = 'OredrOps_Test.csv (테스트)';
      fileInfo.style.display = 'flex';
    }
  } catch (error) {
    console.error('테스트 데이터 로드 오류:', error);
    alert('테스트 데이터를 불러오는 중 오류가 발생했습니다.');
  }
}

function processFile(file) {
  const extension = file.name.split('.').pop().toLowerCase();

  if (!['csv', 'xlsx', 'xls'].includes(extension)) {
    alert('CSV 또는 Excel 파일만 업로드 가능합니다.');
    return;
  }

  fileName.textContent = file.name;
  fileInfo.style.display = 'flex';

  if (extension === 'csv') {
    readCSV(file);
  } else {
    readExcel(file);
  }
}

function readCSV(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const text = e.target.result;
      parseCSV(text);
    } catch (error) {
      console.error('CSV 파싱 오류:', error);
      alert('CSV 파일 읽기에 실패했습니다: ' + error.message);
    }
  };
  reader.onerror = (e) => {
    console.error('파일 읽기 오류:', e);
    alert('파일 읽기에 실패했습니다.');
  };
  reader.readAsText(file, 'UTF-8');
}

function readExcel(file) {
  const reader = new FileReader();
  reader.onload = (e) => {
    try {
      const data = new Uint8Array(e.target.result);
      const workbook = XLSX.read(data, { type: 'array' });

      // 첫 번째 시트 또는 '발송처리' 시트 사용
      let sheetName = workbook.SheetNames[0];
      if (workbook.SheetNames.includes('발송처리')) {
        sheetName = '발송처리';
      }

      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1 });

      parseExcelData(jsonData);
    } catch (error) {
      console.error('Excel 파싱 오류:', error);
      alert('Excel 파일 읽기에 실패했습니다.');
    }
  };
  reader.readAsArrayBuffer(file);
}

function parseExcelData(data) {
  if (data.length < 2) {
    alert('데이터가 없습니다.');
    return;
  }

  const headers = data[0];
  ordersData = [];
  rawOrderData = []; // 원본 데이터 초기화

  for (let i = 1; i < data.length; i++) {
    const values = data[i];
    if (!values || values.length === 0) continue;

    const row = {};
    headers.forEach((header, index) => {
      row[header] = values[index] !== undefined ? String(values[index]).trim() : '';
    });

    // 원본 데이터 저장
    rawOrderData.push(row);

    const order = parseOrderData(row);
    if (order) ordersData.push(order);
  }

  filteredOrders = [...ordersData];
  updateUI();
}

/**
 * CSV 텍스트를 rows 배열로 파싱 (줄바꿈/콤마가 포함된 큰따옴표 필드 대응)
 */
function parseCSVToRows(text) {
  const rows = [];
  let currentRow = [];
  let currentField = '';
  let inQuotes = false;

  for (let i = 0; i < text.length; i++) {
    const char = text[i];
    const nextChar = text[i + 1];

    if (inQuotes) {
      if (char === '"' && nextChar === '"') {
        currentField += '"';
        i++; // 다음 " 건너뜀
      } else if (char === '"') {
        inQuotes = false;
      } else {
        currentField += char;
      }
    } else {
      if (char === '"') {
        inQuotes = true;
      } else if (char === ',') {
        currentRow.push(currentField.trim());
        currentField = '';
      } else if (char === '\n' || char === '\r') {
        currentRow.push(currentField.trim());
        if (currentRow.some(field => field !== '')) {
          rows.push(currentRow);
        }
        currentRow = [];
        currentField = '';
        if (char === '\r' && nextChar === '\n') {
          i++; // \r\n 대응
        }
      } else {
        currentField += char;
      }
    }
  }

  // 마지막 남은 데이터 처리
  if (currentRow.length > 0 || currentField !== '') {
    currentRow.push(currentField.trim());
    if (currentRow.some(field => field !== '')) {
      rows.push(currentRow);
    }
  }

  return rows;
}

/**
 * CSV 파싱 (줄바꿈/콤마가 포함된 큰따옴표 필드 대응)
 */
function parseCSV(text) {
  // BOM 제거 (UTF-8 BOM: \uFEFF)
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
  rawOrderData = []; // 원본 데이터 초기화

  for (let i = 1; i < rows.length; i++) {
    const rowValues = rows[i];
    const rowObj = {};
    headers.forEach((header, index) => {
      rowObj[header] = rowValues[index] || '';
    });

    // 원본 데이터 저장
    rawOrderData.push(rowObj);

    const order = parseOrderData(rowObj);
    if (order) ordersData.push(order);
  }

  filteredOrders = [...ordersData];
  updateUI();
}


function parseOrderData(row) {
  try {
    // 필수 데이터 검증 (상품주문번호나 상품명이 없으면 로드하지 않음/빈 행 제외)
    if (!row['상품주문번호'] || !row['상품명']) {
      return null;
    }

    const optionInfo = row['옵션정보'] || '';
    const productName = row['상품명'] || '';
    const parsed = parseOptionInfo(optionInfo);
    const parsedFromName = parseProductName(productName);
    const productId = row['상품번호'] || '';
    const deliveryMemo = row['배송메세지'] || '';

    // 옵션정보 파싱 결과가 없으면 상품명에서 추출한 값 사용
    const design = parsed.design || parsedFromName.design || '';
    const thickness = parsed.thickness || parsedFromName.thickness || '';
    const width = parsed.width || parsedFromName.width || '';
    let length = parsed.length || parsedFromName.length || '';

    // 상품번호 기반 제품 타입 판별
    const productMapping = PRODUCT_TYPE_MAP[productId];
    const productType = productMapping ? productMapping.name : getProductTypeByName(productName);
    const productCategory = productMapping ? productMapping.type : '기타';
    const productCategoryCode = productMapping ? productMapping.category : 'etc';

    // 확인 필요 감지 (상품코드 4200445704 애견롤매트에만 적용)
    const hasCuttingRequest = productId === '4200445704' && detectCuttingRequest(deliveryMemo);

    // 마감재 요청 감지 (퍼즐매트 5994906898, 5994903887에만 적용)
    const hasFinishingRequest = (productId === '5994906898' || productId === '5994903887') && detectFinishingRequest(deliveryMemo);

    // 길이값 추출 (m 단위)
    if (productCategory === 'petRoll') {
      if (!length) {
        length = '50cm';
      }
    }

    const lengthM = parseLengthToMeters(length);

    // 두께값 추출 (T 제거)
    const thicknessNum = parseThickness(thickness);

    // 폭값 추출 (cm 단위 숫자)
    const widthNum = parseWidth(width);

    return {
      id: row['상품주문번호'] || '',
      orderId: row['주문번호'] || '',
      productId: productId,
      customerName: row['수취인명'] || row['구매자명'] || '',
      productName: productName,
      optionInfo: optionInfo, // 옵션정보 추가 (플러스 감지용)
      productType: productType,
      productCategory: productCategory,
      productCategoryCode: productCategoryCode,
      design: design,
      thickness: thickness,
      thicknessNum: thicknessNum,
      width: width,
      widthNum: widthNum,
      length: length,
      lengthM: lengthM,
      quantity: parseInt(row['수량']) || 1,
      price: parsePrice(row['최종 상품별 총 주문금액']),
      status: row['주문상태'] || '',
      gift: row['사은품'] || '',
      deliveryMemo: deliveryMemo,
      address: row['통합배송지'] || row['배송지'] || '',
      zipCode: row['우편번호'] || extractZipCode(row['통합배송지'] || ''),
      phone: row['수취인연락처1'] || '',
      orderDate: row['주문일시'] || '',
      // 자동 감지 결과
      hasCuttingRequest: hasCuttingRequest,
      hasFinishingRequest: hasFinishingRequest,
      // 원본 CSV 데이터 저장
      rawRow: row
    };
  } catch (e) {
    console.error('파싱 오류:', e, row);
    return null;
  }
}

// 우편번호 추출
function extractZipCode(address) {
  const match = address.match(/\((\d{5})\)/);
  return match ? match[1] : '';
}

// 두께값 파싱 (예: "17T" → 17, "1.7cm" → 17)
function parseThickness(thicknessStr) {
  if (!thicknessStr) return 0;
  // T 형식 (예: 17T)
  let match = thicknessStr.match(/(\d+)T/i);
  if (match) return parseInt(match[1]);
  // cm 형식 (예: 1.7cm)
  match = thicknessStr.match(/(\d+\.?\d*)cm/i);
  if (match) return Math.round(parseFloat(match[1]) * 10);
  // mm 형식 (예: 25mm)
  match = thicknessStr.match(/(\d+)mm/i);
  if (match) return Math.round(parseInt(match[1]) / 10 * 10);
  return 0;
}

// 폭값 파싱 (예: "140cm" → 140)
function parseWidth(widthStr) {
  if (!widthStr) return 0;
  const match = widthStr.match(/(\d+)/);
  return match ? parseInt(match[1]) : 0;
}

// 길이값을 m 단위로 변환 (예: "800cm" → 8, "8m" → 8)
function parseLengthToMeters(lengthStr) {
  if (!lengthStr) return 0;
  // m 단위
  let match = lengthStr.match(/(\d+\.?\d*)m(?!m)/i); // mm가 아닌 m만 찾음
  if (match) return parseFloat(match[1]);

  // cm 단위 (명시적)
  match = lengthStr.match(/(\d+)cm/i);
  if (match) return parseInt(match[1]) / 100;

  // 단위 없는 숫자
  // 30 이상이면 cm로 간주 (0.3m), 30 미만이면 m로 간주
  // 롤매트는 보통 m단위 판매가 많지만 100단위는 cm일 확률 높음
  match = lengthStr.match(/(\d+(\.\d+)?)/);
  if (match) {
    const val = parseFloat(match[1]);
    if (val >= 30) return val / 100; // 30cm 이상은 cm로 간주
    return val; // 그 외는 m
  }
  return 0;
}

// 재단 요청 감지
function detectCuttingRequest(text) {
  if (!text) return false;
  const lowerText = text.toLowerCase();

  // 숫자 포함 여부
  const hasNumber = /\d/.test(text);

  // 키워드 검색
  const hasKeyword = CUTTING_KEYWORDS.some(keyword => lowerText.includes(keyword.toLowerCase()));

  return hasKeyword || hasNumber;
}

// 마감재 요청 감지
function detectFinishingRequest(text) {
  if (!text) return false;
  return FINISHING_KEYWORDS.some(keyword => text.includes(keyword));
}

// ========== 디자인 코드 변환 ==========

// 색상명 축약
function shortenColor(colorName) {
  if (!colorName) return '';
  // 축약 규칙에 있으면 축약 적용
  for (const [full, short] of Object.entries(COLOR_ABBREVIATIONS)) {
    if (colorName.includes(full)) {
      return short;
    }
  }
  // 없으면 앞 4글자만
  return colorName.length > 4 ? colorName.substring(0, 4) : colorName;
}

// 프로모션 문구 제거
function removePromotions(text) {
  if (!text) return '';
  return text
    .replace(/⚡[^⚡]*⚡/g, '')
    .replace(/⛄[^⛄]*⛄/g, '')
    .replace(/✨[^✨]*✨/g, '')
    .replace(/\[특가\]/g, '')
    .replace(/\[할인\]/g, '')
    .replace(/\[강세일\]/g, '')
    .replace(/\[홀리데이\]/g, '')
    .trim();
}

// 송장용 디자인 코드 생성 (메인 함수)
function generateDesignCode(order) {
  const productId = order.productId;
  const optionInfo = removePromotions(order.productName + ' ' + (order.optionInfo || ''));

  // 제품 타입별 분기
  if (PUZZLE_PRODUCT_IDS.includes(productId) || order.productCategory === '퍼즐') {
    return generatePuzzleCode(order);
  } else if (productId === '282153807' || productId === '561709916' || order.productCategory === '벽지') {
    return generateWallpaperCode(order);
  } else if (TAPE_PRODUCT_IDS.includes(productId) || order.productCategory === '테이프') {
    return generateTapeCode(order);
  } else {
    return generateRollMatCode(order);
  }
}

// 롤매트 디자인 코드: (폭)두께T디자인명
function generateRollMatCode(order) {
  const width = order.widthNum || 0;
  const thickness = order.thicknessNum || 0;
  const design = shortenColor(order.design || '');
  const lengthM = order.lengthM || 0;

  if (width && thickness && design) {
    return `(${width})${thickness}T${design}`;
  }
  // 폴백: 원본 디자인 사용
  return order.design || order.productName || '';
}

// 롤매트 디자인+길이+수량 코드 (로젠용)
function generateRollMatCodeWithLength(order) {
  const baseCode = generateRollMatCode(order);
  const lengthM = order.lengthM || 0;
  const qty = order.quantity || 1;

  if (lengthM > 0) {
    return `${baseCode}${lengthM}mx${qty}`;
  }
  return `${baseCode}x${qty}`;
}

// 퍼즐매트 디자인 코드: 두께T(퍼즐)색상(사이즈) 또는 플러스♥두께T색상(사이즈)
function generatePuzzleCode(order) {
  const thickness = order.thicknessNum || 25;
  const color = shortenColor(order.design || '아이보리');
  const optionInfo = order.productName + ' ' + (order.optionInfo || '');

  // 사이즈 추출 (100x100 또는 50x50)
  let size = '100';
  if (optionInfo.includes('50x50') || optionInfo.includes('50cm')) {
    size = '50';
  }

  // PLUS+ 제품 확인
  if (optionInfo.includes('PLUS') || optionInfo.includes('플러스')) {
    return `플러스(PLUS)${thickness}T${color}(${size})`;
  }

  return `${thickness}T(퍼즐)${color}(${size})`;
}

// 퍼즐매트 코드+수량 (로젠용)
function generatePuzzleCodeWithQty(order) {
  const baseCode = generatePuzzleCode(order);
  const qty = order.quantity || 1;
  return `${baseCode}x${qty}`;
}

// 벽지매트 디자인 코드: 두께T벽지:디자인명
function generateWallpaperCode(order) {
  const thickness = order.thicknessNum || 10;
  const design = shortenColor(order.design || '');
  return `${thickness}T벽지:${design}`;
}

// 테이프 코드
function generateTapeCode(order) {
  const optionInfo = order.productName + ' ' + (order.optionInfo || '');
  // 길이 추출 (20m, 10m 등)
  const lengthMatch = optionInfo.match(/(\d+)m/);
  const length = lengthMatch ? lengthMatch[1] : '20';

  // 사은품인지 판매인지 구분 (기본: 판매)
  if (order.gift || optionInfo.includes('증정') || optionInfo.includes('사은품')) {
    return `★증정★테이프${length}m`;
  }
  return `★판매★테이프${length}m`;
}

// 길이 변환 (cm → m)
function formatLengthToMeters(lengthCm) {
  if (!lengthCm || lengthCm === 0) return '';
  const meters = lengthCm / 100;
  // 소수점 처리
  if (meters === Math.floor(meters)) {
    return `${meters}m`;
  }
  return `${meters.toFixed(1)}m`;
}

function parseOptionInfo(optionInfo) {
  const result = { design: '', thickness: '', width: '', length: '' };

  // 디자인 추출 (디자인선택, 디자인: 모두 지원)
  let designMatch = optionInfo.match(/디자인선택:\s*([^/]+)/);
  if (!designMatch) {
    designMatch = optionInfo.match(/디자인:\s*([^/]+)/);
  }
  if (designMatch) result.design = designMatch[1].trim();

  // 두께 추출 (두께선택, 두께(폭), 두께 / 폭 모두 지원)
  let thicknessMatch = optionInfo.match(/두께선택:\s*([0-9.]+)T/i);
  if (!thicknessMatch) {
    thicknessMatch = optionInfo.match(/두께\s*\(?폭\)?:\s*([0-9.]+)(mm|cm|T)/i);
  }
  if (!thicknessMatch) {
    thicknessMatch = optionInfo.match(/두께\s*\/?\s*폭:\s*([0-9.]+)(cm|mm|T)/i);
  }
  if (thicknessMatch) {
    const value = parseFloat(thicknessMatch[1]);
    const unit = thicknessMatch[2] ? thicknessMatch[2].toLowerCase() : 'T';
    if (unit === 't') {
      result.thickness = value + 'T';
    } else if (unit === 'mm') {
      result.thickness = value + 'mm';
    } else {
      result.thickness = value + 'cm';
    }
  }

  // 폭 추출 (두께/폭 패턴 또는 길이선택의 첫 번째 숫자)
  const thicknessWidthMatch = optionInfo.match(/두께\s*\/?\s*폭:\s*([0-9.]+)cm\/([0-9]+)cm/);
  if (thicknessWidthMatch) {
    result.width = thicknessWidthMatch[2] + 'cm';
  } else {
    // 두께(폭): 6mm(폭140cm) 패턴
    const widthInThicknessMatch = optionInfo.match(/\(폭([0-9]+)cm\)/);
    if (widthInThicknessMatch) {
      result.width = widthInThicknessMatch[1] + 'cm';
    }
  }

  // 길이 추출 (다양한 패턴 지원)
  // 1. "길이선택: 110x200" 패턴 (폭x길이)
  let lengthMatch = optionInfo.match(/길이선택:\s*(\d+)x(\d+)/);
  if (lengthMatch) {
    const widthNum = parseInt(lengthMatch[1]);
    const lengthNum = parseInt(lengthMatch[2]);
    // 폭이 이미 설정되지 않았으면 첫 번째 숫자를 폭으로
    if (!result.width) {
      result.width = widthNum + 'cm';
    }
    // 두 번째 숫자가 100 이상이면 길이로 사용 (롤매트는 보통 1m 이상)
    if (lengthNum >= 100) {
      result.length = lengthNum + 'cm';
    }
  }

  // 2. "길이: 800cm" 패턴
  if (!result.length) {
    lengthMatch = optionInfo.match(/길이:\s*([0-9.]+)\s*(cm|m|미터|메터)/i);
    if (lengthMatch) {
      const value = parseFloat(lengthMatch[1]);
      const unit = lengthMatch[2] ? lengthMatch[2].toLowerCase() : 'cm';
      // 'cm'가 포함되어 있으면 m로 오인하지 않도록 체크
      if (unit === 'm' || unit === '미터' || unit === '메터') {
        result.length = value + 'm';
      } else {
        result.length = value + 'cm';
      }
    }
  }

  // 3. "길이(수량추가): 50cm" 패턴 (애견롤 등)
  if (!result.length) {
    lengthMatch = optionInfo.match(/길이\(수량추가\):\s*([0-9.]+)\s*(cm|m|미터|메터)/i);
    if (lengthMatch) {
      const value = parseFloat(lengthMatch[1]);
      const unit = lengthMatch[2] ? lengthMatch[2].toLowerCase() : 'cm';
      if (unit === 'm' || unit.includes('미터') || unit.includes('메터')) {
        result.length = value + 'm';
      } else {
        result.length = value + 'cm';
      }
    }
  }

  // 4. "길이:" 없이 숫자+m/cm 패턴도 찾기 (예: "3m", "300cm", "3미터 롤")
  if (!result.length) {
    lengthMatch = optionInfo.match(/([0-9.]+)\s*(m|미터|메터)(\s*롤)?/i);
    if (lengthMatch) {
      const value = parseFloat(lengthMatch[1]);
      result.length = value + 'm';
    }
  }

  // 5. cm 단위로도 찾기 (예: "300cm") - 3자리 이상 숫자
  if (!result.length) {
    lengthMatch = optionInfo.match(/([0-9]{3,})\s*cm/i);
    if (lengthMatch) {
      result.length = lengthMatch[1] + 'cm';
    }
  }

  // 퍼즐매트 파싱
  const puzzleMatch = optionInfo.match(/\(([0-9]+)mm\)\s*([0-9]+)x([0-9]+)\s*(\d+)장/);
  if (puzzleMatch) {
    result.thickness = puzzleMatch[1] + 'mm';
    result.width = puzzleMatch[2] + 'x' + puzzleMatch[3];
    result.length = puzzleMatch[4] + '장';
  }

  // 색상 추출
  const colorMatch = optionInfo.match(/색상:\s*([^/]+)/);
  if (colorMatch && !result.design) result.design = colorMatch[1].trim();

  return result;
}

// 상품명에서 디자인/두께/폭/길이 추출 (예: "마블아이보리 6T 110x50cm")
function parseProductName(productName) {
  const result = { design: '', thickness: '', width: '', length: '' };
  if (!productName) return result;

  // 알려진 디자인명 패턴
  const designPatterns = [
    '마블아이보리', '퓨어아이보리', '그레이캔버스', '딜라이트우드', '모던그레이',
    '스노우화이트', '내추럴우드', '바닐라아이보리', '라이트그레이', '스카이블루',
    '파스텔핑크', '무직타이거', '코지베어', '플라워가든', '스타라이트',
    '헬로베어', '트로피칼', '사파리', '포레스트', '클라우드'
  ];

  // 디자인명 추출
  for (const design of designPatterns) {
    if (productName.includes(design)) {
      result.design = design;
      break;
    }
  }

  // 두께 추출: 6T, 9T, 10T, 12T, 17T, 22T 또는 x2.5(cm단위) 등
  let thicknessMatch = productName.match(/(\d+)T\b/i);
  if (thicknessMatch) {
    result.thickness = thicknessMatch[1] + 'T';
  } else {
    // 퍼즐매트용 x2.5 소수점 패턴 (cm -> mm 변환)
    thicknessMatch = productName.match(/x(\d+\.\d+)/);
    if (thicknessMatch) {
      result.thickness = Math.round(parseFloat(thicknessMatch[1]) * 10) + 'mm';
    }
  }

  // 폭x길이 추출: 110x50cm, 140x100cm 등
  // 주의: 두 번째 숫자가 너무 작으면(100cm 이하) 길이가 아닐 수 있음
  // 롤매트는 보통 길이가 1m 이상이므로, 100cm 이하는 길이로 사용하지 않음
  const sizeMatch = productName.match(/(\d+)x(\d+)cm/i);
  if (sizeMatch) {
    const widthNum = parseInt(sizeMatch[1]);
    const lengthNum = parseInt(sizeMatch[2]);
    result.width = sizeMatch[1] + 'cm';
    // 두 번째 숫자가 100cm 이상이거나, 폭보다 큰 경우에만 길이로 사용
    // (롤매트는 보통 길이가 폭보다 길거나 같음)
    if (lengthNum >= 100 || lengthNum >= widthNum) {
      result.length = sizeMatch[2] + 'cm';
    }
    // 그 외의 경우는 길이 정보가 아닐 가능성이 높음 (예: 110x50cm는 폭x폭 또는 다른 의미)
  }

  // 상품명에서 m 단위 길이 직접 추출 (예: "3m", "5미터")
  if (!result.length) {
    const lengthMMatch = productName.match(/([0-9.]+)\s*(m|미터|메터)(\s*롤)?/i);
    if (lengthMMatch) {
      result.length = lengthMMatch[1] + 'm';
    }
  }

  return result;
}

// 상품명 기반 제품타입 판별 (폴백)
function getProductTypeByName(productName) {
  if (!productName) return '기타';
  if (productName.includes('퍼즐')) return '퍼즐매트';
  if (productName.includes('애견') || productName.includes('펫')) return '애견롤매트';
  if (productName.includes('유아') || productName.includes('아기')) return '유아롤매트';
  if (productName.includes('TPU')) return 'TPU매트';
  if (productName.includes('PE')) return 'PE롤';
  if (productName.includes('벽지')) return '벽지';
  if (productName.includes('테이프')) return '테이프';
  return '기타';
}

/**
 * 금액 문자열을 숫자로 변환
 * @param {string} priceStr - "₩125,300", "23000" 등
 * @returns {number}
 */
function parsePrice(priceStr) {
  if (!priceStr) return 0;
  // 숫자와 소수점만 남기고 제거 (콤마, 통화문자 등 제거)
  const cleaned = String(priceStr).replace(/[^0-9.]/g, '');
  return parseInt(cleaned) || 0;
}

const tabBtns = document.querySelectorAll('.tab-btn');
const contentSections = document.querySelectorAll('.content-section');

// 탭 전환 함수
function switchTab(tabId) {
  currentTabId = tabId;
  // 탭 버튼 활성화
  tabBtns.forEach(btn => {
    if (btn.dataset.tab === tabId) {
      btn.classList.add('active');
    } else {
      btn.classList.remove('active');
    }
  });

  // 섹션 표시
  contentSections.forEach(section => {
    section.style.display = 'none';
    section.classList.remove('active');
  });

  if (tabId === 'combined') {
    document.getElementById('combinedSection').style.display = 'block';
    renderCombinedTable();
  } else if (tabId === 'kyungdong') {
    document.getElementById('kyungdongSection').style.display = 'block';
    renderKyungdongTable();
  } else if (tabId === 'lozen') {
    document.getElementById('lozenSection').style.display = 'block';
    renderLozenTable();
  } else if (tabId === 'raw') {
    document.getElementById('rawSection').style.display = 'block';
    renderRawData();
  }
}

function rerenderActiveTab() {
  if (currentTabId === 'combined') {
    renderCombinedTable();
  } else if (currentTabId === 'kyungdong') {
    renderKyungdongTable();
  } else if (currentTabId === 'lozen') {
    renderLozenTable();
  } else if (currentTabId === 'raw') {
    renderRawData();
  }
}

// 탭 클릭 이벤트
tabBtns.forEach(btn => {
  btn.addEventListener('click', () => {
    switchTab(btn.dataset.tab);
  });
});

// 원본 데이터 렌더링
function getFilteredRawOrderData() {
  if (!rawOrderData || rawOrderData.length === 0) return [];
  if (!filteredOrders || filteredOrders.length === 0) return [];

  const rawRowsSet = new Set(filteredOrders.map(o => o.rawRow).filter(Boolean));
  if (rawRowsSet.size === 0) return rawOrderData;
  return rawOrderData.filter(row => rawRowsSet.has(row));
}

function renderRawData() {
  const table = document.getElementById('rawTable');
  const thead = table.querySelector('thead');
  const tbody = table.querySelector('tbody');

  thead.innerHTML = '';
  tbody.innerHTML = '';

  const rowsToRender = getFilteredRawOrderData();
  if (!rowsToRender || rowsToRender.length === 0) {
    tbody.innerHTML = '<tr><td colspan="20" style="text-align: center; padding: 2rem; color: var(--text-muted);">데이터가 없습니다. 파일을 업로드해주세요.</td></tr>';
    return;
  }

  // 헤더 생성 (첫 번째 데이터의 키 사용)
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

  // 데이터 생성 (같은 수취인끼리 그룹화하여 배경색 적용)
  let currentCustomer = '';
  let groupIndex = 0;

  rowsToRender.forEach(row => {
    const tr = document.createElement('tr');
    tr.style.cursor = 'pointer';

    // 수취인명 추출
    const customerName = row['수취인명'] || row['구매자명'] || '';

    // 수취인명이 변경되면 그룹 인덱스 증가
    if (customerName && customerName !== currentCustomer) {
      currentCustomer = customerName;
      groupIndex++;
    }

    // 그룹별 배경색 적용 (흰색/회색 번갈아)
    if (groupIndex % 2 === 0) {
      tr.classList.add('raw-row-group-even');
    } else {
      tr.classList.add('raw-row-group-odd');
    }

    headers.forEach(header => {
      const td = document.createElement('td');
      let cellValue = row[header] || '';

      // \ 문자 제거
      if (typeof cellValue === 'string') {
        cellValue = cellValue.replace(/\\/g, '');
      }

      // 비용 컬럼 처리 (금액 관련 컬럼)
      const costColumns = ['최종 상품별 총 주문금액', '상품별 총 주문금액', '주문금액', '금액', '가격', '비용'];
      if (costColumns.some(col => header.includes(col))) {
        // 금액 형식 정리 (숫자만 남기거나 포맷팅)
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
}

function clearFile() {
  fileInput.value = '';
  fileInfo.style.display = 'none';
  ordersData = [];
  filteredOrders = [];
  packedOrders = [];
  rawOrderData = []; // 원본 데이터 초기화
  updateUI();
}

// ========== UI 업데이트 ==========
function updateUI() {
  if (ordersData.length === 0) {
    summarySection.style.display = 'none';
    filterSection.style.display = 'none';
    ordersSection.style.display = 'none';
    uploadArea.style.display = 'flex'; // 데이터가 없으면 업로드 영역 다시 표시
    uploadArea.style.flexDirection = 'column';
    uploadArea.style.alignItems = 'center';
    uploadArea.style.justifyContent = 'center';
    return;
  }
  
  // 데이터가 있으면 업로드 영역 숨기기
  uploadArea.style.display = 'none';
  
  summarySection.style.display = 'block';
  filterSection.style.display = 'flex';
  ordersSection.style.display = 'block';

  groupOrdersByRecipient();
  processPacking(); // 자동 패킹 처리
  updateSummary();
  renderOrders();
  switchTab(currentTabId); // 활성 탭의 발주서 테이블 렌더링 및 섹션 표시
}

function updateSummary() {
  const uniqueCustomers = Object.keys(groupedOrders).length;

  // 사은품 대상자 계산 (9.5만원+, 50만원+)
  const giftEligible = Object.values(groupedOrders).filter(g => g.totalPrice >= 95000).length;

  // 택배사별 박스 수 계산
  const kyungdongBoxes = packedOrders.filter(box => box.courier === 'kyungdong').length;
  const lozenBoxes = packedOrders.filter(box => box.courier === 'lozen').length;

  kyungdongCount.textContent = kyungdongBoxes;
  lozenCount.textContent = lozenBoxes;
  totalCustomers.textContent = uniqueCustomers;
  giftOrders.textContent = giftEligible;
}

// ========== 패킹 알고리즘 ==========

// 동일 수취인 그룹화 (이름 + 주소 + 연락처)
function groupOrdersByRecipient() {
  groupedOrders = {};

  filteredOrders.forEach(order => {
    // 동일 수취인 키 생성 (이름 + 주소 + 연락처)
    const key = `${order.customerName}|${order.address}|${order.phone}`;

    if (!groupedOrders[key]) {
      groupedOrders[key] = {
        key: key,
        orderId: order.orderId,
        customerName: order.customerName,
        address: order.address,
        zipCode: order.zipCode,
        phone: order.phone,
        status: order.status,
        orderDate: order.orderDate,
        items: [],
        rollItems: [], // 롤매트 아이템
        puzzleItems: [], // 퍼즐매트 아이템
        tapeItems: [], // 테이프 아이템
        otherItems: [], // 기타 아이템
        totalPrice: 0,
        hasGift: false,
        giftInfo: '',
        deliveryMemo: order.deliveryMemo,
        hasCuttingRequest: false,
        hasFinishingRequest: false,
      };
    }

    groupedOrders[key].items.push(order);
    groupedOrders[key].totalPrice += order.price;

    // 제품군별 분류
    const fullProductName = (order.productName + ' ' + (order.rawRow['옵션정보'] || '')).toLowerCase();

    if (fullProductName.includes('테이프')) {
      groupedOrders[key].tapeItems.push(order);
    } else if (ROLL_PRODUCT_IDS.includes(order.productId) || order.productCategory === '롤') {
      groupedOrders[key].rollItems.push(order);
    } else if (PUZZLE_PRODUCT_IDS.includes(order.productId) || order.productCategory === '퍼즐') {
      groupedOrders[key].puzzleItems.push(order);
    } else {
      groupedOrders[key].otherItems.push(order);
    }

    if (order.gift && order.gift.trim() !== '') {
      groupedOrders[key].hasGift = true;
      groupedOrders[key].giftInfo = order.gift;
    }

    if (order.deliveryMemo && order.deliveryMemo.trim() !== '') {
      // 배송메모를 배열로 수집 (중복 제거)
      if (!groupedOrders[key].deliveryMemos) {
        groupedOrders[key].deliveryMemos = [];
      }
      const trimmedMemo = order.deliveryMemo.trim();
      if (!groupedOrders[key].deliveryMemos.includes(trimmedMemo)) {
        groupedOrders[key].deliveryMemos.push(trimmedMemo);
      }
    }

    if (order.hasCuttingRequest) {
      groupedOrders[key].hasCuttingRequest = true;
    }
    if (order.hasFinishingRequest) {
      groupedOrders[key].hasFinishingRequest = true;
    }
  });

  // 사은품 대상자 마킹
  Object.values(groupedOrders).forEach(group => {
    // 롤매트(유아/애견) 상품만 필터링
    const rollMatOrders = group.items.filter(item =>
      item.productId === '6092903705' || item.productId === '4200445704'
    );

    // 롤매트 상품의 총 금액
    const rollMatTotalPrice = rollMatOrders.reduce((sum, item) => sum + item.price, 0);

    if (rollMatTotalPrice >= 500000) {
      group.giftEligible = '실리콘테이프 2개';
    } else if (rollMatTotalPrice >= 95000) {
      group.giftEligible = '실리콘테이프 1개';
    }
  });
}

// 롤매트 포장 방식 결정
function determineRollPackaging(item) {
  const thickness = item.thicknessNum;
  const length = item.lengthM;
  const width = item.widthNum;

  // 두께별 기준 찾기
  const thresholds = PACKAGING_THRESHOLDS[String(thickness)] || PACKAGING_THRESHOLDS["17"];

  // 140폭 8m 이상 → 비닐 + ★마킹
  if (width >= 140 && length >= 8) {
    return { type: 'vinyl', needsStar: true };
  }

  // 두께별 기준 적용
  if (length >= thresholds.vinyl) {
    return { type: 'vinyl', needsStar: true, canCombine: false };
  } else if (length >= thresholds.large) {
    return { type: 'largeBox', needsStar: false, canCombine: false };
  } else if (length >= thresholds.small) {
    return { type: 'smallBox', needsStar: false, canCombine: false };
  } else {
    // 기준 미달(소형) -> 소박스지만 합포장 가능
    return { type: 'smallBox', needsStar: false, canCombine: true };
  }
}

// 퍼즐매트 박스 수 계산
function calculatePuzzleBoxes(items) {
  let totalCount = 0;
  let thickness = 25; // 기본값

  items.forEach(item => {
    totalCount += item.quantity;
    if (item.thicknessNum === 40) {
      thickness = 40;
    }
  });

  const capacity = PUZZLE_BOX_CAPACITY[String(thickness)] || 6;
  return Math.ceil(totalCount / capacity);
}

function parseCuttingSegmentsFromMemo(memo) {
  const text = String(memo || '');
  if (!text) return [];

  const segments = [];
  const regex = /(\d+(?:\.\d+)?)\s*(?:m|미터|메터)\s*(?:씩)?\s*(\d+)\s*롤/gi;
  let match;
  while ((match = regex.exec(text)) !== null) {
    const lengthM = parseFloat(match[1]);
    const count = parseInt(match[2], 10);
    if (Number.isFinite(lengthM) && lengthM > 0 && Number.isFinite(count) && count > 0) {
      segments.push({ lengthM, count });
    }
  }
  return segments;
}

function getDistinctLengthRequestsFromMemo(memo) {
  const text = String(memo || '').toLowerCase();
  if (!text) return [];

  const matches = text.match(/(\d+(?:\.\d+)?)\s*(?:m|미터|메터)/gi) || [];
  const values = matches
    .map(m => parseFloat(m.replace(/[^0-9.]/g, '')))
    .filter(v => Number.isFinite(v) && v > 0);

  const distinct = [];
  values.forEach(v => {
    if (!distinct.some(x => Math.abs(x - v) < 0.05)) distinct.push(v);
  });
  return distinct;
}

function collectWarningsFromItems(items) {
  const warningsSet = new Set();
  (items || []).forEach(item => {
    (item.cuttingWarnings || []).forEach(w => warningsSet.add(w));
  });
  return Array.from(warningsSet);
}

// 박스 내 아이템들의 배송메모를 수집 (중복 제거)
function collectDeliveryMemos(items) {
  const memosSet = new Set();
  (items || []).forEach(item => {
    if (item.deliveryMemo && item.deliveryMemo.trim() !== '') {
      memosSet.add(item.deliveryMemo.trim());
    }
  });
  return Array.from(memosSet);
}

// 스마트 재단 로직 (애견롤매트 등 50cm 단위 판매 제품)
function applySmartCutting(item) {
  // 애견롤매트(petRoll)이면서 옵션에 50cm 단위가 포함된 경우
  // 예: "길이(수량추가): 50cm"
  const isPetRoll = item.productCategory === 'petRoll' ||
    (item.rawRow['옵션정보'] && item.rawRow['옵션정보'].includes('50cm'));

  if (!isPetRoll) {
    return [item];
  }

  // 총 길이 계산 (수량 * 단위길이)
  const unitLengthM = item.lengthM > 0 ? item.lengthM : 0.5;
  const totalLengthM = item.quantity * unitLengthM;

  // 배송메모 분석
  const memo = item.deliveryMemo || '';
  const warnings = [];

  // "5m 2롤 + 9m 1롤" 같이 케이스가 명확한 경우는 그대로 반영
  const cuttingSegments = parseCuttingSegmentsFromMemo(memo);
  if (cuttingSegments.length > 0) {
    const totalRequestedM = cuttingSegments.reduce((sum, seg) => sum + (seg.lengthM * seg.count), 0);
    const epsilon = 0.05; // 5cm 허용 오차
    if (Math.abs(totalRequestedM - totalLengthM) <= epsilon) {
      const resultItems = [];
      cuttingSegments.forEach(seg => {
        for (let i = 0; i < seg.count; i++) {
          resultItems.push({
            ...item,
            quantity: 1,
            lengthM: seg.lengthM,
            length: `${seg.lengthM}m`,
            cuttingWarnings: []
          });
        }
      });
      return resultItems;
    }

    warnings.push(`재단요청 길이 합계(${totalRequestedM}m)가 주문 길이(${totalLengthM}m)와 일치하지 않습니다.`);
  } else {
    const distinctLengths = getDistinctLengthRequestsFromMemo(memo);
    if (distinctLengths.length >= 2) {
      warnings.push('재단요청에 여러 길이(m)가 포함되어 단순 파서로 자동 분리가 불확실합니다.');
    }
  }

  let splitCount = 1;
  let splitLength = 0;

  // "N롤" 또는 "N등분" 패턴 (예: "2롤", "2등분", "반씩")
  const rollMatch = memo.match(/(\d+)롤|(\d+)등분|반씩|반으로/);
  if (rollMatch) {
    if (rollMatch[1]) splitCount = parseInt(rollMatch[1]);
    else if (rollMatch[2]) splitCount = parseInt(rollMatch[2]);
    else splitCount = 2; // 반씩/반으로
  }
  // "N m" 패턴 (예: "5m씩", "5미터로")
  else {
    const lengthMatch = memo.match(/(\d+)m|(\d+)미터|(\d+)메터/i);
    if (lengthMatch) {
      const reqLength = parseInt(lengthMatch[1] || lengthMatch[2] || lengthMatch[3]);
      if (reqLength > 0 && reqLength < totalLengthM) {
        // 요청 길이로 나누기 (예: 10m를 5m씩 -> 2롤)
        splitLength = reqLength;
      }
    }
  }

  const resultItems = [];

  if (splitLength > 0) {
    // 특정 길이로 자르기
    let remaining = totalLengthM;
    while (remaining >= splitLength) {
      resultItems.push({
        ...item,
        quantity: 1, // 박스 수량 기준으로는 1개
        lengthM: splitLength,
        length: `${splitLength}m`,
        cuttingWarnings: [...warnings]
      });
      remaining -= splitLength;
    }
    if (remaining > 0) {
      resultItems.push({
        ...item,
        quantity: 1,
        lengthM: remaining,
        length: `${remaining}m`,
        cuttingWarnings: [...warnings]
      });
    }
  } else if (splitCount > 1) {
    // N등분
    const lengthPerRoll = totalLengthM / splitCount;
    for (let i = 0; i < splitCount; i++) {
      resultItems.push({
        ...item,
        quantity: 1,
        lengthM: lengthPerRoll,
        length: `${lengthPerRoll}m`,
        cuttingWarnings: [...warnings]
      });
    }
  } else {
    // 분할 없음 -> 통으로 1롤
    resultItems.push({
      ...item,
      quantity: 1,
      lengthM: totalLengthM,
      length: `${totalLengthM}m`,
      cuttingWarnings: [...warnings]
    });
  }

  return resultItems;
}

// 패킹 처리 메인 함수
function processPacking() {
  packedOrders = [];

  Object.values(groupedOrders).forEach(group => {
    group._warnings = [];
    const boxes = [];

    // 1. 롤매트 처리
    let standaloneBoxes = [];

    // 스마트 재단 로직 적용: 애견롤매트 등 50cm 단위 판매 제품을 메모에 따라 분리/병합
    const processedRollItems = group.rollItems.flatMap(item => applySmartCutting(item));
    group._warnings = collectWarningsFromItems(processedRollItems);

    // 확인 필요 주문 경고 추가 (배송메모 원문 포함) - 프론트엔드 표시용
    if (group.hasCuttingRequest) {
      const memos = group.items.map(o => o.deliveryMemo).filter(m => m && m.trim());
      const uniqueMemos = [...new Set(memos)];
      group._warnings.unshift(`⚠️ 확인: ${uniqueMemos.join(' / ')}`);
    }

    // 롤매트 합포장 로직: 소형 롤들을 모아서 하나의 박스에
    // 1) 비닐 포장 대상 (대형)은 단독 박스
    // 2) 소형/중형은 합포장 시도
    const vinylItems = []; // 비닐 포장 (단독)
    const combinableItems = []; // 합포장 가능 (소박스/대박스)

    processedRollItems.forEach(item => {
      const packaging = determineRollPackaging(item);
      for (let i = 0; i < item.quantity; i++) {
        if (packaging.type === 'vinyl') {
          vinylItems.push({ ...item, quantity: 1, _packaging: packaging });
        } else {
          combinableItems.push({ ...item, quantity: 1, _packaging: packaging });
        }
      }
    });

    // 비닐 포장 아이템은 단독 박스
    vinylItems.forEach(item => {
      const designCode = generateRollMatCode(item);
      standaloneBoxes.push({
        type: 'roll',
        packagingType: 'vinyl',
        needsStar: true,
        items: [item],
        designText: `${designCode}${item.lengthM}m`,
        isCombined: false,
        remark: '',
        deliveryMemo: item.deliveryMemo,
        warnings: collectWarningsFromItems([item])
      });
    });

    // 합포장 가능 아이템들 처리 (두께가 달라도 합포장 가능)
    if (combinableItems.length > 0) {
      // 가장 두꺼운 아이템의 기준으로 박스 용량 결정 (두꺼울수록 용량 작음)
      const maxThickness = Math.max(...combinableItems.map(i => i.thicknessNum || 17));
      const thresholds = PACKAGING_THRESHOLDS[String(maxThickness)] || PACKAGING_THRESHOLDS["17"];

      // 길이 합산하여 박스에 배치
      let currentBox = { items: [], totalLength: 0 };
      const finishedBoxes = [];

      combinableItems.forEach(item => {
        const len = item.lengthM || 0;
        // 현재 박스에 추가했을 때 대박스 기준 초과하면 새 박스
        if (currentBox.totalLength + len > thresholds.large) {
          if (currentBox.items.length > 0) {
            finishedBoxes.push(currentBox);
          }
          currentBox = { items: [item], totalLength: len };
        } else {
          currentBox.items.push(item);
          currentBox.totalLength += len;
        }
      });
      if (currentBox.items.length > 0) {
        finishedBoxes.push(currentBox);
      }

      // 각 박스의 포장 타입 결정
      finishedBoxes.forEach(box => {
        const totalLen = box.totalLength;
        // 박스 내 가장 두꺼운 아이템 기준으로 포장 타입 결정
        const boxMaxThickness = Math.max(...box.items.map(i => i.thicknessNum || 17));
        const boxThresholds = PACKAGING_THRESHOLDS[String(boxMaxThickness)] || PACKAGING_THRESHOLDS["17"];

        let packagingType = 'smallBox';
        let needsStar = false;

        if (totalLen >= boxThresholds.vinyl) {
          packagingType = 'vinyl';
          needsStar = true;
        } else if (totalLen > boxThresholds.small) {
          packagingType = 'largeBox';
        }

        const designText = box.items.map(i => {
          const code = generateRollMatCode(i);
          return `${code}${i.lengthM}m`;
        }).join('+');

        standaloneBoxes.push({
          type: 'roll',
          packagingType: packagingType,
          needsStar: needsStar,
          items: box.items,
          designText: designText,
          isCombined: box.items.length > 1,
          remark: box.items.length > 1 ? '합' : '',
          deliveryMemo: box.items.find(i => i.deliveryMemo)?.deliveryMemo || '',
          warnings: collectWarningsFromItems(box.items)
        });
      });
    }

    // 2. 퍼즐매트 처리 (비닐 포장)
    if (group.puzzleItems.length > 0) {
      const puzzleItemsCopy = group.puzzleItems.map(item => ({ ...item }));
      const puzzleBoxCount = calculatePuzzleBoxes(puzzleItemsCopy);
      const puzzleDesign = puzzleItemsCopy.map(p => generatePuzzleCode(p)).join('+');

      const thickness = puzzleItemsCopy[0].thicknessNum || 25;
      const capacity = PUZZLE_BOX_CAPACITY[String(thickness)] || 6;

      let itemIndex = 0;
      for (let i = 0; i < puzzleBoxCount; i++) {
        const boxItems = [];
        let remainingCapacity = capacity;

        while (itemIndex < puzzleItemsCopy.length && remainingCapacity > 0) {
          const item = puzzleItemsCopy[itemIndex];
          const takeQty = Math.min(item.quantity, remainingCapacity);

          if (takeQty > 0) {
            boxItems.push({ ...item, quantity: takeQty });
            remainingCapacity -= takeQty;
            if (takeQty >= item.quantity) {
              itemIndex++;
            } else {
              item.quantity -= takeQty;
            }
          } else {
            itemIndex++;
          }
        }

        if (boxItems.length > 0) {
          // 박스 내 아이템들의 메모 중 하나를 사용 (보통 퍼즐매트는 동일 메모)
          const boxMemo = boxItems.find(i => i.deliveryMemo)?.deliveryMemo || '';

          standaloneBoxes.push({
            type: 'puzzle',
            packagingType: 'vinyl',
            needsStar: false,
            items: boxItems,
            designText: puzzleDesign,
            isCombined: false,
            remark: '',
            deliveryMemo: boxMemo
          });
        }
      }
    }

    // 3. 기타 아이템 처리
    group.otherItems.forEach(item => {
      for (let i = 0; i < item.quantity; i++) {
        standaloneBoxes.push({
          type: 'other',
          packagingType: 'box',
          needsStar: false,
          items: [item],
          designText: generateDesignCode(item),
          isCombined: false,
          remark: '',
          deliveryMemo: item.deliveryMemo
        });
      }
    });

    // 4. 테이프 처리 (어디든 낑겨넣기)
    if (group.tapeItems.length > 0) {
      if (standaloneBoxes.length > 0) {
        // 기존 박스(마지막 박스)에 합포장
        const targetBox = standaloneBoxes[standaloneBoxes.length - 1];
        group.tapeItems.forEach(item => {
          targetBox.items.push({ ...item });
        });
        targetBox.isCombined = true;
        targetBox.remark = '합';

        // 디자인 텍스트 업데이트
        const tapeDesigns = group.tapeItems.map(i => generateDesignCode(i)).join('+');
        targetBox.designText += '+' + tapeDesigns;

        // 메모 업데이트: 기존 메모가 있으면 유지, 없으면 테이프 메모 사용 (혹은 병합?)
        // 보통 테이프는 '증정'이거나 별도 메모가 중요치 않으므로 기존 박스 메모 유지.
        // 하지만 만약 기존 박스 메모가 비어있고 테이프에 메모가 있다면?
        if (!targetBox.deliveryMemo) {
          const tapeMemo = group.tapeItems.find(i => i.deliveryMemo)?.deliveryMemo;
          if (tapeMemo) targetBox.deliveryMemo = tapeMemo;
        }

      } else {
        // 박스가 전혀 없으면 새 소박스 생성
        const tapeDesigns = group.tapeItems.map(i => generateDesignCode(i)).join('+');
        const tapeMemo = group.tapeItems.find(i => i.deliveryMemo)?.deliveryMemo || '';

        standaloneBoxes.push({
          type: 'roll', // 분류상 롤/소품으로 처리
          packagingType: 'smallBox',
          needsStar: false,
          items: group.tapeItems.map(i => ({ ...i })),
          designText: tapeDesigns,
          isCombined: true,
          remark: '합',
          deliveryMemo: tapeMemo
        });
      }
    }

    // 최종 박스 리스트에 추가
    boxes.push(...standaloneBoxes);

    // 4. N-n 라벨링
    const totalBoxes = boxes.length;
    boxes.forEach((box, index) => {
      const n = index + 1;
      let recipientLabel = group.customerName;

      if (totalBoxes > 1) {
        recipientLabel = `${totalBoxes}-${n}${group.customerName}`;
      }

      if (box.needsStar) {
        recipientLabel += '★';
      }

      box.recipientLabel = recipientLabel;
      box.boxNumber = n;
      box.totalBoxes = totalBoxes;
    });

    // 결과 저장
    boxes.forEach(box => {
      // 박스 내 총 아이템 수량 계산
      const totalQtyInBox = box.items.reduce((sum, item) => sum + (item.quantity || 0), 0);

      // 박스 내 아이템들의 가격 합계 계산 (실제결제금액용)
      const boxTotalPrice = box.items.reduce((sum, item) => sum + ((item.price || 0) * (item.quantity || 1)), 0);

      // 박스 내 아이템들의 배송메모만 수집 (중복 제거)
      const boxMemos = collectDeliveryMemos(box.items);

      // 택배사 결정: 소박스 → 로젠, 대박스/비닐 → 경동
      const courier = box.packagingType === 'smallBox' ? 'lozen' : 'kyungdong';

      packedOrders.push({
        ...box,
        group: group,
        customerName: group.customerName,
        address: group.address,
        zipCode: group.zipCode,
        phone: group.phone,
        totalQtyInBox: totalQtyInBox, // 추가된 필드
        deliveryMemos: boxMemos, // 배송메모 배열 (중복 제거)
        totalPrice: boxTotalPrice, // 박스별 실제 결제 금액 (아이템 가격 합계)
        giftEligible: group.giftEligible,
        hasCuttingRequest: group.hasCuttingRequest,
        hasFinishingRequest: group.hasFinishingRequest,
        warnings: [...(group._warnings || []), ...(box.warnings && box.warnings.length ? box.warnings : collectWarningsFromItems(box.items))],
        courier: courier // 택배사 (lozen/kyungdong)
      });
    });
  });

}

// 패킹 결과 렌더링
function renderPackedResults() {
  if (packedOrders.length === 0) return;

  let html = '<div class="packed-results">';
  html += '<h3>패킹 결과</h3>';
  html += '<table class="packed-table">';
  html += '<thead><tr>';
  html += '<th>수취인</th><th>디자인</th><th>포장</th><th>비고</th>';
  html += '</tr></thead><tbody>';

  packedOrders.forEach(box => {
    const packagingText = box.packagingType === 'vinyl' ? '비닐' :
      box.packagingType === 'largeBox' ? '대박스' : '소박스';
    html += `<tr>
      <td>${box.recipientLabel}</td>
      <td>${box.designText}</td>
      <td>${packagingText}</td>
      <td>${box.remark || ''}</td>
    </tr>`;
  });

  html += '</tbody></table></div>';

  // 모달로 표시
  modalBody.innerHTML = html;
  orderModal.style.display = 'flex';
}

// ========== 배송비 데이터 로드 ==========
async function loadShippingFees() {
  const startTime = Date.now();
  console.log('=== 배송비 CSV 로드 시작 ===', new Date().toISOString());

  try {
    console.log('[1/6] fetch 요청 시작: shipping_fees.csv');
    const response = await fetch('shipping_fees.csv');
    console.log('[2/6] fetch 응답 수신:', {
      status: response.status,
      ok: response.ok,
      statusText: response.statusText,
      url: response.url
    });

    if (!response.ok) {
      console.error('❌ 배송비 데이터 파일을 찾을 수 없습니다!', {
        status: response.status,
        statusText: response.statusText
      });
      return;
    }

    console.log('[3/6] CSV 텍스트 읽기 시작...');
    const text = await response.text();
    console.log('[3/6] CSV 텍스트 읽기 완료:', {
      길이: text.length,
      첫200자: text.substring(0, 200)
    });

    console.log('[4/6] CSV 파싱 시작...');
    const rows = parseCSVToRows(text);
    console.log('[4/6] CSV 파싱 완료:', {
      총행수: rows.length,
      헤더: rows[0],
      첫데이터행: rows[1]
    });

    if (rows.length < 2) {
      console.error('❌ 배송비 데이터가 없습니다! 행 수:', rows.length);
      return;
    }

    const headers = rows[0].map(h => h.trim());
    shippingFees = [];

    console.log('[5/6] 데이터 변환 시작...');
    console.log('배송비 CSV 헤더:', headers);

    for (let i = 1; i < rows.length; i++) {
      const rowValues = rows[i];
      const rowObj = {};
      headers.forEach((header, index) => {
        rowObj[header] = rowValues[index] || '';
      });

      // 배송비가 있는 경우만 추가
      if (rowObj['배송비'] && rowObj['배송비'].trim() !== '') {
        const feeData = {
          순번: rowObj['순번'] || i,
          productGroup: rowObj['제품군'] || '',
          packageType: rowObj['포장종류'] || '',
          fee: parseInt(rowObj['배송비'].replace(/[^0-9]/g, '')) || 0,
          width: rowObj['폭(cm)'] ? parseFloat(rowObj['폭(cm)']) : null,
          thickness: rowObj['두께(cm)'] ? parseFloat(rowObj['두께(cm)']) : null,
          lengthMin: rowObj['길이_최소(m)'] ? parseFloat(rowObj['길이_최소(m)']) : null,
          lengthMax: rowObj['길이_최대(m)'] ? parseFloat(rowObj['길이_최대(m)']) : null
        };
        shippingFees.push(feeData);

        // 첫 3개 데이터는 상세 로그
        if (i <= 3) {
          console.log(`  행 ${i}:`, feeData);
        }
      }
    }

    const loadTime = Date.now() - startTime;
    console.log(`[6/6] ✅ 배송비 데이터 로드 완료! (${loadTime}ms)`);
    console.log(`총 ${shippingFees.length}건 로드됨`);

    // 제품군별 개수 확인
    const groupCounts = {};
    shippingFees.forEach(fee => {
      groupCounts[fee.productGroup] = (groupCounts[fee.productGroup] || 0) + 1;
    });
    console.log('제품군별 데이터 수:', groupCounts);

    // 퍼즐매트 데이터 특별 확인
    const puzzleData = shippingFees.filter(f => f.productGroup === '퍼즐매트');
    console.log(`퍼즐매트 데이터 ${puzzleData.length}건:`, puzzleData);

    // 전역 변수 확인
    console.log('전역 shippingFees 배열 길이:', shippingFees.length);
    console.log('=== 배송비 CSV 로드 완료 ===', new Date().toISOString());
    
    // 택배 운임비 테이블 렌더링
    renderShippingFeesTable();
  } catch (error) {
    console.error('❌ 배송비 데이터 로드 오류:', error);
    console.error('에러 스택:', error.stack);
  }
}

// ========== 페이지 전환 ==========
function switchPage(page) {
  // 사이드바 아이템 활성화
  sidebarItems.forEach(item => {
    if (item.getAttribute('data-page') === page) {
      item.classList.add('active');
    } else {
      item.classList.remove('active');
    }
  });

  // 페이지 표시/숨김
  if (page === 'orders') {
    // 주문 처리 페이지
    if (ordersSection) ordersSection.style.display = ordersData.length > 0 ? 'block' : 'none';
    if (summarySection) summarySection.style.display = ordersData.length > 0 ? 'block' : 'none';
    if (filterSection) filterSection.style.display = ordersData.length > 0 ? 'flex' : 'none';
    if (uploadArea) uploadArea.style.display = ordersData.length > 0 ? 'none' : 'flex';
    if (shippingFeesPage) shippingFeesPage.style.display = 'none';
  } else if (page === 'shipping-fees') {
    // 택배 운임비 페이지
    if (ordersSection) ordersSection.style.display = 'none';
    if (summarySection) summarySection.style.display = 'none';
    if (filterSection) filterSection.style.display = 'none';
    if (uploadArea) uploadArea.style.display = 'none';
    if (shippingFeesPage) shippingFeesPage.style.display = 'block';
  }
}

// ========== 택배 운임비 테이블 렌더링 ==========
function renderShippingFeesTable() {
  if (!shippingFeesTableBody) return;

  shippingFeesTableBody.innerHTML = '';

  if (shippingFees.length === 0) {
    shippingFeesTableBody.innerHTML = '<tr><td colspan="9" style="text-align: center; padding: 2rem; color: var(--text-muted);">배송비 데이터를 불러오는 중...</td></tr>';
    return;
  }

  // 기존 데이터 렌더링
  shippingFees.forEach((fee, index) => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td>${fee.순번 || index + 1}</td>
      <td>${fee.productGroup || ''}</td>
      <td>${fee.packageType || ''}</td>
      <td>${fee.fee ? formatCurrency(fee.fee) : ''}</td>
      <td>${fee.width !== null && fee.width !== undefined ? fee.width : ''}</td>
      <td>${fee.thickness !== null && fee.thickness !== undefined ? fee.thickness : ''}</td>
      <td>${fee.lengthMin !== null && fee.lengthMin !== undefined ? fee.lengthMin : ''}</td>
      <td>${fee.lengthMax !== null && fee.lengthMax !== undefined ? fee.lengthMax : ''}</td>
      <td></td>
    `;
    shippingFeesTableBody.appendChild(tr);
  });

  // 누락된 운임비 항목 찾기 및 추가
  const missingFees = findMissingShippingFees();
  missingFees.forEach((missingFee, index) => {
    const tr = document.createElement('tr');
    tr.classList.add('missing-fee');
    tr.innerHTML = `
      <td>-</td>
      <td>${missingFee.productGroup}</td>
      <td>${missingFee.packageType}</td>
      <td><strong>담당자 확인 필요</strong></td>
      <td>${missingFee.width || ''}</td>
      <td>${missingFee.thickness || ''}</td>
      <td>${missingFee.lengthMin || ''}</td>
      <td>${missingFee.lengthMax || ''}</td>
      <td><span style="color: var(--danger);">⚠️ ${missingFee.note}</span></td>
    `;
    shippingFeesTableBody.appendChild(tr);
  });
}

// 누락된 운임비 항목 찾기
function findMissingShippingFees() {
  const missingFees = [];
  
  // 제품군별로 그룹화
  const groupedByProduct = {};
  shippingFees.forEach(fee => {
    if (!groupedByProduct[fee.productGroup]) {
      groupedByProduct[fee.productGroup] = [];
    }
    groupedByProduct[fee.productGroup].push(fee);
  });

  // 각 제품군별로 대박스 초과 운임비 확인
  Object.keys(groupedByProduct).forEach(productGroup => {
    const fees = groupedByProduct[productGroup];
    
    // 대박스 찾기
    const largeBoxes = fees.filter(f => f.packageType && f.packageType.includes('대박스'));
    
    if (largeBoxes.length > 0) {
      // 각 대박스의 최대 길이 확인
      largeBoxes.forEach(largeBox => {
        if (largeBox.lengthMax) {
          // 대박스 초과 운임비 누락 확인
          const hasOverLargeBox = fees.some(f => 
            f.packageType && f.packageType.includes('대박스') &&
            f.lengthMin && f.lengthMin > largeBox.lengthMax
          );
          
          if (!hasOverLargeBox && largeBox.lengthMax < 20) { // 20m 미만인 경우만 체크
            missingFees.push({
              productGroup: productGroup,
              packageType: `${largeBox.packageType} 초과`,
              width: largeBox.width,
              thickness: largeBox.thickness,
              lengthMin: largeBox.lengthMax + 0.5,
              lengthMax: 20, // 임시로 20m까지
              note: '대박스 초과 운임비 필요'
            });
          }
        }
      });
    }

    // 합포장 운임비 확인 (소박스가 2개 이상 합포장되는 경우)
    const smallBoxes = fees.filter(f => f.packageType && f.packageType.includes('소박스'));
    if (smallBoxes.length > 0) {
      // 합포장 운임비가 있는지 확인
      const hasCombinedFee = fees.some(f => 
        f.packageType && (f.packageType.includes('합포장') || f.packageType.includes('합배송'))
      );
      
      if (!hasCombinedFee) {
        // 각 소박스 조합별로 합포장 운임비 필요
        const uniqueWidths = [...new Set(smallBoxes.map(f => f.width).filter(w => w))];
        uniqueWidths.forEach(width => {
          missingFees.push({
            productGroup: productGroup,
            packageType: `${width}cm 소박스 합포장`,
            width: width,
            thickness: null,
            lengthMin: null,
            lengthMax: null,
            note: '합포장 운임비 필요 (2개 이상 소박스 합포장 시)'
          });
        });
      }
    }
  });

  return missingFees;
}

// 배송비 계산 함수
function calculateShippingFee(box) {
  if (!box || !box.items || box.items.length === 0) {
    return 0;
  }

  // 첫 번째 아이템 기준으로 제품 타입 결정
  const item = box.items[0];
  const productType = item.productType || '';
  const productCategory = item.productCategoryCode || '';

  // 합포장된 박스의 총 길이 계산 (롤매트의 경우)
  const totalLength = box.items.reduce((sum, i) => sum + ((i.lengthM || 0) * (i.quantity || 1)), 0);

  // 합포장된 박스에서 가장 넓은 폭 찾기
  const widthValues = box.items.map(i => i.widthNum || 0).filter(w => w > 0);
  const maxWidth = widthValues.length > 0 ? Math.max(...widthValues) : 0;

  // 가장 두꺼운 두께 찾기 (배송비 계산용)
  const thicknessValues = box.items.map(i => i.thicknessNum || 0).filter(t => t > 0);
  const maxThickness = thicknessValues.length > 0 ? Math.max(...thicknessValues) : 0;

  // 퍼즐매트의 경우 총 장수 계산
  const totalPuzzleCount = box.items.reduce((sum, i) => sum + (i.quantity || 1), 0);

  // 제품군 매핑 (productCategoryCode 기반으로 우선 판별)
  let productGroup = '';
  const categoryCode = item.productCategoryCode || '';

  // productCategoryCode 기반 매핑 (정확한 매핑)
  if (categoryCode === 'babyRoll' || categoryCode === 'petRoll' || categoryCode === 'roll') {
    productGroup = 'PVC롤매트';
  } else if (categoryCode === 'peRoll') {
    productGroup = 'PE롤매트';
  } else if (categoryCode === 'puzzle') {
    productGroup = '퍼즐매트';
  } else if (categoryCode === 'tpu') {
    productGroup = 'TPU매트';
  } else if (categoryCode === 'wallpaper') {
    productGroup = '단열벽지';
  } else if (productType.includes('롤매트') || productCategory === 'roll' || productCategory === 'babyRoll' || productCategory === 'petRoll') {
    // 폴백: productType 또는 productCategory 기반 매핑
    if (productType.includes('PE') || productCategory === 'peRoll') {
      productGroup = 'PE롤매트';
    } else {
      productGroup = 'PVC롤매트';
    }
  } else if (productType.includes('퍼즐') || productCategory === 'puzzle') {
    productGroup = '퍼즐매트';
  } else if (productType.includes('TPU') || productCategory === 'tpu') {
    productGroup = 'TPU매트';
  } else if (productType.includes('벽지') || productCategory === 'wallpaper') {
    productGroup = '단열벽지';
  } else {
    return 0; // 매칭되는 제품군이 없으면 0 반환
  }

  // 포장종류는 길이 범위에 따라 자동 결정 - box.packagingType에 의존하지 않음
  // 배송비 CSV에서 폭/두께/길이 조건에 맞는 레코드를 찾아 포장종류 결정
  // 퍼즐매트만 비닐 포장으로 고정
  let packageType = '';
  if (productGroup === '퍼즐매트') {
    packageType = '강화비닐(100x100cm)';
  }
  // 롤매트는 길이 범위로 자동 결정되므로 packageType을 빈 문자열로 둠

  // 폭, 두께, 길이 정보 (합포장 반영)
  const width = maxWidth > 0 ? maxWidth : null;
  const thickness = maxThickness > 0 ? maxThickness / 10 : null; // T를 cm로 변환 (예: 17T → 1.7cm)
  const length = productGroup === '퍼즐매트' ? totalPuzzleCount : totalLength; // 퍼즐은 장수, 롤매트는 총 길이

  // 디버깅용 로그
  console.log('배송비 계산 시작:', {
    productGroup,
    packageType,
    width,
    thickness: maxThickness > 0 ? maxThickness / 10 : null,
    length,
    totalLength,
    boxItems: box.items.map(i => ({
      widthNum: i.widthNum,
      thicknessNum: i.thicknessNum,
      lengthM: i.lengthM,
      quantity: i.quantity,
      productType: i.productType
    }))
  });

  // 매칭되는 배송비 찾기 (상세 디버깅 로그 포함)
  console.log('=== 배송비 매칭 시작 ===');
  console.log('조건:', { productGroup, packageType, width, thickness, length });
  console.log(`로드된 배송비 데이터 수: ${shippingFees.length}`);

  // 퍼즐매트 데이터 확인
  const puzzleFees = shippingFees.filter(f => f.productGroup === '퍼즐매트');
  console.log(`퍼즐매트 배송비 데이터:`, puzzleFees);

  let matchedFee = null;
  let largeBoxFee = null; // 대박스 요금 저장 (비닐 계산용)
  let failReasons = [];

  for (const fee of shippingFees) {
    const reasons = [];

    // 1. 제품군 매칭
    if (fee.productGroup !== productGroup) {
      continue; // 제품군 다르면 아예 스킵
    }

    // 제품군 일치하면 상세 비교 로그
    console.log(`비교중: ${fee.packageType}, 폭=${fee.width}, 두께=${fee.thickness}, 길이=${fee.lengthMin}-${fee.lengthMax}`);


    // 2. 포장종류 매칭 (packageType이 비어있으면 스킵 - 롤매트는 폭/두께/길이로만 매칭)
    let packageMatch = false;

    if (packageType === '') {
      packageMatch = true;
    } else if (fee.packageType === packageType) {
      packageMatch = true;
    } else {
      const feeBaseType = fee.packageType.replace(/\d+cm\s*/, '').replace(/강화비닐\([^)]+\)/, '강화비닐').trim();
      const packageBaseType = packageType.replace(/\d+cm\s*/, '').replace(/강화비닐\([^)]+\)/, '강화비닐').trim();

      if (feeBaseType === packageBaseType ||
        (feeBaseType.includes('강화비닐') && packageBaseType.includes('강화비닐'))) {
        packageMatch = true;
      }
    }

    if (!packageMatch) {
      reasons.push(`포장종류 불일치: ${fee.packageType} vs ${packageType}`);
    }

    // 3. 폭 매칭
    let widthMatch = true;
    if (width !== null && fee.width !== null) {
      if (Math.abs(fee.width - width) > 5) {
        widthMatch = false;
        reasons.push(`폭 불일치: ${fee.width} vs ${width}`);
      }
    } else if (width === null && fee.width !== null) {
      if (productGroup === 'PVC롤매트' || productGroup === 'PE롤매트') {
        widthMatch = false;
        reasons.push(`폭 정보 없음 (필수)`);
      }
    }

    // 4. 두께 매칭
    let thicknessMatch = true;
    if (thickness !== null && fee.thickness !== null) {
      if (Math.abs(fee.thickness - thickness) > 0.1) {
        thicknessMatch = false;
        reasons.push(`두께 불일치: ${fee.thickness} vs ${thickness}`);
      }
    } else if (thickness === null && fee.thickness !== null) {
      if (productGroup === 'PVC롤매트' || productGroup === 'PE롤매트') {
        thicknessMatch = false;
        reasons.push(`두께 정보 없음 (필수)`);
      }
    }

    // 5. 길이 범위 매칭
    let lengthMatch = true;
    if (length > 0 && fee.lengthMin !== null && fee.lengthMax !== null) {
      if (length < fee.lengthMin) {
        lengthMatch = false;
        reasons.push(`길이 부족: ${length} < ${fee.lengthMin}`);
      } else if (length > fee.lengthMax) {
        lengthMatch = false;
        reasons.push(`길이 초과: ${length} > ${fee.lengthMax}`);

        // 대박스인데 길이 초과하면, 비닐 계산용으로 대박스 요금 저장
        if (fee.packageType.includes('대박스') && widthMatch && thicknessMatch) {
          largeBoxFee = fee;
          console.log('대박스 길이 초과 - 비닐 후보:', fee);
        }
      }
    } else if (length === 0 && (fee.lengthMin !== null || fee.lengthMax !== null)) {
      if (productGroup === 'PVC롤매트' || productGroup === 'PE롤매트') {
        lengthMatch = false;
        reasons.push(`길이 정보 없음 (필수)`);
      }
    }

    // 모든 조건 만족?
    if (packageMatch && widthMatch && thicknessMatch && lengthMatch) {
      matchedFee = fee;
      console.log('✓ 매칭 성공:', fee);
      break;
    } else if (reasons.length > 0) {
      failReasons.push({ fee: `${fee.packageType} (${fee.width}cm, ${fee.thickness}cm, ${fee.lengthMin}-${fee.lengthMax}m)`, reasons });
    }
  }

  // 매칭 실패 시
  if (!matchedFee) {
    // 대박스 길이 초과로 비닐 처리
    if (largeBoxFee) {
      const vinylFee = largeBoxFee.fee + 5000;
      console.log('✓ 비닐 처리 (대박스+5000원):', { 대박스요금: largeBoxFee.fee, 비닐요금: vinylFee });
      return vinylFee;
    }

    console.warn('✗ 배송비 매칭 실패');
    console.table(failReasons.slice(0, 5)); // 상위 5개 실패 이유만 표시
  }

  return matchedFee ? matchedFee.fee : 0;
}

// ========== 컬럼 설명 (Tooltip) ==========
const COLUMN_TOOLTIPS = {
  // 송장화 컬럼
  '상품주문번호': '스마트스토어에서 발급한 주문번호입니다. 같은 주문번호는 하나의 행으로 합쳐집니다.',
  '수취인명': '상품을 받는 사람의 이름입니다.',
  '운임타입': '배송비 결제 방식입니다. 현재는 모두 "선불"로 설정됩니다.',
  '송장수량': '이 주문이 몇 개의 박스(송장)로 나뉘어 배송되는지 표시합니다. 합포장이면 2개 이상입니다.',
  '디자인': '주문한 상품의 디자인명입니다. 첫 번째 상품의 디자인이 표시됩니다.',
  '길이': '롤매트의 경우 길이 정보입니다. 첫 번째 상품의 길이가 표시됩니다.',
  '수량': '이 주문의 전체 상품 수량입니다. 모든 상품의 수량을 합산한 값입니다.',
  '디자인+수량': '같은 디자인코드끼리 합쳐서 표시합니다. 예: "코드3mx2"는 코드3m이 2개라는 뜻입니다.',
  '배송메세지': '구매자가 입력한 배송 메시지입니다. 여러 개면 "/"로 구분됩니다.',
  '합포장': '여러 박스로 나뉘어 배송되면 "Y", 하나면 "N"입니다.',
  '비고': '추가 메모를 입력할 수 있는 컬럼입니다.',
  '주문하신분 핸드폰': '주문을 한 사람의 전화번호입니다.',
  '받는분 핸드폰': '상품을 받는 사람의 전화번호입니다.',
  '우편번호': '배송지의 우편번호입니다.',
  '받는분 주소': '상품을 받을 주소입니다.',
  '주문자명': '주문을 한 사람의 이름입니다.',
  '실제결제금액': '이 주문에서 실제로 결제된 금액입니다. 모든 상품의 금액을 합산한 값입니다.',
  
  // 경동 발주서 컬럼
  '받는분': '상품을 받는 사람의 이름입니다. 여러 박스면 "2-1홍길동"처럼 표시됩니다.',
  '주소': '배송지의 전체 주소입니다.',
  '상세주소': '배송 메시지가 여기에 표시됩니다.',
  '운송장번호': '택배사에서 발급하는 운송장 번호입니다. 발송 후 채워집니다.',
  '고객사주문번호': '스마트스토어 주문번호입니다.',
  '도착영업소': '배송지 근처 택배 영업소입니다. 발송 후 채워집니다.',
  '전화번호': '받는 사람의 전화번호입니다.',
  '기타전화번호': '추가 연락처입니다.',
  '선불후불': '배송비 결제 방식입니다. 현재는 모두 "선불"입니다.',
  '품목명': '디자인코드와 수량을 합쳐서 표시합니다. 예: "코드3mx2"',
  '수량': '이 박스에 들어있는 상품 수량입니다. 경동은 박스당 1개로 처리됩니다.',
  '포장상태': '비닐 포장이면 "비닐", 박스 포장이면 "박스"입니다.',
  '가로': '박스의 가로 길이(cm)입니다.',
  '세로': '박스의 세로 길이(cm)입니다.',
  '높이': '박스의 높이(cm)입니다.',
  '무게': '박스의 무게(kg)입니다.',
  '개별단가': '고정값 50원입니다.',
  '배송운임': '이 박스의 배송비입니다. 제품 크기와 무게로 자동 계산됩니다.',
  '기타운임': '고정값 100원입니다.',
  '별도운임': '고정값 0원입니다.',
  '할증운임': '고정값 0원입니다.',
  '도서운임': '고정값 0원입니다.',
  '메모': '디자인코드가 다시 표시됩니다.',
  
  // 로젠 발주서 컬럼
  '받는분 전화번호': '받는 사람의 전화번호입니다.',
  '받는분 핸드폰': '받는 사람의 휴대폰 번호입니다.',
  '운임타입': '배송비 결제 방식입니다. 모두 "선불"입니다.',
  '송장수량': '이 주문의 박스 개수입니다. 로젠은 박스당 1개로 처리됩니다.',
  '디자인+수량': '같은 디자인코드끼리 합쳐서 표시합니다.',
  '합배송여부': '여러 박스로 나뉘어 배송되면 "Y", 하나면 "N"입니다.',
  '배송운임': '이 박스의 배송비입니다. 제품 크기로 자동 계산됩니다.'
};

// ========== 발주서 생성 ==========

// 경동택배 발주서 데이터 생성
function getKyungdongData() {
  if (packedOrders.length === 0) {
    processPacking();
  }

  // 경동택배 대상만 필터 (대박스, 비닐 → courier === 'kyungdong')
  const kyungdongOrders = packedOrders.filter(box => box.courier === 'kyungdong');

  // 실무 양식 컬럼: 받는분, 주소, 상세주소, 운송장번호, 고객사주문번호, 우편번호, 도착영업소, 전화번호, 기타전화번호, 선불후불, 품목명, 수량, 포장상태, 가로, 세로, 높이, 무게, 개별단가, 배송운임, 기타운임, 별도운임, 할증운임, 도서운임, 메모
  // ※ 실무 규칙: 상세주소 칸에는 '배송메모'를 기록, 마지막 메모 칸에는 '디자인코드(줄임말)' 기록
  const headers = ["받는분", "주소", "상세주소", "운송장번호", "고객사주문번호", "우편번호", "도착영업소", "전화번호", "기타전화번호", "선불후불", "품목명", "수량", "포장상태", "가로", "세로", "높이", "무게", "개별단가", "배송운임", "기타운임", "별도운임", "할증운임", "도서운임", "메모"];
  const data = [headers];

  kyungdongOrders.forEach((box, index) => {
    const packagingStatus = box.packagingType === 'vinyl' ? '비닐' : '박스';
    const dimensions = getDimensions(box);
    const shippingFee = calculateShippingFee(box); // 배송비 계산

    // 디자인+수량 형식 생성 (같은 디자인코드는 합쳐서 수량 표시)
    let designWithQty = box.designText;
    if (box.items && box.items.length > 0) {
      const codeCountMap = new Map();
      box.items.forEach(item => {
        let code = '';
        if (PUZZLE_PRODUCT_IDS.includes(item.productId) || item.productCategory === '퍼즐') {
          code = generatePuzzleCode(item);
        } else if (ROLL_PRODUCT_IDS.includes(item.productId) || item.productCategory === '롤') {
          code = generateRollMatCode(item);
          if (item.lengthM > 0) code += `${item.lengthM}m`;
        } else {
          code = generateDesignCode(item);
        }
        const qty = item.quantity || 1;
        codeCountMap.set(code, (codeCountMap.get(code) || 0) + qty);
      });
      designWithQty = Array.from(codeCountMap.entries())
        .map(([code, qty]) => qty > 1 ? `${code}x${qty}` : code)
        .join('+');
    }

    // 상세주소(배송메모) - 출력용이므로 경고 표시 없음
    const deliveryMemoText = box.deliveryMemos && box.deliveryMemos.length > 0 ? box.deliveryMemos.join(' / ') : '';

    const rowData = [
      box.recipientLabel,   // 받는분
      box.address,          // 주소 (전체 주소)
      deliveryMemoText, // 상세주소 (실무팀 요청: 배송메모 기록)
      '', // 운송장번호 (공란)
      box.group.orderId,
      box.zipCode,
      '', // 도착영업소 (공란)
      box.phone,
      '', // 기타전화번호 (공란)
      '선불', // 선불후불 (기본 선불)
      designWithQty,       // 품목명 (디자인코드x수량)
      1, // 경동은 박스당 1개로 처리
      packagingStatus,
      dimensions.width,
      dimensions.height,
      dimensions.depth,
      '', // 무게
      50, // 개별단가 (고정값)
      shippingFee || '', // 배송운임 (배송비 계산값)
      100, // 기타운임 (고정값)
      0, // 별도운임 (고정값)
      0, // 할증운임 (고정값)
      0, // 도서운임 (고정값)
      designWithQty        // 메모 (실무팀 요청: 디자인코드/주문내용 줄임말)
    ];

    // 고객 식별 정보를 함께 저장 (색상 구분용)
    // 주소 + 전화번호 조합으로 고객 식별 (박스 단위가 아닌 고객 단위)
    rowData._customerKey = makeCustomerKey({ address: box.address, phone: box.phone });
    rowData._customerName = box.customerName || '';
    rowData._orderId = box.group.orderId || '';
    rowData._address = box.address || '';
    rowData._phone = box.phone || '';
    rowData._boxIndex = index;
    rowData._warnings = box.warnings || [];
    rowData._hasCuttingRequest = box.hasCuttingRequest || false;
    rowData._hasFinishingRequest = box.hasFinishingRequest || false;

    data.push(rowData);
  });

  return data;
}

// 경동택배 발주서 생성 (Excel 다운로드)
function exportToKyungdong() {
  const data = getKyungdongData();
  downloadExcel(data, '경동발주서');
}

// 로젠택배 발주서 데이터 생성
function getLozenData() {
  if (packedOrders.length === 0) {
    processPacking();
  }

  // 로젠택배 대상만 필터 (소박스 → courier === 'lozen')
  const lozenOrders = packedOrders.filter(box => box.courier === 'lozen');

  // 실무 양식 컬럼: 받는분, 받는분 전화번호, 받는분 핸드폰, 우편번호, 받는분 주소, 운임타입, 송장수량, 디자인+수량, 배송메세지, 합배송여부, 배송운임
  const headers = ['받는분', '받는분 전화번호', '받는분 핸드폰', '우편번호', '받는분 주소', '운임타입', '송장수량', '디자인+수량', '배송메세지', '합배송여부', '배송운임'];
  const data = [headers];

  lozenOrders.forEach((box, index) => {
    // 디자인+수량 형식 생성 (같은 디자인코드는 합쳐서 수량 표시)
    let designWithQty = box.designText;
    if (box.items && box.items.length > 0) {
      // 아이템별 디자인코드 생성 후 같은 코드끼리 수량 합산
      const codeCountMap = new Map();
      box.items.forEach(item => {
        let code = '';
        if (PUZZLE_PRODUCT_IDS.includes(item.productId) || item.productCategory === '퍼즐') {
          code = generatePuzzleCode(item);
        } else if (ROLL_PRODUCT_IDS.includes(item.productId) || item.productCategory === '롤') {
          code = generateRollMatCode(item);
          if (item.lengthM > 0) code += `${item.lengthM}m`;
        } else {
          code = generateDesignCode(item);
        }
        const qty = item.quantity || 1;
        codeCountMap.set(code, (codeCountMap.get(code) || 0) + qty);
      });
      // 코드x수량 형태로 조합
      designWithQty = Array.from(codeCountMap.entries())
        .map(([code, qty]) => qty > 1 ? `${code}x${qty}` : code)
        .join('+');
    }

    const shippingFee = calculateShippingFee(box); // 배송비 계산

    // 배송메세지 - 출력용이므로 경고 표시 없음
    const deliveryMemoText = box.deliveryMemos && box.deliveryMemos.length > 0 ? box.deliveryMemos.join(' / ') : '';

    const rowData = [
      box.recipientLabel,
      box.phone,
      box.phone,
      box.zipCode,
      box.address,
      '선불',
      1,
      designWithQty,
      deliveryMemoText,
      box.isCombined ? 'Y' : 'N',
      shippingFee || '' // 배송운임 (배송비 계산값)
    ];

    // 고객 식별 정보를 함께 저장 (색상 구분용)
    // 주소 + 전화번호 조합으로 고객 식별 (박스 단위가 아닌 고객 단위)
    rowData._customerKey = makeCustomerKey({ address: box.address, phone: box.phone });
    rowData._customerName = box.customerName || '';
    rowData._orderId = box.group.orderId || '';
    rowData._address = box.address || '';
    rowData._phone = box.phone || '';
    rowData._boxIndex = index;
    rowData._warnings = box.warnings || [];
    rowData._hasCuttingRequest = box.hasCuttingRequest || false;
    rowData._hasFinishingRequest = box.hasFinishingRequest || false;

    data.push(rowData);
  });

  return data;
}

// 로젠택배 발주서 생성 (Excel 다운로드)
function exportToLozen() {
  const data = getLozenData();
  downloadExcel(data, '로젠발주서');
}

// 송장화 데이터 생성 (Excel 다운로드)
function exportToCombined() {
  const data = getCombinedData();
  downloadExcel(data, '송장화');
}

// 가공데이터(송장화) 통합 데이터 생성 (주문별)
function getCombinedData() {
  if (packedOrders.length === 0) {
    processPacking();
  }

  // 컬럼: 상품주문번호, 수취인명, 운임타입, 송장수량, 디자인, 길이, 수량, 디자인+수량, 배송메세지, 합포장, 비고, 주문하신분 핸드폰, 받는분 핸드폰, 우편번호, 받는분 주소, 주문자명, 실제결제금액
  const headers = ['상품주문번호', '수취인명', '운임타입', '송장수량', '디자인', '길이', '수량', '디자인+수량', '배송메세지', '합포장', '비고', '주문하신분 핸드폰', '받는분 핸드폰', '우편번호', '받는분 주소', '주문자명', '실제결제금액'];
  const data = [headers];

  // 주문별로 그룹화 (groupedOrders 기반)
  Object.values(groupedOrders).forEach(group => {
    // 해당 그룹의 박스 수 계산 (송장수량)
    const groupBoxes = packedOrders.filter(box => box.group === group);
    const invoiceCount = groupBoxes.length || 1;

    // 주문의 모든 아이템을 합쳐서 디자인+수량 생성
    const codeCountMap = new Map();
    let totalQty = 0;
    let design = '';
    let length = '';

    group.items.forEach(item => {
      let code = '';
      if (PUZZLE_PRODUCT_IDS.includes(item.productId) || item.productCategory === '퍼즐') {
        code = generatePuzzleCode(item);
      } else if (ROLL_PRODUCT_IDS.includes(item.productId) || item.productCategory === '롤') {
        code = generateRollMatCode(item);
        if (item.lengthM > 0) code += `${item.lengthM}m`;
      } else {
        code = generateDesignCode(item);
      }
      const qty = item.quantity || 1;
      codeCountMap.set(code, (codeCountMap.get(code) || 0) + qty);
      totalQty += qty;

      // 첫 번째 아이템에서 디자인, 길이 추출
      if (!design && item.design) design = item.design;
      if (!length && item.length) length = item.length;
    });

    const designWithQty = Array.from(codeCountMap.entries())
      .map(([code, qty]) => qty > 1 ? `${code}x${qty}` : code)
      .join('+');

    // 배송메모 수집 (중복 제거)
    const deliveryMemos = group.items
      .map(item => item.deliveryMemo)
      .filter(memo => memo && memo.trim())
      .filter((memo, index, arr) => arr.indexOf(memo) === index);
    const deliveryMemoText = deliveryMemos.join(' / ');

    // 합포장 여부 (박스가 2개 이상이면 합포장)
    const isCombined = invoiceCount > 1 ? 'Y' : 'N';

    // 실제결제금액: 주문의 모든 상품의 '최종 상품별 총 주문금액' 합산
    const actualPayment = group.items.reduce((sum, item) => sum + (item.price || 0), 0);

    const rowData = [
      group.orderId || '',                // 상품주문번호 (주문번호)
      group.customerName || '',            // 수취인명
      '선불',                              // 운임타입
      invoiceCount,                        // 송장수량 (박스 개수)
      design,                              // 디자인
      length,                              // 길이
      totalQty,                           // 수량 (전체 수량)
      designWithQty,                      // 디자인+수량
      deliveryMemoText,                   // 배송메세지
      isCombined,                         // 합포장
      '',                                 // 비고
      group.phone || '',                  // 주문하신분 핸드폰
      group.phone || '',                  // 받는분 핸드폰
      group.zipCode || '',                // 우편번호
      group.address || '',                // 받는분 주소
      group.customerName || '',           // 주문자명
      actualPayment                       // 실제결제금액 (정산예정금액)
    ];

    // 메타데이터 (화면 표시용)
    rowData._customerKey = makeCustomerKey({ address: group.address, phone: group.phone });
    rowData._hasCuttingRequest = group.hasCuttingRequest || false;
    rowData._hasFinishingRequest = group.hasFinishingRequest || false;

    data.push(rowData);
  });

  return data;
}

// 가공데이터(송장화) 테이블 렌더링
function renderCombinedTable() {
  const table = document.getElementById('combinedTable');
  const thead = table.querySelector('thead');
  const tbody = document.getElementById('combinedTableBody');

  thead.innerHTML = '';
  tbody.innerHTML = '';

  if (packedOrders.length === 0) {
    tbody.innerHTML = '<tr><td colspan="17" style="text-align: center; padding: 2rem; color: var(--text-muted);">데이터가 없습니다.</td></tr>';
    return;
  }

  const data = getCombinedData();
  if (data.length < 2) {
    tbody.innerHTML = '<tr><td colspan="17" style="text-align: center; padding: 2rem; color: var(--text-muted);">유효한 데이터가 없습니다.</td></tr>';
    return;
  }

  // 헤더 생성
  const headers = data[0];
  const trHead = document.createElement('tr');
  headers.forEach(header => {
    const th = document.createElement('th');
    th.className = 'tooltip-header';
    
    // 헤더 텍스트와 아이콘 추가
    const headerText = document.createTextNode(header);
    th.appendChild(headerText);
    
    // Tooltip 추가
    if (COLUMN_TOOLTIPS[header]) {
      th.setAttribute('data-tooltip', COLUMN_TOOLTIPS[header]);
      th.style.cursor = 'help';
      
      // "i" 아이콘 추가
      const icon = document.createElement('span');
      icon.className = 'tooltip-icon';
      icon.textContent = 'ⓘ';
      icon.setAttribute('aria-label', '도움말');
      th.appendChild(icon);
    }
    
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);

  // 데이터 생성 (고객별로 색상 구분)
  const customerGroupIndexByKey = new Map();
  let nextGroupIndex = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tr = document.createElement('tr');

    // 고객 식별 키 생성
    const customerKey = row._customerKey || '';

    // 같은 고객 키는 항상 같은 그룹 인덱스 사용
    let groupIndex = 0;
    if (customerKey) {
      if (!customerGroupIndexByKey.has(customerKey)) {
        nextGroupIndex += 1;
        customerGroupIndexByKey.set(customerKey, nextGroupIndex);
      }
      groupIndex = customerGroupIndexByKey.get(customerKey);
    }

    // 그룹별 배경색 적용 (흰색/회색 번갈아)
    if (groupIndex % 2 === 0) {
      tr.classList.add('raw-row-group-even');
    } else {
      tr.classList.add('raw-row-group-odd');
    }

    // 확인 필요 행 (붉은색) - 애견롤매트 재단요청
    if (row._hasCuttingRequest) {
      tr.classList.add('anomaly-row');
    }
    // 마감재 요청 행 (노란색) - 퍼즐매트 마감재
    if (row._hasFinishingRequest) {
      tr.classList.add('finishing-row');
    }

    headers.forEach((header, index) => {
      const td = document.createElement('td');
      td.textContent = row[index] !== undefined && row[index] !== null ? row[index] : '';
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  }
}

// 경동 발주서 테이블 렌더링
function renderKyungdongTable() {
  const table = document.getElementById('kyungdongTable');
  const thead = table.querySelector('thead');
  const tbody = document.getElementById('kyungdongTableBody');

  thead.innerHTML = '';
  tbody.innerHTML = '';

  // 경동택배 물량이 없는 경우 (대박스/비닐 없음)
  const kyungdongOrders = packedOrders.filter(box => box.courier === 'kyungdong');
  if (kyungdongOrders.length === 0) {
    tbody.innerHTML = '<tr><td colspan="24" style="text-align: center; padding: 2rem; color: var(--text-muted);">경동택배 물량이 없습니다. (대박스/비닐 대상 없음)</td></tr>';
    return;
  }

  const data = getKyungdongData();
  if (data.length < 2) {
    tbody.innerHTML = '<tr><td colspan="24" style="text-align: center; padding: 2rem; color: var(--text-muted);">유효한 데이터가 없습니다.</td></tr>';
    return;
  }

  // 헤더 생성
  const headers = data[0];
  const trHead = document.createElement('tr');
  headers.forEach(header => {
    const th = document.createElement('th');
    th.className = 'tooltip-header';
    
    // 헤더 텍스트와 아이콘 추가
    const headerText = document.createTextNode(header);
    th.appendChild(headerText);
    
    // Tooltip 추가
    if (COLUMN_TOOLTIPS[header]) {
      th.setAttribute('data-tooltip', COLUMN_TOOLTIPS[header]);
      th.style.cursor = 'help';
      
      // "i" 아이콘 추가
      const icon = document.createElement('span');
      icon.className = 'tooltip-icon';
      icon.textContent = 'ⓘ';
      icon.setAttribute('aria-label', '도움말');
      th.appendChild(icon);
    }
    
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);

  // 데이터 생성 (고객별로 색상 구분)
  // 같은 고객이면 어디에 있든 동일한 색상(그룹) 유지
  const customerGroupIndexByKey = new Map();
  let nextGroupIndex = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tr = document.createElement('tr');

    // 고객 식별 키 생성 (우선순위: 저장된 키 > 주소+전화번호 조합 > 받는분 라벨 파싱)
    let customerKey = '';

    if (row._customerKey) {
      // getKyungdongData에서 저장한 고객 키 사용
      customerKey = row._customerKey;
    } else {
      // 폴백: 주소 + 전화번호 조합
      const address = row._address || row[1] || ''; // 주소 컬럼
      const phone = row._phone || row[7] || ''; // 전화번호 컬럼
      customerKey = makeCustomerKey({ address, phone });

      // 여전히 없으면 받는분 라벨에서 고객명 추출하여 키 생성
      if (!customerKey || customerKey === '||') {
        const recipientLabel = row[0] || '';
        const customerName = recipientLabel.replace(/^\d+-\d+/, '').replace(/★$/, '').trim();
        customerKey = customerName ? `name:${customerName}` : '';
      }
    }

    // 같은 고객 키는 항상 같은 그룹 인덱스 사용
    let groupIndex = 0;
    if (customerKey) {
      if (!customerGroupIndexByKey.has(customerKey)) {
        nextGroupIndex += 1;
        customerGroupIndexByKey.set(customerKey, nextGroupIndex);
      }
      groupIndex = customerGroupIndexByKey.get(customerKey);
    }

    // 그룹별 배경색 적용 (흰색/회색 번갈아)
    if (groupIndex % 2 === 0) {
      tr.classList.add('raw-row-group-even');
    } else {
      tr.classList.add('raw-row-group-odd');
    }

    // 확인 필요 행 (붉은색) - 애견롤매트 재단요청
    if (row._hasCuttingRequest) {
      tr.classList.add('anomaly-row');
    }
    // 마감재 요청 행 (노란색) - 퍼즐매트 마감재
    if (row._hasFinishingRequest) {
      tr.classList.add('finishing-row');
    }

    headers.forEach((header, index) => {
      const td = document.createElement('td');
      td.textContent = row[index] || '';
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  }
}

// 로젠 발주서 테이블 렌더링
function renderLozenTable() {
  const table = document.getElementById('lozenTable');
  const thead = table.querySelector('thead');
  const tbody = document.getElementById('lozenTableBody');

  thead.innerHTML = '';
  tbody.innerHTML = '';

  // 로젠택배 물량이 없는 경우 (소박스 없음)
  const lozenOrders = packedOrders.filter(box => box.courier === 'lozen');
  if (lozenOrders.length === 0) {
    tbody.innerHTML = '<tr><td colspan="11" style="text-align: center; padding: 2rem; color: var(--text-muted);">로젠택배 물량이 없습니다. (소박스 대상 없음)</td></tr>';
    return;
  }

  const data = getLozenData();
  if (data.length < 2) {
    tbody.innerHTML = '<tr><td colspan="11" style="text-align: center; padding: 2rem; color: var(--text-muted);">유효한 데이터가 없습니다.</td></tr>';
    return;
  }

  // 헤더 생성
  const headers = data[0];
  const trHead = document.createElement('tr');
  headers.forEach(header => {
    const th = document.createElement('th');
    th.className = 'tooltip-header';
    
    // 헤더 텍스트와 아이콘 추가
    const headerText = document.createTextNode(header);
    th.appendChild(headerText);
    
    // Tooltip 추가
    if (COLUMN_TOOLTIPS[header]) {
      th.setAttribute('data-tooltip', COLUMN_TOOLTIPS[header]);
      th.style.cursor = 'help';
      
      // "i" 아이콘 추가
      const icon = document.createElement('span');
      icon.className = 'tooltip-icon';
      icon.textContent = 'ⓘ';
      icon.setAttribute('aria-label', '도움말');
      th.appendChild(icon);
    }
    
    trHead.appendChild(th);
  });
  thead.appendChild(trHead);

  // 데이터 생성 (고객별로 색상 구분)
  // 같은 고객이면 어디에 있든 동일한 색상(그룹) 유지
  const customerGroupIndexByKey = new Map();
  let nextGroupIndex = 0;

  for (let i = 1; i < data.length; i++) {
    const row = data[i];
    const tr = document.createElement('tr');

    // 고객 식별 키 생성 (우선순위: 저장된 키 > 주소+전화번호 조합 > 받는분 라벨 파싱)
    let customerKey = '';

    if (row._customerKey) {
      // getLozenData에서 저장한 고객 키 사용
      customerKey = row._customerKey;
    } else {
      // 로젠 데이터 구조: 받는분, 받는분 전화번호, 받는분 핸드폰, 우편번호, 받는분 주소, ...
      const phone = row._phone || row[1] || row[2] || ''; // 전화번호 컬럼
      const address = row._address || row[4] || ''; // 주소 컬럼
      customerKey = makeCustomerKey({ address, phone });

      // 여전히 없으면 받는분 라벨에서 고객명 추출하여 키 생성
      if (!customerKey || customerKey === '||') {
        const recipientLabel = row[0] || '';
        const customerName = recipientLabel.replace(/^\d+-\d+/, '').replace(/★$/, '').trim();
        customerKey = customerName ? `name:${customerName}` : '';
      }
    }

    // 같은 고객 키는 항상 같은 그룹 인덱스 사용
    let groupIndex = 0;
    if (customerKey) {
      if (!customerGroupIndexByKey.has(customerKey)) {
        nextGroupIndex += 1;
        customerGroupIndexByKey.set(customerKey, nextGroupIndex);
      }
      groupIndex = customerGroupIndexByKey.get(customerKey);
    }

    // 그룹별 배경색 적용 (흰색/회색 번갈아)
    if (groupIndex % 2 === 0) {
      tr.classList.add('raw-row-group-even');
    } else {
      tr.classList.add('raw-row-group-odd');
    }

    // 확인 필요 행 (붉은색) - 애견롤매트 재단요청
    if (row._hasCuttingRequest) {
      tr.classList.add('anomaly-row');
    }
    // 마감재 요청 행 (노란색) - 퍼즐매트 마감재
    if (row._hasFinishingRequest) {
      tr.classList.add('finishing-row');
    }

    headers.forEach((header, index) => {
      const td = document.createElement('td');
      td.textContent = row[index] || '';
      tr.appendChild(td);
    });

    tbody.appendChild(tr);
  }
}

// 주소 분리
function splitAddress(address) {
  if (!address) return { base: '', detail: '' };

  // 괄호 안의 우편번호 제거
  const cleanAddress = address.replace(/\(\d{5}\)/, '').trim();

  // 시/도, 구/군/시, 동/읍/면/리 까지를 기본주소로
  const match = cleanAddress.match(/^(.+?(?:시|도)\s*.+?(?:구|군|시)\s*.+?(?:동|읍|면|리|로|길))\s*(.*)$/);

  if (match) {
    return {
      base: match[1].trim(),
      detail: match[2].trim()
    };
  }

  // 매칭 실패시 전체를 기본주소로
  return { base: cleanAddress, detail: '' };
}

function normalizePhone(phone) {
  const digits = String(phone || '').replace(/\D/g, '');
  if (!digits) return '';
  // 국내번호 형태를 최대한 유지: +82/82로 시작하면 0을 붙여 정규화
  if (digits.startsWith('82') && digits.length >= 10) {
    return `0${digits.slice(2)}`;
  }
  return digits;
}

function normalizeAddress(address) {
  return String(address || '')
    .replace(/\\/g, '') // CSV/엑셀에서 들어온 '\' 제거
    .replace(/\(\d{5}\)/g, '') // (우편번호) 제거
    .replace(/\s+/g, ' ')
    .trim();
}

function makeCustomerKey({ address, phone }) {
  const normalizedAddress = normalizeAddress(address);
  const normalizedPhone = normalizePhone(phone);
  if (!normalizedAddress && !normalizedPhone) return '';
  return `${normalizedAddress}|${normalizedPhone}`;
}

// 박스 치수 계산 (임시)
function getDimensions(box) {
  // 실제로는 제품별로 치수가 다르지만, 기본값 반환
  if (box.type === 'puzzle') {
    return { width: 100, height: 100, depth: 30 };
  }
  if (box.packagingType === 'vinyl') {
    return { width: 150, height: 20, depth: 20 };
  }
  return { width: 120, height: 15, depth: 15 };
}

// Excel 파일 다운로드
function downloadExcel(data, filename) {
  // 메타데이터 제거 (고객명, 박스 인덱스 등)
  const cleanData = data.map((row, rowIndex) => {
    if (rowIndex === 0) return row; // 헤더는 그대로
    if (Array.isArray(row)) {
      // 배열인 경우 그대로 반환 (메타데이터는 배열에 속성으로 추가되므로 영향 없음)
      return row;
    }
    return row;
  });

  const ws = XLSX.utils.aoa_to_sheet(cleanData);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

  const today = new Date();
  const dateStr = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;

  XLSX.writeFile(wb, `${filename}_${dateStr}.xlsx`);
}

// ========== 렌더링 ==========
function renderOrders() {
  renderTableView();
}

// 패킹 상세 모달
function showPackingDetail(box) {
  const itemsHtml = box.items.map(item => {
    // 원본 데이터 표시용
    const rawData = item.rawRow || {};
    const rawInfo = rawData['옵션정보'] || rawData['상품명'] || '';

    return `
    <div style="padding: 0.5rem; background: #f5f5f5; border-radius: 4px; margin-bottom: 0.5rem;">
      <strong>${item.productType}</strong> ${item.design || item.productName}<br>
      <small>두께: ${item.thickness} / 폭: ${item.width} / 길이: ${item.length} / 수량: ${item.quantity}</small>
      ${rawInfo ? `<div style="margin-top: 0.25rem; font-size: 0.7rem; color: #999; font-style: italic;">원본: ${rawInfo}</div>` : ''}
    </div>
  `;
  }).join('');

  modalBody.innerHTML = `
    <div class="order-detail">
      <div class="detail-group">
        <label>수취인 라벨</label>
        <p><strong style="font-size: 1.2rem;">${box.recipientLabel}</strong></p>
      </div>
      <div class="detail-group">
        <label>연락처</label>
        <p>${box.phone}</p>
      </div>
      <div class="detail-group">
        <label>주소</label>
        <p>${box.address}</p>
      </div>
      <hr>
      <div class="detail-group">
        <label>포장 내용물</label>
        ${itemsHtml}
      </div>
      <div class="detail-group">
        <label>디자인 코드</label>
        <p><code style="background: #e0e0e0; padding: 0.25rem 0.5rem; border-radius: 4px;">${box.designText}</code></p>
      </div>
      <div class="detail-group">
        <label>포장 방식</label>
        <p>${box.packagingType === 'vinyl' ? '비닐 (파손주의)' : box.packagingType === 'largeBox' ? '대박스' : '박스'}</p>
      </div>
	      ${box.deliveryMemos && box.deliveryMemos.length > 0 ? `
	        <div class="detail-group">
	          <label>배송메모${box.deliveryMemos.length > 1 ? ` (${box.deliveryMemos.length}건)` : ''}</label>
	          <p style="color: #c0392b; white-space: pre-line;">${box.deliveryMemos.join('\n')}</p>
	        </div>
	      ` : ''}
	      ${box.warnings && box.warnings.length > 0 ? `
	        <div class="detail-group">
	          <label>이상감지</label>
	          <p style="color: #b91c1c; font-weight: 600; white-space: pre-line;">${box.warnings.join('\n')}</p>
	        </div>
	      ` : ''}
	      ${box.giftEligible ? `
	        <div class="detail-group">
	          <label>사은품</label>
	          <p style="color: #27ae60; font-weight: bold;">${box.giftEligible}</p>
	        </div>
	      ` : ''}
    </div>
  `;

  orderModal.style.display = 'flex';
}

// 테이블 뷰 렌더링
function renderTableView() {
  // ordersTableBody가 HTML에 없으므로 null 체크
  if (!ordersTableBody) return;

  ordersTableBody.innerHTML = '';

  filteredOrders.forEach(order => {
    const tr = document.createElement('tr');
    tr.innerHTML = `
      <td><input type="checkbox" class="order-checkbox" data-id="${order.id}"></td>
      <td>
        <a href="#" class="order-link" data-id="${order.id}">${order.orderId.slice(-8)}</a>
      </td>
      <td>${order.customerName}</td>
      <td>
        <span class="product-badge ${getProductBadgeClass(order.productType)}">${order.productType}</span>
        <span class="design-name">${order.design}</span>
      </td>
      <td class="option-display">
        <strong>${order.thickness}</strong> / ${order.width} / ${order.length}
      </td>
      <td>${order.quantity}</td>
      <td>${formatCurrency(order.price)}</td>
      <td>${order.gift ? '<span class="gift-badge">있음</span>' : '-'}</td>
      <td><span class="status-badge ${getStatusClass(order.status)}">${order.status}</span></td>
      <td class="delivery-memo" title="${order.deliveryMemo}">${order.deliveryMemo || '-'}</td>
    `;

    // 주문 상세 클릭 이벤트
    tr.querySelector('.order-link').addEventListener('click', (e) => {
      e.preventDefault();
      showOrderDetail(order);
    });

    ordersTableBody.appendChild(tr);
  });
}

function getProductBadgeClass(type) {
  if (type === '유아롤매트') return 'roll-baby';
  if (type === '애견롤매트') return 'roll-pet';
  if (type === '퍼즐매트') return 'puzzle';
  return '';
}

function getStatusClass(status) {
  if (status === '발송대기') return 'pending';
  if (status === '배송중') return 'shipping';
  if (status === '배송완료') return 'completed';
  return 'pending';
}

function formatCurrency(amount) {
  return new Intl.NumberFormat('ko-KR', {
    style: 'currency',
    currency: 'KRW',
    maximumFractionDigits: 0
  }).format(amount);
}

// ========== 필터 ==========
function applyFilters() {
  filteredOrders = ordersData.filter(order => {
    // 제품 유형 필터
    if (productTypeFilter.value && order.productType !== productTypeFilter.value) {
      return false;
    }

    // 주문 상태 필터
    if (orderStatusFilter.value && order.status !== orderStatusFilter.value) {
      return false;
    }

    // 사은품 필터
    if (giftFilter.value === '있음' && (!order.gift || order.gift.trim() === '')) {
      return false;
    }
    if (giftFilter.value === '없음' && order.gift && order.gift.trim() !== '') {
      return false;
    }

    // 검색 필터
    if (searchInput.value) {
      const query = searchInput.value.toLowerCase();
      const searchFields = [
        order.orderId,
        order.customerName,
        order.design,
        order.productType
      ].join(' ').toLowerCase();

      if (!searchFields.includes(query)) {
        return false;
      }
    }

    return true;
  });


  // 필터링된 데이터로 다시 그룹화 및 패킹 처리
  groupOrdersByRecipient();
  processPacking();
  updateSummary();
  renderOrders();
  rerenderActiveTab();
}

// ========== 체크박스 ==========
function handleSelectAll(e) {
  const checkboxes = document.querySelectorAll('.order-checkbox');
  checkboxes.forEach(cb => {
    cb.checked = e.target.checked;
    cb.closest('tr').classList.toggle('selected', e.target.checked);
  });
}

// ========== 모달 ==========
function showOrderDetail(order) {
  const packaging = determineRollPackaging(order);
  const packagingText = packaging.type === 'vinyl' ? '비닐' :
    packaging.type === 'largeBox' ? '대박스' :
      packaging.type === 'combine' ? '합포장' : '소박스';

  modalBody.innerHTML = `
    <div class="order-detail">
      <div class="detail-group">
        <label>주문번호</label>
        <p>${order.orderId}</p>
      </div>
      <div class="detail-group">
        <label>고객명 / 연락처</label>
        <p>${order.customerName} / ${order.phone}</p>
      </div>
      <div class="detail-group">
        <label>배송지</label>
        <p>${order.address}</p>
      </div>
      <hr>
      <div class="detail-group">
        <label>제품</label>
        <p><span class="product-badge ${getProductBadgeClass(order.productType)}">${order.productType}</span> ${order.design}</p>
      </div>
      <div class="detail-group">
        <label>옵션</label>
        <p>두께: <strong>${order.thickness}</strong> / 폭: <strong>${order.width}</strong> / 길이: <strong>${order.length}</strong></p>
        ${order.rawRow ? (() => {
      const rawInfo = order.rawRow['옵션정보'] || order.rawRow['상품명'] || '';
      return rawInfo ? `<div style="margin-top: 0.25rem; font-size: 0.7rem; color: #999; font-style: italic;">원본: ${rawInfo}</div>` : '';
    })() : ''}
      </div>
      <div class="detail-group">
        <label>수량</label>
        <p>${order.quantity}</p>
      </div>
      <div class="detail-group">
        <label>포장방식</label>
        <p>${packagingText} ${packaging.needsStar ? '★ (파손주의)' : ''}</p>
      </div>
      <div class="detail-group">
        <label>금액</label>
        <p><strong>${formatCurrency(order.price)}</strong></p>
      </div>
      ${order.gift ? `
        <div class="detail-group">
          <label>사은품</label>
          <p>${order.gift}</p>
        </div>
      ` : ''}
      ${order.deliveryMemo ? `
        <div class="detail-group">
          <label>배송메모</label>
          <p>${order.deliveryMemo}</p>
        </div>
      ` : ''}
    </div>
  `;

  orderModal.style.display = 'flex';
}

function closeModal() {
  orderModal.style.display = 'none';
}

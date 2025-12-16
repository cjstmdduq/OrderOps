// ========== 전역 상태 ==========
let ordersData = [];
let filteredOrders = [];
let groupedOrders = {}; // 고객별 그룹화된 주문
let packedOrders = []; // 패킹 처리된 결과
let currentView = 'packing'; // 'customer' | 'packing'

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
  "40": 4
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
  "퓨어아이보리": "퓨어아이",
  "마블아이보리": "마블아이",
  "그레이캔버스": "그레이캔",
  "딜라이트우드": "딜라이트",
  "모던그레이": "모던그레",
  "스노우화이트": "스노우",
  "내추럴우드": "내추럴"
};

// ========== DOM 요소 ==========
const uploadArea = document.getElementById('uploadArea');
const fileInput = document.getElementById('fileInput');
const uploadBtn = document.getElementById('uploadBtn');
const fileInfo = document.getElementById('fileInfo');
const fileName = document.getElementById('fileName');
const clearBtn = document.getElementById('clearBtn');

const summarySection = document.getElementById('summarySection');
const filterSection = document.getElementById('filterSection');
const ordersSection = document.getElementById('ordersSection');
const emptyState = document.getElementById('emptyState');
const ordersTableBody = document.getElementById('ordersTableBody');
const customerCardsContainer = document.getElementById('customerCardsContainer');
const tableContainer = document.getElementById('tableContainer');

const totalOrders = document.getElementById('totalOrders');
const totalCustomers = document.getElementById('totalCustomers');
const giftOrders = document.getElementById('giftOrders');
const totalAmount = document.getElementById('totalAmount');

const productTypeFilter = document.getElementById('productTypeFilter');
const orderStatusFilter = document.getElementById('orderStatusFilter');
const giftFilter = document.getElementById('giftFilter');
const searchInput = document.getElementById('searchInput');
const selectAll = document.getElementById('selectAll');

const customerViewBtn = document.getElementById('customerViewBtn');
const packingViewBtn = document.getElementById('packingViewBtn');

const orderModal = document.getElementById('orderModal');
const modalClose = document.getElementById('modalClose');
const modalBody = document.getElementById('modalBody');

// ========== 초기화 ==========
document.addEventListener('DOMContentLoaded', () => {
  initEventListeners();
  updateDateDisplay();
});

function initEventListeners() {
  // 파일 업로드
  uploadBtn.addEventListener('click', () => fileInput.click());
  uploadArea.addEventListener('click', (e) => {
    if (e.target !== uploadBtn) fileInput.click();
  });
  fileInput.addEventListener('change', handleFileSelect);
  clearBtn.addEventListener('click', clearFile);

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

  // 뷰 전환
  customerViewBtn.addEventListener('click', () => switchView('customer'));
  packingViewBtn.addEventListener('click', () => switchView('packing'));

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
  const exportKyungdongBtn = document.getElementById('exportKyungdongBtn');
  const exportLozenBtn = document.getElementById('exportLozenBtn');

  if (exportKyungdongBtn) {
    exportKyungdongBtn.addEventListener('click', () => exportToKyungdong());
  }
  if (exportLozenBtn) {
    exportLozenBtn.addEventListener('click', () => exportToLozen());
  }
}

function switchView(view) {
  currentView = view;

  // 버튼 활성화 상태
  customerViewBtn.classList.toggle('active', view === 'customer');
  packingViewBtn.classList.toggle('active', view === 'packing');

  // 뷰 표시 전환
  if (view === 'customer') {
    // 고객별 카드 뷰
    customerCardsContainer.style.display = 'flex';
    document.getElementById('packingResultContainer').style.display = 'none';
  } else {
    // 발주 양식 뷰 (패킹 테이블)
    customerCardsContainer.style.display = 'none';
    document.getElementById('packingResultContainer').style.display = 'block';
  }
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
    const text = e.target.result;
    parseCSV(text);
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

  for (let i = 1; i < data.length; i++) {
    const values = data[i];
    if (!values || values.length === 0) continue;

    const row = {};
    headers.forEach((header, index) => {
      row[header] = values[index] !== undefined ? String(values[index]).trim() : '';
    });

    const order = parseOrderData(row);
    if (order) ordersData.push(order);
  }

  filteredOrders = [...ordersData];
  updateUI();
}

function parseCSV(text) {
  const lines = text.split('\n');
  const headers = parseCSVLine(lines[0]);

  ordersData = [];

  for (let i = 1; i < lines.length; i++) {
    if (lines[i].trim() === '') continue;

    const values = parseCSVLine(lines[i]);
    const row = {};

    headers.forEach((header, index) => {
      row[header.trim()] = values[index] ? values[index].trim() : '';
    });

    // 데이터 정제
    const order = parseOrderData(row);
    if (order) ordersData.push(order);
  }

  filteredOrders = [...ordersData];
  updateUI();
}

function parseCSVLine(line) {
  const result = [];
  let current = '';
  let inQuotes = false;

  for (let i = 0; i < line.length; i++) {
    const char = line[i];

    if (char === '"') {
      inQuotes = !inQuotes;
    } else if (char === ',' && !inQuotes) {
      result.push(current);
      current = '';
    } else {
      current += char;
    }
  }
  result.push(current);

  return result.map(val => val.replace(/^"|"$/g, '').trim());
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
    const length = parsed.length || parsedFromName.length || '';

    // 상품번호 기반 제품 타입 판별
    const productMapping = PRODUCT_TYPE_MAP[productId];
    const productType = productMapping ? productMapping.name : getProductTypeByName(productName);
    const productCategory = productMapping ? productMapping.type : '기타';
    const productCategoryCode = productMapping ? productMapping.category : 'etc';

    // 재단 요청 감지
    const hasCuttingRequest = detectCuttingRequest(deliveryMemo);

    // 마감재 요청 감지
    const hasFinishingRequest = detectFinishingRequest(deliveryMemo);

    // 길이값 추출 (m 단위)
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
    return `플러스♥${thickness}T${color}(${size})`;
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

  // 3. "길이(수량추가): 50cm" 패턴은 수량이므로 무시
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

  // 두께 추출: 6T, 9T, 10T, 12T, 17T, 22T 등
  const thicknessMatch = productName.match(/(\d+)T\b/i);
  if (thicknessMatch) {
    result.thickness = thicknessMatch[1] + 'T';
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

function parsePrice(priceStr) {
  if (!priceStr) return 0;
  return parseInt(priceStr.replace(/[^0-9]/g, '')) || 0;
}

function clearFile() {
  fileInput.value = '';
  fileInfo.style.display = 'none';
  ordersData = [];
  filteredOrders = [];
  packedOrders = [];
  updateUI();
}

// ========== UI 업데이트 ==========
function updateUI() {
  if (ordersData.length === 0) {
    summarySection.style.display = 'none';
    filterSection.style.display = 'none';
    ordersSection.style.display = 'none';
    emptyState.style.display = 'block';
    return;
  }

  summarySection.style.display = 'block';
  filterSection.style.display = 'flex';
  ordersSection.style.display = 'block';
  emptyState.style.display = 'none';

  groupOrdersByRecipient();
  processPacking(); // 자동 패킹 처리
  updateSummary();
  renderOrders();
}

function updateSummary() {
  const uniqueCustomers = Object.keys(groupedOrders).length;
  const total = ordersData.reduce((sum, o) => sum + o.price, 0);

  // 사은품 대상자 계산 (20만원+, 30만원+)
  const giftEligible = Object.values(groupedOrders).filter(g => g.totalPrice >= 200000).length;

  totalOrders.textContent = packedOrders.length; // 총 박스 수
  totalCustomers.textContent = uniqueCustomers; // 고객 수
  giftOrders.textContent = giftEligible;
  totalAmount.textContent = formatCurrency(total);
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
    if (ROLL_PRODUCT_IDS.includes(order.productId) || order.productCategory === '롤') {
      groupedOrders[key].rollItems.push(order);
    } else if (PUZZLE_PRODUCT_IDS.includes(order.productId) || order.productCategory === '퍼즐') {
      groupedOrders[key].puzzleItems.push(order);
    } else if (TAPE_PRODUCT_IDS.includes(order.productId) || order.productCategory === '테이프') {
      groupedOrders[key].tapeItems.push(order);
    } else {
      groupedOrders[key].otherItems.push(order);
    }

    if (order.gift && order.gift.trim() !== '') {
      groupedOrders[key].hasGift = true;
      groupedOrders[key].giftInfo = order.gift;
    }

    if (order.deliveryMemo && order.deliveryMemo.trim() !== '') {
      groupedOrders[key].deliveryMemo = order.deliveryMemo;
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

    if (rollMatTotalPrice >= 300000) {
      group.giftEligible = '발매트+테이프';
    } else if (rollMatTotalPrice >= 200000) {
      group.giftEligible = '실리콘테이프';
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

// 패킹 처리 메인 함수
function processPacking() {
  packedOrders = [];

  Object.values(groupedOrders).forEach(group => {
    const boxes = [];

    // 1. 롤매트 처리
    let combinableItems = []; // 합포장 가능한 아이템들 (별도 박스 없이 합침)
    let standaloneBoxes = []; // 독립적인 박스들

    group.rollItems.forEach(item => {
      const packaging = determineRollPackaging(item);

      if (packaging.canCombine) {
        combinableItems.push({ ...item, packaging });
      } else {
        // 단독 박스 생성
        for (let i = 0; i < item.quantity; i++) {
          const designCode = generateRollMatCode(item);
          standaloneBoxes.push({
            type: 'roll',
            packagingType: packaging.type,
            needsStar: packaging.needsStar,
            items: [item],
            designText: `${designCode}${item.lengthM}m`,
            isCombined: false,
            remark: ''
          });
        }
      }
    });

    // 테이프는 무조건 합포장
    group.tapeItems.forEach(item => {
      combinableItems.push({ ...item, packaging: { type: 'smallBox', canCombine: true } });
    });

    // 합포장 아이템 처리
    if (combinableItems.length > 0) {
      if (standaloneBoxes.length > 0) {
        // 기존 박스(마지막)에 합포장
        // 배송메모에 '합포장 표시'가 필요할 수도 있음
        const lastBox = standaloneBoxes[standaloneBoxes.length - 1];

        combinableItems.forEach(item => {
          lastBox.items.push(item);
        });

        lastBox.isCombined = true;
        lastBox.remark = '합';

        // 디자인 텍스트 업데이트
        const combineDesigns = combinableItems.map(i => generateDesignCode(i)).join('+');
        lastBox.designText += '+' + combineDesigns;

      } else {
        // 담을 박스가 없으면 새 소박스 생성
        const combineDesigns = combinableItems.map(i => generateDesignCode(i)).join('+');
        standaloneBoxes.push({
          type: 'roll',
          packagingType: 'smallBox',
          needsStar: false,
          items: combinableItems,
          designText: combineDesigns,
          isCombined: true,
          remark: '합'
        });
      }
    }

    // 최종 박스 리스트에 추가
    boxes.push(...standaloneBoxes);

    // 2. 퍼즐매트 처리 (롤매트와 별도)
    if (group.puzzleItems.length > 0) {
      // 원본 배열을 복사해서 사용 (원본 수정 방지)
      const puzzleItemsCopy = group.puzzleItems.map(item => ({ ...item }));
      const puzzleBoxCount = calculatePuzzleBoxes(puzzleItemsCopy);
      const puzzleDesign = puzzleItemsCopy.map(p => generatePuzzleCode(p)).join('+');

      // 퍼즐매트 아이템을 박스별로 분배
      const thickness = puzzleItemsCopy[0].thicknessNum || 25;
      const capacity = PUZZLE_BOX_CAPACITY[String(thickness)] || 6;

      let itemIndex = 0;
      for (let i = 0; i < puzzleBoxCount; i++) {
        const boxItems = [];
        let remainingCapacity = capacity;

        // 각 박스에 capacity만큼 아이템 할당
        while (itemIndex < puzzleItemsCopy.length && remainingCapacity > 0) {
          const item = puzzleItemsCopy[itemIndex];
          const takeQty = Math.min(item.quantity, remainingCapacity);

          if (takeQty > 0) {
            boxItems.push({
              ...item,
              quantity: takeQty
            });
            remainingCapacity -= takeQty;

            // 아이템의 수량을 줄임
            if (takeQty >= item.quantity) {
              itemIndex++;
            } else {
              item.quantity -= takeQty;
            }
          } else {
            itemIndex++;
          }
        }

        // 아이템이 있는 박스만 추가
        if (boxItems.length > 0 && boxItems.reduce((sum, item) => sum + item.quantity, 0) > 0) {
          boxes.push({
            type: 'puzzle',
            packagingType: 'vinyl',
            needsStar: false,
            items: boxItems,
            designText: puzzleDesign,
            isCombined: false,
            remark: ''
          });
        }
      }
    }

    // 3. 기타 아이템 처리
    group.otherItems.forEach(item => {
      for (let i = 0; i < item.quantity; i++) {
        boxes.push({
          type: 'other',
          packagingType: 'box',
          needsStar: false,
          items: [item],
          designText: generateDesignCode(item),
          isCombined: false,
          remark: ''
        });
      }
    });

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
      packedOrders.push({
        ...box,
        group: group,
        customerName: group.customerName,
        address: group.address,
        zipCode: group.zipCode,
        phone: group.phone,
        deliveryMemo: group.deliveryMemo,
        totalPrice: group.totalPrice,
        giftEligible: group.giftEligible,
        hasCuttingRequest: group.hasCuttingRequest,
        hasFinishingRequest: group.hasFinishingRequest
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

// ========== 발주서 생성 ==========

// 경동택배 발주서 생성
function exportToKyungdong() {
  if (packedOrders.length === 0) {
    processPacking();
  }

  const data = [
    ['받는분', '주소', '상세주소', '우편번호', '전화번호', '품목명', '수량', '포장상태', '가로', '세로', '높이', '고객사주문번호']
  ];

  packedOrders.forEach(box => {
    // 주소 분리 (시/도 구/군 까지는 기본주소, 나머지는 상세주소)
    const addressParts = splitAddress(box.address);

    // 포장상태
    const packagingStatus = box.packagingType === 'vinyl' ? '비닐' : '박스';

    // 치수 (임시 기본값)
    const dimensions = getDimensions(box);

    data.push([
      box.recipientLabel,
      addressParts.base,
      addressParts.detail,
      box.zipCode,
      box.phone,
      box.designText,
      1,
      packagingStatus,
      dimensions.width,
      dimensions.height,
      dimensions.depth,
      box.group.orderId
    ]);
  });

  downloadExcel(data, '경동발주서');
}

// 로젠택배 발주서 생성
function exportToLozen() {
  if (packedOrders.length === 0) {
    processPacking();
  }

  const data = [
    ['받는분', '받는분 전화번호', '받는분 핸드폰', '우편번호', '받는분 주소', '운임타입', '송장수량', '디자인+수량', '배송메세지', '합배송여부']
  ];

  packedOrders.forEach(box => {
    // 디자인+수량 형식 생성
    let designWithQty = box.designText;
    if (box.items && box.items.length > 0) {
      designWithQty = box.items.map(item => {
        if (PUZZLE_PRODUCT_IDS.includes(item.productId) || item.productCategory === '퍼즐') {
          return generatePuzzleCodeWithQty(item);
        } else if (ROLL_PRODUCT_IDS.includes(item.productId) || item.productCategory === '롤') {
          return generateRollMatCodeWithLength(item);
        } else {
          return `${generateDesignCode(item)}x${item.quantity || 1}`;
        }
      }).join('+');
    }

    data.push([
      box.recipientLabel,
      box.phone,
      box.phone,
      box.zipCode,
      box.address,
      '선불',
      1,
      designWithQty,
      box.deliveryMemo || '',
      box.isCombined ? 'Y' : 'N'
    ]);
  });

  downloadExcel(data, '로젠발주서');
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
  const ws = XLSX.utils.aoa_to_sheet(data);
  const wb = XLSX.utils.book_new();
  XLSX.utils.book_append_sheet(wb, ws, 'Sheet1');

  const today = new Date();
  const dateStr = `${today.getFullYear()}${String(today.getMonth() + 1).padStart(2, '0')}${String(today.getDate()).padStart(2, '0')}`;

  XLSX.writeFile(wb, `${filename}_${dateStr}.xlsx`);
}

// ========== 렌더링 ==========
function renderOrders() {
  renderPackingTable();
  renderCustomerCards();
  renderTableView();
}

// 패킹 결과 테이블 렌더링
function renderPackingTable() {
  const tbody = document.getElementById('packingTableBody');
  if (!tbody) return;

  tbody.innerHTML = '';

  // 빈 박스 필터링 (items가 없거나, designText가 없거나, items의 총 수량이 0인 경우)
  const validBoxes = packedOrders.filter(box => {
    if (!box.items || box.items.length === 0) return false;
    if (!box.designText || box.designText.trim() === '') return false;
    const totalQty = box.items.reduce((sum, item) => sum + (item.quantity || 0), 0);
    return totalQty > 0;
  });

  validBoxes.forEach((box, index) => {
    const packagingText = box.packagingType === 'vinyl' ? '비닐' :
      box.packagingType === 'largeBox' ? '대박스' : '박스';

    // 비고 내용 조합
    let remarks = [];
    if (box.remark) remarks.push(box.remark);
    if (box.hasCuttingRequest) remarks.push('재단');
    if (box.hasFinishingRequest) remarks.push('마감재');

    // 사은품
    let giftText = '';
    if (box.giftEligible) {
      giftText = box.giftEligible;
    }

    const tr = document.createElement('tr');
    tr.className = box.needsStar ? 'star-row' : '';
    tr.innerHTML = `
      <td class="recipient-cell">
        <span class="recipient-label">${box.recipientLabel}</span>
      </td>
      <td class="design-cell">${box.designText}</td>
      <td class="packaging-cell">
        <span class="pkg-badge ${box.packagingType === 'vinyl' ? 'vinyl' : 'box'}">${packagingText}</span>
      </td>
      <td class="remark-cell">${remarks.join(', ')}</td>
      <td class="gift-cell">${giftText ? `<span class="gift-tag">${giftText}</span>` : ''}</td>
    `;

    // 클릭 시 상세 보기
    tr.addEventListener('click', () => showPackingDetail(box));
    tbody.appendChild(tr);
  });
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
      ${box.deliveryMemo ? `
        <div class="detail-group">
          <label>배송메모</label>
          <p style="color: #c0392b;">${box.deliveryMemo}</p>
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

// 고객 카드 뷰 렌더링
function renderCustomerCards() {
  customerCardsContainer.innerHTML = '';

  Object.values(groupedOrders).forEach(customer => {
    const card = document.createElement('div');
    card.className = 'customer-card';
    card.dataset.key = customer.key;

    // 패킹 요약 계산
    let packStats = {
      smallBox: 0,
      largeBox: 0,
      vinyl: 0,
      total: 0
    };

    // packedOrders에서 해당 고객의 박스 정보 집계 (이미 processPacking()이 실행된 상태여야 함)
    // 현재 구조상 packedOrders는 전역변수이고, customerName 등으로 필터링해야 함
    // 하지만 renderCustomerCards() 시점에는 packedOrders가 이미 채워져 있음
    const customerBoxes = packedOrders.filter(b => b.group.key === customer.key);

    customerBoxes.forEach(box => {
      packStats.total++;
      if (box.packagingType === 'vinyl') packStats.vinyl++;
      else if (box.packagingType === 'largeBox') packStats.largeBox++;
      else packStats.smallBox++;
    });

    // 택배사 결정 로직 (임시: 대박스나 비닐이 있으면 경동, 아니면 로젠)
    let courierName = '로젠택배';
    let courierClass = 'lozen';
    if (packStats.largeBox > 0 || packStats.vinyl > 0) {
      courierName = '경동택배';
      courierClass = 'kyungdong';
    }

    // 패킹 요약 텍스트 생성
    let packSummary = [];
    if (packStats.smallBox > 0) packSummary.push(`소박스 ${packStats.smallBox}`);
    if (packStats.largeBox > 0) packSummary.push(`대박스 ${packStats.largeBox}`);
    if (packStats.vinyl > 0) packSummary.push(`비닐 ${packStats.vinyl}`);
    let packSummaryText = packSummary.join(' / ');
    if (packSummary.length === 0 && customerBoxes.length === 0) packSummaryText = '패킹 대기';


    // 헤더
    const header = document.createElement('div');
    header.className = 'customer-card-header';
    header.innerHTML = `
      <div class="customer-info">
        <div class="customer-avatar">${customer.customerName.charAt(0)}</div>
        <div class="customer-details">
          <h3>
            ${customer.customerName}
            <span class="order-count">${customer.items.length}건</span>
            <span class="courier-badge ${courierClass}">${courierName}</span>
          </h3>
          <div class="customer-address">${customer.address}</div>
        </div>
      </div>
      <div class="customer-meta">
        <div class="customer-meta-item">
          <span class="meta-value packaging-total">${packSummaryText}</span>
          <span class="meta-label">포장 내역</span>
        </div>
        ${customer.giftEligible ? `
          <div class="customer-meta-item gift">
            <span class="meta-value">${customer.giftEligible}</span>
            <span class="meta-label">사은품</span>
          </div>
        ` : ''}
        <div class="customer-meta-item">
          <span class="meta-value">${formatCurrency(customer.totalPrice)}</span>
          <span class="meta-label">총 금액</span>
        </div>
      </div>
    `;

    // 바디 (주문 아이템 목록)
    const body = document.createElement('div');
    body.className = 'customer-card-body';

    const itemsList = document.createElement('div');
    itemsList.className = 'order-items-list';

    customer.items.forEach(item => {
      const packaging = determineRollPackaging(item);
      const itemEl = document.createElement('div');
      itemEl.className = 'order-item';
      itemEl.innerHTML = `
        <div class="order-item-icon">${getProductIcon(item.productType)}</div>
        <div class="order-item-details">
          <div class="order-item-name">
            <span class="product-badge ${getProductBadgeClass(item.productType)}">${item.productType}</span>
            ${item.design}
            ${packaging.needsStar ? '<span class="star-mark">★</span>' : ''}
          </div>
          <div class="order-item-specs">
            두께 <strong>${item.thickness}</strong> / 폭 <strong>${item.width}</strong> / 길이 <strong>${item.length}</strong>
          </div>
        </div>
        <div class="order-item-qty">×${item.quantity}</div>
        <div class="order-item-price">${formatCurrency(item.price)}</div>
      `;
      itemsList.appendChild(itemEl);
    });

    body.appendChild(itemsList);

    // 푸터
    const footer = document.createElement('div');
    footer.className = 'customer-card-footer';
    footer.innerHTML = `
      <div class="footer-info">
        ${customer.hasCuttingRequest ? `
          <div class="footer-info-item request-tag cutting">재단요청</div>
        ` : ''}
        ${customer.hasFinishingRequest ? `
          <div class="footer-info-item request-tag finishing">마감재요청</div>
        ` : ''}
        ${customer.deliveryMemo ? `
          <div class="footer-info-item">
            <span class="delivery-memo-badge">${customer.deliveryMemo}</span>
          </div>
        ` : ''}
      </div>
      <div class="footer-actions">
        <button class="footer-btn detail-btn" data-key="${customer.key}">상세보기</button>
      </div>
    `;

    // 상세보기 버튼 이벤트
    footer.querySelector('.detail-btn').addEventListener('click', () => {
      showCustomerDetail(customer);
    });

    card.appendChild(header);
    card.appendChild(body);
    card.appendChild(footer);
    customerCardsContainer.appendChild(card);
  });
}

function getProductIcon(type) {
  return '';
}

// 테이블 뷰 렌더링
function renderTableView() {
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
}

// 고객 상세 모달 (카드뷰용)
function showCustomerDetail(customer) {
  let itemsHtml = customer.items.map(item => {
    const packaging = determineRollPackaging(item);
    const packagingText = packaging.type === 'vinyl' ? '비닐' :
      packaging.type === 'largeBox' ? '대박스' :
        packaging.type === 'combine' ? '합포장' : '소박스';

    // 원본 데이터 표시용
    const rawData = item.rawRow || {};
    const rawInfo = rawData['옵션정보'] || rawData['상품명'] || '';

    return `
    <div style="padding: 0.75rem; margin-bottom: 0.5rem; background: #f8fafc; border-radius: 8px; border: 1px solid #e2e8f0;">
      <span class="product-badge ${getProductBadgeClass(item.productType)}">${item.productType}</span>
      ${item.design} ${packaging.needsStar ? '★' : ''}<br>
      <small style="color: #64748b;">두께: ${item.thickness} / 폭: ${item.width} / 길이: ${item.length} / 수량: ${item.quantity}</small><br>
      <small style="color: #888;">포장: ${packagingText}</small><br>
      <strong>${formatCurrency(item.price)}</strong>
      ${rawInfo ? `<div style="margin-top: 0.5rem; font-size: 0.7rem; color: #999; font-style: italic; border-top: 1px solid #e2e8f0; padding-top: 0.5rem;">원본: ${rawInfo}</div>` : ''}
    </div>
  `;
  }).join('');

  modalBody.innerHTML = `
    <div class="order-detail">
      <div class="detail-group">
        <label>주문번호</label>
        <p>${customer.orderId}</p>
      </div>
      <div class="detail-group">
        <label>고객명 / 연락처</label>
        <p>${customer.customerName} / ${customer.phone}</p>
      </div>
      <div class="detail-group">
        <label>배송지</label>
        <p>${customer.address}</p>
      </div>
      <hr>
      <div class="detail-group">
        <label>주문 상품 (${customer.items.length}건)</label>
        ${itemsHtml}
      </div>
      <hr>
      <div class="detail-group">
        <label>총 금액</label>
        <p><strong style="font-size: 1.2rem;">${formatCurrency(customer.totalPrice)}</strong></p>
      </div>
      ${customer.giftEligible ? `
        <div class="detail-group">
          <label>사은품 대상</label>
          <p style="color: #e74c3c; font-weight: bold;">${customer.giftEligible}</p>
        </div>
      ` : ''}
      ${customer.deliveryMemo ? `
        <div class="detail-group">
          <label>배송메모</label>
          <p>${customer.deliveryMemo}</p>
        </div>
      ` : ''}
      ${customer.hasCuttingRequest ? `
        <div class="detail-group">
          <label style="color: #e74c3c;">재단요청</label>
          <p>배송메모에 재단 관련 키워드가 포함되어 있습니다.</p>
        </div>
      ` : ''}
    </div>
  `;

  orderModal.style.display = 'flex';
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

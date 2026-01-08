// ========== 유틸리티 함수 ==========

/**
 * 금액을 한국 원화 형식으로 포맷팅
 * @param {number} amount - 금액
 * @returns {string} 포맷팅된 금액 문자열
 */
export function formatCurrency(amount) {
    return new Intl.NumberFormat('ko-KR', {
        style: 'currency',
        currency: 'KRW',
        maximumFractionDigits: 0
    }).format(amount);
}

/**
 * 주소를 기본주소와 상세주소로 분리
 * @param {string} address - 전체 주소
 * @returns {{ base: string, detail: string }}
 */
export function splitAddress(address) {
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

/**
 * 전화번호 정규화 (숫자만 추출, 국제번호 처리)
 * @param {string} phone - 전화번호
 * @returns {string} 정규화된 전화번호
 */
export function normalizePhone(phone) {
    const digits = String(phone || '').replace(/\D/g, '');
    if (!digits) return '';
    // 국내번호 형태를 최대한 유지: +82/82로 시작하면 0을 붙여 정규화
    if (digits.startsWith('82') && digits.length >= 10) {
        return `0${digits.slice(2)}`;
    }
    return digits;
}

/**
 * 주소 정규화 (특수문자 제거, 공백 정리)
 * @param {string} address - 주소
 * @returns {string} 정규화된 주소
 */
export function normalizeAddress(address) {
    return String(address || '')
        .replace(/\\/g, '') // CSV/엑셀에서 들어온 '\' 제거
        .replace(/\(\d{5}\)/g, '') // (우편번호) 제거
        .replace(/\s+/g, ' ')
        .trim();
}

/**
 * 고객 키 생성 (주소 + 전화번호 기반)
 * @param {{ address: string, phone: string }} customer
 * @returns {string} 고객 키
 */
export function makeCustomerKey({ address, phone }) {
    const normalizedAddress = normalizeAddress(address);
    const normalizedPhone = normalizePhone(phone);
    if (!normalizedAddress && !normalizedPhone) return '';
    return `${normalizedAddress}|${normalizedPhone}`;
}

/**
 * 박스 치수 계산
 * @param {object} box - 박스 정보
 * @returns {{ width: number, height: number, depth: number }}
 */
export function getDimensions(box) {
    // 실제로는 제품별로 치수가 다르지만, 기본값 반환
    if (box.type === 'puzzle') {
        return { width: 100, height: 100, depth: 30 };
    }
    if (box.packagingType === 'vinyl') {
        return { width: 150, height: 20, depth: 20 };
    }
    return { width: 120, height: 15, depth: 15 };
}

/**
 * 길이 변환 (cm → m)
 * @param {number} lengthCm - cm 단위 길이
 * @returns {string} m 단위 문자열
 */
export function formatLengthToMeters(lengthCm) {
    if (!lengthCm || lengthCm === 0) return '';
    const meters = lengthCm / 100;
    // 소수점 처리
    if (meters === Math.floor(meters)) {
        return `${meters}m`;
    }
    return `${meters.toFixed(1)}m`;
}

/**
 * 제품 타입에 따른 뱃지 클래스 반환
 * @param {string} type - 제품 타입
 * @returns {string} CSS 클래스명
 */
export function getProductBadgeClass(type) {
    if (type === '유아롤매트') return 'roll-baby';
    if (type === '애견롤매트') return 'roll-pet';
    if (type === '퍼즐매트') return 'puzzle';
    return '';
}

/**
 * 주문 상태에 따른 클래스 반환
 * @param {string} status - 주문 상태
 * @returns {string} CSS 클래스명
 */
export function getStatusClass(status) {
    if (status === '발송대기') return 'pending';
    if (status === '배송중') return 'shipping';
    if (status === '배송완료') return 'completed';
    return 'pending';
}

/**
 * Excel 파일 다운로드 (XLSX 라이브러리 사용)
 * @param {Array} data - 2차원 배열 데이터
 * @param {string} filename - 파일명 (확장자 제외)
 */
export function downloadExcel(data, filename) {
    // 메타데이터 제거 (고객명, 박스 인덱스 등)
    const cleanData = data.map((row, rowIndex) => {
        if (rowIndex === 0) return row; // 헤더는 그대로
        if (Array.isArray(row)) {
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

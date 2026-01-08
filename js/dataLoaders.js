// ========== 데이터 로더 모듈 ==========
import { parseCSVToRows } from './parsers.js';

/**
 * 문자열 정규화 - 이모지를 언더스코어로 치환하고 비교용으로 정리
 * @param {string} str - 원본 문자열
 * @returns {string} 정규화된 문자열
 */
function normalizeForMatching(str) {
    if (!str || typeof str !== 'string') return str;

    return str
        // 이모지를 언더스코어로 치환 (🏆BEST🏆 -> __BEST__)
        .replace(/[\u{1F300}-\u{1F9FF}]/gu, '_')  // Miscellaneous Symbols and Pictographs, Emoticons, etc.
        .replace(/[\u{2600}-\u{26FF}]/gu, '_')    // Miscellaneous Symbols
        .replace(/[\u{2700}-\u{27BF}]/gu, '_')    // Dingbats
        .replace(/[\u{FE00}-\u{FE0F}]/gu, '')     // Variation Selectors (제거)
        .replace(/[\u{1F000}-\u{1F02F}]/gu, '_')  // Mahjong Tiles
        .replace(/[\u{1F0A0}-\u{1F0FF}]/gu, '_')  // Playing Cards
        .replace(/[\u{200D}]/gu, '')              // Zero Width Joiner (제거)
        .replace(/[\u{20E3}]/gu, '')              // Combining Enclosing Keycap (제거)
        .replace(/[\u{E0020}-\u{E007F}]/gu, '')   // Tags (제거)
        // 연속된 언더스코어를 두 개로 통일
        .replace(/_+/g, '__')
        .trim();
}

/**
 * 배송비 CSV 데이터 로드
 * @returns {Promise<Array>} 배송비 데이터 배열
 */
export async function loadShippingFees() {
    const startTime = Date.now();
    console.log('=== 배송비 CSV 로드 시작 ===', new Date().toISOString());

    try {
        const response = await fetch('shipping_fees.csv');

        if (!response.ok) {
            console.error('❌ 배송비 데이터 파일을 찾을 수 없습니다!', {
                status: response.status,
                statusText: response.statusText
            });
            return [];
        }

        const text = await response.text();
        const rows = parseCSVToRows(text);

        if (rows.length < 2) {
            console.error('❌ 배송비 데이터가 없습니다! 행 수:', rows.length);
            return [];
        }

        const headers = rows[0].map(h => h.trim());
        const shippingFees = [];

        for (let i = 1; i < rows.length; i++) {
            const rowValues = rows[i];
            const rowObj = {};
            headers.forEach((header, index) => {
                rowObj[header] = rowValues[index] || '';
            });

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
            }
        }

        const loadTime = Date.now() - startTime;
        console.log(`✅ 배송비 데이터 로드 완료! (${loadTime}ms), 총 ${shippingFees.length}건`);

        return shippingFees;
    } catch (error) {
        console.error('❌ 배송비 데이터 로드 오류:', error);
        return [];
    }
}

/**
 * 제품 DB CSV 데이터 로드
 * @returns {Promise<Array>} 제품 DB 배열
 */
export async function loadProductDb() {
    const startTime = Date.now();
    console.log('=== 제품 DB CSV 로드 시작 ===', new Date().toISOString());

    try {
        const response = await fetch('product_db.csv');

        if (!response.ok) {
            console.error('❌ 제품 DB 파일을 찾을 수 없습니다!', {
                status: response.status,
                statusText: response.statusText
            });
            return [];
        }

        const text = await response.text();
        const rows = parseCSVToRows(text);

        if (rows.length < 2) {
            console.error('❌ 제품 DB 데이터가 없습니다! 행 수:', rows.length);
            return [];
        }

        const productDb = [];

        for (let i = 1; i < rows.length; i++) {
            const rowValues = rows[i];
            const productId = (rowValues[0] || '').trim();
            const optionInfo = (rowValues[1] || '').trim();
            const designCode = (rowValues[2] || '').trim();
            const lengthNum = (rowValues[3] || '').trim();
            const lengthM = (rowValues[4] || '').trim();

            if (productId && optionInfo && designCode) {
                productDb.push({
                    productId,
                    optionInfo,
                    designCode,
                    lengthNum,
                    lengthM
                });
            }
        }

        const loadTime = Date.now() - startTime;
        console.log(`✅ 제품 DB 데이터 로드 완료! (${loadTime}ms), 총 ${productDb.length}건`);

        return productDb;
    } catch (error) {
        console.error('❌ 제품 DB 데이터 로드 오류:', error);
        return [];
    }
}

/**
 * 배송비 계산 함수
 * @param {Object} box - 박스 정보
 * @param {Array} shippingFees - 배송비 데이터 배열
 * @returns {number} 배송비
 */
export function calculateShippingFee(box, shippingFees) {
    if (!box || !box.items || box.items.length === 0 || !shippingFees || shippingFees.length === 0) {
        return 0;
    }

    const item = box.items[0];
    const categoryCode = item.productCategoryCode || '';

    // 합포장된 박스의 총 길이 계산
    const totalLength = box.items.reduce((sum, i) => sum + ((i.lengthM || 0) * (i.quantity || 1)), 0);
    const widthValues = box.items.map(i => i.widthNum || 0).filter(w => w > 0);
    const maxWidth = widthValues.length > 0 ? Math.max(...widthValues) : 0;
    const thicknessValues = box.items.map(i => i.thicknessNum || 0).filter(t => t > 0);
    const maxThickness = thicknessValues.length > 0 ? Math.max(...thicknessValues) : 0;
    const totalPuzzleCount = box.items.reduce((sum, i) => sum + (i.quantity || 1), 0);

    // 제품군 매핑
    let productGroup = '';
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
    } else {
        return 0;
    }

    let packageType = '';
    if (productGroup === '퍼즐매트') {
        packageType = '강화비닐(100x100cm)';
    }

    const width = maxWidth > 0 ? maxWidth : null;
    const thickness = maxThickness > 0 ? maxThickness / 10 : null;
    const length = productGroup === '퍼즐매트' ? totalPuzzleCount : totalLength;

    let matchedFee = null;
    let largeBoxFee = null;

    for (const fee of shippingFees) {
        if (fee.productGroup !== productGroup) continue;

        let packageMatch = packageType === '' || fee.packageType === packageType ||
            (fee.packageType.includes('강화비닐') && packageType.includes('강화비닐'));

        let widthMatch = true;
        if (width !== null && fee.width !== null) {
            if (Math.abs(fee.width - width) > 5) widthMatch = false;
        }

        let thicknessMatch = true;
        if (thickness !== null && fee.thickness !== null) {
            if (Math.abs(fee.thickness - thickness) > 0.1) thicknessMatch = false;
        }

        let lengthMatch = true;
        if (length > 0 && fee.lengthMin !== null && fee.lengthMax !== null) {
            if (length < fee.lengthMin || length > fee.lengthMax) {
                lengthMatch = false;
                if (fee.packageType.includes('대박스') && widthMatch && thicknessMatch && length > fee.lengthMax) {
                    largeBoxFee = fee;
                }
            }
        }

        if (packageMatch && widthMatch && thicknessMatch && lengthMatch) {
            matchedFee = fee;
            break;
        }
    }

    if (!matchedFee && largeBoxFee) {
        return largeBoxFee.fee + 5000; // 비닐 처리
    }

    return matchedFee ? matchedFee.fee : 0;
}

/**
 * 제품 DB에서 디자인 코드 매칭
 * @param {string} productId - 상품번호
 * @param {string} optionInfo - 옵션정보
 * @param {Array} productDb - 제품 DB 배열
 * @returns {Object|null} 매칭된 엔트리 또는 null
 */
export function matchDesignCodeFromDb(productId, optionInfo, productDb) {
    if (!productId || !optionInfo || !productDb || productDb.length === 0) {
        return null;
    }

    const productIdStr = String(productId).trim();
    const candidates = productDb.filter(entry => String(entry.productId).trim() === productIdStr);
    if (candidates.length === 0) return null;

    const optionInfoTrimmed = String(optionInfo).trim();
    // 이모지를 언더스코어로 정규화한 버전 (🏆BEST🏆 -> __BEST__)
    const optionInfoNormalized = normalizeForMatching(optionInfoTrimmed);

    for (const entry of candidates) {
        const dbOptionTrimmed = String(entry.optionInfo).trim();

        // 1차: 정확히 일치
        if (dbOptionTrimmed === optionInfoTrimmed) {
            return entry;
        }

        // 2차: 정규화 후 비교 (이모지 -> 언더스코어)
        const dbOptionNormalized = normalizeForMatching(dbOptionTrimmed);
        if (dbOptionNormalized === optionInfoNormalized) {
            return entry;
        }
    }

    return null;
}

/**
 * 제품 DB 매칭 결과 전체 반환
 * @param {Object} order - 주문 객체
 * @param {Array} productDb - 제품 DB 배열
 * @returns {Object|null}
 */
export function getProductDbMatch(order, productDb) {
    if (!order || !order.productId) return null;
    const optionInfo = order.optionInfo || '';
    if (!optionInfo) return null;
    return matchDesignCodeFromDb(order.productId, optionInfo, productDb);
}

/**
 * 송장용 디자인 코드 생성
 * @param {Object} order - 주문 객체
 * @param {Array} productDb - 제품 DB 배열
 * @returns {string} 디자인 코드
 */
export function generateDesignCode(order, productDb) {
    const match = getProductDbMatch(order, productDb);
    return match ? match.designCode : '[미등록]';
}

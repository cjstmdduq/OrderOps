// ========== 파싱 함수 모듈 ==========
import { PRODUCT_TYPE_MAP, CUTTING_KEYWORDS, FINISHING_KEYWORDS } from './constants.js';

/**
 * CSV 텍스트를 rows 배열로 파싱 (줄바꿈/콤마가 포함된 큰따옴표 필드 대응)
 * @param {string} text - CSV 텍스트
 * @returns {Array<Array<string>>} 파싱된 행 배열
 */
export function parseCSVToRows(text) {
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
 * 우편번호 추출
 * @param {string} address - 주소 문자열
 * @returns {string} 우편번호
 */
export function extractZipCode(address) {
    const match = address.match(/\((\d{5})\)/);
    return match ? match[1] : '';
}

/**
 * 두께값 파싱 (예: "17T" → 17, "1.7cm" → 17)
 * @param {string} thicknessStr - 두께 문자열
 * @returns {number} 두께 숫자
 */
export function parseThickness(thicknessStr) {
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

/**
 * 폭값 파싱 (예: "140cm" → 140)
 * @param {string} widthStr - 폭 문자열
 * @returns {number} 폭 숫자
 */
export function parseWidth(widthStr) {
    if (!widthStr) return 0;
    const match = widthStr.match(/(\d+)/);
    return match ? parseInt(match[1]) : 0;
}

/**
 * 길이값을 m 단위로 변환 (예: "800cm" → 8, "8m" → 8)
 * @param {string} lengthStr - 길이 문자열
 * @returns {number} m 단위 숫자
 */
export function parseLengthToMeters(lengthStr) {
    if (!lengthStr) return 0;
    // m 단위
    let match = lengthStr.match(/(\d+\.?\d*)m(?!m)/i); // mm가 아닌 m만 찾음
    if (match) return parseFloat(match[1]);

    // cm 단위 (명시적)
    match = lengthStr.match(/(\d+)cm/i);
    if (match) return parseInt(match[1]) / 100;

    // 단위 없는 숫자
    match = lengthStr.match(/(\d+(\.\d+)?)/);
    if (match) {
        const val = parseFloat(match[1]);
        if (val >= 30) return val / 100; // 30cm 이상은 cm로 간주
        return val; // 그 외는 m
    }
    return 0;
}

/**
 * 금액 문자열을 숫자로 변환
 * @param {string} priceStr - "₩125,300", "23000" 등
 * @returns {number}
 */
export function parsePrice(priceStr) {
    if (!priceStr) return 0;
    const cleaned = String(priceStr).replace(/[^0-9.]/g, '');
    return parseInt(cleaned) || 0;
}

/**
 * 재단 요청 감지
 * @param {string} text - 배송메세지 등
 * @returns {boolean}
 */
export function detectCuttingRequest(text) {
    if (!text) return false;
    const lowerText = text.toLowerCase();
    const hasNumber = /\d/.test(text);
    const hasKeyword = CUTTING_KEYWORDS.some(keyword => lowerText.includes(keyword.toLowerCase()));
    return hasKeyword || hasNumber;
}

/**
 * 마감재 요청 감지
 * @param {string} text - 배송메세지 등
 * @returns {boolean}
 */
export function detectFinishingRequest(text) {
    if (!text) return false;
    return FINISHING_KEYWORDS.some(keyword => text.includes(keyword));
}

/**
 * 옵션정보에서 디자인/두께/폭/길이 추출
 * @param {string} optionInfo - 옵션정보 문자열
 * @returns {{ design: string, thickness: string, width: string, length: string }}
 */
export function parseOptionInfo(optionInfo) {
    const result = { design: '', thickness: '', width: '', length: '' };

    // 디자인 추출
    let designMatch = optionInfo.match(/디자인선택:\s*([^/]+)/);
    if (!designMatch) {
        designMatch = optionInfo.match(/디자인:\s*([^/]+)/);
    }
    if (designMatch) result.design = designMatch[1].trim();

    // 두께 추출
    let thicknessMatch = optionInfo.match(/두께선택:\s*([0-9.]+)T/i);
    if (!thicknessMatch) {
        thicknessMatch = optionInfo.match(/두께\s*\(?폭\)?:\s*([0-9.]+)(mm|cm|T)/i);
    }
    if (!thicknessMatch) {
        thicknessMatch = optionInfo.match(/두께\s*\/?\s*폭:\s*([0-9.]+)(cm|mm|T)/i);
    }
    // 퍼즐 옵션처럼 "(4.0cm) 100x100" 형식인 경우도 두께로 인식
    if (!thicknessMatch) {
        thicknessMatch = optionInfo.match(/\(([0-9.]+)\s*cm\)/i);
        if (thicknessMatch) {
            // 이 케이스는 항상 cm 단위로 취급
            result.thickness = thicknessMatch[1] + 'cm';
        }
    }
    if (thicknessMatch && !result.thickness) {
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

    // 폭 추출
    const thicknessWidthMatch = optionInfo.match(/두께\s*\/?\s*폭:\s*([0-9.]+)cm\/([0-9]+)cm/);
    if (thicknessWidthMatch) {
        result.width = thicknessWidthMatch[2] + 'cm';
    } else {
        const widthInThicknessMatch = optionInfo.match(/\(폭([0-9]+)cm\)/);
        if (widthInThicknessMatch) {
            result.width = widthInThicknessMatch[1] + 'cm';
        }
    }

    // 길이 추출 - 순차 패턴 매칭
    let lengthMatch = optionInfo.match(/길이선택:\s*(\d+)x(\d+)/);
    if (lengthMatch) {
        const widthNum = parseInt(lengthMatch[1]);
        const lengthNum = parseInt(lengthMatch[2]);
        if (!result.width) {
            result.width = widthNum + 'cm';
        }
        if (lengthNum >= 100) {
            result.length = lengthNum + 'cm';
        }
    }

    if (!result.length) {
        lengthMatch = optionInfo.match(/길이:\s*([0-9.]+)\s*(cm|m|미터|메터)/i);
        if (lengthMatch) {
            const value = parseFloat(lengthMatch[1]);
            const unit = lengthMatch[2] ? lengthMatch[2].toLowerCase() : 'cm';
            if (unit === 'm' || unit === '미터' || unit === '메터') {
                result.length = value + 'm';
            } else {
                result.length = value + 'cm';
            }
        }
    }

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

    if (!result.length) {
        lengthMatch = optionInfo.match(/([0-9.]+)\s*(m|미터|메터)(\s*롤)?/i);
        if (lengthMatch) {
            result.length = parseFloat(lengthMatch[1]) + 'm';
        }
    }

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

/**
 * 상품명에서 디자인/두께/폭/길이 추출 (예: "마블아이보리 6T 110x50cm")
 * @param {string} productName - 상품명
 * @returns {{ design: string, thickness: string, width: string, length: string }}
 */
export function parseProductName(productName) {
    const result = { design: '', thickness: '', width: '', length: '' };
    if (!productName) return result;

    // 알려진 디자인명 패턴
    const designPatterns = [
        '마블아이보리', '퓨어아이보리', '그레이캔버스', '딜라이트우드', '모던그레이',
        '스노우화이트', '내추럴우드', '바닐라아이보리', '라이트그레이', '스카이블루',
        '파스텔핑크', '무직타이거', '코지베어', '플라워가든', '스타라이트',
        '헬로베어', '트로피칼', '사파리', '포레스트', '클라우드'
    ];

    for (const design of designPatterns) {
        if (productName.includes(design)) {
            result.design = design;
            break;
        }
    }

    // 두께 추출
    let thicknessMatch = productName.match(/(\d+)T\b/i);
    if (thicknessMatch) {
        result.thickness = thicknessMatch[1] + 'T';
    } else {
        thicknessMatch = productName.match(/x(\d+\.\d+)/);
        if (thicknessMatch) {
            result.thickness = Math.round(parseFloat(thicknessMatch[1]) * 10) + 'mm';
        }
    }

    // 폭x길이 추출
    const sizeMatch = productName.match(/(\d+)x(\d+)cm/i);
    if (sizeMatch) {
        const widthNum = parseInt(sizeMatch[1]);
        const lengthNum = parseInt(sizeMatch[2]);
        result.width = sizeMatch[1] + 'cm';
        if (lengthNum >= 100 || lengthNum >= widthNum) {
            result.length = sizeMatch[2] + 'cm';
        }
    }

    // m 단위 길이 직접 추출
    if (!result.length) {
        const lengthMMatch = productName.match(/([0-9.]+)\s*(m|미터|메터)(\s*롤)?/i);
        if (lengthMMatch) {
            result.length = lengthMMatch[1] + 'm';
        }
    }

    return result;
}

/**
 * 상품명 기반 제품타입 판별 (폴백)
 * @param {string} productName - 상품명
 * @returns {string} 제품 타입
 */
export function getProductTypeByName(productName) {
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
 * 주문 데이터 행을 파싱하여 주문 객체 생성
 * @param {object} row - CSV 행 객체
 * @param {Array} productDb - 제품 DB (디자인 코드 매칭용)
 * @returns {object|null} 파싱된 주문 객체 또는 null
 */
export function parseOrderData(row, productDb = []) {
    try {
        if (!row['상품주문번호'] || !row['상품명']) {
            return null;
        }

        const optionInfo = row['옵션정보'] || '';
        const productName = row['상품명'] || '';
        const parsed = parseOptionInfo(optionInfo);
        const parsedFromName = parseProductName(productName);
        const productId = row['상품번호'] || '';
        const deliveryMemo = row['배송메세지'] || '';

        const design = parsed.design || parsedFromName.design || '';
        const thickness = parsed.thickness || parsedFromName.thickness || '';
        const width = parsed.width || parsedFromName.width || '';
        let length = parsed.length || parsedFromName.length || '';

        const productMapping = PRODUCT_TYPE_MAP[productId];
        const productType = productMapping ? productMapping.name : getProductTypeByName(productName);
        const productCategory = productMapping ? productMapping.type : '기타';
        const productCategoryCode = productMapping ? productMapping.category : 'etc';

        const hasCuttingRequest = productId === '4200445704' && detectCuttingRequest(deliveryMemo);
        const hasFinishingRequest = (productId === '5994906898' || productId === '5994903887') && detectFinishingRequest(deliveryMemo);

        if (productCategory === 'petRoll' && !length) {
            length = '50cm';
        }

        // 4200445704 상품: 수량 × 단위길이(0.5m)로 실제 길이 계산
        // 주의: 이 로직은 애견롤매트(4200445704)에만 적용되어야 함
        // 유아롤매트(6092903705)는 수량과 길이를 분리해서 처리해야 함
        // 애견롤매트는 1개당 50cm(0.5m) 고정이므로, 수량 × 0.5m = 실제 길이
        let quantity = parseInt(row['수량']) || 1;
        let lengthM = parseLengthToMeters(length);

        if (productId === '4200445704' && lengthM > 0) {
            // 애견롤매트는 항상 0.5m 단위이므로, 수량과 관계없이 계산
            // 예: 수량 10개 × 0.5m = 5m 한 롤
            lengthM = lengthM * quantity;
            // 길이 문자열도 업데이트 (예: 0.5m × 10 = 5m)
            length = lengthM + 'm';
            // 수량은 1롤로 처리 (길이에 이미 합산됨)
            quantity = 1;
        }
        const thicknessNum = parseThickness(thickness);
        const widthNum = parseWidth(width);

        return {
            id: row['상품주문번호'] || '',
            orderId: row['주문번호'] || '',
            productId: productId,
            customerName: row['수취인명'] || row['구매자명'] || '',
            productName: productName,
            optionInfo: optionInfo,
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
            quantity: quantity,
            price: parsePrice(row['최종 상품별 총 주문금액']),
            status: row['주문상태'] || '',
            gift: row['사은품'] || '',
            deliveryMemo: deliveryMemo,
            address: row['통합배송지'] || row['배송지'] || '',
            zipCode: row['우편번호'] || extractZipCode(row['통합배송지'] || ''),
            phone: row['수취인연락처1'] || '',
            orderDate: row['주문일시'] || '',
            hasCuttingRequest: hasCuttingRequest,
            hasFinishingRequest: hasFinishingRequest,
            rawRow: row
        };
    } catch (e) {
        console.error('파싱 오류:', e, row);
        return null;
    }
}

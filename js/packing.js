// ========== 패킹 알고리즘 모듈 ==========
import {
    ROLL_PRODUCT_IDS,
    PUZZLE_PRODUCT_IDS,
    PACKAGING_THRESHOLDS,
    PUZZLE_BOX_CAPACITY
} from './constants.js';

/**
 * 동일 수취인 그룹화 (이름 + 주소 + 연락처)
 * @param {Array} orders - 주문 배열
 * @returns {Object} 그룹화된 주문 객체
 */
export function groupOrdersByRecipient(orders) {
    const groupedOrders = {};

    orders.forEach(order => {
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
                rollItems: [],
                puzzleItems: [],
                tapeItems: [],
                otherItems: [],
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
        const rollMatOrders = group.items.filter(item =>
            item.productId === '6092903705' || item.productId === '4200445704'
        );
        const rollMatTotalPrice = rollMatOrders.reduce((sum, item) => sum + item.price, 0);

        if (rollMatTotalPrice >= 500000) {
            group.giftEligible = '실리콘테이프 2개';
        } else if (rollMatTotalPrice >= 95000) {
            group.giftEligible = '실리콘테이프 1개';
        }
    });

    return groupedOrders;
}

/**
 * 롤매트 포장 방식 결정
 * @param {Object} item - 아이템 정보
 * @returns {{ type: string, needsStar: boolean, canCombine?: boolean }}
 */
export function determineRollPackaging(item) {
    const thickness = item.thicknessNum;
    const length = item.lengthM;
    const width = item.widthNum;

    // 70cm 폭 → 로젠 소박스 (단독 배송 전용, 합포장 불가)
    if (width === 70) {
        return { type: 'smallBox', needsStar: false, canCombine: false };
    }

    const thresholds = PACKAGING_THRESHOLDS[String(thickness)] || PACKAGING_THRESHOLDS["17"];

    // 140폭 8m 이상 → 비닐 + ★마킹
    if (width >= 140 && length >= 8) {
        return { type: 'vinyl', needsStar: true };
    }

    if (length >= thresholds.vinyl) {
        return { type: 'vinyl', needsStar: true, canCombine: false };
    } else if (length >= thresholds.large) {
        return { type: 'largeBox', needsStar: false, canCombine: false };
    } else if (length >= thresholds.small) {
        return { type: 'smallBox', needsStar: false, canCombine: false };
    } else {
        return { type: 'smallBox', needsStar: false, canCombine: true };
    }
}

/**
 * 퍼즐매트 박스 수 계산
 * @param {Array} items - 퍼즐매트 아이템 배열
 * @returns {number} 필요한 박스 수
 */
export function calculatePuzzleBoxes(items) {
    let totalCount = 0;
    let thickness = 25;

    items.forEach(item => {
        totalCount += item.quantity;
        if (item.thicknessNum === 40) {
            thickness = 40;
        }
    });

    const capacity = PUZZLE_BOX_CAPACITY[String(thickness)] || 6;
    return Math.ceil(totalCount / capacity);
}

/**
 * 배송메모에서 재단 요청 세그먼트 파싱
 * @param {string} memo - 배송메모
 * @returns {Array<{ lengthM: number, count: number }>}
 */
export function parseCuttingSegmentsFromMemo(memo) {
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

/**
 * 배송메모에서 고유 길이 요청값 추출
 * @param {string} memo - 배송메모
 * @returns {Array<number>}
 */
export function getDistinctLengthRequestsFromMemo(memo) {
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

/**
 * 아이템들에서 경고 메시지 수집
 * @param {Array} items - 아이템 배열
 * @returns {Array<string>}
 */
export function collectWarningsFromItems(items) {
    const warningsSet = new Set();
    (items || []).forEach(item => {
        (item.cuttingWarnings || []).forEach(w => warningsSet.add(w));
    });
    return Array.from(warningsSet);
}

/**
 * 박스 내 아이템들의 배송메모 수집 (중복 제거)
 * @param {Array} items - 아이템 배열
 * @returns {Array<string>}
 */
export function collectDeliveryMemos(items) {
    const memosSet = new Set();
    (items || []).forEach(item => {
        if (item.deliveryMemo && item.deliveryMemo.trim() !== '') {
            memosSet.add(item.deliveryMemo.trim());
        }
    });
    return Array.from(memosSet);
}

/**
 * 스마트 재단 로직 (애견롤매트 등 50cm 단위 판매 제품)
 * @param {Object} item - 아이템 정보
 * @returns {Array} 분리된 아이템 배열
 */
export function applySmartCutting(item) {
    // 애견롤매트만 적용: productCategory가 'petRoll'이거나 productId가 '4200445704'인 경우만
    // 유아롤매트(6092903705)는 옵션정보에 '50cm'가 포함될 수 있지만(예: "110x50x1.2cm"), 애견롤매트가 아님
    const isPetRoll = item.productCategory === 'petRoll' || item.productId === '4200445704';

    if (!isPetRoll) {
        // 유아롤매트는 그대로 반환 (수량과 길이를 분리해서 처리)
        return [item];
    }

    const unitLengthM = item.lengthM > 0 ? item.lengthM : 0.5;
    const totalLengthM = item.quantity * unitLengthM;
    const memo = item.deliveryMemo || '';
    const warnings = [];

    const cuttingSegments = parseCuttingSegmentsFromMemo(memo);
    if (cuttingSegments.length > 0) {
        const totalRequestedM = cuttingSegments.reduce((sum, seg) => sum + (seg.lengthM * seg.count), 0);
        const epsilon = 0.05;
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

    const rollMatch = memo.match(/(\d+)롤|(\d+)등분|반씩|반으로/);
    if (rollMatch) {
        if (rollMatch[1]) splitCount = parseInt(rollMatch[1]);
        else if (rollMatch[2]) splitCount = parseInt(rollMatch[2]);
        else splitCount = 2;
    } else {
        const lengthMatch = memo.match(/(\d+)m|(\d+)미터|(\d+)메터/i);
        if (lengthMatch) {
            const reqLength = parseInt(lengthMatch[1] || lengthMatch[2] || lengthMatch[3]);
            if (reqLength > 0 && reqLength < totalLengthM) {
                splitLength = reqLength;
            }
        }
    }

    const resultItems = [];

    if (splitLength > 0) {
        let remaining = totalLengthM;
        while (remaining >= splitLength) {
            resultItems.push({
                ...item,
                quantity: 1,
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

/**
 * 롤매트 최적 분할 조합 찾기 (배송비 최소화)
 * @param {Array} items - 롤매트 아이템 배열 (각 아이템은 {lengthM, widthNum, thicknessNum})
 * @param {Object} thresholds - 포장 임계값 {small, large, vinyl}
 * @param {Function} calculateBoxFee - 박스 배송비 계산 함수 (box) => fee
 * @returns {Array<Array>} 최적 분할 조합 (각 배열은 박스에 들어갈 아이템들)
 */
function findOptimalRollSplit(items, thresholds, calculateBoxFee) {
    if (items.length === 0) return [];
    
    // 성능 제한: 아이템이 너무 많으면 기본 분할 사용
    if (items.length > 20) {
        return null; // 기본 로직 사용
    }

    const totalLength = items.reduce((sum, item) => sum + (item.lengthM || 0), 0);
    
    // 단일 박스로 가능하면 그대로 반환
    if (totalLength <= thresholds.large) {
        return [items];
    }

    // 간단한 그리디 최적화: 가능한 조합들을 시도
    const minBoxes = Math.ceil(totalLength / thresholds.large);
    const maxBoxes = Math.min(minBoxes + 2, items.length);
    
    let bestSplit = null;
    let bestTotalFee = Infinity;

    // 길이 순으로 정렬
    const sortedItems = [...items].sort((a, b) => (b.lengthM || 0) - (a.lengthM || 0));

    function generateSplits(remainingItems, boxes, currentSplit) {
        if (boxes === 0) {
            if (remainingItems.length === 0) {
                // 유효한 분할 조합
                const totalFee = currentSplit.reduce((sum, boxItems) => {
                    const mockBox = {
                        type: 'roll',
                        items: boxItems,
                        packagingType: 'box'
                    };
                    return sum + calculateBoxFee(mockBox);
                }, 0);
                if (totalFee < bestTotalFee) {
                    bestTotalFee = totalFee;
                    bestSplit = currentSplit.map(box => [...box]);
                }
            }
            return;
        }

        if (remainingItems.length === 0) return;

        // 현재 박스에 아이템 추가 시도
        const currentBox = [];
        let currentLength = 0;

        for (let i = 0; i < remainingItems.length; i++) {
            const item = remainingItems[i];
            const itemLength = item.lengthM || 0;

            // 대박스 용량을 초과하면 중단
            if (currentLength + itemLength > thresholds.large) {
                break;
            }

            currentBox.push(item);
            currentLength += itemLength;
        }

                if (currentBox.length > 0) {
                    // 주문 정보 보존: 각 아이템의 원본 정보(rawRow, orderId 등)는 그대로 유지
                    const newRemaining = remainingItems.slice(currentBox.length);
                    // 아이템 객체는 참조가 아닌 복사본이므로 원본 주문 정보는 보존됨
                    currentSplit.push([...currentBox]); // 배열 복사
                    generateSplits(newRemaining, boxes - 1, currentSplit);
                    currentSplit.pop();
                }
    }

    // 박스 수를 늘려가며 최적 조합 찾기
    for (let numBoxes = minBoxes; numBoxes <= maxBoxes; numBoxes++) {
        generateSplits(sortedItems, numBoxes, []);
        if (bestSplit && numBoxes > minBoxes + 1) break;
    }

    return bestSplit;
}

/**
 * 퍼즐매트 최적 분할 조합 찾기 (배송비 최소화)
 * @param {number} totalCount - 총 수량
 * @param {number} capacity - 박스 용량
 * @param {Function} calculateBoxFee - 박스 배송비 계산 함수 (count) => fee
 * @returns {Array<number>} 최적 분할 조합 (예: [6, 5, 5, 5])
 */
function findOptimalPuzzleSplit(totalCount, capacity, calculateBoxFee) {
    if (totalCount <= capacity) {
        return [totalCount];
    }

    // 성능 제한: 너무 많은 수량이면 기본 분할 사용
    if (totalCount > 50) {
        const minBoxes = Math.ceil(totalCount / capacity);
        const result = [];
        let remaining = totalCount;
        for (let i = 0; i < minBoxes; i++) {
            const qty = Math.min(capacity, remaining);
            result.push(qty);
            remaining -= qty;
        }
        return result;
    }

    const minBoxes = Math.ceil(totalCount / capacity);
    // 최대 박스 수 제한 (성능 최적화)
    const maxBoxes = Math.min(minBoxes + 3, totalCount);
    
    let bestSplit = null;
    let bestTotalFee = Infinity;

    // 가능한 모든 분할 조합 시도 (제한된 범위 내에서)
    function generateSplits(remaining, boxes, currentSplit) {
        if (boxes === 0) {
            if (remaining === 0) {
                // 유효한 분할 조합
                const totalFee = currentSplit.reduce((sum, count) => sum + calculateBoxFee(count), 0);
                if (totalFee < bestTotalFee) {
                    bestTotalFee = totalFee;
                    bestSplit = [...currentSplit];
                }
            }
            return;
        }

        if (remaining <= 0) return;

        // 각 박스에 1장부터 capacity까지 시도
        const minQty = Math.max(1, Math.ceil(remaining / boxes));
        const maxQty = Math.min(capacity, remaining - (boxes - 1));

        // 범위가 너무 크면 제한
        if (maxQty - minQty > 10) {
            // 대표적인 값들만 시도 (성능 최적화)
            const candidates = [minQty, Math.floor((minQty + maxQty) / 2), maxQty];
            for (const qty of candidates) {
                if (qty >= minQty && qty <= maxQty) {
                    currentSplit.push(qty);
                    generateSplits(remaining - qty, boxes - 1, currentSplit);
                    currentSplit.pop();
                }
            }
        } else {
            for (let qty = minQty; qty <= maxQty; qty++) {
                currentSplit.push(qty);
                generateSplits(remaining - qty, boxes - 1, currentSplit);
                currentSplit.pop();
            }
        }
    }

    // 박스 수를 늘려가며 최적 조합 찾기 (제한된 범위)
    for (let numBoxes = minBoxes; numBoxes <= maxBoxes; numBoxes++) {
        generateSplits(totalCount, numBoxes, []);
        // 이미 최적 조합을 찾았고 더 많은 박스는 비용이 증가할 것이므로 조기 종료
        if (bestSplit && numBoxes > minBoxes + 2) break;
    }

    return bestSplit || [totalCount]; // 최적 조합이 없으면 전체를 하나의 박스로
}

/**
 * 패킹 처리 메인 함수
 * @param {Object} groupedOrders - 그룹화된 주문 객체
 * @param {Function} generateDesignCode - 디자인 코드 생성 함수
 * @param {Array} shippingFees - 배송비 데이터 (퍼즐매트 최적화용)
 * @returns {Array} 패킹된 주문 배열
 */
export function processPacking(groupedOrders, generateDesignCode, shippingFees = []) {
    const packedOrders = [];

    Object.values(groupedOrders).forEach(group => {
        group._warnings = [];
        const boxes = [];
        let standaloneBoxes = [];

        // 스마트 재단 로직 적용
        const processedRollItems = group.rollItems.flatMap(item => applySmartCutting(item));
        group._warnings = collectWarningsFromItems(processedRollItems);

        if (group.hasCuttingRequest) {
            const memos = group.items.map(o => o.deliveryMemo).filter(m => m && m.trim());
            const uniqueMemos = [...new Set(memos)];
            group._warnings.unshift(`⚠️ 확인: ${uniqueMemos.join(' / ')}`);
        }

        // 롤매트 합포장 로직
        // 모든 롤매트 아이템을 합포장 대상으로 먼저 고려
        // 합포장 후 총 길이가 비닐 임계값을 넘으면 비닐로 처리
        const allRollItems = [];
        processedRollItems.forEach(item => {
            for (let i = 0; i < item.quantity; i++) {
                allRollItems.push({ ...item, quantity: 1 });
            }
        });

        if (allRollItems.length > 0) {
            const maxThickness = Math.max(...allRollItems.map(i => i.thicknessNum || 17));
            const thresholds = PACKAGING_THRESHOLDS[String(maxThickness)] || PACKAGING_THRESHOLDS["17"];

            // 롤매트 최적화는 복잡하므로 일단 기본 로직 사용
            // (롤매트는 길이 기반이고 합포장 시 포장 타입이 변경될 수 있어 최적화가 더 복잡함)
            let currentBox = { items: [], totalLength: 0 };
            const finishedBoxes = [];

            // 합포장 시도: 대박스 용량 또는 비닐 최대 길이를 초과하면 새 박스로 분리
            // 실무자 판단: 대박스 용량 초과 시 각각 개별 박스로 처리
            allRollItems.forEach(item => {
                const len = item.lengthM || 0;
                const itemThickness = item.thicknessNum || maxThickness;
                const itemThresholds = PACKAGING_THRESHOLDS[String(itemThickness)] || thresholds;
                
                // 현재 박스의 총 길이 + 새 아이템 길이
                const newTotalLength = currentBox.totalLength + len;
                
                // 대박스 용량 또는 비닐 최대 길이를 초과하면 새 박스로 분리
                const exceedsLarge = newTotalLength > thresholds.large;
                const exceedsVinylMax = newTotalLength > (itemThresholds.vinylMax || Infinity);
                
                if (exceedsLarge || exceedsVinylMax) {
                    if (currentBox.items.length > 0) {
                        finishedBoxes.push(currentBox);
                    }
                    currentBox = { items: [item], totalLength: len };
                } else {
                    // 대박스 용량 및 비닐 최대 길이 내에서 합포장
                    currentBox.items.push(item);
                    currentBox.totalLength += len;
                }
            });
            if (currentBox.items.length > 0) {
                finishedBoxes.push(currentBox);
            }

            // 최적화 실패 시 기본 로직 사용
            if (finishedBoxes.length === 0) {
                let currentBox = { items: [], totalLength: 0 };

                // 합포장 시도: 대박스 용량 또는 비닐 최대 길이를 초과하면 새 박스로 분리
                // 실무자 판단: 대박스 용량 초과 시 각각 개별 박스로 처리
                allRollItems.forEach(item => {
                    const len = item.lengthM || 0;
                    const itemThickness = item.thicknessNum || maxThickness;
                    const itemThresholds = PACKAGING_THRESHOLDS[String(itemThickness)] || thresholds;
                    
                    // 현재 박스의 총 길이 + 새 아이템 길이
                    const newTotalLength = currentBox.totalLength + len;
                    
                    // 대박스 용량 또는 비닐 최대 길이를 초과하면 새 박스로 분리
                    const exceedsLarge = newTotalLength > thresholds.large;
                    const exceedsVinylMax = newTotalLength > (itemThresholds.vinylMax || Infinity);
                    
                    if (exceedsLarge || exceedsVinylMax) {
                        if (currentBox.items.length > 0) {
                            finishedBoxes.push(currentBox);
                        }
                        currentBox = { items: [item], totalLength: len };
                    } else {
                        // 대박스 용량 및 비닐 최대 길이 내에서 합포장
                        currentBox.items.push(item);
                        currentBox.totalLength += len;
                    }
                });
                if (currentBox.items.length > 0) {
                    finishedBoxes.push(currentBox);
                }
            }

            // 합포장된 박스들을 최종 포장 타입 결정
            finishedBoxes.forEach(box => {
                const totalLen = box.totalLength;
                const boxMaxThickness = Math.max(...box.items.map(i => i.thicknessNum || 17));
                const boxThresholds = PACKAGING_THRESHOLDS[String(boxMaxThickness)] || PACKAGING_THRESHOLDS["17"];

                let packagingType = 'smallBox';
                let needsStar = false;

                // 합포장 후 총 길이가 비닐 임계값을 넘으면 비닐로 처리
                if (totalLen >= boxThresholds.vinyl) {
                    packagingType = 'vinyl';
                    needsStar = true;
                } else if (totalLen > boxThresholds.small) {
                    packagingType = 'largeBox';
                }

                // 같은 디자인코드와 길이를 가진 아이템들을 그룹핑하여 수량 합산
                const itemsByKey = new Map();
                box.items.forEach(i => {
                    const code = generateDesignCode(i);
                    const key = `${code}_${i.lengthM || 0}`;
                    if (!itemsByKey.has(key)) {
                        itemsByKey.set(key, {
                            code: code,
                            lengthM: i.lengthM || 0,
                            qty: 0
                        });
                    }
                    itemsByKey.get(key).qty += (i.quantity || 1);
                });

                const designText = Array.from(itemsByKey.values()).map(item => {
                    const lengthSuffix = item.lengthM > 0 ? `${item.lengthM}m` : '';
                    const qtySuffix = item.qty > 1 ? ` x${item.qty}` : '';
                    return `${item.code}${lengthSuffix}${qtySuffix}`;
                }).join(' / ');

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

        // 퍼즐매트 처리 (배송비 최적화 적용)
        if (group.puzzleItems.length > 0) {
            const puzzleItemsCopy = group.puzzleItems.map(item => ({ ...item }));
            
            // 총 수량 계산
            const totalCount = puzzleItemsCopy.reduce((sum, item) => sum + item.quantity, 0);
            const thickness = puzzleItemsCopy[0].thicknessNum || 25;
            const capacity = PUZZLE_BOX_CAPACITY[String(thickness)] || 6;
            const width = puzzleItemsCopy[0].widthNum || 100;
            const thicknessCm = thickness / 10;

            // 배송비 계산 헬퍼 함수
            const calculateBoxFee = (count) => {
                if (!shippingFees || shippingFees.length === 0) return 0;
                
                for (const fee of shippingFees) {
                    if (!fee.productGroup.includes('퍼즐매트')) continue;
                    if (!fee.packageType.includes('강화비닐')) continue;
                    
                    const widthMatch = !fee.width || Math.abs(fee.width - width) <= 5;
                    const thicknessMatch = !fee.thickness || Math.abs(fee.thickness - thicknessCm) <= 0.1;
                    const lengthMatch = !fee.lengthMin || !fee.lengthMax || 
                        (count >= fee.lengthMin && count <= fee.lengthMax);
                    
                    if (widthMatch && thicknessMatch && lengthMatch) {
                        return fee.fee;
                    }
                }
                return 0;
            };

            // 최적 분할 조합 찾기
            const optimalSplit = findOptimalPuzzleSplit(totalCount, capacity, calculateBoxFee);

            // 최적 조합에 따라 박스 생성
            let itemIndex = 0;
            for (let i = 0; i < optimalSplit.length; i++) {
                const targetCount = optimalSplit[i];
                const boxItems = [];
                let remainingCount = targetCount;

                while (itemIndex < puzzleItemsCopy.length && remainingCount > 0) {
                    const item = puzzleItemsCopy[itemIndex];
                    const takeQty = Math.min(item.quantity, remainingCount);

                    if (takeQty > 0) {
                        boxItems.push({ ...item, quantity: takeQty });
                        remainingCount -= takeQty;
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
                    const boxMemo = boxItems.find(i => i.deliveryMemo)?.deliveryMemo || '';
                    // 각 박스에 담긴 아이템 기준으로 디자인 코드 생성
                    const boxDesignText = boxItems.map(item => generateDesignCode(item)).join(' ');
                    standaloneBoxes.push({
                        type: 'puzzle',
                        packagingType: 'vinyl',
                        needsStar: false,
                        items: boxItems,
                        designText: boxDesignText,
                        isCombined: false,
                        remark: '',
                        deliveryMemo: boxMemo
                    });
                }
            }
        }

        // 기타 아이템 처리
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

        // 테이프 처리
        if (group.tapeItems.length > 0) {
            if (standaloneBoxes.length > 0) {
                const targetBox = standaloneBoxes[standaloneBoxes.length - 1];
                group.tapeItems.forEach(item => {
                    targetBox.items.push({ ...item });
                });
                targetBox.isCombined = true;
                targetBox.remark = '합';

                const tapeDesigns = group.tapeItems.map(i => generateDesignCode(i)).join('+');
                targetBox.designText += '+' + tapeDesigns;

                if (!targetBox.deliveryMemo) {
                    const tapeMemo = group.tapeItems.find(i => i.deliveryMemo)?.deliveryMemo;
                    if (tapeMemo) targetBox.deliveryMemo = tapeMemo;
                }
            } else {
                const tapeDesigns = group.tapeItems.map(i => generateDesignCode(i)).join('+');
                const tapeMemo = group.tapeItems.find(i => i.deliveryMemo)?.deliveryMemo || '';

                standaloneBoxes.push({
                    type: 'roll',
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

        boxes.push(...standaloneBoxes);

        // N-n 라벨링
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
            const totalQtyInBox = box.items.reduce((sum, item) => sum + (item.quantity || 0), 0);
            const boxTotalPrice = box.items.reduce((sum, item) => sum + ((item.price || 0) * (item.quantity || 1)), 0);
            const boxMemos = collectDeliveryMemos(box.items);

            // 택배사 결정 로직
            let courier = 'kyungdong'; // 기본값

            if (box.packagingType === 'smallBox') {
                // 소박스는 기본적으로 로젠
                courier = 'logen';
            } else {
                // 대박스/비닐인데 70cm 폭 단독 배송이면 로젠
                // 70cm 단독 = 박스 내 모든 아이템이 70cm 폭이고 합포장되지 않음
                const all70cm = box.items.every(item => item.widthNum === 70);
                const isSingle = !box.isCombined && box.items.length === 1;

                if (all70cm && isSingle) {
                    courier = 'logen';
                } else {
                    courier = 'kyungdong';
                }
            }

            packedOrders.push({
                ...box,
                group: group,
                customerName: group.customerName,
                address: group.address,
                zipCode: group.zipCode,
                phone: group.phone,
                totalQtyInBox: totalQtyInBox,
                deliveryMemos: boxMemos,
                totalPrice: boxTotalPrice,
                giftEligible: group.giftEligible,
                hasCuttingRequest: group.hasCuttingRequest,
                hasFinishingRequest: group.hasFinishingRequest,
                warnings: [...(group._warnings || []), ...(box.warnings && box.warnings.length ? box.warnings : collectWarningsFromItems(box.items))],
                courier: courier
            });
        });
    });

    return packedOrders;
}

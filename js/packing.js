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
    const isPetRoll = item.productCategory === 'petRoll' ||
        (item.rawRow['옵션정보'] && item.rawRow['옵션정보'].includes('50cm'));

    if (!isPetRoll) {
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
 * 패킹 처리 메인 함수
 * @param {Object} groupedOrders - 그룹화된 주문 객체
 * @param {Function} generateDesignCode - 디자인 코드 생성 함수
 * @returns {Array} 패킹된 주문 배열
 */
export function processPacking(groupedOrders, generateDesignCode) {
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
        const vinylItems = [];
        const combinableItems = [];

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
            const designCode = generateDesignCode(item);
            const lengthSuffix = item.lengthM > 0 ? `${item.lengthM}m` : '';
            standaloneBoxes.push({
                type: 'roll',
                packagingType: 'vinyl',
                needsStar: true,
                items: [item],
                designText: `${designCode}${lengthSuffix}`,
                isCombined: false,
                remark: '',
                deliveryMemo: item.deliveryMemo,
                warnings: collectWarningsFromItems([item])
            });
        });

        // 합포장 가능 아이템들 처리
        if (combinableItems.length > 0) {
            const maxThickness = Math.max(...combinableItems.map(i => i.thicknessNum || 17));
            const thresholds = PACKAGING_THRESHOLDS[String(maxThickness)] || PACKAGING_THRESHOLDS["17"];

            let currentBox = { items: [], totalLength: 0 };
            const finishedBoxes = [];

            combinableItems.forEach(item => {
                const len = item.lengthM || 0;
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

            finishedBoxes.forEach(box => {
                const totalLen = box.totalLength;
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
                    const code = generateDesignCode(i);
                    const lengthSuffix = i.lengthM > 0 ? `${i.lengthM}m` : '';
                    return `${code}${lengthSuffix}`;
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

        // 퍼즐매트 처리
        if (group.puzzleItems.length > 0) {
            const puzzleItemsCopy = group.puzzleItems.map(item => ({ ...item }));
            const puzzleBoxCount = calculatePuzzleBoxes(puzzleItemsCopy);

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
            const courier = box.packagingType === 'smallBox' ? 'logen' : 'kyungdong';

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

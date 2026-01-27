import {
    parseSoldierPile,
    calculatePileWeight,
    isSoldierPile,
    isPrimarySecantPile,
    isSecondarySecantPile,
    isTangentPile,
    isSheetPile,
    isTimberLagging,
    isTimberSheeting,
    isWaler,
    isRaker,
    isUpperRaker,
    isLowerRaker,
    isStandOff,
    isKicker,
    isChannel,
    isRollChock,
    isStudBeam,
    isInnerCornerBrace,
    isKneeBrace,
    parseSoeItem
} from '../parsers/soeParser'
import { SHEET_PILE_WEIGHTS } from '../constants/sheetPileWeight'

/**
 * Processes soldier pile items and groups them appropriately
 */
export const processSoldierPileItems = (rawDataRows, headers) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const allItems = []
    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = row[unitIdx]

        if (isSoldierPile(digitizerItem)) {
            const parsed = parseSoldierPile(digitizerItem)
            const weight = calculatePileWeight(parsed)

            allItems.push({
                particulars: digitizerItem,
                takeoff: total,
                unit: unit,
                parsed: parsed,
                weight: weight,
                rawRowNumber: rowIndex + 2
            })
        }
    })

    const groupMap = new Map()
    allItems.forEach(item => {
        const groupKey = item.parsed.groupKey
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                type: item.parsed.type,
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    let groups = Array.from(groupMap.values())
    const hpGroups = groups.filter(g => g.type === 'hp')
    const drilledGroups = groups.filter(g => g.type === 'drilled')

    const heightCounts = new Map()
    hpGroups.forEach(group => {
        const height = group.parsed.heightRaw
        heightCounts.set(height, (heightCounts.get(height) || 0) + group.items.length)
    })

    const duplicateHPGroups = []
    const uniqueHPItems = []
    hpGroups.forEach(group => {
        const height = group.parsed.heightRaw
        if (heightCounts.get(height) > 1) {
            duplicateHPGroups.push(group)
        } else {
            uniqueHPItems.push(...group.items)
        }
    })

    const regroupedHP = []
    if (uniqueHPItems.length > 0) {
        regroupedHP.push({
            groupKey: 'HP-UNIQUE',
            type: 'hp',
            items: uniqueHPItems,
            parsed: uniqueHPItems[0].parsed
        })
    }
    regroupedHP.push(...duplicateHPGroups)
    groups = [...drilledGroups, ...regroupedHP]

    groups.sort((a, b) => {
        if (a.type !== b.type) return a.type === 'drilled' ? -1 : 1
        if (a.type === 'drilled') {
            if (a.parsed.diameter !== b.parsed.diameter) return a.parsed.diameter - b.parsed.diameter
            if (a.parsed.thickness !== b.parsed.thickness) return a.parsed.thickness - b.parsed.thickness
            const patternOrder = { 'E': 1, 'E+RS': 2, 'RS': 3, 'H': 4 }
            const aPattern = a.groupKey.includes('E+RS') ? 'E+RS' : a.groupKey.includes('-E-') ? 'E' : a.groupKey.includes('-RS-') ? 'RS' : 'H'
            const bPattern = b.groupKey.includes('E+RS') ? 'E+RS' : b.groupKey.includes('-E-') ? 'E' : b.groupKey.includes('-RS-') ? 'RS' : 'H'
            return (patternOrder[aPattern] || 0) - (patternOrder[bPattern] || 0)
        } else {
            if (a.groupKey === 'HP-UNIQUE') return -1
            if (b.groupKey === 'HP-UNIQUE') return 1
            return a.parsed.calculatedHeight - b.parsed.calculatedHeight
        }
    })

    return groups
}

/**
 * Generic process function for simple SOE items
 */
const processGenericSoeItems = (rawDataRows, headers, identifierFn) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')
    const qtyIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'count') // Assume count for QTY

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const items = []
    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = row[unitIdx]
        const qty = qtyIdx !== -1 ? parseFloat(row[qtyIdx]) || '' : ''

        if (identifierFn(digitizerItem)) {
            const parsed = parseSoeItem(digitizerItem)

            // For sheet piles, lookup weight
            if (parsed.type === 'sheet_pile') {
                const sheetPileNames = Object.keys(SHEET_PILE_WEIGHTS).sort((a, b) => b.length - a.length)
                for (const name of sheetPileNames) {
                    let pattern = name
                        .replace(/[.*+?^${}()|[\]\\]/g, '\\$&')
                        .replace(/[-\s]/g, '[- ]?')

                    const regex = new RegExp(pattern, 'i')
                    if (regex.test(digitizerItem)) {
                        parsed.weight = SHEET_PILE_WEIGHTS[name]
                        break
                    }
                }
            }

            items.push({
                particulars: digitizerItem,
                takeoff: total,
                unit: unit,
                qty: qty,
                parsed: parsed,
                weight: parsed.weight,
                rawRowNumber: rowIndex + 2
            })
        }
    })
    return items
}

export const processPrimarySecantItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isPrimarySecantPile)
export const processSecondarySecantItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isSecondarySecantPile)
export const processTangentPileItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isTangentPile)
export const processSheetPileItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isSheetPile)
export const processTimberLaggingItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isTimberLagging)
export const processTimberSheetingItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isTimberSheeting)
export const processWalerItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isWaler)
export const processRakerItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isRaker)
export const processUpperRakerItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isUpperRaker)
export const processLowerRakerItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isLowerRaker)
export const processStandOffItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isStandOff)
export const processKickerItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isKicker)
export const processChannelItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isChannel)
export const processRollChockItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isRollChock)
export const processStudBeamItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isStudBeam)
export const processInnerCornerBraceItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isInnerCornerBrace)
export const processKneeBraceItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isKneeBrace)

/**
 * Generates formulas for SOE items
 */
export const generateSoeFormulas = (itemType, rowNum, itemData) => {
    const formulas = {
        ft: null,
        sqFt: null,
        lbs: null,
        cy: null,
        qtyFinal: null,
        takeoff: null
    }

    const type = itemData.parsed?.type || itemType

    switch (type) {
        case 'hp':
        case 'drilled':
        case 'soldier_pile':
            formulas.ft = `H${rowNum}*C${rowNum}`
            formulas.lbs = `I${rowNum}*${(itemData.weight || 0).toFixed(3)}`
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'primary_secant':
        case 'tangent':
            formulas.ft = `H${rowNum}*C${rowNum}`
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'secondary_secant':
            formulas.ft = `H${rowNum}*C${rowNum}`
            formulas.lbs = `I${rowNum}*${(itemData.weight || 0).toFixed(3)}`
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'sheet_pile':
            // Sheet pile: FT(I)=C, SQ FT(J)=I*H, LBS(K)=J*Wt
            formulas.ft = `C${rowNum}`
            formulas.sqFt = `I${rowNum}*H${rowNum}`
            formulas.lbs = `J${rowNum}*${(itemData.weight || 0).toFixed(3)}`
            break

        case 'timber_logging':
        case 'timber_lagging':
        case 'timber_sheeting':
            // Timber lagging/sheeting: FT(I)=C, SQ FT(J)=I*H
            formulas.ft = `C${rowNum}`
            formulas.sqFt = `I${rowNum}*H${rowNum}`
            break

        case 'backpacking_item':
            if (itemData.timberLaggingSumRow) {
                formulas.takeoff = `J${itemData.timberLaggingSumRow}`
            }
            formulas.sqFt = `C${rowNum}`
            formulas.ft = null
            break

        case 'waler':
        case 'inner_corner_brace':
            // FT(I)=C, LBS(K)=I*Wt, QTY(M)=E
            formulas.ft = `C${rowNum}`
            formulas.lbs = `I${rowNum}*${(itemData.weight || 0).toFixed(3)}`
            formulas.qtyFinal = `E${rowNum}`
            break

        case 'raker':
        case 'upper_raker':
        case 'lower_raker':
            // FT(I)=C*1.15, LBS(K)=I*Wt, QTY(M)=E
            formulas.ft = `C${rowNum}*1.15`
            formulas.lbs = `I${rowNum}*${(itemData.weight || 0).toFixed(3)}`
            formulas.qtyFinal = `E${rowNum}`
            break

        case 'stand_off':
        case 'kicker':
            // FT(I)=C, LBS(K)=I*Wt, QTY(M)=E
            formulas.ft = `C${rowNum}`
            formulas.lbs = `I${rowNum}*${(itemData.weight || 0).toFixed(3)}`
            formulas.qtyFinal = `E${rowNum}`
            break

        case 'channel':
        case 'roll_chock':
        case 'stud_beam':
        case 'knee_brace':
            // FT(I)=C*H (or H*C), LBS(K)=I*Wt, QTY(M)=C
            // Image shows H*C or C*H. Let's use H*C as standard if taking EA
            formulas.ft = `C${rowNum}*H${rowNum}`
            formulas.lbs = `I${rowNum}*${(itemData.weight || 0).toFixed(3)}`
            formulas.qtyFinal = `C${rowNum}`
            break
    }

    return formulas
}

export default {
    processSoldierPileItems,
    processPrimarySecantItems,
    processSecondarySecantItems,
    processTangentPileItems,
    processSheetPileItems,
    processTimberLaggingItems,
    processTimberSheetingItems,
    processWalerItems,
    processRakerItems,
    processUpperRakerItems,
    processLowerRakerItems,
    processStandOffItems,
    processKickerItems,
    processChannelItems,
    processRollChockItems,
    processStudBeamItems,
    processInnerCornerBraceItems,
    processKneeBraceItems,
    generateSoeFormulas
}

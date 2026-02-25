import { mergeSingleItemGroups } from '../groupingUtils.js'
import {
    parseSoldierPile,
    parseTimberSoldierPile,
    parseTimberPlank,
    parseTimberRaker,
    parseTimberBrace,
    parseTimberWaler,
    parseTimberStringer,
    parseTimberPost,
    parseTimberSheets,
    parseVerticalTimberSheets,
    parseHorizontalTimberSheets,
    calculatePileWeight,
    isSoldierPile,
    isTimberSoldierPile,
    isTimberPlank,
    isTimberRaker,
    isTimberBrace,
    isTimberWaler,
    isTimberStringer,
    isTimberPost,
    isTimberSheets,
    isVerticalTimberSheets,
    isHorizontalTimberSheets,
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
    isCornerbrace,
    isKneeBrace,
    isSupportingAngle,
    isParging,
    isHeelBlock,
    isUnderpinning,
    isRockAnchor,
    isRockBolt,
    isAnchor,
    isTieBack,
    isConcreteSoilRetentionPier,
    isGuideWall,
    isDowelBar,
    isRockPin,
    isShotcrete,
    isPermissionGrouting,
    isButton,
    isRockStabilization,
    isFormBoard,
    isDrilledHoleGrout,
    parseDrilledHoleGrout,
    parseSoeItem
} from '../parsers/soeParser'
import { SHEET_PILE_WEIGHTS } from '../constants/sheetPileWeight'

/**
 * Processes soldier pile items and groups them appropriately
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 */
export const processSoldierPileItems = (rawDataRows, headers, tracker = null) => {
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

            // Mark this row as used
            if (tracker) {
                tracker.markUsed(rowIndex)
            }
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

    return mergeSingleItemGroups(groups)
}

/**
 * Processes timber soldier pile items and groups them (same structure as drilled soldier piles)
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 */
export const processTimberSoldierPileItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const allItems = []
    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = row[unitIdx]

        if (isTimberSoldierPile(digitizerItem)) {
            const parsed = parseTimberSoldierPile(digitizerItem)

            allItems.push({
                particulars: digitizerItem,
                takeoff: total,
                unit: unit,
                parsed: parsed,
                weight: 0,
                rawRowNumber: rowIndex + 2
            })

            if (tracker) {
                tracker.markUsed(rowIndex)
            }
        }
    })

    const groupMap = new Map()
    allItems.forEach(item => {
        const groupKey = item.parsed.groupKey
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                type: 'timber',
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    let groups = Array.from(groupMap.values()).sort((a, b) => {
        if (a.parsed.heightRaw !== b.parsed.heightRaw) return a.parsed.heightRaw - b.parsed.heightRaw
        return (a.parsed.embedment || 0) - (b.parsed.embedment || 0)
    })
    return mergeSingleItemGroups(groups)
}

/**
 * Processes timber plank items and groups them (same structure as timber soldier piles, no column K, H not rounded)
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 */
export const processTimberPlankItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const allItems = []
    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = row[unitIdx]

        if (isTimberPlank(digitizerItem)) {
            const parsed = parseTimberPlank(digitizerItem)

            allItems.push({
                particulars: digitizerItem,
                takeoff: total,
                unit: unit,
                parsed: parsed,
                weight: 0,
                rawRowNumber: rowIndex + 2
            })

            if (tracker) {
                tracker.markUsed(rowIndex)
            }
        }
    })

    const groupMap = new Map()
    allItems.forEach(item => {
        const groupKey = item.parsed.groupKey
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                type: 'timber_plank',
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    let groups = Array.from(groupMap.values()).sort((a, b) => {
        return (a.parsed.heightRaw || 0) - (b.parsed.heightRaw || 0)
    })
    return mergeSingleItemGroups(groups)
}

/**
 * Processes timber raker (wood raker) items and groups them (same structure as timber planks: no column K, H not rounded)
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 */
export const processTimberRakerItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const allItems = []
    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = row[unitIdx]

        if (isTimberRaker(digitizerItem)) {
            const parsed = parseTimberRaker(digitizerItem)

            allItems.push({
                particulars: digitizerItem,
                takeoff: total,
                unit: unit,
                parsed: parsed,
                weight: 0,
                rawRowNumber: rowIndex + 2
            })

            if (tracker) {
                tracker.markUsed(rowIndex)
            }
        }
    })

    const groupMap = new Map()
    allItems.forEach(item => {
        const groupKey = item.parsed.groupKey
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                type: 'timber_raker',
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    let groups = Array.from(groupMap.values()).sort((a, b) => {
        return (a.parsed.heightRaw || 0) - (b.parsed.heightRaw || 0)
    })
    return mergeSingleItemGroups(groups)
}

/**
 * Processes timber brace items (Horizontal wood brace and Timber Corner brace) and groups them
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 */
export const processTimberBraceItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const allItems = []
    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = row[unitIdx]

        if (isTimberBrace(digitizerItem)) {
            const parsed = parseTimberBrace(digitizerItem)

            allItems.push({
                particulars: digitizerItem,
                takeoff: total,
                unit: unit,
                parsed: parsed,
                weight: 0,
                rawRowNumber: rowIndex + 2
            })

            if (tracker) {
                tracker.markUsed(rowIndex)
            }
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

    let groups = Array.from(groupMap.values()).sort((a, b) => {
        const aH = a.parsed.heightRaw || 0
        const bH = b.parsed.heightRaw || 0
        if (aH !== bH) return aH - bH
        return (a.parsed.qty || 0) - (b.parsed.qty || 0)
    })
    return mergeSingleItemGroups(groups)
}

/**
 * Processes timber post items and groups them (same structure as timber soldier piles, no column K)
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 */
export const processTimberPostItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const allItems = []
    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = row[unitIdx]

        if (isTimberPost(digitizerItem)) {
            const parsed = parseTimberPost(digitizerItem)

            allItems.push({
                particulars: digitizerItem,
                takeoff: total,
                unit: unit,
                parsed: parsed,
                weight: 0,
                rawRowNumber: rowIndex + 2
            })

            if (tracker) {
                tracker.markUsed(rowIndex)
            }
        }
    })

    const groupMap = new Map()
    allItems.forEach(item => {
        const groupKey = item.parsed.groupKey
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                type: 'timber_post',
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    let groups = Array.from(groupMap.values()).sort((a, b) => {
        if (a.parsed.heightRaw !== b.parsed.heightRaw) return a.parsed.heightRaw - b.parsed.heightRaw
        return (a.parsed.embedment || 0) - (b.parsed.embedment || 0)
    })
    return mergeSingleItemGroups(groups)
}

/**
 * Processes timber sheets items (vertical, horizontal, wood, wooden, or plain)
 * and groups them by dimension + H + E — same logic as before, just one unified section.
 * Formulas: FT(I)=C, SQ FT(J)=I*H, no column K
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 */
export const processTimberSheetsItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const allItems = []
    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = row[unitIdx]

        if (isTimberSheets(digitizerItem)) {
            const parsed = parseTimberSheets(digitizerItem)

            allItems.push({
                particulars: digitizerItem,
                takeoff: total,
                unit: unit,
                parsed: parsed,
                weight: 0,
                rawRowNumber: rowIndex + 2
            })

            if (tracker) {
                tracker.markUsed(rowIndex)
            }
        }
    })

    const groupMap = new Map()
    allItems.forEach(item => {
        const groupKey = item.parsed.groupKey
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                type: 'timber_sheets',
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    let groups = Array.from(groupMap.values()).sort((a, b) => {
        if (a.parsed.heightRaw !== b.parsed.heightRaw) return a.parsed.heightRaw - b.parsed.heightRaw
        return (a.parsed.embedment || 0) - (b.parsed.embedment || 0)
    })
    return mergeSingleItemGroups(groups)
}

// Keep old names as aliases so any lingering call-sites don't break
export const processVerticalTimberSheetsItems = processTimberSheetsItems
export const processHorizontalTimberSheetsItems = processTimberSheetsItems

/**
 * Processes timber waler items and groups them
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 */
export const processTimberWalerItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const allItems = []
    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = row[unitIdx]

        if (isTimberWaler(digitizerItem)) {
            const parsed = parseTimberWaler(digitizerItem)

            allItems.push({
                particulars: digitizerItem,
                takeoff: total,
                unit: unit,
                parsed: parsed,
                weight: 0,
                rawRowNumber: rowIndex + 2
            })

            if (tracker) {
                tracker.markUsed(rowIndex)
            }
        }
    })

    const groupMap = new Map()
    allItems.forEach(item => {
        const groupKey = item.parsed.groupKey
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                type: 'timber_waler',
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    let groups = Array.from(groupMap.values()).sort((a, b) => {
        return (a.parsed.qty || 0) - (b.parsed.qty || 0)
    })
    return mergeSingleItemGroups(groups)
}

/**
 * Processes timber stringer items and groups them by size
 * E=qty from (N) or 1, I=C*E. Per-group: multiple items -> sum I, items black; single item -> no sum, item row red
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 */
export const processTimberStringerItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const allItems = []
    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = row[unitIdx]

        if (isTimberStringer(digitizerItem)) {
            const parsed = parseTimberStringer(digitizerItem)

            allItems.push({
                particulars: digitizerItem,
                takeoff: total,
                unit: unit,
                parsed: parsed,
                weight: 0,
                rawRowNumber: rowIndex + 2
            })

            if (tracker) {
                tracker.markUsed(rowIndex)
            }
        }
    })

    const groupMap = new Map()
    allItems.forEach(item => {
        const groupKey = item.parsed.groupKey
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                type: 'timber_stringer',
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    let groups = Array.from(groupMap.values()).sort((a, b) => {
        return (a.parsed.qty || 0) - (b.parsed.qty || 0)
    })
    return mergeSingleItemGroups(groups)
}

/**
 * Processes drilled hole grout items and groups them by H (height)
 * F=G=SQRT((d/12)^2*3.14/4), H=height, I=C*H, J=G*F*C, L=J*H/27, M=C
 * Per-group: multiple items -> sum row, items black; single item -> no sum, item row red
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 */
export const processDrilledHoleGroutItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    let totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    if (totalIdx === -1) totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'takeoff')
    let unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')
    if (unitIdx === -1) unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'unit')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const allItems = []
    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = row[unitIdx]

        if (isDrilledHoleGrout(digitizerItem)) {
            const parsed = parseDrilledHoleGrout(digitizerItem)

            allItems.push({
                particulars: digitizerItem,
                takeoff: total,
                unit: unit,
                parsed: parsed,
                weight: 0,
                rawRowNumber: rowIndex + 2
            })

            if (tracker) {
                tracker.markUsed(rowIndex)
            }
        }
    })

    const groupMap = new Map()
    allItems.forEach(item => {
        const groupKey = item.parsed.groupKey
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                type: 'drilled_hole_grout',
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    let groups = Array.from(groupMap.values()).sort((a, b) => {
        return (a.parsed.heightRaw || 0) - (b.parsed.heightRaw || 0)
    })
    return mergeSingleItemGroups(groups)
}

/**
 * Generic process function for simple SOE items
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 */
const processGenericSoeItems = (rawDataRows, headers, identifierFn, tracker = null) => {
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

            // Mark this row as used
            if (tracker) {
                tracker.markUsed(rowIndex)
            }
        }
    })
    return items
}

export const processPrimarySecantItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericSoeItems(rawDataRows, headers, isPrimarySecantPile, tracker)

    // Generate proposal text line that can be made
    if (items.length > 0) {
        const firstItem = items[0]
        const totalQty = items.reduce((sum, item) => sum + (parseFloat(item.qty) || 0), 0)
        const totalTakeoff = items.reduce((sum, item) => sum + (item.takeoff || 0), 0)

        // Extract diameter from particulars
        let diameter = null
        const diameterMatch = firstItem.particulars?.match(/([0-9.]+)["\s]*Ø/i)
        if (diameterMatch) {
            diameter = parseFloat(diameterMatch[1])
        }

        // Calculate average height
        let totalHeight = 0
        let heightCount = 0
        items.forEach(item => {
            const height = item.parsed?.calculatedHeight || item.parsed?.heightRaw || 0
            if (height > 0) {
                totalHeight += height * (item.takeoff || 0)
                heightCount += (item.takeoff || 0)
            }
        })
        const avgHeight = heightCount > 0 ? totalHeight / heightCount : 0

        // Extract embedment
        let totalEmbedment = 0
        let embedmentCount = 0
        items.forEach(item => {
            const particulars = item.particulars || ''
            const eMatch = particulars.match(/E=([0-9'"\-]+)/i)
            if (eMatch) {
                const embedmentStr = eMatch[1]
                const embedmentMatch = embedmentStr.match(/(\d+)(?:'-?)?(\d+)?/)
                if (embedmentMatch) {
                    const feet = parseInt(embedmentMatch[1]) || 0
                    const inches = parseInt(embedmentMatch[2]) || 0
                    const embedment = feet + (inches / 12)
                    totalEmbedment += embedment * (item.takeoff || 0)
                    embedmentCount += (item.takeoff || 0)
                }
            }
        })
        const avgEmbedment = embedmentCount > 0 ? totalEmbedment / embedmentCount : 0

        // Format proposal text line
        const roundToMultipleOf5 = (value) => Math.ceil(value / 5) * 5
        const avgHeightRounded = roundToMultipleOf5(avgHeight)
        const avgEmbedmentRounded = roundToMultipleOf5(avgEmbedment)

        const heightFeet = Math.floor(avgHeightRounded)
        const heightInches = Math.round((avgHeightRounded - heightFeet) * 12)
        const heightText = heightInches === 0 ? `${heightFeet}'-0"` : `${heightFeet}'-${heightInches}"`

        const embedmentFeet = Math.floor(avgEmbedmentRounded)
        const embedmentInches = Math.round((avgEmbedmentRounded - embedmentFeet) * 12)
        const embedmentText = embedmentInches === 0 ? `${embedmentFeet}'-0"` : `${embedmentFeet}'-${embedmentInches}"`

        const proposalLine = diameter
            ? `F&I new (${Math.round(totalQty || totalTakeoff)})no [${diameter}" Ø] primary secant piles (Havg=${heightText}, ${embedmentText} embedment) as per SOE-#.## & details on SOE-#.##`
            : `F&I new (${Math.round(totalQty || totalTakeoff)})no [#" Ø] primary secant piles (Havg=${heightText}, ${embedmentText} embedment) as per SOE-#.## & details on SOE-#.##`
    }

    return items
}
export const processSecondarySecantItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isSecondarySecantPile, tracker)
export const processTangentPileItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isTangentPile, tracker)
export const processSheetPileItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericSoeItems(rawDataRows, headers, isSheetPile, tracker)

    if (items.length > 0) {
    }

    return items
}
export const processTimberLaggingItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isTimberLagging, tracker)
export const processTimberSheetingItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isTimberSheeting, tracker)
export const processWalerItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isWaler, tracker)
export const processRakerItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isRaker, tracker)
export const processUpperRakerItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isUpperRaker, tracker)
export const processLowerRakerItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isLowerRaker, tracker)
export const processStandOffItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isStandOff, tracker)
export const processKickerItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isKicker, tracker)
export const processChannelItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isChannel, tracker)
export const processRollChockItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isRollChock, tracker)
export const processStudBeamItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isStudBeam, tracker)
export const processInnerCornerBraceItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isCornerbrace, tracker)
export const processKneeBraceItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isKneeBrace, tracker)
export const processSupportingAngleItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericSoeItems(rawDataRows, headers, isSupportingAngle, tracker)
    // Group by groupKey (what's after @)
    const groups = new Map()
    items.forEach(item => {
        const key = item.parsed.groupKey || 'Other'
        if (!groups.has(key)) groups.set(key, [])
        groups.get(key).push(item)
    })
    return Array.from(groups.entries()).map(([name, items]) => ({ name, items }))
}
export const processPargingItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isParging, tracker)
export const processHeelBlockItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isHeelBlock, tracker)
export const processUnderpinningItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isUnderpinning, tracker)
export const processRockAnchorItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isRockAnchor, tracker)
export const processRockBoltItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericSoeItems(rawDataRows, headers, isRockBolt, tracker)

    // Group by groupKey (spacing value)
    const groupMap = new Map()
    items.forEach(item => {
        const groupKey = item.parsed?.groupKey || 'OTHER'
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    let groups = Array.from(groupMap.values())

    // Merge single-item groups if there are multiple
    const singleItemGroups = groups.filter(g => g.items.length === 1)
    const multiItemGroups = groups.filter(g => g.items.length > 1)

    if (singleItemGroups.length > 1) {
        const mergedItems = []
        singleItemGroups.forEach(group => {
            mergedItems.push(...group.items)
        })

        return [...multiItemGroups, {
            groupKey: 'MERGED',
            items: mergedItems,
            parsed: mergedItems[0].parsed,
            isMerged: true
        }]
    }

    return groups.length > 0 ? groups : items
}
export const processAnchorItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isAnchor, tracker)
export const processTieBackItems = (rawDataRows, headers, tracker = null) => processGenericSoeItems(rawDataRows, headers, isTieBack, tracker)
export const processConcreteSoilRetentionPierItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericSoeItems(rawDataRows, headers, isConcreteSoilRetentionPier, tracker)


    return items
}
export const processGuideWallItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericSoeItems(rawDataRows, headers, isGuideWall, tracker)

    // Generate proposal text template
    if (items.length > 0) {
        // Extract widths from items (e.g., "4'-6½" & 5'-3½"")
        const widths = new Set()
        let height = ''
        items.forEach(item => {
            const particulars = item.particulars || ''
            // Extract from bracket: Guide wall (4'-6½"x3'-0")
            const bracketMatch = particulars.match(/\(([^)]+)\)/)
            if (bracketMatch) {
                const dimsStr = bracketMatch[1].split('x').map(d => d.trim())
                if (dimsStr.length >= 1) {
                    widths.add(dimsStr[0]) // First dimension is width
                }
                if (dimsStr.length >= 2) {
                    height = dimsStr[1] // Second dimension is height
                }
            }
        })
        const widthText = Array.from(widths).join(' & ')

        // Get SOE page references
        let soePageMain = 'SOE-100.00'
        let soePageDetails = 'SOE-300.00'
        const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
        if (pageIdx !== -1) {
            for (let rowIndex = 0; rowIndex < rawDataRows.length; rowIndex++) {
                const row = rawDataRows[rowIndex]
                const digitizerItem = row[headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')]
                if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('guide wall')) {
                        const pageValue = row[pageIdx]
                        if (pageValue) {
                            const pageStr = String(pageValue).trim()
                            const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                            if (soeMatches && soeMatches.length > 0) {
                                soePageMain = soeMatches[0]
                                if (soeMatches.length > 1) {
                                    soePageDetails = soeMatches[1]
                                }
                            }
                        }
                    }
                }
            }
        }

        // Format: F&I new (4'-6½" & 5'-3½" wide) guide wall (H=3'-0") as per SOE-100.00 & details on SOE-300.00
        const proposalText = widthText
            ? `F&I new (${widthText} wide) guide wall (H=${height || '3\'-0"'}) as per ${soePageMain} & details on ${soePageDetails}`
            : `F&I new guide wall (H=${height || '3\'-0"'}) as per ${soePageMain} & details on ${soePageDetails}`
    }

    return items
}

export const processDowelBarItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericSoeItems(rawDataRows, headers, isDowelBar, tracker)

    // Generate proposal text template
    if (items.length > 0) {

        // Calculate totals
        const totalQty = items.reduce((sum, item) => sum + (parseFloat(item.qty) || 0), 0)
        const totalTakeoff = items.reduce((sum, item) => sum + (item.takeoff || 0), 0)
        const totalQtyValue = Math.round(totalQty || totalTakeoff || 0)

        // Extract bar size (e.g., #9)
        let barSize = ''
        let rockSocket = ''
        items.forEach(item => {
            const particulars = item.particulars || ''
            // Extract bar size: "4 - #9 Steel dowels bar"
            const barMatch = particulars.match(/#(\d+)/)
            if (barMatch && !barSize) {
                barSize = `#${barMatch[1]}`
            }
            // Extract rock socket: "RS=4'-0""
            const rsMatch = particulars.match(/RS=([0-9'"\-]+)/i)
            if (rsMatch && !rockSocket) {
                rockSocket = rsMatch[1]
            }
        })

        // Calculate average height
        let totalHeight = 0
        let heightCount = 0
        items.forEach(item => {
            const height = item.parsed?.heightRaw || 0
            if (height > 0) {
                totalHeight += height * (item.takeoff || 0)
                heightCount += (item.takeoff || 0)
            }
        })
        const avgHeight = heightCount > 0 ? totalHeight / heightCount : 0
        const heightFeet = Math.floor(avgHeight)
        const heightInches = Math.round((avgHeight - heightFeet) * 12)
        const heightText = heightInches === 0 ? `${heightFeet}'-0"` : `${heightFeet}'-${heightInches}"`

        // Get SOE page references
        let soePageMain = 'SOESK-01'
        let soePageDetails = 'SOESK-02'
        const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
        if (pageIdx !== -1) {
            for (let rowIndex = 0; rowIndex < rawDataRows.length; rowIndex++) {
                const row = rawDataRows[rowIndex]
                const digitizerItem = row[headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')]
                if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('dowel')) {
                        const pageValue = row[pageIdx]
                        if (pageValue) {
                            const pageStr = String(pageValue).trim()
                            const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                            if (soeMatches && soeMatches.length > 0) {
                                soePageMain = soeMatches[0]
                                if (soeMatches.length > 1) {
                                    soePageDetails = soeMatches[1]
                                }
                            }
                        }
                    }
                }
            }
        }

        // Format: F&I new (48)no 4-#9 steel dowels bar (Havg=6'-3", 4'-0" rock socket) as per SOESK-01 & details on SOESK-02
        const proposalText = `F&I new (${totalQtyValue})no 4-${barSize || '#9'} steel dowels bar (Havg=${heightText}, ${rockSocket || '4\'-0"'} rock socket) as per ${soePageMain} & details on ${soePageDetails}`
    }

    return items
}

export const processRockPinItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericSoeItems(rawDataRows, headers, isRockPin, tracker)

    // Generate proposal text template
    if (items.length > 0) {

        // Calculate totals
        const totalQty = items.reduce((sum, item) => sum + (parseFloat(item.qty) || 0), 0)
        const totalTakeoff = items.reduce((sum, item) => sum + (item.takeoff || 0), 0)
        const totalQtyValue = Math.round(totalQty || totalTakeoff || 0)

        // Extract rock socket
        let rockSocket = ''
        items.forEach(item => {
            const particulars = item.particulars || ''
            // Extract rock socket: "RS=4'-0""
            const rsMatch = particulars.match(/RS=([0-9'"\-]+)/i)
            if (rsMatch && !rockSocket) {
                rockSocket = rsMatch[1]
            }
        })

        // Calculate average height
        let totalHeight = 0
        let heightCount = 0
        items.forEach(item => {
            const height = item.parsed?.heightRaw || 0
            if (height > 0) {
                totalHeight += height * (item.takeoff || 0)
                heightCount += (item.takeoff || 0)
            }
        })
        const avgHeight = heightCount > 0 ? totalHeight / heightCount : 0
        const heightFeet = Math.floor(avgHeight)
        const heightInches = Math.round((avgHeight - heightFeet) * 12)
        const heightText = heightInches === 0 ? `${heightFeet}'-0"` : `${heightFeet}'-${heightInches}"`

        // Get SOE page references
        let soePageMain = 'SOESK-01'
        let soePageDetails = 'SOESK-02'
        const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
        if (pageIdx !== -1) {
            for (let rowIndex = 0; rowIndex < rawDataRows.length; rowIndex++) {
                const row = rawDataRows[rowIndex]
                const digitizerItem = row[headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')]
                if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('rock pin')) {
                        const pageValue = row[pageIdx]
                        if (pageValue) {
                            const pageStr = String(pageValue).trim()
                            const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                            if (soeMatches && soeMatches.length > 0) {
                                soePageMain = soeMatches[0]
                                if (soeMatches.length > 1) {
                                    soePageDetails = soeMatches[1]
                                }
                            }
                        }
                    }
                }
            }
        }

        // Format: F&I new (24)no rock pins (Havg=6'-3", 4'-0" rock socket) as per SOESK-01 & details on SOESK-02
        const proposalText = `F&I new (${totalQtyValue})no rock pins (Havg=${heightText}, ${rockSocket || '4\'-0"'} rock socket) as per ${soePageMain} & details on ${soePageDetails}`
    }

    return items
}

export const processShotcreteItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericSoeItems(rawDataRows, headers, isShotcrete, tracker)

    // Group by groupKey (H value)
    const groupMap = new Map()
    items.forEach(item => {
        const groupKey = item.parsed?.groupKey || 'OTHER'
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    let groups = Array.from(groupMap.values())

    // Merge single-item groups if there are multiple
    const singleItemGroups = groups.filter(g => g.items.length === 1)
    const multiItemGroups = groups.filter(g => g.items.length > 1)

    if (singleItemGroups.length > 1) {
        const mergedItems = []
        singleItemGroups.forEach(group => {
            mergedItems.push(...group.items)
        })

        groups = [...multiItemGroups, {
            groupKey: 'MERGED',
            items: mergedItems,
            parsed: mergedItems[0].parsed,
            isMerged: true
        }]
    }

    const allItems = groups.length > 0 ? groups.flatMap(g => g.items) : items

    // Generate proposal text template
    if (allItems.length > 0) {

        // Extract thickness and wire mesh
        let thickness = ''
        let wireMesh = ''
        allItems.forEach(item => {
            const particulars = item.particulars || ''
            // Extract thickness: "6" thick"
            const thickMatch = particulars.match(/(\d+(?:\/\d+)?)"?\s*thick/i)
            if (thickMatch && !thickness) {
                thickness = `${thickMatch[1]}"`
            }
            // Extract wire mesh: "6x6 wire mesh"
            const meshMatch = particulars.match(/(\d+x\d+)\s*wire\s*mesh/i)
            if (meshMatch && !wireMesh) {
                wireMesh = meshMatch[1]
            }
        })

        // Calculate average height
        let totalHeight = 0
        let heightCount = 0
        items.forEach(item => {
            const height = item.parsed?.heightRaw || 0
            if (height > 0) {
                totalHeight += height * (item.takeoff || 0)
                heightCount += (item.takeoff || 0)
            }
        })
        const avgHeight = heightCount > 0 ? totalHeight / heightCount : 0
        const heightFeet = Math.floor(avgHeight)
        const heightInches = Math.round((avgHeight - heightFeet) * 12)
        const heightText = heightInches === 0 ? `${heightFeet}'-0"` : `${heightFeet}'-${heightInches}"`

        // Get SOE page references
        let soePageMain = 'SOESK-01'
        let soePageDetails = 'SOESK-02'
        const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
        if (pageIdx !== -1) {
            for (let rowIndex = 0; rowIndex < rawDataRows.length; rowIndex++) {
                const row = rawDataRows[rowIndex]
                const digitizerItem = row[headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')]
                if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('shotcrete')) {
                        const pageValue = row[pageIdx]
                        if (pageValue) {
                            const pageStr = String(pageValue).trim()
                            const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                            if (soeMatches && soeMatches.length > 0) {
                                soePageMain = soeMatches[0]
                                if (soeMatches.length > 1) {
                                    soePageDetails = soeMatches[1]
                                }
                            }
                        }
                    }
                }
            }
        }

        // Format: F&I new (6" thick) shotcrete w/ 6x6 wire mesh (Havg=23'-0") as per SOESK-01 & details on SOESK-02
        const proposalText = `F&I new (${thickness || '6"'} thick) shotcrete w/ ${wireMesh || '6x6'} wire mesh (Havg=${heightText}) as per ${soePageMain} & details on ${soePageDetails}`
    }

    return items
}

export const processPermissionGroutingItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericSoeItems(rawDataRows, headers, isPermissionGrouting, tracker)

    // Group by groupKey (H value)
    const groupMap = new Map()
    items.forEach(item => {
        const groupKey = item.parsed?.groupKey || 'OTHER'
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    let groups = Array.from(groupMap.values())

    // Merge single-item groups if there are multiple
    const singleItemGroups = groups.filter(g => g.items.length === 1)
    const multiItemGroups = groups.filter(g => g.items.length > 1)

    if (singleItemGroups.length > 1) {
        const mergedItems = []
        singleItemGroups.forEach(group => {
            mergedItems.push(...group.items)
        })

        groups = [...multiItemGroups, {
            groupKey: 'MERGED',
            items: mergedItems,
            parsed: mergedItems[0].parsed,
            isMerged: true
        }]
    }

    const allItems = groups.length > 0 ? groups.flatMap(g => g.items) : items

    // Generate proposal text template
    if (allItems.length > 0) {

        // Calculate average height
        let totalHeight = 0
        let heightCount = 0
        allItems.forEach(item => {
            const height = item.parsed?.heightRaw || 0
            if (height > 0) {
                totalHeight += height * (item.takeoff || 0)
                heightCount += (item.takeoff || 0)
            }
        })
        const avgHeight = heightCount > 0 ? totalHeight / heightCount : 0
        const heightFeet = Math.floor(avgHeight)
        const heightInches = Math.round((avgHeight - heightFeet) * 12)
        const heightText = heightInches === 0 ? `${heightFeet}'-0"` : `${heightFeet}'-${heightInches}"`

        // Get SOE page references
        let soePageMain = 'SOESK-01'
        let soePageDetails = 'SOESK-02'
        const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
        if (pageIdx !== -1) {
            for (let rowIndex = 0; rowIndex < rawDataRows.length; rowIndex++) {
                const row = rawDataRows[rowIndex]
                const digitizerItem = row[headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')]
                if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('permission grouting')) {
                        const pageValue = row[pageIdx]
                        if (pageValue) {
                            const pageStr = String(pageValue).trim()
                            const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                            if (soeMatches && soeMatches.length > 0) {
                                soePageMain = soeMatches[0]
                                if (soeMatches.length > 1) {
                                    soePageDetails = soeMatches[1]
                                }
                            }
                        }
                    }
                }
            }
        }

        // Format: F&I new permission grouting (Havg=16'-0") as per SOESK-01 & details on SOESK-02
        const proposalText = `F&I new permission grouting (Havg=${heightText}) as per ${soePageMain} & details on ${soePageDetails}`
    }

    return items
}
export const processButtonItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericSoeItems(rawDataRows, headers, isButton, tracker)

    // Generate proposal text template
    if (items.length > 0) {

        // Calculate totals for proposal text
        const totalQty = items.reduce((sum, item) => sum + (parseFloat(item.qty) || 0), 0)
        const totalTakeoff = items.reduce((sum, item) => sum + (item.takeoff || 0), 0)
        const totalQtyValue = Math.round(totalQty || totalTakeoff || 0)

        // Extract width from items (for format: 3'-0"x3'-0" wide)
        let widthText = ''
        if (items.length > 0) {
            const firstItem = items[0]
            const particulars = firstItem.particulars || ''

            // Extract dimensions from bracket format: Concrete button (3'-0"x3'-0"x1'-0")
            const bracketMatch = particulars.match(/\(([^)]+)\)/)
            if (bracketMatch) {
                const dimsStr = bracketMatch[1].split('x').map(d => d.trim())
                if (dimsStr.length >= 2) {
                    // Use first two dimensions (length x width)
                    widthText = `${dimsStr[0]}x${dimsStr[1]}`
                }
            } else {
                // Fallback to parsed values if bracket not found
                const length = firstItem.parsed?.length
                const width = firstItem.parsed?.width
                if (length && width) {
                    // Format numeric values back to dimension strings
                    const formatDim = (val) => {
                        if (typeof val === 'number') {
                            const feet = Math.floor(val)
                            const inches = Math.round((val - feet) * 12)
                            return inches === 0 ? `${feet}'-0"` : `${feet}'-${inches}"`
                        }
                        return val
                    }
                    widthText = `${formatDim(length)}x${formatDim(width)}`
                }
            }
        }

        // Calculate average height from parsed height values
        let totalHeight = 0
        let heightCount = 0
        items.forEach(item => {
            // For buttons, height is in parsed.height (from bracket dimensions)
            const height = item.parsed?.height || item.parsed?.heightRaw || 0
            if (height > 0) {
                totalHeight += height * (item.takeoff || 0)
                heightCount += (item.takeoff || 0)
            }
        })
        const avgHeight = heightCount > 0 ? totalHeight / heightCount : 0

        // Format average height (use actual average, format as feet and inches)
        const heightFeet = Math.floor(avgHeight)
        const heightInches = Math.round((avgHeight - heightFeet) * 12)
        const heightText = heightInches === 0 ? `${heightFeet}'-0"` : `${heightFeet}'-${heightInches}"`

        // Get SOE page reference from raw data (could be SOESK-01 or SOE-XXX.XX)
        let soePageMain = 'SOE-101.00' // Default
        const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
        if (pageIdx !== -1) {
            for (let rowIndex = 0; rowIndex < rawDataRows.length; rowIndex++) {
                const row = rawDataRows[rowIndex]
                const digitizerItem = row[headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')]
                if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('button')) {
                        const pageValue = row[pageIdx]
                        if (pageValue) {
                            const pageStr = String(pageValue).trim()
                            // Match SOE-XXX.XX or SOESK-XX format
                            const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                            if (soeMatches && soeMatches.length > 0) {
                                soePageMain = soeMatches[0]
                                break
                            }
                        }
                    }
                }
            }
        }

        // Generate proposal text template
        // Format: F&I new (12)no (3'-0"x3'-0" wide) concrete buttons (Havg=3'-3") as per SOESK-01
        const proposalText = widthText
            ? `F&I new (${totalQtyValue})no (${widthText} wide) concrete buttons (Havg=${heightText}) as per ${soePageMain}`
            : `F&I new (${totalQtyValue})no concrete buttons (Havg=${heightText}) as per ${soePageMain}`
    }

    return items
}
export const processRockStabilizationItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericSoeItems(rawDataRows, headers, isRockStabilization, tracker)

    // Group by groupKey (H value)
    const groupMap = new Map()
    items.forEach(item => {
        const groupKey = item.parsed?.groupKey || 'OTHER'
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    let groups = Array.from(groupMap.values())

    // Merge single-item groups if there are multiple
    const singleItemGroups = groups.filter(g => g.items.length === 1)
    const multiItemGroups = groups.filter(g => g.items.length > 1)

    if (singleItemGroups.length > 1) {
        const mergedItems = []
        singleItemGroups.forEach(group => {
            mergedItems.push(...group.items)
        })

        groups = [...multiItemGroups, {
            groupKey: 'MERGED',
            items: mergedItems,
            parsed: mergedItems[0].parsed,
            isMerged: true
        }]
    }

    const allItems = groups.length > 0 ? groups.flatMap(g => g.items) : items

    // Generate proposal text template
    if (allItems.length > 0) {

        // Get SOE page references
        let soePageMain = 'SOE-100.00'
        let soePageDetails = 'SOE-300.00'
        const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
        if (pageIdx !== -1) {
            for (let rowIndex = 0; rowIndex < rawDataRows.length; rowIndex++) {
                const row = rawDataRows[rowIndex]
                const digitizerItem = row[headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')]
                if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('rock stabilization')) {
                        const pageValue = row[pageIdx]
                        if (pageValue) {
                            const pageStr = String(pageValue).trim()
                            const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                            if (soeMatches && soeMatches.length > 0) {
                                soePageMain = soeMatches[0]
                                if (soeMatches.length > 1) {
                                    soePageDetails = soeMatches[1]
                                }
                            }
                        }
                    }
                }
            }
        }

        // Format: F&I new rock stabilization as per SOE-100.00 & details on SOE-300.00
        const proposalText = `F&I new rock stabilization as per ${soePageMain} & details on ${soePageDetails}`
    }

    return items
}
export const processFormBoardItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericSoeItems(rawDataRows, headers, isFormBoard, tracker)

    // Group by groupKey (thickness at start)
    const groupMap = new Map()
    items.forEach(item => {
        const groupKey = item.parsed?.groupKey || 'OTHER'
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    let groups = Array.from(groupMap.values())

    // Merge single-item groups if there are multiple
    const singleItemGroups = groups.filter(g => g.items.length === 1)
    const multiItemGroups = groups.filter(g => g.items.length > 1)

    if (singleItemGroups.length > 1) {
        const mergedItems = []
        singleItemGroups.forEach(group => {
            mergedItems.push(...group.items)
        })

        groups = [...multiItemGroups, {
            groupKey: 'MERGED',
            items: mergedItems,
            parsed: mergedItems[0].parsed,
            isMerged: true
        }]
    }

    const allItems = groups.length > 0 ? groups.flatMap(g => g.items) : items

    // Generate proposal text template
    if (allItems.length > 0) {

        // Calculate totals for proposal text
        const totalQty = allItems.reduce((sum, item) => sum + (parseFloat(item.qty) || 0), 0)
        const totalTakeoff = allItems.reduce((sum, item) => sum + (item.takeoff || 0), 0)
        const totalQtyValue = Math.round(totalQty || totalTakeoff || 0)

        // Extract thickness from items (e.g., "1" form board" -> "1"")
        let thickness = '1"' // Default thickness
        if (allItems.length > 0) {
            const firstItem = allItems[0]
            const particulars = (firstItem.particulars || '').trim()
            // Match pattern like "1" form board" or "1\" form board"
            const thicknessMatch = particulars.match(/(\d+(?:\/\d+)?)"?\s*form\s*board/i)
            if (thicknessMatch) {
                thickness = `${thicknessMatch[1]}"`
            }
        }

        // Get SOE page reference from raw data
        let soePageMain = 'SOE-300.00' // Default for form board
        const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
        if (pageIdx !== -1) {
            for (let rowIndex = 0; rowIndex < rawDataRows.length; rowIndex++) {
                const row = rawDataRows[rowIndex]
                const digitizerItem = row[headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')]
                if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('form board')) {
                        const pageValue = row[pageIdx]
                        if (pageValue) {
                            const pageStr = String(pageValue).trim()
                            const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                            if (soeMatches && soeMatches.length > 0) {
                                soePageMain = soeMatches[0]
                                break
                            }
                        }
                    }
                }
            }
        }

        // Generate proposal text template
        const proposalText = `F&I new (${thickness} thick) form board w/ filter fabric between tunnel and retention pier as per ${soePageMain}`
    }

    return items
}

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

    // Use parsed.type for formula logic (e.g. drilled_hole_grout_item -> parsed.type is drilled_hole_grout)
    const type = itemData.parsed?.type || (itemType === 'drilled_hole_grout_item' ? 'drilled_hole_grout' : itemType)

    switch (type) {
        case 'hp':
        case 'drilled':
        case 'soldier_pile':
            formulas.ft = `H${rowNum}*C${rowNum}`
            formulas.lbs = `I${rowNum}*${(itemData.weight || 0).toFixed(3)}`
            formulas.qtyFinal = `C${rowNum}`
            break
        case 'timber':
            // Timber soldier piles: no column K (LBS)
            formulas.ft = `H${rowNum}*C${rowNum}`
            formulas.qtyFinal = `C${rowNum}`
            break
        case 'timber_plank':
            // Timber planks: no column K. If qty from (N) at start: E=qty, I=H*C*E, M=C*E. Else: I=H*C, M=C
            if (itemData.parsed?.qty != null) {
                formulas.qty = itemData.parsed.qty
                formulas.ft = `H${rowNum}*C${rowNum}*E${rowNum}`
                formulas.qtyFinal = `C${rowNum}*E${rowNum}`
            } else {
                formulas.ft = `H${rowNum}*C${rowNum}`
                formulas.qtyFinal = `C${rowNum}`
            }
            break
        case 'timber_raker':
            // Timber raker: no column K (LBS), same as timber soldier piles
            formulas.ft = `H${rowNum}*C${rowNum}`
            formulas.qtyFinal = `C${rowNum}`
            break
        case 'timber_brace_horizontal':
            // 2"x4" Horizontal wood brace: I=C*H, M=C, height from bracket
            formulas.height = itemData.parsed?.calculatedHeight
            formulas.ft = `C${rowNum}*H${rowNum}`
            formulas.qtyFinal = `C${rowNum}`
            break
        case 'timber_brace_corner':
            // (2) 6"x6" Timber Corner brace: E=qty, F=manual, I=C*E, M=F*E
            formulas.qty = itemData.parsed?.qty ?? 1
            formulas.ft = `C${rowNum}*E${rowNum}`
            formulas.qtyFinal = `F${rowNum}*E${rowNum}`
            break
        case 'timber_waler':
            // (2) 6"x6" Timber waler: E=qty from (N) or 1, F=manual, I=E*C, M=E*F
            formulas.qty = itemData.parsed?.qty ?? 1
            formulas.ft = `E${rowNum}*C${rowNum}`
            formulas.qtyFinal = `E${rowNum}*F${rowNum}`
            break
        case 'timber_stringer':
            // (2) 6"x6" Timber stringer: E=qty from (N) or 1, I=C*E
            formulas.qty = itemData.parsed?.qty ?? 1
            formulas.ft = `C${rowNum}*E${rowNum}`
            break
        case 'drilled_hole_grout':
            // 5-5/8" Ø Drilled hole grout H=22'-6": F=G=SQRT((d/12)^2*3.14/4), H=height, I=C*H, J=G*F*C, L=J*H/27, M=C
            const diam = itemData.parsed?.diameter || 5.625
            const fgFormula = `SQRT((${diam}/12)^2*3.14/4)`
            formulas.length = fgFormula
            formulas.width = fgFormula
            formulas.height = itemData.parsed?.heightRaw
            formulas.ft = `C${rowNum}*H${rowNum}`
            formulas.sqFt = `G${rowNum}*F${rowNum}*C${rowNum}`
            formulas.cy = `J${rowNum}*H${rowNum}/27`
            formulas.qtyFinal = `C${rowNum}`
            break
        case 'timber_post':
            // Timber post: no column K (LBS), same as timber soldier piles
            formulas.ft = `H${rowNum}*C${rowNum}`
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'primary_secant':
        case 'tangent':
            // No column K (LBS) for Primary secant piles and Tangent piles
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
        case 'timber_sheets':
        case 'vertical_timber_sheets':
        case 'horizontal_timber_sheets':
            // Timber lagging/sheeting/timber sheets: FT(I)=C, SQ FT(J)=I*H
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

        case 'supporting_angle':
            // FT(I)=H*E*C, LBS(K)=I*Wt, QTY(M)=C*E
            formulas.ft = `H${rowNum}*E${rowNum}*C${rowNum}`
            formulas.lbs = `I${rowNum}*${(itemData.weight || 0).toFixed(3)}`
            formulas.qtyFinal = `C${rowNum}*E${rowNum}`
            break

        case 'parging':
            // FT(I)=C, SQ FT(J)=I*H
            formulas.ft = `C${rowNum}`
            formulas.sqFt = `I${rowNum}*H${rowNum}`
            break

        case 'heel_block':
            // SQ FT(J)=C*H*G, CY(L)=J*F/27, QTY(M)=C
            formulas.sqFt = `C${rowNum}*H${rowNum}*G${rowNum}`
            formulas.cy = `J${rowNum}*F${rowNum}/27`
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'underpinning':
            // FT(I)=F*C, SQ FT(J)=C*H*G, CY(L)=J*F/27, QTY(M)=C
            formulas.ft = `F${rowNum}*C${rowNum}`
            formulas.sqFt = `C${rowNum}*H${rowNum}*G${rowNum}`
            formulas.cy = `J${rowNum}*F${rowNum}/27`
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'shims':
            // FT(I)=Reference, SQ FT(J)=I*G
            // reference is passed in itemData
            if (itemData.underpinningSumRow) {
                formulas.ft = `I${itemData.underpinningSumRow}`
            }
            formulas.sqFt = `I${rowNum}*G${rowNum}`
            break

        case 'rock_anchor':
            // FT(I)=F*C, QTY(M)=C
            // F should be calculated height (sum of free length + bond length, rounded to multiple of 5, then +5)
            if (itemData.parsed?.calculatedHeight !== undefined) {
                formulas.length = itemData.parsed.calculatedHeight
            }
            formulas.ft = `F${rowNum}*C${rowNum}`
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'rock_bolt':
            // QTY(E)=ROUNDUP(C/OC,0)+1, Length(F)=bond length+5, FT(I)=F*E, QTY(M)=E
            // OC spacing and bond length are in parsed data
            if (itemData.parsed?.ocSpacing) {
                formulas.qty = `ROUNDUP(C${rowNum}/${itemData.parsed.ocSpacing},0)+1`
            }
            if (itemData.parsed?.bondLength !== undefined) {
                formulas.length = `${itemData.parsed.bondLength}+5`
            }
            formulas.ft = `F${rowNum}*E${rowNum}`
            formulas.qtyFinal = `E${rowNum}`
            break

        case 'anchor':
            // Height(H)=calculated height, FT(I)=H*C, QTY(M)=C
            if (itemData.parsed?.calculatedHeight) {
                formulas.height = itemData.parsed.calculatedHeight
            }
            formulas.ft = `H${rowNum}*C${rowNum}`
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'tie_back':
            // Height(H)=calculated height, FT(I)=H*C, QTY(M)=C
            if (itemData.parsed?.calculatedHeight) {
                formulas.height = itemData.parsed.calculatedHeight
            }
            formulas.ft = `H${rowNum}*C${rowNum}`
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'concrete_soil_retention_pier':
            // Length(F), Width(G), Height(H) from bracket, SQ FT(J)=C*H*G, CY(L)=J*F/27, QTY(M)=C
            if (itemData.parsed?.length) formulas.length = itemData.parsed.length
            if (itemData.parsed?.width) formulas.width = itemData.parsed.width
            if (itemData.parsed?.height) formulas.height = itemData.parsed.height
            formulas.sqFt = `C${rowNum}*H${rowNum}*G${rowNum}`
            formulas.cy = `J${rowNum}*F${rowNum}/27`
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'guide_wall':
            // Width(G)=formula from bracket (e.g., =4+(6.5/12)), Height(H)=from bracket, FT(I)=C, SQ FT(J)=I*G, CY(L)=J*H/27
            if (itemData.parsed?.widthFormula) {
                formulas.width = itemData.parsed.widthFormula
            } else if (itemData.parsed?.width) {
                formulas.width = itemData.parsed.width
            }
            if (itemData.parsed?.heightRaw) {
                formulas.height = itemData.parsed.heightRaw
            }
            formulas.ft = `C${rowNum}`
            formulas.sqFt = `I${rowNum}*G${rowNum}`
            formulas.cy = `J${rowNum}*H${rowNum}/27`
            break

        case 'dowel_bar':
            // QTY(E) from name, Height(H)=H+RS, FT(I)=C*E*H, QTY(M)=C*E
            if (itemData.parsed?.qty) formulas.qty = itemData.parsed.qty
            if (itemData.parsed?.heightRaw) formulas.height = itemData.parsed.heightRaw
            formulas.ft = `C${rowNum}*E${rowNum}*H${rowNum}`
            formulas.qtyFinal = `C${rowNum}*E${rowNum}`
            break

        case 'rock_pin':
            // QTY(E)=1, Height(H)=H+RS, FT(I)=C*E*H, QTY(M)=C*E
            formulas.qty = 1
            if (itemData.parsed?.heightRaw) formulas.height = itemData.parsed.heightRaw
            formulas.ft = `C${rowNum}*E${rowNum}*H${rowNum}`
            formulas.qtyFinal = `C${rowNum}*E${rowNum}`
            break

        case 'shotcrete':
            // Length(F) and Width(G) should be empty, Height(H) from name, FT(I)=C, SQ FT(J)=C*H, CY(L)=J*G/27
            // Don't set formulas.length or formulas.width - they should be empty
            if (itemData.parsed?.heightRaw) formulas.height = itemData.parsed.heightRaw
            formulas.ft = `C${rowNum}`
            formulas.sqFt = `C${rowNum}*H${rowNum}`
            formulas.cy = `J${rowNum}*G${rowNum}/27`
            break

        case 'permission_grouting':
            // Height(H) from name, FT(I)=C, SQ FT(J)=C*H
            if (itemData.parsed?.heightRaw) formulas.height = itemData.parsed.heightRaw
            formulas.ft = `C${rowNum}`
            formulas.sqFt = `C${rowNum}*H${rowNum}`
            break

        case 'button':
            // Length(F), Width(G), Height(H) from bracket, SQ FT(J)=C*H*G, CY(L)=J*F/27, QTY(M)=C
            if (itemData.parsed?.length) formulas.length = itemData.parsed.length
            if (itemData.parsed?.width) formulas.width = itemData.parsed.width
            if (itemData.parsed?.height) formulas.height = itemData.parsed.height
            formulas.sqFt = `C${rowNum}*H${rowNum}*G${rowNum}`
            formulas.cy = `J${rowNum}*F${rowNum}/27`
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'rock_stabilization':
            // Height(H) from name, FT(I) should be empty, SQ FT(J)=C, CY(L)=J*H/27
            if (itemData.parsed?.heightRaw) formulas.height = itemData.parsed.heightRaw
            // Don't set formulas.ft - it should be empty
            formulas.sqFt = `C${rowNum}`
            formulas.cy = `J${rowNum}*H${rowNum}/27`
            break

        case 'form_board':
            // Height(H) from name, FT(I)=C, SQ FT(J)=C*H
            if (itemData.parsed?.heightRaw) formulas.height = itemData.parsed.heightRaw
            formulas.ft = `C${rowNum}`
            formulas.sqFt = `C${rowNum}*H${rowNum}`
            break
    }

    return formulas
}

/**
 * Formats drilled soldier pile proposal text
 * @param {Array} drilledGroups - Array of drilled soldier pile groups
 * @returns {string|null} - Formatted proposal text or null if no groups
 */
export const formatDrilledSoldierPileProposalText = (drilledGroups) => {
    if (!drilledGroups || drilledGroups.length === 0) return null

    // Get all drilled items from all groups
    const allDrilledItems = []
    drilledGroups.forEach(group => {
        if (group.type === 'drilled' && group.items) {
            allDrilledItems.push(...group.items)
        }
    })

    if (allDrilledItems.length === 0) return null

    // Get diameter and thickness from first item (should be same for all in a group)
    const firstItem = allDrilledItems[0]
    const parsed = firstItem.parsed
    const diameter = parsed.diameter
    const thickness = parsed.thickness

    // Calculate average height
    let totalHeight = 0
    let heightCount = 0
    let embedment = null

    allDrilledItems.forEach(item => {
        if (item.parsed && item.parsed.heightRaw) {
            totalHeight += item.parsed.heightRaw
            heightCount++
        }
        // Get embedment from first item that has it
        if (!embedment && item.parsed && item.parsed.embedment) {
            embedment = item.parsed.embedment
        }
    })

    if (heightCount === 0) return null

    const avgHeight = totalHeight / heightCount
    // Round average height to nearest foot
    const avgHeightRounded = Math.round(avgHeight)

    // Format embedment (embedment is already in feet from parseDimension)
    let embedmentText = ''
    if (embedment) {
        const embedmentFeet = Math.floor(embedment)
        const embedmentInches = Math.round((embedment - embedmentFeet) * 12)
        if (embedmentInches === 0) {
            embedmentText = `${embedmentFeet}'-0"`
        } else {
            embedmentText = `${embedmentFeet}'-${embedmentInches}"`
        }
    } else {
        embedmentText = "0'-0\""
    }

    // Format height (heightRaw is already in feet from parseDimension)
    const heightFeet = Math.floor(avgHeightRounded)
    const heightInches = Math.round((avgHeightRounded - heightFeet) * 12)
    let heightText = ''
    if (heightInches === 0) {
        heightText = `${heightFeet}'-0"`
    } else {
        heightText = `${heightFeet}'-${heightInches}"`
    }

    // Count total number of items (sum of all takeoff values)
    const totalCount = Math.round(allDrilledItems.reduce((sum, item) => sum + (item.takeoff || 0), 0))

    // Format the text: F&I new (##)no [9.625" Øx0.545" thick] drilled soldier piles (H=30'-0", 15'-0" embedment) as per SOE-101.00
    const proposalText = `F&I new (${totalCount})no [${diameter}" Øx${thickness}" thick] drilled soldier piles (H=${heightText}, ${embedmentText} embedment) as per SOE-101.00`

    return proposalText
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
    processSupportingAngleItems,
    processPargingItems,
    processHeelBlockItems,
    processUnderpinningItems,
    processRockAnchorItems,
    processRockBoltItems,
    processAnchorItems,
    processTieBackItems,
    processConcreteSoilRetentionPierItems,
    processGuideWallItems,
    processDowelBarItems,
    processRockPinItems,
    processShotcreteItems,
    processPermissionGroutingItems,
    processButtonItems,
    processRockStabilizationItems,
    processFormBoardItems,
    generateSoeFormulas,
    formatDrilledSoldierPileProposalText
}
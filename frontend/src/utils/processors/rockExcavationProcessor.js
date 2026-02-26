import { parseExcavationItem, getExcavationItemType } from '../parsers/excavationParser'
import { normalizeUnit } from '../parsers/dimensionParser'

/**
 * Identifies if a digitizer item belongs to Rock Excavation section
 * @param {string} digitizerItem - The digitizer item text
 * @returns {boolean} - True if it's a rock excavation item
 */
export const isRockExcavationItem = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false

    const itemLower = digitizerItem.toLowerCase()

    // Rock excavation specific keywords (excavation, exc, or exc.)
    const rockKeywords = [
        'duplex sewage ejector pit slab',
        'rock excavation',
        'rock exc',
        'rock exc.',
        'line drill'
    ]

    return rockKeywords.some(keyword => itemLower.includes(keyword))
}

/**
 * Generates formulas for rock excavation items
 * @param {string} itemType - Type of rock excavation item
 * @param {number} rowNum - Excel row number (1-based)
 * @param {object} parsedData - Parsed data
 * @returns {object} - Formulas
 */
export const generateRockExcavationFormulas = (itemType, rowNum, parsedData) => {
    const formulas = {
        ft: null,
        sqFt: null,
        lbs: null,
        cy: null,
        qtyFinal: null
    }

    switch (itemType) {
        case 'concrete_pier':
            // SQ FT (J) = C * F * G, CY (L) = J * H / 27
            formulas.sqFt = `C${rowNum}*F${rowNum}*G${rowNum}`
            formulas.cy = `J${rowNum}*H${rowNum}/27`
            formulas.lbs = null // Column K empty
            break

        case 'sewage_pit_slab':
        case 'rock_exc':
            // SQ FT (J) = C, CY (L) = J * H / 27
            formulas.sqFt = `C${rowNum}`
            formulas.cy = `J${rowNum}*H${rowNum}/27`
            formulas.lbs = null // Column K empty
            break

        case 'sump_pit':
            // SQ FT (J) = 16 * C, CY (L) = 1.3 * C
            formulas.sqFt = `16*C${rowNum}`
            formulas.cy = `1.3*C${rowNum}`
            formulas.lbs = null // Column K empty
            break

        case 'line_drill_concrete_pier':
            // takeoff = =((G+F)*2)*C, height = =H, col E is =ROUNDUP(H/2,0), col I is =E*C
            // These will be generated in generateCalculationSheet using external row references
            break

        case 'line_drilling':
            // col E is =ROUNDUP(H/2,0), col I is =E*C
            formulas.qty = `ROUNDUP(H${rowNum}/2,0)` // Col E
            formulas.ft = `E${rowNum}*C${rowNum}` // Col I
            break

        default:
            break
    }

    return formulas
}

/**
 * Processes all rock excavation items
 * @param {Array} rawDataRows - Raw data rows
 * @param {Array} headers - Column headers
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 * @returns {Array} - Processed items
 */
export const processRockExcavationItems = (rawDataRows, headers, tracker = null) => {
    const rockExcavationItems = []

    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return rockExcavationItems

    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = normalizeUnit(row[unitIdx] || '')

        if (isRockExcavationItem(digitizerItem)) {
            const itemLower = digitizerItem.toLowerCase()
            // Exclude line drill items from primary excavation subsection
            if (itemLower.includes('line drill')) return

            const itemType = getExcavationItemType(digitizerItem, 'rock_excavation')
            const parsed = parseExcavationItem(digitizerItem, total, unit, itemType, 'rock_excavation')

            const itemData = {
                ...parsed,
                id: `rock_exc_${itemType}_${rockExcavationItems.length}`,
                itemType,
                subsection: 'rock_excavation',
                rawRow: row,
                rawRowNumber: rowIndex + 2
            }

            // Aggregate sewage_pit_slab items
            if (itemType === 'sewage_pit_slab') {
                const existingItem = rockExcavationItems.find(item =>
                    item.itemType === 'sewage_pit_slab' &&
                    item.particulars === digitizerItem
                )
                if (existingItem) {
                    existingItem.takeoff += total
                    // Mark this row as used even though we're aggregating
                    if (tracker) {
                        tracker.markUsed(rowIndex)
                    }
                    return // Skip adding new item
                }
            }

            // Mark this row as used
            if (tracker) {
                tracker.markUsed(rowIndex)
            }

            rockExcavationItems.push(itemData)
        }
    })

    // Add "Sump pit" item - takeoff from "Sump pit @ elevator pit" raw data, or 0 if not available
    let sumpPitTakeoff = 0
    rawDataRows.forEach((row) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const itemLower = (digitizerItem || '').toLowerCase().trim()
        if (itemLower.includes('sump pit @ elevator pit')) {
            sumpPitTakeoff += total
        }
    })

    rockExcavationItems.push({
        id: 'rock_exc_sump_pit_manual',
        particulars: 'Sump pit',
        takeoff: sumpPitTakeoff,
        unit: 'EA',
        qty: 0,
        length: 0,
        width: 0,
        height: 0,
        itemType: 'sump_pit',
        subsection: 'rock_excavation',
        manualHeight: false
    })

    return rockExcavationItems
}

/**
 * Processes line drilling items from raw data
 * @param {Array} rawDataRows - Raw data rows
 * @param {Array} headers - Column headers
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 * @returns {Array} - Processed line drilling items
 */
export const processLineDrillItems = (rawDataRows, headers, tracker = null) => {
    const lineDrillItems = []

    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return lineDrillItems

    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = normalizeUnit(row[unitIdx] || '')

        const itemLower = digitizerItem?.toLowerCase() || ''
        if (itemLower.includes('line drill')) {
            const itemType = 'line_drilling'
            const parsed = parseExcavationItem(digitizerItem, total, unit, itemType, 'rock_excavation')

            lineDrillItems.push({
                ...parsed,
                itemType,
                subsection: 'line_drill',
                rawRow: row,
                rawRowNumber: rowIndex + 2
            })

            // Mark this row as used
            if (tracker) {
                tracker.markUsed(rowIndex)
            }
        }
    })

    return lineDrillItems
}

/**
 * Calculates total SQFT and CY from processed rock excavation items
 * @param {Array} rockExcavationItems - Processed rock excavation items
 * @returns {object} - Object with totalSQFT and totalCY
 */
export const calculateRockExcavationTotals = (rockExcavationItems) => {
    let totalSQFT = 0
    let totalCY = 0

    rockExcavationItems.forEach((item) => {
        const { itemType, takeoff, length, width, height } = item
        let sqft = 0
        let cy = 0

        switch (itemType) {
            case 'concrete_pier':
                // SQ FT (J) = C * F * G, CY (L) = J * H / 27
                // C = takeoff, F = length, G = width, H = height
                sqft = (takeoff || 0) * (length || 0) * (width || 0)
                cy = sqft * (height || 0) / 27
                break

            case 'sewage_pit_slab':
            case 'rock_exc':
                // SQ FT (J) = C, CY (L) = J * H / 27
                // C = takeoff, H = height
                sqft = takeoff || 0
                cy = sqft * (height || 0) / 27
                break

            case 'sump_pit':
                // SQ FT (J) = 16 * C, CY (L) = 1.3 * C
                // C = takeoff
                sqft = 16 * (takeoff || 0)
                cy = 1.3 * (takeoff || 0)
                break

            default:
                break
        }

        totalSQFT += sqft
        totalCY += cy
    })

    return {
        totalSQFT: parseFloat(totalSQFT.toFixed(2)),
        totalCY: parseFloat(totalCY.toFixed(2))
    }
}

/**
 * Calculates total FT from processed line drilling items
 * Formula: FT = ROUNDUP(height/2, 0) * takeoff
 * @param {Array} lineDrillItems - Processed line drilling items
 * @returns {number} - Total FT
 */
export const calculateLineDrillTotalFT = (lineDrillItems) => {
    let totalFT = 0

    lineDrillItems.forEach((item) => {
        const { takeoff, height } = item
        // FT = ROUNDUP(height/2, 0) * takeoff
        const qty = Math.ceil((height || 0) / 2)
        const ft = qty * (takeoff || 0)
        totalFT += ft
    })

    const result = parseFloat(totalFT.toFixed(2))


    return result
}

export default {
    isRockExcavationItem,
    generateRockExcavationFormulas,
    processRockExcavationItems,
    processLineDrillItems,
    calculateRockExcavationTotals,
    calculateLineDrillTotalFT
}

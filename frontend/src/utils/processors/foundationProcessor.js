import { mergeSingleItemGroupsIfAll } from '../groupingUtils.js'
import {
    isDrilledFoundationPile,
    isHelicalFoundationPile,
    isDrivenFoundationPile,
    isStelcorDrilledDisplacementPile,
    isCFAPile,
    isMiscellaneousFoundationPile,
    isPileCap,
    isStripFooting,
    isIsolatedFooting,
    isPilaster,
    isGradeBeam,
    isTieBeam,
    isThickenedSlab,
    isButtress,
    isPier,
    isCorbel,
    isLinearWall,
    isFoundationWall,
    isRetainingWall,
    isBarrierWall,
    isStemWall,
    isElevatorPit,
    isServiceElevatorPit,
    isDetentionTank,
    isDuplexSewageEjectorPit,
    isDeepSewageEjectorPit,
    isGreaseTrap,
    isHouseTrap,
    isMatSlab,
    isMudSlabFoundation,
    isSOG,
    isStairsOnGrade,
    isElectricConduit,
    parseDrilledFoundationPile,
    parseHelicalFoundationPile,
    parseDrivenFoundationPile,
    parseStelcorDrilledDisplacementPile,
    parseCFAPile,
    parsePileCap,
    parseStripFooting,
    parseIsolatedFooting,
    parsePilaster,
    parseGradeBeam,
    parseTieBeam,
    parseThickenedSlab,
    parsePier,
    parseCorbel,
    parseLinearWall,
    parseFoundationWall,
    parseRetainingWall,
    parseBarrierWall,
    parseStemWall,
    parseElevatorPit,
    parseServiceElevatorPit,
    parseDetentionTank,
    parseDuplexSewageEjectorPit,
    parseDeepSewageEjectorPit,
    parseGreaseTrap,
    parseHouseTrap,
    parseMatSlab,
    parseMudSlabFoundation,
    parseSOG,
    parseStairsOnGrade,
    parseElectricConduit
} from '../parsers/foundationParser'

/**
 * Generic process function for Foundation items
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 */
const processGenericFoundationItems = (rawDataRows, headers, identifierFn, parserFn, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const particularsIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'particulars')
    const itemNameIdx = digitizerIdx >= 0 ? digitizerIdx : particularsIdx
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (itemNameIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const items = []
    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[itemNameIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = row[unitIdx]

        if (identifierFn(digitizerItem)) {
            const parsed = parserFn ? parserFn(digitizerItem) : {}

            items.push({
                particulars: digitizerItem,
                takeoff: total,
                unit: unit,
                parsed: parsed,
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

/**
 * Processes drilled foundation pile items and groups them
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 */
export const processDrilledFoundationPileItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')
    // Find column that might contain "Influ" - check common column names
    const influIdx = headers.findIndex(h => {
        if (!h) return false
        const hLower = h.toLowerCase().trim()
        return hLower.includes('influ') || hLower.includes('influence') || hLower.includes('note')
    })

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const groups = []
    let currentGroup = null

    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const digitizerText = digitizerItem ? String(digitizerItem).trim() : ''
        const isRowEmpty = !digitizerItem || digitizerText === ''

        // Also check if this might be a sum row (has formula indicators or is just totals)
        // Sum rows typically have empty digitizer item but might have values in other columns
        const takeoffValue = row[totalIdx]
        const isSumRow = isRowEmpty && (takeoffValue === '' || takeoffValue === null || takeoffValue === undefined)

        if (isRowEmpty || isSumRow) {
            // Empty row or sum row - save current group if it exists and start a new one
            // This is how we separate groups - by empty rows, NOT by Influ
            if (currentGroup && currentGroup.items.length > 0) {
                groups.push(currentGroup)
                currentGroup = null
            }
            return
        }

        // Check if this is a drilled foundation pile
        if (isDrilledFoundationPile(digitizerItem)) {
            const total = parseFloat(row[totalIdx]) || 0
            const unit = row[unitIdx]
            const influValue = influIdx !== -1 ? row[influIdx] : null
            const hasInflu = influValue && String(influValue).toLowerCase().includes('influ')
            const parsed = parseDrilledFoundationPile(digitizerItem)

            // Mark this row as used
            if (tracker) {
                tracker.markUsed(rowIndex)
            }

            // If no current group, start a new one (new group starts when we encounter a drilled pile after an empty row)
            if (!currentGroup) {
                currentGroup = {
                    groupKey: parsed.groupKey || 'OTHER',
                    items: [{
                        particulars: digitizerItem,
                        takeoff: 0,
                        unit: unit,
                        parsed: parsed,
                        rawRowNumber: rowIndex + 2,
                        hasInflu: hasInflu
                    }],
                    parsed: parsed,
                    hasInflu: hasInflu
                }
            } else {
                // Add to current group - sum the takeoff
                // This item belongs to the same group (no empty row between them)
                currentGroup.items[0].takeoff += total
                // If this item has Influ, mark the whole group
                if (hasInflu) {
                    currentGroup.hasInflu = true
                }
            }
        } else {
            // Not a drilled foundation pile - if we have a current group, save it
            // This handles the case where we encounter a different item type, which also ends the group
            if (currentGroup && currentGroup.items.length > 0) {
                groups.push(currentGroup)
                currentGroup = null
            }
        }
    })

    // Don't forget to add the last group if it exists
    if (currentGroup && currentGroup.items.length > 0) {
        groups.push(currentGroup)
    }

    return groups
}

/**
 * Processes helical foundation pile items and groups them by diameter x thickness
 */
export const processHelicalFoundationPileItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericFoundationItems(rawDataRows, headers, isHelicalFoundationPile, parseHelicalFoundationPile, tracker)

    // Group by groupKey (diameter x thickness)
    const groupMap = new Map()
    items.forEach(item => {
        const groupKey = item.parsed.groupKey || 'OTHER'
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    return mergeSingleItemGroupsIfAll(Array.from(groupMap.values()))
}

/**
 * Processes driven foundation pile items
 */
export const processDrivenFoundationPileItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isDrivenFoundationPile, parseDrivenFoundationPile, tracker)
}

/**
 * Processes stelcor drilled displacement pile items
 */
export const processStelcorDrilledDisplacementPileItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isStelcorDrilledDisplacementPile, parseStelcorDrilledDisplacementPile, tracker)
}

/**
 * Processes CFA pile items
 */
export const processCFAPileItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isCFAPile, parseCFAPile, tracker)
}

/**
 * Processes miscellaneous pile items (pile items that don't fit other foundation or SOE subsections).
 * Run after other foundation pile processors; only processes rows not yet used.
 */
export const processMiscellaneousPileItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const items = []
    rawDataRows.forEach((row, rowIndex) => {
        if (tracker && tracker.isUsed(rowIndex)) return
        const digitizerItem = row[digitizerIdx]
        const total = row[totalIdx]
        const takeoff = total !== '' && total !== null && total !== undefined ? parseFloat(total) : ''
        const unit = unitIdx >= 0 ? (row[unitIdx] || '') : ''

        if (isMiscellaneousFoundationPile(digitizerItem)) {
            items.push({
                particulars: digitizerItem || '',
                takeoff,
                unit: unit || '',
                parsed: {}
            })
            if (tracker) tracker.markUsed(rowIndex)
        }
    })
    return items
}

/**
 * Processes pile cap items
 */
export const processPileCapItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isPileCap, parsePileCap, tracker)
}

/**
 * Processes strip footing items and groups them by size
 */
export const processStripFootingItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericFoundationItems(rawDataRows, headers, isStripFooting, parseStripFooting, tracker)

    // Group by groupKey (length x width)
    const groupMap = new Map()
    items.forEach(item => {
        const groupKey = item.parsed.groupKey || 'OTHER'
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    return mergeSingleItemGroupsIfAll(Array.from(groupMap.values()))
}

/**
 * Processes isolated footing items
 */
export const processIsolatedFootingItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isIsolatedFooting, parseIsolatedFooting, tracker)
}

/**
 * Processes pilaster items
 */
export const processPilasterItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isPilaster, parsePilaster, tracker)
}

/**
 * Processes grade beam items.
 * For Grade beams subsection we want ALL items together under a single group/sum,
 * so we simply return the flat list (no grouping by size here).
 */
export const processGradeBeamItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isGradeBeam, parseGradeBeam, tracker)
}

/**
 * Processes tie beam items and groups them by size
 */
export const processTieBeamItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericFoundationItems(rawDataRows, headers, isTieBeam, parseTieBeam, tracker)

    // Group by groupKey (width x height in inches)
    const groupMap = new Map()
    items.forEach(item => {
        const groupKey = item.parsed.groupKey || 'OTHER'
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    return mergeSingleItemGroupsIfAll(Array.from(groupMap.values()))
}

/**
 * Processes thickened slab items and groups them by size
 */
export const processThickenedSlabItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericFoundationItems(rawDataRows, headers, isThickenedSlab, parseThickenedSlab, tracker)

    // Group by groupKey (width x height)
    const groupMap = new Map()
    items.forEach(item => {
        const groupKey = item.parsed.groupKey || 'OTHER'
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    return mergeSingleItemGroupsIfAll(Array.from(groupMap.values()))
}

/**
 * Processes buttress items
 */
export const processButtressItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return null

    let buttressItem = null
    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = row[unitIdx]

        if (isButtress(digitizerItem)) {
            buttressItem = {
                particulars: digitizerItem,
                takeoff: total,
                unit: unit,
                rawRowNumber: rowIndex + 2
            }
            // Mark this row as used
            if (tracker) {
                tracker.markUsed(rowIndex)
            }
        }
    })

    return buttressItem
}

/**
 * Processes pier items and groups them by size (first bracket value)
 */
export const processPierItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericFoundationItems(rawDataRows, headers, isPier, parsePier, tracker)

    // Group by groupKey (first dimension - length)
    const groupMap = new Map()
    items.forEach(item => {
        const groupKey = item.parsed.groupKey || 'OTHER'
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    return mergeSingleItemGroupsIfAll(Array.from(groupMap.values()))
}

/**
 * Processes corbel items and groups them by size (first bracket value)
 */
export const processCorbelItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericFoundationItems(rawDataRows, headers, isCorbel, parseCorbel, tracker)

    // Group by groupKey (width x height)
    const groupMap = new Map()
    items.forEach(item => {
        const groupKey = item.parsed.groupKey || 'OTHER'
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    return mergeSingleItemGroupsIfAll(Array.from(groupMap.values()))
}

/**
 * Processes linear wall items and groups them by size (first bracket value)
 */
export const processLinearWallItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericFoundationItems(rawDataRows, headers, isLinearWall, parseLinearWall, tracker)

    // Group by groupKey (width x height)
    const groupMap = new Map()
    items.forEach(item => {
        const groupKey = item.parsed.groupKey || 'OTHER'
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    return mergeSingleItemGroupsIfAll(Array.from(groupMap.values()))
}

/**
 * Processes foundation wall items and groups them by size (first bracket value - width)
 */
export const processFoundationWallItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericFoundationItems(rawDataRows, headers, isFoundationWall, parseFoundationWall, tracker)

    // Group by groupKey (width x height)
    const groupMap = new Map()
    items.forEach(item => {
        const groupKey = item.parsed.groupKey || 'OTHER'
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    return mergeSingleItemGroupsIfAll(Array.from(groupMap.values()))
}

/**
 * Processes retaining wall items and groups them by size
 */
export const processRetainingWallItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericFoundationItems(rawDataRows, headers, isRetainingWall, parseRetainingWall, tracker)

    // Group by groupKey (width x height)
    const groupMap = new Map()
    items.forEach(item => {
        const groupKey = item.parsed.groupKey || 'OTHER'
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    return mergeSingleItemGroupsIfAll(Array.from(groupMap.values()))
}

/**
 * Processes barrier wall items and groups them by size
 */
export const processBarrierWallItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericFoundationItems(rawDataRows, headers, isBarrierWall, parseBarrierWall, tracker)

    // Group by groupKey (width x height)
    const groupMap = new Map()
    items.forEach(item => {
        const groupKey = item.parsed.groupKey || 'OTHER'
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                items: [],
                parsed: item.parsed
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    return mergeSingleItemGroupsIfAll(Array.from(groupMap.values()))
}

/**
 * Processes stem wall items
 */
export const processStemWallItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isStemWall, parseStemWall, tracker)
}

/**
 * Processes elevator pit items
 */
export const processElevatorPitItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isElevatorPit, parseElevatorPit, tracker)
}

/**
 * Processes service elevator pit items
 */
export const processServiceElevatorPitItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isServiceElevatorPit, parseServiceElevatorPit, tracker)
}

/**
 * Processes detention tank items
 */
export const processDetentionTankItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isDetentionTank, parseDetentionTank, tracker)
}

/**
 * Processes duplex sewage ejector pit items
 */
export const processDuplexSewageEjectorPitItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isDuplexSewageEjectorPit, parseDuplexSewageEjectorPit, tracker)
}

/**
 * Processes deep sewage ejector pit items
 */
export const processDeepSewageEjectorPitItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isDeepSewageEjectorPit, parseDeepSewageEjectorPit, tracker)
}

/**
 * Processes grease trap items
 */
export const processGreaseTrapItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isGreaseTrap, parseGreaseTrap, tracker)
}

/**
 * Processes house trap items
 */
export const processHouseTrapItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isHouseTrap, parseHouseTrap, tracker)
}

/**
 * Processes mat slab items
 */
export const processMatSlabItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isMatSlab, parseMatSlab, tracker)
}

/**
 * Processes mud slab items (Foundation section)
 */
export const processMudSlabFoundationItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isMudSlabFoundation, parseMudSlabFoundation, tracker)
}

/**
 * Processes SOG items
 */
export const processSOGItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isSOG, parseSOG, tracker)
}

/**
 * Processes stairs on grade items: one group per "stairs on grade" item.
 * Prioritize attaching one landing to each stairs group when available.
 * Landings are never in separate groups; they are always included with a stairs group.
 */
export const processStairsOnGradeItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const stairsList = []
    const landingsList = []

    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = row[unitIdx]

        if (isStairsOnGrade(digitizerItem)) {
            const parsed = parseStairsOnGrade(digitizerItem)
            const item = {
                particulars: digitizerItem,
                takeoff: total,
                unit: unit,
                parsed: parsed,
                rawRowNumber: rowIndex + 2
            }
            if (parsed.itemSubType === 'stairs') {
                stairsList.push(item)
            } else if (parsed.itemSubType === 'landings') {
                landingsList.push(item)
            }
            // Mark this row as used
            if (tracker) {
                tracker.markUsed(rowIndex)
            }
        }
    })

    // One group per stairs. Attach one landing to each stairs group when available.
    // Attach landings to the last N stairs (so if we have more stairs than landings, first stairs have no landing).
    const groups = []
    const stairIdentifiers = ['A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M', 'N', 'O', 'P', 'Q', 'R', 'S', 'T', 'U', 'V', 'W', 'X', 'Y', 'Z']

    stairsList.forEach((stairsItem, i) => {
        const items = []
        // Landing index for this group: attach landings to the last N stairs
        const landingIndex = i - (stairsList.length - landingsList.length)
        if (landingIndex >= 0 && landingIndex < landingsList.length) {
            items.push(landingsList[landingIndex])
        }
        items.push(stairsItem)

        groups.push({
            stairIdentifier: stairIdentifiers[i] || `Group${i + 1}`,
            items,
            hasStairs: true,
            hasLandings: items.length > 1
        })
    })

    return groups
}

/**
 * Processes electric conduit items (Foundation section)
 */
export const processElectricConduitItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isElectricConduit, parseElectricConduit, tracker)
}

/**
 * Generates formulas for Foundation items
 */
export const generateFoundationFormulas = (itemType, rowNum, itemData) => {
    const formulas = {
        ft: null,
        sqFt: null,
        lbs: null,
        cy: null,
        qtyFinal: null,
        takeoff: null,
        length: null,
        width: null,
        height: null,
        qty: null
    }

    const type = itemData.parsed?.type || itemType

    switch (type) {
        case 'drilled_foundation_pile':
            if (itemData.parsed?.isDualDiameter) {
                // Isolation casing (dual diameter):
                // H is height, I = H*C, J = E*C, K = (I*wt1) + (J*wt2)
                // E (QTY) is manual input, so don't set it.
                formulas.height = itemData.parsed.calculatedHeight || ''
                formulas.ft = `H${rowNum}*C${rowNum}` // I column
                formulas.sqFt = `E${rowNum}*C${rowNum}` // J column
                if (itemData.parsed?.weight && itemData.parsed?.weight2) {
                    formulas.lbs = `(I${rowNum}*${itemData.parsed.weight.toFixed(3)})+(J${rowNum}*${itemData.parsed.weight2.toFixed(3)})`
                }
                formulas.qtyFinal = `C${rowNum}`
            } else {
                // Single diameter: I=H*C, K=I*weight, M=C, J is empty
                formulas.height = itemData.parsed.calculatedHeight || '' // H column
                formulas.ft = `H${rowNum}*C${rowNum}` // I column (FT column, but formula is H*C)
                // J column should be empty - no formula, no value
                if (itemData.parsed?.weight) {
                    formulas.lbs = `I${rowNum}*${itemData.parsed.weight.toFixed(3)}`
                }
                formulas.qtyFinal = `C${rowNum}`
            }
            break

        case 'helical_foundation_pile':
            // I=H*C, K=I*weight, M=C
            formulas.height = itemData.parsed.calculatedHeight || ''
            formulas.ft = `H${rowNum}*C${rowNum}`
            if (itemData.parsed?.weight) {
                formulas.lbs = `I${rowNum}*${itemData.parsed.weight.toFixed(3)}`
            }
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'driven_foundation_pile':
            // I=H*C, K=I*weight, M=C
            formulas.height = itemData.parsed.calculatedHeight || ''
            formulas.ft = `H${rowNum}*C${rowNum}`
            if (itemData.parsed?.weight) {
                formulas.lbs = `I${rowNum}*${itemData.parsed.weight.toFixed(3)}`
            }
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'stelcor_drilled_displacement_pile':
            // I=H*C, K=I*weight, M=C
            formulas.height = itemData.parsed.calculatedHeight || ''
            formulas.ft = `H${rowNum}*C${rowNum}`
            if (itemData.parsed?.weight) {
                formulas.lbs = `I${rowNum}*${itemData.parsed.weight.toFixed(3)}`
            }
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'cfa_pile':
            // I=H*C, M=C
            formulas.height = itemData.parsed.calculatedHeight || ''
            formulas.ft = `H${rowNum}*C${rowNum}`
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'pile_cap':
            // J=C*H*G, L=J*F/27, M=C
            if (itemData.parsed?.length) formulas.length = itemData.parsed.length
            if (itemData.parsed?.width) formulas.width = itemData.parsed.width
            if (itemData.parsed?.height) formulas.height = itemData.parsed.height
            formulas.sqFt = `C${rowNum}*H${rowNum}*G${rowNum}`
            formulas.cy = `J${rowNum}*F${rowNum}/27`
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'strip_footing':
            // Width goes to column G, Height goes to column H
            // I=C, J=G*I (for SF/WF) or J=H*I (for ST)
            // L=J*H/27 (for SF/WF) or L=J*G/27 (for ST)
            if (itemData.parsed?.width) formulas.width = itemData.parsed.width
            if (itemData.parsed?.height) formulas.height = itemData.parsed.height
            formulas.ft = `C${rowNum}` // Column I = C
            // Column J: G*I for SF/WF, H*I for ST
            if (itemData.parsed?.itemType === 'ST') {
                formulas.sqFt = `H${rowNum}*I${rowNum}` // ST: J = H*I
                formulas.cy = `J${rowNum}*G${rowNum}/27` // ST: L = J*G/27
            } else {
                formulas.sqFt = `G${rowNum}*I${rowNum}` // SF/WF: J = G*I
                formulas.cy = `J${rowNum}*H${rowNum}/27` // SF/WF: L = J*H/27
            }
            break

        case 'isolated_footing':
            // J=C*F*G, L=J*H/27, M=C
            if (itemData.parsed?.length) formulas.length = itemData.parsed.length
            if (itemData.parsed?.width) formulas.width = itemData.parsed.width
            if (itemData.parsed?.height) formulas.height = itemData.parsed.height
            formulas.sqFt = `C${rowNum}*F${rowNum}*G${rowNum}`
            formulas.cy = `J${rowNum}*H${rowNum}/27`
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'pilaster':
            // J=C*H*G, L=J*F/27, M=C
            // Set length, width, height values (check for undefined, not falsy, since 0 is valid)
            if (itemData.parsed?.length !== undefined) formulas.length = itemData.parsed.length
            if (itemData.parsed?.width !== undefined) formulas.width = itemData.parsed.width
            if (itemData.parsed?.height !== undefined) formulas.height = itemData.parsed.height
            formulas.sqFt = `C${rowNum}*H${rowNum}*G${rowNum}`
            formulas.cy = `J${rowNum}*F${rowNum}/27`
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'grade_beam':
            // I=C, J=H*I, L=J*G/27
            if (itemData.parsed?.width) formulas.width = itemData.parsed.width
            if (itemData.parsed?.height) formulas.height = itemData.parsed.height
            formulas.ft = `C${rowNum}`
            formulas.sqFt = `H${rowNum}*I${rowNum}`
            formulas.cy = `J${rowNum}*G${rowNum}/27`
            break

        case 'tie_beam':
            // I=C, J=H*I, L=J*G/27
            if (itemData.parsed?.width) formulas.width = itemData.parsed.width
            if (itemData.parsed?.height) formulas.height = itemData.parsed.height
            formulas.ft = `C${rowNum}`
            formulas.sqFt = `H${rowNum}*I${rowNum}`
            formulas.cy = `J${rowNum}*G${rowNum}/27`
            break

        case 'thickened_slab':
            // I=C, J=H*I, L=J*G/27
            if (itemData.parsed?.width) formulas.width = itemData.parsed.width
            if (itemData.parsed?.height) formulas.height = itemData.parsed.height
            formulas.ft = `C${rowNum}`
            formulas.sqFt = `H${rowNum}*I${rowNum}`
            formulas.cy = `J${rowNum}*G${rowNum}/27`
            break

        case 'buttress_takeoff':
            // As per Takeoff count row:
            // F=1, G=1, H=1 are set in generateCalculationSheet.
            // I = H*C, J = C*H*G, L = J*F/27, M = C
            formulas.ft = `H${rowNum}*C${rowNum}`                    // Column I
            formulas.sqFt = `C${rowNum}*H${rowNum}*G${rowNum}`       // Column J
            formulas.cy = `J${rowNum}*F${rowNum}/27`                 // Column L
            formulas.qtyFinal = `C${rowNum}`                         // Column M
            break

        case 'pier':
            // J=C*F*G, L=J*H/27, M=C
            if (itemData.parsed?.length) formulas.length = itemData.parsed.length
            if (itemData.parsed?.width) formulas.width = itemData.parsed.width
            if (itemData.parsed?.height) formulas.height = itemData.parsed.height
            formulas.sqFt = `C${rowNum}*F${rowNum}*G${rowNum}`
            formulas.cy = `J${rowNum}*H${rowNum}/27`
            formulas.qtyFinal = `C${rowNum}`
            break

        case 'corbel':
            // I=C, J=I*H, L=J*G/27
            if (itemData.parsed?.width) formulas.width = itemData.parsed.width
            if (itemData.parsed?.height) formulas.height = itemData.parsed.height
            formulas.ft = `C${rowNum}`
            formulas.sqFt = `I${rowNum}*H${rowNum}`
            formulas.cy = `J${rowNum}*G${rowNum}/27`
            break

        case 'linear_wall':
            // I=C, J=I*H, L=J*G/27
            if (itemData.parsed?.width) formulas.width = itemData.parsed.width
            if (itemData.parsed?.height) formulas.height = itemData.parsed.height
            formulas.ft = `C${rowNum}`
            formulas.sqFt = `I${rowNum}*H${rowNum}`
            formulas.cy = `J${rowNum}*G${rowNum}/27`
            break

        case 'foundation_wall':
            // I=C, J=I*H, L=J*G/27
            if (itemData.parsed?.width) formulas.width = itemData.parsed.width
            if (itemData.parsed?.height) formulas.height = itemData.parsed.height
            formulas.ft = `C${rowNum}`
            formulas.sqFt = `I${rowNum}*H${rowNum}`
            formulas.cy = `J${rowNum}*G${rowNum}/27`
            break

        case 'retaining_wall':
            // I=C, J=I*H, L=J*G/27
            if (itemData.parsed?.width) formulas.width = itemData.parsed.width
            if (itemData.parsed?.height) formulas.height = itemData.parsed.height
            formulas.ft = `C${rowNum}`
            formulas.sqFt = `I${rowNum}*H${rowNum}`
            formulas.cy = `J${rowNum}*G${rowNum}/27`
            break

        case 'barrier_wall':
            // I=C, J=I*H, L=J*G/27
            if (itemData.parsed?.width) formulas.width = itemData.parsed.width
            if (itemData.parsed?.height) formulas.height = itemData.parsed.height
            formulas.ft = `C${rowNum}`
            formulas.sqFt = `I${rowNum}*H${rowNum}`
            formulas.cy = `J${rowNum}*G${rowNum}/27`
            break

        case 'stem_wall':
            // I=C, J=I*H, L=J*G/27
            if (itemData.parsed?.width) formulas.width = itemData.parsed.width
            if (itemData.parsed?.height) formulas.height = itemData.parsed.height
            formulas.ft = `C${rowNum}`
            formulas.sqFt = `I${rowNum}*H${rowNum}`
            formulas.cy = `J${rowNum}*G${rowNum}/27`
            break

        case 'elevator_pit':
        case 'service_elevator_pit':
            const subType = itemData.parsed?.itemSubType
            if (subType === 'sump_pit') {
                // Sump pit: J=16*C, L=C*1.3, M=C
                formulas.sqFt = `16*C${rowNum}`
                formulas.cy = `C${rowNum}*1.3`
                formulas.qtyFinal = `C${rowNum}`
            } else if (subType === 'slab' || subType === 'mat') {
                // Elev. pit slab / Elev. pit mat: I empty, J=C, L=J*H/27
                formulas.sqFt = `C${rowNum}` // J = C
                if (itemData.parsed?.heightFromH !== undefined) {
                    formulas.height = itemData.parsed.heightFromH
                }
                formulas.cy = `J${rowNum}*H${rowNum}/27`
            } else if (subType === 'wall' || subType === 'slope_transition') {
                // Elev. pit wall/slope: I=C, J=I*H, L=J*G/27
                formulas.ft = `C${rowNum}`
                formulas.sqFt = `I${rowNum}*H${rowNum}`
                if (itemData.parsed?.width !== undefined) formulas.width = itemData.parsed.width
                if (itemData.parsed?.height !== undefined) formulas.height = itemData.parsed.height
                formulas.cy = `J${rowNum}*G${rowNum}/27`
            }
            break

        case 'detention_tank':
            const dtSubType = itemData.parsed?.itemSubType
            if (dtSubType === 'slab' || dtSubType === 'lid_slab') {
                // Detention tank slab: J=C, L=J*H/27
                formulas.sqFt = `C${rowNum}`
                if (itemData.parsed?.heightFromName !== undefined) {
                    formulas.height = itemData.parsed.heightFromName
                }
                formulas.cy = `J${rowNum}*H${rowNum}/27`
            } else if (dtSubType === 'wall') {
                // Detention tank wall: I=C, J=I*H, L=J*G/27
                // Note: F (Length) should be empty, not set
                formulas.ft = `C${rowNum}`
                formulas.sqFt = `I${rowNum}*H${rowNum}`
                // Do not set formulas.length for Detention tank wall (F should be empty)
                if (itemData.parsed?.width !== undefined) formulas.width = itemData.parsed.width
                if (itemData.parsed?.height !== undefined) formulas.height = itemData.parsed.height
                formulas.cy = `J${rowNum}*G${rowNum}/27`
            }
            break

        case 'duplex_sewage_ejector_pit':
            const dsepSubType = itemData.parsed?.itemSubType
            if (dsepSubType === 'slab') {
                // Duplex sewage ejector pit slab: J=C, L=J*H/27
                formulas.sqFt = `C${rowNum}`
                if (itemData.parsed?.heightFromName !== undefined) {
                    formulas.height = itemData.parsed.heightFromName
                }
                formulas.cy = `J${rowNum}*H${rowNum}/27`
            } else if (dsepSubType === 'wall') {
                // Duplex sewage ejector pit wall: I=C, J=I*H, L=J*G/27
                formulas.ft = `C${rowNum}`
                formulas.sqFt = `I${rowNum}*H${rowNum}`
                if (itemData.parsed?.width !== undefined) formulas.width = itemData.parsed.width
                if (itemData.parsed?.height !== undefined) formulas.height = itemData.parsed.height
                formulas.cy = `J${rowNum}*G${rowNum}/27`
            }
            break

        case 'deep_sewage_ejector_pit':
            const deepSepSubType = itemData.parsed?.itemSubType
            if (deepSepSubType === 'slab') {
                // Deep sewage ejector pit slab: J=C, L=J*H/27
                formulas.sqFt = `C${rowNum}`
                if (itemData.parsed?.heightFromName !== undefined) {
                    formulas.height = itemData.parsed.heightFromName
                }
                formulas.cy = `J${rowNum}*H${rowNum}/27`
            } else if (deepSepSubType === 'wall') {
                // Deep sewage ejector pit wall: I=C, J=I*H, L=J*G/27
                formulas.ft = `C${rowNum}`
                formulas.sqFt = `I${rowNum}*H${rowNum}`
                if (itemData.parsed?.width !== undefined) formulas.width = itemData.parsed.width
                if (itemData.parsed?.height !== undefined) formulas.height = itemData.parsed.height
                formulas.cy = `J${rowNum}*G${rowNum}/27`
            }
            break

        case 'grease_trap':
            const gtSubType = itemData.parsed?.itemSubType
            if (gtSubType === 'slab') {
                // Grease trap slab: J=C, L=J*H/27
                formulas.sqFt = `C${rowNum}`
                if (itemData.parsed?.heightFromName !== undefined) {
                    formulas.height = itemData.parsed.heightFromName
                }
                formulas.cy = `J${rowNum}*H${rowNum}/27`
            } else if (gtSubType === 'wall') {
                // Grease trap wall: I=C, J=I*H, L=J*G/27
                formulas.ft = `C${rowNum}`
                formulas.sqFt = `I${rowNum}*H${rowNum}`
                if (itemData.parsed?.width !== undefined) formulas.width = itemData.parsed.width
                if (itemData.parsed?.height !== undefined) formulas.height = itemData.parsed.height
                formulas.cy = `J${rowNum}*G${rowNum}/27`
            }
            break

        case 'house_trap':
            const htSubType = itemData.parsed?.itemSubType
            if (htSubType === 'slab') {
                // House trap pit slab: J=C, L=J*H/27
                // G (Width) should be empty, not set
                formulas.sqFt = `C${rowNum}`
                if (itemData.parsed?.heightFromName !== undefined) {
                    formulas.height = itemData.parsed.heightFromName
                }
                // Do not set formulas.width for pit slab items (G should be empty)
                formulas.cy = `J${rowNum}*H${rowNum}/27`
            } else if (htSubType === 'wall') {
                // House trap pit: I=C, J=I*H, L=J*G/27
                formulas.ft = `C${rowNum}`
                formulas.sqFt = `I${rowNum}*H${rowNum}`
                if (itemData.parsed?.width !== undefined) formulas.width = itemData.parsed.width
                if (itemData.parsed?.height !== undefined) formulas.height = itemData.parsed.height
                formulas.cy = `J${rowNum}*G${rowNum}/27`
            }
            break

        case 'mat_slab':
            const msSubType = itemData.parsed?.itemSubType
            if (msSubType === 'mat') {
                // Mat items: J=C, L=J*H/27
                formulas.sqFt = `C${rowNum}`
                if (itemData.parsed?.heightFromH !== undefined) {
                    formulas.height = itemData.parsed.heightFromH
                }
                formulas.cy = `J${rowNum}*H${rowNum}/27`
            } else if (msSubType === 'haunch') {
                // Haunch items: I=C, J=I*H, L=J*G/27
                formulas.ft = `C${rowNum}`
                formulas.sqFt = `I${rowNum}*H${rowNum}`
                if (itemData.parsed?.width !== undefined) formulas.width = itemData.parsed.width
                if (itemData.parsed?.height !== undefined) formulas.height = itemData.parsed.height
                formulas.cy = `J${rowNum}*G${rowNum}/27`
            }
            break

        case 'mud_slab_foundation':
            // Mud slab: J=C, L=J*H/27
            // C and H are manual entries (blank)
            formulas.sqFt = `C${rowNum}`
            formulas.cy = `J${rowNum}*H${rowNum}/27`
            break

        case 'sog':
            const sogSubType = itemData.parsed?.itemSubType
            if (sogSubType === 'gravel') {
                // Gravel: J=C, L=J*H/27
                // H is blank (manual entry - they will fill it themselves)
                formulas.sqFt = `C${rowNum}`
                // Do not set formulas.height for Gravel (H should be blank)
                formulas.cy = `J${rowNum}*H${rowNum}/27`
            } else if (sogSubType === 'gravel_backfill') {
                // Gravel backfill: J=C, L=J*H/27, H from H=X'-Y" format
                formulas.sqFt = `C${rowNum}`
                if (itemData.parsed?.heightFromH !== undefined) {
                    formulas.height = itemData.parsed.heightFromH
                }
                formulas.cy = `J${rowNum}*H${rowNum}/27`
            } else if (sogSubType === 'geotextile') {
                // Geotextile filter fabric: J=C, L blank, H blank
                formulas.sqFt = `C${rowNum}`
                // H and L should be blank
            } else if (sogSubType === 'sog_slab') {
                // SOG slabs: J=C, L=J*H/27, H from name (4", 6", 5")
                formulas.sqFt = `C${rowNum}`
                if (itemData.parsed?.heightFromName !== undefined) {
                    formulas.height = itemData.parsed.heightFromName
                }
                formulas.cy = `J${rowNum}*H${rowNum}/27`
            } else if (sogSubType === 'sog_step') {
                // SOG step: I=C, J=I*H, L=J*G/27
                formulas.ft = `C${rowNum}`
                formulas.sqFt = `I${rowNum}*H${rowNum}`
                if (itemData.parsed?.width !== undefined) formulas.width = itemData.parsed.width
                if (itemData.parsed?.height !== undefined) formulas.height = itemData.parsed.height
                formulas.cy = `J${rowNum}*G${rowNum}/27`
            }
            break

        case 'stairs_on_grade':
            const sogStairsSubType = itemData.parsed?.itemSubType
            if (sogStairsSubType === 'stairs') {
                // Stairs on grade: F=11/12, H=7/12, J=C*G*F, L=J*H/27, M=C
                formulas.length = '11/12' // Formula: =11/12
                formulas.height = '7/12' // Formula: =7/12
                if (itemData.parsed?.widthFromName !== undefined) {
                    formulas.width = itemData.parsed.widthFromName
                }
                // G (Width) is either from name or manual input (will be set in generateCalculationSheet)
                formulas.sqFt = `C${rowNum}*G${rowNum}*F${rowNum}`
                formulas.cy = `J${rowNum}*H${rowNum}/27`
                formulas.qtyFinal = `C${rowNum}`
            } else if (sogStairsSubType === 'landings') {
                // Landings: J=C, L=J*H/27, H=0.67 (8" = 0.67 feet)
                formulas.height = 0.67 // Constant: 8" = 0.67 feet
                formulas.sqFt = `C${rowNum}`
                formulas.cy = `J${rowNum}*H${rowNum}/27`
            } else if (sogStairsSubType === 'stair_slab') {
                // Stair slab: C=C[stairs_row]*1.3, I=C, J=I*H, H=0.67, G = G[stairs_row]
                const stairsRow = itemData.parsed?.stairsRow
                const hasWidthFromName = itemData.parsed?.hasWidthFromName
                if (stairsRow) {
                    formulas.takeoff = `C${stairsRow}*1.3` // C = C[stairs_row]*1.3
                    // Width (G) for stair slab is always =G of the stairs on grade row
                    formulas.width = `G${stairsRow}`
                }
                formulas.ft = `C${rowNum}`
                formulas.height = 0.67 // Constant: 8" = 0.67 feet
                formulas.sqFt = `I${rowNum}*H${rowNum}`
                // L formula: if hasWidthFromName, L=J*G/27, else L=J*F/27
                if (hasWidthFromName) {
                    formulas.cy = `J${rowNum}*G${rowNum}/27`
                } else {
                    formulas.cy = `J${rowNum}*F${rowNum}/27`
                    if (stairsRow) {
                        formulas.length = `G${stairsRow}` // F = G[stairs_row]
                    }
                }
            }
            break

        case 'electric_conduit':
            // Electric conduit: I (FT) = C (Takeoff)
            formulas.ft = `C${rowNum}`
            break
    }

    return formulas
}

export default {
    processDrilledFoundationPileItems,
    processHelicalFoundationPileItems,
    processDrivenFoundationPileItems,
    processStelcorDrilledDisplacementPileItems,
    processCFAPileItems,
    processPileCapItems,
    processStripFootingItems,
    processIsolatedFootingItems,
    processPilasterItems,
    processGradeBeamItems,
    processTieBeamItems,
    processThickenedSlabItems,
    processButtressItems,
    processPierItems,
    processCorbelItems,
    processLinearWallItems,
    processFoundationWallItems,
    processRetainingWallItems,
    processBarrierWallItems,
    processStemWallItems,
    processElevatorPitItems,
    processDetentionTankItems,
    processDuplexSewageEjectorPitItems,
    processDeepSewageEjectorPitItems,
    processGreaseTrapItems,
    processHouseTrapItems,
    processMatSlabItems,
    processMudSlabFoundationItems,
    processSOGItems,
    processStairsOnGradeItems,
    processElectricConduitItems,
    generateFoundationFormulas
}

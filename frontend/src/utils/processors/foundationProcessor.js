import { mergeSingleItemGroupsIfAll } from '../groupingUtils.js'
import {
    isDrilledFoundationPile,
    isHelicalFoundationPile,
    isDrivenFoundationPile,
    isStelcorDrilledDisplacementPile,
    isCFAPile,
    isMiscellaneousFoundationPile,
    tryMatchPileStructure,
    isPileCap,
    isStripFooting,
    isIsolatedFooting,
    isPilaster,
    isGradeBeam,
    isTieBeam,
    isStrapBeam,
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
    isSumpPumpPit,
    isGreaseTrap,
    isHouseTrap,
    isMatSlab,
    isMudSlabFoundation,
    isSOG,
    isROG,
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
    parseStrapBeam,
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
    parseSumpPumpPit,
    parseGreaseTrap,
    parseHouseTrap,
    parseMatSlab,
    parseMudSlabFoundation,
    parseSOG,
    parseROG,
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
            const parsed = parseDrilledFoundationPile(digitizerItem)
            
            const influValue = influIdx !== -1 ? row[influIdx] : null
            const hasInfluFromColumn = influValue && String(influValue).toLowerCase().includes('influ')
            const hasInflu = parsed.hasInfluence || hasInfluFromColumn

            // Mark this row as used
            if (tracker) {
                tracker.markUsed(rowIndex)
            }

            const itemGroupKey = parsed.groupKey || 'OTHER'

            // If current group exists but has different structure (e.g. single vs dual diameter), save it and start new group
            if (currentGroup && currentGroup.groupKey !== itemGroupKey) {
                groups.push(currentGroup)
                currentGroup = null
            }

            // If no current group, start a new one
            if (!currentGroup) {
                currentGroup = {
                    groupKey: itemGroupKey,
                    items: [{
                        particulars: digitizerItem,
                        takeoff: total,
                        unit: unit,
                        parsed: parsed,
                        rawRowNumber: rowIndex + 2,
                        hasInflu: hasInflu
                    }],
                    parsed: parsed,
                    hasInflu: hasInflu
                }
            } else {
                // Same structure - add as separate item (do not sum takeoffs)
                currentGroup.items.push({
                    particulars: digitizerItem,
                    takeoff: total,
                    unit: unit,
                    parsed: parsed,
                    rawRowNumber: rowIndex + 2,
                    hasInflu: hasInflu
                })
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
                parsed: item.parsed,
                hasInflu: false
            })
        }
        groupMap.get(groupKey).items.push(item)
        // Set hasInflu flag if any item has influence
        if (item.parsed.hasInfluence) {
            groupMap.get(groupKey).hasInflu = true
        }
    })

    // Set hasInflu flag on each item based on their group
    const groups = Array.from(groupMap.values())
    groups.forEach(group => {
        group.items.forEach(item => {
            item.hasInflu = group.hasInflu
        })
    })

    return mergeSingleItemGroupsIfAll(groups)
}

/**
 * Processes driven foundation pile items
 */
export const processDrivenFoundationPileItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericFoundationItems(rawDataRows, headers, isDrivenFoundationPile, parseDrivenFoundationPile, tracker)

    // Add hasInflu flag to each item based on parsed data
    items.forEach(item => {
        item.hasInflu = item.parsed.hasInfluence || false
    })

    if (items.length === 0) return []
    const lastItem = items[items.length - 1]
    if (lastItem.hasInflu) {
        let firstInfluIdx = items.length - 1
        while (firstInfluIdx - 1 >= 0 && items[firstInfluIdx - 1].hasInflu) {
            firstInfluIdx--
        }

        const nonInfluItems = items.slice(0, firstInfluIdx)
        const influItems = items.slice(firstInfluIdx)

        const groups = []

        if (nonInfluItems.length > 0) {
            groups.push({
                groupKey: 'DRIVEN_REMAINING',
                items: nonInfluItems,
                parsed: nonInfluItems[0].parsed || {},
                hasInflu: false
            })
        }

        groups.push({
            groupKey: 'DRIVEN_INFLUENCE',
            items: influItems,
            parsed: influItems[0].parsed || {},
            hasInflu: true
        })

        groups.forEach(g => g.items.forEach(it => { it.hasInflu = g.hasInflu }))

        return groups
    }

    const groupMap = new Map()
    items.forEach(item => {
        const groupKey = item.parsed.groupKey || 'OTHER'
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                items: [],
                parsed: item.parsed,
                hasInflu: false
            })
        }
        groupMap.get(groupKey).items.push(item)
        if (item.parsed.hasInfluence) {
            groupMap.get(groupKey).hasInflu = true
        }
    })

    const groups = Array.from(groupMap.values())
    groups.forEach(group => {
        group.items.forEach(item => {
            item.hasInflu = group.hasInflu
        })
    })

    return mergeSingleItemGroupsIfAll(groups)
}

/**
 * Processes stelcor drilled displacement pile items and groups by empty rows and influence status
 */
export const processStelcorDrilledDisplacementPileItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const items = []

    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const digitizerText = digitizerItem ? String(digitizerItem).trim() : ''
        const isRowEmpty = !digitizerItem || digitizerText === ''

        // Check if this is a sum row (empty digitizer and empty takeoff)
        const takeoffValue = row[totalIdx]
        const isSumRow = isRowEmpty && (takeoffValue === '' || takeoffValue === null || takeoffValue === undefined)

        if (isRowEmpty || isSumRow) {
            return
        }

        // Check if this is a stelcor drilled displacement pile
        if (isStelcorDrilledDisplacementPile(digitizerItem)) {
            const total = parseFloat(row[totalIdx]) || 0
            const unit = row[unitIdx]
            const parsed = parseStelcorDrilledDisplacementPile(digitizerItem)

            // Mark this row as used
            if (tracker) {
                tracker.markUsed(rowIndex)
            }

            items.push({
                particulars: digitizerItem,
                takeoff: total,
                unit: unit,
                parsed: parsed,
                rawRowNumber: rowIndex + 2,
                hasInflu: parsed.hasInfluence || false
            })
        }
    })

    if (items.length === 0) return []

    const lastItem = items[items.length - 1]
    if (lastItem.hasInflu) {
        // Group by size (groupKey) first, then split by influence status
        const groupMap = new Map()
        items.forEach(item => {
            const groupKey = item.parsed.groupKey || 'OTHER'
            if (!groupMap.has(groupKey)) {
                groupMap.set(groupKey, { nonInflu: [], influ: [] })
            }
            if (item.hasInflu) {
                groupMap.get(groupKey).influ.push(item)
            } else {
                groupMap.get(groupKey).nonInflu.push(item)
            }
        })

        const groups = []
        groupMap.forEach((sizeGroup, groupKey) => {
            if (sizeGroup.nonInflu.length > 0) {
                groups.push({
                    groupKey: groupKey,
                    items: sizeGroup.nonInflu,
                    parsed: sizeGroup.nonInflu[0].parsed || {},
                    hasInflu: false
                })
            }
            if (sizeGroup.influ.length > 0) {
                groups.push({
                    groupKey: `INFLU-${groupKey}`,
                    items: sizeGroup.influ,
                    parsed: sizeGroup.influ[0].parsed || {},
                    hasInflu: true
                })
            }
        })
        return groups
    }

    // No influence items - group by size only
    const groupMap = new Map()
    items.forEach(item => {
        const groupKey = item.parsed.groupKey || 'OTHER'
        if (!groupMap.has(groupKey)) {
            groupMap.set(groupKey, {
                groupKey: groupKey,
                items: [],
                parsed: item.parsed,
                hasInflu: false
            })
        }
        groupMap.get(groupKey).items.push(item)
    })

    return Array.from(groupMap.values())
}

/**
 * Processes CFA pile items and groups by influence status
 */
export const processCFAPileItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const items = []

    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const digitizerText = digitizerItem ? String(digitizerItem).trim() : ''
        const isRowEmpty = !digitizerItem || digitizerText === ''

        // Check if this is a sum row (empty digitizer and empty takeoff)
        const takeoffValue = row[totalIdx]
        const isSumRow = isRowEmpty && (takeoffValue === '' || takeoffValue === null || takeoffValue === undefined)

        if (isRowEmpty || isSumRow) {
            return
        }

        // Check if this is a CFA pile
        if (isCFAPile(digitizerItem)) {
            const total = parseFloat(row[totalIdx]) || 0
            const unit = row[unitIdx]
            const parsed = parseCFAPile(digitizerItem)

            // Mark this row as used
            if (tracker) {
                tracker.markUsed(rowIndex)
            }

            items.push({
                particulars: digitizerItem,
                takeoff: total,
                unit: unit,
                parsed: parsed,
                rawRowNumber: rowIndex + 2,
                hasInflu: parsed.hasInfluence || false
            })
        }
    })

    if (items.length === 0) return []

    // Split items into non-influence and influence groups
    const nonInfluItems = items.filter(item => !item.hasInflu)
    const influItems = items.filter(item => item.hasInflu)

    const groups = []

    if (nonInfluItems.length > 0) {
        groups.push({
            groupKey: 'CFA_REMAINING',
            items: nonInfluItems,
            parsed: nonInfluItems[0].parsed || {},
            hasInflu: false
        })
    }

    if (influItems.length > 0) {
        groups.push({
            groupKey: 'CFA_INFLUENCE',
            items: influItems,
            parsed: influItems[0].parsed || {},
            hasInflu: true
        })
    }

    return groups
}

/**
 * Processes miscellaneous pile items and groups them by empty rows
 * Pile items that don't fit other foundation or SOE subsections.
 * Run after other foundation pile processors; only processes rows not yet used.
 */
export const processMiscellaneousPileItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const groups = []
    let currentGroup = null

    rawDataRows.forEach((row, rowIndex) => {
        if (tracker && tracker.isUsed(rowIndex)) return

        const digitizerItem = row[digitizerIdx]
        const digitizerText = digitizerItem ? String(digitizerItem).trim() : ''
        const isRowEmpty = !digitizerItem || digitizerText === ''

        // Check if this is a sum row (empty digitizer and empty takeoff)
        const takeoffValue = row[totalIdx]
        const isSumRow = isRowEmpty && (takeoffValue === '' || takeoffValue === null || takeoffValue === undefined)

        if (isRowEmpty || isSumRow) {
            // Empty row or sum row - save current group if it exists and start a new one
            if (currentGroup && currentGroup.items.length > 0) {
                groups.push(currentGroup)
                currentGroup = null
            }
            return
        }

        // Check if this is a miscellaneous pile item
        if (isMiscellaneousFoundationPile(digitizerItem)) {
            // Only add if item structure matches one of the known pile types
            const match = tryMatchPileStructure(digitizerItem)
            if (match) {
                const total = parseFloat(row[totalIdx]) || 0
                const unit = row[unitIdx]
                const hasInflu = match.parsed.hasInfluence || false

                // Mark this row as used
                if (tracker) {
                    tracker.markUsed(rowIndex)
                }

                // Start a new group for each item
                if (currentGroup && currentGroup.items.length > 0) {
                    groups.push(currentGroup)
                }

                currentGroup = {
                    groupKey: match.parsed.groupKey || 'OTHER',
                    items: [{
                        particulars: digitizerItem,
                        takeoff: total,
                        unit: unit,
                        parsed: match.parsed,
                        matchedPileType: match.matchedType,
                        rawRowNumber: rowIndex + 2,
                        hasInflu: hasInflu
                    }],
                    parsed: match.parsed,
                    hasInflu: hasInflu
                }
            }
            // If no match: do not add, do not mark as used - item stays in unused data
        } else {
            // Not a miscellaneous pile - if we have a current group, save it
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
 * Processes strap beam items.
 * Same logic as grade beams: flat list, all items under single sum.
 */
export const processStrapBeamItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isStrapBeam, parseStrapBeam, tracker)
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
 * Processes buttress items - returns takeoff from Buttress raw data, or { takeoff: 0, unit: 'EA' } if not available (matches Columns subsection logic)
 */
export const processButtressItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) {
        return { particulars: 'Buttress', takeoff: 0, unit: 'EA' }
    }

    let buttressItem = { particulars: 'Buttress', takeoff: 0, unit: 'EA' }
    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = row[unitIdx]

        if (isButtress(digitizerItem)) {
            buttressItem = {
                particulars: digitizerItem,
                takeoff: total,
                unit: unit || 'EA',
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
 * Process sump pump pit items (same grouping and formulas as Duplex sewage ejector pit)
 */
export const processSumpPumpPitItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isSumpPumpPit, parseSumpPumpPit, tracker)
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
 * Processes mat slab items.
 * Groups mats by H value; each haunch is assigned to the mat group that immediately precedes it in the raw order.
 * So: Mat-1, Mat-2, Mat-3, Haunch -> group1; Mat-1, Mat-2, Mat-3, Haunch -> group2 (each haunch in its own group).
 */
export const processMatSlabItems = (rawDataRows, headers, tracker = null) => {
    const items = processGenericFoundationItems(rawDataRows, headers, isMatSlab, parseMatSlab, tracker)
    if (items.length === 0) return []

    const groups = []
    let currentGroup = null

    items.forEach(item => {
        const subType = item.parsed?.itemSubType
        if (subType === 'mat') {
            const groupKey = item.parsed.groupKey || 'OTHER'
            if (currentGroup && currentGroup.groupKey !== groupKey) {
                groups.push(currentGroup)
                currentGroup = null
            }
            if (!currentGroup) {
                currentGroup = { groupKey, items: [] }
            }
            currentGroup.items.push(item)
        } else if (subType === 'haunch') {
            if (currentGroup) {
                currentGroup.items.push(item)
            }
            // Haunch does not start a new group; next mat will
        }
    })
    if (currentGroup) groups.push(currentGroup)
    return groups
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
 * Processes ROG (Ramp on grade) items
 */
export const processROGItems = (rawDataRows, headers, tracker = null) => {
    return processGenericFoundationItems(rawDataRows, headers, isROG, parseROG, tracker)
}

/**
 * Processes stairs on grade items (Foundation section).
 * Same logic as Demo stair on grade: group by text after @, or NO_AT.
 * Raw items: "Stairs on grade @ Stair A1", "Landings on grade @ Stair A1", "Stairs on grade", "Landings on grade"
 */
export const processStairsOnGradeItems = (rawDataRows, headers, tracker = null) => {
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
    const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

    if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return []

    const groupMap = new Map() // key -> { heading, stairs, landings }

    rawDataRows.forEach((row, rowIndex) => {
        const digitizerItem = row[digitizerIdx]
        const total = parseFloat(row[totalIdx]) || 0
        const unit = row[unitIdx]

        // Exclude demo stair items - they belong in Demolition section only
        const itemLower = (digitizerItem || '').toLowerCase()
        if (itemLower.includes('demo') && (itemLower.includes('stairs on grade') || itemLower.includes('landings on grade'))) {
            return
        }

        if (isStairsOnGrade(digitizerItem)) {
            const parsed = parseStairsOnGrade(digitizerItem)
            const key = parsed.groupKey || 'NO_AT'
            const heading = key !== 'NO_AT' ? key : null

            if (!groupMap.has(key)) {
                groupMap.set(key, { heading, groupKey: key, stairs: null, landings: null })
            }
            const g = groupMap.get(key)
            if (parsed.itemSubType === 'stairs') {
                g.stairs = { particulars: digitizerItem, takeoff: total, unit: unit, parsed, rawRowNumber: rowIndex + 2 }
            } else if (parsed.itemSubType === 'landings') {
                g.landings = { particulars: digitizerItem, takeoff: total, unit: unit, parsed, rawRowNumber: rowIndex + 2 }
            }
            if (tracker) tracker.markUsed(rowIndex)
        }
    })

    const groups = Array.from(groupMap.values()).filter(g => g.stairs || g.landings)
    // Stair col (or group key containing "stair col") should come last, after normal stairs (Stair A1, A2, etc.)
    const isStairCol = (g) => {
        const k = (g.groupKey || g.heading || '').toLowerCase()
        return k.includes('stair col') || k === 'col'
    }
    groups.sort((a, b) => {
        const aLast = isStairCol(a)
        const bLast = isStairCol(b)
        if (aLast && !bLast) return 1
        if (!aLast && bLast) return -1
        return 0
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

        case 'strap_beam':
            // Same as grade_beam: I=C, J=H*I, L=J*G/27
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
            // As per Takeoff count row: F, G, H, I, J, L empty; M = C
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
            } else if (subType === 'slab' || subType === 'mat' || subType === 'mat_slab') {
                // Elev. pit slab / Elev. pit mat / Elev. pit mat slab: I empty, J=C, L=J*H/27
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
            if (dsepSubType === 'slab' || dsepSubType === 'mat' || dsepSubType === 'mat_slab') {
                // Duplex sewage ejector pit slab/mat/mat_slab: J=C, L=J*H/27
                formulas.sqFt = `C${rowNum}`
                if (itemData.parsed?.heightFromName !== undefined) {
                    formulas.height = itemData.parsed.heightFromName
                }
                formulas.cy = `J${rowNum}*H${rowNum}/27`
            } else if (dsepSubType === 'wall' || dsepSubType === 'slope_transition') {
                // Duplex sewage ejector pit wall/slope: I=C, J=I*H, L=J*G/27
                formulas.ft = `C${rowNum}`
                formulas.sqFt = `I${rowNum}*H${rowNum}`
                if (itemData.parsed?.width !== undefined) formulas.width = itemData.parsed.width
                if (itemData.parsed?.height !== undefined) formulas.height = itemData.parsed.height
                formulas.cy = `J${rowNum}*G${rowNum}/27`
            }
            break

        case 'deep_sewage_ejector_pit':
            const deepSepSubType = itemData.parsed?.itemSubType
            if (deepSepSubType === 'slab' || deepSepSubType === 'mat' || deepSepSubType === 'mat_slab') {
                // Deep sewage ejector pit slab/mat/mat_slab: J=C, L=J*H/27
                formulas.sqFt = `C${rowNum}`
                if (itemData.parsed?.heightFromName !== undefined) {
                    formulas.height = itemData.parsed.heightFromName
                }
                formulas.cy = `J${rowNum}*H${rowNum}/27`
            } else if (deepSepSubType === 'wall' || deepSepSubType === 'slope_transition') {
                // Deep sewage ejector pit wall/slope: I=C, J=I*H, L=J*G/27
                formulas.ft = `C${rowNum}`
                formulas.sqFt = `I${rowNum}*H${rowNum}`
                if (itemData.parsed?.width !== undefined) formulas.width = itemData.parsed.width
                if (itemData.parsed?.height !== undefined) formulas.height = itemData.parsed.height
                formulas.cy = `J${rowNum}*G${rowNum}/27`
            }
            break

        case 'sump_pump_pit':
            const sppSubType = itemData.parsed?.itemSubType
            if (sppSubType === 'slab' || sppSubType === 'mat' || sppSubType === 'mat_slab') {
                // Sump pump pit slab/mat/mat_slab: J=C, L=J*H/27
                formulas.sqFt = `C${rowNum}`
                if (itemData.parsed?.heightFromName !== undefined) {
                    formulas.height = itemData.parsed.heightFromName
                }
                formulas.cy = `J${rowNum}*H${rowNum}/27`
            } else if (sppSubType === 'wall' || sppSubType === 'slope_transition') {
                // Sump pump pit wall/slope: I=C, J=I*H, L=J*G/27
                formulas.ft = `C${rowNum}`
                formulas.sqFt = `I${rowNum}*H${rowNum}`
                if (itemData.parsed?.width !== undefined) formulas.width = itemData.parsed.width
                if (itemData.parsed?.height !== undefined) formulas.height = itemData.parsed.height
                formulas.cy = `J${rowNum}*G${rowNum}/27`
            }
            break

        case 'grease_trap':
            const gtSubType = itemData.parsed?.itemSubType
            if (gtSubType === 'slab' || gtSubType === 'mat' || gtSubType === 'mat_slab') {
                // Grease trap slab/mat/mat_slab: J=C, L=J*H/27
                formulas.sqFt = `C${rowNum}`
                if (itemData.parsed?.heightFromName !== undefined) {
                    formulas.height = itemData.parsed.heightFromName
                }
                formulas.cy = `J${rowNum}*H${rowNum}/27`
            } else if (gtSubType === 'wall' || gtSubType === 'slope_transition') {
                // Grease trap wall/slope: I=C, J=I*H, L=J*G/27
                formulas.ft = `C${rowNum}`
                formulas.sqFt = `I${rowNum}*H${rowNum}`
                if (itemData.parsed?.width !== undefined) formulas.width = itemData.parsed.width
                if (itemData.parsed?.height !== undefined) formulas.height = itemData.parsed.height
                formulas.cy = `J${rowNum}*G${rowNum}/27`
            }
            break

        case 'house_trap':
            const htSubType = itemData.parsed?.itemSubType
            if (htSubType === 'slab' || htSubType === 'mat' || htSubType === 'mat_slab') {
                // House trap slab/mat/mat_slab: J=C, L=J*H/27
                formulas.sqFt = `C${rowNum}`
                if (itemData.parsed?.heightFromName !== undefined) {
                    formulas.height = itemData.parsed.heightFromName
                }
                formulas.cy = `J${rowNum}*H${rowNum}/27`
            } else if (htSubType === 'wall' || htSubType === 'slope_transition') {
                // House trap wall/slope: I=C, J=I*H, L=J*G/27
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

        case 'rog':
            // ROG (Ramp on grade): same formulas as Patch SOG - J=C, L=J*H/27, H from name
            formulas.sqFt = `C${rowNum}`
            if (itemData.parsed?.heightFromName !== undefined) {
                formulas.height = itemData.parsed.heightFromName
            }
            formulas.cy = `J${rowNum}*H${rowNum}/27`
            break

        case 'stairs_on_grade':
            const sogStairsSubType = itemData.parsed?.itemSubType
            if (sogStairsSubType === 'stairs') {
                // Stairs on grade: F=11/12, H=7/12 or from name, J=C*G*F, L=J*H/27, M=C
                formulas.length = '11/12' // Formula: =11/12
                formulas.height = itemData.parsed?.heightFromName != null ? itemData.parsed.heightFromName : '7/12'
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
    processStrapBeamItems,
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
    processSumpPumpPitItems,
    processGreaseTrapItems,
    processHouseTrapItems,
    processMatSlabItems,
    processMudSlabFoundationItems,
    processSOGItems,
    processROGItems,
    processStairsOnGradeItems,
    processElectricConduitItems,
    generateFoundationFormulas
}
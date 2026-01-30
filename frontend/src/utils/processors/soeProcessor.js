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
export const processSupportingAngleItems = (rawDataRows, headers) => {
    const items = processGenericSoeItems(rawDataRows, headers, isSupportingAngle)
    // Group by groupKey (what's after @)
    const groups = new Map()
    items.forEach(item => {
        const key = item.parsed.groupKey || 'Other'
        if (!groups.has(key)) groups.set(key, [])
        groups.get(key).push(item)
    })
    return Array.from(groups.entries()).map(([name, items]) => ({ name, items }))
}
export const processPargingItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isParging)
export const processHeelBlockItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isHeelBlock)
export const processUnderpinningItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isUnderpinning)
export const processRockAnchorItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isRockAnchor)
export const processRockBoltItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isRockBolt)
export const processAnchorItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isAnchor)
export const processTieBackItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isTieBack)
export const processConcreteSoilRetentionPierItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isConcreteSoilRetentionPier)
export const processGuideWallItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isGuideWall)
export const processDowelBarItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isDowelBar)
export const processRockPinItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isRockPin)
export const processShotcreteItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isShotcrete)
export const processPermissionGroutingItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isPermissionGrouting)
export const processButtonItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isButton)
export const processRockStabilizationItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isRockStabilization)
export const processFormBoardItems = (rawDataRows, headers) => processGenericSoeItems(rawDataRows, headers, isFormBoard)

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
    
    console.log('=== Drilled Soldier Pile Proposal Text ===')
    console.log(proposalText)
    console.log(`Total Count: ${totalCount}`)
    console.log(`Diameter: ${diameter}"`)
    console.log(`Thickness: ${thickness}"`)
    console.log(`Average Height: ${avgHeight.toFixed(2)}' (${heightText})`)
    console.log(`Embedment: ${embedmentText}`)
    
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


import { isExteriorSideItem, parseExteriorSideHeight, isExteriorSidePitItem, parseExteriorSidePitDimensions, isNegativeSideWallItem, parseNegativeSideWallHeight, getExteriorSidePitRefKey, isNegativeSideSlabItem, extractWaterproofingGroupKey } from '../parsers/waterproofingParser.js'

/**
 * Processes Waterproofing - Exterior side items from raw data
 * Remark: "2nd value from the bracket is considered as height & 2 FT extra is added."
 * Formulas: I (FT) = C (Takeoff), H = parsed height, J (SQ FT) = H * I
 * @param {Array} rawDataRows - Rows from raw Excel (excluding header)
 * @param {Array} headers - Column headers
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 * @returns {Array} - Items with { particulars, takeoff, unit, parsed: { heightFromBracketPlus2 } }
 */
export const processExteriorSideItems = (rawDataRows, headers, tracker = null) => {
  const items = []
  const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
  const totalIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'total')
  const unitIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'units')
  const estimateIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'estimate')

  if (digitizerIdx === -1 || totalIdx === -1) return items

  rawDataRows.forEach((row, rowIndex) => {
    const particulars = row[digitizerIdx]
    const total = row[totalIdx]
    const takeoff = total !== '' && total !== null && total !== undefined ? parseFloat(total) : ''
    const unit = unitIdx >= 0 ? (row[unitIdx] || '') : ''

    if (estimateIdx >= 0 && row[estimateIdx]) {
      const estimateVal = String(row[estimateIdx]).trim()
      if (estimateVal !== 'Waterproofing') return
    }

    if (!isExteriorSideItem(particulars)) return

    const heightFromBracketPlus2 = parseExteriorSideHeight(particulars)
    const groupKey = extractWaterproofingGroupKey(particulars)

    items.push({
      particulars: particulars || '',
      takeoff,
      unit,
      groupKey,
      parsed: {
        heightFromBracketPlus2: heightFromBracketPlus2 != null ? heightFromBracketPlus2 : ''
      }
    })

    // Mark this row as used
    if (tracker) {
      tracker.markUsed(rowIndex)
    }
  })

  // Group items by their grouping key
  const groupMap = new Map()
  items.forEach(item => {
    const key = item.groupKey || 'OTHER'
    if (!groupMap.has(key)) {
      groupMap.set(key, [])
    }
    groupMap.get(key).push(item)
  })

  let groups = Array.from(groupMap.values())

  // Merge single-item groups if there are multiple
  const singleItemGroups = groups.filter(g => g.length === 1)
  const multiItemGroups = groups.filter(g => g.length > 1)

  if (singleItemGroups.length > 1) {
    const mergedItems = []
    singleItemGroups.forEach(group => {
      mergedItems.push(...group)
    })
    mergedItems.forEach(item => item.groupKey = 'MERGED')

    return [...multiItemGroups.flat(), ...mergedItems]
  }

  return groups.flat()
}

/**
 * Processes Waterproofing - Exterior side pit/wall items (height = 2nd value from bracket + 2 FT).
 * @param {Array} rawDataRows - Rows from raw Excel (excluding header)
 * @param {Array} headers - Column headers
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 * @returns {Array} - Items with { particulars, takeoff, unit, parsed: { heightFromBracketPlus2, firstValueFeet, heightRefKey } }
 */
export const processExteriorSidePitItems = (rawDataRows, headers, tracker = null) => {
  const items = []
  const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
  const totalIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'total')
  const unitIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'units')
  const estimateIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'estimate')

  if (digitizerIdx === -1 || totalIdx === -1) return items

  rawDataRows.forEach((row, rowIndex) => {
    const particulars = row[digitizerIdx]
    const total = row[totalIdx]
    const takeoff = total !== '' && total !== null && total !== undefined ? parseFloat(total) : ''
    const unit = unitIdx >= 0 ? (row[unitIdx] || '') : ''

    if (estimateIdx >= 0 && row[estimateIdx]) {
      const estimateVal = String(row[estimateIdx]).trim()
      if (estimateVal !== 'Waterproofing') return
    }

    if (!isExteriorSidePitItem(particulars)) return

    const { heightFromBracketPlus2, firstValueFeet, heightRefKey } = parseExteriorSidePitDimensions(particulars)
    items.push({
      particulars: particulars || '',
      takeoff,
      unit,
      parsed: {
        heightFromBracketPlus2,
        firstValueFeet: firstValueFeet != null && firstValueFeet > 0 ? firstValueFeet : null,
        heightRefKey
      }
    })

    // Mark this row as used
    if (tracker) {
      tracker.markUsed(rowIndex)
    }
  })

  return items
}

/**
 * Processes Waterproofing - Negative side wall items (height = 2nd value from bracket only; no G; L = J*G/27 only for Elev. pit wall and Detention tank wall).
 * @param {Array} rawDataRows - Rows from raw Excel (excluding header)
 * @param {Array} headers - Column headers
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 * @returns {Array} - Items with { particulars, takeoff, unit, parsed: { secondValueFeet, heightRefKey } }
 */
export const processNegativeSideWallItems = (rawDataRows, headers, tracker = null) => {
  const items = []
  const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
  const totalIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'total')
  const unitIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'units')
  const estimateIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'estimate')

  if (digitizerIdx === -1 || totalIdx === -1) return items

  rawDataRows.forEach((row, rowIndex) => {
    const particulars = row[digitizerIdx]
    const total = row[totalIdx]
    const takeoff = total !== '' && total !== null && total !== undefined ? parseFloat(total) : ''
    const unit = unitIdx >= 0 ? (row[unitIdx] || '') : 'FT'

    if (estimateIdx >= 0 && row[estimateIdx]) {
      const estimateVal = String(row[estimateIdx]).trim()
      if (estimateVal !== 'Waterproofing') return
    }

    if (!isNegativeSideWallItem(particulars)) return

    const secondValueFeet = parseNegativeSideWallHeight(particulars)
    const heightRefKey = getExteriorSidePitRefKey(particulars)
    items.push({
      particulars: particulars || '',
      takeoff,
      unit: unit || 'FT',
      parsed: { secondValueFeet, heightRefKey }
    })

    // Mark this row as used
    if (tracker) {
      tracker.markUsed(rowIndex)
    }
  })

  return items
}

/**
 * Processes Waterproofing - Negative side slab items (Unit SQ FT, J = C).
 * @param {Array} rawDataRows - Rows from raw Excel (excluding header)
 * @param {Array} headers - Column headers
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 * @returns {Array} - Items with { particulars, takeoff, unit }
 */
export const processNegativeSideSlabItems = (rawDataRows, headers, tracker = null) => {
  const items = []
  const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
  const totalIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'total')
  const unitIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'units')
  const estimateIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'estimate')

  if (digitizerIdx === -1 || totalIdx === -1) return items

  rawDataRows.forEach((row, rowIndex) => {
    const particulars = row[digitizerIdx]
    const total = row[totalIdx]
    const takeoff = total !== '' && total !== null && total !== undefined ? parseFloat(total) : ''
    const unit = unitIdx >= 0 ? (row[unitIdx] || '') : 'SQ FT'

    if (estimateIdx >= 0 && row[estimateIdx]) {
      const estimateVal = String(row[estimateIdx]).trim()
      if (estimateVal !== 'Waterproofing') return
    }

    if (!isNegativeSideSlabItem(particulars)) return

    const groupKey = extractWaterproofingGroupKey(particulars)

    items.push({
      particulars: particulars || '',
      takeoff,
      unit: unit || 'SQ FT',
      groupKey
    })

    // Mark this row as used
    if (tracker) {
      tracker.markUsed(rowIndex)
    }
  })

  // Group items by their grouping key
  const groupMap = new Map()
  items.forEach(item => {
    const key = item.groupKey || 'OTHER'
    if (!groupMap.has(key)) {
      groupMap.set(key, [])
    }
    groupMap.get(key).push(item)
  })

  let groups = Array.from(groupMap.values())

  // Merge single-item groups if there are multiple
  const singleItemGroups = groups.filter(g => g.length === 1)
  const multiItemGroups = groups.filter(g => g.length > 1)

  if (singleItemGroups.length > 1) {
    const mergedItems = []
    singleItemGroups.forEach(group => {
      mergedItems.push(...group)
    })
    mergedItems.forEach(item => item.groupKey = 'MERGED')

    return [...multiItemGroups.flat(), ...mergedItems]
  }

  return groups.flat()
}

/**
 * Formulas for Exterior side: I = C, J = H * I (SQ FT)
 * For pit items: height = secondValueFeet + 2; no width (G).
 * Column CY (L) only for Elev. pit wall and Detention tank wall: L = J*G/27 (G left empty).
 * @param {string} itemType - e.g. 'waterproofing_exterior_side', 'waterproofing_exterior_side_pit'
 * @param {number} rowNum - 1-based row number
 * @param {object} itemData - Item with parsed data
 * @param {object} [options] - Not used anymore (kept for compatibility)
 * @returns {object} - { ft, sqFt, height, heightFormula, cy }
 */
export const generateWaterproofingFormulas = (itemType, rowNum, itemData, options = {}) => {
  const formulas = {
    ft: null,
    sqFt: null,
    height: null,
    heightFormula: null,
    cy: null
  }

  if (itemType === 'waterproofing_exterior_side') {
    formulas.ft = `C${rowNum}`
    formulas.height = itemData?.parsed?.heightFromBracketPlus2
    formulas.sqFt = `H${rowNum}*I${rowNum}`
  }

  if (itemType === 'waterproofing_exterior_side_pit') {
    formulas.ft = `C${rowNum}`
    // Height is now static: 2nd bracket value + 2 feet
    formulas.height = itemData?.parsed?.heightFromBracketPlus2
    formulas.sqFt = `H${rowNum}*I${rowNum}`
    const heightRefKey = itemData?.parsed?.heightRefKey
    if (heightRefKey === 'elevatorPit' || heightRefKey === 'detentionTank') {
      formulas.cy = `J${rowNum}*G${rowNum}/27`
    }
  }

  if (itemType === 'waterproofing_negative_side_wall') {
    formulas.ft = `C${rowNum}`
    formulas.height = itemData?.parsed?.secondValueFeet
    formulas.sqFt = `I${rowNum}*H${rowNum}`
  }

  if (itemType === 'waterproofing_negative_side_slab') {
    formulas.sqFt = `C${rowNum}`
  }

  return formulas
}

export default {
  processExteriorSideItems,
  processExteriorSidePitItems,
  processNegativeSideWallItems,
  processNegativeSideSlabItems,
  generateWaterproofingFormulas
}

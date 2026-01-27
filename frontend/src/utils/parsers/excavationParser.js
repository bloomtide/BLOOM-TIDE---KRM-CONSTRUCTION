import { convertToFeet, extractDimensions } from './dimensionParser'

/**
 * Extracts height from patterns like "H=11'-0"" or "H=2'-4""
 * @param {string} text - Text containing height
 * @returns {number} - Height in feet
 */
export const extractHeightFromH = (text) => {
  if (!text) return 0

  const match = text.match(/H=([^)]+)/)
  if (match) {
    return convertToFeet(match[1])
  }
  return 0
}

/**
 * Extracts height from mud slab specification like "w/ 2" mud slab"
 * @param {string} text - Text containing mud slab specification
 * @returns {number} - Height in feet
 */
export const extractMudSlabHeight = (text) => {
  if (!text) return 0

  // Match patterns like "w/ 2" mud slab" or "w/ 2\" mud slab"
  const match = text.match(/w\/\s*(\d+)["']?\s*mud\s*slab/i)
  if (match) {
    const inches = parseFloat(match[1]) || 0
    return inches / 12 // Convert inches to feet
  }
  return 0
}

/**
 * Determines excavation item type
 * @param {string} digitizerItem - The digitizer item text
 * @param {string} subsection - Subsection context
 * @returns {string} - Type identifier
 */
export const getExcavationItemType = (digitizerItem, subsection = 'excavation') => {
  const itemLower = digitizerItem.toLowerCase()

  if (itemLower.includes('underground piping')) return 'underground_piping'

  // SF, WF, ST items - only use first value from bracket as width
  if (itemLower.match(/^sf\s*\(/)) return 'sf'
  if (itemLower.match(/^wf-\d+\s*\(/)) return 'wf'
  if (itemLower.match(/^st-\d+\s*\(/)) return 'st'

  // Heel block items
  if (itemLower.includes('heel block')) return 'heel_block'

  // PC items with numeric progression
  if (itemLower.match(/^pc-\d+\s*\(/)) return 'pc'

  // F items with numeric progression (F-1, F-2, ..., F-n)
  if (itemLower.match(/^f-\d+\s*\(/)) return 'f'

  // Prioritize excavation/backfill types in their respective subsections
  // This ensures items like "Slope exc & backfill ... w/ 2" mud slab" get the depth H from bracket
  if (subsection === 'excavation' || subsection === 'backfill') {
    // Exc items with H= pattern
    if (itemLower.match(/^exc\s*\(h=/)) return 'exc'

    // Slope exc & backfill items
    if (itemLower.includes('slope exc') || itemLower.includes('slope excavation')) return 'slope_exc'

    // Exc & backfill
    if (itemLower.includes('exc & backfill') || itemLower.includes('excavation & backfill')) return 'exc_backfill'

    // Backfill items (just "Backfill" with H=)
    if (itemLower.match(/^backfill\s*\(h=/)) return 'backfill'
  }

  // Mud slab items
  if (itemLower.match(/w\/\s*\d+["']?\s*mud\s*slab/)) return 'mud_slab'

  // If not handled by subsection priority, check these types anyway
  if (itemLower.match(/^exc\s*\(h=/)) return 'exc'
  if (itemLower.includes('slope exc') || itemLower.includes('slope excavation')) return 'slope_exc'
  if (itemLower.includes('exc & backfill') || itemLower.includes('excavation & backfill')) return 'exc_backfill'
  if (itemLower.match(/^backfill\s*\(h=/)) return 'backfill'

  // Duplex sewage ejector pit slab
  if (itemLower.includes('duplex sewage ejector pit slab')) return 'sewage_pit_slab'

  // Concrete pier
  if (itemLower.includes('concrete pier')) return 'concrete_pier'

  // Rock excavation
  if (itemLower.startsWith('rock excavation')) return 'rock_exc'

  // Sump pit
  if (itemLower === 'sump pit') return 'sump_pit'

  return 'other'
}

/**
 * Parses excavation items with special rules
 * @param {string} digitizerItem - Full digitizer item text
 * @param {number} total - Total from raw data
 * @param {string} unit - Unit from raw data
 * @param {string} itemType - Type of excavation item
 * @param {string} subsection - Subsection context ('excavation' or 'backfill')
 * @returns {object} - Parsed data
 */
export const parseExcavationItem = (digitizerItem, total, unit, itemType, subsection = 'excavation') => {
  const result = {
    particulars: digitizerItem,
    takeoff: total,
    unit: unit,
    qty: 0,
    length: 0,
    width: 0,
    height: 0,
    manualHeight: false
  }

  switch (itemType) {
    case 'underground_piping':
      // Width and height are constant
      // Excavation subsection: width=3, height=2.5
      // Backfill subsection: width=3, height=2
      result.width = 3
      result.height = subsection === 'backfill' ? 2 : 2.5
      break

    case 'sf':
    case 'wf':
    case 'st':
      // Extract only 1st value from bracket as Width
      // Depth/Height is manual input
      const sfDims = extractDimensions(digitizerItem)
      if (sfDims.length && sfDims.width) {
        // Two dimensions found - use first as width
        result.width = sfDims.length
      } else if (sfDims.width) {
        result.width = sfDims.width
      } else if (sfDims.length) {
        result.width = sfDims.length
      }
      result.height = 0 // Manual input
      result.manualHeight = true
      break

    case 'heel_block':
    case 'pc':
    case 'f':
      // Consider 1st & 2nd values from bracket as Length & Width
      // Height is manual input (ignore 3rd dimension in bracket)
      const dimensions = extractDimensions(digitizerItem)
      if (dimensions.length) result.length = dimensions.length
      if (dimensions.width) result.width = dimensions.width
      // Height is manual - will be input separately
      result.manualHeight = true
      break

    case 'slope_exc':
    case 'exc_backfill':
      // Extract height from H= pattern in brackets
      // Remark: "Wherever exc is mentioned in Excavation bucket list, consider H= from bracket"
      result.height = extractHeightFromH(digitizerItem)
      break

    case 'exc':
    case 'backfill':
      // Extract height from H= pattern for standalone Exc items
      result.height = extractHeightFromH(digitizerItem)
      break

    case 'mud_slab':
      // Extract height from mud slab specification (e.g., "w/ 2" mud slab")
      result.height = extractMudSlabHeight(digitizerItem)
      // QTY, Length, Width are manual inputs - leave empty
      result.manualHeight = false // Height is auto-calculated from mud slab spec
      break

    case 'sewage_pit_slab':
      // E, F, G, H remain empty
      result.qty = 0
      result.length = 0
      result.width = 0
      result.height = 0
      result.manualHeight = true
      break

    case 'concrete_pier':
      // Extract length and width from brackets (e.g., "(4'-0"x4'-0"x14'-0")")
      // Only extract 1st and 2nd dimensions
      const pierDims = extractDimensions(digitizerItem)
      if (pierDims.length) result.length = pierDims.length // Column F
      if (pierDims.width) result.width = pierDims.width   // Column G
      result.height = 0 // Column H is manual
      result.manualHeight = true
      break

    case 'rock_exc':
    case 'line_drilling':
      // Extract height from H= pattern
      result.height = extractHeightFromH(digitizerItem)
      break

    default:
      // Generic case
      const genericDims = extractDimensions(digitizerItem)
      if (genericDims.length) result.length = genericDims.length
      if (genericDims.width) result.width = genericDims.width
      if (genericDims.height) result.height = genericDims.height
  }

  // For EA items, set QTY (but leave column E empty)
  if (unit === 'EA') {
    result.qty = total
  }

  return result
}

export default {
  extractHeightFromH,
  getExcavationItemType,
  parseExcavationItem
}

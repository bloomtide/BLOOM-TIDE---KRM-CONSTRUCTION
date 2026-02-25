/**
 * Utility functions to parse dimensions from digitizer item strings
 */

/**
 * Converts a dimension string to feet
 * Examples: "4"" → 0.33, "2'-0"" → 2.0, "1'-6"" → 1.5
 * @param {string} dimStr - Dimension string like "4"", "2'-0"", "1'-6""
 * @returns {number} - Value in feet
 */
export const convertToFeet = (dimStr) => {
  if (!dimStr || typeof dimStr !== 'string') return 0

  dimStr = dimStr.trim()

  // Check if it's just inches (e.g., "4"")
  if (dimStr.includes('"') && !dimStr.includes("'")) {
    const inches = parseFloat(dimStr.replace('"', ''))
    return inches / 12
  }

  // Check if it's feet and inches (e.g., "2'-6"" or "1'-0"")
  if (dimStr.includes("'")) {
    const parts = dimStr.split("'")
    const feet = parseFloat(parts[0]) || 0

    if (parts[1]) {
      // Remove quotes and dashes, then parse inches
      let inchesStr = parts[1].replace('"', '').trim()
      // Remove leading dash if present (e.g., "-6" → "6")
      if (inchesStr.startsWith('-')) {
        inchesStr = inchesStr.substring(1)
      }
      const inches = parseFloat(inchesStr) || 0
      return feet + inches / 12
    }

    return feet
  }

  // If no units, assume feet
  return parseFloat(dimStr) || 0
}

/**
 * Extracts dimensions from parentheses in digitizer item.
 * When there are multiple parenthesis groups (e.g. "Pilaster (P3) (22\"x16\"x6'-0\")"),
 * uses the group that contains 'x' (dimension separator), same as foundation parseBracketDimensions.
 * Examples:
 * - "(2'-0"x1'-0")" → { width: 2, height: 1 }
 * - "(2'-0"x3'-0"x1'-6")" → { length: 2, width: 3, height: 1.5 }
 * - "Demo Pilaster (P3) (22\"x16\"x6'-0\")" → uses "(22\"x16\"x6'-0\")" → length, width, height
 * @param {string} text - Text containing dimensions
 * @returns {object} - Object with length, width, height (in feet)
 */
export const extractDimensions = (text) => {
  if (!text) return {}

  // Find all parenthetical groups (same logic as foundation parseBracketDimensions)
  const matches = Array.from(text.matchAll(/\(([^)]+)\)/g))
  if (!matches.length) return {}

  const candidates = matches.map(m => (m[1] || '').trim()).filter(Boolean)
  if (!candidates.length) return {}

  // Prefer the bracket that contains 'x' (dimension separator), e.g. "(22\"x16\"x6'-0\")" not "(P3)"
  const dimStr =
    [...candidates].reverse().find(c => c.toLowerCase().includes('x')) || candidates[0]

  // Split by 'x' (multiplication symbol)
  const parts = dimStr.split('x').map(p => p.trim())

  const result = {}

  if (parts.length === 2) {
    // 2 values: width x height
    result.width = convertToFeet(parts[0])
    result.height = convertToFeet(parts[1])
  } else if (parts.length === 3) {
    // 3 values: length x width x height
    result.length = convertToFeet(parts[0])
    result.width = convertToFeet(parts[1])
    result.height = convertToFeet(parts[2])
  }

  return result
}

/**
 * Extracts thickness from text like "SOG 4" thick" or "SOG 6" thick"
 * @param {string} text - Text containing thickness
 * @returns {number} - Thickness in feet
 */
export const extractThickness = (text) => {
  if (!text) return 0

  // Match patterns like "4" thick", '6" thick', etc.
  const match = text.match(/(\d+(?:\.\d+)?)\s*["']\s*thick/i)
  if (match) {
    const value = parseFloat(match[1])
    // Check if it has double quote (inches) or single quote (feet)
    if (text.includes('" thick')) {
      return value / 12 // Convert inches to feet
    } else if (text.includes("' thick")) {
      return value
    }
  }

  return 0
}

/**
 * Extracts quantity from digitizer item for EA units
 * For items like "Demo isolated footing", the quantity comes from the Total column
 * @param {number} total - Total value from raw data
 * @param {string} unit - Unit from raw data
 * @returns {number} - Quantity
 */
export const extractQuantity = (total, unit) => {
  if (unit === 'EA') {
    return total
  }
  return 0
}

/**
 * Parses a complete digitizer item for demolition
 * @param {string} digitizerItem - Full digitizer item text
 * @param {number} total - Total from raw data
 * @param {string} unit - Unit from raw data
 * @param {string} subsection - Subsection name
 * @returns {object} - Parsed data
 */
export const parseDemolitionItem = (digitizerItem, total, unit, subsection) => {
  const result = {
    particulars: digitizerItem,
    takeoff: total,
    unit: unit,
    qty: 0,  // Column E - always leave empty
    length: 0,  // Column F - always leave empty
    width: 0,   // Column G - always leave empty
    height: 0   // Column H - fill based on subsection
  }

  switch (subsection) {
    case 'Demo slab on grade':
    case 'Demo Ramp on grade':
      // Extract thickness from item name (e.g., "Demo SOG 4" thick", "Demo ROG 4" thick")
      // If thickness not mentioned, use 4" typ. (0.33 feet)
      const thickness = extractThickness(digitizerItem)
      result.height = thickness > 0 ? thickness : 4 / 12 // Default to 4" = 0.33 feet
      // Columns E, F, G remain empty
      break

    case 'Demo strip footing':
      // Extract dimensions from brackets (e.g., "(2'-0"x1'-0")"); same as Foundation: SF/WF vs ST
      const sfDimensions = extractDimensions(digitizerItem)
      if (sfDimensions.width) result.width = sfDimensions.width  // Column G
      if (sfDimensions.height) result.height = sfDimensions.height  // Column H
      const itemLower = (digitizerItem || '').toLowerCase()
      if (itemLower.startsWith('st ') || /^st\s*\(/i.test(itemLower) || itemLower.includes('strap')) {
        result.itemType = 'ST'
      } else if (itemLower.startsWith('wf') || itemLower.includes('wall footing')) {
        result.itemType = 'WF'
      } else {
        result.itemType = 'SF'
      }
      break

    case 'Demo foundation wall':
    case 'Demo retaining wall':
      // Extract dimensions from brackets (e.g., "(1'-0"x3'-0")")
      // First value is width, second value is height
      const fwDimensions = extractDimensions(digitizerItem)
      if (fwDimensions.width) result.width = fwDimensions.width  // Column G
      if (fwDimensions.height) result.height = fwDimensions.height  // Column H
      // Columns E, F remain empty
      break

    case 'Demo isolated footing':
      // Extract dimensions from brackets (e.g., "(2'-0"x3'-0"x1'-6")")
      const footingDimensions = extractDimensions(digitizerItem)
      if (footingDimensions.length) result.length = footingDimensions.length
      if (footingDimensions.width) result.width = footingDimensions.width
      if (footingDimensions.height) result.height = footingDimensions.height
      break

    // ── Pile caps, pilaster, buttress, pier, corbel: L×W×H (same as isolated footing) ──
    case 'Demo pile caps':
    case 'Demo pilaster':
    case 'Demo buttress':
    case 'Demo pier':
    case 'Demo corbel':
      {
        const dims = extractDimensions(digitizerItem)
        if (dims.length) result.length = dims.length
        if (dims.width) result.width = dims.width
        if (dims.height) result.height = dims.height
      }
      break

    // ── Grade beam, tie beam, strap beam, liner wall, barrier wall, stem wall: W×H (linear; takeoff = length) ──
    case 'Demo grade beam':
    case 'Demo tie beam':
    case 'Demo strap beam':
    case 'Demo liner wall':
    case 'Demo barrier wall':
    case 'Demo stem wall':
      {
        const dims = extractDimensions(digitizerItem)
        if (dims.width) result.width = dims.width
        if (dims.height) result.height = dims.height
      }
      break

    // ── Thickened slab, mat slab: thickness as height (SQ FT = takeoff) ──
    case 'Demo thickened slab':
    case 'Demo mat slab':
      {
        const thickness = extractThickness(digitizerItem)
        result.height = thickness > 0 ? thickness : 4 / 12
      }
      break

    // ── Mud slab ──
    case 'Demo mud slab':
      {
        const thickness = extractThickness(digitizerItem)
        result.height = thickness > 0 ? thickness : 4 / 12
      }
      break

    // ── Pit items: L×W×H (or W×H; use same extractDimensions) ──
    case 'Demo elevator pit':
    case 'Demo service elevator pit':
    case 'Demo detention tank':
    case 'Demo duplex sewage ejector pit':
    case 'Demo deep sewage ejector pit':
    case 'Demo sewage ejector pit':
    case 'Demo sump pump pit':
    case 'Demo grease trap pit':
    case 'Demo house trap pit':
      {
        const dims = extractDimensions(digitizerItem)
        if (dims.length) result.length = dims.length
        if (dims.width) result.width = dims.width
        if (dims.height) result.height = dims.height
      }
      break

    case 'Demo stair on grade':
      // Pass-through: particulars, takeoff, unit - no dimension extraction
      break

    default:
      break
  }

  return result
}

export default {
  convertToFeet,
  extractDimensions,
  extractThickness,
  extractQuantity,
  parseDemolitionItem
}
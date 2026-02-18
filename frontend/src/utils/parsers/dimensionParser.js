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
 * Extracts dimensions from parentheses in digitizer item
 * Examples:
 * - "(2'-0"x1'-0")" → { width: 2, height: 1 }
 * - "(2'-0"x3'-0"x1'-6")" → { length: 2, width: 3, height: 1.5 }
 * @param {string} text - Text containing dimensions
 * @returns {object} - Object with length, width, height (in feet)
 */
export const extractDimensions = (text) => {
  if (!text) return {}

  // Find content within parentheses
  const match = text.match(/\(([^)]+)\)/)
  if (!match) return {}

  const dimStr = match[1]

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
      // Extract dimensions from brackets (e.g., "(2'-0"x1'-0")")
      // First value is width, second value is height
      const sfDimensions = extractDimensions(digitizerItem)
      if (sfDimensions.width) result.width = sfDimensions.width  // Column G
      if (sfDimensions.height) result.height = sfDimensions.height  // Column H
      // Columns E, F remain empty
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
      // First value is length, second is width, third is height
      const footingDimensions = extractDimensions(digitizerItem)
      if (footingDimensions.length) result.length = footingDimensions.length  // Column F
      if (footingDimensions.width) result.width = footingDimensions.width  // Column G
      if (footingDimensions.height) result.height = footingDimensions.height  // Column H
      // Column E remains empty
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
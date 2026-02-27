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
 * Extracts thickness from text like "SOG 4" thick", "6" thk", etc.
 * @param {string} text - Text containing thickness
 * @returns {number} - Thickness in feet
 */
export const extractThickness = (text) => {
  if (!text) return 0

  // Match patterns like "4" thick", "6" thk", etc. (thick or thk)
  const match = text.match(/(\d+(?:\.\d+)?)\s*["']\s*(?:thick|thk)/i)
  if (match) {
    const value = parseFloat(match[1])
    // Check if it has double quote (inches) or single quote (feet)
    if (text.includes('" thick') || text.includes('" thk')) {
      return value / 12 // Convert inches to feet
    } else if (text.includes("' thick") || text.includes("' thk")) {
      return value
    }
  }

  return 0
}

/**
 * Extracts height (in feet) from pit/slab names like "House trap pit slab 12"" or "slab 8" typ."
 * Same pattern as Foundation: (\d+)" with optional "typ." - no "thick" required.
 * @param {string} text - Item name
 * @returns {number} - Height in feet, or 0 if not found
 */
export const extractInchesFromName = (text) => {
  if (!text) return 0
  // Same as Foundation parseHouseTrap/parseSumpPumpPit slab: (\d+)" optional typ.
  const match = text.match(/(\d+(?:\.\d+)?)\s*["']\s*(?:typ\.)?/i)
  if (match) {
    const inches = parseFloat(match[1])
    return inches / 12 // Convert inches to feet (e.g. "12"" -> 1)
  }
  return 0
}

/**
 * Normalizes "thk" to "thick" for UI/display. Raw data may contain "thk"; show as "thick" everywhere.
 * @param {string} text - Text that may contain "thk" (e.g. "8\" thk", "6\" thk.")
 * @returns {string} - Same text with word "thk" replaced by "thick"
 */
export const normalizeThickForDisplay = (text) => {
  if (text == null || typeof text !== 'string') return text
  return text.replace(/\bthk\.?/gi, 'thick');
}

/**
 * Normalizes unit from raw data. No. can be EA, No, or No. — all treated as EA for downstream use.
 * Use this when reading unit from row[unitIdx] so that all sections treat No/No. like EA.
 * @param {string} unit - Raw unit value (e.g. "EA", "No", "No.", "SF", "LF")
 * @returns {string} - Normalized unit ("EA" for No/No./EA, otherwise trimmed original)
 */
export const normalizeUnit = (unit) => {
  if (unit == null || unit === '') return ''
  const u = String(unit).trim()
  const norm = u.replace(/\.$/, '').toUpperCase()
  if (norm === 'EA' || norm === 'NO') return 'EA'
  return u
}

/**
 * Extracts quantity from digitizer item for EA units
 * For items like "Demo isolated footing", the quantity comes from the Total column
 * @param {number} total - Total value from raw data
 * @param {string} unit - Unit from raw data
 * @returns {number} - Quantity
 */
export const extractQuantity = (total, unit) => {
  // No. can be EA, No, or No.
  const unitNorm = (unit && String(unit).trim().replace(/\.$/, '').toUpperCase())
  if (unitNorm === 'EA' || unitNorm === 'NO') {
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
      // Same as Foundation SOG/ROG: height from "4" thick", "6"" or "4" thick" — match Foundation parseSOG/parseROG
      const sogRogThickness = extractThickness(digitizerItem) || extractInchesFromName(digitizerItem)
      result.height = sogRogThickness > 0 ? sogRogThickness : 4 / 12 // Default to 4" = 0.33 feet
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

    // ── Thickened slab: same as Foundation — width & height from brackets (2 values), or thickness from name ──
    case 'Demo thickened slab':
      {
        const tsDims = extractDimensions(digitizerItem)
        if (tsDims.width != null) result.width = tsDims.width
        if (tsDims.height != null) result.height = tsDims.height
        if (!result.height) {
          const th = extractThickness(digitizerItem) || extractInchesFromName(digitizerItem)
          result.height = th > 0 ? th : 4 / 12
        }
      }
      break

    // ── Mat slab: same as Foundation — thickness as height from "24"" or "8" thick" ──
    case 'Demo mat slab':
      {
        const matThickness = extractThickness(digitizerItem) || extractInchesFromName(digitizerItem)
        result.height = matThickness > 0 ? matThickness : 4 / 12
      }
      break

    // ── Mud slab: same as Foundation — thickness from "4"" or "mud mat 4"" ──
    case 'Demo mud slab':
      {
        const mudThickness = extractThickness(digitizerItem) || extractInchesFromName(digitizerItem)
        result.height = mudThickness > 0 ? mudThickness : 4 / 12
      }
      break

    // ── Pit items: same sub-types and formulas as Foundation (slab/mat/mat_slab → J=C, L=J*H/27; wall/slope → I=C, J=I*H, L=J*G/27) ──
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
        const pitLower = (digitizerItem || '').toLowerCase()
        // Sub-type for formula alignment with Foundation (slab/mat/wall/slope get different formulas)
        // Height for slab/mat: same as Foundation — from bracket dims, or from "12"" / "8" typ." in name (extractInchesFromName), or extractThickness ("12" thick")
        const slabHeightFromName = () => {
          if (dims.height) return dims.height
          const fromInches = extractInchesFromName(digitizerItem)
          if (fromInches > 0) return fromInches
          const th = extractThickness(digitizerItem)
          return th > 0 ? th : 0
        }
        if (pitLower.includes('mat slab')) {
          result.itemSubType = 'mat_slab'
          if (!result.height) result.height = slabHeightFromName()
        } else if (pitLower.includes('mat') && !pitLower.includes('mat slab')) {
          result.itemSubType = 'mat'
          if (!result.height) result.height = slabHeightFromName()
        } else if (subsection === 'Demo detention tank' && (pitLower.includes('lid') || pitLower.includes('lid slab'))) {
          result.itemSubType = 'lid_slab'
          if (!result.height) result.height = slabHeightFromName()
        } else if ((subsection === 'Demo elevator pit' || subsection === 'Demo service elevator pit') && pitLower.includes('sump')) {
          result.itemSubType = 'sump_pit'
        } else if (pitLower.includes('slope') || pitLower.includes('haunch') || pitLower.includes('transition')) {
          result.itemSubType = 'slope_transition'
          if (dims.width) result.width = dims.width
          if (dims.height) result.height = dims.height
          if (dims.width != null && dims.height != null) result.groupKey = `${Number(dims.width).toFixed(2)}x${Number(dims.height).toFixed(2)}`
        } else if (pitLower.includes('wall')) {
          result.itemSubType = 'wall'
          if (dims.width) result.width = dims.width
          if (dims.height) result.height = dims.height
          if (dims.width != null && dims.height != null) result.groupKey = `${Number(dims.width).toFixed(2)}x${Number(dims.height).toFixed(2)}`
        } else if (pitLower.includes('slab')) {
          result.itemSubType = 'slab'
          if (!result.height) result.height = slabHeightFromName()
        } else {
          // Default: treat as slab (J=C, L=J*H/27) when no keyword found
          result.itemSubType = 'slab'
          if (!result.height) result.height = slabHeightFromName()
        }
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
  parseDemolitionItem,
  normalizeThickForDisplay
}
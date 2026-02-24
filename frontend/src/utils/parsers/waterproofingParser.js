import { convertToFeet } from './dimensionParser.js'

/**
 * Parser for Waterproofing section - Exterior side subsection
 * Remark: "2nd value from the bracket is considered as height & 2 FT extra is added."
 */

/**
 * Checks if particulars matches Exterior side item patterns:
 * FW (...), RW (...), Vehicle barrier wall (...), Concrete liner wall (...), Stem wall (...)
 * @param {string} particulars - Item description
 * @returns {boolean}
 */
export const isExteriorSideItem = (particulars) => {
  if (!particulars || typeof particulars !== 'string') return false
  const p = particulars.trim()
  if (p.startsWith('FW (') || p.startsWith('FW(')) return true
  if (p.startsWith('RW (') || p.startsWith('RW(')) return true
  if (p.toLowerCase().includes('vehicle barrier wall (')) return true
  if (p.toLowerCase().includes('concrete liner wall (')) return true
  if (p.toLowerCase().includes('stem wall (')) return true
  return false
}

/**
 * Parses height for Exterior side: 2nd value from bracket (in feet) + 2 FT
 * Examples: (1'-0"x10'-0") -> 10+2=12; (10"x12'-0") -> 12+2=14; (8"x2'-6") -> 2.5+2=4.5
 * @param {string} particulars - Item description containing bracket dimensions
 * @returns {number|null} - Height in feet for column H, or null if not parseable
 */
export const parseExteriorSideHeight = (particulars) => {
  if (!particulars || typeof particulars !== 'string') return null
  const match = particulars.match(/\(([^)]+)\)/)
  if (!match) return null
  const inner = match[1]
  const parts = inner.split('x').map((s) => s.trim())
  if (parts.length < 2) return null
  const secondValueFeet = convertToFeet(parts[1])
  if (secondValueFeet === 0 && parts[1] !== '0') return null
  return secondValueFeet + 2
}

/**
 * Mapping from item name patterns to foundation slab ref keys.
 * Using regex patterns to support flexible naming (pit is optional, abbreviations like elev/elev./elevator)
 */
const PIT_PATTERN_MAP = {
  deepSewageEjectorPit: /deep\s+sewage\s+ejector(?:\s+pit)?\s+wall/i,
  elevatorPit: /(elev\.?|elevator)(?:\s+pit)?\s+wall/i,
  detentionTank: /detention\s+tank\s+wall/i,
  duplexSewageEjectorPit: /duplex\s+sewage\s+ejector(?:\s+pit)?\s+wall/i,
  greaseTrap: /grease\s+trap(?:\s+pit)?\s+(wall|slab)/i,
  houseTrap: /house\s+trap(?:\s+pit)?\s+(wall|slab)/i
}

/**
 * Checks if particulars is an Exterior side pit/wall item that uses Foundation slab height.
 * Excludes: Deep sewage ejector pit slab 12", Grease trap pit slab 12", House trap pit slab 12"
 * @param {string} particulars - Item description
 * @returns {boolean}
 */
export const isExteriorSidePitItem = (particulars) => {
  if (!particulars || typeof particulars !== 'string') return false
  const p = particulars.trim().toLowerCase()
  if (p.includes('slab')) return false
  return Object.values(PIT_PATTERN_MAP).some(pattern => pattern.test(p))
}

/**
 * Returns foundation slab ref key for pit item, or null.
 * @param {string} particulars - Item description
 * @returns {string|null}
 */
export const getExteriorSidePitRefKey = (particulars) => {
  if (!particulars || typeof particulars !== 'string') return null
  const p = particulars.trim().toLowerCase()
  for (const [refKey, pattern] of Object.entries(PIT_PATTERN_MAP)) {
    if (pattern.test(p)) return refKey
  }
  return null
}

/**
 * Parses dimensions for Exterior side pit items: 2nd value from bracket + 2 feet = height, 1st = width (feet) if present.
 * Height = 2nd value + 2 FT (same as other exterior side items).
 * @param {string} particulars - Item description containing bracket dimensions e.g. (6"x8'-6"), (1'-0"x5'-0")
 * @returns {{ heightFromBracketPlus2: number, firstValueFeet: number|null, heightRefKey: string|null }}
 */
export const parseExteriorSidePitDimensions = (particulars) => {
  const heightRefKey = getExteriorSidePitRefKey(particulars)
  const match = particulars.match(/\(([^)]+)\)/)
  if (!match) return { heightFromBracketPlus2: 0, firstValueFeet: null, heightRefKey }
  const inner = match[1]
  const parts = inner.split('x').map((s) => s.trim())
  if (parts.length < 2) return { heightFromBracketPlus2: 0, firstValueFeet: null, heightRefKey }
  const secondValueFeet = convertToFeet(parts[1])
  const firstValueFeet = convertToFeet(parts[0])
  const heightFromBracketPlus2 = secondValueFeet + 2
  return { heightFromBracketPlus2, firstValueFeet, heightRefKey }
}

/**
 * Negative side subsection
 * Wall items: same patterns as Exterior pit (Elev. pit wall, Detention tank wall, etc.) - 2nd value from bracket = height, no G, J = I*H, L = J*G/27 only for Elev. pit wall and Detention tank wall.
 * Slab items: House trap pit slab 12", Grease trap pit slab 12", Deep sewage ejector pit slab 12", Duplex sewage ejector pit slab 8" typ., Detention tank lid slab 8", Detention tank slab 12", Elev. pit slab (H=3'-0") (1.60') - Unit SQ FT, J = C.
 */

/** Negative side wall: same patterns as Exterior pit (no slab). */
export const isNegativeSideWallItem = (particulars) => isExteriorSidePitItem(particulars)

/** Parses height for Negative side wall: 2nd value from bracket only (no +2 feet). */
export const parseNegativeSideWallHeight = (particulars) => {
  if (!particulars || typeof particulars !== 'string') return null
  const match = particulars.match(/\(([^)]+)\)/)
  if (!match) return null
  const inner = match[1]
  const parts = inner.split('x').map((s) => s.trim())
  if (parts.length < 2) return null
  const secondValueFeet = convertToFeet(parts[1])
  return secondValueFeet
}

/** Negative side slab items: contain "slab" and match known patterns. Supports flexible naming (pit is optional). */
export const isNegativeSideSlabItem = (particulars) => {
  if (!particulars || typeof particulars !== 'string') return false
  const p = particulars.trim().toLowerCase()
  if (!p.includes('slab')) return false
  if (/house\s+trap(?:\s+pit)?\s+slab/i.test(p)) return true
  if (/grease\s+trap(?:\s+pit)?\s+slab/i.test(p)) return true
  if (/deep\s+sewage\s+ejector(?:\s+pit)?\s+slab/i.test(p)) return true
  if (/duplex\s+sewage\s+ejector(?:\s+pit)?\s+slab/i.test(p)) return true
  if (/detention\s+tank\s+lid\s+slab/i.test(p)) return true
  if (/detention\s+tank(?:\s+pit)?\s+slab(?!\s+lid)/i.test(p)) return true
  if (/(elev\.?|elevator)(?:\s+pit)?\s+slab/i.test(p)) return true
  return false
}

/**
 * Extracts grouping key for waterproofing items
 * Groups by thickness for slab items (e.g., "Detention tank lid slab 8"" -> groups by 8")
 * Groups by first bracket value for wall items (e.g., "FW (1'-0"x10'-0")" -> groups by 1'-0")
 * @param {string} particulars - Item description
 * @returns {string} - Grouping key
 */
export const extractWaterproofingGroupKey = (particulars) => {
  if (!particulars || typeof particulars !== 'string') return 'OTHER'

  const p = particulars.trim()
  const pLower = p.toLowerCase()

  // For slab items with thickness at the end (e.g., "Detention tank lid slab 8"")
  if (pLower.includes('slab')) {
    const thickMatch = p.match(/(\d+)["']?\s*$/)
    if (thickMatch) {
      return `THICK_${thickMatch[1]}`
    }
  }

  // For wall items with bracket dimensions, group by first value
  if (p.includes('(') && p.includes('x')) {
    const bracketMatch = p.match(/\(([^x)]+)/)
    if (bracketMatch) {
      return `DIM_${bracketMatch[1].trim()}`
    }
  }

  return 'OTHER'
}

export default {
  isExteriorSideItem,
  parseExteriorSideHeight,
  isExteriorSidePitItem,
  getExteriorSidePitRefKey,
  parseExteriorSidePitDimensions,
  isNegativeSideWallItem,
  parseNegativeSideWallHeight,
  isNegativeSideSlabItem,
  extractWaterproofingGroupKey
}
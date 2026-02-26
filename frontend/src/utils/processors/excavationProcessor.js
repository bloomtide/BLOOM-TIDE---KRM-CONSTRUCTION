import { parseExcavationItem, getExcavationItemType } from '../parsers/excavationParser'
import {
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
  isElevatorPit,
  isServiceElevatorPit,
  isDetentionTank,
  isDuplexSewageEjectorPit,
  isDeepSewageEjectorPit,
  isSewageEjectorPit,
  isSumpPumpPit,
  isGreaseTrap,
  isHouseTrap,
  isMatSlab,
  isMudSlabFoundation,
  isSOG,
  isROG,
  isStairsOnGrade,
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
  parseElevatorPit,
  parseServiceElevatorPit,
  parseDetentionTank,
  parseDuplexSewageEjectorPit,
  parseDeepSewageEjectorPit,
  parseSewageEjectorPit,
  parseSumpPumpPit,
  parseGreaseTrap,
  parseHouseTrap,
  parseMatSlab,
  parseMudSlabFoundation,
  parseSOG,
  parseROG,
  parseStairsOnGrade
} from '../parsers/foundationParser'

/**
 * Identifies if a digitizer item belongs to Excavation section
 * @param {string} digitizerItem - The digitizer item text
 * @returns {boolean} - True if it's an excavation item
 */
export const isExcavationItem = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return false

  const itemLower = digitizerItem.toLowerCase()

  // Check for excavation keywords
  const excavationKeywords = [
    'underground piping',
    /^sf\s*\(/,
    /^wf-\d+/,
    /^st-\d+/,
    'heel block',
    /^pc-\d+/,
    /^f-\d+\s*\(/,
    /^exc\s*\(/,
    'slope exc',
    'exc & backfill',
    'duplex sewage ejector pit slab',
    'sewage ejector pit slab'
  ]

  // Exclude gravel items and pure backfill items from excavation subsection
  if (itemLower.includes('gravel')) return false
  if (itemLower.startsWith('backfill')) return false

  // Exclude rock excavation items (they belong to Rock Excavation section)
  if (itemLower.includes('rock excavation')) return false
  if (itemLower.includes('concrete pier')) return false

  return excavationKeywords.some(keyword => {
    if (typeof keyword === 'string') {
      return itemLower.includes(keyword)
    } else {
      return keyword.test(itemLower)
    }
  })
}

/**
 * Identifies if a digitizer item belongs to Backfill subsection
 * @param {string} digitizerItem - The digitizer item text
 * @returns {boolean} - True if it's a backfill item
 */
export const isBackfillItem = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return false

  const itemLower = digitizerItem.toLowerCase()

  // Check for backfill keywords
  const backfillKeywords = [
    'underground piping',
    /^backfill\s*\(/,
    'slope exc',
    'exc & backfill'
  ]

  return backfillKeywords.some(keyword => {
    if (typeof keyword === 'string') {
      return itemLower.includes(keyword)
    } else {
      return keyword.test(itemLower)
    }
  })
}

/**
 * Identifies if a digitizer item belongs to Mud slab subsection
 * @param {string} digitizerItem - The digitizer item text
 * @returns {boolean} - True if it's a mud slab item
 */
export const isMudSlabItem = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return false

  const itemLower = digitizerItem.toLowerCase()

  // Check for mud slab pattern (e.g., "w/ 2" mud slab")
  return /w\/\s*\d+["']?\s*mud\s*slab/.test(itemLower)
}

/**
 * Generates formulas for excavation items
 * @param {string} itemType - Type of excavation item
 * @param {number} rowNum - Excel row number (1-based)
 * @param {object} parsedData - Parsed data with dimensions
 * @returns {object} - Formula strings for each calculated column
 */
export const generateExcavationFormulas = (itemType, rowNum, parsedData) => {
  const formulas = {
    ft: null,      // Column I
    sqFt: null,    // Column J
    lbs: null,     // Column K (will have CY)
    cy: null,      // Column L (will have 1.3*CY or J*H/27 for backfill)
    qtyFinal: null // Column M
  }

  const { unit, width, height, length, subsection } = parsedData
  const isBackfill = subsection === 'backfill'

  switch (itemType) {
    case 'underground_piping':
      // SQ FT = Takeoff * Width
      formulas.sqFt = `C${rowNum}*G${rowNum}`

      if (isBackfill) {
        // Backfill: Column K empty, Column L = J*H/27
        formulas.cy = `J${rowNum}*H${rowNum}/27`
      } else {
        // Excavation: CY = SQ FT * Height / 27, 1.3*CY = CY * 1.3
        formulas.lbs = `J${rowNum}*H${rowNum}/27`  // CY in column K
        formulas.cy = `K${rowNum}*1.3`              // 1.3*CY in column L
      }
      break

    case 'sf':
    case 'wf':
    case 'st':
      // SQ FT = Takeoff * Width, CY = SQ FT * Height / 27, 1.3*CY = CY * 1.3
      // Height column is empty but formulas will work when height is manually entered
      formulas.sqFt = `C${rowNum}*G${rowNum}`
      formulas.lbs = `J${rowNum}*H${rowNum}/27`  // CY in column K
      formulas.cy = `K${rowNum}*1.3`              // 1.3*CY in column L
      break

    case 'heel_block':
    case 'pc':
    case 'f':
      // SQ FT = Length * Width * Takeoff, CY = SQ FT * Height / 27, 1.3*CY = CY * 1.3
      if (unit === 'EA') {
        formulas.sqFt = `F${rowNum}*G${rowNum}*C${rowNum}`
        formulas.lbs = `J${rowNum}*H${rowNum}/27`  // CY in column K
        formulas.cy = `K${rowNum}*1.3`              // 1.3*CY in column L
        // Leave QTY column M empty for excavation items
      }
      break

    case 'slope_exc':
    case 'exc_backfill':
      // SQ FT = Takeoff
      formulas.sqFt = `C${rowNum}`

      if (isBackfill) {
        // Backfill: Column K empty, Column L = J*H/27
        formulas.cy = `J${rowNum}*H${rowNum}/27`
      } else {
        // Excavation subsection: 
        // formulas now include the 1.3 multiplier as requested
        formulas.lbs = `J${rowNum}*H${rowNum}/27` // Column K: CY
        formulas.cy = `K${rowNum}*1.3`             // Column L: 1.3*CY
      }
      break

    case 'exc':
    case 'backfill':
    case 'backfill_item': // generic backfill
      // SQ FT = Takeoff
      formulas.sqFt = `C${rowNum}`

      if (isBackfill) {
        // Backfill subsection: Column K empty, Column L = J*H/27
        formulas.lbs = null
        formulas.cy = `J${rowNum}*H${rowNum}/27`
      } else {
        // Standard Excavation: CY in K, 1.3*CY in L
        formulas.lbs = `J${rowNum}*H${rowNum}/27`
        formulas.cy = `K${rowNum}*1.3`
      }
      break

    case 'mud_slab':
      // Mud slab: Column J = Takeoff * 1.2, Column L = J*H/27
      // QTY, Length, Width are manual inputs (columns E, F, G)
      formulas.sqFt = `C${rowNum}*1.2`  // Column J = C * 1.2
      formulas.cy = `J${rowNum}*H${rowNum}/27`  // Column L = J*H/27
      break

    case 'sewage_pit_slab':
      // SQ FT = Takeoff, CY = J*H/27, 1.3*CY = K*1.3
      formulas.sqFt = `C${rowNum}` // Column J: =C
      formulas.lbs = `J${rowNum}*H${rowNum}/27` // Column K: CY
      formulas.cy = `K${rowNum}*1.3` // Column L: 1.3*CY
      break

    default:
      // Generic calculation
      if (unit === 'SQ FT' || unit === 'SF') {
        formulas.sqFt = `C${rowNum}`
        if (height > 0) {
          formulas.lbs = `J${rowNum}*H${rowNum}/27`
          formulas.cy = `K${rowNum}*1.3`
        }
      }
  }

  return formulas
}

/**
 * Processes all excavation items from raw data
 * @param {Array} rawDataRows - Array of rows from raw Excel data (excluding header)
 * @param {Array} headers - Column headers from raw data
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 * @returns {Array} - Processed excavation items with formulas
 */
export const processExcavationItems = (rawDataRows, headers, tracker = null) => {
  const excavationItems = []

  // Find column indices
  const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
  const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
  const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

  if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) {
    return excavationItems
  }

  // Process each row
  rawDataRows.forEach((row, rowIndex) => {
    const digitizerItem = row[digitizerIdx]
    const total = parseFloat(row[totalIdx]) || 0
    const unit = row[unitIdx]

    if (isExcavationItem(digitizerItem)) {
      const itemType = getExcavationItemType(digitizerItem, 'excavation')
      const parsed = parseExcavationItem(digitizerItem, total, unit, itemType, 'excavation')

      const itemData = {
        ...parsed,
        itemType,
        subsection: 'excavation', // Mark as excavation subsection
        rawRow: row,
        rawRowNumber: rowIndex + 2 // +2 because: +1 for header row, +1 for 1-based indexing
      }

      // Aggregate sewage_pit_slab items
      if (itemType === 'sewage_pit_slab') {
        const existingItem = excavationItems.find(item =>
          item.itemType === 'sewage_pit_slab' &&
          item.particulars === digitizerItem
        )
        if (existingItem) {
          existingItem.takeoff += total
          // Mark this row as used even though we're aggregating
          if (tracker) {
            tracker.markUsed(rowIndex)
          }
          return // Skip adding new item, already aggregated
        }
      }

      // Mark this row as used
      if (tracker) {
        tracker.markUsed(rowIndex)
      }

      excavationItems.push(itemData)
    }
  })

  return excavationItems
}

/**
 * Processes all backfill items from raw data
 * @param {Array} rawDataRows - Array of rows from raw Excel data (excluding header)
 * @param {Array} headers - Column headers from raw data
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 * @returns {Array} - Processed backfill items with formulas
 */
export const processBackfillItems = (rawDataRows, headers, tracker = null) => {
  const backfillItems = []

  // Find column indices
  const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
  const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
  const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

  if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) {
    return backfillItems
  }

  // Process each row
  rawDataRows.forEach((row, rowIndex) => {
    const digitizerItem = row[digitizerIdx]
    const total = parseFloat(row[totalIdx]) || 0
    const unit = row[unitIdx]

    if (isBackfillItem(digitizerItem)) {
      const itemType = getExcavationItemType(digitizerItem, 'backfill')
      const parsed = parseExcavationItem(digitizerItem, total, unit, itemType, 'backfill')

      backfillItems.push({
        ...parsed,
        itemType,
        subsection: 'backfill', // Mark as backfill subsection
        rawRow: row,
        rawRowNumber: rowIndex + 2 // +2 because: +1 for header row, +1 for 1-based indexing
      })

      // Mark this row as used
      if (tracker) {
        tracker.markUsed(rowIndex)
      }
    }
  })

  return backfillItems
}

/**
 * Processes all mud slab items from raw data
 * @param {Array} rawDataRows - Array of rows from raw Excel data (excluding header)
 * @param {Array} headers - Column headers from raw data
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 * @returns {Array} - Processed mud slab items with formulas
 */
export const processMudSlabItems = (rawDataRows, headers, tracker = null) => {
  const mudSlabItems = []

  // Find column indices
  const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
  const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
  const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

  if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) {
    return mudSlabItems
  }

  // Process each row
  rawDataRows.forEach((row, rowIndex) => {
    const digitizerItem = row[digitizerIdx]
    const total = parseFloat(row[totalIdx]) || 0
    const unit = row[unitIdx]

    if (isMudSlabItem(digitizerItem)) {
      const itemType = getExcavationItemType(digitizerItem, 'mud_slab')
      const parsed = parseExcavationItem(digitizerItem, total, unit, itemType, 'mud_slab')

      mudSlabItems.push({
        ...parsed,
        itemType,
        subsection: 'mud_slab', // Mark as mud slab subsection
        rawRow: row,
        rawRowNumber: rowIndex + 2 // +2 because: +1 for header row, +1 for 1-based indexing
      })

      // Mark this row as used
      if (tracker) {
        tracker.markUsed(rowIndex)
      }
    }
  })

  return mudSlabItems
}

/**
 * Excavation foundation-type: exclude items that contain "wall" (e.g. linear wall, pit wall, elev. pit wall).
 * Items like "elev. pit mat slab", "elev slab" do not contain "wall" and are included.
 */
const excavationExcludeWall = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return true
  return digitizerItem.toLowerCase().includes('wall')
}

/**
 * Exclude rows that belong to other sections (BPP Alternate, Civil asphalt, etc.).
 * Only Foundation subsection types (pile cap, strip footing, pilaster, SOG, pits, etc.) should be added to Excavation.
 */
const excavationExcludeNonFoundationSections = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return true
  const itemLower = digitizerItem.toLowerCase()
  if (itemLower.includes('bpp alternate') || itemLower.includes('- bpp ') || itemLower.includes(' bpp ')) return true
  if (itemLower.includes('full depth asphalt pavement') || itemLower.includes('full depth asphalt')) return true
  if (itemLower.includes('surface course') && itemLower.includes('base course')) return true
  if (itemLower.includes('asphalt pavement') && (itemLower.includes('surface course') || itemLower.includes('base course') || itemLower.includes('thick'))) return true
  if (itemLower.includes('bpp concrete sidewalk') || itemLower.includes('bpp concrete driveway') || itemLower.includes('bpp concrete curb')) return true
  if (itemLower.includes('bpp expansion joint') || itemLower.includes('bpp roadway')) return true
  return false
}

/**
 * Collects foundation-type items from UNUSED rows for the Excavation subsection.
 * Same types as Demolition/Foundation (pile caps, strip footings, pilaster, grade beam, pits, SOG, etc.)
 * but excludes: demolition items, and any item whose name contains "wall" (e.g. linear wall, pit wall).
 * Renders in one Excavation subsection with Foundation-style calculations.
 * Does NOT mark rows as used, so the same rows also appear in the Foundation section (e.g. Deep sewage ejector pit).
 * Call before Foundation so items are collected for Excavation; Foundation will then process the same rows for its subsections.
 * @param {Array} rawDataRows
 * @param {Array} headers
 * @param {UsedRowTracker} tracker
 * @returns {Array} Items with { particulars, takeoff, unit, parsed, foundationType, rawRowNumber }
 */
export const processExcavationFoundationTypeItems = (rawDataRows, headers, tracker = null) => {
  const items = []
  const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
  const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
  const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')
  if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) return items

  rawDataRows.forEach((row, rowIndex) => {
    if (tracker && tracker.isUsed(rowIndex)) return
    const digitizerItem = row[digitizerIdx]
    const total = parseFloat(row[totalIdx]) || 0
    const unit = row[unitIdx]
    if (!digitizerItem || typeof digitizerItem !== 'string') return
    const itemLower = digitizerItem.toLowerCase()
    if (itemLower.startsWith('demo ') || itemLower.startsWith('demo.') || itemLower.startsWith('demolition ') || itemLower.startsWith('remove ')) return
    if (excavationExcludeWall(digitizerItem)) return
    if (excavationExcludeNonFoundationSections(digitizerItem)) return

    let foundationType = null
    let parsed = {}

    if (isPileCap(digitizerItem)) { foundationType = 'pile_cap'; parsed = parsePileCap(digitizerItem) }
    else if (isStripFooting(digitizerItem)) { foundationType = 'strip_footing'; parsed = parseStripFooting(digitizerItem) }
    else if (isIsolatedFooting(digitizerItem)) { foundationType = 'isolated_footing'; parsed = parseIsolatedFooting(digitizerItem) }
    else if (isPilaster(digitizerItem)) { foundationType = 'pilaster'; parsed = parsePilaster(digitizerItem) }
    else if (isGradeBeam(digitizerItem)) { foundationType = 'grade_beam'; parsed = parseGradeBeam(digitizerItem) }
    else if (isTieBeam(digitizerItem)) { foundationType = 'tie_beam'; parsed = parseTieBeam(digitizerItem) }
    else if (isStrapBeam(digitizerItem)) { foundationType = 'strap_beam'; parsed = parseStrapBeam(digitizerItem) }
    else if (isThickenedSlab(digitizerItem)) { foundationType = 'thickened_slab'; parsed = parseThickenedSlab(digitizerItem) }
    else if (isButtress(digitizerItem)) { foundationType = 'buttress_takeoff'; parsed = {} }
    else if (isPier(digitizerItem)) { foundationType = 'pier'; parsed = parsePier(digitizerItem) }
    else if (isCorbel(digitizerItem)) { foundationType = 'corbel'; parsed = parseCorbel(digitizerItem) }
    else if (isElevatorPit(digitizerItem)) { foundationType = 'elevator_pit'; parsed = parseElevatorPit(digitizerItem) }
    else if (isServiceElevatorPit(digitizerItem)) { foundationType = 'service_elevator_pit'; parsed = parseServiceElevatorPit(digitizerItem) }
    else if (isDetentionTank(digitizerItem)) { foundationType = 'detention_tank'; parsed = parseDetentionTank(digitizerItem) }
    else if (isDuplexSewageEjectorPit(digitizerItem)) { foundationType = 'duplex_sewage_ejector_pit'; parsed = parseDuplexSewageEjectorPit(digitizerItem) }
    else if (isDeepSewageEjectorPit(digitizerItem)) { foundationType = 'deep_sewage_ejector_pit'; parsed = parseDeepSewageEjectorPit(digitizerItem) }
    else if (isSewageEjectorPit(digitizerItem)) { foundationType = 'sewage_ejector_pit'; parsed = parseSewageEjectorPit(digitizerItem) }
    else if (isSumpPumpPit(digitizerItem)) { foundationType = 'sump_pump_pit'; parsed = parseSumpPumpPit(digitizerItem) }
    else if (isGreaseTrap(digitizerItem)) { foundationType = 'grease_trap'; parsed = parseGreaseTrap(digitizerItem) }
    else if (isHouseTrap(digitizerItem)) { foundationType = 'house_trap'; parsed = parseHouseTrap(digitizerItem) }
    else if (isMatSlab(digitizerItem)) { foundationType = 'mat_slab'; parsed = parseMatSlab(digitizerItem) }
    else if (isMudSlabFoundation(digitizerItem)) { foundationType = 'mud_slab_foundation'; parsed = parseMudSlabFoundation(digitizerItem) }
    else if (isSOG(digitizerItem)) { foundationType = 'sog'; parsed = parseSOG(digitizerItem) }
    else if (isROG(digitizerItem)) { foundationType = 'rog'; parsed = parseROG(digitizerItem) }
    else if (isStairsOnGrade(digitizerItem)) { foundationType = 'stairs_on_grade'; parsed = parseStairsOnGrade(digitizerItem) }

    if (!foundationType) return

    items.push({
      particulars: digitizerItem,
      takeoff: total,
      unit: unit || '',
      parsed,
      foundationType,
      rawRowNumber: rowIndex + 2,
      subsection: 'excavation'
    })
    // Do NOT mark as used — same row should also appear in Foundation section (e.g. Deep sewage ejector pit subsection)
  })

  return items
}

export default {
  isExcavationItem,
  getExcavationItemType,
  generateExcavationFormulas,
  processExcavationItems,
  processExcavationFoundationTypeItems
}

import { parseDemolitionItem } from '../parsers/dimensionParser'

/**
 * Identifies if a digitizer item belongs to Demolition section
 * @param {string} digitizerItem - The digitizer item text
 * @returns {boolean} - True if it's a demolition item
 */
export const isDemolitionItem = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return false

  const itemLower = digitizerItem.toLowerCase()
  return itemLower.startsWith('demo ')
}

/**
 * Determines which demolition subsection an item belongs to
 * @param {string} digitizerItem - The digitizer item text
 * @returns {string|null} - Subsection name or null
 */
export const getDemolitionSubsection = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return null

  const itemLower = digitizerItem.toLowerCase()

  // Demo slab on grade
  if (itemLower.includes('demo sog')) {
    return 'Demo slab on grade'
  }

  // Demo Ramp on grade
  if (itemLower.includes('demo rog')) {
    return 'Demo Ramp on grade'
  }

  // Demo strip footing
  if (itemLower.includes('demo sf')) {
    return 'Demo strip footing'
  }

  // Demo foundation wall
  if (itemLower.includes('demo fw')) {
    return 'Demo foundation wall'
  }

  // Demo retaining wall
  if (itemLower.includes('demo rw')) {
    return 'Demo retaining wall'
  }

  // Demo isolated footing
  if (itemLower.includes('demo isolated footing')) {
    return 'Demo isolated footing'
  }

  // Demo stair on grade (stairs and landings)
  if (itemLower.includes('demo stairs on grade') || itemLower.includes('demo stair landings on grade') || itemLower.includes('demo stair landing on grade')) {
    return 'Demo stair on grade'
  }

  return null
}

/**
 * Generates formulas for demolition items based on subsection and row number
 * @param {string} subsection - Subsection name
 * @param {number} rowNum - Excel row number (1-based)
 * @param {object} parsedData - Parsed data with dimensions
 * @returns {object} - Formula strings for each calculated column
 */
export const generateDemolitionFormulas = (type, rowNum, parsedData) => {
  const formulas = {
    ft: null,      // Column I
    sqFt: null,    // Column J
    lbs: null,     // Column K
    cy: null,      // Column L
    qtyFinal: null // Column M
  }

  // Handle both subsection names and explicit item types
  switch (type) {
    case 'Demo slab on grade':
    case 'Demo Ramp on grade':
    case 'demo_extra_sqft':
    case 'demo_extra_rog_sqft':
      // SQ FT (J) = Takeoff, CY (L) = SQ FT * Height / 27
      formulas.sqFt = `C${rowNum}`
      formulas.cy = `J${rowNum}*H${rowNum}/27`
      break

    case 'Demo strip footing':
    case 'Demo foundation wall':
    case 'Demo retaining wall':
    case 'demo_extra_ft':
    case 'demo_extra_rw':
      // SQ FT (J) = Takeoff * Width (G)
      // Note: For Demo foundation wall / Demo retaining wall, C*G for J, J*H/27 for L
      formulas.sqFt = `C${rowNum}*G${rowNum}`
      formulas.cy = `J${rowNum}*H${rowNum}/27`
      break

    case 'Demo isolated footing':
    case 'demo_extra_ea':
      // SQ FT (J) = Length (F) * Width (G) * Takeoff (C)
      formulas.sqFt = `F${rowNum}*G${rowNum}*C${rowNum}`
      formulas.cy = `J${rowNum}*H${rowNum}/27`
      formulas.qtyFinal = `C${rowNum}`
      break

    default:
      break
  }

  return formulas
}

/**
 * Processes all demolition items from raw data, grouped by subsection
 * @param {Array} rawDataRows - Array of rows from raw Excel data (excluding header)
 * @param {Array} headers - Column headers from raw data
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 * @returns {object} - Object with subsection names as keys and arrays of items as values
 */
export const processDemolitionItems = (rawDataRows, headers, tracker = null) => {
  const demolitionItemsBySubsection = {
    'Demo slab on grade': [],
    'Demo Ramp on grade': [],
    'Demo strip footing': [],
    'Demo foundation wall': [],
    'Demo retaining wall': [],
    'Demo isolated footing': [],
    'Demo stair on grade': []
  }

  // Find column indices
  const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
  const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
  const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

  if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) {
    return demolitionItemsBySubsection
  }

  // Process each row
  rawDataRows.forEach((row, rowIndex) => {
    const digitizerItem = row[digitizerIdx]
    const total = parseFloat(row[totalIdx]) || 0
    const unit = row[unitIdx]

    if (isDemolitionItem(digitizerItem)) {
      const subsection = getDemolitionSubsection(digitizerItem)

      if (subsection && demolitionItemsBySubsection[subsection] !== undefined) {
        const parsed = parseDemolitionItem(digitizerItem, total, unit, subsection)

        // Extract grouping key based on subsection type
        let groupKey = 'DEFAULT'
        const itemLower = (digitizerItem || '').toLowerCase()

        if (subsection === 'Demo stair on grade') {
          // Demo stair on grade: group by text after @, or "NO_AT" if no @
          const atMatch = digitizerItem.match(/@\s*(.+)$/i)
          groupKey = atMatch ? atMatch[1].trim() : 'NO_AT'
          parsed.itemSubType = itemLower.includes('landings') ? 'landings' : 'stairs'
        } else if ((subsection === 'Demo slab on grade' || subsection === 'Demo Ramp on grade') && itemLower.includes('"')) {
          // Group by thickness for Demo SOG / Demo ROG
          const thickMatch = digitizerItem.match(/(\d+)["']?\s*thick/i)
          if (thickMatch) {
            groupKey = `THICK_${thickMatch[1]}`
          }
        } else if ((subsection === 'Demo strip footing' || subsection === 'Demo foundation wall' || subsection === 'Demo retaining wall' || subsection === 'Demo isolated footing') && digitizerItem.includes('(')) {
          // Group by first bracket value
          const bracketMatch = digitizerItem.match(/\(([^x)]+)/)
          if (bracketMatch) {
            groupKey = `DIM_${bracketMatch[1].trim()}`
          }
        }

        demolitionItemsBySubsection[subsection].push({
          ...parsed,
          subsection,
          groupKey,
          rawRow: row,
          rawRowNumber: rowIndex + 2 // +2 because: +1 for header row, +1 for 1-based indexing
        })

        // Mark this row as used
        if (tracker) {
          tracker.markUsed(rowIndex)
        }
      }
    }
  })

  // Group items within each subsection by their grouping key
  Object.keys(demolitionItemsBySubsection).forEach(subsection => {
    const items = demolitionItemsBySubsection[subsection]
    if (items.length === 0) return

    // Demo stair on grade: build groups { heading, stairs, landings } - do not merge
    if (subsection === 'Demo stair on grade') {
      const groupMap = new Map() // key -> { heading, stairs, landings }
      items.forEach(item => {
        const key = item.groupKey || 'NO_AT'
        const heading = key !== 'NO_AT' ? key : null
        if (!groupMap.has(key)) {
          groupMap.set(key, { heading, stairs: null, landings: null })
        }
        const g = groupMap.get(key)
        if (item.itemSubType === 'stairs') {
          g.stairs = item
        } else if (item.itemSubType === 'landings') {
          g.landings = item
        }
      })
      demolitionItemsBySubsection[subsection] = Array.from(groupMap.values()).filter(g => g.stairs || g.landings)
      return
    }

    const groupMap = new Map()
    items.forEach(item => {
      const key = item.groupKey || 'DEFAULT'
      if (!groupMap.has(key)) {
        groupMap.set(key, [])
      }
      groupMap.get(key).push(item)
    })

    // Convert to grouped array
    const groups = Array.from(groupMap.values())

    // Merge single-item groups if there are multiple
    const singleItemGroups = groups.filter(g => g.length === 1)
    const multiItemGroups = groups.filter(g => g.length > 1)

    if (singleItemGroups.length > 1) {
      const mergedItems = []
      singleItemGroups.forEach(group => {
        mergedItems.push(...group)
      })
      mergedItems.forEach(item => item.groupKey = 'MERGED')

      demolitionItemsBySubsection[subsection] = [...multiItemGroups.flat(), ...mergedItems]
    } else {
      demolitionItemsBySubsection[subsection] = groups.flat()
    }
  })

  return demolitionItemsBySubsection
}

export default {
  isDemolitionItem,
  getDemolitionSubsection,
  generateDemolitionFormulas,
  processDemolitionItems
}
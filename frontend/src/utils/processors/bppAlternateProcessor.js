/**
 * Processes B.P.P. Alternate #2 scope section items from raw data.
 * Items are organized by street name (e.g., "West Street", "Maple Avenue")
 * 
 * Subsections:
 * - Gravel (manual entry, 2 rows per street for 4" and 6" gravel)
 * - Concrete sidewalk
 * - Concrete driveway
 * - Concrete curb
 * - Concrete flush curb
 * - Expansion joint
 * - Conc road base (manual entry)
 * - Full depth asphalt pavement
 */

/**
 * Identifies if a digitizer item belongs to B.P.P. Alternate #2 scope section
 * @param {string} digitizerItem - The digitizer item text
 * @returns {boolean}
 */
export const isBPPAlternateItem = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return false
  const itemLower = digitizerItem.toLowerCase()
  return itemLower.includes('- bpp ')
}

/**
 * Extracts street name from item (e.g., "West Street - BPP ..." -> "West Street")
 * @param {string} digitizerItem - The digitizer item text
 * @returns {string|null}
 */
export const extractStreetName = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return null
  const match = digitizerItem.match(/^(.+?)\s*-\s*BPP\s/i)
  return match ? match[1].trim() : null
}

/**
 * Determines which B.P.P. subsection an item belongs to
 * @param {string} digitizerItem - The digitizer item text
 * @returns {string|null} - Subsection name or null
 */
export const getBPPSubsection = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return null
  const itemLower = digitizerItem.toLowerCase()

  if (itemLower.includes('bpp concrete sidewalk')) return 'Concrete sidewalk'
  if (itemLower.includes('bpp concrete driveway')) return 'Concrete driveway'
  if (itemLower.includes('bpp concrete curb') || itemLower.includes('bpp concrete flush curb') || itemLower.includes('bpp concrete drop curb')) return 'Concrete curb'
  if (itemLower.includes('bpp expansion joint')) return 'Expansion joint'
  if (itemLower.includes('bpp full depth asphalt') || itemLower.includes('bpp asphalt') || itemLower.includes('bpp roadway')) return 'Full depth asphalt pavement'

  return null
}

/**
 * Parses thickness from item name like "4" thick" -> 4
 * @param {string} particulars - Digitizer item text
 * @returns {number|null}
 */
const parseThicknessInches = (particulars) => {
  // Match patterns like "4" thick" or "7" thick"
  const match = particulars.match(/(\d+(?:\.\d+)?)\s*"\s*thick/i)
  return match ? parseFloat(match[1]) : null
}

/**
 * Parses width and height from curb item like "8" wide, Height=1'-6""
 * @param {string} particulars - Digitizer item text
 * @returns {{ widthInches: number, heightFeet: number }|null}
 */
const parseCurbDimensions = (particulars) => {
  // Match width like "8" wide" or "6" wide"
  const widthMatch = particulars.match(/(\d+(?:\.\d+)?)\s*"\s*wide/i)
  // Match height like "Height=1'-6"" -> 1.5 feet
  const heightMatch = particulars.match(/Height\s*=\s*(\d+)'\s*-?\s*(\d+)"/i)

  if (!widthMatch || !heightMatch) return null

  const widthInches = parseFloat(widthMatch[1])
  const heightFeet = parseInt(heightMatch[1]) + parseInt(heightMatch[2]) / 12

  return { widthInches, heightFeet }
}

/**
 * Parses asphalt pavement dimensions from item name
 * "1.5" thick surface course + 3" thick base course+ 6" thick gravel"
 * @param {string} particulars - Digitizer item text
 * @returns {{ surfaceInches: number, baseInches: number, gravelInches: number }|null}
 */
const parseAsphaltDimensions = (particulars) => {
  const surfaceMatch = particulars.match(/(\d+(?:\.\d+)?)\s*"\s*(?:thick\s+)?surface/i)
  const baseMatch = particulars.match(/(\d+(?:\.\d+)?)\s*"\s*(?:thick\s+)?base/i)
  const gravelMatch = particulars.match(/(\d+(?:\.\d+)?)\s*"\s*(?:thick\s+)?gravel/i)

  return {
    surfaceInches: surfaceMatch ? parseFloat(surfaceMatch[1]) : 0,
    baseInches: baseMatch ? parseFloat(baseMatch[1]) : 0,
    gravelInches: gravelMatch ? parseFloat(gravelMatch[1]) : 0
  }
}

/**
 * Generates formulas for B.P.P. items based on subsection and row number
 * @param {string} itemType - Item type identifier
 * @param {number} rowNum - Excel row number (1-based)
 * @param {object} parsedData - Parsed data with dimensions
 * @returns {object} - Formula strings for each calculated column
 */
export const generateBPPFormulas = (itemType, rowNum, parsedData) => {
  const formulas = {
    ft: null,      // Column I
    sqFt: null,    // Column J
    lbs: null,     // Column K
    cy: null,      // Column L
    qtyFinal: null // Column M
  }

  switch (itemType) {
    case 'bpp_gravel':
      // Gravel: FT = Takeoff (C), CY = SQ_FT * Height / 27
      formulas.ft = `C${rowNum}`
      formulas.cy = `J${rowNum}*H${rowNum}/27`
      break

    case 'bpp_concrete_sidewalk':
    case 'bpp_concrete_driveway':
      // SQ FT items: FT = Takeoff (C), CY = SQ_FT * Height / 27
      formulas.ft = `C${rowNum}`
      formulas.cy = `J${rowNum}*H${rowNum}/27`
      break

    case 'bpp_concrete_curb':
    case 'bpp_concrete_flush_curb':
      // FT items: FT = Takeoff (C), SQ_FT = FT * Height, CY = FT * Width / 27
      formulas.ft = `C${rowNum}`
      formulas.sqFt = `I${rowNum}*H${rowNum}`
      formulas.cy = `I${rowNum}*G${rowNum}/27`
      break

    case 'bpp_expansion_joint':
      // Expansion joint: FT = Takeoff (C)
      formulas.ft = `C${rowNum}`
      break

    case 'bpp_conc_road_base':
      // Conc road base: SQ_FT = Takeoff (C), CY = SQ_FT * Height / 27
      formulas.sqFt = `C${rowNum}`
      formulas.cy = `J${rowNum}*H${rowNum}/27`
      break

    case 'bpp_full_depth_asphalt':
      // Full depth asphalt: FT = Takeoff (C), CY = SQ_FT * Height / 27
      formulas.ft = `C${rowNum}`
      formulas.cy = `J${rowNum}*H${rowNum}/27`
      break

    case 'bpp_sum':
      // Sum row formulas are applied separately
      break

    default:
      break
  }

  return formulas
}

/**
 * Processes all B.P.P. Alternate #2 scope items from raw data, grouped by street and subsection
 * @param {Array} rawDataRows - Array of rows from raw Excel data (excluding header)
 * @param {Array} headers - Column headers from raw data
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 * @returns {object} - Object with street names as keys, each containing subsection items
 */
export const processBPPAlternateItems = (rawDataRows, headers, tracker = null) => {
  const itemsByStreet = {}

  // Find column indices
  const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
  const totalIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'total')
  const unitIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'units')
  const estimateIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'estimate')

  if (digitizerIdx === -1 || totalIdx === -1) {
    return itemsByStreet
  }

  // Process each row
  rawDataRows.forEach((row, rowIndex) => {
    const digitizerItem = row[digitizerIdx]
    const total = row[totalIdx]
    const takeoff = total !== '' && total !== null && total !== undefined ? parseFloat(total) : ''
    const unit = unitIdx >= 0 ? (row[unitIdx] || '') : ''

    // Check if this is a BPP item (either by estimate column or by item name)
    if (estimateIdx >= 0 && row[estimateIdx]) {
      const estimateVal = String(row[estimateIdx]).trim()
      if (estimateVal !== 'B.P.P. Alternate #2 scope' && !isBPPAlternateItem(digitizerItem)) return
    } else if (!isBPPAlternateItem(digitizerItem)) {
      return
    }

    const streetName = extractStreetName(digitizerItem)
    const subsection = getBPPSubsection(digitizerItem)

    if (!streetName || !subsection) return

    // Initialize street object if needed
    if (!itemsByStreet[streetName]) {
      itemsByStreet[streetName] = {
        streetName,
        'Concrete sidewalk': [],
        'Concrete driveway': [],
        'Concrete curb': [],
        'Expansion joint': [],
        'Full depth asphalt pavement': []
      }
    }

    // Parse dimensions based on subsection
    let parsed = {}

    if (subsection === 'Concrete sidewalk' || subsection === 'Concrete driveway') {
      const thickness = parseThicknessInches(digitizerItem)
      parsed = {
        heightValue: thickness ? thickness / 12 : null,
        heightFormula: thickness ? `${thickness}/12` : null
      }
    } else if (subsection === 'Concrete curb' || subsection === 'Concrete flush curb') {
      const dims = parseCurbDimensions(digitizerItem)
      if (dims) {
        parsed = {
          widthValue: dims.widthInches / 12,
          widthFormula: `${dims.widthInches}/12`,
          heightValue: dims.heightFeet
        }
      }
    } else if (subsection === 'Full depth asphalt pavement') {
      const dims = parseAsphaltDimensions(digitizerItem)
      if (dims) {
        const totalHeight = dims.surfaceInches + dims.baseInches
        parsed = {
          heightValue: totalHeight / 12,
          heightFormula: `(${dims.surfaceInches}+${dims.baseInches})/12`,
          surfaceInches: dims.surfaceInches,
          baseInches: dims.baseInches,
          gravelInches: dims.gravelInches
        }
      }
    }

    // Determine unit
    let itemUnit = unit
    if (!itemUnit) {
      if (subsection === 'Concrete curb' || subsection === 'Concrete flush curb' || subsection === 'Expansion joint') {
        itemUnit = 'FT'
      } else {
        itemUnit = 'SQ FT'
      }
    }

    // Add item to appropriate subsection
    itemsByStreet[streetName][subsection].push({
      particulars: digitizerItem,
      takeoff,
      unit: itemUnit,
      parsed,
      subsection,
      streetName
    })

    // Mark this row as used
    if (tracker) {
      tracker.markUsed(rowIndex)
    }
  })

  return itemsByStreet
}

export default {
  isBPPAlternateItem,
  extractStreetName,
  getBPPSubsection,
  generateBPPFormulas,
  processBPPAlternateItems
}
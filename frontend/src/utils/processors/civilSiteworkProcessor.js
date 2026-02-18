/**
 * Processes Civil / Sitework section items from raw data.
 * 
 * Subsections:
 * - Demo
 *   - Demo asphalt
 *   - Demo curb
 *   - Demo fence
 *   - Demo wall
 *   - Demo pipe
 *   - Demo rail
 *   - Demo sign
 *   - Demo manhole
 *   - Demo fire hydrant
 *   - Demo utility pole
 *   - Demo valve
 *   - Demo inlet
 * - Excavation
 * - Gravel
 * - Concrete Pavement
 * - Asphalt
 * - Pads
 * - Soil Erosion
 * - Fence
 * - Concrete filled steel pipe bollard
 * - Site (Hydrant, Wheel stop, Drain, Protection, Signages, Main line)
 * - Ele (Excavation, Backfill, Gravel)
 * - Gas (Excavation, Backfill, Gravel)
 * - Water (Excavation, Backfill, Gravel)
 * - Drains & Utilities
 * - Alternate
 */

/**
 * Identifies if a digitizer item belongs to Civil / Sitework section
 * @param {string} digitizerItem - The digitizer item text
 * @param {string} estimate - The estimate column value
 * @returns {boolean}
 */
export const isCivilSiteworkItem = (digitizerItem, estimate) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return false

  // Check estimate column
  if (estimate && String(estimate).trim() === 'Civil / Sitework') return true

  const itemLower = digitizerItem.toLowerCase()

  // Demo items - exclude (Add/Alt) items for utility pole
  if (itemLower.includes('(add/alt)') && itemLower.includes('utility pole')) return false

  // Demo items
  if (itemLower.includes('remove existing') || itemLower.includes('protect existing') || itemLower.includes('relocate existing')) return true
  if (itemLower.startsWith('remove ') && (itemLower.includes('wall') || itemLower.includes('rail'))) return true

  return false
}

/**
 * Determines which Demo sub-subsection an item belongs to
 * @param {string} digitizerItem - The digitizer item text
 * @returns {string|null} - Sub-subsection name or null
 */
export const getDemoSubSubsection = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return null
  const itemLower = digitizerItem.toLowerCase()

  // Demo asphalt
  if (itemLower.includes('asphalt pavement') || (itemLower.includes('remove') && itemLower.includes('asphalt') && !itemLower.includes('curb'))) {
    return 'Demo asphalt'
  }

  // Demo curb
  if (itemLower.includes('curb')) {
    return 'Demo curb'
  }

  // Demo fence
  if (itemLower.includes('fence')) {
    return 'Demo fence'
  }

  // Demo wall
  if (itemLower.includes('wall') && !itemLower.includes('stormwater')) {
    return 'Demo wall'
  }

  // Demo pipe
  if (itemLower.includes('pipe') || itemLower.includes('hdpe') || itemLower.includes('rcp') || itemLower.includes('stormwater main')) {
    return 'Demo pipe'
  }

  // Demo rail
  if (itemLower.includes('rail') || itemLower.includes('guide rail')) {
    return 'Demo rail'
  }

  // Demo sign
  if (itemLower.includes('sign')) {
    return 'Demo sign'
  }

  // Demo manhole
  if (itemLower.includes('manhole')) {
    return 'Demo manhole'
  }

  // Demo fire hydrant (includes relocate)
  if (itemLower.includes('fire hydrant') || itemLower.includes('relocate existing fire hydrant')) {
    return 'Demo fire hydrant'
  }

  // Demo utility pole - exclude (Add/Alt) items
  if ((itemLower.includes('utility pole') || (itemLower.includes('pole') && !itemLower.includes('hydrant'))) &&
    !itemLower.includes('(add/alt)') && !itemLower.startsWith('(add/alt)')) {
    return 'Demo utility pole'
  }

  // Demo valve
  if (itemLower.includes('valve')) {
    return 'Demo valve'
  }

  // Demo inlet
  if (itemLower.includes('inlet')) {
    return 'Demo inlet'
  }

  return null
}

/**
 * Determines the sign type for grouping
 * @param {string} digitizerItem - The digitizer item text
 * @returns {string} - Sign type identifier
 */
export const getSignType = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return 'other'
  const itemLower = digitizerItem.toLowerCase()

  if (itemLower.includes('row of sign')) {
    return 'row_of_signs'
  }
  return 'single_sign'
}

/**
 * Determines the inlet type for grouping
 * @param {string} digitizerItem - The digitizer item text
 * @returns {string} - Inlet type identifier
 */
export const getInletType = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return 'other'
  const itemLower = digitizerItem.toLowerCase()

  if (itemLower.includes('protect')) {
    return 'protect'
  }
  if (itemLower.includes('remove')) {
    return 'remove'
  }
  return 'other'
}

/**
 * Determines the fence type for grouping
 * @param {string} digitizerItem - The digitizer item text
 * @returns {string} - Fence type identifier
 */
export const getFenceType = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return 'other'
  const itemLower = digitizerItem.toLowerCase()

  if (itemLower.includes('chain link') || itemLower.includes('vinyl')) {
    return 'chain_link_vinyl'
  }
  if (itemLower.includes('wood')) {
    return 'wood'
  }
  return 'other'
}

/**
 * Determines the pipe type for grouping
 * @param {string} digitizerItem - The digitizer item text
 * @returns {string} - Pipe type identifier
 */
export const getPipeType = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return 'other'
  const itemLower = digitizerItem.toLowerCase()

  if (itemLower.includes('remove') && itemLower.includes('pipe')) {
    return 'remove_pipe'
  }
  if (itemLower.includes('protect')) {
    return 'protect'
  }
  return 'other'
}

/**
 * Generates formulas for Civil/Sitework Demo items
 * @param {string} itemType - Item type identifier
 * @param {number} rowNum - Excel row number (1-based)
 * @param {object} parsedData - Parsed data with dimensions
 * @returns {object} - Formula strings for each calculated column
 */
export const generateCivilDemoFormulas = (itemType, rowNum, parsedData) => {
  const formulas = {
    ft: null,      // Column I
    sqFt: null,    // Column J
    lbs: null,     // Column K
    cy: null,      // Column L
    qtyFinal: null // Column M
  }

  switch (itemType) {
    case 'civil_demo_asphalt':
      // SQ FT items: J = C, L = J * H / 27
      formulas.sqFt = `C${rowNum}`
      formulas.cy = `J${rowNum}*H${rowNum}/27`
      break

    case 'civil_demo_curb':
      // FT items: I = C, J = I * H, L = J * G / 27
      formulas.ft = `C${rowNum}`
      formulas.sqFt = `I${rowNum}*H${rowNum}`
      formulas.cy = `J${rowNum}*G${rowNum}/27`
      break

    case 'civil_demo_fence':
      // FT items: I = C, J = I * H
      formulas.ft = `C${rowNum}`
      formulas.sqFt = `I${rowNum}*H${rowNum}`
      break

    case 'civil_demo_wall':
      // FT items: I = C, J = I * H, L = J * G / 27
      formulas.ft = `C${rowNum}`
      formulas.sqFt = `I${rowNum}*H${rowNum}`
      formulas.cy = `J${rowNum}*G${rowNum}/27`
      break

    case 'civil_demo_pipe':
    case 'civil_demo_rail':
      // FT items: I = C
      formulas.ft = `C${rowNum}`
      break

    default:
      break
  }

  return formulas
}

/**
 * Processes all Civil / Sitework Demo items from raw data
 * @param {Array} rawDataRows - Array of rows from raw Excel data (excluding header)
 * @param {Array} headers - Column headers from raw data
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 * @returns {object} - Object with sub-subsection names as keys and arrays of items as values
 */
export const processCivilDemoItems = (rawDataRows, headers, tracker = null) => {
  const demoItems = {
    'Demo asphalt': [],
    'Demo curb': [],
    'Demo fence': { 'chain_link_vinyl': [], 'wood': [], 'other': [] },
    'Demo wall': [],
    'Demo pipe': { 'remove_pipe': [], 'protect': [], 'other': [] },
    'Demo rail': [],
    'Demo sign': { 'single_sign': [], 'row_of_signs': [] },
    'Demo manhole': [],
    'Demo fire hydrant': [],
    'Demo utility pole': [],
    'Demo valve': [],
    'Demo inlet': { 'protect': [], 'remove': [], 'other': [] }
  }

  // Find column indices
  const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
  const totalIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'total')
  const unitIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'units')
  const estimateIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'estimate')

  if (digitizerIdx === -1 || totalIdx === -1) {
    return demoItems
  }

  // Process each row
  rawDataRows.forEach((row, rowIndex) => {
    const digitizerItem = row[digitizerIdx]
    const total = row[totalIdx]
    const takeoff = total !== '' && total !== null && total !== undefined ? parseFloat(total) : ''
    const unit = unitIdx >= 0 ? (row[unitIdx] || '') : ''
    const estimate = estimateIdx >= 0 ? row[estimateIdx] : ''

    // Check if this is a Civil/Sitework Demo item
    if (!isCivilSiteworkItem(digitizerItem, estimate)) return

    const subSubsection = getDemoSubSubsection(digitizerItem)
    if (!subSubsection) return

    // Parse dimensions based on sub-subsection
    let parsed = {}
    let itemUnit = unit

    if (subSubsection === 'Demo asphalt') {
      parsed = { heightValue: 0.25 }
      itemUnit = itemUnit || 'SQ FT'
    } else if (subSubsection === 'Demo curb') {
      // Width: 8" = 0.67 ft, Height: 1'-6" = 1.5 ft
      parsed = { widthValue: 0.67, heightValue: 1.5 }
      itemUnit = itemUnit || 'FT'
    } else if (subSubsection === 'Demo fence') {
      // Height: 6 ft (typical for both chain link and wood)
      parsed = { heightValue: 6 }
      itemUnit = itemUnit || 'FT'
    } else if (subSubsection === 'Demo wall') {
      // Width: 1.5 ft, Height: 3.5 ft
      parsed = { widthValue: 1.5, heightValue: 3.5 }
      itemUnit = itemUnit || 'FT'
    } else if (subSubsection === 'Demo pipe' || subSubsection === 'Demo rail') {
      itemUnit = itemUnit || 'FT'
    } else if (subSubsection === 'Demo sign' || subSubsection === 'Demo manhole' ||
      subSubsection === 'Demo fire hydrant' || subSubsection === 'Demo utility pole' ||
      subSubsection === 'Demo valve' || subSubsection === 'Demo inlet') {
      // EA items - formula goes in column M
      itemUnit = itemUnit || 'EA'
    }

    const item = {
      particulars: digitizerItem,
      takeoff,
      unit: itemUnit,
      parsed,
      subSubsection
    }

    // Add to appropriate group
    if (subSubsection === 'Demo fence') {
      const fenceType = getFenceType(digitizerItem)
      demoItems['Demo fence'][fenceType].push(item)
    } else if (subSubsection === 'Demo pipe') {
      const pipeType = getPipeType(digitizerItem)
      demoItems['Demo pipe'][pipeType].push(item)
    } else if (subSubsection === 'Demo sign') {
      const signType = getSignType(digitizerItem)
      demoItems['Demo sign'][signType].push(item)
    } else if (subSubsection === 'Demo inlet') {
      const inletType = getInletType(digitizerItem)
      demoItems['Demo inlet'][inletType].push(item)
    } else if (demoItems[subSubsection] && !Array.isArray(demoItems[subSubsection])) {
      // This shouldn't happen but handle it
      console.warn(`Unexpected structure for ${subSubsection}`)
    } else if (demoItems[subSubsection]) {
      demoItems[subSubsection].push(item)
    }

    // Mark this row as used
    if (tracker) {
      tracker.markUsed(rowIndex)
    }
  })

  return demoItems
}

/**
 * Determines which Civil/Sitework subsection an item belongs to (non-Demo)
 * @param {string} digitizerItem - The digitizer item text
 * @returns {string|null} - Subsection name or null
 */
export const getCivilSubsection = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return null
  const itemLower = digitizerItem.toLowerCase()

  // Skip Demo items (handled separately)
  if (itemLower.includes('remove existing') || itemLower.includes('protect existing') ||
    itemLower.includes('relocate existing') || itemLower.startsWith('remove ')) {
    return null
  }

  // Fence subsection
  if (itemLower.includes('construction fence') ||
    (itemLower.includes('proposed fence') && itemLower.includes('height=')) ||
    itemLower.includes('proposed guiderail')) {
    return 'Fence'
  }

  // Soil Erosion subsection
  if (itemLower.includes('stabilized construction entrance') ||
    itemLower.includes('silt fence') ||
    itemLower.includes('inlet filter')) {
    return 'Soil Erosion'
  }

  // Pads subsection
  if (itemLower.includes('proposed transformer concrete pad')) {
    return 'Pads'
  }

  // Asphalt subsection (full depth asphalt pavement - not in gravel or excavation context)
  if (itemLower.includes('full depth asphalt pavement') && itemLower.includes('surface course')) {
    return 'Asphalt'
  }

  // Concrete Pavement subsection
  if (itemLower.includes('proposed reinforced concrete sidewalk')) {
    return 'Concrete Pavement'
  }

  // Gravel subsection - items that need gravel calculations
  // Note: These items appear in multiple sections - this is for the Gravel subsection specifically

  // Excavation subsection - items that need excavation calculations
  if (itemLower.includes('proposed') && (
    itemLower.includes('bollard') ||
    itemLower.includes('transformer concrete pad') ||
    itemLower.includes('reinforced concrete sidewalk') ||
    itemLower.includes('full depth asphalt pavement')
  )) {
    return 'Excavation'
  }

  return null
}

/**
 * Parses QTY from bracket like "(2 No.)" or "(1 No.)"
 * @param {string} particulars - Digitizer item text
 * @returns {number|null}
 */
export const parseQtyFromBracket = (particulars) => {
  const match = particulars.match(/\(\s*(\d+)\s*(?:No\.?|EA)?\s*\)/i)
  return match ? parseInt(match[1], 10) : null
}

/**
 * Parses height from name like "Height=6'-0"" or "Height=10'-0""
 * @param {string} particulars - Digitizer item text
 * @returns {number|null}
 */
export const parseHeightFromName = (particulars) => {
  // Match "Height=X'-Y"" pattern
  const match = particulars.match(/Height\s*=\s*(\d+)'\s*-?\s*(\d+)"/i)
  if (match) {
    return parseInt(match[1]) + parseInt(match[2]) / 12
  }
  // Match "Height=X'-Y" typ" pattern
  const match2 = particulars.match(/Height\s*=\s*(\d+)'\s*-?\s*(\d+)"\s*typ/i)
  if (match2) {
    return parseInt(match2[1]) + parseInt(match2[2]) / 12
  }
  return null
}

/**
 * Parses thickness from name like "6" thick" or "8" thick"
 * @param {string} particulars - Digitizer item text
 * @returns {number|null} - thickness in inches
 */
export const parseThicknessFromName = (particulars) => {
  const match = particulars.match(/(\d+(?:\.\d+)?)\s*"\s*thick/i)
  return match ? parseFloat(match[1]) : null
}

/**
 * Parses asphalt layer thicknesses from name
 * @param {string} particulars - Digitizer item text
 * @returns {{ surface: number, base: number }|null}
 */
export const parseAsphaltThicknesses = (particulars) => {
  const surfaceMatch = particulars.match(/(\d+(?:\.\d+)?)\s*"\s*(?:thick\s+)?surface/i)
  const baseMatch = particulars.match(/(\d+(?:\.\d+)?)\s*"\s*(?:thick\s+)?base/i)

  if (surfaceMatch && baseMatch) {
    return {
      surface: parseFloat(surfaceMatch[1]),
      base: parseFloat(baseMatch[1])
    }
  }
  return null
}

/**
 * Parses bollard dimensions from name like "(6" ∅, H=6'-6") w/ footing (18" ∅, H=4'-0")"
 * @param {string} particulars - Digitizer item text
 * @returns {{ bollardDia: number, bollardH: number, footingDia: number, footingH: number }|null}
 */
export const parseBollardDimensions = (particulars) => {
  // Match bollard diameter and height
  const bollardMatch = particulars.match(/\((\d+(?:\.\d+)?)\s*"\s*[∅Ø],\s*H\s*=\s*(\d+)'\s*-?\s*(\d+)"\)/i)
  // Match footing diameter and height
  const footingMatch = particulars.match(/footing\s*\((\d+(?:\.\d+)?)\s*"\s*[∅Ø],\s*H\s*=\s*(\d+)'\s*-?\s*(\d+)"\)/i)

  if (bollardMatch && footingMatch) {
    return {
      bollardDia: parseFloat(bollardMatch[1]),
      bollardH: parseInt(bollardMatch[2]) + parseInt(bollardMatch[3]) / 12,
      footingDia: parseFloat(footingMatch[1]),
      footingH: parseInt(footingMatch[2]) + parseInt(footingMatch[3]) / 12
    }
  }
  return null
}

/**
 * Parses silt fence height from name like "Height=2'-6""
 * @param {string} particulars - Digitizer item text
 * @returns {number|null}
 */
export const parseSiltFenceHeight = (particulars) => {
  const match = particulars.match(/Height\s*=\s*(\d+)'\s*-?\s*(\d+)"/i)
  if (match) {
    return parseInt(match[1]) + parseInt(match[2]) / 12
  }
  return null
}

/**
 * Gets the fence group type for Civil/Sitework Fence subsection
 * @param {string} digitizerItem - The digitizer item text
 * @returns {string}
 */
export const getCivilFenceType = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return 'other'
  const itemLower = digitizerItem.toLowerCase()

  if (itemLower.includes('construction fence')) {
    return 'construction_fence'
  }
  if (itemLower.includes('proposed fence') && itemLower.includes('height=')) {
    return 'proposed_fence'
  }
  if (itemLower.includes('guiderail')) {
    return 'guiderail'
  }
  return 'other'
}

/**
 * Processes Civil / Sitework non-Demo items (Excavation, Gravel, Concrete Pavement, Asphalt, Pads, Soil Erosion, Fence)
 * @param {Array} rawDataRows - Array of rows from raw Excel data (excluding header)
 * @param {Array} headers - Column headers from raw data
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 * @returns {object}
 */
export const processCivilOtherItems = (rawDataRows, headers, tracker = null) => {
  const items = {
    'Excavation': {
      'transformer_pad': [],
      'reinforced_sidewalk': [],
      'bollard': [],
      'asphalt': []
    },
    'Gravel': {
      'transformer_pad': [],
      'transformer_pad_8': [],
      'reinforced_sidewalk': [],
      'asphalt': []
    },
    'Concrete Pavement': [],
    'Asphalt': [],
    'Pads': [],
    'Soil Erosion': {
      'stabilized_entrance': [],
      'silt_fence': [],
      'inlet_filter': []
    },
    'Fence': {
      'construction_fence': [],
      'proposed_fence': [],
      'guiderail': []
    },
    'Concrete filled steel pipe bollard': {
      'footing': [],
      'simple': []
    },
    'Site': {
      'Hydrant': [],
      'Wheel stop': [],
      'Drain': {
        'Area': [],
        'Floor': []
      },
      'Protection': [],
      'Signages': [],
      'Main line': {
        'Gas': [],
        'Sanitary': [],
        'Water': []
      }
    },
    'Drains & Utilities': [],
    'Alternate': []
  }

  // Find column indices
  const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
  const totalIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'total')
  const unitIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'units')
  const estimateIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'estimate')
  const qtyIdx = headers.findIndex(h => {
    const t = h && String(h).toLowerCase().trim()
    return t === 'qty' || t === 'quantity'
  })

  if (digitizerIdx === -1 || totalIdx === -1) {
    return items
  }

  // Process each row
  rawDataRows.forEach((row, rowIndex) => {
    const digitizerItem = row[digitizerIdx]
    if (!digitizerItem || typeof digitizerItem !== 'string') return

    const total = row[totalIdx]
    const takeoff = total !== '' && total !== null && total !== undefined ? parseFloat(total) : ''
    const unit = unitIdx >= 0 ? (row[unitIdx] || '') : ''
    const estimate = estimateIdx >= 0 ? row[estimateIdx] : ''
    const qtyVal = qtyIdx >= 0 && row[qtyIdx] !== '' && row[qtyIdx] != null ? parseFloat(row[qtyIdx]) : null

    const itemLower = digitizerItem.toLowerCase()

    // Skip Demo items
    if (itemLower.includes('remove existing') || itemLower.includes('protect existing') ||
      itemLower.includes('relocate existing') || itemLower.startsWith('remove ')) {
      // Allow (Add/Alt) items to pass through
      if (!itemLower.includes('(add/alt)')) {
        return
      }
    }




    // Check if item belongs to Civil/Sitework by estimate column
    if (estimate && String(estimate).trim() !== 'Civil / Sitework') {
      // Also check for specific Civil/Sitework patterns if estimate doesn't match
      if (!itemLower.includes('proposed') && !itemLower.includes('construction fence')) {
        return
      }
    }

    // Process Fence items
    if (itemLower.includes('construction fence') ||
      (itemLower.includes('proposed fence') && itemLower.includes('height=')) ||
      itemLower.includes('proposed guiderail')) {
      const height = parseHeightFromName(digitizerItem)
      const fenceType = getCivilFenceType(digitizerItem)
      const item = {
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'FT',
        parsed: { heightValue: height }
      }
      if (fenceType !== 'other') {
        items['Fence'][fenceType].push(item)
      }
      if (tracker) tracker.markUsed(rowIndex)
      return
    }

    // Process Soil Erosion items
    if (itemLower.includes('stabilized construction entrance')) {
      const thickness = parseThicknessFromName(digitizerItem)
      items['Soil Erosion']['stabilized_entrance'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'SQ FT',
        parsed: { heightValue: thickness ? thickness / 12 : 0.5 }
      })
      if (tracker) tracker.markUsed(rowIndex)
      return
    }
    if (itemLower.includes('silt fence')) {
      const height = parseSiltFenceHeight(digitizerItem)
      items['Soil Erosion']['silt_fence'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'FT',
        parsed: { heightValue: height || 2.5 }
      })
      if (tracker) tracker.markUsed(rowIndex)
      return
    }
    if (itemLower.includes('inlet filter') && !itemLower.includes('protection')) {
      items['Soil Erosion']['inlet_filter'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'EA',
        parsed: {}
      })
      if (tracker) tracker.markUsed(rowIndex)
      return
    }

    // Process Pads items
    if (itemLower.includes('proposed transformer concrete pad')) {
      const qty = parseQtyFromBracket(digitizerItem) || qtyVal
      const thickness = parseThicknessFromName(digitizerItem)
      items['Pads'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'SQ FT',
        parsed: {
          heightValue: thickness ? thickness / 12 : 0.5,
          qty: qty
        }
      })
      const gravelItem = {
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'SQ FT',
        parsed: { qty: qty }
      }
      // Also add to Excavation
      items['Excavation']['transformer_pad'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'SQ FT',
        parsed: { qty: qty }
      })
      // Gravel: 8" thick items (e.g. "Proposed transformer concrete pad 8" thick (2 No.)") go to transformer_pad_8, others to transformer_pad
      if (itemLower.includes('8" thick') || thickness === 8) {
        items['Gravel']['transformer_pad_8'].push(gravelItem)
      } else {
        items['Gravel']['transformer_pad'].push(gravelItem)
      }
      if (tracker) tracker.markUsed(rowIndex)
      return
    }

    // Process Asphalt items - exclude BPP items (they're handled in BPP section)
    if (itemLower.includes('full depth asphalt pavement') && itemLower.includes('surface course')) {
      // Skip BPP items - they contain street names like "West Street" or "Maple Avenue"
      if (itemLower.includes('west street') || itemLower.includes('maple avenue') ||
        itemLower.includes('- bpp') || itemLower.includes('bpp ')) {
        return
      }

      const thicknesses = parseAsphaltThicknesses(digitizerItem)
      const totalHeight = thicknesses ? (thicknesses.surface + thicknesses.base) / 12 : 4.5 / 12
      const heightFormula = thicknesses ? `(${thicknesses.surface}+${thicknesses.base})/12` : '4.5/12'

      items['Asphalt'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'SQ FT',
        parsed: {
          heightValue: totalHeight,
          heightFormula: heightFormula
        }
      })
      // Also add to Gravel (non-BPP full depth asphalt pavement)
      items['Gravel']['asphalt'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'SQ FT',
        parsed: {}
      })
      if (tracker) tracker.markUsed(rowIndex)
      return
    }

    // Process Concrete Pavement items (Proposed Reinforced concrete sidewalk - also goes to Excavation and Gravel)
    if (itemLower.includes('proposed reinforced concrete sidewalk')) {
      const thickness = parseThicknessFromName(digitizerItem)

      items['Concrete Pavement'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'SQ FT',
        parsed: {
          heightValue: thickness ? thickness / 12 : 0.5
        }
      })
      // Also add to Excavation (height is manual input, leave empty)
      items['Excavation']['reinforced_sidewalk'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'SQ FT',
        parsed: {}
      })
      // Also add to Gravel
      items['Gravel']['reinforced_sidewalk'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'SQ FT',
        parsed: {}
      })
      if (tracker) tracker.markUsed(rowIndex)
      return
    }

    // Process Concrete filled steel pipe bollard items
    if (itemLower.includes('bollard')) {
      const bollardDims = parseBollardDimensions(digitizerItem)

      // Categorize: Footing vs Simple
      // If it mentions "footing", treat as footing item (Group 1)
      if (itemLower.includes('footing') || (bollardDims && bollardDims.footingDia)) {
        // Add to Excavation subsection (for excavation calcs) - keep existing logic
        items['Excavation']['bollard'].push({
          particulars: digitizerItem,
          takeoff,
          unit: unit || 'EA',
          parsed: {
            bollardDimensions: bollardDims
          }
        })

        // Add to Bollard subsection - Footing Group
        items['Concrete filled steel pipe bollard']['footing'].push({
          particulars: digitizerItem,
          takeoff,
          unit: unit || 'EA',
          parsed: {
            bollardDimensions: bollardDims,
            qty: qtyVal
          }
        })
      } else {
        // Simple items (Group 2)
        // Parse height from name for simple items if needed (though requirement says Col M=C, straightforward)
        const simpleHeight = parseHeightFromName(digitizerItem)

        // Add to Bollard subsection - Simple Group
        items['Concrete filled steel pipe bollard']['simple'].push({
          particulars: digitizerItem,
          takeoff,
          unit: unit || 'EA',
          parsed: {
            heightValue: simpleHeight
          }
        })
      }
      if (tracker) tracker.markUsed(rowIndex)
      return
    }

    // Process Site items
    // Hydrant
    if (itemLower.includes('fire hydrant')) {
      items['Site']['Hydrant'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'EA',
        parsed: {}
      })
      if (tracker) tracker.markUsed(rowIndex)
      return
    }
    // Wheel stop
    if (itemLower.includes('wheel stop')) {
      items['Site']['Wheel stop'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'EA',
        parsed: {}
      })
      if (tracker) tracker.markUsed(rowIndex)
      return
    }
    // Drain
    if (itemLower.includes('area drain')) {
      items['Site']['Drain']['Area'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'EA',
        parsed: {}
      })
      if (tracker) tracker.markUsed(rowIndex)
      return
    }
    if (itemLower.includes('floor drain')) {
      items['Site']['Drain']['Floor'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'EA',
        parsed: {}
      })
      if (tracker) tracker.markUsed(rowIndex)
      return
    }

    // Protection
    if (itemLower.includes('inlet filter protection')) {
      items['Site']['Protection'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'EA',
        parsed: {}
      })
      if (tracker) tracker.markUsed(rowIndex)
      return
    }
    // Signages
    if (itemLower.includes('signages')) {
      items['Site']['Signages'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'EA',
        parsed: {}
      })
      if (tracker) tracker.markUsed(rowIndex)
      return
    }
    // Main line
    if (itemLower.includes('connection to existing') && (itemLower.includes('gas') || itemLower.includes('sanitary') || itemLower.includes('water'))) {
      const item = {
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'EA',
        parsed: {}
      }

      if (itemLower.includes('gas')) {
        items['Site']['Main line']['Gas'].push(item)
      } else if (itemLower.includes('sanitary')) {
        items['Site']['Main line']['Sanitary'].push(item)
      } else if (itemLower.includes('water')) {
        items['Site']['Main line']['Water'].push(item)
      }
      if (tracker) tracker.markUsed(rowIndex)
      return
    }

    // Drains & Utilities
    const drainsUtilitiesKeywords = [
      'storm sewer piping',
      'fire service lateral',
      'water service lateral',
      'gas service lateral',
      'sanitary sewer service',
      'underground water main',
      'electrical conduit',
      'underslab drainage',
      'connection to existing utility pole',
      'sanitary invert',
      'backwater valve'
    ]
    if (drainsUtilitiesKeywords.some(k => itemLower.includes(k))) {
      items['Drains & Utilities'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'EA', // Unit will be handled in sheet gen (FT->Col I, EA->Col M)
        parsed: {}
      })
      if (tracker) tracker.markUsed(rowIndex)
      return
    }

    // Alternate
    if (itemLower.includes('(add/alt)')) {
      items['Alternate'].push({
        particulars: digitizerItem,
        takeoff,
        unit: unit || 'EA',
        parsed: {}
      })
      if (tracker) tracker.markUsed(rowIndex)
      return
    }
  })

  return items
}

export default {
  isCivilSiteworkItem,
  getDemoSubSubsection,
  getFenceType,
  getPipeType,
  getSignType,
  getInletType,
  generateCivilDemoFormulas,
  processCivilDemoItems,
  getCivilSubsection,
  parseQtyFromBracket,
  parseHeightFromName,
  parseThicknessFromName,
  parseAsphaltThicknesses,
  parseBollardDimensions,
  parseSiltFenceHeight,
  getCivilFenceType,
  processCivilOtherItems
}

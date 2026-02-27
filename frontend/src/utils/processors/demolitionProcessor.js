import { parseDemolitionItem, normalizeUnit } from '../parsers/dimensionParser'

/**
 * Identifies if a digitizer item belongs to Demolition section
 */
export const isDemolitionItem = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return false
  const itemLower = digitizerItem.toLowerCase()
  return (
    itemLower.startsWith('demo ') ||
    itemLower.startsWith('remove ') ||
    itemLower.startsWith('demo. ') ||
    itemLower.startsWith('demolition ')
  )
}

/**
 * Helper: check if item (lowercased) contains a demolition-prefixed keyword
 */
const hasDemoPrefix = (itemLower, keyword) => {
  return (
    itemLower.includes(`demo ${keyword}`) ||
    itemLower.includes(`demo. ${keyword}`) ||
    itemLower.includes(`demolition ${keyword}`) ||
    itemLower.includes(`demolition. ${keyword}`) ||
    itemLower.includes(`remove ${keyword}`)
  )
}

/**
 * Determines which demolition subsection an item belongs to.
 * Order matters — more specific checks first.
 */
export const getDemolitionSubsection = (digitizerItem) => {
  if (!digitizerItem || typeof digitizerItem !== 'string') return null
  const itemLower = digitizerItem.toLowerCase()

  // ── Existing subsections ────────────────────────────────────────────────

  // Demo slab on grade — SOG, slab on grade, gravel, patio slab on grade, pressure slab, geotextile filter fabric
  if (
    hasDemoPrefix(itemLower, 'sog') ||
    hasDemoPrefix(itemLower, 'slab on grade') ||
    hasDemoPrefix(itemLower, 'gravel') ||
    hasDemoPrefix(itemLower, 'patio slab on grade') ||
    hasDemoPrefix(itemLower, 'pressure slab on grade') ||
    hasDemoPrefix(itemLower, 'pressure slab') ||
    hasDemoPrefix(itemLower, 'slab on grade step') ||
    hasDemoPrefix(itemLower, 'geotextile filter fabric') ||
    hasDemoPrefix(itemLower, 'patio sog') ||
    hasDemoPrefix(itemLower, 'pressure sog')
  ) {
    return 'Demo slab on grade'
  }

  // Demo Ramp on grade
  if (hasDemoPrefix(itemLower, 'rog') || hasDemoPrefix(itemLower, 'ramp on grade')) {
    return 'Demo Ramp on grade'
  }

  // Demo stair on grade (check before strip footing to avoid 'st' confusion)
  if (
    hasDemoPrefix(itemLower, 'stairs on grade') ||
    hasDemoPrefix(itemLower, 'stair landings on grade') ||
    hasDemoPrefix(itemLower, 'stair landing on grade')
  ) {
    return 'Demo stair on grade'
  }

  // Demo strip footing — SF, ST-, WF-, strip footing, wall footing (mirrors isStripFooting)
  if (
    hasDemoPrefix(itemLower, 'sf') ||
    hasDemoPrefix(itemLower, 'st-') ||
    hasDemoPrefix(itemLower, 'wf-') ||
    hasDemoPrefix(itemLower, 'strip footing') ||
    hasDemoPrefix(itemLower, 'wall footing') ||
    hasDemoPrefix(itemLower, 'wf')
  ) {
    return 'Demo strip footing'
  }

  // Demo foundation wall — FW, foundation wall, fndt wall
  if (
    hasDemoPrefix(itemLower, 'fw') ||
    hasDemoPrefix(itemLower, 'foundation wall') ||
    hasDemoPrefix(itemLower, 'fndt wall')
  ) {
    return 'Demo foundation wall'
  }

  // Demo retaining wall
  if (
    hasDemoPrefix(itemLower, 'rw') ||
    hasDemoPrefix(itemLower, 'retaining wall')
  ) {
    return 'Demo retaining wall'
  }

  // Demo isolated footing — F-, isolated footing, footing (but not foundation) (mirrors isIsolatedFooting)
  if (
    hasDemoPrefix(itemLower, 'f-') ||
    hasDemoPrefix(itemLower, 'isolated footing') ||
    (hasDemoPrefix(itemLower, 'footing') && !itemLower.includes('foundation'))
  ) {
    return 'Demo isolated footing'
  }

  // ── New subsections ─────────────────────────────────────────────────────

  // Demo pile caps — PC-, pile cap (mirrors isPileCap)
  if (
    hasDemoPrefix(itemLower, 'pc-') ||
    hasDemoPrefix(itemLower, 'pc') ||
    hasDemoPrefix(itemLower, 'pile cap')
  ) {
    return 'Demo pile caps'
  }

  // Demo service elevator pit — must come before elevator pit (mirrors isServiceElevatorPit)
  if (
    hasDemoPrefix(itemLower, 'sump pit @ service elevator') ||
    hasDemoPrefix(itemLower, 'service elev. pit') ||
    hasDemoPrefix(itemLower, 'service elevator pit') ||
    hasDemoPrefix(itemLower, 'service elev') ||
    hasDemoPrefix(itemLower, 'service elevator') ||
    /(?:demo|demolition|remove)\s+service\s+(elev\.?|elevator)\s+(slab|mat|wall|slope|haunch|sump)/i.test(itemLower)
  ) {
    return 'Demo service elevator pit'
  }

  // Demo elevator pit (mirrors isElevatorPit)
  if (
    hasDemoPrefix(itemLower, 'elev. pit') ||
    hasDemoPrefix(itemLower, 'elevator pit') ||
    hasDemoPrefix(itemLower, 'elev pit') ||
    hasDemoPrefix(itemLower, 'sump pit @ elevator') ||
    hasDemoPrefix(itemLower, 'sump pit') ||
    hasDemoPrefix(itemLower, 'elev slab') ||
    hasDemoPrefix(itemLower, 'elev. slab') ||
    hasDemoPrefix(itemLower, 'elevator slab') ||
    hasDemoPrefix(itemLower, 'elev mat') ||
    hasDemoPrefix(itemLower, 'elev. mat') ||
    hasDemoPrefix(itemLower, 'elev wall') ||
    hasDemoPrefix(itemLower, 'elev. wall') ||
    hasDemoPrefix(itemLower, 'elevator wall') ||
    /(?:demo|demolition|remove)\s+(elev\.?|elevator)\s+(slab|mat|wall|slope|haunch|sump)/i.test(itemLower)
  ) {
    return 'Demo elevator pit'
  }

  // Demo duplex sewage ejector pit — before generic sewage ejector
  if (
    hasDemoPrefix(itemLower, 'duplex sewage ejector pit') ||
    hasDemoPrefix(itemLower, 'duplex sewage ejector')
  ) {
    return 'Demo duplex sewage ejector pit'
  }

  // Demo deep sewage ejector pit — before generic sewage ejector
  if (
    hasDemoPrefix(itemLower, 'deep sewage ejector pit') ||
    hasDemoPrefix(itemLower, 'deep sewage ejector')
  ) {
    return 'Demo deep sewage ejector pit'
  }

  // Demo sewage ejector pit (generic)
  if (
    hasDemoPrefix(itemLower, 'sewage ejector pit') ||
    hasDemoPrefix(itemLower, 'sewage ejector')
  ) {
    return 'Demo sewage ejector pit'
  }

  // Demo sump pump pit
  if (
    hasDemoPrefix(itemLower, 'sump pump pit') ||
    hasDemoPrefix(itemLower, 'sump pump')
  ) {
    return 'Demo sump pump pit'
  }

  // Demo grease trap pit
  if (
    hasDemoPrefix(itemLower, 'grease trap pit') ||
    hasDemoPrefix(itemLower, 'grease trap')
  ) {
    return 'Demo grease trap pit'
  }

  // Demo house trap pit
  if (
    hasDemoPrefix(itemLower, 'house trap pit') ||
    hasDemoPrefix(itemLower, 'house trap')
  ) {
    return 'Demo house trap pit'
  }

  // Demo detention tank
  if (
    hasDemoPrefix(itemLower, 'detention tank')
  ) {
    return 'Demo detention tank'
  }

  // Demo grade beam — GB, GB-, grade beam (mirrors isGradeBeam)
  if (
    hasDemoPrefix(itemLower, 'gb-') ||
    hasDemoPrefix(itemLower, 'gb') ||
    hasDemoPrefix(itemLower, 'gb(') ||
    hasDemoPrefix(itemLower, 'grade beam')
  ) {
    return 'Demo grade beam'
  }

  // Demo tie beam — TB, tie beam (mirrors isTieBeam)
  if (
    hasDemoPrefix(itemLower, 'tb') ||
    hasDemoPrefix(itemLower, 'tb-') ||
    hasDemoPrefix(itemLower, 'tb(') ||
    hasDemoPrefix(itemLower, 'tie beam')
  ) {
    return 'Demo tie beam'
  }

  // Demo strap beam — ST (space/bracket), strap beam — but NOT st- (that's strip footing)
  // Mirrors isStrapBeam: starts with 'st ' or /^st\s*\(/ or includes 'strap beam'
  if (
    hasDemoPrefix(itemLower, 'strap beam') ||
    (hasDemoPrefix(itemLower, 'st ') && !itemLower.includes('st-')) ||
    /(?:demo|demolition|remove)\s+st\s*\(/i.test(itemLower)
  ) {
    return 'Demo strap beam'
  }

  // Demo thickened slab
  if (hasDemoPrefix(itemLower, 'thickened slab')) {
    return 'Demo thickened slab'
  }

  // Demo buttress
  if (hasDemoPrefix(itemLower, 'buttress')) {
    return 'Demo buttress'
  }

  // Demo pier — must start with pier or concrete pier after demo prefix (mirrors isPier: startsWith)
  if (
    /(?:demo|demolition|remove)\.?\s+pier\b/i.test(itemLower) ||
    /(?:demo|demolition|remove)\.?\s+concrete pier\b/i.test(itemLower)
  ) {
    return 'Demo pier'
  }

  // Demo corbel
  if (hasDemoPrefix(itemLower, 'corbel')) {
    return 'Demo corbel'
  }

  // Demo liner wall (linear wall / liner wall / concrete liner wall)
  if (
    hasDemoPrefix(itemLower, 'liner wall') ||
    hasDemoPrefix(itemLower, 'linear wall') ||
    hasDemoPrefix(itemLower, 'concrete liner wall')
  ) {
    return 'Demo liner wall'
  }

  // Demo barrier wall
  if (
    hasDemoPrefix(itemLower, 'barrier wall') ||
    hasDemoPrefix(itemLower, 'vehicle barrier wall') ||
    hasDemoPrefix(itemLower, 'vehicle barrier')
  ) {
    return 'Demo barrier wall'
  }

  // Demo stem wall
  if (hasDemoPrefix(itemLower, 'stem wall')) {
    return 'Demo stem wall'
  }

  // Demo mud slab — exact mud slab / mud mat (must come before mat slab, mirrors isMudSlabFoundation)
  if (
    hasDemoPrefix(itemLower, 'mud slab') ||
    hasDemoPrefix(itemLower, 'mud mat')
  ) {
    return 'Demo mud slab'
  }

  // Demo mat slab — mat + haunch OR mat + number (mirrors isMatSlab, excludes pit types)
  if (
    !itemLower.includes('elevator pit') && !itemLower.includes('elev. pit') &&
    !itemLower.includes('service elevator') &&
    !itemLower.includes('duplex sewage ejector') && !itemLower.includes('deep sewage ejector') &&
    !itemLower.includes('sump pump pit') && !itemLower.includes('grease trap') &&
    !itemLower.includes('house trap') && !itemLower.includes('mud') &&
    (
      hasDemoPrefix(itemLower, 'mat slab') ||
      (hasDemoPrefix(itemLower, 'mat') && hasDemoPrefix(itemLower, 'haunch')) ||
      (hasDemoPrefix(itemLower, 'mat') && /demo\.?\s+mat[-\s]*\d+/i.test(itemLower))
    )
  ) {
    return 'Demo mat slab'
  }

  // Demo pilaster
  if (hasDemoPrefix(itemLower, 'pilaster')) {
    return 'Demo pilaster'
  }

  return null
}

/**
 * All known demolition subsection names (used for initialising the result object).
 */
const ALL_DEMO_SUBSECTIONS = [
  'Demo slab on grade',
  'Demo Ramp on grade',
  'Demo strip footing',
  'Demo foundation wall',
  'Demo retaining wall',
  'Demo isolated footing',
  'Demo stair on grade',
  'Demo pile caps',
  'Demo pilaster',
  'Demo grade beam',
  'Demo tie beam',
  'Demo strap beam',
  'Demo thickened slab',
  'Demo buttress',
  'Demo pier',
  'Demo corbel',
  'Demo liner wall',
  'Demo barrier wall',
  'Demo stem wall',
  'Demo elevator pit',
  'Demo service elevator pit',
  'Demo detention tank',
  'Demo duplex sewage ejector pit',
  'Demo deep sewage ejector pit',
  'Demo sewage ejector pit',
  'Demo sump pump pit',
  'Demo grease trap pit',
  'Demo house trap pit',
  'Demo mat slab',
  'Demo mud slab'
]

/**
 * Returns formula column assignments for a given demo subsection / item type.
 *
 * Formula conventions (mirrors foundation patterns):
 *   • Wall / beam items  → FT (I) = takeoff; SQ FT (J) = takeoff * width; CY (L) = J*H/27
 *   • Area / slab items  → SQ FT (J) = takeoff; CY (L) = J*H/27
 *   • Volumetric / pit items with L×W → SQ FT (J) = F*G*C; CY (L) = J*H/27; QTY (M) = C
 *   • FT-only items     → FT (I) = takeoff only (e.g. pile caps counted in EA)
 */
export const generateDemolitionFormulas = (type, rowNum, parsedData) => {
  const formulas = {
    ft: null,       // Column I
    sqFt: null,     // Column J
    lbs: null,      // Column K
    cy: null,       // Column L
    qtyFinal: null  // Column M
  }

  switch (type) {
    // ── existing ────────────────────────────────────────────────────────────

    case 'Demo slab on grade':
    case 'Demo Ramp on grade':
    case 'demo_extra_sqft':
    case 'demo_extra_rog_sqft':
      formulas.sqFt = `C${rowNum}`
      formulas.cy = `J${rowNum}*H${rowNum}/27`
      break

    case 'Demo strip footing':
      // Same as Foundation: I=C; ST → J=H*I, L=J*G/27; SF/WF → J=G*I, L=J*H/27
      formulas.ft = `C${rowNum}`
      if (parsedData?.itemType === 'ST') {
        formulas.sqFt = `H${rowNum}*I${rowNum}`
        formulas.cy = `J${rowNum}*G${rowNum}/27`
      } else {
        formulas.sqFt = `G${rowNum}*I${rowNum}`
        formulas.cy = `J${rowNum}*H${rowNum}/27`
      }
      break

    case 'Demo foundation wall':
    case 'Demo retaining wall':
    case 'demo_extra_ft':
    case 'demo_extra_rw':
      // Same as Foundation linear wall: I=C, J=I*H, L=J*G/27
      formulas.ft = `C${rowNum}`
      formulas.sqFt = `H${rowNum}*I${rowNum}`
      formulas.cy = `J${rowNum}*G${rowNum}/27`
      break

    case 'Demo isolated footing':
    case 'demo_extra_ea':
      formulas.sqFt = `F${rowNum}*G${rowNum}*C${rowNum}`
      formulas.cy = `J${rowNum}*H${rowNum}/27`
      formulas.qtyFinal = `C${rowNum}`
      break

    // ── wall / linear items — same as Foundation: I=C, J=H*I, L=J*G/27 ─────────────────────────
    case 'Demo grade beam':
    case 'Demo tie beam':
    case 'Demo strap beam':
    case 'Demo liner wall':
    case 'Demo barrier wall':
    case 'Demo stem wall':
    case 'Demo corbel':
      formulas.ft = `C${rowNum}`
      formulas.sqFt = `H${rowNum}*I${rowNum}`
      formulas.cy = `J${rowNum}*G${rowNum}/27`
      break

    // ── Thickened slab — same as Foundation: I=C, J=H*I, L=J*G/27 ─────────────────────────
    case 'Demo thickened slab':
      formulas.ft = `C${rowNum}`
      formulas.sqFt = `H${rowNum}*I${rowNum}`
      formulas.cy = `J${rowNum}*G${rowNum}/27`
      break

    // ── slab / area items — same as Foundation mat/mud: J=C, L=J*H/27 ─────────────────────────
    case 'Demo mat slab':
    case 'Demo mud slab':
      formulas.sqFt = `C${rowNum}`
      formulas.cy = `J${rowNum}*H${rowNum}/27`
      break

    // ── pile_cap / pilaster: J=C*H*G, L=J*F/27, M=C (same as Foundation) ────────────────────────
    case 'Demo pile caps':
    case 'Demo pilaster':
      formulas.sqFt = `C${rowNum}*H${rowNum}*G${rowNum}`
      formulas.cy = `J${rowNum}*F${rowNum}/27`
      formulas.qtyFinal = `C${rowNum}`
      break

    // ── pier: J=C*F*G, L=J*H/27, M=C (same as Foundation) ────────────────────────
    case 'Demo pier':
      formulas.sqFt = `C${rowNum}*F${rowNum}*G${rowNum}`
      formulas.cy = `J${rowNum}*H${rowNum}/27`
      formulas.qtyFinal = `C${rowNum}`
      break

    // ── buttress: M=C only (same as Foundation buttress_takeoff) ────────────────────────
    case 'Demo buttress':
      formulas.qtyFinal = `C${rowNum}`
      break

    // ── Pit items: same formulas as Foundation (slab/mat/mat_slab → J=C, L=J*H/27; wall/slope → I=C, J=I*H, L=J*G/27; sump_pit/lid_slab as per Foundation) ────────────────────────────
    case 'Demo elevator pit':
    case 'Demo service elevator pit': {
      const elevSub = parsedData?.itemSubType
      if (elevSub === 'sump_pit') {
        formulas.sqFt = `16*C${rowNum}`
        formulas.cy = `C${rowNum}*1.3`
        formulas.qtyFinal = `C${rowNum}`
      } else if (elevSub === 'slab' || elevSub === 'mat' || elevSub === 'mat_slab') {
        formulas.sqFt = `C${rowNum}`
        if (parsedData?.height !== undefined && parsedData?.height !== '') formulas.height = parsedData.height
        formulas.cy = `J${rowNum}*H${rowNum}/27`
      } else if (elevSub === 'wall' || elevSub === 'slope_transition') {
        formulas.ft = `C${rowNum}`
        formulas.sqFt = `I${rowNum}*H${rowNum}`
        if (parsedData?.width !== undefined && parsedData?.width !== '') formulas.width = parsedData.width
        if (parsedData?.height !== undefined && parsedData?.height !== '') formulas.height = parsedData.height
        formulas.cy = `J${rowNum}*G${rowNum}/27`
      } else {
        formulas.sqFt = `C${rowNum}`
        if (parsedData?.height !== undefined && parsedData?.height !== '') formulas.height = parsedData.height
        formulas.cy = `J${rowNum}*H${rowNum}/27`
      }
      break
    }
    case 'Demo detention tank': {
      const dtSub = parsedData?.itemSubType
      if (dtSub === 'slab' || dtSub === 'lid_slab') {
        formulas.sqFt = `C${rowNum}`
        if (parsedData?.height !== undefined && parsedData?.height !== '') formulas.height = parsedData.height
        formulas.cy = `J${rowNum}*H${rowNum}/27`
      } else if (dtSub === 'wall') {
        formulas.ft = `C${rowNum}`
        formulas.sqFt = `I${rowNum}*H${rowNum}`
        if (parsedData?.width !== undefined && parsedData?.width !== '') formulas.width = parsedData.width
        if (parsedData?.height !== undefined && parsedData?.height !== '') formulas.height = parsedData.height
        formulas.cy = `J${rowNum}*G${rowNum}/27`
      } else {
        formulas.sqFt = `C${rowNum}`
        if (parsedData?.height !== undefined && parsedData?.height !== '') formulas.height = parsedData.height
        formulas.cy = `J${rowNum}*H${rowNum}/27`
      }
      break
    }
    case 'Demo duplex sewage ejector pit':
    case 'Demo deep sewage ejector pit':
    case 'Demo sewage ejector pit':
    case 'Demo sump pump pit':
    case 'Demo grease trap pit':
    case 'Demo house trap pit': {
      const pitSub = parsedData?.itemSubType
      if (pitSub === 'slab' || pitSub === 'mat' || pitSub === 'mat_slab') {
        formulas.sqFt = `C${rowNum}`
        if (parsedData?.height !== undefined && parsedData?.height !== '') formulas.height = parsedData.height
        formulas.cy = `J${rowNum}*H${rowNum}/27`
      } else if (pitSub === 'wall' || pitSub === 'slope_transition') {
        formulas.ft = `C${rowNum}`
        formulas.sqFt = `I${rowNum}*H${rowNum}`
        if (parsedData?.width !== undefined && parsedData?.width !== '') formulas.width = parsedData.width
        if (parsedData?.height !== undefined && parsedData?.height !== '') formulas.height = parsedData.height
        formulas.cy = `J${rowNum}*G${rowNum}/27`
      } else {
        formulas.sqFt = `C${rowNum}`
        if (parsedData?.height !== undefined && parsedData?.height !== '') formulas.height = parsedData.height
        formulas.cy = `J${rowNum}*H${rowNum}/27`
      }
      break
    }

    default:
      break
  }

  return formulas
}

/**
 * Processes all demolition items from raw data, grouped by subsection.
 */
export const processDemolitionItems = (rawDataRows, headers, tracker = null) => {
  // Initialise result with every known subsection
  const demolitionItemsBySubsection = {}
  ALL_DEMO_SUBSECTIONS.forEach(name => { demolitionItemsBySubsection[name] = [] })

  // Find column indices
  const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
  const totalIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'total')
  const unitIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'units')

  if (digitizerIdx === -1 || totalIdx === -1 || unitIdx === -1) {
    return demolitionItemsBySubsection
  }

  rawDataRows.forEach((row, rowIndex) => {
    const digitizerItem = row[digitizerIdx]
    const total = parseFloat(row[totalIdx]) || 0
    const unit = normalizeUnit(row[unitIdx] || '')

    if (!isDemolitionItem(digitizerItem)) return

    const subsection = getDemolitionSubsection(digitizerItem)
    if (!subsection || demolitionItemsBySubsection[subsection] === undefined) return

    const parsed = parseDemolitionItem(digitizerItem, total, unit, subsection)

    // Grouping key
    let groupKey = 'DEFAULT'
    const itemLower = (digitizerItem || '').toLowerCase()

    if (subsection === 'Demo stair on grade') {
      const atMatch = digitizerItem.match(/@\s*(.+)$/i)
      groupKey = atMatch ? atMatch[1].trim() : 'NO_AT'
      parsed.itemSubType = itemLower.includes('landings') ? 'landings' : 'stairs'
    } else if (
      (subsection === 'Demo slab on grade' || subsection === 'Demo Ramp on grade') &&
      itemLower.includes('"')
    ) {
      const thickMatch = digitizerItem.match(/(\d+)["']?\s*(?:thick|thk)/i)
      if (thickMatch) groupKey = `THICK_${thickMatch[1]}`
    } else if (digitizerItem.includes('(')) {
      // Group by first dimension bracket value for all subsections that use brackets
      const bracketMatch = digitizerItem.match(/\(([^x)]+)/)
      if (bracketMatch) groupKey = `DIM_${bracketMatch[1].trim()}`
    }
    // For pit wall/slope, use dimension-based groupKey (same as Foundation) so grouping matches
    const pitSubsections = ['Demo elevator pit', 'Demo service elevator pit', 'Demo detention tank', 'Demo duplex sewage ejector pit', 'Demo deep sewage ejector pit', 'Demo sewage ejector pit', 'Demo sump pump pit', 'Demo grease trap pit', 'Demo house trap pit']
    if (pitSubsections.includes(subsection) && parsed.groupKey) {
      groupKey = parsed.groupKey
    }

    demolitionItemsBySubsection[subsection].push({
      ...parsed,
      subsection,
      groupKey,
      rawRow: row,
      rawRowNumber: rowIndex + 2
    })

    if (tracker) tracker.markUsed(rowIndex)
  })

  // Group / merge within each subsection
  Object.keys(demolitionItemsBySubsection).forEach(subsection => {
    const items = demolitionItemsBySubsection[subsection]
    if (items.length === 0) return

    // Demo stair on grade: build groups { heading, stairs, landings }
    if (subsection === 'Demo stair on grade') {
      const groupMap = new Map()
      items.forEach(item => {
        const key = item.groupKey || 'NO_AT'
        const heading = key !== 'NO_AT' ? key : null
        if (!groupMap.has(key)) groupMap.set(key, { heading, stairs: null, landings: null })
        const g = groupMap.get(key)
        if (item.itemSubType === 'stairs') g.stairs = item
        else if (item.itemSubType === 'landings') g.landings = item
      })
      demolitionItemsBySubsection[subsection] = Array.from(groupMap.values()).filter(g => g.stairs || g.landings)
      return
    }

    const groupMap = new Map()
    items.forEach(item => {
      const key = item.groupKey || 'DEFAULT'
      if (!groupMap.has(key)) groupMap.set(key, [])
      groupMap.get(key).push(item)
    })

    const groups = Array.from(groupMap.values())
    const singleItemGroups = groups.filter(g => g.length === 1)
    const multiItemGroups = groups.filter(g => g.length > 1)

    if (singleItemGroups.length > 1) {
      const mergedItems = []
      singleItemGroups.forEach(g => mergedItems.push(...g))
      mergedItems.forEach(item => (item.groupKey = 'MERGED'))
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
import { convertToFeet, normalizeUnit } from '../parsers/dimensionParser'

/**
 * Processes Superstructure section items from raw data.
 * Subsections: CIP Slabs, Balcony slab, Terrace slab, Patch slab, Slab steps, LW concrete fill.
 */

/**
 * Parses bracket dimensions like (2'-0"x1'-3") into first and second values in feet.
 * @param {string} particulars - Digitizer item text
 * @returns {{ first: number, second: number }|null}
 */
const parseBracketDimensions = (particulars) => {
  const match = particulars.match(/\(([^)]+)\)/)
  if (!match) return null
  const parts = match[1].split(/x/i).map(p => p.trim())
  if (parts.length < 2) return null
  const first = convertToFeet(parts[0])
  const second = convertToFeet(parts[1])
  return { first, second }
}

/**
 * Parses bracket dimensions from the bracket that contains "x" (e.g. for "2B-1 (Upturned) (10"x1'-6")" finds "(10"x1'-6")").
 * @param {string} particulars - Digitizer item text
 * @returns {{ first: number, second: number }|null}
 */
const parseBracketDimensionsContainingX = (particulars) => {
  const match = particulars.match(/\(([^)]*x[^)]+)\)/)
  if (!match) return null
  const parts = match[1].split(/x/i).map(p => p.trim())
  if (parts.length < 2) return null
  const first = convertToFeet(parts[0])
  const second = convertToFeet(parts[1])
  return { first, second }
}

/**
 * Parses trailing group id like (1) or (2) at end of item name.
 * @param {string} particulars - Digitizer item text
 * @returns {number|null}
 */
const parseTrailingGroupId = (particulars) => {
  const match = particulars.trim().match(/\(\s*(\d+)\s*\)\s*$/)
  return match ? parseInt(match[1], 10) : null
}

/**
 * Parses three dimensions from bracket like (12"x12"x10'-1") into length, width, height in feet.
 * @param {string} particulars - Digitizer item text
 * @returns {{ length: number, width: number, height: number }|null}
 */
const parseThreeDimensions = (particulars) => {
  const match = particulars.match(/\(([^)]+)\)/)
  if (!match) return null
  const parts = match[1].split(/x/i).map(p => p.trim())
  if (parts.length < 3) return null
  return {
    length: convertToFeet(parts[0]),
    width: convertToFeet(parts[1]),
    height: convertToFeet(parts[2])
  }
}

/**
 * Parses QTY from Thermal break item name. "3rd FL to 7th FL - Thermal break" -> 5 (7-3+1).
 * "10th FL to 13th FL - Thermal break" -> 4. Single "14th FL - Thermal break" -> null (col E empty).
 * @param {string} particulars - Digitizer item text
 * @returns {number|null}
 */
const parseThermalBreakQtyFromName = (particulars) => {
  if (!particulars || typeof particulars !== 'string') return null
  const rangeMatch = particulars.match(/(\d+)(?:st|nd|rd|th)\s*FL\s*to\s*(\d+)(?:st|nd|rd|th)\s*FL/i)
  if (rangeMatch) {
    const first = parseInt(rangeMatch[1], 10)
    const second = parseInt(rangeMatch[2], 10)
    if (!isNaN(first) && !isNaN(second) && second >= first) {
      return second - first + 1
    }
  }
  return null
}

/**
 * Parses Height/Ht/H= value from text like "LW concrete fill, H=1'-1"" or "Height=1'-1"" or "Ht.=4""
 * @param {string} particulars - Digitizer item text
 * @returns {number|null}
 */
const parseHeightFromName = (particulars) => {
  const match = particulars.match(/(?:Height|Ht\.?|H)\s*=\s*([^,\s]+)/i)
  if (!match) return null
  return convertToFeet(match[1].trim())
}

/**
 * Parses QTY from bracket like "(4 No.)" or "(1 No.)" in concrete pad items.
 * @param {string} particulars - Digitizer item text
 * @returns {number|null}
 */
const parseQtyFromBracket = (particulars) => {
  const match = particulars.match(/\(\s*(\d+)\s*(?:No\.?|EA)?\s*\)/i)
  return match ? parseInt(match[1], 10) : null
}

/**
 * Parses a dimension string that may include fractions (e.g. "4 ½", "2", "4.5") to inches as number.
 * @param {string} str - String like "4 ½", "2", "4.5"
 * @returns {number}
 */
const parseInchesFromText = (str) => {
  if (!str || typeof str !== 'string') return 0
  let s = str.trim()
  s = s.replace(/\s*½/g, '.5').replace(/\s*¼/g, '.25').replace(/\s*¾/g, '.75')
  const wholeFrac = s.match(/(\d+)\s+(\d+)\/(\d+)/)
  if (wholeFrac) {
    const whole = parseFloat(wholeFrac[1]) || 0
    const num = parseFloat(wholeFrac[2]) || 0
    const den = parseFloat(wholeFrac[3]) || 1
    return whole + (num / den)
  }
  const num = parseFloat(s.replace(/\s/g, '').replace(/[^\d.-]/g, ''))
  if (!isNaN(num)) return num
  return 0
}

/**
 * Extracts first and second inch values from SOMD item name like "SOMD S3 4 ½" LW concrete topping over 2" MD x18GA".
 * Returns { firstValueInches, secondValueInches } in inches (as numbers). Heights in formulas use first/12 and second/12.
 * @param {string} particulars - Digitizer item text
 * @returns {{ firstValueInches: number, secondValueInches: number }|null}
 */
const parseSOMDDimensions = (particulars) => {
  if (!particulars || typeof particulars !== 'string') return null
  const parts = particulars.split('"')
  if (parts.length < 3) return null
  const beforeFirst = parts[0].trim()
  const beforeSecond = (parts[1] || '').trim()
  const firstMatch = beforeFirst.match(/(\d+(?:\s*[½¼¾]|\s*\d+\/\d+)?)\s*$/)
  const secondMatch = beforeSecond.match(/(\d+(?:\s*[½¼¾]|\s*\d+\/\d+)?)\s*$/)
  if (!firstMatch || !secondMatch) return null
  const firstValueInches = parseInchesFromText(firstMatch[1].trim())
  const secondValueInches = parseInchesFromText(secondMatch[1].trim())
  if (firstValueInches <= 0 || secondValueInches <= 0) return null
  return { firstValueInches, secondValueInches }
}

/**
 * Returns unit string for a superstructure groupKey.
 * @param {string} groupKey - Item group key
 * @param {string} unit - Raw unit from row
 * @returns {string}
 */
const getUnitForSuperstructureGroup = (groupKey, unit) => {
  const unitMap = {
    slabStep: 'FT',
    lwConcreteFill: 'SQ FT',
    toppingSlab: 'SQ FT',
    thermalBreak: 'FT',
    raisedKneeWall: 'FT',
    raisedSlab: 'SQ FT',
    builtUpKneeWall: 'FT',
    builtUpSlab: 'SQ FT',
    builtupRampsKneeWall: 'FT',
    builtupRamp: 'SQ FT',
    builtUpStairs: 'Treads',
    concreteHanger: 'EA',
    shearWalls: 'FT',
    parapetWalls: 'FT',
    columnsTakeoff: 'EA',
    concretePost: 'EA',
    concreteEncasement: 'EA',
    dropPanelBracket: 'EA',
    dropPanelH: 'SQ FT',
    beams: 'FT',
    curbs: 'FT',
    nonShrinkGrout: 'EA'
  }
  // For concrete pad, preserve the raw unit from data (can be SQ FT or EA)
  if (groupKey === 'concretePad' || groupKey === 'concretePadNoBracket') {
    return unit || 'SQ FT'
  }
  if (groupKey === 'repairScope') return (unit || 'FT')
  return unitMap[groupKey] ?? (unit || 'SQ FT')
}


/**
 * Returns which subsection and group an item belongs to from particulars.
 * @param {string} particulars - Digitizer item text
 * @returns {{ subsection: string, groupKey: string, heightFormula: string|null, heightValue: number|null, widthValue: number|null, qty: number|null }|null}
 */
export const getSuperstructureItemType = (particulars) => {
  if (!particulars || typeof particulars !== 'string') return null
  const p = particulars.trim().toLowerCase()
  if (p.includes('detention tank lid slab')) return null
  if (p.includes('duplex sewage ejector pit slab')) return null
  if (p.includes('slab step')) {
    const dims = parseBracketDimensions(particulars)
    if (dims) {
      return { subsection: 'Slab steps', groupKey: 'slabStep', heightFormula: null, heightValue: dims.second, widthValue: dims.first, qty: 2 }
    }
    const wideMatch = particulars.match(/(?:slab step)\s*([0-9'"\-]+)\s*(?:wide|width)\s*,?\s*(?:Height|Ht|H)=([0-9'"\-]+)/i)
    if (wideMatch) {
      const widthValue = convertToFeet(wideMatch[1].trim())
      const heightValue = convertToFeet(wideMatch[2].trim().replace(/["\s]+$/, '').trim())
      return { subsection: 'Slab steps', groupKey: 'slabStep', heightFormula: null, heightValue, widthValue, qty: 2 }
    }
  }
  if (p.includes('lw concrete fill') || p.includes('light weight concrete fill')) {
    let heightValue = parseHeightFromName(particulars)
    if (heightValue == null) {
      const thickMatch = particulars.match(/lw concrete fill\s*([0-9'"\-]+)"?\s*(?:thick|thk)?\.?/i)
      if (thickMatch) {
        const raw = thickMatch[1].trim()
        heightValue = raw.includes("'") || raw.includes('"') ? convertToFeet(raw) : convertToFeet(`${raw}"`)
      }
    }
    return { subsection: 'LW concrete fill', groupKey: 'lwConcreteFill', heightFormula: null, heightValue: heightValue ?? 1 + 1 / 12, widthValue: null, qty: null }
  }
  // Landing SOMD for Stairs - Infilled tads (exclude from Slab on metal deck)
  if (p.includes('landing') && (p.includes('somd') || p.includes('slab on metal deck')) && p.includes('lw concrete topping')) {
    const dims = parseSOMDDimensions(particulars)
    if (dims) {
      const groupKey = `somd_${dims.firstValueInches}_${dims.secondValueInches}`
      return { subsection: 'Stairs – Infilled tads', groupKey: 'infilledLanding', heightFormula: null, heightValue: 0.67, widthValue: null, qty: null, firstValueInches: dims.firstValueInches, secondValueInches: dims.secondValueInches, somdGroupKey: groupKey }
    }
  }
  if (p.includes('slab on metal deck') || p.includes('somd') || p.includes('corrugated slab on metal deck') || p.includes('corrugated somd')) {
    const dims = parseSOMDDimensions(particulars)
    if (dims) {
      const groupKey = `somd_${dims.firstValueInches}_${dims.secondValueInches}`
      return { subsection: 'Slab on metal deck', groupKey, heightFormula: null, heightValue: null, widthValue: null, qty: null, firstValueInches: dims.firstValueInches, secondValueInches: dims.secondValueInches }
    }
  }
  if (p.includes('cast in place slab 8"') || p.includes('cast in place slab 8\"')) {
    return { subsection: 'CIP Slabs', groupKey: 'castInPlaceSlab8', heightFormula: '8/12', heightValue: null, widthValue: null, qty: null }
  }
  if (p.includes('cip slab 8"') || p.includes('cip slab 8\"')) {
    return { subsection: 'CIP Slabs', groupKey: 'cipSlabVar', heightFormula: '8/12', heightValue: null, widthValue: null, qty: null }
  }
  if (p.includes('roof slab 8"') || p.includes('roof slab 8\"')) {
    return { subsection: 'CIP Slabs', groupKey: 'roofSlab8', heightFormula: '8/12', heightValue: null, widthValue: null, qty: null }
  }
  if (p.includes('balcony slab 8"') || p.includes('balcony slab 8\"')) {
    return { subsection: 'Balcony slab', groupKey: 'balcony', heightFormula: '8/12', heightValue: null, widthValue: null, qty: null }
  }
  if (p.includes('terrace slab 8"') || p.includes('terrace slab 8\"')) {
    return { subsection: 'Terrace slab', groupKey: 'terrace', heightFormula: '8/12', heightValue: null, widthValue: null, qty: null }
  }
  // Exclude Sump pump pit slab 8" from CIP Slabs
  if (p.includes('sump pump pit slab 8')) {
    return null
  }
  if (/^slab\s+[0-9'"\-]/.test(p) && !p.includes('slab step') && !p.includes('patch slab') && !p.includes('balcony slab') && !p.includes('terrace slab') && !p.includes('topping slab') && !p.includes('overpour slab') && !p.includes('slab on ')) {
    const slabThickMatch = particulars.match(/^slab\s+([0-9'"\-]+)"?\s*(?:thick|thk)?\.?/i)
    if (slabThickMatch) {
      const raw = slabThickMatch[1].trim()
      const heightValue = raw.includes("'") || raw.includes('"') ? convertToFeet(raw) : convertToFeet(`${raw}"`)
      const groupKey = Math.abs(heightValue - 8 / 12) < 0.001 ? 'slab8' : 'slabVar'
      return { subsection: 'CIP Slabs', groupKey, heightFormula: null, heightValue, widthValue: null, qty: null }
    }
  }
  if (p.includes('slab 8"') || p.includes('slab 8\"')) {
    return { subsection: 'CIP Slabs', groupKey: 'slab8', heightFormula: '8/12', heightValue: null, widthValue: null, qty: null }
  }
  if (p.includes('patch slab')) {
    return { subsection: 'Patch slab', groupKey: 'patch', heightFormula: null, heightValue: 0.5, widthValue: null, qty: null }
  }
  if ((p.includes('topping slab') || p.includes('overpour slab')) && (p.includes('thick') || p.includes('thk') || p.includes('"'))) {
    const thickMatch = particulars.match(/(\d+(?:\.\d+)?)\s*"\s*(?:thick|thk)/i) || particulars.match(/(\d+(?:\.\d+)?)\s*"/)
    const inches = thickMatch ? parseFloat(thickMatch[1]) : 2
    return { subsection: 'Topping slab', groupKey: 'toppingSlab', heightFormula: null, heightValue: inches / 12, widthValue: null, qty: null }
  }
  if (p.includes('thermal break')) {
    const qtyFromName = parseThermalBreakQtyFromName(particulars)
    return { subsection: 'Thermal break', groupKey: 'thermalBreak', heightFormula: null, heightValue: null, widthValue: null, qty: qtyFromName ?? null }
  }
  if (p.includes('knee wall') && (p.includes('builtup') || p.includes('built up') || p.includes('@ builtup'))) {
    const dims = parseBracketDimensions(particulars)
    if (dims) {
      return { subsection: 'Built-up slab', groupKey: 'builtUpKneeWall', heightFormula: null, heightValue: dims.second, widthValue: dims.first, qty: null }
    }
  }
  if (p.includes('concrete hanger')) {
    const dims = parseThreeDimensions(particulars)
    if (dims) {
      return { subsection: 'Concrete hanger', groupKey: 'concreteHanger', heightFormula: null, heightValue: dims.height, widthValue: dims.width, lengthValue: dims.length, qty: null }
    }
    const wideMatch = particulars.match(/concrete hanger\s*([0-9'"\-]+)\s*x\s*([0-9'"\-]+)\s*(?:wide|width)\s*,?\s*(?:Height|Ht|H)=([0-9'"\-]+)/i)
    if (wideMatch) {
      const lengthValue = convertToFeet(wideMatch[1].trim())
      const widthValue = convertToFeet(wideMatch[2].trim())
      const heightValue = convertToFeet(wideMatch[3].trim().replace(/["\s]+$/, '').trim())
      return { subsection: 'Concrete hanger', groupKey: 'concreteHanger', heightFormula: null, heightValue, widthValue, lengthValue, qty: null }
    }
  }
  if (p.includes('built up stairs') || (p.includes('built up stair') && !p.includes('@'))) {
    return { subsection: 'Built-up stair', groupKey: 'builtUpStairs', heightFormula: '7/12', heightValue: null, widthFormula: '11/12', widthValue: null, lengthValue: 3, qty: null }
  }
  if (p.includes('builtup ramp') || p.includes('built up ramp')) {
    const thickMatch = particulars.match(/(?:builtup ramp|built up ramp)\s*([0-9'"\-]+)"?\s*(?:thick|thk)?/i)
    let heightValue = 3 / 12
    if (thickMatch) {
      const raw = thickMatch[1].trim()
      heightValue = raw.includes("'") || raw.includes('"') ? convertToFeet(raw) : convertToFeet(`${raw}"`)
    } else {
      const fallback = particulars.match(/(\d+(?:\.\d+)?)\s*"\s*(?:thick|thk)/i) || particulars.match(/(\d+(?:\.\d+)?)\s*"/)
      if (fallback) heightValue = parseFloat(fallback[1]) / 12
    }
    const groupId = parseTrailingGroupId(particulars)
    return { subsection: 'Builtup ramps', groupKey: 'builtupRamp', heightFormula: null, heightValue, widthValue: null, qty: null, groupId: groupId ?? 1 }
  }
  if (p.includes('knee wall') && p.includes('x') && /\)\s*\(\d+\)\s*$/.test(particulars.trim())) {
    const dims = parseBracketDimensions(particulars)
    if (dims) {
      const groupId = parseTrailingGroupId(particulars)
      return { subsection: 'Builtup ramps', groupKey: 'builtupRampsKneeWall', heightFormula: null, heightValue: dims.second, widthValue: dims.first, qty: null, groupId: groupId ?? 1 }
    }
  }
  if (p.includes('knee wall') && p.includes('x')) {
    const dims = parseBracketDimensions(particulars)
    if (dims) {
      return { subsection: 'Raised slab', groupKey: 'raisedKneeWall', heightFormula: null, heightValue: dims.second, widthValue: dims.first, qty: null }
    }
  }
  if (p.includes('raised slab')) {
    const thickMatch = particulars.match(/raised slab\s*([0-9'"\-]+)"?\s*(?:thick|thk)?/i)
    let heightValue = 4 / 12
    if (thickMatch) {
      const raw = thickMatch[1].trim()
      heightValue = raw.includes("'") || raw.includes('"') ? convertToFeet(raw) : convertToFeet(`${raw}"`)
    } else {
      const fallback = particulars.match(/(\d+(?:\.\d+)?)\s*"/) || particulars.match(/(\d+(?:\.\d+)?)\s*(?:thick|thk)/i)
      if (fallback) heightValue = parseFloat(fallback[1]) / 12
    }
    return { subsection: 'Raised slab', groupKey: 'raisedSlab', heightFormula: null, heightValue, widthValue: null, qty: null }
  }
  if (p.includes('builtup slab') || p.includes('built up slab')) {
    const thickMatch = particulars.match(/(?:builtup slab|built up slab)\s*([0-9'"\-]+)"?\s*(?:thick|thk)?/i)
    let heightValue = 3 / 12
    if (thickMatch) {
      const raw = thickMatch[1].trim()
      heightValue = raw.includes("'") || raw.includes('"') ? convertToFeet(raw) : convertToFeet(`${raw}"`)
    } else {
      const fallback = particulars.match(/(\d+(?:\.\d+)?)\s*"\s*(?:thick|thk)/i) || particulars.match(/(\d+(?:\.\d+)?)\s*"/) || particulars.match(/(\d+(?:\.\d+)?)\s*(?:thick|thk)/i)
      if (fallback) heightValue = parseFloat(fallback[1]) / 12
    }
    return { subsection: 'Built-up slab', groupKey: 'builtUpSlab', heightFormula: null, heightValue, widthValue: null, qty: null }
  }
  if (p.includes('sw ') || p.includes('shear wall') || p.includes('concrete wall')) {
    const dims = parseBracketDimensions(particulars)
    if (dims) {
      return { subsection: 'Shear Walls', groupKey: 'shearWalls', heightFormula: null, heightValue: dims.second, widthValue: dims.first, qty: null }
    }
    const swWideMatch = particulars.match(/(?:shear wall|sw|concrete wall)\s*([0-9'"\-]+)\s*(?:wide|width)\s*,?\s*(?:Height|Ht|H)=([0-9'"\-]+)/i)
    if (swWideMatch) {
      const widthValue = convertToFeet(swWideMatch[1].trim())
      const heightValue = convertToFeet(swWideMatch[2].trim().replace(/["\s]+$/, '').trim())
      return { subsection: 'Shear Walls', groupKey: 'shearWalls', heightFormula: null, heightValue, widthValue, qty: null }
    }
  }
  if (p.includes('parapet wall')) {
    const dims = parseBracketDimensions(particulars)
    if (dims) {
      return { subsection: 'Parapet walls', groupKey: 'parapetWalls', heightFormula: null, heightValue: dims.second, widthValue: dims.first, qty: null }
    }
    const parapetWideMatch = particulars.match(/parapet wall\s*([0-9'"\-]+)\s*(?:wide|width)\s*,?\s*(?:Height|Ht|H)=([0-9'"\-]+)/i)
    if (parapetWideMatch) {
      const widthValue = convertToFeet(parapetWideMatch[1].trim())
      const heightValue = convertToFeet(parapetWideMatch[2].trim().replace(/["\s]+$/, '').trim())
      return { subsection: 'Parapet walls', groupKey: 'parapetWalls', heightFormula: null, heightValue, widthValue, qty: null }
    }
  }
  if (p.includes('as per takeoff count') || (p.includes('column') && p.includes('count'))) {
    return { subsection: 'Columns', groupKey: 'columnsTakeoff', heightFormula: null, heightValue: null, widthValue: null, qty: null }
  }
  if (p.includes('concrete post')) {
    const dims = parseThreeDimensions(particulars)
    if (dims) {
      return { subsection: 'Concrete post', groupKey: 'concretePost', heightFormula: null, heightValue: dims.height, widthValue: dims.width, lengthValue: dims.length, qty: null }
    }
    const postWideMatch = particulars.match(/concrete post\s*([0-9'"\-]+)\s*x\s*([0-9'"\-]+)\s*(?:wide|width)\s*,?\s*(?:Height|Ht|H)=([0-9'"\-]+)/i)
    if (postWideMatch) {
      const lengthValue = convertToFeet(postWideMatch[1].trim())
      const widthValue = convertToFeet(postWideMatch[2].trim())
      const heightValue = convertToFeet(postWideMatch[3].trim().replace(/["\s]+$/, '').trim())
      return { subsection: 'Concrete post', groupKey: 'concretePost', heightFormula: null, heightValue, widthValue, lengthValue, qty: null }
    }
  }
  if (p.includes('concrete encasement')) {
    const dims = parseThreeDimensions(particulars)
    if (dims) {
      return { subsection: 'Concrete encasement', groupKey: 'concreteEncasement', heightFormula: null, heightValue: dims.height, widthValue: dims.width, lengthValue: dims.length, qty: null }
    }
    const encWideMatch = particulars.match(/concrete encasement\s*([0-9'"\-]+)\s*x\s*([0-9'"\-]+)\s*(?:wide|width)\s*,?\s*(?:Height|Ht|H)=([0-9'"\-]+)/i)
    if (encWideMatch) {
      const lengthValue = convertToFeet(encWideMatch[1].trim())
      const widthValue = convertToFeet(encWideMatch[2].trim())
      const heightValue = convertToFeet(encWideMatch[3].trim().replace(/["\s]+$/, '').trim())
      return { subsection: 'Concrete encasement', groupKey: 'concreteEncasement', heightFormula: null, heightValue, widthValue, lengthValue, qty: null }
    }
  }
  if (p.includes('drop panel')) {
    const dims = parseThreeDimensions(particulars)
    if (dims) {
      return { subsection: 'Drop panel', groupKey: 'dropPanelBracket', heightFormula: null, heightValue: dims.height, widthValue: dims.width, lengthValue: dims.length, qty: null }
    }
    const dropWideMatch = particulars.match(/drop panel\s*([0-9'"\-]+)\s*x\s*([0-9'"\-]+)\s*(?:wide|width)\s*,?\s*(?:Height|Ht|H)=([0-9'"\-]+)/i)
    if (dropWideMatch) {
      const lengthValue = convertToFeet(dropWideMatch[1].trim())
      const widthValue = convertToFeet(dropWideMatch[2].trim())
      const heightValue = convertToFeet(dropWideMatch[3].trim().replace(/["\s]+$/, '').trim())
      return { subsection: 'Drop panel', groupKey: 'dropPanelBracket', heightFormula: null, heightValue, widthValue, lengthValue, qty: null }
    }
    const dropHBracketMatch = particulars.match(/drop panel\s*\(\s*(?:Height|Ht|H)\s*=\s*([^)]+)\)/i)
    if (dropHBracketMatch) {
      const heightValue = convertToFeet(dropHBracketMatch[1].trim())
      return { subsection: 'Drop panel', groupKey: 'dropPanelH', heightFormula: null, heightValue, widthValue: null, qty: null }
    }
    if (particulars.match(/(?:Height|Ht|H)\s*=/i)) {
      const heightValue = parseHeightFromName(particulars)
      return { subsection: 'Drop panel', groupKey: 'dropPanelH', heightFormula: null, heightValue: heightValue ?? 0.67, widthValue: null, qty: null }
    }
  }
  // Beams: item must start with digits+B- (e.g. 1B-1, 2B-1) or RB- (e.g. RB-1, RB-2) or BHB- (e.g. BHB-1) only
  if (
    !p.includes('secant pile') &&
    !p.includes('core beam') &&
    !p.includes('pile w/') &&
    /^(?:'?\s*)?(\d+B-\d+|RB-\d+|BHB-\d+)/i.test(particulars.trim()) &&
    particulars.match(/\([^)]+x[^)]+\)/) &&
    !particulars.match(/\([^)]+x[^)]+x[^)]+\)/)
  ) {
    const dims = parseBracketDimensionsContainingX(particulars)
    if (dims) {
      return { subsection: 'Beams', groupKey: 'beams', heightFormula: null, heightValue: dims.second, widthValue: dims.first, qty: null }
    }
  }
  if (p.includes('concrete curb') || p.includes('curb')) {
    const dims = parseBracketDimensions(particulars)
    if (dims) {
      return { subsection: 'Curbs', groupKey: 'curbs', heightFormula: null, heightValue: dims.second, widthValue: dims.first, qty: null }
    }
    const curbWideMatch = particulars.match(/(?:concrete curb|curb)\s*([0-9'"\-]+)\s*(?:wide|width)\s*,?\s*(?:Height|Ht|H)=([0-9'"\-]+)/i)
    if (curbWideMatch) {
      const widthValue = convertToFeet(curbWideMatch[1].trim())
      const heightValue = convertToFeet(curbWideMatch[2].trim().replace(/["\s]+$/, '').trim())
      return { subsection: 'Curbs', groupKey: 'curbs', heightFormula: null, heightValue, widthValue, qty: null }
    }
  }
  if ((p.includes('concrete pad') || p.includes('pad') || p.includes('housekeeping pad')) && !p.includes('transformer')) {
    const padNoMatch = particulars.match(/(?:concrete pad|pad|housekeeping pad)\s*\(\s*(\d+)\s*\)\s*no\.?/i)
    if (padNoMatch) {
      const qty = parseInt(padNoMatch[1], 10)
      return { subsection: 'Concrete pad', groupKey: 'concretePad', heightFormula: null, heightValue: 4 / 12, widthValue: null, qty }
    }
    const thickMatch = particulars.match(/(?:concrete pad|pad|housekeeping pad)\s*([0-9'"\-]+)"?\s*(?:thick|thk)?/i) || particulars.match(/(\d+)\s*"/) || particulars.match(/(\d+)"\s*$/)
    if (thickMatch) {
      let heightValue = 4 / 12
      const raw = thickMatch[1].trim()
      if (raw.includes("'") || raw.includes('"')) {
        heightValue = convertToFeet(raw)
      } else if (/^\d+(\.\d+)?$/.test(raw)) {
        heightValue = parseFloat(raw) / 12
      }
      const qtyFromBracket = parseQtyFromBracket(particulars)
      if (qtyFromBracket != null) {
        return { subsection: 'Concrete pad', groupKey: 'concretePad', heightFormula: null, heightValue, widthValue: null, qty: qtyFromBracket }
      }
      return { subsection: 'Concrete pad', groupKey: 'concretePadNoBracket', heightFormula: null, heightValue, widthValue: null, qty: null, noBracket: true }
    }
  }
  if (p.includes('non-shrink grout') || p.includes('non shrink grout')) {
    return { subsection: 'Non-shrink grout', groupKey: 'nonShrinkGrout', heightFormula: null, heightValue: null, widthValue: null, qty: null }
  }
  if (p.includes('concrete wall crack repair')) {
    return { subsection: 'Repair scope', groupKey: 'repairScope', heightFormula: null, heightValue: null, widthValue: null, qty: null, itemSubType: 'wall' }
  }
  if (p.includes('slab crack repair')) {
    return { subsection: 'Repair scope', groupKey: 'repairScope', heightFormula: null, heightValue: null, widthValue: null, qty: null, itemSubType: 'slab' }
  }
  if (p.includes('column crack repair')) {
    return { subsection: 'Repair scope', groupKey: 'repairScope', heightFormula: null, heightValue: null, widthValue: null, qty: null, itemSubType: 'column' }
  }
  return null
}

/**
 * Processes Superstructure items from raw data.
 * @param {Array} rawDataRows - Rows from raw Excel (excluding header)
 * @param {Array} headers - Column headers
 * @param {UsedRowTracker} tracker - Optional tracker to mark used row indices
 * @returns {{ cipSlab8: Array, ... toppingSlab: Array, thermalBreak: Array, raisedSlab: Object, builtUpSlab: Object }}
 */
export const processSuperstructureItems = (rawDataRows, headers, tracker = null) => {
  const cipSlab8 = []
  const cipRoofSlab8 = []
  const cipCastInPlaceSlab8 = []
  const cipCIPSlabVar = []
  const balconySlab = []
  const terraceSlab = []
  const patchSlab = []
  const slabSteps = []
  const lwConcreteFill = []
  const slabOnMetalDeckByKey = {}
  const infilledLandingItems = []
  const toppingSlab = []
  const thermalBreak = []
  const raisedSlab = { kneeWall: [], raisedSlab: [] }
  const builtUpSlab = { kneeWall: [], builtUpSlab: [] }
  const builtUpStair = { kneeWall: [], builtUpStairs: [] }
  const builtupRamps = { kneeWall: [], ramp: [] }
  const concreteHanger = []
  const shearWalls = []
  const parapetWalls = []
  const columnsTakeoff = []
  const concretePost = []
  const concreteEncasement = []
  const dropPanelBracket = []
  const dropPanelH = []
  const beams = []
  const curbs = []
  const concretePad = []
  const nonShrinkGrout = []
  const repairScope = []

  const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
  const totalIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'total')
  const unitIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'units')
  const estimateIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'estimate')
  const qtyIdx = headers.findIndex(h => {
    const t = h && String(h).toLowerCase().trim()
    return t === 'qty' || t === 'quantity'
  })

  if (digitizerIdx === -1 || totalIdx === -1) {
    return { cipSlab8, cipRoofSlab8, cipCastInPlaceSlab8, cipCIPSlabVar, balconySlab, terraceSlab, patchSlab, slabSteps, lwConcreteFill, slabOnMetalDeck: [], infilledLandingItems: [], toppingSlab, thermalBreak, raisedSlab, builtUpSlab, builtUpStair, builtupRamps, concreteHanger, shearWalls, parapetWalls, columnsTakeoff, cipStairsGroups: [], concretePost, concreteEncasement, dropPanelBracket, dropPanelH, beams, curbs, concretePad, nonShrinkGrout, repairScope }
  }

  rawDataRows.forEach((row, rowIndex) => {
    const particulars = row[digitizerIdx]
    const total = row[totalIdx]
    const takeoff = total !== '' && total !== null && total !== undefined ? parseFloat(total) : ''
    const unit = unitIdx >= 0 ? normalizeUnit(row[unitIdx] || '') : 'SQ FT'

    if (estimateIdx >= 0 && row[estimateIdx]) {
      const estimateVal = String(row[estimateIdx]).trim()
      if (estimateVal !== 'Superstructure') return
    }

    const itemType = getSuperstructureItemType(particulars)
    if (!itemType) return

    // Mark this row as used
    if (tracker) {
      tracker.markUsed(rowIndex)
    }

    if (itemType.groupKey === 'infilledLanding') {
      infilledLandingItems.push({
        particulars: particulars || '',
        takeoff,
        unit: 'SQ FT',
        parsed: {
          somdGroupKey: itemType.somdGroupKey,
          firstValueInches: itemType.firstValueInches,
          secondValueInches: itemType.secondValueInches
        }
      })
      return
    }
    if (itemType.groupKey && itemType.groupKey.startsWith('somd_')) {
      const key = itemType.groupKey
      if (!slabOnMetalDeckByKey[key]) {
        slabOnMetalDeckByKey[key] = {
          groupKey: key,
          firstValueInches: itemType.firstValueInches,
          secondValueInches: itemType.secondValueInches,
          particulars: particulars || '',
          items: []
        }
      }
      slabOnMetalDeckByKey[key].items.push({
        particulars: particulars || '',
        takeoff,
        unit: 'SQ FT'
      })
      return
    }

    const qtyVal = qtyIdx >= 0 && row[qtyIdx] !== '' && row[qtyIdx] != null ? parseFloat(row[qtyIdx]) : undefined
    const qtyForItem = itemType.groupKey === 'thermalBreak' ? (itemType.qty ?? undefined) : (itemType.groupKey === 'dropPanelH' || itemType.groupKey === 'concretePad' ? (itemType.qty ?? qtyVal ?? undefined) : (itemType.qty ?? qtyVal ?? undefined))
    const unitForItem = getUnitForSuperstructureGroup(itemType.groupKey, unit)
    const item = {
      particulars: particulars || '',
      takeoff,
      unit: unitForItem,
      parsed: {
        heightFormula: itemType.heightFormula,
        heightValue: itemType.heightValue,
        widthValue: itemType.widthValue ?? undefined,
        widthFormula: itemType.widthFormula ?? undefined,
        lengthValue: itemType.lengthValue ?? undefined,
        qty: qtyForItem,
        groupId: itemType.groupId ?? undefined,
        noBracket: itemType.noBracket ?? undefined,
        itemSubType: itemType.itemSubType ?? undefined
      }
    }

    switch (itemType.groupKey) {
      case 'slab8':
        cipSlab8.push(item)
        break
      case 'castInPlaceSlab8':
        cipCastInPlaceSlab8.push(item)
        break
      case 'cipSlabVar':
        cipCIPSlabVar.push(item)
        break
      case 'roofSlab8':
        cipRoofSlab8.push(item)
        break
      case 'balcony':
        balconySlab.push(item)
        break
      case 'terrace':
        terraceSlab.push(item)
        break
      case 'patch':
        patchSlab.push(item)
        break
      case 'slabStep':
        slabSteps.push(item)
        break
      case 'lwConcreteFill':
        lwConcreteFill.push(item)
        break
      case 'toppingSlab':
        toppingSlab.push(item)
        break
      case 'thermalBreak':
        thermalBreak.push(item)
        break
      case 'raisedKneeWall':
        raisedSlab.kneeWall.push(item)
        break
      case 'raisedSlab':
        raisedSlab.raisedSlab.push(item)
        break
      case 'builtUpKneeWall':
        if (builtUpSlab.kneeWall.length === 0) {
          builtUpSlab.kneeWall.push(item)
        } else {
          builtUpStair.kneeWall.push(item)
        }
        break
      case 'builtUpSlab':
        builtUpSlab.builtUpSlab.push(item)
        break
      case 'builtupRampsKneeWall':
        builtupRamps.kneeWall.push(item)
        break
      case 'builtupRamp':
        builtupRamps.ramp.push(item)
        break
      case 'builtUpStairs':
        builtUpStair.builtUpStairs.push(item)
        break
      case 'concreteHanger':
        concreteHanger.push(item)
        break
      case 'shearWalls':
        shearWalls.push(item)
        break
      case 'parapetWalls':
        parapetWalls.push(item)
        break
      case 'columnsTakeoff':
        columnsTakeoff.push(item)
        break
      case 'concretePost':
        concretePost.push(item)
        break
      case 'concreteEncasement':
        concreteEncasement.push(item)
        break
      case 'dropPanelBracket':
        dropPanelBracket.push(item)
        break
      case 'dropPanelH':
        dropPanelH.push(item)
        break
      case 'beams':
        beams.push(item)
        break
      case 'curbs':
        curbs.push(item)
        break
      case 'concretePad':
      case 'concretePadNoBracket':
        concretePad.push(item)
        break
      case 'nonShrinkGrout':
        nonShrinkGrout.push(item)
        break
      case 'repairScope':
        repairScope.push(item)
        break
      default:
        break
    }
  })

  const slabOnMetalDeck = Object.values(slabOnMetalDeckByKey)
  return { cipSlab8, cipRoofSlab8, cipCastInPlaceSlab8, cipCIPSlabVar, balconySlab, terraceSlab, patchSlab, slabSteps, lwConcreteFill, slabOnMetalDeck, infilledLandingItems, toppingSlab, thermalBreak, raisedSlab, builtUpSlab, builtUpStair, builtupRamps, concreteHanger, shearWalls, parapetWalls, columnsTakeoff, cipStairsGroups: [], concretePost, concreteEncasement, dropPanelBracket, dropPanelH, beams, curbs, concretePad, nonShrinkGrout, repairScope }
}

export default {
  getSuperstructureItemType,
  processSuperstructureItems
}
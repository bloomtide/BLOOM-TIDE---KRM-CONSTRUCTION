import { ANGLE_WEIGHTS } from '../constants/angleWeights'
/**
 * Parser for SOE (Shoring of Excavation) items
 */

/**
 * Extracts numeric value from dimension string (e.g., "27'-10"" -> 27.833)
 * @param {string} dimStr - Dimension string
 * @returns {number} - Decimal feet value
 */
const parseDimension = (dimStr) => {
    if (!dimStr) return 0
    const match = dimStr.match(/(\d+)(?:'-?)?(\d+)?/)
    if (!match) return 0
    const feet = parseInt(match[1]) || 0
    const inches = parseInt(match[2]) || 0
    return feet + (inches / 12)
}

/**
 * Rounds up to nearest multiple of 5
 * @param {number} value - Value to round
 * @returns {number} - Rounded value
 */
const roundToMultipleOf5 = (value) => {
    return Math.ceil(value / 5) * 5
}

/**
 * Extracts parameters from drilled soldier pile name
 * @param {string} itemName - Item name
 * @returns {object} - Parsed parameters
 */
export const parseSoldierPile = (itemName) => {
    const result = {
        type: null,
        diameter: null,
        thickness: null,
        hpWeight: null,
        height: 0,
        embedment: null,
        rockSocket: null,
        heightRaw: 0,
        calculatedHeight: 0,
        groupKey: null
    }

    // Check if it's HP type
    const hpMatch = itemName.match(/HP(\d+)x(\d+)/i)
    if (hpMatch) {
        result.type = 'hp'
        result.hpWeight = parseFloat(hpMatch[2])

        // Extract H value - move hyphen to end of character class
        const hMatch = itemName.match(/H=([0-9'"\-]+)/)
        if (hMatch) {
            result.heightRaw = parseDimension(hMatch[1])
            result.calculatedHeight = roundToMultipleOf5(result.heightRaw)
        }

        result.groupKey = `HP-${result.heightRaw}`
        return result
    }

    // Check if it's drilled type
    const drilledMatch = itemName.match(/([0-9.]+)Ø\s*x\s*([0-9.]+)/i)
    if (drilledMatch) {
        result.type = 'drilled'
        result.diameter = parseFloat(drilledMatch[1])
        result.thickness = parseFloat(drilledMatch[2])

        // Extract H, E, RS values - move hyphen to end of character class
        const hMatch = itemName.match(/H=([0-9'"\-]+)/)
        const eMatch = itemName.match(/E=([0-9'"\-]+)/)
        const rsMatch = itemName.match(/RS=([0-9'"\-]+)/)

        if (hMatch) result.heightRaw = parseDimension(hMatch[1])
        if (eMatch) result.embedment = parseDimension(eMatch[1])
        if (rsMatch) result.rockSocket = parseDimension(rsMatch[1])

        // Calculate total height based on parameters
        if (result.rockSocket && result.heightRaw) {
            // H + RS (ignore E if present)
            result.calculatedHeight = roundToMultipleOf5(result.heightRaw + result.rockSocket)
        } else if (result.heightRaw) {
            // H only (ignore E if present)
            result.calculatedHeight = roundToMultipleOf5(result.heightRaw)
        }

        // Create group key
        const pattern = result.embedment && result.rockSocket ? 'E+RS' :
            result.embedment ? 'E' :
                result.rockSocket ? 'RS' : 'H'

        const eValue = result.embedment ? Math.round(result.embedment * 12) : 0
        const rsValue = result.rockSocket ? Math.round(result.rockSocket * 12) : 0

        result.groupKey = `${result.diameter}-${result.thickness}-${pattern}-${eValue}-${rsValue}`
    }

    return result
}

/**
 * Calculates weight of pile
 * @param {object} parsed - Parsed soldier pile data
 * @returns {number} - Weight in lbs/ft
 */
export const calculatePileWeight = (parsed) => {
    if (parsed.type === 'hp') {
        return parsed.hpWeight || 0
    } else if (parsed.type === 'drilled') {
        // (Diameter - Thickness) * Thickness * 10.69
        return (parsed.diameter - parsed.thickness) * parsed.thickness * 10.69
    }
    // Timber has no weight
    return 0
}

export const isSoldierPile = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false
    const itemLower = digitizerItem.toLowerCase()

    // Exclude supporting angle items
    if (itemLower.includes('supporting angle')) return false
    // Exclude timber soldier piles (handled in separate subsection)
    if (itemLower.includes('timber soldier pile')) return false

    return (itemLower.includes('drilled soldier pile') || itemLower.includes('soldier pile')) &&
        !itemLower.includes('secant') &&
        !itemLower.includes('tangent')
}

/**
 * Identifies if item is a timber soldier pile
 * e.g. 4"x4" Timber soldier piles (H=11'-2", E=5'-0)
 */
export const isTimberSoldierPile = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false
    const itemLower = digitizerItem.toLowerCase()
    const regex = /\b(timber\s+soldier\s+piles|timber\s+piles?|wood(en)?\s+pile)\b/
    return regex.test(itemLower)
}

/**
 * Identifies if item is a timber plank
 * e.g. 3"x12" Timber planks (H=3'-0")
 */
export const isTimberPlank = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false
    const itemLower = digitizerItem.toLowerCase()
    return /\b(timber|wood|wooden)\s+planks?\b/i.test(digitizerItem)
}

/**
 * Parses timber plank parameters
 * e.g. 3"x12" Timber planks (H=3'-0") or (2) 3"x12" Timber planks (H=3'-0")
 * Same structure as timber soldier piles but: no E (embedment), calculatedHeight = H from bracket (no rounding)
 * Optional qty from (N) at start
 */
export const parseTimberPlank = (itemName) => {
    const result = {
        type: 'timber_plank',
        size: null,
        heightRaw: 0,
        calculatedHeight: 0,
        qty: null,
        groupKey: null
    }

    const qtyMatch = itemName.match(/^\((\d+)\)/)
    if (qtyMatch) {
        result.qty = parseInt(qtyMatch[1], 10) || 1
    }

    const sizeMatch = itemName.match(/(\d+)"\s*x\s*(\d+)"/i)
    if (sizeMatch) {
        result.size = `${sizeMatch[1]}x${sizeMatch[2]}`
    }

    const hMatch = itemName.match(/H=([0-9'"\-]+)/)
    if (hMatch) result.heightRaw = parseDimension(hMatch[1])

    // Timber planks: use H from bracket directly, do NOT round up
    result.calculatedHeight = result.heightRaw

    const hValue = Math.round(result.heightRaw * 12)
    const qtyPart = result.qty != null ? `-q${result.qty}` : ''
    result.groupKey = `timber-plank-${result.size || 'other'}-${hValue}${qtyPart}`

    return result
}

/**
 * Parses timber soldier pile parameters
 * e.g. 4"x4" Timber soldier piles (H=11'-2", E=5'-0)
 * Same structure as drilled: H, E, calculatedHeight, groupKey
 */
export const parseTimberSoldierPile = (itemName) => {
    const result = {
        type: 'timber',
        size: null,
        heightRaw: 0,
        embedment: null,
        calculatedHeight: 0,
        groupKey: null
    }

    const sizeMatch = itemName.match(/(\d+)"\s*x\s*(\d+)"/i)
    if (sizeMatch) {
        result.size = `${sizeMatch[1]}x${sizeMatch[2]}`
    }

    const hMatch = itemName.match(/H=([0-9'"\-]+)/)
    const eMatch = itemName.match(/E=([0-9'"\-]+)/)

    if (hMatch) result.heightRaw = parseDimension(hMatch[1])
    if (eMatch) result.embedment = parseDimension(eMatch[1])

    result.calculatedHeight = result.heightRaw > 0 ? roundToMultipleOf5(result.heightRaw) : 0

    const hValue = Math.round(result.heightRaw * 12)
    const eValue = result.embedment ? Math.round(result.embedment * 12) : 0
    result.groupKey = `timber-${result.size || 'other'}-${hValue}-${eValue}`

    return result
}

/**
 * Identifies if item is a timber post
 * e.g. 6"x6" Timber post (H=5'-6", E=4'-0")
 */
export const isTimberPost = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false
    const itemLower = digitizerItem.toLowerCase()
    return /\b(timber|wood|wooden)\s+post\b/.test(itemLower)
}

/**
 * Parses timber post parameters
 * e.g. 6"x6" Timber post (H=5'-6", E=4'-0")
 * Same structure as timber soldier pile but: calculatedHeight = H from bracket (no rounding)
 */
export const parseTimberPost = (itemName) => {
    const result = {
        type: 'timber_post',
        size: null,
        heightRaw: 0,
        embedment: null,
        calculatedHeight: 0,
        groupKey: null
    }

    const sizeMatch = itemName.match(/(\d+)"\s*x\s*(\d+)"/i)
    if (sizeMatch) {
        result.size = `${sizeMatch[1]}x${sizeMatch[2]}`
    }

    const hMatch = itemName.match(/H=([0-9'"\-]+)/)
    const eMatch = itemName.match(/E=([0-9'"\-]+)/)

    if (hMatch) result.heightRaw = parseDimension(hMatch[1])
    if (eMatch) result.embedment = parseDimension(eMatch[1])

    // Timber post: use H from bracket directly, do NOT round up
    result.calculatedHeight = result.heightRaw

    const hValue = Math.round(result.heightRaw * 12)
    const eValue = result.embedment ? Math.round(result.embedment * 12) : 0
    result.groupKey = `timber-post-${result.size || 'other'}-${hValue}-${eValue}`

    return result
}

/**
 * Identifies if item is a timber sheet (vertical, horizontal, wood, wooden, or plain timber sheet/sheets)
 * e.g. 3"x10" Vertical timber sheets H=6'-0", E=4'-0"
 *      3"x10" Horizontal timber sheets H=6'-0", E=4'-0"
 *      3"x10" Wood timber sheets H=6'-0"
 */
export const isTimberSheets = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false
    const itemLower = digitizerItem.toLowerCase()
    return /\b(?:vertical|horizontal|wood(?:en)?)\s+timber\s+sheets?\b|\btimber\s+sheets?\b/.test(itemLower)
}

// Keep old names as aliases so any lingering imports don't break
export const isVerticalTimberSheets = isTimberSheets
export const isHorizontalTimberSheets = isTimberSheets

/**
 * Parses timber sheet parameters (vertical, horizontal, wood, wooden, or plain)
 * e.g. 3"x10" Vertical timber sheets H=6'-0", E=4'-0" or H=7'-9" (E optional)
 * No grouping — all timber sheets shown in one block with one sum (groupKey unused).
 * Formulas: FT(I)=C, SQ FT(J)=I*H
 */
export const parseTimberSheets = (itemName) => {
    const result = {
        type: 'timber_sheets',
        size: null,
        heightRaw: 0,
        embedment: null,
        calculatedHeight: 0,
        groupKey: 'timber_sheets'
    }

    const sizeMatch = itemName.match(/(\d+)"\s*x\s*(\d+)"/i)
    if (sizeMatch) {
        result.size = `${sizeMatch[1]}x${sizeMatch[2]}`
    }

    const hMatch = itemName.match(/H=([0-9'"\-]+)/)
    const eMatch = itemName.match(/E=([0-9'"\-]+)/)

    if (hMatch) result.heightRaw = parseDimension(hMatch[1])
    if (eMatch) result.embedment = parseDimension(eMatch[1])

    result.calculatedHeight = result.heightRaw

    return result
}

// Keep old parsers as aliases so any lingering imports don't break
export const parseVerticalTimberSheets = parseTimberSheets
export const parseHorizontalTimberSheets = parseTimberSheets

/**
 * Identifies if item is a timber raker (wood raker)
 * e.g. 4"x4" Wood raker (H=2'-1" typ.)
 */
export const isTimberRaker = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false
    const itemLower = digitizerItem.toLowerCase()
    return /\b(timber|wood|wooden)\s+raker\b/.test(itemLower)
}

/**
 * Parses timber raker parameters
 * e.g. 4"x4" Wood raker (H=2'-1" typ.)
 * Same structure as timber planks: no column K, calculatedHeight = H from bracket (no rounding)
 */
export const parseTimberRaker = (itemName) => {
    const result = {
        type: 'timber_raker',
        size: null,
        heightRaw: 0,
        calculatedHeight: 0,
        groupKey: null
    }

    const sizeMatch = itemName.match(/(\d+)"\s*x\s*(\d+)"/i)
    if (sizeMatch) {
        result.size = `${sizeMatch[1]}x${sizeMatch[2]}`
    }

    const hMatch = itemName.match(/H=([0-9'"\-]+)/)
    if (hMatch) result.heightRaw = parseDimension(hMatch[1])

    // Timber raker: use H from bracket directly, do NOT round up
    result.calculatedHeight = result.heightRaw

    const hValue = Math.round(result.heightRaw * 12)
    result.groupKey = `timber-raker-${result.size || 'other'}-${hValue}`

    return result
}

/**
 * Identifies if item is a timber brace (Horizontal wood brace or Timber Corner brace)
 * e.g. 2"x4" Horizontal wood brace (H=2'-0" typ.) or (2) 6"x6" Timber Corner brace
 */
export const isTimberBrace = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false
    const itemLower = digitizerItem.toLowerCase()
    return /\b(timber|wood|wooden)(?:\s+corner)?\s+brace\b/.test(itemLower)
}

/**
 * Parses timber brace parameters
 * Two types: horizontal_wood_brace (H from bracket) or timber_corner_brace (qty from (N) at start)
 */
export const parseTimberBrace = (itemName) => {
    const itemLower = itemName.toLowerCase()
    const isCornerBrace = itemLower.includes('timber corner brace')

    const result = {
        type: isCornerBrace ? 'timber_brace_corner' : 'timber_brace_horizontal',
        size: null,
        heightRaw: 0,
        calculatedHeight: 0,
        qty: 1,
        groupKey: null
    }

    const sizeMatch = itemName.match(/(\d+)"\s*x\s*(\d+)"/i)
    if (sizeMatch) {
        result.size = `${sizeMatch[1]}x${sizeMatch[2]}`
    }

    if (isCornerBrace) {
        // (2) 6"x6" Timber Corner brace - qty from (N) at start
        const qtyMatch = itemName.match(/^\((\d+)\)/)
        if (qtyMatch) {
            result.qty = parseInt(qtyMatch[1], 10) || 1
        }
        result.groupKey = `timber-brace-corner-${result.size || 'other'}-${result.qty}`
    } else {
        // 2"x4" Horizontal wood brace (H=2'-0" typ.)
        const hMatch = itemName.match(/H=([0-9'"\-]+)/)
        if (hMatch) result.heightRaw = parseDimension(hMatch[1])
        result.calculatedHeight = result.heightRaw
        const hValue = Math.round(result.heightRaw * 12)
        result.groupKey = `timber-brace-h-${result.size || 'other'}-${hValue}`
    }

    return result
}

/**
 * Identifies if item is a timber waler
 * e.g. (2) 6"x6" Timber waler
 */
export const isTimberWaler = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false
    const itemLower = digitizerItem.toLowerCase()
    return /\b(timber|wood|wooden)\s+waler\b/.test(itemLower)
}

/**
 * Parses timber waler parameters
 * e.g. (2) 6"x6" Timber waler - qty from (N) at start, or 1 if not present
 */
export const parseTimberWaler = (itemName) => {
    const result = {
        type: 'timber_waler',
        size: null,
        qty: 1,
        groupKey: null
    }

    const sizeMatch = itemName.match(/(\d+)"\s*x\s*(\d+)"/i)
    if (sizeMatch) {
        result.size = `${sizeMatch[1]}x${sizeMatch[2]}`
    }

    const qtyMatch = itemName.match(/^\((\d+)\)/)
    if (qtyMatch) {
        result.qty = parseInt(qtyMatch[1], 10) || 1
    }

    result.groupKey = `timber-waler-${result.size || 'other'}-${result.qty}`

    return result
}

/**
 * Identifies if item is a timber stringer
 * e.g. (2) 6"x6" Timber stringer
 */
export const isTimberStringer = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false
    const itemLower = digitizerItem.toLowerCase()
    return /\b(timber|wood|wooden)\s+stringer\b/.test(itemLower)
}

/**
 * Parses timber stringer parameters
 * e.g. (2) 6"x6" Timber stringer - qty from (N) at start, or 1 if not present
 */
export const parseTimberStringer = (itemName) => {
    const result = {
        type: 'timber_stringer',
        size: null,
        qty: 1,
        groupKey: null
    }

    const sizeMatch = itemName.match(/(\d+)"\s*x\s*(\d+)"/i)
    if (sizeMatch) {
        result.size = `${sizeMatch[1]}x${sizeMatch[2]}`
    }

    const qtyMatch = itemName.match(/^\((\d+)\)/)
    if (qtyMatch) {
        result.qty = parseInt(qtyMatch[1], 10) || 1
    }

    result.groupKey = `timber-stringer-${result.size || 'other'}`

    return result
}

/**
 * Identifies if item is a primary secant pile
 */
export const isPrimarySecantPile = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false
    const itemLower = digitizerItem.toLowerCase()
    const primaryRegex = /\bprimary\s+secant(?:\s+pile)?\b/
    return primaryRegex.test(itemLower)
}

/**
 * Identifies if item is a secondary secant pile
 */
export const isSecondarySecantPile = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false
    const itemLower = digitizerItem.toLowerCase()
    const secondaryRegex = /\bsecondary\s+secant(?:\s+pile)?\b/
    return secondaryRegex.test(itemLower)
}

/**
 * Identifies if item is a tangent pile
 */
export const isTangentPile = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false
    const itemLower = digitizerItem.toLowerCase()
    return itemLower.includes('tangent pile')
}

/**
 * Identifies if item is a sheet pile
 */
export const isSheetPile = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false
    const itemLower = digitizerItem.toLowerCase()
    return itemLower.includes('sheet pile')
}

/**
 * Identifies if item is timber lagging
 */
export const isTimberLagging = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false
    const itemLower = digitizerItem.toLowerCase()

    // Exclude supporting angle items
    if (itemLower.includes('supporting angle')) return false

    return /\b(timber|wood|wooden)\s+lagging\b/.test(itemLower)
}

/**
 * New identifiers for additional SOE sections
 */
export const isTimberSheeting = (item) => {
    if (!item || typeof item !== 'string') return false
    const lower = item.toLowerCase()
    return /\b(timber|wood|wooden)\s+sheeting\b/.test(lower)
}
export const isWaler = (item) => item?.toLowerCase().includes('waler') && !item?.toLowerCase().includes('timber waler')
export const isRaker = (item) => item?.toLowerCase().includes('raker') && !item?.toLowerCase().includes('upper') && !item?.toLowerCase().includes('lower') && !item?.toLowerCase().includes('wood raker')
export const isUpperRaker = (item) => item?.toLowerCase().includes('upper raker')
export const isLowerRaker = (item) => item?.toLowerCase().includes('lower raker')
export const isStandOff = (item) => item?.toLowerCase().includes('stand off')
export const isKicker = (item) => item?.toLowerCase().includes('kicker')
export const isChannel = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('channel') && !itemLower.includes('bollard')
}
export const isRollChock = (item) => item?.toLowerCase().includes('roll chock')
export const isStudBeam = (item) => item?.toLowerCase().includes('stud beam')
export const isCornerbrace = (item) => item?.toLowerCase().includes('corner brace')
export const isKneeBrace = (item) => item?.toLowerCase().includes('knee brace')
export const isSupportingAngle = (item) => item?.toLowerCase().includes('supporting angle')
export const isParging = (item) => item?.toLowerCase().startsWith('parging') || item?.toLowerCase().startsWith('parge')
export const isHeelBlock = (item) => item?.toLowerCase().includes('heel block')
export const isUnderpinning = (item) => item?.toLowerCase().includes('underpinning') || item?.toLowerCase().includes('underpin')
export const isRockAnchor = (item) => item?.toLowerCase().includes('rock anchor')
export const isRockBolt = (item) => item?.toLowerCase().includes('rock bolt')
export const isAnchor = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('anchor') && !itemLower.includes('rock anchor') && !itemLower.includes('tie back') && !itemLower.includes('hollow down anchor')
}
export const isTieBack = (item) => item?.toLowerCase().includes('tie back') || item?.toLowerCase().includes('hollow down anchor')
export const isConcreteSoilRetentionPier = (item) => item?.toLowerCase().includes('concrete soil retention pier')
export const isGuideWall = (item) => item?.toLowerCase().includes('guide wall') || item?.toLowerCase().includes('guilde wall')
export const isDowelBar = (item) => item?.toLowerCase().includes('dowel bar') || item?.toLowerCase().includes('steel dowels bar')
export const isRockPin = (item) => item?.toLowerCase().includes('rock pin') || item?.toLowerCase().includes('rock pins')
export const isShotcrete = (item) => item?.toLowerCase().includes('shotcrete')
export const isPermissionGrouting = (item) => item?.toLowerCase().includes('permission grouting')
export const isButton = (item) => item?.toLowerCase().includes('concrete button') || item?.toLowerCase().includes('buttons')
export const isRockStabilization = (item) => item?.toLowerCase().includes('rock stabilization')
export const isFormBoard = (item) => item?.toLowerCase().includes('form board')
export const isDrilledHoleGrout = (item) => item?.toLowerCase().includes('drilled hole grout') || item?.toLowerCase().includes('drill hole grout')

/**
 * Parses diameter in inches from string like "5-5/8" Ø" or "6" Ø"
 * @param {string} str - String containing diameter
 * @returns {number} - Diameter in inches
 */
const parseDiameterInches = (str) => {
    if (!str || typeof str !== 'string') return 0
    // Match 5-5/8 or 5-5/8" format (a + b/c inches)
    const fracMatch = str.match(/(\d+)-(\d+)\/(\d+)["']?\s*Ø/i)
    if (fracMatch) {
        const a = parseInt(fracMatch[1], 10) || 0
        const b = parseInt(fracMatch[2], 10) || 0
        const c = parseInt(fracMatch[3], 10) || 1
        return a + (b / c)
    }
    // Match 6" Ø or 5.625" Ø
    const decMatch = str.match(/(\d+\.?\d*)["']?\s*Ø/i)
    if (decMatch) {
        return parseFloat(decMatch[1]) || 0
    }
    return 0
}

/**
 * Parses drilled hole grout parameters
 * e.g. 5-5/8" Ø Drilled hole grout H=22'-6", typ.
 * Grouping by H. F=G=SQRT((d/12)^2*3.14/4), H=height, I=C*H, J=G*F*C, L=J*H/27, M=C
 */
export const parseDrilledHoleGrout = (itemName) => {
    const result = {
        type: 'drilled_hole_grout',
        diameter: 0,
        heightRaw: 0,
        calculatedHeight: 0,
        groupKey: null
    }

    result.diameter = parseDiameterInches(itemName)

    const hMatch = itemName.match(/H=([0-9'"\-]+)/)
    if (hMatch) result.heightRaw = parseDimension(hMatch[1])

    result.calculatedHeight = result.heightRaw

    const hValue = Math.round(result.heightRaw * 12)
    result.groupKey = `drilled-hole-grout-H${hValue}`

    return result
}

/**
 * Universal SOE parser for Secant, Tangent, Sheet piles and Timber lagging
 */
export const parseSoeItem = (itemName) => {
    const result = {
        heightRaw: 0,
        calculatedHeight: 0,
        weight: 0,
        type: null,
    }

    const itemLower = itemName.toLowerCase()

    // Height/LF extraction patterns
    // 1. H=XX'-XX"
    // 2. LF=XX'-XX"
    const hMatch = itemName.match(/H=([0-9'"\-]+)/)
    const lfMatch = itemName.match(/LF=([0-9'"\-]+)/)

    if (hMatch) {
        result.heightRaw = parseDimension(hMatch[1])
    } else if (lfMatch) {
        result.heightRaw = parseDimension(lfMatch[1])
    }

    // Weight extraction from name e.g. W12x58 or MC18x42.7 or WT4x17.5
    // Pattern matches W, MC, or WT followed by digits, then x, then weight (digits and possibly a dot)
    const weightMatch = itemName.match(/(?:W|MC|WT)\d+(?:\.\d+)?x([0-9.]+)/i)
    if (weightMatch) {
        result.weight = parseFloat(weightMatch[1])
    }

    // Type classification and rounding logic
    if (itemLower.includes('supporting angle')) {
        result.type = 'supporting_angle'
        // Extract group key from @ part
        const groupMatch = itemName.match(/@\s*([^)]+)/)
        if (groupMatch) {
            result.groupKey = groupMatch[1].trim()
        }
        // Extract quantity if present e.g. "2 - L8x4x1/2"
        const qtyMatch = itemName.match(/^(\d+)\s*-/)
        if (qtyMatch) {
            result.qty = parseFloat(qtyMatch[1])
        } else {
            result.qty = 1
        }
        // Extract angle size and look up weight
        // Pattern matches L followed by digits, then x, digits, then x, fraction or decimal
        const angleMatch = itemName.match(/L(\d+)x(\d+)x([0-9./½]+)/i)
        if (angleMatch) {
            let d1 = angleMatch[1]
            let d2 = angleMatch[2]
            let d3 = angleMatch[3]
            if (d3 === '½' || d3 === '1/2') d3 = '0.500'
            if (d3 === '3/8') d3 = '0.375'
            if (d3 === '1/4') d3 = '0.250'
            if (d3 === '5/8') d3 = '0.625'
            if (d3 === '3/4') d3 = '0.750'
            if (d3 === '7/8') d3 = '0.875'

            // Handle cases where might be .5 or .500
            const angleKey = `${d1}x${d2}x${parseFloat(d3).toFixed(3)}`
            const angleKeySimple = `${d1}x${d2}x${parseFloat(d3)}`
            result.weight = ANGLE_WEIGHTS[angleKey] || ANGLE_WEIGHTS[angleKeySimple] || 0
        }
    } else if (itemLower.includes('primary secant')) {
        result.type = 'primary_secant'
        result.calculatedHeight = roundToMultipleOf5(result.heightRaw)
        // Extract diameter for concrete weight (e.g. 24" Ø or 24Ø)
        const secantDiameterMatch = itemName.match(/([0-9.]+)["\s]*Ø/i)
        if (secantDiameterMatch) {
            result.diameter = parseFloat(secantDiameterMatch[1])
            // Concrete weight per LF: π * (d/12)² / 4 * 150 pcf ≈ 0.818 * d² lbs/LF
            result.weight = result.weight || (Math.PI * Math.pow(result.diameter / 12, 2) / 4 * 150)
        }
    } else if (itemLower.includes('secondary secant')) {
        result.type = 'secondary_secant'
        result.calculatedHeight = roundToMultipleOf5(result.heightRaw)
    } else if (itemLower.includes('tangent pile')) {
        result.type = 'tangent'
        result.calculatedHeight = roundToMultipleOf5(result.heightRaw)
        // Extract diameter for concrete weight (e.g. 24" Ø or 24Ø)
        const tangentDiameterMatch = itemName.match(/([0-9.]+)["\s]*Ø/i)
        if (tangentDiameterMatch) {
            result.diameter = parseFloat(tangentDiameterMatch[1])
            result.weight = result.weight || (Math.PI * Math.pow(result.diameter / 12, 2) / 4 * 150)
        }
    } else if (itemLower.includes('sheet pile')) {
        result.type = 'sheet_pile'
        result.calculatedHeight = roundToMultipleOf5(result.heightRaw)
    } else if (/\b(timber|wood|wooden)\s+lagging\b/.test(itemLower)) {
        result.type = 'timber_lagging'
        result.calculatedHeight = result.heightRaw
    } else if (itemLower.includes('timber sheeting')) {
        result.type = 'timber_sheeting'
        result.calculatedHeight = result.heightRaw
    } else if (itemLower.includes('waler')) {
        result.type = 'waler'
    } else if (itemLower.includes('upper raker')) {
        result.type = 'upper_raker'
    } else if (itemLower.includes('lower raker')) {
        result.type = 'lower_raker'
    } else if (itemLower.includes('raker')) {
        result.type = 'raker'
    } else if (itemLower.includes('stand off')) {
        result.type = 'stand_off'
    } else if (itemLower.includes('kicker')) {
        result.type = 'kicker'
    } else if (itemLower.includes('channel')) {
        result.type = 'channel'
    } else if (itemLower.includes('roll chock')) {
        result.type = 'roll_chock'
    } else if (itemLower.includes('stud beam')) {
        result.type = 'stud_beam'
    } else if (itemLower.includes('corner brace')) {
        result.type = 'corner_brace'
    } else if (itemLower.includes('knee brace')) {
        result.type = 'knee_brace'
    } else if (itemLower.startsWith('parging')) {
        result.type = 'parging'
    } else if (itemLower.includes('heel block')) {
        result.type = 'heel_block'
        // Dimensions in bracket (4'-0"x5'-0"x3'-0")
        const bracketMatch = itemName.match(/\(([^)]+)\)/)
        if (bracketMatch) {
            const dims = bracketMatch[1].split('x').map(p => parseDimension(p.trim()))
            if (dims.length === 3) {
                result.length = dims[0]
                result.width = dims[1]
                result.height = dims[2]
            }
        }
    } else if (itemLower.includes('underpinning')) {
        result.type = 'underpinning'
        // Underpinning 2'-4"x1'-0" wide, Height=4'-7"
        const lengthMatch = itemName.match(/Underpinning\s*([0-9'"\-]+)/i)
        const widthMatch = itemName.match(/([0-9'"\-]+)\s*wide/i)
        const hMatchUnder = itemName.match(/Height=([0-9'"\-]+)/i)
        if (lengthMatch) result.length = parseDimension(lengthMatch[1])
        if (widthMatch) result.width = parseDimension(widthMatch[1])
        if (hMatchUnder) result.heightRaw = parseDimension(hMatchUnder[1])
    } else if (itemLower.includes('rock anchor')) {
        result.type = 'rock_anchor'
        // Rock anchor (Free length=13'-3" + Bond length= 10'-6")
        const freeLengthMatch = itemName.match(/Free length=([0-9'"\-]+)/i)
        const bondLengthMatch = itemName.match(/Bond length=\s*([0-9'"\-]+)/i)
        if (freeLengthMatch && bondLengthMatch) {
            const freeLength = parseDimension(freeLengthMatch[1])
            const bondLength = parseDimension(bondLengthMatch[1])
            const total = freeLength + bondLength
            result.heightRaw = total
            result.calculatedHeight = roundToMultipleOf5(total) + 5
            result.bondLength = bondLength // Store bond length for formula
        }
    } else if (itemLower.includes('rock bolt')) {
        result.type = 'rock_bolt'
        // Rock bolt @ 7'-0" O.C. (Bond length=10'-0")
        const ocMatch = itemName.match(/@\s*([0-9'"\-]+)\s*['"]?\s*O\.?C\.?/i)
        const bondLengthMatch = itemName.match(/Bond length=([0-9'"\-]+)/i)
        if (ocMatch) result.ocSpacing = parseDimension(ocMatch[1])
        if (bondLengthMatch) {
            const bondLength = parseDimension(bondLengthMatch[1])
            result.bondLength = bondLength // Store bond length for formula
            result.calculatedLength = bondLength + 5
        }
    } else if (itemLower.includes('anchor') && !itemLower.includes('rock anchor') && !itemLower.includes('tie back') && !itemLower.includes('hollow down anchor')) {
        result.type = 'anchor'
        // DSI R51N Hollow bar anchor (Free length=28'-0" + Bond length=20'-0")
        const freeLengthMatch = itemName.match(/Free length=([0-9'"\-]+)/i)
        const bondLengthMatch = itemName.match(/Bond length=\s*([0-9'"\-]+)/i)
        if (freeLengthMatch && bondLengthMatch) {
            const freeLength = parseDimension(freeLengthMatch[1])
            const bondLength = parseDimension(bondLengthMatch[1])
            const total = freeLength + bondLength
            result.heightRaw = total
            result.calculatedHeight = roundToMultipleOf5(total) + 5
        }
    } else if (itemLower.includes('tie back') || itemLower.includes('hollow down anchor')) {
        result.type = 'tie_back'
        // Hollow down anchor (Free length=28'-0" + Bond length=20'-0")
        const freeLengthMatch = itemName.match(/Free length=([0-9'"\-]+)/i)
        const bondLengthMatch = itemName.match(/Bond length=\s*([0-9'"\-]+)/i)
        if (freeLengthMatch && bondLengthMatch) {
            const freeLength = parseDimension(freeLengthMatch[1])
            const bondLength = parseDimension(bondLengthMatch[1])
            const total = freeLength + bondLength
            result.heightRaw = total
            result.calculatedHeight = roundToMultipleOf5(total) + 5
        }
    } else if (itemLower.includes('concrete soil retention pier') || itemLower.includes('retention pier')) {
        result.type = 'concrete_soil_retention_pier'
        // Concrete soil retention pier (4'-0"x4'-0"x14'-2")
        const bracketMatch = itemName.match(/\(([^)]+)\)/)
        if (bracketMatch) {
            const dims = bracketMatch[1].split('x').map(p => parseDimension(p.trim()))
            if (dims.length === 3) {
                result.length = dims[0]
                result.width = dims[1]
                result.height = dims[2]
            }
        }
    } else if (itemLower.includes('guide wall')) {
        result.type = 'guide_wall'
        // Guide wall (4'-6½"x3'-0") - parse width and height
        const bracketMatch = itemName.match(/\(([^)]+)\)/)
        if (bracketMatch) {
            const parts = bracketMatch[1].split('x')
            if (parts.length >= 2) {
                // Parse width - handle 6½" format (e.g., 4'-6½")
                const widthStr = parts[0].trim()
                // Match patterns like: 4'-6½", 4'-6.5", 5'-3½"
                const widthMatch = widthStr.match(/(\d+)['"]?\s*[-]?\s*(\d+)?\s*([½1\/2]|3\/8|1\/4|5\/8|3\/4|7\/8|\.\d+)?["']?/i)
                if (widthMatch) {
                    const feet = parseFloat(widthMatch[1]) || 0
                    let inches = 0
                    let inchesStr = ''
                    const wholeInches = widthMatch[2] ? parseFloat(widthMatch[2]) : 0
                    const fraction = widthMatch[3]

                    if (fraction) {
                        if (fraction === '½' || fraction === '1/2' || fraction === '1/2') {
                            inches = wholeInches + 0.5
                            inchesStr = `${wholeInches}.5`
                        } else if (fraction === '3/8') {
                            inches = wholeInches + 0.375
                            inchesStr = `${wholeInches}.375`
                        } else if (fraction === '1/4') {
                            inches = wholeInches + 0.25
                            inchesStr = `${wholeInches}.25`
                        } else if (fraction === '5/8') {
                            inches = wholeInches + 0.625
                            inchesStr = `${wholeInches}.625`
                        } else if (fraction === '3/4') {
                            inches = wholeInches + 0.75
                            inchesStr = `${wholeInches}.75`
                        } else if (fraction === '7/8') {
                            inches = wholeInches + 0.875
                            inchesStr = `${wholeInches}.875`
                        } else if (fraction.startsWith('.')) {
                            inches = wholeInches + parseFloat(fraction)
                            inchesStr = `${wholeInches}${fraction}`
                        }
                    } else if (wholeInches) {
                        inches = wholeInches
                        inchesStr = `${wholeInches}`
                    }

                    result.width = feet + (inches / 12)
                    // Create formula for width: =feet+(inches/12)
                    if (inchesStr) {
                        result.widthFormula = `${feet}+(${inchesStr}/12)`
                    }
                }
                // Height is typically 3'-0" or similar
                result.heightRaw = parseDimension(parts[1].trim())
            }
        }
    } else if (itemLower.includes('dowel bar') || itemLower.includes('steel dowels bar')) {
        result.type = 'dowel_bar'
        // 4 - #9 Steel dowels bar (H=1'-0" + RS=4'-0")
        const hMatch = itemName.match(/H=([0-9'"\-]+)/i)
        const rsMatch = itemName.match(/RS=([0-9'"\-]+)/i)
        const qtyMatch = itemName.match(/^(\d+)\s*-/)
        if (hMatch) result.hValue = parseDimension(hMatch[1])
        if (rsMatch) result.rsValue = parseDimension(rsMatch[1])
        if (qtyMatch) result.qty = parseFloat(qtyMatch[1])
        if (result.hValue && result.rsValue) {
            result.heightRaw = result.hValue + result.rsValue
        }
    } else if (itemLower.includes('rock pin') || itemLower.includes('rock pins')) {
        result.type = 'rock_pin'
        // Rock pin (H=1'-0" + RS=4'-0")
        const hMatch = itemName.match(/H=([0-9'"\-]+)/i)
        const rsMatch = itemName.match(/RS=([0-9'"\-]+)/i)
        if (hMatch) result.hValue = parseDimension(hMatch[1])
        if (rsMatch) result.rsValue = parseDimension(rsMatch[1])
        if (result.hValue && result.rsValue) {
            result.heightRaw = result.hValue + result.rsValue
        }
        result.qty = 1 // Rock pins typically have QTY = 1
    } else if (itemLower.includes('shotcrete')) {
        result.type = 'shotcrete'
        // Shotcrete w/ wire mesh H=15'-0"
        const hMatch = itemName.match(/H=([0-9'"\-]+)/i)
        if (hMatch) {
            result.heightRaw = parseDimension(hMatch[1])
            // Group by H value
            result.groupKey = `H_${hMatch[1].trim()}`
        }
    } else if (itemLower.includes('permission grouting')) {
        result.type = 'permission_grouting'
        // Permission grouting H=18'-0"
        const hMatch = itemName.match(/H=([0-9'"\-]+)/i)
        if (hMatch) {
            result.heightRaw = parseDimension(hMatch[1])
            // Group by H value
            result.groupKey = `H_${hMatch[1].trim()}`
        }
    } else if (itemLower.includes('concrete button')) {
        result.type = 'button'
        // Concrete button (3'-0"x3'-0"x1'-0")
        const bracketMatch = itemName.match(/\(([^)]+)\)/)
        if (bracketMatch) {
            const dims = bracketMatch[1].split('x').map(p => parseDimension(p.trim()))
            if (dims.length === 3) {
                result.length = dims[0]
                result.width = dims[1]
                result.height = dims[2]
            }
        }
    } else if (itemLower.includes('rock stabilization')) {
        result.type = 'rock_stabilization'
        // Rock stabilization (H=2'-4")
        const hMatch = itemName.match(/H=([0-9'"\-]+)/i)
        if (hMatch) {
            result.heightRaw = parseDimension(hMatch[1])
            // Group by H value
            result.groupKey = `H_${hMatch[1].trim()}`
        }
    } else if (itemLower.includes('form board')) {
        result.type = 'form_board'
        // 1" form board (H=14'-2")
        const thickMatch = itemName.match(/^(\d+["'])/)
        const hMatch = itemName.match(/H=([0-9'"\-]+)/i)
        if (thickMatch) {
            // Group by thickness at start
            result.thickness = thickMatch[1]
            result.groupKey = `THICK_${thickMatch[1]}`
        }
        if (hMatch) result.heightRaw = parseDimension(hMatch[1])
    } else if (itemLower.includes('rock bolt')) {
        result.type = 'rock_bolt'
        // Rock bolt @ 7'-0" O.C. (Bond length=10'-0")
        const spacingMatch = itemName.match(/@\s*([0-9'\-"]+)\s*o\.?c/i)
        const bondLengthMatch = itemName.match(/Bond length=([0-9'"\-]+)/i)
        if (spacingMatch) {
            result.spacing = parseDimension(spacingMatch[1])
            result.ocSpacing = result.spacing // For QTY formula: ROUNDUP(C/OC,0)+1
            // Group by spacing value
            result.groupKey = `SPACING_${spacingMatch[1].trim()}`
        }
        if (bondLengthMatch) {
            result.bondLength = parseDimension(bondLengthMatch[1])
            result.calculatedLength = result.bondLength + 5 // Length(F)=bond length+5
        }
    }

    return result
}

export default {
    parseSoldierPile,
    parseTimberSoldierPile,
    parseTimberPlank,
    parseTimberRaker,
    parseTimberBrace,
    parseTimberWaler,
    parseTimberStringer,
    parseTimberPost,
    parseVerticalTimberSheets,
    parseHorizontalTimberSheets,
    calculatePileWeight,
    isSoldierPile,
    isTimberSoldierPile,
    isVerticalTimberSheets,
    isHorizontalTimberSheets,
    isTimberPlank,
    isTimberRaker,
    isTimberPost,
    isTimberBrace,
    isTimberWaler,
    isTimberStringer,
    isPrimarySecantPile,
    isSecondarySecantPile,
    isTangentPile,
    isSheetPile,
    isTimberLagging,
    isTimberSheeting,
    isWaler,
    isRaker,
    isUpperRaker,
    isLowerRaker,
    isStandOff,
    isKicker,
    isChannel,
    isRollChock,
    isStudBeam,
    isCornerbrace,
    isKneeBrace,
    isSupportingAngle,
    isParging,
    isHeelBlock,
    isUnderpinning,
    isRockAnchor,
    isRockBolt,
    isAnchor,
    isTieBack,
    isConcreteSoilRetentionPier,
    isGuideWall,
    isDowelBar,
    isRockPin,
    isShotcrete,
    isPermissionGrouting,
    isButton,
    isRockStabilization,
    isFormBoard,
    isDrilledHoleGrout,
    parseDrilledHoleGrout,
    parseSoeItem,
    parseDimension,
    roundToMultipleOf5
}
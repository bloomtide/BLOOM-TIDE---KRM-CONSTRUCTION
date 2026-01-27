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
    return 0
}

export const isSoldierPile = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false
    const itemLower = digitizerItem.toLowerCase()

    // Exclude supporting angle items
    if (itemLower.includes('supporting angle')) return false

    return (itemLower.includes('drilled soldier pile') || itemLower.includes('soldier pile')) &&
        !itemLower.includes('secant') &&
        !itemLower.includes('tangent')
}

/**
 * Identifies if item is a primary secant pile
 */
export const isPrimarySecantPile = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false
    const itemLower = digitizerItem.toLowerCase()
    return itemLower.includes('primary secant pile')
}

/**
 * Identifies if item is a secondary secant pile
 */
export const isSecondarySecantPile = (digitizerItem) => {
    if (!digitizerItem || typeof digitizerItem !== 'string') return false
    const itemLower = digitizerItem.toLowerCase()
    return itemLower.includes('secondary secant pile')
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

    return itemLower.includes('timber lagging')
}

/**
 * New identifiers for additional SOE sections
 */
export const isTimberSheeting = (item) => item?.toLowerCase().includes('timber sheeting')
export const isWaler = (item) => item?.toLowerCase().includes('waler')
export const isRaker = (item) => item?.toLowerCase().includes('raker') && !item?.toLowerCase().includes('upper') && !item?.toLowerCase().includes('lower')
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
export const isInnerCornerBrace = (item) => item?.toLowerCase().includes('inner corner brace')
export const isKneeBrace = (item) => item?.toLowerCase().includes('knee brace')

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
    if (itemLower.includes('primary secant')) {
        result.type = 'primary_secant'
        result.calculatedHeight = roundToMultipleOf5(result.heightRaw)
    } else if (itemLower.includes('secondary secant')) {
        result.type = 'secondary_secant'
        result.calculatedHeight = roundToMultipleOf5(result.heightRaw)
    } else if (itemLower.includes('tangent pile')) {
        result.type = 'tangent'
        result.calculatedHeight = roundToMultipleOf5(result.heightRaw)
    } else if (itemLower.includes('sheet pile')) {
        result.type = 'sheet_pile'
        result.calculatedHeight = roundToMultipleOf5(result.heightRaw)
    } else if (itemLower.includes('timber lagging')) {
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
    } else if (itemLower.includes('inner corner brace')) {
        result.type = 'inner_corner_brace'
    } else if (itemLower.includes('knee brace')) {
        result.type = 'knee_brace'
    }

    return result
}

export default {
    parseSoldierPile,
    calculatePileWeight,
    isSoldierPile,
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
    isInnerCornerBrace,
    isKneeBrace,
    parseSoeItem,
    parseDimension,
    roundToMultipleOf5
}

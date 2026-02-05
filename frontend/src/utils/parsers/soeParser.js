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
export const isSupportingAngle = (item) => item?.toLowerCase().includes('supporting angle')
export const isParging = (item) => item?.toLowerCase().startsWith('parging')
export const isHeelBlock = (item) => item?.toLowerCase().includes('heel block')
export const isUnderpinning = (item) => item?.toLowerCase().includes('underpinning')
export const isRockAnchor = (item) => item?.toLowerCase().includes('rock anchor')
export const isRockBolt = (item) => item?.toLowerCase().includes('rock bolt')
export const isAnchor = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('anchor') && !itemLower.includes('rock anchor') && !itemLower.includes('tie back') && !itemLower.includes('hollow down anchor')
}
export const isTieBack = (item) => item?.toLowerCase().includes('tie back') || item?.toLowerCase().includes('hollow down anchor')
export const isConcreteSoilRetentionPier = (item) => item?.toLowerCase().includes('concrete soil retention pier')
export const isGuideWall = (item) => item?.toLowerCase().includes('guide wall')
export const isDowelBar = (item) => item?.toLowerCase().includes('dowel bar') || item?.toLowerCase().includes('steel dowels bar')
export const isRockPin = (item) => item?.toLowerCase().includes('rock pin')
export const isShotcrete = (item) => item?.toLowerCase().includes('shotcrete')
export const isPermissionGrouting = (item) => item?.toLowerCase().includes('permission grouting')
export const isButton = (item) => item?.toLowerCase().includes('concrete button')
export const isRockStabilization = (item) => item?.toLowerCase().includes('rock stabilization')
export const isFormBoard = (item) => item?.toLowerCase().includes('form board')

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
    } else if (itemLower.includes('concrete soil retention pier')) {
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
    } else if (itemLower.includes('rock pin')) {
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
        if (hMatch) result.heightRaw = parseDimension(hMatch[1])
    } else if (itemLower.includes('permission grouting')) {
        result.type = 'permission_grouting'
        // Permission grouting H=18'-0"
        const hMatch = itemName.match(/H=([0-9'"\-]+)/i)
        if (hMatch) result.heightRaw = parseDimension(hMatch[1])
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
        if (hMatch) result.heightRaw = parseDimension(hMatch[1])
    } else if (itemLower.includes('form board')) {
        result.type = 'form_board'
        // 1" form board (H=14'-2")
        const hMatch = itemName.match(/H=([0-9'"\-]+)/i)
        if (hMatch) result.heightRaw = parseDimension(hMatch[1])
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
    parseSoeItem,
    parseDimension,
    roundToMultipleOf5
}

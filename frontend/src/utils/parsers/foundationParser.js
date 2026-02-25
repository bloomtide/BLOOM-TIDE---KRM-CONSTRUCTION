import { convertToFeet } from './dimensionParser.js'

/**
 * Extracts numeric value from dimension string (e.g., "27'-10"" -> 27.833)
 * Uses the proper convertToFeet function from dimensionParser
 * @param {string} dimStr - Dimension string
 * @returns {number} - Decimal feet value
 */
const parseDimension = (dimStr) => {
    if (!dimStr) return 0
    return convertToFeet(dimStr)
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
 * Parser for Foundation items
 */

/**
 * Calculates weight of pile using formula: (diameter - thickness) * thickness * 10.69
 * @param {number} diameter - Diameter in inches
 * @param {number} thickness - Thickness in inches (default 0.5 if not specified)
 * @returns {number} - Weight in lbs/ft
 */
export const calculatePileWeight = (diameter, thickness = 0.5) => {
    return (diameter - thickness) * thickness * 10.69
}

/**
 * Parses diameter and thickness from string like "9-5/8" Øx0.545"" or "4½" Øx0.408""
 */
const parseDiameterThickness = (str) => {
    // First, handle Unicode fractions like "4½" by converting to "4-1/2"
    let normalizedStr = str
    const unicodeFractions = {
        '½': '1/2',
        '⅓': '1/3',
        '⅔': '2/3',
        '¼': '1/4',
        '¾': '3/4',
        '⅕': '1/5',
        '⅖': '2/5',
        '⅗': '3/5',
        '⅘': '4/5',
        '⅙': '1/6',
        '⅚': '5/6',
        '⅛': '1/8',
        '⅜': '3/8',
        '⅝': '5/8',
        '⅞': '7/8'
    }
    
    // Replace Unicode fractions with standard fractions
    for (const [unicode, fraction] of Object.entries(unicodeFractions)) {
        normalizedStr = normalizedStr.replace(new RegExp(`(\\d+)${unicode.replace(/[.*+?^${}()|[\]\\]/g, '\\$&')}`, 'g'), `$1-${fraction}`)
    }
    
    // Match patterns like "9-5/8" Øx0.545"" or "7"Øx0.408"" or "4-1/2" Øx0.408""
    const match = normalizedStr.match(/(\d+(?:-\d+\/\d+)?)["']?\s*[ØØ]\s*x\s*([0-9.]+)/i)
    if (match) {
        const diameterStr = match[1]
        let thickness = parseFloat(match[2])
        // Normalize typo "0545" (missing decimal) -> 0.545; "0408" -> 0.408
        if (thickness >= 100 && thickness < 1000 && Number.isInteger(thickness)) {
            thickness = thickness / 1000
        }
        
        // Parse diameter - handle fractions like "9-5/8" or "4-1/2"
        let diameter = 0
        if (diameterStr.includes('-')) {
            const parts = diameterStr.split('-')
            const whole = parseFloat(parts[0]) || 0
            const fraction = parts[1]
            if (fraction) {
                const fracMatch = fraction.match(/(\d+)\/(\d+)/)
                if (fracMatch) {
                    const num = parseFloat(fracMatch[1])
                    const den = parseFloat(fracMatch[2])
                    diameter = whole + (num / den)
                } else {
                    diameter = whole
                }
            } else {
                diameter = whole
            }
        } else {
            diameter = parseFloat(diameterStr)
        }
        
        return { diameter, thickness }
    }
    return null
}

/**
 * Parses single diameter like "13-3/8" Ø"
 */
const parseSingleDiameter = (str) => {
    const match = str.match(/(\d+(?:-\d+\/\d+)?)["']?\s*[ØØ]/i)
    if (match) {
        const diameterStr = match[1]
        let diameter = 0
        if (diameterStr.includes('-')) {
            const parts = diameterStr.split('-')
            const whole = parseFloat(parts[0]) || 0
            const fraction = parts[1]
            if (fraction) {
                const fracMatch = fraction.match(/(\d+)\/(\d+)/)
                if (fracMatch) {
                    const num = parseFloat(fracMatch[1])
                    const den = parseFloat(fracMatch[2])
                    diameter = whole + (num / den)
                } else {
                    diameter = whole
                }
            } else {
                diameter = whole
            }
        } else {
            diameter = parseFloat(diameterStr)
        }
        return diameter
    }
    return null
}

/**
 * Identifies if item is a drilled foundation pile
 */
export const isDrilledFoundationPile = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return (itemLower.includes('drilled') && itemLower.includes('foundation pile')) ||
           (itemLower.includes('drilled') && itemLower.includes('cassion pile')) ||
           /\b(drilled|structural|foundation|fndt)\s+piles?\b/i.test(item)
}

/**
 * Identifies if item is a helical foundation pile
 */
export const isHelicalFoundationPile = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('helical') && (itemLower.includes('foundation pile') || itemLower.includes('pile'))
}

/**
 * Identifies if item is a driven foundation pile
 */
export const isDrivenFoundationPile = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('driven') && (itemLower.includes('foundation pile') || itemLower.includes('pile'))
}

/**
 * Identifies if item is a stelcor drilled displacement pile
 */
export const isStelcorDrilledDisplacementPile = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('stelcor') && itemLower.includes('drilled displacement pile')
}

/**
 * Identifies if item is a CFA pile
 */
export const isCFAPile = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('cfa pile')
}

/**
 * Identifies if item is a miscellaneous pile (contains "pile" but doesn't fit
 * Drilled/Helical/Driven/Stelcor/CFA foundation piles or SOE pile types).
 * Excludes: drilled soldier pile, soldier pile, primary/secondary secant pile, tangent pile, sheet pile.
 */
export const isMiscellaneousFoundationPile = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    if (!itemLower.includes('pile')) return false
    if (isDrilledFoundationPile(item) || isHelicalFoundationPile(item) || isDrivenFoundationPile(item) ||
        isStelcorDrilledDisplacementPile(item) || isCFAPile(item)) return false
    if (itemLower.includes('drilled soldier pile') || itemLower.includes('soldier pile')) return false
    if (itemLower.includes('secant pile') || itemLower.includes('tangent pile') || itemLower.includes('sheet pile')) return false
    return true
}

/**
 * Identifies if item is a pile cap
 */
export const isPileCap = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('pile cap') || itemLower.startsWith('pc-')
}

/**
 * Identifies if item is a strip footing
 */
export const isStripFooting = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('strip footing') || itemLower.startsWith('sf') || itemLower.startsWith('st-') || itemLower.startsWith('wf-') || itemLower.includes('wall footing')
}

/**
 * Identifies if item is an isolated footing
 */
export const isIsolatedFooting = (item) => {
  if (!item || typeof item !== 'string') return false
  const itemLower = item.toLowerCase()

  return (
    (itemLower.startsWith('f-') || itemLower.includes('footing')) && !itemLower.includes('foundation')
  )
}

/**
 * Identifies if item is a pilaster
 */
export const isPilaster = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('pilaster')
}

/**
 * Identifies if item is a grade beam
 */
export const isGradeBeam = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('grade beam') || itemLower.startsWith('gb') || itemLower.startsWith('gb-')
}

/**
 * Identifies if item is a tie beam
 */
export const isTieBeam = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('tie beam') || itemLower.startsWith('tb')
}

/**
 * Identifies if item is a strap beam
 * Format: ST (3'-10"x2'-9") typ. or ST (2'-8"x3'-0") (87.86')
 */
export const isStrapBeam = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase().trim()
    return (itemLower.startsWith('st ') || /^st\s*\(/i.test(itemLower) || itemLower.includes('strap beam')) && !itemLower.startsWith('st-')
}

/**
 * Identifies if item is a thickened slab
 */
export const isThickenedSlab = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('thickened slab')
}

/**
 * Identifies if item is a buttress
 */
export const isButtress = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('buttress')
}

/**
 * Identifies if item is a pier
 */
export const isPier = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase().trim()
    // Only treat items explicitly named "Pier ..." as piers.
    // This excludes things like "Concrete soil retention pier".
    return itemLower.startsWith('pier') || itemLower.startsWith('concrete pier')
}

/**
 * Identifies if item is a corbel
 */
export const isCorbel = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('corbel')
}

/**
 * Identifies if item is a linear wall
 */
export const isLinearWall = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('linear wall') || itemLower.includes('liner wall')
}

/**
 * Identifies if item is a foundation wall
 */
export const isFoundationWall = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return (itemLower.includes('foundation wall') || itemLower.startsWith('fw') || itemLower.includes('fndt wall')) && !itemLower.includes('retaining')
}

/**
 * Identifies if item is a retaining wall
 */
export const isRetainingWall = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('retaining wall') || itemLower.startsWith('rw')
}

/**
 * Identifies if item is a barrier wall
 */
export const isBarrierWall = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('barrier wall') || itemLower.includes('vehicle barrier')
}

/**
 * Identifies if item is a stem wall
 */
export const isStemWall = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('stem wall')
}

/**
 * Identifies if item is an elevator pit item
 * Excludes service elevator pit items - used in Service elevator pit subsection
 * Note: "pit" is optional - "elev. slab" is treated as "elev. pit slab"
 */
export const isElevatorPit = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    // Exclude service elevator pit items
    if (itemLower.includes('sump pit @ service elevator')) return false
    if (itemLower.includes('service elev. pit') || itemLower.includes('service elevator pit')) return false
    if (itemLower.includes('service elev.') || itemLower.includes('service elevator')) return false
    // Check for explicit elevator pit
    if (itemLower.includes('elev. pit') || itemLower.includes('elevator pit')) return true
    // Check for sump pit @ elevator variants
    if (itemLower.includes('sump pit @ elevator')) return true
    if (itemLower.includes('sump pit')) return true
    // Check for elev/elevator followed by keywords (pit is optional)
    const elevPattern = /(elev\.?|elevator)\s+(slab|mat|wall|slope|haunch|sump)/i
    return elevPattern.test(itemLower)
}

/**
 * Identifies if item is a service elevator pit item
 * Sump pit @ service elevator, Service Elev. pit slab, Service Elev. pit wall, etc.
 * Note: "pit" is optional - "service elev. slab" is treated as "service elev. pit slab"
 */
export const isServiceElevatorPit = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    // Check for "sump pit @ service elevator" variants
    if (itemLower.includes('sump pit @ service elevator')) return true
    // Check for explicit service elevator pit
    if (itemLower.includes('service elev. pit') || itemLower.includes('service elevator pit')) return true
    // Check for service elev/elevator followed by keywords (pit is optional)
    const serviceElevPattern = /service\s+(elev\.?|elevator)\s+(slab|mat|wall|slope|haunch|sump)/i
    return serviceElevPattern.test(itemLower)
}

/**
 * Identifies if item is a detention tank item
 */
export const isDetentionTank = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('detention tank')
}

/**
 * Identifies if item is a duplex sewage ejector pit item
 * Note: "pit" is optional - "duplex sewage ejector slab" = "duplex sewage ejector pit slab"
 */
export const isDuplexSewageEjectorPit = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    // Explicit check
    if (itemLower.includes('duplex sewage ejector pit')) return true
    // Make pit optional: duplex sewage ejector + keyword
    const dupPattern = /duplex\s+sewage\s+ejector\s+(slab|mat|wall|slope|haunch|sump)/i
    return dupPattern.test(itemLower)
}

/**
 * Identifies if item is a deep sewage ejector pit item
 * Note: "pit" is optional - "deep sewage ejector slab" = "deep sewage ejector pit slab"
 */
export const isDeepSewageEjectorPit = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    // Explicit check
    if (itemLower.includes('deep sewage ejector pit')) return true
    // Make pit optional: deep sewage ejector + keyword
    const deepPattern = /(?:deep\s+sewage\s+)?ejector\s+(slab|mat|wall|slope|haunch|sump|pit)/i
    return deepPattern.test(itemLower)
}

/**
 * Identifies if item is a sump pump pit item
 * Note: "pit" is optional - "sump pump slab" = "sump pump pit slab"
 */
export const isSumpPumpPit = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    // Exclude items if they're pumps but not pits (e.g., sump pump equipment)
    if (itemLower.includes('sump pump') && !itemLower.includes('pit')) {
        // Check if it has pit-related keywords
        const sumpPattern = /sump\s+pump(?:\s+pit)?\s+(slab|mat|wall|slope|haunch)/i
        return sumpPattern.test(itemLower)
    }
    return itemLower.includes('sump pump pit') || itemLower.includes('sump pump')
}

/**
 * Identifies if item is a grease trap item
 * Note: "pit" is optional - "grease trap slab" = "grease trap pit slab"
 */
export const isGreaseTrap = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    // Explicit check
    if (itemLower.includes('grease trap')) return true
    // Make pit optional: grease trap + keyword
    const greasePattern = /grease\s+trap(?:\s+pit)?\s+(slab|mat|wall|slope|haunch)/i
    return greasePattern.test(itemLower)
}

/**
 * Identifies if item is a house trap item
 * Note: "pit" is optional - "house trap slab" = "house trap pit slab"
 */
export const isHouseTrap = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    // Explicit check
    if (itemLower.includes('house trap')) return true
    // Make pit optional: house trap + keyword
    const housePattern = /house\s+trap(?:\s+pit)?\s+(slab|mat|wall|slope|haunch)/i
    return housePattern.test(itemLower)
}

/**
 * Identifies if item is a mat slab item
 */
export const isMatSlab = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    // Exclude all pit-type items as they should be handled by their specific identifiers
    if (itemLower.includes('elevator pit') || itemLower.includes('elev. pit') || 
        itemLower.includes('service elevator pit') ||
        itemLower.includes('duplex sewage ejector') ||
        itemLower.includes('deep sewage ejector') ||
        itemLower.includes('sump pump pit') ||
        itemLower.includes('grease trap') ||
        itemLower.includes('house trap')) return false
    return itemLower.includes('mat') && (itemLower.includes('haunch') || itemLower.match(/mat(?:[-\s]+slab)?[-\s]*\d+/i))
}

/**
 * Identifies if item is a mud slab item (Foundation section)
 */
export const isMudSlabFoundation = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    // Only match exact "mud slab" for Foundation section (not "w/ X" mud slab" which is Excavation)
    return itemLower === 'mud slab' || itemLower.trim() === 'mud slab' || itemLower === 'mud mat' || itemLower.trim() === 'mud mat'
}

/**
 * Identifies if item is a SOG item
 */
export const isSOG = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    // Exclude items with "Demo" in the name
    if (itemLower.includes('demo')) return false
    return itemLower.includes('sog') || itemLower.includes('gravel') || itemLower.includes('geotextile filter fabric') || itemLower.includes('slab on grade')
}

/**
 * Identifies if item is a ROG (Ramp on grade) item
 */
export const isROG = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    if (itemLower.includes('demo')) return false
    return itemLower.includes('rog') || itemLower.includes('ramp on grade')
}

/**
 * Identifies if item is a stairs on grade item
 */
export const isStairsOnGrade = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('stairs on grade') || itemLower.includes('landings on grade')
}

/**
 * Identifies if item is an electric conduit item (Foundation section)
 */
export const isElectricConduit = (item) => {
    if (!item || typeof item !== 'string') return false
    const itemLower = item.toLowerCase()
    return itemLower.includes('underground electric conduit') ||
        itemLower.includes('electric conduit in slab') ||
        itemLower.includes('trench drain') ||
        itemLower.includes('perforated pipe')
}

/**
 * Parses electric conduit item (no extra fields; type only)
 */
export const parseElectricConduit = (itemName) => ({
    type: 'electric_conduit'
})

/**
 * Builds groupKey for drilled foundation pile based on H and RS values only.
 * Groups items with same height and rock socket regardless of diameter string typos (e.g. 0545 vs 0.545).
 */
const buildDrilledPileGroupKey = (parsed, isDual, hasInfluence = false) => {
    let prefix = isDual ? 'DUAL' : 'SINGLE'
    if (hasInfluence) {
        prefix = `INFLU-${prefix}`
    }
    const parts = []
    if (parsed.height != null && parsed.height > 0) {
        parts.push(`H${parsed.height.toFixed(2)}`)
    }
    if (parsed.rockSocket != null && parsed.rockSocket > 0) {
        parts.push(`RS${parsed.rockSocket.toFixed(2)}`)
    }
    return parts.length > 0 ? `${prefix}-${parts.join('-')}` : `${prefix}-OTHER`
}

/**
 * Parses drilled foundation pile
 */
export const parseDrilledFoundationPile = (itemName) => {
    const result = {
        type: 'drilled_foundation_pile',
        diameter: null,
        thickness: null,
        height: 0,
        rockSocket: null,
        calculatedHeight: 0,
        weight: 0,
        isDualDiameter: false,
        diameter2: null,
        weight2: 0,
        hasInfluence: false,
        groupKey: null
    }

    const itemLower = itemName.toLowerCase()
    if (itemLower.includes('influence')) {
        result.hasInfluence = true
    }

    // Extract H and RS values
    // Supports formats like:
    // - "H=32'-6\"+ 7'-0\" RS"  (RS without '=', after +)
    // - "H=32'-6\" RS=7'-0\""  (RS with '=')
    // - "H=32'-6\"+18'-3\" RS"  (no space after +)
    let hStr = null
    let rsStr = null

    // First, try to match the full pattern: H=<dim>+ <dim> RS or H=<dim>+<dim> RS
    // Match H= followed by dimension (digits, dash, digits, quote, dash, digits, quote), 
    // then optionally + followed by dimension and RS
    // Pattern: H=32'-6"+ 7'-0" RS or H=32'-6"+18'-3" RS
    // Use a more specific pattern: match the exact dimension format
    // Dimension format: \d+'-\d+" (e.g., "32'-6"")
    // This ensures we capture the full dimension, not just the first digit
    const fullMatch = itemName.match(/H\s*=\s*(\d+'-\d+")\s*(?:\+\s*(\d+'-\d+")\s*RS\b)?/i)
    if (fullMatch) {
        hStr = (fullMatch[1] || '').trim()
        // Remove trailing quote if present
        hStr = hStr.replace(/["\s]+$/, '').trim()
        if (fullMatch[2]) {
            rsStr = (fullMatch[2] || '').trim()
            rsStr = rsStr.replace(/["\s]+$/, '').trim()
        }
    }

    // If RS wasn't captured above, try "RS=<dim>" format
    if (!rsStr) {
        const rsEqMatch = itemName.match(/RS\s*=\s*(\d+'-\d+")/i)
        if (rsEqMatch) {
            rsStr = (rsEqMatch[1] || '').trim()
            rsStr = rsStr.replace(/["\s]+$/, '').trim()
        }
    }

    // Fallback: try to find "+ <dim> RS" pattern anywhere (more flexible)
    if (!rsStr) {
        // Match: + followed by dimension followed by RS
        // Dimension pattern: \d+'-\d+" which matches "7'-0"" or "18'-3""
        const rsPlusMatch = itemName.match(/\+\s*(\d+'-\d+")\s*RS\b/i)
        if (rsPlusMatch) {
            rsStr = (rsPlusMatch[1] || '').trim()
            rsStr = rsStr.replace(/["\s]+$/, '').trim()
        }
    }

    // Parse the extracted strings using convertToFeet
    if (hStr) {
        result.height = parseDimension(hStr)
    }
    if (rsStr) {
        result.rockSocket = parseDimension(rsStr)
    }

    // Calculate total height
    if (result.rockSocket && result.height) {
        const total = result.height + result.rockSocket
        result.calculatedHeight = roundToMultipleOf5(total)
    } else if (result.height) {
        result.calculatedHeight = roundToMultipleOf5(result.height)
    }

    // Check for dual diameter (e.g., "9-5/8" Øx0.545" & 13-3/8" Ø")
    const dualMatch = itemName.match(/([^&]+)\s*&\s*([^&]+)/i)
    if (dualMatch) {
        result.isDualDiameter = true
        const firstPart = dualMatch[1].trim()
        const secondPart = dualMatch[2].trim()
        
        const dt1 = parseDiameterThickness(firstPart)
        if (dt1) {
            result.diameter = dt1.diameter
            result.thickness = dt1.thickness
            result.weight = calculatePileWeight(dt1.diameter, dt1.thickness)
        }
        
        const diameter2 = parseSingleDiameter(secondPart)
        if (diameter2) {
            result.diameter2 = diameter2
            result.weight2 = calculatePileWeight(diameter2, 0.5) // Default thickness 0.5
        }
        
        // Group key: dual diameter + H + RS (ignores diameter typos like 0545 vs 0.545)
        result.groupKey = buildDrilledPileGroupKey(result, true, result.hasInfluence)
    } else {
        // Single diameter
        const dt = parseDiameterThickness(itemName)
        if (dt) {
            result.diameter = dt.diameter
            result.thickness = dt.thickness
            result.weight = calculatePileWeight(dt.diameter, dt.thickness)
        }
        // Group key: single diameter + H + RS (ignores diameter typos like 0545 vs 0.545)
        result.groupKey = buildDrilledPileGroupKey(result, false, result.hasInfluence)
    }

    return result
}

/**
 * Parses helical foundation pile
 */
export const parseHelicalFoundationPile = (itemName) => {
    const result = {
        type: 'helical_foundation_pile',
        diameter: null,
        thickness: null,
        height: 0,
        calculatedHeight: 0,
        weight: 0,
        hasInfluence: false,
        groupKey: null
    }

    // Detect influence keyword
    const itemLower = itemName.toLowerCase()
    if (itemLower.includes('influence')) {
        result.hasInfluence = true
    }

    // Extract H value
    const hMatch = itemName.match(/H=([0-9'"\-]+)/i)
    if (hMatch) {
        result.height = parseDimension(hMatch[1])
        result.calculatedHeight = roundToMultipleOf5(result.height)
    }

    // Parse diameter and thickness
    const dt = parseDiameterThickness(itemName)
    if (dt) {
        result.diameter = dt.diameter
        result.thickness = dt.thickness
        result.weight = calculatePileWeight(dt.diameter, dt.thickness)
        let groupKeyBase = `${dt.diameter.toFixed(3)}x${dt.thickness}`
        // Add INFLU- prefix to groupKey if item has influence
        result.groupKey = result.hasInfluence ? `INFLU-${groupKeyBase}` : groupKeyBase
    }

    return result
}

/**
 * Parses driven foundation pile
 */
export const parseDrivenFoundationPile = (itemName) => {
    const result = {
        type: 'driven_foundation_pile',
        hpSize: null,
        weight: 0,
        height: 0,
        calculatedHeight: 0,
        hasInfluence: false,
        groupKey: null
    }

    // Detect influence keyword
    const itemLower = itemName.toLowerCase()
    if (itemLower.includes('influence')) {
        result.hasInfluence = true
    }

    // Extract HP size (e.g., HP12x74)
    const hpMatch = itemName.match(/HP(\d+)x(\d+)/i)
    if (hpMatch) {
        result.hpSize = hpMatch[0]
        result.weight = parseFloat(hpMatch[2]) // Weight is the second number
    }

    // Extract H value
    const hMatch = itemName.match(/H=([0-9'"\-]+)/i)
    if (hMatch) {
        result.height = parseDimension(hMatch[1])
        result.calculatedHeight = roundToMultipleOf5(result.height)
    }

    let groupKeyBase = result.hpSize || 'OTHER'
    // Add INFLU- prefix to groupKey if item has influence
    result.groupKey = result.hasInfluence ? `INFLU-${groupKeyBase}` : groupKeyBase

    return result
}

/**
 * Parses stelcor drilled displacement pile
 */
export const parseStelcorDrilledDisplacementPile = (itemName) => {
    const result = {
        type: 'stelcor_drilled_displacement_pile',
        diameter: null,
        thickness: null,
        height: 0,
        calculatedHeight: 0,
        weight: 0,
        hasInfluence: false,
        groupKey: null
    }

    // Detect influence keyword
    const itemLower = itemName.toLowerCase()
    if (itemLower.includes('influence')) {
        result.hasInfluence = true
    }

    // Extract H value
    const hMatch = itemName.match(/H=([0-9'"\-]+)/i)
    if (hMatch) {
        result.height = parseDimension(hMatch[1])
        result.calculatedHeight = roundToMultipleOf5(result.height)
    }

    // Parse diameter and thickness
    const dt = parseDiameterThickness(itemName)
    if (dt) {
        result.diameter = dt.diameter
        result.thickness = dt.thickness
        result.weight = calculatePileWeight(dt.diameter, dt.thickness)
        let groupKeyBase = `${dt.diameter.toFixed(3)}x${dt.thickness}`
        result.groupKey = result.hasInfluence ? `INFLU-${groupKeyBase}` : groupKeyBase
    }

    return result
}

/**
 * Parses CFA pile
 */
export const parseCFAPile = (itemName) => {
    const result = {
        type: 'cfa_pile',
        height: 0,
        calculatedHeight: 0,
        hasInfluence: false
    }

    // Detect influence keyword
    const itemLower = itemName.toLowerCase()
    if (itemLower.includes('influence')) {
        result.hasInfluence = true
    }

    // Extract H value
    const hMatch = itemName.match(/H=([0-9'"\-]+)/i)
    if (hMatch) {
        result.height = parseDimension(hMatch[1])
        result.calculatedHeight = roundToMultipleOf5(result.height)
    }

    return result
}

/**
 * Tries to match a miscellaneous pile item's structure to one of the known pile types.
 * Used for Piles subsection: if structure matches Drilled/Helical/Driven/Stelcor/CFA, use that pile's formulas.
 * If no match, the item should NOT be added (stays in unused data).
 * @param {string} itemName - The item name (e.g. from Digitizer Item)
 * @returns {{ matchedType: string, parsed: object } | null} - Matched pile type and parsed data, or null if no structure match
 */
export const tryMatchPileStructure = (itemName) => {
    if (!itemName || typeof itemName !== 'string') return null

    // Try Drilled - needs (diameter or dual diameter) and (height or calculatedHeight)
    const drilled = parseDrilledFoundationPile(itemName)
    if ((drilled.diameter || drilled.isDualDiameter) && (drilled.height || drilled.calculatedHeight)) {
        return { matchedType: 'drilled_foundation_pile', parsed: drilled }
    }

    // Try Driven - needs hpSize (HP12x74 pattern) and height
    const driven = parseDrivenFoundationPile(itemName)
    if (driven.hpSize && (driven.height || driven.calculatedHeight)) {
        return { matchedType: 'driven_foundation_pile', parsed: driven }
    }

    // Try Helical - needs diameter and height
    const helical = parseHelicalFoundationPile(itemName)
    if (helical.diameter && (helical.height || helical.calculatedHeight)) {
        return { matchedType: 'helical_foundation_pile', parsed: helical }
    }

    // Try Stelcor - needs diameter and height
    const stelcor = parseStelcorDrilledDisplacementPile(itemName)
    if (stelcor.diameter && (stelcor.height || stelcor.calculatedHeight)) {
        return { matchedType: 'stelcor_drilled_displacement_pile', parsed: stelcor }
    }

    // Try CFA - needs height only (H= dimension)
    const cfa = parseCFAPile(itemName)
    if (cfa.height || cfa.calculatedHeight) {
        return { matchedType: 'cfa_pile', parsed: cfa }
    }

    return null
}

/**
 * Parses dimensions from bracket (e.g., "(4'-6"x4"-6"x3'-4")")
 * Handles both feet-inches format and inches-only format
 */
const parseBracketDimensions = (itemName) => {
    // Some items have multiple parentheses groups, e.g. "Pilaster (P3) (22\"x16\"x6'-0\")".
    // We want the group that actually contains dimensions (the one with "x").
    const matches = Array.from(itemName.matchAll(/\(([^)]+)\)/g))
    if (!matches.length) return null

    const candidates = matches.map(m => (m[1] || '').trim()).filter(Boolean)
    if (!candidates.length) return null

    // Prefer the last parentheses group that contains an 'x' (dimension separator)
    const bracketContent =
        [...candidates].reverse().find(c => c.toLowerCase().includes('x')) || candidates[0]

    // Split by 'x'
    const parts = bracketContent.split('x').map(p => p.trim())
    const dims = parts.map((p, index) => {
        let trimmed = p.trim()

        // Handle case where there's a leading apostrophe before a digit (e.g., "'1'-0"" -> "1'-0"")
        trimmed = trimmed.replace(/^'+(?=\d)/, '')

        // Feet-inches format
        if (trimmed.includes("'")) {
            const result = parseDimension(trimmed)
            return result
        }

        // Inches-only format (e.g., 22")
        if (trimmed.match(/^\d+["']?$/)) {
            const result = parseFloat(trimmed.replace(/["']/g, '')) / 12
            return result
        }

        // Fallback
        const result = parseDimension(trimmed)
        return result
    })

    return dims
}

/**
 * Parses pile cap
 */
export const parsePileCap = (itemName) => {
    const result = {
        type: 'pile_cap',
        length: 0,
        width: 0,
        height: 0
    }

    const dims = parseBracketDimensions(itemName)
    if (dims && dims.length >= 3) {
        result.length = dims[0]
        result.width = dims[1]
        result.height = dims[2]
    }

    return result
}

/**
 * Parses strip footing
 * From brackets, extracts width (first value) and height (second value)
 */
export const parseStripFooting = (itemName) => {
    const result = {
        type: 'strip_footing',
        width: 0,
        height: 0,
        groupKey: null,
        itemType: null // 'SF', 'WF', or 'ST' to determine formula
    }

    // Determine item type (SF, WF, or ST)
    const itemLower = itemName.toLowerCase()
    if (itemLower.startsWith('sf') || itemLower.includes('strip footing')) {
        result.itemType = 'SF'
    } else if (itemLower.startsWith('wf-')) {
        result.itemType = 'WF'
    } else if (itemLower.startsWith('st-')) {
        result.itemType = 'ST'
    }

    const dims = parseBracketDimensions(itemName)
    if (dims && dims.length >= 2) {
        result.width = dims[0]  // First value is width (goes to column G)
        result.height = dims[1]  // Second value is height (goes to column H)
        result.groupKey = `${dims[0].toFixed(2)}`
        
        // Debug logging
    }

    return result
}

/**
 * Parses isolated footing
 */
export const parseIsolatedFooting = (itemName) => {
    const result = {
        type: 'isolated_footing',
        length: 0,
        width: 0,
        height: 0
    }

    const dims = parseBracketDimensions(itemName)
    if (dims && dims.length >= 3) {
        result.length = dims[0]
        result.width = dims[1]
        result.height = dims[2]
    }

    return result
}

/**
 * Parses pilaster
 * From brackets, extracts Length, Width, Height
 * Example: Pilaster (P3) (22"x16"x6'-0") -> Length: 1.833, Width: 1.3333, Height: 6
 */
export const parsePilaster = (itemName) => {
    const result = {
        type: 'pilaster',
        length: 0,
        width: 0,
        height: 0
    }

    const dims = parseBracketDimensions(itemName)
    if (dims && dims.length >= 3) {
        // parseBracketDimensions already returns values in feet
        // For inches-only (like 22", 16"), it converts to feet (22/12 = 1.833)
        // For feet-inches (like 6'-0"), it converts to feet (6.0)
        result.length = dims[0]  // Length (F) - already in feet
        result.width = dims[1]   // Width (G) - already in feet
        result.height = dims[2]  // Height (H) - already in feet
        
    }

    return result
}

/**
 * Parses grade beam
 */
export const parseGradeBeam = (itemName) => {
    const result = {
        type: 'grade_beam',
        width: 0,
        height: 0,
        groupKey: null
    }

    const dims = parseBracketDimensions(itemName)
    if (dims && dims.length >= 2) {
        result.width = dims[0]
        result.height = dims[1]
        result.groupKey = `${dims[0].toFixed(2)}`
    }

    return result
}

/**
 * Parses tie beam
 */
export const parseTieBeam = (itemName) => {
    const result = {
        type: 'tie_beam',
        width: 0,
        height: 0,
        groupKey: null
    }

    const dims = parseBracketDimensions(itemName)
    if (dims && dims.length >= 2) {
        // parseBracketDimensions already returns feet
        // For TB1 (20"x32"): dims[0] = 20/12 = 1.666..., dims[1] = 32/12 = 2.666...
        result.width = dims[0]
        result.height = dims[1]
        result.groupKey = `${dims[0].toFixed(2)}`
    }

    return result
}

/**
 * Parses strap beam
 * Format: ST (3'-10"x2'-9") typ. or ST (2'-8"x3'-0") (87.86')
 * Same structure as grade beam: width x height from bracket
 */
export const parseStrapBeam = (itemName) => {
    const result = {
        type: 'strap_beam',
        width: 0,
        height: 0,
        groupKey: null
    }

    const dims = parseBracketDimensions(itemName)
    if (dims && dims.length >= 2) {
        result.width = dims[0]
        result.height = dims[1]
        result.groupKey = `${dims[0].toFixed(2)}`
    }

    return result
}

/**
 * Parses thickened slab
 */
export const parseThickenedSlab = (itemName) => {
    const result = {
        type: 'thickened_slab',
        width: 0,
        height: 0,
        groupKey: null
    }

    const dims = parseBracketDimensions(itemName)
    if (dims && dims.length >= 2) {
        result.width = dims[0]
        result.height = dims[1]
        result.groupKey = `${dims[0].toFixed(2)}`
    }

    return result
}

/**
 * Parses pier
 */
export const parsePier = (itemName) => {
    const result = {
        type: 'pier',
        length: 0,
        width: 0,
        height: 0,
        groupKey: null
    }

    const dims = parseBracketDimensions(itemName)
    if (dims && dims.length >= 3) {
        // parseBracketDimensions already returns feet
        result.length = dims[0]
        result.width = dims[1]
        result.height = dims[2]
        // Group by first bracket value (length)
        result.groupKey = `${dims[0].toFixed(2)}`
    }

    return result
}

/**
 * Parses corbel
 */
export const parseCorbel = (itemName) => {
    const result = {
        type: 'corbel',
        width: 0,
        height: 0,
        groupKey: null
    }

    const dims = parseBracketDimensions(itemName)
    if (dims && dims.length >= 2) {
        // parseBracketDimensions already returns values in feet
        result.width = dims[0]
        result.height = dims[1]
        // Group by first value only (width)
        result.groupKey = `${dims[0].toFixed(2)}`
    }

    return result
}

/**
 * Parses linear wall
 */
export const parseLinearWall = (itemName) => {
    const result = {
        type: 'linear_wall',
        width: 0,
        height: 0,
        groupKey: null
    }

    const dims = parseBracketDimensions(itemName)
    if (dims && dims.length >= 2) {
        // parseBracketDimensions already returns values in feet
        result.width = dims[0]
        result.height = dims[1]
        // Group by first value only (width)
        result.groupKey = `${dims[0].toFixed(2)}`
    }

    return result
}

/**
 * Parses foundation wall
 */
export const parseFoundationWall = (itemName) => {
    const result = {
        type: 'foundation_wall',
        width: 0,
        height: 0,
        groupKey: null
    }

    const dims = parseBracketDimensions(itemName)
    if (dims && dims.length >= 2) {
        result.width = dims[0]
        result.height = dims[1]
        // Group by first value only (width)
        result.groupKey = `${dims[0].toFixed(2)}`
    }

    return result
}

/**
 * Parses retaining wall
 */
export const parseRetainingWall = (itemName) => {
    const result = {
        type: 'retaining_wall',
        width: 0,
        height: 0,
        groupKey: null
    }

    const dims = parseBracketDimensions(itemName)
    if (dims && dims.length >= 2) {
        // parseBracketDimensions already returns values in feet
        result.width = dims[0]
        result.height = dims[1]
        // Group by width (first dimension) only (in feet)
        result.groupKey = `${dims[0].toFixed(2)}`
    }

    return result
}

/**
 * Parses barrier wall
 */
export const parseBarrierWall = (itemName) => {
    const result = {
        type: 'barrier_wall',
        width: 0,
        height: 0,
        groupKey: null
    }

    const dims = parseBracketDimensions(itemName)
    if (dims && dims.length >= 2) {
        // parseBracketDimensions already returns values in feet
        result.width = dims[0]
        result.height = dims[1]
        // Group by width (first dimension) only (in feet)
        result.groupKey = `${dims[0].toFixed(2)}`
    }

    return result
}

/**
 * Parses stem wall
 * Example: Stem wall (10"x3'-10") -> width: 0.833, height: 3.833
 */
export const parseStemWall = (itemName) => {
    const result = {
        type: 'stem_wall',
        width: 0,
        height: 0,
        groupKey: null
    }

    const dims = parseBracketDimensions(itemName)
    if (dims && dims.length >= 2) {
        // parseBracketDimensions already returns values in feet
        result.width = dims[0]  // Width (G) - first value from bracket
        result.height = dims[1] // Height (H) - second value from bracket
        result.groupKey = `${dims[0].toFixed(2)}`
    }

    return result
}

/**
 * Parses elevator pit items
 * Handles: Sump pit, Elev. pit slab, Elev. pit mat, Elev. pit mat slab, Elev. pit wall, Elev. pit slope transition/haunch
 * Note: "elev"/"elevator"/"elev." are treated as same; "pit" is optional (e.g., "elev slab" = "elev pit slab")
 */
export const parseElevatorPit = (itemName) => {
    const result = {
        type: 'elevator_pit',
        itemSubType: null, // 'sump_pit', 'slab', 'mat', 'mat_slab', 'wall', 'slope_transition'
        length: 0,
        width: 0,
        height: 0,
        heightFromH: null, // For slab items with H=3'-0" format
        groupKey: null
    }

    const itemLower = itemName.toLowerCase()

    // Identify sub-type (sump pit: "sump pit @ elevator", "sump pit @ elevator pit", or "sump pit")
    if (itemLower.includes('sump pit @ elevator') || itemLower.includes('sump pit @ elevator pit') || itemLower.includes('sump pit')) {
        result.itemSubType = 'sump_pit'
        return result
    } else if (itemLower.includes('mat slab')) {
        result.itemSubType = 'mat_slab'
        // Extract height from H=3'-0" format
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromH = parseDimension(hMatch[1])
        }
        return result
    } else if (itemLower.includes('mat')) {
        result.itemSubType = 'mat'
        // Extract height from H=3'-0" format or thickness like "30""
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromH = parseDimension(hMatch[1])
        } else {
            // Try to extract thickness from patterns like "mat 30""
            const thicknessMatch = itemName.match(/mat\s+(\d+)"?/i)
            if (thicknessMatch) {
                result.heightFromH = parseFloat(thicknessMatch[1]) / 12 // Convert inches to feet
            }
        }
        return result
    } else if (itemLower.includes('slab')) {
        result.itemSubType = 'slab'
        // Extract height from H=3'-0" format
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromH = parseDimension(hMatch[1])
        }
        return result
    } else if (itemLower.includes('wall')) {
        result.itemSubType = 'wall'
        const dims = parseBracketDimensions(itemName)
        if (dims && dims.length >= 2) {
            // Two parsed values: width (G) and height (H)
            result.width = dims[0]   // Width (G)
            result.height = dims[1]  // Height (H)
            result.groupKey = `${dims[0].toFixed(2)}`
        }
        return result
    } else if (itemLower.includes('slope transition') || itemLower.includes('haunch')) {
        result.itemSubType = 'slope_transition'
        const dims = parseBracketDimensions(itemName)
        if (dims && dims.length >= 2) {
            // Two parsed values: width (G) and height (H)
            result.width = dims[0]   // Width (G)
            result.height = dims[1]  // Height (H)
            result.groupKey = `${dims[0].toFixed(2)}`
        }
        return result
    }

    return result
}

/**
 * Parses service elevator pit items
 * Same structure as elevator pit: Sump pit @ service elevator, Service Elev. pit slab, Service Elev. pit mat, Service Elev. pit mat slab, Service Elev. pit wall, slope transition
 * Note: "elev"/"elevator"/"elev." are treated as same; "pit" is optional
 */
export const parseServiceElevatorPit = (itemName) => {
    const result = {
        type: 'service_elevator_pit',
        itemSubType: null, // 'sump_pit', 'slab', 'mat', 'mat_slab', 'wall', 'slope_transition'
        length: 0,
        width: 0,
        height: 0,
        heightFromH: null,
        groupKey: null
    }

    const itemLower = itemName.toLowerCase()

    if (itemLower.includes('sump pit @ service elevator') || itemLower.includes('sump pit @ service elevator pit')) {
        result.itemSubType = 'sump_pit'
        return result
    } else if (itemLower.includes('mat slab')) {
        // New group: mat slab 
        result.itemSubType = 'mat_slab'
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromH = parseDimension(hMatch[1])
        }
        return result
    } else if (itemLower.includes('mat')) {
        result.itemSubType = 'mat'
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromH = parseDimension(hMatch[1])
        } else {
            const thicknessMatch = itemName.match(/mat\s+(\d+)"?/i)
            if (thicknessMatch) {
                result.heightFromH = parseFloat(thicknessMatch[1]) / 12
            }
        }
        return result
    } else if (itemLower.includes('slab')) {
        result.itemSubType = 'slab'
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromH = parseDimension(hMatch[1])
        }
        return result
    } else if (itemLower.includes('wall')) {
        result.itemSubType = 'wall'
        const dims = parseBracketDimensions(itemName)
        if (dims && dims.length >= 2) {
            result.width = dims[0]
            result.height = dims[1]
            result.groupKey = `${dims[0].toFixed(2)}`
        }
        return result
    } else if (itemLower.includes('slope transition') || itemLower.includes('haunch')) {
        result.itemSubType = 'slope_transition'
        const dims = parseBracketDimensions(itemName)
        if (dims && dims.length >= 2) {
            result.width = dims[0]
            result.height = dims[1]
            result.groupKey = `${dims[0].toFixed(2)}`
        }
        return result
    }

    return result
}

/**
 * Parses detention tank items
 * Handles: slab items (extract height from name like "12"", "8""), wall items (parse from bracket)
 */
export const parseDetentionTank = (itemName) => {
    const result = {
        type: 'detention_tank',
        itemSubType: null, // 'slab', 'lid_slab', 'wall'
        width: 0,
        height: 0,
        length: 0,
        heightFromName: null, // For slab items extracted from name
        groupKey: null
    }

    const itemLower = itemName.toLowerCase()

    if (itemLower.includes('slab')) {
        if (itemLower.includes('lid')) {
            result.itemSubType = 'lid_slab'
        } else {
            result.itemSubType = 'slab'
        }
        // Extract height from name like "12"", "8""
        const inchMatch = itemName.match(/(\d+)"\s*(?:typ\.)?/i)
        if (inchMatch) {
            const inches = parseFloat(inchMatch[1])
            result.heightFromName = inches / 12 // Convert to feet
        }
        return result
    } else if (itemLower.includes('wall')) {
        result.itemSubType = 'wall'
        const dims = parseBracketDimensions(itemName)
        if (dims && dims.length >= 2) {
            // For Detention tank wall: F=first dim, G=first dim (width), H=second dim
            result.length = dims[0]  // Length (F)
            result.width = dims[0]  // Width (G) - same as first dimension
            result.height = dims[1] // Height (H) - second dimension
            result.groupKey = `${dims[0].toFixed(2)}`
        }
        return result
    }

    return result
}

/**
 * Parses duplex sewage ejector pit items
 * Handles: slab, mat, mat_slab, wall, slope items with flexible naming (pit is optional)
 */
export const parseDuplexSewageEjectorPit = (itemName) => {
    const result = {
        type: 'duplex_sewage_ejector_pit',
        itemSubType: null, // 'slab', 'mat', 'mat_slab', 'wall', 'slope_transition'
        width: 0,
        height: 0,
        heightFromName: null,
        groupKey: null
    }

    const itemLower = itemName.toLowerCase()

    if (itemLower.includes('mat slab')) {
        result.itemSubType = 'mat_slab'
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromName = parseDimension(hMatch[1])
        }
        return result
    } else if (itemLower.includes('mat')) {
        result.itemSubType = 'mat'
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromName = parseDimension(hMatch[1])
        } else {
            const inchMatch = itemName.match(/(\d+)"\s*(?:typ\.)?/i)
            if (inchMatch) {
                const inches = parseFloat(inchMatch[1])
                result.heightFromName = inches / 12
            }
        }
        return result
    } else if (itemLower.includes('slab')) {
        result.itemSubType = 'slab'
        const inchMatch = itemName.match(/(\d+)"\s*(?:typ\.)?/i)
        if (inchMatch) {
            const inches = parseFloat(inchMatch[1])
            result.heightFromName = inches / 12
        }
        return result
    } else if (itemLower.includes('wall')) {
        result.itemSubType = 'wall'
        const dims = parseBracketDimensions(itemName)
        if (dims && dims.length >= 2) {
            result.width = dims[0]
            result.height = dims[1]
            result.groupKey = `${dims[0].toFixed(2)}x${dims[1].toFixed(2)}`
        }
        return result
    } else if (itemLower.includes('slope transition') || itemLower.includes('haunch')) {
        result.itemSubType = 'slope_transition'
        const dims = parseBracketDimensions(itemName)
        if (dims && dims.length >= 2) {
            result.width = dims[0]
            result.height = dims[1]
            result.groupKey = `${dims[0].toFixed(2)}x${dims[1].toFixed(2)}`
        }
        return result
    }

    return result
}

/**
 * Parses deep sewage ejector pit items
 * Handles: slab, mat, mat_slab, wall, slope items with flexible naming (pit is optional)
 */
export const parseDeepSewageEjectorPit = (itemName) => {
    const result = {
        type: 'deep_sewage_ejector_pit',
        itemSubType: null, // 'slab', 'mat', 'mat_slab', 'wall', 'slope_transition'
        width: 0,
        height: 0,
        heightFromName: null,
        groupKey: null
    }

    const itemLower = itemName.toLowerCase()

    if (itemLower.includes('mat slab')) {
        result.itemSubType = 'mat_slab'
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromName = parseDimension(hMatch[1])
        }
        return result
    } else if (itemLower.includes('mat')) {
        result.itemSubType = 'mat'
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromName = parseDimension(hMatch[1])
        } else {
            const inchMatch = itemName.match(/(\d+)"\s*(?:typ\.)?/i)
            if (inchMatch) {
                const inches = parseFloat(inchMatch[1])
                result.heightFromName = inches / 12
            }
        }
        return result
    } else if (itemLower.includes('slab')) {
        result.itemSubType = 'slab'
        const inchMatch = itemName.match(/(\d+)"\s*(?:typ\.)?/i)
        if (inchMatch) {
            const inches = parseFloat(inchMatch[1])
            result.heightFromName = inches / 12
        }
        return result
    } else if (itemLower.includes('wall') || itemLower.includes('slope transition') || itemLower.includes('haunch')) {
        if (itemLower.includes('slope transition') || itemLower.includes('haunch')) {
            result.itemSubType = 'slope_transition'
        } else {
            result.itemSubType = 'wall'
        }
        const dims = parseBracketDimensions(itemName)
        if (dims && dims.length >= 2) {
            result.width = dims[0]
            result.height = dims[1]
            result.groupKey = `${dims[0].toFixed(2)}x${dims[1].toFixed(2)}`
        }
        return result
    }

    return result
}

/**
 * Parses sump pump pit items
 * Handles: slab items (extract height from name), wall items (parse from bracket)
 * Same grouping and formulas as Duplex sewage ejector pit
 */
export const parseSumpPumpPit = (itemName) => {
    const result = {
        type: 'sump_pump_pit',
        itemSubType: null, // 'slab', 'mat', 'mat_slab', 'wall', 'slope_transition'
        width: 0,
        height: 0,
        heightFromName: null,
        groupKey: null
    }

    const itemLower = itemName.toLowerCase()

    if (itemLower.includes('mat slab')) {
        result.itemSubType = 'mat_slab'
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromName = parseDimension(hMatch[1])
        }
        return result
    } else if (itemLower.includes('mat')) {
        result.itemSubType = 'mat'
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromName = parseDimension(hMatch[1])
        } else {
            const inchMatch = itemName.match(/(\d+)"\s*(?:typ\.)?/i)
            if (inchMatch) {
                const inches = parseFloat(inchMatch[1])
                result.heightFromName = inches / 12
            }
        }
        return result
    } else if (itemLower.includes('slab')) {
        result.itemSubType = 'slab'
        const inchMatch = itemName.match(/(\d+)"\s*(?:typ\.)?/i)
        if (inchMatch) {
            const inches = parseFloat(inchMatch[1])
            result.heightFromName = inches / 12 // Convert to feet
        }
        return result
    } else if (itemLower.includes('wall') || itemLower.includes('slope transition') || itemLower.includes('haunch')) {
        if (itemLower.includes('slope transition') || itemLower.includes('haunch')) {
            result.itemSubType = 'slope_transition'
        } else {
            result.itemSubType = 'wall'
        }
        const dims = parseBracketDimensions(itemName)
        if (dims && dims.length >= 2) {
            result.width = dims[0]
            result.height = dims[1]
            result.groupKey = `${dims[0].toFixed(2)}x${dims[1].toFixed(2)}`
        }
        return result
    }

    return result
}

/**
 * Parses grease trap items
 * Handles: slab, mat, mat_slab, wall, slope items with flexible naming (pit is optional)
 */
export const parseGreaseTrap = (itemName) => {
    const result = {
        type: 'grease_trap',
        itemSubType: null, // 'slab', 'mat', 'mat_slab', 'wall', 'slope_transition'
        width: 0,
        height: 0,
        heightFromName: null,
        groupKey: null
    }

    const itemLower = itemName.toLowerCase()

    if (itemLower.includes('mat slab')) {
        result.itemSubType = 'mat_slab'
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromName = parseDimension(hMatch[1])
        }
        return result
    } else if (itemLower.includes('mat')) {
        result.itemSubType = 'mat'
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromName = parseDimension(hMatch[1])
        } else {
            const inchMatch = itemName.match(/(\d+)"\s*(?:typ\.)?/i)
            if (inchMatch) {
                const inches = parseFloat(inchMatch[1])
                result.heightFromName = inches / 12
            }
        }
        return result
    } else if (itemLower.includes('slab')) {
        result.itemSubType = 'slab'
        const inchMatch = itemName.match(/(\d+)"\s*(?:typ\.)?/i)
        if (inchMatch) {
            const inches = parseFloat(inchMatch[1])
            result.heightFromName = inches / 12 // Convert to feet
        }
        return result
    } else if (itemLower.includes('wall') || itemLower.includes('slope transition') || itemLower.includes('haunch')) {
        if (itemLower.includes('slope transition') || itemLower.includes('haunch')) {
            result.itemSubType = 'slope_transition'
        } else {
            result.itemSubType = 'wall'
        }
        const dims = parseBracketDimensions(itemName)
        if (dims && dims.length >= 2) {
            result.width = dims[0]
            result.height = dims[1]
            result.groupKey = `${dims[0].toFixed(2)}x${dims[1].toFixed(2)}`
        }
        return result
    }

    return result
}

/**
 * Parses house trap items
 * Handles: slab, mat, mat_slab, wall, slope items with flexible naming (pit is optional)
 */
export const parseHouseTrap = (itemName) => {
    const result = {
        type: 'house_trap',
        itemSubType: null, // 'slab', 'mat', 'mat_slab', 'wall', 'slope_transition'
        width: 0,
        height: 0,
        heightFromName: null,
        groupKey: null
    }

    const itemLower = itemName.toLowerCase()

    if (itemLower.includes('mat slab')) {
        result.itemSubType = 'mat_slab'
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromName = parseDimension(hMatch[1])
        }
        return result
    } else if (itemLower.includes('mat')) {
        result.itemSubType = 'mat'
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromName = parseDimension(hMatch[1])
        } else {
            const inchMatch = itemName.match(/(\d+)"\s*(?:typ\.)?/i)
            if (inchMatch) {
                const inches = parseFloat(inchMatch[1])
                result.heightFromName = inches / 12
            }
        }
        return result
    } else if (itemLower.includes('slab')) {
        result.itemSubType = 'slab'
        const inchMatch = itemName.match(/(\d+)"\s*(?:typ\.)?/i)
        if (inchMatch) {
            const inches = parseFloat(inchMatch[1])
            result.heightFromName = inches / 12 // Convert to feet
        }
        return result
    } else if (itemLower.includes('wall') || itemLower.includes('slope transition') || itemLower.includes('haunch')) {
        if (itemLower.includes('slope transition') || itemLower.includes('haunch')) {
            result.itemSubType = 'slope_transition'
        } else {
            result.itemSubType = 'wall'
        }
        const dims = parseBracketDimensions(itemName)
        if (dims && dims.length >= 2) {
            result.width = dims[0]
            result.height = dims[1]
            result.groupKey = `${dims[0].toFixed(2)}x${dims[1].toFixed(2)}`
        }
        return result
    }

    return result
}

/**
 * Parses mat slab items
 * Handles: mat items (extract height from H=X'-Y" format), haunch items (parse from bracket)
 */
export const parseMatSlab = (itemName) => {
    const result = {
        type: 'mat_slab',
        itemSubType: null, // 'mat', 'haunch'
        width: 0,
        height: 0,
        heightFromH: null, // For mat items with H=X'-Y" format
        groupKey: null
    }

    const itemLower = itemName.toLowerCase()

    if (itemLower.includes('haunch')) {
        result.itemSubType = 'haunch'
        const dims = parseBracketDimensions(itemName)
        if (dims && dims.length >= 2) {
            // Width (G) and Height (H) from bracket
            result.width = dims[0]   // Width (G)
            result.height = dims[1]  // Height (H)
            result.groupKey = `${dims[0].toFixed(2)}x${dims[1].toFixed(2)}`
        }
        return result
    } else if (itemLower.includes('mat')) {
        result.itemSubType = 'mat'
        // Extract height from H=X'-Y" format
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromH = parseDimension(hMatch[1])
            result.groupKey = `H${result.heightFromH.toFixed(2)}` // Group by height
        }
        return result
    }

    return result
}

/**
 * Parses mud slab item (Foundation section)
 * This is a manual entry item with no parsing needed
 */
export const parseMudSlabFoundation = (itemName) => {
    return {
        type: 'mud_slab_foundation',
        itemSubType: 'mud_slab',
        width: 0,
        height: 0
    }
}

/**
 * Parses SOG items
 * Handles: Gravel, Gravel backfill, Geotextile filter fabric, SOG slabs, SOG step
 */
export const parseSOG = (itemName) => {    
    const result = {
        type: 'sog',
        itemSubType: null, // 'gravel', 'gravel_backfill', 'geotextile', 'sog_slab', 'sog_step'
        width: 0,
        height: 0,
        heightFromName: null, // For items with height in name like "4"", "6""
        heightFromH: null, // For items with H=X'-Y" format
        groupKey: null
    }

    const itemLower = itemName.toLowerCase()

    if (itemLower.includes('gravel backfill')) {
        result.itemSubType = 'gravel_backfill'
        // Extract height from H=X'-Y" format
        const hMatch = itemName.match(/H\s*=\s*(\d+'-\d+")/i)
        if (hMatch) {
            result.heightFromH = parseDimension(hMatch[1])
        }
        result.groupKey = 'gravel_backfill'
        return result
    } else if (itemLower.includes('gravel')) {
        result.itemSubType = 'gravel'
        result.groupKey = 'gravel'
        return result
    } else if (itemLower.includes('geotextile filter fabric')) {
        result.itemSubType = 'geotextile'
        result.groupKey = 'geotextile'
        return result
    } else if (itemLower.includes('step')) {
        result.itemSubType = 'sog_step'
        const dims = parseBracketDimensions(itemName)
        if (dims && dims.length >= 2) {
            // Width (G) and Height (H) from bracket
            result.width = dims[0]   // Width (G)
            result.height = dims[1]  // Height (H)
            result.groupKey = `${dims[0].toFixed(2)}x${dims[1].toFixed(2)}`
        }
        return result
    } else if (itemLower.includes('sog') || itemLower.includes('slab on grade') || itemLower.includes('pressure slab')) {
        result.itemSubType = 'sog_slab'
        // Extract height from name like "SOG 4"", "SOG 6"", "Patio SOG 6"", "Patch SOG 5""
        const inchMatch = itemName.match(/(\d+)"\s*(?:thick)?/i)
        if (inchMatch) {
            const inches = parseFloat(inchMatch[1])
            result.heightFromName = inches / 12 // Convert to feet
        }
        // Group by prefix and height (e.g., "SOG 4"", "SOG 6"", "Patio SOG 6"", "Patch SOG 5"", "Pressure SOG 5"")
        let prefix = 'sog'
        if (itemLower.includes('patio')) {
            prefix = 'patio_sog'
        } else if (itemLower.includes('patch')) {
            prefix = 'patch_sog'
        } else if (itemLower.includes('pressure')) {
            prefix = 'pressure_sog'
        }
        result.groupKey = `${prefix}_${result.heightFromName?.toFixed(2) || 'other'}`
        return result
    }

    return result
}

/**
 * Parses ROG (Ramp on grade) items
 * Handles: ROG 6", ROG 4", etc. - same structure as SOG slabs
 */
export const parseROG = (itemName) => {
    const result = {
        type: 'rog',
        itemSubType: 'rog_slab',
        heightFromName: null,
        groupKey: null
    }

    const itemLower = itemName.toLowerCase()
    // Extract height from name like "ROG 6"", "ROG 4""
    const inchMatch = itemName.match(/(\d+)"\s*(?:thick)?/i)
    if (inchMatch) {
        const inches = parseFloat(inchMatch[1])
        result.heightFromName = inches / 12 // Convert to feet
    }
    result.groupKey = `rog_${result.heightFromName?.toFixed(2) || 'other'}`
    return result
}

/**
 * Parses stairs on grade items
 * Handles: Stairs on grade (with/without width in name), Landings on grade
 * Raw items: "Stairs on grade @ Stair A1", "Landings on grade @ Stair A1", "Stairs on grade", "Landings on grade"
 * Special: "Stairs on grade 5'-5" wide @ stair A2" - parse width and height from name for col G and H
 * Grouping: by text after @ (e.g. "Stair A1"), or "NO_AT" if no @
 */
export const parseStairsOnGrade = (itemName) => {
    const result = {
        type: 'stairs_on_grade',
        itemSubType: null, // 'stairs', 'landings'
        width: 0,
        widthFromName: null, // Width parsed from name like "5'-5" wide"
        heightFromName: null, // Height (col H) parsed from name like "7" riser" -> 7/12 feet
        stairIdentifier: null,
        groupKey: null
    }

    const itemLower = itemName.toLowerCase()

    // Group by text after @, or "NO_AT" if no @ (same logic as Demo stair on grade)
    const atMatch = itemName.match(/@\s*(.+)$/i)
    const groupKey = atMatch ? atMatch[1].trim() : 'NO_AT'
    result.groupKey = groupKey

    if (itemLower.includes('landings')) {
        result.itemSubType = 'landings'
        return result
    } else if (itemLower.includes('stairs on grade')) {
        result.itemSubType = 'stairs'
        // Extract width from name like "Stairs on grade 5'-5" wide"
        const widthMatch = itemName.match(/(\d+'-?\d*")\s*wide/i)
        if (widthMatch) {
            result.widthFromName = parseDimension(widthMatch[1])
        }
        // Extract height (col H) from name - e.g. "7" riser", "7\" riser", "7 inch riser"
        const riserMatch = itemName.match(/(\d+(?:\.\d+)?)\s*["']?\s*riser/i)
        if (riserMatch) {
            const inches = parseFloat(riserMatch[1])
            result.heightFromName = inches / 12
        }
        return result
    }

    return result
}

export default {
    isDrilledFoundationPile,
    isHelicalFoundationPile,
    isDrivenFoundationPile,
    isStelcorDrilledDisplacementPile,
    isCFAPile,
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
    isLinearWall,
    isFoundationWall,
    isRetainingWall,
    isBarrierWall,
    isStemWall,
    isElevatorPit,
    isDetentionTank,
    isDuplexSewageEjectorPit,
    isDeepSewageEjectorPit,
    isSumpPumpPit,
    isGreaseTrap,
    isHouseTrap,
    isMatSlab,
    isMudSlabFoundation,
    isSOG,
    isROG,
    isStairsOnGrade,
    isElectricConduit,
    parseDrilledFoundationPile,
    parseHelicalFoundationPile,
    parseDrivenFoundationPile,
    parseStelcorDrilledDisplacementPile,
    parseCFAPile,
    tryMatchPileStructure,
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
    parseLinearWall,
    parseFoundationWall,
    parseRetainingWall,
    parseBarrierWall,
    parseStemWall,
    parseElevatorPit,
    parseDetentionTank,
    parseDuplexSewageEjectorPit,
    parseDeepSewageEjectorPit,
    parseSumpPumpPit,
    parseGreaseTrap,
    parseHouseTrap,
    parseMatSlab,
    parseMudSlabFoundation,
    parseSOG,
    parseROG,
    parseStairsOnGrade,
    parseElectricConduit,
    calculatePileWeight
}
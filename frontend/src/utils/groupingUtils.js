/**
 * Centralized grouping utility for extracting grouping keys from item descriptions
 * This module provides functions to extract the appropriate grouping parameter from various item types
 */

/**
 * Extracts the primary grouping key from an item description
 * @param {string} itemDescription - The item description (particulars)
 * @returns {string} - The grouping key
 */
export const extractGroupingKey = (itemDescription) => {
  if (!itemDescription || typeof itemDescription !== 'string') {
    return 'OTHER'
  }

  const itemLower = itemDescription.toLowerCase()

  // Demo SOG 4" thick / Demo ROG 4" thick - group by thickness
  if ((itemLower.includes('demo sog') || itemLower.includes('demo rog')) && itemLower.includes('"')) {
    const thickMatch = itemDescription.match(/(\d+)["']?\s*thick/i)
    if (thickMatch) {
      return `THICK_${thickMatch[1]}`
    }
  }

  // Detention tank lid slab 8" - group by thickness at end
  if (itemLower.includes('detention tank') && itemLower.includes('lid slab')) {
    const thickMatch = itemDescription.match(/(\d+)["']?\s*$/i)
    if (thickMatch) {
      return `THICK_${thickMatch[1]}`
    }
  }

  // Demo SF (2'-0"x1'-0") / Demo FW (1'-0"x3'-0") / Demo RW (1'-0"x3'-0") - group by first bracket value
  if ((itemLower.includes('demo sf') || itemLower.includes('demo fw') || itemLower.includes('demo rw')) && itemDescription.includes('(')) {
    const bracketMatch = itemDescription.match(/\(([^x)]+)/)
    if (bracketMatch) {
      return `DIM_${bracketMatch[1].trim()}`
    }
  }

  // Rock bolt @ 7'-0" O.C. (Bond length=10'-0") - group by spacing value
  if (itemLower.includes('rock bolt') && itemDescription.includes('@')) {
    const spacingMatch = itemDescription.match(/@\s*([0-9'"\-]+)\s*o\.?c/i)
    if (spacingMatch) {
      return `SPACING_${spacingMatch[1].trim()}`
    }
  }

  // Shotcrete w/ wire mesh H=15'-0" - group by H parameter
  if (itemLower.includes('shotcrete') && itemDescription.includes('H=')) {
    const hMatch = itemDescription.match(/H=([0-9'"\-]+)/i)
    if (hMatch) {
      return `H_${hMatch[1].trim()}`
    }
  }

  // Rock stabilization (H=2'-4") - group by H parameter
  if (itemLower.includes('rock stabilization') && itemDescription.includes('H=')) {
    const hMatch = itemDescription.match(/H=([0-9'"\-]+)/i)
    if (hMatch) {
      return `H_${hMatch[1].trim()}`
    }
  }

  // 1" form board (H=14'-2") - group by thickness at start
  if (itemLower.includes('form board') && itemDescription.match(/^\d+["']/)) {
    const thickMatch = itemDescription.match(/^(\d+["'])/)
    if (thickMatch) {
      return `THICK_${thickMatch[1]}`
    }
  }

  // Pier (12"x24"x3'-4"), Corbel (8"x2'-3"), Concrete liner wall (8"x11'-1"), 
  // FW (1'-0"x10'-10"), Vehicle barrier wall (8"x2'-6") - group by first bracket value
  const itemTypes = [
    'pier',
    'corbel',
    'concrete liner wall',
    'fw',
    'foundation wall',
    'vehicle barrier wall',
    'barrier wall',
    'retaining wall',
    'demo fw',
    'demo isolated footing'
  ]

  for (const type of itemTypes) {
    if (itemLower.includes(type) && itemDescription.includes('(')) {
      const bracketMatch = itemDescription.match(/\(([^x)]+)/)
      if (bracketMatch) {
        return `DIM_${bracketMatch[1].trim()}`
      }
    }
  }

  // Permission grouting H=18'-0" - extract H value for potential merging
  if (itemLower.includes('grouting') && itemDescription.includes('H=')) {
    const hMatch = itemDescription.match(/H=([0-9'"\-]+)/i)
    if (hMatch) {
      return `H_${hMatch[1].trim()}`
    }
  }

  // Generic H= parameter extraction for other items
  if (itemDescription.includes('H=')) {
    const hMatch = itemDescription.match(/H=([0-9'"\-]+)/i)
    if (hMatch) {
      return `H_${hMatch[1].trim()}`
    }
  }

  // Generic bracket extraction (first value) for other items
  if (itemDescription.includes('(') && itemDescription.includes('x')) {
    const bracketMatch = itemDescription.match(/\(([^x)]+)/)
    if (bracketMatch) {
      return `DIM_${bracketMatch[1].trim()}`
    }
  }

  return 'OTHER'
}

/**
 * Groups items by their grouping key and handles single-item group merging
 * @param {Array} items - Array of items to group
 * @param {Function} keyExtractor - Function to extract grouping key from item
 * @returns {Array} - Array of grouped items
 */
export const groupItemsByKey = (items, keyExtractor = null) => {
  if (!items || items.length === 0) return []

  // Use provided key extractor or default to parsed.groupKey or extractGroupingKey
  const getKey = keyExtractor || ((item) => {
    if (item.parsed?.groupKey) return item.parsed.groupKey
    return extractGroupingKey(item.particulars)
  })

  // Group items by their key
  const groupMap = new Map()
  items.forEach(item => {
    const groupKey = getKey(item)
    if (!groupMap.has(groupKey)) {
      groupMap.set(groupKey, {
        groupKey: groupKey,
        items: [],
        parsed: item.parsed
      })
    }
    groupMap.get(groupKey).items.push(item)
  })

  // Convert to array
  let groups = Array.from(groupMap.values())

  // Identify single-item groups for potential merging
  const singleItemGroups = groups.filter(g => g.items.length === 1)
  const multiItemGroups = groups.filter(g => g.items.length > 1)

  // If there are multiple single-item groups, merge them
  if (singleItemGroups.length > 1) {
    const mergedItems = []
    singleItemGroups.forEach(group => {
      mergedItems.push(...group.items)
    })

    // Create a merged group
    const mergedGroup = {
      groupKey: 'MERGED_SINGLES',
      items: mergedItems,
      parsed: mergedItems[0].parsed,
      isMerged: true
    }

    return [...multiItemGroups, mergedGroup]
  }

  return groups
}

/**
 * If all groups have exactly one item each, merge them into a single group.
 * Each element in the merged group remains distinct.
 * @param {Array} groups - Array of groups with .items array
 * @returns {Array} - Groups (merged into one if all were single-item)
 */
export const mergeSingleItemGroupsIfAll = (groups) => {
  if (!groups || groups.length <= 1) return groups
  const allSingleItem = groups.every(g => g.items && g.items.length === 1)
  if (!allSingleItem) return groups

  const firstGroup = groups[0]
  return [{
    ...firstGroup,
    groupKey: firstGroup.groupKey ? `${firstGroup.groupKey}_MERGED` : 'MERGED_SINGLES',
    items: groups.flatMap(g => g.items),
    parsed: firstGroup.parsed,
    isMerged: true
  }]
}

/**
 * Checks if items should be merged based on similar characteristics
 * Used for merging single-item groups with same parameters but different secondary values
 * @param {Array} groups - Array of groups
 * @returns {Array} - Array of groups with merging applied
 */
export const mergeSimilarSingleItemGroups = (groups) => {
  if (!groups || groups.length === 0) return []

  // Separate single-item and multi-item groups
  const singleItemGroups = groups.filter(g => g.items.length === 1)
  const multiItemGroups = groups.filter(g => g.items.length > 1)

  // If only one or no single-item groups, no merging needed
  if (singleItemGroups.length <= 1) {
    return groups
  }

  // Group single-item groups by their base type (without the specific dimension/height)
  const baseTypeMap = new Map()
  singleItemGroups.forEach(group => {
    const item = group.items[0]
    const particulars = item.particulars || ''
    
    // Extract base type (everything before the parameters)
    let baseType = particulars.replace(/\([^)]*\)/g, '').replace(/H=.*$/i, '').trim()
    baseType = baseType.replace(/\s+\d+["'].*$/, '').trim() // Remove trailing dimensions
    
    if (!baseTypeMap.has(baseType)) {
      baseTypeMap.set(baseType, [])
    }
    baseTypeMap.get(baseType).push(group)
  })

  // Merge groups that have the same base type
  const mergedGroups = []
  baseTypeMap.forEach((groupList, baseType) => {
    if (groupList.length > 1) {
      // Merge these groups
      const mergedItems = []
      groupList.forEach(group => {
        mergedItems.push(...group.items)
      })

      mergedGroups.push({
        groupKey: `${baseType}_MERGED`,
        items: mergedItems,
        parsed: mergedItems[0].parsed,
        isMerged: true,
        baseType: baseType
      })
    } else {
      // Keep as is
      mergedGroups.push(groupList[0])
    }
  })

  return [...multiItemGroups, ...mergedGroups]
}

export default {
  extractGroupingKey,
  groupItemsByKey,
  mergeSingleItemGroupsIfAll,
  mergeSimilarSingleItemGroups
}
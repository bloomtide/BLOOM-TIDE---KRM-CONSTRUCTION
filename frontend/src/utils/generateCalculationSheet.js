import capstoneTemplate from './templates/capstoneTemplate'
import { processDemolitionItems } from './processors/demolitionProcessor'
import { processExcavationItems, processBackfillItems, processMudSlabItems } from './processors/excavationProcessor'
import { processRockExcavationItems, processLineDrillItems } from './processors/rockExcavationProcessor'
import {
  processSoldierPileItems,
  processPrimarySecantItems,
  processSecondarySecantItems,
  processTangentPileItems,
  processSheetPileItems,
  processTimberLaggingItems,
  processTimberSheetingItems,
  processWalerItems,
  processRakerItems,
  processUpperRakerItems,
  processLowerRakerItems,
  processStandOffItems,
  processKickerItems,
  processChannelItems,
  processRollChockItems,
  processStudBeamItems,
  processInnerCornerBraceItems,
  processKneeBraceItems,
  processSupportingAngleItems,
  processPargingItems,
  processHeelBlockItems,
  processUnderpinningItems,
  processRockAnchorItems,
  processRockBoltItems,
  processAnchorItems,
  processTieBackItems,
  processConcreteSoilRetentionPierItems,
  processGuideWallItems,
  processDowelBarItems,
  processRockPinItems,
  processShotcreteItems,
  processPermissionGroutingItems,
  processButtonItems,
  processRockStabilizationItems,
  processFormBoardItems,
  generateSoeFormulas
} from './processors/soeProcessor'

/**
 * Generates the Calculations Sheet structure based on the selected template
 * @param {string} templateId - The template ID (e.g., 'capstone')
 * @param {Array} rawData - The raw Excel data (headers and rows)
 * @returns {object} - Object containing rows and formulas
 */
export const generateCalculationSheet = (templateId, rawData = null) => {
  let template = null

  // Select the appropriate template
  switch (templateId) {
    case 'capstone':
      template = capstoneTemplate
      break
    // Add other templates here as they're implemented
    default:
      template = capstoneTemplate
  }

  if (!template) {
    throw new Error(`Template ${templateId} not found`)
  }

  const rows = []
  const formulas = [] // Store formulas with row numbers

  // Add header row
  rows.push(template.columns)

  // Add empty row after header
  rows.push(Array(template.columns.length).fill(''))

  // Process raw data if provided
  let demolitionItemsBySubsection = {}
  let excavationItems = []
  let backfillItems = []
  let mudSlabItems = []
  let rockExcavationItems = []
  let lineDrillItems = []
  let soldierPileGroups = [] // Initialize here
  let primarySecantItems = [] // Initialize here
  let secondarySecantItems = [] // Initialize here
  let tangentPileItems = []
  let sheetPileItems = []
  let timberLaggingItems = []
  let timberSheetingItems = []
  let walerItems = []
  let rakerItems = []
  let upperRakerItems = []
  let lowerRakerItems = []
  let standOffItems = []
  let kickerItems = []
  let channelItems = []
  let rollChockItems = []
  let studBeamItems = []
  let innerCornerBraceItems = []
  let kneeBraceItems = []
  let supportingAngleGroups = []
  let pargingItems = []
  let heelBlockItems = []
  let underpinningItems = []
  let rockAnchorItems = []
  let rockBoltItems = []
  let anchorItems = []
  let tieBackItems = []
  let concreteSoilRetentionPierItems = []
  let guideWallItems = []
  let dowelBarItems = []
  let rockPinItems = []
  let shotcreteItems = []
  let permissionGroutingItems = []
  let buttonItems = []
  let rockStabilizationItems = []
  let formBoardItems = []
  let hasBackpacking = false
  let rockExcavationRowRefs = {} // To store row references for line drill
  if (rawData && rawData.length > 1) {
    const headers = rawData[0]
    const dataRows = rawData.slice(1)
    demolitionItemsBySubsection = processDemolitionItems(dataRows, headers)
    excavationItems = processExcavationItems(dataRows, headers)
    backfillItems = processBackfillItems(dataRows, headers)
    mudSlabItems = processMudSlabItems(dataRows, headers)
    rockExcavationItems = processRockExcavationItems(dataRows, headers)
    lineDrillItems = processLineDrillItems(dataRows, headers)
    soldierPileGroups = processSoldierPileItems(dataRows, headers)
    primarySecantItems = processPrimarySecantItems(dataRows, headers)
    secondarySecantItems = processSecondarySecantItems(dataRows, headers)
    tangentPileItems = processTangentPileItems(dataRows, headers)
    sheetPileItems = processSheetPileItems(dataRows, headers)
    timberLaggingItems = processTimberLaggingItems(dataRows, headers)
    timberSheetingItems = processTimberSheetingItems(dataRows, headers)
    walerItems = processWalerItems(dataRows, headers)
    rakerItems = processRakerItems(dataRows, headers)
    upperRakerItems = processUpperRakerItems(dataRows, headers)
    lowerRakerItems = processLowerRakerItems(dataRows, headers)
    standOffItems = processStandOffItems(dataRows, headers)
    kickerItems = processKickerItems(dataRows, headers)
    channelItems = processChannelItems(dataRows, headers)
    rollChockItems = processRollChockItems(dataRows, headers)
    studBeamItems = processStudBeamItems(dataRows, headers)
    innerCornerBraceItems = processInnerCornerBraceItems(dataRows, headers)
    kneeBraceItems = processKneeBraceItems(dataRows, headers)
    supportingAngleGroups = processSupportingAngleItems(dataRows, headers)
    pargingItems = processPargingItems(dataRows, headers)
    heelBlockItems = processHeelBlockItems(dataRows, headers)
    underpinningItems = processUnderpinningItems(dataRows, headers)
    rockAnchorItems = processRockAnchorItems(dataRows, headers)
    rockBoltItems = processRockBoltItems(dataRows, headers)
    anchorItems = processAnchorItems(dataRows, headers)
    tieBackItems = processTieBackItems(dataRows, headers)
    concreteSoilRetentionPierItems = processConcreteSoilRetentionPierItems(dataRows, headers)
    guideWallItems = processGuideWallItems(dataRows, headers)
    dowelBarItems = processDowelBarItems(dataRows, headers)
    rockPinItems = processRockPinItems(dataRows, headers)
    shotcreteItems = processShotcreteItems(dataRows, headers)
    permissionGroutingItems = processPermissionGroutingItems(dataRows, headers)
    buttonItems = processButtonItems(dataRows, headers)
    rockStabilizationItems = processRockStabilizationItems(dataRows, headers)
    formBoardItems = processFormBoardItems(dataRows, headers)

    hasBackpacking = timberLaggingItems.some(item =>
      item.particulars.toLowerCase().includes('w/backpacking')
    )
  }

  // Generate the structure
  template.structure.forEach((section) => {
    // Add section header
    const sectionRow = Array(template.columns.length).fill('')
    sectionRow[0] = section.section

    // Special handling for Excavation and Rock Excavation sections - add CY and 1.3*CY headers
    if (section.section === 'Excavation' || section.section === 'Rock Excavation') {
      sectionRow[10] = 'CY'        // Column K (LBS column becomes CY)
      sectionRow[11] = '1.3*CY'    // Column L (CY column becomes 1.3*CY)
    }

    rows.push(sectionRow)

    // Add empty row after section
    rows.push(Array(template.columns.length).fill(''))

    // Handle Demolition section with subsections
    if (section.section === 'Demolition') {
      section.subsections.forEach((subsection) => {
        const subsectionItems = demolitionItemsBySubsection[subsection.name] || []
        if (subsectionItems.length > 0 || subsection.name.includes('Extra line item')) {
          // Add subsection header (indented)
          const subsectionRow = Array(template.columns.length).fill('')
          subsectionRow[1] = subsection.name + ':'
          rows.push(subsectionRow)

          if (subsection.name === 'For demo Extra line item use this') {
            const extraItems = [
              { name: 'In SQ FT', unit: 'SQ FT', h: 1, type: 'demo_extra_sqft' },
              { name: 'In FT', unit: 'FT', g: 1, h: 1, type: 'demo_extra_ft' },
              { name: 'In EA', unit: 'EA', f: 1, g: 1, h: 1, type: 'demo_extra_ea' }
            ]

            extraItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.name
              itemRow[2] = 1 // Default Takeoff
              itemRow[3] = item.unit
              if (item.f) itemRow[5] = item.f
              if (item.g) itemRow[6] = item.g
              if (item.h) itemRow[7] = item.h
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: item.type, section: 'demolition' })
            })
          } else {
            const firstItemRow = rows.length + 1
            subsectionItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[5] = item.length || ''
              itemRow[6] = item.width || ''
              itemRow[7] = item.height || ''
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'demolition_item', parsedData: item, section: 'demolition', subsection: subsection.name })
            })

            // Add sum row
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'demolition_sum', section: 'demolition', subsection: subsection.name, firstDataRow: firstItemRow, lastDataRow: rows.length - 1 })
          }

          // Add empty row
          rows.push(Array(template.columns.length).fill(''))
        }
      })
    } else if (section.section === 'Excavation') {
      // Handle Excavation section
      section.subsections.forEach((subsection) => {
        let subsectionItems = []
        if (subsection.name === 'Excavation') {
          subsectionItems = excavationItems
        } else if (subsection.name === 'Backfill') {
          subsectionItems = backfillItems
        } else if (subsection.name === 'Mud slab') {
          subsectionItems = mudSlabItems
        }

        if (subsectionItems.length > 0 || subsection.name.includes('Extra line item')) {
          // Add subsection header (indented)
          const subsectionRow = Array(template.columns.length).fill('')
          subsectionRow[1] = subsection.name + ':'
          rows.push(subsectionRow)

          if (subsection.name === 'For soil excavation Extra line item use this') {
            const extraItems = [
              { name: 'In SQ FT', unit: 'SQ FT', h: 1, type: 'soil_exc_extra_sqft' },
              { name: 'In FT', unit: 'FT', g: 1, h: 1, type: 'soil_exc_extra_ft' },
              { name: 'In EA', unit: 'EA', f: 1, g: 1, h: 1, type: 'soil_exc_extra_ea' }
            ]

            extraItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.name
              itemRow[2] = 1 // Default Takeoff
              itemRow[3] = item.unit
              if (item.f) itemRow[5] = item.f
              if (item.g) itemRow[6] = item.g
              if (item.h) itemRow[7] = item.h
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: item.type, section: 'excavation' })
            })
          } else if (subsection.name === 'For Backfill Extra line item use this') {
            const extraItems = [
              { name: 'In SQ FT', unit: 'SQ FT', h: 1, type: 'backfill_extra_sqft' },
              { name: 'In FT', unit: 'FT', g: 1, h: 1, type: 'backfill_extra_ft' },
              { name: 'In EA', unit: 'EA', f: 1, g: 1, h: 1, type: 'backfill_extra_ea' }
            ]

            extraItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.name
              itemRow[2] = 1 // Default Takeoff
              itemRow[3] = item.unit
              if (item.f) itemRow[5] = item.f
              if (item.g) itemRow[6] = item.g
              if (item.h) itemRow[7] = item.h
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: item.type, section: 'excavation' })
            })
          } else {
            const firstItemRow = rows.length + 1
            subsectionItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[5] = item.length || ''
              itemRow[6] = item.width || ''
              itemRow[7] = item.height || ''
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'excavation_item', parsedData: item, section: 'excavation' })
            })

            // Add sum row
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            const sumRowNumber = rows.length
            formulas.push({
              row: sumRowNumber,
              itemType: 'excavation_sum',
              section: 'excavation',
              firstDataRow: firstItemRow,
              lastDataRow: sumRowNumber - 1,
              subsection: subsection.name
            })

            // Add Havg row for the 'Excavation' subsection only
            if (subsection.name === 'Excavation') {
              // Space of 1 row above Havg
              rows.push(Array(template.columns.length).fill(''))

              const havgRow = Array(template.columns.length).fill('')
              havgRow[1] = 'Havg'
              rows.push(havgRow)
              formulas.push({
                row: rows.length,
                itemType: 'excavation_havg',
                section: 'excavation',
                sumRowNumber: sumRowNumber
              })
            }
          }
        }

        // Add empty row after subsection
        rows.push(Array(template.columns.length).fill(''))
      })
    } else if (section.section === 'Rock Excavation') {
      // Handle Rock Excavation section
      section.subsections.forEach((subsection) => {
        let subsectionItems = []
        if (subsection.name === 'Excavation') {
          subsectionItems = rockExcavationItems
        } else if (subsection.name === 'Line drill') {
          subsectionItems = lineDrillItems
        }

        if (subsectionItems.length > 0 || subsection.name === 'For rock excavation Extra line item use this') {
          // Add subsection header (indented)
          const subsectionRow = Array(template.columns.length).fill('')
          subsectionRow[1] = subsection.name + ':'
          rows.push(subsectionRow)

          if (subsection.name === 'For rock excavation Extra line item use this') {
            const extraItems = [
              { name: 'In SQ FT', unit: 'SQ FT', h: 1, type: 'rock_exc_extra_sqft' },
              { name: 'In FT', unit: 'FT', g: 1, h: 1, type: 'rock_exc_extra_ft' },
              { name: 'In EA', unit: 'EA', f: 1, g: 1, h: 1, type: 'rock_exc_extra_ea' }
            ]

            extraItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.name
              itemRow[2] = 1 // Default Takeoff
              itemRow[3] = item.unit
              if (item.f) itemRow[5] = item.f
              if (item.g) itemRow[6] = item.g
              if (item.h) itemRow[7] = item.h
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: item.type, section: 'rock_excavation' })
            })
          } else if (subsection.name === 'Excavation') {
            const firstItemRow = rows.length + 1
            subsectionItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[5] = item.length || ''
              itemRow[6] = item.width || ''
              itemRow[7] = item.height || ''
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'rock_excavation_item', parsedData: item, section: 'rock_excavation' })
              if (item.id) rockExcavationRowRefs[item.id] = rows.length
            })

            // Add sum row
            const sumRowNumber = rows.length + 1
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: sumRowNumber, itemType: 'rock_excavation_sum', section: 'rock_excavation', firstDataRow: firstItemRow, lastDataRow: sumRowNumber - 1 })

            // Add Havg row for the 'Excavation' subsection
            rows.push(Array(template.columns.length).fill('')) // Space
            const havgRow = Array(template.columns.length).fill('')
            havgRow[1] = 'Havg'
            rows.push(havgRow)
            formulas.push({ row: rows.length, itemType: 'rock_excavation_havg', section: 'rock_excavation', sumRowNumber })
          } else if (subsection.name === 'Line drill') {
            // Add Lifts/Height header row once
            const subHeaderRow = Array(template.columns.length).fill('')
            subHeaderRow[4] = 'Lifts'
            subHeaderRow[7] = 'Height'
            rows.push(subHeaderRow)
            formulas.push({ row: rows.length, itemType: 'line_drill_sub_header', section: 'rock_excavation' })

            const firstItemRow = rows.length + 1

            // 1. Referenced items from Rock Excavation -> Excavation subsection
            rockExcavationItems.forEach(item => {
              if (['concrete_pier', 'sewage_pit_slab', 'sump_pit'].includes(item.itemType)) {
                const itemRow = Array(template.columns.length).fill('')
                itemRow[1] = item.particulars // Manually set label
                rows.push(itemRow)
                const refRow = rockExcavationRowRefs[item.id]
                let type = item.itemType === 'concrete_pier' ? 'line_drill_concrete_pier' :
                  item.itemType === 'sewage_pit_slab' ? 'line_drill_sewage_pit' :
                    'line_drill_sump_pit'

                formulas.push({
                  row: rows.length,
                  itemType: type,
                  section: 'rock_excavation',
                  refRow: refRow || 0
                })
              }
            })

            // 2. Standalone Line drill items from raw data
            lineDrillItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[7] = item.height || ''
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'line_drilling', parsedData: item, section: 'rock_excavation' })
            })

            // Add sum row
            const sumRowNumber = rows.length + 1
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: sumRowNumber, itemType: 'line_drill_sum', section: 'rock_excavation', firstDataRow: firstItemRow, lastDataRow: sumRowNumber - 1 })
          }
        }

        // Add empty row after subsection
        rows.push(Array(template.columns.length).fill(''))
      })
    } else if (section.section === 'SOE') {
      // Handle SOE section with grouped soldier pile items and other subsections
      let timberLaggingSumRow = null

      section.subsections.forEach((subsection) => {
        // Skip Backpacking if not needed
        if (subsection.name === 'Backpacking' && !hasBackpacking) {
          return
        }

        // Add subsection header (indented)
        const subsectionRow = Array(template.columns.length).fill('')
        subsectionRow[1] = subsection.name + ':'
        rows.push(subsectionRow)

        // Add space row after Underpinning heading
        if (subsection.name === 'Underpinning') {
          rows.push(Array(template.columns.length).fill(''))
        }

        let subsectionItems = []
        let itemType = 'soe_generic_item'

        if (subsection.name === 'Drilled soldier pile' && soldierPileGroups.length > 0) {
          // Process each group for soldier piles
          soldierPileGroups.forEach((group, groupIndex) => {
            const firstGroupRow = rows.length + 1
            group.items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[7] = item.parsed.calculatedHeight || ''
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'soldier_pile_item', parsedData: item, section: 'soe' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'soldier_pile_group_sum', section: 'soe', firstDataRow: firstGroupRow, lastDataRow: rows.length - 1 })
            if (groupIndex < soldierPileGroups.length - 1) rows.push(Array(template.columns.length).fill(''))
          })
        } else {
          // Other SOE subsections
          if (subsection.name === 'Primary secant piles') subsectionItems = primarySecantItems
          else if (subsection.name === 'Secondary secant piles') subsectionItems = secondarySecantItems
          else if (subsection.name === 'Tangent piles') subsectionItems = tangentPileItems
          else if (subsection.name === 'Sheet pile') subsectionItems = sheetPileItems
          else if (subsection.name === 'Timber lagging') subsectionItems = timberLaggingItems
          else if (subsection.name === 'Timber sheeting') subsectionItems = timberSheetingItems
          else if (subsection.name === 'Waler') subsectionItems = walerItems
          else if (subsection.name === 'Raker') subsectionItems = rakerItems
          else if (subsection.name === 'Upper Raker') subsectionItems = upperRakerItems
          else if (subsection.name === 'Lower Raker') subsectionItems = lowerRakerItems
          else if (subsection.name === 'Stand off') subsectionItems = standOffItems
          else if (subsection.name === 'Kicker') subsectionItems = kickerItems
          else if (subsection.name === 'Channel') subsectionItems = channelItems
          else if (subsection.name === 'Roll chock') subsectionItems = rollChockItems
          else if (subsection.name === 'Stud beam') subsectionItems = studBeamItems
          else if (subsection.name === 'Inner corner brace') subsectionItems = innerCornerBraceItems
          else if (subsection.name === 'Knee brace') subsectionItems = kneeBraceItems
          else if (subsection.name === 'Supporting angle') subsectionItems = [] // Handled specially below
          else if (subsection.name === 'Parging') subsectionItems = [] // Handled specially below
          else if (subsection.name === 'Heel blocks') subsectionItems = [] // Handled specially below
          else if (subsection.name === 'Underpinning') subsectionItems = [] // Handled specially below
          else if (subsection.name === 'Rock anchors') subsectionItems = [] // Handled specially below
          else if (subsection.name === 'Rock bolts') subsectionItems = [] // Handled specially below
          else if (subsection.name === 'Anchor') subsectionItems = [] // Handled specially below
          else if (subsection.name === 'Tie back') subsectionItems = [] // Handled specially below
          else if (subsection.name === 'Concrete soil retention piers') subsectionItems = [] // Handled specially below
          else if (subsection.name === 'Guide wall') subsectionItems = [] // Handled specially below
          else if (subsection.name === 'Dowel bar') subsectionItems = [] // Handled specially below
          else if (subsection.name === 'Rock pins') subsectionItems = [] // Handled specially below
          else if (subsection.name === 'Shotcrete') subsectionItems = [] // Handled specially below
          else if (subsection.name === 'Permission grouting') subsectionItems = [] // Handled specially below
          else if (subsection.name === 'Buttons') subsectionItems = [] // Handled specially below
          else if (subsection.name === 'Rock stabilization') subsectionItems = [] // Handled specially below
          else if (subsection.name === 'Form board') subsectionItems = [] // Handled specially below

          if (subsectionItems.length > 0) {
            const firstItemRow = rows.length + 1
            subsectionItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit

              // QTY for Waler and Rakers might be manual, but if parsed, we put it in Col E
              // User said "Col E should be empty it is manual input"
              if (['Waler', 'Raker', 'Upper Raker', 'Lower Raker', 'Inner corner brace'].includes(subsection.name)) {
                itemRow[4] = '' // Empty Col E
              } else {
                itemRow[4] = item.qty || ''
              }

              itemRow[7] = item.parsed.calculatedHeight || item.parsed.heightRaw || ''
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'soe_generic_item', parsedData: item, section: 'soe' })
            })

            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({
              row: rows.length,
              itemType: 'soe_generic_sum',
              section: 'soe',
              firstDataRow: firstItemRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name
            })

            if (subsection.name === 'Timber lagging') {
              timberLaggingSumRow = rows.length
            }
          }

          if (subsection.name === 'Supporting angle' && supportingAngleGroups.length > 0) {
            supportingAngleGroups.forEach((group, gIdx) => {
              const firstGroupRow = rows.length + 1
              group.items.forEach(item => {
                const itemRow = Array(template.columns.length).fill('')
                itemRow[1] = item.particulars
                itemRow[2] = item.takeoff
                itemRow[3] = item.unit
                itemRow[4] = (item.parsed.qty !== undefined && item.parsed.qty !== null) ? item.parsed.qty : 1
                itemRow[7] = item.parsed.heightRaw || ''
                rows.push(itemRow)
                formulas.push({ row: rows.length, itemType: 'supporting_angle', parsedData: item, section: 'soe' })
              })
              const sumRow = Array(template.columns.length).fill('')
              rows.push(sumRow)
              formulas.push({ row: rows.length, itemType: 'soe_generic_sum', section: 'soe', firstDataRow: firstGroupRow, lastDataRow: rows.length - 1, subsectionName: subsection.name })
              if (gIdx < supportingAngleGroups.length - 1) rows.push(Array(template.columns.length).fill(''))
            })
          } else if (subsection.name === 'Parging' && pargingItems.length > 0) {
            const firstItemRow = rows.length + 1
            pargingItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[7] = item.parsed.heightRaw || ''
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'parging', parsedData: item, section: 'soe' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'soe_generic_sum', section: 'soe', firstDataRow: firstItemRow, lastDataRow: rows.length - 1, subsectionName: subsection.name })
          } else if (subsection.name === 'Heel blocks' && heelBlockItems.length > 0) {
            const firstItemRow = rows.length + 1
            heelBlockItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[5] = item.parsed.length || ''
              itemRow[6] = item.parsed.width || ''
              itemRow[7] = item.parsed.height || ''
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'heel_block', parsedData: item, section: 'soe' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'soe_generic_sum', section: 'soe', firstDataRow: firstItemRow, lastDataRow: rows.length - 1, subsectionName: subsection.name })
          } else if (subsection.name === 'Underpinning' && underpinningItems.length > 0) {
            const firstItemRow = rows.length + 1
            underpinningItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[5] = item.parsed.length || ''
              itemRow[6] = item.parsed.width || ''
              itemRow[7] = item.parsed.heightRaw || ''
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'underpinning', parsedData: item, section: 'soe' })
            })
            const sumRowNumber = rows.length + 1
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: sumRowNumber, itemType: 'soe_generic_sum', section: 'soe', firstDataRow: firstItemRow, lastDataRow: sumRowNumber - 1, subsectionName: subsection.name })

            // Add empty row before Shims
            rows.push(Array(template.columns.length).fill(''))

            // Add Shims item below underpinning
            const shimRow = Array(template.columns.length).fill('')
            shimRow[1] = 'Shims'
            shimRow[6] = 4 // Width is constant
            rows.push(shimRow)
            formulas.push({ row: rows.length, itemType: 'shims', section: 'soe', underpinningSumRow: sumRowNumber })
          } else if (subsection.name === 'Rock anchors' && rockAnchorItems.length > 0) {
            const firstItemRow = rows.length + 1
            rockAnchorItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[5] = item.parsed.calculatedHeight || '' // Length (F) = calculated height
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'rock_anchor', parsedData: item, section: 'soe' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'soe_generic_sum', section: 'soe', firstDataRow: firstItemRow, lastDataRow: rows.length - 1, subsectionName: subsection.name })
          } else if (subsection.name === 'Rock bolts' && rockBoltItems.length > 0) {
            const firstItemRow = rows.length + 1
            rockBoltItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[5] = item.parsed.calculatedLength || '' // Length (F) = bond length + 5
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'rock_bolt', parsedData: item, section: 'soe' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'soe_generic_sum', section: 'soe', firstDataRow: firstItemRow, lastDataRow: rows.length - 1, subsectionName: subsection.name })
          } else if (subsection.name === 'Anchor' && anchorItems.length > 0) {
            const firstItemRow = rows.length + 1
            anchorItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[7] = item.parsed.calculatedHeight || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'anchor', parsedData: item, section: 'soe' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'soe_generic_sum', section: 'soe', firstDataRow: firstItemRow, lastDataRow: rows.length - 1, subsectionName: subsection.name })
          } else if (subsection.name === 'Tie back' && tieBackItems.length > 0) {
            const firstItemRow = rows.length + 1
            tieBackItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[7] = item.parsed.calculatedHeight || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'tie_back', parsedData: item, section: 'soe' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'soe_generic_sum', section: 'soe', firstDataRow: firstItemRow, lastDataRow: rows.length - 1, subsectionName: subsection.name })
          } else if (subsection.name === 'Concrete soil retention piers' && concreteSoilRetentionPierItems.length > 0) {
            const firstItemRow = rows.length + 1
            concreteSoilRetentionPierItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[5] = item.parsed.length || ''
              itemRow[6] = item.parsed.width || ''
              itemRow[7] = item.parsed.height || ''
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'concrete_soil_retention_pier', parsedData: item, section: 'soe' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'soe_generic_sum', section: 'soe', firstDataRow: firstItemRow, lastDataRow: rows.length - 1, subsectionName: subsection.name })
          } else if (subsection.name === 'Guide wall' && guideWallItems.length > 0) {
            const firstItemRow = rows.length + 1
            guideWallItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[6] = item.parsed.width || '' // Width (G) - calculated from bracket
              itemRow[7] = item.parsed.heightRaw || '' // Height (H) - from bracket
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'guide_wall', parsedData: item, section: 'soe' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'soe_generic_sum', section: 'soe', firstDataRow: firstItemRow, lastDataRow: rows.length - 1, subsectionName: subsection.name })
          } else if (subsection.name === 'Dowel bar' && dowelBarItems.length > 0) {
            const firstItemRow = rows.length + 1
            dowelBarItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[4] = item.parsed.qty || ''
              itemRow[7] = item.parsed.heightRaw || '' // Height (H) = H + RS
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'dowel_bar', parsedData: item, section: 'soe' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'soe_generic_sum', section: 'soe', firstDataRow: firstItemRow, lastDataRow: rows.length - 1, subsectionName: subsection.name })
          } else if (subsection.name === 'Rock pins' && rockPinItems.length > 0) {
            const firstItemRow = rows.length + 1
            rockPinItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[4] = item.parsed.qty || 1
              itemRow[7] = item.parsed.heightRaw || '' // Height (H) = H + RS
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'rock_pin', parsedData: item, section: 'soe' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'soe_generic_sum', section: 'soe', firstDataRow: firstItemRow, lastDataRow: rows.length - 1, subsectionName: subsection.name })
          } else if (subsection.name === 'Shotcrete' && shotcreteItems.length > 0) {
            const firstItemRow = rows.length + 1
            shotcreteItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              // Length (F) and Width (G) should be empty
              itemRow[7] = item.parsed.heightRaw || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'shotcrete', parsedData: item, section: 'soe' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'soe_generic_sum', section: 'soe', firstDataRow: firstItemRow, lastDataRow: rows.length - 1, subsectionName: subsection.name })
          } else if (subsection.name === 'Permission grouting' && permissionGroutingItems.length > 0) {
            const firstItemRow = rows.length + 1
            permissionGroutingItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[7] = item.parsed.heightRaw || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'permission_grouting', parsedData: item, section: 'soe' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'soe_generic_sum', section: 'soe', firstDataRow: firstItemRow, lastDataRow: rows.length - 1, subsectionName: subsection.name })
          } else if (subsection.name === 'Buttons' && buttonItems.length > 0) {
            const firstItemRow = rows.length + 1
            buttonItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[5] = item.parsed.length || ''
              itemRow[6] = item.parsed.width || ''
              itemRow[7] = item.parsed.height || ''
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'button', parsedData: item, section: 'soe' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'soe_generic_sum', section: 'soe', firstDataRow: firstItemRow, lastDataRow: rows.length - 1, subsectionName: subsection.name })
          } else if (subsection.name === 'Rock stabilization' && rockStabilizationItems.length > 0) {
            const firstItemRow = rows.length + 1
            rockStabilizationItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[7] = item.parsed.heightRaw || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'rock_stabilization', parsedData: item, section: 'soe' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'soe_generic_sum', section: 'soe', firstDataRow: firstItemRow, lastDataRow: rows.length - 1, subsectionName: subsection.name })
          } else if (subsection.name === 'Form board' && formBoardItems.length > 0) {
            const firstItemRow = rows.length + 1
            formBoardItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[7] = item.parsed.heightRaw || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'form_board', parsedData: item, section: 'soe' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'soe_generic_sum', section: 'soe', firstDataRow: firstItemRow, lastDataRow: rows.length - 1, subsectionName: subsection.name })
          } else if (subsection.name === 'Backpacking' && hasBackpacking) {
            // Add Backpacking item
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = 'Backpacking'
            itemRow[3] = 'SQ FT'
            rows.push(itemRow)

            formulas.push({
              row: rows.length,
              itemType: 'backpacking_item',
              section: 'soe',
              timberLaggingSumRow: timberLaggingSumRow
            })
          }
        }

        // Add empty row after subsection
        rows.push(Array(template.columns.length).fill(''))
      })
    } else if (section.subsections && section.subsections.length > 0) {
      // Handle other sections with subsections
      section.subsections.forEach((subsection) => {
        // Add subsection header (indented)
        const subsectionRow = Array(template.columns.length).fill('')
        subsectionRow[1] = subsection.name + ':'
        rows.push(subsectionRow)

        // If subsection has sub-subsections
        if (subsection.subSubsections && subsection.subSubsections.length > 0) {
          subsection.subSubsections.forEach((subSubsection) => {
            // Add sub-subsection header (double indented)
            const subSubsectionRow = Array(template.columns.length).fill('')
            subSubsectionRow[1] = '  ' + subSubsection.name + ':'
            rows.push(subSubsectionRow)
            rows.push(Array(template.columns.length).fill(''))
          })
        } else {
          rows.push(Array(template.columns.length).fill(''))
        }
      })
    } else {
      // For sections without defined subsections in template
      rows.push(Array(template.columns.length).fill(''))
    }
  })

  return { rows, formulas }
}

/**
 * Generates configuration for spreadsheet columns
 * @returns {Array} - Array of column configurations
 */
export const generateColumnConfigs = () => {
  return [
    { width: 110 },  // Estimate
    { width: 500 },  // Particulars
    { width: 80 },   // Takeoff
    { width: 60 },   // Unit
    { width: 60 },   // QTY
    { width: 80 },   // Length
    { width: 80 },   // Width
    { width: 80 },   // Height
    { width: 80 },   // FT
    { width: 80 },   // SQ FT
    { width: 80 },   // LBS
    { width: 80 },   // CY
    { width: 60 }    // QTY
  ]
}

/**
 * Applies formatting to the calculation sheet
 * @param {Object} spreadsheet - Syncfusion Spreadsheet component reference
 * @param {number} sheetIndex - The sheet index to format
 */
export const applyCalculationSheetFormatting = (spreadsheet, sheetIndex = 0) => {
  if (!spreadsheet) return

  try {
    // Format header row (row 0)
    spreadsheet.cellFormat(
      {
        fontWeight: 'bold',
        textAlign: 'center',
        backgroundColor: '#4472C4',
        color: '#FFFFFF'
      },
      `A1:M1`
    )

    // TODO: Add more formatting
    // - Bold section headers
    // - Indent subsections
    // - Number formatting for numeric columns
    // - Border styling
  } catch (error) {
    console.error('Error applying formatting:', error)
  }
}

export default {
  generateCalculationSheet,
  generateColumnConfigs,
  applyCalculationSheetFormatting
}

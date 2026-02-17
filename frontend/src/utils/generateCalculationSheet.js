import capstoneTemplate from './templates/capstoneTemplate'
import { processDemolitionItems } from './processors/demolitionProcessor'
import { processExcavationItems, processBackfillItems, processMudSlabItems } from './processors/excavationProcessor'
import { processRockExcavationItems, processLineDrillItems, calculateRockExcavationTotals, calculateLineDrillTotalFT } from './processors/rockExcavationProcessor'
import {
  processSoldierPileItems,
  processPrimarySecantItems,
  processSecondarySecantItems,
  processTangentPileItems,
  processSheetPileItems,
  processTimberLaggingItems,
  processTimberSheetingItems,
  processTimberSoldierPileItems,
  processTimberPlankItems,
  processTimberWalerItems,
  processTimberRakerItems,
  processTimberBraceItems,
  processTimberPostItems,
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
import {
  processDrilledFoundationPileItems,
  processHelicalFoundationPileItems,
  processDrivenFoundationPileItems,
  processStelcorDrilledDisplacementPileItems,
  processCFAPileItems,
  processMiscellaneousPileItems,
  processPileCapItems,
  processStripFootingItems,
  processIsolatedFootingItems,
  processPilasterItems,
  processGradeBeamItems,
  processTieBeamItems,
  processThickenedSlabItems,
  processButtressItems,
  processPierItems,
  processCorbelItems,
  processLinearWallItems,
  processFoundationWallItems,
  processRetainingWallItems,
  processBarrierWallItems,
  processStemWallItems,
  processElevatorPitItems,
  processDetentionTankItems,
  processDuplexSewageEjectorPitItems,
  processDeepSewageEjectorPitItems,
  processGreaseTrapItems,
  processHouseTrapItems,
  processMatSlabItems,
  processMudSlabFoundationItems,
  processSOGItems,
  processStairsOnGradeItems,
  processElectricConduitItems,
  generateFoundationFormulas
} from './processors/foundationProcessor'
import { processExteriorSideItems, processExteriorSidePitItems, processNegativeSideWallItems, processNegativeSideSlabItems } from './processors/waterproofingProcessor'
import { processSuperstructureItems } from './processors/superstructureProcessor'
import { processBPPAlternateItems } from './processors/bppAlternateProcessor'
import { processCivilDemoItems, processCivilOtherItems } from './processors/civilSiteworkProcessor'
import UsedRowTracker from './usedRowTracker'
import { mergeSingleItemGroupsIfAll } from './groupingUtils'

/** Convert Map<key, items[]> to groups, merging into one if all groups have a single item */
const getMergedGroupsIfNeeded = (groupMap) => {
  if (!groupMap || groupMap.size === 0) return []
  const groups = Array.from(groupMap.entries()).map(([key, items]) => ({ groupKey: key, items }))
  return mergeSingleItemGroupsIfAll(groups)
}

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
  let timberSoldierPileGroups = []
  let timberPlankGroups = []
  let timberWalerGroups = []
  let timberRakerGroups = []
  let timberBraceGroups = []
  let timberPostGroups = []
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
  let drilledFoundationPileGroups = []
  let helicalFoundationPileGroups = []
  let drivenFoundationPileItems = []
  let stelcorDrilledDisplacementPileItems = []
  let cfaPileItems = []
  let miscellaneousPileItems = []
  let pileCapItems = []
  let stripFootingGroups = []
  let isolatedFootingItems = []
  let pilasterItems = []
  let gradeBeamGroups = []
  let tieBeamGroups = []
  let thickenedSlabGroups = []
  let buttressItem = null
  let pierItems = []
  let corbelGroups = []
  let linearWallGroups = []
  let foundationWallGroups = []
  let retainingWallGroups = []
  let barrierWallGroups = []
  let stemWallItems = []
  let elevatorPitItems = []
  let detentionTankItems = []
  let duplexSewageEjectorPitItems = []
  let deepSewageEjectorPitItems = []
  let greaseTrapItems = []
  let houseTrapItems = []
  let matSlabItems = []
  let mudSlabFoundationItems = []
  let sogItems = []
  let stairsOnGradeGroups = []
  let electricConduitItems = []
  let exteriorSideItems = []
  let exteriorSidePitItems = []
  let negativeSideWallItems = []
  let negativeSideSlabItems = []
  let trenchingTakeoff = ''
  let superstructureItems = { cipSlab8: [], cipRoofSlab8: [], balconySlab: [], terraceSlab: [], patchSlab: [], slabSteps: [], lwConcreteFill: [], slabOnMetalDeck: [], toppingSlab: [], thermalBreak: [], raisedSlab: { kneeWall: [], raisedSlab: [] }, builtUpSlab: { kneeWall: [], builtUpSlab: [] }, builtUpStair: { kneeWall: [], builtUpStairs: [] }, builtupRamps: { kneeWall: [], ramp: [] }, concreteHanger: [], shearWalls: [], parapetWalls: [], columnsTakeoff: [], concretePost: [], concreteEncasement: [], dropPanelBracket: [], dropPanelH: [], beams: [], curbs: [], concretePad: [], nonShrinkGrout: [], repairScope: [] }
  let bppAlternateItemsByStreet = {}
  let civilDemoItems = {
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
  let civilOtherItems = {
    'Excavation': { 'transformer_pad': [], 'reinforced_sidewalk': [], 'bollard': [] },
    'Gravel': { 'transformer_pad': [], 'reinforced_sidewalk': [], 'asphalt': [] },
    'Concrete Pavement': [],
    'Asphalt': [],
    'Pads': [],
    'Soil Erosion': { 'stabilized_entrance': [], 'silt_fence': [], 'inlet_filter': [] },
    'Fence': { 'construction_fence': [], 'proposed_fence': [], 'guiderail': [] },
    'Concrete filled steel pipe bollard': []
  }
  const foundationSlabRows = {} // Populated when building Foundation section; used by Waterproofing Exterior side pit items
  let rockExcavationTotals = { totalSQFT: 0, totalCY: 0 } // Initialize rock excavation totals
  let lineDrillTotalFT = 0 // Initialize line drill total FT
  let unusedRawDataRows = [] // Store unused raw data rows for tracking
  if (rawData && rawData.length > 1) {
    const headers = rawData[0]
    const dataRows = rawData.slice(1)

    // Create tracker to track used row indices across all processors
    const tracker = new UsedRowTracker()

    demolitionItemsBySubsection = processDemolitionItems(dataRows, headers, tracker)
    excavationItems = processExcavationItems(dataRows, headers, tracker)
    backfillItems = processBackfillItems(dataRows, headers, tracker)
    mudSlabItems = processMudSlabItems(dataRows, headers, tracker)
    rockExcavationItems = processRockExcavationItems(dataRows, headers, tracker)
    // Calculate rock excavation totals
    if (rockExcavationItems.length > 0) {
      rockExcavationTotals = calculateRockExcavationTotals(rockExcavationItems)
    }
    lineDrillItems = processLineDrillItems(dataRows, headers, tracker)
    // Calculate line drill total FT
    if (lineDrillItems.length > 0) {
      lineDrillTotalFT = calculateLineDrillTotalFT(lineDrillItems)
    }
    soldierPileGroups = processSoldierPileItems(dataRows, headers, tracker)
    primarySecantItems = processPrimarySecantItems(dataRows, headers, tracker)
    secondarySecantItems = processSecondarySecantItems(dataRows, headers, tracker)
    tangentPileItems = processTangentPileItems(dataRows, headers, tracker)
    sheetPileItems = processSheetPileItems(dataRows, headers, tracker)
    timberLaggingItems = processTimberLaggingItems(dataRows, headers, tracker)
    timberSheetingItems = processTimberSheetingItems(dataRows, headers, tracker)
    timberSoldierPileGroups = processTimberSoldierPileItems(dataRows, headers, tracker)
    timberPlankGroups = processTimberPlankItems(dataRows, headers, tracker)
    timberWalerGroups = processTimberWalerItems(dataRows, headers, tracker)
    timberRakerGroups = processTimberRakerItems(dataRows, headers, tracker)
    timberBraceGroups = processTimberBraceItems(dataRows, headers, tracker)
    timberPostGroups = processTimberPostItems(dataRows, headers, tracker)
    walerItems = processWalerItems(dataRows, headers, tracker)
    rakerItems = processRakerItems(dataRows, headers, tracker)
    upperRakerItems = processUpperRakerItems(dataRows, headers, tracker)
    lowerRakerItems = processLowerRakerItems(dataRows, headers, tracker)
    standOffItems = processStandOffItems(dataRows, headers, tracker)
    kickerItems = processKickerItems(dataRows, headers, tracker)
    channelItems = processChannelItems(dataRows, headers, tracker)
    rollChockItems = processRollChockItems(dataRows, headers, tracker)
    studBeamItems = processStudBeamItems(dataRows, headers, tracker)
    innerCornerBraceItems = processInnerCornerBraceItems(dataRows, headers, tracker)
    kneeBraceItems = processKneeBraceItems(dataRows, headers, tracker)
    supportingAngleGroups = processSupportingAngleItems(dataRows, headers, tracker)
    pargingItems = processPargingItems(dataRows, headers, tracker)
    heelBlockItems = processHeelBlockItems(dataRows, headers, tracker)
    underpinningItems = processUnderpinningItems(dataRows, headers, tracker)
    rockAnchorItems = processRockAnchorItems(dataRows, headers, tracker)
    rockBoltItems = processRockBoltItems(dataRows, headers, tracker)
    anchorItems = processAnchorItems(dataRows, headers, tracker)
    tieBackItems = processTieBackItems(dataRows, headers, tracker)
    concreteSoilRetentionPierItems = processConcreteSoilRetentionPierItems(dataRows, headers, tracker)
    guideWallItems = processGuideWallItems(dataRows, headers, tracker)
    dowelBarItems = processDowelBarItems(dataRows, headers, tracker)
    rockPinItems = processRockPinItems(dataRows, headers, tracker)
    shotcreteItems = processShotcreteItems(dataRows, headers, tracker)
    permissionGroutingItems = processPermissionGroutingItems(dataRows, headers, tracker)
    buttonItems = processButtonItems(dataRows, headers, tracker)
    rockStabilizationItems = processRockStabilizationItems(dataRows, headers, tracker)
    formBoardItems = processFormBoardItems(dataRows, headers, tracker)

    hasBackpacking = timberLaggingItems.some(item =>
      item.particulars.toLowerCase().includes('w/backpacking')
    )

    // Process Foundation items
    drilledFoundationPileGroups = processDrilledFoundationPileItems(dataRows, headers, tracker)
    helicalFoundationPileGroups = processHelicalFoundationPileItems(dataRows, headers, tracker)
    drivenFoundationPileItems = processDrivenFoundationPileItems(dataRows, headers, tracker)
    stelcorDrilledDisplacementPileItems = processStelcorDrilledDisplacementPileItems(dataRows, headers, tracker)
    cfaPileItems = processCFAPileItems(dataRows, headers, tracker)
    miscellaneousPileItems = processMiscellaneousPileItems(dataRows, headers, tracker)
    pileCapItems = processPileCapItems(dataRows, headers, tracker)
    stripFootingGroups = processStripFootingItems(dataRows, headers, tracker)
    isolatedFootingItems = processIsolatedFootingItems(dataRows, headers, tracker)
    pilasterItems = processPilasterItems(dataRows, headers, tracker)
    gradeBeamGroups = processGradeBeamItems(dataRows, headers, tracker)
    tieBeamGroups = processTieBeamItems(dataRows, headers, tracker)
    thickenedSlabGroups = processThickenedSlabItems(dataRows, headers, tracker)
    buttressItem = processButtressItems(dataRows, headers, tracker)
    pierItems = processPierItems(dataRows, headers, tracker)
    corbelGroups = processCorbelItems(dataRows, headers, tracker)
    linearWallGroups = processLinearWallItems(dataRows, headers, tracker)
    foundationWallGroups = processFoundationWallItems(dataRows, headers, tracker)
    retainingWallGroups = processRetainingWallItems(dataRows, headers, tracker)
    barrierWallGroups = processBarrierWallItems(dataRows, headers, tracker)
    stemWallItems = processStemWallItems(dataRows, headers, tracker)
    elevatorPitItems = processElevatorPitItems(dataRows, headers, tracker)
    detentionTankItems = processDetentionTankItems(dataRows, headers, tracker)
    duplexSewageEjectorPitItems = processDuplexSewageEjectorPitItems(dataRows, headers, tracker)
    deepSewageEjectorPitItems = processDeepSewageEjectorPitItems(dataRows, headers, tracker)
    greaseTrapItems = processGreaseTrapItems(dataRows, headers, tracker)
    houseTrapItems = processHouseTrapItems(dataRows, headers, tracker)
    matSlabItems = processMatSlabItems(dataRows, headers, tracker)
    mudSlabFoundationItems = processMudSlabFoundationItems(dataRows, headers, tracker)
    sogItems = processSOGItems(dataRows, headers, tracker)
    stairsOnGradeGroups = processStairsOnGradeItems(dataRows, headers, tracker)
    electricConduitItems = processElectricConduitItems(dataRows, headers, tracker)
    exteriorSideItems = processExteriorSideItems(dataRows, headers, tracker)
    exteriorSidePitItems = processExteriorSidePitItems(dataRows, headers, tracker)
    negativeSideWallItems = processNegativeSideWallItems(dataRows, headers, tracker)
    negativeSideSlabItems = processNegativeSideSlabItems(dataRows, headers, tracker)
    superstructureItems = processSuperstructureItems(dataRows, headers, tracker)
    bppAlternateItemsByStreet = processBPPAlternateItems(dataRows, headers, tracker)
    civilDemoItems = processCivilDemoItems(dataRows, headers, tracker)
    civilOtherItems = processCivilOtherItems(dataRows, headers, tracker)

    // Get unused rows
    unusedRawDataRows = tracker.getUnusedRows(dataRows, headers)
    const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
    const totalIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'total')
    if (digitizerIdx >= 0 && totalIdx >= 0) {
      for (const row of dataRows) {
        const particulars = row[digitizerIdx]
        if (particulars && String(particulars).trim().toLowerCase() === 'trenching') {
          const total = row[totalIdx]
          trenchingTakeoff = total !== '' && total != null && total !== undefined ? parseFloat(total) : ''
          break
        }
      }
    }
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
    if (section.section === 'Trenching') {
      sectionRow[2] = '' // formula =L{patchbackRow} applied in Spreadsheet
      sectionRow[3] = 'CY'
    }

    // Determine if section has data
    let hasSectionData = false

    if (section.section === 'Demolition') {
      // Check if any subsection in Demolition has items
      hasSectionData = section.subsections.some(sub => (demolitionItemsBySubsection[sub.name] || []).length > 0 || sub.name.includes('Extra line item'))
    } else if (section.section === 'Excavation') {
      hasSectionData = excavationItems.length > 0 || backfillItems.length > 0 || mudSlabItems.length > 0 || section.subsections.some(sub => sub.name.includes('Extra line item'))
    } else if (section.section === 'Rock Excavation') {
      hasSectionData = rockExcavationItems.length > 0 || lineDrillItems.length > 0 || section.subsections.some(sub => sub.name.includes('Extra line item'))
    } else if (section.section === 'SOE') {
      hasSectionData = [
        soldierPileGroups, timberSoldierPileGroups, timberPlankGroups, timberWalerGroups, timberRakerGroups, timberBraceGroups, timberPostGroups, primarySecantItems, secondarySecantItems, tangentPileItems, sheetPileItems,
        timberLaggingItems, timberSheetingItems, walerItems, rakerItems, upperRakerItems, lowerRakerItems,
        standOffItems, kickerItems, channelItems, rollChockItems, studBeamItems, innerCornerBraceItems,
        kneeBraceItems, supportingAngleGroups, pargingItems, heelBlockItems, underpinningItems,
        rockAnchorItems, rockBoltItems, anchorItems, tieBackItems, concreteSoilRetentionPierItems,
        guideWallItems, dowelBarItems, rockPinItems, shotcreteItems, permissionGroutingItems,
        buttonItems, rockStabilizationItems, formBoardItems
      ].some(arr => arr && arr.length > 0) || hasBackpacking
    } else if (section.section === 'Foundation') {
      hasSectionData = [
        miscellaneousPileItems, drilledFoundationPileGroups, helicalFoundationPileGroups, drivenFoundationPileItems,
        stelcorDrilledDisplacementPileItems, cfaPileItems, pileCapItems, stripFootingGroups,
        isolatedFootingItems, pilasterItems, gradeBeamGroups, tieBeamGroups, thickenedSlabGroups,
        pierItems, corbelGroups, linearWallGroups, foundationWallGroups, retainingWallGroups,
        barrierWallGroups, stemWallItems, elevatorPitItems, detentionTankItems, duplexSewageEjectorPitItems,
        deepSewageEjectorPitItems, greaseTrapItems, houseTrapItems, matSlabItems, mudSlabFoundationItems,
        sogItems, stairsOnGradeGroups, electricConduitItems
      ].some(arr => arr && arr.length > 0) || !!buttressItem
    } else if (section.section === 'Waterproofing') {
      hasSectionData = [
        exteriorSideItems, exteriorSidePitItems, negativeSideWallItems, negativeSideSlabItems
      ].some(arr => arr && arr.length > 0) || section.subsections.some(sub => sub.name.includes('Extra line item'))
    } else if (section.section === 'Superstructure') {
      hasSectionData = Object.values(superstructureItems).some(val => Array.isArray(val) ? val.length > 0 : (typeof val === 'object' && val !== null ? Object.values(val).some(v => Array.isArray(v) && v.length > 0) : false))
    } else if (section.section === 'B.P.P. Alternate #2 scope') {
      hasSectionData = Object.keys(bppAlternateItemsByStreet).length > 0
    } else if (section.section === 'Civil / Sitework') {
      const hasDemo = Object.values(civilDemoItems).some(val => Array.isArray(val) ? val.length > 0 : Object.values(val).some(v => Array.isArray(v) && v.length > 0))
      const hasOther = Object.values(civilOtherItems).some(val => Array.isArray(val) ? val.length > 0 : Object.values(val).some(v => Array.isArray(v) && v.length > 0))
      hasSectionData = hasDemo || hasOther
    } else if (section.section === 'Trenching') {
      hasSectionData = trenchingTakeoff !== '' && trenchingTakeoff !== null && trenchingTakeoff !== undefined
    } else {
      // Default to true for unknown sections to avoid hiding things unnecessarily
      hasSectionData = true
    }

    if (hasSectionData) {
      rows.push(sectionRow)
    }

    // Add empty row after section (or Trenching items)
    if (section.section === 'Trenching' && hasSectionData) {
      const trenchingHeaderRow = rows.length
      rows.push(Array(template.columns.length).fill(''))
      const demoRow = trenchingHeaderRow + 2
      const trenchingItems = [
        { name: 'Demo', takeoff: trenchingTakeoff, g: 2.5, hFormula: '4/12', unit: 'FT', isFirst: true },
        { name: 'Excavation', takeoffRefRow: demoRow, g: 2.5, h: 2.5, unit: 'FT' },
        { name: 'Backfill', takeoffRefRow: demoRow, g: 2.5, h: 1.67, unit: 'FT' },
        { name: 'Gravel', takeoffRefRow: demoRow, g: 2.5, h: 0.5, unit: 'FT', lYellow: true },
        { name: 'Patchback', takeoffRefRow: demoRow, g: 2.5, h: 0.33, unit: 'FT' }
      ]
      trenchingItems.forEach((item) => {
        const itemRow = Array(template.columns.length).fill('')
        itemRow[1] = item.name
        if (item.isFirst && item.takeoff !== undefined && item.takeoff !== '') itemRow[2] = item.takeoff
        itemRow[3] = item.unit || 'FT'
        itemRow[6] = item.g
        if (item.h !== undefined) itemRow[7] = item.h
        if (item.hFormula) itemRow[7] = '' // formula applied in Spreadsheet
        rows.push(itemRow)
        formulas.push({ row: rows.length, itemType: 'trenching_item', section: 'trenching', ...item })
      })
      formulas.push({ row: trenchingHeaderRow, itemType: 'trenching_section_header', section: 'trenching', patchbackRow: rows.length })
      // One empty row after Trenching is added by the final else (sections without subsections)
    } else if (hasSectionData) {
      rows.push(Array(template.columns.length).fill(''))
    }

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
            // Customized demo templates for extra line items
            const extraItems = [
              {
                name: 'Demo SOG 4\" thick',
                unit: 'SQ FT',
                h: 1,
                type: 'demo_extra_sqft'
              },
              {
                name: 'Demo ROG 4\" thick',
                unit: 'SQ FT',
                h: 1,
                type: 'demo_extra_rog_sqft'
              },
              {
                name: 'Demo SF (2\'-0\"x1\'-0\")',
                unit: 'SQ FT',
                g: 1,
                h: 1,
                type: 'demo_extra_ft'
              },
              {
                name: 'Demo FW (1\'-0\"x3\'-0\")',
                unit: 'SQ FT',
                g: 1,
                h: 1,
                type: 'demo_extra_ft'
              },
              {
                name: 'Demo RW (1\'-0\"x3\'-0\")',
                unit: 'SQ FT',
                g: 1,
                h: 1,
                type: 'demo_extra_rw'
              },
              {
                name: 'Demo isolated footing (2\'-0\"x3\'-0\"x1\'-6\")',
                unit: 'EA',
                f: 1,
                g: 1,
                h: 1,
                type: 'demo_extra_ea'
              }
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
        if (subsectionItems.length > 0 || subsection.name.includes('Extra line item')) {
          rows.push(Array(template.columns.length).fill(''))
        }
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
        if (subsectionItems.length > 0 || subsection.name === 'For rock excavation Extra line item use this') {
          rows.push(Array(template.columns.length).fill(''))
        }
      })
    } else if (section.section === 'SOE') {
      // Handle SOE section with grouped soldier pile items and other subsections
      let timberLaggingSumRow = null

      section.subsections.forEach((subsection) => {
        // Skip Backpacking if not needed
        if (subsection.name === 'Backpacking' && !hasBackpacking) {
          return
        }

        // Determine if subsection has items
        let headerCheckItems = []
        if (subsection.name === 'Drilled soldier pile') headerCheckItems = soldierPileGroups // Group check
        else if (subsection.name === 'Timber soldier piles') headerCheckItems = timberSoldierPileGroups // Group check
        else if (subsection.name === 'Timber planks') headerCheckItems = timberPlankGroups // Group check
        else if (subsection.name === 'Timber waler') headerCheckItems = timberWalerGroups // Group check
        else if (subsection.name === 'Timber raker') headerCheckItems = timberRakerGroups // Group check
        else if (subsection.name === 'Timber brace') headerCheckItems = timberBraceGroups // Group check
        else if (subsection.name === 'Timber post') headerCheckItems = timberPostGroups // Group check
        else if (subsection.name === 'Primary secant piles') headerCheckItems = primarySecantItems
        else if (subsection.name === 'Secondary secant piles') headerCheckItems = secondarySecantItems
        else if (subsection.name === 'Tangent piles') headerCheckItems = tangentPileItems
        else if (subsection.name === 'Sheet pile') headerCheckItems = sheetPileItems
        else if (subsection.name === 'Timber lagging') headerCheckItems = timberLaggingItems
        else if (subsection.name === 'Timber sheeting') headerCheckItems = timberSheetingItems
        else if (subsection.name === 'Waler') headerCheckItems = walerItems
        else if (subsection.name === 'Raker') headerCheckItems = rakerItems
        else if (subsection.name === 'Upper Raker') headerCheckItems = upperRakerItems
        else if (subsection.name === 'Lower Raker') headerCheckItems = lowerRakerItems
        else if (subsection.name === 'Stand off') headerCheckItems = standOffItems
        else if (subsection.name === 'Kicker') headerCheckItems = kickerItems
        else if (subsection.name === 'Channel') headerCheckItems = channelItems
        else if (subsection.name === 'Roll chock') headerCheckItems = rollChockItems
        else if (subsection.name === 'Stud beam') headerCheckItems = studBeamItems
        else if (subsection.name === 'Inner corner brace') headerCheckItems = innerCornerBraceItems
        else if (subsection.name === 'Knee brace') headerCheckItems = kneeBraceItems
        else if (subsection.name === 'Supporting angle') headerCheckItems = supportingAngleGroups
        else if (subsection.name === 'Parging') headerCheckItems = pargingItems
        else if (subsection.name === 'Heel blocks') headerCheckItems = heelBlockItems
        else if (subsection.name === 'Underpinning') headerCheckItems = underpinningItems
        else if (subsection.name === 'Rock anchors') headerCheckItems = rockAnchorItems
        else if (subsection.name === 'Rock bolts') headerCheckItems = rockBoltItems
        else if (subsection.name === 'Anchor') headerCheckItems = anchorItems
        else if (subsection.name === 'Tie back') headerCheckItems = tieBackItems
        else if (subsection.name === 'Concrete soil retention piers') headerCheckItems = concreteSoilRetentionPierItems
        else if (subsection.name === 'Guide wall') headerCheckItems = guideWallItems
        else if (subsection.name === 'Dowel bar') headerCheckItems = dowelBarItems
        else if (subsection.name === 'Rock pins') headerCheckItems = rockPinItems
        else if (subsection.name === 'Shotcrete') headerCheckItems = shotcreteItems
        else if (subsection.name === 'Permission grouting') headerCheckItems = permissionGroutingItems
        else if (subsection.name === 'Buttons') headerCheckItems = buttonItems
        else if (subsection.name === 'Rock stabilization') headerCheckItems = rockStabilizationItems
        else if (subsection.name === 'Form board') headerCheckItems = formBoardItems

        let hasSubsectionData = headerCheckItems.length > 0
        if (subsection.name === 'Backpacking' && hasBackpacking) hasSubsectionData = true

        if (hasSubsectionData) {
          // Add subsection header (indented)
          const subsectionRow = Array(template.columns.length).fill('')
          subsectionRow[1] = subsection.name + ':'
          rows.push(subsectionRow)
        }

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
        } else if (subsection.name === 'Timber soldier piles' && timberSoldierPileGroups.length > 0) {
          // Process each group for timber soldier piles (same structure as Drilled soldier pile, but no column K)
          timberSoldierPileGroups.forEach((group, groupIndex) => {
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
            formulas.push({ row: rows.length, itemType: 'timber_soldier_pile_group_sum', section: 'soe', firstDataRow: firstGroupRow, lastDataRow: rows.length - 1 })
            if (groupIndex < timberSoldierPileGroups.length - 1) rows.push(Array(template.columns.length).fill(''))
          })
        } else if (subsection.name === 'Timber planks' && timberPlankGroups.length > 0) {
          // Process each group for timber planks (no column K, H not rounded). If (N) at start, E=qty, I=H*C*E, M=C*E
          timberPlankGroups.forEach((group, groupIndex) => {
            const firstGroupRow = rows.length + 1
            group.items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[7] = item.parsed.calculatedHeight || ''
              if (item.parsed.qty != null) {
                itemRow[4] = item.parsed.qty
              }
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'soldier_pile_item', parsedData: item, section: 'soe' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'timber_plank_group_sum', section: 'soe', firstDataRow: firstGroupRow, lastDataRow: rows.length - 1 })
            if (groupIndex < timberPlankGroups.length - 1) rows.push(Array(template.columns.length).fill(''))
          })
        } else if (subsection.name === 'Timber waler' && timberWalerGroups.length > 0) {
          // Timber waler: E=qty from (N) or 1, F=manual, I=E*C, M=E*F. Sum I and M only if multiple items.
          const timberWalerTotalItems = timberWalerGroups.reduce((sum, g) => sum + g.items.length, 0)
          const hasMultipleTimberWalerItems = timberWalerTotalItems > 1
          timberWalerGroups.forEach((group, groupIndex) => {
            const firstGroupRow = rows.length + 1
            group.items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[4] = item.parsed.qty ?? 1
              // Col F (index 5) = length, empty for manual input
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'timber_waler_item', parsedData: item, section: 'soe', hasMultipleItems: hasMultipleTimberWalerItems })
            })
            if (hasMultipleTimberWalerItems) {
              const sumRow = Array(template.columns.length).fill('')
              rows.push(sumRow)
              formulas.push({ row: rows.length, itemType: 'timber_waler_group_sum', section: 'soe', firstDataRow: firstGroupRow, lastDataRow: rows.length - 1 })
            }
            if (groupIndex < timberWalerGroups.length - 1) rows.push(Array(template.columns.length).fill(''))
          })
        } else if (subsection.name === 'Timber raker' && timberRakerGroups.length > 0) {
          // Process each group for timber raker (same as timber planks: no column K, H not rounded)
          timberRakerGroups.forEach((group, groupIndex) => {
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
            formulas.push({ row: rows.length, itemType: 'timber_raker_group_sum', section: 'soe', firstDataRow: firstGroupRow, lastDataRow: rows.length - 1 })
            if (groupIndex < timberRakerGroups.length - 1) rows.push(Array(template.columns.length).fill(''))
          })
        } else if (subsection.name === 'Timber brace' && timberBraceGroups.length > 0) {
          // Timber brace: per-group - multiple same-type items -> sum I and M, items black; single item or merged (different types) -> no sum, I and M red
          timberBraceGroups.forEach((group, groupIndex) => {
            const firstGroupRow = rows.length + 1
            const hasMultipleItemsInGroup = group.items.length > 1 && !group.isMerged
            group.items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              if (item.parsed.type === 'timber_brace_corner') {
                itemRow[4] = item.parsed.qty ?? 1
                // Col F (index 5) = length, empty for manual input
              } else {
                itemRow[7] = item.parsed.calculatedHeight || ''
              }
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'timber_brace_item', parsedData: item, section: 'soe', hasMultipleItems: hasMultipleItemsInGroup })
            })
            if (hasMultipleItemsInGroup) {
              const sumRow = Array(template.columns.length).fill('')
              rows.push(sumRow)
              formulas.push({ row: rows.length, itemType: 'timber_brace_group_sum', section: 'soe', firstDataRow: firstGroupRow, lastDataRow: rows.length - 1 })
            }
            if (groupIndex < timberBraceGroups.length - 1) rows.push(Array(template.columns.length).fill(''))
          })
        } else if (subsection.name === 'Timber post' && timberPostGroups.length > 0) {
          // Process each group for timber post (same as timber soldier piles: no column K)
          timberPostGroups.forEach((group, groupIndex) => {
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
            formulas.push({ row: rows.length, itemType: 'timber_post_group_sum', section: 'soe', firstDataRow: firstGroupRow, lastDataRow: rows.length - 1 })
            if (groupIndex < timberPostGroups.length - 1) rows.push(Array(template.columns.length).fill(''))
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
        if (hasSubsectionData) {
          rows.push(Array(template.columns.length).fill(''))
        }
      })
    } else if (section.section === 'Foundation') {
      // Foundation section: put CY in col D of heading row; col C will get sum formula later
      const foundationHeaderRow = rows.length - 1 // 1-based row of section header (last pushed was empty row)
      if (rows.length >= 2) {
        rows[rows.length - 2][3] = 'CY' // Col D = "CY" on section header row
      }
      // Handle Foundation section
      section.subsections.forEach((subsection) => {
        // Check if subsection has data
        let hasSubsectionData = false
        if (subsection.name === 'Piles') hasSubsectionData = miscellaneousPileItems.length > 0
        else if (subsection.name === 'Drilled foundation pile') hasSubsectionData = drilledFoundationPileGroups.length > 0
        else if (subsection.name === 'Helical foundation pile') hasSubsectionData = helicalFoundationPileGroups.length > 0
        else if (subsection.name === 'Driven foundation pile') hasSubsectionData = drivenFoundationPileItems.length > 0
        else if (subsection.name === 'Stelcor drilled displacement pile') hasSubsectionData = stelcorDrilledDisplacementPileItems.length > 0
        else if (subsection.name === 'CFA pile') hasSubsectionData = cfaPileItems.length > 0
        else if (subsection.name === 'Pile caps') hasSubsectionData = pileCapItems.length > 0
        else if (subsection.name === 'Strip Footings') hasSubsectionData = stripFootingGroups.length > 0
        else if (subsection.name === 'Isolated Footings') hasSubsectionData = isolatedFootingItems.length > 0
        else if (subsection.name === 'Pilaster') hasSubsectionData = pilasterItems.length > 0
        else if (subsection.name === 'Grade beams') hasSubsectionData = gradeBeamGroups.length > 0
        else if (subsection.name === 'Tie beam') hasSubsectionData = tieBeamGroups.length > 0
        else if (subsection.name === 'Thickened slab') hasSubsectionData = thickenedSlabGroups.length > 0
        else if (subsection.name === 'Buttresses') hasSubsectionData = !!buttressItem
        else if (subsection.name === 'Pier') hasSubsectionData = pierItems.length > 0
        else if (subsection.name === 'Corbel') hasSubsectionData = corbelGroups.length > 0
        else if (subsection.name === 'Linear Wall') hasSubsectionData = linearWallGroups.length > 0
        else if (subsection.name === 'Foundation Wall') hasSubsectionData = foundationWallGroups.length > 0
        else if (subsection.name === 'Retaining walls') hasSubsectionData = retainingWallGroups.length > 0
        else if (subsection.name === 'Barrier wall') hasSubsectionData = barrierWallGroups.length > 0
        else if (subsection.name === 'Stem wall') hasSubsectionData = stemWallItems.length > 0
        else if (subsection.name === 'Elevator Pit') hasSubsectionData = elevatorPitItems.length > 0
        else if (subsection.name === 'Detention tank') hasSubsectionData = detentionTankItems.length > 0
        else if (subsection.name === 'Duplex sewage ejector pit') hasSubsectionData = duplexSewageEjectorPitItems.length > 0
        else if (subsection.name === 'Deep sewage ejector pit') hasSubsectionData = deepSewageEjectorPitItems.length > 0
        else if (subsection.name === 'Grease trap') hasSubsectionData = greaseTrapItems.length > 0
        else if (subsection.name === 'House trap') hasSubsectionData = houseTrapItems.length > 0
        else if (subsection.name === 'Mat slab') hasSubsectionData = matSlabItems.length > 0
        else if (subsection.name === 'Mud slab') hasSubsectionData = mudSlabFoundationItems.length > 0
        else if (subsection.name === 'SOG') hasSubsectionData = sogItems.length > 0
        else if (subsection.name === 'Stairs on grade') hasSubsectionData = stairsOnGradeGroups.length > 0
        else if (subsection.name === 'Proposed underground electrical conduit') hasSubsectionData = electricConduitItems.length > 0
        else if (subsection.name === 'For foundation Extra line item use this') hasSubsectionData = true // Always show extra line item placeholder if section exists? Or only if needed? Treating as always show if section active

        if (hasSubsectionData) {
          // Add subsection header (indented)
          const subsectionRow = Array(template.columns.length).fill('')
          subsectionRow[1] = subsection.name + ':'
          rows.push(subsectionRow)
        }

        if (subsection.name === 'Piles' && miscellaneousPileItems.length > 0) {
          miscellaneousPileItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || ''
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'foundation_piles_misc', parsedData: item, section: 'foundation', subsectionName: subsection.name })
          })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Drilled foundation pile' && drilledFoundationPileGroups.length > 0) {
          // Process each group for drilled foundation piles.
          // - Identical items are merged into one data row per group (C column).
          // - Each group gets its own sum row (separated).
          drilledFoundationPileGroups.forEach((group, groupIndex) => {
            const firstGroupRow = rows.length + 1
            const item = group.items[0]

            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff // Combined quantity (from raw)
            itemRow[3] = item.unit

            // Height is always in column H for this subsection (including isolation casing)
            itemRow[7] = item.parsed.calculatedHeight || ''
            // Column E remains manual input for isolation casing items (dual diameter) and is left blank here.
            // Leave F/G empty.

            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'drilled_foundation_pile', parsedData: item, section: 'foundation' })

            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: firstGroupRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              isDualDiameter: !!item.parsed?.isDualDiameter
            })

            if (groupIndex < drilledFoundationPileGroups.length - 1) {
              rows.push(Array(template.columns.length).fill(''))
            }
          })
        } else if (subsection.name === 'Helical foundation pile' && helicalFoundationPileGroups.length > 0) {
          // Process each group for helical foundation piles
          helicalFoundationPileGroups.forEach((group, groupIndex) => {
            const firstGroupRow = rows.length + 1
            group.items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[7] = item.parsed.calculatedHeight || '' // Height (G)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'helical_foundation_pile', parsedData: item, section: 'foundation' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: firstGroupRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name
            })
            if (groupIndex < helicalFoundationPileGroups.length - 1) {
              rows.push(Array(template.columns.length).fill(''))
            }
          })
        } else if (subsection.name === 'Driven foundation pile' && drivenFoundationPileItems.length > 0) {
          const firstItemRow = rows.length + 1
          drivenFoundationPileItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit
            itemRow[7] = item.parsed.calculatedHeight || '' // Height (G)
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'driven_foundation_pile', parsedData: item, section: 'foundation' })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({
            row: rows.length,
            itemType: 'foundation_sum',
            section: 'foundation',
            firstDataRow: firstItemRow,
            lastDataRow: rows.length - 1,
            subsectionName: subsection.name,
            foundationCySumRow: false // Do not include Driven foundation pile CY in Foundation section total
          })
        } else if (subsection.name === 'Stelcor drilled displacement pile' && stelcorDrilledDisplacementPileItems.length > 0) {
          const firstItemRow = rows.length + 1
          stelcorDrilledDisplacementPileItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit
            itemRow[7] = item.parsed.calculatedHeight || '' // Height (G)
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'stelcor_drilled_displacement_pile', parsedData: item, section: 'foundation' })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({
            row: rows.length,
            itemType: 'foundation_sum',
            section: 'foundation',
            firstDataRow: firstItemRow,
            lastDataRow: rows.length - 1,
            subsectionName: subsection.name
          })
        } else if (subsection.name === 'CFA pile' && cfaPileItems.length > 0) {
          const firstItemRow = rows.length + 1
          cfaPileItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit
            itemRow[7] = item.parsed.calculatedHeight || '' // Height (G)
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'cfa_pile', parsedData: item, section: 'foundation' })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({
            row: rows.length,
            itemType: 'foundation_sum',
            section: 'foundation',
            firstDataRow: firstItemRow,
            lastDataRow: rows.length - 1,
            subsectionName: subsection.name,
            foundationCySumRow: false // Do not include CFA pile CY in Foundation section total
          })
        } else if (subsection.name === 'Pile caps' && pileCapItems.length > 0) {
          const firstItemRow = rows.length + 1
          pileCapItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit
            itemRow[5] = item.parsed.length || '' // Length (F)
            itemRow[6] = item.parsed.width || '' // Width (G)
            itemRow[7] = item.parsed.height || '' // Height (H)
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'pile_cap', parsedData: item, section: 'foundation' })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({
            row: rows.length,
            itemType: 'foundation_sum',
            section: 'foundation',
            firstDataRow: firstItemRow,
            lastDataRow: rows.length - 1,
            subsectionName: subsection.name,
            foundationCySumRow: true
          })
        } else if (subsection.name === 'Strip Footings' && stripFootingGroups.length > 0) {
          // Process each group for strip footings - no space between items
          const firstItemRow = rows.length + 1
          stripFootingGroups.forEach((group, groupIndex) => {
            group.items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[6] = item.parsed.width || '' // Width (G) - from first bracket value
              itemRow[7] = item.parsed.height || '' // Height (H) - from second bracket value
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'strip_footing', parsedData: item, section: 'foundation' })
            })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({
            row: rows.length,
            itemType: 'foundation_sum',
            section: 'foundation',
            firstDataRow: firstItemRow,
            lastDataRow: rows.length - 1,
            subsectionName: subsection.name,
            foundationCySumRow: true
          })
        } else if (subsection.name === 'Isolated Footings' && isolatedFootingItems.length > 0) {
          const firstItemRow = rows.length + 1
          isolatedFootingItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit
            itemRow[5] = item.parsed.length || '' // Length (F)
            itemRow[6] = item.parsed.width || '' // Width (G)
            itemRow[7] = item.parsed.height || '' // Height (H)
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'isolated_footing', parsedData: item, section: 'foundation' })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({
            row: rows.length,
            itemType: 'foundation_sum',
            section: 'foundation',
            firstDataRow: firstItemRow,
            lastDataRow: rows.length - 1,
            subsectionName: subsection.name,
            foundationCySumRow: true
          })
        } else if (subsection.name === 'Pilaster' && pilasterItems.length > 0) {
          const firstItemRow = rows.length + 1
          pilasterItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit
            itemRow[4] = 1 // QTY (E) = 1
            itemRow[5] = item.parsed.length || '' // Length (F)
            itemRow[6] = item.parsed.width || '' // Width (G)
            itemRow[7] = item.parsed.height || '' // Height (H)
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'pilaster', parsedData: item, section: 'foundation' })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({
            row: rows.length,
            itemType: 'foundation_sum',
            section: 'foundation',
            firstDataRow: firstItemRow,
            lastDataRow: rows.length - 1,
            subsectionName: subsection.name,
            foundationCySumRow: true
          })
        } else if (subsection.name === 'Grade beams' && gradeBeamGroups.length > 0) {
          // All Grade beams items should be together under a single sum.
          const firstItemRow = rows.length + 1
          gradeBeamGroups.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit
            itemRow[6] = item.parsed.width || '' // Width (G)
            itemRow[7] = item.parsed.height || '' // Height (H)
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'grade_beam', parsedData: item, section: 'foundation' })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({
            row: rows.length,
            itemType: 'foundation_sum',
            section: 'foundation',
            firstDataRow: firstItemRow,
            lastDataRow: rows.length - 1,
            subsectionName: subsection.name,
            foundationCySumRow: true
          })
        } else if (subsection.name === 'Tie beam' && tieBeamGroups.length > 0) {
          // Process each group for tie beams
          tieBeamGroups.forEach((group, groupIndex) => {
            const firstGroupRow = rows.length + 1
            group.items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[6] = item.parsed.width || '' // Width (G)
              itemRow[7] = item.parsed.height || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'tie_beam', parsedData: item, section: 'foundation' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: firstGroupRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              foundationCySumRow: true
            })
            if (groupIndex < tieBeamGroups.length - 1) {
              rows.push(Array(template.columns.length).fill(''))
            }
          })
        } else if (subsection.name === 'Thickened slab' && thickenedSlabGroups.length > 0) {
          // Process each group for thickened slabs
          thickenedSlabGroups.forEach((group, groupIndex) => {
            const firstGroupRow = rows.length + 1
            group.items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[6] = item.parsed.width || '' // Width (G)
              itemRow[7] = item.parsed.height || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'thickened_slab', parsedData: item, section: 'foundation' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: firstGroupRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              foundationCySumRow: true
            })
            if (groupIndex < thickenedSlabGroups.length - 1) {
              rows.push(Array(template.columns.length).fill(''))
            }
          })
        } else if (subsection.name === 'Buttresses' && buttressItem) {
          // Add "As per Takeoff count" row
          const itemRow = Array(template.columns.length).fill('')
          itemRow[1] = 'As per Takeoff count'
          itemRow[2] = buttressItem.takeoff
          itemRow[3] = buttressItem.unit
          itemRow[4] = 1
          itemRow[5] = 1
          itemRow[6] = 1
          itemRow[7] = 1
          rows.push(itemRow)
          const buttressRow = rows.length
          formulas.push({ row: rows.length, itemType: 'buttress_takeoff', parsedData: buttressItem, section: 'foundation' })

          // Add empty row
          rows.push(Array(template.columns.length).fill(''))

          // Add "Final as per schedule count" row (manual entry)
          const finalRow = Array(template.columns.length).fill('')
          finalRow[1] = 'Final as per schedule count'
          finalRow[3] = 'EA'
          finalRow[8] = 144 // FT (I) - manual
          finalRow[9] = 148.5 // SQ FT (J) - manual
          finalRow[11] = 12.1204 // CY (L) - manual
          finalRow[12] = 14 // QTY (M) - manual, but should reference buttress count
          rows.push(finalRow)
          formulas.push({ row: rows.length, itemType: 'buttress_final', section: 'foundation', buttressRow: buttressRow, foundationCySumRow: true })
        } else if (subsection.name === 'Pier' && pierItems.length > 0) {
          // Process each group for pier items
          pierItems.forEach((group, groupIndex) => {
            const firstGroupRow = rows.length + 1
            group.items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[4] = 1 // QTY (E) = 1
              itemRow[5] = item.parsed.length || '' // Length (F)
              itemRow[6] = item.parsed.width || '' // Width (G)
              itemRow[7] = item.parsed.height || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'pier', parsedData: item, section: 'foundation' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: firstGroupRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              foundationCySumRow: true
            })
            if (groupIndex < pierItems.length - 1) {
              rows.push(Array(template.columns.length).fill(''))
            }
          })
        } else if (subsection.name === 'Corbel' && corbelGroups.length > 0) {
          // Process each group for corbel items
          corbelGroups.forEach((group, groupIndex) => {
            const firstGroupRow = rows.length + 1
            group.items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[6] = item.parsed.width || '' // Width (G)
              itemRow[7] = item.parsed.height || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'corbel', parsedData: item, section: 'foundation' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: firstGroupRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              foundationCySumRow: true
            })
            if (groupIndex < corbelGroups.length - 1) {
              rows.push(Array(template.columns.length).fill(''))
            }
          })
        } else if (subsection.name === 'Linear Wall' && linearWallGroups.length > 0) {
          // Process each group for linear wall items
          linearWallGroups.forEach((group, groupIndex) => {
            const firstGroupRow = rows.length + 1
            group.items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[6] = item.parsed.width || '' // Width (G)
              itemRow[7] = item.parsed.height || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'linear_wall', parsedData: item, section: 'foundation' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: firstGroupRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              foundationCySumRow: true
            })
            if (groupIndex < linearWallGroups.length - 1) {
              rows.push(Array(template.columns.length).fill(''))
            }
          })
        } else if (subsection.name === 'Foundation Wall' && foundationWallGroups.length > 0) {
          // Process each group for foundation wall items
          foundationWallGroups.forEach((group, groupIndex) => {
            const firstGroupRow = rows.length + 1
            group.items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[6] = item.parsed.width || '' // Width (G)
              itemRow[7] = item.parsed.height || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'foundation_wall', parsedData: item, section: 'foundation' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: firstGroupRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              foundationCySumRow: true
            })
            if (groupIndex < foundationWallGroups.length - 1) {
              rows.push(Array(template.columns.length).fill(''))
            }
          })
        } else if (subsection.name === 'Retaining walls' && retainingWallGroups.length > 0) {
          // Process each group for retaining walls
          retainingWallGroups.forEach((group, groupIndex) => {
            const firstGroupRow = rows.length + 1
            group.items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[6] = item.parsed.width || '' // Width (G)
              itemRow[7] = item.parsed.height || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'retaining_wall', parsedData: item, section: 'foundation' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: firstGroupRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              foundationCySumRow: true
            })
            if (groupIndex < retainingWallGroups.length - 1) {
              rows.push(Array(template.columns.length).fill(''))
            }
          })
        } else if (subsection.name === 'Barrier wall' && barrierWallGroups.length > 0) {
          // Process each group for barrier walls
          barrierWallGroups.forEach((group, groupIndex) => {
            const firstGroupRow = rows.length + 1
            group.items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[6] = item.parsed.width || '' // Width (G)
              itemRow[7] = item.parsed.height || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'barrier_wall', parsedData: item, section: 'foundation' })
            })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: firstGroupRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              foundationCySumRow: true
            })
            if (groupIndex < barrierWallGroups.length - 1) {
              rows.push(Array(template.columns.length).fill(''))
            }
          })
        } else if (subsection.name === 'Stem wall' && stemWallItems.length > 0) {
          const firstItemRow = rows.length + 1
          stemWallItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit
            itemRow[6] = item.parsed.width || '' // Width (G)
            itemRow[7] = item.parsed.height || '' // Height (H)
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'stem_wall', parsedData: item, section: 'foundation' })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({
            row: rows.length,
            itemType: 'foundation_sum',
            section: 'foundation',
            firstDataRow: firstItemRow,
            lastDataRow: rows.length - 1,
            subsectionName: subsection.name,
            foundationCySumRow: true
          })
        } else if (subsection.name === 'Elevator Pit' && elevatorPitItems.length > 0) {
          const firstItemRow = rows.length + 1

          // Group elevator pit items by sub-type for proper rendering
          const slabItems = []
          const wallItems = []
          const slopeItems = []

          elevatorPitItems.forEach(item => {
            const subType = item.parsed?.itemSubType
            if (subType === 'slab') {
              slabItems.push(item)
            } else if (subType === 'wall') {
              wallItems.push(item)
            } else if (subType === 'slope_transition') {
              slopeItems.push(item)
            }
            // Note: sump_pit items are not added from data, only the manual one with 2 EA
          })

          // Add manual "Sump pit" item with value 2 EA
          const sumpPitRow = Array(template.columns.length).fill('')
          sumpPitRow[1] = 'Sump pit'
          sumpPitRow[2] = 2
          sumpPitRow[3] = 'EA'
          rows.push(sumpPitRow)
          formulas.push({
            row: rows.length,
            itemType: 'elevator_pit',
            parsedData: { particulars: 'Sump pit', takeoff: 2, unit: 'EA', parsed: { type: 'elevator_pit', itemSubType: 'sump_pit' } },
            section: 'foundation',
            foundationCySumRow: true
          })

          // Add slab items
          const slabFirstRow = slabItems.length > 0 ? rows.length + 1 : null
          if (slabFirstRow) foundationSlabRows.elevatorPit = slabFirstRow
          slabItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit
            itemRow[7] = item.parsed.heightFromH || '' // Height (H)
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'elevator_pit', parsedData: item, section: 'foundation' })
          })

          // Add sum row for slab items
          if (slabItems.length > 0) {
            const slabSumRow = Array(template.columns.length).fill('')
            rows.push(slabSumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: slabFirstRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              foundationCySumRow: true
            })
          }

          // Add empty row between slab and wall
          if (wallItems.length > 0) {
            rows.push(Array(template.columns.length).fill(''))
          }

          // Group wall items by size (groupKey)
          const wallGroups = new Map()
          wallItems.forEach(item => {
            const groupKey = item.parsed.groupKey || 'OTHER'
            if (!wallGroups.has(groupKey)) {
              wallGroups.set(groupKey, [])
            }
            wallGroups.get(groupKey).push(item)
          })

          // Add wall items grouped by size
          const wallGroupsToIterate = getMergedGroupsIfNeeded(wallGroups)
          wallGroupsToIterate.forEach((group, groupIndex) => {
            const { items } = group
            const wallGroupFirstRow = rows.length + 1
            items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[6] = item.parsed.width || '' // Width (G)
              itemRow[7] = item.parsed.height || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'elevator_pit', parsedData: item, section: 'foundation' })
            })
            // Add sum row for this wall group
            const wallSumRow = Array(template.columns.length).fill('')
            rows.push(wallSumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: wallGroupFirstRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              foundationCySumRow: true
            })
            // Add empty row between wall groups
            if (groupIndex < wallGroupsToIterate.length - 1 || slopeItems.length > 0) {
              rows.push(Array(template.columns.length).fill(''))
            }
          })

          // Add empty row between wall and slope
          if (slopeItems.length > 0) {
            rows.push(Array(template.columns.length).fill(''))
          }

          // Group slope items by size (groupKey)
          const slopeGroups = new Map()
          slopeItems.forEach(item => {
            const groupKey = item.parsed.groupKey || 'OTHER'
            if (!slopeGroups.has(groupKey)) {
              slopeGroups.set(groupKey, [])
            }
            slopeGroups.get(groupKey).push(item)
          })

          // Add slope items grouped by size
          getMergedGroupsIfNeeded(slopeGroups).forEach((group) => {
            const { items } = group
            const slopeGroupFirstRow = rows.length + 1
            items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[6] = item.parsed.width || '' // Width (G)
              itemRow[7] = item.parsed.height || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'elevator_pit', parsedData: item, section: 'foundation' })
            })
            // Add sum row for this slope group
            const slopeSumRow = Array(template.columns.length).fill('')
            rows.push(slopeSumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: slopeGroupFirstRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              foundationCySumRow: true
            })
          })
        } else if (subsection.name === 'Detention tank' && detentionTankItems.length > 0) {
          // Group items by type
          const slabItems = []
          const lidSlabItems = []
          const wallItems = []

          detentionTankItems.forEach(item => {
            const subType = item.parsed?.itemSubType
            if (subType === 'slab') {
              slabItems.push(item)
            } else if (subType === 'lid_slab') {
              lidSlabItems.push(item)
            } else if (subType === 'wall') {
              wallItems.push(item)
            }
          })

          // Add slab items
          if (slabItems.length > 0) {
            const slabFirstRow = rows.length + 1
            foundationSlabRows.detentionTank = slabFirstRow
            slabItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[7] = item.parsed.heightFromName || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'detention_tank', parsedData: item, section: 'foundation' })
            })
            const slabSumRow = Array(template.columns.length).fill('')
            rows.push(slabSumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: slabFirstRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              excludeISum: true, // Exclude I sum for slab items
              foundationCySumRow: true
            })
            rows.push(Array(template.columns.length).fill(''))
          }

          // Add lid slab items
          if (lidSlabItems.length > 0) {
            const lidSlabFirstRow = rows.length + 1
            lidSlabItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[7] = item.parsed.heightFromName || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'detention_tank', parsedData: item, section: 'foundation' })
            })
            const lidSlabSumRow = Array(template.columns.length).fill('')
            rows.push(lidSlabSumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: lidSlabFirstRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              excludeISum: true, // Exclude I sum for lid slab items
              foundationCySumRow: true
            })
            rows.push(Array(template.columns.length).fill(''))
          }

          // Group wall items by size
          if (wallItems.length > 0) {
            const wallGroups = new Map()
            wallItems.forEach(item => {
              const groupKey = item.parsed.groupKey || 'OTHER'
              if (!wallGroups.has(groupKey)) {
                wallGroups.set(groupKey, [])
              }
              wallGroups.get(groupKey).push(item)
            })

            getMergedGroupsIfNeeded(wallGroups).forEach((group) => {
              const { items } = group
              const wallGroupFirstRow = rows.length + 1
              items.forEach(item => {
                const itemRow = Array(template.columns.length).fill('')
                itemRow[1] = item.particulars
                itemRow[2] = item.takeoff
                itemRow[3] = item.unit
                // F (Length) should be empty for Detention tank wall items
                itemRow[6] = item.parsed.width || '' // Width (G)
                itemRow[7] = item.parsed.height || '' // Height (H)
                rows.push(itemRow)
                formulas.push({ row: rows.length, itemType: 'detention_tank', parsedData: item, section: 'foundation' })
              })
              const wallSumRow = Array(template.columns.length).fill('')
              rows.push(wallSumRow)
              formulas.push({
                row: rows.length,
                itemType: 'foundation_sum',
                section: 'foundation',
                firstDataRow: wallGroupFirstRow,
                lastDataRow: rows.length - 1,
                subsectionName: subsection.name,
                foundationCySumRow: true
              })
            })
          }
        } else if (subsection.name === 'Duplex sewage ejector pit' && duplexSewageEjectorPitItems.length > 0) {
          // Group items by type
          const slabItems = []
          const wallItems = []

          duplexSewageEjectorPitItems.forEach(item => {
            const subType = item.parsed?.itemSubType
            if (subType === 'slab') {
              slabItems.push(item)
            } else if (subType === 'wall') {
              wallItems.push(item)
            }
          })

          // Add slab items
          if (slabItems.length > 0) {
            const slabFirstRow = rows.length + 1
            foundationSlabRows.duplexSewageEjectorPit = slabFirstRow
            slabItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[7] = item.parsed.heightFromName || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'duplex_sewage_ejector_pit', parsedData: item, section: 'foundation' })
            })
            const slabSumRow = Array(template.columns.length).fill('')
            rows.push(slabSumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: slabFirstRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              excludeISum: true, // Exclude I sum for slab items
              foundationCySumRow: true
            })
            rows.push(Array(template.columns.length).fill(''))
          }

          // Group wall items by size
          if (wallItems.length > 0) {
            const wallGroups = new Map()
            wallItems.forEach(item => {
              const groupKey = item.parsed.groupKey || 'OTHER'
              if (!wallGroups.has(groupKey)) {
                wallGroups.set(groupKey, [])
              }
              wallGroups.get(groupKey).push(item)
            })

            getMergedGroupsIfNeeded(wallGroups).forEach((group) => {
              const { items } = group
              const wallGroupFirstRow = rows.length + 1
              items.forEach(item => {
                const itemRow = Array(template.columns.length).fill('')
                itemRow[1] = item.particulars
                itemRow[2] = item.takeoff
                itemRow[3] = item.unit
                itemRow[6] = item.parsed.width || '' // Width (G)
                itemRow[7] = item.parsed.height || '' // Height (H)
                rows.push(itemRow)
                formulas.push({ row: rows.length, itemType: 'duplex_sewage_ejector_pit', parsedData: item, section: 'foundation' })
              })
              const wallSumRow = Array(template.columns.length).fill('')
              rows.push(wallSumRow)
              formulas.push({
                row: rows.length,
                itemType: 'foundation_sum',
                section: 'foundation',
                firstDataRow: wallGroupFirstRow,
                lastDataRow: rows.length - 1,
                subsectionName: subsection.name,
                foundationCySumRow: true
              })
            })
          }
        } else if (subsection.name === 'Deep sewage ejector pit' && deepSewageEjectorPitItems.length > 0) {
          // Group items by type
          const slabItems = []
          const wallItems = []

          deepSewageEjectorPitItems.forEach(item => {
            const subType = item.parsed?.itemSubType
            if (subType === 'slab') {
              slabItems.push(item)
            } else if (subType === 'wall') {
              wallItems.push(item)
            }
          })

          // Add slab items
          if (slabItems.length > 0) {
            const slabFirstRow = rows.length + 1
            foundationSlabRows.deepSewageEjectorPit = slabFirstRow
            slabItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[7] = item.parsed.heightFromName || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'deep_sewage_ejector_pit', parsedData: item, section: 'foundation' })
            })
            const slabSumRow = Array(template.columns.length).fill('')
            rows.push(slabSumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: slabFirstRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              excludeISum: true, // Exclude I sum for slab items
              foundationCySumRow: true
            })
            rows.push(Array(template.columns.length).fill(''))
          }

          // Group wall items by size
          if (wallItems.length > 0) {
            const wallGroups = new Map()
            wallItems.forEach(item => {
              const groupKey = item.parsed.groupKey || 'OTHER'
              if (!wallGroups.has(groupKey)) {
                wallGroups.set(groupKey, [])
              }
              wallGroups.get(groupKey).push(item)
            })

            getMergedGroupsIfNeeded(wallGroups).forEach((group) => {
              const { items } = group
              const wallGroupFirstRow = rows.length + 1
              items.forEach(item => {
                const itemRow = Array(template.columns.length).fill('')
                itemRow[1] = item.particulars
                itemRow[2] = item.takeoff
                itemRow[3] = item.unit
                itemRow[6] = item.parsed.width || '' // Width (G)
                itemRow[7] = item.parsed.height || '' // Height (H)
                rows.push(itemRow)
                formulas.push({ row: rows.length, itemType: 'deep_sewage_ejector_pit', parsedData: item, section: 'foundation' })
              })
              const wallSumRow = Array(template.columns.length).fill('')
              rows.push(wallSumRow)
              formulas.push({
                row: rows.length,
                itemType: 'foundation_sum',
                section: 'foundation',
                firstDataRow: wallGroupFirstRow,
                lastDataRow: rows.length - 1,
                subsectionName: subsection.name,
                foundationCySumRow: true
              })
            })
          }
        } else if (subsection.name === 'Grease trap' && greaseTrapItems.length > 0) {
          // Group items by type
          const slabItems = []
          const wallItems = []

          greaseTrapItems.forEach(item => {
            const subType = item.parsed?.itemSubType
            if (subType === 'slab') {
              slabItems.push(item)
            } else if (subType === 'wall') {
              wallItems.push(item)
            }
          })

          // Add slab items
          if (slabItems.length > 0) {
            const slabFirstRow = rows.length + 1
            foundationSlabRows.greaseTrap = slabFirstRow
            slabItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[7] = item.parsed.heightFromName || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'grease_trap', parsedData: item, section: 'foundation' })
            })
            const slabSumRow = Array(template.columns.length).fill('')
            rows.push(slabSumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: slabFirstRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              excludeISum: true, // Exclude I sum for slab items
              foundationCySumRow: true
            })
            rows.push(Array(template.columns.length).fill(''))
          }

          // Group wall items by size
          if (wallItems.length > 0) {
            const wallGroups = new Map()
            wallItems.forEach(item => {
              const groupKey = item.parsed.groupKey || 'OTHER'
              if (!wallGroups.has(groupKey)) {
                wallGroups.set(groupKey, [])
              }
              wallGroups.get(groupKey).push(item)
            })

            getMergedGroupsIfNeeded(wallGroups).forEach((group) => {
              const { items } = group
              const wallGroupFirstRow = rows.length + 1
              items.forEach(item => {
                const itemRow = Array(template.columns.length).fill('')
                itemRow[1] = item.particulars
                itemRow[2] = item.takeoff
                itemRow[3] = item.unit
                itemRow[6] = item.parsed.width || '' // Width (G)
                itemRow[7] = item.parsed.height || '' // Height (H)
                rows.push(itemRow)
                formulas.push({ row: rows.length, itemType: 'grease_trap', parsedData: item, section: 'foundation' })
              })
              const wallSumRow = Array(template.columns.length).fill('')
              rows.push(wallSumRow)
              formulas.push({
                row: rows.length,
                itemType: 'foundation_sum',
                section: 'foundation',
                firstDataRow: wallGroupFirstRow,
                lastDataRow: rows.length - 1,
                subsectionName: subsection.name,
                foundationCySumRow: true
              })
            })
          }
        } else if (subsection.name === 'House trap' && houseTrapItems.length > 0) {
          // Group items by type
          const slabItems = []
          const wallItems = []

          houseTrapItems.forEach(item => {
            const subType = item.parsed?.itemSubType
            if (subType === 'slab') {
              slabItems.push(item)
            } else if (subType === 'wall') {
              wallItems.push(item)
            }
          })

          // Add slab items
          if (slabItems.length > 0) {
            const slabFirstRow = rows.length + 1
            foundationSlabRows.houseTrap = slabFirstRow
            slabItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              // G (Width) should be empty for pit slab items
              itemRow[7] = item.parsed.heightFromName || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'house_trap', parsedData: item, section: 'foundation' })
            })
            const slabSumRow = Array(template.columns.length).fill('')
            rows.push(slabSumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: slabFirstRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              excludeISum: true, // Exclude I sum for slab items
              foundationCySumRow: true
            })
            rows.push(Array(template.columns.length).fill(''))
          }

          // Group wall items by size
          if (wallItems.length > 0) {
            const wallGroups = new Map()
            wallItems.forEach(item => {
              const groupKey = item.parsed.groupKey || 'OTHER'
              if (!wallGroups.has(groupKey)) {
                wallGroups.set(groupKey, [])
              }
              wallGroups.get(groupKey).push(item)
            })

            getMergedGroupsIfNeeded(wallGroups).forEach((group) => {
              const { items } = group
              const wallGroupFirstRow = rows.length + 1
              items.forEach(item => {
                const itemRow = Array(template.columns.length).fill('')
                itemRow[1] = item.particulars
                itemRow[2] = item.takeoff
                itemRow[3] = item.unit
                itemRow[6] = item.parsed.width || '' // Width (G)
                itemRow[7] = item.parsed.height || '' // Height (H)
                rows.push(itemRow)
                formulas.push({ row: rows.length, itemType: 'house_trap', parsedData: item, section: 'foundation' })
              })
              const wallSumRow = Array(template.columns.length).fill('')
              rows.push(wallSumRow)
              formulas.push({
                row: rows.length,
                itemType: 'foundation_sum',
                section: 'foundation',
                firstDataRow: wallGroupFirstRow,
                lastDataRow: rows.length - 1,
                subsectionName: subsection.name,
                foundationCySumRow: true
              })
            })
          }
        } else if (subsection.name === 'Mat slab' && matSlabItems.length > 0) {
          // Group items by type and height
          const matItems = []
          const haunchItems = []

          matSlabItems.forEach(item => {
            const subType = item.parsed?.itemSubType
            if (subType === 'mat') {
              matItems.push(item)
            } else if (subType === 'haunch') {
              haunchItems.push(item)
            }
          })

          // Group mat items by height (groupKey)
          const matGroups = new Map()
          matItems.forEach(item => {
            const groupKey = item.parsed.groupKey || 'OTHER'
            if (!matGroups.has(groupKey)) {
              matGroups.set(groupKey, [])
            }
            matGroups.get(groupKey).push(item)
          })

          // Process each mat group with its associated haunch
          const matGroupsToIterate = getMergedGroupsIfNeeded(matGroups)
          matGroupsToIterate.forEach((group, groupIndex) => {
            const matGroupItems = group.items
            // Add mat items for this group
            const matFirstRow = rows.length + 1
            matGroupItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[7] = item.parsed.heightFromH || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'mat_slab', parsedData: item, section: 'foundation' })
            })
            const lastMatRow = rows.length

            // Add haunch item(s) for this group (if available) - when merged, add all corresponding haunch items
            let lastDataRow = lastMatRow
            const haunchIndices = matGroupItems.length === 1 ? [groupIndex] : Array.from({ length: matGroupItems.length }, (_, i) => i)
            haunchIndices.forEach((idx) => {
              if (haunchItems.length > idx) {
                const haunchItem = haunchItems[idx]
                const itemRow = Array(template.columns.length).fill('')
                itemRow[1] = haunchItem.particulars
                itemRow[2] = haunchItem.takeoff
                itemRow[3] = haunchItem.unit
                itemRow[6] = haunchItem.parsed.width || '' // Width (G)
                itemRow[7] = haunchItem.parsed.height || '' // Height (H)
                rows.push(itemRow)
                formulas.push({ row: rows.length, itemType: 'mat_slab', parsedData: haunchItem, section: 'foundation' })
                lastDataRow = rows.length // Update to include haunch item
              }
            })

            // Add sum row for mat items (J only, no I, no L)
            const matSumRow = Array(template.columns.length).fill('')
            rows.push(matSumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: matFirstRow,
              lastDataRow: lastMatRow,
              subsectionName: subsection.name,
              excludeISum: true, // Exclude I sum for mat items
              excludeLSum: true, // Exclude L sum (will be in combined sum)
              matSumOnly: true // Only sum J for mat items
            })

            // Add sum row for mat + haunch (L only, includes both mat and haunch)
            // This comes right after the mat sum row (no gap)
            const combinedSumRow = Array(template.columns.length).fill('')
            rows.push(combinedSumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: matFirstRow,
              lastDataRow: lastDataRow,
              subsectionName: subsection.name,
              excludeISum: true, // Exclude I sum
              excludeJSum: true, // Exclude J sum (only for mat items)
              cySumOnly: true, // Only sum L (CY) for mat + haunch combined
              foundationCySumRow: true
            })

            // Add empty row between groups
            if (groupIndex < matGroupsToIterate.length - 1) {
              rows.push(Array(template.columns.length).fill(''))
            }
          })
        } else if (subsection.name === 'Mud Slab') {
          // Mud Slab: Add a row with item name "Mud slab", Unit "SQ FT"
          const itemRow = Array(template.columns.length).fill('')
          itemRow[1] = 'Mud slab'
          itemRow[2] = '' // C is blank (manual entry)
          itemRow[3] = 'SQ FT' // Unit
          // H is blank (manual entry)
          rows.push(itemRow)
          formulas.push({
            row: rows.length,
            itemType: 'mud_slab_foundation',
            parsedData: { particulars: 'Mud slab', takeoff: 0, unit: 'SQ FT', parsed: { type: 'mud_slab_foundation', itemSubType: 'mud_slab' } },
            section: 'foundation',
            foundationCySumRow: true
          })
        } else if (subsection.name === 'SOG' && sogItems.length >= 0) {
          // Group items by type
          const gravelItems = []
          const gravelBackfillItems = []
          const geotextileItems = []
          const sogSlabItems = []
          const sogStepItems = []

          sogItems.forEach(item => {
            const subType = item.parsed?.itemSubType
            if (subType === 'gravel') {
              gravelItems.push(item)
            } else if (subType === 'gravel_backfill') {
              gravelBackfillItems.push(item)
            } else if (subType === 'geotextile') {
              geotextileItems.push(item)
            } else if (subType === 'sog_slab') {
              sogSlabItems.push(item)
            } else if (subType === 'sog_step') {
              sogStepItems.push(item)
            }
          })

          // Gravel Group
          const gravelFirstRow = rows.length + 1
          // Add manual "Gravel" item
          const gravelRow = Array(template.columns.length).fill('')
          gravelRow[1] = 'Gravel'
          gravelRow[2] = '' // C is blank
          gravelRow[3] = 'SQ FT'
          // H is blank (manual entry - they will fill it themselves)
          rows.push(gravelRow)
          formulas.push({
            row: rows.length,
            itemType: 'sog',
            parsedData: { particulars: 'Gravel', takeoff: 0, unit: 'SQ FT', parsed: { type: 'sog', itemSubType: 'gravel' } },
            section: 'foundation'
          })

          // Add Gravel backfill items
          gravelBackfillItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit
            itemRow[7] = item.parsed.heightFromH || '' // Height (H)
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'sog', parsedData: item, section: 'foundation' })
          })

          // Add sum row for Gravel group
          const gravelSumRow = Array(template.columns.length).fill('')
          rows.push(gravelSumRow)
          formulas.push({
            row: rows.length,
            itemType: 'foundation_sum',
            section: 'foundation',
            firstDataRow: gravelFirstRow,
            lastDataRow: rows.length - 1,
            subsectionName: subsection.name,
            excludeISum: true, // Exclude I sum
            foundationCySumRow: false // Do not include gravel group in Foundation CY total
          })
          rows.push(Array(template.columns.length).fill(''))

          // Geotextile Filter Fabric Group
          const geotextileFirstRow = rows.length + 1
          // Add manual "Geotextile filter fabric" items (2 items)
          for (let i = 0; i < 2; i++) {
            const geotextileRow = Array(template.columns.length).fill('')
            geotextileRow[1] = 'Geotextile filter fabric'
            geotextileRow[2] = '' // C is blank
            geotextileRow[3] = 'SQ FT'
            // H is blank
            rows.push(geotextileRow)
            formulas.push({
              row: rows.length,
              itemType: 'sog',
              parsedData: { particulars: 'Geotextile filter fabric', takeoff: 0, unit: 'SQ FT', parsed: { type: 'sog', itemSubType: 'geotextile' } },
              section: 'foundation'
            })
          }

          // Add existing Geotextile filter fabric items from data
          geotextileItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit
            // H is blank
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'sog', parsedData: item, section: 'foundation' })
          })

          // Add sum row for Geotextile group (J only, L is blank)
          const geotextileSumRow = Array(template.columns.length).fill('')
          rows.push(geotextileSumRow)
          formulas.push({
            row: rows.length,
            itemType: 'foundation_sum',
            section: 'foundation',
            firstDataRow: geotextileFirstRow,
            lastDataRow: rows.length - 1,
            subsectionName: subsection.name,
            excludeISum: true, // Exclude I sum
            excludeLSum: true // Exclude L sum (L is blank for geotextile)
          })
          rows.push(Array(template.columns.length).fill(''))

          // Group SOG slab items by groupKey
          const sogSlabGroups = new Map()
          sogSlabItems.forEach(item => {
            const groupKey = item.parsed.groupKey || 'OTHER'
            if (!sogSlabGroups.has(groupKey)) {
              sogSlabGroups.set(groupKey, [])
            }
            sogSlabGroups.get(groupKey).push(item)
          })

          // Add SOG slab groups
          const sogSlabGroupsToIterate = getMergedGroupsIfNeeded(sogSlabGroups)
          sogSlabGroupsToIterate.forEach((group, groupIndex) => {
            const { items } = group
            const slabGroupFirstRow = rows.length + 1
            items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit
              itemRow[7] = item.parsed.heightFromName || '' // Height (H)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'sog', parsedData: item, section: 'foundation' })
            })
            const slabSumRow = Array(template.columns.length).fill('')
            rows.push(slabSumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: slabGroupFirstRow,
              lastDataRow: rows.length - 1,
              subsectionName: subsection.name,
              excludeISum: true, // Exclude I sum
              foundationCySumRow: true
            })
            // Add empty row between groups
            if (groupIndex < sogSlabGroupsToIterate.length - 1 || sogStepItems.length > 0) {
              rows.push(Array(template.columns.length).fill(''))
            }
          })

          // Group SOG step items by size
          if (sogStepItems.length > 0) {
            const sogStepGroups = new Map()
            sogStepItems.forEach(item => {
              const groupKey = item.parsed.groupKey || 'OTHER'
              if (!sogStepGroups.has(groupKey)) {
                sogStepGroups.set(groupKey, [])
              }
              sogStepGroups.get(groupKey).push(item)
            })

            getMergedGroupsIfNeeded(sogStepGroups).forEach((group) => {
              const { items } = group
              const stepGroupFirstRow = rows.length + 1
              items.forEach(item => {
                const itemRow = Array(template.columns.length).fill('')
                itemRow[1] = item.particulars
                itemRow[2] = item.takeoff
                itemRow[3] = item.unit
                itemRow[6] = item.parsed.width || '' // Width (G)
                itemRow[7] = item.parsed.height || '' // Height (H)
                rows.push(itemRow)
                formulas.push({ row: rows.length, itemType: 'sog', parsedData: item, section: 'foundation' })
              })
              const stepSumRow = Array(template.columns.length).fill('')
              rows.push(stepSumRow)
              formulas.push({
                row: rows.length,
                itemType: 'foundation_sum',
                section: 'foundation',
                firstDataRow: stepGroupFirstRow,
                lastDataRow: rows.length - 1,
                subsectionName: subsection.name,
                foundationCySumRow: true
              })
            })
          }
        } else if (subsection.name === 'Stairs on grade Stairs' && stairsOnGradeGroups.length > 0) {
          // Process each stair group
          stairsOnGradeGroups.forEach((group, groupIndex) => {
            // Add group header (e.g., "Stair A:")
            const groupHeaderRow = Array(template.columns.length).fill('')
            groupHeaderRow[1] = `Stair ${group.stairIdentifier}:`
            rows.push(groupHeaderRow)
            formulas.push({ row: rows.length, itemType: 'stairs_on_grade_group_header', section: 'foundation' })

            const groupFirstRow = rows.length + 1
            let stairsOnGradeRow = null
            let stairsOnGradeItem = null

            // Process items in order: Landings first, then Stairs on grade
            const landingsItems = group.items.filter(item => item.parsed?.itemSubType === 'landings')
            const stairsItems = group.items.filter(item => item.parsed?.itemSubType === 'stairs')

            // Add Landings items first
            landingsItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = 'Landings'
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || ''
              // H = 0.67 (will be set by formula)
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'stairs_on_grade', parsedData: item, section: 'foundation' })
            })

            // Add extra row space after the landings row when both landings and stairs exist
            if (landingsItems.length > 0 && stairsItems.length > 0) {
              rows.push(Array(template.columns.length).fill(''))
            }

            // Add Stairs on grade items
            stairsItems.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars.includes('wide') ? item.particulars : 'Stairs on grade'
              itemRow[2] = item.takeoff
              itemRow[3] = 'Treads' // Replace "EA" with "Treads"
              itemRow[4] = item.takeoff // E (QTY) = takeoff value
              // F = 11/12 (will be set by formula)
              // G = width from name or manual input (will be set by formula or manual)
              if (item.parsed?.widthFromName !== undefined) {
                itemRow[6] = item.parsed.widthFromName // G (Width) from name
              }
              // H = 7/12 (will be set by formula)
              rows.push(itemRow)
              stairsOnGradeRow = rows.length
              stairsOnGradeItem = item
              formulas.push({ row: rows.length, itemType: 'stairs_on_grade', parsedData: item, section: 'foundation' })
            })

            // Generate Stair slab item if there's a Stairs on grade item
            if (stairsOnGradeItem && stairsOnGradeRow) {
              const stairSlabRow = Array(template.columns.length).fill('')
              stairSlabRow[1] = 'Stair slab'
              // C = C[stairs_row]*1.3 (will be set as formula)
              stairSlabRow[3] = 'FT'
              // F and G depend on whether stairs has width in name or manual input
              // G for stair slab is always formula =G[stairs_row] (set in foundationProcessor / Spreadsheet)
              if (!(stairsOnGradeItem.parsed?.widthFromName !== undefined)) {
                // Manual width: F = G[stairs_row] (formula reference), G = empty
                // F will be set as formula =G[stairs_row] in Spreadsheet.jsx
              }
              // H = 0.67 (will be set by formula)
              rows.push(stairSlabRow)
              formulas.push({
                row: rows.length,
                itemType: 'stairs_on_grade',
                parsedData: {
                  particulars: 'Stair slab',
                  takeoff: 0,
                  unit: 'FT',
                  parsed: {
                    type: 'stairs_on_grade',
                    itemSubType: 'stair_slab',
                    stairsRow: stairsOnGradeRow,
                    hasWidthFromName: stairsOnGradeItem.parsed?.widthFromName !== undefined
                  }
                },
                section: 'foundation'
              })
            }

            // Add sum row for the group (I/J/M exclude landings; L sum includes landings + space row + stairs + stair slab)
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            const firstDataRowForSum = landingsItems.length > 0
              ? groupFirstRow + landingsItems.length + (stairsItems.length > 0 ? 1 : 0) // skip landings rows + empty row
              : groupFirstRow
            const lastDataRowForGroup = rows.length - 1
            // Explicit L range so landings L is always included: from first group data row (landings) to last data row (stair slab).
            // Build as explicit cell list L52,L53,... so every row (landing, space, stairs, stair slab) is included.
            const lSumCells = []
            for (let r = groupFirstRow; r <= lastDataRowForGroup; r++) lSumCells.push(`L${r}`)
            const lSumRange = lSumCells.join(',')
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: firstDataRowForSum,
              lastDataRow: lastDataRowForGroup,
              subsectionName: subsection.name,
              foundationCySumRow: true,
              firstDataRowForL: groupFirstRow,
              lastDataRowForL: lastDataRowForGroup,
              lSumRange // explicit range string for L sum (includes landings)
            })

            // Add empty row between groups
            if (groupIndex < stairsOnGradeGroups.length - 1) {
              rows.push(Array(template.columns.length).fill(''))
            }
          })
        } else if (subsection.name === 'Electric conduit') {
          const conduitGroup1 = electricConduitItems.filter(item => {
            const p = (item.particulars || '').toLowerCase()
            return p.includes('underground electric conduit') || p.includes('electric conduit in slab')
          })
          const conduitGroup2 = electricConduitItems.filter(item => {
            const p = (item.particulars || '').toLowerCase()
            return p.includes('trench drain') || p.includes('perforated pipe')
          })
          const electricFirstRow = rows.length + 1
          conduitGroup1.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'FT'
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'electric_conduit', parsedData: item, section: 'foundation' })
          })
          const electricLastRow = conduitGroup1.length > 0 ? rows.length : electricFirstRow - 1
          if (conduitGroup1.length > 0) {
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({
              row: rows.length,
              itemType: 'foundation_sum',
              section: 'foundation',
              firstDataRow: electricFirstRow,
              lastDataRow: electricLastRow,
              subsectionName: subsection.name
            })
          }
          if (conduitGroup2.length > 0) {
            rows.push(Array(template.columns.length).fill(''))
            conduitGroup2.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'FT'
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'electric_conduit', parsedData: item, section: 'foundation' })
            })
          }
        } else if (subsection.name === 'For foundation Extra line item use this') {
          const extraItems = [
            { name: 'In SQ FT', unit: 'SQ FT', h: 1, type: 'foundation_extra_sqft' },
            { name: 'In FT', unit: 'FT', g: 1, h: 1, type: 'foundation_extra_ft' },
            { name: 'In EA', unit: 'EA', f: 1, g: 1, h: 1, type: 'foundation_extra_ea' }
          ]
          extraItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.name
            itemRow[2] = 1
            itemRow[3] = item.unit
            if (item.f) itemRow[5] = item.f
            if (item.g) itemRow[6] = item.g
            if (item.h) itemRow[7] = item.h
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: item.type, section: 'foundation' })
          })
        }

        // Add empty row after subsection
        if (hasSubsectionData) {
          rows.push(Array(template.columns.length).fill(''))
        }
      })
      // Foundation CY total: sum column L from included subsection sum rows (and data rows for Buttresses Final, Mud Slab, Elevator sump pit)
      const foundationCySumRows = formulas.filter(
        f => f.section === 'foundation' && f.foundationCySumRow === true
      ).map(f => f.row)
      formulas.push({
        row: foundationHeaderRow,
        itemType: 'foundation_section_cy_sum',
        section: 'foundation',
        sumRows: foundationCySumRows
      })
    } else if (section.section === 'Waterproofing') {
      section.subsections.forEach((subsection) => {
        let hasSubsectionData = false
        if (subsection.name === 'Exterior side') hasSubsectionData = exteriorSideItems.length > 0 || exteriorSidePitItems.length > 0
        else if (subsection.name === 'Negative side') hasSubsectionData = negativeSideWallItems.length > 0 || negativeSideSlabItems.length > 0
        else if (subsection.name === 'Horizontal') hasSubsectionData = exteriorSideItems.length > 0 || negativeSideWallItems.length > 0
        else if (subsection.name === 'For foundation Extra line item use this') hasSubsectionData = true

        if (hasSubsectionData) {
          const subsectionRow = Array(template.columns.length).fill('')
          subsectionRow[1] = subsection.name + ':'
          if (subsection.name === 'Exterior side') subsectionRow[7] = 'add +2'
          rows.push(subsectionRow)
          if (subsection.name === 'Exterior side') {
            formulas.push({ row: rows.length, itemType: 'waterproofing_exterior_side_header', section: 'waterproofing' })
          }
        }

        if (subsection.name === 'Exterior side' && (exteriorSideItems.length > 0 || exteriorSidePitItems.length > 0)) {
          const firstItemRow = rows.length + 1
          exteriorSideItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || ''
            if (item.parsed?.heightFromBracketPlus2 !== undefined && item.parsed?.heightFromBracketPlus2 !== '') {
              itemRow[7] = item.parsed.heightFromBracketPlus2
            }
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'waterproofing_exterior_side', parsedData: item, section: 'waterproofing', subsectionName: subsection.name })
          })
          exteriorSidePitItems.forEach(item => {
            const slabRow = item.parsed?.heightRefKey ? foundationSlabRows[item.parsed.heightRefKey] : null
            if (slabRow == null) return
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || ''
            rows.push(itemRow)
            formulas.push({
              row: rows.length,
              itemType: 'waterproofing_exterior_side_pit',
              parsedData: item,
              section: 'waterproofing',
              subsectionName: subsection.name,
              foundationSlabRow: slabRow
            })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({
            row: rows.length,
            itemType: 'waterproofing_exterior_side_sum',
            section: 'waterproofing',
            subsectionName: subsection.name,
            firstDataRow: firstItemRow,
            lastDataRow: rows.length - 1
          })
        } else if (subsection.name === 'Negative side' && (negativeSideWallItems.length > 0 || negativeSideSlabItems.length > 0)) {
          const firstItemRow = rows.length + 1
          negativeSideWallItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'FT'
            if (item.parsed?.secondValueFeet != null && item.parsed.secondValueFeet !== '') {
              itemRow[7] = item.parsed.secondValueFeet
            }
            rows.push(itemRow)
            formulas.push({
              row: rows.length,
              itemType: 'waterproofing_negative_side_wall',
              parsedData: item,
              section: 'waterproofing',
              subsectionName: subsection.name
            })
          })
          negativeSideSlabItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'SQ FT'
            rows.push(itemRow)
            formulas.push({
              row: rows.length,
              itemType: 'waterproofing_negative_side_slab',
              parsedData: item,
              section: 'waterproofing',
              subsectionName: subsection.name
            })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({
            row: rows.length,
            itemType: 'waterproofing_negative_side_sum',
            section: 'waterproofing',
            subsectionName: subsection.name,
            firstDataRow: firstItemRow,
            lastDataRow: rows.length - 1
          })
        } else if (subsection.name === 'Horizontal') {
          const firstWPRow = rows.length + 1
          const wpItems = [
            { particulars: 'WP @ SOG', unit: 'SQ FT' },
            { particulars: 'WP @ Grade beam', unit: 'SQ FT' },
            { particulars: 'WP @ pile cap', unit: 'SQ FT' }
          ]
          wpItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = 0
            itemRow[3] = item.unit || 'SQ FT'
            rows.push(itemRow)
            formulas.push({
              row: rows.length,
              itemType: 'waterproofing_horizontal_wp',
              section: 'waterproofing',
              subsectionName: subsection.name
            })
          })
          const wpSumRow = Array(template.columns.length).fill('')
          rows.push(wpSumRow)
          formulas.push({
            row: rows.length,
            itemType: 'waterproofing_horizontal_wp_sum',
            section: 'waterproofing',
            subsectionName: subsection.name,
            firstDataRow: firstWPRow,
            lastDataRow: firstWPRow + 2
          })
          rows.push(Array(template.columns.length).fill(''))
          const firstInsulRow = rows.length + 1
          const insulItems = [
            { particulars: '2"XPS Rigid insulation @ SOG & GB' },
            { particulars: '2"XPS Rigid insulation @ SOG & GB' },
            { particulars: '2"XPS Rigid insulation @ SOG & GB' }
          ]
          insulItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = 0
            itemRow[3] = 'SQ FT'
            rows.push(itemRow)
            formulas.push({
              row: rows.length,
              itemType: 'waterproofing_horizontal_insulation',
              section: 'waterproofing',
              subsectionName: subsection.name
            })
          })
          const insulSumRow = Array(template.columns.length).fill('')
          rows.push(insulSumRow)
          formulas.push({
            row: rows.length,
            itemType: 'waterproofing_horizontal_insulation_sum',
            section: 'waterproofing',
            subsectionName: subsection.name,
            firstDataRow: firstInsulRow,
            lastDataRow: firstInsulRow + 2
          })
        } else if (subsection.name === 'For foundation Extra line item use this') {
          const extraItems = [
            { name: 'In SQ FT', unit: 'SQ FT', type: 'waterproofing_extra_sqft' },
            { name: 'In FT', unit: 'FT', h: 1, type: 'waterproofing_extra_ft' }
          ]
          extraItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.name
            itemRow[2] = 1
            itemRow[3] = item.unit
            if (item.h) itemRow[7] = item.h
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: item.type, section: 'waterproofing' })
          })
        } else {
          rows.push(Array(template.columns.length).fill(''))
        }

        if (hasSubsectionData) {
          rows.push(Array(template.columns.length).fill(''))
        }
      })
    } else if (section.section === 'Superstructure') {
      section.subsections.forEach((subsection) => {
        let hasSubsectionData = false
        // Determine if subsection has data
        if (subsection.name === 'CIP Slabs') hasSubsectionData = superstructureItems.cipSlab8.length > 0 || superstructureItems.cipRoofSlab8.length > 0
        else if (subsection.name === 'Balcony slab') hasSubsectionData = superstructureItems.balconySlab.length > 0
        else if (subsection.name === 'Terrace slab') hasSubsectionData = superstructureItems.terraceSlab.length > 0
        else if (subsection.name === 'Patch slab') hasSubsectionData = superstructureItems.patchSlab.length > 0
        else if (subsection.name === 'Slab steps') hasSubsectionData = superstructureItems.slabSteps.length > 0
        else if (subsection.name === 'LW concrete fill') hasSubsectionData = superstructureItems.lwConcreteFill.length > 0
        else if (subsection.name === 'Slab on metal deck') hasSubsectionData = superstructureItems.slabOnMetalDeck && superstructureItems.slabOnMetalDeck.length > 0
        else if (subsection.name === 'Topping slab') hasSubsectionData = superstructureItems.toppingSlab && superstructureItems.toppingSlab.length > 0
        else if (subsection.name === 'Thermal break') hasSubsectionData = superstructureItems.thermalBreak && superstructureItems.thermalBreak.length > 0
        else if (subsection.name === 'Raised slab') hasSubsectionData = superstructureItems.raisedSlab && (superstructureItems.raisedSlab.kneeWall.length > 0 || superstructureItems.raisedSlab.raisedSlab.length > 0)
        else if (subsection.name === 'Built up slab') hasSubsectionData = superstructureItems.builtUpSlab && (superstructureItems.builtUpSlab.kneeWall.length > 0 || superstructureItems.builtUpSlab.builtUpSlab.length > 0)
        else if (subsection.name === 'Built up stair') hasSubsectionData = superstructureItems.builtUpStair && (superstructureItems.builtUpStair.kneeWall.length > 0 || superstructureItems.builtUpStair.builtUpStairs.length > 0)
        else if (subsection.name === 'Built up ramps') hasSubsectionData = superstructureItems.builtupRamps && (superstructureItems.builtupRamps.kneeWall.length > 0 || superstructureItems.builtupRamps.ramp.length > 0)
        else if (subsection.name === 'Concrete hanger') hasSubsectionData = superstructureItems.concreteHanger && superstructureItems.concreteHanger.length > 0
        else if (subsection.name === 'Shear Walls') hasSubsectionData = superstructureItems.shearWalls && superstructureItems.shearWalls.length > 0
        else if (subsection.name === 'Parapet walls') hasSubsectionData = superstructureItems.parapetWalls && superstructureItems.parapetWalls.length > 0
        else if (subsection.name === 'Columns') hasSubsectionData = superstructureItems.columnsTakeoff && superstructureItems.columnsTakeoff.length > 0
        else if (subsection.name === 'Concrete post') hasSubsectionData = superstructureItems.concretePost && superstructureItems.concretePost.length > 0
        else if (subsection.name === 'Concrete encasement') hasSubsectionData = superstructureItems.concreteEncasement && superstructureItems.concreteEncasement.length > 0
        else if (subsection.name === 'Drop panel') hasSubsectionData = (superstructureItems.dropPanelBracket && superstructureItems.dropPanelBracket.length > 0) || (superstructureItems.dropPanelH && superstructureItems.dropPanelH.length > 0)
        else if (subsection.name === 'Beams') hasSubsectionData = superstructureItems.beams && superstructureItems.beams.length > 0
        else if (subsection.name === 'Curbs') hasSubsectionData = superstructureItems.curbs && superstructureItems.curbs.length > 0
        else if (subsection.name === 'Concrete pad') hasSubsectionData = superstructureItems.concretePad && superstructureItems.concretePad.length > 0
        else if (subsection.name === 'Non-shrink grout') hasSubsectionData = superstructureItems.nonShrinkGrout && superstructureItems.nonShrinkGrout.length > 0
        else if (subsection.name === 'Repair scope') hasSubsectionData = superstructureItems.repairScope && superstructureItems.repairScope.length > 0
        else if (subsection.name === 'CIP Stairs') hasSubsectionData = true // Always show CIP stairs? 
        else if (subsection.name === 'Stairs \u2013 Infilled tads') hasSubsectionData = true // Always show?
        else if (subsection.name === 'Ele') hasSubsectionData = civilOtherItems['Drains & Utilities'] && civilOtherItems['Drains & Utilities'].some(i => i.particulars.toLowerCase().includes('electrical conduit'))
        else if (subsection.name === 'For Superstructure Extra line item use this') hasSubsectionData = true

        if (hasSubsectionData) {
          const subsectionRow = Array(template.columns.length).fill('')
          subsectionRow[1] = subsection.name + ':'
          rows.push(subsectionRow)
        }

        if (subsection.name === 'CIP Slabs') {
          const slab8FirstRow = rows.length + 1
          superstructureItems.cipSlab8.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'SQ FT'
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_item', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const slab8LastRow = rows.length
          if (superstructureItems.cipSlab8.length > 0) {
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: slab8FirstRow, lastDataRow: slab8LastRow })
          }
          rows.push(Array(template.columns.length).fill(''))
          const roofFirstRow = rows.length + 1
          superstructureItems.cipRoofSlab8.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'SQ FT'
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_item', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const roofLastRow = rows.length
          if (superstructureItems.cipRoofSlab8.length > 0) {
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: roofFirstRow, lastDataRow: roofLastRow })
          }
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Balcony slab' && superstructureItems.balconySlab.length > 0) {
          const firstRow = rows.length + 1
          superstructureItems.balconySlab.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'SQ FT'
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_item', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({ row: rows.length, itemType: 'superstructure_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: firstRow, lastDataRow: rows.length - 1 })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Terrace slab' && superstructureItems.terraceSlab.length > 0) {
          const firstRow = rows.length + 1
          superstructureItems.terraceSlab.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'SQ FT'
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_item', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({ row: rows.length, itemType: 'superstructure_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: firstRow, lastDataRow: rows.length - 1 })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Patch slab' && superstructureItems.patchSlab.length > 0) {
          const firstRow = rows.length + 1
          superstructureItems.patchSlab.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'SQ FT'
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_item', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({ row: rows.length, itemType: 'superstructure_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: firstRow, lastDataRow: rows.length - 1 })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Slab steps' && superstructureItems.slabSteps.length > 0) {
          const firstRow = rows.length + 1
          superstructureItems.slabSteps.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'FT'
            if (item.parsed?.widthValue != null) itemRow[6] = item.parsed.widthValue
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_slab_step', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({ row: rows.length, itemType: 'superstructure_slab_steps_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: firstRow, lastDataRow: rows.length - 1 })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'LW concrete fill' && superstructureItems.lwConcreteFill.length > 0) {
          const firstRow = rows.length + 1
          superstructureItems.lwConcreteFill.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'SQ FT'
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_lw_concrete_fill', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({ row: rows.length, itemType: 'superstructure_lw_concrete_fill_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: firstRow, lastDataRow: rows.length - 1 })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Slab on metal deck' && superstructureItems.slabOnMetalDeck && superstructureItems.slabOnMetalDeck.length > 0) {
          const groups = superstructureItems.slabOnMetalDeck
          groups.forEach((group, groupIndex) => {
            const itemFirstRow = rows.length + 1
            group.items.forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'SQ FT'
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'superstructure_somd_item', section: 'superstructure', subsectionName: subsection.name })
            })
            const itemLastRow = rows.length
            rows.push(Array(template.columns.length).fill(''))
            const gen1Row = rows.length + 1
            const gen1RowData = Array(template.columns.length).fill('')
            gen1RowData[1] = group.particulars
            gen1RowData[3] = 'SQ FT'
            rows.push(gen1RowData)
            formulas.push({ row: rows.length, itemType: 'superstructure_somd_gen1', section: 'superstructure', subsectionName: subsection.name, firstDataRow: itemFirstRow, lastDataRow: itemLastRow, heightFormula: `${group.firstValueInches}/12` })
            const gen2Row = rows.length + 1
            const gen2RowData = Array(template.columns.length).fill('')
            gen2RowData[3] = 'SQ FT'
            rows.push(gen2RowData)
            formulas.push({ row: rows.length, itemType: 'superstructure_somd_gen2', section: 'superstructure', subsectionName: subsection.name, takeoffRefRow: gen1Row, heightValue: group.secondValueInches / 12 })
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_somd_sum', section: 'superstructure', subsectionName: subsection.name, gen1Row, gen2Row })
            rows.push(Array(template.columns.length).fill(''))
            if (groupIndex < groups.length - 1) {
              rows.push(Array(template.columns.length).fill(''))
              rows.push(Array(template.columns.length).fill(''))
            }
          })
        } else if (subsection.name === 'Topping slab' && superstructureItems.toppingSlab && superstructureItems.toppingSlab.length > 0) {
          const firstRow = rows.length + 1
          superstructureItems.toppingSlab.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'SQ FT'
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_topping_slab', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({ row: rows.length, itemType: 'superstructure_topping_slab_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: firstRow, lastDataRow: rows.length - 1 })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Thermal break' && superstructureItems.thermalBreak && superstructureItems.thermalBreak.length > 0) {
          const firstRow = rows.length + 1
          superstructureItems.thermalBreak.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'FT'
            if (item.parsed?.qty != null) itemRow[4] = item.parsed.qty
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_thermal_break', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({ row: rows.length, itemType: 'superstructure_thermal_break_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: firstRow, lastDataRow: rows.length - 1 })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Raised slab' && superstructureItems.raisedSlab && (superstructureItems.raisedSlab.kneeWall.length > 0 || superstructureItems.raisedSlab.raisedSlab.length > 0)) {
          const kneeWallFirstRow = rows.length + 1
          superstructureItems.raisedSlab.kneeWall.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'FT'
            if (item.parsed?.widthValue != null) itemRow[6] = item.parsed.widthValue
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_raised_knee_wall', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const hasKneeWall = superstructureItems.raisedSlab.kneeWall.length > 0
          const hasRaisedSlab = superstructureItems.raisedSlab.raisedSlab.length > 0
          if (hasKneeWall && hasRaisedSlab) {
            const raisedSlabFirstRow = rows.length + 2
            const styrofoamRow = Array(template.columns.length).fill('')
            styrofoamRow[1] = 'Styrofoam'
            rows.push(styrofoamRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_raised_styrofoam', section: 'superstructure', subsectionName: subsection.name, takeoffRefRow: raisedSlabFirstRow, heightRefRow: kneeWallFirstRow })
          }
          superstructureItems.raisedSlab.raisedSlab.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'SQ FT'
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_raised_slab', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Built-up slab' && superstructureItems.builtUpSlab && (superstructureItems.builtUpSlab.kneeWall.length > 0 || superstructureItems.builtUpSlab.builtUpSlab.length > 0)) {
          const kneeWallFirstRow = rows.length + 1
          superstructureItems.builtUpSlab.kneeWall.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'FT'
            if (item.parsed?.widthValue != null) itemRow[6] = item.parsed.widthValue
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_builtup_knee_wall', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const hasKneeWall = superstructureItems.builtUpSlab.kneeWall.length > 0
          const hasBuiltUpSlab = superstructureItems.builtUpSlab.builtUpSlab.length > 0
          if (hasKneeWall && hasBuiltUpSlab) {
            const builtUpSlabFirstRow = rows.length + 2
            const styrofoamRow = Array(template.columns.length).fill('')
            styrofoamRow[1] = 'Styrofoam'
            rows.push(styrofoamRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_builtup_styrofoam', section: 'superstructure', subsectionName: subsection.name, takeoffRefRow: builtUpSlabFirstRow, heightRefRow: kneeWallFirstRow })
          }
          superstructureItems.builtUpSlab.builtUpSlab.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'SQ FT'
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_builtup_slab', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Builtup ramps' && superstructureItems.builtupRamps && (superstructureItems.builtupRamps.kneeWall.length > 0 || superstructureItems.builtupRamps.ramp.length > 0)) {
          const kneeWallFirstRow = rows.length + 1
          const kneeWalls = [...superstructureItems.builtupRamps.kneeWall].sort((a, b) => (a.parsed?.groupId ?? 1) - (b.parsed?.groupId ?? 1))
          kneeWalls.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'FT'
            if (item.parsed?.widthValue != null) itemRow[6] = item.parsed.widthValue
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_builtup_ramps_knee_wall', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const kneeLastRow = rows.length
          if (kneeWalls.length > 0) {
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_builtup_ramps_knee_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: kneeWallFirstRow, lastDataRow: kneeLastRow })
          }
          rows.push(Array(template.columns.length).fill(''))
          const ramps = [...superstructureItems.builtupRamps.ramp].sort((a, b) => (a.parsed?.groupId ?? 1) - (b.parsed?.groupId ?? 1))
          const styroFirstRow = rows.length + 1
          const rampFirstRow = rows.length + ramps.length + 3
          ramps.forEach((_, idx) => {
            const styrofoamRow = Array(template.columns.length).fill('')
            const kneeWall = kneeWalls[idx]
            const sizeInches = kneeWall?.parsed?.heightValue != null ? (kneeWall.parsed.heightValue * 12).toFixed(0) : ''
            const groupId = kneeWall?.parsed?.groupId ?? ramps[idx].parsed?.groupId ?? idx + 1
            styrofoamRow[1] = `Styrofoam ${sizeInches}" (${groupId})`
            styrofoamRow[3] = 'SQ FT'
            rows.push(styrofoamRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_builtup_ramps_styrofoam', section: 'superstructure', subsectionName: subsection.name, takeoffRefRow: rampFirstRow + idx, heightRefRow: kneeWallFirstRow + idx })
          })
          const rampFirstRowActual = rows.length + 2
          if (ramps.length > 0) {
            const styroSumRow = Array(template.columns.length).fill('')
            rows.push(styroSumRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_builtup_ramps_styro_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: styroFirstRow, lastDataRow: styroFirstRow + ramps.length - 1 })
          }
          rows.push(Array(template.columns.length).fill(''))
          ramps.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'SQ FT'
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_builtup_ramps_ramp', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const rampLastRow = rows.length
          if (ramps.length > 0) {
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_builtup_ramps_ramp_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: rampFirstRowActual, lastDataRow: rampLastRow })
          }
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Built-up stair' && superstructureItems.builtUpStair && (superstructureItems.builtUpStair.kneeWall.length > 0 || superstructureItems.builtUpStair.builtUpStairs.length > 0)) {
          const kneeWallFirstRow = rows.length + 1
          superstructureItems.builtUpStair.kneeWall.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'FT'
            if (item.parsed?.widthValue != null) itemRow[6] = item.parsed.widthValue
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_builtup_stair_knee_wall', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const styrofoamRow = Array(template.columns.length).fill('')
          styrofoamRow[1] = 'Styrofoam'
          styrofoamRow[3] = 'SQ FT'
          rows.push(styrofoamRow)
          const stairFirstRow = rows.length + 1
          const stairLastRowForJSum = rows.length + 1 + superstructureItems.builtUpStair.builtUpStairs.length
          formulas.push({ row: rows.length, itemType: 'superstructure_builtup_stair_styrofoam', section: 'superstructure', subsectionName: subsection.name, takeoffJSumFirstRow: stairFirstRow, takeoffJSumLastRow: stairLastRowForJSum, heightRefRow: kneeWallFirstRow })
          superstructureItems.builtUpStair.builtUpStairs.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'Treads'
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_builtup_stairs', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const stairSlabDataRow = Array(template.columns.length).fill('')
          stairSlabDataRow[1] = 'Stair slab'
          stairSlabDataRow[3] = 'FT'
          rows.push(stairSlabDataRow)
          formulas.push({ row: rows.length, itemType: 'superstructure_stair_slab', section: 'superstructure', subsectionName: subsection.name, takeoffRefRow: stairFirstRow, widthRefRow: stairFirstRow, heightValue: 0.5 })
          const stairLastRow = rows.length
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({ row: rows.length, itemType: 'superstructure_builtup_stair_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: stairFirstRow, lastDataRow: stairLastRow })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Concrete hanger' && superstructureItems.concreteHanger && superstructureItems.concreteHanger.length > 0) {
          const firstRow = rows.length + 1
          superstructureItems.concreteHanger.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'EA'
            if (item.parsed?.lengthValue != null) itemRow[5] = item.parsed.lengthValue
            if (item.parsed?.widthValue != null) itemRow[6] = item.parsed.widthValue
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_concrete_hanger', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({ row: rows.length, itemType: 'superstructure_concrete_hanger_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: firstRow, lastDataRow: rows.length - 1 })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Shear Walls' && superstructureItems.shearWalls && superstructureItems.shearWalls.length > 0) {
          const shearGroups = {}
          superstructureItems.shearWalls.forEach((item) => {
            const w = item.parsed?.widthValue
            const key = w != null ? String(w) : 'other'
            if (!shearGroups[key]) shearGroups[key] = []
            shearGroups[key].push(item)
          })
          const groupKeys = Object.keys(shearGroups)
          groupKeys.forEach((key, groupIndex) => {
            const groupItems = shearGroups[key]
            const firstRow = rows.length + 1
            groupItems.forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'FT'
              if (item.parsed?.widthValue != null) itemRow[6] = item.parsed.widthValue
              if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'superstructure_shear_wall', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
            })
            const lastRow = rows.length
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_shear_walls_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: firstRow, lastDataRow: lastRow })
            rows.push(Array(template.columns.length).fill(''))
            if (groupIndex < groupKeys.length - 1) {
              rows.push(Array(template.columns.length).fill(''))
            }
          })
        } else if (subsection.name === 'Parapet walls' && superstructureItems.parapetWalls && superstructureItems.parapetWalls.length > 0) {
          const firstRow = rows.length + 1
          superstructureItems.parapetWalls.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'FT'
            if (item.parsed?.qty != null) itemRow[4] = item.parsed.qty
            if (item.parsed?.widthValue != null) itemRow[6] = item.parsed.widthValue
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_parapet_wall', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({ row: rows.length, itemType: 'superstructure_parapet_walls_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: firstRow, lastDataRow: rows.length - 1 })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Columns') {
          const firstItem = superstructureItems.columnsTakeoff?.[0]
          const takeoffRow = rows.length + 1
          const row1 = Array(template.columns.length).fill('')
          row1[1] = 'Columns'
          row1[2] = firstItem?.takeoff || 0
          row1[3] = 'EA'
          rows.push(row1)
          formulas.push({ row: rows.length, itemType: 'superstructure_columns_takeoff', parsedData: firstItem, section: 'superstructure', subsectionName: subsection.name })
          rows.push(Array(template.columns.length).fill(''))
          const row2 = Array(template.columns.length).fill('')
          row2[1] = 'Final as per schedule count'
          row2[3] = 'EA'
          rows.push(row2)
          formulas.push({ row: rows.length, itemType: 'superstructure_columns_final', section: 'superstructure', subsectionName: subsection.name, takeoffRefRow: takeoffRow })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Concrete post' && superstructureItems.concretePost && superstructureItems.concretePost.length > 0) {
          const firstRow = rows.length + 1
          superstructureItems.concretePost.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'EA'
            if (item.parsed?.lengthValue != null) itemRow[5] = item.parsed.lengthValue
            if (item.parsed?.widthValue != null) itemRow[6] = item.parsed.widthValue
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_concrete_post', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({ row: rows.length, itemType: 'superstructure_concrete_post_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: firstRow, lastDataRow: rows.length - 1 })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Concrete encasement' && superstructureItems.concreteEncasement && superstructureItems.concreteEncasement.length > 0) {
          const firstRow = rows.length + 1
          superstructureItems.concreteEncasement.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'EA'
            if (item.parsed?.lengthValue != null) itemRow[5] = item.parsed.lengthValue
            if (item.parsed?.widthValue != null) itemRow[6] = item.parsed.widthValue
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_concrete_encasement', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({ row: rows.length, itemType: 'superstructure_concrete_encasement_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: firstRow, lastDataRow: rows.length - 1 })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Drop panel' && (superstructureItems.dropPanelBracket?.length > 0 || superstructureItems.dropPanelH?.length > 0)) {
          const firstRow = rows.length + 1
            ; (superstructureItems.dropPanelBracket || []).forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'EA'
              if (item.parsed?.lengthValue != null) itemRow[5] = item.parsed.lengthValue
              if (item.parsed?.widthValue != null) itemRow[6] = item.parsed.widthValue
              if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'superstructure_drop_panel_bracket', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
            })
            ; (superstructureItems.dropPanelH || []).forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'SQ FT'
              if (item.parsed?.qty != null) itemRow[4] = item.parsed.qty
              if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'superstructure_drop_panel_h', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
            })
          const lastRow = rows.length
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({ row: rows.length, itemType: 'superstructure_drop_panel_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: firstRow, lastDataRow: lastRow })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Beams' && superstructureItems.beams && superstructureItems.beams.length > 0) {
          const firstRow = rows.length + 1
          superstructureItems.beams.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'FT'
            if (item.parsed?.widthValue != null) itemRow[6] = item.parsed.widthValue
            if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_beam', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({ row: rows.length, itemType: 'superstructure_beams_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: firstRow, lastDataRow: rows.length - 1 })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Curbs' && superstructureItems.curbs && superstructureItems.curbs.length > 0) {
          const curbGroups = {}
          superstructureItems.curbs.forEach((item) => {
            const key = String(item.parsed?.widthValue ?? '')
            if (!curbGroups[key]) curbGroups[key] = []
            curbGroups[key].push(item)
          })
          Object.keys(curbGroups).forEach((key, groupIndex) => {
            const groupItems = curbGroups[key]
            const firstRow = rows.length + 1
            groupItems.forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'FT'
              if (item.parsed?.widthValue != null) itemRow[6] = item.parsed.widthValue
              if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'superstructure_curb', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
            })
            const lastRow = rows.length
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_curbs_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: firstRow, lastDataRow: lastRow })
            rows.push(Array(template.columns.length).fill(''))
            if (groupIndex < Object.keys(curbGroups).length - 1) rows.push(Array(template.columns.length).fill(''))
          })
        } else if (subsection.name === 'Concrete pad' && superstructureItems.concretePad && superstructureItems.concretePad.length > 0) {
          const padGroups = {}
          superstructureItems.concretePad.forEach((item) => {
            const h = item.parsed?.heightValue ?? 0
            const key = String(h)
            if (!padGroups[key]) padGroups[key] = []
            padGroups[key].push(item)
          })
          Object.keys(padGroups).forEach((key, groupIndex) => {
            const groupItems = padGroups[key]
            const firstRow = rows.length + 1
            groupItems.forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'SQ FT'
              if (item.parsed?.qty != null) itemRow[4] = item.parsed.qty
              // Set height (H column) for SQ FT items
              const unitLower = String(item.unit || '').toLowerCase().trim()
              if ((unitLower.includes('sq') || unitLower === 'sf') && item.parsed?.heightValue != null) {
                itemRow[7] = item.parsed.heightValue
              }
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'superstructure_concrete_pad', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
            })
            const lastRow = rows.length
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_concrete_pad_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: firstRow, lastDataRow: lastRow })
            rows.push(Array(template.columns.length).fill(''))
            if (groupIndex < Object.keys(padGroups).length - 1) rows.push(Array(template.columns.length).fill(''))
          })
        } else if (subsection.name === 'Non-shrink grout' && superstructureItems.nonShrinkGrout && superstructureItems.nonShrinkGrout.length > 0) {
          const firstRow = rows.length + 1
          superstructureItems.nonShrinkGrout.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'EA'
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_non_shrink_grout', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          const sumRow = Array(template.columns.length).fill('')
          rows.push(sumRow)
          formulas.push({ row: rows.length, itemType: 'superstructure_non_shrink_grout_sum', section: 'superstructure', subsectionName: subsection.name, firstDataRow: firstRow, lastDataRow: rows.length - 1 })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'Repair scope' && superstructureItems.repairScope && superstructureItems.repairScope.length > 0) {
          superstructureItems.repairScope.forEach((item) => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.particulars
            itemRow[2] = item.takeoff
            itemRow[3] = item.unit || 'FT'
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: 'superstructure_repair_scope', parsedData: item, section: 'superstructure', subsectionName: subsection.name })
          })
          rows.push(Array(template.columns.length).fill(''))
        } else if (subsection.name === 'CIP Stairs') {
          // CIP Stairs - formulas in columns I-M; raw values in C/F/G/H
          const stairGroups = [
            {
              name: 'Stair A, B & D:',
              hasLandings: true,
              landingTakeoff: 21,
              landingHeight: 0.67,
              stairTakeoff: 262,
              stairF: '11/12',
              stairG: 3,
              stairHeight: '7/12',
              slabHeight: 0.5,
              slabG: 3,
              slabCMultiplier: 1.3
            },
            {
              name: 'Stair E & F:',
              hasLandings: false,
              stairTakeoff: '=18+18',
              stairF: '11/12',
              stairG: 4,
              stairHeight: '7/12',
              slabHeight: 0.5,
              slabG: 4,
              slabCMultiplier: 1.3
            },
            {
              name: 'Misc. stair:',
              hasLandings: false,
              stairTakeoff: 16,
              stairF: '11/12',
              stairG: 3,
              stairHeight: '7/12',
              slabHeight: 0.5,
              slabG: 3,
              slabCMultiplier: 1.3
            },
            {
              name: 'Ext. stair:',
              hasLandings: true,
              landingTakeoff: 20.1,
              landingHeight: 0.67,
              stairTakeoff: 16,
              stairF: '11/12',
              stairG: 3,
              stairHeight: '7/12',
              slabHeight: 0.5,
              slabG: 3,
              slabCMultiplier: 1.3
            }
          ]

          stairGroups.forEach((group) => {
            // Group Header (underlined, no other data in this row)
            const headerRow = Array(template.columns.length).fill('')
            headerRow[1] = group.name
            rows.push(headerRow)
            formulas.push({
              row: rows.length,
              itemType: 'superstructure_manual_cip_stairs_header',
              section: 'superstructure',
              subsectionName: subsection.name
            })

            // Landings (if present)
            if (group.hasLandings) {
              const landingRow = Array(template.columns.length).fill('')
              landingRow[1] = 'Landings'
              rows.push(landingRow)
              formulas.push({
                row: rows.length,
                itemType: 'superstructure_manual_cip_stairs_landing',
                section: 'superstructure',
                subsectionName: subsection.name,
                takeoff: group.landingTakeoff,
                height: group.landingHeight
              })

              // Gap row after landings
              rows.push(Array(template.columns.length).fill(''))
            }

            // Stairs
            const stairRow = Array(template.columns.length).fill('')
            stairRow[1] = 'Stairs'
            rows.push(stairRow)
            const stairRowNum = rows.length
            formulas.push({
              row: rows.length,
              itemType: 'superstructure_manual_cip_stairs_stair',
              section: 'superstructure',
              subsectionName: subsection.name,
              takeoff: group.stairTakeoff,
              stairF: group.stairF,
              stairG: group.stairG,
              height: group.stairHeight
            })

            // Stair slab
            const slabRow = Array(template.columns.length).fill('')
            slabRow[1] = 'Stair slab'
            rows.push(slabRow)
            formulas.push({
              row: rows.length,
              itemType: 'superstructure_manual_cip_stairs_slab',
              section: 'superstructure',
              subsectionName: subsection.name,
              height: group.slabHeight,
              slabG: group.slabG,
              stairsRowNum: stairRowNum,
              slabCMultiplier: group.slabCMultiplier,
              hasLandings: group.hasLandings
            })

            // Sum row - sums only Stairs and Stair slab (not Landings)
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({
              row: rows.length,
              itemType: 'superstructure_manual_cip_stairs_sum',
              section: 'superstructure',
              subsectionName: subsection.name,
              firstDataRow: stairRowNum,
              lastDataRow: stairRowNum + 1
            })

            // Blank row
            rows.push(Array(template.columns.length).fill(''))
          })

        } else if (subsection.name === 'Stairs \u2013 Infilled tads') {
          // Stairs - Infilled tads - similar to CIP Stairs, C-H empty
          // Two identical "Stair E:" groups

          const addStairE = () => {
            // Header (underlined, no other data)
            const headerRow = Array(template.columns.length).fill('')
            headerRow[1] = 'Stair E:'
            rows.push(headerRow)
            formulas.push({
              row: rows.length,
              itemType: 'superstructure_manual_infilled_header',
              section: 'superstructure',
              subsectionName: subsection.name
            })

            // Landings row 1
            const landing1Row = Array(template.columns.length).fill('')
            landing1Row[1] = 'Landings'
            rows.push(landing1Row)
            const landing1RowNum = rows.length
            formulas.push({
              row: rows.length,
              itemType: 'superstructure_manual_infilled_landing_1',
              section: 'superstructure',
              subsectionName: subsection.name
            })

            // Landings row 2 (no name, C and H set via formulas)
            const landing2Row = Array(template.columns.length).fill('')
            rows.push(landing2Row)
            formulas.push({
              row: rows.length,
              itemType: 'superstructure_manual_infilled_landing_2',
              section: 'superstructure',
              subsectionName: subsection.name,
              landing1RowNum: landing1RowNum
            })

            // Sum for Landings
            const landingSumRow = Array(template.columns.length).fill('')
            rows.push(landingSumRow)
            formulas.push({
              row: rows.length,
              itemType: 'superstructure_manual_infilled_landing_sum',
              section: 'superstructure',
              subsectionName: subsection.name,
              landing1RowNum: landing1RowNum,
              firstDataRow: landing1RowNum,
              lastDataRow: landing1RowNum + 1
            })

            // Blank row
            rows.push(Array(template.columns.length).fill(''))

            // Stairs
            const stairsRow = Array(template.columns.length).fill('')
            stairsRow[1] = 'Stairs'
            rows.push(stairsRow)
            formulas.push({
              row: rows.length,
              itemType: 'superstructure_manual_infilled_stair',
              section: 'superstructure',
              subsectionName: subsection.name
            })

            // Blank row
            rows.push(Array(template.columns.length).fill(''))
          }

          // Add two identical Stair E groups
          addStairE()
          addStairE()

        } else if (subsection.name === 'Ele') {
          // Superstructure Ele subsection with Excavation, Backfill, Gravel
          const createEleGroup = (subSubName, takeoffSource, lMultiplier = 1) => {
            // Sub-subsection header
            const headerRow = Array(template.columns.length).fill('')
            headerRow[1] = `${subSubName}:`
            rows.push(headerRow)

            // Data row
            const dataRowIndex = rows.length + 1
            const dataRow = Array(template.columns.length).fill('')
            dataRow[0] = 'Ele'
            dataRow[1] = subSubName
            dataRow[3] = 'FT'
            rows.push(dataRow)
            formulas.push({
              row: rows.length,
              itemType: 'superstructure_ele_item',
              section: 'superstructure',
              subsectionName: subsection.name,
              subSubsectionName: subSubName,
              takeoffSource
            })

            // Sum row (for J and L)
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({
              row: rows.length,
              itemType: 'superstructure_ele_sum',
              section: 'superstructure',
              subsectionName: subsection.name,
              subSubsectionName: subSubName,
              firstDataRow: dataRowIndex,
              lastDataRow: dataRowIndex,
              lMultiplier
            })

            // Blank row between groups
            rows.push(Array(template.columns.length).fill(''))
          }

          // Excavation - takeoff from C of Proposed underground electrical conduit item
          createEleGroup('Excavation', 'drains_conduit', 1.25)
          // Backfill - takeoff from C of Excavation item in Drains & Utilities
          createEleGroup('Backfill', 'drains_excavation', 1)
          // Gravel - takeoff from C of Proposed underground electrical conduit item
          createEleGroup('Gravel', 'drains_conduit', 1)
        } else if (subsection.name === 'For Superstructure Extra line item use this') {
          const extraItems = [
            { name: 'In SQ FT', unit: 'SQ FT', h: 1, type: 'superstructure_extra_sqft' },
            { name: 'In FT', unit: 'FT', g: 1, h: 1, type: 'superstructure_extra_ft' },
            { name: 'In EA', unit: 'EA', f: 1, g: 1, h: 1, type: 'superstructure_extra_ea' }
          ]
          extraItems.forEach(item => {
            const itemRow = Array(template.columns.length).fill('')
            itemRow[1] = item.name
            itemRow[2] = 1
            itemRow[3] = item.unit
            if (item.f) itemRow[5] = item.f
            if (item.g) itemRow[6] = item.g
            if (item.h) itemRow[7] = item.h
            rows.push(itemRow)
            formulas.push({ row: rows.length, itemType: item.type, section: 'superstructure', subsectionName: subsection.name })
          })
          rows.push(Array(template.columns.length).fill(''))
        } else if (hasSubsectionData) {
          rows.push(Array(template.columns.length).fill(''))
        }
      })
    } else if (section.section === 'B.P.P. Alternate #2 scope') {
      // B.P.P. Alternate #2 scope - organized by street name
      const streetNames = Object.keys(bppAlternateItemsByStreet)

      if (streetNames.length > 0) {
        streetNames.forEach((streetName, streetIndex) => {
          const streetData = bppAlternateItemsByStreet[streetName]

          // Street name header row
          const streetHeaderRow = Array(template.columns.length).fill('')
          streetHeaderRow[1] = `Street name: ${streetName}`
          rows.push(streetHeaderRow)
          formulas.push({ row: rows.length, itemType: 'bpp_street_header', section: 'bpp_alternate', streetName })

          // Gravel rows (2 rows - 4" and 6" gravel, manual entry, col H empty)
          const gravel4Row = Array(template.columns.length).fill('')
          gravel4Row[1] = 'Gravel'
          gravel4Row[3] = 'SQ FT'
          // col H (height) is left empty for manual entry
          rows.push(gravel4Row)
          formulas.push({ row: rows.length, itemType: 'bpp_gravel', section: 'bpp_alternate', streetName, gravelType: '4inch' })

          const gravel6Row = Array(template.columns.length).fill('')
          gravel6Row[1] = 'Gravel'
          gravel6Row[3] = 'SQ FT'
          // col H (height) is left empty for manual entry
          rows.push(gravel6Row)
          formulas.push({ row: rows.length, itemType: 'bpp_gravel', section: 'bpp_alternate', streetName, gravelType: '6inch' })

          rows.push(Array(template.columns.length).fill(''))

          // Concrete sidewalk items
          if (streetData['Concrete sidewalk'] && streetData['Concrete sidewalk'].length > 0) {
            const firstRow = rows.length + 1
            streetData['Concrete sidewalk'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'SQ FT'
              if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'bpp_concrete_sidewalk', parsedData: item, section: 'bpp_alternate', subsectionName: 'Concrete sidewalk', streetName })
            })
            // Sum row
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'bpp_sum', section: 'bpp_alternate', subsectionName: 'Concrete sidewalk', firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['J', 'L'], streetName })
            rows.push(Array(template.columns.length).fill(''))
          }

          // Concrete driveway items
          if (streetData['Concrete driveway'] && streetData['Concrete driveway'].length > 0) {
            const firstRow = rows.length + 1
            streetData['Concrete driveway'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'SQ FT'
              if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'bpp_concrete_driveway', parsedData: item, section: 'bpp_alternate', subsectionName: 'Concrete driveway', streetName })
            })
            // Sum row
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'bpp_sum', section: 'bpp_alternate', subsectionName: 'Concrete driveway', firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['J', 'L'], streetName })
            rows.push(Array(template.columns.length).fill(''))
          }

          // Concrete curb items
          if (streetData['Concrete curb'] && streetData['Concrete curb'].length > 0) {
            const firstRow = rows.length + 1
            streetData['Concrete curb'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'FT'
              if (item.parsed?.widthValue != null) itemRow[6] = item.parsed.widthValue
              if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'bpp_concrete_curb', parsedData: item, section: 'bpp_alternate', subsectionName: 'Concrete curb', streetName })
            })
            // Sum row
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'bpp_sum', section: 'bpp_alternate', subsectionName: 'Concrete curb', firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['I', 'J', 'L'], streetName })
            rows.push(Array(template.columns.length).fill(''))
          }

          // Concrete flush curb items
          if (streetData['Concrete flush curb'] && streetData['Concrete flush curb'].length > 0) {
            const firstRow = rows.length + 1
            streetData['Concrete flush curb'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'FT'
              if (item.parsed?.widthValue != null) itemRow[6] = item.parsed.widthValue
              if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'bpp_concrete_flush_curb', parsedData: item, section: 'bpp_alternate', subsectionName: 'Concrete flush curb', streetName })
            })
            // Sum row
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'bpp_sum', section: 'bpp_alternate', subsectionName: 'Concrete flush curb', firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['I', 'J', 'L'], streetName })
            rows.push(Array(template.columns.length).fill(''))
          }

          // Expansion joint items
          if (streetData['Expansion joint'] && streetData['Expansion joint'].length > 0) {
            const firstRow = rows.length + 1
            streetData['Expansion joint'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'FT'
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'bpp_expansion_joint', parsedData: item, section: 'bpp_alternate', subsectionName: 'Expansion joint', streetName })
            })
            // Sum row
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'bpp_sum', section: 'bpp_alternate', subsectionName: 'Expansion joint', firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['I'], streetName })
            rows.push(Array(template.columns.length).fill(''))
          }

          // Conc road base row (manual entry, col H empty)
          const concRoadBaseRow = Array(template.columns.length).fill('')
          concRoadBaseRow[1] = 'Conc road base'
          concRoadBaseRow[3] = 'SQ FT'
          // col H (height) is left empty for manual entry
          rows.push(concRoadBaseRow)
          formulas.push({ row: rows.length, itemType: 'bpp_conc_road_base', section: 'bpp_alternate', streetName })

          // Full depth asphalt pavement items
          if (streetData['Full depth asphalt pavement'] && streetData['Full depth asphalt pavement'].length > 0) {
            const firstRow = rows.length + 1
            streetData['Full depth asphalt pavement'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'SQ FT'
              if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'bpp_full_depth_asphalt', parsedData: item, section: 'bpp_alternate', subsectionName: 'Full depth asphalt pavement', streetName })
            })
            // Sum row
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'bpp_sum', section: 'bpp_alternate', subsectionName: 'Full depth asphalt pavement', firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['J', 'L'], streetName })
            rows.push(Array(template.columns.length).fill(''))
          }

          // Add spacing between streets
          if (streetIndex < streetNames.length - 1) {
            rows.push(Array(template.columns.length).fill(''))
          }
        })
      }
    } else if (section.section === 'Civil / Sitework') {
      // Civil / Sitework section
      section.subsections.forEach((subsection) => {
        // Check if subsection has items (Civil)
        let hasSubsectionData = false
        if (subsection.name === 'Demo') hasSubsectionData = Object.values(civilDemoItems).some(val => Array.isArray(val) ? val.length > 0 : Object.values(val).some(v => Array.isArray(v) && v.length > 0))
        else if (subsection.name === 'Excavation') hasSubsectionData = Object.values(civilOtherItems['Excavation']).some(arr => arr.length > 0)
        else if (subsection.name === 'Gravel') hasSubsectionData = Object.values(civilOtherItems['Gravel']).some(arr => arr.length > 0)
        else if (subsection.name === 'Concrete Pavement') hasSubsectionData = civilOtherItems['Concrete Pavement'].length > 0
        else if (subsection.name === 'Asphalt') hasSubsectionData = civilOtherItems['Asphalt'].length > 0
        else if (subsection.name === 'Pads') hasSubsectionData = civilOtherItems['Pads'].length > 0
        else if (subsection.name === 'Soil Erosion') hasSubsectionData = Object.values(civilOtherItems['Soil Erosion']).some(arr => arr.length > 0)
        else if (subsection.name === 'Fence') hasSubsectionData = Object.values(civilOtherItems['Fence']).some(arr => arr.length > 0)
        else if (subsection.name === 'Concrete filled steel pipe bollard') hasSubsectionData = Object.values(civilOtherItems['Concrete filled steel pipe bollard']).some(arr => arr.length > 0)
        else if (subsection.name === 'Site') hasSubsectionData = Object.values(civilOtherItems['Site']).some(val => Array.isArray(val) ? val.length > 0 : Object.values(val).some(v => Array.isArray(v) && v.length > 0))
        else if (subsection.name === 'Alternate') hasSubsectionData = true // Handled inside

        // Add subsection header for all subsections except Alternate,
        // which has custom header/spacing handled in its own block below.
        if (subsection.name !== 'Alternate' && hasSubsectionData) {
          const subsectionRow = Array(template.columns.length).fill('')
          subsectionRow[1] = subsection.name + ':'
          rows.push(subsectionRow)
        }

        if (subsection.name === 'Demo' && subsection.subSubsections) {
          // Demo subsection with sub-subsections
          subsection.subSubsections.forEach((subSubsection) => {
            const subSubName = subSubsection.name

            // Determine if sub-subsection has data
            let hasSubData = false
            if (subSubName === 'Demo asphalt') hasSubData = civilDemoItems['Demo asphalt'].length > 0
            else if (subSubName === 'Demo curb') hasSubData = civilDemoItems['Demo curb'].length > 0
            else if (subSubName === 'Demo fence') hasSubData = civilDemoItems['Demo fence']['chain_link_vinyl'].length > 0 || civilDemoItems['Demo fence']['wood'].length > 0 || civilDemoItems['Demo fence']['other'].length > 0
            else if (subSubName === 'Demo wall') hasSubData = civilDemoItems['Demo wall'].length > 0
            else if (subSubName === 'Demo pipe') hasSubData = civilDemoItems['Demo pipe']['remove_pipe'].length > 0 || civilDemoItems['Demo pipe']['protect'].length > 0
            else if (subSubName === 'Demo rail') hasSubData = civilDemoItems['Demo rail'].length > 0
            else if (subSubName === 'Demo sign') hasSubData = civilDemoItems['Demo sign']['single_sign'].length > 0 || civilDemoItems['Demo sign']['row_of_signs'].length > 0
            else if (subSubName === 'Demo manhole') hasSubData = civilDemoItems['Demo manhole'].length > 0
            else if (subSubName === 'Demo fire hydrant') hasSubData = civilDemoItems['Demo fire hydrant'].length > 0
            else if (subSubName === 'Demo utility pole') hasSubData = civilDemoItems['Demo utility pole'].length > 0
            else if (subSubName === 'Demo valve') hasSubData = civilDemoItems['Demo valve'].length > 0
            else if (subSubName === 'Demo inlet') hasSubData = civilDemoItems['Demo inlet']['protect'].length > 0 || civilDemoItems['Demo inlet']['remove'].length > 0

            if (hasSubData) {
              // Add sub-subsection header
              const subSubsectionRow = Array(template.columns.length).fill('')
              subSubsectionRow[1] = subSubName + ':'
              rows.push(subSubsectionRow)
            }

            // Handle Demo asphalt
            if (subSubName === 'Demo asphalt' && civilDemoItems['Demo asphalt'].length > 0) {
              const firstRow = rows.length + 1
              civilDemoItems['Demo asphalt'].forEach((item) => {
                const itemRow = Array(template.columns.length).fill('')
                itemRow[1] = item.particulars
                itemRow[2] = item.takeoff
                itemRow[3] = item.unit || 'SQ FT'
                if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
                rows.push(itemRow)
                formulas.push({ row: rows.length, itemType: 'civil_demo_asphalt', parsedData: item, section: 'civil_sitework', subsectionName: subSubName })
              })
              // Sum row
              const sumRow = Array(template.columns.length).fill('')
              rows.push(sumRow)
              formulas.push({ row: rows.length, itemType: 'civil_demo_sum', section: 'civil_sitework', subsectionName: subSubName, firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['J', 'L'] })
              rows.push(Array(template.columns.length).fill(''))
            }
            // Handle Demo curb
            else if (subSubName === 'Demo curb' && civilDemoItems['Demo curb'].length > 0) {
              const firstRow = rows.length + 1
              civilDemoItems['Demo curb'].forEach((item) => {
                const itemRow = Array(template.columns.length).fill('')
                itemRow[1] = item.particulars
                itemRow[2] = item.takeoff
                itemRow[3] = item.unit || 'FT'
                if (item.parsed?.widthValue != null) itemRow[6] = item.parsed.widthValue
                if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
                rows.push(itemRow)
                formulas.push({ row: rows.length, itemType: 'civil_demo_curb', parsedData: item, section: 'civil_sitework', subsectionName: subSubName })
              })
              // Sum row
              const sumRow = Array(template.columns.length).fill('')
              rows.push(sumRow)
              formulas.push({ row: rows.length, itemType: 'civil_demo_sum', section: 'civil_sitework', subsectionName: subSubName, firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['I', 'J', 'L'] })
              rows.push(Array(template.columns.length).fill(''))
            }
            // Handle Demo fence (grouped by type)
            else if (subSubName === 'Demo fence') {
              const fenceGroups = civilDemoItems['Demo fence']
              // Chain link / Vinyl fence group
              if (fenceGroups['chain_link_vinyl'].length > 0) {
                const firstRow = rows.length + 1
                fenceGroups['chain_link_vinyl'].forEach((item) => {
                  const itemRow = Array(template.columns.length).fill('')
                  itemRow[1] = item.particulars
                  itemRow[2] = item.takeoff
                  itemRow[3] = item.unit || 'FT'
                  if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
                  rows.push(itemRow)
                  formulas.push({ row: rows.length, itemType: 'civil_demo_fence', parsedData: item, section: 'civil_sitework', subsectionName: subSubName })
                })
                // Sum row
                const sumRow = Array(template.columns.length).fill('')
                rows.push(sumRow)
                formulas.push({ row: rows.length, itemType: 'civil_demo_sum', section: 'civil_sitework', subsectionName: subSubName, firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['I', 'J'] })
                rows.push(Array(template.columns.length).fill(''))
              }
              // Wood fence group
              if (fenceGroups['wood'].length > 0) {
                const firstRow = rows.length + 1
                fenceGroups['wood'].forEach((item) => {
                  const itemRow = Array(template.columns.length).fill('')
                  itemRow[1] = item.particulars
                  itemRow[2] = item.takeoff
                  itemRow[3] = item.unit || 'FT'
                  if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
                  rows.push(itemRow)
                  formulas.push({ row: rows.length, itemType: 'civil_demo_fence', parsedData: item, section: 'civil_sitework', subsectionName: subSubName })
                })
                // Sum row
                const sumRow = Array(template.columns.length).fill('')
                rows.push(sumRow)
                formulas.push({ row: rows.length, itemType: 'civil_demo_sum', section: 'civil_sitework', subsectionName: subSubName, firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['I', 'J'] })
                rows.push(Array(template.columns.length).fill(''))
              }
              // Other fence group
              if (fenceGroups['other'].length > 0) {
                const firstRow = rows.length + 1
                fenceGroups['other'].forEach((item) => {
                  const itemRow = Array(template.columns.length).fill('')
                  itemRow[1] = item.particulars
                  itemRow[2] = item.takeoff
                  itemRow[3] = item.unit || 'FT'
                  if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
                  rows.push(itemRow)
                  formulas.push({ row: rows.length, itemType: 'civil_demo_fence', parsedData: item, section: 'civil_sitework', subsectionName: subSubName })
                })
                // Sum row
                const sumRow = Array(template.columns.length).fill('')
                rows.push(sumRow)
                formulas.push({ row: rows.length, itemType: 'civil_demo_sum', section: 'civil_sitework', subsectionName: subSubName, firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['I', 'J'] })
                rows.push(Array(template.columns.length).fill(''))
              }
            }
            // Handle Demo wall
            else if (subSubName === 'Demo wall' && civilDemoItems['Demo wall'].length > 0) {
              const firstRow = rows.length + 1
              civilDemoItems['Demo wall'].forEach((item) => {
                const itemRow = Array(template.columns.length).fill('')
                itemRow[1] = item.particulars
                itemRow[2] = item.takeoff
                itemRow[3] = item.unit || 'FT'
                if (item.parsed?.widthValue != null) itemRow[6] = item.parsed.widthValue
                if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
                rows.push(itemRow)
                formulas.push({ row: rows.length, itemType: 'civil_demo_wall', parsedData: item, section: 'civil_sitework', subsectionName: subSubName })
              })
              // Sum row
              const sumRow = Array(template.columns.length).fill('')
              rows.push(sumRow)
              formulas.push({ row: rows.length, itemType: 'civil_demo_sum', section: 'civil_sitework', subsectionName: subSubName, firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['I', 'J', 'L'] })
              rows.push(Array(template.columns.length).fill(''))
            }
            // Handle Demo pipe (grouped by type)
            else if (subSubName === 'Demo pipe') {
              const pipeGroups = civilDemoItems['Demo pipe']
              // Remove pipe group
              if (pipeGroups['remove_pipe'].length > 0) {
                const firstRow = rows.length + 1
                pipeGroups['remove_pipe'].forEach((item) => {
                  const itemRow = Array(template.columns.length).fill('')
                  itemRow[1] = item.particulars
                  itemRow[2] = item.takeoff
                  itemRow[3] = item.unit || 'FT'
                  rows.push(itemRow)
                  formulas.push({ row: rows.length, itemType: 'civil_demo_pipe', parsedData: item, section: 'civil_sitework', subsectionName: subSubName })
                })
                // Sum row
                const sumRow = Array(template.columns.length).fill('')
                rows.push(sumRow)
                formulas.push({ row: rows.length, itemType: 'civil_demo_sum', section: 'civil_sitework', subsectionName: subSubName, firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['I'] })
                rows.push(Array(template.columns.length).fill(''))
              }
              // Protect pipe group
              if (pipeGroups['protect'].length > 0) {
                const firstRow = rows.length + 1
                pipeGroups['protect'].forEach((item) => {
                  const itemRow = Array(template.columns.length).fill('')
                  itemRow[1] = item.particulars
                  itemRow[2] = item.takeoff
                  itemRow[3] = item.unit || 'FT'
                  rows.push(itemRow)
                  formulas.push({ row: rows.length, itemType: 'civil_demo_pipe', parsedData: item, section: 'civil_sitework', subsectionName: subSubName })
                })
                // Sum row
                const sumRow = Array(template.columns.length).fill('')
                rows.push(sumRow)
                formulas.push({ row: rows.length, itemType: 'civil_demo_sum', section: 'civil_sitework', subsectionName: subSubName, firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['I'] })
                rows.push(Array(template.columns.length).fill(''))
              }
            }
            // Handle Demo rail
            else if (subSubName === 'Demo rail' && civilDemoItems['Demo rail'].length > 0) {
              const firstRow = rows.length + 1
              civilDemoItems['Demo rail'].forEach((item) => {
                const itemRow = Array(template.columns.length).fill('')
                itemRow[1] = item.particulars
                itemRow[2] = item.takeoff
                itemRow[3] = item.unit || 'FT'
                rows.push(itemRow)
                formulas.push({ row: rows.length, itemType: 'civil_demo_rail', parsedData: item, section: 'civil_sitework', subsectionName: subSubName })
              })
              // Sum row
              const sumRow = Array(template.columns.length).fill('')
              rows.push(sumRow)
              formulas.push({ row: rows.length, itemType: 'civil_demo_sum', section: 'civil_sitework', subsectionName: subSubName, firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['I'] })
              rows.push(Array(template.columns.length).fill(''))
            }
            // Handle Demo sign (grouped by type)
            else if (subSubName === 'Demo sign') {
              const signGroups = civilDemoItems['Demo sign']
              // Single sign group
              if (signGroups['single_sign'].length > 0) {
                const firstRow = rows.length + 1
                signGroups['single_sign'].forEach((item) => {
                  const itemRow = Array(template.columns.length).fill('')
                  itemRow[1] = item.particulars
                  itemRow[2] = item.takeoff
                  itemRow[3] = item.unit || 'EA'
                  rows.push(itemRow)
                  formulas.push({ row: rows.length, itemType: 'civil_demo_ea', parsedData: item, section: 'civil_sitework', subsectionName: subSubName })
                })
                // Sum row
                const sumRow = Array(template.columns.length).fill('')
                rows.push(sumRow)
                formulas.push({ row: rows.length, itemType: 'civil_demo_sum', section: 'civil_sitework', subsectionName: subSubName, firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['M'] })
                rows.push(Array(template.columns.length).fill(''))
              }
              // Row of signs group
              if (signGroups['row_of_signs'].length > 0) {
                const firstRow = rows.length + 1
                signGroups['row_of_signs'].forEach((item) => {
                  const itemRow = Array(template.columns.length).fill('')
                  itemRow[1] = item.particulars
                  itemRow[2] = item.takeoff
                  itemRow[3] = item.unit || 'EA'
                  rows.push(itemRow)
                  formulas.push({ row: rows.length, itemType: 'civil_demo_ea', parsedData: item, section: 'civil_sitework', subsectionName: subSubName })
                })
                // Sum row
                const sumRow = Array(template.columns.length).fill('')
                rows.push(sumRow)
                formulas.push({ row: rows.length, itemType: 'civil_demo_sum', section: 'civil_sitework', subsectionName: subSubName, firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['M'] })
                rows.push(Array(template.columns.length).fill(''))
              }
            }
            // Handle Demo manhole
            else if (subSubName === 'Demo manhole' && civilDemoItems['Demo manhole'].length > 0) {
              const firstRow = rows.length + 1
              civilDemoItems['Demo manhole'].forEach((item) => {
                const itemRow = Array(template.columns.length).fill('')
                itemRow[1] = item.particulars
                itemRow[2] = item.takeoff
                itemRow[3] = item.unit || 'EA'
                rows.push(itemRow)
                formulas.push({ row: rows.length, itemType: 'civil_demo_ea', parsedData: item, section: 'civil_sitework', subsectionName: subSubName })
              })
              // Sum row
              const sumRow = Array(template.columns.length).fill('')
              rows.push(sumRow)
              formulas.push({ row: rows.length, itemType: 'civil_demo_sum', section: 'civil_sitework', subsectionName: subSubName, firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['M'] })
              rows.push(Array(template.columns.length).fill(''))
            }
            // Handle Demo fire hydrant
            else if (subSubName === 'Demo fire hydrant' && civilDemoItems['Demo fire hydrant'].length > 0) {
              const firstRow = rows.length + 1
              civilDemoItems['Demo fire hydrant'].forEach((item) => {
                const itemRow = Array(template.columns.length).fill('')
                itemRow[1] = item.particulars
                itemRow[2] = item.takeoff
                itemRow[3] = item.unit || 'EA'
                rows.push(itemRow)
                formulas.push({ row: rows.length, itemType: 'civil_demo_ea', parsedData: item, section: 'civil_sitework', subsectionName: subSubName })
              })
              // Sum row
              const sumRow = Array(template.columns.length).fill('')
              rows.push(sumRow)
              formulas.push({ row: rows.length, itemType: 'civil_demo_sum', section: 'civil_sitework', subsectionName: subSubName, firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['M'] })
              rows.push(Array(template.columns.length).fill(''))
            }
            // Handle Demo utility pole
            else if (subSubName === 'Demo utility pole' && civilDemoItems['Demo utility pole'].length > 0) {
              const firstRow = rows.length + 1
              civilDemoItems['Demo utility pole'].forEach((item) => {
                const itemRow = Array(template.columns.length).fill('')
                itemRow[1] = item.particulars
                itemRow[2] = item.takeoff
                itemRow[3] = item.unit || 'EA'
                rows.push(itemRow)
                formulas.push({ row: rows.length, itemType: 'civil_demo_ea', parsedData: item, section: 'civil_sitework', subsectionName: subSubName })
              })
              // Sum row
              const sumRow = Array(template.columns.length).fill('')
              rows.push(sumRow)
              formulas.push({ row: rows.length, itemType: 'civil_demo_sum', section: 'civil_sitework', subsectionName: subSubName, firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['M'] })
              rows.push(Array(template.columns.length).fill(''))
            }
            // Handle Demo valve
            else if (subSubName === 'Demo valve' && civilDemoItems['Demo valve'].length > 0) {
              const firstRow = rows.length + 1
              civilDemoItems['Demo valve'].forEach((item) => {
                const itemRow = Array(template.columns.length).fill('')
                itemRow[1] = item.particulars
                itemRow[2] = item.takeoff
                itemRow[3] = item.unit || 'EA'
                rows.push(itemRow)
                formulas.push({ row: rows.length, itemType: 'civil_demo_ea', parsedData: item, section: 'civil_sitework', subsectionName: subSubName })
              })
              // Sum row
              const sumRow = Array(template.columns.length).fill('')
              rows.push(sumRow)
              formulas.push({ row: rows.length, itemType: 'civil_demo_sum', section: 'civil_sitework', subsectionName: subSubName, firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['M'] })
              rows.push(Array(template.columns.length).fill(''))
            }
            // Handle Demo inlet (grouped by type)
            else if (subSubName === 'Demo inlet') {
              const inletGroups = civilDemoItems['Demo inlet']
              // Protect inlet group
              if (inletGroups['protect'].length > 0) {
                const firstRow = rows.length + 1
                inletGroups['protect'].forEach((item) => {
                  const itemRow = Array(template.columns.length).fill('')
                  itemRow[1] = item.particulars
                  itemRow[2] = item.takeoff
                  itemRow[3] = item.unit || 'EA'
                  rows.push(itemRow)
                  formulas.push({ row: rows.length, itemType: 'civil_demo_ea', parsedData: item, section: 'civil_sitework', subsectionName: subSubName })
                })
                // Sum row
                const sumRow = Array(template.columns.length).fill('')
                rows.push(sumRow)
                formulas.push({ row: rows.length, itemType: 'civil_demo_sum', section: 'civil_sitework', subsectionName: subSubName, firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['M'] })
                rows.push(Array(template.columns.length).fill(''))
              }
              // Remove inlet group
              if (inletGroups['remove'].length > 0) {
                const firstRow = rows.length + 1
                inletGroups['remove'].forEach((item) => {
                  const itemRow = Array(template.columns.length).fill('')
                  itemRow[1] = item.particulars
                  itemRow[2] = item.takeoff
                  itemRow[3] = item.unit || 'EA'
                  rows.push(itemRow)
                  formulas.push({ row: rows.length, itemType: 'civil_demo_ea', parsedData: item, section: 'civil_sitework', subsectionName: subSubName })
                })
                // Sum row
                const sumRow = Array(template.columns.length).fill('')
                rows.push(sumRow)
                formulas.push({ row: rows.length, itemType: 'civil_demo_sum', section: 'civil_sitework', subsectionName: subSubName, firstDataRow: firstRow, lastDataRow: rows.length - 1, sumColumns: ['M'] })
                rows.push(Array(template.columns.length).fill(''))
              }
            }

          })
        } else if (subsection.name === 'Excavation') {
          // Excavation subsection
          const excItems = civilOtherItems['Excavation']
          let hasItems = false
          const firstDataRow = rows.length + 1

          // Transformer pad items
          if (excItems['transformer_pad'].length > 0) {
            hasItems = true
            excItems['transformer_pad'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'SQ FT'
              // Height is manual input, leave empty
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_exc_transformer', parsedData: item, section: 'civil_sitework', subsectionName: 'Excavation' })
            })
          }

          // Reinforced sidewalk items
          if (excItems['reinforced_sidewalk'].length > 0) {
            hasItems = true
            excItems['reinforced_sidewalk'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'SQ FT'
              // Height is manual input, leave empty
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_exc_sidewalk', parsedData: item, section: 'civil_sitework', subsectionName: 'Excavation' })
            })
          }

          // Bollard items
          if (excItems['bollard'].length > 0) {
            hasItems = true
            excItems['bollard'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'EA'
              // Height is manual input, leave empty
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_exc_bollard', parsedData: item, section: 'civil_sitework', subsectionName: 'Excavation' })
            })
          }

          // Sum row for all excavation items
          if (hasItems) {
            const lastDataRow = rows.length
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'civil_exc_sum', section: 'civil_sitework', subsectionName: 'Excavation', firstDataRow: firstDataRow, lastDataRow: lastDataRow })
            rows.push(Array(template.columns.length).fill(''))
          } else {
            rows.push(Array(template.columns.length).fill(''))
          }
        } else if (subsection.name === 'Gravel') {
          // Gravel subsection
          const gravelItems = civilOtherItems['Gravel']
          let hasItems = false
          const firstDataRow = rows.length + 1

          // Transformer pad items
          if (gravelItems['transformer_pad'].length > 0) {
            hasItems = true
            gravelItems['transformer_pad'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[3] = item.unit || 'SQ FT'
              // Height is manual entry, leave blank initially
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_gravel_item', parsedData: item, section: 'civil_sitework', subsectionName: 'Gravel' })
            })
          }

          // Reinforced sidewalk items
          if (gravelItems['reinforced_sidewalk'].length > 0) {
            hasItems = true
            gravelItems['reinforced_sidewalk'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[3] = item.unit || 'SQ FT'
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_gravel_item', parsedData: item, section: 'civil_sitework', subsectionName: 'Gravel' })
            })
          }

          // Asphalt items (Full depth asphalt pavement - non-BPP)
          if (gravelItems['asphalt'].length > 0) {
            hasItems = true
            gravelItems['asphalt'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[3] = item.unit || 'SQ FT'
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_gravel_item', parsedData: item, section: 'civil_sitework', subsectionName: 'Gravel' })
            })
          }

          // Sum row for all gravel items
          if (hasItems) {
            const lastDataRow = rows.length
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'civil_gravel_sum', section: 'civil_sitework', subsectionName: 'Gravel', firstDataRow: firstDataRow, lastDataRow: lastDataRow })
            rows.push(Array(template.columns.length).fill(''))
          } else {
            rows.push(Array(template.columns.length).fill(''))
          }
        } else if (subsection.name === 'Concrete Pavement') {
          // Concrete Pavement subsection
          const cpItems = civilOtherItems['Concrete Pavement']
          if (cpItems.length > 0) {
            const firstRow = rows.length + 1
            cpItems.forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'SQ FT'
              if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_concrete_pavement', parsedData: item, section: 'civil_sitework', subsectionName: 'Concrete Pavement' })
            })
            // Sum row
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'civil_concrete_pavement_sum', section: 'civil_sitework', subsectionName: 'Concrete Pavement', firstDataRow: firstRow, lastDataRow: rows.length - 1 })
            rows.push(Array(template.columns.length).fill(''))
          } else {
            rows.push(Array(template.columns.length).fill(''))
          }
        } else if (subsection.name === 'Asphalt') {
          // Asphalt subsection
          const asphaltItems = civilOtherItems['Asphalt']
          if (asphaltItems.length > 0) {
            const firstRow = rows.length + 1
            asphaltItems.forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'SQ FT'
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_asphalt', parsedData: item, section: 'civil_sitework', subsectionName: 'Asphalt' })
            })
            // Sum row
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'civil_asphalt_sum', section: 'civil_sitework', subsectionName: 'Asphalt', firstDataRow: firstRow, lastDataRow: rows.length - 1 })
            rows.push(Array(template.columns.length).fill(''))
          } else {
            rows.push(Array(template.columns.length).fill(''))
          }
        } else if (subsection.name === 'Pads') {
          // Pads subsection
          const padsItems = civilOtherItems['Pads']
          if (padsItems.length > 0) {
            const firstRow = rows.length + 1
            padsItems.forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'SQ FT'
              if (item.parsed?.qty != null) itemRow[4] = item.parsed.qty
              if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_pads', parsedData: item, section: 'civil_sitework', subsectionName: 'Pads' })
            })
            // Sum row
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'civil_pads_sum', section: 'civil_sitework', subsectionName: 'Pads', firstDataRow: firstRow, lastDataRow: rows.length - 1 })
            rows.push(Array(template.columns.length).fill(''))
          } else {
            rows.push(Array(template.columns.length).fill(''))
          }
        } else if (subsection.name === 'Soil Erosion') {
          // Soil Erosion subsection
          const seItems = civilOtherItems['Soil Erosion']
          let hasItems = false

          // Stabilized entrance items
          if (seItems['stabilized_entrance'].length > 0) {
            hasItems = true
            seItems['stabilized_entrance'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'SQ FT'
              if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_soil_stabilized', parsedData: item, section: 'civil_sitework', subsectionName: 'Soil Erosion' })
            })
          }

          // Silt fence items
          if (seItems['silt_fence'].length > 0) {
            hasItems = true
            seItems['silt_fence'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'FT'
              if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_soil_silt_fence', parsedData: item, section: 'civil_sitework', subsectionName: 'Soil Erosion' })
            })
          }

          // Inlet filter items
          if (seItems['inlet_filter'].length > 0) {
            hasItems = true
            seItems['inlet_filter'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'EA'
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_soil_inlet_filter', parsedData: item, section: 'civil_sitework', subsectionName: 'Soil Erosion' })
            })
          }

          if (hasItems) {
            rows.push(Array(template.columns.length).fill(''))
          }
        } else if (subsection.name === 'Fence') {
          // Fence subsection
          const fenceItems = civilOtherItems['Fence']
          let hasItems = false

          // Construction fence items
          if (fenceItems['construction_fence'].length > 0) {
            hasItems = true
            const firstRow = rows.length + 1
            fenceItems['construction_fence'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'FT'
              if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_fence', parsedData: item, section: 'civil_sitework', subsectionName: 'Fence' })
            })
            // Sum row
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'civil_fence_sum', section: 'civil_sitework', subsectionName: 'Fence', firstDataRow: firstRow, lastDataRow: rows.length - 1 })
            rows.push(Array(template.columns.length).fill(''))
          }

          // Proposed fence items
          if (fenceItems['proposed_fence'].length > 0) {
            hasItems = true
            const firstRow = rows.length + 1
            fenceItems['proposed_fence'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'FT'
              if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_fence', parsedData: item, section: 'civil_sitework', subsectionName: 'Fence' })
            })
            // Sum row
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'civil_fence_sum', section: 'civil_sitework', subsectionName: 'Fence', firstDataRow: firstRow, lastDataRow: rows.length - 1 })
            rows.push(Array(template.columns.length).fill(''))
          }

          // Guiderail items
          if (fenceItems['guiderail'].length > 0) {
            hasItems = true
            const firstRow = rows.length + 1
            fenceItems['guiderail'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'FT'
              if (item.parsed?.heightValue != null) itemRow[7] = item.parsed.heightValue
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_fence', parsedData: item, section: 'civil_sitework', subsectionName: 'Fence' })
            })
            // Sum row
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'civil_fence_sum', section: 'civil_sitework', subsectionName: 'Fence', firstDataRow: firstRow, lastDataRow: rows.length - 1 })
            rows.push(Array(template.columns.length).fill(''))
          }


        } else if (subsection.name === 'Concrete filled steel pipe bollard') {
          // Bollard subsection
          const bollardGroups = civilOtherItems['Concrete filled steel pipe bollard']

          // Group 1: Items with Footing
          if (bollardGroups['footing'] && bollardGroups['footing'].length > 0) {
            const firstDataRow = rows.length + 1
            bollardGroups['footing'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'EA'
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_bollard_footing_item', parsedData: item, section: 'civil_sitework', subsectionName: 'Concrete filled steel pipe bollard' })
            })

            // Sum row for Footing Items
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'civil_bollard_footing_sum', section: 'civil_sitework', firstDataRow, lastDataRow: rows.length - 1 })
            rows.push(Array(template.columns.length).fill(''))
          }

          // Group 2: Simple Items
          if (bollardGroups['simple'] && bollardGroups['simple'].length > 0) {
            const firstDataRow = rows.length + 1
            bollardGroups['simple'].forEach((item) => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'EA'
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_bollard_simple_item', parsedData: item, section: 'civil_sitework', subsectionName: 'Concrete filled steel pipe bollard' })
            })

            // Sum row for Simple Items
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({ row: rows.length, itemType: 'civil_bollard_simple_sum', section: 'civil_sitework', firstDataRow, lastDataRow: rows.length - 1 })
            rows.push(Array(template.columns.length).fill(''))
          }


        } else if (subsection.name === 'Site') {
          const siteItems = civilOtherItems['Site']
          const siteOrder = ['Hydrant', 'Wheel stop', 'Drain', 'Protection', 'Signages', 'Main line']

          siteOrder.forEach(subName => {
            const groupData = siteItems[subName]
            let hasData = false
            if (Array.isArray(groupData)) {
              hasData = groupData && groupData.length > 0
            } else {
              hasData = Object.values(groupData).some(sub => sub && sub.length > 0)
            }

            if (hasData) {
              // Render Group Header
              const headerRow = Array(template.columns.length).fill('')
              headerRow[1] = '  ' + subName + ':'
              rows.push(headerRow)
            }

            if (Array.isArray(groupData)) {
              if (groupData && groupData.length > 0) {
                const firstDataRow = rows.length + 1
                groupData.forEach(item => {
                  const itemRow = Array(template.columns.length).fill('')
                  itemRow[1] = item.particulars
                  itemRow[2] = item.takeoff
                  itemRow[3] = item.unit || 'EA'
                  rows.push(itemRow)
                  formulas.push({ row: rows.length, itemType: 'civil_site_item', parsedData: item, section: 'civil_sitework', subsectionName: 'Site' })
                })

                // Sum Row
                const sumRow = Array(template.columns.length).fill('')
                rows.push(sumRow)
                formulas.push({ row: rows.length, itemType: 'civil_site_sum', section: 'civil_sitework', siteGroupKey: subName, firstDataRow, lastDataRow: rows.length - 1 })
              }
            } else {
              // Nested group (Drain, Main line)
              Object.keys(groupData).forEach(key => {
                const subItems = groupData[key]
                if (subItems && subItems.length > 0) {
                  const firstDataRow = rows.length + 1
                  subItems.forEach(item => {
                    const itemRow = Array(template.columns.length).fill('')
                    itemRow[1] = item.particulars
                    itemRow[2] = item.takeoff
                    itemRow[3] = item.unit || 'EA'
                    rows.push(itemRow)
                    formulas.push({ row: rows.length, itemType: 'civil_site_item', parsedData: item, section: 'civil_sitework', subsectionName: 'Site' })
                  })

                  // Sum Row
                  const sumRow = Array(template.columns.length).fill('')
                  rows.push(sumRow)
                  formulas.push({ row: rows.length, itemType: 'civil_site_sum', section: 'civil_sitework', siteGroupKey: key, firstDataRow, lastDataRow: rows.length - 1 })

                  // Gap
                  rows.push(Array(template.columns.length).fill(''))
                }
              })
              // Remove last gap if added by nested loop
              if (rows.length > 0 && rows[rows.length - 1].every(c => c === '')) {
                rows.pop()
              }
            }
            if (hasData) {
              rows.push(Array(template.columns.length).fill(''))
            }
          })
        } else if (subsection.name === 'Drains & Utilities') {
          const items = civilOtherItems['Drains & Utilities']
          if (items && items.length > 0) {
            items.forEach(item => {
              if (item.particulars.toLowerCase().includes('backwater valve')) {
                rows.push(Array(template.columns.length).fill(''))
              }

              const itemRow = Array(template.columns.length).fill('')
              if (item.particulars.toLowerCase().includes('backwater valve')) {
                itemRow[0] = 'Concrete'
              }
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'EA'
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_drains_utilities_item', parsedData: item, section: 'civil_sitework', subsectionName: 'Drains & Utilities' })
            })
          }
        } else if (subsection.name === 'Alternate') {
          const items = civilOtherItems['Alternate']
          if (items && items.length > 0) {
            // Gap before Alternate
            rows.push(Array(template.columns.length).fill(''))

            // Render Alternate Header
            const altHeaderRow = Array(template.columns.length).fill('')
            altHeaderRow[1] = 'Alternate:'
            rows.push(altHeaderRow)
            formulas.push({ row: rows.length, itemType: 'civil_alternate_header', section: 'civil_sitework', subsectionName: 'Alternate' })

            items.forEach(item => {
              const itemRow = Array(template.columns.length).fill('')
              itemRow[1] = item.particulars
              itemRow[2] = item.takeoff
              itemRow[3] = item.unit || 'EA'
              rows.push(itemRow)
              formulas.push({ row: rows.length, itemType: 'civil_alternate_item', parsedData: item, section: 'civil_sitework', subsectionName: 'Alternate' })
            })
          }
        } else if (subsection.name === 'Ele') {
          // Civil/Sitework Ele subsection with Excavation, Backfill, Gravel
          const createCivilEleGroup = (subSubName, takeoffSourceType) => {
            // Sub-subsection header
            const headerRow = Array(template.columns.length).fill('')
            headerRow[1] = `  ${subSubName}:`
            rows.push(headerRow)

            // Data row
            const dataRow = Array(template.columns.length).fill('')
            dataRow[1] = subSubName
            dataRow[3] = 'FT'
            rows.push(dataRow)
            formulas.push({
              row: rows.length,
              itemType: 'civil_ele_item',
              section: 'civil_sitework',
              subsectionName: subsection.name,
              subSubsectionName: subSubName,
              takeoffSourceType
            })

            // Sum row (for J and L)
            const sumRow = Array(template.columns.length).fill('')
            rows.push(sumRow)
            formulas.push({
              row: rows.length,
              itemType: 'civil_ele_sum',
              section: 'civil_sitework',
              subsectionName: subsection.name,
              subSubsectionName: subSubName,
              firstDataRow: rows.length - 1,
              lastDataRow: rows.length - 1
            })

            // Blank row between groups
            rows.push(Array(template.columns.length).fill(''))
          }

          // Excavation - takeoff from C of Proposed underground electrical conduit item from Drains & Utilities
          createCivilEleGroup('Excavation', 'drains_conduit')
          // Backfill - takeoff from C of Excavation item from this Ele section
          createCivilEleGroup('Backfill', 'ele_excavation')
          // Gravel - takeoff from C of Proposed underground electrical conduit item from Drains & Utilities
          createCivilEleGroup('Gravel', 'drains_conduit')
        } else if (subsection.name === 'Gas') {
          // Civil/Sitework Gas subsection with Excavation, Backfill, Gravel

          // Excavation sub-subsection
          const excavationHeaderRow = Array(template.columns.length).fill('')
          excavationHeaderRow[1] = '  Excavation:'
          rows.push(excavationHeaderRow)

          const excavationDataRow = Array(template.columns.length).fill('')
          excavationDataRow[1] = 'Proposed Underground gas service lateral'
          excavationDataRow[3] = 'FT'
          rows.push(excavationDataRow)
          formulas.push({
            row: rows.length,
            itemType: 'civil_gas_water_item',
            section: 'civil_sitework',
            subsectionName: 'Gas',
            subSubsectionName: 'Excavation',
            takeoffSourceType: 'drains_gas_lateral',
            isGravel: false
          })

          const excavationSumRow = Array(template.columns.length).fill('')
          rows.push(excavationSumRow)
          formulas.push({
            row: rows.length,
            itemType: 'civil_gas_water_sum',
            section: 'civil_sitework',
            subsectionName: 'Gas',
            subSubsectionName: 'Excavation',
            firstDataRow: rows.length - 1,
            lastDataRow: rows.length - 1,
            isGravel: false
          })
          rows.push(Array(template.columns.length).fill(''))

          // Backfill sub-subsection
          const backfillHeaderRow = Array(template.columns.length).fill('')
          backfillHeaderRow[1] = '  Backfill:'
          rows.push(backfillHeaderRow)

          const backfillDataRow = Array(template.columns.length).fill('')
          backfillDataRow[1] = 'Proposed Underground gas service lateral'
          backfillDataRow[3] = 'FT'
          rows.push(backfillDataRow)
          formulas.push({
            row: rows.length,
            itemType: 'civil_gas_water_item',
            section: 'civil_sitework',
            subsectionName: 'Gas',
            subSubsectionName: 'Backfill',
            takeoffSourceType: 'gas_excavation',
            isGravel: false
          })

          const backfillSumRow = Array(template.columns.length).fill('')
          rows.push(backfillSumRow)
          formulas.push({
            row: rows.length,
            itemType: 'civil_gas_water_sum',
            section: 'civil_sitework',
            subsectionName: 'Gas',
            subSubsectionName: 'Backfill',
            firstDataRow: rows.length - 1,
            lastDataRow: rows.length - 1,
            isGravel: false
          })
          rows.push(Array(template.columns.length).fill(''))

          // Gravel sub-subsection
          const gravelHeaderRow = Array(template.columns.length).fill('')
          gravelHeaderRow[1] = '  Gravel:'
          rows.push(gravelHeaderRow)

          const gravelFirstDataRow = rows.length + 1

          // Gravel item 1
          const gravelDataRow1 = Array(template.columns.length).fill('')
          gravelDataRow1[1] = 'Proposed Underground gas service lateral'
          gravelDataRow1[3] = 'FT'
          rows.push(gravelDataRow1)
          formulas.push({
            row: rows.length,
            itemType: 'civil_gas_water_item',
            section: 'civil_sitework',
            subsectionName: 'Gas',
            subSubsectionName: 'Gravel',
            takeoffSourceType: 'drains_gas_lateral',
            isGravel: true
          })

          // Gravel item 2
          const gravelDataRow2 = Array(template.columns.length).fill('')
          gravelDataRow2[1] = 'Proposed Underground water main'
          gravelDataRow2[3] = 'FT'
          rows.push(gravelDataRow2)
          formulas.push({
            row: rows.length,
            itemType: 'civil_gas_water_item',
            section: 'civil_sitework',
            subsectionName: 'Gas',
            subSubsectionName: 'Gravel',
            takeoffSourceType: 'drains_water_main',
            isGravel: true
          })

          const gravelSumRow = Array(template.columns.length).fill('')
          rows.push(gravelSumRow)
          formulas.push({
            row: rows.length,
            itemType: 'civil_gas_water_sum',
            section: 'civil_sitework',
            subsectionName: 'Gas',
            subSubsectionName: 'Gravel',
            firstDataRow: gravelFirstDataRow,
            lastDataRow: rows.length - 1,
            isGravel: true
          })
          rows.push(Array(template.columns.length).fill(''))

        } else if (subsection.name === 'Water') {
          // Civil/Sitework Water subsection with Excavation, Backfill, Gravel

          // Excavation sub-subsection
          const excavationHeaderRow = Array(template.columns.length).fill('')
          excavationHeaderRow[1] = '  Excavation:'
          rows.push(excavationHeaderRow)

          const excavationFirstDataRow = rows.length + 1

          // Excavation item 1
          const excavationDataRow1 = Array(template.columns.length).fill('')
          excavationDataRow1[1] = 'Proposed Underground 8" water service lateral'
          excavationDataRow1[3] = 'FT'
          rows.push(excavationDataRow1)
          formulas.push({
            row: rows.length,
            itemType: 'civil_gas_water_item',
            section: 'civil_sitework',
            subsectionName: 'Water',
            subSubsectionName: 'Excavation',
            takeoffSourceType: 'drains_sanitary_sewer',
            isGravel: false
          })

          // Excavation item 2
          const excavationDataRow2 = Array(template.columns.length).fill('')
          excavationDataRow2[1] = 'Proposed Underground water main'
          excavationDataRow2[3] = 'FT'
          rows.push(excavationDataRow2)
          formulas.push({
            row: rows.length,
            itemType: 'civil_gas_water_item',
            section: 'civil_sitework',
            subsectionName: 'Water',
            subSubsectionName: 'Excavation',
            takeoffSourceType: 'drains_water_main',
            isGravel: false
          })

          const excavationSumRow = Array(template.columns.length).fill('')
          rows.push(excavationSumRow)
          formulas.push({
            row: rows.length,
            itemType: 'civil_gas_water_sum',
            section: 'civil_sitework',
            subsectionName: 'Water',
            subSubsectionName: 'Excavation',
            firstDataRow: excavationFirstDataRow,
            lastDataRow: rows.length - 1,
            isGravel: false
          })
          rows.push(Array(template.columns.length).fill(''))

          // Backfill sub-subsection
          const backfillHeaderRow = Array(template.columns.length).fill('')
          backfillHeaderRow[1] = '  Backfill:'
          rows.push(backfillHeaderRow)

          const backfillDataRow = Array(template.columns.length).fill('')
          backfillDataRow[1] = 'Proposed Underground 8" water service lateral'
          backfillDataRow[3] = 'FT'
          rows.push(backfillDataRow)
          formulas.push({
            row: rows.length,
            itemType: 'civil_gas_water_item',
            section: 'civil_sitework',
            subsectionName: 'Water',
            subSubsectionName: 'Backfill',
            takeoffSourceType: 'water_backfill_5_rows_above',
            isGravel: false
          })

          const backfillSumRow = Array(template.columns.length).fill('')
          rows.push(backfillSumRow)
          formulas.push({
            row: rows.length,
            itemType: 'civil_gas_water_sum',
            section: 'civil_sitework',
            subsectionName: 'Water',
            subSubsectionName: 'Backfill',
            firstDataRow: rows.length - 1,
            lastDataRow: rows.length - 1,
            isGravel: false
          })
          rows.push(Array(template.columns.length).fill(''))

          // Gravel sub-subsection
          const gravelHeaderRow = Array(template.columns.length).fill('')
          gravelHeaderRow[1] = '  Gravel:'
          rows.push(gravelHeaderRow)

          const gravelFirstDataRow = rows.length + 1

          // Gravel item 1
          const gravelDataRow1 = Array(template.columns.length).fill('')
          gravelDataRow1[1] = 'Proposed Underground 8" water service lateral'
          gravelDataRow1[3] = 'FT'
          rows.push(gravelDataRow1)
          formulas.push({
            row: rows.length,
            itemType: 'civil_gas_water_item',
            section: 'civil_sitework',
            subsectionName: 'Water',
            subSubsectionName: 'Gravel',
            takeoffSourceType: 'drains_sanitary_sewer',
            isGravel: true
          })

          // Gravel item 2
          const gravelDataRow2 = Array(template.columns.length).fill('')
          gravelDataRow2[1] = 'Proposed Underground water main'
          gravelDataRow2[3] = 'FT'
          rows.push(gravelDataRow2)
          formulas.push({
            row: rows.length,
            itemType: 'civil_gas_water_item',
            section: 'civil_sitework',
            subsectionName: 'Water',
            subSubsectionName: 'Gravel',
            takeoffSourceType: 'drains_water_main',
            isGravel: true
          })

          const gravelSumRow = Array(template.columns.length).fill('')
          rows.push(gravelSumRow)
          formulas.push({
            row: rows.length,
            itemType: 'civil_gas_water_sum',
            section: 'civil_sitework',
            subsectionName: 'Water',
            subSubsectionName: 'Gravel',
            firstDataRow: gravelFirstDataRow,
            lastDataRow: rows.length - 1,
            isGravel: true
          })
          rows.push(Array(template.columns.length).fill(''))

        } else if (subsection.subSubsections && subsection.subSubsections.length > 0) {
          // Other subsections with sub-subsections
          subsection.subSubsections.forEach((subSubsection) => {
            const subSubsectionRow = Array(template.columns.length).fill('')
            subSubsectionRow[1] = '  ' + subSubsection.name + ':'
            rows.push(subSubsectionRow)
            rows.push(Array(template.columns.length).fill(''))
          })
        } else {
          // Simple subsection without sub-subsections
          rows.push(Array(template.columns.length).fill(''))
        }
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

  return {
    rows,
    formulas,
    rockExcavationTotals,
    lineDrillTotalFT,
    soldierPileGroups,
    primarySecantItems,
    secondarySecantItems,
    tangentPileItems,
    pargingItems,
    guideWallItems,
    dowelBarItems,
    rockPinItems,
    rockStabilizationItems,
    shotcreteItems,
    permissionGroutingItems,
    buttonItems,
    mudSlabItems,
    drilledFoundationPileGroups,
    helicalFoundationPileGroups,
    drivenFoundationPileItems,
    stelcorDrilledDisplacementPileItems,
    cfaPileItems,
    unusedRawDataRows // Return unused data rows
  }
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
    // Silently handle formatting errors
  }
}

export default {
  generateCalculationSheet,
  generateColumnConfigs,
  applyCalculationSheetFormatting
}
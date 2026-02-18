import React, { useState, useEffect, useRef } from 'react'
import { useLocation, useNavigate } from 'react-router-dom'
import { SpreadsheetComponent, SheetsDirective, SheetDirective, ColumnsDirective, ColumnDirective } from '@syncfusion/ej2-react-spreadsheet'
import { FiArrowLeft, FiSave } from 'react-icons/fi'
import { generateCalculationSheet, generateColumnConfigs } from '../utils/generateCalculationSheet'
import { generateDemolitionFormulas } from '../utils/processors/demolitionProcessor'
import { generateExcavationFormulas } from '../utils/processors/excavationProcessor'
import { generateRockExcavationFormulas } from '../utils/processors/rockExcavationProcessor'
import { generateSoeFormulas, formatDrilledSoldierPileProposalText } from '../utils/processors/soeProcessor'
import { generateFoundationFormulas } from '../utils/processors/foundationProcessor'
import { generateWaterproofingFormulas } from '../utils/processors/waterproofingProcessor'

const Spreadsheet = () => {
  const location = useLocation()
  const navigate = useNavigate()
  const { rawData, fileName, sheetName, headers, rows, selectedTemplate } = location.state || {}
  const [hasData, setHasData] = useState(false)
  const [calculationData, setCalculationData] = useState([])
  const [formulaData, setFormulaData] = useState([])
  const [rockExcavationTotals, setRockExcavationTotals] = useState({ totalSQFT: 0, totalCY: 0 })
  const [lineDrillTotalFT, setLineDrillTotalFT] = useState(0)
  const [sheetCreated, setSheetCreated] = useState(false)
  const spreadsheetRef = useRef(null)
  const LAST_ACTIVE_SHEET_KEY = 'krm:lastActiveSheet'

  useEffect(() => {
    if (rawData && rawData.length > 0) {
      setHasData(true)

      // Generate the Calculations Sheet structure with data
      const template = selectedTemplate || 'capstone'
      const result = generateCalculationSheet(template, rawData)
      setCalculationData(result.rows)
      setFormulaData(result.formulas)
      setRockExcavationTotals(result.rockExcavationTotals || { totalSQFT: 0, totalCY: 0 })
      setLineDrillTotalFT(result.lineDrillTotalFT || 0)
      // Store soldier pile groups for proposal sheet
      window.soldierPileGroups = result.soldierPileGroups || []
      // Store SOE subsection items
      window.soeSubsectionItems = new Map()
      // Store primary secant items for proposal sheet
      window.primarySecantItems = result.primarySecantItems || []
      // Store secondary secant items for proposal sheet
      window.secondarySecantItems = result.secondarySecantItems || []
      // Store tangent pile items for proposal sheet
      window.tangentPileItems = result.tangentPileItems || []
      // Store parging items for proposal sheet
      window.pargingItems = result.pargingItems || []
      // Store guide wall items for proposal sheet
      window.guideWallItems = result.guideWallItems || []
      // Store dowel bar items for proposal sheet
      window.dowelBarItems = result.dowelBarItems || []
      // Store rock pin items for proposal sheet
      window.rockPinItems = result.rockPinItems || []
      // Store rock stabilization items for proposal sheet
      window.rockStabilizationItems = result.rockStabilizationItems || []
      // Store button items for proposal sheet
      window.buttonItems = result.buttonItems || []
      // Store foundation processor results for proposal sheet
      window.drilledFoundationPileGroups = result.drilledFoundationPileGroups || []
      window.helicalFoundationPileGroups = result.helicalFoundationPileGroups || []
      window.drivenFoundationPileItems = result.drivenFoundationPileItems || []
      window.stelcorDrilledDisplacementPileItems = result.stelcorDrilledDisplacementPileItems || []
      window.cfaPileItems = result.cfaPileItems || []
      // Store Foundation subsection items
      window.foundationSubsectionItems = new Map()

      // Check which waterproofing items are present
      if (rawData && rawData.length > 1) {
        const headers = rawData[0]
        const dataRows = rawData.slice(1)
        const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
        const estimateIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'estimate')

        const itemsToCheck = {
          'foundation walls': /^FW\s*\(/i,
          'retaining wall': /^RW\s*\(/i,
          'vehicle barrier wall': /vehicle\s+barrier\s+wall\s*\(/i,
          'concrete liner wall': /concrete\s+liner\s+wall\s*\(/i,
          'stem wall': /stem\s+wall\s*\(/i,
          'grease trap pit wall': /grease\s+trap\s+pit\s+wall/i,
          'house trap pit wall': /house\s+trap\s+pit\s+wall/i,
          'detention tank wall': /detention\s+tank\s+wall/i,
          'elevator pit walls': /elev(?:ator)?\s+pit\s+wall/i,
          'duplex sewage ejector pit wall': /duplex\s+sewage\s+ejector\s+pit\s+wall/i
        }

        // Negative side items to check
        const negativeSideItemsToCheck = {
          'detention tank wall': /detention\s+tank\s+wall/i,
          'elevator pit walls': /elev(?:ator)?\s+pit\s+wall/i,
          'detention tank slab': /detention\s+tank\s+slab(?!\s+lid)/i,
          'duplex sewage ejector pit wall': /duplex\s+sewage\s+ejector\s+pit\s+wall/i,
          'duplex sewage ejector pit slab': /duplex\s+sewage\s+ejector\s+pit\s+slab/i,
          'elevator pit slab': /elev(?:ator)?\s+pit\s+slab/i
        }

        const presentItems = []
        const presentNegativeSideItems = []

        dataRows.forEach(row => {
          const digitizerItem = row[digitizerIdx]
          if (!digitizerItem) return

          // Check if it's a waterproofing item
          if (estimateIdx >= 0 && row[estimateIdx]) {
            const estimateVal = String(row[estimateIdx]).trim()
            if (estimateVal !== 'Waterproofing') return
          }

          const itemText = String(digitizerItem).trim()

          // Check each pattern for exterior side items
          Object.entries(itemsToCheck).forEach(([itemName, pattern]) => {
            if (pattern.test(itemText) && !presentItems.includes(itemName)) {
              presentItems.push(itemName)
            }
          })

          // Check each pattern for negative side items
          Object.entries(negativeSideItemsToCheck).forEach(([itemName, pattern]) => {
            if (pattern.test(itemText) && !presentNegativeSideItems.includes(itemName)) {
              presentNegativeSideItems.push(itemName)
            }
          })
        })

        // Store present items for use in proposal sheet
        window.waterproofingPresentItems = presentItems
        window.waterproofingNegativeSideItems = presentNegativeSideItems
      }
    } else {
      // No data - create empty structure
      const template = selectedTemplate || 'capstone'
      const result = generateCalculationSheet(template, null)
      setCalculationData(result.rows)
      setFormulaData(result.formulas || [])
      setRockExcavationTotals(result.rockExcavationTotals || { totalSQFT: 0, totalCY: 0 })
      setLineDrillTotalFT(result.lineDrillTotalFT || 0)
      // Store soldier pile groups for proposal sheet
      window.soldierPileGroups = result.soldierPileGroups || []
      // Store SOE subsection items
      window.soeSubsectionItems = new Map()
      // Store primary secant items for proposal sheet
      window.primarySecantItems = result.primarySecantItems || []
      // Store secondary secant items for proposal sheet
      window.secondarySecantItems = result.secondarySecantItems || []
      // Store tangent pile items for proposal sheet
      window.tangentPileItems = result.tangentPileItems || []
      // Store parging items for proposal sheet
      window.pargingItems = result.pargingItems || []
      // Store guide wall items for proposal sheet
      window.guideWallItems = result.guideWallItems || []
      // Store dowel bar items for proposal sheet
      window.dowelBarItems = result.dowelBarItems || []
      // Store rock pin items for proposal sheet
      window.rockPinItems = result.rockPinItems || []
      // Store rock stabilization items for proposal sheet
      window.rockStabilizationItems = result.rockStabilizationItems || []
      // Store shotcrete items for proposal sheet
      window.shotcreteItems = result.shotcreteItems || []
      // Store permission grouting items for proposal sheet
      window.permissionGroutingItems = result.permissionGroutingItems || []
      // Store button items for proposal sheet
      window.buttonItems = result.buttonItems || []
      // Store mud slab items for proposal sheet
      window.mudSlabItems = result.mudSlabItems || []
      // Store foundation processor results for proposal sheet
      window.drilledFoundationPileGroups = result.drilledFoundationPileGroups || []
      window.helicalFoundationPileGroups = result.helicalFoundationPileGroups || []
      window.drivenFoundationPileItems = result.drivenFoundationPileItems || []
      window.stelcorDrilledDisplacementPileItems = result.stelcorDrilledDisplacementPileItems || []
      window.cfaPileItems = result.cfaPileItems || []
      // Store Foundation subsection items
      window.foundationSubsectionItems = new Map()
    }
  }, [rawData, fileName, sheetName, rows, selectedTemplate])

  // Effect to apply data and formulas when calculationData or formulaData changes
  useEffect(() => {
    if (spreadsheetRef.current && calculationData.length > 0) {
      applyDataToSpreadsheet()
    }
  }, [calculationData, formulaData])

  const applyDataToSpreadsheet = () => {
    if (!spreadsheetRef.current) return
    const spreadsheet = spreadsheetRef.current
    const prevSheetName = spreadsheet.getActiveSheet?.().name
    const prevSelectedRange = spreadsheet.getActiveSheet?.().selectedRange

    // Always write calculations to the Calculations Sheet so Proposal Sheet stays independent.
    // We also restore the user's previously active sheet after we're done.
    try {
      spreadsheet.goTo('Calculations Sheet!A1')
    } catch (e) {
      // If goTo fails, we still attempt to write; but this should be rare.
    }

    // Set data for Calculations Sheet
    calculationData.forEach((row, rowIndex) => {
      row.forEach((cellValue, colIndex) => {
        const colLetter = String.fromCharCode(65 + colIndex) // A, B, C, etc.
        const cellAddress = `${colLetter}${rowIndex + 1}`

        try {
          if (cellValue !== '' && cellValue !== null && cellValue !== undefined) {
            let valueToSet = cellValue

            // For Particulars column (B), prevent date conversion
            if (colIndex === 1 && typeof cellValue === 'string') {
              if (/^[A-Z]+-\d+/.test(cellValue) || /^\d+-\d+/.test(cellValue)) {
                valueToSet = "'" + cellValue
              }
              spreadsheet.cellFormat({ format: '@' }, cellAddress)
            }

            spreadsheet.updateCell({ value: valueToSet }, cellAddress)
          }
        } catch (error) {
          // Ignore errors
        }
      })
    })

    // Apply formulas (defer foundation_sum L so data-row L values like landings are set first)
    const deferredFoundationSumL = []
    formulaData.forEach((formulaInfo) => {
      const { row, itemType, parsedData, section, subsection } = formulaInfo

      let formulas
      if (section === 'demolition') {
        if (itemType === 'demolition_sum') {
          // Handle sum formulas for demolition subsection
          const { firstDataRow, lastDataRow, subsection } = formulaInfo
          try {
            // Sum for SQ FT (column J) - sum all items in subsection
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)

            // Sum for CY (column L) - sum all items in subsection
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)

            // For Demo isolated footing, also sum column M (QTY)
            if (subsection === 'Demo isolated footing') {
              spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `M${row}`)
            }
          } catch (error) {
            // Ignore errors
          }
          return
        }

        // Handle "For demo Extra line item use this" rows
        if (itemType === 'demo_extra_sqft') {
          // Row 1: In SQ FT
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`) // Col J: =C
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`) // Col L: J*H/27
            // Apply red color to B, J, L, M
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `M${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'demo_extra_ft') {
          // Row 2: In FT
          try {
            spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, `J${row}`) // Col J: =C*G
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`) // Col L: J*H/27
            // Apply red color to B, J, L, M
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `M${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'demo_extra_ea') {
          // Row 3: In EA
          try {
            spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, `J${row}`) // Col J: =C*F*G
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`) // Col L: J*H/27
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`) // Col M: =C
            // Apply red color to B, J, L, M
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `M${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'demo_extra_sqft') {
          // Row 1: In SQ FT
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`) // Col J: =C
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`) // Col L: J*H/27
            // Apply red color to B, J, L
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'demo_extra_ft') {
          // Row 2: In FT
          try {
            spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, `J${row}`) // Col J: =C*G
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`) // Col L: J*H/27
            // Apply red color to B, J, L
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'demo_extra_ea') {
          // Row 3: In EA
          try {
            spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, `J${row}`) // Col J: =C*F*G
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`) // Col L: J*H/27
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`) // Col M: =C
            // Apply red color to B, J, L, M
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `M${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        // Generate formulas based on subsection
        formulas = generateDemolitionFormulas(itemType === 'demolition_item' ? subsection : itemType, row, parsedData)
      } else if (section === 'excavation') {
        if (itemType === 'excavation_sum') {
          // Handle sum formulas for excavation
          const { firstDataRow, lastDataRow, sqFtSumRows } = formulaInfo
          try {
            // Sum for SQ FT (column J) - sum all items
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)

            // Sum for 1.3*CY (column L) - sum all items
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'excavation_havg') {
          // Handle Havg formula: (1.3*CY sum * 27) / SQ FT sum
          const { sumRowNumber } = formulaInfo
          try {
            // Column C: Formula = (L{sumRow} * 27) / J{sumRow}
            spreadsheet.updateCell({ formula: `=(L${sumRowNumber}*27)/J${sumRowNumber}` }, `C${row}`)
            // Column B: red and not bold
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'normal', fontStyle: 'normal' }, `B${row}`)
            // Column C: black and not bold
            spreadsheet.cellFormat({ color: '#000000', fontWeight: 'normal' }, `C${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'backfill_sum') {
          // Handle sum formulas for backfill subsection
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            // Sum for SQ FT (column J) - sum all backfill items
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)

            // Sum for CY (column L) - sum all backfill items
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'mud_slab_sum') {
          // Handle sum formulas for mud slab subsection
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            // Sum for SQ FT (column J) - sum all mud slab items
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)

            // Sum for CY (column L) - sum all mud slab items
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'backfill_extra_sqft') {
          // Row 1: In SQ FT
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`) // Col J: =C
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`) // Col L: J*H/27
            // Apply red color to B, J, L, M
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `M${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'backfill_extra_ft') {
          // Row 2: In FT
          try {
            spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, `J${row}`) // Col J: =C*G
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`) // Col L: J*H/27
            // Apply red color to B, J, L, M
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `M${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'backfill_extra_ea') {
          // Row 3: In EA
          try {
            spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, `J${row}`) // Col J: =C*F*G
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`) // Col L: J*H/27
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`) // Col M: =C
            // Apply red color to B, J, L, M
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `M${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'soil_exc_extra_sqft') {
          // Row 1: In SQ FT
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`) // Col J: =C
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `K${row}`) // Col K: CY
            spreadsheet.updateCell({ formula: `=K${row}*1.3` }, `L${row}`) // Col L: 1.3*CY
            // Apply red color to B, J, K, L
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', textDecoration: 'line-through', fontWeight: 'bold' }, `K${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'soil_exc_extra_ft') {
          // Row 2: In FT
          try {
            spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, `J${row}`) // Col J: =C*G
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `K${row}`) // Col K: CY
            spreadsheet.updateCell({ formula: `=K${row}*1.3` }, `L${row}`) // Col L: 1.3*CY
            // Apply red color to B, J, K, L
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', textDecoration: 'line-through', fontWeight: 'bold' }, `K${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'soil_exc_extra_ea') {
          // Row 3: In EA
          try {
            spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, `J${row}`) // Col J: =C*F*G
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `K${row}`) // Col K: CY
            spreadsheet.updateCell({ formula: `=K${row}*1.3` }, `L${row}`) // Col L: 1.3*CY
            // Column M is empty as requested
            // Apply red color to B, J, K, L
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', textDecoration: 'line-through', fontWeight: 'bold' }, `K${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        formulas = generateExcavationFormulas(itemType === 'excavation_item' ? parsedData.itemType : itemType, row, parsedData)
      } else if (section === 'rock_excavation') {
        if (itemType === 'line_drill_sub_header') {
          // Lifts and Height header row - black and italic
          try {
            spreadsheet.cellFormat({ color: '#000000', fontStyle: 'italic', fontWeight: 'normal' }, `E${row}`)
            spreadsheet.cellFormat({ color: '#000000', fontStyle: 'italic', fontWeight: 'normal' }, `H${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'line_drill_concrete_pier') {
          const { refRow } = formulaInfo
          try {
            if (refRow) {
              spreadsheet.updateCell({ formula: `=B${refRow}` }, `B${row}`)
              spreadsheet.updateCell({ formula: `=((G${refRow}+F${refRow})*2)*C${refRow}` }, `C${row}`)
              spreadsheet.updateCell({ formula: `=H${refRow}` }, `H${row}`)
            }
            spreadsheet.updateCell({ value: 'FT' }, `D${row}`)
            spreadsheet.updateCell({ formula: `=ROUNDUP(H${row}/2,0)` }, `E${row}`)
            spreadsheet.updateCell({ formula: `=E${row}*C${row}` }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'line_drill_sewage_pit') {
          const { refRow } = formulaInfo
          try {
            if (refRow) {
              spreadsheet.updateCell({ formula: `=B${refRow}` }, `B${row}`)
              spreadsheet.updateCell({ formula: `=SQRT(C${refRow})*4` }, `C${row}`)
              spreadsheet.updateCell({ formula: `=H${refRow}` }, `H${row}`)
            }
            spreadsheet.updateCell({ value: 'FT' }, `D${row}`)
            spreadsheet.updateCell({ formula: `=ROUNDUP(H${row}/2,0)` }, `E${row}`)
            spreadsheet.updateCell({ formula: `=E${row}*C${row}` }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'line_drill_sump_pit') {
          const { refRow } = formulaInfo
          try {
            if (refRow) {
              spreadsheet.updateCell({ formula: `=B${refRow}` }, `B${row}`)
              spreadsheet.updateCell({ formula: `=C${refRow}*8` }, `C${row}`)
            }
            spreadsheet.updateCell({ value: 'FT' }, `D${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'line_drilling') {
          const rockFormulas = generateRockExcavationFormulas(itemType, row, parsedData)
          try {
            if (rockFormulas.qty) spreadsheet.updateCell({ formula: `=${rockFormulas.qty}` }, `E${row}`)
            if (rockFormulas.ft) spreadsheet.updateCell({ formula: `=${rockFormulas.ft}` }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'line_drill_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})*2` }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `I${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'rock_excavation_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }
        if (itemType === 'rock_excavation_havg') {
          const { sumRowNumber } = formulaInfo
          try {
            // Havg = (CY sum * 27) / SQ FT sum
            spreadsheet.updateCell({ formula: `=(L${sumRowNumber}*27)/J${sumRowNumber}` }, `C${row}`)
            // Column B: red and not bold
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'normal', fontStyle: 'normal' }, `B${row}`)
            // Column C: black and not bold
            spreadsheet.cellFormat({ color: '#000000', fontWeight: 'normal' }, `C${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'rock_exc_extra_sqft') {
          // Row 1: In SQ FT
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`) // Col J: =C
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`) // Col L: J*H/27
            // Apply red color to B, J, L
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'rock_exc_extra_ft') {
          // Row 2: In FT
          try {
            spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, `J${row}`) // Col J: =C*G
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`) // Col L: J*H/27
            // Apply red color to B, J, L
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'rock_exc_extra_ea') {
          // Row 3: In EA
          try {
            spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, `J${row}`) // Col J: =C*F*G
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`) // Col L: J*H/27
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`) // Col M: =C
            // Apply red color to B, J, L, M
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `M${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        formulas = generateRockExcavationFormulas(itemType === 'rock_excavation_item' ? parsedData.itemType : itemType, row, parsedData)
      } else if (section === 'soe') {
        if (['soldier_pile_item', 'timber_brace_item', 'timber_waler_item', 'timber_stringer_item', 'drilled_hole_grout_item', 'soe_generic_item', 'backpacking_item', 'supporting_angle', 'parging', 'heel_block', 'underpinning', 'shims', 'rock_anchor', 'rock_bolt', 'anchor', 'tie_back', 'concrete_soil_retention_pier', 'guide_wall', 'dowel_bar', 'rock_pin', 'shotcrete', 'permission_grouting', 'button', 'rock_stabilization', 'form_board'].includes(itemType)) {
          try {
            const soeFormulas = generateSoeFormulas(itemType, row, parsedData || formulaInfo)
            if (soeFormulas.takeoff) spreadsheet.updateCell({ formula: `=${soeFormulas.takeoff}` }, `C${row}`)
            if (soeFormulas.length !== undefined) {
              if (typeof soeFormulas.length === 'string') {
                spreadsheet.updateCell({ formula: `=${soeFormulas.length}` }, `F${row}`)
              } else {
                spreadsheet.updateCell({ value: soeFormulas.length }, `F${row}`)
              }
            }
            if (soeFormulas.width !== undefined) {
              if (typeof soeFormulas.width === 'string') {
                spreadsheet.updateCell({ formula: `=${soeFormulas.width}` }, `G${row}`)
              } else {
                spreadsheet.updateCell({ value: soeFormulas.width }, `G${row}`)
              }
            }
            if (soeFormulas.height !== undefined) {
              if (typeof soeFormulas.height === 'string') {
                spreadsheet.updateCell({ formula: `=${soeFormulas.height}` }, `H${row}`)
              } else {
                spreadsheet.updateCell({ value: soeFormulas.height }, `H${row}`)
              }
            }
            if (soeFormulas.qty !== undefined) {
              if (typeof soeFormulas.qty === 'string') {
                spreadsheet.updateCell({ formula: `=${soeFormulas.qty}` }, `E${row}`)
              } else {
                spreadsheet.updateCell({ value: soeFormulas.qty }, `E${row}`)
              }
            }
            if (soeFormulas.ft) {
              spreadsheet.updateCell({ formula: `=${soeFormulas.ft}` }, `I${row}`)
            } else {
              spreadsheet.updateCell({ value: '' }, `I${row}`)
            }
            if (soeFormulas.sqFt) {
              spreadsheet.updateCell({ formula: `=${soeFormulas.sqFt}` }, `J${row}`)
            } else {
              spreadsheet.updateCell({ value: '' }, `J${row}`)
            }
            if (soeFormulas.lbs) {
              spreadsheet.updateCell({ formula: `=${soeFormulas.lbs}` }, `K${row}`)
            } else {
              spreadsheet.updateCell({ value: '' }, `K${row}`)
            }
            if (soeFormulas.cy) {
              spreadsheet.updateCell({ formula: `=${soeFormulas.cy}` }, `L${row}`)
            } else {
              spreadsheet.updateCell({ value: '' }, `L${row}`)
            }
            if (soeFormulas.qtyFinal) spreadsheet.updateCell({ formula: `=${soeFormulas.qtyFinal}` }, `M${row}`)

            // Apply red color to item names
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)

            // Special formatting for Backpacking
            if (itemType === 'backpacking_item') {
              spreadsheet.cellFormat({ color: '#000000' }, `C${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            }

            // Special formatting for Shims - columns I and J in red
            if (itemType === 'shims') {
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `I${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            }
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'soldier_pile_group_sum' || itemType === 'timber_soldier_pile_group_sum' || itemType === 'timber_plank_group_sum' || itemType === 'timber_raker_group_sum' || itemType === 'timber_brace_group_sum' || itemType === 'timber_waler_group_sum' || itemType === 'timber_stringer_group_sum' || itemType === 'timber_post_group_sum' || itemType === 'vertical_timber_sheets_group_sum' || itemType === 'horizontal_timber_sheets_group_sum' || itemType === 'drilled_hole_grout_group_sum' || itemType === 'soe_generic_sum') {
          const { firstDataRow, lastDataRow, subsectionName } = formulaInfo
          try {
            // Standard sum for FT (I) for all SOE, except Heel blocks
            const ftSumSubsections = ['Rock anchors', 'Rock bolts', 'Tie back anchor', 'Tie down anchor', 'Dowel bar', 'Rock pins', 'Shotcrete', 'Permission grouting', 'Form board', 'Guide wall']
            if ((itemType === 'timber_waler_group_sum' || itemType === 'timber_brace_group_sum' || itemType === 'timber_stringer_group_sum' || itemType === 'drilled_hole_grout_group_sum') || (itemType !== 'timber_brace_group_sum' && itemType !== 'timber_stringer_group_sum' && itemType !== 'drilled_hole_grout_group_sum' && subsectionName !== 'Heel blocks' && (ftSumSubsections.includes(subsectionName) || !['Concrete soil retention piers', 'Buttons', 'Rock stabilization'].includes(subsectionName)))) {
              spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `I${row}`)
            }

            // Sum for SQ FT (J)
            const sqFtSubsections = ['Sheet pile', 'Timber lagging', 'Timber sheeting', 'Vertical timber sheets', 'Horizontal timber sheets', 'Parging', 'Heel blocks', 'Underpinning', 'Concrete soil retention piers', 'Guide wall', 'Shotcrete', 'Permission grouting', 'Buttons', 'Form board', 'Rock stabilization']
            if (sqFtSubsections.includes(subsectionName) || itemType === 'drilled_hole_grout_group_sum') {
              spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            }

            // Sum for LBS (K)
            const lbsSubsections = [
              'Primary secant piles',
              'Secondary secant piles',
              'Tangent piles',
              'Sheet pile',
              'Waler',
              'Raker',
              'Upper Raker',
              'Lower Raker',
              'Stand off',
              'Kicker',
              'Channel',
              'Roll chock',
              'Stud beam',
              'Inner corner brace',
              'Knee brace',
              'Supporting angle'
            ]
            if (itemType === 'soldier_pile_group_sum' || lbsSubsections.includes(subsectionName)) {
              spreadsheet.updateCell({ formula: `=SUM(K${firstDataRow}:K${lastDataRow})` }, `K${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `K${row}`)
            }

            // Sum for QTY (M)
            const qtySubsections = [
              'Primary secant piles',
              'Secondary secant piles',
              'Tangent piles',
              'Waler',
              'Raker',
              'Upper Raker',
              'Lower Raker',
              'Stand off',
              'Kicker',
              'Channel',
              'Roll chock',
              'Stud beam',
              'Inner corner brace',
              'Knee brace',
              'Supporting angle',
              'Heel blocks',
              'Underpinning',
              'Rock anchors',
              'Rock bolts',
              'Tie back anchor',
              'Tie down anchor',
              'Concrete soil retention piers',
              'Dowel bar',
              'Rock pins',
              'Buttons'
            ]
            if ((itemType === 'soldier_pile_group_sum' || itemType === 'timber_soldier_pile_group_sum' || itemType === 'timber_plank_group_sum' || itemType === 'timber_raker_group_sum' || itemType === 'timber_waler_group_sum' || itemType === 'timber_brace_group_sum' || itemType === 'timber_post_group_sum' || itemType === 'drilled_hole_grout_group_sum') || qtySubsections.includes(subsectionName)) {
              spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `M${row}`)
            }

            // Sum for CY (L)
            const cySubsections = ['Heel blocks', 'Underpinning', 'Concrete soil retention piers', 'Guide wall', 'Shotcrete', 'Buttons', 'Rock stabilization']
            if (cySubsections.includes(subsectionName) || itemType === 'drilled_hole_grout_group_sum') {
              spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
            }
          } catch (error) {
            // Ignore errors
          }
          return
        }
      } else if (section === 'foundation') {
        if (itemType === 'foundation_section_cy_sum') {
          const { sumRows } = formulaInfo
          try {
            if (sumRows && sumRows.length > 0) {
              const sumRefs = sumRows.map((r) => `L${r}`).join(',')
              spreadsheet.updateCell({ formula: `=SUM(${sumRefs})` }, `C${row}`)
            }
          } catch (error) {
            console.error(`Error applying Foundation section CY sum at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'stairs_on_grade_group_header') {
          try {
            spreadsheet.cellFormat({ textDecoration: 'underline' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying stairs group header format at row ${row}:`, error)
          }
          return
        }
        // Handle "For foundation Extra line item use this" rows (same as Demo extra)
        if (itemType === 'foundation_extra_sqft') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `M${row}`)
          } catch (error) {
            console.error(`Error applying foundation extra SQ FT formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'foundation_extra_ft') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `M${row}`)
          } catch (error) {
            console.error(`Error applying foundation extra FT formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'foundation_extra_ea') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `M${row}`)
          } catch (error) {
            console.error(`Error applying foundation extra EA formula at row ${row}:`, error)
          }
          return
        }
        if (['drilled_foundation_pile', 'helical_foundation_pile', 'driven_foundation_pile', 'stelcor_drilled_displacement_pile', 'cfa_pile', 'pile_cap', 'strip_footing', 'isolated_footing', 'pilaster', 'grade_beam', 'tie_beam', 'strap_beam', 'thickened_slab', 'buttress_takeoff', 'buttress_final', 'pier', 'corbel', 'linear_wall', 'foundation_wall', 'retaining_wall', 'barrier_wall', 'stem_wall', 'elevator_pit', 'service_elevator_pit', 'detention_tank', 'duplex_sewage_ejector_pit', 'deep_sewage_ejector_pit', 'sump_pump_pit', 'grease_trap', 'house_trap', 'mat_slab', 'mud_slab_foundation', 'sog', 'rog', 'stairs_on_grade', 'electric_conduit'].includes(itemType)) {
          try {
            const foundationFormulas = generateFoundationFormulas(itemType, row, parsedData || formulaInfo)
            if (foundationFormulas.takeoff) spreadsheet.updateCell({ formula: `=${foundationFormulas.takeoff}` }, `C${row}`)
            // Only override Length (F) if a non-null formula/value is provided
            if (foundationFormulas.length != null) {
              if (typeof foundationFormulas.length === 'string') {
                // Check if it's a formula reference (starts with G, I, etc.) or a fraction like "11/12"
                if (foundationFormulas.length.match(/^[A-Z]\d+$/)) {
                  // It's a cell reference like "G733" - use as formula
                  spreadsheet.updateCell({ formula: `=${foundationFormulas.length}` }, `F${row}`)
                } else {
                  // It's a formula string like "11/12" - use as formula
                  spreadsheet.updateCell({ formula: `=${foundationFormulas.length}` }, `F${row}`)
                }
              } else {
                spreadsheet.updateCell({ value: foundationFormulas.length }, `F${row}`)
              }
            }
            // Only override Width (G) if a non-null formula/value is provided
            if (foundationFormulas.width != null) {
              if (typeof foundationFormulas.width === 'string') {
                spreadsheet.updateCell({ formula: `=${foundationFormulas.width}` }, `G${row}`)
              } else {
                spreadsheet.updateCell({ value: foundationFormulas.width }, `G${row}`)
              }
            }
            // Only override Height (H) if a non-null formula/value is provided
            if (foundationFormulas.height != null) {
              if (typeof foundationFormulas.height === 'string') {
                spreadsheet.updateCell({ formula: `=${foundationFormulas.height}` }, `H${row}`)
              } else {
                spreadsheet.updateCell({ value: foundationFormulas.height }, `H${row}`)
              }
            }
            if (foundationFormulas.qty !== undefined) {
              if (typeof foundationFormulas.qty === 'string') {
                spreadsheet.updateCell({ formula: `=${foundationFormulas.qty}` }, `E${row}`)
              } else {
                spreadsheet.updateCell({ value: foundationFormulas.qty }, `E${row}`)
              }
            }
            // For single diameter drilled foundation pile, ft formula (H*C) goes to I
            if (itemType === 'drilled_foundation_pile' && !(parsedData || formulaInfo)?.parsed?.isDualDiameter) {
              // Single diameter: I=H*C
              if (foundationFormulas.ft) {
                spreadsheet.updateCell({ formula: `=${foundationFormulas.ft}` }, `I${row}`)
              } else {
                spreadsheet.updateCell({ value: '' }, `I${row}`)
              }
              // J column should be empty for single diameter
              spreadsheet.updateCell({ value: '' }, `J${row}`)
            } else {
              // For other items, ft goes to I, sqFt goes to J
              if (foundationFormulas.ft) {
                spreadsheet.updateCell({ formula: `=${foundationFormulas.ft}` }, `I${row}`)
              } else {
                spreadsheet.updateCell({ value: '' }, `I${row}`)
              }
              if (foundationFormulas.sqFt) {
                spreadsheet.updateCell({ formula: `=${foundationFormulas.sqFt}` }, `J${row}`)
              } else {
                spreadsheet.updateCell({ value: '' }, `J${row}`)
              }
              // Handle sqFt2 for dual diameter drilled foundation piles (column J)
              if (foundationFormulas.sqFt2) {
                spreadsheet.updateCell({ formula: `=${foundationFormulas.sqFt2}` }, `J${row}`)
              }
            }
            if (foundationFormulas.lbs) {
              spreadsheet.updateCell({ formula: `=${foundationFormulas.lbs}` }, `K${row}`)
            } else {
              spreadsheet.updateCell({ value: '' }, `K${row}`)
            }
            if (foundationFormulas.cy) {
              spreadsheet.updateCell({ formula: `=${foundationFormulas.cy}` }, `L${row}`)
            } else {
              spreadsheet.updateCell({ value: '' }, `L${row}`)
            }
            if (foundationFormulas.qtyFinal) {
              spreadsheet.updateCell({ formula: `=${foundationFormulas.qtyFinal}` }, `M${row}`)
            }

            // Apply red color to item names
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)

            // Special formatting for buttress takeoff row - strikethrough
            if (itemType === 'buttress_takeoff') {
              spreadsheet.cellFormat({ textDecoration: 'line-through' }, `B${row}:M${row}`)
            }

            // Special formatting for buttress final row - reference M to C of takeoff row; I, J, L, M red
            if (itemType === 'buttress_final') {
              if (formulaInfo.buttressRow) {
                spreadsheet.updateCell({ formula: `=C${formulaInfo.buttressRow}` }, `M${row}`)
              }
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `I${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `M${row}`)
            }

            // Special formatting for Elevator Pit - Sump pit row (J, L, M red)
            if (
              (itemType === 'elevator_pit' || itemType === 'service_elevator_pit') &&
              (parsedData || formulaInfo)?.parsed?.itemSubType === 'sump_pit'
            ) {
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `M${row}`)
            }

            // Special formatting for Mud Slab - J and L red
            if (itemType === 'mud_slab_foundation') {
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
            }

            // Special formatting for Stairs on grade - Stair slab and Landings J and L red
            if (itemType === 'stairs_on_grade') {
              const subType = (parsedData || formulaInfo)?.parsed?.itemSubType
              if (subType === 'stair_slab' || subType === 'landings') {
                spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
                spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
              }
            }

            // Special formatting for Electric conduit - Trench drain and Perforated pipe column I red
            if (itemType === 'electric_conduit') {
              const particulars = (parsedData || formulaInfo)?.particulars || ''
              const p = String(particulars).toLowerCase()
              if (p.includes('trench drain') || p.includes('perforated pipe')) {
                spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `I${row}`)
              }
            }
          } catch (error) {
            console.error(`Error applying Foundation formula at row ${row}:`, error)
          }
          return
        }

        if (itemType === 'foundation_sum') {
          const { firstDataRow, lastDataRow, subsectionName, isDualDiameter, excludeISum, excludeJSum, excludeKSum, matSumOnly, cySumOnly, firstDataRowForL, lastDataRowForL, lSumRange, matSlabCombinedSum, lastDataRowForJ } = formulaInfo
          try {
            // Mat slab: J (mat only) + L (mat+haunch) on same row - no gap in column L
            if (matSlabCombinedSum && subsectionName === 'Mat slab') {
              const jEndRow = lastDataRowForJ != null ? lastDataRowForJ : lastDataRow
              spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${jEndRow})` }, `J${row}`)
              spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
              return
            }

            // Sum for FT (I) - exclude if excludeISum is true (for slab items)
            if (!excludeISum) {
              const ftSumSubsections = ['Piles', 'Helical foundation pile', 'Driven foundation pile', 'Stelcor drilled displacement pile', 'CFA pile', 'Grade beams', 'Tie beam', 'Strap beams', 'Thickened slab', 'Corbel', 'Linear Wall', 'Foundation Wall', 'Retaining walls', 'Barrier wall', 'Drilled foundation pile', 'Strip Footings', 'Stem wall', 'Elevator Pit', 'Service elevator pit', 'Detention tank', 'Duplex sewage ejector pit', 'Deep sewage ejector pit', 'Sump pump pit', 'Grease trap', 'House trap', 'SOG', 'Ramp on grade', 'Stairs on grade Stairs', 'Electric conduit']
              if (ftSumSubsections.includes(subsectionName)) {
                spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
                spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `I${row}`)
              }
            }

            // Sum for SQ FT (I and J) - for drilled foundation pile dual diameter
            if (subsectionName === 'Drilled foundation pile' && isDualDiameter) {
              spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
              spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `I${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
            }

            // Sum for SQ FT (J) - exclude Drilled foundation pile (single diameter)
            // For Mat slab, only sum J for mat items (not haunch)
            if (!excludeJSum && !cySumOnly) {
              const sqFtSubsections = ['Piles', 'Pile caps', 'Isolated Footings', 'Pilaster', 'Pier', 'Strip Footings', 'Grade beams', 'Tie beam', 'Strap beams', 'Thickened slab', 'Corbel', 'Linear Wall', 'Foundation Wall', 'Retaining walls', 'Barrier wall', 'Stem wall', 'Elevator Pit', 'Service elevator pit', 'Detention tank', 'Duplex sewage ejector pit', 'Deep sewage ejector pit', 'Sump pump pit', 'Grease trap', 'House trap', 'Mat slab', 'SOG', 'Ramp on grade', 'Stairs on grade Stairs']
              if (sqFtSubsections.includes(subsectionName)) {
                spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
                spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
              }
            }

            // For Drilled foundation pile single diameter, J should be empty (no sum)
            if (subsectionName === 'Drilled foundation pile' && !isDualDiameter) {
              // J column should be empty - do nothing
            }

            // Sum for LBS (K)
            const lbsSubsections = ['Piles', 'Drilled foundation pile', 'Helical foundation pile', 'Driven foundation pile', 'Stelcor drilled displacement pile']
            if (!excludeKSum && lbsSubsections.includes(subsectionName)) {
              spreadsheet.updateCell({ formula: `=SUM(K${firstDataRow}:K${lastDataRow})` }, `K${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `K${row}`)
            }

            // Sum for QTY (M)
            const qtySubsections = ['Piles', 'Drilled foundation pile', 'Helical foundation pile', 'Driven foundation pile', 'Stelcor drilled displacement pile', 'CFA pile', 'Pile caps', 'Isolated Footings', 'Pilaster', 'Pier', 'Stairs on grade Stairs']
            if (qtySubsections.includes(subsectionName)) {
              spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `M${row}`)
            }

            // Sum for CY (L)
            // When lSumRange is provided (Stairs on grade Stairs), use it so landings L is included
            const lStartRow = firstDataRowForL != null ? firstDataRowForL : firstDataRow
            const lEndRow = lastDataRowForL != null ? lastDataRowForL : lastDataRow
            // For Mat slab with cySumOnly, sum L includes both mat and haunch
            if (cySumOnly) {
              // Only sum L (CY) for mat + haunch combined
              spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
              spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
            } else if (!formulaInfo.excludeLSum) {
              const cySubsections = ['Piles', 'Pile caps', 'Strip Footings', 'Isolated Footings', 'Pilaster', 'Grade beams', 'Tie beam', 'Strap beams', 'Thickened slab', 'Pier', 'Corbel', 'Linear Wall', 'Foundation Wall', 'Retaining walls', 'Barrier wall', 'Stem wall', 'Elevator Pit', 'Service elevator pit', 'Detention tank', 'Duplex sewage ejector pit', 'Deep sewage ejector pit', 'Sump pump pit', 'Grease trap', 'House trap', 'Mat slab', 'SOG', 'Ramp on grade', 'Stairs on grade Stairs']
              if (cySubsections.includes(subsectionName)) {
                const lFormula = lSumRange ? `=SUM(${lSumRange})` : `=SUM(L${lStartRow}:L${lEndRow})`
                deferredFoundationSumL.push({ row, lFormula })
                spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `L${row}`)
              }
            }
          } catch (error) {
            console.error(`Error applying Foundation sum formula at row ${row}:`, error)
          }
          return
        }
      } else if (section === 'waterproofing') {
        if (itemType === 'waterproofing_exterior_side_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
          } catch (error) {
            console.error(`Error applying Waterproofing Exterior side sum at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'waterproofing_exterior_side') {
          const waterproofingFormulas = generateWaterproofingFormulas(itemType, row, parsedData || formulaInfo)
          try {
            if (waterproofingFormulas.ft) {
              spreadsheet.updateCell({ formula: `=${waterproofingFormulas.ft}` }, `I${row}`)
            }
            if (waterproofingFormulas.height != null && waterproofingFormulas.height !== '') {
              spreadsheet.updateCell({ value: waterproofingFormulas.height }, `H${row}`)
            }
            if (waterproofingFormulas.sqFt) {
              spreadsheet.updateCell({ formula: `=${waterproofingFormulas.sqFt}` }, `J${row}`)
            }
          } catch (error) {
            console.error(`Error applying Waterproofing Exterior side formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'waterproofing_exterior_side_pit') {
          const foundationSlabRow = formulaInfo.foundationSlabRow
          const waterproofingFormulas = generateWaterproofingFormulas(itemType, row, parsedData || formulaInfo, { foundationSlabRow })
          try {
            if (waterproofingFormulas.ft) {
              spreadsheet.updateCell({ formula: `=${waterproofingFormulas.ft}` }, `I${row}`)
            }
            if (waterproofingFormulas.heightFormula) {
              spreadsheet.updateCell({ formula: `=${waterproofingFormulas.heightFormula}` }, `H${row}`)
            }
            if (waterproofingFormulas.sqFt) {
              spreadsheet.updateCell({ formula: `=${waterproofingFormulas.sqFt}` }, `J${row}`)
            }
            if (waterproofingFormulas.cy) {
              spreadsheet.updateCell({ formula: `=${waterproofingFormulas.cy}` }, `L${row}`)
            }
          } catch (error) {
            console.error(`Error applying Waterproofing Exterior side pit formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'waterproofing_negative_side_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
          } catch (error) {
            console.error(`Error applying Waterproofing Negative side sum at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'waterproofing_negative_side_wall') {
          const waterproofingFormulas = generateWaterproofingFormulas(itemType, row, parsedData || formulaInfo)
          try {
            if (waterproofingFormulas.ft) {
              spreadsheet.updateCell({ formula: `=${waterproofingFormulas.ft}` }, `I${row}`)
            }
            if (waterproofingFormulas.height != null && waterproofingFormulas.height !== '') {
              spreadsheet.updateCell({ value: waterproofingFormulas.height }, `H${row}`)
            }
            if (waterproofingFormulas.sqFt) {
              spreadsheet.updateCell({ formula: `=${waterproofingFormulas.sqFt}` }, `J${row}`)
            }
          } catch (error) {
            console.error(`Error applying Waterproofing Negative side wall formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'waterproofing_negative_side_slab') {
          const waterproofingFormulas = generateWaterproofingFormulas(itemType, row, parsedData || formulaInfo)
          try {
            if (waterproofingFormulas.sqFt) {
              spreadsheet.updateCell({ formula: `=${waterproofingFormulas.sqFt}` }, `J${row}`)
            }
          } catch (error) {
            console.error(`Error applying Waterproofing Negative side slab formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'waterproofing_horizontal_wp') {
          try {
            spreadsheet.updateCell({ value: 0 }, `C${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
          } catch (error) {
            console.error(`Error applying Waterproofing Horizontal WP formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'waterproofing_horizontal_wp_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
          } catch (error) {
            console.error(`Error applying Waterproofing Horizontal WP sum at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'waterproofing_horizontal_insulation') {
          try {
            spreadsheet.updateCell({ value: 0 }, `C${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
          } catch (error) {
            console.error(`Error applying Waterproofing Horizontal insulation formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'waterproofing_horizontal_insulation_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
          } catch (error) {
            console.error(`Error applying Waterproofing Horizontal insulation sum at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'waterproofing_extra_sqft') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
          } catch (error) {
            console.error(`Error applying Waterproofing extra SQ FT formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'waterproofing_extra_ft') {
          try {
            spreadsheet.updateCell({ formula: `=H${row}*C${row}` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'bold' }, `J${row}`)
          } catch (error) {
            console.error(`Error applying Waterproofing extra FT formula at row ${row}:`, error)
          }
          return
        }
      } else {
        formulas = generateDemolitionFormulas(itemType, row, parsedData)
      }

      try {
        // Column E (QTY)
        if (formulas.qty) {
          spreadsheet.updateCell({ formula: `=${formulas.qty}` }, `E${row}`)
        }

        // Column I (FT)
        if (formulas.ft) {
          spreadsheet.updateCell({ formula: `=${formulas.ft}` }, `I${row}`)
          if (itemType === 'line_drilling') {
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          }
        }

        // Column J (SQ FT)
        if (formulas.sqFt) {
          spreadsheet.updateCell({ formula: `=${formulas.sqFt}` }, `J${row}`)
        }

        // Column K (LBS or CY for excavation)
        if (formulas.lbs) {
          spreadsheet.updateCell({ formula: `=${formulas.lbs}` }, `K${row}`)

          // Apply strikethrough to excavation subsection CY values in column K (not backfill)
          if (section === 'excavation' && parsedData.subsection !== 'backfill') {
            spreadsheet.cellFormat(
              { textDecoration: 'line-through' },
              `K${row}`
            )
          }
        }

        // Column L (CY or 1.3*CY for excavation)
        if (formulas.cy) {
          spreadsheet.updateCell({ formula: `=${formulas.cy}` }, `L${row}`)
        }

        // Column M (QTY)
        if (formulas.qtyFinal) {
          spreadsheet.updateCell({ formula: `=${formulas.qtyFinal}` }, `M${row}`)
        }
      } catch (error) {
        // Ignore errors
      }
    })

    // Apply foundation_sum L formulas after all data-row formulas (so landings L etc. are set first)
    deferredFoundationSumL.forEach(({ row, lFormula }) => {
      try {
        spreadsheet.updateCell({ formula: lFormula }, `L${row}`)
      } catch (e) {
        console.error(`Error applying deferred Foundation sum L at row ${row}:`, e)
      }
    })

    // Apply formatting
    try {
      // Format header row - only Estimate column (A) has yellow background
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          textAlign: 'center',
          backgroundColor: '#FFFF00',
          color: '#000000'
        },
        'A1'
      )
      // Format other header columns with no background
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          textAlign: 'center',
          color: '#000000'
        },
        'B1:M1'
      )

      // Format section headers (find rows where column A has content and B is empty)
      calculationData.forEach((row, rowIndex) => {
        const rowNum = rowIndex + 1
        if (row[0] && !row[1]) {
          // This is a section header
          const sectionName = String(row[0])
          let backgroundColor = '#F4B084'
          if (sectionName === 'Demolition') {
            backgroundColor = '#E5B7AF'
          } else if (sectionName === 'Excavation') {
            backgroundColor = '#C6E0B4'
          } else if (sectionName === 'Rock Excavation') {
            backgroundColor = '#C6E0B4'
          } else if (sectionName === 'SOE') {
            backgroundColor = '#C6E0B4'
          } else if (sectionName === 'Foundation') {
            backgroundColor = '#C6E0B4'
          } else if (sectionName === 'Waterproofing') {
            backgroundColor = '#C6E0B4'
          }
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              backgroundColor: backgroundColor,
              fontSize: '11pt'
            },
            `A${rowNum}:M${rowNum}`
          )
          // Foundation header: sum (C) and CY (D) should not be bold
          if (sectionName === 'Foundation') {
            spreadsheet.cellFormat(
              { fontWeight: 'normal', backgroundColor: backgroundColor, fontSize: '11pt' },
              `C${rowNum}:D${rowNum}`
            )
          }
        }
        // Format Ele, Gas, Water sub-subsection headers (column A and B bold)
        if (row[0] && row[1]) {
          const aContent = String(row[0]).trim()
          const bContent = String(row[1])
          if ((aContent === 'Ele' || aContent === 'Gas' || aContent === 'Water') && /^\s*(Excavation|Backfill|Gravel):/.test(bContent)) {
            spreadsheet.cellFormat({ fontWeight: 'bold', fontStyle: 'italic' }, `A${rowNum}`)
            spreadsheet.cellFormat({ fontWeight: 'bold', fontStyle: 'italic' }, `B${rowNum}`)
          } else if (aContent === 'Concrete') {
            spreadsheet.cellFormat({ fontWeight: 'bold', fontStyle: 'italic' }, `A${rowNum}`)
          }
        }
        // Format subsection and sub-subsection headers (column B has content ending with ':' or starting with spaces)
        if (!row[0] && row[1]) {
          const bContent = String(row[1])
          // Check if it's a subsection/sub-subsection header (ends with ':' or has indentation)
          if (bContent.endsWith(':') || bContent.startsWith('  ')) {
            // Special formatting for extra line item headers
            if (bContent.includes('For demo Extra line item use this') ||
              bContent.includes('For Backfill Extra line item use this') ||
              bContent.includes('For soil excavation Extra line item use this') ||
              bContent.includes('For rock excavation Extra line item use this') ||
              bContent.includes('For foundation Extra line item use this')) {
              // Format only column B with yellow background
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  fontStyle: 'italic',
                  backgroundColor: '#FFFF00'
                },
                `B${rowNum}`
              )
            } else if (bContent.includes('Line drill:')) {
              // Line drill subsection header color
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  fontStyle: 'italic',
                  backgroundColor: '#E2EFDA'
                },
                `B${rowNum}`
              )
            } else if (bContent.includes('Underpinning:')) {
              // Underpinning subsection header color
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  fontStyle: 'italic',
                  backgroundColor: '#E2EFDA'
                },
                `B${rowNum}`
              )
            } else {
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  fontStyle: 'italic'
                },
                `B${rowNum}`
              )
            }
          } else {
            // Format data items in column B with red color (not headers/subsections or Havg)
            if (bContent !== 'Havg') {
              spreadsheet.cellFormat(
                {
                  color: '#FF0000'
                },
                `B${rowNum}`
              )
            }
          }
        }
      })

      // Center align the Estimate column (A)
      spreadsheet.cellFormat({ textAlign: 'center' }, 'A:A')

      // Format numeric values: whole numbers without decimals, others up to 2 decimal places (display only; calculations use full precision)
      const calcSheetRange = 'Calculations Sheet!A1:M1000'
      try {
        spreadsheet.numberFormat('0.##', calcSheetRange)
      } catch (e) {
        try { spreadsheet.cellFormat({ format: '0.##' }, calcSheetRange) } catch (e2) { }
      }
    } catch (error) {
      // Ignore errors
    }

    // Restore whatever sheet the user was on (so we don't force-switch them).
    try {
      if (prevSheetName) {
        const topLeft = typeof prevSelectedRange === 'string' && prevSelectedRange.includes(':')
          ? prevSelectedRange.split(':')[0]
          : (prevSelectedRange || 'A1')
        spreadsheet.goTo(`${prevSheetName}!${topLeft}`)
      }
    } catch (e) {
      // ignore
    }
  }

  const applyProposalFirstRow = (spreadsheet) => {
    // Full Proposal template (match screenshot layout).
    const proposalSheetIndex = 1
    const pfx = 'Proposal Sheet!'

    // Clear everything (values + formats) on Proposal sheet so it never retains other sheet data.
    try {
      spreadsheet.clear({ range: `${pfx}A1:Z1000`, type: 'Clear All' })
    } catch (e) {
      try { spreadsheet.clear({ range: `${pfx}A1:Z1000` }) } catch (e2) { /* ignore */ }
    }

    // Column widths to mirror screenshot
    const colWidths = {
      A: 40,  // margin
      B: 700, // main left block (wider per request)
      C: 75,  // LF
      D: 75,  // SF
      E: 85,  // LBS
      F: 85,  // CY
      G: 85,  // QTY
      H: 70,  // $/1000 spacer
      I: 65,  // LF
      J: 65,  // SF
      K: 75,  // LBS
      L: 75,  // CY
      M: 75,  // QTY
      N: 75   // LS
    }
    Object.entries(colWidths).forEach(([col, width]) => {
      spreadsheet.setColWidth(width, col.charCodeAt(0) - 65, proposalSheetIndex)
    })

    // Helper function to calculate row height based on text content in column B
    const calculateRowHeight = (text) => {
      if (!text || typeof text !== 'string') {
        return 22 // Default height for empty cells
      }

      // Column B width is approximately 500 pixels (based on colWidths)
      const columnWidth = 500
      const fontSize = 11 // Default font size
      const lineHeight = 14 // More accurate line height in pixels
      const padding = 8 // Top and bottom padding
      const charWidth = 6.5 // More accurate character width in pixels (for font size 11)

      // Calculate how many characters fit per line (accounting for padding)
      const availableWidth = columnWidth - 20 // Account for cell padding
      const charsPerLine = Math.floor(availableWidth / charWidth)

      // Calculate number of lines needed (accounting for word wrapping)
      const textLength = text.length
      const estimatedLines = Math.max(1, Math.ceil(textLength / charsPerLine))

      // Calculate height: (lines * lineHeight) + padding
      const calculatedHeight = Math.max(22, Math.ceil(estimatedLines * lineHeight + padding))

      // Cap at reasonable maximum (60px for most cases, 80px for very long text)
      // This prevents excessively tall rows
      return Math.min(calculatedHeight, 80)
    }


    // Styles
    const headerGray = { backgroundColor: '#D9D9D9', fontWeight: 'bold', textAlign: 'center', verticalAlign: 'middle' }
    const boxGray = { backgroundColor: '#D9D9D9', fontWeight: 'bold', verticalAlign: 'middle' }
    const thick = { border: '2px solid #000000' }
    const thin = { border: '1px solid #000000' }

      // Top header row (row 1)
      ;[
        ['C1', 'LF'], ['D1', 'SF'], ['E1', 'LBS'], ['F1', 'CY'], ['G1', 'QTY'],
        ['H1', '$/1000'],
        ['I1', 'LF'], ['J1', 'SF'], ['K1', 'LBS'], ['L1', 'CY'], ['M1', 'QTY'], ['N1', 'LS']
      ].forEach(([cell, value]) => spreadsheet.updateCell({ value }, `${pfx}${cell}`))
    spreadsheet.cellFormat(headerGray, `${pfx}C1:N1`)
    spreadsheet.cellFormat(thick, `${pfx}B1:N1`)
    // Make $/1000 cell background white
    spreadsheet.cellFormat({ backgroundColor: 'white' }, `${pfx}H1`)

    // Main outer frame around top content block (PERIMETER ONLY).
    // Avoid applying a border to the full range, since that creates thick internal grid lines.
    spreadsheet.cellFormat(thick, `${pfx}B3:G3`)     // top edge
    spreadsheet.cellFormat(thick, `${pfx}B11:G11`)   // bottom edge
    spreadsheet.cellFormat(thick, `${pfx}B3:B11`)    // left edge
    spreadsheet.cellFormat(thick, `${pfx}G3:G11`)    // right edge

    // Per request: rows 3 to 8, columns B to E should look like "no border configured"
    // (i.e., rely on default grid lines), and text should be normal weight in black.
    try {
      spreadsheet.clear({ range: `${pfx}B3:E8`, type: 'Clear Formats' })
    } catch (e) {
      // ignore (different versions may have different clear type strings)
    }
    spreadsheet.cellFormat({ fontWeight: 'normal', color: '#000000' }, `${pfx}B3:E8`)

    // Left address block
    spreadsheet.updateCell({ value: '37-24 24th Street, Suite 132, Long Island City, NY 11101' }, `${pfx}B4`)
    spreadsheet.updateCell({ value: 'Tel: 718 726-1525 | Fax: 718 726-1601 | Cell: 917 600-3958' }, `${pfx}B5`)
    spreadsheet.updateCell({ value: 'Email:' }, `${pfx}B6`)
    spreadsheet.cellFormat({ fontWeight: 'bold' }, `${pfx}B4:B5`)
    spreadsheet.cellFormat({ color: '#0B76C3', textDecoration: 'underline' }, `${pfx}B6`)
    spreadsheet.cellFormat(thin, `${pfx}F4:G6`)
    // Remove borders from B4:E8 - ensure no border color
    try {
      spreadsheet.clear({ range: `${pfx}B4:E8`, type: 'Clear Formats' })
    } catch (e) {
      // ignore
    }
    // Reapply formatting without borders
    spreadsheet.cellFormat({ fontWeight: 'normal', color: '#000000' }, `${pfx}B4:E8`)
    spreadsheet.cellFormat({ fontWeight: 'bold' }, `${pfx}B4:B5`)
    spreadsheet.cellFormat({ color: '#0B76C3', textDecoration: 'underline' }, `${pfx}B6`)

    // Logo (image) near center top-left; uses production URL for server-side save compatibility
    try {
      // Use production URL so Syncfusion server-side save can fetch the image via HTTP
      // The save server runs on production and cannot reach localhost
      const imgSrc = 'https://krmestimators.com/images/templateimage.png'
      // Start around mid of column B and extend roughly to column E
      // (B width is large; left offset pushes the image start into the middle of B)
      spreadsheet.insertImage(
        [{ src: imgSrc, width: 460, height: 110, left: 450, top: 50 }],
        `${pfx}B4`
      )
    } catch (e) {
      // ignore
    }

    // Date/Project/Client block
    spreadsheet.updateCell({ value: "Date: Today's date" }, `${pfx}B9`)
    spreadsheet.updateCell({ value: 'Project: ###' }, `${pfx}B10`)
    spreadsheet.updateCell({ value: 'Client: ###' }, `${pfx}B11`)
    spreadsheet.cellFormat({ fontWeight: 'bold' }, `${pfx}B9:B11`)
    spreadsheet.cellFormat(thin, `${pfx}B9:G11`)

    // Right grey box (Estimate / Drawings Dated / lines) - start from column F
    // Individual cell styling - customize each cell separately

    // Row 3: Estimate #25-1150
    spreadsheet.merge(`${pfx}F3:G3`)
    spreadsheet.updateCell({ value: 'Estimate #25-1150' }, `${pfx}F3`)
    spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontSize: '11pt', fontWeight: 'bold', borderTop: '1px solid #000000', borderLeft: '1px solid #000000', borderRight: '1px solid #000000' }, `${pfx}F3`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontSize: '11pt', fontWeight: 'bold', borderTop: '1px solid #000000', borderRight: '1px solid #000000' }, `${pfx}G3`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontSize: '11pt', fontWeight: 'normal' }, `${pfx}H3`)

    // Row 4: Empty row
    spreadsheet.merge(`${pfx}F4:G4`)
    spreadsheet.cellFormat({ backgroundColor: 'white', borderLeft: '1px solid #000000' }, `${pfx}F4`)
    spreadsheet.cellFormat({ backgroundColor: 'white', borderRight: '1px solid #000000' }, `${pfx}G4`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal' }, `${pfx}H4`)

    // Row 5: Drawings Dated:
    spreadsheet.merge(`${pfx}F5:G5`)
    spreadsheet.updateCell({ value: 'Drawings Dated:' }, `${pfx}F5`)
    spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '1px solid #000000', borderTop: '1px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}F5`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'bold', borderRight: '1px solid #000000', borderTop: '1px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}G5`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal' }, `${pfx}H5`)

    // Row 6: SOE:
    spreadsheet.merge(`${pfx}F6:G6`)
    spreadsheet.updateCell({ value: 'SOE:' }, `${pfx}F6`)
    spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '1px solid #000000', borderTop: '1px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}F6`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'bold', borderRight: '1px solid #000000', borderTop: '1px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}G6`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal' }, `${pfx}H6`)

    // Row 7: Structural:
    spreadsheet.merge(`${pfx}F7:G7`)
    spreadsheet.updateCell({ value: 'Structural:' }, `${pfx}F7`)
    spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '1px solid #000000', borderTop: '1px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}F7`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'bold', borderRight: '1px solid #000000', borderTop: '1px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}G7`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal' }, `${pfx}H7`)

    // Row 8: Architectural:
    spreadsheet.merge(`${pfx}F8:G8`)
    spreadsheet.updateCell({ value: 'Architectural:' }, `${pfx}F8`)
    spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '1px solid #000000', borderTop: '1px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}F8`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'bold', borderRight: '1px solid #000000', borderTop: '1px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}G8`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal' }, `${pfx}H8`)

    // Row 9: Plumbing:
    spreadsheet.merge(`${pfx}F9:G9`)
    spreadsheet.updateCell({ value: 'Plumbing:' }, `${pfx}F9`)
    spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '1px solid #000000', borderTop: '1px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}F9`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'bold', borderRight: '1px solid #000000', borderTop: '1px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}G9`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal' }, `${pfx}H9`)

    // Row 10: Mechanical
    spreadsheet.merge(`${pfx}F10:G10`)
    spreadsheet.updateCell({ value: 'Mechanical' }, `${pfx}F10`)
    spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderBottom: '1px solid #000000', borderLeft: '1px solid #000000' }, `${pfx}F10`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'bold', borderBottom: '1px solid #000000', borderRight: '1px solid #000000' }, `${pfx}G10`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal' }, `${pfx}H10`)

    // Bottom header row (row 12)
    spreadsheet.updateCell({ value: 'DESCRIPTION' }, `${pfx}B12`)
      ;[
        ['C12', 'LF'], ['D12', 'SF'], ['E12', 'LBS'], ['F12', 'CY'], ['G12', 'QTY'],
        ['H12', '$/1000'],
        ['I12', 'LF'], ['J12', 'SF'], ['K12', 'LBS'], ['L12', 'CY'], ['M12', 'QTY']
      ].forEach(([cell, value]) => spreadsheet.updateCell({ value }, `${pfx}${cell}`))
    spreadsheet.updateCell({ value: 'LS' }, `${pfx}N12`)

    spreadsheet.cellFormat(headerGray, `${pfx}B12:N12`)
    spreadsheet.cellFormat(thick, `${pfx}B12:N12`)
    // Make $/1000 cell background white
    spreadsheet.cellFormat({ backgroundColor: 'white' }, `${pfx}H12`)

    // Row 14: Demolition scope
    spreadsheet.updateCell({ value: 'Demolition scope:' }, `${pfx}B14`)
    spreadsheet.cellFormat({
      backgroundColor: '#BDD7EE',
      textAlign: 'center',
      verticalAlign: 'middle',
      textDecoration: 'underline',
      fontWeight: 'normal',
      border: '1px solid #000000'
    }, `${pfx}B14`)

    // Demolition scope lines from Calculations Sheet:
    // For each Demolition subsection, take the first item description in column B
    // (e.g. "Allow to saw-cut/demo/remove/dispose existing (4\" thick) slab on grade @ existing building as per DM-106.00 & details on DM-107.00")
    // and show it under Demolition scope on the Proposal sheet.
    if (Array.isArray(calculationData) && calculationData.length > 0) {
      const linesBySubsection = new Map()
      const rowsBySubsection = new Map() // Track all data rows for each subsection to calculate SF
      const sumRowsBySubsection = new Map() // Track sum row for each subsection to get QTY
      let inDemolitionSection = false
      let inExcavationSection = false
      let inExcavationSubsection = false // Track if we're in the "excavation" subsection
      let inBackfillSubsection = false // Track if we're in the "backfill" subsection
      let currentSubsection = null
      let dataRowCount = 0 // Track number of data rows in current subsection
      let excavationTotalSQFT = 0 // Track total SQFT for excavation section
      let excavationRunningSum = 0 // Track running sum of calculated SQFT values
      let excavationTotalCY = 0 // Track total CY for excavation section
      let excavationRunningCYSum = 0 // Track running sum of calculated CY values
      let foundEmptyParticulars = false // Track if we've found the first empty Particulars row
      let excavationEmptyRowIndex = null // Store the row index where empty Particulars row was found
      let backfillTotalSQFT = 0 // Track total SQFT for backfill section
      let backfillRunningSum = 0 // Track running sum of calculated SQFT values for backfill
      let backfillTotalCY = 0 // Track total CY for backfill section
      let backfillRunningCYSum = 0 // Track running sum of calculated CY values for backfill
      let foundBackfillEmptyParticulars = false // Track if we've found the first empty Particulars row for backfill
      let backfillEmptyRowIndex = null // Store the row index where empty Particulars row was found for backfill
      let inRockExcavationSection = false // Track if we're in the Rock Excavation section
      let inRockExcavationSubsection = false // Track if we're in the "rock excavation" subsection
      let rockExcavationTotalSQFT = 0 // Track total SQFT for rock excavation section
      let rockExcavationRunningSum = 0 // Track running sum of calculated SQFT values for rock excavation
      let rockExcavationTotalCY = 0 // Track total CY for rock excavation section
      let rockExcavationRunningCYSum = 0 // Track running sum of calculated CY values for rock excavation
      let foundRockExcavationEmptyParticulars = false // Track if we've found the first empty Particulars row for rock excavation
      let rockExcavationEmptyRowIndex = null // Store the row index where empty Particulars row was found for rock excavation
      let inSOESection = false // Track if we're in the SOE section
      let inDrilledSoldierPileSubsection = false // Track if we're in the "Drilled soldier pile" subsection
      let drilledSoldierPileItems = [] // Collect drilled soldier pile items
      let inHPSoldierPileSubsection = false // Track if we're in the "HP" or "H-pile" subsection
      let hpSoldierPileItems = [] // Collect HP soldier pile items
      // Store SOE subsection items by subsection name
      window.soeSubsectionItems = new Map() // Map<subsectionName, items[]>
      let currentSOESubsectionItems = [] // Current subsection items being collected
      let currentSOESubsectionName = null // Current subsection name
      // Store Foundation subsection items by subsection name
      window.foundationSubsectionItems = new Map() // Map<subsectionName, items[]>
      let currentFoundationSubsectionItems = [] // Current subsection items being collected
      let currentFoundationSubsectionName = null // Current subsection name
      let inFoundationSection = false // Track if we're in the Foundation section

      // Initialize workbook early so we can read evaluated formula values
      let workbook = null
      let calcSheetIndex = 0
      try {
        workbook = spreadsheet.getWorkbook()
      } catch (e) {
        // Ignore errors - workbook might not be available yet
      }

      // Helper function to calculate SQFT for a single row (needed for real-time calculation)
      const calculateRowSQFT = (row) => {
        // First try to use column J (SQFT) if it exists and is a number
        const sqftValue = parseFloat(row[9]) || 0 // Column J (index 9)

        if (sqftValue > 0) {
          return sqftValue
        }

        // If column J is empty, calculate from takeoff, length, width
        const takeoff = parseFloat(row[2]) || 0 // Column C (index 2)
        const length = parseFloat(row[5]) || 0   // Column F (index 5)
        const width = parseFloat(row[6]) || 0    // Column G (index 6)

        if (takeoff > 0) {
          if (width > 0 && length > 0) {
            // Has both width and length: SQFT = takeoff * length * width
            return takeoff * length * width
          } else if (width > 0) {
            // Has width but no length: SQFT = takeoff * width
            return takeoff * width
          } else if (length > 0) {
            // Has length but no width: SQFT = takeoff * length
            return takeoff * length
          } else {
            // No width and no length: SQFT = takeoff
            return takeoff
          }
        }
        return 0
      }

      // Helper function to calculate CY for a single row
      const calculateRowCY = (row) => {
        // First try to use column L (CY) if it exists and is a number
        const cyValue = parseFloat(row[11]) || 0 // Column L (index 11)

        if (cyValue > 0) {
          return cyValue
        }

        // If column L is empty, calculate from SF and height: CY = SF * height / 27
        const sf = calculateRowSQFT(row)
        const height = parseFloat(row[7]) || 0 // Column H (index 7)

        if (sf > 0 && height > 0) {
          return (sf * height) / 27
        }

        return 0
      }

      calculationData.forEach((row, rowIndex) => {
        const colA = row[0]
        const colB = row[1]

        if (colA && String(colA).trim().toLowerCase() === 'demolition') {
          inDemolitionSection = true
          inExcavationSection = false
          currentSubsection = null
          dataRowCount = 0
          return
        }

        if (colA && String(colA).trim().toLowerCase() === 'excavation') {
          inDemolitionSection = false
          inExcavationSection = true
          inRockExcavationSection = false
          inExcavationSubsection = false
          inBackfillSubsection = false
          currentSubsection = null
          dataRowCount = 0
          return
        }

        if (colA && String(colA).trim().toLowerCase() === 'rock excavation') {
          inDemolitionSection = false
          inExcavationSection = false
          inRockExcavationSection = true
          inSOESection = false
          inExcavationSubsection = false
          inBackfillSubsection = false
          inRockExcavationSubsection = false
          inDrilledSoldierPileSubsection = false
          inHPSoldierPileSubsection = false
          currentSubsection = null
          dataRowCount = 0
          rockExcavationRunningSum = 0
          rockExcavationRunningCYSum = 0
          foundRockExcavationEmptyParticulars = false
          drilledSoldierPileItems = []
          hpSoldierPileItems = []
          return
        }

        if (colA && String(colA).trim().toLowerCase() === 'soe') {
          inDemolitionSection = false
          inExcavationSection = false
          inRockExcavationSection = false
          inSOESection = true
          inFoundationSection = false
          inExcavationSubsection = false
          inBackfillSubsection = false
          inRockExcavationSubsection = false
          inDrilledSoldierPileSubsection = false
          inHPSoldierPileSubsection = false
          currentSubsection = null
          dataRowCount = 0
          drilledSoldierPileItems = []
          hpSoldierPileItems = []
          window.soeSubsectionItems = new Map()
          currentSOESubsectionItems = []
          currentSOESubsectionName = null
          window.foundationSubsectionItems = new Map()
          currentFoundationSubsectionItems = []
          currentFoundationSubsectionName = null
          return
        }

        if (colA && String(colA).trim().toLowerCase() === 'foundation') {
          inDemolitionSection = false
          inExcavationSection = false
          inRockExcavationSection = false
          inSOESection = false
          inFoundationSection = true
          inExcavationSubsection = false
          inBackfillSubsection = false
          inRockExcavationSubsection = false
          inDrilledSoldierPileSubsection = false
          inHPSoldierPileSubsection = false
          currentSubsection = null
          dataRowCount = 0
          drilledSoldierPileItems = []
          hpSoldierPileItems = []
          window.soeSubsectionItems = new Map()
          currentSOESubsectionItems = []
          currentSOESubsectionName = null
          window.foundationSubsectionItems = new Map()
          currentFoundationSubsectionItems = []
          currentFoundationSubsectionName = null
          return
        }

        if (!inDemolitionSection && !inExcavationSection && !inRockExcavationSection && !inSOESection && !inFoundationSection) return

        // If we hit another main section header in column A, stop collecting.
        if (colA && String(colA).trim() &&
          String(colA).trim().toLowerCase() !== 'demolition' &&
          String(colA).trim().toLowerCase() !== 'excavation' &&
          String(colA).trim().toLowerCase() !== 'rock excavation' &&
          String(colA).trim().toLowerCase() !== 'soe' &&
          String(colA).trim().toLowerCase() !== 'foundation') {
          // Save current SOE subsection items if any
          if (inSOESection && currentSOESubsectionName && currentSOESubsectionItems.length > 0) {
            if (!window.soeSubsectionItems.has(currentSOESubsectionName)) {
              window.soeSubsectionItems.set(currentSOESubsectionName, [])
            }
            window.soeSubsectionItems.get(currentSOESubsectionName).push([...currentSOESubsectionItems])
          }

          // Save current Foundation subsection items if any
          if (inFoundationSection && currentFoundationSubsectionName && currentFoundationSubsectionItems.length > 0) {
            if (!window.foundationSubsectionItems.has(currentFoundationSubsectionName)) {
              window.foundationSubsectionItems.set(currentFoundationSubsectionName, [])
            }
            window.foundationSubsectionItems.get(currentFoundationSubsectionName).push([...currentFoundationSubsectionItems])
          }

          // Reset flags when moving to next section
          inDemolitionSection = false
          inExcavationSection = false
          inRockExcavationSection = false
          inSOESection = false
          inFoundationSection = false
          inExcavationSubsection = false
          inBackfillSubsection = false
          inRockExcavationSubsection = false
          inDrilledSoldierPileSubsection = false
          inHPSoldierPileSubsection = false
          currentSubsection = null
          dataRowCount = 0
          drilledSoldierPileItems = []
          hpSoldierPileItems = []
          currentSOESubsectionItems = []
          currentSOESubsectionName = null
          currentFoundationSubsectionItems = []
          currentFoundationSubsectionName = null
          return
        }

        const bText = colB ? String(colB).trim() : ''

        // Subsection header (ends with ':')
        if (bText.endsWith(':')) {
          currentSubsection = bText.slice(0, -1).trim()
          if (!rowsBySubsection.has(currentSubsection)) {
            rowsBySubsection.set(currentSubsection, [])
          }

          // Check if this is the "excavation" subsection within the Excavation section
          if (inExcavationSection && currentSubsection.toLowerCase() === 'excavation') {
            inExcavationSubsection = true
            excavationRunningSum = 0 // Reset running sum
            excavationRunningCYSum = 0 // Reset CY running sum
            foundEmptyParticulars = false // Reset flag
          } else if (inExcavationSubsection && inExcavationSection) {
            // We've moved to a different subsection, so the excavation subsection has ended
            inExcavationSubsection = false
          }

          // Check if this is the "backfill" subsection within the Excavation section
          if (inExcavationSection && currentSubsection.toLowerCase() === 'backfill') {
            inBackfillSubsection = true
            backfillRunningSum = 0 // Reset running sum
            backfillRunningCYSum = 0 // Reset CY running sum
            foundBackfillEmptyParticulars = false // Reset flag
          } else if (inBackfillSubsection && inExcavationSection) {
            // We've moved to a different subsection, so the backfill subsection has ended
            inBackfillSubsection = false
          }

          // Check if this is the "rock excavation" subsection within the Rock Excavation section
          if (inRockExcavationSection && currentSubsection.toLowerCase() === 'rock excavation') {
            inRockExcavationSubsection = true
            rockExcavationRunningSum = 0 // Reset running sum
            rockExcavationRunningCYSum = 0 // Reset CY running sum
            foundRockExcavationEmptyParticulars = false // Reset flag
          } else if (inRockExcavationSubsection && inRockExcavationSection) {
            // We've moved to a different subsection, so the rock excavation subsection has ended
            inRockExcavationSubsection = false
          }

          // Check if this is the "Drilled soldier pile" subsection within the SOE section
          if (inSOESection && currentSubsection.toLowerCase() === 'drilled soldier pile') {
            inDrilledSoldierPileSubsection = true
            drilledSoldierPileItems = [] // Reset items
            // Save previous subsection items if any
            if (currentSOESubsectionName && currentSOESubsectionItems.length > 0) {
              if (!window.soeSubsectionItems.has(currentSOESubsectionName)) {
                window.soeSubsectionItems.set(currentSOESubsectionName, [])
              }
              window.soeSubsectionItems.get(currentSOESubsectionName).push([...currentSOESubsectionItems])
              currentSOESubsectionItems = []
            }
            currentSOESubsectionName = null
          } else if (inDrilledSoldierPileSubsection && inSOESection) {
            // We've moved to a different subsection, so the drilled soldier pile subsection has ended
            // Save the last group if there are items
            if (drilledSoldierPileItems.length > 0) {
              if (!window.drilledSoldierPileGroups) {
                window.drilledSoldierPileGroups = []
              }
              window.drilledSoldierPileGroups.push([...drilledSoldierPileItems])
            }

            // Also save HP items if any were collected in this subsection
            if (hpSoldierPileItems && hpSoldierPileItems.length > 0) {
              if (!window.hpSoldierPileGroups) {
                window.hpSoldierPileGroups = []
              }
              window.hpSoldierPileGroups.push([...hpSoldierPileItems])
            }

            inDrilledSoldierPileSubsection = false
            drilledSoldierPileItems = []
            hpSoldierPileItems = []
          }

          // Check if this is the "HP" or "H-pile" subsection within the SOE section
          const hpSubsectionNames = ['hp', 'h-pile', 'hp soldier pile', 'hp pile']
          if (inSOESection && hpSubsectionNames.includes(currentSubsection.toLowerCase())) {
            inHPSoldierPileSubsection = true
            hpSoldierPileItems = [] // Reset items
            // Save previous subsection items if any
            if (currentSOESubsectionName && currentSOESubsectionItems.length > 0) {
              if (!window.soeSubsectionItems.has(currentSOESubsectionName)) {
                window.soeSubsectionItems.set(currentSOESubsectionName, [])
              }
              window.soeSubsectionItems.get(currentSOESubsectionName).push([...currentSOESubsectionItems])
              currentSOESubsectionItems = []
            }
            currentSOESubsectionName = null
          } else if (inHPSoldierPileSubsection && inSOESection) {
            // We've moved to a different subsection, so the HP soldier pile subsection has ended
            // Save the last group if there are items
            if (hpSoldierPileItems.length > 0) {
              if (!window.hpSoldierPileGroups) {
                window.hpSoldierPileGroups = []
              }
              window.hpSoldierPileGroups.push([...hpSoldierPileItems])
            }
            inHPSoldierPileSubsection = false
            hpSoldierPileItems = []
          }

          // Track other SOE subsections (not soldier piles)
          if (inSOESection && !inDrilledSoldierPileSubsection && !inHPSoldierPileSubsection) {
            const soldierPileSubsectionNames = ['drilled soldier pile', 'hp', 'h-pile', 'hp soldier pile', 'hp pile']
            if (!soldierPileSubsectionNames.includes(currentSubsection.toLowerCase())) {
              // Save previous subsection items if any
              if (currentSOESubsectionName && currentSOESubsectionName !== currentSubsection && currentSOESubsectionItems.length > 0) {
                if (!window.soeSubsectionItems.has(currentSOESubsectionName)) {
                  window.soeSubsectionItems.set(currentSOESubsectionName, [])
                }
                window.soeSubsectionItems.get(currentSOESubsectionName).push([...currentSOESubsectionItems])
                currentSOESubsectionItems = []
              }
              // Start collecting for new subsection
              currentSOESubsectionName = currentSubsection
              currentSOESubsectionItems = []

            }
          }

          // Track Foundation subsections
          if (inFoundationSection) {
            // Save previous subsection items if any
            if (currentFoundationSubsectionName && currentFoundationSubsectionName !== currentSubsection && currentFoundationSubsectionItems.length > 0) {
              if (!window.foundationSubsectionItems.has(currentFoundationSubsectionName)) {
                window.foundationSubsectionItems.set(currentFoundationSubsectionName, [])
              }
              window.foundationSubsectionItems.get(currentFoundationSubsectionName).push([...currentFoundationSubsectionItems])
              currentFoundationSubsectionItems = []
            }
            // Start collecting for new subsection
            currentFoundationSubsectionName = currentSubsection
            currentFoundationSubsectionItems = []
          }

          dataRowCount = 0
          return
        }

        // Collect drilled soldier pile items until empty row
        if (inDrilledSoldierPileSubsection && inSOESection) {
          const bText = colB ? String(colB).trim() : ''
          const isParticularsEmpty = !colB || bText === ''

          if (isParticularsEmpty) {
            // Found empty row - save current group and continue collecting if there are more items
            if (drilledSoldierPileItems.length > 0) {
              // Store current group
              if (!window.drilledSoldierPileGroups) {
                window.drilledSoldierPileGroups = []
              }
              window.drilledSoldierPileGroups.push([...drilledSoldierPileItems])
              // Reset for next group
              drilledSoldierPileItems = []
              // Continue collecting - don't set inDrilledSoldierPileSubsection to false
              // The next non-empty row will be part of the next group
            }

            // Also save HP items if any were collected in this subsection
            if (hpSoldierPileItems && hpSoldierPileItems.length > 0) {
              if (!window.hpSoldierPileGroups) {
                window.hpSoldierPileGroups = []
              }
              window.hpSoldierPileGroups.push([...hpSoldierPileItems])
              hpSoldierPileItems = []
            }

            return // Skip further processing for this row
          } else if (bText && !bText.endsWith(':')) {
            // Check if this is an HP item
            const isHPItem = /HP\d+x\d+/i.test(bText)

            // Collect HP items even if they're in the drilled soldier pile subsection
            if (isHPItem) {
              const takeoff = parseFloat(row[2]) || 0 // Column C (index 2)
              const unit = row[3] || '' // Column D (index 3)
              const height = parseFloat(row[7]) || 0 // Column H (index 7)

              hpSoldierPileItems.push({
                particulars: bText,
                takeoff: takeoff,
                unit: unit,
                height: height,
                rawRow: row,
                rawRowNumber: rowIndex + 2
              })
              return // Skip further processing for this row
            }

            // This is a drilled soldier pile data row - collect it
            const takeoff = parseFloat(row[2]) || 0 // Column C (index 2)
            const unit = row[3] || '' // Column D (index 3)
            const height = parseFloat(row[7]) || 0 // Column H (index 7)

            drilledSoldierPileItems.push({
              particulars: bText,
              takeoff: takeoff,
              unit: unit,
              height: height,
              rawRow: row,
              rawRowNumber: rowIndex + 2
            })
            return // Skip further processing for this row
          }
        }

        // Collect HP soldier pile items until empty row
        if (inHPSoldierPileSubsection && inSOESection) {
          const bText = colB ? String(colB).trim() : ''
          const isParticularsEmpty = !colB || bText === ''

          if (isParticularsEmpty) {
            // Found empty row - save current group and continue collecting if there are more items
            if (hpSoldierPileItems.length > 0) {
              // Store current group
              if (!window.hpSoldierPileGroups) {
                window.hpSoldierPileGroups = []
              }
              window.hpSoldierPileGroups.push([...hpSoldierPileItems])
              // Reset for next group
              hpSoldierPileItems = []
            }
            return // Skip further processing for this row
          } else if (bText && !bText.endsWith(':')) {
            // This is an HP data row - collect it
            const takeoff = parseFloat(row[2]) || 0 // Column C (index 2)
            const unit = row[3] || '' // Column D (index 3)
            const height = parseFloat(row[7]) || 0 // Column H (index 7)

            hpSoldierPileItems.push({
              particulars: bText,
              takeoff: takeoff,
              unit: unit,
              height: height,
              rawRow: row,
              rawRowNumber: rowIndex + 2
            })
            return // Skip further processing for this row
          }
        }

        // Collect items for other SOE subsections (not soldier piles)
        if (inSOESection && !inDrilledSoldierPileSubsection && !inHPSoldierPileSubsection && currentSOESubsectionName) {
          const bText = colB ? String(colB).trim() : ''
          const isParticularsEmpty = !colB || bText === ''


          if (isParticularsEmpty) {
            // Found empty row - save current group
            if (currentSOESubsectionItems.length > 0) {
              if (!window.soeSubsectionItems.has(currentSOESubsectionName)) {
                window.soeSubsectionItems.set(currentSOESubsectionName, [])
              }
              window.soeSubsectionItems.get(currentSOESubsectionName).push([...currentSOESubsectionItems])


              currentSOESubsectionItems = []
            }
            return // Skip further processing for this row
          } else if (bText && !bText.endsWith(':')) {
            // This is a data row - collect it
            const takeoff = parseFloat(row[2]) || 0 // Column C (index 2)
            const unit = row[3] || '' // Column D (index 3)
            const height = parseFloat(row[7]) || 0 // Column H (index 7)
            const sqft = parseFloat(row[9]) || 0 // Column J (index 9)
            const lbs = parseFloat(row[10]) || 0 // Column K (index 10)
            const qty = parseFloat(row[12]) || 0 // Column M (index 12)

            const item = {
              particulars: bText,
              takeoff: takeoff,
              unit: unit,
              height: height,
              sqft: sqft,
              lbs: lbs,
              qty: qty,
              rawRow: row,
              rawRowNumber: rowIndex + 2
            }

            currentSOESubsectionItems.push(item)


            return // Skip further processing for this row
          }
        }

        // Collect items for Foundation subsections
        if (inFoundationSection && currentFoundationSubsectionName) {
          const bText = colB ? String(colB).trim() : ''
          const isParticularsEmpty = !colB || bText === ''

          if (isParticularsEmpty) {
            // Found empty row - save current group
            if (currentFoundationSubsectionItems.length > 0) {
              if (!window.foundationSubsectionItems.has(currentFoundationSubsectionName)) {
                window.foundationSubsectionItems.set(currentFoundationSubsectionName, [])
              }
              window.foundationSubsectionItems.get(currentFoundationSubsectionName).push([...currentFoundationSubsectionItems])
              currentFoundationSubsectionItems = []
            }
            return // Skip further processing for this row
          } else if (bText && !bText.endsWith(':')) {
            // This is a data row - collect it
            const takeoff = parseFloat(row[2]) || 0 // Column C (index 2)
            const unit = row[3] || '' // Column D (index 3)
            const height = parseFloat(row[7]) || 0 // Column H (index 7)
            const sqft = parseFloat(row[9]) || 0 // Column J (index 9)
            const lbs = parseFloat(row[10]) || 0 // Column K (index 10)
            const qty = parseFloat(row[12]) || 0 // Column M (index 12)

            currentFoundationSubsectionItems.push({
              particulars: bText,
              takeoff: takeoff,
              unit: unit,
              height: height,
              sqft: sqft,
              lbs: lbs,
              qty: qty,
              rawRow: row,
              rawRowNumber: rowIndex + 2
            })
            return // Skip further processing for this row
          }
        }

        // Check if this row is a sum row (has total SQFT in column J - this is the row with totals for the subsection)
        // Column J = 10th column = index 9 (SQFT)
        // Column M = 13th column = index 12 (QTY)
        if (currentSubsection) {
          // Try multiple ways to parse the values (handle formulas, strings, numbers)
          const sqftRaw = row[9] // Column J (index 9) - SQFT total
          const qtyRaw = row[12] // Column M (index 12) - QTY

          let sqftTotal = 0
          let qtyValue = 0

          // Parse SQFT
          if (sqftRaw !== undefined && sqftRaw !== null && sqftRaw !== '') {
            sqftTotal = parseFloat(sqftRaw) || 0
            if (isNaN(sqftTotal) && typeof sqftRaw === 'string') {
              // Try to extract number from string
              const numMatch = sqftRaw.match(/[\d.]+/)
              if (numMatch) sqftTotal = parseFloat(numMatch[0]) || 0
            }
          }

          // Parse QTY
          if (qtyRaw !== undefined && qtyRaw !== null && qtyRaw !== '') {
            qtyValue = parseFloat(qtyRaw) || 0
            if (isNaN(qtyValue) && typeof qtyRaw === 'string') {
              // Try to extract number from string
              const numMatch = qtyRaw.match(/[\d.]+/)
              if (numMatch) qtyValue = parseFloat(numMatch[0]) || 0
            }
          }

          if (sqftTotal > 0 && (!colB || colB.trim() === '')) {
            // This is the sum row for the current subsection (has SQFT total, empty or no B column)
            // QTY will be in this same row in column M
            sumRowsBySubsection.set(currentSubsection, row)
            return
          } else {
            // Still store it if it has QTY
            if (qtyValue > 0) {
              sumRowsBySubsection.set(currentSubsection, row)
            }
          }
        }

        // Find the first row with empty Particulars (column B) after excavation subsection appears
        // Calculate SQFT in real time and sum until we find the empty row
        if (inExcavationSubsection && !foundEmptyParticulars) {
          const bText = colB ? String(colB).trim() : ''
          const isParticularsEmpty = !colB || bText === ''

          // Calculate SQFT and CY for this row in real time
          const calculatedSQFT = calculateRowSQFT(row)
          const calculatedCY = calculateRowCY(row)

          // Check if column B (Particulars) is empty
          if (isParticularsEmpty) {
            // Found first empty Particulars row - get CY from column L of this row
            let cyRaw = row[11] // Column L (index 11) - CY

            // Try to get evaluated value from workbook if it's a formula
            if (workbook) {
              try {
                const evaluatedCY = workbook.getValueRowCol(calcSheetIndex, rowIndex, 11)
                if (evaluatedCY !== undefined && evaluatedCY !== null && evaluatedCY !== '') {
                  cyRaw = evaluatedCY
                }
              } catch (e) {
                // Ignore errors
              }
            }

            // Parse CY value
            let cyValue = 0
            if (cyRaw !== undefined && cyRaw !== null && cyRaw !== '') {
              cyValue = parseFloat(cyRaw) || 0
              if (isNaN(cyValue) && typeof cyRaw === 'string') {
                const numMatch = cyRaw.match(/[\d.]+/)
                if (numMatch) cyValue = parseFloat(numMatch[0]) || 0
              }
            }

            // Use the running sum as the total for SQFT, and CY from column L
            if (excavationRunningSum > 0) {
              excavationTotalSQFT = excavationRunningSum
              excavationTotalCY = cyValue > 0 ? cyValue : excavationRunningCYSum
              excavationEmptyRowIndex = rowIndex + 1 // Store Excel row number (1-based)
              foundEmptyParticulars = true
            }
          } else {
            // Particulars is filled - calculate and add to running sum
            if (calculatedSQFT > 0) {
              excavationRunningSum += calculatedSQFT
            }
            if (calculatedCY > 0) {
              excavationRunningCYSum += calculatedCY
            }
          }
        }

        // Find the first row with empty Particulars (column B) after backfill subsection appears
        // Calculate SQFT and CY in real time and sum until we find the empty row
        if (inBackfillSubsection && !foundBackfillEmptyParticulars) {
          const bText = colB ? String(colB).trim() : ''
          const isParticularsEmpty = !colB || bText === ''

          // Calculate SQFT and CY for this row in real time
          const calculatedSQFT = calculateRowSQFT(row)
          const calculatedCY = calculateRowCY(row)

          // Check if column B (Particulars) is empty
          if (isParticularsEmpty) {
            // Found first empty Particulars row - get CY from column L of this row
            let cyRaw = row[11] // Column L (index 11) - CY

            // Try to get evaluated value from workbook if it's a formula
            if (workbook) {
              try {
                const evaluatedCY = workbook.getValueRowCol(calcSheetIndex, rowIndex, 11)
                if (evaluatedCY !== undefined && evaluatedCY !== null && evaluatedCY !== '') {
                  cyRaw = evaluatedCY
                }
              } catch (e) {
                // Ignore errors
              }
            }

            // Parse CY value
            let cyValue = 0
            if (cyRaw !== undefined && cyRaw !== null && cyRaw !== '') {
              cyValue = parseFloat(cyRaw) || 0
              if (isNaN(cyValue) && typeof cyRaw === 'string') {
                const numMatch = cyRaw.match(/[\d.]+/)
                if (numMatch) cyValue = parseFloat(numMatch[0]) || 0
              }
            }

            // Use the running sum as the total for SQFT, and CY from column L
            if (backfillRunningSum > 0) {
              backfillTotalSQFT = backfillRunningSum
              backfillTotalCY = cyValue > 0 ? cyValue : backfillRunningCYSum
              backfillEmptyRowIndex = rowIndex + 1 // Store Excel row number (1-based)
              foundBackfillEmptyParticulars = true
            }
          } else {
            // Particulars is filled - calculate and add to running sum
            if (calculatedSQFT > 0) {
              backfillRunningSum += calculatedSQFT
            }
            if (calculatedCY > 0) {
              backfillRunningCYSum += calculatedCY
            }
          }
        }

        // Find the first row with empty Particulars (column B) after rock excavation subsection appears
        // Calculate SQFT and CY in real time and sum until we find the empty row
        if (inRockExcavationSubsection && !foundRockExcavationEmptyParticulars) {
          const bText = colB ? String(colB).trim() : ''
          const isParticularsEmpty = !colB || bText === ''

          // Calculate SQFT and CY for this row in real time
          const calculatedSQFT = calculateRowSQFT(row)
          const calculatedCY = calculateRowCY(row)

          // Check if column B (Particulars) is empty
          if (isParticularsEmpty) {
            // Found first empty Particulars row - get CY from column L of this row
            let cyRaw = row[11] // Column L (index 11) - CY

            // Try to get evaluated value from workbook if it's a formula
            if (workbook) {
              try {
                const evaluatedCY = workbook.getValueRowCol(calcSheetIndex, rowIndex, 11)
                if (evaluatedCY !== undefined && evaluatedCY !== null && evaluatedCY !== '') {
                  cyRaw = evaluatedCY
                }
              } catch (e) {
                // Ignore errors
              }
            }

            // Parse CY value
            let cyValue = 0
            if (cyRaw !== undefined && cyRaw !== null && cyRaw !== '') {
              cyValue = parseFloat(cyRaw) || 0
              if (isNaN(cyValue) && typeof cyRaw === 'string') {
                const numMatch = cyRaw.match(/[\d.]+/)
                if (numMatch) cyValue = parseFloat(numMatch[0]) || 0
              }
            }

            // Use the running sum as the total for SQFT, and CY from column L
            if (rockExcavationRunningSum > 0) {
              rockExcavationTotalSQFT = rockExcavationRunningSum
              rockExcavationTotalCY = cyValue > 0 ? cyValue : rockExcavationRunningCYSum
              rockExcavationEmptyRowIndex = rowIndex + 1 // Store Excel row number (1-based)
              foundRockExcavationEmptyParticulars = true
            }
          } else {
            // Particulars is filled - calculate and add to running sum
            if (calculatedSQFT > 0) {
              rockExcavationRunningSum += calculatedSQFT
            }
            if (calculatedCY > 0) {
              rockExcavationRunningCYSum += calculatedCY
            }
          }
        }

        // Data row for current subsection – collect first item for parsing and track row for SF calculation
        if (!currentSubsection || !colB) return

        // Only process demolition subsections for linesBySubsection
        if (inDemolitionSection) {
          if (!linesBySubsection.has(currentSubsection)) {
            linesBySubsection.set(currentSubsection, bText)
          }
        }

        // Add row to subsection for SF calculation (both demolition and excavation)
        if (!rowsBySubsection.has(currentSubsection)) {
          rowsBySubsection.set(currentSubsection, [])
        }
        rowsBySubsection.get(currentSubsection).push(row)
        dataRowCount++
      })


      // Helper function to extract DM reference from raw data for a given subsection
      const getDMReferenceFromRawData = (subsectionName) => {
        if (!rawData || !Array.isArray(rawData) || rawData.length < 2) {
          return 'DM-106.00' // Default fallback
        }

        const headers = rawData[0]
        const dataRows = rawData.slice(1)

        // Find the "digitizer item" column index
        const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
        if (digitizerIdx === -1) {
          return 'DM-106.00' // Default fallback
        }

        // Map subsection names to search patterns
        const subsectionPatterns = {
          'Demo slab on grade': /demo\s+sog/i,
          'Demo strip footing': /demo\s+sf/i,
          'Demo foundation wall': /demo\s+fw/i,
          'Demo isolated footing': /demo\s+isolated\s+footing/i
        }

        const pattern = subsectionPatterns[subsectionName]
        if (!pattern) {
          return 'DM-106.00' // Default fallback
        }

        // Search through raw data rows to find matching item
        for (const row of dataRows) {
          const digitizerItem = row[digitizerIdx]
          if (digitizerItem && pattern.test(String(digitizerItem))) {
            // Extract DM reference from the digitizer item text
            const text = String(digitizerItem).trim()
            const dmMatch = text.match(/DM-[\d.]+/i)
            if (dmMatch) {
              return dmMatch[0]
            }
          }
        }

        return 'DM-106.00' // Default fallback if not found
      }

      // Helper function to extract thickness/dimensions and build template
      const buildDemolitionTemplate = (subsectionName, itemText) => {
        // Extract slab type from subsection name (everything after "Demo")
        // e.g., "Demo slab on grade" → "slab on grade"
        const slabTypeMatch = subsectionName.match(/^Demo\s+(.+)$/i)
        const slabType = slabTypeMatch ? slabTypeMatch[1].trim() : subsectionName.replace(/^Demo\s+/i, '').trim()

        // Extract DM reference from raw data (not from itemText)
        const dmReference = getDMReferenceFromRawData(subsectionName)

        if (!itemText) {
          // If no item text, return template with default thickness for slab on grade
          if (slabType.toLowerCase().includes('slab on grade')) {
            return `Allow to saw-cut/demo/remove/dispose existing (4" thick) ${slabType} @ existing building as per ${dmReference}`
          }
          return `Allow to saw-cut/demo/remove/dispose existing ${slabType} @ existing building as per ${dmReference}`
        }

        const text = String(itemText).trim()
        let thicknessPart = ''

        // Check if it's slab on grade (has thickness) or other types (has dimensions)
        if (slabType.toLowerCase().includes('slab on grade')) {
          // Extract thickness (e.g., "4" thick" or "4\" thick")
          const thicknessMatch = text.match(/(\d+["\"]?\s*thick)/i) || text.match(/(\d+["\"]?)/)
          thicknessPart = thicknessMatch ? `(${thicknessMatch[1]})` : '(4" thick)'
        } else {
          // Extract dimensions (e.g., "(2'-0"x1'-0")" or "(2'-0"x3'-0"x1'-6")")
          const dimMatch = text.match(/\(([^)]+)\)/)
          thicknessPart = dimMatch ? `(${dimMatch[1]})` : ''
        }

        // Build template: "Allow to saw-cut/demo/remove/dispose existing {thickness} {slab_type} @ existing building as per {dmReference}"
        return `Allow to saw-cut/demo/remove/dispose existing ${thicknessPart} ${slabType} @ existing building as per ${dmReference}`
      }

      // Helper function to calculate SF for a single row
      const calculateRowSF = (row) => {
        // First try to use column J (SQFT) if it exists and is a number
        const sqftValue = parseFloat(row[9]) || 0 // Column J (index 9)

        if (sqftValue > 0) {
          return sqftValue
        }

        // If column J is empty, calculate from takeoff, length, width
        const takeoff = parseFloat(row[2]) || 0 // Column C (index 2)
        const length = parseFloat(row[5]) || 0   // Column F (index 5)
        const width = parseFloat(row[6]) || 0    // Column G (index 6)

        if (takeoff > 0) {
          if (width > 0 && length > 0) {
            // Has both width and length: SQFT = takeoff * length * width
            return takeoff * length * width
          } else if (width > 0) {
            // Has width but no length: SQFT = takeoff * width
            return takeoff * width
          } else if (length > 0) {
            // Has length but no width: SQFT = takeoff * length
            return takeoff * length
          } else {
            // No width and no length: SQFT = takeoff
            return takeoff
          }
        }
        return 0
      }

      // Helper function to calculate SF for a set of rows
      const calculateSF = (rows) => {
        let totalSF = 0
        if (!rows || rows.length === 0) return 0

        rows.forEach(row => {
          totalSF += calculateRowSF(row)
        })
        return totalSF
      }

      // Helper function to calculate CY for a set of rows
      // CY = sum of (SF * height / 27) for each row
      const calculateCY = (rows) => {
        let totalCY = 0
        if (!rows || rows.length === 0) return 0

        rows.forEach(row => {
          const sf = calculateRowSF(row)
          const height = parseFloat(row[7]) || 0 // Column H (index 7)

          if (sf > 0 && height > 0) {
            const cy = (sf * height) / 27
            totalCY += cy
          }
        })
        return totalCY
      }

      // Helper function to get QTY from sum row for a subsection
      // Returns QTY (column M, index 12) from the sum row if it exists, otherwise returns null
      const getQTYFromSumRow = (subsectionName) => {
        const sumRow = sumRowsBySubsection.get(subsectionName)
        if (!sumRow) {
          return null
        }

        const qty = parseFloat(sumRow[12]) || 0 // Column M (index 12)
        return qty > 0 ? qty : null
      }



      // Workbook is already initialized earlier, just try to refresh it if needed
      if (!workbook) {
        try {
          workbook = spreadsheet.getWorkbook()
        } catch (e) {
          // Ignore errors
        }
      }

      // Also check rows to find demolition subsection sum rows (for cases where they weren't caught in the first loop)
      for (let excelRow = 0; excelRow < calculationData.length; excelRow++) {
        const rowIndex = excelRow // 0-based index
        const row = calculationData[rowIndex]

        if (!row) continue

        // If this looks like a sum row (has SQFT, empty B), update our map
        // Try to get values from spreadsheet if calculationData is empty
        let sqftValue = row[9] || null
        let bValue = row[1] || null
        let qtyValue = row[12] || null

        // If values are empty in calculationData, try reading from spreadsheet
        if ((!sqftValue || sqftValue === '') && workbook) {
          try {
            sqftValue = workbook.getValueRowCol(calcSheetIndex, rowIndex, 9)
          } catch (e) { }
        }
        if ((!bValue || bValue === '') && workbook) {
          try {
            bValue = workbook.getValueRowCol(calcSheetIndex, rowIndex, 1)
          } catch (e) { }
        }
        if ((!qtyValue || qtyValue === '') && workbook) {
          try {
            qtyValue = workbook.getValueRowCol(calcSheetIndex, rowIndex, 12)
          } catch (e) { }
        }

        const sqftNum = parseFloat(sqftValue) || 0
        const bText = String(bValue || '').trim()
        const qtyNum = parseFloat(qtyValue) || 0

        if (sqftNum > 0 && (!bText || bText === '')) {
          // Find which subsection this belongs to by checking nearby rows
          let foundSubsection = null
          for (let checkRow = excelRow; checkRow >= Math.max(0, excelRow - 10); checkRow--) {
            if (calculationData[checkRow]) {
              const checkBValue = calculationData[checkRow][1] || ''
              const checkBText = String(checkBValue).trim()
              if (checkBText.endsWith(':')) {
                const subsectionName = checkBText.slice(0, -1).trim()
                if (subsectionName.startsWith('Demo')) {
                  foundSubsection = subsectionName
                  break
                }
              }
            }
          }

          if (foundSubsection) {
            // Create a row array with the actual values (from spreadsheet if available)
            const actualRow = Array(13).fill('')
            actualRow[9] = sqftNum
            actualRow[12] = qtyNum
            sumRowsBySubsection.set(foundSubsection, actualRow)
          }
        }
      }

      // Render specific demolition lines into fixed rows B14–B17 (moved up one row)
      const orderedSubsections = [
        'Demo slab on grade',
        'Demo strip footing',
        'Demo foundation wall',
        'Demo isolated footing'
      ]

      orderedSubsections.forEach((name, index) => {
        const rowIndex = 14 + index // 14,15,16,17 (moved up one row to remove empty row after Demolition scope)
        const originalText = linesBySubsection.get(name)
        const templateText = buildDemolitionTemplate(name, originalText)
        const cellRef = `${pfx}B${rowIndex}`

        // Always set the template text, even if no original data found
        spreadsheet.updateCell({ value: templateText }, cellRef)

        // Enable text wrapping using Syncfusion's wrap method
        try {
          spreadsheet.wrap(cellRef, true)
        } catch (e) {
          // Fallback if wrap method doesn't exist
        }

        // Calculate and set row height based on content
        const dynamicHeight = calculateRowHeight(templateText)
        spreadsheet.setRowHeight(dynamicHeight, rowIndex - 1, proposalSheetIndex) // rowIndex is 1-based, setRowHeight uses 0-based

        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            verticalAlign: 'top'
          },
          cellRef
        )

        // Calculate and display individual SF for this subsection in column D
        const subsectionRows = rowsBySubsection.get(name) || []
        const subsectionSF = calculateSF(subsectionRows)
        const formattedSF = parseFloat(subsectionSF.toFixed(2))

        spreadsheet.updateCell({ value: formattedSF }, `${pfx}D${rowIndex}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'right',
            format: '#,##0.00'
          },
          `${pfx}D${rowIndex}`
        )

        // Calculate and display individual CY for this subsection in column F
        // CY = sum of (SF * height / 27) for all rows in subsection
        const subsectionCY = calculateCY(subsectionRows)
        const formattedCY = parseFloat(subsectionCY.toFixed(2))

        spreadsheet.updateCell({ value: formattedCY }, `${pfx}F${rowIndex}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'right',
            format: '#,##0.00'
          },
          `${pfx}F${rowIndex}`
        )

        // Get QTY from sum row for this subsection in column G (only if sum row has QTY)
        const subsectionQTY = getQTYFromSumRow(name)
        if (subsectionQTY !== null) {
          const formattedQTY = parseFloat(subsectionQTY.toFixed(2))
          spreadsheet.updateCell({ value: formattedQTY }, `${pfx}G${rowIndex}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'right',
              format: '#,##0.00'
            },
            `${pfx}G${rowIndex}`
          )
        }

        // Fill SF (column J) with 7 for rows 14-17
        spreadsheet.updateCell({ value: 7 }, `${pfx}J${rowIndex}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'right',
            format: '$#,##0.00'
          },
          `${pfx}J${rowIndex}`
        )

        // Fill CY (column L) with 500 for rows 14-17
        spreadsheet.updateCell({ value: 500 }, `${pfx}L${rowIndex}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'right',
            format: '$#,##0.00'
          },
          `${pfx}L${rowIndex}`
        )

        // Add $/1000 formula in column H: =ROUNDUP(MAX(C*I,D*J,E*K,F*L,G*M,N)/1000,1)
        const dollarFormula = `=ROUNDUP(MAX(C${rowIndex}*I${rowIndex},D${rowIndex}*J${rowIndex},E${rowIndex}*K${rowIndex},F${rowIndex}*L${rowIndex},G${rowIndex}*M${rowIndex},N${rowIndex})/1000,1)`
        spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${rowIndex}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'right',
            format: '$#,##0.00'
          },
          `${pfx}H${rowIndex}`
        )
      })

      // Apply background color #E2EFDA to columns I-N (LF, SF, LBS, CY, QTY, LS) for all demolition rows
      // Extend to all rows below demolition scope (rows 14-19)
      for (let row = 14; row <= 19; row++) {
        const columns = ['I', 'J', 'K', 'L', 'M', 'N'] // LF, SF, LBS, CY, QTY, LS
        columns.forEach(col => {
          spreadsheet.cellFormat(
            { backgroundColor: '#E2EFDA' },
            `${pfx}${col}${row}`
          )
        })
      }

      // Add note below all demolition items (row 18)
      spreadsheet.updateCell({ value: 'Note: Site/building demolition by others.' }, `${pfx}B18`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: 'white'
        },
        `${pfx}B18`
      )

      // Add Demolition Total row (row 19)
      spreadsheet.merge(`${pfx}D19:E19`)
      spreadsheet.updateCell({ value: 'Demolition Total:' }, `${pfx}D19`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: '#BDD7EE',
          border: '1px solid #000000'
        },
        `${pfx}D19:E19`
      )

      spreadsheet.merge(`${pfx}F19:G19`)
      const totalFormula = `=SUM(H14:H17)*1000`
      spreadsheet.updateCell({ formula: totalFormula }, `${pfx}F19`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          backgroundColor: '#BDD7EE',
          border: '1px solid #000000',
          format: '$#,##0.00'
        },
        `${pfx}F19:G19`
      )

      // Apply background color to entire row B19:G19
      spreadsheet.cellFormat(
        {
          backgroundColor: '#BDD7EE'
        },
        `${pfx}B19:G19`
      )

      // Row 20: Empty row - ensure it's white (not blue)
      spreadsheet.cellFormat({ backgroundColor: 'white' }, `${pfx}B20:G20`)

      // Row 21: Excavation scope (moved up one row to remove extra empty line)
      spreadsheet.updateCell({ value: 'Excavation scope:' }, `${pfx}B21`)
      spreadsheet.cellFormat({
        backgroundColor: '#BDD7EE',
        textAlign: 'center',
        verticalAlign: 'middle',
        textDecoration: 'underline',
        fontWeight: 'normal'
      }, `${pfx}B21`)
      // Ensure surrounding cells are white (not blue)
      spreadsheet.cellFormat({ backgroundColor: 'white' }, `${pfx}C21:G21`)
      spreadsheet.setRowHeight(20, 22, proposalSheetIndex) // Set row height to 22 (row 21 is index 20) - set after formatting

      // Row 22: Empty row - ensure it's white (not blue)
      spreadsheet.cellFormat({ backgroundColor: 'white' }, `${pfx}B22:G22`)
      spreadsheet.setRowHeight(21, 22, proposalSheetIndex) // Set row height to 22 (row 22 is index 21)

      // Row 23: Soil excavation scope
      spreadsheet.updateCell({ value: 'Soil excavation scope:' }, `${pfx}B23`)
      spreadsheet.cellFormat({
        backgroundColor: '#FFF2CC',
        textAlign: 'center',
        verticalAlign: 'middle',
        textDecoration: 'underline',
        fontWeight: 'bold',
        border: '1px solid #000000'
      }, `${pfx}B23`)
      spreadsheet.setRowHeight(22, 22, proposalSheetIndex) // Set fixed small height for row 23 (index 22)

      // Row 24: First soil excavation line
      const soilExcavationText = 'Allow to perform soil excavation, trucking & disposal (Havg=16\'-9") as per SOE-101.00, P-301.01 & details on SOE-201.01 to SOE-204.00'
      spreadsheet.updateCell({ value: soilExcavationText }, `${pfx}B24`)
      spreadsheet.wrap(`${pfx}B24`, true)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: 'white',
          verticalAlign: 'top'
        },
        `${pfx}B24`
      )
      // Calculate and set row height based on content
      const soilExcavationHeight = calculateRowHeight(soilExcavationText)
      spreadsheet.setRowHeight(soilExcavationHeight, 23, proposalSheetIndex) // Row 24 is index 23

      // Add SF value from excavation total to column D (using formula to reference calculation sheet)
      if (excavationEmptyRowIndex) {
        // Use formula to reference the calculation sheet
        spreadsheet.updateCell({ formula: `='Calculations Sheet'!J${excavationEmptyRowIndex}` }, `${pfx}D25`)
      } else {
        // Fallback to value if row index not found
        const formattedExcavationSF = parseFloat(excavationTotalSQFT.toFixed(2))
        spreadsheet.updateCell({ value: formattedExcavationSF }, `${pfx}D25`)
      }
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '#,##0.00'
        },
        `${pfx}D24`
      )

      // Add CY value from excavation total to column F (using formula to reference calculation sheet)
      if (excavationEmptyRowIndex) {
        // Use formula to reference the calculation sheet
        spreadsheet.updateCell({ formula: `='Calculations Sheet'!L${excavationEmptyRowIndex}` }, `${pfx}F25`)
      } else {
        // Fallback to value if row index not found
        const formattedExcavationCY = parseFloat(excavationTotalCY.toFixed(2))
        spreadsheet.updateCell({ value: formattedExcavationCY }, `${pfx}F25`)
      }
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '#,##0.00'
        },
        `${pfx}F24`
      )

      // Fill CY (column L) with 75 for row 24
      spreadsheet.updateCell({ value: 75 }, `${pfx}L24`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '$#,##0.00'
        },
        `${pfx}L24`
      )

      // Add $/1000 formula in column H for row 24
      const dollarFormula24 = `=ROUNDUP(MAX(C24*I24,D24*J24,E24*K24,F24*L24,G24*M24,N24)/1000,1)`

      // Read formulas and evaluate them in real-time with longer timeout
      setTimeout(() => {
        try {
          // Try multiple ways to access the workbook
          let wb = null
          try {
            wb = spreadsheet.getWorkbook()
          } catch (e) {
            try {
              wb = spreadsheet.workbook
            } catch (e2) {
              wb = workbook // Fallback to workbook variable
            }
          }

          // If still null, try accessing through spreadsheet's sheets
          if (!wb && spreadsheet.sheets) {
            wb = { sheets: spreadsheet.sheets }
          }

          if (wb && wb.sheets) {
            const propSheetIndex = proposalSheetIndex
            const row24 = 23 // 0-based index for row 24

            // Helper function to get cell formula or value from workbook
            const getCellInfo = (row, col, colLetter) => {
              try {
                // Try to get the cell from the sheet
                const sheet = wb.sheets[propSheetIndex]
                if (sheet && sheet.rows && sheet.rows[row] && sheet.rows[row].cells) {
                  const cell = sheet.rows[row].cells[col]
                  if (cell) {
                    if (cell.formula) {
                      return { type: 'formula', value: cell.formula, displayValue: (cell.value !== undefined && cell.value !== null) ? cell.value : 0 }
                    }
                    return { type: 'value', value: (cell.value !== undefined && cell.value !== null) ? cell.value : 0 }
                  }
                }
                // Fallback: try to get value directly if method exists
                if (typeof wb.getValueRowCol === 'function') {
                  const value = wb.getValueRowCol(propSheetIndex, row, col)
                  return { type: 'value', value: value || 0 }
                }
                // If no value found, return 0
                return { type: 'value', value: 0 }
              } catch (e) {
                return { type: 'value', value: 0 }
              }
            }

            // Helper function to evaluate a cell reference (e.g., 'Calculations Sheet'!J75 or J75)
            const evaluateCellRef = (ref) => {
              try {
                let sheetIdx = propSheetIndex
                let cellRef = ref.trim()

                // Remove quotes from sheet name if present
                if (ref.includes('!')) {
                  const parts = ref.split('!')
                  let sheetName = parts[0].trim()
                  // Remove single quotes if present
                  sheetName = sheetName.replace(/^'|'$/g, '')
                  cellRef = parts[1].trim()

                  const sheet = wb.sheets.find(s => s.name === sheetName)
                  if (sheet) {
                    sheetIdx = wb.sheets.indexOf(sheet)
                  }
                }

                const col = cellRef.charCodeAt(0) - 65 // A=0, B=1, etc.
                const row = parseInt(cellRef.substring(1)) - 1 // Convert to 0-based

                // Try to read value from sheet structure
                let value = 0
                try {
                  const sheet = wb.sheets[sheetIdx]
                  if (sheet && sheet.rows && sheet.rows[row] && sheet.rows[row].cells) {
                    const cell = sheet.rows[row].cells[col]
                    if (cell) {
                      // If cell has a formula, try to get the evaluated value
                      if (cell.formula) {
                        value = (cell.value !== undefined && cell.value !== null) ? cell.value : 0
                      } else {
                        value = (cell.value !== undefined && cell.value !== null) ? cell.value : 0
                      }
                    }
                  }

                  // Fallback: try getValueRowCol if available
                  if (value === 0 && typeof wb.getValueRowCol === 'function') {
                    value = wb.getValueRowCol(sheetIdx, row, col) || 0
                  }
                } catch (e2) {
                  // Ignore errors
                }

                return value || 0
              } catch (e) {
                return 0
              }
            }

            // Get cell info for each cell
            const cells = {
              c24: getCellInfo(row24, 2, 'C'),
              i24: getCellInfo(row24, 8, 'I'),
              d24: getCellInfo(row24, 3, 'D'),
              j24: getCellInfo(row24, 9, 'J'),
              e24: getCellInfo(row24, 4, 'E'),
              k24: getCellInfo(row24, 10, 'K'),
              f24: getCellInfo(row24, 5, 'F'),
              l24: getCellInfo(row24, 11, 'L'),
              g24: getCellInfo(row24, 6, 'G'),
              m24: getCellInfo(row24, 12, 'M'),
              n24: getCellInfo(row24, 13, 'N')
            }

            // Evaluate each cell
            const values = {}
            Object.keys(cells).forEach(key => {
              const cell = cells[key]
              if (cell.type === 'formula') {
                // Extract cell references from formula and evaluate
                const formula = cell.value.replace(/^=/, '')

                // Match formulas like 'Calculations Sheet'!J92 or J92
                // The full match (refMatch[0]) includes the sheet reference if present
                const refMatch = formula.match(/(?:'[^']+'!)?[A-Z]+\d+/)
                if (refMatch) {
                  const fullRef = refMatch[0]
                  values[key] = evaluateCellRef(fullRef)
                } else {
                  values[key] = cell.displayValue || 0
                }
              } else {
                values[key] = cell.value
              }
            })
          }
        } catch (e) {
          console.error('Row 24: Error reading values:', e)
          console.error('Row 24: Error stack:', e.stack)
        }
      }, 2000) // Increased timeout to 2 seconds

      spreadsheet.updateCell({ formula: dollarFormula24 }, `${pfx}H24`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '$#,##0.00'
        },
        `${pfx}H24`
      )

      // Row 27: Second soil excavation line
      spreadsheet.updateCell({ value: 'Allow to import new clean soil to backfill and compact as per SOE-101.00, P-301.01 & details on SOE-201.01 to SOE-204.00' }, `${pfx}B27`)
      spreadsheet.wrap(`${pfx}B27`, true)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: 'white',
          verticalAlign: 'top'
        },
        `${pfx}B27`
      )

      // Add SF value from backfill total to column D (using formula to reference calculation sheet)
      if (backfillEmptyRowIndex) {
        // Use formula to reference the calculation sheet
        spreadsheet.updateCell({ formula: `='Calculations Sheet'!J${backfillEmptyRowIndex}` }, `${pfx}D27`)
      } else {
        // Fallback to value if row index not found
        const formattedBackfillSF = parseFloat(backfillTotalSQFT.toFixed(2))
        spreadsheet.updateCell({ value: formattedBackfillSF }, `${pfx}D27`)
      }
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '#,##0.00'
        },
        `${pfx}D27`
      )

      // Add CY value from backfill total to column F (using formula to reference calculation sheet)
      if (backfillEmptyRowIndex) {
        // Use formula to reference the calculation sheet
        spreadsheet.updateCell({ formula: `='Calculations Sheet'!L${backfillEmptyRowIndex}` }, `${pfx}F27`)
      } else {
        // Fallback to value if row index not found
        const formattedBackfillCY = parseFloat(backfillTotalCY.toFixed(2))
        spreadsheet.updateCell({ value: formattedBackfillCY }, `${pfx}F27`)
      }
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '#,##0.00'
        },
        `${pfx}F27`
      )

      // Fill CY (column L) with 85 for row 27
      spreadsheet.updateCell({ value: 85 }, `${pfx}L27`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '$#,##0.00'
        },
        `${pfx}L27`
      )

      // Add $/1000 formula in column H for row 27
      const dollarFormula27 = `=ROUNDUP(MAX(C27*I27,D27*J27,E27*K27,F27*L27,G27*M27,N27)/1000,1)`

      // Read formulas and evaluate them in real-time with longer timeout
      setTimeout(() => {
        try {
          // Try multiple ways to access the workbook
          let wb = null
          try {
            wb = spreadsheet.getWorkbook()
          } catch (e) {
            try {
              wb = spreadsheet.workbook
            } catch (e2) {
              wb = workbook // Fallback to workbook variable
            }
          }

          // If still null, try accessing through spreadsheet's sheets
          if (!wb && spreadsheet.sheets) {
            wb = { sheets: spreadsheet.sheets }
          }

          if (wb && wb.sheets) {
            const propSheetIndex = proposalSheetIndex
            const row27 = 26 // 0-based index for row 27

            // Helper function to get cell formula or value from workbook
            const getCellInfo = (row, col, colLetter) => {
              try {
                // Try to get the cell from the sheet
                const sheet = wb.sheets[propSheetIndex]
                if (sheet && sheet.rows && sheet.rows[row] && sheet.rows[row].cells) {
                  const cell = sheet.rows[row].cells[col]
                  if (cell) {
                    if (cell.formula) {
                      return { type: 'formula', value: cell.formula, displayValue: (cell.value !== undefined && cell.value !== null) ? cell.value : 0 }
                    }
                    return { type: 'value', value: (cell.value !== undefined && cell.value !== null) ? cell.value : 0 }
                  }
                }
                // Fallback: try to get value directly if method exists
                if (typeof wb.getValueRowCol === 'function') {
                  const value = wb.getValueRowCol(propSheetIndex, row, col)
                  return { type: 'value', value: value || 0 }
                }
                // If no value found, return 0
                return { type: 'value', value: 0 }
              } catch (e) {
                return { type: 'value', value: 0 }
              }
            }

            // Helper function to evaluate a cell reference (e.g., 'Calculations Sheet'!J75 or J75)
            const evaluateCellRef = (ref) => {
              try {
                let sheetIdx = propSheetIndex
                let cellRef = ref.trim()

                // Remove quotes from sheet name if present
                if (ref.includes('!')) {
                  const parts = ref.split('!')
                  let sheetName = parts[0].trim()
                  // Remove single quotes if present
                  sheetName = sheetName.replace(/^'|'$/g, '')
                  cellRef = parts[1].trim()

                  const sheet = wb.sheets.find(s => s.name === sheetName)
                  if (sheet) {
                    sheetIdx = wb.sheets.indexOf(sheet)
                  }
                }

                const col = cellRef.charCodeAt(0) - 65 // A=0, B=1, etc.
                const row = parseInt(cellRef.substring(1)) - 1 // Convert to 0-based

                // Try to read value from sheet structure
                let value = 0
                try {
                  const sheet = wb.sheets[sheetIdx]
                  if (sheet && sheet.rows && sheet.rows[row] && sheet.rows[row].cells) {
                    const cell = sheet.rows[row].cells[col]
                    if (cell) {
                      // If cell has a formula, try to get the evaluated value
                      if (cell.formula) {
                        value = cell.value || 0
                      } else {
                        value = cell.value || 0
                      }
                    }
                  }

                  // Fallback: try getValueRowCol if available
                  if (value === 0 && typeof wb.getValueRowCol === 'function') {
                    value = wb.getValueRowCol(sheetIdx, row, col) || 0
                  }
                } catch (e2) {
                  // Ignore errors
                }

                return value || 0
              } catch (e) {
                return 0
              }
            }
          }
        } catch (e) {
          // Ignore errors
        }
      }, 2000) // Increased timeout to 2 seconds

      spreadsheet.updateCell({ formula: dollarFormula27 }, `${pfx}H27`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '$#,##0.00'
        },
        `${pfx}H27`
      )

      // Row 28: Note 1
      spreadsheet.updateCell({ value: 'Note: Backfill SOE voids by others' }, `${pfx}B28`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: 'white'
        },
        `${pfx}B28`
      )

      // Row 29: Note 2
      spreadsheet.updateCell({ value: 'Note: NJ Res Soil included, contaminated, mixed, hazardous, petroleum impacted not incl.' }, `${pfx}B29`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: 'white'
        },
        `${pfx}B29`
      )

      // Row 30: Note 3
      spreadsheet.updateCell({ value: 'Note: Bedrock not included, see add alt unit rate if required' }, `${pfx}B30`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: 'white'
        },
        `${pfx}B30`
      )

      // Row 31: Soil Excavation Total
      spreadsheet.merge(`${pfx}D31:E31`)
      spreadsheet.updateCell({ value: 'Soil Excavation Total:' }, `${pfx}D31`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: '#FFF2CC',
          border: '1px solid #000000'
        },
        `${pfx}D31:E31`
      )

      spreadsheet.merge(`${pfx}F31:G31`)
      const soilExcavationTotalFormula = `=SUM(H24:H28)*1000`
      spreadsheet.updateCell({ formula: soilExcavationTotalFormula }, `${pfx}F31`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          backgroundColor: '#FFF2CC',
          border: '1px solid #000000'
        },
        `${pfx}F31:G31`
      )
      // Apply number format with comma separator
      try {
        spreadsheet.numberFormat('$#,##0.00', `${pfx}F31:G31`)
      } catch (e) {
        // Fallback to cellFormat if numberFormat doesn't work
        spreadsheet.cellFormat({ format: '$#,##0.00' }, `${pfx}F31`)
      }

      // Apply background color to entire row B31:G31
      spreadsheet.cellFormat(
        {
          backgroundColor: '#FFF2CC'
        },
        `${pfx}B31:G31`
      )

      // Row 32: Empty row (skip)
      // (No content needed - empty row)

      // Apply green background color #E2EFDA to columns I-N for all rows between demolition and rock excavation (rows 22-32)
      for (let row = 22; row <= 32; row++) {
        const columns = ['I', 'J', 'K', 'L', 'M', 'N'] // LF, SF, LBS, CY, QTY, LS
        columns.forEach(col => {
          spreadsheet.cellFormat(
            { backgroundColor: '#E2EFDA' },
            `${pfx}${col}${row}`
          )
        })
      }

      // Row 33: Rock excavation scope
      spreadsheet.updateCell({ value: 'Rock excavation scope:' }, `${pfx}B33`)
      spreadsheet.cellFormat({
        backgroundColor: '#FFF2CC',
        textAlign: 'center',
        verticalAlign: 'middle',
        textDecoration: 'underline',
        fontWeight: 'bold',
        border: '1px solid #000000'
      }, `${pfx}B33`)

      // Apply green background color #E2EFDA to columns I-N for row 33
      const columns33 = ['I', 'J', 'K', 'L', 'M', 'N'] // LF, SF, LBS, CY, QTY, LS
      columns33.forEach(col => {
        spreadsheet.cellFormat(
          { backgroundColor: '#E2EFDA' },
          `${pfx}${col}33`
        )
      })

      // Row 34: First rock excavation line
      spreadsheet.updateCell({ value: 'Allow to perform rock excavation, trucking & disposal for building (Havg=2\'-9") as per SOE-100.00 & details on SOE-A-202.00' }, `${pfx}B34`)
      spreadsheet.wrap(`${pfx}B34`, true)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: 'white',
          verticalAlign: 'top'
        },
        `${pfx}B34`
      )

      // Add SF value from rock excavation total to column D (using processor totals)
      const formattedRockExcavationSF = parseFloat(rockExcavationTotals.totalSQFT.toFixed(2))
      spreadsheet.updateCell({ value: formattedRockExcavationSF }, `${pfx}D34`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '#,##0.00'
        },
        `${pfx}D34`
      )

      // Add CY value from rock excavation total to column F (using processor totals)
      const formattedRockExcavationCY = parseFloat(rockExcavationTotals.totalCY.toFixed(2))
      spreadsheet.updateCell({ value: formattedRockExcavationCY }, `${pfx}F34`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '#,##0.00'
        },
        `${pfx}F34`
      )

      // Row 35: Second rock excavation line
      spreadsheet.updateCell({ value: 'Allow to perform line drilling as per SOE-100.00' }, `${pfx}B35`)
      spreadsheet.wrap(`${pfx}B35`, true)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: 'white',
          verticalAlign: 'top'
        },
        `${pfx}B35`
      )

      // Add CY value from rock excavation total to column F (using processor totals)
      const formattedRockExcavationCY35 = parseFloat(rockExcavationTotals.totalCY.toFixed(2))
      spreadsheet.updateCell({ value: formattedRockExcavationCY35 }, `${pfx}F35`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '#,##0.00'
        },
        `${pfx}F35`
      )

      // Fill SF (column J) with 7 for row 34
      spreadsheet.updateCell({ value: 7 }, `${pfx}J34`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '$#,##0.00'
        },
        `${pfx}J34`
      )

      // Fill CY (column L) with 350 for row 34 (green background) and line drill total FT for row 35
      spreadsheet.updateCell({ value: 350 }, `${pfx}L34`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '$#,##0.00',
          backgroundColor: '#90EE90' // Light green background
        },
        `${pfx}L34`
      )
      // Use line drill total FT for row 35 (column C - first LF column), multiplied by 2
      const formattedLineDrillFT = parseFloat((lineDrillTotalFT * 2).toFixed(2))
      spreadsheet.updateCell({ value: formattedLineDrillFT }, `${pfx}C35`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '#,##0.00'
        },
        `${pfx}C35`
      )

      // Fill second LF column (I) with 30 for row 35
      spreadsheet.updateCell({ value: 30 }, `${pfx}I35`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '#,##0.00'
        },
        `${pfx}I35`
      )

      // Add $/1000 formula in column H for row 34
      const dollarFormula34 = `=ROUNDUP(MAX(C34*I34,D34*J34,E34*K34,F34*L34,G34*M34,N34)/1000,1)`
      spreadsheet.updateCell({ formula: dollarFormula34 }, `${pfx}H34`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '$#,##0.00'
        },
        `${pfx}H34`
      )

      // Apply green background color #E2EFDA to columns I-N (LF, SF, LBS, CY, QTY, LS) for row 34
      const columns34 = ['I', 'J', 'K', 'L', 'M', 'N'] // LF, SF, LBS, CY, QTY, LS
      columns34.forEach(col => {
        spreadsheet.cellFormat(
          { backgroundColor: '#E2EFDA' },
          `${pfx}${col}34`
        )
      })

      // Add $/1000 formula in column H for row 35
      const dollarFormula35 = `=ROUNDUP(MAX(C35*I35,D35*J35,E35*K35,F35*L35,G35*M35,N35)/1000,1)`
      spreadsheet.updateCell({ formula: dollarFormula35 }, `${pfx}H35`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '$#,##0.00'
        },
        `${pfx}H35`
      )

      // Apply green background color #E2EFDA to columns I-N (LF, SF, LBS, CY, QTY, LS) for row 35
      const columns = ['I', 'J', 'K', 'L', 'M', 'N'] // LF, SF, LBS, CY, QTY, LS
      columns.forEach(col => {
        spreadsheet.cellFormat(
          { backgroundColor: '#E2EFDA' },
          `${pfx}${col}35`
        )
      })

      // Row 36: Rock excavation Total
      spreadsheet.merge(`${pfx}D36:E36`)
      spreadsheet.updateCell({ value: 'Rock excavation Total:' }, `${pfx}D36`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: '#FFF2CC',
          border: '1px solid #000000'
        },
        `${pfx}D36:E36`
      )

      spreadsheet.merge(`${pfx}F36:G36`)
      const rockExcavationTotalFormula = `=SUM(H32:H34)*1000`
      spreadsheet.updateCell({ formula: rockExcavationTotalFormula }, `${pfx}F36`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          backgroundColor: '#FFF2CC',
          border: '1px solid #000000'
        },
        `${pfx}F36:G36`
      )
      // Apply number format with comma separator
      try {
        spreadsheet.numberFormat('$#,##0.00', `${pfx}F36:G36`)
      } catch (e) {
        // Fallback to cellFormat if numberFormat doesn't work
        spreadsheet.cellFormat({ format: '$#,##0.00' }, `${pfx}F36`)
      }

      // Apply background color to columns B and C as well
      spreadsheet.cellFormat(
        {
          backgroundColor: '#FFF2CC',
          border: '1px solid #000000'
        },
        `${pfx}B36:C36`
      )

      // Row 37: Empty row

      // Row 38: Excavation Total
      spreadsheet.merge(`${pfx}D38:E38`)
      spreadsheet.updateCell({ value: 'Excavation Total:' }, `${pfx}D38`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: '#BDD7EE',
          border: '1px solid #000000'
        },
        `${pfx}D38:E38`
      )

      spreadsheet.merge(`${pfx}F38:G38`)
      // Sum F31 (Soil Excavation Total) and F36 (Rock excavation Total)
      const excavationTotalFormula = `=SUM(F31,F36)`
      spreadsheet.updateCell({ formula: excavationTotalFormula }, `${pfx}F38`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          backgroundColor: '#BDD7EE',
          border: '1px solid #000000'
        },
        `${pfx}F38:G38`
      )
      // Apply number format with comma separator
      try {
        spreadsheet.numberFormat('$#,##0.00', `${pfx}F38:G38`)
      } catch (e) {
        // Fallback to cellFormat if numberFormat doesn't work
        spreadsheet.cellFormat({ format: '$#,##0.00' }, `${pfx}F38`)
      }

      // Apply background color to columns B and C as well
      spreadsheet.cellFormat(
        {
          backgroundColor: '#BDD7EE',
          border: '1px solid #000000'
        },
        `${pfx}B38:C38`
      )

      // Apply number format with comma separators to column H ($/1000) and all columns after it (I, J, K, L, M, N)
      // Format entire column ranges for data rows (16-38)
      try {
        // Column H ($/1000) - currency format with comma
        spreadsheet.numberFormat('$#,##0.00', `${pfx}H16:H38`)
      } catch (e) {
        // Ignore errors
      }

      // Columns I, J, K, L, M, N - apply comma separator format
      const formatColumns = ['I', 'J', 'K', 'L', 'M', 'N']
      formatColumns.forEach(col => {
        try {
          // Use currency format for columns that show dollar values (I, J, L)
          // and number format with comma for others
          if (col === 'I' || col === 'J' || col === 'L') {
            spreadsheet.numberFormat('$#,##0.00', `${pfx}${col}16:${col}38`)
          } else {
            spreadsheet.numberFormat('#,##0.00', `${pfx}${col}16:${col}38`)
          }
        } catch (e) {
          // Ignore formatting errors
        }
      })

      // Row 39: Empty row

      // Row 40: SOE scope
      spreadsheet.updateCell({ value: 'SOE scope:' }, `${pfx}B40`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'center',
          backgroundColor: '#BDD7EE',
          border: '1px solid #000000'
        },
        `${pfx}B40`
      )

      // Add SOE subsection headings (rows 41-64)
      // We'll add content after each heading, so headings will be at:
      // Row 41: Soldier drilled piles:
      // Row 42: Soldier driven piles: (will be shifted down if drilled piles are added)
      // etc.
      const soeHeadings = [
        'Soldier drilled piles:',
        'Soldier driven piles:',
        'Primary secant piles:',
        'Secondary secant piles:',
        'Tangent pile:',
        'Sheet piles:',
        'Timber lagging:',
        'Timber sheeting:',
        'Bracing:',
        'Tie back anchor:',
        'Tie down anchor:',
        'Parging:',
        'Heel blocks:',
        'Underpinning:',
        'Concrete soil retention pier:',
        'Form board:',
        'Buttons:',
        'Concrete buttons:',
        'Guide wall:',
        'Concrete buttons:',
        'Dowels:',
        'Rock pins:',
        'Rock stabilization:',
        'Shotcrete:',
        'Permission grouting:',
        'Mud slab:',
        'Misc.:'
      ]

      // Store heading row numbers for later reference
      const headingRows = {}
      soeHeadings.forEach((heading, index) => {
        const rowNum = 41 + index
        headingRows[heading] = rowNum
        spreadsheet.updateCell({ value: heading }, `${pfx}B${rowNum}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            backgroundColor: '#D0CECE',
            textDecoration: 'underline',
            border: '1px solid #000000'
          },
          `${pfx}B${rowNum}`
        )
        spreadsheet.setRowHeight(rowNum, 22, proposalSheetIndex)
      })

      // Track row shifts for dynamic content insertion
      // This will help us place content correctly after headings
      let rowShift = 0

      // Helper function to extract SOE page number from raw data
      const getSOEPageFromRawData = (diameter, thickness) => {
        if (!rawData || !Array.isArray(rawData) || rawData.length < 2) {
          return 'SOE-101.00' // Default fallback
        }

        const headers = rawData[0]
        const dataRows = rawData.slice(1)

        // Find the "digitizer item" and "page" column indices
        const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
        const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

        if (digitizerIdx === -1 || pageIdx === -1) {
          return 'SOE-101.00' // Default fallback
        }

        // Create pattern to match diameter and thickness (e.g., "9.625Ø x0.545")
        const pattern = new RegExp(`${diameter}\\s*Ø\\s*x\\s*${thickness}`, 'i')

        // Search through raw data rows to find matching item
        for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
          const row = dataRows[rowIndex]
          const digitizerItem = row[digitizerIdx]
          if (digitizerItem && pattern.test(String(digitizerItem))) {
            // Extract SOE page number from the page column
            const pageValue = row[pageIdx]
            if (pageValue) {
              const pageStr = String(pageValue).trim()
              // Extract SOE reference (e.g., "SOE-101.00")
              const soeMatch = pageStr.match(/SOE-[\d.]+/i)
              if (soeMatch) {
                return soeMatch[0]
              }
            }
          }
        }

        return 'SOE-101.00' // Default fallback if not found
      }

      // Add drilled soldier pile proposal text from collected groups
      const collectedGroups = window.drilledSoldierPileGroups || []

      // Helper function to parse dimension string (e.g., "27'-10"" -> 27.833)
      const parseDimension = (dimStr) => {
        if (!dimStr) return 0
        const match = dimStr.match(/(\d+)(?:'-?)?(\d+)?/)
        if (!match) return 0
        const feet = parseInt(match[1]) || 0
        const inches = parseInt(match[2]) || 0
        return feet + (inches / 12)
      }

      // Calculate totals using formulas from soeProcessor.js
      // FT = H * C (Height * Takeoff) - formula from soeProcessor.js line 255
      // LBS = I * Weight (FT * Weight) - formula from soeProcessor.js line 256
      // Weight for drilled: (Diameter - Thickness) * Thickness * 10.69 - from calculatePileWeight
      collectedGroups.forEach((group, idx) => {
        let groupFT = 0
        let groupLBS = 0
        let groupWeight = null

        group.forEach(item => {
          const particulars = item.particulars || ''
          const takeoff = item.takeoff || 0
          const height = item.height || 0 // Height from column H (already calculated/rounded)

          // Calculate FT for this item: FT = H * C (formula from soeProcessor.js line 255)
          const itemFT = height * takeoff
          groupFT += itemFT

          // Calculate weight for drilled soldier piles
          if (!groupWeight) {
            const drilledMatch = particulars.match(/([0-9.]+)Ø\s*x\s*([0-9.]+)/i)
            if (drilledMatch) {
              const diameter = parseFloat(drilledMatch[1])
              const thickness = parseFloat(drilledMatch[2])
              // Weight formula from calculatePileWeight (soeParser.js line 113-114)
              groupWeight = (diameter - thickness) * thickness * 10.69
            }
          }

          // Calculate LBS for this item: LBS = I * Weight (formula from soeProcessor.js line 256)
          if (groupWeight) {
            const itemLBS = itemFT * groupWeight
            groupLBS += itemLBS
          }
        })

        // Show only the sum totals with formula range
        const firstRowNumber = Math.min(...group.map(item => item.rawRowNumber || 0))
        const lastRowNumber = Math.max(...group.map(item => item.rawRowNumber || 0))

      })

      // Separate HP items from drilled soldier pile items
      const drilledGroups = []
      const hpGroupsFromDrilled = []

      collectedGroups.forEach((group) => {
        const hasHPItems = group.some(item => {
          const particulars = item.particulars || ''
          return /HP\d+x\d+/i.test(particulars)
        })

        if (hasHPItems) {
          // This group contains HP items - move to HP groups
          hpGroupsFromDrilled.push(group)
        } else {
          // This is a drilled soldier pile group
          drilledGroups.push(group)
        }
      })


      // Add HP groups from drilled soldier pile section to HP groups
      if (hpGroupsFromDrilled.length > 0) {
        if (!window.hpSoldierPileGroups) {
          window.hpSoldierPileGroups = []
        }
        window.hpSoldierPileGroups.push(...hpGroupsFromDrilled)
      }

      // Initialize currentRow for both drilled and HP groups
      // Start drilled soldier piles right after "Soldier drilled piles:" heading (row 41)
      let currentRow = 42

      if (drilledGroups.length > 0) {
        // Process each group separately
        drilledGroups.forEach((collectedItems, groupIndex) => {
          if (collectedItems.length > 0) {
            // Parse items to extract diameter, thickness, heights, and embedment
            let diameter = null
            let thickness = null
            let totalHeight = 0
            let heightCount = 0
            let embedment = null
            let totalCount = 0

            // Helper function to parse dimension string (e.g., "27'-10"" -> 27.833)
            const parseDimension = (dimStr) => {
              if (!dimStr) return 0
              const match = dimStr.match(/(\d+)(?:'-?)?(\d+)?/)
              if (!match) return 0
              const feet = parseInt(match[1]) || 0
              const inches = parseInt(match[2]) || 0
              return feet + (inches / 12)
            }

            // Group items by particulars
            const groupedItems = new Map()
            collectedItems.forEach(item => {
              const particulars = item.particulars || ''
              if (!groupedItems.has(particulars)) {
                groupedItems.set(particulars, {
                  particulars: particulars,
                  takeoff: 0,
                  hValue: null,
                  rsValue: null,
                  combinedHeight: null,
                  embedment: null
                })
              }
              const group = groupedItems.get(particulars)
              group.takeoff += (item.takeoff || 0)
            })

            // Process each group to extract values
            groupedItems.forEach((group, particulars) => {
              // Extract diameter and thickness (e.g., "9.625Ø x0.545")
              if (!diameter || !thickness) {
                const drilledMatch = particulars.match(/([0-9.]+)Ø\s*x\s*([0-9.]+)/i)
                if (drilledMatch) {
                  diameter = parseFloat(drilledMatch[1])
                  thickness = parseFloat(drilledMatch[2])
                }
              }

              // Extract H value (e.g., "H=27'-10"" or "H=27'-10"+ RS=15'-0"")
              const hMatch = particulars.match(/H=([0-9'"\-]+)/)
              if (hMatch) {
                let heightValue = parseDimension(hMatch[1])
                group.hValue = hMatch[1]

                // Check if there's a "+" followed by RS (Rock Socket) - if so, add RS to H
                const rsMatch = particulars.match(/\+\s*RS=([0-9'"\-]+)/i)
                if (rsMatch) {
                  const rsValue = parseDimension(rsMatch[1])
                  group.rsValue = rsMatch[1]
                  heightValue = heightValue + rsValue
                }
                group.combinedHeight = heightValue

                // Add height for each takeoff (if takeoff is 8, add height 8 times)
                for (let i = 0; i < group.takeoff; i++) {
                  totalHeight += heightValue
                  heightCount++
                }
              }

              // Extract E value (embedment) (e.g., "E=15'-0"")
              if (!embedment) {
                const eMatch = particulars.match(/E=([0-9'"\-]+)/)
                if (eMatch) {
                  embedment = parseDimension(eMatch[1])
                  group.embedment = eMatch[1]
                }
              }

              // Sum total count
              totalCount += group.takeoff
            })

            if (heightCount > 0) {
              const avgHeight = totalHeight / heightCount
              // Round up to multiple of 5
              const roundToMultipleOf5 = (value) => {
                return Math.ceil(value / 5) * 5
              }
              const avgHeightRoundedTo5 = roundToMultipleOf5(avgHeight)
            }

            if (diameter && thickness && heightCount > 0) {
              // Calculate average height
              const avgHeight = totalHeight / heightCount
              // Round up to multiple of 5 (same as soeProcessor.js)
              const roundToMultipleOf5 = (value) => {
                return Math.ceil(value / 5) * 5
              }
              const avgHeightRounded = roundToMultipleOf5(avgHeight)

              // Format embedment
              let embedmentText = ''
              if (embedment) {
                const embedmentFeet = Math.floor(embedment)
                const embedmentInches = Math.round((embedment - embedmentFeet) * 12)
                if (embedmentInches === 0) {
                  embedmentText = `${embedmentFeet}'-0"`
                } else {
                  embedmentText = `${embedmentFeet}'-${embedmentInches}"`
                }
              } else {
                embedmentText = "0'-0\""
              }

              // Format height
              const heightFeet = Math.floor(avgHeightRounded)
              const heightInches = Math.round((avgHeightRounded - heightFeet) * 12)
              let heightText = ''
              if (heightInches === 0) {
                heightText = `${heightFeet}'-0"`
              } else {
                heightText = `${heightFeet}'-${heightInches}"`
              }

              // Get SOE page number from raw data
              const soePage = getSOEPageFromRawData(diameter, thickness)

              // Find the sum row for this group
              // The sum formula is in the same row as the last item (not the next row)
              const lastRowNumber = Math.max(...collectedItems.map(item => item.rawRowNumber || 0))
              const sumRowIndex = lastRowNumber // Sum row is the same as last item row

              // Create cell references to the sum row in calculation sheet
              // FT is in column I, LBS is in column K, QTY is in column M
              const calcSheetName = 'Calculations Sheet'
              const ftCellRef = `'${calcSheetName}'!I${sumRowIndex}`
              const lbsCellRef = `'${calcSheetName}'!K${sumRowIndex}`
              const qtyCellRef = `'${calcSheetName}'!M${sumRowIndex}`

              // Format the proposal text
              const proposalText = `F&I new (${totalCount})no [${diameter}" Øx${thickness}" thick] drilled soldier piles (H=${heightText}, ${embedmentText} embedment) as per ${soePage}`

              spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  color: '#000000',
                  textAlign: 'left',
                  backgroundColor: 'white',
                  verticalAlign: 'top'
                },
                `${pfx}B${currentRow}`
              )

              // Calculate and set row height based on content
              const dynamicHeight = calculateRowHeight(proposalText)
              spreadsheet.setRowHeight(dynamicHeight, currentRow, proposalSheetIndex)

              // Add FT (LF) to column C - reference to calculation sheet sum row
              spreadsheet.updateCell({ formula: `=${ftCellRef}` }, `${pfx}C${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white'
                },
                `${pfx}C${currentRow}`
              )

              // Add LBS to column E - reference to calculation sheet sum row
              spreadsheet.updateCell({ formula: `=${lbsCellRef}` }, `${pfx}E${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white'
                },
                `${pfx}E${currentRow}`
              )

              // Add QTY to column G - reference to calculation sheet sum row (column M)
              spreadsheet.updateCell({ formula: `=${qtyCellRef}` }, `${pfx}G${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white'
                },
                `${pfx}G${currentRow}`
              )

              // Add 180 to second LF column (column I)
              spreadsheet.updateCell({ value: 180 }, `${pfx}I${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: '#E2EFDA'
                },
                `${pfx}I${currentRow}:N${currentRow}`
              )

              // Add $/1000 formula in column H
              const dollarFormula = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
              spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white',
                  format: '$#,##0.0'
                },
                `${pfx}H${currentRow}`
              )

              currentRow++ // Move to next row for next group
            }
          }
        })
      }

      // Move to row after "Soldier driven piles:" heading for HP piles
      // If drilled piles were added, currentRow is already after the last drilled pile
      // If no drilled piles, start HP piles at row 43 (right after "Soldier driven piles:" at row 42)
      if (drilledGroups.length === 0 || currentRow <= 42) {
        currentRow = 43 // Start HP piles right after "Soldier driven piles:" heading
      } else {
        // Drilled piles were added and we're past row 42
        // Ensure "Soldier driven piles:" heading is visible before HP piles
        // Add it at the current row, then increment for HP piles
        spreadsheet.updateCell({ value: 'Soldier driven piles:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            backgroundColor: '#D0CECE',
            textDecoration: 'underline',
            border: '1px solid #000000'
          },
          `${pfx}B${currentRow}`
        )
        spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
        currentRow++ // Move to next row for HP piles
      }

      // Add HP soldier pile proposal text from collected groups
      const hpCollectedGroups = window.hpSoldierPileGroups || []

      // Calculate totals using formulas from soeProcessor.js
      // FT = H * C (Height * Takeoff) - formula from soeProcessor.js line 255
      // LBS = I * Weight (FT * Weight) - formula from soeProcessor.js line 256
      // Weight for HP: hpWeight from HP12x63 format - from calculatePileWeight
      hpCollectedGroups.forEach((group, idx) => {
        let groupFT = 0
        let groupLBS = 0
        let groupHPWeight = null

        group.forEach(item => {
          const particulars = item.particulars || ''
          const takeoff = item.takeoff || 0
          const height = item.height || 0 // Height from column H (already calculated/rounded)

          // Extract HP weight
          if (!groupHPWeight) {
            const hpMatch = particulars.match(/HP\d+x(\d+)/i)
            if (hpMatch) {
              groupHPWeight = parseFloat(hpMatch[1])
            }
          }

          // Calculate FT for this item: FT = H * C (formula from soeProcessor.js line 255)
          const itemFT = height * takeoff
          groupFT += itemFT

          // Calculate LBS for this item: LBS = I * Weight (formula from soeProcessor.js line 256)
          if (groupHPWeight) {
            const itemLBS = itemFT * groupHPWeight
            groupLBS += itemLBS
          }
        })

        // Show only the sum totals with formula range
        const firstRowNumber = Math.min(...group.map(item => item.rawRowNumber || 0))
        const lastRowNumber = Math.max(...group.map(item => item.rawRowNumber || 0))

      })

      if (hpCollectedGroups.length > 0) {
        // Process each HP group separately
        hpCollectedGroups.forEach((hpCollectedItems, groupIndex) => {
          if (hpCollectedItems.length > 0) {
            // Reset variables for each group
            let hpType = null
            let hpWeight = null
            let totalHeight = 0
            let heightCount = 0
            let totalCount = 0
            let totalFT = 0
            let totalLBS = 0

            // Helper function to parse dimension string (e.g., "24'-9"" -> 24.75)
            const parseDimension = (dimStr) => {
              if (!dimStr) return 0
              const match = dimStr.match(/(\d+)(?:'-?)?(\d+)?/)
              if (!match) return 0
              const feet = parseInt(match[1]) || 0
              const inches = parseInt(match[2]) || 0
              return feet + (inches / 12)
            }

            // Group items by particulars
            const groupedItems = new Map()
            hpCollectedItems.forEach(item => {
              const particulars = item.particulars || ''
              if (!groupedItems.has(particulars)) {
                groupedItems.set(particulars, {
                  particulars: particulars,
                  takeoff: 0,
                  hValue: null
                })
              }
              const group = groupedItems.get(particulars)
              group.takeoff += (item.takeoff || 0)
            })

            // Process each group to extract values
            groupedItems.forEach((group, particulars) => {
              // Extract HP type and weight (e.g., "HP12x63")
              if (!hpType || !hpWeight) {
                const hpMatch = particulars.match(/HP(\d+)x(\d+)/i)
                if (hpMatch) {
                  hpType = `HP${hpMatch[1]}x${hpMatch[2]}`
                  hpWeight = parseFloat(hpMatch[2]) // Weight is the number after 'x'
                }
              }

              // Extract H value (e.g., "H=24'-9"")
              const hMatch = particulars.match(/H=([0-9'"\-]+)/)
              if (hMatch) {
                const heightValue = parseDimension(hMatch[1])
                group.hValue = hMatch[1]

                // Calculate FT for this item: FT = H * C (height * takeoff)
                const itemFT = heightValue * group.takeoff
                totalFT += itemFT

                // Calculate LBS for this item: LBS = FT * weight
                if (hpWeight) {
                  const itemLBS = itemFT * hpWeight
                  totalLBS += itemLBS
                }

                // Add height for each takeoff (if takeoff is 8, add height 8 times)
                for (let i = 0; i < group.takeoff; i++) {
                  totalHeight += heightValue
                  heightCount++
                }
              }

              // Sum total count
              totalCount += group.takeoff
            })


            if (hpType && heightCount > 0) {
              // Calculate average height
              const avgHeight = totalHeight / heightCount

              // Round up to multiple of 5
              const roundToMultipleOf5 = (value) => {
                return Math.ceil(value / 5) * 5
              }
              const avgHeightRounded = roundToMultipleOf5(avgHeight)

              // Format height
              const heightFeet = Math.floor(avgHeightRounded)
              const heightInches = Math.round((avgHeightRounded - heightFeet) * 12)
              let heightText = ''
              if (heightInches === 0) {
                heightText = `${heightFeet}'-0"`
              } else {
                heightText = `${heightFeet}'-${heightInches}"`
              }

              // Get SOE page number from raw data (search for HP type)
              const getHPPageFromRawData = (hpTypeValue) => {
                if (!rawData || !Array.isArray(rawData) || rawData.length < 2) {
                  return 'SOE-B-100.00' // Default fallback
                }

                const headers = rawData[0]
                const dataRows = rawData.slice(1)

                const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
                const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

                if (digitizerIdx === -1 || pageIdx === -1) {
                  return 'SOE-B-100.00' // Default fallback
                }

                // Create pattern to match HP type (e.g., "HP12x63")
                const pattern = new RegExp(hpTypeValue.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i')

                for (const row of dataRows) {
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && pattern.test(String(digitizerItem))) {
                    const pageValue = row[pageIdx]
                    if (pageValue) {
                      const pageStr = String(pageValue).trim()
                      const soeMatch = pageStr.match(/SOE-[\w-]+/i)
                      if (soeMatch) {
                        return soeMatch[0]
                      }
                    }
                  }
                }

                return 'SOE-B-100.00' // Default fallback
              }

              const soePage = getHPPageFromRawData(hpType)

              // Find the sum row for this group
              // The sum formula is in the same row as the last item (not the next row)
              const lastRowNumber = Math.max(...hpCollectedItems.map(item => item.rawRowNumber || 0))
              const sumRowIndex = lastRowNumber // Sum row is the same as last item row

              // Create cell references to the sum row in calculation sheet
              // FT is in column I, LBS is in column K, QTY is in column M
              const calcSheetName = 'Calculations Sheet'
              const ftCellRef = `'${calcSheetName}'!I${sumRowIndex}`
              const lbsCellRef = `'${calcSheetName}'!K${sumRowIndex}`
              const qtyCellRef = `'${calcSheetName}'!M${sumRowIndex}`

              // Format the proposal text: F&I new (101)no [HP12x63] driven pile (Havg=42'-9") as per SOE-B-100.00
              const proposalText = `F&I new (${totalCount})no [${hpType}] driven pile (Havg=${heightText}) as per ${soePage}`


              spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  color: '#000000',
                  textAlign: 'left',
                  backgroundColor: 'white',
                  verticalAlign: 'top'
                },
                `${pfx}B${currentRow}`
              )

              // Add FT (LF) to column C - reference to calculation sheet sum row
              spreadsheet.updateCell({ formula: `=${ftCellRef}` }, `${pfx}C${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white'
                },
                `${pfx}C${currentRow}`
              )

              // Add LBS to column E - reference to calculation sheet sum row
              spreadsheet.updateCell({ formula: `=${lbsCellRef}` }, `${pfx}E${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white'
                },
                `${pfx}E${currentRow}`
              )

              // Add QTY to column G - reference to calculation sheet sum row (column M)
              spreadsheet.updateCell({ formula: `=${qtyCellRef}` }, `${pfx}G${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white'
                },
                `${pfx}G${currentRow}`
              )

              // Add 145 to second LF column (column I) for HP groups
              spreadsheet.updateCell({ value: 145 }, `${pfx}I${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: '#E2EFDA'
                },
                `${pfx}I${currentRow}:N${currentRow}`
              )

              // Add $/1000 formula in column H
              const dollarFormulaHP = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
              spreadsheet.updateCell({ formula: dollarFormulaHP }, `${pfx}H${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white',
                  format: '$#,##0.0'
                },
                `${pfx}H${currentRow}`
              )

              // Calculate and set row height based on content in column B
              const dynamicHeightHP = calculateRowHeight(proposalText)
              spreadsheet.setRowHeight(dynamicHeightHP, currentRow, proposalSheetIndex)

              currentRow++ // Move to next row for next group
              rowShift++ // Track that we've added a row
            }
          }
        })
      }

      // Update rowShift after HP piles are added
      // Now continue with other SOE subsections, accounting for all rows added so far

      // Add Primary secant piles proposal text below "Primary secant piles:" heading (row 43)
      // Try to get from processed items first, then fall back to soeSubsectionItems
      let primarySecantItems = window.primarySecantItems || []
      const primarySecantItemsMap = window.soeSubsectionItems || new Map()
      const primarySecantGroupsFromMap = primarySecantItemsMap.get('Primary secant piles') || []


      // If we have processed items, group them (similar to how soldier piles are grouped)
      // Otherwise, use groups from soeSubsectionItems
      let primarySecantGroups = []
      if (primarySecantItems.length > 0) {
        // Group items by empty rows (similar to drilled foundation piles)
        let currentGroup = []
        primarySecantItems.forEach((item, index) => {
          currentGroup.push(item)
          // If next item doesn't exist or we're at the end, save the group
          if (index === primarySecantItems.length - 1) {
            if (currentGroup.length > 0) {
              primarySecantGroups.push([...currentGroup])
              currentGroup = []
            }
          }
        })
        // If there's a remaining group, add it
        if (currentGroup.length > 0) {
          primarySecantGroups.push([...currentGroup])
        }
      } else if (primarySecantGroupsFromMap.length > 0) {
        primarySecantGroups = primarySecantGroupsFromMap
      }

      // Move to row after "Primary secant piles:" heading (row 43) for primary secant piles
      // If HP piles were added, currentRow is already after the last HP pile
      // If no HP piles, start primary secant piles at row 44 (right after "Primary secant piles:" at row 43)
      const primarySecantHeadingRow = headingRows['Primary secant piles:']
      if (hpCollectedGroups.length === 0 || currentRow <= primarySecantHeadingRow) {
        currentRow = primarySecantHeadingRow + 1 // Start primary secant piles right after heading
      } else {
        // HP piles were added and we're past the heading row
        // Ensure "Primary secant piles:" heading is visible before primary secant content
        spreadsheet.updateCell({ value: 'Primary secant piles:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            backgroundColor: '#D0CECE',
            textDecoration: 'underline',
            border: '1px solid #000000'
          },
          `${pfx}B${currentRow}`
        )
        spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
        currentRow++ // Move to next row for primary secant content
      }

      // If still no groups, try to find primary secant items from calculation sheet data
      if (primarySecantGroups.length === 0 && calculationData && calculationData.length > 0) {
        const calcSheetName = 'Calculations Sheet'
        let foundPrimarySecantRows = []
        let primarySecantHeaderRow = null
        let primarySecantSumRow = null

        // Search through calculationData for "Primary secant piles" subsection
        let inPrimarySecantSubsection = false
        calculationData.forEach((row, index) => {
          const rowNum = index + 2 // Excel row number (row 1 is header, so index 0 = row 2)
          const colB = row[1] // Column B (index 1)

          if (colB && typeof colB === 'string') {
            const bText = colB.trim()
            // Check if this is exactly the "Primary secant piles:" subsection header (case-insensitive)
            if (bText.toLowerCase() === 'primary secant piles:') {
              inPrimarySecantSubsection = true
              primarySecantHeaderRow = rowNum
              console.log('Found Primary secant piles: header at row', rowNum)
              return
            }

            // If we're in the subsection and hit another subsection or section, stop
            if (inPrimarySecantSubsection) {
              // Stop if we hit another subsection header (ends with ':')
              if (bText.endsWith(':') && bText.toLowerCase() !== 'primary secant piles:') {
                console.log('Stopping at subsection:', bText, 'row', rowNum)
                // The sum row should be the row before this new subsection
                if (foundPrimarySecantRows.length > 0) {
                  primarySecantSumRow = rowNum - 1
                }
                inPrimarySecantSubsection = false
                return
              }

              // This is a data row in the primary secant subsection
              // Only collect rows that have data (takeoff > 0) and don't end with ':'
              if (bText && !bText.endsWith(':')) {
                const takeoff = parseFloat(row[2]) || 0 // Column C
                if (takeoff > 0) {
                  foundPrimarySecantRows.push({
                    rowNum: rowNum,
                    particulars: bText,
                    takeoff: takeoff,
                    qty: parseFloat(row[12]) || 0, // Column M
                    height: parseFloat(row[7]) || 0, // Column H
                    rawRowNumber: rowNum
                  })
                }
              }
            }
          }
        })

        // If we found items but didn't hit another subsection, the sum row is after the last item
        if (inPrimarySecantSubsection && foundPrimarySecantRows.length > 0 && !primarySecantSumRow) {
          const lastItemRow = foundPrimarySecantRows[foundPrimarySecantRows.length - 1].rowNum
          primarySecantSumRow = lastItemRow + 1
        }

        console.log('Found primary secant rows:', foundPrimarySecantRows.length)
        console.log('Primary secant header row:', primarySecantHeaderRow)
        console.log('Primary secant sum row:', primarySecantSumRow)

        if (foundPrimarySecantRows.length > 0) {
          // Group them (all in one group for now) and store the sum row
          primarySecantGroups = [{
            items: foundPrimarySecantRows,
            sumRow: primarySecantSumRow
          }]
        }
      }

      if (primarySecantGroups.length > 0) {
        primarySecantGroups.forEach((group, groupIndex) => {
          // Handle both array format and object format with items property
          const groupItems = Array.isArray(group) ? group : (group.items || [])
          const groupSumRow = group.sumRow || null

          if (groupItems.length === 0) return

          // Calculate totals for the group
          let totalTakeoff = 0
          let totalQty = 0
          let totalFT = 0
          let totalHeight = 0
          let heightCount = 0
          let totalEmbedment = 0
          let embedmentCount = 0
          let lastRowNumber = 0
          let firstRowNumber = Infinity
          let diameter = null

          // Helper function to parse dimension string (e.g., "27'-10"" -> 27.833)
          const parseDimension = (dimStr) => {
            if (!dimStr) return 0
            const match = dimStr.match(/(\d+)(?:'-?)?(\d+)?/)
            if (!match) return 0
            const feet = parseInt(match[1]) || 0
            const inches = parseInt(match[2]) || 0
            return feet + (inches / 12)
          }

          groupItems.forEach(item => {
            totalTakeoff += item.takeoff || 0
            totalQty += item.qty || 0
            // Get height from parsed calculatedHeight or from calculation sheet
            const itemHeight = item.parsed?.calculatedHeight || item.height || 0
            const itemFT = itemHeight * (item.takeoff || 0)
            totalFT += itemFT
            totalHeight += itemHeight * (item.takeoff || 0)
            heightCount += (item.takeoff || 0)
            lastRowNumber = Math.max(lastRowNumber, item.rawRowNumber || 0)
            firstRowNumber = Math.min(firstRowNumber, item.rawRowNumber || Infinity)

            // Debug: log item structure if needed

            // Extract diameter from particulars if not set (just diameter, not thickness)
            if (!diameter) {
              const particulars = item.particulars || ''
              // Match patterns like "24" Ø" or "24Ø" or "24"Ø"
              const diameterMatch = particulars.match(/([0-9.]+)["\s]*Ø/i)
              if (diameterMatch) {
                diameter = parseFloat(diameterMatch[1])
              }
            }

            // Extract embedment (E=XX'-XX")
            const particulars = item.particulars || ''
            const eMatch = particulars.match(/E=([0-9'"\-]+)/i)
            if (eMatch) {
              const embedmentValue = parseDimension(eMatch[1])
              totalEmbedment += embedmentValue * (item.takeoff || 0)
              embedmentCount += (item.takeoff || 0)
            }
          })

          // Calculate average height
          const avgHeight = heightCount > 0 ? totalHeight / heightCount : 0
          const roundToMultipleOf5 = (value) => Math.ceil(value / 5) * 5
          const avgHeightRounded = roundToMultipleOf5(avgHeight)

          // Format height
          const heightFeet = Math.floor(avgHeightRounded)
          const heightInches = Math.round((avgHeightRounded - heightFeet) * 12)
          let heightText = ''
          if (heightInches === 0) {
            heightText = `${heightFeet}'-0"`
          } else {
            heightText = `${heightFeet}'-${heightInches}"`
          }

          // Calculate average embedment
          const avgEmbedment = embedmentCount > 0 ? totalEmbedment / embedmentCount : 0
          const avgEmbedmentRounded = roundToMultipleOf5(avgEmbedment)

          // Format embedment
          const embedmentFeet = Math.floor(avgEmbedmentRounded)
          const embedmentInches = Math.round((avgEmbedmentRounded - embedmentFeet) * 12)
          let embedmentText = ''
          if (embedmentInches === 0) {
            embedmentText = `${embedmentFeet}'-0"`
          } else {
            embedmentText = `${embedmentFeet}'-${embedmentInches}"`
          }

          // Get SOE page references from raw data
          let soePageMain = 'SOE-100.00'
          let soePageDetails = 'SOE-200.00'

          if (rawData && Array.isArray(rawData) && rawData.length > 1) {
            const headers = rawData[0]
            const dataRows = rawData.slice(1)
            const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
            const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

            if (digitizerIdx !== -1 && pageIdx !== -1) {
              // Search for primary secant pile items
              for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                const row = dataRows[rowIndex]
                const digitizerItem = row[digitizerIdx]
                if (digitizerItem && typeof digitizerItem === 'string') {
                  const itemText = digitizerItem.toLowerCase()
                  if (itemText.includes('primary secant')) {
                    const pageValue = row[pageIdx]
                    if (pageValue) {
                      const pageStr = String(pageValue).trim()
                      // Extract all SOE references
                      const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                      if (soeMatches && soeMatches.length > 0) {
                        if (!soePageMain || soePageMain === 'SOE-100.00') {
                          soePageMain = soeMatches[0]
                        }
                        if (soeMatches.length > 1 && (!soePageDetails || soePageDetails === 'SOE-200.00')) {
                          soePageDetails = soeMatches[1]
                        }
                      }
                    }
                  }
                }
              }
            }
          }

          // Find the sum row for this group
          // First try to use the sum row from the search, otherwise find it in the calculation sheet
          let sumRowIndex = groupSumRow

          // If we don't have a sum row from search, try to find it in the calculation sheet
          if (!sumRowIndex && calculationData && calculationData.length > 0) {
            // First, search backwards from firstRowNumber to find "Primary secant piles:" header
            let primarySecantHeaderFound = false
            for (let i = Math.max(0, firstRowNumber - 20); i < firstRowNumber; i++) {
              const rowData = calculationData[i - 1] // calculationData is 0-indexed
              const colB = rowData?.[1]
              if (colB && typeof colB === 'string' && colB.trim().toLowerCase() === 'primary secant piles:') {
                primarySecantHeaderFound = true
                console.log('Found Primary secant piles: header at row', i + 1, 'by searching backwards')
                break
              }
            }

            // Search for the sum row after the last item row
            // The sum row should be the first row after lastRowNumber that has a formula in column I or M
            // or is an empty row before the next subsection
            for (let i = lastRowNumber; i < Math.min(calculationData.length + 1, lastRowNumber + 10); i++) {
              const rowData = calculationData[i - 1] // calculationData is 0-indexed
              const colB = rowData?.[1]
              const colI = rowData?.[8] // Column I
              const colM = rowData?.[12] // Column M

              // Check if this row is empty in column B (sum rows are usually empty in B)
              // and has a value or formula in column I or M
              if ((!colB || colB === '') && (colI !== undefined && colI !== '' || colM !== undefined && colM !== '')) {
                sumRowIndex = i + 1 // Convert to 1-indexed
                console.log('Found sum row at:', sumRowIndex, 'colI:', colI, 'colM:', colM)
                break
              }

              // Also check if we hit another subsection header
              if (colB && typeof colB === 'string' && colB.trim().endsWith(':')) {
                // The sum row should be the row before this subsection
                sumRowIndex = i
                console.log('Found sum row before subsection:', colB, 'at row:', sumRowIndex)
                break
              }
            }

            // Fallback: if still not found, use lastRowNumber + 1
            if (!sumRowIndex) {
              sumRowIndex = lastRowNumber + 1
              console.log('Using fallback sum row:', sumRowIndex)
            }
          } else if (!sumRowIndex) {
            // Fallback: use lastRowNumber + 1
            sumRowIndex = lastRowNumber + 1
            console.log('Using fallback sum row (no calculationData):', sumRowIndex)
          }

          const calcSheetName = 'Calculations Sheet'

          console.log('=== Primary Secant Piles Debug ===')
          console.log('firstRowNumber:', firstRowNumber)
          console.log('lastRowNumber:', lastRowNumber)
          console.log('sumRowIndex:', sumRowIndex)
          console.log('groupSumRow from search:', groupSumRow)
          console.log('FT formula:', `='${calcSheetName}'!I${sumRowIndex}`)
          console.log('QTY formula:', `='${calcSheetName}'!M${sumRowIndex}`)
          console.log('Group items:', groupItems.map(item => ({
            rawRowNumber: item.rawRowNumber,
            particulars: item.particulars,
            takeoff: item.takeoff,
            qty: item.qty
          })))

          // Check what's actually in the calculation sheet at the sum row
          if (calculationData && calculationData.length >= sumRowIndex) {
            const sumRowData = calculationData[sumRowIndex - 1] // calculationData is 0-indexed, sumRowIndex is 1-indexed
            console.log('Sum row data from calculation sheet:', {
              rowIndex: sumRowIndex - 1,
              colB: sumRowData?.[1],
              colI: sumRowData?.[8], // Column I (index 8)
              colM: sumRowData?.[12] // Column M (index 12)
            })
          }

          // Also check the rows around the sum row
          if (calculationData && calculationData.length >= sumRowIndex + 2) {
            console.log('Rows around sum row:')
            for (let i = Math.max(0, sumRowIndex - 3); i < Math.min(calculationData.length, sumRowIndex + 2); i++) {
              const rowData = calculationData[i]
              console.log(`Row ${i + 1}:`, {
                colB: rowData?.[1],
                colI: rowData?.[8],
                colM: rowData?.[12]
              })
            }
          }

          // Format proposal text: F&I new (#)no [#" Ø] primary secant piles (Havg=#'-#", #'-#" embedment) as per SOE-#.##
          // Use totalQty if available, otherwise use totalTakeoff
          const qtyValue = Math.round(totalQty || totalTakeoff)
          let proposalText = ''
          if (diameter) {
            proposalText = `F&I new (${qtyValue})no [${diameter}" Ø] primary secant piles (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain}`
          } else {
            proposalText = `F&I new (${qtyValue})no [#" Ø] primary secant piles (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain}`
          }


          spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: 'white',
              verticalAlign: 'top'
            },
            `${pfx}B${currentRow}`
          )

          // Add FT (LF) to column C - reference to calculation sheet sum row
          if (sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${sumRowIndex}` }, `${pfx}C${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}C${currentRow}`
            )
          }

          // Add QTY to column G - reference to calculation sheet sum row (column M)
          if (sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}G${currentRow}`
            )
          }

          // Add 180 to second LF column (column I) - same as drilled soldier piles
          spreadsheet.updateCell({ value: 180 }, `${pfx}I${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: '#E2EFDA'
            },
            `${pfx}I${currentRow}:N${currentRow}`
          )

          // Add $/1000 formula in column H
          const dollarFormula = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
          spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: 'white',
              format: '$#,##0.0'
            },
            `${pfx}H${currentRow}`
          )

          // Apply currency format
          try {
            spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}:N${currentRow}`)
          } catch (e) {
            // Fallback already applied in cellFormat
          }

          // Calculate and set row height based on content in column B
          const proposalTextForHeight = proposalText || (spreadsheet.getCellValue ? spreadsheet.getCellValue(`${pfx}B${currentRow}`) : '')
          const dynamicHeight = calculateRowHeight(proposalTextForHeight)
          spreadsheet.setRowHeight(dynamicHeight, currentRow, proposalSheetIndex)

          currentRow++ // Move to next row for next group
        })
      }

      // Add Secondary secant piles proposal text below "Secondary secant piles:" heading
      // Try to get from processed items first, then fall back to soeSubsectionItems
      let secondarySecantItems = window.secondarySecantItems || []
      const secondarySecantItemsMap = window.soeSubsectionItems || new Map()
      const secondarySecantGroupsFromMap = secondarySecantItemsMap.get('Secondary secant piles') || []

      // If we have processed items, group them (similar to how soldier piles are grouped)
      // Otherwise, use groups from soeSubsectionItems
      let secondarySecantGroups = []
      if (secondarySecantItems.length > 0) {
        // Group items by empty rows (similar to drilled foundation piles)
        let currentGroup = []
        secondarySecantItems.forEach((item, index) => {
          currentGroup.push(item)
          // If next item doesn't exist or we're at the end, save the group
          if (index === secondarySecantItems.length - 1) {
            if (currentGroup.length > 0) {
              secondarySecantGroups.push([...currentGroup])
              currentGroup = []
            }
          }
        })
        // If there's a remaining group, add it
        if (currentGroup.length > 0) {
          secondarySecantGroups.push([...currentGroup])
        }
      } else if (secondarySecantGroupsFromMap.length > 0) {
        secondarySecantGroups = secondarySecantGroupsFromMap
      }

      // Move to row after "Secondary secant piles:" heading for secondary secant piles
      const secondarySecantHeadingRow = headingRows['Secondary secant piles:']
      if (currentRow <= secondarySecantHeadingRow) {
        currentRow = secondarySecantHeadingRow + 1 // Start secondary secant piles right after heading
      } else {
        // Ensure "Secondary secant piles:" heading is visible before secondary secant content
        spreadsheet.updateCell({ value: 'Secondary secant piles:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            backgroundColor: '#D0CECE',
            textDecoration: 'underline',
            border: '1px solid #000000'
          },
          `${pfx}B${currentRow}`
        )
        spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
        currentRow++ // Move to next row for secondary secant content
      }

      // If still no groups, try to find secondary secant items from calculation sheet data
      if (secondarySecantGroups.length === 0 && calculationData && calculationData.length > 0) {
        const calcSheetName = 'Calculations Sheet'
        let foundSecondarySecantRows = []

        // Search through calculationData for "Secondary secant piles" subsection
        let inSecondarySecantSubsection = false
        calculationData.forEach((row, index) => {
          const rowNum = index + 2 // Excel row number (row 1 is header, so index 0 = row 2)
          const colB = row[1] // Column B (index 1)

          if (colB && typeof colB === 'string') {
            const bText = colB.trim()
            // Check if this is the "Secondary secant piles" subsection header
            if (bText.toLowerCase().includes('secondary secant piles') && bText.endsWith(':')) {
              inSecondarySecantSubsection = true
              return
            }

            // If we're in the subsection and hit another subsection or section, stop
            if (inSecondarySecantSubsection) {
              if (bText.endsWith(':') && !bText.toLowerCase().includes('secondary secant')) {
                inSecondarySecantSubsection = false
                return
              }

              // This is a data row in the secondary secant subsection
              if (bText && !bText.endsWith(':') && row[2] !== undefined) {
                const takeoff = parseFloat(row[2]) || 0 // Column C
                if (takeoff > 0 || bText.toLowerCase().includes('secondary secant')) {
                  foundSecondarySecantRows.push({
                    rowNum: rowNum,
                    particulars: bText,
                    takeoff: takeoff,
                    qty: parseFloat(row[12]) || 0, // Column M
                    height: parseFloat(row[7]) || 0, // Column H
                    rawRowNumber: rowNum
                  })
                }
              }
            }
          }
        })

        if (foundSecondarySecantRows.length > 0) {
          // Group them (all in one group for now)
          secondarySecantGroups = [foundSecondarySecantRows]
        }
      }

      if (secondarySecantGroups.length > 0) {
        secondarySecantGroups.forEach((group, groupIndex) => {
          if (group.length === 0) return

          // Calculate totals for the group
          let totalTakeoff = 0
          let totalQty = 0
          let totalFT = 0
          let totalHeight = 0
          let heightCount = 0
          let totalEmbedment = 0
          let embedmentCount = 0
          let lastRowNumber = 0
          let diameter = null

          // Helper function to parse dimension string (e.g., "27'-10"" -> 27.833)
          const parseDimension = (dimStr) => {
            if (!dimStr) return 0
            const match = dimStr.match(/(\d+)(?:'-?)?(\d+)?/)
            if (!match) return 0
            const feet = parseInt(match[1]) || 0
            const inches = parseInt(match[2]) || 0
            return feet + (inches / 12)
          }

          group.forEach(item => {
            totalTakeoff += item.takeoff || 0
            totalQty += item.qty || 0
            // Get height from parsed calculatedHeight or from calculation sheet
            const itemHeight = item.parsed?.calculatedHeight || item.height || 0
            const itemFT = itemHeight * (item.takeoff || 0)
            totalFT += itemFT
            totalHeight += itemHeight * (item.takeoff || 0)
            heightCount += (item.takeoff || 0)
            lastRowNumber = Math.max(lastRowNumber, item.rawRowNumber || 0)

            // Extract diameter from particulars if not set (just diameter, not thickness)
            if (!diameter) {
              const particulars = item.particulars || ''
              // Match patterns like "24" Ø" or "24Ø" or "24"Ø"
              const diameterMatch = particulars.match(/([0-9.]+)["\s]*Ø/i)
              if (diameterMatch) {
                diameter = parseFloat(diameterMatch[1])
              }
            }

            // Extract embedment (E=XX'-XX")
            const particulars = item.particulars || ''
            const eMatch = particulars.match(/E=([0-9'"\-]+)/i)
            if (eMatch) {
              const embedmentValue = parseDimension(eMatch[1])
              totalEmbedment += embedmentValue * (item.takeoff || 0)
              embedmentCount += (item.takeoff || 0)
            }
          })

          // Calculate average height
          const avgHeight = heightCount > 0 ? totalHeight / heightCount : 0
          const roundToMultipleOf5 = (value) => Math.ceil(value / 5) * 5
          const avgHeightRounded = roundToMultipleOf5(avgHeight)

          // Format height
          const heightFeet = Math.floor(avgHeightRounded)
          const heightInches = Math.round((avgHeightRounded - heightFeet) * 12)
          let heightText = ''
          if (heightInches === 0) {
            heightText = `${heightFeet}'-0"`
          } else {
            heightText = `${heightFeet}'-${heightInches}"`
          }

          // Calculate average embedment
          const avgEmbedment = embedmentCount > 0 ? totalEmbedment / embedmentCount : 0
          const avgEmbedmentRounded = roundToMultipleOf5(avgEmbedment)

          // Format embedment
          const embedmentFeet = Math.floor(avgEmbedmentRounded)
          const embedmentInches = Math.round((avgEmbedmentRounded - embedmentFeet) * 12)
          let embedmentText = ''
          if (embedmentInches === 0) {
            embedmentText = `${embedmentFeet}'-0"`
          } else {
            embedmentText = `${embedmentFeet}'-${embedmentInches}"`
          }

          // Get SOE page references from raw data
          let soePageMain = 'SOE-100.00'

          if (rawData && Array.isArray(rawData) && rawData.length > 1) {
            const headers = rawData[0]
            const dataRows = rawData.slice(1)
            const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
            const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

            if (digitizerIdx !== -1 && pageIdx !== -1) {
              // Search for secondary secant pile items
              for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                const row = dataRows[rowIndex]
                const digitizerItem = row[digitizerIdx]
                if (digitizerItem && typeof digitizerItem === 'string') {
                  const itemText = digitizerItem.toLowerCase()
                  if (itemText.includes('secondary secant')) {
                    const pageValue = row[pageIdx]
                    if (pageValue) {
                      const pageStr = String(pageValue).trim()
                      // Extract all SOE references
                      const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                      if (soeMatches && soeMatches.length > 0) {
                        if (!soePageMain || soePageMain === 'SOE-100.00') {
                          soePageMain = soeMatches[0]
                        }
                      }
                    }
                  }
                }
              }
            }
          }

          // Find the sum row for this group
          const sumRowIndex = lastRowNumber
          const calcSheetName = 'Calculations Sheet'

          // Format proposal text: F&I new (#)no [#" Ø] secondary secant piles (Havg=#'-#", #'-#" embedment) as per SOE-#.##
          // Use totalQty if available, otherwise use totalTakeoff
          const qtyValue = Math.round(totalQty || totalTakeoff)
          let proposalText = ''
          if (diameter) {
            proposalText = `F&I new (${qtyValue})no [${diameter}" Ø] secondary secant piles (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain}`
          } else {
            proposalText = `F&I new (${qtyValue})no [#" Ø] secondary secant piles (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain}`
          }

          spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: 'white',
              verticalAlign: 'top'
            },
            `${pfx}B${currentRow}`
          )

          // Add FT (LF) to column C - reference to calculation sheet sum row
          if (sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${sumRowIndex}` }, `${pfx}C${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}C${currentRow}`
            )
          }

          // Add QTY to column G - reference to calculation sheet sum row (column M)
          if (sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}G${currentRow}`
            )
          }

          // Add 180 to second LF column (column I) - same as primary secant piles
          spreadsheet.updateCell({ value: 180 }, `${pfx}I${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: '#E2EFDA'
            },
            `${pfx}I${currentRow}:N${currentRow}`
          )

          // Add $/1000 formula in column H
          const dollarFormula = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
          spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: 'white',
              format: '$#,##0.0'
            },
            `${pfx}H${currentRow}`
          )

          // Apply currency format
          try {
            spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}:N${currentRow}`)
          } catch (e) {
            // Fallback already applied in cellFormat
          }

          // Calculate and set row height based on content in column B
          const proposalTextForHeight = proposalText || (spreadsheet.getCellValue ? spreadsheet.getCellValue(`${pfx}B${currentRow}`) : '')
          const dynamicHeight = calculateRowHeight(proposalTextForHeight)
          spreadsheet.setRowHeight(dynamicHeight, currentRow, proposalSheetIndex)

          currentRow++ // Move to next row for next group
        })
      }

      // Add Tangent pile proposal text below "Tangent pile:" heading
      // Try to get from processed items first, then fall back to soeSubsectionItems
      let tangentPileItems = window.tangentPileItems || []
      const tangentPileItemsMap = window.soeSubsectionItems || new Map()
      const tangentPileGroupsFromMap = tangentPileItemsMap.get('Tangent piles') || tangentPileItemsMap.get('Tangent pile') || []

      // If we have processed items, group them (similar to how soldier piles are grouped)
      // Otherwise, use groups from soeSubsectionItems
      let tangentPileGroups = []
      if (tangentPileItems.length > 0) {
        // Group items by empty rows (similar to drilled foundation piles)
        let currentGroup = []
        tangentPileItems.forEach((item, index) => {
          currentGroup.push(item)
          // If next item doesn't exist or we're at the end, save the group
          if (index === tangentPileItems.length - 1) {
            if (currentGroup.length > 0) {
              tangentPileGroups.push([...currentGroup])
              currentGroup = []
            }
          }
        })
        // If there's a remaining group, add it
        if (currentGroup.length > 0) {
          tangentPileGroups.push([...currentGroup])
        }
      } else if (tangentPileGroupsFromMap.length > 0) {
        tangentPileGroups = tangentPileGroupsFromMap
      }

      // Move to row after "Tangent pile:" heading for tangent piles
      const tangentPileHeadingRow = headingRows['Tangent pile:']
      if (currentRow <= tangentPileHeadingRow) {
        currentRow = tangentPileHeadingRow + 1 // Start tangent piles right after heading
      } else {
        // Ensure "Tangent pile:" heading is visible before tangent pile content
        spreadsheet.updateCell({ value: 'Tangent pile:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            backgroundColor: '#D0CECE',
            textDecoration: 'underline',
            border: '1px solid #000000'
          },
          `${pfx}B${currentRow}`
        )
        spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
        currentRow++ // Move to next row for tangent pile content
      }

      // If still no groups, try to find tangent pile items from calculation sheet data
      if (tangentPileGroups.length === 0 && calculationData && calculationData.length > 0) {
        const calcSheetName = 'Calculations Sheet'
        let foundTangentPileRows = []

        // Search through calculationData for "Tangent pile" or "Tangent piles" subsection
        let inTangentPileSubsection = false
        calculationData.forEach((row, index) => {
          const rowNum = index + 2 // Excel row number (row 1 is header, so index 0 = row 2)
          const colB = row[1] // Column B (index 1)

          if (colB && typeof colB === 'string') {
            const bText = colB.trim()
            // Check if this is the "Tangent pile" or "Tangent piles" subsection header
            if ((bText.toLowerCase().includes('tangent pile') || bText.toLowerCase().includes('tangent piles')) && bText.endsWith(':')) {
              inTangentPileSubsection = true
              return
            }

            // If we're in the subsection and hit another subsection or section, stop
            if (inTangentPileSubsection) {
              if (bText.endsWith(':') && !bText.toLowerCase().includes('tangent')) {
                inTangentPileSubsection = false
                return
              }

              // This is a data row in the tangent pile subsection
              if (bText && !bText.endsWith(':') && row[2] !== undefined) {
                const takeoff = parseFloat(row[2]) || 0 // Column C
                if (takeoff > 0 || bText.toLowerCase().includes('tangent')) {
                  foundTangentPileRows.push({
                    rowNum: rowNum,
                    particulars: bText,
                    takeoff: takeoff,
                    qty: parseFloat(row[12]) || 0, // Column M
                    height: parseFloat(row[7]) || 0, // Column H
                    rawRowNumber: rowNum
                  })
                }
              }
            }
          }
        })

        if (foundTangentPileRows.length > 0) {
          // Group them (all in one group for now)
          tangentPileGroups = [foundTangentPileRows]
        }
      }

      if (tangentPileGroups.length > 0) {
        tangentPileGroups.forEach((group, groupIndex) => {
          if (group.length === 0) return

          // Calculate totals for the group
          let totalTakeoff = 0
          let totalQty = 0
          let totalFT = 0
          let totalHeight = 0
          let heightCount = 0
          let totalEmbedment = 0
          let embedmentCount = 0
          let lastRowNumber = 0
          let diameter = null

          // Helper function to parse dimension string (e.g., "27'-10"" -> 27.833)
          const parseDimension = (dimStr) => {
            if (!dimStr) return 0
            const match = dimStr.match(/(\d+)(?:'-?)?(\d+)?/)
            if (!match) return 0
            const feet = parseInt(match[1]) || 0
            const inches = parseInt(match[2]) || 0
            return feet + (inches / 12)
          }

          group.forEach(item => {
            totalTakeoff += item.takeoff || 0
            totalQty += item.qty || 0
            // Get height from parsed calculatedHeight or from calculation sheet
            const itemHeight = item.parsed?.calculatedHeight || item.height || 0
            const itemFT = itemHeight * (item.takeoff || 0)
            totalFT += itemFT
            totalHeight += itemHeight * (item.takeoff || 0)
            heightCount += (item.takeoff || 0)
            lastRowNumber = Math.max(lastRowNumber, item.rawRowNumber || 0)

            // Extract diameter from particulars if not set (just diameter, not thickness)
            if (!diameter) {
              const particulars = item.particulars || ''
              // Match patterns like "24" Ø" or "24Ø" or "24"Ø"
              const diameterMatch = particulars.match(/([0-9.]+)["\s]*Ø/i)
              if (diameterMatch) {
                diameter = parseFloat(diameterMatch[1])
              }
            }

            // Extract embedment (E=XX'-XX")
            const particulars = item.particulars || ''
            const eMatch = particulars.match(/E=([0-9'"\-]+)/i)
            if (eMatch) {
              const embedmentValue = parseDimension(eMatch[1])
              totalEmbedment += embedmentValue * (item.takeoff || 0)
              embedmentCount += (item.takeoff || 0)
            }
          })

          // Calculate average height
          const avgHeight = heightCount > 0 ? totalHeight / heightCount : 0
          const roundToMultipleOf5 = (value) => Math.ceil(value / 5) * 5
          const avgHeightRounded = roundToMultipleOf5(avgHeight)

          // Format height
          const heightFeet = Math.floor(avgHeightRounded)
          const heightInches = Math.round((avgHeightRounded - heightFeet) * 12)
          let heightText = ''
          if (heightInches === 0) {
            heightText = `${heightFeet}'-0"`
          } else {
            heightText = `${heightFeet}'-${heightInches}"`
          }

          // Calculate average embedment
          const avgEmbedment = embedmentCount > 0 ? totalEmbedment / embedmentCount : 0
          const avgEmbedmentRounded = roundToMultipleOf5(avgEmbedment)

          // Format embedment
          const embedmentFeet = Math.floor(avgEmbedmentRounded)
          const embedmentInches = Math.round((avgEmbedmentRounded - embedmentFeet) * 12)
          let embedmentText = ''
          if (embedmentInches === 0) {
            embedmentText = `${embedmentFeet}'-0"`
          } else {
            embedmentText = `${embedmentFeet}'-${embedmentInches}"`
          }

          // Get SOE page references from raw data
          let soePageMain = 'SOE-100.00'

          if (rawData && Array.isArray(rawData) && rawData.length > 1) {
            const headers = rawData[0]
            const dataRows = rawData.slice(1)
            const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
            const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

            if (digitizerIdx !== -1 && pageIdx !== -1) {
              // Search for tangent pile items
              for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                const row = dataRows[rowIndex]
                const digitizerItem = row[digitizerIdx]
                if (digitizerItem && typeof digitizerItem === 'string') {
                  const itemText = digitizerItem.toLowerCase()
                  if (itemText.includes('tangent')) {
                    const pageValue = row[pageIdx]
                    if (pageValue) {
                      const pageStr = String(pageValue).trim()
                      // Extract all SOE references
                      const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                      if (soeMatches && soeMatches.length > 0) {
                        if (!soePageMain || soePageMain === 'SOE-100.00') {
                          soePageMain = soeMatches[0]
                        }
                      }
                    }
                  }
                }
              }
            }
          }

          // Find the sum row for this group
          const sumRowIndex = lastRowNumber
          const calcSheetName = 'Calculations Sheet'

          // Format proposal text: F&I new (#)no [#" Ø] tangent pile (Havg=#'-#", #'-#" embedment) as per SOE-#.##
          // Use totalQty if available, otherwise use totalTakeoff
          const qtyValue = Math.round(totalQty || totalTakeoff)
          let proposalText = ''
          if (diameter) {
            proposalText = `F&I new (${qtyValue})no [${diameter}" Ø] tangent pile (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain}`
          } else {
            proposalText = `F&I new (${qtyValue})no [#" Ø] tangent pile (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain}`
          }

          spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: 'white',
              verticalAlign: 'top'
            },
            `${pfx}B${currentRow}`
          )

          // Add FT (LF) to column C - reference to calculation sheet sum row
          if (sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${sumRowIndex}` }, `${pfx}C${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}C${currentRow}`
            )
          }

          // Add QTY to column G - reference to calculation sheet sum row (column M)
          if (sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}G${currentRow}`
            )
          }

          // Add 180 to second LF column (column I) - same as primary secant piles
          spreadsheet.updateCell({ value: 180 }, `${pfx}I${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: '#E2EFDA'
            },
            `${pfx}I${currentRow}:N${currentRow}`
          )

          // Add $/1000 formula in column H
          const dollarFormula = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
          spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: 'white',
              format: '$#,##0.0'
            },
            `${pfx}H${currentRow}`
          )

          // Apply currency format
          try {
            spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}:N${currentRow}`)
          } catch (e) {
            // Fallback already applied in cellFormat
          }

          // Calculate and set row height based on content in column B
          const proposalTextForHeight = proposalText || (spreadsheet.getCellValue ? spreadsheet.getCellValue(`${pfx}B${currentRow}`) : '')
          const dynamicHeight = calculateRowHeight(proposalTextForHeight)
          spreadsheet.setRowHeight(dynamicHeight, currentRow, proposalSheetIndex)

          currentRow++ // Move to next row for next group
        })
      }

      // Add SOE subsection headers that come after soldier pile subsections
      // These should be displayed as grey headers, not processed groups
      const soeSubsectionItems = window.soeSubsectionItems || new Map()
      const subsectionOrder = [
        'Primary secant piles',
        'Secondary secant piles',
        'Tangent piles',
        'Tangent pile',
        'Sheet pile',
        'Sheet piles',
        'Timber lagging',
        'Timber sheeting',
        'Vertical timber sheets',
        'Horizontal timber sheets',
        'Timber stringer',
        'Bracing',
        'Tie back',
        'Tie back anchor',
        'Tie down',
        'Tie down anchor',
        'Parging',
        'Heel blocks',
        'Underpinning',
        'Concrete soil retention piers',
        'Concrete soil retention pier',
        'Guide wall',
        'Concrete buttons',
        'Dowels',
        'Rock pins',
        'Rock stabilization',
        'Shotcrete',
        'Permission grouting',
        'Form board',
        'Drilled hole grout',
        'Mud slab',
        'Misc.'
      ]

      // Subsections to completely exclude from display
      const excludedSubsections = new Set([
        'Form board',
        'Buttons',
        'Dowel bar',
        'Backpacking',
        'Upper Raker',
        'Upper raker',
        'Lower Raker',
        'Lower raker',
        'Stud beam',
        'Stub beam',
        'Rock anchors',
        'Rock bolts',
        'Tie down anchor'
      ])

      // Get all unique subsection names from collected items
      // Exclude "Primary secant piles" since it's already processed above
      const collectedSubsections = new Set()
      soeSubsectionItems.forEach((groups, name) => {
        if (groups.length > 0 && name !== 'Primary secant piles') {
          // Map "Dowel bar" to "Dowels" for consistency
          if (name.toLowerCase() === 'dowel bar') {
            collectedSubsections.add('Dowels')
          } else if (name.toLowerCase() === 'buttons' || name.toLowerCase().includes('button')) {
            // Map "Buttons" to "Concrete buttons" for consistency
            collectedSubsections.add('Concrete buttons')
          } else {
            collectedSubsections.add(name)
          }
        }
      })

      // Check if parging is in soeSubsectionItems
      const pargingGroups = soeSubsectionItems.get('Parging') || []
      const pargingItemsFromWindow = window.pargingItems || []
      const hasPargingItems = pargingGroups.length > 0 && pargingGroups.some(g => g.length > 0) || pargingItemsFromWindow.length > 0

      // If parging items exist but Parging is not in collectedSubsections, add it
      if (hasPargingItems && !collectedSubsections.has('Parging')) {
        collectedSubsections.add('Parging')
      }

      // Check if guide wall is in soeSubsectionItems or window
      const guideWallGroups = soeSubsectionItems.get('Guide wall') || soeSubsectionItems.get('Guilde wall') || []
      const guideWallItemsFromWindow = window.guideWallItems || []
      const hasGuideWallItems = guideWallGroups.length > 0 && guideWallGroups.some(g => g.length > 0) || guideWallItemsFromWindow.length > 0

      // If guide wall items exist but Guide wall is not in collectedSubsections, add it
      if (hasGuideWallItems && !collectedSubsections.has('Guide wall')) {
        collectedSubsections.add('Guide wall')
      }

      // Check if dowels/dowel bar is in soeSubsectionItems or window
      const dowelBarGroups = soeSubsectionItems.get('Dowel bar') || soeSubsectionItems.get('Dowels') || []
      const dowelBarItemsFromWindow = window.dowelBarItems || []
      const hasDowelBarItems = dowelBarGroups.length > 0 && dowelBarGroups.some(g => g.length > 0) || dowelBarItemsFromWindow.length > 0

      // If dowel bar items exist but Dowels is not in collectedSubsections, add it
      if (hasDowelBarItems && !collectedSubsections.has('Dowels')) {
        collectedSubsections.add('Dowels')
      }

      // Check if rock pins is in soeSubsectionItems or window
      const rockPinGroups = soeSubsectionItems.get('Rock pins') || soeSubsectionItems.get('Rock pin') || []
      const rockPinItemsFromWindow = window.rockPinItems || []
      const hasRockPinItems = rockPinGroups.length > 0 && rockPinGroups.some(g => g.length > 0) || rockPinItemsFromWindow.length > 0

      // If rock pin items exist but Rock pins is not in collectedSubsections, add it
      if (hasRockPinItems && !collectedSubsections.has('Rock pins')) {
        collectedSubsections.add('Rock pins')
      }

      // Check if rock stabilization is in soeSubsectionItems or window
      const rockStabilizationGroups = soeSubsectionItems.get('Rock stabilization') || []
      const rockStabilizationItemsFromWindow = window.rockStabilizationItems || []
      const hasRockStabilizationItems = rockStabilizationGroups.length > 0 && rockStabilizationGroups.some(g => g.length > 0) || rockStabilizationItemsFromWindow.length > 0

      // If rock stabilization items exist but Rock stabilization is not in collectedSubsections, add it
      if (hasRockStabilizationItems && !collectedSubsections.has('Rock stabilization')) {
        collectedSubsections.add('Rock stabilization')
      }

      // Check if shotcrete is in soeSubsectionItems or window
      const shotcreteGroups = soeSubsectionItems.get('Shotcrete') || []
      const shotcreteItemsFromWindow = window.shotcreteItems || []
      const hasShotcreteItems = shotcreteGroups.length > 0 && shotcreteGroups.some(g => g.length > 0) || shotcreteItemsFromWindow.length > 0

      // If shotcrete items exist but Shotcrete is not in collectedSubsections, add it
      if (hasShotcreteItems && !collectedSubsections.has('Shotcrete')) {
        collectedSubsections.add('Shotcrete')
      }

      // Also check calculationData directly for shotcrete
      if (!hasShotcreteItems && calculationData && calculationData.length > 0) {
        let foundShotcrete = false
        for (const row of calculationData) {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('shotcrete') && (bText.endsWith(':') || parseFloat(row[2]) > 0)) {
              foundShotcrete = true
              break
            }
          }
        }
        if (foundShotcrete && !collectedSubsections.has('Shotcrete')) {
          collectedSubsections.add('Shotcrete')
        }
      }

      // Check if permission grouting is in soeSubsectionItems or window
      const permissionGroutingGroups = soeSubsectionItems.get('Permission grouting') || []
      const permissionGroutingItemsFromWindow = window.permissionGroutingItems || []
      const hasPermissionGroutingItems = permissionGroutingGroups.length > 0 && permissionGroutingGroups.some(g => g.length > 0) || permissionGroutingItemsFromWindow.length > 0

      // If permission grouting items exist but Permission grouting is not in collectedSubsections, add it
      if (hasPermissionGroutingItems && !collectedSubsections.has('Permission grouting')) {
        collectedSubsections.add('Permission grouting')
      }

      // Also check calculationData directly for permission grouting
      if (!hasPermissionGroutingItems && calculationData && calculationData.length > 0) {
        let foundPermissionGrouting = false
        for (const row of calculationData) {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('permission grouting') && (bText.endsWith(':') || parseFloat(row[2]) > 0)) {
              foundPermissionGrouting = true
              break
            }
          }
        }
        if (foundPermissionGrouting && !collectedSubsections.has('Permission grouting')) {
          collectedSubsections.add('Permission grouting')
        }
      }

      // Check if mud slab is in soeSubsectionItems or window
      const mudSlabGroups = soeSubsectionItems.get('Mud slab') || []
      const mudSlabItemsFromWindow = window.mudSlabItems || []
      const hasMudSlabItems = mudSlabGroups.length > 0 && mudSlabGroups.some(g => g.length > 0) || mudSlabItemsFromWindow.length > 0

      // If mud slab items exist but Mud slab is not in collectedSubsections, add it
      if (hasMudSlabItems && !collectedSubsections.has('Mud slab')) {
        collectedSubsections.add('Mud slab')
      }

      // Also check calculationData directly for mud slab
      if (!hasMudSlabItems && calculationData && calculationData.length > 0) {
        let foundMudSlab = false
        for (const row of calculationData) {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('mud slab') && (bText.endsWith(':') || parseFloat(row[2]) > 0)) {
              foundMudSlab = true
              break
            }
          }
        }
        if (foundMudSlab && !collectedSubsections.has('Mud slab')) {
          collectedSubsections.add('Mud slab')
        }
      }

      // Always add Misc. subsection to collectedSubsections
      if (!collectedSubsections.has('Misc.') && !collectedSubsections.has('Misc')) {
        collectedSubsections.add('Misc.')
      }

      // Check if sheet pile is in soeSubsectionItems or window
      const sheetPileGroups = soeSubsectionItems.get('Sheet pile') || soeSubsectionItems.get('Sheet piles') || []
      const sheetPileItemsFromWindow = window.sheetPileItems || []
      const hasSheetPileItems = sheetPileGroups.length > 0 && sheetPileGroups.some(g => g.length > 0) || sheetPileItemsFromWindow.length > 0

      // If sheet pile items exist but Sheet pile is not in collectedSubsections, add it
      if (hasSheetPileItems && !collectedSubsections.has('Sheet pile') && !collectedSubsections.has('Sheet piles')) {
        collectedSubsections.add('Sheet pile')
      }

      // Also check calculationData directly for sheet pile
      if (!hasSheetPileItems && calculationData && calculationData.length > 0) {
        let foundSheetPile = false
        for (const row of calculationData) {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('sheet pile') && (bText.endsWith(':') || parseFloat(row[2]) > 0)) {
              foundSheetPile = true
              break
            }
          }
        }
        if (foundSheetPile && !collectedSubsections.has('Sheet pile') && !collectedSubsections.has('Sheet piles')) {
          collectedSubsections.add('Sheet pile')
        }
      }

      // Check if timber lagging is in soeSubsectionItems
      const timberLaggingGroups = soeSubsectionItems.get('Timber lagging') || []
      const hasTimberLaggingItems = timberLaggingGroups.length > 0 && timberLaggingGroups.some(g => g.length > 0)

      // If timber lagging items exist but Timber lagging is not in collectedSubsections, add it
      if (hasTimberLaggingItems && !collectedSubsections.has('Timber lagging')) {
        collectedSubsections.add('Timber lagging')
      }

      // Also check calculationData directly for timber lagging
      if (!hasTimberLaggingItems && calculationData && calculationData.length > 0) {
        let foundTimberLagging = false
        for (const row of calculationData) {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('timber lagging') && (bText.endsWith(':') || parseFloat(row[2]) > 0)) {
              foundTimberLagging = true
              break
            }
          }
        }
        if (foundTimberLagging && !collectedSubsections.has('Timber lagging')) {
          collectedSubsections.add('Timber lagging')
        }
      }

      // Check if timber sheeting is in soeSubsectionItems
      const timberSheetingGroups = soeSubsectionItems.get('Timber sheeting') || []
      const hasTimberSheetingItems = timberSheetingGroups.length > 0 && timberSheetingGroups.some(g => g.length > 0)

      // If timber sheeting items exist but Timber sheeting is not in collectedSubsections, add it
      if (hasTimberSheetingItems && !collectedSubsections.has('Timber sheeting')) {
        collectedSubsections.add('Timber sheeting')
      }

      // Also check calculationData directly for timber sheeting
      if (!hasTimberSheetingItems && calculationData && calculationData.length > 0) {
        let foundTimberSheeting = false
        for (const row of calculationData) {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('timber sheeting') && (bText.endsWith(':') || parseFloat(row[2]) > 0)) {
              foundTimberSheeting = true
              break
            }
          }
        }
        if (foundTimberSheeting && !collectedSubsections.has('Timber sheeting')) {
          collectedSubsections.add('Timber sheeting')
        }
      }

      // Check if vertical timber sheets is in soeSubsectionItems
      const verticalTimberSheetsGroups = soeSubsectionItems.get('Vertical timber sheets') || []
      const hasVerticalTimberSheetsItems = verticalTimberSheetsGroups.length > 0 && verticalTimberSheetsGroups.some(g => g.length > 0)

      // If vertical timber sheets items exist but Vertical timber sheets is not in collectedSubsections, add it
      if (hasVerticalTimberSheetsItems && !collectedSubsections.has('Vertical timber sheets')) {
        collectedSubsections.add('Vertical timber sheets')
      }

      // Also check calculationData directly for vertical timber sheets
      if (!hasVerticalTimberSheetsItems && calculationData && calculationData.length > 0) {
        let foundVerticalTimberSheets = false
        for (const row of calculationData) {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('vertical timber sheets') && (bText.endsWith(':') || parseFloat(row[2]) > 0)) {
              foundVerticalTimberSheets = true
              break
            }
          }
        }
        if (foundVerticalTimberSheets && !collectedSubsections.has('Vertical timber sheets')) {
          collectedSubsections.add('Vertical timber sheets')
        }
      }

      // Check if horizontal timber sheets is in soeSubsectionItems
      const horizontalTimberSheetsGroups = soeSubsectionItems.get('Horizontal timber sheets') || []
      const hasHorizontalTimberSheetsItems = horizontalTimberSheetsGroups.length > 0 && horizontalTimberSheetsGroups.some(g => g.length > 0)

      // If horizontal timber sheets items exist but Horizontal timber sheets is not in collectedSubsections, add it
      if (hasHorizontalTimberSheetsItems && !collectedSubsections.has('Horizontal timber sheets')) {
        collectedSubsections.add('Horizontal timber sheets')
      }

      // Also check calculationData directly for horizontal timber sheets
      if (!hasHorizontalTimberSheetsItems && calculationData && calculationData.length > 0) {
        let foundHorizontalTimberSheets = false
        for (const row of calculationData) {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('horizontal timber sheets') && (bText.endsWith(':') || parseFloat(row[2]) > 0)) {
              foundHorizontalTimberSheets = true
              break
            }
          }
        }
        if (foundHorizontalTimberSheets && !collectedSubsections.has('Horizontal timber sheets')) {
          collectedSubsections.add('Horizontal timber sheets')
        }
      }

      // Check if timber stringer is in soeSubsectionItems
      const timberStringerGroupsFromMap = soeSubsectionItems.get('Timber stringer') || []
      const hasTimberStringerItems = timberStringerGroupsFromMap.length > 0 && timberStringerGroupsFromMap.some(g => g.length > 0)

      // If timber stringer items exist but Timber stringer is not in collectedSubsections, add it
      if (hasTimberStringerItems && !collectedSubsections.has('Timber stringer')) {
        collectedSubsections.add('Timber stringer')
      }

      // Also check calculationData directly for timber stringer
      if (!hasTimberStringerItems && calculationData && calculationData.length > 0) {
        let foundTimberStringer = false
        for (const row of calculationData) {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('timber stringer') && (bText.endsWith(':') || parseFloat(row[2]) > 0)) {
              foundTimberStringer = true
              break
            }
          }
        }
        if (foundTimberStringer && !collectedSubsections.has('Timber stringer')) {
          collectedSubsections.add('Timber stringer')
        }
      }
      
      // Check if heel blocks are in soeSubsectionItems
      const heelBlockGroups = soeSubsectionItems.get('Heel blocks') || []
      const hasHeelBlockItems = heelBlockGroups.length > 0 && heelBlockGroups.some(g => g.length > 0)

      // If heel block items exist but Heel blocks is not in collectedSubsections, add it
      if (hasHeelBlockItems && !collectedSubsections.has('Heel blocks')) {
        collectedSubsections.add('Heel blocks')
      }

      // Also check calculationData directly for heel blocks
      if (!hasHeelBlockItems && calculationData && calculationData.length > 0) {
        let foundHeelBlocks = false
        for (const row of calculationData) {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('heel block') && (bText.endsWith(':') || parseFloat(row[2]) > 0)) {
              foundHeelBlocks = true
              break
            }
          }
        }
        if (foundHeelBlocks && !collectedSubsections.has('Heel blocks')) {
          collectedSubsections.add('Heel blocks')
        }
      }

      // Check if underpinning is in soeSubsectionItems
      const underpinningGroups = soeSubsectionItems.get('Underpinning') || []
      const hasUnderpinningItems = underpinningGroups.length > 0 && underpinningGroups.some(g => g.length > 0)

      // If underpinning items exist but Underpinning is not in collectedSubsections, add it
      if (hasUnderpinningItems && !collectedSubsections.has('Underpinning')) {
        collectedSubsections.add('Underpinning')
      }

      // Also check calculationData directly for underpinning
      if (!hasUnderpinningItems && calculationData && calculationData.length > 0) {
        let foundUnderpinning = false
        for (const row of calculationData) {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('underpinning') && !bText.includes('shim') && (bText.endsWith(':') || parseFloat(row[2]) > 0)) {
              foundUnderpinning = true
              break
            }
          }
        }
        if (foundUnderpinning && !collectedSubsections.has('Underpinning')) {
          collectedSubsections.add('Underpinning')
        }
      }

      // Check if concrete soil retention pier is in soeSubsectionItems
      const concreteSoilRetentionGroups = soeSubsectionItems.get('Concrete soil retention piers') ||
        soeSubsectionItems.get('Concrete soil retention pier') || []
      const hasConcreteSoilRetentionItems = concreteSoilRetentionGroups.length > 0 && concreteSoilRetentionGroups.some(g => g.length > 0)

      // If concrete soil retention pier items exist but not in collectedSubsections, add it
      if (hasConcreteSoilRetentionItems && !collectedSubsections.has('Concrete soil retention piers') && !collectedSubsections.has('Concrete soil retention pier')) {
        collectedSubsections.add('Concrete soil retention pier')
      }

      // Also check calculationData directly for concrete soil retention pier
      if (!hasConcreteSoilRetentionItems && calculationData && calculationData.length > 0) {
        let foundConcreteSoilRetention = false
        for (const row of calculationData) {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('concrete soil retention pier') && (bText.endsWith(':') || parseFloat(row[2]) > 0)) {
              foundConcreteSoilRetention = true
              break
            }
          }
        }
        if (foundConcreteSoilRetention && !collectedSubsections.has('Concrete soil retention piers') && !collectedSubsections.has('Concrete soil retention pier')) {
          collectedSubsections.add('Concrete soil retention pier')
        }
      }

      // Check if form board is in soeSubsectionItems
      const formBoardGroups = soeSubsectionItems.get('Form board') || []
      const hasFormBoardItems = formBoardGroups.length > 0 && formBoardGroups.some(g => g.length > 0)

      // If form board items exist but not in collectedSubsections, add it
      if (hasFormBoardItems && !collectedSubsections.has('Form board')) {
        collectedSubsections.add('Form board')
      }

      // Check if drilled hole grout is in soeSubsectionItems
      const drilledHoleGroutGroups = soeSubsectionItems.get('Drilled hole grout') || []
      const hasDrilledHoleGroutItems = drilledHoleGroutGroups.length > 0 && drilledHoleGroutGroups.some(g => g.length > 0)

      // If drilled hole grout items exist but not in collectedSubsections, add it
      if (hasDrilledHoleGroutItems && !collectedSubsections.has('Drilled hole grout')) {
        collectedSubsections.add('Drilled hole grout')
      }

      // Also check calculationData directly for form board
      if (!hasFormBoardItems && calculationData && calculationData.length > 0) {
        let foundFormBoard = false
        for (const row of calculationData) {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('form board') && (bText.endsWith(':') || parseFloat(row[2]) > 0)) {
              foundFormBoard = true
              break
            }
          }
        }
        if (foundFormBoard && !collectedSubsections.has('Form board')) {
          collectedSubsections.add('Form board')
        }
      }

      // Check if buttons/concrete buttons are in soeSubsectionItems
      const buttonGroups = soeSubsectionItems.get('Buttons') || soeSubsectionItems.get('Concrete buttons') || []
      const hasButtonItems = buttonGroups.length > 0 && buttonGroups.some(g => g.length > 0)

      // If button items exist, add both "Buttons" and "Concrete buttons" to collectedSubsections
      if (hasButtonItems) {
        if (!collectedSubsections.has('Buttons')) {
          collectedSubsections.add('Buttons')
        }
        if (!collectedSubsections.has('Concrete buttons')) {
          collectedSubsections.add('Concrete buttons')
        }
      }

      // Also check calculationData directly for buttons
      if (!hasButtonItems && calculationData && calculationData.length > 0) {
        let foundButtons = false
        for (const row of calculationData) {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('button') && (bText.endsWith(':') || parseFloat(row[2]) > 0)) {
              foundButtons = true
              break
            }
          }
        }
        if (foundButtons) {
          if (!collectedSubsections.has('Buttons')) {
            collectedSubsections.add('Buttons')
          }
          if (!collectedSubsections.has('Concrete buttons')) {
            collectedSubsections.add('Concrete buttons')
          }
        }
      }

      // Check if any bracing-related subsections exist
      // Match the exact names from the template, plus common variations
      const bracingSubsections = ['Waler', 'Raker', 'Upper Raker', 'Lower Raker', 'Stand off', 'Kicker', 'Channel', 'Roll chock', 'Stub beam', 'Stud beam', 'Inner corner brace', 'Knee brace', 'Supporting angle']
      // Also check case-insensitive matches
      let hasBracingItems = bracingSubsections.some(name => {
        return collectedSubsections.has(name) ||
          Array.from(collectedSubsections).some(collected =>
            collected.toLowerCase() === name.toLowerCase()
          )
      })

      // If not found in collectedSubsections, check soeSubsectionItems directly
      if (!hasBracingItems) {
        hasBracingItems = bracingSubsections.some(name => {
          // Check exact match
          if (soeSubsectionItems.has(name)) {
            const groups = soeSubsectionItems.get(name) || []
            if (groups.length > 0 && groups.some(g => g.length > 0)) return true
          }
          // Check case-insensitive match
          for (const [key, value] of soeSubsectionItems.entries()) {
            if (key.toLowerCase() === name.toLowerCase()) {
              const groups = value || []
              if (groups.length > 0 && groups.some(g => g.length > 0)) return true
            }
          }
          return false
        })
      }

      // If still not found, search calculation data directly
      if (!hasBracingItems && calculationData && calculationData.length > 0) {
        for (const row of calculationData) {
          const colB = row[1] // Column B
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bracingSubsections.some(name => bText.includes(name.toLowerCase()))) {
              hasBracingItems = true
              break
            }
          }
        }
      }

      // Debug: log collected subsections to see what we have

      // Display subsections in order, prioritizing the order list
      const subsectionsToDisplay = []
      subsectionOrder.forEach(name => {
        // Skip excluded subsections
        if (excludedSubsections.has(name)) {
          return
        }

        // Special handling for Bracing - show it if any bracing items exist
        if (name === 'Bracing' && hasBracingItems) {
          subsectionsToDisplay.push(name)
          // Don't remove bracing subsections from collectedSubsections yet
        } else if (name === 'Parging' && hasPargingItems) {
          // Special handling for Parging - show it if any parging items exist
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (name === 'Heel blocks' && hasHeelBlockItems) {
          // Special handling for Heel blocks - show it if any heel block items exist
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (name === 'Underpinning' && hasUnderpinningItems) {
          // Special handling for Underpinning - show it if any underpinning items exist
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (name === 'Concrete soil retention pier' && hasConcreteSoilRetentionItems) {
          // Special handling for Concrete soil retention pier - show it if any items exist
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
          collectedSubsections.delete('Concrete soil retention piers')
        } else if (name === 'Form board' && hasFormBoardItems) {
          // Special handling for Form board - show it if any items exist
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (name === 'Drilled hole grout' && hasDrilledHoleGroutItems) {
          // Special handling for Drilled hole grout - show it if any items exist
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if ((name === 'Guide wall' || name === 'Guilde wall') && hasGuideWallItems) {
          // Special handling for Guide wall - show it if any items exist
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
          collectedSubsections.delete('Guilde wall') // Also remove typo version
        } else if ((name === 'Dowels' || name === 'Dowel bar') && hasDowelBarItems) {
          // Special handling for Dowels - show it if any items exist
          subsectionsToDisplay.push(name === 'Dowel bar' ? 'Dowels' : name)
          collectedSubsections.delete(name)
          collectedSubsections.delete('Dowel bar') // Also remove "Dowel bar" version
          collectedSubsections.delete('Dowels')
        } else if ((name === 'Rock pins' || name === 'Rock pin') && hasRockPinItems) {
          // Special handling for Rock pins - show it if any items exist
          subsectionsToDisplay.push(name === 'Rock pin' ? 'Rock pins' : name)
          collectedSubsections.delete(name)
          collectedSubsections.delete('Rock pin') // Also remove "Rock pin" version
          collectedSubsections.delete('Rock pins')
        } else if (name === 'Rock stabilization' && hasRockStabilizationItems) {
          // Special handling for Rock stabilization - show it if any items exist
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (name === 'Shotcrete' && hasShotcreteItems) {
          // Special handling for Shotcrete - show it if any items exist
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if ((name === 'Sheet pile' || name === 'Sheet piles') && hasSheetPileItems) {
          // Special handling for Sheet pile - show it if any items exist
          subsectionsToDisplay.push('Sheet pile')
          collectedSubsections.delete(name)
          collectedSubsections.delete('Sheet piles') // Also remove plural version
        } else if (name === 'Timber lagging' && hasTimberLaggingItems) {
          // Special handling for Timber lagging - show it if any items exist
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (name === 'Timber sheeting' && hasTimberSheetingItems) {
          // Special handling for Timber sheeting - show it if any items exist
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (name === 'Vertical timber sheets' && hasVerticalTimberSheetsItems) {
          // Special handling for Vertical timber sheets - show it if any items exist
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (name === 'Horizontal timber sheets' && hasHorizontalTimberSheetsItems) {
          // Special handling for Horizontal timber sheets - show it if any items exist
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (name === 'Timber stringer' && hasTimberStringerItems) {
          // Special handling for Timber stringer - show it if any items exist
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (collectedSubsections.has(name)) {
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        }
      })

      // If bracing items exist but Bracing wasn't added yet, add it now
      if (hasBracingItems && !subsectionsToDisplay.includes('Bracing')) {
        // Find the right position to insert Bracing (after Tangent pile, before Tie back)
        const bracingIndex = subsectionOrder.indexOf('Bracing')
        if (bracingIndex !== -1) {
          // Find where to insert it based on subsectionOrder
          let insertIndex = subsectionsToDisplay.length
          for (let i = 0; i < subsectionsToDisplay.length; i++) {
            const currentIndex = subsectionOrder.indexOf(subsectionsToDisplay[i])
            if (currentIndex !== -1 && currentIndex > bracingIndex) {
              insertIndex = i
              break
            }
          }
          subsectionsToDisplay.splice(insertIndex, 0, 'Bracing')
        } else {
          // If not in order, add it after Tangent pile or at a reasonable position
          subsectionsToDisplay.push('Bracing')
        }
      }

      // If parging items exist but Parging wasn't added yet, add it now
      if (hasPargingItems && !subsectionsToDisplay.includes('Parging')) {
        // Find the right position to insert Parging (after Tie down anchor, before Heel blocks)
        const pargingIndex = subsectionOrder.indexOf('Parging')
        if (pargingIndex !== -1) {
          // Find where to insert it based on subsectionOrder
          let insertIndex = subsectionsToDisplay.length
          for (let i = 0; i < subsectionsToDisplay.length; i++) {
            const currentIndex = subsectionOrder.indexOf(subsectionsToDisplay[i])
            if (currentIndex !== -1 && currentIndex > pargingIndex) {
              insertIndex = i
              break
            }
          }
          subsectionsToDisplay.splice(insertIndex, 0, 'Parging')
        } else {
          // If not in order, add it at a reasonable position
          subsectionsToDisplay.push('Parging')
        }
      }

      // If heel block items exist but Heel blocks wasn't added yet, add it now (right after Parging)
      if (hasHeelBlockItems && !subsectionsToDisplay.includes('Heel blocks')) {
        // Find the right position to insert Heel blocks (after Parging, before Underpinning)
        const heelBlockIndex = subsectionOrder.indexOf('Heel blocks')
        if (heelBlockIndex !== -1) {
          // Find where to insert it based on subsectionOrder
          let insertIndex = subsectionsToDisplay.length
          for (let i = 0; i < subsectionsToDisplay.length; i++) {
            const currentIndex = subsectionOrder.indexOf(subsectionsToDisplay[i])
            if (currentIndex !== -1 && currentIndex > heelBlockIndex) {
              insertIndex = i
              break
            }
          }
          subsectionsToDisplay.splice(insertIndex, 0, 'Heel blocks')
        } else {
          // If not in order, try to insert after Parging if it exists
          const pargingPos = subsectionsToDisplay.indexOf('Parging')
          if (pargingPos !== -1) {
            subsectionsToDisplay.splice(pargingPos + 1, 0, 'Heel blocks')
          } else {
            subsectionsToDisplay.push('Heel blocks')
          }
        }
      }

      // If underpinning items exist but Underpinning wasn't added yet, add it now (right after Heel blocks)
      if (hasUnderpinningItems && !subsectionsToDisplay.includes('Underpinning')) {
        // Find the right position to insert Underpinning (after Heel blocks)
        const underpinningIndex = subsectionOrder.indexOf('Underpinning')
        if (underpinningIndex !== -1) {
          // Find where to insert it based on subsectionOrder
          let insertIndex = subsectionsToDisplay.length
          for (let i = 0; i < subsectionsToDisplay.length; i++) {
            const currentIndex = subsectionOrder.indexOf(subsectionsToDisplay[i])
            if (currentIndex !== -1 && currentIndex > underpinningIndex) {
              insertIndex = i
              break
            }
          }
          subsectionsToDisplay.splice(insertIndex, 0, 'Underpinning')
        } else {
          // If not in order, try to insert after Heel blocks if it exists
          const heelBlockPos = subsectionsToDisplay.indexOf('Heel blocks')
          if (heelBlockPos !== -1) {
            subsectionsToDisplay.splice(heelBlockPos + 1, 0, 'Underpinning')
          } else {
            subsectionsToDisplay.push('Underpinning')
          }
        }
      }

      // If concrete soil retention pier items exist but Concrete soil retention pier wasn't added yet, add it now (right after Underpinning)
      if (hasConcreteSoilRetentionItems && !subsectionsToDisplay.includes('Concrete soil retention pier')) {
        // Find the right position to insert Concrete soil retention pier (after Underpinning)
        const concreteSoilRetentionIndex = subsectionOrder.indexOf('Concrete soil retention pier')
        if (concreteSoilRetentionIndex !== -1) {
          // Find where to insert it based on subsectionOrder
          let insertIndex = subsectionsToDisplay.length
          for (let i = 0; i < subsectionsToDisplay.length; i++) {
            const currentIndex = subsectionOrder.indexOf(subsectionsToDisplay[i])
            if (currentIndex !== -1 && currentIndex > concreteSoilRetentionIndex) {
              insertIndex = i
              break
            }
          }
          subsectionsToDisplay.splice(insertIndex, 0, 'Concrete soil retention pier')
        } else {
          // If not in order, try to insert after Underpinning if it exists
          const underpinningPos = subsectionsToDisplay.indexOf('Underpinning')
          if (underpinningPos !== -1) {
            subsectionsToDisplay.splice(underpinningPos + 1, 0, 'Concrete soil retention pier')
          } else {
            subsectionsToDisplay.push('Concrete soil retention pier')
          }
        }
      }

      // If guide wall items exist but Guide wall wasn't added yet, add it now
      if (hasGuideWallItems && !subsectionsToDisplay.includes('Guide wall') && !subsectionsToDisplay.includes('Guilde wall')) {
        // Find the right position to insert Guide wall (after Concrete soil retention pier, before Concrete buttons)
        const guideWallIndex = subsectionOrder.indexOf('Guide wall')
        if (guideWallIndex !== -1) {
          // Find where to insert it based on subsectionOrder
          let insertIndex = subsectionsToDisplay.length
          for (let i = 0; i < subsectionsToDisplay.length; i++) {
            const currentIndex = subsectionOrder.indexOf(subsectionsToDisplay[i])
            if (currentIndex !== -1 && currentIndex > guideWallIndex) {
              insertIndex = i
              break
            }
          }
          subsectionsToDisplay.splice(insertIndex, 0, 'Guide wall')
        } else {
          // If not in order, try to insert after Concrete soil retention pier if it exists
          const concreteSoilRetentionPos = subsectionsToDisplay.indexOf('Concrete soil retention pier')
          if (concreteSoilRetentionPos !== -1) {
            subsectionsToDisplay.splice(concreteSoilRetentionPos + 1, 0, 'Guide wall')
          } else {
            subsectionsToDisplay.push('Guide wall')
          }
        }
      }

      // If form board items exist but Form board wasn't added yet, add it now (right after Concrete soil retention pier)
      if (hasFormBoardItems && !subsectionsToDisplay.includes('Form board')) {
        // Find the right position to insert Form board (after Concrete soil retention pier)
        const formBoardIndex = subsectionOrder.indexOf('Form board')
        if (formBoardIndex !== -1) {
          // Find where to insert it based on subsectionOrder
          let insertIndex = subsectionsToDisplay.length
          for (let i = 0; i < subsectionsToDisplay.length; i++) {
            const currentIndex = subsectionOrder.indexOf(subsectionsToDisplay[i])
            if (currentIndex !== -1 && currentIndex > formBoardIndex) {
              insertIndex = i
              break
            }
          }
          subsectionsToDisplay.splice(insertIndex, 0, 'Form board')
        } else {
          // If not in order, try to insert after Concrete soil retention pier if it exists
          const concreteSoilRetentionPos = subsectionsToDisplay.indexOf('Concrete soil retention pier')
          if (concreteSoilRetentionPos !== -1) {
            subsectionsToDisplay.splice(concreteSoilRetentionPos + 1, 0, 'Form board')
          } else {
            subsectionsToDisplay.push('Form board')
          }
        }
      }

      // If drilled hole grout items exist but Drilled hole grout wasn't added yet, add it now (right after Form board)
      if (hasDrilledHoleGroutItems && !subsectionsToDisplay.includes('Drilled hole grout')) {
        const drilledHoleGroutIndex = subsectionOrder.indexOf('Drilled hole grout')
        if (drilledHoleGroutIndex !== -1) {
          let insertIndex = subsectionsToDisplay.length
          for (let i = 0; i < subsectionsToDisplay.length; i++) {
            const currentIndex = subsectionOrder.indexOf(subsectionsToDisplay[i])
            if (currentIndex !== -1 && currentIndex > drilledHoleGroutIndex) {
              insertIndex = i
              break
            }
          }
          subsectionsToDisplay.splice(insertIndex, 0, 'Drilled hole grout')
        } else {
          const formBoardPos = subsectionsToDisplay.indexOf('Form board')
          if (formBoardPos !== -1) {
            subsectionsToDisplay.splice(formBoardPos + 1, 0, 'Drilled hole grout')
          } else {
            subsectionsToDisplay.push('Drilled hole grout')
          }
        }
      }

      // If rock stabilization items exist but Rock stabilization wasn't added yet, add it now
      if (hasRockStabilizationItems && !subsectionsToDisplay.includes('Rock stabilization')) {
        // Find the right position to insert Rock stabilization (after Rock pins, before Shotcrete)
        const rockStabilizationIndex = subsectionOrder.indexOf('Rock stabilization')
        if (rockStabilizationIndex !== -1) {
          // Find where to insert it based on subsectionOrder
          let insertIndex = subsectionsToDisplay.length
          for (let i = 0; i < subsectionsToDisplay.length; i++) {
            const currentIndex = subsectionOrder.indexOf(subsectionsToDisplay[i])
            if (currentIndex !== -1 && currentIndex > rockStabilizationIndex) {
              insertIndex = i
              break
            }
          }
          subsectionsToDisplay.splice(insertIndex, 0, 'Rock stabilization')
        } else {
          // If not in order, try to insert after Rock pins if it exists
          const rockPinsPos = subsectionsToDisplay.indexOf('Rock pins')
          if (rockPinsPos !== -1) {
            subsectionsToDisplay.splice(rockPinsPos + 1, 0, 'Rock stabilization')
          } else {
            subsectionsToDisplay.push('Rock stabilization')
          }
        }
      }

      // If shotcrete items exist but Shotcrete wasn't added yet, add it now
      if (hasShotcreteItems && !subsectionsToDisplay.includes('Shotcrete')) {
        // Find the right position to insert Shotcrete (after Rock stabilization, before Permission grouting)
        const shotcreteIndex = subsectionOrder.indexOf('Shotcrete')
        if (shotcreteIndex !== -1) {
          // Find where to insert it based on subsectionOrder
          let insertIndex = subsectionsToDisplay.length
          for (let i = 0; i < subsectionsToDisplay.length; i++) {
            const currentIndex = subsectionOrder.indexOf(subsectionsToDisplay[i])
            if (currentIndex !== -1 && currentIndex > shotcreteIndex) {
              insertIndex = i
              break
            }
          }
          subsectionsToDisplay.splice(insertIndex, 0, 'Shotcrete')
        } else {
          // If not in order, try to insert after Rock stabilization if it exists
          const rockStabilizationPos = subsectionsToDisplay.indexOf('Rock stabilization')
          if (rockStabilizationPos !== -1) {
            subsectionsToDisplay.splice(rockStabilizationPos + 1, 0, 'Shotcrete')
          } else {
            subsectionsToDisplay.push('Shotcrete')
          }
        }
      }

      // If sheet pile items exist but Sheet pile wasn't added yet, add it now (right after Tangent pile)
      if (hasSheetPileItems && !subsectionsToDisplay.includes('Sheet pile')) {
        // Find the right position to insert Sheet pile (after Tangent pile, before Bracing)
        const sheetPileIndex = subsectionOrder.indexOf('Sheet pile')
        if (sheetPileIndex !== -1) {
          // Find where to insert it based on subsectionOrder
          let insertIndex = subsectionsToDisplay.length
          for (let i = 0; i < subsectionsToDisplay.length; i++) {
            const currentIndex = subsectionOrder.indexOf(subsectionsToDisplay[i])
            if (currentIndex !== -1 && currentIndex > sheetPileIndex) {
              insertIndex = i
              break
            }
          }
          subsectionsToDisplay.splice(insertIndex, 0, 'Sheet pile')
        } else {
          // If not in order, try to insert after Tangent pile if it exists
          const tangentPilePos = subsectionsToDisplay.indexOf('Tangent pile')
          if (tangentPilePos !== -1) {
            subsectionsToDisplay.splice(tangentPilePos + 1, 0, 'Sheet pile')
          } else {
            // Try after Tangent piles (plural)
            const tangentPilesPos = subsectionsToDisplay.indexOf('Tangent piles')
            if (tangentPilesPos !== -1) {
              subsectionsToDisplay.splice(tangentPilesPos + 1, 0, 'Sheet pile')
            } else {
              subsectionsToDisplay.push('Sheet pile')
            }
          }
        }
      }

      // If timber lagging items exist but Timber lagging wasn't added yet, add it now (right after Sheet pile)
      if (hasTimberLaggingItems && !subsectionsToDisplay.includes('Timber lagging')) {
        // Find the right position to insert Timber lagging (after Sheet pile, before Timber sheeting)
        const timberLaggingIndex = subsectionOrder.indexOf('Timber lagging')
        if (timberLaggingIndex !== -1) {
          // Find where to insert it based on subsectionOrder
          let insertIndex = subsectionsToDisplay.length
          for (let i = 0; i < subsectionsToDisplay.length; i++) {
            const currentIndex = subsectionOrder.indexOf(subsectionsToDisplay[i])
            if (currentIndex !== -1 && currentIndex > timberLaggingIndex) {
              insertIndex = i
              break
            }
          }
          subsectionsToDisplay.splice(insertIndex, 0, 'Timber lagging')
        } else {
          // If not in order, try to insert after Sheet pile if it exists
          const sheetPilePos = subsectionsToDisplay.indexOf('Sheet pile')
          if (sheetPilePos !== -1) {
            subsectionsToDisplay.splice(sheetPilePos + 1, 0, 'Timber lagging')
          } else {
            subsectionsToDisplay.push('Timber lagging')
          }
        }
      }

      // If timber sheeting items exist but Timber sheeting wasn't added yet, add it now (right after Timber lagging)
      if (hasTimberSheetingItems && !subsectionsToDisplay.includes('Timber sheeting')) {
        // Find the right position to insert Timber sheeting (after Timber lagging, before Bracing)
        const timberSheetingIndex = subsectionOrder.indexOf('Timber sheeting')
        if (timberSheetingIndex !== -1) {
          // Find where to insert it based on subsectionOrder
          let insertIndex = subsectionsToDisplay.length
          for (let i = 0; i < subsectionsToDisplay.length; i++) {
            const currentIndex = subsectionOrder.indexOf(subsectionsToDisplay[i])
            if (currentIndex !== -1 && currentIndex > timberSheetingIndex) {
              insertIndex = i
              break
            }
          }
          subsectionsToDisplay.splice(insertIndex, 0, 'Timber sheeting')
        } else {
          // If not in order, try to insert after Timber lagging if it exists
          const timberLaggingPos = subsectionsToDisplay.indexOf('Timber lagging')
          if (timberLaggingPos !== -1) {
            subsectionsToDisplay.splice(timberLaggingPos + 1, 0, 'Timber sheeting')
          } else {
            // Try after Sheet pile if Timber lagging not found
            const sheetPilePos = subsectionsToDisplay.indexOf('Sheet pile')
            if (sheetPilePos !== -1) {
              subsectionsToDisplay.splice(sheetPilePos + 1, 0, 'Timber sheeting')
            } else {
              subsectionsToDisplay.push('Timber sheeting')
            }
          }
        }
      }

      // If vertical timber sheets items exist but Vertical timber sheets wasn't added yet, add it now (right after Timber sheeting)
      if (hasVerticalTimberSheetsItems && !subsectionsToDisplay.includes('Vertical timber sheets')) {
        const verticalTimberSheetsIndex = subsectionOrder.indexOf('Vertical timber sheets')
        if (verticalTimberSheetsIndex !== -1) {
          let insertIndex = subsectionsToDisplay.length
          for (let i = 0; i < subsectionsToDisplay.length; i++) {
            const currentIndex = subsectionOrder.indexOf(subsectionsToDisplay[i])
            if (currentIndex !== -1 && currentIndex > verticalTimberSheetsIndex) {
              insertIndex = i
              break
            }
          }
          subsectionsToDisplay.splice(insertIndex, 0, 'Vertical timber sheets')
        } else {
          const timberSheetingPos = subsectionsToDisplay.indexOf('Timber sheeting')
          if (timberSheetingPos !== -1) {
            subsectionsToDisplay.splice(timberSheetingPos + 1, 0, 'Vertical timber sheets')
          } else {
            const timberLaggingPos = subsectionsToDisplay.indexOf('Timber lagging')
            if (timberLaggingPos !== -1) {
              subsectionsToDisplay.splice(timberLaggingPos + 1, 0, 'Vertical timber sheets')
            } else {
              subsectionsToDisplay.push('Vertical timber sheets')
            }
          }
        }
      }

      // If horizontal timber sheets items exist but Horizontal timber sheets wasn't added yet, add it now (right after Vertical timber sheets)
      if (hasHorizontalTimberSheetsItems && !subsectionsToDisplay.includes('Horizontal timber sheets')) {
        const horizontalTimberSheetsIndex = subsectionOrder.indexOf('Horizontal timber sheets')
        if (horizontalTimberSheetsIndex !== -1) {
          let insertIndex = subsectionsToDisplay.length
          for (let i = 0; i < subsectionsToDisplay.length; i++) {
            const currentIndex = subsectionOrder.indexOf(subsectionsToDisplay[i])
            if (currentIndex !== -1 && currentIndex > horizontalTimberSheetsIndex) {
              insertIndex = i
              break
            }
          }
          subsectionsToDisplay.splice(insertIndex, 0, 'Horizontal timber sheets')
        } else {
          const verticalTimberSheetsPos = subsectionsToDisplay.indexOf('Vertical timber sheets')
          if (verticalTimberSheetsPos !== -1) {
            subsectionsToDisplay.splice(verticalTimberSheetsPos + 1, 0, 'Horizontal timber sheets')
          } else {
            const timberSheetingPos = subsectionsToDisplay.indexOf('Timber sheeting')
            if (timberSheetingPos !== -1) {
              subsectionsToDisplay.splice(timberSheetingPos + 1, 0, 'Horizontal timber sheets')
            } else {
              subsectionsToDisplay.push('Horizontal timber sheets')
            }
          }
        }
      }

      // If timber stringer items exist but Timber stringer wasn't added yet, add it now (right after Horizontal timber sheets)
      if (hasTimberStringerItems && !subsectionsToDisplay.includes('Timber stringer')) {
        const timberStringerIndex = subsectionOrder.indexOf('Timber stringer')
        if (timberStringerIndex !== -1) {
          let insertIndex = subsectionsToDisplay.length
          for (let i = 0; i < subsectionsToDisplay.length; i++) {
            const currentIndex = subsectionOrder.indexOf(subsectionsToDisplay[i])
            if (currentIndex !== -1 && currentIndex > timberStringerIndex) {
              insertIndex = i
              break
            }
          }
          subsectionsToDisplay.splice(insertIndex, 0, 'Timber stringer')
        } else {
          const horizontalTimberSheetsPos = subsectionsToDisplay.indexOf('Horizontal timber sheets')
          if (horizontalTimberSheetsPos !== -1) {
            subsectionsToDisplay.splice(horizontalTimberSheetsPos + 1, 0, 'Timber stringer')
          } else {
            subsectionsToDisplay.push('Timber stringer')
          }
        }
      }

      // Add any remaining subsections (but skip bracing items if Bracing header was added, and skip excluded ones)
      if (hasBracingItems && subsectionsToDisplay.includes('Bracing')) {
        // Remove bracing subsections from collectedSubsections so they don't appear twice
        bracingSubsections.forEach(name => collectedSubsections.delete(name))
      }
      collectedSubsections.forEach(name => {
        // Skip excluded subsections
        if (!excludedSubsections.has(name)) {
          subsectionsToDisplay.push(name)
        }
      })

      // Search for Timber lagging and Timber sheeting in calculation data (always run)
      if (calculationData && calculationData.length > 0) {
        const timberLaggingRows = []
        const timberSheetingRows = []
        calculationData.forEach((row, index) => {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('timber lagging')) {
              timberLaggingRows.push({
                rowIndex: index + 2,
                row: row,
                colB: colB
              })
            }
            if (bText.includes('timber sheeting')) {
              timberSheetingRows.push({
                rowIndex: index + 2,
                row: row,
                colB: colB
              })
            }
          }
        })
      }

      subsectionsToDisplay.forEach((subsectionName) => {

        // Special handling for Bracing - show it as a heading and then process all bracing items below
        if (subsectionName === 'Bracing' && hasBracingItems) {
          // Add Bracing header
          const subsectionText = `Bracing: including shims, stiffeners, and/or wale seats as required`
          spreadsheet.updateCell({ value: subsectionText }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: '#D0CECE',
              textDecoration: 'underline',
              border: '1px solid #000000'
            },
            `${pfx}B${currentRow}`
          )
          spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
          currentRow++

          // Process all bracing-related subsections and display their items
          // Map subsection names to display names (matching template names)
          const bracingDisplayNames = {
            'Waler': 'Waler',
            'Raker': 'Raker',
            'Upper Raker': 'Upper Raker',
            'Lower Raker': 'Lower Raker',
            'Stand off': 'Stand Off',
            'Kicker': 'Kicker',
            'Channel': 'Channel',
            'Roll chock': 'Roll Chock',
            'Stub beam': 'Stub Beam',
            'Stud beam': 'Stub Beam',
            'Inner corner brace': 'Inner Corner Brace',
            'Knee brace': 'Knee Brace',
            'Supporting angle': 'Supporting Angle'
          }

          // Helper function to find bracing groups with case-insensitive matching
          const getBracingGroups = (subsectionName) => {
            // Try exact match first
            let groups = soeSubsectionItems.get(subsectionName) || []
            if (groups.length > 0 && groups.some(g => g.length > 0)) return groups

            // Try case-insensitive match
            for (const [key, value] of soeSubsectionItems.entries()) {
              if (key.toLowerCase() === subsectionName.toLowerCase()) {
                const groups = value || []
                if (groups.length > 0 && groups.some(g => g.length > 0)) return groups
              }
            }
            return []
          }

          // Helper function to find bracing items from calculation data
          const findBracingItemsFromCalculationData = (subsectionName) => {
            if (!calculationData || calculationData.length === 0) return []

            const foundRows = []
            let inSubsection = false

            calculationData.forEach((row, index) => {
              const rowNum = index + 2
              const colB = row[1]

              if (colB && typeof colB === 'string') {
                const bText = colB.trim()
                // Check if this is the subsection header
                if ((bText.toLowerCase().includes(subsectionName.toLowerCase()) ||
                  subsectionName.toLowerCase().includes(bText.toLowerCase())) &&
                  bText.endsWith(':')) {
                  inSubsection = true
                  return
                }

                // If we're in the subsection and hit another subsection or section, stop
                if (inSubsection) {
                  if (bText.endsWith(':') && !bText.toLowerCase().includes(subsectionName.toLowerCase())) {
                    inSubsection = false
                    return
                  }

                  // This is a data row in the subsection
                  if (bText && !bText.endsWith(':') && row[2] !== undefined) {
                    const takeoff = parseFloat(row[2]) || 0
                    if (takeoff > 0) {
                      foundRows.push({
                        rowNum: rowNum,
                        particulars: bText,
                        takeoff: takeoff,
                        qty: parseFloat(row[12]) || 0,
                        height: parseFloat(row[7]) || 0,
                        lbs: parseFloat(row[10]) || 0,
                        rawRowNumber: rowNum
                      })
                    }
                  }
                }
              }
            })

            return foundRows.length > 0 ? [foundRows] : []
          }

          bracingSubsections.forEach(bracingSubsectionName => {
            let bracingGroups = getBracingGroups(bracingSubsectionName)

            // If not found in soeSubsectionItems, try calculation data
            if (bracingGroups.length === 0 || !bracingGroups.some(g => g.length > 0)) {
              bracingGroups = findBracingItemsFromCalculationData(bracingSubsectionName)
            }

            if (bracingGroups.length === 0) return // Skip if no groups found

            // Special handling for Supporting angle - group by size and location
            if (bracingSubsectionName.toLowerCase() === 'supporting angle') {
              // Collect all supporting angle items
              const allSupportingAngleItems = []
              bracingGroups.forEach(group => {
                group.forEach(item => {
                  if (item.particulars) {
                    allSupportingAngleItems.push(item)
                  }
                })
              })

              // Group by size and location
              const groupedBySizeAndLocation = new Map()
              allSupportingAngleItems.forEach(item => {
                const particulars = item.particulars || ''
                // Extract size (e.g., L4x4x½, L8x4x½)
                const sizeMatch = particulars.match(/(L\d+x\d+x[½¼¾\d\/]+)/i)
                const size = sizeMatch ? sizeMatch[1] : ''

                // Extract location (@ timber lagging, @ soldier pile, etc.)
                let location = ''
                if (particulars.toLowerCase().includes('@ timber lagging')) {
                  location = '@ timber lagging'
                } else if (particulars.toLowerCase().includes('@ soldier pile')) {
                  location = '@ soldier piles'
                } else if (particulars.toLowerCase().includes('@')) {
                  const locationMatch = particulars.match(/@\s*([^,]+)/i)
                  if (locationMatch) {
                    location = `@ ${locationMatch[1].trim()}`
                  }
                }

                const key = `${size}|${location}`
                if (!groupedBySizeAndLocation.has(key)) {
                  groupedBySizeAndLocation.set(key, [])
                }
                groupedBySizeAndLocation.get(key).push(item)
              })

              // Process each group separately
              groupedBySizeAndLocation.forEach((group, key) => {
                if (group.length === 0) return

                const [size, location] = key.split('|')

                // Calculate totals for this group
                let totalTakeoff = 0
                let totalQty = 0
                let totalFT = 0
                let totalLBS = 0
                let lastRowNumber = 0

                group.forEach(item => {
                  totalTakeoff += item.takeoff || 0
                  totalQty += item.qty || 0
                  const itemFT = (item.height || 0) * (item.takeoff || 0)
                  totalFT += itemFT
                  totalLBS += item.lbs || 0
                  lastRowNumber = Math.max(lastRowNumber, item.rawRowNumber || 0)
                })

                // Find the sum row for this group
                const sumRowIndex = lastRowNumber
                const calcSheetName = 'Calculations Sheet'

                // Get SOE page references from raw data
                let soePageMain = 'SOE-101.00'
                let soePageDetails = 'SOE-301.00'
                if (rawData && Array.isArray(rawData) && rawData.length > 1) {
                  const headers = rawData[0]
                  const dataRows = rawData.slice(1)
                  const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
                  const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

                  if (digitizerIdx !== -1 && pageIdx !== -1) {
                    for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                      const row = dataRows[rowIndex]
                      const digitizerItem = row[digitizerIdx]
                      if (digitizerItem && typeof digitizerItem === 'string') {
                        const itemText = digitizerItem.toLowerCase()
                        if (itemText.includes('supporting angle')) {
                          const pageValue = row[pageIdx]
                          if (pageValue) {
                            const pageStr = String(pageValue).trim()
                            const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                            if (soeMatches && soeMatches.length > 0) {
                              soePageMain = soeMatches[0]
                              if (soeMatches.length > 1) {
                                soePageDetails = soeMatches[1]
                              }
                            }
                          }
                        }
                      }
                    }
                  }
                }

                // Generate proposal text: F&I new (X)no [size] supporting angle @ [location] as per SOE-XXX.XX & details on SOE-XXX.XX
                const qtyValue = Math.round(totalQty || totalTakeoff)
                const proposalText = `F&I new (${qtyValue})no ${size} supporting angle ${location} as per ${soePageMain} & details on ${soePageDetails}`

                // Add proposal text
                spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
                spreadsheet.wrap(`${pfx}B${currentRow}`, true)
                spreadsheet.cellFormat(
                  {
                    fontWeight: 'bold',
                    color: '#000000',
                    textAlign: 'left',
                    backgroundColor: 'white',
                    verticalAlign: 'top'
                  },
                  `${pfx}B${currentRow}`
                )

                // Calculate and set row height based on content
                const dynamicHeight = calculateRowHeight(proposalText)
                spreadsheet.setRowHeight(dynamicHeight, currentRow, proposalSheetIndex)

                // Add FT (LF) to column C - reference to calculation sheet sum row
                if (sumRowIndex > 0) {
                  spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${sumRowIndex}` }, `${pfx}C${currentRow}`)
                  spreadsheet.cellFormat(
                    {
                      fontWeight: 'bold',
                      textAlign: 'right',
                      backgroundColor: 'white'
                    },
                    `${pfx}C${currentRow}`
                  )
                }

                // Add LBS to column E - reference to calculation sheet sum row
                if (sumRowIndex > 0) {
                  spreadsheet.updateCell({ formula: `='${calcSheetName}'!K${sumRowIndex}` }, `${pfx}E${currentRow}`)
                  spreadsheet.cellFormat(
                    {
                      fontWeight: 'bold',
                      textAlign: 'right',
                      backgroundColor: 'white'
                    },
                    `${pfx}E${currentRow}`
                  )
                }

                // Add QTY to column G - reference to calculation sheet sum row
                if (sumRowIndex > 0) {
                  spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
                  spreadsheet.cellFormat(
                    {
                      fontWeight: 'bold',
                      textAlign: 'right',
                      backgroundColor: 'white'
                    },
                    `${pfx}G${currentRow}`
                  )
                }

                // Add 180 to second LF column (column I) - same as other SOE items
                spreadsheet.updateCell({ value: 180 }, `${pfx}I${currentRow}`)
                spreadsheet.cellFormat(
                  {
                    fontWeight: 'bold',
                    textAlign: 'right',
                    backgroundColor: '#E2EFDA'
                  },
                  `${pfx}I${currentRow}:N${currentRow}`
                )

                // Add $/1000 formula in column H
                const dollarFormula = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
                spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
                spreadsheet.cellFormat(
                  {
                    fontWeight: 'bold',
                    textAlign: 'right',
                    backgroundColor: 'white',
                    format: '$#,##0.0'
                  },
                  `${pfx}H${currentRow}`
                )

                // Apply currency format
                try {
                  spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}:N${currentRow}`)
                } catch (e) {
                  // Fallback already applied in cellFormat
                }

                // Row height already set above based on proposal text

                currentRow++ // Move to next row
              })

              return // Skip normal processing for supporting angles
            }

            // Normal processing for other bracing items
            bracingGroups.forEach((group) => {
              if (group.length === 0) return

              // Calculate totals for the group
              let totalTakeoff = 0
              let totalQty = 0
              let totalFT = 0
              let totalLBS = 0
              let lastRowNumber = 0
              let itemDescription = ''

              group.forEach(item => {
                totalTakeoff += item.takeoff || 0
                totalQty += item.qty || 0
                const itemFT = (item.height || 0) * (item.takeoff || 0)
                totalFT += itemFT
                totalLBS += item.lbs || 0
                lastRowNumber = Math.max(lastRowNumber, item.rawRowNumber || 0)
                // Extract item description from particulars if available
                if (!itemDescription && item.particulars) {
                  itemDescription = item.particulars
                }
              })

              // Find the sum row for this group
              const sumRowIndex = lastRowNumber
              const calcSheetName = 'Calculations Sheet'

              // Get display name for this subsection - try both the original name and case variations
              let displayName = bracingDisplayNames[bracingSubsectionName] ||
                bracingDisplayNames[bracingSubsectionName.toLowerCase()] ||
                bracingDisplayNames[bracingSubsectionName.charAt(0).toUpperCase() + bracingSubsectionName.slice(1).toLowerCase()] ||
                bracingSubsectionName

              // Generate proposal text for bracing item
              // Format: F&I new (X)no [item description] [item type] as per SOE-XXX.XX
              const qtyValue = Math.round(totalQty || totalTakeoff)
              let proposalText = ''

              // Extract size/type from item description if available
              let sizeInfo = ''
              if (itemDescription) {
                // Try to extract size information (e.g., "W12x26", "HP12x63", etc.)
                const sizeMatch = itemDescription.match(/([A-Z]+\d+x\d+[A-Z]?)/i)
                if (sizeMatch) {
                  sizeInfo = `[${sizeMatch[1]}] `
                }
              }

              // Get SOE page reference from raw data
              let soePageMain = 'SOE-100.00'
              if (rawData && Array.isArray(rawData) && rawData.length > 1) {
                const headers = rawData[0]
                const dataRows = rawData.slice(1)
                const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
                const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

                if (digitizerIdx !== -1 && pageIdx !== -1) {
                  for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                    const row = dataRows[rowIndex]
                    const digitizerItem = row[digitizerIdx]
                    if (digitizerItem && typeof digitizerItem === 'string') {
                      const itemText = digitizerItem.toLowerCase()
                      if (itemText.includes(bracingSubsectionName.toLowerCase())) {
                        const pageValue = row[pageIdx]
                        if (pageValue) {
                          const pageStr = String(pageValue).trim()
                          const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                          if (soeMatches && soeMatches.length > 0) {
                            soePageMain = soeMatches[0]
                            break
                          }
                        }
                      }
                    }
                  }
                }
              }

              if (sizeInfo) {
                proposalText = `F&I new (${qtyValue})no ${sizeInfo}${displayName.toLowerCase()} as per ${soePageMain}`
              } else {
                proposalText = `F&I new (${qtyValue})no ${displayName.toLowerCase()} as per ${soePageMain}`
              }

              // Add proposal text
              spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  color: '#000000',
                  textAlign: 'left',
                  backgroundColor: 'white',
                  verticalAlign: 'top'
                },
                `${pfx}B${currentRow}`
              )

              // Calculate and set row height based on content
              const dynamicHeight = calculateRowHeight(proposalText)
              spreadsheet.setRowHeight(dynamicHeight, currentRow, proposalSheetIndex)

              // Add FT (LF) to column C - reference to calculation sheet sum row
              if (sumRowIndex > 0) {
                spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${sumRowIndex}` }, `${pfx}C${currentRow}`)
                spreadsheet.cellFormat(
                  {
                    fontWeight: 'bold',
                    textAlign: 'right',
                    backgroundColor: 'white'
                  },
                  `${pfx}C${currentRow}`
                )
              }

              // Add LBS to column E - reference to calculation sheet sum row
              if (sumRowIndex > 0) {
                spreadsheet.updateCell({ formula: `='${calcSheetName}'!K${sumRowIndex}` }, `${pfx}E${currentRow}`)
                spreadsheet.cellFormat(
                  {
                    fontWeight: 'bold',
                    textAlign: 'right',
                    backgroundColor: 'white'
                  },
                  `${pfx}E${currentRow}`
                )
              }

              // Add QTY to column G - reference to calculation sheet sum row
              if (sumRowIndex > 0) {
                spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
                spreadsheet.cellFormat(
                  {
                    fontWeight: 'bold',
                    textAlign: 'right',
                    backgroundColor: 'white'
                  },
                  `${pfx}G${currentRow}`
                )
              }

              // Add 180 to second LF column (column I) - same as other SOE items
              spreadsheet.updateCell({ value: 180 }, `${pfx}I${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: '#E2EFDA'
                },
                `${pfx}I${currentRow}:N${currentRow}`
              )

              // Add $/1000 formula in column H
              const dollarFormula = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
              spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white',
                  format: '$#,##0.0'
                },
                `${pfx}H${currentRow}`
              )

              // Apply currency format
              try {
                spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}:N${currentRow}`)
              } catch (e) {
                // Fallback already applied in cellFormat
              }

              // Row height already set above based on proposal text

              currentRow++ // Move to next row
            })
          })
          // Add Tie back anchor section right after Bracing
          // Add Tie back anchor header
          const tieBackText = `Tie back anchor: including applicable washers, steel bearing plates, locking hex nuts as required`
          spreadsheet.updateCell({ value: tieBackText }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: '#D0CECE',
              textDecoration: 'underline',
              border: '1px solid #000000'
            },
            `${pfx}B${currentRow}`
          )
          spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
          currentRow++

          // Get tie back anchor groups
          let tieBackGroups = soeSubsectionItems.get('Tie back') || []
          const tieBackAnchorGroups = soeSubsectionItems.get('Tie back anchor') || []
          tieBackGroups = [...tieBackGroups, ...tieBackAnchorGroups]

          // If not found, try to find from calculation data
          if (tieBackGroups.length === 0 && calculationData && calculationData.length > 0) {
            const foundRows = []
            let inSubsection = false

            calculationData.forEach((row, index) => {
              const rowNum = index + 2
              const colB = row[1]

              if (colB && typeof colB === 'string') {
                const bText = colB.trim().toLowerCase()
                if ((bText.includes('tie back') || bText.includes('tie back anchor')) && bText.endsWith(':')) {
                  inSubsection = true
                  return
                }

                if (inSubsection) {
                  if (bText.endsWith(':') && !bText.includes('tie back')) {
                    inSubsection = false
                    return
                  }

                  if (bText && !bText.endsWith(':') && row[2] !== undefined) {
                    const takeoff = parseFloat(row[2]) || 0
                    if (takeoff > 0) {
                      foundRows.push({
                        rowNum: rowNum,
                        particulars: colB.trim(),
                        takeoff: takeoff,
                        qty: parseFloat(row[12]) || 0,
                        height: parseFloat(row[7]) || 0,
                        lbs: parseFloat(row[10]) || 0,
                        rawRowNumber: rowNum
                      })
                    }
                  }
                }
              }
            })

            if (foundRows.length > 0) {
              tieBackGroups = [foundRows]
            }
          }

          // Process tie back anchor groups
          if (tieBackGroups.length > 0) {
            tieBackGroups.forEach((group) => {
              if (group.length === 0) return

              // Calculate totals
              let totalTakeoff = 0
              let totalQty = 0
              let totalFT = 0
              let totalLBS = 0
              let lastRowNumber = 0

              group.forEach(item => {
                totalTakeoff += item.takeoff || 0
                totalQty += item.qty || 0
                const itemFT = (item.height || 0) * (item.takeoff || 0)
                totalFT += itemFT
                totalLBS += item.lbs || 0
                lastRowNumber = Math.max(lastRowNumber, item.rawRowNumber || 0)
              })

              const sumRowIndex = lastRowNumber
              const calcSheetName = 'Calculations Sheet'

              // Get first item's particulars for extracting details
              let firstItemParticulars = ''
              group.forEach(item => {
                if (!firstItemParticulars && item.particulars) {
                  firstItemParticulars = item.particulars
                }
              })

              // Extract free length and bond length from particulars
              let freeLength = "28'-0\""
              let bondLength = "20'-0\""
              if (firstItemParticulars) {
                const freeLengthMatch = firstItemParticulars.match(/Free\s+length\s*=\s*([0-9'\"\-\s]+)/i)
                if (freeLengthMatch) {
                  freeLength = freeLengthMatch[1].trim()
                }
                const bondLengthMatch = firstItemParticulars.match(/Bond\s+length\s*=\s*([0-9'\"\-\s]+)/i)
                if (bondLengthMatch) {
                  bondLength = bondLengthMatch[1].trim()
                }
              }

              // Extract drill hole size (default to 5 ½"Ø)
              let drillHole = '5 ½"Ø'
              if (firstItemParticulars) {
                const drillMatch = firstItemParticulars.match(/([\d\s½¼¾]+)"?\s*Ø/i)
                if (drillMatch) {
                  drillHole = drillMatch[1].trim() + '"Ø'
                }
              }

              // Get SOE page references from raw data
              let soePageMain = 'SOE-002.00'
              if (rawData && Array.isArray(rawData) && rawData.length > 1) {
                const headers = rawData[0]
                const dataRows = rawData.slice(1)
                const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
                const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

                if (digitizerIdx !== -1 && pageIdx !== -1) {
                  for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                    const row = dataRows[rowIndex]
                    const digitizerItem = row[digitizerIdx]
                    if (digitizerItem && typeof digitizerItem === 'string') {
                      const itemText = digitizerItem.toLowerCase()
                      if (itemText.includes('tie back') || itemText.includes('hollow bar')) {
                        const pageValue = row[pageIdx]
                        if (pageValue) {
                          const pageStr = String(pageValue).trim()
                          const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                          if (soeMatches && soeMatches.length > 0) {
                            soePageMain = soeMatches[0]
                            break
                          }
                        }
                      }
                    }
                  }
                }
              }

              // Generate proposal text: F&I new (X)no tie back anchors (L=XX'-XX" + XX'-XX" bond length), XX ½"Ø drill hole as per SOE-XXX.XX
              const qtyValue = Math.round(totalQty || totalTakeoff)
              const proposalText = `F&I new (${qtyValue})no tie back anchors (L=${freeLength} + ${bondLength} bond length), ${drillHole} drill hole as per ${soePageMain}`

              spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  color: '#000000',
                  textAlign: 'left',
                  backgroundColor: 'white',
                  verticalAlign: 'top'
                },
                `${pfx}B${currentRow}`
              )

              // Add FT (LF) to column C
              if (sumRowIndex > 0) {
                spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${sumRowIndex}` }, `${pfx}C${currentRow}`)
                spreadsheet.cellFormat(
                  {
                    fontWeight: 'bold',
                    textAlign: 'right',
                    backgroundColor: 'white'
                  },
                  `${pfx}C${currentRow}`
                )
              }

              // Add LBS to column E
              if (sumRowIndex > 0) {
                spreadsheet.updateCell({ formula: `='${calcSheetName}'!K${sumRowIndex}` }, `${pfx}E${currentRow}`)
                spreadsheet.cellFormat(
                  {
                    fontWeight: 'bold',
                    textAlign: 'right',
                    backgroundColor: 'white'
                  },
                  `${pfx}E${currentRow}`
                )
              }

              // Add QTY to column G
              if (sumRowIndex > 0) {
                spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
                spreadsheet.cellFormat(
                  {
                    fontWeight: 'bold',
                    textAlign: 'right',
                    backgroundColor: 'white'
                  },
                  `${pfx}G${currentRow}`
                )
              }

              // Add 180 to second LF column (column I)
              spreadsheet.updateCell({ value: 180 }, `${pfx}I${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: '#E2EFDA'
                },
                `${pfx}I${currentRow}:N${currentRow}`
              )

              // Add $/1000 formula in column H
              const dollarFormula = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
              spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white',
                  format: '$#,##0.0'
                },
                `${pfx}H${currentRow}`
              )

              // Apply currency format
              try {
                spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}:N${currentRow}`)
              } catch (e) {
                // Fallback already applied in cellFormat
              }

              // Calculate and set row height based on content in column B
              const proposalTextForHeight = proposalText || (spreadsheet.getCellValue ? spreadsheet.getCellValue(`${pfx}B${currentRow}`) : '')
              const dynamicHeight = calculateRowHeight(proposalTextForHeight)
              spreadsheet.setRowHeight(dynamicHeight, currentRow, proposalSheetIndex)

              currentRow++
            })
          }

          return // Skip the normal processing for Bracing
        }

        // Add Tie back anchor section if it's in the list (but it should already be handled above)
        if (subsectionName === 'Tie back' || subsectionName === 'Tie back anchor') {
          // Add Tie back anchor header
          const tieBackText = `Tie back anchor: including applicable washers, steel bearing plates, locking hex nuts as required`
          spreadsheet.updateCell({ value: tieBackText }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: '#D0CECE',
              textDecoration: 'underline',
              border: '1px solid #000000'
            },
            `${pfx}B${currentRow}`
          )
          spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
          currentRow++

          // Get tie back anchor groups
          let tieBackGroups = soeSubsectionItems.get('Tie back') || []
          const tieBackAnchorGroups = soeSubsectionItems.get('Tie back anchor') || []
          tieBackGroups = [...tieBackGroups, ...tieBackAnchorGroups]

          // If not found, try to find from calculation data
          if (tieBackGroups.length === 0 && calculationData && calculationData.length > 0) {
            const foundRows = []
            let inSubsection = false

            calculationData.forEach((row, index) => {
              const rowNum = index + 2
              const colB = row[1]

              if (colB && typeof colB === 'string') {
                const bText = colB.trim().toLowerCase()
                if ((bText.includes('tie back') || bText.includes('tie back anchor')) && bText.endsWith(':')) {
                  inSubsection = true
                  return
                }

                if (inSubsection) {
                  if (bText.endsWith(':') && !bText.includes('tie back')) {
                    inSubsection = false
                    return
                  }

                  if (bText && !bText.endsWith(':') && row[2] !== undefined) {
                    const takeoff = parseFloat(row[2]) || 0
                    if (takeoff > 0) {
                      foundRows.push({
                        rowNum: rowNum,
                        particulars: colB.trim(),
                        takeoff: takeoff,
                        qty: parseFloat(row[12]) || 0,
                        height: parseFloat(row[7]) || 0,
                        lbs: parseFloat(row[10]) || 0,
                        rawRowNumber: rowNum
                      })
                    }
                  }
                }
              }
            })

            if (foundRows.length > 0) {
              tieBackGroups = [foundRows]
            }
          }

          // Process tie back anchor groups
          if (tieBackGroups.length > 0) {
            tieBackGroups.forEach((group) => {
              if (group.length === 0) return

              // Calculate totals
              let totalTakeoff = 0
              let totalQty = 0
              let totalFT = 0
              let totalLBS = 0
              let lastRowNumber = 0

              group.forEach(item => {
                totalTakeoff += item.takeoff || 0
                totalQty += item.qty || 0
                const itemFT = (item.height || 0) * (item.takeoff || 0)
                totalFT += itemFT
                totalLBS += item.lbs || 0
                lastRowNumber = Math.max(lastRowNumber, item.rawRowNumber || 0)
              })

              const sumRowIndex = lastRowNumber
              const calcSheetName = 'Calculations Sheet'

              // Get first item's particulars for extracting details
              let firstItemParticulars = ''
              group.forEach(item => {
                if (!firstItemParticulars && item.particulars) {
                  firstItemParticulars = item.particulars
                }
              })

              // Extract free length and bond length from particulars
              let freeLength = "28'-0\""
              let bondLength = "20'-0\""
              if (firstItemParticulars) {
                const freeLengthMatch = firstItemParticulars.match(/Free\s+length\s*=\s*([0-9'\"\-\s]+)/i)
                if (freeLengthMatch) {
                  freeLength = freeLengthMatch[1].trim()
                }
                const bondLengthMatch = firstItemParticulars.match(/Bond\s+length\s*=\s*([0-9'\"\-\s]+)/i)
                if (bondLengthMatch) {
                  bondLength = bondLengthMatch[1].trim()
                }
              }

              // Extract drill hole size (default to 5 ½"Ø)
              let drillHole = '5 ½"Ø'
              if (firstItemParticulars) {
                const drillMatch = firstItemParticulars.match(/([\d\s½¼¾]+)"?\s*Ø/i)
                if (drillMatch) {
                  drillHole = drillMatch[1].trim() + '"Ø'
                }
              }

              // Get SOE page references from raw data
              let soePageMain = 'SOE-002.00'
              if (rawData && Array.isArray(rawData) && rawData.length > 1) {
                const headers = rawData[0]
                const dataRows = rawData.slice(1)
                const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
                const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

                if (digitizerIdx !== -1 && pageIdx !== -1) {
                  for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                    const row = dataRows[rowIndex]
                    const digitizerItem = row[digitizerIdx]
                    if (digitizerItem && typeof digitizerItem === 'string') {
                      const itemText = digitizerItem.toLowerCase()
                      if (itemText.includes('tie back') || itemText.includes('hollow bar')) {
                        const pageValue = row[pageIdx]
                        if (pageValue) {
                          const pageStr = String(pageValue).trim()
                          const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                          if (soeMatches && soeMatches.length > 0) {
                            soePageMain = soeMatches[0]
                            break
                          }
                        }
                      }
                    }
                  }
                }
              }

              // Generate proposal text: F&I new (X)no tie back anchors (L=XX'-XX" + XX'-XX" bond length), XX ½"Ø drill hole as per SOE-XXX.XX
              const qtyValue = Math.round(totalQty || totalTakeoff)
              const proposalText = `F&I new (${qtyValue})no tie back anchors (L=${freeLength} + ${bondLength} bond length), ${drillHole} drill hole as per ${soePageMain}`

              spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  color: '#000000',
                  textAlign: 'left',
                  backgroundColor: 'white',
                  verticalAlign: 'top'
                },
                `${pfx}B${currentRow}`
              )

              // Add FT (LF) to column C
              if (sumRowIndex > 0) {
                spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${sumRowIndex}` }, `${pfx}C${currentRow}`)
                spreadsheet.cellFormat(
                  {
                    fontWeight: 'bold',
                    textAlign: 'right',
                    backgroundColor: 'white'
                  },
                  `${pfx}C${currentRow}`
                )
              }

              // Add LBS to column E
              if (sumRowIndex > 0) {
                spreadsheet.updateCell({ formula: `='${calcSheetName}'!K${sumRowIndex}` }, `${pfx}E${currentRow}`)
                spreadsheet.cellFormat(
                  {
                    fontWeight: 'bold',
                    textAlign: 'right',
                    backgroundColor: 'white'
                  },
                  `${pfx}E${currentRow}`
                )
              }

              // Add QTY to column G
              if (sumRowIndex > 0) {
                spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
                spreadsheet.cellFormat(
                  {
                    fontWeight: 'bold',
                    textAlign: 'right',
                    backgroundColor: 'white'
                  },
                  `${pfx}G${currentRow}`
                )
              }

              // Add 180 to second LF column (column I)
              spreadsheet.updateCell({ value: 180 }, `${pfx}I${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: '#E2EFDA'
                },
                `${pfx}I${currentRow}:N${currentRow}`
              )

              // Add $/1000 formula in column H
              const dollarFormula = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
              spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white',
                  format: '$#,##0.0'
                },
                `${pfx}H${currentRow}`
              )

              // Apply currency format
              try {
                spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}:N${currentRow}`)
              } catch (e) {
                // Fallback already applied in cellFormat
              }

              // Calculate and set row height based on content in column B
              const proposalTextForHeight = proposalText || (spreadsheet.getCellValue ? spreadsheet.getCellValue(`${pfx}B${currentRow}`) : '')
              const dynamicHeight = calculateRowHeight(proposalTextForHeight)
              spreadsheet.setRowHeight(dynamicHeight, currentRow, proposalSheetIndex)

              currentRow++
            })
          }

          // Add Tie down anchor section right after Tie back anchor (only if items exist)
          // Get tie down anchor groups first to check if any exist
          let tieDownGroups = soeSubsectionItems.get('Tie down') || soeSubsectionItems.get('Anchor') || []
          const tieDownAnchorGroups = soeSubsectionItems.get('Tie down anchor') || []
          tieDownGroups = [...tieDownGroups, ...tieDownAnchorGroups]

          // If not found, try to find from calculation data
          if (tieDownGroups.length === 0 && calculationData && calculationData.length > 0) {
            const foundRows = []
            let inSubsection = false

            calculationData.forEach((row, index) => {
              const rowNum = index + 2
              const colB = row[1]

              if (colB && typeof colB === 'string') {
                const bText = colB.trim().toLowerCase()
                if ((bText.includes('tie down') || bText.includes('tie down anchor')) && bText.endsWith(':')) {
                  inSubsection = true
                  return
                }

                if (inSubsection) {
                  if (bText.endsWith(':') && !bText.includes('tie down')) {
                    inSubsection = false
                    return
                  }

                  if (bText && !bText.endsWith(':') && row[2] !== undefined) {
                    const takeoff = parseFloat(row[2]) || 0
                    if (takeoff > 0) {
                      foundRows.push({
                        rowNum: rowNum,
                        particulars: colB.trim(),
                        takeoff: takeoff,
                        qty: parseFloat(row[12]) || 0,
                        height: parseFloat(row[7]) || 0,
                        lbs: parseFloat(row[10]) || 0,
                        rawRowNumber: rowNum
                      })
                    }
                  }
                }
              }
            })

            if (foundRows.length > 0) {
              tieDownGroups = [foundRows]
            }
          }

          // Only add Tie down anchor section if items exist
          if (tieDownGroups.length > 0 && tieDownGroups.some(g => g.length > 0)) {
            // Add Tie down anchor header
            const tieDownText = `Tie down anchor: including applicable washers, steel bearing plates, locking hex nuts as required`
            spreadsheet.updateCell({ value: tieDownText }, `${pfx}B${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                color: '#000000',
                textAlign: 'left',
                backgroundColor: '#D0CECE',
                textDecoration: 'underline',
                border: '1px solid #000000'
              },
              `${pfx}B${currentRow}`
            )
            spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
            currentRow++

            // Process tie down anchor groups
            tieDownGroups.forEach((group) => {
              if (group.length === 0) return

              // Calculate totals
              let totalTakeoff = 0
              let totalQty = 0
              let totalFT = 0
              let totalLBS = 0
              let lastRowNumber = 0
              let firstItemParticulars = ''

              group.forEach(item => {
                totalTakeoff += item.takeoff || 0
                totalQty += item.qty || 0
                const itemFT = (item.height || 0) * (item.takeoff || 0)
                totalFT += itemFT
                totalLBS += item.lbs || 0
                lastRowNumber = Math.max(lastRowNumber, item.rawRowNumber || 0)
                // Get first item's particulars for extracting details
                if (!firstItemParticulars && item.particulars) {
                  firstItemParticulars = item.particulars
                }
              })

              const sumRowIndex = lastRowNumber
              const calcSheetName = 'Calculations Sheet'

              // Extract free length and bond length from particulars
              let freeLength = "28'-0\""
              let bondLength = "20'-0\""
              if (firstItemParticulars) {
                const freeLengthMatch = firstItemParticulars.match(/Free\s+length\s*=\s*([0-9'\"\-\s]+)/i)
                if (freeLengthMatch) {
                  freeLength = freeLengthMatch[1].trim()
                }
                const bondLengthMatch = firstItemParticulars.match(/Bond\s+length\s*=\s*([0-9'\"\-\s]+)/i)
                if (bondLengthMatch) {
                  bondLength = bondLengthMatch[1].trim()
                }
              }

              // Extract drill hole size (default to 5 ½"Ø)
              let drillHole = '5 ½"Ø'
              if (firstItemParticulars) {
                const drillMatch = firstItemParticulars.match(/([\d\s½¼¾]+)"?\s*Ø/i)
                if (drillMatch) {
                  drillHole = drillMatch[1].trim() + '"Ø'
                }
              }

              // Get SOE page references from raw data
              let soePageMain = 'SOE-002.00'
              if (rawData && Array.isArray(rawData) && rawData.length > 1) {
                const headers = rawData[0]
                const dataRows = rawData.slice(1)
                const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
                const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

                if (digitizerIdx !== -1 && pageIdx !== -1) {
                  for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                    const row = dataRows[rowIndex]
                    const digitizerItem = row[digitizerIdx]
                    if (digitizerItem && typeof digitizerItem === 'string') {
                      const itemText = digitizerItem.toLowerCase()
                      if (itemText.includes('tie down') || itemText.includes('hollow down')) {
                        const pageValue = row[pageIdx]
                        if (pageValue) {
                          const pageStr = String(pageValue).trim()
                          const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                          if (soeMatches && soeMatches.length > 0) {
                            soePageMain = soeMatches[0]
                            break
                          }
                        }
                      }
                    }
                  }
                }
              }

              // Generate proposal text: F&I new (X)no tie down anchors (L=XX'-XX" + XX'-XX" bond length), XX ½"Ø drill hole as per SOE-XXX.XX
              const qtyValue = Math.round(totalQty || totalTakeoff)
              const proposalText = `F&I new (${qtyValue})no tie down anchors (L=${freeLength} + ${bondLength} bond length), ${drillHole} drill hole as per ${soePageMain}`

              spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  color: '#000000',
                  textAlign: 'left',
                  backgroundColor: 'white',
                  verticalAlign: 'top'
                },
                `${pfx}B${currentRow}`
              )

              // Add FT (LF) to column C
              if (sumRowIndex > 0) {
                spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${sumRowIndex}` }, `${pfx}C${currentRow}`)
                spreadsheet.cellFormat(
                  {
                    fontWeight: 'bold',
                    textAlign: 'right',
                    backgroundColor: 'white'
                  },
                  `${pfx}C${currentRow}`
                )
              }

              // Add LBS to column E
              if (sumRowIndex > 0) {
                spreadsheet.updateCell({ formula: `='${calcSheetName}'!K${sumRowIndex}` }, `${pfx}E${currentRow}`)
                spreadsheet.cellFormat(
                  {
                    fontWeight: 'bold',
                    textAlign: 'right',
                    backgroundColor: 'white'
                  },
                  `${pfx}E${currentRow}`
                )
              }

              // Add QTY to column G
              if (sumRowIndex > 0) {
                spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
                spreadsheet.cellFormat(
                  {
                    fontWeight: 'bold',
                    textAlign: 'right',
                    backgroundColor: 'white'
                  },
                  `${pfx}G${currentRow}`
                )
              }

              // Add 180 to second LF column (column I)
              spreadsheet.updateCell({ value: 180 }, `${pfx}I${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: '#E2EFDA'
                },
                `${pfx}I${currentRow}:N${currentRow}`
              )

              // Add $/1000 formula in column H
              const dollarFormula = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
              spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white',
                  format: '$#,##0.0'
                },
                `${pfx}H${currentRow}`
              )

              // Apply currency format
              try {
                spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}:N${currentRow}`)
              } catch (e) {
                // Fallback already applied in cellFormat
              }

              // Calculate and set row height based on content in column B
              const proposalTextForHeight = proposalText || (spreadsheet.getCellValue ? spreadsheet.getCellValue(`${pfx}B${currentRow}`) : '')
              const dynamicHeight = calculateRowHeight(proposalTextForHeight)
              spreadsheet.setRowHeight(dynamicHeight, currentRow, proposalSheetIndex)

              currentRow++
            })
          }

          return // Skip normal processing for Tie back anchor (Tie down anchor is already added above)
        }

        // Skip normal processing for Tie down anchor (it's already added right after Tie back anchor above)
        if (subsectionName === 'Tie down' || subsectionName === 'Tie down anchor') {
          return
        }

        // Special handling for Misc. subsection - handle it early and return
        if (subsectionName.toLowerCase() === 'misc.' || subsectionName.toLowerCase() === 'misc') {
          // Add subsection header
          spreadsheet.updateCell({ value: `${subsectionName}:` }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: '#D0CECE',
              textDecoration: 'underline',
              border: '1px solid #000000'
            },
            `${pfx}B${currentRow}`
          )
          spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
          currentRow++

          // Add hardcoded text items
          const miscItems = [
            'Cut-off to grade included',
            'Pilings will be threaded at both ends and installed in 10\' or 15\' increments',
            'Obstruction removal drilling using DTHH per hour: $995/h'
          ]

          miscItems.forEach((item) => {
            spreadsheet.updateCell({ value: item }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat(
              {
                fontWeight: 'normal',
                color: '#000000',
                textAlign: 'left',
                backgroundColor: 'white',
                verticalAlign: 'top'
              },
              `${pfx}B${currentRow}`
            )
            spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
            currentRow++
          })

          // Skip all other processing for Misc.
          return
        }

        // Add subsection header with grey background
        let subsectionText = `${subsectionName}:`

        // Add special notes for certain subsections
        if (subsectionName.toLowerCase().includes('tie back')) {
          subsectionText = `Tie back anchor: including applicable washers, steel bearing plates, locking hex nuts as required`
        } else if (subsectionName.toLowerCase().includes('tie down')) {
          subsectionText = `Tie down anchor: including applicable washers, steel bearing plates, locking hex nuts as required`
        }

        spreadsheet.updateCell({ value: subsectionText }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            backgroundColor: '#D0CECE',
            textDecoration: 'underline',
            border: '1px solid #000000'
          },
          `${pfx}B${currentRow}`
        )

        // Increase row height
        spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)

        currentRow++ // Move to next row

        // Process items for this subsection and generate proposal text
        // Handle mapping: if subsectionName is "Dowels", also check for "Dowel bar"
        // If subsectionName is "Concrete buttons", also check for "Buttons"
        let groups = soeSubsectionItems.get(subsectionName) || []

        // Also try case-insensitive match for soeSubsectionItems
        if (groups.length === 0) {
          for (const [key, value] of soeSubsectionItems.entries()) {
            if (key.toLowerCase() === subsectionName.toLowerCase()) {
              groups = value || []
              break
            }
          }
        }


        if (subsectionName === 'Dowels') {
          const dowelBarGroups = soeSubsectionItems.get('Dowel bar') || []
          groups = [...groups, ...dowelBarGroups]
        } else if (subsectionName === 'Concrete buttons' || subsectionName === 'Buttons') {
          // Also check for "Buttons" subsection name if looking for "Concrete buttons"
          if (subsectionName === 'Concrete buttons') {
            const buttonsGroups = soeSubsectionItems.get('Buttons') || []
            groups = [...groups, ...buttonsGroups]
          }
          // Also check for "Concrete buttons" if looking for "Buttons"
          if (subsectionName === 'Buttons') {
            const concreteButtonsGroups = soeSubsectionItems.get('Concrete buttons') || []
            groups = [...groups, ...concreteButtonsGroups]
          }
        } else if (subsectionName === 'Guilde wall' || subsectionName === 'Guide wall') {
          // Also check for "Guide wall" if looking for "Guilde wall" (typo)
          if (subsectionName === 'Guilde wall') {
            const guideWallGroups = soeSubsectionItems.get('Guide wall') || []
            groups = [...groups, ...guideWallGroups]
          }
          // Also check for "Guilde wall" if looking for "Guide wall"
          if (subsectionName === 'Guide wall') {
            const guildeWallGroups = soeSubsectionItems.get('Guilde wall') || []
            groups = [...groups, ...guildeWallGroups]
          }
        }


        // For Parging, Heel blocks, Underpinning, Concrete soil retention pier, Form board, Guide wall, Dowels, Rock pins, Rock stabilization, Shotcrete, and Buttons/Concrete buttons, first check window items, then soeSubsectionItems, then calculationData
        if ((subsectionName.toLowerCase() === 'parging' ||
          subsectionName.toLowerCase() === 'heel blocks' ||
          subsectionName.toLowerCase() === 'underpinning' ||
          subsectionName.toLowerCase() === 'concrete soil retention piers' ||
          subsectionName.toLowerCase() === 'concrete soil retention pier' ||
          subsectionName.toLowerCase() === 'form board' ||
          subsectionName.toLowerCase() === 'guide wall' ||
          subsectionName.toLowerCase() === 'guilde wall' ||
          subsectionName.toLowerCase() === 'dowels' ||
          subsectionName.toLowerCase() === 'dowel bar' ||
          subsectionName.toLowerCase() === 'rock pins' ||
          subsectionName.toLowerCase() === 'rock pin' ||
          subsectionName.toLowerCase() === 'rock stabilization' ||
          subsectionName.toLowerCase() === 'shotcrete' ||
          subsectionName.toLowerCase() === 'permission grouting' ||
          subsectionName.toLowerCase() === 'mud slab' ||
          subsectionName.toLowerCase() === 'buttons' ||
          subsectionName.toLowerCase() === 'concrete buttons' ||
          subsectionName.toLowerCase() === 'sheet pile' ||
          subsectionName.toLowerCase() === 'sheet piles' ||
          subsectionName.toLowerCase() === 'timber lagging' ||
          subsectionName.toLowerCase() === 'timber sheeting')) {

          // For these subsections, check window items first even if groups is not empty
          // This ensures we get the processed items from generateCalculationSheet
          const isParging = subsectionName.toLowerCase() === 'parging'
          const isHeelBlocks = subsectionName.toLowerCase() === 'heel blocks'
          const isUnderpinning = subsectionName.toLowerCase() === 'underpinning'
          const isConcreteSoilRetention = subsectionName.toLowerCase() === 'concrete soil retention piers' || subsectionName.toLowerCase() === 'concrete soil retention pier'
          const isFormBoard = subsectionName.toLowerCase() === 'form board'
          const isGuideWall = subsectionName.toLowerCase() === 'guide wall' || subsectionName.toLowerCase() === 'guilde wall'
          const isDowels = subsectionName.toLowerCase() === 'dowels' || subsectionName.toLowerCase() === 'dowel bar'
          const isRockPins = subsectionName.toLowerCase() === 'rock pins' || subsectionName.toLowerCase() === 'rock pin'
          const isRockStabilization = subsectionName.toLowerCase() === 'rock stabilization'
          const isShotcrete = subsectionName.toLowerCase() === 'shotcrete'
          const isPermissionGrouting = subsectionName.toLowerCase() === 'permission grouting'
          const isMudSlab = subsectionName.toLowerCase() === 'mud slab'
          const isButtons = subsectionName.toLowerCase() === 'buttons' || subsectionName.toLowerCase() === 'concrete buttons'
          const isSheetPile = subsectionName.toLowerCase() === 'sheet pile' || subsectionName.toLowerCase() === 'sheet piles'
          const isTimberLagging = subsectionName.toLowerCase() === 'timber lagging'
          const isTimberSheeting = subsectionName.toLowerCase() === 'timber sheeting'

          // Check window items first (these are processed items from generateCalculationSheet)
          if (isParging && window.pargingItems && window.pargingItems.length > 0) {
            const formattedItems = window.pargingItems.map(item => ({
              particulars: item.particulars || '',
              takeoff: item.takeoff || 0,
              unit: item.unit || '',
              height: item.parsed?.heightRaw || item.height || 0,
              sqft: 0,
              lbs: 0,
              qty: item.qty || 0,
              rawRowNumber: item.rawRowNumber || 0
            }))
            groups = [formattedItems]
          } else if (isGuideWall && window.guideWallItems && window.guideWallItems.length > 0) {
            const formattedItems = window.guideWallItems.map(item => ({
              particulars: item.particulars || '',
              takeoff: item.takeoff || 0,
              unit: item.unit || '',
              height: item.parsed?.heightRaw || item.height || 0,
              sqft: 0,
              lbs: 0,
              qty: item.qty || 0,
              parsed: item.parsed || {},
              rawRowNumber: item.rawRowNumber || 0
            }))
            groups = [formattedItems]
          } else if (isDowels && window.dowelBarItems && window.dowelBarItems.length > 0) {
            const formattedItems = window.dowelBarItems.map(item => ({
              particulars: item.particulars || '',
              takeoff: item.takeoff || 0,
              unit: item.unit || '',
              height: item.parsed?.heightRaw || item.height || 0,
              sqft: 0,
              lbs: 0,
              qty: item.qty || 0,
              parsed: item.parsed || {},
              rawRowNumber: item.rawRowNumber || 0
            }))
            groups = [formattedItems]
          } else if (isRockPins && window.rockPinItems && window.rockPinItems.length > 0) {
            const formattedItems = window.rockPinItems.map(item => ({
              particulars: item.particulars || '',
              takeoff: item.takeoff || 0,
              unit: item.unit || '',
              height: item.parsed?.heightRaw || item.height || 0,
              sqft: 0,
              lbs: 0,
              qty: item.qty || 0,
              parsed: item.parsed || {},
              rawRowNumber: item.rawRowNumber || 0
            }))
            groups = [formattedItems]
          } else if (isRockStabilization && window.rockStabilizationItems && window.rockStabilizationItems.length > 0) {
            const formattedItems = window.rockStabilizationItems.map(item => ({
              particulars: item.particulars || '',
              takeoff: item.takeoff || 0,
              unit: item.unit || '',
              height: item.parsed?.heightRaw || item.height || 0,
              sqft: 0,
              lbs: 0,
              qty: item.qty || 0,
              parsed: item.parsed || {},
              rawRowNumber: item.rawRowNumber || 0
            }))
            groups = [formattedItems]
          } else if (isShotcrete && window.shotcreteItems && window.shotcreteItems.length > 0) {
            const formattedItems = window.shotcreteItems.map(item => ({
              particulars: item.particulars || '',
              takeoff: item.takeoff || 0,
              unit: item.unit || '',
              height: item.parsed?.heightRaw || item.height || 0,
              sqft: 0,
              lbs: 0,
              qty: item.qty || 0,
              parsed: item.parsed || {},
              rawRowNumber: item.rawRowNumber || 0
            }))
            groups = [formattedItems]
          } else if (isPermissionGrouting && window.permissionGroutingItems && window.permissionGroutingItems.length > 0) {
            const formattedItems = window.permissionGroutingItems.map(item => ({
              particulars: item.particulars || '',
              takeoff: item.takeoff || 0,
              unit: item.unit || '',
              height: item.parsed?.heightRaw || item.height || 0,
              sqft: 0,
              lbs: 0,
              qty: item.qty || 0,
              parsed: item.parsed || {},
              rawRowNumber: item.rawRowNumber || 0
            }))
            groups = [formattedItems]
          } else if (isMudSlab && window.mudSlabItems && window.mudSlabItems.length > 0) {
            const formattedItems = window.mudSlabItems.map(item => ({
              particulars: item.particulars || '',
              takeoff: item.takeoff || 0,
              unit: item.unit || '',
              height: item.parsed?.heightRaw || item.height || 0,
              sqft: 0,
              lbs: 0,
              qty: item.qty || 0,
              parsed: item.parsed || {},
              rawRowNumber: item.rawRowNumber || 0
            }))
            groups = [formattedItems]
          } else if (isButtons && window.buttonItems && window.buttonItems.length > 0) {
            const formattedItems = window.buttonItems.map(item => ({
              particulars: item.particulars || '',
              takeoff: item.takeoff || 0,
              unit: item.unit || '',
              height: item.parsed?.height || item.parsed?.heightRaw || item.height || 0,
              sqft: 0,
              lbs: 0,
              qty: item.qty || 0,
              parsed: item.parsed || {},
              rawRowNumber: item.rawRowNumber || 0
            }))
            groups = [formattedItems]
          } else if (isSheetPile && window.sheetPileItems && window.sheetPileItems.length > 0) {
            const formattedItems = window.sheetPileItems.map(item => ({
              particulars: item.particulars || '',
              takeoff: item.takeoff || 0,
              unit: item.unit || '',
              height: item.parsed?.heightRaw || item.height || 0,
              sqft: 0,
              lbs: 0,
              qty: item.qty || 0,
              parsed: item.parsed || {},
              rawRowNumber: item.rawRowNumber || 0
            }))
            groups = [formattedItems]
          }

          // If still no groups, then check soeSubsectionItems and calculationData
          if (groups.length === 0) {
            const searchTerm = isParging ? 'parging' : isHeelBlocks ? 'heel block' : isUnderpinning ? 'underpinning' : isConcreteSoilRetention ? 'concrete soil retention pier' : isFormBoard ? 'form board' : isGuideWall ? 'guide wall' : isDowels ? 'dowel' : isRockPins ? 'rock pin' : isRockStabilization ? 'rock stabilization' : isShotcrete ? 'shotcrete' : isPermissionGrouting ? 'permission grouting' : isMudSlab ? 'mud slab' : isButtons ? 'button' : isSheetPile ? 'sheet pile' : isTimberLagging ? 'timber lagging' : isTimberSheeting ? 'timber sheeting' : ''

            // For parging, check window.pargingItems (already checked above, but keep for fallback)
            if (isParging) {
              const pargingItemsFromWindow = window.pargingItems || []
              if (pargingItemsFromWindow.length > 0) {
                // Convert pargingItems to the format expected by groups
                const formattedItems = pargingItemsFromWindow.map(item => ({
                  particulars: item.particulars || '',
                  takeoff: item.takeoff || 0,
                  unit: item.unit || '',
                  height: item.parsed?.heightRaw || item.height || 0,
                  sqft: 0,
                  lbs: 0,
                  qty: item.qty || 0,
                  rawRowNumber: item.rawRowNumber || 0
                }))
                groups = [formattedItems]
              }
            }

            // For guide wall, check window.guideWallItems
            if (isGuideWall) {
              const guideWallItemsFromWindow = window.guideWallItems || []
              if (guideWallItemsFromWindow.length > 0) {
                // Convert guideWallItems to the format expected by groups
                const formattedItems = guideWallItemsFromWindow.map(item => ({
                  particulars: item.particulars || '',
                  takeoff: item.takeoff || 0,
                  unit: item.unit || '',
                  height: item.parsed?.heightRaw || item.height || 0,
                  sqft: 0,
                  lbs: 0,
                  qty: item.qty || 0,
                  parsed: item.parsed || {},
                  rawRowNumber: item.rawRowNumber || 0
                }))
                groups = [formattedItems]
              }
            }

            // For dowels, check window.dowelBarItems
            if (isDowels) {
              const dowelBarItemsFromWindow = window.dowelBarItems || []
              if (dowelBarItemsFromWindow.length > 0) {
                // Convert dowelBarItems to the format expected by groups
                const formattedItems = dowelBarItemsFromWindow.map(item => ({
                  particulars: item.particulars || '',
                  takeoff: item.takeoff || 0,
                  unit: item.unit || '',
                  height: item.parsed?.heightRaw || item.height || 0,
                  sqft: 0,
                  lbs: 0,
                  qty: item.qty || 0,
                  parsed: item.parsed || {},
                  rawRowNumber: item.rawRowNumber || 0
                }))
                groups = [formattedItems]
              }
            }

            // For rock pins, check window.rockPinItems
            if (isRockPins) {
              const rockPinItemsFromWindow = window.rockPinItems || []
              if (rockPinItemsFromWindow.length > 0) {
                // Convert rockPinItems to the format expected by groups
                const formattedItems = rockPinItemsFromWindow.map(item => ({
                  particulars: item.particulars || '',
                  takeoff: item.takeoff || 0,
                  unit: item.unit || '',
                  height: item.parsed?.heightRaw || item.height || 0,
                  sqft: 0,
                  lbs: 0,
                  qty: item.qty || 0,
                  parsed: item.parsed || {},
                  rawRowNumber: item.rawRowNumber || 0
                }))
                groups = [formattedItems]
              }
            }

            // For rock stabilization, check window.rockStabilizationItems
            if (isRockStabilization) {
              const rockStabilizationItemsFromWindow = window.rockStabilizationItems || []
              if (rockStabilizationItemsFromWindow.length > 0) {
                // Convert rockStabilizationItems to the format expected by groups
                const formattedItems = rockStabilizationItemsFromWindow.map(item => ({
                  particulars: item.particulars || '',
                  takeoff: item.takeoff || 0,
                  unit: item.unit || '',
                  height: item.parsed?.heightRaw || item.height || 0,
                  sqft: 0,
                  lbs: 0,
                  qty: item.qty || 0,
                  parsed: item.parsed || {},
                  rawRowNumber: item.rawRowNumber || 0
                }))
                groups = [formattedItems]
              }
            }

            // For shotcrete, check window.shotcreteItems
            if (isShotcrete) {
              const shotcreteItemsFromWindow = window.shotcreteItems || []
              if (shotcreteItemsFromWindow.length > 0) {
                // Convert shotcreteItems to the format expected by groups
                const formattedItems = shotcreteItemsFromWindow.map(item => ({
                  particulars: item.particulars || '',
                  takeoff: item.takeoff || 0,
                  unit: item.unit || '',
                  height: item.parsed?.heightRaw || item.height || 0,
                  sqft: 0,
                  lbs: 0,
                  qty: item.qty || 0,
                  parsed: item.parsed || {},
                  rawRowNumber: item.rawRowNumber || 0
                }))
                groups = [formattedItems]
              }
            }

            // For permission grouting, check window.permissionGroutingItems
            if (isPermissionGrouting) {
              const permissionGroutingItemsFromWindow = window.permissionGroutingItems || []
              if (permissionGroutingItemsFromWindow.length > 0) {
                // Convert permissionGroutingItems to the format expected by groups
                const formattedItems = permissionGroutingItemsFromWindow.map(item => ({
                  particulars: item.particulars || '',
                  takeoff: item.takeoff || 0,
                  unit: item.unit || '',
                  height: item.parsed?.heightRaw || item.height || 0,
                  sqft: 0,
                  lbs: 0,
                  qty: item.qty || 0,
                  parsed: item.parsed || {},
                  rawRowNumber: item.rawRowNumber || 0
                }))
                groups = [formattedItems]
              }
            }

            // For mud slab, check window.mudSlabItems
            if (isMudSlab) {
              const mudSlabItemsFromWindow = window.mudSlabItems || []
              if (mudSlabItemsFromWindow.length > 0) {
                // Convert mudSlabItems to the format expected by groups
                const formattedItems = mudSlabItemsFromWindow.map(item => ({
                  particulars: item.particulars || '',
                  takeoff: item.takeoff || 0,
                  unit: item.unit || '',
                  height: item.parsed?.heightRaw || item.height || 0,
                  sqft: 0,
                  lbs: 0,
                  qty: item.qty || 0,
                  parsed: item.parsed || {},
                  rawRowNumber: item.rawRowNumber || 0
                }))
                groups = [formattedItems]
              }
            }

            // For sheet pile, check window.sheetPileItems
            if (isSheetPile) {
              const sheetPileItemsFromWindow = window.sheetPileItems || []
              if (sheetPileItemsFromWindow.length > 0) {
                // Convert sheetPileItems to the format expected by groups
                const formattedItems = sheetPileItemsFromWindow.map(item => ({
                  particulars: item.particulars || '',
                  takeoff: item.takeoff || 0,
                  unit: item.unit || '',
                  height: item.parsed?.heightRaw || item.height || 0,
                  sqft: 0,
                  lbs: 0,
                  qty: item.qty || 0,
                  parsed: item.parsed || {},
                  rawRowNumber: item.rawRowNumber || 0
                }))
                groups = [formattedItems]
              }
            }

            // For buttons, check window.buttonItems
            if (isButtons) {
              const buttonItemsFromWindow = window.buttonItems || []
              if (buttonItemsFromWindow.length > 0) {
                // Convert buttonItems to the format expected by groups
                const formattedItems = buttonItemsFromWindow.map(item => ({
                  particulars: item.particulars || '',
                  takeoff: item.takeoff || 0,
                  unit: item.unit || '',
                  height: item.parsed?.height || item.parsed?.heightRaw || item.height || 0,
                  sqft: 0,
                  lbs: 0,
                  qty: item.qty || 0,
                  parsed: item.parsed || {},
                  rawRowNumber: item.rawRowNumber || 0
                }))
                groups = [formattedItems]
              }
            }

            // If still no groups, try soeSubsectionItems
            if (groups.length === 0) {
              let groupsFromMap = soeSubsectionItems.get(subsectionName) || []
              // Also try case-insensitive match
              if (groupsFromMap.length === 0) {
                for (const [key, value] of soeSubsectionItems.entries()) {
                  if (key.toLowerCase() === subsectionName.toLowerCase()) {
                    groupsFromMap = value || []
                    break
                  }
                }
              }
              if (groupsFromMap.length > 0 && groupsFromMap[0].length > 0) {
                groups = groupsFromMap
              } else {
                // Search calculationData directly
                const foundRows = []
                let inSubsection = false

                if (calculationData && calculationData.length > 0) {
                  calculationData.forEach((row, index) => {
                    const rowNum = index + 2
                    const colB = row[1]

                    if (colB && typeof colB === 'string') {
                      const bText = colB.trim()
                      const bTextLower = bText.toLowerCase()
                      // For underpinning, exclude shims from the main underpinning group
                      if (isUnderpinning && bTextLower.includes('shim')) {
                        // Skip shims - they'll be handled separately
                        return
                      }
                      if (bTextLower.includes(searchTerm) && bText.endsWith(':')) {
                        inSubsection = true
                        return
                      }

                      if (inSubsection) {
                        if (bText.endsWith(':') && !bTextLower.includes(searchTerm)) {
                          inSubsection = false
                          return
                        }

                        if (bText && !bText.endsWith(':') && row[2] !== undefined) {
                          const takeoff = parseFloat(row[2]) || 0
                          // For underpinning, exclude shims
                          if (isUnderpinning && bTextLower.includes('shim')) {
                            return
                          }
                          if (takeoff > 0 || bTextLower.includes(searchTerm)) {
                            // Parse height from row or from particulars for Timber lagging/Timber sheeting
                            let parsedHeight = parseFloat(row[7]) || 0
                            // Try to extract height from particulars if not in row
                            if (parsedHeight === 0 && (isTimberLagging || isTimberSheeting)) {
                              const hMatch = bText.match(/H=([0-9'"\-]+)/i)
                              if (hMatch) {
                                // Parse dimension string to feet
                                const parseDim = (dimStr) => {
                                  if (!dimStr) return 0
                                  const match = dimStr.match(/(\d+)(?:'-?)?(\d+)?/)
                                  if (!match) return 0
                                  const feet = parseInt(match[1]) || 0
                                  const inches = parseInt(match[2]) || 0
                                  return feet + (inches / 12)
                                }
                                parsedHeight = parseDim(hMatch[1])
                              }
                            }

                            foundRows.push({
                              rowNum: rowNum,
                              particulars: bText,
                              takeoff: takeoff,
                              qty: parseFloat(row[12]) || 0,
                              height: parsedHeight,
                              sqft: parseFloat(row[9]) || 0,
                              lbs: parseFloat(row[10]) || 0,
                              rawRowNumber: rowNum,
                              parsed: {
                                heightRaw: parsedHeight
                              }
                            })
                          }
                        }
                      }
                    }
                  })

                  if (foundRows.length > 0) {
                    groups = [foundRows]
                  }
                }
              }
            }
          }
        }

        // If no groups found but subsection is in display list, skip processing
        if (groups.length === 0 || groups.every(g => g.length === 0)) {
          return
        }

        // Find form board items related to concrete soil retention pier
        if (subsectionName.toLowerCase() === 'concrete soil retention piers' || subsectionName.toLowerCase() === 'concrete soil retention pier') {
          const formBoardItems = []

          // Check soeSubsectionItems for form board
          const formBoardGroups = soeSubsectionItems.get('Form board') || []
          if (formBoardGroups.length > 0) {
            formBoardGroups.forEach(group => {
              group.forEach(item => {
                formBoardItems.push({
                  particulars: item.particulars || '',
                  takeoff: item.takeoff || 0,
                  qty: item.qty || 0,
                  height: item.height || item.parsed?.heightRaw || 0,
                  sqft: item.sqft || 0,
                  rowIndex: item.rawRowNumber || 0
                })
              })
            })
          }

          // Also search calculation data for form board items
          if (calculationData && calculationData.length > 0) {
            let foundConcreteSoilRetention = false
            let inFormBoardSection = false

            for (let i = 0; i < calculationData.length; i++) {
              const row = calculationData[i]
              const colB = row[1]

              if (colB && typeof colB === 'string') {
                const bText = colB.trim()
                const bTextLower = bText.toLowerCase()

                // Check if we're in concrete soil retention pier section
                if (bTextLower.includes('concrete soil retention pier') && bText.endsWith(':')) {
                  foundConcreteSoilRetention = true
                  inFormBoardSection = false
                  continue
                }

                // Check if we've moved to form board section (after concrete soil retention pier)
                if (foundConcreteSoilRetention && bTextLower.includes('form board') && bText.endsWith(':')) {
                  inFormBoardSection = true
                  continue
                }

                // If we're in form board section and it's a data row (not a header)
                if (inFormBoardSection && bText && !bText.endsWith(':')) {
                  if (bTextLower.includes('form board')) {
                    const takeoff = parseFloat(row[2]) || 0
                    if (takeoff > 0 || bTextLower.includes('form board')) {
                      // Check if we already have this item (avoid duplicates)
                      const existing = formBoardItems.find(fb => fb.particulars === bText && fb.takeoff === takeoff)
                      if (!existing) {
                        formBoardItems.push({
                          particulars: bText,
                          takeoff: takeoff,
                          qty: parseFloat(row[12]) || 0,
                          height: parseFloat(row[7]) || 0,
                          sqft: parseFloat(row[9]) || 0,
                          rowIndex: i + 2
                        })
                      }
                    }
                  }
                }

                // Stop if we hit another subsection after form board
                if (inFormBoardSection && bText.endsWith(':') && !bTextLower.includes('form board')) {
                  break
                }
              }
            }
          }

        }

        // For Underpinning, separate underpinning items from shims
        let underpinningGroups = []
        let shimsGroups = []
        if (subsectionName.toLowerCase() === 'underpinning') {
          groups.forEach(group => {
            // Separate items within each group
            const underpinningItems = []
            const shimsItems = []

            group.forEach(item => {
              const particulars = (item.particulars || '').trim()

              // Check if this is a shims item
              // An item is a shim if:
              // 1. It starts with "Shims" (case-insensitive) as a standalone word, OR
              // 2. It contains "shim" or "wedge" as a standalone word
              // AND it does NOT start with "Underpinning" followed by dimensions
              const startsWithShims = /^shims\b/i.test(particulars)
              const hasShimWord = /\b(shim|wedge)\b/i.test(particulars)
              const isUnderpinningWithDimensions = /^underpinning\s+\d+['"]?\s*x\s*\d+['"]?/i.test(particulars)

              // If it starts with "Underpinning" followed by dimensions, it's definitely an underpinning item
              if (isUnderpinningWithDimensions) {
                underpinningItems.push(item)
              } else if (startsWithShims || hasShimWord) {
                // This is a shims item (has shim/wedge keyword and is not an underpinning item with dimensions)
                shimsItems.push(item)
              } else {
                // Default: treat as underpinning item if we're in the underpinning subsection
                underpinningItems.push(item)
              }
            })

            // Add non-empty groups
            if (underpinningItems.length > 0) {
              underpinningGroups.push(underpinningItems)
            }
            if (shimsItems.length > 0) {
              shimsGroups.push(shimsItems)
            }
          })
        } else {
          underpinningGroups = groups
        }

        // Process underpinning items first
        underpinningGroups.forEach((group) => {
          if (group.length === 0) return

          // Calculate totals for the group
          let totalTakeoff = 0
          let totalQty = 0
          let totalFT = 0
          let totalLBS = 0
          let totalSQFT = 0
          let lastRowNumber = 0
          let avgHeight = 0
          let heightCount = 0

          // Helper function to parse dimension string (e.g., "21'-1"" -> 21.083)
          const parseDimension = (dimStr) => {
            if (!dimStr) return 0
            const match = dimStr.match(/(\d+)(?:'-?)?(\d+)?/)
            if (!match) return 0
            const feet = parseInt(match[1]) || 0
            const inches = parseInt(match[2]) || 0
            return feet + (inches / 12)
          }

          group.forEach(item => {
            totalTakeoff += item.takeoff || 0
            totalQty += item.qty || 0
            // For parging, FT(I) = C, so FT is just the takeoff value, not height * takeoff
            let itemFT = 0
            if (subsectionName.toLowerCase() === 'parging') {
              itemFT = item.takeoff || 0
            } else {
              itemFT = (item.height || 0) * (item.takeoff || 0)
            }
            totalFT += itemFT
            totalLBS += item.lbs || 0
            totalSQFT += item.sqft || 0

            // For Parging and Underpinning, extract height from particulars if not in height field
            let itemHeight = item.height || 0
            if ((subsectionName.toLowerCase() === 'parging' || subsectionName.toLowerCase() === 'underpinning') && !itemHeight && item.particulars) {
              const particulars = item.particulars || ''
              // Try H= pattern first
              let hMatch = particulars.match(/H=([0-9'"\-]+)/i) || particulars.match(/Height=([0-9'"\-]+)/i)
              if (hMatch) {
                itemHeight = parseDimension(hMatch[1])
              } else if (subsectionName.toLowerCase() === 'underpinning') {
                // For underpinning, try to parse from format like "2'-4"x4'-0" wide, Height=4'-7""
                const heightMatch = particulars.match(/Height=([0-9'"\-]+)/i)
                if (heightMatch) {
                  itemHeight = parseDimension(heightMatch[1])
                }
              }
            }

            if (itemHeight > 0) {
              avgHeight += itemHeight * (item.takeoff || 1)
              heightCount += item.takeoff || 1
            }
            lastRowNumber = Math.max(lastRowNumber, item.rawRowNumber || 0)
          })

          if (heightCount > 0) {
            avgHeight = avgHeight / heightCount
          }

          // Find the sum row for this group
          // For parging, heel blocks, underpinning, concrete soil retention piers, and form board, the sum row is the row after the last item (where the sum formula is)
          let sumRowIndex = lastRowNumber
          const subsectionLower = subsectionName.toLowerCase()
          if (subsectionLower === 'parging' ||
            subsectionLower === 'heel blocks' ||
            subsectionLower === 'underpinning' ||
            subsectionLower === 'concrete soil retention piers' ||
            subsectionLower === 'concrete soil retention pier' ||
            subsectionLower === 'form board') {
            // For these subsections, sum row is after the last item row
            sumRowIndex = lastRowNumber + 1
          }
          const calcSheetName = 'Calculations Sheet'

          // Generate proposal text based on subsection type
          let proposalText = ''

          // Special handling for Underpinning
          if (subsectionName.toLowerCase() === 'underpinning') {
            // Parse underpinning items to extract width and height
            let minWidth = Infinity
            let maxWidth = 0
            let totalWidth = 0
            let widthCount = 0

            group.forEach(item => {
              const particulars = item.particulars || ''
              // Parse width from particulars like "2'-4"x1'-0" wide" or "2'-4"x4'-0" wide"
              const widthMatch = particulars.match(/(\d+)['"]?\s*x\s*(\d+)['"]?\s*-\s*(\d+)['"]?\s*wide/i) ||
                particulars.match(/(\d+)['"]?\s*x\s*(\d+)['"]?\s*wide/i)
              if (widthMatch) {
                // Width is the second dimension (after x)
                const widthFeet = parseInt(widthMatch[2]) || 0
                const widthInches = parseInt(widthMatch[3]) || 0
                const width = widthFeet + (widthInches / 12)

                if (width > 0) {
                  minWidth = Math.min(minWidth, width)
                  maxWidth = Math.max(maxWidth, width)
                  totalWidth += width * (item.takeoff || 1)
                  widthCount += item.takeoff || 1
                }
              } else {
                // Try to get width from parsed data (column G)
                const width = item.parsed?.width || item.width || 0
                if (width > 0) {
                  minWidth = Math.min(minWidth, width)
                  maxWidth = Math.max(maxWidth, width)
                  totalWidth += width * (item.takeoff || 1)
                  widthCount += item.takeoff || 1
                }
              }
            })

            // Format width range
            let widthText = ''
            if (minWidth < Infinity && maxWidth > 0) {
              if (minWidth === maxWidth) {
                const wFeet = Math.floor(minWidth)
                const wInches = Math.round((minWidth - wFeet) * 12)
                widthText = wInches === 0 ? `${wFeet}'-0"` : `${wFeet}'-${wInches}"`
              } else {
                const minWFeet = Math.floor(minWidth)
                const minWInches = Math.round((minWidth - minWFeet) * 12)
                const maxWFeet = Math.floor(maxWidth)
                const maxWInches = Math.round((maxWidth - maxWFeet) * 12)
                const minWText = minWInches === 0 ? `${minWFeet}'-0"` : `${minWFeet}'-${minWInches}"`
                const maxWText = maxWInches === 0 ? `${maxWFeet}'-0"` : `${maxWFeet}'-${maxWInches}"`
                widthText = `${minWText} to ${maxWText}`
              }
            }

            // Format average height
            const roundToMultipleOf5 = (value) => Math.ceil(value / 5) * 5
            const avgHeightRounded = heightCount > 0 ? roundToMultipleOf5(avgHeight) : 0
            const heightFeet = Math.floor(avgHeightRounded)
            const heightInches = Math.round((avgHeightRounded - heightFeet) * 12)
            const heightText = heightInches === 0 ? `${heightFeet}'-0"` : `${heightFeet}'-${heightInches}"`

            // Get total quantity
            const totalQtyValue = Math.round(totalQty || totalTakeoff || 0)

            // Get SOE page reference from raw data
            let soePageMain = 'SOE-101.00'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                  const row = dataRows[rowIndex]
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('underpinning') && !itemText.includes('shim')) {
                      const pageValue = row[pageIdx]
                      if (pageValue) {
                        const pageStr = String(pageValue).trim()
                        const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                        if (soeMatches && soeMatches.length > 0) {
                          soePageMain = soeMatches[0]
                          break
                        }
                      }
                    }
                  }
                }
              }
            }

            // Format: F&I new (##)no. (#'-0" to #'-0" wide) underpinning piers (Havg=#'-#"), reinforced w/#5@12"O.C., E.W., E.F.
            if (widthText) {
              proposalText = `F&I new (${totalQtyValue})no. (${widthText} wide) underpinning piers (Havg=${heightText}), reinforced w/#5@12"O.C., E.W., E.F.`
            } else {
              proposalText = `F&I new (${totalQtyValue})no. underpinning piers (Havg=${heightText}), reinforced w/#5@12"O.C., E.W., E.F.`
            }
          } else if (subsectionName.toLowerCase() === 'heel blocks') {
            // Get total quantity from sum row (column M)
            const totalQtyValue = Math.round(totalQty || totalTakeoff || 0)

            // Get SOE page reference from raw data
            let soePageMain = 'SOE-101.00'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                  const row = dataRows[rowIndex]
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('heel block')) {
                      const pageValue = row[pageIdx]
                      if (pageValue) {
                        const pageStr = String(pageValue).trim()
                        const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                        if (soeMatches && soeMatches.length > 0) {
                          soePageMain = soeMatches[0]
                          break
                        }
                      }
                    }
                  }
                }
              }
            }

            // Format: F&I new (X)no typ. heel blocks for raker support as per SOE-XXX.XX
            proposalText = `F&I new (${totalQtyValue})no typ. heel blocks for raker support as per ${soePageMain}`
          } else if (subsectionName.toLowerCase() === 'concrete soil retention piers' || subsectionName.toLowerCase() === 'concrete soil retention pier') {
            // Get total quantity from sum row (column M)
            const totalQtyValue = Math.round(totalQty || totalTakeoff || 0)

            // Get SOE page reference from raw data
            let soePageMain = 'SOE-101.00'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                  const row = dataRows[rowIndex]
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('concrete soil retention pier')) {
                      const pageValue = row[pageIdx]
                      if (pageValue) {
                        const pageStr = String(pageValue).trim()
                        const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                        if (soeMatches && soeMatches.length > 0) {
                          soePageMain = soeMatches[0]
                          break
                        }
                      }
                    }
                  }
                }
              }
            }

            // Format: F&I new (X)no concrete soil retention pier as per SOE-XXX.XX
            proposalText = `F&I new (${totalQtyValue})no concrete soil retention pier as per ${soePageMain}`

            // Find and add form board proposal text right after concrete soil retention pier
            // Search for form board items in calculation data
            const formBoardItems = []
            if (calculationData && calculationData.length > 0) {
              let foundConcreteSoilRetention = false
              let inFormBoardSection = false

              for (let i = 0; i < calculationData.length; i++) {
                const row = calculationData[i]
                const colB = row[1]

                if (colB && typeof colB === 'string') {
                  const bText = colB.trim()
                  const bTextLower = bText.toLowerCase()

                  // Check if we're in concrete soil retention pier section
                  if (bTextLower.includes('concrete soil retention pier') && bText.endsWith(':')) {
                    foundConcreteSoilRetention = true
                    inFormBoardSection = false
                    continue
                  }

                  // Check if we've moved to form board section (after concrete soil retention pier)
                  if (foundConcreteSoilRetention && bTextLower.includes('form board') && bText.endsWith(':')) {
                    inFormBoardSection = true
                    continue
                  }

                  // If we're in form board section and it's a data row (not a header)
                  if (inFormBoardSection && bText && !bText.endsWith(':')) {
                    if (bTextLower.includes('form board')) {
                      const takeoff = parseFloat(row[2]) || 0
                      if (takeoff > 0 || bTextLower.includes('form board')) {
                        formBoardItems.push({
                          particulars: bText,
                          takeoff: takeoff,
                          qty: parseFloat(row[12]) || 0,
                          height: parseFloat(row[7]) || 0,
                          sqft: parseFloat(row[9]) || 0,
                          rowIndex: i + 2
                        })
                      }
                    }
                  }

                  // Stop if we hit another subsection after form board
                  if (inFormBoardSection && bText.endsWith(':') && !bTextLower.includes('form board')) {
                    break
                  }
                }
              }
            }

            // Also check soeSubsectionItems for form board
            const formBoardGroups = soeSubsectionItems.get('Form board') || []
            if (formBoardGroups.length > 0) {
              formBoardGroups.forEach(group => {
                group.forEach(item => {
                  formBoardItems.push({
                    particulars: item.particulars || '',
                    takeoff: item.takeoff || 0,
                    qty: item.qty || 0,
                    height: item.height || item.parsed?.heightRaw || 0,
                    sqft: item.sqft || 0,
                    rowIndex: item.rawRowNumber || 0
                  })
                })
              })
            }

            // Generate form board proposal text if items found
            if (formBoardItems.length > 0) {
              // Calculate totals
              let formBoardTotalQty = 0
              let formBoardTotalTakeoff = 0
              formBoardItems.forEach(item => {
                formBoardTotalQty += item.qty || 0
                formBoardTotalTakeoff += item.takeoff || 0
              })
              const formBoardTotalQtyValue = Math.round(formBoardTotalQty || formBoardTotalTakeoff || 0)

              // Extract thickness from items
              let formBoardThickness = '1"' // Default thickness
              if (formBoardItems.length > 0) {
                const firstItem = formBoardItems[0]
                const particulars = (firstItem.particulars || '').trim()
                const thicknessMatch = particulars.match(/(\d+(?:\/\d+)?)"?\s*form\s*board/i)
                if (thicknessMatch) {
                  formBoardThickness = `${thicknessMatch[1]}"`
                }
              }

              // Get SOE page reference for form board
              let formBoardSoePage = 'SOE-300.00' // Default for form board
              if (rawData && Array.isArray(rawData) && rawData.length > 1) {
                const headers = rawData[0]
                const dataRows = rawData.slice(1)
                const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
                const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

                if (digitizerIdx !== -1 && pageIdx !== -1) {
                  for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                    const row = dataRows[rowIndex]
                    const digitizerItem = row[digitizerIdx]
                    if (digitizerItem && typeof digitizerItem === 'string') {
                      const itemText = digitizerItem.toLowerCase()
                      if (itemText.includes('form board')) {
                        const pageValue = row[pageIdx]
                        if (pageValue) {
                          const pageStr = String(pageValue).trim()
                          const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                          if (soeMatches && soeMatches.length > 0) {
                            formBoardSoePage = soeMatches[0]
                            break
                          }
                        }
                      }
                    }
                  }
                }
              }

              // Store form board proposal text to add after concrete soil retention pier
              window.formBoardProposalText = `F&I new (${formBoardThickness} thick) form board w/ filter fabric between tunnel and retention pier as per ${formBoardSoePage}`
              window.formBoardItems = formBoardItems
            }

            // Console log the proposal text line and data
          } else if (subsectionName.toLowerCase() === 'parging') {
            // Calculate average height
            const roundToMultipleOf5 = (value) => Math.ceil(value / 5) * 5
            const avgHeightRounded = roundToMultipleOf5(avgHeight)

            // Format height
            const heightFeet = Math.floor(avgHeightRounded)
            const heightInches = Math.round((avgHeightRounded - heightFeet) * 12)
            let heightText = ''
            if (heightInches === 0) {
              heightText = `${heightFeet}'-0"`
            } else {
              heightText = `${heightFeet}'-${heightInches}"`
            }

            // Get SOE page reference from raw data
            let soePageMain = 'SOE-101.00'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                  const row = dataRows[rowIndex]
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('parging')) {
                      const pageValue = row[pageIdx]
                      if (pageValue) {
                        const pageStr = String(pageValue).trim()
                        const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                        if (soeMatches && soeMatches.length > 0) {
                          soePageMain = soeMatches[0]
                          break
                        }
                      }
                    }
                  }
                }
              }
            }

            // Format: New parging (Havg=XX'-XX") as per SOE-XXX.XX
            proposalText = `New parging (Havg=${heightText}) as per ${soePageMain}`
          } else if (subsectionName.toLowerCase() === 'form board') {
            // Get total quantity from sum row (column M)
            const totalQtyValue = Math.round(totalQty || totalTakeoff || 0)

            // Extract thickness from items (e.g., "1" form board" -> "1"")
            let thickness = '1"' // Default thickness
            if (group.length > 0) {
              const firstItem = group[0]
              const particulars = (firstItem.particulars || '').trim()
              // Match pattern like "1" form board" or "1\" form board"
              const thicknessMatch = particulars.match(/(\d+(?:\/\d+)?)"?\s*form\s*board/i)
              if (thicknessMatch) {
                thickness = `${thicknessMatch[1]}"`
              }
            }

            // Get SOE page reference from raw data
            let soePageMain = 'SOE-300.00' // Default for form board
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                  const row = dataRows[rowIndex]
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('form board')) {
                      const pageValue = row[pageIdx]
                      if (pageValue) {
                        const pageStr = String(pageValue).trim()
                        const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                        if (soeMatches && soeMatches.length > 0) {
                          soePageMain = soeMatches[0]
                          break
                        }
                      }
                    }
                  }
                }
              }
            }

            // Format: F&I new (1" thick) form board w/ filter fabric between tunnel and retention pier as per SOE-300.00
            proposalText = `F&I new (${thickness} thick) form board w/ filter fabric between tunnel and retention pier as per ${soePageMain}`
          } else if (subsectionName.toLowerCase() === 'concrete buttons' || subsectionName.toLowerCase() === 'buttons') {
            // Get total quantity from sum row (column M)
            const totalQtyValue = Math.round(totalQty || totalTakeoff || 0)

            // Extract width from items (for format: 3'-0"x3'-0" wide)
            let widthText = ''
            if (group.length > 0) {
              const firstItem = group[0]
              const particulars = firstItem.particulars || ''

              // Extract dimensions from bracket format: Concrete button (3'-0"x3'-0"x1'-0")
              const bracketMatch = particulars.match(/\(([^)]+)\)/)
              if (bracketMatch) {
                const dimsStr = bracketMatch[1].split('x').map(d => d.trim())
                if (dimsStr.length >= 2) {
                  // Use first two dimensions (length x width)
                  widthText = `${dimsStr[0]}x${dimsStr[1]}`
                }
              } else {
                // Fallback to parsed values if bracket not found
                const length = firstItem.parsed?.length
                const width = firstItem.parsed?.width
                if (length && width) {
                  // Format numeric values back to dimension strings
                  const formatDim = (val) => {
                    if (typeof val === 'number') {
                      const feet = Math.floor(val)
                      const inches = Math.round((val - feet) * 12)
                      return inches === 0 ? `${feet}'-0"` : `${feet}'-${inches}"`
                    }
                    return val
                  }
                  widthText = `${formatDim(length)}x${formatDim(width)}`
                }
              }
            }

            // Calculate average height from parsed height values
            let totalHeight = 0
            let heightCount = 0
            group.forEach(item => {
              // For buttons, height is in parsed.height (from bracket dimensions)
              const height = item.parsed?.height || item.parsed?.heightRaw || item.height || 0
              if (height > 0) {
                totalHeight += height * (item.takeoff || 0)
                heightCount += (item.takeoff || 0)
              }
            })
            const avgHeight = heightCount > 0 ? totalHeight / heightCount : 0

            // Format average height (use actual average, format as feet and inches)
            const heightFeet = Math.floor(avgHeight)
            const heightInches = Math.round((avgHeight - heightFeet) * 12)
            const heightText = heightInches === 0 ? `${heightFeet}'-0"` : `${heightFeet}'-${heightInches}"`

            // Get SOE page reference from raw data (could be SOESK-01 or SOE-XXX.XX)
            let soePageMain = 'SOE-101.00' // Default
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                  const row = dataRows[rowIndex]
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('button')) {
                      const pageValue = row[pageIdx]
                      if (pageValue) {
                        const pageStr = String(pageValue).trim()
                        // Match SOE-XXX.XX or SOESK-XX format
                        const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                        if (soeMatches && soeMatches.length > 0) {
                          soePageMain = soeMatches[0]
                          break
                        }
                      }
                    }
                  }
                }
              }
            }

            // Format: F&I new (12)no (3'-0"x3'-0" wide) concrete buttons (Havg=3'-3") as per SOESK-01
            proposalText = widthText
              ? `F&I new (${totalQtyValue})no (${widthText} wide) concrete buttons (Havg=${heightText}) as per ${soePageMain}`
              : `F&I new (${totalQtyValue})no concrete buttons (Havg=${heightText}) as per ${soePageMain}`
          } else if (subsectionName.toLowerCase() === 'guide wall' || subsectionName.toLowerCase() === 'guilde wall') {
            // Extract widths from items
            const widths = new Set()
            let height = ''
            if (group.length > 0) {
              group.forEach(item => {
                const particulars = item.particulars || ''
                const bracketMatch = particulars.match(/\(([^)]+)\)/)
                if (bracketMatch) {
                  const dimsStr = bracketMatch[1].split('x').map(d => d.trim())
                  if (dimsStr.length >= 1) {
                    widths.add(dimsStr[0])
                  }
                  if (dimsStr.length >= 2) {
                    height = dimsStr[1]
                  }
                }
              })
            }
            const widthText = Array.from(widths).join(' & ')

            // Get SOE page references
            let soePageMain = 'SOE-100.00'
            let soePageDetails = 'SOE-300.00'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                  const row = dataRows[rowIndex]
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('guide wall')) {
                      const pageValue = row[pageIdx]
                      if (pageValue) {
                        const pageStr = String(pageValue).trim()
                        const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                        if (soeMatches && soeMatches.length > 0) {
                          soePageMain = soeMatches[0]
                          if (soeMatches.length > 1) {
                            soePageDetails = soeMatches[1]
                          }
                        }
                      }
                    }
                  }
                }
              }
            }

            // Format: F&I new (4'-6½" & 5'-3½" wide) guide wall (H=3'-0") as per SOE-100.00 & details on SOE-300.00
            proposalText = widthText
              ? `F&I new (${widthText} wide) guide wall (H=${height || '3\'-0"'}) as per ${soePageMain}${soePageDetails ? ` & details on ${soePageDetails}` : ''}`
              : `F&I new guide wall (H=${height || '3\'-0"'}) as per ${soePageMain}${soePageDetails ? ` & details on ${soePageDetails}` : ''}`
          } else if (subsectionName.toLowerCase() === 'dowels' || subsectionName.toLowerCase() === 'dowel bar') {
            // Get total quantity
            const totalQtyValue = Math.round(totalQty || totalTakeoff || 0)

            // Extract bar size and rock socket
            let barSize = ''
            let rockSocket = ''
            group.forEach(item => {
              const particulars = item.particulars || ''
              const barMatch = particulars.match(/#(\d+)/)
              if (barMatch && !barSize) {
                barSize = `#${barMatch[1]}`
              }
              const rsMatch = particulars.match(/RS=([0-9'"\-]+)/i)
              if (rsMatch && !rockSocket) {
                rockSocket = rsMatch[1]
              }
            })

            // Calculate average height
            let totalHeight = 0
            let heightCount = 0
            group.forEach(item => {
              const height = item.parsed?.heightRaw || 0
              if (height > 0) {
                totalHeight += height * (item.takeoff || 0)
                heightCount += (item.takeoff || 0)
              }
            })
            const avgHeight = heightCount > 0 ? totalHeight / heightCount : 0
            const heightFeet = Math.floor(avgHeight)
            const heightInches = Math.round((avgHeight - heightFeet) * 12)
            const heightText = heightInches === 0 ? `${heightFeet}'-0"` : `${heightFeet}'-${heightInches}"`

            // Get SOE page references
            let soePageMain = 'SOESK-01'
            let soePageDetails = 'SOESK-02'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                  const row = dataRows[rowIndex]
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('dowel')) {
                      const pageValue = row[pageIdx]
                      if (pageValue) {
                        const pageStr = String(pageValue).trim()
                        const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                        if (soeMatches && soeMatches.length > 0) {
                          soePageMain = soeMatches[0]
                          if (soeMatches.length > 1) {
                            soePageDetails = soeMatches[1]
                          }
                        }
                      }
                    }
                  }
                }
              }
            }

            // Format: F&I new (48)no 4-#9 steel dowels bar (Havg=6'-3", 4'-0" rock socket) as per SOESK-01
            proposalText = `F&I new (${totalQtyValue})no 4-${barSize || '#9'} steel dowels bar (Havg=${heightText}, ${rockSocket || '4\'-0"'} rock socket) as per ${soePageMain}`
          } else if (subsectionName.toLowerCase() === 'rock pins' || subsectionName.toLowerCase() === 'rock pin') {
            // Get total quantity
            const totalQtyValue = Math.round(totalQty || totalTakeoff || 0)

            // Extract rock socket
            let rockSocket = ''
            group.forEach(item => {
              const particulars = item.particulars || ''
              const rsMatch = particulars.match(/RS=([0-9'"\-]+)/i)
              if (rsMatch && !rockSocket) {
                rockSocket = rsMatch[1]
              }
            })

            // Calculate average height
            let totalHeight = 0
            let heightCount = 0
            group.forEach(item => {
              const height = item.parsed?.heightRaw || 0
              if (height > 0) {
                totalHeight += height * (item.takeoff || 0)
                heightCount += (item.takeoff || 0)
              }
            })
            const avgHeight = heightCount > 0 ? totalHeight / heightCount : 0
            const heightFeet = Math.floor(avgHeight)
            const heightInches = Math.round((avgHeight - heightFeet) * 12)
            const heightText = heightInches === 0 ? `${heightFeet}'-0"` : `${heightFeet}'-${heightInches}"`

            // Get SOE page references
            let soePageMain = 'SOESK-01'
            let soePageDetails = 'SOESK-02'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                  const row = dataRows[rowIndex]
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('rock pin')) {
                      const pageValue = row[pageIdx]
                      if (pageValue) {
                        const pageStr = String(pageValue).trim()
                        const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                        if (soeMatches && soeMatches.length > 0) {
                          soePageMain = soeMatches[0]
                          if (soeMatches.length > 1) {
                            soePageDetails = soeMatches[1]
                          }
                        }
                      }
                    }
                  }
                }
              }
            }

            // Format: F&I new (24)no rock pins (Havg=6'-3", 4'-0" rock socket) as per SOESK-01
            proposalText = `F&I new (${totalQtyValue})no rock pins (Havg=${heightText}, ${rockSocket || '4\'-0"'} rock socket) as per ${soePageMain}`
          } else if (subsectionName.toLowerCase() === 'rock stabilization') {
            // Get SOE page references
            let soePageMain = 'SOE-100.00'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                  const row = dataRows[rowIndex]
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('rock stabilization')) {
                      const pageValue = row[pageIdx]
                      if (pageValue) {
                        const pageStr = String(pageValue).trim()
                        const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                        if (soeMatches && soeMatches.length > 0) {
                          soePageMain = soeMatches[0]
                        }
                      }
                    }
                  }
                }
              }
            }

            // Format: F&I new rock stabilization as per SOE-100.00
            proposalText = `F&I new rock stabilization as per ${soePageMain}`
          } else if (subsectionName.toLowerCase() === 'shotcrete') {
            // Extract thickness and wire mesh
            let thickness = ''
            let wireMesh = ''
            group.forEach(item => {
              const particulars = item.particulars || ''
              const thickMatch = particulars.match(/(\d+(?:\/\d+)?)"?\s*thick/i)
              if (thickMatch && !thickness) {
                thickness = `${thickMatch[1]}"`
              }
              const meshMatch = particulars.match(/(\d+x\d+)\s*wire\s*mesh/i)
              if (meshMatch && !wireMesh) {
                wireMesh = meshMatch[1]
              }
            })

            // Calculate average height
            let totalHeight = 0
            let heightCount = 0
            group.forEach(item => {
              const height = item.parsed?.heightRaw || 0
              if (height > 0) {
                totalHeight += height * (item.takeoff || 0)
                heightCount += (item.takeoff || 0)
              }
            })
            const avgHeight = heightCount > 0 ? totalHeight / heightCount : 0
            const heightFeet = Math.floor(avgHeight)
            const heightInches = Math.round((avgHeight - heightFeet) * 12)
            const heightText = heightInches === 0 ? `${heightFeet}'-0"` : `${heightFeet}'-${heightInches}"`

            // Get SOE page references
            let soePageMain = 'SOESK-01'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                  const row = dataRows[rowIndex]
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('shotcrete')) {
                      const pageValue = row[pageIdx]
                      if (pageValue) {
                        const pageStr = String(pageValue).trim()
                        const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                        if (soeMatches && soeMatches.length > 0) {
                          soePageMain = soeMatches[0]
                        }
                      }
                    }
                  }
                }
              }
            }

            // Format: F&I new (6" thick) shotcrete w/ 6x6 wire mesh (Havg=23'-0") as per SOESK-01
            proposalText = `F&I new (${thickness || '6"'} thick) shotcrete w/ ${wireMesh || '6x6'} wire mesh (Havg=${heightText}) as per ${soePageMain}`
          } else if (subsectionName.toLowerCase() === 'permission grouting') {
            // Calculate average height
            let totalHeight = 0
            let heightCount = 0
            group.forEach(item => {
              const height = item.parsed?.heightRaw || 0
              if (height > 0) {
                totalHeight += height * (item.takeoff || 0)
                heightCount += (item.takeoff || 0)
              }
            })
            const avgHeight = heightCount > 0 ? totalHeight / heightCount : 0
            const heightFeet = Math.floor(avgHeight)
            const heightInches = Math.round((avgHeight - heightFeet) * 12)
            const heightText = heightInches === 0 ? `${heightFeet}'-0"` : `${heightFeet}'-${heightInches}"`

            // Get SOE page references
            let soePageMain = 'SOESK-01'
            let soePageDetails = 'SOESK-02'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                  const row = dataRows[rowIndex]
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('permission grouting')) {
                      const pageValue = row[pageIdx]
                      if (pageValue) {
                        const pageStr = String(pageValue).trim()
                        const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                        if (soeMatches && soeMatches.length > 0) {
                          soePageMain = soeMatches[0]
                          if (soeMatches.length > 1) {
                            soePageDetails = soeMatches[1]
                          }
                        }
                      }
                    }
                  }
                }
              }
            }

            // Format: F&I new permission grouting (Havg=16'-0") as per SOESK-01
            proposalText = `F&I new permission grouting (Havg=${heightText}) as per ${soePageMain}`
          } else if (subsectionName.toLowerCase() === 'mud slab') {
            // Extract thickness from mud slab items (e.g., "w/ 2" mud slab" -> "2"")
            let thickness = '2"' // Default thickness
            if (group.length > 0) {
              for (const item of group) {
                const particulars = (item.particulars || '').trim()
                // Match pattern like "w/ 2" mud slab" or "w/ 2\" mud slab"
                const thicknessMatch = particulars.match(/w\/\s*(\d+(?:\/\d+)?)["']?\s*mud\s*slab/i)
                if (thicknessMatch) {
                  thickness = `${thicknessMatch[1]}"`
                  break
                }
              }
            }

            // Get SOE page reference from raw data
            let soePageMain = 'SOE-101.00'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                  const row = dataRows[rowIndex]
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('mud slab')) {
                      const pageValue = row[pageIdx]
                      if (pageValue) {
                        const pageStr = String(pageValue).trim()
                        const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                        if (soeMatches && soeMatches.length > 0) {
                          soePageMain = soeMatches[0]
                          break
                        }
                      }
                    }
                  }
                }
              }
            }

            // Format: Allow to (2" thick) mud slab as per SOE-101.00
            proposalText = `Allow to (${thickness} thick) mud slab as per ${soePageMain}`
          } else if (subsectionName.toLowerCase() === 'sheet pile' || subsectionName.toLowerCase() === 'sheet piles') {
            // Extract sheet pile type (e.g., "NZ-14", "PZC-12", etc.)
            let sheetPileType = ''
            if (group.length > 0) {
              const firstItem = group[0]
              const particulars = (firstItem.particulars || '').trim()

              // Match sheet pile type patterns: NZ-14, PZC-12, PZ-22, AZ-18, etc.
              const typeMatch = particulars.match(/\b([A-Z]{2,3}[- ]?\d+(?:[-/]\d+)?(?:[-/]\d+)?)\b/i)
              if (typeMatch) {
                sheetPileType = typeMatch[1].replace(/\s+/g, '-') // Normalize spaces to hyphens
              }
            }

            // Extract height (H) from items
            let height = ''
            let heightRaw = 0
            group.forEach(item => {
              const h = item.parsed?.heightRaw || 0
              if (h > 0 && !heightRaw) {
                heightRaw = h
                const heightFeet = Math.floor(h)
                const heightInches = Math.round((h - heightFeet) * 12)
                height = heightInches === 0 ? `${heightFeet}'-0"` : `${heightFeet}'-${heightInches}"`
              }
            })

            // Extract embedment (E) from items
            let embedment = ''
            if (group.length > 0) {
              const firstItem = group[0]
              const particulars = (firstItem.particulars || '').trim()
              const eMatch = particulars.match(/E=([0-9'"\-]+)/i)
              if (eMatch) {
                // Parse dimension string to feet
                const parseDim = (dimStr) => {
                  if (!dimStr) return 0
                  const match = dimStr.match(/(\d+)(?:'-?)?(\d+)?/)
                  if (!match) return 0
                  const feet = parseInt(match[1]) || 0
                  const inches = parseInt(match[2]) || 0
                  return feet + (inches / 12)
                }
                const embedmentRaw = parseDim(eMatch[1])
                if (embedmentRaw > 0) {
                  const embedmentFeet = Math.floor(embedmentRaw)
                  const embedmentInches = Math.round((embedmentRaw - embedmentFeet) * 12)
                  embedment = embedmentInches === 0 ? `${embedmentFeet}'-0"` : `${embedmentFeet}'-${embedmentInches}"`
                }
              }
            }

            // Get SOE page reference from raw data
            let soePageMain = 'SOE-102.00'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                  const row = dataRows[rowIndex]
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('sheet pile')) {
                      const pageValue = row[pageIdx]
                      if (pageValue) {
                        const pageStr = String(pageValue).trim()
                        const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                        if (soeMatches && soeMatches.length > 0) {
                          soePageMain = soeMatches[0]
                          break
                        }
                      }
                    }
                  }
                }
              }
            }

            // Format: F&I new [NZ-14] sheet pile (H=27'-0", E=18'-6"embedment) as per SOE-102.00
            if (embedment) {
              proposalText = `F&I new [${sheetPileType || '#'}] sheet pile (H=${height || '#'}, E=${embedment}embedment) as per ${soePageMain}`
            } else {
              proposalText = `F&I new [${sheetPileType || '#'}] sheet pile (H=${height || '#'}) as per ${soePageMain}`
            }
          } else if (subsectionName.toLowerCase() === 'timber lagging') {
            // Extract dimensions from items (e.g., "3"x10"" from "Timber lagging 3"x10"")
            let dimensions = ''
            if (group.length > 0) {
              const firstItem = group[0]
              const particulars = (firstItem.particulars || '').trim()

              // Match pattern like "3"x10"" or "3\"x10\"" or "3x10"
              const dimMatch = particulars.match(/(\d+(?:\/\d+)?)["']?\s*x\s*(\d+(?:\/\d+)?)["']?/i)
              if (dimMatch) {
                dimensions = `${dimMatch[1]}"x${dimMatch[2]}"`
              }
            }

            // Calculate average height
            let totalHeight = 0
            let heightCount = 0
            group.forEach(item => {
              const height = item.parsed?.heightRaw || item.height || 0
              if (height > 0) {
                totalHeight += height * (item.takeoff || 0)
                heightCount += (item.takeoff || 0)
              }
            })
            const avgHeight = heightCount > 0 ? totalHeight / heightCount : 0
            const heightFeet = Math.floor(avgHeight)
            const heightInches = Math.round((avgHeight - heightFeet) * 12)
            const heightText = heightInches === 0 ? `${heightFeet}'-0"` : `${heightFeet}'-${heightInches}"`

            // Get SOE page reference from raw data
            let soePageMain = 'SOE-101.00'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                  const row = dataRows[rowIndex]
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('timber lagging')) {
                      const pageValue = row[pageIdx]
                      if (pageValue) {
                        const pageStr = String(pageValue).trim()
                        const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                        if (soeMatches && soeMatches.length > 0) {
                          soePageMain = soeMatches[0]
                          break
                        }
                      }
                    }
                  }
                }
              }
            }

            // Format: F&I new 3"x10" timber lagging for the exposed depths (Havg=10'-6") as per SOE-101.00
            proposalText = `F&I new ${dimensions || '##'} timber lagging for the exposed depths (Havg=${heightText || '##'}) as per ${soePageMain}`

            // Store soePageMain for backpacking line
            window.timberLaggingSoePage = soePageMain
          } else if (subsectionName.toLowerCase() === 'timber sheeting') {
            // Extract dimensions from items (e.g., "3"x10"" from "Timber sheeting 3"x10"")
            let dimensions = ''
            if (group.length > 0) {
              const firstItem = group[0]
              const particulars = (firstItem.particulars || '').trim()

              // Match pattern like "3"x10"" or "3\"x10\"" or "3x10"
              const dimMatch = particulars.match(/(\d+(?:\/\d+)?)["']?\s*x\s*(\d+(?:\/\d+)?)["']?/i)
              if (dimMatch) {
                dimensions = `${dimMatch[1]}"x${dimMatch[2]}"`
              }
            }

            // Calculate average height
            let totalHeight = 0
            let heightCount = 0
            group.forEach(item => {
              const height = item.parsed?.heightRaw || item.height || 0
              if (height > 0) {
                totalHeight += height * (item.takeoff || 0)
                heightCount += (item.takeoff || 0)
              }
            })
            const avgHeight = heightCount > 0 ? totalHeight / heightCount : 0
            const heightFeet = Math.floor(avgHeight)
            const heightInches = Math.round((avgHeight - heightFeet) * 12)
            const heightText = heightInches === 0 ? `${heightFeet}'-0"` : `${heightFeet}'-${heightInches}"`

            // Get SOE page reference from raw data
            let soePageMain = 'SOE-101.00'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                  const row = dataRows[rowIndex]
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('timber sheeting')) {
                      const pageValue = row[pageIdx]
                      if (pageValue) {
                        const pageStr = String(pageValue).trim()
                        const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                        if (soeMatches && soeMatches.length > 0) {
                          soePageMain = soeMatches[0]
                          break
                        }
                      }
                    }
                  }
                }
              }
            }

            // Format: F&I new 3"x10" timber sheeting (Havg=10'-6") as per SOE-101.00
            proposalText = `F&I new ${dimensions || '##'} timber sheeting (Havg=${heightText || '##'}) as per ${soePageMain}`
          } else {
            // Default proposal text for other subsections
            proposalText = `${subsectionName} item: ${Math.round(totalQty)} nos`
          }

          // Add proposal text
          spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: 'white',
              verticalAlign: 'top'
            },
            `${pfx}B${currentRow}`
          )

          // Calculate and set row height based on content
          const dynamicHeight = calculateRowHeight(proposalText)
          spreadsheet.setRowHeight(dynamicHeight, currentRow, proposalSheetIndex)

          // For Timber lagging, add second line for backpacking
          if (subsectionName.toLowerCase() === 'timber lagging') {
            currentRow++
            const backpackingSoePage = window.timberLaggingSoePage || 'SOE-101.00'
            const backpackingText = `F&I new backpacking @ timber lagging ${backpackingSoePage}`
            spreadsheet.updateCell({ value: backpackingText }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                color: '#000000',
                textAlign: 'left',
                backgroundColor: 'white',
                verticalAlign: 'top'
              },
              `${pfx}B${currentRow}`
            )
            spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
          }

          // Add FT (LF) to column C - reference to calculation sheet sum row if available
          // For parging, FT(I) = C, so we should reference column I from the sum row (which sums column C)
          if (subsectionName.toLowerCase() === 'parging' && sumRowIndex > 0) {
            // For parging, column I in the sum row contains the sum of column C (takeoff)
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${sumRowIndex}` }, `${pfx}C${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}C${currentRow}`
            )
          } else if (totalFT > 0 && sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${sumRowIndex}` }, `${pfx}C${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}C${currentRow}`
            )
          }

          // Add SQFT to column D for Parging, Heel blocks, Underpinning, Concrete soil retention piers, and Form board - reference to calculation sheet sum row
          const subsectionLowerName2 = subsectionName.toLowerCase()
          if ((subsectionLowerName2 === 'parging' ||
            subsectionLowerName2 === 'heel blocks' ||
            subsectionLowerName2 === 'underpinning' ||
            subsectionLowerName2 === 'concrete soil retention piers' ||
            subsectionLowerName2 === 'concrete soil retention pier' ||
            subsectionLowerName2 === 'form board') && sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${sumRowIndex}` }, `${pfx}D${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}D${currentRow}`
            )
          }

          // Add CY to column F for Underpinning and Concrete soil retention piers - reference to calculation sheet sum row
          if ((subsectionLowerName2 === 'underpinning' ||
            subsectionLowerName2 === 'concrete soil retention piers' ||
            subsectionLowerName2 === 'concrete soil retention pier') && sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!L${sumRowIndex}` }, `${pfx}F${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}F${currentRow}`
            )
          }

          // Add QTY to column G for Underpinning and Concrete soil retention piers - reference to calculation sheet sum row
          if ((subsectionLowerName2 === 'underpinning' ||
            subsectionLowerName2 === 'concrete soil retention piers' ||
            subsectionLowerName2 === 'concrete soil retention pier') && sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}G${currentRow}`
            )
          }

          // Add LBS to column E - reference to calculation sheet sum row if available
          if (totalLBS > 0 && sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!K${sumRowIndex}` }, `${pfx}E${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}E${currentRow}`
            )
          }

          // Add QTY to column G - reference to calculation sheet sum row if available
          // For heel blocks, always add QTY from sum row
          if ((totalQty > 0 || subsectionName.toLowerCase() === 'heel blocks') && sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}G${currentRow}`
            )
          }

          // Add 180 to second LF column (column I) - same as other SOE items
          spreadsheet.updateCell({ value: 180 }, `${pfx}I${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: '#E2EFDA'
            },
            `${pfx}I${currentRow}:N${currentRow}`
          )

          // Add $/1000 formula in column H
          const dollarFormula = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
          spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: 'white',
              format: '$#,##0.0'
            },
            `${pfx}H${currentRow}`
          )

          // Apply currency format
          try {
            spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}:N${currentRow}`)
          } catch (e) {
            // Fallback already applied in cellFormat
          }

          // Row height already set above based on proposal text

          currentRow++ // Move to next row

          // If this is concrete soil retention pier and form board items exist, add form board text on next row
          if ((subsectionName.toLowerCase() === 'concrete soil retention piers' || subsectionName.toLowerCase() === 'concrete soil retention pier') && window.formBoardProposalText) {
            const formBoardText = window.formBoardProposalText
            const formBoardItems = window.formBoardItems || []

            // Add form board proposal text
            spreadsheet.updateCell({ value: formBoardText }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                color: '#000000',
                textAlign: 'left',
                backgroundColor: 'white',
                verticalAlign: 'top'
              },
              `${pfx}B${currentRow}`
            )

            // Calculate and set row height based on content
            const formBoardDynamicHeight = calculateRowHeight(formBoardText)
            spreadsheet.setRowHeight(formBoardDynamicHeight, currentRow, proposalSheetIndex)

            // Calculate form board totals for columns
            let formBoardTotalTakeoff = 0
            let formBoardTotalSQFT = 0
            let formBoardLastRowNumber = 0
            formBoardItems.forEach(item => {
              formBoardTotalTakeoff += item.takeoff || 0
              formBoardTotalSQFT += item.sqft || 0
              formBoardLastRowNumber = Math.max(formBoardLastRowNumber, item.rowIndex || 0)
            })

            // Find form board sum row in calculation sheet
            let formBoardSumRowIndex = formBoardLastRowNumber + 1

            // Add FT (LF) to column C - reference to calculation sheet sum row
            if (formBoardSumRowIndex > 0) {
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${formBoardSumRowIndex}` }, `${pfx}C${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white'
                },
                `${pfx}C${currentRow}`
              )
            }

            // Add SQFT to column D - reference to calculation sheet sum row
            if (formBoardSumRowIndex > 0) {
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${formBoardSumRowIndex}` }, `${pfx}D${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white'
                },
                `${pfx}D${currentRow}`
              )
            }

            // Add 180 to second LF column (column I)
            spreadsheet.updateCell({ value: 180 }, `${pfx}I${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: '#E2EFDA'
              },
              `${pfx}I${currentRow}:N${currentRow}`
            )

            // Add $/1000 formula in column H
            const formBoardDollarFormula = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
            spreadsheet.updateCell({ formula: formBoardDollarFormula }, `${pfx}H${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white',
                format: '$#,##0.0'
              },
              `${pfx}H${currentRow}`
            )

            // Apply currency format
            try {
              spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}:N${currentRow}`)
            } catch (e) {
              // Fallback already applied in cellFormat
            }

            // Row height already set above based on form board proposal text

            // Clear the stored form board data
            window.formBoardProposalText = null
            window.formBoardItems = null

            currentRow++ // Move to next row after form board
          }
        })

        // Process shims groups if this is underpinning
        if (subsectionName.toLowerCase() === 'underpinning' && shimsGroups.length > 0) {
          shimsGroups.forEach((shimsGroup) => {
            if (shimsGroup.length === 0) return

            // Calculate totals for shims
            let shimsTotalTakeoff = 0
            let shimsTotalQty = 0
            let shimsLastRowNumber = 0

            shimsGroup.forEach(item => {
              shimsTotalTakeoff += item.takeoff || 0
              shimsTotalQty += item.qty || 0
              shimsLastRowNumber = Math.max(shimsLastRowNumber, item.rawRowNumber || 0)
            })

            // Find the sum row for shims (after the last shims item)
            const shimsSumRowIndex = shimsLastRowNumber + 1

            // Generate shims proposal text
            const shimsQtyValue = Math.round(shimsTotalQty || shimsTotalTakeoff || 0)
            const shimsProposalText = `F&I new (${shimsQtyValue})no. sets of 2"x4" @2'-0" O.C. min shim/wedge plates for transfer of vertical load & fill the gab between B.O. existing foundation & T.O. new underpinning w/non-shrink grout/dry pack`

            // Add shims proposal text
            spreadsheet.updateCell({ value: shimsProposalText }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                color: '#000000',
                textAlign: 'left',
                backgroundColor: 'white',
                verticalAlign: 'top'
              },
              `${pfx}B${currentRow}`
            )

            // Calculate and set row height based on content
            const shimsDynamicHeight = calculateRowHeight(shimsProposalText)
            spreadsheet.setRowHeight(shimsDynamicHeight, currentRow, proposalSheetIndex)

            // Add QTY to column G for shims
            if (shimsSumRowIndex > 0) {
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${shimsSumRowIndex}` }, `${pfx}G${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white'
                },
                `${pfx}G${currentRow}`
              )
            }

            // Add 180 to second LF column (column I)
            spreadsheet.updateCell({ value: 180 }, `${pfx}I${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: '#E2EFDA'
              },
              `${pfx}I${currentRow}:N${currentRow}`
            )

            // Add $/1000 formula in column H
            const shimsDollarFormula = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
            spreadsheet.updateCell({ formula: shimsDollarFormula }, `${pfx}H${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white',
                format: '$#,##0.0'
              },
              `${pfx}H${currentRow}`
            )

            // Apply currency format
            try {
              spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}:N${currentRow}`)
            } catch (e) {
              // Fallback already applied in cellFormat
            }

            // Row height already set above based on shims proposal text

            currentRow++ // Move to next row
          })
        }
      })

      // Add Rock anchor scope heading after SOE scope
      spreadsheet.updateCell({ value: 'Rock anchor scope:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'center',
          backgroundColor: '#BDD7EE',
          border: '1px solid #000000'
        },
        `${pfx}B${currentRow}`
      )
      spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
      currentRow++

      // Add Rock anchors subheading
      spreadsheet.updateCell({ value: 'Rock anchors: including applicable washers, steel bearing plates, locking hex nuts as required' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: '#D0CECE',
          textDecoration: 'underline',
          border: '1px solid #000000'
        },
        `${pfx}B${currentRow}`
      )
      spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
      currentRow++

      // Process Rock anchor items from calculation data
      const rockAnchorItems = []
      if (calculationData && calculationData.length > 0) {
        let inRockAnchorSection = false
        calculationData.forEach((row, index) => {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('rock anchor') && bText.endsWith(':')) {
              inRockAnchorSection = true
              return
            }

            if (inRockAnchorSection) {
              if (bText.endsWith(':') && !bText.includes('rock anchor')) {
                inRockAnchorSection = false
                return
              }

              if (bText && !bText.endsWith(':') && row[2] !== undefined) {
                const takeoff = parseFloat(row[2]) || 0
                if (takeoff > 0 || bText.includes('rock anchor')) {
                  // Try to extract free length and bond length from particulars
                  let freeLength = ''
                  let bondLength = ''
                  const freeLengthMatch = bText.match(/free length=([0-9'"\-]+)/i)
                  const bondLengthMatch = bText.match(/bond length=\s*([0-9'"\-]+)/i)
                  if (freeLengthMatch) {
                    freeLength = freeLengthMatch[1].trim()
                  }
                  if (bondLengthMatch) {
                    bondLength = bondLengthMatch[1].trim()
                  }

                  rockAnchorItems.push({
                    rowIndex: index + 2,
                    particulars: colB.trim(),
                    takeoff: takeoff,
                    qty: parseFloat(row[12]) || 0,
                    freeLength: freeLength,
                    bondLength: bondLength,
                    rawRowNumber: index + 2
                  })
                }
              }
            }
          }
        })
      }

      // Process and display Rock anchor items (grouped)
      if (rockAnchorItems.length > 0) {
        // Group rock anchor items by similar characteristics (drill hole, free length, bond length)
        const rockAnchorGroups = new Map()
        rockAnchorItems.forEach((item) => {
          // Extract drill hole size
          let drillHole = '##"Ø'
          const drillMatch = item.particulars.match(/([\d\s½¼¾]+)"?\s*Ø/i)
          if (drillMatch) {
            drillHole = `${drillMatch[1].trim()}"Ø`
          }

          // Create group key based on drill hole, free length, and bond length
          const freeLengthText = item.freeLength || '##'
          const bondLengthText = item.bondLength || '##'
          const groupKey = `${drillHole}|${freeLengthText}|${bondLengthText}`

          if (!rockAnchorGroups.has(groupKey)) {
            rockAnchorGroups.set(groupKey, [])
          }
          rockAnchorGroups.get(groupKey).push(item)
        })

        // Process each group and generate one proposal text per group
        rockAnchorGroups.forEach((group, groupKey) => {
          // Calculate total quantity for the group
          let totalQty = 0
          let totalTakeoff = 0
          let lastRowNumber = 0
          let drillHole = '##"Ø'
          let freeLengthText = '##'
          let bondLengthText = '##'

          group.forEach((item) => {
            totalQty += item.qty || 0
            totalTakeoff += item.takeoff || 0
            lastRowNumber = Math.max(lastRowNumber, item.rawRowNumber || 0)

            // Get values from first item in group
            if (group.indexOf(item) === 0) {
              const drillMatch = item.particulars.match(/([\d\s½¼¾]+)"?\s*Ø/i)
              if (drillMatch) {
                drillHole = `${drillMatch[1].trim()}"Ø`
              }
              freeLengthText = item.freeLength || '##'
              bondLengthText = item.bondLength || '##'
            }
          })

          const totalQtyValue = Math.round(totalQty || totalTakeoff || 0)

          // Get SOE/FO page reference from raw data
          let soePageMain = 'FO-101.00'
          if (rawData && Array.isArray(rawData) && rawData.length > 1) {
            const headers = rawData[0]
            const dataRows = rawData.slice(1)
            const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
            const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

            if (digitizerIdx !== -1 && pageIdx !== -1) {
              for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                const row = dataRows[rowIndex]
                const digitizerItem = row[digitizerIdx]
                if (digitizerItem && typeof digitizerItem === 'string') {
                  const itemText = digitizerItem.toLowerCase()
                  if (itemText.includes('rock anchor')) {
                    const pageValue = row[pageIdx]
                    if (pageValue) {
                      const pageStr = String(pageValue).trim()
                      const soeMatches = pageStr.match(/(SOE|SOESK|FO)-[\d.]+/gi)
                      if (soeMatches && soeMatches.length > 0) {
                        soePageMain = soeMatches[0]
                        break
                      }
                    }
                  }
                }
              }
            }
          }

          // Generate proposal text with placeholders for missing values
          // Format: F&I new (41)no (1.25"Ø thick) threaded bar (or approved equal) (100 Kips tension load capacity & 80 Kips lock off load capacity) 150KSI rock anchors (5-KSI grout infilled) (L=13'-3" + 10'-6" bond length), 4"Ø drill hole as per FO-101.00
          const proposalText = `F&I new (${totalQtyValue})no (##"Ø thick) threaded bar (or approved equal) (## Kips tension load capacity & ## Kips lock off load capacity) ##KSI rock anchors (##-KSI grout infilled) (L=${freeLengthText} + ${bondLengthText} bond length), ${drillHole} drill hole as per ${soePageMain}`

          spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: 'white',
              verticalAlign: 'top'
            },
            `${pfx}B${currentRow}`
          )

          // Calculate and set row height based on content
          const rockAnchorDynamicHeight = calculateRowHeight(proposalText)
          spreadsheet.setRowHeight(rockAnchorDynamicHeight, currentRow, proposalSheetIndex)

          // Add FT (LF) to column C - sum of all items in group
          let totalFT = 0
          group.forEach(item => {
            if (item.rawRowNumber > 0) {
              const calcSheetName = 'Calculations Sheet'
              // We'll sum all FT values from the group
              totalFT += parseFloat(calculationData[item.rawRowNumber - 2]?.[8] || 0) || 0
            }
          })

          // Find sum row for rock anchors (last row number + 1)
          const sumRowIndex = lastRowNumber + 1
          if (sumRowIndex > 0) {
            const calcSheetName = 'Calculations Sheet'
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${sumRowIndex}` }, `${pfx}C${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}C${currentRow}`
            )
          }

          // Add QTY to column G - sum of all items in group
          if (sumRowIndex > 0) {
            const calcSheetName = 'Calculations Sheet'
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}G${currentRow}`
            )
          }

          // Row height already set above based on rock anchor proposal text
          currentRow++
        })
      }

      // Process Rock bolt items from calculation data
      const rockBoltItems = []
      if (calculationData && calculationData.length > 0) {
        let inRockBoltSection = false
        calculationData.forEach((row, index) => {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('rock bolt') && bText.endsWith(':')) {
              inRockBoltSection = true
              return
            }

            if (inRockBoltSection) {
              if (bText.endsWith(':') && !bText.includes('rock bolt')) {
                inRockBoltSection = false
                return
              }

              if (bText && !bText.endsWith(':') && row[2] !== undefined) {
                const takeoff = parseFloat(row[2]) || 0
                if (takeoff > 0 || bText.includes('rock bolt')) {
                  // Extract bond length from particulars: "Rock bolt @ 7'-0" O.C. (Bond length=10'-0")"
                  let bondLength = ''
                  const bondLengthMatch = bText.match(/bond length=([0-9'"\-]+)/i)
                  if (bondLengthMatch) {
                    bondLength = bondLengthMatch[1].trim()
                  }

                  // Extract O.C. spacing (for reference, but not in output format)
                  let ocSpacing = ''
                  const ocMatch = bText.match(/@\s*([0-9'"\-]+)\s*['"]?\s*o\.?c\.?/i)
                  if (ocMatch) {
                    ocSpacing = ocMatch[1].trim()
                  }

                  rockBoltItems.push({
                    rowIndex: index + 2,
                    particulars: colB.trim(),
                    takeoff: takeoff,
                    qty: parseFloat(row[12]) || 0,
                    bondLength: bondLength,
                    ocSpacing: ocSpacing,
                    rawRowNumber: index + 2
                  })
                }
              }
            }
          }
        })
      }

      // Process and display Rock bolt items (grouped)
      if (rockBoltItems.length > 0) {
        // Group rock bolt items by similar characteristics (drill hole, bond length)
        const rockBoltGroups = new Map()
        rockBoltItems.forEach((item) => {
          // Extract drill hole size
          let drillHole = '##"Ø'
          const drillMatch = item.particulars.match(/([\d\s½¼¾]+)"?\s*Ø/i)
          if (drillMatch) {
            drillHole = `${drillMatch[1].trim()}"Ø`
          }

          // Create group key based on drill hole and bond length
          const bondLengthText = item.bondLength || '##'
          const groupKey = `${drillHole}|${bondLengthText}`

          if (!rockBoltGroups.has(groupKey)) {
            rockBoltGroups.set(groupKey, [])
          }
          rockBoltGroups.get(groupKey).push(item)
        })

        // Process each group and generate one proposal text per group
        rockBoltGroups.forEach((group, groupKey) => {
          // Calculate total quantity for the group
          let totalQty = 0
          let totalTakeoff = 0
          let lastRowNumber = 0
          let drillHole = '##"Ø'
          let bondLengthText = '##'

          group.forEach((item) => {
            totalQty += item.qty || 0
            totalTakeoff += item.takeoff || 0
            lastRowNumber = Math.max(lastRowNumber, item.rawRowNumber || 0)

            // Get values from first item in group
            if (group.indexOf(item) === 0) {
              const drillMatch = item.particulars.match(/([\d\s½¼¾]+)"?\s*Ø/i)
              if (drillMatch) {
                drillHole = `${drillMatch[1].trim()}"Ø`
              }
              bondLengthText = item.bondLength || '##'
            }
          })

          const totalQtyValue = Math.round(totalQty || totalTakeoff || 0)

          // Get SOE page reference from raw data
          let soePageMain = 'SOE-A-100.00'
          if (rawData && Array.isArray(rawData) && rawData.length > 1) {
            const headers = rawData[0]
            const dataRows = rawData.slice(1)
            const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
            const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

            if (digitizerIdx !== -1 && pageIdx !== -1) {
              for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                const row = dataRows[rowIndex]
                const digitizerItem = row[digitizerIdx]
                if (digitizerItem && typeof digitizerItem === 'string') {
                  const itemText = digitizerItem.toLowerCase()
                  if (itemText.includes('rock bolt')) {
                    const pageValue = row[pageIdx]
                    if (pageValue) {
                      const pageStr = String(pageValue).trim()
                      const soeMatches = pageStr.match(/(SOE|SOESK|FO|SOE-A)-[\d.]+/gi)
                      if (soeMatches && soeMatches.length > 0) {
                        soePageMain = soeMatches[0]
                        break
                      }
                    }
                  }
                }
              }
            }
          }

          // Generate proposal text
          // Format: F&I new (11)no threaded bar (or approved equal) 150KSI (20°) rock bolts/dowel (L=10'-0"), 3"Ø drill hole as per SOE-A-100.00
          const proposalText = `F&I new (${totalQtyValue})no threaded bar (or approved equal) ##KSI (##°) rock bolts/dowel (L=${bondLengthText}), ${drillHole} drill hole as per ${soePageMain}`

          spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: 'white',
              verticalAlign: 'top'
            },
            `${pfx}B${currentRow}`
          )

          // Calculate and set row height based on content
          const rockBoltDynamicHeight = calculateRowHeight(proposalText)
          spreadsheet.setRowHeight(rockBoltDynamicHeight, currentRow, proposalSheetIndex)

          // Find sum row for rock bolts (last row number + 1)
          const sumRowIndex = lastRowNumber + 1
          if (sumRowIndex > 0) {
            const calcSheetName = 'Calculations Sheet'
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${sumRowIndex}` }, `${pfx}C${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}C${currentRow}`
            )
          }

          // Add QTY to column G - sum of all items in group
          if (sumRowIndex > 0) {
            const calcSheetName = 'Calculations Sheet'
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}G${currentRow}`
            )
          }

          // Row height already set above based on rock bolt proposal text
          currentRow++
        })
      }

      // Add empty row to separate Rock anchor scope from Foundation scope
      currentRow++

      // Note: Green background for columns I-N will be applied at the end after all data is written

      // Add Foundation subsection headers (similar to SOE subsections)
      const foundationSubsectionItems = window.foundationSubsectionItems || new Map()
      const foundationSubsectionOrder = [
        'Drilled foundation pile',  // Will display as "Foundation drilled piles scope:"
        'Driven foundation pile',   // Will display as "Foundation driven piles scope:"
        'Helical foundation pile',  // Will display as "Foundation helical piles scope:"
        'Stelcor drilled displacement pile', // Will display as "Stelcor piles scope:"
        'CFA pile'                  // Will display as "CAF piles scope:"
      ]

      // Get all unique foundation subsection names from collected items
      // Map new format names to old format for display
      const foundationSubsectionDisplayMap = {
        'Drilled foundation pile': 'Foundation drilled piles',
        'Driven foundation pile': 'Foundation driven piles',
        'Helical foundation pile': 'Foundation helical piles',
        'Stelcor drilled displacement pile': 'Stelcor piles',
        'CFA pile': 'CAF piles'
      }

      const collectedFoundationSubsections = new Set()
      foundationSubsectionItems.forEach((groups, name) => {
        if (groups.length > 0) {
          collectedFoundationSubsections.add(name)
        }
      })

      // Display foundation subsections in order (only up to CFA pile)
      const foundationSubsectionsToDisplay = []
      foundationSubsectionOrder.forEach(name => {
        if (collectedFoundationSubsections.has(name)) {
          foundationSubsectionsToDisplay.push(name)
          collectedFoundationSubsections.delete(name)
        }
      })
      // Do not add any remaining subsections - only show up to CFA pile

      foundationSubsectionsToDisplay.forEach((subsectionName) => {
        // Map new format names to old format for display
        let displayName = subsectionName
        if (foundationSubsectionDisplayMap[subsectionName]) {
          displayName = foundationSubsectionDisplayMap[subsectionName]
        }

        // Add subsection header with blue background and "scope:" suffix
        const subsectionText = `${displayName} scope:`

        spreadsheet.updateCell({ value: subsectionText }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'center',
            backgroundColor: '#BDD7EE',
            border: '1px solid #000000'
          },
          `${pfx}B${currentRow}`
        )

        // Increase row height
        spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)

        currentRow++ // Move to next row

        // Track the start row for this subsection (for total calculation)
        // This is the first row where proposal text will be written
        const subsectionStartRow = currentRow

        // Calculate QTY sum from processor groups (sum of all takeoff values) - BEFORE processing groups
        // Use the summed takeoff from processDrilledFoundationPileItems (or other processors)
        let qtySumValue = 0
        let pileType = 'drilled' // default
        let diameter = null
        let thickness = null

        // Map subsection names to processor group arrays
        const processorGroupsMap = {
          'Drilled foundation pile': window.drilledFoundationPileGroups || [],
          'Foundation drilled piles': window.drilledFoundationPileGroups || [],
          'Driven foundation pile': window.drivenFoundationPileItems || [],
          'Foundation driven piles': window.drivenFoundationPileItems || [],
          'Helical foundation pile': window.helicalFoundationPileGroups || [],
          'Foundation helical piles': window.helicalFoundationPileGroups || [],
          'Stelcor drilled displacement pile': window.stelcorDrilledDisplacementPileItems || [],
          'CFA pile': window.cfaPileItems || []
        }

        // Get the processor groups for this subsection
        const processorGroups = processorGroupsMap[subsectionName] || []

        // Set pileType based on subsection
        if (subsectionName === 'Drilled foundation pile' || subsectionName === 'Foundation drilled piles') {
          pileType = 'drilled'
        } else if (subsectionName === 'Driven foundation pile' || subsectionName === 'Foundation driven piles') {
          pileType = 'driven'
        } else if (subsectionName === 'Helical foundation pile' || subsectionName === 'Foundation helical piles') {
          pileType = 'helical'
        } else if (subsectionName === 'Stelcor drilled displacement pile') {
          pileType = 'stelcor'
        } else if (subsectionName === 'CFA pile') {
          pileType = 'cfa'
        }

        // Process each group individually (for drilled piles, generate separate proposal text for each group)
        // For other types, still sum all groups
        const isDrilledPiles = subsectionName === 'Drilled foundation pile' || subsectionName === 'Foundation drilled piles'

        if (!isDrilledPiles) {
          // For non-drilled piles, sum all groups (existing logic)
          processorGroups.forEach(itemOrGroup => {
            // Check if it's a grouped structure (has items array) or flat item structure
            if (itemOrGroup.items && Array.isArray(itemOrGroup.items) && itemOrGroup.items.length > 0) {
              // For helical piles: need to sum all items in the group
              const groupTakeoff = itemOrGroup.items.reduce((sum, item) => sum + (item.takeoff || 0), 0)
              qtySumValue += groupTakeoff
            } else if (itemOrGroup.takeoff !== undefined) {
              // For flat array items (like driven piles, stelcor, CFA), use takeoff directly
              qtySumValue += itemOrGroup.takeoff || 0
            }
          })

          // Round the final sum
          qtySumValue = Math.round(qtySumValue)
        }

        // Log the sum for debugging


        // Generate proposal text - for drilled piles, generate per group; for others, generate once
        if (isDrilledPiles && processorGroups.length > 0) {
          // Process each group individually for drilled piles
          processorGroups.forEach((itemOrGroup, groupIndex) => {
            if (!itemOrGroup.items || !Array.isArray(itemOrGroup.items) || itemOrGroup.items.length === 0) return

            const groupTakeoff = itemOrGroup.items.reduce((sum, item) => sum + (item.takeoff || 0), 0)
            if (groupTakeoff === 0) return

            const groupQty = Math.round(groupTakeoff)
            const groupHasInflu = itemOrGroup.hasInflu || false


            // Extract data for this specific group
            const firstItem = itemOrGroup.items[0]
            let groupDiameter = null
            let groupThickness = null

            if (firstItem?.particulars) {
              const particulars = firstItem.particulars
              const dtMatch = particulars.match(/([\d.]+)"\s*Ø\s*x\s*([\d.]+)"/i)
              if (dtMatch) {
                groupDiameter = parseFloat(dtMatch[1])
                groupThickness = parseFloat(dtMatch[2])
              }
            }

            // Calculate height and rock socket for this group
            let groupAvgHeight = 0
            let groupHeightCount = 0
            let groupAvgRockSocket = 0
            let groupRockSocketCount = 0

            itemOrGroup.items.forEach(item => {
              if (item.parsed?.calculatedHeight) {
                const heightValue = parseFloat(item.parsed.calculatedHeight)
                if (!isNaN(heightValue)) {
                  const itemTakeoff = item.takeoff || 1
                  groupAvgHeight += heightValue * itemTakeoff
                  groupHeightCount += itemTakeoff
                }
              }

              if (item.particulars) {
                const particulars = String(item.particulars)
                const rockSocketMatch = particulars.match(/([\d.]+)['"]?-?([\d.]+)?["']?\s*(?:rock\s*socket|rock)/i)
                if (rockSocketMatch) {
                  let rockSocketFeet = parseFloat(rockSocketMatch[1])
                  if (rockSocketMatch[2]) {
                    rockSocketFeet += parseFloat(rockSocketMatch[2]) / 12
                  }
                  const itemTakeoff = item.takeoff || 1
                  groupAvgRockSocket += rockSocketFeet * itemTakeoff
                  groupRockSocketCount += itemTakeoff
                }
              }
            })

            if (groupHeightCount > 0) {
              groupAvgHeight = groupAvgHeight / groupHeightCount
            }
            if (groupRockSocketCount > 0) {
              groupAvgRockSocket = groupAvgRockSocket / groupRockSocketCount
            }

            // Round to nearest multiple of 5
            const roundToMultipleOf5 = (value) => Math.ceil(value / 5) * 5
            const avgHeightRounded = roundToMultipleOf5(groupAvgHeight)
            const avgRockSocketRounded = roundToMultipleOf5(groupAvgRockSocket)

            // Format height
            const heightFeet = Math.floor(avgHeightRounded)
            const heightInches = Math.round((avgHeightRounded - heightFeet) * 12)
            let heightText = heightInches === 0 ? `${heightFeet}'-0"` : `${heightFeet}'-${heightInches}"`

            // Format rock socket
            const rockSocketFeet = Math.floor(avgRockSocketRounded)
            const rockSocketInches = Math.round((avgRockSocketRounded - rockSocketFeet) * 12)
            let rockSocketText = groupRockSocketCount > 0
              ? (rockSocketInches === 0 ? `${rockSocketFeet}'-0"` : `${rockSocketFeet}'-${rockSocketInches}"`)
              : "0'-0\""

            // Extract reference codes (same for all groups)
            let foReference = 'FO-101.00'
            let soeReference = 'SOE-101.00'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (const row of dataRows) {
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && typeof digitizerItem === 'string') {
                    const itemText = digitizerItem.toLowerCase()
                    if (itemText.includes('drilled') && (itemText.includes('foundation') || itemText.includes('pile'))) {
                      const pageValue = row[pageIdx]
                      if (pageValue) {
                        const pageStr = String(pageValue).trim()
                        const foMatch = pageStr.match(/FO-[\d.]+/i)
                        if (foMatch && !foReference) foReference = foMatch[0]
                        const soeMatch = pageStr.match(/SOE-[\d.]+/i)
                        if (soeMatch && !soeReference) soeReference = soeMatch[0]
                      }
                    }
                  }
                }
              }
            }

            // Format diameter and thickness
            let diameterThicknessText = ''
            if (groupDiameter && groupThickness) {
              const formatDiameter = (d) => {
                if (d % 1 === 0) return `${d}`
                const fractionMap = {
                  0.125: '1/8', 0.25: '1/4', 0.375: '3/8', 0.5: '1/2',
                  0.625: '5/8', 0.75: '3/4', 0.875: '7/8'
                }
                const whole = Math.floor(d)
                const decimal = d - whole
                const fraction = fractionMap[Math.round(decimal * 1000) / 1000]
                if (fraction) {
                  return whole > 0 ? `${whole}-${fraction}` : fraction
                }
                return d.toString()
              }
              const formattedDiam = formatDiameter(groupDiameter)
              diameterThicknessText = `(${formattedDiam}" ØX${groupThickness}" thick)`
            }

            // If group has Influ, add empty row with background color BEFORE the proposal text
            if (groupHasInflu) {
              // Add empty row with background color #FCE4D6
              spreadsheet.cellFormat(
                { backgroundColor: '#FCE4D6' },
                `${pfx}B${currentRow}:${pfx}H${currentRow}`
              )
              spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
              currentRow++
            }

            // Generate proposal text for this group
            let proposalText = `F&I new (${groupQty})no drilled cassion pile`
            proposalText += ` (# tons design compression, # tons design tension & # ton design lateral load)`
            if (diameterThicknessText) {
              proposalText += `, ${diameterThicknessText}`
            }
            proposalText += `, (#-KSI grout infilled)`
            proposalText += ` with (#)qty #00-#" ## Ksi full length reinforcement`
            proposalText += ` (H=${heightText}, ${rockSocketText} rock socket)`
            proposalText += ` as per ${foReference}, ${soeReference}`

            // Add proposal text
            spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)

            // Apply color if group has Influ
            const cellFormat = {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              verticalAlign: 'top'
            }
            if (groupHasInflu) {
              cellFormat.backgroundColor = '#FCE4D6'
            } else {
              cellFormat.backgroundColor = 'white'
            }

            spreadsheet.cellFormat(cellFormat, `${pfx}B${currentRow}`)

            // Add other columns (FT, LBS, QTY, etc.)
            const calcSheetName = 'Calculations Sheet'

            // Find the sum row for this group in the calculation sheet
            // The sum row contains the FT sum formula for this group
            let groupSumRowIndex = 0

            // Find the formulaData entry for this group
            // Groups are processed in the same order in both generateCalculationSheet and here
            // So we can match by groupIndex, but we need to ensure we're getting the right one
            if (formulaData && Array.isArray(formulaData)) {
              // Get all foundation_sum formulas for drilled foundation pile, sorted by row
              const drilledFoundationSumFormulas = formulaData
                .filter(f =>
                  f.itemType === 'foundation_sum' &&
                  f.subsectionName === 'Drilled foundation pile'
                )
                .sort((a, b) => a.firstDataRow - b.firstDataRow) // Sort by firstDataRow to match group order

              // Use groupIndex to get the corresponding sum row
              if (drilledFoundationSumFormulas[groupIndex]) {
                const formulaInfo = drilledFoundationSumFormulas[groupIndex]
                groupSumRowIndex = formulaInfo.row


                // Reference the FT sum from column I of the sum row
                spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${groupSumRowIndex}` }, `${pfx}C${currentRow}`)
              }
            }

            // Format FT column
            if (groupSumRowIndex > 0 || groupFTSum > 0) {
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: groupHasInflu ? '#FCE4D6' : 'white'
                },
                `${pfx}C${currentRow}`
              )
            }

            // Add LBS to column E - reference to calculation sheet sum row (column K)
            if (groupSumRowIndex > 0) {
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!K${groupSumRowIndex}` }, `${pfx}E${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: groupHasInflu ? '#FCE4D6' : 'white'
                },
                `${pfx}E${currentRow}`
              )
            }

            // Add QTY
            spreadsheet.updateCell({ value: groupQty }, `${pfx}G${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: groupHasInflu ? '#FCE4D6' : 'white'
              },
              `${pfx}G${currentRow}`
            )

            // Add hardcoded second LF value in column I for drilled piles: 150
            spreadsheet.updateCell({ value: 150 }, `${pfx}I${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: groupHasInflu ? '#FCE4D6' : 'white'
              },
              `${pfx}I${currentRow}`
            )

            // Add $/1000 formula
            const dollarFormula = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
            spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: groupHasInflu ? '#FCE4D6' : 'white',
                format: '$#,##0.0'
              },
              `${pfx}H${currentRow}`
            )

            // Row height already set above based on proposal text
            currentRow++
          })
        } else if (qtySumValue > 0) {
          // For non-drilled piles, generate single proposal text (existing logic)
          // Get height and other details from groups if available
          const groups = foundationSubsectionItems.get(subsectionName) || []
          let avgHeight = 0
          let heightCount = 0
          let lastRowNumber = 0
          let sumRowIndex = 0
          const calcSheetName = 'Calculations Sheet'

          // Calculate averages from processor groups first
          let avgRockSocket = 0
          let rockSocketCount = 0

          // Extract height and rock socket from processor groups
          processorGroups.forEach(itemOrGroup => {
            if (itemOrGroup.items && Array.isArray(itemOrGroup.items)) {
              itemOrGroup.items.forEach(item => {
                // Extract height from parsed data or particulars
                if (item.parsed?.calculatedHeight) {
                  const heightValue = parseFloat(item.parsed.calculatedHeight)
                  if (!isNaN(heightValue)) {
                    const itemTakeoff = item.takeoff || 1
                    avgHeight += heightValue * itemTakeoff
                    heightCount += itemTakeoff
                  }
                }

                // Extract rock socket from particulars
                if (item.particulars) {
                  const particulars = String(item.particulars)
                  // Look for rock socket pattern like "7'-0" rock socket" or "7' rock socket" or "4'-0" rock socket"
                  const rockSocketMatch = particulars.match(/([\d.]+)['"]?-?([\d.]+)?["']?\s*(?:rock\s*socket|rock)/i)
                  if (rockSocketMatch) {
                    let rockSocketFeet = parseFloat(rockSocketMatch[1])
                    if (rockSocketMatch[2]) {
                      rockSocketFeet += parseFloat(rockSocketMatch[2]) / 12 // Add inches
                    }
                    const itemTakeoff = item.takeoff || 1
                    avgRockSocket += rockSocketFeet * itemTakeoff
                    rockSocketCount += itemTakeoff
                  }
                }
              })
            }
          })

          groups.forEach((group) => {
            if (group.length === 0) return
            group.forEach(item => {
              if (item.height > 0) {
                avgHeight += item.height * (item.takeoff || 1)
                heightCount += item.takeoff || 1
              }
              // Extract rock socket from particulars or height column
              if (item.particulars) {
                const particulars = String(item.particulars)
                // Look for rock socket pattern like "7'-0" rock socket" or "7' rock socket"
                const rockSocketMatch = particulars.match(/([\d.]+)['"]?-?([\d.]+)?["']?\s*(?:rock\s*socket|rock)/i)
                if (rockSocketMatch) {
                  let rockSocketFeet = parseFloat(rockSocketMatch[1])
                  if (rockSocketMatch[2]) {
                    rockSocketFeet += parseFloat(rockSocketMatch[2]) / 12 // Add inches
                  }
                  avgRockSocket += rockSocketFeet * (item.takeoff || 1)
                  rockSocketCount += item.takeoff || 1
                }
              }
              lastRowNumber = Math.max(lastRowNumber, item.rawRowNumber || 0)
            })
          })

          if (heightCount > 0) {
            avgHeight = avgHeight / heightCount
          }

          if (rockSocketCount > 0) {
            avgRockSocket = avgRockSocket / rockSocketCount
          }

          // Round height to nearest multiple of 5
          const roundToMultipleOf5 = (value) => {
            return Math.ceil(value / 5) * 5
          }
          const avgHeightRounded = roundToMultipleOf5(avgHeight)
          const avgRockSocketRounded = roundToMultipleOf5(avgRockSocket)

          // Format height
          const heightFeet = Math.floor(avgHeightRounded)
          const heightInches = Math.round((avgHeightRounded - heightFeet) * 12)
          let heightText = ''
          if (heightInches === 0) {
            heightText = `${heightFeet}'-0"`
          } else {
            heightText = `${heightFeet}'-${heightInches}"`
          }

          // Format rock socket
          const rockSocketFeet = Math.floor(avgRockSocketRounded)
          const rockSocketInches = Math.round((avgRockSocketRounded - rockSocketFeet) * 12)
          let rockSocketText = "0'-0\""
          if (rockSocketCount > 0) {
            if (rockSocketInches === 0) {
              rockSocketText = `${rockSocketFeet}'-0"`
            } else {
              rockSocketText = `${rockSocketFeet}'-${rockSocketInches}"`
            }
          }

          // Find the sum row for this subsection from formulaData
          // The sum row contains the FT sum formula for this subsection
          if (formulaData && Array.isArray(formulaData)) {
            // Find the foundation_sum formula for this subsection
            const subsectionSumFormula = formulaData.find(f =>
              f.itemType === 'foundation_sum' &&
              f.subsectionName === subsectionName
            )

            if (subsectionSumFormula) {
              sumRowIndex = subsectionSumFormula.row
            } else {
              // Fallback: use lastRowNumber + 1
              sumRowIndex = lastRowNumber + 1
            }
          } else {
            // Fallback: use lastRowNumber + 1
            sumRowIndex = lastRowNumber + 1
          }

          // Extract reference codes (FO-101.00, SOE-101.00) from raw data
          let foReference = ''
          let soeReference = ''

          if (rawData && Array.isArray(rawData) && rawData.length > 1) {
            const headers = rawData[0]
            const dataRows = rawData.slice(1)
            const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
            const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')

            if (digitizerIdx !== -1 && pageIdx !== -1) {
              // Search for foundation drilled pile items
              for (const row of dataRows) {
                const digitizerItem = row[digitizerIdx]
                if (digitizerItem && typeof digitizerItem === 'string') {
                  const itemText = digitizerItem.toLowerCase()
                  // Check if it's a drilled foundation pile
                  if (itemText.includes('drilled') && (itemText.includes('foundation') || itemText.includes('pile'))) {
                    const pageValue = row[pageIdx]
                    if (pageValue) {
                      const pageStr = String(pageValue).trim()
                      // Extract FO reference
                      const foMatch = pageStr.match(/FO-[\d.]+/i)
                      if (foMatch && !foReference) {
                        foReference = foMatch[0]
                      }
                      // Extract SOE reference
                      const soeMatch = pageStr.match(/SOE-[\d.]+/i)
                      if (soeMatch && !soeReference) {
                        soeReference = soeMatch[0]
                      }
                    }
                  }
                }
              }
            }
          }

          // Default references if not found
          if (!foReference) foReference = 'FO-101.00'
          if (!soeReference) soeReference = 'SOE-101.00'

          // Format diameter and thickness for display (e.g., "9-5/8" ØX0.545" thick")
          let diameterThicknessText = ''
          if (diameter && thickness) {
            // Convert decimal diameter to fraction format if needed (e.g., 9.625 -> 9-5/8)
            const formatDiameter = (d) => {
              if (d % 1 === 0) return `${d}`
              // Check common fractions
              const fractionMap = {
                0.125: '1/8', 0.25: '1/4', 0.375: '3/8', 0.5: '1/2',
                0.625: '5/8', 0.75: '3/4', 0.875: '7/8'
              }
              const whole = Math.floor(d)
              const decimal = d - whole
              const fraction = fractionMap[Math.round(decimal * 1000) / 1000]
              if (fraction) {
                return whole > 0 ? `${whole}-${fraction}` : fraction
              }
              return d.toString()
            }
            const formattedDiam = formatDiameter(diameter)
            diameterThicknessText = `(${formattedDiam}" ØX${thickness}" thick)`
          }

          // Generate proposal text
          // Format: F&I new (8)no drilled cassion pile (140 tons design compression, 70 tons design tension & 1 ton design lateral load), (9-5/8" ØX0.545" thick), (6-KSI grout infilled) with (1)qty #24" 80 Ksi full length reinforcement (H=40'-0", 7'-0" rock socket) as per FO-101.00, SOE-101.00
          let pileTypeText = pileType === 'drilled' ? 'drilled cassion pile' : `${pileType} mini-piles`
          let proposalText = `F&I new (${qtySumValue})no ${pileTypeText}`

          // Add capacity (keep as placeholders # with "design" prefix)
          proposalText += ` (# tons design compression, # tons design tension & # ton design lateral load)`

          // Add steel casing if diameter and thickness available
          if (diameterThicknessText) {
            proposalText += `, ${diameterThicknessText}`
          }

          // Add grout (keep 6 as #)
          proposalText += `, (#-KSI grout infilled)`

          // Add reinforcement (placeholder)
          proposalText += ` with (#)qty #00-#" ## Ksi full length reinforcement`

          // Add height and rock socket (comma separated, not +)
          proposalText += ` (H=${heightText}, ${rockSocketText} rock socket)`

          // Add reference codes
          proposalText += ` as per ${foReference}, ${soeReference}`

          // Add proposal text
          spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: 'white',
              verticalAlign: 'top'
            },
            `${pfx}B${currentRow}`
          )

          // Calculate and set row height based on content
          const foundationPileDynamicHeight = calculateRowHeight(proposalText)
          spreadsheet.setRowHeight(foundationPileDynamicHeight, currentRow, proposalSheetIndex)

          // Add FT (LF) to column C - reference to calculation sheet sum row if available
          // Column I (index 8) contains FT for foundation items
          if (sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${sumRowIndex}` }, `${pfx}C${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}C${currentRow}`
            )
          }

          // Add LBS to column E - reference to calculation sheet sum row if available
          if (sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!K${sumRowIndex}` }, `${pfx}E${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}E${currentRow}`
            )
          }

          // Add QTY to column G - use the calculated sum value
          spreadsheet.updateCell({ value: qtySumValue }, `${pfx}G${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: 'white'
            },
            `${pfx}G${currentRow}`
          )

          // Add hardcoded second LF value in column I based on subsection
          const secondLFValues = {
            'Drilled foundation pile': 150,
            'Driven foundation pile': 115,
            'Helical foundation pile': 120,
            'Stelcor drilled displacement pile': 90,
            'CFA pile': 90
          }
          const secondLFValue = secondLFValues[subsectionName] || 0
          if (secondLFValue > 0) {
            spreadsheet.updateCell({ value: secondLFValue }, `${pfx}I${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}I${currentRow}`
            )
          }

          // Add $/1000 formula in column H
          const dollarFormula = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
          spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: 'white',
              format: '$#,##0.0'
            },
            `${pfx}H${currentRow}`
          )

          // Row height already set above based on foundation pile proposal text

          currentRow++ // Move to next row
        }

        // Add misc. section for this foundation subsection (after all proposal text items)
        const miscSections = {
          'Drilled foundation pile': {
            title: 'Drilled pile misc.:',
            included: [
              'Plates & locking nuts included',
              'Pilings will be threaded at both ends and installed in 5\' or 10\' increments',
              'Single mobilization & demobilization of drilling equipment included',
              'Surveying, stakeout, pile numbering plan & as-built plan included',
              'Two lateral reactionary load tests included',
              'Video inspections included',
              'Engineering & Shop drawings included'
            ],
            additional: [
              'Additional linear foot of pile: $155/LF additional if required',
              'Reactionary load test (using production piles as reactionary piles): $15,000 additional if required',
              'Obstruction removal drilling per hour: $995/hour additional if required'
            ]
          },
          'Driven foundation pile': {
            title: 'Driven piles misc.:',
            included: [
              'The Entire Length of each driven pile is charged including any cut off length',
              'Single mobilization & demobilization of drilling equipment included',
              'Surveying, stakeout, pile numbering plan & as-built plan included',
              'Single compression reactionary load tests included',
              'Single lateral reactionary load tests included',
              'Engineering & Shop drawings included'
            ],
            additional: [
              'Additional linear foot of pile: $120/LF additional if required'
            ]
          },
          'Helical foundation pile': {
            title: 'Helical pile misc.:',
            included: [
              'Upon completion: cut-off of helical pile tops after grade marks & capping with and applicable flat plate',
              'If torque is not adequate upon reaching the 7\' depth, additional lengths will be added & billed',
              'Pile location, stake out, deviation, pile numbering plans, grade marks for cut-offs included'
            ],
            additional: []
          },
          'Stelcor drilled displacement pile': {
            title: 'Stelcor pile misc.:',
            included: [
              'Plates & locking nuts included',
              'Single mobilization & demobilization of drilling equipment included',
              'Surveying, stakeout, pile numbering plan & as-built plan included',
              'Single compression reactionary load test included',
              'Engineering & Shop drawings included'
            ],
            additional: [
              'Additional linear foot of pile: $95/LF additional if required',
              'Reactionary lateral load test (using production piles as reactionary piles): $6,000 additional if required',
              'Obstruction removal drilling per hour: $995/hour additional if required'
            ]
          },
          'CFA pile': {
            title: 'CFA pile misc.:',
            included: [
              'Plates & locking nuts included',
              'Single mobilization & demobilization of drilling equipment included',
              'Surveying, stakeout, pile numbering plan & as-built plan included',
              'Single compression reactionary load test included',
              'Engineering & Shop drawings included'
            ],
            additional: [
              'Additional linear foot of pile: $95/LF additional if required',
              'Reactionary lateral load test (using production piles as reactionary piles): $6,000 additional if required',
              'Obstruction removal drilling per hour: $995/hour additional if required'
            ]
          }
        }

        // Check if we have a misc section for this subsection
        const miscSection = miscSections[subsectionName]
        if (miscSection) {
          // Add misc. title
          spreadsheet.updateCell({ value: miscSection.title }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: 'white',
              textDecoration: 'underline'
            },
            `${pfx}B${currentRow}`
          )
          spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
          currentRow++

          // Add included items
          miscSection.included.forEach(item => {
            spreadsheet.updateCell({ value: item }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat(
              {
                fontWeight: 'normal',
                color: '#000000',
                textAlign: 'left',
                backgroundColor: 'white',
                verticalAlign: 'top'
              },
              `${pfx}B${currentRow}`
            )
            spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
            currentRow++
          })

          // Add additional cost items
          miscSection.additional.forEach(item => {
            spreadsheet.updateCell({ value: item }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat(
              {
                fontWeight: 'normal',
                color: '#000000',
                textAlign: 'left',
                backgroundColor: 'white',
                verticalAlign: 'top'
              },
              `${pfx}B${currentRow}`
            )
            spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
            currentRow++
          })
        }

        // Add individual total row for this subsection (after misc section)
        // Calculate the end row (current row - 1, before the total row we're about to add)
        const subsectionEndRow = currentRow - 1

        // Only add total if there are proposal rows for this subsection
        if (subsectionEndRow >= subsectionStartRow) {
          // Determine the total label based on subsection
          const totalLabels = {
            'Drilled foundation pile': 'Foundation Piles Total:',
            'Helical foundation pile': 'Helical Pile Total:',
            'Stelcor drilled displacement pile': 'Stelcor Piles Total:',
            'CFA pile': 'CFA Piles Total:',
            'Driven foundation pile': 'Foundation Piles Total:' // Driven piles use same label as drilled
          }
          const totalLabel = totalLabels[subsectionName] || 'Foundation Piles Total:'

          currentRow++
          // Format like Demolition Total: label in D:E, total in F:G
          spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
          spreadsheet.updateCell({ value: totalLabel }, `${pfx}D${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: '#BDD7EE',
              border: '1px solid #000000'
            },
            `${pfx}D${currentRow}:E${currentRow}`
          )

          // Add total formula: =SUM(H{startRow}:H{endRow})*1000 in merged F:G
          spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
          spreadsheet.updateCell({ formula: `=SUM(H${subsectionStartRow}:H${subsectionEndRow})*1000` }, `${pfx}F${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'right',
              backgroundColor: '#BDD7EE',
              border: '1px solid #000000',
              format: '$#,##0.00'
            },
            `${pfx}F${currentRow}:G${currentRow}`
          )

          // Apply currency format using numberFormat
          try {
            spreadsheet.numberFormat('$#,##0.00', `${pfx}F${currentRow}:G${currentRow}`)
          } catch (e) {
            // Fallback to cellFormat if numberFormat doesn't work
            spreadsheet.cellFormat(
              {
                format: '$#,##0.00'
              },
              `${pfx}F${currentRow}:G${currentRow}`
            )
          }

          // Apply background color to entire row B:G (like Demolition Total)
          spreadsheet.cellFormat(
            {
              backgroundColor: '#BDD7EE'
            },
            `${pfx}B${currentRow}:G${currentRow}`
          )

          currentRow++

          // Add empty row after the total (like Demolition Total)
          currentRow++
        }
      })

      // Add empty row after all foundation sections
      currentRow++

      // Add Substructure concrete scope header
      spreadsheet.updateCell({ value: 'Substructure concrete scope:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'center',
          backgroundColor: '#BDD7EE',
          border: '1px solid #000000',
          textDecoration: 'underline'
        },
        `${pfx}B${currentRow}`
      )

      // Increase row height
      spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)

      currentRow++

      // Add empty row after Substructure concrete scope
      currentRow++

      // Add Below grade waterproofing scope header
      spreadsheet.updateCell({ value: 'Below grade waterproofing scope:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'center',
          backgroundColor: '#BDD7EE',
          border: '1px solid #000000',
          textDecoration: 'underline'
        },
        `${pfx}B${currentRow}`
      )

      // Increase row height
      spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)

      currentRow++

      // Add proposal text with present waterproofing items
      const waterproofingItems = window.waterproofingPresentItems || []

      if (waterproofingItems.length > 0) {
        // Map item names to display format
        const itemDisplayMap = {
          'foundation walls': 'foundation walls',
          'retaining wall': 'retaining wall',
          'vehicle barrier wall': 'vehicle barrier wall',
          'concrete liner wall': 'concrete liner wall',
          'stem wall': 'stem wall',
          'grease trap pit wall': 'grease trap pit wall',
          'house trap pit wall': 'house trap pit wall',
          'detention tank wall': 'detention tank wall',
          'elevator pit walls': 'elevator pit walls',
          'duplex sewage ejector pit wall': 'duplex sewage ejector pit wall'
        }

        // Convert to display names
        const displayItems = waterproofingItems.map(item => itemDisplayMap[item] || item)

        // Join with " & " for the last two, and commas for others
        let itemsText = ''
        if (displayItems.length === 1) {
          itemsText = displayItems[0]
        } else if (displayItems.length === 2) {
          itemsText = displayItems.join(' & ')
        } else {
          // Join all but last with commas, then " & " for the last one
          itemsText = displayItems.slice(0, -1).join(', ') + ' & ' + displayItems[displayItems.length - 1]
        }

        const proposalText = `F&I new Precon by WR Meadows waterproofing membrane (vertical only) with Hydro duct protection/drainage board & R15 2"XPS Rigid Insulation @ ${itemsText}`

        spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            backgroundColor: 'white',
            verticalAlign: 'top'
          },
          `${pfx}B${currentRow}`
        )

        // Calculate and set row height based on content
        const waterproofingDynamicHeight = calculateRowHeight(proposalText)
        spreadsheet.setRowHeight(waterproofingDynamicHeight, currentRow, proposalSheetIndex)

        // Find the FT and SQFT totals from waterproofing sum rows (columns I and J)
        let waterproofingFTSumRow = 0
        const calcSheetName = 'Calculations Sheet'

        if (formulaData && Array.isArray(formulaData)) {
          // Find waterproofing sum rows (exterior_side_sum or negative_side_sum)
          const waterproofingSumFormulas = formulaData.filter(f =>
            (f.itemType === 'waterproofing_exterior_side_sum' || f.itemType === 'waterproofing_negative_side_sum') &&
            f.section === 'waterproofing'
          )

          // Use the first sum row found (or combine if needed)
          if (waterproofingSumFormulas.length > 0) {
            // For now, use the first sum row. If there are multiple, we might need to sum them
            waterproofingFTSumRow = waterproofingSumFormulas[0].row

            // Add FT (LF) to column C - reference to calculation sheet sum row (column I)
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${waterproofingFTSumRow}` }, `${pfx}C${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}C${currentRow}`
            )
            // Apply 2 decimal format
            try {
              spreadsheet.numberFormat('#,##0.00', `${pfx}C${currentRow}`)
            } catch (e) {
              spreadsheet.cellFormat({ format: '#,##0.00' }, `${pfx}C${currentRow}`)
            }

            // Add SQFT to column D - reference to calculation sheet sum row (column J)
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${waterproofingFTSumRow}` }, `${pfx}D${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}D${currentRow}`
            )
            // Apply 2 decimal format
            try {
              spreadsheet.numberFormat('#,##0.00', `${pfx}D${currentRow}`)
            } catch (e) {
              spreadsheet.cellFormat({ format: '#,##0.00' }, `${pfx}D${currentRow}`)
            }

            // Add 15 to column J (SF)
            spreadsheet.updateCell({ value: 15 }, `${pfx}J${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}J${currentRow}`
            )

            // Add $/1000 formula in column H
            const dollarFormula = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
            spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white',
                format: '$#,##0.0'
              },
              `${pfx}H${currentRow}`
            )
            // Apply currency format using numberFormat
            try {
              spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}`)
            } catch (e) {
              // Fallback already applied in cellFormat
            }
          }
        }

        // Row height already set above based on waterproofing proposal text
        currentRow++

        // Add negative side proposal text if negative side items are present
        const negativeSideItems = window.waterproofingNegativeSideItems || []
        if (negativeSideItems.length > 0) {
          // Map item names to display format
          const negativeSideDisplayMap = {
            'detention tank wall': 'detention tank wall',
            'elevator pit walls': 'elevator pit walls',
            'detention tank slab': 'detention tank slab',
            'duplex sewage ejector pit wall': 'duplex sewage ejector pit wall',
            'duplex sewage ejector pit slab': 'duplex sewage ejector pit slab',
            'elevator pit slab': 'elevator pit slab'
          }

          // Convert to display names
          const displayNegativeItems = negativeSideItems.map(item => negativeSideDisplayMap[item] || item)

          // Join with " & " for the last two, and commas for others
          let negativeItemsText = ''
          if (displayNegativeItems.length === 1) {
            negativeItemsText = displayNegativeItems[0]
          } else if (displayNegativeItems.length === 2) {
            negativeItemsText = displayNegativeItems.join(' & ')
          } else {
            // Join all but last with commas, then " & " for the last one
            negativeItemsText = displayNegativeItems.slice(0, -1).join(', ') + ' & ' + displayNegativeItems[displayNegativeItems.length - 1]
          }

          const negativeProposalText = `F&I new Aquafin waterproofing @ negative side of the ${negativeItemsText}`

          spreadsheet.updateCell({ value: negativeProposalText }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: 'white',
              verticalAlign: 'top'
            },
            `${pfx}B${currentRow}`
          )

          // Calculate and set row height based on content
          const negativeSideDynamicHeight = calculateRowHeight(negativeProposalText)
          spreadsheet.setRowHeight(negativeSideDynamicHeight, currentRow, proposalSheetIndex)

          // Find the FT and SQFT totals from negative side sum row (columns I and J)
          let negativeSideSumRow = 0

          if (formulaData && Array.isArray(formulaData)) {
            // Find negative side sum row
            const negativeSideSumFormula = formulaData.find(f =>
              f.itemType === 'waterproofing_negative_side_sum' &&
              f.section === 'waterproofing'
            )

            if (negativeSideSumFormula) {
              negativeSideSumRow = negativeSideSumFormula.row

              // Add FT (LF) to column C - reference to calculation sheet sum row (column I)
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${negativeSideSumRow}` }, `${pfx}C${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white'
                },
                `${pfx}C${currentRow}`
              )
              // Apply 2 decimal format
              try {
                spreadsheet.numberFormat('#,##0.00', `${pfx}C${currentRow}`)
              } catch (e) {
                spreadsheet.cellFormat({ format: '#,##0.00' }, `${pfx}C${currentRow}`)
              }

              // Add SQFT to column D - reference to calculation sheet sum row (column J)
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${negativeSideSumRow}` }, `${pfx}D${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white'
                },
                `${pfx}D${currentRow}`
              )
              // Apply 2 decimal format
              try {
                spreadsheet.numberFormat('#,##0.00', `${pfx}D${currentRow}`)
              } catch (e) {
                spreadsheet.cellFormat({ format: '#,##0.00' }, `${pfx}D${currentRow}`)
              }
            }
          }

          // Add 20 to column J (SF) for negative side
          spreadsheet.updateCell({ value: 20 }, `${pfx}J${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: 'white'
            },
            `${pfx}J${currentRow}`
          )

          // Add $/1000 formula in column H
          const negativeDollarFormula = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
          spreadsheet.updateCell({ formula: negativeDollarFormula }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: 'white',
              format: '$#,##0.0'
            },
            `${pfx}H${currentRow}`
          )
          // Apply currency format using numberFormat
          try {
            spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}`)
          } catch (e) {
            // Fallback already applied in cellFormat
          }

          // Row height already set above based on negative side proposal text
          currentRow++
        }
      } else {
        // No items found, show "No details provided"
        spreadsheet.updateCell({ value: 'No details provided' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'normal',
            color: '#000000',
            textAlign: 'left',
            backgroundColor: 'white'
          },
          `${pfx}B${currentRow}`
        )
        spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
        currentRow++
      }

      // Add horizontal insulation proposal text if present (check calculation sheet for horizontal insulation sum)
      // Horizontal insulation items are always created in the calculation sheet, so check formulaData
      let hasHorizontalInsulation = false
      if (formulaData && Array.isArray(formulaData)) {
        const horizontalInsulationSumFormula = formulaData.find(f =>
          f.itemType === 'waterproofing_horizontal_insulation_sum' &&
          f.section === 'waterproofing'
        )
        hasHorizontalInsulation = !!horizontalInsulationSumFormula
      }

      if (hasHorizontalInsulation) {
        const horizontalProposalText = `F&I new (2" thick) XPS rigid insulation @ SOG & grade beam`

        spreadsheet.updateCell({ value: horizontalProposalText }, `${pfx}B${currentRow}`)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            backgroundColor: 'white',
            verticalAlign: 'top'
          },
          `${pfx}B${currentRow}`
        )

        // Calculate and set row height based on content
        const horizontalInsulationDynamicHeight = calculateRowHeight(horizontalProposalText)
        spreadsheet.setRowHeight(horizontalInsulationDynamicHeight, currentRow, proposalSheetIndex)

        // Find the SQFT total from horizontal insulation sum row (column J)
        let horizontalInsulationSumRow = 0
        const calcSheetName = 'Calculations Sheet'

        if (formulaData && Array.isArray(formulaData)) {
          // Find horizontal insulation sum row
          const horizontalInsulationSumFormula = formulaData.find(f =>
            f.itemType === 'waterproofing_horizontal_insulation_sum' &&
            f.section === 'waterproofing'
          )

          if (horizontalInsulationSumFormula) {
            horizontalInsulationSumRow = horizontalInsulationSumFormula.row

            // Add SQFT to column D - reference to calculation sheet sum row (column J)
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${horizontalInsulationSumRow}` }, `${pfx}D${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}D${currentRow}`
            )
            // Apply 2 decimal format
            try {
              spreadsheet.numberFormat('#,##0.00', `${pfx}D${currentRow}`)
            } catch (e) {
              spreadsheet.cellFormat({ format: '#,##0.00' }, `${pfx}D${currentRow}`)
            }
          }
        }

        // Add 15 to column J (SF)
        spreadsheet.updateCell({ value: 15 }, `${pfx}J${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            textAlign: 'right',
            backgroundColor: 'white'
          },
          `${pfx}J${currentRow}`
        )

        // Add $/1000 formula in column H
        const horizontalDollarFormula = `=ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)`
        spreadsheet.updateCell({ formula: horizontalDollarFormula }, `${pfx}H${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            textAlign: 'right',
            backgroundColor: 'white',
            format: '$#,##0.0'
          },
          `${pfx}H${currentRow}`
        )
        // Apply currency format using numberFormat
        try {
          spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}`)
        } catch (e) {
          // Fallback already applied in cellFormat
        }

        // Row height already set above based on horizontal insulation proposal text
        currentRow++
      }

      // Add note about WR Meadows Installer
      spreadsheet.updateCell({ value: 'Note: Capstone Contracting Corp is a licensed WR Meadows Installer ' }, `${pfx}B${currentRow}`)
      spreadsheet.wrap(`${pfx}B${currentRow}`, true)
      spreadsheet.cellFormat(
        {
          fontWeight: 'normal',
          color: '#000000',
          textAlign: 'center',
          backgroundColor: 'white',
          verticalAlign: 'middle'
        },
        `${pfx}B${currentRow}`
      )
      spreadsheet.setRowHeight(currentRow, 22, proposalSheetIndex)
      currentRow++

      // Individual totals are now added after each subsection's misc section

      // Apply green background and $ prefix format to columns I-N for all data rows (from row 2 to currentRow-1)
      // Also apply $ prefix format to column H ($/1000)
      // Row 1 is the header, so data starts from row 2
      if (currentRow > 2) {
        const lastDataRow = currentRow - 1
        // Apply green background to columns I-N
        spreadsheet.cellFormat(
          {
            backgroundColor: '#E2EFDA'
          },
          `${pfx}I2:N${lastDataRow}`
        )
        // Apply $ prefix format to column H ($/1000) using numberFormat
        try {
          spreadsheet.numberFormat('$#,##0.0', `${pfx}H2:H${lastDataRow}`)
        } catch (e) {
          // Fallback to cellFormat if numberFormat doesn't work
          spreadsheet.cellFormat(
            {
              format: '$#,##0.0'
            },
            `${pfx}H2:H${lastDataRow}`
          )
        }
        // Apply $ prefix format to columns I-N using numberFormat
        try {
          spreadsheet.numberFormat('$#,##0.0', `${pfx}I2:N${lastDataRow}`)
        } catch (e) {
          // Fallback to cellFormat if numberFormat doesn't work
          spreadsheet.cellFormat(
            {
              format: '$#,##0.0'
            },
            `${pfx}I2:N${lastDataRow}`
          )
        }
      }
    }
  }

  const onCreated = () => {
    if (sheetCreated) return
    if (calculationData.length > 0) {
      applyDataToSpreadsheet()
      setSheetCreated(true)
    }

    // Proposal sheet should stay independent: only seed the first row headers.
    if (spreadsheetRef.current) {
      try {
        applyProposalFirstRow(spreadsheetRef.current)
      } catch (e) {
        // Ignore errors
      }
    }

    // Restore last active sheet (so if you were on Proposal, refresh stays on Proposal)
    try {
      const lastSheet = localStorage.getItem(LAST_ACTIVE_SHEET_KEY)
      if (lastSheet && spreadsheetRef.current) {
        spreadsheetRef.current.goTo(`${lastSheet}!A1`)
      }
    } catch (e) {
      // ignore
    }
  }

  const handleActionComplete = () => {
    // Persist which sheet the user is on, so refresh returns to it.
    try {
      const sheetName = spreadsheetRef.current?.getActiveSheet?.().name
      if (sheetName) localStorage.setItem(LAST_ACTIVE_SHEET_KEY, sheetName)
    } catch (e) {
      // ignore
    }
  }

  const handleGoBack = () => {
    navigate('/dashboard')
  }

  const handleSave = () => {
    // TODO: Implement save to MongoDB
  }

  const columnConfigs = generateColumnConfigs()

  return (
    <div className="h-screen flex flex-col">
      <div className="bg-white border-b shadow-sm">
        <div className="flex items-center justify-between px-4 py-3">
          <div className="flex items-center space-x-4">
            <button
              onClick={handleGoBack}
              className="flex items-center space-x-2 px-3 py-2 text-gray-600 hover:text-gray-900 hover:bg-gray-100 rounded-lg transition-colors"
            >
              <FiArrowLeft />
              <span>Back</span>
            </button>
            <div className="h-6 w-px bg-gray-300"></div>
            <h1 className="text-xl font-bold text-gray-900">
              {hasData && fileName ? fileName : 'Excel Processing Workbook'}
            </h1>
            {hasData && (
              <span className="text-sm text-gray-500">
                ({rows?.length || 0} raw data rows | {calculationData.length} calculation rows)
              </span>
            )}
          </div>
          <button
            onClick={handleSave}
            className="flex items-center space-x-2 px-4 py-2 bg-green-600 text-white rounded-lg hover:bg-green-700 transition-colors"
          >
            <FiSave />
            <span>Save to Database</span>
          </button>
        </div>
      </div>
      <div className="flex-1 overflow-hidden">
        <SpreadsheetComponent
          ref={spreadsheetRef}
          created={onCreated}
          actionComplete={handleActionComplete}
          allowOpen={true}
          allowSave={true}
          enableClipboard={true}
          openUrl="https://services.syncfusion.com/react/production/api/spreadsheet/open"
          saveUrl="https://services.syncfusion.com/react/production/api/spreadsheet/save"
          height="100%"
        >
          <SheetsDirective>
            <SheetDirective name="Calculations Sheet" rowCount={1000}>
              <ColumnsDirective>
                {columnConfigs.map((config, index) => (
                  <ColumnDirective key={index} width={config.width}></ColumnDirective>
                ))}
              </ColumnsDirective>
            </SheetDirective>
            <SheetDirective name="Proposal Sheet" rowCount={500}>
              <ColumnsDirective>
                <ColumnDirective width={150}></ColumnDirective>
              </ColumnsDirective>
            </SheetDirective>
          </SheetsDirective>
        </SpreadsheetComponent>
      </div>
    </div>
  )
}

export default Spreadsheet

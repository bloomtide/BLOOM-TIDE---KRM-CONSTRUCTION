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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)

            // Sum for CY (column L) - sum all items in subsection
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)

            // For Demo isolated footing, also sum column M (QTY)
            if (subsection === 'Demo isolated footing') {
              spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)

            // Sum for 1.3*CY (column L) - sum all items
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)

            // Sum for CY (column L) - sum all backfill items
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)

            // Sum for CY (column L) - sum all mud slab items
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', textDecoration: 'line-through' }, `K${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', textDecoration: 'line-through' }, `K${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', textDecoration: 'line-through' }, `K${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'rock_excavation_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
          } catch (error) {
            // Ignore errors
          }
          return
        }

        formulas = generateRockExcavationFormulas(itemType === 'rock_excavation_item' ? parsedData.itemType : itemType, row, parsedData)
      } else if (section === 'soe') {
        if (['soldier_pile_item', 'soe_generic_item', 'backpacking_item', 'supporting_angle', 'parging', 'heel_block', 'underpinning', 'shims', 'rock_anchor', 'rock_bolt', 'anchor', 'tie_back', 'concrete_soil_retention_pier', 'guide_wall', 'dowel_bar', 'rock_pin', 'shotcrete', 'permission_grouting', 'button', 'rock_stabilization', 'form_board'].includes(itemType)) {
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
              spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            }

            // Special formatting for Shims - columns I and J in red
            if (itemType === 'shims') {
              spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            }
          } catch (error) {
            // Ignore errors
          }
          return
        }

        if (itemType === 'soldier_pile_group_sum' || itemType === 'soe_generic_sum') {
          const { firstDataRow, lastDataRow, subsectionName } = formulaInfo
          try {
            // Standard sum for FT (I) for all SOE, except Heel blocks
            const ftSumSubsections = ['Rock anchors', 'Rock bolts', 'Anchor', 'Tie back', 'Dowel bar', 'Rock pins', 'Shotcrete', 'Permission grouting', 'Form board', 'Guide wall']
            if (subsectionName !== 'Heel blocks' && (ftSumSubsections.includes(subsectionName) || !['Concrete soil retention piers', 'Buttons', 'Rock stabilization'].includes(subsectionName))) {
              spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            }

            // Sum for SQ FT (J)
            const sqFtSubsections = ['Sheet pile', 'Timber lagging', 'Timber sheeting', 'Parging', 'Heel blocks', 'Underpinning', 'Concrete soil retention piers', 'Guide wall', 'Shotcrete', 'Permission grouting', 'Buttons', 'Form board', 'Rock stabilization']
            if (sqFtSubsections.includes(subsectionName)) {
              spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            }

            // Sum for LBS (K)
            const lbsSubsections = [
              'Secondary secant piles',
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
              spreadsheet.cellFormat({ color: '#FF0000' }, `K${row}`)
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
              'Anchor',
              'Tie back',
              'Concrete soil retention piers',
              'Dowel bar',
              'Rock pins',
              'Buttons'
            ]
            if (itemType === 'soldier_pile_group_sum' || qtySubsections.includes(subsectionName)) {
              spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            }

            // Sum for CY (L)
            const cySubsections = ['Heel blocks', 'Underpinning', 'Concrete soil retention piers', 'Guide wall', 'Shotcrete', 'Buttons', 'Rock stabilization']
            if (cySubsections.includes(subsectionName)) {
              spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
          } catch (error) {
            console.error(`Error applying foundation extra EA formula at row ${row}:`, error)
          }
          return
        }
        if (['drilled_foundation_pile', 'helical_foundation_pile', 'driven_foundation_pile', 'stelcor_drilled_displacement_pile', 'cfa_pile', 'pile_cap', 'strip_footing', 'isolated_footing', 'pilaster', 'grade_beam', 'tie_beam', 'thickened_slab', 'buttress_takeoff', 'buttress_final', 'pier', 'corbel', 'linear_wall', 'foundation_wall', 'retaining_wall', 'barrier_wall', 'stem_wall', 'elevator_pit', 'detention_tank', 'duplex_sewage_ejector_pit', 'deep_sewage_ejector_pit', 'grease_trap', 'house_trap', 'mat_slab', 'mud_slab_foundation', 'sog', 'stairs_on_grade', 'electric_conduit'].includes(itemType)) {
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
              spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            }

            // Special formatting for Elevator Pit - Sump pit row (J, L, M red)
            if (
              itemType === 'elevator_pit' &&
              (parsedData || formulaInfo)?.parsed?.itemSubType === 'sump_pit'
            ) {
              spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            }

            // Special formatting for Mud Slab - J and L red
            if (itemType === 'mud_slab_foundation') {
              spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            }

            // Special formatting for Stairs on grade - Stair slab and Landings J and L red
            if (itemType === 'stairs_on_grade') {
              const subType = (parsedData || formulaInfo)?.parsed?.itemSubType
              if (subType === 'stair_slab' || subType === 'landings') {
                spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
                spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
              }
            }
          } catch (error) {
            console.error(`Error applying Foundation formula at row ${row}:`, error)
          }
          return
        }

        if (itemType === 'foundation_sum') {
          const { firstDataRow, lastDataRow, subsectionName, isDualDiameter, excludeISum, excludeJSum, matSumOnly, cySumOnly, firstDataRowForL, lastDataRowForL, lSumRange } = formulaInfo
          try {
            // Sum for FT (I) - exclude if excludeISum is true (for slab items)
            if (!excludeISum) {
              const ftSumSubsections = ['Helical foundation pile', 'Driven foundation pile', 'Stelcor drilled displacement pile', 'CFA pile', 'Grade beams', 'Tie beam', 'Thickened slab', 'Corbel', 'Linear Wall', 'Foundation Wall', 'Retaining walls', 'Barrier wall', 'Drilled foundation pile', 'Strip Footings', 'Stem wall', 'Detention tank', 'Duplex sewage ejector pit', 'Deep sewage ejector pit', 'Grease trap', 'House trap', 'SOG', 'Stairs on grade Stairs', 'Electric conduit']
              if (ftSumSubsections.includes(subsectionName)) {
                spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
                spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
              }
            }

            // Sum for SQ FT (I and J) - for drilled foundation pile dual diameter
            if (subsectionName === 'Drilled foundation pile' && isDualDiameter) {
              spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
              spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            }

            // Sum for SQ FT (J) - exclude Drilled foundation pile (single diameter)
            // For Mat slab, only sum J for mat items (not haunch)
            if (!excludeJSum && !cySumOnly) {
              const sqFtSubsections = ['Pile caps', 'Isolated Footings', 'Pilaster', 'Pier', 'Strip Footings', 'Grade beams', 'Tie beam', 'Thickened slab', 'Corbel', 'Linear Wall', 'Foundation Wall', 'Retaining walls', 'Barrier wall', 'Stem wall', 'Elevator Pit', 'Detention tank', 'Duplex sewage ejector pit', 'Deep sewage ejector pit', 'Grease trap', 'House trap', 'Mat slab', 'SOG', 'Stairs on grade Stairs']
              if (sqFtSubsections.includes(subsectionName)) {
                spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
                spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
              }
            }

            // For Drilled foundation pile single diameter, J should be empty (no sum)
            if (subsectionName === 'Drilled foundation pile' && !isDualDiameter) {
              // J column should be empty - do nothing
            }

            // Sum for LBS (K)
            const lbsSubsections = ['Drilled foundation pile', 'Helical foundation pile', 'Driven foundation pile', 'Stelcor drilled displacement pile']
            if (lbsSubsections.includes(subsectionName)) {
              spreadsheet.updateCell({ formula: `=SUM(K${firstDataRow}:K${lastDataRow})` }, `K${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `K${row}`)
            }

            // Sum for QTY (M)
            const qtySubsections = ['Drilled foundation pile', 'Helical foundation pile', 'Driven foundation pile', 'Stelcor drilled displacement pile', 'CFA pile', 'Pile caps', 'Isolated Footings', 'Pilaster', 'Pier', 'Stairs on grade Stairs']
            if (qtySubsections.includes(subsectionName)) {
              spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            }

            // Sum for CY (L)
            // When lSumRange is provided (Stairs on grade Stairs), use it so landings L is included
            const lStartRow = firstDataRowForL != null ? firstDataRowForL : firstDataRow
            const lEndRow = lastDataRowForL != null ? lastDataRowForL : lastDataRow
            // For Mat slab with cySumOnly, sum L includes both mat and haunch
            if (cySumOnly) {
              // Only sum L (CY) for mat + haunch combined
              spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            } else if (!formulaInfo.excludeLSum) {
              const cySubsections = ['Pile caps', 'Strip Footings', 'Isolated Footings', 'Pilaster', 'Grade beams', 'Tie beam', 'Thickened slab', 'Pier', 'Corbel', 'Linear Wall', 'Foundation Wall', 'Retaining walls', 'Barrier wall', 'Stem wall', 'Elevator Pit', 'Detention tank', 'Duplex sewage ejector pit', 'Deep sewage ejector pit', 'Grease trap', 'House trap', 'Mat slab', 'SOG', 'Stairs on grade Stairs']
              if (cySubsections.includes(subsectionName)) {
                const lFormula = lSumRange ? `=SUM(${lSumRange})` : `=SUM(L${lStartRow}:L${lEndRow})`
                deferredFoundationSumL.push({ row, lFormula })
                spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
          } catch (error) {
            console.error(`Error applying Waterproofing Horizontal WP formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'waterproofing_horizontal_wp_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
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
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
          } catch (error) {
            console.error(`Error applying Waterproofing Horizontal insulation formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'waterproofing_horizontal_insulation_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
          } catch (error) {
            console.error(`Error applying Waterproofing Horizontal insulation sum at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'waterproofing_extra_sqft') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
          } catch (error) {
            console.error(`Error applying Waterproofing extra SQ FT formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'waterproofing_extra_ft') {
          try {
            spreadsheet.updateCell({ formula: `=H${row}*C${row}` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
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

      // Format number columns to show 2 decimal places
      const numColumns = ['E', 'F', 'G', 'H', 'I', 'J', 'K', 'L', 'M']
      numColumns.forEach(col => {
        try {
          spreadsheet.numberFormat('0.00', `${col}:${col}`)
        } catch (e) {
          // Ignore formatting errors for individual columns
        }
      })
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

    // All rows have the same height
    const uniformRowHeight = 22
    for (let r = 0; r < 20; r++) {
      spreadsheet.setRowHeight(uniformRowHeight, r, proposalSheetIndex)
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

    // Logo (image) near center top-left; uses public path
    try {
      const imgSrc = encodeURI('/images/templateimage.png')
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
    spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '1px solid #000000',borderTop: '1px solid #000000',borderBottom: '1px solid #D0CECE' }, `${pfx}F5`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'bold', borderRight: '1px solid #000000',borderTop: '1px solid #000000',borderBottom: '1px solid #D0CECE' }, `${pfx}G5`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal' }, `${pfx}H5`)

    // Row 6: SOE:
    spreadsheet.merge(`${pfx}F6:G6`)
    spreadsheet.updateCell({ value: 'SOE:' }, `${pfx}F6`)
    spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '1px solid #000000',borderTop: '1px solid #000000',borderBottom: '1px solid #D0CECE' }, `${pfx}F6`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'bold', borderRight: '1px solid #000000',borderTop: '1px solid #000000',borderBottom: '1px solid #D0CECE' }, `${pfx}G6`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal' }, `${pfx}H6`)

    // Row 7: Structural:
    spreadsheet.merge(`${pfx}F7:G7`)
    spreadsheet.updateCell({ value: 'Structural:' }, `${pfx}F7`)
    spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '1px solid #000000',borderTop: '1px solid #000000',borderBottom: '1px solid #D0CECE' }, `${pfx}F7`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'bold', borderRight: '1px solid #000000',borderTop: '1px solid #000000',borderBottom: '1px solid #D0CECE' }, `${pfx}G7`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal' }, `${pfx}H7`)

    // Row 8: Architectural:
    spreadsheet.merge(`${pfx}F8:G8`)
    spreadsheet.updateCell({ value: 'Architectural:' }, `${pfx}F8`)
    spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '1px solid #000000',borderTop: '1px solid #000000',borderBottom: '1px solid #D0CECE' }, `${pfx}F8`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'bold', borderRight: '1px solid #000000',borderTop: '1px solid #000000',borderBottom: '1px solid #D0CECE' }, `${pfx}G8`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal' }, `${pfx}H8`)

    // Row 9: Plumbing:
    spreadsheet.merge(`${pfx}F9:G9`)
    spreadsheet.updateCell({ value: 'Plumbing:' }, `${pfx}F9`)
    spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '1px solid #000000',borderTop: '1px solid #000000',borderBottom: '1px solid #D0CECE' }, `${pfx}F9`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'bold', borderRight: '1px solid #000000',borderTop: '1px solid #000000',borderBottom: '1px solid #D0CECE' }, `${pfx}G9`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal' }, `${pfx}H9`)

    // Row 10: Mechanical
    spreadsheet.merge(`${pfx}F10:G10`)
    spreadsheet.updateCell({ value: 'Mechanical' }, `${pfx}F10`)
    spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderBottom: '1px solid #000000', borderLeft: '1px solid #000000' }, `${pfx}F10`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'bold', borderBottom: '1px solid #000000', borderRight: '1px solid #000000' }, `${pfx}G10`)
    spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal' }, `${pfx}H10`)

    // Bottom header row (row 13)
    spreadsheet.updateCell({ value: 'DESCRIPTION' }, `${pfx}B13`)
    ;[
      ['C13', 'LF'], ['D13', 'SF'], ['E13', 'LBS'], ['F13', 'CY'], ['G13', 'QTY'],
      ['H13', '$/1000'],
      ['I13', 'LF'], ['J13', 'SF'], ['K13', 'LBS'], ['L13', 'CY'], ['M13', 'QTY']
    ].forEach(([cell, value]) => spreadsheet.updateCell({ value }, `${pfx}${cell}`))
    spreadsheet.updateCell({ value: 'LS' }, `${pfx}N13`)

    spreadsheet.cellFormat(headerGray, `${pfx}B13:N13`)
    spreadsheet.cellFormat(thick, `${pfx}B13:N13`)
    // Make $/1000 cell background white
    spreadsheet.cellFormat({ backgroundColor: 'white' }, `${pfx}H13`)

    // Row 14: Empty row
    // (No content needed)

    // Row 15: Demolition scope
    spreadsheet.updateCell({ value: 'Demolition scope:' }, `${pfx}B15`)
    spreadsheet.cellFormat({ 
      backgroundColor: '#BDD7EE', 
      textAlign: 'center', 
      verticalAlign: 'middle',
      textDecoration: 'underline',
      fontWeight: 'normal',
      border: '1px solid #000000'
    }, `${pfx}B15`)

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
          return
        }

        if (!inDemolitionSection && !inExcavationSection && !inRockExcavationSection && !inSOESection) return

        // If we hit another main section header in column A, stop collecting.
        if (colA && String(colA).trim() && 
            String(colA).trim().toLowerCase() !== 'demolition' && 
            String(colA).trim().toLowerCase() !== 'excavation' &&
            String(colA).trim().toLowerCase() !== 'rock excavation' &&
            String(colA).trim().toLowerCase() !== 'soe') {
          // Save current SOE subsection items if any
          if (inSOESection && currentSOESubsectionName && currentSOESubsectionItems.length > 0) {
            if (!window.soeSubsectionItems.has(currentSOESubsectionName)) {
              window.soeSubsectionItems.set(currentSOESubsectionName, [])
            }
            window.soeSubsectionItems.get(currentSOESubsectionName).push([...currentSOESubsectionItems])
          }
          
          // Reset flags when moving to next section
          inDemolitionSection = false
          inExcavationSection = false
          inRockExcavationSection = false
          inSOESection = false
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
            
            currentSOESubsectionItems.push({
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
          } catch (e) {}
        }
        if ((!bValue || bValue === '') && workbook) {
          try {
            bValue = workbook.getValueRowCol(calcSheetIndex, rowIndex, 1)
          } catch (e) {}
        }
        if ((!qtyValue || qtyValue === '') && workbook) {
          try {
            qtyValue = workbook.getValueRowCol(calcSheetIndex, rowIndex, 12)
          } catch (e) {}
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

      // Render specific demolition lines into fixed rows B16–B19
      const orderedSubsections = [
        'Demo slab on grade',
        'Demo strip footing',
        'Demo foundation wall',
        'Demo isolated footing'
      ]

      orderedSubsections.forEach((name, index) => {
        const rowIndex = 16 + index // 16,17,18,19
        const originalText = linesBySubsection.get(name)
        const templateText = buildDemolitionTemplate(name, originalText)
        const cellRef = `${pfx}B${rowIndex}`

        // Increase row height to accommodate wrapped text
        spreadsheet.setRowHeight(44, rowIndex - 1, proposalSheetIndex) // rowIndex is 1-based, setRowHeight uses 0-based

        // Always set the template text, even if no original data found
        spreadsheet.updateCell({ value: templateText }, cellRef)
        
        // Enable text wrapping using Syncfusion's wrap method
        try {
          spreadsheet.wrap(cellRef, true)
        } catch (e) {
          // Fallback if wrap method doesn't exist
        }
        
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

        // Fill SF (column J) with 7 for rows 16-19
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

        // Fill CY (column L) with 500 for rows 16-19
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
      // Extend to all rows below demolition scope (rows 16-21)
      for (let row = 16; row <= 21; row++) {
        const columns = ['I', 'J', 'K', 'L', 'M', 'N'] // LF, SF, LBS, CY, QTY, LS
        columns.forEach(col => {
          spreadsheet.cellFormat(
            { backgroundColor: '#E2EFDA' },
            `${pfx}${col}${row}`
          )
        })
      }

      // Add note below all demolition items (row 20)
      spreadsheet.updateCell({ value: 'Note: Site/building demolition by others.' }, `${pfx}B20`)
      spreadsheet.cellFormat(
        { 
          fontWeight: 'bold', 
          color: '#000000', 
          textAlign: 'left',
          backgroundColor: 'white'
        },
        `${pfx}B20`
      )

      // Add Demolition Total row (row 21)
      spreadsheet.merge(`${pfx}D21:E21`)
      spreadsheet.updateCell({ value: 'Demolition Total:' }, `${pfx}D21`)
      spreadsheet.cellFormat(
        { 
          fontWeight: 'bold', 
          color: '#000000', 
          textAlign: 'left',
          backgroundColor: '#BDD7EE',
          border: '1px solid #000000'
        },
        `${pfx}D21:E21`
      )

      spreadsheet.merge(`${pfx}F21:G21`)
      const totalFormula = `=SUM(H16:H19)*1000`
      spreadsheet.updateCell({ formula: totalFormula }, `${pfx}F21`)
      spreadsheet.cellFormat(
        { 
          fontWeight: 'bold', 
          color: '#000000', 
          textAlign: 'right',
          backgroundColor: '#BDD7EE',
          border: '1px solid #000000',
          format: '$#,##0.00'
        },
        `${pfx}F21:G21`
      )

      // Apply background color to entire row B21:G21
      spreadsheet.cellFormat(
        { 
          backgroundColor: '#BDD7EE'
        },
        `${pfx}B21:G21`
      )

      // Add empty row below Demolition Total (row 22)
      // (No content needed - empty row)

      // Row 23: Excavation scope
      spreadsheet.updateCell({ value: 'Excavation scope:' }, `${pfx}B23`)
      spreadsheet.cellFormat({ 
        backgroundColor: '#BDD7EE', 
        textAlign: 'center', 
        verticalAlign: 'middle',
        textDecoration: 'underline',
        fontWeight: 'bold',
        border: '1px solid #000000'
      }, `${pfx}B23`)

      // Row 24: Empty row
      // (No content needed - empty row)

      // Row 25: Soil excavation scope
      spreadsheet.updateCell({ value: 'Soil excavation scope:' }, `${pfx}B25`)
      spreadsheet.cellFormat({ 
        backgroundColor: '#FFF2CC', 
        textAlign: 'center', 
        verticalAlign: 'middle',
        textDecoration: 'underline',
        fontWeight: 'bold',
        border: '1px solid #000000'
      }, `${pfx}B25`)

      // Row 26: First soil excavation line
      spreadsheet.updateCell({ value: 'Allow to perform soil excavation, trucking & disposal (Havg=16\'-9") as per SOE-101.00, P-301.01 & details on SOE-201.01 to SOE-204.00' }, `${pfx}B26`)
      spreadsheet.wrap(`${pfx}B26`, true)
      spreadsheet.cellFormat(
        { 
          fontWeight: 'bold', 
          color: '#000000', 
          textAlign: 'left',
          backgroundColor: 'white',
          verticalAlign: 'top'
        },
        `${pfx}B26`
      )
      
      // Add SF value from excavation total to column D (using formula to reference calculation sheet)
      if (excavationEmptyRowIndex) {
        // Use formula to reference the calculation sheet
        spreadsheet.updateCell({ formula: `='Calculations Sheet'!J${excavationEmptyRowIndex}` }, `${pfx}D26`)
      } else {
        // Fallback to value if row index not found
        const formattedExcavationSF = parseFloat(excavationTotalSQFT.toFixed(2))
        spreadsheet.updateCell({ value: formattedExcavationSF }, `${pfx}D26`)
      }
      spreadsheet.cellFormat(
        { 
          fontWeight: 'bold', 
          color: '#000000', 
          textAlign: 'right',
          format: '#,##0.00'
        },
        `${pfx}D26`
      )
      
      // Add CY value from excavation total to column F (using formula to reference calculation sheet)
      if (excavationEmptyRowIndex) {
        // Use formula to reference the calculation sheet
        spreadsheet.updateCell({ formula: `='Calculations Sheet'!L${excavationEmptyRowIndex}` }, `${pfx}F26`)
      } else {
        // Fallback to value if row index not found
        const formattedExcavationCY = parseFloat(excavationTotalCY.toFixed(2))
        spreadsheet.updateCell({ value: formattedExcavationCY }, `${pfx}F26`)
      }
      spreadsheet.cellFormat(
        { 
          fontWeight: 'bold', 
          color: '#000000', 
          textAlign: 'right',
          format: '#,##0.00'
        },
        `${pfx}F26`
      )

      // Fill CY (column L) with 75 for row 26
      spreadsheet.updateCell({ value: 75 }, `${pfx}L26`)
      spreadsheet.cellFormat(
        { 
          fontWeight: 'bold', 
          color: '#000000', 
          textAlign: 'right',
          format: '$#,##0.00'
        },
        `${pfx}L26`
      )

      // Add $/1000 formula in column H for row 26
      const dollarFormula26 = `=ROUNDUP(MAX(C26*I26,D26*J26,E26*K26,F26*L26,G26*M26,N26)/1000,1)`
      
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
            const row26 = 25 // 0-based index for row 26
            
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
              c26: getCellInfo(row26, 2, 'C'),
              i26: getCellInfo(row26, 8, 'I'),
              d26: getCellInfo(row26, 3, 'D'),
              j26: getCellInfo(row26, 9, 'J'),
              e26: getCellInfo(row26, 4, 'E'),
              k26: getCellInfo(row26, 10, 'K'),
              f26: getCellInfo(row26, 5, 'F'),
              l26: getCellInfo(row26, 11, 'L'),
              g26: getCellInfo(row26, 6, 'G'),
              m26: getCellInfo(row26, 12, 'M'),
              n26: getCellInfo(row26, 13, 'N')
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
          console.error('Row 26: Error reading values:', e)
          console.error('Row 26: Error stack:', e.stack)
        }
      }, 2000) // Increased timeout to 2 seconds
      
      spreadsheet.updateCell({ formula: dollarFormula26 }, `${pfx}H26`)
      spreadsheet.cellFormat(
        { 
          fontWeight: 'bold', 
          color: '#000000', 
          textAlign: 'right',
          format: '$#,##0.00'
        },
        `${pfx}H26`
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
      const soilExcavationTotalFormula = `=SUM(H25:H29)*1000`
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

      // Row 41: Drilled soldier pile subsection
      spreadsheet.updateCell({ value: 'Drilled soldier pile:' }, `${pfx}B41`)
      spreadsheet.cellFormat(
        { 
          fontWeight: 'bold', 
          color: '#000000', 
          textAlign: 'left',
          backgroundColor: '#D0CECE',
          textDecoration: 'underline',
          border: '1px solid #000000'
        },
        `${pfx}B41`
      )

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
        for (const row of dataRows) {
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
      console.log('=== All Drilled Soldier Pile Groups ===')
      console.log(`Total groups: ${collectedGroups.length}`)
      
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
        
        console.log(`Group ${idx + 1} - Total FT: ${groupFT.toFixed(2)} (SUM(I${firstRowNumber}:I${lastRowNumber}))`)
        console.log(`Group ${idx + 1} - Total LBS: ${groupLBS.toFixed(2)} (SUM(K${firstRowNumber}:K${lastRowNumber}))`)
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
      
      console.log('=== Separated Groups ===')
      console.log(`Drilled Soldier Pile Groups: ${drilledGroups.length}`)
      console.log(`HP Groups from Drilled Section: ${hpGroupsFromDrilled.length}`)
      
      // Add HP groups from drilled soldier pile section to HP groups
      if (hpGroupsFromDrilled.length > 0) {
        if (!window.hpSoldierPileGroups) {
          window.hpSoldierPileGroups = []
        }
        window.hpSoldierPileGroups.push(...hpGroupsFromDrilled)
      }
      
      // Initialize currentRow for both drilled and HP groups
      let currentRow = 42 // Start at row 42
      
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
              
              // Increase row height to accommodate wrapped text
              spreadsheet.setRowHeight(currentRow, 41, proposalSheetIndex)
              
              currentRow++ // Move to next row for next group
            }
          }
        })
      }
      
      // Add HP soldier pile proposal text from collected groups
      const hpCollectedGroups = window.hpSoldierPileGroups || []
      console.log('=== All HP Groups ===')
      console.log(`Total HP groups: ${hpCollectedGroups.length}`)
      
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
        
        console.log(`HP Group ${idx + 1} - Total FT: ${groupFT.toFixed(2)} (SUM(I${firstRowNumber}:I${lastRowNumber}))`)
        console.log(`HP Group ${idx + 1} - Total LBS: ${groupLBS.toFixed(2)} (SUM(K${firstRowNumber}:K${lastRowNumber}))`)
      })
      
      if (hpCollectedGroups.length > 0) {
        // Process each HP group separately
        hpCollectedGroups.forEach((hpCollectedItems, groupIndex) => {
          console.log(`Processing HP Group ${groupIndex + 1} with ${hpCollectedItems.length} items`)
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
            
            console.log(`Group ${groupIndex + 1} - hpType: ${hpType}, hpWeight: ${hpWeight} lbs/ft, totalCount: ${totalCount}, totalFT: ${totalFT.toFixed(2)}, totalLBS: ${totalLBS.toFixed(2)}`)
            
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
              
              console.log(proposalText)
              
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
              
              // Increase row height to accommodate wrapped text
              spreadsheet.setRowHeight(currentRow, 41, proposalSheetIndex)
              
              currentRow++ // Move to next row for next group
            }
          }
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
        'Concrete buttons',
        'Guide wall',
        'Dowels',
        'Rock pins',
        'Rock stabilization',
        'Shotcrete',
        'Permission grouting',
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
        'Anchor'
      ])
      
      // Get all unique subsection names from collected items
      const collectedSubsections = new Set()
      soeSubsectionItems.forEach((groups, name) => {
        if (groups.length > 0) {
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
      
      // Check if any bracing-related subsections exist
      const bracingSubsections = ['Waler', 'Raker', 'Upper raker', 'Lower raker', 'Stand off', 'Kicker', 'Channel', 'Roll chock', 'Stub beam', 'Inner corner brace', 'Knee brace', 'Supporting angle']
      const hasBracingItems = bracingSubsections.some(name => collectedSubsections.has(name))
      
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
        } else if (collectedSubsections.has(name)) {
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        }
      })
      
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
          bracingSubsections.forEach(bracingSubsectionName => {
            const bracingGroups = soeSubsectionItems.get(bracingSubsectionName) || []
            bracingGroups.forEach((group) => {
              if (group.length === 0) return
              
              // Calculate totals for the group
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
              
              // Generate proposal text for bracing item
              let proposalText = `${bracingSubsectionName} item: ${Math.round(totalQty)} nos`
              
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
              
              // Add FT (LF) to column C if available
              if (totalFT > 0 && sumRowIndex > 0) {
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
              
              // Add LBS to column E if available
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
              
              // Add QTY to column G if available
              if (totalQty > 0 && sumRowIndex > 0) {
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
              
              // Increase row height
              spreadsheet.setRowHeight(currentRow, 41, proposalSheetIndex)
              
              currentRow++ // Move to next row
            })
          })
          return // Skip the normal processing for Bracing
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
        if (subsectionName === 'Dowels') {
          const dowelBarGroups = soeSubsectionItems.get('Dowel bar') || []
          groups = [...groups, ...dowelBarGroups]
        } else if (subsectionName === 'Concrete buttons') {
          // Also check for "Buttons" subsection name
          const buttonsGroups = soeSubsectionItems.get('Buttons') || []
          groups = [...groups, ...buttonsGroups]
        }
        groups.forEach((group) => {
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
          
          group.forEach(item => {
            totalTakeoff += item.takeoff || 0
            totalQty += item.qty || 0
            const itemFT = (item.height || 0) * (item.takeoff || 0)
            totalFT += itemFT
            totalLBS += item.lbs || 0
            totalSQFT += item.sqft || 0
            if (item.height > 0) {
              avgHeight += item.height * (item.takeoff || 1)
              heightCount += item.takeoff || 1
            }
            lastRowNumber = Math.max(lastRowNumber, item.rawRowNumber || 0)
          })
          
          if (heightCount > 0) {
            avgHeight = avgHeight / heightCount
          }
          
          // Find the sum row for this group (same row as last item)
          const sumRowIndex = lastRowNumber
          const calcSheetName = 'Calculations Sheet'
          
          // Generate proposal text based on subsection type
          // This will be enhanced later with proper parsing and formatting
          let proposalText = `${subsectionName} item: ${Math.round(totalQty)} nos`
          
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
          
          // Add FT (LF) to column C - reference to calculation sheet sum row if available
          if (totalFT > 0 && sumRowIndex > 0) {
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
          if (totalQty > 0 && sumRowIndex > 0) {
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
          
          // Increase row height
          spreadsheet.setRowHeight(currentRow, 41, proposalSheetIndex)
          
          currentRow++ // Move to next row
        })
      })
      
      // Apply green background to rows 36-56, columns I-N
      spreadsheet.cellFormat(
        { 
          backgroundColor: '#E2EFDA'
        },
        `${pfx}I36:N56`
      )
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

import React, { useState, useEffect, useRef } from 'react'
import { useLocation, useNavigate } from 'react-router-dom'
import { SpreadsheetComponent, SheetsDirective, SheetDirective, ColumnsDirective, ColumnDirective } from '@syncfusion/ej2-react-spreadsheet'
import { FiArrowLeft, FiSave } from 'react-icons/fi'
import { generateCalculationSheet, generateColumnConfigs } from '../utils/generateCalculationSheet'
import { generateDemolitionFormulas } from '../utils/processors/demolitionProcessor'
import { generateExcavationFormulas } from '../utils/processors/excavationProcessor'
import { generateRockExcavationFormulas } from '../utils/processors/rockExcavationProcessor'
import { generateSoeFormulas } from '../utils/processors/soeProcessor'
import { generateFoundationFormulas } from '../utils/processors/foundationProcessor'
import { generateWaterproofingFormulas } from '../utils/processors/waterproofingProcessor'

const Spreadsheet = () => {
  const location = useLocation()
  const navigate = useNavigate()
  const { rawData, fileName, sheetName, headers, rows, selectedTemplate } = location.state || {}
  const [hasData, setHasData] = useState(false)
  const [calculationData, setCalculationData] = useState([])
  const [formulaData, setFormulaData] = useState([])
  const [sheetCreated, setSheetCreated] = useState(false)
  const spreadsheetRef = useRef(null)

  useEffect(() => {
    if (rawData && rawData.length > 0) {
      setHasData(true)
      console.log('Raw data received:', { fileName, sheetName, rowCount: rows?.length, template: selectedTemplate })

      // Generate the Calculations Sheet structure with data
      const template = selectedTemplate || 'capstone'
      const result = generateCalculationSheet(template, rawData)
      setCalculationData(result.rows)
      setFormulaData(result.formulas)

      console.log('Calculation sheet generated with', result.rows.length, 'rows and', result.formulas.length, 'formula items')
    } else {
      // No data - create empty structure
      const template = selectedTemplate || 'capstone'
      const result = generateCalculationSheet(template, null)
      setCalculationData(result.rows)
      setFormulaData(result.formulas || [])
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
          console.error(`Error updating cell ${cellAddress}:`, error)
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
            console.error(`Error applying demolition sum formula at row ${row}:`, error)
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
            console.error(`Error applying demo extra SQ FT formula at row ${row}:`, error)
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
            console.error(`Error applying demo extra FT formula at row ${row}:`, error)
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
            console.error(`Error applying demo extra EA formula at row ${row}:`, error)
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
            console.error(`Error applying demo extra SQ FT formula at row ${row}:`, error)
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
            console.error(`Error applying demo extra FT formula at row ${row}:`, error)
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
            console.error(`Error applying demo extra EA formula at row ${row}:`, error)
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
            console.error(`Error applying sum formula at row ${row}:`, error)
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
            console.error(`Error applying Havg formula at row ${row}:`, error)
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
            console.error(`Error applying backfill sum formula at row ${row}:`, error)
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
            console.error(`Error applying mud slab sum formula at row ${row}:`, error)
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
            console.error(`Error applying backfill extra SQ FT formula at row ${row}:`, error)
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
            console.error(`Error applying backfill extra FT formula at row ${row}:`, error)
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
            console.error(`Error applying backfill extra EA formula at row ${row}:`, error)
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
            console.error(`Error applying soil exc extra SQ FT formula at row ${row}:`, error)
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
            console.error(`Error applying soil exc extra FT formula at row ${row}:`, error)
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
            console.error(`Error applying soil exc extra EA formula at row ${row}:`, error)
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
            console.error(`Error formatting line drill sub header at row ${row}:`, error)
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
            console.error(`Error applying line drill pier formula at row ${row}:`, error)
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
            console.error(`Error applying line drill sewage formula at row ${row}:`, error)
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
            console.error(`Error applying line drill sump formula at row ${row}:`, error)
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
            console.error(`Error applying line drilling formula at row ${row}:`, error)
          }
          return
        }

        if (itemType === 'line_drill_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})*2` }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
          } catch (error) {
            console.error(`Error applying line drill sum formula at row ${row}:`, error)
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
            console.error(`Error applying rock excavation sum formula at row ${row}:`, error)
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
            console.error(`Error applying rock excavation Havg formula at row ${row}:`, error)
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
            console.error(`Error applying rock extra SQ FT formula at row ${row}:`, error)
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
            console.error(`Error applying rock extra FT formula at row ${row}:`, error)
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
            console.error(`Error applying rock extra EA formula at row ${row}:`, error)
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
            console.error(`Error applying SOE formula at row ${row}:`, error)
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
            console.error(`Error applying SOE sum formula at row ${row}:`, error)
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
      } else if (section === 'superstructure') {
        if (itemType === 'superstructure_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Superstructure sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_slab_steps_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Superstructure slab steps sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_lw_concrete_fill_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Superstructure LW concrete fill sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_item') {
          const parsed = parsedData || formulaInfo
          const heightFormula = parsed?.parsed?.heightFormula
          const heightValue = parsed?.parsed?.heightValue
          try {
            if (heightFormula) {
              spreadsheet.updateCell({ formula: `=${heightFormula}` }, `H${row}`)
            }
            if (heightValue != null) {
              spreadsheet.updateCell({ value: heightValue }, `H${row}`)
            }
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Superstructure item formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_slab_step') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Superstructure slab step formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_lw_concrete_fill') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Superstructure LW concrete fill formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_somd_item') {
          try {
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying SOMD item format at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_somd_gen1') {
          const { firstDataRow, lastDataRow, heightFormula } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(C${firstDataRow}:C${lastDataRow})` }, `C${row}`)
            spreadsheet.updateCell({ formula: `=${heightFormula}` }, `H${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=(J${row}*H${row})/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ backgroundColor: '#DDEBF7' }, `H${row}`)
          } catch (error) {
            console.error(`Error applying SOMD gen1 formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_somd_gen2') {
          const { takeoffRefRow, heightValue } = formulaInfo
          try {
            if (takeoffRefRow != null) {
              spreadsheet.updateCell({ formula: `=C${takeoffRefRow}` }, `C${row}`)
            }
            spreadsheet.updateCell({ value: heightValue }, `H${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=(J${row}*H${row})/27/2` }, `L${row}`)
            spreadsheet.cellFormat({ backgroundColor: '#DDEBF7' }, `H${row}`)
          } catch (error) {
            console.error(`Error applying SOMD gen2 formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_somd_sum') {
          const { gen1Row, gen2Row } = formulaInfo
          try {
            if (gen1Row != null) {
              spreadsheet.updateCell({ formula: `=J${gen1Row}` }, `J${row}`)
            }
            if (gen1Row != null && gen2Row != null) {
              spreadsheet.updateCell({ formula: `=SUM(L${gen1Row}:L${gen2Row})` }, `L${row}`)
            }
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying SOMD sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_topping_slab') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Topping slab formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_topping_slab_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Topping slab sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_thermal_break') {
          const parsed = parsedData || formulaInfo
          const qty = parsed?.parsed?.qty
          try {
            if (qty != null) {
              spreadsheet.updateCell({ formula: `=C${row}*E${row}` }, `I${row}`)
            } else {
              spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            }
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Thermal break formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_thermal_break_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
          } catch (error) {
            console.error(`Error applying Thermal break sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_raised_knee_wall') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Raised knee wall formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_raised_styrofoam') {
          const { takeoffRefRow, heightRefRow } = formulaInfo
          try {
            if (takeoffRefRow != null) {
              spreadsheet.updateCell({ formula: `=C${takeoffRefRow}` }, `C${row}`)
            }
            if (heightRefRow != null) {
              spreadsheet.updateCell({ formula: `=H${heightRefRow}` }, `H${row}`)
            }
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Raised styrofoam formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_raised_slab') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Raised slab formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_raised_slab_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Raised slab sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_builtup_knee_wall') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Built-up knee wall formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_builtup_styrofoam') {
          const { takeoffRefRow, heightRefRow } = formulaInfo
          try {
            if (takeoffRefRow != null) {
              spreadsheet.updateCell({ formula: `=C${takeoffRefRow}` }, `C${row}`)
            }
            if (heightRefRow != null) {
              spreadsheet.updateCell({ formula: `=H${heightRefRow}` }, `H${row}`)
            }
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Built-up styrofoam formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_builtup_slab') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Built-up slab formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_builtup_slab_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Built-up slab sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_builtup_ramps_knee_wall') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Builtup ramps knee wall formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_builtup_ramps_knee_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Builtup ramps knee sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_builtup_ramps_styrofoam') {
          const { takeoffRefRow, heightRefRow } = formulaInfo
          try {
            if (takeoffRefRow != null) {
              spreadsheet.updateCell({ formula: `=C${takeoffRefRow}` }, `C${row}`)
            }
            if (heightRefRow != null) {
              spreadsheet.updateCell({ formula: `=H${heightRefRow}` }, `H${row}`)
            }
            // col I empty for builtup ramps styrofoam
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Builtup ramps styrofoam formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_builtup_ramps_styro_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Builtup ramps styro sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_builtup_ramps_ramp') {
          try {
            // col G and col I empty for builtup ramp items
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Builtup ramps ramp formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_builtup_ramps_ramp_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Builtup ramps ramp sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_builtup_stair_knee_wall') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Built-up stair knee wall formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_builtup_stair_styrofoam') {
          const { takeoffJSumFirstRow, takeoffJSumLastRow, heightRefRow } = formulaInfo
          try {
            if (takeoffJSumFirstRow != null && takeoffJSumLastRow != null) {
              spreadsheet.updateCell({ formula: `=SUM(J${takeoffJSumFirstRow}:J${takeoffJSumLastRow})` }, `C${row}`)
            }
            if (heightRefRow != null) {
              spreadsheet.updateCell({ formula: `=H${heightRefRow}` }, `H${row}`)
            }
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Built-up stair styrofoam formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_builtup_stairs') {
          try {
            spreadsheet.updateCell({ formula: `=11/12` }, `F${row}`)
            spreadsheet.updateCell({ formula: `=7/12` }, `H${row}`)
            spreadsheet.updateCell({ formula: `=C${row}*G${row}*F${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Built up Stairs formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_stair_slab') {
          const { takeoffRefRow, widthRefRow, heightValue } = formulaInfo
          try {
            if (takeoffRefRow != null) {
              spreadsheet.updateCell({ formula: `=C${takeoffRefRow}*1.3` }, `C${row}`)
            }
            if (widthRefRow != null) {
              spreadsheet.updateCell({ formula: `=G${widthRefRow}` }, `G${row}`)
            }
            spreadsheet.updateCell({ value: heightValue }, `H${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Stair slab formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_builtup_stair_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
          } catch (error) {
            console.error(`Error applying Built-up stair sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_concrete_hanger') {
          try {
            spreadsheet.updateCell({ formula: `=G${row}*F${row}*C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Concrete hanger formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_concrete_hanger_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
          } catch (error) {
            console.error(`Error applying Concrete hanger sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_shear_wall') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Shear wall formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_shear_walls_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Shear Walls sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_parapet_wall') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Parapet wall formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_parapet_walls_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Parapet walls sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_columns_takeoff') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', textDecoration: 'line-through' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#000000', textDecoration: 'line-through' }, `C${row}`)
            spreadsheet.cellFormat({ color: '#000000', textDecoration: 'line-through' }, `D${row}`)
            spreadsheet.cellFormat({ color: '#000000', textDecoration: 'line-through' }, `M${row}`)
          } catch (error) {
            console.error(`Error applying Columns takeoff row formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_columns_final') {
          const { takeoffRefRow } = formulaInfo
          try {
            if (takeoffRefRow != null) {
              spreadsheet.updateCell({ formula: `=C${takeoffRefRow}` }, `C${row}`)
            }
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'normal' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'normal' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'normal' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
          } catch (error) {
            console.error(`Error applying Columns final row formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_concrete_post') {
          try {
            spreadsheet.updateCell({ formula: `=G${row}*H${row}*C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*F${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Concrete post formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_concrete_post_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
          } catch (error) {
            console.error(`Error applying Concrete post sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_concrete_encasement') {
          try {
            spreadsheet.updateCell({ formula: `=G${row}*H${row}*C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*F${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Concrete encasement formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_concrete_encasement_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
          } catch (error) {
            console.error(`Error applying Concrete encasement sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_drop_panel_bracket') {
          try {
            spreadsheet.updateCell({ formula: `=G${row}*H${row}*C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*F${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Drop panel bracket formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_drop_panel_h') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=E${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Drop panel H formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_drop_panel_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
          } catch (error) {
            console.error(`Error applying Drop panel sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_beam') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Beam formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_beams_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Beams sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_curb') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Curb formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_curbs_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
          } catch (error) {
            console.error(`Error applying Curbs sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_concrete_pad') {
          try {
            const noBracket = formulaInfo.parsedData?.parsed?.noBracket
            if (noBracket) {
              spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            } else {
              spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
              spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
              spreadsheet.updateCell({ formula: `=E${row}` }, `M${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            }
          } catch (error) {
            console.error(`Error applying Concrete pad formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_concrete_pad_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
          } catch (error) {
            console.error(`Error applying Concrete pad sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_non_shrink_grout') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
          } catch (error) {
            console.error(`Error applying Non-shrink grout formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_non_shrink_grout_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
          } catch (error) {
            console.error(`Error applying Non-shrink grout sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_repair_scope') {
          try {
            const subType = formulaInfo.parsedData?.parsed?.itemSubType
            if (subType === 'wall') {
              spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            } else if (subType === 'slab') {
              spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            } else if (subType === 'column') {
              spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            }
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'normal', fontStyle: 'normal' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'normal', fontStyle: 'normal' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000', fontWeight: 'normal', fontStyle: 'normal' }, `M${row}`)
          } catch (error) {
            console.error(`Error applying Repair scope formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_repair_scope_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          try {
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
          } catch (error) {
            console.error(`Error applying Repair scope sum formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_extra_sqft') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
          } catch (error) {
            console.error(`Error applying Superstructure extra SQ FT formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_extra_ft') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
          } catch (error) {
            console.error(`Error applying Superstructure extra FT formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'superstructure_extra_ea') {
          try {
            spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.updateCell({ formula: `=C${row}` }, `M${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
          } catch (error) {
            console.error(`Error applying Superstructure extra EA formula at row ${row}:`, error)
          }
          return
        }
      } else if (section === 'trenching') {
        if (itemType === 'trenching_section_header') {
          const { patchbackRow } = formulaInfo
          try {
            if (patchbackRow != null) {
              spreadsheet.updateCell({ formula: `=L${patchbackRow}` }, `C${row}`)
            }
          } catch (error) {
            console.error(`Error applying Trenching section header formula at row ${row}:`, error)
          }
          return
        }
        if (itemType === 'trenching_item') {
          const { takeoffRefRow, hFormula, h, lYellow } = formulaInfo
          try {
            if (takeoffRefRow != null) {
              spreadsheet.updateCell({ formula: `=C${takeoffRefRow}` }, `C${row}`)
            }
            if (hFormula) {
              spreadsheet.updateCell({ formula: `=${hFormula}` }, `H${row}`)
            }
            spreadsheet.updateCell({ formula: `=C${row}` }, `I${row}`)
            spreadsheet.updateCell({ formula: `=I${row}*G${row}` }, `J${row}`)
            spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, `L${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `J${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `L${row}`)
            if (lYellow) {
              spreadsheet.cellFormat({ backgroundColor: '#FFF2CC' }, `L${row}`)
            }
          } catch (error) {
            console.error(`Error applying Trenching item formula at row ${row}:`, error)
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
        console.error(`Error applying formula at row ${row}:`, error)
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
          } else if (sectionName === 'Trenching') {
            backgroundColor = '#C6E0B4'
          } else if (sectionName === 'Superstructure') {
            backgroundColor = '#C6E0B4'
          } else if (sectionName === 'B.P.P. Alternate #2 scope') {
            backgroundColor = '#C6E0B4'
          } else if (sectionName === 'Civil / Sitework') {
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
          // Foundation and Trenching header: sum (C) and CY (D) should not be bold
          if (sectionName === 'Foundation' || sectionName === 'Trenching') {
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
              bContent.includes('For foundation Extra line item use this') ||
              bContent.includes('For Superstructure Extra line item use this')) {
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
      console.error('Error applying formatting:', error)
    }
  }

  const onCreated = () => {
    if (sheetCreated) return
    if (calculationData.length > 0) {
      applyDataToSpreadsheet()
      setSheetCreated(true)
    }
  }

  const handleGoBack = () => {
    navigate('/dashboard')
  }

  const handleSave = () => {
    // TODO: Implement save to MongoDB
    console.log('Save functionality to be implemented')
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

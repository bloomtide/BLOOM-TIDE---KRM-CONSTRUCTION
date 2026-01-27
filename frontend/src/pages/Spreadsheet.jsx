import React, { useState, useEffect, useRef } from 'react'
import { useLocation, useNavigate } from 'react-router-dom'
import { SpreadsheetComponent, SheetsDirective, SheetDirective, ColumnsDirective, ColumnDirective } from '@syncfusion/ej2-react-spreadsheet'
import { FiArrowLeft, FiSave } from 'react-icons/fi'
import { generateCalculationSheet, generateColumnConfigs } from '../utils/generateCalculationSheet'
import { generateDemolitionFormulas } from '../utils/processors/demolitionProcessor'
import { generateExcavationFormulas } from '../utils/processors/excavationProcessor'
import { generateRockExcavationFormulas } from '../utils/processors/rockExcavationProcessor'
import { generateSoeFormulas } from '../utils/processors/soeProcessor'

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

    // Apply formulas
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
        if (itemType === 'soldier_pile_item' || itemType === 'soe_generic_item' || itemType === 'backpacking_item') {
          try {
            const soeFormulas = generateSoeFormulas(itemType, row, parsedData || formulaInfo)
            if (soeFormulas.takeoff) spreadsheet.updateCell({ formula: `=${soeFormulas.takeoff}` }, `C${row}`)
            if (soeFormulas.ft) spreadsheet.updateCell({ formula: `=${soeFormulas.ft}` }, `I${row}`)
            if (soeFormulas.sqFt) spreadsheet.updateCell({ formula: `=${soeFormulas.sqFt}` }, `J${row}`)
            if (soeFormulas.lbs) spreadsheet.updateCell({ formula: `=${soeFormulas.lbs}` }, `K${row}`)
            if (soeFormulas.qtyFinal) spreadsheet.updateCell({ formula: `=${soeFormulas.qtyFinal}` }, `M${row}`)

            // Apply red color to item names
            spreadsheet.cellFormat({ color: '#FF0000' }, `B${row}`)

            // Special formatting for Backpacking
            if (itemType === 'backpacking_item') {
              spreadsheet.cellFormat({ color: '#000000' }, `C${row}`)
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
            // Standard sum for FT (I) for all SOE
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, `I${row}`)
            spreadsheet.cellFormat({ color: '#FF0000' }, `I${row}`)

            // Sum for SQ FT (J)
            const sqFtSubsections = ['Sheet pile', 'Timber lagging', 'Timber sheeting']
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
              'Knee brace'
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
              'Knee brace'
            ]
            if (itemType === 'soldier_pile_group_sum' || qtySubsections.includes(subsectionName)) {
              spreadsheet.updateCell({ formula: `=SUM(M${firstDataRow}:M${lastDataRow})` }, `M${row}`)
              spreadsheet.cellFormat({ color: '#FF0000' }, `M${row}`)
            }
          } catch (error) {
            console.error(`Error applying SOE sum formula at row ${row}:`, error)
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
        'B1:N1'
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
          }
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              backgroundColor: backgroundColor,
              fontSize: '11pt'
            },
            `A${rowNum}:N${rowNum}`
          )
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
              bContent.includes('For rock excavation Extra line item use this')) {
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

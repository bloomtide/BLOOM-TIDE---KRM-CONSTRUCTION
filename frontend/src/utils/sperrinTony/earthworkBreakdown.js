/**
 * Sperrin Tony – Earthwork Breakdown tab.
 * Same structure as calculation sheet for Demolition, Excavation, and Rock Excavation.
 * Columns I–M are filled by applying the same formula logic as Capstone (demolition/excavation/rock processors).
 */

import { generateDemolitionFormulas } from '../processors/demolitionProcessor'
import { generateExcavationFormulas } from '../processors/excavationProcessor'
import { generateRockExcavationFormulas } from '../processors/rockExcavationProcessor'

const EARTHWORK_SECTIONS = ['Demolition', 'Excavation', 'Rock Excavation']

// Column headers A–H same as capstone; I–M order: FT, SQ FT, CY, QTY, LBS
const COLUMN_HEADERS = [
  'Estimate',
  'Particulars',
  'Takeoff',
  'Unit',
  'QTY',
  'Length',
  'Width',
  'Height',
  'FT',
  'SQ FT',
  'CY',
  'QTY',
  'LBS'
]

const ESTIMATE_HEADER_BG = '#ffff00'   // Column A header
const B_TO_M_HEADER_BG = '#c6e0b4'   // Columns B–M header
const COL_B_WIDTH_PX = 533
const COL_REST_WIDTH_PX = 154
const ROW_HEIGHT_PX = 30
const RED_TEXT = '#FF0000'  // red for particulars and sum rows
const NUM_COLS = 13
const FONT_SIZE = '13pt'
const BASE_FORMAT = { fontSize: FONT_SIZE }
// Rows 1–3 only: center-aligned and bold
const ROWS_1_3_FORMAT = { fontSize: FONT_SIZE, textAlign: 'center', fontWeight: 'bold' }

/** Round numeric values to 2 decimal places for display; leave non-numbers as-is */
function toTwoDecimals(val) {
  if (val === '' || val == null || val === undefined) return val
  if (typeof val === 'number') {
    if (Number.isInteger(val)) return val
    return Math.round(val * 100) / 100
  }
  const n = parseFloat(val)
  if (!Number.isNaN(n)) return Math.round(n * 100) / 100
  return val
}

/**
 * Get column letter for 0-based index (A, B, ... M for 0–12)
 */
function colLetter(c) {
  if (c <= 25) return String.fromCharCode(65 + c)
  return String.fromCharCode(64 + Math.floor(c / 26)) + String.fromCharCode(65 + (c % 26))
}

// I–M on Earthwork Breakdown: FT, SQ FT, CY, QTY, LBS (calculation data has I=FT, J=SQ FT, K=LBS, L=CY, M=QTY)
// Map logical index 8–12 to physical column for writing: 8→I, 9→J, 10→M(LBS), 11→K(CY), 12→L(QTY)
const EARTHWORK_I_M_MAP = { 8: 'I', 9: 'J', 10: 'M', 11: 'K', 12: 'L' }
function earthworkColLetter(c) {
  if (c >= 8 && c <= 12) return EARTHWORK_I_M_MAP[c]
  return colLetter(c)
}

/**
 * Extract earthwork rows from full calculation sheet data: header row, empty row, then all rows
 * from Demolition through Rock Excavation (until next section e.g. SOE).
 */
function getEarthworkRows(calculationData) {
  if (!calculationData || !Array.isArray(calculationData) || calculationData.length < 2) return []
  const rows = calculationData
  let endIdx = rows.length
  for (let i = 2; i < rows.length; i++) {
    const sectionName = (rows[i][0] != null && rows[i][0] !== '') ? String(rows[i][0]).trim() : ''
    if (sectionName && !EARTHWORK_SECTIONS.includes(sectionName)) {
      endIdx = i
      break
    }
  }
  return rows.slice(0, endIdx)
}

/**
 * Build set of 1-based row numbers that are sum rows within earthwork (from formulaData).
 */
function getEarthworkSumRowNumbers(formulaData) {
  if (!formulaData || !Array.isArray(formulaData)) return new Set()
  const sections = new Set(['demolition', 'excavation', 'rock_excavation'])
  const set = new Set()
  formulaData.forEach((f) => {
    if (!f.row) return
    const section = (f.section || '').toLowerCase()
    if (!sections.has(section)) return
    const type = (f.itemType || '').toLowerCase()
    if (type.includes('sum')) set.add(f.row)
  })
  return set
}

/**
 * Apply I–M (and E where needed) formulas to Earthwork Breakdown sheet using formulaData.
 * Mirrors ProposalDetail applyFormulasAndStyles for demolition, excavation, rock_excavation only.
 */
function applyEarthworkFormulas(spreadsheet, pfx, formulaData) {
  if (!formulaData || !Array.isArray(formulaData)) return
  const sections = new Set(['demolition', 'excavation', 'rock_excavation'])

  formulaData.forEach((formulaInfo) => {
    const { row, itemType, parsedData, section, subsection } = formulaInfo
    if (!row || !section || !sections.has(section.toLowerCase())) return

    const cell = (col) => `${pfx}${col}${row}`

    try {
      if (section === 'demolition') {
        if (itemType === 'demolition_sum') {
          const { firstDataRow, lastDataRow, subsection: subName } = formulaInfo
          const hasFtColumn = ['Demo strip footing', 'Demo foundation wall', 'Demo retaining wall'].includes(subName)
          if (hasFtColumn) {
            spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})` }, cell('I'))
          }
          spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, cell('J'))
          spreadsheet.updateCell({ formula: `=SUM(K${firstDataRow}:K${lastDataRow})` }, cell('K'))
          if (subName === 'Demo isolated footing') {
            spreadsheet.updateCell({ formula: `=SUM(L${firstDataRow}:L${lastDataRow})` }, cell('L'))
          }
          return
        }
        if (itemType === 'demo_stair_on_grade_landing') {
          spreadsheet.updateCell({ formula: `=C${row}` }, cell('J'))
          spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
          return
        }
        if (itemType === 'demo_stair_on_grade_stairs') {
          spreadsheet.updateCell({ formula: '=11/12' }, cell('F'))
          spreadsheet.updateCell({ formula: '=7/12' }, cell('H'))
          spreadsheet.updateCell({ formula: `=C${row}*G${row}*F${row}` }, cell('J'))
          spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
          spreadsheet.updateCell({ formula: `=C${row}` }, cell('L'))
          return
        }
        if (itemType === 'demo_stair_on_grade_stair_slab') {
          const { stairsRefRow } = formulaInfo
          if (stairsRefRow) {
            spreadsheet.updateCell({ formula: `=1.3*C${stairsRefRow}` }, cell('C'))
            spreadsheet.updateCell({ formula: `=G${stairsRefRow}` }, cell('G'))
          }
          spreadsheet.updateCell({ formula: `=C${row}` }, cell('I'))
          spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, cell('J'))
          spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, cell('K'))
          return
        }
        if (itemType === 'demo_stair_on_grade_sum') {
          const { sumRanges } = formulaInfo
          if (sumRanges && sumRanges.length > 0) {
            const sumI = sumRanges.map(([f, l]) => `SUM(I${f}:I${l})`).join('+')
            const sumJ = sumRanges.map(([f, l]) => `SUM(J${f}:J${l})`).join('+')
            const sumK = sumRanges.map(([f, l]) => `SUM(K${f}:K${l})`).join('+')
            const sumL = sumRanges.map(([f, l]) => `SUM(L${f}:L${l})`).join('+')
            const sumM = sumRanges.map(([f, l]) => `SUM(M${f}:M${l})`).join('+')
            spreadsheet.updateCell({ formula: `=${sumI}` }, cell('I'))
            spreadsheet.updateCell({ formula: `=${sumJ}` }, cell('J'))
            spreadsheet.updateCell({ formula: `=${sumK}` }, cell('K'))
            spreadsheet.updateCell({ formula: `=${sumL}` }, cell('L'))
            spreadsheet.updateCell({ formula: `=${sumM}` }, cell('M'))
          }
          return
        }
        if (itemType === 'demo_extra_sqft' || itemType === 'demo_extra_rog_sqft') {
          spreadsheet.updateCell({ formula: `=C${row}` }, cell('J'))
          spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
          return
        }
        if (itemType === 'demo_extra_ft' || itemType === 'demo_extra_rw') {
          spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, cell('J'))
          spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
          return
        }
        if (itemType === 'demo_extra_ea') {
          spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, cell('J'))
          spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
          spreadsheet.updateCell({ formula: `=C${row}` }, cell('L'))
          return
        }
        const formulas = generateDemolitionFormulas(itemType === 'demolition_item' ? subsection : itemType, row, parsedData)
        if (formulas?.ft) spreadsheet.updateCell({ formula: `=${formulas.ft}` }, cell('I'))
        if (formulas?.sqFt) spreadsheet.updateCell({ formula: `=${formulas.sqFt}` }, cell('J'))
        if (formulas?.cy) spreadsheet.updateCell({ formula: `=${formulas.cy}` }, cell('K'))
        if (formulas?.qtyFinal) spreadsheet.updateCell({ formula: `=${formulas.qtyFinal}` }, cell('L'))
        if (formulas?.lbs) {
          spreadsheet.updateCell({ formula: `=${formulas.lbs}` }, cell('M'))
          spreadsheet.cellFormat({ format: '#,##0.00', textAlign: 'right', backgroundColor: '#d9e1f2' }, cell('M'))
        }
        return
      }

      if (section === 'excavation') {
        if (itemType === 'excavation_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, cell('J'))
          spreadsheet.updateCell({ formula: `=SUM(K${firstDataRow}:K${lastDataRow})` }, cell('K'))
          return
        }
        if (itemType === 'excavation_havg') {
          const { sumRowNumber } = formulaInfo
          spreadsheet.updateCell({ formula: `=(K${sumRowNumber}*27)/J${sumRowNumber}` }, cell('C'))
          return
        }
        if (itemType === 'backfill_sum' || itemType === 'mud_slab_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, cell('J'))
          spreadsheet.updateCell({ formula: `=SUM(K${firstDataRow}:K${lastDataRow})` }, cell('K'))
          return
        }
        if (itemType === 'backfill_extra_sqft') {
          spreadsheet.updateCell({ formula: `=C${row}` }, cell('J'))
          spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
          return
        }
        if (itemType === 'backfill_extra_ft') {
          spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, cell('J'))
          spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
          return
        }
        if (itemType === 'backfill_extra_ea') {
          spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, cell('J'))
          spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
          spreadsheet.updateCell({ formula: `=C${row}` }, cell('L'))
          return
        }
        if (itemType === 'soil_exc_extra_sqft') {
          spreadsheet.updateCell({ formula: `=C${row}` }, cell('J'))
          spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
          spreadsheet.updateCell({ formula: `=K${row}*1.3` }, cell('L'))
          return
        }
        if (itemType === 'soil_exc_extra_ft') {
          spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, cell('J'))
          spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
          spreadsheet.updateCell({ formula: `=K${row}*1.3` }, cell('L'))
          return
        }
        if (itemType === 'soil_exc_extra_ea') {
          spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, cell('J'))
          spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
          spreadsheet.updateCell({ formula: `=K${row}*1.3` }, cell('L'))
          return
        }
        const formulas = generateExcavationFormulas(itemType === 'excavation_item' ? (parsedData?.itemType || itemType) : itemType, row, parsedData)
        if (formulas?.ft) spreadsheet.updateCell({ formula: `=${formulas.ft}` }, cell('I'))
        if (formulas?.sqFt) spreadsheet.updateCell({ formula: `=${formulas.sqFt}` }, cell('J'))
        if (formulas?.cy) spreadsheet.updateCell({ formula: `=${formulas.cy}` }, cell('K'))
        if (formulas?.qtyFinal) spreadsheet.updateCell({ formula: `=${formulas.qtyFinal}` }, cell('L'))
        if (formulas?.lbs) {
          spreadsheet.updateCell({ formula: `=${formulas.lbs}` }, cell('M'))
          spreadsheet.cellFormat({ format: '#,##0.00', textAlign: 'right', backgroundColor: '#d9e1f2' }, cell('M'))
        }
        return
      }

      if (section === 'rock_excavation') {
        if (itemType === 'rock_excavation_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          spreadsheet.updateCell({ formula: `=SUM(J${firstDataRow}:J${lastDataRow})` }, cell('J'))
          spreadsheet.updateCell({ formula: `=SUM(K${firstDataRow}:K${lastDataRow})` }, cell('K'))
          return
        }
        if (itemType === 'rock_excavation_havg') {
          const { sumRowNumber } = formulaInfo
          spreadsheet.updateCell({ formula: `=(K${sumRowNumber}*27)/J${sumRowNumber}` }, cell('C'))
          return
        }
        if (itemType === 'line_drill_concrete_pier') {
          const { refRow } = formulaInfo
          if (refRow) {
            spreadsheet.updateCell({ formula: `=B${refRow}` }, cell('B'))
            spreadsheet.updateCell({ formula: `=((G${refRow}+F${refRow})*2)*C${refRow}` }, cell('C'))
            spreadsheet.updateCell({ formula: `=H${refRow}` }, cell('H'))
          }
          spreadsheet.updateCell({ value: 'FT' }, cell('D'))
          spreadsheet.updateCell({ formula: `=ROUNDUP(H${row}/2,0)` }, cell('E'))
          spreadsheet.updateCell({ formula: `=E${row}*C${row}` }, cell('I'))
          return
        }
        if (itemType === 'line_drill_sewage_pit') {
          const { refRow } = formulaInfo
          if (refRow) {
            spreadsheet.updateCell({ formula: `=B${refRow}` }, cell('B'))
            spreadsheet.updateCell({ formula: `=SQRT(C${refRow})*4` }, cell('C'))
            spreadsheet.updateCell({ formula: `=H${refRow}` }, cell('H'))
          }
          spreadsheet.updateCell({ value: 'FT' }, cell('D'))
          spreadsheet.updateCell({ formula: `=ROUNDUP(H${row}/2,0)` }, cell('E'))
          spreadsheet.updateCell({ formula: `=E${row}*C${row}` }, cell('I'))
          return
        }
        if (itemType === 'line_drill_sump_pit') {
          const { refRow } = formulaInfo
          if (refRow) {
            spreadsheet.updateCell({ formula: `=B${refRow}` }, cell('B'))
            spreadsheet.updateCell({ formula: `=C${refRow}*8` }, cell('C'))
          }
          spreadsheet.updateCell({ value: 'FT' }, cell('D'))
          spreadsheet.updateCell({ formula: `=C${row}` }, cell('I'))
          return
        }
        if (itemType === 'line_drilling') {
          const rockFormulas = generateRockExcavationFormulas(itemType, row, parsedData)
          if (rockFormulas?.qty) spreadsheet.updateCell({ formula: `=${rockFormulas.qty}` }, cell('E'))
          if (rockFormulas?.ft) spreadsheet.updateCell({ formula: `=${rockFormulas.ft}` }, cell('I'))
          return
        }
        if (itemType === 'line_drill_sum') {
          const { firstDataRow, lastDataRow } = formulaInfo
          spreadsheet.updateCell({ formula: `=SUM(I${firstDataRow}:I${lastDataRow})*2` }, cell('I'))
          return
        }
        if (itemType === 'rock_exc_extra_sqft') {
          spreadsheet.updateCell({ formula: `=C${row}` }, cell('J'))
          spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
          return
        }
        if (itemType === 'rock_exc_extra_ft') {
          spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, cell('J'))
          spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
          return
        }
        if (itemType === 'rock_exc_extra_ea') {
          spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, cell('J'))
          spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
          spreadsheet.updateCell({ formula: `=C${row}` }, cell('L'))
          return
        }
        const formulas = generateRockExcavationFormulas(itemType === 'rock_excavation_item' ? (parsedData?.itemType || itemType) : itemType, row, parsedData)
        if (formulas?.qty) spreadsheet.updateCell({ formula: `=${formulas.qty}` }, cell('E'))
        if (formulas?.ft) spreadsheet.updateCell({ formula: `=${formulas.ft}` }, cell('I'))
        if (formulas?.sqFt) spreadsheet.updateCell({ formula: `=${formulas.sqFt}` }, cell('J'))
        if (formulas?.cy) spreadsheet.updateCell({ formula: `=${formulas.cy}` }, cell('K'))
        if (formulas?.qtyFinal) spreadsheet.updateCell({ formula: `=${formulas.qtyFinal}` }, cell('L'))
        if (formulas?.lbs) {
          spreadsheet.updateCell({ formula: `=${formulas.lbs}` }, cell('M'))
          spreadsheet.cellFormat({ format: '#,##0.00', textAlign: 'right', backgroundColor: '#d9e1f2' }, cell('M'))
        }
      }
    } catch (e) {
      // ignore per-row formula errors
    }
  })
}

/**
 * @param {object} spreadsheet - Syncfusion spreadsheet instance
 * @param {object} ctx - { calculationData, formulaData, rawData, currentRow, pfx, ... }
 * @returns {number} next row index after this section
 */
export function buildEarthworkBreakdown(spreadsheet, ctx) {
  const pfx = ctx.pfx ?? 'Earthwork Breakdown!'
  const calculationData = ctx.calculationData
  const formulaData = ctx.formulaData || []

  const earthworkRows = getEarthworkRows(calculationData)
  const sumRowNumbers = getEarthworkSumRowNumbers(formulaData)

  const sheetIndex = ctx.sheetIndex ?? 1

  if (earthworkRows.length === 0) {
    spreadsheet.updateCell({ value: 'Earthwork Breakdown' }, `${pfx}B1`)
    spreadsheet.cellFormat(
      { ...ROWS_1_3_FORMAT, color: '#000000', backgroundColor: B_TO_M_HEADER_BG, textDecoration: 'underline' },
      `${pfx}B1`
    )
    setEarthworkColumnWidthsAndRowHeights(spreadsheet, sheetIndex, 1)
    return 1
  }

  // Row 1: all headers in order A–M; COLUMN_HEADERS already has I–M = FT, SQ FT, CY, QTY, LBS
  for (let c = 0; c < NUM_COLS; c++) {
    spreadsheet.updateCell({ value: COLUMN_HEADERS[c] }, `${pfx}${colLetter(c)}1`)
  }
  spreadsheet.cellFormat(
    { ...ROWS_1_3_FORMAT, color: '#000000', backgroundColor: ESTIMATE_HEADER_BG },
    `${pfx}A1`
  )
  spreadsheet.cellFormat(
    { ...ROWS_1_3_FORMAT, color: '#000000', backgroundColor: B_TO_M_HEADER_BG },
    `${pfx}B1:${colLetter(NUM_COLS - 1)}1`
  )

  // Row 2: empty – A2 white, B2:M2 center + bold
  spreadsheet.cellFormat(
    { ...ROWS_1_3_FORMAT, color: '#000000', backgroundColor: 'white' },
    `${pfx}A2`
  )
  spreadsheet.cellFormat(ROWS_1_3_FORMAT, `${pfx}B2:${colLetter(NUM_COLS - 1)}2`)

  // Row 3: section name in A green (only A1 is yellow), B–M headers #c6e0b4
  if (earthworkRows.length >= 3) {
    const sectionRow = earthworkRows[2]
    const sectionName = (sectionRow[0] != null && sectionRow[0] !== '') ? sectionRow[0] : 'Demolition'
    spreadsheet.updateCell({ value: sectionName }, `${pfx}A3`)
    for (let c = 1; c < COLUMN_HEADERS.length; c++) {
      spreadsheet.updateCell({ value: COLUMN_HEADERS[c] }, `${pfx}${colLetter(c)}3`)
    }
    spreadsheet.cellFormat(
      { ...ROWS_1_3_FORMAT, color: '#000000', backgroundColor: B_TO_M_HEADER_BG },
      `${pfx}A3`
    )
    spreadsheet.cellFormat(
      { ...ROWS_1_3_FORMAT, color: '#000000', backgroundColor: B_TO_M_HEADER_BG },
      `${pfx}B3:${colLetter(NUM_COLS - 1)}3`
    )
  }

  // Apply base format (fontSize 11, center) to entire used range
  for (let r = 0; r < earthworkRows.length; r++) {
    const rowNum = r + 1
    for (let c = 0; c < NUM_COLS; c++) {
      try {
        spreadsheet.cellFormat(BASE_FORMAT, `${pfx}${colLetter(c)}${rowNum}`)
      } catch (e) { /* ignore */ }
    }
  }

  // Write all rows from earthwork data. I–M use earthwork order (FT, SQ FT, CY, QTY, LBS).
  // Row 1 (r=0): only write A–H (c 0–7) so I–M are never overwritten with calculation header order (FT, SQ FT, LBS, CY, QTY).
  for (let r = 0; r < earthworkRows.length; r++) {
    const row = earthworkRows[r]
    const rowNum = r + 1
    const maxC = (r === 0) ? 8 : Math.min(row.length, NUM_COLS)
    for (let c = 0; c < maxC; c++) {
      const val = row[c]
      const cellRef = `${pfx}${earthworkColLetter(c)}${rowNum}`
      if (val !== '' && val !== null && val !== undefined) {
        try {
          const displayVal = c >= 2 ? toTwoDecimals(val) : val
          spreadsheet.updateCell({ value: displayVal }, cellRef)
        } catch (e) { /* ignore */ }
      }
    }
  }

  // Re-apply row 2: leave as-is. Section header rows below.
  // Row 3 we already set; if there are more section header rows (Excavation, Rock Excavation), style them and fill B–M with headers
  for (let r = 2; r < earthworkRows.length; r++) {
    const row = earthworkRows[r]
    const rowNum = r + 1
    const sectionName = (row[0] != null && row[0] !== '') ? String(row[0]).trim() : ''
    const isSectionHeader = EARTHWORK_SECTIONS.includes(sectionName)
    const isSubsectionHeader = !isSectionHeader && row[1] != null && String(row[1]).trim().endsWith(':')
    const isSumRow = sumRowNumbers.has(rowNum)

    if (isSectionHeader) {
      spreadsheet.updateCell({ value: sectionName }, `${pfx}A${rowNum}`)
      for (let c = 1; c < COLUMN_HEADERS.length; c++) {
        spreadsheet.updateCell({ value: COLUMN_HEADERS[c] }, `${pfx}${colLetter(c)}${rowNum}`)
      }
      // Section headers after row 3: col A green (only A1 is yellow), B–M green, bold, no center
      spreadsheet.cellFormat(
        { ...BASE_FORMAT, fontWeight: 'bold', color: '#000000', backgroundColor: B_TO_M_HEADER_BG },
        `${pfx}A${rowNum}`
      )
      spreadsheet.cellFormat(
        { ...BASE_FORMAT, fontWeight: 'bold', color: '#000000', backgroundColor: B_TO_M_HEADER_BG },
        `${pfx}B${rowNum}:${colLetter(NUM_COLS - 1)}${rowNum}`
      )
    } else if (isSubsectionHeader) {
      spreadsheet.cellFormat(
        { ...BASE_FORMAT, fontWeight: 'bold', color: '#000000', textAlign: 'left' },
        `${pfx}B${rowNum}`
      )
    } else if (isSumRow) {
      spreadsheet.cellFormat(
        { ...BASE_FORMAT, fontWeight: 'bold', color: RED_TEXT, textAlign: 'right' },
        `${pfx}A${rowNum}:${colLetter(NUM_COLS - 1)}${rowNum}`
      )
      for (let c = 2; c < NUM_COLS; c++) {
        spreadsheet.cellFormat({ ...BASE_FORMAT, format: '#,##0.00', textAlign: 'right' }, `${pfx}${colLetter(c)}${rowNum}`)
      }
      spreadsheet.updateCell({ formula: `=IFERROR(ROUND(I${rowNum}/L${rowNum},2),"")` }, `${pfx}N${rowNum}`)
      spreadsheet.cellFormat({ ...BASE_FORMAT, format: '#,##0.00', textAlign: 'right', color: '#000000', fontWeight: 'bold' }, `${pfx}N${rowNum}`)
    } else {
      // Data row: Particulars (B) in red, left; numeric columns C–M always two decimals
      if (row[1] !== '' && row[1] != null && row[1] !== undefined) {
        spreadsheet.cellFormat(
          { ...BASE_FORMAT, color: RED_TEXT, textAlign: 'left' },
          `${pfx}B${rowNum}`
        )
      }
      for (let c = 2; c < NUM_COLS; c++) {
        spreadsheet.cellFormat(
          { ...BASE_FORMAT, format: '#,##0.00', textAlign: 'right' },
          `${pfx}${colLetter(c)}${rowNum}`
        )
      }
      // LBS (column M): #d9e1f2 only when this row has an LBS value (formula-filled M is styled in applyEarthworkFormulas)
      const hasLbsValue = row[10] !== '' && row[10] != null && row[10] !== undefined
      if (hasLbsValue) {
        spreadsheet.cellFormat(
          { ...BASE_FORMAT, format: '#,##0.00', textAlign: 'right', backgroundColor: '#d9e1f2' },
          `${pfx}M${rowNum}`
        )
      }
    }
  }

  // Apply I–M formulas from formulaData (same logic as Capstone Calculations Sheet; Sperrin Tony has no calc sheet so we apply here)
  applyEarthworkFormulas(spreadsheet, pfx, formulaData)

  setEarthworkColumnWidthsAndRowHeights(spreadsheet, sheetIndex, earthworkRows.length)

  // Red vertical line to the left of column I (FT column)
  const lastRow = earthworkRows.length
  if (lastRow > 0) {
    try {
      spreadsheet.cellFormat({ borderLeft: '3px solid #FF0000' }, `${pfx}I1:I${lastRow}`)
    } catch (e) { /* ignore */ }
  }

  // Final re-apply of row 1 headers so I–M are always FT, SQ FT, CY, QTY, LBS (no other code can overwrite after this)
  for (let c = 0; c < NUM_COLS; c++) {
    try {
      spreadsheet.updateCell({ value: COLUMN_HEADERS[c] }, `${pfx}${colLetter(c)}1`)
    } catch (e) { /* ignore */ }
  }

  return 1
}

/**
 * Set column widths: B = 533px, A and C–M = 154px. Set row height 30px for rows 0..rowCount-1.
 */
function setEarthworkColumnWidthsAndRowHeights(spreadsheet, sheetIndex, rowCount) {
  try {
    for (let c = 0; c < NUM_COLS; c++) {
      const w = c === 1 ? COL_B_WIDTH_PX : COL_REST_WIDTH_PX
      spreadsheet.setColWidth(w, c, sheetIndex)
    }
    for (let r = 0; r < rowCount; r++) {
      spreadsheet.setRowHeight(ROW_HEIGHT_PX, r, sheetIndex)
    }
  } catch (e) { /* ignore */ }
}

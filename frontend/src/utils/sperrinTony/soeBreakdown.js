/**
 * Sperrin Tony – SOE Breakdown tab.
 * Same structure as calculation sheet for Support of Excavation.
 * Columns I–M are filled by applying the same formula logic as Capstone (soeProcessor).
 */

import { generateSoeFormulas } from '../processors/soeProcessor'

const SOE_SECTION_NAME = 'SOE'

// Column headers A–H same as capstone; I–M order on SOE: FT, SQ FT, CY, QTY, LBS (same as Earthwork Breakdown)
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

const ESTIMATE_HEADER_BG = '#ffff00'
const B_TO_M_HEADER_BG = '#c6e0b4'
const COL_B_WIDTH_PX = 533
const COL_REST_WIDTH_PX = 154
const ROW_HEIGHT_PX = 30
const RED_TEXT = '#FF0000'
const NUM_COLS = 13
const FONT_SIZE = '13pt'
const BASE_FORMAT = { fontSize: FONT_SIZE }
const ROWS_1_3_FORMAT = { fontSize: FONT_SIZE, textAlign: 'center', fontWeight: 'bold' }

/** Round numeric values to 2 decimal places for display */
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

function colLetter(c) {
  if (c <= 25) return String.fromCharCode(65 + c)
  return String.fromCharCode(64 + Math.floor(c / 26)) + String.fromCharCode(65 + (c % 26))
}

// I–M on SOE Breakdown: FT, SQ FT, CY, QTY, LBS (calculation has I=FT, J=SQ FT, K=LBS, L=CY, M=QTY)
// Map calculation index 8–12 to sheet column: 8→I, 9→J, 10→M(LBS), 11→K(CY), 12→L(QTY)
const SOE_I_M_MAP = { 8: 'I', 9: 'J', 10: 'M', 11: 'K', 12: 'L' }
function soeColLetter(c) {
  if (c >= 8 && c <= 12) return SOE_I_M_MAP[c]
  return colLetter(c)
}

/**
 * Extract SOE rows from full calculation sheet: find SOE section start/end, return
 * [headerRow, emptyRow, ...soeSectionRows] and the 0-based start index in calculationData.
 */
function getSoeRows(calculationData) {
  if (!calculationData || !Array.isArray(calculationData) || calculationData.length < 2) {
    return { rows: [], soeStartIdx: -1 }
  }
  const rows = calculationData
  let soeStartIdx = -1
  let endIdx = rows.length
  for (let i = 2; i < rows.length; i++) {
    const sectionName = (rows[i][0] != null && rows[i][0] !== '') ? String(rows[i][0]).trim() : ''
    if (sectionName === SOE_SECTION_NAME && soeStartIdx === -1) {
      soeStartIdx = i
    }
    if (soeStartIdx >= 0 && sectionName && sectionName !== SOE_SECTION_NAME) {
      endIdx = i
      break
    }
  }
  if (soeStartIdx < 0) return { rows: [], soeStartIdx: -1 }
  const headerRow = rows[1] && rows[1].length ? rows[1].slice(0, NUM_COLS) : COLUMN_HEADERS.slice()
  const emptyRow = Array(NUM_COLS).fill('')
  const soeSectionRows = rows.slice(soeStartIdx, endIdx)
  return {
    rows: [headerRow, emptyRow, ...soeSectionRows],
    soeStartIdx
  }
}

/**
 * Build set of 1-based sheet row numbers that are sum rows within SOE (from formulaData).
 * Sheet row = formulaInfo.row - soeStartIdx + 2 (row 1=header, 2=empty, 3=first SOE row).
 */
function getSoeSumRowNumbers(formulaData, soeStartIdx) {
  if (!formulaData || !Array.isArray(formulaData) || soeStartIdx < 0) return new Set()
  const set = new Set()
  formulaData.forEach((f) => {
    if (!f.row || (f.section || '').toLowerCase() !== 'soe') return
    const type = (f.itemType || '').toLowerCase()
    if (type.includes('sum')) {
      const sheetRow = f.row - soeStartIdx + 2
      set.add(sheetRow)
    }
  })
  return set
}

/**
 * Map calculation sheet 1-based row to SOE Breakdown sheet 1-based row.
 * soeStartIdx is 0-based; first SOE row is at sheet row 3.
 */
function calcRowToSheetRow(calcRow1Based, soeStartIdx) {
  return calcRow1Based - soeStartIdx + 2
}

/**
 * Apply I–M (and C,E,F,G,H where needed) formulas to SOE Breakdown sheet using formulaData.
 * Mirrors ProposalDetail applyFormulasAndStyles for section === 'soe' only.
 */
function applySoeFormulas(spreadsheet, pfx, formulaData, soeStartIdx) {
  if (!formulaData || !Array.isArray(formulaData) || soeStartIdx < 0) return
  const soeItemTypes = ['soldier_pile_item', 'timber_brace_item', 'timber_waler_item', 'timber_stringer_item', 'drilled_hole_grout_item', 'soe_generic_item', 'backpacking_item', 'supporting_angle', 'parging', 'heel_block', 'underpinning', 'shims', 'rock_anchor', 'rock_bolt', 'anchor', 'tie_back', 'concrete_soil_retention_pier', 'guide_wall', 'dowel_bar', 'rock_pin', 'shotcrete', 'permission_grouting', 'button', 'rock_stabilization', 'form_board']

  formulaData.forEach((formulaInfo) => {
    const { row: calcRow, itemType, parsedData, section } = formulaInfo
    if (!calcRow || (section || '').toLowerCase() !== 'soe') return
    const row = calcRowToSheetRow(calcRow, soeStartIdx)
    const cell = (col) => `${pfx}${col}${row}`

    try {
      if (soeItemTypes.includes(itemType)) {
        const soeFormulas = generateSoeFormulas(itemType, row, parsedData || formulaInfo)
        if (soeFormulas.takeoff) spreadsheet.updateCell({ formula: `=${soeFormulas.takeoff}` }, cell('C'))
        if (soeFormulas.length !== undefined) {
          if (typeof soeFormulas.length === 'string') {
            spreadsheet.updateCell({ formula: `=${soeFormulas.length}` }, cell('F'))
          } else {
            spreadsheet.updateCell({ value: soeFormulas.length }, cell('F'))
          }
        }
        if (soeFormulas.width !== undefined) {
          if (typeof soeFormulas.width === 'string') {
            spreadsheet.updateCell({ formula: `=${soeFormulas.width}` }, cell('G'))
          } else {
            spreadsheet.updateCell({ value: soeFormulas.width }, cell('G'))
          }
        }
        if (soeFormulas.height !== undefined) {
          if (typeof soeFormulas.height === 'string') {
            spreadsheet.updateCell({ formula: `=${soeFormulas.height}` }, cell('H'))
          } else {
            spreadsheet.updateCell({ value: soeFormulas.height }, cell('H'))
          }
        }
        if (soeFormulas.qty !== undefined) {
          if (typeof soeFormulas.qty === 'string') {
            spreadsheet.updateCell({ formula: `=${soeFormulas.qty}` }, cell('E'))
          } else {
            spreadsheet.updateCell({ value: soeFormulas.qty }, cell('E'))
          }
        }
        // SOE sheet I–M order: FT, SQ FT, CY, QTY, LBS
        if (soeFormulas.ft) spreadsheet.updateCell({ formula: `=${soeFormulas.ft}` }, cell('I'))
        if (soeFormulas.sqFt) spreadsheet.updateCell({ formula: `=${soeFormulas.sqFt}` }, cell('J'))
        if (soeFormulas.cy) spreadsheet.updateCell({ formula: `=${soeFormulas.cy}` }, cell('K'))
        if (soeFormulas.qtyFinal) spreadsheet.updateCell({ formula: `=${soeFormulas.qtyFinal}` }, cell('L'))
        if (soeFormulas.lbs) {
          spreadsheet.updateCell({ formula: `=${soeFormulas.lbs}` }, cell('M'))
          spreadsheet.cellFormat({ format: '#,##0.00', textAlign: 'right', backgroundColor: '#d9e1f2' }, cell('M'))
        }
        if (itemType === 'backpacking_item') {
          spreadsheet.cellFormat({ color: '#000000' }, cell('C'))
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('J'))
        }
        if (itemType === 'shims') {
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('I'))
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('J'))
        }
        return
      }

      const groupSumTypes = ['soldier_pile_group_sum', 'timber_soldier_pile_group_sum', 'timber_plank_group_sum', 'timber_raker_group_sum', 'timber_brace_group_sum', 'timber_waler_group_sum', 'timber_stringer_group_sum', 'timber_post_group_sum', 'vertical_timber_sheets_group_sum', 'horizontal_timber_sheets_group_sum', 'drilled_hole_grout_group_sum', 'soe_generic_sum']
      if (groupSumTypes.includes(itemType)) {
        const { firstDataRow, lastDataRow, subsectionName } = formulaInfo
        const firstSheet = calcRowToSheetRow(firstDataRow, soeStartIdx)
        const lastSheet = calcRowToSheetRow(lastDataRow, soeStartIdx)
        const ftSumSubsections = ['Rock anchors', 'Rock bolts', 'Tie back anchor', 'Tie down anchor', 'Dowel bar', 'Rock pins', 'Shotcrete', 'Permission grouting', 'Form board', 'Guide wall']
        const sqFtSubsections = ['Sheet pile', 'Timber lagging', 'Timber sheeting', 'Vertical timber sheets', 'Horizontal timber sheets', 'Parging', 'Heel blocks', 'Underpinning', 'Concrete soil retention piers', 'Guide wall', 'Shotcrete', 'Permission grouting', 'Buttons', 'Form board', 'Rock stabilization']
        const lbsSubsections = ['Secondary secant piles', 'Sheet pile', 'Waler', 'Raker', 'Upper Raker', 'Lower Raker', 'Stand off', 'Kicker', 'Channel', 'Roll chock', 'Stud beam', 'Inner corner brace', 'Knee brace', 'Supporting angle']
        const qtySubsections = ['Primary secant piles', 'Secondary secant piles', 'Tangent piles', 'Waler', 'Raker', 'Upper Raker', 'Lower Raker', 'Stand off', 'Kicker', 'Channel', 'Roll chock', 'Stud beam', 'Inner corner brace', 'Knee brace', 'Supporting angle', 'Heel blocks', 'Underpinning', 'Rock anchors', 'Rock bolts', 'Tie back anchor', 'Tie down anchor', 'Concrete soil retention piers', 'Dowel bar', 'Rock pins', 'Buttons']
        const cySubsections = ['Heel blocks', 'Underpinning', 'Concrete soil retention piers', 'Guide wall', 'Shotcrete', 'Buttons', 'Rock stabilization']

        const doISum = (itemType === 'timber_waler_group_sum' || itemType === 'timber_brace_group_sum' || itemType === 'timber_stringer_group_sum' || itemType === 'drilled_hole_grout_group_sum') ||
          (subsectionName !== 'Heel blocks' && (ftSumSubsections.includes(subsectionName) || !['Concrete soil retention piers', 'Buttons', 'Rock stabilization'].includes(subsectionName)))
        if (doISum) {
          spreadsheet.updateCell({ formula: `=SUM(I${firstSheet}:I${lastSheet})` }, cell('I'))
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('I'))
        }
        if (sqFtSubsections.includes(subsectionName) || itemType === 'drilled_hole_grout_group_sum') {
          spreadsheet.updateCell({ formula: `=SUM(J${firstSheet}:J${lastSheet})` }, cell('J'))
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('J'))
        }
        // SOE sheet columns: I=FT, J=SQ FT, K=CY, L=QTY, M=LBS
        if (itemType === 'soldier_pile_group_sum' || lbsSubsections.includes(subsectionName)) {
          spreadsheet.updateCell({ formula: `=SUM(M${firstSheet}:M${lastSheet})` }, cell('M'))
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('M'))
        }
        if ((itemType === 'soldier_pile_group_sum' || itemType === 'timber_soldier_pile_group_sum' || itemType === 'timber_plank_group_sum' || itemType === 'timber_raker_group_sum' || itemType === 'timber_waler_group_sum' || itemType === 'timber_brace_group_sum' || itemType === 'timber_post_group_sum' || itemType === 'drilled_hole_grout_group_sum') || qtySubsections.includes(subsectionName)) {
          spreadsheet.updateCell({ formula: `=SUM(L${firstSheet}:L${lastSheet})` }, cell('L'))
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('L'))
        }
        if (cySubsections.includes(subsectionName) || itemType === 'drilled_hole_grout_group_sum') {
          spreadsheet.updateCell({ formula: `=SUM(K${firstSheet}:K${lastSheet})` }, cell('K'))
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('K'))
        }
      }
    } catch (e) {
      // ignore per-row formula errors
    }
  })
}

function setSoeColumnWidthsAndRowHeights(spreadsheet, sheetIndex, rowCount) {
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

/**
 * @param {object} spreadsheet - Syncfusion spreadsheet instance
 * @param {object} ctx - { calculationData, formulaData, rawData, currentRow, pfx, sheetIndex, ... }
 * @returns {number} next row index after this section
 */
export function buildSoeBreakdown(spreadsheet, ctx) {
  const pfx = ctx.pfx ?? 'SOE Breakdown!'
  const calculationData = ctx.calculationData || []
  const formulaData = ctx.formulaData || []
  const sheetIndex = ctx.sheetIndex ?? 2

  const { rows: soeRows, soeStartIdx } = getSoeRows(calculationData)
  const sumRowNumbers = getSoeSumRowNumbers(formulaData, soeStartIdx)

  if (soeRows.length === 0) {
    spreadsheet.updateCell({ value: 'SOE Breakdown' }, `${pfx}B1`)
    spreadsheet.cellFormat(
      { ...ROWS_1_3_FORMAT, color: '#000000', backgroundColor: B_TO_M_HEADER_BG, textDecoration: 'underline' },
      `${pfx}B1`
    )
    setSoeColumnWidthsAndRowHeights(spreadsheet, sheetIndex, 1)
    return 1
  }

  // Row 1: A1–M1 all headings
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

  // Row 2: blank
  spreadsheet.cellFormat(
    { ...ROWS_1_3_FORMAT, color: '#000000', backgroundColor: 'white' },
    `${pfx}A2`
  )
  spreadsheet.cellFormat(ROWS_1_3_FORMAT, `${pfx}B2:${colLetter(NUM_COLS - 1)}2`)

  // Row 3: A3 = SOE only, B3–M3 blank with green background
  spreadsheet.updateCell({ value: SOE_SECTION_NAME }, `${pfx}A3`)
  spreadsheet.cellFormat(
    { ...BASE_FORMAT, fontWeight: 'bold', color: '#000000', backgroundColor: B_TO_M_HEADER_BG },
    `${pfx}A3`
  )
  spreadsheet.cellFormat(
    { ...BASE_FORMAT, backgroundColor: B_TO_M_HEADER_BG },
    `${pfx}B3:${colLetter(NUM_COLS - 1)}3`
  )

  // Base format for entire used range
  for (let r = 0; r < soeRows.length; r++) {
    const rowNum = r + 1
    for (let c = 0; c < NUM_COLS; c++) {
      try {
        spreadsheet.cellFormat(BASE_FORMAT, `${pfx}${colLetter(c)}${rowNum}`)
      } catch (e) { /* ignore */ }
    }
  }

  // Write data. I–M use soeColLetter so order is FT, SQ FT, CY, QTY, LBS. Row 1 (r=0): only write A–H so I–M are not overwritten.
  for (let r = 0; r < soeRows.length; r++) {
    const row = soeRows[r]
    const rowNum = r + 1
    const maxC = (r === 0) ? 8 : Math.min(row.length, NUM_COLS)
    for (let c = 0; c < maxC; c++) {
      const val = row[c]
      const cellRef = `${pfx}${soeColLetter(c)}${rowNum}`
      if (val !== '' && val !== null && val !== undefined) {
        try {
          const displayVal = c >= 2 ? toTwoDecimals(val) : val
          spreadsheet.updateCell({ value: displayVal }, cellRef)
        } catch (e) { /* ignore */ }
      }
    }
  }

  // Style section headers, subsection headers, sum rows, data rows (rows 4+)
  for (let r = 3; r < soeRows.length; r++) {
    const row = soeRows[r]
    const rowNum = r + 1
    const sectionName = (row[0] != null && row[0] !== '') ? String(row[0]).trim() : ''
    const isSectionHeader = sectionName === SOE_SECTION_NAME
    const isSubsectionHeader = !isSectionHeader && row[1] != null && String(row[1]).trim().endsWith(':')
    const isSumRow = sumRowNumbers.has(rowNum)

    if (isSectionHeader) {
      // Section row: A = SOE, B–M left blank (already set for row 3; other section rows e.g. from data get A only)
      spreadsheet.updateCell({ value: sectionName }, `${pfx}A${rowNum}`)
      spreadsheet.cellFormat(
        { ...BASE_FORMAT, fontWeight: 'bold', color: '#000000', backgroundColor: B_TO_M_HEADER_BG },
        `${pfx}A${rowNum}`
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
      // LBS (column M): #d9e1f2 only when this row has an LBS value (formula-filled M is styled in applySoeFormulas)
      const hasLbsValue = row[10] !== '' && row[10] != null && row[10] !== undefined
      if (hasLbsValue) {
        spreadsheet.cellFormat(
          { ...BASE_FORMAT, format: '#,##0.00', textAlign: 'right', backgroundColor: '#d9e1f2' },
          `${pfx}M${rowNum}`
        )
      }
    }
  }

  applySoeFormulas(spreadsheet, pfx, formulaData, soeStartIdx)
  setSoeColumnWidthsAndRowHeights(spreadsheet, sheetIndex, soeRows.length)

  // Red vertical line to the left of column I
  if (soeRows.length > 0) {
    try {
      spreadsheet.cellFormat({ borderLeft: '3px solid #FF0000' }, `${pfx}I1:I${soeRows.length}`)
    } catch (e) { /* ignore */ }
  }

  // Final re-apply of row 1 headers so I–M are always FT, SQ FT, CY, QTY, LBS
  for (let c = 0; c < NUM_COLS; c++) {
    try {
      spreadsheet.updateCell({ value: COLUMN_HEADERS[c] }, `${pfx}${colLetter(c)}1`)
    } catch (e) { /* ignore */ }
  }

  return 1
}

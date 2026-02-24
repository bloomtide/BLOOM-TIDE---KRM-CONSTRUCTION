/**
 * Sperrin Tony – Superstructure Breakdown tab.
 * Same structure as Capstone calculation sheet for Superstructure: CIP Slabs, Roof Slab, Balcony,
 * Terrace, Patch slab, Slab steps, LW concrete fill, SOMD, Topping slab, Thermal break, Raised slab,
 * Built-up slab, Builtup ramps, Built-up stair, Concrete hanger, Shear Walls, Columns, Concrete post,
 * Concrete encasement, Drop panel, Beams, CIP Stairs, Infilled stairs, Curbs, Concrete pad,
 * Non-shrink grout, Repair scope. Data rows from calculationData; columns A–M.
 */

// Column order matches Capstone calculation sheet: I=FT, J=SQ FT, K=LBS, L=CY, M=QTY
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
  'LBS',
  'CY',
  'QTY'
]

// Row 1: same as other Sperrin Tony tabs (Earthwork, SOE, Substructure)
const ESTIMATE_HEADER_BG = '#ffff00'   // Column A header
const B_TO_M_HEADER_BG = '#c6e0b4'    // Columns B–M header
const SECTION_ROW_BG = '#c6e0b4'      // Row 3 section header background
const COL_B_WIDTH_PX = 533
const COL_REST_WIDTH_PX = 154
const ROW_HEIGHT_PX = 30
const RED_TEXT = '#FF0000'
const NUM_COLS = 13
const FONT_SIZE = '13pt'
const BASE_FORMAT = { fontSize: FONT_SIZE }
const ROWS_1_3_FORMAT = { fontSize: FONT_SIZE, textAlign: 'center', fontWeight: 'bold' }

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
 * Column B (Particulars) color: red when row has takeoff value > 0, else black.
 * Matches Capstone getColumnBColor logic in ProposalDetail.
 */
function getColumnBColorForRow(row) {
  const takeoffValue = row && row[2]
  if (takeoffValue == null || takeoffValue === '' || takeoffValue === 0 || takeoffValue === '0') return '#000000'
  const valueStr = String(takeoffValue).trim()
  if (valueStr === '' || valueStr.startsWith('=')) return '#000000'
  const numValue = parseFloat(valueStr)
  if (!Number.isNaN(numValue)) return numValue > 0 ? RED_TEXT : '#000000'
  return RED_TEXT
}

/** Parse slab thickness from Particulars (e.g. "Slab 8\"", "Terrace Slab 8\"", "8\"") and return height in feet, or null. */
function parseHeightFromParticulars(particulars) {
  if (particulars == null || particulars === '') return null
  const p = String(particulars).trim()
  const match = p.match(/(\d+)\s*[""]\s*(?:typ\.?)?\s*$/i) || p.match(/(\d+)\s*[""]/) || p.match(/(\d+)\s*inch(?:es)?/i)
  if (match) {
    const inches = parseInt(match[1], 10)
    if (!Number.isNaN(inches) && inches > 0) return Math.round((inches / 12) * 100) / 100
  }
  return null
}

function colLetter(c) {
  if (c <= 25) return String.fromCharCode(65 + c)
  return String.fromCharCode(64 + Math.floor(c / 26)) + String.fromCharCode(65 + (c % 26))
}

/**
 * Get superstructure rows from full calculation data using formulaData to determine range.
 * Includes subsection header rows (e.g. "CIP Slabs:") which are pushed before data in the calc sheet but have no formula.
 */
function getSuperstructureRows(calculationData, formulaData) {
  if (!calculationData || !Array.isArray(calculationData) || calculationData.length === 0) return []
  const superstructureFormulas = (formulaData || []).filter(
    (f) => (f.section || '').toLowerCase() === 'superstructure' && f.row
  )
  if (superstructureFormulas.length === 0) return []
  const rows = superstructureFormulas.map((f) => f.row)
  const firstFormulaRow = Math.min(...rows)
  const lastRow = Math.max(...rows)
  let startIndex = firstFormulaRow - 1 // 0-based index of first formula row
  if (startIndex > 0) {
    const prevRow = calculationData[startIndex - 1]
    const particulars = (prevRow && prevRow[1] != null) ? String(prevRow[1]).trim() : ''
    if (particulars.endsWith(':')) {
      startIndex = startIndex - 1
    }
  }
  return calculationData.slice(startIndex, lastRow)
}

/** Data starts at sheet row 4 (rows 1–3 are headers). */
const DATA_START_ROW = 4

/**
 * Build set of sheet row numbers that are sum/final rows (for styling).
 */
function getSuperstructureSumRowNumbers(formulaData, startIndex) {
  if (!formulaData || !Array.isArray(formulaData)) return new Set()
  const set = new Set()
  formulaData.forEach((f) => {
    if ((f.section || '').toLowerCase() !== 'superstructure') return
    const type = (f.itemType || '').toLowerCase()
    if (type.includes('sum') || type.includes('_final')) {
      const sheetRow = f.row - startIndex + DATA_START_ROW - 1
      set.add(sheetRow)
    }
  })
  return set
}

/**
 * Build set of sheet row numbers that should show Particulars (B) in red:
 * - Rows where C is set by formula (empty in raw data) so getColumnBColorForRow would return black.
 * - Rows that are always data rows (e.g. Columns "As per Takeoff count" / "Final as per schedule count") even when takeoff is 0.
 */
const FORMULA_VALUE_ITEM_TYPES = new Set([
  'superstructure_somd_gen1',
  'superstructure_somd_gen2',
  'superstructure_raised_styrofoam',
  'superstructure_builtup_styrofoam',
  'superstructure_builtup_ramps_styrofoam',
  'superstructure_builtup_stair_styrofoam',
  'superstructure_stair_slab',
  'superstructure_columns_takeoff',   // As per Takeoff count (red even when takeoff is 0)
  'superstructure_columns_final',     // Final as per schedule count (C from formula)
  'superstructure_manual_cip_stairs_landing',
  'superstructure_manual_cip_stairs_stair',
  'superstructure_manual_cip_stairs_slab',
  'superstructure_infilled_landing_item',
  'superstructure_infilled_landing_second',
  'superstructure_manual_infilled_stair',
  'superstructure_ele_item'
])

function getSuperstructureFormulaValueRowNumbers(formulaData, startIndex) {
  if (!formulaData || !Array.isArray(formulaData)) return new Set()
  const set = new Set()
  formulaData.forEach((f) => {
    if ((f.section || '').toLowerCase() !== 'superstructure') return
    const type = (f.itemType || '').toLowerCase()
    if (FORMULA_VALUE_ITEM_TYPES.has(type)) {
      const sheetRow = f.row - startIndex + DATA_START_ROW - 1
      set.add(sheetRow)
    }
  })
  return set
}

/** Rows that should have columns I, J, K, L (FT, SQ FT, LBS, CY) in red: Raised slab / Built-up slab data rows + Columns takeoff/final. */
const RED_I_TO_L_ITEM_TYPES = new Set([
  'superstructure_raised_knee_wall',
  'superstructure_raised_styrofoam',
  'superstructure_raised_slab',
  'superstructure_builtup_knee_wall',
  'superstructure_builtup_styrofoam',
  'superstructure_builtup_slab',
  'superstructure_builtup_ramps_knee_wall',
  'superstructure_builtup_ramps_styrofoam',
  'superstructure_builtup_ramps_ramp',
  'superstructure_builtup_stair_knee_wall',
  'superstructure_builtup_stair_styrofoam',
  'superstructure_builtup_stairs',
  'superstructure_stair_slab',
  'superstructure_columns_takeoff',
  'superstructure_columns_final'
])

function getSuperstructureRedIToLRowNumbers(formulaData, startIndex) {
  if (!formulaData || !Array.isArray(formulaData)) return new Set()
  const set = new Set()
  formulaData.forEach((f) => {
    if ((f.section || '').toLowerCase() !== 'superstructure') return
    const type = (f.itemType || '').toLowerCase()
    if (RED_I_TO_L_ITEM_TYPES.has(type)) {
      const sheetRow = f.row - startIndex + DATA_START_ROW - 1
      set.add(sheetRow)
    }
  })
  return set
}

/** Map calculation sheet 1-based row to Superstructure Breakdown sheet row. */
function calcRowToSheetRow(calcRow, startIndex) {
  return calcRow - startIndex + DATA_START_ROW - 1
}

/**
 * Apply I–M (FT, SQ FT, CY, QTY, LBS) formulas to Superstructure Breakdown using formulaData.
 * Mirrors ProposalDetail applyFormulasAndStyles for section === 'superstructure' with row mapping.
 */
function applySuperstructureFormulas(spreadsheet, pfx, formulaData, startIndex) {
  if (!formulaData || !Array.isArray(formulaData)) return
  formulaData.forEach((formulaInfo) => {
    const { row, itemType, parsedData, section } = formulaInfo
    if (!row || (section || '').toLowerCase() !== 'superstructure') return
    const r = calcRowToSheetRow(row, startIndex)
    const cell = (col) => `${pfx}${col}${r}`

    try {
      if (itemType === 'superstructure_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('J'))
        spreadsheet.updateCell({ formula: `=SUM(L${sFirst}:L${sLast})` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_slab_steps_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(I${sFirst}:I${sLast})` }, cell('I'))
        spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('J'))
        spreadsheet.updateCell({ formula: `=SUM(L${sFirst}:L${sLast})` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_lw_concrete_fill_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('J'))
        spreadsheet.updateCell({ formula: `=SUM(L${sFirst}:L${sLast})` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_item') {
        const parsed = parsedData || formulaInfo
        const heightFormula = parsed?.parsed?.heightFormula
        const heightValue = parsed?.parsed?.heightValue
        if (heightFormula) spreadsheet.updateCell({ formula: `=${heightFormula}` }, cell('H'))
        if (heightValue != null) spreadsheet.updateCell({ value: heightValue }, cell('H'))
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*H${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_slab_step') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('I'))
        spreadsheet.updateCell({ formula: `=I${r}*H${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*G${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_lw_concrete_fill') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*H${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_topping_slab') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*H${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_topping_slab_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('J'))
        spreadsheet.updateCell({ formula: `=SUM(L${sFirst}:L${sLast})` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_thermal_break') {
        const parsed = parsedData || formulaInfo
        const qty = parsed?.parsed?.qty
        if (qty != null) {
          spreadsheet.updateCell({ formula: `=C${r}*E${r}` }, cell('I'))
        } else {
          spreadsheet.updateCell({ formula: `=C${r}` }, cell('I'))
        }
        return
      }
      if (itemType === 'superstructure_thermal_break_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(I${sFirst}:I${sLast})` }, cell('I'))
        return
      }
      if (itemType === 'superstructure_raised_knee_wall') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('I'))
        spreadsheet.updateCell({ formula: `=I${r}*H${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*G${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_raised_styrofoam') {
        const { takeoffRefRow, heightRefRow } = formulaInfo
        const refC = takeoffRefRow != null ? calcRowToSheetRow(takeoffRefRow, startIndex) : null
        const refH = heightRefRow != null ? calcRowToSheetRow(heightRefRow, startIndex) : null
        if (refC != null) spreadsheet.updateCell({ formula: `=C${refC}` }, cell('C'))
        if (refH != null) spreadsheet.updateCell({ formula: `=H${refH}` }, cell('H'))
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*H${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_raised_slab') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*H${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_raised_slab_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(I${sFirst}:I${sLast})` }, cell('I'))
        spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('J'))
        spreadsheet.updateCell({ formula: `=SUM(L${sFirst}:L${sLast})` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_builtup_knee_wall') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('I'))
        spreadsheet.updateCell({ formula: `=I${r}*H${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*G${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_builtup_styrofoam') {
        const { takeoffRefRow, heightRefRow } = formulaInfo
        const refC = takeoffRefRow != null ? calcRowToSheetRow(takeoffRefRow, startIndex) : null
        const refH = heightRefRow != null ? calcRowToSheetRow(heightRefRow, startIndex) : null
        if (refC != null) spreadsheet.updateCell({ formula: `=C${refC}` }, cell('C'))
        if (refH != null) spreadsheet.updateCell({ formula: `=H${refH}` }, cell('H'))
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*H${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_builtup_slab') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*H${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_builtup_slab_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(I${sFirst}:I${sLast})` }, cell('I'))
        spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('J'))
        spreadsheet.updateCell({ formula: `=SUM(L${sFirst}:L${sLast})` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_builtup_ramps_knee_wall') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('I'))
        spreadsheet.updateCell({ formula: `=I${r}*H${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*G${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_builtup_ramps_styrofoam') {
        const { takeoffRefRow, heightRefRow } = formulaInfo
        const refC = takeoffRefRow != null ? calcRowToSheetRow(takeoffRefRow, startIndex) : null
        const refH = heightRefRow != null ? calcRowToSheetRow(heightRefRow, startIndex) : null
        if (refC != null) spreadsheet.updateCell({ formula: `=C${refC}` }, cell('C'))
        if (refH != null) spreadsheet.updateCell({ formula: `=H${refH}` }, cell('H'))
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*H${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_builtup_ramps_ramp') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*H${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_builtup_stair_knee_wall') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('I'))
        spreadsheet.updateCell({ formula: `=I${r}*H${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*G${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_builtup_stair_styrofoam') {
        const { takeoffJSumFirstRow, takeoffJSumLastRow, heightRefRow } = formulaInfo
        const sFirst = takeoffJSumFirstRow != null ? calcRowToSheetRow(takeoffJSumFirstRow, startIndex) : null
        const sLast = takeoffJSumLastRow != null ? calcRowToSheetRow(takeoffJSumLastRow, startIndex) : null
        if (sFirst != null && sLast != null) spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('C'))
        const refH = heightRefRow != null ? calcRowToSheetRow(heightRefRow, startIndex) : null
        if (refH != null) spreadsheet.updateCell({ formula: `=H${refH}` }, cell('H'))
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*H${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_builtup_stairs') {
        spreadsheet.updateCell({ formula: '=11/12' }, cell('F'))
        spreadsheet.updateCell({ formula: '=7/12' }, cell('H'))
        spreadsheet.updateCell({ formula: `=C${r}*G${r}*F${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*H${r}/27` }, cell('L'))
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_stair_slab') {
        const { takeoffRefRow, widthRefRow, heightValue } = formulaInfo
        const refC = takeoffRefRow != null ? calcRowToSheetRow(takeoffRefRow, startIndex) : null
        const refG = widthRefRow != null ? calcRowToSheetRow(widthRefRow, startIndex) : null
        if (refC != null) spreadsheet.updateCell({ formula: `=C${refC}*1.3` }, cell('C'))
        if (refG != null) spreadsheet.updateCell({ formula: `=G${refG}` }, cell('G'))
        if (heightValue != null) spreadsheet.updateCell({ value: heightValue }, cell('H'))
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('I'))
        spreadsheet.updateCell({ formula: `=I${r}*H${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*G${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_somd_gen1') {
        const { firstDataRow, lastDataRow, heightFormula } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(C${sFirst}:C${sLast})` }, cell('C'))
        if (heightFormula) spreadsheet.updateCell({ formula: `=${heightFormula}` }, cell('H'))
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=(J${r}*H${r})/27` }, cell('L'))
        spreadsheet.cellFormat({ backgroundColor: '#DDEBF7' }, cell('H'))
        return
      }
      if (itemType === 'superstructure_somd_gen2') {
        const { takeoffRefRow, heightValue } = formulaInfo
        const refR = takeoffRefRow != null ? calcRowToSheetRow(takeoffRefRow, startIndex) : null
        if (refR != null) spreadsheet.updateCell({ formula: `=C${refR}` }, cell('C'))
        if (heightValue != null) spreadsheet.updateCell({ value: heightValue }, cell('H'))
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=(J${r}*H${r})/27/2` }, cell('L'))
        spreadsheet.cellFormat({ backgroundColor: '#DDEBF7' }, cell('H'))
        return
      }
      if (itemType === 'superstructure_somd_sum') {
        const { gen1Row, gen2Row } = formulaInfo
        const g1 = gen1Row != null ? calcRowToSheetRow(gen1Row, startIndex) : null
        const g2 = gen2Row != null ? calcRowToSheetRow(gen2Row, startIndex) : null
        if (g1 != null) spreadsheet.updateCell({ formula: `=J${g1}` }, cell('J'))
        if (g1 != null && g2 != null) spreadsheet.updateCell({ formula: `=SUM(L${g1}:L${g2})` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_columns_takeoff') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_columns_final') {
        const { takeoffRefRow } = formulaInfo
        const refC = takeoffRefRow != null ? calcRowToSheetRow(takeoffRefRow, startIndex) : null
        if (refC != null) spreadsheet.updateCell({ formula: `=C${refC}` }, cell('C'))
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_concrete_hanger') {
        spreadsheet.updateCell({ formula: `=G${r}*F${r}*C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*H${r}/27` }, cell('L'))
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_concrete_hanger_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('J'))
        spreadsheet.updateCell({ formula: `=SUM(L${sFirst}:L${sLast})` }, cell('L'))
        spreadsheet.updateCell({ formula: `=SUM(M${sFirst}:M${sLast})` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_shear_wall') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('I'))
        spreadsheet.updateCell({ formula: `=I${r}*H${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*G${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_parapet_wall') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('I'))
        spreadsheet.updateCell({ formula: `=I${r}*H${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*G${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_shear_walls_sum' || itemType === 'superstructure_parapet_walls_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(I${sFirst}:I${sLast})` }, cell('I'))
        spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('J'))
        spreadsheet.updateCell({ formula: `=SUM(L${sFirst}:L${sLast})` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_concrete_post') {
        spreadsheet.updateCell({ formula: `=G${r}*H${r}*C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*F${r}/27` }, cell('L'))
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_concrete_encasement') {
        spreadsheet.updateCell({ formula: `=G${r}*H${r}*C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*F${r}/27` }, cell('L'))
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_drop_panel_bracket') {
        spreadsheet.updateCell({ formula: `=G${r}*H${r}*C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*F${r}/27` }, cell('L'))
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_drop_panel_h') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*H${r}/27` }, cell('L'))
        spreadsheet.updateCell({ formula: `=E${r}` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_builtup_stair_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(I${sFirst}:I${sLast})` }, cell('I'))
        spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('J'))
        spreadsheet.updateCell({ formula: `=SUM(L${sFirst}:L${sLast})` }, cell('L'))
        spreadsheet.updateCell({ formula: `=SUM(M${sFirst}:M${sLast})` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_builtup_ramps_knee_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        if (firstDataRow != null && lastDataRow != null) {
          const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
          const sLast = calcRowToSheetRow(lastDataRow, startIndex)
          spreadsheet.updateCell({ formula: `=SUM(I${sFirst}:I${sLast})` }, cell('I'))
          spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('J'))
          spreadsheet.updateCell({ formula: `=SUM(L${sFirst}:L${sLast})` }, cell('L'))
        }
        return
      }
      if (itemType === 'superstructure_builtup_ramps_styro_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        if (firstDataRow != null && lastDataRow != null) {
          const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
          const sLast = calcRowToSheetRow(lastDataRow, startIndex)
          spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('J'))
          spreadsheet.updateCell({ formula: `=SUM(L${sFirst}:L${sLast})` }, cell('L'))
        }
        return
      }
      if (itemType === 'superstructure_builtup_ramps_ramp_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        if (firstDataRow != null && lastDataRow != null) {
          const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
          const sLast = calcRowToSheetRow(lastDataRow, startIndex)
          spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('J'))
          spreadsheet.updateCell({ formula: `=SUM(L${sFirst}:L${sLast})` }, cell('L'))
        }
        return
      }
      if (itemType === 'superstructure_beam') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('I'))
        spreadsheet.updateCell({ formula: `=I${r}*H${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*G${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_drop_panel_sum' || itemType === 'superstructure_concrete_post_sum' || itemType === 'superstructure_concrete_encasement_sum' || itemType === 'superstructure_beams_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(I${sFirst}:I${sLast})` }, cell('I'))
        spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('J'))
        spreadsheet.updateCell({ formula: `=SUM(L${sFirst}:L${sLast})` }, cell('L'))
        spreadsheet.updateCell({ formula: `=SUM(M${sFirst}:M${sLast})` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_curb') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('I'))
        spreadsheet.updateCell({ formula: `=I${r}*H${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*G${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_curbs_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(I${sFirst}:I${sLast})` }, cell('I'))
        spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('J'))
        spreadsheet.updateCell({ formula: `=SUM(L${sFirst}:L${sLast})` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_concrete_pad') {
        const unit = (formulaInfo.parsedData?.unit || '').toString().toLowerCase().trim()
        if (unit.includes('sq') || unit === 'sf') {
          spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
          spreadsheet.updateCell({ formula: `=J${r}*H${r}/27` }, cell('L'))
          const hasQty = formulaInfo.parsedData?.parsed?.qty != null
          if (hasQty) spreadsheet.updateCell({ formula: `=E${r}` }, cell('M'))
        } else {
          spreadsheet.updateCell({ formula: `=C${r}` }, cell('M'))
        }
        return
      }
      if (itemType === 'superstructure_concrete_pad_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('J'))
        spreadsheet.updateCell({ formula: `=SUM(L${sFirst}:L${sLast})` }, cell('L'))
        spreadsheet.updateCell({ formula: `=SUM(M${sFirst}:M${sLast})` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_non_shrink_grout') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_non_shrink_grout_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(M${sFirst}:M${sLast})` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_repair_scope') {
        const subType = formulaInfo.parsedData?.parsed?.itemSubType
        if (subType === 'wall') spreadsheet.updateCell({ formula: `=C${r}` }, cell('I'))
        else if (subType === 'slab') spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        else if (subType === 'column') spreadsheet.updateCell({ formula: `=C${r}` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_repair_scope_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(I${sFirst}:I${sLast})` }, cell('I'))
        spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('J'))
        return
      }
      if (itemType === 'superstructure_manual_cip_stairs_landing') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*H${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_manual_cip_stairs_stair') {
        spreadsheet.updateCell({ formula: '=11/12' }, cell('F'))
        spreadsheet.updateCell({ formula: '=7/12' }, cell('H'))
        spreadsheet.updateCell({ formula: `=C${r}*G${r}*F${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*H${r}/27` }, cell('L'))
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_manual_cip_stairs_slab') {
        const { stairsRowNum, slabCMultiplier } = formulaInfo
        const sRow = stairsRowNum != null ? calcRowToSheetRow(stairsRowNum, startIndex) : null
        if (sRow != null) spreadsheet.updateCell({ formula: `=G${sRow}` }, cell('G'))
        if (sRow != null && slabCMultiplier != null && slabCMultiplier !== 1) {
          spreadsheet.updateCell({ formula: `=C${sRow}*${slabCMultiplier}` }, cell('C'))
        } else if (sRow != null) {
          spreadsheet.updateCell({ formula: `=C${sRow}` }, cell('C'))
        }
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('I'))
        spreadsheet.updateCell({ formula: `=I${r}*H${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*G${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_manual_cip_stairs_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(I${sFirst}:I${sLast})` }, cell('I'))
        spreadsheet.updateCell({ formula: `=SUM(J${sFirst}:J${sLast})` }, cell('J'))
        spreadsheet.updateCell({ formula: `=SUM(L${sFirst}:L${sLast})` }, cell('L'))
        spreadsheet.updateCell({ formula: `=SUM(M${sFirst}:M${sLast})` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_infilled_landing_item') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=(J${r}*H${r})/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_infilled_landing_second') {
        const { landingRowNum } = formulaInfo
        const landR = landingRowNum != null ? calcRowToSheetRow(landingRowNum, startIndex) : null
        if (landR != null) spreadsheet.updateCell({ formula: `=C${landR}` }, cell('C'))
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=(J${r}*H${r})/27/2` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_infilled_landing_sum') {
        const { landingRowNum, firstDataRow, lastDataRow } = formulaInfo
        const landR = landingRowNum != null ? calcRowToSheetRow(landingRowNum, startIndex) : null
        if (landR != null) spreadsheet.updateCell({ formula: `=J${landR}` }, cell('J'))
        if (firstDataRow != null && lastDataRow != null) {
          const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
          const sLast = calcRowToSheetRow(lastDataRow, startIndex)
          spreadsheet.updateCell({ formula: `=SUM(L${sFirst}:L${sLast})` }, cell('L'))
        }
        return
      }
      if (itemType === 'superstructure_manual_infilled_stair') {
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${r}*H${r}/27` }, cell('L'))
        return
      }
      if (itemType === 'superstructure_ele_item') {
        const takeoffSource = formulaInfo.parsedData?.takeoffSource ?? formulaInfo.takeoffSource
        const refR = takeoffSource != null ? calcRowToSheetRow(takeoffSource, startIndex) : null
        if (refR != null) spreadsheet.updateCell({ formula: `=C${refR}` }, cell('C'))
        spreadsheet.updateCell({ formula: `=C${r}` }, cell('M'))
        return
      }
      if (itemType === 'superstructure_ele_sum') {
        const { firstDataRow, lastDataRow } = formulaInfo
        const sFirst = calcRowToSheetRow(firstDataRow, startIndex)
        const sLast = calcRowToSheetRow(lastDataRow, startIndex)
        spreadsheet.updateCell({ formula: `=SUM(M${sFirst}:M${sLast})` }, cell('M'))
        return
      }
    } catch (e) {
      // ignore per-row errors
    }
  })
}

/**
 * @param {object} spreadsheet - Syncfusion spreadsheet instance
 * @param {object} ctx - { calculationData, formulaData, rawData, currentRow, pfx, ... }
 * @returns {number} next row index after this section
 */
export function buildSuperstructureBreakdown(spreadsheet, ctx) {
  const pfx = ctx.pfx ?? 'Superstructure Breakdown!'
  const calculationData = ctx.calculationData || []
  const formulaData = ctx.formulaData || []

  const superstructureFormulas = (formulaData || []).filter(
    (f) => (f.section || '').toLowerCase() === 'superstructure' && f.row
  )
  const firstFormulaRow = superstructureFormulas.length > 0
    ? Math.min(...superstructureFormulas.map((f) => f.row))
    : 1
  let startIndex = firstFormulaRow - 1
  if (startIndex > 0 && calculationData[startIndex - 1]) {
    const particulars = (calculationData[startIndex - 1][1] != null) ? String(calculationData[startIndex - 1][1]).trim() : ''
    if (particulars.endsWith(':')) startIndex = startIndex - 1
  }
  const superstructureRows = getSuperstructureRows(calculationData, formulaData)
  const sumRowNumbers = getSuperstructureSumRowNumbers(formulaData, startIndex)
  const formulaValueRowNumbers = getSuperstructureFormulaValueRowNumbers(formulaData, startIndex)
  const redIToLRowNumbers = getSuperstructureRedIToLRowNumbers(formulaData, startIndex)
  const sheetIndex = ctx.sheetIndex ?? 4

  if (superstructureRows.length === 0) {
    spreadsheet.updateCell({ value: 'Superstructure Breakdown' }, `${pfx}B1`)
    for (let c = 0; c < NUM_COLS; c++) {
      spreadsheet.updateCell({ value: COLUMN_HEADERS[c] }, `${pfx}${colLetter(c)}1`)
    }
    spreadsheet.cellFormat(
      { ...ROWS_1_3_FORMAT, color: '#000000', backgroundColor: B_TO_M_HEADER_BG, textDecoration: 'underline' },
      `${pfx}B1`
    )
    spreadsheet.cellFormat(
      { ...ROWS_1_3_FORMAT, color: '#000000', backgroundColor: ESTIMATE_HEADER_BG },
      `${pfx}A1`
    )
    spreadsheet.cellFormat(
      { ...ROWS_1_3_FORMAT, color: '#000000', backgroundColor: B_TO_M_HEADER_BG },
      `${pfx}B1:${colLetter(NUM_COLS - 1)}1`
    )
    try {
      for (let c = 0; c < NUM_COLS; c++) {
        const w = c === 1 ? COL_B_WIDTH_PX : COL_REST_WIDTH_PX
        spreadsheet.setColWidth(w, c, sheetIndex)
      }
    } catch (e) { /* ignore */ }
    return 1
  }

  // Row 1: column headers – same as other tabs (A1 yellow, B1:M1 green, black text)
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

  // Row 2: empty
  spreadsheet.cellFormat({ ...BASE_FORMAT, backgroundColor: 'white' }, `${pfx}A2:${colLetter(NUM_COLS - 1)}2`)

  // Row 3: section header "Superstructure" and column labels (same font/alignment as Capstone)
  spreadsheet.updateCell({ value: 'Superstructure' }, `${pfx}A3`)
  for (let c = 1; c < COLUMN_HEADERS.length; c++) {
    spreadsheet.updateCell({ value: COLUMN_HEADERS[c] }, `${pfx}${colLetter(c)}3`)
  }
  spreadsheet.cellFormat(
    { ...BASE_FORMAT, fontWeight: 'bold', textAlign: 'center', color: '#000000', backgroundColor: SECTION_ROW_BG },
    `${pfx}A3:${colLetter(NUM_COLS - 1)}3`
  )

  // Apply base format to entire used range (headers + data)
  const totalRows = DATA_START_ROW - 1 + superstructureRows.length
  for (let rowNum = 1; rowNum <= totalRows; rowNum++) {
    for (let c = 0; c < NUM_COLS; c++) {
      try {
        spreadsheet.cellFormat(BASE_FORMAT, `${pfx}${colLetter(c)}${rowNum}`)
      } catch (e) { /* ignore */ }
    }
  }

  // Write data rows: superstructureRows[0] -> sheet row DATA_START_ROW, etc.
  const HEIGHT_COL = 7
  for (let r = 0; r < superstructureRows.length; r++) {
    const row = superstructureRows[r]
    const rowNum = r + DATA_START_ROW
    const maxC = Math.min(row.length, NUM_COLS)
    for (let c = 0; c < maxC; c++) {
      let val = row[c]
      if (c === HEIGHT_COL && (val === '' || val == null || val === undefined)) {
        const inferred = parseHeightFromParticulars(row[1])
        if (inferred != null) val = inferred
      }
      const cellRef = `${pfx}${colLetter(c)}${rowNum}`
      if (val !== '' && val !== null && val !== undefined) {
        try {
          const displayVal = c >= 2 ? toTwoDecimals(val) : val
          spreadsheet.updateCell({ value: displayVal }, cellRef)
        } catch (e) { /* ignore */ }
      }
    }
  }

  // Style subsection headers (B ends with ':') and sum rows
  for (let r = 0; r < superstructureRows.length; r++) {
    const row = superstructureRows[r]
    const rowNum = r + DATA_START_ROW
    const particulars = (row[1] != null && row[1] !== '') ? String(row[1]).trim() : ''
    const isSubsectionHeader = particulars.endsWith(':')
    const isSumRow = sumRowNumbers.has(rowNum)

    if (isSubsectionHeader) {
      spreadsheet.cellFormat(
        { ...BASE_FORMAT, fontWeight: 'bold', color: '#000000', textAlign: 'left' },
        `${pfx}B${rowNum}`
      )
    } else if (isSumRow) {
      // Same as Capstone: red bold only on I, J, L (calculated sum columns)
      for (let c = 2; c < NUM_COLS; c++) {
        spreadsheet.cellFormat({ ...BASE_FORMAT, format: '#,##0.00', textAlign: 'right' }, `${pfx}${colLetter(c)}${rowNum}`)
      }
      spreadsheet.cellFormat({ ...BASE_FORMAT, color: RED_TEXT, fontWeight: 'bold', textAlign: 'right' }, `${pfx}I${rowNum}`)
      spreadsheet.cellFormat({ ...BASE_FORMAT, color: RED_TEXT, fontWeight: 'bold', textAlign: 'right' }, `${pfx}J${rowNum}`)
      spreadsheet.cellFormat({ ...BASE_FORMAT, color: RED_TEXT, fontWeight: 'bold', textAlign: 'right' }, `${pfx}L${rowNum}`)
    } else {
      // Data row: Particulars (B) – red when takeoff > 0 or row has C from formula (Styrofoam, stair slab, columns final, etc.), black otherwise
      if (row[1] !== '' && row[1] != null && row[1] !== undefined) {
        const bColor = formulaValueRowNumbers.has(rowNum) ? RED_TEXT : getColumnBColorForRow(row)
        spreadsheet.cellFormat(
          { ...BASE_FORMAT, color: bColor, textAlign: 'left' },
          `${pfx}B${rowNum}`
        )
      }
      for (let c = 2; c < NUM_COLS; c++) {
        spreadsheet.cellFormat(
          { ...BASE_FORMAT, format: '#,##0.00', textAlign: 'right' },
          `${pfx}${colLetter(c)}${rowNum}`
        )
      }
      // Raised slab / Built-up slab / Columns: I, J, K, L in red (same as sum rows for those columns)
      if (redIToLRowNumbers.has(rowNum)) {
        spreadsheet.cellFormat({ ...BASE_FORMAT, color: RED_TEXT, format: '#,##0.00', textAlign: 'right' }, `${pfx}I${rowNum}`)
        spreadsheet.cellFormat({ ...BASE_FORMAT, color: RED_TEXT, format: '#,##0.00', textAlign: 'right' }, `${pfx}J${rowNum}`)
        spreadsheet.cellFormat({ ...BASE_FORMAT, color: RED_TEXT, format: '#,##0.00', textAlign: 'right' }, `${pfx}K${rowNum}`)
        spreadsheet.cellFormat({ ...BASE_FORMAT, color: RED_TEXT, format: '#,##0.00', textAlign: 'right' }, `${pfx}L${rowNum}`)
      }
    }
  }

  // Apply FT, SQ FT, CY, QTY, LBS formulas (I–M) from formulaData with row mapping
  applySuperstructureFormulas(spreadsheet, pfx, formulaData, startIndex)

  // Column widths and row heights
  try {
    for (let c = 0; c < NUM_COLS; c++) {
      const w = c === 1 ? COL_B_WIDTH_PX : COL_REST_WIDTH_PX
      spreadsheet.setColWidth(w, c, sheetIndex)
    }
    for (let r = 0; r < totalRows; r++) {
      spreadsheet.setRowHeight(ROW_HEIGHT_PX, r, sheetIndex)
    }
  } catch (e) { /* ignore */ }

  // Red vertical line left of column I
  const lastRow = totalRows
  if (lastRow > 0) {
    try {
      spreadsheet.cellFormat({ borderLeft: '3px solid #FF0000' }, `${pfx}I1:I${lastRow}`)
    } catch (e) { /* ignore */ }
  }

  return 1
}

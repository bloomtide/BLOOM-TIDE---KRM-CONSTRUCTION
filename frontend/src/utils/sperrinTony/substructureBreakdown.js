/**
 * Sperrin Tony – Substructure Breakdown tab.
 * Filled from calculation sheet Foundation section. Same column layout as Earthwork/SOE (I–M = FT, SQ FT, CY, QTY, LBS).
 */

import { generateFoundationFormulas } from '../processors/foundationProcessor'

const FOUNDATION_SECTION_NAME = 'Foundation'
const SUBSTRUCTURE_LABEL = 'Substructure'
// Section names that follow Foundation in template order (stop when we hit one)
const SECTIONS_AFTER_FOUNDATION = ['Waterproofing', 'Trenching', 'Superstructure', 'B.P.P. Alternate #2 scope', 'Civil / Sitework', 'Masonry', 'Fireproofing']

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

// I–M on sheet: FT, SQ FT, CY, QTY, LBS (calculation 8=FT, 9=SQ FT, 10=LBS, 11=CY, 12=QTY)
const SUBSTRUCTURE_I_M_MAP = { 8: 'I', 9: 'J', 10: 'M', 11: 'K', 12: 'L' }
function substructureColLetter(c) {
  if (c >= 8 && c <= 12) return SUBSTRUCTURE_I_M_MAP[c]
  return colLetter(c)
}

/**
 * Extract Foundation section rows from calculation data.
 * Returns { rows: [headerRow, emptyRow, ...foundationRows], foundationStartIdx (0-based) }.
 */
function getSubstructureRows(calculationData) {
  if (!calculationData || !Array.isArray(calculationData) || calculationData.length < 2) {
    return { rows: [], foundationStartIdx: -1 }
  }
  const rows = calculationData
  let foundationStartIdx = -1
  let endIdx = rows.length
  for (let i = 2; i < rows.length; i++) {
    const sectionName = (rows[i][0] != null && rows[i][0] !== '') ? String(rows[i][0]).trim() : ''
    if (sectionName === FOUNDATION_SECTION_NAME && foundationStartIdx === -1) {
      foundationStartIdx = i
    }
    if (foundationStartIdx >= 0 && sectionName && SECTIONS_AFTER_FOUNDATION.includes(sectionName)) {
      endIdx = i
      break
    }
  }
  if (foundationStartIdx < 0) return { rows: [], foundationStartIdx: -1 }
  const headerRow = rows[1] && rows[1].length ? rows[1].slice(0, NUM_COLS) : COLUMN_HEADERS.slice()
  const emptyRow = Array(NUM_COLS).fill('')
  const foundationRows = rows.slice(foundationStartIdx, endIdx)
  return {
    rows: [headerRow, emptyRow, ...foundationRows],
    foundationStartIdx
  }
}

/** Convert calculation 1-based row to Substructure sheet 1-based row. */
function calcRowToSheetRow(calcRow1Based, foundationStartIdx) {
  return calcRow1Based - foundationStartIdx + 2
}

function getSubstructureSumRowNumbers(formulaData, foundationStartIdx) {
  if (!formulaData || !Array.isArray(formulaData) || foundationStartIdx < 0) return new Set()
  const set = new Set()
  formulaData.forEach((f) => {
    if (!f.row || (f.section || '').toLowerCase() !== 'foundation') return
    const type = (f.itemType || '').toLowerCase()
    if (type.includes('sum')) {
      set.add(calcRowToSheetRow(f.row, foundationStartIdx))
    }
  })
  return set
}

function getColumnBColor(calculationData, calcRow) {
  if (!calculationData || calcRow < 1) return '#000000'
  const rowIndex = calcRow - 1
  if (rowIndex >= calculationData.length) return '#000000'
  const takeoffValue = calculationData[rowIndex][2]
  if (takeoffValue == null || takeoffValue === '' || takeoffValue === 0 || takeoffValue === '0') return '#000000'
  if (typeof takeoffValue === 'string' && takeoffValue.trim().startsWith('=')) return '#000000'
  const n = parseFloat(takeoffValue)
  return (n > 0) ? RED_TEXT : '#000000'
}

/** Rewrite formula string: replace all calculation row numbers with sheet rows (so e.g. H462*C462 becomes H15*C15 on this sheet). */
function formulaCalcRowsToSheetRows(formulaStr, toRowFn) {
  if (formulaStr == null || typeof formulaStr !== 'string') return formulaStr
  return formulaStr.replace(/([A-Z]+)(\d+)/g, (_, col, num) => col + toRowFn(parseInt(num, 10)))
}

/**
 * Apply foundation formulas to Substructure Breakdown sheet.
 * Mirrors ProposalDetail applyFormulasAndStyles for section === 'foundation'; row refs use sheet rows.
 */
function applySubstructureFormulas(spreadsheet, pfx, formulaData, foundationStartIdx, calculationData) {
  if (!formulaData || !Array.isArray(formulaData) || foundationStartIdx < 0) return
  const toRow = (calcRow) => calcRowToSheetRow(calcRow, foundationStartIdx)

  formulaData.forEach((formulaInfo) => {
    const { row: calcRow, itemType, parsedData, section, subsectionName } = formulaInfo
    if (!calcRow || (section || '').toLowerCase() !== 'foundation') return
    const row = toRow(calcRow)
    const cell = (col) => `${pfx}${col}${row}`

    try {
      if (itemType === 'foundation_piles_misc') {
        spreadsheet.cellFormat({ color: RED_TEXT }, cell('B'))
        return
      }
      if (itemType === 'foundation_stairs_on_grade_heading') {
        spreadsheet.cellFormat({ fontWeight: 'bold', fontStyle: 'italic', textDecoration: 'underline', color: '#000000' }, cell('B'))
        return
      }
      if (itemType === 'foundation_stairs_on_grade_landing') {
        spreadsheet.updateCell({ formula: `=C${row}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
        spreadsheet.cellFormat({ color: getColumnBColor(calculationData, calcRow) }, cell('B'))
        spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('J'))
        spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('K'))
        return
      }
      if (itemType === 'foundation_stairs_on_grade_stairs') {
        spreadsheet.updateCell({ formula: '=11/12' }, cell('F'))
        const heightFromName = parsedData?.parsed?.heightFromName
        if (heightFromName != null) {
          spreadsheet.updateCell({ value: heightFromName }, cell('H'))
        } else {
          spreadsheet.updateCell({ formula: '=7/12' }, cell('H'))
        }
        spreadsheet.updateCell({ formula: `=C${row}*G${row}*F${row}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
        spreadsheet.updateCell({ formula: `=C${row}` }, cell('L'))
        spreadsheet.cellFormat({ color: getColumnBColor(calculationData, calcRow) }, cell('B'))
        return
      }
      if (itemType === 'foundation_stairs_on_grade_stair_slab') {
        const { stairsRefRow } = formulaInfo
        if (stairsRefRow) {
          const refSheetRow = toRow(stairsRefRow)
          spreadsheet.updateCell({ formula: `=1.3*C${refSheetRow}` }, cell('C'))
          spreadsheet.updateCell({ formula: `=G${refSheetRow}` }, cell('G'))
        }
        spreadsheet.updateCell({ formula: `=C${row}` }, cell('I'))
        spreadsheet.updateCell({ formula: `=I${row}*H${row}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${row}*G${row}/27` }, cell('K'))
        return
      }
      if (itemType === 'foundation_stairs_on_grade_sum') {
        const { sumRanges } = formulaInfo
        if (sumRanges && sumRanges.length > 0) {
          const sumI = sumRanges.map(([f, l]) => `SUM(I${toRow(f)}:I${toRow(l)})`).join('+')
          const sumJ = sumRanges.map(([f, l]) => `SUM(J${toRow(f)}:J${toRow(l)})`).join('+')
          const sumK = sumRanges.map(([f, l]) => `SUM(K${toRow(f)}:K${toRow(l)})`).join('+')
          const sumL = sumRanges.map(([f, l]) => `SUM(L${toRow(f)}:L${toRow(l)})`).join('+')
          spreadsheet.updateCell({ formula: `=${sumI}` }, cell('I'))
          spreadsheet.updateCell({ formula: `=${sumJ}` }, cell('J'))
          spreadsheet.updateCell({ formula: `=${sumK}` }, cell('K'))
          spreadsheet.updateCell({ formula: `=${sumL}` }, cell('L'))
          spreadsheet.cellFormat({ color: RED_TEXT }, cell('I'))
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('J'))
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('K'))
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('L'))
        }
        return
      }
      if (itemType === 'foundation_sum') {
        const { firstDataRow, lastDataRow, subsectionName: subName, isDualDiameter, excludeISum, excludeJSum, excludeKSum, cySumOnly, lSumRange, matSlabCombinedSum, lastDataRowForJ, excludeLSum } = formulaInfo
        const first = toRow(firstDataRow)
        const last = toRow(lastDataRow)

        if (matSlabCombinedSum && subName === 'Mat slab') {
          const jEnd = lastDataRowForJ != null ? toRow(lastDataRowForJ) : last
          spreadsheet.updateCell({ formula: `=SUM(J${first}:J${jEnd})` }, cell('J'))
          spreadsheet.updateCell({ formula: `=SUM(K${first}:K${last})` }, cell('K'))
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('J'))
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('K'))
          return
        }
        if (!excludeISum) {
          const ftSumSubsections = ['Piles', 'Helical foundation pile', 'Driven foundation pile', 'Drilled displacement pile', 'CFA pile', 'Grade beams', 'Tie beam', 'Strap beams', 'Thickened slab', 'Corbel', 'Linear Wall', 'Foundation Wall', 'Retaining walls', 'Barrier wall', 'Drilled foundation pile', 'Strip Footings', 'Stem wall', 'Elevator Pit', 'Service elevator pit', 'Detention tank', 'Duplex sewage ejector pit', 'Deep sewage ejector pit', 'Sump pump pit', 'Grease trap', 'House trap', 'SOG', 'Ramp on grade', 'Stairs on grade Stairs', 'Electric conduit']
          if (ftSumSubsections.includes(subName)) {
            spreadsheet.updateCell({ formula: `=SUM(I${first}:I${last})` }, cell('I'))
            spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('I'))
          }
        }
        if (subName === 'Drilled foundation pile' && isDualDiameter) {
          spreadsheet.updateCell({ formula: `=SUM(I${first}:I${last})` }, cell('I'))
          spreadsheet.updateCell({ formula: `=SUM(J${first}:J${last})` }, cell('J'))
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('I'))
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('J'))
        }
        if (!excludeJSum && !cySumOnly) {
          const sqFtSubsections = ['Piles', 'Pile caps', 'Isolated Footings', 'Pilaster', 'Pier', 'Strip Footings', 'Grade beams', 'Tie beam', 'Strap beams', 'Thickened slab', 'Corbel', 'Linear Wall', 'Foundation Wall', 'Retaining walls', 'Barrier wall', 'Stem wall', 'Elevator Pit', 'Service elevator pit', 'Detention tank', 'Duplex sewage ejector pit', 'Deep sewage ejector pit', 'Sump pump pit', 'Grease trap', 'House trap', 'Mat slab', 'SOG', 'Ramp on grade', 'Stairs on grade Stairs']
          if (sqFtSubsections.includes(subName)) {
            spreadsheet.updateCell({ formula: `=SUM(J${first}:J${last})` }, cell('J'))
            spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('J'))
          }
        }
        const lbsSubsections = ['Piles', 'Drilled foundation pile', 'Helical foundation pile', 'Driven foundation pile', 'Drilled displacement pile']
        if (!excludeKSum && lbsSubsections.includes(subName)) {
          spreadsheet.updateCell({ formula: `=SUM(M${first}:M${last})` }, cell('M'))
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('M'))
        }
        const qtySubsections = ['Piles', 'Drilled foundation pile', 'Helical foundation pile', 'Driven foundation pile', 'Drilled displacement pile', 'CFA pile', 'Pile caps', 'Isolated Footings', 'Pilaster', 'Pier', 'Stairs on grade Stairs']
        if (qtySubsections.includes(subName)) {
          spreadsheet.updateCell({ formula: `=SUM(L${first}:L${last})` }, cell('L'))
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('L'))
        }
        if (cySumOnly) {
          spreadsheet.updateCell({ formula: `=SUM(K${first}:K${last})` }, cell('K'))
          spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('K'))
        } else if (!excludeLSum) {
          const cySubsections = ['Piles', 'Pile caps', 'Strip Footings', 'Isolated Footings', 'Pilaster', 'Grade beams', 'Tie beam', 'Strap beams', 'Thickened slab', 'Pier', 'Corbel', 'Linear Wall', 'Foundation Wall', 'Retaining walls', 'Barrier wall', 'Stem wall', 'Elevator Pit', 'Service elevator pit', 'Detention tank', 'Duplex sewage ejector pit', 'Deep sewage ejector pit', 'Sump pump pit', 'Grease trap', 'House trap', 'Mat slab', 'SOG', 'Ramp on grade', 'Stairs on grade Stairs']
          if (cySubsections.includes(subName)) {
            const kFormula = lSumRange
              ? '=SUM(' + lSumRange.replace(/L(\d+)/gi, (_, n) => `K${toRow(Number(n))}`) + ')'
              : `=SUM(K${first}:K${last})`
            spreadsheet.updateCell({ formula: kFormula }, cell('K'))
            spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('K'))
          }
        }
        return
      }
      if (itemType === 'foundation_section_cy_sum') {
        const { sumRows } = formulaInfo
        if (sumRows && sumRows.length > 0) {
          const sumRefs = sumRows.map((r) => `K${toRow(r)}`).join(',')
          spreadsheet.updateCell({ formula: `=SUM(${sumRefs})` }, cell('C'))
        }
        return
      }
      if (itemType === 'foundation_extra_sqft') {
        spreadsheet.updateCell({ formula: `=C${row}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
        spreadsheet.cellFormat({ color: getColumnBColor(calculationData, calcRow) }, cell('B'))
        spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('J'))
        spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('K'))
        return
      }
      if (itemType === 'foundation_extra_ft') {
        spreadsheet.updateCell({ formula: `=C${row}*G${row}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
        spreadsheet.cellFormat({ color: getColumnBColor(calculationData, calcRow) }, cell('B'))
        spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('J'))
        spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('K'))
        spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('L'))
        return
      }
      if (itemType === 'foundation_extra_ea') {
        spreadsheet.updateCell({ formula: `=C${row}*F${row}*G${row}` }, cell('J'))
        spreadsheet.updateCell({ formula: `=J${row}*H${row}/27` }, cell('K'))
        spreadsheet.updateCell({ formula: `=C${row}` }, cell('L'))
        spreadsheet.cellFormat({ color: getColumnBColor(calculationData, calcRow) }, cell('B'))
        spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('J'))
        spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('K'))
        spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('L'))
        return
      }
      if ((itemType === 'elevator_pit' || itemType === 'service_elevator_pit') && parsedData?.parsed?.itemSubType === 'sump_pit') {
        const foundationFormulas = generateFoundationFormulas(itemType, calcRow, parsedData || formulaInfo)
        const f = (str) => (typeof str === 'string' ? formulaCalcRowsToSheetRows(str, toRow) : str)
        if (foundationFormulas.sqFt) spreadsheet.updateCell({ formula: `=${f(foundationFormulas.sqFt)}` }, cell('J'))
        if (foundationFormulas.cy) spreadsheet.updateCell({ formula: `=${f(foundationFormulas.cy)}` }, cell('K'))
        if (foundationFormulas.qtyFinal) spreadsheet.updateCell({ formula: `=${f(foundationFormulas.qtyFinal)}` }, cell('L'))
        spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('J'))
        spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('K'))
        spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('L'))
        return
      }
      if (itemType === 'buttress_final') {
        const refRow = formulaInfo.buttressRow ?? formulaInfo.takeoffRefRow
        if (refRow != null) spreadsheet.updateCell({ formula: `=C${toRow(refRow)}` }, cell('C'))
        spreadsheet.updateCell({ formula: `=C${row}` }, cell('L'))
        spreadsheet.cellFormat({ color: getColumnBColor(calculationData, calcRow) }, cell('B'))
        spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('I'))
        spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('J'))
        spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('K'))
        spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('L'))
        return
      }
      const foundationItemTypes = ['drilled_foundation_pile', 'helical_foundation_pile', 'driven_foundation_pile', 'stelcor_drilled_displacement_pile', 'cfa_pile', 'pile_cap', 'strip_footing', 'isolated_footing', 'pilaster', 'grade_beam', 'tie_beam', 'strap_beam', 'thickened_slab', 'buttress_takeoff', 'buttress_final', 'pier', 'corbel', 'linear_wall', 'foundation_wall', 'retaining_wall', 'barrier_wall', 'stem_wall', 'elevator_pit', 'service_elevator_pit', 'detention_tank', 'duplex_sewage_ejector_pit', 'deep_sewage_ejector_pit', 'sump_pump_pit', 'grease_trap', 'house_trap', 'mat_slab', 'mud_slab_foundation', 'sog', 'rog', 'stairs_on_grade', 'electric_conduit']
      if (foundationItemTypes.includes(itemType)) {
        const foundationFormulas = generateFoundationFormulas(itemType, calcRow, parsedData || formulaInfo)
        const f = (str) => (typeof str === 'string' ? formulaCalcRowsToSheetRows(str, toRow) : str)
        if (foundationFormulas.takeoff) spreadsheet.updateCell({ formula: `=${f(foundationFormulas.takeoff)}` }, cell('C'))
        if (foundationFormulas.length != null) {
          if (typeof foundationFormulas.length === 'string') {
            spreadsheet.updateCell({ formula: `=${f(foundationFormulas.length)}` }, cell('F'))
          } else {
            spreadsheet.updateCell({ value: foundationFormulas.length }, cell('F'))
          }
        }
        if (foundationFormulas.width != null) {
          if (typeof foundationFormulas.width === 'string') {
            spreadsheet.updateCell({ formula: `=${f(foundationFormulas.width)}` }, cell('G'))
          } else {
            spreadsheet.updateCell({ value: foundationFormulas.width }, cell('G'))
          }
        }
        if (foundationFormulas.height != null) {
          if (typeof foundationFormulas.height === 'string') {
            spreadsheet.updateCell({ formula: `=${f(foundationFormulas.height)}` }, cell('H'))
          } else {
            spreadsheet.updateCell({ value: foundationFormulas.height }, cell('H'))
          }
        }
        if (foundationFormulas.qty !== undefined) {
          if (typeof foundationFormulas.qty === 'string') {
            spreadsheet.updateCell({ formula: `=${f(foundationFormulas.qty)}` }, cell('E'))
          } else {
            spreadsheet.updateCell({ value: foundationFormulas.qty }, cell('E'))
          }
        }
        if (foundationFormulas.ft) spreadsheet.updateCell({ formula: `=${f(foundationFormulas.ft)}` }, cell('I'))
        if (foundationFormulas.sqFt) spreadsheet.updateCell({ formula: `=${f(foundationFormulas.sqFt)}` }, cell('J'))
        if (foundationFormulas.sqFt2) spreadsheet.updateCell({ formula: `=${f(foundationFormulas.sqFt2)}` }, cell('J'))
        if (foundationFormulas.lbs) {
          spreadsheet.updateCell({ formula: `=${f(foundationFormulas.lbs)}` }, cell('M'))
          spreadsheet.cellFormat({ format: '#,##0.00', textAlign: 'right', backgroundColor: '#d9e1f2' }, cell('M'))
        }
        if (foundationFormulas.cy) spreadsheet.updateCell({ formula: `=${f(foundationFormulas.cy)}` }, cell('K'))
        if (foundationFormulas.qtyFinal) spreadsheet.updateCell({ formula: `=${f(foundationFormulas.qtyFinal)}` }, cell('L'))
        spreadsheet.cellFormat({ color: getColumnBColor(calculationData, calcRow) }, cell('B'))
        if (itemType === 'buttress_takeoff') {
          spreadsheet.cellFormat({ textDecoration: 'line-through' }, `${pfx}A${row}:${pfx}M${row}`)
        }
        if (itemType === 'electric_conduit') {
          const particulars = (parsedData || formulaInfo)?.particulars || ''
          const p = String(particulars).toLowerCase()
          if (p.includes('trench drain') || p.includes('perforated pipe')) {
            spreadsheet.cellFormat({ color: RED_TEXT, fontWeight: 'bold' }, cell('I'))
          }
        }
      }
    } catch (e) {
      // ignore per-row formula errors
    }
  })
}

function setSubstructureColumnWidthsAndRowHeights(spreadsheet, sheetIndex, rowCount) {
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
 * @param {object} ctx - { calculationData, formulaData, currentRow, pfx, sheetIndex, ... }
 * @returns {number} next row index after this section
 */
export function buildSubstructureBreakdown(spreadsheet, ctx) {
  const pfx = ctx.pfx ?? 'Substructure Breakdown!'
  const calculationData = ctx.calculationData || []
  const formulaData = ctx.formulaData || []
  const sheetIndex = ctx.sheetIndex ?? 3

  const { rows: substructureRows, foundationStartIdx } = getSubstructureRows(calculationData)
  const sumRowNumbers = getSubstructureSumRowNumbers(formulaData, foundationStartIdx)

  if (substructureRows.length === 0) {
    spreadsheet.updateCell({ value: 'Substructure Breakdown' }, `${pfx}B1`)
    spreadsheet.cellFormat(
      { ...ROWS_1_3_FORMAT, color: '#000000', backgroundColor: B_TO_M_HEADER_BG, textDecoration: 'underline' },
      `${pfx}B1`
    )
    setSubstructureColumnWidthsAndRowHeights(spreadsheet, sheetIndex, 1)
    return 1
  }

  // Row 1: headers
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
  spreadsheet.cellFormat(
    { ...ROWS_1_3_FORMAT, color: '#000000', backgroundColor: 'white' },
    `${pfx}A2`
  )
  spreadsheet.cellFormat(ROWS_1_3_FORMAT, `${pfx}B2:${colLetter(NUM_COLS - 1)}2`)

  // Row 3: A3 = "Substructure", B3–M3 green
  spreadsheet.updateCell({ value: SUBSTRUCTURE_LABEL }, `${pfx}A3`)
  spreadsheet.cellFormat(
    { ...BASE_FORMAT, fontWeight: 'bold', color: '#000000', backgroundColor: B_TO_M_HEADER_BG },
    `${pfx}A3`
  )
  spreadsheet.cellFormat(
    { ...BASE_FORMAT, backgroundColor: B_TO_M_HEADER_BG },
    `${pfx}B3:${colLetter(NUM_COLS - 1)}3`
  )

  for (let r = 0; r < substructureRows.length; r++) {
    const rowNum = r + 1
    for (let c = 0; c < NUM_COLS; c++) {
      try {
        spreadsheet.cellFormat(BASE_FORMAT, `${pfx}${colLetter(c)}${rowNum}`)
      } catch (e) { /* ignore */ }
    }
  }

  // Write data; I–M use substructureColLetter (FT, SQ FT, CY, QTY, LBS). Row 1: only A–H.
  for (let r = 0; r < substructureRows.length; r++) {
    const row = substructureRows[r]
    const rowNum = r + 1
    const maxC = (r === 0) ? 8 : Math.min(row.length, NUM_COLS)
    for (let c = 0; c < maxC; c++) {
      const val = row[c]
      const cellRef = `${pfx}${substructureColLetter(c)}${rowNum}`
      if (val !== '' && val !== null && val !== undefined) {
        try {
          const displayVal = c >= 2 ? toTwoDecimals(val) : val
          spreadsheet.updateCell({ value: displayVal }, cellRef)
        } catch (e) { /* ignore */ }
      }
    }
  }

  // Overwrite row 3 first data row (index 2) with "Substructure" in A; section header styling for Foundation row
  if (substructureRows.length >= 3) {
    const sectionRow = substructureRows[2]
    const sectionName = (sectionRow[0] != null && sectionRow[0] !== '') ? sectionRow[0] : SUBSTRUCTURE_LABEL
    spreadsheet.updateCell({ value: sectionName }, `${pfx}A3`)
  }

  // Style rows 3 onward: section header, subsection header, sum row, data row
  for (let r = 2; r < substructureRows.length; r++) {
    const row = substructureRows[r]
    const rowNum = r + 1
    const sectionName = (row[0] != null && row[0] !== '') ? String(row[0]).trim() : ''
    const isSectionHeader = sectionName === FOUNDATION_SECTION_NAME || (r === 2 && sectionName !== '')
    const isSubsectionHeader = !isSectionHeader && row[1] != null && String(row[1]).trim().endsWith(':')
    const isSumRow = sumRowNumbers.has(rowNum)

    if (isSectionHeader) {
      spreadsheet.updateCell({ value: sectionName || SUBSTRUCTURE_LABEL }, `${pfx}A${rowNum}`)
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
      const hasLbsValue = row[10] !== '' && row[10] != null && row[10] !== undefined
      if (hasLbsValue) {
        spreadsheet.cellFormat(
          { ...BASE_FORMAT, format: '#,##0.00', textAlign: 'right', backgroundColor: '#d9e1f2' },
          `${pfx}M${rowNum}`
        )
      }
    }
  }

  applySubstructureFormulas(spreadsheet, pfx, formulaData, foundationStartIdx, calculationData)
  setSubstructureColumnWidthsAndRowHeights(spreadsheet, sheetIndex, substructureRows.length)

  if (substructureRows.length > 0) {
    try {
      spreadsheet.cellFormat({ borderLeft: '3px solid #FF0000' }, `${pfx}I1:I${substructureRows.length}`)
    } catch (e) { /* ignore */ }
  }

  for (let c = 0; c < NUM_COLS; c++) {
    try {
      spreadsheet.updateCell({ value: COLUMN_HEADERS[c] }, `${pfx}${colLetter(c)}1`)
    } catch (e) { /* ignore */ }
  }

  return 1
}

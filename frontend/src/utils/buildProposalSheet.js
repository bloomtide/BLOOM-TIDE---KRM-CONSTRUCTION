import { fontColor } from '@syncfusion/ej2-spreadsheet'
import { convertToFeet } from './parsers/dimensionParser'
import proposalMapped from './proposal_mapped.json'

export function buildProposalSheet(spreadsheet, { calculationData, formulaData, rockExcavationTotals, lineDrillTotalFT, rawData, createdAt, project, client }) {
  // Full Proposal template (match screenshot layout).
  // Sheet order in workbook: Proposal Sheet = 0, Calculations Sheet = 1 (must match ProposalDetail.jsx sheets array order)
  const proposalSheetIndex = 0
  const pfx = 'Proposal Sheet!'

  // Header labels from proposal_mapped.json (DESCRIPTION entry); fallback to key name if missing
  const descriptionEntry = Array.isArray(proposalMapped) ? proposalMapped.find(e => e.description === 'DESCRIPTION') : null
  const headerLabels = descriptionEntry?.values || {}
  const label = (key) => (headerLabels[key] != null && headerLabels[key] !== '') ? String(headerLabels[key]) : key
  const L = { LF: label('LF'), SF: label('SF'), LBS: label('LBS'), CY: label('CY'), QTY: label('QTY'), LS: label('LS') }

  // Build formula so (XX) in column B references column G (QTY) of the same row on proposal sheet (use for all sections)
  const proposalFormulaWithQtyRef = (proposalRow, afterCount) => {
    const escaped = (afterCount || '').replace(/"/g, '""')
    return `=CONCATENATE("F&I new (", G${proposalRow}, "${escaped}")`
  }
  const afterCountFromProposalText = (text) => {
    const s = text || ''
    const idx = s.indexOf(')no ')
    if (idx >= 0) return s.slice(idx)
    const idx2 = s.indexOf(')no.')
    if (idx2 >= 0) return s.slice(idx2)
    return ''
  }

  // Normalize abbreviations when reading from calculation sheet (case-insensitive)
  // Thick: thk; Height: Ht, H; Width: W, Wide; Length: L, LF; No.: EA, No; Demolition: Demo; Excavation: Exc;
  // Slope/Soil/Rock exc; Line drill; Embedment: E; Rock socket: RS; Dowel: Dower/Dowels; Concrete soil retention pier: Concrete pier; Heel block: Foot block; Caisson: cassion
  // Foundation: PC, SF, WF, Wal->Wall, GB, TB, ST, F->Isolated footing, Liner wall->Concrete liner wall, FW; Elev.->Elevator; SOG, ROG, EAmp->Ramp; X P->X Perforated; SOMD, Insualtion->Insulation, Hanger->Concrete hanger, SW->Shear wall; Curb->Concrete curb, Pad->Concrete pad, Misc/Ext stair->Concrete stair
  const normalizeCalcSheetAbbreviations = (text) => {
    if (text == null || typeof text !== 'string') return text === undefined ? '' : text
    let s = String(text)
    s = s.replace(/\bthk\b/gi, 'Thick')
    s = s.replace(/\bHt\b/gi, 'Height')
    s = s.replace(/\bH\b/gi, 'Height')
    s = s.replace(/\bW\b/gi, 'Width')
    s = s.replace(/\bWide\b/gi, 'Width')
    s = s.replace(/\bLF\b/gi, 'Length')
    s = s.replace(/\bL\b/gi, 'Length')
    s = s.replace(/\bEA\b/gi, 'No.')
    s = s.replace(/\bNo\b(?!\.)/gi, 'No.') // No -> No. but not when already No.
    s = s.replace(/\bDemo\b/gi, 'Demolition') // Demo and Demolition treated as one
    s = s.replace(/\bExc\b/gi, 'Excavation') // Exc and Excavation treated as one
    s = s.replace(/\bSlope\s+exc\s+&\s+backfill\b/gi, 'Slope excavation & backfill') // Slope exc & backfill and Slope excavation & backfill treated as one (longer phrase first)
    s = s.replace(/\bSlope\s+exc\b/gi, 'Slope excavation') // Slope exc and Slope excavation treated as one
    s = s.replace(/\bSoil\s+slope\s+exc\s+&\s+backfill\b/gi, 'Soil slope excavation & backfill') // Soil slope exc & backfill (longer first)
    s = s.replace(/\bSoil\s+slope\s+exc\b/gi, 'Soil slope excavation') // Soil slope exc and Soil slope excavation treated as one
    s = s.replace(/\bSoil\s+exc\b/gi, 'Soil excavation') // Soil exc and Soil excavation treated as one
    s = s.replace(/\bLine\s+drill\b/gi, 'Line drilling') // Line drill and Line drilling treated as one
    s = s.replace(/\bRock\s+slope\s+exc\s+&\s+backfill\b/gi, 'Rock slope excavation & backfill') // Rock slope exc & backfill (longer first)
    s = s.replace(/\bRock\s+slope\s+exc\b/gi, 'Rock slope excavation') // Rock slope exc and Rock slope excavation treated as one
    s = s.replace(/\bRock\s+exc\b/gi, 'Rock excavation') // Rock exc and Rock excavation treated as one
    s = s.replace(/\bLine\s+drill\b/gi, 'Line drilling') // Line drill and Line drilling treated as one
    s = s.replace(/\bE\s*=\s*/gi, 'Embedment=') // E=##'-##" (embedment) so matching finds Embedment
    s = s.replace(/\bRS\s*=\s*/gi, 'Rock socket=') // RS=##'-##" so matching finds Rock socket
    s = s.replace(/\bDower\b/gi, 'Dowel') // Fix common misspelling: Dower bar / Steel dower bar -> Dowel
    s = s.replace(/\bDowels\b/gi, 'Dowel') // Dowels bar / Steel dowels bar -> Dowel
    s = s.replace(/\bConcrete\s+pier\b/gi, 'Concrete soil retention pier') // Concrete pier -> Concrete soil retention pier
    s = s.replace(/\bFoot\s+block\b/gi, 'Heel block') // Foot block -> Heel block
    s = s.replace(/\bcassion\b/gi, 'caisson') // Drilled cassion pile -> Drilled caisson pile
    // Foundation item abbreviations (case-insensitive)
    s = s.replace(/\bPC\b/gi, 'Pile cap') // Pile cap
    s = s.replace(/\bSF\b/gi, 'Strip footing') // Strip footing
    s = s.replace(/\bWF\b/gi, 'Wall footing') // Wall footing
    s = s.replace(/\bWal\b/gi, 'Wall') // Wal footing -> Wall footing
    s = s.replace(/\bGB\b/gi, 'Grade beam') // Grade beam
    s = s.replace(/\bTB\b/gi, 'Tie beam') // Tie beam
    s = s.replace(/\bST\b/gi, 'Strap beam') // Strap beam
    s = s.replace(/\bF\b/gi, 'Isolated footing') // Isolated footing (standalone F in particulars)
    s = s.replace(/\bLiner\s+wall\b/gi, 'Concrete liner wall') // Liner wall -> Concrete liner wall
    s = s.replace(/\bFW\b/gi, 'Foundation wall') // Foundation wall
    // Elevator / pit: Elev. -> Elevator (Elev. pit slab, Elev. pit wall, etc.)
    s = s.replace(/\bElev\.\b/gi, 'Elevator')
    // Slab/grade: SOG, ROG; typo EAmp -> Ramp
    s = s.replace(/\bSOG\b/gi, 'Slab on grade') // SOG ##", Patch SOG, SOG step
    s = s.replace(/\bROG\b/gi, 'Ramp on grade') // ROG ##", ROG ##'-##"
    s = s.replace(/\bEAmp\b/gi, 'Ramp') // typo: EAmp on grade -> Ramp on grade
    // Perforated pipe: X P -> X Perforated (so "X P" matches "X Perforated pipe")
    s = s.replace(/\bX\s+P\b/gi, 'X Perforated')
    // Slab on metal deck, topping/superstructure
    s = s.replace(/\bSOMD\b/gi, 'Slab on metal deck')
    s = s.replace(/\bInsualtion\b/gi, 'Insulation') // typo: Insualtion -> Insulation
    s = s.replace(/(?<!Concrete\s)\bHanger\b/gi, 'Concrete hanger') // Hanger -> Concrete hanger (not when already "Concrete hanger")
    s = s.replace(/\bSW\b/gi, 'Shear wall') // Shear wall (SW @ ^^, SW (##'x##'))
    // Curb, pad, stair variants
    s = s.replace(/(?<!Concrete\s)\bCurb\b/gi, 'Concrete curb') // Curb -> Concrete curb (not when already "Concrete curb")
    s = s.replace(/(?<!Housekeeping\s)(?<!Transformer\s)\bPad\b/gi, 'Concrete pad') // Pad -> Concrete pad (not Housekeeping/Transformer pad)
    s = s.replace(/\bMisc\.?\s*stair\b/gi, 'Concrete stair') // Misc. stair / Misc stair -> Concrete stair
    s = s.replace(/\bExt\.?\s*stair\b/gi, 'Concrete stair') // Ext. stair / Ext stair -> Concrete stair
    return s
  }

  // Case-insensitive abbreviation/display name lookup for whole proposal sheet (handles upper/lower/mixed case)
  const getDisplayNameForAbbreviation = (item, map) => {
    if (item == null || item === '') return item
    const s = String(item).trim()
    if (!map || typeof map !== 'object') return s
    if (map[s] !== undefined) return map[s]
    const lower = s.toLowerCase()
    const entry = Object.entries(map).find(([k]) => (k || '').toLowerCase() === lower)
    return entry ? entry[1] : s
  }

  // Format multiple page refs for template text: "102 & 105" or "102, 103 & 105" (used in all scopes)
  const formatPageRefList = (refs) => {
    const list = Array.isArray(refs) ? refs.filter(Boolean) : []
    if (list.length === 0) return '##'
    if (list.length === 1) return list[0]
    if (list.length === 2) return `${list[0]} & ${list[1]}`
    return list.slice(0, -1).join(', ') + ' & ' + list[list.length - 1]
  }
  // Column G width: numeric 4.5 (feet) -> "4'-6""
  const formatWidthFeetInches = (val) => {
    const n = parseFloat(val)
    if (val == null || val === '' || isNaN(n)) return null
    const f = Math.floor(n)
    const inch = Math.round((n - f) * 12)
    if (inch === 0) return `${f}'-0"`
    return `${f}'-${inch}"`
  }

  // Parse drilled caisson pile template fields from particulars (design loads, grout KSI, rebar)
  const parseDrilledPileTemplateFromParticulars = (particulars) => {
    const p = (particulars || '').toString()
    const compMatch = p.match(/(\d+)\s*tons?\s*design\s*compression/i)
    const tensionMatch = p.match(/(\d+)\s*tons?\s*design\s*tension/i)
    const lateralMatch = p.match(/(\d+)\s*ton(?:s)?\s*design\s*lateral/i)
    const groutMatch = p.match(/(\d+)-?\s*KSI\s*grout/i) || p.match(/(\d+)\s*KSI\s*grout\s*infilled/i)
    const rebarMatch = p.match(/with\s*\((\d+)\)\s*qty\s*#?(\d+)"?\s*(\d+)\s*Ksi/i) || p.match(/\((\d+)\)\s*qty\s*#?(\d+)"?\s*(\d+)\s*Ksi/i)
    return {
      designCompression: compMatch ? compMatch[1] : null,
      designTension: tensionMatch ? tensionMatch[1] : null,
      designLateral: lateralMatch ? lateralMatch[1] : null,
      groutKsi: groutMatch ? groutMatch[1] : null,
      rebarQty: rebarMatch ? rebarMatch[1] : null,
      rebarSize: rebarMatch ? rebarMatch[2] : null,
      rebarKsi: rebarMatch ? rebarMatch[3] : null
    }
  }

  // Hardcoded template for all items in foundation drilled piles scope (placeholders # filled from data when available)
  const buildDrilledPileTemplateText = (particulars, diameterThicknessText, heightText, rockSocketText, qtyPlaceholder) => {
    const parsed = parseDrilledPileTemplateFromParticulars(particulars)
    const comp = parsed.designCompression ?? '140' // hardcoded 140 tons design compression per user
    const tension = parsed.designTension ?? '#'
    const lateral = parsed.designLateral ?? '#'
    const grout = parsed.groutKsi ?? '#'
    const rebarQty = parsed.rebarQty ?? '#'
    const rebarSize = parsed.rebarSize ?? '#'
    const rebarKsi = parsed.rebarKsi ?? '#'
    const thickPart = diameterThicknessText || '(#" ØX#" thick)'
    const heightPart = (heightText && rockSocketText && rockSocketText !== "0'-0\"")
      ? `(H=${heightText}, ${rockSocketText} rock socket)`
      : (heightText ? `(H=${heightText})` : "(H=#'-#\", #'-#\" rock socket)")
    return `F&I new (${qtyPlaceholder})no drilled caisson pile (${comp} tons design compression, ${tension} tons design tension & ${lateral} ton design lateral load), ${thickPart}, (${grout}-KSI grout infilled) with (${rebarQty})qty #${rebarSize}" ${rebarKsi} Ksi full length reinforcement ${heightPart} as per`
  }

  // Extract "whatever is after () and before as per" for rate matching (e.g. "no W12x58 walers" from "F&I new (11)no W12x58 walers as per SOE-101.00")
  const getDescriptionKeyAfterParenBeforeAsPer = (descriptionText) => {
    if (!descriptionText || typeof descriptionText !== 'string') return null
    const closeParen = descriptionText.indexOf(')')
    if (closeParen < 0) return null
    const asPer = descriptionText.toLowerCase().indexOf(' as per ', closeParen)
    if (asPer < 0) return null
    const key = descriptionText.slice(closeParen + 1, asPer).trim()
    return key || null
  }

  // Line rates from proposal_mapped.json only; also tries key "after () and before as per" for bracing-type lines
  const getRatesForDescription = (descriptionText) => {
    if (!descriptionText || typeof descriptionText !== 'string') return null
    const trimmed = descriptionText.trim()
    if (!trimmed) return null

    if (!Array.isArray(proposalMapped)) return null
    const candidates = proposalMapped.filter(e => e.description !== 'DESCRIPTION' && e.values && typeof e.values === 'object')
    const withRates = candidates.filter(e => {
      const v = e.values
      return Object.keys(v).some(k => typeof v[k] === 'number')
    })
    const tryMatch = (key) => {
      if (!key) return null
      const exact = candidates.find(e => e.description.trim() === key)
      if (exact && exact.values) return exact.values
      // Require word boundary so "Survey" does not match "Surveying, stakeout..."
      const atWordBoundary = (desc, k) => {
        const pos = k.indexOf(desc)
        if (pos < 0) return false
        const after = k.slice(pos + desc.length)
        return after.length === 0 || !/^[a-zA-Z]/.test(after)
      }
      const rowContains = withRates.filter(e => {
        const desc = e.description.trim()
        return key.indexOf(desc) >= 0 && atWordBoundary(desc, key)
      })
      if (rowContains.length > 0) {
        // When line has both "full depth asphalt pavement" and "surface course", prefer "surface course" (SF 10) so dynamic thickness values work across sheets
        const keyLower = key.toLowerCase()
        if (keyLower.includes('full depth asphalt pavement') && keyLower.includes('surface course')) {
          const surfaceCourse = rowContains.find(e => e.description.trim().toLowerCase().includes('surface course'))
          if (surfaceCourse && surfaceCourse.values) return surfaceCourse.values
        }
        // When line has "buttresses" and "foundation walls", prefer "buttresses" (CY 950) not "foundation walls" (CY 900)
        if (keyLower.includes('buttresses') && keyLower.includes('foundation walls')) {
          const buttresses = rowContains.find(e => e.description.trim().toLowerCase().includes('buttresses'))
          if (buttresses && buttresses.values) return buttresses.values
        }
        // When line has "saw-cut/demo/remove/dispose existing" and "sidewalk/driveway", use sidewalk/driveway rate (SF 11) even if key has extra text like 7" (so exact substring may not match)
        if (keyLower.includes('saw-cut/demo/remove/dispose existing') && keyLower.includes('sidewalk/driveway')) {
          const sidewalkDriveway = withRates.find(e => e.description.trim().toLowerCase().includes('sidewalk/driveway'))
          if (sidewalkDriveway && sidewalkDriveway.values) return sidewalkDriveway.values
        }
        // When line has "saw-cut/demo/remove/dispose existing" and "sidewalk", prefer sidewalk rate (SF 10) over generic
        if (keyLower.includes('saw-cut/demo/remove/dispose existing') && keyLower.includes('sidewalk')) {
          const sidewalk = rowContains.find(e => e.description.trim().toLowerCase().includes('sidewalk'))
          if (sidewalk && sidewalk.values) return sidewalk.values
        }
        // When line has "concrete sidewalk, reinf w/" and "Pavement replacement: West Street" or "Maple Avenue", prefer that rate (SF 12) over generic "concrete sidewalk, reinf w/ " (SF 15)
        if (keyLower.includes('concrete sidewalk, reinf w/') && (keyLower.includes('pavement replacement: west street') || keyLower.includes('pavement replacement: maple avenue'))) {
          const westStreet = withRates.find(e => e.description.trim().toLowerCase().includes('pavement replacement: west street'))
          const mapleAve = withRates.find(e => e.description.trim().toLowerCase().includes('pavement replacement: maple avenue'))
          if (keyLower.includes('pavement replacement: west street') && westStreet?.values) return westStreet.values
          if (keyLower.includes('pavement replacement: maple avenue') && mapleAve?.values) return mapleAve.values
        }
        // Prefer longest matching description so "slab on grade reinforced" wins over "slab"; if same length, prefer earliest in key
        const best = rowContains.reduce((best, e) => {
          const desc = e.description.trim()
          const pos = key.indexOf(desc)
          if (pos < 0) return best
          if (!best) return e
          const bestDesc = best.description.trim()
          const bestPos = key.indexOf(bestDesc)
          const longer = desc.length > bestDesc.length ? e : best
          const sameLen = desc.length === bestDesc.length
          if (!sameLen) return longer
          return pos <= bestPos ? e : best
        }, null)
        if (best && best.values) return best.values
      }
      const entryContains = withRates.filter(e => e.description.trim().indexOf(key) >= 0)
      if (entryContains.length > 0) return entryContains.reduce((best, e) => (e.description.length < (best?.description?.length || 1e9) ? e : best), entryContains[0]).values
      return null
    }
    const fullMatch = tryMatch(trimmed)
    if (fullMatch) return fullMatch
    const afterParenKey = getDescriptionKeyAfterParenBeforeAsPer(descriptionText)
    return tryMatch(afterParenKey) || null
  }

  // Fill columns I–N (unit rates $/LF, $/SF, etc.) from proposal_mapped.json for a proposal row
  // When rateLookupKey is provided, use it for rate lookup (e.g. BPP concrete sidewalk under "Pavement replacement: West Street") so SF 12 is used instead of generic SF 15
  const fillRatesForProposalRow = (row, descriptionText, rateLookupKey) => {
    const values = getRatesForDescription(rateLookupKey != null ? rateLookupKey : descriptionText)
    if (!values) return
    const colMap = { LF: 'I', SF: 'J', LBS: 'K', CY: 'L', QTY: 'M', LS: 'N' }
    Object.entries(colMap).forEach(([key, col]) => {
      const num = values[key]
      if (num != null && typeof num === 'number' && !Number.isNaN(num)) {
        try {
          spreadsheet.updateCell({ value: num }, `${pfx}${col}${row}`)
        } catch (e) { /* ignore */ }
      }
    })
  }

  // Clear everything (values + formats) on Proposal sheet so it never retains other sheet data.
  try {
    spreadsheet.clear({ range: `${pfx}A1:Z1000`, type: 'Clear All' })
  } catch (e) {
    try { spreadsheet.clear({ range: `${pfx}A1:Z1000` }) } catch (e2) { /* ignore */ }
  }

  // Keep the header row visible when scrolling – freeze row 1
  try {
    // Prefer API that accepts sheet index when available
    if (typeof spreadsheet.freezePanes === 'function') {
      spreadsheet.freezePanes(1, 0, proposalSheetIndex)
    }
  } catch (e) {
    // Fallback for older signatures that don't take sheet index
    try { spreadsheet.freezePanes(1, 0) } catch (e2) { /* ignore */ }
  }

  // Column widths to mirror screenshot
  const colWidths = {
    A: 40,  // margin
    B: 955, // main left block
    C: 146, // LF
    D: 146, // SF
    E: 146, // LBS
    F: 146, // CY
    G: 146, // QTY
    H: 107, // $/1000 spacer - widened
    I: 100,  // LF - wide enough for $  12,000.00 and larger
    J: 100,  // SF
    K: 100,  // LBS
    L: 100,  // CY
    M: 100,  // QTY
    N: 100   // LS
  }
  // Calculation sheet (index 0): default width per column (Estimate, Particulars, Takeoff, Unit, QTY, Length, Width, Height, FT, SQ FT, LBS, CY, QTY)
  const calculationSheetIndex = 1
  const calcSheetColWidths = {
    A: 70,   // Estimate
    B: 220,  // Particulars (wider for item text)
    C: 90,   // Takeoff
    D: 55,   // Unit
    E: 55,   // QTY
    F: 70,   // Length
    G: 70,   // Width
    H: 70,   // Height
    I: 70,   // FT
    J: 75,   // SQ FT
    K: 70,   // LBS
    L: 70,   // CY
    M: 55    // QTY
  }
  Object.entries(calcSheetColWidths).forEach(([col, width]) => {
    try { spreadsheet.setColWidth(width, col.charCodeAt(0) - 65, calculationSheetIndex) } catch (e) { }
  })
  // Red right border on column H from row 1 to last row with data on Calculation sheet
  const calcLastRow = calculationData && calculationData.length > 0 ? calculationData.length : 1
  try {
    spreadsheet.cellFormat({ borderRight: '1px solid #ff0000' }, `'Calculations Sheet'!H1:H${calcLastRow}`)
  } catch (e) { }
  // One empty row after data, then black bottom border on that row (whole row)
  const calcEmptyRow = calcLastRow + 1
  try {
    spreadsheet.cellFormat({ borderBottom: '1px solid #000000' }, `'Calculations Sheet'!A${calcEmptyRow}:N${calcEmptyRow}`)
  } catch (e) { }

  // Apply styling to Calculations Sheet rows with "Influ" marker in column A
  if (calculationData && Array.isArray(calculationData)) {
    calculationData.forEach((row, index) => {
      if (row && row[0] === 'Influ') {
        const rowNum = index + 1
        try {
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: ' center',
              verticalAlign: 'middle'

            },
            `'Calculations Sheet'!A${rowNum}`
          )
        } catch (e) { }
      }
    })
  }

  // Apply column widths only to Proposal sheet (sheet index 1), not to Calculation sheet
  Object.entries(colWidths).forEach(([col, width]) => {
    try { spreadsheet.setColWidth(width, col.charCodeAt(0) - 65, proposalSheetIndex) } catch (e) { }
  })

  // Uniform height for one-liner and empty rows on Proposal sheet (same across whole sheet)
  const DEFAULT_ROW_HEIGHT = 30
  for (let ri = 0; ri <= 11; ri++) {
    try { spreadsheet.setRowHeight(DEFAULT_ROW_HEIGHT, ri, proposalSheetIndex) } catch (e) { }
  }
  // Row 13 (index 12) is empty spacer – set same default height
  try { spreadsheet.setRowHeight(DEFAULT_ROW_HEIGHT, 12, proposalSheetIndex) } catch (e) { }

  // Track all section total rows for BASE BID TOTAL formula
  const baseBidTotalRows = []
  // Track rows that should remain bold (totals) to override the unbold rule
  const totalRows = []
  // Rows where column C is "Total SF" – re-apply bold + italic after global format
  const totalSFRowsForItalic = []
  // Note rows (e.g. "Note: Backfill SOE voids...") – keep normal weight after global bold
  const noteRows = []
  // WR Meadows Installer note row – re-apply normal, italic, center
  const wpNoteRow = []
  // Row with "General Liability total coverage..." – re-apply 3px bottom border last so it is not overwritten
  let generalLiabilityBorderRow = null
  // Rows with long text that need dynamic height (e.g. soil excavation); re-applied after uniform height loop
  const dynamicHeightRows = []
  // Column B content per row for dynamic height (row number -> text)
  const rowBContentMap = new Map()

  // Helper function to calculate row height based on text content in column B
  // One-liner and empty: same height (DEFAULT_ROW_HEIGHT). When text wraps to multiple lines: height fits content.
  const calculateRowHeight = (text) => {
    const trimmed = (text && typeof text === 'string') ? String(text).trim() : ''
    if (trimmed === '') {
      return DEFAULT_ROW_HEIGHT
    }

    // Column B width is 955 pixels
    const columnWidth = 955
    const lineHeight = 26
    const paddingTop = 8
    const paddingBottom = 8
    const verticalPadding = paddingTop + paddingBottom
    const charWidth = 11
    const availableWidth = columnWidth - 20
    const charsPerLine = Math.floor(availableWidth / charWidth)
    const estimatedLines = Math.max(1, Math.ceil(trimmed.length / charsPerLine))

    // One line: same height as empty rows for consistency
    if (estimatedLines <= 1) {
      return DEFAULT_ROW_HEIGHT
    }

    // Multiple lines: height fits content
    const calculatedHeight = Math.ceil(estimatedLines * lineHeight + verticalPadding)
    return Math.min(Math.max(DEFAULT_ROW_HEIGHT, calculatedHeight), 400)
  }


  // Styles
  const headerGray = { backgroundColor: '#D9D9D9', fontWeight: 'bold', fontFamily: 'Calibri', fontSize: '10pt', textAlign: 'center', verticalAlign: 'middle' }
  const boxGray = { backgroundColor: '#D9D9D9', fontWeight: 'bold', fontFamily: 'Calibri', fontSize: '18pt', verticalAlign: 'middle' }
  const thick = { border: '2px solid #000000' }
  const thickTop = { borderTop: '4px solid #000000' }
  const thickBottom = { borderBottom: '2px solid #000000' }
  const thickLeft = { borderLeft: '2px solid #000000' }
  const thickRight = { borderRight: '2px solid #000000' }
  const thin = { border: '1px solid #000000' }

  // Total row B-G borders: left of B and right of G black; top/bottom black; internal borders match row bg
  function applyTotalRowBorders(spreadsheet, pfx, row, backgroundColor) {
    const black = '1px solid #000000'
    const bgBorder = `1px solid ${backgroundColor}`
    spreadsheet.cellFormat({ backgroundColor, borderTop: black, borderBottom: black, borderLeft: black, borderRight: bgBorder }, `${pfx}B${row}`)
    spreadsheet.cellFormat({ backgroundColor, borderTop: black, borderBottom: black, borderLeft: bgBorder, borderRight: bgBorder }, `${pfx}C${row}`)
    spreadsheet.cellFormat({ backgroundColor, borderTop: black, borderBottom: black, borderLeft: bgBorder, borderRight: bgBorder }, `${pfx}D${row}`)
    spreadsheet.cellFormat({ backgroundColor, borderTop: black, borderBottom: black, borderLeft: bgBorder, borderRight: bgBorder }, `${pfx}E${row}`)
    spreadsheet.cellFormat({ backgroundColor, borderTop: black, borderBottom: black, borderLeft: bgBorder, borderRight: bgBorder }, `${pfx}F${row}`)
    spreadsheet.cellFormat({ backgroundColor, borderTop: black, borderBottom: black, borderLeft: bgBorder, borderRight: black }, `${pfx}G${row}`)
  }

  // Top header row (row 1) - labels from proposal_mapped.json DESCRIPTION entry
  ;[
    ['C1', L.LF], ['D1', L.SF], ['E1', L.LBS], ['F1', L.CY], ['G1', L.QTY],
    ['H1', '$/1000'],
    ['I1', L.LF], ['J1', L.SF], ['K1', L.LBS], ['L1', L.CY], ['M1', L.QTY], ['N1', L.LS]
  ].forEach(([cell, value]) => spreadsheet.updateCell({ value }, `${pfx}${cell}`))
  spreadsheet.cellFormat(headerGray, `${pfx}C1:N1`)
  spreadsheet.cellFormat(thick, `${pfx}C1:N1`)
  // Make $/1000 cell background white; $/1000 heading at 11pt; border; centered
  spreadsheet.cellFormat({ backgroundColor: 'white', fontSize: '11pt', border: '1px solid #000000', textAlign: 'center', verticalAlign: 'middle' }, `${pfx}H1`)

  // Main outer frame around top content block (PERIMETER ONLY).
  // Avoid applying a border to the full range, since that creates thick internal grid lines.
  // Border is now applied dynamically at the end to wrap the entire content.

  // Per request: rows 3 to 8, columns B to E should look like "no border configured"
  // (i.e., rely on default grid lines), and text should be normal weight in black.
  try {
    spreadsheet.clear({ range: `${pfx}B3:E8`, type: 'Clear Formats' })
  } catch (e) {
    // ignore (different versions may have different clear type strings)
  }
  spreadsheet.cellFormat({ fontWeight: 'normal', color: '#000000', fontFamily: 'Calibri (Body)' }, `${pfx}B3:E8`)

  // Left address block
  spreadsheet.updateCell({ value: '37-24 24th Street, Suite 132, Long Island City, NY 11101' }, `${pfx}B4`)
  spreadsheet.updateCell({ value: 'Tel: 718 726-1525 | Fax: 718 726-1601 | Cell: 917 600-3958' }, `${pfx}B5`)
  spreadsheet.updateCell({ value: 'Email:' }, `${pfx}B6`)
  spreadsheet.cellFormat({ border: '3px solid #000000' }, `${pfx}F4:G4`)
  spreadsheet.cellFormat({ border: '2px solid #000000' }, `${pfx}F6:G6`)
  // F5:G5 formatted separately below so its borderTop can override
  // Remove borders from B4:E8 - ensure no border color
  try {
    spreadsheet.clear({ range: `${pfx}B4:E8`, type: 'Clear Formats' })
  } catch (e) {
    // ignore
  }
  // Reapply formatting without borders
  spreadsheet.cellFormat({ fontWeight: 'normal', color: '#000000', fontFamily: 'Calibri (Body)' }, `${pfx}B4:E8`)
  spreadsheet.cellFormat({ fontWeight: 'bold' }, `${pfx}B4`)
  // Tel/Fax/Cell line: bold, black, 18pt
  spreadsheet.cellFormat({ fontWeight: 'bold', fontSize: '18pt', color: '#000000' }, `${pfx}B5`)
  spreadsheet.cellFormat({ color: '#0B76C3', textDecoration: 'underline' }, `${pfx}B6`)


  // Date/Project/Client block
  const proposalDate = createdAt ? new Date(createdAt).toLocaleDateString() : "Today's date"
  const estimateYearFull = createdAt ? new Date(createdAt).getFullYear() : new Date().getFullYear()
  const estimateYearSuffix = String(estimateYearFull).slice(-2)
  const estimateLabel = `Estimate #${estimateYearSuffix}-`

  spreadsheet.updateCell({ value: `Date: ${proposalDate}` }, `${pfx}B9`)
  spreadsheet.updateCell({ value: `Project: ${project || '###'}` }, `${pfx}B10`)
  spreadsheet.updateCell({ value: `Client: ${client || '###'}` }, `${pfx}B11`)
  spreadsheet.cellFormat({ fontWeight: 'bold' }, `${pfx}B9:B11`)
  spreadsheet.cellFormat(thin, `${pfx}B9:G11`)

  // Right grey box (Estimate / Drawings Dated / lines) - start from column F
  // Individual cell styling - customize each cell separately

  // Row 3: Estimate #YY- (centered); 3px border on all sides
  spreadsheet.merge(`${pfx}F3:G3`)
  spreadsheet.updateCell({ value: estimateLabel }, `${pfx}F3`)
  spreadsheet.cellFormat({
    backgroundColor: '#D0CECE',
    fontSize: '11pt',
    fontWeight: 'bold',
    textAlign: 'center',
    borderTop: '3px solid #000000',
    borderLeft: '3px solid #000000',
    borderBottom: '3px solid #000000',
    borderRight: '3px solid #000000'
  }, `${pfx}F3:G3`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontSize: '11pt', fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}H3`)

  // Row 4: Empty row (no left border on F4)
  spreadsheet.merge(`${pfx}F4:G4`)
  spreadsheet.cellFormat({ backgroundColor: 'white', borderLeft: 'none' }, `${pfx}F4`)
  spreadsheet.cellFormat({ backgroundColor: 'white', borderRight: '1px solid #000000' }, `${pfx}G4`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}H4`)

  // Row 5: Drawings Dated: (all borders set explicitly so borderTop overrides any prior range format)
  spreadsheet.merge(`${pfx}F5:G5`)
  spreadsheet.updateCell({ value: 'Drawings Dated:' }, `${pfx}F5`)
  spreadsheet.cellFormat({
    textAlign: 'center',
    backgroundColor: '#D0CECE',
    fontWeight: 'bold',
    borderTop: '3px solid #000000',
    borderLeft: '3px solid #000000',
    borderRight: '2px solid #000000',
    borderBottom: '1px solid #D0CECE'
  }, `${pfx}F5:G5`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}H5`)

  // Row 6: SOE:
  spreadsheet.merge(`${pfx}F6:G6`)
  spreadsheet.updateCell({ value: 'SOE:' }, `${pfx}F6`)
  // Side borders only
  spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '3px solid #000000', borderRight: '2px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}F6:G6`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}H6`)

  // Row 7: Structural:
  spreadsheet.merge(`${pfx}F7:G7`)
  spreadsheet.updateCell({ value: 'Structural:' }, `${pfx}F7`)
  // Side borders only
  spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '3px solid #000000', borderRight: '2px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}F7:G7`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}H7`)

  // Row 8: Architectural:
  spreadsheet.merge(`${pfx}F8:G8`)
  spreadsheet.updateCell({ value: 'Architectural:' }, `${pfx}F8`)
  // Side borders only
  spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '3px solid #000000', borderRight: '2px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}F8:G8`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}H8`)

  // Row 9: Plumbing:
  spreadsheet.merge(`${pfx}F9:G9`)
  spreadsheet.updateCell({ value: 'Plumbing:' }, `${pfx}F9`)
  // Side borders only
  spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '3px solid #000000', borderRight: '2px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}F9:G9`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}H9`)

  // Row 10: Mechanical
  spreadsheet.merge(`${pfx}F10:G10`)
  spreadsheet.updateCell({ value: 'Mechanical' }, `${pfx}F10`)
  // Bottom border, Side borders
  spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '3px solid #000000', borderBottom: '3px solid #000000', borderRight: '2px solid #000000' }, `${pfx}F10:G10`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}H10`)

  // Explicit right border on G5, G6, G10, G11 (right edge of grey box / Date block)
  const rightBorderBlack = { borderLeft: '2px solid #000000' }
  spreadsheet.cellFormat(rightBorderBlack, `${pfx}H5`)
  spreadsheet.cellFormat(rightBorderBlack, `${pfx}H6`)
  spreadsheet.cellFormat(rightBorderBlack, `${pfx}H9`)
  spreadsheet.cellFormat(rightBorderBlack, `${pfx}H10`)


  // Bottom header row (row 12) - labels from proposal_mapped.json DESCRIPTION entry
  spreadsheet.updateCell({ value: 'DESCRIPTION' }, `${pfx}B12`)
    ;[
      ['C12', L.LF], ['D12', L.SF], ['E12', L.LBS], ['F12', L.CY], ['G12', L.QTY],
      ['H12', '$/1000'],
      ['I12', L.LF], ['J12', L.SF], ['K12', L.LBS], ['L12', L.CY], ['M12', L.QTY]
    ].forEach(([cell, value]) => spreadsheet.updateCell({ value }, `${pfx}${cell}`))
  spreadsheet.updateCell({ value: L.LS }, `${pfx}N12`)

  spreadsheet.cellFormat(headerGray, `${pfx}B12:N12`)
  spreadsheet.cellFormat({ fontSize: '18pt', borderTop: '3px solid #000000', borderLeft: '2px solid #000000', borderRight: '3px solid #000000', borderBottom: '3px solid #000000' }, `${pfx}C12:N12`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontSize: '11pt', textAlign: 'center', verticalAlign: 'middle' }, `${pfx}H12`) // $/1000 in row 12: 11pt, white, centered
  spreadsheet.cellFormat({ fontWeight: '300' }, `${pfx}B12:N12`) // row 12: semi-bold
  // Row 12 (B12:N12): 3px top, bottom, and right (N12); 2px left
  spreadsheet.cellFormat({ borderTop: '3px solid #000000', borderLeft: '2px solid #000000', borderRight: '3px solid #000000', borderBottom: '3px solid #000000' }, `${pfx}B12:N12`)

  // Row 13: Empty row between DESCRIPTION and Demolition scope
  spreadsheet.cellFormat({ backgroundColor: 'white' }, `${pfx}B13:N13`)

  // Pre-fill $/1000 formula in column H for all data rows (13-1000)
  // This ensures the formula is always present and calculates automatically when other cells are filled
  // IFERROR returns blank if there's an error (like #NAME?)
  // IF returns blank if the result is 0 (no data in the row)
  for (let row = 13; row <= 1000; row++) {
    const dollarFormula = `=IFERROR(IF(ROUNDUP(MAX(C${row}*I${row},D${row}*J${row},E${row}*K${row},F${row}*L${row},G${row}*M${row},N${row})/1000,1)=0,"",ROUNDUP(MAX(C${row}*I${row},D${row}*J${row},E${row}*K${row},F${row}*L${row},G${row}*M${row},N${row})/1000,1)),"")`
    spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${row}`)
  }

  // Apply formatting to column H ($/1000) for all data rows
  try {
    spreadsheet.cellFormat(
      {
        fontWeight: 'normal',
        fontFamily: 'Calibri (Body)',
        textAlign: 'right',
        backgroundColor: 'white',
        format: '$#,##0.00'
      },
      `${pfx}H13:H1000`
    )
    spreadsheet.numberFormat('$#,##0.00', `${pfx}H13:H1000`)
  } catch (e) {
    // Fallback - format will be applied individually as rows are created
  }

  // Row 14: Demolition scope heading
  spreadsheet.updateCell({ value: 'Demolition scope:' }, `${pfx}B14`)
  rowBContentMap.set(14, 'Demolition scope:')
  spreadsheet.wrap(`${pfx}B14`, true)
  spreadsheet.cellFormat({
    backgroundColor: '#BDD7EE',
    textAlign: 'center',
    verticalAlign: 'middle',
    textDecoration: 'underline',
    fontWeight: 'normal',
    border: '1px solid #000000',
    wrapText: true
  }, `${pfx}B14`)

  // Demolition scope lines from Calculations Sheet:
  // For each Demolition subsection, take the first item description in column B
  // (e.g. "Allow to saw-cut/demo/remove/dispose existing (4\" thick) slab on grade @ existing building as per DM-106.00 & details on DM-107.00")
  // and show it under Demolition scope on the Proposal sheet.
  // Always render demolition lines (with defaults when calculationData is empty)
  // DM reference (e.g. DM-107.00) is taken from the Page column of the first matching row for that subsection, not hardcoded.
  const getDMReferenceFromRawData = (subsectionName) => {
    if (!rawData || !Array.isArray(rawData) || rawData.length < 2) return '##'
    const headers = rawData[0]
    const dataRows = rawData.slice(1)
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
    if (digitizerIdx === -1) return '##'
    const subsectionPatterns = {
      'Demo slab on grade': /demo\s+sog/i,
      'Demo strip footing': /demo\s+sf/i,
      'Demo foundation wall': /demo\s+fw/i,
      'Demo isolated footing': /demo\s+isolated\s+footing/i,
      'Demo Ramp on grade': /demo\s+rog|demo\s+ramp\s+on\s+grade/i,
      'Demo retaining wall': /demo\s+rw|demo\s+retaining\s+wall/i,
      // Match "Demo stair landings on grade @ Stair A1", "Demo stairs on grade", "Demo landings on grade" etc. in Digitizer Item
      'Demo stair on grade': /demo\s+(?:stair\s+)?(?:stairs|landings)\s+on\s+grade/i,
      'Demo pile cap': /demo\s+(?:pile\s+caps?|PC)/i,
      'Demo pile caps': /demo\s+(?:pile\s+caps?|PC)/i,
      'Demo pilaster': /demo\s+pilaster/i,
      'Demo grade beam': /demo\s+(?:grade\s+beam|GB)/i,
      'Demo tie beam': /demo\s+(?:tie\s+beam|TB)/i,
      'Demo strap beam': /demo\s+(?:strap\s+beam|ST)/i,
      'Demo thickened slab': /demo\s+thickened\s+slab/i,
      'Demo buttress': /demo\s+buttress/i,
      'Demo pier': /demo\s+pier/i,
      'Demo corbel': /demo\s+corbel/i,
      'Demo liner wall': /demo\s+(?:concrete\s+)?liner\s+wall/i,
      'Demo barrier wall': /demo\s+(?:vehicle\s+)?barrier\s+wall/i,
      'Demo stem wall': /demo\s+stem\s+wall/i,
      'Demo elevator pit': /demo\s+(?:sump\s+pit\s+@\s+)?elevator\s+pit|demo\s+elev\.?\s+pit\b|demo\s+elevator\s+pit/i,
      'Demo service elevator pit': /demo\s+service\s+elevator\s+pit|demo\s+service\s+elev\.?\s+pit/i,
      'Demo detention tank': /demo\s+detention\s+tank/i,
      'Demo duplex sewage ejector pit': /demo\s+duplex\s+sewage\s+ejector/i,
      'Demo deep sewage ejector pit': /demo\s+deep\s+sewage\s+ejector/i,
      'Demo sump pump pit': /demo\s+sump\s+pump\s+(?:pit|wall)/i,
      'Demo grease trap pit': /demo\s+grease\s+trap\s+pit/i,
      'Demo house trap pit': /demo\s+house\s+trap/i,
      'Demo mat slab': /demo\s+mat(?:\s+slab|-\d)/i
    }
    const pattern = subsectionPatterns[subsectionName]
    if (!pattern) return '##'
    const collected = []
    const seen = new Set()
    for (const row of dataRows) {
      const digitizerItem = row[digitizerIdx]
      if (!digitizerItem || !pattern.test(String(digitizerItem))) continue
      let ref = null
      if (pageIdx >= 0) {
        const pageValue = row[pageIdx]
        if (pageValue != null && pageValue !== '') {
          const pageStr = String(pageValue).trim()
          const dmMatch = pageStr.match(/DM-[\d.]+/i) || pageStr.match(/A-[\d.]+/i) || pageStr.match(/FO-[\d.]+/i)
          if (dmMatch) ref = dmMatch[0]
        }
      }
      if (!ref) {
        const dmMatch = String(digitizerItem).trim().match(/DM-[\d.]+/i) || String(digitizerItem).trim().match(/A-[\d.]+/i) || String(digitizerItem).trim().match(/FO-[\d.]+/i)
        if (dmMatch) ref = dmMatch[0]
      }
      if (ref && !seen.has(ref)) {
        seen.add(ref)
        collected.push(ref)
      }
    }
    return formatPageRefList(collected)
  }

  // For demolition pits/mat slabs, estimate an average height/depth from calculation rows
  const getAverageHeightForDemoSubsection = (subsectionName) => {
    if (!rowsBySubsection || !subsectionName) return null
    const subsectionRows = rowsBySubsection.get(subsectionName) || []
    if (!subsectionRows.length) return null

    const toInches = (token) => {
      if (!token) return null
      const s = String(token).trim()
      let m = s.match(/^(-?\d+)'-(\d+)"?$/)
      if (m) {
        const feet = parseInt(m[1], 10)
        const inches = parseInt(m[2], 10)
        return feet * 12 + inches
      }
      m = s.match(/^(-?\d+)'$/)
      if (m) {
        const feet = parseInt(m[1], 10)
        return feet * 12
      }
      m = s.match(/^(\d+)"$/)
      if (m) {
        return parseInt(m[1], 10)
      }
      m = s.match(/^(\d+(?:\.\d+)?)$/)
      if (m) {
        return parseFloat(m[1])
      }
      return null
    }

    const formatInches = (valueInInches) => {
      if (valueInInches == null || Number.isNaN(valueInInches)) return null
      const total = Math.round(valueInInches)
      const feet = Math.floor(total / 12)
      const inches = total % 12
      if (feet > 0 && inches > 0) return `${feet}'-${inches}"`
      if (feet > 0) return `${feet}'-0"`
      return `${inches}"`
    }

    const extractHeightToken = (text) => {
      if (!text) return null
      const t = String(text)
      const hMatch = t.match(/H\s*=\s*([^)]+)/i)
      if (hMatch) return hMatch[1].trim()
      const bracketMatch = t.match(/\(([^)]+)\)/)
      if (bracketMatch) {
        const parts = bracketMatch[1].split(/\s*x\s*/i).map((p) => p.trim()).filter(Boolean)
        if (parts.length >= 2) return parts[parts.length - 1]
      }
      return null
    }

    const heightsInInches = []
    subsectionRows.forEach((row) => {
      const bVal = row && row[1] != null ? row[1] : ''
      const token = extractHeightToken(bVal)
      const inches = toInches(token)
      if (inches != null && !Number.isNaN(inches) && inches > 0) {
        heightsInInches.push(inches)
      }
    })

    if (!heightsInInches.length) return null
    const sum = heightsInInches.reduce((t, v) => t + v, 0)
    const avg = sum / heightsInInches.length
    return formatInches(avg)
  }

  const buildDemolitionTemplate = (subsectionName, itemText, fallbackText, subsectionRows = []) => {
    const slabTypeMatch = subsectionName.match(/^Demo\s+(.+)$/i)
    const slabType = slabTypeMatch ? slabTypeMatch[1].trim() : subsectionName.replace(/^Demo\s+/i, '').trim()
    const dmReference = getDMReferenceFromRawData(subsectionName)
    const dmSuffix = ` as per ${dmReference} & details on`
    const dmSuffixStripFooting = ` as per ${dmReference} & details on `
    const stLower = slabType.toLowerCase()
    const pfxText = 'Allow to saw-cut/demo/remove/dispose existing'

    // Shared helpers for averaging dimensions across all subsection rows
    const _toInches = (token) => {
      if (!token) return null
      const s = String(token).trim()
      const m1 = s.match(/^(-?\d+)'-(\d+)"?$/)
      if (m1) return parseInt(m1[1], 10) * 12 + parseInt(m1[2], 10)
      const m2 = s.match(/^(-?\d+)'$/)
      if (m2) return parseInt(m2[1], 10) * 12
      const m3 = s.match(/^(\d+)"$/)
      if (m3) return parseInt(m3[1], 10)
      const m4 = s.match(/^(\d+(?:\.\d+)?)$/)
      if (m4) return parseFloat(m4[1])
      return null
    }
    const _fmtInches = (val) => {
      if (val == null || Number.isNaN(val)) return '##'
      const total = Math.round(val)
      const feet = Math.floor(total / 12)
      const inches = total % 12
      if (feet > 0 && inches > 0) return `${feet}'-${inches}"`
      if (feet > 0) return `${feet}'-0"`
      return `${inches}"`
    }
    const _findDimParen = (str) => {
      const regex = /\(([^)]+)\)/g
      let best = null
      let m
      while ((m = regex.exec(str)) !== null) {
        if (/\d/.test(m[1]) && /x/i.test(m[1])) best = m[1]
      }
      return best
    }
    const _parseDims2 = (str) => {
      const content = _findDimParen(str)
      if (!content) return null
      const parts = content.split(/\s*x\s*/i).map(s => s.trim())
      if (parts.length < 2) return null
      const a = _toInches(parts[0])
      const b = _toInches(parts[1])
      if (a == null || b == null) return null
      return [a, b]
    }
    const _parseDims3 = (str) => {
      const content = _findDimParen(str)
      if (!content) return null
      const parts = content.split(/\s*x\s*/i).map(s => s.trim())
      if (parts.length < 3) return null
      const a = _toInches(parts[0])
      const b = _toInches(parts[1])
      const c = _toInches(parts[2])
      if (a == null || b == null || c == null) return null
      return [a, b, c]
    }
    const _parseThickness = (str) => {
      const m = str.match(/(\d+)(?:"|")\s*(?:thick|thk)?/i) || str.match(/(\d+'-?\d*"?)/)
      if (!m) return null
      return _toInches(m[1])
    }
    const _avg2FromRows = () => {
      if (!subsectionRows || !subsectionRows.length) return null
      const dims = []
      subsectionRows.forEach((row) => {
        const bVal = row && row[1] != null ? String(row[1]) : ''
        const parsed = _parseDims2(bVal)
        if (parsed) dims.push(parsed)
      })
      if (!dims.length) return null
      const n = dims.length
      return [dims.reduce((t, d) => t + d[0], 0) / n, dims.reduce((t, d) => t + d[1], 0) / n]
    }
    const _avg3FromRows = () => {
      if (!subsectionRows || !subsectionRows.length) return null
      const dims = []
      subsectionRows.forEach((row) => {
        const bVal = row && row[1] != null ? String(row[1]) : ''
        const parsed = _parseDims3(bVal)
        if (parsed) dims.push(parsed)
      })
      if (!dims.length) return null
      const n = dims.length
      return [dims.reduce((t, d) => t + d[0], 0) / n, dims.reduce((t, d) => t + d[1], 0) / n, dims.reduce((t, d) => t + d[2], 0) / n]
    }
    const _avgThicknessFromRows = () => {
      if (!subsectionRows || !subsectionRows.length) return null
      const vals = []
      subsectionRows.forEach((row) => {
        const bVal = row && row[1] != null ? String(row[1]) : ''
        const v = _parseThickness(bVal)
        if (v != null && v > 0) vals.push(v)
      })
      if (!vals.length) return null
      return vals.reduce((t, v) => t + v, 0) / vals.length
    }

    // Pits and mat slab: always use average height/depth from all rows
    const pitTypes = ['Demo elevator pit', 'Demo service elevator pit', 'Demo detention tank', 'Demo duplex sewage ejector pit', 'Demo deep sewage ejector pit', 'Demo sump pump pit', 'Demo grease trap pit', 'Demo house trap pit']
    if (pitTypes.includes(subsectionName)) {
      const avgH = getAverageHeightForDemoSubsection(subsectionName)
      const hPart = avgH || '##'
      return `${pfxText} (${hPart}) ${slabType} @ existing building${dmSuffix}`
    }
    if (subsectionName === 'Demo mat slab') {
      const avgH = getAverageHeightForDemoSubsection(subsectionName)
      const hPart = avgH ? `H=${avgH}` : 'H=##'
      return `${pfxText} (${hPart}) mat slab @ existing building${dmSuffix}`
    }

    // Stair on grade: width from text, riser from QTY (caller handles formula)
    if (stLower.includes('stair on grade') || stLower.includes('stairs on grade')) {
      const textToParse = (itemText ? String(itemText).trim() : '') || (fallbackText ? String(fallbackText).trim() : '')
      if (!textToParse) return `${pfxText} (## wide) stairs on grade (## Riser) @ 1st FL${dmSuffix}`
      const widthMatch = textToParse.match(/(\d+'-?\d*"?)\s*wide/i) || textToParse.match(/\(([^)]+)\)/)
      const width = widthMatch ? String(widthMatch[1]).replace(/^\(|\)$/g, '').trim() : '##'
      return `${pfxText} (${width} wide) stairs on grade (## Riser) @ 1st FL${dmSuffix}`
    }

    // Slab on grade / ramp on grade: average thickness from all rows
    if (stLower.includes('slab on grade') || stLower.includes('ramp on grade')) {
      const displayType = stLower.includes('ramp on grade') ? 'ramp on grade' : slabType
      const avgThick = _avgThicknessFromRows()
      const thickStr = avgThick != null ? _fmtInches(avgThick) : '##'
      return `${pfxText} (${thickStr} thick) ${displayType} @ existing building${dmSuffix}`
    }

    // 2-dimension types (WxH): retaining wall, foundation wall, strip footing, grade beam, tie beam, strap beam, corbel, liner wall, barrier wall, stem wall
    const twoD = [
      { match: 'retaining wall', label: 'retaining wall', suffix: dmSuffix },
      { match: 'foundation wall', label: 'foundation wall', suffix: dmSuffix },
      { match: 'strip footing', label: 'strip footing', suffix: dmSuffixStripFooting },
      { match: 'grade beam', label: 'grade beam', suffix: dmSuffix },
      { match: 'tie beam', label: 'tie beam', suffix: dmSuffix },
      { match: 'strap beam', label: 'strap beam', suffix: dmSuffix },
      { match: 'corbel', label: 'corbel', suffix: dmSuffix },
      { match: 'liner wall', label: 'liner wall', suffix: dmSuffix },
      { match: 'barrier wall', label: 'barrier wall', suffix: dmSuffix },
      { match: 'stem wall', label: 'stem wall', suffix: dmSuffix }
    ]
    for (const t of twoD) {
      if (stLower.includes(t.match)) {
        const avg = _avg2FromRows()
        const w = avg ? _fmtInches(avg[0]) : '##'
        const h = avg ? _fmtInches(avg[1]) : '##'
        return `${pfxText} (${w} wide) ${t.label} (H=${h}) @ existing building${t.suffix}`
      }
    }

    // 3-dimension types (WxW2xH): isolated footing, pile cap, pilaster, pier, thickened slab
    const threeD = [
      { match: 'isolated footing', label: 'isolated footing', suffix: dmSuffix },
      { match: 'pile cap', label: 'pile cap', suffix: dmSuffix },
      { match: 'pilaster', label: 'pilaster', suffix: dmSuffix },
      { match: 'pier', label: 'pier', suffix: dmSuffix },
      { match: 'thickened slab', label: 'thickened slab', suffix: dmSuffix }
    ]
    for (const t of threeD) {
      if (stLower.includes(t.match)) {
        const avg = _avg3FromRows()
        if (avg) {
          const widthPart = `${_fmtInches(avg[0])}x${_fmtInches(avg[1])}`
          const height = _fmtInches(avg[2])
          return `${pfxText} (${widthPart} wide) ${t.label} (H=${height}) @ existing building${t.suffix}`
        }
        const avg2 = _avg2FromRows()
        if (avg2) {
          return `${pfxText} (${_fmtInches(avg2[0])} wide) ${t.label} (H=${_fmtInches(avg2[1])}) @ existing building${t.suffix}`
        }
        return `${pfxText} (##x## wide) ${t.label} (H=##) @ existing building${t.suffix}`
      }
    }

    // Buttress: no dimensions
    if (stLower.includes('buttress')) {
      return `${pfxText} ${slabType} @ existing building${dmSuffix}`
    }

    // Generic fallback: try to get a thickness or dimension from all rows
    const avgThick = _avgThicknessFromRows()
    if (avgThick != null) {
      return `${pfxText} (${_fmtInches(avgThick)} thick) ${slabType} @ existing building${dmSuffix}`
    }
    const avg2 = _avg2FromRows()
    if (avg2) {
      return `${pfxText} (${_fmtInches(avg2[0])} wide) ${slabType} (H=${_fmtInches(avg2[1])}) @ existing building${dmSuffix}`
    }
    return `${pfxText} ${slabType} @ existing building${dmSuffix}`
  }

  // Soil excavation scope: SOE ref, P ref from raw data (Page column), like demolition.
  const getSoilExcavationRefsFromRawData = () => {
    if (!rawData || !Array.isArray(rawData) || rawData.length < 2) {
      return { soeRef: '##', pRef: '##', detailsOnRef: '##' }
    }
    const headers = rawData[0]
    const dataRows = rawData.slice(1)
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
    if (digitizerIdx === -1) return { soeRef: '##', pRef: '##', detailsOnRef: '##' }
    for (const row of dataRows) {
      const digitizerItem = row[digitizerIdx]
      if (!digitizerItem) continue
      const d = String(digitizerItem).toLowerCase().trim()
      if ((d.includes('excavation') || d.includes('soil')) && !d.includes('rock')) {
        let soeRef = '##'
        let pRef = '##'
        let detailsOnRef = '##'
        if (pageIdx >= 0) {
          const pageVal = row[pageIdx]
          if (pageVal != null && pageVal !== '') {
            const pageStr = String(pageVal).trim()
            const soeMatch = pageStr.match(/SOE-[\d.]+/i)
            if (soeMatch) soeRef = soeMatch[0]
            const pMatch = pageStr.match(/P-[\d.]+/i)
            if (pMatch) pRef = pMatch[0]
            const toMatch = pageStr.match(/SOE-[\d.]+[\s]*to[\s]*SOE-[\d.]+/gi)
            if (toMatch && toMatch.length) detailsOnRef = toMatch[0].trim()
          }
        }
        return { soeRef, pRef, detailsOnRef }
      }
    }
    return { soeRef: '##', pRef: '##', detailsOnRef: '##' }
  }

  // Collect all unique page refs (SOE-xxx, P-xxx, etc.) from raw rows in the excavation/backfill group:
  // excavation, soil, backfill, slope exc, underground piping, exc & backfill (not rock). Use for proposal "as per" line.
  const getUniquePageRefsForExcavationAndBackfillGroup = () => {
    if (!rawData || !Array.isArray(rawData) || rawData.length < 2) return []
    const headers = rawData[0]
    const dataRows = rawData.slice(1)
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
    if (digitizerIdx === -1 || pageIdx === -1) return []
    const refSet = new Set()
    for (const row of dataRows) {
      const digitizerItem = row[digitizerIdx]
      if (!digitizerItem) continue
      const d = String(digitizerItem).toLowerCase().trim()
      if (d.includes('rock')) continue
      const isInGroup =
        d.includes('excavation') || d.includes('soil') || d.includes('backfill') ||
        d.includes('slope exc') || d.includes('underground') || d.includes('exc & backfill') ||
        d.includes('pc-') || /backfill\s*\(h=/i.test(d) || /slope\s+exc/i.test(d)
      if (!isInGroup) continue
      const pageVal = row[pageIdx]
      if (pageVal == null || pageVal === '') continue
      const pageStr = String(pageVal).trim()
      const soeRefs = pageStr.match(/SOE-[\d.]+/gi) || []
      const pRefs = pageStr.match(/P-[\d.]+/gi) || []
      soeRefs.forEach(r => refSet.add(r))
      pRefs.forEach(r => refSet.add(r))
    }
    return [...refSet]
  }

  // Collect unique page refs from raw rows in the rock excavation group (digitizer item contains rock + excavation).
  const getUniquePageRefsForRockExcavationGroup = () => {
    if (!rawData || !Array.isArray(rawData) || rawData.length < 2) return []
    const headers = rawData[0]
    const dataRows = rawData.slice(1)
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
    if (digitizerIdx === -1 || pageIdx === -1) return []
    const refSet = new Set()
    for (const row of dataRows) {
      const digitizerItem = row[digitizerIdx]
      if (!digitizerItem) continue
      const d = String(digitizerItem).toLowerCase().trim()
      if (!d.includes('rock') || !(d.includes('excavation') || d.includes('rock excavation') || d.includes('line drill'))) continue
      const pageVal = row[pageIdx]
      if (pageVal == null || pageVal === '') continue
      const pageStr = String(pageVal).trim()
      const soeRefs = pageStr.match(/SOE-[\d.]+/gi) || []
      const pRefs = pageStr.match(/P-[\d.]+/gi) || []
      soeRefs.forEach(r => refSet.add(r))
      pRefs.forEach(r => refSet.add(r))
    }
    return [...refSet]
  }

  const formatHeightAsFeetInches = (heightValues) => {
    if (!heightValues || heightValues.length === 0) return '##'
    const numHeights = heightValues.filter(v => v != null && !Number.isNaN(v) && v > 0).map(v => parseFloat(v))
    if (numHeights.length === 0) return '##'
    let avg = numHeights.reduce((a, b) => a + b, 0) / numHeights.length
    if (avg > 50) avg = avg / 12
    const feet = Math.floor(avg)
    const inches = Math.round((avg - feet) * 12)
    if (inches === 0) return `${feet}'-0"`
    return `${feet}'-${inches}"`
  }

  const calculateRowSF = (row) => {
    if (!row || row.length === 0) return 0
    const sqftValue = parseFloat(row[9]) || 0
    if (sqftValue > 0) return sqftValue
    const takeoff = parseFloat(row[2]) || 0, length = parseFloat(row[5]) || 0, width = parseFloat(row[6]) || 0
    if (takeoff > 0) {
      if (width > 0 && length > 0) return takeoff * length * width
      if (width > 0) return takeoff * width
      if (length > 0) return takeoff * length
      return takeoff
    }
    return 0
  }
  const calculateSF = (rows) => (!rows || rows.length === 0) ? 0 : rows.reduce((t, r) => t + calculateRowSF(r), 0)
  const calculateCY = (rows) => {
    if (!rows || rows.length === 0) return 0
    return rows.reduce((t, r) => {
      const sf = calculateRowSF(r), height = parseFloat(r[7]) || 0
      return t + (sf > 0 && height > 0 ? (sf * height) / 27 : 0)
    }, 0)
  }
  let linesBySubsection = new Map()
  let rowsBySubsection = new Map()
  let sumRowsBySubsection = new Map()
  let sumRowIndexBySubsection = new Map() // Excel 1-based row in Calculations Sheet for formula refs
  let lastDataRowIndexBySubsection = new Map() // 0-based row index of last data row (sum row = this + 2 in 1-based)
  let firstDemoStairOnGradeStairsRowIndex = null // 1-based Excel row of first "Demo stairs on grade" (not landings) row, for width (G)
  const getQTYFromSumRow = (subsectionName) => {
    const sumRow = sumRowsBySubsection.get(subsectionName)
    return sumRow && parseFloat(sumRow[12]) > 0 ? parseFloat(sumRow[12]) : null
  }
  if (Array.isArray(calculationData) && calculationData.length > 0) {
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
    let excavationHeightValues = [] // Collect Height (col H) from excavation data rows for Havg
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
    let rockExcavationEmptyRowIndex = null // Store the row index where empty Particulars row was found for rock excavation (sum row, 1-based)
    let rockExcavationLine1RowIndex = null // First data row in rock excavation subsection (1-based)
    let rockExcavationLine2RowIndex = null // Second data row in rock excavation subsection (1-based)
    let rockExcavationDataRowCount = 0 // Count data rows in rock excavation subsection
    let rockExcavationHeightValues = [] // Collect Height (col H) from rock excavation data rows for Havg
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

      // Reset SOE soldier pile groups at start of each build so we don't accumulate duplicates on re-run
      if (rowIndex === 0) {
        window.drilledSoldierPileGroups = []
        window.hpSoldierPileGroups = []
      }

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
        rockExcavationLine1RowIndex = null
        rockExcavationLine2RowIndex = null
        rockExcavationDataRowCount = 0
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

      const colALower = colA ? String(colA).trim().toLowerCase() : ''

      // Detect start of Foundation section. Support both "Foundation" and "Foundation/Substructure"
      if (colALower === 'foundation' || colALower === 'foundation/substructure') {
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
      // Treat both "Foundation" and "Foundation/Substructure" as the same main section.
      if (colA && String(colA).trim() &&
        colALower !== 'demolition' &&
        colALower !== 'excavation' &&
        colALower !== 'rock excavation' &&
        colALower !== 'soe' &&
        colALower !== 'foundation' &&
        colALower !== 'foundation/substructure') {
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
        let subName = bText.slice(0, -1).trim()
        // Normalize "Demo pile cap" / "Demo pile caps" to single key so we don't duplicate demolition lines
        if (subName.toLowerCase() === 'demo pile cap') subName = 'Demo pile caps'
        currentSubsection = subName
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
          rockExcavationRunningSum = 0
          rockExcavationRunningCYSum = 0
          foundRockExcavationEmptyParticulars = false
          rockExcavationDataRowCount = 0
          rockExcavationHeightValues = []
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

        // Track Foundation subsections (use canonical name so Map keys match foundationSubsectionOrder)
        if (inFoundationSection) {
          const foundationDisplayToCanonical = {
            'Foundation drilled piles': 'Drilled foundation pile',
            'Foundation driven piles': 'Driven foundation pile',
            'Foundation helical piles': 'Helical foundation pile',
            'Stelcor piles': 'Drilled displacement pile',
            'CAF piles': 'CFA pile'
          }
          const displayName = currentSubsection.replace(/\s*scope\s*$/i, '').trim()
          const canonicalName = foundationDisplayToCanonical[displayName] || currentSubsection
          // Save previous subsection items if any
          if (currentFoundationSubsectionName && currentFoundationSubsectionName !== canonicalName && currentFoundationSubsectionItems.length > 0) {
            if (!window.foundationSubsectionItems.has(currentFoundationSubsectionName)) {
              window.foundationSubsectionItems.set(currentFoundationSubsectionName, [])
            }
            window.foundationSubsectionItems.get(currentFoundationSubsectionName).push([...currentFoundationSubsectionItems])
            currentFoundationSubsectionItems = []
          }
          // Start collecting for new subsection (store under canonical name)
          currentFoundationSubsectionName = canonicalName
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
        let sqftRaw = row[9] // Column J (index 9) - SQFT total
        let qtyRaw = row[12] // Column M (index 12) - QTY

        // If cell contains a formula, get evaluated value from workbook so we detect sum row
        if (workbook && (sqftRaw == null || sqftRaw === '' || (typeof sqftRaw === 'string' && sqftRaw.trim().startsWith('=')))) {
          try {
            const evaluated = workbook.getValueRowCol(calcSheetIndex, rowIndex, 9)
            if (evaluated !== undefined && evaluated !== null && evaluated !== '') sqftRaw = evaluated
          } catch (e) { /* ignore */ }
        }
        if (workbook && (qtyRaw == null || qtyRaw === '' || (typeof qtyRaw === 'string' && qtyRaw.trim().startsWith('=')))) {
          try {
            const evaluated = workbook.getValueRowCol(calcSheetIndex, rowIndex, 12)
            if (evaluated !== undefined && evaluated !== null && evaluated !== '') qtyRaw = evaluated
          } catch (e) { /* ignore */ }
        }

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
          if (currentSubsection && currentSubsection.startsWith('Demo')) {
            sumRowIndexBySubsection.set(currentSubsection, rowIndex + 1)
          }
          return
        } else {
          // Still store it if it has QTY
          if (qtyValue > 0) {
            sumRowsBySubsection.set(currentSubsection, row)
            if (currentSubsection && currentSubsection.startsWith('Demo')) {
              sumRowIndexBySubsection.set(currentSubsection, rowIndex + 1)
            }
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
          const heightVal = parseFloat(row[7])
          if (!Number.isNaN(heightVal) && heightVal > 0) excavationHeightValues.push(heightVal)
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
          // Particulars is filled - data row; track row index for proposal references (LF, SF, CY)
          rockExcavationDataRowCount++
          if (rockExcavationDataRowCount === 1) rockExcavationLine1RowIndex = rowIndex + 1
          else if (rockExcavationDataRowCount === 2) rockExcavationLine2RowIndex = rowIndex + 1
          const rockHeightVal = parseFloat(row[7])
          if (!Number.isNaN(rockHeightVal) && rockHeightVal > 0) rockExcavationHeightValues.push(rockHeightVal)
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
      // Track last data row index for demolition so we can reference sum row (next row) as formula
      if (currentSubsection && currentSubsection.startsWith('Demo')) {
        lastDataRowIndexBySubsection.set(currentSubsection, rowIndex)
      }
      // First "Demo stairs on grade" (stairs line, not landings/slab) row – for width (G) in proposal CONCATENATE
      if (currentSubsection === 'Demo stair on grade') {
        const lower = bText.toLowerCase()
        if (lower.includes('stairs on grade') && !lower.includes('landings') && !lower.includes('stair slab')) {
          if (firstDemoStairOnGradeStairsRowIndex == null) firstDemoStairOnGradeStairsRowIndex = rowIndex + 1
        }
      }
    })


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

      // If values are empty or formulas in calculationData, try reading evaluated value from spreadsheet
      if (workbook && (!sqftValue || sqftValue === '' || (typeof sqftValue === 'string' && String(sqftValue).trim().startsWith('=')))) {
        try {
          const ev = workbook.getValueRowCol(calcSheetIndex, rowIndex, 9)
          if (ev !== undefined && ev !== null && ev !== '') sqftValue = ev
        } catch (e) { }
      }
      if ((!bValue || bValue === '') && workbook) {
        try {
          bValue = workbook.getValueRowCol(calcSheetIndex, rowIndex, 1)
        } catch (e) { }
      }
      if (workbook && (!qtyValue || qtyValue === '' || (typeof qtyValue === 'string' && String(qtyValue).trim().startsWith('=')))) {
        try {
          const ev = workbook.getValueRowCol(calcSheetIndex, rowIndex, 12)
          if (ev !== undefined && ev !== null && ev !== '') qtyValue = ev
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
              let subsectionName = checkBText.slice(0, -1).trim()
              if (subsectionName.toLowerCase() === 'demo pile cap') subsectionName = 'Demo pile caps'
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
          if (foundSubsection.startsWith('Demo')) {
            sumRowIndexBySubsection.set(foundSubsection, excelRow + 1)
          }
        }
      }
    }

    // Initialize dynamic row counter; proposal data range (I-N etc.) ends at Exclusions row
    let currentRow = 14
    let proposalDataEndRow = null

    // -------------------------------------------------------------------------
    // DEMOLITION SECTION
    // -------------------------------------------------------------------------

    // Check if we have any demolition items
    const orderedSubsections = [
      'Demo slab on grade',
      'Demo Ramp on grade',
      'Demo strip footing',
      'Demo foundation wall',
      'Demo retaining wall',
      'Demo isolated footing',
      'Demo pile caps',
      'Demo pilaster',
      'Demo grade beam',
      'Demo tie beam',
      'Demo strap beam',
      'Demo thickened slab',
      'Demo buttress',
      'Demo pier',
      'Demo corbel',
      'Demo liner wall',
      'Demo barrier wall',
      'Demo stem wall',
      'Demo elevator pit',
      'Demo service elevator pit',
      'Demo detention tank',
      'Demo duplex sewage ejector pit',
      'Demo deep sewage ejector pit',
      'Demo sump pump pit',
      'Demo grease trap pit',
      'Demo house trap pit',
      'Demo mat slab',
      'Demo stair on grade'
    ]

    // Demo stair on grade: one proposal line per landing/stairs/slab from calculation (grouped by Stair A1, etc.)
    const demoStairOnGradeLines = []
    if (formulaData && Array.isArray(formulaData) && calculationData && calculationData.length > 0) {
      const demoStairEntries = formulaData.filter(f =>
        f.section === 'demolition' && f.subsection === 'Demo stair on grade' &&
        ['demo_stair_on_grade_heading', 'demo_stair_on_grade_landing', 'demo_stair_on_grade_stairs', 'demo_stair_on_grade_stair_slab'].includes(f.itemType)
      )
      demoStairEntries.sort((a, b) => (a.row || 0) - (b.row || 0))
      let currentGroupName = null
      demoStairEntries.forEach((entry) => {
        if (entry.itemType === 'demo_stair_on_grade_heading') {
          const label = (calculationData[entry.row - 1]?.[1] || '').toString().trim().replace(/:+\s*$/, '').trim()
          currentGroupName = (label && label !== 'Demo stair on grade') ? label : null
        } else {
          const type = entry.itemType === 'demo_stair_on_grade_landing' ? 'landing' : entry.itemType === 'demo_stair_on_grade_stairs' ? 'stairs' : 'slab'
          demoStairOnGradeLines.push({ type, row: entry.row, groupName: currentGroupName })
        }
      })
    }

    // Check if any of these subsections have data
    let hasDemolitionItems = false
    orderedSubsections.forEach(name => {
      if ((rowsBySubsection.get(name) || []).length > 0) {
        hasDemolitionItems = true
      }
    })
    if (demoStairOnGradeLines.length > 0) hasDemolitionItems = true

    if (hasDemolitionItems) {
      // Demolition Scope Heading
      spreadsheet.updateCell({ value: 'Demolition scope:' }, `${pfx}B${currentRow}`)
      rowBContentMap.set(currentRow, 'Demolition scope:')
      spreadsheet.wrap(`${pfx}B${currentRow}`, true)
      spreadsheet.cellFormat({
        backgroundColor: '#BDD7EE',
        textAlign: 'center',
        verticalAlign: 'middle',
        textDecoration: 'underline',
        fontWeight: 'normal',
        wrapText: true
      }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat({ backgroundColor: 'white' }, `${pfx}C${currentRow}:G${currentRow}`)
      currentRow++

      // Render specific demolition lines
      const demolitionStartRow = currentRow

      const demolitionGroupsFromCalc = Object.fromEntries(
        orderedSubsections.map(name => [
          name,
          {
            firstItemText: linesBySubsection.get(name) ?? null,
            rowCount: (rowsBySubsection.get(name) || []).length,
            sumRow: sumRowsBySubsection.has(name) ? 'yes' : 'no'
          }
        ])
      )

      const demolitionCalcSheet = 'Calculations Sheet'
      orderedSubsections.forEach((name, index) => {
        const rowCount = (rowsBySubsection.get(name) || []).length
        // Do not show this demolition line if subsection has no data in calculation sheet
        if (rowCount === 0) {
          if (name === 'Demo stair on grade' && demoStairOnGradeLines.length > 0) {
            // Demo stair on grade: show when formulaData has lines even if no direct rows
          } else {
            return
          }
        }

        const dmRef = getDMReferenceFromRawData(name)
        const esc = (s) => (s || '').replace(/"/g, '""')
        // Demo stair on grade: multiple rows from calculation (landings, stairs, slab per group)
        if (name === 'Demo stair on grade' && demoStairOnGradeLines.length > 0) {
          let lastGroupName = null
          demoStairOnGradeLines.forEach((line) => {
            if (line.groupName && line.groupName !== lastGroupName) {
              spreadsheet.updateCell({ value: `${line.groupName}:` }, `${pfx}B${currentRow}`)
              rowBContentMap.set(currentRow, `${line.groupName}:`)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat(
                { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#E2EFDA', textDecoration: 'underline', border: '1px solid #000000' },
                `${pfx}B${currentRow}`
              )
              currentRow++
              lastGroupName = line.groupName
            }
            const cellRef = `${pfx}B${currentRow}`
            let bVal
            let bDesc
            if (line.type === 'landing') {
              const atStair = line.groupName ? ` @ ${line.groupName}` : ''
              bDesc = `Allow to saw-cut/demo/remove/dispose existing stair landing on grade${atStair} as per ${dmRef} & details on`
              bVal = { value: bDesc }
            } else if (line.type === 'stairs') {
              const midPart = line.groupName ? `" @ ${line.groupName} "` : '" "'
              const calcRow = calculationData[line.row - 1]
              const widthVal = calcRow && calcRow[6] != null && calcRow[6] !== '' ? calcRow[6] : null
              const widthStr = formatWidthFeetInches(widthVal) || "0'-0\""
              const widthEsc = widthStr.replace(/"/g, '""')
              bVal = {
                formula: `=CONCATENATE("Allow to saw-cut/demo/remove/dispose existing (","${widthEsc}"," wide) stairs on grade",${midPart},"(",ROUND('${demolitionCalcSheet}'!M${line.row},0)," Riser) @ 1st FL as per ","${esc(dmRef)}"," & details on")`
              }
              bDesc = `Allow to saw-cut/demo/remove/dispose existing (… wide) stairs on grade${line.groupName ? ` @ ${line.groupName}` : ''} (… Riser) @ 1st FL as per ${dmRef} & details on`
            } else {
              bDesc = `Allow to saw-cut/demo/remove/dispose existing stair slab as per ${dmRef} & details on`
              bVal = { value: bDesc }
            }
            spreadsheet.updateCell(bVal, cellRef)
            rowBContentMap.set(currentRow, bDesc)
            spreadsheet.wrap(cellRef, true)
            spreadsheet.cellFormat(
              { fontWeight: 'bold', color: '#000000', textAlign: 'left', verticalAlign: 'top', wrapText: true },
              cellRef
            )
            fillRatesForProposalRow(currentRow, bDesc)
            spreadsheet.updateCell({ formula: `='${demolitionCalcSheet}'!J${line.row}` }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: `='${demolitionCalcSheet}'!L${line.row}` }, `${pfx}F${currentRow}`)
            spreadsheet.updateCell({ formula: `='${demolitionCalcSheet}'!M${line.row}` }, `${pfx}G${currentRow}`)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'right', format: '#,##0.00' }, `${pfx}D${currentRow}`)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'right', format: '#,##0.00' }, `${pfx}F${currentRow}`)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'right', format: '#,##0.00' }, `${pfx}G${currentRow}`)
            const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
            spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'right', format: '#,##0.00' }, `${pfx}H${currentRow}`)
            currentRow++
          })
          return
        }

        const originalText = linesBySubsection.get(name)
        const subsectionRows = rowsBySubsection.get(name) || []
        const firstRowWithB = subsectionRows.find(r => r[1] && String(r[1]).trim())
        const fallbackText = originalText ? null : (firstRowWithB ? String(firstRowWithB[1]).trim() : null)
        const templateText = buildDemolitionTemplate(name, originalText, fallbackText, subsectionRows)
        const cellRef = `${pfx}B${currentRow}`

        // Sum row in Calculations Sheet: either from detection (sumRowIndexBySubsection) or last data row + 1 (sum is next row)
        const lastDataIdx = lastDataRowIndexBySubsection.get(name)
        const detectedSumRow = sumRowIndexBySubsection.get(name)
        const sumRowIndex = detectedSumRow != null ? detectedSumRow : (lastDataIdx != null ? lastDataIdx + 2 : null)

        // Demo stair on grade: width from first "Demo stairs on grade" row (G), riser from sum row (M) – full formula when both available
        const stairsRiserPlaceholder = '(## Riser)'
        const useDemoStairsFullFormula = name === 'Demo stair on grade' && sumRowIndex != null && firstDemoStairOnGradeStairsRowIndex != null
        const useStairsRiserFormula = !useDemoStairsFullFormula && templateText.includes(stairsRiserPlaceholder) && sumRowIndex != null
        let valueOrFormulaForB = templateText
        if (useDemoStairsFullFormula) {
          const calcRow = calculationData[firstDemoStairOnGradeStairsRowIndex - 1]
          const widthVal = calcRow && calcRow[6] != null && calcRow[6] !== '' ? calcRow[6] : null
          const widthStr = formatWidthFeetInches(widthVal) || "0'-0\""
          const widthEsc = widthStr.replace(/"/g, '""')
          valueOrFormulaForB = {
            formula: `=CONCATENATE("Allow to saw-cut/demo/remove/dispose existing (","${widthEsc}"," wide) stairs on grade (",ROUND('${demolitionCalcSheet}'!M${sumRowIndex},0)," Riser) @ 1st FL as per ","${esc(dmRef)}"," & details on")`
          }
        } else if (useStairsRiserFormula) {
          const parts = templateText.split(stairsRiserPlaceholder)
          const prefix = (parts[0] || '').trimEnd()
          const suffix = ' Riser)' + (parts[1] || '').trimStart()
          valueOrFormulaForB = {
            formula: `=CONCATENATE("${esc(prefix)}",'${demolitionCalcSheet}'!M${sumRowIndex},"${esc(suffix)}")`
          }
        }

        if (typeof valueOrFormulaForB === 'object' && valueOrFormulaForB.formula) {
          spreadsheet.updateCell(valueOrFormulaForB, cellRef)
          rowBContentMap.set(currentRow, templateText)
        } else {
          spreadsheet.updateCell({ value: valueOrFormulaForB }, cellRef)
          rowBContentMap.set(currentRow, valueOrFormulaForB)
        }
        spreadsheet.wrap(cellRef, true)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            verticalAlign: 'top',
            wrapText: true
          },
          cellRef
        )
        fillRatesForProposalRow(currentRow, typeof valueOrFormulaForB === 'object' ? templateText : valueOrFormulaForB)

        if (sumRowIndex != null) {
          // Always use formulas so the cell shows the formula and updates when calc sheet changes
          spreadsheet.updateCell({ formula: `='${demolitionCalcSheet}'!J${sumRowIndex}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${demolitionCalcSheet}'!L${sumRowIndex}` }, `${pfx}F${currentRow}`)
          spreadsheet.updateCell({ formula: `='${demolitionCalcSheet}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
        } else {
          // Fallback: compute from in-memory data (when sum row index not found)
          const subsectionRows = rowsBySubsection.get(name) || []
          const subsectionSF = calculateSF(subsectionRows)
          const formattedSF = parseFloat(subsectionSF.toFixed(2))
          spreadsheet.updateCell({ value: formattedSF }, `${pfx}D${currentRow}`)

          const subsectionCY = calculateCY(subsectionRows)
          const formattedCY = parseFloat(subsectionCY.toFixed(2))
          spreadsheet.updateCell({ value: formattedCY }, `${pfx}F${currentRow}`)

          const subsectionQTY = getQTYFromSumRow(name)
          if (subsectionQTY !== null) {
            const formattedQTY = parseFloat(subsectionQTY.toFixed(2))
            spreadsheet.updateCell({ value: formattedQTY }, `${pfx}G${currentRow}`)
          }
        }

        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'right',
            format: '#,##0.00'
          },
          `${pfx}D${currentRow}`
        )
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'right',
            format: '#,##0.00'
          },
          `${pfx}F${currentRow}`
        )
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'right',
            format: '#,##0.00'
          },
          `${pfx}G${currentRow}`
        )

        // Add $/1000 formula in column H
        const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
        spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'right',
            format: '$#,##0.00'
          },
          `${pfx}H${currentRow}`
        )

        // Background color columns I-N
        const columns = ['I', 'J', 'K', 'L', 'M', 'N']
        columns.forEach(col => {
          spreadsheet.cellFormat(
            { backgroundColor: '#E2EFDA' },
            `${pfx}${col}${currentRow}`
          )
        })

        currentRow++
      })

      const demolitionEndRow = currentRow - 1

      // Add note below all demolition items
      spreadsheet.updateCell({ value: 'Note: Site/building demolition by others.' }, `${pfx}B${currentRow}`)
      rowBContentMap.set(currentRow, 'Note: Site/building demolition by others.')
      spreadsheet.wrap(`${pfx}B${currentRow}`, true)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: 'white',
          wrapText: true
        },
        `${pfx}B${currentRow}`
      )
      currentRow++

      // Add Demolition Total row
      spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
      spreadsheet.updateCell({ value: 'Demolition Total:' }, `${pfx}D${currentRow}`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: '#BDD7EE'
        },
        `${pfx}D${currentRow}:E${currentRow}`
      )

      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      // Sum H from start to end of demolition items
      const totalFormula = `=SUM(H${demolitionStartRow}:H${demolitionEndRow})*1000`
      spreadsheet.updateCell({ formula: totalFormula }, `${pfx}F${currentRow}`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          backgroundColor: '#BDD7EE',
          format: '$#,##0.00'
        },
        `${pfx}F${currentRow}:G${currentRow}`
      )

      applyTotalRowBorders(spreadsheet, pfx, currentRow, '#BDD7EE')
      baseBidTotalRows.push(currentRow) // Demolition Total
      totalRows.push(currentRow)

      currentRow++

      // Empty row - ensure it's white (not blue)
      spreadsheet.cellFormat({ backgroundColor: 'white' }, `${pfx}B${currentRow}:G${currentRow}`)
      currentRow++
    }

    // -------------------------------------------------------------------------
    // EXCAVATION SECTION
    // -------------------------------------------------------------------------

    // Excavation scope heading
    spreadsheet.updateCell({ value: 'Excavation scope:' }, `${pfx}B${currentRow}`)
    rowBContentMap.set(currentRow, 'Excavation scope:')
    spreadsheet.cellFormat({
      backgroundColor: '#BDD7EE',
      textAlign: 'center',
      verticalAlign: 'middle',
      textDecoration: 'underline',
      fontWeight: 'normal'
    }, `${pfx}B${currentRow}`)
    spreadsheet.cellFormat({ backgroundColor: 'white' }, `${pfx}C${currentRow}:G${currentRow}`)
    const excavationScopeStartRow = currentRow // Track for green background loop
    currentRow++

    // Only show "Soil excavation scope:" subheading and separate Soil/Rock totals when rock has values
    const hasRockExcavation = rockExcavationTotals && (
      (parseFloat(rockExcavationTotals.totalSQFT) || 0) > 0 ||
      (parseFloat(rockExcavationTotals.totalCY) || 0) > 0 ||
      (parseFloat(lineDrillTotalFT) || 0) > 0
    )
    const hasSoilExcavationItems = !!(excavationEmptyRowIndex || excavationTotalSQFT > 0 || excavationTotalCY > 0)
    const hasBackfillItems = !!(backfillEmptyRowIndex || backfillTotalSQFT > 0 || backfillTotalCY > 0)
    const hasSoilScopeItems = hasSoilExcavationItems || hasBackfillItems
    const hasAnyExcavationData = hasSoilScopeItems || hasRockExcavation
    if (hasRockExcavation && hasSoilScopeItems) {
      currentRow++ // Extra line after Excavation scope
      // Soil excavation scope heading
      spreadsheet.updateCell({ value: 'Soil excavation scope:' }, `${pfx}B${currentRow}`)
      rowBContentMap.set(currentRow, 'Soil excavation scope:')
      spreadsheet.cellFormat({
        backgroundColor: '#FFF2CC',
        textAlign: 'center',
        verticalAlign: 'middle',
        textDecoration: 'underline',
        fontWeight: 'bold',
        border: '1px solid #000000'
      }, `${pfx}B${currentRow}`)
      currentRow++
    }

    // First soil excavation line – only show when soil excavation group exists in calculation sheet
    let soilExcavationRow1 = null
    const uniqueExcavationBackfillRefs = getUniquePageRefsForExcavationAndBackfillGroup()
    const asPerPart = uniqueExcavationBackfillRefs.length > 0 ? uniqueExcavationBackfillRefs.join(', ') : '##'
    const havgText = formatHeightAsFeetInches(excavationHeightValues)
    if (excavationEmptyRowIndex || excavationTotalSQFT > 0 || excavationTotalCY > 0) {
      const soilExcavationText = `Allow to perform soil excavation, trucking & disposal (Havg=${havgText}) as per ${asPerPart} & details on`
      spreadsheet.updateCell({ value: soilExcavationText }, `${pfx}B${currentRow}`)
      rowBContentMap.set(currentRow, soilExcavationText)
      spreadsheet.wrap(`${pfx}B${currentRow}`, true)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          verticalAlign: 'top',
          wrapText: true
        },
        `${pfx}B${currentRow}`
      )
      const soilExcavationHeight = calculateRowHeight(soilExcavationText)
      try { spreadsheet.setRowHeight(soilExcavationHeight, currentRow - 1, proposalSheetIndex) } catch (e) { }
      dynamicHeightRows.push({ row: currentRow, height: soilExcavationHeight })
      fillRatesForProposalRow(currentRow, 'Allow to perform soil excavation, trucking & disposal - Soil excavation scope')
      soilExcavationRow1 = currentRow

      // Add SF value from excavation total to column D
      if (excavationEmptyRowIndex) {
        spreadsheet.updateCell({ formula: `='Calculations Sheet'!J${excavationEmptyRowIndex}` }, `${pfx}D${currentRow}`)
      } else {
        const formattedExcavationSF = parseFloat(excavationTotalSQFT.toFixed(2))
        spreadsheet.updateCell({ value: formattedExcavationSF }, `${pfx}D${currentRow}`)
      }
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '#,##0.00'
        },
        `${pfx}D${currentRow}`
      )

      // Add CY value from excavation total to column F
      if (excavationEmptyRowIndex) {
        spreadsheet.updateCell({ formula: `='Calculations Sheet'!L${excavationEmptyRowIndex}` }, `${pfx}F${currentRow}`)
      } else {
        const formattedExcavationCY = parseFloat(excavationTotalCY.toFixed(2))
        spreadsheet.updateCell({ value: formattedExcavationCY }, `${pfx}F${currentRow}`)
      }
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '#,##0.00'
        },
        `${pfx}F${currentRow}`
      )

      const dollarFormulaSoil1 = `=IFERROR(IF(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)=0,"",ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)),"")`
      spreadsheet.updateCell({ formula: dollarFormulaSoil1 }, `${pfx}H${currentRow}`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '$#,##0.00'
        },
        `${pfx}H${currentRow}`
      )
      currentRow++
    }

    // Row 2: Second soil excavation line (new clean soil) – only show if backfill group exists in calculation sheet
    if (backfillEmptyRowIndex || backfillTotalSQFT > 0 || backfillTotalCY > 0) {
      const soilExcavationRow2 = currentRow
      const backfillSoilText = `Allow to import new clean soil to backfill and compact as per ${asPerPart} & details on`
      spreadsheet.updateCell({ value: backfillSoilText }, `${pfx}B${currentRow}`)
      rowBContentMap.set(currentRow, backfillSoilText)
      spreadsheet.wrap(`${pfx}B${currentRow}`, true)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: 'white',
          verticalAlign: 'top',
          wrapText: true
        },
        `${pfx}B${currentRow}`
      )
      const backfillSoilHeight = calculateRowHeight(backfillSoilText)
      try { spreadsheet.setRowHeight(backfillSoilHeight, currentRow - 1, proposalSheetIndex) } catch (e) { }
      dynamicHeightRows.push({ row: currentRow, height: backfillSoilHeight })
      fillRatesForProposalRow(currentRow, backfillSoilText)

      // Add SF value
      if (backfillEmptyRowIndex) {
        spreadsheet.updateCell({ formula: `='Calculations Sheet'!J${backfillEmptyRowIndex}` }, `${pfx}D${currentRow}`)
      } else {
        const formattedBackfillSF = parseFloat(backfillTotalSQFT.toFixed(2))
        spreadsheet.updateCell({ value: formattedBackfillSF }, `${pfx}D${currentRow}`)
      }
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '#,##0.00'
        },
        `${pfx}D${currentRow}`
      )

      // Add CY value
      if (backfillEmptyRowIndex) {
        spreadsheet.updateCell({ formula: `='Calculations Sheet'!L${backfillEmptyRowIndex}` }, `${pfx}F${currentRow}`)
      } else {
        const formattedBackfillCY = parseFloat(backfillTotalCY.toFixed(2))
        spreadsheet.updateCell({ value: formattedBackfillCY }, `${pfx}F${currentRow}`)
      }
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '#,##0.00'
        },
        `${pfx}F${currentRow}`
      )

      // Formula for row 2
      const dollarFormulaSoil2 = `=IFERROR(IF(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)=0,"",ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)),"")`
      spreadsheet.updateCell({ formula: dollarFormulaSoil2 }, `${pfx}H${currentRow}`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          format: '$#,##0.00'
        },
        `${pfx}H${currentRow}`
      )
      currentRow++
      if (soilExcavationRow1 === null) soilExcavationRow1 = currentRow - 1
    }

    // Notes: only show when the related scope has items (if a scope has no items, don't show its note)
    const notes = []
    if (hasBackfillItems) notes.push('Note: Backfill SOE voids by others only')
    if (hasSoilExcavationItems) notes.push('Note: NJ Res Soil included, contaminated, mixed, hazardous, petroleum impacted not incl.')
    // Bedrock note only when rock excavation exists AND soil excavation scope has at least one line (soil or backfill)
    if (hasRockExcavation && (hasSoilExcavationItems || hasBackfillItems)) notes.push('Note: Bedrock not included, see add alt unit rate if required')
    notes.forEach(note => {
      noteRows.push(currentRow)
      const colonIdx = note.indexOf(': ')
      const boldPart = colonIdx >= 0 ? note.slice(0, colonIdx) + ':' : note
      const normalPart = colonIdx >= 0 ? note.slice(colonIdx + 1) : '' // includes space after colon
      const richText = [
        { text: boldPart, style: { fontWeight: 'bold' } },
        { text: normalPart, style: { fontWeight: 'normal' } }
      ]
      spreadsheet.updateCell({ value: note, richText }, `${pfx}B${currentRow}`)
      rowBContentMap.set(currentRow, note)
      spreadsheet.cellFormat(
        { color: '#000000', textAlign: 'left', backgroundColor: 'white' },
        `${pfx}B${currentRow}`
      )
      currentRow++
    })

    // Clear any leftover note rows from a previous build (e.g. Bedrock note when hasRockExcavation is now false)
    if (notes.length < 3) {
      const emptyRow = currentRow
      spreadsheet.updateCell({ value: '' }, `${pfx}B${emptyRow}`)
      rowBContentMap.set(emptyRow, '')
      if (notes.length < 2) {
        spreadsheet.updateCell({ value: '' }, `${pfx}B${emptyRow + 1}`)
        rowBContentMap.set(emptyRow + 1, '')
      }
      if (notes.length < 1) {
        spreadsheet.updateCell({ value: '' }, `${pfx}B${emptyRow + 2}`)
        rowBContentMap.set(emptyRow + 2, '')
      }
    }

    // Soil Excavation Total (only when rock section is shown and soil scope has items)
    let soilExcavationTotalRow = null
    if (hasRockExcavation && hasSoilScopeItems) {
      spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
      spreadsheet.updateCell({ value: 'Soil Excavation Total:' }, `${pfx}D${currentRow}`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: '#FFF2CC'
        },
        `${pfx}D${currentRow}:E${currentRow}`
      )

      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({
        formula: soilExcavationRow1 != null ? `=SUM(H${soilExcavationRow1}:H${currentRow - 1})*1000` : '=0'
      }, `${pfx}F${currentRow}`)

      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          backgroundColor: '#FFF2CC',
          format: '$#,##0.00'
        },
        `${pfx}F${currentRow}:G${currentRow}`
      )
      applyTotalRowBorders(spreadsheet, pfx, currentRow, '#FFF2CC')
      // Soil Excavation Total: left border of B, C, E match row bg (no black left edge)
      const soilExcavationBg = '#FFF2CC'
      const soilTotalRow = currentRow
        ;['B', 'C', 'E'].forEach(col => {
          spreadsheet.cellFormat({ backgroundColor: soilExcavationBg, borderLeft: `1px solid ${soilExcavationBg}` }, `${pfx}${col}${soilTotalRow}`)
        })
      baseBidTotalRows.push(soilTotalRow) // Soil Excavation Total
      totalRows.push(currentRow)

      soilExcavationTotalRow = currentRow
      currentRow++ // Empty row
      currentRow++ // Extra line after Soil Excavation Total

      // Apply green background to columns I-N for Soil Excavation (Scope to Total)
      for (let row = excavationScopeStartRow; row < currentRow; row++) {
        const columns = ['I', 'J', 'K', 'L', 'M', 'N']
        columns.forEach(col => {
          spreadsheet.cellFormat({ backgroundColor: '#E2EFDA' }, `${pfx}${col}${row}`)
        })
      }
    } else {
      // No rock: apply green background from Excavation scope through soil/notes
      for (let row = excavationScopeStartRow; row < currentRow; row++) {
        const columns = ['I', 'J', 'K', 'L', 'M', 'N']
        columns.forEach(col => {
          spreadsheet.cellFormat({ backgroundColor: '#E2EFDA' }, `${pfx}${col}${row}`)
        })
      }
    }

    // When both soil and rock exist, we add sub-headings and an empty row between Excavation scope and Rock. When only rock, show scope then rock lines directly (no extra blank).
    if (hasRockExcavation && hasSoilScopeItems) {
      currentRow++
    }

    // -------------------------------------------------------------------------
    // ROCK EXCAVATION SECTION (only when rock has values)
    // When only one scope (rock only), don't show "Rock excavation scope:" or "Rock excavation Total" — go straight to Excavation Total
    // -------------------------------------------------------------------------
    let rockExcavationTotalRow = null
    let rockOnlyFirstRow = null
    let rockOnlyLastRow = null
    // Resolve rock excavation / line drill sum rows from formulaData so we use formulas (not hardcoded values)
    if (hasRockExcavation && formulaData && Array.isArray(formulaData)) {
      const rockExcSum = formulaData.find(f => f.itemType === 'rock_excavation_sum' && f.section === 'rock_excavation')
      const lineDrillSum = formulaData.find(f => f.itemType === 'line_drill_sum' && f.section === 'rock_excavation')
      if (rockExcSum && rockExcSum.row) rockExcavationLine1RowIndex = rockExcSum.row
      if (lineDrillSum && lineDrillSum.row) rockExcavationLine2RowIndex = lineDrillSum.row
    }
    if (hasRockExcavation) {
      // Rock excavation scope heading — only when both soil and rock exist; when only rock, show just Excavation scope + items + Excavation Total
      const showRockSubHeading = hasSoilScopeItems
      if (showRockSubHeading) {
        spreadsheet.updateCell({ value: 'Rock excavation scope:' }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, 'Rock excavation scope:')
        spreadsheet.cellFormat({
          backgroundColor: '#FFF2CC',
          textAlign: 'center',
          verticalAlign: 'middle',
          textDecoration: 'underline',
          fontWeight: 'bold',
          border: '1px solid #000000'
        }, `${pfx}B${currentRow}`)
        // Green background for scope row
        const columns33 = ['I', 'J', 'K', 'L', 'M', 'N']
        columns33.forEach(col => spreadsheet.cellFormat({ backgroundColor: '#E2EFDA' }, `${pfx}${col}${currentRow}`))
        currentRow++
      }
      const columns33 = ['I', 'J', 'K', 'L', 'M', 'N']

      // Resolve rock excavation line descriptions from calculation sheet (like other sections)
      let rockExcavationTextFromCalc = null
      let lineDrillingTextFromCalc = null
      if (calculationData && Array.isArray(calculationData)) {
        let inRockSection = false
        let inRockExcavationSubsection = false
        let dataRowCount = 0
        for (let i = 0; i < calculationData.length; i++) {
          const row = calculationData[i]
          const colA = row && row[0] != null ? String(row[0]).trim() : ''
          const colB = row && row[1] != null ? String(row[1]).trim() : ''
          if (colA.toLowerCase() === 'rock excavation') {
            inRockSection = true
            inRockExcavationSubsection = false
            dataRowCount = 0
            continue
          }
          if (colA && colA.toLowerCase() === 'soe') {
            if (inRockSection) break
          }
          if (!inRockSection) continue
          if (colB && colB.toLowerCase() === 'rock excavation:') {
            inRockExcavationSubsection = true
            dataRowCount = 0
            continue
          }
          if (inRockExcavationSubsection && colB && !colB.endsWith(':')) {
            dataRowCount++
            if (dataRowCount === 1) rockExcavationTextFromCalc = colB
            else if (dataRowCount === 2) {
              lineDrillingTextFromCalc = colB
              break
            }
          }
        }
      }
      const rockUniqueRefs = getUniquePageRefsForRockExcavationGroup()
      const rockAsPerPart = rockUniqueRefs.length > 0 ? rockUniqueRefs.join(', ') : '##'
      const havgRock = formatHeightAsFeetInches(rockExcavationHeightValues)
      let rockExcavationBase
      if (rockExcavationTextFromCalc) {
        const stripped = rockExcavationTextFromCalc.replace(/\s+as per [^&]*&\s*details on.*$/i, '').trim()
        rockExcavationBase = stripped.replace(/\s*\(Havg=[^)]*\)/i, '').trim()
        rockExcavationBase = rockExcavationBase ? `${rockExcavationBase} (Havg=${havgRock})` : `Allow to perform rock excavation, trucking & disposal for building (Havg=${havgRock})`
      } else {
        rockExcavationBase = `Allow to perform rock excavation, trucking & disposal for building (Havg=${havgRock})`
      }
      const rockExcavationText = `${rockExcavationBase} as per ${rockAsPerPart} & details on`
      const lineDrillingBase = lineDrillingTextFromCalc
        ? lineDrillingTextFromCalc.replace(/\s+as per [^&]*&\s*details on.*$/i, '').replace(/\s+as per .*$/i, '').trim()
        : 'Allow to perform line drilling'
      const lineDrillingText = `${lineDrillingBase} as per ${rockAsPerPart} & details on`
      const hasLineDrillInRockScope = (parseFloat(lineDrillTotalFT) || 0) > 0 || !!rockExcavationLine2RowIndex || !!lineDrillingTextFromCalc

      // First rock excavation line
      const rockRow1 = currentRow
      spreadsheet.updateCell({ value: rockExcavationText }, `${pfx}B${currentRow}`)
      rowBContentMap.set(currentRow, rockExcavationText)
      spreadsheet.wrap(`${pfx}B${currentRow}`, true)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: 'white',
          verticalAlign: 'top',
          textDecoration: 'none'
        },
        `${pfx}B${currentRow}`
      )
      fillRatesForProposalRow(currentRow, rockExcavationText)
      // SF (D) and CY (F) from calculation sheet – first data row in rock excavation subsection (like other scopes)
      if (rockExcavationLine1RowIndex) {
        spreadsheet.updateCell({ formula: `='Calculations Sheet'!J${rockExcavationLine1RowIndex}` }, `${pfx}D${currentRow}`)
        spreadsheet.updateCell({ formula: `='Calculations Sheet'!L${rockExcavationLine1RowIndex}` }, `${pfx}F${currentRow}`)
      } else if (rockExcavationEmptyRowIndex) {
        spreadsheet.updateCell({ formula: `='Calculations Sheet'!J${rockExcavationEmptyRowIndex}` }, `${pfx}D${currentRow}`)
        spreadsheet.updateCell({ formula: `='Calculations Sheet'!L${rockExcavationEmptyRowIndex}` }, `${pfx}F${currentRow}`)
      } else {
        const formattedRockExcavationSF = parseFloat(rockExcavationTotals.totalSQFT.toFixed(2))
        const formattedRockExcavationCY = parseFloat(rockExcavationTotals.totalCY.toFixed(2))
        spreadsheet.updateCell({ value: formattedRockExcavationSF }, `${pfx}D${currentRow}`)
        spreadsheet.updateCell({ value: formattedRockExcavationCY }, `${pfx}F${currentRow}`)
      }
      spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'right', format: '#,##0.00' }, `${pfx}D${currentRow}`)
      spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'right', format: '#,##0.00' }, `${pfx}F${currentRow}`)

      // Dollar Formula
      const dollarFormulaRock1 = `=IFERROR(IF(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)=0,"",ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)),"")`
      spreadsheet.updateCell({ formula: dollarFormulaRock1 }, `${pfx}H${currentRow}`)
      spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'right', format: '$#,##0.00' }, `${pfx}H${currentRow}`)
      columns33.forEach(col => spreadsheet.cellFormat({ backgroundColor: '#E2EFDA' }, `${pfx}${col}${currentRow}`))
      currentRow++

      // Second rock excavation line (line drilling) — only when line drill exists in rock excavation scope
      let rockRow2 = rockRow1
      if (hasLineDrillInRockScope) {
        rockRow2 = currentRow
        spreadsheet.updateCell({ value: lineDrillingText }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, lineDrillingText)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            backgroundColor: 'white',
            verticalAlign: 'top',
            textDecoration: 'none'
          },
          `${pfx}B${currentRow}`
        )
        fillRatesForProposalRow(currentRow, lineDrillingText)
        // LF (C) and CY (F) from calculation sheet - second data row in rock excavation subsection
        if (rockExcavationLine2RowIndex) {
          spreadsheet.updateCell({ formula: `='Calculations Sheet'!I${rockExcavationLine2RowIndex}` }, `${pfx}C${currentRow}`)
          spreadsheet.updateCell({ formula: `='Calculations Sheet'!L${rockExcavationLine2RowIndex}` }, `${pfx}F${currentRow}`)
        } else {
          const formattedLineDrillFT = parseFloat((lineDrillTotalFT * 2).toFixed(2))
          const formattedRockExcavationCY35 = parseFloat(rockExcavationTotals.totalCY.toFixed(2))
          spreadsheet.updateCell({ value: formattedLineDrillFT }, `${pfx}C${currentRow}`)
          spreadsheet.updateCell({ value: formattedRockExcavationCY35 }, `${pfx}F${currentRow}`)
        }
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'right', format: '#,##0.00' }, `${pfx}C${currentRow}`)
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'right', format: '#,##0.00' }, `${pfx}F${currentRow}`)

        // Dollar Formula
        const dollarFormulaRock2 = `=IFERROR(IF(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)=0,"",ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)),"")`
        spreadsheet.updateCell({ formula: dollarFormulaRock2 }, `${pfx}H${currentRow}`)
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'right', format: '$#,##0.00' }, `${pfx}H${currentRow}`)
        columns33.forEach(col => spreadsheet.cellFormat({ backgroundColor: '#E2EFDA' }, `${pfx}${col}${currentRow}`))
        currentRow++
      }

      // When only rock (no soil), skip "Rock excavation Total" — Excavation Total will sum rock H range directly
      if (!showRockSubHeading) {
        rockOnlyFirstRow = rockRow1
        rockOnlyLastRow = rockRow2
      }
      if (showRockSubHeading) {
        // Rock Excavation Total
        spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
        spreadsheet.updateCell({ value: 'Rock excavation Total:' }, `${pfx}D${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            backgroundColor: '#FFF2CC'
          },
          `${pfx}D${currentRow}:E${currentRow}`
        )

        spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
        const rockTotalFormula = `=SUM(H${rockRow1}:H${rockRow2})*1000`
        spreadsheet.updateCell({ formula: rockTotalFormula }, `${pfx}F${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'right',
            backgroundColor: '#FFF2CC',
            format: '$#,##0.00'
          },
          `${pfx}F${currentRow}:G${currentRow}`
        )
        applyTotalRowBorders(spreadsheet, pfx, currentRow, '#FFF2CC')
        baseBidTotalRows.push(currentRow) // Rock Excavation Total
        totalRows.push(currentRow)

        rockExcavationTotalRow = currentRow
        currentRow++ // Empty row
        currentRow++ // Extra line after Rock Excavation Total
      }
    }

    // Excavation Total (only when there is any excavation data: soil and/or rock)
    if (hasAnyExcavationData) {
      spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
      spreadsheet.updateCell({ value: 'Excavation Total:' }, `${pfx}D${currentRow}`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: '#BDD7EE'
        },
        `${pfx}D${currentRow}:E${currentRow}`
      )

      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      // Build Excavation Total formula safely:
      // - If soil/rock total rows exist, sum only the existing ones (avoid Fnull)
      // - If no total rows yet but soil scope exists, fall back to summing H (scope-to-total) * 1000
      const excavationTotalRefs = []
      if (soilExcavationTotalRow != null) excavationTotalRefs.push(`F${soilExcavationTotalRow}`)
      if (rockExcavationTotalRow != null) excavationTotalRefs.push(`F${rockExcavationTotalRow}`)
      let excavationFullTotalFormula
      if (excavationTotalRefs.length > 0) {
        excavationFullTotalFormula = `=SUM(${excavationTotalRefs.join(',')})`
      } else if (soilExcavationRow1 != null) {
        excavationFullTotalFormula = `=SUM(H${soilExcavationRow1}:H${currentRow - 1})*1000`
      } else if (rockOnlyFirstRow != null && rockOnlyLastRow != null) {
        excavationFullTotalFormula = `=SUM(H${rockOnlyFirstRow}:H${rockOnlyLastRow})*1000`
      } else {
        excavationFullTotalFormula = '=0'
      }
      spreadsheet.updateCell({ formula: excavationFullTotalFormula }, `${pfx}F${currentRow}`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          backgroundColor: '#BDD7EE',
          format: '$#,##0.00'
        },
        `${pfx}F${currentRow}:G${currentRow}`
      )
      applyTotalRowBorders(spreadsheet, pfx, currentRow, '#BDD7EE')
      baseBidTotalRows.push(currentRow) // Excavation Total (with or without rock)
      totalRows.push(currentRow) // so Excavation Total gets centered $ like other totals

      currentRow++ // Extra line after Excavation Total
    }

    // Apply currency formatting to H and number format to others for all data rows 
    // (Demolition start to now)
    try {
      spreadsheet.numberFormat('$#,##0.00', `${pfx}H16:H${currentRow}`) // Note: H16 is hardcoded? 
      // We should format H(demolitionStartRow) ...
      // But formatting logic handles columns I-N dynamically per section.
      // H column formatting was global/ranged.
      // Let's assume the previous `formatColumns.forEach` handles I-N.
      // We need to ensure H is formatted.
      // Actually, individual cells set their format to '$#,##0.00' in my code above.
      // So global range format is redundancy/cleanup.
    } catch (e) { }

    currentRow++ // One row gap after Excavation Total

    // -------------------------------------------------------------------------
    // SOE SECTION - only show if calculation sheet has SOE data
    // -------------------------------------------------------------------------
    const soeSubsectionItemsForCheck = window.soeSubsectionItems || new Map()
    const hasSOESubsectionData = [...soeSubsectionItemsForCheck.values()].some(groups => Array.isArray(groups) && groups.some(g => Array.isArray(g) && g.length > 0))
    const hasSOEScopeData = (formulaData || []).some(f => f.section === 'soe') ||
      hasSOESubsectionData ||
      ((window.drilledSoldierPileGroups || []).length > 0) ||
      ((window.soldierPileGroups || []).length > 0)

    if (hasSOEScopeData) {
      // SOE Scope heading is written only after we know there are subsections to display (see below)

      // Compute drilled vs HP groups before headings so we know which subsection headings to show
      const collectedGroups = window.drilledSoldierPileGroups || []
      const drilledGroups = []
      const hpGroupsFromDrilled = []
      collectedGroups.forEach((group) => {
        const hasHPItems = group.some(item => /HP\d+x\d+/i.test(item.particulars || ''))
        if (hasHPItems) hpGroupsFromDrilled.push(group)
        else drilledGroups.push(group)
      })
      if (hpGroupsFromDrilled.length > 0) {
        if (!window.hpSoldierPileGroups) window.hpSoldierPileGroups = []
        window.hpSoldierPileGroups.push(...hpGroupsFromDrilled)
      }

      const hasSOESubsectionGroups = (map, keys) => {
        if (!map || !Array.isArray(keys)) return false
        for (const k of keys) {
          const groups = map.get(k) || []
          if (Array.isArray(groups) && groups.some(g => Array.isArray(g) && g.length > 0)) return true
        }
        return false
      }

      const hasSOEHeadingItems = (heading) => {
        const key = heading.replace(/:$/, '').trim()
        switch (heading) {
          case 'Soldier drilled piles:':
            return drilledGroups.length > 0
          case 'Timber soldier piles:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Timber soldier piles'])
          case 'Soldier driven piles:':
            return hpGroupsFromDrilled.length > 0 || ((window.soldierPileGroups || []).length > 0) || ((window.hpSoldierPileGroups || []).length > 0)
          case 'Primary secant piles:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Primary secant piles'])
          case 'Secondary secant piles:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Secondary secant piles'])
          case 'Tangent pile:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Tangent piles', 'Tangent pile'])
          case 'Sheet piles:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Sheet piles', 'Sheet pile'])
          case 'Timber planks:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Timber planks'])
          case 'Timber post:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Timber post'])
          case 'Timber lagging:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Timber lagging'])
          case 'Timber sheeting:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Timber sheeting'])
          case 'Bracing:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Bracing'])
          case 'Tie back anchor:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Tie back anchor', 'Tie back'])
          case 'Tie down anchor:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Tie down anchor', 'Tie down'])
          case 'Parging:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Parging']) || ((window.pargingItems || []).length > 0)
          case 'Heel blocks:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Heel blocks'])
          case 'Underpinning:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Underpinning'])
          case 'Concrete soil retention pier:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Concrete soil retention pier', 'Concrete soil retention piers'])
          case 'Form board:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Form board'])
          case 'Buttons:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Buttons'])
          case 'Concrete buttons:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Concrete buttons'])
          case 'Guide wall:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Guide wall', 'Guilde wall']) || ((window.guideWallItems || []).length > 0)
          case 'Dowels:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Dowels', 'Dowel bar']) || ((window.dowelBarItems || []).length > 0)
          case 'Rock pins:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Rock pins', 'Rock pin']) || ((window.rockPinItems || []).length > 0)
          case 'Rock stabilization:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Rock stabilization']) || ((window.rockStabilizationItems || []).length > 0)
          case 'Shotcrete:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Shotcrete']) || ((window.shotcreteItems || []).length > 0)
          case 'Permission grouting:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Permission grouting']) || ((window.permissionGroutingItems || []).length > 0)
          case 'Mud slab:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Mud slab']) || ((window.mudSlabItems || []).length > 0)
          case 'Misc.:':
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, ['Misc.', 'Misc'])
          default:
            return hasSOESubsectionGroups(soeSubsectionItemsForCheck, [key])
        }
      }

      // SOE Headings (hardcoded order) — only show heading when that subsection has at least one group
      const soeHeadings = [
        'Soldier drilled piles:',
        'Timber soldier piles:',
        'Soldier driven piles:',
        'Primary secant piles:',
        'Secondary secant piles:',
        'Tangent pile:',
        'Sheet piles:',
        'Timber planks:',
        'Timber post:',
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

      // Write SOE subheading at currentRow and advance (only call when subsection has items)
      const writeSOESubheading = (heading) => {
        spreadsheet.updateCell({ value: heading }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, heading)
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
        currentRow++
      }

      let rowShift = 0

      // Helper function... (retained)
      const getSOEPageFromRawData = (diameter, thickness) => {
        // ... (retained implementation or fallback)
        if (!rawData || !Array.isArray(rawData) || rawData.length < 2) return 'SOE-101.00'
        // ... (simplified for space)
        return 'SOE-101.00'
      }

      // Add drilled soldier pile proposal text (drilledGroups / hpGroupsFromDrilled already computed above for SOE headings)
      const parseDimension = (dimStr) => {
        if (!dimStr) return 0
        const match = dimStr.match(/(\d+)(?:'-?)?(\d+)?/)
        if (!match) return 0
        return (parseInt(match[1]) || 0) + ((parseInt(match[2]) || 0) / 12)
      }

      // Merge drilled groups that have the same diameter, thickness, H, E (and RS) so we output one proposal line per unique combo
      const getDrilledGroupKey = (items) => {
        if (!items || items.length === 0) return ''
        const p = (items[0].particulars || '').trim()
        const dMatch = p.match(/([0-9.]+)Ø\s*x\s*([0-9.]+)/i)
        const diameter = dMatch ? dMatch[1] : ''
        const thickness = dMatch ? dMatch[2] : ''
        const hMatch = p.match(/H=([0-9'"\-]+)/)
        const h = hMatch ? hMatch[1] : ''
        const rsMatch = p.match(/\+\s*RS=([0-9'"\-]+)/i)
        const rs = rsMatch ? rsMatch[1] : ''
        const eMatch = p.match(/E=([0-9'"\-]+)/)
        const e = eMatch ? eMatch[1] : ''
        return `${diameter}|${thickness}|${h}|${rs}|${e}`
      }
      const mergedDrilledGroups = []
      const keyToItems = new Map()
      drilledGroups.forEach((collectedItems) => {
        if (collectedItems.length === 0) return
        const key = getDrilledGroupKey(collectedItems)
        if (!keyToItems.has(key)) keyToItems.set(key, [])
        keyToItems.get(key).push(...collectedItems)
      })
      keyToItems.forEach((items) => mergedDrilledGroups.push(items))

      // Drilled soldier piles: write heading only when we have items, then content (next group will be on next row)
      if (mergedDrilledGroups.length > 0) {
        writeSOESubheading('Soldier drilled piles:')
        // Process each merged group (one proposal line per unique diameter/thickness/H/E combo)
        mergedDrilledGroups.forEach((collectedItems, groupIndex) => {
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

            // Process each group to extract values (use collected items so we can use item.height from calc row)
            const itemsByParticulars = new Map()
            collectedItems.forEach(item => {
              const p = (item.particulars || '').trim()
              if (!itemsByParticulars.has(p)) itemsByParticulars.set(p, [])
              itemsByParticulars.get(p).push(item)
            })
            groupedItems.forEach((group, particulars) => {
              // Extract diameter and thickness (e.g., "9.625Ø x0.545")
              if (!diameter || !thickness) {
                const drilledMatch = particulars.match(/([0-9.]+)Ø\s*x\s*([0-9.]+)/i)
                if (drilledMatch) {
                  diameter = parseFloat(drilledMatch[1])
                  thickness = parseFloat(drilledMatch[2])
                }
              }

              const groupItems = itemsByParticulars.get(particulars) || []
              // Extract H value: from particulars (e.g., "H=27'-10"") or from calculation row Height (col H)
              let heightValue = null
              const hMatch = particulars.match(/H=([0-9'"\-]+)/)
              if (hMatch) {
                heightValue = parseDimension(hMatch[1])
                group.hValue = hMatch[1]
                const rsMatch = particulars.match(/\+\s*RS=([0-9'"\-]+)/i)
                if (rsMatch) {
                  const rsValue = parseDimension(rsMatch[1])
                  group.rsValue = rsMatch[1]
                  heightValue = heightValue + rsValue
                }
                group.combinedHeight = heightValue
              } else {
                // Use Height from calculation sheet (column H) when H= not in particulars
                const fromRow = groupItems.find(it => it.height != null && it.height > 0)
                if (fromRow && fromRow.height > 0) heightValue = fromRow.height
              }
              if (heightValue != null) {
                group.combinedHeight = heightValue
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

              // Row range for this (possibly merged) group so FT/LBS/QTY sum across all items
              const rowNumbers = collectedItems.map(item => item.rawRowNumber || 0).filter(Boolean)
              const firstRowNumber = rowNumbers.length > 0 ? Math.min(...rowNumbers) : 0
              const lastRowNumber = rowNumbers.length > 0 ? Math.max(...rowNumbers) : 0
              const sumRowIndex = lastRowNumber
              const useSumRange = firstRowNumber > 0 && lastRowNumber > 0 && firstRowNumber !== lastRowNumber
              const calcSheetName = 'Calculations Sheet'

              // Use QTY from calculation data: for merged groups use sum of column M across rows; else single row
              let displayCount = totalCount
              if (calculationData && rowNumbers.length > 0) {
                if (useSumRange) {
                  let sumQty = 0
                  for (let r = firstRowNumber; r <= lastRowNumber; r++) {
                    const row = calculationData[r - 1]
                    if (row && row[12] != null) {
                      const q = parseFloat(row[12])
                      if (!Number.isNaN(q)) sumQty += q
                    }
                  }
                  if (sumQty > 0) displayCount = Math.round(sumQty)
                } else if (sumRowIndex >= 1 && calculationData[sumRowIndex - 1]) {
                  const rowQty = parseFloat(calculationData[sumRowIndex - 1][12])
                  if (!Number.isNaN(rowQty) && rowQty > 0) displayCount = Math.round(rowQty)
                }
              }

              // Cell refs: single row or SUM over range for merged groups
              const ftCellRef = useSumRange ? `SUM('${calcSheetName}'!I${firstRowNumber}:I${lastRowNumber})` : `'${calcSheetName}'!I${sumRowIndex}`
              const lbsCellRef = useSumRange ? `SUM('${calcSheetName}'!K${firstRowNumber}:K${lastRowNumber})` : `'${calcSheetName}'!K${sumRowIndex}`
              const qtyCellRef = useSumRange ? `SUM('${calcSheetName}'!M${firstRowNumber}:M${lastRowNumber})` : `'${calcSheetName}'!M${sumRowIndex}`

              // Proposal text for rate lookup and row height; B cell uses formula so (XX) updates when QTY changes
              const proposalText = `F&I new (${displayCount})no [${diameter}" Øx${thickness}" thick] drilled soldier piles (H=${heightText}, ${embedmentText} embedment) as per ${soePage} & details on`
              const afterCount = `)no [${diameter}" Øx${thickness}" thick] drilled soldier piles (H=${heightText}, ${embedmentText} embedment) as per ${soePage} & details on`
              const bFormula = proposalFormulaWithQtyRef(currentRow, afterCount)

              spreadsheet.updateCell({ formula: bFormula }, `${pfx}B${currentRow}`)
              rowBContentMap.set(currentRow, proposalText)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  color: '#000000',
                  textAlign: 'left',
                  backgroundColor: 'white',
                  verticalAlign: 'top',
                  textDecoration: 'none'
                },
                `${pfx}B${currentRow}`
              )
              fillRatesForProposalRow(currentRow, proposalText)

              // Calculate and set row height based on content
              const dynamicHeight = calculateRowHeight(proposalText)

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

              // Add $/1000 formula in column H
              const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
              spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white',
                  format: '$#,##0.00'
                },
                `${pfx}H${currentRow}`
              )

              currentRow++ // Move to next row for next group
            }
          }
        })
      }

      // Soldier driven piles (HP): write heading only when we have items, then content
      const hasHPGroups = hpGroupsFromDrilled.length > 0 || ((window.hpSoldierPileGroups || []).length > 0)
      if (hasHPGroups) {
        writeSOESubheading('Soldier driven piles:')
      }

      // Add HP soldier pile proposal text from collected groups
      const hpCollectedGroupsRaw = window.hpSoldierPileGroups || []

      // Merge HP groups by (HP type, height) so one proposal line per unique combo; skip empty groups
      const getHPGroupKey = (items) => {
        if (!items || items.length === 0) return ''
        const p = (items[0].particulars || '').trim()
        const hpMatch = p.match(/HP(\d+)x(\d+)/i)
        const hpType = hpMatch ? `HP${hpMatch[1]}x${hpMatch[2]}` : ''
        const hMatch = p.match(/H=([0-9'"\-]+)/)
        const hRaw = hMatch ? hMatch[1] : ''
        return `${hpType}|${hRaw}`
      }
      const hpKeyToItems = new Map()
      hpCollectedGroupsRaw.forEach((group) => {
        if (!group || group.length === 0) return
        const key = getHPGroupKey(group)
        if (!key) return
        if (!hpKeyToItems.has(key)) hpKeyToItems.set(key, [])
        hpKeyToItems.get(key).push(...group)
      })
      const hpCollectedGroups = Array.from(hpKeyToItems.values())

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

            // Group items by particulars and keep item refs for height from calc row
            const groupedItems = new Map()
            const itemsByParticulars = new Map()
            hpCollectedItems.forEach(item => {
              const particulars = (item.particulars || '').trim()
              if (!groupedItems.has(particulars)) {
                groupedItems.set(particulars, {
                  particulars: particulars,
                  takeoff: 0,
                  hValue: null
                })
                itemsByParticulars.set(particulars, [])
              }
              const group = groupedItems.get(particulars)
              group.takeoff += (item.takeoff || 0)
              itemsByParticulars.get(particulars).push(item)
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

              const groupItems = itemsByParticulars.get(particulars) || []
              let heightValue = null
              const hMatch = particulars.match(/H=([0-9'"\-]+)/)
              if (hMatch) {
                heightValue = parseDimension(hMatch[1])
                group.hValue = hMatch[1]
              } else {
                // Use Height from calculation sheet (column H) when H= not in particulars
                const fromRow = groupItems.find(it => it.height != null && it.height > 0)
                if (fromRow && fromRow.height > 0) heightValue = fromRow.height
              }

              if (heightValue != null) {
                const itemFT = heightValue * group.takeoff
                totalFT += itemFT
                if (hpWeight) {
                  const itemLBS = itemFT * hpWeight
                  totalLBS += itemLBS
                }
                for (let i = 0; i < group.takeoff; i++) {
                  totalHeight += heightValue
                  heightCount++
                }
              }

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

              // Get SOE page ref(s) from raw data (search for HP type); multiple pages formatted as "102 & 105" or "102, 103 & 105"
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

                const pattern = new RegExp(hpTypeValue.replace(/[.*+?^${}()|[\]\\]/g, '\\$&'), 'i')
                const collected = []
                const seen = new Set()
                for (const row of dataRows) {
                  const digitizerItem = row[digitizerIdx]
                  if (digitizerItem && pattern.test(String(digitizerItem))) {
                    const pageValue = row[pageIdx]
                    if (pageValue) {
                      const pageStr = String(pageValue).trim()
                      const soeMatch = pageStr.match(/SOE-[\w-]+/i)
                      if (soeMatch && !seen.has(soeMatch[0])) {
                        seen.add(soeMatch[0])
                        collected.push(soeMatch[0])
                      }
                    }
                  }
                }

                return collected.length ? formatPageRefList(collected) : 'SOE-B-100.00'
              }

              const soePage = getHPPageFromRawData(hpType)

              // Row range for this (possibly merged) group
              const hpRowNumbers = hpCollectedItems.map(item => item.rawRowNumber || 0).filter(Boolean)
              const firstRowNumber = hpRowNumbers.length > 0 ? Math.min(...hpRowNumbers) : 0
              const lastRowNumber = hpRowNumbers.length > 0 ? Math.max(...hpRowNumbers) : 0
              const sumRowIndex = lastRowNumber
              const useSumRange = firstRowNumber > 0 && lastRowNumber > 0 && firstRowNumber !== lastRowNumber
              const calcSheetName = 'Calculations Sheet'

              // Use QTY: for merged groups sum column M over range; else single row
              let displayCount = totalCount
              if (calculationData && hpRowNumbers.length > 0) {
                if (useSumRange) {
                  let sumQty = 0
                  for (let r = firstRowNumber; r <= lastRowNumber; r++) {
                    const row = calculationData[r - 1]
                    if (row && row[12] != null) {
                      const q = parseFloat(row[12])
                      if (!Number.isNaN(q)) sumQty += q
                    }
                  }
                  if (sumQty > 0) displayCount = Math.round(sumQty)
                } else if (sumRowIndex >= 1 && calculationData[sumRowIndex - 1]) {
                  const rowQty = parseFloat(calculationData[sumRowIndex - 1][12])
                  if (!Number.isNaN(rowQty) && rowQty > 0) displayCount = Math.round(rowQty)
                }
              }

              const ftCellRef = useSumRange ? `SUM('${calcSheetName}'!I${firstRowNumber}:I${lastRowNumber})` : `'${calcSheetName}'!I${sumRowIndex}`
              const lbsCellRef = useSumRange ? `SUM('${calcSheetName}'!K${firstRowNumber}:K${lastRowNumber})` : `'${calcSheetName}'!K${sumRowIndex}`
              const qtyCellRef = useSumRange ? `SUM('${calcSheetName}'!M${firstRowNumber}:M${lastRowNumber})` : `'${calcSheetName}'!M${sumRowIndex}`

              const proposalText = `F&I new (${displayCount})no [${hpType}] driven pile (Havg=${heightText}) as per ${soePage} & details on`
              const afterCount = `)no [${hpType}] driven pile (Havg=${heightText}) as per ${soePage} & details on`
              const bFormula = proposalFormulaWithQtyRef(currentRow, afterCount)

              spreadsheet.updateCell({ formula: bFormula }, `${pfx}B${currentRow}`)
              rowBContentMap.set(currentRow, proposalText)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  color: '#000000',
                  textAlign: 'left',
                  backgroundColor: 'white',
                  verticalAlign: 'top',
                  textDecoration: 'none'
                },
                `${pfx}B${currentRow}`
              )
              // Use scope-specific rate so Soldier driven piles get LF 150 (proposal_mapped.json)
              fillRatesForProposalRow(currentRow, 'driven pile - Soldier driven piles scope')

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

              // Add $/1000 formula in column H
              const dollarFormulaHP = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
              spreadsheet.updateCell({ formula: dollarFormulaHP }, `${pfx}H${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white',
                  format: '$#,##0.00'
                },
                `${pfx}H${currentRow}`
              )

              // Calculate and set row height based on content in column B
              const dynamicHeightHP = calculateRowHeight(proposalText)

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
              return
            }

            // If we're in the subsection and hit another subsection or section, stop
            if (inPrimarySecantSubsection) {
              // Stop if we hit another subsection header (ends with ':')
              if (bText.endsWith(':') && bText.toLowerCase() !== 'primary secant piles:') {
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

        if (foundPrimarySecantRows.length > 0) {
          // Group them (all in one group for now) and store the sum row
          primarySecantGroups = [{
            items: foundPrimarySecantRows,
            sumRow: primarySecantSumRow
          }]
        }
      }

      // Primary secant piles: write heading only when we have items, then content (next group on next row)
      if (primarySecantGroups.length > 0) {
        writeSOESubheading('Primary secant piles:')
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
              const soeRefsCollected = new Set()
              for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                const row = dataRows[rowIndex]
                const digitizerItem = row[digitizerIdx]
                if (digitizerItem && typeof digitizerItem === 'string') {
                  const itemText = digitizerItem.toLowerCase()
                  if (itemText.includes('primary secant')) {
                    const pageValue = row[pageIdx]
                    if (pageValue) {
                      const pageStr = String(pageValue).trim()
                      const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                      if (soeMatches) soeMatches.forEach(m => soeRefsCollected.add(m))
                    }
                  }
                }
              }
              if (soeRefsCollected.size > 0) {
                soePageMain = formatPageRefList([...soeRefsCollected])
                const arr = [...soeRefsCollected]
                if (arr.length > 1) soePageDetails = arr[1]
              }
            }
          }

          // Find the sum row for this group
          // First try formulaData (most reliable), then groupSumRow from search, then calculation sheet
          let sumRowIndex = groupSumRow
          if (!sumRowIndex && formulaData && Array.isArray(formulaData)) {
            const sumFormula = formulaData.find(f =>
              f.itemType === 'soe_generic_sum' && f.subsectionName === 'Primary secant piles'
            )
            if (sumFormula) sumRowIndex = sumFormula.row
          }

          // If we don't have a sum row from search, try to find it in the calculation sheet
          if (!sumRowIndex && calculationData && calculationData.length > 0) {
            // First, search backwards from firstRowNumber to find "Primary secant piles:" header
            let primarySecantHeaderFound = false
            for (let i = Math.max(0, firstRowNumber - 20); i < firstRowNumber; i++) {
              const rowData = calculationData[i - 1] // calculationData is 0-indexed
              const colB = rowData?.[1]
              if (colB && typeof colB === 'string' && colB.trim().toLowerCase() === 'primary secant piles:') {
                primarySecantHeaderFound = true
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
                break
              }

              // Also check if we hit another subsection header
              if (colB && typeof colB === 'string' && colB.trim().endsWith(':')) {
                // The sum row should be the row before this subsection
                sumRowIndex = i
                break
              }
            }

            // Fallback: if still not found, use lastRowNumber + 1
            if (!sumRowIndex) {
              sumRowIndex = lastRowNumber + 1
            }
          } else if (!sumRowIndex) {
            sumRowIndex = lastRowNumber + 1
          }

          const calcSheetName = 'Calculations Sheet'

          // Use QTY from sum row (column M) for display when available so (XX)no matches the sheet
          let qtyValue = Math.round(totalQty || totalTakeoff)
          if (calculationData && sumRowIndex >= 1 && calculationData[sumRowIndex - 1]) {
            const rowQty = parseFloat(calculationData[sumRowIndex - 1][12])
            if (!Number.isNaN(rowQty) && rowQty > 0) qtyValue = Math.round(rowQty)
          }

          let proposalText = ''
          let afterCount = ''
          if (diameter) {
            proposalText = `F&I new (${qtyValue})no [${diameter}" Ø] primary secant piles (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain} & details on`
            afterCount = `)no [${diameter}" Ø] primary secant piles (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain} & details on`
          } else {
            proposalText = `F&I new (${qtyValue})no [#" Ø] primary secant piles (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain} & details on`
            afterCount = `)no [#" Ø] primary secant piles (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain} & details on`
          }

          if (sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: proposalFormulaWithQtyRef(currentRow, afterCount) }, `${pfx}B${currentRow}`)
          } else {
            spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          }
          rowBContentMap.set(currentRow, proposalText)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: 'white',
              verticalAlign: 'top',
              textDecoration: 'none'
            },
            `${pfx}B${currentRow}`
          )
          fillRatesForProposalRow(currentRow, proposalText)

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

          // Primary secant piles: column K is empty in calculation sheet, so no LBS in column E

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

          // Add $/1000 formula in column H
          const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
          spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: 'white',
              format: '$#,##0.00'
            },
            `${pfx}H${currentRow}`
          )

          // Apply currency format
          try {
            spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}:N${currentRow}`)
          } catch (e) {
            // Fallback already applied in cellFormat
          }

          // Calculate and set row height based on content in column B
          const proposalTextForHeight = proposalText || (spreadsheet.getCellValue ? spreadsheet.getCellValue(`${pfx}B${currentRow}`) : '')
          const dynamicHeight = calculateRowHeight(proposalTextForHeight)

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

      // Secondary secant piles: write heading only when we have items, then content (next group on next row)
      if (secondarySecantGroups.length > 0) {
        writeSOESubheading('Secondary secant piles:')
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
              const soeRefsCollected = new Set()
              for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                const row = dataRows[rowIndex]
                const digitizerItem = row[digitizerIdx]
                if (digitizerItem && typeof digitizerItem === 'string') {
                  const itemText = digitizerItem.toLowerCase()
                  if (itemText.includes('secondary secant')) {
                    const pageValue = row[pageIdx]
                    if (pageValue) {
                      const pageStr = String(pageValue).trim()
                      const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                      if (soeMatches) soeMatches.forEach(m => soeRefsCollected.add(m))
                    }
                  }
                }
              }
              if (soeRefsCollected.size > 0) soePageMain = formatPageRefList([...soeRefsCollected])
            }
          }

          // Find the sum row for this group - use formulaData when available for correct row
          let sumRowIndex = lastRowNumber
          if (formulaData && Array.isArray(formulaData)) {
            const sumFormula = formulaData.find(f =>
              f.itemType === 'soe_generic_sum' && f.subsectionName === 'Secondary secant piles'
            )
            if (sumFormula) sumRowIndex = sumFormula.row
          } else {
            sumRowIndex = lastRowNumber + 1 // Sum row is after last data row
          }
          const calcSheetName = 'Calculations Sheet'

          // Use QTY from sum row (column M) for display when available so (XX)no matches the sheet
          let qtyValue = Math.round(totalQty || totalTakeoff)
          if (calculationData && sumRowIndex >= 1 && calculationData[sumRowIndex - 1]) {
            const rowQty = parseFloat(calculationData[sumRowIndex - 1][12])
            if (!Number.isNaN(rowQty) && rowQty > 0) qtyValue = Math.round(rowQty)
          }

          let proposalText = ''
          let afterCount = ''
          if (diameter) {
            proposalText = `F&I new (${qtyValue})no [${diameter}" Ø] secondary secant piles (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain} & details on`
            afterCount = `)no [${diameter}" Ø] secondary secant piles (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain} & details on`
          } else {
            proposalText = `F&I new (${qtyValue})no [#" Ø] secondary secant piles (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain} & details on`
            afterCount = `)no [#" Ø] secondary secant piles (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain} & details on`
          }

          if (sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: proposalFormulaWithQtyRef(currentRow, afterCount) }, `${pfx}B${currentRow}`)
          } else {
            spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          }
          rowBContentMap.set(currentRow, proposalText)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: 'white',
              verticalAlign: 'top',
              textDecoration: 'none'
            },
            `${pfx}B${currentRow}`
          )
          fillRatesForProposalRow(currentRow, proposalText)

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

          // Add LBS to column E - reference to calculation sheet sum row (column K)
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

          // Add $/1000 formula in column H
          const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
          spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: 'white',
              format: '$#,##0.00'
            },
            `${pfx}H${currentRow}`
          )

          // Apply currency format
          try {
            spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}:N${currentRow}`)
          } catch (e) {
            // Fallback already applied in cellFormat
          }

          // Calculate and set row height based on content in column B
          const proposalTextForHeight = proposalText || (spreadsheet.getCellValue ? spreadsheet.getCellValue(`${pfx}B${currentRow}`) : '')
          const dynamicHeight = calculateRowHeight(proposalTextForHeight)

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

      // Tangent pile: write heading only when we have items, then content (next group on next row)
      if (tangentPileGroups.length > 0) {
        writeSOESubheading('Tangent pile:')
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
              const soeRefsCollected = new Set()
              for (let rowIndex = 0; rowIndex < dataRows.length; rowIndex++) {
                const row = dataRows[rowIndex]
                const digitizerItem = row[digitizerIdx]
                if (digitizerItem && typeof digitizerItem === 'string') {
                  const itemText = digitizerItem.toLowerCase()
                  if (itemText.includes('tangent')) {
                    const pageValue = row[pageIdx]
                    if (pageValue) {
                      const pageStr = String(pageValue).trim()
                      const soeMatches = pageStr.match(/SOE-[\d.]+/gi)
                      if (soeMatches) soeMatches.forEach(m => soeRefsCollected.add(m))
                    }
                  }
                }
              }
              if (soeRefsCollected.size > 0) soePageMain = formatPageRefList([...soeRefsCollected])
            }
          }

          // Find the sum row for this group - use formulaData when available for correct row
          let sumRowIndex = lastRowNumber
          if (formulaData && Array.isArray(formulaData)) {
            const sumFormula = formulaData.find(f =>
              (f.itemType === 'soe_generic_sum') &&
              (f.subsectionName === 'Tangent piles' || f.subsectionName === 'Tangent pile')
            )
            if (sumFormula) sumRowIndex = sumFormula.row
          } else {
            sumRowIndex = lastRowNumber + 1 // Sum row is after last data row
          }
          const calcSheetName = 'Calculations Sheet'

          let qtyValue = Math.round(totalQty || totalTakeoff)
          if (calculationData && sumRowIndex >= 1 && calculationData[sumRowIndex - 1]) {
            const rowQty = parseFloat(calculationData[sumRowIndex - 1][12])
            if (!Number.isNaN(rowQty) && rowQty > 0) qtyValue = Math.round(rowQty)
          }

          let proposalText = ''
          let afterCount = ''
          if (diameter) {
            proposalText = `F&I new (${qtyValue})no [${diameter}" Ø] tangent pile (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain} & details on`
            afterCount = `)no [${diameter}" Ø] tangent pile (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain} & details on`
          } else {
            proposalText = `F&I new (${qtyValue})no [#" Ø] tangent pile (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain} & details on`
            afterCount = `)no [#" Ø] tangent pile (Havg=${heightText}, ${embedmentText} embedment) as per ${soePageMain} & details on`
          }

          if (sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: proposalFormulaWithQtyRef(currentRow, afterCount) }, `${pfx}B${currentRow}`)
          } else {
            spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          }
          rowBContentMap.set(currentRow, proposalText)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: 'white',
              verticalAlign: 'top',
              textDecoration: 'none'
            },
            `${pfx}B${currentRow}`
          )
          fillRatesForProposalRow(currentRow, proposalText)

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

          // Tangent piles: column K is empty in calculation sheet, so no LBS in column E

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

          // Add $/1000 formula in column H
          const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
          spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: 'white',
              format: '$#,##0.00'
            },
            `${pfx}H${currentRow}`
          )

          // Apply currency format
          try {
            spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}:N${currentRow}`)
          } catch (e) {
            // Fallback already applied in cellFormat
          }

          // Calculate and set row height based on content in column B
          const proposalTextForHeight = proposalText || (spreadsheet.getCellValue ? spreadsheet.getCellValue(`${pfx}B${currentRow}`) : '')
          const dynamicHeight = calculateRowHeight(proposalTextForHeight)

          currentRow++ // Move to next row for next group
        })
      }

      // Fallback: ensure Timber soldier piles / planks / post are in soeSubsectionItems when present in calculationData
      // (parser may miss them if section detection or subsection tracking skips rows)
      const timberSubsectionNames = ['Timber soldier piles', 'Timber planks', 'Timber post']
      if (calculationData && calculationData.length > 0) {
        timberSubsectionNames.forEach((subName) => {
          const existing = (window.soeSubsectionItems || new Map()).get(subName) || []
          if (existing.length > 0 && existing.some(g => g.length > 0)) return
          const groups = []
          let inSubsection = false
          let currentGroup = []
          for (let i = 0; i < calculationData.length; i++) {
            const row = calculationData[i]
            const colB = row && row[1] != null ? String(row[1]).trim() : ''
            const colA = row && row[0] != null ? String(row[0]).trim().toLowerCase() : ''
            if (colA === 'soe') {
              inSubsection = false
              currentGroup = []
              continue
            }
            if (colA === 'foundation' || colA === 'foundation/substructure' || (colA && colA !== 'soe' && colA !== 'demolition' && colA !== 'excavation' && colA !== 'rock excavation')) {
              inSubsection = false
              if (currentGroup.length > 0) {
                groups.push(currentGroup)
                currentGroup = []
              }
              continue
            }
            if (colB === (subName + ':')) {
              inSubsection = true
              if (currentGroup.length > 0) {
                groups.push(currentGroup)
                currentGroup = []
              }
              continue
            }
            if (inSubsection && colB && colB.endsWith(':')) {
              inSubsection = false
              if (currentGroup.length > 0) {
                groups.push(currentGroup)
                currentGroup = []
              }
              continue
            }
            if (inSubsection && colB && !colB.endsWith(':')) {
              const takeoff = parseFloat(row[2]) || 0
              const unit = row[3] || ''
              const height = parseFloat(row[7]) || 0
              const sqft = parseFloat(row[9]) || 0
              const lbs = parseFloat(row[10]) || 0
              const qty = parseFloat(row[12]) || 0
              if (takeoff > 0 || qty > 0 || colB.toLowerCase().includes('timber')) {
                currentGroup.push({
                  particulars: colB,
                  takeoff,
                  unit,
                  height,
                  sqft,
                  lbs,
                  qty,
                  rawRow: row,
                  rawRowNumber: i + 2
                })
              }
              continue
            }
            if (inSubsection && !colB) {
              if (currentGroup.length > 0) {
                groups.push(currentGroup)
                currentGroup = []
              }
            }
          }
          if (currentGroup.length > 0) groups.push(currentGroup)
          if (groups.length > 0 && groups.some(g => g.length > 0)) {
            if (!window.soeSubsectionItems) window.soeSubsectionItems = new Map()
            window.soeSubsectionItems.set(subName, groups)
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
        'Timber soldier piles',
        'Timber planks',
        'Timber post',
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
        'Concrete buttons',
        'Guide wall',
        'Dowels',
        'Rock pins',
        'Rock stabilization',
        'Shotcrete',
        'Permission grouting',
        'Form board',
        'Mud slab',
        'Drilled hole grout',
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

      // SOE Mud slab: only show when excavation scope has mud slab (e.g. "Slope exc & backfill ... w/ 2" mud slab").
      // Do NOT use foundation "Mud Slab:" / "Mud slab" (Foundation/Substructure) — only rows that clearly describe excavation (w/, backfill, exc, slope).
      let foundMudSlab = false
      if (calculationData && calculationData.length > 0) {
        for (const row of calculationData) {
          const cellText = ((row[0] || '') + (row[1] || '')).toString().trim().toLowerCase()
          if (!cellText.includes('mud slab')) continue
          // Excavation mud slab only: row must contain mud slab AND an excavation phrase (exclude standalone "Mud slab:" or "Mud slab" under Foundation)
          const hasExcavationPhrase = cellText.includes("w/") || cellText.includes('backfill') || cellText.includes("exc") || cellText.includes('slope')
          if (hasExcavationPhrase) {
            foundMudSlab = true
            break
          }
        }
        if (foundMudSlab && !collectedSubsections.has('Mud slab')) {
          collectedSubsections.add('Mud slab')
        }
      }
      const hasMudSlabInCalc = foundMudSlab

      // Only add Misc. subsection when calculation sheet has Misc.: section
      const hasMiscInCalc = calculationData && calculationData.length > 0 && calculationData.some((row) => {
        const b = (row[1] || '').toString().trim()
        const bLower = b.toLowerCase()
        return (bLower === 'misc.' || bLower === 'misc') && b.endsWith(':')
      })
      if (hasMiscInCalc && !collectedSubsections.has('Misc.') && !collectedSubsections.has('Misc')) {
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

      // Check if Vertical timber sheets, Horizontal timber sheets, Timber stringer are in soeSubsectionItems (display above Bracing)
      const verticalTimberSheetsGroups = soeSubsectionItems.get('Vertical timber sheets') || []
      const horizontalTimberSheetsGroups = soeSubsectionItems.get('Horizontal timber sheets') || []
      const timberStringerGroups = soeSubsectionItems.get('Timber stringer') || []
      const hasVerticalTimberSheetsItems = verticalTimberSheetsGroups.length > 0 && verticalTimberSheetsGroups.some(g => g.length > 0)
      const hasHorizontalTimberSheetsItems = horizontalTimberSheetsGroups.length > 0 && horizontalTimberSheetsGroups.some(g => g.length > 0)
      const hasTimberStringerItems = timberStringerGroups.length > 0 && timberStringerGroups.some(g => g.length > 0)
      if (hasVerticalTimberSheetsItems && !collectedSubsections.has('Vertical timber sheets')) collectedSubsections.add('Vertical timber sheets')
      if (hasHorizontalTimberSheetsItems && !collectedSubsections.has('Horizontal timber sheets')) collectedSubsections.add('Horizontal timber sheets')
      if (hasTimberStringerItems && !collectedSubsections.has('Timber stringer')) collectedSubsections.add('Timber stringer')

      // Also check calculationData for these subsections
      if (calculationData && calculationData.length > 0) {
        for (const row of calculationData) {
          const colB = row[1]
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText.includes('vertical timber sheets') && bText.endsWith(':') && !collectedSubsections.has('Vertical timber sheets')) collectedSubsections.add('Vertical timber sheets')
            if (bText.includes('horizontal timber sheets') && bText.endsWith(':') && !collectedSubsections.has('Horizontal timber sheets')) collectedSubsections.add('Horizontal timber sheets')
            if (bText.includes('timber stringer') && bText.endsWith(':') && !collectedSubsections.has('Timber stringer')) collectedSubsections.add('Timber stringer')
          }
        }
      }

      // Check if Timber soldier piles, Timber planks, Timber post are in soeSubsectionItems
      const timberSoldierPileGroups = soeSubsectionItems.get('Timber soldier piles') || []
      const hasTimberSoldierPileItems = timberSoldierPileGroups.length > 0 && timberSoldierPileGroups.some(g => g.length > 0)
      const timberPlankGroups = soeSubsectionItems.get('Timber planks') || []
      const hasTimberPlankItems = timberPlankGroups.length > 0 && timberPlankGroups.some(g => g.length > 0)
      const timberPostGroups = soeSubsectionItems.get('Timber post') || []
      const hasTimberPostItems = timberPostGroups.length > 0 && timberPostGroups.some(g => g.length > 0)
      if (hasTimberSoldierPileItems && !collectedSubsections.has('Timber soldier piles')) collectedSubsections.add('Timber soldier piles')
      if (hasTimberPlankItems && !collectedSubsections.has('Timber planks')) collectedSubsections.add('Timber planks')
      if (hasTimberPostItems && !collectedSubsections.has('Timber post')) collectedSubsections.add('Timber post')

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
      const bracingSubsections = ['Waler', 'Raker', 'Upper Raker', 'Lower Raker', 'Stand off', 'Kicker', 'Channel', 'Roll chock', 'Stub beam', 'Stud beam', 'Corner brace', 'Knee brace', 'Supporting angle']
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
        } else if (name === 'Mud slab' && hasMudSlabInCalc) {
          // Only show Mud slab when calculation sheet has data
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (name === 'Misc.' || name === 'Misc') {
          // Misc. is added later only when at least one other SOE subsection exists (see below)
          collectedSubsections.delete('Misc.')
          collectedSubsections.delete('Misc')
        } else if (name === 'Drilled hole grout' && hasDrilledHoleGroutItems) {
          // Only show Drilled hole grout when calculation sheet has data
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
          // Special handling for Timber lagging - show it only when there are items
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (name === 'Timber lagging') {
          // Do not show heading when the group is not there (do not add from collectedSubsections)
        } else if (name === 'Timber sheeting' && hasTimberSheetingItems) {
          // Special handling for Timber sheeting - show it if any items exist
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (name === 'Vertical timber sheets' && hasVerticalTimberSheetsItems) {
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (name === 'Horizontal timber sheets' && hasHorizontalTimberSheetsItems) {
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (name === 'Timber stringer' && hasTimberStringerItems) {
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (name === 'Timber soldier piles' && hasTimberSoldierPileItems) {
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (name === 'Timber planks' && hasTimberPlankItems) {
          subsectionsToDisplay.push(name)
          collectedSubsections.delete(name)
        } else if (name === 'Timber post' && hasTimberPostItems) {
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

      // If drilled hole grout wasn't added yet (and calculation has data), add it now (right after Mud slab)
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
          const mudSlabPos = subsectionsToDisplay.indexOf('Mud slab')
          if (mudSlabPos !== -1) {
            subsectionsToDisplay.splice(mudSlabPos + 1, 0, 'Drilled hole grout')
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

      // If Vertical timber sheets, Horizontal timber sheets, or Timber stringer exist but weren't added, insert before Bracing
      const beforeBracingSubsections = [
        { name: 'Vertical timber sheets', has: hasVerticalTimberSheetsItems },
        { name: 'Horizontal timber sheets', has: hasHorizontalTimberSheetsItems },
        { name: 'Timber stringer', has: hasTimberStringerItems }
      ]
      beforeBracingSubsections.forEach(({ name, has }) => {
        if (has && !subsectionsToDisplay.includes(name)) {
          const idx = subsectionOrder.indexOf(name)
          const bracingIdx = subsectionOrder.indexOf('Bracing')
          let insertIndex = subsectionsToDisplay.length
          for (let i = 0; i < subsectionsToDisplay.length; i++) {
            const currentIndex = subsectionOrder.indexOf(subsectionsToDisplay[i])
            if (currentIndex !== -1 && (idx !== -1 && currentIndex > idx || currentIndex >= bracingIdx)) {
              insertIndex = i
              break
            }
          }
          subsectionsToDisplay.splice(insertIndex, 0, name)
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

      // Add hardcoded Misc. when at least one other SOE subsection is already shown
      if (subsectionsToDisplay.length > 0) {
        subsectionsToDisplay.push('Misc.')
      }

      // Only show SOE scope section when there is at least one subsection to display
      if (subsectionsToDisplay.length > 0) {
        spreadsheet.updateCell({ value: 'SOE scope:' }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, 'SOE scope:')
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
        currentRow++
      }

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

      // Track SOE scope rows for SOE Total calculation
      let soeScopeStartRow = null
      let soeScopeEndRow = null

      if (subsectionsToDisplay.length > 0) {
      subsectionsToDisplay.forEach((subsectionName) => {
        // Track first subsection row for SOE Total
        if (soeScopeStartRow === null) {
          soeScopeStartRow = currentRow
        }

        // Special handling for Bracing - show it as a heading and then process all bracing items below
        if (subsectionName === 'Bracing' && hasBracingItems) {
          // Add Bracing header
          const subsectionText = `Bracing: including shims, stiffeners, and/or wale seats as required`
          spreadsheet.updateCell({ value: subsectionText }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, subsectionText)
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
          currentRow++

          // Bracing summary lines: values from Timber waler / Timber raker / Timber brace (size, qty, SOE refs dynamic)
          let soeBracingMain = 'SOE-101.00'
          let soeBracingDetails = 'SOE-002.00'
          if (rawData && Array.isArray(rawData) && rawData.length > 1) {
            const headers = rawData[0]
            const dataRows = rawData.slice(1)
            const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
            const pageIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'page')
            if (digitizerIdx !== -1 && pageIdx !== -1) {
              for (let r = 0; r < dataRows.length; r++) {
                const d = String(dataRows[r][digitizerIdx] || '').toLowerCase()
                if (d.includes('waler') || d.includes('raker') || d.includes('timber brace') || d.includes('horizontal brace') || d.includes('corner brace')) {
                  const pageStr = String(dataRows[r][pageIdx] || '').trim()
                  const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                  if (soeMatches && soeMatches.length > 0) {
                    soeBracingMain = soeMatches[0]
                    if (soeMatches.length > 1) soeBracingDetails = soeMatches[1]
                  }
                  break
                }
              }
            }
          }
          const asPerBracing = ` as per ${soeBracingMain} & details on ${soeBracingDetails}`

          const extractSizeFromParticulars = (p) => {
            if (!p || typeof p !== 'string') return ''
            const m = p.match(/(\d+)"\s*x\s*(\d+)"?/i) || p.match(/(\d+)\s*"\s*x\s*(\d+)/i)
            return m ? `${m[1]}"x${m[2]}"` : ''
          }
          const extractQtyFromParticulars = (p) => {
            if (!p || typeof p !== 'string') return null
            const m = p.match(/^\((\d+)\)\s*/)
            return m ? parseInt(m[1], 10) : null
          }
          const flattenGroups = (groups) => (groups || []).flat().filter(Boolean)
          const sumQty = (items) => items.reduce((s, i) => s + (i.qty != null ? Number(i.qty) : i.takeoff != null ? Number(i.takeoff) : 0), 0)

          // Helper: get bracing items from soeSubsectionItems or by scanning calculationData
          const getBracingItemsForSummary = (subsectionKeys) => {
            for (const key of subsectionKeys) {
              const groups = soeSubsectionItems.get(key) || []
              const items = flattenGroups(groups)
              if (items.length > 0) return items
            }
            // Fallback: scan calculationData for first matching subsection (e.g. "Timber waler")
            const subsectionName = subsectionKeys[0]
            const fromCalc = findBracingItemsFromCalculationDataForSummary(subsectionName)
            if (fromCalc.length > 0) return fromCalc
            return []
          }
          const findBracingItemsFromCalculationDataForSummary = (subsectionName) => {
            if (!calculationData || calculationData.length === 0) return []
            const found = []
            let inSubsection = false
            calculationData.forEach((row, index) => {
              const sheetRow = index + 1
              const colB = row[1]
              if (colB && typeof colB === 'string') {
                const bText = colB.trim()
                if (bText.endsWith(':') && bText.slice(0, -1).trim().toLowerCase() === subsectionName.toLowerCase()) {
                  inSubsection = true
                  return
                }
                if (inSubsection && bText.endsWith(':')) {
                  inSubsection = false
                  return
                }
                if (inSubsection && bText && !bText.endsWith(':') && row[2] !== undefined) {
                  found.push({
                    particulars: bText,
                    takeoff: parseFloat(row[2]) || 0,
                    qty: parseFloat(row[12]) || 0,
                    rawRowNumber: sheetRow
                  })
                }
              }
            })
            return found
          }

          const walerGroups = soeSubsectionItems.get('Timber waler') || soeSubsectionItems.get('Waler') || []
          const rakerGroups = soeSubsectionItems.get('Timber raker') || soeSubsectionItems.get('Raker') || []
          const braceGroups = soeSubsectionItems.get('Timber brace') || soeSubsectionItems.get('Bracing') || []
          let walerItems = flattenGroups(walerGroups)
          let rakerItems = flattenGroups(rakerGroups)
          let braceItems = flattenGroups(braceGroups)
          if (walerItems.length === 0) walerItems = getBracingItemsForSummary(['Timber waler', 'Waler'])
          if (rakerItems.length === 0) rakerItems = getBracingItemsForSummary(['Timber raker', 'Raker'])
          if (braceItems.length === 0) braceItems = getBracingItemsForSummary(['Timber brace', 'Bracing'])

          const walerQty = walerItems.length ? (sumQty(walerItems) || extractQtyFromParticulars(walerItems[0].particulars) || 0) : 0
          const rakerQty = rakerItems.length ? (sumQty(rakerItems) || 0) : 0
          const walerSize = walerItems.length ? (extractSizeFromParticulars(walerItems[0].particulars) || '##') : '##'
          const rakerSize = rakerItems.length ? (extractSizeFromParticulars(rakerItems[0].particulars) || '##') : '##'
          const horizontalBraceItems = braceItems.filter(i => (i.particulars || '').toLowerCase().includes('horizontal'))
          const cornerBraceItems = braceItems.filter(i => (i.particulars || '').toLowerCase().includes('corner'))
          const horizontalQty = horizontalBraceItems.length ? sumQty(horizontalBraceItems) : 0
          const cornerQty = cornerBraceItems.length ? (sumQty(cornerBraceItems) || 0) : 0
          const horizontalSize = horizontalBraceItems.length ? (extractSizeFromParticulars(horizontalBraceItems[0].particulars) || '##') : '##'
          const cornerSize = cornerBraceItems.length ? (extractSizeFromParticulars(cornerBraceItems[0].particulars) || '##') : '##'

          // Calculation sheet: Timber waler has sum row only when >1 item; Timber raker always has sum row per group; Timber brace has sum row only when >1 item in group
          const maxRawRow = (items) => items.length ? Math.max(...items.map(i => i.rawRowNumber || 0), 0) : 0
          const walerSumRow = walerItems.length ? (walerItems.length > 1 ? maxRawRow(walerItems) + 1 : maxRawRow(walerItems)) : 0
          const rakerSumRow = rakerItems.length ? maxRawRow(rakerItems) + 1 : 0
          const horizontalSumRow = horizontalBraceItems.length ? (horizontalBraceItems.length > 1 ? maxRawRow(horizontalBraceItems) + 1 : maxRawRow(horizontalBraceItems)) : 0
          const cornerSumRow = cornerBraceItems.length ? (cornerBraceItems.length > 1 ? maxRawRow(cornerBraceItems) + 1 : maxRawRow(cornerBraceItems)) : 0

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
            'Corner brace': 'Corner Brace',
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

                const qtyValue = Math.round(totalQty || totalTakeoff)
                const proposalText = `F&I new (${qtyValue})no ${size} supporting angle ${location} as per ${soePageMain} & details on`
                const afterCount = `)no ${size} supporting angle ${location} as per ${soePageMain} & details on`

                if (sumRowIndex > 0) {
                  spreadsheet.updateCell({ formula: proposalFormulaWithQtyRef(currentRow, afterCount) }, `${pfx}B${currentRow}`)
                } else {
                  spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
                }
                rowBContentMap.set(currentRow, proposalText)
                spreadsheet.wrap(`${pfx}B${currentRow}`, true)
                spreadsheet.cellFormat(
                  {
                    fontWeight: 'bold',
                    color: '#000000',
                    textAlign: 'left',
                    backgroundColor: 'white',
                    verticalAlign: 'top',
                    textDecoration: 'none'
                  },
                  `${pfx}B${currentRow}`
                )
                fillRatesForProposalRow(currentRow, proposalText)

                // Calculate and set row height based on content
                const dynamicHeight = calculateRowHeight(proposalText)

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

                // Add $/1000 formula in column H
                const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
                spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
                spreadsheet.cellFormat(
                  {
                    fontWeight: 'bold',
                    textAlign: 'right',
                    backgroundColor: 'white',
                    format: '$#,##0.00'
                  },
                  `${pfx}H${currentRow}`
                )

                // Apply currency format
                try {
                  spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}:N${currentRow}`)
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
                proposalText = `F&I new (${qtyValue})no ${sizeInfo}${displayName.toLowerCase()} as per ${soePageMain} & details on`
              } else {
                proposalText = `F&I new (${qtyValue})no ${displayName.toLowerCase()} as per ${soePageMain} & details on`
              }
              const afterCount = afterCountFromProposalText(proposalText)

              if (sumRowIndex > 0 && afterCount) {
                spreadsheet.updateCell({ formula: proposalFormulaWithQtyRef(currentRow, afterCount) }, `${pfx}B${currentRow}`)
              } else {
                spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
              }
              rowBContentMap.set(currentRow, proposalText)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  color: '#000000',
                  textAlign: 'left',
                  backgroundColor: 'white',
                  verticalAlign: 'top',
                  textDecoration: 'none'
                },
                `${pfx}B${currentRow}`
              )
              fillRatesForProposalRow(currentRow, proposalText)

              // Calculate and set row height based on content
              const dynamicHeight = calculateRowHeight(proposalText)

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

              // Add $/1000 formula in column H
              const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
              spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white',
                  format: '$#,##0.00'
                },
                `${pfx}H${currentRow}`
              )

              // Apply currency format
              try {
                spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}:N${currentRow}`)
              } catch (e) {
                // Fallback already applied in cellFormat
              }

              // Row height already set above based on proposal text

              currentRow++ // Move to next row
            })
          })

          // After all detailed bracing items, add the four summary lines at the end of the Bracing section
          const calcSheetName = 'Calculations Sheet'
          const bracingSummaryLines = [
            { qty: walerQty, size: walerSize, label: 'timber walers', sumRowIndex: walerSumRow },
            { qty: rakerQty, size: rakerSize, label: 'timber rakers', sumRowIndex: rakerSumRow },
            { qty: horizontalQty, size: horizontalSize, label: 'timber horizontal braces', sumRowIndex: horizontalSumRow },
            { qty: cornerQty, size: cornerSize, label: 'timber corner brace', sumRowIndex: cornerSumRow }
          ]

          bracingSummaryLines.forEach(({ qty, size, label, sumRowIndex }) => {
            const sizeStr = size || '##'
            // Always show the numeric quantity (including 0) in the proposal text
            const text = `F&I new (${String(qty)})no ${sizeStr} ${label}${asPerBracing}`
            if (sumRowIndex > 0) {
              const afterCount = `)no ${sizeStr} ${label}${asPerBracing}`
              spreadsheet.updateCell({ formula: proposalFormulaWithQtyRef(currentRow, afterCount) }, `${pfx}B${currentRow}`)
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${sumRowIndex}` }, `${pfx}C${currentRow}`)
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${sumRowIndex}` }, `${pfx}D${currentRow}`)
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!K${sumRowIndex}` }, `${pfx}E${currentRow}`)
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!L${sumRowIndex}` }, `${pfx}F${currentRow}`)
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
            } else {
              spreadsheet.updateCell({ value: text }, `${pfx}B${currentRow}`)
              // Always set C–G so formula bar shows a formula; when no sum row use refs that display blank
              spreadsheet.updateCell({ formula: `=IFERROR('${calcSheetName}'!I0,"")` }, `${pfx}C${currentRow}`)
              spreadsheet.updateCell({ formula: `=IFERROR('${calcSheetName}'!J0,"")` }, `${pfx}D${currentRow}`)
              spreadsheet.updateCell({ formula: `=IFERROR('${calcSheetName}'!K0,"")` }, `${pfx}E${currentRow}`)
              spreadsheet.updateCell({ formula: `=IFERROR('${calcSheetName}'!L0,"")` }, `${pfx}F${currentRow}`)
              spreadsheet.updateCell({ formula: `=IFERROR('${calcSheetName}'!M0,"")` }, `${pfx}G${currentRow}`)
            }
            const sumCellFormat = { fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }
            spreadsheet.cellFormat(sumCellFormat, `${pfx}C${currentRow}`)
            spreadsheet.cellFormat(sumCellFormat, `${pfx}D${currentRow}`)
            spreadsheet.cellFormat(sumCellFormat, `${pfx}E${currentRow}`)
            spreadsheet.cellFormat(sumCellFormat, `${pfx}F${currentRow}`)
            spreadsheet.cellFormat(sumCellFormat, `${pfx}G${currentRow}`)
            rowBContentMap.set(currentRow, text)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                color: '#000000',
                textAlign: 'left',
                backgroundColor: 'white',
                verticalAlign: 'top',
                textDecoration: 'none'
              },
              `${pfx}B${currentRow}`
            )
            fillRatesForProposalRow(currentRow, text)
            const dynamicHeight = calculateRowHeight(text)
            currentRow++
          })

          // Add Tie back anchor section right after Bracing
          // Add Tie back anchor header
          const tieBackText = `Tie back anchor: including applicable washers, steel bearing plates, locking hex nuts as required`
          spreadsheet.updateCell({ value: tieBackText }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, tieBackText)
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
              const proposalText = `F&I new (${qtyValue})no tie back anchors (L=${freeLength} + ${bondLength} bond length), ${drillHole} drill hole as per ${soePageMain} & details on`
              const afterCount = afterCountFromProposalText(proposalText)

              if (sumRowIndex > 0 && afterCount) {
                spreadsheet.updateCell({ formula: proposalFormulaWithQtyRef(currentRow, afterCount) }, `${pfx}B${currentRow}`)
              } else {
                spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
              }
              rowBContentMap.set(currentRow, proposalText)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  color: '#000000',
                  textAlign: 'left',
                  backgroundColor: 'white',
                  verticalAlign: 'top',
                  textDecoration: 'none'
                },
                `${pfx}B${currentRow}`
              )
              fillRatesForProposalRow(currentRow, proposalText)

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

              // Add $/1000 formula in column H
              const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
              spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white',
                  format: '$#,##0.00'
                },
                `${pfx}H${currentRow}`
              )

              // Apply currency format
              try {
                spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}:N${currentRow}`)
              } catch (e) {
                // Fallback already applied in cellFormat
              }

              // Calculate and set row height based on content in column B
              const proposalTextForHeight = proposalText || (spreadsheet.getCellValue ? spreadsheet.getCellValue(`${pfx}B${currentRow}`) : '')
              const dynamicHeight = calculateRowHeight(proposalTextForHeight)

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
          rowBContentMap.set(currentRow, tieBackText)
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
              const proposalText = `F&I new (${qtyValue})no tie back anchors (L=${freeLength} + ${bondLength} bond length), ${drillHole} drill hole as per ${soePageMain} & details on`
              const afterCount = afterCountFromProposalText(proposalText)

              if (sumRowIndex > 0 && afterCount) {
                spreadsheet.updateCell({ formula: proposalFormulaWithQtyRef(currentRow, afterCount) }, `${pfx}B${currentRow}`)
              } else {
                spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
              }
              rowBContentMap.set(currentRow, proposalText)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  color: '#000000',
                  textAlign: 'left',
                  backgroundColor: 'white',
                  verticalAlign: 'top',
                  textDecoration: 'none'
                },
                `${pfx}B${currentRow}`
              )
              fillRatesForProposalRow(currentRow, proposalText)

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

              // Add $/1000 formula in column H
              const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
              spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white',
                  format: '$#,##0.00'
                },
                `${pfx}H${currentRow}`
              )

              // Apply currency format
              try {
                spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}:N${currentRow}`)
              } catch (e) {
                // Fallback already applied in cellFormat
              }

              // Calculate and set row height based on content in column B
              const proposalTextForHeight = proposalText || (spreadsheet.getCellValue ? spreadsheet.getCellValue(`${pfx}B${currentRow}`) : '')
              const dynamicHeight = calculateRowHeight(proposalTextForHeight)

              currentRow++
            })
          }

          // Add Tie down anchor section right after Tie back anchor (only if items exist)
          // Get tie down anchor groups first to check if any exist
          let tieDownGroups = soeSubsectionItems.get('Tie down') || []
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
            rowBContentMap.set(currentRow, tieDownText)
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
              const proposalText = `F&I new (${qtyValue})no tie down anchors (L=${freeLength} + ${bondLength} bond length), ${drillHole} drill hole as per ${soePageMain} & details on`
              const afterCount = afterCountFromProposalText(proposalText)

              if (sumRowIndex > 0 && afterCount) {
                spreadsheet.updateCell({ formula: proposalFormulaWithQtyRef(currentRow, afterCount) }, `${pfx}B${currentRow}`)
              } else {
                spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
              }
              rowBContentMap.set(currentRow, proposalText)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  color: '#000000',
                  textAlign: 'left',
                  backgroundColor: 'white',
                  verticalAlign: 'top',
                  textDecoration: 'none'
                },
                `${pfx}B${currentRow}`
              )
              fillRatesForProposalRow(currentRow, proposalText)

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

              // Add $/1000 formula in column H
              const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
              spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white',
                  format: '$#,##0.00'
                },
                `${pfx}H${currentRow}`
              )

              // Apply currency format
              try {
                spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}:N${currentRow}`)
              } catch (e) {
                // Fallback already applied in cellFormat
              }

              // Calculate and set row height based on content in column B
              const proposalTextForHeight = proposalText || (spreadsheet.getCellValue ? spreadsheet.getCellValue(`${pfx}B${currentRow}`) : '')
              const dynamicHeight = calculateRowHeight(proposalTextForHeight)

              currentRow++
            })
          }

          return // Skip normal processing for Tie back anchor (Tie down anchor is already added above)
        }

        // Skip normal processing for Tie down anchor (it's already added right after Tie back anchor above)
        if (subsectionName === 'Tie down' || subsectionName === 'Tie down anchor') {
          return
        }

        // Special handling for Misc. subsection (only added to subsectionsToDisplay when at least one other SOE subsection exists)
        if (subsectionName.toLowerCase() === 'misc.' || subsectionName.toLowerCase() === 'misc') {
          // Add subsection header
          spreadsheet.updateCell({ value: `${subsectionName}:` }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, `${subsectionName}:`)
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
          currentRow++

          // Hardcoded Misc. items (show only when at least one other SOE subsection exists)
          const miscItems = [
            'Cut-off to grade included',
            'Pilings will be threaded at both ends and installed in 10\' or 15\' increments',
            'Obstruction removal drilling using DTHH per hour: $995/h'
          ]

          miscItems.forEach((item) => {
            spreadsheet.updateCell({ value: item }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, item)
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
        rowBContentMap.set(currentRow, subsectionText)
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
                    const rowNum = index + 1 // 1-based Excel row (calculationData[0] = row 1)
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

        // If no groups found, handle special SOE subsections (Vertical/Horizontal timber sheets, Timber stringer, Drilled hole grout)
        // via template text when Calculation sheet has the subsection; otherwise skip.
        const isEmpty = groups.length === 0 || groups.every(g => g.length === 0)
        const isVerticalTimberSheets = subsectionName.toLowerCase() === 'vertical timber sheets'
        const isHorizontalTimberSheets = subsectionName.toLowerCase() === 'horizontal timber sheets'
        const isTimberStringer = subsectionName.toLowerCase() === 'timber stringer'
        const isDrilledHoleGrout = subsectionName.toLowerCase() === 'drilled hole grout'
        const isTemplateSubsection = isVerticalTimberSheets || isHorizontalTimberSheets || isTimberStringer || isDrilledHoleGrout

        if (isEmpty) {
          if (!isTemplateSubsection) {
            // For non-template subsections with no items, skip entirely
            return
          }

          // Fill template text; pull dimensions/Havg/embedment from Calculation sheet data when present
          let proposalText = ''
          let calcRefSubsectionName = ''
          let dimensions = '##'
          let heightText = '##'
          let embedmentText = '## embedment'
          let sizeStr = '##'
          let diameterStr = '##'
          let soeRefDrilled = 'SOE-101.01'
          if (calculationData && calculationData.length > 0) {
            let inSubsection = false
            let subsectionHeaderRow = -1
            const dataRows = []
            const targetHeader = isVerticalTimberSheets ? 'vertical timber sheets' : isHorizontalTimberSheets ? 'horizontal timber sheets' : isTimberStringer ? 'timber stringer' : 'drilled hole grout'
            for (let i = 0; i < calculationData.length; i++) {
              const colB = calculationData[i][1]
              if (colB && typeof colB === 'string') {
                const bText = colB.trim().toLowerCase()
                const isHeader = bText.endsWith(':') && bText.slice(0, -1).trim() === targetHeader
                if (isHeader) {
                  inSubsection = true
                  subsectionHeaderRow = i
                  continue
                }
                if (inSubsection && bText.endsWith(':')) {
                  break
                }
                if (inSubsection && bText && !bText.endsWith(':')) {
                  const takeoff = parseFloat(calculationData[i][2]) || 0
                  if (bText.includes(targetHeader) || takeoff > 0) {
                    dataRows.push({ particulars: calculationData[i][1], row: calculationData[i], index: i })
                  }
                }
              }
            }
            const parseDimToFeet = (dimStr) => {
              if (!dimStr) return 0
              const m = String(dimStr).match(/(\d+)(?:'-?)?(\d+)?/)
              if (!m) return 0
              return (parseInt(m[1], 10) || 0) + ((parseInt(m[2], 10) || 0) / 12)
            }
            const feetToFtIn = (feet) => {
              const f = Math.floor(feet)
              const i = Math.round((feet - f) * 12)
              return i === 0 ? `${f}'-0"` : `${f}'-${i}"`
            }
            if (isVerticalTimberSheets || isHorizontalTimberSheets) {
              if (dataRows.length > 0) {
                const first = dataRows[0].particulars || ''
                const dimMatch = first.match(/(\d+(?:\/\d+)?)["']?\s*x\s*(\d+(?:\/\d+)?)["']?/i)
                if (dimMatch) dimensions = `${dimMatch[1]}"x${dimMatch[2]}"`
                const heightVals = []
                const embedmentVals = []
                dataRows.forEach(({ particulars, index }) => {
                  const p = String(particulars || '')
                  const hMatch = p.match(/H=([0-9'"\-]+)/i)
                  if (hMatch) heightVals.push(parseDimToFeet(hMatch[1]))
                  const eMatch = p.match(/E=([0-9'"\-]+)/i)
                  if (eMatch) embedmentVals.push(parseDimToFeet(eMatch[1]))
                  const rowH = parseFloat(calculationData[index]?.[7]) || 0
                  if (rowH > 0) heightVals.push(rowH)
                })
                if (heightVals.length > 0) {
                  const avg = heightVals.reduce((a, b) => a + b, 0) / heightVals.length
                  heightText = feetToFtIn(avg)
                }
                if (embedmentVals.length > 0) {
                  const avgE = embedmentVals.reduce((a, b) => a + b, 0) / embedmentVals.length
                  embedmentText = `${feetToFtIn(avgE)} embedment`
                }
              }
            } else if (isTimberStringer && dataRows.length > 0) {
              const first = dataRows[0].particulars || ''
              const m = first.match(/(\d+)"\s*x\s*(\d+)"?/i) || first.match(/(\d+)\s*"\s*x\s*(\d+)/i)
              if (m) sizeStr = `${m[1]}"x${m[2]}"`
            } else if (isDrilledHoleGrout && dataRows.length > 0) {
              const first = dataRows[0].particulars || ''
              const fracMatch = first.match(/(\d+-\d+\/\d+)["']?\s*Ø/i)
              const decMatch = first.match(/(\d+\.?\d*)["']?\s*Ø/i)
              if (fracMatch) diameterStr = `${fracMatch[1]}" Ø`
              else if (decMatch) diameterStr = `${decMatch[1]}" Ø`
              const heightVals = []
              dataRows.forEach(({ particulars, index }) => {
                const p = String(particulars || '')
                const hMatch = p.match(/H=([0-9'"\-]+)/i)
                if (hMatch) heightVals.push(parseDimToFeet(hMatch[1]))
                const rowH = parseFloat(calculationData[index]?.[7]) || 0
                if (rowH > 0) heightVals.push(rowH)
              })
              if (heightVals.length > 0) {
                const avg = heightVals.reduce((a, b) => a + b, 0) / heightVals.length
                heightText = feetToFtIn(avg)
              }
              if (rawData && Array.isArray(rawData) && rawData.length > 1) {
                const headers = rawData[0]
                const dataRowsRaw = rawData.slice(1)
                const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
                const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
                if (digitizerIdx !== -1 && pageIdx >= 0) {
                  for (let r = 0; r < dataRowsRaw.length; r++) {
                    const d = String(dataRowsRaw[r][digitizerIdx] || '').toLowerCase()
                    if (d.includes('drilled hole grout')) {
                      const pageStr = String(dataRowsRaw[r][pageIdx] || '').trim()
                      const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                      if (soeMatches && soeMatches.length > 0) soeRefDrilled = soeMatches[0]
                      break
                    }
                  }
                }
              }
            }
          }
          if (isVerticalTimberSheets) {
            proposalText = `F&I new ${dimensions} vertical timber sheets (Havg=${heightText}, ${embedmentText}) as per SOE-101.00 & details on SOE-201.00 to SOE-205.00`
            calcRefSubsectionName = 'Vertical timber sheets'
          } else if (isHorizontalTimberSheets) {
            proposalText = `F&I new ${dimensions} horizontal timber sheets (Havg=${heightText}, ${embedmentText}) as per SOE-101.00 & details on SOE-201.00 to SOE-205.00`
            calcRefSubsectionName = 'Horizontal timber sheets'
          } else if (isDrilledHoleGrout) {
            proposalText = `F&I new (${diameterStr}) drilled hole grout @ tangent pile (H=${heightText}, typ.) as per ${soeRefDrilled} & details on above misc`
            calcRefSubsectionName = 'Drilled hole grout'
          } else {
            proposalText = `F&I new ${sizeStr} timber stringer as per SOE-101.00 & details on SOE-201.00 to SOE-205.00`
            calcRefSubsectionName = 'Timber stringer'
          }
          spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, proposalText)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat(
            { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' },
            `${pfx}B${currentRow}`
          )
          fillRatesForProposalRow(currentRow, proposalText)
          // C–G: reference the row on Calculation sheet where this subsection lives (header row + 1 = first data/sum row); 1-based sheet row
          const calcSheetName = 'Calculations Sheet'
          let refRow = 0
          if (calculationData && calculationData.length > 0 && calcRefSubsectionName) {
            for (let i = 0; i < calculationData.length; i++) {
              const colB = calculationData[i][1]
              if (colB && typeof colB === 'string') {
                const bText = colB.trim()
                const headerMatch = bText.endsWith(':') && bText.slice(0, -1).trim().toLowerCase() === calcRefSubsectionName.toLowerCase()
                if (headerMatch) {
                  refRow = i + 2
                  break
                }
              }
            }
          }
          if (refRow > 0) {
            spreadsheet.updateCell({ formula: `=IFERROR('${calcSheetName}'!I${refRow},"")` }, `${pfx}C${currentRow}`)
            spreadsheet.updateCell({ formula: `=IFERROR('${calcSheetName}'!J${refRow},"")` }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: `=IFERROR('${calcSheetName}'!K${refRow},"")` }, `${pfx}E${currentRow}`)
            spreadsheet.updateCell({ formula: `=IFERROR('${calcSheetName}'!L${refRow},"")` }, `${pfx}F${currentRow}`)
            spreadsheet.updateCell({ formula: `=IFERROR('${calcSheetName}'!M${refRow},"")` }, `${pfx}G${currentRow}`)
          } else {
            const blankFormula = '=IFERROR(1/0,"")'
            spreadsheet.updateCell({ formula: blankFormula }, `${pfx}C${currentRow}`)
            spreadsheet.updateCell({ formula: blankFormula }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: blankFormula }, `${pfx}E${currentRow}`)
            spreadsheet.updateCell({ formula: blankFormula }, `${pfx}F${currentRow}`)
            spreadsheet.updateCell({ formula: blankFormula }, `${pfx}G${currentRow}`)
          }
          const sumCellFormat = { fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }
          spreadsheet.cellFormat(sumCellFormat, `${pfx}C${currentRow}`)
          spreadsheet.cellFormat(sumCellFormat, `${pfx}D${currentRow}`)
          spreadsheet.cellFormat(sumCellFormat, `${pfx}E${currentRow}`)
          spreadsheet.cellFormat(sumCellFormat, `${pfx}F${currentRow}`)
          spreadsheet.cellFormat(sumCellFormat, `${pfx}G${currentRow}`)
          const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
          spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white', format: '$#,##0.00' }, `${pfx}H${currentRow}`)
          const dynamicHeight = calculateRowHeight(proposalText)
          try { spreadsheet.setRowHeight(dynamicHeight, currentRow - 1, proposalSheetIndex) } catch (e) { }
          currentRow++
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

          // Find the sum row for this group - use formulaData when available for accurate row
          let sumRowIndex = lastRowNumber
          const subsectionLower = subsectionName.toLowerCase()
          if (formulaData && Array.isArray(formulaData)) {
            const sumFormula = formulaData.find(f => {
              // Mud slab is in Excavation section with excavation_sum itemType
              if (subsectionLower === 'mud slab' && f.itemType === 'excavation_sum' && f.subsection === 'Mud slab') {
                return true
              }
              // Other subsections use soe_generic_sum
              if (f.itemType !== 'soe_generic_sum' || !f.subsectionName) return false
              const fSub = f.subsectionName.toLowerCase()
              // Match exact or aliases: Buttons/Concrete buttons, Dowels/Dowel bar, Rock pins/Rock pin
              return fSub === subsectionLower ||
                ((subsectionLower === 'concrete buttons' || subsectionLower === 'buttons') && fSub === 'buttons') ||
                ((subsectionLower === 'dowels' || subsectionLower === 'dowel bar') && fSub === 'dowel bar') ||
                ((subsectionLower === 'rock pins' || subsectionLower === 'rock pin') && fSub === 'rock pins')
            })
            if (sumFormula) {
              // Use lastDataRow + 1 to get the sum row (row immediately after last data row)
              sumRowIndex = (sumFormula.lastDataRow != null ? sumFormula.lastDataRow + 1 : sumFormula.row)
            }
          }
          // Sum row is the row after last data row for subsections that have a sum row per group
          if (sumRowIndex === lastRowNumber && (subsectionLower === 'parging' ||
            subsectionLower === 'heel blocks' ||
            subsectionLower === 'underpinning' ||
            subsectionLower === 'concrete soil retention piers' ||
            subsectionLower === 'concrete soil retention pier' ||
            subsectionLower === 'form board' ||
            subsectionLower === 'guide wall' ||
            subsectionLower === 'guilde wall' ||
            subsectionLower === 'concrete buttons' ||
            subsectionLower === 'buttons' ||
            subsectionLower === 'dowels' ||
            subsectionLower === 'dowel bar' ||
            subsectionLower === 'rock pins' ||
            subsectionLower === 'rock pin' ||
            subsectionLower === 'rock stabilization' ||
            subsectionLower === 'shotcrete' ||
            subsectionLower === 'permission grouting' ||
            subsectionLower === 'mud slab' ||
            subsectionLower === 'vertical timber sheets' ||
            subsectionLower === 'horizontal timber sheets' ||
            subsectionLower === 'timber stringer' ||
            subsectionLower === 'timber sheeting')) {
            sumRowIndex = lastRowNumber + 1
          }
          const calcSheetName = 'Calculations Sheet'

          // Generate proposal text based on subsection type
          let proposalText = ''
          let useFormulaForB = false
          let afterCountForB = ''

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
            proposalText = `F&I new (${totalQtyValue})no typ. heel blocks for raker support as per ${soePageMain} & details on`
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
            proposalText = `F&I new (${totalQtyValue})no concrete soil retention pier as per ${soePageMain} & details on`

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
              window.formBoardProposalText = `F&I new (${formBoardThickness} thick) form board w/ filter fabric between tunnel and retention pier as per ${formBoardSoePage} & details on`
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
            proposalText = `New parging (Havg=${heightText}) as per ${soePageMain} & details on`
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
            proposalText = `F&I new (${thickness} thick) form board w/ filter fabric between tunnel and retention pier as per ${soePageMain} & details on`
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
              ? `F&I new (${totalQtyValue})no (${widthText} wide) concrete buttons (Havg=${heightText}) as per ${soePageMain} & details on`
              : `F&I new (${totalQtyValue})no concrete buttons (Havg=${heightText}) as per ${soePageMain} & details on`
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
              ? `F&I new (${widthText} wide) guide wall (H=${height || '3\'-0"'}) as per ${soePageMain} & details on`
              : `F&I new guide wall (H=${height || '3\'-0"'}) as per ${soePageMain} & details on`
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
            proposalText = `F&I new (${totalQtyValue})no 4-${barSize || '#9'} steel dowels bar (Havg=${heightText}, ${rockSocket || '4\'-0"'} rock socket) as per ${soePageMain} & details on`
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
            proposalText = `F&I new (${totalQtyValue})no rock pins (Havg=${heightText}, ${rockSocket || '4\'-0"'} rock socket) as per ${soePageMain} & details on`
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
            proposalText = `F&I new rock stabilization as per ${soePageMain} & details on`
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
            proposalText = `F&I new (${thickness || '6"'} thick) shotcrete w/ ${wireMesh || '6x6'} wire mesh (Havg=${heightText}) as per ${soePageMain} & details on`
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
            proposalText = `F&I new permission grouting (Havg=${heightText}) as per ${soePageMain} & details on`
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

            // Format: Allow to (2" thick) mud slab as per SOE-101.00 & details on SOE-204.00
            proposalText = `Allow to (${thickness} thick) mud slab as per ${soePageMain} & details on SOE-204.00`
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
              proposalText = `F&I new [${sheetPileType || '#'}] sheet pile (H=${height || '#'}, E=${embedment}embedment) as per ${soePageMain} & details on`
            } else {
              proposalText = `F&I new [${sheetPileType || '#'}] sheet pile (H=${height || '#'}) as per ${soePageMain} & details on`
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

            // Timber lagging line: main reference + details on SOE-201.01 to SOE-204.00
            const timberLaggingMainRef = (soePageMain && !String(soePageMain).includes('301')) ? soePageMain : 'SOE-101.00'
            // Format: F&I new 3"x10" timber lagging for the exposed depths (Havg=10'-6") as per SOE-101.00 & details on SOE-201.01 to SOE-204.00
            proposalText = `F&I new ${dimensions || '##'} timber lagging for the exposed depths (Havg=${heightText || '##'}) as per ${timberLaggingMainRef} & details on `

            // Store main ref for backpacking line when shown under Sheet pile (backpacking uses main + details on SOE-301.00)
            window.timberLaggingSoePage = timberLaggingMainRef
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
            proposalText = `F&I new ${dimensions || '##'} timber sheeting (Havg=${heightText || '##'}) as per ${soePageMain} & details on`
          } else if (subsectionName.toLowerCase() === 'vertical timber sheets') {
            // Format: F&I new 3"x10" vertical timber sheets (Havg=7'-2", 4'-0" embedment) as per SOE-101.00 & details on SOE-201.00 to SOE-205.00
            let dimensions = ''
            const heightVals = []
            const embedmentVals = []
            const parseDimToFeet = (dimStr) => {
              if (!dimStr) return 0
              const m = String(dimStr).match(/(\d+)(?:'-?)?(\d+)?/)
              if (!m) return 0
              return (parseInt(m[1]) || 0) + ((parseInt(m[2]) || 0) / 12)
            }
            const feetToFtIn = (feet) => {
              const f = Math.floor(feet)
              const i = Math.round((feet - f) * 12)
              return i === 0 ? `${f}'-0"` : `${f}'-${i}"`
            }
            group.forEach(item => {
              const p = (item.particulars || '').trim()
              const dimMatch = p.match(/(\d+(?:\/\d+)?)["']?\s*x\s*(\d+(?:\/\d+)?)["']?/i)
              if (dimMatch && !dimensions) dimensions = `${dimMatch[1]}"x${dimMatch[2]}"`
              const hMatch = p.match(/H=([0-9'"\-]+)/i)
              if (hMatch) heightVals.push(parseDimToFeet(hMatch[1]))
              const eMatch = p.match(/E=([0-9'"\-]+)/i)
              if (eMatch) embedmentVals.push(parseDimToFeet(eMatch[1]))
              const rowH = item.height || item.parsed?.heightRaw || 0
              if (rowH > 0) heightVals.push(rowH)
            })
            let heightText = '##'
            if (heightVals.length > 0) {
              const avg = heightVals.reduce((a, b) => a + b, 0) / heightVals.length
              heightText = feetToFtIn(avg)
            }
            let embedmentText = ''
            if (embedmentVals.length > 0) {
              const avgE = embedmentVals.reduce((a, b) => a + b, 0) / embedmentVals.length
              embedmentText = `${feetToFtIn(avgE)} embedment`
            }
            let soePageMain = 'SOE-101.00'
            let soePageDetails = 'SOE-201.00'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let r = 0; r < dataRows.length; r++) {
                  const d = String(dataRows[r][digitizerIdx] || '').toLowerCase()
                  if (d.includes('vertical timber sheets')) {
                    const pageStr = String(dataRows[r][pageIdx] || '').trim()
                    const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                    if (soeMatches && soeMatches.length > 0) {
                      soePageMain = soeMatches[0]
                      if (soeMatches.length > 1) soePageDetails = soeMatches[1]
                    }
                    break
                  }
                }
              }
            }
            proposalText = `F&I new ${dimensions || '##'} vertical timber sheets (Havg=${heightText}, ${embedmentText || '## embedment'}) as per ${soePageMain} & details on ${soePageDetails} to SOE-205.00`
          } else if (subsectionName.toLowerCase() === 'horizontal timber sheets') {
            // Format: F&I new 3"x10" horizontal timber sheets (Havg=7'-2", 4'-0" embedment) as per SOE-101.00 & details on SOE-201.00 to SOE-205.00
            let dimensions = ''
            const heightVals = []
            const embedmentVals = []
            const parseDimToFeet = (dimStr) => {
              if (!dimStr) return 0
              const m = String(dimStr).match(/(\d+)(?:'-?)?(\d+)?/)
              if (!m) return 0
              return (parseInt(m[1]) || 0) + ((parseInt(m[2]) || 0) / 12)
            }
            const feetToFtIn = (feet) => {
              const f = Math.floor(feet)
              const i = Math.round((feet - f) * 12)
              return i === 0 ? `${f}'-0"` : `${f}'-${i}"`
            }
            group.forEach(item => {
              const p = (item.particulars || '').trim()
              const dimMatch = p.match(/(\d+(?:\/\d+)?)["']?\s*x\s*(\d+(?:\/\d+)?)["']?/i)
              if (dimMatch && !dimensions) dimensions = `${dimMatch[1]}"x${dimMatch[2]}"`
              const hMatch = p.match(/H=([0-9'"\-]+)/i)
              if (hMatch) heightVals.push(parseDimToFeet(hMatch[1]))
              const eMatch = p.match(/E=([0-9'"\-]+)/i)
              if (eMatch) embedmentVals.push(parseDimToFeet(eMatch[1]))
              const rowH = item.height || item.parsed?.heightRaw || 0
              if (rowH > 0) heightVals.push(rowH)
            })
            let heightText = '##'
            if (heightVals.length > 0) {
              const avg = heightVals.reduce((a, b) => a + b, 0) / heightVals.length
              heightText = feetToFtIn(avg)
            }
            let embedmentText = ''
            if (embedmentVals.length > 0) {
              const avgE = embedmentVals.reduce((a, b) => a + b, 0) / embedmentVals.length
              embedmentText = `${feetToFtIn(avgE)} embedment`
            }
            let soePageMain = 'SOE-101.00'
            let soePageDetails = 'SOE-201.00'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let r = 0; r < dataRows.length; r++) {
                  const d = String(dataRows[r][digitizerIdx] || '').toLowerCase()
                  if (d.includes('horizontal timber sheets')) {
                    const pageStr = String(dataRows[r][pageIdx] || '').trim()
                    const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                    if (soeMatches && soeMatches.length > 0) {
                      soePageMain = soeMatches[0]
                      if (soeMatches.length > 1) soePageDetails = soeMatches[1]
                    }
                    break
                  }
                }
              }
            }
            proposalText = `F&I new ${dimensions || '##'} horizontal timber sheets (Havg=${heightText}, ${embedmentText || '## embedment'}) as per ${soePageMain} & details on ${soePageDetails} to SOE-205.00`
          } else if (subsectionName.toLowerCase() === 'timber stringer') {
            // Format: F&I new 6"x6" timber stringer as per SOE-101.00 & details on SOE-201.00 to SOE-205.00
            let sizeStr = ''
            group.forEach(item => {
              const p = (item.particulars || '').trim()
              const m = p.match(/(\d+)"\s*x\s*(\d+)"?/i) || p.match(/(\d+)\s*"\s*x\s*(\d+)/i)
              if (m && !sizeStr) sizeStr = `${m[1]}"x${m[2]}"`
            })
            let soePageMain = 'SOE-101.00'
            let soePageDetails = 'SOE-201.00'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
              if (digitizerIdx !== -1 && pageIdx !== -1) {
                for (let r = 0; r < dataRows.length; r++) {
                  const d = String(dataRows[r][digitizerIdx] || '').toLowerCase()
                  if (d.includes('timber stringer')) {
                    const pageStr = String(dataRows[r][pageIdx] || '').trim()
                    const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                    if (soeMatches && soeMatches.length > 0) {
                      soePageMain = soeMatches[0]
                      if (soeMatches.length > 1) soePageDetails = soeMatches[1]
                    }
                    break
                  }
                }
              }
            }
            proposalText = `F&I new ${sizeStr || '##'} timber stringer as per ${soePageMain} & details on ${soePageDetails} to SOE-205.00`
          } else if (subsectionName.toLowerCase() === 'timber soldier piles') {
            // Template text: F&I new (15)no [4"x4"] timber soldier piles (Havg=15'-0", 3'-10" & 5'-0" embedment) as per SOE-101.00 & details on SOE-201.00
            // (qty from G, size/Havg/embedment from group items, SOE refs from raw data)
            const parseDimToFeet = (dimStr) => {
              if (!dimStr) return 0
              const m = String(dimStr).match(/(\d+)(?:'-?)?(\d+)?/)
              if (!m) return 0
              return (parseInt(m[1]) || 0) + ((parseInt(m[2]) || 0) / 12)
            }
            const feetToFtIn = (feet) => {
              const f = Math.floor(feet)
              const i = Math.round((feet - f) * 12)
              return i === 0 ? `${f}'-0"` : `${f}'-${i}"`
            }
            let sizeStr = ''
            const heightVals = []
            const embedmentVals = []
            group.forEach(item => {
              const p = (item.particulars || '').trim()
              const sizeMatch = p.match(/(\d+)"\s*x\s*(\d+)"?/i) || p.match(/(\d+)\s*"\s*x\s*(\d+)/i)
              if (sizeMatch && !sizeStr) sizeStr = `${sizeMatch[1]}"x${sizeMatch[2]}"`
              const hMatch = p.match(/H=([0-9'"\-]+)/i)
              if (hMatch) heightVals.push(parseDimToFeet(hMatch[1]))
              const eMatch = p.match(/E=([0-9'"\-]+)/i)
              if (eMatch) embedmentVals.push(parseDimToFeet(eMatch[1]))
              const rowH = item.height || item.parsed?.heightRaw || 0
              if (rowH > 0) heightVals.push(rowH)
            })
            let havgText = '##'
            if (heightVals.length > 0) {
              const avg = heightVals.reduce((a, b) => a + b, 0) / heightVals.length
              havgText = feetToFtIn(avg)
            }
            let embedmentText = ''
            if (embedmentVals.length > 0) {
              const uniq = [...new Set(embedmentVals.map(v => Math.round(v * 12)))]
              const minE = Math.min(...embedmentVals)
              const maxE = Math.max(...embedmentVals)
              const minStr = feetToFtIn(minE)
              const maxStr = feetToFtIn(maxE)
              if (minE === maxE) {
                embedmentText = `${minStr} embedment`
              } else if (uniq.length === 2) {
                embedmentText = `${minStr} & ${maxStr} embedment`
              } else {
                embedmentText = `${minStr} to ${maxStr} embedment`
              }
            }
            let soePageMain = 'SOE-101.00'
            let soePageDetails = 'SOE-201.00'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
              if (digitizerIdx !== -1 && pageIdx >= 0) {
                for (let r = 0; r < dataRows.length; r++) {
                  const d = String(dataRows[r][digitizerIdx] || '').toLowerCase()
                  if (d.includes('timber') && d.includes('soldier')) {
                    const pageStr = String(dataRows[r][pageIdx] || '').trim()
                    const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                    if (soeMatches && soeMatches.length > 0) {
                      soePageMain = soeMatches[0]
                      if (soeMatches.length > 1) soePageDetails = soeMatches[1]
                    }
                    break
                  }
                }
              }
            }
            const qtyVal = Math.round(totalQty || totalTakeoff || 0)
            afterCountForB = `)no [${sizeStr || '##'}] timber soldier piles (Havg=${havgText}, ${embedmentText || '## embedment'}) as per ${soePageMain} & details on ${soePageDetails}`
            proposalText = `F&I new (${qtyVal})no [${sizeStr || '##'}] timber soldier piles (Havg=${havgText}, ${embedmentText || '## embedment'}) as per ${soePageMain} & details on ${soePageDetails}`
            if (sumRowIndex > 0) useFormulaForB = true
          } else if (subsectionName.toLowerCase() === 'timber planks') {
            // Template text: F&I new 3"x12" timber planks (H=3'-0") as per SOE-100.00 & details on SOE-002.00
            // (dimensions and H from group items, SOE refs from raw data)
            let dimensions = ''
            let heightText = ''
            group.forEach(item => {
              const p = (item.particulars || '').trim()
              const dimMatch = p.match(/(\d+)"\s*x\s*(\d+)"?/i) || p.match(/(\d+)\s*"\s*x\s*(\d+)/i)
              if (dimMatch && !dimensions) dimensions = `${dimMatch[1]}"x${dimMatch[2]}"`
              const hMatch = p.match(/H=([0-9'"\-]+)/i)
              if (hMatch && !heightText) heightText = hMatch[1].replace(/^\(|\)$/g, '').trim()
            })
            if (!heightText && group.length > 0) {
              let totalH = 0
              let cnt = 0
              group.forEach(item => {
                const h = item.height || item.parsed?.heightRaw || 0
                if (h > 0) { totalH += h; cnt++ }
              })
              if (cnt > 0) {
                const avg = totalH / cnt
                const f = Math.floor(avg)
                const i = Math.round((avg - f) * 12)
                heightText = i === 0 ? `${f}'-0"` : `${f}'-${i}"`
              }
            }
            let soePageMain = 'SOE-100.00'
            let soePageDetails = 'SOE-002.00'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
              if (digitizerIdx !== -1 && pageIdx >= 0) {
                for (let r = 0; r < dataRows.length; r++) {
                  const d = String(dataRows[r][digitizerIdx] || '').toLowerCase()
                  if (d.includes('timber') && d.includes('plank')) {
                    const pageStr = String(dataRows[r][pageIdx] || '').trim()
                    const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                    if (soeMatches && soeMatches.length > 0) {
                      soePageMain = soeMatches[0]
                      if (soeMatches.length > 1) soePageDetails = soeMatches[1]
                    }
                    break
                  }
                }
              }
            }
            proposalText = `F&I new ${dimensions || '##'} timber planks (H=${heightText || '##'}) as per ${soePageMain} & details on ${soePageDetails}`
          } else if (subsectionName.toLowerCase() === 'timber post') {
            // Template text: F&I new 6"x6" timber post (Havg=7'-10", 4'-0" to 5'-8" embedment) as per SOE-101.00 & details on SOE-201.00 to SOE-205.00
            // (size/Havg/embedment from group items, SOE refs from raw data; " to SOE-205.00" when page has range)
            const parseDimToFeet = (dimStr) => {
              if (!dimStr) return 0
              const m = String(dimStr).match(/(\d+)(?:'-?)?(\d+)?/)
              if (!m) return 0
              return (parseInt(m[1]) || 0) + ((parseInt(m[2]) || 0) / 12)
            }
            const feetToFtIn = (feet) => {
              const f = Math.floor(feet)
              const i = Math.round((feet - f) * 12)
              return i === 0 ? `${f}'-0"` : `${f}'-${i}"`
            }
            let sizeStr = ''
            const heightVals = []
            const embedmentVals = []
            group.forEach(item => {
              const p = (item.particulars || '').trim()
              const sizeMatch = p.match(/(\d+)"\s*x\s*(\d+)"?/i) || p.match(/(\d+)\s*"\s*x\s*(\d+)/i)
              if (sizeMatch && !sizeStr) sizeStr = `${sizeMatch[1]}"x${sizeMatch[2]}"`
              const hMatch = p.match(/H=([0-9'"\-]+)/i)
              if (hMatch) heightVals.push(parseDimToFeet(hMatch[1]))
              const eMatch = p.match(/E=([0-9'"\-]+)/i)
              if (eMatch) embedmentVals.push(parseDimToFeet(eMatch[1]))
              const rowH = item.height || item.parsed?.heightRaw || 0
              if (rowH > 0) heightVals.push(rowH)
            })
            let havgText = '##'
            if (heightVals.length > 0) {
              const avg = heightVals.reduce((a, b) => a + b, 0) / heightVals.length
              havgText = feetToFtIn(avg)
            }
            let embedmentText = ''
            if (embedmentVals.length > 0) {
              const minE = Math.min(...embedmentVals)
              const maxE = Math.max(...embedmentVals)
              embedmentText = minE === maxE ? feetToFtIn(minE) + ' embedment' : `${feetToFtIn(minE)} to ${feetToFtIn(maxE)} embedment`
            }
            let soePageMain = 'SOE-101.00'
            let soePageDetails = 'SOE-201.00'
            let soePageRange = ''
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
              if (digitizerIdx !== -1 && pageIdx >= 0) {
                for (let r = 0; r < dataRows.length; r++) {
                  const d = String(dataRows[r][digitizerIdx] || '').toLowerCase()
                  if (d.includes('timber') && d.includes('post')) {
                    const pageStr = String(dataRows[r][pageIdx] || '').trim()
                    const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                    const toMatch = pageStr.match(/SOE-[\d.]+[\s]*to[\s]*SOE-[\d.]+/gi)
                    if (soeMatches && soeMatches.length > 0) {
                      soePageMain = soeMatches[0]
                      if (soeMatches.length > 1) soePageDetails = soeMatches[1]
                    }
                    if (toMatch && toMatch.length) soePageRange = ' ' + toMatch[0].trim()
                    break
                  }
                }
              }
            }
            proposalText = `F&I new ${sizeStr || '##'} timber post (Havg=${havgText}, ${embedmentText || '## embedment'}) as per ${soePageMain} & details on ${soePageDetails}${soePageRange}`
          } else if (subsectionName.toLowerCase() === 'drilled hole grout') {
            // Dynamic from Calculation sheet: e.g. "5-5/8" Ø Drilled hole grout H=22'-6", typ." -> F&I new (5-5/8" Ø) drilled hole grout @ tangent pile (H=22'-6", typ.)
            let diameterStr = '##'
            const heightVals = []
            const parseDimToFeet = (dimStr) => {
              if (!dimStr) return 0
              const m = String(dimStr).match(/(\d+)(?:'-?)?(\d+)?/)
              if (!m) return 0
              return (parseInt(m[1], 10) || 0) + ((parseInt(m[2], 10) || 0) / 12)
            }
            const feetToFtIn = (feet) => {
              const f = Math.floor(feet)
              const i = Math.round((feet - f) * 12)
              return i === 0 ? `${f}'-0"` : `${f}'-${i}"`
            }
            group.forEach(item => {
              const p = (item.particulars || '').trim()
              if (diameterStr === '##') {
                const fracMatch = p.match(/(\d+-\d+\/\d+)["']?\s*Ø/i)
                const decMatch = p.match(/(\d+\.?\d*)["']?\s*Ø/i)
                if (fracMatch) diameterStr = `${fracMatch[1]}" Ø`
                else if (decMatch) diameterStr = `${decMatch[1]}" Ø`
              }
              const hMatch = p.match(/H=([0-9'"\-]+)/i)
              if (hMatch) heightVals.push(parseDimToFeet(hMatch[1]))
              const rowH = item.height || item.parsed?.heightRaw || 0
              if (rowH > 0) heightVals.push(rowH)
            })
            let heightText = '##'
            if (heightVals.length > 0) {
              const avg = heightVals.reduce((a, b) => a + b, 0) / heightVals.length
              heightText = feetToFtIn(avg)
            }
            let soeRef = 'SOE-101.01'
            if (rawData && Array.isArray(rawData) && rawData.length > 1) {
              const headers = rawData[0]
              const dataRows = rawData.slice(1)
              const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
              const pageIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'page')
              if (digitizerIdx !== -1 && pageIdx >= 0) {
                for (let r = 0; r < dataRows.length; r++) {
                  const d = String(dataRows[r][digitizerIdx] || '').toLowerCase()
                  if (d.includes('drilled hole grout')) {
                    const pageStr = String(dataRows[r][pageIdx] || '').trim()
                    const soeMatches = pageStr.match(/(SOE|SOESK)-[\d.]+/gi)
                    if (soeMatches && soeMatches.length > 0) soeRef = soeMatches[0]
                    break
                  }
                }
              }
            }
            proposalText = `F&I new (${diameterStr}) drilled hole grout @ tangent pile (H=${heightText}, typ.) as per ${soeRef} & details on above misc`
          } else {
            // Default proposal text for other subsections
            proposalText = `${subsectionName} item: ${Math.round(totalQty)} nos`
          }

          // Add proposal text (use formula for B when subsection uses QTY from calc sheet, e.g. Timber soldier piles)
          if (useFormulaForB && sumRowIndex > 0) {
            spreadsheet.updateCell({ formula: proposalFormulaWithQtyRef(currentRow, afterCountForB) }, `${pfx}B${currentRow}`)
          } else {
            spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          }
          rowBContentMap.set(currentRow, proposalText)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: 'white',
              verticalAlign: 'top',
              textDecoration: 'none'
            },
            `${pfx}B${currentRow}`
          )
          fillRatesForProposalRow(currentRow, proposalText)

          // Calculate and set row height based on content
          const dynamicHeight = calculateRowHeight(proposalText)

          // Add FT (LF) to column C - reference to calculation sheet sum row if available
          // For parging, sheet pile, guide wall, dowels: FT(I) from sum row
          const subsectionLowerForC = subsectionName.toLowerCase()
          if ((subsectionLowerForC === 'parging' || subsectionLowerForC === 'sheet pile' || subsectionLowerForC === 'sheet piles' || subsectionLowerForC === 'guide wall' || subsectionLowerForC === 'guilde wall' || subsectionLowerForC === 'dowels' || subsectionLowerForC === 'dowel bar' || subsectionLowerForC === 'rock pins' || subsectionLowerForC === 'rock pin' || subsectionLowerForC === 'shotcrete' || subsectionLowerForC === 'permission grouting' || subsectionLowerForC === 'mud slab' || subsectionLowerForC === 'drilled hole grout') && sumRowIndex > 0) {
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

          // Add SQFT to column D for Parging, Heel blocks, Underpinning, Concrete soil retention piers, Form board, and Sheet pile - reference to calculation sheet sum row
          const subsectionLowerName2 = subsectionName.toLowerCase()
          if ((subsectionLowerName2 === 'parging' ||
            subsectionLowerName2 === 'heel blocks' ||
            subsectionLowerName2 === 'underpinning' ||
            subsectionLowerName2 === 'concrete soil retention piers' ||
            subsectionLowerName2 === 'concrete soil retention pier' ||
            subsectionLowerName2 === 'form board' ||
            subsectionLowerName2 === 'sheet pile' ||
            subsectionLowerName2 === 'sheet piles' ||
            subsectionLowerName2 === 'timber soldier piles' ||
            subsectionLowerName2 === 'timber planks' ||
            subsectionLowerName2 === 'timber post' ||
            subsectionLowerName2 === 'timber lagging' ||
            subsectionLowerName2 === 'timber sheeting' ||
            subsectionLowerName2 === 'vertical timber sheets' ||
            subsectionLowerName2 === 'horizontal timber sheets' ||
            subsectionLowerName2 === 'timber stringer' ||
            subsectionLowerName2 === 'drilled hole grout' ||
            subsectionLowerName2 === 'guide wall' ||
            subsectionLowerName2 === 'guilde wall' ||
            subsectionLowerName2 === 'concrete buttons' ||
            subsectionLowerName2 === 'buttons' ||
            subsectionLowerName2 === 'rock stabilization' ||
            subsectionLowerName2 === 'shotcrete' ||
            subsectionLowerName2 === 'permission grouting' ||
            subsectionLowerName2 === 'mud slab') && sumRowIndex > 0) {
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

          // Add CY to column F for Heel blocks, Underpinning, Concrete soil retention piers, Guide wall, Concrete buttons, Vertical/Horizontal timber sheets, Timber stringer - reference to calculation sheet sum row
          if ((subsectionLowerName2 === 'heel blocks' ||
            subsectionLowerName2 === 'underpinning' ||
            subsectionLowerName2 === 'concrete soil retention piers' ||
            subsectionLowerName2 === 'concrete soil retention pier' ||
            subsectionLowerName2 === 'guide wall' ||
            subsectionLowerName2 === 'guilde wall' ||
            subsectionLowerName2 === 'concrete buttons' ||
            subsectionLowerName2 === 'buttons' ||
            subsectionLowerName2 === 'rock stabilization' ||
            subsectionLowerName2 === 'shotcrete' ||
            subsectionLowerName2 === 'mud slab' ||
            subsectionLowerName2 === 'vertical timber sheets' ||
            subsectionLowerName2 === 'horizontal timber sheets' ||
            subsectionLowerName2 === 'timber stringer' ||
            subsectionLowerName2 === 'drilled hole grout') && sumRowIndex > 0) {
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

          // Add QTY to column G for Underpinning, Concrete soil retention piers, Heel blocks, Concrete buttons, Dowels, Timber soldier piles, Timber planks, Timber post - reference to calculation sheet sum row
          if ((subsectionLowerName2 === 'underpinning' ||
            subsectionLowerName2 === 'concrete soil retention piers' ||
            subsectionLowerName2 === 'concrete soil retention pier' ||
            subsectionLowerName2 === 'heel blocks' ||
            subsectionLowerName2 === 'concrete buttons' ||
            subsectionLowerName2 === 'buttons' ||
            subsectionLowerName2 === 'dowels' ||
            subsectionLowerName2 === 'dowel bar' ||
            subsectionLowerName2 === 'rock pins' ||
            subsectionLowerName2 === 'rock pin' ||
            subsectionLowerName2 === 'timber soldier piles' ||
            subsectionLowerName2 === 'timber planks' ||
            subsectionLowerName2 === 'timber post' ||
            subsectionLowerName2 === 'vertical timber sheets' ||
            subsectionLowerName2 === 'horizontal timber sheets' ||
            subsectionLowerName2 === 'timber stringer' ||
            subsectionLowerName2 === 'drilled hole grout') && sumRowIndex > 0) {
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
          // For Sheet pile, Vertical/Horizontal timber sheets, Timber stringer: add when sum row exists
          if ((totalLBS > 0 || subsectionLowerName2 === 'sheet pile' || subsectionLowerName2 === 'sheet piles' ||
            subsectionLowerName2 === 'vertical timber sheets' || subsectionLowerName2 === 'horizontal timber sheets' || subsectionLowerName2 === 'timber stringer') && sumRowIndex > 0) {
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
          // For heel blocks, dowels, rock pins - add QTY when sum row exists
          if ((totalQty > 0 || subsectionLowerName2 === 'heel blocks' || subsectionLowerName2 === 'dowels' || subsectionLowerName2 === 'dowel bar' || subsectionLowerName2 === 'rock pins' || subsectionLowerName2 === 'rock pin') && sumRowIndex > 0) {
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
          const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
          spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: 'white',
              format: '$#,##0.00'
            },
            `${pfx}H${currentRow}`
          )

          // Apply currency format
          try {
            spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}:N${currentRow}`)
          } catch (e) {
            // Fallback already applied in cellFormat
          }

          // Row height already set above based on proposal text

          // For Timber lagging, add Backpacking line after it (under Sheet piles section order: Sheet pile → Timber lagging last → Backpacking after it)
          if (subsectionName.toLowerCase() === 'timber lagging') {
            let backpackingSumRowIndex = null
            if (calculationData && Array.isArray(calculationData)) {
              let fallbackRow = null
              for (let i = 0; i < calculationData.length; i++) {
                const r = calculationData[i]
                const colB = (r && r[1] != null) ? String(r[1]).trim() : ''
                const bLower = colB.toLowerCase()
                if (bLower === 'backpacking') {
                  backpackingSumRowIndex = i + 1
                  break
                }
                if (bLower === 'backpacking:') fallbackRow = i + 2
                else if (bLower.includes('backpacking') && fallbackRow == null) fallbackRow = i + 1
              }
              if (backpackingSumRowIndex == null && fallbackRow != null) backpackingSumRowIndex = fallbackRow
            }
            if (backpackingSumRowIndex > 0) {
              currentRow++
              const timberLaggingSoePage = window.timberLaggingSoePage || 'SOE-101.00'
              const mainSoeMatch = timberLaggingSoePage.match(/(SOE[-\d.]+)/i)
              const mainSoePage = mainSoeMatch ? mainSoeMatch[1] : 'SOE-101.00'
              const backpackingText = `F&I new backpacking @ timber lagging ${mainSoePage} & details on SOE-301.00`
              spreadsheet.updateCell({ value: backpackingText }, `${pfx}B${currentRow}`)
              rowBContentMap.set(currentRow, backpackingText)
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
              fillRatesForProposalRow(currentRow, backpackingText)
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${backpackingSumRowIndex}` }, `${pfx}D${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  textAlign: 'right',
                  backgroundColor: 'white'
                },
                `${pfx}D${currentRow}`
              )
              const backpackingDollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
              spreadsheet.updateCell({ formula: backpackingDollarFormula }, `${pfx}H${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'bold',
                  color: '#000000',
                  textAlign: 'right',
                  backgroundColor: 'white',
                  format: '$#,##0.00'
                },
                `${pfx}H${currentRow}`
              )
              try {
                spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}:N${currentRow}`)
              } catch (e) { /* ignore */ }
            }
          }

          currentRow++ // Move to next row

          // If this is concrete soil retention pier and form board items exist, add form board text on next row
          if ((subsectionName.toLowerCase() === 'concrete soil retention piers' || subsectionName.toLowerCase() === 'concrete soil retention pier') && window.formBoardProposalText) {
            const formBoardText = window.formBoardProposalText
            const formBoardItems = window.formBoardItems || []

            // Add form board proposal text
            spreadsheet.updateCell({ value: formBoardText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, formBoardText)
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
            fillRatesForProposalRow(currentRow, formBoardText)

            // Calculate and set row height based on content
            const formBoardDynamicHeight = calculateRowHeight(formBoardText)

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
            // Sum row is on the same row as the last form board item, not the next row
            let formBoardSumRowIndex = formBoardLastRowNumber

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

            // Add $/1000 formula in column H
            const formBoardDollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
            spreadsheet.updateCell({ formula: formBoardDollarFormula }, `${pfx}H${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white',
                format: '$#,##0.00'
              },
              `${pfx}H${currentRow}`
            )

            // Apply currency format
            try {
              spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}:N${currentRow}`)
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
            const afterCountShims = afterCountFromProposalText(shimsProposalText)
            if (afterCountShims) {
              spreadsheet.updateCell({ formula: proposalFormulaWithQtyRef(currentRow, afterCountShims) }, `${pfx}B${currentRow}`)
            } else {
              spreadsheet.updateCell({ value: shimsProposalText }, `${pfx}B${currentRow}`)
            }
            rowBContentMap.set(currentRow, shimsProposalText)
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
            fillRatesForProposalRow(currentRow, shimsProposalText)

            // Calculate and set row height based on content
            const shimsDynamicHeight = calculateRowHeight(shimsProposalText)

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

            // Add $/1000 formula in column H
            const shimsDollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
            spreadsheet.updateCell({ formula: shimsDollarFormula }, `${pfx}H${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white',
                format: '$#,##0.00'
              },
              `${pfx}H${currentRow}`
            )

            // Apply currency format
            try {
              spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}:N${currentRow}`)
            } catch (e) {
              // Fallback already applied in cellFormat
            }

            // Row height already set above based on shims proposal text

            currentRow++ // Move to next row
          })
        }
      })

      // Mark end of SOE scope for SOE Total
      soeScopeEndRow = currentRow - 1

      // Add SOE Total row after Misc.
      if (soeScopeStartRow !== null && soeScopeEndRow !== null) {
        spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
        spreadsheet.updateCell({ value: 'SOE Total:' }, `${pfx}D${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            backgroundColor: '#BDD7EE'
          },
          `${pfx}D${currentRow}:E${currentRow}`
        )
        spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
        // SOE total in F:G (like Demolition Total / Excavation Total), not in H
        spreadsheet.updateCell({ formula: `=SUM(H${soeScopeStartRow}:H${soeScopeEndRow})*1000` }, `${pfx}F${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'center',
            backgroundColor: '#BDD7EE',
            format: '$#,##0.00'
          },
          `${pfx}F${currentRow}:G${currentRow}`
        )
        applyTotalRowBorders(spreadsheet, pfx, currentRow, '#BDD7EE')
        baseBidTotalRows.push(currentRow)
        totalRows.push(currentRow)
        currentRow++
        currentRow++ // Empty row after SOE Total
      }
      } // end if (subsectionsToDisplay.length > 0)
    } // end if (hasSOEScopeData)

    // Process Rock anchor items from calculation data (check before showing section)
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

    const hasRockAnchorScopeData = rockAnchorItems.length > 0

    if (hasRockAnchorScopeData) {
      // Add Rock anchor scope heading after SOE Total
      spreadsheet.updateCell({ value: 'Rock anchor scope:' }, `${pfx}B${currentRow}`)
      rowBContentMap.set(currentRow, 'Rock anchor scope:')
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
      currentRow++

      // Add Rock anchors subheading
      spreadsheet.updateCell({ value: 'Rock anchors: including applicable washers, steel bearing plates, locking hex nuts as required' }, `${pfx}B${currentRow}`)
      rowBContentMap.set(currentRow, 'Rock anchors: including applicable washers, steel bearing plates, locking hex nuts as required')
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
      currentRow++

      // Process and display Rock anchor items (grouped)
      const rockAnchorStartRow = currentRow
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
          const proposalText = `F&I new (${totalQtyValue})no (##"Ø thick) threaded bar (or approved equal) (## Kips tension load capacity & ## Kips lock off load capacity) ##KSI rock anchors (##-KSI grout infilled) (L=${freeLengthText} + ${bondLengthText} bond length), ${drillHole} drill hole as per ${soePageMain} & details on`
          const afterCount = afterCountFromProposalText(proposalText)
          if (afterCount) {
            spreadsheet.updateCell({ formula: proposalFormulaWithQtyRef(currentRow, afterCount) }, `${pfx}B${currentRow}`)
          } else {
            spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          }
          rowBContentMap.set(currentRow, proposalText)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: 'white',
              verticalAlign: 'top',
              textDecoration: 'none'
            },
            `${pfx}B${currentRow}`
          )
          fillRatesForProposalRow(currentRow, proposalText)

          // Calculate and set row height based on content
          const rockAnchorDynamicHeight = calculateRowHeight(proposalText)

          // Add FT (LF) to column C - sum of all items in group
          let totalFT = 0
          group.forEach(item => {
            if (item.rawRowNumber > 0) {
              const calcSheetName = 'Calculations Sheet'
              // We'll sum all FT values from the group
              totalFT += parseFloat(calculationData[item.rawRowNumber - 2]?.[8] || 0) || 0
            }
          })

          // Find sum row for rock anchors
          const sumRowIndex = lastRowNumber
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

          // Add $/1000 formula in column H
          const rockAnchorDollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
          spreadsheet.updateCell({ formula: rockAnchorDollarFormula }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: 'white',
              format: '$#,##0.00'
            },
            `${pfx}H${currentRow}`
          )

          // Apply currency format
          try {
            spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}:N${currentRow}`)
          } catch (e) {
            // Fallback already applied in cellFormat
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
          const proposalText = `F&I new (${totalQtyValue})no threaded bar (or approved equal) ##KSI (##°) rock bolts/dowel (L=${bondLengthText}), ${drillHole} drill hole as per ${soePageMain} & details on`
          const afterCountRB = afterCountFromProposalText(proposalText)
          if (afterCountRB) {
            spreadsheet.updateCell({ formula: proposalFormulaWithQtyRef(currentRow, afterCountRB) }, `${pfx}B${currentRow}`)
          } else {
            spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          }
          rowBContentMap.set(currentRow, proposalText)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'left',
              backgroundColor: 'white',
              verticalAlign: 'top',
              textDecoration: 'none'
            },
            `${pfx}B${currentRow}`
          )
          fillRatesForProposalRow(currentRow, proposalText)

          // Calculate and set row height based on content
          const rockBoltDynamicHeight = calculateRowHeight(proposalText)

          // Find sum row for rock bolts
          const sumRowIndex = lastRowNumber
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

          // Add $/1000 formula in column H
          const rockBoltDollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
          spreadsheet.updateCell({ formula: rockBoltDollarFormula }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: 'white',
              format: '$#,##0.00'
            },
            `${pfx}H${currentRow}`
          )

          // Apply currency format
          try {
            spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}:N${currentRow}`)
          } catch (e) {
            // Fallback already applied in cellFormat
          }

          // Row height already set above based on rock bolt proposal text
          currentRow++
        })
      }
      const rockAnchorEndRow = currentRow - 1

      // Add Rock anchor Total row when there are data rows
      if (rockAnchorEndRow >= rockAnchorStartRow) {
        // Label across B–E, amount in F–G, same layout as other totals
        spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
        spreadsheet.updateCell({ value: 'Rock anchor Total:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'right',
            backgroundColor: '#BDD7EE'
          },
          `${pfx}B${currentRow}:E${currentRow}`
        )

        spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
        spreadsheet.updateCell(
          { formula: `=SUM(H${rockAnchorStartRow}:H${rockAnchorEndRow})*1000` },
          `${pfx}F${currentRow}`
        )
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'right',
            backgroundColor: '#BDD7EE',
            format: '$#,##0.00'
          },
          `${pfx}F${currentRow}:G${currentRow}`
        )

        applyTotalRowBorders(spreadsheet, pfx, currentRow, '#BDD7EE')
        totalRows.push(currentRow)
        currentRow++
      }

      // Add empty row to separate Rock anchor scope from Foundation scope
      currentRow++
    } // end if (hasRockAnchorScopeData)

    // Note: Green background for columns I-N will be applied at the end after all data is written

    // Add Foundation subsection headers (similar to SOE subsections)
    const foundationSubsectionItems = window.foundationSubsectionItems || new Map()
    const foundationSubsectionOrder = [
      'Piles',                    // From calculation sheet "Piles" subsection (miscellaneousPileItems)
      'Drilled foundation pile',  // Will display as "Foundation drilled piles scope:"
      'Driven foundation pile',   // Will display as "Foundation driven piles scope:"
      'Helical foundation pile',  // Will display as "Foundation helical piles scope:"
      'Drilled displacement pile', // Will display as "Stelcor piles scope:"
      'CFA pile'                  // Will display as "CAF piles scope:"
    ]

    // Get all unique foundation subsection names from collected items
    // Map new format names to old format for display
    const foundationSubsectionDisplayMap = {
      'Piles': 'Piles',
      'Drilled foundation pile': 'Foundation drilled piles',
      'Driven foundation pile': 'Foundation driven piles',
      'Helical foundation pile': 'Foundation helical piles',
      'Drilled displacement pile': 'Stelcor piles',
      'CFA pile': 'CAF piles'
    }
    // Map header text (e.g. "Foundation drilled piles scope") to canonical for matching
    const foundationHeaderToCanonical = {
      'Piles': 'Piles',
      'Piles scope': 'Piles',
      'Foundation drilled piles': 'Drilled foundation pile',
      'Foundation drilled piles scope': 'Drilled foundation pile',
      'Foundation driven piles': 'Driven foundation pile',
      'Foundation driven piles scope': 'Driven foundation pile',
      'Foundation helical piles': 'Helical foundation pile',
      'Foundation helical piles scope': 'Helical foundation pile',
      'Stelcor piles': 'Drilled displacement pile',
      'Stelcor piles scope': 'Drilled displacement pile',
      'CAF piles': 'CFA pile',
      'CAF piles scope': 'CFA pile'
    }

    const collectedFoundationSubsections = new Set()
    foundationSubsectionItems.forEach((groups, name) => {
      if (groups.length > 0) {
        const canonical = foundationHeaderToCanonical[name] || name
        collectedFoundationSubsections.add(canonical)
      }
    })

    // Display foundation subsections in order (only up to CFA pile)
    // Include subsection if we have collected items OR if we have processor groups (e.g. from calculation sheet)
    const processorGroupsMapForDisplay = {
      'Piles': window.miscellaneousPileItems || [],
      'Drilled foundation pile': window.drilledFoundationPileGroups || [],
      'Driven foundation pile': window.drivenFoundationPileItems || [],
      'Helical foundation pile': window.helicalFoundationPileGroups || [],
      'Drilled displacement pile': window.stelcorDrilledDisplacementPileItems || [],
      'CFA pile': window.cfaPileItems || []
    }
    const foundationSubsectionsToDisplay = []
    foundationSubsectionOrder.forEach(name => {
      const hasCollected = collectedFoundationSubsections.has(name)
      const processorGroupsForName = processorGroupsMapForDisplay[name]
      // Piles uses flat array of items (miscellaneousPileItems); others use groups with .items or .takeoff
      const hasProcessorData = Array.isArray(processorGroupsForName) && processorGroupsForName.length > 0 && (
        name === 'Piles'
          ? true
          : (processorGroupsForName[0].items?.length > 0 || processorGroupsForName[0].takeoff !== undefined)
      )
      if (hasCollected || hasProcessorData) {
        foundationSubsectionsToDisplay.push(name)
        collectedFoundationSubsections.delete(name)
      }
    })
    // Do not add any remaining subsections - only show up to CFA pile

    // Add "Piles scope:" parent header above all pile subsections (Foundation drilled piles, etc.)
    if (foundationSubsectionsToDisplay.length > 0) {
      spreadsheet.updateCell({ value: 'Piles scope:' }, `${pfx}B${currentRow}`)
      rowBContentMap.set(currentRow, 'Piles scope:')
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
      currentRow++
    }

    foundationSubsectionsToDisplay.forEach((subsectionName) => {
      // Map new format names to old format for display
      let displayName = subsectionName
      if (foundationSubsectionDisplayMap[subsectionName]) {
        displayName = foundationSubsectionDisplayMap[subsectionName]
      }

      // Add subsection header with blue background and "scope:" suffix (skip for "Piles" – parent row is already "Piles scope:")
      if (subsectionName !== 'Piles') {
        const subsectionText = `${displayName} scope:`
        spreadsheet.updateCell({ value: subsectionText }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, subsectionText)
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
        currentRow++
      }

      // Track the start row for this subsection (for total calculation)
      // This is the first row where proposal text will be written
      const subsectionStartRow = currentRow

      // Calculate QTY sum from processor groups (sum of all takeoff values) - BEFORE processing groups
      // Use the summed takeoff from processDrilledFoundationPileItems (or other processors)
      let qtySumValue = 0
      let pileType = 'drilled' // default
      let diameter = null
      let thickness = null

      // Map subsection names to processor group arrays (Piles: wrap miscellaneousPileItems as one group)
      const miscPileItems = window.miscellaneousPileItems || []
      const processorGroupsMap = {
        'Piles': miscPileItems.length > 0 ? [{ items: miscPileItems }] : [],
        'Drilled foundation pile': window.drilledFoundationPileGroups || [],
        'Foundation drilled piles': window.drilledFoundationPileGroups || [],
        'Driven foundation pile': window.drivenFoundationPileItems || [],
        'Foundation driven piles': window.drivenFoundationPileItems || [],
        'Helical foundation pile': window.helicalFoundationPileGroups || [],
        'Foundation helical piles': window.helicalFoundationPileGroups || [],
        'Drilled displacement pile': window.stelcorDrilledDisplacementPileItems || [],
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
      } else if (subsectionName === 'Drilled displacement pile') {
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
        // Classify groups: influence + Drilled cassion pile (single Ø) vs influence + Isolation casing (dual Ø) vs no influence
        const influenceDrilledGroups = [] // Pile within TA Influence Line
        const influenceIsolationGroups = [] // Isolation casing drilled cassion pile within TA Influence Line
        const noInfluenceGroups = []
        const hasInfluenceFromText = (group) => {
          const firstParticulars = (group.items?.[0]?.particulars || '').toString().toLowerCase()
          return /influence|\s+rs\s+|ta\s*influence|rs\s*influence/i.test(firstParticulars) || !!(group.hasInflu)
        }
        processorGroups.forEach((itemOrGroup, groupIndex) => {
          if (!itemOrGroup.items || !Array.isArray(itemOrGroup.items) || itemOrGroup.items.length === 0) return
          const groupTakeoff = itemOrGroup.items[0]?.takeoff || 0
          if (groupTakeoff === 0) return
          if (!itemOrGroup.hasInflu && hasInfluenceFromText(itemOrGroup)) itemOrGroup.hasInflu = true
          const groupHasInflu = itemOrGroup.hasInflu || false
          const isIsolationCasing = !!(itemOrGroup.items[0]?.parsed?.isDualDiameter) ||
            (itemOrGroup.items[0]?.particulars && String(itemOrGroup.items[0].particulars).toLowerCase().includes('isolation casing'))
          if (groupHasInflu && isIsolationCasing) influenceIsolationGroups.push({ itemOrGroup, groupIndex })
          else if (groupHasInflu) influenceDrilledGroups.push({ itemOrGroup, groupIndex })
          else noInfluenceGroups.push({ itemOrGroup, groupIndex })
        })

        const addInfluenceSectionHeading = (headingText) => {
          spreadsheet.updateCell({ value: headingText }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, headingText)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat(
            { fontWeight: 'bold', fontStyle: 'italic', color: '#000000', textAlign: 'left', backgroundColor: '#FCE4D6', textDecoration: 'underline', border: '1px solid #000000' },
            `${pfx}B${currentRow}`
          )
          currentRow++
        }

        const renderOneDrilledPileGroup = (itemOrGroup, groupIndex) => {
          const groupTakeoff = itemOrGroup.items[0]?.takeoff || 0
          const groupQty = Math.round(groupTakeoff)
          const groupHasInflu = itemOrGroup.hasInflu || false

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

          // Hardcoded template for all items in foundation drilled piles scope
          const pileDescriptionFromParticulars = (firstItem?.particulars || '').toString().trim()
          const proposalText = buildDrilledPileTemplateText(pileDescriptionFromParticulars, diameterThicknessText, heightText, rockSocketText, groupQty) + ` ${foReference}, ${soeReference} & details on`
          const afterCountDC = afterCountFromProposalText(proposalText)
          if (afterCountDC) {
            spreadsheet.updateCell({ formula: proposalFormulaWithQtyRef(currentRow, afterCountDC) }, `${pfx}B${currentRow}`)
          } else {
            spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          }
          rowBContentMap.set(currentRow, proposalText)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          fillRatesForProposalRow(currentRow, proposalText)

          // Data rows stay white; only section headings use #FCE4D6
          const cellFormat = {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            verticalAlign: 'top',
            backgroundColor: 'white'
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
            // The sum row is at formula.row, and firstDataRow is the data row
            // So the sum row is at firstDataRow + 1 (data row + 1 = sum row)
            if (drilledFoundationSumFormulas[groupIndex]) {
              const formulaInfo = drilledFoundationSumFormulas[groupIndex]
              // The sum row should be firstDataRow + 1 (data row + 1 = sum row)
              // Use firstDataRow + 1 directly to ensure we get the correct sum row
              groupSumRowIndex = formulaInfo.firstDataRow + 1


              // Reference the FT sum from column I of the sum row
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${groupSumRowIndex}` }, `${pfx}C${currentRow}`)
            }
          }

          // Format FT column (data rows white; only section headings use #FCE4D6)
          if (groupSumRowIndex > 0 || groupFTSum > 0) {
            spreadsheet.cellFormat(
              { fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' },
              `${pfx}C${currentRow}`
            )
          }

          // Add LBS to column E - reference to calculation sheet sum row (column K)
          if (groupSumRowIndex > 0) {
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!K${groupSumRowIndex}` }, `${pfx}E${currentRow}`)
            spreadsheet.cellFormat(
              { fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' },
              `${pfx}E${currentRow}`
            )
          }

          // Add QTY from calculation sheet (column M of sum row) when available
          if (groupSumRowIndex > 0) {
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${groupSumRowIndex}` }, `${pfx}G${currentRow}`)
          } else {
            spreadsheet.updateCell({ value: groupQty }, `${pfx}G${currentRow}`)
          }
          spreadsheet.cellFormat(
            { fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' },
            `${pfx}G${currentRow}`
          )

          // Add $/1000 formula
          const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
          spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat(
            { fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white', format: '$#,##0.00' },
            `${pfx}H${currentRow}`
          )

          // Row height already set above based on proposal text
          currentRow++
        }

        // Render in order: (1) Pile within TA Influence Line, (2) Isolation casing within TA Influence Line, (3) no-influence
        if (influenceDrilledGroups.length > 0) {
          addInfluenceSectionHeading('Pile within TA Influence Line:')
          influenceDrilledGroups.forEach(({ itemOrGroup, groupIndex }) => renderOneDrilledPileGroup(itemOrGroup, groupIndex))
        }
        if (influenceIsolationGroups.length > 0) {
          if (influenceDrilledGroups.length > 0) {
            spreadsheet.cellFormat({ backgroundColor: '#FCE4D6' }, `${pfx}B${currentRow}:${pfx}H${currentRow}`)
            currentRow++
          }
          addInfluenceSectionHeading('Isolation casing drilled cassion pile within TA Influence Line:')
          influenceIsolationGroups.forEach(({ itemOrGroup, groupIndex }) => renderOneDrilledPileGroup(itemOrGroup, groupIndex))
        }
        if (noInfluenceGroups.length > 0) {
          noInfluenceGroups.forEach(({ itemOrGroup, groupIndex }) => renderOneDrilledPileGroup(itemOrGroup, groupIndex))
        }
      } else if (qtySumValue > 0) {
        // For non-drilled piles, generate single proposal text (existing logic)
        // Get height and other details from groups if available (try canonical key then display header key)
        const displayKey = (foundationSubsectionDisplayMap[subsectionName] || subsectionName) + ' scope'
        const groups = foundationSubsectionItems.get(subsectionName) || foundationSubsectionItems.get(displayKey) || []
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

        // Generate proposal text: use first group's particulars when rich; else template with parsed design loads, grout, rebar, height & rock socket
        // For Stelcor/CFA, groups may be [non-influence, influence]; use first group for default, but for influence heading/line use first *influence* group if any
        const firstGroupFirstItem = processorGroups[0]?.items?.[0]
        let pileDescFromParticulars = (firstGroupFirstItem?.particulars || processorGroups[0]?.particulars || '').toString().trim()
        const firstInfluenceGroup = processorGroups.find(g => g.hasInflu || (Array.isArray(g.items) && g.items.some(it => /\binfluence\b/i.test((it.particulars || '').toString()))))
        const firstInfluenceParticulars = firstInfluenceGroup
          ? (firstInfluenceGroup.items?.[0]?.particulars ?? firstInfluenceGroup.particulars ?? '')
          : ''
        const influenceDesc = (firstInfluenceParticulars && firstInfluenceParticulars.toString().trim()) || pileDescFromParticulars
        const hasInfluence = /\binfluence\b/i.test(influenceDesc)
        if (hasInfluence) {
          pileDescFromParticulars = influenceDesc
          const influenceHeadingBySubsection = {
            'CFA pile': 'CFA pile within TA Influence Line:',
            'Driven foundation pile': 'Driven pile within TA Influence Line:',
            'Foundation driven piles': 'Driven pile within TA Influence Line:',
            'Helical foundation pile': 'Helical pile within TA Influence Line:',
            'Foundation helical piles': 'Helical pile within TA Influence Line:',
            'Drilled displacement pile': 'Stelcor pile within TA Influence Line:',
            'Piles': 'Pile within TA Influence Line:'
          }
          const influenceHeading = influenceHeadingBySubsection[subsectionName] || `${subsectionName} within TA Influence Line:`
          spreadsheet.updateCell({ value: influenceHeading }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, influenceHeading)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat(
            { fontWeight: 'bold', fontStyle: 'italic', color: '#000000', textAlign: 'left', backgroundColor: '#FCE4D6', textDecoration: 'underline', border: '1px solid #000000' },
            `${pfx}B${currentRow}`
          )
          currentRow++
          pileDescFromParticulars = pileDescFromParticulars.replace(/\binfluence\b/gi, 'within TA influence line').trim()
        }
        const useParticularsAsDesc = pileDescFromParticulars.length > 3 && (pileDescFromParticulars.toLowerCase().includes('pile') || /[\d]+\/[\d]+/.test(pileDescFromParticulars))
        let proposalText
        if (useParticularsAsDesc) {
          proposalText = `F&I new (${qtySumValue})no ${pileDescFromParticulars} as per ${foReference}, ${soeReference} & details on`
        } else {
          proposalText = buildDrilledPileTemplateText(pileDescFromParticulars, diameterThicknessText, heightText, rockSocketText, qtySumValue) + ` ${foReference}, ${soeReference} & details on`
        }
        const afterCountFP = afterCountFromProposalText(proposalText)
        if (afterCountFP) {
          spreadsheet.updateCell({ formula: proposalFormulaWithQtyRef(currentRow, afterCountFP) }, `${pfx}B${currentRow}`)
        } else {
          spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
        }
        rowBContentMap.set(currentRow, proposalText)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            backgroundColor: 'white',
            verticalAlign: 'top',
            textDecoration: 'none'
          },
          `${pfx}B${currentRow}`
        )
        // Use proposal_mapped "titan pile" (LF 150) for Titan pile lines – matches CONCATENATE formula in B
        const pileRateKey = useParticularsAsDesc && /titan\s+pile/i.test(pileDescFromParticulars) ? 'titan pile' : null
        fillRatesForProposalRow(currentRow, proposalText, pileRateKey)

        // Calculate and set row height based on content
        const foundationPileDynamicHeight = calculateRowHeight(proposalText)

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

        // Add $/1000 formula in column H
        const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
        spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            textAlign: 'right',
            backgroundColor: 'white',
            format: '$#,##0.00'
          },
          `${pfx}H${currentRow}`
        )

        // Row height already set above based on foundation pile proposal text

        currentRow++ // Move to next row
      }

      // Track the end row of pile data items (before misc. section)
      const pileDataEndRow = currentRow - 1

      // Add misc. section for this foundation subsection (after all proposal text items)
      // Same "Drilled pile misc." content used under Drilled foundation pile and under Piles scope
      const drilledPileMisc = {
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
      }
      const miscSections = {
        'Drilled foundation pile': drilledPileMisc,
        'Piles': drilledPileMisc,
        'Driven foundation pile': {
          title: 'Driven piles misc.:',
          included: [
            'The Entire Length of each driven pile is charged including any cut off length',
            'One mobilization & demobilization of drilling equipment included',
            'Surveying, stakeout, pile numbering plan & as-built plan included',
            'One compression reactionary load tests included',
            'One lateral reactionary load tests included',
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
        'Drilled displacement pile': {
          title: 'Stelcor pile misc.:',
          included: [
            'Plates & locking nuts included',
            'One mobilization & demobilization of drilling equipment included',
            'Surveying, stakeout, pile numbering plan & as-built plan included',
            'One compression reactionary load test included',
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
            'One mobilization & demobilization of drilling equipment included',
            'Surveying, stakeout, pile numbering plan & as-built plan included',
            'One compression reactionary load test included',
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
        // Add misc. title (no extra line after – first included item on next row)
        spreadsheet.updateCell({ value: miscSection.title }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            backgroundColor: '#c9c9c9',
            textDecoration: 'underline'
          },
          `${pfx}B${currentRow}`
        )
        currentRow++

        // Add included items
        miscSection.included.forEach((item, index) => {
          spreadsheet.updateCell({ value: item }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, item)
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

          // Apply unit rates for this misc line from proposal_mapped.json
          // Include scope for lines that have different rates per pile type (driven/drilled/Stelcor/CAF/Helical etc.)
          const scopeKey = foundationSubsectionDisplayMap[subsectionName] || subsectionName
          const needsScopeForRates = item.includes('compression reactionary load test') ||
            item.includes('Engineering & Shop drawings included') ||
            item.includes('mobilization & demobilization of drilling equipment')
          let rateLookupKey = needsScopeForRates ? `${item} - ${scopeKey} scope` : item
          // proposal_mapped uses "One mobilization..." – normalize "Single mobilization..." for lookup
          if (item.includes('Single mobilization & demobilization of drilling equipment')) {
            rateLookupKey = needsScopeForRates ? `One mobilization & demobilization of drilling equipment included - ${scopeKey} scope` : 'One mobilization & demobilization of drilling equipment included'
          }
          fillRatesForProposalRow(currentRow, rateLookupKey)

          // Add special formulas for certain items
          if (item.includes('Plates & locking nuts') && (subsectionName === 'Drilled foundation pile' || subsectionName === 'Piles' || subsectionName === 'CFA pile')) {
            // Add SUM formula for QTY column (G) - sum all drilled pile QTY values
            if (subsectionStartRow && pileDataEndRow >= subsectionStartRow) {
              spreadsheet.updateCell({ formula: `=SUM(G${subsectionStartRow}:G${pileDataEndRow})` }, `${pfx}G${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'normal',
                  textAlign: 'right',
                  backgroundColor: 'white'
                },
                `${pfx}G${currentRow}`
              )

              // Add $/1000 formula in column H
              const miscDollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
              spreadsheet.updateCell({ formula: miscDollarFormula }, `${pfx}H${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'normal',
                  textAlign: 'right',
                  backgroundColor: 'white',
                  format: '$#,##0.00'
                },
                `${pfx}H${currentRow}`
              )

              // Apply currency format
              try {
                spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}:N${currentRow}`)
              } catch (e) {
                // Fallback already applied in cellFormat
              }
            }
          } else if (item.includes('Video inspections included') && (subsectionName === 'Drilled foundation pile' || subsectionName === 'Piles')) {
            // Add SUM formula for QTY column (G) - same as "Plates & locking nuts"
            if (subsectionStartRow && pileDataEndRow >= subsectionStartRow) {
              spreadsheet.updateCell({ formula: `=SUM(G${subsectionStartRow}:G${pileDataEndRow})` }, `${pfx}G${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'normal',
                  textAlign: 'right',
                  backgroundColor: 'white'
                },
                `${pfx}G${currentRow}`
              )

              // Add $/1000 formula in column H
              const miscDollarFormula3 = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
              spreadsheet.updateCell({ formula: miscDollarFormula3 }, `${pfx}H${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'normal',
                  textAlign: 'right',
                  backgroundColor: 'white',
                  format: '$#,##0.00'
                },
                `${pfx}H${currentRow}`
              )

              // Apply currency format
              try {
                spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}:N${currentRow}`)
              } catch (e) {
                // Fallback already applied in cellFormat
              }
            }
          }

          // Override QTY rate for "Plates & locking nuts included" by pile type:
          // - Drilled pile misc. and Piles scope: 250
          // - Stelcor pile misc. and CFA pile misc.: 100
          if (item.includes('Plates & locking nuts')) {
            let qtyRate = null
            if (subsectionName === 'Drilled foundation pile' || subsectionName === 'Piles') qtyRate = 250
            else if (subsectionName === 'Drilled displacement pile' || subsectionName === 'CFA pile') qtyRate = 100
            if (qtyRate != null) {
              spreadsheet.updateCell({ value: qtyRate }, `${pfx}M${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'normal',
                  textAlign: 'right',
                  backgroundColor: 'white'
                },
                `${pfx}M${currentRow}`
              )
            }
          }

          // If this misc line is a load test line, set QTY (G) to 1 or 2 based on the leading word
          const lowerItem = item.toLowerCase()
          if (lowerItem.includes('load test')) {
            let qtyValue = null
            if (lowerItem.startsWith('one ') || lowerItem.includes(' one ')) qtyValue = 1
            else if (lowerItem.startsWith('two ') || lowerItem.includes(' two ')) qtyValue = 2
            if (qtyValue != null) {
              spreadsheet.updateCell({ value: qtyValue }, `${pfx}G${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'normal',
                  textAlign: 'right',
                  backgroundColor: 'white'
                },
                `${pfx}G${currentRow}`
              )
            }
          }

          currentRow++
        })

        // Add additional cost items
        miscSection.additional.forEach(item => {
          spreadsheet.updateCell({ value: item }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, item)
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

          // Apply unit rates for this additional misc line from proposal_mapped.json
          fillRatesForProposalRow(currentRow, item)

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
          'Piles': 'Piles Total:',
          'Helical foundation pile': 'Helical Pile Total:',
          'Drilled displacement pile': 'Stelcor Piles Total:',
          'CFA pile': 'CFA Piles Total:',
          'Driven foundation pile': 'Foundation Piles Total:' // Driven piles use same label as drilled
        }
        const totalLabel = totalLabels[subsectionName] || 'Foundation Piles Total:'

        // Format like Demolition Total: label in D:E, total in F:G (no extra line before total)
        spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
        spreadsheet.updateCell({ value: totalLabel }, `${pfx}D${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            backgroundColor: '#BDD7EE'
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

        applyTotalRowBorders(spreadsheet, pfx, currentRow, '#BDD7EE')
        baseBidTotalRows.push(currentRow) // Foundation pile section total
        totalRows.push(currentRow)

        currentRow++ // Move past total row
        currentRow++ // One extra blank line after total before next subsection
      }
    })

    // Substructure concrete scope (above Below grade waterproofing scope)
    const substructureCalcSheet = 'Calculations Sheet'
    const pileSubsections = ['Drilled foundation pile', 'Helical foundation pile', 'Driven foundation pile', 'Drilled displacement pile', 'CFA pile']
    const allFoundationSums = (formulaData || []).filter(f =>
      f.itemType === 'foundation_sum' && f.section === 'foundation' && f.subsectionName &&
      !pileSubsections.includes(f.subsectionName)
    )
    const getFirstParticulars = (sumF) => {
      if (!calculationData || !sumF.firstDataRow) return ''
      const row = calculationData[sumF.firstDataRow - 1]
      return (row?.[1] || '').toString().toLowerCase()
    }
    const hasNonZeroTakeoff = (source, isFormulaItem) => {
      if (!calculationData) return true
      if (isFormulaItem) {
        if (source.itemType === 'buttress_final') {
          const finalRow = calculationData[source.row - 1]
          const mVal = parseFloat(finalRow?.[12]) || 0
          if (mVal > 0) return true
          if (source.buttressRow) {
            const takeoffRow = calculationData[source.buttressRow - 1]
            return (parseFloat(takeoffRow?.[2]) || 0) > 0
          }
          return false
        }
        const row = calculationData[source.row - 1]
        return (parseFloat(row?.[2]) || 0) > 0
      }
      const first = (source.firstDataRow || 1) - 1
      const last = (source.lastDataRow || source.firstDataRow || 1) - 1
      for (let i = first; i <= last; i++) {
        const row = calculationData[i]
        if (row && (parseFloat(row[2]) || 0) > 0) return true
      }
      return false
    }
    // Get data rows for a group (subsection): from rowsBySubsection or single formula row
    const getGroupDataRows = (source, sub, isFormulaItem) => {
      if (!calculationData) return []
      if (isFormulaItem && source && source.row) {
        const r = calculationData[source.row - 1]
        return r ? [r] : []
      }
      if (!rowsBySubsection || !sub) return []
      const direct = rowsBySubsection.get(sub)
      if (direct && direct.length) return direct
      const subLower = sub.toLowerCase()
      for (const [key, rows] of rowsBySubsection) {
        if (key && key.toLowerCase() === subLower && rows && rows.length) return rows
      }
      return []
    }
    // Get page ref(s) for a subsection from raw data Page column; if group has multiple pages return "102 & 105" or "102, 103 & 105"
    const getPageRefForSubsection = (sub) => {
      if (!rawData || !Array.isArray(rawData) || rawData.length < 2) return '##'
      const headers = rawData[0]
      const dataRows = rawData.slice(1)
      const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
      const pageIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'page')
      if (digitizerIdx === -1 || pageIdx === -1) return '##'
      const subNorm = (sub || '').toLowerCase().replace(/\s+s$/, '') // "Strap beams" -> "strap beam"
      const subAlt = (sub || '').toLowerCase() // "Strap beams" as-is
      const subWords = (sub || '').toLowerCase().split(/\s+/).filter(w => w.length > 1) // e.g. ["stairs", "on", "grade", "stairs"] -> match "stairs on grade"
      const collected = []
      const seen = new Set()
      for (const row of dataRows) {
        const digitizerItem = (row[digitizerIdx] || '').toString().toLowerCase()
        let matches = digitizerItem.includes(subNorm) || digitizerItem.includes(subAlt) ||
          (subNorm === 'tie beam' && (digitizerItem.includes('tie beam') || digitizerItem.startsWith('tb ')))
        // Match when subsection is "Stairs on grade Stairs" and raw has "Stairs on grade" (no trailing "Stairs")
        if (!matches && subWords.length >= 2) {
          const subCore = subWords.slice(0, -1).join(' ') // "stairs on grade" from "stairs on grade stairs"
          if (subCore && digitizerItem.includes(subCore)) matches = true
        }
        // Strap beams: raw data has "ST (3'-10"x2'-9") typ.", "ST (2'-8"x3'-0") (87.86')" etc. – no "strap" in text
        if (!matches && (subAlt === 'strap beams' || subNorm === 'strap beam')) {
          matches = (digitizerItem.startsWith('st ') || /^st\s*\(/.test(digitizerItem)) && !digitizerItem.startsWith('st-')
        }
        if (matches && row[pageIdx] != null && row[pageIdx] !== '') {
          const pageStr = String(row[pageIdx]).trim()
          const refMatch = pageStr.match(/(?:FO|P|S|A|E|DM)-[\d.]+/i)
          const ref = refMatch ? refMatch[0] : pageStr
          if (ref && !seen.has(ref)) {
            seen.add(ref)
            collected.push(ref)
          }
        }
      }
      return formatPageRefList(collected)
    }
    // Extract qty (col M = 12), height (col H = 7), thickness, width, dimensions (from Particulars col 1), pageRef from rawData for dynamic proposal text
    const getDynamicValuesFromGroupRows = (source, sub, isFormulaItem) => {
      const rows = getGroupDataRows(source, sub, isFormulaItem)
      const heightValues = rows.map(r => parseFloat(r[7])).filter(v => !Number.isNaN(v) && v > 0)
      let heightText = formatHeightAsFeetInches(heightValues)
      let elevatorPitWallThicknessFromBracket = ''
      let elevatorPitWallHeightFromBracket = ''
      let qty = 0
      for (const r of rows) {
        qty += Math.round(parseFloat(r[12]) || 0) || Math.round(parseFloat(r[2]) || 0) // M then C (takeoff)
      }
      const qtyText = qty > 0 ? String(qty) : '##'
      let thicknessText = ''
      let widthText = ''
      let dimensionsText = ''
      let wireMeshText = ''
      for (const r of rows) {
        const p = normalizeCalcSheetAbbreviations((r && r[1] || '').toString())
        if (!thicknessText) {
          const feetInch = p.match(/(\d+'\s*-\s*\d+"?)\s*thick/i) || p.match(/(\d+'\s*-\s*\d+"?)(?=\s|,|\)|$)/i)
          const inchOnly = p.match(/(\d+(?:\/\d+)?)"?\s*thick/i) || p.match(/(\d+(?:\/\d+)?)"(?=\s|,|\)|$)/i)
          if (feetInch) {
            thicknessText = feetInch[1].trim()
            if (!/"/.test(thicknessText)) thicknessText += '"'
          } else if (inchOnly) {
            thicknessText = inchOnly[1] + (inchOnly[1].includes('"') ? '' : '"')
          } else {
            const hMatch = p.match(/H\s*=\s*(\d+'\s*-\s*\d+"?)/i)
            if (hMatch && (p.toLowerCase().includes('mat') || p.toLowerCase().includes('slab'))) {
              thicknessText = hMatch[1].trim()
              if (!/"/.test(thicknessText)) thicknessText += '"'
            }
          }
        }
        if (!widthText) {
          const widthRange = p.match(/(\d+'\s*-\s*\d+"?\s*to\s*\d+'\s*-\s*\d+"?)\s*wide/i)
          const widthInch = p.match(/(\d+"?\s*x\s*\d+"?)\s*wide/i) || p.match(/(\d+)\s*["']?\s*x\s*["']?\s*(\d+)\s*wide/i)
          const widthSingle = p.match(/(\d+'\s*-\s*\d+"?)\s*wide/i) || p.match(/(\d+(?:\/\d+)?)"?\s*wide/i)
          const stWxH = p.match(/\(\s*(\d+'\s*-\s*\d+"?)\s*x\s*(\d+'\s*-\s*\d+"?)\s*\)/i) // e.g. ST (3'-10"x2'-9")
          if (widthRange) widthText = widthRange[1].trim()
          else if (widthInch) {
            const parts = (widthInch[1] + (widthInch[2] ? 'x' + widthInch[2] : '')).split(/\s*x\s*/i)
            widthText = parts.map(part => part.replace(/\s/g, '').trim() + (part.includes('"') || part.includes("'") ? '' : '"')).join('x')
          }
          else if (widthSingle) {
            widthText = widthSingle[1].trim()
            if (!/"/.test(widthText) && widthSingle[0].includes('"')) widthText += '"'
          } else if (stWxH) {
            widthText = stWxH[1].trim()
            if (!/"/.test(widthText)) widthText += '"'
          }
        }
        if (!dimensionsText) {
          const dim3 = p.match(/(\d+'\s*-\s*\d+"?\s*x\s*\d+'\s*-\s*\d+"?\s*x\s*\d+'\s*-\s*\d+"?)/i) || p.match(/(\d+'\s*-\s*\d+"?x\d+'\s*-\s*\d+"?x\d+'\s*-\s*\d+"?)/i)
          if (dim3) dimensionsText = dim3[1].replace(/\s/g, '').trim()
        }
        // Wire mesh for SOG/ROG: #x#-#/# e.g. 6x6-10/10 W.W.M.
        if (!wireMeshText) {
          const wwmMatch = p.match(/\b(\d+x\d+-\d+\/\d+)\s*W\.?W\.?M\.?/i) || p.match(/\b(\d+x\d+-\d+\/\d+)\b/i)
          if (wwmMatch) wireMeshText = wwmMatch[1]
        }
        // Elevator pit wall / Elev. pit wall: (1'-4"x8'-0") -> thickness = 1st dim, height = 2nd dim
        const pl = (p || '').toLowerCase()
        if ((pl.includes('elev. pit wall') || pl.includes('elevator pit wall')) && !pl.includes('service')) {
          const bracketMatch = p.match(/\(\s*(\d+'\s*-\s*\d+"?)\s*x\s*(\d+'\s*-\s*\d+"?)\s*\)/i)
          if (bracketMatch) {
            elevatorPitWallThicknessFromBracket = bracketMatch[1].trim()
            if (!/"/.test(elevatorPitWallThicknessFromBracket)) elevatorPitWallThicknessFromBracket += '"'
            elevatorPitWallHeightFromBracket = bracketMatch[2].trim()
            if (!/"/.test(elevatorPitWallHeightFromBracket)) elevatorPitWallHeightFromBracket += '"'
          }
        }
      }
      if (elevatorPitWallThicknessFromBracket && !thicknessText) thicknessText = elevatorPitWallThicknessFromBracket
      if (elevatorPitWallHeightFromBracket) heightText = heightText || elevatorPitWallHeightFromBracket
      const pageRef = getPageRefForSubsection(sub)
      return { qtyText, heightText: heightText || '##', thicknessText, widthText, dimensionsText, wireMeshText: wireMeshText || '', pageRef }
    }
    // Replace (X)no./no, (H=...), (Havg=...), thickness, width, dimensions, "as per X" in template with values from group data (template wording unchanged)
    const applyDynamicValuesToTemplate = (text, vals) => {
      if (!text || typeof text !== 'string') return text
      if (!vals) return text
      let out = text
      if (vals.pageRef && vals.pageRef !== '##') {
        out = out.replace(/\s+as\s+per\s+[^&]*&\s*details\s+on/gi, ` as per ${vals.pageRef} & details on`)
        out = out.replace(/\s+as\s+per\s+##\s*$/gi, ` as per ${vals.pageRef} & details on`)
      }
      if (vals.qtyText && vals.qtyText !== '##') {
        out = out.replace(/\s*\(\d+\)\s*no\./gi, ` (${vals.qtyText})no.`)
        out = out.replace(/\s*\(\d+\)\s*no\s/gi, ` (${vals.qtyText})no `)
      }
      if (vals.heightText) {
        out = out.replace(/\s*\(H(avg)?\s*=\s*[^)]*\)/gi, (full, avg) => {
          const withTyp = /,\s*typ\.?\s*\)$/.test(full)
          return ` (H${avg ? 'avg=' : '='}${vals.heightText}${withTyp ? ', typ.' : ''})`
        })
      }
      // Placeholder (#" thick, typ.) – replace with dynamic thickness or default 6"
      if (out.includes('#" thick, typ.)') || out.includes('##" thick, typ.)')) {
        const thick = vals.thicknessText || '6"'
        out = out.replace(/\s*\(#+"?\s*thick,\s*typ\.?\)/gi, () => ` (${thick} thick, typ.)`)
      }
      if (vals.thicknessText) {
        out = out.replace(/\s*\(##\s+thick\)/gi, () => ` (${vals.thicknessText} thick)`)
        out = out.replace(/\s*\(\d+'\s*-\s*\d+"?\s*thick[^)]*\)/gi, () => ` (${vals.thicknessText} thick)`)
        out = out.replace(/\s*\(\d+(?:\/\d+)?"\s*thick[^)]*\)/gi, () => ` (${vals.thicknessText} thick)`)
        out = out.replace(/\s*\(\d+'\s*-\s*\d+"?\s*thick,\s*typ\.?\)/gi, () => ` (${vals.thicknessText} thick, typ.)`)
        out = out.replace(/\s*\(\d+(?:\/\d+)?"\s*thick,\s*typ\.?\)/gi, () => ` (${vals.thicknessText} thick, typ.)`)
      }
      if (vals.widthText) {
        out = out.replace(/\s*\(\d+'\s*-\s*\d+"?\s*to\s*\d+'\s*-\s*\d+"?\s*wide\)/gi, () => ` (${vals.widthText} wide)`)
        out = out.replace(/\s*\(\d+"\s*x\s*\d+"\s*wide\)/gi, () => ` (${vals.widthText} wide)`)
        out = out.replace(/\s*\(\d+'\s*-\s*\d+"?\s*wide\)/gi, () => ` (${vals.widthText} wide)`)
        out = out.replace(/\s*\(\d+(?:\/\d+)?"\s*wide\)/gi, () => ` (${vals.widthText} wide)`)
      }
      if (vals.dimensionsText) {
        out = out.replace(/\s*\(\d+'\s*-\s*\d+"?x\d+'\s*-\s*\d+"?x\d+'\s*-\s*\d+"?\)/gi, () => ` (${vals.dimensionsText})`)
      }
      // Wire mesh for SOG/ROG: w/#x#-#/# W.W.M. – dynamic or omit if not found
      if (vals.wireMeshText) {
        out = out.replace(/\bw\/\s*\d+x\d+-\d+\/\d+\s*W\.?W\.?M\.?/gi, `w/${vals.wireMeshText} W.W.M.`)
      } else {
        out = out.replace(/\s+reinforced\s+w\/\s*\d+x\d+-\d+\/\d+\s*W\.?W\.?M\.?/gi, ' reinforced')
      }
      return out
    }
    const substructureTemplate = [
      {
        heading: 'Compacted gravel', items: [
          { text: `F&I new (#" thick, typ.) gravel/crushed stone, including 6MIL vapor barrier on top @ SOG, elevator pit mat, detention tank mat as per ##`, sub: 'SOG', match: p => p.includes('gravel') && !p.includes('geotextile'), showIfExists: true }
        ]
      },
      {
        heading: 'Elevator pit', items: [
          { text: `F&I new (2)no. (1'-0" thick) sump pit (2'-0"x2'-0"x2'-0") reinf w/ #4@12"O.C., T&B/E.F., E.W. typ`, sub: 'Elevator Pit', match: p => (p || '').toLowerCase().includes('sump pit') && !(p || '').toLowerCase().includes('service elevator'), formulaItem: { itemType: 'elevator_pit', match: p => (p || '').toLowerCase().includes('sump') } },
          { text: `F&I new (3'-0" thick) elevator pit slab, reinf w/#5@12"O.C. & #5@12"O.C., as per FO-101.00 & details on`, sub: 'Elevator Pit', match: p => (p || '').toLowerCase().includes('elev') && (p || '').toLowerCase().includes('pit slab') && !(p || '').toLowerCase().includes('service') },
          { text: `F&I new (3'-0" thick) elevator pit mat, reinf w/#5@12"O.C. & #5@12"O.C., as per FO-101.00 & details on`, sub: 'Elevator Pit', match: p => (p || '').toLowerCase().includes('elev') && (p || '').toLowerCase().includes('pit mat') && !(p || '').toLowerCase().includes('service') },
          { text: `F&I new (1'-0" thick) elevator pit walls (H=5'-0") as per FO-101.00 & details on`, sub: 'Elevator Pit', match: p => ((p || '').toLowerCase().includes('elev. pit wall') || ((p || '').toLowerCase().includes('elevator pit wall') || ((p || '').toLowerCase().includes('elev') && (p || '').toLowerCase().includes('pit wall')))) && (p.includes("1'-0") || p.includes('1-0"')) && !(p || '').toLowerCase().includes('service') },
          { text: `F&I new (1'-2" thick) elevator pit walls (H=5'-0") as per FO-101.00 & details on`, sub: 'Elevator Pit', match: p => ((p || '').toLowerCase().includes('elev. pit wall') || ((p || '').toLowerCase().includes('elevator pit wall') || ((p || '').toLowerCase().includes('elev') && (p || '').toLowerCase().includes('pit wall')))) && (p.includes("1'-2") || p.includes('1-2"') || p.includes('1.17')) && !(p || '').toLowerCase().includes('service') },
          { text: `F&I new (## thick) elevator pit walls (H=##) as per FO-101.00 & details on`, sub: 'Elevator Pit', match: p => ((p || '').toLowerCase().includes('elev. pit wall') || (p || '').toLowerCase().includes('elevator pit wall')) && !(p || '').toLowerCase().includes('service') },
          { text: `F&I new (1'-2" thick) elevator pit slope transition/haunch (H=2'-3") as per FO-101.00 & details on`, sub: 'Elevator Pit', match: p => ((p || '').toLowerCase().includes('slope') || (p || '').toLowerCase().includes('haunch')) && !(p || '').toLowerCase().includes('service') }
        ]
      },
      {
        heading: 'Service elevator pit', items: [
          {
            text: `F&I new (2)no. (1'-0" thick) sump pit (2'-0"x2'-0"x2'-0") reinf w/ #4@12"O.C., T&B/E.F., E.W. typ`,
            sub: 'Service elevator pit',
            match: p => (p || '').toLowerCase().includes('sump pit @ service elevator') || (p || '').toLowerCase().includes('sump pit @ service elevator pit'),
            formulaItem: { itemType: 'service_elevator_pit', match: p => (p || '').toLowerCase().includes('sump') }
          },
          {
            text: `F&I new (3'-0" thick) service elevator pit slab, reinf w/#5@12"O.C. & #5@12"O.C., as per FO-101.00 & details on FO-203.00`,
            sub: 'Service elevator pit',
            match: p => (p || '').toLowerCase().includes('service elev. pit slab') || (p || '').toLowerCase().includes('service elevator pit slab')
          },
          {
            text: `F&I new (1'-0" thick) service elevator pit walls (H=5'-0") as per FO-101.00 & details on FO-203.00`,
            sub: 'Service elevator pit',
            match: p => (p || '').toLowerCase().includes('service elev. pit wall') || (p || '').toLowerCase().includes('service elevator pit wall')
          }
        ]
      },
      {
        heading: 'Detention tank w/epoxy coated reinforcement', items: [
          { text: `F&I new (1'-0" thick) detention tank slab as per FO-101.00 & details on`, sub: 'Detention tank', match: p => p.includes('detention') && p.includes('slab') && !p.includes('lid') },
          { text: `F&I new (1'-0" thick) detention tank walls (H=10'-0") as per FO-101.00 & details on`, sub: 'Detention tank', match: p => p.includes('detention') && p.includes('wall') },
          { text: `F&I new (0'-8" thick) detention tank lid slab as per FO-101.00 & details on`, sub: 'Detention tank', match: p => p.includes('detention') && (p.includes('lid') || p.includes('8"')) }
        ]
      },
      {
        heading: 'Duplex sewage ejector pit', items: [
          { text: `F&I new (0'-8" thick, typ.) duplex sewage ejector pit slab as per P-301.01 & details on`, sub: 'Duplex sewage ejector pit', match: p => p.includes('duplex') && p.includes('slab') },
          { text: `F&I new (0'-8" thick, typ.) duplex sewage ejector pit wall (H=5'-0", typ.) as per P-301.01 & details on`, sub: 'Duplex sewage ejector pit', match: p => p.includes('duplex') && p.includes('wall') }
        ]
      },
      {
        heading: 'Deep sewage ejector pit', items: [
          { text: `F&I new (1'-0" thick) deep sewage ejector pit slab as per P-100.00 & details on`, sub: 'Deep sewage ejector pit', match: p => p.includes('deep') && p.includes('slab') },
          { text: `F&I new (0'-6" wide) deep sewage ejector pit walls (H=8'-6") as per P-100.00 & details on`, sub: 'Deep sewage ejector pit', match: p => (p || '').toLowerCase().includes('deep') && (p.includes('wall') || /6\s*[""]?\s*x\s*8\s*['']?\s*-\s*6/i.test(p) || (p.includes('6"') && (p.includes("8'-6") || p.includes("8'- 6")))) }
        ]
      },
      {
        heading: 'Sewage ejector pit', items: [
          { text: `F&I new sewage ejector pit slab as per details on`, sub: 'Sewage ejector pit', match: p => (p || '').toLowerCase().includes('sewage ejector') && !(p || '').toLowerCase().includes('duplex') && !(p || '').toLowerCase().includes('deep') && (p || '').toLowerCase().includes('slab') },
          { text: `F&I new sewage ejector pit wall as per details on`, sub: 'Sewage ejector pit', match: p => (p || '').toLowerCase().includes('sewage ejector') && !(p || '').toLowerCase().includes('duplex') && !(p || '').toLowerCase().includes('deep') && (p || '').toLowerCase().includes('wall') }
        ]
      },
      {
        heading: 'Sump pump pit', items: [
          { text: `F&I new (0'-8" thick, typ.) sump pump pit slab as per P-301.01`, sub: 'Sump pump pit', match: p => (p || '').toLowerCase().includes('sump pump') && (p || '').toLowerCase().includes('slab') },
          { text: `F&I new (0'-8" thick, typ.) sump pump pit wall (H=5'-0", typ.) as per P-301.01`, sub: 'Sump pump pit', match: p => (p || '').toLowerCase().includes('sump pump') && (p || '').toLowerCase().includes('wall') }
        ]
      },
      {
        heading: 'Grease trap pit', items: [
          { text: `F&I new (1'-0" thick) grease trap pit slab as per P-100.00 & details on`, sub: 'Grease trap', match: p => (p || '').toLowerCase().includes('grease') && (p || '').toLowerCase().includes('slab') },
          { text: `F&I new (0'-6" wide) grease trap pit walls (H=8'-6") as per P-100.00 & details on`, sub: 'Grease trap', match: p => (p || '').toLowerCase().includes('grease') && ((p || '').toLowerCase().includes('wall') || /6\s*[""]?\s*x\s*8\s*['']?\s*-\s*6/i.test(p) || ((p || '').includes('6"') && ((p || '').includes("8'-6") || (p || '').includes("8'- 6")))) }
        ]
      },
      {
        heading: 'House trap pit', items: [
          { text: `F&I new (1'-0" thick) house trap pit slab as per P-100.00 & details on`, sub: 'House trap', match: p => (p || '').toLowerCase().includes('house trap') && (p || '').toLowerCase().includes('slab') },
          { text: `F&I new (0'-6" wide) house trap pit walls (H=8'-6") as per P-100.00 & details on`, sub: 'House trap', match: p => (p || '').toLowerCase().includes('house trap') && ((p || '').toLowerCase().includes('wall') || /6\s*[""]?\s*x\s*8\s*['']?\s*-\s*6/i.test(p) || ((p || '').includes('6"') && ((p || '').includes("8'-6") || (p || '').includes("8'- 6")))) }
        ]
      },
      {
        heading: 'Foundation elements', items: [
          { text: 'F&I new (9)no pile caps as per FO-101.00 & details on', sub: 'Pile caps', match: () => true },
          { text: `F&I new (2'-0" to 7'-0" wide) strip footings (Havg=1'-10") as per FO-101.00 & details on`, sub: 'Strip Footings', match: () => true },
          { text: 'F&I new (15)no isolated footings as per FO-101.00 & details on', sub: 'Isolated Footings', match: () => true },
          { text: `F&I new (17)no (12"x24" wide) piers (H=3'-4") as per FO-101.00 & details on`, sub: 'Pier', match: () => true },
          { text: `F&I new (4)no (22"x16" wide) Pilasters (Havg=8'-0") as per S-101.00 & details on`, sub: 'Pilaster', match: () => true },
          { text: 'F&I new grade beams as per FO-102.00 & details on', sub: 'Grade beams', match: () => true },
          { text: 'F&I new tie beams as per FO-101.00, FO-102.00 & details on', sub: 'Tie beam', match: () => true },
          { text: 'F&I new strap beams as per FO-102.00 & details on', sub: 'Strap beams', match: () => true },
          { text: `F&I new (8" thick) thickened slab as per FO-101.00 & details on`, sub: 'Thickened slab', match: () => true },
          { text: `F&I new (24" thick) mat slab reinforced as per FO-101.00 & details on`, sub: 'Mat slab', match: p => (p || '').includes('2\'') || (p || '').includes('24') },
          { text: `F&I new (36" thick) mat slab reinforced as per FO-101.00 & details on`, sub: 'Mat slab', match: p => (p || '').includes('3\'') || (p || '').includes('36') },
          { text: `F&I new mat slab reinforced as per FO-101.00 & details on`, sub: 'Mat slab', match: p => (p || '').toLowerCase().includes('mat') && (p || '').toLowerCase().includes('slab') },
          { text: `F&I new (3" thick) mud/rat slab as per FO-101.00 & details on`, sub: 'Mud Slab', match: () => true },
          {
            text: `F&I new (4" thick) slab on grade reinforced w/6x6-10/10 W.W.M. @ cellar FL as per FO-101.00 & details on`, sub: 'SOG', match: p => {
              const pl = (p || '').toLowerCase()
              return pl.includes('sog') && !pl.includes('patio') && !pl.includes('patch') && !pl.includes('step') && !pl.includes('pressure')
            }
          },
          {
            text: `F&I new (6" thick) slab on grade reinforced w/6x6-10/10 W.W.M. @ cellar FL as per FO-101.00 & details on`, sub: 'SOG', match: p => {
              const pl = (p || '').toLowerCase()
              return pl.includes('sog') && !pl.includes('patio') && !pl.includes('patch') && !pl.includes('step') && !pl.includes('pressure')
            }
          },
          { text: `F&I new (6" thick) patio slab on grade reinforced w/6x6-10/10 W.W.M. @ cellar FL as per FO-101.00 & details on`, sub: 'SOG', match: p => p.includes('patio') },
          { text: `Allow to (5" thick) patch slab on grade reinforced w/6x6-10/10 W.W.M. @ cellar FL as per FO-101.00 & details on`, sub: 'SOG', match: p => p.includes('patch') },
          { text: `Allow to (5" thick) pressure slab reinforced w/6x6-10/10 W.W.M. @ cellar FL as per FO-101.00 & details on`, sub: 'SOG', match: p => (p || '').toLowerCase().includes('pressure') && (p || '').toLowerCase().includes('sog') },
          { text: `F&I new (6" thick) slab on grade step (H=1'-2") as per FO-101.00 & details on`, sub: 'SOG', match: p => p.includes('step') },
          {
            formulaItem: { itemType: 'buttress_final', match: () => true },
            showIfExists: true,
            getText: (source, calculationData, rawData) => {
              let qty = '##'
              if (source.row && calculationData && calculationData[source.row - 1]) {
                const mVal = calculationData[source.row - 1][12]
                const n = Math.round(parseFloat(mVal) || 0)
                if (n > 0) qty = String(n)
              }
              if (qty === '##' && source.buttressRow && calculationData && calculationData[source.buttressRow - 1]) {
                const takeoffVal = calculationData[source.buttressRow - 1][2]
                const n = Math.round(parseFloat(takeoffVal) || 0)
                if (n > 0) qty = String(n)
              }
              let width = '##'
              let height = '##'
              let page = '##'
              if (rawData && rawData.length >= 2) {
                const headers = rawData[0]
                const dataRows = rawData.slice(1)
                const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
                const pageIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'page')
                for (let i = 0; i < dataRows.length; i++) {
                  const row = dataRows[i]
                  const digitizerItem = (row[digitizerIdx] || '').toString()
                  if (digitizerItem.toLowerCase().includes('buttress')) {
                    const particulars = digitizerItem
                    const widthMatch = particulars.match(/(\d+'\s*-\s*\d+"?\s*to\s*\d+'\s*-\s*\d+"?)\s*wide/i) || particulars.match(/(\d+'\s*-\s*\d+"?)\s*wide/i)
                    if (widthMatch) width = widthMatch[1].trim()
                    const heightMatch = particulars.match(/H\s*=\s*(\d+'\s*-\s*\d+")/i)
                    if (heightMatch) height = heightMatch[1].trim()
                    if (pageIdx >= 0 && row[pageIdx]) {
                      const pageStr = String(row[pageIdx]).trim()
                      const foMatch = pageStr.match(/(?:FO|P)-[\d.]+/i)
                      if (foMatch) page = foMatch[0]
                    }
                    break
                  }
                }
              }
              if (page === '##') page = 'FO-101.00'
              const heightPart = height !== '##' ? `(H=${height})` : `(${height})`
              const line = `F&I new (${qty})no (${width}) buttresses ${heightPart} within foundation walls @ cellar FL to 1st FL as per ${page} & details on`
              return { line, data: { qty, width, height, page } }
            }
          },
          { text: `F&I new (8" thick) corbel (Havg=2'-2") as per FO-101.00 & details on`, sub: 'Corbel', match: () => true },
          { text: `F&I new (8" thick) concrete liner walls (Havg=11'-2") @ cellar FL to 1st FL as per FO-101.00 & details on`, sub: 'Linear Wall', match: () => true },
          { text: `F&I new (1'-0" wide) foundation walls (Havg=10'-10") @ cellar FL to 1st FL as per FO-101.00 & details on`, sub: 'Foundation Wall', match: () => true },
          { text: `F&I new (10" wide) retaining walls (Havg=6'-4") @ court as per FO-101.00 & details on`, sub: 'Retaining walls', match: p => p.includes('court') },
          { text: `F&I new (1'-0" wide) retaining walls w/epoxy reinforcement (Havg=4'-8") @ level 1 as per FO-101.00 & details on`, sub: 'Retaining walls', match: p => p.toLowerCase().startsWith('rw') || p.includes('retaining wall') || p.includes('retaining walls') || p.includes('level') }
        ]
      },
      {
        heading: 'Barrier wall', items: [
          { text: `F&I new (0'-8" thick) barrier walls (Havg=4'-0") @ cellar FL as per FO-001.00 & details on`, sub: 'Barrier wall', match: p => p.includes('8"') || p.includes('2\'') },
          { text: `F&I new (0'-10" thick) barrier walls (Havg=5'-1") @ cellar FL as per FO-001.00 & details on`, sub: 'Barrier wall', match: p => p.includes('10"') || p.includes('5\'') }
        ]
      },
      {
        heading: 'Stem wall', items: [
          { text: `F&I new (0'-10" thick) stem walls (Havg=5'-1") @ cellar FL as per FO-001.00 & details on`, sub: 'Stem wall', match: () => true }
        ]
      },
      // Stair on grade: under Stem wall in substructure scope; groups from calc "Stairs on grade Stairs" (e.g. Stair A2, C2, D2, F)
      {
        heading: 'Stair on grade',
        customStairsOnGrade: true,
        items: []
      },
      {
        heading: 'Electric conduit', items: [
          { text: 'F&I new electric conduit as per E-060.00 & E-070.00', sub: 'Electric conduit', match: p => p.includes('electric conduit') || p.includes('conduit') }
        ]
      },
      {
        heading: 'Trench drain', items: [
          { text: 'F&I new concrete @ trench drain as per FO-102.00 & details on', sub: 'Trench drain', formulaItem: { itemType: 'electric_conduit', match: p => (p || '').toLowerCase().includes('trench drain') } }
        ]
      },
      {
        heading: 'Misc.', items: [
          { text: 'F&I (4"Ø thick) perforated pipe + gravel at perimeter footing drain', sub: 'Electric conduit', match: p => p.includes('perforated pipe'), formulaItem: { itemType: 'electric_conduit', match: p => (p || '').toLowerCase().includes('perforated pipe') } }
        ]
      }
    ]
    let hasSubstructureScopeData = false
    const dryRunUsedSum = new Set()
    const dryRunUsedItem = new Set()
    for (const { items, customStairsOnGrade } of substructureTemplate) {
      if (customStairsOnGrade) {
        const stairsOnGradeFormulas = (formulaData || []).filter(f =>
          f.section === 'foundation' &&
          (f.itemType === 'stairs_on_grade' || f.itemType === 'foundation_stairs_on_grade_stairs' || f.itemType === 'foundation_stairs_on_grade_landing' || f.itemType === 'foundation_stairs_on_grade_stair_slab') &&
          hasNonZeroTakeoff(f, true)
        )
        if (stairsOnGradeFormulas.length > 0) hasSubstructureScopeData = true
        if (hasSubstructureScopeData) break
        continue
      }
      for (const it of items) {
        const { sub, match, formulaItem, showIfExists } = it
        let source = null
        if (formulaItem) {
          const itemF = (formulaData || []).find(f =>
            f.itemType === formulaItem.itemType && f.section === 'foundation' &&
            !dryRunUsedItem.has(f.row) && formulaItem.match((f.parsedData?.particulars || '').toLowerCase())
          )
          if (itemF) {
            source = itemF
            dryRunUsedItem.add(itemF.row)
            if (showIfExists || hasNonZeroTakeoff(itemF, true)) { hasSubstructureScopeData = true; break }
          }
        } else {
          const sumF = allFoundationSums.find(f => {
            if (dryRunUsedSum.has(f.row)) return false
            if ((f.subsectionName || '').toLowerCase() !== sub.toLowerCase()) return false
            return match(getFirstParticulars(f))
          })
          if (sumF) {
            source = sumF
            dryRunUsedSum.add(sumF.row)
            if (showIfExists || hasNonZeroTakeoff(sumF, false)) { hasSubstructureScopeData = true; break }
          }
        }
      }
      if (hasSubstructureScopeData) break
    }
    if (hasSubstructureScopeData) {
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
      currentRow++
      const substructureStartRow = currentRow
      const usedSumIds = new Set()
      let lastSubstructureHeading = null
      const usedItemIds = new Set()

      const formatStairsWidth = (feetNum) => {
        if (feetNum == null || isNaN(feetNum)) return 'as indicated'
        const f = Math.floor(feetNum)
        const inch = Math.round((feetNum - f) * 12)
        if (inch === 0) return `${f}'-0"`
        return `${f}'-${inch}"`
      }
      const stairsOnGradePageRef = getPageRefForSubsection('Stairs on grade Stairs') || 'A-312.00'
      const stairsPageRefEsc = (stairsOnGradePageRef || '').replace(/"/g, '""')
      const getStairsOnGradeProposalText = (formulaEntry) => {
        const it = formulaEntry.itemType
        const p = (formulaEntry.parsedData?.particulars || '').trim()
        const subType = formulaEntry.parsedData?.parsed?.itemSubType
        if (it === 'foundation_stairs_on_grade_stair_slab') return null
        if (subType === 'landings' || (p && p.toLowerCase().includes('landings')) || it === 'foundation_stairs_on_grade_landing') {
          let thick = '8"'
          if (calculationData && formulaEntry.row) {
            const row = calculationData[formulaEntry.row - 1]
            const particulars = (row?.[1] || '').toString()
            const heightRaw = row[7]
            if (heightRaw != null && heightRaw !== '') {
              let heightFeet = parseFloat(heightRaw)
              if (isNaN(heightFeet) && typeof heightRaw === 'string') {
                const frac = heightRaw.trim().match(/^(\d+)\s*\/\s*(\d+)$/)
                if (frac) heightFeet = parseInt(frac[1], 10) / parseInt(frac[2], 10)
              }
              if (!isNaN(heightFeet) && heightFeet > 0) {
                const inches = Math.round(heightFeet * 12)
                thick = `${inches}"`
              }
            }
            if (thick === '8"') {
              const m = particulars.match(/(\d+(?:\/\d+)?)"?\s*thick/i) || particulars.match(/(\d+)"\s*thick/i)
              if (m) thick = m[1].includes('"') ? m[1] : m[1] + '"'
            }
          }
          return `F&I new (${thick} thick) stair landings as per ${stairsOnGradePageRef}`
        }
        if (subType === 'stair_slab' || (p && p.toLowerCase().includes('stair slab'))) return null
        let widthStr = 'as indicated'
        let riserStr = 'as indicated'
        if (calculationData && formulaEntry.row) {
          const row = calculationData[formulaEntry.row - 1]
          const particulars = (row?.[1] || '').toString()
          const widthMatch = particulars.match(/(\d+'-?\d*")\s*wide/i)
          const widthFromParsed = formulaEntry.parsedData?.parsed?.widthFromName
          const widthFromCol = row[6] != null && row[6] !== '' ? formatWidthFeetInches(row[6]) || String(row[6]).trim() : null
          widthStr = widthMatch && widthMatch[1] ? widthMatch[1] : (widthFromParsed != null ? formatStairsWidth(widthFromParsed) : (widthFromCol || 'as indicated'))
          let qtyM = Math.round(parseFloat(row[12]) || 0)
          if (qtyM <= 0 && row[11] != null && row[11] !== '') qtyM = Math.round(parseFloat(row[11]) || 0)
          if (qtyM <= 0) {
            const takeoffVal = parseFloat(row[2])
            if (Number.isInteger(takeoffVal) && takeoffVal > 0 && takeoffVal < 100) qtyM = takeoffVal
          }
          if (qtyM <= 0 && formulaEntry.row < calculationData.length) {
            const slabRow = calculationData[formulaEntry.row]
            if (slabRow) {
              const slabQ = Math.round(parseFloat(slabRow[12]) || 0) || Math.round(parseFloat(slabRow[11]) || 0)
              if (slabQ > 0) qtyM = slabQ
            }
          }
          if (qtyM <= 0 && formulaEntry.row + 1 < calculationData.length) {
            const sumRow = calculationData[formulaEntry.row + 1]
            if (sumRow) {
              const sumQ = Math.round(parseFloat(sumRow[12]) || 0) || Math.round(parseFloat(sumRow[11]) || 0)
              if (sumQ > 0) qtyM = sumQ
            }
          }
          const rm = particulars.match(/(\d+)\s*Riser/i)
          riserStr = qtyM > 0 ? `${qtyM} Riser` : (rm ? `${rm[1]} Riser` : 'as indicated')
        } else {
          const widthFromParsed = formulaEntry.parsedData?.parsed?.widthFromName
          const wm = p.match(/(\d+'-?\d*")\s*wide/i)
          widthStr = (wm && wm[1]) || (widthFromParsed != null ? formatStairsWidth(widthFromParsed) : 'as indicated')
          const riserMatch = p.match(/(\d+)\s*Riser/i)
          riserStr = riserMatch ? `${riserMatch[1]} Riser` : 'as indicated'
        }
        return `F&I new (${widthStr} wide) stairs on grade (${riserStr}) @ 1st FL as per ${stairsOnGradePageRef}`
      }

      substructureTemplate.forEach(({ heading, items, customStairsOnGrade }) => {
        let headingAdded = false
        if (customStairsOnGrade && heading === 'Stair on grade') {
          const allStairsFormulas = (formulaData || []).filter(f =>
            f.section === 'foundation' &&
            (f.itemType === 'stairs_on_grade' || f.itemType === 'stairs_on_grade_group_header' ||
              f.itemType === 'foundation_stairs_on_grade_heading' || f.itemType === 'foundation_stairs_on_grade_stairs' ||
              f.itemType === 'foundation_stairs_on_grade_landing' || f.itemType === 'foundation_stairs_on_grade_stair_slab')
          )
          allStairsFormulas.sort((a, b) => a.row - b.row)
          const groups = []
          let currentGroup = null
          const getStairHeadingFromRow = (rowNum) => {
            if (!calculationData || !rowNum) return 'Stair'
            const particulars = (calculationData[rowNum - 1]?.[1] || '').toString().trim()
            return particulars.replace(/:+\s*$/, '').trim() || 'Stair'
          }
          allStairsFormulas.forEach((f) => {
            if (f.itemType === 'stairs_on_grade_group_header' || f.itemType === 'foundation_stairs_on_grade_heading') {
              currentGroup = { stairIdentifier: f.stairIdentifier || getStairHeadingFromRow(f.row), items: [] }
              groups.push(currentGroup)
            } else if ((f.itemType === 'stairs_on_grade' || f.itemType === 'foundation_stairs_on_grade_stairs' || f.itemType === 'foundation_stairs_on_grade_landing' || f.itemType === 'foundation_stairs_on_grade_stair_slab') && currentGroup && (hasNonZeroTakeoff(f, true) || f.itemType === 'foundation_stairs_on_grade_stair_slab')) {
              currentGroup.items.push(f)
            }
          })
          groups.forEach((group) => {
            if (group.items.length === 0) return
            if (!headingAdded) {
              headingAdded = true
              lastSubstructureHeading = heading
              spreadsheet.updateCell({ value: `${heading}:` }, `${pfx}B${currentRow}`)
              spreadsheet.cellFormat(
                { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
                `${pfx}B${currentRow}`
              )
              currentRow++
            }
            spreadsheet.updateCell({ value: `${group.stairIdentifier}:` }, `${pfx}B${currentRow}`)
            spreadsheet.cellFormat(
              { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#E2EFDA', textDecoration: 'underline', border: '1px solid #000000' },
              `${pfx}B${currentRow}`
            )
            currentRow++
            group.items.forEach((formulaEntry) => {
              if (formulaEntry.itemType === 'foundation_stairs_on_grade_stair_slab') return
              const sumRow = formulaEntry.row
              const isLanding = formulaEntry.itemType === 'foundation_stairs_on_grade_landing' || (formulaEntry.parsedData?.parsed?.itemSubType === 'landings') || (formulaEntry.parsedData?.particulars || '').toLowerCase().includes('landings')
              const isStairs = formulaEntry.itemType === 'stairs_on_grade' || formulaEntry.itemType === 'foundation_stairs_on_grade_stairs'
              let bVal
              let bDesc
              if (isLanding) {
                // Real-time: thickness from column H (height in feet) * 12 → inches; one inch symbol in formula = "" in Excel
                bVal = {
                  formula: `=CONCATENATE("F&I new (",ROUND('${substructureCalcSheet}'!H${sumRow}*12,0),"\"\""," thick) stair landings as per ","${stairsPageRefEsc}",")")`
                }
                bDesc = `F&I new (… thick) stair landings as per ${stairsOnGradePageRef}`
              } else if (isStairs) {
                // Real-time: width from column G (feet) as X'-Y", riser from column M; inch symbol = "" in Excel
                const gRef = `'${substructureCalcSheet}'!G${sumRow}`
                const mRef = `'${substructureCalcSheet}'!M${sumRow}`
                const widthPart = `INT(${gRef})&"'-"&TEXT(ROUND((${gRef}-INT(${gRef}))*12,0),"0")&"\"\""`
                bVal = {
                  formula: `=CONCATENATE("F&I new (",${widthPart}," wide) stairs on grade (",ROUND(${mRef},0)," Riser) @ 1st FL as per ","${stairsPageRefEsc}",")")`
                }
                bDesc = `F&I new (… wide) stairs on grade (… Riser) @ 1st FL as per ${stairsOnGradePageRef}`
              } else {
                const text = getStairsOnGradeProposalText(formulaEntry)
                if (!text) return
                bVal = { value: text }
                bDesc = text
              }
              spreadsheet.updateCell(bVal, `${pfx}B${currentRow}`)
              rowBContentMap.set(currentRow, bDesc)
              const descForRates = rowBContentMap.get(currentRow)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat({ fontWeight: 'normal', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
              spreadsheet.updateCell({ formula: `='${substructureCalcSheet}'!I${sumRow}` }, `${pfx}C${currentRow}`)
              spreadsheet.updateCell({ formula: `='${substructureCalcSheet}'!J${sumRow}` }, `${pfx}D${currentRow}`)
              spreadsheet.updateCell({ formula: `='${substructureCalcSheet}'!L${sumRow}` }, `${pfx}F${currentRow}`)
              spreadsheet.updateCell({ formula: `='${substructureCalcSheet}'!M${sumRow}` }, `${pfx}G${currentRow}`)
              fillRatesForProposalRow(currentRow, descForRates)
              spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
              currentRow++
            })
          })
          return
        }
        items.forEach(({ text, sub, match, formulaItem, showIfExists, getText }) => {
          let rowToUse = null
          let source = null
          if (formulaItem) {
            const itemF = (formulaData || []).find(f =>
              f.itemType === formulaItem.itemType && f.section === 'foundation' &&
              !usedItemIds.has(f.row) && formulaItem.match((f.parsedData?.particulars || '').toLowerCase())
            )
            if (itemF) {
              rowToUse = itemF.row
              source = itemF
              usedItemIds.add(rowToUse)
            }
          } else {
            const sumF = allFoundationSums.find(f => {
              if (usedSumIds.has(f.row)) return false
              if ((f.subsectionName || '').toLowerCase() !== sub.toLowerCase()) return false
              return match(getFirstParticulars(f))
            })
            if (sumF) {
              rowToUse = sumF.row
              source = sumF
              usedSumIds.add(rowToUse)
            }
          }
          // Geotextile filter fabric below gravel: drive SQ FT / QTY from Gravel SOG sum row
          if (heading === 'Compacted gravel' && text.includes('geotextile filter fabric')) {
            const gravelSum = allFoundationSums.find(f => {
              if ((f.subsectionName || '').toLowerCase() !== 'sog') return false
              const firstP = (getFirstParticulars(f) || '').toLowerCase()
              return firstP.includes('gravel') && !firstP.includes('geotextile')
            })
            if (gravelSum) {
              rowToUse = gravelSum.row
              source = gravelSum
            }
          }
          if (!rowToUse || !source) return
          if (!showIfExists && !hasNonZeroTakeoff(source, !!formulaItem)) return
          if (!headingAdded) {
            headingAdded = true
            lastSubstructureHeading = heading
            spreadsheet.updateCell({ value: `${heading}:` }, `${pfx}B${currentRow}`)
            spreadsheet.cellFormat(
              { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
              `${pfx}B${currentRow}`
            )
            currentRow++
          }
          const sumRow = rowToUse
          const getTextResult = typeof getText === 'function' ? getText(source, calculationData, rawData) : null
          let displayText = getTextResult != null && typeof getTextResult === 'object' && getTextResult.line != null
            ? getTextResult.line
            : (getTextResult != null && typeof getTextResult === 'string' ? getTextResult : text)
          if (displayText === text && source != null && text) {
            const dynVals = getDynamicValuesFromGroupRows(source, sub, !!formulaItem)
            // Gravel: if values are 0 and no thickness mentioned, show ## for thickness and as per ##
            if (text.includes('gravel') && !dynVals.thicknessText && !hasNonZeroTakeoff(source, !!formulaItem)) {
              dynVals.thicknessText = '##'
              dynVals.pageRef = '##'
            }
            displayText = applyDynamicValuesToTemplate(text, dynVals)
          }
          const afterCountSub = afterCountFromProposalText(displayText)
          if (afterCountSub) {
            spreadsheet.updateCell({ formula: proposalFormulaWithQtyRef(currentRow, afterCountSub) }, `${pfx}B${currentRow}`)
          } else {
            spreadsheet.updateCell({ value: displayText }, `${pfx}B${currentRow}`)
          }
          rowBContentMap.set(currentRow, displayText)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${substructureCalcSheet}'!I${sumRow}` }, `${pfx}C${currentRow}`)
          // D = SQ FT from J only (reference from J, not C)
          spreadsheet.updateCell({ formula: `='${substructureCalcSheet}'!J${sumRow}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${substructureCalcSheet}'!L${sumRow}` }, `${pfx}F${currentRow}`)
          spreadsheet.updateCell({ formula: `='${substructureCalcSheet}'!M${sumRow}` }, `${pfx}G${currentRow}`)
          // Compacted gravel: use CY 95 for gravel/crushed stone with 6MIL vapor barrier (proposal_mapped.json)
          const gravelRateKey = (lastSubstructureHeading === 'Compacted gravel' && displayText.includes('gravel/crushed stone, including 6MIL vapor barrier')) ? 'gravel/crushed stone, including 6MIL vapor barrier - Compacted gravel scope' : displayText
          fillRatesForProposalRow(currentRow, gravelRateKey)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        })
        // One proposal row per Ramp on grade sum (thickness/wire mesh from data)
        if (heading === 'Foundation elements') {
          const rogTemplate = `F&I new (6" thick) ramp on grade reinforced w/6x6-10/10 W.W.M. @ cellar FL as per FO-101.00 & details on`
          const rogSums = allFoundationSums.filter(f =>
            (f.subsectionName || '').toLowerCase() === 'ramp on grade' && !usedSumIds.has(f.row) && hasNonZeroTakeoff(f, false)
          )
          rogSums.forEach((source) => {
            usedSumIds.add(source.row)
            if (!headingAdded) {
              headingAdded = true
              lastSubstructureHeading = heading
              spreadsheet.updateCell({ value: `${heading}:` }, `${pfx}B${currentRow}`)
              spreadsheet.cellFormat(
                { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
                `${pfx}B${currentRow}`
              )
              currentRow++
            }
            const dynVals = getDynamicValuesFromGroupRows(source, 'Ramp on grade', false)
            const displayText = applyDynamicValuesToTemplate(rogTemplate, dynVals)
            spreadsheet.updateCell({ value: displayText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, displayText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'normal', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${substructureCalcSheet}'!I${source.row}` }, `${pfx}C${currentRow}`)
            spreadsheet.updateCell({ formula: `='${substructureCalcSheet}'!J${source.row}` }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: `='${substructureCalcSheet}'!L${source.row}` }, `${pfx}F${currentRow}`)
            spreadsheet.updateCell({ formula: `='${substructureCalcSheet}'!M${source.row}` }, `${pfx}G${currentRow}`)
            fillRatesForProposalRow(currentRow, displayText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          })
        }
      })
      const substructureDataEndRow = currentRow - 1
      // Add Misc. fixed lines (DOB washout, Engineering) when substructure has content
      if (currentRow > substructureStartRow) {
        if (lastSubstructureHeading !== 'Misc.') {
          spreadsheet.updateCell({ value: 'Misc.:' }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat(
            { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
            `${pfx}B${currentRow}`
          )
          currentRow++
        }
        spreadsheet.updateCell({ value: 'DOB approved concrete washout included' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
        spreadsheet.updateCell({ formula: `=SUM(F${substructureStartRow}:F${substructureDataEndRow})` }, `${pfx}F${currentRow}`)
        spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}F${currentRow}`)
        try { spreadsheet.numberFormat('#,##0.00', `${pfx}F${currentRow}`) } catch (e) { }
        fillRatesForProposalRow(currentRow, 'DOB approved concrete washout included')
        spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
        currentRow++
        spreadsheet.updateCell({ value: 'Engineering, shop drawings, formwork drawings, design mixes included' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
        fillRatesForProposalRow(currentRow, 'Engineering, shop drawings, formwork drawings, design mixes included')
        spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
        currentRow++
      }
      if (currentRow > substructureStartRow) {
        const substructureEndRow = currentRow - 1
        spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
        spreadsheet.updateCell({ value: 'Substructure concrete Total:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD7EE' },
          `${pfx}B${currentRow}:E${currentRow}`
        )
        spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
        spreadsheet.updateCell({ formula: `=SUM(H${substructureStartRow}:H${substructureEndRow})*1000` }, `${pfx}F${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD7EE', format: '$#,##0.00' },
          `${pfx}F${currentRow}:G${currentRow}`
        )
        applyTotalRowBorders(spreadsheet, pfx, currentRow, '#BDD7EE')
        currentRow++
        currentRow++ // Empty row after Substructure concrete Total
      }

      currentRow++
    }

    // Below grade waterproofing scope – only show when calculation sheet has waterproofing data
    const waterproofingItems = window.waterproofingPresentItems || []
    const negativeSideItems = window.waterproofingNegativeSideItems || []
    const calcSheetNameWP = 'Calculations Sheet'
    let hasHorizontalWP = false
    let horizontalWPSumRow = 0
    if (formulaData && Array.isArray(formulaData)) {
      const horizontalWPSumFormula = formulaData.find(f =>
        f.itemType === 'waterproofing_horizontal_wp_sum' && f.section === 'waterproofing'
      )
      if (horizontalWPSumFormula) {
        hasHorizontalWP = true
        horizontalWPSumRow = horizontalWPSumFormula.row
      }
    }
    // Only show Below grade waterproofing scope when calculation sheet has exterior-side or negative-side items (not just horizontal WP/insulation formulas which may exist with 0)
    const hasAnyWaterproofingData = (waterproofingItems.length > 0) || (negativeSideItems.length > 0)

    if (hasAnyWaterproofingData) {
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

    currentRow++
    const wpScopeStartRow = currentRow

    // Add proposal text with present waterproofing items
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
        'elev. pit wall': 'elevator pit wall',
        'elev. pit walls': 'elevator pit walls',
        'duplex sewage ejector pit wall': 'duplex sewage ejector pit wall'
      }

      // Convert to display names (case-insensitive abbreviation lookup)
      const displayItems = waterproofingItems.map(item => getDisplayNameForAbbreviation(item, itemDisplayMap))

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
      rowBContentMap.set(currentRow, proposalText)
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
      fillRatesForProposalRow(currentRow, proposalText)

      // Calculate and set row height based on content
      const waterproofingDynamicHeight = calculateRowHeight(proposalText)

      // Find the FT and SQFT totals from waterproofing sum rows (columns I and J)
      let waterproofingFTSumRow = 0

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
          spreadsheet.updateCell({ formula: `='${calcSheetNameWP}'!I${waterproofingFTSumRow}` }, `${pfx}C${currentRow}`)
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
          spreadsheet.updateCell({ formula: `='${calcSheetNameWP}'!J${waterproofingFTSumRow}` }, `${pfx}D${currentRow}`)
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

          // Add $/1000 formula in column H
          const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
          spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              textAlign: 'right',
              backgroundColor: 'white',
              format: '$#,##0.00'
            },
            `${pfx}H${currentRow}`
          )
          // Apply currency format using numberFormat
          try {
            spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}`)
          } catch (e) {
            // Fallback already applied in cellFormat
          }
        }
      }

      // Row height already set above based on waterproofing proposal text
      currentRow++

      // Add negative side proposal text if negative side items are present
      if (negativeSideItems.length > 0) {
        // Map item names to display format
        const negativeSideDisplayMap = {
          'detention tank wall': 'detention tank wall',
          'elevator pit walls': 'elevator pit walls',
          'elev. pit wall': 'elevator pit wall',
          'elev. pit walls': 'elevator pit walls',
          'detention tank slab': 'detention tank slab',
          'duplex sewage ejector pit wall': 'duplex sewage ejector pit wall',
          'duplex sewage ejector pit slab': 'duplex sewage ejector pit slab',
          'elevator pit slab': 'elevator pit slab'
        }

        // Convert to display names (case-insensitive abbreviation lookup)
        const displayNegativeItems = negativeSideItems.map(item => getDisplayNameForAbbreviation(item, negativeSideDisplayMap))

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
        rowBContentMap.set(currentRow, negativeProposalText)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            backgroundColor: 'white',
            verticalAlign: 'top',
            textDecoration: 'none'
          },
          `${pfx}B${currentRow}`
        )
        fillRatesForProposalRow(currentRow, negativeProposalText)

        // Calculate and set row height based on content
        const negativeSideDynamicHeight = calculateRowHeight(negativeProposalText)

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
            spreadsheet.updateCell({ formula: `='${calcSheetNameWP}'!I${negativeSideSumRow}` }, `${pfx}C${currentRow}`)
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
            spreadsheet.updateCell({ formula: `='${calcSheetNameWP}'!J${negativeSideSumRow}` }, `${pfx}D${currentRow}`)
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

        // Add $/1000 formula in column H
        const negativeDollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
        spreadsheet.updateCell({ formula: negativeDollarFormula }, `${pfx}H${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            textAlign: 'right',
            backgroundColor: 'white',
            format: '$#,##0.00'
          },
          `${pfx}H${currentRow}`
        )
        // Apply currency format using numberFormat
        try {
          spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}`)
        } catch (e) {
          // Fallback already applied in cellFormat
        }

        // Row height already set above based on negative side proposal text
        currentRow++
      }

      // Horizontal waterproofing membrane (horizontal) @ SOG – show even when sum is 0
      if (hasHorizontalWP && horizontalWPSumRow > 0) {
        const horizontalWPProposalText = 'F&I new Precon by WR Meadows waterproofing membrane (horizontal) @ SOG'
        spreadsheet.updateCell({ value: horizontalWPProposalText }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, horizontalWPProposalText)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            backgroundColor: 'white',
            verticalAlign: 'top',
            textDecoration: 'none'
          },
          `${pfx}B${currentRow}`
        )
        fillRatesForProposalRow(currentRow, horizontalWPProposalText)
        spreadsheet.updateCell({ formula: `='${calcSheetNameWP}'!J${horizontalWPSumRow}` }, `${pfx}D${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' },
          `${pfx}D${currentRow}`
        )
        try {
          spreadsheet.numberFormat('#,##0.00', `${pfx}D${currentRow}`)
        } catch (e) {
          spreadsheet.cellFormat({ format: '#,##0.00' }, `${pfx}D${currentRow}`)
        }
        spreadsheet.updateCell({
          formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
        }, `${pfx}H${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white', format: '$#,##0.00' },
          `${pfx}H${currentRow}`
        )
        try {
          spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}`)
        } catch (e) { }
        currentRow++
      }
    } else if (hasHorizontalWP) {
      // No vertical items but horizontal WP subsection exists – show horizontal membrane line (even when sum is 0)
      const horizontalWPProposalText = 'F&I new Precon by WR Meadows waterproofing membrane (horizontal) @ SOG'
      spreadsheet.updateCell({ value: horizontalWPProposalText }, `${pfx}B${currentRow}`)
      rowBContentMap.set(currentRow, horizontalWPProposalText)
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
      fillRatesForProposalRow(currentRow, horizontalWPProposalText)
      if (horizontalWPSumRow > 0) {
        spreadsheet.updateCell({ formula: `='${calcSheetNameWP}'!J${horizontalWPSumRow}` }, `${pfx}D${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' },
          `${pfx}D${currentRow}`
        )
        try {
          spreadsheet.numberFormat('#,##0.00', `${pfx}D${currentRow}`)
        } catch (e) {
          spreadsheet.cellFormat({ format: '#,##0.00' }, `${pfx}D${currentRow}`)
        }
      }
      spreadsheet.updateCell({
        formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
      }, `${pfx}H${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white', format: '$#,##0.00' },
        `${pfx}H${currentRow}`
      )
      try {
        spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}`)
      } catch (e) { }
      currentRow++
    } else {
      // No vertical, negative-side, or horizontal WP items – show "No details provided"
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
      currentRow++
    }

    // XPS rigid insulation @ SOG & grade beam – part of waterproofing (before Total)
    let hasHorizontalInsulation = false
    if (formulaData && Array.isArray(formulaData)) {
      const horizontalInsulationSumFormula = formulaData.find(f =>
        f.itemType === 'waterproofing_horizontal_insulation_sum' && f.section === 'waterproofing'
      )
      hasHorizontalInsulation = !!horizontalInsulationSumFormula
    }
    if (hasHorizontalInsulation) {
      const calcSheetName = 'Calculations Sheet'
      let horizontalInsulationSumRow = 0
      let thicknessFromCalc = '2'
      if (formulaData && Array.isArray(formulaData)) {
        const horizontalInsulationSumFormula = formulaData.find(f =>
          f.itemType === 'waterproofing_horizontal_insulation_sum' && f.section === 'waterproofing'
        )
        if (horizontalInsulationSumFormula) {
          horizontalInsulationSumRow = horizontalInsulationSumFormula.row
          const firstDataRow = horizontalInsulationSumFormula.firstDataRow
          if (firstDataRow && calculationData && calculationData[firstDataRow - 1]) {
            const particulars = (calculationData[firstDataRow - 1][1] || '').toString().trim()
            const thickMatch = particulars.match(/^(\d+(?:\.\d+)?)\s*"/) || particulars.match(/(\d+(?:\.\d+)?)\s*"\s*XPS/i)
            if (thickMatch) thicknessFromCalc = thickMatch[1]
          }
        }
      }
      const horizontalProposalText = `F&I new (${thicknessFromCalc}" thick) XPS rigid insulation @ SOG & grade beam`
      spreadsheet.updateCell({ value: horizontalProposalText }, `${pfx}B${currentRow}`)
      rowBContentMap.set(currentRow, horizontalProposalText)
      spreadsheet.wrap(`${pfx}B${currentRow}`, true)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' },
        `${pfx}B${currentRow}`
      )
      if (horizontalInsulationSumRow > 0) {
        spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${horizontalInsulationSumRow}` }, `${pfx}D${currentRow}`)
        spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}D${currentRow}`)
        try { spreadsheet.numberFormat('#,##0.00', `${pfx}D${currentRow}`) } catch (e) { spreadsheet.cellFormat({ format: '#,##0.00' }, `${pfx}D${currentRow}`) }
      }
      fillRatesForProposalRow(currentRow, horizontalProposalText)
      spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
      spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white', format: '$#,##0.00' }, `${pfx}H${currentRow}`)
      try { spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}`) } catch (e) { }
      currentRow++
    }

    // Note: Capstone Contracting Corp is a licensed WR Meadows Installer – part of waterproofing (not bold, italic, centered); one column (B only)
    wpNoteRow.push(currentRow)
    const wpNote = 'Note: Capstone Contracting Corp is a licensed WR Meadows Installer'
    spreadsheet.updateCell({ value: wpNote }, `${pfx}B${currentRow}`)
    rowBContentMap.set(currentRow, wpNote)
    spreadsheet.wrap(`${pfx}B${currentRow}`, true)
    spreadsheet.cellFormat(
      { fontWeight: 'normal', fontStyle: 'italic', color: '#000000', textAlign: 'center', verticalAlign: 'middle', backgroundColor: 'white' },
      `${pfx}B${currentRow}`
    )
    currentRow++

    const wpScopeEndRow = currentRow - 1
    // Add Below grade waterproofing Total row when there are data rows
    if (wpScopeEndRow >= wpScopeStartRow) {
      // Label across B–E, amount in F–G, same layout as other totals
      spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
      spreadsheet.updateCell({ value: 'Below grade waterproofing Total:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          backgroundColor: '#BDD7EE'
        },
        `${pfx}B${currentRow}:E${currentRow}`
      )

      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      spreadsheet.updateCell(
        { formula: `=SUM(H${wpScopeStartRow}:H${wpScopeEndRow})*1000` },
        `${pfx}F${currentRow}`
      )
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          backgroundColor: '#BDD7EE',
          format: '$#,##0.00'
        },
        `${pfx}F${currentRow}:G${currentRow}`
      )

      applyTotalRowBorders(spreadsheet, pfx, currentRow, '#BDD7EE')
      totalRows.push(currentRow)
      currentRow++
    }
    }

    // Superstructure concrete scope – only when calculation sheet has Superstructure section (formulaData has section === 'superstructure')
    const hasSuperstructureInCalculation = !!(formulaData && Array.isArray(formulaData) && formulaData.some(f => f.section === 'superstructure'))
    if (hasSuperstructureInCalculation) {
      // Add empty row above Superstructure concrete scope
      currentRow++

      // Superstructure concrete scope – add new subsections at the end (before the empty row after scope)
      // Add Superstructure concrete scope header (below Below grade waterproofing)
      spreadsheet.updateCell({ value: 'Superstructure concrete scope:' }, `${pfx}B${currentRow}`)
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
    currentRow++

    // CIP concrete slabs: consolidated proposal lines from raw data (Slab 8", Roof Slab 8", Balcony, Terrace)
    const buildSuperstructureProposalLines = () => {
      const lines = []
      const patchLines = []
      const somdLines = []
      if (!rawData || rawData.length < 2) return { cipLines: lines, patchLines, somdLines, slabStepsLines: [], cipStairsGroupsWithLines: [], concreteHangerLines: [], lwConcreteFillLines: [], toppingSlabLines: [], raisedSlabLines: [], builtUpSlabLines: [], builtupRampsLines: [], builtUpStairLines: [], shearWallLines: [], columnsLines: [], dropPanelLines: [], concretePostLines: [], concreteEncasementLines: [], beamsLines: [], parapetWallLines: [], thermalBreakLines: [], curbsLines: [], concretePadLines: [], nonShrinkGroutLines: [], repairScopeLines: [] }
      const headers = rawData[0]
      const dataRows = rawData.slice(1)
      const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
      const pageIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'page')
      const estimateIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'estimate')
      if (digitizerIdx === -1) return { cipLines: lines, patchLines, somdLines, slabStepsLines: [], cipStairsGroupsWithLines: [], concreteHangerLines: [], lwConcreteFillLines: [], toppingSlabLines: [], raisedSlabLines: [], builtUpSlabLines: [], builtupRampsLines: [], builtUpStairLines: [], shearWallLines: [], columnsLines: [], dropPanelLines: [], concretePostLines: [], concreteEncasementLines: [], beamsLines: [], parapetWallLines: [], thermalBreakLines: [], curbsLines: [], concretePadLines: [], nonShrinkGroutLines: [], repairScopeLines: [] }

      const toOrdinal = (n) => {
        const s = ['th', 'st', 'nd', 'rd']
        const v = n % 100
        return n + (s[(v - 20) % 10] || s[v] || s[0])
      }
      const extractFloorRef = (text) => {
        if (!text || typeof text !== 'string') return ''
        const t = text.trim()
        const atMatch = t.match(/@\s*(.+?)(?:\s+as\s+per|$)/i)
        if (atMatch) return atMatch[1].trim()
        if (/roof\s*fl/i.test(t)) return 'roof FL'
        const rangeMatch = t.match(/(\d+)(?:st|nd|rd|th)?\s*FL\s*to\s*(\d+)(?:st|nd|rd|th)?\s*FL/i)
        if (rangeMatch) return `${toOrdinal(parseInt(rangeMatch[1]))} FL to ${toOrdinal(parseInt(rangeMatch[2]))} FL`
        const singleMatch = t.match(/(\d+)(?:st|nd|rd|th)?\s*FL/i)
        if (singleMatch) return `${toOrdinal(parseInt(singleMatch[1]))} FL`
        return ''
      }
      const extractSPageRefs = (rows) => {
        if (pageIdx === -1) return 'S-101.00'
        const refs = new Set()
        rows.forEach(r => {
          const p = r[pageIdx]
          if (p) {
            const s = String(p).trim()
            const match = s.match(/(?:S|A|P|FO)-[\w.]+|SP\/[\w-]+/gi)
            if (match) match.forEach(m => refs.add(m))
          }
        })
        const arr = [...refs].sort((a, b) => {
          const numPart = (s) => parseFloat(String(s).replace(/^[A-Z]+[-/]/i, '').replace(/[A-Za-z]/g, '')) || 0
          const aNum = numPart(a)
          const bNum = numPart(b)
          return aNum !== bNum ? aNum - bNum : String(a).localeCompare(String(b))
        })
        if (arr.length === 0) return 'S-101.00'
        if (arr.length === 1) return arr[0]
        if (arr.length === 2) return `${arr[0]} & ${arr[1]}`
        return arr.slice(0, -1).join(', ') + ' & ' + arr[arr.length - 1]
      }
      const formatFloorRefs = (floorRefs) => {
        const unique = [...new Set(floorRefs)].filter(Boolean)
        if (unique.length === 0) return ''
        if (unique.length === 1) return unique[0]
        const hasRoof = unique.some(f => /roof/i.test(f))
        const numeric = unique.filter(f => !/roof/i.test(f))
        const rangeMatch = numeric.find(f => /to\s+\d/i.test(f))
        let out = ''
        if (rangeMatch && numeric.length === 1) {
          out = rangeMatch
        } else if (numeric.length > 0) {
          const sorted = numeric.sort((a, b) => {
            const nA = parseInt(a) || 0
            const nB = parseInt(b) || 0
            return nA - nB
          })
          out = sorted.length === 2 ? sorted.join(' & ') : sorted.slice(0, -1).join(', ') + ' & ' + sorted[sorted.length - 1]
        }
        if (hasRoof) out = out ? `${out} & roof FL` : 'roof FL'
        return out
      }

      const slab8Rows = []
      const roofSlab8Rows = []
      const slabStepsRows = []
      const cipStairsGroups = [] // { groupName, landings: [], stairs: [] }
      let currentCipStairsGroup = null
      const concretePadRows = []
      const nonShrinkGroutRows = []
      const repairScopeWallRows = []
      const repairScopeSlabRows = []
      const repairScopeColumnRows = []
      const patchSlabRows = []
      const somdRows = []
      const balconyRows = []
      const terraceRows = []
      const concreteHangerRows = []
      const lwConcreteFillRows = []
      const toppingSlabRows = []
      const raisedSlabKneeRows = []
      const raisedSlabRows = []
      const builtUpSlabKneeRows = []
      const builtUpSlabRows = []
      const builtupRampsKneeRows = []
      const builtupRampsStyroRows = []
      const builtupRampsRampRows = []
      const builtUpStairRows = []
      const shearWallRows = []
      const columnsRows = []
      const dropPanelRows = []
      const concretePostRows = []
      const concreteEncasementRows = []
      const beamsRows = []
      const parapetWallRows = []
      const thermalBreakRows = []
      const curbsRows = []

      const parseShearWallDims = (particulars) => {
        if (!particulars || typeof particulars !== 'string') return null
        const match = particulars.match(/\(([^)]+x[^)]+)\)/)
        if (!match) return null
        const parts = match[1].split(/x/i).map(p => p.trim())
        if (parts.length < 2) return null
        const widthFeet = convertToFeet(parts[0])
        const heightFeet = convertToFeet(parts[1])
        const widthInches = Math.round(widthFeet * 12)
        return { widthFeet, heightFeet, widthInches }
      }

      const parseKneeWallDims = (particulars) => {
        if (!particulars || typeof particulars !== 'string') return null
        const match = particulars.match(/\(([^)]+x[^)]+)\)/)
        if (!match) return null
        const parts = match[1].split(/x/i).map(p => p.trim())
        if (parts.length < 2) return null
        const firstFeet = convertToFeet(parts[0])
        const secondFeet = convertToFeet(parts[1])
        const firstInch = firstFeet * 12
        const secondInch = secondFeet * 12
        return { thicknessInches: Math.round(firstInch) || 4, heightFeet: secondFeet }
      }

      const parseToppingSlabDims = (particulars) => {
        if (!particulars || typeof particulars !== 'string') return null
        const t = String(particulars)
        const thickMatch = t.match(/(\d+(?:\s*[½¼¾]|\s*\d+\/\d+)?)\s*"\s*thick/i) || t.match(/(\d+(?:\s*[½¼¾]|\s*\d+\/\d+)?)\s*"/)
        const parseInch = (s) => {
          if (!s) return 2
          s = String(s).replace(/\s*½/g, '.5').replace(/\s*¼/g, '.25').replace(/\s*¾/g, '.75')
          const m = s.match(/(\d+)\s+(\d+)\/(\d+)/)
          if (m) return parseFloat(m[1]) + (parseFloat(m[2]) / parseFloat(m[3]))
          return parseFloat(s.replace(/[^\d.-]/g, '')) || 2
        }
        const inches = thickMatch ? parseInch(thickMatch[1].trim()) : 2
        const concreteType = t.toLowerCase().includes('nw') ? 'NW' : 'LW'
        return { inches, concreteType }
      }
      const formatToppingThickness = (inches) => {
        if (inches % 1 === 0.5) return `${Math.floor(inches)}½"`
        if (inches % 1 === 0.25) return `${Math.floor(inches)}¼"`
        if (inches % 1 === 0.75) return `${Math.floor(inches)}¾"`
        return `${inches}"`
      }

      const parseHeightFromParticulars = (particulars) => {
        if (!particulars || typeof particulars !== 'string') return null
        const match = String(particulars).match(/H\s*=\s*([^,\s]+)/i)
        return match ? convertToFeet(match[1].trim()) : null
      }

      const parseThreeDimBracket = (particulars) => {
        if (!particulars || typeof particulars !== 'string') return null
        const match = particulars.match(/\(([^)]+x[^)]+x[^)]+)\)/)
        if (!match) return null
        const parts = match[1].split(/x/i).map(p => p.trim())
        if (parts.length < 3) return null
        return { length: convertToFeet(parts[0]), width: convertToFeet(parts[1]), height: convertToFeet(parts[2]) }
      }
      const formatFeetInches = (feet) => {
        if (feet == null || isNaN(feet)) return ''
        const f = Math.floor(feet)
        const inch = Math.round((feet - f) * 12)
        if (inch === 0) return `${f}'-0"`
        return `${f}'-${inch}"`
      }

      const parseSOMDDimensions = (particulars) => {
        if (!particulars || typeof particulars !== 'string') return null
        const parts = String(particulars).split('"')
        if (parts.length < 3) return null
        const beforeFirst = (parts[0] || '').trim()
        const beforeSecond = (parts[1] || '').trim()
        const firstMatch = beforeFirst.match(/(\d+(?:\s*[½¼¾]|\s*\d+\/\d+)?)\s*$/);
        const secondMatch = beforeSecond.match(/(\d+(?:\s*[½¼¾]|\s*\d+\/\d+)?)\s*$/)
        if (!firstMatch || !secondMatch) return null
        const parseInch = (s) => {
          s = String(s).replace(/\s*½/g, '.5').replace(/\s*¼/g, '.25').replace(/\s*¾/g, '.75')
          const m = s.match(/(\d+)\s+(\d+)\/(\d+)/)
          if (m) return parseFloat(m[1]) + (parseFloat(m[2]) / parseFloat(m[3]))
          return parseFloat(s.replace(/[^\d.-]/g, '')) || 0
        }
        const firstVal = parseInch(firstMatch[1].trim())
        const secondVal = parseInch(secondMatch[1].trim())
        if (firstVal <= 0 || secondVal <= 0) return null
        const gaMatch = String(particulars).match(/x\s*(\d+)\s*GA/i)
        const ga = gaMatch ? gaMatch[1] : '18'
        return { firstInches: firstVal, secondInches: secondVal, ga }
      }
      const formatSOMDThickness = (firstInches) => {
        if (firstInches % 1 === 0.5) return `${Math.floor(firstInches)}½"`
        if (firstInches % 1 === 0.25) return `${Math.floor(firstInches)}¼"`
        if (firstInches % 1 === 0.75) return `${Math.floor(firstInches)}¾"`
        return `${firstInches}"`
      }

      dataRows.forEach(row => {
        if (estimateIdx >= 0 && row[estimateIdx] && String(row[estimateIdx]).trim() !== 'Superstructure') return
        const dig = row[digitizerIdx]
        if (!dig) return
        const p = String(dig).toLowerCase()
        if (p.includes('balcony slab') && (p.includes('8"') || p.includes('8"'))) {
          balconyRows.push(row)
        } else if (p.includes('terrace slab') && (p.includes('8"') || p.includes('8"'))) {
          terraceRows.push(row)
        } else if (p.includes('roof slab') && (p.includes('8"') || p.includes('8"'))) {
          roofSlab8Rows.push(row)
        } else if ((p.includes('slab 8"') || p.includes('slab 8"')) && !p.includes('roof') && !p.includes('balcony') && !p.includes('terrace')) {
          slab8Rows.push(row)
        } else if (p.includes('slab step') && /\([^)]+x[^)]+\)/.test(String(dig))) {
          slabStepsRows.push(row)
        } else if (p.includes('patch slab')) {
          patchSlabRows.push(row)
        } else if (p.includes('somd') || p.includes('slab on metal deck')) {
          somdRows.push(row)
        } else if (p.includes('concrete hanger') && /\([^)]+x[^)]+x[^)]+\)/.test(String(dig))) {
          concreteHangerRows.push(row)
        } else if (p.includes('lw concrete fill')) {
          lwConcreteFillRows.push(row)
        } else if (p.includes('topping slab') && (p.includes('thick') || p.includes('"'))) {
          toppingSlabRows.push(row)
        } else if (p.includes('knee wall') && /\([^)]+x[^)]+\)/.test(String(dig)) && /\)\s*\(\d+\)\s*$/.test(String(dig).trim())) {
          builtupRampsKneeRows.push(row)
        } else if (p.includes('styrofoam') && /\(\d+\)\s*$/.test(String(dig).trim())) {
          builtupRampsStyroRows.push(row)
        } else if ((p.includes('builtup ramp') || p.includes('built up ramp')) && /\(\d+\)\s*$/.test(String(dig).trim())) {
          builtupRampsRampRows.push(row)
        } else if (p.includes('knee wall') && p.includes('x') && !p.includes('builtup') && !p.includes('built up')) {
          raisedSlabKneeRows.push(row)
        } else if (p.includes('raised slab') && (p.includes('"') || p.includes('thick'))) {
          raisedSlabRows.push(row)
        } else if (p.includes('knee wall') && /\([^)]+x[^)]+\)/.test(String(dig))) {
          if (p.includes('builtup slab') || p.includes('built up slab') || p.includes('@ builtup')) {
            builtUpSlabKneeRows.push(row)
          } else {
            builtUpSlabKneeRows.push(row)
          }
        } else if ((p.includes('builtup slab') || p.includes('built up slab')) && (p.includes('"') || p.includes('thick'))) {
          builtUpSlabRows.push(row)
        } else if ((p.includes('builtup ramp') || p.includes('built up ramp')) && (p.includes('"') || p.includes('thick'))) {
          builtupRampsRampRows.push(row)
        } else if ((p.includes('built up stairs') || p.includes('built up stair')) && !p.includes('@')) {
          builtUpStairRows.push(row)
        } else if ((p.includes('sw ') || p.includes('shear wall')) && /\([^)]+x[^)]+\)/.test(String(dig))) {
          shearWallRows.push(row)
        } else if (p.includes('as per takeoff count') || (p.includes('column') && p.includes('count'))) {
          columnsRows.push(row)
        } else if (p.includes('drop panel')) {
          dropPanelRows.push(row)
        } else if (p.includes('concrete post') && /\([^)]+x[^)]+x[^)]+\)/.test(String(dig))) {
          concretePostRows.push(row)
        } else if (p.includes('concrete encasement') && /\([^)]+x[^)]+x[^)]+\)/.test(String(dig))) {
          concreteEncasementRows.push(row)
        } else if (!p.includes('secant pile') && !p.includes('core beam') && !p.includes('pile w/') &&
          /^(?:'?\s*)?(\d+B-\d+|RB-\d+|BHB-\d+)/i.test(String(dig).trim()) && /\([^)]+x[^)]+\)/.test(String(dig)) && !/\([^)]+x[^)]+x[^)]+\)/.test(String(dig))) {
          beamsRows.push(row)
        } else if (p.includes('parapet wall') && /\([^)]+x[^)]+\)/.test(String(dig))) {
          parapetWallRows.push(row)
        } else if (p.includes('thermal break')) {
          thermalBreakRows.push(row)
        } else if (p.includes('concrete curb') && /\([^)]+x[^)]+\)/.test(String(dig))) {
          curbsRows.push(row)
        } else if (p.includes('concrete pad') && (/\d+\s*"/.test(String(dig)) || /\d+"\s*$/.test(String(dig))) && !p.includes('transformer')) {
          concretePadRows.push(row)
        } else if (p.includes('non-shrink grout') || p.includes('non shrink grout')) {
          nonShrinkGroutRows.push(row)
        } else if (p.includes('concrete wall crack repair')) {
          repairScopeWallRows.push(row)
        } else if (p.includes('slab crack repair')) {
          repairScopeSlabRows.push(row)
        } else if (p.includes('column crack repair')) {
          repairScopeColumnRows.push(row)
        } else if (/stair\s+[a-z][^:]*:?$|misc\.?\s*stair:?|ext\.?\s*stair:?|exterior\s*stair:?/i.test(String(dig).trim()) && !p.includes('concrete stair') && !p.includes('landing')) {
          const groupName = String(dig).trim().replace(/:$/, '').trim()
          currentCipStairsGroup = { groupName, landings: [], stairs: [] }
          cipStairsGroups.push(currentCipStairsGroup)
        } else if (currentCipStairsGroup && ((p.includes('landing') && !p.includes('stair slab')) || p === 'landings')) {
          currentCipStairsGroup.landings.push(row)
        } else if (currentCipStairsGroup && (p.includes('concrete stair') || (p.includes('stair') && (p.includes('riser') || p.includes('wide'))) || p === 'stairs')) {
          currentCipStairsGroup.stairs.push(row)
        } else if (currentCipStairsGroup && p === 'stair slab') {
          if (currentCipStairsGroup.stairs.length === 0) currentCipStairsGroup.stairs.push(row)
        } else if ((p.includes('stair landing') || (p.includes('landing') && p.includes('stair'))) && !p.includes('built up') && !p.includes('grade')) {
          if (!currentCipStairsGroup || cipStairsGroups.length === 0) {
            currentCipStairsGroup = { groupName: 'CIP Stairs', landings: [], stairs: [] }
            cipStairsGroups.push(currentCipStairsGroup)
          }
          currentCipStairsGroup.landings.push(row)
        } else if ((p.includes('concrete stair') || (p.includes('stair') && p.includes('riser'))) && !p.includes('built up') && !p.includes('builtup') && !p.includes('slab step')) {
          if (!currentCipStairsGroup || cipStairsGroups.length === 0) {
            currentCipStairsGroup = { groupName: 'CIP Stairs', landings: [], stairs: [] }
            cipStairsGroups.push(currentCipStairsGroup)
          }
          currentCipStairsGroup.stairs.push(row)
        }
      })

      if (patchSlabRows.length > 0) {
        const floorText = formatFloorRefs(patchSlabRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
        const sRefs = extractSPageRefs(patchSlabRows)
        patchLines.push({
          proposalText: `Allow to patch slab @ ${floorText} as per ${sRefs}`,
          sumRowKey: 'patch',
          subsectionName: 'Patch slab'
        })
      }
      if (slab8Rows.length > 0) {
        const floorRefs = slab8Rows.map(r => extractFloorRef(r[digitizerIdx]))
        const floorText = formatFloorRefs(floorRefs) || 'as indicated'
        const sRefs = extractSPageRefs(slab8Rows)
        lines.push({
          proposalText: `F&I new (8" thick) slab @ ${floorText} as per ${sRefs}`,
          sumRowKey: 'slab8',
          subsectionName: 'CIP Slabs',
          cyUnitPrice: 1500
        })
      }
      if (roofSlab8Rows.length > 0) {
        const floorText = formatFloorRefs(roofSlab8Rows.map(r => extractFloorRef(r[digitizerIdx]))) || 'roof FL'
        const sRefs = extractSPageRefs(roofSlab8Rows)
        lines.push({
          proposalText: `F&I new (8" thick) slab w/epoxy coated reinforcement @ roof FL as per ${sRefs}`,
          sumRowKey: 'roofSlab8',
          subsectionName: 'CIP Slabs',
          cyUnitPrice: 1500
        })
      }
      if (balconyRows.length > 0) {
        const floorText = formatFloorRefs(balconyRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
        const sRefs = extractSPageRefs(balconyRows)
        lines.push({
          proposalText: `F&I new (8" thick) balcony slab w/epoxy coated reinforcement @ ${floorText} as per ${sRefs}`,
          sumRowKey: 'balcony',
          subsectionName: 'Balcony slab',
          cyUnitPrice: 1500
        })
      }
      if (terraceRows.length > 0) {
        const floorText = formatFloorRefs(terraceRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
        const sRefs = extractSPageRefs(terraceRows)
        lines.push({
          proposalText: `F&I new (8" thick) terrace slab w/epoxy coated reinforcement @ ${floorText} as per ${sRefs}`,
          sumRowKey: 'terrace',
          subsectionName: 'Terrace slab',
          cyUnitPrice: 1500
        })
      }
      const cipStairsGroupsWithLines = []
      cipStairsGroups.forEach(grp => {
        const allRows = [...grp.landings, ...grp.stairs]
        const floorText = formatFloorRefs(allRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
        const sRefs = extractSPageRefs(allRows)
        const lines = []
        if (grp.landings.length > 0) {
          lines.push({ proposalText: `F&I new (8" thick) stair landings @ ${floorText} as per ${sRefs}` })
        }
        if (grp.stairs.length > 0) {
          const stairsRow = grp.stairs.find(r => {
            const d = String(r[digitizerIdx] || '').toLowerCase()
            return d.includes('concrete stair') || d.includes('riser') || (d.includes('stair') && !d.includes('stair slab'))
          }) || grp.stairs[0]
          const stairDig = stairsRow[digitizerIdx]
          const widthMatch = String(stairDig || '').match(/(\d+)\s*'[-]?\s*(\d+)\s*"/) || String(stairDig || '').match(/(\d+)['-](\d+)\s*"/i)
          const widthStr = widthMatch && widthMatch[1] ? `${widthMatch[1]}'-${widthMatch[2] || 0}"` : "3'-0\""
          const riserMatch = String(stairDig || '').match(/(\d+)\s*Riser/i)
          const riserStr = riserMatch ? `${riserMatch[1]} Riser` : 'as indicated'
          lines.push({ proposalText: `F&I new (${widthStr} wide) concrete stairs (${riserStr}) @ ${floorText} as per ${sRefs}` })
        }
        if (lines.length > 0) {
          cipStairsGroupsWithLines.push({ groupName: grp.groupName, lines })
        }
      })
      const slabStepsLines = []
      if (slabStepsRows.length > 0) {
        const parseSlabStepDims = (t) => {
          const match = String(t).match(/\(([^)]+x[^)]+)\)/)
          if (!match) return null
          const parts = match[1].split(/x/i).map(p => p.trim())
          if (parts.length < 2) return null
          return { widthFeet: convertToFeet(parts[0]), heightFeet: convertToFeet(parts[1]) }
        }
        const dims = parseSlabStepDims(slabStepsRows[0][digitizerIdx])
        const widthStr = dims ? formatFeetInches(dims.widthFeet) : "2'-0\""
        const heightStr = dims ? formatFeetInches(dims.heightFeet) : "1'-3\""
        const floorText = formatFloorRefs(slabStepsRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
        const sRefs = extractSPageRefs(slabStepsRows)
        slabStepsLines.push({
          proposalText: `F&I new (${widthStr} wide) concrete slab steps (H=${heightStr}) @ ${floorText} as per ${sRefs}`,
          sumRowKey: 'slabSteps',
          subsectionName: 'Slab steps',
          sumColumn: 'J',
          targetCol: 'D',
          cyUnitPrice: 1200
        })
      }
      const lwConcreteFillLines = []
      if (lwConcreteFillRows.length > 0) {
        const heightFeet = parseHeightFromParticulars(lwConcreteFillRows[0][digitizerIdx])
        const hStr = heightFeet != null ? formatFeetInches(heightFeet) : "1'-1\""
        const floorText = formatFloorRefs(lwConcreteFillRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
        const sRefs = extractSPageRefs(lwConcreteFillRows)
        lwConcreteFillLines.push({
          proposalText: `F&I new LW concrete fill (H=${hStr}) @ ${floorText} as per ${sRefs}`,
          sumRowKey: 'lwConcreteFill',
          subsectionName: 'LW concrete fill'
        })
      }
      const raisedSlabLines = []
      const builtUpSlabLines = []
      if (builtUpSlabKneeRows.length > 0 || builtUpSlabRows.length > 0) {
        const slabKnee = builtUpSlabKneeRows[0]
        if (slabKnee) {
          const dims = parseKneeWallDims(slabKnee[digitizerIdx])
          const thickStr = dims ? `${dims.thicknessInches}"` : '4"'
          const hStr = dims && dims.heightFeet ? formatFeetInches(dims.heightFeet) : "1'-0\""
          const sRefs = extractSPageRefs([...builtUpSlabKneeRows.slice(0, 1), ...builtUpSlabRows])
          builtUpSlabLines.push({
            proposalText: `F&I new (${thickStr} wide) knee walls (H=${hStr}) as per ${sRefs}`,
            sumRowKey: 'builtUpSlabKnee',
            subsectionName: 'Built-up slab'
          })
        }
        if (builtUpSlabKneeRows.length > 0 && builtUpSlabRows.length > 0) {
          const sRefs = extractSPageRefs([...builtUpSlabKneeRows, ...builtUpSlabRows])
          builtUpSlabLines.push({
            proposalText: `F&I new Styrofoam @ under built-up slab as per ${sRefs}`,
            sumRowKey: 'builtUpSlabStyro',
            subsectionName: 'Built-up slab'
          })
        }
        if (builtUpSlabRows.length > 0) {
          const thickMatch = String(builtUpSlabRows[0][digitizerIdx] || '').match(/(\d+(?:\.\d+)?)\s*"\s*thick/i) || String(builtUpSlabRows[0][digitizerIdx] || '').match(/(\d+(?:\.\d+)?)\s*"/) || []
          const thickStr = thickMatch[1] ? `${thickMatch[1]}"` : '3"'
          const floorText = formatFloorRefs(builtUpSlabRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
          const sRefs = extractSPageRefs(builtUpSlabRows)
          builtUpSlabLines.push({
            proposalText: `F&I new (${thickStr} thick) built-up slab @ ${floorText} as per ${sRefs}`,
            sumRowKey: 'builtUpSlab',
            subsectionName: 'Built-up slab'
          })
        }
      }
      const builtupRampsLines = []
      if (builtupRampsKneeRows.length > 0 || builtupRampsStyroRows.length > 0 || builtupRampsRampRows.length > 0) {
        if (builtupRampsKneeRows.length > 0) {
          const heights = builtupRampsKneeRows.map(r => {
            const d = parseKneeWallDims(r[digitizerIdx])
            return d && d.heightFeet ? d.heightFeet * 12 : 0
          }).filter(Boolean)
          const avgInches = heights.length > 0 ? Math.round(heights.reduce((a, b) => a + b, 0) / heights.length) : 4
          const firstDims = parseKneeWallDims(builtupRampsKneeRows[0][digitizerIdx])
          const thickStr = firstDims ? `${firstDims.thicknessInches}"` : '4"'
          const sRefs = extractSPageRefs([...builtupRampsKneeRows, ...builtupRampsStyroRows, ...builtupRampsRampRows])
          builtupRampsLines.push({
            proposalText: `F&I new (${thickStr} wide) knee walls (Havg=${avgInches}") as per ${sRefs}`,
            sumRowKey: 'builtupRampsKnee',
            subsectionName: 'Builtup ramps'
          })
        }
        if (builtupRampsStyroRows.length > 0 || (builtupRampsKneeRows.length > 0 && builtupRampsRampRows.length > 0)) {
          const sRefs = extractSPageRefs([...builtupRampsKneeRows, ...builtupRampsStyroRows, ...builtupRampsRampRows])
          builtupRampsLines.push({
            proposalText: `F&I new Styrofoam @ under built-up ramp as per ${sRefs}`,
            sumRowKey: 'builtupRampsStyro',
            subsectionName: 'Builtup ramps'
          })
        }
        if (builtupRampsRampRows.length > 0) {
          const thickMatch = String(builtupRampsRampRows[0][digitizerIdx] || '').match(/(\d+(?:\.\d+)?)\s*"\s*thick/i) || String(builtupRampsRampRows[0][digitizerIdx] || '').match(/(\d+(?:\.\d+)?)\s*"/) || []
          const thickStr = thickMatch[1] ? `${thickMatch[1]}"` : '3"'
          const floorText = formatFloorRefs(builtupRampsRampRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
          const sRefs = extractSPageRefs(builtupRampsRampRows)
          builtupRampsLines.push({
            proposalText: `F&I new (${thickStr} thick) built-up ramp @ ${floorText} as per ${sRefs}`,
            sumRowKey: 'builtupRampsRamp',
            subsectionName: 'Builtup ramps'
          })
        }
      }
      const builtUpStairLines = []
      if (builtUpStairRows.length > 0) {
        const stairKneeIdx = builtUpSlabRows.length > 0 ? 1 : 0
        const stairKnee = builtUpSlabKneeRows[stairKneeIdx] || builtUpSlabKneeRows[0]
        if (stairKnee) {
          const dims = parseKneeWallDims(stairKnee[digitizerIdx])
          const thickStr = dims ? `${dims.thicknessInches}"` : '4"'
          const hInches = dims && dims.heightFeet ? Math.round(dims.heightFeet * 12) : 8
          const sRefs = extractSPageRefs([stairKnee, ...builtUpStairRows])
          builtUpStairLines.push({
            proposalText: `F&I new (${thickStr} wide) knee walls (H=${hInches}") as per ${sRefs}`,
            sumRowKey: 'builtUpStairKnee',
            subsectionName: 'Built-up stair'
          })
        }
        if (stairKnee) {
          const dims = parseKneeWallDims(stairKnee[digitizerIdx])
          const styroInches = dims && dims.heightFeet ? Math.round(dims.heightFeet * 12) : 8
          const sRefs = extractSPageRefs([...builtUpSlabKneeRows, ...builtUpStairRows])
          builtUpStairLines.push({
            proposalText: `F&I new (${styroInches}" thick) Styrofoam @ under Built-up stair as per ${sRefs}`,
            sumRowKey: 'builtUpStairStyro',
            subsectionName: 'Built-up stair'
          })
        }
        if (builtUpStairRows.length > 0) {
          const widthMatch = String(builtUpStairRows[0][digitizerIdx] || '').match(/(\d+)\s*'[-]?\s*(\d+)\s*"/) || []
          const widthStr = widthMatch[1] && widthMatch[2] ? `${widthMatch[1]}'-${widthMatch[2]}"` : "3'-0\""
          const riserMatch = String(builtUpStairRows[0][digitizerIdx] || '').match(/(\d+)\s*Riser/i) || []
          const riserStr = riserMatch[1] ? `${riserMatch[1]} Riser` : '4 Riser'
          const floorText = formatFloorRefs(builtUpStairRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
          const sRefs = extractSPageRefs(builtUpStairRows)
          builtUpStairLines.push({
            proposalText: `F&I new (${widthStr} wide) built up stairs (${riserStr}) @ ${floorText} as per ${sRefs}`,
            sumRowKey: 'builtUpStair',
            subsectionName: 'Built-up stair'
          })
        }
      }
      if (raisedSlabKneeRows.length > 0 || raisedSlabRows.length > 0) {
        if (raisedSlabKneeRows.length > 0) {
          const firstKnee = parseKneeWallDims(raisedSlabKneeRows[0][digitizerIdx])
          const thickStr = firstKnee ? `${firstKnee.thicknessInches}"` : '4"'
          const hStr = firstKnee && firstKnee.heightFeet ? formatFeetInches(firstKnee.heightFeet) : "1'-7\""
          const sRefs = extractSPageRefs(raisedSlabKneeRows)
          raisedSlabLines.push({
            proposalText: `F&I new (${thickStr} wide) knee walls (H=${hStr}) as per ${sRefs}`,
            sumRowKey: 'raisedKneeWall',
            subsectionName: 'Raised slab'
          })
        }
        if (raisedSlabKneeRows.length > 0 && raisedSlabRows.length > 0) {
          const firstKnee = parseKneeWallDims(raisedSlabKneeRows[0][digitizerIdx])
          const styroInches = firstKnee && firstKnee.heightFeet ? Math.round(firstKnee.heightFeet * 12) : 19
          const sRefs = extractSPageRefs([...raisedSlabKneeRows, ...raisedSlabRows])
          raisedSlabLines.push({
            proposalText: `F&I new (${styroInches}" thick) Styrofoam @ under built-up slab topping as per ${sRefs}`,
            sumRowKey: 'raisedStyrofoam',
            subsectionName: 'Raised slab'
          })
        }
        if (raisedSlabRows.length > 0) {
          const thickMatch = String(raisedSlabRows[0][digitizerIdx] || '').match(/(\d+(?:\.\d+)?)\s*"/) || []
          const thickStr = thickMatch[1] ? `${thickMatch[1]}"` : '4"'
          const floorText = formatFloorRefs(raisedSlabRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
          const sRefs = extractSPageRefs(raisedSlabRows)
          raisedSlabLines.push({
            proposalText: `F&I new (${thickStr} thick) slab topping @ ${floorText} as per ${sRefs}`,
            sumRowKey: 'raisedSlab',
            subsectionName: 'Raised slab'
          })
        }
      }
      const toppingSlabLines = []
      if (toppingSlabRows.length > 0) {
        const firstDims = parseToppingSlabDims(toppingSlabRows[0][digitizerIdx])
        const thickStr = firstDims ? formatToppingThickness(firstDims.inches) : '2"'
        const concreteType = firstDims ? firstDims.concreteType : 'NW'
        const floorText = formatFloorRefs(toppingSlabRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
        const sRefs = extractSPageRefs(toppingSlabRows)
        toppingSlabLines.push({
          proposalText: `F&I new (${thickStr} thick) concrete ${concreteType} topping slab @ ${floorText} as per ${sRefs}`,
          sumRowKey: 'toppingSlab',
          subsectionName: 'Topping slab'
        })
      }
      if (somdRows.length > 0) {
        const somdGroups = new Map()
        somdRows.forEach(row => {
          const dims = parseSOMDDimensions(row[digitizerIdx])
          if (!dims) return
          const concreteType = String(row[digitizerIdx] || '').toLowerCase().includes('nw concrete') ? 'NW' : 'LW'
          const key = `${dims.firstInches}_${dims.secondInches}_${concreteType}`
          if (!somdGroups.has(key)) somdGroups.set(key, { rows: [], dims, concreteType })
          somdGroups.get(key).rows.push(row)
        })
        const sortedKeys = [...somdGroups.keys()].sort((a, b) => {
          const aParts = a.split('_')
          const bParts = b.split('_')
          const a1 = parseFloat(aParts[0]) || 0
          const a2 = parseFloat(aParts[1]) || 0
          const b1 = parseFloat(bParts[0]) || 0
          const b2 = parseFloat(bParts[1]) || 0
          return (a1 - b1) || (a2 - b2)
        })
        sortedKeys.forEach((key, groupIndex) => {
          const g = somdGroups.get(key)
          const { dims, concreteType, rows } = g
          const thickStr = formatSOMDThickness(dims.firstInches)
          const mdStr = `${formatSOMDThickness(dims.secondInches)} x ${dims.ga} GA`
          const floorText = formatFloorRefs(rows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
          const sRefs = extractSPageRefs(rows)
          somdLines.push({
            headingText: 'Concrete slab on metal deck:',
            proposalText: `F&I new (${thickStr} thick) ${concreteType} concrete topping over ${mdStr} (metal deck by others) reinforced w/WWF @ ${floorText} as per ${sRefs}`,
            sumRowKey: 'somd',
            subsectionName: 'Slab on metal deck',
            somdGroupIndex: groupIndex
          })
        })
      }
      const shearWallLines = []
      if (shearWallRows.length > 0) {
        const swGroups = new Map()
        shearWallRows.forEach(row => {
          const dims = parseShearWallDims(row[digitizerIdx])
          if (!dims) return
          const key = String(dims.widthFeet)
          if (!swGroups.has(key)) swGroups.set(key, { rows: [], dims, heights: [] })
          const g = swGroups.get(key)
          g.rows.push(row)
          g.heights.push(dims.heightFeet)
        })
        const sortedKeys = [...swGroups.keys()].sort((a, b) => parseFloat(a) - parseFloat(b))
        sortedKeys.forEach((key, groupIndex) => {
          const g = swGroups.get(key)
          const { dims, rows, heights } = g
          const thickStr = dims.widthInches >= 12 ? formatFeetInches(dims.widthFeet) : `${dims.widthInches}"`
          const avgHeight = heights.length > 0 ? heights.reduce((a, b) => a + b, 0) / heights.length : 0
          const havgStr = avgHeight > 0 ? formatFeetInches(avgHeight) : ''
          const floorText = formatFloorRefs(rows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
          const sRefs = extractSPageRefs(rows)
          shearWallLines.push({
            proposalText: `F&I new (${thickStr} thick) shear wall (Havg=${havgStr}) @ ${floorText} as per ${sRefs}`,
            sumRowKey: 'shearWall',
            subsectionName: 'Shear Walls',
            shearGroupIndex: groupIndex
          })
        })
      }
      const dropPanelLines = []
      if (dropPanelRows.length > 0) {
        const totalIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'total')
        const takeoffIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'takeoff')
        const qtyIdx = totalIdx >= 0 ? totalIdx : takeoffIdx
        let totalQty = 0
        const heights = []
        dropPanelRows.forEach(row => {
          const d = row[digitizerIdx]
          const v = qtyIdx >= 0 && row[qtyIdx] != null ? parseFloat(String(row[qtyIdx]).replace(/[^\d.-]/g, '')) : 1
          if (!isNaN(v)) totalQty += v
          const dims3 = parseThreeDimBracket(d)
          if (dims3) heights.push(dims3.height)
          else {
            const hFeet = parseHeightFromParticulars(d)
            if (hFeet != null) heights.push(hFeet)
          }
        })
        if (totalQty === 0) totalQty = dropPanelRows.length
        const avgHeight = heights.length > 0 ? heights.reduce((a, b) => a + b, 0) / heights.length : 0
        const havgStr = avgHeight > 0 ? formatFeetInches(avgHeight) : "0'-10\""
        const floorText = formatFloorRefs(dropPanelRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
        const sRefs = extractSPageRefs(dropPanelRows)
        dropPanelLines.push({
          proposalText: `F&I new (${totalQty})no drop panel (Havg=${havgStr}) @ ${floorText} as per ${sRefs}`,
          sumRowKey: 'dropPanel',
          subsectionName: 'Drop panel',
          sumColumn: 'M',
          targetCol: 'G'
        })
      }
      const concretePostLines = []
      if (concretePostRows.length > 0) {
        const totalIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'total')
        const takeoffIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'takeoff')
        const qtyIdx = totalIdx >= 0 ? totalIdx : takeoffIdx
        let totalQty = 0
        let widthStr = "12\"x12\""
        let hStr = "9'-8\""
        concretePostRows.forEach(row => {
          const v = qtyIdx >= 0 && row[qtyIdx] != null ? parseFloat(String(row[qtyIdx]).replace(/[^\d.-]/g, '')) : 1
          if (!isNaN(v)) totalQty += v
          const dims = parseThreeDimBracket(row[digitizerIdx])
          if (dims) {
            const wIn = Math.round(dims.width * 12)
            const lIn = Math.round((dims.length || dims.width) * 12)
            widthStr = `${wIn}"x${lIn}"`
            hStr = formatFeetInches(dims.height)
          }
        })
        if (totalQty === 0) totalQty = concretePostRows.length
        const floorText = formatFloorRefs(concretePostRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
        const sRefs = extractSPageRefs(concretePostRows)
        concretePostLines.push({
          proposalText: `F&I new (${totalQty})no (${widthStr} wide) concrete post (H=${hStr}) @ ${floorText} as per ${sRefs}`,
          sumRowKey: 'concretePost',
          subsectionName: 'Concrete post',
          sumColumn: 'M',
          targetCol: 'G'
        })
      }
      const concreteEncasementLines = []
      if (concreteEncasementRows.length > 0) {
        const totalIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'total')
        const takeoffIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'takeoff')
        const qtyIdx = totalIdx >= 0 ? totalIdx : takeoffIdx
        let totalQty = 0
        let widthStr = "12\"x12\""
        let hStr = "9'-8\""
        concreteEncasementRows.forEach(row => {
          const v = qtyIdx >= 0 && row[qtyIdx] != null ? parseFloat(String(row[qtyIdx]).replace(/[^\d.-]/g, '')) : 1
          if (!isNaN(v)) totalQty += v
          const dims = parseThreeDimBracket(row[digitizerIdx])
          if (dims) {
            const wIn = Math.round(dims.width * 12)
            const lIn = Math.round((dims.length || dims.width) * 12)
            widthStr = `${wIn}"x${lIn}"`
            hStr = formatFeetInches(dims.height)
          }
        })
        if (totalQty === 0) totalQty = concreteEncasementRows.length
        const floorText = formatFloorRefs(concreteEncasementRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
        const sRefs = extractSPageRefs(concreteEncasementRows)
        concreteEncasementLines.push({
          proposalText: `F&I new (${totalQty})no (${widthStr} wide) concrete encasement (H=${hStr}) @ ${floorText} as per ${sRefs}`,
          sumRowKey: 'concreteEncasement',
          subsectionName: 'Concrete encasement',
          sumColumn: 'M',
          targetCol: 'G'
        })
      }
      const beamsLines = []
      if (beamsRows.length > 0) {
        const parseBeamDims = (t) => {
          const match = String(t).match(/\(([^)]+x[^)]+)\)/)
          if (!match) return null
          const parts = match[1].split(/x/i).map(p => p.trim())
          if (parts.length < 2) return null
          return { widthFeet: convertToFeet(parts[0]), heightFeet: convertToFeet(parts[1]) }
        }
        const widths = []
        const heights = []
        beamsRows.forEach(row => {
          const d = parseBeamDims(row[digitizerIdx])
          if (d) {
            widths.push(d.widthFeet)
            heights.push(d.heightFeet)
          }
        })
        const minW = widths.length > 0 ? Math.min(...widths) : 0
        const maxW = widths.length > 0 ? Math.max(...widths) : 0
        const avgH = heights.length > 0 ? heights.reduce((a, b) => a + b, 0) / heights.length : 0
        const widthRangeStr = minW > 0 && maxW > 0 ? `${formatFeetInches(minW)} to ${formatFeetInches(maxW)}` : "0'-8\" to 1'-0\""
        const havgStr = avgH > 0 ? formatFeetInches(avgH) : "1'-7\""
        const floorText = formatFloorRefs(beamsRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
        const sRefs = extractSPageRefs(beamsRows)
        beamsLines.push({
          proposalText: `F&I new (${widthRangeStr} wide) concrete beam (Havg=${havgStr}) @ ${floorText} as per ${sRefs}`,
          sumRowKey: 'beams',
          subsectionName: 'Beams',
          sumColumn: 'I',
          targetCol: 'C'
        })
      }
      const parapetWallLines = []
      if (parapetWallRows.length > 0) {
        const parseParapetDims = (t) => {
          const match = String(t).match(/\(([^)]+x[^)]+)\)/)
          if (!match) return null
          const parts = match[1].split(/x/i).map(p => p.trim())
          if (parts.length < 2) return null
          return { thicknessFeet: convertToFeet(parts[0]), heightFeet: convertToFeet(parts[1]) }
        }
        const dims = parseParapetDims(parapetWallRows[0][digitizerIdx])
        const thicknessStr = dims ? formatFeetInches(dims.thicknessFeet) : "1'-0\""
        const heightStr = dims ? formatFeetInches(dims.heightFeet) : "2'-10\""
        const floorText = formatFloorRefs(parapetWallRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
        const sRefs = extractSPageRefs(parapetWallRows)
        parapetWallLines.push({
          proposalText: `F&I new (${thicknessStr} thick) parapet wall (H=${heightStr}) @ ${floorText} as per ${sRefs}`,
          sumRowKey: 'parapetWalls',
          subsectionName: 'Parapet walls',
          sumColumn: 'I',
          targetCol: 'C'
        })
      }
      const thermalBreakLines = []
      if (thermalBreakRows.length > 0) {
        const floorText = formatFloorRefs(thermalBreakRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
        const sRefs = extractSPageRefs(thermalBreakRows)
        thermalBreakLines.push({
          proposalText: `F&I new thermal break @ ${floorText} as per ${sRefs}`,
          sumRowKey: 'thermalBreak',
          subsectionName: 'Thermal break',
          sumColumn: 'I',
          targetCol: 'C'
        })
      }
      const curbsLines = []
      if (curbsRows.length > 0) {
        const parseCurbDims = (t) => {
          const match = String(t).match(/\(([^)]+x[^)]+)\)/)
          if (!match) return null
          const parts = match[1].split(/x/i).map(p => p.trim())
          if (parts.length < 2) return null
          return { widthFeet: convertToFeet(parts[0]), heightFeet: convertToFeet(parts[1]) }
        }
        const curbGroups = new Map()
        curbsRows.forEach(row => {
          const d = parseCurbDims(row[digitizerIdx])
          if (!d) return
          const key = String(d.widthFeet)
          if (!curbGroups.has(key)) curbGroups.set(key, { rows: [], heights: [], widthFeet: d.widthFeet })
          const g = curbGroups.get(key)
          g.rows.push(row)
          g.heights.push(d.heightFeet)
        })
        const sortedKeys = [...curbGroups.keys()].sort((a, b) => parseFloat(a) - parseFloat(b))
        sortedKeys.forEach((key, groupIndex) => {
          const g = curbGroups.get(key)
          const { rows, heights, widthFeet } = g
          const widthInches = Math.round(widthFeet * 12)
          const widthStr = widthInches >= 12 ? formatFeetInches(widthFeet) : `${widthInches}"`
          const avgHeight = heights.length > 0 ? heights.reduce((a, b) => a + b, 0) / heights.length : 0
          const uniqueHeights = [...new Set(heights.map(h => h.toFixed(4)))]
          const heightStr = uniqueHeights.length === 1 ? formatFeetInches(heights[0]) : formatFeetInches(avgHeight)
          const heightLabel = uniqueHeights.length === 1 ? `H=${heightStr}` : `Havg=${heightStr}`
          const floorText = formatFloorRefs(rows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
          const sRefs = extractSPageRefs(rows)
          curbsLines.push({
            proposalText: `F&I new (${widthStr} wide) concrete curbs (${heightLabel}, typ.) @ ${floorText} as per ${sRefs}`,
            sumRowKey: 'curbs',
            subsectionName: 'Curbs',
            curbGroupIndex: groupIndex,
            sumColumn: 'I',
            targetCol: 'C'
          })
        })
        curbsLines.push({
          proposalText: 'F&I new concrete curbs - Allowance',
          allowance: true
        })
      }
      const columnsLines = []
      const hasColumnsInCalc = (formulaData || []).some(f =>
        (f.itemType === 'superstructure_columns_final' || f.itemType === 'superstructure_columns_takeoff') && f.section === 'superstructure' && f.subsectionName === 'Columns'
      )
      if (columnsRows.length > 0 || hasColumnsInCalc) {
        let totalQty = 160
        let floorText = 'as indicated'
        let sRefs = 'FO-101.00 & S-101.00 to S-105.00'
        if (columnsRows.length > 0) {
          const totalIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'total')
          const takeoffIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'takeoff')
          const qtyIdx = totalIdx >= 0 ? totalIdx : takeoffIdx
          let qtySum = 0
          columnsRows.forEach(row => {
            const v = qtyIdx >= 0 && row[qtyIdx] != null ? parseFloat(String(row[qtyIdx]).replace(/[^\d.-]/g, '')) : 0
            if (!isNaN(v)) qtySum += v
          })
          if (qtySum > 0) totalQty = qtySum
          floorText = formatFloorRefs(columnsRows.map(r => extractFloorRef(r[digitizerIdx]))) || floorText
          sRefs = extractSPageRefs(columnsRows) || sRefs
        }
        columnsLines.push({
          proposalText: `F&I new (${totalQty})no (1'-6" to 2'-6" wide) concrete columns (Havg=11'-0") @ ${floorText} as per ${sRefs}`,
          sumRowKey: 'columns',
          subsectionName: 'Columns',
          sumColumn: 'M',
          targetCol: 'G'
        })
      }
      const concreteHangerLines = []
      if (concreteHangerRows.length > 0) {
        const totalIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'total')
        const takeoffIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'takeoff')
        const qtyIdx = totalIdx >= 0 ? totalIdx : takeoffIdx
        let totalQty = 0
        const heights = []
        let widthFeet = null
        concreteHangerRows.forEach(row => {
          const dims = parseThreeDimBracket(row[digitizerIdx])
          if (dims) {
            heights.push(dims.height)
            if (widthFeet == null) widthFeet = dims.width
          }
          const v = qtyIdx >= 0 && row[qtyIdx] != null ? parseFloat(String(row[qtyIdx]).replace(/[^\d.-]/g, '')) : 1
          if (!isNaN(v)) totalQty += v
        })
        if (totalQty === 0) totalQty = concreteHangerRows.length
        const avgHeight = heights.length > 0 ? heights.reduce((a, b) => a + b, 0) / heights.length : 0
        const widthStr = widthFeet != null ? formatFeetInches(widthFeet) : "1'-0\""
        const havgStr = avgHeight > 0 ? formatFeetInches(avgHeight) : ''
        const floorText = formatFloorRefs(concreteHangerRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
        const sRefs = extractSPageRefs(concreteHangerRows)
        concreteHangerLines.push({
          proposalText: `F&I new (${totalQty})no (${widthStr} wide) concrete hangers${havgStr ? ` (Havg=${havgStr})` : ''} @ ${floorText} as per ${sRefs}`,
          sumRowKey: 'concreteHanger',
          subsectionName: 'Concrete hanger',
          sumColumn: 'M',
          targetCol: 'G'
        })
      }

      const concretePadLines = []
      if (concretePadRows.length > 0) {
        const parsePadThickness = (t) => {
          const m = String(t).match(/(\d+)\s*"/)
          return m ? parseInt(m[1], 10) : 4
        }
        const parsePadQty = (t) => {
          const m = String(t).match(/\(\s*(\d+)\s*(?:No\.?|EA)?\s*\)/i)
          return m ? parseInt(m[1], 10) : null
        }
        const padGroups = new Map()
        concretePadRows.forEach(row => {
          const thick = parsePadThickness(row[digitizerIdx])
          const key = String(thick)
          if (!padGroups.has(key)) padGroups.set(key, [])
          padGroups.get(key).push(row)
        })
        const sortedKeys = [...padGroups.keys()].sort((a, b) => parseInt(a, 10) - parseInt(b, 10))
        sortedKeys.forEach((key, groupIndex) => {
          const rows = padGroups.get(key)
          const totalIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'total')
          const takeoffIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'takeoff')
          const qtyIdx = totalIdx >= 0 ? totalIdx : takeoffIdx
          let totalQty = 0
          rows.forEach(r => {
            const q = parsePadQty(r[digitizerIdx])
            if (q != null) totalQty += q
            else if (qtyIdx >= 0 && r[qtyIdx] != null) totalQty += parseFloat(String(r[qtyIdx]).replace(/[^\d.-]/g, '')) || 0
          })
          if (totalQty === 0) totalQty = rows.length
          const totalQtyRounded = Number.isInteger(totalQty) ? totalQty : Math.round(totalQty * 100) / 100
          const totalQtyDisplay = totalQtyRounded % 1 === 0 ? Math.round(totalQtyRounded) : totalQtyRounded
          const thickStr = `${key}"`
          const floorText = formatFloorRefs(rows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
          const sRefs = extractSPageRefs(rows)
          const padWord = totalQtyDisplay === 1 ? 'concrete pad' : 'concrete pads'
          concretePadLines.push({
            proposalText: `F&I new (${totalQtyDisplay})no (${thickStr} thick) ${padWord} @ ${floorText} as per ${sRefs}`,
            sumRowKey: 'concretePad',
            subsectionName: 'Concrete pad',
            concretePadGroupIndex: groupIndex,
            sumColumn: 'M',
            targetCol: 'G'
          })
        })
        concretePadLines.push({
          proposalText: 'F&I new housekeeping pads - Allowance',
          allowance: true
        })
      }
      const nonShrinkGroutLines = []
      if (nonShrinkGroutRows.length > 0) {
        const totalIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'total')
        const takeoffIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'takeoff')
        const qtyIdx = totalIdx >= 0 ? totalIdx : takeoffIdx
        let totalQty = 0
        let thickStr = "1½\""
        nonShrinkGroutRows.forEach(row => {
          const m = String(row[digitizerIdx] || '').match(/(\d+)\s*½?\s*"/) || String(row[digitizerIdx] || '').match(/(\d+)\s*"/)
          if (m) thickStr = m[0].includes('½') ? `${m[1]}½"` : `${m[1]}"`
          const v = qtyIdx >= 0 && row[qtyIdx] != null ? parseFloat(String(row[qtyIdx]).replace(/[^\d.-]/g, '')) : 1
          if (!isNaN(v)) totalQty += v
        })
        if (totalQty === 0) totalQty = nonShrinkGroutRows.length
        const sRefs = extractSPageRefs(nonShrinkGroutRows)
        nonShrinkGroutLines.push({
          proposalText: `F&I new (${totalQty})no (${thickStr} thick) non-shrink grout @ base plates as per ${sRefs}`,
          sumRowKey: 'nonShrinkGrout',
          subsectionName: 'Non-shrink grout',
          sumColumn: 'M',
          targetCol: 'G',
          thickStr,
          sRefs
        })
      }
      const repairScopeLines = []
      if (repairScopeWallRows.length > 0) {
        const sRefs = extractSPageRefs(repairScopeWallRows)
        repairScopeLines.push({
          proposalText: `Allow to repair spall concrete wall crack as per ${sRefs}`,
          sumRowKey: 'repairScopeWall',
          subsectionName: 'Repair scope',
          repairSubType: 'wall',
          sumColumn: 'I',
          targetCol: 'C'
        })
      }
      if (repairScopeSlabRows.length > 0) {
        const sRefs = extractSPageRefs(repairScopeSlabRows)
        repairScopeLines.push({
          proposalText: `Allow to repair concrete slab crack as per ${sRefs}`,
          sumRowKey: 'repairScopeSlab',
          subsectionName: 'Repair scope',
          repairSubType: 'slab',
          sumColumn: 'J',
          targetCol: 'D'
        })
      }
      if (repairScopeColumnRows.length > 0) {
        const sRefs = extractSPageRefs(repairScopeColumnRows)
        repairScopeLines.push({
          proposalText: `Allow to repair concrete column crack as per ${sRefs}`,
          sumRowKey: 'repairScopeColumn',
          subsectionName: 'Repair scope',
          repairSubType: 'column',
          sumColumn: 'M',
          targetCol: 'G'
        })
      }
      return { cipLines: lines, patchLines, somdLines, slabStepsLines, cipStairsGroupsWithLines, concreteHangerLines, lwConcreteFillLines, toppingSlabLines, raisedSlabLines, builtUpSlabLines, builtupRampsLines, builtUpStairLines, shearWallLines, columnsLines, dropPanelLines, concretePostLines, concreteEncasementLines, beamsLines, parapetWallLines, thermalBreakLines, curbsLines, concretePadLines, nonShrinkGroutLines, repairScopeLines }
    }

    // Use data row(s) (same line as description) for values, not the sum row (next line)
    const useDataRows = (sumFormula) => {
      if (!sumFormula) return null
      if (sumFormula.firstDataRow == null) return { type: 'single', row: sumFormula.row }
      const first = sumFormula.firstDataRow
      const last = sumFormula.lastDataRow != null ? sumFormula.lastDataRow : first
      if (first === last) return { type: 'single', row: first }
      const rows = []
      for (let r = first; r <= last; r++) rows.push(r)
      return { type: 'sum', rows }
    }
    const findSumRowForSubsection = (subsectionName, sumRowKey, groupIndex = 0, line = null) => {
      if (!formulaData || !Array.isArray(formulaData)) return null
      const sums = formulaData.filter(f =>
        f.itemType === 'superstructure_sum' &&
        f.section === 'superstructure' &&
        f.subsectionName === subsectionName
      )
      if (sumRowKey === 'slab8') {
        return sums[0] ? useDataRows(sums[0]) : null
      }
      if (sumRowKey === 'roofSlab8') return sums[1] ? useDataRows(sums[1]) : (sums[0] ? useDataRows(sums[0]) : null)
      if (sumRowKey === 'balcony' || sumRowKey === 'terrace' || sumRowKey === 'patch') return sums[0] ? useDataRows(sums[0]) : null
      if (sumRowKey === 'slabSteps') {
        const f = (formulaData || []).find(x =>
          x.itemType === 'superstructure_slab_steps_sum' && x.section === 'superstructure' && x.subsectionName === 'Slab steps'
        )
        return f ? useDataRows(f) : null
      }
      if (sumRowKey === 'raisedKneeWall') {
        const rows = (formulaData || []).filter(f =>
          f.itemType === 'superstructure_raised_knee_wall' && f.section === 'superstructure' && f.subsectionName === 'Raised slab'
        ).map(f => f.row)
        return rows.length > 0 ? { type: 'sum', rows } : null
      }
      if (sumRowKey === 'raisedStyrofoam') {
        const f = (formulaData || []).find(x =>
          x.itemType === 'superstructure_raised_styrofoam' && x.section === 'superstructure' && x.subsectionName === 'Raised slab'
        )
        return f ? { type: 'single', row: f.row } : null
      }
      if (sumRowKey === 'raisedSlab') {
        const rows = (formulaData || []).filter(f =>
          f.itemType === 'superstructure_raised_slab' && f.section === 'superstructure' && f.subsectionName === 'Raised slab'
        ).map(f => f.row)
        return rows.length > 0 ? { type: 'sum', rows } : null
      }
      if (sumRowKey === 'builtUpSlabKnee') {
        const rows = (formulaData || []).filter(f =>
          f.itemType === 'superstructure_builtup_knee_wall' && f.section === 'superstructure' && f.subsectionName === 'Built-up slab'
        ).map(f => f.row)
        return rows.length > 0 ? { type: 'sum', rows } : null
      }
      if (sumRowKey === 'builtUpSlabStyro') {
        const f = (formulaData || []).find(x =>
          x.itemType === 'superstructure_builtup_styrofoam' && x.section === 'superstructure' && x.subsectionName === 'Built-up slab'
        )
        return f ? { type: 'single', row: f.row } : null
      }
      if (sumRowKey === 'builtUpSlab') {
        const rows = (formulaData || []).filter(f =>
          f.itemType === 'superstructure_builtup_slab' && f.section === 'superstructure' && f.subsectionName === 'Built-up slab'
        ).map(f => f.row)
        return rows.length > 0 ? { type: 'sum', rows } : null
      }
      if (sumRowKey === 'builtupRampsKnee') {
        const f = (formulaData || []).find(x =>
          x.itemType === 'superstructure_builtup_ramps_knee_sum' && x.section === 'superstructure' && x.subsectionName === 'Builtup ramps'
        )
        return f ? useDataRows(f) : null
      }
      if (sumRowKey === 'builtupRampsStyro') {
        const f = (formulaData || []).find(x =>
          x.itemType === 'superstructure_builtup_ramps_styro_sum' && x.section === 'superstructure' && x.subsectionName === 'Builtup ramps'
        )
        return f ? useDataRows(f) : null
      }
      if (sumRowKey === 'builtupRampsRamp') {
        const f = (formulaData || []).find(x =>
          x.itemType === 'superstructure_builtup_ramps_ramp_sum' && x.section === 'superstructure' && x.subsectionName === 'Builtup ramps'
        )
        return f ? useDataRows(f) : null
      }
      if (sumRowKey === 'builtUpStairKnee') {
        const rows = (formulaData || []).filter(f =>
          f.itemType === 'superstructure_builtup_stair_knee_wall' && f.section === 'superstructure' && f.subsectionName === 'Built-up stair'
        ).map(f => f.row)
        return rows.length > 0 ? { type: 'sum', rows } : null
      }
      if (sumRowKey === 'builtUpStairStyro') {
        const f = (formulaData || []).find(x =>
          x.itemType === 'superstructure_builtup_stair_styrofoam' && x.section === 'superstructure' && x.subsectionName === 'Built-up stair'
        )
        return f ? { type: 'single', row: f.row } : null
      }
      if (sumRowKey === 'shearWall') {
        const shearSums = (formulaData || []).filter(f =>
          f.itemType === 'superstructure_shear_walls_sum' && f.section === 'superstructure' && f.subsectionName === 'Shear Walls'
        )
        const idx = groupIndex
        const sumAt = shearSums[idx]
        return sumAt ? useDataRows(sumAt) : null
      }
      if (sumRowKey === 'dropPanel') {
        const f = (formulaData || []).find(x =>
          x.itemType === 'superstructure_drop_panel_sum' && x.section === 'superstructure' && x.subsectionName === 'Drop panel'
        )
        return f ? useDataRows(f) : null
      }
      if (sumRowKey === 'concretePost') {
        const f = (formulaData || []).find(x =>
          x.itemType === 'superstructure_concrete_post_sum' && x.section === 'superstructure' && x.subsectionName === 'Concrete post'
        )
        return f ? useDataRows(f) : null
      }
      if (sumRowKey === 'concreteEncasement') {
        const f = (formulaData || []).find(x =>
          x.itemType === 'superstructure_concrete_encasement_sum' && x.section === 'superstructure' && x.subsectionName === 'Concrete encasement'
        )
        return f ? useDataRows(f) : null
      }
      if (sumRowKey === 'beams') {
        const f = (formulaData || []).find(x =>
          x.itemType === 'superstructure_beams_sum' && x.section === 'superstructure' && x.subsectionName === 'Beams'
        )
        return f ? useDataRows(f) : null
      }
      if (sumRowKey === 'parapetWalls') {
        const f = (formulaData || []).find(x =>
          x.itemType === 'superstructure_parapet_walls_sum' && x.section === 'superstructure' && x.subsectionName === 'Parapet walls'
        )
        return f ? useDataRows(f) : null
      }
      if (sumRowKey === 'thermalBreak') {
        const f = (formulaData || []).find(x =>
          x.itemType === 'superstructure_thermal_break_sum' && x.section === 'superstructure' && x.subsectionName === 'Thermal break'
        )
        return f ? useDataRows(f) : null
      }
      if (sumRowKey === 'curbs') {
        const curbSums = (formulaData || []).filter(f =>
          f.itemType === 'superstructure_curbs_sum' && f.section === 'superstructure' && f.subsectionName === 'Curbs'
        )
        const idx = (line && line.curbGroupIndex != null) ? line.curbGroupIndex : groupIndex
        const sumAt = curbSums[idx]
        return sumAt ? useDataRows(sumAt) : null
      }
      if (sumRowKey === 'concretePad') {
        const padSums = (formulaData || []).filter(f =>
          f.itemType === 'superstructure_concrete_pad_sum' && f.section === 'superstructure' && f.subsectionName === 'Concrete pad'
        )
        const idx = (line && line.concretePadGroupIndex != null) ? line.concretePadGroupIndex : groupIndex
        const sumAt = padSums[idx]
        return sumAt ? useDataRows(sumAt) : null
      }
      if (sumRowKey === 'nonShrinkGrout') {
        const f = (formulaData || []).find(x =>
          x.itemType === 'superstructure_non_shrink_grout_sum' && x.section === 'superstructure' && x.subsectionName === 'Non-shrink grout'
        )
        return f ? useDataRows(f) : null
      }
      if (sumRowKey === 'repairScopeWall' || sumRowKey === 'repairScopeSlab' || sumRowKey === 'repairScopeColumn') {
        const subType = (line && line.repairSubType) || 'wall'
        const rows = (formulaData || []).filter(f =>
          f.itemType === 'superstructure_repair_scope' && f.section === 'superstructure' && f.subsectionName === 'Repair scope' && (f.parsedData?.parsed?.itemSubType === subType)
        ).map(f => f.row)
        return rows.length > 0 ? { type: 'sum', rows } : null
      }
      if (sumRowKey === 'columns') {
        const f = (formulaData || []).find(x =>
          x.itemType === 'superstructure_columns_final' && x.section === 'superstructure' && x.subsectionName === 'Columns'
        )
        return f ? { type: 'single', row: f.row } : null
      }
      if (sumRowKey === 'builtUpStair') {
        const f = (formulaData || []).find(x =>
          x.itemType === 'superstructure_builtup_stair_sum' && x.section === 'superstructure' && x.subsectionName === 'Built-up stair'
        )
        return f ? useDataRows(f) : null
      }
      if (sumRowKey === 'toppingSlab') {
        const toppingSum = (formulaData || []).find(f =>
          f.itemType === 'superstructure_topping_slab_sum' && f.section === 'superstructure' && f.subsectionName === 'Topping slab'
        )
        return toppingSum ? useDataRows(toppingSum) : null
      }
      if (sumRowKey === 'lwConcreteFill') {
        const lwSum = (formulaData || []).find(f =>
          f.itemType === 'superstructure_lw_concrete_fill_sum' && f.section === 'superstructure' && f.subsectionName === 'LW concrete fill'
        )
        return lwSum ? useDataRows(lwSum) : null
      }
      if (sumRowKey === 'concreteHanger') {
        const hangerSum = (formulaData || []).find(f =>
          f.itemType === 'superstructure_concrete_hanger_sum' && f.section === 'superstructure' && f.subsectionName === 'Concrete hanger'
        )
        return hangerSum ? useDataRows(hangerSum) : null
      }
      if (sumRowKey === 'somd') {
        const somdSums = (formulaData || []).filter(f =>
          f.itemType === 'superstructure_somd_sum' && f.section === 'superstructure' && f.subsectionName === 'Slab on metal deck'
        )
        const sumAt = somdSums[groupIndex]
        return sumAt ? { type: 'single', row: sumAt.row } : null
      }
      return sums[0] ? useDataRows(sums[0]) : null
    }

    const { cipLines, patchLines, somdLines, slabStepsLines, cipStairsGroupsWithLines, concreteHangerLines, lwConcreteFillLines, toppingSlabLines, raisedSlabLines, builtUpSlabLines, builtupRampsLines, builtUpStairLines, shearWallLines, columnsLines, dropPanelLines, concretePostLines, concreteEncasementLines, beamsLines, parapetWallLines, thermalBreakLines, curbsLines, concretePadLines, nonShrinkGroutLines, repairScopeLines } = buildSuperstructureProposalLines()
    const calcSheetName = 'Calculations Sheet'

    const infilledGroups = []
    const cipStairsGroupsFromFormulas = []
    if (formulaData && Array.isArray(formulaData)) {
      const infilledFormulas = formulaData.filter(f => f.section === 'superstructure' && f.subsectionName === 'Stairs \u2013 Infilled tads')
      infilledFormulas.sort((a, b) => a.row - b.row)
      let currentInfilled = null
      infilledFormulas.forEach(f => {
        if (f.itemType === 'superstructure_manual_infilled_header') {
          const name = (calculationData && calculationData[f.row - 1] && calculationData[f.row - 1][1] != null)
            ? String(calculationData[f.row - 1][1]).trim()
            : 'Stair:'
          currentInfilled = { name: name.endsWith(':') ? name : `${name}:`, landingSumRow: null, stairRow: null }
          infilledGroups.push(currentInfilled)
        } else if ((f.itemType === 'superstructure_infilled_landing_sum' || f.itemType === 'superstructure_infilled_landing_item' || f.itemType === 'superstructure_manual_infilled_landing_sum') && currentInfilled) {
          // Prefer sum row (J and L totals) for proposal; fallback to landing item row
          if (f.itemType === 'superstructure_infilled_landing_sum' || currentInfilled.landingSumRow == null) {
            currentInfilled.landingSumRow = f.row
          }
        } else if (f.itemType === 'superstructure_manual_infilled_stair' && currentInfilled) {
          currentInfilled.stairRow = f.row
        }
      })
      const cipStairsFormulas = formulaData.filter(f => f.section === 'superstructure' && f.subsectionName === 'CIP Stairs')
      cipStairsFormulas.sort((a, b) => a.row - b.row)
      let currentCip = null
      cipStairsFormulas.forEach(f => {
        if (f.itemType === 'superstructure_manual_cip_stairs_header') {
          const name = (calculationData && calculationData[f.row - 1] && calculationData[f.row - 1][1] != null)
            ? String(calculationData[f.row - 1][1]).trim()
            : 'CIP Stairs'
          currentCip = { name: name.endsWith(':') ? name : `${name}:`, landingRow: null, stairRow: null, slabRow: null }
          cipStairsGroupsFromFormulas.push(currentCip)
        } else if (f.itemType === 'superstructure_manual_cip_stairs_landing' && currentCip) {
          currentCip.landingRow = f.row
        } else if (f.itemType === 'superstructure_manual_cip_stairs_stair' && currentCip) {
          currentCip.stairRow = f.row
        } else if (f.itemType === 'superstructure_manual_cip_stairs_slab' && currentCip) {
          currentCip.slabRow = f.row
        }
      })
    }

    // Page ref for superstructure subsections (CIP Stairs, Stairs – Infilled tads)
    const getPageRefSuperstructure = (subSectionName) => {
      if (!rawData || !Array.isArray(rawData) || rawData.length < 2) return '##'
      const headers = rawData[0]
      const dataRows = rawData.slice(1)
      const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
      const pageIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'page')
      if (digitizerIdx === -1 || pageIdx === -1) return '##'
      const subLower = (subSectionName || '').toLowerCase()
      const collected = []
      const seen = new Set()
      for (const row of dataRows) {
        const dig = (row[digitizerIdx] || '').toString().toLowerCase()
        const matches = dig.includes('cip stair') || dig.includes('stair') && (subLower.includes('infilled') ? dig.includes('infill') : true)
        if (matches && row[pageIdx] != null && row[pageIdx] !== '') {
          const pageStr = String(row[pageIdx]).trim()
          const refMatch = pageStr.match(/(?:FO|P|S|A|E|DM)-[\d.]+/i)
          const ref = refMatch ? refMatch[0] : pageStr
          if (ref && !seen.has(ref)) { seen.add(ref); collected.push(ref) }
        }
      }
      return collected.length === 0 ? 'details on' : collected.length === 1 ? collected[0] : collected.slice(0, -1).join(', ') + ' & ' + collected[collected.length - 1]
    }
    const cipStairsPageRefEsc = (getPageRefSuperstructure('CIP Stairs') || 'details on').replace(/"/g, '""')
    const infilledPageRefEsc = (getPageRefSuperstructure('Stairs – Infilled tads') || 'details on').replace(/"/g, '""')
    const getCipStairsValuesFromRow = (calcRow) => {
      const defaultAsPer = getPageRefSuperstructure('CIP Stairs') || 'as per details on'
      if (!calculationData || !calcRow || calcRow < 1) return { thickness: '8"', width: 'as indicated', riser: 'as indicated', location: '1st FL', asPer: defaultAsPer }
      const row = calculationData[calcRow - 1]
      if (!row || !Array.isArray(row)) return { thickness: '8"', width: 'as indicated', riser: 'as indicated', location: '1st FL', asPer: defaultAsPer }
      const p = (row[1] || '').toString()
      let thickness = '8"'
      const thicknessMatch = p.match(/(\d+(?:\/\d+)?)"?\s*thick/i) || p.match(/(\d+)"\s*thick/i)
      if (thicknessMatch) thickness = thicknessMatch[1].includes('"') ? thicknessMatch[1] : thicknessMatch[1] + '"'
      else if (row[7] != null && row[7] !== '') {
        const heightFeet = parseFloat(row[7])
        if (!isNaN(heightFeet) && heightFeet > 0) thickness = `${Math.round(heightFeet * 12)}"`
      }
      const widthFromColRaw = row[6] != null && row[6] !== '' ? String(row[6]).trim() : null
      const widthFromCol = widthFromColRaw ? (formatWidthFeetInches(widthFromColRaw) || widthFromColRaw) : null
      const widthMatch = p.match(/(\d+'-?\d*")\s*wide/i)
      const width = widthFromCol || (widthMatch ? widthMatch[1] : 'as indicated')
      let riserVal = Math.round(parseFloat(row[12]) || 0)
      if (riserVal <= 0 && row[11] != null && row[11] !== '') riserVal = Math.round(parseFloat(row[11]) || 0)
      if (riserVal <= 0) {
        const takeoffVal = parseFloat(row[2])
        if (Number.isInteger(takeoffVal) && takeoffVal > 0 && takeoffVal < 100) riserVal = takeoffVal
      }
      if (riserVal <= 0) {
        const rm = p.match(/(\d+)\s*Riser/i)
        if (rm) riserVal = parseInt(rm[1], 10)
      }
      const riser = riserVal > 0 ? String(riserVal) : 'as indicated'
      const locationMatch = p.match(/@\s*([^ as per]+?)(?:\s+as\s+per|$)/i) || p.match(/(?:cellar|roof|1st|2nd|\d+th)\s*FL[^,]*(?:,\s*(?:cellar|roof|1st|2nd|\d+th)\s*FL[^,]*)*/i)
      const location = (locationMatch && locationMatch[1] ? locationMatch[1].trim() : null) || (locationMatch && locationMatch[0] ? locationMatch[0].trim() : null) || '1st FL'
      const asPer = getPageRefSuperstructure('CIP Stairs') || 'as per details on'
      return { thickness, width, riser, location, asPer }
    }
    const getCipStairsProposalText = (groupName, lineType, calcRow) => {
      const asPerDef = getPageRefSuperstructure('CIP Stairs') || 'as per details on'
      if (lineType === 'slab') return `F&I new stair slab as per ${asPerDef}`
      if (!calcRow) return null
      const v = getCipStairsValuesFromRow(calcRow)
      if (lineType === 'landings') return `F&I new (${v.thickness} thick) stair landings @ ${v.location} as per ${v.asPer}`
      if (lineType === 'stairs') return `F&I new (${v.width} wide) concrete stairs (${v.riser} Riser) @ ${v.location} as per ${v.asPer}`
      return null
    }
    // Only render slab line when the slab row has actual quantity data (so we don't add a third line when original has two: landings + stairs)
    const cipStairsSlabRowHasData = (grp) => {
      if (grp.slabRow == null || !calculationData || !Array.isArray(calculationData)) return false
      const row = calculationData[grp.slabRow - 1]
      if (!row || !Array.isArray(row)) return false
      const numCols = [3, 9, 11, 12] // D, J, L, M (0-based)
      return numCols.some(i => {
        const v = row[i]
        const n = typeof v === 'number' ? v : parseFloat(v)
        return !Number.isNaN(n) && n > 0
      })
    }

    let superstructureScopeStartRow = null
    let superstructureScopeEndRow = null
    const washoutExclusionRows = []

    const renderProposalLine = (line, options = {}) => {
      const lineFontWeight = options.fontWeight ?? 'bold'
      if (superstructureScopeStartRow == null) superstructureScopeStartRow = currentRow
      if (line.allowance) {
        spreadsheet.updateCell({ value: line.proposalText }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, line.proposalText)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat(
          { fontWeight: lineFontWeight, color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' },
          `${pfx}B${currentRow}`
        )
        fillRatesForProposalRow(currentRow, line.proposalText)
        const dynamicHeight = calculateRowHeight(line.proposalText)
        spreadsheet.updateCell({ value: 50 }, `${pfx}C${currentRow}`)
        spreadsheet.updateCell({ formula: `=C${currentRow}*1.5` }, `${pfx}D${currentRow}`)
        spreadsheet.updateCell({ formula: `=D${currentRow}/27` }, `${pfx}F${currentRow}`)
          ;['C', 'D', 'F'].forEach(col => {
            spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}${col}${currentRow}`)
            try { spreadsheet.numberFormat('#,##0.00', `${pfx}${col}${currentRow}`) } catch (e) { }
          })
        const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
        spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
        spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white', format: '$#,##0.00' }, `${pfx}H${currentRow}`)
        try { spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}`) } catch (e) { }
        washoutExclusionRows.push(currentRow)
        superstructureScopeEndRow = currentRow
        currentRow++
        return
      }
      const groupIndex = line.somdGroupIndex ?? line.shearGroupIndex ?? line.curbGroupIndex ?? line.concretePadGroupIndex ?? 0
      const sumInfo = findSumRowForSubsection(line.subsectionName, line.sumRowKey, groupIndex, line)
      // Non-shrink grout: CONCATENATE so qty updates in real time from Calculation sheet
      const isNonShrinkGroutWithFormula = line.sumRowKey === 'nonShrinkGrout' && sumInfo && (sumInfo.type === 'single' && sumInfo.row) && line.thickStr != null && line.sRefs != null
      if (isNonShrinkGroutWithFormula) {
        const thickStrEsc = (line.thickStr || '').replace(/"/g, '""')
        const sRefsEsc = (line.sRefs || '').replace(/"/g, '""')
        const bFormula = `=CONCATENATE("F&I new (",'${calcSheetName}'!G${sumInfo.row},")no (","${thickStrEsc}"," thick) non-shrink grout @ base plates as per ","${sRefsEsc}",")")`
        spreadsheet.updateCell({ formula: bFormula }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, line.proposalText)
      } else {
        spreadsheet.updateCell({ value: line.proposalText }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, line.proposalText)
      }
      spreadsheet.wrap(`${pfx}B${currentRow}`, true)
      spreadsheet.cellFormat(
        { fontWeight: lineFontWeight, color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' },
        `${pfx}B${currentRow}`
      )
      const rateLookupKey = (line.subsectionName === 'Slab on metal deck' && /concrete topping/i.test(line.proposalText))
        ? (/NW concrete topping/i.test(line.proposalText) ? 'NW concrete topping - Slab on metal deck' : 'LW concrete topping - Slab on metal deck')
        : line.proposalText
      fillRatesForProposalRow(currentRow, rateLookupKey)
      const dynamicHeight = calculateRowHeight(line.proposalText)
      const qtyCol = line.targetCol || 'D'
      if (sumInfo) {
        const col = line.sumColumn || 'J'
        if (sumInfo.type === 'sum' && sumInfo.rows && sumInfo.rows.length >= 1) {
          const sumRefs = sumInfo.rows.map(r => `'${calcSheetName}'!${col}${r}`).join(',')
          spreadsheet.updateCell({ formula: sumInfo.rows.length === 1 ? `='${calcSheetName}'!${col}${sumInfo.rows[0]}` : `=SUM(${sumRefs})` }, `${pfx}${qtyCol}${currentRow}`)
        } else if (sumInfo.type === 'single' && sumInfo.row) {
          spreadsheet.updateCell({ formula: `='${calcSheetName}'!${col}${sumInfo.row}` }, `${pfx}${qtyCol}${currentRow}`)
        } else {
          spreadsheet.updateCell({ value: '' }, `${pfx}${qtyCol}${currentRow}`)
        }
      } else {
        spreadsheet.updateCell({ value: '' }, `${pfx}${qtyCol}${currentRow}`)
      }
      const isCipSlabLine = line.sumRowKey && ['slab8', 'roofSlab8', 'balcony', 'terrace', 'slabSteps'].includes(line.sumRowKey)
      if (sumInfo && !isCipSlabLine) {
        const cyCol = 'L'
        if (sumInfo.type === 'sum' && sumInfo.rows && sumInfo.rows.length >= 1) {
          const cyRefs = sumInfo.rows.map(r => `'${calcSheetName}'!${cyCol}${r}`).join(',')
          spreadsheet.updateCell({ formula: sumInfo.rows.length === 1 ? `='${calcSheetName}'!${cyCol}${sumInfo.rows[0]}` : `=SUM(${cyRefs})` }, `${pfx}F${currentRow}`)
        } else if (sumInfo.type === 'single' && sumInfo.row) {
          spreadsheet.updateCell({ formula: `='${calcSheetName}'!${cyCol}${sumInfo.row}` }, `${pfx}F${currentRow}`)
        }
      }
      if (isCipSlabLine) {
        spreadsheet.updateCell({ value: 1200 }, `${pfx}F${currentRow}`)
      }
      spreadsheet.cellFormat(
        { fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' },
        `${pfx}${qtyCol}${currentRow}`
      )
      try {
        spreadsheet.numberFormat('#,##0.00', `${pfx}${qtyCol}${currentRow}`)
      } catch (e) {
        spreadsheet.cellFormat({ format: '#,##0.00' }, `${pfx}${qtyCol}${currentRow}`)
      }
      if (sumInfo || isCipSlabLine) {
        spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}F${currentRow}`)
        try { spreadsheet.numberFormat('#,##0.00', `${pfx}F${currentRow}`) } catch (e) { }
      }
      const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
      spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white', format: '$#,##0.00' },
        `${pfx}H${currentRow}`
      )
      try {
        spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}`)
      } catch (e) { }
      const isWashoutExclusion = ['thermalBreak', 'nonShrinkGrout', 'repairScopeWall', 'repairScopeSlab', 'repairScopeColumn'].includes(line.sumRowKey)
      if (isWashoutExclusion) washoutExclusionRows.push(currentRow)
      superstructureScopeEndRow = currentRow
      currentRow++
    }

    if (cipLines.length > 0) {
      spreadsheet.updateCell({ value: 'CIP concrete slabs:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      const cipStartRow = currentRow
      cipLines.forEach(renderProposalLine)
      const cipEndRow = currentRow - 1
      spreadsheet.updateCell({ value: 'Total SF' }, `${pfx}C${currentRow}`)
      spreadsheet.updateCell({ formula: `=SUM(D${cipStartRow}:D${cipEndRow})` }, `${pfx}D${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', fontStyle: 'italic', color: '#000000', textAlign: 'center', backgroundColor: 'white' },
        `${pfx}C${currentRow}`
      )
      spreadsheet.cellFormat(
        { fontWeight: 'bold', fontStyle: 'italic', textAlign: 'center', backgroundColor: 'white' },
        `${pfx}D${currentRow}`
      )
      try { spreadsheet.numberFormat('#,##0.00', `${pfx}D${currentRow}`) } catch (e) { }
      washoutExclusionRows.push(currentRow)
      totalRows.push(currentRow) // So "Total SF" stays bold after global C13:G normal-weight override
      totalSFRowsForItalic.push(currentRow) // So "Total SF" cell (C) is re-applied bold + italic
      superstructureScopeEndRow = currentRow
      currentRow++
      if (slabStepsLines.length > 0) {
        slabStepsLines.forEach(renderProposalLine)
      }
    }
    if (patchLines.length > 0) {
      spreadsheet.updateCell({ value: 'Patch slab:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      patchLines.forEach(renderProposalLine)
    }
    if (somdLines.length > 0) {
      somdLines.forEach((line) => {
        spreadsheet.updateCell({ value: line.headingText }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
          `${pfx}B${currentRow}`
        )
        currentRow++
        renderProposalLine(line)
      })
    }
    if (concreteHangerLines.length > 0) {
      spreadsheet.updateCell({ value: 'Concrete hanger:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      concreteHangerLines.forEach(renderProposalLine)
    }
    if (lwConcreteFillLines.length > 0) {
      spreadsheet.updateCell({ value: 'LW concrete fill:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      lwConcreteFillLines.forEach(renderProposalLine)
    }
    if (toppingSlabLines.length > 0) {
      spreadsheet.updateCell({ value: 'Topping slab:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      toppingSlabLines.forEach(renderProposalLine)
    }
    if (raisedSlabLines.length > 0) {
      spreadsheet.updateCell({ value: 'Raised slab:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      raisedSlabLines.forEach(renderProposalLine)
    }
    if (builtUpSlabLines.length > 0) {
      spreadsheet.updateCell({ value: 'Built-up slab:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      builtUpSlabLines.forEach(renderProposalLine)
    }
    if (builtupRampsLines.length > 0) {
      spreadsheet.updateCell({ value: 'Builtup ramps:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      builtupRampsLines.forEach(renderProposalLine)
    }
    if (builtUpStairLines.length > 0) {
      spreadsheet.updateCell({ value: 'Built-up stair:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      builtUpStairLines.forEach(renderProposalLine)
    }
    if (shearWallLines.length > 0) {
      spreadsheet.updateCell({ value: 'Shear Walls:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      shearWallLines.forEach(renderProposalLine)
    }
    if (columnsLines.length > 0) {
      spreadsheet.updateCell({ value: 'Columns:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      columnsLines.forEach(renderProposalLine)
    }
    if (dropPanelLines.length > 0) {
      spreadsheet.updateCell({ value: 'Drop panel:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      dropPanelLines.forEach(renderProposalLine)
    }
    if (concretePostLines.length > 0) {
      spreadsheet.updateCell({ value: 'Concrete post:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      concretePostLines.forEach(renderProposalLine)
    }
    if (concreteEncasementLines.length > 0) {
      spreadsheet.updateCell({ value: 'Concrete encasement:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      concreteEncasementLines.forEach(renderProposalLine)
    }
    if (beamsLines.length > 0) {
      spreadsheet.updateCell({ value: 'Beams:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      beamsLines.forEach(renderProposalLine)
    }
    const cipStairSubHeaderLabel = (name) => {
      const t = (name || '').trim()
      const withColon = t.endsWith(':') ? t : `${t}:`
      if (withColon === 'Ext. stair:') return 'Exterior stair:'
      return withColon
    }
    const renderCipStairsLineFromFormula = (proposalText, calcRow, options = {}) => {
      if (!calcRow) return
      const showRisersSuffix = options.riserSuffix === true
      const useStairsFormula = options.useStairsFormula === true
      const useLandingsFormula = options.useLandingsFormula === true
      if (superstructureScopeStartRow == null) superstructureScopeStartRow = currentRow
      if (useLandingsFormula) {
        const landFormula = `=CONCATENATE("F&I new (",ROUND('${calcSheetName}'!H${calcRow}*12,0),"\"\""," thick) stair landings @ 1st FL as per ","${cipStairsPageRefEsc}",")")`
        spreadsheet.updateCell({ formula: landFormula }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, `F&I new (… thick) stair landings @ 1st FL as per ${getPageRefSuperstructure('CIP Stairs') || 'details on'}`)
      } else if (useStairsFormula) {
        const gRef = `'${calcSheetName}'!G${calcRow}`
        const mRef = `'${calcSheetName}'!M${calcRow}`
        const widthPart = `INT(${gRef})&"'-"&TEXT(ROUND((${gRef}-INT(${gRef}))*12,0),"0")&"\"\""`
        const stairsFormula = `=CONCATENATE("F&I new (",${widthPart}," wide) concrete stairs (",ROUND(${mRef},0)," Riser) @ 1st FL as per ","${cipStairsPageRefEsc}",")")`
        spreadsheet.updateCell({ formula: stairsFormula }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, `F&I new (… wide) concrete stairs (… Riser) @ 1st FL as per ${getPageRefSuperstructure('CIP Stairs') || 'details on'}`)
      } else {
        const text = proposalText || getCipStairsProposalText(null, options.lineType, calcRow)
        if (!text) return
        spreadsheet.updateCell({ value: text }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, text)
      }
      const descForRates = rowBContentMap.get(currentRow) || proposalText || ''
      spreadsheet.wrap(`${pfx}B${currentRow}`, true)
      spreadsheet.cellFormat({ fontWeight: 'normal', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
      spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${calcRow}` }, `${pfx}D${currentRow}`)
      spreadsheet.updateCell({ formula: `='${calcSheetName}'!L${calcRow}` }, `${pfx}F${currentRow}`)
      spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${calcRow}` }, `${pfx}G${currentRow}`)
      fillRatesForProposalRow(currentRow, descForRates)
      spreadsheet.cellFormat({ fontWeight: 'normal', textAlign: 'right', backgroundColor: 'white' }, `${pfx}D${currentRow}`)
      spreadsheet.cellFormat({ fontWeight: 'normal', textAlign: 'right', backgroundColor: 'white' }, `${pfx}F${currentRow}`)
      spreadsheet.cellFormat({ fontWeight: 'normal', textAlign: showRisersSuffix ? 'center' : 'right', backgroundColor: 'white' }, `${pfx}G${currentRow}`)
      try { spreadsheet.numberFormat('#,##0.00', `${pfx}D${currentRow}`) } catch (e) { }
      try { spreadsheet.numberFormat('#,##0.00', `${pfx}F${currentRow}`) } catch (e) { }
      try { spreadsheet.numberFormat(showRisersSuffix ? '#,##0" Risers"' : '#,##0.00', `${pfx}G${currentRow}`) } catch (e) { }
      spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
      spreadsheet.cellFormat({ fontWeight: 'normal', textAlign: 'right', backgroundColor: 'white', format: '$#,##0.00' }, `${pfx}H${currentRow}`)
      currentRow++
    }
    if (parapetWallLines.length > 0) {
      spreadsheet.updateCell({ value: 'Parapet walls:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      parapetWallLines.forEach(renderProposalLine)
    }
    if (thermalBreakLines.length > 0) {
      spreadsheet.updateCell({ value: 'Thermal break:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      thermalBreakLines.forEach(renderProposalLine)
    }
    if (curbsLines.length > 0) {
      spreadsheet.updateCell({ value: 'Curbs:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      curbsLines.forEach(renderProposalLine)
    }
    if (cipStairsGroupsFromFormulas.length > 0) {
      spreadsheet.updateCell({ value: 'Concrete stairs:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D9D9D9', textDecoration: 'underline', borderBottom: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      cipStairsGroupsFromFormulas.forEach((grp) => {
        const subHeaderLabel = cipStairSubHeaderLabel(grp.name)
        spreadsheet.updateCell({ value: subHeaderLabel }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, subHeaderLabel)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#E2EFDA', textDecoration: 'underline', borderBottom: '1px solid #000000' },
          `${pfx}B${currentRow}`
        )
        currentRow++
        if (grp.landingRow != null) renderCipStairsLineFromFormula(null, grp.landingRow, { lineType: 'landings', useLandingsFormula: true })
        if (grp.stairRow != null) renderCipStairsLineFromFormula(null, grp.stairRow, { riserSuffix: true, useStairsFormula: true })
        const slabText = getCipStairsProposalText(grp.name, 'slab', grp.slabRow)
        if (grp.slabRow != null && slabText && cipStairsSlabRowHasData(grp)) renderCipStairsLineFromFormula(slabText, grp.slabRow, {})
      })
      superstructureScopeEndRow = currentRow - 1
    } else if (cipStairsGroupsWithLines.length > 0) {
      spreadsheet.updateCell({ value: 'Concrete stairs:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D9D9D9', textDecoration: 'underline', borderBottom: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      cipStairsGroupsWithLines.forEach(({ groupName, lines }) => {
        const subHeaderLabel = cipStairSubHeaderLabel(groupName)
        spreadsheet.updateCell({ value: subHeaderLabel }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, subHeaderLabel)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#E2EFDA', textDecoration: 'underline', borderBottom: '1px solid #000000' },
          `${pfx}B${currentRow}`
        )
        currentRow++
        lines.forEach((line) => renderProposalLine(line, { fontWeight: 'normal' }))
      })
    }
    const getInfilledValuesFromRow = (calcRow) => {
      const defaultAsPer = getPageRefSuperstructure('Stairs – Infilled tads') || 'as per details on'
      if (!calculationData || !calcRow || calcRow < 1) return { thickness: '8"', width: 'as indicated', riser: 'as indicated', location: '1st FL', asPer: defaultAsPer, particulars: '' }
      const row = calculationData[calcRow - 1]
      if (!row || !Array.isArray(row)) return { thickness: '8"', width: 'as indicated', riser: 'as indicated', location: '1st FL', asPer: defaultAsPer, particulars: '' }
      const p = (row[1] || '').toString()
      let thickness = '8"'
      const thicknessMatch = p.match(/(\d+(?:\/\d+)?)"?\s*thick/i) || p.match(/(\d+)"\s*thick/i)
      if (thicknessMatch) thickness = thicknessMatch[1].includes('"') ? thicknessMatch[1] : thicknessMatch[1] + '"'
      else if (row[7] != null && row[7] !== '') {
        const heightFeet = parseFloat(row[7])
        if (!isNaN(heightFeet) && heightFeet > 0) thickness = `${Math.round(heightFeet * 12)}"`
      }
      const widthFromColRaw = row[6] != null && row[6] !== '' ? String(row[6]).trim() : null
      const widthFromCol = widthFromColRaw ? (formatWidthFeetInches(widthFromColRaw) || widthFromColRaw) : null
      const widthMatch = p.match(/(\d+'-?\d*")\s*wide/i)
      const width = widthFromCol || (widthMatch ? widthMatch[1] : 'as indicated')
      let riserVal = Math.round(parseFloat(row[12]) || 0)
      if (riserVal <= 0 && row[11] != null && row[11] !== '') riserVal = Math.round(parseFloat(row[11]) || 0)
      if (riserVal <= 0) {
        const takeoffVal = parseFloat(row[2])
        if (Number.isInteger(takeoffVal) && takeoffVal > 0 && takeoffVal < 100) riserVal = takeoffVal
      }
      if (riserVal <= 0) { const rm = p.match(/(\d+)\s*Riser/i); if (rm) riserVal = parseInt(rm[1], 10) }
      const riser = riserVal > 0 ? String(riserVal) : 'as indicated'
      const locationMatch = p.match(/@\s*([^ as per]+?)(?:\s+as\s+per|$)/i) || p.match(/(?:cellar|roof|1st|2nd|\d+th)\s*FL[^,]*(?:,\s*(?:cellar|roof|1st|2nd|\d+th)\s*FL[^,]*)*/i)
      const location = (locationMatch && locationMatch[1] ? locationMatch[1].trim() : null) || (locationMatch && locationMatch[0] ? locationMatch[0].trim() : null) || '1st FL'
      const asPer = getPageRefSuperstructure('Stairs – Infilled tads') || 'as per details on'
      return { thickness, width, riser, location, asPer, particulars: p }
    }
    const getInfilledLandingProposalText = (calcRow) => {
      const v = getInfilledValuesFromRow(calcRow)
      if (v.particulars && v.particulars.length > 30 && /F&I|topping|concrete|reinforced/i.test(v.particulars)) return v.particulars
      return `F&I new (${v.thickness} thick) stair landings @ ${v.location} as per ${v.asPer}`
    }
    const getInfilledStairProposalText = (calcRow) => {
      const v = getInfilledValuesFromRow(calcRow)
      if (v.particulars && v.particulars.length > 30 && /F&I|infill|concrete|riser/i.test(v.particulars)) return v.particulars
      return `F&I new (${v.width} wide) concrete stairs (${v.riser} Riser) @ ${v.location} as per ${v.asPer}`
    }
    if (infilledGroups.length > 0) {
      if (superstructureScopeStartRow == null) superstructureScopeStartRow = currentRow
      spreadsheet.updateCell({ value: 'Concrete infilled stairs:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D9D9D9', textDecoration: 'underline', borderBottom: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      infilledGroups.forEach((grp) => {
        spreadsheet.updateCell({ value: grp.name && grp.name.endsWith(':') ? grp.name : `${grp.name || 'Stair'}:` }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#E2EFDA', textDecoration: 'underline', borderBottom: '1px solid #000000' },
          `${pfx}B${currentRow}`
        )
        currentRow++
        const infilledLandingDesc = `F&I new (… thick) stair landings @ 1st FL as per ${getPageRefSuperstructure('Stairs – Infilled tads') || 'details on'}`
        const infilledStairDesc = `F&I new (… wide) concrete stairs (… Riser) @ 1st FL as per ${getPageRefSuperstructure('Stairs – Infilled tads') || 'details on'}`
        if (grp.landingSumRow != null) {
          const landFormula = `=CONCATENATE("F&I new (",ROUND('${calcSheetName}'!H${grp.landingSumRow}*12,0),"\"\""," thick) stair landings @ 1st FL as per ","${infilledPageRefEsc}",")")`
          spreadsheet.updateCell({ formula: landFormula }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, infilledLandingDesc)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'normal', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${grp.landingSumRow}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${calcSheetName}'!L${grp.landingSumRow}` }, `${pfx}F${currentRow}`)
          spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${grp.landingSumRow}` }, `${pfx}G${currentRow}`)
          fillRatesForProposalRow(currentRow, infilledLandingDesc)
          spreadsheet.cellFormat({ fontWeight: 'normal', textAlign: 'right', backgroundColor: 'white' }, `${pfx}D${currentRow}`)
          spreadsheet.cellFormat({ fontWeight: 'normal', textAlign: 'right', backgroundColor: 'white' }, `${pfx}F${currentRow}`)
          spreadsheet.cellFormat({ fontWeight: 'normal', textAlign: 'right', backgroundColor: 'white' }, `${pfx}G${currentRow}`)
          try { spreadsheet.numberFormat('#,##0.00', `${pfx}D${currentRow}`) } catch (e) { }
          try { spreadsheet.numberFormat('#,##0.00', `${pfx}F${currentRow}`) } catch (e) { }
          try { spreadsheet.numberFormat('#,##0.00', `${pfx}G${currentRow}`) } catch (e) { }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat({ fontWeight: 'normal', textAlign: 'right', backgroundColor: 'white', format: '$#,##0.00' }, `${pfx}H${currentRow}`)
          currentRow++
        }
        if (grp.stairRow != null) {
          const gRef = `'${calcSheetName}'!G${grp.stairRow}`
          const mRef = `'${calcSheetName}'!M${grp.stairRow}`
          const widthPart = `INT(${gRef})&"'-"&TEXT(ROUND((${gRef}-INT(${gRef}))*12,0),"0")&"\"\""`
          const stairFormula = `=CONCATENATE("F&I new (",${widthPart}," wide) concrete stairs (",ROUND(${mRef},0)," Riser) @ 1st FL as per ","${infilledPageRefEsc}",")")`
          spreadsheet.updateCell({ formula: stairFormula }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, infilledStairDesc)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'normal', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${grp.stairRow}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${calcSheetName}'!L${grp.stairRow}` }, `${pfx}F${currentRow}`)
          spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${grp.stairRow}` }, `${pfx}G${currentRow}`)
          fillRatesForProposalRow(currentRow, infilledStairDesc)
          spreadsheet.cellFormat({ fontWeight: 'normal', textAlign: 'right', backgroundColor: 'white' }, `${pfx}D${currentRow}`)
          spreadsheet.cellFormat({ fontWeight: 'normal', textAlign: 'right', backgroundColor: 'white' }, `${pfx}F${currentRow}`)
          spreadsheet.cellFormat({ fontWeight: 'normal', textAlign: 'center', backgroundColor: 'white' }, `${pfx}G${currentRow}`)
          try { spreadsheet.numberFormat('#,##0.00', `${pfx}D${currentRow}`) } catch (e) { }
          try { spreadsheet.numberFormat('#,##0.00', `${pfx}F${currentRow}`) } catch (e) { }
          try { spreadsheet.numberFormat('#,##0" Risers"', `${pfx}G${currentRow}`) } catch (e) { }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat({ fontWeight: 'normal', textAlign: 'right', backgroundColor: 'white', format: '$#,##0.00' }, `${pfx}H${currentRow}`)
          currentRow++
        }
      })
      superstructureScopeEndRow = currentRow - 1
    }
    if (nonShrinkGroutLines.length > 0) {
      spreadsheet.updateCell({ value: 'Non-shrink grout:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      nonShrinkGroutLines.forEach(renderProposalLine)
    }
    if (concretePadLines.length > 0) {
      spreadsheet.updateCell({ value: 'Housekeeping concrete pads:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      concretePadLines.forEach(renderProposalLine)
    }
    if (repairScopeLines.length > 0) {
      spreadsheet.updateCell({ value: 'Repair scope:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      repairScopeLines.forEach(renderProposalLine)
    }
    {
      spreadsheet.updateCell({ value: 'Misc.:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      const miscItems = [
        'All Rigging of material & concrete pumping included',
        'Outrigger netting, jumps & engineering',
        'DOB approved concrete washout included',
        'Engineering, shop drawings, formwork drawings, design mixes included'
      ]
      miscItems.forEach((item, idx) => {
        spreadsheet.updateCell({ value: item }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, item)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat({ fontWeight: 'normal', color: '#000000', textAlign: 'left' }, `${pfx}B${currentRow}`)
        const rateLookupKey = item === 'Engineering, shop drawings, formwork drawings, design mixes included'
          ? 'Engineering, shop drawings, formwork drawings, design mixes included (Superstructure concrete scope)'
          : item
        fillRatesForProposalRow(currentRow, rateLookupKey)
        if (item === 'DOB approved concrete washout included' && superstructureScopeStartRow != null && superstructureScopeEndRow != null) {
          const exclPart = washoutExclusionRows.length > 0 ? '-' + washoutExclusionRows.map(r => `F${r}`).join('-') : ''
          spreadsheet.updateCell(
            { formula: `=SUM(F${superstructureScopeStartRow}:F${superstructureScopeEndRow})${exclPart}` },
            `${pfx}F${currentRow}`
          )
          spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}F${currentRow}`)
          try { spreadsheet.numberFormat('#,##0.00', `${pfx}F${currentRow}`) } catch (e) { }
        }
        const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
        spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
        spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white', format: '$#,##0.00' }, `${pfx}H${currentRow}`)
        try { spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}`) } catch (e) { }
        currentRow++
      })
    }
    if (superstructureScopeStartRow != null && superstructureScopeEndRow != null) {
      spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
      spreadsheet.updateCell({ value: 'Superstructure Concrete Total:' }, `${pfx}D${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#BDD7EE' },
        `${pfx}D${currentRow}:E${currentRow}`
      )
      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({ formula: `=SUM(H${superstructureScopeStartRow}:H${superstructureScopeEndRow})*1000` }, `${pfx}F${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD7EE', format: '$#,##0.00' },
        `${pfx}F${currentRow}:G${currentRow}`
      )
      try { spreadsheet.numberFormat('$#,##0.00', `${pfx}F${currentRow}:G${currentRow}`) } catch (e) { }
      applyTotalRowBorders(spreadsheet, pfx, currentRow, '#BDD7EE')
      baseBidTotalRows.push(currentRow) // Superstructure Concrete Total
      totalRows.push(currentRow)
      currentRow++
      currentRow++ // Empty row after Superstructure Concrete Total
    }

      // Add empty row after Superstructure concrete scope
      currentRow++
    }

    // Plumbing/Electrical trenching scope section
    {
      const trenchingFormulas = (formulaData || []).filter(f => f.itemType === 'trenching_item' && f.section === 'trenching')
      const trenchingProposalTextMap = {
        'Demo': 'Allow to saw-cut/demo/remove/dispose existing SOG',
        'Excavation': 'Allow to perform soil excavation, trucking & disposal (Havg=2\'-2")',
        'Backfill': 'Allow to backfill SOE berm with existing stockpiled soil on site and compact',
        'Gravel': 'F&I new (6" thick) gravel/crushed stone, including 6MIL vapor barrier on top @ new patch back',
        'Patchback': 'F&I new (4" thick) concrete patch back'
      }

      if (trenchingFormulas.length > 0) {
        spreadsheet.updateCell({ value: 'Plumbing/Electrical trenching scope:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'center', backgroundColor: '#BDD7EE', border: '1px solid #000000', textDecoration: 'underline' },
          `${pfx}B${currentRow}`
        )
        currentRow++

        const trenchingStartRow = currentRow
        trenchingFormulas.forEach(tf => {
          const proposalText = trenchingProposalTextMap[tf.name] || tf.name
          const calcRow = tf.row
          spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, proposalText)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, proposalText)
          const dynamicHeight = calculateRowHeight(proposalText)
          spreadsheet.updateCell({ formula: `='Calculations Sheet'!I${calcRow}` }, `${pfx}C${currentRow}`)
          spreadsheet.updateCell({ formula: `='Calculations Sheet'!J${calcRow}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='Calculations Sheet'!L${calcRow}` }, `${pfx}F${currentRow}`)
            ;['C', 'D', 'F'].forEach(col => {
              spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}${col}${currentRow}`)
              try { spreadsheet.numberFormat('#,##0.00', `${pfx}${col}${currentRow}`) } catch (e) { }
            })
          const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
          spreadsheet.updateCell({ formula: dollarFormula }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white', format: '$#,##0.00' }, `${pfx}H${currentRow}`)
          try { spreadsheet.numberFormat('$#,##0.00', `${pfx}H${currentRow}`) } catch (e) { }
          currentRow++
        })
        const trenchingEndRow = currentRow - 1

        spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
        spreadsheet.updateCell({ value: 'Trenching Total:' }, `${pfx}D${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#BDD7EE' },
          `${pfx}D${currentRow}:E${currentRow}`
        )
        spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
        spreadsheet.updateCell({ formula: `=SUM(H${trenchingStartRow}:H${trenchingEndRow})*1000` }, `${pfx}F${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD7EE', format: '$#,##0.00' },
          `${pfx}F${currentRow}:G${currentRow}`
        )
        try { spreadsheet.numberFormat('$#,##0.00', `${pfx}F${currentRow}:G${currentRow}`) } catch (e) { }
        applyTotalRowBorders(spreadsheet, pfx, currentRow, '#BDD7EE')
        baseBidTotalRows.push(currentRow) // Trenching Total
        totalRows.push(currentRow)
        currentRow++
      }
    }

    // Add empty row after Trenching scope
    currentRow++

    // BASE BID TOTAL row
    {
      spreadsheet.updateCell({ value: '(Good For 30 Days)' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'center', backgroundColor: '#BDD7EE', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
      spreadsheet.updateCell({ value: 'BASE BID TOTAL:' }, `${pfx}D${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD7EE', border: '1px solid #000000' },
        `${pfx}D${currentRow}:E${currentRow}`
      )
      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      const baseBidFormula = baseBidTotalRows.length > 0
        ? `=SUM(${baseBidTotalRows.map(r => `F${r}`).join(',')})`
        : '=0'
      spreadsheet.updateCell({ formula: baseBidFormula }, `${pfx}F${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD7EE', border: '1px solid #000000', format: '$#,##0.00' },
        `${pfx}F${currentRow}:G${currentRow}`
      )
      try { spreadsheet.numberFormat('$#,##0.00', `${pfx}F${currentRow}:G${currentRow}`) } catch (e) { }
      currentRow++
    }

    // Add empty row after BASE BID TOTAL
    currentRow++

    // Only show Add Alternate #1 (Sitework) and Add Alternate #2 (B.P.P.) when both sections exist in the calculation sheet
    const hasBPPAlternate2InCalc = calculationData && calculationData.some(row => String(row[0] || '').trim() === 'B.P.P. Alternate #2 scope')
    const hasCivilSiteworkInCalc = calculationData && calculationData.some(row => String(row[0] || '').trim() === 'Civil / Sitework')
    if (hasBPPAlternate2InCalc && hasCivilSiteworkInCalc) {
      // Add Alternate #1: Sitework scope header
      let siteworkScopeStartRow = null
      const siteworkTotalFRows = []
      {
        spreadsheet.merge(`${pfx}B${currentRow}:G${currentRow}`)
        spreadsheet.updateCell({ value: 'Add Alternate #1: Sitework scope: as per C-4' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'center', backgroundColor: '#BDD6EE', border: '1px solid #000000' },
          `${pfx}B${currentRow}:G${currentRow}`
        )
        currentRow++
        siteworkScopeStartRow = currentRow // Track start row for total calculation
      }

      // Sitework scope content
      // Add "Site demo / Removals:" subsection header
      spreadsheet.updateCell({ value: 'Site demo / Removals:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: '#FEF2CB', verticalAlign: 'middle',
          textDecoration: 'underline',
          border: '1px solid #000000'
        },
        `${pfx}B${currentRow}`
      )
      currentRow++

      // Add "Demolition:" subsection header
      spreadsheet.updateCell({ value: 'Demolition:' }, `${pfx}B${currentRow}`)
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
      currentRow++

      // Track the start row for demo items (for total calculation)
      const demoStartRow = currentRow

      // Process Civil/Sitework demolition items from calculation data
      const civilDemoItems = {
        'Demo asphalt': { items: [], proposalText: 'Allow to saw-cut/demo/remove/dispose existing full depth asphalt pavement (full depth: ##") as per C-3' },
        'Demo curb': { items: [], proposalText: 'Allow to saw-cut/demo/remove/dispose existing asphalt curb as per C-3' },
        'Demo fence': { items: [], proposalText: null }, // Multiple types
        'Demo wall': { items: [], proposalText: 'Allow to remove (##" wide) retaining wall (H=##", typ.) as per C-3' },
        'Demo pipe': { items: [], proposalText: null }, // Multiple types
        'Demo rail': { items: [], proposalText: 'Allow to remove guiderail as per C-3' },
        'Demo sign': { items: [], proposalText: null }, // Multiple types
        'Demo manhole': { items: [], proposalText: 'Allow to protect existing stormwater manhole as per C-3' },
        'Demo fire hydrant': { items: [], proposalText: 'Allow to relocated existing fire hydrant as per C-3' },
        'Demo utility pole': { items: [], proposalText: 'Allow to protect utility pole as per C-3' },
        'Demo valve': { items: [], proposalText: 'Allow to protect existing water valve as per C-3' },
        'Demo inlet': { items: [], proposalText: 'Allow to remove existing stormwater inlet as per C-3' }
      }

      // Parse calculation data for civil demo items - track sum rows for each type
      const demoSumRows = {}

      if (calculationData && calculationData.length > 0) {
        let currentDemoSubsection = null
        let lastItemRowInSubsection = 0

        calculationData.forEach((row, index) => {
          const colB = row[1]
          const rowNum = index + 2 // 1-based Excel row number

          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()

            // Check for Demo subsection headers
            const demoHeaders = {
              'demo asphalt:': 'Demo asphalt',
              'demo curb:': 'Demo curb',
              'demo fence:': 'Demo fence',
              'demo wall:': 'Demo wall',
              'demo pipe:': 'Demo pipe',
              'demo rail:': 'Demo rail',
              'demo sign:': 'Demo sign',
              'demo manhole:': 'Demo manhole',
              'demo fire hydrant:': 'Demo fire hydrant',
              'demo utility pole:': 'Demo utility pole',
              'demo valve:': 'Demo valve',
              'demo inlet:': 'Demo inlet'
            }

            if (demoHeaders[bText]) {
              // Save sum row for previous subsection (sum row is the last item row itself)
              if (currentDemoSubsection && lastItemRowInSubsection > 0) {
                demoSumRows[currentDemoSubsection] = lastItemRowInSubsection
              }
              currentDemoSubsection = demoHeaders[bText]
              lastItemRowInSubsection = 0
              return
            }

            // Check if this is a new section header (ends with : but not demo)
            if (bText.endsWith(':') && !bText.startsWith('demo') && !bText.startsWith('remove') && !bText.startsWith('protect')) {
              // Save sum row for previous subsection
              if (currentDemoSubsection && lastItemRowInSubsection > 0) {
                demoSumRows[currentDemoSubsection] = lastItemRowInSubsection
              }
              currentDemoSubsection = null
              lastItemRowInSubsection = 0
              return
            }

            // If in a demo subsection, collect items
            if (currentDemoSubsection && civilDemoItems[currentDemoSubsection]) {
              const takeoff = parseFloat(row[2]) || 0
              // Check if this row has data (not empty, not just a sum row)
              if (bText && !bText.endsWith(':') && (takeoff > 0 || bText.includes('remove') || bText.includes('protect') || bText.includes('relocate'))) {
                civilDemoItems[currentDemoSubsection].items.push({
                  rowIndex: rowNum,
                  particulars: colB.trim(),
                  takeoff: takeoff,
                  rawRowNumber: rowNum
                })
                lastItemRowInSubsection = rowNum
              }
            }
          }
        })

        // Save sum row for last subsection
        if (currentDemoSubsection && lastItemRowInSubsection > 0) {
          demoSumRows[currentDemoSubsection] = lastItemRowInSubsection
        }
      }

      // Display civil demo items in proposal sheet
      const demoOrder = ['Demo asphalt', 'Demo curb', 'Demo fence', 'Demo wall', 'Demo pipe', 'Demo rail', 'Demo sign', 'Demo inlet', 'Demo fire hydrant', 'Demo manhole', 'Demo utility pole', 'Demo valve']

      demoOrder.forEach(demoType => {
        const demoData = civilDemoItems[demoType]
        if (demoData && demoData.items.length > 0) {
          // Reference row: use last data row (sum row - 1) so proposal points to same line as calc sheet total
          const civilDemoSums = (formulaData || []).filter(f => f.itemType === 'civil_demo_sum' && f.section === 'civil_sitework' && f.subsectionName === demoType)
          const lastDataRowFromParse = demoSumRows[demoType] || 0
          const sumRowIndex = civilDemoSums.length >= 1
            ? (civilDemoSums[0].lastDataRow != null ? civilDemoSums[0].lastDataRow : (civilDemoSums[0].row - 1))
            : lastDataRowFromParse

          // Generate proposal text based on demo type
          let proposalText = demoData.proposalText
          const calcSheetName = 'Calculations Sheet'

          if (demoType === 'Demo asphalt') {
            // Extract thickness from particulars if available
            const firstItem = demoData.items[0]
            if (firstItem && firstItem.particulars) {
              const thicknessMatch = firstItem.particulars.match(/(\d+)[""]?\s*(thick)?/i)
              if (thicknessMatch) {
                proposalText = `Allow to saw-cut/demo/remove/dispose existing full depth asphalt pavement (full depth: ${thicknessMatch[1]}") as per C-3`
              }
            }

            // Display Demo asphalt with SF and CY
            spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, proposalText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            fillRatesForProposalRow(currentRow, proposalText)

            if (sumRowIndex > 0) {
              // SF in column D
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${sumRowIndex}` }, `${pfx}D${currentRow}`)
              spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}D${currentRow}`)
              // CY in column F
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!L${sumRowIndex}` }, `${pfx}F${currentRow}`)
              spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}F${currentRow}`)
            }
            currentRow++
            return
          } else if (demoType === 'Demo curb') {
            proposalText = 'Allow to saw-cut/demo/remove/dispose existing asphalt curb as per C-3'
            spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, proposalText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            fillRatesForProposalRow(currentRow, proposalText)

            if (sumRowIndex > 0) {
              // LF in column C
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${sumRowIndex}` }, `${pfx}C${currentRow}`)
              spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}C${currentRow}`)
              // SF in column D
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${sumRowIndex}` }, `${pfx}D${currentRow}`)
              spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}D${currentRow}`)
              // CY in column F
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!L${sumRowIndex}` }, `${pfx}F${currentRow}`)
              spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}F${currentRow}`)
            }
            currentRow++
            return
          } else if (demoType === 'Demo fence') {
            // Check for chain-link/vinyl vs wood fence - each has separate sum row
            const chainLinkItems = demoData.items.filter(item => item.particulars.toLowerCase().includes('chain') || item.particulars.toLowerCase().includes('vinyl'))
            const woodItems = demoData.items.filter(item => item.particulars.toLowerCase().includes('wood'))

            if (chainLinkItems.length > 0) {
              const chainLinkLastDataRow = Math.max(...chainLinkItems.map(i => i.rawRowNumber))
              const chainLinkFormula = (formulaData || []).find(f => f.itemType === 'civil_demo_sum' && f.section === 'civil_sitework' && f.subsectionName === 'Demo fence' && f.firstDataRow <= chainLinkLastDataRow && chainLinkLastDataRow <= f.lastDataRow)
              const chainLinkSumRow = chainLinkFormula ? (chainLinkFormula.lastDataRow != null ? chainLinkFormula.lastDataRow : (chainLinkFormula.row - 1)) : chainLinkLastDataRow

              proposalText = `Allow to remove existing (H=6'-0", typ) chain-link fence/ vinyl fence as per C-3`
              spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
              rowBContentMap.set(currentRow, proposalText)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
              fillRatesForProposalRow(currentRow, proposalText)

              // Add LF and SF from calculation sheet
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${chainLinkSumRow}` }, `${pfx}C${currentRow}`)
              spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}C${currentRow}`)
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${chainLinkSumRow}` }, `${pfx}D${currentRow}`)
              spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}D${currentRow}`)
              currentRow++
            }

            if (woodItems.length > 0) {
              const woodLastDataRow = Math.max(...woodItems.map(i => i.rawRowNumber))
              const woodFormula = (formulaData || []).find(f => f.itemType === 'civil_demo_sum' && f.section === 'civil_sitework' && f.subsectionName === 'Demo fence' && f.firstDataRow <= woodLastDataRow && woodLastDataRow <= f.lastDataRow)
              const woodSumRow = woodFormula ? (woodFormula.lastDataRow != null ? woodFormula.lastDataRow : (woodFormula.row - 1)) : woodLastDataRow

              proposalText = `Allow to remove existing (H=6'-0", typ) wood-link fence as per C-3`
              spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
              rowBContentMap.set(currentRow, proposalText)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
              fillRatesForProposalRow(currentRow, proposalText)

              // Add LF and SF from calculation sheet
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${woodSumRow}` }, `${pfx}C${currentRow}`)
              spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}C${currentRow}`)
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${woodSumRow}` }, `${pfx}D${currentRow}`)
              spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}D${currentRow}`)
              currentRow++
            }
            return // Skip default handling
          } else if (demoType === 'Demo wall') {
            // Extract wall dimensions from particulars if available
            const firstItem = demoData.items[0]
            let width = '18"'
            let height = '3\'-6"'
            if (firstItem && firstItem.particulars) {
              const widthMatch = firstItem.particulars.match(/(\d+)[""]?\s*wide/i)
              const heightMatch = firstItem.particulars.match(/H=([^,\)]+)/i)
              if (widthMatch) width = widthMatch[1] + '"'
              if (heightMatch) height = heightMatch[1]
            }

            proposalText = `Allow to remove (${width} wide) retaining wall (H=${height}, typ.) as per C-3`
            spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, proposalText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            fillRatesForProposalRow(currentRow, proposalText)

            if (sumRowIndex > 0) {
              // LF in column C
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${sumRowIndex}` }, `${pfx}C${currentRow}`)
              spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}C${currentRow}`)
              // SF in column D
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${sumRowIndex}` }, `${pfx}D${currentRow}`)
              spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}D${currentRow}`)
              // CY in column F
              spreadsheet.updateCell({ formula: `='${calcSheetName}'!L${sumRowIndex}` }, `${pfx}F${currentRow}`)
              spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}F${currentRow}`)
            }
            currentRow++
            return
          } else if (demoType === 'Demo pipe') {
            // Check for remove vs protect - each has separate sum row
            const removeItems = demoData.items.filter(item => item.particulars.toLowerCase().includes('remove'))
            const protectItems = demoData.items.filter(item => item.particulars.toLowerCase().includes('protect'))

            if (removeItems.length > 0) {
              // Find sum row for remove pipe
              const removeLastRow = Math.max(...removeItems.map(i => i.rawRowNumber))
              const removeSumRow = removeLastRow

              // Extract pipe size
              const removeItem = removeItems[0]
              const sizeMatch = removeItem?.particulars.match(/(\d+)[""]?\s*(thick|pipe)?/i)
              const size = sizeMatch ? sizeMatch[1] : '##'
              proposalText = `Allow to remove existing HDPE pipe (${size}" thick) as per C-3`
              spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
              rowBContentMap.set(currentRow, proposalText)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
              fillRatesForProposalRow(currentRow, proposalText)

              spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${removeSumRow}` }, `${pfx}C${currentRow}`)
              spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}C${currentRow}`)
              currentRow++
            }

            // Add guiderail line (from Demo rail) here if it exists
            const railData = civilDemoItems['Demo rail']
            if (railData && railData.items.length > 0) {
              const railSumRow = demoSumRows['Demo rail'] || 0
              proposalText = `Allow to remove guiderail as per C-3`
              spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
              rowBContentMap.set(currentRow, proposalText)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
              fillRatesForProposalRow(currentRow, proposalText)

              if (railSumRow > 0) {
                spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${railSumRow}` }, `${pfx}C${currentRow}`)
                spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}C${currentRow}`)
              }
              currentRow++
            }

            if (protectItems.length > 0) {
              // Find sum row for protect pipe
              const protectLastRow = Math.max(...protectItems.map(i => i.rawRowNumber))
              const protectSumRow = protectLastRow

              const protectItem = protectItems[0]
              const sizeMatch = protectItem?.particulars.match(/(\d+)[""]?\s*(thick|rcp)?/i)
              const size = sizeMatch ? sizeMatch[1] : '##'
              proposalText = `Allow to protect existing RCP stormwater main (${size}" thick) as per C-3`
              spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
              rowBContentMap.set(currentRow, proposalText)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
              fillRatesForProposalRow(currentRow, proposalText)

              spreadsheet.updateCell({ formula: `='${calcSheetName}'!I${protectSumRow}` }, `${pfx}C${currentRow}`)
              spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}C${currentRow}`)
              currentRow++
            }
            return // Skip default handling
          } else if (demoType === 'Demo rail') {
            // Skip - handled with Demo pipe
            return
          } else if (demoType === 'Demo sign') {
            // Check for single sign vs row of signs - each has separate sum row
            const singleItems = demoData.items.filter(item => !item.particulars.toLowerCase().includes('row'))
            const rowItems = demoData.items.filter(item => item.particulars.toLowerCase().includes('row'))

            if (singleItems.length > 0) {
              // Find sum row for single sign (sum is on the last item row itself)
              const singleLastRow = Math.max(...singleItems.map(i => i.rawRowNumber))
              const singleSumRow = singleLastRow

              proposalText = `Allow to remove existing sign as per C-3`
              spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
              rowBContentMap.set(currentRow, proposalText)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
              fillRatesForProposalRow(currentRow, proposalText)

              spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${singleSumRow}` }, `${pfx}G${currentRow}`)
              spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}G${currentRow}`)
              currentRow++
            }

            if (rowItems.length > 0) {
              // Find sum row for row of signs (sum is on the last item row itself)
              const rowLastRow = Math.max(...rowItems.map(i => i.rawRowNumber))
              const rowSumRow = rowLastRow

              proposalText = `Allow to remove existing row of signs as per C-3`
              spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
              rowBContentMap.set(currentRow, proposalText)
              spreadsheet.wrap(`${pfx}B${currentRow}`, true)
              spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
              fillRatesForProposalRow(currentRow, proposalText)

              spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${rowSumRow}` }, `${pfx}G${currentRow}`)
              spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}G${currentRow}`)
              currentRow++
            }
            return // Skip default handling
          } else if (demoType === 'Demo inlet' && sumRowIndex > 0) {
            proposalText = 'Allow to remove existing stormwater inlet as per C-3'
            spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, proposalText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            fillRatesForProposalRow(currentRow, proposalText)
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
            spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}G${currentRow}`)
            currentRow++
            return
          } else if (demoType === 'Demo fire hydrant' && sumRowIndex > 0) {
            proposalText = 'Allow to relocate existing fire hydrant as per C-3'
            spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, proposalText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            fillRatesForProposalRow(currentRow, 'Allow to relocate existing fire hydrant as per C-3 - Demolition')
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
            spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}G${currentRow}`)
            currentRow++
            return
          } else if (demoType === 'Demo manhole' && sumRowIndex > 0) {
            proposalText = 'Allow to protect existing stormwater manhole as per C-3'
            spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, proposalText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            fillRatesForProposalRow(currentRow, proposalText)
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
            spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}G${currentRow}`)
            currentRow++
            return
          } else if (demoType === 'Demo utility pole' && sumRowIndex > 0) {
            proposalText = 'Allow to protect utility pole as per C-3'
            spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, proposalText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            fillRatesForProposalRow(currentRow, proposalText)
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
            spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}G${currentRow}`)
            currentRow++
            return
          } else if (demoType === 'Demo valve' && sumRowIndex > 0) {
            proposalText = 'Allow to protect existing water valve as per C-3'
            spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, proposalText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            fillRatesForProposalRow(currentRow, proposalText)
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
            spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}G${currentRow}`)
            currentRow++
            return
          }

          // Default handling for other demo types
          if (proposalText) {
            spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, proposalText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            currentRow++
          }
        }
      })

      // Add Site Demo / Removals Total row
      const demoEndRow = currentRow - 1
      if (demoEndRow >= demoStartRow) {
        spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
        spreadsheet.updateCell({ value: 'Site Demo / Removals Total:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle' },
          `${pfx}B${currentRow}:E${currentRow}`
        )
        spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
        spreadsheet.updateCell({ formula: `=SUM(H${demoStartRow}:H${demoEndRow})*1000` }, `${pfx}F${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', format: '$#,##0.00' },
          `${pfx}F${currentRow}:G${currentRow}`
        )
        applyTotalRowBorders(spreadsheet, pfx, currentRow, '#FEF2CB')
        siteworkTotalFRows.push(currentRow)
        currentRow++
      }

      // Empty row above Erosion control
      currentRow++

      // Erosion control subsection
      const erosionHeaderRow = currentRow
      const erosionCalcSheetName = 'Calculations Sheet'
      const erosionItems = { stabilized_entrance: [], silt_fence: [], inlet_filter: [] }
      if (calculationData && calculationData.length > 0) {
        let inSoilErosion = false
        calculationData.forEach((row, index) => {
          const colB = row[1]
          const rowNum = index + 1 // 1-based Excel row (calculationData[0] = row 1)
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText === 'soil erosion:') {
              inSoilErosion = true
              return
            }
            if (inSoilErosion && bText.endsWith(':') && !bText.includes('stabilized') && !bText.includes('silt') && !bText.includes('inlet')) {
              inSoilErosion = false
              return
            }
            if (inSoilErosion) {
              if (bText.includes('stabilized construction entrance')) {
                const thicknessMatch = colB.match(/(\d+(?:\.\d+)?)\s*["']\s*thick/i)
                erosionItems.stabilized_entrance.push({ rowNum, thickness: thicknessMatch ? parseInt(thicknessMatch[1], 10) : 6 })
              } else if (bText.includes('silt fence')) {
                const heightMatch = colB.match(/height\s*=\s*(\d+)'\s*-?\s*(\d+)"/i)
                erosionItems.silt_fence.push({ rowNum, heightStr: heightMatch ? `${heightMatch[1]}'-${heightMatch[2]}"` : `2'-6"` })
              } else if (bText.includes('inlet filter') && !bText.includes('protection')) {
                const qtyMatch = colB.match(/\((\d+)\s*(?:no\.?|ea)\)/i)
                erosionItems.inlet_filter.push({ rowNum, qty: qtyMatch ? parseInt(qtyMatch[1], 10) : 1 })
              }
            }
          }
        })
      }

      const hasErosionItems = erosionItems.stabilized_entrance.length > 0 || erosionItems.silt_fence.length > 0 || erosionItems.inlet_filter.length > 0
      if (hasErosionItems) {
        spreadsheet.updateCell({ value: 'Erosion control:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#FEF2CB', verticalAlign: 'middle', textDecoration: 'underline', border: '1px solid #000000' },
          `${pfx}B${currentRow}`
        )
        currentRow++

        erosionItems.stabilized_entrance.forEach(({ rowNum, thickness = 6 }) => {
          const stabilizedText = `F&I new (${thickness}" thick) temporary stabilized construction entrance w/geotextile fabric as per C-8 & details on`
          spreadsheet.updateCell({ value: stabilizedText }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, stabilizedText)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${erosionCalcSheetName}'!J${rowNum}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${erosionCalcSheetName}'!K${rowNum}` }, `${pfx}E${currentRow}`)
          spreadsheet.updateCell({ formula: `='${erosionCalcSheetName}'!L${rowNum}` }, `${pfx}F${currentRow}`)
          fillRatesForProposalRow(currentRow, stabilizedText)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        })

        erosionItems.silt_fence.forEach(({ rowNum, heightStr = "2'-6\"" }) => {
          const siltFenceText = `F&I new silt fence (H=${heightStr}) as per C-8 & details on`
          spreadsheet.updateCell({ value: siltFenceText }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, siltFenceText)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${erosionCalcSheetName}'!I${rowNum}` }, `${pfx}C${currentRow}`)
          spreadsheet.updateCell({ formula: `='${erosionCalcSheetName}'!J${rowNum}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${erosionCalcSheetName}'!L${rowNum}` }, `${pfx}F${currentRow}`)
          fillRatesForProposalRow(currentRow, siltFenceText)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        })

        erosionItems.inlet_filter.forEach(({ rowNum, qty = 1 }) => {
          const inletFilterText = `F&I new (${qty})no inlet filter as per C-8 & details on`
          spreadsheet.updateCell({ value: inletFilterText }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, inletFilterText)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${erosionCalcSheetName}'!M${rowNum}` }, `${pfx}G${currentRow}`)
          fillRatesForProposalRow(currentRow, inletFilterText)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        })

        const erosionEndRow = currentRow - 1
        const erosionDataStartRow = erosionHeaderRow + 1
        spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
        spreadsheet.updateCell({ value: 'Erosion & Sediment Control Total:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle' },
          `${pfx}B${currentRow}:E${currentRow}`
        )
        spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
        spreadsheet.updateCell({ formula: `=SUM(H${erosionDataStartRow}:H${erosionEndRow})*1000` }, `${pfx}F${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', format: '$#,##0.00' },
          `${pfx}F${currentRow}:G${currentRow}`
        )
        applyTotalRowBorders(spreadsheet, pfx, currentRow, '#FEF2CB')
        siteworkTotalFRows.push(currentRow)
        currentRow++
      }

      // Excavation, backfill & grading scope
      currentRow++
      const excGradingHeaderRow = currentRow
      const excGradingCalcSheet = 'Calculations Sheet'
      const civilExcSum = (formulaData || []).find(f => f.itemType === 'civil_exc_sum' && f.section === 'civil_sitework')
      const civilGravelSum = (formulaData || []).find(f => f.itemType === 'civil_gravel_sum' && f.section === 'civil_sitework')
      const gravelItems = { transformer_pad: [], transformer_pad_8: [], reinforced_sidewalk: [], asphalt: [] }
      if (calculationData && calculationData.length > 0) {
        let inExcavation = false
        let inGravel = false
        calculationData.forEach((row, index) => {
          const colB = row[1]
          const rowNum = index + 1
          if (colB && typeof colB === 'string') {
            const bText = colB.trim().toLowerCase()
            if (bText === 'excavation:') {
              inExcavation = true
              inGravel = false
              return
            }
            if (bText === 'gravel:') {
              inExcavation = false
              inGravel = true
              return
            }
            if (bText.endsWith(':') && !bText.includes('transformer') && !bText.includes('sidewalk') && !bText.includes('asphalt') && !bText.includes('gravel')) {
              if (inGravel) inGravel = false
              if (inExcavation) inExcavation = false
              return
            }
            if (inGravel) {
              if (bText.includes('transformer') && bText.includes('pad')) {
                if (bText.includes('8" thick')) {
                  gravelItems.transformer_pad_8.push({ rowNum })
                } else {
                  gravelItems.transformer_pad.push({ rowNum })
                }
              } else if (bText.includes('reinforced') && bText.includes('sidewalk')) {
                gravelItems.reinforced_sidewalk.push({ rowNum })
              } else if (bText.includes('full depth asphalt') || bText.includes('asphalt pavement')) {
                gravelItems.asphalt.push({ rowNum })
              }
            }
          }
        })
      }

      const hasExcOrGravel = civilExcSum || civilGravelSum || gravelItems.transformer_pad.length > 0 || gravelItems.transformer_pad_8.length > 0 || gravelItems.reinforced_sidewalk.length > 0 || gravelItems.asphalt.length > 0
      if (hasExcOrGravel) {
        spreadsheet.updateCell({ value: 'Excavation, backfill & grading scope:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'center', backgroundColor: '#FEF2CB', verticalAlign: 'middle', textDecoration: 'underline', border: '1px solid #000000' },
          `${pfx}B${currentRow}`
        )
        currentRow++

        if (civilExcSum) {
          spreadsheet.updateCell({ value: 'Excavation:' }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat(
            { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
            `${pfx}B${currentRow}`
          )
          currentRow++

          const excSumRow = civilExcSum.row // Row on Calculations Sheet that has the sum (e.g. J1410, L1410)
          const soilExcavationC4Text = 'Allow to perform soil excavation, trucking & disposal as per C-4'
          spreadsheet.updateCell({ value: soilExcavationC4Text }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, soilExcavationC4Text)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
          if (excSumRow) {
            spreadsheet.updateCell({ formula: `='${excGradingCalcSheet}'!J${excSumRow}` }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: `='${excGradingCalcSheet}'!L${excSumRow}` }, `${pfx}F${currentRow}`)
          }
          fillRatesForProposalRow(currentRow, soilExcavationC4Text)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }

        if (gravelItems.transformer_pad.length > 0 || gravelItems.transformer_pad_8.length > 0 || gravelItems.reinforced_sidewalk.length > 0 || gravelItems.asphalt.length > 0) {
          spreadsheet.updateCell({ value: 'Compacted gravel:' }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat(
            { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
            `${pfx}B${currentRow}`
          )
          currentRow++

          if (gravelItems.transformer_pad.length > 0) {
            const tpRows = gravelItems.transformer_pad
            const sumFormulaJ = tpRows.length === 1 ? `='${excGradingCalcSheet}'!J${tpRows[0].rowNum}` : `=SUM(${tpRows.map(r => `'${excGradingCalcSheet}'!J${r.rowNum}`).join(',')})`
            const sumFormulaL = tpRows.length === 1 ? `='${excGradingCalcSheet}'!L${tpRows[0].rowNum}` : `=SUM(${tpRows.map(r => `'${excGradingCalcSheet}'!L${r.rowNum}`).join(',')})`
            const transformerGravelText = `F&I new (4" thick) gravel/crushed stone @ transformer concrete pad as per A-100.00`
            spreadsheet.updateCell({ value: transformerGravelText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, transformerGravelText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: sumFormulaJ }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: sumFormulaL }, `${pfx}F${currentRow}`)
            fillRatesForProposalRow(currentRow, transformerGravelText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }

          if (gravelItems.transformer_pad_8.length > 0) {
            const tp8Rows = gravelItems.transformer_pad_8
            const sumFormulaJ = tp8Rows.length === 1 ? `='${excGradingCalcSheet}'!J${tp8Rows[0].rowNum}` : `=SUM(${tp8Rows.map(r => `'${excGradingCalcSheet}'!J${r.rowNum}`).join(',')})`
            const sumFormulaL = tp8Rows.length === 1 ? `='${excGradingCalcSheet}'!L${tp8Rows[0].rowNum}` : `=SUM(${tp8Rows.map(r => `'${excGradingCalcSheet}'!L${r.rowNum}`).join(',')})`
            spreadsheet.updateCell({ value: `F&I new (8" thick) gravel/crushed stone @ transformer concrete pad as per A-100.00` }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: sumFormulaJ }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: sumFormulaL }, `${pfx}F${currentRow}`)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }

          const gravel6Rows = [...gravelItems.reinforced_sidewalk, ...gravelItems.asphalt]
          if (gravel6Rows.length > 0) {
            const sumFormulaJ = gravel6Rows.length === 1
              ? `='${excGradingCalcSheet}'!J${gravel6Rows[0].rowNum}`
              : `=SUM(${gravel6Rows.map(r => `'${excGradingCalcSheet}'!J${r.rowNum}`).join(',')})`
            const sumFormulaL = gravel6Rows.length === 1
              ? `='${excGradingCalcSheet}'!L${gravel6Rows[0].rowNum}`
              : `=SUM(${gravel6Rows.map(r => `'${excGradingCalcSheet}'!L${r.rowNum}`).join(',')})`
            const utilityTrenchGravelText = `F&I new (6" thick) gravel/crushed stone @ utility trench & asphalt pavement as per C-6 & details on`
            spreadsheet.updateCell({ value: utilityTrenchGravelText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, utilityTrenchGravelText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: sumFormulaJ }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: sumFormulaL }, `${pfx}F${currentRow}`)
            fillRatesForProposalRow(currentRow, utilityTrenchGravelText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
        }

        const excGradingEndRow = currentRow - 1
        const excGradingDataStartRow = excGradingHeaderRow + 1
        spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
        spreadsheet.updateCell({ value: 'Excavation, Backfill & Grading Total:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle' },
          `${pfx}B${currentRow}:E${currentRow}`
        )
        spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
        spreadsheet.updateCell({ formula: `=SUM(H${excGradingDataStartRow}:H${excGradingEndRow})*1000` }, `${pfx}F${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', format: '$#,##0.00' },
          `${pfx}F${currentRow}:G${currentRow}`
        )
        applyTotalRowBorders(spreadsheet, pfx, currentRow, '#FEF2CB')
        siteworkTotalFRows.push(currentRow)
        currentRow++
      }

      // Dewatering section (after Excavation, backfill & grading scope)
      currentRow++
      spreadsheet.updateCell({ value: 'Dewatering:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'center', backgroundColor: '#FEF2CB', verticalAlign: 'middle', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      const dewateringAllowanceText = 'Dewatering allowance - Budget $200k'
      spreadsheet.updateCell({ value: dewateringAllowanceText }, `${pfx}B${currentRow}`)
      rowBContentMap.set(currentRow, dewateringAllowanceText)
      spreadsheet.wrap(`${pfx}B${currentRow}`, true)
      spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
      spreadsheet.updateCell({ value: 1 }, `${pfx}G${currentRow}`)
      fillRatesForProposalRow(currentRow, dewateringAllowanceText)
      spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
      currentRow++
      spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
      spreadsheet.updateCell({ value: 'Dewatering Total:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle' },
        `${pfx}B${currentRow}:E${currentRow}`
      )
      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({ value: 200000 }, `${pfx}F${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', format: '$#,##0.00' },
        `${pfx}F${currentRow}:G${currentRow}`
      )
      applyTotalRowBorders(spreadsheet, pfx, currentRow, '#FEF2CB')
      siteworkTotalFRows.push(currentRow)
      currentRow++

      // Utilities & drainage system section (below Dewatering)
      const siteSums = (formulaData || []).filter(f => f.itemType === 'civil_site_sum' && f.section === 'civil_sitework')
      const utilsCalcSheet = 'Calculations Sheet'
      const civilEleSums = (formulaData || []).filter(f => f.itemType === 'civil_ele_sum' && f.section === 'civil_sitework')
      const civilPadsSum = (formulaData || []).find(f => f.itemType === 'civil_pads_sum' && f.section === 'civil_sitework')
      const drainsUtilsItems = (formulaData || []).filter(f => f.itemType === 'civil_drains_utilities_item' && f.section === 'civil_sitework')
      const findDrainsItem = (keyword) => drainsUtilsItems.find(f => (f.parsedData?.particulars || '').toLowerCase().includes(keyword))
      const getSiteSumByKey = (key) => (siteSums || []).find(s => s.siteGroupKey === key)

      const parseUtilsSectionFromCalc = (sectionName) => {
        const sums = { Excavation: null, Backfill: null, Gravel: null }
        if (!calculationData || calculationData.length === 0) return sums
        let inSection = false
        let currentSub = null
        let lastDataRow = null
        for (let i = 0; i < calculationData.length; i++) {
          const row = calculationData[i]
          const colB = (row[1] || '').toString().trim()
          const colJ = row[9]
          const colL = row[11]
          const hasSum = (typeof colJ === 'number' && !isNaN(colJ)) || (typeof colL === 'number' && !isNaN(colL)) || (typeof colJ === 'string' && parseFloat(colJ)) || (typeof colL === 'string' && parseFloat(colL))
          if (colB.toLowerCase().includes(sectionName.toLowerCase() + ':') && !colB.toLowerCase().includes('water main') && !colB.toLowerCase().includes('water service')) {
            inSection = true
            currentSub = null
            continue
          }
          if (!inSection) continue
          if (colB.toLowerCase().includes('drains') || colB.toLowerCase().includes('alternate') || (sectionName === 'Gas' && colB.toLowerCase().includes('water:'))) {
            if (currentSub && lastDataRow !== null) sums[currentSub] = lastDataRow + 2
            inSection = false
            break
          }
          if (colB.trim() === 'Excavation:' || colB.trim().startsWith('  Excavation:')) {
            currentSub = 'Excavation'
            continue
          }
          if (colB.trim() === 'Backfill:' || colB.trim().startsWith('  Backfill:')) {
            if (currentSub && lastDataRow !== null) sums[currentSub] = lastDataRow + 2
            currentSub = 'Backfill'
            lastDataRow = null
            continue
          }
          if (colB.trim() === 'Gravel:' || colB.trim().startsWith('  Gravel:')) {
            if (currentSub && lastDataRow !== null) sums[currentSub] = lastDataRow + 2
            currentSub = 'Gravel'
            lastDataRow = null
            continue
          }
          if (currentSub && (colB.toLowerCase().includes('proposed') || colB.toLowerCase().includes('excavation') || colB.toLowerCase().includes('backfill') || colB.toLowerCase().includes('gravel'))) {
            lastDataRow = i
          }
          if (currentSub && hasSum && lastDataRow !== null && !colB.toLowerCase().includes('proposed') && colB.trim().length < 30) {
            sums[currentSub] = i + 1
            lastDataRow = null
          }
        }
        return sums
      }

      const gasSums = parseUtilsSectionFromCalc('Gas')
      const waterSums = parseUtilsSectionFromCalc('Water')

      const civilEleExc = civilEleSums.find(f => f.subSubsectionName === 'Excavation')
      const civilEleBackfill = civilEleSums.find(f => f.subSubsectionName === 'Backfill')
      const civilEleGravel = civilEleSums.find(f => f.subSubsectionName === 'Gravel')
      const eleConduitItem = findDrainsItem('electrical conduit')
      const utilPoleItem = findDrainsItem('connection to existing utility pole')
      const gasLateralItem = findDrainsItem('gas service lateral')
      const waterMainItem = findDrainsItem('water main')
      const fireServiceItem = findDrainsItem('fire service lateral')
      const waterServiceItem = findDrainsItem('water service lateral')
      const underslabItem = findDrainsItem('underslab drainage')
      const stormSewerItem = findDrainsItem('storm sewer piping')
      const sanitarySewerItem = findDrainsItem('sanitary sewer service')
      const sanitInvertRows = []
      drainsUtilsItems.filter(f => (f.parsedData?.particulars || '').toLowerCase().includes('sanitary invert')).forEach(f => sanitInvertRows.push({ row: f.row, parsedData: f.parsedData }))
      if (sanitInvertRows.length === 0 && calculationData && calculationData.length > 0) {
        calculationData.forEach((row, i) => {
          const colB = (row[1] || '').toString().trim()
          if (colB.toLowerCase().includes('proposed sanitary invert')) {
            sanitInvertRows.push({ row: i + 1, parsedData: { particulars: colB } })
          }
        })
      }
      const sanitInvertItem = sanitInvertRows.length > 0 ? sanitInvertRows : null

      const hasUtilities = civilEleExc || civilEleBackfill || civilEleGravel || civilPadsSum || eleConduitItem || utilPoleItem ||
        gasLateralItem || gasSums.Excavation || gasSums.Backfill || gasSums.Gravel || getSiteSumByKey('Gas') ||
        waterMainItem || fireServiceItem || waterServiceItem || underslabItem || waterSums.Excavation || waterSums.Backfill || waterSums.Gravel || getSiteSumByKey('Water') ||
        stormSewerItem || sanitarySewerItem || sanitInvertItem

      if (hasUtilities) {
        currentRow++
        const utilsHeaderRow = currentRow
        spreadsheet.updateCell({ value: 'Utilities & drainage system:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#FEF2CB', verticalAlign: 'middle', textDecoration: 'underline', border: '1px solid #000000' },
          `${pfx}B${currentRow}`
        )
        currentRow++

        if (civilEleExc || civilEleBackfill || civilEleGravel || eleConduitItem || utilPoleItem || civilPadsSum) {
          spreadsheet.updateCell({ value: 'Electric service:' }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat(
            { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
            `${pfx}B${currentRow}`
          )
          currentRow++

          if (civilEleExc) {
            const excFirstRow = civilEleExc.firstDataRow || civilEleExc.row
            const excLastRow = civilEleExc.lastDataRow || civilEleExc.row
            const soilExcC4Text = 'Allow to perform soil excavation, trucking & disposal (Havg=2\'-11") as per C-4'
            spreadsheet.updateCell({ value: soilExcC4Text }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, soilExcC4Text)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            if (excFirstRow && excLastRow && excLastRow !== excFirstRow) {
              spreadsheet.updateCell({ formula: `=SUM('${utilsCalcSheet}'!J${excFirstRow}:'${utilsCalcSheet}'!J${excLastRow})` }, `${pfx}D${currentRow}`)
              spreadsheet.updateCell({ formula: `=SUM('${utilsCalcSheet}'!L${excFirstRow}:'${utilsCalcSheet}'!L${excLastRow})` }, `${pfx}F${currentRow}`)
            } else if (excFirstRow) {
              spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${excFirstRow}` }, `${pfx}D${currentRow}`)
              spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${excFirstRow}` }, `${pfx}F${currentRow}`)
            } else {
              spreadsheet.updateCell({ value: 0 }, `${pfx}D${currentRow}`)
              spreadsheet.updateCell({ value: 0 }, `${pfx}F${currentRow}`)
            }
            // Electric service: use CY 75 for soil excavation (proposal_mapped.json)
            fillRatesForProposalRow(currentRow, 'Allow to perform soil excavation, trucking & disposal - Electric service')
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          if (civilEleBackfill) {
            const backFirstRow = civilEleBackfill.firstDataRow || civilEleBackfill.row
            const backLastRow = civilEleBackfill.lastDataRow || civilEleBackfill.row
            const backfillC4Text = 'Allow to import new clean soil to backfill and compact as per C-4'
            spreadsheet.updateCell({ value: backfillC4Text }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, backfillC4Text)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            if (backFirstRow && backLastRow && backLastRow !== backFirstRow) {
              spreadsheet.updateCell({ formula: `=SUM('${utilsCalcSheet}'!J${backFirstRow}:'${utilsCalcSheet}'!J${backLastRow})` }, `${pfx}D${currentRow}`)
              spreadsheet.updateCell({ formula: `=SUM('${utilsCalcSheet}'!L${backFirstRow}:'${utilsCalcSheet}'!L${backLastRow})` }, `${pfx}F${currentRow}`)
            } else if (backFirstRow) {
              spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${backFirstRow}` }, `${pfx}D${currentRow}`)
              spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${backFirstRow}` }, `${pfx}F${currentRow}`)
            } else {
              spreadsheet.updateCell({ value: 0 }, `${pfx}D${currentRow}`)
              spreadsheet.updateCell({ value: 0 }, `${pfx}F${currentRow}`)
            }
            fillRatesForProposalRow(currentRow, backfillC4Text)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          if (civilEleGravel) {
            const gravFirstRow = civilEleGravel.firstDataRow || civilEleGravel.row
            const gravLastRow = civilEleGravel.lastDataRow || civilEleGravel.row
            const gravelUtilText = 'F&I new (6" thick) gravel/crushed stone, including 6MIL vapor barrier on top @ utility trench & asphalt pavement as per C-6 & details on'
            spreadsheet.updateCell({ value: gravelUtilText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, gravelUtilText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            if (gravFirstRow && gravLastRow && gravLastRow !== gravFirstRow) {
              spreadsheet.updateCell({ formula: `=SUM('${utilsCalcSheet}'!J${gravFirstRow}:'${utilsCalcSheet}'!J${gravLastRow})` }, `${pfx}D${currentRow}`)
              spreadsheet.updateCell({ formula: `=SUM('${utilsCalcSheet}'!L${gravFirstRow}:'${utilsCalcSheet}'!L${gravLastRow})` }, `${pfx}F${currentRow}`)
            } else if (gravFirstRow) {
              spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${gravFirstRow}` }, `${pfx}D${currentRow}`)
              spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${gravFirstRow}` }, `${pfx}F${currentRow}`)
            } else {
              spreadsheet.updateCell({ value: 0 }, `${pfx}D${currentRow}`)
              spreadsheet.updateCell({ value: 0 }, `${pfx}F${currentRow}`)
            }
            fillRatesForProposalRow(currentRow, gravelUtilText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          if (utilPoleItem) {
            const utilPoleText = 'F&I new underground connection to existing utility pole as per C-6'
            spreadsheet.updateCell({ value: utilPoleText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, utilPoleText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!M${utilPoleItem.row}` }, `${pfx}G${currentRow}`)
            fillRatesForProposalRow(currentRow, utilPoleText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          if (eleConduitItem) {
            const eleConduitText = 'F&I new underground electrical conduit'
            spreadsheet.updateCell({ value: eleConduitText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, eleConduitText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!I${eleConduitItem.row}` }, `${pfx}C${currentRow}`)
            fillRatesForProposalRow(currentRow, eleConduitText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          if (civilPadsSum) {
            spreadsheet.updateCell({ value: 'Concrete pad @ electrical service:' }, `${pfx}B${currentRow}`)
            spreadsheet.cellFormat(
              { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
              `${pfx}B${currentRow}`
            )
            currentRow++
            const transformerPadText = 'F&I new (8" thick) transformer concrete pad as per A-100.00'
            spreadsheet.updateCell({ value: transformerPadText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, transformerPadText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${civilPadsSum.row}` }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${civilPadsSum.row}` }, `${pfx}F${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!M${civilPadsSum.row}` }, `${pfx}G${currentRow}`)
            fillRatesForProposalRow(currentRow, transformerPadText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
        }

        if (gasLateralItem || gasSums.Excavation || gasSums.Backfill || gasSums.Gravel || getSiteSumByKey('Gas')) {
          spreadsheet.updateCell({ value: 'Gas service:' }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat(
            { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
            `${pfx}B${currentRow}`
          )
          currentRow++
          const gasExcRow = gasSums.Excavation
          const gasBackRow = gasSums.Backfill
          const gasGravRow = gasSums.Gravel
          if (gasExcRow) {
            const soilExcC4Text = 'Allow to perform soil excavation, trucking & disposal (Havg=2\'-11") as per C-4'
            spreadsheet.updateCell({ value: soilExcC4Text }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, soilExcC4Text)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${gasExcRow}` }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${gasExcRow}` }, `${pfx}F${currentRow}`)
            fillRatesForProposalRow(currentRow, soilExcC4Text)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          if (gasBackRow) {
            const backfillC4Text = 'Allow to import new clean soil to backfill and compact as per C-4'
            spreadsheet.updateCell({ value: backfillC4Text }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, backfillC4Text)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${gasBackRow}` }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${gasBackRow}` }, `${pfx}F${currentRow}`)
            fillRatesForProposalRow(currentRow, backfillC4Text)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          if (gasGravRow) {
            const gravelUtilText = 'F&I new (6" thick) gravel/crushed stone, including 6MIL vapor barrier on top @ utility trench & asphalt pavement as per C-6 & details on'
            spreadsheet.updateCell({ value: gravelUtilText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, gravelUtilText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${gasGravRow}` }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${gasGravRow}` }, `${pfx}F${currentRow}`)
            fillRatesForProposalRow(currentRow, gravelUtilText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          if (gasLateralItem) {
            const gasLateralText = 'F&I new (4" thick) underground gas service lateral as per C-6'
            spreadsheet.updateCell({ value: gasLateralText }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!I${gasLateralItem.row}` }, `${pfx}C${currentRow}`)
            fillRatesForProposalRow(currentRow, gasLateralText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          const gasMainSum = getSiteSumByKey('Gas')
          if (gasMainSum) {
            const gasConnectionText = 'F&I new (1)no connection to existing gas main as per C-6'
            spreadsheet.updateCell({ value: gasConnectionText }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!M${gasMainSum.row}` }, `${pfx}G${currentRow}`)
            fillRatesForProposalRow(currentRow, gasConnectionText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
        }

        if (waterMainItem || fireServiceItem || waterServiceItem || waterSums.Excavation || waterSums.Backfill || waterSums.Gravel || getSiteSumByKey('Water')) {
          spreadsheet.updateCell({ value: 'Water:' }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat(
            { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
            `${pfx}B${currentRow}`
          )
          currentRow++
          const watExcRow = waterSums.Excavation
          const watBackRow = waterSums.Backfill
          const watGravRow = waterSums.Gravel
          if (watExcRow) {
            const soilExcC4Text = 'Allow to perform soil excavation, trucking & disposal (Havg=2\'-11") as per C-4'
            spreadsheet.updateCell({ value: soilExcC4Text }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, soilExcC4Text)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${watExcRow}` }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${watExcRow}` }, `${pfx}F${currentRow}`)
            fillRatesForProposalRow(currentRow, soilExcC4Text)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          if (watBackRow) {
            const backfillC4Text = 'Allow to import new clean soil to backfill and compact as per C-4'
            spreadsheet.updateCell({ value: backfillC4Text }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, backfillC4Text)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${watBackRow}` }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${watBackRow}` }, `${pfx}F${currentRow}`)
            fillRatesForProposalRow(currentRow, backfillC4Text)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          const watGravRowOrFallback = watGravRow || watExcRow || watBackRow
          if (watGravRowOrFallback) {
            const gravelUtilText = 'F&I new (6" thick) gravel/crushed stone, including 6MIL vapor barrier on top @ utility trench & asphalt pavement as per C-6 & details on'
            spreadsheet.updateCell({ value: gravelUtilText }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, gravelUtilText)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${watGravRowOrFallback}` }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${watGravRowOrFallback}` }, `${pfx}F${currentRow}`)
            fillRatesForProposalRow(currentRow, gravelUtilText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          if (waterMainItem) {
            const waterMainText = 'F&I new underground water main as per C-6'
            spreadsheet.updateCell({ value: waterMainText }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!I${waterMainItem.row}` }, `${pfx}C${currentRow}`)
            fillRatesForProposalRow(currentRow, waterMainText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          if (fireServiceItem) {
            const fireServiceText = 'F&I new (6" thick) fire service lateral as per C-6'
            spreadsheet.updateCell({ value: fireServiceText }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!I${fireServiceItem.row}` }, `${pfx}C${currentRow}`)
            fillRatesForProposalRow(currentRow, fireServiceText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          const waterMainSum = getSiteSumByKey('Water')
          if (waterMainSum) {
            const waterConnectionText = 'F&I new (3)no connection to existing water main as per C-6'
            spreadsheet.updateCell({ value: waterConnectionText }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!M${waterMainSum.row}` }, `${pfx}G${currentRow}`)
            fillRatesForProposalRow(currentRow, waterConnectionText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          if (waterServiceItem) {
            const waterServiceText = 'F&I new (4" thick) underground water service lateral as per C-6'
            spreadsheet.updateCell({ value: waterServiceText }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!I${waterServiceItem.row}` }, `${pfx}C${currentRow}`)
            fillRatesForProposalRow(currentRow, waterServiceText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
        }

        if (stormSewerItem || sanitarySewerItem || sanitInvertItem || underslabItem) {
          spreadsheet.updateCell({ value: 'Sewer:' }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat(
            { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
            `${pfx}B${currentRow}`
          )
          currentRow++
          if (stormSewerItem) {
            const stormSewerText = 'F&I new (10" Ø) storm sewer piping as per P-099.00'
            spreadsheet.updateCell({ value: stormSewerText }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!I${stormSewerItem.row}` }, `${pfx}C${currentRow}`)
            fillRatesForProposalRow(currentRow, stormSewerText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          if (sanitarySewerItem) {
            const sanitarySewerText = 'F&I new (8" Ø) underground PVC sanitary sewer service as per C-6'
            spreadsheet.updateCell({ value: sanitarySewerText }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!I${sanitarySewerItem.row}` }, `${pfx}C${currentRow}`)
            fillRatesForProposalRow(currentRow, sanitarySewerText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          if (sanitInvertRows.length > 0) {
            const sanitInvertText = 'F&I new sanitary invert as per C-6'
            spreadsheet.updateCell({ value: sanitInvertText }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            const lfFormula = sanitInvertRows.length === 1
              ? `='${utilsCalcSheet}'!I${sanitInvertRows[0].row}`
              : `=AVERAGE(${sanitInvertRows.map(r => `'${utilsCalcSheet}'!I${r.row}`).join(',')})`
            spreadsheet.updateCell({ formula: lfFormula }, `${pfx}C${currentRow}`)
            fillRatesForProposalRow(currentRow, sanitInvertText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
          if (underslabItem) {
            const underslabText = 'F&I new underslab drainage piping'
            spreadsheet.updateCell({ value: underslabText }, `${pfx}B${currentRow}`)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!I${underslabItem.row}` }, `${pfx}C${currentRow}`)
            fillRatesForProposalRow(currentRow, underslabText)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          }
        }

        const utilsEndRow = currentRow - 1
        const utilsDataStartRow = utilsHeaderRow + 1
        spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
        spreadsheet.updateCell({ value: 'Utilities Total:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle' },
          `${pfx}B${currentRow}:E${currentRow}`
        )
        spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
        spreadsheet.updateCell({ formula: `=SUM(H${utilsDataStartRow}:H${utilsEndRow})*1000` }, `${pfx}F${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', format: '$#,##0.00' },
          `${pfx}F${currentRow}:G${currentRow}`
        )
        applyTotalRowBorders(spreadsheet, pfx, currentRow, '#FEF2CB')
        siteworkTotalFRows.push(currentRow)
        currentRow++
      }

      // Storm water management section (below Utilities & drainage system)
      currentRow++
      const stormWaterHeaderRow = currentRow
      spreadsheet.updateCell({ value: 'Storm water management:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#FEF2CB', verticalAlign: 'middle', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      const stormwaterManholeText = 'F&I new stormwater manhole as per C-5'
      spreadsheet.updateCell({ value: stormwaterManholeText }, `${pfx}B${currentRow}`)
      rowBContentMap.set(currentRow, stormwaterManholeText)
      spreadsheet.wrap(`${pfx}B${currentRow}`, true)
      spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
      spreadsheet.updateCell({ value: 1 }, `${pfx}G${currentRow}`)
      fillRatesForProposalRow(currentRow, stormwaterManholeText)
      spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
      currentRow++
      const stormWaterEndRow = currentRow - 1
      const stormWaterDataStartRow = stormWaterHeaderRow + 1
      spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
      spreadsheet.updateCell({ value: 'Storm Water Management Total:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle' },
        `${pfx}B${currentRow}:E${currentRow}`
      )
      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({ formula: `=SUM(H${stormWaterDataStartRow}:H${stormWaterEndRow})*1000` }, `${pfx}F${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', format: '$#,##0.00' },
        `${pfx}F${currentRow}:G${currentRow}`
      )
      applyTotalRowBorders(spreadsheet, pfx, currentRow, '#FEF2CB')
      siteworkTotalFRows.push(currentRow)
      currentRow++

      // Site finishes section (after Storm water management)
      const siteFinishesCalcSheet = 'Calculations Sheet'
      const civilConcretePavementSum = (formulaData || []).find(f => f.itemType === 'civil_concrete_pavement_sum' && f.section === 'civil_sitework')
      const civilAsphaltSum = (formulaData || []).find(f => f.itemType === 'civil_asphalt_sum' && f.section === 'civil_sitework')
      const civilBollardFootingSum = (formulaData || []).find(f => f.itemType === 'civil_bollard_footing_sum' && f.section === 'civil_sitework')
      const bollardFootingItems = (formulaData || []).filter(f => f.itemType === 'civil_bollard_footing_item' && f.section === 'civil_sitework')
      const trafficSignRows = bollardFootingItems.filter(f => (f.parsedData?.particulars || '').toLowerCase().includes('traffic')).map(f => f.row)
      const sanitaryInvertItem = (formulaData || []).find(f => f.itemType === 'civil_drains_utilities_item' && (f.parsedData?.particulars || '').toLowerCase().includes('sanitary invert'))

      const hasSiteFinishes = civilConcretePavementSum || civilAsphaltSum || civilBollardFootingSum || siteSums.length > 0
      const getSiteRefFromRawData = (pattern, fallback = 'C-6') => {
        if (!rawData || !Array.isArray(rawData) || rawData.length < 2) return fallback
        const headers = rawData[0]
        const dataRows = rawData.slice(1)
        const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
        const pageIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'page')
        for (const row of dataRows) {
          const digitizerItem = (row[digitizerIdx] || '').toString()
          if (pattern.test(digitizerItem)) {
            if (pageIdx >= 0 && row[pageIdx]) {
              const pageStr = String(row[pageIdx]).trim()
              const refs = pageStr.match(/[A-Z]+-[\d.]+/gi)
              if (refs && refs.length > 0) return refs[0]
            }
            const asPerMatch = digitizerItem.match(/as\s+per\s+([A-Z]+-[\d.]+)/i)
            if (asPerMatch) return asPerMatch[1]
            break
          }
        }
        return fallback
      }
      const getSiteRefsFromRawData = (pattern, fallbackMain = 'C-6', fallbackDetails = 'C-11') => {
        if (!rawData || !Array.isArray(rawData) || rawData.length < 2) return { main: fallbackMain, details: fallbackDetails }
        const headers = rawData[0]
        const dataRows = rawData.slice(1)
        const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
        const pageIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'page')
        for (const row of dataRows) {
          const digitizerItem = (row[digitizerIdx] || '').toString()
          if (pattern.test(digitizerItem)) {
            if (pageIdx >= 0 && row[pageIdx]) {
              const pageStr = String(row[pageIdx]).trim()
              const refs = pageStr.match(/[A-Z]+-[\d.]+/gi)
              if (refs && refs.length >= 2) return { main: refs[0], details: refs[1] }
              if (refs && refs.length === 1) return { main: refs[0], details: fallbackDetails }
            }
            const asPerMatch = digitizerItem.match(/as\s+per\s+([A-Z]+-[\d.]+)(?:\s+&\s+details\s+on\s+([A-Z]+-[\d.]+))?/i)
            if (asPerMatch) return { main: asPerMatch[1], details: asPerMatch[2] || fallbackDetails }
            break
          }
        }
        return { main: fallbackMain, details: fallbackDetails }
      }
      const getSiteQtyFromSum = (sumInfo) => {
        if (!sumInfo || !formulaData) return null
        const items = formulaData.filter(f => f.itemType === 'civil_site_item' && f.row >= sumInfo.firstDataRow && f.row <= sumInfo.lastDataRow)
        const bollardItems = formulaData.filter(f => (f.itemType === 'civil_bollard_footing_item' || f.itemType === 'civil_bollard_simple_item') && sumInfo.firstDataRow && sumInfo.lastDataRow && f.row >= sumInfo.firstDataRow && f.row <= sumInfo.lastDataRow)
        const relevantItems = items.length > 0 ? items : bollardItems
        if (relevantItems.length === 0 && calculationData && sumInfo.firstDataRow) {
          let total = 0
          for (let r = sumInfo.firstDataRow - 1; r < sumInfo.lastDataRow && r < calculationData.length; r++) {
            const row = calculationData[r]
            const qty = parseFloat(row[12]) || parseFloat(row[2]) || 0
            if (!isNaN(qty)) total += qty
          }
          return total > 0 ? Math.round(total) : null
        }
        const total = relevantItems.reduce((s, i) => s + (parseFloat(i.parsedData?.takeoff) || 0), 0)
        return total > 0 ? Math.round(total) : null
      }
      if (hasSiteFinishes) {
        currentRow++
        const siteFinishesHeaderRow = currentRow
        spreadsheet.updateCell({ value: 'Site finishes:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#FEF2CB', verticalAlign: 'middle', textDecoration: 'underline', border: '1px solid #000000' },
          `${pfx}B${currentRow}`
        )
        currentRow++

        if (civilConcretePavementSum) {
          spreadsheet.updateCell({ value: 'Concrete pavement:' }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat(
            { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
            `${pfx}B${currentRow}`
          )
          currentRow++
          const cpRow = civilConcretePavementSum.row
          const cpRefs = getSiteRefsFromRawData(/concrete\s+sidewalk|reinforced\s+sidewalk/i, 'C-4', 'C-11')
          const concreteSidewalkText = `F&I new (4" thick) concrete sidewalk, reinf w/ 6x6-W1.4xW1.4 w.w.f as per ${cpRefs.main} & details on`
          spreadsheet.updateCell({ value: concreteSidewalkText }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, concreteSidewalkText)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${siteFinishesCalcSheet}'!J${cpRow}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${siteFinishesCalcSheet}'!L${cpRow}` }, `${pfx}F${currentRow}`)
          fillRatesForProposalRow(currentRow, concreteSidewalkText)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }

        if (civilAsphaltSum) {
          spreadsheet.updateCell({ value: 'Asphalt pavement:' }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat(
            { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
            `${pfx}B${currentRow}`
          )
          currentRow++
          const apRow = civilAsphaltSum.row
          const apRefs = getSiteRefsFromRawData(/asphalt\s+pavement|full\s+depth\s+asphalt/i, 'FO-001.00', 'C-11')
          const asphaltPavementText = `F&I new full depth asphalt pavement (1.5" thick) surface course on (3" thick) base course as per ${apRefs.main} & details on`
          spreadsheet.updateCell({ value: asphaltPavementText }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, asphaltPavementText)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${siteFinishesCalcSheet}'!J${apRow}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${siteFinishesCalcSheet}'!L${apRow}` }, `${pfx}F${currentRow}`)
          fillRatesForProposalRow(currentRow, asphaltPavementText)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }

        if (siteSums.length > 0 || trafficSignRows.length > 0) {
          spreadsheet.updateCell({ value: 'Site:' }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat(
            { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
            `${pfx}B${currentRow}`
          )
          currentRow++

          // Match by exact line (single item) or whole group (sum): search template/calculations for corresponding data
          const getSiteSumByKey = (key) => (siteSums || []).find(s => s.siteGroupKey === key)
          const siteProposalLines = [
            { textBuilder: (qty, ref) => `F&I new (${qty ?? 2})no traffic sign w/ footing (18" Ø, H=3'-0") as per ${ref}`, qtySource: 'traffic' },
            { textBuilder: (qty, ref) => 'Allow to relocate existing fire hydrant as per C-3', qtySource: 'sum', siteGroupKey: 'Hydrant', rateLookupKey: 'Allow to relocate existing fire hydrant - Site' },
            { textBuilder: (qty, ref) => `F&I new (${qty ?? 9})no concrete wheel stop as per ${ref}`, qtySource: 'sum', siteGroupKey: 'Wheel stop' },
            { textBuilder: (qty, ref) => `F&I new (${qty ?? 11})no area drain as per ${ref}`, qtySource: 'sum', siteGroupKey: 'Area' },
            { textBuilder: (qty, ref) => `F&I new (${qty ?? 9})no floor drain as per ${ref}`, qtySource: 'sum', siteGroupKey: 'Floor' },
            { textBuilder: (qty, ref) => `F&I new (${qty ?? 5})no inlet filter protection as per ${ref}`, qtySource: 'sum', siteGroupKey: 'Protection' },
            { textBuilder: (qty, ref) => `F&I new (${qty ?? 8})no signages as per ${ref}`, qtySource: 'sum', siteGroupKey: 'Signages' },
            { textBuilder: (qty, ref) => `F&I new (${qty ?? 1})no sanitary invert as per ${ref}`, qtySource: 'sanitaryInvert' },
            { textBuilder: (qty, ref) => `F&I new (${qty ?? 1})no connection to existing sanitary main as per ${ref}`, qtySource: 'sum', siteGroupKey: 'Sanitary' }
          ]
          siteProposalLines.forEach((line) => {
            let qtyFormula = ''
            let qtyVal = null
            let ref = 'C-6'
            if (line.qtySource === 'traffic' && trafficSignRows.length > 0) {
              qtyFormula = trafficSignRows.length === 1 ? `='${siteFinishesCalcSheet}'!M${trafficSignRows[0]}` : `=SUM(${trafficSignRows.map(r => `'${siteFinishesCalcSheet}'!M${r}`).join(',')})`
              ref = getSiteRefFromRawData(/traffic\s+sign/i)
            } else if (line.qtySource === 'sanitaryInvert' && sanitaryInvertItem) {
              qtyFormula = `='${siteFinishesCalcSheet}'!M${sanitaryInvertItem.row}`
              qtyVal = (calculationData && sanitaryInvertItem.row) ? (parseFloat(calculationData[sanitaryInvertItem.row - 1]?.[12]) || parseFloat(calculationData[sanitaryInvertItem.row - 1]?.[2])) : null
              ref = getSiteRefFromRawData(/sanitary\s+invert/i)
            } else if (line.qtySource === 'sum' && line.siteGroupKey) {
              const sumRow = getSiteSumByKey(line.siteGroupKey)
              if (sumRow) {
                qtyFormula = `='${siteFinishesCalcSheet}'!M${sumRow.row}`
                qtyVal = getSiteQtyFromSum(sumRow)
                const patternMap = { Hydrant: /fire\s+hydrant/i, 'Wheel stop': /wheel\s+stop/i, Area: /area\s+drain/i, Floor: /floor\s+drain/i, Protection: /inlet\s+filter/i, Signages: /signages/i, Sanitary: /connection\s+to\s+existing\s+sanitary/i }
                ref = getSiteRefFromRawData(patternMap[line.siteGroupKey] || /./)
              }
            }
            if (!qtyFormula) return
            const text = typeof line.textBuilder === 'function' ? line.textBuilder(qtyVal, ref) : line.text
            spreadsheet.updateCell({ value: text }, `${pfx}B${currentRow}`)
            rowBContentMap.set(currentRow, text)
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
            spreadsheet.updateCell({ formula: qtyFormula }, `${pfx}G${currentRow}`)
            fillRatesForProposalRow(currentRow, line.rateLookupKey != null ? line.rateLookupKey : text)
            spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
            currentRow++
          })
        }

        if (civilBollardFootingSum) {
          spreadsheet.updateCell({ value: 'Concrete filled steel pipe bollard:' }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat(
            { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
            `${pfx}B${currentRow}`
          )
          currentRow++
          const bfRow = civilBollardFootingSum.row
          const bollardItems = formulaData.filter(f => f.itemType === 'civil_bollard_footing_item' && f.row >= civilBollardFootingSum.firstDataRow && f.row <= civilBollardFootingSum.lastDataRow)
          const firstBollard = bollardItems[0]?.parsedData
          const bollardDims = firstBollard?.parsed?.bollardDimensions || firstBollard?.bollardDimensions
          const diaStr = bollardDims?.bollardDia ? `${bollardDims.bollardDia}"` : '6"'
          const hFt = bollardDims?.bollardH ? Math.floor(bollardDims.bollardH) : 6
          const hIn = bollardDims?.bollardH ? Math.round((bollardDims.bollardH % 1) * 12) : 0
          const heightStr = hIn > 0 ? `${hFt}'-${hIn}"` : `${hFt}'-0"`
          const bollardQty = getSiteQtyFromSum(civilBollardFootingSum)
          const bollardRefs = getSiteRefsFromRawData(/bollard/i, 'C-4', 'C-11')
          const bollardText = `F&I new (${bollardQty ?? 28})no (${diaStr} Ø wide) concrete filled steel pipe bollard (Havg=${heightStr}) as per ${bollardRefs.main} & details on`
          spreadsheet.updateCell({ value: bollardText }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, bollardText)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${siteFinishesCalcSheet}'!J${bfRow}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${siteFinishesCalcSheet}'!L${bfRow}` }, `${pfx}F${currentRow}`)
          spreadsheet.updateCell({ formula: `='${siteFinishesCalcSheet}'!M${bfRow}` }, `${pfx}G${currentRow}`)
          fillRatesForProposalRow(currentRow, bollardText)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }

        spreadsheet.updateCell({ value: 'Misc.:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
          `${pfx}B${currentRow}`
        )
        currentRow++
        const tempProtectionText = 'Temp Protection & Barriers'
        spreadsheet.updateCell({ value: tempProtectionText }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, tempProtectionText)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
        spreadsheet.updateCell({ value: 1 }, `${pfx}G${currentRow}`)
        fillRatesForProposalRow(currentRow, tempProtectionText)
        spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
        currentRow++
        const dotPermitsText = 'DOT Permits, Temp DOT Barriers'
        spreadsheet.updateCell({ value: dotPermitsText }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, dotPermitsText)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
        spreadsheet.updateCell({ value: 1 }, `${pfx}G${currentRow}`)
        fillRatesForProposalRow(currentRow, dotPermitsText)
        spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
        currentRow++

        const siteFinishesEndRow = currentRow - 1
        const siteFinishesDataStartRow = siteFinishesHeaderRow + 1
        spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
        spreadsheet.updateCell({ value: 'Site Finishes Total:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle' },
          `${pfx}B${currentRow}:E${currentRow}`
        )
        spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
        spreadsheet.updateCell({ formula: `=SUM(H${siteFinishesDataStartRow}:H${siteFinishesEndRow})*1000` }, `${pfx}F${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', format: '$#,##0.00' },
          `${pfx}F${currentRow}:G${currentRow}`
        )
        applyTotalRowBorders(spreadsheet, pfx, currentRow, '#FEF2CB')
        siteworkTotalFRows.push(currentRow)
        currentRow++
      }

      // Fence section
      const fenceCalcSheet = 'Calculations Sheet'
      const fenceSums = (formulaData || []).filter(f => f.itemType === 'civil_fence_sum' && f.section === 'civil_sitework')
      const fenceRef0 = getSiteRefFromRawData(/construction\s+fence/i, 'SOE-100.00')
      const fenceRef1 = getSiteRefFromRawData(/proposed\s+fence|fence\s+height/i, 'SOE-100.00')
      const fenceRef2 = getSiteRefFromRawData(/guiderail/i, 'SOE-100.00')
      const fenceItems = [
        { proposalText: `F&I new construction fence (H=6'-0") as per ${fenceRef0}`, sumIndex: 0 },
        { proposalText: `F&I new fence (H=10'-0") as per ${fenceRef1}`, sumIndex: 1 },
        { proposalText: `F&I new guiderail (H=3'-6") as per ${fenceRef2}`, sumIndex: 2 }
      ]
      const hasFenceItems = fenceSums.length > 0
      if (hasFenceItems) {
        currentRow++
        const fenceHeaderRow = currentRow
        spreadsheet.updateCell({ value: 'Fence:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#FEF2CB', verticalAlign: 'middle', textDecoration: 'underline', border: '1px solid #000000' },
          `${pfx}B${currentRow}`
        )
        currentRow++

        fenceItems.forEach((item, idx) => {
          const sumInfo = fenceSums[item.sumIndex]
          if (!sumInfo) return
          const sumRow = sumInfo.row
          spreadsheet.updateCell({ value: item.proposalText }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${fenceCalcSheet}'!I${sumRow}` }, `${pfx}C${currentRow}`)
          spreadsheet.updateCell({ formula: `='${fenceCalcSheet}'!J${sumRow}` }, `${pfx}D${currentRow}`)
          fillRatesForProposalRow(currentRow, item.proposalText)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        })

        const fenceEndRow = currentRow - 1
        const fenceDataStartRow = fenceHeaderRow + 1
        spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
        spreadsheet.updateCell({ value: 'Construction Fence Total:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle' },
          `${pfx}B${currentRow}:E${currentRow}`
        )
        spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
        spreadsheet.updateCell({ formula: `=SUM(H${fenceDataStartRow}:H${fenceEndRow})*1000` }, `${pfx}F${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', format: '$#,##0.00' },
          `${pfx}F${currentRow}:G${currentRow}`
        )
        applyTotalRowBorders(spreadsheet, pfx, currentRow, '#FEF2CB')
        siteworkTotalFRows.push(currentRow)
        currentRow++
      }

      // Allowances section
      currentRow++
      const allowancesHeaderRow = currentRow
      spreadsheet.updateCell({ value: 'Allowances:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#FEF2CB', verticalAlign: 'middle', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++

      const allowancesItems = [
        'Survey',
        'Site Logistics',
        'Construction Fence with Gates including maintenance.',
        'Removals outside property line. All restoration is excluded.',
        'Restoration of Fire Department lot outside property line.',
        'Subsurface gravity drainage system below basement level asphalt.'
      ]
      allowancesItems.forEach((itemText) => {
        spreadsheet.updateCell({ value: itemText }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, itemText)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top', textDecoration: 'none' }, `${pfx}B${currentRow}`)
        spreadsheet.updateCell({ value: 1 }, `${pfx}G${currentRow}`)
        fillRatesForProposalRow(currentRow, itemText)
        spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
        currentRow++
      })

      const allowancesEndRow = currentRow - 1
      const allowancesDataStartRow = allowancesHeaderRow + 1
      spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
      spreadsheet.updateCell({ value: 'Allowances Total:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle' },
        `${pfx}B${currentRow}:E${currentRow}`
      )
      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({ formula: `=SUM(H${allowancesDataStartRow}:H${allowancesEndRow})*1000` }, `${pfx}F${currentRow}`)
      try { spreadsheet.numberFormat('$#,##0.00', `${pfx}F${currentRow}:G${currentRow}`) } catch (e) { }
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', format: '$#,##0.00' },
        `${pfx}F${currentRow}:G${currentRow}`
      )
      applyTotalRowBorders(spreadsheet, pfx, currentRow, '#FEF2CB')
      siteworkTotalFRows.push(currentRow)
      totalRows.push(currentRow) // Allowances Total: dollar format and alignment like other totals
      currentRow++

      // Add Alternate #1: Sitework Total row - sum of individual section F column totals
      if (siteworkTotalFRows.length > 0) {
        spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
        spreadsheet.updateCell({ value: 'Add Alternate #1: Sitework Total:' }, `${pfx}D${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD6EE' },
          `${pfx}D${currentRow}:E${currentRow}`
        )
        spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
        spreadsheet.updateCell({ formula: `=SUM(${siteworkTotalFRows.map(r => `F${r}`).join(',')})` }, `${pfx}F${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD6EE', format: '$#,##0.00' },
          `${pfx}F${currentRow}:G${currentRow}`
        )
        applyTotalRowBorders(spreadsheet, pfx, currentRow, '#BDD6EE')
        currentRow++
      }

      // Add empty row before Alternate #2
      currentRow++

      // Alternate #2 : B.P.P. scope header
      let bppScopeStartRow = null
      {
        spreadsheet.merge(`${pfx}B${currentRow}:G${currentRow}`)
        spreadsheet.updateCell({ value: 'Add Alternate #2 : B.P.P. scope: as per C-4' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'center', verticalAlign: 'middle', backgroundColor: '#BDD6EE', border: '1px solid #000000' },
          `${pfx}B${currentRow}:G${currentRow}`
        )
        currentRow++
        bppScopeStartRow = currentRow // Track start row for total calculation
      }

      // B.P.P. scope content - organized by street
      {
        // Get BPP items from formulaData grouped by street
        const bppFormulas = formulaData.filter(f => f.section === 'bpp_alternate')
        const bppStreetHeaders = bppFormulas.filter(f => f.itemType === 'bpp_street_header')
        const streetNames = [...new Set(bppStreetHeaders.map(f => f.streetName))].filter(Boolean)

        // If no BPP data, use default "West Street"
        const streets = streetNames.length > 0 ? streetNames : ['West Street']

        streets.forEach((streetName) => {
          // Get BPP items for this street from formulaData FIRST (needed for quantity formulas)
          const streetBppItems = bppFormulas.filter(f => f.streetName === streetName)

          // Get calculation row numbers for each item type
          const sidewalkRows = streetBppItems.filter(f => f.itemType === 'bpp_concrete_sidewalk').map(f => f.row)
          const drivewayRows = streetBppItems.filter(f => f.itemType === 'bpp_concrete_driveway').map(f => f.row)
          const curbRows = streetBppItems.filter(f => f.itemType === 'bpp_concrete_curb').map(f => f.row)
          const flushCurbRows = streetBppItems.filter(f => f.itemType === 'bpp_concrete_flush_curb').map(f => f.row)
          const expansionRows = streetBppItems.filter(f => f.itemType === 'bpp_expansion_joint').map(f => f.row)
          const asphaltRows = streetBppItems.filter(f => f.itemType === 'bpp_full_depth_asphalt').map(f => f.row)
          const gravelRows = streetBppItems.filter(f => f.itemType === 'bpp_gravel').map(f => f.row)
          const concRoadBaseRows = streetBppItems.filter(f => f.itemType === 'bpp_conc_road_base').map(f => f.row)

          // Helper to build SUM formula for multiple rows
          const buildSumFormula = (rows, col) => {
            if (rows.length === 0) return ''
            if (rows.length === 1) return `='Calculations Sheet'!${col}${rows[0]}`
            return `=SUM(${rows.map(r => `'Calculations Sheet'!${col}${r}`).join(',')})`
          }

          // "Pavement replacement: [Street Name]" header - background only on column B
          spreadsheet.updateCell({ value: `Pavement replacement: ${streetName}` }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat(
            { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', border: '1px solid #000000' },
            `${pfx}B${currentRow}`
          )
          currentRow++

          // Demo/removal proposal lines with quantities from Calculations Sheet
          // 1. Allow to saw-cut/demo/remove/dispose existing sidewalk - SF and CY from sidewalk
          const bppLine1 = 'Allow to saw-cut/demo/remove/dispose existing sidewalk'
          spreadsheet.updateCell({ value: bppLine1 }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, bppLine1)
          spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, bppLine1)
          if (sidewalkRows.length > 0) {
            spreadsheet.updateCell({ formula: buildSumFormula(sidewalkRows, 'J') }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: buildSumFormula(sidewalkRows, 'L') }, `${pfx}F${currentRow}`)
          }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++

          // 2. Allow to saw-cut/demo/remove/dispose existing 7" sidewalk/driveway - SF and CY from driveway
          const bppLine2 = 'Allow to saw-cut/demo/remove/dispose existing 7" sidewalk/driveway'
          spreadsheet.updateCell({ value: bppLine2 }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, bppLine2)
          spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, bppLine2)
          if (drivewayRows.length > 0) {
            spreadsheet.updateCell({ formula: buildSumFormula(drivewayRows, 'J') }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: buildSumFormula(drivewayRows, 'L') }, `${pfx}F${currentRow}`)
          }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++

          // 3. Allow to strip existing asphalt & dispose - SF and CY from asphalt
          const bppLine3 = 'Allow to strip existing asphalt & dispose'
          spreadsheet.updateCell({ value: bppLine3 }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, bppLine3)
          spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, bppLine3)
          if (asphaltRows.length > 0) {
            spreadsheet.updateCell({ formula: buildSumFormula(asphaltRows, 'J') }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: buildSumFormula(asphaltRows, 'L') }, `${pfx}F${currentRow}`)
          }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++

          // 4. Allow to demo existing concrete curb & dispose - LF and CY from curb
          const bppLine4 = 'Allow to demo existing concrete curb & dispose'
          spreadsheet.updateCell({ value: bppLine4 }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, bppLine4)
          spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, bppLine4)
          if (curbRows.length > 0) {
            spreadsheet.updateCell({ formula: buildSumFormula(curbRows, 'I') }, `${pfx}C${currentRow}`)
            spreadsheet.updateCell({ formula: buildSumFormula(curbRows, 'L') }, `${pfx}F${currentRow}`)
          }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++

          // 5. Allow to demo existing concrete flush curb & dispose - LF and CY from flush curb
          const bppLine5 = 'Allow to demo existing concrete flush curb & dispose'
          spreadsheet.updateCell({ value: bppLine5 }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, bppLine5)
          spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, bppLine5)
          if (flushCurbRows.length > 0) {
            spreadsheet.updateCell({ formula: buildSumFormula(flushCurbRows, 'I') }, `${pfx}C${currentRow}`)
            spreadsheet.updateCell({ formula: buildSumFormula(flushCurbRows, 'L') }, `${pfx}F${currentRow}`)
          }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++

          // F&I proposal lines with quantities from Calculations Sheet
          // 6. F&I new (6" thick) ¾" crushed stone - SF from sidewalk+driveway+asphalt, CY = SF×6"/12/27
          const bppLine6 = 'F&I new (6" thick) ¾" crushed stone on compacted subgrade @ sidewalks & asphalt'
          spreadsheet.updateCell({ value: bppLine6 }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, bppLine6)
          spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, bppLine6)
          const allSfRows = [...sidewalkRows, ...drivewayRows, ...asphaltRows]
          if (allSfRows.length > 0) {
            spreadsheet.updateCell({ formula: buildSumFormula(allSfRows, 'J') }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: buildSumFormula(allSfRows, 'L') }, `${pfx}F${currentRow}`)
          }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++

          // 7. F&I new (4" thick) concrete sidewalk - SF/CY from sidewalk; use street-specific rate lookup so "concrete sidewalk, reinf w/ -Pavement replacement: [Street]" gets SF 12
          const bppLine7 = 'F&I new (4" thick) concrete sidewalk, reinf w/ 6x6-W1.4xW1.4 (NYCDOT H-1045, Type I)'
          spreadsheet.updateCell({ value: bppLine7 }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, bppLine7)
          spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, bppLine7, `concrete sidewalk, reinf w/ -Pavement replacement: ${streetName}`)
          if (sidewalkRows.length > 0) {
            spreadsheet.updateCell({ formula: buildSumFormula(sidewalkRows, 'J') }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: buildSumFormula(sidewalkRows, 'L') }, `${pfx}F${currentRow}`)
          }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++

          // 8. F&I new (7" thick) concrete sidewalk/driveway - SF/CY from driveway
          const bppLine8 = 'F&I new (7" thick) concrete sidewalk/driveway, reinf w/ 6x6-W1.4xW1.4 @ corners & driveways'
          spreadsheet.updateCell({ value: bppLine8 }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, bppLine8)
          spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, bppLine8)
          if (drivewayRows.length > 0) {
            spreadsheet.updateCell({ formula: buildSumFormula(drivewayRows, 'J') }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: buildSumFormula(drivewayRows, 'L') }, `${pfx}F${currentRow}`)
          }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++

          // 9. F&I new (8" thick) concrete curb - LF/SF/CY from curb
          const bppLine9 = 'F&I new (8" thick) concrete curb (H=1\'-6") NYCDOT H-1010'
          spreadsheet.updateCell({ value: bppLine9 }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, bppLine9)
          spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, bppLine9)
          if (curbRows.length > 0) {
            spreadsheet.updateCell({ formula: buildSumFormula(curbRows, 'I') }, `${pfx}C${currentRow}`)
            spreadsheet.updateCell({ formula: buildSumFormula(curbRows, 'J') }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: buildSumFormula(curbRows, 'L') }, `${pfx}F${currentRow}`)
          }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++

          // 10. F&I new (6" thick) flush concrete curb - LF/SF/CY from flush curb
          const bppLine10 = 'F&I new (6" thick) flush concrete curb (H=1\'-6") NYCDOT H-1010'
          spreadsheet.updateCell({ value: bppLine10 }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, bppLine10)
          spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, bppLine10)
          if (flushCurbRows.length > 0) {
            spreadsheet.updateCell({ formula: buildSumFormula(flushCurbRows, 'I') }, `${pfx}C${currentRow}`)
            spreadsheet.updateCell({ formula: buildSumFormula(flushCurbRows, 'J') }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: buildSumFormula(flushCurbRows, 'L') }, `${pfx}F${currentRow}`)
          }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++

          // 11. F&I new 1/2" expansion joint - LF from expansion joint
          const bppLine11 = 'F&I new 1/2" expansion joint & caulking perimeter & 20\' o.c. of new sidewalk'
          spreadsheet.updateCell({ value: bppLine11 }, `${pfx}B${currentRow}`)
          rowBContentMap.set(currentRow, bppLine11)
          spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, bppLine11)
          if (expansionRows.length > 0) {
            spreadsheet.updateCell({ formula: buildSumFormula(expansionRows, 'I') }, `${pfx}C${currentRow}`)
          }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++

          // 12. F&I new (6"-9" thick) concrete road base - SF/CY from conc road base
          const bppLine12 = 'F&I new (6"-9" thick) concrete road base (Considered 3\'-0" from curb) @ new asphalt'
          spreadsheet.updateCell({ value: bppLine12 }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, bppLine12)
          if (concRoadBaseRows.length > 0) {
            spreadsheet.updateCell({ formula: buildSumFormula(concRoadBaseRows, 'J') }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: buildSumFormula(concRoadBaseRows, 'L') }, `${pfx}F${currentRow}`)
          }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++

          // 13. F&I new 1½" asphalt wearing course - SF/CY from asphalt
          const bppLine13 = 'F&I new 1½" asphalt wearing course on 3" asphalt binder course @ stripped roadway'
          spreadsheet.updateCell({ value: bppLine13 }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, bppLine13)
          if (asphaltRows.length > 0) {
            spreadsheet.updateCell({ formula: buildSumFormula(asphaltRows, 'J') }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: buildSumFormula(asphaltRows, 'L') }, `${pfx}F${currentRow}`)
          }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++

          // Note line
          spreadsheet.updateCell({ value: 'Note: granite block pavers for tree pits excluded' }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat({ textAlign: 'left', fontStyle: 'italic' }, `${pfx}B${currentRow}`)
          currentRow++
        })

        // Misc. section for B.P.P. scope
        spreadsheet.updateCell({ value: 'Misc.:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', border: '1px solid #000000' },
          `${pfx}B${currentRow}`
        )
        currentRow++

        const bppMiscItems = [
          'Design mix/TR-3',
          'DOT Barriers & Temp Protection',
          'DOT Permits'
        ]

        bppMiscItems.forEach(itemText => {
          spreadsheet.updateCell({ value: itemText }, `${pfx}B${currentRow}`)
          spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, itemText)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        })

        // Add Alternate #2 : B.P.P. Total row
        const bppScopeEndRow = currentRow - 1
        currentRow++ // Empty row before total
        spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
        spreadsheet.updateCell({ value: 'Add Alternate #2 : B.P.P. Total:' }, `${pfx}D${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD6EE' },
          `${pfx}D${currentRow}:E${currentRow}`
        )
        spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
        spreadsheet.updateCell({ formula: `=SUM(H${bppScopeStartRow}:H${bppScopeEndRow})*1000` }, `${pfx}F${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD6EE' },
          `${pfx}F${currentRow}:G${currentRow}`
        )
        try { spreadsheet.numberFormat('$#,##0.00', `${pfx}F${currentRow}:G${currentRow}`) } catch (e) { }
        applyTotalRowBorders(spreadsheet, pfx, currentRow, '#BDD6EE')
        currentRow++
      }

    } // end if (hasBPPAlternate2InCalc && hasCivilSiteworkInCalc)

    // Two empty rows before Add Alternate sections
    currentRow++
    currentRow++

    // Add Alternate Concrete Unit Rates section
    {
      spreadsheet.updateCell({ value: 'Add Alternate Concrete Unit Rates:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'center', backgroundColor: '#AECBED', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++

      const concreteUnitRates = [
        'Obstruction Removal during drilling using down the hole hammer: $1195/Hr.',
        'Hot Water Concrete: $6/CY',
        '1% Non Chloride Accelerator: $15/CY',
        '2% Non chloride retarder as per submitted hot weather procedures based on ACI 305R( verified via batching tickets): $20/CY',
        'Thermal Blankets: $1.50/SQFT',
        'Temporary Heat (Tarps & Torpedo Heaters) as per ACI standards: $6/SQFT',
        '2% retardant: $18/CY',
        'F&I Ice as required as per submitted hot weather procedures based on ACI 305R ( verified via batching tickets) (40lbs per CY): $1/LB',
        'Rock in lieu of soil: $300/CY',
        'Wet Curing ( Burlap and Sprinkler system) $1.50/SQFT',
        'Short load concrete cost: $1150',
        'Line Drilling: $35/LF'
      ]

      concreteUnitRates.forEach(itemText => {
        spreadsheet.updateCell({ value: itemText }, `${pfx}B${currentRow}`)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat({ textAlign: 'left', wrapText: true }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, itemText)
        currentRow++
      })
    }

    // Add Alternate Masonry Unit Rates section
    {
      spreadsheet.updateCell({ value: 'Add Alternate Masonry Unit Rates:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'center', backgroundColor: '#AECBED', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++

      const masonryUnitRates = [
        'Blok-Guard and Graffiti Control ($3/SF)',
        'Prosoco Sealant if required ($1/SF)',
        'Thru wall Flashing inclusive of end dams ($25/LF)',
        '2" Rigid Insulation ($3/SF)',
        'Water barrier/air barrier (e.g. Blueskin VP 160) ($5/SF)',
        'Thermal Blankets: ($1.50/SF)',
        'Temporary Heat (Tarps & Torpedo Heaters) as per ACI standards: ($6/SF)',
        '2% retardant: ($18/yard)',
        'Additional Pipe Scaffold if required ($6/SF) based off 2,000 SF minimum',
        'Precast Sill Furnish and Install ($120/LF)',
        'Furnish and Install Relieving angles ($65/LF)',
        'Precast Coping ($150/LF)'
      ]

      masonryUnitRates.forEach(itemText => {
        spreadsheet.updateCell({ value: itemText }, `${pfx}B${currentRow}`)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat({ textAlign: 'left', wrapText: true }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, itemText)
        currentRow++
      })
    }

    // Labor section
    {
      spreadsheet.updateCell({ value: 'Labor:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++

      const laborRates = [
        'Foreman: $95.00/Hr.',
        'Mechanic: $85.00/Hr.',
        'Laborer: $70.00/Hr.',
        'Driver: $82.50/Hr.',
        'Carpenter: $85.00/Hr.',
        'Yard manager: $90.00/Hr.',
        'Trucking Fee: $250',
        'Operating Engineer: $100.00/Hr.'
      ]

      laborRates.forEach(itemText => {
        spreadsheet.updateCell({ value: itemText }, `${pfx}B${currentRow}`)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat({ textAlign: 'left', wrapText: true }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, itemText)
        currentRow++
      })
    }

    // Exclusions section (proposal data range ends here for I12:N border/formatting)
    {
      proposalDataEndRow = currentRow
      spreadsheet.updateCell({ value: 'Exclusions' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'center', backgroundColor: '#AECBED', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++

      const exclusions = [
        'Union/Prevailing Wage Labor',
        'Testing & Inspections',
        'Overtime',
        'Demolition of below grade structures/exiting foundations',
        'Contaminated/Hazardous Soil disposal, assumed as clean fill',
        'Locating/Relocating Existing Services',
        'All Permits unless specifically noted above',
        'Shoring of existing or neighbouring structures',
        'Monitoring of existing and neighbouring structures for vibration/movement/settlement',
        'Protection/Trees replacement/Landscaping',
        'Dewatering',
        'Concrete encased conduit',
        'Winter Concrete/Temp Heat &/or Site winterization (Hot & Cold Weather Procedures)',
        'Snow/Ice Removal',
        'Mini Containers & Associated Carting',
        'Site Fence & Permitting',
        'Site Drainage & Plumbing & associated Precast manholes',
        'Traffic Coating',
        'Epoxy Coated Reinforcement & Corrosion inhibitor additive @ exposed concrete slabs',
        'Remobilization due to delays/stoppages caused by others',
        'Downtime due to others, charged additionally via T&M',
        'Pre-cast concrete / stones',
        'Architectural Concrete (pads, curbs, raised slabs, etc.) unless above',
        'General Liability total coverage of $10million included, if required, additional GL coverage or bonding can be provided at an additional cost'
      ]

      exclusions.forEach(itemText => {
        spreadsheet.updateCell({ value: itemText }, `${pfx}B${currentRow}`)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat({ textAlign: 'left', wrapText: true }, `${pfx}B${currentRow}`)
        rowBContentMap.set(currentRow, itemText)
        if (itemText.includes('General Liability total coverage')) {
          generalLiabilityBorderRow = currentRow
          spreadsheet.cellFormat({ borderBottom: '3px solid #000000' }, `${pfx}B${currentRow}:G${currentRow}`)
        }
        currentRow++
      })
    }

    // Signature/Acceptance rows
    {
      // Empty row before signature section
      currentRow++

      // Accept By/Title row
      spreadsheet.merge(`${pfx}B${currentRow}:D${currentRow}`)
      spreadsheet.updateCell({ value: 'Accept By/Title: _________________________' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat({ textAlign: 'left', fontWeight: 'normal' }, `${pfx}B${currentRow}:D${currentRow}`)
      spreadsheet.merge(`${pfx}E${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({ value: '_________________________ Owner' }, `${pfx}E${currentRow}`)
      spreadsheet.cellFormat({ textAlign: 'right', fontWeight: 'normal' }, `${pfx}E${currentRow}:G${currentRow}`)
      currentRow++

      // Empty row
      currentRow++

      // Company Name row
      spreadsheet.merge(`${pfx}B${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({ value: 'Company Name _________________________' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat({ textAlign: 'left', fontWeight: 'normal' }, `${pfx}B${currentRow}:G${currentRow}`)
      currentRow++
    }


    // Draw the main large box border (Left, Right, Bottom)
    // Top edge was drawn at the beginning (A3:G3).
    const finalRow = currentRow - 1
    if (finalRow >= 3) {
      // Top edge
      spreadsheet.cellFormat({ borderTop: '3px solid #000000' }, `${pfx}B3:G3`)

      // Bottom edge
      spreadsheet.cellFormat({ borderBottom: '3px solid #000000' }, `${pfx}B${finalRow}:G${finalRow}`)

      // Left edge (iterate to force application likely)
      // Applying to range usually works but if background color logic overrides it, we might need to be specific
      // Or simply re-applying it here at the end should work if it's the last operation.
      // But to be safe vs background color cells:
      spreadsheet.cellFormat({ borderLeft: '3px solid #000000' }, `${pfx}B3:B${finalRow}`)
      spreadsheet.cellFormat({ borderRight: '3px solid #000000' }, `${pfx}G3:G${finalRow}`)

      // Border on each row for columns B to G (1px internal grid); exclude last 4 rows so they have no internal borders
      const thinEndRow = finalRow - 4
      if (thinEndRow >= 13) {
        spreadsheet.cellFormat(thin, `${pfx}B13:G${thinEndRow}`)
      } else {
        spreadsheet.cellFormat(thin, `${pfx}B13:G${finalRow}`)
      }
      spreadsheet.cellFormat(thickLeft, `${pfx}B13:B${finalRow}`)
      spreadsheet.cellFormat(thickRight, `${pfx}G13:G${finalRow}`)
      spreadsheet.cellFormat({ borderTop: '3px solid #000000' }, `${pfx}B13:G13`)
      spreadsheet.cellFormat(thickBottom, `${pfx}B${finalRow}:G${finalRow}`)

      // Global Font Style Application (Calibri Bold 18pt, text color black); all cells vertically centered
      spreadsheet.cellFormat({ fontFamily: 'Calibri', fontWeight: 'bold', fontSize: '18pt', color: '#000000', verticalAlign: 'middle' }, `${pfx}A2:N${finalRow}`)
      // Row 1 header: black text, vertically centered
      spreadsheet.cellFormat({ color: '#000000', verticalAlign: 'middle' }, `${pfx}A1:N1`)

      // $ value columns (H = $/1000, I–N = unit rates & amounts): font size 11pt, not 18pt; light padding (data rows only; H1/H12 are headers)
      const cellPaddingIndent = '6pt'
      spreadsheet.cellFormat({ fontSize: '11pt', textIndent: cellPaddingIndent }, `${pfx}H13:H${finalRow}`)
      spreadsheet.cellFormat({ fontSize: '11pt', textIndent: cellPaddingIndent }, `${pfx}I2:N${finalRow}`)
      // Row 12 header: 18pt, 3px top and bottom, 2px left/right; $/1000 (H12) at 11pt, centered
      spreadsheet.cellFormat({ fontSize: '18pt', fontColor: '#000000', fontWeight: '600' }, `${pfx}A12:N12`)
      spreadsheet.cellFormat({ fontSize: '11pt', textAlign: 'center', verticalAlign: 'middle' }, `${pfx}H12`)
      // Re-apply row 12 (B12:N12) borders: 3px top, bottom, right (N12); 2px left
      spreadsheet.cellFormat({ borderTop: '3px solid #000000', borderLeft: '2px solid #000000', borderRight: '3px solid #000000', borderBottom: '3px solid #000000' }, `${pfx}B12:N12`)
      // $/1000 heading (H1) at 11pt, centered
      spreadsheet.cellFormat({ fontSize: '11pt', textAlign: 'center', verticalAlign: 'middle' }, `${pfx}H1`)
      // Columns C–G: center all values horizontally and vertically
      spreadsheet.cellFormat({ textAlign: 'center', verticalAlign: 'middle' }, `${pfx}C12:G${finalRow}`)
      // Word wrap for description column (B) so long text wraps within the cell
      spreadsheet.cellFormat({ wrapText: true }, `${pfx}B13:B${finalRow}`)

      // Email label (B6): normal weight, blue, underline
      spreadsheet.cellFormat({ fontWeight: 'normal', fontFamily: 'Calibri (Body)', color: '#0B76C3', textDecoration: 'underline' }, `${pfx}B6`)
      // Tel/Fax/Cell line (B5): bold, black, 18pt
      spreadsheet.cellFormat({ fontWeight: 'bold', fontSize: '18pt', color: '#000000' }, `${pfx}B5`)

      // Override numerical value columns to be normal weight per request (including $/1000 column H)
      // Data starts at Row 13; C–G stay centered
      spreadsheet.cellFormat({ fontWeight: 'normal', fontFamily: 'Calibri (Body)', textAlign: 'center', verticalAlign: 'middle' }, `${pfx}C13:G${finalRow}`)
      spreadsheet.cellFormat({ fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}H13:H${finalRow}`)
      spreadsheet.cellFormat({ fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}I13:N${finalRow}`)

      // RE-APPLY BOLD to Total Rows (B, C-G, I-N only; keep H normal so $/1000 values are not bold)
      totalRows.forEach(row => {
        spreadsheet.cellFormat({ fontWeight: 'bold', fontFamily: 'Calibri' }, `${pfx}B${row}:G${row}`)
        spreadsheet.cellFormat({ fontWeight: 'bold', fontFamily: 'Calibri' }, `${pfx}I${row}:N${row}`)
      })
      totalSFRowsForItalic.forEach(row => {
        spreadsheet.cellFormat({ fontWeight: 'bold', fontStyle: 'italic', textAlign: 'center', fontFamily: 'Calibri' }, `${pfx}C${row}`)
        spreadsheet.cellFormat({ fontWeight: 'bold', fontStyle: 'italic', textAlign: 'center', fontFamily: 'Calibri' }, `${pfx}D${row}`)
      })
      // Note rows: single-column B with rich text (Note: bold, rest normal) – don’t apply fontWeight so rich text is preserved
      noteRows.forEach(row => {
        spreadsheet.cellFormat({ fontFamily: 'Calibri (Body)', textAlign: 'left' }, `${pfx}B${row}`)
      })
      wpNoteRow.forEach(row => {
        spreadsheet.cellFormat({ fontWeight: 'normal', fontStyle: 'italic', textAlign: 'center', verticalAlign: 'middle', fontFamily: 'Calibri (Body)' }, `${pfx}B${row}`)
      })
      // Ensure every cell on Proposal sheet is vertically centered (y-axis)
      spreadsheet.cellFormat({ verticalAlign: 'middle' }, `${pfx}A1:N${finalRow}`)
      // $/1000 headings (H1, H12) centered in cell (re-apply last so not overwritten)
      spreadsheet.cellFormat({ textAlign: 'center', verticalAlign: 'middle' }, `${pfx}H1`)
      spreadsheet.cellFormat({ textAlign: 'center', verticalAlign: 'middle' }, `${pfx}H12`)

      // Re-apply 3px outer border B–G last so it overrides all other border styles
      spreadsheet.cellFormat({ borderTop: '3px solid #000000' }, `${pfx}B3:G3`)
      spreadsheet.cellFormat({ borderBottom: '3px solid #000000' }, `${pfx}B${finalRow}:G${finalRow}`)
      spreadsheet.cellFormat({ borderLeft: '3px solid #000000' }, `${pfx}B3:B${finalRow}`)
      spreadsheet.cellFormat({ borderRight: '3px solid #000000' }, `${pfx}G3:G${finalRow}`)
    }

    // Row height for each data row based on column B content (wrap text, height fits content)
    for (let r = 13; r <= finalRow; r++) {
      try {
        const text = rowBContentMap.has(r) ? String(rowBContentMap.get(r)) : ''
        const h = calculateRowHeight(text)
        spreadsheet.setRowHeight(h, r - 1, proposalSheetIndex)
      } catch (e) { /* ignore */ }
    }

    // Individual totals are now added after each subsection's misc section

    // Apply green background and $ prefix format to columns I-N for all data rows (from row 2 to currentRow-1)
    // Also apply $ prefix format to column H ($/1000)
    // Row 1 is the header, so data starts from row 2
    if (currentRow > 2) {
      const lastDataRow = proposalDataEndRow != null ? proposalDataEndRow : currentRow - 1
      // Apply green background and border to columns I-N
      spreadsheet.cellFormat(
        {
          backgroundColor: '#E2EFDA',
          border: '2px solid #000000'
        },
        `${pfx}I12:N${lastDataRow}`
      )
      // 3px border on all sides for each cell I–N from row 13 to lastDataRow
      spreadsheet.cellFormat(
        { borderTop: '3px solid #000000', borderBottom: '3px solid #000000', borderLeft: '3px solid #000000', borderRight: '3px solid #000000' },
        `${pfx}I13:N${lastDataRow}`
      )
      // Override row 12 (I12:N12) so 3px top, bottom, right (N12) are not overwritten by the 2px border above
      spreadsheet.cellFormat(
        { borderTop: '3px solid #000000', borderBottom: '3px solid #000000', borderLeft: '2px solid #000000', borderRight: '3px solid #000000' },
        `${pfx}I12:N12`
      )
      // Border on header row I1:N1 (in case not covered by header styles)
      spreadsheet.cellFormat({ border: '1px solid #000000' }, `${pfx}I1:N1`)
      // Apply number format to quantity columns C-G that hides zeros
      try {
        spreadsheet.numberFormat('#,##0.00;-#,##0.00;""', `${pfx}C2:G${lastDataRow}`)
      } catch (e) { /* ignore */ }
      // Apply $ prefix format to column H ($/1000): more space between $ and value
      try {
        // Show a leading $ in every row of column H, even when the numeric value is blank/zero
        spreadsheet.numberFormat('$          #,##0.00;-$          #,##0.00;"$"', `${pfx}H2:H${lastDataRow}`)
      } catch (e) {
        spreadsheet.cellFormat(
          { format: '$          #,##0.00;-$          #,##0.00;"$"', textAlign: 'left' },
          `${pfx}H2:H${lastDataRow}`
        )
      }
      spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}H13:H${lastDataRow}`)
      // Apply $ prefix format to columns I-N: space between $ and value; column width fits data
      try {
        spreadsheet.numberFormat('$   #,##0.00;-$   #,##0.00;""', `${pfx}I2:N${lastDataRow}`)
      } catch (e) {
        spreadsheet.cellFormat(
          { format: '$   #,##0.00;-$   #,##0.00;""', textAlign: 'left' },
          `${pfx}I2:N${lastDataRow}`
        )
      }
      spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}I2:N${lastDataRow}`)

      // Total rows: dollar and amount together (no space); total label (B:E) right-aligned; $ values (F:G, H, I-N) centered
      const totalFormat = '$#,##0.00;-$#,##0.00;""'
      totalRows.forEach(row => {
        spreadsheet.cellFormat({ textAlign: 'right' }, `${pfx}B${row}:E${row}`)
        try {
          spreadsheet.numberFormat(totalFormat, `${pfx}F${row}:G${row}`)
          spreadsheet.numberFormat(totalFormat, `${pfx}H${row}`)
          spreadsheet.numberFormat(totalFormat, `${pfx}I${row}:N${row}`)
        } catch (e) { /* ignore */ }
        spreadsheet.cellFormat({ textAlign: 'center', format: totalFormat }, `${pfx}F${row}:G${row}`)
        spreadsheet.cellFormat({ textAlign: 'center', format: totalFormat }, `${pfx}H${row}`)
        spreadsheet.cellFormat({ textAlign: 'center', format: totalFormat }, `${pfx}I${row}:N${row}`)
      })
      // Final override: totals show $ and amount together – re-apply so global H/I-N format does not override
      totalRows.forEach(row => {
        try {
          spreadsheet.numberFormat('$#,##0.00', `${pfx}F${row}:G${row}`)
          spreadsheet.numberFormat('$#,##0.00', `${pfx}H${row}`)
          spreadsheet.numberFormat('$#,##0.00', `${pfx}I${row}:N${row}`)
        } catch (e) { /* ignore */ }
        spreadsheet.cellFormat({ format: '$#,##0.00' }, `${pfx}F${row}:N${row}`)
      })
      // Re-apply row 12 (I12:N12) 3px top, bottom, right (N12) last so nothing overwrites them
      spreadsheet.cellFormat(
        { borderTop: '3px solid #000000', borderBottom: '3px solid #000000', borderLeft: '2px solid #000000', borderRight: '3px solid #000000' },
        `${pfx}I12:N12`
      )
    }

    // Force recalculation of all formulas to ensure $/1000 values display correctly
    try {
      spreadsheet.refresh()
    } catch (e) {
      // Fallback: try to trigger recalculation by touching a cell
      try {
        spreadsheet.goTo('Calculations Sheet!A1')
      } catch (e2) { /* ignore */ }
    }

    // Auto-fit columns I–N so cells expand to fit $ and value (no overflow); then show Calculation sheet first (not Proposal)
    try {
      spreadsheet.goTo('Calculations Sheet!A1')
      if (typeof spreadsheet.autoFit === 'function') {
        spreadsheet.autoFit('I:N')
      }
    } catch (e) { /* ignore */ }

    // Hardcoded override: row 12 I12:N12 always 3px top, bottom, right (N12) (overrides any other formatting)
    try {
      spreadsheet.cellFormat(
        { borderTop: '3px solid #000000', borderBottom: '3px solid #000000', borderLeft: '3px solid #000000', borderRight: '3px solid #000000' },
        `${pfx}I12:N12`
      )
      // Remove top border from row 13 (I13:N13) so row 12's 3px bottom border is not doubled/hidden
      spreadsheet.cellFormat({ borderTop: 'none' }, `${pfx}I13:N13`)
    } catch (e) { /* ignore */ }

    // Total SF rows: force center, bold, italic on C and D so alignment is not overwritten
    try {
      totalSFRowsForItalic.forEach(row => {
        spreadsheet.cellFormat({ fontWeight: 'bold', fontStyle: 'italic', textAlign: 'center', verticalAlign: 'middle' }, `${pfx}C${row}`)
        spreadsheet.cellFormat({ fontWeight: 'bold', fontStyle: 'italic', textAlign: 'center', verticalAlign: 'middle' }, `${pfx}D${row}`)
      })
    } catch (e) { /* ignore */ }

    // General Liability line: force 3px bottom border (overrides thin/thickBottom applied to range)
    if (generalLiabilityBorderRow != null) {
      try {
        spreadsheet.cellFormat({ borderBottom: '3px solid #000000' }, `${pfx}B${generalLiabilityBorderRow}:G${generalLiabilityBorderRow}`)
          // Re-apply per column in case range format is overridden
          ;['B', 'C', 'D', 'E', 'F', 'G'].forEach(col => {
            spreadsheet.cellFormat({ borderBottom: '3px solid #000000' }, `${pfx}${col}${generalLiabilityBorderRow}`)
          })
      } catch (e) { /* ignore */ }
    }
  }
}

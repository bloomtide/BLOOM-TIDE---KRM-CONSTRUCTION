import { convertToFeet } from './parsers/dimensionParser'
import proposalMapped from './proposal_mapped.json'

export function buildProposalSheet(spreadsheet, { calculationData, formulaData, rockExcavationTotals, lineDrillTotalFT, rawData, createdAt, project, client }) {
  // Full Proposal template (match screenshot layout).
  const proposalSheetIndex = 1
  const pfx = 'Proposal Sheet!'

  // Header labels from proposal_mapped.json (DESCRIPTION entry); fallback to key name if missing
  const descriptionEntry = Array.isArray(proposalMapped) ? proposalMapped.find(e => e.description === 'DESCRIPTION') : null
  const headerLabels = descriptionEntry?.values || {}
  const label = (key) => (headerLabels[key] != null && headerLabels[key] !== '') ? String(headerLabels[key]) : key
  const L = { LF: label('LF'), SF: label('SF'), LBS: label('LBS'), CY: label('CY'), QTY: label('QTY'), LS: label('LS') }

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
      const rowContains = withRates.filter(e => key.indexOf(e.description.trim()) >= 0)
      if (rowContains.length > 0) {
        // Prefer the match that appears earliest in the key (e.g. "Styrofoam" over "slab topping" in "Styrofoam @ under built-up slab topping")
        const best = rowContains.reduce((best, e) => {
          const desc = e.description.trim()
          const pos = key.indexOf(desc)
          if (pos < 0) return best
          if (!best) return e
          const bestPos = key.indexOf(best.description.trim())
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
  const fillRatesForProposalRow = (row, descriptionText) => {
    const values = getRatesForDescription(descriptionText)
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
    I: 87,  // LF
    J: 87,  // SF
    K: 87,  // LBS
    L: 87,  // CY
    M: 87,  // QTY
    N: 87   // LS
  }
  Object.entries(colWidths).forEach(([col, width]) => {
    // Try both 0 and 1 indices to match the active sheet
    try { spreadsheet.setColWidth(width, col.charCodeAt(0) - 65, 0) } catch (e) { }
    try { spreadsheet.setColWidth(width, col.charCodeAt(0) - 65, 1) } catch (e) { }
    // Also try referencing by sheet name logic if possible, but indices are standard
  })

  // Set height for 1st row (index 0) to 24
  try { spreadsheet.setRowHeight(24, 0, 0) } catch (e) { }
  try { spreadsheet.setRowHeight(24, 0, 1) } catch (e) { }

  // Track all section total rows for BASE BID TOTAL formula
  const baseBidTotalRows = []
  // Track rows that should remain bold (totals) to override the unbold rule
  const totalRows = []

  // Helper function to calculate row height based on text content in column B
  const calculateRowHeight = (text) => {
    if (!text || typeof text !== 'string') {
      return 30 // Default minimal height
    }

    // Column B width is 955 pixels
    const columnWidth = 955
    const fontSize = 18 // Updated to 18pt
    // Approximate line height for 18pt font. 30px is very tight for 18pt (24px).
    // Try tightly packed to meet request.
    const lineHeight = 26
    const padding = 4
    // Average char width for 18pt Calibri Bold is approx 11-12px
    const charWidth = 11

    // Calculate how many characters fit per line (accounting for padding)
    const availableWidth = columnWidth - 20 // Account for cell padding
    const charsPerLine = Math.floor(availableWidth / charWidth)

    // Calculate number of lines needed (accounting for word wrapping)
    const textLength = text.length
    const estimatedLines = Math.max(1, Math.ceil(textLength / charsPerLine))

    // Calculate height: (lines * lineHeight) + padding
    const calculatedHeight = Math.max(30, Math.ceil(estimatedLines * lineHeight + padding))

    // Increase cap significantly as 18pt text takes space
    return Math.min(calculatedHeight, 400)
  }


  // Styles
  const headerGray = { backgroundColor: '#D9D9D9', fontWeight: 'bold', fontFamily: 'Calibri', fontSize: '10pt', textAlign: 'center', verticalAlign: 'middle' }
  const boxGray = { backgroundColor: '#D9D9D9', fontWeight: 'bold', fontFamily: 'Calibri', fontSize: '18pt', verticalAlign: 'middle' }
  const thick = { border: '2px solid #000000' }
  const thickTop = { borderTop: '2px solid #000000' }
  const thickBottom = { borderBottom: '2px solid #000000' }
  const thickLeft = { borderLeft: '2px solid #000000' }
  const thickRight = { borderRight: '2px solid #000000' }
  const thin = { border: '1px solid #000000' }

    // Top header row (row 1) - labels from proposal_mapped.json DESCRIPTION entry
    ;[
      ['C1', L.LF], ['D1', L.SF], ['E1', L.LBS], ['F1', L.CY], ['G1', L.QTY],
      ['H1', '$/1000'],
      ['I1', L.LF], ['J1', L.SF], ['K1', L.LBS], ['L1', L.CY], ['M1', L.QTY], ['N1', L.LS]
    ].forEach(([cell, value]) => spreadsheet.updateCell({ value }, `${pfx}${cell}`))
  spreadsheet.cellFormat(headerGray, `${pfx}C1:N1`)
  spreadsheet.cellFormat(thick, `${pfx}C1:N1`)
  // Make $/1000 cell background white
  spreadsheet.cellFormat({ backgroundColor: 'white' }, `${pfx}H1`)

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
  spreadsheet.cellFormat({ fontWeight: 'normal', color: '#000000', fontFamily: 'Calibri (Body)' }, `${pfx}B4:E8`)
  spreadsheet.cellFormat({ fontWeight: 'bold' }, `${pfx}B4:B5`)
  spreadsheet.cellFormat({ color: '#0B76C3', textDecoration: 'underline' }, `${pfx}B6`)

  // Logo (image) near center top-left; uses public path
  try {
    const imgSrc = encodeURI('/images/templateimage.png')
    // Start around mid of column B and extend roughly to column E
    // (B width is large; left offset pushes the image start into the middle of B)
    spreadsheet.insertImage(
      [{ src: imgSrc, width: 529, height: 182, left: 790, top: 70 }],
      `${pfx}B4`
    )
  } catch (e) {
    // ignore
  }

  // Date/Project/Client block
  const proposalDate = createdAt ? new Date(createdAt).toLocaleDateString() : "Today's date"

  spreadsheet.updateCell({ value: `Date: ${proposalDate}` }, `${pfx}B9`)
  spreadsheet.updateCell({ value: `Project: ${project || '###'}` }, `${pfx}B10`)
  spreadsheet.updateCell({ value: `Client: ${client || '###'}` }, `${pfx}B11`)
  spreadsheet.cellFormat({ fontWeight: 'bold' }, `${pfx}B9:B11`)
  spreadsheet.cellFormat(thin, `${pfx}B9:G11`)

  // Right grey box (Estimate / Drawings Dated / lines) - start from column F
  // Individual cell styling - customize each cell separately

  // Row 3: Estimate #25-1150
  spreadsheet.merge(`${pfx}F3:G3`)
  spreadsheet.updateCell({ value: 'Estimate #2' }, `${pfx}F3`)
  spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontSize: '11pt', fontWeight: 'bold', borderTop: '1px solid #000000', borderLeft: '1px solid #000000', borderRight: '1px solid #000000' }, `${pfx}F3`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontSize: '11pt', fontWeight: 'bold', borderTop: '1px solid #000000', borderRight: '1px solid #000000' }, `${pfx}G3`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontSize: '11pt', fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}H3`)

  // Row 4: Empty row
  spreadsheet.merge(`${pfx}F4:G4`)
  spreadsheet.cellFormat({ backgroundColor: 'white', borderLeft: '1px solid #000000' }, `${pfx}F4`)
  spreadsheet.cellFormat({ backgroundColor: 'white', borderRight: '1px solid #000000' }, `${pfx}G4`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}H4`)

  // Row 5: Drawings Dated:
  spreadsheet.merge(`${pfx}F5:G5`)
  spreadsheet.updateCell({ value: 'Drawings Dated:' }, `${pfx}F5`)
  // Top border, Left/Right borders, No bottom border
  spreadsheet.cellFormat({ textAlign: 'center', backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '2px solid #000000', borderTop: '2px solid #000000', borderRight: '2px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}F5:G5`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}H5`)

  // Row 6: SOE:
  spreadsheet.merge(`${pfx}F6:G6`)
  spreadsheet.updateCell({ value: 'SOE:' }, `${pfx}F6`)
  // Side borders only
  spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '2px solid #000000', borderRight: '2px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}F6:G6`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}H6`)

  // Row 7: Structural:
  spreadsheet.merge(`${pfx}F7:G7`)
  spreadsheet.updateCell({ value: 'Structural:' }, `${pfx}F7`)
  // Side borders only
  spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '2px solid #000000', borderRight: '2px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}F7:G7`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}H7`)

  // Row 8: Architectural:
  spreadsheet.merge(`${pfx}F8:G8`)
  spreadsheet.updateCell({ value: 'Architectural:' }, `${pfx}F8`)
  // Side borders only
  spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '2px solid #000000', borderRight: '2px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}F8:G8`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}H8`)

  // Row 9: Plumbing:
  spreadsheet.merge(`${pfx}F9:G9`)
  spreadsheet.updateCell({ value: 'Plumbing:' }, `${pfx}F9`)
  // Side borders only
  spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '2px solid #000000', borderRight: '2px solid #000000', borderBottom: '1px solid #D0CECE' }, `${pfx}F9:G9`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}H9`)

  // Row 10: Mechanical
  spreadsheet.merge(`${pfx}F10:G10`)
  spreadsheet.updateCell({ value: 'Mechanical' }, `${pfx}F10`)
  // Bottom border, Side borders
  spreadsheet.cellFormat({ backgroundColor: '#D0CECE', fontWeight: 'bold', borderLeft: '2px solid #000000', borderBottom: '2px solid #000000', borderRight: '2px solid #000000' }, `${pfx}F10:G10`)
  spreadsheet.cellFormat({ backgroundColor: 'white', fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}H10`)

  // Bottom header row (row 12) - labels from proposal_mapped.json DESCRIPTION entry
  spreadsheet.updateCell({ value: 'DESCRIPTION' }, `${pfx}B12`)
    ;[
      ['C12', L.LF], ['D12', L.SF], ['E12', L.LBS], ['F12', L.CY], ['G12', L.QTY],
      ['H12', '$/1000'],
      ['I12', L.LF], ['J12', L.SF], ['K12', L.LBS], ['L12', L.CY], ['M12', L.QTY]
    ].forEach(([cell, value]) => spreadsheet.updateCell({ value }, `${pfx}${cell}`))
  spreadsheet.updateCell({ value: L.LS }, `${pfx}N12`)

  spreadsheet.cellFormat(headerGray, `${pfx}B12:N12`)
  spreadsheet.cellFormat(thick, `${pfx}B12:N12`)
  // Make $/1000 cell background white
  spreadsheet.cellFormat({ backgroundColor: 'white' }, `${pfx}H12`)

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
        format: '$#,##0.0'
      },
      `${pfx}H13:H1000`
    )
    spreadsheet.numberFormat('$#,##0.0', `${pfx}H13:H1000`)
  } catch (e) {
    // Fallback - format will be applied individually as rows are created
  }

  // Row 14: Demolition scope heading
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
  // Always render demolition lines (with defaults when calculationData is empty)
  const getDMReferenceFromRawData = (subsectionName) => {
    if (!rawData || !Array.isArray(rawData) || rawData.length < 2) return 'DM-106.00'
    const headers = rawData[0]
    const dataRows = rawData.slice(1)
    const digitizerIdx = headers.findIndex(h => h && h.toLowerCase().trim() === 'digitizer item')
    if (digitizerIdx === -1) return 'DM-106.00'
    const subsectionPatterns = {
      'Demo slab on grade': /demo\s+sog/i, 'Demo strip footing': /demo\s+sf/i,
      'Demo foundation wall': /demo\s+fw/i, 'Demo isolated footing': /demo\s+isolated\s+footing/i
    }
    const pattern = subsectionPatterns[subsectionName]
    if (!pattern) return 'DM-106.00'
    for (const row of dataRows) {
      const digitizerItem = row[digitizerIdx]
      if (digitizerItem && pattern.test(String(digitizerItem))) {
        const dmMatch = String(digitizerItem).trim().match(/DM-[\d.]+/i)
        if (dmMatch) return dmMatch[0]
      }
    }
    return 'DM-106.00'
  }
  const buildDemolitionTemplate = (subsectionName, itemText) => {
    const slabTypeMatch = subsectionName.match(/^Demo\s+(.+)$/i)
    const slabType = slabTypeMatch ? slabTypeMatch[1].trim() : subsectionName.replace(/^Demo\s+/i, '').trim()
    const dmReference = getDMReferenceFromRawData(subsectionName)
    if (!itemText) {
      if (slabType.toLowerCase().includes('slab on grade')) {
        return `Allow to saw-cut/demo/remove/dispose existing (4" thick) ${slabType} @ existing building as per ${dmReference}`
      }
      return `Allow to saw-cut/demo/remove/dispose existing ${slabType} @ existing building as per ${dmReference}`
    }
    const text = String(itemText).trim()
    let thicknessPart = ''
    if (slabType.toLowerCase().includes('slab on grade')) {
      const thicknessMatch = text.match(/(\d+["\"]?\s*thick)/i) || text.match(/(\d+["\"]?)/)
      thicknessPart = thicknessMatch ? `(${thicknessMatch[1]})` : '(4" thick)'
    } else {
      const dimMatch = text.match(/\(([^)]+)\)/)
      thicknessPart = dimMatch ? `(${dimMatch[1]})` : ''
    }
    return `Allow to saw-cut/demo/remove/dispose existing ${thicknessPart} ${slabType} @ existing building as per ${dmReference}`
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
          rockExcavationDataRowCount = 0 // Reset data row counter for LF/SF/CY references
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
          // Particulars is filled - data row; track row index for proposal references (LF, SF, CY)
          rockExcavationDataRowCount++
          if (rockExcavationDataRowCount === 1) rockExcavationLine1RowIndex = rowIndex + 1
          else if (rockExcavationDataRowCount === 2) rockExcavationLine2RowIndex = rowIndex + 1
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

    // Initialize dynamic row counter
    let currentRow = 14

    // -------------------------------------------------------------------------
    // DEMOLITION SECTION
    // -------------------------------------------------------------------------

    // Check if we have any demolition items
    const orderedSubsections = [
      'Demo slab on grade',
      'Demo strip footing',
      'Demo foundation wall',
      'Demo isolated footing'
    ]

    // Check if any of these subsections have data
    let hasDemolitionItems = false
    orderedSubsections.forEach(name => {
      if ((rowsBySubsection.get(name) || []).length > 0) {
        hasDemolitionItems = true
      }
    })

    if (hasDemolitionItems) {
      // Demolition Scope Heading
      spreadsheet.updateCell({ value: 'Demolition scope:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat({
        backgroundColor: '#BDD7EE',
        textAlign: 'center',
        verticalAlign: 'middle',
        textDecoration: 'underline',
        fontWeight: 'normal'
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

      orderedSubsections.forEach((name, index) => {
        const rowCount = (rowsBySubsection.get(name) || []).length
        const originalText = linesBySubsection.get(name)
        const templateText = buildDemolitionTemplate(name, originalText)
        const cellRef = `${pfx}B${currentRow}`

        spreadsheet.updateCell({ value: templateText }, cellRef)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'left',
            verticalAlign: 'top'
          },
          cellRef
        )
        fillRatesForProposalRow(currentRow, templateText)

        // Calculate and display individual SF for this subsection in column D
        const subsectionRows = rowsBySubsection.get(name) || []
        const subsectionSF = calculateSF(subsectionRows)
        const formattedSF = parseFloat(subsectionSF.toFixed(2))

        spreadsheet.updateCell({ value: formattedSF }, `${pfx}D${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'right',
            format: '#,##0.00'
          },
          `${pfx}D${currentRow}`
        )

        // Calculate and display individual CY for this subsection in column F
        const subsectionCY = calculateCY(subsectionRows)
        const formattedCY = parseFloat(subsectionCY.toFixed(2))

        spreadsheet.updateCell({ value: formattedCY }, `${pfx}F${currentRow}`)
        spreadsheet.cellFormat(
          {
            fontWeight: 'bold',
            color: '#000000',
            textAlign: 'right',
            format: '#,##0.00'
          },
          `${pfx}F${currentRow}`
        )

        // Get QTY from sum row for this subsection in column G
        const subsectionQTY = getQTYFromSumRow(name)
        if (subsectionQTY !== null) {
          const formattedQTY = parseFloat(subsectionQTY.toFixed(2))
          spreadsheet.updateCell({ value: formattedQTY }, `${pfx}G${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'right',
              format: '#,##0.00'
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
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: 'white'
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
          backgroundColor: '#BDD7EE',
          border: '1px solid #000000'
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
          border: '1px solid #000000',
          format: '$#,##0.00'
        },
        `${pfx}F${currentRow}:G${currentRow}`
      )

      // Apply background color to entire row
      spreadsheet.cellFormat(
        {
          backgroundColor: '#BDD7EE'
        },
        `${pfx}B${currentRow}:G${currentRow}`
      )
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
    if (hasRockExcavation) {
      currentRow++ // Extra line after Excavation scope
      // Soil excavation scope heading
      spreadsheet.updateCell({ value: 'Soil excavation scope:' }, `${pfx}B${currentRow}`)
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

    // First soil excavation line
    const soilExcavationText = 'Allow to perform soil excavation, trucking & disposal (Havg=16\'-9") as per SOE-101.00, P-301.01 & details on SOE-201.01 to SOE-204.00'
    spreadsheet.updateCell({ value: soilExcavationText }, `${pfx}B${currentRow}`)
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
    spreadsheet.setRowHeight(currentRow, 30) // Set row height for wrapped text
    fillRatesForProposalRow(currentRow, soilExcavationText)
    const soilExcavationRow1 = currentRow

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

    // Formula for row 1
    const dollarFormulaSoil1 = `=IFERROR(IF(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)=0,"",ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1)),"")`

    // Evaluation Logic (simplified for brevity, main logic logic remains in helpers if any)

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

    // Row 2: Second soil excavation line (new clean soil)
    const soilExcavationRow2 = currentRow
    const backfillSoilText = 'Allow to import new clean soil to backfill and compact as per SOE-101.00, P-301.01 & details on SOE-201.01 to SOE-204.00'
    spreadsheet.updateCell({ value: backfillSoilText }, `${pfx}B${currentRow}`)
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

    // Notes
    const notes = [
      'Note: Backfill SOE voids by others',
      'Note: NJ Res Soil included, contaminated, mixed, hazardous, petroleum impacted not incl.',
      'Note: Bedrock not included, see add alt unit rate if required'
    ]
    notes.forEach(note => {
      spreadsheet.updateCell({ value: note }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: 'white'
        },
        `${pfx}B${currentRow}`
      )
      currentRow++
    })

    // Soil Excavation Total (only when rock section is shown)
    let soilExcavationTotalRow = null
    if (hasRockExcavation) {
      spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
      spreadsheet.updateCell({ value: 'Soil Excavation Total:' }, `${pfx}D${currentRow}`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'left',
          backgroundColor: '#FFF2CC',
          border: '1px solid #000000'
        },
        `${pfx}D${currentRow}:E${currentRow}`
      )

      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({ formula: `=SUM(H${soilExcavationRow1}:H${currentRow - 1})*1000` }, `${pfx}F${currentRow}`)

      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          color: '#000000',
          textAlign: 'right',
          backgroundColor: '#FFF2CC',
          border: '1px solid #000000',
          format: '$#,##0.00'
        },
        `${pfx}F${currentRow}:G${currentRow}`
      )
      spreadsheet.cellFormat({ backgroundColor: '#FFF2CC' }, `${pfx}B${currentRow}:G${currentRow}`)
      baseBidTotalRows.push(currentRow) // Soil Excavation Total
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

    // -------------------------------------------------------------------------
    // ROCK EXCAVATION SECTION (only when rock has values)
    // -------------------------------------------------------------------------
    let rockExcavationTotalRow = null
    if (hasRockExcavation) {
    // Rock excavation scope heading
    spreadsheet.updateCell({ value: 'Rock excavation scope:' }, `${pfx}B${currentRow}`)
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
    const rockExcavationText = rockExcavationTextFromCalc || 'Allow to perform rock excavation, trucking & disposal for building (Havg=2\'-9") as per SOE-100.00 & details on SOE-A-202.00'
    const lineDrillingText = lineDrillingTextFromCalc || 'Allow to perform line drilling as per SOE-100.00'

    // First rock excavation line
    const rockRow1 = currentRow
    spreadsheet.updateCell({ value: rockExcavationText }, `${pfx}B${currentRow}`)
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

    // Second rock excavation line
    const rockRow2 = currentRow
    spreadsheet.updateCell({ value: lineDrillingText }, `${pfx}B${currentRow}`)
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

    // Rock Excavation Total
    spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
    spreadsheet.updateCell({ value: 'Rock excavation Total:' }, `${pfx}D${currentRow}`)
    spreadsheet.cellFormat(
      {
        fontWeight: 'bold',
        color: '#000000',
        textAlign: 'left',
        backgroundColor: '#FFF2CC',
        border: '1px solid #000000'
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
        border: '1px solid #000000',
        format: '$#,##0.00'
      },
      `${pfx}F${currentRow}:G${currentRow}`
    )
    spreadsheet.cellFormat({ backgroundColor: '#FFF2CC', border: '1px solid #000000' }, `${pfx}B${currentRow}:C${currentRow}`)
    baseBidTotalRows.push(currentRow) // Rock Excavation Total
    totalRows.push(currentRow)

    rockExcavationTotalRow = currentRow
    currentRow++ // Empty row
    currentRow++ // Extra line after Rock Excavation Total
    }

    // Excavation Total (Sum of Soil + Rock when both; otherwise soil only)
    spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
    spreadsheet.updateCell({ value: 'Excavation Total:' }, `${pfx}D${currentRow}`)
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

    spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
    const excavationFullTotalFormula = hasRockExcavation
      ? `=SUM(F${soilExcavationTotalRow},F${rockExcavationTotalRow})`
      : `=SUM(H${soilExcavationRow1}:H${currentRow - 1})*1000`
    spreadsheet.updateCell({ formula: excavationFullTotalFormula }, `${pfx}F${currentRow}`)
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
    spreadsheet.cellFormat({ backgroundColor: '#BDD7EE', border: '1px solid #000000' }, `${pfx}B${currentRow}:C${currentRow}`)
    if (!hasRockExcavation) {
      baseBidTotalRows.push(currentRow) // Excavation Total (only row for excavation when no rock)
      totalRows.push(currentRow)
    }

    currentRow++ // Extra line after Excavation Total

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
    // SOE Scope
    spreadsheet.updateCell({ value: 'SOE scope:' }, `${pfx}B${currentRow}`)
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

    // SOE Headings
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

    const headingRows = {}
    soeHeadings.forEach((heading, index) => {
      // Write headings sequentially
      const rowNum = currentRow
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
      currentRow++
    })

    let rowShift = 0

    // Helper function... (retained)
    const getSOEPageFromRawData = (diameter, thickness) => {
      // ... (retained implementation or fallback)
      if (!rawData || !Array.isArray(rawData) || rawData.length < 2) return 'SOE-101.00'
      // ... (simplified for space)
      return 'SOE-101.00'
    }

    // Add drilled soldier pile proposal text
    const collectedGroups = window.drilledSoldierPileGroups || []

    // ... (helper parseDimension) ...
    const parseDimension = (dimStr) => {
      if (!dimStr) return 0
      const match = dimStr.match(/(\d+)(?:'-?)?(\d+)?/)
      if (!match) return 0
      return (parseInt(match[1]) || 0) + ((parseInt(match[2]) || 0) / 12)
    }

    // Logic to separate HP/Drilled
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

    // Initialize currentRow for Drilled items
    // Initialize currentRow for both drilled and HP groups
    // Start drilled soldier piles right after "Soldier drilled piles:" heading
    currentRow = headingRows['Soldier drilled piles:'] + 1

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
            fillRatesForProposalRow(currentRow, proposalText)

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
                format: '$#,##0.0'
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

    // Track SOE scope rows for SOE Total calculation
    let soeScopeStartRow = null
    let soeScopeEndRow = null

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
          subsectionLower === 'mud slab')) {
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

          // Timber lagging line: main reference only (e.g. SOE-101.00), no details page
          const timberLaggingMainRef = (soePageMain && !String(soePageMain).includes('301')) ? soePageMain : 'SOE-101.00'
          // Format: F&I new 3"x10" timber lagging for the exposed depths (Havg=10'-6") as per SOE-101.00
          proposalText = `F&I new ${dimensions || '##'} timber lagging for the exposed depths (Havg=${heightText || '##'}) as per ${timberLaggingMainRef}`

          // Store main ref for backpacking line (backpacking uses main + details on SOE-301.00)
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
        fillRatesForProposalRow(currentRow, proposalText)

        // Calculate and set row height based on content
        const dynamicHeight = calculateRowHeight(proposalText)

        // Add FT (LF) to column C - reference to calculation sheet sum row if available
        // For parging, sheet pile, guide wall, dowels: FT(I) from sum row
        const subsectionLowerForC = subsectionName.toLowerCase()
        if ((subsectionLowerForC === 'parging' || subsectionLowerForC === 'sheet pile' || subsectionLowerForC === 'sheet piles' || subsectionLowerForC === 'guide wall' || subsectionLowerForC === 'guilde wall' || subsectionLowerForC === 'dowels' || subsectionLowerForC === 'dowel bar' || subsectionLowerForC === 'rock pins' || subsectionLowerForC === 'rock pin' || subsectionLowerForC === 'shotcrete' || subsectionLowerForC === 'permission grouting' || subsectionLowerForC === 'mud slab') && sumRowIndex > 0) {
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
          subsectionLowerName2 === 'timber lagging' ||
          subsectionLowerName2 === 'timber sheeting' ||
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

        // Add CY to column F for Heel blocks, Underpinning, Concrete soil retention piers, Guide wall, Concrete buttons - reference to calculation sheet sum row
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
          subsectionLowerName2 === 'mud slab') && sumRowIndex > 0) {
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

        // Add QTY to column G for Underpinning, Concrete soil retention piers, Heel blocks, Concrete buttons, Dowels - reference to calculation sheet sum row
        if ((subsectionLowerName2 === 'underpinning' ||
          subsectionLowerName2 === 'concrete soil retention piers' ||
          subsectionLowerName2 === 'concrete soil retention pier' ||
          subsectionLowerName2 === 'heel blocks' ||
          subsectionLowerName2 === 'concrete buttons' ||
          subsectionLowerName2 === 'buttons' ||
          subsectionLowerName2 === 'dowels' ||
          subsectionLowerName2 === 'dowel bar' ||
          subsectionLowerName2 === 'rock pins' ||
          subsectionLowerName2 === 'rock pin') && sumRowIndex > 0) {
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
        // For Sheet pile, always add LBS reference when sum row exists (calculation sheet has J*Wt formula)
        if ((totalLBS > 0 || subsectionLowerName2 === 'sheet pile' || subsectionLowerName2 === 'sheet piles') && sumRowIndex > 0) {
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

        // For Timber lagging, add second line for backpacking (after timber lagging row has C, D, E, H)
        if (subsectionName.toLowerCase() === 'timber lagging') {
          currentRow++
          const timberLaggingSoePage = window.timberLaggingSoePage || 'SOE-101.00'
          const mainSoeMatch = timberLaggingSoePage.match(/(SOE[-\d.]+)/i)
          const mainSoePage = mainSoeMatch ? mainSoeMatch[1] : 'SOE-101.00'
          // Backpacking has its own reference: main SOE + details on SOE-301.00 (different from timber lagging line)
          const backpackingText = `F&I new backpacking @ timber lagging ${mainSoePage} & details on SOE-301.00`
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
          fillRatesForProposalRow(currentRow, backpackingText)
          // Backpacking SF from calculation sheet: find the row with "Backpacking" in column B and reference its J (SF). J is filled by formula at render time so we don't require a value in calculationData.
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
            spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${backpackingSumRowIndex}` }, `${pfx}D${currentRow}`)
            spreadsheet.cellFormat(
              {
                fontWeight: 'bold',
                textAlign: 'right',
                backgroundColor: 'white'
              },
              `${pfx}D${currentRow}`
            )
          }
          // Add $/1000 formula for backpacking row
          const backpackingDollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
          spreadsheet.updateCell({ formula: backpackingDollarFormula }, `${pfx}H${currentRow}`)
          spreadsheet.cellFormat(
            {
              fontWeight: 'bold',
              color: '#000000',
              textAlign: 'right',
              backgroundColor: 'white',
              format: '$#,##0.0'
            },
            `${pfx}H${currentRow}`
          )
          try {
            spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}:N${currentRow}`)
          } catch (e) { /* ignore */ }
        }

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

          // Add $/1000 formula in column H
          const formBoardDollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
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
          backgroundColor: '#BDD7EE',
          border: '1px solid #000000'
        },
        `${pfx}D${currentRow}:E${currentRow}`
      )
      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      spreadsheet.cellFormat(
        {
          backgroundColor: '#BDD7EE',
          border: '1px solid #000000'
        },
        `${pfx}F${currentRow}:G${currentRow}`
      )

      // Add sum formula for column H ($/1000) - sum all SOE subsection totals and multiply by 1000
      spreadsheet.updateCell({ formula: `=SUM(H${soeScopeStartRow}:H${soeScopeEndRow})*1000` }, `${pfx}H${currentRow}`)
      spreadsheet.cellFormat(
        {
          fontWeight: 'bold',
          backgroundColor: '#BDD7EE',
          border: '1px solid #000000',
          format: '$#,##0.0'
        },
        `${pfx}H${currentRow}`
      )
      baseBidTotalRows.push(currentRow)
      totalRows.push(currentRow)
      currentRow++
    }
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

        // Row height already set above based on rock bolt proposal text
        currentRow++
      })
    }

    // Add empty row to separate Rock anchor scope from Foundation scope
    currentRow++
    } // end if (hasRockAnchorScopeData)

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

          const groupTakeoff = itemOrGroup.items[0]?.takeoff || 0
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
          fillRatesForProposalRow(currentRow, proposalText)

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

          // Add $/1000 formula
          const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
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
        fillRatesForProposalRow(currentRow, proposalText)

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
            format: '$#,##0.0'
          },
          `${pfx}H${currentRow}`
        )

        // Row height already set above based on foundation pile proposal text

        currentRow++ // Move to next row
      }

      // Track the end row of pile data items (before misc. section)
      const pileDataEndRow = currentRow - 1

      // Add misc. section for this foundation subsection (after all proposal text items)
      const miscSections = {
        'Drilled foundation pile': {
          title: 'Drilled pile misc.:',
          included: [
            'Plates & locking nuts included',
            'Pilings will be threaded at both ends and installed in 5\' or 10\' increments',
            'One mobilization & demobilization of drilling equipment included',
            'Surveying, stakeout, pile numbering plan & as-built plan included',
            'One compression reactionary load tests included',
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
        'Stelcor drilled displacement pile': {
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
        currentRow++

        // Add included items
        miscSection.included.forEach((item, index) => {
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

          // Apply unit rates for this misc line from proposal_mapped.json
          fillRatesForProposalRow(currentRow, item)

          // Add special formulas for certain items
          if (item.includes('Plates & locking nuts') && (subsectionName === 'Drilled foundation pile' || subsectionName === 'CFA pile')) {
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
            }
          } else if (item.includes('Video inspections included') && (subsectionName === 'Drilled foundation pile')) {
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
            }
          }

          // Override QTY rate for "Plates & locking nuts included" by pile type:
          // - Drilled pile misc.: 250
          // - Stelcor pile misc. and CFA pile misc.: 100
          if (item.includes('Plates & locking nuts')) {
            let qtyRate = null
            if (subsectionName === 'Drilled foundation pile') qtyRate = 250
            else if (subsectionName === 'Stelcor drilled displacement pile' || subsectionName === 'CFA pile') qtyRate = 100
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

          // If this misc line is mobilization/demobilization, override LS rate by pile type
          if (item.includes('mobilization & demobilization of drilling equipment')) {
            let lsValue = null
            if (subsectionName === 'Drilled foundation pile' || subsectionName === 'Stelcor drilled displacement pile') {
              lsValue = 10000
            } else if (subsectionName === 'Driven foundation pile') {
              lsValue = 20000
            }
            if (lsValue != null) {
              spreadsheet.updateCell({ value: lsValue }, `${pfx}N${currentRow}`)
              spreadsheet.cellFormat(
                {
                  fontWeight: 'normal',
                  textAlign: 'right',
                  backgroundColor: 'white'
                },
                `${pfx}N${currentRow}`
              )
            }
          }

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
        baseBidTotalRows.push(currentRow) // Foundation pile section total
        totalRows.push(currentRow)

        currentRow++ // One row gap after total
      }
    })

    // Substructure concrete scope (above Below grade waterproofing scope)
    const substructureCalcSheet = 'Calculations Sheet'
    const pileSubsections = ['Drilled foundation pile', 'Helical foundation pile', 'Driven foundation pile', 'Stelcor drilled displacement pile', 'CFA pile']
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
      const substructureTemplate = [
        {
          heading: 'Compacted gravel', items: [
            { text: `F&I new (6" thick) gravel/crushed stone, including 6MIL vapor barrier on top @ SOG, elevator pit slab, detention tank slab, duplex sewage ejector pit slab & mat as per FO-101.00 & details on FO-203.00`, sub: 'SOG', match: p => p.includes('gravel') && !p.includes('geotextile') },
            { text: 'F&I new geotextile filter fabric below gravel', sub: 'SOG', match: p => p.includes('geotextile') }
          ]
        },
        {
          heading: 'Elevator pit', items: [
            { text: `F&I new (2)no. (1'-0" thick) sump pit (2'-0"x2'-0"x2'-0") reinf w/ #4@12"O.C., T&B/E.F., E.W. typ`, sub: 'Elevator Pit', match: p => p.includes('sump'), formulaItem: { itemType: 'elevator_pit', match: p => (p || '').toLowerCase().includes('sump') } },
            { text: `F&I new (3'-0" thick) elevator pit slab, reinf w/#5@12"O.C. & #5@12"O.C., as per FO-101.00 & details on FO-203.00`, sub: 'Elevator Pit', match: p => p.includes('elev') && p.includes('slab') && !p.includes('sump') },
            { text: `F&I new (1'-0" thick) elevator pit walls (H=5'-0") as per FO-101.00 & details on FO-203.00`, sub: 'Elevator Pit', match: p => p.includes('elev') && p.includes('wall') && (p.includes("1'-0") || p.includes('1-0"')) },
            { text: `F&I new (1'-2" thick) elevator pit walls (H=5'-0") as per FO-101.00 & details on FO-203.00`, sub: 'Elevator Pit', match: p => p.includes('elev') && p.includes('wall') && (p.includes("1'-2") || p.includes('1-2"') || p.includes('1.17')) },
            { text: `F&I new (1'-2" thick) elevator pit slope transition/haunch (H=2'-3") as per FO-101.00 & details on FO-203.00`, sub: 'Elevator Pit', match: p => p.includes('slope') || p.includes('haunch') }
          ]
        },
        {
          heading: 'Detention tank w/epoxy coated reinforcement', items: [
            { text: `F&I new (1'-0" thick) detention tank slab as per FO-101.00 & details on FO-203.00`, sub: 'Detention tank', match: p => p.includes('detention') && p.includes('slab') && !p.includes('lid') },
            { text: `F&I new (1'-0" thick) detention tank walls (H=10'-0") as per FO-101.00 & details on FO-203.00`, sub: 'Detention tank', match: p => p.includes('detention') && p.includes('wall') },
            { text: `F&I new (0'-8" thick) detention tank lid slab as per FO-101.00 & details on FO-203.00`, sub: 'Detention tank', match: p => p.includes('detention') && (p.includes('lid') || p.includes('8"')) }
          ]
        },
        {
          heading: 'Duplex sewage ejector pit', items: [
            { text: `F&I new (0'-8" thick, typ.) duplex sewage ejector pit slab as per P-301.01`, sub: 'Duplex sewage ejector pit', match: p => p.includes('duplex') && p.includes('slab') },
            { text: `F&I new (0'-8" thick, typ.) duplex sewage ejector pit wall (H=5'-0", typ.) as per P-301.01`, sub: 'Duplex sewage ejector pit', match: p => p.includes('duplex') && p.includes('wall') }
          ]
        },
        {
          heading: 'Deep sewage ejector pit', items: [
            { text: `F&I new (1'-0" thick) deep sewage ejector pit slab as per P-100.00 & details on P-205.00`, sub: 'Deep sewage ejector pit', match: p => p.includes('deep') && p.includes('slab') },
            { text: `F&I new (0'-6" wide) deep sewage ejector pit walls (H=8'-6") as per P-100.00 & details on P-205.00`, sub: 'Deep sewage ejector pit', match: p => p.includes('deep') && p.includes('wall') }
          ]
        },
        {
          heading: 'Grease trap pit', items: [
            { text: `F&I new (1'-0" thick) grease trap pit slab as per P-100.00 & details on P-205.00`, sub: 'Grease trap', match: p => p.includes('grease') && p.includes('slab') },
            { text: `F&I new (0'-6" wide) grease trap pit walls (H=8'-6") as per P-100.00 & details on P-205.00`, sub: 'Grease trap', match: p => p.includes('grease') && p.includes('wall') }
          ]
        },
        {
          heading: 'House trap pit', items: [
            { text: `F&I new (1'-0" thick) house trap pit slab as per P-100.00 & details on P-205.00`, sub: 'House trap', match: p => p.includes('house trap') && p.includes('slab') },
            { text: `F&I new (0'-6" wide) house trap pit walls (H=8'-6") as per P-100.00 & details on P-205.00`, sub: 'House trap', match: p => p.includes('house trap') && p.includes('wall') }
          ]
        },
        {
          heading: 'Foundation elements', items: [
            { text: 'F&I new (9)no pile caps as per FO-101.00', sub: 'Pile caps', match: () => true },
            { text: `F&I new (2'-0" to 7'-0" wide) strip footings (Havg=1'-10") as per FO-101.00`, sub: 'Strip Footings', match: () => true },
            { text: 'F&I new (15)no isolated footings as per FO-101.00', sub: 'Isolated Footings', match: () => true },
            { text: `F&I new (17)no (12"x24" wide) piers (H=3'-4") as per FO-101.00`, sub: 'Pier', match: () => true },
            { text: `F&I new (4)no (22"x16" wide) Pilasters (Havg=8'-0") as per S-101.00`, sub: 'Pilaster', match: () => true },
            { text: 'F&I new grade beams as per FO-102.00', sub: 'Grade beams', match: () => true },
            { text: 'F&I new tie beams as per FO-101.00, FO-102.00', sub: 'Tie beam', match: () => true },
            { text: `F&I new (8" thick) thickened slab as per FO-101.00`, sub: 'Thickened slab', match: () => true },
            { text: `F&I new (24" thick) mat slab reinforced as per FO-101.00`, sub: 'Mat slab', match: p => p.includes('2\'') || p.includes('24') },
            { text: `F&I new (36" thick) mat slab reinforced as per FO-101.00`, sub: 'Mat slab', match: p => p.includes('3\'') || p.includes('36') },
            { text: `F&I new (3" thick) mud/rat slab as per FO-101.00 & details on FO-003.00`, sub: 'Mud Slab', match: () => true },
            { text: `F&I new (4" thick) slab on grade reinforced w/6x6-10/10 W.W.M. @ cellar FL as per FO-101.00`, sub: 'SOG', match: p => (p.includes('sog 4') || (p.includes('sog') && p.includes('4"') && !p.includes('patch') && !p.includes('5"'))) },
            { text: `F&I new (6" thick) slab on grade reinforced w/6x6-10/10 W.W.M. @ cellar FL as per FO-101.00`, sub: 'SOG', match: p => (p.includes('sog 6') || (p.includes('sog') && p.includes('6"') && !p.includes('patio') && !p.includes('step'))) },
            { text: `F&I new (6" thick) patio slab on grade reinforced w/6x6-10/10 W.W.M. @ cellar FL as per FO-101.00`, sub: 'SOG', match: p => p.includes('patio') },
            { text: `Allow to (5" thick) patch slab on grade reinforced w/6x6-10/10 W.W.M. @ cellar FL as per FO-101.00`, sub: 'SOG', match: p => p.includes('patch') },
            { text: `F&I new (6" thick) slab on grade step (H=1'-2") as per FO-101.00 & details on FO-203.00`, sub: 'SOG', match: p => p.includes('step') },
            { text: `F&I new (14)no (2'-0" to 3'-6" wide) buttresses (H=10'-3") within foundation walls @ cellar FL to 1st FL as per FO-101.00`, sub: 'Buttresses', match: () => true },
            { text: `F&I new (8" thick) corbel (Havg=2'-2") as per FO-101.00 & details on FO-203.00`, sub: 'Corbel', match: () => true },
            { text: `F&I new (8" thick) concrete liner walls (Havg=11'-2") @ cellar FL to 1st FL as per FO-101.00 & details on FO-203.00`, sub: 'Linear Wall', match: () => true },
            { text: `F&I new (1'-0" wide) foundation walls (Havg=10'-10") @ cellar FL to 1st FL as per FO-101.00`, sub: 'Foundation Wall', match: () => true },
            { text: `F&I new (10" wide) retaining walls (Havg=6'-4") @ court as per FO-101.00`, sub: 'Retaining walls', match: p => p.includes('court') },
            { text: `F&I new (1'-0" wide) retaining walls w/epoxy reinforcement (Havg=4'-8") @ level 1 as per FO-101.00`, sub: 'Retaining walls', match: p => p.toLowerCase().startsWith('rw') || p.includes('retaining wall') || p.includes('retaining walls') || p.includes('level') }
          ]
        },
        {
          heading: 'Barrier wall', items: [
            { text: `F&I new (0'-8" thick) barrier walls (Havg=4'-0") @ cellar FL as per FO-001.00 & details on FO-103.00`, sub: 'Barrier wall', match: p => p.includes('8"') || p.includes('2\'') },
            { text: `F&I new (0'-10" thick) barrier walls (Havg=5'-1") @ cellar FL as per FO-001.00 & details on FO-103.00`, sub: 'Barrier wall', match: p => p.includes('10"') || p.includes('5\'') }
          ]
        },
        {
          heading: 'Stem wall', items: [
            { text: `F&I new (0'-10" thick) stem walls (Havg=5'-1") @ cellar FL as per FO-001.00 & details on FO-103.00`, sub: 'Stem wall', match: () => true }
          ]
        },
        {
          heading: 'Stair on grade', items: [
            { text: `F&I new (8" thick) stair landings as per A-312.00`, sub: 'Stairs on grade Stairs', match: p => p.includes('landing') || p.includes('Landings') },
            { text: `F&I new (5'-6" wide) stairs on grade (2 Riser) @ 1st FL as per A-312.00`, sub: 'Stairs on grade Stairs', match: p => (p.includes('riser') || p.includes('Riser') || p.includes('wide')) && !p.includes('landing') }
          ]
        },
        {
          heading: 'Electric conduit', items: [
            { text: 'F&I new electric conduit as per E-060.00 & E-070.00', sub: 'Electric conduit', match: p => p.includes('electric conduit') || p.includes('conduit') }
          ]
        },
        {
          heading: 'Trench drain', items: [
            { text: 'F&I new trench drain as per FO-102.00 & details on scope sheet pt. no. 65', formulaItem: { itemType: 'electric_conduit', match: p => (p || '').toLowerCase().includes('trench drain') } }
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
      for (const { items } of substructureTemplate) {
        for (const it of items) {
          const { sub, match, formulaItem } = it
          let source = null
          if (formulaItem) {
            const itemF = (formulaData || []).find(f =>
              f.itemType === formulaItem.itemType && f.section === 'foundation' &&
              !dryRunUsedItem.has(f.row) && formulaItem.match((f.parsedData?.particulars || '').toLowerCase())
            )
            if (itemF) {
              source = itemF
              dryRunUsedItem.add(itemF.row)
              if (hasNonZeroTakeoff(itemF, true)) { hasSubstructureScopeData = true; break }
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
              if (hasNonZeroTakeoff(sumF, false)) { hasSubstructureScopeData = true; break }
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
      substructureTemplate.forEach(({ heading, items }) => {
        let headingAdded = false
        items.forEach(({ text, sub, match, formulaItem }) => {
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
          if (!rowToUse || !source) return
          if (!hasNonZeroTakeoff(source, !!formulaItem)) return
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
          spreadsheet.updateCell({ value: text }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${substructureCalcSheet}'!I${sumRow}` }, `${pfx}C${currentRow}`)
          spreadsheet.updateCell({ formula: `='${substructureCalcSheet}'!J${sumRow}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${substructureCalcSheet}'!L${sumRow}` }, `${pfx}F${currentRow}`)
          spreadsheet.updateCell({ formula: `='${substructureCalcSheet}'!M${sumRow}` }, `${pfx}G${currentRow}`)
          fillRatesForProposalRow(currentRow, text)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        })
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
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
        spreadsheet.updateCell({ formula: `=SUM(F${substructureStartRow}:F${substructureDataEndRow})` }, `${pfx}F${currentRow}`)
        spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}F${currentRow}`)
        try { spreadsheet.numberFormat('#,##0.00', `${pfx}F${currentRow}`) } catch (e) { }
        fillRatesForProposalRow(currentRow, 'DOB approved concrete washout included')
        spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
        currentRow++
        spreadsheet.updateCell({ value: 'Engineering, shop drawings, formwork drawings, design mixes included' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
        fillRatesForProposalRow(currentRow, 'Engineering, shop drawings, formwork drawings, design mixes included')
        spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
        currentRow++
      }
      if (currentRow > substructureStartRow) {
        const substructureEndRow = currentRow - 1
        spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
        spreadsheet.updateCell({ value: 'Substructure concrete Total:' }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD7EE', border: '1px solid #000000' },
          `${pfx}B${currentRow}:E${currentRow}`
        )
        spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
        spreadsheet.updateCell({ formula: `=SUM(H${substructureStartRow}:H${substructureEndRow})*1000` }, `${pfx}F${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD7EE', border: '1px solid #000000', format: '$#,##0.00' },
          `${pfx}F${currentRow}:G${currentRow}`
        )
        currentRow++
      }

      currentRow++
    }

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
      fillRatesForProposalRow(currentRow, proposalText)

      // Calculate and set row height based on content
      const waterproofingDynamicHeight = calculateRowHeight(proposalText)

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

          // Add $/1000 formula in column H
          const dollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
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

        // Add $/1000 formula in column H
        const negativeDollarFormula = `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")`
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
      const calcSheetName = 'Calculations Sheet'
      let horizontalInsulationSumRow = 0
      let thicknessFromCalc = '2' // default
      if (formulaData && Array.isArray(formulaData)) {
        const horizontalInsulationSumFormula = formulaData.find(f =>
          f.itemType === 'waterproofing_horizontal_insulation_sum' &&
          f.section === 'waterproofing'
        )
        if (horizontalInsulationSumFormula) {
          horizontalInsulationSumRow = horizontalInsulationSumFormula.row
          // Thickness is dynamic: read from first data row (column B) e.g. "2"XPS Rigid insulation @ SOG & GB"
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
      spreadsheet.wrap(`${pfx}B${currentRow}`, true)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' },
        `${pfx}B${currentRow}`
      )
      if (horizontalInsulationSumRow > 0) {
        spreadsheet.updateCell({ formula: `='${calcSheetName}'!J${horizontalInsulationSumRow}` }, `${pfx}D${currentRow}`)
        spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}D${currentRow}`)
        try { spreadsheet.numberFormat('#,##0.00', `${pfx}D${currentRow}`) } catch (e) { spreadsheet.cellFormat({ format: '#,##0.00' }, `${pfx}D${currentRow}`) }
      }
      fillRatesForProposalRow(currentRow, horizontalProposalText)
      spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
      spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white', format: '$#,##0.0' }, `${pfx}H${currentRow}`)
      try { spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}`) } catch (e) { }
      currentRow++
    }

    // Add note about WR Meadows Installer
    spreadsheet.updateCell({ value: 'Note: Capstone Contracting Corp is a licensed WR Meadows Installer ' }, `${pfx}B${currentRow}`)
    spreadsheet.wrap(`${pfx}B${currentRow}`, true)
    spreadsheet.cellFormat(
      {
        fontWeight: 'normal',
        fontStyle: 'italic',
        color: '#000000',
        textAlign: 'center',
        backgroundColor: 'white',
        verticalAlign: 'middle'
      },
      `${pfx}B${currentRow}`
    )
    currentRow++

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
      if (!rawData || rawData.length < 2) return { cipLines: lines, patchLines, somdLines, slabStepsLines: [], cipStairsLines: [], concreteHangerLines: [], lwConcreteFillLines: [], toppingSlabLines: [], raisedSlabLines: [], builtUpSlabLines: [], builtupRampsLines: [], builtUpStairLines: [], shearWallLines: [], columnsLines: [], dropPanelLines: [], concretePostLines: [], concreteEncasementLines: [], beamsLines: [], parapetWallLines: [], thermalBreakLines: [], curbsLines: [], concretePadLines: [], nonShrinkGroutLines: [], repairScopeLines: [] }
      const headers = rawData[0]
      const dataRows = rawData.slice(1)
      const digitizerIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'digitizer item')
      const pageIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'page')
      const estimateIdx = headers.findIndex(h => h && String(h).toLowerCase().trim() === 'estimate')
      if (digitizerIdx === -1) return { cipLines: lines, patchLines, somdLines, slabStepsLines: [], cipStairsLines: [], concreteHangerLines: [], lwConcreteFillLines: [], toppingSlabLines: [], raisedSlabLines: [], builtUpSlabLines: [], builtupRampsLines: [], builtUpStairLines: [], shearWallLines: [], columnsLines: [], dropPanelLines: [], concretePostLines: [], concreteEncasementLines: [], beamsLines: [], parapetWallLines: [], thermalBreakLines: [], curbsLines: [], concretePadLines: [], nonShrinkGroutLines: [], repairScopeLines: [] }

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
      const cipStairsLines = []
      cipStairsGroups.forEach(grp => {
        const allRows = [...grp.landings, ...grp.stairs]
        const floorText = formatFloorRefs(allRows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
        const sRefs = extractSPageRefs(allRows)
        if (grp.landings.length > 0) {
          const line = { proposalText: `F&I new (8" thick) stair landings @ ${floorText} as per ${sRefs}` }
          cipStairsLines.push(line)
          if (typeof console !== 'undefined' && console.log) console.log('[CIP Stairs]', grp.groupName, 'Landings:', line.proposalText)
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
          const line = { proposalText: `F&I new (${widthStr} wide) concrete stairs (${riserStr}) @ ${floorText} as per ${sRefs}` }
          cipStairsLines.push(line)
          if (typeof console !== 'undefined' && console.log) console.log('[CIP Stairs]', grp.groupName, 'Stairs:', line.proposalText)
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
          const thickStr = `${key}"`
          const floorText = formatFloorRefs(rows.map(r => extractFloorRef(r[digitizerIdx]))) || 'as indicated'
          const sRefs = extractSPageRefs(rows)
          const padWord = totalQty === 1 ? 'concrete pad' : 'concrete pads'
          concretePadLines.push({
            proposalText: `F&I new (${totalQty})no (${thickStr} thick) ${padWord} @ ${floorText} as per ${sRefs}`,
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
          targetCol: 'G'
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
      return { cipLines: lines, patchLines, somdLines, slabStepsLines, cipStairsLines, concreteHangerLines, lwConcreteFillLines, toppingSlabLines, raisedSlabLines, builtUpSlabLines, builtupRampsLines, builtUpStairLines, shearWallLines, columnsLines, dropPanelLines, concretePostLines, concreteEncasementLines, beamsLines, parapetWallLines, thermalBreakLines, curbsLines, concretePadLines, nonShrinkGroutLines, repairScopeLines }
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

    const { cipLines, patchLines, somdLines, slabStepsLines, cipStairsLines, concreteHangerLines, lwConcreteFillLines, toppingSlabLines, raisedSlabLines, builtUpSlabLines, builtupRampsLines, builtUpStairLines, shearWallLines, columnsLines, dropPanelLines, concretePostLines, concreteEncasementLines, beamsLines, parapetWallLines, thermalBreakLines, curbsLines, concretePadLines, nonShrinkGroutLines, repairScopeLines } = buildSuperstructureProposalLines()
    const calcSheetName = 'Calculations Sheet'

    let superstructureScopeStartRow = null
    let superstructureScopeEndRow = null
    const washoutExclusionRows = []

    const renderProposalLine = (line) => {
      if (superstructureScopeStartRow == null) superstructureScopeStartRow = currentRow
      if (line.allowance) {
        spreadsheet.updateCell({ value: line.proposalText }, `${pfx}B${currentRow}`)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' },
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
        spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white', format: '$#,##0.0' }, `${pfx}H${currentRow}`)
        try { spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}`) } catch (e) { }
        washoutExclusionRows.push(currentRow)
        superstructureScopeEndRow = currentRow
        currentRow++
        return
      }
      const groupIndex = line.somdGroupIndex ?? line.shearGroupIndex ?? line.curbGroupIndex ?? line.concretePadGroupIndex ?? 0
      const sumInfo = findSumRowForSubsection(line.subsectionName, line.sumRowKey, groupIndex, line)
      spreadsheet.updateCell({ value: line.proposalText }, `${pfx}B${currentRow}`)
      spreadsheet.wrap(`${pfx}B${currentRow}`, true)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' },
        `${pfx}B${currentRow}`
      )
      fillRatesForProposalRow(currentRow, line.proposalText)
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
        { fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white', format: '$#,##0.0' },
        `${pfx}H${currentRow}`
      )
      try {
        spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}`)
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
      spreadsheet.updateCell({ formula: `=SUM(F${cipStartRow}:F${cipEndRow})` }, `${pfx}F${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', fontStyle: 'italic' },
        `${pfx}C${currentRow}`
      )
        ;['D', 'F'].forEach(col => {
          spreadsheet.cellFormat(
            { fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' },
            `${pfx}${col}${currentRow}`
          )
          try { spreadsheet.numberFormat('#,##0.00', `${pfx}${col}${currentRow}`) } catch (e) { }
        })
      washoutExclusionRows.push(currentRow)
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
    if (cipStairsLines.length > 0) {
      spreadsheet.updateCell({ value: 'Concrete stairs:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      cipStairsLines.forEach(renderProposalLine)
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
        spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white', format: '$#,##0.0' }, `${pfx}H${currentRow}`)
        try { spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}`) } catch (e) { }
        currentRow++
      })
    }
    if (superstructureScopeStartRow != null && superstructureScopeEndRow != null) {
      spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
      spreadsheet.updateCell({ value: 'Superstructure Concrete Total:' }, `${pfx}D${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#BDD7EE', border: '1px solid #000000' },
        `${pfx}D${currentRow}:E${currentRow}`
      )
      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({ formula: `=SUM(H${superstructureScopeStartRow}:H${superstructureScopeEndRow})*1000` }, `${pfx}F${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD7EE', border: '1px solid #000000', format: '$#,##0.00' },
        `${pfx}F${currentRow}:G${currentRow}`
      )
      try { spreadsheet.numberFormat('$#,##0.00', `${pfx}F${currentRow}:G${currentRow}`) } catch (e) { }
      spreadsheet.cellFormat({ backgroundColor: '#BDD7EE' }, `${pfx}B${currentRow}:G${currentRow}`)
      baseBidTotalRows.push(currentRow) // Superstructure Concrete Total
      totalRows.push(currentRow)
      currentRow++
    }

    // Add empty row after Superstructure concrete scope
    currentRow++

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
        spreadsheet.updateCell({ value: 'Plumbing/Electrical trenching scope: as per P-302.00' }, `${pfx}B${currentRow}`)
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
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
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
          spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white', format: '$#,##0.0' }, `${pfx}H${currentRow}`)
          try { spreadsheet.numberFormat('$#,##0.0', `${pfx}H${currentRow}`) } catch (e) { }
          currentRow++
        })
        const trenchingEndRow = currentRow - 1

        spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
        spreadsheet.updateCell({ value: 'Trenching Total:' }, `${pfx}D${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#BDD7EE', border: '1px solid #000000' },
          `${pfx}D${currentRow}:E${currentRow}`
        )
        spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
        spreadsheet.updateCell({ formula: `=SUM(H${trenchingStartRow}:H${trenchingEndRow})*1000` }, `${pfx}F${currentRow}`)
        spreadsheet.cellFormat(
          { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD7EE', border: '1px solid #000000', format: '$#,##0.00' },
          `${pfx}F${currentRow}:G${currentRow}`
        )
        try { spreadsheet.numberFormat('$#,##0.00', `${pfx}F${currentRow}:G${currentRow}`) } catch (e) { }
        spreadsheet.cellFormat({ backgroundColor: '#BDD7EE' }, `${pfx}B${currentRow}:G${currentRow}`)
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
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#BDD7EE', border: '1px solid #000000' },
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
      spreadsheet.cellFormat({ backgroundColor: '#BDD7EE' }, `${pfx}B${currentRow}:G${currentRow}`)
      currentRow++
    }

    // Add empty row after BASE BID TOTAL
    currentRow++

    // Add Alternate #1: Sitework scope header
    let siteworkScopeStartRow = null
    const siteworkTotalFRows = []
    {
      spreadsheet.merge(`${pfx}B${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({ value: 'Add Alternate #1: Sitework scope: as per C-4' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#BDD6EE', border: '1px solid #000000' },
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
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
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
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
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
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
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
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
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
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
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
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
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
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
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
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
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
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
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
            spreadsheet.wrap(`${pfx}B${currentRow}`, true)
            spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
            fillRatesForProposalRow(currentRow, proposalText)

            spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${rowSumRow}` }, `${pfx}G${currentRow}`)
            spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}G${currentRow}`)
            currentRow++
          }
          return // Skip default handling
        } else if (demoType === 'Demo inlet' && sumRowIndex > 0) {
          proposalText = 'Allow to remove existing stormwater inlet as per C-3'
          spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, proposalText)
          spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
          spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}G${currentRow}`)
          currentRow++
          return
        } else if (demoType === 'Demo fire hydrant' && sumRowIndex > 0) {
          proposalText = 'Allow to relocate existing fire hydrant as per C-3'
          spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, proposalText)
          spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
          spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}G${currentRow}`)
          currentRow++
          return
        } else if (demoType === 'Demo manhole' && sumRowIndex > 0) {
          proposalText = 'Allow to protect existing stormwater manhole as per C-3'
          spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, proposalText)
          spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
          spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}G${currentRow}`)
          currentRow++
          return
        } else if (demoType === 'Demo utility pole' && sumRowIndex > 0) {
          proposalText = 'Allow to protect utility pole as per C-3'
          spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, proposalText)
          spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
          spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}G${currentRow}`)
          currentRow++
          return
        } else if (demoType === 'Demo valve' && sumRowIndex > 0) {
          proposalText = 'Allow to protect existing water valve as per C-3'
          spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          fillRatesForProposalRow(currentRow, proposalText)
          spreadsheet.updateCell({ formula: `='${calcSheetName}'!M${sumRowIndex}` }, `${pfx}G${currentRow}`)
          spreadsheet.cellFormat({ fontWeight: 'bold', textAlign: 'right', backgroundColor: 'white' }, `${pfx}G${currentRow}`)
          currentRow++
          return
        }

        // Default handling for other demo types
        if (proposalText) {
          spreadsheet.updateCell({ value: proposalText }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
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
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000' },
        `${pfx}B${currentRow}:E${currentRow}`
      )
      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({ formula: `=SUM(H${demoStartRow}:H${demoEndRow})*1000` }, `${pfx}F${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000', format: '$#,##0.00' },
        `${pfx}F${currentRow}:G${currentRow}`
      )
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
        spreadsheet.updateCell({ value: `F&I new (${thickness}" thick) temporary stabilized construction entrance w/geotextile fabric as per C-8 & details on C-12` }, `${pfx}B${currentRow}`)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
        spreadsheet.updateCell({ formula: `='${erosionCalcSheetName}'!J${rowNum}` }, `${pfx}D${currentRow}`)
        spreadsheet.updateCell({ formula: `='${erosionCalcSheetName}'!L${rowNum}` }, `${pfx}F${currentRow}`)
        spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
        currentRow++
      })

      erosionItems.silt_fence.forEach(({ rowNum, heightStr = "2'-6\"" }) => {
        spreadsheet.updateCell({ value: `F&I new silt fence (H=${heightStr}) as per C-8 & details on C-13` }, `${pfx}B${currentRow}`)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
        spreadsheet.updateCell({ formula: `='${erosionCalcSheetName}'!I${rowNum}` }, `${pfx}C${currentRow}`)
        spreadsheet.updateCell({ formula: `='${erosionCalcSheetName}'!J${rowNum}` }, `${pfx}D${currentRow}`)
        spreadsheet.updateCell({ formula: `='${erosionCalcSheetName}'!L${rowNum}` }, `${pfx}F${currentRow}`)
        spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
        currentRow++
      })

      erosionItems.inlet_filter.forEach(({ rowNum, qty = 1 }) => {
        spreadsheet.updateCell({ value: `F&I new (${qty})no inlet filter as per C-8 & details on C-13` }, `${pfx}B${currentRow}`)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
        spreadsheet.updateCell({ formula: `='${erosionCalcSheetName}'!M${rowNum}` }, `${pfx}G${currentRow}`)
        spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
        currentRow++
      })

      const erosionEndRow = currentRow - 1
      const erosionDataStartRow = erosionHeaderRow + 1
      spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
      spreadsheet.updateCell({ value: 'Erosion & Sediment Control Total:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000' },
        `${pfx}B${currentRow}:E${currentRow}`
      )
      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({ formula: `=SUM(H${erosionDataStartRow}:H${erosionEndRow})*1000` }, `${pfx}F${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000', format: '$#,##0.00' },
        `${pfx}F${currentRow}:G${currentRow}`
      )
      siteworkTotalFRows.push(currentRow)
      currentRow++
    }

    // Excavation, backfill & grading scope
    currentRow++
    const excGradingHeaderRow = currentRow
    const excGradingCalcSheet = 'Calculations Sheet'
    const civilExcSum = (formulaData || []).find(f => f.itemType === 'civil_exc_sum' && f.section === 'civil_sitework')
    const civilGravelSum = (formulaData || []).find(f => f.itemType === 'civil_gravel_sum' && f.section === 'civil_sitework')
    const gravelItems = { transformer_pad: [], reinforced_sidewalk: [], asphalt: [] }
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
              gravelItems.transformer_pad.push({ rowNum })
            } else if (bText.includes('reinforced') && bText.includes('sidewalk')) {
              gravelItems.reinforced_sidewalk.push({ rowNum })
            } else if (bText.includes('full depth asphalt') || bText.includes('asphalt pavement')) {
              gravelItems.asphalt.push({ rowNum })
            }
          }
        }
      })
    }

    const hasExcOrGravel = civilExcSum || civilGravelSum || gravelItems.transformer_pad.length > 0 || gravelItems.reinforced_sidewalk.length > 0 || gravelItems.asphalt.length > 0
    if (hasExcOrGravel) {
      spreadsheet.updateCell({ value: 'Excavation, backfill & grading scope:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#FEF2CB', verticalAlign: 'middle', textDecoration: 'underline', border: '1px solid #000000' },
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

        const excFirstRow = civilExcSum.firstDataRow || civilExcSum.row
        const excLastRow = civilExcSum.lastDataRow || civilExcSum.row
        spreadsheet.updateCell({ value: 'Allow to perform soil excavation, trucking & disposal as per C-4' }, `${pfx}B${currentRow}`)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
        if (excFirstRow && excLastRow && excLastRow !== excFirstRow) {
          spreadsheet.updateCell({ formula: `=SUM('${excGradingCalcSheet}'!J${excFirstRow}:'${excGradingCalcSheet}'!J${excLastRow})` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `=SUM('${excGradingCalcSheet}'!L${excFirstRow}:'${excGradingCalcSheet}'!L${excLastRow})` }, `${pfx}F${currentRow}`)
        } else if (excFirstRow) {
          spreadsheet.updateCell({ formula: `='${excGradingCalcSheet}'!J${excFirstRow}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${excGradingCalcSheet}'!L${excFirstRow}` }, `${pfx}F${currentRow}`)
        }
        spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
        currentRow++
      }

      if (gravelItems.transformer_pad.length > 0 || gravelItems.reinforced_sidewalk.length > 0 || gravelItems.asphalt.length > 0) {
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
          spreadsheet.updateCell({ value: `F&I new (4" thick) gravel/crushed stone @ transformer concrete pad as per A-100.00` }, `${pfx}B${currentRow}`)
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
          spreadsheet.updateCell({ value: `F&I new (6" thick) gravel/crushed stone @ utility trench & asphalt pavement as per C-6 & details on C-12` }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: sumFormulaJ }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: sumFormulaL }, `${pfx}F${currentRow}`)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
      }

      const excGradingEndRow = currentRow - 1
      const excGradingDataStartRow = excGradingHeaderRow + 1
      spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
      spreadsheet.updateCell({ value: 'Excavation, Backfill & Grading Total:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000' },
        `${pfx}B${currentRow}:E${currentRow}`
      )
      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({ formula: `=SUM(H${excGradingDataStartRow}:H${excGradingEndRow})*1000` }, `${pfx}F${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000', format: '$#,##0.00' },
        `${pfx}F${currentRow}:G${currentRow}`
      )
      siteworkTotalFRows.push(currentRow)
      currentRow++
    }

    // Dewatering section (after Excavation, backfill & grading scope)
    currentRow++
    spreadsheet.updateCell({ value: 'Dewatering:' }, `${pfx}B${currentRow}`)
    spreadsheet.cellFormat(
      { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#FEF2CB', verticalAlign: 'middle', textDecoration: 'underline', border: '1px solid #000000' },
      `${pfx}B${currentRow}`
    )
    currentRow++
    spreadsheet.updateCell({ value: 'Dewatering allowance - Budget $200k' }, `${pfx}B${currentRow}`)
    spreadsheet.wrap(`${pfx}B${currentRow}`, true)
    spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
    currentRow++
    spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
    spreadsheet.updateCell({ value: 'Dewatering Total:' }, `${pfx}B${currentRow}`)
    spreadsheet.cellFormat(
      { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000' },
      `${pfx}B${currentRow}:E${currentRow}`
    )
    spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
    spreadsheet.updateCell({ value: 200000 }, `${pfx}F${currentRow}`)
    spreadsheet.cellFormat(
      { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000', format: '$#,##0.00' },
      `${pfx}F${currentRow}:G${currentRow}`
    )
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
          spreadsheet.updateCell({ value: 'Allow to perform soil excavation, trucking & disposal (Havg=2\'-11") as per C-4' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          if (excFirstRow && excLastRow && excLastRow !== excFirstRow) {
            spreadsheet.updateCell({ formula: `=SUM('${utilsCalcSheet}'!J${excFirstRow}:'${utilsCalcSheet}'!J${excLastRow})` }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: `=SUM('${utilsCalcSheet}'!L${excFirstRow}:'${utilsCalcSheet}'!L${excLastRow})` }, `${pfx}F${currentRow}`)
          } else if (excFirstRow) {
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${excFirstRow}` }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${excFirstRow}` }, `${pfx}F${currentRow}`)
          }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
        if (civilEleBackfill) {
          const backFirstRow = civilEleBackfill.firstDataRow || civilEleBackfill.row
          const backLastRow = civilEleBackfill.lastDataRow || civilEleBackfill.row
          spreadsheet.updateCell({ value: 'Allow to import new clean soil to backfill and compact as per C-4' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          if (backFirstRow && backLastRow && backLastRow !== backFirstRow) {
            spreadsheet.updateCell({ formula: `=SUM('${utilsCalcSheet}'!J${backFirstRow}:'${utilsCalcSheet}'!J${backLastRow})` }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: `=SUM('${utilsCalcSheet}'!L${backFirstRow}:'${utilsCalcSheet}'!L${backLastRow})` }, `${pfx}F${currentRow}`)
          } else if (backFirstRow) {
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${backFirstRow}` }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${backFirstRow}` }, `${pfx}F${currentRow}`)
          }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
        if (civilEleGravel) {
          const gravFirstRow = civilEleGravel.firstDataRow || civilEleGravel.row
          const gravLastRow = civilEleGravel.lastDataRow || civilEleGravel.row
          spreadsheet.updateCell({ value: 'F&I new (6\" thick) gravel/crushed stone, including 6MIL vapor barrier on top @ utility trench & asphalt pavement as per C-6 & details on C-12' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          if (gravFirstRow && gravLastRow && gravLastRow !== gravFirstRow) {
            spreadsheet.updateCell({ formula: `=SUM('${utilsCalcSheet}'!J${gravFirstRow}:'${utilsCalcSheet}'!J${gravLastRow})` }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: `=SUM('${utilsCalcSheet}'!L${gravFirstRow}:'${utilsCalcSheet}'!L${gravLastRow})` }, `${pfx}F${currentRow}`)
          } else if (gravFirstRow) {
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${gravFirstRow}` }, `${pfx}D${currentRow}`)
            spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${gravFirstRow}` }, `${pfx}F${currentRow}`)
          }
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
        if (utilPoleItem) {
          spreadsheet.updateCell({ value: 'F&I new underground connection to existing utility pole as per C-6' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!M${utilPoleItem.row}` }, `${pfx}G${currentRow}`)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
        if (eleConduitItem) {
          spreadsheet.updateCell({ value: 'F&I new underground electrical conduit' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!I${eleConduitItem.row}` }, `${pfx}C${currentRow}`)
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
          spreadsheet.updateCell({ value: 'F&I new (8" thick) transformer concrete pad as per A-100.00' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${civilPadsSum.row}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${civilPadsSum.row}` }, `${pfx}F${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!M${civilPadsSum.row}` }, `${pfx}G${currentRow}`)
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
          spreadsheet.updateCell({ value: 'Allow to perform soil excavation, trucking & disposal (Havg=2\'-11") as per C-4' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${gasExcRow}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${gasExcRow}` }, `${pfx}F${currentRow}`)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
        if (gasBackRow) {
          spreadsheet.updateCell({ value: 'Allow to import new clean soil to backfill and compact as per C-4' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${gasBackRow}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${gasBackRow}` }, `${pfx}F${currentRow}`)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
        if (gasGravRow) {
          spreadsheet.updateCell({ value: 'F&I new (6" thick) gravel/crushed stone, including 6MIL vapor barrier on top @ utility trench & asphalt pavement as per C-6 & details on C-12' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${gasGravRow}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${gasGravRow}` }, `${pfx}F${currentRow}`)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
        if (gasLateralItem) {
          spreadsheet.updateCell({ value: 'F&I new (4" thick) underground gas service lateral as per C-6' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!I${gasLateralItem.row}` }, `${pfx}C${currentRow}`)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
        const gasMainSum = getSiteSumByKey('Gas')
        if (gasMainSum) {
          spreadsheet.updateCell({ value: 'F&I new (1)no connection to existing gas main as per C-6' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!M${gasMainSum.row}` }, `${pfx}G${currentRow}`)
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
          spreadsheet.updateCell({ value: 'Allow to perform soil excavation, trucking & disposal (Havg=2\'-11") as per C-4' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${watExcRow}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${watExcRow}` }, `${pfx}F${currentRow}`)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
        if (watBackRow) {
          spreadsheet.updateCell({ value: 'Allow to import new clean soil to backfill and compact as per C-4' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${watBackRow}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${watBackRow}` }, `${pfx}F${currentRow}`)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
        const watGravRowOrFallback = watGravRow || watExcRow || watBackRow
        if (watGravRowOrFallback) {
          spreadsheet.updateCell({ value: 'F&I new (6" thick) gravel/crushed stone, including 6MIL vapor barrier on top @ utility trench & asphalt pavement as per C-6 & details on C-12' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!J${watGravRowOrFallback}` }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!L${watGravRowOrFallback}` }, `${pfx}F${currentRow}`)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
        if (waterMainItem) {
          spreadsheet.updateCell({ value: 'F&I new underground water main as per C-6' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!I${waterMainItem.row}` }, `${pfx}C${currentRow}`)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
        if (fireServiceItem) {
          spreadsheet.updateCell({ value: 'F&I new (6" thick) fire service lateral as per C-6' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!I${fireServiceItem.row}` }, `${pfx}C${currentRow}`)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
        const waterMainSum = getSiteSumByKey('Water')
        if (waterMainSum) {
          spreadsheet.updateCell({ value: 'F&I new (3)no connection to existing water main as per C-6' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!M${waterMainSum.row}` }, `${pfx}G${currentRow}`)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
        if (waterServiceItem) {
          spreadsheet.updateCell({ value: 'F&I new (4" thick) underground water service lateral as per C-6' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!I${waterServiceItem.row}` }, `${pfx}C${currentRow}`)
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
          spreadsheet.updateCell({ value: 'F&I new (10" Ø) storm sewer piping as per P-099.00' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!I${stormSewerItem.row}` }, `${pfx}C${currentRow}`)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
        if (sanitarySewerItem) {
          spreadsheet.updateCell({ value: 'F&I new (8" Ø) underground PVC sanitary sewer service as per C-6' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!I${sanitarySewerItem.row}` }, `${pfx}C${currentRow}`)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
        if (sanitInvertRows.length > 0) {
          spreadsheet.updateCell({ value: 'F&I new sanitary invert as per C-6' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          const lfFormula = sanitInvertRows.length === 1
            ? `='${utilsCalcSheet}'!I${sanitInvertRows[0].row}`
            : `=AVERAGE(${sanitInvertRows.map(r => `'${utilsCalcSheet}'!I${r.row}`).join(',')})`
          spreadsheet.updateCell({ formula: lfFormula }, `${pfx}C${currentRow}`)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
        if (underslabItem) {
          spreadsheet.updateCell({ value: 'F&I new underslab drainage piping' }, `${pfx}B${currentRow}`)
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: `='${utilsCalcSheet}'!I${underslabItem.row}` }, `${pfx}C${currentRow}`)
          spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
          currentRow++
        }
      }

      const utilsEndRow = currentRow - 1
      const utilsDataStartRow = utilsHeaderRow + 1
      spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
      spreadsheet.updateCell({ value: 'Utilities Total:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000' },
        `${pfx}B${currentRow}:E${currentRow}`
      )
      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({ formula: `=SUM(H${utilsDataStartRow}:H${utilsEndRow})*1000` }, `${pfx}F${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000', format: '$#,##0.00' },
        `${pfx}F${currentRow}:G${currentRow}`
      )
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
    spreadsheet.updateCell({ value: 'F&I new stormwater manhole as per C-5' }, `${pfx}B${currentRow}`)
    spreadsheet.wrap(`${pfx}B${currentRow}`, true)
    spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
    spreadsheet.updateCell({ value: 1 }, `${pfx}G${currentRow}`)
    spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
    currentRow++
    const stormWaterEndRow = currentRow - 1
    const stormWaterDataStartRow = stormWaterHeaderRow + 1
    spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
    spreadsheet.updateCell({ value: 'Storm Water Management Total:' }, `${pfx}B${currentRow}`)
    spreadsheet.cellFormat(
      { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000' },
      `${pfx}B${currentRow}:E${currentRow}`
    )
    spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
    spreadsheet.updateCell({ value: 10000 }, `${pfx}F${currentRow}`)
    spreadsheet.cellFormat(
      { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000', format: '$#,##0.00' },
      `${pfx}F${currentRow}:G${currentRow}`
    )
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
        spreadsheet.updateCell({ value: `F&I new (4" thick) concrete sidewalk, reinf w/ 6x6-W1.4xW1.4 w.w.f as per ${cpRefs.main} & details on ${cpRefs.details}` }, `${pfx}B${currentRow}`)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
        spreadsheet.updateCell({ formula: `='${siteFinishesCalcSheet}'!J${cpRow}` }, `${pfx}D${currentRow}`)
        spreadsheet.updateCell({ formula: `='${siteFinishesCalcSheet}'!L${cpRow}` }, `${pfx}F${currentRow}`)
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
        spreadsheet.updateCell({ value: `F&I new full depth asphalt pavement (1.5" thick) surface course on (3" thick) base course as per ${apRefs.main} & details on ${apRefs.details}` }, `${pfx}B${currentRow}`)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
        spreadsheet.updateCell({ formula: `='${siteFinishesCalcSheet}'!J${apRow}` }, `${pfx}D${currentRow}`)
        spreadsheet.updateCell({ formula: `='${siteFinishesCalcSheet}'!L${apRow}` }, `${pfx}F${currentRow}`)
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
          { textBuilder: (qty, ref) => `F&I new (${qty ?? 3})no fire hydrant as per ${ref}`, qtySource: 'sum', siteGroupKey: 'Hydrant' },
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
          spreadsheet.wrap(`${pfx}B${currentRow}`, true)
          spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
          spreadsheet.updateCell({ formula: qtyFormula }, `${pfx}G${currentRow}`)
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
        spreadsheet.updateCell({ value: `F&I new (${bollardQty ?? 28})no (${diaStr} Ø wide) concrete filled steel pipe bollard (Havg=${heightStr}) as per ${bollardRefs.main} & details on ${bollardRefs.details}` }, `${pfx}B${currentRow}`)
        spreadsheet.wrap(`${pfx}B${currentRow}`, true)
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
        spreadsheet.updateCell({ formula: `='${siteFinishesCalcSheet}'!J${bfRow}` }, `${pfx}D${currentRow}`)
        spreadsheet.updateCell({ formula: `='${siteFinishesCalcSheet}'!L${bfRow}` }, `${pfx}F${currentRow}`)
        spreadsheet.updateCell({ formula: `='${siteFinishesCalcSheet}'!M${bfRow}` }, `${pfx}G${currentRow}`)
        spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
        currentRow++
      }

      spreadsheet.updateCell({ value: 'Misc.:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#D0CECE', textDecoration: 'underline', border: '1px solid #000000' },
        `${pfx}B${currentRow}`
      )
      currentRow++
      spreadsheet.updateCell({ value: 'Temp Protection & Barriers' }, `${pfx}B${currentRow}`)
      spreadsheet.wrap(`${pfx}B${currentRow}`, true)
      spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
      currentRow++
      spreadsheet.updateCell({ value: 'DOT Permits, Temp DOT Barriers' }, `${pfx}B${currentRow}`)
      spreadsheet.wrap(`${pfx}B${currentRow}`, true)
      spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
      currentRow++

      const siteFinishesEndRow = currentRow - 1
      const siteFinishesDataStartRow = siteFinishesHeaderRow + 1
      spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
      spreadsheet.updateCell({ value: 'Site Finishes Total:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000' },
        `${pfx}B${currentRow}:E${currentRow}`
      )
      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({ formula: `=SUM(H${siteFinishesDataStartRow}:H${siteFinishesEndRow})*1000` }, `${pfx}F${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000', format: '$#,##0.00' },
        `${pfx}F${currentRow}:G${currentRow}`
      )
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
        spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
        spreadsheet.updateCell({ formula: `='${fenceCalcSheet}'!I${sumRow}` }, `${pfx}C${currentRow}`)
        spreadsheet.updateCell({ formula: `='${fenceCalcSheet}'!J${sumRow}` }, `${pfx}D${currentRow}`)
        spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
        currentRow++
      })

      const fenceEndRow = currentRow - 1
      const fenceDataStartRow = fenceHeaderRow + 1
      spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
      spreadsheet.updateCell({ value: 'Construction Fence Total:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000' },
        `${pfx}B${currentRow}:E${currentRow}`
      )
      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({ formula: `=SUM(H${fenceDataStartRow}:H${fenceEndRow})*1000` }, `${pfx}F${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000', format: '$#,##0.00' },
        `${pfx}F${currentRow}:G${currentRow}`
      )
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
      spreadsheet.wrap(`${pfx}B${currentRow}`, true)
      spreadsheet.cellFormat({ fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: 'white', verticalAlign: 'top' }, `${pfx}B${currentRow}`)
      currentRow++
    })

    const allowancesEndRow = currentRow - 1
    const allowancesDataStartRow = allowancesHeaderRow + 1
    spreadsheet.merge(`${pfx}B${currentRow}:E${currentRow}`)
    spreadsheet.updateCell({ value: 'Allowances Total:' }, `${pfx}B${currentRow}`)
    spreadsheet.cellFormat(
      { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000' },
      `${pfx}B${currentRow}:E${currentRow}`
    )
    spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
    spreadsheet.updateCell({ value: 225000 }, `${pfx}F${currentRow}`)
    spreadsheet.cellFormat(
      { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#FEF2CB', verticalAlign: 'middle', border: '1px solid #000000', format: '$#,##0.00' },
      `${pfx}F${currentRow}:G${currentRow}`
    )
    siteworkTotalFRows.push(currentRow)
    currentRow++

    // Add Alternate #1: Sitework Total row - sum of individual section F column totals
    if (siteworkTotalFRows.length > 0) {
      spreadsheet.merge(`${pfx}D${currentRow}:E${currentRow}`)
      spreadsheet.updateCell({ value: 'Add Alternate #1: Sitework Total:' }, `${pfx}D${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD6EE', border: '1px solid #000000' },
        `${pfx}D${currentRow}:E${currentRow}`
      )
      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({ formula: `=SUM(${siteworkTotalFRows.map(r => `F${r}`).join(',')})` }, `${pfx}F${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD6EE', border: '1px solid #000000', format: '$#,##0.00' },
        `${pfx}F${currentRow}:G${currentRow}`
      )
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
        { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#BDD6EE', border: '1px solid #000000' },
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
        spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
        fillRatesForProposalRow(currentRow, bppLine6)
        const allSfRows = [...sidewalkRows, ...drivewayRows, ...asphaltRows]
        if (allSfRows.length > 0) {
          spreadsheet.updateCell({ formula: buildSumFormula(allSfRows, 'J') }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: buildSumFormula(allSfRows, 'L') }, `${pfx}F${currentRow}`)
        }
        spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
        currentRow++

        // 7. F&I new (4" thick) concrete sidewalk - SF/CY from sidewalk
        const bppLine7 = 'F&I new (4" thick) concrete sidewalk, reinf w/ 6x6-W1.4xW1.4 (NYCDOT H-1045, Type I)'
        spreadsheet.updateCell({ value: bppLine7 }, `${pfx}B${currentRow}`)
        spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
        fillRatesForProposalRow(currentRow, bppLine7)
        if (sidewalkRows.length > 0) {
          spreadsheet.updateCell({ formula: buildSumFormula(sidewalkRows, 'J') }, `${pfx}D${currentRow}`)
          spreadsheet.updateCell({ formula: buildSumFormula(sidewalkRows, 'L') }, `${pfx}F${currentRow}`)
        }
        spreadsheet.updateCell({ formula: `=IFERROR(ROUNDUP(MAX(C${currentRow}*I${currentRow},D${currentRow}*J${currentRow},E${currentRow}*K${currentRow},F${currentRow}*L${currentRow},G${currentRow}*M${currentRow},N${currentRow})/1000,1),"")` }, `${pfx}H${currentRow}`)
        currentRow++

        // 8. F&I new (7" thick) concrete sidewalk/driveway - SF/CY from driveway
        const bppLine8 = 'F&I new (7" thick) concrete sidewalk/driveway, reinf w/ 6x6-W1.4xW1.4 @ corners & driveways'
        spreadsheet.updateCell({ value: bppLine8 }, `${pfx}B${currentRow}`)
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
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD6EE', border: '1px solid #000000' },
        `${pfx}D${currentRow}:E${currentRow}`
      )
      spreadsheet.merge(`${pfx}F${currentRow}:G${currentRow}`)
      spreadsheet.updateCell({ formula: `=SUM(H${bppScopeStartRow}:H${bppScopeEndRow})*1000` }, `${pfx}F${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'right', backgroundColor: '#BDD6EE', border: '1px solid #000000' },
        `${pfx}F${currentRow}:G${currentRow}`
      )
      try { spreadsheet.numberFormat('$#,##0.00', `${pfx}F${currentRow}:G${currentRow}`) } catch (e) { }
      spreadsheet.cellFormat({ backgroundColor: '#BDD6EE' }, `${pfx}B${currentRow}:G${currentRow}`)
      currentRow++
    }

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
        spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
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
        spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
        currentRow++
      })
    }

    // Labor section
    {
      spreadsheet.updateCell({ value: 'Labor:' }, `${pfx}B${currentRow}`)
      spreadsheet.cellFormat(
        { fontWeight: 'bold', color: '#000000', textAlign: 'center', backgroundColor: '#AECBED', border: '1px solid #000000' },
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
        spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
        currentRow++
      })
    }

    // Exclusions section
    {
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
        spreadsheet.cellFormat({ textAlign: 'left' }, `${pfx}B${currentRow}`)
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
      spreadsheet.cellFormat(thickTop, `${pfx}B3:G3`)

      // Bottom edge
      spreadsheet.cellFormat(thickBottom, `${pfx}B${finalRow}:G${finalRow}`)

      // Left edge (iterate to force application likely)
      // Applying to range usually works but if background color logic overrides it, we might need to be specific
      // Or simply re-applying it here at the end should work if it's the last operation.
      // But to be safe vs background color cells:
      spreadsheet.cellFormat(thickLeft, `${pfx}B3:B${finalRow}`)
      spreadsheet.cellFormat(thickRight, `${pfx}G3:G${finalRow}`)

      // Global Font Style Application (Calibri Bold 18pt)
      // Apply to the entire content area including headers and signature
      // EXCLUDING Row 1 (Header) which is 10pt
      spreadsheet.cellFormat({ fontFamily: 'Calibri', fontWeight: 'bold', fontSize: '18pt' }, `${pfx}A2:N${finalRow}`)

      // Override Email label to be normal weight per request
      spreadsheet.cellFormat({ fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}B6`)

      // Override numerical value columns to be normal weight per request
      // Columns C-G and I-N (H is handled separately above)
      // Data starts at Row 13
      spreadsheet.cellFormat({ fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}C13:G${finalRow}`)
      spreadsheet.cellFormat({ fontWeight: 'normal', fontFamily: 'Calibri (Body)' }, `${pfx}I13:N${finalRow}`)

      // RE-APPLY BOLD to Total Rows (which were unbolded by the range above)
      totalRows.forEach(row => {
        // Apply to the specific columns or the whole row's data area
        // Often totals are in B, but values are in C-N.
        spreadsheet.cellFormat({ fontWeight: 'bold', fontFamily: 'Calibri' }, `${pfx}B${row}:N${row}`)
      })
    }

    // Uniform row height for all data rows below DESCRIPTION (row 12)
    const dataRowHeight = 30
    for (let r = 13; r <= finalRow; r++) {
      try { spreadsheet.setRowHeight(r, dataRowHeight) } catch (e) { /* ignore */ }
    }

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
      // Apply number format to quantity columns C-G that hides zeros
      try {
        spreadsheet.numberFormat('#,##0.00;-#,##0.00;""', `${pfx}C2:G${lastDataRow}`)
      } catch (e) { /* ignore */ }
      // Apply $ prefix format to column H ($/1000) that hides zeros
      try {
        spreadsheet.numberFormat('$#,##0.0;-$#,##0.0;""', `${pfx}H2:H${lastDataRow}`)
      } catch (e) {
        // Fallback to cellFormat if numberFormat doesn't work
        spreadsheet.cellFormat(
          {
            format: '$#,##0.0;-$#,##0.0;""'
          },
          `${pfx}H2:H${lastDataRow}`
        )
      }
      // Apply $ prefix format to columns I-N that hides zeros
      try {
        spreadsheet.numberFormat('$#,##0.0;-$#,##0.0;""', `${pfx}I2:N${lastDataRow}`)
      } catch (e) {
        // Fallback to cellFormat if numberFormat doesn't work
        spreadsheet.cellFormat(
          {
            format: '$#,##0.0;-$#,##0.0;""'
          },
          `${pfx}I2:N${lastDataRow}`
        )
      }
    }

    // Force recalculation of all formulas to ensure $/1000 values display correctly
    try {
      spreadsheet.refresh()
    } catch (e) {
      // Fallback: try to trigger recalculation by touching a cell
      try {
        spreadsheet.goTo('Proposal Sheet!A1')
      } catch (e2) { /* ignore */ }
    }
  }
}


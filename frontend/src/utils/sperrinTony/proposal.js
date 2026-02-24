/**
 * Sperrin Tony – Proposal tab.
 * Renders the main Proposal section (header / summary content).
 * No dependency on buildProposalSheet or other templates.
 */

import { PROPOSAL_SHEET_PFX } from './shared'

/**
 * @param {object} spreadsheet - Syncfusion spreadsheet instance
 * @param {object} ctx - { calculationData, formulaData, rawData, project, client, createdAt, currentRow, pfx }
 * @returns {number} next row index after this section
 */
export function buildProposalTab(spreadsheet, ctx) {
  const pfx = ctx.pfx ?? PROPOSAL_SHEET_PFX
  let row = ctx.currentRow ?? 1

  // Placeholder: Proposal tab content (header, project/client, etc.)
  spreadsheet.updateCell({ value: 'Proposal' }, `${pfx}B${row}`)
  spreadsheet.cellFormat(
    { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#E2EFDA' },
    `${pfx}B${row}`
  )
  row++

  if (ctx.project) {
    spreadsheet.updateCell({ value: `Project: ${ctx.project}` }, `${pfx}B${row}`)
    row++
  }
  if (ctx.client) {
    spreadsheet.updateCell({ value: `Client: ${ctx.client}` }, `${pfx}B${row}`)
    row++
  }

  row++ // spacing
  return row
}

/**
 * Sperrin Tony – Sitework Breakdown tab.
 * Renders sitework / civil sitework breakdown section.
 * No dependency on buildProposalSheet or other templates.
 */

import { PROPOSAL_SHEET_PFX } from './shared'

/**
 * @param {object} spreadsheet - Syncfusion spreadsheet instance
 * @param {object} ctx - { calculationData, formulaData, rawData, currentRow, pfx, ... }
 * @returns {number} next row index after this section
 */
export function buildSiteworkBreakdown(spreadsheet, ctx) {
  const pfx = ctx.pfx ?? PROPOSAL_SHEET_PFX
  let row = ctx.currentRow ?? 1

  spreadsheet.updateCell({ value: 'Sitework Breakdown' }, `${pfx}B${row}`)
  spreadsheet.cellFormat(
    { fontWeight: 'bold', color: '#000000', textAlign: 'left', backgroundColor: '#E2EFDA', textDecoration: 'underline' },
    `${pfx}B${row}`
  )
  row++

  // Section content to be implemented per Sperrin Tony spec
  row++
  return row
}

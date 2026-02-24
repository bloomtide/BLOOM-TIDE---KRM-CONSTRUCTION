/**
 * Sperrin Tony – sheet generation entry.
 * One sheet per tab (Proposal, Earthwork Breakdown, SOE Breakdown, etc.).
 * No dependency on buildProposalSheet or other templates.
 */

import { SPERRIN_TONY_SHEET_NAMES } from './shared'

export { SPERRIN_TONY_SHEET_NAMES } from './shared'
import { buildProposalTab } from './proposal'
import { buildEarthworkBreakdown } from './earthworkBreakdown'
import { buildSoeBreakdown } from './soeBreakdown'
import { buildSubstructureBreakdown } from './substructureBreakdown'
import { buildSuperstructureBreakdown } from './superstructureBreakdown'
import { buildMasonryBreakdown } from './masonryBreakdown'
import { buildFireproofingBreakdown } from './fireproofingBreakdown'
import { buildBppBreakdown } from './bppBreakdown'
import { buildSiteworkBreakdown } from './siteworkBreakdown'

const BUILDERS = [
  buildProposalTab,
  buildEarthworkBreakdown,
  buildSoeBreakdown,
  buildSubstructureBreakdown,
  buildSuperstructureBreakdown,
  buildMasonryBreakdown,
  buildFireproofingBreakdown,
  buildBppBreakdown,
  buildSiteworkBreakdown
]

/**
 * Build the full Sperrin Tony workbook: one sheet per tab.
 * Same options shape as buildProposalSheet for drop-in use from ProposalDetail.
 *
 * @param {object} spreadsheet - Syncfusion spreadsheet instance
 * @param {object} options - { calculationData, formulaData, rockExcavationTotals, lineDrillTotalFT, rawData, project, client, createdAt }
 */
export function buildSperrinTonySheet(spreadsheet, options = {}) {
  const {
    calculationData = [],
    formulaData = [],
    rockExcavationTotals = { totalSQFT: 0, totalCY: 0 },
    lineDrillTotalFT = 0,
    rawData = null,
    project = '',
    client = '',
    createdAt = null
  } = options

  for (let i = 0; i < SPERRIN_TONY_SHEET_NAMES.length && i < BUILDERS.length; i++) {
    const sheetName = SPERRIN_TONY_SHEET_NAMES[i]
    const pfx = `${sheetName}!`
    const ctx = {
      calculationData,
      formulaData,
      rockExcavationTotals,
      lineDrillTotalFT,
      rawData,
      project,
      client,
      createdAt,
      currentRow: 1,
      pfx,
      sheetIndex: i
    }
    BUILDERS[i](spreadsheet, ctx)
  }
}

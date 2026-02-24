/**
 * Shared helpers for Sperrin Tony sheet generation.
 * No dependencies on buildProposalSheet or other proposal utils.
 */

export const PROPOSAL_SHEET_PFX = 'Proposal Sheet!'
export const CALC_SHEET_NAME = 'Calculations Sheet'

/**
 * Sperrin Tony tab names – one sheet per tab. Order is fixed (do not reorder).
 * Display order: Proposal → Earthwork → SOE → Substructure → Superstructure → Masonry → Fireproofing → BPP → Sitework.
 */
export const SPERRIN_TONY_SHEET_NAMES = [
  'Proposal',
  'Earthwork Breakdown',
  'SOE Breakdown',
  'Substructure Breakdown',
  'Superstructure Breakdown',
  'Masonry Breakdown',
  'Fireproofing Breakdown',
  'BPP Breakdown',
  'Sitework Breakdown'
]

/**
 * Format multiple page refs: "102 & 105" or "102, 103 & 105"
 */
export function formatPageRefList(refs) {
  const list = Array.isArray(refs) ? refs.filter(Boolean) : []
  if (list.length === 0) return '##'
  if (list.length === 1) return list[0]
  if (list.length === 2) return `${list[0]} & ${list[1]}`
  return list.slice(0, -1).join(', ') + ' & ' + list[list.length - 1]
}

/**
 * Escape double quotes for use inside formula strings
 */
export function esc(s) {
  return (s ?? '').toString().replace(/"/g, '""')
}

/**
 * Numeric feet (e.g. 4.5) -> "4'-6""
 */
export function formatWidthFeetInches(val) {
  const n = parseFloat(val)
  if (val == null || val === '' || isNaN(n)) return null
  const f = Math.floor(n)
  const inch = Math.round((n - f) * 12)
  if (inch === 0) return `${f}'-0"`
  return `${f}'-${inch}"`
}

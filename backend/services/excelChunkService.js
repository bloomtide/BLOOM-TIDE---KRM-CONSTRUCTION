import ProposalExcelChunk, { EXCEL_CHUNK_SIZE } from '../models/ProposalExcelChunk.js';

/**
 * Save raw Excel rows as chunks (e.g. 200 rows per chunk).
 * Replaces any existing chunks for this proposal + sheetId.
 * @param {string} proposalId - Proposal _id
 * @param {string} sheetId - Sheet name / id (e.g. 'Sheet1')
 * @param {Array<Array>} rows - 2D array of cell values
 */
export async function saveRawExcelChunks(proposalId, sheetId, rows) {
    if (!proposalId || !sheetId) return;
    const rowList = Array.isArray(rows) ? rows : [];
    await ProposalExcelChunk.deleteMany({ proposalId, sheetId });
    if (rowList.length === 0) return;
    const chunks = [];
    for (let i = 0; i < rowList.length; i += EXCEL_CHUNK_SIZE) {
        chunks.push({
            proposalId,
            sheetId,
            chunkIndex: Math.floor(i / EXCEL_CHUNK_SIZE),
            rows: rowList.slice(i, i + EXCEL_CHUNK_SIZE),
        });
    }
    if (chunks.length > 0) {
        await ProposalExcelChunk.insertMany(chunks);
    }
}

/**
 * Load all chunks for a proposal + sheet and return concatenated rows.
 * @param {string} proposalId - Proposal _id
 * @param {string} sheetId - Sheet name / id
 * @returns {Promise<Array<Array>>} 2D array of cell values
 */
export async function loadRawExcelChunks(proposalId, sheetId) {
    if (!proposalId || !sheetId) return [];
    const chunks = await ProposalExcelChunk.find({ proposalId, sheetId })
        .sort({ chunkIndex: 1 })
        .lean();
    return chunks.reduce((acc, c) => acc.concat(c.rows || []), []);
}

/**
 * Check if any chunks exist for this proposal (so we know to load from chunks instead of Proposal.rawExcelData.rows).
 */
export async function hasChunks(proposalId, sheetId = 'Sheet1') {
    const count = await ProposalExcelChunk.countDocuments({ proposalId, sheetId }).limit(1);
    return count > 0;
}

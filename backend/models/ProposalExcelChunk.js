import mongoose from 'mongoose';

/** Chunk size for raw Excel rows (keeps docs small, fetch only needed chunks) */
export const EXCEL_CHUNK_SIZE = 200;

const proposalExcelChunkSchema = new mongoose.Schema(
    {
        proposalId: {
            type: mongoose.Schema.Types.ObjectId,
            ref: 'Proposal',
            required: true,
            index: true,
        },
        sheetId: {
            type: String,
            required: true,
            default: 'Sheet1',
        },
        chunkIndex: {
            type: Number,
            required: true,
            min: 0,
        },
        rows: {
            type: [[mongoose.Schema.Types.Mixed]],
            required: true,
            default: [],
        },
    },
    { timestamps: true }
);

// Fetch chunks for a proposal+sheet in one query; compound index for list by proposalId + sheetId + chunkIndex
proposalExcelChunkSchema.index({ proposalId: 1, sheetId: 1, chunkIndex: 1 });

const ProposalExcelChunk = mongoose.model('ProposalExcelChunk', proposalExcelChunkSchema);
export default ProposalExcelChunk;

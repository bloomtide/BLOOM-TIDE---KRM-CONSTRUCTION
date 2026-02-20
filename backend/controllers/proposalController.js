import zlib from 'zlib';
import Proposal from '../models/Proposal.js';
import ProposalExcelChunk from '../models/ProposalExcelChunk.js';
import XLSX from 'xlsx';
import { saveRawExcelChunks, loadRawExcelChunks, hasChunks } from '../services/excelChunkService.js';
import { uploadBufferToS3, getBufferFromS3, getS3UrlForKey } from '../services/s3Service.js';

// @desc    Create new proposal
// @route   POST /api/proposals
// @access  Private
export const createProposal = async (req, res) => {
    try {
        const { name, client, project, template } = req.body;

        // Validate required fields
        if (!name || !client || !project || !template) {
            return res.status(400).json({
                success: false,
                message: 'Please provide name, client, project, and template',
            });
        }

        // Check if file was uploaded
        if (!req.file) {
            return res.status(400).json({
                success: false,
                message: 'Please upload an Excel file',
            });
        }

        // Parse Excel file from buffer
        const workbook = XLSX.read(req.file.buffer, { type: 'buffer' });
        const firstSheetName = workbook.SheetNames[0];
        const worksheet = workbook.Sheets[firstSheetName];
        const jsonData = XLSX.utils.sheet_to_json(worksheet, { header: 1, defval: '' });

        // Extract headers and rows
        const headers = jsonData[0] || [];
        const rows = jsonData.slice(1);

        // Create proposal (rows stored in chunks, not in main doc)
        const proposal = await Proposal.create({
            name,
            client,
            project,
            template,
            rawExcelData: {
                fileName: req.file.originalname,
                sheetName: firstSheetName,
                headers,
                rows: [], // rows stored in ProposalExcelChunk
            },
            createdBy: req.user._id,
        });
        await saveRawExcelChunks(proposal._id, firstSheetName, rows);

        // Upload raw Excel file to S3 and persist key/URL on proposal (source of truth for GET)
        const safeName = (req.file.originalname || 'upload.xlsx').replace(/[^a-zA-Z0-9._-]/g, '_');
        const s3Key = `proposals/${proposal._id}/raw/${safeName}`;
        let rawExcelS3Url = null;
        try {
            rawExcelS3Url = await uploadBufferToS3(req.file.buffer, s3Key);
            if (rawExcelS3Url) {
                await Proposal.findByIdAndUpdate(proposal._id, {
                    rawExcelS3Key: s3Key,
                    rawExcelS3Url,
                });
                console.log('[Proposal] Raw Excel uploaded to S3:', rawExcelS3Url);
            }
        } catch (s3Err) {
            console.error('[Proposal] S3 upload failed:', s3Err.message);
        }

        // Attach assembled rows and S3 URL for response
        const proposalObj = (await Proposal.findById(proposal._id)).toObject();
        proposalObj.rawExcelData = proposalObj.rawExcelData || {};
        proposalObj.rawExcelData.rows = rows;
        if (rawExcelS3Url) proposalObj.rawExcelS3Url = rawExcelS3Url;
        res.status(201).json({
            success: true,
            proposal: proposalObj,
        });
    } catch (error) {
        console.error('Error creating proposal:', error);
        res.status(500).json({
            success: false,
            message: error.message || 'Error creating proposal',
        });
    }
};

// @desc    Get all proposals with pagination and filtering
// @route   GET /api/proposals
// @access  Private
export const getProposals = async (req, res) => {
    try {
        const {
            page = 1,
            limit = 10,
            search,
            startDate,
            endDate
        } = req.query;

        // Build query
        const query = {};

        // Search functionality
        if (search) {
            const searchRegex = new RegExp(search, 'i');
            query.$or = [
                { name: searchRegex },
                { client: searchRegex },
                { project: searchRegex },
                { template: searchRegex }
            ];
        }

        // Date range filter
        if (startDate && endDate) {
            query.createdAt = {
                $gte: new Date(startDate),
                $lte: new Date(endDate)
            };
        } else if (startDate) {
            query.createdAt = { $gte: new Date(startDate) };
        } else if (endDate) {
            query.createdAt = { $lte: new Date(endDate) };
        }

        // Pagination
        const pageNumber = parseInt(page, 10);
        const limitNumber = parseInt(limit, 10);
        const skip = (pageNumber - 1) * limitNumber;

        const proposals = await Proposal.find(query)
            .select('-rawExcelData -spreadsheetJson -images -unusedRawDataRows -rawExcelS3Key -spreadsheetS3Key') // Exclude inner/large data for list view
            .sort({ createdAt: -1 })
            .skip(skip)
            .limit(limitNumber)
            .allowDiskUse(true)
            .lean();

        const total = await Proposal.countDocuments(query);

        res.status(200).json({
            success: true,
            count: proposals.length,
            pagination: {
                page: pageNumber,
                pages: Math.ceil(total / limitNumber),
                total
            },
            proposals,
        });
    } catch (error) {
        console.error('Error fetching proposals:', error.message, error.stack);
        res.status(500).json({
            success: false,
            message: error.message || 'Error fetching proposals',
        });
    }
};

// @desc    Get single proposal by ID. Raw Excel: from S3 (fetch + parse in real time) or from chunks.
// @route   GET /api/proposals/:id
// @access  Private
export const getProposalById = async (req, res) => {
    try {
        const proposal = await Proposal.findById(req.params.id);

        if (!proposal) {
            return res.status(404).json({
                success: false,
                message: 'Proposal not found',
            });
        }

        const out = proposal.toObject();

        // When S3 is used, do NOT embed large payloads; frontend fetches from file URLs and parses
        if (out.rawExcelS3Key) {
            out.rawExcelFileUrl = `/api/proposals/${req.params.id}/raw-file`;
            if (out.rawExcelS3Url == null) out.rawExcelS3Url = getS3UrlForKey(out.rawExcelS3Key);
            if (out.rawExcelData) {
                out.rawExcelData.rows = []; // keep metadata only; client fetches file and parses
            }
        } else {
            // Fallback: assemble rows from chunks (no S3)
            if (out.rawExcelData && out.rawExcelData.sheetName && (!out.rawExcelData.rows || out.rawExcelData.rows.length === 0)) {
                const sheetId = out.rawExcelData.sheetName;
                const useChunks = await hasChunks(proposal._id, sheetId);
                if (useChunks) {
                    out.rawExcelData.rows = await loadRawExcelChunks(proposal._id, sheetId);
                }
            }
        }

        if (out.spreadsheetS3Key) {
            out.spreadsheetFileUrl = `/api/proposals/${req.params.id}/spreadsheet-file`;
            if (out.spreadsheetS3Url == null) out.spreadsheetS3Url = getS3UrlForKey(out.spreadsheetS3Key);
            out.spreadsheetJson = null; // client fetches file and loads into Syncfusion
        }

        res.status(200).json({
            success: true,
            proposal: out,
        });
    } catch (error) {
        console.error('Error fetching proposal:', error);
        res.status(500).json({
            success: false,
            message: 'Error fetching proposal',
        });
    }
};

// @desc    Stream raw Excel file from S3 (client parses to get rows)
// @route   GET /api/proposals/:id/raw-file
// @access  Private
export const getRawFile = async (req, res) => {
    try {
        const proposal = await Proposal.findById(req.params.id).select('rawExcelS3Key').lean();
        if (!proposal?.rawExcelS3Key) {
            return res.status(404).json({ success: false, message: 'Raw file not found' });
        }
        const buffer = await getBufferFromS3(proposal.rawExcelS3Key);
        if (!buffer) return res.status(404).json({ success: false, message: 'Raw file not found' });
        res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        res.send(buffer);
    } catch (err) {
        console.error('[Proposal] getRawFile:', err.message);
        res.status(500).json({ success: false, message: 'Error fetching raw file' });
    }
};

// @desc    Stream spreadsheet JSON (gzipped) from S3 (client decompresses and loads into Syncfusion)
// @route   GET /api/proposals/:id/spreadsheet-file
// @access  Private
export const getSpreadsheetFile = async (req, res) => {
    try {
        const proposal = await Proposal.findById(req.params.id).select('spreadsheetS3Key').lean();
        if (!proposal?.spreadsheetS3Key) {
            return res.status(404).json({ success: false, message: 'Spreadsheet file not found' });
        }
        const buffer = await getBufferFromS3(proposal.spreadsheetS3Key);
        if (!buffer) return res.status(404).json({ success: false, message: 'Spreadsheet file not found' });
        res.setHeader('Content-Type', 'application/gzip');
        res.send(buffer);
    } catch (err) {
        console.error('[Proposal] getSpreadsheetFile:', err.message);
        res.status(500).json({ success: false, message: 'Error fetching spreadsheet file' });
    }
};

// @desc    Update proposal
// @route   PUT /api/proposals/:id
// @access  Private
export const updateProposal = async (req, res) => {
    try {
        const { name, client, project, spreadsheetJson, spreadsheetJsonCompressed, images, unusedRawDataRows, rawExcelData } = req.body;

        // Build update object - only include fields that were sent
        const update = {};
        if (name !== undefined) update.name = name;
        if (client !== undefined) update.client = client;
        if (project !== undefined) update.project = project;
        let spreadsheetJsonObj = null; // for S3 upload
        if (spreadsheetJsonCompressed !== undefined) {
            try {
                const buf = Buffer.from(spreadsheetJsonCompressed, 'base64');
                const decompressed = zlib.gunzipSync(buf);
                spreadsheetJsonObj = JSON.parse(decompressed.toString('utf8'));
            } catch (err) {
                console.error('Error decompressing spreadsheetJsonCompressed:', err);
                return res.status(400).json({
                    success: false,
                    message: 'Invalid spreadsheetJsonCompressed payload',
                });
            }
        } else if (spreadsheetJson !== undefined) {
            spreadsheetJsonObj = spreadsheetJson;
        }
        if (spreadsheetJsonObj !== null) {
            update._uploadSpreadsheetToS3 = spreadsheetJsonObj;
            update.spreadsheetJson = null; // store only URL in DB
        }
        if (images !== undefined) update.images = images;
        if (unusedRawDataRows !== undefined) update.unusedRawDataRows = unusedRawDataRows;
        if (rawExcelData !== undefined) {
            const rows = Array.isArray(rawExcelData.rows) ? rawExcelData.rows : [];
            update.rawExcelData = {
                fileName: rawExcelData.fileName ?? '',
                sheetName: rawExcelData.sheetName ?? 'Sheet1',
                headers: Array.isArray(rawExcelData.headers) ? rawExcelData.headers : [],
                rows: [], // rows stored in ProposalExcelChunk
            };
            // Save rows to chunks (after update we'll have proposal id)
            update._saveRawExcelChunks = { sheetId: update.rawExcelData.sheetName, rows };
        }

        if (Object.keys(update).length === 0) {
            const proposal = await Proposal.findById(req.params.id);
            if (!proposal) {
                return res.status(404).json({
                    success: false,
                    message: 'Proposal not found',
                });
            }
            return res.status(200).json({
                success: true,
                proposal,
            });
        }

        const saveChunks = update._saveRawExcelChunks;
        delete update._saveRawExcelChunks;
        const uploadSpreadsheetToS3 = update._uploadSpreadsheetToS3;
        delete update._uploadSpreadsheetToS3;

        const existingProposal = await Proposal.findById(req.params.id).select('rawExcelS3Key spreadsheetS3Key').lean();

        // Use findByIdAndUpdate to ensure MongoDB persistence (bypasses Mongoose doc state)
        const proposal = await Proposal.findByIdAndUpdate(
            req.params.id,
            { $set: update },
            { new: true, runValidators: true }
        );

        if (!proposal) {
            return res.status(404).json({
                success: false,
                message: 'Proposal not found',
            });
        }

        // Upload spreadsheet JSON (Proposal + Calculations sheets) to S3; keep only URL in DB
        if (uploadSpreadsheetToS3) {
            try {
                const jsonStr = JSON.stringify(uploadSpreadsheetToS3);
                const gzipBuffer = zlib.gzipSync(Buffer.from(jsonStr, 'utf8'));
                const s3Key = existingProposal?.spreadsheetS3Key || `proposals/${proposal._id}/spreadsheet.json.gz`;
                const url = await uploadBufferToS3(gzipBuffer, s3Key, 'application/gzip');
                if (url) {
                    await Proposal.findByIdAndUpdate(proposal._id, {
                        spreadsheetS3Key: s3Key,
                        spreadsheetS3Url: url,
                    });
                }
            } catch (s3Err) {
                console.error('[Proposal] S3 spreadsheet upload failed:', s3Err.message);
            }
        }

        if (saveChunks) {
            await saveRawExcelChunks(proposal._id, saveChunks.sheetId, saveChunks.rows);

            // Re-upload Excel to S3 so GET always uses S3 and converts in real time
            const headers = (update.rawExcelData && update.rawExcelData.headers) || [];
            const sheetName = (update.rawExcelData && update.rawExcelData.sheetName) || 'Sheet1';
            const fileName = (update.rawExcelData && update.rawExcelData.fileName) || 'upload.xlsx';
            const aoa = [headers, ...saveChunks.rows];
            const wb = XLSX.utils.book_new();
            const ws = XLSX.utils.aoa_to_sheet(aoa);
            XLSX.utils.book_append_sheet(wb, ws, sheetName);
            const xlsxBuffer = XLSX.write(wb, { type: 'buffer', bookType: 'xlsx' });
            const safeName = fileName.replace(/[^a-zA-Z0-9._-]/g, '_') || 'upload.xlsx';
            const s3Key = existingProposal?.rawExcelS3Key || `proposals/${proposal._id}/raw/${safeName}`;
            try {
                const url = await uploadBufferToS3(xlsxBuffer, s3Key);
                if (url) {
                    await Proposal.findByIdAndUpdate(proposal._id, { rawExcelS3Key: s3Key, rawExcelS3Url: url });
                }
            } catch (s3Err) {
                console.error('[Proposal] S3 re-upload on update failed:', s3Err.message);
            }
        }

        const out = proposal.toObject();
        if (saveChunks) {
            out.rawExcelFileUrl = `/api/proposals/${req.params.id}/raw-file`;
            if (out.rawExcelData) out.rawExcelData.rows = [];
        }
        if (uploadSpreadsheetToS3) {
            out.spreadsheetFileUrl = `/api/proposals/${req.params.id}/spreadsheet-file`;
            out.spreadsheetJson = null;
        }

        res.status(200).json({
            success: true,
            proposal: out,
        });
    } catch (error) {
        console.error('Error updating proposal:', error);
        res.status(500).json({
            success: false,
            message: 'Error updating proposal',
        });
    }
};

// @desc    Update unused row status (toggle isUsed checkbox)
// @route   PATCH /api/proposals/:id/unused-rows/:rowIndex
// @access  Private
export const updateUnusedRowStatus = async (req, res) => {
    try {
        const { id, rowIndex } = req.params;
        const { isUsed } = req.body;

        if (typeof isUsed !== 'boolean') {
            return res.status(400).json({
                success: false,
                message: 'isUsed must be a boolean value',
            });
        }

        const rowIndexNum = parseInt(rowIndex, 10);
        if (isNaN(rowIndexNum)) {
            return res.status(400).json({
                success: false,
                message: 'Invalid row index',
            });
        }

        // Find the proposal and update the specific row's isUsed status
        const proposal = await Proposal.findOneAndUpdate(
            { _id: id, 'unusedRawDataRows.rowIndex': rowIndexNum },
            { $set: { 'unusedRawDataRows.$.isUsed': isUsed } },
            { new: true }
        );

        if (!proposal) {
            return res.status(404).json({
                success: false,
                message: 'Proposal or row not found',
            });
        }

        res.status(200).json({
            success: true,
            message: 'Row status updated successfully',
            unusedRawDataRows: proposal.unusedRawDataRows,
        });
    } catch (error) {
        console.error('Error updating unused row status:', error);
        res.status(500).json({
            success: false,
            message: 'Error updating row status',
        });
    }
};

// @desc    Bulk update unused row statuses
// @route   PATCH /api/proposals/:id/unused-rows/bulk
// @access  Private
export const updateUnusedRowStatusBulk = async (req, res) => {
    try {
        const { id } = req.params;
        const { updates } = req.body; // updates is array of { rowIndex, isUsed }

        if (!Array.isArray(updates)) {
            return res.status(400).json({
                success: false,
                message: 'Updates must be an array',
            });
        }

        const proposal = await Proposal.findById(id);
        if (!proposal) {
            return res.status(404).json({
                success: false,
                message: 'Proposal not found',
            });
        }

        // Apply all updates
        // We can do this efficiently by iterating and updating the map/array
        // Since unusedRawDataRows is an array of objects, we can update in memory then save
        let updatedCount = 0;

        updates.forEach(update => {
            const { rowIndex, isUsed } = update;
            if (rowIndex !== undefined && typeof isUsed === 'boolean') {
                const rowIndexNum = parseInt(rowIndex, 10);
                const row = proposal.unusedRawDataRows.find(r => r.rowIndex === rowIndexNum);
                if (row) {
                    row.isUsed = isUsed;
                    updatedCount++;
                }
            }
        });

        if (updatedCount > 0) {
            await proposal.save();
        }

        res.status(200).json({
            success: true,
            message: `${updatedCount} rows updated successfully`,
            unusedRawDataRows: proposal.unusedRawDataRows,
        });

    } catch (error) {
        console.error('Error bulk updating unused row status:', error);
        res.status(500).json({
            success: false,
            message: 'Error bulk updating row status',
        });
    }
};

// @desc    Delete proposal
// @route   DELETE /api/proposals/:id
// @access  Private
export const deleteProposal = async (req, res) => {
    try {
        const proposal = await Proposal.findById(req.params.id);

        if (!proposal) {
            return res.status(404).json({
                success: false,
                message: 'Proposal not found',
            });
        }

        await ProposalExcelChunk.deleteMany({ proposalId: proposal._id });
        await proposal.deleteOne();

        res.status(200).json({
            success: true,
            message: 'Proposal deleted successfully',
        });
    } catch (error) {
        console.error('Error deleting proposal:', error);
        res.status(500).json({
            success: false,
            message: 'Error deleting proposal',
        });
    }
};

// @desc    Duplicate proposal
// @route   POST /api/proposals/:id/duplicate
// @access  Private
export const duplicateProposal = async (req, res) => {
    try {
        const proposal = await Proposal.findById(req.params.id);

        if (!proposal) {
            return res.status(404).json({
                success: false,
                message: 'Proposal not found',
            });
        }

        let rawExcelData = proposal.rawExcelData;
        let rowsToChunk = [];
        if (proposal.rawExcelData && proposal.rawExcelData.sheetName) {
            const useChunks = await hasChunks(proposal._id, proposal.rawExcelData.sheetName);
            if (useChunks) {
                rowsToChunk = await loadRawExcelChunks(proposal._id, proposal.rawExcelData.sheetName);
                const raw = proposal.rawExcelData.toObject ? proposal.rawExcelData.toObject() : { ...proposal.rawExcelData };
                rawExcelData = { ...raw, rows: [] };
            }
        }

        const duplicated = await Proposal.create({
            name: `${proposal.name} (Copy)`,
            client: proposal.client,
            project: proposal.project,
            template: proposal.template,
            rawExcelData: rawExcelData || proposal.rawExcelData,
            spreadsheetJson: proposal.spreadsheetS3Key ? null : proposal.spreadsheetJson,
            images: proposal.images,
            createdBy: req.user._id,
        });

        if (rowsToChunk.length > 0 && duplicated.rawExcelData && duplicated.rawExcelData.sheetName) {
            await saveRawExcelChunks(duplicated._id, duplicated.rawExcelData.sheetName, rowsToChunk);
        }

        // If source had spreadsheet on S3, copy to duplicated proposal's key (Proposal + Calculations sheet JSON)
        if (proposal.spreadsheetS3Key) {
            try {
                const buffer = await getBufferFromS3(proposal.spreadsheetS3Key);
                if (buffer) {
                    const newKey = `proposals/${duplicated._id}/spreadsheet.json.gz`;
                    const url = await uploadBufferToS3(buffer, newKey, 'application/gzip');
                    if (url) {
                        await Proposal.findByIdAndUpdate(duplicated._id, {
                            spreadsheetS3Key: newKey,
                            spreadsheetS3Url: url,
                        });
                    }
                }
            } catch (s3Err) {
                console.error('[Proposal] S3 spreadsheet copy on duplicate failed:', s3Err.message);
            }
        }

        // If source had S3 file, copy object to duplicated proposal's key so GET uses S3
        let copiedRows = null;
        if (proposal.rawExcelS3Key) {
            try {
                const buffer = await getBufferFromS3(proposal.rawExcelS3Key);
                if (buffer) {
                    const safeName = (duplicated.rawExcelData?.fileName || 'upload.xlsx').replace(/[^a-zA-Z0-9._-]/g, '_');
                    const newKey = `proposals/${duplicated._id}/raw/${safeName}`;
                    const url = await uploadBufferToS3(buffer, newKey);
                    if (url) {
                        await Proposal.findByIdAndUpdate(duplicated._id, { rawExcelS3Key: newKey, rawExcelS3Url: url });
                    }
                    const wb = XLSX.read(buffer, { type: 'buffer' });
                    const sn = wb.SheetNames[0];
                    const ws = wb.Sheets[sn];
                    const jsonData = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '' });
                    copiedRows = jsonData.slice(1);
                }
            } catch (s3Err) {
                console.error('[Proposal] S3 copy on duplicate failed:', s3Err.message);
            }
        }

        const out = (await Proposal.findById(duplicated._id)).toObject();
        if (out.rawExcelS3Key) {
            out.rawExcelFileUrl = `/api/proposals/${duplicated._id}/raw-file`;
            if (out.rawExcelData) out.rawExcelData.rows = [];
        } else if (out.rawExcelData && rowsToChunk.length > 0) {
            out.rawExcelData.rows = rowsToChunk;
        } else if (out.rawExcelData && copiedRows) {
            out.rawExcelData.rows = copiedRows;
        }
        if (out.spreadsheetS3Key) {
            out.spreadsheetFileUrl = `/api/proposals/${duplicated._id}/spreadsheet-file`;
            out.spreadsheetJson = null;
        }

        res.status(201).json({
            success: true,
            proposal: out,
        });
    } catch (error) {
        console.error('Error duplicating proposal:', error);
        res.status(500).json({
            success: false,
            message: error.message || 'Error duplicating proposal',
        });
    }
};

// @desc    Delete multiple proposals (bulk delete)
// @route   POST /api/proposals/bulk-delete
// @access  Private
export const deleteProposals = async (req, res) => {
    try {
        const { ids } = req.body;

        if (!ids || ids.length === 0) {
            return res.status(400).json({
                success: false,
                message: 'No proposal IDs provided',
            });
        }

        await ProposalExcelChunk.deleteMany({ proposalId: { $in: ids } });
        const result = await Proposal.deleteMany({ _id: { $in: ids } });

        res.status(200).json({
            success: true,
            message: `${result.deletedCount} proposals deleted successfully`,
        });
    } catch (error) {
        console.error('Error deleting proposals:', error);
        res.status(500).json({
            success: false,
            message: 'Error deleting proposals',
        });
    }
};

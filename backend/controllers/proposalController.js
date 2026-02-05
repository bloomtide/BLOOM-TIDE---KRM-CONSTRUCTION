import Proposal from '../models/Proposal.js';
import XLSX from 'xlsx';

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

        // Create proposal
        const proposal = await Proposal.create({
            name,
            client,
            project,
            template,
            rawExcelData: {
                fileName: req.file.originalname,
                sheetName: firstSheetName,
                headers,
                rows,
            },
            createdBy: req.user._id,
        });

        res.status(201).json({
            success: true,
            proposal,
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
            .select('-rawExcelData -spreadsheetJson -images') // Exclude large data for list view
            .sort({ createdAt: -1 })
            .skip(skip)
            .limit(limitNumber);

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
        console.error('Error fetching proposals:', error);
        res.status(500).json({
            success: false,
            message: 'Error fetching proposals',
        });
    }
};

// @desc    Get single proposal by ID
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

        res.status(200).json({
            success: true,
            proposal,
        });
    } catch (error) {
        console.error('Error fetching proposal:', error);
        res.status(500).json({
            success: false,
            message: 'Error fetching proposal',
        });
    }
};

// @desc    Update proposal
// @route   PUT /api/proposals/:id
// @access  Private
export const updateProposal = async (req, res) => {
    try {
        const { name, client, project, spreadsheetJson, images } = req.body;

        // Build update object - only include fields that were sent
        const update = {};
        if (name !== undefined) update.name = name;
        if (client !== undefined) update.client = client;
        if (project !== undefined) update.project = project;
        if (spreadsheetJson !== undefined) update.spreadsheetJson = spreadsheetJson;
        if (images !== undefined) update.images = images;

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

        res.status(200).json({
            success: true,
            proposal,
        });
    } catch (error) {
        console.error('Error updating proposal:', error);
        res.status(500).json({
            success: false,
            message: 'Error updating proposal',
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

        // Delete multiple proposals
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

import express from 'express';
import {
    createProposal,
    getProposals,
    getProposalById,
    updateProposal,
    deleteProposal,
    deleteProposals,
    duplicateProposal,
} from '../controllers/proposalController.js';
import { protect } from '../middleware/auth.js';
import upload from '../middleware/upload.js';

const router = express.Router();

// All routes require authentication
router.use(protect);

// Routes
router.route('/')
    .get(getProposals)
    .post(upload.single('excelFile'), createProposal);

router.post('/bulk-delete', deleteProposals);
router.post('/:id/duplicate', duplicateProposal);

router.route('/:id')
    .get(getProposalById)
    .put(updateProposal)
    .delete(deleteProposal);

export default router;

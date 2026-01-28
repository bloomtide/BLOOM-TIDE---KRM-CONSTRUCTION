import express from 'express';
import {
    getUsers,
    createUser,
    getUserById,
    updateUser,
    deleteUser,
} from '../controllers/userController.js';
import { protect } from '../middleware/auth.js';
import { adminOnly } from '../middleware/roleCheck.js';

const router = express.Router();

// All routes require authentication and admin role
router.use(protect, adminOnly);

router.route('/').get(getUsers).post(createUser);
router.route('/:id').get(getUserById).put(updateUser).delete(deleteUser);

export default router;

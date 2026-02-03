import User from '../models/User.js';

// @desc    Get all users with pagination and search
// @route   GET /api/users
// @access  Private/Admin
export const getUsers = async (req, res) => {
    try {
        const { page = 1, limit = 10, search } = req.query;

        // Build query
        const query = {};

        // Search functionality
        if (search) {
            const searchRegex = new RegExp(search, 'i');
            query.$or = [
                { name: searchRegex },
                { email: searchRegex }
            ];
        }

        // Pagination
        const pageNumber = parseInt(page, 10);
        const limitNumber = parseInt(limit, 10);
        const skip = (pageNumber - 1) * limitNumber;

        const users = await User.find(query)
            .select('-password')
            .sort('-createdAt')
            .skip(skip)
            .limit(limitNumber);

        const total = await User.countDocuments(query);

        res.status(200).json({
            success: true,
            count: users.length,
            pagination: {
                page: pageNumber,
                pages: Math.ceil(total / limitNumber),
                total
            },
            users,
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            message: error.message,
        });
    }
};

// @desc    Create new user
// @route   POST /api/users
// @access  Private/Admin
export const createUser = async (req, res) => {
    try {
        const { name, email, password, role } = req.body;

        // Validate name
        if (!name || !name.trim()) {
            return res.status(400).json({
                success: false,
                message: 'Name is required',
            });
        }
        if (name.trim().length < 2) {
            return res.status(400).json({
                success: false,
                message: 'Name must be at least 2 characters',
            });
        }

        // Validate email
        const emailRegex = /^\S+@\S+\.\S+$/;
        if (!email || !email.trim()) {
            return res.status(400).json({
                success: false,
                message: 'Email is required',
            });
        }
        if (!emailRegex.test(email)) {
            return res.status(400).json({
                success: false,
                message: 'Please provide a valid email',
            });
        }

        // Check if user exists
        const userExists = await User.findOne({ email: email.toLowerCase() });
        if (userExists) {
            return res.status(400).json({
                success: false,
                message: 'User already exists with this email',
            });
        }

        // Validate password
        if (!password) {
            return res.status(400).json({
                success: false,
                message: 'Password is required',
            });
        }
        if (password.length < 6) {
            return res.status(400).json({
                success: false,
                message: 'Password must be at least 6 characters',
            });
        }

        // Create user
        const user = await User.create({
            name: name.trim(),
            email: email.toLowerCase(),
            password,
            role: role || 'user',
            createdBy: req.user._id,
        });

        res.status(201).json({
            success: true,
            user: {
                id: user._id,
                name: user.name,
                email: user.email,
                role: user.role,
            },
            message: 'User created successfully',
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            message: error.message,
        });
    }
};

// @desc    Get user by ID
// @route   GET /api/users/:id
// @access  Private/Admin
export const getUserById = async (req, res) => {
    try {
        const user = await User.findById(req.params.id).select('-password');

        if (!user) {
            return res.status(404).json({
                success: false,
                message: 'User not found',
            });
        }

        res.status(200).json({
            success: true,
            user,
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            message: error.message,
        });
    }
};

// @desc    Update user
// @route   PUT /api/users/:id
// @access  Private/Admin
export const updateUser = async (req, res) => {
    try {
        const { name, email, password, role } = req.body;

        const user = await User.findById(req.params.id);

        if (!user) {
            return res.status(404).json({
                success: false,
                message: 'User not found',
            });
        }

        // Validate name
        if (name !== undefined) {
            if (!name.trim()) {
                return res.status(400).json({
                    success: false,
                    message: 'Name cannot be empty',
                });
            }
            if (name.trim().length < 2) {
                return res.status(400).json({
                    success: false,
                    message: 'Name must be at least 2 characters',
                });
            }
            user.name = name.trim();
        }

        // Validate and check for duplicate email
        if (email !== undefined && email !== user.email) {
            const emailRegex = /^\S+@\S+\.\S+$/;
            if (!emailRegex.test(email)) {
                return res.status(400).json({
                    success: false,
                    message: 'Please provide a valid email',
                });
            }

            // Check if email is already taken by another user
            const emailExists = await User.findOne({ email: email.toLowerCase() });
            if (emailExists && emailExists._id.toString() !== req.params.id) {
                return res.status(400).json({
                    success: false,
                    message: 'Email is already in use by another user',
                });
            }

            user.email = email.toLowerCase();
        }

        // Update password if provided
        if (password) {
            if (password.length < 6) {
                return res.status(400).json({
                    success: false,
                    message: 'Password must be at least 6 characters',
                });
            }
            user.password = password;
        }

        // Update role if provided
        if (role !== undefined) {
            user.role = role;
        }

        const updatedUser = await user.save();

        res.status(200).json({
            success: true,
            user: {
                id: updatedUser._id,
                name: updatedUser.name,
                email: updatedUser.email,
                role: updatedUser.role,
            },
            message: 'User updated successfully',
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            message: error.message,
        });
    }
};

// @desc    Delete user (hard delete)
// @route   DELETE /api/users/:id
// @access  Private/Admin
export const deleteUser = async (req, res) => {
    try {
        const user = await User.findById(req.params.id);

        if (!user) {
            return res.status(404).json({
                success: false,
                message: 'User not found',
            });
        }

        // Hard delete - permanently remove user
        await user.deleteOne();

        res.status(200).json({
            success: true,
            message: 'User deleted successfully',
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            message: error.message,
        });
    }
};

// @desc    Delete multiple users (bulk hard delete)
// @route   POST /api/users/bulk-delete
// @access  Private/Admin
export const deleteUsers = async (req, res) => {
    try {
        const { ids } = req.body;

        if (!ids || ids.length === 0) {
            return res.status(400).json({
                success: false,
                message: 'No user IDs provided',
            });
        }

        // Hard delete multiple users
        const result = await User.deleteMany({ _id: { $in: ids } });

        res.status(200).json({
            success: true,
            message: `${result.deletedCount} users deleted successfully`,
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            message: error.message,
        });
    }
};
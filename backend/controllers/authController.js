import User from '../models/User.js';
import Session from '../models/Session.js';
import generateToken from '../utils/generateToken.js';


export const login = async (req, res) => {
    try {
        const { email, password } = req.body;

        // Validate input
        if (!email || !password) {
            return res.status(400).json({
                success: false,
                message: 'Please provide email and password',
            });
        }

        // Find user and include password
        const user = await User.findOne({ email }).select('+password');

        if (!user) {
            return res.status(401).json({
                success: false,
                message: 'Invalid credentials',
            });
        }

        // Check password
        const isPasswordCorrect = await user.matchPassword(password);

        if (!isPasswordCorrect) {
            return res.status(401).json({
                success: false,
                message: 'Invalid credentials',
            });
        }

        // Generate token
        const token = generateToken(user._id);

        // Create session
        const expiresAt = new Date();
        expiresAt.setDate(expiresAt.getDate() + parseInt(process.env.JWT_COOKIE_EXPIRE));

        await Session.create({
            userId: user._id,
            token,
            ipAddress: req.ip,
            userAgent: req.get('user-agent'),
            expiresAt,
        });

        // Update last login
        user.lastLogin = Date.now();
        await user.save({ validateBeforeSave: false });

        // Set cookie
        const cookieOptions = {
            expires: expiresAt,
            httpOnly: true,
            secure: process.env.NODE_ENV === 'production',
            sameSite: 'strict',
        };

        res.cookie('token', token, cookieOptions);

        res.status(200).json({
            success: true,
            token,
            user: {
                id: user._id,
                name: user.name,
                email: user.email,
                role: user.role,
            },
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            message: error.message,
        });
    }
};

// @desc    Logout user
// @route   POST /api/auth/logout
// @access  Private
export const logout = async (req, res) => {
    try {
        const token = req.headers.authorization?.split(' ')[1] || req.cookies.token;

        // Deactivate session
        await Session.findOneAndUpdate({ token }, { isActive: false });

        // Clear cookie
        res.cookie('token', '', {
            httpOnly: true,
            expires: new Date(0),
        });

        res.status(200).json({
            success: true,
            message: 'Logged out successfully',
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            message: error.message,
        });
    }
};

// @desc    Get current user
// @route   GET /api/auth/me
// @access  Private
export const getMe = async (req, res) => {
    try {
        const user = await User.findById(req.user.id);

        res.status(200).json({
            success: true,
            user: {
                id: user._id,
                name: user.name,
                email: user.email,
                role: user.role,
                lastLogin: user.lastLogin,
            },
        });
    } catch (error) {
        res.status(500).json({
            success: false,
            message: error.message,
        });
    }
};
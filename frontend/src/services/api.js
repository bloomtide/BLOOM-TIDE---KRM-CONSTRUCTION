import api from '../utils/api';

// ============================================
// AUTH SERVICES
// ============================================

/**
 * Login user
 * @param {Object} credentials - { email, password }
 * @returns {Promise} Response with token and user data
 */
export const login = async (credentials) => {
    const response = await api.post('/auth/login', credentials);
    return response.data;
};

/**
 * Logout current user
 * @returns {Promise} Response
 */
export const logout = async () => {
    const response = await api.post('/auth/logout');
    return response.data;
};

/**
 * Get current user
 * @returns {Promise} Current user data
 */
export const getCurrentUser = async () => {
    const response = await api.get('/auth/me');
    return response.data;
};

// ============================================
// USER SERVICES (Admin only)
// ============================================

/**
 * Get all users
 * @returns {Promise} List of all users
 */
export const getAllUsers = async () => {
    const response = await api.get('/users');
    return response.data;
};

/**
 * Create new user
 * @param {Object} userData - { name, email, password, role }
 * @returns {Promise} Created user data
 */
export const createUser = async (userData) => {
    const response = await api.post('/users', userData);
    return response.data;
};

/**
 * Get user by ID
 * @param {string} userId - User ID
 * @returns {Promise} User data
 */
export const getUserById = async (userId) => {
    const response = await api.get(`/users/${userId}`);
    return response.data;
};

/**
 * Update user
 * @param {string} userId - User ID
 * @param {Object} userData - Updated user data
 * @returns {Promise} Updated user data
 */
export const updateUser = async (userId, userData) => {
    const response = await api.put(`/users/${userId}`, userData);
    return response.data;
};

/**
 * Delete user
 * @param {string} userId - User ID
 * @returns {Promise} Response
 */
export const deleteUser = async (userId) => {
    const response = await api.delete(`/users/${userId}`);
    return response.data;
};

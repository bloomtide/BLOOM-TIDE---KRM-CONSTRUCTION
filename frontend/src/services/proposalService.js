import axios from 'axios';

const API_URL = import.meta.env.VITE_API_URL || 'http://localhost:3001/api';

// Configure axios defaults
axios.defaults.withCredentials = true;

export const proposalAPI = {
    /**
     * Create a new proposal
     * @param {FormData} formData - Form data containing Excel file and metadata
     * @returns {Promise} Response with created proposal
     */
    create: async (formData) => {
        const response = await axios.post(`${API_URL}/proposals`, formData, {
            headers: {
                'Content-Type': 'multipart/form-data',
            },
        });
        return response.data;
    },

    /**
     * Get all proposals with pagination and filters
     * @param {Object} params - Query parameters (page, limit, search, startDate, endDate)
     * @returns {Promise} Response with list of proposals and pagination info
     */
    list: async (params = {}) => {
        const response = await axios.get(`${API_URL}/proposals`, { params });
        return response.data;
    },

    /**
     * Get a single proposal by ID
     * @param {string} id - Proposal ID
     * @returns {Promise} Response with proposal data
     */
    getById: async (id) => {
        const response = await axios.get(`${API_URL}/proposals/${id}`);
        return response.data;
    },

    /**
     * Update a proposal
     * @param {string} id - Proposal ID
     * @param {Object} data - Data to update (name, client, project, spreadsheetJson)
     * @returns {Promise} Response with updated proposal
     */
    update: async (id, data) => {
        const response = await axios.put(`${API_URL}/proposals/${id}`, data);
        return response.data;
    },

    /**
     * Delete a proposal
     * @param {string} id - Proposal ID
     * @returns {Promise} Response confirming deletion
     */
    delete: async (id) => {
        const response = await axios.delete(`${API_URL}/proposals/${id}`);
        return response.data;
    },
    /**
     * Delete multiple proposals
     * @param {Array} ids - Array of Proposal IDs
     * @returns {Promise} Response confirming deletion
     */
    bulkDelete: async (ids) => {
        const response = await axios.post(`${API_URL}/proposals/bulk-delete`, { ids });
        return response.data;
    },
};

export default proposalAPI;

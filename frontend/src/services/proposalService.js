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

    /** Get raw Excel file as ArrayBuffer (for parsing client-side when rawExcelFileUrl is set) */
    getRawFile: async (id) => {
        const response = await axios.get(`${API_URL}/proposals/${id}/raw-file`, {
            responseType: 'arraybuffer',
        });
        return response.data;
    },

    /** Get spreadsheet JSON file (gzipped) as ArrayBuffer (decompress and parse client-side) */
    getSpreadsheetFile: async (id) => {
        const response = await axios.get(`${API_URL}/proposals/${id}/spreadsheet-file`, {
            responseType: 'arraybuffer',
        });
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

    /**
     * Update the status of an unused raw data row
     * @param {string} id - Proposal ID
     * @param {number} rowIndex - The index of the row to update
     * @param {boolean} isUsed - The new status
     * @returns {Promise} Response with updated proposal data
     */
    updateUnusedRowStatus: async (id, rowIndex, isUsed) => {
        const response = await axios.patch(`${API_URL}/proposals/${id}/unused-rows/${rowIndex}`, { isUsed });
        return response.data;
    },

    /**
     * Bulk update status of unused raw data rows
     * @param {string} id - Proposal ID
     * @param {Array} updates - Array of update objects { rowIndex, isUsed }
     * @returns {Promise} Response with updated proposal data
     */
    updateUnusedRowStatusBulk: async (id, updates) => {
        const response = await axios.patch(`${API_URL}/proposals/${id}/unused-rows/bulk`, { updates });
        return response.data;
    },
};

export default proposalAPI;

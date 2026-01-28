import axios from 'axios';

// Create axios instance with default config
const api = axios.create({
    baseURL: import.meta.env.VITE_API_URL || 'http://localhost:5000/api',
    headers: {
        'Content-Type': 'application/json',
    },
    withCredentials: true, // Important: This sends cookies automatically
});

// Response interceptor - handle errors globally
api.interceptors.response.use(
    (response) => response,
    (error) => {
        if (error.response?.status === 401) {
            // Unauthorized
            localStorage.removeItem('user');

            // Only redirect if we're not already on the login page
            // and if this wasn't a silent check
            if (window.location.pathname !== '/') {
                window.location.href = '/';
            }
        }
        return Promise.reject(error);
    }
);

export default api;

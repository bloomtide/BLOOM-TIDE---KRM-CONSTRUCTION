import React, { createContext, useState, useContext, useEffect } from 'react';
import Cookies from 'js-cookie';
import { login as loginAPI, logout as logoutAPI, getCurrentUser } from '../services/api';

const AuthContext = createContext(null);

export const AuthProvider = ({ children }) => {
    const [user, setUser] = useState(null);
    const [loading, setLoading] = useState(true);
    const [isAuthenticated, setIsAuthenticated] = useState(false);

    // Check if user is logged in on mount
    useEffect(() => {
        // Check if token cookie exists (even though we can't read httpOnly cookies,
        // we can try to verify with the backend)
        const savedUser = localStorage.getItem('user');

        if (savedUser) {
            setUser(JSON.parse(savedUser));
            setIsAuthenticated(true);
            // Verify token with backend
            verifyToken();
        } else {
            // Try to verify anyway in case cookie exists
            verifyToken();
        }
    }, []);

    const verifyToken = async () => {
        try {
            const response = await getCurrentUser();
            if (response.success) {
                setUser(response.user);
                setIsAuthenticated(true);
                // Save user data to localStorage for UI purposes
                localStorage.setItem('user', JSON.stringify(response.user));
            }
        } catch (error) {
            console.error('Token verification failed:', error);
            handleLogout();
        } finally {
            setLoading(false);
        }
    };

    const login = async (email, password) => {
        try {
            const response = await loginAPI({ email, password });

            if (response.success) {
                const { user } = response;
                // Note: Token is stored in httpOnly cookie by backend
                // We only save user data to localStorage for UI
                localStorage.setItem('user', JSON.stringify(user));

                // Update state
                setUser(user);
                setIsAuthenticated(true);

                return { success: true, user };
            }
        } catch (error) {
            console.error('Login Error Details:', error);

            let message = 'Login failed';

            if (error.response) {
                // Server responded with a status code
                console.error('Data:', error.response.data);
                console.error('Status:', error.response.status);
                message = error.response.data?.message || `Server Error: ${error.response.status}`;
            } else if (error.request) {
                // Request was made but no response received
                console.error('No response received:', error.request);
                message = 'Network Error: Cannot connect to server';
            } else {
                // Setup error
                message = error.message;
            }

            return { success: false, message };
        }
    };

    const logout = async () => {
        try {
            await logoutAPI();
        } catch (error) {
            console.error('Logout error:', error);
        } finally {
            handleLogout();
        }
    };

    const handleLogout = () => {
        // Clear user data from localStorage
        localStorage.removeItem('user');
        // Cookie will be cleared by backend
        setUser(null);
        setIsAuthenticated(false);
    };

    const value = {
        user,
        loading,
        isAuthenticated,
        login,
        logout,
    };

    return <AuthContext.Provider value={value}>{children}</AuthContext.Provider>;
};

// Custom hook to use auth context
export const useAuth = () => {
    const context = useContext(AuthContext);
    if (!context) {
        throw new Error('useAuth must be used within AuthProvider');
    }
    return context;
};

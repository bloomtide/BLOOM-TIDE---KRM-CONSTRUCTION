import React, { createContext, useState, useContext } from 'react';

const SidebarContext = createContext(null);

export const SidebarProvider = ({ children }) => {
    const [sidebarCollapsed, setSidebarCollapsed] = useState(false);

    const toggleSidebar = () => {
        setSidebarCollapsed(prev => !prev);
    };

    const value = {
        sidebarCollapsed,
        setSidebarCollapsed,
        toggleSidebar,
    };

    return <SidebarContext.Provider value={value}>{children}</SidebarContext.Provider>;
};

// Custom hook to use sidebar context
export const useSidebar = () => {
    const context = useContext(SidebarContext);
    if (!context) {
        throw new Error('useSidebar must be used within SidebarProvider');
    }
    return context;
};

export default SidebarContext;

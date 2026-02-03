import React, { useState, useEffect, useRef } from 'react'
import { FiBell, FiLogOut, FiUser } from 'react-icons/fi'
import { useAuth } from '../context/AuthContext'
import { useNavigate } from 'react-router-dom'

const TopBar = ({ leftContent, actionButtons }) => {
    const { user, logout } = useAuth();
    const navigate = useNavigate();
    const [showProfileMenu, setShowProfileMenu] = useState(false);
    const profileMenuRef = useRef(null);

    // Get user initials
    const getInitials = (name) => {
        if (!name) return 'U';
        return name
            .split(' ')
            .map((n) => n[0])
            .join('')
            .toUpperCase()
            .substring(0, 2);
    };

    // Close dropdown when clicking outside
    useEffect(() => {
        const handleClickOutside = (event) => {
            if (profileMenuRef.current && !profileMenuRef.current.contains(event.target)) {
                setShowProfileMenu(false);
            }
        };

        if (showProfileMenu) {
            document.addEventListener('mousedown', handleClickOutside);
        }

        return () => {
            document.removeEventListener('mousedown', handleClickOutside);
        };
    }, [showProfileMenu]);

    const handleLogout = async () => {
        await logout();
        navigate('/');
    };

    return (
        <div className="bg-white border-b border-gray-200 px-8 py-[18px]">
            <div className={`flex items-center ${leftContent ? 'justify-between' : 'justify-end'}`}>
                {/* Left Section: Custom content (e.g., proposal info) */}
                {leftContent && <div className="flex-1">{leftContent}</div>}
                
                {/* Right Section: Action Buttons and User */}
                <div className="flex items-center gap-3">
                    {/* Custom Action Buttons (e.g., Preview, Settings) */}
                    {actionButtons}
                    
                    {/* User Profile */}
                    <div className="relative" ref={profileMenuRef}>
                        <button
                            onClick={() => setShowProfileMenu(!showProfileMenu)}
                            className="flex items-center gap-3 bg-gray-100 rounded-full px-4 py-2 hover:bg-gray-200 transition-colors"
                        >
                            <div className="w-9 h-9 rounded-full overflow-hidden bg-gradient-to-br from-blue-500 to-blue-600 flex items-center justify-center text-white text-sm font-semibold">
                                {getInitials(user?.name)}
                            </div>
                            <span className="text-sm font-medium text-gray-900 pr-2">{user?.name || 'User'}</span>
                        </button>

                        {/* Dropdown Menu */}
                        {showProfileMenu && (
                            <div className="absolute right-0 mt-2 w-48 bg-white rounded-lg shadow-lg py-1 z-50 border border-gray-100">
                                <div className="px-4 py-2 border-b border-gray-100">
                                    <p className="text-sm font-medium text-gray-900">{user?.name}</p>
                                    <p className="text-xs text-gray-500 truncate">{user?.email}</p>
                                </div>
                                <button
                                    onClick={handleLogout}
                                    className="w-full text-left px-4 py-2 text-sm text-red-600 hover:bg-red-50 flex items-center gap-2"
                                >
                                    <FiLogOut size={16} />
                                    Sign out
                                </button>
                            </div>
                        )}
                    </div>
                </div>
            </div>
        </div>
    )
}

export default TopBar

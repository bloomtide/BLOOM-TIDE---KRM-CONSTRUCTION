import React from 'react'
import { Link, useLocation } from 'react-router-dom'
import { FiFileText, FiUsers, FiChevronLeft, FiChevronRight } from 'react-icons/fi'
import Logo from '../utils/logo/Logo.svg'

const Sidebar = ({ collapsed = false, onToggle }) => {
    const location = useLocation()

    const menuItems = [
        {
            name: 'Proposal',
            icon: FiFileText,
            path: '/proposals',
            active: location.pathname === '/proposals'
        },
        {
            name: 'Users',
            icon: FiUsers,
            path: '/users',
            active: location.pathname === '/users'
        }
    ]

    return (
        <div
            className={`bg-[#1E4976] min-h-screen transition-all duration-300 flex flex-col ${collapsed ? 'w-20' : 'w-64'
                }`}
        >
            {/* Logo Section */}
            <div className="p-6 flex items-center justify-between border-b border-[#2B5A8F]">
                {!collapsed && (
                    <img src={Logo} alt="KRM Construction" className="h-12 w-full object-contain" />
                )}
                {collapsed && (
                    <div className="w-full flex justify-center">
                        <img src={Logo} alt="KRM" className="h-12 w-auto" />
                    </div>
                )}
            </div>

            {/* Menu Section */}
            <div className="flex-1 py-6">
                {/* Menu Header */}
                <div className="px-6 mb-4 flex items-center justify-between">
                    {!collapsed && (
                        <span className="text-xs font-semibold text-white/60 tracking-wider">
                            MENU
                        </span>
                    )}
                    <button
                        onClick={onToggle}
                        className="p-1 rounded hover:bg-[#2B5A8F] transition-colors text-white/60 hover:text-white"
                        aria-label={collapsed ? 'Expand sidebar' : 'Collapse sidebar'}
                    >
                        {collapsed ? <FiChevronRight size={16} /> : <FiChevronLeft size={16} />}
                    </button>
                </div>

                {/* Menu Items */}
                <nav className="space-y-1 px-3">
                    {menuItems.map((item) => {
                        const Icon = item.icon
                        return (
                            <Link
                                key={item.path}
                                to={item.path}
                                className={`flex items-center gap-3 px-3 py-3 rounded-lg transition-all ${item.active
                                    ? 'bg-[#2B5A8F] text-white'
                                    : 'text-white/70 hover:bg-[#2B5A8F]/50 hover:text-white'
                                    }`}
                                title={collapsed ? item.name : ''}
                            >
                                <Icon size={20} className="flex-shrink-0" />
                                {!collapsed && (
                                    <span className="text-sm font-medium">{item.name}</span>
                                )}
                            </Link>
                        )
                    })}
                </nav>
            </div>
        </div>
    )
}

export default Sidebar

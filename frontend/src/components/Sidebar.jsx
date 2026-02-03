import React from 'react'
import { Link, useLocation } from 'react-router-dom'
import { FiFileText, FiUsers, FiChevronLeft, FiChevronRight } from 'react-icons/fi'
import Logo from '../utils/logo/Logo.svg'
import LogoImage from '../utils/logo/LogoImage.svg'
import { useAuth } from '../context/AuthContext'

const Sidebar = ({ collapsed = false, onToggle }) => {
    const location = useLocation()
    const { user } = useAuth()

    const menuItems = [
        {
            name: 'Proposals',
            icon: FiFileText,
            path: '/proposals',
            active: location.pathname === '/proposals' || location.pathname.startsWith('/proposals/')
        },
        ...(user?.role === 'admin' ? [{
            name: 'Users',
            icon: FiUsers,
            path: '/users',
            active: location.pathname === '/users'
        }] : [])
    ]

    return (
        <div
            className={`bg-[#1E4976] min-h-screen transition-all duration-300 flex flex-col ${collapsed ? 'w-20' : 'w-64'
                }`}
        >
            {/* Logo Section */}
            <div className="p-6 h-[88px] flex items-center justify-center border-b border-[#2B5A8F]">
                {!collapsed && (
                    <img src={Logo} alt="KRM Construction" className="h-12 w-full object-contain" />
                )}
                {collapsed && (
                    <img src={LogoImage} alt="KRM" className="h-12 w-12 object-contain" />
                )}
            </div>

            {/* Menu Section */}
            <div className="flex-1 py-6">
                {/* Menu Header */}
                <div className={`px-6 mb-4 flex items-center ${collapsed ? 'justify-center' : 'justify-between'}`}>
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
                                className={`flex items-center ${collapsed ? 'justify-center' : 'gap-3'} px-3 py-3 rounded-lg transition-all ${item.active
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

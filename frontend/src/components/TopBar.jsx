import React from 'react'
import { FiBell } from 'react-icons/fi'

const TopBar = () => {
    return (
        <div className="bg-white border-b border-gray-200 px-8 py-4">
            <div className="flex items-center justify-end">
                {/* Right Section: Notifications and User */}
                <div className="flex items-center gap-4">
                    {/* Notification Bell */}
                    <button className="relative w-12 h-12 rounded-full bg-gray-100 flex items-center justify-center text-gray-700 hover:bg-gray-200 transition-colors">
                        <FiBell size={20} />
                        <span className="absolute top-2 right-2 w-2 h-2 bg-red-500 rounded-full"></span>
                    </button>

                    {/* User Profile */}
                    <div className="flex items-center gap-3 bg-gray-100 rounded-full px-4 py-2">
                        <div className="w-9 h-9 rounded-full overflow-hidden bg-gradient-to-br from-blue-500 to-blue-600 flex items-center justify-center text-white text-sm font-semibold">
                            KA
                        </div>
                        <span className="text-sm font-medium text-gray-900 pr-2">KRM Admin</span>
                    </div>
                </div>
            </div>
        </div>
    )
}

export default TopBar

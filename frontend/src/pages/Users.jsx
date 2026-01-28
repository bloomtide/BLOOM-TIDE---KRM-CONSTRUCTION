import React, { useState } from 'react'
import { Link } from 'react-router-dom'
import Sidebar from '../components/Sidebar'
import TopBar from '../components/TopBar'
import DataTable from '../components/DataTable'
import AddUserModal from '../components/AddUserModal'
import { FiPlus, FiSearch, FiHome } from 'react-icons/fi'

const Users = () => {
    const [sidebarCollapsed, setSidebarCollapsed] = useState(false)
    const [isModalOpen, setIsModalOpen] = useState(false)

    // Sample user data matching the design
    const userData = [
        {
            id: 1,
            name: 'John Doe',
            email: 'johndoe23@gmail.com'
        },
        {
            id: 2,
            name: 'John Doe',
            email: 'johndoe23@gmail.com'
        },
        {
            id: 3,
            name: 'John Doe',
            email: 'johndoe23@gmail.com'
        }
    ]

    // Table columns configuration
    const columns = [
        {
            header: 'Name',
            key: 'name',
        },
        {
            header: 'Email',
            key: 'email'
        }
    ]

    // Action handlers
    const handleEdit = (row) => {
        console.log('Edit user:', row)
    }

    const handleDelete = (row) => {
        console.log('Delete user:', row)
    }

    const handleAddUser = (userData) => {
        console.log('New user added:', userData)
        // Add logic to handle the new user data
    }

    return (
        <div className="flex h-screen bg-gray-50">
            {/* Sidebar */}
            <Sidebar
                collapsed={sidebarCollapsed}
                onToggle={() => setSidebarCollapsed(!sidebarCollapsed)}
            />

            {/* Main Content */}
            <div className="flex-1 flex flex-col overflow-hidden">
                {/* Top Bar */}
                <TopBar />

                {/* Page Content */}
                <div className="flex-1 overflow-auto">
                    <div className="p-8">
                        {/* Breadcrumb Navigation */}
                        <div className="flex items-center gap-2 text-sm mb-6">
                            <Link to="/dashboard" className="text-gray-500 hover:text-gray-700">
                                <FiHome size={18} />
                            </Link>
                            <span className="text-gray-400">&gt;</span>
                            <span className="text-gray-900 font-medium">Users</span>
                        </div>
                        {/* Page Header */}
                        <div className="flex items-start justify-between mb-6">
                            <div>
                                <h1 className="text-2xl font-bold text-gray-900 mb-2">Users</h1>
                                <p className="text-sm text-gray-500">
                                    View and manage users in your organization.
                                </p>
                            </div>
                            <button
                                onClick={() => setIsModalOpen(true)}
                                className="flex items-center gap-2 bg-[#1A72B9] hover:bg-[#1565C0] text-white px-4 py-2.5 rounded-lg transition-colors text-sm font-medium shadow-sm"
                            >
                                <FiPlus size={18} />
                                Add User
                            </button>
                        </div>

                        {/* Search Section */}
                        <div className="bg-white rounded-lg border border-gray-200 p-4 mb-6">
                            <div className="flex items-center gap-4">
                                {/* Search Input */}
                                <div className="flex-1 relative">
                                    <FiSearch className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
                                    <input
                                        type="text"
                                        placeholder="Search here"
                                        className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
                                    />
                                </div>
                            </div>
                        </div>

                        {/* Data Table */}
                        <DataTable
                            columns={columns}
                            data={userData}
                            onEdit={handleEdit}
                            onDelete={handleDelete}
                            showCheckbox={true}
                            showActions={true}
                        />
                    </div>
                </div>
            </div>

            {/* Add User Modal */}
            <AddUserModal
                isOpen={isModalOpen}
                onClose={() => setIsModalOpen(false)}
                onSubmit={handleAddUser}
            />
        </div>
    )
}

export default Users

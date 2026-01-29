import React, { useState, useEffect } from 'react'
import { Link } from 'react-router-dom'
import { toast } from 'react-hot-toast'
import Sidebar from '../components/Sidebar'
import TopBar from '../components/TopBar'
import DataTable from '../components/DataTable'
import UserModal from '../components/UserModal'
import { FiPlus, FiSearch, FiHome } from 'react-icons/fi'
import { getAllUsers, createUser, updateUser, deleteUser } from '../services/api'

const Users = () => {
    const [sidebarCollapsed, setSidebarCollapsed] = useState(false)
    const [isModalOpen, setIsModalOpen] = useState(false)
    const [users, setUsers] = useState([])
    const [loading, setLoading] = useState(true)
    const [actionLoading, setActionLoading] = useState(false)
    const [searchTerm, setSearchTerm] = useState('')

    // Track who we are editing
    const [selectedUser, setSelectedUser] = useState(null)

    // Fetch users on mount
    useEffect(() => {
        fetchUsers()
    }, [])

    const fetchUsers = async () => {
        try {
            setLoading(true)
            const response = await getAllUsers()
            if (response.success) {
                setUsers(response.users)
            } else {
                toast.error('Failed to fetch users')
            }
        } catch (error) {
            console.error('Error fetching users:', error)
            toast.error('Error loading users')
        } finally {
            setLoading(false)
        }
    }

    // Filter users based on search and exclude admin users
    const filteredUsers = users
        .filter(user => user.role !== 'admin') // Don't show admin users
        .filter(user =>
            user.name.toLowerCase().includes(searchTerm.toLowerCase()) ||
            user.email.toLowerCase().includes(searchTerm.toLowerCase())
        )

    // Table columns configuration
    const columns = [
        {
            header: 'Name',
            key: 'name',
            className: 'w-1/2',
            render: (row) => (
                <div className="font-medium text-gray-900">{row.name}</div>
            )
        },
        {
            header: 'Email',
            key: 'email',
            className: 'w-1/2',
        }
    ]

    // OPEN MODAL: For New User
    const openAddModal = () => {
        setSelectedUser(null)
        setIsModalOpen(true)
    }

    // OPEN MODAL: For Edit User
    const handleEdit = (row) => {
        setSelectedUser(row)
        setIsModalOpen(true)
    }

    // SAVE USER (Handles both Add and Edit)
    const handleSaveUser = async (formData) => {
        try {
            setActionLoading(true)

            let response;
            if (selectedUser) {
                // UPDATE EXISTING
                response = await updateUser(selectedUser.id || selectedUser._id, formData)
            } else {
                // CREATE NEW
                response = await createUser(formData)
            }

            if (response.success) {
                toast.success(selectedUser ? 'User updated successfully' : 'User added successfully')
                setIsModalOpen(false)
                fetchUsers() // Refresh list
            } else {
                toast.error(response.message || (selectedUser ? 'Failed to update user' : 'Failed to add user'))
            }
        } catch (error) {
            const msg = error.response?.data?.message || 'Error saving user';
            toast.error(msg)
        } finally {
            setActionLoading(false)
        }
    }

    const handleDelete = async (row) => {
        if (!window.confirm(`Are you sure you want to delete ${row.name}?`)) return

        try {
            const response = await deleteUser(row._id || row.id)
            if (response.success) {
                toast.success('User deleted successfully')
                fetchUsers()
            } else {
                toast.error('Failed to delete user')
            }
        } catch (error) {
            toast.error('Error deleting user')
        }
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
                                onClick={openAddModal}
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
                                        placeholder="Search by name or email"
                                        value={searchTerm}
                                        onChange={(e) => setSearchTerm(e.target.value)}
                                        className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
                                    />
                                </div>
                            </div>
                        </div>

                        {/* Data Table */}
                        {loading ? (
                            <div className="flex justify-center items-center h-64">
                                <div className="animate-spin rounded-full h-12 w-12 border-b-2 border-blue-500"></div>
                            </div>
                        ) : (
                            <DataTable
                                columns={columns}
                                data={filteredUsers}
                                onEdit={handleEdit}
                                onDelete={handleDelete}
                                showCheckbox={true}
                                showActions={true}
                            />
                        )}
                    </div>
                </div>
            </div>

            {/* Shared User Modal (Add/Edit) */}
            <UserModal
                isOpen={isModalOpen}
                onClose={() => setIsModalOpen(false)}
                onSubmit={handleSaveUser}
                loading={actionLoading}
                initialData={selectedUser}
                isEditing={!!selectedUser}
            />
        </div>
    )
}

export default Users

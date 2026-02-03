import React, { useState, useEffect } from 'react'
import { toast } from 'react-hot-toast'
import Sidebar from '../components/Sidebar'
import TopBar from '../components/TopBar'
import DataTable from '../components/DataTable'
import UserModal from '../components/UserModal'
import { FiPlus, FiSearch } from 'react-icons/fi'
import { getAllUsers, createUser, updateUser, deleteUser, bulkDeleteUsers } from '../services/api'
import { useSidebar } from '../context/SidebarContext'

const Users = () => {
    const { sidebarCollapsed, toggleSidebar } = useSidebar()
    const [isModalOpen, setIsModalOpen] = useState(false)
    const [pagination, setPagination] = useState({
        page: 1,
        limit: 10,
        totalPages: 0,
        totalItems: 0
    })
    const [debouncedSearch, setDebouncedSearch] = useState('')
    const [searchTerm, setSearchTerm] = useState('')
    const [users, setUsers] = useState([])
    const [loading, setLoading] = useState(true)
    const [actionLoading, setActionLoading] = useState(false)
    const [selectedUser, setSelectedUser] = useState(null)

    // Debounce search query
    useEffect(() => {
        const timer = setTimeout(() => {
            setDebouncedSearch(searchTerm)
            setPagination(prev => ({ ...prev, page: 1 }))
        }, 500)

        return () => clearTimeout(timer)
    }, [searchTerm])

    // Fetch users on mount
    useEffect(() => {
        fetchUsers()
    }, [debouncedSearch, pagination.page, pagination.limit])

    const fetchUsers = async () => {
        try {
            setLoading(true)
            const params = {
                page: pagination.page,
                limit: pagination.limit,
                search: debouncedSearch
            }
            const response = await getAllUsers(params)

            if (response.success) {
                setUsers(response.users)
                if (response.pagination) {
                    setPagination(prev => ({
                        ...prev,
                        totalPages: response.pagination.pages,
                        totalItems: response.pagination.total
                    }))
                }
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

    const handlePageChange = (newPage) => {
        setPagination(prev => ({ ...prev, page: newPage }))
    }

    // Filter out admin users from display (if needed, though ideally backend handles strict filtering)
    // For now we'll just display what the backend sends, assuming backend sends all. 
    // If we need to filter admins out client side AFTER pagination, it messes up pagination counts.
    // Ideally backend should have a filter for 'role' if we want to hide admins.
    // Based on previous code: .filter(user => user.role !== 'admin')
    // We will keep this visual filter but note it might make page size seem smaller than limit.
    const filteredUsers = users.filter(user => user.role !== 'admin')

    // Table columns configuration
    const columns = [
        {
            header: 'Name',
            key: 'name',
            className: 'w-1/2',
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

    const handleBulkDelete = async (selectedUsers) => {
        if (!window.confirm(`Are you sure you want to delete ${selectedUsers.length} users?`)) return

        try {
            const ids = selectedUsers.map(u => u._id || u.id)
            const response = await bulkDeleteUsers(ids)
            if (response.success) {
                toast.success('Users deleted successfully')
                fetchUsers()
            } else {
                toast.error('Failed to delete users')
            }
        } catch (error) {
            toast.error('Error deleting users')
        }
    }

    return (
        <div className="flex h-screen bg-gray-50">
            {/* Sidebar */}
            <Sidebar
                collapsed={sidebarCollapsed}
                onToggle={toggleSidebar}
            />

            {/* Main Content */}
            <div className="flex-1 flex flex-col overflow-hidden">
                {/* Top Bar */}
                <TopBar />

                {/* Page Content */}
                <div className="flex-1 overflow-auto">
                    <div className="p-8">
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
                                onBulkDelete={handleBulkDelete}
                                showCheckbox={true}
                                showActions={true}
                                pagination={{
                                    currentPage: pagination.page,
                                    totalPages: pagination.totalPages,
                                    onPageChange: handlePageChange,
                                    totalItems: pagination.totalItems,
                                    itemsPerPage: pagination.limit
                                }}
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

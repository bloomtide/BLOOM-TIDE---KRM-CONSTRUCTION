import React, { useState, useEffect } from 'react'
import Modal from './Modal'

const UserModal = ({ isOpen, onClose, onSubmit, loading, initialData, isEditing }) => {
    const [formData, setFormData] = useState({
        name: '',
        email: '',
        password: '',
        role: 'user', // Default role
        isActive: true
    })

    // Populate form when initialData changes (for editing)
    useEffect(() => {
        if (isOpen && initialData) {
            setFormData({
                name: initialData.name || '',
                email: initialData.email || '',
                password: '', // Keep empty unless changing
                role: initialData.role || 'user',
                isActive: initialData.isActive ?? true
            })
        } else if (isOpen) {
            // Reset for new user
            setFormData({
                name: '',
                email: '',
                password: '',
                role: 'user',
                isActive: true
            })
        }
    }, [isOpen, initialData])

    const handleInputChange = (e) => {
        const { name, value, type, checked } = e.target
        const val = type === 'checkbox' ? checked : value
        setFormData(prev => ({ ...prev, [name]: val }))
    }

    const handleSubmit = (e) => {
        e.preventDefault()
        if (onSubmit) {
            onSubmit(formData)
        }
    }

    return (
        <Modal
            isOpen={isOpen}
            onClose={onClose}
            title={isEditing ? "Edit User" : "New User"}
            subtitle={isEditing ? "Update user details." : "Add a user to your organization."}
            maxWidth="max-w-md"
        >
            <form onSubmit={handleSubmit}>
                {/* Name */}
                <div className="mb-4">
                    <label className="block text-sm font-medium text-gray-700 mb-1">Name*</label>
                    <input
                        type="text"
                        name="name"
                        value={formData.name}
                        onChange={handleInputChange}
                        required
                        className="w-full px-3 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm border-gray-300"
                    />
                </div>

                {/* Email */}
                <div className="mb-4">
                    <label className="block text-sm font-medium text-gray-700 mb-1">Email*</label>
                    <input
                        type="email"
                        name="email"
                        value={formData.email}
                        onChange={handleInputChange}
                        required
                        className="w-full px-3 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm border-gray-300"
                    />
                </div>

                {/* Password (Optional in Edit) */}
                <div className="mb-4">
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                        {isEditing ? 'Password (leave blank to keep current)' : 'Password*'}
                    </label>
                    <input
                        type="password"
                        name="password"
                        value={formData.password}
                        onChange={handleInputChange}
                        required={!isEditing}
                        placeholder={isEditing ? "********" : "Create Password"}
                        className="w-full px-3 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm border-gray-300"
                    />
                </div>



                {/* Status Checkbox (Only for Edit) */}
                {isEditing && (
                    <div className="mb-6 flex items-center gap-2">
                        <input
                            type="checkbox"
                            name="isActive"
                            checked={formData.isActive}
                            onChange={handleInputChange}
                            id="isActive"
                            className="w-4 h-4 text-blue-600 rounded"
                        />
                        <label htmlFor="isActive" className="text-sm text-gray-700">Active Account</label>
                    </div>
                )}

                <div className="flex justify-end gap-3 mt-6">
                    <button
                        type="button"
                        onClick={onClose}
                        disabled={loading}
                        className="px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors disabled:opacity-50"
                    >
                        Cancel
                    </button>
                    <button
                        type="submit"
                        disabled={loading}
                        className="px-4 py-2 text-sm font-medium text-white bg-[#1A72B9] hover:bg-[#1565C0] rounded-lg transition-colors disabled:opacity-50"
                    >
                        {loading ? 'Saving...' : (isEditing ? 'Save Changes' : 'Add User')}
                    </button>
                </div>
            </form>
        </Modal>
    )
}

export default UserModal

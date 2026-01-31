import React, { useState, useEffect } from 'react'
import { FiEye, FiEyeOff } from 'react-icons/fi'
import Modal from './Modal'

const UserModal = ({ isOpen, onClose, onSubmit, loading, initialData, isEditing }) => {
    const [formData, setFormData] = useState({
        name: '',
        email: '',
        password: '',
        role: 'user', // Default role
    })

    const [showPassword, setShowPassword] = useState(false)
    const [errors, setErrors] = useState({})

    // Populate form when initialData changes (for editing)
    useEffect(() => {
        if (isOpen && initialData) {
            setFormData({
                name: initialData.name || '',
                email: initialData.email || '',
                password: '', // Keep empty unless changing
                role: initialData.role || 'user',
            })
        } else if (isOpen) {
            // Reset for new user - all fields empty
            setFormData({
                name: '',
                email: '',
                password: '',
                role: 'user',
            })
        }
        // Reset errors and password visibility when modal opens/closes
        setErrors({})
        setShowPassword(false)
    }, [isOpen, initialData])

    const validateForm = () => {
        const newErrors = {}

        // Name validation
        if (!formData.name.trim()) {
            newErrors.name = 'Name is required'
        } else if (formData.name.trim().length < 2) {
            newErrors.name = 'Name must be at least 2 characters'
        }

        // Email validation
        const emailRegex = /^\S+@\S+\.\S+$/
        if (!formData.email.trim()) {
            newErrors.email = 'Email is required'
        } else if (!emailRegex.test(formData.email)) {
            newErrors.email = 'Please enter a valid email'
        }

        // Password validation
        if (!isEditing) {
            // For new users, password is required
            if (!formData.password) {
                newErrors.password = 'Password is required'
            } else if (formData.password.length < 6) {
                newErrors.password = 'Password must be at least 6 characters'
            }
        } else {
            // For editing, only validate if password is provided
            if (formData.password && formData.password.length < 6) {
                newErrors.password = 'Password must be at least 6 characters'
            }
        }

        setErrors(newErrors)
        return Object.keys(newErrors).length === 0
    }

    const handleInputChange = (e) => {
        const { name, value } = e.target
        setFormData(prev => ({ ...prev, [name]: value }))
        // Clear error for this field when user starts typing
        if (errors[name]) {
            setErrors(prev => ({ ...prev, [name]: '' }))
        }
    }

    const handleSubmit = (e) => {
        e.preventDefault()
        if (validateForm() && onSubmit) {
            // Remove password from submission if it's empty during edit
            const dataToSubmit = { ...formData }
            if (isEditing && !dataToSubmit.password) {
                delete dataToSubmit.password
            }
            onSubmit(dataToSubmit)
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
            <form onSubmit={handleSubmit} autoComplete="off">
                {/* Name */}
                <div className="mb-4">
                    <label className="block text-sm font-medium text-gray-700 mb-1">Name*</label>
                    <input
                        type="text"
                        name="name"
                        value={formData.name}
                        onChange={handleInputChange}
                        autoComplete="off"
                        className={`w-full px-3 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm ${
                            errors.name ? 'border-red-500' : 'border-gray-300'
                        }`}
                    />
                    {errors.name && (
                        <p className="mt-1 text-xs text-red-600">{errors.name}</p>
                    )}
                </div>

                {/* Email */}
                <div className="mb-4">
                    <label className="block text-sm font-medium text-gray-700 mb-1">Email*</label>
                    <input
                        type="email"
                        name="email"
                        value={formData.email}
                        onChange={handleInputChange}
                        autoComplete="off"
                        className={`w-full px-3 py-2 border rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm ${
                            errors.email ? 'border-red-500' : 'border-gray-300'
                        }`}
                    />
                    {errors.email && (
                        <p className="mt-1 text-xs text-red-600">{errors.email}</p>
                    )}
                </div>

                {/* Password (Optional in Edit) */}
                <div className="mb-6">
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                        {isEditing ? 'Password (leave blank to keep current)' : 'Password*'}
                    </label>
                    <div className="relative">
                        <input
                            type={showPassword ? 'text' : 'password'}
                            name="password"
                            value={formData.password}
                            onChange={handleInputChange}
                            autoComplete="new-password"
                            placeholder={isEditing ? "Enter new password" : "Create password"}
                            className={`w-full px-3 py-2 pr-10 border rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm ${
                                errors.password ? 'border-red-500' : 'border-gray-300'
                            }`}
                        />
                        <button
                            type="button"
                            onClick={() => setShowPassword(!showPassword)}
                            className="absolute right-3 top-1/2 -translate-y-1/2 text-gray-400 hover:text-gray-600"
                        >
                            {showPassword ? <FiEyeOff size={18} /> : <FiEye size={18} />}
                        </button>
                    </div>
                    {errors.password && (
                        <p className="mt-1 text-xs text-red-600">{errors.password}</p>
                    )}
                </div>

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

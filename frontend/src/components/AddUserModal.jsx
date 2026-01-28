import React, { useState } from 'react'
import Modal from './Modal'

const AddUserModal = ({ isOpen, onClose, onSubmit, loading }) => {
    const [formData, setFormData] = useState({
        name: '',
        email: '',
        password: ''
    })

    const handleInputChange = (e) => {
        const { name, value } = e.target
        setFormData(prev => ({ ...prev, [name]: value }))
    }

    const handleSubmit = (e) => {
        e.preventDefault()
        if (onSubmit) {
            onSubmit(formData)
        }
        // Reset form
        setFormData({
            name: '',
            email: '',
            password: ''
        })
        onClose()
    }

    return (
        <Modal
            isOpen={isOpen}
            onClose={onClose}
            title="New User"
            subtitle="Add a user to your organization."
            maxWidth="max-w-md"
        >
            <form onSubmit={handleSubmit}>
                {/* Name Field */}
                <div className="mb-4">
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                        Name*
                    </label>
                    <input
                        type="text"
                        name="name"
                        value={formData.name}
                        onChange={handleInputChange}
                        placeholder="Enter Employee Name"
                        required
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
                    />
                </div>

                {/* Email Field */}
                <div className="mb-4">
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                        Email*
                    </label>
                    <input
                        type="email"
                        name="email"
                        value={formData.email}
                        onChange={handleInputChange}
                        placeholder="Enter Employee Email"
                        required
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
                    />
                </div>

                {/* Password Field */}
                <div className="mb-6">
                    <label className="block text-sm font-medium text-gray-700 mb-1">
                        Password*
                    </label>
                    <input
                        type="password"
                        name="password"
                        value={formData.password}
                        onChange={handleInputChange}
                        placeholder="Create Password"
                        required
                        className="w-full px-3 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
                    />
                </div>

                {/* Action Buttons */}
                <div className="flex justify-end gap-3">
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
                        className="px-4 py-2 text-sm font-medium text-white bg-[#1A72B9] hover:bg-[#1565C0] rounded-lg transition-colors disabled:opacity-50 flex items-center gap-2"
                    >
                        {loading ? 'Adding...' : 'Add User'}
                    </button>
                </div>
            </form>
        </Modal>
    )
}

export default AddUserModal

import React, { useState, useEffect } from 'react'
import Modal from './Modal'

const ProposalSettingsModal = ({ isOpen, onClose, proposal, onSave }) => {
    const [formData, setFormData] = useState({
        name: '',
        client: '',
        project: '',
    })
    const [isSaving, setIsSaving] = useState(false)

    useEffect(() => {
        if (proposal) {
            setFormData({
                name: proposal.name || '',
                client: proposal.client || '',
                project: proposal.project || '',
            })
        }
    }, [proposal])

    const handleInputChange = (e) => {
        const { name, value } = e.target
        setFormData(prev => ({ ...prev, [name]: value }))
    }

    const handleSubmit = async (e) => {
        e.preventDefault()
        setIsSaving(true)
        
        try {
            await onSave(formData)
            onClose()
        } catch (error) {
            console.error('Error saving settings:', error)
        } finally {
            setIsSaving(false)
        }
    }

    return (
        <Modal
            isOpen={isOpen}
            onClose={onClose}
            title="Proposal Settings"
            subtitle="Edit proposal information"
            maxWidth="max-w-2xl"
        >
            <form onSubmit={handleSubmit}>
                <div className="space-y-5">
                    {/* Proposal Name */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Proposal Name*
                        </label>
                        <input
                            type="text"
                            name="name"
                            value={formData.name}
                            onChange={handleInputChange}
                            placeholder="Enter proposal name"
                            required
                            className="w-full px-4 py-2.5 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
                        />
                    </div>

                    {/* Client */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Client*
                        </label>
                        <input
                            type="text"
                            name="client"
                            value={formData.client}
                            onChange={handleInputChange}
                            placeholder="Enter client name"
                            required
                            className="w-full px-4 py-2.5 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
                        />
                    </div>

                    {/* Project */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Project*
                        </label>
                        <input
                            type="text"
                            name="project"
                            value={formData.project}
                            onChange={handleInputChange}
                            placeholder="Enter project name"
                            required
                            className="w-full px-4 py-2.5 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
                        />
                    </div>
                </div>

                {/* Action Buttons */}
                <div className="flex justify-end gap-3 mt-6">
                    <button
                        type="button"
                        onClick={onClose}
                        disabled={isSaving}
                        className="px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors disabled:opacity-50"
                    >
                        Cancel
                    </button>
                    <button
                        type="submit"
                        disabled={isSaving}
                        className="px-4 py-2 text-sm font-medium text-white bg-[#1A72B9] hover:bg-[#1565C0] rounded-lg transition-colors disabled:opacity-50"
                    >
                        {isSaving ? 'Saving...' : 'Save Changes'}
                    </button>
                </div>
            </form>
        </Modal>
    )
}

export default ProposalSettingsModal

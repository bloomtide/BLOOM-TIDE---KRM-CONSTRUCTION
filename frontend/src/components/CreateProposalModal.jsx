import React, { useState } from 'react'
import Modal from './Modal'
import { FiUploadCloud } from 'react-icons/fi'
import { DatePicker } from './DatePicker'

const CreateProposalModal = ({ isOpen, onClose, onSubmit }) => {
    const [formData, setFormData] = useState({
        proposalName: '',
        client: '',
        project: '',
        address: '',
        email: '',
        phone: '',
        estimateNo: '',
        proposalDate: '',
        structuralDate: '',
        architecturalDate: '',
        template: ''
    })
    const [file, setFile] = useState(null)
    const [dragActive, setDragActive] = useState(false)

    const handleInputChange = (e) => {
        const { name, value } = e.target
        setFormData(prev => ({ ...prev, [name]: value }))
    }

    const handleDrag = (e) => {
        e.preventDefault()
        e.stopPropagation()
        if (e.type === 'dragenter' || e.type === 'dragover') {
            setDragActive(true)
        } else if (e.type === 'dragleave') {
            setDragActive(false)
        }
    }

    const handleDrop = (e) => {
        e.preventDefault()
        e.stopPropagation()
        setDragActive(false)

        if (e.dataTransfer.files && e.dataTransfer.files[0]) {
            const droppedFile = e.dataTransfer.files[0]
            if (droppedFile.name.endsWith('.xlsx') || droppedFile.name.endsWith('.xls')) {
                setFile(droppedFile)
            } else {
                alert('Please upload an Excel file (.xlsx)')
            }
        }
    }

    const handleFileChange = (e) => {
        if (e.target.files && e.target.files[0]) {
            setFile(e.target.files[0])
        }
    }

    const handleSubmit = (e) => {
        e.preventDefault()
        if (onSubmit) {
            onSubmit({ ...formData, file })
        }
        // Reset form
        setFormData({
            proposalName: '',
            client: '',
            project: '',
            address: '',
            email: '',
            phone: '',
            estimateNo: '',
            proposalDate: '',
            structuralDate: '',
            architecturalDate: '',
            template: ''
        })
        setFile(null)
        onClose()
    }

    return (
        <Modal
            isOpen={isOpen}
            onClose={onClose}
            title="New Proposal"
            subtitle="Upload your Excel sheet to generate and review the proposal data."
            maxWidth="max-w-4xl"
        >
            <form onSubmit={handleSubmit}>
                {/* File Upload Area */}
                <div
                    className={`border-2 border-dashed rounded-lg p-12 mb-8 text-center transition-colors ${dragActive
                        ? 'border-blue-500 bg-blue-50'
                        : 'border-gray-300 bg-gray-50'
                        }`}
                    onDragEnter={handleDrag}
                    onDragLeave={handleDrag}
                    onDragOver={handleDrag}
                    onDrop={handleDrop}
                >
                    <div className="flex flex-col items-center">
                        <div className="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center mb-4">
                            <FiUploadCloud className="text-blue-600" size={28} />
                        </div>
                        <p className="text-sm text-gray-600 mb-1">
                            <label htmlFor="file-upload" className="text-blue-600 hover:text-blue-700 cursor-pointer font-medium">
                                Click to Upload
                            </label>
                            {' '}or drag and drop
                        </p>
                        <p className="text-xs text-gray-500">.xlsx (Max. size: 20 MB)</p>
                        {file && (
                            <p className="mt-3 text-sm text-gray-700 font-medium">
                                Selected: {file.name}
                            </p>
                        )}
                        <input
                            id="file-upload"
                            type="file"
                            className="hidden"
                            accept=".xlsx,.xls"
                            onChange={handleFileChange}
                        />
                    </div>
                </div>

                {/* Form Fields - Two Columns */}
                <div className="grid grid-cols-2 gap-x-6 gap-y-5 mb-8">
                    {/* Proposal Name */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Proposal Name*
                        </label>
                        <input
                            type="text"
                            name="proposalName"
                            value={formData.proposalName}
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

                    {/* Address */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Address*
                        </label>
                        <input
                            type="text"
                            name="address"
                            value={formData.address}
                            onChange={handleInputChange}
                            placeholder="Enter Address"
                            required
                            className="w-full px-4 py-2.5 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
                        />
                    </div>

                    {/* Email */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Email*
                        </label>
                        <input
                            type="email"
                            name="email"
                            value={formData.email}
                            onChange={handleInputChange}
                            placeholder="Enter email"
                            required
                            className="w-full px-4 py-2.5 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
                        />
                    </div>

                    {/* Phone */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Phone*
                        </label>
                        <div className="flex gap-2">
                            <select className="w-20 px-2 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 text-sm">
                                <option>🇺🇸 +1</option>
                            </select>
                            <input
                                type="tel"
                                name="phone"
                                value={formData.phone}
                                onChange={handleInputChange}
                                placeholder="000-000-0000"
                                required
                                className="flex-1 px-4 py-2.5 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
                            />
                        </div>
                    </div>

                    {/* Estimate No */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Estimate No.*
                        </label>
                        <input
                            type="text"
                            name="estimateNo"
                            value={formData.estimateNo}
                            onChange={handleInputChange}
                            placeholder="Enter Estimate No."
                            required
                            className="w-full px-4 py-2.5 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
                        />
                    </div>

                    {/* Structural Date */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Structural Date*
                        </label>
                        <DatePicker
                            value={formData.structuralDate}
                            onChange={(date) => setFormData(prev => ({ ...prev, structuralDate: date }))}
                            placeholder="Select structural date"
                        />
                    </div>

                    {/* Architectural Date */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Architectural Date*
                        </label>
                        <DatePicker
                            value={formData.architecturalDate}
                            onChange={(date) => setFormData(prev => ({ ...prev, architecturalDate: date }))}
                            placeholder="Select architectural date"
                        />
                    </div>

                    {/* Proposal Date */}
                    <div>
                        <label className="block text-sm font-medium text-gray-700 mb-2">
                            Proposal Date*
                        </label>
                        <DatePicker
                            value={formData.proposalDate}
                            onChange={(date) => setFormData(prev => ({ ...prev, proposalDate: date }))}
                            placeholder="Select proposal date"
                        />
                    </div>
                </div>

                {/* Template - Full Width */}
                <div className="mb-6">
                    <label className="block text-sm font-medium text-gray-700 mb-2">
                        Template*
                    </label>
                    <select
                        name="template"
                        value={formData.template}
                        onChange={handleInputChange}
                        required
                        className="w-full px-4 py-2.5 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
                    >
                        <option value="">Select Template</option>
                        <option value="template1">Template 1</option>
                        <option value="template2">Template 2</option>
                        <option value="template3">Template 3</option>
                    </select>
                </div>

                {/* Action Buttons */}
                <div className="flex justify-end gap-3">
                    <button
                        type="button"
                        onClick={onClose}
                        className="px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors"
                    >
                        Cancel
                    </button>
                    <button
                        type="submit"
                        className="px-4 py-2 text-sm font-medium text-white bg-[#1A72B9] hover:bg-[#1565C0] rounded-lg transition-colors"
                    >
                        Create Proposal
                    </button>
                </div>
            </form>
        </Modal>
    )
}

export default CreateProposalModal

import React, { useState } from 'react'
import { useNavigate } from 'react-router-dom'
import * as XLSX from 'xlsx'
import toast from 'react-hot-toast'
import Modal from './Modal'
import { FiUploadCloud, FiArrowLeft, FiArrowRight } from 'react-icons/fi'
import { proposalAPI } from '../services/proposalService'

const CreateProposalModal = ({ isOpen, onClose, onSuccess }) => {
    const navigate = useNavigate()
    const [step, setStep] = useState(1) // 1 = Form, 2 = Preview
    const [formData, setFormData] = useState({
        proposalName: '',
        client: '',
        project: '',
        template: 'capstone'
    })
    const [file, setFile] = useState(null)
    const [dragActive, setDragActive] = useState(false)
    const [previewData, setPreviewData] = useState(null)
    const [isCreating, setIsCreating] = useState(false)

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
                toast.error('Please upload an Excel file (.xlsx or .xls)')
            }
        }
    }

    const handleFileChange = (e) => {
        if (e.target.files && e.target.files[0]) {
            setFile(e.target.files[0])
        }
    }

    const handleNextStep = (e) => {
        e.preventDefault()

        // Validate form data
        if (!formData.proposalName || !formData.client || !formData.project || !formData.template) {
            toast.error('Please fill in all required fields')
            return
        }

        if (!file) {
            toast.error('Please upload an Excel file')
            return
        }

        // Parse Excel file to show preview
        const reader = new FileReader()
        reader.onload = (e) => {
            try {
                const data = new Uint8Array(e.target.result)
                const workbook = XLSX.read(data, { type: 'array' })
                const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
                const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' })

                const headers = jsonData[0] || []
                const rows = jsonData.slice(1)

                setPreviewData({
                    fileName: file.name,
                    sheetName: workbook.SheetNames[0],
                    headers,
                    rows
                })
                setStep(2)
            } catch (error) {
                console.error('Error parsing Excel file:', error)
                toast.error('Error parsing Excel file. Please make sure it is a valid Excel file.')
            }
        }

        reader.onerror = () => {
            toast.error('Error reading file. Please try again.')
        }

        reader.readAsArrayBuffer(file)
    }

    const handleCreateProposal = async () => {
        setIsCreating(true)

        try {
            const formDataToSend = new FormData()
            formDataToSend.append('name', formData.proposalName)
            formDataToSend.append('client', formData.client)
            formDataToSend.append('project', formData.project)
            formDataToSend.append('template', formData.template)
            formDataToSend.append('excelFile', file)

            const response = await proposalAPI.create(formDataToSend)

            toast.success('Proposal created successfully!')

            // Reset form
            setFormData({
                proposalName: '',
                client: '',
                project: '',
                template: 'capstone'
            })
            setFile(null)
            setStep(1)
            setPreviewData(null)

            onClose()

            if (onSuccess) {
                onSuccess()
            }

            // Navigate to proposal detail page
            navigate(`/proposals/${response.proposal._id}`)
        } catch (error) {
            console.error('Error creating proposal:', error)
            toast.error(error.response?.data?.message || 'Error creating proposal')
        } finally {
            setIsCreating(false)
        }
    }

    const handleBackToForm = () => {
        setStep(1)
    }

    const handleClose = () => {
        if (!isCreating) {
            setStep(1)
            setPreviewData(null)
            onClose()
        }
    }

    return (
        <Modal
            isOpen={isOpen}
            onClose={handleClose}
            title={step === 1 ? "New Proposal" : "Preview Data"}
            subtitle={step === 1
                ? "Upload your Excel sheet to generate and review the proposal data."
                : `${previewData?.fileName} • ${previewData?.rows?.length || 0} rows`
            }
            maxWidth="max-w-5xl"
        >
            {step === 1 ? (
                // Step 1: Form
                <form onSubmit={handleNextStep}>
                    {/* File Upload Area */}
                    <div
                        className={`border-2 border-dashed rounded-lg p-12 mb-8 text-center transition-colors cursor-pointer ${dragActive
                            ? 'border-blue-500 bg-blue-50'
                            : 'border-gray-300 bg-gray-50 hover:border-blue-400 hover:bg-blue-50'
                            }`}
                        onDragEnter={handleDrag}
                        onDragLeave={handleDrag}
                        onDragOver={handleDrag}
                        onDrop={handleDrop}
                    >
                        <label htmlFor="file-upload" className="flex flex-col items-center cursor-pointer">
                            <div className="w-16 h-16 bg-blue-100 rounded-full flex items-center justify-center mb-4">
                                <FiUploadCloud className="text-blue-600" size={28} />
                            </div>
                            <p className="text-sm text-gray-600 mb-1">
                                <span className="text-blue-600 hover:text-blue-700 font-medium">
                                    Click to Upload
                                </span>
                                {' '}or drag and drop
                            </p>
                            <p className="text-xs text-gray-500">.xlsx (Max. size: 20 MB)</p>
                            {file && (
                                <p className="mt-3 text-sm text-gray-700 font-medium">
                                    Selected: {file.name}
                                </p>
                            )}
                        </label>
                        <input
                            id="file-upload"
                            type="file"
                            className="hidden"
                            accept=".xlsx,.xls"
                            onChange={handleFileChange}
                        />
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

                        {/* Template */}
                        <div>
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
                                <option value="capstone">Capstone</option>
                                <option value="dbi" disabled>DBI</option>
                                <option value="pd_steel" disabled>PD Steel</option>
                                <option value="sperrin_tony">Sperrin Tony</option>
                                <option value="tristate_martin" disabled>Tristate Martin</option>
                            </select>
                        </div>
                    </div>

                    {/* Action Buttons */}
                    <div className="flex justify-end gap-3">
                        <button
                            type="button"
                            onClick={handleClose}
                            className="px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors"
                        >
                            Cancel
                        </button>
                        <button
                            type="submit"
                            className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-white bg-[#1A72B9] hover:bg-[#1565C0] rounded-lg transition-colors"
                        >
                            <span>Next: Preview Data</span>
                            <FiArrowRight />
                        </button>
                    </div>
                </form>
            ) : (
                // Step 2: Preview
                <div>
                    {/* Preview Table */}
                    <div className="max-h-[500px] overflow-auto mb-6 border rounded-lg">
                        <table className="w-full border-collapse">
                            <thead className="bg-gray-100 sticky top-0 z-10">
                                <tr>
                                    <th className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider border-b-2 border-gray-300">
                                        #
                                    </th>
                                    {previewData?.headers?.map((header, index) => (
                                        <th
                                            key={index}
                                            className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider border-b-2 border-gray-300 whitespace-nowrap"
                                        >
                                            {header || `Column ${index + 1}`}
                                        </th>
                                    ))}
                                </tr>
                            </thead>
                            <tbody className="divide-y divide-gray-200 bg-white">
                                {previewData?.rows?.length === 0 ? (
                                    <tr>
                                        <td colSpan={(previewData?.headers?.length || 0) + 1} className="px-4 py-8 text-center text-gray-500">
                                            No data found in the Excel file
                                        </td>
                                    </tr>
                                ) : (
                                    previewData?.rows?.map((row, rowIndex) => (
                                        <tr key={rowIndex} className="hover:bg-gray-50 transition-colors">
                                            <td className="px-4 py-2.5 text-sm text-gray-500 font-medium border-r border-gray-200">
                                                {rowIndex + 1}
                                            </td>
                                            {previewData?.headers?.map((_, colIndex) => {
                                                const cellValue = row[colIndex]
                                                const isNumber = !isNaN(cellValue) && cellValue !== '' && cellValue !== null

                                                return (
                                                    <td
                                                        key={colIndex}
                                                        className={`px-4 py-2.5 text-sm ${isNumber ? 'text-right font-mono text-gray-900' : 'text-left text-gray-700'
                                                            }`}
                                                    >
                                                        {cellValue !== null && cellValue !== undefined && cellValue !== ''
                                                            ? String(cellValue)
                                                            : <span className="text-gray-300">-</span>
                                                        }
                                                    </td>
                                                )
                                            })}
                                        </tr>
                                    ))
                                )}
                            </tbody>
                        </table>
                    </div>



                    {/* Action Buttons */}
                    <div className="flex justify-between gap-3">
                        <button
                            type="button"
                            onClick={handleBackToForm}
                            disabled={isCreating}
                            className="flex items-center gap-2 px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors disabled:opacity-50"
                        >
                            <FiArrowLeft />
                            <span>Back</span>
                        </button>
                        <button
                            type="button"
                            onClick={handleCreateProposal}
                            disabled={isCreating}
                            className="px-6 py-2 text-sm font-medium text-white bg-[#1A72B9] hover:bg-[#1565C0] rounded-lg transition-colors disabled:opacity-50"
                        >
                            {isCreating ? 'Creating...' : 'Create Proposal'}
                        </button>
                    </div>
                </div>
            )}
        </Modal>
    )
}

export default CreateProposalModal

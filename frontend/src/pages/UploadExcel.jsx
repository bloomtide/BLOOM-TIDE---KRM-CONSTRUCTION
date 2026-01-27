import React, { useState } from 'react'
import * as XLSX from 'xlsx'
import { useNavigate } from 'react-router-dom'
import { FiUpload, FiFile, FiX } from 'react-icons/fi'

const UploadExcel = () => {
  const [file, setFile] = useState(null)
  const [isDragging, setIsDragging] = useState(false)
  const [isProcessing, setIsProcessing] = useState(false)
  const navigate = useNavigate()

  const handleFileChange = (e) => {
    const selectedFile = e.target.files[0]
    validateAndSetFile(selectedFile)
  }

  const handleDragOver = (e) => {
    e.preventDefault()
    setIsDragging(true)
  }

  const handleDragLeave = (e) => {
    e.preventDefault()
    setIsDragging(false)
  }

  const handleDrop = (e) => {
    e.preventDefault()
    setIsDragging(false)
    const droppedFile = e.dataTransfer.files[0]
    validateAndSetFile(droppedFile)
  }

  const validateAndSetFile = (file) => {
    if (!file) return

    const validExtensions = ['.xlsx', '.xls']
    const fileExtension = file.name.substring(file.name.lastIndexOf('.')).toLowerCase()

    if (validExtensions.includes(fileExtension)) {
      setFile(file)
    } else {
      alert('Please upload a valid Excel file (.xlsx or .xls)')
    }
  }

  const handleRemoveFile = () => {
    setFile(null)
  }

  const handleUpload = () => {
    if (!file) return

    setIsProcessing(true)
    const reader = new FileReader()

    reader.onload = (e) => {
      try {
        const data = new Uint8Array(e.target.result)
        const workbook = XLSX.read(data, { type: 'array' })
        const firstSheet = workbook.Sheets[workbook.SheetNames[0]]
        const jsonData = XLSX.utils.sheet_to_json(firstSheet, { header: 1, defval: '' })
        
        // Navigate to template selection with data
        navigate('/template-select', { 
          state: { 
            rawData: jsonData, 
            fileName: file.name,
            sheetName: workbook.SheetNames[0]
          } 
        })
      } catch (error) {
        console.error('Error parsing Excel file:', error)
        alert('Error parsing Excel file. Please make sure it is a valid Excel file.')
        setIsProcessing(false)
      }
    }

    reader.onerror = () => {
      alert('Error reading file. Please try again.')
      setIsProcessing(false)
    }

    reader.readAsArrayBuffer(file)
  }

  const formatFileSize = (bytes) => {
    if (bytes === 0) return '0 Bytes'
    const k = 1024
    const sizes = ['Bytes', 'KB', 'MB']
    const i = Math.floor(Math.log(bytes) / Math.log(k))
    return Math.round(bytes / Math.pow(k, i) * 100) / 100 + ' ' + sizes[i]
  }

  return (
    <div className="min-h-screen bg-gray-50 p-8">
      <div className="max-w-4xl mx-auto">
        <div className="mb-8">
          <h1 className="text-3xl font-bold text-gray-900 mb-2">Upload Raw Excel Data</h1>
          <p className="text-gray-600">Upload your construction project data in Excel format (.xlsx or .xls)</p>
        </div>

        <div className="bg-white rounded-lg shadow-md p-8">
          {!file ? (
            <div
              onDragOver={handleDragOver}
              onDragLeave={handleDragLeave}
              onDrop={handleDrop}
              className={`border-2 border-dashed rounded-lg p-12 text-center transition-colors ${
                isDragging 
                  ? 'border-blue-500 bg-blue-50' 
                  : 'border-gray-300 hover:border-gray-400'
              }`}
            >
              <FiUpload className="mx-auto text-5xl text-gray-400 mb-4" />
              <h3 className="text-lg font-semibold text-gray-700 mb-2">
                Drag and drop your Excel file here
              </h3>
              <p className="text-gray-500 mb-4">or</p>
              <label className="inline-block px-6 py-3 bg-blue-600 text-white rounded-lg hover:bg-blue-700 cursor-pointer transition-colors">
                Browse Files
                <input
                  type="file"
                  accept=".xlsx,.xls"
                  onChange={handleFileChange}
                  className="hidden"
                />
              </label>
              <p className="text-sm text-gray-400 mt-4">Supported formats: .xlsx, .xls</p>
            </div>
          ) : (
            <div className="space-y-6">
              <div className="flex items-center justify-between p-4 bg-gray-50 rounded-lg">
                <div className="flex items-center space-x-4">
                  <FiFile className="text-3xl text-blue-600" />
                  <div>
                    <p className="font-semibold text-gray-900">{file.name}</p>
                    <p className="text-sm text-gray-500">{formatFileSize(file.size)}</p>
                  </div>
                </div>
                <button
                  onClick={handleRemoveFile}
                  className="text-red-600 hover:text-red-700 transition-colors"
                  disabled={isProcessing}
                >
                  <FiX className="text-2xl" />
                </button>
              </div>

              <div className="flex space-x-4">
                <button
                  onClick={handleUpload}
                  disabled={isProcessing}
                  className={`flex-1 px-6 py-3 rounded-lg font-semibold transition-colors ${
                    isProcessing
                      ? 'bg-gray-400 text-gray-200 cursor-not-allowed'
                      : 'bg-blue-600 text-white hover:bg-blue-700'
                  }`}
                >
                  {isProcessing ? 'Processing...' : 'Select Template'}
                </button>
                <button
                  onClick={() => navigate('/dashboard')}
                  disabled={isProcessing}
                  className="px-6 py-3 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300 transition-colors disabled:opacity-50"
                >
                  Cancel
                </button>
              </div>
            </div>
          )}
        </div>

        <div className="mt-8 bg-blue-50 border border-blue-200 rounded-lg p-4">
          <h4 className="font-semibold text-blue-900 mb-2">Expected Data Format:</h4>
          <p className="text-sm text-blue-800">
            Your Excel file should contain columns: <strong>Page</strong>, <strong>Digitizer Item</strong>, <strong>Total</strong>, and <strong>Units</strong>
          </p>
        </div>
      </div>
    </div>
  )
}

export default UploadExcel

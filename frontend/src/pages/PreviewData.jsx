import React, { useState, useEffect } from 'react'
import { useLocation, useNavigate } from 'react-router-dom'
import { FiArrowLeft, FiArrowRight, FiFileText } from 'react-icons/fi'

const PreviewData = () => {
  const location = useLocation()
  const navigate = useNavigate()
  const { rawData, fileName, sheetName, selectedTemplate } = location.state || {}
  const [headers, setHeaders] = useState([])
  const [rows, setRows] = useState([])

  useEffect(() => {
    if (!rawData || rawData.length === 0) {
      navigate('/upload')
      return
    }

    // First row is headers
    setHeaders(rawData[0])
    // Rest are data rows
    setRows(rawData.slice(1))
  }, [rawData, navigate])

  const handleGoBack = () => {
    navigate('/template-select', { 
      state: { rawData, fileName, sheetName }
    })
  }

  const handleProceed = () => {
    // Pass raw data to spreadsheet for processing
    navigate('/spreadsheet', { 
      state: { 
        rawData, 
        fileName,
        sheetName,
        headers,
        rows,
        selectedTemplate
      } 
    })
  }

  if (!rawData) {
    return null
  }

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">
      {/* Header */}
      <div className="bg-white border-b shadow-sm">
        <div className="max-w-7xl mx-auto px-6 py-4">
          <div className="flex items-center justify-between">
            <div className="flex items-center space-x-3">
              <FiFileText className="text-2xl text-blue-600" />
              <div>
                <h1 className="text-2xl font-bold text-gray-900">Preview Raw Data</h1>
                <p className="text-sm text-gray-600">
                  {fileName} • {sheetName} • {rows.length} rows • Template: {selectedTemplate || 'capstone'}
                </p>
              </div>
            </div>
            <div className="flex space-x-3">
              <button
                onClick={handleGoBack}
                className="flex items-center space-x-2 px-4 py-2 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300 transition-colors"
              >
                <FiArrowLeft />
                <span>Go Back</span>
              </button>
              <button
                onClick={handleProceed}
                className="flex items-center space-x-2 px-6 py-2 bg-blue-600 text-white rounded-lg hover:bg-blue-700 transition-colors"
              >
                <span>Proceed to Calculations</span>
                <FiArrowRight />
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* Data Table */}
      <div className="flex-1 overflow-auto p-6">
        <div className="max-w-7xl mx-auto">
          <div className="bg-white rounded-lg shadow-md overflow-hidden">
            <div className="overflow-x-auto">
              <table className="w-full">
                <thead className="bg-gray-100 border-b border-gray-200">
                  <tr>
                    <th className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider sticky top-0 bg-gray-100">
                      #
                    </th>
                    {headers.map((header, index) => (
                      <th
                        key={index}
                        className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider sticky top-0 bg-gray-100"
                      >
                        {header || `Column ${index + 1}`}
                      </th>
                    ))}
                  </tr>
                </thead>
                <tbody className="divide-y divide-gray-200">
                  {rows.length === 0 ? (
                    <tr>
                      <td colSpan={headers.length + 1} className="px-4 py-8 text-center text-gray-500">
                        No data found in the Excel file
                      </td>
                    </tr>
                  ) : (
                    rows.map((row, rowIndex) => (
                      <tr key={rowIndex} className="hover:bg-gray-50 transition-colors">
                        <td className="px-4 py-3 text-sm text-gray-500 font-medium">
                          {rowIndex + 1}
                        </td>
                        {headers.map((_, colIndex) => {
                          const cellValue = row[colIndex]
                          const isNumber = !isNaN(cellValue) && cellValue !== '' && cellValue !== null
                          
                          return (
                            <td
                              key={colIndex}
                              className={`px-4 py-3 text-sm ${
                                isNumber ? 'text-right font-mono text-gray-900' : 'text-left text-gray-700'
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
          </div>

          {/* Summary Info */}
          <div className="mt-4 bg-blue-50 border border-blue-200 rounded-lg p-4">
            <div className="flex items-start space-x-2">
              <div className="text-blue-600 mt-0.5">ℹ️</div>
              <div>
                <h4 className="font-semibold text-blue-900 mb-1">Next Steps:</h4>
                <p className="text-sm text-blue-800">
                  Review the data above. When ready, click <strong>"Proceed to Calculations"</strong> to generate 
                  the Calculations Sheet and Proposal Sheet based on this raw data.
                </p>
              </div>
            </div>
          </div>
        </div>
      </div>
    </div>
  )
}

export default PreviewData

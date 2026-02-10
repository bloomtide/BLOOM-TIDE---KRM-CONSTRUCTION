import React from 'react'
import Modal from './Modal'
import { FiX, FiDownload } from 'react-icons/fi'
import * as XLSX from 'xlsx'

const RawDataPreviewModal = ({ isOpen, onClose, rawExcelData }) => {
    if (!rawExcelData) return null

    const { fileName, sheetName, headers, rows } = rawExcelData

    const handleDownload = () => {
        // Create workbook and worksheet
        const wb = XLSX.utils.book_new()

        // Combine headers and rows
        const wsData = [
            headers || [],
            ...(rows || [])
        ]

        const ws = XLSX.utils.aoa_to_sheet(wsData)

        // Add worksheet to workbook
        XLSX.utils.book_append_sheet(wb, ws, sheetName || 'Raw Data')

        // Generate Excel file
        XLSX.writeFile(wb, fileName || 'raw_data.xlsx')
    }

    return (
        <Modal
            isOpen={isOpen}
            onClose={onClose}
            title="Raw Excel Data Preview"
            subtitle={`${fileName}${sheetName ? ` • ${sheetName}` : ''} • ${rows?.length || 0} rows`}
            headerActions={
                <button
                    onClick={handleDownload}
                    className="flex items-center gap-2 px-3 py-1.5 text-sm font-medium text-white bg-green-600 rounded-md hover:bg-green-700 transition-colors"
                    title="Download as Excel"
                >
                    <FiDownload size={16} />
                    <span className="hidden sm:inline">Download</span>
                </button>
            }
            maxWidth="max-w-6xl"
        >
            <div className="max-h-[600px] overflow-auto">
                <table className="w-full border-collapse">
                    <thead className="bg-gray-100 sticky top-0 z-10">
                        <tr>
                            <th className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider border-b-2 border-gray-300">
                                #
                            </th>
                            {headers?.map((header, index) => (
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
                        {rows?.length === 0 ? (
                            <tr>
                                <td colSpan={(headers?.length || 0) + 1} className="px-4 py-8 text-center text-gray-500">
                                    No data found in the Excel file
                                </td>
                            </tr>
                        ) : (
                            rows?.map((row, rowIndex) => (
                                <tr key={rowIndex} className="hover:bg-gray-50 transition-colors">
                                    <td className="px-4 py-3 text-sm text-gray-500 font-medium border-r border-gray-200">
                                        {rowIndex + 1}
                                    </td>
                                    {headers?.map((_, colIndex) => {
                                        const cellValue = row[colIndex]
                                        const isNumber = !isNaN(cellValue) && cellValue !== '' && cellValue !== null

                                        return (
                                            <td
                                                key={colIndex}
                                                className={`px-4 py-3 text-sm ${isNumber ? 'text-right font-mono text-gray-900' : 'text-left text-gray-700'
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

            {/* Footer */}
            <div className="mt-6 flex justify-end">
                <button
                    onClick={onClose}
                    className="px-4 py-2 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-lg hover:bg-gray-50 transition-colors"
                >
                    Close
                </button>
            </div>
        </Modal>
    )
}

export default RawDataPreviewModal

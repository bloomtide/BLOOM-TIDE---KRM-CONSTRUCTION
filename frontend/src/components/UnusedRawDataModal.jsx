import React, { useState } from 'react'
import Modal from './Modal'
import { FiCheckSquare, FiSquare } from 'react-icons/fi'

const UnusedRawDataModal = ({ isOpen, onClose, unusedRows = [], onUpdateRowStatus, onBulkUpdateRowStatus, headers = [] }) => {
    const [updatingParams, setUpdatingParams] = useState(null)
    const [stagedUpdates, setStagedUpdates] = useState({}) // { rowIndex: isUsed }

    if (!isOpen) return null

    // Determine the effective status of a row (staged > original)
    const getRowStatus = (row) => {
        if (stagedUpdates.hasOwnProperty(row.rowIndex)) {
            return stagedUpdates[row.rowIndex]
        }
        return row.isUsed
    }

    const handleCheckboxClick = (rowIndex, currentEffectiveStatus) => {
        // Toggle the status in staged updates
        setStagedUpdates(prev => ({
            ...prev,
            [rowIndex]: !currentEffectiveStatus
        }))
    }

    const handleSave = async () => {
        setUpdatingParams('bulk')
        try {
            const updates = Object.entries(stagedUpdates).map(([rowIndex, isUsed]) => ({
                rowIndex: parseInt(rowIndex, 10),
                isUsed
            }))
            await onBulkUpdateRowStatus(updates)
            setStagedUpdates({}) // Clear staged updates on success
            onClose() // Close modal after save
        } catch (error) {
            console.error("Failed to save bulk updates", error)
        } finally {
            setUpdatingParams(null)
        }
    }

    const handleReset = () => {
        setStagedUpdates({})
    }

    // Filter to show only rows that are NOT marked as used, unless we want to show all and let them toggle back?
    // Requirement said: "display unused raw data rows... with persistent checkboxes for clients to mark rows as manually processed."
    // So we should show the rows that the system identified as unused. The checkbox allows them to say "I've handled this".
    // So we list all rows in `unusedRawDataRows`.

    // Sort by rowIndex
    const sortedRows = [...unusedRows].sort((a, b) => a.rowIndex - b.rowIndex)
    const hasChanges = Object.keys(stagedUpdates).length > 0

    return (
        <Modal
            isOpen={isOpen}
            onClose={onClose}
            title="Unused Raw Data Rows"
            subtitle={`${sortedRows.length} rows found that were not automatically processed`}
            maxWidth="max-w-7xl"
        >
            <div className="max-h-[600px] overflow-auto">
                <table className="w-full border-collapse">
                    <thead className="bg-gray-100 sticky top-0 z-10">
                        <tr>
                            <th className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider border-b-2 border-gray-300 w-16">
                                Status
                            </th>
                            <th className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider border-b-2 border-gray-300 w-16">
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
                        {sortedRows.length === 0 ? (
                            <tr>
                                <td colSpan={(headers?.length || 0) + 2} className="px-4 py-8 text-center text-gray-500">
                                    No unused rows found. All data has been processed!
                                </td>
                            </tr>
                        ) : (
                            sortedRows.map((row) => {
                                const isUsed = getRowStatus(row)
                                return (
                                    <tr key={row.rowIndex} className={`hover:bg-gray-50 transition-colors ${isUsed ? 'bg-gray-50 opacity-60' : ''}`}>
                                        <td className="px-4 py-3 text-center border-r border-gray-200">
                                            <button
                                                onClick={() => handleCheckboxClick(row.rowIndex, isUsed)}
                                                disabled={updatingParams === 'bulk'}
                                                className={`focus:outline-none transition-colors ${updatingParams === 'bulk' ? 'opacity-50 cursor-wait' : 'cursor-pointer'}`}
                                                title={isUsed ? "Mark as unprocessed" : "Mark as processed"}
                                            >
                                                {isUsed ? (
                                                    <FiCheckSquare className="w-5 h-5 text-green-600" />
                                                ) : (
                                                    <FiSquare className="w-5 h-5 text-gray-400 hover:text-gray-600" />
                                                )}
                                            </button>
                                        </td>
                                        <td className="px-4 py-3 text-sm text-gray-500 font-medium border-r border-gray-200">
                                            {row.rowIndex + 1}
                                        </td>
                                        {headers?.map((_, colIndex) => {
                                            const cellValue = row.rowData && row.rowData[colIndex]
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
                                )
                            })
                        )}
                    </tbody>
                </table>
            </div>

            {/* Footer */}
            <div className="mt-6 flex justify-between items-center bg-gray-50 p-4 rounded-lg border border-gray-200">
                <div className="text-sm text-gray-600">
                    <span className="font-medium text-gray-900">{sortedRows.filter(r => !getRowStatus(r)).length}</span> rows remaining to review
                </div>
                <div className="flex gap-2">
                    {hasChanges && (
                        <>
                            <button
                                onClick={handleReset}
                                disabled={updatingParams === 'bulk'}
                                className="px-4 py-2 text-sm font-medium text-red-600 bg-white border border-red-300 rounded-lg hover:bg-red-50 transition-colors shadow-sm disabled:opacity-50"
                            >
                                Reset Changes
                            </button>
                            <button
                                onClick={handleSave}
                                disabled={updatingParams === 'bulk'}
                                className="px-4 py-2 text-sm font-medium text-white bg-blue-600 border border-transparent rounded-lg hover:bg-blue-700 transition-colors shadow-sm disabled:opacity-50 flex items-center gap-2"
                            >
                                {updatingParams === 'bulk' && <div className="w-4 h-4 border-2 border-white border-t-transparent rounded-full animate-spin"></div>}
                                Save Changes
                            </button>
                        </>
                    )}
                </div>
            </div>
        </Modal>
    )
}

export default UnusedRawDataModal

import React, { useState, useEffect } from 'react'
import Modal from './Modal'
import { FiDownload, FiEdit2, FiSave } from 'react-icons/fi'
import * as XLSX from 'xlsx'

const RawDataPreviewModal = ({ isOpen, onClose, rawExcelData, onSave }) => {
    const [isEditing, setIsEditing] = useState(false)
    const [editedHeaders, setEditedHeaders] = useState([])
    const [editedRows, setEditedRows] = useState([])
    const [isSaving, setIsSaving] = useState(false)

    // Sync editable state when modal opens or rawExcelData changes
    useEffect(() => {
        if (!isOpen || !rawExcelData) return
        const { headers = [], rows = [] } = rawExcelData
        setEditedHeaders([...headers])
        setEditedRows(rows.map(row => Array.isArray(row) ? [...row] : []))
        setIsEditing(false)
    }, [isOpen, rawExcelData])

    if (!rawExcelData) return null

    const { fileName, sheetName, headers, rows } = rawExcelData
    const displayHeaders = isEditing ? editedHeaders : (headers || [])
    const displayRows = isEditing ? editedRows : (rows || [])

    const handleDownload = () => {
        const wb = XLSX.utils.book_new()
        const wsData = [displayHeaders, ...displayRows]
        const ws = XLSX.utils.aoa_to_sheet(wsData)
        XLSX.utils.book_append_sheet(wb, ws, sheetName || 'Raw Data')
        XLSX.writeFile(wb, fileName || 'raw_data.xlsx')
    }

    const updateHeader = (colIndex, value) => {
        setEditedHeaders(prev => {
            const next = [...prev]
            while (next.length <= colIndex) next.push('')
            next[colIndex] = value
            return next
        })
    }

    const updateCell = (rowIndex, colIndex, value) => {
        setEditedRows(prev => {
            const next = prev.map((row, i) => (i === rowIndex ? [...row] : row))
            const row = next[rowIndex] || []
            while (row.length <= colIndex) row.push('')
            row[colIndex] = value
            next[rowIndex] = row
            return next
        })
    }

    const handleSave = async () => {
        if (!onSave) return
        const colCount = editedHeaders.length
        const valid = editedRows.every(row => !row || row.length <= colCount)
        if (!valid) {
            // optional: toast or message
            return
        }
        setIsSaving(true)
        try {
            await onSave({
                fileName: fileName || 'raw_data.xlsx',
                sheetName: sheetName || 'Raw Data',
                headers: editedHeaders,
                rows: editedRows,
            })
            setIsEditing(false)
            onClose()
        } catch (e) {
            console.error('Save raw data failed', e)
        } finally {
            setIsSaving(false)
        }
    }

    const headerActions = (
        <div className="flex items-center gap-2">
            <button
                onClick={handleDownload}
                className="flex items-center gap-2 px-3 py-1.5 text-sm font-medium text-white bg-green-600 rounded-md hover:bg-green-700 transition-colors"
                title="Download as Excel"
            >
                <FiDownload size={16} />
                <span className="hidden sm:inline">Download</span>
            </button>
            {onSave && (
                isEditing ? (
                    <>
                        <button
                            onClick={() => setIsEditing(false)}
                            className="flex items-center gap-2 px-3 py-1.5 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-md hover:bg-gray-50 transition-colors"
                        >
                            Cancel
                        </button>
                        <button
                            onClick={handleSave}
                            disabled={isSaving}
                            className="flex items-center gap-2 px-3 py-1.5 text-sm font-medium text-white bg-blue-600 rounded-md hover:bg-blue-700 transition-colors disabled:opacity-50"
                        >
                            <FiSave size={16} />
                            <span className="hidden sm:inline">{isSaving ? 'Saving…' : 'Save'}</span>
                        </button>
                    </>
                ) : (
                    <button
                        onClick={() => setIsEditing(true)}
                        className="flex items-center gap-2 px-3 py-1.5 text-sm font-medium text-white bg-indigo-600 rounded-md hover:bg-indigo-700 transition-colors"
                        title="Edit raw data"
                    >
                        <FiEdit2 size={16} />
                        <span className="hidden sm:inline">Edit</span>
                    </button>
                )
            )}
        </div>
    )

    return (
        <Modal
            isOpen={isOpen}
            onClose={onClose}
            title="Raw Excel Data Preview"
            subtitle={`${fileName}${sheetName ? ` • ${sheetName}` : ''} • ${displayRows?.length || 0} rows${isEditing ? ' • Editing' : ''}`}
            headerActions={headerActions}
            maxWidth="max-w-6xl"
        >
            <div className="max-h-[600px] overflow-auto">
                <table className="w-full border-collapse">
                    <thead className="bg-gray-100 sticky top-0 z-10">
                        <tr>
                            <th className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider border-b-2 border-gray-300 w-12">
                                #
                            </th>
                            {displayHeaders?.map((header, index) => (
                                <th
                                    key={index}
                                    className="px-4 py-3 text-left text-xs font-semibold text-gray-600 uppercase tracking-wider border-b-2 border-gray-300 whitespace-nowrap"
                                >
                                    {isEditing ? (
                                        <input
                                            value={header ?? ''}
                                            onChange={e => updateHeader(index, e.target.value)}
                                            className="w-full min-w-[80px] px-2 py-1 text-sm border border-gray-300 rounded bg-white"
                                        />
                                    ) : (
                                        (header ?? `Column ${index + 1}`)
                                    )}
                                </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-200 bg-white">
                        {displayRows?.length === 0 ? (
                            <tr>
                                <td colSpan={(displayHeaders?.length || 0) + 1} className="px-4 py-8 text-center text-gray-500">
                                    No data found in the Excel file
                                </td>
                            </tr>
                        ) : (
                            displayRows?.map((row, rowIndex) => (
                                <tr key={rowIndex} className="hover:bg-gray-50 transition-colors">
                                    <td className="px-4 py-3 text-sm text-gray-500 font-medium border-r border-gray-200">
                                        {rowIndex + 1}
                                    </td>
                                    {displayHeaders?.map((_, colIndex) => {
                                        const cellValue = row[colIndex]
                                        if (isEditing) {
                                            return (
                                                <td key={colIndex} className="p-1">
                                                    <input
                                                        value={cellValue !== null && cellValue !== undefined ? String(cellValue) : ''}
                                                        onChange={e => updateCell(rowIndex, colIndex, e.target.value)}
                                                        className="w-full min-w-[60px] px-2 py-1 text-sm border border-gray-300 rounded bg-white"
                                                    />
                                                </td>
                                            )
                                        }
                                        const isNumber = !isNaN(cellValue) && cellValue !== '' && cellValue !== null
                                        return (
                                            <td
                                                key={colIndex}
                                                className={`px-4 py-3 text-sm ${isNumber ? 'text-right font-mono text-gray-900' : 'text-left text-gray-700'}`}
                                            >
                                                {cellValue !== null && cellValue !== undefined && cellValue !== ''
                                                    ? String(cellValue)
                                                    : <span className="text-gray-300">-</span>}
                                            </td>
                                        )
                                    })}
                                </tr>
                            ))
                        )}
                    </tbody>
                </table>
            </div>

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

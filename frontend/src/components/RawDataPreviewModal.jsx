import React, { useState, useEffect } from 'react'
import toast from 'react-hot-toast'
import Modal from './Modal'
import { FiDownload, FiSave, FiRefreshCw } from 'react-icons/fi'
import * as XLSX from 'xlsx'
import { proposalAPI } from '../services/proposalService'

const RawDataPreviewModal = ({ isOpen, onClose, rawExcelData, proposalId, onSaveSuccess, onRebuildWithRawData }) => {
    const [editableHeaders, setEditableHeaders] = useState([])
    const [editableRows, setEditableRows] = useState([])
    const [saving, setSaving] = useState(false)
    const [saveError, setSaveError] = useState(null)

    useEffect(() => {
        if (isOpen && rawExcelData) {
            setEditableHeaders(Array.isArray(rawExcelData.headers) ? [...rawExcelData.headers] : [])
            setEditableRows(Array.isArray(rawExcelData.rows)
                ? rawExcelData.rows.map(row => Array.isArray(row) ? [...row] : [])
                : [])
            setSaveError(null)
        }
    }, [isOpen, rawExcelData])

    if (!rawExcelData) return null

    const { fileName, sheetName } = rawExcelData
    const canSave = !!proposalId && !!onSaveSuccess
    const canRebuild = !!onRebuildWithRawData

    const handleDownload = () => {
        const wb = XLSX.utils.book_new()
        const wsData = [editableHeaders, ...editableRows]
        const ws = XLSX.utils.aoa_to_sheet(wsData)
        XLSX.utils.book_append_sheet(wb, ws, sheetName || 'Raw Data')
        XLSX.writeFile(wb, fileName || 'raw_data.xlsx')
    }

    const setHeader = (colIndex, value) => {
        setEditableHeaders(prev => {
            const next = [...prev]
            while (next.length <= colIndex) next.push('')
            next[colIndex] = value
            return next
        })
    }

    const setCell = (rowIndex, colIndex, value) => {
        setEditableRows(prev => {
            const next = prev.map((row, r) => {
                if (r !== rowIndex) return row
                const newRow = [...(row || [])]
                while (newRow.length <= colIndex) newRow.push('')
                newRow[colIndex] = value
                return newRow
            })
            return next
        })
    }

    const normalizeValue = (v) => {
        if (v === null || v === undefined || v === '') return ''
        const s = String(v).trim()
        const num = Number(s)
        if (s !== '' && !Number.isNaN(num)) return num
        return s
    }

    const getEditedRawExcelData = () => ({
        fileName: fileName || 'raw_data.xlsx',
        sheetName: sheetName || 'Sheet1',
        headers: editableHeaders.map(h => (h != null && h !== '') ? String(h) : ''),
        rows: editableRows.map(row => (row || []).map(cell => normalizeValue(cell))),
    })

    const handleRebuild = () => {
        if (!onRebuildWithRawData) return
        onRebuildWithRawData(getEditedRawExcelData())
        onClose()
    }

    const handleSave = async () => {
        if (!proposalId || !onSaveSuccess) return
        setSaving(true)
        setSaveError(null)
        try {
            const normalizedRows = editableRows.map(row =>
                (row || []).map(cell => normalizeValue(cell))
            )
            const normalizedHeaders = editableHeaders.map(h => (h != null && h !== '') ? String(h) : '')
            await proposalAPI.update(proposalId, {
                rawExcelData: {
                    fileName: fileName || 'raw_data.xlsx',
                    sheetName: sheetName || 'Sheet1',
                    headers: normalizedHeaders,
                    rows: normalizedRows,
                },
            })
            toast.success('Raw data saved to proposal')
            onSaveSuccess()
            onClose()
        } catch (err) {
            setSaveError(err.response?.data?.message || err.message || 'Failed to save')
        } finally {
            setSaving(false)
        }
    }

    const numCols = Math.max(editableHeaders.length, ...editableRows.map(r => (r || []).length), 1)

    return (
        <Modal
            isOpen={isOpen}
            onClose={onClose}
            title="Raw Excel Data Preview"
            subtitle={`${fileName}${sheetName ? ` • ${sheetName}` : ''} • ${editableRows?.length || 0} rows • Editable`}
            headerActions={
                <div className="flex items-center gap-2">
                    {canRebuild && (
                        <button
                            type="button"
                            onClick={handleRebuild}
                            className="flex items-center gap-2 px-3 py-1.5 text-sm font-medium text-white bg-amber-600 rounded-md hover:bg-amber-700 transition-colors"
                            title="Rebuild calculation sheet from this edited raw data (not from DB)"
                        >
                            <FiRefreshCw size={16} />
                            <span className="hidden sm:inline">Rebuild calculation sheet</span>
                        </button>
                    )}
                    <button
                        type="button"
                        onClick={handleDownload}
                        className="flex items-center gap-2 px-3 py-1.5 text-sm font-medium text-gray-700 bg-white border border-gray-300 rounded-md hover:bg-gray-50 transition-colors"
                        title="Download as Excel"
                    >
                        <FiDownload size={16} />
                        <span className="hidden sm:inline">Download</span>
                    </button>
                    {canSave && (
                        <button
                            type="button"
                            onClick={handleSave}
                            disabled={saving}
                            className="flex items-center gap-2 px-3 py-1.5 text-sm font-medium text-white bg-blue-600 rounded-md hover:bg-blue-700 disabled:opacity-50 transition-colors"
                            title="Save changes to proposal"
                        >
                            <FiSave size={16} />
                            <span className="hidden sm:inline">{saving ? 'Saving…' : 'Save'}</span>
                        </button>
                    )}
                </div>
            }
            maxWidth="max-w-6xl"
        >
            <div className="max-h-[600px] overflow-auto">
                <table className="w-full border-collapse">
                    <thead className="bg-gray-100 sticky top-0 z-10">
                        <tr>
                            <th className="px-2 py-2 text-left text-xs font-semibold text-gray-600 uppercase border-b-2 border-gray-300 w-12">
                                #
                            </th>
                            {Array.from({ length: numCols }, (_, colIndex) => (
                                <th key={colIndex} className="border-b-2 border-gray-300 p-0">
                                    <input
                                        type="text"
                                        value={editableHeaders[colIndex] ?? ''}
                                        onChange={(e) => setHeader(colIndex, e.target.value)}
                                        className="w-full min-w-[80px] px-2 py-2 text-xs font-semibold text-gray-700 bg-gray-50 border-0 border-r border-gray-200 focus:ring-1 focus:ring-blue-500 focus:bg-white"
                                        placeholder={`Col ${colIndex + 1}`}
                                    />
                                </th>
                            ))}
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-200 bg-white">
                        {editableRows.length === 0 ? (
                            <tr>
                                <td colSpan={numCols + 1} className="px-4 py-8 text-center text-gray-500">
                                    No data. Add rows by editing and saving.
                                </td>
                            </tr>
                        ) : (
                            editableRows.map((row, rowIndex) => (
                                <tr key={rowIndex} className="hover:bg-gray-50">
                                    <td className="px-2 py-1 text-sm text-gray-500 font-medium border-r border-gray-200 sticky left-0 bg-white">
                                        {rowIndex + 1}
                                    </td>
                                    {Array.from({ length: numCols }, (_, colIndex) => {
                                        const cellValue = row[colIndex]
                                        const display = cellValue !== null && cellValue !== undefined && cellValue !== '' ? String(cellValue) : ''
                                        return (
                                            <td key={colIndex} className="p-0 border-r border-gray-100">
                                                <input
                                                    type="text"
                                                    value={display}
                                                    onChange={(e) => setCell(rowIndex, colIndex, e.target.value)}
                                                    className="w-full min-w-[80px] px-2 py-1.5 text-sm text-gray-900 border-0 border-b border-gray-100 focus:ring-1 focus:ring-blue-500 focus:z-10"
                                                />
                                            </td>
                                        )
                                    })}
                                </tr>
                            ))
                        )}
                    </tbody>
                </table>
            </div>

            {saveError && (
                <p className="mt-3 text-sm text-red-600">{saveError}</p>
            )}

            <div className="mt-6 flex justify-end gap-2 flex-wrap">
                {canRebuild && (
                    <button
                        type="button"
                        onClick={handleRebuild}
                        className="px-4 py-2 text-sm font-medium text-white bg-amber-600 rounded-lg hover:bg-amber-700 transition-colors flex items-center gap-2"
                        title="Rebuild from the edited raw data above (not from DB)"
                    >
                        <FiRefreshCw size={16} />
                        Rebuild calculation sheet
                    </button>
                )}
                {canSave && (
                    <button
                        type="button"
                        onClick={handleSave}
                        disabled={saving}
                        className="px-4 py-2 text-sm font-medium text-white bg-blue-600 rounded-lg hover:bg-blue-700 disabled:opacity-50 transition-colors flex items-center gap-2"
                    >
                        <FiSave size={16} />
                        {saving ? 'Saving…' : 'Save to proposal'}
                    </button>
                )}
                <button
                    type="button"
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

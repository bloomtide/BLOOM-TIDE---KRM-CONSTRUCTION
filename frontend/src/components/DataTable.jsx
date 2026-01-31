import React, { useState } from 'react'
import { FiEdit2, FiDownload, FiTrash2 } from 'react-icons/fi'

const DataTable = ({
    columns = [],
    data = [],
    onEdit,
    onDownload,
    onDelete,
    onBulkDelete,
    showActions = true,
    showCheckbox = true
}) => {
    const [selectedRows, setSelectedRows] = useState([])
    const [selectAll, setSelectAll] = useState(false)

    const handleSelectAll = () => {
        if (selectAll) {
            setSelectedRows([])
        } else {
            setSelectedRows(data.map((_, index) => index))
        }
        setSelectAll(!selectAll)
    }

    const handleSelectRow = (index) => {
        if (selectedRows.includes(index)) {
            setSelectedRows(selectedRows.filter(i => i !== index))
        } else {
            setSelectedRows([...selectedRows, index])
        }
    }

    return (
        <div className="bg-white rounded-lg border border-gray-200 overflow-hidden">
            <div className="overflow-x-auto">
                {selectedRows.length > 0 && onBulkDelete && (
                    <div className="bg-red-50 px-6 py-2 flex items-center justify-between border-b border-red-100">
                        <span className="text-sm text-red-700 font-medium">
                            {selectedRows.length} selected
                        </span>
                        <button
                            onClick={() => onBulkDelete(selectedRows.map(index => data[index]))}
                            className="text-red-600 hover:text-red-700 text-sm font-medium flex items-center gap-2"
                        >
                            <FiTrash2 size={16} />
                            Delete Selected
                        </button>
                    </div>
                )}
                <table className="w-full">
                    <thead className="bg-gray-50 border-b border-gray-200">
                        <tr>
                            {showCheckbox && (
                                <th className="px-6 py-3 text-left w-12">
                                    <input
                                        type="checkbox"
                                        checked={selectAll}
                                        onChange={handleSelectAll}
                                        className="w-4 h-4 text-blue-600 border-gray-300 rounded focus:ring-2 focus:ring-blue-500"
                                    />
                                </th>
                            )}
                            {columns.map((column, index) => (
                                <th
                                    key={index}
                                    className={`px-6 py-3 text-left text-xs font-bold text-gray-700 uppercase tracking-wider ${column.className || ''}`}
                                >
                                    {column.header}
                                </th>
                            ))}
                            {showActions && (
                                <th className="px-6 py-3 text-left text-xs font-bold text-gray-700 uppercase tracking-wider">
                                    Actions
                                </th>
                            )}
                        </tr>
                    </thead>
                    <tbody className="divide-y divide-gray-200">
                        {data.length === 0 ? (
                            <tr>
                                <td
                                    colSpan={columns.length + (showCheckbox ? 1 : 0) + (showActions ? 1 : 0)}
                                    className="px-6 py-8 text-center text-gray-500"
                                >
                                    No data available
                                </td>
                            </tr>
                        ) : (
                            data.map((row, rowIndex) => (
                                <tr
                                    key={rowIndex}
                                    className="hover:bg-gray-50 transition-colors"
                                >
                                    {showCheckbox && (
                                        <td className="px-6 py-4">
                                            <input
                                                type="checkbox"
                                                checked={selectedRows.includes(rowIndex)}
                                                onChange={() => handleSelectRow(rowIndex)}
                                                className="w-4 h-4 text-blue-600 border-gray-300 rounded focus:ring-2 focus:ring-blue-500"
                                            />
                                        </td>
                                    )}
                                    {columns.map((column, colIndex) => (
                                        <td
                                            key={colIndex}
                                            className="px-6 py-4 text-sm text-gray-900"
                                        >
                                            {column.render ? column.render(row) : row[column.key]}
                                        </td>
                                    ))}
                                    {showActions && (
                                        <td className="px-6 py-4">
                                            <div className="flex items-center gap-3">
                                                {onEdit && (
                                                    <button
                                                        onClick={() => onEdit(row, rowIndex)}
                                                        className="text-gray-500 hover:text-blue-600 transition-colors"
                                                        title="Edit"
                                                    >
                                                        <FiEdit2 size={18} />
                                                    </button>
                                                )}
                                                {onDownload && (
                                                    <button
                                                        onClick={() => onDownload(row, rowIndex)}
                                                        className="text-gray-500 hover:text-green-600 transition-colors"
                                                        title="Download"
                                                    >
                                                        <FiDownload size={18} />
                                                    </button>
                                                )}
                                                {onDelete && (
                                                    <button
                                                        onClick={() => onDelete(row, rowIndex)}
                                                        className="text-red-600 hover:text-red-800 transition-colors"
                                                        title="Delete"
                                                    >
                                                        <FiTrash2 size={18} />
                                                    </button>
                                                )}
                                            </div>
                                        </td>
                                    )}
                                </tr>
                            ))
                        )}
                    </tbody>
                </table>
            </div>
        </div>
    )
}

export default DataTable

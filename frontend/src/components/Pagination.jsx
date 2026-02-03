import React from 'react'
import { FiChevronLeft, FiChevronRight } from 'react-icons/fi'

const Pagination = ({
    currentPage,
    totalPages,
    onPageChange,
    totalItems,
    itemsPerPage
}) => {
    // Generate page numbers
    const getPageNumbers = () => {
        const pages = []
        const maxVisiblePages = 5

        if (totalPages <= maxVisiblePages) {
            for (let i = 1; i <= totalPages; i++) {
                pages.push(i)
            }
        } else {
            // Always show first page
            pages.push(1)

            // Calculate start and end of visible pages around current page
            let start = Math.max(2, currentPage - 1)
            let end = Math.min(totalPages - 1, currentPage + 1)

            // Adjust if we are near the beginning
            if (currentPage <= 3) {
                end = 4
            }

            // Adjust if we are near the end
            if (currentPage >= totalPages - 2) {
                start = totalPages - 3
            }

            // Add ellipsis if needed
            if (start > 2) {
                pages.push('...')
            }

            // Add visible pages
            for (let i = start; i <= end; i++) {
                pages.push(i)
            }

            // Add ellipsis if needed
            if (end < totalPages - 1) {
                pages.push('...')
            }

            // Always show last page
            if (totalPages > 1) {
                pages.push(totalPages)
            }
        }

        return pages
    }

    if (totalPages <= 1) return null

    return (
        <div className="flex items-center justify-between px-6 py-3 border-t border-gray-200 bg-white">
            <div className="flex items-center text-sm text-gray-500">
                <span className="mr-1">Showing</span>
                <span className="font-medium text-gray-900">
                    {Math.min((currentPage - 1) * itemsPerPage + 1, totalItems)}
                </span>
                <span className="mx-1">to</span>
                <span className="font-medium text-gray-900">
                    {Math.min(currentPage * itemsPerPage, totalItems)}
                </span>
                <span className="mx-1">of</span>
                <span className="font-medium text-gray-900">{totalItems}</span>
                <span className="ml-1">results</span>
            </div>

            <div className="flex items-center gap-2">
                <button
                    onClick={() => onPageChange(currentPage - 1)}
                    disabled={currentPage === 1}
                    className="p-2 border border-gray-300 rounded-md text-gray-600 hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
                >
                    <FiChevronLeft size={16} />
                </button>

                <div className="flex items-center gap-1">
                    {getPageNumbers().map((page, index) => (
                        <button
                            key={index}
                            onClick={() => typeof page === 'number' ? onPageChange(page) : null}
                            disabled={page === '...'}
                            className={`min-w-[32px] h-8 px-2 rounded-md text-sm font-medium transition-colors ${page === currentPage
                                    ? 'bg-[#1A72B9] text-white'
                                    : page === '...'
                                        ? 'text-gray-500 cursor-default'
                                        : 'text-gray-700 hover:bg-gray-100 border border-gray-300'
                                }`}
                        >
                            {page}
                        </button>
                    ))}
                </div>

                <button
                    onClick={() => onPageChange(currentPage + 1)}
                    disabled={currentPage === totalPages}
                    className="p-2 border border-gray-300 rounded-md text-gray-600 hover:bg-gray-50 disabled:opacity-50 disabled:cursor-not-allowed transition-colors"
                >
                    <FiChevronRight size={16} />
                </button>
            </div>
        </div>
    )
}

export default Pagination

/**
 * UsedRowTracker - Utility class to track which raw data row indices are used
 * across all processors during calculation sheet generation.
 * 
 * This ensures 100% precision in identifying unused rows by having each
 * processor mark the exact rows it processes.
 */
class UsedRowTracker {
    constructor() {
        this.usedIndices = new Set()
    }

    /**
     * Mark a single row index as used
     * @param {number} rowIndex - The 0-based index in the rawDataRows array (not Excel row number)
     */
    markUsed(rowIndex) {
        if (typeof rowIndex === 'number' && !isNaN(rowIndex)) {
            this.usedIndices.add(rowIndex)
        }
    }

    /**
     * Mark multiple row indices as used
     * @param {number[]} indices - Array of 0-based indices
     */
    markMultipleUsed(indices) {
        if (Array.isArray(indices)) {
            indices.forEach(i => this.markUsed(i))
        }
    }

    /**
     * Check if a row index has been marked as used
     * @param {number} rowIndex - The 0-based index
     * @returns {boolean}
     */
    isUsed(rowIndex) {
        return this.usedIndices.has(rowIndex)
    }

    /**
     * Get all used row indices
     * @returns {Set<number>}
     */
    getUsedIndices() {
        return this.usedIndices
    }

    /**
     * Get unused rows from raw data
     * @param {Array[]} rawDataRows - Array of rows (excluding header)
     * @param {Array} headers - Column headers
     * @returns {Array<{rowIndex: number, rowData: Array, isUsed: boolean}>}
     */
    getUnusedRows(rawDataRows, headers) {
        const unused = []

        if (!rawDataRows || !Array.isArray(rawDataRows)) {
            return unused
        }

        rawDataRows.forEach((row, index) => {
            // Only include rows that:
            // 1. Were not marked as used by any processor
            // 2. Have some content (not completely empty)
            if (!this.usedIndices.has(index)) {
                const hasContent = row.some(cell =>
                    cell !== null && cell !== undefined && cell !== ''
                )
                if (hasContent) {
                    unused.push({
                        rowIndex: index,
                        rowData: row,
                        isUsed: false // Default checkbox state
                    })
                }
            }
        })

        return unused
    }

    /**
     * Reset the tracker (useful for testing or re-processing)
     */
    reset() {
        this.usedIndices.clear()
    }

    /**
     * Get statistics about row usage
     * @param {number} totalRows - Total number of raw data rows
     * @returns {{used: number, unused: number, total: number}}
     */
    getStats(totalRows) {
        return {
            used: this.usedIndices.size,
            unused: totalRows - this.usedIndices.size,
            total: totalRows
        }
    }
}

export default UsedRowTracker

import React, { useState, useEffect } from 'react'
import { useNavigate } from 'react-router-dom'
import { format } from 'date-fns'
import toast from 'react-hot-toast'
import Sidebar from '../components/Sidebar'
import TopBar from '../components/TopBar'
import DataTable from '../components/DataTable'
import CreateProposalModal from '../components/CreateProposalModal'
import { DateRangePicker } from '../components/DateRangePicker'
import { FiPlus, FiSearch } from 'react-icons/fi'
import { proposalAPI } from '../services/proposalService'
import { useSidebar } from '../context/SidebarContext'

const Proposals = () => {
    const navigate = useNavigate()
    const { sidebarCollapsed, toggleSidebar } = useSidebar()
    const [isModalOpen, setIsModalOpen] = useState(false)
    const [dateRange, setDateRange] = useState({ from: undefined, to: undefined })
    const [pagination, setPagination] = useState({
        page: 1,
        limit: 10,
        totalPages: 0,
        totalItems: 0
    })
    const [debouncedSearch, setDebouncedSearch] = useState('')
    const [searchQuery, setSearchQuery] = useState('')
    const [isLoading, setIsLoading] = useState(true)
    const [proposals, setProposals] = useState([])

    // Debounce search query
    useEffect(() => {
        const timer = setTimeout(() => {
            setDebouncedSearch(searchQuery)
            setPagination(prev => ({ ...prev, page: 1 }))
        }, 500)

        return () => clearTimeout(timer)
    }, [searchQuery])

    // Update pagination when date range changes
    useEffect(() => {
        setPagination(prev => ({ ...prev, page: 1 }))
    }, [dateRange])

    // Fetch proposals
    useEffect(() => {
        fetchProposals()
    }, [debouncedSearch, dateRange, pagination.page, pagination.limit])

    const fetchProposals = async () => {
        try {
            setIsLoading(true)
            const params = {
                page: pagination.page,
                limit: pagination.limit,
                search: debouncedSearch,
                startDate: dateRange.from?.toISOString(),
                endDate: dateRange.to?.toISOString()
            }

            const response = await proposalAPI.list(params)
            setProposals(response.proposals || [])

            if (response.pagination) {
                setPagination(prev => ({
                    ...prev,
                    totalPages: response.pagination.pages,
                    totalItems: response.pagination.total
                }))
            }
        } catch (error) {
            console.error('Error fetching proposals:', error)
            toast.error('Error loading proposals')
        } finally {
            setIsLoading(false)
        }
    }

    const handlePageChange = (newPage) => {
        setPagination(prev => ({ ...prev, page: newPage }))
    }

    const formatProposalData = (proposals) => {
        return proposals.map(proposal => ({
            id: proposal._id,
            proposalName: proposal.name,
            client: proposal.client,
            project: proposal.project,
            template: proposal.template,
            createdDate: format(new Date(proposal.createdAt), 'MMM dd, yyyy'),
            _original: proposal
        }))
    }

    // Table columns configuration
    const columns = [
        {
            header: 'Proposal Name',
            key: 'proposalName',
        },
        {
            header: 'Client',
            key: 'client'
        },
        {
            header: 'Project',
            key: 'project'
        },
        {
            header: 'Template',
            key: 'template',
            render: (value) => {
                const templateNames = {
                    capstone: 'Capstone',
                    dbi: 'DBI',
                    pd_steel: 'PD Steel',
                    sperrin_tony: 'Sperrin Tony',
                    tristate_martin: 'Tristate Martin'
                }
                return templateNames[value] || value
            }
        },
        {
            header: 'Created Date',
            key: 'createdDate'
        }
    ]

    // Action handlers
    const handleRowClick = (row) => {
        navigate(`/proposals/${row.id}`)
    }

    const handleEdit = (row) => {
        navigate(`/proposals/${row.id}`)
    }

    const handleDownload = (row) => {
        console.log('Download:', row)
        toast.info('Download feature coming soon')
    }

    const handleDelete = async (row) => {
        if (!window.confirm('Are you sure you want to delete this proposal?')) {
            return
        }

        try {
            await proposalAPI.delete(row.id)
            toast.success('Proposal deleted successfully')
            fetchProposals() // Refresh the list
        } catch (error) {
            console.error('Error deleting proposal:', error)
            toast.error(error.response?.data?.message || 'Error deleting proposal')
        }
    }

    const handleBulkDelete = async (selectedProposals) => {
        if (!window.confirm(`Are you sure you want to delete ${selectedProposals.length} proposals?`)) return

        try {
            const ids = selectedProposals.map(p => p.id)
            await proposalAPI.bulkDelete(ids)
            toast.success('Proposals deleted successfully')
            fetchProposals()
        } catch (error) {
            console.error('Error deleting proposals:', error)
            toast.error('Error deleting proposals')
        }
    }

    const handleProposalCreated = () => {
        fetchProposals() // Refresh the list after creating a new proposal
    }

    return (
        <div className="flex h-screen bg-gray-50">
            {/* Sidebar */}
            <Sidebar
                collapsed={sidebarCollapsed}
                onToggle={toggleSidebar}
            />

            {/* Main Content */}
            <div className="flex-1 flex flex-col overflow-hidden">
                {/* Top Bar */}
                <TopBar />

                {/* Page Content */}
                <div className="flex-1 overflow-auto">
                    <div className="p-8">
                        {/* Page Header */}
                        <div className="flex items-start justify-between mb-6">
                            <div>
                                <h1 className="text-2xl font-bold text-gray-900 mb-2">Proposals</h1>
                                <p className="text-sm text-gray-500">
                                    View all proposals and create new ones.
                                </p>
                            </div>
                            <button
                                onClick={() => setIsModalOpen(true)}
                                className="flex items-center gap-2 bg-[#1A72B9] hover:bg-[#1565C0] text-white px-4 py-2.5 rounded-lg transition-colors text-sm font-medium shadow-sm"
                            >
                                <FiPlus size={18} />
                                New Proposal
                            </button>
                        </div>

                        {/* Search and Filter Section */}
                        <div className="bg-white rounded-lg border border-gray-200 p-4 mb-6">
                            <div className="flex items-center gap-4">
                                {/* Search Input */}
                                <div className="flex-1 relative">
                                    <FiSearch className="absolute left-3 top-1/2 -translate-y-1/2 text-gray-400" size={18} />
                                    <input
                                        type="text"
                                        placeholder="Search by name, client, project, or template"
                                        value={searchQuery}
                                        onChange={(e) => setSearchQuery(e.target.value)}
                                        className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
                                    />
                                </div>

                                {/* Date Range Picker */}
                                <DateRangePicker
                                    value={dateRange}
                                    onChange={setDateRange}
                                    placeholder="Select date range"
                                />
                            </div>
                        </div>

                        {/* Loading State */}
                        {isLoading ? (
                            <div className="bg-white rounded-lg border border-gray-200 p-12 text-center">
                                <p className="text-gray-500">Loading proposals...</p>
                            </div>
                        ) : proposals.length === 0 ? (
                            <div className="bg-white rounded-lg border border-gray-200 p-12 text-center">
                                <p className="text-gray-500 mb-2">
                                    {debouncedSearch || dateRange.from
                                        ? 'No proposals found matching your search criteria.'
                                        : 'No proposals yet. Create your first proposal to get started.'}
                                </p>
                                {!debouncedSearch && !dateRange.from && (
                                    <button
                                        onClick={() => setIsModalOpen(true)}
                                        className="inline-flex items-center gap-2 mt-4 text-blue-600 hover:text-blue-700 font-medium"
                                    >
                                        <FiPlus />
                                        Create New Proposal
                                    </button>
                                )}
                            </div>
                        ) : (
                            <DataTable
                                columns={columns}
                                data={formatProposalData(proposals)}
                                onEdit={handleEdit}
                                onDelete={handleDelete}
                                onRowClick={handleRowClick}
                                showCheckbox={true}
                                showActions={true}
                                onBulkDelete={handleBulkDelete}
                                pagination={{
                                    currentPage: pagination.page,
                                    totalPages: pagination.totalPages,
                                    onPageChange: handlePageChange,
                                    totalItems: pagination.totalItems,
                                    itemsPerPage: pagination.limit
                                }}
                            />
                        )}
                    </div>
                </div>
            </div>

            {/* Create Proposal Modal */}
            <CreateProposalModal
                isOpen={isModalOpen}
                onClose={() => setIsModalOpen(false)}
                onSuccess={handleProposalCreated}
            />
        </div>
    )
}

export default Proposals

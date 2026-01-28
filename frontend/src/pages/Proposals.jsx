import React, { useState } from 'react'
import { Link } from 'react-router-dom'
import Sidebar from '../components/Sidebar'
import TopBar from '../components/TopBar'
import DataTable from '../components/DataTable'
import CreateProposalModal from '../components/CreateProposalModal'
import { FiPlus, FiSearch, FiCalendar, FiHome } from 'react-icons/fi'

const Proposals = () => {
    const [sidebarCollapsed, setSidebarCollapsed] = useState(false)
    const [isModalOpen, setIsModalOpen] = useState(false)

    // Sample data matching the design
    const proposalData = [
        {
            id: 1,
            proposalName: 'Proposal 1',
            client: 'Capstone Contracting Corp.',
            project: 'Project 1',
            createdDate: 'Jan 05, 2026'
        },
        {
            id: 2,
            proposalName: 'Proposal 2',
            client: 'PD Steel',
            project: 'Project 1',
            createdDate: 'Jan 05, 2026'
        },
        {
            id: 3,
            proposalName: 'Proposal 3',
            client: 'DBI',
            project: 'Project 1',
            createdDate: 'Jan 05, 2026'
        }
    ]

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
            header: 'Created Date',
            key: 'createdDate'
        }
    ]

    // Action handlers
    const handleEdit = (row) => {
        console.log('Edit:', row)
    }

    const handleDownload = (row) => {
        console.log('Download:', row)
    }

    const handleDelete = (row) => {
        console.log('Delete:', row)
    }

    const handleCreateProposal = (proposalData) => {
        console.log('New proposal created:', proposalData)
        // Add logic to handle the new proposal data
    }

    return (
        <div className="flex h-screen bg-gray-50">
            {/* Sidebar */}
            <Sidebar
                collapsed={sidebarCollapsed}
                onToggle={() => setSidebarCollapsed(!sidebarCollapsed)}
            />

            {/* Main Content */}
            <div className="flex-1 flex flex-col overflow-hidden">
                {/* Top Bar */}
                <TopBar />

                {/* Page Content */}
                <div className="flex-1 overflow-auto">
                    <div className="p-8">
                        {/* Breadcrumb Navigation */}
                        <div className="flex items-center gap-2 text-sm mb-6">
                            <Link to="/dashboard" className="text-gray-500 hover:text-gray-700">
                                <FiHome size={18} />
                            </Link>
                            <span className="text-gray-400">&gt;</span>
                            <span className="text-gray-900 font-medium">Proposals</span>
                        </div>
                        {/* Page Header */}
                        <div className="flex items-start justify-between mb-6">
                            <div>
                                <h1 className="text-2xl font-bold text-gray-900 mb-2">Proposals</h1>
                                <p className="text-sm text-gray-500">
                                    Lorem ipsum dolor sit amet consectetur. Facilisi diam ullamcorper arcu risus laoreet.
                                </p>
                            </div>
                            <button
                                onClick={() => setIsModalOpen(true)}
                                className="flex items-center gap-2 bg-[#1A72B9] hover:bg-[#1565C0] text-white px-4 py-2.5 rounded-lg transition-colors text-sm font-medium shadow-sm"
                            >
                                <FiPlus size={18} />
                                Create New Proposal
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
                                        placeholder="Search here"
                                        className="w-full pl-10 pr-4 py-2 border border-gray-300 rounded-lg focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent text-sm"
                                    />
                                </div>

                                {/* Filter Dropdown */}
                                <div className="relative">
                                    <select className="appearance-none bg-white border border-gray-300 rounded-lg px-4 py-2 pr-10 text-sm text-gray-700 focus:outline-none focus:ring-2 focus:ring-blue-500 focus:border-transparent cursor-pointer">
                                        <option>Last 30 days</option>
                                        <option>Last 7 days</option>
                                        <option>Last 90 days</option>
                                    </select>
                                    <div className="absolute right-3 top-1/2 -translate-y-1/2 pointer-events-none">
                                        <svg className="w-4 h-4 text-gray-400" fill="none" stroke="currentColor" viewBox="0 0 24 24">
                                            <path strokeLinecap="round" strokeLinejoin="round" strokeWidth={2} d="M19 9l-7 7-7-7" />
                                        </svg>
                                    </div>
                                </div>

                                {/* Date Range Display */}
                                <div className="flex items-center gap-2 px-4 py-2 border border-gray-300 rounded-lg text-sm text-gray-700">
                                    <FiCalendar className="text-gray-400" size={16} />
                                    <span>15 May - 15 Jun</span>
                                </div>
                            </div>
                        </div>

                        {/* Data Table */}
                        <DataTable
                            columns={columns}
                            data={proposalData}
                            onEdit={handleEdit}
                            onDownload={handleDownload}
                            onDelete={handleDelete}
                            showCheckbox={true}
                            showActions={true}
                        />
                    </div>
                </div>
            </div>

            {/* Create Proposal Modal */}
            <CreateProposalModal
                isOpen={isModalOpen}
                onClose={() => setIsModalOpen(false)}
                onSubmit={handleCreateProposal}
            />
        </div>
    )
}

export default Proposals

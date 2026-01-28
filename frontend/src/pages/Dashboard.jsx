import React from 'react'
import { Link } from 'react-router-dom'
import { FiUpload, FiFileText } from 'react-icons/fi'

const Dashboard = () => {
  return (
    <div className="min-h-screen bg-gray-50 p-8">
      <div className="max-w-6xl mx-auto">
        <h1 className="text-3xl font-bold mb-8 text-gray-900">Dashboard</h1>

        <div className="grid grid-cols-1 md:grid-cols-2 gap-6">
          {/* Upload New Project Card */}
          <Link
            to="/upload"
            className="block p-6 bg-white rounded-lg shadow-md hover:shadow-lg transition-shadow border-2 border-transparent hover:border-blue-500"
          >
            <div className="flex items-start space-x-4">
              <div className="p-3 bg-blue-100 rounded-lg">
                <FiUpload className="text-2xl text-blue-600" />
              </div>
              <div>
                <h2 className="text-xl font-semibold text-gray-900 mb-2">New Project</h2>
                <p className="text-gray-600">Upload raw Excel data to create a new construction project workbook</p>
              </div>
            </div>
          </Link>

          {/* View Spreadsheet Card */}
          <Link
            to="/spreadsheet"
            className="block p-6 bg-white rounded-lg shadow-md hover:shadow-lg transition-shadow border-2 border-transparent hover:border-green-500"
          >
            <div className="flex items-start space-x-4">
              <div className="p-3 bg-green-100 rounded-lg">
                <FiFileText className="text-2xl text-green-600" />
              </div>
              <div>
                <h2 className="text-xl font-semibold text-gray-900 mb-2">View Spreadsheet</h2>
                <p className="text-gray-600">Open the spreadsheet view directly (empty workbook)</p>
              </div>
            </div>
          </Link>
        </div>

        {/* Info Section */}
        <div className="mt-8 bg-blue-50 border border-blue-200 rounded-lg p-6">
          <h3 className="font-semibold text-blue-900 mb-3">How it works:</h3>
          <ol className="space-y-2 text-blue-800">
            <li>1. Click <strong>New Project</strong> to upload your raw Excel data</li>
            <li>2. Preview and verify your data</li>
            <li>3. The system will generate Calculations Sheet and Proposal Sheet</li>
            <li>4. Export your completed workbook</li>
          </ol>
        </div>
      </div>
    </div>
  )
}

export default Dashboard
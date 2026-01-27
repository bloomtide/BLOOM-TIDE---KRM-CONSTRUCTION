import React, { useState } from 'react'
import { useLocation, useNavigate } from 'react-router-dom'
import { FiArrowLeft, FiArrowRight, FiCheck } from 'react-icons/fi'

const templates = [
  {
    id: 'capstone',
    name: 'Capstone Template',
    description: 'Standard construction template with comprehensive sections for excavation, foundation, and superstructure work',
    active: true
  },
  {
    id: 'dbi',
    name: 'DBI Template',
    description: 'Coming soon',
    active: false
  },
  {
    id: 'pd_steel',
    name: 'PD Steel Template',
    description: 'Coming soon',
    active: false
  },
  {
    id: 'sperrin_tony',
    name: 'Sperrin Tony Template',
    description: 'Coming soon',
    active: false
  },
  {
    id: 'tristate_martin',
    name: 'Tristate Martin Template',
    description: 'Coming soon',
    active: false
  }
]

const TemplateSelector = () => {
  const location = useLocation()
  const navigate = useNavigate()
  const { rawData, fileName, sheetName } = location.state || {}
  const [selectedTemplate, setSelectedTemplate] = useState('capstone')

  const handleGoBack = () => {
    navigate('/upload')
  }

  const handleProceed = () => {
    if (!selectedTemplate) {
      alert('Please select a template')
      return
    }

    // Navigate to preview with template selection
    navigate('/preview', {
      state: {
        rawData,
        fileName,
        sheetName,
        selectedTemplate
      }
    })
  }

  if (!rawData) {
    navigate('/upload')
    return null
  }

  return (
    <div className="min-h-screen bg-gray-50 flex flex-col">
      {/* Header */}
      <div className="bg-white border-b shadow-sm">
        <div className="max-w-7xl mx-auto px-6 py-4">
          <div className="flex items-center justify-between">
            <div>
              <h1 className="text-2xl font-bold text-gray-900">Select Template</h1>
              <p className="text-sm text-gray-600 mt-1">
                Choose a calculation template for: {fileName}
              </p>
            </div>
            <div className="flex space-x-3">
              <button
                onClick={handleGoBack}
                className="flex items-center space-x-2 px-4 py-2 bg-gray-200 text-gray-700 rounded-lg hover:bg-gray-300 transition-colors"
              >
                <FiArrowLeft />
                <span>Go Back</span>
              </button>
              <button
                onClick={handleProceed}
                disabled={!selectedTemplate}
                className={`flex items-center space-x-2 px-6 py-2 rounded-lg transition-colors ${
                  selectedTemplate
                    ? 'bg-blue-600 text-white hover:bg-blue-700'
                    : 'bg-gray-300 text-gray-500 cursor-not-allowed'
                }`}
              >
                <span>Continue</span>
                <FiArrowRight />
              </button>
            </div>
          </div>
        </div>
      </div>

      {/* Template Grid */}
      <div className="flex-1 p-6">
        <div className="max-w-7xl mx-auto">
          <div className="grid grid-cols-1 md:grid-cols-2 lg:grid-cols-3 gap-6">
            {templates.map((template) => (
              <button
                key={template.id}
                onClick={() => template.active && setSelectedTemplate(template.id)}
                disabled={!template.active}
                className={`relative p-6 rounded-lg border-2 text-left transition-all ${
                  selectedTemplate === template.id
                    ? 'border-blue-600 bg-blue-50 shadow-md'
                    : template.active
                    ? 'border-gray-300 bg-white hover:border-gray-400 hover:shadow-md'
                    : 'border-gray-200 bg-gray-100 opacity-60 cursor-not-allowed'
                }`}
              >
                {selectedTemplate === template.id && (
                  <div className="absolute top-4 right-4 bg-blue-600 text-white rounded-full p-1">
                    <FiCheck className="text-lg" />
                  </div>
                )}
                
                <h3 className="text-xl font-bold text-gray-900 mb-2 pr-10">
                  {template.name}
                </h3>
                <p className="text-sm text-gray-600">{template.description}</p>
                
                {!template.active && (
                  <div className="mt-3 inline-block px-2 py-1 bg-gray-300 text-gray-600 text-xs font-semibold rounded">
                    Coming Soon
                  </div>
                )}
              </button>
            ))}
          </div>

          {/* Template Info */}
          {selectedTemplate === 'capstone' && (
            <div className="mt-8 bg-blue-50 border border-blue-200 rounded-lg p-6">
              <h4 className="font-semibold text-blue-900 mb-3">Capstone Template Includes:</h4>
              <div className="grid grid-cols-2 md:grid-cols-4 gap-3 text-sm text-blue-800">
                <div>• Demolition</div>
                <div>• Excavation</div>
                <div>• Rock Excavation</div>
                <div>• SOE</div>
                <div>• Foundation</div>
                <div>• Waterproofing</div>
                <div>• Superstructure</div>
                <div>• B.P.P.</div>
                <div>• Civil/Sitework</div>
              </div>
            </div>
          )}
        </div>
      </div>
    </div>
  )
}

export default TemplateSelector

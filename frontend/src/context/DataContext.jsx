import React, { createContext, useState, useContext } from 'react'

const DataContext = createContext()

export const DataProvider = ({ children }) => {
  const [rawData, setRawData] = useState(null)
  const [calculationsData, setCalculationsData] = useState(null)
  const [proposalData, setProposalData] = useState(null)
  const [projectName, setProjectName] = useState('')
  const [fileName, setFileName] = useState('')
  const [selectedTemplate, setSelectedTemplate] = useState(null)

  const clearData = () => {
    setRawData(null)
    setCalculationsData(null)
    setProposalData(null)
    setProjectName('')
    setFileName('')
    setSelectedTemplate(null)
  }

  return (
    <DataContext.Provider
      value={{
        rawData,
        setRawData,
        calculationsData,
        setCalculationsData,
        proposalData,
        setProposalData,
        projectName,
        setProjectName,
        fileName,
        setFileName,
        selectedTemplate,
        setSelectedTemplate,
        clearData,
      }}
    >
      {children}
    </DataContext.Provider>
  )
}

export const useData = () => {
  const context = useContext(DataContext)
  if (!context) {
    throw new Error('useData must be used within a DataProvider')
  }
  return context
}

export default DataContext

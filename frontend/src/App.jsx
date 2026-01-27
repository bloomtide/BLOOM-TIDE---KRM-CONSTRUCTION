import React from 'react'
import { Routes, Route } from 'react-router-dom'

import Login from './pages/Login.jsx'
import UpdatePassword from './pages/UpdatePassword.jsx'
import Dashboard from './pages/Dashboard.jsx'
import UploadExcel from './pages/UploadExcel.jsx'
import TemplateSelector from './pages/TemplateSelector.jsx'
import PreviewData from './pages/PreviewData.jsx'
import Spreadsheet from './pages/Spreadsheet.jsx'

function App() {
  return (
    <div>
        <Routes>
          <Route path="/" element={<Login />} />
          <Route path="update-password/:token" element={ <UpdatePassword />} />
          <Route path="dashboard" element={<Dashboard />} />
          <Route path="/upload" element={<UploadExcel />} />
          <Route path="/template-select" element={<TemplateSelector />} />
          <Route path="/preview" element={<PreviewData />} />
          <Route path="/spreadsheet" element={<Spreadsheet />} />
        </Routes>
    </div>
  )
}

export default App
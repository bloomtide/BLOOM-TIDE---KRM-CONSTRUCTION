import React from 'react'
import { Routes, Route } from 'react-router-dom'
import { AuthProvider } from './context/AuthContext'

import Login from './pages/Login.jsx'
import UpdatePassword from './pages/UpdatePassword.jsx'
import Dashboard from './pages/Dashboard.jsx'
import Proposals from './pages/Proposals.jsx'
import Users from './pages/Users.jsx'
import UploadExcel from './pages/UploadExcel.jsx'
import TemplateSelector from './pages/TemplateSelector.jsx'
import PreviewData from './pages/PreviewData.jsx'
import Spreadsheet from './pages/Spreadsheet.jsx'

function App() {
  return (
    <AuthProvider>
      <Routes>
        <Route path="/" element={<Login />} />
        <Route path="update-password/:token" element={<UpdatePassword />} />
        <Route path="dashboard" element={<Dashboard />} />
        <Route path="proposals" element={<Proposals />} />
        <Route path="users" element={<Users />} />
        <Route path="/upload" element={<UploadExcel />} />
        <Route path="/template-select" element={<TemplateSelector />} />
        <Route path="/preview" element={<PreviewData />} />
        <Route path="/spreadsheet" element={<Spreadsheet />} />
      </Routes>
    </AuthProvider>
  )
}

export default App
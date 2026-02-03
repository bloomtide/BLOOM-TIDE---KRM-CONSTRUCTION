import React from 'react'
import { Routes, Route } from 'react-router-dom'
import { AuthProvider } from './context/AuthContext'
import { SidebarProvider } from './context/SidebarContext'

import Login from './pages/Login.jsx'
import Proposals from './pages/Proposals.jsx'
import ProposalDetail from './pages/ProposalDetail.jsx'
import Users from './pages/Users.jsx'

function App() {
  return (
    <AuthProvider>
      <SidebarProvider>
        <Routes>
          <Route path="/" element={<Login />} />
          <Route path="proposals" element={<Proposals />} />
          <Route path="proposals/:proposalId" element={<ProposalDetail />} />
          <Route path="users" element={<Users />} />
        </Routes>
      </SidebarProvider>
    </AuthProvider>
  )
}

export default App
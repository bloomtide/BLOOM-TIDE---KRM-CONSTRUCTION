import React from 'react'
import { Routes, Route } from 'react-router-dom'

import Login from './pages/Login.jsx'
import UpdatePassword from './pages/UpdatePassword.jsx'
import Dashboard from './pages/Dashboard.jsx'

function App() {
  return (
    <div>
        <Routes>
          <Route path="/" element={<Login />} />
          <Route path="update-password/:token" element={ <UpdatePassword />} />
          <Route path="dashboard" element={<Dashboard />} />
        </Routes>
    </div>
  )
}

export default App
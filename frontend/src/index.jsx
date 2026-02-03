import React from 'react'
import { createRoot } from 'react-dom/client'
import { BrowserRouter } from 'react-router-dom'
import { Toaster } from 'react-hot-toast'
import { registerLicense } from '@syncfusion/ej2-base'
import { DataProvider } from './context/DataContext.jsx'
import './index.css'
import App from './App.jsx'

// Register Syncfusion license
registerLicense(import.meta.env.VITE_APP_SYNCFUSION_LICENSE_KEY)


createRoot(document.getElementById('root')).render(
  <BrowserRouter>
    <DataProvider>
      <App />
      <Toaster position="top-center" />
    </DataProvider>
  </BrowserRouter>
)
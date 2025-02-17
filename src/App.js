// src/App.js
import React from 'react';
import ProtectedDashboard from './components/ProtectedDashboard';
import { AppProvider } from './context/AppContext';
import { ThemeProvider } from '@mui/material/styles';
import { darkTheme } from './theme';
import { CssBaseline } from '@mui/material';
import './App.css';

function App() {
  return (
    <AppProvider>
      <ThemeProvider theme={darkTheme}>
        {/* CssBaseline setzt die globalen Styles für Dark Mode */}
        <CssBaseline />
        {/* ProtectedDashboard schützt das gesamte Dashboard */}
        <ProtectedDashboard />
      </ThemeProvider>
    </AppProvider>
  );
}

export default App;

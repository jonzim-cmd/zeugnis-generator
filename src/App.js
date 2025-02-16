import React from 'react';
import Dashboard from './components/Dashboard';
import { AppProvider } from './context/AppContext';
import { ThemeProvider } from '@mui/material/styles';
import { darkTheme } from './theme';
import './App.css';

function App() {
  return (
    <AppProvider>
      <ThemeProvider theme={darkTheme}>
        <Dashboard />
      </ThemeProvider>
    </AppProvider>
  );
}

export default App;

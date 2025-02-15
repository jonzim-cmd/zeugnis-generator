// src/context/AppContext.js
import React, { createContext, useState } from 'react';

export const AppContext = createContext();

export const AppProvider = ({ children }) => {
  const [excelData, setExcelData] = useState([]);
  const [dashboardData, setDashboardData] = useState({
    klassenleitung: '',
    schulleitung: '',
    schuljahr: '',
    datum: '',
    zeugnisart: 'Jahreszeugnis'  // Optionen: "Zwischenzeugnis", "Jahreszeugnis", "Abschlusszeugnis"
  });

  return (
    <AppContext.Provider value={{ excelData, setExcelData, dashboardData, setDashboardData }}>
      {children}
    </AppContext.Provider>
  );
};

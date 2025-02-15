// src/context/AppContext.js
import React, { createContext, useState } from 'react';

export const AppContext = createContext();

export const AppProvider = ({ children }) => {
  const [excelData, setExcelData] = useState([]);
  const [dashboardData, setDashboardData] = useState({
    klassenleitung: '',     // Klassenleitung aus dem Dashboard
    schulleitung: '',       // Schulleitung aus dem Dashboard
    schuljahr: '',          // Schuljahr aus dem Dashboard
    datum: '',              // Datum (falls benötigt)
    zeugnisart: 'Jahreszeugnis',  // Optionen: "Zwischenzeugnis", "Jahreszeugnis", "Abschlusszeugnis"
    KL: ''                  // Klasse aus dem Dashboard (wird als "KL" in der WordTemplateProcessor erwartet)
  });

  return (
    <AppContext.Provider value={{ excelData, setExcelData, dashboardData, setDashboardData }}>
      {children}
    </AppContext.Provider>
  );
};

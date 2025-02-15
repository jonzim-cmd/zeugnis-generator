import React, { createContext, useState } from 'react';

// Der zentrale State-Context für Excel-Daten und Dashboard-Eingaben
export const AppContext = createContext();

export const AppProvider = ({ children }) => {
  const [excelData, setExcelData] = useState([]);
  const [dashboardData, setDashboardData] = useState({
    klassenleitung: '',
    schulleitung: '',
    schuljahr: '',
    datum: '',
    zeugnisart: 'Jahreszeugnis' // Standardwert
  });

  return (
    <AppContext.Provider value={{ excelData, setExcelData, dashboardData, setDashboardData }}>
      {children}
    </AppContext.Provider>
  );
};

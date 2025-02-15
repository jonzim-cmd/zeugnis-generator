import React, { createContext, useState } from 'react';

export const AppContext = createContext();

export const AppProvider = ({ children }) => {
  const [excelData, setExcelData] = useState([]);
  const [dashboardData, setDashboardData] = useState({
    klassenleitung: '',
    schulleitung: '',
    sl_titel: '',
    kl_titel: '',
    schuljahr: '',
    datum: '',
    zeugnisart: 'Jahreszeugnis',
    KL: ''
  });

  return (
    <AppContext.Provider value={{ excelData, setExcelData, dashboardData, setDashboardData }}>
      {children}
    </AppContext.Provider>
  );
};

import React, { createContext, useState } from 'react';

export const AppContext = createContext();

export const AppProvider = ({ children }) => {
  const [excelData, setExcelData] = useState([]);
  const [dashboardData, setDashboardData] = useState({
    klassenleitung: '',     // Platzhalter: {{Klassenleitung}}
    schulleitung: '',       // Platzhalter: {{Schulleitung}}
    schuljahr: '',          // Platzhalter: {{SJ}}
    datum: '',              // Wird für {{Zeugnisdatum}} genutzt
    zeugnisart: 'Jahreszeugnis',  // Optionen: "Zwischenzeugnis", "Jahreszeugnis", "Abschlusszeugnis"
    KL: '',                 // Klasse (wird als "KL" in Word erwartet)
    sl_titel: ''            // Funktionsbezeichnung (Schulleitung & Klassenleitung) – Platzhalter: {{Sl_Titel}}
  });

  return (
    <AppContext.Provider value={{ excelData, setExcelData, dashboardData, setDashboardData }}>
      {children}
    </AppContext.Provider>
  );
};

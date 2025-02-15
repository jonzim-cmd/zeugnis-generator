import React, { useContext } from 'react';
import { Button, Typography } from '@mui/material';
import * as XLSX from 'xlsx';
import { AppContext } from '../context/AppContext';

const ExcelUpload = () => {
  const { setExcelData, setDashboardData } = useContext(AppContext);

  const handleFileUpload = (event) => {
    const file = event.target.files[0];
    if (!file) return;

    const reader = new FileReader();
    reader.onload = (evt) => {
      const data = evt.target.result;
      const workbook = XLSX.read(data, { type: 'binary' });
      const sheetName = workbook.SheetNames[0];
      const worksheet = workbook.Sheets[sheetName];
      const jsonData = XLSX.utils.sheet_to_json(worksheet, { defval: '' });
      setExcelData(jsonData);

      // Falls im ersten Datensatz "Zeugnisart" den Wert "ZZ" hat,
      // wird im Dashboard die Auswahl auf "Zwischenzeugnis" voreingestellt.
      if (jsonData.length > 0 && jsonData[0].Zeugnisart === "ZZ") {
        setDashboardData(prev => ({ ...prev, zeugnisart: "Zwischenzeugnis" }));
      }
    };
    reader.readAsBinaryString(file);
  };

  return (
    <div>
      <Typography variant="h6" gutterBottom>
        Excel-Datei hochladen
      </Typography>
      <input
        type="file"
        accept=".xlsx, .xls"
        onChange={handleFileUpload}
        style={{ marginBottom: '20px' }}
      />
      <Typography variant="body2" color="textSecondary">
        Bitte stellen Sie sicher, dass die Excel-Datei folgende Spalten enthält: Klasse, Vorname, Nachname, Geburtsdatum, Geburtsort, Fächer, Noten der Fächer, Zeugnisbemerkung 1 und Zeugnisbemerkung 2.
      </Typography>
    </div>
  );
};

export default ExcelUpload;

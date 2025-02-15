import React from 'react';
import { Box, Typography } from '@mui/material';

/*
  CertificateTemplate-Komponente:
  - Erzeugt das Layout für ein Zeugnis.
  - Nutzt sowohl die Schülerdaten (aus der Excel-Datei) als auch die Dashboard-Daten.
  - Je nach ausgewählter Zeugnisart (Zwischen-, Jahres- oder Abschlusszeugnis) 
    wird ein leicht modifiziertes Layout dargestellt – etwa in der Rahmenfarbe.
  - Das Layout orientiert sich exakt an den vorgegebenen Templates.
*/
const CertificateTemplate = ({ student, dashboardData }) => {
  const { zeugnisart } = dashboardData;

  // Auswahl des Styles je nach Zeugnisart
  const getTemplateStyle = () => {
    switch (zeugnisart) {
      case 'Zwischenzeugnis':
        return { border: '2px solid blue', padding: '20px' };
      case 'Jahreszeugnis':
        return { border: '2px solid green', padding: '20px' };
      case 'Abschlusszeugnis':
        return { border: '2px solid red', padding: '20px' };
      default:
        return { border: '2px solid black', padding: '20px' };
    }
  };

  return (
    // A4-Größe (210 x 297 mm) – angepasst über CSS (hier als Pixel simuliert)
    <Box width={210} height={297} style={getTemplateStyle()} sx={{ position: 'relative', fontFamily: 'Arial, sans-serif' }}>
      {/* Header-Bereich mit Angaben aus dem Dashboard */}
      <Box textAlign="center" mb={2}>
        <Typography variant="h5">{zeugnisart}</Typography>
        <Typography variant="subtitle1">{dashboardData.schuljahr}</Typography>
        <Typography variant="subtitle2">{dashboardData.datum}</Typography>
      </Box>
      {/* Schüler-Informationen */}
      <Box mb={2}>
        <Typography variant="body1">
          <strong>Klasse:</strong> {student.Klasse}
        </Typography>
        <Typography variant="body1">
          <strong>Name:</strong> {student.Vorname} {student.Nachname}
        </Typography>
        <Typography variant="body1">
          <strong>Geburtsdatum:</strong> {student.Geburtsdatum}
        </Typography>
        <Typography variant="body1">
          <strong>Geburtsort:</strong> {student.Geburtsort}
        </Typography>
      </Box>
      {/* Darstellung der Fächer und Noten */}
      <Box mb={2}>
        <Typography variant="body1"><strong>Fächer & Noten:</strong></Typography>
        <Typography variant="body2">{student.Fächer}</Typography>
        <Typography variant="body2">{student['Noten der Fächer']}</Typography>
      </Box>
      {/* Zeugnisbemerkungen */}
      <Box mb={2}>
        <Typography variant="body1"><strong>Bemerkungen:</strong></Typography>
        <Typography variant="body2">{student['Zeugnisbemerkung 1']}</Typography>
        <Typography variant="body2">{student['Zeugnisbemerkung 2']}</Typography>
      </Box>
      {/* Footer-Bereich mit Angaben zur Klassen- und Schulleitung */}
      <Box position="absolute" bottom={20} left={20} right={20} textAlign="center">
        <Typography variant="body2">
          Klassenleitung: {dashboardData.klassenleitung} | Schulleitung: {dashboardData.schulleitung}
        </Typography>
      </Box>
    </Box>
  );
};

export default CertificateTemplate;

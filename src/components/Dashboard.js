import React, { useContext, useState } from 'react';
import { Container, Grid, Typography, Paper, Box, Divider } from '@mui/material';
import { AppContext } from '../context/AppContext';
import ExcelUpload from './ExcelUpload';
import TemplateImport from './TemplateImport';
import WordTemplateProcessor from './WordTemplateProcessor';

const Dashboard = () => {
  const { dashboardData, setDashboardData, excelData } = useContext(AppContext);
  const [customTemplate, setCustomTemplate] = useState(null);

  const handleChange = (e) => {
    setDashboardData({ ...dashboardData, [e.target.name]: e.target.value });
  };

  // Optionen für Funktionsbezeichnungen
  const funktionsOptions = [
    'StR',
    'StRin',
    'OStR',
    'OStRin',
    'StD',
    'StDin',
    'OStD',
    'OStDin'
  ];

  return (
    <Container maxWidth="md" sx={{ mt: 4, mb: 4 }}>
      <Paper elevation={3} sx={{ p: 3 }}>
        <Typography 
          variant="h4" 
          align="center" 
          gutterBottom 
          sx={{ fontWeight: 'bold' }}>
          Zeugnisgenerator
        </Typography>

        {/* Section: Excel Upload */}
        <Box sx={{ mt: 3, p: 2, backgroundColor: '#e3f2fd', borderRadius: 1 }}>
          <Typography variant="h6" gutterBottom>
            Excel-Datei hochladen
          </Typography>
          <ExcelUpload />
        </Box>

        <Divider sx={{ my: 3 }} />

        {/* Section: Dashboard Eingaben */}
        <Box sx={{ mt: 3, p: 2, backgroundColor: '#f1f8e9', borderRadius: 1 }}>
          <Typography variant="h6" gutterBottom>
            Dashboard Eingaben
          </Typography>
          <Grid container spacing={2}>
            {/* Klassenleitung & Funktionsbezeichnung Klassenleitung */}
            <Grid item xs={12} sm={6}>
              <Typography variant="subtitle1">Klassenleitung</Typography>
              <input
                type="text"
                name="klassenleitung"
                value={dashboardData.klassenleitung || ''}
                onChange={handleChange}
                placeholder="Klassenleitung"
                style={{ width: '100%', padding: '8px', boxSizing: 'border-box' }}
              />
            </Grid>
            <Grid item xs={12} sm={6}>
              <Typography variant="subtitle1">Funktionsbezeichnung Klassenleitung</Typography>
              <select
                name="kl_titel"
                value={dashboardData.kl_titel || ''}
                onChange={handleChange}
                style={{ width: '100%', padding: '8px', boxSizing: 'border-box' }}
              >
                <option value="">Bitte wählen</option>
                {funktionsOptions.map((option) => (
                  <option key={option} value={option}>
                    {option}
                  </option>
                ))}
              </select>
            </Grid>
            {/* Schulleitung & Funktionsbezeichnung Schulleitung */}
            <Grid item xs={12} sm={6}>
              <Typography variant="subtitle1">Schulleitung</Typography>
              <input
                type="text"
                name="schulleitung"
                value={dashboardData.schulleitung || ''}
                onChange={handleChange}
                placeholder="Schulleitung"
                style={{ width: '100%', padding: '8px', boxSizing: 'border-box' }}
              />
            </Grid>
            <Grid item xs={12} sm={6}>
              <Typography variant="subtitle1">Funktionsbezeichnung Schulleitung</Typography>
              <select
                name="sl_titel"
                value={dashboardData.sl_titel || ''}
                onChange={handleChange}
                style={{ width: '100%', padding: '8px', boxSizing: 'border-box' }}
              >
                <option value="">Bitte wählen</option>
                {funktionsOptions.map((option) => (
                  <option key={option} value={option}>
                    {option}
                  </option>
                ))}
              </select>
            </Grid>
            {/* Weitere Felder */}
            <Grid item xs={12} sm={6}>
              <Typography variant="subtitle1">Schuljahr</Typography>
              <input
                type="text"
                name="schuljahr"
                value={dashboardData.schuljahr || ''}
                onChange={handleChange}
                placeholder="Schuljahr"
                style={{ width: '100%', padding: '8px', boxSizing: 'border-box' }}
              />
            </Grid>
            <Grid item xs={12} sm={6}>
              <Typography variant="subtitle1">Zeugnisdatum</Typography>
              <input
                type="date"
                name="datum"
                value={dashboardData.datum || ''}
                onChange={handleChange}
                style={{ width: '100%', padding: '8px', boxSizing: 'border-box' }}
              />
            </Grid>
            <Grid item xs={12}>
              <Typography variant="subtitle1">Zeugnisart</Typography>
              <select
                name="zeugnisart"
                value={dashboardData.zeugnisart || ''}
                onChange={handleChange}
                style={{ width: '100%', padding: '8px', boxSizing: 'border-box' }}
              >
                <option value="">Bitte wählen</option>
                <option value="Zwischenzeugnis">Zwischenzeugnis</option>
                <option value="Jahreszeugnis">Jahreszeugnis</option>
                <option value="Abschlusszeugnis">Abschlusszeugnis</option>
              </select>
            </Grid>
          </Grid>
        </Box>

        <Divider sx={{ my: 3 }} />

        {/* Section: Word-Template Upload */}
        <Box sx={{ mt: 3, p: 2, backgroundColor: '#ede7f6', borderRadius: 1 }}>
          <Typography variant="h6" gutterBottom>
            Word-Template Upload
          </Typography>
          <TemplateImport onTemplateLoaded={setCustomTemplate} />
        </Box>

        <Divider sx={{ my: 3 }} />

        {/* Section: Dokumentgenerierung */}
        <Box sx={{ mt: 3, p: 2, backgroundColor: '#fff3e0', borderRadius: 1, textAlign: 'center' }}>
          <Typography variant="h6" gutterBottom>
            Dokumentgenerierung
          </Typography>
          <WordTemplateProcessor
            dashboardData={dashboardData}
            customTemplate={customTemplate}
            excelData={excelData && excelData.length > 0 ? excelData : [{ KL: '', gdat: '' }]}
          />
        </Box>
      </Paper>
    </Container>
  );
};

export default Dashboard;

import React, { useContext, useState } from 'react';
import { Container, Grid, Typography, Paper, Box, Divider } from '@mui/material';
import { AppContext } from '../context/AppContext';
import TemplateImport from './TemplateImport';
import WordTemplateProcessor from './WordTemplateProcessor';

const Dashboard = () => {
  const { dashboardData, setDashboardData } = useContext(AppContext);
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
        <Typography variant="h4" align="center" gutterBottom>
          Zeugnis Generator
        </Typography>

        {/* Section 1: Zusätzliche Eingaben */}
        <Box sx={{ mt: 3 }}>
          <Typography variant="h6" gutterBottom>
            Zusätzliche Eingaben
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

        {/* Section 2: Upload & Generierung */}
        <Box sx={{ mt: 3 }}>
          <Typography variant="h6" gutterBottom>
            Word-Template Upload & Dokumentgenerierung
          </Typography>
          <TemplateImport onTemplateLoaded={setCustomTemplate} />
          <Box sx={{ mt: 2 }}>
            {/* Hier übergeben wir ein Dummy-Datensatz-Array (eine leeres Objekt) an WordTemplateProcessor */}
            <WordTemplateProcessor
              dashboardData={dashboardData}
              customTemplate={customTemplate}
              excelData={[{ KL: '', gdat: '' }]}
            />
          </Box>
        </Box>
      </Paper>
    </Container>
  );
};

export default Dashboard;

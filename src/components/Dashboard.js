import React, { useContext, useState } from 'react';
import { TextField, MenuItem, Grid, Typography } from '@mui/material';
import { AppContext } from '../context/AppContext';
import TemplateImport from './TemplateImport';
import WordTemplateProcessor from './WordTemplateProcessor';

const Dashboard = () => {
  const { dashboardData, setDashboardData } = useContext(AppContext);
  const [customTemplate, setCustomTemplate] = useState(null);
  const [excelData, setExcelData] = useState([]); // Excel-Daten bleiben hier erhalten

  const handleChange = (e) => {
    setDashboardData({ ...dashboardData, [e.target.name]: e.target.value });
  };

  // Dropdown-Optionen für Funktionsbezeichnungen
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
    <div>
      <Typography variant="h6" gutterBottom>
        Zusätzliche Eingaben
      </Typography>
      <Grid container spacing={2}>
        {/* Klassenleitung und Funktionsbezeichnung Klassenleitung */}
        <Grid item xs={12} sm={6}>
          <TextField
            label="Klassenleitung"
            name="klassenleitung"
            value={dashboardData.klassenleitung || ''}
            onChange={handleChange}
            fullWidth
          />
        </Grid>
        <Grid item xs={12} sm={6}>
          <TextField
            select
            label="Funktionsbezeichnung Klassenleitung"
            name="kl_titel"
            value={dashboardData.kl_titel || ''}
            onChange={handleChange}
            fullWidth
          >
            {funktionsOptions.map((option) => (
              <MenuItem key={option} value={option}>
                {option}
              </MenuItem>
            ))}
          </TextField>
        </Grid>
        {/* Schulleitung und Funktionsbezeichnung Schulleitung */}
        <Grid item xs={12} sm={6}>
          <TextField
            label="Schulleitung"
            name="schulleitung"
            value={dashboardData.schulleitung || ''}
            onChange={handleChange}
            fullWidth
          />
        </Grid>
        <Grid item xs={12} sm={6}>
          <TextField
            select
            label="Funktionsbezeichnung Schulleitung"
            name="sl_titel"
            value={dashboardData.sl_titel || ''}
            onChange={handleChange}
            fullWidth
          >
            {funktionsOptions.map((option) => (
              <MenuItem key={option} value={option}>
                {option}
              </MenuItem>
            ))}
          </TextField>
        </Grid>
        {/* Weitere Felder */}
        <Grid item xs={12} sm={6}>
          <TextField
            label="Schuljahr"
            name="schuljahr"
            value={dashboardData.schuljahr || ''}
            onChange={handleChange}
            fullWidth
          />
        </Grid>
        <Grid item xs={12} sm={6}>
          <TextField
            label="Zeugnisdatum"
            name="datum"
            type="date"
            value={dashboardData.datum || ''}
            onChange={handleChange}
            InputLabelProps={{ shrink: true }}
            fullWidth
          />
        </Grid>
        <Grid item xs={12} sm={6}>
          <TextField
            select
            label="Zeugnisart"
            name="zeugnisart"
            value={dashboardData.zeugnisart || ''}
            onChange={handleChange}
            fullWidth
          >
            <MenuItem value="Zwischenzeugnis">Zwischenzeugnis</MenuItem>
            <MenuItem value="Jahreszeugnis">Jahreszeugnis</MenuItem>
            <MenuItem value="Abschlusszeugnis">Abschlusszeugnis</MenuItem>
          </TextField>
        </Grid>
      </Grid>
      {/* Template-Upload */}
      <TemplateImport onTemplateLoaded={setCustomTemplate} />
      {/* Word-Generator */}
      <WordTemplateProcessor
        excelData={excelData}
        dashboardData={dashboardData}
        customTemplate={customTemplate}
      />
    </div>
  );
};

export default Dashboard;

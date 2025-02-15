import React, { useContext } from 'react';
import { TextField, MenuItem, Grid, Typography } from '@mui/material';
import { AppContext } from '../context/AppContext';

/*
  Dashboard-Komponente:
  - Bietet Eingabefelder für zusätzliche Daten:
    Klassenleitung, Schulleitung, Schuljahr, Datum und Zeugnisart.
  - Aktualisiert den globalen State entsprechend.
*/
const Dashboard = () => {
  const { dashboardData, setDashboardData } = useContext(AppContext);

  const handleChange = (e) => {
    setDashboardData({ ...dashboardData, [e.target.name]: e.target.value });
  };

  return (
    <div>
      <Typography variant="h6" gutterBottom>
        Zusätzliche Eingaben
      </Typography>
      <Grid container spacing={2}>
        <Grid item xs={12} sm={6}>
          <TextField
            label="Klassenleitung"
            name="klassenleitung"
            value={dashboardData.klassenleitung}
            onChange={handleChange}
            fullWidth
          />
        </Grid>
        <Grid item xs={12} sm={6}>
          <TextField
            label="Schulleitung"
            name="schulleitung"
            value={dashboardData.schulleitung}
            onChange={handleChange}
            fullWidth
          />
        </Grid>
        <Grid item xs={12} sm={6}>
          <TextField
            label="Schuljahr"
            name="schuljahr"
            value={dashboardData.schuljahr}
            onChange={handleChange}
            fullWidth
          />
        </Grid>
        <Grid item xs={12} sm={6}>
          <TextField
            label="Datum"
            name="datum"
            type="date"
            value={dashboardData.datum}
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
            value={dashboardData.zeugnisart}
            onChange={handleChange}
            fullWidth
          >
            <MenuItem value="Zwischenzeugnis">Zwischenzeugnis</MenuItem>
            <MenuItem value="Jahreszeugnis">Jahreszeugnis</MenuItem>
            <MenuItem value="Abschlusszeugnis">Abschlusszeugnis</MenuItem>
          </TextField>
        </Grid>
      </Grid>
    </div>
  );
};

export default Dashboard;

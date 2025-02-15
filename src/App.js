// src/App.js
import React, { useContext, useState } from 'react';
import { Container, Typography, Stepper, Step, StepLabel, Box, Button } from '@mui/material';
import ExcelUpload from './components/ExcelUpload';
import Dashboard from './components/Dashboard';
import WordTemplateProcessor from './components/WordTemplateProcessor';
import { AppContext } from './context/AppContext';

function App() {
  const { excelData, dashboardData } = useContext(AppContext);
  const steps = ['Excel Upload', 'Dashboard Eingaben', 'Word-Dokument Generierung'];
  const [activeStep, setActiveStep] = useState(0);

  const handleNext = () => {
    if (activeStep === 0 && excelData.length === 0) {
      alert("Bitte laden Sie zuerst die Excel-Datei hoch.");
      return;
    }
    setActiveStep(prev => prev + 1);
  };

  const handleBack = () => {
    setActiveStep(prev => prev - 1);
  };

  return (
    <Container maxWidth="md" sx={{ mt: 4, mb: 4 }}>
      <Typography variant="h4" align="center" gutterBottom>
        Zeugnis Generator
      </Typography>
      <Stepper activeStep={activeStep} alternativeLabel sx={{ mb: 4 }}>
        {steps.map((label, index) => (
          <Step key={index}>
            <StepLabel>{label}</StepLabel>
          </Step>
        ))}
      </Stepper>
      <Box>
        {activeStep === 0 && <ExcelUpload />}
        {activeStep === 1 && <Dashboard />}
        {activeStep === 2 && (
          <div>
            {excelData.length > 0 ? (
              <WordTemplateProcessor student={excelData[0]} dashboardData={dashboardData} />
            ) : (
              <Typography variant="body1">Keine Daten verfügbar.</Typography>
            )}
          </div>
        )}
      </Box>
      <Box display="flex" justifyContent="space-between" mt={4}>
        {activeStep > 0 && (
          <Button variant="contained" color="secondary" onClick={handleBack}>
            Zurück
          </Button>
        )}
        {activeStep < steps.length - 1 && (
          <Button variant="contained" onClick={handleNext}>
            Weiter
          </Button>
        )}
      </Box>
    </Container>
  );
}

export default App;

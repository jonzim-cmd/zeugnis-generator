import React, { useContext, useState } from 'react';
import { Button, Typography } from '@mui/material';
import { jsPDF } from 'jspdf';
import html2canvas from 'html2canvas';
import { AppContext } from '../context/AppContext';
import CertificateTemplate from './CertificateTemplate';

/*
  PDFGenerator-Komponente:
  - Rendert einen Button, um die PDF-Erstellung zu starten.
  - Für jeden Schüler aus den Excel-Daten wird eine (unsichtbare) 
    CertificateTemplate-Komponente gerendert, die mit html2canvas in ein Bild umgewandelt wird.
  - Die Bilder werden dann seitenweise in ein jsPDF-Dokument eingefügt.
*/
const PDFGenerator = () => {
  const { excelData, dashboardData } = useContext(AppContext);
  const [isGenerating, setIsGenerating] = useState(false);

  const generatePDF = async () => {
    setIsGenerating(true);
    const doc = new jsPDF('p', 'mm', 'a4');
    for (let i = 0; i < excelData.length; i++) {
      const certificateElement = document.getElementById(`certificate-${i}`);
      if (certificateElement) {
        // Erfassen der CertificateTemplate-Komponente als Canvas
        const canvas = await html2canvas(certificateElement, { scale: 2 });
        const imgData = canvas.toDataURL('image/png');
        const pdfWidth = doc.internal.pageSize.getWidth();
        const pdfHeight = doc.internal.pageSize.getHeight();
        if (i > 0) {
          doc.addPage();
        }
        // Das Bild wird so eingefügt, dass es die gesamte Seite ausfüllt.
        doc.addImage(imgData, 'PNG', 0, 0, pdfWidth, pdfHeight);
      }
    }
    doc.save('zeugnisse.pdf');
    setIsGenerating(false);
  };

  return (
    <div>
      <Typography variant="h6" gutterBottom>
        PDF Generierung
      </Typography>
      <Button variant="contained" color="primary" onClick={generatePDF} disabled={isGenerating}>
        {isGenerating ? 'Generiere PDF...' : 'PDF generieren'}
      </Button>
      {/* Unsichtbarer Container, in dem für jeden Schüler eine Zertifikat-Komponente gerendert wird */}
      <div style={{ position: 'absolute', top: '-10000px', left: '-10000px' }}>
        {excelData.map((student, index) => (
          <div key={index} id={`certificate-${index}`}>
            <CertificateTemplate student={student} dashboardData={dashboardData} />
          </div>
        ))}
      </div>
    </div>
  );
};

export default PDFGenerator;

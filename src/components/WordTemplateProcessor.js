import React, { useState, useContext } from 'react';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';
import { AppContext } from '../context/AppContext';

const WordTemplateProcessor = ({ student }) => {
  const [processing, setProcessing] = useState(false);
  const { dashboardData } = useContext(AppContext);

  const getTemplateFileName = () => {
    const art = dashboardData.zeugnisart || '';
    if (art === 'Zwischenzeugnis') {
      return `${process.env.PUBLIC_URL}/template_zwischen.docx`;
    } else if (art === 'Abschlusszeugnis') {
      return `${process.env.PUBLIC_URL}/template_abschluss.docx`;
    }
    return `${process.env.PUBLIC_URL}/template_jahr.docx`;
  };

  const generateDocx = async () => {
    setProcessing(true);
    try {
      const templateFile = getTemplateFileName();
      const response = await fetch(templateFile);
      if (!response.ok) throw new Error('Template nicht gefunden');
      const arrayBuffer = await response.arrayBuffer();

      const zip = new PizZip(arrayBuffer);
      const doc = new Docxtemplater();
      doc.loadZip(zip);

      // Zusammenführen von Excel-Daten (student) und Dashboard-Daten
      const data = {
        ...student,
        ...dashboardData,
        Zeugnisdatum: dashboardData.datum,
        Sl_Titel: dashboardData.sl_titel,
        Kl_Titel: dashboardData.kl_titel
      };

      doc.setData(data);
      doc.render();

      const out = doc.getZip().generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });

      saveAs(out, `zeugnis_${student.Nachname}.docx`);
    } catch (error) {
      console.error('Fehler:', error);
      alert('Fehler bei der Generierung!');
    } finally {
      setProcessing(false);
    }
  };

  return (
    <div>
      <button onClick={generateDocx} disabled={processing}>
        {processing ? "Generiere..." : "Word-Dokument erstellen"}
      </button>
    </div>
  );
};

export default WordTemplateProcessor;

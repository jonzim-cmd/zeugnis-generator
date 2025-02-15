import React, { useState, useContext } from 'react';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';
import { AppContext } from '../context/AppContext';

// Hilfsfunktion, um regex-sichere Strings zu erzeugen
const escapeRegExp = (string) => {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
};

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
      if (!response.ok) {
        throw new Error(`Template nicht gefunden: ${templateFile}`);
      }
      const arrayBuffer = await response.arrayBuffer();
      const zip = new PizZip(arrayBuffer);

      // Bearbeite ausschließlich das Hauptdokument (word/document.xml)
      const documentXmlPath = 'word/document.xml';
      if (!zip.file(documentXmlPath)) {
        throw new Error('Dokumentstruktur ungültig: word/document.xml nicht gefunden');
      }
      let xmlContent = zip.file(documentXmlPath).asText();

      // Mapping: Ersetze die exakten Platzhalter durch die entsprechenden Werte.
      const mapping = {
        'Placeholder_SJ': dashboardData.schuljahr || '',
        'Placeholder_KL': dashboardData.KL || '',
        'Placeholder_Vorname': student.Vorname || '',
        'Placeholder_Nachname': student.Nachname || '',
        'F1': student.F1 || '',
        'F2': student.F2 || '',
        'F3': student.F3 || '',
        'F4': student.F4 || '',
        'F5': student.F5 || '',
        'F6': student.F6 || '',
        'F7': student.F7 || '',
        'F8': student.F8 || '',
        'F9': student.F9 || '',
        'F1_N': student.F1_N || '',
        'F2_N': student.F2_N || '',
        'F3_N': student.F3_N || '',
        'F4_N': student.F4_N || '',
        'F5_N': student.F5_N || '',
        'F6_N': student.F6_N || '',
        'F7_N': student.F7_N || '',
        'F8_N': student.F8_N || '',
        'F9_N': student.F9_N || '',
        'BU': student.BU || '',
        'BU2': student.BU2 || '',
        'Zeugnisdatum': dashboardData.datum || '',
        'Placeholder_SL': dashboardData.schulleitung || '',
        'SL_Titel': dashboardData.sl_titel || '',
        'Placeholder_Klassenleitung': dashboardData.klassenleitung || '',
        'KL_Titel': dashboardData.kl_titel || ''
      };

      // Ersetze alle Platzhalter global in der XML
      Object.entries(mapping).forEach(([placeholder, value]) => {
        const regex = new RegExp(escapeRegExp(placeholder), 'g');
        xmlContent = xmlContent.replace(regex, value);
      });

      // Speichere den bearbeiteten XML-Content zurück in das ZIP
      zip.file(documentXmlPath, xmlContent);

      const out = zip.generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });

      saveAs(out, `zeugnis_${student.Nachname}.docx`);
    } catch (error) {
      console.error('Fehler beim Generieren der Word-Datei:', error);
      alert(`Fehler bei der Generierung: ${error.message || 'Unbekannter Fehler'}`);
    } finally {
      setProcessing(false);
    }
  };

  return (
    <div>
      <button 
        onClick={generateDocx} 
        disabled={processing}
        className="px-4 py-2 bg-blue-500 text-white rounded disabled:bg-gray-400"
      >
        {processing ? 'Generiere...' : 'Word-Dokument erstellen'}
      </button>
    </div>
  );
};

export default WordTemplateProcessor;

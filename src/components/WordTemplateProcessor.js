import React, { useState, useContext } from 'react';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';
import { AppContext } from '../context/AppContext';

// Helper: Escape regex-special characters
const escapeRegExp = (string) => {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
};

// Diese Funktion ersetzt ALLE <<...>>-Tags im Inhalt global durch {{...}}.
const sanitizeXmlContent = (content) => {
  return content.replace(/<<([^<>]+)>>/g, (match, p1) => {
    return `{{${p1.trim()}}}`;
  });
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
      const xmlContent = zip.file(documentXmlPath).asText();
      const sanitizedContent = sanitizeXmlContent(xmlContent);
      zip.file(documentXmlPath, sanitizedContent);

      // Konfiguriere und lade Docxtemplater
      const doc = new Docxtemplater();
      doc.loadZip(zip);
      
      // nullGetter sorgt dafür, dass fehlende Werte nicht zu Fehlern führen
      // und der parser entfernt ggf. überflüssige Klammern.
      doc.setOptions({
        nullGetter: () => '',
        parser: (tag) => {
          tag = tag.replace(/[<>]/g, '').trim();
          return {
            get: (scope) => scope[tag] || ''
          };
        }
      });

      const data = {
        ...student,
        ...dashboardData,
        Zeugnisdatum: dashboardData.datum
      };

      // Vorverarbeiten: Alle Werte als Strings
      const processedData = Object.entries(data).reduce((acc, [key, value]) => {
        acc[key] = value?.toString() || '';
        return acc;
      }, {});

      doc.setData(processedData);
      doc.render();

      const out = doc.getZip().generate({
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

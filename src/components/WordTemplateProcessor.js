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

  const sanitizeXmlContent = (content) => {
    // First, collect all unique placeholders
    const placeholders = new Set();
    const regex = /<<([^<>]+)>>/g;
    let match;
    
    while ((match = regex.exec(content)) !== null) {
      placeholders.add(match[0]);
    }

    // Then replace each placeholder exactly once
    let sanitizedContent = content;
    placeholders.forEach(placeholder => {
      const innerContent = placeholder.slice(2, -2).trim();
      const replacement = `{{${innerContent}}}`;
      // Replace only the first occurrence to avoid duplicate replacements
      sanitizedContent = sanitizedContent.replace(placeholder, replacement);
    });

    return sanitizedContent;
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

      // Process main document
      const documentXmlPath = 'word/document.xml';
      if (!zip.file(documentXmlPath)) {
        throw new Error('Dokumentstruktur ungültig: word/document.xml nicht gefunden');
      }

      const xmlContent = zip.file(documentXmlPath).asText();
      const sanitizedContent = sanitizeXmlContent(xmlContent);
      zip.file(documentXmlPath, sanitizedContent);

      // Configure and render document
      const doc = new Docxtemplater();
      doc.loadZip(zip);
      
      doc.setOptions({
        nullGetter: () => '',
        parser: (tag) => {
          // Remove any remaining << >> or extra spaces
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

      // Pre-process data to ensure all values are strings
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

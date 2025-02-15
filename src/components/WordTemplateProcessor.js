import React, { useState, useContext } from 'react';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';
import { AppContext } from '../context/AppContext';

const WordTemplateProcessor = ({ student }) => {
  const [processing, setProcessing] = useState(false);
  const { dashboardData } = useContext(AppContext);

  // Wähle das richtige DOCX-Template basierend auf der Zeugnisart.
  const getTemplateFileName = () => {
    const art = dashboardData.zeugnisart || '';
    if (art === 'Zwischenzeugnis') {
      return `${process.env.PUBLIC_URL}/template_zwischen.docx`;
    } else if (art === 'Abschlusszeugnis') {
      return `${process.env.PUBLIC_URL}/template_abschluss.docx`;
    }
    // Standard: Jahreszeugnis
    return `${process.env.PUBLIC_URL}/template_jahr.docx`;
  };

  const generateDocx = async () => {
    setProcessing(true);
    try {
      const templateFile = getTemplateFileName();
      const response = await fetch(templateFile);
      if (!response.ok) throw new Error('Template nicht gefunden');
      const arrayBuffer = await response.arrayBuffer();

      // Erstelle ein ZIP-Objekt aus dem ArrayBuffer.
      let zip = new PizZip(arrayBuffer);

      // Nur in der Hauptdokument-Datei (word/document.xml)
      const documentXmlPath = 'word/document.xml';
      if (zip.file(documentXmlPath)) {
        let xmlContent = zip.file(documentXmlPath).asText();
        // Ersetze alle Vorkommen von <<…>> durch {{…}}.
        // Dabei greift der Regex nur auf Zeichen zwischen << und >>, die nicht die Zeichen < oder > enthalten.
        xmlContent = xmlContent.replace(/<<([^<>]+)>>/g, (match, p1) => {
          return '{{' + p1.trim() + '}}';
        });
        zip.file(documentXmlPath, xmlContent);
      } else {
        console.warn('word/document.xml nicht gefunden');
      }

      // Lade Docxtemplater mit dem aktualisierten ZIP.
      const doc = new Docxtemplater();
      doc.loadZip(zip);

      // Konfiguriere den nullGetter, damit fehlende Platzhalter nicht Fehler verursachen.
      doc.setOptions({
        nullGetter: (part) => ''
      });

      // Füge Daten aus Excel (student) und Dashboard zusammen.
      const data = {
        ...student,
        ...dashboardData,
        Zeugnisdatum: dashboardData.datum
      };

      doc.setData(data);
      doc.render();

      const out = doc.getZip().generate({
        type: 'blob',
        mimeType:
          'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });
      saveAs(out, `zeugnis_${student.Nachname}.docx`);
    } catch (error) {
      console.error('Fehler beim Generieren der Word-Datei:', error);
      alert('Fehler bei der Generierung!');
    }
    setProcessing(false);
  };

  return (
    <div>
      <button onClick={generateDocx} disabled={processing}>
        {processing ? 'Generiere...' : 'Word-Dokument erstellen'}
      </button>
    </div>
  );
};

export default WordTemplateProcessor;

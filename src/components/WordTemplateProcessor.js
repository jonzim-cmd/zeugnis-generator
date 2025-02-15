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

      // Durchlaufe alle .xml-Dateien und ersetze alle <<...>> durch {{...}}
      const xmlFiles = Object.keys(zip.files).filter((fileName) =>
        fileName.endsWith('.xml')
      );
      xmlFiles.forEach((fileName) => {
        let xmlContent = zip.file(fileName).asText();
        xmlContent = xmlContent.replace(/<<([^>]+)>>/g, '{{$1}}');
        zip.file(fileName, xmlContent);
      });

      // Lade Docxtemplater mit dem aktualisierten ZIP.
      const doc = new Docxtemplater();
      doc.loadZip(zip);

      // Konfiguriere den nullGetter, damit fehlende Platzhalter durch einen leeren String ersetzt werden.
      doc.setOptions({
        nullGetter: (part) => {
          return '';
        }
      });

      // Füge Daten aus Excel (student) und Dashboard zusammen.
      const data = {
        ...student,
        ...dashboardData,
        Zeugnisdatum: dashboardData.datum
        // Falls du im Template z. B. auch {{SJ}} nutzt, kannst du hier ergänzen:
        // SJ: dashboardData.schuljahr
      };

      doc.setData(data);
      doc.render();

      // Erzeuge die finale DOCX als Blob und starte den Download.
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

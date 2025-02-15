// src/components/WordTemplateProcessor.js
import React, { useState } from 'react';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';

const WordTemplateProcessor = ({ student, dashboardData }) => {
  const [processing, setProcessing] = useState(false);

  // Wähle das richtige DOCX-Template basierend auf der Zeugnisart.
  // Mit process.env.PUBLIC_URL wird der korrekte Pfad zum public-Ordner erzeugt.
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
      if (!response.ok) {
        throw new Error(`Template file not found: ${templateFile}`);
      }
      const arrayBuffer = await response.arrayBuffer();
      const zip = new PizZip(arrayBuffer);
      
      // Lese den Inhalt der Hauptdokument-XML
      let documentXml = zip.file("word/document.xml").asText();

      // Daten für die Platzhalter – diese kommen aus Excel (student) und dem Dashboard
      const data = {
        Vorname: student.Vorname || '',
        Nachname: student.Nachname || '',
        SJ: dashboardData.schuljahr || '',
        Kl: student.Klasse || '',
        GDat: student.Geburtsdatum || '',
        GOrt: student.Geburtsort || ''
      };

      // Ersetze Platzhalter in beiden Formaten: {{...}} und <<...>>
      // Der reguläre Ausdruck sucht nach EITHER {{Name}} oder <<Name>>
      documentXml = documentXml
        .replace(/(?:{{|<<)Vorname(?:}}|>>)/g, data.Vorname)
        .replace(/(?:{{|<<)Nachname(?:}}|>>)/g, data.Nachname)
        .replace(/(?:{{|<<)SJ(?:}}|>>)/g, data.SJ)
        .replace(/(?:{{|<<)Kl(?:}}|>>)/g, data.Kl)
        .replace(/(?:{{|<<)GDat(?:}}|>>)/g, data.GDat)
        .replace(/(?:{{|<<)GOrt(?:}}|>>)/g, data.GOrt);

      // Schreibe die modifizierte XML zurück in das Zip-Archiv
      zip.file("word/document.xml", documentXml);

      // Erzeuge das finale DOCX als Blob
      const out = zip.generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      });
      
      // Starte den Download
      saveAs(out, "zeugnis.docx");
    } catch (error) {
      console.error("Fehler beim Generieren der Word-Datei:", error);
    }
    setProcessing(false);
  };

  return (
    <div>
      <button onClick={generateDocx} disabled={processing}>
        {processing ? "Erstelle Word-Datei..." : "Word-Datei generieren"}
      </button>
    </div>
  );
};

export default WordTemplateProcessor;

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

      // Lade den Inhalt der Dokument-XML
      let documentXml = zip.file("word/document.xml").asText();

      // Daten zum Ersetzen – Beachte:
      // - Vorname und Nachname kommen aus dem student-Objekt (Excel)
      // - SJ, KL, Schulleitung und Klassenleitung kommen aus dem dashboardData-Objekt
      const data = {
        Vorname: student.Vorname || '',
        Nachname: student.Nachname || '',
        SJ: dashboardData.schuljahr || '',           // Schuljahr aus dem Dashboard
        KL: dashboardData.KL || '',                     // Klasse aus dem Dashboard
        Schulleitung: dashboardData.schulleitung || '', // Schulleitung aus dem Dashboard
        Klassenleitung: dashboardData.klassenleitung || ''// Klassenleitung aus dem Dashboard
      };

      // Ersetze in der XML alle Platzhalter – unterstützt sowohl {{...}} als auch <<...>>
      documentXml = documentXml
        .replace(/(?:{{|<<)Vorname(?:}}|>>)/g, data.Vorname)
        .replace(/(?:{{|<<)Nachname(?:}}|>>)/g, data.Nachname)
        .replace(/(?:{{|<<)SJ(?:}}|>>)/g, data.SJ)
        .replace(/(?:{{|<<)KL(?:}}|>>)/g, data.KL)
        .replace(/(?:{{|<<)Schulleitung(?:}}|>>)/g, data.Schulleitung)
        .replace(/(?:{{|<<)Klassenleitung(?:}}|>>)/g, data.Klassenleitung);

      // Schreibe die modifizierte XML zurück in das ZIP-Archiv
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

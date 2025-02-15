// src/components/WordTemplateProcessor.js
import React, { useState } from 'react';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';

const WordTemplateProcessor = ({ student, dashboardData }) => {
  const [processing, setProcessing] = useState(false);

  // Wähle das richtige DOCX-Template basierend auf der Zeugnisart.
  // Mit process.env.PUBLIC_URL wird der richtige Pfad zum public-Ordner erzeugt.
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

      // Lese den Inhalt der document.xml, um zu prüfen, welches Platzhalterschema verwendet wird
      const documentXML = zip.file("word/document.xml")?.asText();
      let delimiters = { start: '{{', end: '}}' }; // Standardformat
      if (documentXML && documentXML.includes("<<") && documentXML.includes(">>")) {
        delimiters = { start: '<<', end: '>>' };
      }

      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
        delimiters: delimiters,
      });

      const data = {
        SJ: dashboardData.schuljahr || '',
        Kl: student.Klasse || '',
        Vorname: student.Vorname || '',
        Nachname: student.Nachname || '',
        GDat: student.Geburtsdatum || '',
        GOrt: student.Geburtsort || ''
        // Füge hier weitere Platzhalter hinzu, falls benötigt.
      };

      doc.setData(data);

      try {
        doc.render();
      } catch (error) {
        console.error("Fehler beim Rendern des Dokuments:", error);
        setProcessing(false);
        return;
      }

      const out = doc.getZip().generate({
        type: "blob",
        mimeType:
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });

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

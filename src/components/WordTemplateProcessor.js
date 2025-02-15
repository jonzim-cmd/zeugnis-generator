import React, { useState } from 'react';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';

const WordTemplateProcessor = ({ student, dashboardData }) => {
  const [processing, setProcessing] = useState(false);

  // Wählt das richtige DOCX-Template basierend auf der Zeugnisart
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

      // Lese den Inhalt der document.xml
      let documentXml = zip.file("word/document.xml").asText();

      // Zusammenführen von Excel-Daten (student) und Dashboard-Daten
      const data = {
        // Excel-Daten
        Zeugnisart: student.Zeugnisart || '',
        KL: student.KL || '',
        Vorname: student.Vorname || '',
        Nachname: student.Nachname || '',
        GDat: student.GDat || '',
        GOrt: student.GOrt || '',
        F1_ZZ: student.F1_ZZ || '',
        F2_ZZ: student.F2_ZZ || '',
        F3_ZZ: student.F3_ZZ || '',
        F4_ZZ: student.F4_ZZ || '',
        F5_ZZ: student.F5_ZZ || '',
        F6_ZZ: student.F6_ZZ || '',
        F7_ZZ: student.F7_ZZ || '',
        F8_ZZ: student.F8_ZZ || '',
        F9_ZZ: student.F9_ZZ || '',
        BU_ZZ: student.BU_ZZ || '',
        BU2_ZZ: student.BU2_ZZ || '',
        F1: student.F1 || '',
        F2: student.F2 || '',
        F3: student.F3 || '',
        F4: student.F4 || '',
        F5: student.F5 || '',
        F6: student.F6 || '',
        F7: student.F7 || '',
        F8: student.F8 || '',
        F9: student.F9 || '',
        // Dashboard-Daten
        Schulleitung: dashboardData.schulleitung || '',
        Klassenleitung: dashboardData.klassenleitung || '',
        Sl_Titel: dashboardData.sl_titel || '',
        Zeugnisdatum: dashboardData.datum || '',
        SJ: dashboardData.schuljahr || ''
      };

      // Ersetze alle Platzhalter (unterstützt sowohl {{...}} als auch <<...>>)
      for (const [key, value] of Object.entries(data)) {
        const regex = new RegExp(`(?:{{|<<)${key}(?:}}|>>)`, 'g');
        documentXml = documentXml.replace(regex, value);
      }

      // Überschreibe die document.xml im ZIP-Archiv
      zip.file("word/document.xml", documentXml);

      // Generiere die finale DOCX-Datei als Blob
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

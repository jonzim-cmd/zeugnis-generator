// src/components/WordTemplateProcessor.js
import React, { useState } from 'react';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';

const WordTemplateProcessor = ({ student, dashboardData }) => {
  const [processing, setProcessing] = useState(false);

  // Wähle das richtige Template basierend auf der Zeugnisart
  const getTemplateFileName = () => {
    const art = dashboardData.zeugnisart || '';
    if (art === 'Zwischenzeugnis') {
      return '/template_zwischen.docx';
    } else if (art === 'Abschlusszeugnis') {
      return '/template_abschluss.docx';
    }
    // Standard: Jahreszeugnis
    return '/template_jahr.docx';
  };

  const generateDocx = async () => {
    setProcessing(true);
    try {
      // Hole den Dateinamen für das gewählte Template
      const templateFile = getTemplateFileName();
      // Lade die DOCX-Vorlage aus dem public-Ordner
      const response = await fetch(templateFile);
      if (!response.ok) {
        throw new Error(`Template file not found: ${templateFile}`);
      }
      const arrayBuffer = await response.arrayBuffer();

      // Erstelle ein Zip-Objekt mit PizZip
      const zip = new PizZip(arrayBuffer);

      // Initialisiere docxtemplater mit dem Zip-Objekt
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });

      // Bereite die Daten vor, um die Platzhalter in der DOCX zu ersetzen.
      // Stelle sicher, dass in deiner DOCX-Datei dieselben Platzhalternamen (z. B. {{SJ}}, {{Kl}}, etc.) verwendet werden.
      const data = {
        SJ: dashboardData.schuljahr || '',
        Kl: student.Klasse || '',
        Vorname: student.Vorname || '',
        Nachname: student.Nachname || '',
        GDat: student.Geburtsdatum || '',
        GOrt: student.Geburtsort || ''
        // Füge weitere Platzhalter hinzu, falls benötigt.
      };

      // Setze die Daten in das Template
      doc.setData(data);

      try {
        doc.render();
      } catch (error) {
        console.error("Fehler beim Rendern des Dokuments:", error);
        setProcessing(false);
        return;
      }

      // Generiere das bearbeitete DOCX als Blob
      const out = doc.getZip().generate({
        type: "blob",
        mimeType: "application/vnd.openxmlformats-officedocument.wordprocessingml.document",
      });

      // Lasse den Nutzer die fertige Datei herunterladen
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

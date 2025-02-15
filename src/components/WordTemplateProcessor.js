// src/components/WordTemplateProcessor.js
import React, { useState } from 'react';
import PizZip from 'pizzip';
import Docxtemplater from 'docxtemplater';
import { saveAs } from 'file-saver';

const WordTemplateProcessor = ({ student, dashboardData }) => {
  const [processing, setProcessing] = useState(false);

  const generateDocx = async () => {
    setProcessing(true);
    try {
      // Lade die DOCX-Vorlage aus dem public-Ordner (Dateiname: template.docx)
      const response = await fetch('/template.docx');
      if (!response.ok) {
        throw new Error('Template file not found');
      }
      const arrayBuffer = await response.arrayBuffer();

      // Erstelle ein Zip-Objekt mit PizZip
      const zip = new PizZip(arrayBuffer);

      // Initialisiere docxtemplater mit dem Zip-Objekt
      const doc = new Docxtemplater(zip, {
        paragraphLoop: true,
        linebreaks: true,
      });

      // Bereite die Daten vor, um die Platzhalter in der DOCX zu ersetzen
      // Stelle sicher, dass die Platzhalternamen in deiner DOCX (z.B. {{SJ}}, {{Kl}}, etc.) mit diesen Schlüsseln übereinstimmen
      const data = {
        SJ: dashboardData.schuljahr || '',
        Kl: student.Klasse || '',
        Vorname: student.Vorname || '',
        Nachname: student.Nachname || '',
        GDat: student.Geburtsdatum || '',
        GOrt: student.Geburtsort || ''
        // Füge hier weitere Platzhalter hinzu, falls nötig
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

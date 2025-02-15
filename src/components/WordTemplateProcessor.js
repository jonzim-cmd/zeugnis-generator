import React, { useState } from 'react';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';

// Hilfsfunktion, um Regex-Sonderzeichen zu escapen.
// So können wir Platzhalternamen sicher in RegExps verwenden.
const escapeRegExp = (string) => {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
};

/*
  WordTemplateProcessor:
  - Bekommt jetzt "excelData" als Array (alle Zeilen aus der Excel)
  - Für jede Zeile wird ein eigenes Word-Dokument erzeugt.
  - "dashboardData" enthält die Angaben aus dem Dashboard.
*/
const WordTemplateProcessor = ({ excelData, dashboardData }) => {
  const [processing, setProcessing] = useState(false);

  // Wähle das richtige DOCX-Template anhand der Zeugnisart.
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
        throw new Error(`Template nicht gefunden: ${templateFile}`);
      }
      const arrayBuffer = await response.arrayBuffer();

      // Für jede Zeile in der Excel wird ein eigenes Dokument erzeugt.
      for (let i = 0; i < excelData.length; i++) {
        const student = excelData[i];

        // Kopiere das ArrayBuffer, damit wir pro Zeile eine frische Vorlage haben.
        const zip = new PizZip(arrayBuffer.slice(0));

        // Hauptdokument laden
        const documentXmlPath = 'word/document.xml';
        if (!zip.file(documentXmlPath)) {
          throw new Error('Dokumentstruktur ungültig: word/document.xml nicht gefunden');
        }
        let xmlContent = zip.file(documentXmlPath).asText();

        // Definiere das Mapping (Platzhalter -> Wert).
        // Achte darauf, dass die Platzhalternamen exakt mit denen im Word-Dokument übereinstimmen!
        const mapping = {
          // Aus dem Dashboard
          'Placeholder_SJ': dashboardData.schuljahr || '',
          'Placeholder_KL': dashboardData.KL || '',
          'Placeholder_SL': dashboardData.schulleitung || '',
          'SL_Titel': dashboardData.sl_titel || '',
          'Placeholder_Klassenleitung': dashboardData.klassenleitung || '',
          'KL_Titel': dashboardData.kl_titel || '',
          'Zeugnisdatum': dashboardData.datum || '',

          // Aus der Excel (student)
          'Placeholder_Vorname': student.Vorname || '',
          'Placeholder_Nachname': student.Nachname || '',
          'GDat': student.GDat || '',
          'GOrt': student.GOrt || '',
          'F1': student.F1 || '',
          'F2': student.F2 || '',
          'F3': student.F3 || '',
          'F4': student.F4 || '',
          'F5': student.F5 || '',
          'F6': student.F6 || '',
          'F7': student.F7 || '',
          'F8': student.F8 || '',
          'F9': student.F9 || '',
          'F1_N': student.F1_N || '',
          'F2_N': student.F2_N || '',
          'F3_N': student.F3_N || '',
          'F4_N': student.F4_N || '',
          'F5_N': student.F5_N || '',
          'F6_N': student.F6_N || '',
          'F7_N': student.F7_N || '',
          'F8_N': student.F8_N || '',
          'F9_N': student.F9_N || '',
          'BU': student.BU || '',
          'BU2': student.BU2 || ''
        };

        // Ersetze alle Platzhalter im XML.
        for (const [placeholder, value] of Object.entries(mapping)) {
          const regex = new RegExp(escapeRegExp(placeholder), 'g');
          xmlContent = xmlContent.replace(regex, value);
        }

        // Speichere den aktualisierten Inhalt zurück ins ZIP
        zip.file(documentXmlPath, xmlContent);

        // Erzeuge das finale DOCX als Blob
        const out = zip.generate({
          type: 'blob',
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });

        // Lade das Dokument herunter; z.B. "zeugnis_Meyer_1.docx"
        const dateiname = `zeugnis_${student.Nachname || 'unbekannt'}_${i + 1}.docx`;
        saveAs(out, dateiname);
      }
    } catch (error) {
      console.error('Fehler beim Generieren der Word-Datei:', error);
      alert(`Fehler bei der Generierung: ${error.message || 'Unbekannter Fehler'}`);
    } finally {
      setProcessing(false);
    }
  };

  return (
    <div>
      <button onClick={generateDocx} disabled={processing}>
        {processing ? 'Generiere...' : 'Word-Dokument(e) erstellen'}
      </button>
    </div>
  );
};

export default WordTemplateProcessor;

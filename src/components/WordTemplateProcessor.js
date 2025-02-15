import React, { useState } from 'react';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';

// Hilfsfunktion: Escapen von Regex-Sonderzeichen
const escapeRegExp = (string) => {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
};

// Konvertiert Excel-Serial in ein Datum im deutschen Format (TT.MM.JJJJ)
const formatExcelDate = (dateVal) => {
  if (typeof dateVal === 'number') {
    // Excel verwendet üblicherweise das 1900-Datumssystem
    const utcDays = Math.floor(dateVal - 25569);
    const utcValue = utcDays * 86400; // Sekunden
    const date = new Date(utcValue * 1000);
    const day = ('0' + date.getDate()).slice(-2);
    const month = ('0' + (date.getMonth() + 1)).slice(-2);
    const year = date.getFullYear();
    return `${day}.${month}.${year}`;
  }
  return dateVal;
};

// Konvertiert einen ISO-Datum-String (z. B. "2025-06-10") in das Format TT.MM.JJJJ
const formatIsoDate = (isoStr) => {
  if (!isoStr) return '';
  const date = new Date(isoStr);
  const day = ('0' + date.getDate()).slice(-2);
  const month = ('0' + (date.getMonth() + 1)).slice(-2);
  const year = date.getFullYear();
  return `${day}.${month}.${year}`;
};

const WordTemplateProcessor = ({ excelData, dashboardData }) => {
  const [processing, setProcessing] = useState(false);

  const getTemplateFileName = () => {
    const art = dashboardData.zeugnisart || '';
    if (art === 'Zwischenzeugnis') {
      return `${process.env.PUBLIC_URL}/template_zwischen.docx`;
    } else if (art === 'Abschlusszeugnis') {
      return `${process.env.PUBLIC_URL}/template_abschluss.docx`;
    }
    return `${process.env.PUBLIC_URL}/template_jahr.docx`;
  };

  // Reine Textersetzung: Für jede Zeile in der Excel wird ein eigenes DOCX erzeugt.
  const generateDocx = async () => {
    setProcessing(true);
    try {
      const templateFile = getTemplateFileName();
      const response = await fetch(templateFile);
      if (!response.ok) {
        throw new Error(`Template nicht gefunden: ${templateFile}`);
      }
      const arrayBuffer = await response.arrayBuffer();

      // Für jede Zeile in Excel (jeden Schüler) ein eigenes Dokument erzeugen:
      for (let i = 0; i < excelData.length; i++) {
        const student = excelData[i];
        // Kopiere den ArrayBuffer, damit jede Zeile eine frische Vorlage bekommt.
        const zip = new PizZip(arrayBuffer.slice(0));

        const documentXmlPath = 'word/document.xml';
        if (!zip.file(documentXmlPath)) {
          throw new Error('Dokumentstruktur ungültig: word/document.xml nicht gefunden');
        }
        let xmlContent = zip.file(documentXmlPath).asText();

        // Mapping: Die Keys müssen exakt den Platzhaltertexten in deinem Word-Dokument entsprechen.
        // Dashboard-Daten:
        const mapping = {
          'placeholdersj': dashboardData.schuljahr || '',
          'placeholdersl': dashboardData.schulleitung || '',
          'sltitel': dashboardData.sl_titel || '',
          'kltitel': dashboardData.kl_titel || '',
          'zeugnisdatum': formatIsoDate(dashboardData.datum) || '',
          // Dashboard-Wert für Klassenleitung:
          'placeholderkl': dashboardData.klassenleitung || '',

          // Excel-Daten:
          'placeholdervn': student.placeholdervn || '',
          'placeholdernm': student.placeholdernm || '',
          // Hier verwenden wir den Excel-Wert aus der Spalte "KL":
          'placeholderklasse': student.KL || '',
          'gdat': formatExcelDate(student.gdat) || '',
          'gort': student.gort || '',
          'f1': student.f1 || '',
          'f1n': student.f1n || '',
          'f2': student.f2 || '',
          'f2n': student.f2n || '',
          'f3': student.f3 || '',
          'f3n': student.f3n || '',
          'f4': student.f4 || '',
          'f4n': student.f4n || '',
          'f5': student.f5 || '',
          'f5n': student.f5n || '',
          'f6': student.f6 || '',
          'f6n': student.f6n || '',
          'f7': student.f7 || '',
          'f7n': student.f7n || '',
          'f8': student.f8 || '',
          'f8n': student.f8n || '',
          'f9': student.f9 || '',
          'f9n': student.f9n || '',
          'bueins': student.bueins || '',
          'buzwei': student.buzwei || ''
        };

        // Ersetze zuerst den Excel-Klassenplatzhalter ("placeholderklasse")
        const regexExcelKl = new RegExp(escapeRegExp('placeholderklasse'), 'g');
        xmlContent = xmlContent.replace(regexExcelKl, mapping['placeholderklasse']);
        // Ersetze danach den Dashboard-Wert für Klassenleitung ("placeholderkl")
        const regexDashKl = new RegExp(escapeRegExp('placeholderkl'), 'g');
        xmlContent = xmlContent.replace(regexDashKl, mapping['placeholderkl']);

        // Um Überschneidungen zu vermeiden (z. B. "f8" und "f8n"), sortieren wir die übrigen Keys nach Länge (längere zuerst)
        const keys = Object.keys(mapping)
          .filter(key => key !== 'placeholderklasse' && key !== 'placeholderkl')
          .sort((a, b) => b.length - a.length);
        keys.forEach((key) => {
          const regex = new RegExp(escapeRegExp(key), 'g');
          xmlContent = xmlContent.replace(regex, mapping[key]);
        });

        zip.file(documentXmlPath, xmlContent);

        const out = zip.generate({
          type: 'blob',
          mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
        });

        const filename = `zeugnis_${student.placeholdernm || 'unbekannt'}_${i + 1}.docx`;
        saveAs(out, filename);
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

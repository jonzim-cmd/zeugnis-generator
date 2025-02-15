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

  // Für jede Excelzeile wird nun nicht mehr ein eigenes Dokument erstellt, sondern
  // ein einziger Dokumenteninhalt mit mehrfach dupliziertem Schülerbereich erzeugt.
  const generateDocx = async () => {
    setProcessing(true);
    try {
      const templateFile = getTemplateFileName();
      const response = await fetch(templateFile);
      if (!response.ok) {
        throw new Error(`Template nicht gefunden: ${templateFile}`);
      }
      const arrayBuffer = await response.arrayBuffer();

      // Erstelle ein neues Zip-Objekt aus dem ArrayBuffer
      const zip = new PizZip(arrayBuffer);
      const documentXmlPath = 'word/document.xml';
      if (!zip.file(documentXmlPath)) {
        throw new Error('Dokumentstruktur ungültig: word/document.xml nicht gefunden');
      }
      let xmlContent = zip.file(documentXmlPath).asText();

      // Suche nach den STUDENT_SECTION Markern
      const startMarker = '<!-- STUDENT_SECTION_START -->';
      const endMarker = '<!-- STUDENT_SECTION_END -->';
      const startIndex = xmlContent.indexOf(startMarker);
      const endIndex = xmlContent.indexOf(endMarker);
      if (startIndex === -1 || endIndex === -1) {
        throw new Error('Student section markers nicht gefunden');
      }

      // Teile das Dokument in drei Bereiche: Vor dem Wiederholungsbereich, der wiederholbare Bereich, und danach
      const beforeSection = xmlContent.substring(0, startIndex + startMarker.length);
      const studentSectionTemplate = xmlContent.substring(startIndex + startMarker.length, endIndex);
      const afterSection = xmlContent.substring(endIndex);

      // Erzeuge den neuen Inhalt, indem du für jede Excelzeile den wiederholbaren Abschnitt kopierst
      let newStudentSections = "";
      for (let i = 0; i < excelData.length; i++) {
        const student = excelData[i];
        // Mapping: Dashboard- und Excel-Daten
        const mapping = {
          'placeholdersj': dashboardData.schuljahr || '',
          'placeholdersl': dashboardData.schulleitung || '',
          'sltitel': dashboardData.sl_titel || '',
          'kltitel': dashboardData.kl_titel || '',
          'zeugnisdatum': formatIsoDate(dashboardData.datum) || '',
          'placeholderkl': dashboardData.klassenleitung || '',
          'placeholdervn': student.placeholdervn || '',
          'placeholdernm': student.placeholdernm || '',
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

        // Ersetze zuerst "placeholderklasse" und "placeholderkl"
        let studentSection = studentSectionTemplate;
        studentSection = studentSection.replace(new RegExp(escapeRegExp('placeholderklasse'), 'g'), mapping['placeholderklasse']);
        studentSection = studentSection.replace(new RegExp(escapeRegExp('placeholderkl'), 'g'), mapping['placeholderkl']);

        // Ersetze anschließend alle übrigen Platzhalter – sortiert nach Länge, damit längere Namen zuerst ersetzt werden
        const keys = Object.keys(mapping).filter(key => key !== 'placeholderklasse' && key !== 'placeholderkl').sort((a, b) => b.length - a.length);
        keys.forEach((key) => {
          const regex = new RegExp(escapeRegExp(key), 'g');
          studentSection = studentSection.replace(regex, mapping[key]);
        });

        // Füge diesen Schülerabschnitt hinzu.
        // Falls in deinem Template bereits ein Seitenumbruch im Student-Block enthalten ist, wird dieser übernommen.
        newStudentSections += studentSection;
      }

      // Erzeuge den finalen XML-Inhalt, indem du den wiederholten Bereich einsetzt.
      const finalXml = beforeSection + newStudentSections + afterSection;
      zip.file(documentXmlPath, finalXml);

      const out = zip.generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });

      // Da nun alle Schüler in einem Dokument enthalten sind, vergeben wir einen generischen Namen.
      saveAs(out, 'zeugnisse_gesamt.docx');
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

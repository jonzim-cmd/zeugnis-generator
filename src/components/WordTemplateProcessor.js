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

  // Neue Logik: Ein einziges Dokument wird erstellt, in dem der Bereich zwischen den
  // Textmarkern (Word-Bookmarks) STUDENT_SECTION_START und STUDENT_SECTION_END
  // für jede Excelzeile dupliziert wird.
  const generateDocx = async () => {
    setProcessing(true);
    try {
      const templateFile = getTemplateFileName();
      const response = await fetch(templateFile);
      if (!response.ok) {
        throw new Error(`Template nicht gefunden: ${templateFile}`);
      }
      const arrayBuffer = await response.arrayBuffer();

      // Erstelle ein Zip-Objekt aus dem ArrayBuffer
      const zip = new PizZip(arrayBuffer);
      const documentXmlPath = 'word/document.xml';
      if (!zip.file(documentXmlPath)) {
        throw new Error('Dokumentstruktur ungültig: word/document.xml nicht gefunden');
      }
      let xmlContent = zip.file(documentXmlPath).asText();

      // Finde den BookmarkStart für STUDENT_SECTION_START
      const startMarkerRegex = /<w:bookmarkStart\s+[^>]*w:name="STUDENT_SECTION_START"[^>]*\/>/;
      const startMatch = xmlContent.match(startMarkerRegex);
      if (!startMatch) {
        throw new Error('Student section start marker nicht gefunden');
      }
      const startMarkerStr = startMatch[0];
      const startIndex = xmlContent.indexOf(startMarkerStr);

      // Extrahiere die w:id des Start-Bookmarks
      const idMatch = startMarkerStr.match(/w:id="(\d+)"/);
      if (!idMatch) {
        throw new Error('Kein w:id im STUDENT_SECTION_START Bookmark gefunden');
      }
      const bookmarkId = idMatch[1];

      // Finde den korrespondierenden BookmarkEnd, der den gleichen w:id hat
      const endMarkerRegex = new RegExp(`<w:bookmarkEnd\\s+[^>]*w:id="${bookmarkId}"[^>]*\\/?>`);
      const endMatch = xmlContent.match(endMarkerRegex);
      if (!endMatch) {
        throw new Error('Student section end marker nicht gefunden');
      }
      const endMarkerStr = endMatch[0];
      const endIndex = xmlContent.indexOf(endMarkerStr);

      // Teile das Dokument in drei Bereiche: vor dem wiederholbaren Bereich, der wiederholbare Bereich und danach
      const beforeSection = xmlContent.substring(0, startIndex + startMarkerStr.length);
      const studentSectionTemplate = xmlContent.substring(startIndex + startMarkerStr.length, endIndex);
      const afterSection = xmlContent.substring(endIndex);

      // Erzeuge den neuen, duplizierten Bereich
      let newStudentSections = "";
      for (let i = 0; i < excelData.length; i++) {
        const student = excelData[i];
        const mapping = {
          // Dashboard-Daten:
          'placeholdersj': dashboardData.schuljahr || '',
          'placeholdersl': dashboardData.schulleitung || '',
          'sltitel': dashboardData.sl_titel || '',
          'kltitel': dashboardData.kl_titel || '',
          'zeugnisdatum': formatIsoDate(dashboardData.datum) || '',
          'placeholderkl': dashboardData.klassenleitung || '',
          // Excel-Daten:
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

        // Ersetze die restlichen Platzhalter (längere zuerst, um Überschneidungen zu vermeiden)
        const keys = Object.keys(mapping).filter(key => key !== 'placeholderklasse' && key !== 'placeholderkl')
          .sort((a, b) => b.length - a.length);
        keys.forEach((key) => {
          const regex = new RegExp(escapeRegExp(key), 'g');
          studentSection = studentSection.replace(regex, mapping[key]);
        });

        newStudentSections += studentSection;
      }

      // Setze den finalen XML-Inhalt zusammen
      const finalXml = beforeSection + newStudentSections + afterSection;
      zip.file(documentXmlPath, finalXml);

      const out = zip.generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });

      // Ein einzelnes Dokument, das alle Schülerseiten enthält
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

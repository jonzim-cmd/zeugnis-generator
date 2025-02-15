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
    const utcDays = Math.floor(dateVal - 25569);
    const utcValue = utcDays * 86400;
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

  // Generiert ein einzelnes Word-Dokument, das alle Schülerabschnitte enthält.
  const generateDocx = async () => {
    setProcessing(true);
    try {
      const templateFile = getTemplateFileName();
      const response = await fetch(templateFile);
      if (!response.ok) {
        throw new Error(`Template nicht gefunden: ${templateFile}`);
      }
      const arrayBuffer = await response.arrayBuffer();

      // Zip öffnen und XML-Inhalt laden
      const zip = new PizZip(arrayBuffer);
      const documentXmlPath = 'word/document.xml';
      if (!zip.file(documentXmlPath)) {
        throw new Error('Dokumentstruktur ungültig: word/document.xml nicht gefunden');
      }
      let xmlContent = zip.file(documentXmlPath).asText();

      // --- 1. Bereich zwischen den Lesezeichen extrahieren ---
      // Wir gehen hier stringbasiert vor. Die Position des Endes des Start-Lesezeichens
      // und der Beginn des End-Lesezeichens bestimmen den zu duplizierenden Bereich.
      const startBookmarkRegex = /(<w:bookmarkStart[^>]*w:name="STUDENT_SECTION_START"[^>]*>)/g;
      const endBookmarkRegex = /(<w:bookmarkStart[^>]*w:name="STUDENT_SECTION_END"[^>]*>)/g;
      
      const startMatch = startBookmarkRegex.exec(xmlContent);
      if (!startMatch) {
        throw new Error('Lesezeichen "STUDENT_SECTION_START" nicht gefunden');
      }
      // Startindex = direkt nach dem Start-Bookmark
      const startIndex = startMatch.index + startMatch[0].length;
      
      const endMatch = endBookmarkRegex.exec(xmlContent);
      if (!endMatch) {
        throw new Error('Lesezeichen "STUDENT_SECTION_END" nicht gefunden');
      }
      const endIndex = endMatch.index;

      // Extrahiere den Template-Abschnitt, der für jeden Schüler dupliziert wird.
      const sectionTemplate = xmlContent.substring(startIndex, endIndex);

      // --- 2. Für jeden Schüler den Template-Bereich bearbeiten ---
      let allStudentSections = "";
      for (let i = 0; i < excelData.length; i++) {
        const student = excelData[i];
        // Mapping: Keys entsprechen exakt den Platzhaltertexten im Word-Dokument
        const mapping = {
          // Dashboard-Daten
          'placeholdersj': dashboardData.schuljahr || '',
          'placeholdersl': dashboardData.schulleitung || '',
          'sltitel': dashboardData.sl_titel || '',
          'kltitel': dashboardData.kl_titel || '',
          'zeugnisdatum': formatIsoDate(dashboardData.datum) || '',
          'placeholderkl': dashboardData.klassenleitung || '',
          // Excel-Daten
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

        // Beginne mit einer Kopie des Template-Abschnitts
        let studentSection = sectionTemplate;

        // Zuerst den Excel-Klassenplatzhalter ersetzen
        studentSection = studentSection.replace(
          new RegExp(escapeRegExp('placeholderklasse'), 'g'),
          mapping['placeholderklasse']
        );
        // Danach den Dashboard-Wert für Klassenleitung
        studentSection = studentSection.replace(
          new RegExp(escapeRegExp('placeholderkl'), 'g'),
          mapping['placeholderkl']
        );
        // Die übrigen Platzhalter ersetzen – längere Schlüssel zuerst, um Überschneidungen zu vermeiden
        Object.keys(mapping)
          .filter(key => key !== 'placeholderklasse' && key !== 'placeholderkl')
          .sort((a, b) => b.length - a.length)
          .forEach(key => {
            const regex = new RegExp(escapeRegExp(key), 'g');
            studentSection = studentSection.replace(regex, mapping[key]);
          });

        // Optional: Füge einen Seitenumbruch ein, damit jeder Schülerabschnitt auf einer neuen Seite beginnt
        if (i < excelData.length - 1) {
          studentSection += `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
        }

        allStudentSections += studentSection;
      }

      // --- 3. Zusammenbau des neuen Dokuments ---
      // Teile den ursprünglichen XML-Inhalt in drei Bereiche:
      // 1. Den Teil vor dem Ende des Start-Lesezeichens
      // 2. Den neuen, zusammengefügten Schülerabschnitt
      // 3. Den Teil ab dem Beginn des End-Lesezeichens
      const beforeSection = xmlContent.substring(0, startIndex);
      const afterSection = xmlContent.substring(endIndex);

      const newXmlContent = beforeSection + allStudentSections + afterSection;
      zip.file(documentXmlPath, newXmlContent);

      // Generiere den Blob und speichere die fertige Datei
      const out = zip.generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });
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
        {processing ? 'Generiere...' : 'Word-Dokument erstellen'}
      </button>
    </div>
  );
};

export default WordTemplateProcessor;

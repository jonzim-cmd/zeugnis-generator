import React, { useState } from 'react';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';

// Hilfsfunktion: Escapen von Regex-Sonderzeichen
const escapeRegExp = (string) => {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
};

// Escapen von XML-Sonderzeichen
const escapeXml = (unsafe) => {
  return unsafe.toString().replace(/[<>&'"]/g, (c) => {
    switch (c) {
      case '<': return '&lt;';
      case '>': return '&gt;';
      case '&': return '&amp;';
      default: return c;
    }
  });
};

// Konvertiert Excel-Serial in ein Datum im Format TT.MM.JJJJ
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

// Konvertiert einen ISO-Datum-String in das Format TT.MM.JJJJ
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

  const generateDocx = async () => {
    setProcessing(true);
    try {
      const templateFile = getTemplateFileName();
      const response = await fetch(templateFile);
      if (!response.ok) {
        throw new Error(`Template nicht gefunden: ${templateFile}`);
      }
      const arrayBuffer = await response.arrayBuffer();

      // Lade das DOCX (ZIP-Archiv)
      const zip = new PizZip(arrayBuffer);
      const documentXmlPath = 'word/document.xml';
      if (!zip.file(documentXmlPath)) {
        throw new Error('Dokumentstruktur ungültig: word/document.xml nicht gefunden');
      }
      let xmlContent = zip.file(documentXmlPath).asText();

      // --- Bookmark-Extraktion analog zur Originalversion ---
      // Suche das bookmarkStart-Element für STUDENT_SECTION_START
      const startBookmarkStartRegex = /<w:bookmarkStart[^>]*w:name="STUDENT_SECTION_START"[^>]*>/;
      const startBookmarkStartMatch = xmlContent.match(startBookmarkStartRegex);
      if (!startBookmarkStartMatch) {
        throw new Error('Lesezeichen "STUDENT_SECTION_START" nicht gefunden');
      }
      const startBookmarkStartIndex = xmlContent.indexOf(startBookmarkStartMatch[0]);
      // Das Ende des bookmarkStart-Elements
      const sectionStart = xmlContent.indexOf('>', startBookmarkStartIndex) + 1;

      // Suche das zugehörige bookmarkEnd für STUDENT_SECTION_START
      const startBookmarkEndRegex = /<w:bookmarkEnd[^>]*w:id\s*=\s*"(\d+)"[^>]*>/;
      const afterStartBookmarkStart = xmlContent.slice(sectionStart);
      const startBookmarkEndMatch = afterStartBookmarkStart.match(startBookmarkEndRegex);
      if (!startBookmarkEndMatch) {
        throw new Error('bookmarkEnd für "STUDENT_SECTION_START" nicht gefunden');
      }
      const startBookmarkEndIndex =
        sectionStart + afterStartBookmarkStart.indexOf(startBookmarkEndMatch[0]) + startBookmarkEndMatch[0].length;

      // Suche das bookmarkStart-Element für STUDENT_SECTION_END
      const endBookmarkStartRegex = /<w:bookmarkStart[^>]*w:name="STUDENT_SECTION_END"[^>]*>/;
      const afterStartBookmarkEnd = xmlContent.slice(startBookmarkEndIndex);
      const endBookmarkStartMatch = afterStartBookmarkEnd.match(endBookmarkStartRegex);
      if (!endBookmarkStartMatch) {
        throw new Error('Lesezeichen "STUDENT_SECTION_END" nicht gefunden');
      }
      const endBookmarkStartIndex = startBookmarkEndIndex + afterStartBookmarkEnd.indexOf(endBookmarkStartMatch[0]);

      // Extrahiere den Studentensektionsbereich (ohne Lesezeichen)
      let sectionTemplate = xmlContent.substring(startBookmarkEndIndex, endBookmarkStartIndex)
        .replace(/<w:bookmark(Start|End)[^>]*>/g, '');

      // --- Platzhalter-Ersetzung für jeden Schüler ---
      if (!Array.isArray(excelData)) {
        excelData = [excelData];
      }
      let allStudentSections = "";
      excelData.forEach((student, i) => {
        let studentSection = sectionTemplate;

        const mapping = {
          'placeholdersj': escapeXml(dashboardData.schuljahr || ''),
          'placeholdersl': escapeXml(dashboardData.schulleitung || ''),
          'sltitel': escapeXml(dashboardData.sl_titel || ''),
          'kltitel': escapeXml(dashboardData.kl_titel || ''),
          'zeugnisdatum': escapeXml(formatIsoDate(dashboardData.datum) || ''),
          'placeholderkl': escapeXml(dashboardData.klassenleitung || ''),
          // Excel-spezifisch:
          'placeholderklasse': escapeXml(student.KL || ''),
          'gdat': escapeXml(formatExcelDate(student.gdat) || '')
        };

        // Weitere Excel-Werte einfügen, ohne bereits definierte Schlüssel zu überschreiben
        Object.entries(student).forEach(([key, value]) => {
          if (['KL', 'gdat'].includes(key)) return;
          mapping[key] = escapeXml(value);
        });

        // Ersetze zuerst den Excel-Platzhalter "placeholderklasse"
        studentSection = studentSection.replace(
          new RegExp(escapeRegExp('placeholderklasse'), 'g'),
          mapping['placeholderklasse']
        );
        // Dann den Dashboard-Platzhalter "placeholderkl"
        studentSection = studentSection.replace(
          new RegExp(escapeRegExp('placeholderkl'), 'g'),
          mapping['placeholderkl']
        );
        // Ersetze alle übrigen Platzhalter exakt
        Object.keys(mapping)
          .filter(key => key !== 'placeholderklasse' && key !== 'placeholderkl')
          .sort((a, b) => b.length - a.length)
          .forEach(key => {
            const regex = new RegExp(escapeRegExp(key), 'g');
            studentSection = studentSection.replace(regex, mapping[key]);
          });

        // Füge zwischen den Schülern einen Seitenumbruch ein (außer beim letzten)
        if (i < excelData.length - 1) {
          studentSection += `
            <w:p>
              <w:r>
                <w:br w:type="page"/>
              </w:r>
            </w:p>
          `;
        }
        allStudentSections += studentSection;
      });

      // --- Zusammensetzen des neuen Dokuments ---
      const newXmlContent =
        xmlContent.substring(0, startBookmarkEndIndex) +
        allStudentSections +
        xmlContent.substring(endBookmarkStartIndex);

      // Keine Änderung am XML-Header oder den Namespaces – nur die exakten Platzhalter werden ersetzt.
      // Optional: Überprüfung der XML-Struktur (nur zur Diagnose)
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(newXmlContent, "text/xml");
      if (xmlDoc.getElementsByTagName("parsererror").length > 0) {
        console.error("XML ist invalide:", xmlDoc.getElementsByTagName("parsererror")[0].textContent);
        throw new Error("Generiertes XML ist fehlerhaft");
      }

      // Aktualisiere document.xml im ZIP
      zip.file(documentXmlPath, newXmlContent);
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

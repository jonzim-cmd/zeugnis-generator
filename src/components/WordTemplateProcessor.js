import React, { useState } from 'react';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';

// --- Hilfsfunktionen ---

// Escapen von Regex-Sonderzeichen
const escapeRegExp = (string) => {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
};

// Escapen von XML-Sonderzeichen
const escapeXml = (unsafe) => {
  return unsafe.toString().replace(/[<>&'"]/g, (c) => {
    switch(c) {
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

      // --- ZIP-Handling: Original-ZIP behalten, um fehlende Dateien später zu kopieren ---
      const originalZip = new PizZip(arrayBuffer);
      const zip = new PizZip(arrayBuffer);

      // Hole das document.xml
      const documentXmlPath = 'word/document.xml';
      if (!zip.file(documentXmlPath)) {
        throw new Error('Dokumentstruktur ungültig: word/document.xml nicht gefunden');
      }
      let xmlContent = zip.file(documentXmlPath).asText();

      // --- 1. Verbesserte Lesezeichen-Extraktion ---
      // Regex zur Erkennung der Bookmarks:
      const startBookmarkStartRegex = /<w:bookmarkStart[^>]*w:name="STUDENT_SECTION_START"[^>]*>/;
      // Prüfe, ob auch ein bookmarkEnd zum STUDENT_SECTION_START existiert:
      const startBookmarkEndRegex = /<w:bookmarkEnd[^>]*w:id\s*=\s*"(\d+)"[^>]*>/;
      const endBookmarkStartRegex = /<w:bookmarkStart[^>]*w:name="STUDENT_SECTION_END"[^>]*>/;

      // Finde das bookmarkStart-Tag für STUDENT_SECTION_START
      const startBookmarkStartMatch = xmlContent.match(startBookmarkStartRegex);
      if (!startBookmarkStartMatch) {
        throw new Error('Lesezeichen "STUDENT_SECTION_START" nicht gefunden');
      }
      const startBookmarkStartIndex = xmlContent.indexOf(startBookmarkStartMatch[0]);
      const afterStartBookmarkStartIndex = startBookmarkStartIndex + startBookmarkStartMatch[0].length;

      // Finde das zugehörige bookmarkEnd-Tag für STUDENT_SECTION_START
      const startBookmarkEndSlice = xmlContent.slice(afterStartBookmarkStartIndex);
      const startBookmarkEndMatch = startBookmarkEndSlice.match(startBookmarkEndRegex);
      if (!startBookmarkEndMatch) {
        throw new Error('bookmarkEnd für "STUDENT_SECTION_START" nicht gefunden');
      }
      const startBookmarkEndIndex =
        afterStartBookmarkStartIndex +
        startBookmarkEndSlice.indexOf(startBookmarkEndMatch[0]) +
        startBookmarkEndMatch[0].length;

      // Finde das bookmarkStart-Tag für STUDENT_SECTION_END (nach dem bookmarkEnd)
      const endBookmarkStartSlice = xmlContent.slice(startBookmarkEndIndex);
      const endBookmarkStartMatch = endBookmarkStartSlice.match(endBookmarkStartRegex);
      if (!endBookmarkStartMatch) {
        throw new Error('Lesezeichen "STUDENT_SECTION_END" nicht gefunden');
      }
      const endBookmarkStartIndex =
        startBookmarkEndIndex + endBookmarkStartSlice.indexOf(endBookmarkStartMatch[0]);

      // Extrahiere den Template-Bereich:
      // Von NACH dem bookmarkEnd von STUDENT_SECTION_START bis VOR dem bookmarkStart von STUDENT_SECTION_END
      let sectionTemplate = xmlContent.substring(startBookmarkEndIndex, endBookmarkStartIndex)
        // Entferne alle Bookmark-Tags, damit keine doppelten Bookmarks in den duplizierten Abschnitten entstehen
        .replace(/<w:bookmark(Start|End)[^>]*>/g, '');

      // --- 2. Platzhalter-Ersetzung für jeden Schüler ---
      let allStudentSections = "";
      for (let i = 0; i < excelData.length; i++) {
        const student = excelData[i];
        let studentSection = sectionTemplate;

        // Temporärer Namespace-Alias: Sicherstellen, dass XML-Namensräume konsistent bleiben
        studentSection = studentSection
          .replace(/<w:/g, '<ns:')
          .replace(/<\/w:/g, '</ns:')
          .replace(/ xmlns:w="[^"]*"/, '');

        // Erstelle ein Mapping (alle Werte werden mittels escapeXml geschützt)
        const mapping = {
          'placeholdersj': escapeXml(dashboardData.schuljahr || ''),
          'placeholdersl': escapeXml(dashboardData.schulleitung || ''),
          'sltitel': escapeXml(dashboardData.sl_titel || ''),
          'kltitel': escapeXml(dashboardData.kl_titel || ''),
          'zeugnisdatum': escapeXml(formatIsoDate(dashboardData.datum) || ''),
          'placeholderkl': escapeXml(dashboardData.klassenleitung || '')
        };

        // Für alle Felder aus den Excel-Daten – hier iterieren wir über alle Schlüssel des Schülereintrags
        Object.entries(student).forEach(([key, val]) => {
          mapping[key] = escapeXml(val);
        });

        // Ersetze zuerst den Excel-Klassenplatzhalter
        studentSection = studentSection.replace(
          new RegExp(escapeRegExp('placeholderklasse'), 'g'),
          mapping['placeholderklasse'] || ''
        );
        // Ersetze dann den Dashboard-Platzhalter für Klassenleitung
        studentSection = studentSection.replace(
          new RegExp(escapeRegExp('placeholderkl'), 'g'),
          mapping['placeholderkl'] || ''
        );
        // Ersetze alle übrigen Platzhalter (längere Schlüssel zuerst, um Überschneidungen zu vermeiden)
        Object.keys(mapping)
          .filter(key => key !== 'placeholderklasse' && key !== 'placeholderkl')
          .sort((a, b) => b.length - a.length)
          .forEach(key => {
            const regex = new RegExp(escapeRegExp(key), 'g');
            studentSection = studentSection.replace(regex, mapping[key]);
          });

        // Namespace-Alias zurückübersetzen
        studentSection = studentSection
          .replace(/<ns:/g, '<w:')
          .replace(/<\/ns:/g, '</w:');

        // Füge einen robusten Seitenumbruch ein (vollständiger Paragraph)
        const pageBreak = `
          <w:p>
            <w:r>
              <w:br w:type="page" w:subType="page"/>
            </w:r>
            <w:pPr>
              <w:spacing w:after="0" w:line="240" w:lineRule="auto"/>
            </w:pPr>
          </w:p>
        `;
        if (i < excelData.length - 1) {
          studentSection += pageBreak;
        }
        allStudentSections += studentSection;
      }

      // --- 3. Zusammenbau des neuen Dokuments ---
      const beforeSection = xmlContent.substring(0, startBookmarkEndIndex);
      const afterSection = xmlContent.substring(endBookmarkStartIndex);
      let newXmlContent = beforeSection + allStudentSections + afterSection;

      // --- 4. Namespace-Konsistenz sicherstellen ---
      // Füge den Namespace beim Öffnungs-Tag des Dokuments ein, falls nicht vorhanden
      newXmlContent = newXmlContent.replace(
        /<w:document/,
        '<w:document xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"'
      );

      // --- 5. XML-Validierung ---
      // Mit DOMParser prüfen wir, ob das generierte XML wohlgeformt ist.
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(newXmlContent, "text/xml");
      if (xmlDoc.getElementsByTagName("parsererror").length > 0) {
        console.error("XML ist invalide:", xmlDoc.getElementsByTagName("parsererror")[0].textContent);
        throw new Error("Generiertes XML ist fehlerhaft");
      }

      // --- 6. ZIP-Struktur vervollständigen ---
      // Stelle sicher, dass alle originalen Dateien im ZIP erhalten bleiben.
      originalZip.forEach((relativePath, file) => {
        if (!zip.file(relativePath)) {
          zip.file(relativePath, file.asUint8Array());
        }
      });

      // Aktualisiere die document.xml im ZIP
      zip.file(documentXmlPath, newXmlContent);

      // --- 7. Ergebnis generieren und speichern ---
      const out = zip.generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });
      saveAs(out, 'zeugnisse_gesamt.docx');

      // Optional: Zum Debuggen – Ausgabe eines XML-Ausschnitts
      // console.log(newXmlContent.substring(0, 1000));

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

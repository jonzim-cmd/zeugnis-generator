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
      // Suche die Positionen des STUDENT_SECTION_START und STUDENT_SECTION_END
      const startBookmarkRegex = /(<w:bookmarkStart[^>]*w:name="STUDENT_SECTION_START"[^>]*>)/g;
      const endBookmarkRegex = /(<w:bookmarkStart[^>]*w:name="STUDENT_SECTION_END"[^>]*>)/g;
      
      const startMatch = startBookmarkRegex.exec(xmlContent);
      if (!startMatch) {
        throw new Error('Lesezeichen "STUDENT_SECTION_START" nicht gefunden');
      }
      const startIndex = startMatch.index + startMatch[0].length;
      
      const endMatch = endBookmarkRegex.exec(xmlContent);
      if (!endMatch) {
        throw new Error('Lesezeichen "STUDENT_SECTION_END" nicht gefunden');
      }
      const endIndex = endMatch.index;
      
      // Extrahiere den Template-Abschnitt und entferne alle Bookmark-Tags
      let sectionTemplate = xmlContent.substring(startIndex, endIndex)
        .replace(/<w:bookmark(Start|End)[^>]*>/g, '');

      // --- 2. Für jeden Schüler den Template-Bereich bearbeiten ---
      let allStudentSections = "";
      for (let i = 0; i < excelData.length; i++) {
        const student = excelData[i];
        let studentSection = sectionTemplate;

        // Namespace-Alias: Vor der Platzhalter-Ersetzung
        studentSection = studentSection
          .replace(/<w:/g, '<ns:')
          .replace(/<\/w:/g, '</ns:')
          .replace(/ xmlns:w="[^"]*"/, '');

        // Mapping: Keys entsprechen den Platzhaltertexten im Template
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

        // Zuerst den Excel-Klassenplatzhalter ersetzen
        studentSection = studentSection.replace(
          new RegExp(escapeRegExp('placeholderklasse'), 'g'),
          mapping['placeholderklasse']
        );
        // Dann den Dashboard-Platzhalter für Klassenleitung
        studentSection = studentSection.replace(
          new RegExp(escapeRegExp('placeholderkl'), 'g'),
          mapping['placeholderkl']
        );
        // Die übrigen Platzhalter ersetzen (längere Schlüssel zuerst, um Überschneidungen zu vermeiden)
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

        // Robuster Seitenumbruch: kompletter Paragraph
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
      const beforeSection = xmlContent.substring(0, startIndex);
      const afterSection = xmlContent.substring(endIndex);
      let newXmlContent = beforeSection + allStudentSections + afterSection;

      // Falls im Development-Modus: XML validieren (sofern xmllint vorhanden)
      if (process.env.NODE_ENV === 'development') {
        try {
          const { validateXML } = await import('xmllint');
          const wordSchema = ''; // Hier muss das Office Open XML-Schema definiert bzw. geladen werden.
          const validation = validateXML({ xml: newXmlContent, schema: wordSchema });
          if (validation.errors && validation.errors.length > 0) {
            console.error('XML Validation Errors:', validation.errors);
          }
        } catch (err) {
          console.warn('XML Validierung konnte nicht durchgeführt werden:', err);
        }
      }

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

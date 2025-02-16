import React, { useState } from 'react';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';

// Hilfsfunktion: Regex-Escaping
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

  // Bestimme anhand der Zeugnisart das Template
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

      // Lade das ZIP-Archiv (DOCX)
      const zip = new PizZip(arrayBuffer);
      const documentXmlPath = 'word/document.xml';
      if (!zip.file(documentXmlPath)) {
        throw new Error('Dokumentstruktur ungültig: word/document.xml nicht gefunden');
      }
      let xmlContent = zip.file(documentXmlPath).asText();

      // Hier nehmen wir an, dass im Template ein Bereich definiert ist, 
      // in dem der Studentensektion steht, z. B. zwischen den Lesezeichen
      const startBookmark = '<w:bookmarkStart w:name="STUDENT_SECTION_START"';
      const endBookmark = '<w:bookmarkStart w:name="STUDENT_SECTION_END"';
      const startIdx = xmlContent.indexOf(startBookmark);
      const endIdx = xmlContent.indexOf(endBookmark);

      if (startIdx === -1 || endIdx === -1 || endIdx <= startIdx) {
        throw new Error('Die benötigten Lesezeichen für den Studentensektionsbereich wurden nicht gefunden.');
      }

      // Erhalte den zu ersetzenden Abschnitt, ohne XML-Manipulation des Headers oder der Namespaces
      const sectionStartEnd = xmlContent.indexOf('>', startIdx) + 1;
      const sectionTemplate = xmlContent.substring(sectionStartEnd, endIdx);

      // Erstelle den zusammengesetzten Inhalt für alle Schüler
      let allStudentSections = "";
      if (!Array.isArray(excelData)) {
        excelData = [excelData];
      }
      excelData.forEach((student, i) => {
        let studentSection = sectionTemplate;
        // Mapping aus Dashboard-Daten und Excel-Daten
        const mapping = {
          'placeholdersj': escapeXml(dashboardData.schuljahr || ''),
          'placeholdersl': escapeXml(dashboardData.schulleitung || ''),
          'sltitel': escapeXml(dashboardData.sl_titel || ''),
          'kltitel': escapeXml(dashboardData.kl_titel || ''),
          'zeugnisdatum': escapeXml(formatIsoDate(dashboardData.datum) || ''),
          // Dashboard-Wert für Klassenleitung
          'placeholderkl': escapeXml(dashboardData.klassenleitung || ''),
          // Excel-spezifisch:
          // Ersetze "placeholderklasse" mit dem Wert aus Spalte "KL"
          'placeholderklasse': escapeXml(student.KL || ''),
          // Datum "gdat" formatiert
          'gdat': escapeXml(formatExcelDate(student.gdat) || '')
        };

        // Ergänze weitere Excel-Felder, ohne Überschreibung der obigen Keys
        Object.entries(student).forEach(([key, value]) => {
          if (['KL', 'gdat'].includes(key)) return;
          let safeValue = escapeXml(value);
          mapping[key] = safeValue;
        });

        // Ersetze zuerst den Excel-Platzhalter "placeholderklasse"
        studentSection = studentSection.replace(
          new RegExp(escapeRegExp('placeholderklasse'), 'g'),
          mapping['placeholderklasse']
        );
        // Ersetze danach den Dashboard-Platzhalter "placeholderkl"
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

      // Setze den neuen Inhalt wieder in das Dokument ein
      const newXmlContent =
        xmlContent.substring(0, sectionStartEnd) +
        allStudentSections +
        xmlContent.substring(endIdx);

      // (Optional) Überprüfung der XML-Struktur – hier ohne Eingriff in Header/Namespaces
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(newXmlContent, "text/xml");
      if (xmlDoc.getElementsByTagName("parsererror").length > 0) {
        console.error("XML ist invalide:", xmlDoc.getElementsByTagName("parsererror")[0].textContent);
        throw new Error("Generiertes XML ist fehlerhaft");
      }

      // Überschreibe document.xml im ZIP und speichere das fertige Dokument
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

import React, { useState } from 'react';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';

// Escapen von Regex-Sonderzeichen
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

const WordTemplateProcessor = ({ excelData, dashboardData, customTemplate }) => {
  const [processing, setProcessing] = useState(false);

  // Angepasst: Bestimme anhand der Zeugnisart das Template aus dem privaten Submodul-Ordner.
  const getTemplateFileName = () => {
    const art = dashboardData.zeugnisart || '';
    if (art === 'Zwischenzeugnis') {
      return `${process.env.PUBLIC_URL}/private-templates/template_zwischen.docx`;
    } else if (art === 'Abschlusszeugnis') {
      return `${process.env.PUBLIC_URL}/private-templates/template_abschluss.docx`;
    }
    return `${process.env.PUBLIC_URL}/private-templates/template_jahr.docx`;
  };

  const generateDocx = async () => {
    setProcessing(true);
    try {
      let arrayBuffer;
      if (customTemplate) {
        // Verwende das hochgeladene Template
        arrayBuffer = customTemplate;
      } else {
        // Statt des direkten Aufrufs des privaten Templates wird jetzt der API-Endpoint aufgerufen:
        const response = await fetch('/api/get-template');
        if (!response.ok) {
          throw new Error(`Template nicht gefunden`);
        }
        arrayBuffer = await response.arrayBuffer();
      }

      // Lade das DOCX (ZIP-Archiv)
      const zip = new PizZip(arrayBuffer);
      const documentXmlPath = 'word/document.xml';
      if (!zip.file(documentXmlPath)) {
        throw new Error('Dokumentstruktur ungültig: word/document.xml nicht gefunden');
      }
      const xmlContent = zip.file(documentXmlPath).asText();

      // --- Extrahiere den <w:body>-Bereich als Vorlage ---
      const bodyStartTag = '<w:body>';
      const bodyEndTag = '</w:body>';
      const bodyStartIndex = xmlContent.indexOf(bodyStartTag);
      const bodyEndIndex = xmlContent.indexOf(bodyEndTag);
      if (bodyStartIndex === -1 || bodyEndIndex === -1) {
        throw new Error('Die benötigten <w:body>-Tags wurden nicht gefunden.');
      }
      // Behalte den Originalheader und alles vor <w:body>
      const preBody = xmlContent.substring(0, bodyStartIndex + bodyStartTag.length);
      // Und alles nach </w:body> (inklusive </w:document> etc.)
      const postBody = xmlContent.substring(bodyEndIndex);
      // Extrahiere den Body-Inhalt (das Zeugnis-Template)
      let studentTemplate = xmlContent.substring(bodyStartIndex + bodyStartTag.length, bodyEndIndex).trim();

      // Entferne gegebenenfalls den <w:sectPr>-Block am Ende, damit er nicht mehrfach eingefügt wird.
      let sectPr = '';
      const sectPrIndex = studentTemplate.lastIndexOf('<w:sectPr');
      if (sectPrIndex !== -1) {
        sectPr = studentTemplate.substring(sectPrIndex);
        studentTemplate = studentTemplate.substring(0, sectPrIndex);
      }

      // --- Erzeuge für jeden Schüler einen Abschnitt ---
      let allStudentSections = "";
      if (!Array.isArray(excelData)) {
        excelData = [excelData];
      }
      excelData.forEach((student, i) => {
        let studentSection = studentTemplate;
        // Mapping aus Dashboard-Daten und Excel-Daten
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

        // Ergänze weitere Excel-Werte, ohne die oben definierten Keys zu überschreiben
        Object.entries(student).forEach(([key, value]) => {
          if (['KL', 'gdat'].includes(key)) return;
          mapping[key] = escapeXml(value);
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
        // Ersetze alle übrigen Platzhalter exakt – längere zuerst
        Object.keys(mapping)
          .filter(key => key !== 'placeholderklasse' && key !== 'placeholderkl')
          .sort((a, b) => b.length - a.length)
          .forEach(key => {
            const regex = new RegExp(escapeRegExp(key), 'g');
            studentSection = studentSection.replace(regex, mapping[key]);
          });

        // **Abschnittswechsel einfügen statt Seitenumbruch:**
        const sectionBreak = `<w:p><w:pPr>${sectPr}</w:pPr></w:p>`;
        const paragraphRegex = /(<w:p\b[^>]*>[\s\S]*?<w:t[^>]*>Studen End<\/w:t>[\s\S]*?)(<\/w:p>)/g;
        if (paragraphRegex.test(studentSection) && i < excelData.length - 1) {
          studentSection = studentSection.replace(
            paragraphRegex,
            `$1$2${sectionBreak}`
          );
        } else if (i < excelData.length - 1) {
          // Fallback: Hänge den Abschnittswechsel als eigenen Absatz an.
          studentSection += sectionBreak;
        }
        
        allStudentSections += studentSection;
      });

      // Hänge zum Ende der zusammengesetzten Schülerabschnitte einmalig den <w:sectPr>-Block an,
      // damit die Sektionseinstellungen erhalten bleiben.
      const newBodyContent = allStudentSections + sectPr;

      // Füge den neuen Body wieder in das komplette Dokument ein
      const newXmlContent = preBody + newBodyContent + postBody;

      // (Optional) Überprüfung der XML-Struktur – nur zur Diagnose
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
    <div style={{ textAlign: 'center' }}>
      <button onClick={generateDocx} disabled={processing}>
        {processing ? 'Generiere...' : 'Word-Dokument erstellen'}
      </button>
    </div>
  );
};

export default WordTemplateProcessor;

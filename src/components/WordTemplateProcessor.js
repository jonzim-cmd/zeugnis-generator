// src/components/WordTemplateProcessor.js
import React, { useState } from 'react';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';

// Escapen von Regex-Sonderzeichen
const escapeRegExp = (string) => string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');

// Escapen von XML-Sonderzeichen
const escapeXml = (unsafe) =>
  unsafe.toString().replace(/[<>&'"]/g, (c) => {
    switch (c) {
      case '<':
        return '&lt;';
      case '>':
        return '&gt;';
      case '&':
        return '&amp;';
      default:
        return c;
    }
  });

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

  const generateDocx = async () => {
    setProcessing(true);
    try {
      let arrayBuffer;
      if (customTemplate) {
        arrayBuffer = customTemplate;
      } else {
        // Lese den Wert aus dashboardData.zeugnisart, Standard: "Jahreszeugnis"
        const zeugnisart = dashboardData.zeugnisart || 'Jahreszeugnis';
        const response = await fetch(
          `/api/get-template?zeugnisart=${encodeURIComponent(zeugnisart)}`
        );
        if (!response.ok) {
          throw new Error(`Template nicht gefunden: ${zeugnisart}`);
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
      const preBody = xmlContent.substring(0, bodyStartIndex + bodyStartTag.length);
      const postBody = xmlContent.substring(bodyEndIndex);
      let studentTemplate = xmlContent.substring(bodyStartIndex + bodyStartTag.length, bodyEndIndex).trim();

      // Entferne ggf. den <w:sectPr>-Block am Ende, damit er nicht mehrfach eingefügt wird.
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
        const mapping = {
          'placeholdersj': escapeXml(dashboardData.schuljahr || ''),
          'placeholdersl': escapeXml(dashboardData.schulleitung || ''),
          'sltitel': escapeXml(dashboardData.sl_titel || ''),
          'kltitel': escapeXml(dashboardData.kl_titel || ''),
          'zeugnisdatum': escapeXml(formatIsoDate(dashboardData.datum) || ''),
          'placeholderkl': escapeXml(dashboardData.klassenleitung || ''),
          'placeholderklasse': escapeXml(student.KL || ''),
          'gdat': escapeXml(formatExcelDate(student.gdat) || '')
        };

        Object.entries(student).forEach(([key, value]) => {
          if (['KL', 'gdat'].includes(key)) return;
          mapping[key] = escapeXml(value);
        });

        studentSection = studentSection.replace(
          new RegExp(escapeRegExp('placeholderklasse'), 'g'),
          mapping['placeholderklasse']
        );
        studentSection = studentSection.replace(
          new RegExp(escapeRegExp('placeholderkl'), 'g'),
          mapping['placeholderkl']
        );
        Object.keys(mapping)
          .filter(key => key !== 'placeholderklasse' && key !== 'placeholderkl')
          .sort((a, b) => b.length - a.length)
          .forEach(key => {
            const regex = new RegExp(escapeRegExp(key), 'g');
            studentSection = studentSection.replace(regex, mapping[key]);
          });

        const sectionBreak = `<w:p><w:pPr>${sectPr}</w:pPr></w:p>`;
        const paragraphRegex = /(<w:p\b[^>]*>[\s\S]*?<w:t[^>]*>Studen End<\/w:t>[\s\S]*?)(<\/w:p>)/g;
        if (paragraphRegex.test(studentSection) && i < excelData.length - 1) {
          studentSection = studentSection.replace(
            paragraphRegex,
            `$1$2${sectionBreak}`
          );
        } else if (i < excelData.length - 1) {
          studentSection += sectionBreak;
        }
        
        allStudentSections += studentSection;
      });

      const newBodyContent = allStudentSections + sectPr;
      const newXmlContent = preBody + newBodyContent + postBody;

      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(newXmlContent, "text/xml");
      if (xmlDoc.getElementsByTagName("parsererror").length > 0) {
        console.error("XML ist invalide:", xmlDoc.getElementsByTagName("parsererror")[0].textContent);
        throw new Error("Generiertes XML ist fehlerhaft");
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
    <div style={{ textAlign: 'center' }}>
      <button onClick={generateDocx} disabled={processing}>
        {processing ? 'Generiere...' : 'Word-Dokument erstellen'}
      </button>
    </div>
  );
};

export default WordTemplateProcessor;

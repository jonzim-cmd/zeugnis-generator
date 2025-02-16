import React, { useState } from 'react';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';

const escapeRegExp = (string) => {
  return string.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
};

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

      const originalZip = new PizZip(arrayBuffer);
      const zip = new PizZip(arrayBuffer);

      const documentXmlPath = 'word/document.xml';
      if (!zip.file(documentXmlPath)) {
        throw new Error('Dokumentstruktur ungültig: word/document.xml nicht gefunden');
      }
      let xmlContent = zip.file(documentXmlPath).asText();

      const startBookmarkStartRegex = /<w:bookmarkStart[^>]*w:name="STUDENT_SECTION_START"[^>]*>/;
      const startBookmarkEndRegex = /<w:bookmarkEnd[^>]*w:id\s*=\s*"(\d+)"[^>]*>/;
      const endBookmarkStartRegex = /<w:bookmarkStart[^>]*w:name="STUDENT_SECTION_END"[^>]*>/;

      const startBookmarkStartMatch = xmlContent.match(startBookmarkStartRegex);
      if (!startBookmarkStartMatch) {
        throw new Error('Lesezeichen "STUDENT_SECTION_START" nicht gefunden');
      }
      const startBookmarkStartIndex = xmlContent.indexOf(startBookmarkStartMatch[0]);
      const afterStartBookmarkStartIndex = startBookmarkStartIndex + startBookmarkStartMatch[0].length;

      const startBookmarkEndSlice = xmlContent.slice(afterStartBookmarkStartIndex);
      const startBookmarkEndMatch = startBookmarkEndSlice.match(startBookmarkEndRegex);
      if (!startBookmarkEndMatch) {
        throw new Error('bookmarkEnd für "STUDENT_SECTION_START" nicht gefunden');
      }
      const startBookmarkEndIndex =
        afterStartBookmarkStartIndex +
        startBookmarkEndSlice.indexOf(startBookmarkEndMatch[0]) +
        startBookmarkEndMatch[0].length;

      const endBookmarkStartSlice = xmlContent.slice(startBookmarkEndIndex);
      const endBookmarkStartMatch = endBookmarkStartSlice.match(endBookmarkStartRegex);
      if (!endBookmarkStartMatch) {
        throw new Error('Lesezeichen "STUDENT_SECTION_END" nicht gefunden');
      }
      const endBookmarkStartIndex =
        startBookmarkEndIndex + endBookmarkStartSlice.indexOf(endBookmarkStartMatch[0]);

      let sectionTemplate = xmlContent.substring(startBookmarkEndIndex, endBookmarkStartIndex)
        .replace(/<w:bookmark(Start|End)[^>]*>/g, '');

      if (!Array.isArray(excelData)) {
        excelData = [excelData];
      }
      let allStudentSections = "";
      for (let i = 0; i < excelData.length; i++) {
        const student = excelData[i];
        let studentSection = sectionTemplate;

        const mapping = {
          'placeholdersj': escapeXml(dashboardData.schuljahr || ''),
          'placeholdersl': escapeXml(dashboardData.schulleitung || ''),
          'sltitel': escapeXml(dashboardData.sl_titel || ''),
          'kltitel': escapeXml(dashboardData.kl_titel || ''),
          'zeugnisdatum': escapeXml(formatIsoDate(dashboardData.datum) || ''),
          'placeholderkl': escapeXml(dashboardData.klassenleitung || '')
        };

        Object.entries(student).forEach(([key, value]) => {
          let formattedValue = value;
          if (key === 'gdat') {
            formattedValue = formatExcelDate(value);
          }
          let safeValue = escapeXml(formattedValue);
          if (safeValue.includes('\u0000')) {
            safeValue = safeValue.replace(/\u0000/g, '');
          }
          mapping[key] = safeValue;
        });

        mapping['placeholderklasse'] = escapeXml(student.KL || '');

        studentSection = studentSection.replace(
          new RegExp(escapeRegExp('placeholderklasse'), 'g'),
          mapping['placeholderklasse'] || ''
        );
        studentSection = studentSection.replace(
          new RegExp(escapeRegExp('placeholderkl'), 'g'),
          mapping['placeholderkl'] || ''
        );
        Object.keys(mapping)
          .filter(key => key !== 'placeholderklasse' && key !== 'placeholderkl')
          .sort((a, b) => b.length - a.length)
          .forEach(key => {
            const regex = new RegExp(escapeRegExp(key), 'g');
            studentSection = studentSection.replace(regex, mapping[key]);
          });

        const pageBreak = `
          <w:p>
            <w:r>
              <w:br w:type="page"/>
            </w:r>
          </w:p>
        `;
        if (i < excelData.length - 1) {
          studentSection += pageBreak;
        }
        allStudentSections += studentSection;
      }

      const beforeSection = xmlContent.substring(0, startBookmarkEndIndex);
      const afterSection = xmlContent.substring(endBookmarkStartIndex);
      let newXmlContent = beforeSection + allStudentSections + afterSection;

      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(newXmlContent, "text/xml");
      if (xmlDoc.getElementsByTagName("parsererror").length > 0) {
        console.error("XML ist invalide:", xmlDoc.getElementsByTagName("parsererror")[0].textContent);
        throw new Error("Generiertes XML ist fehlerhaft");
      }

      Object.keys(originalZip.files).forEach((relativePath) => {
        const file = originalZip.file(relativePath);
        if (!zip.file(relativePath)) {
          zip.file(relativePath, file.asUint8Array());
        }
      });
      zip.file(documentXmlPath, newXmlContent);

      const out = zip.generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document',
        compression: 'DEFLATE',
        compressionOptions: { level: 9 }
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

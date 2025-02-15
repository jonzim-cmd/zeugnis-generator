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

// Funktion zum Auslesen des Inhalts zwischen zwei Lesezeichen (Bookmarks) mittels DOMParser und XPath
const findBookmarkContent = (xmlContent, startBookmarkName, endBookmarkName) => {
  const parser = new DOMParser();
  const xmlDoc = parser.parseFromString(xmlContent, "text/xml");
  const nsResolver = xmlDoc.createNSResolver(xmlDoc.documentElement);

  // Finde das Ende des Start-Lesezeichens (bookmarkEnd, dessen w:id entspricht dem w:id des bookmarkStart)
  const startBookmark = xmlDoc.evaluate(
    `//w:bookmarkStart[@w:name='${startBookmarkName}']`,
    xmlDoc,
    nsResolver,
    XPathResult.ANY_UNORDERED_NODE_TYPE,
    null
  ).singleNodeValue;
  if (!startBookmark) {
    throw new Error(`Lesezeichen "${startBookmarkName}" nicht gefunden`);
  }
  const startBookmarkEnd = xmlDoc.evaluate(
    `//w:bookmarkEnd[@w:id=(//w:bookmarkStart[@w:name='${startBookmarkName}']/@w:id)]`,
    xmlDoc,
    nsResolver,
    XPathResult.ANY_UNORDERED_NODE_TYPE,
    null
  ).singleNodeValue;
  if (!startBookmarkEnd) {
    throw new Error(`bookmarkEnd für "${startBookmarkName}" nicht gefunden`);
  }

  // Finde das Start-Lesezeichen des Endbereichs anhand des Namens
  const endBookmarkStart = xmlDoc.evaluate(
    `//w:bookmarkStart[@w:name='${endBookmarkName}']`,
    xmlDoc,
    nsResolver,
    XPathResult.ANY_UNORDERED_NODE_TYPE,
    null
  ).singleNodeValue;
  if (!endBookmarkStart) {
    throw new Error(`Lesezeichen "${endBookmarkName}" nicht gefunden`);
  }

  // Extrahiere den Inhalt zwischen startBookmarkEnd und endBookmarkStart
  let content = "";
  let currentNode = startBookmarkEnd.nextSibling;
  while (currentNode && currentNode !== endBookmarkStart) {
    content += new XMLSerializer().serializeToString(currentNode);
    currentNode = currentNode.nextSibling;
  }
  return content;
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

  // Dieses Skript erzeugt ein einzelnes Dokument, in dem der wiederholbare Bereich
  // (zwischen den Lesezeichen STUDENT_SECTION_START und STUDENT_SECTION_END)
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

      // Hole den wiederholbaren Abschnitt mithilfe der Lesezeichen-Funktion
      const studentSectionTemplate = findBookmarkContent(xmlContent, "STUDENT_SECTION_START", "STUDENT_SECTION_END");

      // Erzeuge den neuen, duplizierten Bereich für alle Excelzeilen
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

        // Beginne mit dem extrahierten Template-Bereich und ersetze zuerst spezielle Platzhalter
        let studentSection = studentSectionTemplate;
        studentSection = studentSection.replace(new RegExp(escapeRegExp('placeholderklasse'), 'g'), mapping['placeholderklasse']);
        studentSection = studentSection.replace(new RegExp(escapeRegExp('placeholderkl'), 'g'), mapping['placeholderkl']);

        // Ersetze alle weiteren Platzhalter (längere zuerst)
        const keys = Object.keys(mapping)
          .filter(key => key !== 'placeholderklasse' && key !== 'placeholderkl')
          .sort((a, b) => b.length - a.length);
        keys.forEach((key) => {
          const regex = new RegExp(escapeRegExp(key), 'g');
          studentSection = studentSection.replace(regex, mapping[key]);
        });
        newStudentSections += studentSection;
      }

      // DOM-Manipulation: Parsen des gesamten XML in einen DOM-Baum
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(xmlContent, "text/xml");
      const nsResolver = xmlDoc.createNSResolver(xmlDoc.documentElement);

      // Finde das bookmarkEnd von STUDENT_SECTION_START
      const startBookmarkEnd = xmlDoc.evaluate(
        `//w:bookmarkEnd[@w:id=(//w:bookmarkStart[@w:name='STUDENT_SECTION_START']/@w:id)]`,
        xmlDoc,
        nsResolver,
        XPathResult.ANY_UNORDERED_NODE_TYPE,
        null
      ).singleNodeValue;
      if (!startBookmarkEnd) {
        throw new Error('bookmarkEnd für "STUDENT_SECTION_START" nicht gefunden');
      }

      // Finde das bookmarkStart von STUDENT_SECTION_END
      const endBookmarkStart = xmlDoc.evaluate(
        `//w:bookmarkStart[@w:name='STUDENT_SECTION_END']`,
        xmlDoc,
        nsResolver,
        XPathResult.ANY_UNORDERED_NODE_TYPE,
        null
      ).singleNodeValue;
      if (!endBookmarkStart) {
        throw new Error('Lesezeichen "STUDENT_SECTION_END" nicht gefunden');
      }

      // Entferne alle Knoten zwischen startBookmarkEnd und endBookmarkStart
      let currentNode = startBookmarkEnd.nextSibling;
      while (currentNode && currentNode !== endBookmarkStart) {
        const next = currentNode.nextSibling;
        currentNode.parentNode.removeChild(currentNode);
        currentNode = next;
      }

      // Parsen des neuen Inhalts als XML – Wrapper mit explizitem Word-Namespace
      const fragmentWrapper = parser.parseFromString(
        `<wrapper xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main">${newStudentSections}</wrapper>`,
        "text/xml"
      ).documentElement;
      // Füge alle Kindelemente des Wrappers zwischen startBookmarkEnd und endBookmarkStart ein
      while (fragmentWrapper.firstChild) {
        const child = fragmentWrapper.firstChild;
        startBookmarkEnd.parentNode.insertBefore(xmlDoc.importNode(child, true), endBookmarkStart);
      }

      // Serialisiere den DOM zurück in einen String
      xmlContent = new XMLSerializer().serializeToString(xmlDoc);
      zip.file(documentXmlPath, xmlContent);

      const out = zip.generate({
        type: 'blob',
        mimeType: 'application/vnd.openxmlformats-officedocument.wordprocessingml.document'
      });

      // Erstelle ein einzelnes Dokument, das alle Schülerseiten enthält
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

import React, { useState } from 'react';
import PizZip from 'pizzip';
import { saveAs } from 'file-saver';

// Hilfsfunktionen zum Escapen von Sonderzeichen
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

// Extrahiere den ersten <w:sectPr>-Block aus dem XML (angenommen, es gibt nur einen im Template)
const extractOriginalSectPr = (xmlContent) => {
  const sectPrRegex = /<w:sectPr[\s\S]*?<\/w:sectPr>/;
  const match = xmlContent.match(sectPrRegex);
  return match ? match[0] : '';
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

      // Lade das DOCX (ZIP-Archiv)
      const zip = new PizZip(arrayBuffer);
      const documentXmlPath = 'word/document.xml';
      if (!zip.file(documentXmlPath)) {
        throw new Error('Dokumentstruktur ungültig: word/document.xml nicht gefunden');
      }
      const xmlContent = zip.file(documentXmlPath).asText();

      // Extrahiere den <w:body>-Bereich als Vorlage
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

      // Extrahiere den originalen <w:sectPr>-Block und entferne ihn aus dem Template
      const originalSectPr = extractOriginalSectPr(studentTemplate);
      if (originalSectPr) {
        studentTemplate = studentTemplate.replace(originalSectPr, '');
      }

      // Erzeuge einen kontinuierlichen Sektionseinstellungsblock, 
      // indem der w:type auf "continuous" gesetzt wird:
      const continuousSectPr = originalSectPr
        ? originalSectPr.replace(/w:type="[^"]*"/, 'w:type="continuous"')
        : '';

      // --- Erzeuge für jeden Schüler einen Abschnitt ---
      let allStudentSections = "";
      if (!Array.isArray(excelData)) {
        excelData = [excelData];
      }
      excelData.forEach((student, i) => {
        let studentSection = studentTemplate;

        // Mapping aus Dashboard- und Excel-Daten
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

        // Ersetze zuerst "placeholderklasse" und dann "placeholderkl"
        studentSection = studentSection.replace(
          new RegExp(escapeRegExp('placeholderklasse'), 'g'),
          mapping['placeholderklasse']
        );
        studentSection = studentSection.replace(
          new RegExp(escapeRegExp('placeholderkl'), 'g'),
          mapping['placeholderkl']
        );
        // Ersetze alle übrigen Platzhalter (längere zuerst)
        Object.keys(mapping)
          .filter(key => key !== 'placeholderklasse' && key !== 'placeholderkl')
          .sort((a, b) => b.length - a.length)
          .forEach(key => {
            const regex = new RegExp(escapeRegExp(key), 'g');
            studentSection = studentSection.replace(regex, mapping[key]);
          });

        // Für alle Schüler außer dem letzten: Füge einen Absatz mit kontinuierlichen Sektionseinstellungen und Seitenumbruch ein
        if (i < excelData.length - 1 && continuousSectPr) {
          studentSection += `
            <w:p>
              <w:pPr>
                <w:sectPr>
                  ${continuousSectPr}
                </w:sectPr>
              </w:pPr>
              <w:r>
                <w:br w:type="page"/>
              </w:r>
            </w:p>`;
        }
        
        allStudentSections += studentSection;
      });

      // Am Ende des Dokuments: Hänge den finalen Absatz mit den originalen Sektionseinstellungen an
      const finalSection = originalSectPr
        ? `<w:p>
             <w:pPr>
               ${originalSectPr}
             </w:pPr>
           </w:p>`
        : '';
      const newBodyContent = allStudentSections + finalSection;

      // Füge den neuen Body wieder in das komplette Dokument ein
      const newXmlContent = preBody + newBodyContent + postBody;

      // (Optional) Überprüfe die XML-Struktur
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

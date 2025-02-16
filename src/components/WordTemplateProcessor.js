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
      const zip = new PizZip(arrayBuffer);
      
      const documentXmlPath = 'word/document.xml';
      if (!zip.file(documentXmlPath)) {
        throw new Error('Dokumentstruktur ungültig: word/document.xml nicht gefunden');
      }
      const xmlContent = zip.file(documentXmlPath).asText();

      const bodyStartTag = '<w:body>';
      const bodyEndTag = '</w:body>';
      const bodyStartIndex = xmlContent.indexOf(bodyStartTag);
      const bodyEndIndex = xmlContent.indexOf(bodyEndTag);
      
      if (bodyStartIndex === -1 || bodyEndIndex === -1) {
        throw new Error('Die benötigten <w:body>-Tags wurden nicht gefunden.');
      }

      const preBody = xmlContent.substring(0, bodyStartIndex + bodyStartTag.length);
      const postBody = xmlContent.substring(bodyEndIndex);
      let templateContent = xmlContent.substring(
        bodyStartIndex + bodyStartTag.length, 
        bodyEndIndex
      ).trim();

      // Extrahiere die ursprünglichen Sektionseinstellungen und speichere alle Formatierungen
      const sectPrMatch = templateContent.match(/<w:sectPr[\s\S]*?<\/w:sectPr>/);
      const originalSectPr = sectPrMatch ? sectPrMatch[0] : '';
      
      // Extrahiere die Header/Footer Referenzen aus dem originalSectPr
      const headerFooterRefs = originalSectPr.match(/<w:headerReference[^>]*>[\s\S]*?<\/w:headerReference>|<w:footerReference[^>]*>[\s\S]*?<\/w:footerReference>/g) || [];
      
      // Erstelle eine modifizierte Version für kontinuierliche Seitenumbrüche
      let continuousSectPr = originalSectPr
        .replace(/w:type="[^"]*"/g, 'w:type="continuous"')
        .replace(/<w:headerReference[^>]*>[\s\S]*?<\/w:headerReference>|<w:footerReference[^>]*>[\s\S]*?<\/w:footerReference>/g, '');

      // Entferne das ursprüngliche sectPr aus dem Template
      templateContent = templateContent.replace(/<w:sectPr[\s\S]*?<\/w:sectPr>/, '');

      const studentSections = (Array.isArray(excelData) ? excelData : [excelData])
        .map((student, index, array) => {
          let section = templateContent;
          
          const mapping = {
            'placeholdersj': escapeXml(dashboardData.schuljahr || ''),
            'placeholdersl': escapeXml(dashboardData.schulleitung || ''),
            'sltitel': escapeXml(dashboardData.sl_titel || ''),
            'kltitel': escapeXml(dashboardData.kl_titel || ''),
            'zeugnisdatum': escapeXml(formatIsoDate(dashboardData.datum) || ''),
            'placeholderkl': escapeXml(dashboardData.klassenleitung || ''),
            'placeholderklasse': escapeXml(student.KL || ''),
            'gdat': escapeXml(formatExcelDate(student.gdat) || ''),
            ...Object.fromEntries(
              Object.entries(student)
                .filter(([key]) => !['KL', 'gdat'].includes(key))
                .map(([key, value]) => [key, escapeXml(value)])
            )
          };

          Object.entries(mapping)
            .sort(([a], [b]) => b.length - a.length)
            .forEach(([key, value]) => {
              const regex = new RegExp(escapeRegExp(key), 'g');
              section = section.replace(regex, value);
            });

          // Für alle außer dem letzten Schüler fügen wir einen Seitenumbruch mit kontinuierlicher Sektion ein
          if (index < array.length - 1) {
            // Füge die Header/Footer Referenzen zur kontinuierlichen Sektion hinzu
            const sectionWithRefs = continuousSectPr.replace('</w:sectPr>', 
              `${headerFooterRefs.join('\n')}</w:sectPr>`);
            
            // Finde den letzten Absatz und füge dort den Seitenumbruch ein
            const lastParaMatch = section.match(/<w:p[^>]*>(?:(?!<w:p[^>]*>).)*?<\/w:p>(?=[^]*$)/);
            if (lastParaMatch) {
              const lastParagraph = lastParaMatch[0];
              const modifiedParagraph = lastParagraph.replace(
                /<\/w:p>$/,
                `<w:r><w:br w:type="page"/></w:r></w:p>`
              );
              section = section.replace(lastParaMatch[0], modifiedParagraph);
              
              // Füge die Sektionseinstellungen in einem separaten Absatz hinzu
              section += `<w:p><w:pPr>${sectionWithRefs}</w:pPr></w:p>`;
            }
          }

          return section;
        })
        .join('');

      // Füge die finale Sektion mit allen ursprünglichen Einstellungen hinzu
      const finalSection = `<w:p><w:pPr>${originalSectPr}</w:pPr></w:p>`;
      const newXmlContent = `${preBody}${studentSections}${finalSection}${postBody}`;

      // XML Validierung
      const parser = new DOMParser();
      const xmlDoc = parser.parseFromString(newXmlContent, "text/xml");
      if (xmlDoc.getElementsByTagName("parsererror").length > 0) {
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
    <div>
      <button onClick={generateDocx} disabled={processing}>
        {processing ? 'Generiere...' : 'Word-Dokument erstellen'}
      </button>
    </div>
  );
};

export default WordTemplateProcessor;

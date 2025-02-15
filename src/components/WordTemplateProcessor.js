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
    // Suche nach dem Start-Lesezeichen (STUDENT_SECTION_START)
    const startBookmarkRegex = new RegExp(
      `<w:bookmarkStart[^>]*w:name="STUDENT_SECTION_START"[^>]*>`,
      'g'
    );
    const startMatch = startBookmarkRegex.exec(xmlContent);
    if (!startMatch) {
      throw new Error('Lesezeichen "STUDENT_SECTION_START" nicht gefunden');
    }
    const startIndex = startMatch.index + startMatch[0].length;

    // Suche nach dem End-Lesezeichen (STUDENT_SECTION_END)
    const endBookmarkRegex = new RegExp(
      `<w:bookmarkStart[^>]*w:name="STUDENT_SECTION_END"[^>]*>`,
      'g'
    );
    const endMatch = endBookmarkRegex.exec(xmlContent);
    if (!endMatch) {
      throw new Error('Lesezeichen "STUDENT_SECTION_END" nicht gefunden');
    }
    const endIndex = endMatch.index;

    // Der Template-Abschnitt, der dupliziert werden soll
    const sectionTemplate = xmlContent.substring(startIndex, endIndex);

    // --- 2. Für jeden Schüler den Template-Bereich mit Platzhalterersetzung generieren ---
    let allStudentSections = "";
    for (let i = 0; i < excelData.length; i++) {
      const student = excelData[i];

      // Mapping: Die Keys müssen exakt den Platzhaltertexten im Word-Dokument entsprechen
      const mapping = {
        // Dashboard-Daten
        'placeholdersj': dashboardData.schuljahr || '',
        'placeholdersl': dashboardData.schulleitung || '',
        'sltitel': dashboardData.sl_titel || '',
        'kltitel': dashboardData.kl_titel || '',
        'zeugnisdatum': formatIsoDate(dashboardData.datum) || '',
        'placeholderkl': dashboardData.klassenleitung || '',
        // Excel-Daten
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

      let studentSection = sectionTemplate;
      // Erst den Excel-Klassenplatzhalter ersetzen, falls beide Varianten vorhanden sind
      studentSection = studentSection.replace(
        new RegExp(escapeRegExp('placeholderklasse'), 'g'),
        mapping['placeholderklasse']
      );
      // Dann den Dashboard-Platzhalter für Klassenleitung
      studentSection = studentSection.replace(
        new RegExp(escapeRegExp('placeholderkl'), 'g'),
        mapping['placeholderkl']
      );
      // Alle übrigen Platzhalter ersetzen – längere zuerst, um Überschneidungen zu vermeiden
      Object.keys(mapping)
        .filter(key => key !== 'placeholderklasse' && key !== 'placeholderkl')
        .sort((a, b) => b.length - a.length)
        .forEach(key => {
          const regex = new RegExp(escapeRegExp(key), 'g');
          studentSection = studentSection.replace(regex, mapping[key]);
        });

      // Optional: Falls du sicherstellen möchtest, dass jeder Schülerabschnitt auf einer eigenen Seite beginnt,
      // füge hier einen Word-XML-Seitenumbruch ein (z. B. am Ende jedes Abschnitts, außer beim letzten)
      if (i < excelData.length - 1) {
        studentSection += `<w:p><w:r><w:br w:type="page"/></w:r></w:p>`;
      }
      
      allStudentSections += studentSection;
    }

    // --- 3. Neuen XML-Inhalt zusammenbauen ---
    // Vorheriger Teil bis zum Ende des Start-Lesezeichens
    const beforeSection = xmlContent.substring(0, startIndex);
    // Nachfolgender Teil ab dem Beginn des End-Lesezeichens
    const afterSection = xmlContent.substring(endIndex);

    // Setze den neuen Inhalt zusammen: Die originale Struktur bleibt erhalten, 
    // zwischen den Lesezeichen wird der neue, duplizierte Inhalt eingefügt.
    const newXmlContent = beforeSection + allStudentSections + afterSection;
    zip.file(documentXmlPath, newXmlContent);

    // Generiere den Blob und speichere die fertige Datei
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

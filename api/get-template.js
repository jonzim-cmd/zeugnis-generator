// api/get-template.js

export default async function handler(req, res) {
  // Dynamischer Import von node-fetch (ESM-Modul)
  const { default: fetch } = await import('node-fetch');

  // Hole Token und Username aus den Environment-Variablen
  const token = process.env.GITHUB_TOKEN;
  const username = process.env.GITHUB_USERNAME;

  // Mapping der Zeugnisarten auf Dateinamen
  const templateMap = {
    'Zwischenzeugnis': 'template_zwischen.docx',
    'Abschlusszeugnis': 'template_abschluss.docx',
    'Jahreszeugnis': 'template_jahr.docx'
  };

  // Lese den Query-Parameter "zeugnisart"
  const { zeugnisart } = req.query;
  
  // Prüfe, ob eine gültige Zeugnisart übergeben wurde
  if (!zeugnisart || !templateMap[zeugnisart]) {
    return res.status(400).json({ error: 'Ungültige oder fehlende Zeugnisart' });
  }
  
  // Bestimme den Dateinamen basierend auf der Zeugnisart
  const fileName = templateMap[zeugnisart];
  
  // Erstelle die URL zum Abruf der Datei aus dem privaten Repository (mit Verweis auf den Main-Branch)
  const templateUrl = `https://api.github.com/repos/jonzim-cmd/private-template-repo/contents/${fileName}?ref=main`;

  try {
    const response = await fetch(templateUrl, {
      headers: {
        'Authorization': `token ${token}`,
        'Accept': 'application/vnd.github.v3.raw',
        // GitHub erwartet oft einen User-Agent Header
        'User-Agent': username
      }
    });

    if (!response.ok) {
      const errorMsg = await response.text();
      return res.status(response.status).send(`GitHub API Fehler: ${errorMsg}`);
    }

    // Lese die Datei als ArrayBuffer ein
    const arrayBuffer = await response.arrayBuffer();

    // Setze den Content-Type für DOCX
    res.setHeader('Content-Type', 'application/vnd.openxmlformats-officedocument.wordprocessingml.document');
    res.send(Buffer.from(arrayBuffer));
  } catch (error) {
    console.error('Fehler beim Abruf des Templates:', error);
    res.status(500).send('Interner Serverfehler');
  }
}

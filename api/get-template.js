// api/get-template.js
import fetch from 'node-fetch';

export default async function handler(req, res) {
  // Hole den Token und den Username aus den Environment-Variablen
  const token = process.env.GITHUB_TOKEN;
  const username = process.env.GITHUB_USERNAME;

  // URL der Datei im privaten Repository (GitHub API)
  const templateUrl = 'https://api.github.com/repos/jonzim-cmd/private-template-repo/contents/template_zwischen.docx';

  try {
    const response = await fetch(templateUrl, {
      headers: {
        'Authorization': `token ${token}`,
        'Accept': 'application/vnd.github.v3.raw',
        // GitHub verlangt häufig auch einen User-Agent Header
        'User-Agent': username
      }
    });

    if (!response.ok) {
      return res.status(response.status).send('Template nicht gefunden');
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

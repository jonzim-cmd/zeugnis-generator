// api/getTemplate.js
import fs from 'fs';
import path from 'path';

export default function handler(req, res) {
  if (req.method === "GET") {
    // Pfad zum Template in einem nicht-öffentlichen Ordner (nicht im public-Ordner)
    const templatePath = path.join(process.cwd(), 'private-templates', 'template.docx');
    
    fs.readFile(templatePath, (err, data) => {
      if (err) {
        console.error("Template konnte nicht geladen werden:", err);
        res.status(500).json({ error: "Template nicht gefunden" });
      } else {
        res.setHeader(
          "Content-Type",
          "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        );
        res.status(200).send(data);
      }
    });
  } else {
    res.status(405).json({ error: "Method not allowed" });
  }
}

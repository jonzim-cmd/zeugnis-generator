// api/validate.js
export default async function handler(req, res) {
  if (req.method === "POST") {
    const { password } = req.body;
    // MY_SECRET_PASSWORD wird als Environment Variable in Vercel gesetzt
    if (password === process.env.MY_SECRET_PASSWORD) {
      res.status(200).json({ success: true });
    } else {
      res.status(401).json({ success: false, error: "Ungültiges Passwort" });
    }
  } else {
    res.status(405).json({ error: "Method not allowed" });
  }
}

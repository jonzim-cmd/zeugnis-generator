#!/bin/bash
# Dieses Skript erstellt zur Build-Zeit die .gitmodules-Datei für den Zugriff
# auf das private Submodul und nutzt dabei die in Vercel gesetzten Umgebungsvariablen.

if [ -z "$GITHUB_USERNAME" ] || [ -z "$GITHUB_TOKEN" ]; then
  echo "Fehler: GITHUB_USERNAME oder GITHUB_TOKEN sind nicht gesetzt."
  exit 1
fi

# Erstelle eine neue .gitmodules-Datei (die eventuell vorhandene .gitmodules.disabled bleibt unberührt)
cat > .gitmodules <<EOF
[submodule "private-templates"]
  path = private-templates
  url = https://${GITHUB_USERNAME}:${GITHUB_TOKEN}@github.com/jonzim-cmd/private-template-repo.git
EOF

# Optional: Zeige den Inhalt der neuen .gitmodules-Datei zur Kontrolle (ohne den Token sichtbar auszugeben)
echo "Neue .gitmodules-Datei wurde erstellt."
grep "url =" .gitmodules

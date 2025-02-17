#!/bin/bash
# Dieses Skript erstellt zur Build-Zeit die .gitmodules-Datei für den Zugriff auf das private Submodul.
# Es nutzt die in Vercel gesetzten Umgebungsvariablen GITHUB_USERNAME und GITHUB_TOKEN.

cat > .gitmodules <<EOF
[submodule "private-templates"]
  path = private-templates
  url = https://${GITHUB_USERNAME}:${GITHUB_TOKEN}@github.com/jonzim-cmd/private-template-repo.git
EOF

# Optional: Ausgabe zur Kontrolle
cat .gitmodules

#!/usr/bin/env python3
"""<App>-files.sha256 erstellen: SHA256 aller Dateien des ausgecheckten Commits (Inhalt wie im Repo, ohne .github/).
Format wie sha256sum: '<hash>  <pfad>' - Pull.ps1 prueft damit jede geladene Datei.
Ohne Argument wird der Name aus Config/app.json gebildet ("<id>-files.sha256")."""
import hashlib, json, subprocess, sys

if len(sys.argv) > 1:
    out = sys.argv[1]
else:
    out = json.load(open('Config/app.json', encoding='utf-8-sig'))['id'] + '-files.sha256'
files = [f for f in subprocess.check_output(['git', 'ls-files', '-z']).decode('utf-8').split('\0') if f and not f.startswith('.github/')]
lines = []
for f in sorted(files):
    data = subprocess.check_output(['git', 'show', 'HEAD:' + f])
    lines.append(f"{hashlib.sha256(data).hexdigest()}  {f}")
with open(out, 'w', encoding='utf-8', newline='\n') as fh:
    fh.write('\n'.join(lines) + '\n')
print(f"{len(lines)} Dateien -> {out}")

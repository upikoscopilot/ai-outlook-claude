# Claude Mail – Outlook-Add-in (Fork von sctg-development/ai-outlook)

Claude-Assistent für Outlook: Text in einer neuen Mail markieren → Prompt wählen → Ergebnis ersetzt die Markierung.

## Installation (empfohlen: GitHub Pages, kein Server nötig)
1. https://aka.ms/olksideload öffnen und mit dem Microsoft-365-Konto anmelden.
2. „Meine Add-Ins“ → „Benutzerdefiniertes Add-In hinzufügen“ → „Aus Datei hinzufügen…“ →
   `pages\manifest.xml` (Kopie von https://upikoscopilot.github.io/ai-outlook-claude/manifest.xml).
   Wichtig: nicht `dist\manifest.xml` – das zeigt auf localhost.
3. Neue Mail aus dem Microsoft-Konto → Apps → „Claude Mail“ → API-Key von console.anthropic.com eintragen.

Add-ins funktionieren nur bei Microsoft-Konten (Microsoft 365 / outlook.com), nicht bei Gmail-Konten im neuen Outlook.

## Was gegenüber dem Original geändert wurde
- Nur Anthropic/Claude (Sonnet 5.5 Standard, Opus 5.5, Haiku 5.5), Direktaufruf von api.anthropic.com – kein Fremd-Proxy mehr.
- 7 deutsche Prompts (verbessern, förmlicher, kürzen, Antwort, Zusammenfassung, DE→EN, EN→DE).
- Fehler beim Einfügen werden angezeigt; Sonderzeichen < & > bleiben erhalten; API-Key-Felder maskiert.

## Alternative: lokal über https://localhost:3000
Node.js nötig. `einmalig-zertifikat-installieren.cmd`, dann `autostart-einrichten.cmd` (Server startet unsichtbar bei jeder Anmeldung) und `manifest.xml` aus `dist` (Build mit `npm run build`) hinzufügen. `node-pruefen.cmd` prüft die Umgebung.

## Bauen
`npm install`, dann `npm run build:pages` (GitHub Pages) bzw. `npm run build` (localhost).

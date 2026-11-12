# Konsilager-Inventur

Web-App zum Scannen von UDI-Codes (QR / DataMatrix) auf Medizinprodukten im Konsignationslager.
Läuft im Browser auf Android, iPhone und iPad – ohne Installation.

## So funktioniert's

1. **Soll-Liste** antippen und die Excel-Bestandsliste wählen (z. B. aus OneDrive/SharePoint).
   Erkannt werden die Spalten *Artikel-Nr., Artikelbezeichnung, Seite, Größe, Soll, Ist, Chargen-Nr.*
   Gruppen-Überschriften und Leerzeilen werden übersprungen.
2. **Scannen**: UDI-Code vor die Kamera halten. Chargennummer, GTIN (UDI-DI), Verfallsdatum und
   Seriennummer werden automatisch ausgelesen (GS1, GS1 Digital Link und HIBC).
3. Beim **ersten Scan** eines Artikels die passende Zeile (Artikel, Seite, Größe) auswählen.
   Die App merkt sich die Zuordnung und erkennt denselben Code danach automatisch.
4. **Export → In Original-Liste eintragen**: Die importierte Excel-Datei wird mit Ist-Menge,
   Chargen-Nr., Inventurdatum und „Geprüft von“ ausgefüllt. Formatierung und Formeln
   (Differenz, Status) bleiben erhalten. Nicht gescannte Artikel bekommen Ist = 0 → „nachbestellen“.
   Alternativ: eigene Auswertung als Excel oder PDF.

Alle Daten bleiben im Browser des jeweiligen Geräts gespeichert (auch offline nutzbar, sobald geladen).
Tipp: Über „Zum Startbildschirm hinzufügen“ verhält sich die Seite wie eine App.

## Gemeinsame Liste in SharePoint

Nach **Anmelden** mit dem Firmenkonto arbeitet die App direkt in der geteilten Excel-Liste
(Link und App-Kennung in `src/config.ts`):

- Jeder Scan erhöht sofort „Ist“ in der passenden Zeile und ergänzt die Chargen-Nr.
  Es werden nur einzelne Zellen geändert, mehrere Personen können gleichzeitig scannen.
- Der Stand der Kollegen wird jede Minute neu geladen (oder über ⟳).
- Gelernte Zuordnungen (UDI-DI → Zeile) landen im versteckten Blatt „App-Zuordnungen“ und gelten für alle.
- **Abschließen & Exportieren → Inventur abschließen** setzt leere Ist-Felder auf 0 und trägt Datum und Prüfer ein.
- **Einstellungen → Neue Inventur starten** leert Ist und Chargen-Nr. aller Artikelzeilen.

Voraussetzung: App-Registrierung in Microsoft Entra (Single-Page-Anwendung, Umleitungs-URI
`https://janssturmaktiv.github.io/Inventur-APP/`, delegierte Berechtigung `Files.ReadWrite.All`).

## Entwicklung

```bash
npm install
npm run dev      # lokaler Server (für Kamera auf dem Handy wird HTTPS benötigt)
npm test         # Tests für UDI-Auswertung, Abgleich, Excel-Import und -Rückschreiben
npm run build    # fertige Web-App in dist/
```

Bei jedem Push auf `main` baut GitHub Actions die App und veröffentlicht sie auf GitHub Pages
(Einstellung: *Settings → Pages → Source: GitHub Actions*).

## Aufbau

- `src/App.tsx` – Oberfläche
- `src/Scanner.tsx` – Kamera und Code-Erkennung (ZXing, läuft lokal im Browser)
- `src/lib/udi.ts`, `src/lib/gs1.ts` – Auswertung der UDI-Codes
- `src/lib/targetList.ts` – Import der Soll-Liste
- `src/lib/reconcile.ts` – Soll-Ist-Abgleich
- `src/lib/xlsxWriteback.ts` – Werte in die Original-Excel schreiben (formatgetreu)
- `src/lib/exportWeb.ts` – Excel-/PDF-Export

# Gemeinsame Liste ohne Admin: E-Mail + Power Automate

So landen die Scans aller Kollegen in der geteilten Excel-Liste in SharePoint, ohne dass ein
Administrator etwas freigeben muss. Einmalige Einrichtung: ca. 15 Minuten, am besten am PC.

**Ablauf später:** Scannen in der App → „Senden“ → Mail-App öffnet sich mit fertigem Text → „Senden“ tippen →
Power Automate trägt Ist und Chargen-Nr. nach ca. 1 Minute in die Liste ein.

---

## Schritt 1: Skript in Excel anlegen

1. Die Bestandsliste in SharePoint **im Browser** öffnen (Excel für das Web, nicht die Desktop-App).
2. Oben den Reiter **Automatisieren** wählen → **Neues Skript**.
3. Den vorhandenen Beispielcode im Code-Editor rechts komplett löschen.
4. Den Code aus dieser Datei einfügen:
   **[office-script/inventur-scans-eintragen.ts](office-script/inventur-scans-eintragen.ts)**
   (auf GitHub die Datei öffnen, oben rechts auf das Symbol **„Copy raw file“** klicken, dann in Excel einfügen).
5. Oben auf den Skriptnamen klicken und ihn in **`Inventur-Scans eintragen`** ändern, dann **Skript speichern**.

> Kein Reiter „Automatisieren“ da? Dann sind Office Scripts in eurem Konto ausgeschaltet. Bitte Bescheid geben.

## Schritt 2: Power-Automate-Ablauf erstellen

1. [make.powerautomate.com](https://make.powerautomate.com) öffnen (mit dem Firmenkonto).
2. **Erstellen** → **Automatisierter Cloud-Flow**.
   - Name: `Inventur-Scans eintragen`
   - Auslöser suchen: **„Wenn eine neue E-Mail eingeht (V3)“** (Office 365 Outlook) → **Erstellen**.
3. Im Auslöser:
   - Ordner: **Inbox / Posteingang**
   - Unter „Erweiterte Parameter“ bei **Betrefffilter**: `INVENTUR-SCANS`
4. **+ Neuer Schritt / Aktion hinzufügen** → suchen: **„Skript ausführen“** (Excel Online (Business)).
   - Speicherort: **SharePoint-Website „TeamOT“**
   - Dokumentbibliothek: die Bibliothek, in der die Liste liegt (meist **Dokumente**)
   - Datei: die **Bestandsliste** auswählen
   - Skript: **Inventur-Scans eintragen**
   - **mailBody**: ins Feld klicken → dynamischer Inhalt → **Text** (bzw. „Body“) aus dem Auslöser
5. Damit sich mehrere gleichzeitig eintreffende Mails nicht in die Quere kommen:
   beim **Auslöser** auf **⋯ → Einstellungen** → **Parallelitätssteuerung** einschalten → **Grad der Parallelität: 1**.
6. **Speichern**.

*Optional:* Eine weitere Aktion **„E-Mail verschieben (V2)“** hinzufügen (Nachrichten-ID aus dem Auslöser,
Zielordner z. B. „Inventur“), damit die Mails nicht den Posteingang füllen.

## Schritt 3: App einrichten

In der App auf **⚙︎** tippen und bei **„E-Mail-Adresse für die gemeinsame Liste“** deine eigene
Firmen-E-Mail-Adresse eintragen (das Postfach, das der Ablauf überwacht).

**Für Kollegen** gibt es einen Link, der die Adresse gleich mitbringt:

```
https://janssturmaktiv.github.io/Inventur-APP/?an=DEINE-ADRESSE@firma.de
```

## Test

1. In der App die Liste laden, einen Artikel scannen, zuordnen, speichern.
2. Oben auf **Senden** → **E-Mail öffnen** → in der Mail-App **Senden**.
3. Nach ca. 1 Minute in der Excel-Liste nachsehen. In Power Automate unter **Meine Flows → Ausführungsverlauf**
   steht bei jeder Ausführung, wie viele Stück eingetragen wurden (oder welche Zeile nicht gefunden wurde).

## Inventur abschließen

In der App **Export → Inventur abschließen**. Die letzte E-Mail enthält dann zusätzlich den Abschluss:
Alle Artikel ohne Ist-Wert werden auf 0 gesetzt (→ „nachbestellen“), dazu Inventurdatum und „Geprüft von“.

Für eine **neue Inventur** in Excel die Spalten „Ist“ und „Chargen-Nr.“ leeren (bzw. wie bisher eine frische
Kopie der Liste anlegen).

## Empfehlung: Spalte „GTIN“ in der Liste

Füge in der Bestandsliste (Blatt „Inventur“) rechts neben „Bemerkung“ eine Spalte mit der Überschrift **GTIN** hinzu.
Das Skript trägt dort beim ersten Scan eines Artikels automatisch die GTIN aus dem QR-Code ein. Wer danach die
Liste in die App lädt, bekommt alle bereits bekannten Codes mit, und die App ordnet sie ohne Nachfrage zu.

Nach dem Hinzufügen der Spalte das Skript in Excel einmal neu einfügen (Schritt 1), damit die neueste Version läuft.

## Gut zu wissen

- Die App findet die Zeile über **Artikel-Nr. + Seite + Größe**. Diese drei Spalten in der Liste bitte nicht umbenennen.
- Gelernte Zuordnungen speichert jedes Handy für sich. Mit der Spalte „GTIN“ (siehe oben) werden sie über die Liste an alle verteilt.
- Jede E-Mail zählt die Menge **dazu**. Eine Mail doppelt zu senden zählt also doppelt.

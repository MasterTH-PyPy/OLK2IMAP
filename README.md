# OLK2Any

Rettet und exportiert die **Offline-Daten des „new Outlook" (Olk)** nach
**IMAP / Exchange (EWS) / MBOX / PST** – „Olk → Any".

> Hinweis: Das Produkt/der EXE-Name ist **OLK2Any**; interne Namespaces und
> Projektordner heißen aus Historie noch `OLK2IMAP.*`. Repo-/Namespace-Umbenennung
> ist optional und separat.

Das neue Outlook für Windows hat **keinen klassischen Mail-Store** (keine PST/OST).
Seine Offline-Daten liegen in der WebView2-**IndexedDB** der OWA-Web-App:

```
%LocalAppData%\Microsoft\Olk\EBWebView\Default\IndexedDB\
    https_outlook.office.com_0.indexeddb.leveldb   (Chromium LevelDB)
    https_outlook.office.com_0.indexeddb.blob      (externe Blobs)
%LocalAppData%\Microsoft\Olk\Attachments\          (gecachte Anhänge)
```

Nur was **offline vorlag**, kann wiederhergestellt werden – nicht das komplette
Server-Postfach.

## Architektur

Zwei Stufen, bewusst getrennt:

1. **Extraktor** (`extractor/olkextract.py`, als `olkextract.exe` gebündelt) –
   liest den Chromium-LevelDB-/V8-Store (via `ccl_chromium_reader`), rekonstruiert
   Ordnerbaum + Nachrichten und schreibt:
   - `EML/<Ordnerpfad>/<Nachricht>.eml`
   - `manifest.json` (Ordner + Metadaten je Nachricht: Ordner, Message-ID, Datum,
     gelesen/Entwurf)
2. **C#-WinForms-App** (`src/OLK2IMAP.App`) – liest `manifest.json` + EML-Baum und
   überträgt alles ins gewählte Ziel, Ordnerstruktur und Gelesen-Status erhalten.

Die harte Parsing-Arbeit (LevelDB + V8 Structured Clone) macht der erprobte
Python-Extraktor; C# macht GUI + Ziel-Anbindung (MailKit/EWS/Aspose).

## Ziele

| Ziel            | Lib                                    | Status |
|-----------------|----------------------------------------|--------|
| IMAP            | MailKit                                | fertig |
| MBOX            | MimeKit (Thunderbird-Import)           | fertig |
| Exchange (EWS)  | EWS Managed API (.NET-Standard-Port)   | Grundgerüst |
| PST             | Aspose.Email (kommerziell)             | Grundgerüst |

## Bauen

Voraussetzungen: VS 2026, .NET 10 SDK, Python 3.11+ (nur zum Extraktor-Build).

```bat
REM 1) Extraktor-EXE bauen (einmalig / bei Änderungen)
cd extractor
build_extractor.bat

REM 2) Solution bauen
cd ..
dotnet build -c Release
```

Beim App-Build wird `extractor\dist\olkextract\` automatisch nach
`bin\...\extractor\` kopiert. Im Dev-Betrieb ohne EXE ruft die App alternativ
`python extractor\olkextract.py` auf.

## Benutzung

1. **Olk-Profil** wählen (der `...\Microsoft\Olk`-Ordner bzw. ein Backup davon)
   und einen **Arbeitsordner**.
2. **1. Extrahieren** – erzeugt EML-Baum + `manifest.json`.
3. **Ziel** wählen und Zugangsdaten eintragen.
4. **2. Übertragen** – erst mit **Dry-Run** testen, dann scharf.

## Versionierung

Zentrale Version in `Directory.Build.props` (`<Version>`). Bei **jedem Commit**
den Patch-Stand erhöhen; die App zeigt die Version im Titel (`AppVersion.Display`).

## Lizenzhinweis

Aspose.Email (PST) ist kommerziell; ohne Lizenz nur Evaluierungsmodus. Lizenz
einmalig beim Start setzen:
`new Aspose.Email.License().SetLicense("Aspose.Email.lic");`

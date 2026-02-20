# TIA Program Block Exporter

Dieses Tool öffnet ein TIA-Projekt (`.apxx`) über TIA Openness und exportiert alle Programmbausteine in **eine XML-Datei**.

## Voraussetzungen

- TIA Portal inkl. Openness installiert
- .NET 8 SDK installiert
- Optional: Umgebungsvariable `TIA_PUBLIC_API_DIR` setzen, um eine bestimmte TIA-Version zu erzwingen (z. B. V19):

```powershell
$env:TIA_PUBLIC_API_DIR = "C:\Program Files\Siemens\Automation\Portal V19\PublicAPI\V19"
```

Hinweis: Wenn `TIA_PUBLIC_API_DIR` **nicht** gesetzt ist, bevorzugt das Projekt automatisch `Portal V17`, dann `V18`, dann `V19`.

## Build

```powershell
dotnet build
```

## Ausführung

Ohne Parameter startet eine Bedienoberfläche, in der das TIA-Projekt ausgewählt wird. Danach werden die gefundenen PLCs mit Programmstruktur geladen, eine PLC ausgewählt und der Export gestartet. Die Dateien werden automatisch im Projektordner abgelegt und enthalten den PLC-Namen im Dateinamen (z. B. `ProgramBlocksExport_PLCH-1.xml`, `ErrorLog_PLCH-1.txt`).

```powershell
dotnet run
```

Aus dem Projektordner `TiaBlockExporter`:

```powershell
dotnet run -- --project "C:\Users\Etteplan\Documents\Automation\TBMA_VP2_6\TBMA_VP2_6.ap19" --output "C:\Users\Etteplan\Documents\Automation\TBMA_VP2_6\ProgramBlocksExport.xml" --with-ui
```

Beispiel für ein V17-Projekt:

```powershell
dotnet run -- --project "C:\Users\Etteplan\Documents\Automation\Metso\905988-103 Ruifeng\905988-103 Ruifeng.ap17" --output "C:\Users\Etteplan\Documents\Automation\Metso\905988-103 Ruifeng\ProgramBlocksExport.xml" --with-ui
```

Aus dem Workspace-Root `VS-Code`:

```powershell
dotnet run --project ".\TiaBlockExporter\TiaBlockExporter.csproj" -- --project "C:\Users\Etteplan\Documents\Automation\TBMA_VP2_6\TBMA_VP2_6.ap19" --output "C:\Users\Etteplan\Documents\Automation\TBMA_VP2_6\ProgramBlocksExport.xml" --with-ui
```

## Known Issues

- Die PLC-Liste zeigt nur PLCs mit erkannter Programmstruktur; reine HMI-/Visualisierungsgeräte werden nicht exportiert.
- Je nach Projektgröße kann das Laden der PLCs und der Export länger dauern; der Fortschrittsbalken ist aktuell ein Aktivitätsindikator (keine Prozentanzeige).
- Bei mehreren installierten TIA-Versionen kann ohne explizit gesetzte Umgebungsvariable `TIA_PUBLIC_API_DIR` eine unpassende Openness-Version verwendet werden.
- Einzelne Bausteine können wegen Zugriffsrechten, Schutz oder projektspezifischen Einstellungen fehlschlagen; Details stehen in `ErrorLog_<PLC>.txt`.
- Fehlende TIA-Produkte in der aktiven Version (z. B. WinCC Professional) verhindern das Öffnen des Projekts.

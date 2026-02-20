# TIA Program Block Exporter

Dieses Tool öffnet ein TIA-Projekt (`.ap19`) über TIA Openness und exportiert alle Programmbausteine in **eine XML-Datei**.

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

Aus dem Projektordner `TiaBlockExporter`:

```powershell
dotnet run -- --project "C:\Users\Etteplan\Documents\Automation\TBMA_VP2_6\TBMA_VP2_6.ap19" --output "C:\Users\Etteplan\Documents\Automation\TBMA_VP2_6\ProgramBlocksExport.xml" --with-ui
```
...
Ruifeng
dotnet run -- --project "C:\Users\Etteplan\Documents\Automation\Metso\905988-103 Ruifeng\905988-103 Ruifeng.ap17" --output "C:\Users\Etteplan\Documents\Automation\Metso\905988-103 Ruifeng\ProgramBlocksExport.xml" --with-ui

neu

dotnet run -- --project "C:\Users\Etteplan\Documents\Automation\Metso\905988-103 Ruifeng\905988-103 Ruifeng.ap17" --output "C:\Users\Etteplan\Documents\Automation\Metso\905988-103 Ruifeng\ProgramBlocksExport.xml" --with-ui

C:\Users\Etteplan\Documents\Automation\Metso\905988-103 Ruifeng\905988-103 Ruifeng.ap17
Aus dem Workspace-Root `VS-Code`:

```powershell
dotnet run --project ".\TiaBlockExporter\TiaBlockExporter.csproj" -- --project "C:\Users\Etteplan\Documents\Automation\TBMA_VP2_6\TBMA_VP2_6.ap19" --output "C:\Users\Etteplan\Documents\Automation\TBMA_VP2_6\ProgramBlocksExport.xml" --with-ui
```

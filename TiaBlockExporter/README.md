# TIA Program Block Exporter

Dieses Tool öffnet ein TIA-Projekt (`.ap19`) über TIA Openness und exportiert alle Programmbausteine in **eine XML-Datei**.

## Voraussetzungen

- TIA Portal inkl. Openness installiert
- .NET 8 SDK installiert
- Umgebungsvariable `TIA_PUBLIC_API_DIR` gesetzt, z. B.:

```powershell
$env:TIA_PUBLIC_API_DIR = "C:\Program Files\Siemens\Automation\Portal V19\PublicAPI\V19"
```

## Build

```powershell
dotnet build
```

## Ausführung

Aus dem Projektordner `TiaBlockExporter`:

```powershell
dotnet run -- --project "C:\Users\Etteplan\Documents\Automation\TBMA_VP2_6\TBMA_VP2_6.ap19" --output "C:\Users\Etteplan\Documents\Automation\TBMA_VP2_6\ProgramBlocksExport.xml" --with-ui
```

Aus dem Workspace-Root `VS-Code`:

```powershell
dotnet run --project ".\TiaBlockExporter\TiaBlockExporter.csproj" -- --project "C:\Users\Etteplan\Documents\Automation\TBMA_VP2_6\TBMA_VP2_6.ap19" --output "C:\Users\Etteplan\Documents\Automation\TBMA_VP2_6\ProgramBlocksExport.xml" --with-ui
```

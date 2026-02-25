using System.Runtime.Remoting;
using System.Threading.Tasks;
using System.Globalization;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using System.Xml.Linq;
using System.Windows.Forms;
using NPOI.HSSF.Util;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using Siemens.Engineering;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;

internal static class Program
{
    private const string PlcBlowerFromDbName = "091 Signal Exchange From PLC Blower DB";
    private const string PlcBlowerToDbName = "911 Signal Excange To PLC Blower DB";

    [STAThread]
    private static int Main(string[] args)
    {
        var arguments = ParseArguments(args);
        if (!arguments.TryGetValue("project", out var projectPath) || string.IsNullOrWhiteSpace(projectPath))
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new ExportForm());
            return 0;
        }

        if (arguments.TryGetValue("excel", out var excelPath) && !string.IsNullOrWhiteSpace(excelPath)
            && arguments.TryGetValue("worksheet", out var worksheetName) && !string.IsNullOrWhiteSpace(worksheetName))
        {
            var withUiForBuild = !arguments.ContainsKey("without-ui");
            arguments.TryGetValue("plc-device", out var plcDevice);
            arguments.TryGetValue("plc-software", out var plcSoftware);
            arguments.TryGetValue("timeout-minutes", out var timeoutMinutes);
            return RunCliBuild(projectPath!, excelPath!, worksheetName!, withUiForBuild, plcDevice, plcSoftware, timeoutMinutes);
        }

        var outputPath = arguments.TryGetValue("output", out var output) && !string.IsNullOrWhiteSpace(output)
            ? output
            : Path.Combine(Path.GetDirectoryName(Path.GetFullPath(projectPath))!, "ProgramBlocksExport.xml");

        var withUi = arguments.ContainsKey("with-ui");
        return RunCliExport(projectPath!, outputPath!, withUi);
    }

    internal static ExportResult ExportProject(string projectPath, string? outputPath, bool withUi, PlcSelection? selectedPlc = null)
    {
        var resolvedProjectPath = Path.GetFullPath(projectPath);
        var resolvedOutputPath = ResolveOutputPath(resolvedProjectPath, outputPath, selectedPlc?.SoftwareName);

        if (!File.Exists(resolvedProjectPath))
        {
            throw new FileNotFoundException($"Projektdatei nicht gefunden: {resolvedProjectPath}");
        }

        var exportSuccessCount = 0;
        var exportErrorCount = 0;
        var errorLogEntries = new List<string>();
        var warningMessage = string.Empty;

        Project? project = null;
        using var tiaPortal = new TiaPortal(withUi ? TiaPortalMode.WithUserInterface : TiaPortalMode.WithoutUserInterface);
        project = tiaPortal.Projects.Open(new FileInfo(resolvedProjectPath));
        var softwareInfos = EnumeratePlcSoftware(project).ToList();
        if (selectedPlc is not null)
        {
            softwareInfos = softwareInfos
                .Where(info => string.Equals(info.DeviceName, selectedPlc.DeviceName, StringComparison.OrdinalIgnoreCase)
                    && string.Equals(info.SoftwareName, selectedPlc.SoftwareName, StringComparison.OrdinalIgnoreCase))
                .ToList();
        }

        var root = new XElement("TiaProgramBlockExport",
            new XAttribute("project", resolvedProjectPath),
            new XAttribute("exportedAt", DateTimeOffset.Now.ToString("O")));

        if (softwareInfos.Count == 0)
        {
            warningMessage = "Keine PLC-Software gefunden. Prüfe, ob das Projekt PLC-Geräte enthält oder ob die passende TIA-Version über TIA_PUBLIC_API_DIR verwendet wird.";
            root.Add(new XElement("Warning", warningMessage));
        }

        foreach (var softwareInfo in softwareInfos)
        {
            var softwareElement = new XElement("PlcSoftware",
                new XAttribute("device", softwareInfo.DeviceName),
                new XAttribute("software", softwareInfo.SoftwareName));

            ExportBlockGroup(softwareInfo.PlcSoftware.BlockGroup, softwareElement, string.Empty, ref exportSuccessCount, ref exportErrorCount, errorLogEntries);
            root.Add(softwareElement);
        }

        var document = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root);
        Directory.CreateDirectory(Path.GetDirectoryName(resolvedOutputPath) ?? ".");
        document.Save(resolvedOutputPath);

        var errorLogPath = ResolveErrorLogPath(resolvedOutputPath, selectedPlc?.SoftwareName);
        WriteErrorLog(errorLogPath, errorLogEntries);

        TryCloseProject(project);

        return new ExportResult
        {
            OutputPath = resolvedOutputPath,
            ExportSuccessCount = exportSuccessCount,
            ExportErrorCount = exportErrorCount,
            WarningMessage = warningMessage,
            ErrorLogPath = errorLogPath
        };
    }

    private static int RunCliExport(string projectPath, string outputPath, bool withUi)
    {
        var currentDirectory = Directory.GetCurrentDirectory();
        var resolvedProjectPath = Path.GetFullPath(projectPath);
        var resolvedOutputPath = Path.GetFullPath(outputPath);

        Console.WriteLine($"Arbeitsverzeichnis: {currentDirectory}");
        Console.WriteLine($"Parameter --project: {projectPath}");
        Console.WriteLine($"Aufgelöster Projektpfad: {resolvedProjectPath}");
        Console.WriteLine($"Parameter --output: {outputPath}");
        Console.WriteLine($"Aufgelöster Outputpfad: {resolvedOutputPath}");

        try
        {
            var result = ExportProject(projectPath, outputPath, withUi);

            if (!string.IsNullOrWhiteSpace(result.WarningMessage))
            {
                Console.WriteLine($"Warnung: {result.WarningMessage}");
            }

            Console.WriteLine($"Export abgeschlossen: {result.OutputPath}");
            Console.WriteLine($"Error-Log: {result.ErrorLogPath}");
            Console.WriteLine($"Statistik: Erfolgreich={result.ExportSuccessCount}, Fehler={result.ExportErrorCount}, Gesamt={result.ExportSuccessCount + result.ExportErrorCount}");
            return 0;
        }
        catch (RemotingException ex)
        {
            Console.Error.WriteLine("Fehler beim Export: TIA Portal ist nicht mehr erreichbar (Prozess beendet oder Verbindung verloren).");
            Console.Error.WriteLine(ex.Message);
            return 98;
        }
        catch (ObjectDisposedException ex)
        {
            Console.Error.WriteLine("Fehler beim Export: Ein TIA-Objekt wurde während der Verarbeitung ungültig.");
            Console.Error.WriteLine(ex.Message);
            return 98;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Fehler beim Export:");
            Console.Error.WriteLine(ex.Message);
            return 99;
        }
    }

    private static int RunCliBuild(string projectPath, string excelPath, string worksheetName, bool withUi, string? plcDevice, string? plcSoftware, string? timeoutMinutesRaw)
    {
        var resolvedProjectPath = Path.GetFullPath(projectPath);
        var resolvedExcelPath = Path.GetFullPath(excelPath);
        var portalProcessIdsBeforeRun = GetTiaPortalProcessIds();
        var hasTimeout = !string.IsNullOrWhiteSpace(timeoutMinutesRaw);
        var timeoutMinutes = 0;
        if (hasTimeout && (!int.TryParse(timeoutMinutesRaw, NumberStyles.Integer, CultureInfo.InvariantCulture, out timeoutMinutes) || timeoutMinutes <= 0))
        {
            Console.Error.WriteLine("Fehler beim Build: --timeout-minutes muss eine positive ganze Zahl sein.");
            return 96;
        }

        Console.WriteLine($"Build Projekt: {resolvedProjectPath}");
        Console.WriteLine($"Build Excel: {resolvedExcelPath}");
        Console.WriteLine($"Build Worksheet: {worksheetName}");
        if (hasTimeout)
        {
            Console.WriteLine($"Build Timeout: {timeoutMinutes} Minute(n)");
        }

        try
        {
            (PlcSelection SelectedPlc, BuildResult Result) execution;
            if (hasTimeout)
            {
                var timeout = TimeSpan.FromMinutes(timeoutMinutes);
                (PlcSelection SelectedPlc, BuildResult Result)? threadResult = null;
                Exception? threadException = null;

                var workerThread = new Thread(() =>
                {
                    try
                    {
                        threadResult = ExecuteCliBuildOperation(projectPath, withUi, excelPath, worksheetName, plcDevice, plcSoftware);
                    }
                    catch (Exception ex)
                    {
                        threadException = ex;
                    }
                })
                {
                    IsBackground = true,
                    Name = "CliBuildWorker"
                };

                workerThread.Start();
                if (!workerThread.Join(timeout))
                {
                    var timeoutLogPath = WriteCliTimeoutBuildLog(resolvedProjectPath, worksheetName, plcSoftware, timeout);
                    Console.Error.WriteLine($"Build nach {timeoutMinutes} Minute(n) abgebrochen. Build-Log: {timeoutLogPath}");
                    TryCloseSpawnedTiaPortalProcesses(portalProcessIdsBeforeRun);
                    Environment.Exit(124);
                    return 124;
                }

                if (threadException is not null)
                {
                    throw threadException;
                }

                execution = threadResult ?? throw new InvalidOperationException("Build konnte nicht abgeschlossen werden.");
            }
            else
            {
                execution = ExecuteCliBuildOperation(projectPath, withUi, excelPath, worksheetName, plcDevice, plcSoftware);
            }

            Console.WriteLine($"Verwendete PLC: {execution.SelectedPlc.DisplayName}");
            var result = execution.Result;

            Console.WriteLine($"Build abgeschlossen. Strukturen={result.MatchedStructureCount}, Rot={result.RedEntryCount}, Ordner={result.CreatedFolderCount}, Blöcke={result.CreatedBlockCount}, Übersprungen={result.SkippedItemCount}, Fehler={result.ErrorCount}");
            if (!string.IsNullOrWhiteSpace(result.WarningMessage))
            {
                Console.WriteLine($"Warnung: {result.WarningMessage}");
            }

            Console.WriteLine($"Build-Log: {result.BuildLogPath}");
            return 0;
        }
        catch (RemotingException ex)
        {
            Console.Error.WriteLine("Fehler beim Build: TIA Portal ist nicht mehr erreichbar.");
            Console.Error.WriteLine(ex.Message);
            return 98;
        }
        catch (ObjectDisposedException ex)
        {
            Console.Error.WriteLine("Fehler beim Build: Ein TIA-Objekt wurde während der Verarbeitung ungültig.");
            Console.Error.WriteLine(ex.Message);
            return 98;
        }
        catch (Exception ex)
        {
            Console.Error.WriteLine("Fehler beim Build:");
            Console.Error.WriteLine(ex.Message);
            return 99;
        }
        finally
        {
            TryCloseSpawnedTiaPortalProcesses(portalProcessIdsBeforeRun);
        }
    }

    private static (PlcSelection SelectedPlc, BuildResult Result) ExecuteCliBuildOperation(
        string projectPath,
        bool withUi,
        string excelPath,
        string worksheetName,
        string? plcDevice,
        string? plcSoftware)
    {
        var plcs = DiscoverExportablePlcs(projectPath, false);
        if (plcs.Count == 0)
        {
            throw new InvalidOperationException("Keine exportierbare PLC gefunden.");
        }

        PlcSelection selectedPlc;
        if (!string.IsNullOrWhiteSpace(plcDevice) || !string.IsNullOrWhiteSpace(plcSoftware))
        {
            selectedPlc = plcs.FirstOrDefault(plc =>
                    (string.IsNullOrWhiteSpace(plcDevice) || string.Equals(plc.DeviceName, plcDevice, StringComparison.OrdinalIgnoreCase))
                    && (string.IsNullOrWhiteSpace(plcSoftware) || string.Equals(plc.SoftwareName, plcSoftware, StringComparison.OrdinalIgnoreCase)))
                ?? throw new InvalidOperationException("Angegebene PLC nicht gefunden.");
        }
        else
        {
            selectedPlc = plcs[0];
        }

        var result = BuildFromExcelTemplate(projectPath, withUi, selectedPlc, excelPath, worksheetName);
        return (selectedPlc, result);
    }

    private static HashSet<int> GetTiaPortalProcessIds()
    {
        return Process.GetProcessesByName("Siemens.Automation.Portal")
            .Select(process => process.Id)
            .ToHashSet();
    }

    private static void TryCloseSpawnedTiaPortalProcesses(HashSet<int> processIdsBeforeRun)
    {
        try
        {
            var currentIds = Process.GetProcessesByName("Siemens.Automation.Portal")
                .Select(process => process.Id)
                .ToList();

            foreach (var processId in currentIds.Where(id => !processIdsBeforeRun.Contains(id)))
            {
                try
                {
                    using var process = Process.GetProcessById(processId);
                    process.Kill();
                    process.WaitForExit(5000);
                }
                catch
                {
                }
            }
        }
        catch
        {
        }
    }

    private static string WriteCliTimeoutBuildLog(string resolvedProjectPath, string worksheetName, string? plcName, TimeSpan timeout)
    {
        var preferredLogPath = ResolveBuildLogPath(resolvedProjectPath, plcName);
        var lines = new List<string>
        {
            "TIA Program Block Exporter - BuildLog",
            $"Erstellt am: {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
            $"Excel-Tabelle: {worksheetName}",
            string.Empty,
            "Statistik:",
            "- Strukturpaare verarbeitet: 0",
            "- Ordner erstellt: 0",
            "- Blöcke erstellt: 0",
            "- Übersprungen: 0",
            "- Übersprungen (Excel-Whitelist): 0",
            "- Vorlagenfehler: 0",
            "- Fehler: 1",
            string.Empty,
            "Info / Diagnose:",
            $"- Build-Fehler: Timeout nach {(int)timeout.TotalMinutes} Minute(n). Lauf wurde per CLI-Abbruch beendet.",
            string.Empty,
            "Details (gekürzt):",
            "Fehler: 1",
            "- Build-Timeout (CLI)",
            string.Empty
        };

        try
        {
            File.WriteAllLines(preferredLogPath, lines);
            return preferredLogPath;
        }
        catch
        {
            var fallbackLogPath = Path.Combine(Path.GetTempPath(), $"TiaBlockBuildLog_TIMEOUT_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}.txt");
            File.WriteAllLines(fallbackLogPath, lines);
            return fallbackLogPath;
        }
    }

    private static void TryCloseProject(Project? project)
    {
        if (project is null)
        {
            return;
        }

        try
        {
            project.Close();
        }
        catch (RemotingException)
        {
        }
        catch (ObjectDisposedException)
        {
        }
    }

    private static Dictionary<string, string?> ParseArguments(string[] args)
    {
        var result = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

        for (var index = 0; index < args.Length; index++)
        {
            var current = args[index];
            if (!current.StartsWith("--", StringComparison.Ordinal))
            {
                continue;
            }

            var key = current.Substring(2);
            if (index + 1 < args.Length && !args[index + 1].StartsWith("--", StringComparison.Ordinal))
            {
                result[key] = args[index + 1];
                index++;
            }
            else
            {
                result[key] = null;
            }
        }

        return result;
    }

    internal static List<PlcSelection> DiscoverExportablePlcs(string projectPath, bool withUi)
    {
        var resolvedProjectPath = Path.GetFullPath(projectPath);
        if (!File.Exists(resolvedProjectPath))
        {
            throw new FileNotFoundException($"Projektdatei nicht gefunden: {resolvedProjectPath}");
        }

        Project? project = null;
        using var tiaPortal = new TiaPortal(withUi ? TiaPortalMode.WithUserInterface : TiaPortalMode.WithoutUserInterface);
        try
        {
            project = tiaPortal.Projects.Open(new FileInfo(resolvedProjectPath));
            return EnumeratePlcSoftware(project)
                .Where(info => HasProgramStructure(info.PlcSoftware.BlockGroup))
                .Select(info => new PlcSelection
                {
                    DeviceName = info.DeviceName,
                    SoftwareName = info.SoftwareName
                })
                .GroupBy(selection => $"{selection.DeviceName}|{selection.SoftwareName}")
                .Select(group => group.First())
                .OrderBy(selection => selection.DeviceName)
                .ThenBy(selection => selection.SoftwareName)
                .ToList();
        }
        finally
        {
            TryCloseProject(project);
        }
    }

    private static IEnumerable<(string DeviceName, string SoftwareName, PlcSoftware PlcSoftware)> EnumeratePlcSoftware(Project project)
    {
        foreach (var device in EnumerateProjectDevices(project))
        {
            foreach (var deviceItem in TraverseDeviceItems(device.DeviceItems))
            {
                var softwareContainer = deviceItem.GetService<SoftwareContainer>();
                if (softwareContainer?.Software is PlcSoftware plcSoftware)
                {
                    yield return (device.Name, plcSoftware.Name, plcSoftware);
                }
            }
        }
    }

    private static IEnumerable<Device> EnumerateProjectDevices(Project project)
    {
        var yieldedDevices = new HashSet<Device>();

        foreach (Device device in project.Devices)
        {
            if (yieldedDevices.Add(device))
            {
                yield return device;
            }
        }

        if (project.UngroupedDevicesGroup is not null)
        {
            foreach (Device device in project.UngroupedDevicesGroup.Devices)
            {
                if (yieldedDevices.Add(device))
                {
                    yield return device;
                }
            }
        }

        foreach (DeviceUserGroup deviceGroup in project.DeviceGroups)
        {
            foreach (Device device in TraverseDeviceUserGroupDevices(deviceGroup))
            {
                if (yieldedDevices.Add(device))
                {
                    yield return device;
                }
            }
        }
    }

    private static IEnumerable<Device> TraverseDeviceUserGroupDevices(DeviceUserGroup group)
    {
        foreach (Device device in group.Devices)
        {
            yield return device;
        }

        foreach (DeviceUserGroup childGroup in group.Groups)
        {
            foreach (var childDevice in TraverseDeviceUserGroupDevices(childGroup))
            {
                yield return childDevice;
            }
        }
    }

    private static IEnumerable<DeviceItem> TraverseDeviceItems(DeviceItemComposition composition)
    {
        foreach (DeviceItem item in composition)
        {
            yield return item;

            foreach (var child in TraverseDeviceItems(item.DeviceItems))
            {
                yield return child;
            }
        }
    }

    private static bool HasProgramStructure(PlcBlockGroup group)
    {
        if (group.Blocks.Count > 0)
        {
            return true;
        }

        foreach (PlcBlockGroup childGroup in group.Groups)
        {
            if (HasProgramStructure(childGroup))
            {
                return true;
            }
        }

        return false;
    }

    private static void ExportBlockGroup(PlcBlockGroup group, XElement parentElement, string groupPath, ref int exportSuccessCount, ref int exportErrorCount, List<string> errorLogEntries)
    {
        var currentGroupPath = string.IsNullOrEmpty(groupPath) ? group.Name : $"{groupPath}/{group.Name}";

        foreach (PlcBlock block in group.Blocks)
        {
            var blockElement = new XElement("Block",
                new XAttribute("name", block.Name),
                new XAttribute("type", block.GetType().Name),
                new XAttribute("groupPath", currentGroupPath));

            string? tempFile = null;
            try
            {
                tempFile = Path.Combine(Path.GetTempPath(), $"TiaBlockExport_{Guid.NewGuid():N}.xml");
                block.Export(new FileInfo(tempFile), ExportOptions.WithDefaults);
                var exportedXml = XDocument.Load(tempFile);
                blockElement.Add(new XElement("ExportXml", exportedXml.Root));
                exportSuccessCount++;
            }
            catch (Exception ex)
            {
                blockElement.Add(new XElement("ExportError", ex.Message));
                errorLogEntries.Add($"[{DateTime.Now:yyyy-MM-dd HH:mm:ss}] Block-Exportfehler | Gruppe: {currentGroupPath} | Block: {block.Name} | Typ: {block.GetType().Name} | Fehler: {ex.Message}");
                exportErrorCount++;
            }
            finally
            {
                if (!string.IsNullOrWhiteSpace(tempFile) && File.Exists(tempFile))
                {
                    File.Delete(tempFile);
                }
            }

            parentElement.Add(blockElement);
        }

        foreach (PlcBlockGroup childGroup in group.Groups)
        {
            ExportBlockGroup(childGroup, parentElement, currentGroupPath, ref exportSuccessCount, ref exportErrorCount, errorLogEntries);
        }
    }

    private static void WriteErrorLog(string errorLogPath, List<string> errorLogEntries)
    {
        var lines = new List<string>
        {
            "TIA Program Block Exporter - ErrorLog",
            $"Erstellt am: {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
            string.Empty
        };

        if (errorLogEntries.Count == 0)
        {
            lines.Add("Keine Fehler aufgetreten.");
        }
        else
        {
            lines.AddRange(errorLogEntries);
        }

        File.WriteAllLines(errorLogPath, lines);
    }

    internal static BuildResult BuildFromExcelTemplate(string projectPath, bool withUi, PlcSelection selectedPlc, string excelPath, string worksheetName)
    {
        var resolvedProjectPath = Path.GetFullPath(projectPath);
        if (!File.Exists(resolvedProjectPath))
        {
            throw new FileNotFoundException($"Projektdatei nicht gefunden: {resolvedProjectPath}");
        }

        var templateMappings = new List<TemplateMapping>();
        var worksheetAnalysis = new WorksheetAnalysis();
        var blowerDbEntries = new List<BlowerDbEntry>();
        var blowerAnalysis = new BlowerWorksheetAnalysis();
        var buildLogEntries = new List<string>();
        var createdFolderCount = 0;
        var createdBlockCount = 0;
        var skippedItemCount = 0;
        var errorCount = 0;
        var warningMessage = string.Empty;
        var buildException = (Exception?)null;
        var preferredLogPath = ResolveBuildLogPath(resolvedProjectPath, selectedPlc.SoftwareName);
        var buildLogPath = preferredLogPath;

        Project? project = null;
        using var tiaPortal = new TiaPortal(withUi ? TiaPortalMode.WithUserInterface : TiaPortalMode.WithoutUserInterface);
        try
        {
            templateMappings = GetTemplateMappingsFromWorksheet(excelPath, worksheetName, out worksheetAnalysis);
            blowerDbEntries = GetPlcBlowerDbEntriesFromWorksheet(excelPath, worksheetName, out blowerAnalysis);
            buildLogEntries.Add($"Excel-Diagnose: Datenzeilen={worksheetAnalysis.DataRowCount}, I/O TAG gefüllt={worksheetAnalysis.IoTagValueCount}, Function gefüllt={worksheetAnalysis.FunctionValueCount}, gültige Zeilen={worksheetAnalysis.CandidateRowCount}, rot erkannt={worksheetAnalysis.RedMarkedCount}, Function-Gruppen(exakt)={worksheetAnalysis.ExactFunctionGroupCount}, Function-Gruppen(pattern)={worksheetAnalysis.PatternGroupCount}");
            buildLogEntries.Add($"PLC-Blower-Diagnose: System='PLC Blower' Zeilen={blowerAnalysis.PlcBlowerRowCount}, rot markiert={blowerAnalysis.RedMarkedPlcBlowerRowCount}, nicht rot übersprungen={blowerAnalysis.NonRedSkippedCount}, gültige DB-Einträge={blowerDbEntries.Count}, ungültiger Connecting Type={blowerAnalysis.InvalidConnectingTypeCount}, fehlender I/O TAG={blowerAnalysis.MissingIoTagCount}");
            if (worksheetAnalysis.RedColorSamples.Count > 0)
            {
                buildLogEntries.Add($"Excel-Rotfarben (Samples): {string.Join(" | ", worksheetAnalysis.RedColorSamples)}");
            }
            if (templateMappings.Count == 0)
            {
                buildLogEntries.Add("Hinweis: Keine roten Einträge mit passender Referenzstruktur gefunden.");
            }

            if (blowerDbEntries.Count == 0)
            {
                buildLogEntries.Add("Hinweis: Keine gültigen PLC-Blower-DB-Einträge gefunden.");
            }

            if (templateMappings.Count == 0 && blowerDbEntries.Count == 0)
            {
                warningMessage = "Es wurden weder Struktur-Mappings noch PLC-Blower-DB-Einträge in der Excel-Tabelle gefunden.";
            }

            if (templateMappings.Count > 0 || blowerDbEntries.Count > 0)
            {
                project = tiaPortal.Projects.Open(new FileInfo(resolvedProjectPath));
                var selectedSoftwareInfo = EnumeratePlcSoftware(project)
                    .FirstOrDefault(info => string.Equals(info.DeviceName, selectedPlc.DeviceName, StringComparison.OrdinalIgnoreCase)
                        && string.Equals(info.SoftwareName, selectedPlc.SoftwareName, StringComparison.OrdinalIgnoreCase));

                if (selectedSoftwareInfo.PlcSoftware is null)
                {
                    throw new InvalidOperationException($"Ausgewählte PLC wurde im Projekt nicht gefunden: {selectedPlc.DisplayName}");
                }

                if (templateMappings.Count > 0)
                {
                    var replacementPlans = new Dictionary<string, ReplacementPlan>(StringComparer.OrdinalIgnoreCase);

                    foreach (var mapping in templateMappings)
                    {
                        foreach (var targetName in mapping.TargetEntries)
                        {
                            var replacement = FindStructureReplacement(mapping.SourceEntry, targetName);
                            if (replacement is null)
                            {
                                skippedItemCount++;
                                buildLogEntries.Add($"Übersprungen (keine erkennbare Struktur-Ersetzung): Quelle='{mapping.SourceEntry}' Ziel='{targetName}'");
                                continue;
                            }

                            var replacementKey = $"{replacement.Value.SourceStructure}->{replacement.Value.TargetStructure}";
                            if (!replacementPlans.TryGetValue(replacementKey, out var plan))
                            {
                                plan = new ReplacementPlan
                                {
                                    ReplacementKey = replacementKey,
                                    SourceToken = replacement.Value.SourceStructure,
                                    TargetToken = replacement.Value.TargetStructure,
                                    MappingKey = mapping.StructureKey
                                };

                                replacementPlans[replacementKey] = plan;
                            }

                            plan.AllowedTargetIoTags.Add(targetName);
                        }
                    }

                    foreach (var plan in replacementPlans.Values.OrderBy(value => value.ReplacementKey, StringComparer.OrdinalIgnoreCase))
                    {
                        buildLogEntries.Add($"Strukturpaar verarbeitet: {plan.ReplacementKey}");

                        CloneStructureByToken(
                            selectedSoftwareInfo.PlcSoftware.BlockGroup,
                            plan.SourceToken,
                            plan.TargetToken,
                            plan.MappingKey,
                            plan.AllowedTargetIoTags,
                            buildLogEntries,
                            ref createdFolderCount,
                            ref createdBlockCount,
                            ref skippedItemCount,
                            ref errorCount);
                    }
                }

                var createdBlowerDbEntries = 0;
                if (blowerDbEntries.Count > 0)
                {
                    createdBlowerDbEntries = ExtendPlcBlowerSignalExchangeDbs(
                        selectedSoftwareInfo.PlcSoftware.BlockGroup,
                        blowerDbEntries,
                        buildLogEntries,
                        ref skippedItemCount,
                        ref errorCount);
                    buildLogEntries.Add($"PLC-Blower DB-Einträge erstellt: {createdBlowerDbEntries}");
                }

                if (createdFolderCount > 0 || createdBlockCount > 0 || createdBlowerDbEntries > 0)
                {
                    project.Save();
                    buildLogEntries.Add("Projekt gespeichert.");
                }
                else
                {
                    buildLogEntries.Add("Keine Änderungen zum Speichern.");
                }
            }
        }
        catch (Exception ex)
        {
            buildException = ex;
            errorCount++;
            warningMessage = string.IsNullOrWhiteSpace(warningMessage)
                ? "Build mit Fehler beendet. Details im Build-Log."
                : warningMessage;
            buildLogEntries.Add($"Build-Fehler: {BuildExceptionSummary(ex)}");
        }
        finally
        {
            TryCloseProject(project);
            buildLogPath = SafeWriteBuildLog(preferredLogPath, worksheetName, templateMappings, buildLogEntries);
        }

        if (buildException is not null)
        {
            throw new InvalidOperationException($"Build fehlgeschlagen. Build-Log: {buildLogPath}", buildException);
        }

        return new BuildResult
        {
            WarningMessage = warningMessage,
            MatchedStructureCount = templateMappings.Count,
            RedEntryCount = templateMappings.Sum(mapping => mapping.TargetEntries.Count),
            CreatedFolderCount = createdFolderCount,
            CreatedBlockCount = createdBlockCount,
            SkippedItemCount = skippedItemCount,
            ErrorCount = errorCount,
            BuildLogPath = buildLogPath
        };
    }

    private static string SafeWriteBuildLog(string preferredLogPath, string worksheetName, List<TemplateMapping> mappings, List<string> buildLogEntries)
    {
        try
        {
            WriteBuildLog(preferredLogPath, worksheetName, mappings, buildLogEntries);
            return preferredLogPath;
        }
        catch (Exception ex)
        {
            var fallbackLogPath = Path.Combine(Path.GetTempPath(), $"TiaBlockBuildLog_{DateTime.Now:yyyyMMdd_HHmmss}_{Guid.NewGuid():N}.txt");
            var fallbackEntries = new List<string>(buildLogEntries)
            {
                $"Hinweis: Schreiben des Build-Logs am Standardpfad fehlgeschlagen: {CompactLogMessage(ex.Message)}",
                $"Fallback-Logpfad verwendet: {fallbackLogPath}"
            };

            WriteBuildLog(fallbackLogPath, worksheetName, mappings, fallbackEntries);
            return fallbackLogPath;
        }
    }

    private static List<TemplateMapping> GetTemplateMappingsFromWorksheet(string excelPath, string worksheetName, out WorksheetAnalysis analysis)
    {
        analysis = new WorksheetAnalysis();
        var redColorSamples = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        using var stream = File.OpenRead(excelPath);
        IWorkbook workbook = Path.GetExtension(excelPath).Equals(".xls", StringComparison.OrdinalIgnoreCase)
            ? new HSSFWorkbook(stream)
            : new XSSFWorkbook(stream);

        var worksheet = workbook.GetSheet(worksheetName)
            ?? throw new InvalidOperationException($"Die Tabelle '{worksheetName}' wurde in der Excel-Datei nicht gefunden.");

        var formatter = new DataFormatter(CultureInfo.CurrentCulture);
        var headerRow = worksheet.GetRow(worksheet.FirstRowNum)
            ?? throw new InvalidOperationException("Die Excel-Tabelle enthält keine Kopfzeile.");

        var ioTagColumnIndex = FindColumnIndexByHeader(headerRow, formatter, "I/O TAG");
        var functionColumnIndex = FindColumnIndexByHeader(headerRow, formatter, "Function");

        if (ioTagColumnIndex < 0)
        {
            throw new InvalidOperationException("Die Spalte 'I/O TAG' wurde in der Kopfzeile nicht gefunden.");
        }

        if (functionColumnIndex < 0)
        {
            throw new InvalidOperationException("Die Spalte 'Function' wurde in der Kopfzeile nicht gefunden.");
        }

        var entries = new List<WorksheetEntry>();
        for (var rowIndex = headerRow.RowNum + 1; rowIndex <= worksheet.LastRowNum; rowIndex++)
        {
            analysis.DataRowCount++;

            var row = worksheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            var ioTagCell = row.GetCell(ioTagColumnIndex);
            var functionCell = row.GetCell(functionColumnIndex);

            var ioTagValue = formatter.FormatCellValue(ioTagCell)?.Trim();
            if (string.IsNullOrWhiteSpace(ioTagValue))
            {
                continue;
            }

            analysis.IoTagValueCount++;

            var functionValue = formatter.FormatCellValue(functionCell)?.Trim();
            if (string.IsNullOrWhiteSpace(functionValue))
            {
                continue;
            }

            analysis.FunctionValueCount++;

            var isRedMarked = ioTagCell is not null && IsCellMarkedRed(ioTagCell);
            if (isRedMarked)
            {
                analysis.RedMarkedCount++;
                if (ioTagCell is not null)
                {
                    var debugColor = DescribeCellRedColor(ioTagCell);
                    if (!string.IsNullOrWhiteSpace(debugColor))
                    {
                        redColorSamples.Add(debugColor);
                    }
                }
            }

            entries.Add(new WorksheetEntry
            {
                Value = ioTagValue!,
                StructureValue = functionValue!,
                IsRed = isRedMarked,
                CellAddress = $"R{rowIndex + 1}C{ioTagColumnIndex + 1}"
            });

            analysis.CandidateRowCount++;
        }

        analysis.ExactFunctionGroupCount = entries
            .Select(entry => NormalizeFunctionValue(entry.StructureValue))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Count();

        analysis.PatternGroupCount = entries
            .Select(entry => BuildStructureKey(entry.StructureValue))
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .Count();

        analysis.RedColorSamples = redColorSamples.Take(12).ToList();

        var exactFunctionMappings = BuildTemplateMappings(entries, entry => NormalizeFunctionValue(entry.StructureValue), "Function=");
        if (exactFunctionMappings.Count > 0)
        {
            return exactFunctionMappings;
        }

        return BuildTemplateMappings(entries, entry => BuildStructureKey(entry.StructureValue), "Pattern=");
    }

    private static List<BlowerDbEntry> GetPlcBlowerDbEntriesFromWorksheet(string excelPath, string worksheetName, out BlowerWorksheetAnalysis analysis)
    {
        analysis = new BlowerWorksheetAnalysis();
        using var stream = File.OpenRead(excelPath);
        IWorkbook workbook = Path.GetExtension(excelPath).Equals(".xls", StringComparison.OrdinalIgnoreCase)
            ? new HSSFWorkbook(stream)
            : new XSSFWorkbook(stream);

        var worksheet = workbook.GetSheet(worksheetName)
            ?? throw new InvalidOperationException($"Die Tabelle '{worksheetName}' wurde in der Excel-Datei nicht gefunden.");

        var formatter = new DataFormatter(CultureInfo.CurrentCulture);
        var headerRow = worksheet.GetRow(worksheet.FirstRowNum)
            ?? throw new InvalidOperationException("Die Excel-Tabelle enthält keine Kopfzeile.");

        var systemColumnIndex = FindColumnIndexByHeader(headerRow, formatter, "System");
        var connectingTypeColumnIndex = FindColumnIndexByHeader(headerRow, formatter, "Connecting Type");
        var ioTagColumnIndex = FindColumnIndexByHeader(headerRow, formatter, "I/O TAG");
        var loopDescriptionColumnIndex = FindColumnIndexByHeader(headerRow, formatter, "Loop Description");
        var functionColumnIndex = FindColumnIndexByHeader(headerRow, formatter, "Function");
        var signalDescriptionColumnIndex = FindColumnIndexByHeader(headerRow, formatter, "Signal Description");

        if (systemColumnIndex < 0 || connectingTypeColumnIndex < 0 || ioTagColumnIndex < 0 || loopDescriptionColumnIndex < 0 || functionColumnIndex < 0 || signalDescriptionColumnIndex < 0)
        {
            return new List<BlowerDbEntry>();
        }

        var result = new List<BlowerDbEntry>();
        var uniqueKeys = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

        for (var rowIndex = headerRow.RowNum + 1; rowIndex <= worksheet.LastRowNum; rowIndex++)
        {
            var row = worksheet.GetRow(rowIndex);
            if (row is null)
            {
                continue;
            }

            var systemValue = formatter.FormatCellValue(row.GetCell(systemColumnIndex))?.Trim();
            if (!string.Equals(systemValue, "PLC Blower", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            analysis.PlcBlowerRowCount++;

            var ioTagCell = row.GetCell(ioTagColumnIndex);
            var isRedMarked = ioTagCell is not null && IsCellMarkedRed(ioTagCell);
            if (!isRedMarked)
            {
                analysis.NonRedSkippedCount++;
                continue;
            }

            analysis.RedMarkedPlcBlowerRowCount++;

            var connectingTypeValue = formatter.FormatCellValue(row.GetCell(connectingTypeColumnIndex))?.Trim();
            var route = ResolveBlowerRoute(connectingTypeValue);
            if (route is null)
            {
                analysis.InvalidConnectingTypeCount++;
                continue;
            }

            var ioTagValue = formatter.FormatCellValue(row.GetCell(ioTagColumnIndex))?.Trim();
            if (string.IsNullOrWhiteSpace(ioTagValue))
            {
                analysis.MissingIoTagCount++;
                continue;
            }

            var memberName = ToValidDbMemberName(ioTagValue!);
            if (string.IsNullOrWhiteSpace(memberName))
            {
                analysis.MissingIoTagCount++;
                continue;
            }

            var loopDescriptionValue = formatter.FormatCellValue(row.GetCell(loopDescriptionColumnIndex))?.Trim();
            var functionValue = formatter.FormatCellValue(row.GetCell(functionColumnIndex))?.Trim();
            var signalDescriptionValue = formatter.FormatCellValue(row.GetCell(signalDescriptionColumnIndex))?.Trim();
            var comment = BuildBlowerComment(loopDescriptionValue, functionValue, signalDescriptionValue);

            var key = route.Value.TargetDbName + "|" + memberName;
            if (!uniqueKeys.Add(key))
            {
                continue;
            }

            result.Add(new BlowerDbEntry
            {
                TargetDbName = route.Value.TargetDbName,
                MemberName = memberName,
                SourceIoTag = ioTagValue!,
                DataType = route.Value.DataType,
                Comment = comment
            });
        }

        return result;
    }

    private static (string TargetDbName, string DataType)? ResolveBlowerRoute(string? connectingTypeValue)
    {
        if (string.IsNullOrWhiteSpace(connectingTypeValue))
        {
            return null;
        }

        var normalized = (connectingTypeValue ?? string.Empty).Trim().ToUpperInvariant();
        if (normalized.StartsWith("AI", StringComparison.Ordinal))
        {
            return (PlcBlowerFromDbName, "Int");
        }

        if (normalized.StartsWith("DI", StringComparison.Ordinal))
        {
            return (PlcBlowerFromDbName, "Bool");
        }

        if (normalized.StartsWith("AO", StringComparison.Ordinal))
        {
            return (PlcBlowerToDbName, "Int");
        }

        if (normalized.StartsWith("DO", StringComparison.Ordinal))
        {
            return (PlcBlowerToDbName, "Bool");
        }

        return null;
    }

    private static string BuildBlowerComment(string? loopDescription, string? function, string? signalDescription)
    {
        return string.Join(" ", new[] { loopDescription, function, signalDescription }
            .Where(value => !string.IsNullOrWhiteSpace(value))
            .Select(value => value!.Trim()));
    }

    private static string ToValidDbMemberName(string ioTagValue)
    {
        return (ioTagValue ?? string.Empty).Trim();
    }

    private static string ToIoTagDisplayName(string memberName)
    {
        if (string.IsNullOrWhiteSpace(memberName))
        {
            return string.Empty;
        }

        return memberName.Replace('_', '-');
    }

    private static string ToLegacyNormalizedMemberName(string memberName)
    {
        if (string.IsNullOrWhiteSpace(memberName))
        {
            return string.Empty;
        }

        var sanitized = Regex.Replace(memberName.Trim(), @"[^A-Za-z0-9_]", "_", RegexOptions.CultureInvariant);
        sanitized = Regex.Replace(sanitized, @"_+", "_", RegexOptions.CultureInvariant).Trim('_');
        if (string.IsNullOrWhiteSpace(sanitized))
        {
            return string.Empty;
        }

        if (!char.IsLetter(sanitized[0]) && sanitized[0] != '_')
        {
            sanitized = "_" + sanitized;
        }

        return sanitized;
    }

    private static List<TemplateMapping> BuildTemplateMappings(
        List<WorksheetEntry> entries,
        Func<WorksheetEntry, string> keySelector,
        string keyPrefix)
    {
        return entries
            .GroupBy(keySelector, StringComparer.OrdinalIgnoreCase)
            .Where(group => !string.IsNullOrWhiteSpace(group.Key))
            .Select(group =>
            {
                var redEntries = group
                    .Where(entry => entry.IsRed)
                    .Select(entry => entry.Value)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                var sourceEntries = group
                    .Where(entry => !entry.IsRed)
                    .Select(entry => entry.Value)
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (redEntries.Count == 0 || sourceEntries.Count == 0)
                {
                    return null;
                }

                var selectedSource = sourceEntries[0];
                var targets = redEntries
                    .Where(entry => !string.Equals(entry, selectedSource, StringComparison.OrdinalIgnoreCase))
                    .Distinct(StringComparer.OrdinalIgnoreCase)
                    .ToList();

                if (targets.Count == 0)
                {
                    return null;
                }

                return new TemplateMapping
                {
                    StructureKey = keyPrefix + group.Key,
                    SourceEntry = selectedSource,
                    TargetEntries = targets
                };
            })
            .Where(mapping => mapping is not null)
            .Cast<TemplateMapping>()
            .ToList();
    }

    private static string NormalizeFunctionValue(string value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return Regex.Replace(value.Trim(), @"\s+", " ", RegexOptions.CultureInvariant);
    }

    private static int FindColumnIndexByHeader(IRow headerRow, DataFormatter formatter, string headerName)
    {
        var normalizedHeader = NormalizeHeaderName(headerName);
        var firstCell = Math.Max((int)headerRow.FirstCellNum, 0);
        var lastCell = headerRow.LastCellNum;
        for (var cellIndex = firstCell; cellIndex < lastCell; cellIndex++)
        {
            var headerCell = headerRow.GetCell(cellIndex);
            var currentHeader = formatter.FormatCellValue(headerCell);
            if (NormalizeHeaderName(currentHeader) == normalizedHeader)
            {
                return cellIndex;
            }
        }

        return -1;
    }

    private static string NormalizeHeaderName(string? value)
    {
        if (string.IsNullOrWhiteSpace(value))
        {
            return string.Empty;
        }

        return new string(value
            .Where(character => !char.IsWhiteSpace(character))
            .ToArray())
            .Trim()
            .ToUpperInvariant();
    }

    private static string BuildStructureKey(string value)
    {
        var builder = new StringBuilder(value.Length);
        foreach (var character in value)
        {
            if (char.IsLetter(character))
            {
                builder.Append('A');
            }
            else if (char.IsDigit(character))
            {
                builder.Append('#');
            }
            else if (!char.IsWhiteSpace(character))
            {
                builder.Append(character);
            }
        }

        return builder.ToString();
    }

    private static (string SourceStructure, string TargetStructure)? FindStructureReplacement(string sourceValue, string targetValue)
    {
        var sourceTokens = ExtractStructureTokens(sourceValue);
        var targetTokens = ExtractStructureTokens(targetValue);
        if (sourceTokens.Count == 0 || targetTokens.Count == 0)
        {
            return null;
        }

        var sourceTokenDetails = sourceTokens
            .Select(token => new StructureToken(token))
            .ToList();

        var targetTokenDetails = targetTokens
            .Select(token => new StructureToken(token))
            .ToList();

        foreach (var sourceToken in sourceTokenDetails)
        {
            var match = targetTokenDetails.FirstOrDefault(targetToken =>
                string.Equals(targetToken.Prefix, sourceToken.Prefix, StringComparison.OrdinalIgnoreCase)
                && targetToken.NumberPart.Length == sourceToken.NumberPart.Length
                && !string.Equals(targetToken.Raw, sourceToken.Raw, StringComparison.OrdinalIgnoreCase));

            if (match is not null)
            {
                return (sourceToken.Raw, match.Raw);
            }
        }

        return null;
    }

    private static List<string> ExtractStructureTokens(string value)
    {
        var matches = Regex.Matches(value, @"[A-Za-z]+\d+");
        return matches
            .Cast<Match>()
            .Select(match => match.Value)
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();
    }

    private static bool IsCellMarkedRed(ICell cell)
    {
        try
        {
            var style = cell.CellStyle;
            if (style is null)
            {
                return false;
            }

            if (style.FillPattern == FillPattern.NoFill)
            {
                return false;
            }

            if (style is XSSFCellStyle xssfStyle)
            {
                try
                {
                    var xssfColor = xssfStyle.FillForegroundColorColor as XSSFColor;
                    var rgb = xssfColor?.RGB;
                    if (rgb is { Length: >= 3 } && IsExactRgbRed(rgb[0], rgb[1], rgb[2]))
                    {
                        return true;
                    }
                }
                catch
                {
                }
            }

            if (style is HSSFCellStyle hssfStyle)
            {
                try
                {
                    var hssfColor = hssfStyle.FillForegroundColorColor as HSSFColor;
                    var triplet = hssfColor?.RGB;
                    if (triplet is { Length: >= 3 } && IsExactRgbRed((byte)triplet[0], (byte)triplet[1], (byte)triplet[2]))
                    {
                        return true;
                    }
                }
                catch
                {
                }
            }

            return IsIndexedRed(style.FillForegroundColor);
        }
        catch
        {
            return false;
        }
    }

    private static string DescribeCellRedColor(ICell cell)
    {
        try
        {
            var style = cell.CellStyle;
            if (style is null)
            {
                return string.Empty;
            }

            if (style is XSSFCellStyle xssfStyle)
            {
                var fillColor = xssfStyle.FillForegroundColorColor as XSSFColor;
                var fillRgb = fillColor?.RGB;
                if (fillRgb is { Length: >= 3 })
                {
                    return $"FillRGB=#{fillRgb[0]:X2}{fillRgb[1]:X2}{fillRgb[2]:X2}";
                }
            }

            if (style is HSSFCellStyle hssfStyle)
            {
                var fillColor = hssfStyle.FillForegroundColorColor as HSSFColor;
                var triplet = fillColor?.RGB;
                if (triplet is { Length: >= 3 })
                {
                    return $"FillRGB=#{((byte)triplet[0]):X2}{((byte)triplet[1]):X2}{((byte)triplet[2]):X2}";
                }
            }

            var workbook = cell.Sheet.Workbook;
            var font = workbook.GetFontAt(style.FontIndex);
            if (font is XSSFFont xssfFont)
            {
                var fontColor = xssfFont.GetXSSFColor();
                var fontRgb = fontColor?.RGB;
                if (fontRgb is { Length: >= 3 })
                {
                    return $"FontRGB=#{fontRgb[0]:X2}{fontRgb[1]:X2}{fontRgb[2]:X2}";
                }
            }

            if (style.FillForegroundColor > 0)
            {
                return $"FillIndexed={style.FillForegroundColor}";
            }

            return font is not null && font.Color > 0
                ? $"FontIndexed={font.Color}"
                : string.Empty;
        }
        catch
        {
            return string.Empty;
        }
    }

    private static bool IsIndexedRed(short colorIndex)
    {
        return colorIndex == IndexedColors.Red.Index
            || colorIndex == IndexedColors.DarkRed.Index
            || colorIndex == IndexedColors.Rose.Index
            || colorIndex == IndexedColors.Coral.Index;
    }

    private static bool IsExactRgbRed(byte red, byte green, byte blue)
    {
        return red >= 250 && red <= 255
            && green >= 0 && green <= 50
            && blue >= 0 && blue <= 50;
    }

    private static void CloneStructureByToken(
        PlcBlockGroup rootGroup,
        string sourceToken,
        string targetToken,
        string mappingKey,
        HashSet<string> allowedTargetIoTags,
        List<string> buildLogEntries,
        ref int createdFolderCount,
        ref int createdBlockCount,
        ref int skippedItemCount,
        ref int errorCount)
    {
        if (string.Equals(sourceToken, targetToken, StringComparison.OrdinalIgnoreCase))
        {
            skippedItemCount++;
            return;
        }

        var sourceGroups = EnumerateGroups(rootGroup)
            .Where(groupInfo => ContainsInvariant(groupInfo.Path, sourceToken))
            .OrderBy(groupInfo => groupInfo.Path.Count(character => character == '/'))
            .ToList();

        foreach (var sourceGroup in sourceGroups)
        {
            var targetGroupPath = ReplaceToken(sourceGroup.Path, sourceToken, targetToken);
            if (string.Equals(sourceGroup.Path, targetGroupPath, StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            EnsureGroupPath(rootGroup, targetGroupPath, buildLogEntries, ref createdFolderCount);
        }

        var sourceBlocks = EnumerateBlocks(rootGroup)
            .Where(blockInfo => IsCloneableBlock(blockInfo.Block)
                && (ContainsInvariant(blockInfo.Block.Name, sourceToken) || ContainsInvariant(blockInfo.GroupPath, sourceToken)))
            .OrderBy(blockInfo => blockInfo.Block is InstanceDB ? 1 : 0)
            .ThenBy(blockInfo => blockInfo.Block.Name, StringComparer.OrdinalIgnoreCase)
            .ToList();

        foreach (var sourceBlock in sourceBlocks)
        {
            try
            {
                var targetGroupPath = ReplaceToken(sourceBlock.GroupPath, sourceToken, targetToken);
                var targetName = ReplaceToken(sourceBlock.Block.Name, sourceToken, targetToken);
                if (string.Equals(sourceBlock.Block.Name, targetName, StringComparison.OrdinalIgnoreCase))
                {
                    skippedItemCount++;
                    continue;
                }

                if (!IsTargetNameAllowed(targetName, sourceBlock.Block, allowedTargetIoTags))
                {
                    skippedItemCount++;
                    buildLogEntries.Add($"Übersprungen (nicht in Excel-Whitelist): {targetName}");
                    continue;
                }

                var targetGroup = EnsureGroupPath(rootGroup, targetGroupPath, buildLogEntries, ref createdFolderCount);
                if (FindBlockByName(rootGroup, targetName) is not null)
                {
                    skippedItemCount++;
                    buildLogEntries.Add($"Übersprungen (existiert bereits): {targetGroupPath}/{targetName}");
                    continue;
                }

                if (sourceBlock.Block is InstanceDB sourceInstanceDb)
                {
                    var sourceInstanceOfName = sourceInstanceDb.InstanceOfName;
                    if (!string.IsNullOrWhiteSpace(sourceInstanceOfName))
                    {
                        var targetInstanceOfName = ReplaceToken(sourceInstanceOfName, sourceToken, targetToken);
                        if (FindBlockByName(rootGroup, targetInstanceOfName) is null)
                        {
                            skippedItemCount++;
                            buildLogEntries.Add($"Übersprungen (abhängiger FB fehlt): InstanceDB={targetName}, InstanceOf={targetInstanceOfName}");
                            continue;
                        }
                    }
                }

                CopyAndRenameBlock(rootGroup, sourceBlock.Block, targetGroup, targetName, sourceToken, targetToken);
                createdBlockCount++;
                buildLogEntries.Add($"Block erstellt: {targetGroupPath}/{targetName} (Quelle: {sourceBlock.GroupPath}/{sourceBlock.Block.Name})");

                if (sourceBlock.Block is InstanceDB createdInstanceDb)
                {
                    var resolvedDbNumber = ResolveValidDbNumber(TryGetInstanceDbNumber(createdInstanceDb), targetName);
                    var targetInstanceOfName = ReplaceToken(createdInstanceDb.InstanceOfName ?? string.Empty, sourceToken, targetToken);
                    buildLogEntries.Add($"InstanceDB erstellt: Name={targetName}, Nummer={resolvedDbNumber}, InstanceOf={targetInstanceOfName}");
                }
            }
            catch (Exception ex)
            {
                if (IsInconsistentBlockExportError(ex))
                {
                    skippedItemCount++;
                    buildLogEntries.Add($"Vorlagenfehler: [{mappingKey}] {sourceBlock.GroupPath}/{sourceBlock.Block.Name} | Inconsistent blocks and PLC data types (UDT)");
                    continue;
                }

                errorCount++;
                buildLogEntries.Add($"Fehler beim Kopieren: Quelle={sourceBlock.GroupPath}/{sourceBlock.Block.Name} | Fehler={CompactLogMessage(ex.Message)}");
            }
        }
    }

    private static PlcBlock? FindBlockByName(PlcBlockGroup rootGroup, string blockName)
    {
        return EnumerateBlocks(rootGroup)
            .Select(info => info.Block)
            .FirstOrDefault(block => string.Equals(block.Name, blockName, StringComparison.OrdinalIgnoreCase));
    }

    private static bool IsInconsistentBlockExportError(Exception exception)
    {
        return exception.Message.IndexOf("Inconsistent blocks and PLC data types", StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static bool IsTargetNameAllowed(string targetName, PlcBlock sourceBlock, HashSet<string> allowedTargetIoTags)
    {
        if (allowedTargetIoTags.Contains(targetName))
        {
            return true;
        }

        var allowedBases = BuildAllowedIoTagBases(allowedTargetIoTags);
        if (allowedBases.Any(baseName => MatchesIoTagBase(targetName, baseName)))
        {
            return true;
        }

        if (sourceBlock is DataBlock && targetName.EndsWith("_DB", StringComparison.OrdinalIgnoreCase))
        {
            var baseName = targetName.Substring(0, targetName.Length - 3);
            if (allowedTargetIoTags.Contains(baseName))
            {
                return true;
            }

            return allowedBases.Any(allowedBase => MatchesIoTagBase(baseName, allowedBase));
        }

        return false;
    }

    private static int ExtendPlcBlowerSignalExchangeDbs(
        PlcBlockGroup rootGroup,
        List<BlowerDbEntry> entries,
        List<string> buildLogEntries,
        ref int skippedItemCount,
        ref int errorCount)
    {
        if (entries.Count == 0)
        {
            return 0;
        }

        var totalAdded = 0;
        var blocksByName = EnumerateBlocks(rootGroup)
            .Where(info => info.Block is DataBlock)
            .ToDictionary(info => info.Block.Name, info => info, StringComparer.OrdinalIgnoreCase);

        foreach (var dbGroup in entries.GroupBy(entry => entry.TargetDbName, StringComparer.OrdinalIgnoreCase))
        {
            if (!blocksByName.TryGetValue(dbGroup.Key, out var targetDbInfo))
            {
                errorCount++;
                buildLogEntries.Add($"Fehler PLC-Blower DB: Datenbaustein nicht gefunden '{dbGroup.Key}'.");
                continue;
            }

            var dataBlock = targetDbInfo.Block as DataBlock;
            if (dataBlock is null)
            {
                errorCount++;
                buildLogEntries.Add($"Fehler PLC-Blower DB: Block ist kein DataBlock '{dbGroup.Key}'.");
                continue;
            }

            if (!TryProbeDataBlockExport(dataBlock, out var exportProbeError))
            {
                errorCount++;
                buildLogEntries.Add($"PLC-Blower Diagnose: DB '{dbGroup.Key}' ist nicht exportierbar: {exportProbeError}");
                if (exportProbeError.IndexOf("Inconsistent blocks and PLC data types", StringComparison.OrdinalIgnoreCase) >= 0)
                {
                    buildLogEntries.Add($"PLC-Blower Diagnose: DB '{dbGroup.Key}' enthält inkonsistente UDT-Typen. Bitte DB/UDT im TIA übersetzen und Konsistenz prüfen.");
                }

                continue;
            }

            buildLogEntries.Add($"PLC-Blower Diagnose: DB '{dbGroup.Key}' ist exportierbar.");

            try
            {
                var addedCount = AddMembersToDataBlockViaXml(
                    rootGroup,
                    targetDbInfo.Group,
                    dataBlock,
                    dbGroup.ToList(),
                    buildLogEntries,
                    ref skippedItemCount);

                totalAdded += addedCount;
                buildLogEntries.Add($"PLC-Blower DB erweitert: {dbGroup.Key} (+{addedCount})");
            }
            catch (Exception ex)
            {
                errorCount++;
                buildLogEntries.Add($"Fehler PLC-Blower DB: {dbGroup.Key} | Fehler={CompactLogMessage(ex.Message)}");
            }
        }

        return totalAdded;
    }

    private static bool TryProbeDataBlockExport(DataBlock dataBlock, out string errorMessage)
    {
        var tempExportPath = Path.Combine(Path.GetTempPath(), $"TiaBlowerDbProbe_{Guid.NewGuid():N}.xml");
        try
        {
            dataBlock.Export(new FileInfo(tempExportPath), ExportOptions.WithDefaults);
            errorMessage = string.Empty;
            return true;
        }
        catch (Exception ex)
        {
            errorMessage = BuildExceptionSummary(ex);
            return false;
        }
        finally
        {
            if (File.Exists(tempExportPath))
            {
                File.Delete(tempExportPath);
            }
        }
    }

    private static int AddMembersToDataBlockViaXml(
        PlcBlockGroup rootGroup,
        PlcBlockGroup ownerGroup,
        DataBlock dataBlock,
        List<BlowerDbEntry> entries,
        List<string> buildLogEntries,
        ref int skippedItemCount)
    {
        var originalExportPath = Path.Combine(Path.GetTempPath(), $"TiaBlowerDbOriginal_{Guid.NewGuid():N}.xml");
        var modifiedExportPath = Path.Combine(Path.GetTempPath(), $"TiaBlowerDbModified_{Guid.NewGuid():N}.xml");
        var blockName = dataBlock.Name;

        try
        {
            dataBlock.Export(new FileInfo(originalExportPath), ExportOptions.WithDefaults);
            var document = XDocument.Load(originalExportPath, LoadOptions.PreserveWhitespace);

            var staticSection = document
                .Descendants()
                .FirstOrDefault(element => IsElementName(element, "Section")
                    && string.Equals((string?)element.Attribute("Name"), "Static", StringComparison.OrdinalIgnoreCase));

            staticSection ??= document.Descendants().FirstOrDefault(element => IsElementName(element, "Section"));
            if (staticSection is null)
            {
                throw new InvalidOperationException($"Keine DB-Sektion gefunden in '{blockName}'.");
            }

            var existingMembers = staticSection.Elements()
                .Where(element => IsElementName(element, "Member"))
                .ToList();

            var binStructMember = existingMembers.FirstOrDefault(member =>
                string.Equals((string?)member.Attribute("Name"), "BIN", StringComparison.OrdinalIgnoreCase)
                && string.Equals(GetMemberDataType(member), "Struct", StringComparison.OrdinalIgnoreCase));

            var anaStructMember = existingMembers.FirstOrDefault(member =>
                string.Equals((string?)member.Attribute("Name"), "ANA", StringComparison.OrdinalIgnoreCase)
                && string.Equals(GetMemberDataType(member), "Struct", StringComparison.OrdinalIgnoreCase));

            var templateMember = existingMembers.FirstOrDefault();
            if (templateMember is null)
            {
                throw new InvalidOperationException($"Keine Member-Vorlage in DB '{blockName}' gefunden.");
            }

            var scalarTemplateMembers = staticSection
                .Descendants()
                .Where(element => IsElementName(element, "Member"))
                .Where(element => !element.Elements().Any(child => IsElementName(child, "Member")))
                .ToList();

            var scalarTemplateFallback = scalarTemplateMembers.FirstOrDefault();

            var templateByDataType = scalarTemplateMembers
                .GroupBy(member => GetMemberDataType(member), StringComparer.OrdinalIgnoreCase)
                .Where(group => !string.IsNullOrWhiteSpace(group.Key))
                .ToDictionary(group => group.Key!, group => group.First(), StringComparer.OrdinalIgnoreCase);

            var existingNames = new HashSet<string>(
                staticSection
                    .Descendants()
                    .Where(element => IsElementName(element, "Member"))
                    .Select(member => (string?)member.Attribute("Name"))
                    .Where(name => !string.IsNullOrWhiteSpace(name))
                    .Cast<string>(),
                StringComparer.OrdinalIgnoreCase);

            var added = 0;
            var renamed = 0;
            foreach (var entry in entries)
            {
                var ioTagForLog = string.IsNullOrWhiteSpace(entry.SourceIoTag)
                    ? ToIoTagDisplayName(entry.MemberName)
                    : entry.SourceIoTag;

                var legacyMemberName = ToLegacyNormalizedMemberName(entry.MemberName);
                if (!string.Equals(legacyMemberName, entry.MemberName, StringComparison.OrdinalIgnoreCase)
                    && !existingNames.Contains(entry.MemberName)
                    && existingNames.Contains(legacyMemberName))
                {
                    var legacyMember = staticSection
                        .Descendants()
                        .FirstOrDefault(element => IsElementName(element, "Member")
                            && string.Equals((string?)element.Attribute("Name"), legacyMemberName, StringComparison.OrdinalIgnoreCase));

                    if (legacyMember is not null)
                    {
                        legacyMember.SetAttributeValue("Name", entry.MemberName);
                        existingNames.Remove(legacyMemberName);
                        existingNames.Add(entry.MemberName);
                        renamed++;
                        buildLogEntries.Add($"PLC-Blower DB-Eintrag migriert: {blockName}/{legacyMemberName} -> {entry.MemberName} (I/O TAG: {ioTagForLog})");
                        continue;
                    }
                }

                if (existingNames.Contains(entry.MemberName))
                {
                    skippedItemCount++;
                    buildLogEntries.Add($"PLC-Blower Vorabprüfung: Eintrag bereits vorhanden -> {blockName}/{entry.MemberName} (I/O TAG: {ioTagForLog})");
                    continue;
                }

                var templateForEntry = templateByDataType.TryGetValue(entry.DataType, out var typedTemplate)
                    ? typedTemplate
                    : scalarTemplateFallback ?? templateMember;

                var targetParent = ResolveTargetParentForBlowerEntry(staticSection, binStructMember, anaStructMember, entry.DataType);

                var parentScalarTemplateMembers = targetParent
                    .Descendants()
                    .Where(element => IsElementName(element, "Member"))
                    .Where(element => !element.Elements().Any(child => IsElementName(child, "Member")))
                    .ToList();

                if (parentScalarTemplateMembers.Count > 0)
                {
                    var parentTemplateByType = parentScalarTemplateMembers
                        .GroupBy(member => GetMemberDataType(member), StringComparer.OrdinalIgnoreCase)
                        .Where(group => !string.IsNullOrWhiteSpace(group.Key))
                        .ToDictionary(group => group.Key!, group => group.First(), StringComparer.OrdinalIgnoreCase);

                    if (parentTemplateByType.TryGetValue(entry.DataType, out var parentTypedTemplate))
                    {
                        templateForEntry = parentTypedTemplate;
                    }
                    else
                    {
                        templateForEntry = parentScalarTemplateMembers.First();
                    }
                }

                var newMember = new XElement(templateForEntry);
                RemoveInternalXmlIds(newMember);
                RemoveNestedMemberElements(newMember);
                newMember.SetAttributeValue("Name", entry.MemberName);
                SetMemberDataType(newMember, entry.DataType);
                SetMemberComment(newMember, entry.Comment);
                targetParent.Add(newMember);
                existingNames.Add(entry.MemberName);
                added++;
                buildLogEntries.Add($"PLC-Blower DB-Eintrag: {blockName}/{entry.MemberName} (I/O TAG: {ioTagForLog}) [{entry.DataType}] Kommentar='{entry.Comment}'");
            }

            if (added == 0 && renamed == 0)
            {
                return 0;
            }

            document.Save(modifiedExportPath);

            try
            {
                ownerGroup.Blocks.Import(new FileInfo(modifiedExportPath), ImportOptions.Override);
            }
            catch (Exception ex)
            {
                throw new InvalidOperationException($"Import der aktualisierten DB '{blockName}' fehlgeschlagen.", ex);
            }

            if (FindBlockByName(rootGroup, blockName) is null)
            {
                throw new InvalidOperationException($"Import der aktualisierten DB '{blockName}' fehlgeschlagen.");
            }

            return added + renamed;
        }
        finally
        {
            if (File.Exists(originalExportPath))
            {
                File.Delete(originalExportPath);
            }

            if (File.Exists(modifiedExportPath))
            {
                File.Delete(modifiedExportPath);
            }
        }
    }

    private static bool IsElementName(XElement element, string localName)
    {
        return string.Equals(element.Name.LocalName, localName, StringComparison.OrdinalIgnoreCase);
    }

    private static void RemoveInternalXmlIds(XElement element)
    {
        foreach (var node in element.DescendantsAndSelf())
        {
            var attributesToRemove = node.Attributes()
                .Where(attribute => string.Equals(attribute.Name.LocalName, "ID", StringComparison.OrdinalIgnoreCase)
                    || string.Equals(attribute.Name.LocalName, "UId", StringComparison.OrdinalIgnoreCase))
                .ToList();

            foreach (var attribute in attributesToRemove)
            {
                attribute.Remove();
            }
        }
    }

    private static void RemoveNestedMemberElements(XElement element)
    {
        var nestedMembers = element
            .Descendants()
            .Where(node => IsElementName(node, "Member"))
            .ToList();

        foreach (var nestedMember in nestedMembers)
        {
            nestedMember.Remove();
        }
    }

    private static XElement ResolveTargetParentForBlowerEntry(
        XElement staticSection,
        XElement? binStructMember,
        XElement? anaStructMember,
        string dataType)
    {
        if (string.Equals(dataType, "Bool", StringComparison.OrdinalIgnoreCase) && binStructMember is not null)
        {
            return binStructMember;
        }

        if (string.Equals(dataType, "Int", StringComparison.OrdinalIgnoreCase) && anaStructMember is not null)
        {
            return anaStructMember;
        }

        return staticSection;
    }

    private static string? GetMemberDataType(XElement member)
    {
        var dataTypeAttribute = ((string?)member.Attribute("Datatype") ?? (string?)member.Attribute("DataType") ?? string.Empty).Trim();
        if (!string.IsNullOrWhiteSpace(dataTypeAttribute))
        {
            return dataTypeAttribute;
        }

        var dataTypeNode = member.Descendants()
            .FirstOrDefault(node => IsElementName(node, "Datatype") || IsElementName(node, "DataType"));

        var dataTypeText = (dataTypeNode?.Value ?? string.Empty).Trim();
        if (string.IsNullOrWhiteSpace(dataTypeText))
        {
            return null;
        }

        return dataTypeText;
    }

    private static void SetMemberDataType(XElement member, string dataType)
    {
        if (member.Attribute("Datatype") is not null)
        {
            member.SetAttributeValue("Datatype", dataType);
        }

        if (member.Attribute("DataType") is not null)
        {
            member.SetAttributeValue("DataType", dataType);
        }

        foreach (var typeNode in member.Descendants().Where(node => IsElementName(node, "Datatype") || IsElementName(node, "DataType")))
        {
            typeNode.Value = dataType;
        }
    }

    private static void SetMemberComment(XElement member, string comment)
    {
        if (string.IsNullOrWhiteSpace(comment))
        {
            return;
        }

        var commentNodes = member.Descendants().Where(node => IsElementName(node, "Comment")).ToList();
        if (commentNodes.Count == 0)
        {
            return;
        }

        foreach (var commentNode in commentNodes)
        {
            var textNodes = commentNode.Descendants()
                .Where(node => !node.HasElements
                    && (IsElementName(node, "Text") || IsElementName(node, "Value") || IsElementName(node, "MultiLanguageText")))
                .ToList();

            if (textNodes.Count == 0)
            {
                if (!commentNode.HasElements)
                {
                    commentNode.Value = comment;
                }

                continue;
            }

            foreach (var textNode in textNodes)
            {
                textNode.Value = comment;
            }
        }
    }

    private static HashSet<string> BuildAllowedIoTagBases(HashSet<string> allowedTargetIoTags)
    {
        var bases = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
        foreach (var ioTag in allowedTargetIoTags)
        {
            var parts = ioTag.Split(new[] { '-' }, StringSplitOptions.RemoveEmptyEntries);
            if (parts.Length >= 2)
            {
                bases.Add(string.Join("-", parts.Take(parts.Length - 1)));
            }
        }

        return bases;
    }

    private static bool MatchesIoTagBase(string targetName, string ioTagBase)
    {
        if (string.Equals(targetName, ioTagBase, StringComparison.OrdinalIgnoreCase))
        {
            return true;
        }

        return targetName.StartsWith(ioTagBase + "-", StringComparison.OrdinalIgnoreCase)
            || targetName.StartsWith(ioTagBase + "_", StringComparison.OrdinalIgnoreCase);
    }

    private static string CompactLogMessage(string message)
    {
        var singleLine = Regex.Replace(message, @"\s+", " ").Trim();
        const int maxLength = 300;
        if (singleLine.Length <= maxLength)
        {
            return singleLine;
        }

        return singleLine.Substring(0, maxLength) + "...";
    }

    private static string BuildExceptionSummary(Exception exception)
    {
        var details = new List<string>();
        Exception? current = exception;
        var depth = 0;
        while (current is not null && depth < 6)
        {
            var message = string.IsNullOrWhiteSpace(current.Message)
                ? "(keine Meldung)"
                : CompactLogMessage(current.Message);

            details.Add($"{current.GetType().Name}: {message}");
            current = current.InnerException;
            depth++;
        }

        return string.Join(" | ", details);
    }

    private static List<GroupInfo> EnumerateGroups(PlcBlockGroup rootGroup)
    {
        var result = new List<GroupInfo>();

        void Traverse(PlcBlockGroup currentGroup, string currentPath)
        {
            foreach (PlcBlockGroup childGroup in currentGroup.Groups)
            {
                var childPath = string.IsNullOrWhiteSpace(currentPath)
                    ? childGroup.Name
                    : $"{currentPath}/{childGroup.Name}";

                result.Add(new GroupInfo { Group = childGroup, Path = childPath });
                Traverse(childGroup, childPath);
            }
        }

        Traverse(rootGroup, string.Empty);
        return result;
    }

    private static List<BlockInfo> EnumerateBlocks(PlcBlockGroup rootGroup)
    {
        var result = new List<BlockInfo>();

        void Traverse(PlcBlockGroup currentGroup, string currentPath)
        {
            foreach (PlcBlock block in currentGroup.Blocks)
            {
                result.Add(new BlockInfo
                {
                    Block = block,
                    Group = currentGroup,
                    GroupPath = currentPath
                });
            }

            foreach (PlcBlockGroup childGroup in currentGroup.Groups)
            {
                var childPath = string.IsNullOrWhiteSpace(currentPath)
                    ? childGroup.Name
                    : $"{currentPath}/{childGroup.Name}";
                Traverse(childGroup, childPath);
            }
        }

        Traverse(rootGroup, string.Empty);
        return result;
    }

    private static bool IsCloneableBlock(PlcBlock block)
    {
        return block is FB || block is DataBlock;
    }

    private static PlcBlockGroup EnsureGroupPath(PlcBlockGroup rootGroup, string groupPath, List<string> buildLogEntries, ref int createdFolderCount)
    {
        var current = rootGroup;
        if (string.IsNullOrWhiteSpace(groupPath))
        {
            return current;
        }

        var segments = groupPath.Split(new[] { '/' }, StringSplitOptions.RemoveEmptyEntries);
        var currentPath = new StringBuilder();
        foreach (var segment in segments)
        {
            if (currentPath.Length > 0)
            {
                currentPath.Append('/');
            }

            currentPath.Append(segment);

            var next = current.Groups.Find(segment);
            if (next is null)
            {
                next = current.Groups.Create(segment);
                createdFolderCount++;
                buildLogEntries.Add($"Ordner erstellt: {currentPath}");
            }

            current = next;
        }

        return current;
    }

    private static void CopyAndRenameBlock(PlcBlockGroup rootGroup, PlcBlock sourceBlock, PlcBlockGroup targetGroup, string targetName, string sourceToken, string targetToken)
    {
        if (sourceBlock is InstanceDB sourceInstanceDb)
        {
            CreateInstanceDbFromTemplate(rootGroup, sourceInstanceDb, targetGroup, targetName, sourceToken, targetToken);
            return;
        }

        var tempFile = Path.Combine(Path.GetTempPath(), $"TiaBlockClone_{Guid.NewGuid():N}.xml");
        try
        {
            sourceBlock.Export(new FileInfo(tempFile), ExportOptions.WithDefaults);
            RewriteExportedBlockXmlToken(tempFile, sourceToken, targetToken);
            var importedBlocks = targetGroup.Blocks.Import(new FileInfo(tempFile), ImportOptions.None);
            var importedBlock = importedBlocks.FirstOrDefault()
                ?? throw new InvalidOperationException($"Import fehlgeschlagen für Block '{sourceBlock.Name}'.");

            if (!string.Equals(importedBlock.Name, targetName, StringComparison.OrdinalIgnoreCase))
            {
                importedBlock.Name = targetName;
            }
        }
        finally
        {
            if (File.Exists(tempFile))
            {
                File.Delete(tempFile);
            }
        }
    }

    private static void CreateInstanceDbFromTemplate(PlcBlockGroup rootGroup, InstanceDB sourceInstanceDb, PlcBlockGroup targetGroup, string targetName, string sourceToken, string targetToken)
    {
        var instanceOfName = sourceInstanceDb.InstanceOfName;
        if (string.IsNullOrWhiteSpace(instanceOfName))
        {
            throw new InvalidOperationException($"Instance-DB '{sourceInstanceDb.Name}' hat keinen gültigen InstanceOfName.");
        }

        var targetInstanceOfName = ReplaceToken(instanceOfName, sourceToken, targetToken);
        if (string.IsNullOrWhiteSpace(targetInstanceOfName))
        {
            throw new InvalidOperationException($"Ziel-InstanceOfName konnte nicht ermittelt werden für '{sourceInstanceDb.Name}'.");
        }

        if (FindBlockByName(rootGroup, targetInstanceOfName) is null)
        {
            throw new InvalidOperationException($"Block '{targetInstanceOfName}' existiert nicht und kann nicht als InstanceOf verwendet werden.");
        }

        var resolvedDbNumber = ResolveValidDbNumber(TryGetInstanceDbNumber(sourceInstanceDb), targetName);
        var autoNumber = TryGetInstanceDbAutoNumber(sourceInstanceDb);

        try
        {
            targetGroup.Blocks.CreateInstanceDB(targetName, autoNumber, resolvedDbNumber, targetInstanceOfName);
        }
        catch
        {
            targetGroup.Blocks.CreateInstanceDB(targetName, true, resolvedDbNumber, targetInstanceOfName);
        }
    }

    private static int TryGetInstanceDbNumber(InstanceDB sourceInstanceDb)
    {
        try
        {
            return sourceInstanceDb.Number;
        }
        catch
        {
            return 0;
        }
    }

    private static bool TryGetInstanceDbAutoNumber(InstanceDB sourceInstanceDb)
    {
        try
        {
            return sourceInstanceDb.AutoNumber;
        }
        catch
        {
            return true;
        }
    }

    private static int ResolveValidDbNumber(int sourceNumber, string targetName)
    {
        if (sourceNumber > 0)
        {
            return sourceNumber;
        }

        var match = Regex.Match(targetName, @"DB[_-]?(\d+)$", RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
        if (match.Success && int.TryParse(match.Groups[1].Value, out var parsedNumber) && parsedNumber > 0)
        {
            return parsedNumber;
        }

        return 1;
    }

    private static void RewriteExportedBlockXmlToken(string filePath, string sourceToken, string targetToken)
    {
        if (string.IsNullOrWhiteSpace(sourceToken) || string.IsNullOrWhiteSpace(targetToken))
        {
            return;
        }

        var originalContent = File.ReadAllText(filePath, Encoding.UTF8);
        var rewrittenContent = ReplaceToken(originalContent, sourceToken, targetToken);
        if (!string.Equals(originalContent, rewrittenContent, StringComparison.Ordinal))
        {
            File.WriteAllText(filePath, rewrittenContent, Encoding.UTF8);
        }
    }

    private static bool ContainsInvariant(string value, string token)
    {
        return value.IndexOf(token, StringComparison.OrdinalIgnoreCase) >= 0;
    }

    private static string ReplaceToken(string value, string sourceToken, string targetToken)
    {
        return Regex.Replace(value, Regex.Escape(sourceToken), targetToken, RegexOptions.IgnoreCase | RegexOptions.CultureInvariant);
    }

    private static string ResolveBuildLogPath(string resolvedProjectPath, string? plcName)
    {
        var directoryPath = Path.GetDirectoryName(resolvedProjectPath) ?? ".";
        var suffix = string.IsNullOrWhiteSpace(plcName) ? string.Empty : $"_{SanitizeFileNamePart(plcName ?? "PLC")}";
        return Path.Combine(directoryPath, $"BuildLog{suffix}.txt");
    }

    private static void WriteBuildLog(string buildLogPath, string worksheetName, List<TemplateMapping> mappings, List<string> buildLogEntries)
    {
        var structureEntries = buildLogEntries
            .Where(entry => entry.StartsWith("Strukturpaar verarbeitet:", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .OrderBy(entry => entry, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var folderEntries = buildLogEntries
            .Where(entry => entry.StartsWith("Ordner erstellt:", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var blockEntries = buildLogEntries
            .Where(entry => entry.StartsWith("Block erstellt:", StringComparison.OrdinalIgnoreCase))
            .ToList();

        var skippedEntries = buildLogEntries
            .Where(entry => entry.StartsWith("Übersprungen", StringComparison.OrdinalIgnoreCase))
            .ToList();

        var whitelistSkippedEntries = buildLogEntries
            .Where(entry => entry.StartsWith("Übersprungen (nicht in Excel-Whitelist):", StringComparison.OrdinalIgnoreCase))
            .ToList();

        var blowerDiagnoseEntries = buildLogEntries
            .Where(entry => entry.StartsWith("PLC-Blower Diagnose:", StringComparison.OrdinalIgnoreCase))
            .ToList();

        var blowerErrorEntries = buildLogEntries
            .Where(entry => entry.StartsWith("Fehler PLC-Blower DB:", StringComparison.OrdinalIgnoreCase))
            .ToList();

        var errorEntries = buildLogEntries
            .Where(entry => entry.StartsWith("Fehler beim Kopieren:", StringComparison.OrdinalIgnoreCase)
                || entry.StartsWith("Fehler PLC-Blower DB:", StringComparison.OrdinalIgnoreCase)
                || entry.StartsWith("Build-Fehler:", StringComparison.OrdinalIgnoreCase))
            .ToList();

        var templateErrorEntries = buildLogEntries
            .Where(entry => entry.StartsWith("Vorlagenfehler:", StringComparison.OrdinalIgnoreCase))
            .Distinct(StringComparer.OrdinalIgnoreCase)
            .ToList();

        var templateErrorsByFunction = templateErrorEntries
            .GroupBy(GetTemplateErrorFunctionKey, StringComparer.OrdinalIgnoreCase)
            .OrderBy(group => group.Key, StringComparer.OrdinalIgnoreCase)
            .ToList();

        var infoEntries = buildLogEntries
            .Where(entry => !entry.StartsWith("Strukturpaar verarbeitet:", StringComparison.OrdinalIgnoreCase)
                && !entry.StartsWith("Ordner erstellt:", StringComparison.OrdinalIgnoreCase)
                && !entry.StartsWith("Block erstellt:", StringComparison.OrdinalIgnoreCase)
                && !entry.StartsWith("Übersprungen", StringComparison.OrdinalIgnoreCase)
                && !entry.StartsWith("PLC-Blower Diagnose:", StringComparison.OrdinalIgnoreCase)
                && !entry.StartsWith("Vorlagenfehler:", StringComparison.OrdinalIgnoreCase)
                && !entry.StartsWith("Fehler beim Kopieren:", StringComparison.OrdinalIgnoreCase))
            .ToList();

        var lines = new List<string>
        {
            "TIA Program Block Exporter - BuildLog",
            $"Erstellt am: {DateTime.Now:yyyy-MM-dd HH:mm:ss}",
            $"Excel-Tabelle: {worksheetName}",
            string.Empty,
            "Statistik:",
            $"- Strukturpaare verarbeitet: {structureEntries.Count}",
            $"- Ordner erstellt: {folderEntries.Count}",
            $"- Blöcke erstellt: {blockEntries.Count}",
            $"- Übersprungen: {skippedEntries.Count}",
            $"- Übersprungen (Excel-Whitelist): {whitelistSkippedEntries.Count}",
            $"- Vorlagenfehler: {templateErrorEntries.Count}",
            $"- Fehler: {errorEntries.Count}",
            string.Empty,
            "Erkannte Struktur-Mappings:"
        };

        foreach (var mapping in mappings)
        {
            lines.Add($"- Struktur {mapping.StructureKey}: Quelle '{mapping.SourceEntry}' -> Ziele [{string.Join(", ", mapping.TargetEntries)}]");
        }

        lines.Add(string.Empty);
        lines.Add("Verarbeitete Strukturpaare:");
        if (structureEntries.Count == 0)
        {
            lines.Add("Keine Strukturpaare verarbeitet.");
        }
        else
        {
            lines.AddRange(structureEntries);
        }

        lines.Add(string.Empty);
        lines.Add("Info / Diagnose:");
        if (infoEntries.Count == 0)
        {
            lines.Add("Keine zusätzlichen Infos.");
        }
        else
        {
            lines.AddRange(infoEntries.Select(entry => $"- {entry}"));
        }

        lines.Add(string.Empty);
        lines.Add("PLC-Blower Diagnose:");
        if (blowerDiagnoseEntries.Count == 0)
        {
            lines.Add("- Keine");
        }
        else
        {
            lines.AddRange(blowerDiagnoseEntries.Select(entry => $"- {entry}"));
        }

        lines.Add(string.Empty);
        lines.Add("Details (gekürzt):");

        lines.Add("Vorlagenfehler nach Function:");
        if (templateErrorsByFunction.Count == 0)
        {
            lines.Add("- Keine");
        }
        else
        {
            foreach (var functionGroup in templateErrorsByFunction)
            {
                lines.Add($"- {functionGroup.Key}: {functionGroup.Count()} Fehler");
            }
        }

        lines.Add(string.Empty);

        AppendCappedSection(lines, "Ordner", folderEntries, 30);
        AppendCappedSection(lines, "Blöcke", blockEntries, 50);
        AppendCappedSection(lines, "Übersprungen", skippedEntries, 50);
        AppendCappedSection(lines, "Übersprungen (Excel-Whitelist)", whitelistSkippedEntries, 50);
        AppendCappedSection(lines, "PLC-Blower Diagnose", blowerDiagnoseEntries, 50);
        AppendCappedSection(lines, "PLC-Blower Fehler", blowerErrorEntries, 50);
        AppendCappedSection(lines, "Vorlagenfehler", templateErrorEntries, 50);
        AppendCappedSection(lines, "Fehler", errorEntries, 50);

        File.WriteAllLines(buildLogPath, lines);
    }

    private static void AppendCappedSection(List<string> lines, string sectionName, List<string> entries, int maxEntries)
    {
        lines.Add($"{sectionName}: {entries.Count}");
        if (entries.Count == 0)
        {
            lines.Add("- Keine");
            lines.Add(string.Empty);
            return;
        }

        foreach (var entry in entries.Take(maxEntries))
        {
            lines.Add($"- {entry}");
        }

        if (entries.Count > maxEntries)
        {
            lines.Add($"- ... weitere {entries.Count - maxEntries} Einträge");
        }

        lines.Add(string.Empty);
    }

    private static string GetTemplateErrorFunctionKey(string entry)
    {
        var startIndex = entry.IndexOf("[", StringComparison.Ordinal);
        var endIndex = entry.IndexOf("]", StringComparison.Ordinal);
        if (startIndex >= 0 && endIndex > startIndex)
        {
            var key = entry.Substring(startIndex + 1, endIndex - startIndex - 1);
            if (key.StartsWith("Function=", StringComparison.OrdinalIgnoreCase))
            {
                return key.Substring("Function=".Length);
            }

            if (key.StartsWith("Pattern=", StringComparison.OrdinalIgnoreCase))
            {
                return $"Pattern:{key.Substring("Pattern=".Length)}";
            }

            return key;
        }

        return "Unbekannt";
    }

    private static string ResolveOutputPath(string resolvedProjectPath, string? outputPath, string? plcName)
    {
        var basePath = !string.IsNullOrWhiteSpace(outputPath)
            ? Path.GetFullPath(outputPath)
            : Path.Combine(Path.GetDirectoryName(resolvedProjectPath)!, "ProgramBlocksExport.xml");

        if (string.IsNullOrWhiteSpace(plcName))
        {
            return basePath;
        }

        var sanitizedPlcName = SanitizeFileNamePart(plcName ?? "PLC");
        var directoryPath = Path.GetDirectoryName(basePath) ?? ".";
        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(basePath);
        var extension = Path.GetExtension(basePath);
        return Path.Combine(directoryPath, $"{fileNameWithoutExtension}_{sanitizedPlcName}{extension}");
    }

    private static string ResolveErrorLogPath(string resolvedOutputPath, string? plcName)
    {
        var directoryPath = Path.GetDirectoryName(resolvedOutputPath) ?? ".";
        var suffix = string.IsNullOrWhiteSpace(plcName) ? string.Empty : $"_{SanitizeFileNamePart(plcName ?? "PLC")}";
        return Path.Combine(directoryPath, $"ErrorLog{suffix}.txt");
    }

    private static string SanitizeFileNamePart(string value)
    {
        var invalidChars = Path.GetInvalidFileNameChars();
        var sanitized = new string(value.Select(character => invalidChars.Contains(character) ? '_' : character).ToArray());
        return string.IsNullOrWhiteSpace(sanitized) ? "PLC" : sanitized;
    }
}

internal sealed class PlcSelection
{
    public string DeviceName { get; set; } = string.Empty;
    public string SoftwareName { get; set; } = string.Empty;
    public string DisplayName => $"{DeviceName} / {SoftwareName}";
}

internal sealed class ExportResult
{
    public string OutputPath { get; set; } = string.Empty;
    public int ExportSuccessCount { get; set; }
    public int ExportErrorCount { get; set; }
    public string WarningMessage { get; set; } = string.Empty;
    public string ErrorLogPath { get; set; } = string.Empty;
}

internal sealed class BuildResult
{
    public int MatchedStructureCount { get; set; }
    public int RedEntryCount { get; set; }
    public int CreatedFolderCount { get; set; }
    public int CreatedBlockCount { get; set; }
    public int SkippedItemCount { get; set; }
    public int ErrorCount { get; set; }
    public string WarningMessage { get; set; } = string.Empty;
    public string BuildLogPath { get; set; } = string.Empty;
}

internal sealed class WorksheetEntry
{
    public string Value { get; set; } = string.Empty;
    public string StructureValue { get; set; } = string.Empty;
    public bool IsRed { get; set; }
    public string CellAddress { get; set; } = string.Empty;
}

internal sealed class TemplateMapping
{
    public string StructureKey { get; set; } = string.Empty;
    public string SourceEntry { get; set; } = string.Empty;
    public List<string> TargetEntries { get; set; } = new();
}

internal sealed class WorksheetAnalysis
{
    public int DataRowCount { get; set; }
    public int IoTagValueCount { get; set; }
    public int FunctionValueCount { get; set; }
    public int CandidateRowCount { get; set; }
    public int RedMarkedCount { get; set; }
    public int ExactFunctionGroupCount { get; set; }
    public int PatternGroupCount { get; set; }
    public List<string> RedColorSamples { get; set; } = new();
}

internal sealed class ReplacementPlan
{
    public string ReplacementKey { get; set; } = string.Empty;
    public string SourceToken { get; set; } = string.Empty;
    public string TargetToken { get; set; } = string.Empty;
    public string MappingKey { get; set; } = string.Empty;
    public HashSet<string> AllowedTargetIoTags { get; set; } = new(StringComparer.OrdinalIgnoreCase);
}

internal sealed class GroupInfo
{
    public PlcBlockGroup Group { get; set; } = null!;
    public string Path { get; set; } = string.Empty;
}

internal sealed class BlockInfo
{
    public PlcBlock Block { get; set; } = null!;
    public PlcBlockGroup Group { get; set; } = null!;
    public string GroupPath { get; set; } = string.Empty;
}

internal sealed class BlowerDbEntry
{
    public string TargetDbName { get; set; } = string.Empty;
    public string MemberName { get; set; } = string.Empty;
    public string SourceIoTag { get; set; } = string.Empty;
    public string DataType { get; set; } = string.Empty;
    public string Comment { get; set; } = string.Empty;
}

internal sealed class BlowerWorksheetAnalysis
{
    public int PlcBlowerRowCount { get; set; }
    public int RedMarkedPlcBlowerRowCount { get; set; }
    public int NonRedSkippedCount { get; set; }
    public int InvalidConnectingTypeCount { get; set; }
    public int MissingIoTagCount { get; set; }
}

internal sealed class StructureToken
{
    public StructureToken(string raw)
    {
        Raw = raw;
        var splitIndex = 0;
        while (splitIndex < raw.Length && char.IsLetter(raw[splitIndex]))
        {
            splitIndex++;
        }

        Prefix = raw.Substring(0, splitIndex);
        NumberPart = splitIndex < raw.Length ? raw.Substring(splitIndex) : string.Empty;
    }

    public string Raw { get; }
    public string Prefix { get; }
    public string NumberPart { get; }
}

internal sealed class ExportForm : Form
{
    private readonly TextBox _projectPathTextBox;
    private readonly CheckBox _withUiCheckBox;
    private readonly Button _browseButton;
    private readonly Button _loadPlcsButton;
    private readonly Button _startButton;
    private readonly ListBox _plcListBox;
    private readonly Label _plcInfoLabel;
    private readonly TextBox _excelFilePathTextBox;
    private readonly Button _browseExcelButton;
    private readonly Label _worksheetSelectionLabel;
    private readonly ProgressBar _progressBar;
    private readonly Label _progressLabel;
    private string? _selectedWorksheetName;

    public ExportForm()
    {
        Text = "TIA Program Block Exporter";
        StartPosition = FormStartPosition.CenterScreen;
        Width = 760;
        Height = 520;
        FormBorderStyle = FormBorderStyle.FixedDialog;
        MaximizeBox = false;

        var projectLabel = new Label
        {
            Left = 12,
            Top = 20,
            Width = 120,
            Text = "TIA-Projekt:"
        };

        _projectPathTextBox = new TextBox
        {
            Left = 12,
            Top = 45,
            Width = 620
        };

        _browseButton = new Button
        {
            Left = 640,
            Top = 43,
            Width = 90,
            Text = "Auswählen"
        };
        _browseButton.Click += (_, _) => SelectProjectFile();

        _withUiCheckBox = new CheckBox
        {
            Left = 12,
            Top = 80,
            Width = 250,
            Text = "TIA Portal mit Oberfläche öffnen",
            Checked = true
        };

        _loadPlcsButton = new Button
        {
            Left = 270,
            Top = 77,
            Width = 120,
            Text = "PLCs laden"
        };
        _loadPlcsButton.Click += async (_, _) => await LoadPlcsAsync();

        _plcInfoLabel = new Label
        {
            Left = 12,
            Top = 108,
            Width = 720,
            Text = "Gefundene PLCs mit Programmstruktur:"
        };

        _plcListBox = new ListBox
        {
            Left = 12,
            Top = 130,
            Width = 718,
            Height = 150,
            DisplayMember = nameof(PlcSelection.DisplayName)
        };
        _plcListBox.SelectedIndexChanged += (_, _) => UpdateBuildButtonState();

        var excelLabel = new Label
        {
            Left = 12,
            Top = 290,
            Width = 120,
            Text = "Excel-Datei:"
        };

        _excelFilePathTextBox = new TextBox
        {
            Left = 12,
            Top = 312,
            Width = 620,
            ReadOnly = true
        };

        _browseExcelButton = new Button
        {
            Left = 640,
            Top = 310,
            Width = 90,
            Text = "Auswählen"
        };
        _browseExcelButton.Click += (_, _) => SelectExcelFile();

        _worksheetSelectionLabel = new Label
        {
            Left = 12,
            Top = 342,
            Width = 718,
            Text = "Ausgewählte Tabelle: keine"
        };

        _startButton = new Button
        {
            Left = 12,
            Top = 370,
            Width = 120,
            Text = "Build",
            Enabled = false
        };
        _startButton.Click += (_, _) => StartExport();

        _progressBar = new ProgressBar
        {
            Left = 12,
            Top = 408,
            Width = 718,
            Height = 24,
            Minimum = 0,
            Maximum = 100,
            Value = 0,
            Style = ProgressBarStyle.Blocks
        };

        _progressLabel = new Label
        {
            Left = 12,
            Top = 438,
            Width = 718,
            Text = "Bereit"
        };

        Controls.Add(projectLabel);
        Controls.Add(_projectPathTextBox);
        Controls.Add(_browseButton);
        Controls.Add(_withUiCheckBox);
        Controls.Add(_loadPlcsButton);
        Controls.Add(_plcInfoLabel);
        Controls.Add(_plcListBox);
        Controls.Add(excelLabel);
        Controls.Add(_excelFilePathTextBox);
        Controls.Add(_browseExcelButton);
        Controls.Add(_worksheetSelectionLabel);
        Controls.Add(_startButton);
        Controls.Add(_progressBar);
        Controls.Add(_progressLabel);

        UpdateBuildButtonState();
    }

    private void SelectProjectFile()
    {
        using var fileDialog = new OpenFileDialog
        {
            Filter = "TIA Projekt (*.ap*)|*.ap*|Alle Dateien (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false
        };

        if (fileDialog.ShowDialog(this) == DialogResult.OK)
        {
            _projectPathTextBox.Text = fileDialog.FileName;
            _plcListBox.Items.Clear();
            _plcInfoLabel.Text = "Gefundene PLCs mit Programmstruktur:";
            UpdateBuildButtonState();
        }
    }

    private void SelectExcelFile()
    {
        using var fileDialog = new OpenFileDialog
        {
            Filter = "Excel Dateien (*.xlsx;*.xls)|*.xlsx;*.xls|Alle Dateien (*.*)|*.*",
            CheckFileExists = true,
            Multiselect = false
        };

        if (fileDialog.ShowDialog(this) != DialogResult.OK)
        {
            return;
        }

        try
        {
            var worksheetNames = GetWorksheetNames(fileDialog.FileName);
            if (worksheetNames.Count == 0)
            {
                MessageBox.Show(this, "In der ausgewählten Datei wurde keine Tabelle gefunden.", "Keine Tabellen gefunden", MessageBoxButtons.OK, MessageBoxIcon.Warning);
                return;
            }

            var selectedWorksheet = PromptWorksheetSelection(worksheetNames);
            if (string.IsNullOrWhiteSpace(selectedWorksheet))
            {
                return;
            }

            _excelFilePathTextBox.Text = fileDialog.FileName;
            _selectedWorksheetName = selectedWorksheet;
            _worksheetSelectionLabel.Text = $"Ausgewählte Tabelle: {_selectedWorksheetName}";
            UpdateBuildButtonState();
        }
        catch (Exception ex)
        {
            MessageBox.Show(this, $"Fehler beim Lesen der Excel-Datei:\n{ex.Message}", "Excel-Datei fehlerhaft", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
    }

    private void UpdateBuildButtonState()
    {
        var excelSelected = !string.IsNullOrWhiteSpace(_excelFilePathTextBox.Text) && File.Exists(_excelFilePathTextBox.Text);
        var worksheetSelected = !string.IsNullOrWhiteSpace(_selectedWorksheetName);
        var plcSelected = _plcListBox.SelectedItem is PlcSelection;
        _startButton.Enabled = excelSelected && worksheetSelected && plcSelected;
    }

    private static List<string> GetWorksheetNames(string excelFilePath)
    {
        using var stream = File.OpenRead(excelFilePath);
        IWorkbook workbook = Path.GetExtension(excelFilePath).Equals(".xls", StringComparison.OrdinalIgnoreCase)
            ? new HSSFWorkbook(stream)
            : new XSSFWorkbook(stream);

        var sheetNames = new List<string>();
        for (var sheetIndex = 0; sheetIndex < workbook.NumberOfSheets; sheetIndex++)
        {
            var sheetName = workbook.GetSheetName(sheetIndex);
            if (!string.IsNullOrWhiteSpace(sheetName))
            {
                sheetNames.Add(sheetName);
            }
        }

        return sheetNames;
    }

    private string? PromptWorksheetSelection(List<string> worksheetNames)
    {
        using var selectionForm = new Form
        {
            Text = "Tabelle auswählen",
            StartPosition = FormStartPosition.CenterParent,
            Width = 420,
            Height = 360,
            FormBorderStyle = FormBorderStyle.FixedDialog,
            MaximizeBox = false,
            MinimizeBox = false
        };

        var infoLabel = new Label
        {
            Left = 12,
            Top = 12,
            Width = 380,
            Text = "Bitte eine Tabelle aus der Excel-Datei auswählen:"
        };

        var tableListBox = new ListBox
        {
            Left = 12,
            Top = 36,
            Width = 380,
            Height = 230
        };
        tableListBox.Items.AddRange(worksheetNames.Cast<object>().ToArray());
        if (tableListBox.Items.Count > 0)
        {
            tableListBox.SelectedIndex = 0;
        }

        var okButton = new Button
        {
            Left = 236,
            Top = 276,
            Width = 75,
            Text = "OK",
            DialogResult = DialogResult.OK
        };

        var cancelButton = new Button
        {
            Left = 317,
            Top = 276,
            Width = 75,
            Text = "Abbrechen",
            DialogResult = DialogResult.Cancel
        };

        selectionForm.Controls.Add(infoLabel);
        selectionForm.Controls.Add(tableListBox);
        selectionForm.Controls.Add(okButton);
        selectionForm.Controls.Add(cancelButton);
        selectionForm.AcceptButton = okButton;
        selectionForm.CancelButton = cancelButton;

        return selectionForm.ShowDialog(this) == DialogResult.OK
            ? tableListBox.SelectedItem?.ToString()
            : null;
    }

    private async Task LoadPlcsAsync()
    {
        var projectPath = _projectPathTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(projectPath) || !File.Exists(projectPath))
        {
            MessageBox.Show(this, "Bitte eine gültige TIA-Projektdatei auswählen.", "Ungültiger Pfad", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            SetUiBusyState(true, "PLCs werden gesucht...");
            var selections = await Task.Run(() => Program.DiscoverExportablePlcs(projectPath, _withUiCheckBox.Checked));

            _plcListBox.Items.Clear();
            foreach (var selection in selections)
            {
                _plcListBox.Items.Add(selection);
            }

            if (_plcListBox.Items.Count > 0)
            {
                _plcListBox.SelectedIndex = 0;
                _plcInfoLabel.Text = $"Gefundene PLCs mit Programmstruktur: {_plcListBox.Items.Count}";
                _progressLabel.Text = "PLC-Auswahl bereit";
            }
            else
            {
                _plcInfoLabel.Text = "Keine exportierbaren PLCs gefunden.";
                _progressLabel.Text = "Keine PLCs gefunden";
            }

            _progressBar.Style = ProgressBarStyle.Blocks;
            _progressBar.MarqueeAnimationSpeed = 0;
            _progressBar.Value = 100;
        }
        catch (Exception ex)
        {
            _progressBar.Style = ProgressBarStyle.Blocks;
            _progressBar.MarqueeAnimationSpeed = 0;
            _progressBar.Value = 0;
            _progressLabel.Text = "PLC-Suche fehlgeschlagen";
            MessageBox.Show(this, $"Fehler beim Laden der PLCs:\n{ex.Message}", "PLC-Suche fehlgeschlagen", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetUiBusyState(false, _progressLabel.Text);
        }
    }

    private async void StartExport()
    {
        var projectPath = _projectPathTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(projectPath) || !File.Exists(projectPath))
        {
            MessageBox.Show(this, "Bitte eine gültige TIA-Projektdatei auswählen.", "Ungültiger Pfad", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (_plcListBox.SelectedItem is not PlcSelection selectedPlc)
        {
            MessageBox.Show(this, "Bitte zuerst die PLCs laden und eine PLC auswählen.", "PLC auswählen", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        var excelPath = _excelFilePathTextBox.Text.Trim();
        if (string.IsNullOrWhiteSpace(excelPath) || !File.Exists(excelPath))
        {
            MessageBox.Show(this, "Bitte eine gültige Excel-Datei auswählen.", "Excel-Datei auswählen", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        if (string.IsNullOrWhiteSpace(_selectedWorksheetName))
        {
            MessageBox.Show(this, "Bitte eine Tabelle aus der Excel-Datei auswählen.", "Tabelle auswählen", MessageBoxButtons.OK, MessageBoxIcon.Warning);
            return;
        }

        try
        {
            SetUiBusyState(true, "Build läuft...");

            var result = await Task.Run(() => Program.BuildFromExcelTemplate(projectPath, _withUiCheckBox.Checked, selectedPlc, excelPath, _selectedWorksheetName!));

            _progressBar.Style = ProgressBarStyle.Blocks;
            _progressBar.MarqueeAnimationSpeed = 0;
            _progressBar.Value = 100;
            _progressLabel.Text = "Build abgeschlossen";

            var message = $"Build abgeschlossen.\n\nStatistik:\nStrukturen={result.MatchedStructureCount}\nRote Einträge={result.RedEntryCount}\nOrdner erstellt={result.CreatedFolderCount}\nBlöcke erstellt={result.CreatedBlockCount}\nÜbersprungen={result.SkippedItemCount}\nFehler={result.ErrorCount}";
            if (!string.IsNullOrWhiteSpace(result.BuildLogPath))
            {
                message += $"\n\nBuild-Log:\n{result.BuildLogPath}";
            }

            if (!string.IsNullOrWhiteSpace(result.WarningMessage))
            {
                message += $"\n\nWarnung:\n{result.WarningMessage}";
            }

            MessageBox.Show(this, message, "Build abgeschlossen", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            _progressBar.Style = ProgressBarStyle.Blocks;
            _progressBar.MarqueeAnimationSpeed = 0;
            _progressBar.Value = 0;
            _progressLabel.Text = "Build fehlgeschlagen";

            var errorMessage = $"Fehler beim Build:\n{ex.Message}";
            var buildLogPath = ExtractBuildLogPathFromException(ex);
            if (!string.IsNullOrWhiteSpace(buildLogPath))
            {
                errorMessage += $"\n\nBuild-Log:\n{buildLogPath}";
            }

            MessageBox.Show(this, errorMessage, "Build fehlgeschlagen", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetUiBusyState(false, _progressLabel.Text);
        }
    }

    private static string? ExtractBuildLogPathFromException(Exception exception)
    {
        const string marker = "Build-Log:";
        Exception? current = exception;
        while (current is not null)
        {
            var message = current.Message;
            var markerIndex = message.IndexOf(marker, StringComparison.OrdinalIgnoreCase);
            if (markerIndex >= 0)
            {
                var candidatePath = message.Substring(markerIndex + marker.Length).Trim();
                if (!string.IsNullOrWhiteSpace(candidatePath))
                {
                    return candidatePath;
                }
            }

            current = current.InnerException;
        }

        return null;
    }

    private void SetUiBusyState(bool busy, string statusText)
    {
        UseWaitCursor = busy;
        _browseButton.Enabled = !busy;
        _projectPathTextBox.Enabled = !busy;
        _withUiCheckBox.Enabled = !busy;
        _loadPlcsButton.Enabled = !busy;
        _plcListBox.Enabled = !busy;
        _browseExcelButton.Enabled = !busy;
        if (busy)
        {
            _startButton.Enabled = false;
        }
        else
        {
            UpdateBuildButtonState();
        }
        _progressLabel.Text = statusText;

        if (busy)
        {
            _progressBar.Style = ProgressBarStyle.Marquee;
            _progressBar.MarqueeAnimationSpeed = 30;
        }
    }
}

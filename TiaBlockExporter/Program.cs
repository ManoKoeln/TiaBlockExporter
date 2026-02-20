using System.Runtime.Remoting;
using System.Threading.Tasks;
using System.Xml.Linq;
using System.Windows.Forms;
using Siemens.Engineering;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;

internal static class Program
{
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

internal sealed class ExportForm : Form
{
    private readonly TextBox _projectPathTextBox;
    private readonly CheckBox _withUiCheckBox;
    private readonly Button _browseButton;
    private readonly Button _loadPlcsButton;
    private readonly Button _startButton;
    private readonly ListBox _plcListBox;
    private readonly Label _plcInfoLabel;
    private readonly ProgressBar _progressBar;
    private readonly Label _progressLabel;

    public ExportForm()
    {
        Text = "TIA Program Block Exporter";
        StartPosition = FormStartPosition.CenterScreen;
        Width = 760;
        Height = 430;
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

        _startButton = new Button
        {
            Left = 12,
            Top = 292,
            Width = 120,
            Text = "Export starten"
        };
        _startButton.Click += (_, _) => StartExport();

        _progressBar = new ProgressBar
        {
            Left = 12,
            Top = 330,
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
            Top = 360,
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
        Controls.Add(_startButton);
        Controls.Add(_progressBar);
        Controls.Add(_progressLabel);
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
        }
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

        try
        {
            SetUiBusyState(true, "Export läuft...");

            var result = await Task.Run(() => Program.ExportProject(projectPath, null, _withUiCheckBox.Checked, selectedPlc));

            _progressBar.Style = ProgressBarStyle.Blocks;
            _progressBar.MarqueeAnimationSpeed = 0;
            _progressBar.Value = 100;
            _progressLabel.Text = "Export abgeschlossen";

            var message = $"Export abgeschlossen:\n{result.OutputPath}\n\nError-Log:\n{result.ErrorLogPath}\n\nStatistik:\nErfolgreich={result.ExportSuccessCount}, Fehler={result.ExportErrorCount}, Gesamt={result.ExportSuccessCount + result.ExportErrorCount}";
            if (!string.IsNullOrWhiteSpace(result.WarningMessage))
            {
                message += $"\n\nWarnung:\n{result.WarningMessage}";
            }

            MessageBox.Show(this, message, "Export abgeschlossen", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }
        catch (Exception ex)
        {
            _progressBar.Style = ProgressBarStyle.Blocks;
            _progressBar.MarqueeAnimationSpeed = 0;
            _progressBar.Value = 0;
            _progressLabel.Text = "Export fehlgeschlagen";
            MessageBox.Show(this, $"Fehler beim Export:\n{ex.Message}", "Export fehlgeschlagen", MessageBoxButtons.OK, MessageBoxIcon.Error);
        }
        finally
        {
            SetUiBusyState(false, _progressLabel.Text);
        }
    }

    private void SetUiBusyState(bool busy, string statusText)
    {
        UseWaitCursor = busy;
        _browseButton.Enabled = !busy;
        _projectPathTextBox.Enabled = !busy;
        _withUiCheckBox.Enabled = !busy;
        _loadPlcsButton.Enabled = !busy;
        _plcListBox.Enabled = !busy;
        _startButton.Enabled = !busy;
        _progressLabel.Text = statusText;

        if (busy)
        {
            _progressBar.Style = ProgressBarStyle.Marquee;
            _progressBar.MarqueeAnimationSpeed = 30;
        }
    }
}

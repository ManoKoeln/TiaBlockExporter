using System.Xml.Linq;
using System.Runtime.Remoting;
using Siemens.Engineering;
using Siemens.Engineering.HW;
using Siemens.Engineering.HW.Features;
using Siemens.Engineering.SW;
using Siemens.Engineering.SW.Blocks;

var arguments = ParseArguments(args);
if (!arguments.TryGetValue("project", out var projectPath) || string.IsNullOrWhiteSpace(projectPath))
{
    Console.Error.WriteLine("Fehlender Parameter: --project <PfadZur.ap19>");
    return 1;
}

var currentDirectory = Directory.GetCurrentDirectory();
var resolvedProjectPath = Path.GetFullPath(projectPath);

Console.WriteLine($"Arbeitsverzeichnis: {currentDirectory}");
Console.WriteLine($"Parameter --project: {projectPath}");
Console.WriteLine($"Aufgelöster Projektpfad: {resolvedProjectPath}");

if (!File.Exists(resolvedProjectPath))
{
    Console.Error.WriteLine($"Projektdatei nicht gefunden: {resolvedProjectPath}");
    return 2;
}

var outputPath = arguments.TryGetValue("output", out var output) && !string.IsNullOrWhiteSpace(output)
    ? output
    : Path.Combine(Path.GetDirectoryName(resolvedProjectPath)!, "ProgramBlocksExport.xml");
var resolvedOutputPath = Path.GetFullPath(outputPath);

var withUi = arguments.ContainsKey("with-ui");

Console.WriteLine($"Parameter --output: {outputPath}");
Console.WriteLine($"Aufgelöster Outputpfad: {resolvedOutputPath}");

try
{
    var exportSuccessCount = 0;
    var exportErrorCount = 0;

    Project? project = null;
    using var tiaPortal = new TiaPortal(withUi ? TiaPortalMode.WithUserInterface : TiaPortalMode.WithoutUserInterface);
    project = tiaPortal.Projects.Open(new FileInfo(resolvedProjectPath));
    var softwareInfos = EnumeratePlcSoftware(project).ToList();

    var root = new XElement("TiaProgramBlockExport",
        new XAttribute("project", resolvedProjectPath),
        new XAttribute("exportedAt", DateTimeOffset.Now.ToString("O")));

    foreach (var softwareInfo in softwareInfos)
    {
        var softwareElement = new XElement("PlcSoftware",
            new XAttribute("device", softwareInfo.DeviceName),
            new XAttribute("software", softwareInfo.SoftwareName));

        ExportBlockGroup(softwareInfo.PlcSoftware.BlockGroup, softwareElement, string.Empty, ref exportSuccessCount, ref exportErrorCount);
        root.Add(softwareElement);
    }

    var document = new XDocument(new XDeclaration("1.0", "utf-8", "yes"), root);
    Directory.CreateDirectory(Path.GetDirectoryName(resolvedOutputPath) ?? ".");
    document.Save(resolvedOutputPath);

    TryCloseProject(project);

    Console.WriteLine($"Export abgeschlossen: {resolvedOutputPath}");
    Console.WriteLine($"Statistik: Erfolgreich={exportSuccessCount}, Fehler={exportErrorCount}, Gesamt={exportSuccessCount + exportErrorCount}");
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

static void TryCloseProject(Project? project)
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

static Dictionary<string, string?> ParseArguments(string[] args)
{
    var result = new Dictionary<string, string?>(StringComparer.OrdinalIgnoreCase);

    for (var i = 0; i < args.Length; i++)
    {
        var current = args[i];
        if (!current.StartsWith("--", StringComparison.Ordinal))
        {
            continue;
        }

        var key = current.Substring(2);
        if (i + 1 < args.Length && !args[i + 1].StartsWith("--", StringComparison.Ordinal))
        {
            result[key] = args[i + 1];
            i++;
        }
        else
        {
            result[key] = null;
        }
    }

    return result;
}

static IEnumerable<(string DeviceName, string SoftwareName, PlcSoftware PlcSoftware)> EnumeratePlcSoftware(Project project)
{
    foreach (var device in project.Devices)
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

static IEnumerable<DeviceItem> TraverseDeviceItems(DeviceItemComposition composition)
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

static void ExportBlockGroup(PlcBlockGroup group, XElement parentElement, string groupPath, ref int exportSuccessCount, ref int exportErrorCount)
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
        ExportBlockGroup(childGroup, parentElement, currentGroupPath, ref exportSuccessCount, ref exportErrorCount);
    }
}

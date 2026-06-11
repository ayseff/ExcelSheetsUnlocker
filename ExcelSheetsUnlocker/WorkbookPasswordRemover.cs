using System.IO.Compression;
using System.Xml.Linq;

namespace ExcelSheetsUnlocker;

public sealed record SheetUnlockResult(string SheetName, bool ProtectionRemoved);

public sealed record WorkbookUnlockResult(string OutputPath, IReadOnlyList<SheetUnlockResult> Sheets);

internal sealed record WorksheetMetadata(string SheetName, string WorksheetPath);

public sealed class WorkbookPasswordRemover
{
    private static readonly XNamespace RelationshipNamespace = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
    private static readonly XNamespace PackageRelationshipNamespace = "http://schemas.openxmlformats.org/package/2006/relationships";

    public async Task<WorkbookUnlockResult> UnlockWorkbookAsync(string workbookPath, CancellationToken cancellationToken)
    {
        ValidateWorkbookPath(workbookPath);

        var workbookBytes = await File.ReadAllBytesAsync(workbookPath, cancellationToken);

        using var inputStream = new MemoryStream(workbookBytes, writable: false);
        using var inputArchive = new ZipArchive(inputStream, ZipArchiveMode.Read, leaveOpen: true);

        var worksheets = LoadWorksheetMetadata(inputArchive);
        var worksheetMap = worksheets.ToDictionary(sheet => sheet.WorksheetPath, StringComparer.OrdinalIgnoreCase);
        var processedResults = new Dictionary<string, SheetUnlockResult>(StringComparer.OrdinalIgnoreCase);

        using var outputStream = new MemoryStream();

        using (var outputArchive = new ZipArchive(outputStream, ZipArchiveMode.Create, leaveOpen: true))
        {
            foreach (var entry in inputArchive.Entries)
            {
                cancellationToken.ThrowIfCancellationRequested();

                var outputEntry = outputArchive.CreateEntry(entry.FullName, CompressionLevel.Optimal);
                outputEntry.LastWriteTime = entry.LastWriteTime;

                if (string.IsNullOrEmpty(entry.Name))
                {
                    continue;
                }

                await using var sourceStream = entry.Open();
                await using var destinationStream = outputEntry.Open();

                if (!worksheetMap.TryGetValue(entry.FullName, out var worksheet))
                {
                    await sourceStream.CopyToAsync(destinationStream, cancellationToken);
                    continue;
                }

                // Only worksheet XML gets rewritten; every other workbook part is copied byte-for-byte.
                var result = await CopyWorksheetWithoutProtectionAsync(sourceStream, destinationStream, worksheet, cancellationToken);
                processedResults[worksheet.WorksheetPath] = result;
            }
        }

        EnsureEveryWorksheetWasProcessed(worksheets, processedResults);
        var orderedResults = worksheets.Select(worksheet => processedResults[worksheet.WorksheetPath]).ToList();

        var outputPath = BuildOutputPath(workbookPath);
        outputStream.Position = 0;
        await File.WriteAllBytesAsync(outputPath, outputStream.ToArray(), cancellationToken);

        return new WorkbookUnlockResult(outputPath, orderedResults);
    }

    private static void ValidateWorkbookPath(string workbookPath)
    {
        if (!File.Exists(workbookPath))
        {
            throw new FileNotFoundException("The workbook could not be found.", workbookPath);
        }

        var extension = Path.GetExtension(workbookPath);

        if (!string.Equals(extension, ".xlsx", StringComparison.OrdinalIgnoreCase) &&
            !string.Equals(extension, ".xlsm", StringComparison.OrdinalIgnoreCase))
        {
            throw new InvalidOperationException("Only .xlsx and .xlsm workbooks are supported.");
        }
    }

    private static List<WorksheetMetadata> LoadWorksheetMetadata(ZipArchive archive)
    {
        var workbookEntry = archive.GetEntry("xl/workbook.xml")
            ?? throw new InvalidDataException("The workbook is missing xl/workbook.xml.");
        var relationshipsEntry = archive.GetEntry("xl/_rels/workbook.xml.rels")
            ?? throw new InvalidDataException("The workbook is missing xl/_rels/workbook.xml.rels.");

        var workbookDocument = LoadDocument(workbookEntry);
        var relationshipsDocument = LoadDocument(relationshipsEntry);
        var relationshipTargets = LoadRelationshipTargets(relationshipsDocument);

        var workbookNamespace = workbookDocument.Root?.Name.Namespace
            ?? throw new InvalidDataException("The workbook XML has no root element.");
        var sheetsElement = workbookDocument.Root.Element(workbookNamespace + "sheets")
            ?? throw new InvalidDataException("The workbook XML is missing the sheets element.");

        var worksheets = new List<WorksheetMetadata>();

        foreach (var sheetElement in sheetsElement.Elements(workbookNamespace + "sheet"))
        {
            var sheetName = sheetElement.Attribute("name")?.Value
                ?? throw new InvalidDataException("A workbook sheet is missing its name.");
            var relationshipId = sheetElement.Attribute(RelationshipNamespace + "id")?.Value
                ?? throw new InvalidDataException($"Worksheet '{sheetName}' is missing its relationship id.");

            if (!relationshipTargets.TryGetValue(relationshipId, out var worksheetPath))
            {
                throw new InvalidDataException($"Worksheet '{sheetName}' points to an unknown relationship '{relationshipId}'.");
            }

            if (!worksheetPath.StartsWith("xl/worksheets/", StringComparison.OrdinalIgnoreCase))
            {
                continue;
            }

            worksheets.Add(new WorksheetMetadata(sheetName, worksheetPath));
        }

        return worksheets;
    }

    private static Dictionary<string, string> LoadRelationshipTargets(XDocument relationshipsDocument)
    {
        var relationshipsRoot = relationshipsDocument.Root
            ?? throw new InvalidDataException("The workbook relationships XML has no root element.");

        return relationshipsRoot
            .Elements(PackageRelationshipNamespace + "Relationship")
            .ToDictionary(
                relationship => relationship.Attribute("Id")?.Value
                    ?? throw new InvalidDataException("A workbook relationship is missing its Id."),
                relationship =>
                {
                    var target = relationship.Attribute("Target")?.Value
                        ?? throw new InvalidDataException("A workbook relationship is missing its Target.");
                    return ResolveWorksheetPath(target);
                },
                StringComparer.OrdinalIgnoreCase);
    }

    private static XDocument LoadDocument(ZipArchiveEntry entry)
    {
        using var stream = entry.Open();
        return XDocument.Load(stream, LoadOptions.PreserveWhitespace);
    }

    private static string ResolveWorksheetPath(string target)
    {
        // Relationship targets are stored relative to xl/workbook.xml, so resolve them exactly once here.
        var baseUri = new Uri("https://local/xl/workbook.xml");
        var resolvedUri = new Uri(baseUri, target);

        return resolvedUri.AbsolutePath.TrimStart('/');
    }

    private static async Task<SheetUnlockResult> CopyWorksheetWithoutProtectionAsync(
        Stream sourceStream,
        Stream destinationStream,
        WorksheetMetadata worksheet,
        CancellationToken cancellationToken)
    {
        var document = await XDocument.LoadAsync(sourceStream, LoadOptions.PreserveWhitespace, cancellationToken);
        var protectionNodes = document
            .Descendants()
            .Where(node => string.Equals(node.Name.LocalName, "sheetProtection", StringComparison.Ordinal))
            .ToList();

        foreach (var protectionNode in protectionNodes)
        {
            protectionNode.Remove();
        }

        await document.SaveAsync(destinationStream, SaveOptions.DisableFormatting, cancellationToken);

        return new SheetUnlockResult(worksheet.SheetName, protectionNodes.Count > 0);
    }

    private static void EnsureEveryWorksheetWasProcessed(
        IReadOnlyCollection<WorksheetMetadata> worksheets,
        IReadOnlyDictionary<string, SheetUnlockResult> processedResults)
    {
        if (worksheets.Count == processedResults.Count)
        {
            return;
        }

        throw new InvalidDataException("One or more worksheets listed in the workbook were missing from the archive.");
    }

    private static string BuildOutputPath(string workbookPath)
    {
        var directory = Path.GetDirectoryName(workbookPath)
            ?? throw new InvalidOperationException("The workbook path does not have a valid directory.");
        var fileNameWithoutExtension = Path.GetFileNameWithoutExtension(workbookPath);
        var extension = Path.GetExtension(workbookPath);

        return Path.Combine(directory, $"{fileNameWithoutExtension} (password removed){extension}");
    }
}

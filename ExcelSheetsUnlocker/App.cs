using Microsoft.Extensions.Logging;

namespace ExcelSheetsUnlocker;

public sealed class App(WorkbookPasswordRemover workbookPasswordRemover, ILogger<App> logger)
{
    private readonly WorkbookPasswordRemover _workbookPasswordRemover = workbookPasswordRemover;
    private readonly ILogger<App> _logger = logger;

    public async Task RunAsync(string[] args, CancellationToken cancellationToken)
    {
        var workbookPath = GetWorkbookPath(args);
        var result = await _workbookPasswordRemover.UnlockWorkbookAsync(workbookPath, cancellationToken);
        var removedProtection = false;

        foreach (var sheet in result.Sheets)
        {
            if (sheet.ProtectionRemoved)
            {
                removedProtection = true;
                _logger.LogInformation("Removed worksheet protection from '{SheetName}'.", sheet.SheetName);
                continue;
            }

            _logger.LogInformation("No worksheet protection found in '{SheetName}'.", sheet.SheetName);
        }

        if (!removedProtection)
        {
            _logger.LogInformation("No worksheet protection nodes were found in the workbook.");
        }

        _logger.LogInformation("Unlocked workbook written to '{OutputPath}'.", result.OutputPath);
    }

    private static string GetWorkbookPath(string[] args)
    {
        var rawPath = args.Length > 0
            ? string.Join(' ', args)
            : PromptForWorkbookPath();

        return SanitizePath(rawPath);
    }

    private static string PromptForWorkbookPath()
    {
        Console.Write("Enter the path to the Excel workbook: ");
        return Console.ReadLine() ?? throw new ArgumentException("A workbook path is required.");
    }

    private static string SanitizePath(string path)
    {
        if (path is null)
        {
            throw new ArgumentException("A workbook path is required.");
        }

        var sanitizedPath = path.Trim().Trim('"');

        if (string.IsNullOrWhiteSpace(sanitizedPath))
        {
            throw new ArgumentException("A workbook path is required.");
        }

        return sanitizedPath;
    }
}

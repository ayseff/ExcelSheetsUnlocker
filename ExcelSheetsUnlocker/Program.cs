using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using Serilog;

namespace ExcelSheetsUnlocker;

public static class Program
{
    public static async Task<int> Main(string[] args)
    {
        Log.Logger = new LoggerConfiguration()
            .WriteTo.Console()
            .CreateLogger();

        try
        {
            var builder = Host.CreateApplicationBuilder(args);

            builder.Logging.ClearProviders();
            builder.Services.AddSerilog(Log.Logger, dispose: true);

            builder.Services.AddTransient<App>();
            builder.Services.AddTransient<WorkbookPasswordRemover>();

            using var host = builder.Build();

            await host.Services.GetRequiredService<App>().RunAsync(args, CancellationToken.None);

            return 0;
        }
        catch (Exception exception)
        {
            Log.Fatal(exception, "The workbook unlock operation failed.");
            return 1;
        }
        finally
        {
            await Log.CloseAndFlushAsync();
        }
    }
}

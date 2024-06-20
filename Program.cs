using Microsoft.Extensions.DependencyInjection;
using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using VirtualPrinterService;

var builder = Host.CreateDefaultBuilder(args);
builder.ConfigureServices(services =>
{
    services.AddSingleton<PrinterManagementService>();
    services.AddSingleton<PrinterServerService>();
    services.AddHostedService<Worker>();
    services.AddLogging(config =>
    {
        config.AddConsole();
    });
});

var host = builder.Build();
await host.RunAsync();

using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace VirtualPrinterService
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private readonly PrinterServerService _printerServerService;

        public Worker(ILogger<Worker> logger, PrinterServerService printerServerService)
        {
            _logger = logger;
            _printerServerService = printerServerService;
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            try
            {
                _logger.LogInformation("Starting printer server at: {time}", DateTimeOffset.Now);

                // The PrinterServerService should start listening in its own Run method
                _printerServerService.Run();

                _logger.LogInformation("Printer server is running.");

                // Keep the worker running until stopped
                while (!stoppingToken.IsCancellationRequested)
                {
                    await Task.Delay(1000, stoppingToken); // 1-second delay
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "An error occurred while managing the printer.");
            }
            finally
            {
                _logger.LogInformation("Stopping printer server at: {time}", DateTimeOffset.Now);

                // Cleanup and stop the PrinterServerService if necessary
                _printerServerService?.UninstallPrinter();
            }
        }
    }
}

using Microsoft.Extensions.Hosting;
using Microsoft.Extensions.Logging;
using PdfSharpCore.Drawing;
using PdfSharpCore.Pdf;
using System;
using System.IO;
using System.Threading;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;
using System.Xml.Linq;

namespace VirtualPrinterService
{
    public class Worker : BackgroundService
    {
        private readonly ILogger<Worker> _logger;
        private readonly PrinterManagementService _printerManagementService;
        private PrinterServerService _printerServerService;
        private bool isPrintJobCompleted; // Flag to track print job completion

        public Worker(ILogger<Worker> logger, PrinterManagementService printerManagement)
        {
            _logger = logger;
            _printerManagementService = printerManagement;
            isPrintJobCompleted = false; // Initialize flag
        }

        protected override async Task ExecuteAsync(CancellationToken stoppingToken)
        {
            try
            {
                _logger.LogInformation("Starting printer server at: {time}", DateTimeOffset.Now);

                // Define the printer name, IP, and port
                string printerName = "VirtualPrinter";
                string ip = "127.0.0.1";
                int port = 9100;

                // Initialize the PrinterServerService
                _printerServerService = new PrinterServerService(printerName, ip, port, autoInstallPrinter: true, printCallbackFn: PrintCallback);
                _printerServerService.Run();

                _logger.LogInformation("Printer server is running at IP: {ip}, Port: {port}", ip, port);

                // Keep the worker running until print job is completed or stoppingToken is canceled
                while (!isPrintJobCompleted && !stoppingToken.IsCancellationRequested)
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
                // Cleanup the printer server and uninstall the printer if necessary
                _printerServerService?.UninstallPrinter();
            }
        }

        private void PrintCallback(string doc, string title, string author, string filename)
        {
            try
            {
                // Log the details of the print job
                _logger.LogInformation("Received print job: Title={title}, Author={author}, Filename={filename}");

                // Process the print job here (saving document, etc.)
                // Save the document to the user's documents folder as PDF
                string documentsFolder = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);
                string savePath = Path.Combine(documentsFolder, $"{filename}.pdf");

                // Save the document as PDF
                SaveDocumentAsPdf(doc, savePath);

                // Set flag to indicate that print job is completed
                isPrintJobCompleted = true;
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error processing print job");
            }
        }

        private void SaveDocumentAsPdf(string doc, string savePath)
        {
            try
            {
                // Create a new PDF document
                PdfDocument pdfDocument = new PdfDocument();

                // Add a page
                PdfPage page = pdfDocument.AddPage();

                // Create a graphics object from the page
                XGraphics gfx = XGraphics.FromPdfPage(page);

                // Create a font
                XFont font = new XFont("Verdana", 12, XFontStyle.Regular);

                // Draw the document content on the page
                gfx.DrawString(doc, font, XBrushes.Black, new XRect(10, 10, page.Width.Point - 20, page.Height.Point - 20), XStringFormats.TopLeft);

                // Save the PDF document to the specified path
                pdfDocument.Save(savePath);

                _logger.LogInformation("Document saved as PDF to: {savePath}", savePath);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "Error saving document as PDF to: {savePath}", savePath);
            }
        }

    }
}
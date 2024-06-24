using System;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using Aspose.Words.Shaping;

namespace VirtualPrinterService
{
    public delegate void PrintCallbackFunctionType(
        string doc,
        string title,
        string author,
        string filename
    );

    public class PrinterServerService
    {
        private readonly string printerName;
        private string ip;
        private int port;
        private readonly bool autoInstallPrinter;
        private readonly PrintCallbackFunctionType printCallbackFn;
        private bool running;
        private bool keepGoing;
        private PrinterManagementService osPrinterManager;
        private string printerPortName;

        public PrinterServerService(
            string printerName = "My Virtual Printer",
            string ip = "127.0.0.1",
            int port = 9100,
            bool autoInstallPrinter = true,
            PrintCallbackFunctionType printCallbackFn = null
        )
        {
            this.printerName = printerName;
            this.ip = ip;
            this.port = port;
            this.autoInstallPrinter = autoInstallPrinter;
            this.printCallbackFn = printCallbackFn;
            this.running = false;
            this.keepGoing = false;
            this.osPrinterManager = null;
            this.printerPortName = null;
        }

        ~PrinterServerService()
        {
            if (autoInstallPrinter)
            {
                UninstallPrinter();
            }
        }

        private void InstallPrinter(string ip, int port)
        {
            if (Environment.OSVersion.Platform == PlatformID.Win32NT)
            {
                osPrinterManager = new PrinterManagementService();
                printerPortName = $"IP_{ip.Replace(".", "_")}_{port}";
                string printerName = "VirtualPrinter";
                bool makeDefault = false;
                string comment = null;
                osPrinterManager.AddPrinter(printerName, ip, port, printerPortName, makeDefault, comment);
            }
            else
            {
                Console.WriteLine($"WARN: Auto install not implemented for OS {Environment.OSVersion.Platform}");
            }
        }

        public void UninstallPrinter()
        {
            if (osPrinterManager != null)
            {
                osPrinterManager.RemovePrinter(printerName);
                osPrinterManager.RemovePort(printerPortName);
            }
        }

        public void Run()
        {
            if (running)
                return;

            running = true;
            keepGoing = true;

            TcpListener listener = new TcpListener(IPAddress.Parse(ip), port);
            listener.Start();

            IPEndPoint localEndpoint = (IPEndPoint)listener.LocalEndpoint;
            ip = localEndpoint.Address.ToString();
            port = localEndpoint.Port;

            Console.WriteLine($"Opening {ip}:{port}");

            if (autoInstallPrinter)
            {
                InstallPrinter(ip, port);
            }

            while (keepGoing)
            {
                Console.WriteLine("\nListening for incoming print job...");
                if (!keepGoing)
                    continue;

                if (!listener.Pending())
                {
                    Thread.Sleep(1000);
                    continue;
                }

                Console.WriteLine("Incoming job... spooling...");

                TcpClient client = listener.AcceptTcpClient();
                NetworkStream stream = client.GetStream();

                // Read the incoming binary data and process it
                using (MemoryStream memoryStream = new MemoryStream())
                {
                    stream.CopyTo(memoryStream);
                    byte[] documentContent = memoryStream.ToArray();
                    string fulldoc = Convert.ToBase64String(documentContent);
                    Console.WriteLine("Received document content");

                    string pdfText = ExtractPdfText(documentContent);
                    Console.WriteLine(pdfText);

                    // Prompt user to print the document
                    Console.Write("Do you want to print the document? (yes/no): ");
                    string input = Console.ReadLine()?.ToLower();

                    if (input == "yes" || input == "y")
                    {
                        // Call PrintCallback to process the document content
                        printCallbackFn?.Invoke(fulldoc, "Print Job Title", "Print Job Author", "Print Job File");
                        // Save the document content with title and author regardless of print decision
                        SaveDocumentContent(documentContent, "Print Job Title", "Print Job Author");
                    }
                }

                client.Close();
                Thread.Sleep(100);
            }

            listener.Stop();
        }

        private void SaveDocumentContent(byte[] content, string title, string author)
        {
            // Determine the path for saving documents (you can adjust this path as needed)
            string documentsPath = Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments);

            // Generate a unique file name or use a timestamp-based name
            string fileName = GenerateFileName(title, author);

            // Combine the path and file name
            string filePath = Path.Combine(documentsPath, fileName);

            // Write the content to the file
            File.WriteAllBytes(filePath, content);

            Console.WriteLine($"Document saved: {filePath}");
        }

        private string ExtractPdfText(byte[] pdfBytes)
        {
            using (MemoryStream ms = new MemoryStream(pdfBytes))
            {
                iTextSharp.text.pdf.PdfReader reader = new iTextSharp.text.pdf.PdfReader(ms);
                StringBuilder sb = new StringBuilder();

                for (int page = 1; page <= reader.NumberOfPages; page++)
                {
                    iTextSharp.text.pdf.parser.ITextExtractionStrategy extractor = new iTextSharp.text.pdf.parser.SimpleTextExtractionStrategy();
                    string pageText = iTextSharp.text.pdf.parser.PdfTextExtractor.GetTextFromPage(reader, page, extractor);
                    sb.AppendLine(pageText);
                }

                reader.Close();
                return sb.ToString();
            }
        }


        private string GenerateFileName(string title, string author)
        {
            // Use current timestamp as default if title or author information is not provided
            string timestamp = DateTime.Now.ToString("yyyyMMddHHmmss");

            // Use title and author if available, otherwise use timestamp
            string fileName = !string.IsNullOrWhiteSpace(title) && !string.IsNullOrWhiteSpace(author)
                ? $"{SanitizeFileName(title)} - {SanitizeFileName(author)}_{timestamp}.pdf"
                : $"Document_{timestamp}.pdf";

            // Replace invalid characters for file names with underscores
            fileName = SanitizeFileName(fileName);

            return fileName;
        }

        private string SanitizeFileName(string fileName)
        {
            // Remove invalid characters from file name
            foreach (char c in Path.GetInvalidFileNameChars())
            {
                fileName = fileName.Replace(c, '_');
            }

            return fileName;
        }
    }
}

using System;
using System.Collections.Generic;
using System.IO;
using System.Net;
using System.Net.Sockets;
using System.Text;
using System.Threading;

namespace VirtualPrinterService
{
    public delegate void PrintCallbackFunctionType(
        string doc,
        string title,
        string author,
        string filename
    );
    public  class PrinterServerService
    {
        private readonly string ?printerName;
        private  string ip;
        private  int port;
        private readonly bool autoInstallPrinter;
        private readonly PrintCallbackFunctionType? printCallbackFn;
        private bool running;
        private bool keepGoing;
        private PrinterManagementService ?osPrinterManager;
        private string? printerPortName;

        public PrinterServerService(
           string printerName = "My Virtual Printer",
           string ip = "127.0.0.1",
           int port = 9001,
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

                if (printCallbackFn == null)
                {
                    using (FileStream fileStream = new FileStream("I_printed_this.ps", FileMode.Create, FileAccess.Write))
                    {
                        stream.CopyTo(fileStream);
                    }
                }
                else
                {
                    List<string> buffer = new List<string>();
                    using (StreamReader reader = new StreamReader(stream, Encoding.UTF8))
                    {
                        string data;
                        while ((data = reader.ReadLine()) != null)
                        {
                            buffer.Add(data);
                        }
                    }

                    string combinedBuffer = string.Join(Environment.NewLine, buffer);
                    string author = null, title = null, filename = null;
                    string header = "@" + combinedBuffer.Split(new[] { "%!PS-" }, StringSplitOptions.None)[0].Split('@')[1];

                    foreach (var line in header.Split('\n'))
                    {
                        var trimmedLine = line.Trim();
                        if (trimmedLine.StartsWith("@PJL JOB NAME="))
                        {
                            var name = trimmedLine.Split('"')[1];
                            if (File.Exists(name))
                            {
                                filename = name;
                            }
                            else
                            {
                                title = name;
                            }
                        }
                        else if (trimmedLine.StartsWith("@PJL COMMENT"))
                        {
                            var parameters = trimmedLine.Split('"')[1].Split(';');
                            foreach (var param in parameters)
                            {
                                var kv = param.Split(':');
                                if (kv.Length > 1)
                                {
                                    kv[0] = kv[0].Trim().ToLower();
                                    kv[1] = kv[1].Trim();
                                    if (kv[0] == "username")
                                    {
                                        author = kv[1];
                                    }
                                    else if (kv[0] == "app filename")
                                    {
                                        if (title == null)
                                        {
                                            if (File.Exists(kv[1]))
                                            {
                                                filename = kv[1];
                                            }
                                            else
                                            {
                                                title = kv[1];
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (title == null && filename != null)
                    {
                        title = Path.GetFileNameWithoutExtension(filename);
                    }

                    printCallbackFn(combinedBuffer, title, author, filename);
                }

                client.Close();
                Thread.Sleep(100);
            }

            listener.Stop();
        }



    }
}

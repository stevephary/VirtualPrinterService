using System;
using System.Diagnostics;

namespace VirtualPrinterService
{
    public class PrinterManagementService
    {
        private string defaultPrinterDriver = "Microsoft Print To PDF";

        public void RemovePort(string printerPortName)
        {
            string cmd = $"Remove-PrinterPort -Name \"{printerPortName}\"";
            ExecutePowerShellCommand(cmd);
        }

        public void RemovePrinter(string name)
        {
            string cmd = $"Remove-Printer -Name \"{name}\"";
            ExecutePowerShellCommand(cmd);
        }

        public void ListPorts()
        {
            string cmd = "Get-PrinterPort";
            ExecutePowerShellCommand(cmd);
        }

        public void MakePrinterDefault(string name)
        {
            string cmd = $"Set-DefaultPrinter -Name \"{name}\"";
            ExecutePowerShellCommand(cmd);
        }

        public void SetPrinterComment(string name, string comment)
        {
            comment = comment.Replace("\"", "\\\"").Replace("\n", "\\n");
            string cmd = $"Set-Printer -Name \"{name}\" -Comment \"{comment}\"";
            ExecutePowerShellCommand(cmd);
        }

        public string AddPrinter(string name, string host = "127.0.0.1", int port = 9100, string printerPortName = null, bool makeDefault = false, string comment = null)
        {
            string portStr = port.ToString();
            if (printerPortName == null)
            {
                printerPortName = $"IP_{host.Replace(".", "_")}_{portStr}";
            }

            // Ensure driver is installed before creating port
            if (!IsPrinterDriverInstalled(defaultPrinterDriver))
            {
                InstallPrinterDriver(defaultPrinterDriver);
            }

            // Check if printer already exists
            string checkPrinterCmd = $"Get-Printer -Name \"{name}\"";
            string checkPrinterOutput = ExecutePowerShellCommand(checkPrinterCmd);
            Console.WriteLine($"Check Printer Output: {checkPrinterOutput}");

            if (checkPrinterOutput.Contains("Name"))
            {
                Console.WriteLine($"Printer '{name}' is already installed.");
                return name;
            }

            // Check if port already exists
            string listPortsCmd = "Get-PrinterPort";
            string listPortsOutput = ExecutePowerShellCommand(listPortsCmd);
            Console.WriteLine($"List Ports Output: {listPortsOutput}");

            if (!listPortsOutput.Contains(printerPortName))
            {
                // Create the printer port
                string createPortCmd = $"Add-PrinterPort -Name \"{printerPortName}\"   -PrinterHostAddress \"{host}\"   -PortNumber {portStr}";
                string portCreationOutput = ExecutePowerShellCommand(createPortCmd);
                Console.WriteLine($"Port Creation Output: {portCreationOutput}");

                // Check if the port was created successfully
                if (portCreationOutput.Contains("successfully") || portCreationOutput.Contains("already exists"))
                {
                    // Install the printer using the newly created port
                    string installPrinterCmd = $"Add-Printer -Name \"{name}\" -DriverName '{defaultPrinterDriver}' -PortName \"{printerPortName}\" ";
                    Console.WriteLine($"Install Printer Command: {installPrinterCmd}");
                    string installPrinterOutput = ExecutePowerShellCommand(installPrinterCmd);
                    Console.WriteLine($"Install Printer Output: {installPrinterOutput}");

                    // Make the printer default if specified
                    if (makeDefault)
                    {
                        MakePrinterDefault(name);
                    }

                    // Set printer comment if specified
                    if (!string.IsNullOrEmpty(comment))
                    {
                        SetPrinterComment(name, comment);
                    }
                }
                else
                {
                    Console.WriteLine("Error creating printer port. Please check the command syntax and network settings.");
                }
            }
            else
            {
                Console.WriteLine("Port already exists, proceeding to install printer.");
                // Install the printer using the existing port
                string installPrinterCmd = $"Add-Printer -Name \"{name}\" -DriverName '{defaultPrinterDriver}' -PortName \"{printerPortName}\" ";
                Console.WriteLine($"Install Printer Command: {installPrinterCmd}");
                string installPrinterOutput = ExecutePowerShellCommand(installPrinterCmd);
                Console.WriteLine($"Install Printer Output: {installPrinterOutput}");

                // Make the printer default if specified
                if (makeDefault)
                {
                    MakePrinterDefault(name);
                }

                // Set printer comment if specified
                if (!string.IsNullOrEmpty(comment))
                {
                    SetPrinterComment(name, comment);
                }
            }

            return name;
        }


        private bool IsPrinterDriverInstalled(string driverName)
        {
            string cmd = $"Get-PrinterDriver | Where-Object {{$_.Name -eq '{driverName}'}}";
            string result = ExecutePowerShellCommand(cmd);
            return !string.IsNullOrEmpty(result);
        }

        private void InstallPrinterDriver(string driverName)
        {
            string driverInfPath = @"C:\Path\To\Your\Driver\Driver.inf";
            string cmd = $"pnputil /add-driver \"{driverInfPath}\" /install";
            ExecuteCommand(cmd);
        }

        public void PrintTestPage(string name)
        {
            string cmd = $"rundll32 printui.dll,PrintUIEntry /k /n {name}";
            ExecuteCommand(cmd);
        }

        public void ShowSettingsDialog(string name)
        {
            string cmd = $"rundll32 printui.dll,PrintUIEntry /e /n {name}";
            ExecuteCommand(cmd);
        }

        public void SaveSettings(string name, string filename)
        {
            string cmd = $"rundll32 printui.dll,PrintUIEntry /Ss /n {name} /a {filename}";
            ExecuteCommand(cmd);
        }

        public void LoadSettings(string name, string filename)
        {
            string cmd = $"rundll32 printui.dll,PrintUIEntry /Sr /n {name} /a {filename}";
            ExecuteCommand(cmd);
        }

        private string ExecuteCommand(string cmd)
        {
            ProcessStartInfo processStartInfo = new ProcessStartInfo("cmd", $"/c {cmd}")
            {
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process process = new Process { StartInfo = processStartInfo })
            {
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();
                Console.WriteLine(output);
                return output;
            }
        }

        private string ExecutePowerShellCommand(string cmd)
        {
            ProcessStartInfo processStartInfo = new ProcessStartInfo("powershell", $"-Command \"{cmd}\"")
            {
                RedirectStandardOutput = true,
                UseShellExecute = false,
                CreateNoWindow = true
            };

            using (Process process = new Process { StartInfo = processStartInfo })
            {
                process.Start();
                string output = process.StandardOutput.ReadToEnd();
                process.WaitForExit();
                Console.WriteLine(output);
                return output;
            }
        }
    }
}

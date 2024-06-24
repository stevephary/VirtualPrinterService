using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using Microsoft.Office.Interop.Word;

namespace VirtualPrinterService
{
    public class DocumentProcessor
    {
        private static readonly Dictionary<string, Action<byte[]>> Processors = new Dictionary<string, Action<byte[]>>
        {
            { ".txt", ProcessTextFile },
            { ".doc", ProcessWordFile },
            { ".docx", ProcessWordFile }
            // Add more mappings for other file types
        };

        public static void ProcessDocument(byte[] data, string extension)
        {
            if (Processors.TryGetValue(extension.ToLower(), out var processor))
            {
                processor(data);
            }
            else
            {
                Console.WriteLine("Received an unsupported document type.");
            }
        }

        private static void ProcessTextFile(byte[] data)
        {
            string textContent = Encoding.UTF8.GetString(data);
            Console.WriteLine("Text Document Content:");
            Console.WriteLine(textContent);
        }

        private static void ProcessWordFile(byte[] data)
        {
            string tempFile = Path.GetTempFileName();
            File.WriteAllBytes(tempFile, data);

            Application wordApp = new Application();
            Document doc = wordApp.Documents.Open(tempFile);
            string wordContent = doc.Content.Text;
            Console.WriteLine("Word Document Content:");
            Console.WriteLine(wordContent);
            doc.Close();
            wordApp.Quit();

            File.Delete(tempFile);
        }

        // Add more methods for processing other document types
    }
}

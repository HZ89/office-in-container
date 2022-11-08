// Requires office.dll and word interop pia.dll from GAC

using Microsoft.Office.Interop.Word;
using System;
using System.IO;
using System.Linq;

namespace OfficePIATest
{
    internal static class Program
    {
        [STAThread]
        public static void Main(string[] args)
        {
            var input = args.ElementAtOrDefault(0) ?? string.Empty;
            var output = args.ElementAtOrDefault(1) ?? string.Empty;

            if (!File.Exists(input))
            {
                Console.Error.WriteLine("Cannot open the file `{0}`.", input);
                Environment.Exit(1);
                return;
            }

            if (string.IsNullOrWhiteSpace(output))
            {
                Console.Error.WriteLine("Invalid output path.");
                Environment.Exit(2);
                return;
            }

            var outputDir = Path.GetDirectoryName(output);
            if (!Directory.Exists(outputDir))
                Directory.CreateDirectory(outputDir);

            var wordApp = new Application();
            Console.Out.WriteLine("word app initialized.");

            Document wordDocument = null;
            var t = new System.Threading.Tasks.Task(() => wordDocument = wordApp.Documents.Open(input));
            t.Start();
            t.Wait();

            if (wordDocument == null)
            {
                Console.Error.WriteLine("Word document is null reference.");
                Environment.Exit(3);
                return;
            }
            else
                Console.Out.WriteLine("Word document opened.");

            wordDocument.ExportAsFixedFormat(output, WdExportFormat.wdExportFormatPDF);
            Console.Out.WriteLine("Export performed.");

            wordDocument.Close(WdSaveOptions.wdDoNotSaveChanges,
                               WdOriginalFormat.wdOriginalDocumentFormat,
                               false); //Close document
            Console.Write("Closing document");

            wordApp.Quit(); //Important: When you forget this Word keeps running in the background
            Console.Out.WriteLine("Quitting word app.");
        }
    }
}
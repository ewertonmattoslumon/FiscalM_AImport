using System;
using System.IO;
using System.Linq;
using FiscalM_AImport.Importers;
using FiscalM_AImport.Models;
using Microsoft.Extensions.Configuration;
using Microsoft.PowerPlatform.Dataverse.Client;

namespace FiscalM_AImport
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var config = new ConfigurationBuilder()
                .SetBasePath(AppDomain.CurrentDomain.BaseDirectory)
                .AddJsonFile("appsettings.json", optional: false, reloadOnChange: false)
                .AddJsonFile("appsettings.local.json", optional: true, reloadOnChange: false)
                .Build();

            var settings = config.Get<AppSettings>()
                ?? throw new InvalidOperationException("Failed to load appsettings.json.");

            Console.WriteLine("FiscalM AImport - Dynamics 365 Excel Importer");
            Console.WriteLine("==============================================");
            Console.WriteLine();
            //test
            ServiceClient serviceClient;
            try
            {
                serviceClient = new ServiceClient(settings.Dataverse.ConnectionString);
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to initialise Dataverse client: {ex.Message}");
                return;
            }

            using (serviceClient)
            {
                if (!serviceClient.IsReady)
                {
                    Console.WriteLine($"Failed to connect to Dynamics 365: {serviceClient.LastError}");
                    return;
                }

                var url = settings.Dataverse.ConnectionString
                    .Split(';')
                    .Select(p => p.Trim())
                    .FirstOrDefault(p => p.StartsWith("Url=", StringComparison.OrdinalIgnoreCase))
                    ?.Substring(4) ?? "(unknown)";

                Console.WriteLine($"Connected to: {url}");
                Console.Write("Proceed with import? (Y/N): ");
                var answer = Console.ReadLine()?.Trim();
                if (!string.Equals(answer, "Y", StringComparison.OrdinalIgnoreCase))
                {
                    Console.WriteLine("Import aborted.");
                    return;
                }
                Console.WriteLine();

                // Excel files are expected in the current working directory.
                // When running via 'dotnet run', that is the project folder.
                // When running the compiled exe, place the files next to it or run from that folder.
                var baseDir = Directory.GetCurrentDirectory();

                var importer = new TripleImporter(
                    serviceClient,
                    baseDir,
                    settings.Import.ExcelFile,
                    settings.Import.FieldNamesRow);

                try
                {
                    importer.Import();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"Unexpected error during import: {ex.Message}");
                }

                Console.WriteLine("Import completed.");
            }
        }
    }
}

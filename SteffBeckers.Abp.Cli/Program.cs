using DocumentFormat.OpenXml.Packaging;
using Newtonsoft.Json;
using SteffBeckers.Abp.Cli.Localization.Models;
using System;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace SteffBeckers.Abp.Cli
{
    public class Program
    {
        public static async Task<int> Main(string[] args)
        {
            RootCommand rootCommand = new RootCommand("Steff's ABP.io CLI");

            Command localizationCommand = new Command("localization", "Localization commands.");
            Command localizationExportCommand = new Command("export", "Export localization files to other formats.");
            Command localizationExportExcelCommand = new Command("excel", "Export to Excel.");

            localizationExportExcelCommand.Handler = CommandHandler.Create(async () =>
            {
                List<LocalizationFile> localizationFiles = await GetLocalizationFiles();

                string directoryName = new DirectoryInfo(Path.GetDirectoryName(localizationFiles.First().Path)).Name;
                string spreadsheetFileName = $"{directoryName}.xslx";

                await File.WriteAllTextAsync(spreadsheetFileName, "Test");
            });

            localizationExportCommand.AddCommand(localizationExportExcelCommand);
            localizationCommand.AddCommand(localizationExportCommand);
            rootCommand.AddCommand(localizationCommand);

            return await rootCommand.InvokeAsync(args);
        }

        private static async Task<List<LocalizationFile>> GetLocalizationFiles()
        {
            List<LocalizationFile> localizationFiles = new List<LocalizationFile>();

            List<string> localizationFilePaths = Directory.GetFiles(Directory.GetCurrentDirectory())
                .Where(x => x.EndsWith(".json"))
                .ToList();

            foreach (string localizationFilePath in localizationFilePaths)
            {
                string localizationFileJson = await File.ReadAllTextAsync(localizationFilePath);

                LocalizationFile localizationFile = JsonConvert.DeserializeObject<LocalizationFile>(localizationFileJson);
                if (localizationFile.Culture == null)
                {
                    continue;
                }

                localizationFile.Path = localizationFilePath;

                localizationFiles.Add(localizationFile);
            }

            if (localizationFiles.Count == 0)
            {
                throw new Exception("No localization files found.");
            }

            return localizationFiles;
        }
    }
}
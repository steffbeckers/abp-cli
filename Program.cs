using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Newtonsoft.Json;
using SteffBeckers.Abp.Cli.Localization.Models;
using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.CommandLine;
using System.CommandLine.Invocation;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace SteffBeckers.Abp.Cli
{
    public class Program
    {
        public static async Task<int> Main(string[] args)
        {
            RootCommand command = new RootCommand("Steff's ABP.io CLI");
            AddLocalizationCommand(command);
            return await command.InvokeAsync(args);
        }

        private static void AddLocalizationCommand(RootCommand parentCommand)
        {
            Command command = new Command("localization", "Localization commands.");
            AddLocalizationExportCommand(command);
            AddLocalizationImportCommand(command);
            AddLocalizationScanCommand(command);
            parentCommand.AddCommand(command);
        }

        private static void AddLocalizationExportCommand(Command parentCommand)
        {
            Command command = new Command("export", "Export localization files to other formats.");
            AddLocalizationExportExcelCommand(command);
            parentCommand.AddCommand(command);
        }

        private static void AddLocalizationExportExcelCommand(Command parentCommand)
        {
            Command command = new Command("excel", "Export to Excel.");

            command.Handler = CommandHandler.Create(async () =>
            {
                List<LocalizationFile> localizationFiles = await GetLocalizationFiles();

                string directoryName = new DirectoryInfo(Directory.GetCurrentDirectory()).Name;
                string spreadsheetFileName = $"{directoryName}.xlsx";

                // Create a spreadsheet document by supplying the filepath
                // By default, AutoSave = true, Editable = true, and Type = xlsx
                SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Create(spreadsheetFileName, SpreadsheetDocumentType.Workbook);

                // Add a WorkbookPart to the document
                WorkbookPart workbookPart = spreadsheetDocument.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // Add a WorksheetPart to the WorkbookPart
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                worksheetPart.Worksheet = new Worksheet(new SheetData());

                // Add Sheets to the Workbook
                Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

                // Append a new worksheet and associate it with the workbook
                Sheet sheet = new Sheet()
                {
                    Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Localizations"
                };
                sheets.Append(sheet);

                // Add data to worksheet.
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Add header row.
                Row headerRow = new Row();
                headerRow.Append(new Cell()
                {
                    DataType = CellValues.String,
                    CellValue = new CellValue("key")
                });
                foreach (LocalizationFile localizationFile in localizationFiles)
                {
                    headerRow.Append(new Cell()
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(localizationFile.Culture)
                    });
                }
                sheetData.Append(headerRow);

                // Add rows.
                List<string> distinctLocalizationKeys = localizationFiles
                    .SelectMany(x => x.Texts.Select(y => y.Key))
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList();

                foreach (string localizationKey in distinctLocalizationKeys)
                {
                    Row row = new Row();
                    row.Append(new Cell()
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(localizationKey)
                    });

                    foreach (LocalizationFile localizationFile in localizationFiles)
                    {
                        localizationFile.Texts.TryGetValue(localizationKey, out string localizationValue);
                        if (!string.IsNullOrEmpty(localizationValue))
                        {
                            row.Append(new Cell()
                            {
                                DataType = CellValues.String,
                                CellValue = new CellValue(localizationValue)
                            });
                        }
                        else
                        {
                            row.Append(new Cell());
                        }
                    }

                    sheetData.Append(row);
                }

                // Save the Workbook
                workbookPart.Workbook.Save();

                // Close the document
                spreadsheetDocument.Close();

                Console.WriteLine($"Exported localizations to '{spreadsheetFileName}'.");
            });

            parentCommand.AddCommand(command);
        }

        private static void AddLocalizationImportCommand(Command parentCommand)
        {
            Command command = new Command("import", "Import localization files from diverse formats.");
            AddLocalizationImportExcelCommand(command);
            parentCommand.AddCommand(command);
        }

        private static void AddLocalizationImportExcelCommand(Command parentCommand)
        {
            Command command = new Command("excel", "Import from Excel.");

            command.Handler = CommandHandler.Create(async () =>
            {
                string directoryName = new DirectoryInfo(Directory.GetCurrentDirectory()).Name;
                string spreadsheetFileName = $"{directoryName}.xlsx";

                // New localization files to be generated
                List<LocalizationFile> localizationFiles = new List<LocalizationFile>();

                // Open the spreadsheet document by supplying the filepath
                SpreadsheetDocument spreadsheetDocument = SpreadsheetDocument.Open(spreadsheetFileName, false);
                WorkbookPart workbookPart = spreadsheetDocument.WorkbookPart;
                WorksheetPart worksheetPart = workbookPart.WorksheetParts.First();
                SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();

                // Header row
                foreach (Cell cell in sheetData.Elements<Row>().First().Elements<Cell>().Skip(1))
                {
                    string culture = ReadExcelCell(cell, workbookPart);
                    localizationFiles.Add(new LocalizationFile()
                    {
                        Path = $"{culture}.json",
                        Culture = culture,
                        Texts = new Dictionary<string, string>()
                    });
                }

                // Data rows
                foreach (Row row in sheetData.Elements<Row>().Skip(1))
                {
                    string localizationKey = null;
                    foreach (Cell cell in row.Elements<Cell>())
                    {
                        int? columnIndex = GetColumnIndex(cell);
                        string text = ReadExcelCell(cell, workbookPart);

                        if (columnIndex.HasValue && columnIndex.Value == 0)
                        {
                            localizationKey = text;
                            continue;
                        }
                        else if (string.IsNullOrEmpty(localizationKey))
                        {
                            continue;
                        }

                        localizationFiles[columnIndex.Value - 1].Texts.TryAdd(localizationKey, text);
                    }
                }

                // Close the document
                spreadsheetDocument.Close();

                // Write to .json files
                foreach (LocalizationFile localizationFile in localizationFiles)
                {
                    // Sort localizations
                    localizationFile.Texts = localizationFile.Texts
                        .OrderBy(x => x.Key)
                        .ToDictionary(x => x.Key, x => x.Value);

                    string localizationFileJson = JsonConvert.SerializeObject(localizationFile, Formatting.Indented);
                    await File.WriteAllTextAsync(localizationFile.Path, localizationFileJson);
                }

                Console.WriteLine($"Imported localizations to localization files.");
            });

            parentCommand.AddCommand(command);
        }

        private static void AddLocalizationScanCommand(Command parentCommand)
        {
            Command command = new Command("scan", "Scan's all files in current folder based on localization keys.");

            Option localizationsPathOption = new Option<DirectoryInfo>(
                "--localization-files-path",
                "Directory path of localization files."
            )
            {
                IsRequired = true
            };
            localizationsPathOption.AddAlias("-lfp");
            command.AddOption(localizationsPathOption);

            command.Handler = CommandHandler.Create<DirectoryInfo>(async (localizationFilesPath) =>
            {
                List<LocalizationFile> localizationFiles = await GetLocalizationFiles(localizationFilesPath);

                List<string> distinctLocalizationKeys = localizationFiles
                    .SelectMany(x => x.Texts.Select(y => y.Key))
                    .Distinct()
                    .OrderBy(x => x)
                    .ToList();

                DirectoryInfo currentDirectory = new DirectoryInfo(Directory.GetCurrentDirectory());
                IEnumerable<FileInfo> files = currentDirectory.EnumerateFiles("*.*", SearchOption.AllDirectories)
                    .Where(x => !x.DirectoryName.Contains("node_modules"))
                    .Where(x => !x.DirectoryName.Contains("Migrations"))
                    .Where(x =>
                        x.Name.EndsWith(".cs") ||
                        x.Name.EndsWith(".tpl") ||
                        x.Name.EndsWith(".html") ||
                        x.Name.EndsWith(".ts"))
                    .Where(x => !x.Name.EndsWith(".Designer.cs"))
                    .Where(x => !x.Name.EndsWith(".g.cs"))
                    .OrderBy(x => x.FullName);

                Console.WriteLine("Localization scan started.");
                Console.WriteLine();

                ConcurrentBag<string> foundLocalizationKeys = new ConcurrentBag<string>();
                Parallel.ForEach(files, async (file) =>
                {
                    Console.WriteLine($"Scanning: {file.FullName}");

                    string fileContent = await File.ReadAllTextAsync(file.FullName);

                    foreach (string localizationKey in distinctLocalizationKeys)
                    {
                        if (fileContent.Contains(localizationKey))
                        {
                            foundLocalizationKeys.Add(localizationKey);
                        }
                    }
                });

                Console.WriteLine();
                Console.WriteLine($"Localization scan stopped. {files.Count()} files searched.");

                List<string> notFoundLocalizationKeys = distinctLocalizationKeys
                    .Where(x => !foundLocalizationKeys.Any(y => y == x))
                    .ToList();

                Console.WriteLine();
                Console.WriteLine($"{notFoundLocalizationKeys.Count()} localization keys which were not found in files:");

                foreach (string localizationKey in notFoundLocalizationKeys)
                {
                    Console.WriteLine(localizationKey);
                }
            });

            parentCommand.AddCommand(command);
        }

        private static async Task<List<LocalizationFile>> GetLocalizationFiles(DirectoryInfo directoryInfo = null)
        {
            List<LocalizationFile> localizationFiles = new List<LocalizationFile>();

            if (directoryInfo == null)
            {
                directoryInfo = new DirectoryInfo(Directory.GetCurrentDirectory());
            }

            List<string> localizationFilePaths = Directory.GetFiles(directoryInfo.FullName)
                .Where(x => x.EndsWith(".json"))
                .OrderBy(x => x)
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

        private static string ReadExcelCell(Cell cell, WorkbookPart workbookPart)
        {
            CellValue cellValue = cell.CellValue;

            string text = (cellValue == null) ? cell.InnerText : cellValue.Text;

            if ((cell.DataType != null) && (cell.DataType == CellValues.SharedString))
            {
                text = workbookPart.SharedStringTablePart.SharedStringTable
                    .Elements<SharedStringItem>().ElementAt(
                        Convert.ToInt32(cell.CellValue.Text)).InnerText;
            }

            return (text ?? string.Empty).Trim();
        }

        private static int? GetColumnIndex(Cell cell)
        {
            string cellReference = cell.CellReference;

            if (string.IsNullOrEmpty(cellReference))
            {
                return null;
            }

            // Remove digits
            string columnReference = Regex.Replace(cellReference.ToUpper(), @"[\d]", string.Empty);

            int columnNumber = -1;
            int mulitplier = 1;

            // Working from the end of the letters take the ASCII code less 64 (so A = 1, B =2...etc)
            // Then multiply that number by our multiplier (which starts at 1)
            // Multiply our multiplier by 26 as there are 26 letters
            foreach (char c in columnReference.ToCharArray().Reverse())
            {
                columnNumber += mulitplier * ((int)c - 64);
                mulitplier = mulitplier * 26;
            }

            // This will match Excel's COLUMN function
            return columnNumber;
        }
    }
}
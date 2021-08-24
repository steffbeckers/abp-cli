using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
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
        private static Option VerboseOption;

        public static async Task<int> Main(string[] args)
        {
            RootCommand command = new RootCommand("Steff's ABP.io CLI");

            VerboseOption = new Option<bool>("--verbose");
            VerboseOption.AddAlias("-v");
            command.AddOption(VerboseOption);

            AddLocalizationCommand(command);

            return await command.InvokeAsync(args);
        }

        private static void AddLocalizationCommand(RootCommand parentCommand)
        {
            Command command = new Command("localization", "Localization commands.");
            AddLocalizationExportCommand(command);
            AddLocalizationImportCommand(command);
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
            command.AddOption(VerboseOption);

            command.Handler = CommandHandler.Create<bool>(async (verbose) =>
            {
                try
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

                    #region TODO

                    //TableDefinitionPart tableDefinitionPart = worksheetPart.AddNewPart<TableDefinitionPart>("rId" + (worksheetPart.TableDefinitionParts.Count() + 1));
                    //int tableNumber = worksheetPart.TableDefinitionParts.Count();
                    //int colMin = 1;
                    //int colMax = localizationFiles.Count + 1;
                    //int rowMin = 1;
                    //int rowMax = localizationFiles.Max(x => x.Texts.Count) + 1;
                    //string reference = ((char)(64 + colMin)).ToString() + rowMin + ":" + ((char)(64 + colMax)).ToString() + rowMax;

                    //Table table = new Table()
                    //{
                    //    Id = (uint)tableNumber,
                    //    Name = "Table" + tableNumber,
                    //    DisplayName = "Table" + tableNumber,
                    //    Reference = reference,
                    //    TotalsRowShown = false
                    //};
                    //AutoFilter autoFilter = new AutoFilter()
                    //{
                    //    Reference = reference
                    //};

                    //TableColumns tableColumns = new TableColumns()
                    //{
                    //    Count = (uint)(colMax - colMin + 1)
                    //};

                    //foreach (LocalizationFile localizationFile in localizationFiles)
                    //{
                    //    int localizationFileIndex = localizationFiles.IndexOf(localizationFile);

                    //    tableColumns.Append(new TableColumn()
                    //    {
                    //        Id = (uint)(localizationFileIndex),
                    //        Name = localizationFile.Culture
                    //    });
                    //}

                    //TableStyleInfo tableStyleInfo = new TableStyleInfo()
                    //{
                    //    Name = "TableStyleLight1",
                    //    ShowFirstColumn = false,
                    //    ShowLastColumn = false,
                    //    ShowRowStripes = true,
                    //    ShowColumnStripes = false
                    //};

                    //table.Append(autoFilter);
                    //table.Append(tableColumns);
                    //table.Append(tableStyleInfo);

                    //tableDefinitionPart.Table = table;

                    #endregion TODO

                    // Save the Workbook
                    workbookPart.Workbook.Save();

                    // Close the document
                    spreadsheetDocument.Close();

                    Console.WriteLine($"Exported localizations to '{spreadsheetFileName}'.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);

                    string innerExceptionMessage = ex.InnerException?.Message;
                    if (verbose && !string.IsNullOrEmpty(innerExceptionMessage))
                    {
                        Console.WriteLine("ERROR:");
                        Console.WriteLine(innerExceptionMessage);
                    }
                }
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
            command.AddOption(VerboseOption);

            command.Handler = CommandHandler.Create<bool>(async (verbose) =>
            {
                try
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
                        List<Cell> cells = row.Elements<Cell>().ToList();
                        string localizationKey = null;

                        foreach (Cell cell in cells)
                        {
                            int cellIndex = cells.IndexOf(cell);
                            string text = ReadExcelCell(cell, workbookPart);

                            if (cellIndex == 0)
                            {
                                localizationKey = text;
                                continue;
                            }

                            if (!string.IsNullOrEmpty(localizationKey))
                            {
                                localizationFiles[cellIndex - 1].Texts.TryAdd(localizationKey, text);
                            }
                        }
                    }

                    // Close the document
                    spreadsheetDocument.Close();

                    // Write to .json files
                    foreach (LocalizationFile localizationFile in localizationFiles)
                    {
                        string localizationFileJson = JsonConvert.SerializeObject(localizationFile, Formatting.Indented);
                        await File.WriteAllTextAsync(localizationFile.Path, localizationFileJson);
                    }

                    Console.WriteLine($"Imported localizations to localization files.");
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);

                    string innerExceptionMessage = ex.InnerException?.Message;
                    if (verbose && !string.IsNullOrEmpty(innerExceptionMessage))
                    {
                        Console.WriteLine("ERROR:");
                        Console.WriteLine(innerExceptionMessage);
                    }
                }
            });

            parentCommand.AddCommand(command);
        }

        private static async Task<List<LocalizationFile>> GetLocalizationFiles()
        {
            List<LocalizationFile> localizationFiles = new List<LocalizationFile>();

            List<string> localizationFilePaths = Directory.GetFiles(Directory.GetCurrentDirectory())
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
    }
}
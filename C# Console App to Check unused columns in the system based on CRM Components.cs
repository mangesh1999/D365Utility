using System;
using System.Data;
using System.IO;
using System.Linq;
using Microsoft.Xrm.Sdk;
using Microsoft.PowerPlatform.Dataverse.Client;
using Microsoft.Xrm.Sdk.Query;
using ExcelDataReader;
using System.Text.RegularExpressions;
using System.Configuration;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace Validate_Field
{
    class Program
    {
        static void Main(string[] args)
        {
            // Fetching parameters from app.config
            string url = ConfigurationManager.AppSettings["Url"];
            string clientId = ConfigurationManager.AppSettings["ClientId"];
            string clientSecret = ConfigurationManager.AppSettings["ClientSecret"];
            string authority = ConfigurationManager.AppSettings["Authority"];

            string connectionString = $"AuthType=ClientSecret;Url={url};ClientId={clientId};ClientSecret={clientSecret};Authority={authority}";

            // Formfetch XML query
            var Formfetch = $@"<?xml version=""1.0"" encoding=""utf-16""?>
									<fetch>
									  <entity name=""systemform"">
										<attribute name=""formjson"" />
										<attribute name=""formxml"" />
										<attribute name=""objecttypecode"" />
										<attribute name=""formid"" />
										<attribute name=""name"" />
									  </entity>
									</fetch>";

            // Viewfetch XML query
            var ViewFetch = $@"<?xml version=""1.0"" encoding=""utf-16""?>
									<fetch>
									  <entity name=""savedquery"">
										<attribute name=""fetchxml"" />
										<attribute name=""returnedtypecode"" />
										<attribute name=""name"" />
										<attribute name=""savedqueryid"" />
									  </entity>
									</fetch>";

            string excelFilePath = @"Your Excel File Path";

            DataTable fieldsTable = ReadExcelData(excelFilePath);

            using (ServiceClient serviceClient = new ServiceClient(connectionString))
            {
                if (serviceClient.IsReady)
                {
                    Console.WriteLine("Connected to Dynamics 365!");

                    IOrganizationService crmService = (IOrganizationService)serviceClient;
                    if (crmService != null)
                    {
                        Console.WriteLine("IOrganizationService Created !!");

                        var resultFormFetch = crmService.RetrieveMultiple(new FetchExpression(Formfetch));
                        var resultViewFetch = crmService.RetrieveMultiple(new FetchExpression(ViewFetch));

                        DataTable resultTable = new DataTable();
                        resultTable.Columns.AddRange(new[]
                        {
                            new DataColumn("Table Name"),
                            new DataColumn("Attribute Name"),
                            new DataColumn("Is Used"),
                            new DataColumn("Found In"),  // Where it was found (Form/View)
                            new DataColumn("Form/View Name"), // Name of the form/view
                            new DataColumn("Form/View ID")    // ID of the form/view
                        });

                        int totalAttributes = fieldsTable.Rows.Count;
                        int checkCounter = 0;
                        string pattern = "[^a-zA-Z0-9_]"; // Regex pattern to sanitize XML content

                        // Iterate over each attribute and check usage in Formfetch and Viewfetch
                        foreach (DataRow field in fieldsTable.Rows)
                        {
                            string tableName = field["TableName"].ToString();
                            string attributeName = field["logicalname"].ToString();
                            string foundIn = "Not Found";  // Tracks where the attribute was found
                            string formViewName = "";
                            string formViewId = "";

                            // Pre-filter Formfetch results based on tableName (objecttypecode)
                            var filteredForms = resultFormFetch.Entities
                                .Where(f => f.GetAttributeValue<string>("objecttypecode") == tableName)
                                .ToList();

                            // Pre-filter Viewfetch results based on tableName (returnedtypecode)
                            var filteredViews = resultViewFetch.Entities
                                .Where(v => v.GetAttributeValue<string>("returnedtypecode") == tableName)
                                .ToList();

                            // Checking in filtered Formfetch
                            bool isUsedInForm = filteredForms.Any(f =>
                            {
                                string formXml = f.GetAttributeValue<string>("formxml") ?? string.Empty;
                                string sanitizedFormXml = Regex.Replace(formXml, pattern, string.Empty);
                                return sanitizedFormXml.Contains(attributeName);
                            });

                            // Checking in filtered Viewfetch
                            bool isUsedInView = filteredViews.Any(v =>
                            {
                                string fetchXml = v.GetAttributeValue<string>("fetchxml") ?? string.Empty;
                                string sanitizedFetchXml = Regex.Replace(fetchXml, pattern, string.Empty);
                                return sanitizedFetchXml.Contains(attributeName);
                            });

                            // Determine where the attribute was found and fetch form/view details
                            if (isUsedInForm)
                            {
                                foundIn = "FormFetch";
                                var form = filteredForms.First(f =>
                                {
                                    string formXml = f.GetAttributeValue<string>("formxml") ?? string.Empty;
                                    return Regex.Replace(formXml, pattern, string.Empty).Contains(attributeName);
                                });
                                formViewName = form.GetAttributeValue<string>("name");
                                formViewId = form.GetAttributeValue<Guid>("formid").ToString();
                            }
                            else if (isUsedInView)
                            {
                                foundIn = "ViewFetch";
                                var view = filteredViews.First(v =>
                                {
                                    string fetchXml = v.GetAttributeValue<string>("fetchxml") ?? string.Empty;
                                    return Regex.Replace(fetchXml, pattern, string.Empty).Contains(attributeName);
                                });
                                formViewName = view.GetAttributeValue<string>("name");
                                formViewId = view.GetAttributeValue<Guid>("savedqueryid").ToString();
                            }

                            bool isUsed = isUsedInForm || isUsedInView;

                            // Add result to the resultTable
                            resultTable.Rows.Add(tableName, attributeName, isUsed ? "Used" : "Not Found", foundIn, formViewName, formViewId);

                            // Increment the check counter
                            checkCounter++;

                            // Log progress
                            Console.WriteLine($"{checkCounter}/{totalAttributes} attributes checked. Attribute: {attributeName} | Found: {foundIn}");

                            // Log progress every 500 attributes
                            if (checkCounter % 500 == 0)
                                Console.WriteLine($"{checkCounter} attributes processed so far...");

                            // Log pending attributes
                            int pendingAttributes = totalAttributes - checkCounter;
                            Console.WriteLine($"Pending attributes: {pendingAttributes}");
                        }

                        // Save the results to an Excel file
                        string outputFilePath = @"Your output file path";
                        SaveExcelFile(outputFilePath, resultTable);
                        Console.WriteLine($"Results saved to: {outputFilePath}");
                    }
                }
            }
        }

        static DataTable ReadExcelData(string filePath)
        {
            DataTable dataTable = new DataTable();

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    var config = new ExcelDataSetConfiguration
                    {
                        ConfigureDataTable = _ => new ExcelDataTableConfiguration
                        {
                            UseHeaderRow = true
                        }
                    };

                    var dataSet = reader.AsDataSet(config);
                    dataTable = dataSet.Tables[0]; // Assuming the first sheet is the relevant one
                }
            }

            return dataTable;
        }

        static void SaveExcelFile(string filePath, DataTable table)
        {
            using (SpreadsheetDocument document = SpreadsheetDocument.Create(filePath, SpreadsheetDocumentType.Workbook))
            {
                // Create the workbook
                WorkbookPart workbookPart = document.AddWorkbookPart();
                workbookPart.Workbook = new Workbook();

                // Create the worksheet
                WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                Worksheet worksheet = new Worksheet();
                worksheetPart.Worksheet = worksheet;

                // Create the sheet data
                SheetData sheetData = new SheetData();
                worksheet.AppendChild(sheetData);

                // Create the header row
                Row headerRow = new Row();
                foreach (DataColumn column in table.Columns)
                {
                    Cell cell = new Cell
                    {
                        DataType = CellValues.String,
                        CellValue = new CellValue(column.ColumnName)
                    };
                    headerRow.AppendChild(cell);
                }
                sheetData.AppendChild(headerRow);

                // Add data rows
                foreach (DataRow row in table.Rows)
                {
                    Row newRow = new Row();
                    foreach (var cellValue in row.ItemArray)
                    {
                        Cell cell = new Cell
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(cellValue.ToString())
                        };
                        newRow.AppendChild(cell);
                    }
                    sheetData.AppendChild(newRow);
                }

                // Create the Sheets collection
                Sheets sheets = document.WorkbookPart.Workbook.AppendChild(new Sheets());

                // Create a single sheet and append to workbook
                Sheet sheet = new Sheet
                {
                    Id = document.WorkbookPart.GetIdOfPart(worksheetPart),
                    SheetId = 1,
                    Name = "Results"
                };
                sheets.Append(sheet);

                workbookPart.Workbook.Save();
            }
        }
    }
}

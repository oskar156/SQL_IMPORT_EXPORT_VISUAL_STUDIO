/* TABLE OF CONTENTS
 * 
 * METHODS
 * public List<string> GetFilesToImport(string ImportPath, string Extension, bool ConvertExcelToCsv = false, bool ConsoleOutput = true)
 * public DataTable ReadFileIntoDataTableWithColumns(string FilePath, string Delimiter, bool DoubleQuoted, System.Text.Encoding EncodingChoice, bool StripNonPrintableChars, int FieldLimit, string FixedWidthColumnFilePath = "", bool ConsoleOutput = true)
 * public List<DataTable> ReadExcelFileIntoDataTablesWithColumns(string FilePath, ref List<string> TableNames, System.Text.Encoding EncodingChoice, bool StripNonPrintableChars, int FieldLimit, int HeaderRowIndex = 0, bool ConsoleOutput = true)
 * public void CreateTablesInSqlServerVarchar(List<string> TableNames, ConnectionInfo ConnectionInfo, List<DataTable> DataTables, string Delimiter, System.Text.Encoding EncodingChoice, bool StripNonPrintableChars, int FieldLimit, string ColumnTypeMethod = "DEFAULT VARCHAR", string ColumnTypeFilePath = "", bool ConsoleOutput = true)
 * public void ReadFileIntoDataTableWithRowsAndInsertIntoSqlServerTable(string FilePath, string TableName, ConnectionInfo ConnectionInfo, DataTable BaseDtTable, int BatchLimit, string Delimiter, bool DoubleQuoted, System.Text.Encoding EncodingChoice, bool StripNonPrintableChars, int FieldLimit, bool ConsoleOutput = true)
 * public void ReadFileIntoDataTableWithRowsAndInsertIntoSnowflakeTable(string FilePath, string TableName, ConnectionInfo ConnectionInfo, Snowflake Snowflake, DataTable BaseDtTable, int BatchLimit, string Delimiter, bool DoubleQuoted, System.Text.Encoding EncodingChoice, string EncodingChoiceSnowflake, bool StripNonPrintableChars, int FieldLimit, string ColumnTypeMethod = "DEFAULT VARCHAR", string ColumnTypeFilePath = "", bool ImportToExistingTable = false, bool ConsoleOutput = true)
 * public void ReadExcelFilePerSheetIntoDataTablesWithRowsAndInsertIntoSqlServerTables(string FilePath, List<string> TableNames, ConnectionInfo ConnectionInfo, List<DataTable> DataTables, int BatchLimit, string Delimeter, System.Text.Encoding EncodingChoice, bool StripNonPrintableChars, int FieldLimit, int HeaderRowIndex = 0, bool ConsoleOutput = true)
 * public void InsertDataTableUsingSqlBulkCopy(ref string ConnString, ref string TableName, ref DataTable TempDataTable, ref int RowIndex, bool ConsoleOutput = true)
 * 
*/

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.VisualBasic.FileIO; //for TextFieldParser (also right click project > add > references > Microsoft.VisualBasic.FileIO)
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.VisualBasic;
using System.Collections.ObjectModel;
using System.Text.RegularExpressions;

namespace SQL_SERVER_IMPORT_EXPORT
{
    public class Import
    {
        //======================================================================================================================
        // METHODS
        //======================================================================================================================
        public List<string> GetFilesToImport(string ImportPath, string Extension, bool ConvertExcelToCsv = false, string ImportRegexFilter = "", bool ConsoleOutput = true)
        {
            List<string> FilesToImport = new List<string>();

            //detect whether its a directory or file
            bool IsImportPathAFile = true;
            FileAttributes Attr = File.GetAttributes(ImportPath);
            if ((Attr & FileAttributes.Directory) == FileAttributes.Directory)
            {
                IsImportPathAFile = false;
            }

            //if ImportPath is a file, import only that file
            if (IsImportPathAFile)
            {
                //CHECK IF EXT IS XLS AND WE ARE CONVERTING TO CSV
                //IF YES THEN CONVERT TO CSV, ADD EACH OF THE RESULTING FILES
                if (Extension == "xls*" && ConvertExcelToCsv == true)
                {
                    Console.WriteLine("Converting " + ImportPath + " (per tab) to csv...");
                    string FilePathNoName = Path.GetDirectoryName(ImportPath);
                    string FileNameNoExt = Path.GetFileNameWithoutExtension(ImportPath);

                    Microsoft.Office.Interop.Excel.Application App = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook Wb = App.Workbooks.Open(@"" + ImportPath);

                    foreach (Microsoft.Office.Interop.Excel.Worksheet Ws in Wb.Worksheets)
                    {
                        string ConvertedFilePath = @"" + FilePathNoName + FileNameNoExt + "_" + Ws.Name + ".csv";
                        Ws.SaveAs(ConvertedFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows);
                        FilesToImport.Add(ConvertedFilePath);
                    }
                    Wb.Close(false);
                    App.Quit();
                }
                else
                {
                    FilesToImport.Add(ImportPath);
                }
                if (ConsoleOutput) { Console.WriteLine("Getting file to import from " + ImportPath + " ..."); }
            }
            //else (ImportPath must be a direcotry), import only every file in that directory that matches Extension
            else
            {
                if (ConsoleOutput) { Console.WriteLine("Getting file(s) to import from " + ImportPath + "*." + Extension + "..."); }
                string[] FilesInPath = Directory.GetFiles(ImportPath, "*." + Extension);

                //foreach (string File in FilesInPath)
                //{
                //    Console.WriteLine(File);
                //}
                //    Console.WriteLine(FilesInPath.Length);
                //CHECK IF EXT IS XLS AND WE ARE CONVERTING TO CSV
                //IF YES THEN CONVERT EACH TO CSV, ADD EACH OF THE RESULTING FILES
                Regex re = new Regex("");
                if (ImportRegexFilter != "")
                {
                    re = new Regex("(?i)" + ImportRegexFilter); //(?i) makes it case-insensitive
                }

                foreach (string File in FilesInPath)
                {
                    string FilePathNoName = Path.GetDirectoryName(File);
                    string FileNameNoExt = Path.GetFileNameWithoutExtension(File);

                    if (re.IsMatch(FileNameNoExt))
                    {

                        if (Extension == "xls*" && ConvertExcelToCsv == true)
                        {
                            //https://stackoverflow.com/questions/2418270/c-sharp-get-a-list-of-files-excluding-those-that-are-hidden
                            //implement a proper hidden file exclusion!
                            if (!FileNameNoExt.StartsWith("~$"))
                            {
                                Console.WriteLine("Converting " + File + " (per tab) to csv...");
                                Microsoft.Office.Interop.Excel.Application App = new Microsoft.Office.Interop.Excel.Application();
                                Microsoft.Office.Interop.Excel.Workbook Wb = App.Workbooks.Open(@"" + File);

                                foreach (Microsoft.Office.Interop.Excel.Worksheet Ws in Wb.Worksheets)
                                {
                                    string ConvertedFilePath = @"" + FilePathNoName + "\\" + FileNameNoExt + "_" + Ws.Name + ".csv";
                                    Ws.SaveAs(ConvertedFilePath, Microsoft.Office.Interop.Excel.XlFileFormat.xlCSVWindows);
                                    FilesToImport.Add(ConvertedFilePath);
                                }
                                Wb.Close(false);
                                App.Quit();
                            }
                        }
                        else
                        {
                            FilesToImport.Add(File);
                        }
                    }
                }
            }

            if (ConsoleOutput)
            {
                Console.WriteLine(FilesToImport.Count.ToString() + " files found");
                Console.WriteLine("");
            }

            return FilesToImport;
        }        
        public DataTable ReadFileIntoDataTableWithColumns(string FilePath, string Delimiter, bool DoubleQuoted, System.Text.Encoding EncodingChoice, bool StripNonPrintableChars, int FieldLimit, string FixedWidthColumnFilePath = "", string ColumnTypeFilePath = "", bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.Write("Reading file into DataTable with Columns... "); }

            DataTable DtTable = new DataTable();
            Helpers Helpers = new Helpers();

            if (Delimiter == "FIXED WIDTH")
            {
                //get column names and widths from ColumnDefinitionFile
                var ColumnDefinitionFile = File.ReadLines(FixedWidthColumnFilePath, EncodingChoice);
                foreach (var line in ColumnDefinitionFile)
                {
                    Tuple<string, int> ColumnDefinition = Helpers.ParseColumnWidthLine(line);
                    string ColumnName = ColumnDefinition.Item1;
                    int ColumnLength = ColumnDefinition.Item2;

                    //clean ColumnName if needed
                    if (StripNonPrintableChars) Helpers.StripNonPrintableCharsFromValue(ref ColumnName);
                    Helpers.LimitValueLength(ref ColumnName, ref FieldLimit);

                    DataColumn DataColumn = new DataColumn(ColumnName);
                    DataColumn.MaxLength = ColumnLength;
                    DataColumn.AllowDBNull = true;
                    DtTable.Columns.Add(DataColumn);
                }
            }
            else
            {
                if (ColumnTypeFilePath != "")
                {
                    var ColumnDefinitionFile = File.ReadLines(ColumnTypeFilePath, EncodingChoice);
                    foreach (var line in ColumnDefinitionFile)
                    {
                        Tuple<string, string> ColumnDefinition = Helpers.ParseColumnTypeLine(line);
                        string ColumnName = ColumnDefinition.Item1;
                        string ColumnType = ColumnDefinition.Item2;
                        DtTable.Columns.Add(ColumnName);
                    }
                }
                else
                {
                    using (TextFieldParser FileReader = new TextFieldParser(FilePath, EncodingChoice))
                    {
                        FileReader.SetDelimiters(new string[] { Delimiter });
                        FileReader.HasFieldsEnclosedInQuotes = DoubleQuoted;
                        string[] ColFields = FileReader.ReadFields();

                        //get column names and widths from FilePath
                        foreach (string Column in ColFields)
                        {
                            string ColumnName = Column;

                            //clean ColumnName if needed
                            if (StripNonPrintableChars) Helpers.StripNonPrintableCharsFromValue(ref ColumnName);
                            Helpers.LimitValueLength(ref ColumnName, ref FieldLimit);

                            DataColumn DataColumn = new DataColumn(ColumnName);
                            DataColumn.AllowDBNull = true;
                            DtTable.Columns.Add(ColumnName);
                        }
                    }
                }
            }

            if (ConsoleOutput) { Console.WriteLine("DataTable with columns created"); }

            return DtTable;
        }
        public List<DataTable> ReadExcelFileIntoDataTablesWithColumns(string FilePath, ref List<string> TableNames, System.Text.Encoding EncodingChoice, bool StripNonPrintableChars, int FieldLimit, int HeaderRowIndex = 0, bool ConsoleOutput = true)
        {
            Helpers Helpers = new Helpers();
            List<DataTable> DataTables = new List<DataTable>();

            using (SpreadsheetDocument SpreadSheetDocument = SpreadsheetDocument.Open(@"" + FilePath, false))
            {
                IEnumerable<Sheet> Sheets = SpreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string FileName = System.IO.Path.GetFileNameWithoutExtension(FilePath);

                //for each Sheet in the Excel Spreadsheet
                int SheetIndex = 0;
                foreach (Sheet Sheet in Sheets)
                {
                    DataTable DataTable = new DataTable();
                    string SheetName = Sheet.Name;
                    string TableName = FileName + "_" + SheetName;
                    TableNames.Add(TableName);

                    string RelationshipId = Sheets.ElementAt(SheetIndex).Id.Value;
                    WorksheetPart WorksheetPart = (WorksheetPart)SpreadSheetDocument.WorkbookPart.GetPartById(RelationshipId);
                    Worksheet WorkSheet = WorksheetPart.Worksheet;
                    SheetData SheetData = WorkSheet.GetFirstChild<SheetData>();
                    SharedStringTablePart StringTablePart = SpreadSheetDocument.WorkbookPart.SharedStringTablePart;
                    IEnumerable<Row> Rows = SheetData.Descendants<Row>();
                    //IEnumerable<Column> Columns = SheetData.Descendants<Column>();
                    //int ColumnCount = Columns.ToArray().Length;

                    //for each cell of Row HeaderRowIndex
                    int CellIndex = 0;
                    int NormalIndex = 0;
                    List<string> ColumnHeaders = new List<string>();
                    foreach (Cell Cell in Rows.ElementAt(HeaderRowIndex))
                    {
                        //blank excel cells will be skipped using openxml
                        //so, need to check cell reference, add empty values, update index
                        string ColumnLetters = Helpers.GetColumnName(Cell.CellReference);
                        int ColumnLettersIndex = Helpers.GetColumnIndexFromName(ColumnLetters).Value;
                        int NullsEncountered = 0;
                        while ((ColumnLettersIndex - 1) != (CellIndex))
                        {
                            ColumnHeaders.Add("");
                            DataTable.Columns.Add("COLUMN" + (CellIndex + 1).ToString());
                            CellIndex++;
                            NullsEncountered++;
                        }

                        string FinalCellValue = "";

                        //https://stackoverflow.com/questions/36670768/openxml-cell-datetype-is-null
                        if (Cell != null)
                        {
                            if (Cell.DataType != null) //strings and booleans
                            {
                                if (Cell.DataType.Value == CellValues.SharedString)
                                {
                                    string value = Cell.CellValue.InnerXml;
                                    FinalCellValue = StringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                                }
                                else if (Cell.DataType.Value == CellValues.Boolean)
                                {
                                    FinalCellValue = Cell.CellValue.InnerText == "0" ? "FALSE" : "TRUE";
                                }
                                else
                                {
                                    FinalCellValue = Cell.CellValue.InnerText;
                                }
                            }
                            else //numbers and dates
                            {
                                /*
                                    General = 0,
                                    Number = 1,
                                    Decimal = 2,
                                    Currency = 164,
                                    Accounting = 44,
                                    DateShort = 14,
                                    DateLong = 165,
                                    Time = 166,
                                    Percentage = 10,
                                    Fraction = 12,
                                    Scientific = 11,
                                    Text = 49
                                */
                                if (Cell.StyleIndex != null)
                                {
                                    int StyleIndex = (int)Cell.StyleIndex.Value;
                                    CellFormat cellFormat = (CellFormat)SpreadSheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt(StyleIndex);
                                    uint formatId = cellFormat.NumberFormatId.Value;
                                    //Console.WriteLine(formatId);
                                    if ((formatId >= (uint)14 && formatId <= (uint)22) &&
                                        (formatId >= (uint)165 && formatId <= (uint)180))
                                    {
                                        double oaDate;
                                        if (double.TryParse(Cell.InnerText, out oaDate))
                                        {
                                            FinalCellValue = DateTime.FromOADate(oaDate).ToShortDateString();
                                        }
                                    }
                                    else
                                    {
                                        FinalCellValue = Cell.CellValue.InnerText;
                                    }
                                }
                                else //WHEN STYLE IS GENERAL AND ITS A NUMBER
                                {
                                    FinalCellValue = Cell.CellValue.InnerText;
                                }
                            }
                        }
                        else
                        {
                            FinalCellValue = null;
                        }
                        //Console.WriteLine(FinalCellValue);

                        if (StripNonPrintableChars) Helpers.StripNonPrintableCharsFromValue(ref FinalCellValue);
                        if (FieldLimit > 0) Helpers.LimitValueLength(ref FinalCellValue, ref FieldLimit);


                        //ColumnHeaders[CellIndex] = FinalCellValue;
                        CellIndex++;
                        NormalIndex++;

                        DataTable.Columns.Add(FinalCellValue);
                    }

                    //while(CellIndex < ColumnCount)
                    //{
                    //    DataTable.Columns.Add("COLUMN" + (CellIndex + 1).ToString());
                    //    CellIndex++;
                    //}

                    DataTables.Add(DataTable);

                    SheetIndex++;
                }
            }

            return DataTables;
        }
        public void CreateTablesInSqlServerVarchar(List<string> TableNames, ConnectionInfo ConnectionInfo, List<DataTable> DataTables, string Delimiter, System.Text.Encoding EncodingChoice, bool StripNonPrintableChars, int FieldLimit, string ColumnTypeMethod = "DEFAULT VARCHAR", string ColumnTypeFilePath = "", bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Creating " + DataTables.Count.ToString() + " tables in SQL Server... "); }

            List<string> ColumnTypes = new List<string>();
            Helpers Helpers = new Helpers();

            //if defining column types (not using varchar default)
            if (ColumnTypeFilePath != "" && ColumnTypeMethod == "FILE PATH")
            {
                var ColumnDefinitionFile = File.ReadLines(ColumnTypeFilePath, EncodingChoice);
                foreach (var line in ColumnDefinitionFile)
                {
                    Tuple<string, string> ColumnDefinition = Helpers.ParseColumnTypeLine(line);
                    string ColumnType = ColumnDefinition.Item2;
                    ColumnTypes.Add(ColumnType);
                }
            }

            //for each table, create table in SQL Server
            int index = 0;
            foreach (DataTable DataTable in DataTables)
            {
                string ColumnsForTableCreationQuery = "";
                int ColIndex = 0;
                string TableName = TableNames[index];

                //for each column, add to table creation query
                foreach (DataColumn Column in DataTable.Columns)
                {
                    //add [column name]
                    ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + "[" + Column.ColumnName + "] ";

                    //add column data type
                    if (ColumnTypeMethod == "FILE PATH")
                    {
                        ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + " " + ColumnTypes[ColIndex] + ",";
                    }
                    else if (Delimiter == "FIXED WIDTH" && ColumnTypeMethod == "DEFAULT VARCHAR")
                    {
                        ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + " VARCHAR(" + Column.MaxLength.ToString() + "),";
                    }
                    else if (Delimiter != "FIXED WIDTH" && ColumnTypeMethod == "DEFAULT VARCHAR")
                    {
                        ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + " VARCHAR(255),";
                    }
                }
                ColumnsForTableCreationQuery = ColumnsForTableCreationQuery.Substring(0, ColumnsForTableCreationQuery.Length - 1);

                //connect to SQL Server and run table creation query
                string ConnString = @"Server=" + ConnectionInfo.Server + ";Database=" + ConnectionInfo.Database + ";Trusted_Connection = True;";
                using (SqlConnection Conn = new SqlConnection(ConnString))
                {
                    Conn.Open();
                    string TableCreationQuery = "CREATE TABLE [" + TableName + "] (  " + ColumnsForTableCreationQuery + ")";
                    SqlCommand Cmd = new SqlCommand(TableCreationQuery, Conn);
                    Cmd.ExecuteNonQuery();
                }

                //output for user
                if (ConsoleOutput)
                {
                    if (Delimiter != "FIXED WIDTH" && ColumnTypeMethod == "DEFAULT VARCHAR")
                    {
                        Console.WriteLine("Created Table " + ConnectionInfo.Server + "." + ConnectionInfo.Database + "..[" + TableName + "] (all columns VARCHAR(255))");
                    }
                    else if (Delimiter == "FIXED WIDTH" && ColumnTypeMethod == "DEFAULT VARCHAR")
                    {
                        Console.WriteLine("Created Table " + ConnectionInfo.Server + "." + ConnectionInfo.Database + "..[" + TableName + "] (all columns VARCHAR(N))");
                    }
                    else
                    {
                        Console.WriteLine("Created Table " + ConnectionInfo.Server + "." + ConnectionInfo.Database + "..[" + TableName + "] ");
                    }
                }

                index++;
            }
        }
        public void ReadFileIntoDataTableWithRowsAndInsertIntoSqlServerTable(string FilePath, string TableName, ConnectionInfo ConnectionInfo, DataTable BaseDtTable, int BatchLimit, string Delimiter, bool DoubleQuoted, System.Text.Encoding EncodingChoice, bool StripNonPrintableChars, int FieldLimit, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Reading file rows... "); }

            Helpers Helpers = new Helpers();
            string ConnString = @"Server=" + ConnectionInfo.Server + ";Database=" + ConnectionInfo.Database + ";Trusted_Connection = True;";

            //open file
            //just for the header
            string[] FileFields = null;
            using (TextFieldParser FileReader = new TextFieldParser(FilePath, EncodingChoice))
            {
                FileReader.SetDelimiters(new string[] { Delimiter });
                FileReader.HasFieldsEnclosedInQuotes = DoubleQuoted;
                FileFields = FileReader.ReadFields();
            }
            //for the rest of the rows
            using (TextFieldParser FileReader = new TextFieldParser(FilePath, EncodingChoice))
            {
                //set read settings
                if (Delimiter == "FIXED WIDTH")
                {
                    int[] FieldWidths = new int[BaseDtTable.Columns.Count];

                    int c = 0;
                    foreach (DataColumn Column in BaseDtTable.Columns)
                    {
                        FieldWidths[c] = Column.MaxLength;
                        c++;
                    }
                    FileReader.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.FixedWidth;
                    FileReader.SetFieldWidths(FieldWidths);
                    FileReader.HasFieldsEnclosedInQuotes = false;
                }
                else
                {
                    FileReader.SetDelimiters(new string[] { Delimiter });
                    FileReader.HasFieldsEnclosedInQuotes = DoubleQuoted;
                }

                //for each row in the file...
                int Row = 0;
                bool LeftoverData = false;
                DataTable TempDtTable = BaseDtTable;
                

                while (!FileReader.EndOfData)
                {
                    LeftoverData = true;
                    string[] FieldData = null;

                    try
                    {
                        //read the fields
                        FieldData = FileReader.ReadFields();
                        //Console.WriteLine(FieldData.Length);
                        string[] FinalFieldData = new string[BaseDtTable.Columns.Count];
                        int FieldsAdded = 0;
                        for (int cf = 0; cf < FieldData.Length; cf++)
                        {
                            if (BaseDtTable.Columns.Contains(FileFields[cf]))
                            {
                                FinalFieldData[FieldsAdded] = FieldData[cf];
                                //Console.WriteLine(FieldData[cf]);
                                FieldsAdded++;
                            }
                        }

                        //for each field in the row, clean if necessary
                        if (StripNonPrintableChars)
                        {
                            for (int cf = 0; cf < FinalFieldData.Length; cf++)
                            {
                                Helpers.StripNonPrintableCharsFromValue(ref FinalFieldData[cf]);
                            }
                        }
                        if (FieldLimit > 0 && Delimiter != "FIXED WIDTH")
                        {
                            for (int cf = 0; cf < FinalFieldData.Length; cf++)
                            {
                                Helpers.LimitValueLength(ref FinalFieldData[cf], ref FieldLimit);
                            }
                        }

                        //add fields to row in datatable
                        if (Row != 0 || Delimiter == "FIXED WIDTH") //skip header (FIXED WIDTH files dont have headers)
                        {
                            TempDtTable.Rows.Add(FinalFieldData);
                        }
                    }
                    catch
                    {
                        Console.WriteLine("ROW " + Row.ToString() + " SKIPPED.");
                    }

                    /***************************************
                    * INSERT ROWS TO SQL TABLE
                    ***************************************/
                    //when we get to Row BatchLimit, import that chunk into SQL Server
                    if (Row != 0 && Row % BatchLimit == 0)
                    {
                        LeftoverData = false;
                        InsertDataTableUsingSqlBulkCopy(ref ConnString, ref TableName, ref TempDtTable, ref Row);

                        //reset TempDtTable, if we don't do this, then large files (example: 18 columns/4 million rows) will cause the script to run out of memory
                        TempDtTable = BaseDtTable; //not sure if this necessary
                        TempDtTable.Rows.Clear(); //definitely necessary
                    }

                    Row++;
                }

                //Importing the remaining data (necessary because of the batching)
                //the script will only end up coming here to insert data if the current DataTable is not equal to the BatchLimit of rows
                if (LeftoverData == true)
                {
                    InsertDataTableUsingSqlBulkCopy(ref ConnString, ref TableName, ref TempDtTable, ref Row);

                    //reset TempDtTable, if we don't do this, then large files (example: 18 columns/4 million rows) will cause the script to run out of memory
                    TempDtTable = BaseDtTable; //not sure if this necessary
                    TempDtTable.Rows.Clear(); //definitely necessary
                }
            }
            if (ConsoleOutput) { Console.Write("."); }
            if (ConsoleOutput) { Console.WriteLine(""); }
        }
        public void ReadFileIntoDataTableWithRowsAndInsertIntoSnowflakeTable(string FilePath, string TableName, ConnectionInfo ConnectionInfo, Snowflake Snowflake, DataTable BaseDtTable, int BatchLimit, string Delimiter, bool DoubleQuoted, System.Text.Encoding EncodingChoice, string EncodingChoiceSnowflake, bool StripNonPrintableChars, int FieldLimit, string ColumnTypeMethod = "DEFAULT VARCHAR", string ColumnTypeFilePath = "", bool ImportToExistingTable = false, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Staging file... "); }
            string StageName = "WINFORM_TEMP_STAGE";
            Snowflake.StageFile(FilePath, StageName);
            if (ConsoleOutput) { Console.WriteLine("File in staging area. "); }

            if (ConsoleOutput) { Console.WriteLine("Importing file... "); }
            Snowflake.ImportFile(FilePath, StageName, BaseDtTable, TableName, Delimiter, DoubleQuoted, EncodingChoice, EncodingChoiceSnowflake, StripNonPrintableChars, FieldLimit, ColumnTypeMethod, ColumnTypeFilePath, ImportToExistingTable);
            if (ConsoleOutput) { Console.WriteLine("File imported. "); }

            if (ConsoleOutput) { Console.WriteLine(""); }
        }
        public void ReadExcelFilePerSheetIntoDataTablesWithRowsAndInsertIntoSqlServerTables(string FilePath, List<string> TableNames, ConnectionInfo ConnectionInfo, List<DataTable> DataTables, int BatchLimit, string Delimeter, System.Text.Encoding EncodingChoice, bool StripNonPrintableChars, int FieldLimit, int HeaderRowIndex = 0, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Reading Excel file rows... "); }
            Helpers Helpers = new Helpers();

            string ConnString = @"Server=" + ConnectionInfo.Server + ";Database=" + ConnectionInfo.Database + ";Trusted_Connection = True;";

            //https://stackoverflow.com/questions/3321082/from-excel-to-datatable-in-c-sharp-with-open-xml
            using (SpreadsheetDocument SpreadSheetDocument = SpreadsheetDocument.Open(@"" + FilePath, false))
            {

                IEnumerable<Sheet> Sheets = SpreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

                int SheetIndex = 0;
                foreach (Sheet Sheet in Sheets)
                {
                    string SheetName = Sheet.Name;
                    string TableName = TableNames[SheetIndex];
                    if (ConsoleOutput) { Console.WriteLine("Sheet " + SheetIndex.ToString() + 1 + ": " + FilePath + "_" + SheetName); }

                    string RelationshipId = Sheets.ElementAt(SheetIndex).Id.Value;//.First().Id.Value;
                    WorksheetPart WorksheetPart = (WorksheetPart)SpreadSheetDocument.WorkbookPart.GetPartById(RelationshipId);
                    Worksheet WorkSheet = WorksheetPart.Worksheet;
                    SheetData SheetData = WorkSheet.GetFirstChild<SheetData>();
                    IEnumerable<Row> Rows = SheetData.Descendants<Row>();
                    SharedStringTablePart StringTablePart = SpreadSheetDocument.WorkbookPart.SharedStringTablePart;

                    int RowIndex = 0;
                    bool LeftoverData = false;
                    DataTable TempDataTable = DataTables[SheetIndex];


                    //RBT's answer from: https://stackoverflow.com/questions/5115257/openxml-sdk-returning-a-number-for-cellvalue-instead-of-cells-text 
                    foreach (Row Row in Rows)
                    {
                        LeftoverData = true;
                        if (RowIndex > HeaderRowIndex) //skip header row...
                        {
                            DataRow TempRow = TempDataTable.NewRow();
                            int CellIndex = 0;
                            foreach (Cell Cell in Row)
                            {
                                if (CellIndex < TempRow.ItemArray.Length)
                                {
                                    //blank excel cells will be skipped using openxml
                                    //so, need to check cell reference, add empty values, update index
                                    string ColumnLetters = Helpers.GetColumnName(Cell.CellReference);
                                    int ColumnLettersIndex = Helpers.GetColumnIndexFromName(ColumnLetters).Value;

                                    while ((ColumnLettersIndex - 1) != (CellIndex))
                                    {
                                        TempRow.ItemArray.Append("");
                                        TempRow[CellIndex] = "";
                                        CellIndex++;
                                    }

                                    //Console.WriteLine(RowIndex.ToString() + " " + CellIndex.ToString());

                                    string FinalCellValue = "";
                                    //https://stackoverflow.com/questions/36670768/openxml-cell-datetype-is-null
                                    if (Cell != null && Cell.CellValue != null)
                                    {
                                        if (Cell.DataType != null) //strings and booleans
                                        {
                                            if (Cell.DataType.Value == CellValues.SharedString)
                                            {
                                                string value = Cell.CellValue.InnerXml;
                                                FinalCellValue = StringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                                            }
                                            else if (Cell.DataType.Value == CellValues.Boolean)
                                            {
                                                FinalCellValue = Cell.CellValue.InnerText == "0" ? "FALSE" : "TRUE";
                                            }
                                            else
                                            {
                                                FinalCellValue = Cell.CellValue.InnerText;
                                            }
                                        }
                                        else //numbers and dates
                                        {
                                            /*
                                                General = 0,
                                                Number = 1,
                                                Decimal = 2,
                                                Currency = 164,
                                                Accounting = 44,
                                                DateShort = 14,
                                                DateLong = 165,
                                                Time = 166,
                                                Percentage = 10,
                                                Fraction = 12,
                                                Scientific = 11,
                                                Text = 49
                                            */

                                            if (Cell.StyleIndex != null)
                                            {
                                                int StyleIndex = (int)Cell.StyleIndex.Value;
                                                CellFormat cellFormat = (CellFormat)SpreadSheetDocument.WorkbookPart.WorkbookStylesPart.Stylesheet.CellFormats.ElementAt(StyleIndex);
                                                uint formatId = cellFormat.NumberFormatId.Value;
                                                //Console.WriteLine(formatId);
                                                if ((formatId >= (uint)14 && formatId <= (uint)22) ||
                                                    (formatId >= (uint)165 && formatId <= (uint)180))
                                                    //(formatId >= (uint)164 && formatId <= (uint)180))
                                                {
                                                    double oaDate;
                                                    if (double.TryParse(Cell.InnerText, out oaDate))
                                                    {
                                                        //Console.WriteLine(oaDate);
                                                        FinalCellValue = DateTime.FromOADate(oaDate).ToShortDateString();
                                                        //Console.WriteLine(FinalCellValue);
                                                    }
                                                }
                                                else
                                                {
                                                    FinalCellValue = Cell.CellValue.InnerText;
                                                }
                                            }
                                            else //WHEN STYLE IS GENERAL AND ITS A NUMBER
                                            {
                                                FinalCellValue = Cell.CellValue.InnerText;
                                            }
                                        }
                                    }
                                    else
                                    {
                                        FinalCellValue = "";
                                    }

                                    if(FinalCellValue == null)
                                    {
                                        FinalCellValue = "";
                                    }
                                    //Console.WriteLine(FinalCellValue);

                                    //for each field in the row
                                    //if (FasterImport)
                                    //{
                                    //PrepareValueForImport(ref FinalCellValue);
                                    //if(EncodingChoice )
                                    if (StripNonPrintableChars) Helpers.StripNonPrintableCharsFromValue(ref FinalCellValue);
                                    if (FieldLimit > 0) Helpers.LimitValueLength(ref FinalCellValue, ref FieldLimit);
                                    //}


                                    TempRow[CellIndex] = FinalCellValue;
                                    CellIndex++;
                                }

                            }
                            TempDataTable.Rows.Add(TempRow);

                            //when we get to Row BatchLimit, import that chunk into SQL Server
                            //also print to console to help with tracking
                            if (RowIndex != 0 && RowIndex % BatchLimit == 0)
                            {
                                LeftoverData = false;
                                InsertDataTableUsingSqlBulkCopy(ref ConnString, ref TableName, ref TempDataTable, ref RowIndex);

                                //reset TempDataTable
                                //if we don't do this, then large files will cause the program to run out of memory
                                TempDataTable = DataTables[SheetIndex]; //not sure if this necessary
                                TempDataTable.Rows.Clear(); //definitely necessary
                            }
                        }

                        //if (ConsoleOutput) { if (RowIndex > 0 && RowIndex % 1000000 == 0) { Console.WriteLine("ROW " + Row); } }
                        RowIndex++;
                    }

                    //Importing the remaining data (necessary because of the batching)
                    //the script will only end up coming here to insert data if the file is under the BatchLimit of rows
                    if (LeftoverData == true)
                    {
                        InsertDataTableUsingSqlBulkCopy(ref ConnString, ref TableName, ref TempDataTable, ref RowIndex);

                        TempDataTable = DataTables[SheetIndex];
                        TempDataTable.Rows.Clear();
                    }

                    SheetIndex++;
                }
            }
            if (ConsoleOutput) { Console.Write("."); }
            if (ConsoleOutput) { Console.WriteLine(""); }
        }
        public void InsertDataTableUsingSqlBulkCopy(ref string ConnString, ref string TableName, ref DataTable TempDataTable, ref int RowIndex, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.Write("\r" + $"{RowIndex:n0}" + " rows read"); }
            using (SqlConnection Conn = new SqlConnection(ConnString))
            {
                Conn.Open();
                using (SqlBulkCopy SqlBulk = new SqlBulkCopy(Conn))
                {
                    SqlBulk.DestinationTableName = "[dbo].[" + TableName + "]";
                    foreach (var Column in TempDataTable.Columns)
                    {
                        SqlBulk.ColumnMappings.Add(Column.ToString(), Column.ToString());
                    }


                    //if (ConsoleOutput) { Console.Write("Inserting current batch of rows to table... "); }

                    SqlBulk.WriteToServer(TempDataTable);

                    //if (ConsoleOutput) { Console.Write("Inserted"); }
                }
            }
        }
    }
}

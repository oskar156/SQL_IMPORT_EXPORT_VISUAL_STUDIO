/* TABLE OF CONTENTS
 * 
 * IMPORT FUNCTIONS
 * public List<string> GetFilesToImport(string ImportPath, string Extension, bool ConsoleOutput = true)
 * public DataTable ReadFileIntoDataTableWithColumns(string FilePath, string Delimeter, string FixedWidthColumnFilePath, bool ConsoleOutput = true)
 * public List<DataTable> ReadExcelFileIntoDataTablesWithColumns(string FilePath, ref List<string> TableNames, bool ConsoleOutput = true)
 * public void CreateTableInSqlVarchar(string TableName, string Server, string Database, DataTable DtTable, string Delimeter, bool ConsoleOutput = true)
 * public void CreateTablesInSqlVarchar(List<string> TableNames, string Server, string Database, List<DataTable> DataTables, string Delimeter, bool ConsoleOutput = true)
 * public void ReadFileIntoDataTableWithRowsAndInsertIntoSqlTable(string FilePath, string TableName, string Server, string Database, DataTable BaseDtTable, int BatchLimit, string Delimeter, bool ConsoleOutput = true)
 * public void ReadExcelFilePerSheetIntoDataTablesWithRowsAndInsertIntoSqlTables(string FilePath, List<string> TableNames, string Server, string Database, List<DataTable> DataTables, int BatchLimit, string Delimeter, bool ConsoleOutput = true)
 * public void InsertDataTableUsingSqlBulkCopy(ref string ConnString, ref string TableName, ref DataTable TempDataTable, ref int RowIndex, bool ConsoleOutput = true)
 * 
 * EXPORT FUNCTIONS
 * public List<string> GetListofTablesFromSqlDb(string Server, string Database, List<string> ListOfTablesToSearchFor, bool ConsoleOutput = true)
 * public List<string> GetListofTablesFromSqlDb(string Server, string Database, string RegexSearchPattern = "", bool ConsoleOutput = true)
 * public int ExportTableFromSqlToFile(string Server, string Database, string TableToExport, string ExportPath, string Extension, string Delimeter, string Qualifier, bool IncludeHeaders, string FixedWidthColumnLengthMethod, int SizeLimit, string SizeLimitType, bool IncludeHeaderInSplitFiles, string OrderBy = "", bool ConsoleOutput = true)
 * 
 * OTHER FUNCTION
 * public void PrepareValueForImport(ref string Value)
 * 
 * 
*/

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO; //for TextFieldParser
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;

namespace SQL_SERVER_IMPORT_EXPORT
{
    public class Helpers
    {

        //------------------------------------------------------------------------------------
        // IMPORT
        //------------------------------------------------------------------------------------
        public List<string> GetFilesToImport(string ImportPath, string Extension, bool ConsoleOutput = true)
        {
            //https://stackoverflow.com/questions/20759302/upload-csv-file-to-sql-server
            if (ConsoleOutput) { Console.WriteLine("Getting file(s) to import from " + ImportPath + "*." + Extension + "..."); }

            List<string> FilesToImport = new List<string>();

            //detect whether its a directory or file
            //https://stackoverflow.com/questions/1395205/better-way-to-check-if-a-path-is-a-file-or-a-directory
            bool IsImportPathAFile = true;
            FileAttributes Attr = File.GetAttributes(ImportPath);
            if ((Attr & FileAttributes.Directory) == FileAttributes.Directory)
            {
                IsImportPathAFile = false;
            }

            if (IsImportPathAFile) //if ImportPath is a file, import only that file
            {
                FilesToImport.Add(ImportPath);
            }
            else //if ImportPath is a direcotry, import only every file in that directory that matches Extension
            {
                string[] FilesInPath = Directory.GetFiles(ImportPath, "*." + Extension);
                foreach (string File in FilesInPath)
                {

                    FilesToImport.Add(File);
                }
            }

            if (ConsoleOutput)
            {
                Console.WriteLine(FilesToImport.Count.ToString() + " files found");
                Console.WriteLine("");
            }

            return FilesToImport;
        }
        public DataTable ReadFileIntoDataTableWithColumns(string FilePath, string Delimeter, string FixedWidthColumnFilePath, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Reading file into DataTable with Columns... "); }

            DataTable DtTable = new DataTable();

            if (Delimeter == "FIXED WIDTH")
            {
                var ColumnDefinitionFile = File.ReadLines(FixedWidthColumnFilePath);
                foreach (var line in ColumnDefinitionFile)
                {
                    string lineTrimmed = line.Trim();
                    int LastSpaceIndex = lineTrimmed.LastIndexOf(" ");
                    string ColumnName = lineTrimmed.Substring(0, LastSpaceIndex);
                    int ColumnLength = Int32.Parse(lineTrimmed.Substring(LastSpaceIndex).Trim());

                    DataColumn DataColumn = new DataColumn(ColumnName, typeof(string));
                    DataColumn.MaxLength = ColumnLength;
                    DtTable.Columns.Add(DataColumn);
                }
            }
            else
            {
                using (TextFieldParser FileReader = new TextFieldParser(FilePath))
                {
                    FileReader.SetDelimiters(new string[] { Delimeter });
                    FileReader.HasFieldsEnclosedInQuotes = true;
                    string[] ColFields = FileReader.ReadFields();
                    foreach (string Column in ColFields)
                    {
                        DataColumn DataColumn = new DataColumn(Column);
                        DataColumn.AllowDBNull = true;
                        DtTable.Columns.Add(Column);
                    }
                }
            }

            if (ConsoleOutput) { Console.Write("DataTable with columns created"); }

            return DtTable;
        }
        public List<DataTable> ReadExcelFileIntoDataTablesWithColumns(string FilePath, ref List<string> TableNames, bool ConsoleOutput = true)
        {
            List<DataTable> DataTables = new List<DataTable>();

            //https://stackoverflow.com/questions/3321082/from-excel-to-datatable-in-c-sharp-with-open-xml
            using (SpreadsheetDocument SpreadSheetDocument = SpreadsheetDocument.Open(@"" + FilePath, false))
            {

                IEnumerable<Sheet> Sheets = SpreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string FileName = System.IO.Path.GetFileNameWithoutExtension(FilePath);

                int SheetIndex = 0;
                foreach (Sheet Sheet in Sheets)
                {
                    DataTable DataTable = new DataTable();
                    string TableName = FileName + " - " + Sheet.Name;
                    if (TableName.Length > 32)
                    {
                        TableName = TableName.Substring(0, 31);
                    }
                    TableNames.Add(TableName);

                    string RelationshipId = Sheets.ElementAt(SheetIndex).Id.Value;//.First().Id.Value;
                    WorksheetPart WorksheetPart = (WorksheetPart)SpreadSheetDocument.WorkbookPart.GetPartById(RelationshipId);
                    Worksheet WorkSheet = WorksheetPart.Worksheet;
                    string SheetName = Sheet.Name;
                    SheetData SheetData = WorkSheet.GetFirstChild<SheetData>();
                    IEnumerable<Row> Rows = SheetData.Descendants<Row>();

                    //RBT's answer from: https://stackoverflow.com/questions/5115257/openxml-sdk-returning-a-number-for-cellvalue-instead-of-cells-text 
                    foreach (Cell Cell in Rows.ElementAt(0))
                    {
                        SharedStringTablePart StringTablePart = SpreadSheetDocument.WorkbookPart.SharedStringTablePart;
                        string value = Cell.CellValue.InnerXml;
                        string FinalCellValue = "";

                        if (Cell.DataType != null && Cell.DataType.Value == CellValues.SharedString)
                        {
                            FinalCellValue = StringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                        }
                        else
                        {
                            FinalCellValue = value;
                        }

                        DataTable.Columns.Add(FinalCellValue);
                    }

                    DataTables.Add(DataTable);

                    SheetIndex++;
                }
            }

            return DataTables;
        }
        public void CreateTableInSqlVarchar(string TableName, string Server, string Database, DataTable DtTable, string Delimeter, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Creating table in sql... "); }

            string ColumnsForTableCreationQuery = "";

            foreach (DataColumn Column in DtTable.Columns)
            {
                if (Delimeter == "FIXED WIDTH")
                {
                    ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + "[" + Column.ColumnName + "] VARCHAR(" + Column.MaxLength.ToString() + "),";
                }
                else
                {
                    ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + "[" + Column.ColumnName + "] VARCHAR(255),";
                }
            }

            string ConnString = @"Server=" + Server + ";Database=" + Database + ";Trusted_Connection = True;";
            using (SqlConnection Conn = new SqlConnection(ConnString))
            {
                Conn.Open();
                string TableCreationQuery = "CREATE TABLE [" + TableName + "] (  " + ColumnsForTableCreationQuery + ")";
                SqlCommand Cmd = new SqlCommand(TableCreationQuery, Conn);
                Cmd.ExecuteNonQuery();
            }
            if (ConsoleOutput)
            {
                if (Delimeter == "FIXED WIDTH")
                {
                    Console.Write("Created Table " + Server + "." + Database + "..[" + TableName + "] ");
                }
                else
                {
                    Console.Write("Created Table " + Server + "." + Database + "..[" + TableName + "] (all columns VARCHAR(255))");
                }
            }
        }
        public void CreateTablesInSqlVarchar(List<string> TableNames, string Server, string Database, List<DataTable> DataTables, string Delimeter, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Creating " + DataTables.Count.ToString() + "tables in sql... "); }

            int index = 0;
            foreach (DataTable DataTable in DataTables)
            {
                string ColumnsForTableCreationQuery = "";
                string TableName = TableNames[index];

                foreach (DataColumn Column in DataTable.Columns)
                {
                    if (Delimeter == "FIXED WIDTH")
                    {
                        ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + "[" + Column.ColumnName + "] VARCHAR(" + Column.MaxLength.ToString() + "),";
                    }
                    else
                    {
                        ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + "[" + Column.ColumnName + "] VARCHAR(255),";
                    }
                }

                string ConnString = @"Server=" + Server + ";Database=" + Database + ";Trusted_Connection = True;";
                using (SqlConnection Conn = new SqlConnection(ConnString))
                {
                    Conn.Open();
                    string TableCreationQuery = "CREATE TABLE [" + TableName + "] (  " + ColumnsForTableCreationQuery + ")";
                    SqlCommand Cmd = new SqlCommand(TableCreationQuery, Conn);
                    Cmd.ExecuteNonQuery();
                }
                if (ConsoleOutput)
                {
                    if (Delimeter == "FIXED WIDTH")
                    {
                        Console.WriteLine("Created Table " + Server + "." + Database + "..[" + TableName + "] ");
                    }
                    else
                    {
                        Console.WriteLine("Created Table " + Server + "." + Database + "..[" + TableName + "] (all columns VARCHAR(255))");
                    }
                }

                index++;
            }
        }
        public void ReadFileIntoDataTableWithRowsAndInsertIntoSqlTable(string FilePath, string TableName, string Server, string Database, DataTable BaseDtTable, int BatchLimit, string Delimeter, bool DoubleQuoted, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Reading file into DataTable with Rows... "); }

            string ConnString = @"Server=" + Server + ";Database=" + Database + ";Trusted_Connection = True;";

            using (TextFieldParser FileReader = new TextFieldParser(FilePath))
            {
                if (Delimeter == "FIXED WIDTH")
                {
                    int[] FieldWidths = new int[BaseDtTable.Columns.Count];

                    int c = 0;
                    foreach (DataColumn Column in BaseDtTable.Columns)
                    {
                        FieldWidths[c] = Column.MaxLength;
                        c++;
                    }
                    //https://learn.microsoft.com/en-us/dotnet/api/microsoft.visualbasic.fileio.textfieldparser.textfieldtype?view=net-8.0
                    FileReader.TextFieldType = Microsoft.VisualBasic.FileIO.FieldType.FixedWidth;
                    FileReader.SetFieldWidths(FieldWidths);
                    FileReader.HasFieldsEnclosedInQuotes = false;
                }
                else
                {
                    FileReader.SetDelimiters(new string[] { Delimeter });
                    FileReader.HasFieldsEnclosedInQuotes = DoubleQuoted;
                }

                int Row = 0;
                bool LeftoverData = false;
                DataTable TempDtTable = BaseDtTable;

                while (!FileReader.EndOfData)
                {
                    LeftoverData = true;

                    //https://stackoverflow.com/questions/16225909/dealing-with-fields-containing-unescaped-double-quotes-with-textfieldparser
                    string[] FieldData = null;// 

                    try
                    {
                        FieldData = FileReader.ReadFields();

                        //for each field in the row
                        for (int cf = 0; cf < FieldData.Length; cf++)
                        {
                            PrepareValueForImport(ref FieldData[cf]);
                        }
                        if (Row != 0) //skip header
                        {
                            TempDtTable.Rows.Add(FieldData);
                        }
                    }
                    catch
                    {
                        Console.WriteLine("ROW " + Row.ToString() + " SKIPPED.");
                    }

                    /***************************************
                    * INSERT ROWS TO TABLE
                    ***************************************/
                    //when we get to Row BatchLimit, import that chunk into SQL Server
                    //also print to console to help with tracking
                    if (Row != 0 && Row % BatchLimit == 0)
                    {
                        LeftoverData = false;
                        InsertDataTableUsingSqlBulkCopy(ref ConnString, ref TableName, ref TempDtTable, ref Row);

                        //reset TempDtTable
                        //if we don't do this, then large files (example: 18 columns/4 million rows) will cause the script to run out of memory
                        TempDtTable = BaseDtTable; //not sure if this necessary
                        TempDtTable.Rows.Clear(); //definitely necessary
                    }
                    Row++;
                }

                //Importing the remaining data (necessary because of the batching)
                //the script will only end up coming here to insert data if the current DataTable is under the BatchLimit of rows
                if (LeftoverData == true)
                {
                    InsertDataTableUsingSqlBulkCopy(ref ConnString, ref TableName, ref TempDtTable, ref Row);

                    TempDtTable = BaseDtTable;//not sure if either are necessary at this point, because it's after the loop
                    TempDtTable.Rows.Clear();
                }
            }
        }
        public void ReadExcelFilePerSheetIntoDataTablesWithRowsAndInsertIntoSqlTables(string FilePath, List<string> TableNames, string Server, string Database, List<DataTable> DataTables, int BatchLimit, string Delimeter, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Reading Excel file into DataTable with Rows... "); }

            string ConnString = @"Server=" + Server + ";Database=" + Database + ";Trusted_Connection = True;";

            //https://stackoverflow.com/questions/3321082/from-excel-to-datatable-in-c-sharp-with-open-xml
            using (SpreadsheetDocument SpreadSheetDocument = SpreadsheetDocument.Open(@"" + FilePath, false))
            {

                IEnumerable<Sheet> Sheets = SpreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();

                int SheetIndex = 0;
                foreach (Sheet Sheet in Sheets)
                {
                    string SheetName = Sheet.Name;
                    string TableName = TableNames[SheetIndex];
                    if (ConsoleOutput) { Console.WriteLine("Sheet " + SheetIndex.ToString() + 1 + ": " + FilePath + " - " + SheetName); }

                    string RelationshipId = Sheets.ElementAt(SheetIndex).Id.Value;//.First().Id.Value;
                    WorksheetPart WorksheetPart = (WorksheetPart)SpreadSheetDocument.WorkbookPart.GetPartById(RelationshipId);
                    Worksheet WorkSheet = WorksheetPart.Worksheet;
                    SheetData SheetData = WorkSheet.GetFirstChild<SheetData>();
                    IEnumerable<Row> Rows = SheetData.Descendants<Row>();

                    int RowIndex = 0;
                    bool LeftoverData = false;
                    DataTable TempDataTable = DataTables[SheetIndex];


                    //RBT's answer from: https://stackoverflow.com/questions/5115257/openxml-sdk-returning-a-number-for-cellvalue-instead-of-cells-text 
                    foreach (Row Row in Rows)
                    {
                        LeftoverData = true;
                        if (RowIndex > 0) //skip header row...
                        {
                            DataRow TempRow = TempDataTable.NewRow();


                            int CellIndex = 0;
                            foreach (Cell Cell in Row)
                            {
                                SharedStringTablePart StringTablePart = SpreadSheetDocument.WorkbookPart.SharedStringTablePart;
                                string value = Cell.CellValue.InnerXml;
                                string FinalCellValue = "";

                                if (Cell.DataType != null && Cell.DataType.Value == CellValues.SharedString)
                                {
                                    FinalCellValue = StringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                                }
                                else
                                {
                                    FinalCellValue = value;
                                }

                                PrepareValueForImport(ref FinalCellValue);

                                TempRow[CellIndex] = FinalCellValue;
                                CellIndex++;
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
        }
        public void InsertDataTableUsingSqlBulkCopy(ref string ConnString, ref string TableName, ref DataTable TempDataTable, ref int RowIndex, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.Write(RowIndex.ToString() + " rows read"); }
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


                    if (ConsoleOutput) { Console.WriteLine("Inserting rows to table... "); }

                    SqlBulk.WriteToServer(TempDataTable);

                    if (ConsoleOutput) { Console.WriteLine("Inserted"); }
                }
            }
        }

        //------------------------------------------------------------------------------------
        // EXPORT
        //------------------------------------------------------------------------------------
        public List<string> GetListofTablesFromSqlDb(string Server, string Database, List<string> ListOfTablesToSearchFor, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Getting List of Tables from Sql"); }

            List<string> Tables = new List<string>();

            string ConnString = @"Server=" + Server + ";Database=" + Database + ";Trusted_Connection = True;";
            using (SqlConnection Conn = new SqlConnection(ConnString))
            {
                Conn.Open();
                DataTable TablesInSqlDb = Conn.GetSchema("Tables");
                //int TableIndex = 0;
                foreach (DataRow Row in TablesInSqlDb.Rows)
                {
                    string TableName = Row[2].ToString();

                    if (ListOfTablesToSearchFor.Contains(TableName))
                    {
                        Tables.Add(TableName);
                    }
                    //TableIndex++;
                }
            }

            Tables.Sort();

            if (ConsoleOutput) { Console.WriteLine(Tables.Count.ToString() + " tables found"); }
            return Tables;
        }
        public List<string> GetListofTablesFromSqlDb(string Server, string Database, string RegexSearchPattern = "", bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Getting Tables from Sql"); }

            List<string> Tables = new List<string>();

            Regex re = new Regex("");
            if (RegexSearchPattern != "")
            {
                re = new Regex(RegexSearchPattern);
            }

            string ConnString = @"Server=" + Server + ";Database=" + Database + ";Trusted_Connection = True;";
            using (SqlConnection Conn = new SqlConnection(ConnString))
            {
                Conn.Open();
                DataTable TablesInSqlDb = Conn.GetSchema("Tables");
                //int TableIndex = 0;
                foreach (DataRow Row in TablesInSqlDb.Rows)
                {
                    string TableName = Row[2].ToString();

                    if (RegexSearchPattern == "")
                    {
                        Tables.Add(TableName);
                    }
                    else
                    {
                        if (re.IsMatch(TableName))
                        {
                            Tables.Add(TableName);
                        }
                    }
                    //TableIndex++;
                }
            }

            Tables.Sort();

            if (ConsoleOutput)
            {
                string OutputMessage = "";
                if (RegexSearchPattern != "")
                {
                    OutputMessage = " (tables that matched regex: " + RegexSearchPattern + ")";
                }
                else
                {
                    OutputMessage = " (all tables in the databse)";
                }
                Console.WriteLine(Tables.Count.ToString() + " tables found" + OutputMessage);
            }
            return Tables;
        }
        public int ExportTableFromSqlToFile(string Server, string Database, string TableToExport, string ExportPath, string Extension, string Delimeter, string Qualifier, bool IncludeHeaders, string FixedWidthColumnLengthMethod, int SizeLimit, string SizeLimitType, bool IncludeHeaderInSplitFiles, string OrderBy = "", bool ConsoleOutput = true)
        {
            int FilesCreated = 0;
            if (ConsoleOutput) { Console.WriteLine("Reading table from SQL Server"); }

            string ConnString = @"Server=" + Server + ";Database=" + Database + ";Trusted_Connection = True;";

            SqlDataReader DataReader = null;
            using (SqlConnection Conn = new SqlConnection(ConnString))
            {
                Conn.Open();
                string TableExportQuery = "SELECT * FROM [" + TableToExport + "]";
                if(OrderBy != "")
                {
                    TableExportQuery += " ORDER BY " + OrderBy;
                }

                SqlCommand Cmd = new SqlCommand(TableExportQuery, Conn);
                DataReader = Cmd.ExecuteReader();

                //https://learn.microsoft.com/en-us/dotnet/api/system.data.sqlclient.sqldatareader?view=netframework-4.8.1#properties
                int RowCount = 0;// DataReader.RecordsAffected; RETURNS -1 WHEN I TRIED IT
                int ColCount = DataReader.FieldCount;

                if (ConsoleOutput) { Console.WriteLine("Exporting table to file"); }

                //ExportPath
                string FileExportPathBase = ExportPath + "\\" + TableToExport;
                string FileExportPath = FileExportPathBase;
                if (SizeLimit > 0)
                {
                    FileExportPath += "-0";
                }
                FileExportPath += "." + Extension;


                if (Extension == "xlsx")
                {
                    //https://stackoverflow.com/questions/41605649/i-want-to-create-xlsx-excel-file-from-c-sharp
                    //need to get count
                    using (SqlConnection Conn2 = new SqlConnection(ConnString))
                    {
                        Conn2.Open();
                        string CountQuery = "SELECT COUNT(*) as ROW_COUNT FROM [" + TableToExport + "]";
                        SqlCommand CountCmd = new SqlCommand(CountQuery, Conn2);
                        SqlDataReader CountDataReader = CountCmd.ExecuteReader();

                        while (CountDataReader.Read())
                        {
                            RowCount = Int32.Parse(CountDataReader.GetValue(0).ToString());
                        }
                    }
                    
                    Microsoft.Office.Interop.Excel.Application Xlsx = new Microsoft.Office.Interop.Excel.Application();
                    Microsoft.Office.Interop.Excel.Workbook Workbook = Xlsx.Workbooks.Add(System.Reflection.Missing.Value);
                    Microsoft.Office.Interop.Excel.Worksheet Worksheet = (Excel.Worksheet)Workbook.Worksheets.get_Item(1);
                    
                    object[,] OutputRows = new object[RowCount, ColCount];//, DataReader.FieldCount];
                    object[] Output = new object[ColCount];//[DataReader.FieldCount];


                    int ExcelStartRowIndex = 1;

                    //building and writing headers
                    if (IncludeHeaders)
                    {
                        object[] HeaderRow = new object[DataReader.FieldCount];

                        for (int ColumnIndex = 0; ColumnIndex < DataReader.FieldCount; ColumnIndex++)
                        {
                            string ValueToWrite = DataReader.GetName(ColumnIndex);
                            if (ValueToWrite.Length > 0)
                            {
                                if (ValueToWrite.Substring(0, 1) == "0")
                                {
                                    ValueToWrite = "'" + ValueToWrite;
                                }
                            }
                            HeaderRow[ColumnIndex] = ValueToWrite;
                        }
                        Worksheet.Range[Worksheet.Cells[ExcelStartRowIndex, 1], Worksheet.Cells[1, DataReader.FieldCount]].Value = HeaderRow;
                        ExcelStartRowIndex = 2;
                    }

                    //writing rows
                    int ExcelRowIndex = 1;
                    while (DataReader.Read())
                    {
                        DataReader.GetValues(Output);

                        int ExcelColIndex = 1;
                        foreach (object OutputField in Output)
                        {
                            string ValueToWrite = OutputField.ToString();
                            if (ValueToWrite.Length > 0)
                            {
                                if (ValueToWrite.Substring(0, 1) == "0")
                                {
                                    ValueToWrite = "'" + ValueToWrite;
                                }
                            }
                            OutputRows[ExcelRowIndex - 1, ExcelColIndex - 1] = ValueToWrite;
                            ExcelColIndex++;
                        }

                        ExcelRowIndex++;
                    }
                    //we want to write to the sheet as sparingly as possible, because it is slow
                    //so we build 2d list OutputRows and write to sheet once
                    if (OutputRows.Length > 0)
                    {
                        Worksheet.Range[Worksheet.Cells[ExcelStartRowIndex, 1], Worksheet.Cells[RowCount, DataReader.FieldCount]].Value = OutputRows;
                    }

                    Workbook.SaveAs(ExportPath + "\\" + TableToExport + ".xlsx");
                    Workbook.Close();
                    Xlsx.Quit();


                }
                else //anything other than excel
                {
                    //build header
                    List<string> TableColumns = new List<string>();
                    for (int ColumnIndex = 0; ColumnIndex < DataReader.FieldCount; ColumnIndex++)
                    {
                        TableColumns.Add(DataReader.GetName(ColumnIndex));
                    }

                    StreamWriter sw = new StreamWriter(FileExportPath);
                    FilesCreated++;
                    object[] Output = new object[DataReader.FieldCount];

                    List<string> FixedWidthColumnNames = new List<string>();
                    List<int> FixedWidthColumnLengths = new List<int>();

                    string HeaderRow = "";
                    //write header
                    if (IncludeHeaders)
                    {
                        if (Delimeter == "FIXED WIDTH")
                        {

                            for (int ColumnIndex = 0; ColumnIndex < DataReader.FieldCount; ColumnIndex++)
                            {
                                string CurrentColumnName = DataReader.GetName(ColumnIndex);
                                string CurrentColumnLengthQuery = "";

                                if (FixedWidthColumnLengthMethod == "MAX LEN")
                                {
                                    CurrentColumnLengthQuery = "SELECT MAX(LEN([" + CurrentColumnName + "])) FROM [" + TableToExport + "]";
                                }
                                else if (FixedWidthColumnLengthMethod == "COL_LENGTH")
                                {
                                    CurrentColumnLengthQuery = "SELECT COL_LENGTH('[" + TableToExport + "]', '[" + CurrentColumnName + "]')";
                                    Console.Write(CurrentColumnLengthQuery);
                                }

                                using (SqlConnection Conn2 = new SqlConnection(ConnString))
                                {
                                    Conn2.Open();
                                    SqlCommand CurrentColumnLengthCmd = new SqlCommand(CurrentColumnLengthQuery, Conn2);
                                    SqlDataReader CurrentColumnLengthDataReader = CurrentColumnLengthCmd.ExecuteReader();
                                    while (CurrentColumnLengthDataReader.Read())
                                    {
                                        FixedWidthColumnNames.Add(CurrentColumnName);
                                        Console.Write(CurrentColumnLengthDataReader.GetValue(0).ToString());
                                        int FixedWidthColumnLength = Int32.Parse(CurrentColumnLengthDataReader.GetValue(0).ToString());
                                        if (FixedWidthColumnLength == 0)
                                        {
                                            FixedWidthColumnLength = 1;
                                        }
                                        FixedWidthColumnLengths.Add(FixedWidthColumnLength);
                                    }
                                }
                            }

                            //create seperate file _COLUMN_DEFINITIONS.txt
                            //COLUMN NAME LENGTH (for each column)
                            using (TextWriter FixedWidthColumnDefinitionTextWriter = new StreamWriter(ExportPath + "\\" + TableToExport + "_FIXED_WIDTH_COLUMN_DEFINITIONS.txt", true))
                            {
                                for (int FwColIndex = 0; FwColIndex < FixedWidthColumnNames.Count; FwColIndex++)
                                {
                                    FixedWidthColumnDefinitionTextWriter.WriteLine(FixedWidthColumnNames[FwColIndex] + " " + FixedWidthColumnLengths[FwColIndex].ToString());
                                }
                            }
                        }
                        else
                        {
                            HeaderRow = string.Join(Qualifier + Delimeter + Qualifier, TableColumns);
                            HeaderRow = Qualifier + HeaderRow + Qualifier;
                            sw.WriteLine(HeaderRow);
                        }
                    }

                    //write rows
                    int RowIndex = 0;
                    int SplitFileIndex = 0;
                    while (DataReader.Read())
                    {
                        DataReader.GetValues(Output);
                        string CurrentRow = "";

                        if (Delimeter == "FIXED WIDTH")
                        {
                            CurrentRow = "";
                            int FwFieldIndex = 0;
                            foreach (object CurrentField in Output)
                            {
                                int CurrentFieldMaxLength = FixedWidthColumnLengths[FwFieldIndex];
                                CurrentRow += CurrentField;

                                int SpacesNeeded = CurrentFieldMaxLength - CurrentField.ToString().Length;
                                int SpaceIndex = 0;
                                while (SpaceIndex < SpacesNeeded)
                                {
                                    CurrentRow += " ";
                                    SpaceIndex++;
                                }
                                FwFieldIndex++;
                            }
                        }
                        else
                        {
                            CurrentRow = string.Join(Qualifier + Delimeter + Qualifier, Output);
                            CurrentRow = Qualifier + CurrentRow + Qualifier;
                        }

                        if (SizeLimit > 0)
                        {
                            if ((SizeLimitType == "ROW"  && RowIndex % SizeLimit == 0 && RowIndex > 0) ||
                                (SizeLimitType == "SIZE" && (int)sw.BaseStream.Length % (SizeLimit * 1048576) == 0 && RowIndex > 0 && (int)sw.BaseStream.Length > 0))
                            {
                                SplitFileIndex++;
                                sw.Close();
                                sw = new StreamWriter(FileExportPathBase + "-" + SplitFileIndex.ToString() + "." + Extension);
                                FilesCreated++;
                                sw.Write("");
                                if (IncludeHeaders && IncludeHeaderInSplitFiles)
                                {
                                    sw.WriteLine(HeaderRow);
                                }
                            }
                        }
                        sw.WriteLine(CurrentRow);

                        RowIndex++;
                    }

                    sw.Close();

                }

                if (ConsoleOutput) { Console.WriteLine("Exported table to " + FileExportPath); }
            }

            if (ConsoleOutput) { Console.WriteLine(""); }

            return FilesCreated;
        }



        //------------------------------------------------------------------------------------
        // OTHER
        //------------------------------------------------------------------------------------

        public void PrepareValueForImport(ref string Value)
        {
            //remove non-standard chars
            Value = Regex.Replace(Value, @"[^\u0000-\u007F]+", "");

            //limit field length to 255
            if (Value.Length >= 255)
            {
                Value = Value.Substring(0, 255);
            }
        }


    }
}

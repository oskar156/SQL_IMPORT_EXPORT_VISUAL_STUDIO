/* TABLE OF CONTENTS
 * 
 * IMPORT FUNCTIONS
 * public List<string> GetFilesToImport(string ImportPath, string Extension, bool ConsoleOutput = true)
 * public DataTable ReadFileIntoDataTableWithColumns(string FilePath, string Delimeter, string FixedWidthColumnFilePath = "", bool ConsoleOutput = true)
 * public List<DataTable> ReadExcelFileIntoDataTablesWithColumns(string FilePath, ref List<string> TableNames, bool ConsoleOutput = true)
 * public void CreateTableInSqlVarchar(string TableName, string Server, string Database, DataTable DtTable, string Delimeter, string ColumnTypeMethod = "DEFAULT VARCHAR", string ColumnTypeFilePath = "", bool ConsoleOutput = true)
 * public void CreateTablesInSqlVarchar(List<string> TableNames, string Server, string Database, List<DataTable> DataTables, string Delimeter, string ColumnTypeMethod = "DEFAULT VARCHAR", string ColumnTypeFilePath = "", bool ConsoleOutput = true)
 * public void ReadFileIntoDataTableWithRowsAndInsertIntoSqlTable(string FilePath, string TableName, string Server, string Database, DataTable BaseDtTable, int BatchLimit, string Delimeter, bool ConsoleOutput = true)
 * public void ReadFileIntoDataTableWithRowsAndInsertIntoSqlTableFast(string FilePath, string TableName, string Server, string Database, DataTable BaseDtTable, int BatchLimit, string Delimeter, bool ConsoleOutput = true)
 * public void ReadExcelFilePerSheetIntoDataTablesWithRowsAndInsertIntoSqlTables(string FilePath, List<string> TableNames, string Server, string Database, List<DataTable> DataTables, int BatchLimit, string Delimeter, bool ConsoleOutput = true)
 * public void ReadExcelFilePerSheetIntoDataTablesWithRowsAndInsertIntoSqlTablesFast(string FilePath, List<string> TableNames, string Server, string Database, List<DataTable> DataTables, int BatchLimit, string Delimeter, bool ConsoleOutput = true)
 * public void InsertDataTableUsingSqlBulkCopy(ref string ConnString, ref string TableName, ref DataTable TempDataTable, ref int RowIndex, bool ConsoleOutput = true)
 * 
 * EXPORT FUNCTIONS
 * public List<string> GetListofTablesFromSqlDb(ConnectionInfo ConnectionInfo, List<string> ListOfTablesToSearchFor, bool ConsoleOutput = true)
 * public List<string> GetListofTablesFromSqlDb(ConnectionInfo ConnectionInfo, string RegexSearchPattern = "", bool ConsoleOutput = true)
 * public List<string> GetListOfColumnsForTable(ConnectionInfo ConnectionInfo, bool ConsoleOutput = true)
 * public int ExportTableFromSqlTExportTableFromSqlServerToFileoFile(ConnectionInfo ConnectionInfo, string TableToExport, string ExportPath, string Extension, string Delimeter, string Qualifier, bool QualifyEveryField, bool RemoveQualInVal, bool IncludeHeaders, string FixedWidthColumnLengthMethod, decimal SizeLimit, string SizeLimitType, bool IncludeHeaderInSplitFiles, string SelectText = "", string FromText = "", string WhereText = "", string GroupBy = "", string OrderBy = "", bool ConsoleOutput = true)
 *
 * OTHER FUNCTION
 * public void PrepareValueForImport(ref string Value)
 * public Tuple<string, int> ParseColumnWidthLine(string line)
 * public Tuple<string, string> ParseColumnTypeLine(string line)
 * 
*/

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.VisualBasic.FileIO; //for TextFieldParser (also right click project > add > references > Microsoft.VisualBasic.FileIO)
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Data.Common;
using Snowflake.Data.Client; //snowflake

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
                if (ConsoleOutput) { Console.WriteLine("Getting file to import from " + ImportPath + "..."); }
            }
            else //if ImportPath is a direcotry, import only every file in that directory that matches Extension
            {
                if (ConsoleOutput) { Console.WriteLine("Getting file(s) to import from " + ImportPath + "*." + Extension + "..."); }
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
        public DataTable ReadFileIntoDataTableWithColumns(string FilePath, string Delimeter, string FixedWidthColumnFilePath = "", bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.Write("Reading file into DataTable with Columns... "); }

            DataTable DtTable = new DataTable();
            Helpers helpers = new Helpers();

            if (Delimeter == "FIXED WIDTH")
            {
                var ColumnDefinitionFile = File.ReadLines(FixedWidthColumnFilePath);
                foreach (var line in ColumnDefinitionFile)
                {
                    Tuple<string, int> ColumnDefinition = helpers.ParseColumnWidthLine(line);
                    string ColumnName = ColumnDefinition.Item1;
                    int ColumnLength = ColumnDefinition.Item2;
                    //Console.WriteLine(ColumnName + " " + ColumnLength);
                    DataColumn DataColumn = new DataColumn(ColumnName);//, typeof(string));
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

            if (ConsoleOutput) { Console.WriteLine("DataTable with columns created"); }

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
                    //if (TableName.Length > 32)
                    //{
                     //   TableName = TableName.Substring(0, 31);
                    //}
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
        
        public void CreateTablesInSqlServerVarchar(List<string> TableNames, ConnectionInfo ConnectionInfo, List<DataTable> DataTables, string Delimeter, string ColumnTypeMethod = "DEFAULT VARCHAR", string ColumnTypeFilePath = "", bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Creating " + DataTables.Count.ToString() + "tables in sql... "); }

            List<string> ColumnTypes = new List<string>();
            Helpers helpers = new Helpers();
            if (ColumnTypeFilePath != "" && ColumnTypeMethod == "FILE PATH")
            {
                var ColumnDefinitionFile = File.ReadLines(ColumnTypeFilePath);
                foreach (var line in ColumnDefinitionFile)
                {
                    Tuple<string, string> ColumnDefinition = helpers.ParseColumnTypeLine(line);
                    string ColumnName = ColumnDefinition.Item1;
                    string ColumnType = ColumnDefinition.Item2;
                    //Console.WriteLine(ColumnName + " " + ColumnType);
                    //DataColumn DataColumn = new DataColumn(ColumnName);//, typeof(string));
                    //DataColumn.MaxLength = ColumnLength;
                    //DtTable.Columns.Add(DataColumn);
                    ColumnTypes.Add(ColumnType);

                }
            }

            int index = 0;
            foreach (DataTable DataTable in DataTables)
            {
                string ColumnsForTableCreationQuery = "";
                int ColIndex = 0;
                string TableName = TableNames[index];

                foreach (DataColumn Column in DataTable.Columns)
                {
                    ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + "[" + Column.ColumnName + "] ";

                    if (Delimeter == "FIXED WIDTH" && ColumnTypeMethod == "DEFAULT VARCHAR")
                    {
                        ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + " VARCHAR(" + Column.MaxLength.ToString() + "),";
                    }
                    else if (Delimeter == "FIXED WIDTH" && ColumnTypeMethod == "FILE PATH")
                    {
                        ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + " " + ColumnTypes[ColIndex] + ",";
                    }
                    else if (Delimeter != "FIXED WIDTH" && ColumnTypeMethod == "DEFAULT VARCHAR")
                    {
                        ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + " VARCHAR(255),";
                    }
                    else if (Delimeter != "FIXED WIDTH" && ColumnTypeMethod == "FILE PATH")
                    {
                        ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + " " + ColumnTypes[ColIndex] + ",";
                    }
                }
                ColumnsForTableCreationQuery = ColumnsForTableCreationQuery.Substring(0, ColumnsForTableCreationQuery.Length - 1);

                string ConnString = @"Server=" + ConnectionInfo.Server + ";Database=" + ConnectionInfo.Database + ";Trusted_Connection = True;";
                using (SqlConnection Conn = new SqlConnection(ConnString))
                {
                    Conn.Open();
                    string TableCreationQuery = "CREATE TABLE [" + TableName + "] (  " + ColumnsForTableCreationQuery + ")";
                    //Console.WriteLine(TableCreationQuery);
                    SqlCommand Cmd = new SqlCommand(TableCreationQuery, Conn);
                    Cmd.ExecuteNonQuery();
                }
                if (ConsoleOutput)
                {
                    if (Delimeter != "FIXED WIDTH" && ColumnTypeMethod == "DEFAULT VARCHAR")
                    {
                        Console.WriteLine("Created Table " + ConnectionInfo.Server + "." + ConnectionInfo.Database + "..[" + TableName + "] (all columns VARCHAR(255))");
                    }
                    else if (Delimeter == "FIXED WIDTH" && ColumnTypeMethod == "DEFAULT VARCHAR")
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
        public void CreateTablesInSnowflakeVarchar(List<string> TableNames, ConnectionInfo ConnectionInfo, Snowflake Snowflake, List<DataTable> DataTables, string Delimeter, string ColumnTypeMethod = "DEFAULT VARCHAR", string ColumnTypeFilePath = "", bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Creating " + DataTables.Count.ToString() + "tables in sql... "); }

            List<string> ColumnTypes = new List<string>();
            Helpers helpers = new Helpers();
            if (ColumnTypeFilePath != "" && ColumnTypeMethod == "FILE PATH")
            {
                var ColumnDefinitionFile = File.ReadLines(ColumnTypeFilePath);
                foreach (var line in ColumnDefinitionFile)
                {
                    Tuple<string, string> ColumnDefinition = helpers.ParseColumnTypeLine(line);
                    string ColumnName = ColumnDefinition.Item1;
                    string ColumnType = ColumnDefinition.Item2;
                    ColumnTypes.Add(ColumnType);
                }
            }

            int index = 0;
            foreach (DataTable DataTable in DataTables)
            {
                string ColumnsForTableCreationQuery = "";
                int ColIndex = 0;
                string TableName = TableNames[index];

                foreach (DataColumn Column in DataTable.Columns)
                {
                    ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + "[" + Column.ColumnName + "] ";

                    if (Delimeter == "FIXED WIDTH" && ColumnTypeMethod == "DEFAULT VARCHAR")
                    {
                        ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + " VARCHAR(" + Column.MaxLength.ToString() + "),";
                    }
                    else if (Delimeter == "FIXED WIDTH" && ColumnTypeMethod == "FILE PATH")
                    {
                        ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + " " + ColumnTypes[ColIndex] + ",";
                    }
                    else if (Delimeter != "FIXED WIDTH" && ColumnTypeMethod == "DEFAULT VARCHAR")
                    {
                        ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + " VARCHAR(255),";
                    }
                    else if (Delimeter != "FIXED WIDTH" && ColumnTypeMethod == "FILE PATH")
                    {
                        ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + " " + ColumnTypes[ColIndex] + ",";
                    }
                }
                ColumnsForTableCreationQuery = ColumnsForTableCreationQuery.Substring(0, ColumnsForTableCreationQuery.Length - 1);
                ColumnsForTableCreationQuery = ColumnsForTableCreationQuery + ";";

                string TableCreationQuery = "CREATE OR REPLACE TABLE \"" + TableName + "\"  AS (  " + ColumnsForTableCreationQuery + ")";
                Snowflake.Execute(TableCreationQuery);
                if (ConsoleOutput)
                {
                    if (Delimeter != "FIXED WIDTH" && ColumnTypeMethod == "DEFAULT VARCHAR")
                    {
                        Console.WriteLine("Created Table " + ConnectionInfo.Environ + "-" + ConnectionInfo.Server + "-" + ConnectionInfo.Database + "-\"" + TableName + "\" (all columns VARCHAR(255))");
                    }
                    else if (Delimeter == "FIXED WIDTH" && ColumnTypeMethod == "DEFAULT VARCHAR")
                    {
                        Console.WriteLine("Created Table " + ConnectionInfo.Environ + "-" + ConnectionInfo.Server + "." + ConnectionInfo.Database + "-\"" + TableName + "\" (all columns VARCHAR(N))");
                    }
                    else
                    {
                        Console.WriteLine("Created Table " + ConnectionInfo.Environ + "-" + ConnectionInfo.Server + "." + ConnectionInfo.Database + "-\"" + TableName + "\" ");
                    }
                }

                index++;
            }
        }


        public void ReadFileIntoDataTableWithRowsAndInsertIntoSqlServerTable(string FilePath, string TableName, ConnectionInfo ConnectionInfo, DataTable BaseDtTable, int BatchLimit, string Delimeter, bool DoubleQuoted, bool FasterImport, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Reading file rows... "); }

            string ConnString = @"Server=" + ConnectionInfo.Server + ";Database=" + ConnectionInfo.Database + ";Trusted_Connection = True;";

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
                        if(FasterImport)
                        {
                            for (int cf = 0; cf < FieldData.Length; cf++)
                            {
                                PrepareValueForImport(ref FieldData[cf]);
                            }
                        }
                        if (Row != 0 || Delimeter == "FIXED WIDTH") //skip header
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

                    //if (ConsoleOutput) { if (Row > 0 && Row % 1000000 == 0) { Console.WriteLine("ROW " + Row);} }

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
            if (ConsoleOutput) { Console.WriteLine(""); }
        }
        public void ReadFileIntoDataTableWithRowsAndInsertIntoSnowflakeTable(string FilePath, string TableName, ConnectionInfo ConnectionInfo, Snowflake Snowflake, DataTable BaseDtTable, int BatchLimit, string Delimeter, bool DoubleQuoted, bool FasterImport, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Staging file... "); }
            string StageName = "WINFORM_TEMP_STAGE";
            Snowflake.StageFile(FilePath, StageName);
            if (ConsoleOutput) { Console.WriteLine("File in staging area. "); }

            if (ConsoleOutput) { Console.WriteLine("Importing file... "); }
            Snowflake.ImportFile(FilePath, StageName, BaseDtTable, TableName, Delimeter);
            if (ConsoleOutput) { Console.WriteLine("File imported. "); }

            if (ConsoleOutput) { Console.WriteLine(""); }
        }

        public void ReadExcelFilePerSheetIntoDataTablesWithRowsAndInsertIntoSqlServerTables(string FilePath, List<string> TableNames, ConnectionInfo ConnectionInfo, List<DataTable> DataTables, int BatchLimit, string Delimeter, bool FasterImport, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Reading Excel file rows... "); }

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
                                //blank excel cells will be skipped using openxml
                                //so, need to check cell reference, add empty values, update index
                                string ColumnLetters = GetColumnName(Cell.CellReference);
                                int ColumnLettersIndex = GetColumnIndexFromName(ColumnLetters).Value;

                                while ((ColumnLettersIndex - 1) != (CellIndex))
                                {
                                    TempRow.ItemArray.Append("");
                                    TempRow[CellIndex] = "";
                                    CellIndex++;
                                }

                                string FinalCellValue = "";
                                SharedStringTablePart StringTablePart = SpreadSheetDocument.WorkbookPart.SharedStringTablePart;

                                if (Cell.DataType != null)
                                {
                                    string value = Cell.CellValue.InnerXml;
                                    if (Cell.DataType.Value == CellValues.SharedString)
                                    {
                                        FinalCellValue = StringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                                    }
                                    else
                                    {
                                        FinalCellValue = value;
                                    }
                                }
                                else
                                {
                                    FinalCellValue = "";
                                }

                                //for each field in the row
                                if (FasterImport)
                                {
                                    PrepareValueForImport(ref FinalCellValue);
                                }
                                

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
            if (ConsoleOutput) { Console.WriteLine(""); }
        }




        public void InsertDataTableUsingSqlBulkCopy(ref string ConnString, ref string TableName, ref DataTable TempDataTable, ref int RowIndex, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.Write("\r" + $"{RowIndex:n0}" + " rows read. "); }
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

        //------------------------------------------------------------------------------------
        // EXPORT
        //------------------------------------------------------------------------------------
        public List<string> GetListofTablesFromSqlServerDb(ConnectionInfo ConnectionInfo, List<string> ListOfTablesToSearchFor, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Getting List of Tables from Sql Server"); }

            List<string> Tables = new List<string>();

            //convert to upper so we can do case-insensitive matching
            for(int t = 0; t < ListOfTablesToSearchFor.Count; t++)
            {
                ListOfTablesToSearchFor[t] = ListOfTablesToSearchFor[t].ToUpper();
            }

            string ConnString = @"Server=" + ConnectionInfo.Server + ";Database=" + ConnectionInfo.Database + ";Trusted_Connection = True;";
            using (SqlConnection Conn = new SqlConnection(ConnString))
            {
                Conn.Open();
                DataTable TablesInSqlDb = Conn.GetSchema("Tables");

                int TableIndex = 0;
                foreach (DataRow Row in TablesInSqlDb.Rows)
                {
                    string TableName = Row[2].ToString();

                    if (ListOfTablesToSearchFor.Contains(TableName.ToUpper())) //convert to upper so we can do case-insensitive matching
                    {
                        Tables.Add(TableName);
                    }

                    TableIndex++;
                }
            }

            Tables.Sort();

            if (ConsoleOutput) { Console.WriteLine(Tables.Count.ToString() + " tables found"); }
            return Tables;
        }
        public List<string> GetListofTablesFromSqlServerDb(ConnectionInfo ConnectionInfo, string RegexSearchPattern = "", bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Getting Tables from Sql Server"); }

            List<string> Tables = new List<string>();

            Regex re = new Regex("");
            if (RegexSearchPattern != "")
            {
                re = new Regex("(?i)" + RegexSearchPattern); //(?i) makes is case-insensitive
            }

            string ConnString = @"Server=" + ConnectionInfo.Server + ";Database=" + ConnectionInfo.Database + ";Trusted_Connection = True;";
            using (SqlConnection Conn = new SqlConnection(ConnString))
            {
                Conn.Open();
                DataTable TablesInSqlDb = Conn.GetSchema("Tables");

                int TableIndex = 0;
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

                    TableIndex++;
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
        public List<string> GetListofTablesFromSnowflakeDb(Snowflake Snowflake, List<string> ListOfTablesToSearchFor, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Getting List of Tables from Snowflake"); }

            List<string> Tables = new List<string>();

            //convert to upper so we can do case-insensitive matching
            for (int t = 0; t < ListOfTablesToSearchFor.Count; t++)
            {
                ListOfTablesToSearchFor[t] = ListOfTablesToSearchFor[t].ToUpper();
            }

            string Query = "select table_name from information_schema.tables where table_type = 'BASE TABLE';";
            Snowflake.Execute(Query);

            int TableIndex = 0;
            while (Snowflake.Reader.Read())
            {
                string TableName = Snowflake.Reader.GetString(0);

                if (ListOfTablesToSearchFor.Contains(TableName.ToUpper())) //convert to upper so we can do case-insensitive matching
                {
                    Tables.Add(TableName);
                }

                TableIndex++;
            }

            Tables.Sort();

            if (ConsoleOutput) { Console.WriteLine(Tables.Count.ToString() + " tables found"); }
            return Tables;
        }
        public List<string> GetListofTablesFromSnowflakeDb(Snowflake Snowflake, string RegexSearchPattern = "", bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Getting Tables from Snowflake"); }

            List<string> Tables = new List<string>();

            Regex re = new Regex("");
            if (RegexSearchPattern != "")
            {
                re = new Regex("(?i)" + RegexSearchPattern); //(?i) makes is case-insensitive
            }

            string Query = "select table_name from information_schema.tables where table_type = 'BASE TABLE';";
            Snowflake.Execute(Query);

            int TableIndex = 0;
            while (Snowflake.Reader.Read())
            {
                string TableName = Snowflake.Reader.GetString(0);
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

                TableIndex++;
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



        public List<string> GetListOfColumnsForTable(ConnectionInfo ConnectionInfo, string TableName, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Getting Column Names for [" + ConnectionInfo.Server + "].[" + ConnectionInfo.Database + "]..[" + TableName + "]"); }
            List<string> ColumnNames = new List<string>();

            string ConnString = @"Server=" + ConnectionInfo.Server + ";Database=" + ConnectionInfo.Database + ";Trusted_Connection = True;";
            using (SqlConnection Conn = new SqlConnection(ConnString))
            {
                Conn.Open();
                string SqlQuery = " SELECT COLUMN_NAME as [COLUMN_NAMES] FROM [" + ConnectionInfo.Database + "].information_schema.columns WHERE table_name = '" + TableName + "' ";
                SqlCommand Cmd = new SqlCommand(SqlQuery, Conn);
                SqlDataReader DataReader = Cmd.ExecuteReader();

                while (DataReader.Read())
                {
                    ColumnNames.Add(DataReader.GetValue(0).ToString());
                }

                Conn.Close();
            }

            ColumnNames.Sort();

            return ColumnNames;
        }

        public int ExportTableFromSqlServerToFile(ConnectionInfo ConnectionInfo, string TableToExport, string ExportPath, string Extension, string Delimeter, string Qualifier, bool QualifyEveryField, bool RemoveQualInVal, bool IncludeHeaders, string FixedWidthColumnLengthMethod, decimal SizeLimit, string SizeLimitType, bool IncludeHeaderInSplitFiles, string SelectText = "", string FromText = "", string WhereText = "", string GroupBy = "", string OrderBy = "", bool ConsoleOutput = true)
        {
            int FilesCreated = 0;
            if (ConsoleOutput) { Console.WriteLine("Reading table from SQL Server"); }

            //Write the Query
            string TableExportQuery = "SELECT "; //SELECT
            if (SelectText != "")
            {
                TableExportQuery += " " + SelectText + " ";
            }
            else
            {
                TableExportQuery += " * ";
            }

            TableExportQuery += " FROM [" + TableToExport + "] "; //FROM
            if (FromText != "") { TableExportQuery += " " + FromText + " "; } //from extra
            if (WhereText != "") { TableExportQuery += " WHERE " + WhereText + " "; } //WHERE
            if (GroupBy != "") { TableExportQuery += " GROUP BY " + GroupBy + " "; } //GROUP BY
            if (OrderBy != "") { TableExportQuery += " ORDER BY " + OrderBy + " "; } //ORDER BY
            Console.WriteLine("SQL QUERY:\n" + TableExportQuery);

            string ConnString = @"Server=" + ConnectionInfo.Server + ";Database=" + ConnectionInfo.Database + ";Trusted_Connection = True;";

            SqlDataReader DataReader = null;
            using (SqlConnection Conn = new SqlConnection(ConnString))
            {
                Conn.Open();

                //Run the Query
                SqlCommand Cmd = new SqlCommand(TableExportQuery, Conn);

                try
                {
                    DataReader = Cmd.ExecuteReader();
                }
                catch
                {
                    Console.WriteLine("INVALID QUERY!");
                    return 0;
                }

                //https://learn.microsoft.com/en-us/dotnet/api/system.data.sqlclient.sqldatareader?view=netframework-4.8.1#properties
                int RowCount = 0;// DataReader.RecordsAffected; RETURNS -1 WHEN I TRIED IT
                
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
                    int ColCount = DataReader.FieldCount;
                    using (SqlConnection Conn2 = new SqlConnection(ConnString))
                    {
                        Conn2.Open();
                        string CountQuery = "SELECT COUNT(*) as ROW_COUNT FROM [" + TableToExport + "]";
                        SqlCommand CountCmd = new SqlCommand(CountQuery, Conn2);
                        SqlDataReader CountDataReader = CountCmd.ExecuteReader();

                        while (CountDataReader.Read())
                        {
                            RowCount = Int32.Parse(CountDataReader.GetValue(0).ToString()) + 1;
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
                            //if (ExcelRowIndex == 1 || ExcelRowIndex == 2) { Console.Write(ValueToWrite); }
                            OutputRows[ExcelRowIndex - 1, ExcelColIndex - 1] = ValueToWrite;
                            ExcelColIndex++;
                        }

                        if (ConsoleOutput) { if (ExcelRowIndex > 0 && ExcelRowIndex % 1000000 == 0) { Console.Write("\rRows exported: " + $"{ExcelRowIndex:n0}"); } }
                        ExcelRowIndex++;
                    }
                    //we want to write to the sheet as sparingly as possible, because it is slow
                    //so we build 2d list OutputRows and write to sheet once
                    if (OutputRows.Length > 0)
                    {
                        Worksheet.Range[Worksheet.Cells[ExcelStartRowIndex, 1], Worksheet.Cells[RowCount, DataReader.FieldCount]].Value = OutputRows;
                        //insufficient memory - may need to batch the export
                    }

                    Workbook.SaveAs(ExportPath + "\\" + TableToExport + ".xlsx");
                    FilesCreated++;
                    Workbook.Close();
                    Xlsx.Quit();

                    if (ConsoleOutput) {if (ExcelRowIndex % 1000000 != 0) { Console.WriteLine("\rRows exported: " + $"{ExcelRowIndex:n0}"); }}
                }
                else //anything other than excel
                {
                    int FieldCount = DataReader.FieldCount;
                    //build header
                    List<string> TableColumns = new List<string>();
                    for (int ColumnIndex = 0; ColumnIndex < FieldCount; ColumnIndex++)
                    {
                        TableColumns.Add(DataReader.GetName(ColumnIndex));
                    }

                    StreamWriter sw = new StreamWriter(FileExportPath);
                    FilesCreated++;
                    object[] Output = new object[FieldCount];

                    List<string> FixedWidthColumnNames = new List<string>();
                    List<int> FixedWidthColumnLengths = new List<int>();

                    string HeaderRow = "";
                    //write header
                    if (IncludeHeaders)
                    {
                        if (Delimeter == "FIXED WIDTH")
                        {

                            for (int ColumnIndex = 0; ColumnIndex < FieldCount; ColumnIndex++)
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
                                    //Console.Write(CurrentColumnLengthQuery);
                                }

                                using (SqlConnection Conn2 = new SqlConnection(ConnString))
                                {
                                    Conn2.Open();
                                    SqlCommand CurrentColumnLengthCmd = new SqlCommand(CurrentColumnLengthQuery, Conn2);
                                    SqlDataReader CurrentColumnLengthDataReader = CurrentColumnLengthCmd.ExecuteReader();
                                    while (CurrentColumnLengthDataReader.Read())
                                    {
                                        FixedWidthColumnNames.Add(CurrentColumnName);
                                        //Console.Write(CurrentColumnLengthDataReader.GetValue(0).ToString());
                                        int FixedWidthColumnLength = 0;
                                        try
                                        {
                                            FixedWidthColumnLength = Int32.Parse(CurrentColumnLengthDataReader.GetValue(0).ToString());
                                        }
                                        catch
                                        {
                                            FixedWidthColumnLength = 0;
                                        }

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
                            //if(QualifyEveryField)
                            //{
                            //    HeaderRow = string.Join(Qualifier + Delimeter + Qualifier, TableColumns);
                            //    HeaderRow = Qualifier + HeaderRow + Qualifier;
                            //}
                            //else
                            //{
                                foreach (string TblCol in TableColumns)
                                {
                                    //clean
                                    string TblColClean = TblCol;
                                    if(RemoveQualInVal && Qualifier != "" && Qualifier != null)
                                    {
                                        TblColClean = TblColClean.Replace(Qualifier, "");
                                    }

                                    //write to HeaderRow
                                    if(QualifyEveryField || TblColClean.Contains(Delimeter))
                                    {
                                        HeaderRow = HeaderRow + Qualifier + TblColClean + Qualifier + Delimeter;
                                    }
                                    else
                                    {
                                        HeaderRow = HeaderRow + TblColClean + Delimeter;
                                    }

                                }
                                //remove last delimeter
                                HeaderRow = HeaderRow.Remove(HeaderRow.Length - 1);
                            //}
                            sw.WriteLine(HeaderRow);
                        }
                    }

                    //write rows
                    int RowIndex = 0;
                    int SplitFileIndex = 0;
                    while (DataReader.Read())
                    {
                        DataReader.GetValues(Output); //breaks with spatial data
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
                            //if (QualifyEveryField)
                            //{
                            //    CurrentRow = string.Join(Qualifier + Delimeter + Qualifier, Output);
                            //    CurrentRow = Qualifier + CurrentRow + Qualifier;
                            //}
                            //else
                            //{
                                foreach (object CrFld in Output)
                                {
                                    //clean field, handle nulls
                                    string CrFldClean = "";
                                    if(CrFld.GetType() == typeof(string))
                                    {
                                        if (CrFld == null)
                                        {
                                            CrFldClean = "";
                                        }
                                        else
                                        {
                                            CrFldClean = CrFld.ToString();
                                        }
                                    }
                                    else if (CrFld.GetType() != typeof(string))
                                    {
                                        if (CrFld == null)
                                        {
                                            CrFldClean = "";
                                        }
                                        else
                                        {
                                            CrFldClean = CrFld.ToString();
                                        }
                                    }

                                    if(RemoveQualInVal && Qualifier != "" && Qualifier != null)
                                    {
                                        CrFldClean = CrFldClean.Replace(Qualifier, "");
                                    }

                                    //write field to CurrentRow
                                    if (QualifyEveryField || CrFldClean.Contains(Delimeter))
                                    {
                                        CurrentRow = CurrentRow + Qualifier + CrFldClean + Qualifier + Delimeter;
                                    }
                                    else
                                    {
                                        CurrentRow = CurrentRow + CrFldClean + Delimeter;
                                    }
                                }
                                //remoev last delimeter
                                CurrentRow = CurrentRow.Remove(CurrentRow.Length - 1);
                            //}
                        }

                        if (SizeLimit > 0)
                        {
                            if (SizeLimitType == "ROW"  && RowIndex % SizeLimit == 0 && RowIndex > 0)
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
                            else if (SizeLimitType == "SIZE" || SizeLimitType == "SIZE1024")
                            {
                                int MbSize = 0;
                                if(SizeLimitType == "SIZE")
                                {
                                    MbSize = 1000000;
                                }
                                else if(SizeLimitType == "SIZE1024")
                                {
                                    MbSize = 1048576;
                                }
                                /*
                                 * check file current size
                                 * check line current size
                                 * add them
                                 * if greater than split then split
                                 * reset SizeCounter
                                 */
                                int FileSize = (int)sw.BaseStream.Length;
                                int CurrentLineSize = (int)CurrentRow.Length * sizeof(System.Char);
                                int FileProjectedSize = FileSize + CurrentLineSize;

                                if (SizeLimit * MbSize <= FileProjectedSize && RowIndex > 0 && FileSize > 0 && CurrentLineSize > 0) 
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
                        }
                        sw.WriteLine(CurrentRow);
                        //every n rows?
                        //sw.Flush();

                        if (ConsoleOutput) { if (RowIndex > 0 && RowIndex % 1000000 == 0) { Console.Write("\rRows exported: " + $"{RowIndex:n0}"); } }

                        RowIndex++;
                    }

                    if (ConsoleOutput) { if (RowIndex % 1000000 != 0) { Console.WriteLine("\rRows exported: " + $"{RowIndex:n0}"); } }

                    sw.Close();

                }

                if (ConsoleOutput) { Console.WriteLine("Exported table to " + FileExportPath); }
            }

            if (ConsoleOutput) { Console.WriteLine(""); }

            return FilesCreated;
        }

        public int ExportTableFromSnowflakeToFile(Snowflake Snowflake, string TableToExport, string ExportPath, string Extension, string Delimeter, string Qualifier, bool QualifyEveryField, bool RemoveQualInVal, bool IncludeHeaders, string FixedWidthColumnLengthMethod, decimal SizeLimit, string SizeLimitType, bool IncludeHeaderInSplitFiles, string SelectText = "", string FromText = "", string WhereText = "", string GroupBy = "", string OrderBy = "", bool ConsoleOutput = true)
        {
            int FilesCreated = 0;
            if (ConsoleOutput) { Console.WriteLine("Reading table from Snowflake"); }

            //Write the Query
            string TableExportQuery = "SELECT "; //SELECT
            if (SelectText != "")
            {
                TableExportQuery += " " + SelectText + " ";
            }
            else
            {
                TableExportQuery += " * ";
            }

            TableExportQuery += " FROM \"" + TableToExport + "\" "; //FROM
            if (FromText != "") { TableExportQuery += " " + FromText + " "; } //from extra
            if (WhereText != "") { TableExportQuery += " WHERE " + WhereText + " "; } //WHERE
            if (GroupBy != "") { TableExportQuery += " GROUP BY " + GroupBy + " "; } //GROUP BY
            if (OrderBy != "") { TableExportQuery += " ORDER BY " + OrderBy + " "; } //ORDER BY
            TableExportQuery += ";";

            Console.WriteLine("SQL QUERY:\n" + TableExportQuery);

            //string ConnString = @"Server=" + ConnectionInfo.Server + ";Database=" + ConnectionInfo.Database + ";Trusted_Connection = True;";
            try
            {
                Snowflake.Execute(TableExportQuery);
            }
            catch
            {
                Console.WriteLine("INVALID QUERY!");
                return 0;
            }

            int RowCount = 0;// DataReader.RecordsAffected; RETURNS -1 WHEN I TRIED IT

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
                int ColCount = Snowflake.Reader.FieldCount;

                Snowflake Snowflake2 = new Snowflake();
                ConnectionInfo ConnectionInfo = new ConnectionInfo();
                Snowflake2.ConnectToDb(ConnectionInfo);
                Snowflake2.Execute("SELECT COUNT(*) as ROW_COUNT FROM \"" + TableToExport + "\"");
                while (Snowflake2.Reader.Read())
                {
                    RowCount = Int32.Parse(Snowflake2.Reader.GetValue(0).ToString()) + 1;
                }
                Snowflake2.Close();

                Microsoft.Office.Interop.Excel.Application Xlsx = new Microsoft.Office.Interop.Excel.Application();
                Microsoft.Office.Interop.Excel.Workbook Workbook = Xlsx.Workbooks.Add(System.Reflection.Missing.Value);
                Microsoft.Office.Interop.Excel.Worksheet Worksheet = (Excel.Worksheet)Workbook.Worksheets.get_Item(1);

                object[,] OutputRows = new object[RowCount, ColCount];//, DataReader.FieldCount];
                object[] Output = new object[ColCount];//[DataReader.FieldCount];


                int ExcelStartRowIndex = 1;

                //building and writing headers
                if (IncludeHeaders)
                {
                    object[] HeaderRow = new object[Snowflake.Reader.FieldCount];

                    for (int ColumnIndex = 0; ColumnIndex < Snowflake.Reader.FieldCount; ColumnIndex++)
                    {
                        string ValueToWrite = Snowflake.Reader.GetName(ColumnIndex);
                        if (ValueToWrite.Length > 0)
                        {
                            if (ValueToWrite.Substring(0, 1) == "0")
                            {
                                ValueToWrite = "'" + ValueToWrite;
                            }
                        }
                        HeaderRow[ColumnIndex] = ValueToWrite;
                    }
                    Worksheet.Range[Worksheet.Cells[ExcelStartRowIndex, 1], Worksheet.Cells[1, Snowflake.Reader.FieldCount]].Value = HeaderRow;
                    ExcelStartRowIndex = 2;
                }

                //writing rows
                int ExcelRowIndex = 1;
                while (Snowflake.Reader.Read())
                {
                    Snowflake.Reader.GetValues(Output);

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
                        //if (ExcelRowIndex == 1 || ExcelRowIndex == 2) { Console.Write(ValueToWrite); }
                        OutputRows[ExcelRowIndex - 1, ExcelColIndex - 1] = ValueToWrite;
                        ExcelColIndex++;
                    }

                    if (ConsoleOutput) { if (ExcelRowIndex > 0 && ExcelRowIndex % 1000000 == 0) { Console.Write("\rRows exported: " + $"{ExcelRowIndex:n0}"); } }
                    ExcelRowIndex++;
                }
                //we want to write to the sheet as sparingly as possible, because it is slow
                //so we build 2d list OutputRows and write to sheet once
                if (OutputRows.Length > 0)
                {
                    Worksheet.Range[Worksheet.Cells[ExcelStartRowIndex, 1], Worksheet.Cells[RowCount, Snowflake.Reader.FieldCount]].Value = OutputRows;
                    //insufficient memory - may need to batch the export
                }

                Workbook.SaveAs(ExportPath + "\\" + TableToExport + ".xlsx");
                FilesCreated++;
                Workbook.Close();
                Xlsx.Quit();

                if (ConsoleOutput) { if (ExcelRowIndex % 1000000 != 0) { Console.WriteLine("\rRows exported: " + $"{ExcelRowIndex:n0}"); } }
            }
            else //anything other than excel
            {
                int FieldCount = Snowflake.Reader.FieldCount;
                
                //build header
                List<string> TableColumns = new List<string>();
                for (int ColumnIndex = 0; ColumnIndex < FieldCount; ColumnIndex++)
                {
                    TableColumns.Add(Snowflake.Reader.GetName(ColumnIndex));
                }

                StreamWriter sw = new StreamWriter(FileExportPath);
                FilesCreated++;
                object[] Output = new object[FieldCount];

                List<string> FixedWidthColumnNames = new List<string>();
                List<int> FixedWidthColumnLengths = new List<int>();

                string HeaderRow = "";

                //write header
                if (IncludeHeaders)
                {
                    if (Delimeter == "FIXED WIDTH")
                    {
                        Snowflake Snowflake2 = new Snowflake();
                        ConnectionInfo ConnectionInfo  = new ConnectionInfo();
                        Snowflake2.ConnectToDb(ConnectionInfo);

                        for (int ColumnIndex = 0; ColumnIndex < FieldCount; ColumnIndex++)
                        {
                            string CurrentColumnName = Snowflake.Reader.GetName(ColumnIndex);
                            string CurrentColumnLengthQuery = "";

                            if (FixedWidthColumnLengthMethod == "MAX LEN")
                            {
                                CurrentColumnLengthQuery = "SELECT MAX(LEN(\"" + CurrentColumnName + "\")) FROM \"" + TableToExport + "\"";
                            }
                            else if (FixedWidthColumnLengthMethod == "COL_LENGTH")
                            {
                                CurrentColumnLengthQuery = "SELECT COL_LENGTH('\"" + TableToExport + "\"', '\"" + CurrentColumnName + "\"')";
                                //Console.Write(CurrentColumnLengthQuery);
                            }

                            Snowflake2.Execute(CurrentColumnLengthQuery);

                            while (Snowflake2.Reader.Read())
                            {
                                FixedWidthColumnNames.Add(CurrentColumnName);
                                //Console.Write(CurrentColumnLengthDataReader.GetValue(0).ToString());
                                int FixedWidthColumnLength = 0;
                                try
                                {
                                    FixedWidthColumnLength = Int32.Parse(Snowflake2.Reader.GetValue(0).ToString());
                                }
                                catch
                                {
                                    FixedWidthColumnLength = 0;
                                }

                                if (FixedWidthColumnLength == 0)
                                {
                                    FixedWidthColumnLength = 1;
                                }
                                FixedWidthColumnLengths.Add(FixedWidthColumnLength);
                            }
                        }
                        Snowflake2.Close();

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
                        foreach (string TblCol in TableColumns)
                        {
                            //clean
                            string TblColClean = TblCol;
                            if (RemoveQualInVal && Qualifier != "" && Qualifier != null)
                            {
                                TblColClean = TblColClean.Replace(Qualifier, "");
                            }

                            //write to HeaderRow
                            if (QualifyEveryField || TblColClean.Contains(Delimeter))
                            {
                                HeaderRow = HeaderRow + Qualifier + TblColClean + Qualifier + Delimeter;
                            }
                            else
                            {
                                HeaderRow = HeaderRow + TblColClean + Delimeter;
                            }

                        }
                        //remove last delimeter
                        HeaderRow = HeaderRow.Remove(HeaderRow.Length - 1);
                        //}
                        sw.WriteLine(HeaderRow);
                    }
                }

                //write rows
                int RowIndex = 0;
                int SplitFileIndex = 0;
                while (Snowflake.Reader.Read())
                {
                    Snowflake.Reader.GetValues(Output); //breaks with spatial data
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
                        foreach (object CrFld in Output)
                        {
                            //clean field, handle nulls
                            string CrFldClean = "";
                            if (CrFld.GetType() == typeof(string))
                                {
                                if (CrFld == null)
                                {
                                    CrFldClean = "";
                                }
                                else
                                {
                                    CrFldClean = CrFld.ToString();
                                }
                            }
                            else if (CrFld.GetType() != typeof(string))
                            {
                                if (CrFld == null)
                                {
                                    CrFldClean = "";
                                }
                                else
                                {
                                    CrFldClean = CrFld.ToString();
                                }
                            }

                            if (RemoveQualInVal && Qualifier != "" && Qualifier != null)
                            {
                                CrFldClean = CrFldClean.Replace(Qualifier, "");
                            }

                            //write field to CurrentRow
                            if (QualifyEveryField || CrFldClean.Contains(Delimeter))
                            {
                                CurrentRow = CurrentRow + Qualifier + CrFldClean + Qualifier + Delimeter;
                            }
                            else
                            {
                                CurrentRow = CurrentRow + CrFldClean + Delimeter;
                            }
                        }
                        //remoev last delimeter
                        CurrentRow = CurrentRow.Remove(CurrentRow.Length - 1);
                        //}
                    }

                    if (SizeLimit > 0)
                    {
                        if (SizeLimitType == "ROW" && RowIndex % SizeLimit == 0 && RowIndex > 0)
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
                        else if (SizeLimitType == "SIZE" || SizeLimitType == "SIZE1024")
                        {
                            int MbSize = 0;
                            if (SizeLimitType == "SIZE")
                            {
                                MbSize = 1000000;
                            }
                            else if (SizeLimitType == "SIZE1024")
                            {
                                MbSize = 1048576;
                            }
                            /*
                             * check file current size
                             * check line current size
                             * add them
                             * if greater than split then split
                             * reset SizeCounter
                             */
                            int FileSize = (int)sw.BaseStream.Length;
                            int CurrentLineSize = (int)CurrentRow.Length * sizeof(System.Char);
                            int FileProjectedSize = FileSize + CurrentLineSize;

                            if (SizeLimit * MbSize <= FileProjectedSize && RowIndex > 0 && FileSize > 0 && CurrentLineSize > 0)
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
                    }
                    sw.WriteLine(CurrentRow);
                    //every n rows?
                    //sw.Flush();

                    if (ConsoleOutput) { if (RowIndex > 0 && RowIndex % 1000000 == 0) { Console.Write("\rRows exported: " + $"{RowIndex:n0}"); } }

                    RowIndex++;
                }

                if (ConsoleOutput) { if (RowIndex % 1000000 != 0) { Console.WriteLine("\rRows exported: " + $"{RowIndex:n0}"); } }

                sw.Close();

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

        //need to implement
        public void PrepareValueANSI(ref string Value)
        {
            //remove non-standard chars
            Value = Regex.Replace(Value, @"[^\u0000-\u007F]+", "");

            //limit field length to 255
            if (Value.Length >= 255)
            {
                Value = Value.Substring(0, 255);
            }
        }
        //need to implement
        public void PrepareValueUTF8(ref string Value)
        {
            //remove non-standard chars
            Value = Regex.Replace(Value, @"[^\u0000-\u007F]+", "");

            //limit field length to 255
            if (Value.Length >= 255)
            {
                Value = Value.Substring(0, 255);
            }
        }

        public Tuple<string, int> ParseColumnWidthLine(string line)
        {
            string ColumnName = "";
            int ColumnLength = 0;
            string LineTrimmed = line.Trim();

            if (LineTrimmed[0] == '[')
            {
                Regex re = new Regex("\\[.*\\]");
                string ColumnNameRaw = re.Match(line).ToString();
                string ColumnLengthRaw = LineTrimmed.Substring(ColumnNameRaw.Length).Trim();

                ColumnName = ColumnNameRaw.Substring(1, ColumnNameRaw.Length - 2).Trim();
                ColumnLength = Int32.Parse(ColumnLengthRaw);
            }
            else
            {
                int LastSpaceIndex = LineTrimmed.LastIndexOf(" ");

                ColumnName = LineTrimmed.Substring(0, LastSpaceIndex);
                ColumnLength = Int32.Parse(LineTrimmed.Substring(LastSpaceIndex).Trim());
            }

            return new Tuple<string, int>(ColumnName, ColumnLength);
        }
        public Tuple<string, string> ParseColumnTypeLine(string line)
        {
            string ColumnName = "";
            string ColumnType = "";
            string LineTrimmed = line.Trim();

            if (LineTrimmed[0] == '[')
            {
                Regex re = new Regex("\\[.*\\]");
                string ColumnNameRaw = re.Match(line).ToString();

                ColumnName = ColumnNameRaw.Substring(1, ColumnNameRaw.Length - 2).Trim();
                ColumnType = LineTrimmed.Substring(ColumnNameRaw.Length).Trim();
            }
            else
            {
                int LastSpaceIndex = LineTrimmed.LastIndexOf(" ");

                ColumnName = LineTrimmed.Substring(0, LastSpaceIndex);
                ColumnType = LineTrimmed.Substring(LastSpaceIndex).Trim();
            }

            return new Tuple<string, string>(ColumnName, ColumnType);
        }

        /*
        public T ConvertFromDBVal<T>(object obj)
        {
            //https://stackoverflow.com/questions/870697/unable-to-cast-object-of-type-system-dbnull-to-type-system-string
            if (obj == null || obj == DBNull.Value)
            {
                return default(T); // returns the default value for the type
            }
            else
            {
                return (T)obj;
            }
        }
        */


        //https://stackoverflow.com/questions/3837981/reading-excel-open-xml-is-ignoring-blank-cells
        public static string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);

            return match.Value;
        }

        /// <summary>
        /// Given just the column name (no row index), it will return the zero based column index.
        /// Note: This method will only handle columns with a length of up to two (ie. A to Z and AA to ZZ). 
        /// A length of three can be implemented when needed.
        /// </summary>
        /// <param name="columnName">Column Name (ie. A or AB)</param>
        /// <returns>Zero based index if the conversion was successful; otherwise null</returns>

        
        public static int? GetColumnIndexFromName(string columnName)
        {
            const string Letters = "ZABCDEFGHIJKLMNOPQRSTUVWXY";
            int? columnIndex = null;

            string[] colLetters = Regex.Split(columnName, "([A-Z]+)");
            colLetters = colLetters.Where(s => !string.IsNullOrEmpty(s)).ToArray();

            if (colLetters.Count() <= 2)
            {
                int index = 0;
                foreach (string col in colLetters)
                {
                    List<char> col1 = colLetters.ElementAt(index).ToCharArray().ToList();
                    int? indexValue = Letters.IndexOf(col1.ElementAt(index));

                    if (indexValue != -1)
                    {
                        // The first letter of a two digit column needs some extra calculations
                        if (index == 0 && colLetters.Count() == 2)
                        {
                            columnIndex = columnIndex == null ? (indexValue + 1) * 26 : columnIndex + ((indexValue + 1) * 26);
                        }
                        else
                        {
                            columnIndex = columnIndex == null ? indexValue : columnIndex + indexValue;
                        }
                    }

                    index++;
                }
            }
            return columnIndex;

        }
        


    }
}




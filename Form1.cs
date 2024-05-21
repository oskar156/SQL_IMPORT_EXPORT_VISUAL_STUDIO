/* TABLE OF CONTENTS
 * 
 * IMPORT FUNCTIONS
 * private void ImportButton_Click(object sender, EventArgs e)
 * 
 * private List<string> GetFilesToImport(string ImportPath, string Extension, bool ConsoleOutput = true)
 * private DataTable ReadFileIntoDataTableWithColumns(string FilePath, string Delimeter, string FixedWidthColumnFilePath, bool ConsoleOutput = true)
 * private List<DataTable> ReadExcelFileIntoDataTablesWithColumns(string FilePath, ref List<string> TableNames, bool ConsoleOutput = true)
 * private void CreateTableInSqlVarchar(string TableName, string Server, string Database, DataTable DtTable, string Delimeter, bool ConsoleOutput = true)
 * private void CreateTablesInSqlVarchar(List<string> TableNames, string Server, string Database, List<DataTable> DataTables, string Delimeter, bool ConsoleOutput = true)
 * private void ReadFileIntoDataTableWithRowsAndInsertIntoSqlTable(string FilePath, string TableName, string Server, string Database, DataTable BaseDtTable, int BatchLimit, string Delimeter, bool ConsoleOutput = true)
 * private void ReadExcelFilePerSheetIntoDataTablesWithRowsAndInsertIntoSqlTables(string FilePath, List<string> TableNames, string Server, string Database, List<DataTable> DataTables, int BatchLimit, string Delimeter, bool ConsoleOutput = true)
 * private void InsertDataTableUsingSqlBulkCopy(string ConnString, string TableName, DataTable TempDataTable, int RowIndex, bool ConsoleOutput = true
 * 
 * EXPORT FUNCTIONS
 * private void ExportButton_Click(object sender, EventArgs e)
 * 
 * private List<string> GetListofTablesFromSqlDb(string Server, string Database, List<string> ListOfTablesToSearchFor, bool ConsoleOutput = true)
 * private List<string> GetListofTablesFromSqlDb(string Server, string Database, string RegexSearchPattern = "", bool ConsoleOutput = true)
 * private void ExportTableFromSqlToFile(string Server, string Database, string TableToExport, string ExportPath, string Extension, string Delimeter, string Qualifier, bool IncludeHeaders, string FixedWidthColumnLengthMethod, bool ConsoleOutput = true)
 * 
 * OTHER
 * private void LoadSqlTables_Click(object sender, EventArgs e)
 * private void ExtensionListBoxImport_SelectedIndexChanged(object sender, EventArgs e)
 * private void ImportDelimeterListBox_SelectedIndexChanged(object sender, EventArgs e)
 * private void OutputTypeListBox_SelectedIndexChanged(object sender, EventArgs e
 * private void ExportDelimeterListBox_SelectedIndexChanged(object sender, EventArgs e)
 * private void ExportQualifierListBox_SelectedIndexChanged(object sender, EventArgs e)
 * 
 * DEPENDENCIES
 * DocumentFormat.OpenXml.Packaging - import excel
 * Microsoft.Office.Interop.Excel - export excel
 * Microsoft.VisualBasic.FileIO - TextFieldParser
 * (right click Solution in Solution Explorer, open Manage NuGet..., search, download, install it)
*/


using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Data.SqlClient;
using System.Data.Common; //DbDataReader
using System.Data.OleDb; //OleDbConnection
using System.Drawing;
using System.IO; //for Directory
using System.Linq;
using System.Text.RegularExpressions; //Regex
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

using Microsoft.VisualBasic.FileIO; //for TextFieldParser

using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Drawing; //for SpreadsheetDocument
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Wordprocessing;

using Excel = Microsoft.Office.Interop.Excel;

namespace SQL_SERVER_IMPORT_EXPORT
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();

            //default values
            ExtensionListBoxImport.SetSelected(0, true);
            ImportDelimeterListBox.SetSelected(0, true);

            TablePickerRadioButton.Select();
            OutputTypeListBox.SetSelected(0, true);
            ExportDelimeterListBox.SetSelected(0, true);
            ExportQualifierListBox.SetSelected(0, true);
            FixedWidthColumnLengthMethodListBox.SetSelected(0, true);
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }

        //------------------------------------------------------------------------------------
        // IMPORT
        //------------------------------------------------------------------------------------
        private void ImportButton_Click(object sender, EventArgs e)
        {
            Console.WriteLine("------------------------------");
            Console.WriteLine("Import Button Clicked");
            Console.WriteLine("------------------------------");

            /***************************************
             * CONSTANTS
            ***************************************/
            //averages about 30 seconds on SQL04 for meta format (18 cols) for every million records
            //notes: implement fixed filed import method then test it out gpconsumer import
            //exporting n records from gpconsumer as csvc then trying to import it again wont work
            int BatchLimit = 100000; //100k

            /***************************************
             * USER INPUT
            ***************************************/
            string Server = ServerComboBox.Text;
            string Database = DatabaseComboBox.Text;
            string Extension = ExtensionListBoxImport.Text;

            string Delimeter = ImportDelimeterListBox.Text;
            string ActualDelimeter = "";
            if (Delimeter == "COMMA") { ActualDelimeter = ","; }
            else if (Delimeter == "PIPE") { ActualDelimeter = "|"; }
            else if (Delimeter == "TAB") { ActualDelimeter = "\t"; }
            else if (Delimeter == "FIXED WIDTH") { ActualDelimeter = "FIXED WIDTH"; }

            string ImportToSingleTableName = ImportToSingleTableTextBox.Text.Trim();
            bool ImportToSingleTable = false;
            if (ImportToSingleTableName != "")
            {
                ImportToSingleTable = true;
            }

            string FixedWidthColumnFilePath = FixedWidthColumnFilePathTextBox.Text;
            string ImportPath = ImportPathTextBox.Text;

            Console.WriteLine("Server.Databse: " + Server + "." + Database);
            Console.WriteLine("BatchLimit: " + BatchLimit.ToString());

            /***************************************
             * GET FILES TO IMPORT
            ***************************************/
            List<string> FilesToImport = GetFilesToImport(ImportPath, Extension);

            int FilesImported = 0;
            int TablesCreated = 0;
            string TableName = "";
            List<string> TableNames = new List<string>();
            DataTable BaseDtTable = new DataTable();
            List<DataTable> BaseDtTables = new List<DataTable>();

            //for each file to import
            for (int f = 0; f < FilesToImport.Count; f++)
            {
                string FilePath = FilesToImport[f];

                Console.WriteLine(FilePath);
                if (ImportToSingleTable == false)
                {
                    TableName = System.IO.Path.GetFileNameWithoutExtension(FilePath);
                }
                else if (ImportToSingleTable == true && f == 0)
                {
                    TableName = ImportToSingleTableName;
                }

                string FileExtension = System.IO.Path.GetExtension(FilePath);
                Console.WriteLine("File " + (f + 1).ToString() + ": " + FilePath);

                /***************************************
                 * READ FILE COLUMNS
                ***************************************/
                if (ImportToSingleTable == false || (ImportToSingleTable == true && f == 0))
                {
                    //only need to read columns and create table once if we're importing all the files to a single table
                    if (Extension == "xls*")
                    {
                        BaseDtTables = ReadExcelFileIntoDataTablesWithColumns(FilePath, ref TableNames);
                        //will also add to TableNames list
                    }
                    else
                    {
                        BaseDtTable = ReadFileIntoDataTableWithColumns(FilePath, ActualDelimeter, FixedWidthColumnFilePath);
                    }

                    /***************************************
                     * CREATE TABLE IN SQL
                    ***************************************/
                    if (Extension == "xls*")
                    {
                        CreateTablesInSqlVarchar(TableNames, Server, Database, BaseDtTables, ActualDelimeter);
                        TablesCreated += BaseDtTables.Count;
                    }
                    else
                    {
                        CreateTableInSqlVarchar(TableName, Server, Database, BaseDtTable, ActualDelimeter);
                        TablesCreated++;
                    }
                }

                /***************************************
                 * READ FILE ROWS AND INSERT INTO SQL TABLE
                ***************************************/
                if (Extension == "xls*")
                {
                    ReadExcelFilePerSheetIntoDataTablesWithRowsAndInsertIntoSqlTables(FilePath, TableNames, Server, Database, BaseDtTables, BatchLimit, ActualDelimeter);
                    FilesImported += BaseDtTables.Count;
                }
                else
                {
                    ReadFileIntoDataTableWithRowsAndInsertIntoSqlTable(FilePath, TableName, Server, Database, BaseDtTable, BatchLimit, ActualDelimeter);
                    FilesImported++;
                }

                Console.WriteLine("");
            }

            if (ImportToSingleTable == false)
            {
                Console.WriteLine(FilesImported.ToString() + " files imported to " + Server + "." + Database);
            }
            else if (ImportToSingleTable == true)
            {
                Console.WriteLine(FilesImported.ToString() + " files imported to " + Server + "." + Database + ".." + TableName);
            }

            Console.WriteLine("");
        }
        private List<string> GetFilesToImport(string ImportPath, string Extension, bool ConsoleOutput = true)
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
                FilesToImport.Add(ImportPath);  //if ImportPath is a direcotry, import only every file in that directory that matches Extension
            }
            else
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
        private DataTable ReadFileIntoDataTableWithColumns(string FilePath, string Delimeter, string FixedWidthColumnFilePath, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Reading file into DataTable with Columns..."); }

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

            if (ConsoleOutput) { Console.WriteLine("DataTable with columns created"); }

            return DtTable;
        }
        private List<DataTable> ReadExcelFileIntoDataTablesWithColumns(string FilePath, ref List<string> TableNames, bool ConsoleOutput = true)
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
        private void CreateTableInSqlVarchar(string TableName, string Server, string Database, DataTable DtTable, string Delimeter, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Creating table in sql..."); }

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
                    Console.WriteLine("Created Table " + Server + "." + Database + "..[" + TableName + "] ");
                }
                else
                {
                    Console.WriteLine("Created Table " + Server + "." + Database + "..[" + TableName + "] (all columns VARCHAR(255))");
                }
            }
        }
        private void CreateTablesInSqlVarchar(List<string> TableNames, string Server, string Database, List<DataTable> DataTables, string Delimeter, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Creating " + DataTables.Count.ToString() + "tables in sql..."); }

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
        private void ReadFileIntoDataTableWithRowsAndInsertIntoSqlTable(string FilePath, string TableName, string Server, string Database, DataTable BaseDtTable, int BatchLimit, string Delimeter, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Reading file into DataTable with Rows..."); }

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
                    FileReader.HasFieldsEnclosedInQuotes = true;
                }

                int Row = 0;
                bool LeftoverData = false;
                DataTable TempDtTable = BaseDtTable;

                //got the below code from stack overflow, find the link and paste it here
                while (!FileReader.EndOfData)
                {
                    LeftoverData = true;
                    string[] FieldData = FileReader.ReadFields();

                    /*
                    //Making empty value as null
                    for (int i = 0; i < FieldData.Length; i++)
                    {
                        if (FieldData[i] == "")
                        {
                            FieldData[i] = null;
                        }
                    }
                    */

                    TempDtTable.Rows.Add(FieldData);

                    //when we get to Row BatchLimit, import that chunk into SQL Server
                    //also print to console to help with tracking
                    if (Row != 0 && Row % BatchLimit == 0)
                    {
                        LeftoverData = false;
                        if (ConsoleOutput) { Console.WriteLine(Row.ToString() + " rows read"); }

                        /***************************************
                        * INSERT ROWS TO TABLE
                        ***************************************/
                        using (SqlConnection Conn = new SqlConnection(ConnString))
                        {
                            Conn.Open();
                            using (SqlBulkCopy SqlBulk = new SqlBulkCopy(Conn))
                            {
                                SqlBulk.DestinationTableName = "[dbo].[" + TableName + "]";
                                foreach (var Column in TempDtTable.Columns)
                                {
                                    SqlBulk.ColumnMappings.Add(Column.ToString(), Column.ToString());
                                }

                                if (ConsoleOutput) { Console.WriteLine("Inserting rows to table..."); }

                                SqlBulk.WriteToServer(TempDtTable);

                                if (ConsoleOutput) { Console.WriteLine("Inserted"); }
                            }
                        }

                        //reset TempDtTable
                        //if we don't do this files with about 18 columns/4 million rows cause the program to run out of memory
                        TempDtTable = BaseDtTable; //not sure if this necessary
                        TempDtTable.Rows.Clear(); //definitely necessary
                    }
                    Row++;
                }

                //Importing the remaining data (necessary because of the batching)
                //the script will only end up coming here to insert data if the file is under the BatchLimit of rows
                if (LeftoverData == true)
                {

                    if (ConsoleOutput) { Console.WriteLine(Row.ToString() + " rows read"); }

                    using (SqlConnection Conn = new SqlConnection(ConnString))
                    {
                        Conn.Open();
                        using (SqlBulkCopy SqlBulk = new SqlBulkCopy(Conn))
                        {
                            SqlBulk.DestinationTableName = "[dbo].[" + TableName + "]";
                            foreach (var Column in TempDtTable.Columns)
                            {
                                SqlBulk.ColumnMappings.Add(Column.ToString(), Column.ToString());
                            }


                            if (ConsoleOutput) { Console.WriteLine("Inserting rows to table..."); }

                            SqlBulk.WriteToServer(TempDtTable);

                            if (ConsoleOutput) { Console.WriteLine("Inserted"); }
                        }
                    }
                    TempDtTable = BaseDtTable;
                    TempDtTable.Rows.Clear();
                }
            }
        }
        private void ReadExcelFilePerSheetIntoDataTablesWithRowsAndInsertIntoSqlTables(string FilePath, List<string> TableNames, string Server, string Database, List<DataTable> DataTables, int BatchLimit, string Delimeter, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Reading Excel file into DataTable with Rows..."); }

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
                    if (ConsoleOutput) { Console.WriteLine("Sheet " + SheetIndex.ToString() + ": " + SheetName + " - " + TableName); }

                    string RelationshipId = Sheets.ElementAt(SheetIndex).Id.Value;//.First().Id.Value;
                    WorksheetPart WorksheetPart = (WorksheetPart)SpreadSheetDocument.WorkbookPart.GetPartById(RelationshipId);
                    Worksheet WorkSheet = WorksheetPart.Worksheet;
                    SheetData SheetData = WorkSheet.GetFirstChild<SheetData>();
                    IEnumerable<Row> Rows = SheetData.Descendants<Row>();

                    //foreach (Cell Cell in Rows.ElementAt(0))
                    //{
                    //    DataTable.Columns.Add(Cell.CellValue.ToString());
                    //}
                    //DataTables.Add(DataTable);
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

                                TempRow[CellIndex] = FinalCellValue;
                                CellIndex++;
                            }

                            TempDataTable.Rows.Add(TempRow);

                            //when we get to Row BatchLimit, import that chunk into SQL Server
                            //also print to console to help with tracking
                            if (RowIndex != 0 && RowIndex % BatchLimit == 0)
                            {
                                LeftoverData = false;
                                InsertDataTableUsingSqlBulkCopy(ConnString, TableName, TempDataTable, RowIndex);

                                //reset TempDataTable
                                //if we don't do this files with about 18 columns/4 million rows cause the program to run out of memory
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
                        InsertDataTableUsingSqlBulkCopy(ConnString, TableName, TempDataTable, RowIndex);

                        TempDataTable = DataTables[SheetIndex];
                        TempDataTable.Rows.Clear();
                    }

                    SheetIndex++;
                }
            }
        }
        private void InsertDataTableUsingSqlBulkCopy(string ConnString, string TableName, DataTable TempDataTable, int RowIndex, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine(RowIndex.ToString() + " rows read"); }
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


                    if (ConsoleOutput) { Console.WriteLine("Inserting rows to table..."); }

                    SqlBulk.WriteToServer(TempDataTable);

                    if (ConsoleOutput) { Console.WriteLine("Inserted"); }
                }
            }
        }

        //------------------------------------------------------------------------------------
        // EXPORT
        //------------------------------------------------------------------------------------
        private void ExportButton_Click(object sender, EventArgs e)
        {
            Console.WriteLine("------------------------------");
            Console.WriteLine("Export Button Clicked");
            Console.WriteLine("------------------------------");

            /***************************************
             * CONSTANTS
            ***************************************/

            /***************************************
             * USER INPUT
            ***************************************/
            string Server = ServerComboBox.Text;
            string Database = DatabaseComboBox.Text;

            bool TableSearchMethodIsCommaList = CommaSeperatedListTableSearchRadioButton.Checked;
            bool TableSearchMethodIsRegexPattern = RegexPatternTableSearchRadioButton.Checked;
            bool TableSearchMethodIsTablePicker = TablePickerRadioButton.Checked;

            string TablesToExportCommaListText = TablesToExportCommaList.Text;
            List<string> TablesToExportCommaStrList = TablesToExportCommaListText.Split(',').ToList<string>();
            string TablesToExportRegexText = TablesToExportRegex.Text;
            List<string> TablesToExportListFromSqlList = TablesToExportListFromSql.SelectedItems.Cast<string>().ToList();

            string Extension = OutputTypeListBox.Text;

            string Delimeter = ExportDelimeterListBox.Text;
            string ActualDelimeter = "";
            if (Delimeter == "COMMA") { ActualDelimeter = ","; }
            else if (Delimeter == "PIPE") { ActualDelimeter = "|"; }
            else if (Delimeter == "TAB") { ActualDelimeter = "\t"; }
            else if (Delimeter == "FIXED WIDTH") { ActualDelimeter = "FIXED WIDTH"; }

            string ExportQualifier = ExportQualifierListBox.Text;
            if (ExportQualifier == "<NO QUALIFIER>")
            {
                ExportQualifier = "";
            }

            bool IncludeHeaders = IncludeHeadersCheckBox.Checked;
            string FixedWidthColumnLengthMethod = FixedWidthColumnLengthMethodListBox.Text;
            string ExportPath = ExportPathTextBox.Text; //must be a folder

            Console.WriteLine("Server.Databse: " + Server + "." + Database);

            /***************************************
             * GET TABLES TO EXPORT
            ***************************************/
            int FilesExported = 0;
            int TablesExported = 0;
            string FileName = "";

            List<string> TablesFromSqlDb = new List<string>();

            if (TableSearchMethodIsCommaList)
            {
                //User types out comma-seperated list, which is checked against tables that exist in SQL
                //Only the table names that match are returned
                TablesFromSqlDb = GetListofTablesFromSqlDb(Server, Database, TablesToExportCommaStrList);
            }
            else if (TableSearchMethodIsRegexPattern)
            {
                //User types out a regex pattern, which is checked against tables that exist in SQL
                //Only the table names that match are returned
                TablesFromSqlDb = GetListofTablesFromSqlDb(Server, Database, TablesToExportRegexText);
            }
            else if (TableSearchMethodIsTablePicker)
            {
                //User picks from a list of tables that exist in SQLs
                //We move forward with exactly the user input, because it definitely already exists in SQL
                TablesFromSqlDb = TablesToExportListFromSqlList;
            }
            TablesExported = TablesFromSqlDb.Count;

            /***************************************
             * EXPORT TABLE FROM SQL SERVER
            ***************************************/
            //for each table found
            for (int t = 0; t < TablesFromSqlDb.Count; t++)
            {
                string TableName = TablesFromSqlDb[t];
                Console.WriteLine("Table " + (t + 1).ToString() + ": " + TableName);
                ExportTableFromSqlToFile(Server, Database, TableName, ExportPath, Extension, ActualDelimeter, ExportQualifier, IncludeHeaders, FixedWidthColumnLengthMethod);
                FilesExported++;
                Console.WriteLine("");
            }

            Console.WriteLine(TablesExported.ToString() + " tables exported from " + Server + "." + Database);
            Console.WriteLine(FilesExported.ToString() + " files created " + ExportPath);
            Console.WriteLine("");
        }
        private List<string> GetListofTablesFromSqlDb(string Server, string Database, List<string> ListOfTablesToSearchFor, bool ConsoleOutput = true)
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
        private List<string> GetListofTablesFromSqlDb(string Server, string Database, string RegexSearchPattern = "", bool ConsoleOutput = true)
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
        private void ExportTableFromSqlToFile(string Server, string Database, string TableToExport, string ExportPath, string Extension, string Delimeter, string Qualifier, bool IncludeHeaders, string FixedWidthColumnLengthMethod, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Reading table from SQL Server"); }

            string ConnString = @"Server=" + Server + ";Database=" + Database + ";Trusted_Connection = True;";

            SqlDataReader DataReader = null;
            using (SqlConnection Conn = new SqlConnection(ConnString))
            {
                Conn.Open();
                string TableExportQuery = "SELECT * FROM [" + TableToExport + "]";
                SqlCommand Cmd = new SqlCommand(TableExportQuery, Conn);
                DataReader = Cmd.ExecuteReader();

                if (ConsoleOutput) { Console.WriteLine("Exporting table to file"); }

                //ExportPath
                string FileExportPath = ExportPath + "\\" + TableToExport + "." + Extension;

                if (Extension == "xlsx")
                {
                    //https://stackoverflow.com/questions/41605649/i-want-to-create-xlsx-excel-file-from-c-sharp
                    //need to get count

                    int RowCount = 0;
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

                    object[,] OutputRows = new object[RowCount, DataReader.FieldCount];
                    object[] Output = new object[DataReader.FieldCount];


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

                        if (ConsoleOutput && ExcelRowIndex != 0 && ExcelRowIndex % 10000 == 0)
                        {
                            Console.WriteLine("ROW " + ExcelRowIndex.ToString());
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
                else
                {
                    //build header
                    List<string> TableColumns = new List<string>();
                    for (int ColumnIndex = 0; ColumnIndex < DataReader.FieldCount; ColumnIndex++)
                    {
                        TableColumns.Add(DataReader.GetName(ColumnIndex));
                    }

                    StreamWriter sw = new StreamWriter(FileExportPath);
                    object[] Output = new object[DataReader.FieldCount];

                    List<string> FixedWidthColumnNames = new List<string>();
                    List<int> FixedWidthColumnLengths = new List<int>();

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
                            string HeaderRow = string.Join(Qualifier + Delimeter + Qualifier, TableColumns);
                            HeaderRow = Qualifier + HeaderRow + Qualifier;
                            sw.WriteLine(HeaderRow);
                        }
                    }

                    //write rows
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
                        sw.WriteLine(CurrentRow);
                    }
                    sw.Close();
                }

                if (ConsoleOutput) { Console.WriteLine("Exported table to " + FileExportPath); }
            }

            if (ConsoleOutput) { Console.WriteLine(""); }
        }

        //------------------------------------------------------------------------------------
        // FORM EVENT FUNCTIONS
        //------------------------------------------------------------------------------------
        private void LoadSqlTables_Click(object sender, EventArgs e)
        {
            string Server = ServerComboBox.Text;
            string Database = DatabaseComboBox.Text;

            if (Server != "" && Database != "")
            {
                List<string> Tables = GetListofTablesFromSqlDb(Server, Database, "", false);
                for (int i = 0; i < Tables.Count; i++)
                {
                    TablesToExportListFromSql.Items.Add(Tables[i]);
                }
            }
        }
        private void ExtensionListBoxImport_SelectedIndexChanged(object sender, EventArgs e)
        {
            string ValueSelected = ExtensionListBoxImport.SelectedItem.ToString();
            if (ValueSelected == "xls*")
            {
                if (ImportDelimeterListBox.Enabled)
                {
                    string ImportDelimeterCurrentlySelected = ImportDelimeterListBox.SelectedItem.ToString();
                    ImportDelimeterLastSelected.Text = ImportDelimeterCurrentlySelected;
                    ImportDelimeterListBox.Enabled = false;
                    ImportDelimeterListBox.ClearSelected();
                }
            }
            else
            {
                ImportDelimeterListBox.Enabled = true;
                ImportDelimeterListBox.SelectedItem = ImportDelimeterLastSelected.Text;
            }
        }
        private void ImportDelimeterListBox_SelectedIndexChanged(object sender, EventArgs e)
        {

            if (ImportDelimeterListBox.Enabled)
            {
                ImportDelimeterLastSelected.Text = ImportDelimeterListBox.SelectedItem.ToString();
                string ValueSelected = ImportDelimeterListBox.SelectedItem.ToString();
                if (ValueSelected == "FIXED WIDTH")
                {
                    FixedWidthColumnFilePathLabel.Font = new System.Drawing.Font(FixedWidthColumnFilePathLabel.Font, FontStyle.Bold);
                }
                else
                {
                    FixedWidthColumnFilePathLabel.Font = new System.Drawing.Font(FixedWidthColumnFilePathLabel.Font, FontStyle.Regular);
                }
            }
            else
            {
                FixedWidthColumnFilePathLabel.Font = new System.Drawing.Font(FixedWidthColumnFilePathLabel.Font, FontStyle.Regular);
            }
        }
        private void OutputTypeListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            string ValueSelected = OutputTypeListBox.SelectedItem.ToString();
            if (ValueSelected == "xlsx")
            {
                if (ExportDelimeterListBox.Enabled)
                {
                    string ExportDelimeterCurrentlySelected = ExportDelimeterListBox.SelectedItem.ToString();
                    ExportDelimeterListBox.Text = ExportDelimeterCurrentlySelected;
                    ExportDelimeterListBox.Enabled = false;
                    ExportDelimeterListBox.ClearSelected();
                }
                if (ExportQualifierListBox.Enabled)
                {
                    string ExportQualifierCurrentlySelected = ExportQualifierListBox.SelectedItem.ToString();
                    ExportQualifierLastSelected.Text = ExportQualifierCurrentlySelected;
                    ExportQualifierListBox.Enabled = false;
                    ExportQualifierListBox.ClearSelected();
                }
            }
            else
            {
                ExportDelimeterListBox.Enabled = true;
                ExportDelimeterListBox.Text = ExportDelimeterLastSelected.Text;
                //ExportDelimeterListBox.SelectedItem = ExportDelimeterLastSelected.Text;
            }
        }
        private void ExportDelimeterListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ExportDelimeterListBox.Enabled)
            {
                string ValueSelected = ExportDelimeterListBox.SelectedItem.ToString();
                ExportDelimeterLastSelected.Text = ValueSelected;
                if (ValueSelected == "FIXED WIDTH")
                {
                    ExportQualifierListBox.SetSelected(2, true);
                    ExportQualifierListBox.Enabled = false;
                    //FixedWidthColumnFilePathLabel.Font = new System.Drawing.Font(FixedWidthColumnFilePathLabel.Font, FontStyle.Bold);
                }
                else
                {
                    ExportDelimeterListBox.Text = ExportDelimeterLastSelected.Text;
                    ExportQualifierListBox.Enabled = true;
                    ExportQualifierListBox.SelectedItem = ExportQualifierLastSelected.Text;
                    //FixedWidthColumnFilePathLabel.Font = new System.Drawing.Font(FixedWidthColumnFilePathLabel.Font, FontStyle.Regular);
                }
            }
            else
            {
                //FixedWidthColumnFilePathLabel.Font = new System.Drawing.Font(FixedWidthColumnFilePathLabel.Font, FontStyle.Regular);
            }
        }
        private void ExportQualifierListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (ExportQualifierListBox.Enabled)
            {
                if (ExportDelimeterListBox.SelectedItem.ToString() != "FIXED WIDTH")
                {
                    ExportQualifierLastSelected.Text = ExportQualifierListBox.SelectedItem.ToString();
                }
            }
        }
    }
}

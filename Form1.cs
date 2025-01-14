/* TABLE OF CONTENTS
 * 
 * METHODS
 * main:
 * private void ImportButton_Click(object sender, EventArgs e)
 * private void ExportButton_Click(object sender, EventArgs e)
 * 
 * form event helpers...
 * 
 * form events...
 *  
 * NOTES
 * Leading 0s are preserved
 * Leading/Trailing spaces are removed - looks like each field is automatically trimmed (not sure why)
*/

using System;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.IO; //for Directory
using System.Linq;
using System.Windows.Forms;
using System.Security;

namespace SQL_SERVER_IMPORT_EXPORT
{
    public partial class Form1 : Form
    {
        //======================================================================================================================
        // METHODS
        //======================================================================================================================

        public Form1()
        {
            InitializeComponent();

            FormData FormData = new FormData();
            //default values
            ExtensionListBoxImport.SetSelected(0, true); //csv
            ImportDelimeterListBox.SetSelected(0, true); //COMMA

            OutputTypeListBox.SetSelected(0, true); //csv
            ExportDelimeterListBox.SetSelected(0, true); //COMMA
            ExportQualifierListBox.SetSelected(0, true); //"

            TablePickerRadioButton.Select(); //Select tables to export from SQL table picker
            NoSplitRadioButton.Select(); //Select option to not split files on export

            SplitAmounNumericUpDown.Maximum = int.MaxValue - 1; //setting maximum value for split amount

            SplitAmounNumericUpDown.Enabled = false;
            IncludeHeaderInSplitFilesCheckBox.Enabled = false;

            //cant get current path of shortcut that's used to run this!
            string CurrentPath = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
            //set export/import paths to the path where the application is located
            //this is useless unless we can find the location of the shortcut that started the application
            //ImportPathTextBox.Text = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location); ;
            //ExportPathTextBox.Text = System.IO.Directory.GetCurrentDirectory();//Environment.CurrentDirectory;
            //SelectTextBox.Text = System.AppDomain.CurrentDomain.BaseDirectory;// System.IO.Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            //GroupByTextBox.Text = Application.UserAppDataPath;//Application.StartupPath;

            FormData.Environs.ToList().ForEach(n => EnvironComboBox.Items.Add(n));
            EnvironComboBox.Text = "SQL SERVER";
        }

        //------------------------------------------------------------------------------------
        // MAIN
        //------------------------------------------------------------------------------------
        private void ImportButton_Click(object sender, EventArgs e)
        {
            Console.WriteLine("------------------------------");
            Console.WriteLine("Import Button Clicked");
            Console.WriteLine("------------------------------");

            Import Import = new Import();
            Snowflake Snowflake = new Snowflake();
            ConnectionInfo ConnectionInfo = new ConnectionInfo(); //gets connection info straight from the form

            /***************************************
             * CONSTANTS
            ***************************************/
            //averages about 30 seconds on SQL04 for meta format (18 cols) for every million records
            int BatchLimit = 100000; //100k

            /***************************************
             * USER INPUT
            ***************************************/
            if (ConnectionInfo.Environ == "")
            {
                Console.WriteLine("Environ. must not be blank!");
                return;
            }
            else if (ConnectionInfo.Environ == "SQL SERVER" && (ConnectionInfo.Server == "" || ConnectionInfo.Database == ""))
            {
                Console.WriteLine("SQL-SERVER Server and Database must not be blank!");
                return;
            }
            else if (ConnectionInfo.Environ == "SNOWFLAKE")
            {
                if (ConnectionInfo.Database == "")
                {
                    Console.WriteLine("SNOWFLAKE Database must not be blank!");
                    return;
                }
                else
                {
                    Snowflake.ConnectToDb(ConnectionInfo);
                }
            }

            string Extension = ExtensionListBoxImport.Text;

            string Delimeter = ImportDelimeterListBox.Text;
            string ActualDelimeter = "";
            if (Delimeter == "COMMA") { ActualDelimeter = ","; }
            else if (Delimeter == "PIPE") { ActualDelimeter = "|"; }
            else if (Delimeter == "TAB") { ActualDelimeter = "\t"; }
            else if (Delimeter == "FIXED WIDTH") { ActualDelimeter = "FIXED WIDTH"; }

            //int HeaderRowNumber = (int)HeaderRowNumberUpDowm.Value;
            //int RowsToSkipAfterHeader = (int)RowsToSkipAfterHeaderUpDown.Value;
            //header row must be on the first row

            bool IsDoubleQuoted = DoubleQuoted.Checked;

            int ExcelHeaderRow = Convert.ToInt32(ExcelHeaderRowUpDown.Value) - 1; //1-indexed -1 so it's 0-indexed

            bool ConvertExcelToCsv = ConvertExcelToCsvCheckBox.Checked;
            bool DeleteConvertedExcelCsvs = DeleteConvertedExcelCsvsCheckBox.Checked;

            string ImportToSingleTableName = ImportToSingleTableTextBox.Text.Trim();
            bool ImportToSingleTable = false;
            if (ImportToSingleTableName != "") { ImportToSingleTable = true; }

            bool ImportToExistingTable = InsertToExistingTableCheckBox.Checked;

            string FixedWidthColumnFilePath = FixedWidthColumnFilePathTextBox.Text;
            string ColumnTypeFilePath = ColumnTypeFilePathTextBox.Text;
            bool ColumnTypeVarcharDefault = ColumnTypeVarcharDefaultRadioButton.Checked;
            bool ColumnTypeUseFile = ColumnTypeUseFileRadioButton.Checked;
            string ColumnTypeMethod = "DEFAULT VARCHAR";
            if (ColumnTypeUseFile)
            {
                ColumnTypeMethod = "FILE PATH";
            }

            //bool FasterImport = FasterImportCheckBox.Checked;
            string EncodingChoiceInput = ImportEncodingListBox.Text;
            int LimitFieldValue = Convert.ToInt32(LimitFieldLengthUpDown.Value);
            bool StripNonPrintableChars = StripNonPrintableCharsCheckBox.Checked;

            Encoding Encoding = new Encoding();
            System.Text.Encoding EncodingChoice = Encoding.Get(EncodingChoiceInput);
            string EncodingChoiceSnowflake = Encoding.GetSnowflake(EncodingChoiceInput);

            string ImportRegexFilter = ImportRegexFilterTextBox.Text;

            string ImportPath = ImportPathTextBox.Text;

            if (ImportPath == "")
            {
                Console.WriteLine("Import Path must be specified!\n");
                return;
            }

            ConnectionInfo.print();
            //Console.WriteLine("Server.Databse: " + ConnectionInfo.Server + "." + ConnectionInfo.Database);

            if (ConnectionInfo.Environ == "SQL SERVER") { Console.WriteLine("BatchLimit: " + BatchLimit.ToString()); }

            /***************************************
             * GET FILES TO IMPORT
            ***************************************/
            List<string> FilesToImport = Import.GetFilesToImport(ImportPath, Extension, ConvertExcelToCsv, ImportRegexFilter);

            int FilesImported = 0;
            int TablesCreated = 0;
            string TableName = "";
            List<string> TableNames = new List<string>();
            DataTable BaseDtTable = new DataTable();
            List<DataTable> BaseDtTables = new List<DataTable>();

            /***************************************
             * FOR EACH FILE TO IMPORT
            ***************************************/
            //CHECK IF EXT IS XLS AND WE ARE CONVERTING TO CSV
            //IF YES THEN CHANGE EXTENSION TO CSV, CHANGE DELIM TO COMMA, CHANGE QUOTED TO ?
            //OUTPUT USER MESSAGE WARMING OF THE CHANGES
            bool DeleteConvertedExcelCsvsActual = false;
            bool DisplayNotDeletedMsg = false;

            if (Extension == "xls*" && ConvertExcelToCsv == true)
            {
                Console.WriteLine("Excel files converted to csv files. The script will now proceed as if importing csv files like usual");
                Delimeter = "COMMA";
                ActualDelimeter = ",";
                Extension = "csv";
                //IsDoubleQuoted

                if(DeleteConvertedExcelCsvs == true)
                {
                    DeleteConvertedExcelCsvsActual = true;
                }
                else
                {
                    DisplayNotDeletedMsg = true;
                }
            }

            for (int f = 0; f < FilesToImport.Count; f++)
            {
                string FilePath = FilesToImport[f];

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
                        BaseDtTables = Import.ReadExcelFileIntoDataTablesWithColumns(FilePath, ref TableNames, EncodingChoice, StripNonPrintableChars, LimitFieldValue, ExcelHeaderRow);
                        //will also add to TableNames list
                    }
                    else
                    {
                        BaseDtTable = Import.ReadFileIntoDataTableWithColumns(FilePath, ActualDelimeter, IsDoubleQuoted, EncodingChoice, StripNonPrintableChars, LimitFieldValue, FixedWidthColumnFilePath, ColumnTypeFilePath);
                        BaseDtTables.Add(BaseDtTable);
                        TableNames.Add(TableName);
                    }

                    /***************************************
                     * CREATE TABLE IN SQL
                    ***************************************/
                    if (ImportToExistingTable == false)
                    {
                        if (ConnectionInfo.Environ == "SQL SERVER")
                        {
                            Import.CreateTablesInSqlServerVarchar(TableNames, ConnectionInfo, BaseDtTables, ActualDelimeter, EncodingChoice, StripNonPrintableChars, LimitFieldValue, ColumnTypeMethod, ColumnTypeFilePath);
                            TablesCreated += BaseDtTables.Count;
                        }
                        else if(ConnectionInfo.Environ == "SNOWFLAKE")
                        {
                            //Helpers.CreateTablesInSnowflakeVarchar(TableNames, ConnectionInfo, Snowflake, BaseDtTables, ActualDelimeter, ColumnTypeMethod, ColumnTypeFilePath);
                            //THIS IS DONE IN THIS: ReadFileIntoDataTableWithRowsAndInsertIntoSqlServerTable()
                            TablesCreated += BaseDtTables.Count;
                        }
                    }
                }

                /***************************************
                 * READ FILE ROWS AND INSERT INTO SQL TABLE
                ***************************************/
                if (Extension == "xls*")
                {
                    if(ConnectionInfo.Environ == "SQL SERVER")
                    {
                        Import.ReadExcelFilePerSheetIntoDataTablesWithRowsAndInsertIntoSqlServerTables(FilePath, TableNames, ConnectionInfo, BaseDtTables, BatchLimit, ActualDelimeter, EncodingChoice, StripNonPrintableChars, LimitFieldValue, ExcelHeaderRow);
                    }
                    //else if (ConnectionInfo.Environ == "SNOWFLAKE") { }
                }
                else
                {
                    if (ConnectionInfo.Environ == "SQL SERVER")
                    {
                        Import.ReadFileIntoDataTableWithRowsAndInsertIntoSqlServerTable(FilePath, TableName, ConnectionInfo, BaseDtTable, BatchLimit, ActualDelimeter, IsDoubleQuoted, EncodingChoice, StripNonPrintableChars, LimitFieldValue);
                    }
                    else if (ConnectionInfo.Environ == "SNOWFLAKE")
                    {
                        Import.ReadFileIntoDataTableWithRowsAndInsertIntoSnowflakeTable(FilePath, TableName, ConnectionInfo, Snowflake, BaseDtTable, BatchLimit, ActualDelimeter, IsDoubleQuoted, EncodingChoice, EncodingChoiceSnowflake, StripNonPrintableChars, LimitFieldValue, ColumnTypeMethod, ColumnTypeFilePath, ImportToExistingTable);
                    }
                }
                FilesImported += BaseDtTables.Count;

                BaseDtTables.Clear();
                TableNames.Clear();

                Console.WriteLine("");
            }

            if (ConnectionInfo.Environ == "SNOWFLAKE")
            {
                Snowflake.Close();
            }
            //CHECK IF EXT IS XLS AND WE ARE CONVERTING TO CSV AND IF USER WANTS TO DELETE THE CSVS
            //IF YES THEN DELETE THE CSVS, OUTPUT USER MSG
            if (DeleteConvertedExcelCsvsActual == true && ConvertExcelToCsv == true)
            //(Extension == "xls*" && ConvertExcelToCsv == true && DeleteConvertedExcelCsvs == true) //wont work because extensions is changed earlier, near where we define DeleteConvertedExcelCsvsActual
            {
                Console.WriteLine("Deleting the csv files convertred form the excel files...");
                foreach (string FileToDelete in FilesToImport)
                {
                    if (File.Exists(FileToDelete))
                    {
                        File.Delete(FileToDelete);
                    }
                }
                Console.WriteLine("Deleted the csv files.");
                Console.WriteLine("");
            }
            else if(DisplayNotDeletedMsg)
            {
                Console.WriteLine("Csv files convertred form the excel files were NOT deleted.");
                Console.WriteLine("");
            }

            if (ImportToSingleTable == false)
            {
                Console.WriteLine(FilesImported.ToString() + " files imported to " + ConnectionInfo.Environ + "-" +  ConnectionInfo.Server + "-" + ConnectionInfo.Database);
            }
            else if (ImportToSingleTable == true)
            {
                Console.WriteLine(FilesImported.ToString() + " files imported to " + ConnectionInfo.Environ + "-" + ConnectionInfo.Server + "-" + ConnectionInfo.Database + "-" + TableName);
            }

            Console.WriteLine("");
        }
        private void ExportButton_Click(object sender, EventArgs e)
        {
            Console.WriteLine("------------------------------");
            Console.WriteLine("Export Button Clicked");
            Console.WriteLine("------------------------------");

            Export Export = new Export();
            Snowflake Snowflake = new Snowflake();
            ConnectionInfo ConnectionInfo = new ConnectionInfo(); //gets connection info straight from the form

            /***************************************
             * CONSTANTS
            ***************************************/

            /***************************************
             * USER INPUT
            ***************************************/
            if (ConnectionInfo.Environ == "")
            {
                Console.WriteLine("Environ. must not be blank!");
                return;
            }
            else if (ConnectionInfo.Environ == "SQL SERVER " && (ConnectionInfo.Server == "" || ConnectionInfo.Database == ""))
            {
                Console.WriteLine("SQL-SERVER Server and Database must not be blank!");
                return;
            }
            else if (ConnectionInfo.Environ == "SNOWFLAKE")
            {
                if (ConnectionInfo.Database == "")
                {
                    Console.WriteLine("SNOWFLAKE Database must not be blank!");
                    return;
                }
                else
                {
                    Snowflake.ConnectToDb(ConnectionInfo);
                }
            }

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

            bool QualifyAll = QualifyAllRadioButton.Checked;
            bool QualifyIfDelimeter = QualifyIfDelimeterRadioButton.Checked;
            bool QualifyEveryField = false;
            if(QualifyAll)
            {
                QualifyEveryField = true;
            }

            bool RemoveQualInVal = RemoveQualInValCheckBox.Checked;

            bool IncludeHeaders = IncludeHeadersCheckBox.Checked;
            //string FixedWidthColumnLengthMethod = FixedWidthColumnLengthMethodListBox.Text;

            int SizeLimit = Convert.ToInt32(SplitAmounNumericUpDown.Value);
            //decimal SizeLimit = SplitAmounNumericUpDown.Value;
            string SizeLimitType = "";
            bool IncludeHeaderInSplitFiles = IncludeHeaderInSplitFilesCheckBox.Checked;
            bool NoSplitRadioButtonChecked = NoSplitRadioButton.Checked;
            bool RowSplitRadioButtonChecked = RowSplitRadioButton.Checked;
            bool SizeSplitRadioButtonChecked = SizeSplitRadioButton.Checked;
            bool SizeSplit1024RadioButtonChecked = SizeSplit1024RadioButton.Checked;
            if (SizeSplitRadioButtonChecked)
            {
                SizeLimitType = "SIZE";
            }
            else if(SizeSplit1024RadioButtonChecked)
            {
                SizeLimitType = "SIZE1024";
            }
            else if(RowSplitRadioButtonChecked)
            {
                SizeLimitType = "ROW";
                //SizeLimit = Math.Floor(SizeLimit); //round down
            }
            else
            {
                SizeLimitType = "";
                SizeLimit = 0;
            }

            string SelectText = SelectTextBox.Text;
            string FromText = FromTextBox.Text;
            string GroupByText = GroupByTextBox.Text;
            string OrderByText = OrderByTextBox.Text;
            string WhereText = WhereTextBox.Text;


            string EncodingChoiceExport = ExportEncodingListBox.Text;

            Encoding Encoding = new Encoding();
            System.Text.Encoding EncodingChoice = Encoding.Get(EncodingChoiceExport);

            string ExportPath = ExportPathTextBox.Text; //must be a folder

            if(ExportPath == "")
            {
                Console.WriteLine("Export Path must be specified!\n");
                return;
            }
            if (Directory.Exists(ExportPath) == false)
            {
                Console.WriteLine("Export Path does not exist!\nCreating Export Path...\n\n");
                Directory.CreateDirectory(ExportPath);
            }

            ConnectionInfo.print();

            /***************************************
             * GET TABLES TO EXPORT
            ***************************************/
            int FilesExported = 0;
            int TablesExported = 0;
            //string FileName = "";

            List<string> TablesFromSqlDb = new List<string>();

            if (TableSearchMethodIsCommaList)
            {
                //User types out comma-seperated list, which is checked against tables that exist in SQL
                //Only the table names that match are returned
                if (ConnectionInfo.Environ == "SQL SERVER")
                {
                    TablesFromSqlDb = Export.GetListofTablesFromSqlServerDb(ConnectionInfo, TablesToExportCommaStrList);
                }
                else if (ConnectionInfo.Environ == "SNOWFLAKE")
                {
                    TablesFromSqlDb = Export.GetListofTablesFromSnowflakeDb(Snowflake, TablesToExportCommaStrList);
                }
            }
            else if (TableSearchMethodIsRegexPattern)
            {
                //User types out a regex pattern, which is checked against tables that exist in SQL
                //Only the table names that match are returned
                if (ConnectionInfo.Environ == "SQL SERVER")
                {
                    TablesFromSqlDb = Export.GetListofTablesFromSqlServerDb(ConnectionInfo, TablesToExportRegexText);
                }
                else if (ConnectionInfo.Environ == "SNOWFLAKE")
                {
                    TablesFromSqlDb = Export.GetListofTablesFromSnowflakeDb(Snowflake, TablesToExportRegexText);
                }
            }
            else if (TableSearchMethodIsTablePicker)
            {
                //User picks from a list of tables that exist in SQLs
                //We move forward with exactly the user input, because it definitely already exists in SQL
                TablesFromSqlDb = TablesToExportListFromSqlList;
            }
            TablesExported = TablesFromSqlDb.Count;
            Console.WriteLine("");

            /***************************************
             * EXPORT TABLE FROM SQL SERVER
            ***************************************/
            //for each table found
            for (int t = 0; t < TablesFromSqlDb.Count; t++)
            {
                string TableName = TablesFromSqlDb[t];
                Console.WriteLine("Table " + (t + 1).ToString() + ": " + TableName);

                //ideas:

                //add user input - combined file yes/no
                //if combnine, then disable split
                //if split, then disable combine
                //if excel and combine, then enable combine to  single sheet or 1-file-per sheet

                //add text field for combine filename

                //if combine and delimeted - do we care about headers?
                //if combine and fixed - do we allow this?
                //if combine and excel and 1 sheet - do we care about headers? do we care about 1MM row limit?
                //what to name tab? - Sheet?
                //if combine and excel and multi sheet - how to name tab so it's always sub 32 if i <= 9 (left(table name, 28) + "(" + (i) + ")" //// if i >= 10 (left(table name, 27) + "(" + (i) + ")"


                int FilesCreated = 0;
                if (ConnectionInfo.Environ == "SQL SERVER")
                {
                    FilesCreated = Export.ExportTableFromSqlServerToFile(ConnectionInfo, TableName, ExportPath, Extension, ActualDelimeter, EncodingChoice, ExportQualifier, QualifyEveryField, RemoveQualInVal, IncludeHeaders, "MAX LEN", SizeLimit, SizeLimitType, IncludeHeaderInSplitFiles, SelectText, FromText, WhereText, GroupByText, OrderByText);
                }
                else if (ConnectionInfo.Environ == "SNOWFLAKE")
                {
                    FilesCreated = Export.ExportTableFromSnowflakeToFile(Snowflake, TableName, ExportPath, Extension, ActualDelimeter, EncodingChoice, ExportQualifier, QualifyEveryField, RemoveQualInVal, IncludeHeaders, "MAX LEN", SizeLimit, SizeLimitType, IncludeHeaderInSplitFiles, SelectText, FromText, WhereText, GroupByText, OrderByText);
                }

                FilesExported  += FilesCreated;
                Console.WriteLine("");
            }
            if (ConnectionInfo.Environ == "SNOWFLAKE")
            {
                Snowflake.Close();
            }

            Console.WriteLine(TablesExported.ToString() + " tables exported from " + ConnectionInfo.Environ + "-" + ConnectionInfo.Server + "-" + ConnectionInfo.Database);
            Console.WriteLine(FilesExported.ToString() + " files created " + ExportPath);
            Console.WriteLine("");
        }

        //------------------------------------------------------------------------------------
        // FORM EVENTS HELPERS
        //------------------------------------------------------------------------------------
        public void UpdateImportDelimeterListBox()
        {

            if (ImportDelimeterListBox.SelectedItem != null)
            {
                string ValueSelected = ImportDelimeterListBox.SelectedItem.ToString();
                if (ValueSelected == "FIXED WIDTH")
                {
                    FixedWidthColumnFilePathLabel.Font = new System.Drawing.Font(FixedWidthColumnFilePathLabel.Font, FontStyle.Bold);
                    LimitFieldLengthUpDown.Enabled = false;
                    DoubleQuoted.Enabled = false;
                }
                else
                {
                    FixedWidthColumnFilePathLabel.Font = new System.Drawing.Font(FixedWidthColumnFilePathLabel.Font, FontStyle.Regular);
                    LimitFieldLengthUpDown.Enabled = true;
                    DoubleQuoted.Enabled = true;
                }
            }
        }
        public void UpdateOutputTypeListBox(string Value = "")
        {
            string ValueSelected = OutputTypeListBox.SelectedItem.ToString();
            if (ValueSelected == "xlsx")
            {
                ExportQualifierListBox.ClearSelected();
                ExportDelimeterListBox.ClearSelected();
                ExportDelimeterListBox.Enabled = false;
                ExportQualifierListBox.SetSelected(2, true);
                ExportQualifierListBox.Enabled = false;
                QualifyAllRadioButton.Enabled = false;
                QualifyIfDelimeterRadioButton.Enabled = false;
                RemoveQualInValCheckBox.Enabled = false;

                NoSplitRadioButton.Checked = true;
                NoSplitRadioButton.Enabled = false;
                RowSplitRadioButton.Enabled = false;
                SizeSplitRadioButton.Enabled = false;
                SizeSplit1024RadioButton.Enabled = false;
                SplitAmounNumericUpDown.Enabled = false;
                IncludeHeaderInSplitFilesCheckBox.Enabled = false;
                ExportEncodingListBox.Enabled = false;
            }
            else if (ValueSelected == "csv")
            {
                ExportDelimeterListBox.SelectedItem = "COMMA";
                ExportDelimeterListBox.Enabled = true;
                ExportQualifierListBox.SelectedItem = "\"";
                ExportQualifierListBox.Enabled = true;
                QualifyAllRadioButton.Enabled = true;
                QualifyIfDelimeterRadioButton.Enabled = true;
                RemoveQualInValCheckBox.Enabled = true;

                NoSplitRadioButton.Enabled = true;
                RowSplitRadioButton.Enabled = true;
                SizeSplitRadioButton.Enabled = true;
                SizeSplit1024RadioButton.Enabled = true;
                SplitAmounNumericUpDown.Enabled = true;

                if (ExportDelimeterListBox.SelectedItem.ToString() != "FIXED WIDTH")
                {
                    IncludeHeaderInSplitFilesCheckBox.Enabled = true;
                }
                else if (ExportDelimeterListBox.SelectedItem.ToString() != "FIXED WIDTH")
                {
                    IncludeHeaderInSplitFilesCheckBox.Enabled = false;
                }
                ExportEncodingListBox.Enabled = true;
            }
            else
            {
                ExportDelimeterListBox.Enabled = true;
                NoSplitRadioButton.Enabled = true;
                RowSplitRadioButton.Enabled = true;
                SizeSplitRadioButton.Enabled = true;
                SizeSplit1024RadioButton.Enabled = true;
                SplitAmounNumericUpDown.Enabled = true;
                if (ExportDelimeterListBox.SelectedItem == null)
                {
                    ExportDelimeterListBox.SelectedItem = "COMMA";
                }
                if (ExportQualifierListBox.SelectedItem == null)
                {
                    ExportQualifierListBox.SelectedItem = "\"";
                }
                if (ExportDelimeterListBox.SelectedItem.ToString() != "FIXED WIDTH")
                {
                    IncludeHeaderInSplitFilesCheckBox.Enabled = true;
                }
                else if (ExportDelimeterListBox.SelectedItem.ToString() == "FIXED WIDTH")
                {
                    IncludeHeaderInSplitFilesCheckBox.Enabled = false;
                }
                ExportEncodingListBox.Enabled = true;
            }
        }
        public void UpdateExportDelimeterListBox()
        {
            if (ExportDelimeterListBox.SelectedItem != null)
            {
                string ValueSelected = ExportDelimeterListBox.SelectedItem.ToString();

                if (ValueSelected == "FIXED WIDTH")
                {
                    ExportQualifierListBox.SetSelected(2, true);
                    ExportQualifierListBox.Enabled = false;
                    QualifyAllRadioButton.Enabled = false;
                    QualifyIfDelimeterRadioButton.Enabled = false;
                    QualifyIfDelimeterRadioButton.Enabled = false;
                    RemoveQualInValCheckBox.Enabled = false;
                    IncludeHeaderInSplitFilesCheckBox.Enabled = false;
                }
                else if (ValueSelected == "csv")
                {
                    ExportDelimeterListBox.Enabled = true;
                    ExportDelimeterListBox.SelectedItem = "COMMA";
                    ExportDelimeterListBox.Enabled = true;
                    //if ((ExportDelimeterListBox.SelectedItem.ToString() != null && ExportDelimeterListBox.SelectedItem.ToString() != "xslx")
                    //    && (ExportQualifierListBox.SelectedItem.ToString() != null && ExportQualifierListBox.SelectedItem.ToString() != "FIXED WIDTH"))
                    //{
                    IncludeHeaderInSplitFilesCheckBox.Enabled = true;
                    // }
                }
                else
                {
                    //if ((ExportDelimeterListBox.SelectedItem.ToString() != null && ExportDelimeterListBox.SelectedItem.ToString() != "xslx")
                    //    && (ExportQualifierListBox.SelectedItem.ToString() != null && ExportQualifierListBox.SelectedItem.ToString() != "FIXED WIDTH"))
                    //{
                    IncludeHeaderInSplitFilesCheckBox.Enabled = true;
                    //}
                    ExportQualifierListBox.Enabled = true;
                    QualifyAllRadioButton.Enabled = true;
                    QualifyIfDelimeterRadioButton.Enabled = true;
                    RemoveQualInValCheckBox.Enabled = true;

                    if (ExportDelimeterListBox.SelectedItem == null)
                    {
                        ExportDelimeterListBox.SelectedItem = "COMMA";
                    }
                }
            }
        }
        public void UpdateExportQualifierListBox()
        {
            if (ExportQualifierListBox.SelectedItem == null && OutputTypeListBox.SelectedItem.ToString() != "xlsx")
            {
                ExportQualifierListBox.SelectedItem = "\"";
            }
        }
        public void UpdateExtensionListBoxImport(string Value = "")
        {

            string ValueSelected = "";
            if (Value == "")
            {
                if (ExtensionListBoxImport.SelectedItem != null)
                {
                    ValueSelected = ExtensionListBoxImport.SelectedItem.ToString();
                }
            }
            else
            {
                ValueSelected = Value;
                if (ValueSelected == "xls" || ValueSelected == "xlsm" || ValueSelected == "xlsx")
                {
                    ExtensionListBoxImport.SelectedItem = "xls*";

                    if (EnvironComboBox.SelectedItem.ToString() == "SQL SERVER") { ConvertExcelToCsvCheckBox.Enabled = true; }
                    if (EnvironComboBox.SelectedItem.ToString() == "SNOWFLAKE") { ConvertExcelToCsvCheckBox.Enabled = false; }

                    DeleteConvertedExcelCsvsCheckBox.Enabled = true;
                    ExcelHeaderRowUpDown.Enabled = true;
                    if (ConvertExcelToCsvCheckBox.Checked == true)
                    {
                        ImportEncodingListBox.Enabled = true;
                    }
                    else
                    {
                         ImportEncodingListBox.Enabled = false; 
                    }
                }
                else
                {
                    ConvertExcelToCsvCheckBox.Enabled = false;
                    DeleteConvertedExcelCsvsCheckBox.Enabled = false;
                    ExcelHeaderRowUpDown.Enabled = false;
                    ExtensionListBoxImport.SelectedItem = Value;
                }
            }

            if (ValueSelected == "xls*" || ValueSelected == "xls" || ValueSelected == "xlsm" || ValueSelected == "xlsx")
            {
                if (EnvironComboBox.SelectedItem.ToString() == "SQL SERVER") { ConvertExcelToCsvCheckBox.Enabled = true; }
                if (EnvironComboBox.SelectedItem.ToString() == "SNOWFLAKE") { ConvertExcelToCsvCheckBox.Enabled = false; }

                DeleteConvertedExcelCsvsCheckBox.Enabled = true;
                ExcelHeaderRowUpDown.Enabled = true;

                if (ImportDelimeterListBox.Enabled)
                {
                    string ImportDelimeterCurrentlySelected = ImportDelimeterListBox.SelectedItem.ToString();
                    ImportDelimeterListBox.ClearSelected();
                    ImportDelimeterListBox.Enabled = false;
                }
                else
                {
                    if (ImportDelimeterListBox.SelectedItem != null)
                    {
                        string ImportDelimeterCurrentlySelected = ImportDelimeterListBox.SelectedItem.ToString();
                    }
                    ImportDelimeterListBox.ClearSelected();
                }
                FixedWidthColumnFilePathLabel.Font = new System.Drawing.Font(FixedWidthColumnFilePathLabel.Font, FontStyle.Regular);
                DoubleQuoted.Enabled = false;

                if (ConvertExcelToCsvCheckBox.Checked == true)
                {
                    ImportEncodingListBox.Enabled = true;
                }
                else
                {
                    ImportEncodingListBox.Enabled = false;
                }
            }
        
            else if (ValueSelected == "csv")
            {
                ConvertExcelToCsvCheckBox.Enabled = false;
                DeleteConvertedExcelCsvsCheckBox.Enabled = false;
                ExcelHeaderRowUpDown.Enabled = false;

                ImportDelimeterListBox.Enabled = true;
                ImportDelimeterListBox.SelectedItem = "COMMA";
                //ImportDelimeterListBox.Enabled = false;

                if (EnvironComboBox.Text == "SQL SERVER")
                {
                    DoubleQuoted.Enabled = true;
                }
                ImportEncodingListBox.Enabled = true;
            }
            else
            {
                ConvertExcelToCsvCheckBox.Enabled = false;
                DeleteConvertedExcelCsvsCheckBox.Enabled = false;
                ExcelHeaderRowUpDown.Enabled = false;

                ImportDelimeterListBox.Enabled = true;
                if (ImportDelimeterListBox.SelectedItem == null)
                {
                    ImportDelimeterListBox.SelectedItem = "COMMA";
                }
                if (EnvironComboBox.Text == "SQL SERVER")
                {
                    DoubleQuoted.Enabled = true;
                }
                ImportEncodingListBox.Enabled = true;
            }
        }

        //------------------------------------------------------------------------------------
        // FORM EVENTS
        //------------------------------------------------------------------------------------
        private void Form1_Load(object sender, EventArgs e)
        {

        }
        private void LoadSqlTables_Click(object sender, EventArgs e)
        {
            Console.WriteLine("Loading tables...");
            ConnectionInfo ConnectionInfo = new ConnectionInfo();
            Export Export = new Export();

            if (ConnectionInfo.Environ == "SQL SERVER" && ConnectionInfo.Server != "" && ConnectionInfo.Database != "")
            {
                TablesToExportListFromSql.Items.Clear();
                List<string> Tables = Export.GetListofTablesFromSqlServerDb(ConnectionInfo, "", false);
                for (int i = 0; i < Tables.Count; i++)
                {
                    TablesToExportListFromSql.Items.Add(Tables[i]);
                }
            }
            else if (ConnectionInfo.Environ == "SNOWFLAKE" && ConnectionInfo.Database != "")
            {
                Snowflake Snowflake = new Snowflake();
                Snowflake.ConnectToDb(ConnectionInfo);

                TablesToExportListFromSql.Items.Clear();
                List<string> Tables = Export.GetListofTablesFromSnowflakeDb(Snowflake, "", false);
                for (int i = 0; i < Tables.Count; i++)
                {
                    TablesToExportListFromSql.Items.Add(Tables[i]);
                }
                Snowflake.Close();
            }
            
            Console.WriteLine(TablesToExportListFromSql.Items.Count.ToString() + " tables loaded.");
        }
        private void ListColumnsForSelectedTablesButton_Click(object sender, EventArgs e)
        {
            Console.WriteLine("COLUMNS FOR SELECTED TABLES");
            Export Export = new Export();
            ConnectionInfo ConnectionInfo = new ConnectionInfo();

            List<string> UserSelectedTables = Export.GetListOfUserSelectedTables();

            //for every table selected
            for (int t = 0; t < UserSelectedTables.Count; t++)
            {
                string Table = UserSelectedTables[t];
                //get columns
                List<string> ColumnNames = Export.GetListOfColumnsForTable(ConnectionInfo, Table);

                Console.WriteLine(ColumnNames.Count + " columns");

                //list them in output window
                for (int c = 0; c < ColumnNames.Count; c++)
                {
                    Console.WriteLine("[" + ColumnNames[c] + "]");
                }
            }
        }
        private void ExtensionListBoxImport_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateExtensionListBoxImport();
        }
        private void ImportDelimeterListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateImportDelimeterListBox();
        }
        private void OutputTypeListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateOutputTypeListBox();
        }
        private void ExportDelimeterListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateExportDelimeterListBox();
        }
        private void ExportQualifierListBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            UpdateExportQualifierListBox();
        }
        private void ImportChooseFileButton_Click(object sender, EventArgs e)
        {
            //https://learn.microsoft.com/en-us/dotnet/desktop/winforms/controls/how-to-open-files-using-the-openfiledialog-component?view=netframeworkdesktop-4.8
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var FilePath = openFileDialog1.FileName;
                    string Extension = System.IO.Path.GetExtension(FilePath);
                    if (Extension != "")
                    {
                        Extension = Extension.Substring(1, Extension.Length - 1);
                    }

                    List<string> ValidExtensions = new List<string>();

                    foreach(string Ext in ExtensionListBoxImport.Items)
                    {
                        if(Ext == "xls*")
                        {
                            ValidExtensions.Add("xls");
                            ValidExtensions.Add("xlsm");
                            ValidExtensions.Add("xlsx");
                        }
                        else
                        {
                            ValidExtensions.Add(Ext);
                        }
                    }

                    if(ValidExtensions.FirstOrDefault(ValidExtension => ValidExtension.Contains(Extension)) != null)
                    {
                        UpdateExtensionListBoxImport(Extension);
                        ImportPathTextBox.Text = FilePath;
                    }                        
                }
                catch (SecurityException ex)
                {
                    MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                    $"Details:\n\n{ex.StackTrace}");
                }
            }
        }
        private void ExportChooseFileButton_Click(object sender, EventArgs e)
        {
            //https://learn.microsoft.com/en-us/dotnet/api/system.io.path.getfullpath?view=net-8.0
            if (folderBrowserDialog1.ShowDialog() == DialogResult.OK)
            {
                try
                {
                    var FolderPath = folderBrowserDialog1.SelectedPath;
                    ExportPathTextBox.Text = FolderPath.ToLower();
                }
                catch (SecurityException ex)
                {
                    MessageBox.Show($"Security error.\n\nError message: {ex.Message}\n\n" +
                    $"Details:\n\n{ex.StackTrace}");
                }
            }
        }
        private void TablesToExportCommaList_TextChanged(object sender, EventArgs e)
        {
            if (CommaSeperatedListTableSearchRadioButton.Checked == false)
            {
                CommaSeperatedListTableSearchRadioButton.Select();
                TablesToExportCommaList.Select();
            }
        }   
        private void TablesToExportRegex_TextChanged(object sender, EventArgs e)
        {
            if (RegexPatternTableSearchRadioButton.Checked == false)
            {
                RegexPatternTableSearchRadioButton.Select();
                TablesToExportRegex.Select();
            }
        }
        private void TablesToExportListFromSql_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (TablePickerRadioButton.Checked == false)
            {
                TablePickerRadioButton.Select();
                TablesToExportListFromSql.Select();
            }
        }
        private void InsertToExistingTableCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (InsertToExistingTableCheckBox.Checked == true)
            {
                ImportTableNameLabel.Font = new System.Drawing.Font(ImportTableNameLabel.Font, FontStyle.Bold);
            }
            else
            {
                ImportTableNameLabel.Font = new System.Drawing.Font(ImportTableNameLabel.Font, FontStyle.Regular);
            }
        }
        private void ImportTemplate_SelectedIndexChanged(object sender, EventArgs e)
        {
            string ImportTemplateText = ImportTemplate.Text;
            if (ImportTemplateText != "")
            {
                if(ImportTemplateText == "BMW GENESCO (UNQUOTED PIPE)")
                {
                    ServerComboBox.Text = "SQL04";
                    DatabaseComboBox.Text = "TEMP_BMW";
                    ExtensionListBoxImport.Text = "txt";
                    ImportDelimeterListBox.Text = "PIPE";
                    DoubleQuoted.Checked = false;
                    InsertToExistingTableCheckBox.Checked = false;
                    ImportToSingleTableTextBox.Text = "";
                }
                else if (ImportTemplateText == "BMW ALWAYS ON (UNQUOTED PIPE)")
                {
                    ServerComboBox.Text = "SQL04";
                    DatabaseComboBox.Text = "BMW";
                    ExtensionListBoxImport.Text = "txt";
                    ImportDelimeterListBox.Text = "PIPE";
                    DoubleQuoted.Checked = false;
                    InsertToExistingTableCheckBox.Checked = false;
                    ImportToSingleTableTextBox.Text = "";
                }
            }
        }
        private void ExportTemplate_SelectedIndexChanged(object sender, EventArgs e)
        {
            string ExportTemplateText = ExportTemplate.Text;
            if (ExportTemplateText != "")
            {
                if (ExportTemplateText == "BMW GENESCO (UNQUOTED PIPE)")
                {
                    ServerComboBox.Text = "SQL04";
                    DatabaseComboBox.Text = "TEMP_BMW";
                    OutputTypeListBox.Text = "txt";
                    ExportDelimeterListBox.Text = "PIPE";
                    ExportQualifierListBox.Text = "<NO QUALIFIER>";
                    IncludeHeadersCheckBox.Checked = true;
                    NoSplitRadioButton.Checked = true;
                    SplitAmounNumericUpDown.Value = 0;
                    SelectTextBox.Text = "";
                    FromTextBox.Text = "";
                    WhereTextBox.Text = "";
                    GroupByTextBox.Text = "";
                    OrderByTextBox.Text = "";
                }
                else if (ExportTemplateText == "BMW ALWAYS ON (UNQUOTED PIPE)")
                {
                    ServerComboBox.Text = "SQL04";
                    DatabaseComboBox.Text = "BMW";
                    OutputTypeListBox.Text = "txt";
                    ExportDelimeterListBox.Text = "PIPE";
                    ExportQualifierListBox.Text = "<NO QUALIFIER>";
                    IncludeHeadersCheckBox.Checked = true;
                    NoSplitRadioButton.Checked = true;
                    SplitAmounNumericUpDown.Value = 0;
                    SelectTextBox.Text = "";
                    FromTextBox.Text = "";
                    WhereTextBox.Text = "";
                    GroupByTextBox.Text = "";
                    OrderByTextBox.Text = "";
                }
                else if (ExportTemplateText == "ADROLL (NO HEADER, MAX SIZE 10MB)")
                {
                    OutputTypeListBox.Text = "csv";
                    ExportDelimeterListBox.Text = "COMMA";
                    ExportQualifierListBox.Text = "\"";
                    IncludeHeadersCheckBox.Checked = false;
                    SizeSplitRadioButton.Checked = true;
                    SplitAmounNumericUpDown.Value = 10;
                    SelectTextBox.Text = "";
                    FromTextBox.Text = "";
                    WhereTextBox.Text = "";
                    GroupByTextBox.Text = "";
                    OrderByTextBox.Text = "";
                }
                else if (ExportTemplateText == "LINKEDIN (MAX SIZE 19MB)")
                {
                    IncludeHeadersCheckBox.Checked = true;
                    SizeSplitRadioButton.Checked = true;
                    SplitAmounNumericUpDown.Value = 19;
                    SelectTextBox.Text = "";
                    FromTextBox.Text = "";
                    WhereTextBox.Text = "";
                    GroupByTextBox.Text = "";
                    OrderByTextBox.Text = "";
                }
                else if (ExportTemplateText == "TIKTOK (MAX SIZE 1000MB)")
                {
                    OutputTypeListBox.Text = "csv";
                    ExportDelimeterListBox.Text = "COMMA";
                    ExportQualifierListBox.Text = "\"";
                    IncludeHeadersCheckBox.Checked = true;
                    SizeSplitRadioButton.Checked = true;
                    SplitAmounNumericUpDown.Value = 1000;
                    SelectTextBox.Text = "";
                    FromTextBox.Text = "";
                    WhereTextBox.Text = "";
                    GroupByTextBox.Text = "";
                    //OrderByTextBox.Text = "";
                }
                else if (ExportTemplateText == "GOOGLE STRAIGHT FROM META")
                {
                    OutputTypeListBox.Text = "csv";
                    ExportDelimeterListBox.Text = "COMMA";
                    ExportQualifierListBox.Text = "\"";
                    IncludeHeadersCheckBox.Checked = true;
                    SizeSplitRadioButton.Checked = true;
                    SplitAmounNumericUpDown.Value = 1000;
                    SelectTextBox.Text = "  LOWER(EMAIL1) AS Email, LOWER(EMAIL2) AS Email, LOWER(EMAIL3) AS Email, REPLACE(REPLACE(REPLACE(RIGHT(PHONE1, LEN(PHONE1) - 1),'-',''),'(',''),')','') AS Phone, REPLACE(REPLACE(REPLACE(RIGHT(PHONE2, LEN(PHONE1) - 1),'-',''),'(',''),')','') AS Phone,REPLACE(REPLACE(REPLACE(RIGHT(PHONE3, LEN(PHONE1) - 1),'-',''),'(',''),')','') AS Phone, FN AS [First Name], LN AS [Last Name], 'US' AS Country, Zip ";
                    FromTextBox.Text = "";
                    WhereTextBox.Text = "";
                    GroupByTextBox.Text = "";
                    OrderByTextBox.Text = " PRIORITY ";
                }
                else if (ExportTemplateText == "TIKTOK STRAIGHT FROM META")
                {
                    OutputTypeListBox.Text = "csv";
                    ExportDelimeterListBox.Text = "COMMA";
                    ExportQualifierListBox.Text = "\"";
                    SelectTextBox.Text = "  LOWER(EMAIL1) AS email, LOWER(EMAIL2) AS email, LOWER(EMAIL3) AS email, (CASE WHEN PHONE1='' THEN '' ELSE '+'+REPLACE(REPLACE(REPLACE(PHONE1,'-',''),'(',''),')','') END) AS phone,  (CASE WHEN PHONE2='' THEN '' ELSE '+'+REPLACE(REPLACE(REPLACE(PHONE2,'-',''),'(',''),')','') END) AS phone,  (CASE WHEN PHONE3='' THEN '' ELSE '+'+REPLACE(REPLACE(REPLACE(PHONE3,'-',''),'(',''),')','') END) AS phone  ";
                    FromTextBox.Text = "";
                    WhereTextBox.Text = " (EMAIL1 IS NOT NULL AND EMAIL1 <>'') OR (EMAIL2 IS NOT NULL AND EMAIL2 <>'') OR (EMAIL3 IS NOT NULL AND EMAIL3 <>'') OR (PHONE1 IS NOT NULL AND PHONE1 <>'') OR (PHONE2 IS NOT NULL AND PHONE2 <>'') OR (PHONE3 IS NOT NULL AND PHONE3 <>'') ";
                    GroupByTextBox.Text = "";
                    OrderByTextBox.Text = " PRIORITY ";
                }
                else if (ExportTemplateText == "DIGI ORDERED BY PRIORITY")
                {
                    OutputTypeListBox.Text = "csv";
                    ExportDelimeterListBox.Text = "COMMA";
                    ExportQualifierListBox.Text = "\"";
                    IncludeHeadersCheckBox.Checked = true;
                    //SizeSplitRadioButton.Checked = true;
                    //SplitAmounNumericUpDown.Value = 1000;
                    SelectTextBox.Text = "";
                    FromTextBox.Text = "";
                    WhereTextBox.Text = "";
                    GroupByTextBox.Text = "";
                    OrderByTextBox.Text = "PRIORITY ASC";
                }
            }
        }
        private void NoSplitRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            if (NoSplitRadioButton.Checked == true)
            {
                SplitAmounNumericUpDown.Enabled = false;
                IncludeHeaderInSplitFilesCheckBox.Enabled = false;
            }
            else
            {
                SplitAmounNumericUpDown.Enabled = true;
                IncludeHeaderInSplitFilesCheckBox.Enabled = true;
            }
        }
        private void EnvironComboBox_SelectedIndexChanged(object sender, EventArgs e)
        {
            FormData FormData = new FormData();

            if (EnvironComboBox.Text == "SQL SERVER")
            {
                ServerComboBox.Text = string.Empty;
                ServerComboBox.Items.Clear();
                ServerComboBox.Enabled = true;

                DatabaseComboBox.Text = string.Empty;
                DatabaseComboBox.Items.Clear();

                AccountComboBox.Text = string.Empty;
                AccountComboBox.Items.Clear();

                SchemaComboBox.Text = string.Empty;
                SchemaComboBox.Items.Clear();

                UsernameComboBox.Text = string.Empty;
                UsernameComboBox.Items.Clear();

                PasswordComboBox.Text = string.Empty;
                PasswordComboBox.Items.Clear();

                if (ExtensionListBoxImport.SelectedItem.ToString() == "xls*")
                {
                    ConvertExcelToCsvCheckBox.Enabled = true;
                    ConvertExcelToCsvCheckBox.Checked = false;
                    DeleteConvertedExcelCsvsCheckBox.Enabled = true;
                    DeleteConvertedExcelCsvsCheckBox.Checked = false;
                    ExcelHeaderRowUpDown.Enabled = true;
                    ExcelHeaderRowUpDown.Value = 1;
                }

                FormData.SqlServerServers.ToList().ForEach(n => ServerComboBox.Items.Add(n));
                FormData.SqlServerDatabases.ToList().ForEach(n => DatabaseComboBox.Items.Add(n));
                ServerComboBox.Text = FormData.DefaultSqlServerServer;

                ExtensionListBoxImport.ClearSelected();
                ExtensionListBoxImport.Items.Clear();
                FormData.SqlServerImportTypes.ToList().ForEach(n => ExtensionListBoxImport.Items.Add(n));
                ExtensionListBoxImport.Text = "csv";

                ImportDelimeterListBox.ClearSelected();
                ImportDelimeterListBox.Items.Clear();
                FormData.SqlServerImportDelims.ToList().ForEach(n => ImportDelimeterListBox.Items.Add(n));
                ImportDelimeterListBox.Text = "COMMA";

                //DoubleQuoted.Enabled = true;
                //InsertToExistingTableCheckBox.Enabled = true;
                //ImportToSingleTableTextBox.Enabled = true;
                //FixedWidthColumnFilePathTextBox.Enabled = true;
                //ColumnTypeVarcharDefaultRadioButton.Enabled = true;
                //ColumnTypeUseFileRadioButton.Enabled = true;
                //ColumnTypeFilePathTextBox.Enabled = true;
                //FasterImportCheckBox.Enabled = true;
                ListColumnsForSelectedTablesButton.Enabled = true;
                LimitFieldLengthUpDown.Value = 255;

                ServerLabel.Font = new System.Drawing.Font(ServerLabel.Font, FontStyle.Bold);
                AccountLabel.Font = new System.Drawing.Font(AccountLabel.Font, FontStyle.Regular);
                UsernameLabel.Font = new System.Drawing.Font(UsernameLabel.Font, FontStyle.Regular);
                PasswordLabel.Font = new System.Drawing.Font(PasswordLabel.Font, FontStyle.Regular);
            }
            else if (EnvironComboBox.Text == "SNOWFLAKE")
            {
                ServerComboBox.Text = string.Empty;
                ServerComboBox.Items.Clear();
                ServerComboBox.Enabled = false;

                DatabaseComboBox.Text = string.Empty;
                DatabaseComboBox.Items.Clear();

                AccountComboBox.Text = string.Empty;
                AccountComboBox.Items.Clear();

                SchemaComboBox.Text = string.Empty;
                SchemaComboBox.Items.Clear();

                UsernameComboBox.Text = string.Empty;
                UsernameComboBox.Items.Clear();

                PasswordComboBox.Text = string.Empty;
                PasswordComboBox.Items.Clear();

                ConvertExcelToCsvCheckBox.Enabled = false;
                ConvertExcelToCsvCheckBox.Checked = true;
                DeleteConvertedExcelCsvsCheckBox.Enabled = true;
                DeleteConvertedExcelCsvsCheckBox.Checked = false;
                ExcelHeaderRowUpDown.Enabled = true;
                ExcelHeaderRowUpDown.Value = 1;

                ExtensionListBoxImport.ClearSelected();
                ExtensionListBoxImport.Items.Clear();
                FormData.SnowflakeImportTypes.ToList().ForEach(n => ExtensionListBoxImport.Items.Add(n));
                ExtensionListBoxImport.Text = "csv";

                ImportDelimeterListBox.ClearSelected();
                ImportDelimeterListBox.Items.Clear();
                FormData.SnowflakeImportDelims.ToList().ForEach(n => ImportDelimeterListBox.Items.Add(n));
                ImportDelimeterListBox.Text = "COMMA";

                FormData.SnowflakeDatabases.ToList().ForEach(n => DatabaseComboBox.Items.Add(n));
                FormData.SnowflakeAccounts.ToList().ForEach(n => AccountComboBox.Items.Add(n));
                FormData.SnowflakeUsernames.ToList().ForEach(n => UsernameComboBox.Items.Add(n));
                //FormData.SnowflakeSchemas.ToList().ForEach(n => SchemaComboBox.Items.Add(n));
                DatabaseComboBox.Text = FormData.DefaultSnowflakeDatabase;
                AccountComboBox.Text = FormData.DefaultSnowflakeAccount;
                //SchemaComboBox.Text = FormData.DefaultSnowflakeSchema;

                //DoubleQuoted.Enabled = false;
                //InsertToExistingTableCheckBox.Enabled = false;
                //ImportToSingleTableTextBox.Enabled = false;
                //FixedWidthColumnFilePathTextBox.Enabled = false;
                //ColumnTypeVarcharDefaultRadioButton.Enabled = false;
                //ColumnTypeUseFileRadioButton.Enabled = false;
                //ColumnTypeFilePathTextBox.Enabled = false;
                //FasterImportCheckBox.Enabled = false;
                ListColumnsForSelectedTablesButton.Enabled = false;
                LimitFieldLengthUpDown.Value = 0;

                ServerLabel.Font = new System.Drawing.Font(ServerLabel.Font, FontStyle.Regular);
                AccountLabel.Font = new System.Drawing.Font(AccountLabel.Font, FontStyle.Bold);
                UsernameLabel.Font = new System.Drawing.Font(UsernameLabel.Font, FontStyle.Bold);
                PasswordLabel.Font = new System.Drawing.Font(PasswordLabel.Font, FontStyle.Bold);
            }
        }

        private void ConvertExcelToCsvCheckBox_CheckedChanged(object sender, EventArgs e)
        {
            if (ConvertExcelToCsvCheckBox.Checked == true && (string)ExtensionListBoxImport.SelectedItem == "xls*")
            {
                ImportEncodingListBox.Enabled = true;
                DeleteConvertedExcelCsvsCheckBox.Enabled = true;
            }
            else if (ConvertExcelToCsvCheckBox.Checked == false && (string)ExtensionListBoxImport.SelectedItem == "xls*")
            {
                ImportEncodingListBox.Enabled = false;
                DeleteConvertedExcelCsvsCheckBox.Enabled = false;
            }
        }
    }
}

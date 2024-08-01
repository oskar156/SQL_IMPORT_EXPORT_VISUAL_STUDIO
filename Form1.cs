/* TABLE OF CONTENTS
 * 
 * fix writeline/writes
 * fix when you select excel import (and export) doublequotes
 * 
 * IMPORT FUNCTIONS
 * private void ImportButton_Click(object sender, EventArgs e)
 * 
 * EXPORT FUNCTIONS
 * private void ExportButton_Click(object sender, EventArgs e)
 * 
 * OTHER
 * private List<string> GetListOfUserSelectedTables()
 * 
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
 * 
 * NOTES/ISSUES
 * Weird characters will be removed completely
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
using System.Diagnostics;

namespace SQL_SERVER_IMPORT_EXPORT
{
    public partial class Form1 : Form
    {

        //CONSTANTS
        public string[] ValidExtensions = new string[] { "csv", "txt", "xls", "xlsx", "xlsm" };

        public Form1()
        {
            InitializeComponent();

            //default values
            ExtensionListBoxImport.SetSelected(0, true); //csv
            ImportDelimeterListBox.SetSelected(0, true); //COMMA

            OutputTypeListBox.SetSelected(0, true); //csv
            ExportDelimeterListBox.SetSelected(0, true); //COMMA
            ExportQualifierListBox.SetSelected(0, true); //"

            TablePickerRadioButton.Select(); //Select tables to export from SQL table picker
            NoSplitRadioButton.Select(); //Select option to not split files on export

            SplitAmounNumericUpDown.Maximum = int.MaxValue - 1; //setting maximum value for split amount

            //cant get current path of shortcut that's used to run this!
            string CurrentPath = System.IO.Path.GetDirectoryName(System.Windows.Forms.Application.ExecutablePath);
            //set export/import paths to the path where the application is located
            //this is useless unless we can find the location of the shortcut that started the application
            //ImportPathTextBox.Text = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetEntryAssembly().Location); ;
            //ExportPathTextBox.Text = System.IO.Directory.GetCurrentDirectory();//Environment.CurrentDirectory;
            //SelectTextBox.Text = System.AppDomain.CurrentDomain.BaseDirectory;// System.IO.Path.GetDirectoryName(Process.GetCurrentProcess().MainModule.FileName);
            //GroupByTextBox.Text = Application.UserAppDataPath;//Application.StartupPath;
            
        }

        private void Form1_Load(object sender, EventArgs e)
        {

        }




        //------------------------------------------------------------------------------------
        // IMPORT
        /*
         * gets user input
         * 
         * gets file(s) to import
         * 
         * for each file:
         * 
         *   if ImportToSingleTable == false || (ImportToSingleTable == true && 1st file):
         *     reads headers into DataTable
         *     
         *     if ImportToExistingTable == false:
         *       creates SQL table
         *   
         *   for each row:
         *     
         *     PLACED IN HELPER FUNCTION:
         *       reads each row into DataTable
         *       if BatchLimit is reached:
         *         inserts DataTable into SQL table
         *       imports any remaining data not in previous batches into SQL table
         * 
         */
        //------------------------------------------------------------------------------------
        private void ImportButton_Click(object sender, EventArgs e)
        {
            Console.WriteLine("------------------------------");
            Console.WriteLine("Import Button Clicked");
            Console.WriteLine("------------------------------");

            Helpers Helpers = new Helpers();

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

            bool IsDoubleQuoted = DoubleQuoted.Checked;

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

            bool FasterImport = FasterImportCheckBox.Checked;

            string ImportPath = ImportPathTextBox.Text;

            Console.WriteLine("Server.Databse: " + Server + "." + Database);
            Console.WriteLine("BatchLimit: " + BatchLimit.ToString());

            /***************************************
             * GET FILES TO IMPORT
            ***************************************/
            List<string> FilesToImport = Helpers.GetFilesToImport(ImportPath, Extension);

            int FilesImported = 0;
            int TablesCreated = 0;
            string TableName = "";
            List<string> TableNames = new List<string>();
            DataTable BaseDtTable = new DataTable();
            List<DataTable> BaseDtTables = new List<DataTable>();

            /***************************************
             * FOR EACH FILE TO IMPORT
            ***************************************/
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
                        BaseDtTables = Helpers.ReadExcelFileIntoDataTablesWithColumns(FilePath, ref TableNames);
                        //will also add to TableNames list
                    }
                    else
                    {
                        BaseDtTable = Helpers.ReadFileIntoDataTableWithColumns(FilePath, ActualDelimeter, FixedWidthColumnFilePath);
                    }

                    /***************************************
                     * CREATE TABLE IN SQL
                    ***************************************/
                    if (ImportToExistingTable == false)
                    {
                        if (Extension == "xls*")
                        {
                            Helpers.CreateTablesInSqlVarchar(TableNames, Server, Database, BaseDtTables, ActualDelimeter, ColumnTypeMethod, ColumnTypeFilePath);
                            TablesCreated += BaseDtTables.Count;
                        }
                        else
                        {
                            Helpers.CreateTableInSqlVarchar(TableName, Server, Database, BaseDtTable, ActualDelimeter, ColumnTypeMethod, ColumnTypeFilePath);
                            TablesCreated++;
                        }
                    }
                }

                /***************************************
                 * READ FILE ROWS AND INSERT INTO SQL TABLE
                ***************************************/
                if (Extension == "xls*")
                {
                    if (FasterImport == true)
                    {
                        //faster but more error prone to irregular characters and field lengths over 255
                        Helpers.ReadExcelFilePerSheetIntoDataTablesWithRowsAndInsertIntoSqlTablesFast(FilePath, TableNames, Server, Database, BaseDtTables, BatchLimit, ActualDelimeter);
                    }
                    else if (FasterImport == false)
                    {
                        Helpers.ReadExcelFilePerSheetIntoDataTablesWithRowsAndInsertIntoSqlTables(FilePath, TableNames, Server, Database, BaseDtTables, BatchLimit, ActualDelimeter);
                    }
                    FilesImported += BaseDtTables.Count;
                }
                else
                {
                    if(FasterImport == true)
                    {
                        //faster but more error prone to irregular characters and field lengths over 255
                        Helpers.ReadFileIntoDataTableWithRowsAndInsertIntoSqlTableFast(FilePath, TableName, Server, Database, BaseDtTable, BatchLimit, ActualDelimeter, IsDoubleQuoted);
                    }
                    else if (FasterImport == false)
                    {
                        Helpers.ReadFileIntoDataTableWithRowsAndInsertIntoSqlTable(FilePath, TableName, Server, Database, BaseDtTable, BatchLimit, ActualDelimeter, IsDoubleQuoted);
                    }
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





        //------------------------------------------------------------------------------------
        // EXPORT
        /*
         * gets user input
         * 
         * gets table(s) to export
         * 
         * for each table:
         *   export to a new file
         *   
         *   (IDEA: EXPORT TO EXCEL WORKBOOK, 1 TABLE PER TAB)
         *   (IDEA: EXPORT TO SINGLE FILE, SAME HEADERS)
         * 
         */
        //------------------------------------------------------------------------------------
        private void ExportButton_Click(object sender, EventArgs e)
        {
            Console.WriteLine("------------------------------");
            Console.WriteLine("Export Button Clicked");
            Console.WriteLine("------------------------------");

            Helpers Helpers = new Helpers();

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

            string ExportPath = ExportPathTextBox.Text; //must be a folder

            if(ExportPath == "")
            {
                Console.WriteLine("No Export Path specified!\nExport is cancelled.\n\n");
                return;
            }
            if (Directory.Exists(ExportPath) == false)
            {
                Console.WriteLine("Export Path does not exist!\nCreating Export Path...\n\n");
                Directory.CreateDirectory(ExportPath);
            }

            Console.WriteLine("Server.Databse: " + Server + "." + Database);

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
                TablesFromSqlDb = Helpers.GetListofTablesFromSqlDb(Server, Database, TablesToExportCommaStrList);
            }
            else if (TableSearchMethodIsRegexPattern)
            {
                //User types out a regex pattern, which is checked against tables that exist in SQL
                //Only the table names that match are returned
                TablesFromSqlDb = Helpers.GetListofTablesFromSqlDb(Server, Database, TablesToExportRegexText);
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
                int FilesCreated = Helpers.ExportTableFromSqlToFile(Server, Database, TableName, ExportPath, Extension, ActualDelimeter, ExportQualifier, QualifyEveryField, RemoveQualInVal, IncludeHeaders, "MAX LEN", SizeLimit, SizeLimitType, IncludeHeaderInSplitFiles, SelectText, FromText, WhereText, GroupByText, OrderByText);
                FilesExported  += FilesCreated;
                Console.WriteLine("");
            }

            Console.WriteLine(TablesExported.ToString() + " tables exported from " + Server + "." + Database);
            Console.WriteLine(FilesExported.ToString() + " files created " + ExportPath);
            Console.WriteLine("");
        }



        //------------------------------------------------------------------------------------
        // OTHER
        //------------------------------------------------------------------------------------
        private List<string> GetListOfUserSelectedTables()
        {
            List<string> UserSelectedTables = new List<string>();
            Helpers Helpers = new Helpers();

            string Server = ServerComboBox.Text;
            string Database = DatabaseComboBox.Text;
            bool TableSearchMethodIsCommaList = CommaSeperatedListTableSearchRadioButton.Checked;
            bool TableSearchMethodIsRegexPattern = RegexPatternTableSearchRadioButton.Checked;
            bool TableSearchMethodIsTablePicker = TablePickerRadioButton.Checked;

            string TablesToExportCommaListText = TablesToExportCommaList.Text;
            List<string> TablesToExportCommaStrList = TablesToExportCommaListText.Split(',').ToList<string>();
            string TablesToExportRegexText = TablesToExportRegex.Text;
            List<string> TablesToExportListFromSqlList = TablesToExportListFromSql.SelectedItems.Cast<string>().ToList();

            if (TableSearchMethodIsCommaList)
            {
                //User types out comma-seperated list, which is checked against tables that exist in SQL
                //Only the table names that match are returned
                UserSelectedTables = Helpers.GetListofTablesFromSqlDb(Server, Database, TablesToExportCommaStrList);
            }
            else if (TableSearchMethodIsRegexPattern)
            {
                //User types out a regex pattern, which is checked against tables that exist in SQL
                //Only the table names that match are returned
                UserSelectedTables = Helpers.GetListofTablesFromSqlDb(Server, Database, TablesToExportRegexText);
            }
            else if (TableSearchMethodIsTablePicker)
            {
                //User picks from a list of tables that exist in SQLs
                //We move forward with exactly the user input, because it definitely already exists in SQL
                UserSelectedTables = TablesToExportListFromSqlList;
            }

            return UserSelectedTables;
        }

        //------------------------------------------------------------------------------------
        // FORM EVENT FUNCTIONS
        //------------------------------------------------------------------------------------
        private void LoadSqlTables_Click(object sender, EventArgs e)
        {
            string Server = ServerComboBox.Text;
            string Database = DatabaseComboBox.Text;
            Helpers Helpers = new Helpers();

            if (Server != "" && Database != "")
            {
                TablesToExportListFromSql.Items.Clear();
                List<string> Tables = Helpers.GetListofTablesFromSqlDb(Server, Database, "", false);
                for (int i = 0; i < Tables.Count; i++)
                {
                    TablesToExportListFromSql.Items.Add(Tables[i]);
                }
            }
        }


        private void ListColumnsForSelectedTablesButton_Click(object sender, EventArgs e)
        {
            Console.WriteLine("COLUMNS FOR SELECTED TABLES");
            Helpers Helpers = new Helpers();

            string Server = ServerComboBox.Text;
            string Database = DatabaseComboBox.Text;

            List<string> UserSelectedTables = GetListOfUserSelectedTables();

            //for every table selected
            for (int t = 0; t < UserSelectedTables.Count; t++)
            {
                string Table = UserSelectedTables[t];
                //get columns
                List<string> ColumnNames = Helpers.GetListOfColumnsForTable(Server, Database, Table);

                Console.WriteLine(ColumnNames.Count + " columns");

                //list them in output window
                for (int c = 0; c < ColumnNames.Count; c++)
                {
                    Console.WriteLine("[" + ColumnNames[c] + "]");
                }
            }
        }

        //makes changes to the form
        //so the user knows what fields are required, options that are not allowed, etc...
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

                    if (Array.IndexOf(this.ValidExtensions, Extension) >= 0)
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


        //------------------------------------------------------------------------------------
        // FORM
        //------------------------------------------------------------------------------------
        public void UpdateExtensionListBoxImport(string Value = "")
        {

            string ValueSelected = "";
            if (Value == "")
            {
                ValueSelected = ExtensionListBoxImport.SelectedItem.ToString();
            }
            else
            {
                ValueSelected = Value;
                if (ValueSelected == "xls" || ValueSelected == "xlsm" || ValueSelected == "xlsx")
                {
                    ExtensionListBoxImport.SelectedItem = "xls*";
                }
                else
                {
                    ExtensionListBoxImport.SelectedItem = Value;
                }
            }

            if (ValueSelected == "xls*" || ValueSelected == "xls" || ValueSelected == "xlsm" || ValueSelected == "xlsx")
            {
                if (ImportDelimeterListBox.Enabled)
                {
                    string ImportDelimeterCurrentlySelected = ImportDelimeterListBox.SelectedItem.ToString();
                    ImportDelimeterListBox.ClearSelected();
                    ImportDelimeterListBox.Enabled = false;
                }
                else
                { 
                    if(ImportDelimeterListBox.SelectedItem != null)
                    { 
                        string ImportDelimeterCurrentlySelected = ImportDelimeterListBox.SelectedItem.ToString();
                    }
                    ImportDelimeterListBox.ClearSelected();
                }
            }
            else if(ValueSelected == "csv")
            {
                ImportDelimeterListBox.Enabled = true;
                ImportDelimeterListBox.SelectedItem = "COMMA";
                //ImportDelimeterListBox.Enabled = false;
            }
            else
            {
                ImportDelimeterListBox.Enabled = true;
                if (ImportDelimeterListBox.SelectedItem == null)
                {
                    ImportDelimeterListBox.SelectedItem = "COMMA";
                }
            }
        }
        public void UpdateImportDelimeterListBox()
        {

            if (ImportDelimeterListBox.SelectedItem != null)
            {
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

        //TEMPLATES
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
    }
}

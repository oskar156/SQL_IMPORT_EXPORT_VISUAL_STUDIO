/* TABLE OF CONTENTS
 * 
 * METHODS
 * public List<string> GetListOfUserSelectedTables()
 * public List<string> GetListofTablesFromSqlServerDb(ConnectionInfo ConnectionInfo, List<string> ListOfTablesToSearchFor, bool ConsoleOutput = true)
 * public List<string> GetListofTablesFromSqlServerDb(ConnectionInfo ConnectionInfo, string RegexSearchPattern = "", bool ConsoleOutput = true)
 * public List<string> GetListofTablesFromSnowflakeDb(Snowflake Snowflake, ConnectionInfo ConnectionInfo, List<string> ListOfTablesToSearchFor, bool ConsoleOutput = true)
 * public List<string> GetListofTablesFromSnowflakeDb(Snowflake Snowflake, ConnectionInfo ConnectionInfo, string RegexSearchPattern = "", bool ConsoleOutput = true)
 * public List<string> GetListOfColumnsForTable(ConnectionInfo ConnectionInfo, string TableName, bool ConsoleOutput = true)
 * public int ExportTableFromSqlServerToFile(ConnectionInfo ConnectionInfo, string TableToExport, string ExportPath, string Extension, string Delimeter, System.Text.Encoding Encoding, string Qualifier, bool QualifyEveryField, bool RemoveQualInVal, bool IncludeHeaders, string FixedWidthColumnLengthMethod, decimal SizeLimit, string SizeLimitType, bool IncludeHeaderInSplitFiles, string SelectText = "", string FromText = "", string WhereText = "", string GroupBy = "", string OrderBy = "", bool ConsoleOutput = true)
 * public int ExportTableFromSnowflakeToFile(Snowflake Snowflake, string TableToExport, string ExportPath, string Extension, string Delimeter, System.Text.Encoding Encoding, string Qualifier, bool QualifyEveryField, bool RemoveQualInVal, bool IncludeHeaders, string FixedWidthColumnLengthMethod, decimal SizeLimit, string SizeLimitType, bool IncludeHeaderInSplitFiles, string SelectText = "", string FromText = "", string WhereText = "", string GroupBy = "", string OrderBy = "", bool ConsoleOutput = true)
 * 
*/

using DocumentFormat.OpenXml.Spreadsheet;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using System.Windows.Forms;

namespace SQL_SERVER_IMPORT_EXPORT
{
    public class Export
    {
        //======================================================================================================================
        // METHODS
        //======================================================================================================================
        public List<string> GetListOfUserSelectedTables()
        {
            List<string> UserSelectedTables = new List<string>();
            Export Export = new Export();
            ConnectionInfo ConnectionInfo = new ConnectionInfo();

            RadioButton CommaSeperatedListTableSearchRadioButton = Application.OpenForms["Form1"].Controls["CommaSeperatedListTableSearchRadioButton"] as RadioButton;
            RadioButton RegexPatternTableSearchRadioButton = Application.OpenForms["Form1"].Controls["RegexPatternTableSearchRadioButton"] as RadioButton;
            RadioButton TablePickerRadioButton = Application.OpenForms["Form1"].Controls["TablePickerRadioButton"] as RadioButton;
            TextBox TablesToExportCommaList = Application.OpenForms["Form1"].Controls["TablesToExportCommaList"] as TextBox;
            TextBox TablesToExportRegex = Application.OpenForms["Form1"].Controls["TablesToExportRegex"] as TextBox;
            ListBox TablesToExportListFromSql = Application.OpenForms["Form1"].Controls["TablesToExportListFromSql"] as ListBox;

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
                UserSelectedTables = Export.GetListofTablesFromSqlServerDb(ConnectionInfo, TablesToExportCommaStrList);
            }
            else if (TableSearchMethodIsRegexPattern)
            {
                //User types out a regex pattern, which is checked against tables that exist in SQL
                //Only the table names that match are returned
                UserSelectedTables = Export.GetListofTablesFromSqlServerDb(ConnectionInfo, TablesToExportRegexText);
            }
            else if (TableSearchMethodIsTablePicker)
            {
                //User picks from a list of tables that exist in SQLs
                //We move forward with exactly the user input, because it definitely already exists in SQL
                UserSelectedTables = TablesToExportListFromSqlList;
            }

            return UserSelectedTables;
        }
        public List<string> GetListofTablesFromSqlServerDb(ConnectionInfo ConnectionInfo, List<string> ListOfTablesToSearchFor, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Getting List of Tables from Sql Server"); }

            List<string> Tables = new List<string>();

            //convert to upper so we can do case-insensitive matching
            for (int t = 0; t < ListOfTablesToSearchFor.Count; t++)
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
                re = new Regex("(?i)" + RegexSearchPattern); //(?i) makes it case-insensitive
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
        public List<string> GetListofTablesFromSnowflakeDb(Snowflake Snowflake, ConnectionInfo ConnectionInfo, List<string> ListOfTablesToSearchFor, bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Getting List of Tables from Snowflake"); }

            List<string> Tables = new List<string>();

            //convert to upper so we can do case-insensitive matching
            for (int t = 0; t < ListOfTablesToSearchFor.Count; t++)
            {
                ListOfTablesToSearchFor[t] = ListOfTablesToSearchFor[t].ToUpper();
            }

            string Query = "select table_name from information_schema.tables where table_type = 'BASE TABLE'";
            if(ConnectionInfo.Schema != "")
            {
                Query += " AND TABLE_SCHEMA = '" + ConnectionInfo.Schema + "' ";
            }
            Query += ";";
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
        public List<string> GetListofTablesFromSnowflakeDb(Snowflake Snowflake, ConnectionInfo ConnectionInfo, string RegexSearchPattern = "", bool ConsoleOutput = true)
        {
            if (ConsoleOutput) { Console.WriteLine("Getting Tables from Snowflake"); }

            List<string> Tables = new List<string>();

            Regex re = new Regex("");
            if (RegexSearchPattern != "")
            {
                re = new Regex("(?i)" + RegexSearchPattern); //(?i) makes is case-insensitive
            }

            string Query = "select table_name from information_schema.tables where table_type = 'BASE TABLE'";
            if (ConnectionInfo.Schema != "")
            {
                Query += " AND TABLE_SCHEMA = '" + ConnectionInfo.Schema + "' ";
            }
            Query += ";";
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
        public int ExportTableFromSqlServerToFile(ConnectionInfo ConnectionInfo, string TableToExport, string ExportPath, string Extension, string Delimeter, System.Text.Encoding Encoding, string Qualifier, bool QualifyEveryField, bool RemoveQualInVal, bool IncludeHeaders, string FixedWidthColumnLengthMethod, decimal SizeLimit, string SizeLimitType, bool IncludeHeaderInSplitFiles, string SelectText = "", string FromText = "", string WhereText = "", string GroupBy = "", string OrderBy = "", bool ConsoleOutput = true)
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

                    if (ConsoleOutput) { if (ExcelRowIndex % 1000000 != 0) { Console.WriteLine("\rRows exported: " + $"{ExcelRowIndex:n0}"); } }
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

                    StreamWriter sw = new StreamWriter(FileExportPath, true, Encoding);
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
                                if (RemoveQualInVal && Qualifier != "" && Qualifier != null)
                                {
                                    TblColClean = TblColClean.Replace(Qualifier, "");
                                }

                                //write to HeaderRow
                                if (QualifyEveryField || TblColClean.Contains(Delimeter) || TblColClean.Contains("\\"))
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
                            //}s
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
                                sw = new StreamWriter(FileExportPathBase + "-" + SplitFileIndex.ToString() + "." + Extension, true, Encoding);
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
                                    sw = new StreamWriter(FileExportPathBase + "-" + SplitFileIndex.ToString() + "." + Extension, true, Encoding);
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
        public int ExportTableFromSnowflakeToFile(Snowflake Snowflake, string TableToExport, string ExportPath, string Extension, string Delimeter, System.Text.Encoding Encoding, string Qualifier, bool QualifyEveryField, bool RemoveQualInVal, bool IncludeHeaders, string FixedWidthColumnLengthMethod, decimal SizeLimit, string SizeLimitType, bool IncludeHeaderInSplitFiles, string SelectText = "", string FromText = "", string WhereText = "", string GroupBy = "", string OrderBy = "", bool ConsoleOutput = true)
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

                StreamWriter sw = new StreamWriter(FileExportPath, true, Encoding);
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
                        ConnectionInfo ConnectionInfo = new ConnectionInfo();
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
                            sw = new StreamWriter(FileExportPathBase + "-" + SplitFileIndex.ToString() + "." + Extension, true, Encoding);
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
                                sw = new StreamWriter(FileExportPathBase + "-" + SplitFileIndex.ToString() + "." + Extension, true, Encoding);
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
    }
}

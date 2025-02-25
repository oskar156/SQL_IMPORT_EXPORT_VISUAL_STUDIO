/* TABLE OF CONTENTS
 * 
 * FIELDS...
 * 
 * CONSTRUCTORS...
 * 
 * METHODS
 * public void ConnectToDb(ConnectionInfo ConnectionInfo)
 * public void Execute(string Query)
 * public void StageFile(string FilePath, string StageName)
 * public void ImportFile(string FilePath, string StageName, DataTable BaseDtTable, string TableName, string Delimiter, bool DoubleQuoted, System.Text.Encoding EncodingChoice, string EncodingChoiceSnowflake, bool StripNonPrintableChars, int FieldLimit, string ColumnTypeMethod = "DEFAULT VARCHAR", string ColumnTypeFilePath = "", bool ImportToExistingTable = false)
 * public void Close()
 * 
 * NOTES
 * https://github.com/snowflakedb/snowflake-connector-net/issues/895
 * https://community.snowflake.com/s/article/How-to-connect-to-snowflake-using-C-Sharp-application-with-snowflake-NET-Connector-to-perform-SQL-operations-in-windows
 * in your project, open NuGet package manager console (in VisualStudio it's Tools > Nuget Package Manager > Package Manager console), 
 * then after Powershell is loaded, issue this command to install Mono.Unix: PM> NuGet\Install-Package Mono.Unix -Version 7.1.0-final.1.21458.1
 * also installed snowflake w/nuget
 * 
 */

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Threading;
using Snowflake.Data.Client;

namespace SQL_SERVER_IMPORT_EXPORT
{
    public class Snowflake
    {
        //======================================================================================================================
        // FIELDS
        //======================================================================================================================
        public IDbConnection Conn;
        public IDataReader Reader;

        //======================================================================================================================
        // CONSTRUCTORS
        //======================================================================================================================
        public Snowflake()
        {

        }

        //======================================================================================================================
        // METHODS
        //======================================================================================================================
        public void ConnectToDb(ConnectionInfo ConnectionInfo)
        {
            //MinMax Threads are reduced to limit the issue of indefinite Duo MFA requests
            //setting min-max to 1-4 results in nothing
            //setting min-max to 1-5 results in 2-4 MFA requests on average - sometimes we get nothing and it times ou
            //ideally it would just be once, though
            int MinThreadsWorker;
            int MinThreadsCompletionPort;
            int MaxThreadsWorker;
            int MaxThreadsCompletionPort;
            ThreadPool.GetMinThreads(out MinThreadsWorker, out MinThreadsCompletionPort);
            ThreadPool.GetMaxThreads(out MaxThreadsWorker, out MaxThreadsCompletionPort);
            ThreadPool.SetMinThreads(1, 1);
            ThreadPool.SetMaxThreads(5, 5);

            //Open Connection
            this.Conn = new SnowflakeDbConnection();
            string ConnectionString = "account=" + ConnectionInfo.Account + ";user=" + ConnectionInfo.Username + ";password=" + ConnectionInfo.Password;
            this.Conn.ConnectionString = ConnectionString;
            Console.WriteLine("If you have MFA enabled, then please authorize to continue (may take 2-4 taps)...");
            this.Conn.Open();
            Console.WriteLine("Snowflake connection opened!");

            //rest threads to what they were
            ThreadPool.SetMinThreads(MinThreadsWorker, MinThreadsCompletionPort);
            ThreadPool.SetMaxThreads(MaxThreadsWorker, MaxThreadsCompletionPort);

            //Set Database
            string QueryDb = "USE DATABASE " + ConnectionInfo.Database + ";";
            this.Execute(QueryDb);
            
            //Set SCHEMA
            if(ConnectionInfo.Schema != "")
            {
                string QuerySchm = "USE SCHEMA " + ConnectionInfo.Schema + ";";
                this.Execute(QuerySchm);
            }
        }
        public void Execute(string Query)
        {
            IDbCommand Cmd = this.Conn.CreateCommand();
            Cmd.CommandText = Query;
            this.Reader = Cmd.ExecuteReader();
        }
        public void StageFile(string FilePath, string StageName)
        {
            string FilePathForwardSlash = FilePath.Replace("\\", "/");
            string StagingQuery = "PUT 'file://" + FilePathForwardSlash + "' @~/" + StageName + "/ OVERWRITE = TRUE;";
            this.Execute(StagingQuery);
        }
        public void ImportFile(string FilePath, string StageName, DataTable BaseDtTable, string TableName, string Delimiter, bool DoubleQuoted, System.Text.Encoding EncodingChoice, string EncodingChoiceSnowflake, bool StripNonPrintableChars, int FieldLimit, string ColumnTypeMethod = "DEFAULT VARCHAR", string ColumnTypeFilePath = "", bool ImportToExistingTable = false)
        {
            Helpers Helpers = new Helpers();
            List<string> ColumnTypes = new List<string>();

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

            //for each Column, write column select part of the query
            string ColumnSelects = "";

            int ColumnIndex = 1;
            int SubstringIndexStart = 1; //SUBSTRING is 1-indexed in snowflake https://docs.snowflake.com/en/sql-reference/functions/substr

            string ReplaceChar = "";
            if (Delimiter == "FIXED WIDTH")
            {
                ReplaceChar = " ";
            }

            foreach (DataColumn DataColumn in BaseDtTable.Columns)
            {
                string ColName = DataColumn.ColumnName;
                int SubstringLength = DataColumn.MaxLength;

                //either fixed width or field limit, not both!
                if (Delimiter == "FIXED WIDTH") { ColumnSelects += " SUBSTRING( "; }
                else if (FieldLimit > 0)        { ColumnSelects += " LEFT( "; }

                if (StripNonPrintableChars)     { ColumnSelects += " REGEXP_REPLACE( "; }

                ColumnSelects += "$" + ColumnIndex.ToString();

                if (StripNonPrintableChars) { ColumnSelects += " , '[^\u0020-\u007E]+', '" + ReplaceChar + "')"; }

                //either fixed width or field limit, not both!
                if (Delimiter == "FIXED WIDTH") { ColumnSelects += " , " + SubstringIndexStart.ToString() + " " + SubstringLength.ToString() + "  )"; }
                else if (FieldLimit > 0) { ColumnSelects += " , " + FieldLimit.ToString() + ")"; }

                if (ColumnTypeFilePath != "" && ColumnTypeMethod == "FILE PATH")
                {
                    ColumnSelects += "::" + ColumnTypes[ColumnIndex - 1] + " AS \"" + ColName + "\" ,"; //uses column types in file
                }
                else
                {
                    ColumnSelects += "::VARCHAR AS \"" + ColName + "\" ,"; //assums all columns will be VARCHAR
                }
                //fw example:    SUBSTRING(REGEXP_REPLACE($1, '[^\u0020-\u007E]+', ''), 1, 20)::VARCHAR AS "FIRST NAME",
                //limit example: LEFT(REGEXP_REPLACE($1, '[^\u0020-\u007E]+', ''), 255)::VARCHAR AS "FIRST NAME",

                SubstringIndexStart += SubstringLength;
                ColumnIndex++;
            }
            ColumnSelects = ColumnSelects.Substring(0, ColumnSelects.Length - 1); //remove last comma

            string FileFormatName = "TEMP_FILE_FORMAT";
            string FileFormatQuery = " CREATE OR REPLACE TEMPORARY FILE FORMAT " + FileFormatName + " ";
            FileFormatQuery += " type = 'CSV' "; //works for all delimiter types, not just comma
            FileFormatQuery += " field_delimiter = '" + Delimiter + "' ";
            FileFormatQuery += " skip_header=1 "; //skip the header row = true
            FileFormatQuery += " FIELD_OPTIONALLY_ENCLOSED_BY='\"' ";
            FileFormatQuery += " ENCODING='" + EncodingChoiceSnowflake + "' ";
            //if (StripNonPrintableChars) { FileFormatQuery += " REPLACE_INVALID_CHARACTERS=TRUE  "; }
            FileFormatQuery += " ; ";
            //Console.WriteLine(FileFormatQuery);
            this.Execute(FileFormatQuery);

            //https://docs.snowflake.com/en/sql-reference/sql/insert
            string CreateTableFromStagedFileQuery = "";
            if (ImportToExistingTable) { CreateTableFromStagedFileQuery += " INSERT INTO \"" + TableName + "\" AS "; }
            else { CreateTableFromStagedFileQuery += " CREATE OR REPLACE TABLE \"" + TableName + "\" AS "; }
            CreateTableFromStagedFileQuery += " SELECT " + ColumnSelects + " ";
            CreateTableFromStagedFileQuery += " FROM @~/" + StageName + " ";
            CreateTableFromStagedFileQuery += " (file_format => " + FileFormatName + "); ";
            //Console.WriteLine(CreateTableFromStagedFileQuery);
            this.Execute(CreateTableFromStagedFileQuery);

            string RemoveStagingAreaQuery = " REMOVE @~/" + StageName + "/; ";
            //Console.WriteLine(RemoveStagingAreaQuery);
            this.Execute(RemoveStagingAreaQuery);

            string DropFileFormatQuery = " DROP FILE FORMAT " + FileFormatName + "; ";
            //Console.WriteLine(DropFileFormatQuery);
            this.Execute(DropFileFormatQuery);
        }
        public void Close()
        {
            this.Conn.Close();
        }
    }
}

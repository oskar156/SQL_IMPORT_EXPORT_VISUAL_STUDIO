using System;
using System.CodeDom.Compiler;
using System.Collections.Generic;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Security.Cryptography.X509Certificates;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using DocumentFormat.OpenXml.Wordprocessing;
using Mono.Unix.Native;
using Snowflake.Data.Client;
using static System.Windows.Forms.VisualStyles.VisualStyleElement.ToolTip;
/*
 * https://github.com/snowflakedb/snowflake-connector-net/issues/895
 * in your project, open NuGet package manager console (in VisualStudio it's Tools > Nuget Package Manager > Package Manager console), 
 * then after Powershell is loaded, issue this command to install Mono.Unix: PM> NuGet\Install-Package Mono.Unix -Version 7.1.0-final.1.21458.1
 * also installed snowflake w/nuget
 * 
 * https://community.snowflake.com/s/article/How-to-connect-to-snowflake-using-C-Sharp-application-with-snowflake-NET-Connector-to-perform-SQL-operations-in-windows
 */

namespace SQL_SERVER_IMPORT_EXPORT
{
    public class Snowflake
    {
        //fields
        public IDbConnection Conn;
        public IDataReader Reader;

        //constructor
        public Snowflake()
        {

        }

        //methods
        public void ConnectToDb(ConnectionInfo ConnectionInfo)
        {
            //MinMax Threads are reduced to limit the issue of indefinite Duo MFA requests
            //setting min-max to 1-4 results in nothing
            //setting min-max to 1-5 results in 2 MFA requests
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
            Console.WriteLine("If you have MFA enabled, then please authorize to continue (may take 2-3 taps)...");
            this.Conn.Open();
            Console.WriteLine("Snowflake connection opened!");

            //rest threads to what they were
            ThreadPool.SetMinThreads(MinThreadsWorker, MinThreadsCompletionPort);
            ThreadPool.SetMaxThreads(MaxThreadsWorker, MaxThreadsCompletionPort);

            //Set Database
            string Query = "USE DATABASE " + ConnectionInfo.Database + ";";
            this.Execute(Query);
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

        public void ImportFile(string FilePath, string StageName, DataTable BaseDtTable, string TableName, string Delimiter)
        {
            string ColumnSelects = "";

            foreach (DataColumn DataColumn in BaseDtTable.Columns)
            {
                string ColName = DataColumn.ColumnName;
                ColumnSelects += "\"" + ColName + "\" VARCHAR,"; //assums all columns will be VARCHAR
            }
            ColumnSelects = ColumnSelects.Substring(0, ColumnSelects.Length - 1); //remove last comma

            //https://docs.snowflake.com/en/sql-reference/sql/create-file-format
            string FileFormatName = "TEMP_FILE_FORMAT";
            string FileFormatQuery = " CREATE OR REPLACE FILE FORMAT " + FileFormatName + " ";
            FileFormatQuery += " type = 'CSV' "; //works for all delimiter types, not just comma
            FileFormatQuery += " field_delimiter = '" + Delimiter + "' ";
            FileFormatQuery += " skip_header=1; ";
            this.Execute(FileFormatQuery);

            string CreateTableFromStagedFileQuery = " CREATE OR REPLACE FILE FORMAT " + TableName + " ";
            CreateTableFromStagedFileQuery += " SELECT " + ColumnSelects + " ";
            CreateTableFromStagedFileQuery += " FROM @~/" + StageName + " ";
            CreateTableFromStagedFileQuery += " (file_format => " + FileFormatName + "); ";
            this.Execute(CreateTableFromStagedFileQuery);

            string RemoveStagingAreaQuery = "REMOVE @~/" + StageName + "/; ";
            this.Execute(RemoveStagingAreaQuery);

            string DropFileFormatQuery = "DROP FILE FORMAT " + FileFormatName + ";";
            this.Execute(DropFileFormatQuery);
        }

        public void Close()
        {
            this.Conn.Close();
        }

        //https://community.snowflake.com/s/article/How-to-connect-to-snowflake-using-C-Sharp-application-with-snowflake-NET-Connector-to-perform-SQL-operations-in-windows
        //https://community.snowflake.com/s/article/Connect-to-Snowflake-from-Visual-Studio-using-the-NET-Connector

        public static void SnowflakeTest()
        {
            using (IDbConnection conn = new SnowflakeDbConnection())
            {
                
                Console.WriteLine(1);
                // Set Connection String
                //string ConnectionString = "account=UKA60997;user=OSCAR;password=Riguadon74!;Warehouse=mywarehouse;db=DATA_ENGINEERING;schema=public;role=SYSADMIN;warehouse=LOW_PRIORITY;host=UKA60997.us-west-2.snowflakecomputing.com";
                //string ConnectionString = "Account=BTA87963;User=OSCAR;Password=Riguadon74!;Warehouse=DATA_ENGINEERING;Database=DATA_ENGINEERING;Schema=public;Role=SYSADMIN;host=BTA87963.us-west-2.snowflakecomputing.com";
                //string ConnectionString = "account=uka60997;user=oscar;password=Riguadon74!;ROLE=SYSADMIN;db=DATA_ENGINEERING;schema=public";
                //string ConnectionString = "scheme=https;account=uka60997;host=uka60997.snowflakecomputing.com;port=443;role=SYSADMIN;warehouse=DATA_ENGINEERING;user=oscar;password=Riguadon74!;";
                //https://github.com/snowflakedb/snowflake-connector-net/blob/master/doc/Connecting.md
                string ConnectionString = "account=uka60997;password=Riguadon74!;user=oscar;";

                //"port=443;role=SYSADMIN;warehouse=DATA_ENGINEERING;user=oscar;authenticator=snowflake";
                //account  UKA60997  select CURRENT_ACCOUNT();
                //region  AWS_US_WEST_2  select cURRENT_REGION();
                //"account=BTA87963;user=OSCAR;password=Riguadon74!;db=DATA_ENGINEERING;schema=public;role=SYSADMIN;warehouse=LOW_PRIORITY;host=BTA87963.us-west-2.snowflakecomputing.com";
                Console.WriteLine(2);
                conn.ConnectionString = ConnectionString;
                Console.WriteLine(ConnectionString);

                // Initiate the connection
                conn.Open();
                Console.WriteLine(3);

                IDbCommand cmd = conn.CreateCommand();
                Console.WriteLine(4);
                cmd.CommandText = " SELECT TOP 10 * FROM DATA_ENGINEERING.PUBLIC.CONSUMER ";//" select current_user();"; // Set up query
                Console.WriteLine(5);
                IDataReader reader = cmd.ExecuteReader();
                Console.WriteLine(6);

                while (reader.Read())
                {
                    Console.WriteLine(reader.GetString(0)); // Display query result in the console
                }
                Console.WriteLine(7);

                string x = Console.ReadLine();
                cmd.CommandText = " SELECT TOP 20 * FROM DATA_ENGINEERING.PUBLIC.OKTEST ";//" select current_user();"; // Set up query
                Console.WriteLine(5);
                IDataReader reader2 = cmd.ExecuteReader();
                Console.WriteLine(6);
                while (reader2.Read())
                {
                    Console.WriteLine(reader2.GetString(0)); // Display query result in the console
                }
                Console.WriteLine(7);
                conn.Close(); // Close the connection
                string x2 = Console.ReadLine();
                Console.WriteLine(8);
            }
        }
    }
}

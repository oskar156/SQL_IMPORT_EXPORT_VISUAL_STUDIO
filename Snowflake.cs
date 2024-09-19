using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Snowflake.Data.Client;
/*
 * https://github.com/snowflakedb/snowflake-connector-net/issues/895
 * in your project, open NuGet package manager console (in VisualStudio it's Tools > Nuget Package Manager > Package Manager console), 
 * then after Powershell is loaded, issue this command to install Mono.Unix:
PM> NuGet\Install-Package Mono.Unix -Version 7.1.0-final.1.21458.1
 * also installed snowflake w/nuget
 * 
 * https://community.snowflake.com/s/article/How-to-connect-to-snowflake-using-C-Sharp-application-with-snowflake-NET-Connector-to-perform-SQL-operations-in-windows
 */

namespace SQL_SERVER_IMPORT_EXPORT
{
    public class Snowflake
    {
        //https://community.snowflake.com/s/article/How-to-connect-to-snowflake-using-C-Sharp-application-with-snowflake-NET-Connector-to-perform-SQL-operations-in-windows
        //https://community.snowflake.com/s/article/Connect-to-Snowflake-from-Visual-Studio-using-the-NET-Connector

        public static void SnowflakeTest()
        {
            using (IDbConnection conn = new SnowflakeDbConnection())
            {
                
                Console.WriteLine(1);
                // Set Connection String
                string ConnectionString = ""
                conn.ConnectionString = ConnectionString;
                Console.WriteLine(ConnectionString);

                // Initiate the connection
                conn.Open();
                Console.WriteLine(3);

                IDbCommand cmd = conn.CreateCommand();
                Console.WriteLine(4);
                cmd.CommandText = " select current_user();"; // Set up query
                Console.WriteLine(5);
                IDataReader reader = cmd.ExecuteReader();
                Console.WriteLine(6);

                while (reader.Read())
                {
                    Console.WriteLine(reader.GetString(0)); // Display query result in the console
                }
                Console.WriteLine(7);

                conn.Close(); // Close the connection
                Console.WriteLine(8);
            }
        }
    }
}

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
                //string ConnectionString = "account=UKA60997;user=OSCAR;password=Riguadon74!;Warehouse=mywarehouse;db=DATA_ENGINEERING;schema=public;role=SYSADMIN;warehouse=LOW_PRIORITY;host=UKA60997.us-west-2.snowflakecomputing.com";
                string ConnectionString = "Account=BTA87963;User=OSCAR;Password=Riguadon74!;Warehouse=DATA_ENGINEERING;Database=DATA_ENGINEERING;Schema=public;Role=SYSADMIN;host=BTA87963.us-west-2.snowflakecomputing.com";

                //account  UKA60997  select CURRENT_ACCOUNT();
                //region  AWS_US_WEST_2  select cURRENT_REGION();
                //"account=BTA87963;user=OSCAR;password=Riguadon74!;db=DATA_ENGINEERING;schema=public;role=SYSADMIN;warehouse=LOW_PRIORITY;host=BTA87963.us-west-2.snowflakecomputing.com";
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

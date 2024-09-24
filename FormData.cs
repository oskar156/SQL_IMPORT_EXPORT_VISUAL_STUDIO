using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SQL_SERVER_IMPORT_EXPORT
{
    public class FormData
    {
        public string[] Environs = { "SNOWFLAKE", "SQL SERVER" };
        public string DefaultEnviron = "SQL SERVER";

        public string[] SqlServerServers = { "SQL04","SQL05" };
        public string[] SqlServerDatabases = { "BMW", "TEMP_BMW", "TEMP_EA", "TEMP_J","TEMP_JC","TEMP_NS", "TEMP_OK", "TEMP_TN" };
        public string DefaultSqlServerServer = "SQL04";
        public string[] SqlServerImportTypes = { "csv", "txt", "xls*" };
        public string[] SqlServerImportDelims = { "COMMA", "PIPE", "TAB", "FIXED WIDTH" };

        public string[] SnowflakeDatabases = { "DATA_ENGINEERING" };
        public string[] SnowflakeSchemas = { "PUBLIC" };
        public string[] SnowflakeUsernames = { "oscar" };
        public string[] SnowflakeAccounts = { "uka60997" };        
        public string DefaultSnowflakeDatabase = "DATA_ENGINEERING";
        public string DefaultSnowflakeSchema = "PUBLIC";
        public string DefaultSnowflakeAccount = "uka60997";
        public string[] SnowflakeImportTypes = { "csv", "txt" };
        public string[] SnowflakeImportDelims = { "COMMA", "PIPE", "TAB"};




        //at some point move in all info here from Form1() {}

        public FormData() { }
    }
}

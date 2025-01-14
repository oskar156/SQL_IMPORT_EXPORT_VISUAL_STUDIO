/* TABLE OF CONTENTS
 * 
 * METHODS
 * public System.Type Get(string DataTypeChoice)
 * 
 */

namespace SQL_SERVER_IMPORT_EXPORT
{
    public class DataTypes
    {
        //======================================================================================================================
        // METHODS
        //======================================================================================================================
        public System.Type Get(string DataTypeChoice)
        {
            switch (DataTypeChoice)
            {
                case "INT": return System.Type.GetType("System.Int32");
                case "VARCHAR": return System.Type.GetType("System.String");
                case "STRING": return System.Type.GetType("System.String");
                case "DECIMAL": return System.Type.GetType("System.Decimal");
                case "BOOL": return System.Type.GetType("System.Boolean");
                case "BOOLEAN": return System.Type.GetType("System.Boolean");
                case "TIMESPAN": return System.Type.GetType("System.TimeSpan");
                case "DATETIME": return System.Type.GetType("System.DateTime");
                case "BYTEARRAY": return System.Type.GetType("System.Byte[]");
                default: return System.Type.GetType("System.String");
            }
        }
    }
}

/* TABLE OF CONTENTS
 * 
 * METHODS
 * public System.Text.Encoding Get(string EncodingChoice)
 * public string GetSnowflake(string EncodingChoice)
 * 
 */

using System.Text;

namespace SQL_SERVER_IMPORT_EXPORT
{
    public class Encoding
    {
        //======================================================================================================================
        // METHODS
        //======================================================================================================================
        public System.Text.Encoding Get(string EncodingChoice)
        {
            //https://stackoverflow.com/questions/1432064/how-to-read-an-ansi-encoded-file-containing-special-characters
            switch (EncodingChoice)
            {
                case "Default": return System.Text.Encoding.Default; //closest to ANSI (idk if identical)
                case "ANSI": return System.Text.Encoding.GetEncoding("Windows-1252", EncoderFallback.ReplacementFallback, DecoderFallback.ReplacementFallback);
                case "ASCII": return System.Text.Encoding.ASCII;
                case "UTF8": return System.Text.Encoding.UTF8;
                case "UTF16": return System.Text.Encoding.Unicode;
                case "UTF32": return System.Text.Encoding.UTF32;
                default: return System.Text.Encoding.UTF8;
            }
        }
        public string GetSnowflake(string EncodingChoice)
        {
            //https://stackoverflow.com/questions/1432064/how-to-read-an-ansi-encoded-file-containing-special-characters
            switch (EncodingChoice)
            {
                case "Default": return "UTF8";
                case "ANSI": return "WINDOWS1252";
                case "UTF8": return "UTF8";
                case "UTF16": return "UTF16";
                case "UTF32": return "UTF32";
                default: return "UTF8";
            }
        }
    }
}

/* TABLE OF CONTENTS
 * 
 * METHODS
 * public void StripNonPrintableCharsFromValue(ref string Value)
 * public void LimitValueLength(ref string Value, ref int FieldLimit)
 * public Tuple<string, int> ParseColumnWidthLine(string line)
 * public Tuple<string, string> ParseColumnTypeLine(string line)
 * public static string GetColumnName(string cellReference)
 * public static int? GetColumnIndexFromName(string columnName)
 * 
*/

using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.RegularExpressions;

namespace SQL_SERVER_IMPORT_EXPORT
{
    public class Helpers
    {
        //======================================================================================================================
        // METHODS
        //======================================================================================================================
        public void StripNonPrintableCharsFromValue(ref string Value)
        {
            Value = Regex.Replace(Value, @"[^\u0020-\u007E]+", "");
        }
        public void LimitValueLength(ref string Value, ref int FieldLimit)
        {
            //limit field length if necessary
            if (FieldLimit >= 1 && Value.Length >= FieldLimit)
            {
                Value = Value.Substring(0, FieldLimit);
            }
        }
        public Tuple<string, int> ParseColumnWidthLine(string line)
        {
            string ColumnName = "";
            int ColumnLength = 0;
            string LineTrimmed = line.Trim();

            if (LineTrimmed[0] == '[')
            {
                Regex re = new Regex("\\[.*\\]");
                string ColumnNameRaw = re.Match(line).ToString();
                string ColumnLengthRaw = LineTrimmed.Substring(ColumnNameRaw.Length).Trim();

                ColumnName = ColumnNameRaw.Substring(1, ColumnNameRaw.Length - 2).Trim();
                ColumnLength = Int32.Parse(ColumnLengthRaw);
            }
            else
            {
                int LastSpaceIndex = LineTrimmed.LastIndexOf(" ");

                if (LastSpaceIndex >= 0)
                {
                    ColumnName = LineTrimmed.Substring(0, LastSpaceIndex);
                    ColumnLength = Int32.Parse(LineTrimmed.Substring(LastSpaceIndex).Trim());
                }
                else
                {
                    ColumnName = "";
                    ColumnLength = Int32.Parse(LineTrimmed);
                }
            }

            return new Tuple<string, int>(ColumnName, ColumnLength);
        }
        public Tuple<string, string> ParseColumnTypeLine(string line)
        {
            string ColumnName = "";
            string ColumnType = "";
            string LineTrimmed = line.Trim();

            if (LineTrimmed[0] == '[')
            {
                Regex re = new Regex("\\[.*\\]");
                string ColumnNameRaw = re.Match(line).ToString();

                ColumnName = ColumnNameRaw.Substring(1, ColumnNameRaw.Length - 2).Trim();
                ColumnType = LineTrimmed.Substring(ColumnNameRaw.Length).Trim();
            }
            else
            {
                int LastSpaceIndex = LineTrimmed.LastIndexOf(" ");

                if (LastSpaceIndex >= 0)
                {
                    ColumnName = LineTrimmed.Substring(0, LastSpaceIndex);
                    ColumnType = LineTrimmed.Substring(LastSpaceIndex).Trim();
                }
                else
                {
                    ColumnName = "";
                    ColumnType = LineTrimmed;
                }
            }

            return new Tuple<string, string>(ColumnName, ColumnType);
        }
        public static string GetColumnName(string cellReference)
        {
            // Create a regular expression to match the column name portion of the cell name.
            Regex regex = new Regex("[A-Za-z]+");
            Match match = regex.Match(cellReference);

            return match.Value;
        }
        public static int? GetColumnIndexFromName(string columnName)
        {
            const string Letters = "ZABCDEFGHIJKLMNOPQRSTUVWXY";
            int? columnIndex = null;

            string[] colLetters = Regex.Split(columnName, "([A-Z]+)");
            colLetters = colLetters.Where(s => !string.IsNullOrEmpty(s)).ToArray();

            if (colLetters.Count() <= 2)
            {
                int index = 0;
                foreach (string col in colLetters)
                {
                    List<char> col1 = colLetters.ElementAt(index).ToCharArray().ToList();
                    int? indexValue = Letters.IndexOf(col1.ElementAt(index));

                    if (indexValue != -1)
                    {
                        // The first letter of a two digit column needs some extra calculations
                        if (index == 0 && colLetters.Count() == 2)
                        {
                            columnIndex = columnIndex == null ? (indexValue + 1) * 26 : columnIndex + ((indexValue + 1) * 26);
                        }
                        else
                        {
                            columnIndex = columnIndex == null ? indexValue : columnIndex + indexValue;
                        }
                    }

                    index++;
                }
            }
            return columnIndex;

        }
    }
}


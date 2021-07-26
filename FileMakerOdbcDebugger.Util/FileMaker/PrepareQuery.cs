using System.Text.RegularExpressions;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class FileMaker
    {
        /// <summary>
        /// Prepare a SQL query for Filemaker ODBC driver to execute.
        /// </summary>
        public static string PrepareQuery(string sqlQuery)
        {
            // Remove any comments:
            sqlQuery = Regex.Replace(sqlQuery, "(--.+)", "", RegexOptions.IgnoreCase); // inline comments
            sqlQuery = Regex.Replace(sqlQuery, @"\/\*[\s\S]*?\*\/", "", RegexOptions.IgnoreCase); // multiline comments

            var split = new Sql.SplitQuery(sqlQuery);

            // Loop over Statements.
            // Clean spaces (optional really, so they print nicely in logs).
            for (int i = 0; i <= split.Parts.Count - 1; i += 2)
            {
                var s = split.Parts[i];

                // Replace double spaces with a single space:
                while (s.Contains("  "))
                {
                    split.Parts[i] = s.Replace("  ", " ");
                }
                    
                // Remove single white spaces after a new line.
                // Note: we preserve new lines so its still easy to read when debugging.
                while (s.Contains("\r\n "))
                {
                    split.Parts[i] = s.Replace("\r\n ", "\r\n");
                }
            }

            // Loop over String Literals.
            // Replace all newline characters with CR (FileMaker only uses CR for newlines).
            for (int i = 1; i <= split.Parts.Count - 1; i += 2)
            {
                var s = split.Parts[i];
                split.Parts[i] = s.Replace("\r\n", "\r");
                split.Parts[i] = s.Replace("\n", "\r");
            }

            // Loop over Statements.
            // Surround any field names that start with an underscore (_) with double quotes (").
            for (int i = 0; i <= split.Parts.Count - 1; i += 2)
            {
                var s = split.Parts[i];
                split.Parts[i] = EscapeIdentifiersStartingWithUnderscore(s);
            }

            return split.GetJoined().Trim();
        }
    }
}

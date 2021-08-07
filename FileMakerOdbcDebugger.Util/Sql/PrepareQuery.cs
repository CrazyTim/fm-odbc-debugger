using System.Text.RegularExpressions;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Sql
    {
        private static readonly Regex inlineComments = new Regex(@"(--.+)", RegexOptions.IgnoreCase);
        private static readonly Regex multilineComments = new Regex(@"\/\*[\s\S]*?\*\/", RegexOptions.IgnoreCase);

        /// <summary>
        /// Prepare a SQL query for Filemaker ODBC driver to execute.
        /// </summary>
        public static string PrepareQuery(string sqlQuery, bool forFileMaker)
        {
            // Remove any comments:
            sqlQuery = inlineComments.Replace(sqlQuery, "");
            sqlQuery = multilineComments.Replace(sqlQuery, "");

            var split = new SplitQuery(sqlQuery);

            // Clean spaces (just so they print nicely in logs).
            foreach (var s in split.Parts)
            {
                if (s.Type == SqlPartType.Other)
                {
                    // Replace double spaces with a single space:
                    while (s.Value.Contains("  "))
                    {
                        s.Value = s.Value.Replace("  ", " ");
                    }

                    // Remove single white spaces after a new line.
                    // Note: we preserve new lines so its still easy to read when debugging.
                    s.Value.Replace("\r\n ", "\r\n");
                }
            }

            if (forFileMaker)
            {
                foreach (var s in split.Parts)
                {
                    if (s.Type == SqlPartType.StringLiteral)
                    {
                        // Replace all newline characters with CR (FileMaker only uses CR for newlines).
                        while (s.Value.Contains("\r\n"))
                        {
                            s.Value = s.Value.Replace("\r\n", "\r");
                        }
                        s.Value.Replace("\n", "\r");
                    }
                    if (s.Type == SqlPartType.Other)
                    {
                        s.Value = EscapeIdentifiersStartingWithUnderscore(s.Value);
                    }
                }
            }

            return split.Join().Trim();
        }
    }
}

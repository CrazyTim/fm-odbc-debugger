using System;
using System.Text.RegularExpressions;

namespace FileMakerOdbcDebugger.Util {

    public static partial class FileMaker {

        /// <summary>
        /// Return the file version of the driver 'fmodbc64.dll' located in the System32 folder.
        /// Return an empty string if the file doesn't exist.
        /// </summary>
        public static string GetOdbcDriverVersion64Bit() {

            var systemDrive = Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Machine);
            var pathToOdbcDriver = systemDrive + @"\System32\fmodbc64.dll";

            if (System.IO.File.Exists(pathToOdbcDriver)) {

                var myFileVersionInfo = System.Diagnostics.FileVersionInfo.GetVersionInfo(pathToOdbcDriver);
                return myFileVersionInfo.FileVersion;

            }

            return "";

        }

        /// <summary>
        /// Insert double-quotes around identifiyers starting with "_" (which is illegal in odbc, but not for FileMaker).
        /// </summary>
        public static string EscapeIdentifiyersStartingWithUnderscore(string sql) {

            if (string.IsNullOrEmpty(sql)) {
                return "";
            }
                
        
            bool flag = false;
            int startIndex = -1;
            int endIndex = -1;

            // Test first char:
            if (sql[0] == '_') {

                flag = true;
                startIndex = 0;

            }

            // Test middle chars:
            int i = 1;
            while (i <= sql.Length - 1) {

                if (!flag) {

                    // DETECT START OF WORD:
                    // The following characters signify the START of a word:
                    if (sql[i] == '_' & (sql[i - 1] == ' ' | sql[i - 1] == ',' | sql[i - 1] == '(' | sql[i - 1] == '.' | sql[i - 1] == '\t' | sql[i - 1] == '\r' | sql[i - 1] == '\n' | sql[i - 1] == '=')) {
                        flag = true;
                        startIndex = i;
                    }

                // DETECT END OF WORD:
                // The following characters signify the END of a word:
                } else if (sql[i] == ' ' | sql[i] == ',' | sql[i] == ')' | sql[i] == '.' | sql[i] == '\t' | sql[i] == '\r' | sql[i] == '\n' | sql[i] == '=') {
                    endIndex = i;
                    sql = sql.Insert(endIndex, "\"");
                    sql = sql.Insert(startIndex, "\"");
                    i += 2;
                    flag = false;
                }

                i += 1;

            }

            // Test last char:
            if (flag) {

                endIndex = sql.Length;
                sql = sql.Insert(endIndex, "\"");
                sql = sql.Insert(startIndex, "\"");

            }

            return sql;

        }

        /// <summary>
        /// Prepare a query for Filemaker ODBC driver to execute.
        /// </summary>
        public static string PrepareQuery(string Query) {

            // Remove any comments
            Query = Regex.Replace(Query, "(--.+)", "", RegexOptions.IgnoreCase); // inline comments
            Query = Regex.Replace(Query, @"\/\*[\s\S]*?\*\/", "", RegexOptions.IgnoreCase); // multiline comments

            // Split
            var split = Util.Sql.SplitQueryByStringLiteral(Query);

            // Loop over Statements...
            // Clean spaces (optional really, so they print nicely in logs):
            for (int i = 0; i <= split.Parts.Count - 1; i += 2) {

                var s = split.Parts[i];

                // Replace double spaces with a single space:
                while (s.Contains("  "))
                    split.Parts[i] = s.Replace("  ", " ");

                // Remove single white spaces after a new line:
                // Note: we preserve new lines so its still easy to read when debugging.
                while (s.Contains("\r\n "))
                    split.Parts[i] = s.Replace("\r\n ", "\r\n");

            }

            // Loop over String Literals...
            // Replace all newline characters with CR (FileMaker only uses CR for newlines):
            for (int i = 1; i <= split.Parts.Count - 1; i += 2) {

                var s = split.Parts[i];
                split.Parts[i] = s.Replace("\r\n", "\r");
                split.Parts[i] = s.Replace("\n", "\r");

            }

            // Loop over Statements...
            // Surround any field names that start with an underscore (_) with double quotes ("):
            for (int i = 0; i <= split.Parts.Count - 1; i += 2) {

                var s = split.Parts[i];
                split.Parts[i] = EscapeIdentifiyersStartingWithUnderscore(s);

            }

            return split.Rejoin().Trim();

        }

    }

}

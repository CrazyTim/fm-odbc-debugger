namespace FileMakerOdbcDebugger.Util
{
    public static partial class FileMaker
    {
        /// <summary>
        /// Insert double-quotes around identifiyers starting with "_" (which is illegal in odbc, but not for FileMaker).
        /// </summary>
        public static string EscapeIdentifiersStartingWithUnderscore(string sql)
        {
            if (string.IsNullOrEmpty(sql)) return "";

            bool flag = false;
            int startIndex = -1;
            int endIndex;

            // Test first char:
            if (sql[0] == '_')
            {
                flag = true;
                startIndex = 0;
            }

            // Test middle chars:
            int i = 1;
            while (i <= sql.Length - 1)
            {
                if (!flag)
                {
                    // DETECT START OF WORD:
                    // The following characters signify the START of a word:
                    if (sql[i] == '_' & (sql[i - 1] == ' ' | sql[i - 1] == ',' | sql[i - 1] == '(' | sql[i - 1] == '.' | sql[i - 1] == '\t' | sql[i - 1] == '\r' | sql[i - 1] == '\n' | sql[i - 1] == '='))
                    {
                        flag = true;
                        startIndex = i;
                    }
                }
                // DETECT END OF WORD:
                // The following characters signify the END of a word:
                else if (sql[i] == ' ' | sql[i] == ',' | sql[i] == ')' | sql[i] == '.' | sql[i] == '\t' | sql[i] == '\r' | sql[i] == '\n' | sql[i] == '=')
                {
                    endIndex = i;
                    sql = sql.Insert(endIndex, "\"");
                    sql = sql.Insert(startIndex, "\"");
                    i += 2;
                    flag = false;
                }
                i += 1;
            }

            // Test last char:
            if (flag)
            {
                endIndex = sql.Length;
                sql = sql.Insert(endIndex, "\"");
                sql = sql.Insert(startIndex, "\"");
            }

            return sql;
        }
    }
}

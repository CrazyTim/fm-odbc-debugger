using System.Linq;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Sql
    {
        public static readonly char[] IdentifierStartCharacters = { ' ', ',', '(', '.', '\t', '\r', '\n', '=' };
        public static readonly char[] IdentifierEndCharacters = { ' ', ',', ')', '.', '\t', '\r', '\n', '=' };

        /// <summary>
        /// Insert double-quotes around identifiyers starting with "_" (which is illegal in odbc, but not for FileMaker).
        /// In FileMaker its common to name fields starting with a underscore (_).
        /// </summary>
        public static string EscapeIdentifiersStartingWithUnderscore(string sqlStatement)
        {
            if (string.IsNullOrEmpty(sqlStatement)) return "";

            bool isWordFound = false;
            int startIndexOfWord = -1;
            int endIndexOfWord;

            // Test first char:
            if (sqlStatement[0] == '_')
            {
                isWordFound = true;
                startIndexOfWord = 0;
            }

            // Test middle chars:
            int i = 1;
            while (i <= sqlStatement.Length - 1)
            {
                var thisChar = sqlStatement[i];
                var prevChar = sqlStatement[i - 1];

                if (!isWordFound)
                {
                    if (thisChar == '_' & IdentifierStartCharacters.Contains(prevChar))
                    {
                        isWordFound = true;
                        startIndexOfWord = i;
                    }
                }
                else if (IdentifierEndCharacters.Contains(thisChar))
                {
                    endIndexOfWord = i;
                    sqlStatement = sqlStatement.Insert(endIndexOfWord, "\"");
                    sqlStatement = sqlStatement.Insert(startIndexOfWord, "\"");
                    i += 2;
                    isWordFound = false;
                }
                i += 1;
            }

            // Test last char:
            if (isWordFound)
            {
                endIndexOfWord = sqlStatement.Length;
                sqlStatement = sqlStatement.Insert(endIndexOfWord, "\"");
                sqlStatement = sqlStatement.Insert(startIndexOfWord, "\"");
            }

            return sqlStatement;
        }
    }
}

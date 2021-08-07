using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class SqlQuery
    {
        private static readonly char tempStatementSeparator = '\u009f'; // &#159

        /// <summary>
        /// Split a SQL query into a List of statements based on any "statement terminators" (";").
        /// </summary>
        public static List<string> SplitIntoStatements(string sqlQuery)
        {
            int singleQuoteCount = 0;
            var builder = new StringBuilder(sqlQuery);

            for (int i = 0, loopTo = builder.Length - 1; i <= loopTo; i++)
            {
                char s = builder[i];

                if (s == '\'')
                {
                    singleQuoteCount = (singleQuoteCount == 1) ? 0 : 1;
                }
                else if (s == ';')
                {
                    if (singleQuoteCount == 0) builder[i] = tempStatementSeparator;
                }
            }

            var l = builder.ToString().Split(tempStatementSeparator).ToList();
            l.RemoveAll(s => string.IsNullOrWhiteSpace(s));

            return l;
        }
    }
}

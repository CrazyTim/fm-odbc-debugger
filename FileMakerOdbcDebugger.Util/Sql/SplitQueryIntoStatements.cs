using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Sql
    {
        private static readonly char tempSeparator = '\u009f'; // &#159

        /// <summary>
        /// Split a SQL query into an List of statements based on any "statement terminators" (";").
        /// </summary>
        public static List<string> SplitQueryIntoStatements(string sqlQuery)
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
                    if (singleQuoteCount == 0) builder[i] = tempSeparator;
                }
            }

            var l = builder.ToString().Split(tempSeparator).ToList();
            l.RemoveAll(s => string.IsNullOrWhiteSpace(s));

            return l;
        }
    }
}

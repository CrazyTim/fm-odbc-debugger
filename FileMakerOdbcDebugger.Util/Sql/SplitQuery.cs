using System.Collections.Generic;
using System.Linq;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Sql
    {
        /// <summary>
        /// Represents a SQL query that has been split by its string literals.
        /// Can be used to process string literals separate from everything else.
        /// For Example this query: "SELECT * FROM table WHERE field = 'test'"
        /// will be split into:
        /// 1. "SELECT * FROM table WHERE field = "
        /// 2. "test"
        /// </summary>
        public class SplitQuery
        {
            public List<string> Parts { get; set; }

            public SplitQuery(string sqlQuery)
            {
                Parts = new List<string>();

                var singleQuote = '\'';

                if (string.IsNullOrEmpty(sqlQuery)) return;

                // If query starts with a single quote, but its not an escape, add an empty string before it so string literals are always at even numbered indexes.
                if (sqlQuery[0] == singleQuote && sqlQuery[1] != singleQuote)
                {
                    Parts.Add("");
                }

                if (sqlQuery.Length >= 2)
                {
                    // Note: query has to be at least two chars long because we look ahead one char.

                    var i = 0;

                    while (i < sqlQuery.Length)
                    {
                        if ((sqlQuery[i]) == singleQuote)
                        {

                            if (i == sqlQuery.Length - 1)
                            {
                                // Trim last single quote.
                                sqlQuery = sqlQuery.TrimNCharsFromEnd(1);
                                continue;
                            }

                            if (sqlQuery[i + 1] == singleQuote)
                            {
                                // The first single quote was escaping the second.
                                i += 2;
                                continue; // Skip both. No need to split.
                            }

                            if (i == 0)
                            {
                                // Trim first single quote.
                                sqlQuery = sqlQuery.TrimNCharsFromStart(1);
                                continue;
                            }

                            Parts.Add(sqlQuery.Substring(0, i));
                            sqlQuery = sqlQuery.TrimNCharsFromStart(i + 1);
                            i = 0;
                        }

                        i += 1;
                    }
                }

                Parts.Add(sqlQuery); // Add remainder
            }

            public List<string> Statements
            {
                get { return GetParts(0); }
            }

            public List<string> StringLiterals
            {
                get { return GetParts(1); }
            }

            private List<string> GetParts(int partType)
            {
                var r = new List<string>();

                for (var i = partType; i < Parts.Count; i++)
                {
                    r.Add(Parts[i]);
                    i += 2;
                }

                return r;
            }

            /// <summary>
            /// Return the SQL query joined back together.
            /// </summary>
            public string Join()
            {
                if (Parts.Count == 0) return "";

                var joined = string.Join("'", Parts.ToArray<string>()) + "'";

                if (Parts.Count.IsOdd()) joined = joined.TrimNCharsFromEnd(1);

                return joined;
            }
        }
    }
}

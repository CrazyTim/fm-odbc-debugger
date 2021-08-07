using System.Collections.Generic;

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
            public List<SqlPart> Parts { get; set; } = new List<SqlPart>();

            public SplitQuery(string sqlQuery)
            {
                var tempParts = new List<string>();
                var singleQuote = '\'';

                if (string.IsNullOrEmpty(sqlQuery)) return;

                // If query starts with a single quote and its not an escape (even though this is not a valid query),
                // add an empty string before it so string literals are always at even numbered indexes.
                if (sqlQuery[0] == singleQuote && sqlQuery[1] != singleQuote)
                {
                    tempParts.Add("");
                }

                if (sqlQuery.Length >= 2) // Note: query must be at least two chars long because we look ahead one char.
                {
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

                            tempParts.Add(sqlQuery.Substring(0, i));
                            sqlQuery = sqlQuery.TrimNCharsFromStart(i + 1);
                            i = 0;
                        }

                        i += 1;
                    }
                }

                tempParts.Add(sqlQuery); // Add remainder

                // Create Parts
                for (int i = 0; i < tempParts.Count; i += 1)
                {
                    if (i == 0 && tempParts[i] == "") continue; // Skip empty first string.

                    Parts.Add(new SqlPart()
                    {
                        Type = (i + 1).IsOdd() ? SqlPartType.Other : SqlPartType.StringLiteral,
                        Value = tempParts[i],
                    });
                }
            }

            /// <summary>
            /// Return the SQL query joined back together.
            /// </summary>
            public string Join()
            {
                if (Parts.Count == 0) return "";
                return string.Join("", Parts);
            }
        }

        public class SqlPart
        {
            public SqlPartType Type { get; set; }
            public string Value { get; set; }

            public override string ToString()
            {
                if (Type == SqlPartType.StringLiteral)
                {
                    return "'" + Value + "'";
                }
                return Value;
            }
        }

        public enum SqlPartType
        {
            StringLiteral,
            Other,
        }
    }
}

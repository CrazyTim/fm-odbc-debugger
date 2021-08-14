using System.Collections.Generic;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class SqlQuery
    {
        /// <summary>
        /// Clean spaces so the query will print nicely in logs.
        /// </summary>
        public static List<SqlPart> CleanSpaces(this List<SqlPart> parts)
        {
            foreach (var s in parts)
            {
                if (s.Type == SqlPartType.Other)
                {
                    // Replace double spaces with a single space:
                    while (s.Value.Contains("  "))
                    {
                        s.Value = s.Value.Replace("  ", " ");
                    }

                    // Remove single white spaces after a new line:
                    s.Value.Replace("\r\n ", "\r\n");
                }
            }

            return parts;
        }
    }
}

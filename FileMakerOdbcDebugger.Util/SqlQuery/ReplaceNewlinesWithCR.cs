using System.Collections.Generic;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class SqlQuery
    {
        /// <summary>
        /// Replace all newline characters with CR (\r).
        /// Note: FileMaker only uses CR for newlines.
        /// </summary>
        public static void ReplaceNewlinesInStringLiteralsWithCR(List<SqlPart> parts)
        {
            foreach (var s in parts)
            {
                if (s.Type == SqlPartType.StringLiteral)
                {
                    while (s.Value.Contains("\r\n"))
                    {
                        s.Value = s.Value.Replace("\r\n", "\r");
                    }
                    s.Value.Replace("\n", "\r");
                }
            }
        }
    }
}

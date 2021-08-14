using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class SqlQuery
    {
        private static readonly Regex findCrlfAndLfLineBreakSequences = new Regex(@"\r\n?|\n");

        /// <summary>
        /// Replace CRLF and LF linebreak sequences with CR.
        /// Note: FileMaker only uses CR for newlines.
        /// </summary>
        public static void ReplaceLineBreaksInStringLiteralsWithCR(List<SqlPart> parts)
        {
            foreach (var s in parts)
            {
                if (s.Type == SqlPartType.StringLiteral)
                {
                    s.Value = findCrlfAndLfLineBreakSequences.Replace(s.Value, Constants.CR);
                }
            }
        }
    }
}

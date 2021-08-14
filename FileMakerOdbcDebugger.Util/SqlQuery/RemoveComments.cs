using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class SqlQuery
    {
        // Note: `$` only detects Line Feed, so we need a bit of magic for new lines.
        // Also need to use lazy matching with `.` or it will match to the end of the string.
        private static readonly Regex inlineComments = new Regex(@"(--.+?)(?=[\r\n]|\Z)");

        // TODO: add support for nested multiline comments
        private static readonly Regex multilineComments = new Regex(@"\/\*[\s\S]*?\*\/");

        /// <summary>
        /// Remove single and multiline comments
        /// </summary>
        public static void RemoveComments(List<SqlPart> parts)
        {
            foreach (var s in parts)
            {
                if (s.Type == SqlPartType.Other)
                {
                    s.Value = inlineComments.Replace(s.Value, "");
                    s.Value = multilineComments.Replace(s.Value, "");
                }
            }
        }
    }
}

using System.Collections.Generic;
using System.Text.RegularExpressions;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class SqlQuery
    {
        private static readonly Regex findIdentifiersStartingWithUnderscore = new Regex(@"(?<=[\s,""(\.=]|^)(_.*?)(?=[\s,"")\.=]|$)");

        /// <summary>
        /// Insert double-quotes around identifiyers starting with an underscore ("_") (which is illegal for OBDC, but not FileMaker).
        /// Note: in FileMaker its common to name fields starting with a underscore.
        /// </summary>
        public static void EscapeIdentifiersStartingWithUnderscore(List<SqlPart> parts)
        {
            foreach (var s in parts)
            {
                if (s.Type == SqlPartType.Other)
                {
                    s.Value = findIdentifiersStartingWithUnderscore.Replace(s.Value, m => '"' + m.ToString() + '"');
                }
            }
        }
    }
}

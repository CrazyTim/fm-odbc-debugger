using System.Text.RegularExpressions;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Sql
    {
        private static readonly Regex findIdentifiersStartingWithUnderscore = new Regex(@"(?<=[\s,""(\.=]|^)(_.*?)(?=[\s,"")\.=]|$)");

        /// <summary>
        /// Insert double-quotes around identifiyers starting with an underscore ("_") (which is illegal for OBDC, but not FileMaker).
        /// Note: in FileMaker its common to name fields starting with a underscore.
        /// </summary>
        public static string EscapeIdentifiersStartingWithUnderscore(string s)
        {
            return findIdentifiersStartingWithUnderscore.Replace(s, m => '"' + m.ToString() + '"');
        }
    }
}

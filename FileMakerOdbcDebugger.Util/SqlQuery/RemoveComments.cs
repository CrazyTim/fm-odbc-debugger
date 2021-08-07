using System.Text.RegularExpressions;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class SqlQuery
    {
        private static readonly Regex inlineComments = new Regex(@"(--.+)", RegexOptions.IgnoreCase);
        private static readonly Regex multilineComments = new Regex(@"\/\*[\s\S]*?\*\/", RegexOptions.IgnoreCase);

        /// <summary>
        /// Insert double-quotes around identifiyers starting with an underscore ("_") (which is illegal for OBDC, but not FileMaker).
        /// Note: in FileMaker its common to name fields starting with a underscore.
        /// </summary>
        public static string RemoveComments(string sqlQuery)
        {
            sqlQuery = inlineComments.Replace(sqlQuery, "");
            sqlQuery = multilineComments.Replace(sqlQuery, "");
            return sqlQuery;
        }
    }
}

using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Sql
    {
        public enum QueryIssue
        {
            [Description("Warning: FileMaker stores empty strings as NULL, so using \"WHERE column = ''\" or \"WHERE column <> ''\" on a \"Text\" field will always return 0 results. Use \"IS NULL\" instead.")]
            EmptyStringComparisonAlwaysReturns0Results,

            [Description("Syntax error: FileMaker ODBC does not support the keywords \"TRUE\" or \"FALSE\". Use \"1\" or \"0\" instead.")]
            TrueFalseKeywordNotSupported,

            [Description("Syntax error: one or more apostrophes (\"'\") need to be escaped.")]
            ApostrophesNotEscaped,
            
            [Description("Warning: FileMaker ODBC is very slow when comparing dates with the \"BETWEEN\" keyword. Use \">=\" and \"<=\" instead.")]
            BetweenKeywordIsVerySlow,
        }

        private static readonly Regex containsEmptyString1 = new Regex(@"WHERE[\s\S]*=[\s\S]*''", RegexOptions.IgnoreCase);
        private static readonly Regex containsEmptyString2 = new Regex(@"WHERE[\s\S]*<>[\s\S]*''", RegexOptions.IgnoreCase);
        private static readonly Regex containsTrue = new Regex(@"=\s*?TRUE", RegexOptions.IgnoreCase);
        private static readonly Regex containsFalse = new Regex(@"=\s*?FALSE", RegexOptions.IgnoreCase);
        private static readonly Regex containsBetween = new Regex(@"\b(" + "BETWEEN" + @")\b", RegexOptions.IgnoreCase);

        /// <summary>
        /// Check a SQL query for certain issues.
        /// </summary>
        public static List<QueryIssue> CheckQueryForIssues(string sqlQuery, bool forFileMaker)
        {
            var issues = new List<QueryIssue>();
            var split = new Sql.SplitQuery(sqlQuery);

            // Check if there is an odd number of single quotes (') in the query (syntax error):
            if (sqlQuery.Count(c => c == '/').IsOdd()) issues.Add(QueryIssue.ApostrophesNotEscaped);

            if (forFileMaker)
            {
                // Warn about empty string comparisons:
                foreach (var s in split.Statements)
                {
                    if (containsEmptyString1.Match(s).Success | containsEmptyString2.Match(s).Success)
                    {
                        issues.Add(QueryIssue.EmptyStringComparisonAlwaysReturns0Results);
                        break;
                    }
                }

                // Check for TRUE and FALSE keywords (syntax error):
                foreach (var s in split.Statements)
                {
                    if (containsTrue.Match(s).Success | containsFalse.Match(s).Success)
                    {
                        issues.Add(QueryIssue.TrueFalseKeywordNotSupported);
                        break;
                    }
                }

                // Warn about the BETWEEN statement:
                foreach (var s in split.Statements)
                {
                    if (containsBetween.Match(s).Success)
                    {
                        issues.Add(QueryIssue.BetweenKeywordIsVerySlow);
                        break;
                    }
                }
            }

            return issues;
        }
    }
}

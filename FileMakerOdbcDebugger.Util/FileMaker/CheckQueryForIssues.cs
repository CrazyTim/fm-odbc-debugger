using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class FileMaker
    {
        public enum QueryIssue
        {
            [Description("Warning: FileMaker stores empty strings as NULL, so using \"WHERE column = ''\" or \"WHERE column <> ''\" on a \"Text\" field will always return 0 results. Use \"IS NULL\" instead.")]
            EmptyStringComparisonAlwaysReturns0Results = 0,

            [Description("Syntax error: FileMaker ODBC does not support the keywords \"TRUE\" or \"FALSE\". Use \"1\" or \"0\" instead.")]
            TrueFalseKeywordNotSupported = 1,

            [Description("Syntax error: one or more apostrophes (\"'\") need to be escaped.")]
            ApostrophesNotEscaped = 2,

            [Description("Warning: query contains the keyword \"BETWEEN\" and FileMaker ODBC is VERY slow when comparing dates this way. Use \">=\" and \"<=\" instead.")]
            BetweenKeywordIsVerySlow = 3
        }

        private static Regex ContainsEmptyString1 = new Regex(@"WHERE[\s\S]*=[\s\S]*''", RegexOptions.IgnoreCase);
        private static Regex ContainsEmptyString2 = new Regex(@"WHERE[\s\S]*<>[\s\S]*''", RegexOptions.IgnoreCase);
        private static Regex ContainsTrue = new Regex(@"=\s*?TRUE", RegexOptions.IgnoreCase);
        private static Regex ContainsFalse = new Regex(@"=\s*?FALSE", RegexOptions.IgnoreCase);
        private static Regex ContainsBetween = new Regex(@"\b(" + "BETWEEN" + @")\b", RegexOptions.IgnoreCase);

        /// <summary>
        /// Check for certain errors not reported by the FileMaker ODBC driver.
        /// </summary>
        public static List<QueryIssue> CheckQueryForIssues(string sqlQuery)
        {
            var issues = new List<QueryIssue>();
            var split = new Sql.SplitQuery(sqlQuery);

            // Warn about empty string comparisons:
            foreach (var s in split.Statements)
            {
                if (ContainsEmptyString1.Match(s).Success | ContainsEmptyString2.Match(s).Success)
                {
                    issues.Add(QueryIssue.EmptyStringComparisonAlwaysReturns0Results);
                    break;
                }
            }

            // Check for TRUE and FALSE keywords (syntax error):
            foreach (var s in split.Statements)
            {
                if (ContainsTrue.Match(s).Success | ContainsFalse.Match(s).Success)
                {
                    issues.Add(QueryIssue.TrueFalseKeywordNotSupported);
                    break;
                }
            }

            // Check if there is an odd number of single quotes (') in the query (syntax error):
            if (sqlQuery.Count(c => c == '/').IsOdd()) issues.Add(QueryIssue.ApostrophesNotEscaped);

            // Warn about the BETWEEN statement:
            foreach (var s in split.Statements)
            {
                if (ContainsBetween.Match(s).Success)
                {
                    issues.Add(QueryIssue.BetweenKeywordIsVerySlow);
                    break;
                }
            }

            return issues;
        }
    }
}

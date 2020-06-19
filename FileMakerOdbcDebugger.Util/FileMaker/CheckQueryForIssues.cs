using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Text.RegularExpressions;

namespace FileMakerOdbcDebugger.Util {

    public static partial class FileMaker {

        public enum QueryIssue {
            [Description("Warning: FileMaker stores empty strings as NULL, so using \"WHERE column = ''\" or \"WHERE column <> ''\" on a \"Text\" field will always return 0 results. Use \"IS NULL\" instead.")]
            EmptyStringComparisonAlwaysReturns0Results = 0,

            [Description("Syntax error: FileMaker ODBC does not support the keywords \"TRUE\" or \"FALSE\". Use \"1\" or \"0\" instead.")]
            TrueFalseKeywordNotSupported = 1,

            [Description("Syntax error: one or more apostrophes (\"'\") need to be escaped.")]
            ApostrophesNotEscaped = 2,

            [Description("Warning: query contains the keyword \"BETWEEN\" and FileMaker ODBC is VERY slow when comparing dates this way. Use \">=\" and \"<=\" instead.")]
            BetweenKeywordIsVerySlow = 3
        }

        /// <summary> Check for certain errors not reported by the FileMaker ODBC driver. </summary>
        public static List<QueryIssue> CheckQueryForIssues(string Query) {

            var issues = new List<QueryIssue>();

            Regex r1;
            Regex r2;

            // Split
            var split = Util.Sql.SplitQueryByStringLiteral(Query);

            // Warn about empty string comparisons
            r1 = new Regex(@"WHERE[\s\S]*=[\s\S]*''", RegexOptions.IgnoreCase);
            r2 = new Regex(@"WHERE[\s\S]*<>[\s\S]*''", RegexOptions.IgnoreCase);
            foreach (var s in split.Statements) {
                if (r1.Match(s).Success | r2.Match(s).Success) {
                    issues.Add(QueryIssue.EmptyStringComparisonAlwaysReturns0Results);
                    break;
                }
            }

            // Check for TRUE and FALSE keywords (syntax error)
            r1 = new Regex(@"=\s*?TRUE", RegexOptions.IgnoreCase);
            r2 = new Regex(@"=\s*?FALSE", RegexOptions.IgnoreCase);
            foreach (var s in split.Statements) {
                if (r1.Match(s).Success | r2.Match(s).Success) {
                    issues.Add(QueryIssue.TrueFalseKeywordNotSupported);
                    break;
                }
            }

            // Check if there is an even number of single quotes (') in the query (syntax error).
            int singleQuoteCount = 0;
            for (int i = 0, loopTo = Query.Length - 1; i <= loopTo; i++) {
                if (Query[i] == '\'')
                    singleQuoteCount += 1;
            }

            if (singleQuoteCount.IsOdd()) {
                issues.Add(QueryIssue.ApostrophesNotEscaped);
            }

            // Warn about the BETWEEN statement
            r1 = new Regex(@"\b(" + "BETWEEN" + @")\b", RegexOptions.IgnoreCase);
            foreach (var s in split.Statements) {
                if (r1.Match(Query).Success) {
                    issues.Add(QueryIssue.BetweenKeywordIsVerySlow);
                    break;
                }
            }

            return issues;
        }

    }
}

using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text.RegularExpressions;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class SqlQuery
    {
        private static readonly string whereClause = @"(?<=WHERE\s.*)\w+\s+((<>)|=)\s*";
        private static readonly Regex containsEmptyStringComparison = new Regex(whereClause + "''", RegexOptions.IgnoreCase);
        private static readonly Regex containsTrueOrFalse = new Regex(whereClause + "((TRUE)|(FALSE))", RegexOptions.IgnoreCase);
        private static readonly Regex containsBetween = new Regex(@"\bBETWEEN\b", RegexOptions.IgnoreCase);

        /// <summary>
        /// Check a SQL query for certain issues.
        /// </summary>
        public static List<Issue> CheckForIssues(string sqlQuery, bool forFileMaker)
        {
            var issues = new List<Issue>();
            var split = new Split(sqlQuery);

            // Check if there is an odd number of single quotes (') in the query (syntax error):
            if (sqlQuery.Count(c => c == '\'').IsOdd())
            {
                issues.Add(Issue.ApostrophesNotEscaped);
                return issues; // No point in contining, other checks will be wrong until this is fixed.
            }

            if (forFileMaker)
            {
                // Warn about empty string comparisons:
                foreach (var i in split.Parts)
                {
                    if (i.Type == SqlPartType.Other &&
                        containsEmptyStringComparison.Match(i.Value).Success)
                    {
                        issues.Add(Issue.EmptyStringComparisonAlwaysReturnsZeroResults);
                        break;
                    }
                }

                // Check for TRUE and FALSE keywords (syntax error):
                foreach (var i in split.Parts)
                {
                    if (i.Type == SqlPartType.Other &&
                        containsTrueOrFalse.Match(i.Value).Success)
                    {
                        issues.Add(Issue.TrueFalseKeywordNotSupported);
                        break;
                    }
                }

                // Warn about the BETWEEN statement:
                foreach (var i in split.Parts)
                {
                    if (i.Type == SqlPartType.Other &&
                        containsBetween.Match(i.Value).Success)
                    {
                        issues.Add(Issue.BetweenKeywordIsVerySlow);
                        break;
                    }
                }
            }

            return issues;
        }

        public enum Issue
        {
            [Description("Warning: FileMaker stores empty strings as NULL, so using \"WHERE column = ''\" or \"WHERE column <> ''\" on a \"Text\" field will always return 0 results. Use \"IS NULL\" instead.")]
            EmptyStringComparisonAlwaysReturnsZeroResults,

            [Description("Syntax error: FileMaker ODBC does not support the keywords \"TRUE\" or \"FALSE\". Use \"1\" or \"0\" instead.")]
            TrueFalseKeywordNotSupported,

            [Description("Syntax error: one or more apostrophes (\"'\") need to be escaped.")]
            ApostrophesNotEscaped,

            [Description("Warning: FileMaker ODBC is very slow when comparing dates with the \"BETWEEN\" keyword. Use \">=\" and \"<=\" instead.")]
            BetweenKeywordIsVerySlow,
        }
    }
}

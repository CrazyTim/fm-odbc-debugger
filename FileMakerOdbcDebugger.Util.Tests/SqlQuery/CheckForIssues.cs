using Xunit;
using static FileMakerOdbcDebugger.Util.SqlQuery;

namespace FileMakerOdbcDebugger.Util.Tests
{
    public partial class SqlQuery
    {
        public class CheckForIssues
        {
            [Theory]
            [InlineData(Issue.EmptyStringComparisonAlwaysReturnsZeroResults, "WHERE column = ''")]
            [InlineData(Issue.EmptyStringComparisonAlwaysReturnsZeroResults, "WHERE column <> ''")]
            [InlineData(Issue.TrueFalseKeywordNotSupported, "WHERE column = TRUE")]
            [InlineData(Issue.TrueFalseKeywordNotSupported, "WHERE column = FALSE")]
            [InlineData(Issue.ApostrophesNotEscaped, "'''")]
            [InlineData(Issue.BetweenKeywordIsVerySlow, "WHERE column BETWEEN x")]
            public void Query_contains_issue(Issue expectedIssue, string sqlQuery)
            {
                // Act
                var result = Util.SqlQuery.CheckForIssues(sqlQuery, true);

                // Assert
                Assert.Contains(expectedIssue, result);
            }

            [Theory]
            [InlineData(Issue.EmptyStringComparisonAlwaysReturnsZeroResults, "")]
            [InlineData(Issue.TrueFalseKeywordNotSupported, "")]
            [InlineData(Issue.TrueFalseKeywordNotSupported, "SELECT '= TRUE'")]
            [InlineData(Issue.TrueFalseKeywordNotSupported, "SELECT '= FALSE'")]
            [InlineData(Issue.ApostrophesNotEscaped, "")]
            [InlineData(Issue.ApostrophesNotEscaped, "''")]
            [InlineData(Issue.BetweenKeywordIsVerySlow, "")]
            [InlineData(Issue.BetweenKeywordIsVerySlow, "SELECT 'WHERE column BETWEEN x'")]
            public void Query_does_not_contain_issue(Issue expectedIssue, string sqlQuery)
            {
                // Act
                var result = Util.SqlQuery.CheckForIssues(sqlQuery, true);

                // Assert
                Assert.DoesNotContain(expectedIssue, result);
            }
        }
    }
}

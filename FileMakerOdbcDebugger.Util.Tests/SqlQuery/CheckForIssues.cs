using Xunit;

namespace FileMakerOdbcDebugger.Util.Tests
{
    public partial class Sql
    {
        public class CheckForIssues
        {
            [Theory]
            [InlineData(SqlQuery.Issue.EmptyStringComparisonAlwaysReturnsZeroResults, "WHERE column = ''")]
            [InlineData(SqlQuery.Issue.EmptyStringComparisonAlwaysReturnsZeroResults, "WHERE column <> ''")]
            [InlineData(SqlQuery.Issue.TrueFalseKeywordNotSupported, "WHERE column = TRUE")]
            [InlineData(SqlQuery.Issue.TrueFalseKeywordNotSupported, "WHERE column = FALSE")]
            [InlineData(SqlQuery.Issue.ApostrophesNotEscaped, "'''")]
            [InlineData(SqlQuery.Issue.BetweenKeywordIsVerySlow, "WHERE column BETWEEN x")]
            public void Query_contains_issue(SqlQuery.Issue expectedIssue, string query)
            {
                // Act
                var result = SqlQuery.CheckForIssues(query, true);

                // Assert
                Assert.Contains(expectedIssue, result);
            }

            [Theory]
            [InlineData(SqlQuery.Issue.EmptyStringComparisonAlwaysReturnsZeroResults, "")]
            [InlineData(SqlQuery.Issue.TrueFalseKeywordNotSupported, "")]
            [InlineData(SqlQuery.Issue.TrueFalseKeywordNotSupported, "SELECT '= TRUE'")]
            [InlineData(SqlQuery.Issue.TrueFalseKeywordNotSupported, "SELECT '= FALSE'")]
            [InlineData(SqlQuery.Issue.ApostrophesNotEscaped, "")]
            [InlineData(SqlQuery.Issue.ApostrophesNotEscaped, "''")]
            [InlineData(SqlQuery.Issue.BetweenKeywordIsVerySlow, "")]
            [InlineData(SqlQuery.Issue.BetweenKeywordIsVerySlow, "SELECT 'WHERE column BETWEEN x'")]
            public void Query_does_not_contain_issue(SqlQuery.Issue expectedIssue, string query)
            {
                // Act
                var result = SqlQuery.CheckForIssues(query, true);

                // Assert
                Assert.DoesNotContain(expectedIssue, result);
            }
        }
    }
}

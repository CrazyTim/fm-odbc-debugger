using Xunit;

namespace FileMakerOdbcDebugger.Util.Tests
{
    public partial class Sql
    {
        public class CheckQueryForIssues
        {
            [Theory]
            [InlineData(Util.Sql.QueryIssue.EmptyStringComparisonAlwaysReturnsZeroResults, "WHERE column = ''")]
            [InlineData(Util.Sql.QueryIssue.EmptyStringComparisonAlwaysReturnsZeroResults, "WHERE column <> ''")]
            [InlineData(Util.Sql.QueryIssue.TrueFalseKeywordNotSupported, "WHERE column = TRUE")]
            [InlineData(Util.Sql.QueryIssue.TrueFalseKeywordNotSupported, "WHERE column = FALSE")]
            [InlineData(Util.Sql.QueryIssue.ApostrophesNotEscaped, "'''")]
            [InlineData(Util.Sql.QueryIssue.BetweenKeywordIsVerySlow, "WHERE column BETWEEN x")]
            public void Query_contains_issue(Util.Sql.QueryIssue expectedIssue, string query)
            {
                // Act
                var result = Util.Sql.CheckQueryForIssues(query, true);

                // Assert
                Assert.Contains(expectedIssue, result);
            }

            [Theory]
            [InlineData(Util.Sql.QueryIssue.EmptyStringComparisonAlwaysReturnsZeroResults, "")]
            [InlineData(Util.Sql.QueryIssue.TrueFalseKeywordNotSupported, "")]
            [InlineData(Util.Sql.QueryIssue.TrueFalseKeywordNotSupported, "SELECT '= TRUE'")]
            [InlineData(Util.Sql.QueryIssue.TrueFalseKeywordNotSupported, "SELECT '= FALSE'")]
            [InlineData(Util.Sql.QueryIssue.ApostrophesNotEscaped, "")]
            [InlineData(Util.Sql.QueryIssue.ApostrophesNotEscaped, "''")]
            [InlineData(Util.Sql.QueryIssue.BetweenKeywordIsVerySlow, "")]
            [InlineData(Util.Sql.QueryIssue.BetweenKeywordIsVerySlow, "SELECT 'WHERE column BETWEEN x'")]
            public void Query_does_not_contain_issue(Util.Sql.QueryIssue expectedIssue, string query)
            {
                // Act
                var result = Util.Sql.CheckQueryForIssues(query, true);

                // Assert
                Assert.DoesNotContain(expectedIssue, result);
            }
        }
    }
}

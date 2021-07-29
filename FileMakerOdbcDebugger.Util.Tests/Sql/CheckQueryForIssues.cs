using Xunit;

namespace FileMakerOdbcDebugger.Util.Tests
{
    public partial class Sql
    {
        public class CheckQueryForIssues
        {
            [Theory]
            [InlineData(Util.Sql.QueryIssue.EmptyStringComparisonAlwaysReturns0Results, "SELECT * FROM test WHERE column = ''")]
            [InlineData(Util.Sql.QueryIssue.EmptyStringComparisonAlwaysReturns0Results, "SELECT * FROM test WHERE column <> ''")]
            [InlineData(Util.Sql.QueryIssue.TrueFalseKeywordNotSupported, "SELECT * FROM test WHERE column = TRUE")]
            [InlineData(Util.Sql.QueryIssue.TrueFalseKeywordNotSupported, "SELECT * FROM test WHERE column = FALSE")]
            [InlineData(Util.Sql.QueryIssue.ApostrophesNotEscaped, "SELECT ''' FROM test")]
            [InlineData(Util.Sql.QueryIssue.BetweenKeywordIsVerySlow, "WHERE column BETWEEN x AND y")]
            public void Query_contains_issue(Util.Sql.QueryIssue expectedIssue, string query)
            {
                // Act
                var result = Util.Sql.CheckQueryForIssues(query, true);

                // Assert
                Assert.Contains(expectedIssue, result);
            }

            [Theory]
            [InlineData(Util.Sql.QueryIssue.TrueFalseKeywordNotSupported, "SELECT '= TRUE' FROM test")]
            [InlineData(Util.Sql.QueryIssue.TrueFalseKeywordNotSupported, "SELECT '= FALSE' FROM test")]
            [InlineData(Util.Sql.QueryIssue.ApostrophesNotEscaped, "SELECT '' FROM test")]
            [InlineData(Util.Sql.QueryIssue.BetweenKeywordIsVerySlow, "SELECT 'BETWEEN' FROM test")]
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

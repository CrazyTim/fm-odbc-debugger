using Xunit;

namespace FileMakerOdbcDebugger.Util.Tests
{
    public partial class Sql
    {
        public class SplitIntoStatements
        {
            [Theory]
            [InlineData(0, null)]
            [InlineData(0, "")]
            [InlineData(0, ";")] // ignore empty statements
            [InlineData(3, "SELECT;SELECT;SELECT")]
            [InlineData(2, "';';SELECT")] // Should ignore semicolon inside string literal
            public void SQL_query_is_split_into_the_correct_number_of_statements(int expectedStatementCount, string query)
            {
                // Act
                var result = SqlQuery.SplitIntoStatements(query);

                // Assert
                Assert.Equal(expectedStatementCount, result.Count);
            }
        }
    }
}

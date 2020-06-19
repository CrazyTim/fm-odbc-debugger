using Xunit;

namespace FileMakerOdbcDebugger.Util.Tests {

    public partial class Sql {

        public class SplitQueryIntoStatements {

            [Theory]
            [InlineData(0, null)]
            [InlineData(0, "")]
            [InlineData(0, ";")] // ignore empty statements
            [InlineData(3, "SELECT;SELECT;SELECT")]
            [InlineData(2, "';';SELECT")] // Should ignore semicolon inside string literal
            public void ReturnListWithCorrectCount(int expectedCount, string query) {

                // Act
                var actual = Util.Sql.SplitQueryIntoStatements(query);

                // Assert
                Assert.Equal(expectedCount, actual.Count);

            }

        }

    }

}

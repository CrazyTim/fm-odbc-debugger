using Xunit;

namespace FileMakerOdbcDebugger.Util.Tests {

    public partial class Sql {

        public class SplitQueryByStringLiteral {

            [Theory]
            [InlineData(0, "")]
            [InlineData(2, "'str'")] // Starts and ends with a string
            [InlineData(3, "_'str'_")] // String surrounded by other characters
            [InlineData(2, "'str''ing'")] // Escaped single quote inside string
            [InlineData(1, "_''_")] // Escaped single quote surrounded by other characters
            [InlineData(1, "''")] // Escaped single quote
            public void ReturnListWithCorrectCount_Split(int expectedCount, string query) {

                // Act
                var actual = Util.Sql.SplitQueryByStringLiteral(query);

                // Assert
                Assert.Equal(expectedCount, actual.Parts.Count);

            }

            [Theory]
            [InlineData("")]
            [InlineData("'str'")] // Starts and ends with a string
            [InlineData("_'str'_")] // String surrounded by other characters
            [InlineData("'str''ing'")] // Escaped single quote inside string
            [InlineData("_''_")] // Escaped single quote surrounded by other characters
            [InlineData("''")] // Escaped single quote
            public void ReturnListWithCorrectCount_Rejoin(string query) {

                // Act
                var actual = Util.Sql.SplitQueryByStringLiteral(query).Rejoin();

                // Assert
                Assert.Equal(query, actual);

            }

        }

    }

}

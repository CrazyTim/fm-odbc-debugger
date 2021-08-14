using Xunit;

namespace FileMakerOdbcDebugger.Util.Tests
{
    public partial class SqlQuery
    {
        public class Split
        {
            [Theory]
            [InlineData(0, "")]
            [InlineData(1, "'str'")] // Starts and ends with a string
            [InlineData(3, "_'str'_")] // String surrounded by other characters
            [InlineData(1, "'str''ing'")] // Escaped single quote at middle of string
            [InlineData(1, "_''_")] // Empty string surrounded by other characters
            [InlineData(1, "''")] // Empty string
            public void Query_is_split_into_correct_number_of_parts(int expectedPartCount, string sqlQuery)
            {
                // Act
                var result = new Util.SqlQuery.Split(sqlQuery).Parts;

                // Assert
                Assert.Equal(expectedPartCount, result.Count);
            }

            [Theory]
            [InlineData("")]
            [InlineData("'str'")] // Starts and ends with a string
            [InlineData("_'str'_")] // String surrounded by other characters
            [InlineData("'str''ing'")] // Escaped single quote at middle of string
            [InlineData("_''_")] // Empty string surrounded by other characters
            [InlineData("''")] // Empty string
            public void Query_is_unchanged_after_it_is_split_and_rejoined(string sqlQuery)
            {
                // Act
                var result = new Util.SqlQuery.Split(sqlQuery).Join();

                // Assert
                Assert.Equal(sqlQuery, result);
            }
        }
    }
}

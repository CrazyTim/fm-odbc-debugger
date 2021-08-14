using System.Collections.Generic;
using Xunit;
using static FileMakerOdbcDebugger.Util.SqlQuery;

namespace FileMakerOdbcDebugger.Util.Tests
{
    public partial class SqlQuery
    {
        public class EscapeIdentifiersStartingWithUnderscore
        {
            public static IEnumerable<object[]> NonWhiteSpaceTestData = new List<object[]>()
            {
                new object[] {
                    $"{Constants.CR}_test{Constants.CR}",
                    $"{Constants.CR}\"_test\"{Constants.CR}",
                },
                new object[] {
                    '\t' + "_test" + '\t',
                    '\t' + "\"_test\"" + '\t',
                },
                new object[] {
                    $"{Constants.LF}_test{Constants.LF}",
                    $"{Constants.LF}\"_test\"{Constants.LF}",
                },
            };

            [Theory]
            [InlineData("_test", "\"_test\"")]
            [InlineData(" _test ", " \"_test\" ")]
            [InlineData("_test _test", "\"_test\" \"_test\"")] // multiple words
            [InlineData("test_test", "test_test")] // underscore in the middle of a word
            [InlineData(" test_test ", " test_test ")]
            [MemberData(nameof(NonWhiteSpaceTestData))]
            public void Words_starting_with_an_underscore_are_escaped(string sqlPart, string expectedResult)
            {
                // Act
                var result = new List<SqlPart>() { new SqlPart() {
                    Type = SqlPartType.Other,
                    Value = sqlPart,
                }};

                result.EscapeIdentifiersStartingWithUnderscore();

                // Assert
                Assert.Equal(expectedResult, result[0].Value);
            }
        }
    }
}

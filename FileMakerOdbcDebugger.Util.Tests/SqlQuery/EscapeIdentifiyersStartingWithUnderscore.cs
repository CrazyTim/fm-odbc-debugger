using System.Collections.Generic;
using Xunit;
using static FileMakerOdbcDebugger.Util.SqlQuery;

namespace FileMakerOdbcDebugger.Util.Tests
{
    public partial class Sql
    {
        public class EscapeIdentifiyersStartingWithUnderscore
        {
            public static IEnumerable<object[]> Tests = new List<object[]>()
            {
                new object[] { '\r' + "_test" + '\r', '\r' + "\"_test\"" + '\r' },
                new object[] { '\t' + "_test" + '\t', '\t' + "\"_test\"" + '\t' },
                new object[] { '\n' + "_test" + '\n', '\n' + "\"_test\"" + '\n' },
            };

            [Theory]
            [InlineData("_test", "\"_test\"")]
            [InlineData(" _test ", " \"_test\" ")]
            [InlineData("_test _test", "\"_test\" \"_test\"")] // multiple words
            [InlineData("test_test", "test_test")] // underscore in the middle of a word
            [InlineData(" test_test ", " test_test ")]
            [MemberData(nameof(Tests))]
            public void Words_starting_with_an_underscore_are_escaped(string query, string expectedResult)
            {
                // Act
                var result = new List<SqlPart>() { new SqlPart() {
                    Type = SqlPartType.Other,
                    Value = query,
                }};

                EscapeIdentifiersStartingWithUnderscore(result);

                // Assert
                Assert.Equal(expectedResult, result[0].Value);
            }
        }
    }
}

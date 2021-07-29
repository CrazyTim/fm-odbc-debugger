using System.Collections.Generic;
using Xunit;

namespace FileMakerOdbcDebugger.Util.Tests
{
    public partial class Sql
    {
        public class EscapeIdentifiyersStartingWithUnderscore
        {
            public static IEnumerable<object[]> TestData = new List<object[]>()
            {
                new object[] { '\r' + "_test" + '\r', '\r' + "\"_test\"" + '\r' },
                new object[] { '\t' + "_test" + '\t', '\t' + "\"_test\"" + '\t' },
            };

            [Theory]
            [InlineData("_test", "\"_test\"")]
            [InlineData(" _test ", " \"_test\" ")]
            [InlineData("test_test", "test_test")]
            [InlineData(" test_test ", " test_test ")]
            [MemberData(nameof(TestData))]
            public void Identifiers_starting_with_an_underscore_are_escaped(string query, string expectedResult)
            {
                // Act
                var result = Util.Sql.EscapeIdentifiersStartingWithUnderscore(query);

                // Assert
                Assert.Equal(expectedResult, result);
            }
        }
    }
}

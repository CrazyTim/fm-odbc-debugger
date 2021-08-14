using System.Collections.Generic;
using Xunit;
using static FileMakerOdbcDebugger.Util.SqlQuery;

namespace FileMakerOdbcDebugger.Util.Tests
{
    public partial class SqlQuery
    {
        public class ReplaceLineBreaksInStringLiteralsWithCR
        {
            public static IEnumerable<object[]> InlineCommentTestData = new List<object[]>()
            {
                new object[] {
                    $"{Constants.CR}",
                    $"{Constants.CR}",
                },
                new object[] {
                    $"{Constants.LF}",
                    $"{Constants.CR}",
                },
                new object[] {
                    $"{Constants.CR}{Constants.LF}",
                    $"{Constants.CR}",
                },
                new object[] {
                    $"{Constants.CR}{Constants.LF}{Constants.CR}{Constants.LF}",
                    $"{Constants.CR}{Constants.CR}",
                },
            };

            [Theory]
            [MemberData(nameof(InlineCommentTestData))]
            public void All_line_breaks_are_replaced_with_CR(string sqlPart, string expectedResult)
            {
                // Act
                var result = new List<SqlPart>() { new SqlPart() {
                    Type = SqlPartType.StringLiteral,
                    Value = sqlPart,
                }};

                Util.SqlQuery.ReplaceLineBreaksInStringLiteralsWithCR(result);

                // Assert
                Assert.Equal(expectedResult, result[0].Value);
            }
        }
    }
}

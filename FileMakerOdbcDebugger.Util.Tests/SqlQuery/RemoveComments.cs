using System.Collections.Generic;
using Xunit;
using static FileMakerOdbcDebugger.Util.SqlQuery;

namespace FileMakerOdbcDebugger.Util.Tests
{
    public partial class Sql
    {
        public class RemoveComments
        {
            public static IEnumerable<object[]> InlineCommentTestData = new List<object[]>()
            {
                new object[] {
                    "--comment",
                    "",
                },
                new object[] {
                    $"--comment{Constants.CR}--comment",
                    $"{Constants.CR}",
                },
                new object[] {
                    $"--comment{Constants.LF}--comment",
                    $"{Constants.LF}",
                },
                new object[] {
                    $"--comment{Constants.CR}{Constants.LF}--comment",
                    $"{Constants.CR}{Constants.LF}",
                },
                new object[] {
                    $"{Constants.CR}--comment{Constants.CR}start--comment{Constants.CR}end--comment{Constants.CR}--comment{Constants.CR}",
                    $"{Constants.CR}{Constants.CR}start{Constants.CR}end{Constants.CR}{Constants.CR}",
                },
                new object[] {
                    $"{Constants.LF}--comment{Constants.LF}start--comment{Constants.LF}end--comment{Constants.LF}--comment{Constants.LF}",
                    $"{Constants.LF}{Constants.LF}start{Constants.LF}end{Constants.LF}{Constants.LF}",
                },
                new object[] {
                    $"{Constants.CR}{Constants.LF}--comment{Constants.CR}{Constants.LF}start--comment{Constants.CR}{Constants.LF}end--comment{Constants.CR}{Constants.LF}--comment{Constants.CR}{Constants.LF}",
                    $"{Constants.CR}{Constants.LF}{Constants.CR}{Constants.LF}start{Constants.CR}{Constants.LF}end{Constants.CR}{Constants.LF}{Constants.CR}{Constants.LF}",
                },
            };

            [Theory]
            [MemberData(nameof(InlineCommentTestData))]
            public void SQL_query_is_split_into_the_correct_number_of_statements(string sqlPart, string expectedResult)
            {
                // Act
                var result = new List<SqlPart>() { new SqlPart() {
                    Type = SqlPartType.Other,
                    Value = sqlPart,
                }};

                RemoveComments(result);

                // Assert
                Assert.Equal(expectedResult, result[0].Value);
            }
        }
    }
}

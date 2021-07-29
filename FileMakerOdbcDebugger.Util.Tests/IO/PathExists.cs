using Xunit;

namespace FileMakerOdbcDebugger.Util.Tests
{
    public partial class IO
    {
        public class PathExists
        {
            [Fact]
            public void ReturnTrueForValidPath()
            {
                // Arrange
                string path = @"C:\Windows\System32\";

                // Act
                bool actual = Util.IO.PathExists(path);

                // Assert
                Assert.True(actual);
            }

            [Theory]
            [InlineData(@"C:\InvalidFolder\InvalidFile.txt")]
            [InlineData(@"\\InvalidHost\InvalidFolder\")]
            [InlineData(@"invalidpath!")]
            public void ReturnFalseForInvalidPath(string path)
            {
                // Act
                bool actual = Util.IO.PathExists(path);

                // Assert
                Assert.False(actual);
            }
        }
    }
}

using Xunit;

namespace FileMakerOdbcDebugger.Util.Tests {

    public partial class FileMaker {

        public class GetOdbcDriverVersion {

            [Fact]
            public void ReturnsSomething() {

                // Act
                string actual = Util.FileMaker.GetOdbcDriverVersion64Bit();

                // Assert
                Assert.False(string.IsNullOrWhiteSpace(actual));

            }

        }

    }

}

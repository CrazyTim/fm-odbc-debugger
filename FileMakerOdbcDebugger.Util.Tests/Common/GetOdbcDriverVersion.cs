using Xunit;

namespace FileMakerOdbcDebugger.Util.Tests {

    public partial class FileMaker {

        public class GetOdbcDriverVersion {

            [Fact]
            public void ReturnsSomething() {

                // Act
                string actual = Common.GetOdbcDriverVersion64Bit();

                // Assert
                Assert.False(string.IsNullOrWhiteSpace(actual));

            }

        }

    }

}

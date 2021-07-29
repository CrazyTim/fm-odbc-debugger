using Xunit;

namespace FileMakerOdbcDebugger.Util.Tests
{
    public static class Security
    {
        public class Encrypt
        {
            [Theory]
            [InlineData("foo")]
            public void Ciphertext_can_be_decrypted(string plaintext)
            {
                // Act
                var result = Util.Security.Encrypt(plaintext);

                // Assert
                Assert.Equal(plaintext, Util.Security.Decrypt(result));
            }

            [Theory]
            [InlineData("foo")]
            public void Encryption_of_the_same_plaintext_produces_a_different_ciphertext(string plaintext)
            {
                // Act
                var result1 = Util.Security.Encrypt(plaintext);
                var result2 = Util.Security.Encrypt(plaintext);

                // Assert
                Assert.NotEqual(result1, result2);
            }
        }
    }
}

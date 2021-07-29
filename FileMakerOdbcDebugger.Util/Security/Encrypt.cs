using System;
using System.Security.Cryptography;
using System.Text;

namespace FileMakerOdbcDebugger.Util
{
    /// <summary>
    /// See https://stackoverflow.com/a/9034247/737393
    /// </summary>
    public static partial class Security
    {
        public static string Encrypt(this string s)
        {
            return Convert.ToBase64String(
                ProtectedData.Protect(
                    Encoding.Unicode.GetBytes(s), null, DataProtectionScope.CurrentUser)
                );
        }

        public static string Decrypt(this string s)
        {
            return Encoding.Unicode.GetString(
                ProtectedData.Unprotect(
                     Convert.FromBase64String(s), null, DataProtectionScope.CurrentUser)
                );
        }
    }
}

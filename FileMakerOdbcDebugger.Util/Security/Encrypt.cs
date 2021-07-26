using System;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;

namespace FileMakerOdbcDebugger.Util
{
    /// <summary>
    /// See https://stackoverflow.com/a/9034247/737393
    /// </summary>
    [DebuggerStepThrough]
    public static class Security
    {
        public static string Encrypt(this string s)
        {
            return Convert.ToBase64String(
                ProtectedData.Protect(
                    Encoding.Unicode.GetBytes(s), null, DataProtectionScope.CurrentUser)
                );
        }

        public static string Derypt(this string s)
        {
            return Encoding.Unicode.GetString(
                ProtectedData.Unprotect(
                     Convert.FromBase64String(s), null, DataProtectionScope.CurrentUser)
                );
        }
    }
}

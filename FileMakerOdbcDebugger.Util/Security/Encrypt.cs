using System;
using System.Diagnostics;
using System.Security.Cryptography;
using System.Text;


namespace FileMakerOdbcDebugger.Util {

    [DebuggerStepThrough]
    public static class Security {

        // refer: https://stackoverflow.com/a/9034247/737393

        public static string Encrypt(this string text) {
            return Convert.ToBase64String(
                ProtectedData.Protect(
                    Encoding.Unicode.GetBytes(text), null, DataProtectionScope.CurrentUser)
                );
        }

        public static string Derypt(this string text) {
            return Encoding.Unicode.GetString(
                ProtectedData.Unprotect(
                     Convert.FromBase64String(text), null, DataProtectionScope.CurrentUser)
                );
        }

    }

}

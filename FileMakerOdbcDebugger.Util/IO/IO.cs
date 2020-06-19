using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;

namespace FileMakerOdbcDebugger.Util {

    public static class IO {

        // Note: a StringBuilder is required for interops calls that use strings.
        [DllImport("Shlwapi.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool PathFileExists(StringBuilder path);

        /// <summary>
        /// Return True if the given path exists.
        /// Fast! Will not hang if the path or network location doesn't exist.
        /// Won't throw an exception for ilegal paths.
        /// </summary>
        public static bool PathExists(string path) {

            // Source: http://stackoverflow.com/questions/2225415/why-is-file-exists-much-slower-when-the-file-does-not-exist
            
            if (string.IsNullOrWhiteSpace(path)) {
                return false;
            }

            var sBuilder = new StringBuilder(path);

            return PathFileExists(sBuilder);

        }

    }

}

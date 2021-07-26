using System.Runtime.InteropServices;
using System.Text;

namespace FileMakerOdbcDebugger.Util
{
    public static class IO
    {
        // Note: a StringBuilder is required for interops calls that use strings.
        [DllImport("Shlwapi.dll", SetLastError = true, CharSet = CharSet.Auto)]
        private static extern bool PathFileExists(StringBuilder path);

        /// <summary>
        /// Return True if the given path exists.
        /// Fast! Will not hang if the path or network location doesn't exist.
        /// Won't throw an exception for ilegal paths.
        /// See http://stackoverflow.com/questions/2225415/why-is-file-exists-much-slower-when-the-file-does-not-exist.
        /// </summary>
        public static bool PathExists(string path)
        {
            if (string.IsNullOrWhiteSpace(path)) return false;
            return PathFileExists(new StringBuilder(path));
        }
    }
}

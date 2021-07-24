namespace FileMakerOdbcDebugger.Util
{
    public static partial class Extensions
    {
        /// <summary>
        /// Return a string with n characters removed from the start.
        /// </summary>
        public static string TrimNCharsFromStart(this string s, int n)
        {
            if (string.IsNullOrEmpty(s) || s.Length <= n) return "";
            return s.Substring(n, s.Length - n);
        }

        /// <summary>
        /// Return a string with n characters removed from the end.
        /// </summary>
        public static string TrimNCharsFromEnd(this string s, int n)
        {
            if (string.IsNullOrEmpty(s) || s.Length <= n) return "";
            return s.Substring(0, s.Length - n);
        }
    }
}

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Extensions
    {
        /// <summary>
        /// Return True if the int is "odd" (can't be divided evenly by 2).
        /// </summary>
        public static bool IsOdd(this int value)
        {
            return (value % 2 != 0);
        }

        /// <summary>
        /// Return True if the int is "even" (can be divided evenly by 2).
        /// </summary>
        public static bool IsEven(this int value)
        {
            return (value % 2 == 0);
        }
    }
}

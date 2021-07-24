using System;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Extensions
    {
        public static string ToDisplayString(this TimeSpan t)
        {
            return (t.Minutes == 0)
                ? t.Seconds.ToString("#0") + "." + (t.Milliseconds / 10).ToString("00")
                : t.Minutes.ToString("#0") + ":" + t.Seconds.ToString("#0") + "." + (t.Milliseconds / 10).ToString("00");
        }
    }
}

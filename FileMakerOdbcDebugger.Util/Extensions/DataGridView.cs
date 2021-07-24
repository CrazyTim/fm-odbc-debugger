using System.Reflection;
using System.Windows.Forms;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Extensions
    {
        /// <summary>
        /// Turn on double buffering to significantly improve draw performance.
        /// </summary>
        public static void EnableDoubleBuffering(this DataGridView d)
        {
            typeof(DataGridView).InvokeMember(
                "DoubleBuffered",
                BindingFlags.NonPublic | BindingFlags.Instance | BindingFlags.SetProperty,
                default, d, new object[] { true }
                );
        }
    }
}

using System.Windows.Forms;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Extensions
    {
        private delegate void ShowMsg_Delegate(Form f, string message, MessageBoxIcon icon, string title);

        /// <summary>
        /// A thread-safe and reliable method to show a MessageBox over a form.
        /// </summary>
        public static void ShowMsg(
            this Form f,
            string message,
            MessageBoxIcon icon = MessageBoxIcon.Information,
            string title = "")
        {
            if (f.InvokeRequired)
            {
                var d = new ShowMsg_Delegate(ShowMsg);
                f.Invoke(d, new object[] { f, message, icon, title });
                return;
            }
            if (string.IsNullOrEmpty(title))
            {
                title = f.Text;
            }
            if (!f.IsDisposed)
            {
                MessageBox.Show(f, message, title, MessageBoxButtons.OK, icon);
            }
        }
    }
}

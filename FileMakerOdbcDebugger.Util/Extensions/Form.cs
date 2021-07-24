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
            this Form form,
            string message,
            MessageBoxIcon icon = MessageBoxIcon.Information,
            string title = "")
        {
            if (form.InvokeRequired)
            {
                var d = new ShowMsg_Delegate(ShowMsg);
                form.Invoke(d, new object[] { form, message, icon, title });
                return;
            }
            if (string.IsNullOrEmpty(title))
            {
                title = form.Text;
            }
            if (!form.IsDisposed)
            {
                MessageBox.Show(form, message, title, MessageBoxButtons.OK, icon);
            }
        }
    }
}

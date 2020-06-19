using System.Windows.Forms;

namespace FileMakerOdbcDebugger.Util {

    public static partial class Extensions {

        private delegate void ShowMsg_Delegate(Form f, string Msg, MessageBoxIcon Icon, string Title);

        /// <summary>
        /// A thread-safe and reliable method to show a MessageBox over a form.
        /// </summary>
        public static void ShowMsg(this Form f,
                                   string Msg,
                                   MessageBoxIcon Icon = MessageBoxIcon.Information,
                                   string Title = "") {

            if (f.InvokeRequired) {
                var d = new ShowMsg_Delegate(ShowMsg);
                f.Invoke(d, new object[] { f, Msg, Icon, Title });
                return;
            }

            if (string.IsNullOrEmpty(Title)) {
                Title = f.Text;
            }

            if (!f.IsDisposed) {
                MessageBox.Show(f, Msg, Title, MessageBoxButtons.OK, Icon);
            }

        }

    }

}

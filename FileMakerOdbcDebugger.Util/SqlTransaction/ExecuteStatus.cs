using System.ComponentModel;

namespace FileMakerOdbcDebugger.Util
{
    public partial class SqlTransaction
    {
        public enum ExecuteStatus
        {
            [Description("Connecting...")]
            Connecting,

            [Description("Executing Query. Waiting for response...")]
            Executing,

            [Description("Streaming Data...")]
            Streaming,

            [Description("Error")]
            Error,
        }
    }
}

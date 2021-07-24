using System.ComponentModel;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Sql
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

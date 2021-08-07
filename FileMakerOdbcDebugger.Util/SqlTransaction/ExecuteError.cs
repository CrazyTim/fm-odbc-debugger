using System.ComponentModel;

namespace FileMakerOdbcDebugger.Util
{
    public partial class SqlTransaction
    {
        public enum ExecuteError
        {
            [Description("ERROR: An un-representable DateTime value was returned. FileMaker isn't always strict about datatypes, so its likely that a date field contains a non-date value.")]
            UnrepresentableDateTimeValue,

            [Description("ERROR: null server.")]
            NullServer,

            [Description("ERROR: null database name.")]
            NullDatabaseName,

            [Description("ERROR: null username.")]
            NullUsername,

            [Description("ERROR: null connection string.")]
            NullConnectionString,
        }
    }
}

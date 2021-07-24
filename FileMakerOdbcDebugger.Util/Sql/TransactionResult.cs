using System;
using System.Collections.Generic;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Sql
    {
        public class TransactionResult
        {
            public TimeSpan Duration_Connect { get; set; }

            public List<StatementResult> Results { get; set; } = new List<StatementResult>();

            /// <summary>
            /// An error that occured when executing the query, or an empty string.
            /// </summary>
            public string Error { get; set; }
        }
    }
}

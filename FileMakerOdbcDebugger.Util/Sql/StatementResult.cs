using System;
using System.Collections.Generic;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Sql
    {
        public class StatementResult
        {
            public TimeSpan Duration_Execute { get; set; }

            public TimeSpan Duration_Stream { get; set; }

            /// <summary>
            /// Data returned from the query. The first item contains the column headers.
            /// </summary>
            public List<List<string>> Data { get; set; } = new List<List<string>>();

            /// <summary>
            /// The number of rows affected by the query.
            /// </summary>
            public int RowsAffected { get; set; }
        }
    }
}

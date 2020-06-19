using System.Collections.Generic;

namespace FileMakerOdbcDebugger.Util {

    public static partial class Sql {

        /// <summary>
        /// Execute each query in a single transaction. Roll back the transaction if any of them fail.
        /// </summary>
        public static void ExecuteTransaction(List<string> Queries, System.Data.Odbc.OdbcConnection cn) {

            if (Queries.Count == 0) {
                return;
            }

            using (var cmd = new System.Data.Odbc.OdbcCommand()) {

                cmd.Connection = cn; 

                using (var t = cn.BeginTransaction()) {

                    cmd.Transaction = t; // Assign transaction object for a pending local transaction.

                    foreach (var QueryString in Queries) {

                        if (!string.IsNullOrEmpty(QueryString)) { // skip if empty string

                            cmd.CommandText = QueryString;
                            cmd.ExecuteNonQuery();

                        }

                    }

                    try {
                        t.Commit();
                    } catch {
                        t.Rollback();
                        throw;
                    }

                }
            }

        }

    }

}

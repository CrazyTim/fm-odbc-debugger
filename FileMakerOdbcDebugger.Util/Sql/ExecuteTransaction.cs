using System;
using System.Collections.Generic;
using System.Data.Odbc;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Sql
    {
        /// <summary>
        /// Execute one or more SQL statements as a transaction. Roll back the transaction if any of them fail.
        /// </summary>
        public static TransactionResult ExecuteTransaction(List<string> sqlStatements,
                                                           IUi ui,
                                                           int rowLimit,
                                                           string connectionString)
        {
            var transactionResult = new TransactionResult();
            var sw = new System.Diagnostics.Stopwatch();
            var currentStatement = "";

            if (sqlStatements.Count == 0) return transactionResult;

            try
            {
                using (var cn = new OdbcConnection(connectionString))
                {
                    // START CONNECT
                    ui.SetStatusLabel(ExecuteStatus.Connecting.Description());
                    sw.Restart();
                    cn.Open();
                    transactionResult.Duration_Connect = sw.Elapsed;
                    // END CONNECT

                    using (var cmd = new OdbcCommand())
                    {
                        cmd.Connection = cn;

                        // START EXECUTE
                        ui.SetStatusLabel(ExecuteStatus.Executing.Description());
                        sw.Restart();

                        using (var t = cn.BeginTransaction())
                        {
                            cmd.Transaction = t; // Assign transaction object for a pending local transaction.

                            foreach (var statement in sqlStatements)
                            {
                                if (string.IsNullOrEmpty(statement)) continue;
                            
                                currentStatement = statement.Trim();
                                cmd.CommandText = currentStatement;
                                    
                                var statementResult = new StatementResult();

                                if (!currentStatement.StartsWith("select", StringComparison.OrdinalIgnoreCase))
                                {
                                    statementResult.RowsAffected = cmd.ExecuteNonQuery();
                                    statementResult.Duration_Execute = sw.Elapsed;
                                    // END EXECUTE
                                }
                                else
                                {
                                    using (var reader = cmd.ExecuteReader())
                                    {
                                        statementResult.Duration_Execute = sw.Elapsed;
                                        // END EXECUTE

                                        // BEGIN STREAM
                                        ui.SetStatusLabel(ExecuteStatus.Streaming.Description());
                                        sw.Restart();

                                        statementResult.Data.Add(GetColumns(reader));

                                        int readCount = 1;
                                        while (readCount <= rowLimit && reader.Read())
                                        {
                                            statementResult.Data.Add(GetData(reader));
                                            readCount += 1;
                                        }

                                        statementResult.Duration_Stream = sw.Elapsed;
                                        // END STREAM

                                        ui.SetStatusLabel("");

                                        statementResult.RowsAffected = reader.RecordsAffected;
                                    }
                                }
                                transactionResult.Results.Add(statementResult);
                            }

                            try
                            {
                                t.Commit();
                            }
                            catch
                            {
                                t.Rollback();
                                throw;
                            }
                        }
                    }
                }
            }
            catch (ArgumentOutOfRangeException ex)
            {
                if (ex.Message == "Year, Month, and Day parameters describe an un-representable DateTime.")
                {
                    // filemaker allows importing incorrect data into fields, so we need to catch these errors!
                    transactionResult.Error = ExecuteError.UnrepresentableDateTimeValue.Description();
                }
                else
                {
                    transactionResult.Error = ex.Message;
                }

                transactionResult.Error += Environment.NewLine + Environment.NewLine + currentStatement;
            }
            catch (Exception ex)
            {
                transactionResult.Error = ex.Message + Environment.NewLine + Environment.NewLine + currentStatement;
            }

            return transactionResult;
        }

        private static List<string> GetColumns(OdbcDataReader reader)
        {
            var columns = new List<string>();

            for (int i = 0, loopTo = reader.FieldCount - 1; i <= loopTo; i++)
            {
                columns.Add(reader.GetName(i));
            }

            return columns;
        }

        private static List<string> GetData(OdbcDataReader reader)
        {
            var data = new List<string>();

            for (int i = 0, loopTo1 = reader.FieldCount - 1; i <= loopTo1; i++)
            {
                if (reader.GetValue(i).GetType() == typeof(TimeSpan))
                {
                    // Convert timespans into a DateTime.
                    // Filemaker's "time" data type casts to TimeSpan, which isn't supported by DataGridView, so we convert to a string.
                    var d = new DateTime();
                    TimeSpan ts = (TimeSpan)reader.GetValue(i);
                    d = d.Add(ts);
                    data.Add(d.ToString("h:mm:ss tt").ToLower());
                }
                else if (!reader.IsDBNull(i))
                {
                    data.Add((string)reader.GetValue(i));
                }
                else
                {
                    data.Add(null);
                }
            }

            return data;
        }
    }
}

using System;
using System.Collections.Generic;
using System.Data.Common;
using System.Data.Odbc;
using System.Threading.Tasks;

namespace FileMakerOdbcDebugger.Util
{
    public static partial class Sql
    {
        /// <summary>
        /// Execute one or more SQL statements as a transaction. Roll back the transaction if any of them fail.
        /// </summary>
        public async static Task ExecuteTransaction(
            Transaction transaction,
            IUi ui,
            int rowLimit,
            string connectionString)
        {
            List<string> statements = SplitQueryIntoStatements(transaction.Query);
            if (statements.Count == 0) return;

            var sw = new System.Diagnostics.Stopwatch();
            var currentStatement = "";

            try
            {
                using (var cn = new OdbcConnection(connectionString))
                {
                    // START CONNECT
                    ui.SetStatusLabel(ExecuteStatus.Connecting.Description());
                    sw.Restart();
                    cn.Open();
                    transaction.Duration_Connect = sw.Elapsed;
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

                            foreach (var statement in statements)
                            {
                                if (string.IsNullOrEmpty(statement)) continue;
                            
                                currentStatement = statement.Trim();
                                cmd.CommandText = currentStatement;
                                    
                                var statementResult = new StatementResult();

                                if (!currentStatement.StartsWith("select", StringComparison.OrdinalIgnoreCase))
                                {
                                    statementResult.RowsAffected = await cmd.ExecuteNonQueryAsync();
                                    statementResult.Duration_Execute = sw.Elapsed;
                                    // END EXECUTE
                                }
                                else
                                {
                                    using (var reader = await cmd.ExecuteReaderAsync())
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
                                transaction.Results.Add(statementResult);
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
                    transaction.Error = ExecuteError.UnrepresentableDateTimeValue.Description();
                }
                else
                {
                    transaction.Error = ex.Message;
                }

                transaction.Error += Environment.NewLine + Environment.NewLine + currentStatement;
            }
            catch (Exception ex)
            {
                transaction.Error = ex.Message + Environment.NewLine + Environment.NewLine + currentStatement;
            }
        }

        private static List<string> GetColumns(DbDataReader reader)
        {
            var columns = new List<string>();

            for (int i = 0, loopTo = reader.FieldCount - 1; i <= loopTo; i++)
            {
                columns.Add(reader.GetName(i));
            }

            return columns;
        }

        private static List<string> GetData(DbDataReader reader)
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

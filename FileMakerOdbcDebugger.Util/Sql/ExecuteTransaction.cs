using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data.Odbc;

namespace FileMakerOdbcDebugger.Util {

    public static partial class Sql {

        public interface IUiForm {
            void SetStatusLabel(string label);
        }

        public class ExecuteTransactionResult {

            public TimeSpan Duration_Connect { get; set; }

            public List<StatementResult> Results { get; set; } = new List<StatementResult>();

            /// <summary>
            /// Any pre-query warnings.
            /// </summary>
            public List<ArrayList> Warnings { get; set; }

            /// <summary>
            /// A post-query error message if there was one, or it will be blank.
            /// </summary>
            public string Error { get; set; }

        }

        public class StatementResult {

            public TimeSpan Duration_Execute { get; set; }

            public TimeSpan Duration_Stream { get; set; }

            /// <summary>
            /// List of results. First row is column headers.
            /// </summary>
            public List<ArrayList> Data { get; set; } = new List<ArrayList>();

            /// <summary>
            /// Number of rows affected.
            /// </summary>
            public int RowsAffected { get; set; }

        }

        public enum ExecuteStatus {

            [Description("Connecting...")]
            Connecting,

            [Description("Executing Query. Waiting for response...")]
            Executing,

            [Description("Streaming Data...")]
            Streaming,

            [Description("Error")]
            Error

        }

        public enum ExecuteError {

            [Description("ERROR: An un-representable DateTime value was returned. FileMaker isn't always strict about datatypes, so its likely that a date field contains a non-date value.")]
            UnrepresentableDateTimeValue,

            [Description("ERROR: null server.")]
            NullServer,

            [Description("ERROR: null database name.")]
            NullDatabaseName,

            [Description("ERROR: null username.")]
            NullUsername,

            [Description("ERROR: null connection string.")]
            NullConnectionString

        }



        /// <summary>
        /// Execute each query in a single transaction. Roll back the transaction if any of them fail.
        /// </summary>
        public static ExecuteTransactionResult ExecuteTransaction(List<string> statements,
                                                                  IUiForm parent,
                                                                  int rowLimit,
                                                                  string connectionString) {

            var result = new ExecuteTransactionResult();

            var sw = new System.Diagnostics.Stopwatch();

            var currentStatement = "";

            if (statements.Count == 0) {
                return result;
            }

            try {

                using (var cn = new OdbcConnection(connectionString)) {

                    // Connect:
                    parent.SetStatusLabel(ExecuteStatus.Connecting.Description());
                    sw.Restart();
                    cn.Open();
                    result.Duration_Connect = sw.Elapsed;

                    using (var cmd = new System.Data.Odbc.OdbcCommand()) {

                        cmd.Connection = cn;

                        // Execute:
                        parent.SetStatusLabel(ExecuteStatus.Executing.Description());
                        sw.Restart();

                        using (var t = cn.BeginTransaction()) {

                            cmd.Transaction = t; // Assign transaction object for a pending local transaction.

                            foreach (var statement in statements) {

                                if (!string.IsNullOrEmpty(statement)) { // Skip if empty string.

                                    currentStatement = statement.Trim();
                                    cmd.CommandText = currentStatement;
                                    
                                    var statementResult = new StatementResult();

                                    if (!currentStatement.StartsWith("select", StringComparison.OrdinalIgnoreCase)) {

                                        statementResult.RowsAffected = cmd.ExecuteNonQuery();

                                        statementResult.Duration_Execute = sw.Elapsed;

                                    } else {

                                        using (var reader = cmd.ExecuteReader()) {

                                            statementResult.Duration_Execute = sw.Elapsed;

                                            // Stream:
                                            parent.SetStatusLabel(ExecuteStatus.Streaming.Description());
                                            sw.Restart();
                                                
                                            var s = new ArrayList();

                                            // Read column headers:
                                            for (int i = 0, loopTo = reader.FieldCount - 1; i <= loopTo; i++) {

                                                var ColumnName = reader.GetName(i).Trim();

                                                if (s.Contains(ColumnName)) {

                                                    // We can't add duplicate columns to a DataSet
                                                    // So here we incrementally append a special zero-width space on the end of the column name until it is not a duplicate.
                                                    // These characters can be trimmed later from the column caption.

                                                    var incrementChar = "\u200B";

                                                    ColumnName += incrementChar;

                                                    do {
                                                        ColumnName += incrementChar;
                                                    } while (s.Contains(ColumnName));

                                                }

                                                s.Add(ColumnName);
                                            }

                                            statementResult.Data.Add(s);

                                            // Read column data:
                                            int readCount = 1;

                                            while (readCount <= rowLimit && reader.Read()) {

                                                s = new ArrayList();

                                                for (int i = 0, loopTo1 = reader.FieldCount - 1; i <= loopTo1; i++) {

                                                    if (reader.GetValue(i).GetType() == typeof(TimeSpan)) {

                                                        // Convert timespans into a DateTime:
                                                        // Filemake "time" fields return timepsans which aren't supported by the the DataGridView.
                                                        var d = new DateTime();
                                                        TimeSpan ts = (TimeSpan)reader.GetValue(i);
                                                        d = d.Add(ts);
                                                        s.Add(d.ToString("h:mm:ss tt").ToLower());

                                                    } else if (!reader.IsDBNull(i)) {

                                                        s.Add(reader.GetValue(i));

                                                    } else {

                                                        s.Add(null);

                                                    }

                                                }

                                                readCount += 1;

                                                statementResult.Data.Add(s);

                                            }

                                            statementResult.Duration_Stream = sw.Elapsed;

                                            parent.SetStatusLabel("");

                                            statementResult.RowsAffected = reader.RecordsAffected;

                                        }

                                    }

                                    result.Results.Add(statementResult);

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

            } catch (ArgumentOutOfRangeException ex) {

                if (ex.Message == "Year, Month, and Day parameters describe an un-representable DateTime.") { // filemaker allows importing incorrect data into fields, so we need to catch these errors!
                    result.Error = ExecuteError.UnrepresentableDateTimeValue.Description();
                } else {
                    result.Error = ex.Message;
                }

                result.Error += Environment.NewLine + Environment.NewLine + currentStatement;

            } catch (Exception ex) {

                result.Error = ex.Message;

                result.Error += Environment.NewLine + Environment.NewLine + currentStatement;

            }

            return result;

        }

    }

}

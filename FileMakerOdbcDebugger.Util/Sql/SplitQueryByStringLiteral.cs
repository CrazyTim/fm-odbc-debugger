using System;
using System.Collections;
using System.Collections.Generic;
using System.Linq;
using System.Reflection.Emit;
using System.Text;
using System.Threading.Tasks;

namespace FileMakerOdbcDebugger.Util {

    public static partial class Sql {

        /// <summary>
        /// Split a query by the string literals contained in it.
        /// The result can be used to pre-process string literals separate from everything else.
        /// Example result: ["SELECT * FROM table WHERE field = ", "test"]
        /// </summary>
        public static SplitQuery SplitQueryByStringLiteral(string query) {

            var splitQuery = new SplitQuery();
            var singleQuote = '\'';

            if (string.IsNullOrEmpty(query)) {
                return splitQuery;
            }

            // if query starts with a single quote, but its not an escape, add an empty string before it so string literals are always at even numbered indexes 
            if (query[0] == singleQuote && query[1] != singleQuote) {
                splitQuery.Parts.Add("");
            }

            if (query.Length >= 2) { // query has to be at least two chars long because we look ahead one char...

                var i = 0;

                while (i < query.Length) {

                    if ((query[i]) == singleQuote) {

                        if (i == query.Length - 1) {
                            // Trim last single quote
                            query = query.TrimNCharsFromEnd(1);
                            continue;
                        }

                        if (query[i + 1] == singleQuote) { // The first single quote was escaping the second 
                            i += 2;
                            continue; // Skip both. No need to split.
                        }

                        if (i == 0) {
                            // Trim first single quote
                            query = query.TrimNCharsFromStart(1);
                            continue;
                        }

                        splitQuery.Parts.Add(query.Substring(0, i));
                        query = query.TrimNCharsFromStart(i + 1);
                        i = 0;

                    }

                    i += 1;

                }
            }

            splitQuery.Parts.Add(query); // Add remainder

            return splitQuery;

        }

        /// <summary>
        /// Represents a query that has been split by the string literals contained in it.
        /// </summary>
        public class SplitQuery {

            public List<string> Parts { get; set; }

            public List<string> Statements {
                get { return GetParts(0); }
            }

            public List<string> StringLiterals {
               get { return GetParts(1); }
            }

            public SplitQuery() {
                Parts = new List<string>();
            }

            private List<string> GetParts(int partType) {

                var r = new List<string>();

                for (var i = partType; i < Parts.Count; i++) {
                    r.Add(Parts[i]);
                    i += 2;
                }

                return r;

            }

            /// <summary>
            /// Return the query parts joined back together.
            /// </summary>
            public string Rejoin() {

                if (Parts.Count == 0) {
                    return "";
                }

                var joined = "";

                for (var i = 0; i < Parts.Count; i++) {
                    joined += Parts[i] + '\'';
                }

                if (Parts.Count.IsOdd()) {
                    joined = joined.TrimNCharsFromEnd(1);
                }

                return joined;

            }

        }

    }

}

using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace FileMakerOdbcDebugger.Util {

    public static partial class Sql {

        /// <summary> Split a query into an array of statements based on the occurance of any "statement terminators" (";"). </summary>
        public static List<string> SplitQueryIntoStatements(string query) {

            // Replace "query end" identifiers (";") with (char)159 so we can easily split it into parts.

            int singleQuoteCount = 0;
            var builder = new StringBuilder(query);

            for (int i = 0, loopTo = builder.Length - 1; i <= loopTo; i++) {

                char s = builder[i];

                if (s == '\'') {

                    if (singleQuoteCount == 1) {
                        singleQuoteCount = 0;
                    } else {
                        singleQuoteCount = 1;
                    }

                } else if (s == ';') {

                    if (singleQuoteCount == 0) {
                        builder[i] = (char)159; // replace
                    }

                }

            }

            var l = builder.ToString().Split('\u009f').ToList();

            l = l.RemoveEmptyElements();

            return l;

        }


    }

}

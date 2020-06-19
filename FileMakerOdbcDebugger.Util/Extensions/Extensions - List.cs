using System.Collections.Generic;

namespace FileMakerOdbcDebugger.Util {

    public static partial class Extensions {

        /// <summary>
        /// Remove elements from the list that are null, empty, or contain only whitespace characters.
        /// </summary>
        public static List<string> RemoveEmptyElements(this List<string> l) {
            l.RemoveAll(str => string.IsNullOrWhiteSpace(str));
            return l;
        }

    }

}

namespace FileMakerOdbcDebugger.Util
{
    public static partial class SqlQuery
    {
        /// <summary>
        /// Prepare a SQL query to be executed.
        /// </summary>
        public static string Prepare(string sqlQuery, bool forFileMaker)
        {
            var split = new Split(sqlQuery);

            RemoveComments(split.Parts);

            CleanSpaces(split.Parts);

            if (forFileMaker)
            {
                ReplaceLineBreaksInStringLiteralsWithCR(split.Parts);
                EscapeIdentifiersStartingWithUnderscore(split.Parts);
            }

            return split.Join().Trim();
        }
    }
}

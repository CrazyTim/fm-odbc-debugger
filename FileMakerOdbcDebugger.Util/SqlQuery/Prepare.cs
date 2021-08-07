namespace FileMakerOdbcDebugger.Util
{
    public static partial class SqlQuery
    {
        /// <summary>
        /// Prepare a SQL query to be executed.
        /// </summary>
        public static string Prepare(string sqlQuery, bool forFileMaker)
        {
            sqlQuery = RemoveComments(sqlQuery);

            var split = new Split(sqlQuery);

            CleanSpaces(split.Parts);

            if (forFileMaker)
            {
                ReplaceNewlinesInStringLiteralsWithCR(split.Parts);
                EscapeIdentifiersStartingWithUnderscore(split.Parts);
            }

            return split.Join().Trim();
        }
    }
}

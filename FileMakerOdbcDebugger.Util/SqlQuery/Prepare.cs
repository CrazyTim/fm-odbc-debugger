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

            split.Parts
                .RemoveComments()
                .CleanSpaces();

            if (forFileMaker)
            {
                split.Parts
                    .ReplaceLineBreaksInStringLiteralsWithCR()
                    .EscapeIdentifiersStartingWithUnderscore();
            }

            return split.Join().Trim();
        }
    }
}

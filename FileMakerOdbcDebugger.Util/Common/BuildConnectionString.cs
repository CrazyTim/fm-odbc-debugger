namespace FileMakerOdbcDebugger.Util
{
    public static partial class Common
    {
        public static string BuildConnectionString(string serverAddress, string username, string password, string databaseName)
        {
            return $"DRIVER={{FileMaker ODBC}};SERVER={serverAddress};UID={username};PWD={password};DATABASE={databaseName};";
        }
    }
}

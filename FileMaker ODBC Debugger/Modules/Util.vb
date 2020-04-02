Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports System.Xml.Serialization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports FastColoredTextBoxNS

'<DebuggerStepThrough()>
Module Util

    Public ReadOnly AppDataDir As String = My.Application.GetEnvironmentVariable("APPDATA") & "\FileMaker ODBC Debugger"
    Public ReadOnly FmOdbcDriverVersion As String = GetFmOdbcDriverVersion()

    Public Const MAX_TABS As Integer = 7

    ' define syntax regexp
    Public ReadOnly SyntaxSQL_String As New Regex("""""|''|"".*?[^\\]""|'.*?[^\\]'", RegexOptions.Compiled)
    Public ReadOnly SyntaxSQL_Number As New Regex("\b\d+[\.]?\d*([eE]\-?\d+)?\b", RegexOptions.Compiled)
    Public ReadOnly SyntaxSQL_Comment1 As New Regex("--.*$", RegexOptions.Multiline Or RegexOptions.Compiled)
    Public ReadOnly SyntaxSQL_Comment2 As New Regex("(/\*.*?\*/)|(/\*.*)", RegexOptions.Singleline Or RegexOptions.Compiled)
    Public ReadOnly SyntaxSQL_Comment3 As New Regex("(/\*.*?\*/)|(.*\*/)", RegexOptions.Singleline Or RegexOptions.RightToLeft Or RegexOptions.Compiled)
    Public ReadOnly SyntaxSQL_Variables As New Regex("@[a-zA-Z_\d]*\b", RegexOptions.Compiled)
    Public ReadOnly SyntaxSQL_Statements As New Regex("\b(ALTER APPLICATION ROLE|ALTER ASSEMBLY|ALTER ASYMMETRIC KEY|ALTER AUTHORIZATION|ALTER BROKER PRIORITY|ALTER CERTIFICATE|ALTER CREDENTIAL|ALTER CRYPTOGRAPHIC PROVIDER|ALTER DATABASE|ALTER DATABASE AUDIT SPECIFICATION|ALTER DATABASE ENCRYPTION KEY|ALTER ENDPOINT|ALTER EVENT SESSION|ALTER FULLTEXT CATALOG|ALTER FULLTEXT INDEX|ALTER FULLTEXT STOPLIST|ALTER FUNCTION|ALTER INDEX|ALTER LOGIN|ALTER MASTER KEY|ALTER MESSAGE TYPE|ALTER PARTITION FUNCTION|ALTER PARTITION SCHEME|ALTER PROCEDURE|ALTER QUEUE|ALTER REMOTE SERVICE BINDING|ALTER RESOURCE GOVERNOR|ALTER RESOURCE POOL|ALTER ROLE|ALTER ROUTE|ALTER SCHEMA|ALTER SERVER AUDIT|ALTER SERVER AUDIT SPECIFICATION|ALTER SERVICE|ALTER SERVICE MASTER KEY|ALTER SYMMETRIC KEY|ALTER TABLE|ALTER TRIGGER|ALTER USER|ALTER VIEW|ALTER WORKLOAD GROUP|ALTER XML SCHEMA COLLECTION|BULK INSERT|CREATE AGGREGATE|CREATE APPLICATION ROLE|CREATE ASSEMBLY|CREATE ASYMMETRIC KEY|CREATE BROKER PRIORITY|CREATE CERTIFICATE|CREATE CONTRACT|CREATE CREDENTIAL|CREATE CRYPTOGRAPHIC PROVIDER|CREATE DATABASE|CREATE DATABASE AUDIT SPECIFICATION|CREATE DATABASE ENCRYPTION KEY|CREATE DEFAULT|CREATE ENDPOINT|CREATE EVENT NOTIFICATION|CREATE EVENT SESSION|CREATE FULLTEXT CATALOG|CREATE FULLTEXT INDEX|CREATE FULLTEXT STOPLIST|CREATE FUNCTION|CREATE INDEX|CREATE LOGIN|CREATE MASTER KEY|CREATE MESSAGE TYPE|CREATE PARTITION FUNCTION|CREATE PARTITION SCHEME|CREATE PROCEDURE|CREATE QUEUE|CREATE REMOTE SERVICE BINDING|CREATE RESOURCE POOL|CREATE ROLE|CREATE ROUTE|CREATE RULE|CREATE SCHEMA|CREATE SERVER AUDIT|CREATE SERVER AUDIT SPECIFICATION|CREATE SERVICE|CREATE SPATIAL INDEX|CREATE STATISTICS|CREATE SYMMETRIC KEY|CREATE SYNONYM|CREATE TABLE|CREATE TRIGGER|CREATE TYPE|CREATE USER|CREATE VIEW|CREATE WORKLOAD GROUP|CREATE XML INDEX|CREATE XML SCHEMA COLLECTION|DELETE|DISABLE TRIGGER|DROP AGGREGATE|DROP APPLICATION ROLE|DROP ASSEMBLY|DROP ASYMMETRIC KEY|DROP BROKER PRIORITY|DROP CERTIFICATE|DROP CONTRACT|DROP CREDENTIAL|DROP CRYPTOGRAPHIC PROVIDER|DROP DATABASE|DROP DATABASE AUDIT SPECIFICATION|DROP DATABASE ENCRYPTION KEY|DROP DEFAULT|DROP ENDPOINT|DROP EVENT NOTIFICATION|DROP EVENT SESSION|DROP FULLTEXT CATALOG|DROP FULLTEXT INDEX|DROP FULLTEXT STOPLIST|DROP FUNCTION|DROP INDEX|DROP LOGIN|DROP MASTER KEY|DROP MESSAGE TYPE|DROP PARTITION FUNCTION|DROP PARTITION SCHEME|DROP PROCEDURE|DROP QUEUE|DROP REMOTE SERVICE BINDING|DROP RESOURCE POOL|DROP ROLE|DROP ROUTE|DROP RULE|DROP SCHEMA|DROP SERVER AUDIT|DROP SERVER AUDIT SPECIFICATION|DROP SERVICE|DROP SIGNATURE|DROP STATISTICS|DROP SYMMETRIC KEY|DROP SYNONYM|DROP TABLE|DROP TRIGGER|DROP TYPE|DROP USER|DROP VIEW|DROP WORKLOAD GROUP|DROP XML SCHEMA COLLECTION|ENABLE TRIGGER|EXEC|EXECUTE|REPLACE|FROM|INSERT|MERGE|OPTION|OUTPUT|SELECT|TOP|TRUNCATE TABLE|UPDATE|UPDATE STATISTICS|WHERE|WITH|INTO|IN|SET)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public ReadOnly SyntaxSQL_Keywords As New Regex("\b(AND|ADD|ALL|ANY|AS|ASC|AUTHORIZATION|BACKUP|BEGIN|BETWEEN|BREAK|BROWSE|BY|CASCADE|CHECK|CHECKPOINT|CLOSE|CLUSTERED|COLLATE|COLUMN|COMMIT|COMPUTE|CONSTRAINT|CONTAINS|CONTINUE|CROSS|CURRENT|CURRENT_DATE|CURRENT_TIME|CURSOR|DATABASE|DBCC|DEALLOCATE|DECLARE|DEFAULT|DENY|DESC|DISK|DISTINCT|DISTRIBUTED|DOUBLE|DUMP|ELSE|END|ERRLVL|ESCAPE|EXCEPT|EXISTS|EXIT|EXTERNAL|FETCH|FILE|FILLFACTOR|FOR|FOREIGN|FREETEXT|FULL|FUNCTION|GOTO|GRANT|GROUP|HAVING|HOLDLOCK|IDENTITY|IDENTITY_INSERT|IDENTITYCOL|IF|INDEX|INTERSECT|IS|KEY|KILL|LIKE|LINENO|LOAD|NATIONAL|NOCHECK|NONCLUSTERED|NOT|NULL|OF|OFF|OFFSETS|ON|OPEN|OR|ORDER|OUTER|OVER|PERCENT|PIVOT|PLAN|PRECISION|PRIMARY|PRINT|PROC|PROCEDURE|PUBLIC|RAISERROR|READ|READTEXT|RECONFIGURE|REFERENCES|REPLICATION|RESTORE|RESTRICT|RETURN|REVERT|REVOKE|ROLLBACK|ROWCOUNT|ROWGUIDCOL|RULE|SAVE|SCHEMA|SECURITYAUDIT|SHUTDOWN|SOME|STATISTICS|TABLE|TABLESAMPLE|TEXTSIZE|THEN|TO|TRAN|TRANSACTION|TRIGGER|TSEQUAL|UNION|UNIQUE|UNPIVOT|UPDATETEXT|USE|USER|VALUES|VARYING|VIEW|WAITFOR|WHEN|WHILE|WRITETEXT)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public ReadOnly SyntaxSQL_Keywords2 As New Regex("\b(AND|INNER|JOIN)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public ReadOnly SyntaxSQL_Functions As New Regex("\b(ABS|ACOS|APP_NAME|ASCII|ASIN|ASSEMBLYPROPERTY|AsymKey_ID|ASYMKEY_ID|asymkeyproperty|ASYMKEYPROPERTY|ATAN|ATN2|AVG|CASE|CAST|CEILING|Cert_ID|Cert_ID|CertProperty|CHAR|CHARINDEX|CHECKSUM_AGG|COALESCE|COL_LENGTH|COL_NAME|COLLATIONPROPERTY|COLLATIONPROPERTY|COLUMNPROPERTY|COLUMNS_UPDATED|COLUMNS_UPDATED|CONTAINSTABLE|CONVERT|COS|COT|COUNT|COUNT_BIG|CRYPT_GEN_RANDOM|CURRENT_TIMESTAMP|CURRENT_TIMESTAMP|CURRENT_USER|CURRENT_USER|CURSOR_STATUS|DATABASE_PRINCIPAL_ID|DATABASE_PRINCIPAL_ID|DATABASEPROPERTY|DATABASEPROPERTYEX|DATALENGTH|DATALENGTH|DATEADD|DATEDIFF|DATENAME|DATEPART|DAY|DB_ID|DB_NAME|DECRYPTBYASYMKEY|DECRYPTBYCERT|DECRYPTBYKEY|DECRYPTBYKEYAUTOASYMKEY|DECRYPTBYKEYAUTOCERT|DECRYPTBYPASSPHRASE|DEGREES|DENSE_RANK|DIFFERENCE|ENCRYPTBYASYMKEY|ENCRYPTBYCERT|ENCRYPTBYKEY|ENCRYPTBYPASSPHRASE|ERROR_LINE|ERROR_MESSAGE|ERROR_NUMBER|ERROR_PROCEDURE|ERROR_SEVERITY|ERROR_STATE|EVENTDATA|EXP|FILE_ID|FILE_IDEX|FILE_NAME|FILEGROUP_ID|FILEGROUP_NAME|FILEGROUPPROPERTY|FILEPROPERTY|FLOOR|fn_helpcollations|fn_listextendedproperty|fn_servershareddrives|fn_virtualfilestats|fn_virtualfilestats|FORMATMESSAGE|FREETEXTTABLE|FULLTEXTCATALOGPROPERTY|FULLTEXTSERVICEPROPERTY|GETANSINULL|GETDATE|GETUTCDATE|GROUPING|HAS_PERMS_BY_NAME|HOST_ID|HOST_NAME|IDENT_CURRENT|IDENT_CURRENT|IDENT_INCR|IDENT_INCR|IDENT_SEED|IDENTITY\(|INDEX_COL|INDEXKEY_PROPERTY|INDEXPROPERTY|IS_MEMBER|IS_OBJECTSIGNED|IS_SRVROLEMEMBER|ISDATE|ISDATE|ISNULL|ISNUMERIC|Key_GUID|Key_GUID|Key_ID|Key_ID|KEY_NAME|KEY_NAME|LEFT|LEN|LOG|LOG10|LOWER|LTRIM|MAX|MIN|MONTH|NCHAR|NEWID|NTILE|NULLIF|OBJECT_DEFINITION|OBJECT_ID|OBJECT_NAME|OBJECT_SCHEMA_NAME|OBJECTPROPERTY|OBJECTPROPERTYEX|OPENDATASOURCE|OPENQUERY|OPENROWSET|OPENXML|ORIGINAL_LOGIN|ORIGINAL_LOGIN|PARSENAME|PATINDEX|PATINDEX|PERMISSIONS|PI|POWER|PUBLISHINGSERVERNAME|PWDCOMPARE|PWDENCRYPT|QUOTENAME|RADIANS|RAND|RANK|REPLICATE|REVERSE|RIGHT|ROUND|ROW_NUMBER|ROWCOUNT_BIG|RTRIM|SCHEMA_ID|SCHEMA_ID|SCHEMA_NAME|SCHEMA_NAME|SCOPE_IDENTITY|SERVERPROPERTY|SESSION_USER|SESSION_USER|SESSIONPROPERTY|SETUSER|SIGN|SignByAsymKey|SignByCert|SIN|SOUNDEX|SPACE|SQL_VARIANT_PROPERTY|SQRT|SQUARE|STATS_DATE|STDEV|STDEVP|STR|STUFF|SUBSTRING|SUM|SUSER_ID|SUSER_NAME|SUSER_SID|SUSER_SNAME|SWITCHOFFSET|SYMKEYPROPERTY|symkeyproperty|sys\.dm_db_index_physical_stats|sys\.fn_builtin_permissions|sys\.fn_my_permissions|SYSDATETIME|SYSDATETIMEOFFSET|SYSTEM_USER|SYSTEM_USER|SYSUTCDATETIME|TAN|TERTIARY_WEIGHTS|TEXTPTR|TODATETIMEOFFSET|TRIGGER_NESTLEVEL|TYPE_ID|TYPE_NAME|TYPEPROPERTY|UNICODE|UPDATE\(|UPPER|USER_ID|USER_NAME|USER_NAME|VAR|VARP|VerifySignedByAsymKey|VerifySignedByCert|XACT_STATE|YEAR)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public ReadOnly SyntaxSQL_Types As New Regex("\b(BIGINT|NUMERIC|BIT|SMALLINT|DECIMAL|SMALLMONEY|INT|TINYINT|MONEY|FLOAT|REAL|DATE|DATETIMEOFFSET|DATETIME2|SMALLDATETIME|DATETIME|TIME|CHAR|VARCHAR|TEXT|NCHAR|NVARCHAR|NTEXT|BINARY|VARBINARY|IMAGE|TIMESTAMP|HIERARCHYID|TABLE|UNIQUEIDENTIFIER|SQL_VARIANT|XML)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    ' define styles
    Public ReadOnly BlueBoldStyle As Style = New TextStyle(Brushes.Blue, Nothing, FontStyle.Bold)
    Public ReadOnly BlueStyle As Style = New TextStyle(Brushes.Blue, Nothing, FontStyle.Regular)
    Public ReadOnly BoldStyle As Style = New TextStyle(Nothing, Nothing, FontStyle.Bold Or FontStyle.Underline)
    Public ReadOnly BrownStyle As Style = New TextStyle(Brushes.Brown, Nothing, FontStyle.Italic)
    Public ReadOnly GrayStyle As Style = New TextStyle(Brushes.Gray, Nothing, FontStyle.Regular)
    Public ReadOnly GreenStyle As Style = New TextStyle(Brushes.Green, Nothing, FontStyle.Italic)
    Public ReadOnly MagentaStyle As Style = New TextStyle(Brushes.Magenta, Nothing, FontStyle.Regular)
    Public ReadOnly MaroonStyle As Style = New TextStyle(Brushes.Maroon, Nothing, FontStyle.Regular)
    Public ReadOnly RedStyle As Style = New TextStyle(Brushes.Red, Nothing, FontStyle.Regular)
    Public ReadOnly BlackStyle As Style = New TextStyle(Brushes.Black, Nothing, FontStyle.Regular)

    Public Sub InitaliseTextBoxSettings(ByVal TextBox As FastColoredTextBox)
        TextBox.CommentPrefix = "--"
        TextBox.LeftBracket = "("c
        TextBox.RightBracket = ")"c
        TextBox.LeftBracket2 = vbNullChar
        TextBox.RightBracket2 = vbNullChar
        TextBox.AutoIndentCharsPatterns = ""
        TextBox.Paddings = New Padding(1, 8, 2, 8)
        TextBox.ReservedCountOfLineNumberChars = 2
        TextBox.AllowMacroRecording = False
    End Sub

    Public Sub SetRangeStyle(ByVal Range As FastColoredTextBoxNS.Range)

        Range.ClearStyle(GreenStyle, RedStyle, MagentaStyle, BlueStyle, MaroonStyle, BrownStyle, GrayStyle)
        Range.SetStyle(GreenStyle, SyntaxSQL_Comment1)
        Range.SetStyle(GreenStyle, SyntaxSQL_Comment2)
        Range.SetStyle(GreenStyle, SyntaxSQL_Comment3)
        Range.SetStyle(RedStyle, SyntaxSQL_String)
        Range.SetStyle(MagentaStyle, SyntaxSQL_Number)
        Range.SetStyle(BrownStyle, SyntaxSQL_Types)
        Range.SetStyle(MaroonStyle, SyntaxSQL_Variables)
        Range.SetStyle(BlueStyle, SyntaxSQL_Statements)
        Range.SetStyle(BlueStyle, SyntaxSQL_Keywords)
        Range.SetStyle(GrayStyle, SyntaxSQL_Keywords2)
        Range.SetStyle(MaroonStyle, SyntaxSQL_Functions)

        Range.ClearFoldingMarkers()
        Range.SetFoldingMarkers("\bBEGIN\b", "\bEND\b", RegexOptions.IgnoreCase)
        Range.SetFoldingMarkers("/\*", "\*/")

    End Sub

    Public Sub HandleTextBoxChanged(ByVal TextBox As FastColoredTextBox)
        SetRangeStyle(TextBox.Range)
    End Sub

    <Extension()>
    Public Function IsEven(i As Integer) As Boolean
        If i Mod 2 = 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    <Extension()>
    Public Function IsOdd(i As Integer) As Boolean
        If Not i Mod 2 = 0 Then
            Return True
        Else
            Return False
        End If
    End Function

    ''' <summary> Return a string with n characters removed from the start of it. </summary>
    <Extension()>
    Public Function TrimStart(ByVal s As String, ByVal n As Integer) As String
        If String.IsNullOrEmpty(s) Then Return ""
        If s.Length = 0 Then Return "" Else Return Microsoft.VisualBasic.Right(s, s.Length - n)
    End Function

    ''' <summary> Return a string with n characters removed from the end of it. </summary>
    <Extension()>
    Public Function TrimEnd(ByVal s As String, ByVal n As Integer) As String
        If String.IsNullOrEmpty(s) Then Return ""
        If s.Length = 0 Then Return "" Else Return (s.Substring(0, s.Length - n))
    End Function

    Public Function ReplaceChar(str As String, charindex As Integer, ByVal NewChar As String)
        Return str.Remove(charindex, 1).Insert(charindex, NewChar)
    End Function

    ''' <summary> Cut off a portion of string form the beginning upto a certain index, and return the removed portion. </summary>
    <Extension()>
    Public Function CutStartOffAtIndex(ByRef s As String, ByVal index As Integer) As String
        Dim v = s.Substring(0, index)
        s = s.Substring(index)
        Return v
    End Function

    Public Function FormatTime(ByVal t As TimeSpan) As String
        If t.Minutes = 0 Then
            Return t.Seconds.ToString("#0") & "." & (t.Milliseconds / 10).ToString("00")
        Else
            Return t.Minutes.ToString("#0") & ":" & t.Seconds.ToString("#0") & "." & (t.Milliseconds / 10).ToString("00")
        End If
    End Function

    ''' <summary> Return the contents of a text file. </summary>
    Public Function ReadFromFile(ByVal PathToFile As String,
                                 Optional sr As StreamReader = Nothing) As String

        Dim strContents As String

        If sr Is Nothing Then
            sr = New StreamReader(PathToFile)
            strContents = sr.ReadToEnd()
            sr.Close()
            sr.Dispose()
        Else
            strContents = sr.ReadToEnd()
        End If

        Return strContents
    End Function

    ''' <summary> Write text to a file. </summary>
    Public Sub WriteToFile(ByVal PathToFile As String,
                           ByVal Data As String,
                           Optional ByVal Append As Boolean = False,
                           Optional ByVal CreateDirectory As Boolean = True,
                           Optional ByVal sw As StreamWriter = Nothing)

        If CreateDirectory Then
            Dim i As New IO.FileInfo(PathToFile)
            Directory.CreateDirectory(i.DirectoryName)
        End If

        If sw Is Nothing Then
            sw = New StreamWriter(PathToFile, Append)
            sw.Write(Data)
            sw.Close()
            sw.Dispose()
        Else
            sw.Write(Data)
        End If

    End Sub

    ''' <summary> Return True if a file or folder exists. FAST - will not hang if the path or network location doesn't exist. Also it shouldn't raise an exception if the syntax isn't legal. </summary>
    Public Function PathExists(ByVal Path As String) As Boolean
        'source: http://stackoverflow.com/questions/2225415/why-is-file-exists-much-slower-when-the-file-does-not-exist
        ' nb: a StringBuilder is required for interops calls that use strings

        If String.IsNullOrWhiteSpace(Path) Then Return False

        Dim builder As New StringBuilder()
        builder.Append(Path)

        Return PathFileExists(builder)

    End Function

    <DllImport("Shlwapi.dll", SetLastError:=True, CharSet:=CharSet.Auto)>
    Private Function PathFileExists(path As StringBuilder) As Boolean
    End Function

    Public Sub ExportToExcel(ByVal Data As List(Of List(Of String)))

        ' create a temp file
        Dim TempFolder As String = System.IO.Path.GetTempPath & "fm-odbc-debugger"
        IO.Directory.CreateDirectory(TempFolder)

        ' ensure the file name is unique
        Static FileNumber As Long = 0
        Dim TempFile As String

        Do
            TempFile = $"{TempFolder}\export-{FileNumber}.csv"
            FileNumber += 1
        Loop While System.IO.File.Exists(TempFile)

        ' write csv data to file
        Using stream = File.OpenWrite(TempFile)
            Using writer = New StreamWriter(stream, Text.Encoding.UTF8)

                Using csv = New CsvHelper.CsvWriter(writer, Globalization.CultureInfo.InvariantCulture)

                    For Each row In Data

                        For Each i In row
                            csv.WriteField(i)
                        Next

                        csv.NextRecord()
                    Next

                End Using

            End Using
        End Using

        Process.Start(TempFile)

    End Sub

    Private Function GetFmOdbcDriverVersion() As String

        Dim SYSTEM_DRIVE As String = Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Machine)
        Dim PATH_TO_DRIVER As String = SYSTEM_DRIVE & "\System32\fmodbc64.dll"

        If File.Exists(PATH_TO_DRIVER) Then
            Dim myFileVersionInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(PATH_TO_DRIVER)
            Return "v" & myFileVersionInfo.FileMajorPart & "." & myFileVersionInfo.FileMinorPart & "." & myFileVersionInfo.FileBuildPart & "." & myFileVersionInfo.FilePrivatePart
        End If

        Return "driver not installed"

    End Function

End Module

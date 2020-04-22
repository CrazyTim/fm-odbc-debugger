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

    ' syntax regex
    Public ReadOnly SyntaxRegex_String As New Regex("""""|''|"".*?[^\\]""|'.*?[^\\]'", RegexOptions.Compiled)
    Public ReadOnly SyntaxRegex_Number As New Regex("\b\d+[\.]?\d*([eE]\-?\d+)?\b", RegexOptions.Compiled)
    Public ReadOnly SyntaxRegex_Comment1 As New Regex("--.*$", RegexOptions.Multiline Or RegexOptions.Compiled)
    Public ReadOnly SyntaxRegex_Comment2 As New Regex("(/\*.*?\*/)|(/\*.*)", RegexOptions.Singleline Or RegexOptions.Compiled)
    Public ReadOnly SyntaxRegex_Comment3 As New Regex("(/\*.*?\*/)|(.*\*/)", RegexOptions.Singleline Or RegexOptions.RightToLeft Or RegexOptions.Compiled)
    Public ReadOnly SyntaxRegex_Keywords As New Regex("\b(ALTER|CATALOG|CREATE|DELETE|DROP|EXEC|EXECUTE|FROM|IN|INSERT|INTO|OPTION|OUTPUT|SELECT|SET|TRUNCATE|WHERE|WITH|ADD|ARE|AS|ASC|AT|BY|CASE|COLUMN|COMMIT|CONNECT|CURRENT|DEFAULT|DESC|DISTINCT|ELSE|END|EXTERNAL|FETCH|FIRST|FOR|GLOBAL|GO|GOTO|GROUP|HAVING|INDEX|KEY|NO|OF|OFFSET|ON|ONLY|OPEN|ORDER|PERCENT|PRIMARY|ROW|ROWS|TABLE|THEN|TIES|TO|UNION|UNIQUE|UPDATE|VALUE|VALUES|WHEN)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public ReadOnly SyntaxRegex_Operators As New Regex("\b(AND|OR|INNER JOIN|LEFT OUTER JOIN|LIKE|NOT|IS|NULL|BETWEEN|IN|EXISTS|ANY|ALL)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public ReadOnly SyntaxRegex_SpecialKeywords As New Regex("\b(FileMaker_Tables|FileMaker_Fields|ROWMODID|ROWID)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public ReadOnly SyntaxRegex_Functions As New Regex("\b(ABS|ATAN|ATAN2|AVE|AVG|CAST|CEIL|CEILING|CHAR|CHR|CHARACTER_LENGTH|CHAR_LENGTH|COALESCE|CONVERT|COUNT|CURDATE|CURRENT_DATE|CURRENT_TIME|CURRENT_TIMESTAMP|CURRENT_USER|CURTIME|CURTIMESTAMP|DATE|DATEVAL|DAY|DAYNAME|DAYOFWEEK|DEG|DEGREES|EXP|FLOOR|HOUR|INT|LEFT|LENGTH|LOWER|LN|LOG|LTRIM|MAX|MIN|MINUTE|MOD|MONTH|MONTHNAME|NULLIF|NUMVAL|PI|PAD|RADIANS|RIGHT|ROUND|RTRIM|SECOND|SIGN|SIN|SPACE|SQRT|STRVAL|SUBSTR|SUBSTRING|SUM|TAN|TIMESTAMPVAL|TIMEVAL|TODAY|TRIM|UPPER|USER|USERNAME|YEAR)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public ReadOnly SyntaxRegex_Types As New Regex("\b(NUMERIC|DECIMAL|INT|DATE|TIME|TIMESTAMP|VARCHAR|CHARACTER VARYING|BLOB|VARBINARY|LONGVARBINARY|BINARY VARYING)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    ''' <summary> These keywords don't do anything, but they still can't be used as identifiers </summary>
    Public ReadOnly SyntaxRegex_ReservedKeywords As New Regex("\b(BIT|BIT_LENGTH|BOOLEAN|DEC|FLOAT|INTEGER|NCHAR|REAL|SMALLINT|TIMEZONE_HOUR|TIMEZONE_MINUTE|WHENEVER|WORK|WRITE|ZONE|VIEW|USAGE|USING|UNKNOWN|TRAILING|TRANSACTION|TRANSLATE|TRANSLATION|TRUE|TEMPORARY|SCHEMA|SCROLL|SECTION|SESSION|SESSION_USER|SIZE|SOME|SQL|SQLCODE|SQLERROR|SQLSTATE|SYSTEM_USER|PRIOR|PRIVILEGES|PROCEDURE|PUBLIC|READ|REFERENCES|RELATIVE|RESTRICT|REVOKE|ROLLBACK|POSITION|PRECISION|PREPARE|PRESERVE|OUTER|OVERLAPS|PART|PARTIAL|OCTET_LENGTH|LANGUAGE|LAST|LEADING|LEVEL|LOCAL|MATCH|MODULE|NAMES|NATIONAL|NATURAL|NEXT|INDICATOR|INITIALLY|INPUT|INSENSITIVE|INTERSECT|INTERVAL|ISOLATION|IDENTITY|IMMEDIATE|GRANT|FOREIGN|FOUND|FULL|GET|EXTRACT|FALSE|ESCAPE|EVERY|EXCEPT|EXCEPTION|DOMAIN|DOUBLE|DESCRIBE|DESCRIPTOR|DIAGNOSTICS|DISCONNECT|DEFERRABLE|DEFERRED|DECLARE|DEALLOCATE|CURSOR|CORRESPONDING|CONTINUE|CONSTRAINT|CONSTRAINTS|CONNECTION|COLLATION|COLLATE|CLOSE|CHECK|CASCADED|CASCADE|BOTH|ASSERTION|AUTHORIZATION|ALLOCATE|ABSOLUTE|ACTION|BEGIN|END_EXEC|INNER|EXCEPTION|JOIN|CROSS)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    ' styles
    Public ReadOnly BlueStyle As Style = New TextStyle(Brushes.Blue, Nothing, FontStyle.Regular)
    Public ReadOnly GrayStyle As Style = New TextStyle(Brushes.Gray, Nothing, FontStyle.Regular)
    Public ReadOnly GrayStrikeStyle As Style = New TextStyle(Brushes.Gray, Nothing, FontStyle.Strikeout)
    Public ReadOnly GreenStyle As Style = New TextStyle(Brushes.Green, Nothing, FontStyle.Regular)
    Public ReadOnly MagentaStyle As Style = New TextStyle(Brushes.Magenta, Nothing, FontStyle.Regular)
    Public ReadOnly MaroonStyle As Style = New TextStyle(Brushes.Maroon, Nothing, FontStyle.Regular)
    Public ReadOnly RedStyle As Style = New TextStyle(Brushes.Red, Nothing, FontStyle.Regular)

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

        Range.ClearStyle(GreenStyle, RedStyle, MagentaStyle, BlueStyle, MaroonStyle, GrayStyle)
        Range.SetStyle(GreenStyle, SyntaxRegex_Comment1)
        Range.SetStyle(GreenStyle, SyntaxRegex_Comment2)
        Range.SetStyle(GreenStyle, SyntaxRegex_Comment3)
        Range.SetStyle(RedStyle, SyntaxRegex_String)
        Range.SetStyle(GrayStrikeStyle, SyntaxRegex_ReservedKeywords)
        Range.SetStyle(MagentaStyle, SyntaxRegex_SpecialKeywords)
        Range.SetStyle(GrayStyle, SyntaxRegex_Operators)
        Range.SetStyle(BlueStyle, SyntaxRegex_Keywords)
        Range.SetStyle(MagentaStyle, SyntaxRegex_Functions)
        Range.SetStyle(MaroonStyle, SyntaxRegex_Types)

        Range.ClearFoldingMarkers()
        Range.SetFoldingMarkers("\bBEGIN\b", "\bEND\b", RegexOptions.IgnoreCase)
        Range.SetFoldingMarkers("/\*", "\*/")

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

Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text
Imports FastColoredTextBoxNS
Imports FileMakerOdbcDebugger.Util.Json
Imports FileMakerOdbcDebugger.Util.Security
Imports System.ComponentModel

Module Main

    Public ReadOnly Path_AppData As String = My.Application.GetEnvironmentVariable("APPDATA") & "\FileMaker ODBC Debugger"
    Public ReadOnly Path_Settings As String = Path_AppData & "\settings.json"
    Public ReadOnly Path_LastQuery As String = Path_AppData & "\last-query.sql"

    Public ReadOnly FmOdbcDriverVersion As String = Util.FileMaker.GetOdbcDriverVersion64Bit()

    Public Const MAX_TABS As Integer = 7 ' Todo: there is a bug with the tab controls, and tabs disapear if they extend beyond the form.

    ' Style syntaxes:
    Public ReadOnly StyleSyntax_String As New Regex("""""|''|"".*?[^\\]""|'.*?[^\\]'", RegexOptions.Compiled)
    Public ReadOnly StyleSyntax_Number As New Regex("\b\d+[\.]?\d*([eE]\-?\d+)?\b", RegexOptions.Compiled)
    Public ReadOnly StyleSyntax_Comment1 As New Regex("--.*$", RegexOptions.Multiline Or RegexOptions.Compiled)
    Public ReadOnly StyleSyntax_Comment2 As New Regex("(/\*.*?\*/)|(/\*.*)", RegexOptions.Singleline Or RegexOptions.Compiled)
    Public ReadOnly StyleSyntax_Comment3 As New Regex("(/\*.*?\*/)|(.*\*/)", RegexOptions.Singleline Or RegexOptions.RightToLeft Or RegexOptions.Compiled)
    Public ReadOnly StyleSyntax_Keywords As New Regex("\b(ALTER|CATALOG|CREATE|DELETE|DROP|EXEC|EXECUTE|FROM|IN|INSERT|INTO|OPTION|OUTPUT|SELECT|SET|TRUNCATE|WHERE|WITH|ADD|ARE|AS|ASC|AT|BY|CASE|COLUMN|COMMIT|CONNECT|CURRENT|DEFAULT|DESC|DISTINCT|ELSE|END|EXTERNAL|FETCH|FIRST|FOR|GLOBAL|GO|GOTO|GROUP|HAVING|INDEX|KEY|NO|OF|OFFSET|ON|ONLY|OPEN|ORDER|PERCENT|PRIMARY|ROW|ROWS|TABLE|THEN|TIES|TO|UNION|UNIQUE|UPDATE|VALUE|VALUES|WHEN)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public ReadOnly StyleSyntax_Operators As New Regex("\b(AND|OR|INNER JOIN|LEFT OUTER JOIN|LIKE|NOT|IS|NULL|BETWEEN|IN|EXISTS|ANY|ALL)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public ReadOnly StyleSyntax_SpecialKeywords As New Regex("\b(FileMaker_Tables|FileMaker_Fields|ROWMODID|ROWID)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public ReadOnly StyleSyntax_Functions As New Regex("\b(ABS|ATAN|ATAN2|AVE|AVG|CAST|CEIL|CEILING|CHAR|CHR|CHARACTER_LENGTH|CHAR_LENGTH|COALESCE|CONVERT|COUNT|CURDATE|CURRENT_DATE|CURRENT_TIME|CURRENT_TIMESTAMP|CURRENT_USER|CURTIME|CURTIMESTAMP|DATE|DATEVAL|DAY|DAYNAME|DAYOFWEEK|DEG|DEGREES|EXP|FLOOR|GETAS|HOUR|INT|LEFT|LENGTH|LOWER|LN|LOG|LTRIM|MAX|MIN|MINUTE|MOD|MONTH|MONTHNAME|NULLIF|NUMVAL|PI|PAD|RADIANS|RIGHT|ROUND|RTRIM|SECOND|SIGN|SIN|SPACE|SQRT|STRVAL|SUBSTR|SUBSTRING|SUM|TAN|TIMESTAMPVAL|TIMEVAL|TODAY|TRIM|UPPER|USER|USERNAME|YEAR)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    Public ReadOnly StyleSyntax_Types As New Regex("\b(NUMERIC|DECIMAL|INT|DATE|TIME|TIMESTAMP|VARCHAR|CHARACTER VARYING|BLOB|VARBINARY|LONGVARBINARY|BINARY VARYING)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)
    ''' <summary> These keywords don't do anything in FileMaker, but they still can't be used as identifiers. </summary>
    Public ReadOnly StyleSyntax_ReservedKeywords As New Regex("\b(BIT|BIT_LENGTH|BOOLEAN|DEC|FLOAT|INTEGER|NCHAR|REAL|SMALLINT|TIMEZONE_HOUR|TIMEZONE_MINUTE|WHENEVER|WORK|WRITE|ZONE|VIEW|USAGE|USING|UNKNOWN|TRAILING|TRANSACTION|TRANSLATE|TRANSLATION|TRUE|TEMPORARY|SCHEMA|SCROLL|SECTION|SESSION|SESSION_USER|SIZE|SOME|SQL|SQLCODE|SQLERROR|SQLSTATE|SYSTEM_USER|PRIOR|PRIVILEGES|PROCEDURE|PUBLIC|READ|REFERENCES|RELATIVE|RESTRICT|REVOKE|ROLLBACK|POSITION|PRECISION|PREPARE|PRESERVE|OUTER|OVERLAPS|PART|PARTIAL|OCTET_LENGTH|LANGUAGE|LAST|LEADING|LEVEL|LOCAL|MATCH|MODULE|NAMES|NATIONAL|NATURAL|NEXT|INDICATOR|INITIALLY|INPUT|INSENSITIVE|INTERSECT|INTERVAL|ISOLATION|IDENTITY|IMMEDIATE|GRANT|FOREIGN|FOUND|FULL|GET|EXTRACT|FALSE|ESCAPE|EVERY|EXCEPT|EXCEPTION|DOMAIN|DOUBLE|DESCRIBE|DESCRIPTOR|DIAGNOSTICS|DISCONNECT|DEFERRABLE|DEFERRED|DECLARE|DEALLOCATE|CURSOR|CORRESPONDING|CONTINUE|CONSTRAINT|CONSTRAINTS|CONNECTION|COLLATION|COLLATE|CLOSE|CHECK|CASCADED|CASCADE|BOTH|ASSERTION|AUTHORIZATION|ALLOCATE|ABSOLUTE|ACTION|BEGIN|END_EXEC|INNER|EXCEPTION|JOIN|CROSS)\b", RegexOptions.IgnoreCase Or RegexOptions.Compiled)

    ' Style colors:
    Public ReadOnly StyleColor_Blue As Style = New TextStyle(Brushes.Blue, Nothing, FontStyle.Regular)
    Public ReadOnly StyleColor_Gray As Style = New TextStyle(Brushes.Gray, Nothing, FontStyle.Regular)
    Public ReadOnly StyleColor_GrayStrike As Style = New TextStyle(Brushes.Gray, Nothing, FontStyle.Strikeout)
    Public ReadOnly StyleColor_Green As Style = New TextStyle(Brushes.Green, Nothing, FontStyle.Regular)
    Public ReadOnly StyleColor_Magenta As Style = New TextStyle(Brushes.Magenta, Nothing, FontStyle.Regular)
    Public ReadOnly StyleColor_Maroon As Style = New TextStyle(Brushes.Maroon, Nothing, FontStyle.Regular)
    Public ReadOnly StyleColor_Red As Style = New TextStyle(Brushes.Red, Nothing, FontStyle.Regular)

#Region "Settings"

    Public Class Settings
        Public Query As String = ""
        Public ServerAddress As String = "localhost"
        Public DatabaseName As String = ""
        Public Username As String = ""
        Public Password As String = ""
        Public SelectedDriver As SqlControl.SelectedDriverIndex = 1
        Public RowsToReturn As Integer = 1000
        Public ConnectionString As String = ""
    End Class

    ''' <summary> Save settings to a json file and encrypt. </summary>
    Public Sub SaveSettings(Tab As TabControl)

        Dim Settings As New List(Of Settings)

        For Each t As TabPage In Tab.TabPages

            Dim sqlc As SqlControl = t.Form.Controls(0)

            Dim s As New Settings With {
                .Query = sqlc.SQLText,
                .ServerAddress = sqlc.ServerAddress,
                .DatabaseName = sqlc.DatabaseName,
                .Username = sqlc.Username,
                .Password = sqlc.Password,
                .SelectedDriver = sqlc.SelectedDriver,
                .RowsToReturn = sqlc.RowsToReturn,
                .ConnectionString = sqlc.ConnectionString
            }

            Settings.Add(s)

        Next

        IO.File.WriteAllText(Path_Settings, Settings.ToJson().Encrypt())

    End Sub

    ''' <summary> Load settings from a json file and decrypt. </summary>
    Public Function LoadSettings() As List(Of Settings)

        Dim Settings As New List(Of Settings)

        If PathExists(Path_Settings) Then

            Dim s As String = IO.File.ReadAllText(Path_Settings).Derypt()
            Settings = s.FromJson(GetType(List(Of Settings)))

        End If

        Return Settings

    End Function

    Public Function CreateDivider(Optional Dock As DockStyle = DockStyle.Bottom) As Label
        Dim l = New Label
        l.Height = 1
        l.BackColor = Color.Silver
        l.Dock = Dock
        l.BringToFront()
        Return l
    End Function

#End Region

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

        Range.ClearStyle(StyleColor_Green, StyleColor_Red, StyleColor_Magenta, StyleColor_Blue, StyleColor_Maroon, StyleColor_Gray)
        Range.SetStyle(StyleColor_Green, StyleSyntax_Comment1)
        Range.SetStyle(StyleColor_Green, StyleSyntax_Comment2)
        Range.SetStyle(StyleColor_Green, StyleSyntax_Comment3)
        Range.SetStyle(StyleColor_Red, StyleSyntax_String)
        Range.SetStyle(StyleColor_Magenta, StyleSyntax_SpecialKeywords)
        Range.SetStyle(StyleColor_Gray, StyleSyntax_Operators)
        Range.SetStyle(StyleColor_Blue, StyleSyntax_Keywords)
        Range.SetStyle(StyleColor_Magenta, StyleSyntax_Functions)
        Range.SetStyle(StyleColor_Maroon, StyleSyntax_Types)
        Range.SetStyle(StyleColor_GrayStrike, StyleSyntax_ReservedKeywords) ' Do last, so we don't override partial syntax of something that is actually supported.

        Range.ClearFoldingMarkers()
        Range.SetFoldingMarkers("\bBEGIN\b", "\bEND\b", RegexOptions.IgnoreCase)
        Range.SetFoldingMarkers("/\*", "\*/")

    End Sub

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

End Module

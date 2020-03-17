Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports System.Xml.Serialization
Imports System.IO
Imports Microsoft.Office.Interop

Module Module1

    Public AppDataDir As String = My.Application.GetEnvironmentVariable("APPDATA") & "\FileMaker ODBC Debugger"
    Public FMODBCDriverVersion As String = ""

    Public ReadOnly FM_ReservedKeyWords As New List(Of String) From {"ABSOLUTE", "ACTION", "ADD", "ALL", "ALLOCATE", "ALTER", "AND", "ANY", "ARE", "AS", "ASC", "ASSERTION", "AT",
                                                          "AUTHORIZATION", "AVG", "BEGIN", "BETWEEN", "BINARY", "BIT", "BIT_LENGTH", "BLOB", "BOOLEAN", "BOTH", "BY",
                                                          "CASCADE", "CASCADED", "CASE", "CAST", "CATALOG", "CHAR", "CHARACTER", "CHARACTER_LENGTH", "CHAR_LENGTH",
                                                          "CHECK", "CHR", "CLOSE", "COALESCE", "COLLATE", "COLLATION", "COLUMN COMMIT", "CONNECT", "CONNECTION", "CONSTRAINT",
                                                          "CONSTRAINTS", "CONTINUE", "CONVERT", "CORRESPONDING", "COUNT", "CREATE", "CROSS", "CURDATE", "CURRENT", "CURRENT_DATE",
                                                          "CURRENT_TIME", "CURRENT_TIMESTAMP", "CURRENT_USER", "CURSOR", "CURTIME", "CURTIMESTAMP", "DATE", "DATEVAL", "DAY",
                                                          "DAYNAME", "DAYOFWEEK", "DEALLOCATE", "DEC", "DECIMAL", "DECLARE", "DEFAULT", "DEFERRABLE", "DEFERRED", "DELETE",
                                                          "DESC", "DESCRIBE", "DESCRIPTOR", "DIAGNOSTICS", "DISCONNECT", "DISTINCT", "DOMAIN", "DOUBLE", "DROP", "ELSE",
                                                          "END", "END_EXEC", "ESCAPE", "EVERY", "EXCEPT", "EXCEPTION", "EXEC", "EXECUTE", "EXISTS", "EXTERNAL", "EXTRACT",
                                                          "FALSE", "FETCH", "FIRST", "FLOAT", "FOR", "FOREIGN", "FOUND", "FROM", "FULL", "GET", "GLOBAL", "GO", "GOTO",
                                                          "GRANT", "GROUP", "HAVING", "HOUR", "IDENTITY", "IMMEDIATE", "IN", "INDEX", "INDICATOR", "INITIALLY", "INNER",
                                                          "INPUT", "INSENSITIVE", "NSERT", "INT", "INTEGER", "INTERSECT", "INTERVAL", "INTO", "IS", "ISOLATION", "JOIN",
                                                          "KEY", "LANGUAGE", "LAST", "LEADING", "LEFT", "LENGTH", "LEVEL", "LIKE", "LOCAL", "LONGVARBINARY", "LOWER", "LTRIM",
                                                          "MATCH", "MAX", "MIN", "MINUTE", "MODULE", "MONTH", "MONTHNAME", "NAMES", "NATIONAL", "NATURAL", "NCHAR", "NEXT",
                                                          "NO", "NOT", "NULL", "NULLIF", "NUMERIC", "NUMVAL", "OCTET_LENGTH", "OF", "ON", "ONLY", "OPEN", "OPTION", "OR",
                                                          "ORDER", "OUTER", "OUTPUT", "OVERLAPS", "PAD", "PART", "PARTIAL", "POSITION", "PRECISION", "PREPARE", "PRESERVE",
                                                          "PRIMARY", "PRIOR", "PRIVILEGES", "PROCEDURE", "PUBLIC", "READ", "REAL", "REFERENCES", "RELATIVE", "RESTRICT",
                                                          "REVOKE", "RIGHT", "ROLLBACK", "ROUND", "ROWID", "ROWS", "RTRIM", "SCHEMA", "SCROLL", "SECOND", "SECTION", "SELECT",
                                                          "SESSION", "SESSION_USER", "SET", "SIZE", "SMALLINT", "SOME", "SPACE", "SQL", "SQLCODE", "SQLERROR", "SQLSTATE", "STRVAL",
                                                          "SUBSTRING", "SUM", "SYSTEM_USER", "TABLE", "TEMPORARY", "THEN", "TIME", "TIMESTAMP", "TIMESTAMPVAL", "TIMEVAL",
                                                          "TIMEZONE_HOUR", "TIMEZONE_MINUTE", "TO", "TODAY", "TRAILING", "TRANSACTION", "TRANSLATE", "TRANSLATION", "TRIM",
                                                          "TRUE", "UNION", "UNIQUE", "UNKNOWN", "UPDATE", "UPPER", "USAGE", "USER", "USERNAME", "USING", "USAGE", "USER", "USERNAME",
                                                          "USING", "VALUE", "VALUES", "VARBINARY", "VARCHAR", "VARYING", "VIEW", "WHEN", "WHENEVER", "WHERE", "WITH", "WORK",
                                                          "WRITE", "YEAR", "ZONE"}


    ''' <summary> Execute a SELECT SQL query against an ODBC connection </summary>
    Public Function SQL_SELECT(ByVal QueryString As String, ByRef ExeTime As TimeSpan, ByRef StreamTime As TimeSpan, ByRef cn As Odbc.OdbcConnection) As Collection

        ' - handle dropouts and deadlocks by repeating query up to 3 times.
        ' - Handle NULL connections and existing connections (open existing connections if they are closed).
        ' - If sucessfull, return a Collection of ArrayLists (one for each record), otherwise return NOTHING.

        ' validate
        If QueryString.Count = 0 Then Return New Collection

        Dim Col As New Collection
        Dim first As Boolean = False


        Using cmd As New Odbc.OdbcCommand(QueryString, cn)

            Dim sw As New System.Diagnostics.Stopwatch
            sw.Restart()

            Using reader = cmd.ExecuteReader()

                ExeTime = sw.Elapsed
                sw.Restart()

                Dim s As New ArrayList

                If first = False Then ' for the first loop, add the names of the fields to the result

                    For i = 0 To reader.FieldCount - 1
                        s.Add(reader.GetName(i))
                    Next

                    first = True
                    Col.Add(s)
                End If
                Dim cc As Integer = 0
                Do While reader.Read()
                    s = New ArrayList
                    cc += 1
                    For i = 0 To reader.FieldCount - 1
                        If reader.GetValue(i).GetType = GetType(TimeSpan) Then
                            ' filemake "time" fields return timepsans, 
                            ' so here we convert timespans into the more supported date datatype
                            Dim d As New Date
                            Dim t As TimeSpan = reader.GetValue(i)
                            s.Add(d.Add(t))

                        ElseIf Not reader.IsDBNull(i) Then
                            s.Add(reader.GetValue(i))

                        Else
                            s.Add(Nothing)
                        End If
                    Next

                    Col.Add(s)
                Loop

                StreamTime = sw.Elapsed

            End Using
        End Using

        'Catch ave As AccessViolationException
        '    'Log.WriteLine(ave)
        '    'Throw New BSAException("Unknown database error", Nothing, 6)

        'Catch odbcex As Odbc.OdbcException
        '    'Log.WriteLine(odbcex)

        '    'If odbcex.Message.Contains("40001") Then ' DEADLOCKED!
        '    '    RepeatTransaction_Flag = True ' try again
        '    'ElseIf odbcex.Message.Contains("timeout expired") Then
        '    '    Throw New BSAException("The database is taking too long to respond", Nothing, 5)

        '    'ElseIf odbcex.Message.Contains("ERROR [08S01] [FileMaker][FileMaker]") Then
        '    '    ' the FM listener has crashed, or there is a generic datasource error. So need to close the connection and try again.
        '    '    Log.WriteLine("WARNING: LOST CONNECTION TO FM")
        '    '    Log.WriteLine(odbcex)
        '    '    cn.Close()
        '    '    RepeatTransaction_Flag = True ' try again
        '    'Else
        '    '    Throw New BSAException("Unknown database error", Nothing, 6)
        '    'End If

        'Catch exIOE As InvalidOperationException ' ie: the connection to thw SQL server was lost and the connection was disabled

        '    '' dispose of old connection
        '    'If cn.State = System.Data.ConnectionState.Open Then cn.Close()

        '    'If cn Is G_CN_SYRINX Then
        '    '    Log.WriteLine("WARNING: LOST CONNECTION TO SQL")
        '    '    cn.Dispose()
        '    '    cn = Nothing
        '    '    G_CN_SYRINX = New Odbc.OdbcConnection(ODBC_Syrinx)
        '    'ElseIf cn Is G_CN_SCHEDULER Then
        '    '    cn.Dispose()
        '    '    cn = Nothing
        '    '    G_CN_SCHEDULER = New Odbc.OdbcConnection(ODBC_Scheduler)
        '    'Else
        '    '    Throw New BSAException("unable to reset SQL connection", Nothing, 6)
        '    'End If

        '    'RepeatTransaction_Flag = True ' try again

        'Catch ex As Exception
        '    'Log.WriteLine(ex)
        '    'Log.WriteLine("SQL_SELECT:" & vbNewLine & QueryString)
        '    'Throw New BSAException("Unknown database error", Nothing, 6)
        'End Try

        Return Col

    End Function


    ''' <summary> Execute a series of SQL transactions against an ODBC connection. If any of them fail, the transaction is rolled back. If no connection is provided, a default connection is used. </summary>
    Public Sub SQL_TRANS(ByVal Queries As Collection, ByVal cn As Odbc.OdbcConnection)

        ' validate
        If Queries.Count = 0 Then Exit Sub

        Dim CurrentQryString As String = ""

        ' execute transaction 
        Try
            Using cmd As New Odbc.OdbcCommand()
                cmd.Connection = cn                     ' Set the Connection to the new OdbcConnection.

                Using t As Odbc.OdbcTransaction = cn.BeginTransaction() ' Start a local transaction.
                    cmd.Transaction = t           ' Assign transaction object for a pending local transaction.

                    For Each QueryString As String In Queries
                        CurrentQryString = QueryString
                        If QueryString <> "" Then        ' skip if empty string
                            cmd.CommandText = QueryString
                            cmd.ExecuteNonQuery()
                        End If
                    Next

                    Try
                        t.Commit()
                        Dim f = ""
                        'Catch odbcex As Odbc.OdbcException
                    Catch ex As Exception
                        t.Rollback()
                        MsgBox(ex.Message)
                    End Try

                End Using
            End Using

        Catch ex As Exception
            MsgBox(ex.Message)
        End Try


    End Sub


    ''' <summary> Insert double-quotes around words starting with "_" (which is illegal in odbc). </summary>
    Public Function FM_EscUnderscoreIdentifiyers(ByVal sql As String) As String

        If String.IsNullOrEmpty(sql) Then Return ("")

        Dim i As Integer = 1
        Dim Flag As Boolean = False
        Dim StartIndex As Integer = -1
        Dim EndIndex As Integer = -1

        ' test first char:
        If sql(0) = "_" Then
            Flag = True
            StartIndex = 0
        End If

        ' test middle chars:
        Do Until i > sql.Length - 1

            If Not Flag Then
                ' DETECT START OF WORD:
                ' nb: the following characters signify the START of a word:
                If sql(i) = "_" And (sql(i - 1) = " " Or sql(i - 1) = "," Or sql(i - 1) = "(" Or sql(i - 1) = "." Or sql(i - 1) = vbTab Or sql(i - 1) = vbCr Or sql(i - 1) = vbLf Or sql(i - 1) = "=") Then
                    Flag = True
                    StartIndex = i
                End If
            Else
                ' DETECT END OF WORD
                ' nb: the following characters signify the END of a word:
                If sql(i) = " " Or sql(i) = "," Or sql(i) = ")" Or sql(i) = "." Or sql(i) = vbTab Or sql(i) = vbCr Or sql(i) = vbLf Or sql(i) = "=" Then
                    EndIndex = i
                    sql = sql.Insert(EndIndex, """")
                    sql = sql.Insert(StartIndex, """")
                    i += 2
                    Flag = False
                End If
            End If

            i += 1
        Loop

        ' test last char:
        If Flag Then
            EndIndex = sql.Length
            sql = sql.Insert(EndIndex, """")
            sql = sql.Insert(StartIndex, """")
        End If

        Return sql
    End Function


    '' split into an array of strings based on ";"
    Public Function SplitSQLTRANS(ByVal str As String) As String()

        ' --------------------------------------------------------
        ' replace transaction identifiers (";") with a special character (Chr(159))
        ' --------------------------------------------------------
        Dim quoteCounter As Integer = 0
        Dim charCount As Int64 = 0
        For i = 0 To str.Length - 1
            Dim s As Char = str(i)
            If s = "'" Then
                If quoteCounter = 1 Then
                    quoteCounter = 0
                Else
                    quoteCounter = 1
                End If
            ElseIf s = ";" Then
                If quoteCounter = 0 Then
                    str = replaceChar(str, i, Chr(159))
                End If
            End If
            charCount += 1
        Next

        ' trim:
        Dim sp = str.Split(Chr(159))
        For i = 0 To sp.Length - 1
            sp(i) = sp(i).Trim
        Next

        Return sp

    End Function

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

    Public Function replaceChar(str As String, charindex As Integer, ByVal NewChar As String)
        Return str.Remove(charindex, 1).Insert(charindex, NewChar)
    End Function

    ''' <summary> Return a string with n characters removed from the start of it. </summary>
    <Extension()> <DebuggerStepThrough()>
    Public Function TrimStart(ByVal s As String, ByVal n As Integer) As String
        If String.IsNullOrEmpty(s) Then Return ""
        If s.Length = 0 Then Return "" Else Return Microsoft.VisualBasic.Right(s, s.Length - n)
    End Function

    ''' <summary> Return a string with n characters removed from the end of it. </summary>
    <Extension()> <DebuggerStepThrough()>
    Public Function TrimEnd(ByVal s As String, ByVal n As Integer) As String
        If String.IsNullOrEmpty(s) Then Return ""
        If s.Length = 0 Then Return "" Else Return (s.Substring(0, s.Length - n))
    End Function

    ''' <summary> Cut off a portion of string form the beginning upto a certain index, and return the removed portion. </summary>
    <Extension()> <DebuggerStepThrough()>
    Public Function CutStartOffAtIndex(ByRef s As String, ByVal index As Integer) As String
        Dim v = s.Substring(0, index)
        s = s.Substring(index)
        Return v
    End Function

    Private Function JoinSQL(ByVal a As ArrayList)
        Dim r As String = ""
        For i = 0 To a.Count - 1
            r &= a(i) & "'"
        Next
        Return r.TrimEnd(1) ' trim off last apostrophe
    End Function

    Private Function SplitSQL(ByVal sql As String) As ArrayList
        Dim i As Integer = 0
        Dim a As New ArrayList
        Do While i < sql.Length

            If (i + 1) > sql.Length - 1 Then
                i += 1
                Continue Do ' skip if last letter
            End If

            If sql(i) = "'" And sql(i + 1) = "'" Then
                i += 2
                Continue Do ' skip 2
            End If

            If sql(i) = "'" Then
                a.Add(sql.CutStartOffAtIndex(i))
                sql = sql.TrimStart(1)
                i = 0
            End If

            i += 1
        Loop
        a.Add(sql)
        Return a
    End Function

    ''' <summary> Change reserved keywords in a SQL query to Uppercase. </summary>
    Public Function BeautifySQL(ByVal SQL As String) As String
        Dim i As Integer
        Dim r As String = ""
        Dim a As New ArrayList


        ' 1) SPLIT
        a = SplitSQL(SQL)


        ' 2) UPPERCASE RESERVED SQL KEYWORDS
        i = 0
        Do While i < a.Count

            ' nb: \b matches a word boundary (the position between a word and a space).
            For Each w As String In FM_ReservedKeyWords
                a(i) = Regex.Replace(a(i), "\b(" & w & ")\b", w, RegexOptions.IgnoreCase)
            Next

            i += 2
        Loop


        ' 3) JOIN BACK INTO ONE STRING
        r = JoinSQL(a)

        Return r

    End Function

    ''' <summary> Prep SQL ready for filemaker ODBC driver to execute. </summary>
    Public Function FM_PrepSQL(ByVal SQL As String) As String

        Dim i As Integer
        Dim r As String = ""
        Dim a As New ArrayList

        Try


            ' 5) REMOVE ANY COMMENTS
            r = Regex.Replace(SQL, "(--.+)", "", RegexOptions.IgnoreCase) ' inline comments
            r = Regex.Replace(r, "\/\*[\s\S]*?\*\/", "", RegexOptions.IgnoreCase) ' multiline comments

            ' 6) SPLIT AGAIN
            a = SplitSQL(r)


            ' 7) CLEAN UP SPACES
            i = 0
            Do While i < a.Count

                ' A) REPLACE DOUBLE BLANK SPACES WITH SINGLE
                While a(i).contains("  ")
                    a(i) = a(i).Replace("  ", " ")
                End While

                ' nb: we preserve new lines so its easier to read when debugging
                ' B) REMOVE SINGLE WHITE SPACES AFTER A NEW LINE
                While a(i).contains(vbCrLf & " ")
                    a(i) = a(i).Replace(vbCrLf & " ", vbCrLf)
                End While

                i += 2
            Loop


            ' 8) Replace CRLF with CR (which is what filemaker interprets as new lines):
            i = 1
            Do While i < a.Count
                a(i) = a(i).Replace(Environment.NewLine, Chr(13))
                a(i) = a(i).Replace(Chr(10), Chr(13))

                i += 2
            Loop


            ' 9) LASTLY, apply special FileMaker syntax: escape any field names starting with "_" with double quotes
            i = 0
            Do While i < a.Count
                a(i) = FM_EscUnderscoreIdentifiyers(a(i))

                i += 1
            Loop


            ' 10) JOIN INTO ONE STRING AGAIN
            r = JoinSQL(a)


            ' 11) OUTPUT FINAL SQL
            Return r.Trim


        Catch ex As Exception
            ' this shouldn't happen
            Return "[error]"
        End Try

    End Function


    ''' <summary> Check for misc errors not reproted by filemakers odbc driver </summary>
    Public Function FM_CheckErrors(ByVal SQL As String) As ArrayList
        Dim a As New ArrayList
        Dim i As Integer
        Dim Errors As New ArrayList
        Dim r1 As Regex
        Dim r2 As Regex

        ' 1) SPLIT
        a = SplitSQL(SQL)


        ' A) check for WHERE errors
        i = 0
        Do While i < a.Count

            r1 = New Regex("WHERE[\s\S]*=[\s\S]*''", RegexOptions.IgnoreCase)
            r2 = New Regex("WHERE[\s\S]*<>[\s\S]*''", RegexOptions.IgnoreCase)
            If r1.Match(SQL).Success Or r2.Match(SQL).Success Then
                Errors.Add("Note: FileMaker stores empty strings as NULL, so using ""WHERE column = ''"" or ""WHERE column <> ''"" on a ""Text"" field will always return 0 results. Use ""IS NULL"" instead.")
                Exit Do
            End If

            i += 2
        Loop


        ' B) check for TRUE FALSE
        i = 0
        Do While i < a.Count

            r1 = New Regex("=\s*?TRUE", RegexOptions.IgnoreCase)
            r2 = New Regex("=\s*?FALSE", RegexOptions.IgnoreCase)
            If r1.Match(SQL).Success Or r2.Match(SQL).Success Then
                Errors.Add("Syntax error: FileMaker ODBC does not support the keywords ""TRUE"" or ""FALSE"". Use ""1"" or ""0"" instead.")
                Exit Do
            End If

            i += 2
        Loop


        ' C) Check to see if there is an even number of "'" in the query, if not, there is a syntax error.
        Dim AposCounter As Integer = 0
        For i = 0 To SQL.Length - 1
            If SQL(i) = "'" Then AposCounter += 1
        Next
        If AposCounter.IsOdd Then
            Errors.Add("Syntax error: one or more apostrophes (""'"") need to be escaped!")
        End If


        ' D Warn about the BETWEEN statement (very slow in FileMaker)
        i = 0
        Do While i < a.Count

            r1 = New Regex("\b(" & "BETWEEN" & ")\b", RegexOptions.IgnoreCase)
            If r1.Match(SQL).Success Then
                Errors.Add("Query contains the keyword ""BETWEEN"" and FileMaker ODBC is VERY slow when comparing dates this way. Use "">="" and ""<="" instead.")
                Exit Do
            End If

            i += 2
        Loop

        Return Errors

    End Function

    ''' <summary> Write text to a file. </summary>
    Public Sub WriteToFile(ByVal strData As String, ByVal FullPath As String, ByVal Append As Boolean)
        Dim objReader As StreamWriter
        objReader = New StreamWriter(FullPath, Append)
        objReader.Write(strData)
        objReader.Close()
        objReader.Dispose()
    End Sub

    Public Function SaveSerialXML(ByVal FileName As String, ByVal DataToSerialize As Object, ByVal objType As Type) As Boolean

        If FileName = "" Or FileName = Nothing Then Return False

        'set up a blank namespace to eliminate unnecessary junk from the xml
        Dim nsBlank As New XmlSerializerNamespaces
        nsBlank.Add("", "")

        'create an object for the xml settings to control how the xml is written and appears
        Dim xSettings As New System.Xml.XmlWriterSettings
        With xSettings
            .Encoding = System.Text.Encoding.UTF8
            .Indent = True
            .NewLineHandling = Xml.NewLineHandling.Entitize
            .NewLineOnAttributes = False
            .ConformanceLevel = Xml.ConformanceLevel.Document
        End With

        Try
            'create the xmlwriter object that will write the file out
            Dim xw As System.Xml.XmlWriter = Xml.XmlWriter.Create(FileName, xSettings)

            'create the xmlserializer that will serialize the object to XML
            Dim writer As New XmlSerializer(objType)

            'now write it out
            writer.Serialize(xw, DataToSerialize, nsBlank)

            'be sure to close it or it will remain open
            xw.Close()

            Return True

        Catch ex As Exception
            Return False
        End Try
    End Function

    Public Function ReadSerialXML(ByVal sFileName As String, ByVal objType As Type) As Object
        Dim fs As FileStream = Nothing
        Dim xs As XmlSerializer = Nothing
        Dim obj As Object = Nothing

        If My.Computer.FileSystem.FileExists(sFileName) Then
            Try
                fs = New FileStream(sFileName, FileMode.Open, FileAccess.Read, FileShare.ReadWrite)
                xs = New XmlSerializer(objType)
                obj = CType(xs.Deserialize(fs), Object)
                fs.Close()
                Return obj

            Catch ex As InvalidOperationException
                Return Nothing
            End Try
        Else
            Return Nothing
        End If
    End Function

    Public Function CreateDir(ByVal Dir As String) As Boolean
        ' CREATE PROGRAM FILES DIR
        If Not DoesDirExist(Dir) Then
            ' create buildsafe temp folder
            Try
                Directory.CreateDirectory(Dir)
                If Not DoesDirExist(Dir) Then
                    MsgBox("There was a problem creating the folder '" & Dir & "', or the environment variable could not be read.")
                    Return False
                End If
            Catch
            End Try
        End If
        Return True
    End Function

    Public Function DoesDirExist(ByVal Path As String) As Boolean
        On Error Resume Next
        Dim f As String = Dir(Path, FileAttribute.Directory)

        If f = "" Then
            Return False
        Else
            Return True
        End If

    End Function

    Public Sub ExportToExcel2(ByVal Data As List(Of List(Of String)))

        ' generate a temp file to save the csv data to
        Dim TempFolder As String = System.IO.Path.GetTempPath & "fm-odbc-debugger"
        IO.Directory.CreateDirectory(TempFolder)

        Static FileNumber As Long = 0
        Dim TempFile As String

        Do

            TempFile = $"{TempFolder}\export-{FileNumber}.csv"
            FileNumber += 1
        Loop While System.IO.File.Exists(TempFile)

        ' write the csv data to file
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

    Public Sub ExportToExcel(ByVal ColumnHeadings As List(Of String), ByVal Data As List(Of List(Of String)))
        ' -----------------------------------------------------------
        ' LOAD EXCEL:
        ' -----------------------------------------------------------

        Dim PathToTemplate As String = AppDomain.CurrentDomain.BaseDirectory.TrimEnd(1) & "/export.xltx"

        Dim xlsApp As New Excel.Application
        Dim xlsBook As Excel.Workbook
        Dim xlsSheet As Excel.Worksheet

        If File.Exists(PathToTemplate) Then
            xlsBook = xlsApp.Workbooks.Open(PathToTemplate)
        Else
            xlsBook = xlsApp.Application.Workbooks.Add()
        End If

        Dim cRow As Integer = 2
        Dim cCol As Integer = 1

        xlsApp.DisplayAlerts = 0 ' prevent dialog alerts, eg: when deleting worksheets
        'xlsApp.Visible = True
        xlsApp.Visible = False
        xlsApp.Application.ScreenUpdating = False

        xlsSheet = xlsBook.Worksheets("Sheet1")
        xlsSheet.Select()

        With xlsSheet

            ' create columns
            For Each cell In Data(0)
                xlsSheet.Columns(2).Insert(Excel.XlDirection.xlToRight, Excel.XlInsertFormatOrigin.xlFormatFromLeftOrAbove)
            Next
            xlsSheet.Columns(2).Delete(Excel.XlDirection.xlToRight)
            xlsSheet.Columns(2).Delete(Excel.XlDirection.xlToRight)

            ' create column headers
            For Each c In ColumnHeadings
                xlsSheet.Cells(1, cCol).Value = c
                cCol += 1
            Next
            cCol = 1

            ' create rows
            For Each row In Data
                For Each cell In row
                    xlsSheet.Cells(cRow, cCol).Value = cell
                    cCol += 1
                Next
                cCol = 1
                cRow += 1

                ' insert new row
                xlsSheet.Range(cRow + 1 & ":" & cRow + 1).Insert(Excel.XlInsertShiftDirection.xlShiftDown)
            Next

        End With

        ' Delete last rows
        xlsSheet.Range(cRow & ":" & cRow).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)
        xlsSheet.Range(cRow & ":" & cRow).Delete(Excel.XlDeleteShiftDirection.xlShiftUp)

        ' autofit contents
        xlsSheet.Cells.EntireColumn.AutoFit()

        xlsSheet.Cells(1, 1).Select()

        xlsApp.Application.ScreenUpdating = True
        xlsApp.Visible = True
        xlsApp.DisplayAlerts = 1

    End Sub

End Module

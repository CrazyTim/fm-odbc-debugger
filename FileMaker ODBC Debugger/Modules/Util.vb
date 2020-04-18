Imports System.Runtime.CompilerServices
Imports System.Text.RegularExpressions
Imports System.Xml.Serialization
Imports System.IO
Imports System.Runtime.InteropServices
Imports System.Text

<DebuggerStepThrough()>
Module Util

    Public ReadOnly AppDataDir As String = My.Application.GetEnvironmentVariable("APPDATA") & "\FileMaker ODBC Debugger"
    Public ReadOnly FmOdbcDriverVersion As String = GetFmOdbcDriverVersion()

    Public Const MAX_TABS As Integer = 7

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

Imports System.Text.RegularExpressions
Imports System.ComponentModel
Imports FileMakerOdbcDebugger.Util.Extensions

'<DebuggerStepThrough()>
Public Module Sql

    ''' <summary> Insert double-quotes around words starting with "_" (which is illegal in odbc, but not for FileMaker). </summary>
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

    ''' <summary> Prep query for filemaker ODBC driver to execute. </summary>
    Public Function FM_PrepSQL(ByVal Query As String) As String

        ' Remove any comments
        Query = Regex.Replace(Query, "(--.+)", "", RegexOptions.IgnoreCase) ' inline comments
        Query = Regex.Replace(Query, "\/\*[\s\S]*?\*\/", "", RegexOptions.IgnoreCase) ' multiline comments

        ' Split
        Dim Split = FileMakerOdbcDebugger.Util.Sql.SplitQueryByStringLiteral(Query)

        ' Loop over Statements
        ' Clean spaces
        For i = 0 To Split.Parts.Count - 1 Step 2

            Dim s = Split.Parts(i)

            ' Replace double blank spaces with a single space
            While s.Contains("  ")
                Split.Parts(i) = s.Replace("  ", " ")
            End While

            ' Remove single white spaces after a new line
            ' Note: we preserve new lines so its still easy to read when debugging
            While s.Contains(vbCrLf & " ")
                Split.Parts(i) = s.Replace(vbCrLf & " ", vbCrLf)
            End While

        Next

        ' Loop over String Literals
        ' Replace newline characters with CR (FileMaker only uses CR for newlines)
        For i = 1 To Split.Parts.Count - 1 Step 2

            Dim s = Split.Parts(i)

            Split.Parts(i) = s.Replace(vbCrLf, Chr(13))
            Split.Parts(i) = s.Replace(Chr(10), Chr(13))

        Next

        ' Loop over Statements
        ' Surround any field names that start with an underscore (_) with double quotes (")
        For i = 0 To Split.Parts.Count - 1 Step 2

            Dim s = Split.Parts(i)

            Split.Parts(i) = FM_EscUnderscoreIdentifiyers(s)

        Next

        Return Split.Rejoin.Trim

    End Function

    Private Function GetTableNames(ConnectionString As String) As Object

        Try

            Dim MetaData As New List(Of String)
            Dim TableNames As New List(Of String)
            Dim Err As String = ""

            Using cn As New Odbc.OdbcConnection(ConnectionString)
                cn.Open()

                Dim all As DataTable = cn.GetSchema("MetaDataCollections")

                For Each r In all.Rows
                    MetaData.Add(r("CollectionName").ToString)
                    'Console.WriteLine("Meta: " & r("CollectionName").ToString)
                Next

                If Not MetaData.Contains("Tables") Then
                    Err = "Error: driver doesn't support table list"
                    Return Err
                End If

                Dim tables As DataTable = cn.GetSchema("Tables")

                For Each r In tables.Rows
                    TableNames.Add(r("TABLE_NAME").ToString)
                    'Console.WriteLine("Table: " & r("TABLE_NAME").ToString)
                Next

            End Using

            Return TableNames

        Catch ex As Exception
            Return ex.Message
        End Try

    End Function

End Module

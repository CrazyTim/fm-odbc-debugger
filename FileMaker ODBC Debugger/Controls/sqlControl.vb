Imports System.ComponentModel
Imports FastColoredTextBoxNS

Public Class sqlControl

    Public Event DatabaseNameChanged(ByVal e As TextChangedArgs)
    Public Event BeginSQLExecute()

    Private LastQryCount As Integer = 0
    Private LastQrySQL As String = ""

    Private _bw As New BackgroundWorker()

    Private _bw_SQL As String = ""              '    the query we are going to run
    Private _bw_SQL_Formatted As String = ""        ' the query after it has been formatted
    Private _bw_Error As String = ""
    Private _bw_Server As String = ""               ' server name/address
    Private _bw_Database As String = ""             ' db name
    Private _bw_Username As String = ""
    Private _bw_Password As String = ""
    Private _bw_Data As New Collection              ' result
    Private _bw_RecordsAffected As Integer = -1
    Private _bw_Time_Conn As New TimeSpan           ' connection time
    Private _bw_Time_Exe As New TimeSpan            ' execution time   
    Private _bw_Time_Stream As New TimeSpan         ' stream time
    Private _bw_DriverName As String = ""
    Private _bw_RowsToReturn As Long
    Private _bw_ConnectionString As String = ""

    Private Sub Me_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        txtErrorText.Dock = DockStyle.Fill
        txtErrorText.Text = ""
        lblStatus.Text = ""
        lblDurationConnect.Visible = False
        lblDurationExecute.Visible = False
        lblSyntaxWarning.Visible = False
        txtSQL.Select()

        InitaliseTextBoxSettings(txtSQL)

        DisableSearchBox()

        InitaliseBackgroundWorker()

    End Sub

    Public Sub InitaliseBackgroundWorker()

        _bw.WorkerSupportsCancellation = True
        _bw.WorkerReportsProgress = True

        AddHandler _bw.DoWork, AddressOf bw_DoWork
        AddHandler _bw.ProgressChanged, AddressOf bw_ProgressChanged
        AddHandler _bw.RunWorkerCompleted, AddressOf bw_RunWorkerCompleted

    End Sub

    Public Class Settings
        Public Query As String = ""
        Public ServerAddress As String = "localhost"
        Public DatabaseName As String = ""
        Public Username As String = ""
        Public Password As String = ""
        Public SelectedDriver As Integer = 1
        Public RowsToReturn As Integer = 1000
        Public DriverName As String = ""
        Public ConnectionString As String = ""
    End Class

    Public Function GetSettings() As Settings

        Return New Settings With {
            .Query = txtSQL.Text,
            .ServerAddress = ServerIP,
            .DatabaseName = DatabaseName,
            .Username = UserName,
            .Password = Password,
            .SelectedDriver = SelectedDriver,
            .RowsToReturn = RowsToReturn,
            .DriverName = DriverName,
            .ConnectionString = ConnectionString
        }

    End Function

    Private Sub bw_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs)

        Select Case e.ProgressPercentage
            Case 100 ' connecting
                lblStatus.Text = "Connecting..."

            Case 200 ' executing
                lblStatus.Text = "Executing Query. Waiting for response..."

            Case 300 ' reading
                lblStatus.Text = "Reading Data..."

        End Select

    End Sub

    Private Sub DisableSearchBox()
        Label11.Enabled = False
        Panel_Search.Cursor = Cursors.Default
        txtFind.Visible = False
    End Sub

    Private Sub EnableSearchBox()
        Label11.Enabled = True
        Panel_Search.Cursor = Cursors.IBeam
        txtFind.Visible = True
    End Sub

    Public Sub ExecuteSQL()

        If _bw.IsBusy = True Then
            MsgBox("Still busy running a query from last time.", MsgBoxStyle.Information)
        Else

            ' start background worker to execute sql

            ' initalise GUI
            DataGridView1.DataSource = Nothing
            txtErrorText.Visible = False
            'lblReservedWords.Text = ""
            'lblKeywords.Visible = False
            lblDurationConnect.Visible = False
            lblDurationExecute.Visible = False
            lblFindCount.Visible = False
            _LastSearch = ""
            DisableSearchBox()

            lblStatus.Text = ""
            ToolTip1.SetToolTip(lblSyntaxWarning, "")
            lblSyntaxWarning.Visible = False
            Panel1.Enabled = False
            Me.Refresh()

            lblLoading.Visible = True

            SetVars()

            _bw_Data.Clear()
            _bw_Error = ""
            _bw_RecordsAffected = 0
            _bw_Time_Conn = New TimeSpan
            _bw_Time_Exe = New TimeSpan
            _bw_Time_Stream = New TimeSpan


            'If txtSQL.SelectionLength > 0 Then
            '    _bw_SQL = txtSQL.SelectedText ' execute selected text
            'Else
            _bw_SQL = txtSQL.Text ' execute all
            'End If


            If SelectedDriver = 1 Then
                If String.IsNullOrWhiteSpace(_bw_Server) Then
                    _bw_Error = "ERROR: null server."
                    bw_RunWorkerCompleted(Me, New RunWorkerCompletedEventArgs(Nothing, Nothing, Nothing))
                    Return
                End If

                If String.IsNullOrWhiteSpace(_bw_Database) Then
                    _bw_Error = "ERROR: null database name."
                    bw_RunWorkerCompleted(Me, New RunWorkerCompletedEventArgs(Nothing, Nothing, Nothing))
                    Return
                End If

                If String.IsNullOrWhiteSpace(_bw_Username) Then
                    _bw_Error = "ERROR: null username."
                    bw_RunWorkerCompleted(Me, New RunWorkerCompletedEventArgs(Nothing, Nothing, Nothing))
                    Return
                End If

                If String.IsNullOrWhiteSpace(_bw_Password) Then
                    _bw_Error = "ERROR: null password."
                    bw_RunWorkerCompleted(Me, New RunWorkerCompletedEventArgs(Nothing, Nothing, Nothing))
                    Return
                End If
            End If

            If SelectedDriver = 0 Then
                If String.IsNullOrWhiteSpace(_bw_DriverName) Then
                    _bw_Error = "ERROR: null driver name."
                    bw_RunWorkerCompleted(Me, New RunWorkerCompletedEventArgs(Nothing, Nothing, Nothing))
                    Return
                End If

                If String.IsNullOrWhiteSpace(txtConnectionString.Text) Then
                    _bw_Error = "ERROR: null connection string."
                    bw_RunWorkerCompleted(Me, New RunWorkerCompletedEventArgs(Nothing, Nothing, Nothing))
                    Return
                End If
            End If

            If String.IsNullOrWhiteSpace(_bw_SQL) Then
                _bw_Error = "ERROR: type a SQL query first."
                bw_RunWorkerCompleted(Me, New RunWorkerCompletedEventArgs(Nothing, Nothing, Nothing))
                Return
            End If


            ' --------------------------------------------------------------------------------------
            ' Check for syntax errors:
            ' --------------------------------------------------------------------------------------
            If SelectedDriver = 1 Then
                ShowSyntaxWarnings(FM_CheckErrors(_bw_SQL))
            End If


            ' --------------------------------------------------------------------------------------
            ' MAKE SQL LOOK PRETTY
            ' --------------------------------------------------------------------------------------
            ' DISABLED 21/02/2018, doesnt properly, more annoying than useful!
            '_bw_SQL_Formatted = BeautifySQL(_bw_SQL)
            _bw_SQL_Formatted = _bw_SQL



            ' --------------------------------------------------------------------------------------
            ' ensure correct syntax, and prep for FileMaker:
            ' --------------------------------------------------------------------------------------
            If SelectedDriver = 1 Then
                _bw_SQL = FM_PrepSQL(_bw_SQL)
            End If



            Dim s = txtSQL.SelectionStart
            Dim l = txtSQL.SelectionLength
            txtSQL.Text = _bw_SQL_Formatted
            txtSQL.SelectionStart = s
            txtSQL.SelectionLength = l


            ' debugging
            Try
                WriteToFile(_bw_SQL, AppDataDir & "\lastQuery.sql", False)
            Catch ex As Exception
            End Try

            ' run async
            _bw.RunWorkerAsync()
        End If
    End Sub

    Public Sub ShowSyntaxWarnings(ByVal Warnings As ArrayList)

        If Warnings.Count = 0 Then
            lblSyntaxWarning.Visible = False
            Return
        End If

        Dim s As String = ""
        For Each w In Warnings
            s &= "- " & w & vbNewLine
        Next
        s = s.Trim

        ToolTip1.SetToolTip(lblSyntaxWarning, s)
        lblSyntaxWarning.Visible = True

    End Sub

    Private Sub SetVars()

        _bw_Server = txtServerIP.Text
        _bw_Database = txtDatabaseName.Text
        _bw_Username = txtUsername.Text
        _bw_Password = txtPassword.Text

        If IsNumeric(cmbRowLimit.Text) Then
            _bw_RowsToReturn = CLng(cmbRowLimit.Text)
        End If
        If _bw_RowsToReturn <= 0 Then _bw_RowsToReturn = 1 ' return at least one row

        If SelectedDriver = 0 Then
            _bw_DriverName = txtDriverName.Text
            _bw_ConnectionString = "DRIVER={" & _bw_DriverName & "};" & txtConnectionString.Text
        Else
            _bw_DriverName = "FileMaker ODBC"
            _bw_ConnectionString = "DRIVER={" & _bw_DriverName & "};SERVER=" & _bw_Server & ";UID=" & _bw_Username & ";PWD=" & _bw_Password & ";DATABASE=" & _bw_Database & ";"
        End If

    End Sub

    Private Function GetTableNames() As Object

        SetVars()

        Try

            Dim MetaData As New List(Of String)
            Dim TableNames As New List(Of String)
            Dim Err As String = ""

            Using cn As New Odbc.OdbcConnection(_bw_ConnectionString)
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

    Private Sub bw_DoWork(ByVal sender As Object, ByVal e As DoWorkEventArgs)

        If _bw.CancellationPending = True Then
            e.Cancel = True

        Else ' Perform a time consuming operation and report progress

            Dim sw As New System.Diagnostics.Stopwatch

            LastQrySQL = _bw_SQL


            ' are we executing multiple transactions?
            If _bw_SQL.EndsWith(";") Then _bw_SQL = TrimEnd(_bw_SQL, 1)

            Dim Queries = SplitSQLTRANS(_bw_SQL)


            Try

                If Queries.Count > 1 Then

                    ' execute transaction

                    Using cn As New Odbc.OdbcConnection(_bw_ConnectionString)

                        sw.Restart()

                        _bw.ReportProgress(100) ' connecting

                        cn.Open()

                        _bw_Time_Conn = sw.Elapsed
                        sw.Restart()

                        _bw.ReportProgress(200) ' executing

                        SQL_TRANS(Queries.ToList, cn)

                        _bw_Time_Exe = sw.Elapsed

                        _bw_RecordsAffected = 0

                    End Using

                Else

                    ' execute single query

                    Using cn As New Odbc.OdbcConnection(_bw_ConnectionString)

                        _bw.ReportProgress(100) ' connecting

                        sw.Restart()

                        cn.Open()

                        _bw_Time_Conn = sw.Elapsed

                        _bw.ReportProgress(200) ' executing

                        ' ################################################################

                        Dim Col As New Collection
                        Dim first As Boolean = False


                        Using cmd As New Odbc.OdbcCommand(_bw_SQL, cn)

                            sw.Restart()

                            Using reader = cmd.ExecuteReader()

                                _bw_Time_Exe = sw.Elapsed

                                _bw.ReportProgress(300) ' reading

                                sw.Restart()

                                Dim s As New ArrayList

                                If first = False Then ' for the first loop, add the names of the fields to the result

                                    For i = 0 To reader.FieldCount - 1
                                        Dim ColumnName As String = reader.GetName(i).Trim
                                        Dim NewColumnName As String = ""
                                        Dim Incrementor As String = " "

                                        If s.Contains(ColumnName) Then
                                            ' cant add duplicate columns to dataaset
                                            ' incrementally append a blank space on the end of coumn name  until it is not a duplicate:
                                            Do
                                                NewColumnName = ColumnName & Incrementor
                                                Incrementor &= " "
                                            Loop Until Not s.Contains(NewColumnName)
                                            s.Add(NewColumnName)
                                        Else
                                            s.Add(ColumnName)
                                        End If

                                    Next

                                    first = True
                                    Col.Add(s)
                                End If

                                Dim cc As Integer = 1
                                Do While reader.Read() And cc <= _bw_RowsToReturn
                                    s = New ArrayList

                                    For i = 0 To reader.FieldCount - 1
                                        If reader.GetValue(i).GetType = GetType(TimeSpan) Then
                                            ' filemake "time" fields return timepsans which arent supported by the data grid view, 
                                            ' so here we convert timespans into the a more readable date datatype
                                            Dim d As New Date
                                            Dim ts As TimeSpan = reader.GetValue(i)
                                            d = d.Add(ts)
                                            s.Add(d.ToString("h:mm:ss tt").ToLower)

                                        ElseIf Not reader.IsDBNull(i) Then

                                            s.Add(reader.GetValue(i))

                                        Else
                                            s.Add(Nothing)
                                        End If
                                    Next

                                    cc += 1
                                    Col.Add(s)
                                Loop

                                _bw_Time_Stream = sw.Elapsed

                                _bw_RecordsAffected = reader.RecordsAffected

                            End Using

                        End Using

                        _bw_Data = Col

                    End Using

                End If

            Catch ex1 As ArgumentOutOfRangeException
                If ex1.Message = "Year, Month, and Day parameters describe an un-representable DateTime." Then ' filemaker allows importing incorrect data into fields, so we need to catch these errors!
                    _bw_Error = "ERROR: One of the date fields in your query contains an incorrectly formatted value."
                Else
                    _bw_Error = ex1.Message
                End If
                _bw_Error &= vbNewLine & vbNewLine & LastQrySQL
                Return

            Catch ex As Exception
                _bw_Error = ex.Message
                _bw_Error &= vbNewLine & vbNewLine & LastQrySQL
                Return

            End Try

        End If
    End Sub

    Private Sub bw_RunWorkerCompleted(ByVal sender As Object, ByVal e As RunWorkerCompletedEventArgs)

        ' show error (if any)
        If _bw_Error <> "" Then

            ' error before executing
            lblStatus.Text = "Error"
            txtErrorText.Visible = True
            txtErrorText.Text = _bw_Error
            DisableSearchBox()
        Else

            ' show data in data grid view
            ' ################################################################

            Dim r As Collection = _bw_Data

            If r.Count > 0 Then

                Dim DS As New DataSet
                DS.Tables.Add("SELECT")
                Dim t = DS.Tables("SELECT")


                For i = 0 To r(1).Count - 1
                    t.Columns.Add(r(1)(i))
                Next

                If r.Count > 1 Then
                    Dim row As System.Data.DataRow

                    For i = 2 To r.Count
                        row = t.NewRow

                        Dim s(r(1).Count - 1) As Object  ' initalise object with same number of rows

                        For j = 0 To r(i).Count - 1
                            row.Item(j) = r(i)(j)
                        Next

                        t.Rows.Add(row)

                        If i > _bw_RowsToReturn Then
                            Exit For
                        End If

                    Next

                    DataGridView1.DataSource = t

                End If

                LastQryCount = r.Count - 1

                If DataGridView1.Rows.Count > 0 Then
                    'DataGridView1.Select()

                    If txtFind.Text.Count < 2 Then ' select first cell if we are not already searching for something
                        DataGridView1.Rows(0).Cells(0).Selected = True
                    End If

                    RefreshRowCount()
                Else
                    lblStatus.Text = "Row 0 of 0"
                End If

                'Dim jf = DataGridView1.Columns()

                For Each c In DataGridView1.Columns
                    c.AutoSizeMode = DataGridViewAutoSizeColumnMode.AllCells
                    Dim autowidth As Integer = c.Width + 10 ' nb: add 10 pixels on end as auto size doesnt quite fit perfectly
                    c.AutoSizeMode = DataGridViewAutoSizeColumnMode.None

                    If autowidth > 150 Then
                        c.Width = 150
                    Else
                        c.Width = autowidth
                    End If
                Next

            End If

            ' show number records affected:
            If _bw_RecordsAffected > -1 Then
                lblStatus.Text = _bw_RecordsAffected & " Records Affected"
            End If

        End If

        Panel1.Enabled = True
        lblLoading.Visible = False
        btnGo.Text = "Execute (F5)"

        ' show elapsed time
        Dim ElapsedTime As TimeSpan = _bw_Time_Conn.Add(_bw_Time_Exe.Add(_bw_Time_Stream))
        lblDurationConnect.Visible = True
        lblDurationExecute.Visible = True
        'lblDuration2.Text = ElapsedTime.Minutes.ToString("00") & ":" & ElapsedTime.Seconds.ToString("00") & "." & (ElapsedTime.Milliseconds / 10).ToString("00")
        lblDurationConnect.Text = "Connect: " & FormatTime(_bw_Time_Conn)

        Dim runtime = _bw_Time_Exe + _bw_Time_Stream
        lblDurationExecute.Text = "Execute: " & FormatTime(runtime)

        Dim ttt As String = "Process: " & FormatTime(_bw_Time_Exe) &
                            ", Stream: " & FormatTime(_bw_Time_Stream)
        ToolTip1.SetToolTip(lblDurationExecute, ttt)

        ' show/hide status
        If lblStatus.Text = "" Then
            lblStatus.Visible = False
        Else
            lblStatus.Visible = True
        End If

    End Sub

    Private Sub dataGridView1_DataBindingComplete(ByVal sender As Object, ByVal e As DataGridViewBindingCompleteEventArgs) Handles DataGridView1.DataBindingComplete
        ' re-index the search results after new data is loaded
        Search(0)
    End Sub

    Public Property ServerIP() As String
        Get
            Return txtServerIP.Text
        End Get
        Set(ByVal value As String)
            txtServerIP.Text = value
        End Set
    End Property

    Public Property DatabaseName() As String
        Get
            Return txtDatabaseName.Text
        End Get
        Set(ByVal value As String)
            txtDatabaseName.Text = value
        End Set
    End Property

    Public Property UserName() As String
        Get
            Return txtUsername.Text
        End Get
        Set(ByVal value As String)
            txtUsername.Text = value
        End Set
    End Property

    Public Property Password() As String
        Get
            Return txtPassword.Text
        End Get
        Set(ByVal value As String)
            txtPassword.Text = value
        End Set
    End Property

    Public Property SQLText() As String
        Get
            Return txtSQL.Text
        End Get
        Set(ByVal value As String)
            txtSQL.Text = value
        End Set
    End Property

    Public Property DriverName() As String
        Get
            Return txtDriverName.Text
        End Get
        Set(ByVal value As String)
            txtDriverName.Text = value
        End Set
    End Property

    Public Property ConnectionString() As String
        Get
            Return txtConnectionString.Text
        End Get
        Set(ByVal value As String)
            txtConnectionString.Text = value
        End Set
    End Property

    Public Property SelectedDriver() As Integer
        Get
            If cmbDriver.Text.ToLower = "other" Then
                Return 0
            Else
                Return 1
            End If
        End Get
        Set(ByVal value As Integer)
            If value = 0 Then
                cmbDriver.SelectedIndex = 1
            Else
                cmbDriver.SelectedIndex = 0
            End If
        End Set
    End Property

    Public Property RowsToReturn() As Integer
        Get
            If IsNumeric(cmbRowLimit.Text) Then
                Return cmbRowLimit.Text
            Else
                Return 1000
            End If
        End Get
        Set(ByVal value As Integer)
            cmbRowLimit.Text = value
        End Set
    End Property

    Private Sub DataGridView1_CurrentCellChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellChanged
        RefreshRowCount()
    End Sub

    Private Sub RefreshRowCount()

        Dim TotalRows As String = LastQryCount
        If LastQryCount = _bw_RowsToReturn Then ' show "max" label to remind us the rows have been limited
            TotalRows &= " (max)"
        End If

        If DataGridView1.SelectedCells.Count = 0 Then
            lblStatus.Text = "Row 0 of " & TotalRows
        Else
            lblStatus.Text = "Row " & DataGridView1.SelectedCells(0).RowIndex + 1 & " of " & TotalRows
        End If

    End Sub

    Private Sub btnGo_Click(sender As System.Object, e As System.EventArgs) Handles btnGo.Click
        RaiseEvent BeginSQLExecute()
        ExecuteSQL()
    End Sub

    Private Sub TextBox1_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs)
        If e.KeyData = Keys.Tab Then ' allow tabbing inside the text box
            e.IsInputKey = True
        End If
    End Sub

    Private Sub txtDatabaseName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtDatabaseName.TextChanged
        If SelectedDriver = 1 Then
            RaiseEvent DatabaseNameChanged(New TextChangedArgs(txtDatabaseName.Text))
        End If
    End Sub

    Private Sub txtDriverName_TextChanged(sender As Object, e As EventArgs) Handles txtDriverName.TextChanged
        If SelectedDriver = 0 Then
            RaiseEvent DatabaseNameChanged(New TextChangedArgs(txtDriverName.Text))
        End If
    End Sub

    Private Sub txtCredentails_Changed(sender As System.Object, e As System.EventArgs) Handles txtServerIP.TextChanged, txtDatabaseName.TextChanged, txtPassword.TextChanged, txtUsername.TextChanged
        Dim t As TextBox = DirectCast(sender, TextBox)
        If t.Text = "" Then
            t.BackColor = SystemColors.Info
        Else
            t.BackColor = SystemColors.Window
        End If
    End Sub

#Region "Macros"

    Private Sub SECONDToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SECONDToolStripMenuItem.Click
        txtSQL.InsertText("SECOND()")
    End Sub

    Private Sub MINUTEToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MINUTEToolStripMenuItem.Click
        txtSQL.InsertText("MINUTE()")
    End Sub

    Private Sub HOURToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles HOURToolStripMenuItem.Click
        txtSQL.InsertText("HOUR()")
    End Sub

    Private Sub DAYToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DAYToolStripMenuItem.Click
        txtSQL.InsertText("DAY()")
    End Sub

    Private Sub MONTHToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MONTHToolStripMenuItem.Click
        txtSQL.InsertText("MONTH()")
    End Sub

    Private Sub YEARToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles YEARToolStripMenuItem.Click
        txtSQL.InsertText("YEAR()")
    End Sub

    Private Sub CURRENTDATEToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles CURRENTDATEToolStripMenuItem1.Click
        txtSQL.InsertText("CURRENT_DATE")
    End Sub

    Private Sub CURRENTTIMEToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles CURRENTTIMEToolStripMenuItem1.Click
        txtSQL.InsertText("CURRENT_TIME")
    End Sub

    Private Sub CURRENTTIMESTAMPToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles CURRENTTIMESTAMPToolStripMenuItem1.Click
        txtSQL.InsertText("CURRENT_TIMESTAMP")
    End Sub

    Private Sub TODAYToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)
        txtSQL.InsertText("TODAY")
    End Sub

    Private Sub DATEVAL01302011ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DATEVAL01302011ToolStripMenuItem.Click
        txtSQL.InsertText("DATEVAL()")
    End Sub

    Private Sub DateToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DateToolStripMenuItem.Click
        txtSQL.InsertText("{d '" & Now.ToString("yyyy-MM-dd") & "'}")
    End Sub

    Private Sub TimeToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles TimeToolStripMenuItem.Click
        txtSQL.InsertText("{t '" & Now.ToString("hh:mm:ss") & "'}")
    End Sub

    Private Sub TimestampToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles TimestampToolStripMenuItem.Click
        txtSQL.InsertText("{ts '" & Now.ToString("yyyy-MM-dd hh:mm:ss") & "'}")
    End Sub

    Private Sub DAYOFWEEKToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DAYOFWEEKToolStripMenuItem.Click
        txtSQL.InsertText("DAYOFWEEK(")
    End Sub

    Private Sub CHR67ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CHR67ToolStripMenuItem.Click
        txtSQL.InsertText("CHR()")
    End Sub

    Private Sub UPPERToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles UPPERToolStripMenuItem.Click
        txtSQL.InsertText("UPPPER()")
    End Sub

    Private Sub LOWERToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles LOWERToolStripMenuItem.Click
        txtSQL.InsertText("LOWER()")
    End Sub

    Private Sub ROUND1234560ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ROUND1234560ToolStripMenuItem.Click
        txtSQL.InsertText("ROUND()")
    End Sub

    Private Sub LENstringToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles LENstringToolStripMenuItem.Click
        txtSQL.InsertText("LEN()")
    End Sub

    Private Sub MONTHNAMEToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles MONTHNAMEToolStripMenuItem.Click
        txtSQL.InsertText("MONTHNAME()")
    End Sub

    Private Sub USERNAMEToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles USERNAMEToolStripMenuItem.Click
        txtSQL.InsertText("CURRENT_USER")
    End Sub

    Private Sub DAYNAMEdateToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DAYNAMEdateToolStripMenuItem.Click
        txtSQL.InsertText("DAYNAME()")
    End Sub

    Private Sub SUBSTRINGstring23ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SUBSTRINGstring23ToolStripMenuItem.Click
        txtSQL.InsertText("SUBSTR('string', 2, 3)")
    End Sub

    Private Sub CEILToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CEILToolStripMenuItem.Click
        txtSQL.InsertText("CEIL()")
    End Sub

    Private Sub FLOORToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FLOORToolStripMenuItem.Click
        txtSQL.InsertText("FLOOR()")
    End Sub

    Private Sub NUMVALstringToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NUMVALstringToolStripMenuItem.Click
        txtSQL.InsertText("NUMVAL()")
    End Sub

    Private Sub MIN6689ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MIN6689ToolStripMenuItem.Click
        txtSQL.InsertText("MIN()")
    End Sub

    Private Sub MAX6689ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles MAX6689ToolStripMenuItem.Click
        txtSQL.InsertText("MAX()")
    End Sub

    Private Sub INT64321Returns6ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles INT64321Returns6ToolStripMenuItem.Click
        txtSQL.InsertText("INT()")
    End Sub

    Private Sub RTRIMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RTRIMToolStripMenuItem.Click
        txtSQL.InsertText("RTRIM()")
    End Sub

    Private Sub TRIMToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles TRIMToolStripMenuItem.Click
        txtSQL.InsertText("TRIM()")
    End Sub

    Private Sub LEFTstring4ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles LEFTstring4ToolStripMenuItem.Click
        txtSQL.InsertText("LEFT()")
    End Sub

    Private Sub RIGHTstring4ToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RIGHTstring4ToolStripMenuItem.Click
        txtSQL.InsertText("RIGHT()")
    End Sub

    Private Sub STRVALToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles STRVALToolStripMenuItem.Click
        txtSQL.InsertText("STRVAL()")
    End Sub

#End Region

    Private Sub ContextMenuStrip2_Opening(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip2.Opening
        mnuCopy.Enabled = False
        mnuCopyAsCSV.Enabled = False
        mnuExportToExcel.Enabled = False
        mnuCopyWithHeaders.Enabled = False

        If DataGridView1.Rows.Count > 0 Then
            If _ColumnRightClicked <> -1 Then
                mnuCopyAsCSV.Enabled = True
            End If

            If DataGridView1.SelectedCells.Count > 0 Then
                mnuCopy.Enabled = True
                mnuCopyWithHeaders.Enabled = True
            End If

            'If _ColumnRightClicked <> -1 And _RowRightClicked <> -1 Then
            '    mnuCopy.Enabled = True
            'End If

            mnuExportToExcel.Enabled = True

        End If
    End Sub


    Private _ColumnRightClicked As Integer = -1
    Private _RowRightClicked As Integer = -1

    Private Sub mnuCopyAsCSV_Click(sender As System.Object, e As System.EventArgs) Handles mnuCopyAsCSV.Click
        Try
            Dim s = ""
            For Each r As DataGridViewRow In DataGridView1.Rows
                Dim value = r.Cells(_ColumnRightClicked).Value
                If Not IsDBNull(r.Cells(_ColumnRightClicked).Value) Then
                    s &= r.Cells(_ColumnRightClicked).Value.ToString & ","
                End If
            Next
            System.Windows.Forms.Clipboard.SetText(s.TrimEnd(1))
        Catch ex As Exception
            MsgBox(ex.Message)
        End Try
    End Sub

    Private ClipboardCopy As Boolean = False
    Private Sub CopyCellToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles mnuCopy.Click
        Try
            ClipboardCopy = True
            System.Windows.Forms.Clipboard.SetDataObject(DataGridView1.GetClipboardContent)
            ClipboardCopy = False
        Catch ex As Exception
        End Try
    End Sub

    Private Sub mnuCopyWithHeaders_Click(sender As Object, e As EventArgs) Handles mnuCopyWithHeaders.Click
        Try
            DataGridView1.RowHeadersVisible = False
            DataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
            ClipboardCopy = True
            System.Windows.Forms.Clipboard.SetDataObject(DataGridView1.GetClipboardContent)
            ClipboardCopy = False
            DataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
            DataGridView1.RowHeadersVisible = True
        Catch ex As Exception
        End Try
    End Sub

    Private Sub DataGridView1_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseDown

        Dim s As DataGridView = sender
        _ColumnRightClicked = s.HitTest(e.Location.X, e.Location.Y).ColumnIndex
        _RowRightClicked = s.HitTest(e.Location.X, e.Location.Y).RowIndex

    End Sub

    Private Sub mnuPaste_Click(sender As System.Object, e As System.EventArgs) Handles mnuPaste.Click

        ' todo: is this nessecary now that text box suports it natively?

        txtSQL.SelectedText = ""

        Dim s As String = System.Windows.Forms.Clipboard.GetText

        ' replace clean up newlines characters
        s = s.Replace(Environment.NewLine, vbCr)
        s = s.Replace(vbLf, vbCr)
        s = s.Replace(vbCr, Environment.NewLine)

        txtSQL.InsertText(s)

    End Sub

    Private Sub mnuCopy1_Click(sender As System.Object, e As System.EventArgs) Handles mnuCopy1.Click
        Try
            Dim value = txtSQL.SelectedText
            If Not IsDBNull(value) Then
                System.Windows.Forms.Clipboard.SetText(value.ToString)
            End If
        Catch ex As Exception
        End Try
    End Sub

    Private Sub mnuCut1_Click(sender As System.Object, e As System.EventArgs) Handles mnuCut1.Click
        Try
            Dim value = txtSQL.SelectedText
            If Not IsDBNull(value) Then
                System.Windows.Forms.Clipboard.SetText(value.ToString)
            End If
            txtSQL.SelectedText = ""
        Catch ex As Exception
        End Try
    End Sub

    Private Sub ContextMenuStrip1_Opening(sender As System.Object, e As System.ComponentModel.CancelEventArgs) Handles ContextMenuStrip1.Opening
        If String.IsNullOrEmpty(txtSQL.SelectedText) Then
            mnuCut1.Enabled = False
            mnuCopy1.Enabled = False
        Else
            mnuCut1.Enabled = True
            mnuCopy1.Enabled = True
        End If

        If String.IsNullOrEmpty(txtSQL.Text) Then
            mnuSelectAll.Enabled = False
        Else
            mnuSelectAll.Enabled = True
        End If

    End Sub

    Private Sub mnuSelectAll_Click(sender As System.Object, e As System.EventArgs) Handles mnuSelectAll.Click
        txtSQL.SelectAll()
    End Sub

    Private Sub mnuAbout_Click(sender As Object, e As EventArgs) Handles mnuAbout.Click
        frmAbout.ShowDialog(Me)
    End Sub

    Private Sub mnuExportToExcel_Click(sender As Object, e As EventArgs) Handles mnuExportToExcel.Click

        Dim ColumnHeadings As New List(Of String)
        For Each i In DataGridView1.Columns
            ColumnHeadings.Add(i.DataPropertyName)
        Next

        Dim Data As New List(Of List(Of String))
        Data.Add(ColumnHeadings)

        For Each i In DataGridView1.Rows
            Dim Row As New List(Of String)
            For Each c In i.cells
                Row.Add(c.value.ToString)
            Next
            Data.Add(Row)
        Next

        ExportToExcel(Data)

    End Sub

    Private Sub url_sql_v13_Click(sender As Object, e As EventArgs) Handles url_sql_v13.Click
        Process.Start("https://fmhelp.filemaker.com/docs/13/en/fm13_sql_reference.pdf")
    End Sub

    Private Sub url_sql_v14_Click(sender As Object, e As EventArgs) Handles url_sql_v14.Click
        Process.Start("https://fmhelp.filemaker.com/docs/14/en/fm14_sql_reference.pdf")
    End Sub

    Private Sub url_sql_v15_Click(sender As Object, e As EventArgs) Handles url_sql_v15.Click
        Process.Start("https://fmhelp.filemaker.com/docs/15/en/fm15_sql_reference.pdf")
    End Sub

    Private Sub url_sql_v16_Click(sender As Object, e As EventArgs) Handles url_sql_v16.Click
        Process.Start("https://fmhelp.filemaker.com/docs/16/en/fm16_sql_reference.pdf")
    End Sub

    Private Sub url_odbc_v13_Click(sender As Object, e As EventArgs) Handles url_odbc_v13.Click
        Process.Start("https://fmhelp.filemaker.com/docs/13/en/fm13_odbc_jdbc_guide.pdf")
    End Sub

    Private Sub url_odbc_v14_Click(sender As Object, e As EventArgs) Handles url_odbc_v14.Click
        Process.Start("https://fmhelp.filemaker.com/docs/14/en/fm14_odbc_jdbc_guide.pdf")
    End Sub

    Private Sub url_odbc_v15_Click(sender As Object, e As EventArgs) Handles url_odbc_v15.Click
        Process.Start("https://fmhelp.filemaker.com/docs/15/en/fm15_odbc_jdbc_guide.pdf")
    End Sub

    Private Sub url_odbc_v16_Click(sender As Object, e As EventArgs) Handles url_odbc_v16.Click
        Process.Start("https://fmhelp.filemaker.com/docs/16/en/fm16_odbc_jdbc_guide.pdf")
    End Sub

    Private Sub cmbDriver_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbDriver.SelectedIndexChanged
        If SelectedDriver = 0 Then ' other
            Panel_Driver_Custom.Visible = True
            Panel_Driver_FileMaker.Visible = False
            RaiseEvent DatabaseNameChanged(New TextChangedArgs(txtDriverName.Text))
        Else
            Panel_Driver_Custom.Visible = False
            Panel_Driver_FileMaker.Visible = True
            RaiseEvent DatabaseNameChanged(New TextChangedArgs(txtDatabaseName.Text))
        End If
    End Sub

    Private Sub textBox_TextChanged(ByVal sender As Object, ByVal e As TextChangedEventArgs) Handles txtSQL.TextChanged
        SetRangeStyle(txtSQL.Range)
    End Sub

    Private Sub txtFind_KeyDown(sender As Object, e As KeyEventArgs) Handles txtFind.KeyDown
        If e.KeyCode = Keys.F3 And e.Shift Then
            Search(2)
        ElseIf e.KeyCode = Keys.Enter Or e.KeyCode = Keys.F3 Then
            Search(1)
        End If
    End Sub

    Private Sub txtFind_KeyUp(sender As Object, e As KeyEventArgs) Handles txtFind.KeyUp
        If Not e.KeyCode = Keys.Enter And
            Not (e.KeyCode >= 112 And e.KeyCode <= 123) And
            Not e.KeyCode = Keys.Alt And
            Not e.KeyCode = Keys.ControlKey And
            Not e.KeyCode = Keys.ShiftKey Then
            Search(0)
        End If
    End Sub

    Private _LastSearch As String = ""
    Private _CurrentFoundCell As Integer = 0
    Private _FoundCells As New List(Of DataGridViewCell)
    Private Sub Search(ByVal FindMode As Integer)
        Try

            If FindMode = 1 Then ' FIND NEXT
                _CurrentFoundCell += 1

                If _CurrentFoundCell > _FoundCells.Count - 1 Then
                    _CurrentFoundCell = 0
                End If

                HighlightNext()
                Return

            ElseIf FindMode = 2 Then ' FIND PREV
                _CurrentFoundCell -= 1

                If _CurrentFoundCell < 0 Then
                    _CurrentFoundCell = _FoundCells.Count - 1
                End If

                HighlightNext()
                Return
            End If

            Dim t As String = txtFind.Text.ToLower

            If t.StartsWith(_LastSearch) And _LastSearch <> "" And _FoundCells.Count = 0 Then t = _LastSearch ' only if different than last 0-result search

            If _LastSearch <> t Then
                ' re-index items

                _CurrentFoundCell = 0
                _FoundCells.Clear()

                If Not String.IsNullOrEmpty(t) Then
                    For i = 0 To DataGridView1.Rows.Count - 1
                        Dim ro = DataGridView1.Rows(i)
                        For c = 0 To ro.Cells.Count - 1
                            Dim ce = ro.Cells(c)

                            'Console.WriteLine(ce.Value.ToString)
                            If ce.Value.ToString.ToLower.Contains(t) Then
                                _FoundCells.Add(ce)
                            End If
                        Next
                    Next
                End If

                _LastSearch = t

                ' show found count
                If t.Length > 0 Then
                    lblFindCount.Text = _FoundCells.Count
                    lblFindCount.Visible = True

                    If _FoundCells.Count > 0 Then
                        lblFindCount.BackColor = Color.White
                    Else
                        lblFindCount.BackColor = Color.LightSalmon
                    End If
                Else
                    lblFindCount.Visible = False
                End If

            End If

            If DataGridView1.Rows.Count = 0 Then
                DisableSearchBox()
            Else
                EnableSearchBox()
            End If

            HighlightNext()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub HighlightNext()
        If _FoundCells.Count > 0 Then
            DataGridView1.CurrentCell = _FoundCells(_CurrentFoundCell)
        End If
    End Sub

    Public Sub FocusFindTextbox()
        txtFind.Focus()
    End Sub

    Private Sub SplitContainer1_MouseUp(sender As Object, e As MouseEventArgs) Handles SplitContainer1.MouseUp
        lblDurationConnect.Focus()
    End Sub

    Private Sub DataGridView1_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting

        ' this is called when the clipboard retrieves data from the cell as well

        Dim c = DataGridView1.Rows(e.RowIndex).Cells(e.ColumnIndex)

        If c.Value.ToString = "" Then
            c.Style.BackColor = Color.LightYellow
        End If

        If Not ClipboardCopy Then
            If c.Value.ToString.Contains(vbCr) OrElse c.Value.ToString.Contains(vbLf) Then
                ' get first line of string
                Dim s = e.Value
                s = s.Split(vbCr)
                s = s(0).Split(vbLf)

                e.Value = s(0)
            End If
        End If

        ' todo, format dates here as required

    End Sub

    Private Sub txtFind_KeyPress(sender As Object, e As KeyPressEventArgs) Handles txtFind.KeyPress
        ' eat Enter key presses (which prevents a beep that happens when Me.Multiline = False and there is no default button on the form)
        If e.KeyChar = Chr(13) Then
            e.Handled = True
        End If
        MyBase.OnKeyPress(e)
    End Sub

    Private Sub Label11_Click(sender As Object, e As EventArgs) Handles Label11.Click, Panel_Search.Click
        FocusFindTextbox()
    End Sub

End Class

Public Class TextChangedArgs

    Public _NewTextValue As String = ""

    Public Sub New(ByVal NewTextValue As String)
        _NewTextValue = NewTextValue
    End Sub

End Class



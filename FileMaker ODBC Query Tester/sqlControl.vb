Imports System.Text.RegularExpressions
Imports System.ComponentModel

Public Class sqlControl

    Const EM_SETTABSTOPS As Integer = &HCB
    Declare Function SendMessageA Lib "user32" (ByVal TBHandle As IntPtr, _
                                               ByVal EM_SETTABSTOPS As Integer, _
                                               ByVal wParam As Integer, _
                                               ByRef lParam As Integer) As Boolean

    Private TabStopIndent As Integer = 16 'Tab times 4

    Friend Event DatabaseNameChanged(ByVal e As textChangedArgs)
    Friend Event BeginSQLExecute()
    Private LastQryCount As Integer = 0
    Private LastQrySQL As String = ""

    ' parameters for the background worker Printer:
    Private _bw As New BackgroundWorker()
    Private _bw_SQL As String = "" ' the query we are going to run
    Private _bw_SQL_Formatted As String = "" ' the SQL after it has been formatted
    Private _bw_Error As String = ""
    Private _bw_RservedKeywords As String = ""
    Private _bw_ServerIP As String = ""
    Private _bw_Database As String = ""
    Private _bw_Username As String = ""
    Private _bw_Password As String = ""
    Private _bw_Data As New Collection       ' stores results
    Private _bw_RecordsAffected As Integer = -1
    Private _bw_Time_Conn As New TimeSpan
    Private _bw_Time_Exe As New TimeSpan
    Private _bw_Time_Stream As New TimeSpan
    Private _bw_DriverName As String = ""
    Private _bw_RowsToReturn As Long
    Private _bw_ConnectionString As String = ""

    Public Delegate Sub bwPrinterDelegate(ByVal e As RunWorkerCompletedEventArgs)


    Public Sub InitaliseBackgroundWorker()
        _bw.WorkerSupportsCancellation = True
        _bw.WorkerReportsProgress = True

        AddHandler _bw.DoWork, AddressOf bw_DoWork
        AddHandler _bw.ProgressChanged, AddressOf bw_ProgressChanged
        AddHandler _bw.RunWorkerCompleted, AddressOf bw_RunWorkerCompleted
    End Sub

    Private Sub bw_ProgressChanged(ByVal sender As Object, ByVal e As ProgressChangedEventArgs)
        ' nb: In the ProgressChanged event handler, add code to indicate the progress, such as updating the user interface
        ' nb: To determine what percentage of the operation is completed, check the ProgressPercentage property of the ProgressChangedEventArgs object that was passed to the event handler.

        Select Case e.ProgressPercentage
            Case 100 ' connecting
                lblStatus2.Text = "Connecting..."

            Case 200 ' executing
                lblStatus2.Text = "Executing Query. Waiting for response..."

            Case 300 ' reading
                lblStatus2.Text = "Reading Data..."

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
            lblDuration2.Visible = False
            lblDuration3.Visible = False
            lblFindCount.Visible = False
            _LastSearch = ""
            DisableSearchBox()

            lblStatus2.Text = ""
            ToolTip1.SetToolTip(lblSyntaxWarning2, "")
            lblSyntaxWarning2.Visible = False
            Panel1.Enabled = False
            Me.Refresh()

            lblLoading.Visible = True

            SetGlobals()

            _bw_Data.Clear()
            _bw_Error = ""
            _bw_RecordsAffected = 0
            _bw_RservedKeywords = ""
            _bw_Time_Conn = New TimeSpan
            _bw_Time_Exe = New TimeSpan
            _bw_Time_Stream = New TimeSpan


            'If txtSQL.SelectionLength > 0 Then
            '    _bw_SQL = txtSQL.SelectedText ' execute selected text
            'Else
            _bw_SQL = txtSQL.Text ' execute all
            'End If


            If SelectedDriver = 1 Then
                If String.IsNullOrWhiteSpace(_bw_ServerIP) Then
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
                ShowWarning(FM_CheckErrors(_bw_SQL))
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


            ' debugging:
            Try
                WriteToFile(_bw_SQL, AppDataDir & "\Debugging\lastQuery.sql", False)
            Catch ex As Exception
            End Try

            ' run async
            _bw.RunWorkerAsync()
        End If
    End Sub


    Public Sub ShowWarning(ByVal errors As ArrayList)

        If errors.Count = 0 Then
            lblSyntaxWarning2.Visible = False
            Return
        End If

        Dim ttt As String = ""
        For Each s In errors
            ttt &= s & vbNewLine
        Next
        ttt = ttt.Trim

        ToolTip1.SetToolTip(lblSyntaxWarning2, ttt)
        lblSyntaxWarning2.Visible = True

    End Sub

    Private Sub SetGlobals()
        ' set globals

        _bw_ServerIP = txtServerIP.Text
        _bw_Database = txtDatabaseName.Text
        _bw_Username = txtUsername.Text
        _bw_Password = txtPassword.Text

        If IsNumeric(cmboRowsToReturn.Text) Then
            _bw_RowsToReturn = CLng(cmboRowsToReturn.Text)
        End If
        If _bw_RowsToReturn <= 0 Then _bw_RowsToReturn = 1 ' return at least one row

        If SelectedDriver = 0 Then
            _bw_DriverName = txtDriverName.Text
            _bw_ConnectionString = "DRIVER={" & _bw_DriverName & "};" & txtConnectionString.Text
        Else
            _bw_DriverName = "FileMaker ODBC"
            _bw_ConnectionString = "DRIVER={" & _bw_DriverName & "};SERVER=" & _bw_ServerIP & ";UID=" & _bw_Username & ";PWD=" & _bw_Password & ";DATABASE=" & _bw_Database & ";"
        End If

    End Sub

    Private Function GetTableNames() As Object
        SetGlobals()

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

        Else ' Perform a time consuming operation and report progress.


            Dim sw As New System.Diagnostics.Stopwatch

            LastQrySQL = _bw_SQL

            ' --------------------------------------------------------------------------------------
            ' 3) check to see if contains multiple transactions:
            ' --------------------------------------------------------------------------------------
            If _bw_SQL.EndsWith(";") Then _bw_SQL = TrimEnd(_bw_SQL, 1)
            Dim queries = SplitSQLTRANS(_bw_SQL)


            Try

                If queries.Count > 1 Then
                    ' perform a transaction
                    Using cn As New Odbc.OdbcConnection(_bw_ConnectionString)

                        sw.Restart()

                        _bw.ReportProgress(100) ' connecting

                        cn.Open()

                        _bw_Time_Conn = sw.Elapsed
                        sw.Restart()

                        _bw.ReportProgress(200) ' executing

                        Dim c As New Collection
                        For Each s As String In queries
                            c.Add(s)
                        Next
                        SQL_TRANS(c, cn)

                        _bw_Time_Exe = sw.Elapsed

                        _bw_RecordsAffected = 0

                    End Using

                Else

                    ' perform a single query
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
                If ex1.Message = "Year, Month, and Day parameters describe an un-representable DateTime." Then ' filemaker can sometimes allow importing of incorrect data into fields, so need to catch these errors!
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
            lblStatus2.Text = "Error"
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

                    Refreshrowcount()
                Else
                    lblStatus2.Text = "Row 0 of 0"
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
                lblStatus2.Text = _bw_RecordsAffected & " Records Affected"
            End If

        End If

        Panel1.Enabled = True
        lblLoading.Visible = False
        btnGo.Text = "Execute (F5)"

        ' show reserved keywords:
        Dim a As Array = GetSQLReservedKeyWords(_bw_SQL)
        'lblReservedWords.Text = a(0)
        'lblKeywords.Visible = True

        ' show elapsed time
        Dim ElapsedTime As TimeSpan = _bw_Time_Conn.Add(_bw_Time_Exe.Add(_bw_Time_Stream))
        lblDuration2.Visible = True
        lblDuration3.Visible = True
        'lblDuration2.Text = ElapsedTime.Minutes.ToString("00") & ":" & ElapsedTime.Seconds.ToString("00") & "." & (ElapsedTime.Milliseconds / 10).ToString("00")
        lblDuration2.Text = "Connect: " & formattime(_bw_Time_Conn)

        Dim runtime = _bw_Time_Exe + _bw_Time_Stream
        lblDuration3.Text = "Execute: " & formattime(runtime)

        Dim ttt As String = "Process: " & formattime(_bw_Time_Exe) &
                            ", Stream: " & formattime(_bw_Time_Stream)
        ToolTip1.SetToolTip(lblDuration3, ttt)

        ' show/hide status
        If lblStatus2.Text = "" Then
            lblStatus2.Visible = False
        Else
            lblStatus2.Visible = True
        End If
    End Sub

    Private Function formattime(ByVal t As TimeSpan) As String
        If t.Minutes = 0 Then
            Return t.Seconds.ToString("#0") & "." & (t.Milliseconds / 10).ToString("00")
        Else
            Return t.Minutes.ToString("#0") & ":" & t.Seconds.ToString("#0") & "." & (t.Milliseconds / 10).ToString("00")
        End If
    End Function

    Private Sub dataGridView1_DataBindingComplete(ByVal sender As Object, ByVal e As DataGridViewBindingCompleteEventArgs) Handles DataGridView1.DataBindingComplete
        ' re-index the search results after new data is loaded:
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
            If cmbSelectedDriver.Text.ToLower = "other" Then
                Return 0
            Else
                Return 1
            End If
        End Get
        Set(ByVal value As Integer)
            If value = 0 Then
                cmbSelectedDriver.SelectedIndex = 1
            Else
                cmbSelectedDriver.SelectedIndex = 0
            End If
        End Set
    End Property

    Public Property RowsToReturn() As Integer
        Get
            If IsNumeric(cmboRowsToReturn.Text) Then
                Return cmboRowsToReturn.Text
            Else
                Return 1000
            End If
        End Get
        Set(ByVal value As Integer)
            cmboRowsToReturn.Text = value
        End Set
    End Property

    Private Sub UserControl1_Load(sender As Object, e As System.EventArgs) Handles Me.Load
        SendMessageA(txtSQL.Handle, EM_SETTABSTOPS, 1, TabStopIndent) ' set tab width

        ' reset
        txtErrorText.Dock = DockStyle.Fill
        txtErrorText.Text = ""
        lblStatus2.Text = ""
        lblDuration2.Visible = False
        lblDuration3.Visible = False
        lblSyntaxWarning2.Visible = False
        txtSQL.Select()
        'lblReservedWords.Text = ""
        DisableSearchBox()

        'lblKeywords.Visible = False

        InitaliseBackgroundWorker()
    End Sub

    Private Sub DataGridView1_CurrentCellChanged(sender As Object, e As EventArgs) Handles DataGridView1.CurrentCellChanged
        Refreshrowcount()
    End Sub

    ' update status bar
    Private Sub Refreshrowcount()

        Dim TotalRows As String = LastQryCount
        If LastQryCount = _bw_RowsToReturn Then ' show "max" label to remind us the rows have been limited
            TotalRows &= " (max)"
        End If

        If DataGridView1.SelectedCells.Count = 0 Then
            lblStatus2.Text = "Row 0 of " & TotalRows
        Else
            lblStatus2.Text = "Row " & DataGridView1.SelectedCells(0).RowIndex + 1 & " of " & TotalRows
        End If
    End Sub

    Private Sub Button2_Click(sender As System.Object, e As System.EventArgs) Handles btnGo.Click
        RaiseEvent BeginSQLExecute()
        ExecuteSQL()
    End Sub

    Public Function GetSQLReservedKeyWords(ByVal sql As String) As String()

        Dim sc As String = ""
        Dim counter As Integer = 0


        ' check for keywords
        For Each k In FM_ReservedKeyWords
            If sql.ToUpper.Contains(k) Then

                ' determine if it is a complete word by looking at the chars before and after it:

                If k = "POSITION" Then
                    Dim f = ""
                End If

                ' get char after and before the keyword
                Dim w As Integer = sql.ToUpper.IndexOf(k)
                Dim charbefore As String = ""
                Dim charAfter As String = ""

                If w = -1 Then Continue For

                If w <> 0 Then
                    charbefore = sql(w - 1)
                End If
                If w + k.Length <> sql.Length Then
                    charAfter = sql(w + k.Length)
                End If

                ' check if they are quotes
                If charbefore = """" And charAfter = """" Then
                    Continue For
                End If

                Dim CharBefore_NewWordFlag As Boolean = False
                Dim CharAfter_NewWordFlag As String = False

                ' check start of word for new word characters
                If charbefore = "" Or charbefore = " " Or charbefore = "(" Or charbefore = "." Or charbefore = vbLf Then
                    CharBefore_NewWordFlag = True
                End If

                ' check end of word for new word characters
                If charAfter = "" Or charAfter = " " Or charAfter = ")" Or charAfter = "(" Or charAfter = "." Or charAfter = "," Or charAfter = vbCr Then
                    CharAfter_NewWordFlag = True
                End If


                If CharBefore_NewWordFlag And CharAfter_NewWordFlag Then
                    ' this word is a keyword:
                    sc &= k & ", "
                    counter += 1
                End If

            End If

        Next

        sc = TrimEnd(sc, 2) ' trim last comma from the end

        Return New String() {sc, counter}
    End Function



    Private Sub TextBox1_PreviewKeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.PreviewKeyDownEventArgs) Handles txtSQL.PreviewKeyDown
        If e.KeyData = Keys.Tab Then ' captures tab press and allows tabbing inside the text box
            e.IsInputKey = True
        End If
    End Sub


    'Private Sub txtSQL_KeyPress(sender As Object, e As System.Windows.Forms.KeyPressEventArgs) Handles txtSQL.KeyPress
    '    If e.KeyChar = Convert.ToChar(1) Then
    '        DirectCast(sender, TextBox).SelectAll()
    '        e.Handled = True
    '    End If
    'End Sub

    Private Sub txtDatabaseName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtDatabaseName.TextChanged
        If SelectedDriver = 1 Then
            RaiseEvent DatabaseNameChanged(New textChangedArgs(txtDatabaseName.Text))
        End If
    End Sub

    Private Sub txtDriverName_TextChanged(sender As Object, e As EventArgs) Handles txtDriverName.TextChanged
        If SelectedDriver = 0 Then
            RaiseEvent DatabaseNameChanged(New textChangedArgs(txtDriverName.Text))
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

    Private Sub CURRENTDATEToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles CURRENTDATEToolStripMenuItem1.Click
        txtSQL.Paste("CURRENT_DATE")
    End Sub

    Private Sub CURRENTTIMEToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles CURRENTTIMEToolStripMenuItem1.Click
        txtSQL.Paste("CURRENT_TIME")
    End Sub

    Private Sub CURRENTTIMESTAMPToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles CURRENTTIMESTAMPToolStripMenuItem1.Click
        txtSQL.Paste("CURRENT_TIMESTAMP")
    End Sub

    Private Sub TODAYToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs)
        txtSQL.Paste("TODAY")
    End Sub

    Private Sub DATEVAL01302011ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DATEVAL01302011ToolStripMenuItem.Click
        txtSQL.Paste("DATEVAL(")
    End Sub

    Private Sub DateToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DateToolStripMenuItem.Click
        txtSQL.Paste("{d '2001-01-01'}")
    End Sub

    Private Sub TimeToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles TimeToolStripMenuItem.Click
        txtSQL.Paste("{t '01:01:00'}")
    End Sub

    Private Sub TimestampToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles TimestampToolStripMenuItem.Click
        txtSQL.Paste("{ts '2001-01-01 01:01:00'}")
    End Sub

    Private Sub DAYOFWEEKToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DAYOFWEEKToolStripMenuItem.Click
        txtSQL.Paste("DAYOFWEEK(")
    End Sub

    Private Sub CHR67ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles CHR67ToolStripMenuItem.Click
        txtSQL.Paste("CHR()")
    End Sub

    Private Sub UPPERToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles UPPERToolStripMenuItem.Click
        txtSQL.Paste("UPPPER()")
    End Sub

    Private Sub LOWERToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles LOWERToolStripMenuItem.Click
        txtSQL.Paste("LOWER()")
    End Sub

    Private Sub ROUND1234560ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles ROUND1234560ToolStripMenuItem.Click
        txtSQL.Paste("ROUND()")
    End Sub

    Private Sub LENstringToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles LENstringToolStripMenuItem.Click
        txtSQL.Paste("LEN()")
    End Sub

    Private Sub MONTHNAMEToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles MONTHNAMEToolStripMenuItem.Click
        txtSQL.Paste("MONTHNAME()")
    End Sub

    Private Sub USERNAMEToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles USERNAMEToolStripMenuItem.Click
        txtSQL.Paste("USERNAME()")
    End Sub

    Private Sub DAYNAMEToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles DAYNAMEToolStripMenuItem.Click
        txtSQL.Paste("DAYNAME()")
    End Sub

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

    Private Sub DataGridView1_MouseClick(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles DataGridView1.MouseClick
        'Dim s As DataGridView = sender
        '_ColumnRightClicked = s.HitTest(e.Location.X, e.Location.Y).ColumnIndex
        '_RowRightClicked = s.HitTest(e.Location.X, e.Location.Y).RowIndex

        'If e.Button = Windows.Forms.MouseButtons.Right Then
        '    'If _ColumnRightCLicked <> -1 Then s.Columns(_ColumnRightCLicked).Selected = True
        '    ContextMenuStrip2.Show(s.PointToScreen(e.Location))
        'End If
    End Sub

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
            'Dim f = DataGridView1.GetClipboardContent
            ClipboardCopy = True
            System.Windows.Forms.Clipboard.SetDataObject(DataGridView1.GetClipboardContent)
            ClipboardCopy = False
            'Console.WriteLine(Clipboard.GetText())
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

        'If e.Button = Windows.Forms.MouseButtons.Right Then
        '    ContextMenuStrip2.Show(s.PointToScreen(e.Location))

        '    If DataGridView1.SelectedCells.Count <= 1 Then
        '        If _ColumnRightClicked <> -1 And _RowRightClicked <> -1 Then
        '            DataGridView1.ClearSelection()
        '            DataGridView1.Rows(_RowRightClicked).Cells(_ColumnRightClicked).Selected = True
        '        End If
        '    End If

        'End If

    End Sub

    Private Sub mnuPaste_Click(sender As System.Object, e As System.EventArgs) Handles mnuPaste.Click
        txtSQL.SelectedText = ""

        Dim s As String = System.Windows.Forms.Clipboard.GetText

        ' replace dodgy newlines (otherwise can stuff up querie as driver can only recognise CRLF 
        s = s.Replace(Environment.NewLine, vbCr)
        s = s.Replace(vbLf, vbCr)
        s = s.Replace(vbCr, Environment.NewLine)

        txtSQL.Paste(s)
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

    Private Sub ToolStripMenuItem1_Click(sender As System.Object, e As System.EventArgs) Handles mnuSelectAll.Click
        DirectCast(txtSQL, TextBox).SelectAll()
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As System.Object, e As System.Windows.Forms.DataGridViewCellEventArgs) Handles DataGridView1.CellContentClick

    End Sub

    Private Sub Button1_Click(sender As System.Object, e As System.EventArgs) Handles Button1.Click
        Dim f As String = <![CDATA[
   select -- this is a comment

-- this is another comment
'', 'chin
', 'fred''s', f FROM s 

-- this is another comment

where s = 'ssds

d''sdsd'

_jobid BETWEEN 10 AND 50 BATCH 10
            ]]>.Value
        Dim g = ""
        'FormatForFileMaker2(f, 1, g)
    End Sub

    Private Sub SUBSTRINGstring23ToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles SUBSTRINGstring23ToolStripMenuItem.Click
        txtSQL.Paste("SUBSTR('string', 2, 3)")
    End Sub


    Private Sub AboutFMODBCDebuggerToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles AboutFMODBCDebuggerToolStripMenuItem.Click
        About.ShowDialog(Me)
    End Sub

    Private Sub mnuExportToExcel_Click(sender As Object, e As EventArgs) Handles mnuExportToExcel.Click

        Dim rows As New List(Of List(Of String))
        Dim ColumnHeadings As New List(Of String)

        For Each i In DataGridView1.Rows
            Dim Cells As New List(Of String)
            For Each c In i.cells
                Cells.Add(c.value.ToString)
            Next
            rows.Add(Cells)
        Next

        For Each i In DataGridView1.Columns
            ColumnHeadings.Add(i.DataPropertyName)
        Next

        ExportToExcel(ColumnHeadings, rows)
        'Console.WriteLine("done")
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

    Private Sub url_sql_v17_Click(sender As Object, e As EventArgs) Handles url_sql_v17.Click
        Process.Start("https://fmhelp.filemaker.com/docs/17/en/fm17_sql_reference.pdf")
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

    Private Sub url_odbc_v17_Click(sender As Object, e As EventArgs) Handles url_odbc_v17.Click
        Process.Start("https://fmhelp.filemaker.com/docs/17/en/fm16_odbc_jdbc_guide.pdf")
    End Sub

    Private Sub cmbDriverName_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbSelectedDriver.SelectedIndexChanged
        If SelectedDriver = 0 Then ' other
            Panel_Driver_Custom.Visible = True
            Panel_Driver_FileMaker.Visible = False
            RaiseEvent DatabaseNameChanged(New textChangedArgs(txtDriverName.Text))
        Else
            Panel_Driver_Custom.Visible = False
            Panel_Driver_FileMaker.Visible = True
            RaiseEvent DatabaseNameChanged(New textChangedArgs(txtDatabaseName.Text))
        End If
    End Sub

    Private Sub Button2_MouseUp(sender As Object, e As MouseEventArgs) Handles btnGetMetaData1.MouseUp
        Dim p As New System.Drawing.Point(e.X, e.Y)
        ContextMenuStrip3.Show(btnGetMetaData1.PointToScreen(p))
    End Sub

    Private Sub GetTableNamesToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GetTableNamesToolStripMenuItem.Click
        Dim r = GetTableNames()
        If TypeOf (r) Is String Then
            MsgBox(r)
        Else
            Dim s As String = ""
            For Each t In r
                s &= vbNewLine & t
            Next
            txtSQL.Text &= s
        End If
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

    Private Sub txtFind_TextChanged(sender As Object, e As EventArgs) Handles txtFind.TextChanged

    End Sub

    Public Sub focusfind()
        txtFind.Focus()
    End Sub

    Private Sub SplitContainer1_SplitterMoved(sender As Object, e As SplitterEventArgs) Handles SplitContainer1.SplitterMoved

    End Sub

    Private Sub SplitContainer1_MouseUp(sender As Object, e As MouseEventArgs) Handles SplitContainer1.MouseUp
        lblDuration2.Focus()
    End Sub

    Private Sub dataGridView1_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) Handles DataGridView1.CellFormatting

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
        Else
            Dim f = ""
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
        txtFind.Focus()
    End Sub

End Class

Public Class textChangedArgs
    Public _NewTextValue As String = ""
    Public Sub New(ByVal NewTextValue As String)
        _NewTextValue = NewTextValue
    End Sub
End Class

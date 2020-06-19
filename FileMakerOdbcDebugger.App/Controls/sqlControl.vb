Imports System.ComponentModel
Imports FastColoredTextBoxNS
Imports FileMakerOdbcDebugger.Util
Imports System.IO
Imports System.Threading.Tasks

Public Class SqlControl

    Public Event TitleChanged(ByVal e As TitleChangedArgs)
    Public Event BeginExecute()

    Private _Query As String = ""                       ' The query to be executed.

    Private _Result_Data As New List(Of ArrayList)      ' The data returned from the query.
    Private _Result_Error As String = ""                ' Any errors, or it will be empty
    Private _Result_RecordsAffected As Integer = -1
    Private _Result_Duration_Connect As New TimeSpan
    Private _Result_Duration_Execute As New TimeSpan
    Private _Result_Duration_Stream As New TimeSpan

    Private _Search_Text As String = ""
    Private _Search_CurrentCell As Integer = 0
    Private _Search_FoundCells As New List(Of DataGridViewCell)

    Private Sub Me_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        txtErrorText.Dock = DockStyle.Fill
        txtErrorText.Text = ""
        lblStatus.Text = ""
        lblDurationConnect.Visible = False
        lblDurationExecute.Visible = False
        lblDurationStream.Visible = False
        lblWarning.Visible = False
        txtSQL.Select()

        InitaliseTextBoxSettings(txtSQL)

        DisableSearchBox()

        ' Hack to hide the row selection arrow icon
        ' Refer: https://stackoverflow.com/a/5825458/737393
        dgvResults.RowHeadersDefaultCellStyle.Padding = New Padding(dgvResults.RowHeadersWidth)

        dgvResults.RowHeadersWidth = 35

    End Sub

    Private Enum ExecuteStatus

        <Description("Connecting...")>
        Connecting

        <Description("Executing Query. Waiting for response...")>
        Executing

        <Description("Streaming Data...")>
        Streaming

        <Description("Error")>
        [Error]

    End Enum



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

    Public Async Function Execute() As Threading.Tasks.Task

        Query_Prepare()

        Dim ConnString = GetConnectionString()
        Dim RowsToReturn = Me.RowsToReturn

        Await Task.Run(Sub() Query_Execute(ConnString, RowsToReturn))

        Query_DisplayResults()

    End Function

    Public Sub Query_Prepare()

        Try

            ' Initalise UI
            dgvResults.DataSource = Nothing
            txtErrorText.Visible = False
            'lblReservedWords.Text = ""
            'lblKeywords.Visible = False
            lblDurationConnect.Visible = False
            lblDurationExecute.Visible = False
            lblDurationStream.Visible = False
            lblFindCount.Visible = False
            _Search_Text = ""
            DisableSearchBox()

            lblStatus.Text = ""
            ToolTip1.SetToolTip(lblWarning, "")
            lblWarning.Visible = False
            Panel1.Enabled = False
            Me.Refresh()

            lblLoading.Visible = True

            _Result_Data.Clear()
            _Result_Error = ""
            _Result_RecordsAffected = 0
            _Result_Duration_Connect = New TimeSpan
            _Result_Duration_Execute = New TimeSpan
            _Result_Duration_Stream = New TimeSpan


            If txtSQL.SelectionLength > 0 Then
                _Query = txtSQL.SelectedText ' Execute selected text.
            Else
                _Query = txtSQL.Text ' Execute all.
            End If

            If SelectedDriver = SelectedDriverIndex.FileMaker Then

                If String.IsNullOrWhiteSpace(ServerAddress) Then
                    _Result_Error = ExecuteError.NullServer.Description
                    Return
                End If

                If String.IsNullOrWhiteSpace(DatabaseName) Then
                    _Result_Error = ExecuteError.NullDatabaseName.Description
                    Return
                End If

                If String.IsNullOrWhiteSpace(Username) Then
                    _Result_Error = ExecuteError.NullUsername.Description
                    Return
                End If

            End If

            If SelectedDriver = SelectedDriverIndex.Other Then

                If String.IsNullOrWhiteSpace(txtDriverName.Text) Then
                    _Result_Error = ExecuteError.NullDriverName.Description
                    Return
                End If

                If String.IsNullOrWhiteSpace(ConnectionString) Then
                    _Result_Error = ExecuteError.NullConnectionString.Description
                    Return
                End If

            End If

            If SelectedDriver = SelectedDriverIndex.FileMaker Then

                _Query = FM_PrepSQL(_Query)

                UI_SetQueryWarnings(FileMaker.CheckQueryForIssues(_Query))

            End If

            ' Debugging
            Console.WriteLine(_Query)
            File.AppendAllText(Path_LastQuery, _Query)

        Catch ex As Exception
            ParentForm.ShowMsg(ex.Message, MessageBoxIcon.Error)
        End Try

    End Sub

    Public Sub Query_Execute(ConnectionString As String, RowsToReturn As Integer)

        Dim sw As New System.Diagnostics.Stopwatch

        Dim Queries = Util.Sql.SplitQueryIntoStatements(_Query)

        Try

            If Queries.Count > 1 Then

                ' Execute transaction.
                ' Don't show results in data grid.

                Using cn As New Odbc.OdbcConnection(GetConnectionString())

                    UI_SetStatus(ExecuteStatus.Connecting.Description())

                    sw.Restart()

                    cn.Open()

                    _Result_Duration_Connect = sw.Elapsed

                    UI_SetStatus(ExecuteStatus.Executing.Description())

                    sw.Restart()

                    Util.Sql.ExecuteTransaction(Queries, cn)

                    _Result_Duration_Execute = sw.Elapsed

                    _Result_RecordsAffected = 0

                End Using

            Else

                ' Execute a single query.
                ' Show results in data grid.

                If String.IsNullOrEmpty(_Query) Then
                    Return
                End If

                Using cn As New Odbc.OdbcConnection(ConnectionString)

                    UI_SetStatus(ExecuteStatus.Connecting.Description())

                    sw.Restart()

                    cn.Open()

                    _Result_Duration_Connect = sw.Elapsed


                    Dim Data As New List(Of ArrayList)

                    Using cmd As New Odbc.OdbcCommand(_Query, cn)

                        UI_SetStatus(ExecuteStatus.Executing.Description())

                        sw.Restart()

                        Using reader = cmd.ExecuteReader()

                            _Result_Duration_Execute = sw.Elapsed

                            UI_SetStatus(ExecuteStatus.Streaming.Description())

                            sw.Restart()

                            Dim s As New ArrayList

                            ' Read column headers:
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

                            Data.Add(s)

                            ' Read column data:
                            Dim cc As Integer = 1
                            Do While reader.Read() And cc <= RowsToReturn
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
                                Data.Add(s)
                            Loop

                            _Result_Duration_Stream = sw.Elapsed

                            UI_SetStatus("")

                            _Result_RecordsAffected = reader.RecordsAffected

                        End Using

                    End Using

                    _Result_Data = Data

                End Using

            End If

        Catch ex1 As ArgumentOutOfRangeException
            If ex1.Message = "Year, Month, and Day parameters describe an un-representable DateTime." Then ' filemaker allows importing incorrect data into fields, so we need to catch these errors!
                _Result_Error = ExecuteError.UnrepresentableDateTimeValue.Description
            Else
                _Result_Error = ex1.Message
            End If
            _Result_Error &= vbNewLine & vbNewLine & _Query
            Return

        Catch ex As Exception
            _Result_Error = ex.Message
            _Result_Error &= vbNewLine & vbNewLine & _Query
            Return

        End Try

    End Sub

    Public Sub Query_DisplayResults()

        Panel1.Enabled = True
        lblLoading.Visible = False

        If Not String.IsNullOrEmpty(_Result_Error) Then

            lblStatus.Text = ExecuteStatus.Error.Description
            txtErrorText.Visible = True
            txtErrorText.Text = _Result_Error
            DisableSearchBox()
            Return

        End If

        If _Result_Data.Count > 0 Then

            ' Note: its faster if we populate the datagrid view using a DataSet.
            Dim DS As New DataSet
            DS.Tables.Add("SELECT")
            Dim t = DS.Tables("SELECT")


            For i = 0 To _Result_Data(0).Count - 1
                t.Columns.Add(_Result_Data(0)(i))
            Next

            If _Result_Data.Count > 1 Then
                Dim row As System.Data.DataRow

                For i = 1 To _Result_Data.Count - 1
                    row = t.NewRow

                    Dim s(_Result_Data(1).Count - 1) As Object  ' initalise object with same number of rows

                    For j = 0 To _Result_Data(i).Count - 1
                        row.Item(j) = _Result_Data(i)(j)
                    Next

                    t.Rows.Add(row)

                    If i > RowsToReturn Then
                        Exit For
                    End If

                Next

                dgvResults.DataSource = t

            End If

            If dgvResults.Rows.Count > 0 Then

                If txtFind.Text.Count < 2 Then ' select first cell if we are not already searching for something
                    dgvResults.Rows(0).Cells(0).Selected = True
                End If

            End If

            For Each c In dgvResults.Columns
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

        UI_RenderRowCount()

        ' Show duration:
        lblDurationConnect.Visible = True
        lblDurationExecute.Visible = True
        lblDurationStream.Visible = True
        lblDurationConnect.Text = "Connect: " & FormatTime(_Result_Duration_Connect)
        lblDurationExecute.Text = "Execute: " & FormatTime(_Result_Duration_Execute)
        lblDurationStream.Text = "Stream: " & FormatTime(_Result_Duration_Stream)

        ' Show/hide status:
        If lblStatus.Text = "" Then
            lblStatus.Visible = False
        Else
            lblStatus.Visible = True
        End If

    End Sub

    Public Delegate Sub UI_SetStatus_Delegate(Status As String)
    Private Sub UI_SetStatus(Status As String)

        If Me.InvokeRequired Then
            Dim d As New UI_SetStatus_Delegate(AddressOf UI_SetStatus)
            Me.Invoke(d, {Status})
            Return
        End If

        lblStatus.Text = Status

    End Sub

    Public Sub UI_SetQueryWarnings(ByVal Warnings As List(Of FileMaker.QueryIssue))

        If Warnings.Count = 0 Then
            lblWarning.Visible = False
            Return
        End If

        Dim s As String = ""
        For Each i In Warnings
            s &= "• " & i.Description & vbNewLine
        Next
        s = s.Trim

        ToolTip1.SetToolTip(lblWarning, s)
        lblWarning.Visible = True

    End Sub

    Private Function GetConnectionString() As String

        If SelectedDriver = SelectedDriverIndex.Other Then
            Return "DRIVER={" & txtDriverName.Text & "};" & txtConnectionString.Text
        Else
            Return "DRIVER={FileMaker ODBC};SERVER=" & ServerAddress & ";UID=" & Username & ";PWD=" & Password & ";DATABASE=" & DatabaseName & ";"
        End If

    End Function

    Private Sub dgvResults_DataBindingComplete(ByVal sender As Object, ByVal e As DataGridViewBindingCompleteEventArgs) Handles dgvResults.DataBindingComplete
        ' Re-index the search results after new data is loaded.
        Search(0)
    End Sub

    Public Property ServerAddress() As String
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

    Public Property Username() As String
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

    Public Property ConnectionString() As String
        Get
            Return txtConnectionString.Text
        End Get
        Set(ByVal value As String)
            txtConnectionString.Text = value
        End Set
    End Property

    Public Enum SelectedDriverIndex
        Other = 0
        FileMaker = 1
    End Enum

    Public Property SelectedDriver() As SelectedDriverIndex
        Get

            If cmbDriver.Text.ToLower = "other" Then
                Return SelectedDriverIndex.Other
            Else
                Return SelectedDriverIndex.FileMaker
            End If

        End Get
        Set(ByVal value As SelectedDriverIndex)

            If value = SelectedDriverIndex.Other Then
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

    Private Sub UI_RenderRowCount()

        Dim TotalRows As Integer = _Result_Data.Count - 1

        If TotalRows = 0 And _Result_RecordsAffected > -1 Then

            ' Show number records affected:
            lblStatus.Text = _Result_RecordsAffected & " Records Affected"

        Else

            ' Show a "max" label to remind that the rows have been limited.
            Dim MaxLabel As String = ""
            If TotalRows = RowsToReturn Then
                MaxLabel = " (max)"
            End If

            ' Show row count.
            lblStatus.Text = $"{TotalRows} rows{MaxLabel}"

        End If

    End Sub

    Private Async Sub btnGo_Click(sender As System.Object, e As System.EventArgs) Handles btnGo.Click

        RaiseEvent BeginExecute()
        Await Execute()

    End Sub

    Private Sub txtDatabaseName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtDatabaseName.TextChanged

        UI_RenderTitle()

    End Sub

    Private Sub txtDriverName_TextChanged(sender As Object, e As EventArgs) Handles txtDriverName.TextChanged

        UI_RenderTitle()

    End Sub

    Private Sub UI_RenderTitle()

        If SelectedDriver = SelectedDriverIndex.Other Then
            RaiseEvent TitleChanged(New TitleChangedArgs(txtDriverName.Text))
        Else
            RaiseEvent TitleChanged(New TitleChangedArgs(txtDatabaseName.Text))
        End If

    End Sub

    Private Sub txtCredentails_Changed(sender As System.Object, e As System.EventArgs) Handles txtServerIP.TextChanged, txtDatabaseName.TextChanged, txtUsername.TextChanged

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

        If dgvResults.Rows.Count > 0 Then
            If _ColumnRightClicked <> -1 Then
                mnuCopyAsCSV.Enabled = True
            End If

            If dgvResults.SelectedCells.Count > 0 Then
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
            For Each r As DataGridViewRow In dgvResults.Rows
                Dim value = r.Cells(_ColumnRightClicked).Value
                If Not IsDBNull(value) Then
                    s &= value.ToString & ","
                End If
            Next
            Clipboard.SetText(s.TrimNCharsFromEnd(1))
        Catch ex As Exception
            ParentForm.ShowMsg(ex.Message)
        End Try
    End Sub

    Private ClipboardCopy As Boolean = False
    Private Sub CopyCellToolStripMenuItem_Click(sender As System.Object, e As System.EventArgs) Handles mnuCopy.Click
        Try
            ClipboardCopy = True
            Clipboard.SetDataObject(dgvResults.GetClipboardContent)
            ClipboardCopy = False
        Catch ex As Exception
        End Try
    End Sub

    Private Sub mnuCopyWithHeaders_Click(sender As Object, e As EventArgs) Handles mnuCopyWithHeaders.Click
        Try
            dgvResults.RowHeadersVisible = False
            dgvResults.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText
            ClipboardCopy = True
            Clipboard.SetDataObject(dgvResults.GetClipboardContent)
            ClipboardCopy = False
            dgvResults.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableWithoutHeaderText
            dgvResults.RowHeadersVisible = True
        Catch ex As Exception
        End Try
    End Sub

    Private Sub dgvResults_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgvResults.MouseDown

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
        For Each i In dgvResults.Columns
            ColumnHeadings.Add(i.DataPropertyName)
        Next

        Dim Data As New List(Of List(Of String))
        Data.Add(ColumnHeadings)

        For Each i In dgvResults.Rows
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

        If SelectedDriver = SelectedDriverIndex.Other Then
            Panel_Driver_Custom.Visible = True
            Panel_Driver_FileMaker.Visible = False
        Else
            Panel_Driver_Custom.Visible = False
            Panel_Driver_FileMaker.Visible = True
        End If

        UI_RenderTitle()

    End Sub

    Private Sub txtSQL_TextChanged(ByVal sender As Object, ByVal e As TextChangedEventArgs) Handles txtSQL.TextChanged
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

    Private Sub Search(ByVal FindMode As Integer)
        Try

            If FindMode = 1 Then ' FIND NEXT
                _Search_CurrentCell += 1

                If _Search_CurrentCell > _Search_FoundCells.Count - 1 Then
                    _Search_CurrentCell = 0
                End If

                Search_HighlightNext()
                Return

            ElseIf FindMode = 2 Then ' FIND PREV
                _Search_CurrentCell -= 1

                If _Search_CurrentCell < 0 Then
                    _Search_CurrentCell = _Search_FoundCells.Count - 1
                End If

                Search_HighlightNext()
                Return
            End If

            Dim t As String = txtFind.Text.ToLower

            If t.StartsWith(_Search_Text) And _Search_Text <> "" And _Search_FoundCells.Count = 0 Then t = _Search_Text ' only if different than last 0-result search

            If _Search_Text <> t Then
                ' re-index items

                _Search_CurrentCell = 0
                _Search_FoundCells.Clear()

                If Not String.IsNullOrEmpty(t) Then
                    For i = 0 To dgvResults.Rows.Count - 1
                        Dim ro = dgvResults.Rows(i)
                        For c = 0 To ro.Cells.Count - 1
                            Dim ce = ro.Cells(c)

                            'Console.WriteLine(ce.Value.ToString)
                            If ce.Value.ToString.ToLower.Contains(t) Then
                                _Search_FoundCells.Add(ce)
                            End If
                        Next
                    Next
                End If

                _Search_Text = t

                ' show found count
                If t.Length > 0 Then
                    lblFindCount.Text = _Search_FoundCells.Count
                    lblFindCount.Visible = True

                    If _Search_FoundCells.Count > 0 Then
                        lblFindCount.BackColor = Color.White
                    Else
                        lblFindCount.BackColor = Color.LightSalmon
                    End If
                Else
                    lblFindCount.Visible = False
                End If

            End If

            If dgvResults.Rows.Count = 0 Then
                DisableSearchBox()
            Else
                EnableSearchBox()
            End If

            Search_HighlightNext()

        Catch ex As Exception

        End Try
    End Sub

    Private Sub Search_HighlightNext()
        If _Search_FoundCells.Count > 0 Then
            dgvResults.CurrentCell = _Search_FoundCells(_Search_CurrentCell)
        End If
    End Sub

    Public Sub UI_FocusFindTextbox()
        txtFind.Focus()
    End Sub

    Private Sub SplitContainer1_MouseUp(sender As Object, e As MouseEventArgs) Handles SplitContainer1.MouseUp
        lblDurationConnect.Focus()
    End Sub

    Private Sub dgvResults_CellFormatting(ByVal sender As Object, ByVal e As DataGridViewCellFormattingEventArgs) Handles dgvResults.CellFormatting

        ' this is called when the clipboard retrieves data from the cell as well

        Dim c = dgvResults.Rows(e.RowIndex).Cells(e.ColumnIndex)

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
        UI_FocusFindTextbox()
    End Sub

    Private Sub txtSQL_KeyDown(sender As Object, e As KeyEventArgs) Handles txtSQL.KeyDown
        If e.Control And ((e.Shift And e.KeyCode = Keys.Z) Or (e.KeyCode = Keys.Y)) Then
            txtSQL.Redo()
        End If
    End Sub

    Private Sub dgvResults_RowPostPaint(ByVal sender As Object, ByVal e As DataGridViewRowPostPaintEventArgs) Handles dgvResults.RowPostPaint

        ' Paint row numbers on the row header column
        ' Refer: https://stackoverflow.com/a/12840794/737393

        Dim RowNumber = (e.RowIndex + 1).ToString()

        Dim Padding As New Padding(5, 3, 3, 0)

        Dim headerBounds = New Rectangle(e.RowBounds.Left + Padding.Left,
                                         e.RowBounds.Top + Padding.Top,
                                         dgvResults.RowHeadersWidth - Padding.Left - Padding.Right,
                                         e.RowBounds.Height - Padding.Top)

        Dim flags As TextFormatFlags = TextFormatFlags.Left Or TextFormatFlags.EndEllipsis Or TextFormatFlags.ExpandTabs Or TextFormatFlags.SingleLine

        TextRenderer.DrawText(e.Graphics, RowNumber, Me.Font, headerBounds, Me.ForeColor, flags)

    End Sub

End Class

Public Class TitleChangedArgs

    Public ReadOnly NewValue As String = ""

    Public Sub New(ByVal NewTextValue As String)
        Me.NewValue = NewTextValue
    End Sub

End Class

Imports System.ComponentModel
Imports FastColoredTextBoxNS
Imports FileMakerOdbcDebugger.Util
Imports System.IO
Imports System.Threading.Tasks

Public Class SqlControl
    Implements IUi

    Public Event TitleChanged(ByVal e As TitleChangedArgs)
    Public Event BeginExecute()

    Private _Query As String = ""                       ' The query to be executed.
    Private _Statements As List(Of String)                 ' The statements to be executed. 
    Private _Result As Sql.TransactionResult      ' The result of the query.

    Private Sub Me_Load(sender As Object, e As System.EventArgs) Handles Me.Load

        txtSQL.Select()
        InitaliseTextBoxSettings(txtSQL)
        lblLoading.BringToFront() ' Just in case it gets lost in the designer
        SplitContainer1.Panel2Collapsed = True

        Me.DoubleBuffered = True

    End Sub

    Public Async Function Execute() As Threading.Tasks.Task

        If String.IsNullOrWhiteSpace(txtSQL.Text) Then Return

        ' Set UI
        Panel_Results.Controls.Clear()  ' Remove old results
        SplitContainer1.Panel2Collapsed = False
        lblDurationConnect.Visible = False
        lblStatus.Text = ""
        Panel1.Enabled = False
        lblLoading.Visible = True
        RaiseEvent_TitleChanged()
        Me.Refresh()

        Query_Prepare()

        If String.IsNullOrWhiteSpace(_Result.Error) Then ' There may have been errors when preparing the query (see above).

            Await Task.Run(Sub() Query_Execute(ConnectionString:=GetConnectionString(), RowsToReturn:=RowsToReturn))

        End If

        Query_DisplayResults()

        ' Set UI
        Panel1.Enabled = True
        lblLoading.Visible = False
        RaiseEvent_TitleChanged()

    End Function

    Public Sub Query_Prepare()

        Try

            _Result = New Util.Sql.TransactionResult


            If txtSQL.SelectionLength > 0 Then
                _Query = txtSQL.SelectedText ' Execute selected text.
            Else
                _Query = txtSQL.Text ' Execute all.
            End If

            If SelectedDriver = SelectedDriverIndex.Other Then

                If String.IsNullOrWhiteSpace(ConnectionString) Then
                    _Result.Error = Sql.ExecuteError.NullConnectionString.Description
                    Return
                End If

            Else

                If String.IsNullOrWhiteSpace(ServerAddress) Then
                    _Result.Error = Sql.ExecuteError.NullServer.Description
                    Return
                End If

                If String.IsNullOrWhiteSpace(DatabaseName) Then
                    _Result.Error = Sql.ExecuteError.NullDatabaseName.Description
                    Return
                End If

                If String.IsNullOrWhiteSpace(Username) Then
                    _Result.Error = Sql.ExecuteError.NullUsername.Description
                    Return
                End If

                _Query = Sql.PrepareQuery(_Query, SelectedDriver = SelectedDriverIndex.FileMaker)

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

        _Statements = Sql.SplitQueryIntoStatements(_Query)

        _Result = Sql.ExecuteTransaction(_Statements, Me, RowsToReturn, ConnectionString)

    End Sub

    Public Sub Query_DisplayResults()

        lblDurationConnect.Visible = True
        lblDurationConnect.Text = "Connect: " & _Result.Duration_Connect.ToDisplayString

        SplitContainer1.Panel2.Refresh()

        If Not String.IsNullOrEmpty(_Result.Error) Then

            Panel_Results.Controls.Add(CreateDivider(DockStyle.Top))

            Dim c = New TextBox With {
                .Dock = DockStyle.Fill,
                .Font = txtSQL.Font,
                .Text = _Result.Error,
                .BorderStyle = BorderStyle.None,
                .Multiline = True,
                .ScrollBars = ScrollBars.Vertical,
                .BackColor = SystemColors.Control,
                .ReadOnly = True
            }

            Panel_Results.Controls.Add(c)

            lblStatus.Text = Sql.ExecuteStatus.Error.Description

            Return

        End If

        Dim SplitDistance = (Panel_Results.Height / _Result.Results.Count)
        Dim LastSplitContainer As SplitContainer

        For Each r In _Result.Results

            Dim thisSplitContainer As SplitContainer

            If LastSplitContainer Is Nothing Then

                LastSplitContainer = New SplitContainer
                LastSplitContainer.Dock = DockStyle.Fill
                LastSplitContainer.Location = New System.Drawing.Point(0, 1)
                LastSplitContainer.Size = New System.Drawing.Size(50, 50)
                LastSplitContainer.Location = New System.Drawing.Point(0, 0)
                LastSplitContainer.Orientation = System.Windows.Forms.Orientation.Horizontal
                LastSplitContainer.Panel2Collapsed = True

                Panel_Results.Controls.Add(LastSplitContainer)
                thisSplitContainer = LastSplitContainer

            Else

                ' todo: size each split panel evenly

                thisSplitContainer = New SplitContainer
                thisSplitContainer.Dock = DockStyle.Fill
                thisSplitContainer.Orientation = System.Windows.Forms.Orientation.Horizontal
                thisSplitContainer.Panel2Collapsed = True

                LastSplitContainer.Panel2Collapsed = False
                LastSplitContainer.Panel2.Controls.Add(thisSplitContainer)

                LastSplitContainer.Panel1.Controls.Add(CreateDivider())

            End If

            Dim rp As New ResultPanel
            rp.Dock = DockStyle.Fill

            thisSplitContainer.SplitterDistance = SplitDistance

            thisSplitContainer.Panel1.Controls.Add(rp)

            rp.UI_RenderRowCount(r.Data.Count - 1, r.RowsAffected, RowsToReturn)

            rp.UI_RenderQueryWarnings(Sql.CheckQueryForIssues(_Query, SelectedDriver = SelectedDriverIndex.FileMaker))

            Me.Refresh()

            rp.UI_DisplayData(r)

            LastSplitContainer = thisSplitContainer

        Next

    End Sub

    Public Delegate Sub UI_SetStatus_Delegate(Status As String)
    Private Sub UI_SetStatus(Status As String) Implements IUi.SetStatusLabel

        If Me.InvokeRequired Then
            Dim d As New UI_SetStatus_Delegate(AddressOf UI_SetStatus)
            Me.Invoke(d, {Status})
            Return
        End If

        lblStatus.Text = Status

        lblStatus.Visible = (Status <> "")

    End Sub

    Private Function GetConnectionString() As String

        If SelectedDriver = SelectedDriverIndex.Other Then
            Return txtConnectionString.Text
        Else
            Return Common.BuildConnectionString(ServerAddress, Username, Password, DatabaseName)
        End If

    End Function

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

    Private Async Sub btnExecute_Click(sender As System.Object, e As System.EventArgs) Handles btnExecute.Click

        RaiseEvent BeginExecute()
        Await Execute()

    End Sub

    Private Sub txtDatabaseName_TextChanged(sender As System.Object, e As System.EventArgs) Handles txtDatabaseName.TextChanged, txtConnectionString.TextChanged

        RaiseEvent_TitleChanged()

    End Sub

    Private Sub RaiseEvent_TitleChanged()

        Dim NewTitle As String = ""

        ' Prepend asterix to title if query is currently executing:
        If lblLoading.Visible Then
            NewTitle &= "*"
        End If

        If SelectedDriver = SelectedDriverIndex.Other Then
            NewTitle &= txtConnectionString.Text
        Else
            NewTitle &= txtDatabaseName.Text
        End If

        RaiseEvent TitleChanged(New TitleChangedArgs(Me, NewTitle))

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

        RaiseEvent_TitleChanged()

    End Sub

    Private Sub txtSQL_TextChanged(ByVal sender As Object, ByVal e As TextChangedEventArgs) Handles txtSQL.TextChanged
        SetRangeStyle(txtSQL.Range)
    End Sub

    Private Sub SplitContainer1_MouseUp(sender As Object, e As MouseEventArgs) Handles SplitContainer1.MouseUp
        lblDurationConnect.Focus()
    End Sub

    Private Sub txtSQL_KeyDown(sender As Object, e As KeyEventArgs) Handles txtSQL.KeyDown
        If e.Control And ((e.Shift And e.KeyCode = Keys.Z) Or (e.KeyCode = Keys.Y)) Then
            txtSQL.Redo()
        End If
    End Sub

End Class

Public Class TitleChangedArgs

    Public ReadOnly NewValue As String = ""
    Public ReadOnly SqlControl As SqlControl

    Public Sub New(SqlControl As SqlControl, NewTextValue As String)
        Me.SqlControl = SqlControl
        Me.NewValue = NewTextValue
    End Sub

End Class

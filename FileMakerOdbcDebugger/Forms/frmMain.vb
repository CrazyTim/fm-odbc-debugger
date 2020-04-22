Public Class frmMain

    Private _PrevWindowState As FormWindowState
    Private _TabCount As Integer = 0

    Private Sub frmMain_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        SaveSettings()
    End Sub

    ''' <summary> Save tab settings to a json file </summary>
    Public Sub SaveSettings()

        Dim Settings As New List(Of SqlControl.Settings)

        For Each t As TabPage In TabControl1.TabPages
            Dim sqlc As SqlControl = t.Form.Controls(0)
            Settings.Add(sqlc.GetSettings)
        Next

        Try
            WriteToFile(AppDataDir & "\settings.json", Settings.To_JSON)
        Catch ex As Exception
            MsgBox("Error: unable to save settings.", MsgBoxStyle.Critical)
        End Try

    End Sub

    Public Function LoadSettings() As List(Of SqlControl.Settings)
        Dim FileData As String = ReadFromFile(AppDataDir & "\settings.json")
        Return FileData.From_JSON(GetType(List(Of SqlControl.Settings)))
    End Function

    Private Sub frmMain_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F5 Then
            Dim f1 As Form = TabControl1.SelectedForm
            Dim u As SqlControl = f1.Controls(0)
            u.ExecuteSQL()
        End If
    End Sub

    Private Sub sqlControl_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.Control And e.KeyCode = Keys.F Then
            Dim f1 As Form = TabControl1.SelectedForm
            Dim u As SqlControl = f1.Controls(0)
            u.FocusFindTextbox()
        End If
    End Sub

    Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        ' create app data folder
        If Not PathExists(AppDataDir) Then
            Try
                IO.Directory.CreateDirectory(AppDataDir)
            Catch
                MsgBox($"There was a problem creating the folder '{AppDataDir}'.", MsgBoxStyle.Critical, "Error")
            End Try
        End If

        TabControl1.TabCloseButtonImage = My.Resources.btn_close
        TabControl1.TabCloseButtonImageDisabled = My.Resources.btn_close
        TabControl1.TabCloseButtonImageHot = My.Resources.btn_close
        TabControl1.TabCloseButtonSize = New System.Drawing.Point(16, 16)
        TabControl1.FontBoldOnSelect = False

        Try

            ' load saved file
            Dim Settings As List(Of SqlControl.Settings) = LoadSettings()

            If Settings.Count = 0 Then
                CreateNewTab(Nothing, False) ' create default tab
            Else
                Settings.Reverse()
                For Each s In Settings
                    CreateNewTab(s, False)
                Next
            End If

        Catch ex As Exception
            CreateNewTab(Nothing, False) ' create default tab
        End Try

        TabControl1.TabPages(TabControl1.TabPages.Count - 1).Select() ' select first tab

        _PrevWindowState = FormWindowState.Normal

    End Sub

    Private Sub TabControl1_DoubleClick(sender As Object, e As System.EventArgs) Handles TabControl1.DoubleClick
        CreateNewTab(Nothing, True)
    End Sub

    Private Sub CreateNewTab(ByVal Settings As SqlControl.Settings, ByVal Copy As Boolean)

        If TabControl1.TabPages.Count = MAX_TABS Then Return ' limit max tabs (gui will bug out)

        If Settings Is Nothing Then Settings = New SqlControl.Settings

        If Copy Then
            ' copy values from selected tab (if any)
            Dim f1 As Form = TabControl1.SelectedForm
            If f1 IsNot Nothing Then
                Dim sqlc As SqlControl = f1.Controls(0)
                Settings.ServerAddress = sqlc.ServerIP
                Settings.DatabaseName = sqlc.DatabaseName
                Settings.Username = sqlc.UserName
                Settings.Password = sqlc.Password
                Settings.SelectedDriver = sqlc.SelectedDriver
                Settings.RowsToReturn = sqlc.RowsToReturn
                Settings.DriverName = sqlc.DriverName
                Settings.ConnectionString = sqlc.ConnectionString
            End If
        End If

        Dim c As New SqlControl With {
            .Dock = DockStyle.Fill,
            .ServerIP = Settings.ServerAddress,
            .DatabaseName = Settings.DatabaseName,
            .UserName = Settings.Username,
            .Password = Settings.Password,
            .RowsToReturn = Settings.RowsToReturn,
            .DriverName = Settings.DriverName,
            .ConnectionString = Settings.ConnectionString,
            .SQLText = Settings.Query
        }

        c.cmbDriver.Items.Clear()
        c.cmbDriver.Items.Add("FileMaker ODBC (" & FmOdbcDriverVersion & ")")
        c.cmbDriver.Items.Add("Other")

        AddHandler c.DatabaseNameChanged, AddressOf OnDatabaseNameChange
        AddHandler c.BeginSQLExecute, AddressOf OnBeginSQLExecute

        _TabCount += 1
        Dim f As New Form
        f.Text = ""
        f.Controls.Add(c)

        TabControl1.TabPages.Add(f)

        c.SelectedDriver = Settings.SelectedDriver

    End Sub

    Private Sub OnDatabaseNameChange(ByVal e As TextChangedArgs)
        TabControl1.TabPages.SelectedTab.Form.Text = e._NewTextValue
    End Sub

    Private Sub OnBeginSQLExecute()
        SaveSettings()
    End Sub

    Private Sub TabControl1_SelectedTabChanged(sender As Object, e As System.EventArgs) Handles TabControl1.SelectedTabChanged
        If TabControl1.TabPages.Count = 0 Then
            Application.Exit()
        End If
    End Sub

    Private Sub frmMain_SizeChanged(sender As Object, e As System.EventArgs) Handles Me.SizeChanged

        ' fix a bug in tab control where the last tab does not position properly after being minimised
        If _PrevWindowState = FormWindowState.Minimized Then
            TabControl1.ArrangeItems()
        End If

        _PrevWindowState = Me.WindowState
    End Sub

End Class

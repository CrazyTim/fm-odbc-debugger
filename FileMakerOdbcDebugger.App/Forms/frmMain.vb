Imports FileMakerOdbcDebugger.Util.Extensions

Public Class frmMain

    Private _PrevWindowState As FormWindowState
    Private _TabCount As Integer = 0

    Private Sub frmMain_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing

        Try
            SaveSettings(TabControl1)
        Catch ex As Exception
            ShowMsg("Unable to save settings.", MessageBoxIcon.Error)
        End Try

    End Sub

    Private Async Sub frmMain_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown

        ' Execute query:
        If e.KeyCode = Keys.F5 Then
            Dim tab As Form = TabControl1.SelectedForm
            Dim u As SqlControl = tab.Controls(0)
            u.Execute()
        End If

    End Sub

    Private Sub frmMain_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load

        ' Create app data folder:
        If Not Util.IO.PathExists(Path_AppData) Then

            Try
                IO.Directory.CreateDirectory(Path_AppData)
            Catch
                ShowMsg($"There was a problem creating the folder '{Path_AppData}'.", MessageBoxIcon.Error)
            End Try

        End If

        ' Load settings:
        Dim Settings As New List(Of Settings)

        Try
            Settings = LoadSettings()
        Catch ex As Exception
            ShowMsg("Settings are corrupt and could not be loaded.", MessageBoxIcon.Error)
        End Try

        ' Create tabs:
        Try

            If Settings.Count = 0 Then

                ' Create a single default tab.
                UI_CreateTab(Nothing, False)

            Else

                ' Recreate each tab from the settings.
                Settings.Reverse()
                For Each s In Settings
                    UI_CreateTab(s, False)
                Next

            End If

        Catch ex As Exception
            UI_CreateTab(Nothing, False) ' create default tab
        End Try

        ' Style the tab control:
        TabControl1.TabCloseButtonImage = My.Resources.btn_close
        TabControl1.TabCloseButtonImageDisabled = My.Resources.btn_close
        TabControl1.TabCloseButtonImageHot = My.Resources.btn_close
        TabControl1.TabCloseButtonSize = New System.Drawing.Point(16, 16)
        TabControl1.TabMinimumWidth = 80
        TabControl1.TabMaximumWidth = 150
        TabControl1.FontBoldOnSelect = False
        TabControl1.TabPages(TabControl1.TabPages.Count - 1).Select() ' select first tab

        _PrevWindowState = FormWindowState.Normal

    End Sub

    Private Sub TabControl1_DoubleClick(sender As Object, e As System.EventArgs) Handles TabControl1.DoubleClick

        UI_CreateTab(Nothing, True)

    End Sub

    Private Sub UI_CreateTab(ByVal Settings As Settings, ByVal Copy As Boolean)

        If TabControl1.TabPages.Count = MAX_TABS Then Return

        If Settings Is Nothing Then Settings = New Settings

        If Copy Then
            ' copy values from selected tab (if any)
            Dim f1 As Form = TabControl1.SelectedForm
            If f1 IsNot Nothing Then
                Dim sqlc As SqlControl = f1.Controls(0)
                Settings.ServerAddress = sqlc.ServerAddress
                Settings.DatabaseName = sqlc.DatabaseName
                Settings.Username = sqlc.Username
                Settings.Password = sqlc.Password
                Settings.SelectedDriver = sqlc.SelectedDriver
                Settings.RowsToReturn = sqlc.RowsToReturn
                Settings.ConnectionString = sqlc.ConnectionString
            End If
        End If

        Dim c As New SqlControl With {
            .Dock = DockStyle.Fill,
            .ServerAddress = Settings.ServerAddress,
            .DatabaseName = Settings.DatabaseName,
            .Username = Settings.Username,
            .Password = Settings.Password,
            .RowsToReturn = Settings.RowsToReturn,
            .ConnectionString = Settings.ConnectionString,
            .SQLText = Settings.Query
        }

        Dim DriverVersion As String = FmOdbcDriverVersion
        If String.IsNullOrWhiteSpace(DriverVersion) Then
            DriverVersion = "driver not installed"
        End If

        c.cmbDriver.Items.Clear()
        c.cmbDriver.Items.Add("FileMaker ODBC (" & DriverVersion & ")")
        c.cmbDriver.Items.Add("Other")

        AddHandler c.TitleChanged, AddressOf Tab_OnDatabaseNameChange
        AddHandler c.BeginExecute, AddressOf Tab_OnBeginExecute

        _TabCount += 1
        Dim f As New Form
        f.Text = ""
        f.Controls.Add(c)

        TabControl1.TabPages.Add(f)

        c.SelectedDriver = Settings.SelectedDriver

    End Sub

    Private Sub Tab_OnDatabaseNameChange(ByVal e As TitleChangedArgs)

        e.SqlControl.Parent.Text = e.NewValue

    End Sub

    Private Sub Tab_OnBeginExecute()

        Try
            SaveSettings(TabControl1)
        Catch ex As Exception
            ShowMsg("Unable to save settings.", MessageBoxIcon.Error)
        End Try

    End Sub

    Private Sub TabControl1_SelectedTabChanged(sender As Object, e As System.EventArgs) Handles TabControl1.SelectedTabChanged

        If TabControl1.TabPages.Count = 0 Then
            Application.Exit()
        End If

    End Sub

    Private Sub frmMain_SizeChanged(sender As Object, e As System.EventArgs) Handles Me.SizeChanged

        ' Fix bug in tab control where the last tab does not position properly after being minimised:
        If _PrevWindowState = FormWindowState.Minimized Then
            TabControl1.ArrangeItems()
        End If

        _PrevWindowState = Me.WindowState

    End Sub

End Class

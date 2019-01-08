Public Class Form1

    Private _PrevWindowState As FormWindowState
    Private _TabCount As Integer = 0

    Private Sub Form1_FormClosing(sender As Object, e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        SaveSettings()
    End Sub

    Public Sub SaveSettings()
        ' save sql in each tab in a serialised xml file
        Dim SavedTabContent As New ArrayList

        ' get text in each tab, add to 2D array
        For Each t As TabPage In TabControl1.TabPages
            Dim sqlc As sqlControl = t.Form.Controls(0)

            Dim a As New ArrayList
            a.Add(sqlc.txtSQL.Text)
            a.Add(sqlc.ServerIP)
            a.Add(sqlc.DatabaseName)
            a.Add(sqlc.UserName)
            a.Add(sqlc.Password)
            a.Add(sqlc.SelectedDriver)
            a.Add(sqlc.RowsToReturn)
            a.Add(sqlc.DriverName)
            a.Add(sqlc.ConnectionString)

            SavedTabContent.Add(a)
        Next

        ' serialise array
        SaveSerialXML(AppDataDir & "\saved_tabs.xml", SavedTabContent, GetType(ArrayList))
    End Sub

    Private Sub Form1_KeyDown(sender As Object, e As System.Windows.Forms.KeyEventArgs) Handles Me.KeyDown
        If e.KeyCode = Keys.F5 Then
            Dim f1 As Form = TabControl1.SelectedForm
            Dim u As sqlControl = f1.Controls(0)
            u.ExecuteSQL()
        End If
    End Sub

    Private Sub sqlControl_KeyDown(sender As Object, e As KeyEventArgs) Handles Me.KeyDown
        If e.Control And e.KeyCode = Keys.F Then
            Dim f1 As Form = TabControl1.SelectedForm
            Dim u As sqlControl = f1.Controls(0)
            u.focusfind()
        End If
    End Sub

    Private Sub Form1_Load(sender As System.Object, e As System.EventArgs) Handles MyBase.Load
        Me.TabControl1.TabCloseButtonImage = My.Resources.Resources.btn_close
        Me.TabControl1.TabCloseButtonImageDisabled = My.Resources.Resources.btn_close
        Me.TabControl1.TabCloseButtonImageHot = My.Resources.Resources.btn_close
        Me.TabControl1.TabCloseButtonSize = New System.Drawing.Point(16, 16)
        Me.TabControl1.FontBoldOnSelect = False

        CreateDir(AppDataDir)
        CreateDir(AppDataDir & "\Debugging")

        CheckDriverVersion()

        Try

            ' read saved file
            Dim SavedTabContent As ArrayList = ReadSerialXML(AppDataDir & "\saved_tabs.xml", GetType(ArrayList))
            SavedTabContent.Reverse()

            If SavedTabContent Is Nothing OrElse SavedTabContent.Count = 0 Then
                CreateNewTab(, , , , , , , , , False)
            Else
                ' create tabs from saved file
                For Each s As ArrayList In SavedTabContent
                    CreateNewTab(s(0), s(1), s(2), s(3), s(4), s(5), s(6), s(7), s(8), False)
                Next
            End If

        Catch ex As Exception
            ' invalid saved data
            CreateNewTab(, , , , , , , , , False)
        End Try

        ' select first tab
        TabControl1.TabPages(TabControl1.TabPages.Count - 1).Select()

        _PrevWindowState = FormWindowState.Normal

    End Sub

    Public FMODBCVersion As String = ""



    Public Sub CheckDriverVersion()
        Dim PathToDll As String
        Dim SYSTEM_DRIVE As String = Environment.GetEnvironmentVariable("windir", EnvironmentVariableTarget.Machine)

        If System.IO.Directory.Exists(SYSTEM_DRIVE & "\SysWOW64") Then

            PathToDll = SYSTEM_DRIVE & "\SysWOW64\fmodbc32.DLL"

            If System.IO.File.Exists(PathToDll) Then
                Dim myFileVersionInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(PathToDll)
                FMODBCDriverVersion = myFileVersionInfo.FileMajorPart & "." & myFileVersionInfo.FileMinorPart & "." & myFileVersionInfo.FileBuildPart & "." & myFileVersionInfo.FilePrivatePart

                FMODBCVersion &= "v" & FMODBCDriverVersion & " 32bit"
            Else
                FMODBCVersion &= "driver not installed"
            End If

        Else

            PathToDll = SYSTEM_DRIVE & "\System32\fmodbc32.DLL"

            If System.IO.File.Exists(PathToDll) Then
                Dim myFileVersionInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(PathToDll)
                FMODBCDriverVersion = myFileVersionInfo.FileMajorPart & "." & myFileVersionInfo.FileMinorPart & "." & myFileVersionInfo.FileBuildPart & "." & myFileVersionInfo.FilePrivatePart
                FMODBCVersion &= "v" & FMODBCDriverVersion & "64bit"
            Else
                FMODBCVersion &= "driver not installed"
            End If

        End If

    End Sub


    Private Sub TabControl1_DoubleClick(sender As Object, e As System.EventArgs) Handles TabControl1.DoubleClick
        CreateNewTab()
    End Sub

    Private Sub CreateNewTab(Optional ByVal SQL As String = "", Optional ByVal ServerIP As String = "localhost", Optional ByVal DataBaseName As String = "", Optional ByVal UserName As String = "", Optional ByVal Password As String = "", Optional ByVal SelectedDriver As Integer = 1, Optional ByVal RowsToReturn As Integer = 1000, Optional ByVal DriverName As String = "", Optional ByVal ConnectionString As String = "", Optional Copy As Boolean = True)
        'If TabControl1.TabPages.Count = 7 Then Return

        ' copy values from selected tab (if any). don't copy if using parameters set above
        If Copy Then
            Dim f1 As Form = TabControl1.SelectedForm
            If f1 IsNot Nothing Then
                Dim sqlc As sqlControl = f1.Controls(0)

                ServerIP = sqlc.ServerIP
                DataBaseName = sqlc.DatabaseName
                UserName = sqlc.UserName
                Password = sqlc.Password
                SelectedDriver = sqlc.SelectedDriver
                RowsToReturn = sqlc.RowsToReturn
                DriverName = sqlc.DriverName
                ConnectionString = sqlc.ConnectionString
            End If
        End If

        Dim un As New sqlControl
        un.Dock = DockStyle.Fill
        un.ServerIP = ServerIP
        un.DatabaseName = DataBaseName
        un.UserName = UserName
        un.Password = Password
        un.RowsToReturn = RowsToReturn
        un.DriverName = DriverName
        un.ConnectionString = ConnectionString
        un.SQLText = SQL

        un.cmbSelectedDriver.Items.Clear()
        un.cmbSelectedDriver.Items.Add("FileMaker ODBC (" & FMODBCVersion & ")")
        un.cmbSelectedDriver.Items.Add("Other")

        AddHandler un.DatabaseNameChanged, AddressOf OnDatabaseNameChange
        AddHandler un.BeginSQLExecute, AddressOf OnBeginSQLExecute

        _TabCount += 1
        Dim f As New Form
        f.Text = ""
        f.Controls.Add(un)

        TabControl1.TabPages.Add(f)

        un.SelectedDriver = SelectedDriver

    End Sub

    Private Sub OnDatabaseNameChange(ByVal e As textChangedArgs)
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

    Private Sub Form1_SizeChanged(sender As Object, e As System.EventArgs) Handles Me.SizeChanged

        ' fix a bug in tab control where the last tab does not position properly after being minimised
        If _PrevWindowState = FormWindowState.Minimized Then
            TabControl1.ArrangeItems()
        End If

        _PrevWindowState = Me.WindowState
    End Sub
End Class

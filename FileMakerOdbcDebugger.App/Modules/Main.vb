Imports System.Text.RegularExpressions
Imports FastColoredTextBoxNS
Imports FileMakerOdbcDebugger.Util.Json
Imports FileMakerOdbcDebugger.Util.Security
Imports FileMakerOdbcDebugger.Util.Common
Imports FileMakerOdbcDebugger.Util.Sql

Module Main

    Public ReadOnly Path_AppData As String = My.Application.GetEnvironmentVariable("APPDATA") & "\FileMaker ODBC Debugger"
    Public ReadOnly Path_Settings As String = Path_AppData & "\settings.json"
    Public ReadOnly Path_LastQuery As String = Path_AppData & "\last-query.sql"

    Public ReadOnly FmOdbcDriverVersion As String = GetOdbcDriverVersion64Bit()

    Public Const MAX_TABS As Integer = 7 ' todo: there is a bug with the tab controls, and tabs disapear if they extend beyond the form.

    Public ReadOnly StyleColor_Blue As Style = New TextStyle(Brushes.Blue, Nothing, FontStyle.Regular)
    Public ReadOnly StyleColor_Gray As Style = New TextStyle(Brushes.Gray, Nothing, FontStyle.Regular)
    Public ReadOnly StyleColor_GrayStrike As Style = New TextStyle(Brushes.Gray, Nothing, FontStyle.Strikeout)
    Public ReadOnly StyleColor_Green As Style = New TextStyle(Brushes.Green, Nothing, FontStyle.Regular)
    Public ReadOnly StyleColor_Magenta As Style = New TextStyle(Brushes.Magenta, Nothing, FontStyle.Regular)
    Public ReadOnly StyleColor_Maroon As Style = New TextStyle(Brushes.Maroon, Nothing, FontStyle.Regular)
    Public ReadOnly StyleColor_Red As Style = New TextStyle(Brushes.Red, Nothing, FontStyle.Regular)

#Region "Settings"

    Public Class Settings
        Public Query As String = ""
        Public ServerAddress As String = "localhost"
        Public DatabaseName As String = ""
        Public Username As String = ""
        Public Password As String = ""
        Public SelectedDriver As SqlControl.SelectedDriverIndex = 1
        Public RowsToReturn As Integer = 1000
        Public ConnectionString As String = ""
    End Class

    ''' <summary> Save settings to a json file and encrypt. </summary>
    Public Sub SaveSettings(Tab As TabControl)

        Dim Settings As New List(Of Settings)

        For Each t As TabPage In Tab.TabPages

            Dim sqlc As SqlControl = t.Form.Controls(0)

            Dim s As New Settings With {
                .Query = sqlc.SQLText,
                .ServerAddress = sqlc.ServerAddress,
                .DatabaseName = sqlc.DatabaseName,
                .Username = sqlc.Username,
                .Password = sqlc.Password,
                .SelectedDriver = sqlc.SelectedDriver,
                .RowsToReturn = sqlc.RowsToReturn,
                .ConnectionString = sqlc.ConnectionString
            }

            Settings.Add(s)

        Next

        IO.File.WriteAllText(Path_Settings, Settings.ToJson().Encrypt())

    End Sub

    ''' <summary> Load settings from a json file and decrypt. </summary>
    Public Function LoadSettings() As List(Of Settings)

        Dim Settings As New List(Of Settings)

        If Util.IO.PathExists(Path_Settings) Then

            Dim s As String = IO.File.ReadAllText(Path_Settings).Derypt()
            Settings = s.FromJson(GetType(List(Of Settings)))

        End If

        Return Settings

    End Function

    Public Function CreateDivider(Optional Dock As DockStyle = DockStyle.Bottom) As Label
        Dim l = New Label
        l.Height = 1
        l.BackColor = Color.Silver
        l.Dock = Dock
        l.BringToFront()
        Return l
    End Function

#End Region

    Public Sub InitaliseTextBoxSettings(ByVal TextBox As FastColoredTextBox)

        TextBox.CommentPrefix = "--"
        TextBox.LeftBracket = "("c
        TextBox.RightBracket = ")"c
        TextBox.LeftBracket2 = vbNullChar
        TextBox.RightBracket2 = vbNullChar
        TextBox.AutoIndentCharsPatterns = ""
        TextBox.Paddings = New Padding(1, 8, 2, 8)
        TextBox.ReservedCountOfLineNumberChars = 2
        TextBox.AllowMacroRecording = False

    End Sub

    Public Sub SetRangeStyle(ByVal Range As FastColoredTextBoxNS.Range)

        Range.ClearStyle(StyleColor_Green, StyleColor_Red, StyleColor_Magenta, StyleColor_Blue, StyleColor_Maroon, StyleColor_Gray)
        Range.SetStyle(StyleColor_Green, Syntax.Filemaker.Comment1)
        Range.SetStyle(StyleColor_Green, Syntax.Filemaker.Comment2)
        Range.SetStyle(StyleColor_Green, Syntax.Filemaker.Comment3)
        Range.SetStyle(StyleColor_Red, Syntax.Filemaker.String)
        Range.SetStyle(StyleColor_Magenta, Syntax.Filemaker.SpecialKeywords)
        Range.SetStyle(StyleColor_Gray, Syntax.Filemaker.Operators)
        Range.SetStyle(StyleColor_Blue, Syntax.Filemaker.Keywords)
        Range.SetStyle(StyleColor_Magenta, Syntax.Filemaker.Functions)
        Range.SetStyle(StyleColor_Maroon, Syntax.Filemaker.Types)
        Range.SetStyle(StyleColor_GrayStrike, Syntax.Filemaker.ReservedKeywords) ' Do last, so we don't override partial syntax of something that is actually supported.

        Range.ClearFoldingMarkers()
        Range.SetFoldingMarkers("\bBEGIN\b", "\bEND\b", RegexOptions.IgnoreCase)
        Range.SetFoldingMarkers("/\*", "\*/")

    End Sub

End Module

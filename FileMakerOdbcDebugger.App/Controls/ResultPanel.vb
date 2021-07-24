Imports FileMakerOdbcDebugger.Util
Imports FileMakerOdbcDebugger.Util.Common
Imports FileMakerOdbcDebugger.Util.Extensions

Public Class ResultPanel

    Public _DgvBindingSource As New BindingSource

    Private _ColumnRightClicked As Integer = -1
    Private _RowRightClicked As Integer = -1

    Private _Search_Text As String = ""
    Private _Search_CurrentCell As Integer = 0
    Private _Search_FoundCells As New List(Of DataGridViewCell)

    Private Sub ResultsPanel_Load(sender As Object, e As EventArgs) Handles Me.Load

        ' Hack to hide the row selection arrow icon
        ' Refer: https://stackoverflow.com/a/5825458/737393
        dgvResults.RowHeadersDefaultCellStyle.Padding = New Padding(dgvResults.RowHeadersWidth)

        dgvResults.RowHeadersWidth = 35

        dgvResults.DataSource = _DgvBindingSource

        DisableSearchBox()

        dgvResults.EnableDoubleBuffering

        lblDurationStream.Visible = False

    End Sub

    Public Sub UI_DisplayData(Result As Sql.StatementResult)

        If Result.Data.Count > 0 Then

            Dim Table = New DataTable

            ' Prepare unique column names:
            ' We can't add duplicate columns to a DataSet, so instead we incrementally append
            ' a special zero-width space on the end of the column name until its unique.
            Dim ZeroWidthSpace = ChrW(&H200B) '"\u200B"
            Dim UniqueColumnNames As New List(Of String)
            For Each d In Result.Data(0)

                Dim ColumnName = d
                While UniqueColumnNames.Contains(ColumnName)
                    ColumnName &= ZeroWidthSpace
                End While

            Next

            ' Add columns:
            For Each ColumnName In UniqueColumnNames

                Dim NewColumn = New DataColumn With {
                    .ColumnName = ColumnName,
                    .Caption = ColumnName.Replace(ZeroWidthSpace, "")
                }
                Table.Columns.Add(ColumnName)

            Next

            ' Add rows:
            If Result.Data.Count > 1 Then

                Dim row As DataRow

                For i = 1 To Result.Data.Count - 1
                    row = Table.NewRow

                    Dim s(Result.Data(1).Count - 1) As Object  ' initalise object with same number of rows

                    For j = 0 To Result.Data(i).Count - 1
                        row.Item(j) = Result.Data(i)(j)
                    Next

                    Table.Rows.Add(row)

                Next

                _DgvBindingSource.DataSource = Table

            End If

            If dgvResults.Rows.Count > 0 Then

                If txtFind.Text.Count < 2 Then ' select first cell if we are not already searching for something
                    dgvResults.Rows(0).Cells(0).Selected = True
                End If

            End If

        End If

        lblDurationStream.Text = "Stream: " & Result.Duration_Stream.ToDisplayString
        lblDurationStream.Visible = True

        lblDurationExecute.Text = "Execute: " & Result.Duration_Execute.ToDisplayString
        lblDurationExecute.Visible = True

    End Sub

    Public Sub UI_RenderQueryWarnings(ByVal Warnings As List(Of Sql.QueryIssue))

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

    Public Sub UI_RenderRowCount(ResultRowCount As Integer, ResultRowsAffected As Integer, RowLimit As Integer)

        If ResultRowCount <= 0 And ResultRowsAffected > -1 Then

            ' Show number of records affected:
            If ResultRowsAffected = 1 Then
                lblStatus.Text = ResultRowsAffected & " row affected"
            Else
                lblStatus.Text = ResultRowsAffected & " rows affected"
            End If

        Else

            ' Show a "max" label to remind that the rows have been limited.
            Dim MaxLabel As String = ""
            If ResultRowCount = RowLimit Then
                MaxLabel = " (max)"
            End If

            ' Show row count.
            lblStatus.Text = $"{ResultRowCount} rows{MaxLabel}"

        End If

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

    Private Sub dgvResults_MouseDown(sender As Object, e As System.Windows.Forms.MouseEventArgs) Handles dgvResults.MouseDown

        Dim s As DataGridView = sender
        _ColumnRightClicked = s.HitTest(e.Location.X, e.Location.Y).ColumnIndex
        _RowRightClicked = s.HitTest(e.Location.X, e.Location.Y).RowIndex

    End Sub

    Private Sub dgvResults_DataBindingComplete(ByVal sender As Object, ByVal e As DataGridViewBindingCompleteEventArgs) Handles dgvResults.DataBindingComplete
        ' Re-index the search results after new data is loaded.
        Search(0)
    End Sub

    Private Sub mnuCopyAsCSV_Click(sender As System.Object, e As System.EventArgs) Handles mnuCopyAsCSV.Click
        Try
            Dim s = ""
            For Each r As DataGridViewRow In dgvResults.Rows
                Dim value = r.Cells(_ColumnRightClicked).Value
                If Not IsDBNull(value) Then
                    s &= value.ToString & ","
                End If
            Next
            s = s.TrimNCharsFromEnd(1)
            Clipboard.SetText(s)
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

        ExportToCsv(Data)

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
                        lblFindCount.BackColor = Color.LightBlue
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

End Class

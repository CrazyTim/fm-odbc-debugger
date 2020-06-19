<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class ResultPanel
    Inherits System.Windows.Forms.UserControl

    'UserControl overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim DataGridViewCellStyle2 As System.Windows.Forms.DataGridViewCellStyle = New System.Windows.Forms.DataGridViewCellStyle()
        Me.dgvResults = New System.Windows.Forms.DataGridView()
        Me.ContextMenuStrip2 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.mnuCopy = New System.Windows.Forms.ToolStripMenuItem()
        Me.mnuCopyWithHeaders = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator2 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuCopyAsCSV = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripSeparator7 = New System.Windows.Forms.ToolStripSeparator()
        Me.mnuExportToExcel = New System.Windows.Forms.ToolStripMenuItem()
        Me.Panel_Status = New System.Windows.Forms.Panel()
        Me.lblWarning = New System.Windows.Forms.Label()
        Me.lblDurationStream = New System.Windows.Forms.Label()
        Me.lblStatus = New System.Windows.Forms.Label()
        Me.Label12 = New System.Windows.Forms.Label()
        Me.Panel_Search = New System.Windows.Forms.Panel()
        Me.Label13 = New System.Windows.Forms.Label()
        Me.txtFind = New System.Windows.Forms.TextBox()
        Me.Label11 = New System.Windows.Forms.Label()
        Me.lblFindCount = New System.Windows.Forms.Label()
        Me.ToolTip1 = New System.Windows.Forms.ToolTip(Me.components)
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblDurationExecute = New System.Windows.Forms.Label()
        CType(Me.dgvResults, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.ContextMenuStrip2.SuspendLayout()
        Me.Panel_Status.SuspendLayout()
        Me.Panel_Search.SuspendLayout()
        Me.SuspendLayout()
        '
        'dgvResults
        '
        Me.dgvResults.AllowUserToAddRows = False
        Me.dgvResults.AllowUserToDeleteRows = False
        Me.dgvResults.AllowUserToOrderColumns = True
        Me.dgvResults.AllowUserToResizeRows = False
        Me.dgvResults.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.dgvResults.BackgroundColor = System.Drawing.SystemColors.Control
        Me.dgvResults.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.dgvResults.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.DisableResizing
        Me.dgvResults.ContextMenuStrip = Me.ContextMenuStrip2
        DataGridViewCellStyle2.Alignment = System.Windows.Forms.DataGridViewContentAlignment.MiddleLeft
        DataGridViewCellStyle2.BackColor = System.Drawing.SystemColors.Window
        DataGridViewCellStyle2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        DataGridViewCellStyle2.ForeColor = System.Drawing.SystemColors.ControlText
        DataGridViewCellStyle2.Padding = New System.Windows.Forms.Padding(0, 0, 5, 0)
        DataGridViewCellStyle2.SelectionBackColor = System.Drawing.SystemColors.Highlight
        DataGridViewCellStyle2.SelectionForeColor = System.Drawing.SystemColors.HighlightText
        DataGridViewCellStyle2.WrapMode = System.Windows.Forms.DataGridViewTriState.[False]
        Me.dgvResults.DefaultCellStyle = DataGridViewCellStyle2
        Me.dgvResults.Location = New System.Drawing.Point(0, 23)
        Me.dgvResults.Name = "dgvResults"
        Me.dgvResults.ReadOnly = True
        Me.dgvResults.RowHeadersWidth = 25
        Me.dgvResults.RowTemplate.DefaultCellStyle.Font = New System.Drawing.Font("Consolas", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.dgvResults.Size = New System.Drawing.Size(600, 200)
        Me.dgvResults.TabIndex = 1
        '
        'ContextMenuStrip2
        '
        Me.ContextMenuStrip2.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.mnuCopy, Me.mnuCopyWithHeaders, Me.ToolStripSeparator2, Me.mnuCopyAsCSV, Me.ToolStripSeparator7, Me.mnuExportToExcel})
        Me.ContextMenuStrip2.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip2.Size = New System.Drawing.Size(252, 120)
        '
        'mnuCopy
        '
        Me.mnuCopy.Name = "mnuCopy"
        Me.mnuCopy.ShortcutKeys = CType((System.Windows.Forms.Keys.Control Or System.Windows.Forms.Keys.C), System.Windows.Forms.Keys)
        Me.mnuCopy.Size = New System.Drawing.Size(251, 22)
        Me.mnuCopy.Text = "Copy Selected Cells"
        '
        'mnuCopyWithHeaders
        '
        Me.mnuCopyWithHeaders.Name = "mnuCopyWithHeaders"
        Me.mnuCopyWithHeaders.Size = New System.Drawing.Size(251, 22)
        Me.mnuCopyWithHeaders.Text = "Copy Selected Cells With Headers"
        '
        'ToolStripSeparator2
        '
        Me.ToolStripSeparator2.AutoSize = False
        Me.ToolStripSeparator2.Name = "ToolStripSeparator2"
        Me.ToolStripSeparator2.Size = New System.Drawing.Size(248, 14)
        '
        'mnuCopyAsCSV
        '
        Me.mnuCopyAsCSV.Name = "mnuCopyAsCSV"
        Me.mnuCopyAsCSV.Size = New System.Drawing.Size(251, 22)
        Me.mnuCopyAsCSV.Text = "Copy Selected Column(s) As CSV"
        '
        'ToolStripSeparator7
        '
        Me.ToolStripSeparator7.AutoSize = False
        Me.ToolStripSeparator7.Name = "ToolStripSeparator7"
        Me.ToolStripSeparator7.Size = New System.Drawing.Size(248, 14)
        '
        'mnuExportToExcel
        '
        Me.mnuExportToExcel.Name = "mnuExportToExcel"
        Me.mnuExportToExcel.Size = New System.Drawing.Size(251, 22)
        Me.mnuExportToExcel.Text = "Export Results To .Csv File"
        Me.mnuExportToExcel.ToolTipText = "Export results to a .csv file and open using the default program"
        '
        'Panel_Status
        '
        Me.Panel_Status.Controls.Add(Me.lblWarning)
        Me.Panel_Status.Controls.Add(Me.lblDurationStream)
        Me.Panel_Status.Controls.Add(Me.lblDurationExecute)
        Me.Panel_Status.Controls.Add(Me.lblStatus)
        Me.Panel_Status.Controls.Add(Me.Label12)
        Me.Panel_Status.Controls.Add(Me.Label1)
        Me.Panel_Status.Controls.Add(Me.Panel_Search)
        Me.Panel_Status.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel_Status.Location = New System.Drawing.Point(0, 0)
        Me.Panel_Status.Name = "Panel_Status"
        Me.Panel_Status.Size = New System.Drawing.Size(600, 25)
        Me.Panel_Status.TabIndex = 24
        '
        'lblWarning
        '
        Me.lblWarning.AutoSize = True
        Me.lblWarning.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblWarning.Image = Global.FileMakerOdbcDebugger.App.My.Resources.Resources.alert_16x16
        Me.lblWarning.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblWarning.Location = New System.Drawing.Point(310, 1)
        Me.lblWarning.Name = "lblWarning"
        Me.lblWarning.Padding = New System.Windows.Forms.Padding(4, 4, 15, 0)
        Me.lblWarning.Size = New System.Drawing.Size(92, 19)
        Me.lblWarning.TabIndex = 26
        Me.lblWarning.Text = "       Warning"
        '
        'lblDurationStream
        '
        Me.lblDurationStream.AutoSize = True
        Me.lblDurationStream.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblDurationStream.Location = New System.Drawing.Point(190, 1)
        Me.lblDurationStream.Name = "lblDurationStream"
        Me.lblDurationStream.Padding = New System.Windows.Forms.Padding(4, 4, 15, 0)
        Me.lblDurationStream.Size = New System.Drawing.Size(120, 19)
        Me.lblDurationStream.TabIndex = 31
        Me.lblDurationStream.Text = "{duration-stream}"
        '
        'lblStatus
        '
        Me.lblStatus.AutoSize = True
        Me.lblStatus.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblStatus.Location = New System.Drawing.Point(0, 1)
        Me.lblStatus.Name = "lblStatus"
        Me.lblStatus.Padding = New System.Windows.Forms.Padding(4, 4, 15, 0)
        Me.lblStatus.Size = New System.Drawing.Size(65, 19)
        Me.lblStatus.TabIndex = 24
        Me.lblStatus.Text = "{status}"
        '
        'Label12
        '
        Me.Label12.BackColor = System.Drawing.Color.Silver
        Me.Label12.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.Label12.Font = New System.Drawing.Font("Consolas", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label12.Location = New System.Drawing.Point(0, 24)
        Me.Label12.Name = "Label12"
        Me.Label12.Size = New System.Drawing.Size(600, 1)
        Me.Label12.TabIndex = 29
        Me.Label12.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'Panel_Search
        '
        Me.Panel_Search.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.Panel_Search.BackColor = System.Drawing.SystemColors.Window
        Me.Panel_Search.Controls.Add(Me.Label13)
        Me.Panel_Search.Controls.Add(Me.Label11)
        Me.Panel_Search.Controls.Add(Me.lblFindCount)
        Me.Panel_Search.Controls.Add(Me.txtFind)
        Me.Panel_Search.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Panel_Search.Location = New System.Drawing.Point(419, -1)
        Me.Panel_Search.Name = "Panel_Search"
        Me.Panel_Search.Size = New System.Drawing.Size(182, 25)
        Me.Panel_Search.TabIndex = 27
        '
        'Label13
        '
        Me.Label13.BackColor = System.Drawing.Color.Silver
        Me.Label13.Dock = System.Windows.Forms.DockStyle.Left
        Me.Label13.Font = New System.Drawing.Font("Consolas", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label13.Location = New System.Drawing.Point(0, 0)
        Me.Label13.Name = "Label13"
        Me.Label13.Size = New System.Drawing.Size(1, 25)
        Me.Label13.TabIndex = 32
        Me.Label13.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'txtFind
        '
        Me.txtFind.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtFind.BorderStyle = System.Windows.Forms.BorderStyle.None
        Me.txtFind.Font = New System.Drawing.Font("Consolas", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtFind.Location = New System.Drawing.Point(24, 5)
        Me.txtFind.MaxLength = 25
        Me.txtFind.Name = "txtFind"
        Me.txtFind.Size = New System.Drawing.Size(133, 15)
        Me.txtFind.TabIndex = 22
        Me.txtFind.TabStop = False
        '
        'Label11
        '
        Me.Label11.Cursor = System.Windows.Forms.Cursors.IBeam
        Me.Label11.Image = Global.FileMakerOdbcDebugger.App.My.Resources.Resources.search_16x16
        Me.Label11.Location = New System.Drawing.Point(6, 2)
        Me.Label11.Name = "Label11"
        Me.Label11.Padding = New System.Windows.Forms.Padding(0, 1, 7, 2)
        Me.Label11.Size = New System.Drawing.Size(18, 21)
        Me.Label11.TabIndex = 23
        '
        'lblFindCount
        '
        Me.lblFindCount.AutoSize = True
        Me.lblFindCount.BackColor = System.Drawing.Color.LightSalmon
        Me.lblFindCount.Cursor = System.Windows.Forms.Cursors.Default
        Me.lblFindCount.Dock = System.Windows.Forms.DockStyle.Right
        Me.lblFindCount.Location = New System.Drawing.Point(164, 0)
        Me.lblFindCount.Name = "lblFindCount"
        Me.lblFindCount.Padding = New System.Windows.Forms.Padding(2, 5, 3, 5)
        Me.lblFindCount.Size = New System.Drawing.Size(18, 25)
        Me.lblFindCount.TabIndex = 24
        Me.lblFindCount.Text = "5"
        Me.lblFindCount.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        Me.lblFindCount.Visible = False
        '
        'ToolTip1
        '
        Me.ToolTip1.AutoPopDelay = 32000
        Me.ToolTip1.InitialDelay = 100
        Me.ToolTip1.ReshowDelay = 100
        '
        'Label1
        '
        Me.Label1.BackColor = System.Drawing.Color.Silver
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label1.Font = New System.Drawing.Font("Consolas", 9.75!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(0, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(600, 1)
        Me.Label1.TabIndex = 32
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleLeft
        '
        'lblDurationExecute
        '
        Me.lblDurationExecute.AutoSize = True
        Me.lblDurationExecute.Dock = System.Windows.Forms.DockStyle.Left
        Me.lblDurationExecute.Location = New System.Drawing.Point(65, 1)
        Me.lblDurationExecute.Name = "lblDurationExecute"
        Me.lblDurationExecute.Padding = New System.Windows.Forms.Padding(4, 4, 15, 0)
        Me.lblDurationExecute.Size = New System.Drawing.Size(125, 19)
        Me.lblDurationExecute.TabIndex = 33
        Me.lblDurationExecute.Text = "{duration-execute}"
        '
        'ResultPanel
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(7.0!, 15.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Panel_Status)
        Me.Controls.Add(Me.dgvResults)
        Me.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Name = "ResultPanel"
        Me.Size = New System.Drawing.Size(600, 223)
        CType(Me.dgvResults, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ContextMenuStrip2.ResumeLayout(False)
        Me.Panel_Status.ResumeLayout(False)
        Me.Panel_Status.PerformLayout()
        Me.Panel_Search.ResumeLayout(False)
        Me.Panel_Search.PerformLayout()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents dgvResults As DataGridView
    Friend WithEvents ContextMenuStrip2 As ContextMenuStrip
    Friend WithEvents mnuCopy As ToolStripMenuItem
    Friend WithEvents mnuCopyWithHeaders As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator2 As ToolStripSeparator
    Friend WithEvents mnuCopyAsCSV As ToolStripMenuItem
    Friend WithEvents ToolStripSeparator7 As ToolStripSeparator
    Friend WithEvents mnuExportToExcel As ToolStripMenuItem
    Friend WithEvents Panel_Status As Panel
    Private WithEvents lblWarning As Label
    Private WithEvents lblDurationStream As Label
    Private WithEvents lblStatus As Label
    Private WithEvents Label12 As Label
    Friend WithEvents Panel_Search As Panel
    Private WithEvents Label13 As Label
    Private WithEvents txtFind As TextBox
    Private WithEvents Label11 As Label
    Private WithEvents lblFindCount As Label
    Friend WithEvents ToolTip1 As ToolTip
    Private WithEvents Label1 As Label
    Private WithEvents lblDurationExecute As Label
End Class

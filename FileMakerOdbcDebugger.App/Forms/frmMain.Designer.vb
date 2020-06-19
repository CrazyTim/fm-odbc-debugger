<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmMain
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMain))
        Me.ColumnHeader38 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.ColumnHeader1 = CType(New System.Windows.Forms.ColumnHeader(), System.Windows.Forms.ColumnHeader)
        Me.TabControl1 = New FileMakerOdbcDebugger.App.TabControl()
        Me.SuspendLayout()
        '
        'ColumnHeader38
        '
        Me.ColumnHeader38.Text = "-"
        Me.ColumnHeader38.Width = 22
        '
        'ColumnHeader1
        '
        Me.ColumnHeader1.Text = "-"
        Me.ColumnHeader1.Width = 22
        '
        'TabControl1
        '
        Me.TabControl1.BackHighColor = System.Drawing.Color.Silver
        Me.TabControl1.BackLowColor = System.Drawing.Color.Silver
        Me.TabControl1.BorderColorDisabled = System.Drawing.SystemColors.ControlDarkDark
        Me.TabControl1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.TabControl1.DropButtonVisible = False
        Me.TabControl1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.TabControl1.Location = New System.Drawing.Point(0, 0)
        Me.TabControl1.MenuRenderer = Nothing
        Me.TabControl1.Name = "TabControl1"
        Me.TabControl1.Size = New System.Drawing.Size(804, 511)
        Me.TabControl1.TabBackHighColor = System.Drawing.SystemColors.Control
        Me.TabControl1.TabBackHighColorDisabled = System.Drawing.Color.LightGray
        Me.TabControl1.TabBackLowColorDisabled = System.Drawing.Color.LightGray
        Me.TabControl1.TabBorderEnhanceWeight = FileMakerOdbcDebugger.App.TabControl.Weight.Soft
        Me.TabControl1.TabCloseButtonBackHighColorDisabled = System.Drawing.SystemColors.ControlLightLight
        Me.TabControl1.TabCloseButtonBackLowColorDisabled = System.Drawing.SystemColors.ControlLightLight
        Me.TabControl1.TabCloseButtonImage = Nothing
        Me.TabControl1.TabCloseButtonImageDisabled = Nothing
        Me.TabControl1.TabCloseButtonImageHot = Nothing
        Me.TabControl1.TabHeight = 22
        Me.TabControl1.TabIconSize = New System.Drawing.Size(0, 0)
        Me.TabControl1.TabIndex = 3
        Me.TabControl1.TopSeparator = False
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(804, 511)
        Me.Controls.Add(Me.TabControl1)
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.KeyPreview = True
        Me.MinimumSize = New System.Drawing.Size(820, 400)
        Me.Name = "Form1"
        Me.Text = "FileMaker ODBC Debugger"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents ColumnHeader38 As System.Windows.Forms.ColumnHeader
    Friend WithEvents ColumnHeader1 As System.Windows.Forms.ColumnHeader
    Friend WithEvents TabControl1 As FileMakerOdbcDebugger.App.TabControl

End Class

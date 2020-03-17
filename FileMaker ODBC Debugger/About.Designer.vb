<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class About
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(About))
        Me.Panel2 = New System.Windows.Forms.Panel()
        Me.Label6 = New System.Windows.Forms.Label()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.lblAppTitle = New System.Windows.Forms.Label()
        Me.FlowLayoutPanel1 = New System.Windows.Forms.FlowLayoutPanel()
        Me.lblDescription = New System.Windows.Forms.Label()
        Me.lblNotes = New System.Windows.Forms.Label()
        Me.lblFeatures = New System.Windows.Forms.Label()
        Me.lblTodo = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Panel2.SuspendLayout()
        Me.FlowLayoutPanel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.AutoSize = True
        Me.Panel2.BackColor = System.Drawing.Color.White
        Me.Panel2.Controls.Add(Me.Label6)
        Me.Panel2.Controls.Add(Me.Label1)
        Me.Panel2.Controls.Add(Me.lblAppTitle)
        Me.Panel2.Dock = System.Windows.Forms.DockStyle.Top
        Me.Panel2.Location = New System.Drawing.Point(0, 0)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Padding = New System.Windows.Forms.Padding(9, 9, 9, 10)
        Me.Panel2.Size = New System.Drawing.Size(634, 101)
        Me.Panel2.TabIndex = 167
        '
        'Label6
        '
        Me.Label6.AutoSize = True
        Me.Label6.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label6.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label6.Location = New System.Drawing.Point(9, 71)
        Me.Label6.Name = "Label6"
        Me.Label6.Padding = New System.Windows.Forms.Padding(0, 0, 0, 5)
        Me.Label6.Size = New System.Drawing.Size(191, 20)
        Me.Label6.TabIndex = 166
        Me.Label6.Text = "Copyright © 2020 Jason Ballantyne"
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Dock = System.Windows.Forms.DockStyle.Top
        Me.Label1.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label1.Location = New System.Drawing.Point(9, 51)
        Me.Label1.Name = "Label1"
        Me.Label1.Padding = New System.Windows.Forms.Padding(0, 0, 0, 5)
        Me.Label1.Size = New System.Drawing.Size(109, 20)
        Me.Label1.TabIndex = 164
        Me.Label1.Text = "Version: 2020-03-17"
        '
        'lblAppTitle
        '
        Me.lblAppTitle.AutoSize = True
        Me.lblAppTitle.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblAppTitle.Font = New System.Drawing.Font("Segoe UI Semibold", 18.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblAppTitle.Location = New System.Drawing.Point(9, 9)
        Me.lblAppTitle.Name = "lblAppTitle"
        Me.lblAppTitle.Padding = New System.Windows.Forms.Padding(0, 0, 0, 10)
        Me.lblAppTitle.Size = New System.Drawing.Size(309, 42)
        Me.lblAppTitle.TabIndex = 165
        Me.lblAppTitle.Text = "FileMaker ODBC Debugger"
        '
        'FlowLayoutPanel1
        '
        Me.FlowLayoutPanel1.AutoScroll = True
        Me.FlowLayoutPanel1.Controls.Add(Me.lblDescription)
        Me.FlowLayoutPanel1.Controls.Add(Me.lblNotes)
        Me.FlowLayoutPanel1.Controls.Add(Me.lblFeatures)
        Me.FlowLayoutPanel1.Controls.Add(Me.lblTodo)
        Me.FlowLayoutPanel1.Controls.Add(Me.Label2)
        Me.FlowLayoutPanel1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.FlowLayoutPanel1.Location = New System.Drawing.Point(0, 101)
        Me.FlowLayoutPanel1.Name = "FlowLayoutPanel1"
        Me.FlowLayoutPanel1.Size = New System.Drawing.Size(634, 294)
        Me.FlowLayoutPanel1.TabIndex = 169
        '
        'lblDescription
        '
        Me.lblDescription.AutoSize = True
        Me.lblDescription.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblDescription.Location = New System.Drawing.Point(3, 0)
        Me.lblDescription.Name = "lblDescription"
        Me.lblDescription.Padding = New System.Windows.Forms.Padding(9, 2, 9, 11)
        Me.lblDescription.Size = New System.Drawing.Size(608, 103)
        Me.lblDescription.TabIndex = 173
        Me.lblDescription.Text = resources.GetString("lblDescription.Text")
        '
        'lblNotes
        '
        Me.lblNotes.AutoSize = True
        Me.lblNotes.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblNotes.Location = New System.Drawing.Point(3, 103)
        Me.lblNotes.Name = "lblNotes"
        Me.lblNotes.Padding = New System.Windows.Forms.Padding(9, 2, 9, 11)
        Me.lblNotes.Size = New System.Drawing.Size(513, 73)
        Me.lblNotes.TabIndex = 174
        Me.lblNotes.Text = resources.GetString("lblNotes.Text")
        '
        'lblFeatures
        '
        Me.lblFeatures.AutoSize = True
        Me.lblFeatures.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblFeatures.Location = New System.Drawing.Point(3, 176)
        Me.lblFeatures.Name = "lblFeatures"
        Me.lblFeatures.Padding = New System.Windows.Forms.Padding(9, 2, 9, 11)
        Me.lblFeatures.Size = New System.Drawing.Size(610, 268)
        Me.lblFeatures.TabIndex = 175
        Me.lblFeatures.Text = resources.GetString("lblFeatures.Text")
        '
        'lblTodo
        '
        Me.lblTodo.AutoSize = True
        Me.lblTodo.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblTodo.Location = New System.Drawing.Point(3, 444)
        Me.lblTodo.Name = "lblTodo"
        Me.lblTodo.Padding = New System.Windows.Forms.Padding(9, 2, 9, 11)
        Me.lblTodo.Size = New System.Drawing.Size(543, 118)
        Me.lblTodo.TabIndex = 176
        Me.lblTodo.Text = resources.GetString("lblTodo.Text")
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Font = New System.Drawing.Font("Segoe UI", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.Label2.Location = New System.Drawing.Point(3, 562)
        Me.Label2.Name = "Label2"
        Me.Label2.Padding = New System.Windows.Forms.Padding(9, 2, 9, 11)
        Me.Label2.Size = New System.Drawing.Size(515, 58)
        Me.Label2.TabIndex = 177
        Me.Label2.Text = "ACKNOWLEDGEMENTS:" & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "- Icon by David Vignoni (www.iconfinder.com/icons/1230/animal_" &
    "bug_insect_ladybird_icon)." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "- Csv code from joshclose.github.io/CsvHelper, Apach" &
    "e License, Version 2.0."
        '
        'About
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(634, 395)
        Me.Controls.Add(Me.FlowLayoutPanel1)
        Me.Controls.Add(Me.Panel2)
        Me.Font = New System.Drawing.Font("Segoe UI", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "About"
        Me.ShowIcon = False
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent
        Me.Text = "About"
        Me.Panel2.ResumeLayout(False)
        Me.Panel2.PerformLayout()
        Me.FlowLayoutPanel1.ResumeLayout(False)
        Me.FlowLayoutPanel1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents lblAppTitle As System.Windows.Forms.Label
    Friend WithEvents Label6 As Label
    Friend WithEvents FlowLayoutPanel1 As FlowLayoutPanel
    Friend WithEvents lblDescription As Label
    Friend WithEvents lblNotes As Label
    Friend WithEvents lblFeatures As Label
    Friend WithEvents lblTodo As Label
    Friend WithEvents Label2 As Label
End Class

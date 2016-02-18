<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form3
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
        Me.GroupVerbose = New System.Windows.Forms.GroupBox
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.Group1 = New System.Windows.Forms.GroupBox
        Me.Group2 = New System.Windows.Forms.GroupBox
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.ListTable = New System.Windows.Forms.ListBox
        Me.ButtonPurge = New System.Windows.Forms.Button
        Me.Button1 = New System.Windows.Forms.Button
        Me.GroupVerbose.SuspendLayout()
        Me.Group1.SuspendLayout()
        Me.Group2.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupVerbose
        '
        Me.GroupVerbose.Controls.Add(Me.ListBox1)
        Me.GroupVerbose.Location = New System.Drawing.Point(12, 177)
        Me.GroupVerbose.Name = "GroupVerbose"
        Me.GroupVerbose.Size = New System.Drawing.Size(468, 214)
        Me.GroupVerbose.TabIndex = 2
        Me.GroupVerbose.TabStop = False
        Me.GroupVerbose.Text = "Verbose"
        '
        'ListBox1
        '
        Me.ListBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(3, 16)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(462, 186)
        Me.ListBox1.TabIndex = 2
        Me.ListBox1.TabStop = False
        '
        'Group1
        '
        Me.Group1.Controls.Add(Me.ListTable)
        Me.Group1.Location = New System.Drawing.Point(12, 12)
        Me.Group1.Name = "Group1"
        Me.Group1.Size = New System.Drawing.Size(200, 117)
        Me.Group1.TabIndex = 3
        Me.Group1.TabStop = False
        Me.Group1.Text = "Tables"
        '
        'Group2
        '
        Me.Group2.Controls.Add(Me.DataGridView1)
        Me.Group2.Location = New System.Drawing.Point(243, 45)
        Me.Group2.Name = "Group2"
        Me.Group2.Size = New System.Drawing.Size(200, 100)
        Me.Group2.TabIndex = 4
        Me.Group2.TabStop = False
        Me.Group2.Text = "Indexes"
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(3, 16)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(194, 81)
        Me.DataGridView1.TabIndex = 2
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 381)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(492, 22)
        Me.StatusStrip1.TabIndex = 5
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'ListTable
        '
        Me.ListTable.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle
        Me.ListTable.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListTable.FormattingEnabled = True
        Me.ListTable.Location = New System.Drawing.Point(3, 16)
        Me.ListTable.Name = "ListTable"
        Me.ListTable.Size = New System.Drawing.Size(194, 93)
        Me.ListTable.TabIndex = 3
        '
        'ButtonPurge
        '
        Me.ButtonPurge.Location = New System.Drawing.Point(246, 151)
        Me.ButtonPurge.Name = "ButtonPurge"
        Me.ButtonPurge.Size = New System.Drawing.Size(194, 23)
        Me.ButtonPurge.TabIndex = 6
        Me.ButtonPurge.Text = "Shrink Log File"
        Me.ButtonPurge.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(12, 135)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(194, 23)
        Me.Button1.TabIndex = 7
        Me.Button1.Text = "Rebuild Indexes"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'Form3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(492, 403)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.ButtonPurge)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.Group2)
        Me.Controls.Add(Me.Group1)
        Me.Controls.Add(Me.GroupVerbose)
        Me.Name = "Form3"
        Me.Text = "Form3"
        Me.GroupVerbose.ResumeLayout(False)
        Me.Group1.ResumeLayout(False)
        Me.Group2.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupVerbose As System.Windows.Forms.GroupBox
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents Group1 As System.Windows.Forms.GroupBox
    Friend WithEvents Group2 As System.Windows.Forms.GroupBox
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents ListTable As System.Windows.Forms.ListBox
    Friend WithEvents ButtonPurge As System.Windows.Forms.Button
    Friend WithEvents Button1 As System.Windows.Forms.Button

End Class

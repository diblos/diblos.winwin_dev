<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmWatcher
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
        Me.GroupData = New System.Windows.Forms.GroupBox
        Me.GroupVerbose = New System.Windows.Forms.GroupBox
        Me.ListBox1 = New System.Windows.Forms.ListBox
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.DataGridView1 = New System.Windows.Forms.DataGridView
        Me.GroupData.SuspendLayout()
        Me.GroupVerbose.SuspendLayout()
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'GroupData
        '
        Me.GroupData.Controls.Add(Me.DataGridView1)
        Me.GroupData.Location = New System.Drawing.Point(234, 54)
        Me.GroupData.Name = "GroupData"
        Me.GroupData.Size = New System.Drawing.Size(200, 100)
        Me.GroupData.TabIndex = 3
        Me.GroupData.TabStop = False
        Me.GroupData.Text = "Files"
        '
        'GroupVerbose
        '
        Me.GroupVerbose.Controls.Add(Me.ListBox1)
        Me.GroupVerbose.Location = New System.Drawing.Point(12, 199)
        Me.GroupVerbose.Name = "GroupVerbose"
        Me.GroupVerbose.Size = New System.Drawing.Size(422, 100)
        Me.GroupVerbose.TabIndex = 4
        Me.GroupVerbose.TabStop = False
        Me.GroupVerbose.Text = "Verbose"
        '
        'ListBox1
        '
        Me.ListBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(3, 16)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(416, 69)
        Me.ListBox1.TabIndex = 3
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 351)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(460, 22)
        Me.StatusStrip1.TabIndex = 5
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'DataGridView1
        '
        Me.DataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.DataGridView1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.DataGridView1.Location = New System.Drawing.Point(3, 16)
        Me.DataGridView1.Name = "DataGridView1"
        Me.DataGridView1.Size = New System.Drawing.Size(194, 81)
        Me.DataGridView1.TabIndex = 0
        '
        'frmWatcher
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(460, 373)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.GroupVerbose)
        Me.Controls.Add(Me.GroupData)
        Me.Name = "frmWatcher"
        Me.Text = "frmWatcher"
        Me.GroupData.ResumeLayout(False)
        Me.GroupVerbose.ResumeLayout(False)
        CType(Me.DataGridView1, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupData As System.Windows.Forms.GroupBox
    Friend WithEvents GroupVerbose As System.Windows.Forms.GroupBox
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents DataGridView1 As System.Windows.Forms.DataGridView
End Class

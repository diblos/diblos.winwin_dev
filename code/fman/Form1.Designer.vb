<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
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
        Me.StatusStrip1 = New System.Windows.Forms.StatusStrip
        Me.txtPath = New System.Windows.Forms.TextBox
        Me.txtPattern = New System.Windows.Forms.TextBox
        Me.GroupVerbose.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupVerbose
        '
        Me.GroupVerbose.Controls.Add(Me.ListBox1)
        Me.GroupVerbose.Location = New System.Drawing.Point(0, 64)
        Me.GroupVerbose.Name = "GroupVerbose"
        Me.GroupVerbose.Size = New System.Drawing.Size(422, 184)
        Me.GroupVerbose.TabIndex = 5
        Me.GroupVerbose.TabStop = False
        Me.GroupVerbose.Text = "Verbose"
        '
        'ListBox1
        '
        Me.ListBox1.Dock = System.Windows.Forms.DockStyle.Fill
        Me.ListBox1.FormattingEnabled = True
        Me.ListBox1.Location = New System.Drawing.Point(3, 16)
        Me.ListBox1.Name = "ListBox1"
        Me.ListBox1.Size = New System.Drawing.Size(416, 160)
        Me.ListBox1.TabIndex = 3
        '
        'StatusStrip1
        '
        Me.StatusStrip1.Location = New System.Drawing.Point(0, 251)
        Me.StatusStrip1.Name = "StatusStrip1"
        Me.StatusStrip1.Size = New System.Drawing.Size(422, 22)
        Me.StatusStrip1.TabIndex = 6
        Me.StatusStrip1.Text = "StatusStrip1"
        '
        'txtPath
        '
        Me.txtPath.Location = New System.Drawing.Point(0, 12)
        Me.txtPath.Name = "txtPath"
        Me.txtPath.Size = New System.Drawing.Size(422, 20)
        Me.txtPath.TabIndex = 7
        '
        'txtPattern
        '
        Me.txtPattern.Location = New System.Drawing.Point(0, 38)
        Me.txtPattern.Name = "txtPattern"
        Me.txtPattern.Size = New System.Drawing.Size(422, 20)
        Me.txtPattern.TabIndex = 8
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(422, 273)
        Me.Controls.Add(Me.txtPattern)
        Me.Controls.Add(Me.txtPath)
        Me.Controls.Add(Me.StatusStrip1)
        Me.Controls.Add(Me.GroupVerbose)
        Me.Name = "Form1"
        Me.Text = "Form1"
        Me.GroupVerbose.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents GroupVerbose As System.Windows.Forms.GroupBox
    Friend WithEvents ListBox1 As System.Windows.Forms.ListBox
    Friend WithEvents StatusStrip1 As System.Windows.Forms.StatusStrip
    Friend WithEvents txtPath As System.Windows.Forms.TextBox
    Friend WithEvents txtPattern As System.Windows.Forms.TextBox

End Class

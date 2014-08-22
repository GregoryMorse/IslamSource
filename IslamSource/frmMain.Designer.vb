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
        Me.tvwMain = New System.Windows.Forms.TreeView()
        Me.gbMain = New System.Windows.Forms.GroupBox()
        Me.SuspendLayout()
        '
        'tvwMain
        '
        Me.tvwMain.Location = New System.Drawing.Point(0, 6)
        Me.tvwMain.Name = "tvwMain"
        Me.tvwMain.Size = New System.Drawing.Size(161, 409)
        Me.tvwMain.TabIndex = 0
        '
        'gbMain
        '
        Me.gbMain.Location = New System.Drawing.Point(167, 6)
        Me.gbMain.Name = "gbMain"
        Me.gbMain.Size = New System.Drawing.Size(386, 409)
        Me.gbMain.TabIndex = 1
        Me.gbMain.TabStop = False
        Me.gbMain.Text = "Display"
        '
        'frmMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(555, 419)
        Me.Controls.Add(Me.tvwMain)
        Me.Controls.Add(Me.gbMain)
        Me.Name = "frmMain"
        Me.RightToLeftLayout = True
        Me.Text = "Islam Source"
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents tvwMain As System.Windows.Forms.TreeView
    Friend WithEvents gbMain As System.Windows.Forms.GroupBox

End Class

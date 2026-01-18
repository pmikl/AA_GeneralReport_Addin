<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_Test
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
        Me.txtBox_Path = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'txtBox_Path
        '
        Me.txtBox_Path.Location = New System.Drawing.Point(44, 49)
        Me.txtBox_Path.Multiline = True
        Me.txtBox_Path.Name = "txtBox_Path"
        Me.txtBox_Path.Size = New System.Drawing.Size(390, 135)
        Me.txtBox_Path.TabIndex = 0
        '
        'frm_Test
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(478, 227)
        Me.Controls.Add(Me.txtBox_Path)
        Me.Name = "frm_Test"
        Me.Text = "frm_Test"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtBox_Path As Windows.Forms.TextBox
End Class

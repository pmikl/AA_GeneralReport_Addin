<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_transparency
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
        Me.lbl_backPanelWarning = New System.Windows.Forms.Label()
        Me.lbl_percent = New System.Windows.Forms.Label()
        Me.txtBox_transparencyValue = New System.Windows.Forms.TextBox()
        Me.lbl_transparent = New System.Windows.Forms.Label()
        Me.lbl_opaque = New System.Windows.Forms.Label()
        Me.lbl_instraction = New System.Windows.Forms.Label()
        Me.scrl_Transparency = New System.Windows.Forms.HScrollBar()
        Me.SuspendLayout()
        '
        'lbl_backPanelWarning
        '
        Me.lbl_backPanelWarning.AutoSize = True
        Me.lbl_backPanelWarning.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_backPanelWarning.ForeColor = System.Drawing.Color.Red
        Me.lbl_backPanelWarning.Location = New System.Drawing.Point(55, 109)
        Me.lbl_backPanelWarning.Name = "lbl_backPanelWarning"
        Me.lbl_backPanelWarning.Size = New System.Drawing.Size(278, 13)
        Me.lbl_backPanelWarning.TabIndex = 13
        Me.lbl_backPanelWarning.Text = "This section does not have a standard Image Back Panel"
        '
        'lbl_percent
        '
        Me.lbl_percent.AutoSize = True
        Me.lbl_percent.Location = New System.Drawing.Point(219, 80)
        Me.lbl_percent.Name = "lbl_percent"
        Me.lbl_percent.Size = New System.Drawing.Size(15, 13)
        Me.lbl_percent.TabIndex = 12
        Me.lbl_percent.Text = "%"
        '
        'txtBox_transparencyValue
        '
        Me.txtBox_transparencyValue.Enabled = False
        Me.txtBox_transparencyValue.Location = New System.Drawing.Point(167, 77)
        Me.txtBox_transparencyValue.Name = "txtBox_transparencyValue"
        Me.txtBox_transparencyValue.Size = New System.Drawing.Size(46, 20)
        Me.txtBox_transparencyValue.TabIndex = 11
        Me.txtBox_transparencyValue.TabStop = False
        Me.txtBox_transparencyValue.Text = "0"
        Me.txtBox_transparencyValue.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lbl_transparent
        '
        Me.lbl_transparent.AutoSize = True
        Me.lbl_transparent.Location = New System.Drawing.Point(298, 79)
        Me.lbl_transparent.Name = "lbl_transparent"
        Me.lbl_transparent.Size = New System.Drawing.Size(64, 13)
        Me.lbl_transparent.TabIndex = 10
        Me.lbl_transparent.Text = "Transparent"
        '
        'lbl_opaque
        '
        Me.lbl_opaque.AutoSize = True
        Me.lbl_opaque.Location = New System.Drawing.Point(16, 79)
        Me.lbl_opaque.Name = "lbl_opaque"
        Me.lbl_opaque.Size = New System.Drawing.Size(45, 13)
        Me.lbl_opaque.TabIndex = 9
        Me.lbl_opaque.Text = "Opaque"
        '
        'lbl_instraction
        '
        Me.lbl_instraction.AutoSize = True
        Me.lbl_instraction.Location = New System.Drawing.Point(45, 12)
        Me.lbl_instraction.Name = "lbl_instraction"
        Me.lbl_instraction.Size = New System.Drawing.Size(279, 13)
        Me.lbl_instraction.TabIndex = 8
        Me.lbl_instraction.Text = "Slide to set the the transparency for the image back panel"
        '
        'scrl_Transparency
        '
        Me.scrl_Transparency.LargeChange = 1
        Me.scrl_Transparency.Location = New System.Drawing.Point(16, 50)
        Me.scrl_Transparency.Name = "scrl_Transparency"
        Me.scrl_Transparency.Size = New System.Drawing.Size(346, 21)
        Me.scrl_Transparency.TabIndex = 7
        '
        'frm_transparency
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Silver
        Me.ClientSize = New System.Drawing.Size(378, 135)
        Me.Controls.Add(Me.lbl_backPanelWarning)
        Me.Controls.Add(Me.lbl_percent)
        Me.Controls.Add(Me.txtBox_transparencyValue)
        Me.Controls.Add(Me.lbl_transparent)
        Me.Controls.Add(Me.lbl_opaque)
        Me.Controls.Add(Me.lbl_instraction)
        Me.Controls.Add(Me.scrl_Transparency)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_transparency"
        Me.Text = "Image Back Panel Transparency"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lbl_backPanelWarning As Windows.Forms.Label
    Friend WithEvents lbl_percent As Windows.Forms.Label
    Friend WithEvents txtBox_transparencyValue As Windows.Forms.TextBox
    Friend WithEvents lbl_transparent As Windows.Forms.Label
    Friend WithEvents lbl_opaque As Windows.Forms.Label
    Friend WithEvents lbl_instraction As Windows.Forms.Label
    Friend WithEvents scrl_Transparency As Windows.Forms.HScrollBar
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_WhatsNew
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
        Me.txtBox_Instruction = New System.Windows.Forms.TextBox()
        Me.btn_Close = New System.Windows.Forms.Button()
        Me.rTxtBx_WhatsNew = New System.Windows.Forms.RichTextBox()
        Me.SuspendLayout()
        '
        'txtBox_Instruction
        '
        Me.txtBox_Instruction.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBox_Instruction.Location = New System.Drawing.Point(26, 18)
        Me.txtBox_Instruction.Multiline = True
        Me.txtBox_Instruction.Name = "txtBox_Instruction"
        Me.txtBox_Instruction.Size = New System.Drawing.Size(756, 58)
        Me.txtBox_Instruction.TabIndex = 6
        '
        'btn_Close
        '
        Me.btn_Close.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_Close.Location = New System.Drawing.Point(26, 457)
        Me.btn_Close.Name = "btn_Close"
        Me.btn_Close.Size = New System.Drawing.Size(756, 23)
        Me.btn_Close.TabIndex = 5
        Me.btn_Close.Text = "Close ""What's New"""
        Me.btn_Close.UseVisualStyleBackColor = True
        '
        'rTxtBx_WhatsNew
        '
        Me.rTxtBx_WhatsNew.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rTxtBx_WhatsNew.Location = New System.Drawing.Point(26, 82)
        Me.rTxtBx_WhatsNew.Name = "rTxtBx_WhatsNew"
        Me.rTxtBx_WhatsNew.Size = New System.Drawing.Size(756, 360)
        Me.rTxtBx_WhatsNew.TabIndex = 4
        Me.rTxtBx_WhatsNew.Text = ""
        Me.rTxtBx_WhatsNew.WordWrap = False
        '
        'frm_WhatsNew
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(808, 498)
        Me.Controls.Add(Me.txtBox_Instruction)
        Me.Controls.Add(Me.btn_Close)
        Me.Controls.Add(Me.rTxtBx_WhatsNew)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_WhatsNew"
        Me.Text = "frm_WhatsNew"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents txtBox_Instruction As Windows.Forms.TextBox
    Friend WithEvents btn_Close As Windows.Forms.Button
    Friend WithEvents rTxtBx_WhatsNew As Windows.Forms.RichTextBox
End Class

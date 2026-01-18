<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_findTables
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
        Me.components = New System.ComponentModel.Container()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lbl_Instruction_AllowedTypes = New System.Windows.Forms.Label()
        Me.btn_convertAlltoInLine = New System.Windows.Forms.Button()
        Me.lbl_instruction00 = New System.Windows.Forms.Label()
        Me.grpBox_display = New System.Windows.Forms.GroupBox()
        Me.rdBtn_Irregular = New System.Windows.Forms.RadioButton()
        Me.rdBtn_All = New System.Windows.Forms.RadioButton()
        Me.rdBtn_Floating = New System.Windows.Forms.RadioButton()
        Me.rdBtn_inLine = New System.Windows.Forms.RadioButton()
        Me.btn_refresh = New System.Windows.Forms.Button()
        Me.lstBx_plhDetail = New System.Windows.Forms.ListBox()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.GoToTheSelectedPlaceholderToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripSeparator()
        Me.ConvertToInlineToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.GroupBox1.SuspendLayout()
        Me.grpBox_display.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.SuspendLayout()
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.lbl_Instruction_AllowedTypes)
        Me.GroupBox1.Controls.Add(Me.btn_convertAlltoInLine)
        Me.GroupBox1.Controls.Add(Me.lbl_instruction00)
        Me.GroupBox1.Controls.Add(Me.grpBox_display)
        Me.GroupBox1.Controls.Add(Me.btn_refresh)
        Me.GroupBox1.Controls.Add(Me.lstBx_plhDetail)
        Me.GroupBox1.Location = New System.Drawing.Point(12, 13)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(261, 466)
        Me.GroupBox1.TabIndex = 1
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "Placeholders in the current document"
        '
        'lbl_Instruction_AllowedTypes
        '
        Me.lbl_Instruction_AllowedTypes.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbl_Instruction_AllowedTypes.AutoSize = True
        Me.lbl_Instruction_AllowedTypes.ForeColor = System.Drawing.Color.Red
        Me.lbl_Instruction_AllowedTypes.Location = New System.Drawing.Point(6, 16)
        Me.lbl_Instruction_AllowedTypes.Name = "lbl_Instruction_AllowedTypes"
        Me.lbl_Instruction_AllowedTypes.Size = New System.Drawing.Size(115, 13)
        Me.lbl_Instruction_AllowedTypes.TabIndex = 14
        Me.lbl_Instruction_AllowedTypes.Text = "Accessibility instruction"
        '
        'btn_convertAlltoInLine
        '
        Me.btn_convertAlltoInLine.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_convertAlltoInLine.Location = New System.Drawing.Point(6, 48)
        Me.btn_convertAlltoInLine.Name = "btn_convertAlltoInLine"
        Me.btn_convertAlltoInLine.Size = New System.Drawing.Size(243, 23)
        Me.btn_convertAlltoInLine.TabIndex = 13
        Me.btn_convertAlltoInLine.Text = "Force all placeholders to inline"
        Me.btn_convertAlltoInLine.UseVisualStyleBackColor = True
        '
        'lbl_instruction00
        '
        Me.lbl_instruction00.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbl_instruction00.AutoSize = True
        Me.lbl_instruction00.Location = New System.Drawing.Point(4, 346)
        Me.lbl_instruction00.Name = "lbl_instruction00"
        Me.lbl_instruction00.Size = New System.Drawing.Size(39, 13)
        Me.lbl_instruction00.TabIndex = 12
        Me.lbl_instruction00.Text = "Label1"
        '
        'grpBox_display
        '
        Me.grpBox_display.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpBox_display.Controls.Add(Me.rdBtn_Irregular)
        Me.grpBox_display.Controls.Add(Me.rdBtn_All)
        Me.grpBox_display.Controls.Add(Me.rdBtn_Floating)
        Me.grpBox_display.Controls.Add(Me.rdBtn_inLine)
        Me.grpBox_display.Location = New System.Drawing.Point(7, 382)
        Me.grpBox_display.Name = "grpBox_display"
        Me.grpBox_display.Size = New System.Drawing.Size(243, 43)
        Me.grpBox_display.TabIndex = 10
        Me.grpBox_display.TabStop = False
        Me.grpBox_display.Text = "Find and display tables;"
        '
        'rdBtn_Irregular
        '
        Me.rdBtn_Irregular.AutoSize = True
        Me.rdBtn_Irregular.Location = New System.Drawing.Point(175, 16)
        Me.rdBtn_Irregular.Name = "rdBtn_Irregular"
        Me.rdBtn_Irregular.Size = New System.Drawing.Size(63, 17)
        Me.rdBtn_Irregular.TabIndex = 9
        Me.rdBtn_Irregular.TabStop = True
        Me.rdBtn_Irregular.Text = "Irregular"
        Me.rdBtn_Irregular.UseVisualStyleBackColor = True
        '
        'rdBtn_All
        '
        Me.rdBtn_All.AutoSize = True
        Me.rdBtn_All.Location = New System.Drawing.Point(6, 16)
        Me.rdBtn_All.Name = "rdBtn_All"
        Me.rdBtn_All.Size = New System.Drawing.Size(36, 17)
        Me.rdBtn_All.TabIndex = 6
        Me.rdBtn_All.Text = "All"
        Me.rdBtn_All.UseVisualStyleBackColor = True
        '
        'rdBtn_Floating
        '
        Me.rdBtn_Floating.AutoSize = True
        Me.rdBtn_Floating.Checked = True
        Me.rdBtn_Floating.Location = New System.Drawing.Point(48, 16)
        Me.rdBtn_Floating.Name = "rdBtn_Floating"
        Me.rdBtn_Floating.Size = New System.Drawing.Size(62, 17)
        Me.rdBtn_Floating.TabIndex = 7
        Me.rdBtn_Floating.TabStop = True
        Me.rdBtn_Floating.Text = "Floating"
        Me.rdBtn_Floating.UseVisualStyleBackColor = True
        '
        'rdBtn_inLine
        '
        Me.rdBtn_inLine.AutoSize = True
        Me.rdBtn_inLine.Location = New System.Drawing.Point(116, 16)
        Me.rdBtn_inLine.Name = "rdBtn_inLine"
        Me.rdBtn_inLine.Size = New System.Drawing.Size(52, 17)
        Me.rdBtn_inLine.TabIndex = 8
        Me.rdBtn_inLine.Text = "in-line"
        Me.rdBtn_inLine.UseVisualStyleBackColor = True
        '
        'btn_refresh
        '
        Me.btn_refresh.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_refresh.Location = New System.Drawing.Point(6, 435)
        Me.btn_refresh.Name = "btn_refresh"
        Me.btn_refresh.Size = New System.Drawing.Size(244, 23)
        Me.btn_refresh.TabIndex = 4
        Me.btn_refresh.Text = "Refresh Placeholder list"
        Me.btn_refresh.UseVisualStyleBackColor = True
        '
        'lstBx_plhDetail
        '
        Me.lstBx_plhDetail.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstBx_plhDetail.ContextMenuStrip = Me.ContextMenuStrip1
        Me.lstBx_plhDetail.FormattingEnabled = True
        Me.lstBx_plhDetail.HorizontalScrollbar = True
        Me.lstBx_plhDetail.Location = New System.Drawing.Point(6, 77)
        Me.lstBx_plhDetail.Name = "lstBx_plhDetail"
        Me.lstBx_plhDetail.Size = New System.Drawing.Size(244, 264)
        Me.lstBx_plhDetail.TabIndex = 3
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.GoToTheSelectedPlaceholderToolStripMenuItem, Me.ToolStripMenuItem2, Me.ConvertToInlineToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(274, 54)
        '
        'GoToTheSelectedPlaceholderToolStripMenuItem
        '
        Me.GoToTheSelectedPlaceholderToolStripMenuItem.Name = "GoToTheSelectedPlaceholderToolStripMenuItem"
        Me.GoToTheSelectedPlaceholderToolStripMenuItem.Size = New System.Drawing.Size(273, 22)
        Me.GoToTheSelectedPlaceholderToolStripMenuItem.Text = "&Go to the selected placeholder"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(270, 6)
        '
        'ConvertToInlineToolStripMenuItem
        '
        Me.ConvertToInlineToolStripMenuItem.Name = "ConvertToInlineToolStripMenuItem"
        Me.ConvertToInlineToolStripMenuItem.Size = New System.Drawing.Size(273, 22)
        Me.ConvertToInlineToolStripMenuItem.Text = "&Convert selected placeholder to inline"
        '
        'frm_findTables
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(285, 492)
        Me.Controls.Add(Me.GroupBox1)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_findTables"
        Me.Text = "frm_findTables"
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.grpBox_display.ResumeLayout(False)
        Me.grpBox_display.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents lbl_Instruction_AllowedTypes As Windows.Forms.Label
    Friend WithEvents btn_convertAlltoInLine As Windows.Forms.Button
    Friend WithEvents lbl_instruction00 As Windows.Forms.Label
    Friend WithEvents grpBox_display As Windows.Forms.GroupBox
    Friend WithEvents rdBtn_Irregular As Windows.Forms.RadioButton
    Friend WithEvents rdBtn_All As Windows.Forms.RadioButton
    Friend WithEvents rdBtn_Floating As Windows.Forms.RadioButton
    Friend WithEvents rdBtn_inLine As Windows.Forms.RadioButton
    Friend WithEvents btn_refresh As Windows.Forms.Button
    Friend WithEvents lstBx_plhDetail As Windows.Forms.ListBox
    Friend WithEvents ContextMenuStrip1 As Windows.Forms.ContextMenuStrip
    Friend WithEvents GoToTheSelectedPlaceholderToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As Windows.Forms.ToolStripSeparator
    Friend WithEvents ConvertToInlineToolStripMenuItem As Windows.Forms.ToolStripMenuItem
End Class

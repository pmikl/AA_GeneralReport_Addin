<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_tagStyle_Map
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
        Me.btn_refresh = New System.Windows.Forms.Button()
        Me.lbl_Instruction = New System.Windows.Forms.Label()
        Me.btn_Close = New System.Windows.Forms.Button()
        Me.lbl_MainTitle = New System.Windows.Forms.Label()
        Me.lstBx_docMap = New System.Windows.Forms.ListBox()
        Me.ctx_menu_tagStyleActions = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.ctx_menuItem_GoToSection = New System.Windows.Forms.ToolStripMenuItem()
        Me.ctx_menu_tagStyleActions.SuspendLayout()
        Me.SuspendLayout()
        '
        'btn_refresh
        '
        Me.btn_refresh.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_refresh.Location = New System.Drawing.Point(187, 16)
        Me.btn_refresh.Name = "btn_refresh"
        Me.btn_refresh.Size = New System.Drawing.Size(97, 23)
        Me.btn_refresh.TabIndex = 10
        Me.btn_refresh.Text = "Refresh"
        Me.btn_refresh.UseVisualStyleBackColor = True
        '
        'lbl_Instruction
        '
        Me.lbl_Instruction.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbl_Instruction.AutoSize = True
        Me.lbl_Instruction.Location = New System.Drawing.Point(14, 410)
        Me.lbl_Instruction.Name = "lbl_Instruction"
        Me.lbl_Instruction.Size = New System.Drawing.Size(249, 13)
        Me.lbl_Instruction.TabIndex = 9
        Me.lbl_Instruction.Text = "Select and right click the section you want to go to." & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10)
        '
        'btn_Close
        '
        Me.btn_Close.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_Close.Location = New System.Drawing.Point(14, 437)
        Me.btn_Close.Name = "btn_Close"
        Me.btn_Close.Size = New System.Drawing.Size(270, 23)
        Me.btn_Close.TabIndex = 8
        Me.btn_Close.Text = "Close"
        Me.btn_Close.UseVisualStyleBackColor = True
        '
        'lbl_MainTitle
        '
        Me.lbl_MainTitle.AutoSize = True
        Me.lbl_MainTitle.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_MainTitle.Location = New System.Drawing.Point(11, 20)
        Me.lbl_MainTitle.Name = "lbl_MainTitle"
        Me.lbl_MainTitle.Size = New System.Drawing.Size(125, 13)
        Me.lbl_MainTitle.TabIndex = 7
        Me.lbl_MainTitle.Text = "Document TagStyle Map"
        '
        'lstBx_docMap
        '
        Me.lstBx_docMap.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstBx_docMap.ContextMenuStrip = Me.ctx_menu_tagStyleActions
        Me.lstBx_docMap.FormattingEnabled = True
        Me.lstBx_docMap.Location = New System.Drawing.Point(14, 48)
        Me.lstBx_docMap.Name = "lstBx_docMap"
        Me.lstBx_docMap.Size = New System.Drawing.Size(270, 355)
        Me.lstBx_docMap.TabIndex = 6
        '
        'ctx_menu_tagStyleActions
        '
        Me.ctx_menu_tagStyleActions.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.ctx_menuItem_GoToSection})
        Me.ctx_menu_tagStyleActions.Name = "ctx_menu_tagStyleActions"
        Me.ctx_menu_tagStyleActions.Size = New System.Drawing.Size(211, 26)
        '
        'ctx_menuItem_GoToSection
        '
        Me.ctx_menuItem_GoToSection.Name = "ctx_menuItem_GoToSection"
        Me.ctx_menuItem_GoToSection.Size = New System.Drawing.Size(210, 22)
        Me.ctx_menuItem_GoToSection.Text = "&Go to the selected section"
        '
        'frm_tagStyle_Map
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(294, 477)
        Me.ControlBox = False
        Me.Controls.Add(Me.btn_refresh)
        Me.Controls.Add(Me.lbl_Instruction)
        Me.Controls.Add(Me.btn_Close)
        Me.Controls.Add(Me.lbl_MainTitle)
        Me.Controls.Add(Me.lstBx_docMap)
        Me.Name = "frm_tagStyle_Map"
        Me.Text = "frm_tagStyle_Map"
        Me.ctx_menu_tagStyleActions.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents btn_refresh As Windows.Forms.Button
    Friend WithEvents lbl_Instruction As Windows.Forms.Label
    Friend WithEvents btn_Close As Windows.Forms.Button
    Friend WithEvents lbl_MainTitle As Windows.Forms.Label
    Friend WithEvents lstBx_docMap As Windows.Forms.ListBox
    Friend WithEvents ctx_menu_tagStyleActions As Windows.Forms.ContextMenuStrip
    Friend WithEvents ctx_menuItem_GoToSection As Windows.Forms.ToolStripMenuItem
End Class

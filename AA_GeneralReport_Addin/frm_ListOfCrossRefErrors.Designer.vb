<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_ListOfCrossRefErrors
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
        Me.lbl_SourceDocument = New System.Windows.Forms.Label()
        Me.txtBox_SourceDocument = New System.Windows.Forms.TextBox()
        Me.chkBox_ShowAllCrossRefFields = New System.Windows.Forms.CheckBox()
        Me.lbl_ListDescription = New System.Windows.Forms.Label()
        Me.lstBox_CrossRefErrors = New System.Windows.Forms.ListBox()
        Me.btn_Close = New System.Windows.Forms.Button()
        Me.btn_Refresh = New System.Windows.Forms.Button()
        Me.lbl_Instruction = New System.Windows.Forms.Label()
        Me.ctx_FieldsFunctions = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.DeleteSelectedFieldToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DeleteALLOfTheFieldsInTheListToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.RefreshToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem2 = New System.Windows.Forms.ToolStripSeparator()
        Me.CloseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ctx_FieldsFunctions.SuspendLayout()
        Me.SuspendLayout()
        '
        'lbl_SourceDocument
        '
        Me.lbl_SourceDocument.AutoSize = True
        Me.lbl_SourceDocument.Location = New System.Drawing.Point(9, 9)
        Me.lbl_SourceDocument.Name = "lbl_SourceDocument"
        Me.lbl_SourceDocument.Size = New System.Drawing.Size(124, 13)
        Me.lbl_SourceDocument.TabIndex = 4
        Me.lbl_SourceDocument.Text = "Source Document Name"
        '
        'txtBox_SourceDocument
        '
        Me.txtBox_SourceDocument.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBox_SourceDocument.Location = New System.Drawing.Point(12, 25)
        Me.txtBox_SourceDocument.Multiline = True
        Me.txtBox_SourceDocument.Name = "txtBox_SourceDocument"
        Me.txtBox_SourceDocument.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtBox_SourceDocument.Size = New System.Drawing.Size(248, 33)
        Me.txtBox_SourceDocument.TabIndex = 3
        '
        'chkBox_ShowAllCrossRefFields
        '
        Me.chkBox_ShowAllCrossRefFields.AutoSize = True
        Me.chkBox_ShowAllCrossRefFields.Checked = True
        Me.chkBox_ShowAllCrossRefFields.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkBox_ShowAllCrossRefFields.Location = New System.Drawing.Point(12, 397)
        Me.chkBox_ShowAllCrossRefFields.Name = "chkBox_ShowAllCrossRefFields"
        Me.chkBox_ShowAllCrossRefFields.Size = New System.Drawing.Size(218, 17)
        Me.chkBox_ShowAllCrossRefFields.TabIndex = 11
        Me.chkBox_ShowAllCrossRefFields.Text = "Show 'errored' cross reference fields only"
        Me.chkBox_ShowAllCrossRefFields.UseVisualStyleBackColor = True
        '
        'lbl_ListDescription
        '
        Me.lbl_ListDescription.AutoSize = True
        Me.lbl_ListDescription.Location = New System.Drawing.Point(9, 69)
        Me.lbl_ListDescription.Name = "lbl_ListDescription"
        Me.lbl_ListDescription.Size = New System.Drawing.Size(200, 13)
        Me.lbl_ListDescription.TabIndex = 12
        Me.lbl_ListDescription.Text = "List of Orphaned Cross Reference Fields."
        '
        'lstBox_CrossRefErrors
        '
        Me.lstBox_CrossRefErrors.Anchor = CType((((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Bottom) _
            Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lstBox_CrossRefErrors.ContextMenuStrip = Me.ctx_FieldsFunctions
        Me.lstBox_CrossRefErrors.FormattingEnabled = True
        Me.lstBox_CrossRefErrors.Location = New System.Drawing.Point(12, 85)
        Me.lstBox_CrossRefErrors.Name = "lstBox_CrossRefErrors"
        Me.lstBox_CrossRefErrors.Size = New System.Drawing.Size(248, 303)
        Me.lstBox_CrossRefErrors.TabIndex = 13
        '
        'btn_Close
        '
        Me.btn_Close.Anchor = CType(((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.btn_Close.Location = New System.Drawing.Point(128, 418)
        Me.btn_Close.Name = "btn_Close"
        Me.btn_Close.Size = New System.Drawing.Size(133, 23)
        Me.btn_Close.TabIndex = 16
        Me.btn_Close.Text = "Close"
        Me.btn_Close.UseVisualStyleBackColor = True
        '
        'btn_Refresh
        '
        Me.btn_Refresh.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btn_Refresh.Location = New System.Drawing.Point(12, 418)
        Me.btn_Refresh.Name = "btn_Refresh"
        Me.btn_Refresh.Size = New System.Drawing.Size(75, 23)
        Me.btn_Refresh.TabIndex = 15
        Me.btn_Refresh.Text = "Refresh"
        Me.btn_Refresh.UseVisualStyleBackColor = True
        '
        'lbl_Instruction
        '
        Me.lbl_Instruction.AutoSize = True
        Me.lbl_Instruction.Location = New System.Drawing.Point(11, 450)
        Me.lbl_Instruction.Name = "lbl_Instruction"
        Me.lbl_Instruction.Size = New System.Drawing.Size(204, 13)
        Me.lbl_Instruction.TabIndex = 14
        Me.lbl_Instruction.Text = "(Right click list box for available functions)"
        '
        'ctx_FieldsFunctions
        '
        Me.ctx_FieldsFunctions.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.DeleteSelectedFieldToolStripMenuItem, Me.DeleteALLOfTheFieldsInTheListToolStripMenuItem, Me.ToolStripMenuItem1, Me.RefreshToolStripMenuItem, Me.ToolStripMenuItem2, Me.CloseToolStripMenuItem})
        Me.ctx_FieldsFunctions.Name = "ctx_FieldsFunctions"
        Me.ctx_FieldsFunctions.Size = New System.Drawing.Size(247, 104)
        '
        'DeleteSelectedFieldToolStripMenuItem
        '
        Me.DeleteSelectedFieldToolStripMenuItem.Name = "DeleteSelectedFieldToolStripMenuItem"
        Me.DeleteSelectedFieldToolStripMenuItem.Size = New System.Drawing.Size(246, 22)
        Me.DeleteSelectedFieldToolStripMenuItem.Text = "&Delete selected field"
        '
        'DeleteALLOfTheFieldsInTheListToolStripMenuItem
        '
        Me.DeleteALLOfTheFieldsInTheListToolStripMenuItem.Name = "DeleteALLOfTheFieldsInTheListToolStripMenuItem"
        Me.DeleteALLOfTheFieldsInTheListToolStripMenuItem.Size = New System.Drawing.Size(246, 22)
        Me.DeleteALLOfTheFieldsInTheListToolStripMenuItem.Text = "Delete &ALL of the fields in the list"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(243, 6)
        '
        'RefreshToolStripMenuItem
        '
        Me.RefreshToolStripMenuItem.Name = "RefreshToolStripMenuItem"
        Me.RefreshToolStripMenuItem.Size = New System.Drawing.Size(246, 22)
        Me.RefreshToolStripMenuItem.Text = "&Refresh"
        '
        'ToolStripMenuItem2
        '
        Me.ToolStripMenuItem2.Name = "ToolStripMenuItem2"
        Me.ToolStripMenuItem2.Size = New System.Drawing.Size(243, 6)
        '
        'CloseToolStripMenuItem
        '
        Me.CloseToolStripMenuItem.Name = "CloseToolStripMenuItem"
        Me.CloseToolStripMenuItem.Size = New System.Drawing.Size(246, 22)
        Me.CloseToolStripMenuItem.Text = "Clo&se"
        '
        'frm_ListOfCrossRefErrors
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(280, 470)
        Me.ControlBox = False
        Me.Controls.Add(Me.btn_Close)
        Me.Controls.Add(Me.btn_Refresh)
        Me.Controls.Add(Me.lbl_Instruction)
        Me.Controls.Add(Me.lstBox_CrossRefErrors)
        Me.Controls.Add(Me.lbl_ListDescription)
        Me.Controls.Add(Me.chkBox_ShowAllCrossRefFields)
        Me.Controls.Add(Me.lbl_SourceDocument)
        Me.Controls.Add(Me.txtBox_SourceDocument)
        Me.Name = "frm_ListOfCrossRefErrors"
        Me.Text = "Cross Reference Errors"
        Me.ctx_FieldsFunctions.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lbl_SourceDocument As Windows.Forms.Label
    Friend WithEvents txtBox_SourceDocument As Windows.Forms.TextBox
    Friend WithEvents chkBox_ShowAllCrossRefFields As Windows.Forms.CheckBox
    Friend WithEvents lbl_ListDescription As Windows.Forms.Label
    Friend WithEvents lstBox_CrossRefErrors As Windows.Forms.ListBox
    Friend WithEvents btn_Close As Windows.Forms.Button
    Friend WithEvents btn_Refresh As Windows.Forms.Button
    Friend WithEvents lbl_Instruction As Windows.Forms.Label
    Friend WithEvents ctx_FieldsFunctions As Windows.Forms.ContextMenuStrip
    Friend WithEvents DeleteSelectedFieldToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents DeleteALLOfTheFieldsInTheListToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As Windows.Forms.ToolStripSeparator
    Friend WithEvents RefreshToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem2 As Windows.Forms.ToolStripSeparator
    Friend WithEvents CloseToolStripMenuItem As Windows.Forms.ToolStripMenuItem
End Class

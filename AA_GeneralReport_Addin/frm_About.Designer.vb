<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_About
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
        Me.lbl_PublishVersion = New System.Windows.Forms.Label()
        Me.lbl_updateSiteLabel = New System.Windows.Forms.Label()
        Me.txtBox_upDateSite = New System.Windows.Forms.TextBox()
        Me.lbl_author = New System.Windows.Forms.Label()
        Me.btn_OK = New System.Windows.Forms.Button()
        Me.lbl_TemplateVersion = New System.Windows.Forms.Label()
        Me.lbl_firstReleaseDate = New System.Windows.Forms.Label()
        Me.txtBox_attachedTemplate = New System.Windows.Forms.TextBox()
        Me.lbl_template = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'lbl_PublishVersion
        '
        Me.lbl_PublishVersion.AutoSize = True
        Me.lbl_PublishVersion.Location = New System.Drawing.Point(13, 15)
        Me.lbl_PublishVersion.Name = "lbl_PublishVersion"
        Me.lbl_PublishVersion.Size = New System.Drawing.Size(79, 13)
        Me.lbl_PublishVersion.TabIndex = 0
        Me.lbl_PublishVersion.Text = "Publish Version"
        '
        'lbl_updateSiteLabel
        '
        Me.lbl_updateSiteLabel.AutoSize = True
        Me.lbl_updateSiteLabel.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_updateSiteLabel.Location = New System.Drawing.Point(13, 131)
        Me.lbl_updateSiteLabel.Name = "lbl_updateSiteLabel"
        Me.lbl_updateSiteLabel.Size = New System.Drawing.Size(74, 13)
        Me.lbl_updateSiteLabel.TabIndex = 1
        Me.lbl_updateSiteLabel.Text = "Update Site"
        '
        'txtBox_upDateSite
        '
        Me.txtBox_upDateSite.Location = New System.Drawing.Point(16, 148)
        Me.txtBox_upDateSite.Multiline = True
        Me.txtBox_upDateSite.Name = "txtBox_upDateSite"
        Me.txtBox_upDateSite.Size = New System.Drawing.Size(285, 63)
        Me.txtBox_upDateSite.TabIndex = 2
        '
        'lbl_author
        '
        Me.lbl_author.AutoSize = True
        Me.lbl_author.Location = New System.Drawing.Point(13, 235)
        Me.lbl_author.Name = "lbl_author"
        Me.lbl_author.Size = New System.Drawing.Size(78, 13)
        Me.lbl_author.TabIndex = 3
        Me.lbl_author.Text = "Peter Mikelaitis"
        '
        'btn_OK
        '
        Me.btn_OK.Location = New System.Drawing.Point(16, 252)
        Me.btn_OK.Name = "btn_OK"
        Me.btn_OK.Size = New System.Drawing.Size(285, 23)
        Me.btn_OK.TabIndex = 4
        Me.btn_OK.Text = "OK"
        Me.btn_OK.UseVisualStyleBackColor = True
        '
        'lbl_TemplateVersion
        '
        Me.lbl_TemplateVersion.AutoSize = True
        Me.lbl_TemplateVersion.Location = New System.Drawing.Point(13, 39)
        Me.lbl_TemplateVersion.Name = "lbl_TemplateVersion"
        Me.lbl_TemplateVersion.Size = New System.Drawing.Size(89, 13)
        Me.lbl_TemplateVersion.TabIndex = 9
        Me.lbl_TemplateVersion.Text = "Template Version"
        '
        'lbl_firstReleaseDate
        '
        Me.lbl_firstReleaseDate.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.lbl_firstReleaseDate.AutoSize = True
        Me.lbl_firstReleaseDate.Location = New System.Drawing.Point(188, 235)
        Me.lbl_firstReleaseDate.Name = "lbl_firstReleaseDate"
        Me.lbl_firstReleaseDate.Size = New System.Drawing.Size(113, 13)
        Me.lbl_firstReleaseDate.TabIndex = 10
        Me.lbl_firstReleaseDate.Text = "First release Nov 2025"
        Me.lbl_firstReleaseDate.TextAlign = System.Drawing.ContentAlignment.TopRight
        '
        'txtBox_attachedTemplate
        '
        Me.txtBox_attachedTemplate.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.txtBox_attachedTemplate.Location = New System.Drawing.Point(16, 85)
        Me.txtBox_attachedTemplate.Multiline = True
        Me.txtBox_attachedTemplate.Name = "txtBox_attachedTemplate"
        Me.txtBox_attachedTemplate.Size = New System.Drawing.Size(287, 33)
        Me.txtBox_attachedTemplate.TabIndex = 12
        '
        'lbl_template
        '
        Me.lbl_template.AutoSize = True
        Me.lbl_template.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_template.Location = New System.Drawing.Point(15, 68)
        Me.lbl_template.Name = "lbl_template"
        Me.lbl_template.Size = New System.Drawing.Size(118, 13)
        Me.lbl_template.TabIndex = 11
        Me.lbl_template.Text = "Attached Template:"
        '
        'frm_About
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(312, 290)
        Me.Controls.Add(Me.txtBox_attachedTemplate)
        Me.Controls.Add(Me.lbl_template)
        Me.Controls.Add(Me.lbl_firstReleaseDate)
        Me.Controls.Add(Me.lbl_TemplateVersion)
        Me.Controls.Add(Me.btn_OK)
        Me.Controls.Add(Me.lbl_author)
        Me.Controls.Add(Me.txtBox_upDateSite)
        Me.Controls.Add(Me.lbl_updateSiteLabel)
        Me.Controls.Add(Me.lbl_PublishVersion)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "frm_About"
        Me.Text = "About Acil Allen Word Addin (vsto version)"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents lbl_PublishVersion As Windows.Forms.Label
    Friend WithEvents lbl_updateSiteLabel As Windows.Forms.Label
    Friend WithEvents txtBox_upDateSite As Windows.Forms.TextBox
    Friend WithEvents lbl_author As Windows.Forms.Label
    Friend WithEvents btn_OK As Windows.Forms.Button
    Friend WithEvents lbl_TemplateVersion As Windows.Forms.Label
    Friend WithEvents lbl_firstReleaseDate As Windows.Forms.Label
    Friend WithEvents txtBox_attachedTemplate As Windows.Forms.TextBox
    Friend WithEvents lbl_template As Windows.Forms.Label
End Class

<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_colorPicker
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
        Me.btn_noColour = New System.Windows.Forms.Button()
        Me.btn_doAllTblHeaders = New System.Windows.Forms.Button()
        Me.btn_getColours = New System.Windows.Forms.Button()
        Me.txtBox_RGB_Hex = New System.Windows.Forms.TextBox()
        Me.txtBox_RGB = New System.Windows.Forms.TextBox()
        Me.btn_GetColorPicker = New System.Windows.Forms.Button()
        Me.btnSpin_BorderWidth = New System.Windows.Forms.DomainUpDown()
        Me.lbl_borderSize = New System.Windows.Forms.Label()
        Me.lbl_rightClickHere = New System.Windows.Forms.Label()
        Me.btn_Close = New System.Windows.Forms.Button()
        Me.btn_changeThemeForThisWorkBook = New System.Windows.Forms.Button()
        Me.grpBox_Marker2 = New System.Windows.Forms.GroupBox()
        Me.grpBox_Marker = New System.Windows.Forms.GroupBox()
        Me.SuspendLayout()
        '
        'btn_noColour
        '
        Me.btn_noColour.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_noColour.Location = New System.Drawing.Point(487, 10)
        Me.btn_noColour.Name = "btn_noColour"
        Me.btn_noColour.Size = New System.Drawing.Size(113, 23)
        Me.btn_noColour.TabIndex = 25
        Me.btn_noColour.Text = "Cells to 'No Colour'"
        Me.btn_noColour.UseVisualStyleBackColor = True
        Me.btn_noColour.Visible = False
        '
        'btn_doAllTblHeaders
        '
        Me.btn_doAllTblHeaders.Location = New System.Drawing.Point(13, 213)
        Me.btn_doAllTblHeaders.Name = "btn_doAllTblHeaders"
        Me.btn_doAllTblHeaders.Size = New System.Drawing.Size(135, 23)
        Me.btn_doAllTblHeaders.TabIndex = 24
        Me.btn_doAllTblHeaders.Text = "Fill all Table Headers with last colour"
        Me.btn_doAllTblHeaders.UseVisualStyleBackColor = True
        Me.btn_doAllTblHeaders.Visible = False
        '
        'btn_getColours
        '
        Me.btn_getColours.Location = New System.Drawing.Point(427, 10)
        Me.btn_getColours.Name = "btn_getColours"
        Me.btn_getColours.Size = New System.Drawing.Size(173, 23)
        Me.btn_getColours.TabIndex = 23
        Me.btn_getColours.Text = "Get Colours as VB.NET"
        Me.btn_getColours.UseVisualStyleBackColor = True
        '
        'txtBox_RGB_Hex
        '
        Me.txtBox_RGB_Hex.Location = New System.Drawing.Point(34, 12)
        Me.txtBox_RGB_Hex.Name = "txtBox_RGB_Hex"
        Me.txtBox_RGB_Hex.Size = New System.Drawing.Size(70, 20)
        Me.txtBox_RGB_Hex.TabIndex = 22
        Me.txtBox_RGB_Hex.Text = "xxxxxx"
        Me.txtBox_RGB_Hex.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtBox_RGB_Hex.Visible = False
        '
        'txtBox_RGB
        '
        Me.txtBox_RGB.Location = New System.Drawing.Point(91, 12)
        Me.txtBox_RGB.Name = "txtBox_RGB"
        Me.txtBox_RGB.Size = New System.Drawing.Size(135, 20)
        Me.txtBox_RGB.TabIndex = 20
        Me.txtBox_RGB.Text = "RGB = x, x, x"
        Me.txtBox_RGB.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtBox_RGB.Visible = False
        '
        'btn_GetColorPicker
        '
        Me.btn_GetColorPicker.Location = New System.Drawing.Point(194, 10)
        Me.btn_GetColorPicker.Name = "btn_GetColorPicker"
        Me.btn_GetColorPicker.Size = New System.Drawing.Size(75, 23)
        Me.btn_GetColorPicker.TabIndex = 21
        Me.btn_GetColorPicker.Text = "Button1"
        Me.btn_GetColorPicker.UseVisualStyleBackColor = True
        Me.btn_GetColorPicker.Visible = False
        '
        'btnSpin_BorderWidth
        '
        Me.btnSpin_BorderWidth.Items.Add("none")
        Me.btnSpin_BorderWidth.Items.Add("xlHairline")
        Me.btnSpin_BorderWidth.Items.Add("xlThin")
        Me.btnSpin_BorderWidth.Items.Add("xlMedium")
        Me.btnSpin_BorderWidth.Items.Add("xlThick")
        Me.btnSpin_BorderWidth.Location = New System.Drawing.Point(15, 223)
        Me.btnSpin_BorderWidth.Name = "btnSpin_BorderWidth"
        Me.btnSpin_BorderWidth.Size = New System.Drawing.Size(67, 20)
        Me.btnSpin_BorderWidth.TabIndex = 19
        Me.btnSpin_BorderWidth.Text = "none"
        Me.btnSpin_BorderWidth.Visible = False
        '
        'lbl_borderSize
        '
        Me.lbl_borderSize.AutoSize = True
        Me.lbl_borderSize.Location = New System.Drawing.Point(88, 225)
        Me.lbl_borderSize.Name = "lbl_borderSize"
        Me.lbl_borderSize.Size = New System.Drawing.Size(61, 13)
        Me.lbl_borderSize.TabIndex = 18
        Me.lbl_borderSize.Text = "Border Size"
        Me.lbl_borderSize.Visible = False
        '
        'lbl_rightClickHere
        '
        Me.lbl_rightClickHere.AutoSize = True
        Me.lbl_rightClickHere.Location = New System.Drawing.Point(11, 205)
        Me.lbl_rightClickHere.Name = "lbl_rightClickHere"
        Me.lbl_rightClickHere.Size = New System.Drawing.Size(140, 13)
        Me.lbl_rightClickHere.TabIndex = 17
        Me.lbl_rightClickHere.Text = "Right Click here to see more"
        Me.lbl_rightClickHere.Visible = False
        '
        'btn_Close
        '
        Me.btn_Close.Location = New System.Drawing.Point(174, 243)
        Me.btn_Close.Name = "btn_Close"
        Me.btn_Close.Size = New System.Drawing.Size(54, 23)
        Me.btn_Close.TabIndex = 16
        Me.btn_Close.Text = "Close"
        Me.btn_Close.UseVisualStyleBackColor = True
        '
        'btn_changeThemeForThisWorkBook
        '
        Me.btn_changeThemeForThisWorkBook.Location = New System.Drawing.Point(13, 243)
        Me.btn_changeThemeForThisWorkBook.Name = "btn_changeThemeForThisWorkBook"
        Me.btn_changeThemeForThisWorkBook.Size = New System.Drawing.Size(135, 23)
        Me.btn_changeThemeForThisWorkBook.TabIndex = 15
        Me.btn_changeThemeForThisWorkBook.Text = "Change to AAC Theme"
        Me.btn_changeThemeForThisWorkBook.UseVisualStyleBackColor = True
        '
        'grpBox_Marker2
        '
        Me.grpBox_Marker2.Location = New System.Drawing.Point(260, 33)
        Me.grpBox_Marker2.Name = "grpBox_Marker2"
        Me.grpBox_Marker2.Size = New System.Drawing.Size(340, 211)
        Me.grpBox_Marker2.TabIndex = 14
        Me.grpBox_Marker2.TabStop = False
        Me.grpBox_Marker2.Text = "GroupBox2"
        Me.grpBox_Marker2.Visible = False
        '
        'grpBox_Marker
        '
        Me.grpBox_Marker.Location = New System.Drawing.Point(14, 33)
        Me.grpBox_Marker.Name = "grpBox_Marker"
        Me.grpBox_Marker.Size = New System.Drawing.Size(214, 144)
        Me.grpBox_Marker.TabIndex = 13
        Me.grpBox_Marker.TabStop = False
        Me.grpBox_Marker.Text = "GroupBox1"
        Me.grpBox_Marker.Visible = False
        '
        'frm_colorPicker
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(610, 277)
        Me.Controls.Add(Me.btn_noColour)
        Me.Controls.Add(Me.btn_doAllTblHeaders)
        Me.Controls.Add(Me.btn_getColours)
        Me.Controls.Add(Me.txtBox_RGB_Hex)
        Me.Controls.Add(Me.txtBox_RGB)
        Me.Controls.Add(Me.btn_GetColorPicker)
        Me.Controls.Add(Me.btnSpin_BorderWidth)
        Me.Controls.Add(Me.lbl_borderSize)
        Me.Controls.Add(Me.lbl_rightClickHere)
        Me.Controls.Add(Me.btn_Close)
        Me.Controls.Add(Me.btn_changeThemeForThisWorkBook)
        Me.Controls.Add(Me.grpBox_Marker2)
        Me.Controls.Add(Me.grpBox_Marker)
        Me.Name = "frm_colorPicker"
        Me.Text = "frm_colorPicker"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents grpBox_Marker As System.Windows.Forms.GroupBox
    Friend WithEvents grpBox_Marker2 As System.Windows.Forms.GroupBox
    Friend WithEvents btn_changeThemeForThisWorkBook As System.Windows.Forms.Button
    Friend WithEvents btn_Close As System.Windows.Forms.Button
    Friend WithEvents lbl_rightClickHere As System.Windows.Forms.Label
    Friend WithEvents lbl_borderSize As System.Windows.Forms.Label
    Friend WithEvents btnSpin_BorderWidth As System.Windows.Forms.DomainUpDown
    Friend WithEvents mnuCtx_Functions As System.Windows.Forms.ContextMenuStrip
    Friend WithEvents CloseToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As System.Windows.Forms.ToolStripSeparator
    Friend WithEvents CHnageToAAThemeToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents DoSeriesOfChartBordersToolStripMenuItem As System.Windows.Forms.ToolStripMenuItem
    Friend WithEvents txtBox_RGB As System.Windows.Forms.TextBox
    Friend WithEvents btn_GetColorPicker As Windows.Forms.Button
    Friend WithEvents txtBox_RGB_Hex As Windows.Forms.TextBox
    Friend WithEvents btn_getColours As System.Windows.Forms.Button
    Friend WithEvents btn_doAllTblHeaders As System.Windows.Forms.Button
    Friend WithEvents btn_noColour As System.Windows.Forms.Button

End Class

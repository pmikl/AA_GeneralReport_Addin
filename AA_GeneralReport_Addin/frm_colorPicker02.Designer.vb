<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_colorPicker02
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
        Me.grpBox_cellTextSelection = New System.Windows.Forms.GroupBox()
        Me.rdBtn_colourCells = New System.Windows.Forms.RadioButton()
        Me.rdBtn_colourText = New System.Windows.Forms.RadioButton()
        Me.rdBtn_Grid = New System.Windows.Forms.RadioButton()
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
        Me.btn_changeToAATheme = New System.Windows.Forms.Button()
        Me.grpBox_custClrs = New System.Windows.Forms.GroupBox()
        Me.grpBox_thmClrs = New System.Windows.Forms.GroupBox()
        Me.cmBox_themesToChoose = New System.Windows.Forms.ComboBox()
        Me.mnuCtx_Functions = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.CloseToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.ToolStripMenuItem1 = New System.Windows.Forms.ToolStripSeparator()
        Me.CHnageToAAThemeToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.DoSeriesOfChartBordersToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.grpBox_cellTextSelection.SuspendLayout()
        Me.mnuCtx_Functions.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpBox_cellTextSelection
        '
        Me.grpBox_cellTextSelection.Controls.Add(Me.rdBtn_colourCells)
        Me.grpBox_cellTextSelection.Controls.Add(Me.rdBtn_colourText)
        Me.grpBox_cellTextSelection.Controls.Add(Me.rdBtn_Grid)
        Me.grpBox_cellTextSelection.Location = New System.Drawing.Point(234, 77)
        Me.grpBox_cellTextSelection.Name = "grpBox_cellTextSelection"
        Me.grpBox_cellTextSelection.Size = New System.Drawing.Size(179, 36)
        Me.grpBox_cellTextSelection.TabIndex = 60
        Me.grpBox_cellTextSelection.TabStop = False
        '
        'rdBtn_colourCells
        '
        Me.rdBtn_colourCells.AutoSize = True
        Me.rdBtn_colourCells.Location = New System.Drawing.Point(6, 12)
        Me.rdBtn_colourCells.Margin = New System.Windows.Forms.Padding(0)
        Me.rdBtn_colourCells.Name = "rdBtn_colourCells"
        Me.rdBtn_colourCells.Size = New System.Drawing.Size(47, 17)
        Me.rdBtn_colourCells.TabIndex = 42
        Me.rdBtn_colourCells.TabStop = True
        Me.rdBtn_colourCells.Text = "Cells"
        Me.rdBtn_colourCells.UseVisualStyleBackColor = True
        '
        'rdBtn_colourText
        '
        Me.rdBtn_colourText.AutoSize = True
        Me.rdBtn_colourText.Location = New System.Drawing.Point(54, 12)
        Me.rdBtn_colourText.Margin = New System.Windows.Forms.Padding(0)
        Me.rdBtn_colourText.Name = "rdBtn_colourText"
        Me.rdBtn_colourText.Size = New System.Drawing.Size(46, 17)
        Me.rdBtn_colourText.TabIndex = 43
        Me.rdBtn_colourText.TabStop = True
        Me.rdBtn_colourText.Text = "Text"
        Me.rdBtn_colourText.UseVisualStyleBackColor = True
        '
        'rdBtn_Grid
        '
        Me.rdBtn_Grid.AutoSize = True
        Me.rdBtn_Grid.Location = New System.Drawing.Point(100, 12)
        Me.rdBtn_Grid.Margin = New System.Windows.Forms.Padding(0)
        Me.rdBtn_Grid.Name = "rdBtn_Grid"
        Me.rdBtn_Grid.Size = New System.Drawing.Size(81, 17)
        Me.rdBtn_Grid.TabIndex = 44
        Me.rdBtn_Grid.TabStop = True
        Me.rdBtn_Grid.Text = "Cell Borders"
        Me.rdBtn_Grid.UseVisualStyleBackColor = True
        '
        'btn_noColour
        '
        Me.btn_noColour.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.btn_noColour.Location = New System.Drawing.Point(475, 256)
        Me.btn_noColour.Name = "btn_noColour"
        Me.btn_noColour.Size = New System.Drawing.Size(125, 23)
        Me.btn_noColour.TabIndex = 59
        Me.btn_noColour.Text = "Cells to 'Transparent'"
        Me.btn_noColour.UseVisualStyleBackColor = True
        '
        'btn_doAllTblHeaders
        '
        Me.btn_doAllTblHeaders.Location = New System.Drawing.Point(266, 295)
        Me.btn_doAllTblHeaders.Name = "btn_doAllTblHeaders"
        Me.btn_doAllTblHeaders.Size = New System.Drawing.Size(135, 23)
        Me.btn_doAllTblHeaders.TabIndex = 58
        Me.btn_doAllTblHeaders.Text = "Fill all Table Headers with last colour"
        Me.btn_doAllTblHeaders.UseVisualStyleBackColor = True
        Me.btn_doAllTblHeaders.Visible = False
        '
        'btn_getColours
        '
        Me.btn_getColours.Location = New System.Drawing.Point(427, 10)
        Me.btn_getColours.Name = "btn_getColours"
        Me.btn_getColours.Size = New System.Drawing.Size(173, 23)
        Me.btn_getColours.TabIndex = 57
        Me.btn_getColours.Text = "Get Colours as VB.NET"
        Me.btn_getColours.UseVisualStyleBackColor = True
        '
        'txtBox_RGB_Hex
        '
        Me.txtBox_RGB_Hex.Location = New System.Drawing.Point(34, 12)
        Me.txtBox_RGB_Hex.Name = "txtBox_RGB_Hex"
        Me.txtBox_RGB_Hex.Size = New System.Drawing.Size(70, 20)
        Me.txtBox_RGB_Hex.TabIndex = 56
        Me.txtBox_RGB_Hex.Text = "xxxxxx"
        Me.txtBox_RGB_Hex.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtBox_RGB_Hex.Visible = False
        '
        'txtBox_RGB
        '
        Me.txtBox_RGB.Location = New System.Drawing.Point(91, 12)
        Me.txtBox_RGB.Name = "txtBox_RGB"
        Me.txtBox_RGB.Size = New System.Drawing.Size(135, 20)
        Me.txtBox_RGB.TabIndex = 54
        Me.txtBox_RGB.Text = "RGB = x, x, x"
        Me.txtBox_RGB.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtBox_RGB.Visible = False
        '
        'btn_GetColorPicker
        '
        Me.btn_GetColorPicker.Location = New System.Drawing.Point(174, 10)
        Me.btn_GetColorPicker.Name = "btn_GetColorPicker"
        Me.btn_GetColorPicker.Size = New System.Drawing.Size(75, 23)
        Me.btn_GetColorPicker.TabIndex = 55
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
        Me.btnSpin_BorderWidth.Location = New System.Drawing.Point(15, 288)
        Me.btnSpin_BorderWidth.Name = "btnSpin_BorderWidth"
        Me.btnSpin_BorderWidth.Size = New System.Drawing.Size(67, 20)
        Me.btnSpin_BorderWidth.TabIndex = 53
        Me.btnSpin_BorderWidth.Text = "none"
        Me.btnSpin_BorderWidth.Visible = False
        '
        'lbl_borderSize
        '
        Me.lbl_borderSize.AutoSize = True
        Me.lbl_borderSize.Location = New System.Drawing.Point(88, 295)
        Me.lbl_borderSize.Name = "lbl_borderSize"
        Me.lbl_borderSize.Size = New System.Drawing.Size(61, 13)
        Me.lbl_borderSize.TabIndex = 52
        Me.lbl_borderSize.Text = "Border Size"
        Me.lbl_borderSize.Visible = False
        '
        'lbl_rightClickHere
        '
        Me.lbl_rightClickHere.AutoSize = True
        Me.lbl_rightClickHere.Location = New System.Drawing.Point(11, 272)
        Me.lbl_rightClickHere.Name = "lbl_rightClickHere"
        Me.lbl_rightClickHere.Size = New System.Drawing.Size(140, 13)
        Me.lbl_rightClickHere.TabIndex = 51
        Me.lbl_rightClickHere.Text = "Right Click here to see more"
        Me.lbl_rightClickHere.Visible = False
        '
        'btn_Close
        '
        Me.btn_Close.Location = New System.Drawing.Point(207, 203)
        Me.btn_Close.Name = "btn_Close"
        Me.btn_Close.Size = New System.Drawing.Size(54, 23)
        Me.btn_Close.TabIndex = 50
        Me.btn_Close.Text = "Close"
        Me.btn_Close.UseVisualStyleBackColor = True
        '
        'btn_changeToAATheme
        '
        Me.btn_changeToAATheme.Location = New System.Drawing.Point(12, 203)
        Me.btn_changeToAATheme.Name = "btn_changeToAATheme"
        Me.btn_changeToAATheme.Size = New System.Drawing.Size(137, 23)
        Me.btn_changeToAATheme.TabIndex = 49
        Me.btn_changeToAATheme.Text = "Change to current Theme"
        Me.btn_changeToAATheme.UseVisualStyleBackColor = True
        '
        'grpBox_custClrs
        '
        Me.grpBox_custClrs.Location = New System.Drawing.Point(15, 76)
        Me.grpBox_custClrs.Name = "grpBox_custClrs"
        Me.grpBox_custClrs.Size = New System.Drawing.Size(213, 127)
        Me.grpBox_custClrs.TabIndex = 48
        Me.grpBox_custClrs.TabStop = False
        Me.grpBox_custClrs.Text = "GroupBox2"
        Me.grpBox_custClrs.Visible = False
        '
        'grpBox_thmClrs
        '
        Me.grpBox_thmClrs.Location = New System.Drawing.Point(14, 33)
        Me.grpBox_thmClrs.Name = "grpBox_thmClrs"
        Me.grpBox_thmClrs.Size = New System.Drawing.Size(247, 38)
        Me.grpBox_thmClrs.TabIndex = 47
        Me.grpBox_thmClrs.TabStop = False
        Me.grpBox_thmClrs.Text = "GroupBox1"
        Me.grpBox_thmClrs.Visible = False
        '
        'cmBox_themesToChoose
        '
        Me.cmBox_themesToChoose.FormattingEnabled = True
        Me.cmBox_themesToChoose.Items.AddRange(New Object() {"Most current", "Most current (manual)", "2024 theme from file", "2024 theme manual", "Legacy light theme"})
        Me.cmBox_themesToChoose.Location = New System.Drawing.Point(155, 203)
        Me.cmBox_themesToChoose.Name = "cmBox_themesToChoose"
        Me.cmBox_themesToChoose.Size = New System.Drawing.Size(73, 21)
        Me.cmBox_themesToChoose.TabIndex = 61
        '
        'mnuCtx_Functions
        '
        Me.mnuCtx_Functions.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.CloseToolStripMenuItem, Me.ToolStripMenuItem1, Me.CHnageToAAThemeToolStripMenuItem, Me.DoSeriesOfChartBordersToolStripMenuItem})
        Me.mnuCtx_Functions.Name = "mnuCtx_Functions"
        Me.mnuCtx_Functions.Size = New System.Drawing.Size(212, 76)
        '
        'CloseToolStripMenuItem
        '
        Me.CloseToolStripMenuItem.Name = "CloseToolStripMenuItem"
        Me.CloseToolStripMenuItem.Size = New System.Drawing.Size(211, 22)
        Me.CloseToolStripMenuItem.Text = "&Close"
        '
        'ToolStripMenuItem1
        '
        Me.ToolStripMenuItem1.Name = "ToolStripMenuItem1"
        Me.ToolStripMenuItem1.Size = New System.Drawing.Size(208, 6)
        '
        'CHnageToAAThemeToolStripMenuItem
        '
        Me.CHnageToAAThemeToolStripMenuItem.Name = "CHnageToAAThemeToolStripMenuItem"
        Me.CHnageToAAThemeToolStripMenuItem.Size = New System.Drawing.Size(211, 22)
        Me.CHnageToAAThemeToolStripMenuItem.Text = "Change to AA Theme"
        '
        'DoSeriesOfChartBordersToolStripMenuItem
        '
        Me.DoSeriesOfChartBordersToolStripMenuItem.Enabled = False
        Me.DoSeriesOfChartBordersToolStripMenuItem.Name = "DoSeriesOfChartBordersToolStripMenuItem"
        Me.DoSeriesOfChartBordersToolStripMenuItem.Size = New System.Drawing.Size(211, 22)
        Me.DoSeriesOfChartBordersToolStripMenuItem.Text = "Do Series of Chart Borders"
        Me.DoSeriesOfChartBordersToolStripMenuItem.Visible = False
        '
        'frm_colorPicker02
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(610, 329)
        Me.Controls.Add(Me.grpBox_cellTextSelection)
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
        Me.Controls.Add(Me.btn_changeToAATheme)
        Me.Controls.Add(Me.grpBox_custClrs)
        Me.Controls.Add(Me.grpBox_thmClrs)
        Me.Controls.Add(Me.cmBox_themesToChoose)
        Me.Name = "frm_colorPicker02"
        Me.Text = "frm_colorPicker02"
        Me.grpBox_cellTextSelection.ResumeLayout(False)
        Me.grpBox_cellTextSelection.PerformLayout()
        Me.mnuCtx_Functions.ResumeLayout(False)
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents grpBox_cellTextSelection As Windows.Forms.GroupBox
    Friend WithEvents rdBtn_colourCells As Windows.Forms.RadioButton
    Friend WithEvents rdBtn_colourText As Windows.Forms.RadioButton
    Friend WithEvents rdBtn_Grid As Windows.Forms.RadioButton
    Friend WithEvents btn_noColour As Windows.Forms.Button
    Friend WithEvents btn_doAllTblHeaders As Windows.Forms.Button
    Friend WithEvents btn_getColours As Windows.Forms.Button
    Friend WithEvents txtBox_RGB_Hex As Windows.Forms.TextBox
    Friend WithEvents txtBox_RGB As Windows.Forms.TextBox
    Friend WithEvents btn_GetColorPicker As Windows.Forms.Button
    Friend WithEvents btnSpin_BorderWidth As Windows.Forms.DomainUpDown
    Friend WithEvents lbl_borderSize As Windows.Forms.Label
    Friend WithEvents lbl_rightClickHere As Windows.Forms.Label
    Friend WithEvents btn_Close As Windows.Forms.Button
    Friend WithEvents btn_changeToAATheme As Windows.Forms.Button
    Friend WithEvents grpBox_custClrs As Windows.Forms.GroupBox
    Friend WithEvents grpBox_thmClrs As Windows.Forms.GroupBox
    Friend WithEvents cmBox_themesToChoose As Windows.Forms.ComboBox
    Friend WithEvents mnuCtx_Functions As Windows.Forms.ContextMenuStrip
    Friend WithEvents CloseToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents ToolStripMenuItem1 As Windows.Forms.ToolStripSeparator
    Friend WithEvents CHnageToAAThemeToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents DoSeriesOfChartBordersToolStripMenuItem As Windows.Forms.ToolStripMenuItem
End Class

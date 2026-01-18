<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_TableBuilder
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
        Me.grpBox_TableOptions = New System.Windows.Forms.GroupBox()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.chkBx_DataSource = New System.Windows.Forms.CheckBox()
        Me.chkBx_Note = New System.Windows.Forms.CheckBox()
        Me.chkBx_equalColumns = New System.Windows.Forms.CheckBox()
        Me.chkBx_screenUpdatingOn = New System.Windows.Forms.CheckBox()
        Me.grpBox_RowsAndColumns = New System.Windows.Forms.GroupBox()
        Me.scrlBar_numberOfColumns = New System.Windows.Forms.VScrollBar()
        Me.lbl_numColumns = New System.Windows.Forms.Label()
        Me.txtBx_numBodyColumns = New System.Windows.Forms.TextBox()
        Me.lbl_numRows = New System.Windows.Forms.Label()
        Me.scrlBar_numberOfRows = New System.Windows.Forms.VScrollBar()
        Me.txtBx_numBodyRows = New System.Windows.Forms.TextBox()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.chkBx_UnitsRow = New System.Windows.Forms.CheckBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.chkBx_HeaderRow = New System.Windows.Forms.CheckBox()
        Me.txtBx_TableWidth = New System.Windows.Forms.TextBox()
        Me.scrlBar_TableWidth = New System.Windows.Forms.VScrollBar()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.chkBx_Caption = New System.Windows.Forms.CheckBox()
        Me.rdBtn_ES = New System.Windows.Forms.RadioButton()
        Me.rdBtn_Report = New System.Windows.Forms.RadioButton()
        Me.rdBtn_App = New System.Windows.Forms.RadioButton()
        Me.rdBtn_Letter = New System.Windows.Forms.RadioButton()
        Me.chkBx_Envelope = New System.Windows.Forms.CheckBox()
        Me.btn_BuildTable = New System.Windows.Forms.Button()
        Me.btn_Cancel = New System.Windows.Forms.Button()
        Me.lbl_Text2 = New System.Windows.Forms.Label()
        Me.chkBx_doBorders = New System.Windows.Forms.CheckBox()
        Me.lbl_Offset = New System.Windows.Forms.Label()
        Me.txtBx_Offset = New System.Windows.Forms.TextBox()
        Me.rdBtn_TextSmall = New System.Windows.Forms.RadioButton()
        Me.rdBtn_TextStandard = New System.Windows.Forms.RadioButton()
        Me.chkBx_wideTable = New System.Windows.Forms.CheckBox()
        Me.grpBox_TableOptions.SuspendLayout()
        Me.GroupBox4.SuspendLayout()
        Me.grpBox_RowsAndColumns.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'grpBox_TableOptions
        '
        Me.grpBox_TableOptions.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpBox_TableOptions.Controls.Add(Me.GroupBox4)
        Me.grpBox_TableOptions.Controls.Add(Me.grpBox_RowsAndColumns)
        Me.grpBox_TableOptions.Controls.Add(Me.GroupBox3)
        Me.grpBox_TableOptions.Controls.Add(Me.GroupBox2)
        Me.grpBox_TableOptions.Controls.Add(Me.GroupBox1)
        Me.grpBox_TableOptions.Controls.Add(Me.chkBx_Envelope)
        Me.grpBox_TableOptions.Location = New System.Drawing.Point(12, 12)
        Me.grpBox_TableOptions.Name = "grpBox_TableOptions"
        Me.grpBox_TableOptions.Size = New System.Drawing.Size(393, 301)
        Me.grpBox_TableOptions.TabIndex = 0
        Me.grpBox_TableOptions.TabStop = False
        Me.grpBox_TableOptions.Text = "Table Options"
        '
        'GroupBox4
        '
        Me.GroupBox4.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox4.Controls.Add(Me.chkBx_DataSource)
        Me.GroupBox4.Controls.Add(Me.chkBx_Note)
        Me.GroupBox4.Controls.Add(Me.chkBx_equalColumns)
        Me.GroupBox4.Controls.Add(Me.chkBx_screenUpdatingOn)
        Me.GroupBox4.Location = New System.Drawing.Point(48, 234)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(278, 60)
        Me.GroupBox4.TabIndex = 20
        Me.GroupBox4.TabStop = False
        '
        'chkBx_DataSource
        '
        Me.chkBx_DataSource.AutoSize = True
        Me.chkBx_DataSource.Checked = True
        Me.chkBx_DataSource.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkBx_DataSource.Location = New System.Drawing.Point(6, 36)
        Me.chkBx_DataSource.Name = "chkBx_DataSource"
        Me.chkBx_DataSource.Size = New System.Drawing.Size(86, 17)
        Me.chkBx_DataSource.TabIndex = 12
        Me.chkBx_DataSource.Text = "Data Source"
        Me.chkBx_DataSource.UseVisualStyleBackColor = True
        '
        'chkBx_Note
        '
        Me.chkBx_Note.AutoSize = True
        Me.chkBx_Note.Location = New System.Drawing.Point(6, 13)
        Me.chkBx_Note.Name = "chkBx_Note"
        Me.chkBx_Note.Size = New System.Drawing.Size(49, 17)
        Me.chkBx_Note.TabIndex = 11
        Me.chkBx_Note.Text = "Note"
        Me.chkBx_Note.UseVisualStyleBackColor = True
        '
        'chkBx_equalColumns
        '
        Me.chkBx_equalColumns.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.chkBx_equalColumns.AutoSize = True
        Me.chkBx_equalColumns.Location = New System.Drawing.Point(157, 13)
        Me.chkBx_equalColumns.Name = "chkBx_equalColumns"
        Me.chkBx_equalColumns.Size = New System.Drawing.Size(113, 17)
        Me.chkBx_equalColumns.TabIndex = 4
        Me.chkBx_equalColumns.TabStop = False
        Me.chkBx_equalColumns.Text = "Do Equal Columns"
        Me.chkBx_equalColumns.UseVisualStyleBackColor = True
        Me.chkBx_equalColumns.Visible = False
        '
        'chkBx_screenUpdatingOn
        '
        Me.chkBx_screenUpdatingOn.Anchor = System.Windows.Forms.AnchorStyles.Bottom
        Me.chkBx_screenUpdatingOn.AutoSize = True
        Me.chkBx_screenUpdatingOn.Checked = True
        Me.chkBx_screenUpdatingOn.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkBx_screenUpdatingOn.Location = New System.Drawing.Point(157, 36)
        Me.chkBx_screenUpdatingOn.Name = "chkBx_screenUpdatingOn"
        Me.chkBx_screenUpdatingOn.Size = New System.Drawing.Size(135, 17)
        Me.chkBx_screenUpdatingOn.TabIndex = 2
        Me.chkBx_screenUpdatingOn.TabStop = False
        Me.chkBx_screenUpdatingOn.Text = "Screen Updating is ON"
        Me.chkBx_screenUpdatingOn.UseVisualStyleBackColor = True
        Me.chkBx_screenUpdatingOn.Visible = False
        '
        'grpBox_RowsAndColumns
        '
        Me.grpBox_RowsAndColumns.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.grpBox_RowsAndColumns.Controls.Add(Me.scrlBar_numberOfColumns)
        Me.grpBox_RowsAndColumns.Controls.Add(Me.lbl_numColumns)
        Me.grpBox_RowsAndColumns.Controls.Add(Me.txtBx_numBodyColumns)
        Me.grpBox_RowsAndColumns.Controls.Add(Me.lbl_numRows)
        Me.grpBox_RowsAndColumns.Controls.Add(Me.scrlBar_numberOfRows)
        Me.grpBox_RowsAndColumns.Controls.Add(Me.txtBx_numBodyRows)
        Me.grpBox_RowsAndColumns.Location = New System.Drawing.Point(48, 150)
        Me.grpBox_RowsAndColumns.Name = "grpBox_RowsAndColumns"
        Me.grpBox_RowsAndColumns.Size = New System.Drawing.Size(278, 85)
        Me.grpBox_RowsAndColumns.TabIndex = 19
        Me.grpBox_RowsAndColumns.TabStop = False
        '
        'scrlBar_numberOfColumns
        '
        Me.scrlBar_numberOfColumns.LargeChange = 1
        Me.scrlBar_numberOfColumns.Location = New System.Drawing.Point(70, 50)
        Me.scrlBar_numberOfColumns.Maximum = -1
        Me.scrlBar_numberOfColumns.Minimum = -60
        Me.scrlBar_numberOfColumns.Name = "scrlBar_numberOfColumns"
        Me.scrlBar_numberOfColumns.Size = New System.Drawing.Size(20, 20)
        Me.scrlBar_numberOfColumns.TabIndex = 5
        Me.scrlBar_numberOfColumns.Value = -3
        '
        'lbl_numColumns
        '
        Me.lbl_numColumns.AutoSize = True
        Me.lbl_numColumns.Location = New System.Drawing.Point(90, 53)
        Me.lbl_numColumns.Name = "lbl_numColumns"
        Me.lbl_numColumns.Size = New System.Drawing.Size(126, 13)
        Me.lbl_numColumns.TabIndex = 4
        Me.lbl_numColumns.Text = "Number of Body Columns"
        '
        'txtBx_numBodyColumns
        '
        Me.txtBx_numBodyColumns.Location = New System.Drawing.Point(7, 50)
        Me.txtBx_numBodyColumns.Name = "txtBx_numBodyColumns"
        Me.txtBx_numBodyColumns.Size = New System.Drawing.Size(62, 20)
        Me.txtBx_numBodyColumns.TabIndex = 10
        Me.txtBx_numBodyColumns.Text = "3"
        Me.txtBx_numBodyColumns.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lbl_numRows
        '
        Me.lbl_numRows.AutoSize = True
        Me.lbl_numRows.Location = New System.Drawing.Point(90, 24)
        Me.lbl_numRows.Name = "lbl_numRows"
        Me.lbl_numRows.Size = New System.Drawing.Size(113, 13)
        Me.lbl_numRows.TabIndex = 2
        Me.lbl_numRows.Text = "Number of Body Rows"
        '
        'scrlBar_numberOfRows
        '
        Me.scrlBar_numberOfRows.LargeChange = 1
        Me.scrlBar_numberOfRows.Location = New System.Drawing.Point(70, 20)
        Me.scrlBar_numberOfRows.Maximum = -1
        Me.scrlBar_numberOfRows.Minimum = -60
        Me.scrlBar_numberOfRows.Name = "scrlBar_numberOfRows"
        Me.scrlBar_numberOfRows.Size = New System.Drawing.Size(17, 20)
        Me.scrlBar_numberOfRows.TabIndex = 1
        Me.scrlBar_numberOfRows.Value = -3
        '
        'txtBx_numBodyRows
        '
        Me.txtBx_numBodyRows.Location = New System.Drawing.Point(7, 20)
        Me.txtBx_numBodyRows.Name = "txtBx_numBodyRows"
        Me.txtBx_numBodyRows.Size = New System.Drawing.Size(62, 20)
        Me.txtBx_numBodyRows.TabIndex = 9
        Me.txtBx_numBodyRows.Text = "6"
        Me.txtBx_numBodyRows.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GroupBox3
        '
        Me.GroupBox3.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox3.Controls.Add(Me.chkBx_UnitsRow)
        Me.GroupBox3.Location = New System.Drawing.Point(48, 113)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(278, 35)
        Me.GroupBox3.TabIndex = 18
        Me.GroupBox3.TabStop = False
        '
        'chkBx_UnitsRow
        '
        Me.chkBx_UnitsRow.AutoSize = True
        Me.chkBx_UnitsRow.Location = New System.Drawing.Point(6, 14)
        Me.chkBx_UnitsRow.Name = "chkBx_UnitsRow"
        Me.chkBx_UnitsRow.Size = New System.Drawing.Size(113, 17)
        Me.chkBx_UnitsRow.TabIndex = 8
        Me.chkBx_UnitsRow.Text = "Include Units Row"
        Me.chkBx_UnitsRow.UseVisualStyleBackColor = True
        '
        'GroupBox2
        '
        Me.GroupBox2.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox2.Controls.Add(Me.chkBx_HeaderRow)
        Me.GroupBox2.Controls.Add(Me.txtBx_TableWidth)
        Me.GroupBox2.Controls.Add(Me.scrlBar_TableWidth)
        Me.GroupBox2.Controls.Add(Me.Label1)
        Me.GroupBox2.Location = New System.Drawing.Point(8, 78)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(377, 35)
        Me.GroupBox2.TabIndex = 17
        Me.GroupBox2.TabStop = False
        '
        'chkBx_HeaderRow
        '
        Me.chkBx_HeaderRow.AutoSize = True
        Me.chkBx_HeaderRow.Checked = True
        Me.chkBx_HeaderRow.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkBx_HeaderRow.Location = New System.Drawing.Point(6, 13)
        Me.chkBx_HeaderRow.Name = "chkBx_HeaderRow"
        Me.chkBx_HeaderRow.Size = New System.Drawing.Size(124, 17)
        Me.chkBx_HeaderRow.TabIndex = 9
        Me.chkBx_HeaderRow.Text = "Include Header Row"
        Me.chkBx_HeaderRow.UseVisualStyleBackColor = True
        '
        'txtBx_TableWidth
        '
        Me.txtBx_TableWidth.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.txtBx_TableWidth.Location = New System.Drawing.Point(289, 11)
        Me.txtBx_TableWidth.Name = "txtBx_TableWidth"
        Me.txtBx_TableWidth.Size = New System.Drawing.Size(54, 20)
        Me.txtBx_TableWidth.TabIndex = 0
        Me.txtBx_TableWidth.TabStop = False
        Me.txtBx_TableWidth.Text = "174.6"
        Me.txtBx_TableWidth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtBx_TableWidth.Visible = False
        '
        'scrlBar_TableWidth
        '
        Me.scrlBar_TableWidth.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.scrlBar_TableWidth.LargeChange = 1
        Me.scrlBar_TableWidth.Location = New System.Drawing.Point(331, 9)
        Me.scrlBar_TableWidth.Maximum = -20
        Me.scrlBar_TableWidth.Minimum = -300
        Me.scrlBar_TableWidth.Name = "scrlBar_TableWidth"
        Me.scrlBar_TableWidth.Size = New System.Drawing.Size(20, 20)
        Me.scrlBar_TableWidth.TabIndex = 11
        Me.scrlBar_TableWidth.Value = -175
        Me.scrlBar_TableWidth.Visible = False
        '
        'Label1
        '
        Me.Label1.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(196, 14)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(87, 13)
        Me.Label1.TabIndex = 10
        Me.Label1.Text = "Table width (mm)"
        Me.Label1.Visible = False
        '
        'GroupBox1
        '
        Me.GroupBox1.Anchor = CType(((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Left) _
            Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.GroupBox1.Controls.Add(Me.chkBx_Caption)
        Me.GroupBox1.Controls.Add(Me.rdBtn_ES)
        Me.GroupBox1.Controls.Add(Me.rdBtn_Report)
        Me.GroupBox1.Controls.Add(Me.rdBtn_App)
        Me.GroupBox1.Controls.Add(Me.rdBtn_Letter)
        Me.GroupBox1.Location = New System.Drawing.Point(69, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(316, 35)
        Me.GroupBox1.TabIndex = 15
        Me.GroupBox1.TabStop = False
        '
        'chkBx_Caption
        '
        Me.chkBx_Caption.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.chkBx_Caption.AutoSize = True
        Me.chkBx_Caption.Checked = True
        Me.chkBx_Caption.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkBx_Caption.Location = New System.Drawing.Point(6, 14)
        Me.chkBx_Caption.Name = "chkBx_Caption"
        Me.chkBx_Caption.Size = New System.Drawing.Size(100, 17)
        Me.chkBx_Caption.TabIndex = 2
        Me.chkBx_Caption.Text = "Include Caption"
        Me.chkBx_Caption.UseVisualStyleBackColor = True
        '
        'rdBtn_ES
        '
        Me.rdBtn_ES.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rdBtn_ES.AutoSize = True
        Me.rdBtn_ES.Location = New System.Drawing.Point(116, 13)
        Me.rdBtn_ES.Name = "rdBtn_ES"
        Me.rdBtn_ES.Size = New System.Drawing.Size(39, 17)
        Me.rdBtn_ES.TabIndex = 3
        Me.rdBtn_ES.TabStop = True
        Me.rdBtn_ES.Text = "ES"
        Me.rdBtn_ES.UseVisualStyleBackColor = True
        '
        'rdBtn_Report
        '
        Me.rdBtn_Report.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rdBtn_Report.AutoSize = True
        Me.rdBtn_Report.Checked = True
        Me.rdBtn_Report.Location = New System.Drawing.Point(157, 13)
        Me.rdBtn_Report.Name = "rdBtn_Report"
        Me.rdBtn_Report.Size = New System.Drawing.Size(57, 17)
        Me.rdBtn_Report.TabIndex = 4
        Me.rdBtn_Report.TabStop = True
        Me.rdBtn_Report.Text = "Report"
        Me.rdBtn_Report.UseVisualStyleBackColor = True
        '
        'rdBtn_App
        '
        Me.rdBtn_App.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rdBtn_App.AutoSize = True
        Me.rdBtn_App.Location = New System.Drawing.Point(218, 13)
        Me.rdBtn_App.Name = "rdBtn_App"
        Me.rdBtn_App.Size = New System.Drawing.Size(44, 17)
        Me.rdBtn_App.TabIndex = 5
        Me.rdBtn_App.TabStop = True
        Me.rdBtn_App.Text = "App"
        Me.rdBtn_App.UseVisualStyleBackColor = True
        '
        'rdBtn_Letter
        '
        Me.rdBtn_Letter.Anchor = CType((System.Windows.Forms.AnchorStyles.Top Or System.Windows.Forms.AnchorStyles.Right), System.Windows.Forms.AnchorStyles)
        Me.rdBtn_Letter.AutoSize = True
        Me.rdBtn_Letter.Location = New System.Drawing.Point(270, 13)
        Me.rdBtn_Letter.Name = "rdBtn_Letter"
        Me.rdBtn_Letter.Size = New System.Drawing.Size(38, 17)
        Me.rdBtn_Letter.TabIndex = 6
        Me.rdBtn_Letter.TabStop = True
        Me.rdBtn_Letter.Text = "LT"
        Me.rdBtn_Letter.UseVisualStyleBackColor = True
        '
        'chkBx_Envelope
        '
        Me.chkBx_Envelope.AutoSize = True
        Me.chkBx_Envelope.Location = New System.Drawing.Point(14, 55)
        Me.chkBx_Envelope.Name = "chkBx_Envelope"
        Me.chkBx_Envelope.Size = New System.Drawing.Size(355, 17)
        Me.chkBx_Envelope.TabIndex = 16
        Me.chkBx_Envelope.Text = "EncapsulateTable (Caption and Source rows). Use with floating tables"
        Me.chkBx_Envelope.UseVisualStyleBackColor = True
        '
        'btn_BuildTable
        '
        Me.btn_BuildTable.Anchor = CType((System.Windows.Forms.AnchorStyles.Bottom Or System.Windows.Forms.AnchorStyles.Left), System.Windows.Forms.AnchorStyles)
        Me.btn_BuildTable.Location = New System.Drawing.Point(12, 424)
        Me.btn_BuildTable.Name = "btn_BuildTable"
        Me.btn_BuildTable.Size = New System.Drawing.Size(96, 49)
        Me.btn_BuildTable.TabIndex = 1
        Me.btn_BuildTable.Text = "Build Table"
        Me.btn_BuildTable.UseVisualStyleBackColor = True
        '
        'btn_Cancel
        '
        Me.btn_Cancel.Location = New System.Drawing.Point(312, 424)
        Me.btn_Cancel.Name = "btn_Cancel"
        Me.btn_Cancel.Size = New System.Drawing.Size(93, 49)
        Me.btn_Cancel.TabIndex = 2
        Me.btn_Cancel.Text = "Cancel/Close"
        Me.btn_Cancel.UseVisualStyleBackColor = True
        '
        'lbl_Text2
        '
        Me.lbl_Text2.AutoSize = True
        Me.lbl_Text2.ForeColor = System.Drawing.Color.Red
        Me.lbl_Text2.Location = New System.Drawing.Point(29, 381)
        Me.lbl_Text2.Name = "lbl_Text2"
        Me.lbl_Text2.Size = New System.Drawing.Size(355, 26)
        Me.lbl_Text2.TabIndex = 3
        Me.lbl_Text2.Text = "Tables will default to 'between margins'. To build a wider table, just check " & Global.Microsoft.VisualBasic.ChrW(13) & Global.Microsoft.VisualBasic.ChrW(10) & "th" &
    "e 'Wide Table' checkbox above."
        Me.lbl_Text2.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'chkBx_doBorders
        '
        Me.chkBx_doBorders.AutoSize = True
        Me.chkBx_doBorders.Checked = True
        Me.chkBx_doBorders.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkBx_doBorders.Location = New System.Drawing.Point(184, 441)
        Me.chkBx_doBorders.Name = "chkBx_doBorders"
        Me.chkBx_doBorders.Size = New System.Drawing.Size(79, 17)
        Me.chkBx_doBorders.TabIndex = 4
        Me.chkBx_doBorders.Text = "Do Borders"
        Me.chkBx_doBorders.UseVisualStyleBackColor = True
        '
        'lbl_Offset
        '
        Me.lbl_Offset.AutoSize = True
        Me.lbl_Offset.Location = New System.Drawing.Point(184, 465)
        Me.lbl_Offset.Name = "lbl_Offset"
        Me.lbl_Offset.Size = New System.Drawing.Size(120, 13)
        Me.lbl_Offset.TabIndex = 5
        Me.lbl_Offset.Text = "Offset (8mm is standard)"
        Me.lbl_Offset.Visible = False
        '
        'txtBx_Offset
        '
        Me.txtBx_Offset.Location = New System.Drawing.Point(127, 465)
        Me.txtBx_Offset.Name = "txtBx_Offset"
        Me.txtBx_Offset.Size = New System.Drawing.Size(40, 20)
        Me.txtBx_Offset.TabIndex = 6
        Me.txtBx_Offset.TabStop = False
        Me.txtBx_Offset.Text = "0.0"
        Me.txtBx_Offset.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        Me.txtBx_Offset.Visible = False
        '
        'rdBtn_TextSmall
        '
        Me.rdBtn_TextSmall.AutoSize = True
        Me.rdBtn_TextSmall.Location = New System.Drawing.Point(264, 348)
        Me.rdBtn_TextSmall.Name = "rdBtn_TextSmall"
        Me.rdBtn_TextSmall.Size = New System.Drawing.Size(96, 17)
        Me.rdBtn_TextSmall.TabIndex = 15
        Me.rdBtn_TextSmall.Text = "Use Small Text"
        Me.rdBtn_TextSmall.UseVisualStyleBackColor = True
        '
        'rdBtn_TextStandard
        '
        Me.rdBtn_TextStandard.AutoSize = True
        Me.rdBtn_TextStandard.Checked = True
        Me.rdBtn_TextStandard.Location = New System.Drawing.Point(264, 328)
        Me.rdBtn_TextStandard.Name = "rdBtn_TextStandard"
        Me.rdBtn_TextStandard.Size = New System.Drawing.Size(114, 17)
        Me.rdBtn_TextStandard.TabIndex = 17
        Me.rdBtn_TextStandard.TabStop = True
        Me.rdBtn_TextStandard.Text = "Use Standard Text"
        Me.rdBtn_TextStandard.UseVisualStyleBackColor = True
        '
        'chkBx_wideTable
        '
        Me.chkBx_wideTable.AutoSize = True
        Me.chkBx_wideTable.CheckAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkBx_wideTable.Location = New System.Drawing.Point(27, 328)
        Me.chkBx_wideTable.Name = "chkBx_wideTable"
        Me.chkBx_wideTable.Size = New System.Drawing.Size(81, 17)
        Me.chkBx_wideTable.TabIndex = 16
        Me.chkBx_wideTable.Text = "Wide Table"
        Me.chkBx_wideTable.TextAlign = System.Drawing.ContentAlignment.MiddleRight
        Me.chkBx_wideTable.UseVisualStyleBackColor = True
        Me.chkBx_wideTable.Visible = False
        '
        'frm_TableBuilder
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(414, 497)
        Me.Controls.Add(Me.rdBtn_TextSmall)
        Me.Controls.Add(Me.rdBtn_TextStandard)
        Me.Controls.Add(Me.chkBx_wideTable)
        Me.Controls.Add(Me.txtBx_Offset)
        Me.Controls.Add(Me.lbl_Offset)
        Me.Controls.Add(Me.chkBx_doBorders)
        Me.Controls.Add(Me.lbl_Text2)
        Me.Controls.Add(Me.btn_Cancel)
        Me.Controls.Add(Me.btn_BuildTable)
        Me.Controls.Add(Me.grpBox_TableOptions)
        Me.Name = "frm_TableBuilder"
        Me.Text = "frm_TableBuilder"
        Me.grpBox_TableOptions.ResumeLayout(False)
        Me.grpBox_TableOptions.PerformLayout()
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.grpBox_RowsAndColumns.ResumeLayout(False)
        Me.grpBox_RowsAndColumns.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents grpBox_TableOptions As Windows.Forms.GroupBox
    Friend WithEvents btn_BuildTable As Windows.Forms.Button
    Friend WithEvents btn_Cancel As Windows.Forms.Button
    Friend WithEvents lbl_Text2 As Windows.Forms.Label
    Friend WithEvents chkBx_doBorders As Windows.Forms.CheckBox
    Friend WithEvents lbl_Offset As Windows.Forms.Label
    Friend WithEvents txtBx_Offset As Windows.Forms.TextBox
    Friend WithEvents rdBtn_TextSmall As Windows.Forms.RadioButton
    Friend WithEvents rdBtn_TextStandard As Windows.Forms.RadioButton
    Friend WithEvents chkBx_wideTable As Windows.Forms.CheckBox
    Friend WithEvents GroupBox4 As Windows.Forms.GroupBox
    Friend WithEvents chkBx_DataSource As Windows.Forms.CheckBox
    Friend WithEvents chkBx_Note As Windows.Forms.CheckBox
    Friend WithEvents chkBx_equalColumns As Windows.Forms.CheckBox
    Friend WithEvents chkBx_screenUpdatingOn As Windows.Forms.CheckBox
    Friend WithEvents grpBox_RowsAndColumns As Windows.Forms.GroupBox
    Friend WithEvents scrlBar_numberOfColumns As Windows.Forms.VScrollBar
    Friend WithEvents lbl_numColumns As Windows.Forms.Label
    Friend WithEvents txtBx_numBodyColumns As Windows.Forms.TextBox
    Friend WithEvents lbl_numRows As Windows.Forms.Label
    Friend WithEvents scrlBar_numberOfRows As Windows.Forms.VScrollBar
    Friend WithEvents txtBx_numBodyRows As Windows.Forms.TextBox
    Friend WithEvents GroupBox3 As Windows.Forms.GroupBox
    Friend WithEvents chkBx_UnitsRow As Windows.Forms.CheckBox
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents chkBx_HeaderRow As Windows.Forms.CheckBox
    Friend WithEvents txtBx_TableWidth As Windows.Forms.TextBox
    Friend WithEvents scrlBar_TableWidth As Windows.Forms.VScrollBar
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents chkBx_Caption As Windows.Forms.CheckBox
    Friend WithEvents rdBtn_ES As Windows.Forms.RadioButton
    Friend WithEvents rdBtn_Report As Windows.Forms.RadioButton
    Friend WithEvents rdBtn_App As Windows.Forms.RadioButton
    Friend WithEvents rdBtn_Letter As Windows.Forms.RadioButton
    Friend WithEvents chkBx_Envelope As Windows.Forms.CheckBox
End Class

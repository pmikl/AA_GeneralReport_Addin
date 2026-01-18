Imports System.Windows.Forms
Imports System.Drawing
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word
Public Class frm_colorPicker02
    'Public objColorMgr As New cColorMgr()
    Public strFormMode As String            'text_Colour, seriesFill, seriesBorder
    'Public rbn As rbnPMTools
    Public objGlobals As New cGlobals()
    Public rgbColor_Selected As Long
    '
    'Public lstOfThemeButtons As New Collection()
    Public lstOfthemeToolStrips As New Collection()
    '
    Public lst_of_CustomColourButtons As New List(Of ToolStripButton)
    Public lstOfCustomColors As New Collection()
    Public lstOfSeedColors As New Collection()
    Public numColumns As Integer = 0            'Custom colours
    Public numRows As Integer               'Custom colours
    '
    Public btnHeight As Integer = 16
    Public btnWidth As Integer = 16

    Public numColumns_Theme As Integer = 0
    Public numRows_Theme As Integer = 0
    Public extraPaletteHeight As Integer = 0
    Public extraPaletteWidth As Integer = 0
    '
    Public strip As System.Windows.Forms.ToolStrip
    '
    Private _colorMatrix(8, 12) As Integer
    '
    Public Sub New(strFormMode As String)
        ' This call is required by the designer.
        InitializeComponent()

        Dim btn As ToolStripButton
        Dim btn_getClrsXML As System.Windows.Forms.Button
        Dim strCustClrsXML As String
        Dim i As Integer
        Dim btns_thmClrs As List(Of ToolStripButton)
        Dim btns_custClrs As List(Of ToolStripButton)
        Dim titleBarHeight As Integer
        Dim topOffSet, bottomOffSet As Integer
        'Dim custClrs_paletteHeight, paletteHeight As Integer
        'Dim custClrs_paletteWidth As Integer
        '
        Me.objGlobals = New cGlobals()
        '
        strCustClrsXML = ""
        'custClrs_paletteHeight = 0
        'paletteHeight = 0
        '
        'strFormMode = "testMode"
        'Me.strFormMode = strFormMode
        Me.frm_colorPicker_Rename(strFormMode)
        '
        '
        Me.btnSpin_BorderWidth.SelectedIndex = 0
        '
        titleBarHeight = RectangleToScreen(Me.ClientRectangle).Top - Me.Top
        topOffSet = 10
        bottomOffSet = 20
        '
        Me.numColumns_Theme = Me.objGlobals.glb_get_docThemeColours_Actual.Count - 2              'Reject Hyperlink and Followed Hyperlink

        Me.numRows_Theme = 1
        If Me.strFormMode = "testMode" Then Me.numRows_Theme = 6
        '
        btns_thmClrs = frm_build_themePalette(New System.Drawing.Point(12, topOffSet), "Theme Colours", Me.numColumns_Theme, Me.numRows_Theme, Me.btnHeight, Me.btnWidth)
        Me.grpBox_thmClrs.Top = topOffSet
        Me.grpBox_thmClrs.Left = 5
        Me.grpBox_thmClrs.Padding = New Padding(3, 3, 3, 0)

        'Now we must wire up the palette buttons
        For i = 0 To btns_thmClrs.Count - 1
            btn = btns_thmClrs.Item(i)
            AddHandler btn.MouseDown, AddressOf Me.btnHandler_MouseDown
            AddHandler btn.MouseHover, AddressOf Me.btnHandler_Hover
            btn.Visible = True
        Next
        '
        'Make the top row (i.e. standard theme colours) visible
        'For j = 0 To btns_thmClrs.Count - 1 Step numRows_Theme
        'If Not (j > btns_thmClrs.Count - 1) Then
        'btns_thmClrs.Item(j).Visible = True
        'End If
        'Next
        '
        btns_custClrs = frm_build_custClrsPalette(New System.Drawing.Point(260, topOffSet), "Custom Colours")
        Me.grpBox_custClrs.Top = topOffSet
        Me.grpBox_custClrs.Left = Me.grpBox_thmClrs.Left + Me.grpBox_thmClrs.Width + 10
        '
        'Now we must wire up the palette buttons
        For i = 0 To btns_custClrs.Count - 1
            btn = btns_custClrs.Item(i)
            AddHandler btn.MouseDown, AddressOf Me.btnHandler_MouseDown
            AddHandler btn.MouseHover, AddressOf Me.btnHandler_Hover
        Next
        '
        Me.grpBox_cellTextSelection.Text = "Colour.."
        '
        If Me.numRows_Theme = 1 Then
            'Me.grpBox_thmClrs.AutoSize = True
            'Me.grpBox_thmClrs.AutoSizeMode = AutoSizeMode.GrowAndShrink
            'Me.grpBox_thmClrs.Controls.Add(Me.grpBox_custClrs)
            Me.grpBox_custClrs.AutoSize = False
            Me.grpBox_custClrs.Width = Me.grpBox_thmClrs.Width
            Me.grpBox_custClrs.Top = Me.grpBox_thmClrs.Top + Me.grpBox_thmClrs.Height + 5
            Me.grpBox_custClrs.Left = Me.grpBox_thmClrs.Left
            Me.grpBox_custClrs.Text = "Custom Colours"
            'Me.grpBox_custClrs.b
            '
            Me.Width = Me.grpBox_thmClrs.Left + Me.grpBox_thmClrs.Width + 30
            '
            Me.grpBox_cellTextSelection.Text = ""
            Me.grpBox_cellTextSelection.Width = Me.grpBox_thmClrs.Width
            Me.grpBox_cellTextSelection.Top = Me.grpBox_custClrs.Top + Me.grpBox_custClrs.Height + 2
            Me.grpBox_cellTextSelection.Left = Me.grpBox_thmClrs.Left
            '
            '
            Me.btn_noColour.Height = 23
            Me.btn_noColour.Width = Me.grpBox_custClrs.Width
            Me.btn_noColour.Location = New System.Drawing.Point(Me.grpBox_custClrs.Right - Me.btn_noColour.Width, (9))
            Me.btn_noColour.Visible = False
            Me.btn_noColour.Top = Me.grpBox_cellTextSelection.Top + Me.grpBox_cellTextSelection.Height + 5
            '
            Me.btn_changeToAATheme.Left = Me.grpBox_thmClrs.Left
            Me.btn_changeToAATheme.Top = Me.grpBox_thmClrs.Top + Me.grpBox_thmClrs.Height + 5
            'Me.btn_changeThemeForThisWorkBook.Top = Me.btn_noColour.Top + Me.btn_noColour.Height + 5
            Me.btn_changeToAATheme.Top = Me.grpBox_cellTextSelection.Top + Me.grpBox_cellTextSelection.Height + 5
            Me.btn_changeToAATheme.Text = "Change to AA theme"
            Me.btn_changeToAATheme.Width = Me.grpBox_thmClrs.Width / 1.6
            Me.btn_changeToAATheme.BringToFront()
            '
            '
            Me.btn_Close.Width = Me.grpBox_custClrs.Width / 3
            Me.btn_Close.Left = Me.grpBox_custClrs.Right - Me.btn_Close.Width
            Me.btn_Close.Visible = True
            Me.btn_Close.Height = Me.btn_changeToAATheme.Height
            Me.btn_Close.Left = Me.grpBox_thmClrs.Left + Me.grpBox_thmClrs.Width - Me.btn_Close.Width
            Me.btn_Close.Top = Me.btn_changeToAATheme.Top
            Me.btn_Close.Left = Me.grpBox_custClrs.Left + Me.grpBox_custClrs.Width - Me.btn_Close.Width
            '
            Me.cmBox_themesToChoose.Height = Me.btn_changeToAATheme.Height
            Me.cmBox_themesToChoose.Width = (Me.grpBox_thmClrs.Width / 2) - 5 - 10
            Me.cmBox_themesToChoose.Top = Me.btn_Close.Top + Me.btn_Close.Height + 20
            Me.cmBox_themesToChoose.Left = Me.grpBox_thmClrs.Left + Me.grpBox_thmClrs.Width - Me.cmBox_themesToChoose.Width
            Me.cmBox_themesToChoose.SelectedIndex = 0
            Me.cmBox_themesToChoose.Visible = True
            '

            '
            btn_getClrsXML = New System.Windows.Forms.Button()
            btn_getClrsXML.AutoSize = False
            btn_getClrsXML.Font = New System.Drawing.Font("SansSerif", 8.25)
            btn_getClrsXML.Height = 24
            btn_getClrsXML.Width = Me.grpBox_custClrs.Width
            btn_getClrsXML.Margin = New Padding(0, 0, 0, 2)
            btn_getClrsXML.BackColor = Color.White
            btn_getClrsXML.Visible = True
            btn_getClrsXML.Text = "Get CustClrsXML as File"
            btn_getClrsXML.Name = "btn_getClrsXML"
            btn_getClrsXML.Top = cmBox_themesToChoose.Top + cmBox_themesToChoose.Height + 5
            btn_getClrsXML.Left = Me.grpBox_custClrs.Left
            btn_getClrsXML.Visible = False
            '
            Me.Controls.Add(btn_getClrsXML)
            '
            Me.btn_getColours.Left = btn_getClrsXML.Left
            Me.btn_getColours.Top = btn_getClrsXML.Top + btn_getClrsXML.Height + 5
            Me.btn_getColours.Width = Me.grpBox_custClrs.Width
            Me.btn_getColours.Visible = True
            '
            'Me.Height = Me.btn_getColours.Top + Me.btn_getColours.Height + bottomOffSet
            Me.Height = titleBarHeight + Me.btn_Close.Top + Me.btn_Close.Height + bottomOffSet

            '
        Else
            Me.Width = Me.grpBox_custClrs.Left + Me.grpBox_custClrs.Width + 20
            '
            Me.btn_changeToAATheme.Left = Me.grpBox_thmClrs.Left
            Me.btn_changeToAATheme.Top = Me.grpBox_thmClrs.Top + Me.grpBox_thmClrs.Height + 5
            Me.btn_changeToAATheme.Text = "Change to AA theme"
            Me.btn_changeToAATheme.Width = Me.grpBox_thmClrs.Width / 2
            '
            Me.cmBox_themesToChoose.Height = Me.btn_changeToAATheme.Height
            Me.cmBox_themesToChoose.Top = Me.btn_changeToAATheme.Top + 1
            Me.cmBox_themesToChoose.Width = (Me.grpBox_thmClrs.Width / 2) - 5 - 10
            Me.cmBox_themesToChoose.Left = Me.grpBox_thmClrs.Left + Me.grpBox_thmClrs.Width - Me.cmBox_themesToChoose.Width
            Me.cmBox_themesToChoose.SelectedIndex = 0
            '
            '
            Me.btn_Close.Width = Me.grpBox_custClrs.Width / 2
            Me.btn_Close.Left = Me.grpBox_custClrs.Right - Me.btn_Close.Width
            Me.btn_Close.Visible = True
            Me.btn_Close.Height = Me.btn_changeToAATheme.Height
            '
            If grpBox_thmClrs.Height >= grpBox_custClrs.Height Then
                Me.btn_Close.Top = Me.btn_changeToAATheme.Top
            Else
                Me.btn_Close.Top = Me.grpBox_custClrs.Top + Me.grpBox_custClrs.Height + 5
            End If
            If numRows_Theme = 1 Then
                Me.btn_Close.Left = Me.grpBox_thmClrs.Left + Me.grpBox_thmClrs.Width - Me.btn_Close.Width
            Else

            End If

            '
            Me.btn_noColour.Height = 23
            Me.btn_noColour.Width = Me.grpBox_custClrs.Width
            Me.btn_noColour.Location = New System.Drawing.Point(Me.grpBox_custClrs.Right - Me.btn_noColour.Width, (9))
            Me.btn_noColour.Visible = True
            Me.btn_noColour.Top = Me.btn_Close.Top + Me.btn_Close.Height + 5
            '
            btn_getClrsXML = New System.Windows.Forms.Button()
            btn_getClrsXML.AutoSize = False
            btn_getClrsXML.Font = New System.Drawing.Font("SansSerif", 8.25)
            btn_getClrsXML.Height = 24
            btn_getClrsXML.Width = Me.grpBox_custClrs.Width
            btn_getClrsXML.Margin = New Padding(0, 0, 0, 2)
            btn_getClrsXML.BackColor = Color.White
            btn_getClrsXML.Visible = True
            btn_getClrsXML.Text = "Get CustClrsXML as File"
            btn_getClrsXML.Name = "btn_getClrsXML"
            btn_getClrsXML.Top = Me.btn_noColour.Top + Me.btn_noColour.Height + 5
            btn_getClrsXML.Left = Me.grpBox_custClrs.Left
            btn_getClrsXML.Visible = True
            '
            AddHandler btn_getClrsXML.MouseDown, AddressOf Me.btn_getClrsXML_Handler
            Me.btn_getColours.Visible = True
            btn_getClrsXML.Visible = True

            '
            Me.Controls.Add(btn_getClrsXML)
            '
            Me.btn_getColours.Left = btn_getClrsXML.Left
            Me.btn_getColours.Top = btn_getClrsXML.Top + btn_getClrsXML.Height + 5
            Me.btn_getColours.Width = Me.grpBox_custClrs.Width
            Me.btn_getColours.Visible = True
            '
            '
            Me.Height = titleBarHeight + Me.btn_getColours.Top + Me.btn_getColours.Height + bottomOffSet



        End If

        '
        'Select Case Me.strFormMode
        'Case "testMode", "text_Colour"
        'My XML writer
        'AddHandler btn_getClrsXML.MouseDown, AddressOf Me.btn_getClrsXML_Handler
        ' Me.btn_getColours.Visible = True
        'btn_getClrsXML.Visible = True
        '
        'Case "tbl_Header_Colour_all"
        'AddHandler btn_getClrsXML.MouseDown, AddressOf Me.btn_fillHeaders_Handler
        '
        'Case Else
        'My XML writer
        'AddHandler btn_getClrsXML.MouseDown, AddressOf Me.btn_getClrsXML_Handler
        '
        'End Select

        '
        Me.rdBtn_colourCells.AutoSize = True
        Me.rdBtn_colourCells.Font = New System.Drawing.Font("SansSerif", 8.25)
        Me.rdBtn_colourCells.Visible = True
        Me.rdBtn_colourCells.Text = "Cells"
        'Me.rdBtn_colourCells.Checked = True
        'Me.rdBtn_colourCells.Location = New System.Drawing.Point(Me.grpBox_Marker2.Left, (objColorMgr.strip.Top + objColorMgr.strip.Height + 20))
        '
        Me.rdBtn_colourText.AutoSize = True
        Me.rdBtn_colourText.Font = New System.Drawing.Font("SansSerif", 8.25)
        Me.rdBtn_colourText.Visible = True
        Me.rdBtn_colourText.Text = "Text"
        'Me.rdBtn_colourText.Checked = True
        '
        Me.rdBtn_Grid.AutoSize = True
        Me.rdBtn_Grid.Font = New System.Drawing.Font("SansSerif", 8.25)
        Me.rdBtn_Grid.Visible = True
        Me.rdBtn_Grid.Text = "Cell Borders"

        'Me.rdBtn_colourText.Location = New System.Drawing.Point(Me.grpBox_Marker2.Left + Me.rdBtn_colourCells.Width + 5, (objColorMgr.strip.Top + objColorMgr.strip.Height + 20))
        '
        '
        'btn_getClrsXML.Width = Me.grpBox_thmClrs.Width
        'btn_getClrsXML.Location = New System.Drawing.Point(Me.grpBox_thmClrs.Left, (Me.btn_Close.Top))
        '
        'Me.grpBox_cellTextSelection.Text = ""
        'Me.grpBox_cellTextSelection.Width = Me.grpBox_thmClrs.Width
        'Me.grpBox_cellTextSelection.Location = New System.Drawing.Point(Me.grpBox_custClrs.Left, (Me.strip.Top + Me.strip.Height + 0))
        ' Me.grpBox_cellTextSelection.Top = Me.btn_changeThemeForThisWorkBook.Top + Me.btn_changeThemeForThisWorkBook.Height + 5
        'Me.grpBox_cellTextSelection.Left = Me.grpBox_thmClrs.Left
        'Me.Width = Me.grpBox_thmClrs.Width + Me.grpBox_custClrs.Width + 100
        'Me.Height = topOffSet + Me.grpBox_custClrs.Height + 200
        '
        '
        'Me.Height = titleBarHeight + topOffSet + objColorMgr.extraPaletteHeight + 15
        '

        'If Me.strFormMode = "testMode" Then Me.Height = titleBarHeight + Me.btn_getColours.Top + Me.btn_getColours.Height + bottomOffSet

        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        Me.ControlBox = False
        '
        'Me.Text = "AA Theme and Colour Management"
        '

        'Me.Height = Me.Height + 10
        'Me.Width = Me.grpBox_cellTextSelection.Left + Me.grpBox_cellTextSelection.Width + 10
        '
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will set the name of the form depending on which mode it
    ''' was initialised in (i.e. according to strFormMode
    ''' </summary>
    ''' <param name="strFormMode"></param>
    Public Sub frm_colorPicker_Rename(strFormMode As String)
        '
        Me.strFormMode = strFormMode
        Me.rdBtn_colourText.Checked = True
        '
        Select Case strFormMode
            Case "testMode"
                Me.Text = "Test Mode"
            Case "text_Colour"
                Me.Text = "Colour Text Mode"
            Case "backPanel"
                Me.Text = "Image Back Panel fill Mode"
            Case "tbl_Cells"
                Me.Text = "Table Cell(s) fill Mode"
                Me.rdBtn_colourCells.Checked = True
                '
            Case "tbl_cellBorders"
                Me.Text = "Do Table Cell Borders"
                Me.rdBtn_Grid.Checked = True
                '
            Case "tbl_Header_Colour_all"
                Me.Text = "Fill all Table Header Rows"
                '
            Case Else
                Me.Text = "frm_colorPicker"
                Me.rdBtn_colourText.Checked = True
        End Select
        '
finis:
        '
    End Sub
    '    '
    Public Sub frm_refresh_thmPalette()
        'Check out the current list of theme buttons in the palette. If they
        'exist then we can chnage the colour
        Dim lstOfThemeColors As Collection
        Dim objColor As cColorObj
        Dim objColorMgr As New cColorMgr()
        Dim btn As ToolStripButton
        Dim strp As ToolStrip
        Dim column, row As Integer
        Dim rgbColor As Integer
        'Dim rgbColorLng As Long
        Dim baseColor As Color
        Dim variationColor As Color
        Dim strColors As String = ""
        '
        '
        'Colours are verified as correct 20250925
        '
        lstOfThemeColors = Me.objGlobals.glb_get_docThemeColours_Actual()
        'For j = 0 To lstOfThemeColors.Count - 1
        'rgbColorLng = CLng(lstOfThemeColors(CStr(j)))
        'objColor = New cColorObj(rgbColorLng)
        'strColors = strColors + objColor.myColour.ToString() + vbCrLf
        'strColors = strColors + objColor.strColor + vbCrLf
        'Next
        '
        'MsgBox(strColors)
        strColors = ""
        '
        For column = 0 To Me.lstOfthemeToolStrips.Count - 1
            strp = Me.lstOfthemeToolStrips(CStr(column))
            objColor = New cColorObj(CLng(lstOfThemeColors(CStr(column))), Me.numRows_Theme)
            '
            'objColor.strColor
            '
            'Now do first item
            '
            For row = 0 To objColor.lstOfVariations.Count - 1
                btn = strp.Items().Item(row)
                If row = 0 Then
                    'strColors = strColors + objColor.strColor + vbCrLf
                    btn.BackColor = objColor.myColour
                    btn.ToolTipText = Me.getRGB_longForm(btn.BackColor)
                    btn.ToolTipText = Me.colorPalette_Row0_ToolTips(btn, row, column)


                Else
                    rgbColor = CInt(objColor.lstOfVariations(CStr(row)))
                    baseColor = ColorTranslator.FromWin32(rgbColor)
                    variationColor = Color.FromArgb(255, baseColor)
                    'btn = Me.lstOfThemeButtons.Item(k)
                    btn.BackColor = variationColor
                    'btn.BackColor = Color.FromArgb(255, 255, 0, 0)

                    btn.ToolTipText = Me.getRGB_longForm(btn.BackColor)
                End If
            Next
            'MsgBox(strColors)
        Next
        '
    End Sub
    '
    Private Function C(red As Integer, green As Integer, blue As Integer) As Color
        Dim myColor As Color

        myColor = Color.FromArgb(red, green, blue)
        '
        Return myColor
    End Function
    '
    Public Function frmPicker_get_RGB(colourRGB As Integer) As String
        Dim strRGB As String
        '
        strRGB = CStr(Me.frm_get_ARGB_R(colourRGB)) & "," & CStr(Me.frmPicker_get_ARGB_G(colourRGB)) & "," & CStr(Me.frmPicker_get_ARGB_B(colourRGB))
        '
        Return strRGB
    End Function

    '
    Public Function frmPicker_get_ARGB_A(colourARGB As Integer) As Integer
        Dim b As UInt32
        Dim msk As UInt32
        Dim tmp As UInt32
        '
        msk = &HFF000000UI
        b = CUInt(colourARGB)
        tmp = b And msk
        tmp = tmp / (2 ^ 24)
        '
        frmPicker_get_ARGB_A = CInt(tmp)
        'getARGB_A = CStr(tmp)
    End Function
    '
    Public Function frmPicker_get_ARGB_B(colourARGB As Integer) As Integer
        Dim b As UInt32
        Dim msk As UInt32
        Dim tmp As UInt32
        '
        msk = &HFF0000UI
        b = CUInt(colourARGB)
        tmp = b And msk
        tmp = tmp / (2 ^ 16)
        '
        frmPicker_get_ARGB_B = CInt(tmp)
        'getARGB_B = CStr(tmp)
    End Function
    '
    Public Function frmPicker_get_ARGB_G(colourARGB As Integer) As Integer
        Dim b As UInt32
        Dim msk As UInt32
        Dim tmp As UInt32
        '
        msk = &HFF00UI
        b = CUInt(colourARGB)
        tmp = b And msk
        tmp = tmp / (2 ^ 8)
        '
        frmPicker_get_ARGB_G = CInt(tmp)
        'getARGB_G = CStr(tmp)
    End Function
    '
    Public Function frm_get_ARGB_R(colourARGB As Integer) As Integer
        Dim b As UInt32
        Dim msk As UInt32
        Dim tmp As UInt32
        '
        msk = &HFF
        b = CUInt(colourARGB)
        tmp = b And msk
        '
        frm_get_ARGB_R = CInt(tmp)
        'getARGB_R = CStr(tmp)
    End Function


    Public Function frmPicker_get_ARGB(alpha As Integer, colourRGB As UInt32) As Integer
        Dim a, r, g, b As UInt16
        Dim tmp As UInt32
        Dim localColor As Color
        '
        'a = CInt(colourRGB >> 24 & 255)            ' 255
        'msk = 15794176
        tmp = colourRGB
        r = ((tmp) / 2 ^ 16) And 255                '255
        tmp = colourRGB And 65280
        g = ((tmp) / (2 ^ 8)) And (2 ^ 8 - 1)       '122
        b = (colourRGB) And 255                     '15, wrote this a bit different than above just For example
        '
        localColor = Color.FromArgb(a, r, g, b)
        frmPicker_get_ARGB = localColor.ToArgb()
        '
    End Function
    '
    Public Sub btnHandler_Hover(sender As Object, e As EventArgs)
        Dim btn As ToolStripButton
        Dim btnColor As Color
        '
        btn = sender
        btnColor = btn.BackColor
        'btn.ToolTipText = btnColor.ToString
        '
        'strMsg = "hello"
        'Me.txtBox_RGB.Text = strMsg
        '
    End Sub
    '
    '
    Public Sub btnHandler_Click(sender As Object, e As MouseEventArgs)
        Dim btn As ToolStripButton
        Dim btnColor As Color
        '
        btn = sender
        btnColor = btn.BackColor
        '
        If e.Button = MouseButtons.Right Then
            Me.txtBox_RGB.Text = "click"
        End If
        '

        '
        'strMsg = "hello"
        'Me.txtBox_RGB.Text = strMsg
        '
    End Sub
    '
    '
    Public Sub btn_fillHeaders_Handler(sender As Object, e As MouseEventArgs)

    End Sub
    '
    Public Sub btn_getClrsVBNet_Handler(sender As Object, e As MouseEventArgs)
        Dim strResult As String = ""
        'Dim dlg As DialogResult
        Dim dlg_SaveFile As New SaveFileDialog()
        Dim strFilePath As String
        Dim myStream As System.IO.StreamWriter
        '
        'strDocuments = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        '
        strFilePath = ""
        '
        'strResult = Me.colr_build_CustClrsVBNET()
        '
        dlg_SaveFile.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        dlg_SaveFile.FilterIndex = 1
        dlg_SaveFile.RestoreDirectory = True
        dlg_SaveFile.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        '
        Try
            If dlg_SaveFile.ShowDialog() = DialogResult.OK Then
                strFilePath = dlg_SaveFile.FileName
                myStream = New System.IO.StreamWriter(strFilePath, False)
                myStream.Write(strResult)
                myStream.Close()
            End If
        Catch ex As Exception
            MsgBox("Failed to write Custom Colours XML")
        End Try
        MsgBox("Custom Colours XML file successfully written")
        'System.Windows.Forms.FileDialog
        '
        'MsgBox(strResult)
    End Sub


    '
    Public Sub btn_getClrsXML_Handler(sender As Object, e As MouseEventArgs)
        Dim strResult As String
        'Dim dlg As DialogResult
        Dim dlg_SaveFile As New SaveFileDialog()
        Dim strFilePath As String
        Dim myStream As System.IO.StreamWriter
        '
        'strDocuments = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        '
        strFilePath = ""
        '
        strResult = Me.colr_build_CustClrsXML()
        '
        dlg_SaveFile.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        dlg_SaveFile.FilterIndex = 1
        dlg_SaveFile.RestoreDirectory = True
        dlg_SaveFile.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        '
        Try
            If dlg_SaveFile.ShowDialog() = DialogResult.OK Then
                strFilePath = dlg_SaveFile.FileName
                myStream = New System.IO.StreamWriter(strFilePath, False)
                myStream.Write(strResult)
                myStream.Close()
            End If
        Catch ex As Exception
            MsgBox("Failed to write Custom Colours XML")
        End Try
        MsgBox("Custom Colours XML file successfully written")
        'System.Windows.Forms.FileDialog
        '
        'MsgBox(strResult)
    End Sub
    '
    Public Sub btnHandler_MouseDown(sender As Object, e As MouseEventArgs)
        Dim objBckPanelMgr As New cBackPanelMgr()
        Dim objTblsMgr As New cTablesMgr()
        'Dim lstOfShapes As List(Of cShapeMgr)
        'Dim lstOfControls() As Control
        'Dim cShp As cShapeMgr
        Dim sect As Word.Section

        Dim btn As ToolStripButton
        Dim btnColor As Color
        Dim rgbColor As Long
        'Dim wrkSheet As Excel.Worksheet
        'Dim xlApp As Word.Application
        'Dim xlSel As Object
        Dim objColorMgr As New cColorMgr()
        Dim rng As Word.Range
        'Dim drCell As Word.Cell
        'Dim strLineWeight, strBorderWeight As String
        '
        sect = objGlobals.glb_get_wrdSect

        btn = sender
        btnColor = btn.BackColor
        rgbColor = RGB(btnColor.R, btnColor.G, btnColor.B)
        '
        'Store the selected colour for use
        Me.rgbColor_Selected = rgbColor
        '
        'xlApp = Globals.ThisAddIn.Application
        '
        'xlSel = xlApp.Selection
        'strLineWeight = Me.rbn.cmBox_lineWeight.Text
        'strBorderWeight = Me.rbn.cmBox_borderWeight.Text
        '
        If e.Button = MouseButtons.Right Then
            'Me.txtBox_RGB.Text = objColorMgr.getRGB_longForm(btnColor)
            'objColMgr.xdoSelectedItem_Fill(objColorMgr.getRGB(btnColor), "Solid", strLineWeight, strBorderWeight)
            '
        End If

        If e.Button = MouseButtons.Left Then
            'MsgBox(objColorMgr.getRGB_longForm(btnColor))

            Me.txtBox_RGB.Text = objColorMgr.getRGB_longForm(btnColor)
            Me.txtBox_RGB_Hex.Text = objColorMgr.getRGB_Hex(btnColor)
            rng = Me.objGlobals.glb_get_wrdSel.Range
            '
            Select Case Me.strFormMode
                Case "testMode"

                Case "text_Colour"
                    Try
                        Me.objGlobals.glb_get_wrdSelRngAll.Font.Color = RGB(btnColor.R, btnColor.G, btnColor.B)
                    Catch ex As Exception
                        MsgBox("Text colour failed... Have you selected some text to colour?")
                    End Try
                    '
                Case "backPanel"
                    objBckPanelMgr.pnl_reset_BackPanelColour(sect,, RGB(btnColor.R, btnColor.G, btnColor.B))

                Case "tbl_Cells"
                    Try
                        objTblsMgr.tbl_colour_set_colourOfCells(btn)
                    Catch ex As Exception
                        MsgBox("Have you selected some table cells to colour?")
                    End Try
                    '

                Case "tbl_cellBorders"
                    Try
                        If btn.ToolTipText <> "Transparent" Then
                            objTblsMgr.tbl_bordersCell_colourAndVisibility(btn, True)
                            'objTblsMgr.tbl_borders_colourAndVisibility(objGlobals.glb_get_wrdSelRngAll, True, RGB(btnColor.R, btnColor.G, btnColor.B))
                        Else
                            objTblsMgr.tbl_bordersCell_colourAndVisibility(btn, False)

                            'objTblsMgr.tbl_borders_colourAndVisibility(objGlobals.glb_get_wrdSelRngAll, False, RGB(btnColor.R, btnColor.G, btnColor.B))
                        End If
                        'objTblsMgr.tbl_colour_set_colourOfCells(rgbColor)
                    Catch ex As Exception
                        MsgBox("Have you selected some table cells?")
                    End Try
                    '

                Case "tbl_Header_Colour_all"
                    objTblsMgr.tbl_colour_set_HeaderRow(RGB(btnColor.R, btnColor.G, btnColor.B), True)
                    'Me.btn_doAllTblHeaders.Visible = True
                    'lstOfControls = Me.Controls.Find("btn_getClrsXML", True)
                    'Me.Controls.Item("btn_getClrsXML").Text = "Click to fill all Regular/Standard Table Header Rows"
                    '
                    'Me.Controls.Item("btn_getClrsXML").Visible = True
                    'Me.btn_getClrsXML.Visible = True
                Case "doc_security_levels"

                Case "doc_status"

            End Select

            '
            'For Each drCell In rng.Cells
            'drCell.Shading.BackgroundPatternColor = objColorMgr.getRGB_Int(btnColor)
            'drCell.Shading.ForegroundPatternColor = objColorMgr.getRGB_Int(btnColor)
            'Next

            'objColMgr.xdoSelectedItem_Fill(objColorMgr.getRGB(btnColor), "Solid", strLineWeight, strBorderWeight)
        End If

        '
        'strMsg = "hello"
        'Me.txtBox_RGB.Text = strMsg
        '

    End Sub


    Public Sub btnHandler_Colour(sender As Object, e As EventArgs)

    End Sub


    Private Sub btn_Close_Click(sender As Object, e As EventArgs) Handles btn_Close.Click
        Me.Close()
        Globals.ThisAddIn.frm_colorPicker02 = Nothing
    End Sub

    Private Sub DoSeriesOfChartBordersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DoSeriesOfChartBordersToolStripMenuItem.Click

    End Sub


    Private Sub btn_noColour_Click(sender As Object, e As EventArgs) Handles btn_noColour.Click

    End Sub

    Private Sub rdBtn_colourCells_CheckedChanged(sender As Object, e As EventArgs) Handles rdBtn_colourCells.CheckedChanged, rdBtn_colourText.CheckedChanged

    End Sub
    '

    '
    ''' <summary>
    ''' This function will build a colour palette based on the current them. It
    ''' will return a collection of buttons so that their event handlers can be
    ''' wired in the Form that called this building function
    ''' </summary>
    ''' <param name="location"></param>
    ''' <returns></returns>
    Public Function frm_build_custClrsPalette(location As System.Drawing.Point, title As String) As List(Of ToolStripButton)
        Dim lst, lstOfAltColours As Collection
        Dim lstOfButtons As New List(Of ToolStripButton)
        '
        Dim x As Integer = 5
        Dim y As Integer = 15
        '
        Dim btn As ToolStripButton
        Dim left, top As Integer
        Dim rgbColor As Integer
        'Dim lbl As System.Windows.Forms.Label
        Dim objCol As cColorObj
        Dim objGlobals As New cGlobals()
        Dim strActionType As String
        '
        lstOfAltColours = New Collection()
        Me.lstOfCustomColors = New Collection()
        Me.lstOfSeedColors = New Collection()
        '
        'strActionType = "fromSeed"
        strActionType = "customList"
        '
        Select Case strActionType
            Case "fromSeed"
                lstOfSeedColors = Me.colr_set_SeedColours()
                lstOfCustomColors = Me.frmPicker_get_CustomColours(lstOfSeedColors)
                '
                'Modify some colours insitu
                lstOfAltColours = Me.colr_set_CustomColours_AA_02()
                'Me.colr_modify_CustomColour(0, lstOfAltColours("0"), lstOfCustomColors)
                'Me.colr_modify_CustomColour(1, lstOfAltColours("1"), lstOfCustomColors)
                Me.grpBox_custClrs.Text = "Colours From Seed"
                '
            Case "customList"
                '
                'Must generate the seed colours from the first element of each column of the custom colours
                'and then must move every colour in each custom colour column up one (i.e. 1 to 0, 2 to 1 etc)
                'lstOfSeedColors = Me.colr_set_SeedColours()
                'Me.lstOfCustomColors = Me.colr_set_CustomColours_AA_03()
                'Me.lstOfCustomColors = Me.colr_set_CustomColours_AA_7x6_20250926()                   'cols x rows = 7 x 6
                '
                'Preferred
                'Me.lstOfCustomColors = Me.colr_set_CustomColours_AA_02()
                Me.lstOfCustomColors = Me.colr_set_CustomColours_AA_10x5_20250926()                     'cols = 10 x rows = 5 ... AA standard custom colours
                '
                lstOfSeedColors = Me.colr_get_SeedColours(Me.lstOfCustomColors)
                Me.grpBox_custClrs.Text = title
        End Select
        'lstOfSeedColors = Me.colr_set_SeedColours()
        'lstOfCustomColors = Me.getCustomColours(lstOfSeedColors)

        '
        '
        lst = lstOfCustomColors("0")
        '
        'Me.colr_modify_CustomColour(RGB(255, 0, 0), 0, 0, lstOfCustomColors)

        '
        numColumns = lstOfCustomColors.Count
        numRows = lst.Count
        '
        left = location.X
        top = location.Y
        'frm_build_CustomColorPalette_AsBtns = New Collection()
        '
        'lbl = New System.Windows.Forms.Label()
        'lbl.Text = title
        'lbl.AutoSize = True
        'lbl.Location = New System.Drawing.Point(location.X, location.Y - 15)
        'Me.Controls.Add(lbl)
        '
        grpBox_custClrs.AutoSize = True
        grpBox_custClrs.AutoSizeMode = AutoSizeMode.GrowAndShrink
        '
        For column = 0 To numColumns - 1
            strip = New ToolStrip()
            strip.Dock = False
            strip.AutoSize = True
            strip.Location = New System.Drawing.Point(left + column * 18, top)
            strip.Width = 18
            strip.Height = 200
            strip.CanOverflow = True
            strip.Margin = New Padding(0, 0, 0, 0)
            '
            strip.LayoutStyle = ToolStripLayoutStyle.VerticalStackWithOverflow
            strip.BackColor = Color.Transparent
            '
            'Get the first column.. which is stored as a collection, in the collection
            lst = lstOfCustomColors(CStr(column))
            Try
                For row = 0 To numRows
                    'objColor = New cColourObj(CInt(_colorMatrix(0, column)))
                    'objColor = New cColourObj(CInt(lst(CStr(column))))

                    btn = New ToolStripButton()
                    btn.AutoSize = False
                    btn.Height = 16
                    btn.Width = 16
                    btn.Margin = New Padding(0, 0, 0, 2)
                    If row = 0 Then btn.Margin = New Padding(0, 0, 0, 8)
                    btn.BackColor = Color.DarkOrange

                    '
                    'rgbColor = CInt(objColor.lstOfVariations(CStr(row - 1)))
                    'baseColor = ColorTranslator.FromWin32(CInt(lst(CStr(row))))
                    'variationColor = Color.FromArgb(0, baseColor)

                    'strHex = Hex(ColorTranslator.ToWin32(variationColor))
                    'btn.BackColor = baseColor
                    '
                    Select Case row
                        Case 0
                            '
                            Select Case strActionType
                                Case "fromSeed", "customList"
                                    rgbColor = CInt(lstOfSeedColors(CStr(column)))
                                Case ""
                                    rgbColor = CInt(lstOfCustomColors(CStr(column))("0"))
                            End Select
                            objCol = New cColorObj(rgbColor)
                            btn.BackColor = objCol.myColour
                            If rgbColor = RGB(255, 255, 255) And column = 5 Then
                                Select Case Me.strFormMode
                                    Case "testMode"
                                        btn.BackColor = objCol.myColour
                                        btn.ToolTipText = Me.getRGB_longForm(btn.BackColor)
                                    Case Else
                                        btn.BackColor = Color.Transparent
                                        btn.ToolTipText = "Transparent"
                                End Select
                            Else
                                btn.BackColor = objCol.myColour
                                btn.ToolTipText = Me.getRGB_longForm(btn.BackColor)
                            End If

                        Case Else
                            Select Case strActionType
                                Case "fromSeed"
                                    rgbColor = CInt(lst(CStr(row - 1)))
                                Case "customList"
                                    rgbColor = CInt(lst(CStr(row - 1)))
                            End Select

                            objCol = New cColorObj(rgbColor)
                            If rgbColor = RGB(255, 255, 255) And column = 5 Then
                                Select Case Me.strFormMode
                                    Case "testMode"
                                        btn.BackColor = objCol.myColour
                                        btn.ToolTipText = Me.getRGB_longForm(btn.BackColor)
                                    Case Else
                                        btn.BackColor = Color.Transparent
                                        btn.ToolTipText = "Transparent"
                                End Select
                            Else
                                btn.BackColor = objCol.myColour
                                btn.ToolTipText = Me.getRGB_longForm(btn.BackColor)
                            End If
                            'btn.BackColor = Color.Transparent
                            '
                            'Select Case column
                            'Case 4, 5, 6
                            'btn.BackColor = Color.Transparent
                            'End Select

                    End Select
                    '
                    'btn.ToolTipText = Me.getRGB_longForm(btn.BackColor)
                    '
                    strip.Items.Add(btn)
                    lstOfButtons.Add(btn)
                    '
                Next
                '
                Me.grpBox_custClrs.Controls.Add(strip)
                '
                x = column * 18 + 5
                If column = 0 Then x = 5
                '
                strip.Location = New System.Drawing.Point(x, y)
                Me.grpBox_custClrs.Visible = True
                Me.grpBox_custClrs.BringToFront()
                'Me.grpBox_custClrs.BackColor = tra
                'Me.Controls.Add(strip)
                '
                'Include the label height with the palette height
                '
                Me.extraPaletteHeight = strip.Height
                Me.extraPaletteWidth = strip.Width * (numColumns + 1)
                '
                'Me.grpBox_custClrs.Height = strip.Height + 20
                'Me.grpBox_custClrs.Width = Me.extraPaletteWidth
                '
                Me.lst_of_CustomColourButtons = lstOfButtons
                '
            Catch ex As Exception

            End Try

        Next
        '
        Return lstOfButtons
        '
    End Function
    '
    Public Function frm_build_themePalette(location As System.Drawing.Point, title As String, thm_numColumns As Integer, thm_numRows As Integer, btnHeight As Integer, btnWidth As Integer) As List(Of ToolStripButton)
        'Dim thm_numColumns, thm_numRows As Integer
        Dim lstOfButtons As New List(Of ToolStripButton)
        '
        Dim x As Integer = 5
        Dim y As Integer = 15
        '
        Dim lstOfThemeColors As Collection
        Dim strip As ToolStrip
        Dim btn As ToolStripButton
        Dim objColor As cColorObj
        Dim left, top As Integer
        Dim rgbColor As Integer
        Dim baseColor As Color
        Dim variationColor As Color
        'Dim lbl As System.Windows.Forms.Label
        Dim strButtonTip As String
        '
        lstOfThemeColors = Me.objGlobals.glb_get_docThemeColours_Actual()
        Me.lstOfthemeToolStrips = New Collection()
        '
        '
        'Now save these values for later use... if we want to refresh the palette
        'Me.numColumns_Theme = thm_numColumns
        'Me.numRows_Theme = thm_numRows
        '
        left = location.X
        top = location.Y
        lstOfButtons = New List(Of ToolStripButton)
        '
        strButtonTip = ""
        '
        grpBox_thmClrs.AutoSize = True
        grpBox_thmClrs.AutoSizeMode = AutoSizeMode.GrowAndShrink
        grpBox_thmClrs.Text = title
        '
        For column = 0 To thm_numColumns - 1
            strip = New ToolStrip()
            'strip.Name = 
            strip.Dock = False
            strip.AutoSize = True
            strip.Location = New System.Drawing.Point(left + column * 18, top)
            strip.Width = 18
            'strip.Height = thm_numRows * btnHeight + btnHeight
            strip.CanOverflow = True
            strip.Margin = New Padding(0, 0, 0, 0)
            '
            strip.LayoutStyle = ToolStripLayoutStyle.VerticalStackWithOverflow
            strip.BackColor = Color.Transparent
            '
            Me.lstOfthemeToolStrips.Add(strip, CStr(column))
            '
            For row = 0 To thm_numRows - 1
                objColor = New cColorObj(CInt(lstOfThemeColors(CStr(column))), thm_numRows)
                'objColor = New cColorObj(RGB(255, 0, 0))

                btn = New ToolStripButton()
                btn.AutoSize = False
                btn.Height = btnHeight
                btn.Width = btnWidth
                btn.Margin = New Padding(0, 0, 0, 2)
                btn.BackColor = Color.DarkOrange
                'AddHandler btn.Click, AddressOf Me.btnHandler_Colour
                'btn.OnClick(New EventArgs())
                strip.Items.Add(btn)
                '
lbl_firstRowadjustment:
                If row = 0 Then
                    'When we resize the group box, we need to remove the bottom padding to ensure
                    'that for 1 row set of colours, the distance to the bottom is the same as for
                    'multi row... and not an extra 8 pts... See line lbl_resize: below
                    '
                    btn.Margin = New Padding(0, 0, 0, 8)
                    If thm_numRows = 1 Then
                        btn.Margin = New Padding(0, 0, 0, 0)
                    End If
                    btn.BackColor = objColor.myColour
                    '
                    '******
                    strButtonTip = Me.colorPalette_Row0_ToolTips(btn, row, column)
                Else
                    Try
                        rgbColor = CInt(objColor.lstOfVariations(CStr(row)))
                        baseColor = ColorTranslator.FromWin32(rgbColor)
                        variationColor = Color.FromArgb(255, baseColor)
                        btn.BackColor = variationColor
                        '
                        '
                        '*****
                        'strButtonTip = btn.BackColor.ToString()
                        strButtonTip = Me.getRGB_longForm(btn.BackColor)

                    Catch ex As Exception

                    End Try
                End If
                '
                btn.ToolTipText = strButtonTip
                lstOfButtons.Add(btn)                   '0 based
                strButtonTip = ""
            Next
            Me.grpBox_thmClrs.Controls.Add(strip)
            'Me.Controls.Add(strip)
            '
            x = column * 18 + 5
            If column = 0 Then x = 5
            '
            strip.Location = New System.Drawing.Point(x, y)
            Me.grpBox_thmClrs.Visible = True
            Me.grpBox_thmClrs.BringToFront()
            '
            'Me.grpBox_thmClrs.Height = 10

            Me.extraPaletteHeight = strip.Height
            Me.extraPaletteWidth = strip.Width * (thm_numColumns + 1)
            '
        Next
        '
lbl_resize:
        'Resize the Group Box... also see line lbl_firstRowadjustment:
        Dim currentHeight As Integer
        currentHeight = Me.grpBox_thmClrs.Height
        Me.grpBox_thmClrs.AutoSize = False
        Me.grpBox_thmClrs.Height = currentHeight - (btnHeight / 2)
        Me.grpBox_thmClrs.Width = Me.extraPaletteWidth + 3
        '
        '
        Return lstOfButtons
        '
    End Function
    '
    '
    Public Function buildColorPalette(location As System.Drawing.Point, title As String, ByRef frm As frm_colorPicker02) As Collection
        Dim thm_numColumns, thm_numRows As Integer
        Dim lstOfThemeColors As Collection
        Dim strip As ToolStrip
        Dim btn As ToolStripButton
        Dim objColor As cColorObj
        Dim left, top As Integer
        Dim rgbColor As Integer
        Dim baseColor As Color
        Dim variationColor As Color
        Dim lbl As System.Windows.Forms.Label
        Dim strButtonTip As String
        '
        lstOfThemeColors = Me.objGlobals.glb_get_docThemeColours_Actual()
        Me.lstOfthemeToolStrips = New Collection()
        '
        thm_numColumns = lstOfThemeColors.Count
        'numRows = 7
        thm_numRows = 6
        '
        'Now save these values for later use... if we want to refresh the palette
        Me.numColumns_Theme = thm_numColumns
        Me.numRows_Theme = thm_numRows
        '
        left = location.X
        top = location.Y
        buildColorPalette = New Collection()
        '
        lbl = New System.Windows.Forms.Label()
        lbl.Text = title
        lbl.AutoSize = True
        lbl.Location = New System.Drawing.Point(location.X, location.Y - 15)
        lbl.Name = "lbl_ThemePalette"
        frm.Controls.Add(lbl)
        '
        strButtonTip = ""
        '
        For column = 0 To thm_numColumns - 1
            strip = New ToolStrip()
            'strip.Name = 
            strip.Dock = False
            strip.AutoSize = True
            strip.Location = New System.Drawing.Point(left + column * 18, top)
            strip.Width = 18
            strip.Height = 200
            strip.CanOverflow = True
            strip.Margin = New Padding(0, 0, 0, 0)
            '
            strip.LayoutStyle = ToolStripLayoutStyle.VerticalStackWithOverflow
            strip.BackColor = Color.Transparent
            '
            Me.lstOfthemeToolStrips.Add(strip, CStr(column))
            '
            For row = 0 To thm_numRows - 1
                objColor = New cColorObj(CInt(lstOfThemeColors(CStr(column))), thm_numRows)
                'objColor = New cColorObj(RGB(255, 0, 0))

                btn = New ToolStripButton()
                btn.AutoSize = False
                btn.Height = 16
                btn.Width = 16
                btn.Margin = New Padding(0, 0, 0, 2)
                btn.BackColor = Color.DarkOrange
                'AddHandler btn.Click, AddressOf Me.btnHandler_Colour
                'btn.OnClick(New EventArgs())
                strip.Items.Add(btn)
                If row = 0 Then
                    btn.Margin = New Padding(0, 0, 0, 8)
                    btn.BackColor = objColor.myColour
                    '
                    '******
                    strButtonTip = Me.colorPalette_Row0_ToolTips(btn, row, column)
                Else
                    Try
                        rgbColor = CInt(objColor.lstOfVariations(CStr(row)))
                        baseColor = ColorTranslator.FromWin32(rgbColor)
                        variationColor = Color.FromArgb(255, baseColor)
                        btn.BackColor = variationColor
                        '
                        '*****
                        'strButtonTip = btn.BackColor.ToString()
                        strButtonTip = Me.getRGB_longForm(btn.BackColor)

                    Catch ex As Exception

                    End Try
                End If
                '
                '
                '
                'strButtonTip = "(0,0,0)"
                'If row = 0 And column = 0 Then strButtonTip = "Text/Background - Dark 1" + " " + strButtonTip
                ' If row = 0 And column = 1 Then strButtonTip = "Text/Background - Light 1" + " " + strButtonTip
                'If row = 0 And column = 2 Then strButtonTip = "Text/Background - Dark 2" + " " + strButtonTip
                'If row = 0 And column = 3 Then strButtonTip = "Text/Background - Light 2" + " " + strButtonTip
                '
                'If row = 0 And column = 4 Then strButtonTip = "Accent 1" + " " + strButtonTip
                'If row = 0 And column = 5 Then strButtonTip = "Accent 2" + " " + strButtonTip
                'If row = 0 And column = 6 Then strButtonTip = "Accent 3" + " " + strButtonTip
                'If row = 0 And column = 7 Then strButtonTip = "Accent 4" + " " + strButtonTip
                'If row = 0 And column = 8 Then strButtonTip = "Accent 5" + " " + strButtonTip
                'If row = 0 And column = 9 Then strButtonTip = "Accent 6" + " " + strButtonTip
                '
                'If row = 0 And column = 10 Then strButtonTip = "HyperLink" + " " + strButtonTip
                'If row = 0 And column = 11 Then strButtonTip = "Followed HyperLink" + " " + strButtonTip
                '
                btn.ToolTipText = strButtonTip
                buildColorPalette.Add(btn)
                strButtonTip = ""
            Next
            'Me.lstOfThemeButtons = buildColorPalette
            frm.Controls.Add(strip)
            '
            Me.extraPaletteHeight = strip.Height + lbl.Height
            Me.extraPaletteWidth = strip.Width * (thm_numColumns + 1)
            '
            frm.grpBox_thmClrs.Height = strip.Height
            frm.grpBox_thmClrs.Width = Me.extraPaletteWidth
            '
        Next

    End Function
    '
    '
    Public Function colorPalette_Row0_ToolTips(ByRef btn As ToolStripButton, row As Integer, column As Integer) As String
        Dim strButtonTip As String

        'strButtonTip = btn.BackColor.ToString()
        strButtonTip = "(" + Me.getRGB_longForm(btn.BackColor) + ")"
        '
        'strButtonTip = "(0,0,0)"
        If row = 0 And column = 0 Then strButtonTip = "Text/Background - Dark 1" + " " + strButtonTip
        If row = 0 And column = 1 Then strButtonTip = "Text/Background - Light 1" + " " + strButtonTip
        If row = 0 And column = 2 Then strButtonTip = "Text/Background - Dark 2" + " " + strButtonTip
        If row = 0 And column = 3 Then strButtonTip = "Text/Background - Light 2" + " " + strButtonTip
        '
        If row = 0 And column = 4 Then strButtonTip = "Accent 1" + " " + strButtonTip
        If row = 0 And column = 5 Then strButtonTip = "Accent 2" + " " + strButtonTip
        If row = 0 And column = 6 Then strButtonTip = "Accent 3" + " " + strButtonTip
        If row = 0 And column = 7 Then strButtonTip = "Accent 4" + " " + strButtonTip
        If row = 0 And column = 8 Then strButtonTip = "Accent 5" + " " + strButtonTip
        If row = 0 And column = 9 Then strButtonTip = "Accent 6" + " " + strButtonTip
        '
        If row = 0 And column = 10 Then strButtonTip = "HyperLink" + " " + strButtonTip
        If row = 0 And column = 11 Then strButtonTip = "Followed HyperLink" + " " + strButtonTip
        '
        Return strButtonTip

    End Function
    '
    '
    ''' <summary>
    ''' This method will return an RGB string of the form "R=xxx, G=xxx, B=xxx"
    ''' </summary>
    ''' <param name="colourRgb"></param>
    ''' <returns></returns>
    Public Function getRGB_longForm(colourRgb As Color) As String
        Dim strRed, strGreen, strBlue As String
        '
        strRed = colourRgb.R.ToString()
        strGreen = colourRgb.G.ToString()
        strBlue = colourRgb.B.ToString()
        '
        'getRGB_longForm = "R=" + strRed + ", " + "G=" + strGreen + ", " + "B=" + strBlue
        '
        getRGB_longForm = "RGB = " + strRed + "," + strGreen + "," + strBlue


    End Function
    '
    '
    ''' <summary>
    ''' This method will get the list of Seed Colours used to create the
    ''' Custom COours section of the ColorPicker as per excel
    ''' </summary>
    ''' <returns></returns>
    Public Function colr_set_SeedColours() As Collection
        Dim lstOfSeedColours As New Collection()
        '
        lstOfSeedColours.Add(RGB(255, 255, 255), "0")
        lstOfSeedColours.Add(RGB(20, 0, 52), "1")
        lstOfSeedColours.Add(RGB(108, 63, 153), "2")
        lstOfSeedColours.Add(RGB(157, 133, 190), "3")
        lstOfSeedColours.Add(RGB(204, 195, 220), "4")
        lstOfSeedColours.Add(RGB(230, 226, 238), "5")
        lstOfSeedColours.Add(RGB(123, 189, 214), "6")
        lstOfSeedColours.Add(RGB(0, 106, 159), "7")
        lstOfSeedColours.Add(RGB(159, 209, 139), "8")
        lstOfSeedColours.Add(RGB(66, 141, 82), "9")
        lstOfSeedColours.Add(RGB(56, 148, 137), "10")
        'lstOfSeedColours.Add(RGB(20, 84, 188), "11")
        'lstOfSeedColours.Add(RGB(50, 16, 192), "12")
        'lstOfSeedColours.Add(RGB(102, 83, 243), "13")
        'lstOfSeedColours.Add(RGB(93, 46, 162), "14")
        'lstOfSeedColours.Add(RGB(178, 149, 249), "15")
        'lstOfSeedColours.Add(RGB(204, 152, 246), "16")
        'lstOfSeedColours.Add(RGB(60, 146, 148), "17")
        'lstOfSeedColours.Add(RGB(52, 156, 136), "18")

        '    
        Return lstOfSeedColours
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will get the list of Seed Colours used to create the
    ''' Custom COours section of the ColorPicker as per excel
    ''' </summary>
    ''' <returns></returns>
    Public Function colr_set_SeedColours_10x5() As Collection
        Dim lstOfSeedColours As New Collection()
        '
        lstOfSeedColours.Add(RGB(255, 255, 255), "0")
        lstOfSeedColours.Add(RGB(20, 0, 52), "1")
        lstOfSeedColours.Add(RGB(108, 63, 153), "2")
        lstOfSeedColours.Add(RGB(157, 133, 190), "3")
        lstOfSeedColours.Add(RGB(204, 195, 220), "4")
        lstOfSeedColours.Add(RGB(230, 226, 238), "5")
        lstOfSeedColours.Add(RGB(123, 189, 214), "6")
        lstOfSeedColours.Add(RGB(0, 106, 159), "7")
        lstOfSeedColours.Add(RGB(159, 209, 139), "8")
        lstOfSeedColours.Add(RGB(66, 141, 82), "9")
        'lstOfSeedColours.Add(RGB(56, 148, 137), "10")
        'lstOfSeedColours.Add(RGB(20, 84, 188), "11")
        'lstOfSeedColours.Add(RGB(50, 16, 192), "12")
        'lstOfSeedColours.Add(RGB(102, 83, 243), "13")
        'lstOfSeedColours.Add(RGB(93, 46, 162), "14")
        'lstOfSeedColours.Add(RGB(178, 149, 249), "15")
        'lstOfSeedColours.Add(RGB(204, 152, 246), "16")
        'lstOfSeedColours.Add(RGB(60, 146, 148), "17")
        'lstOfSeedColours.Add(RGB(52, 156, 136), "18")
        '    
        Return lstOfSeedColours
        '
    End Function

    '
    '
    ''' <summary>
    ''' This method will get the list of Seed Colours used to create the
    ''' Custom COours section of the ColorPicker as per excel
    ''' </summary>
    ''' <returns></returns>
    Public Function colr_set_SeedColours_02() As Collection
        Dim lstOfSeedColours As New Collection()
        '
        lstOfSeedColours.Add(RGB(255, 255, 255), "0")
        lstOfSeedColours.Add(RGB(20, 0, 52), "1")
        lstOfSeedColours.Add(RGB(108, 63, 153), "2")
        lstOfSeedColours.Add(RGB(157, 133, 190), "3")
        lstOfSeedColours.Add(RGB(204, 195, 220), "4")
        lstOfSeedColours.Add(RGB(230, 226, 238), "5")
        lstOfSeedColours.Add(RGB(123, 189, 214), "6")
        lstOfSeedColours.Add(RGB(0, 106, 159), "7")
        lstOfSeedColours.Add(RGB(159, 209, 139), "8")
        lstOfSeedColours.Add(RGB(66, 141, 82), "9")
        lstOfSeedColours.Add(RGB(56, 148, 137), "10")
        lstOfSeedColours.Add(RGB(20, 84, 188), "11")
        lstOfSeedColours.Add(RGB(50, 16, 192), "12")
        lstOfSeedColours.Add(RGB(102, 83, 243), "13")
        lstOfSeedColours.Add(RGB(93, 46, 162), "14")
        lstOfSeedColours.Add(RGB(178, 149, 249), "15")
        lstOfSeedColours.Add(RGB(204, 152, 246), "16")
        lstOfSeedColours.Add(RGB(60, 146, 148), "17")
        lstOfSeedColours.Add(RGB(52, 156, 136), "18")
        '    
        Return lstOfSeedColours
        '
    End Function
    '
    Public Function colr_get_SeedColours(ByRef lstOfCustomColours As Collection) As Collection
        Dim numROws, numColumns As Integer
        Dim lst As Collection
        Dim lstOfSeedColours As New Collection()
        Dim j As Integer
        '
        lst = lstOfCustomColours("0")
        numColumns = lstOfCustomColours.Count
        numROws = lst.Count
        '
        For j = 0 To numColumns - 1
            lst = lstOfCustomColours(CStr(j))
            lstOfSeedColors.Add(lst.Item(CStr(lst.Count - 1)), CStr(j))
            '
            'Now remove the seed colour from the bottom of the column
            lst.Remove(CStr(lst.Count - 1))
        Next
        '
        Return lstOfSeedColors
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method takes in the collection of custom palette buttons and returns an
    ''' XML file that can be inserted into the Theme File to Add a Custom Colors sectio
    ''' </summary>
    Public Function colr_build_CustClrsXML() As String
        Dim btn As ToolStripButton
        Dim lstOfButtons As List(Of ToolStripButton)
        Dim j, row As Integer
        Dim strXMLCustClrs As String
        '
        '
        '
        strXMLCustClrs = "<a:custClrLst>" + vbCrLf
        '
        'MsgBox("NumRows = " + Me.numRows.ToString() + "NumClumns = " + Me.numColumns.ToString)
        'Seems that we can only display 5 rows in Custom Colours
        'For row = 1 to 10 Step 2
        For row = 1 To 5 Step 1
            lstOfButtons = getRow(row, 5)

            For j = 0 To lstOfButtons.Count - 1
                btn = lstOfButtons.Item(j)
                Me.getRGB_Hex(btn.BackColor)
                '
                strXMLCustClrs = strXMLCustClrs + "<a:custClr name=" + """" + btn.ToolTipText + """" + ">" + vbCrLf
                strXMLCustClrs = strXMLCustClrs + "<a:srgbClr val=" + """" + Me.getRGB_Hex(btn.BackColor) + """" + " />" + vbCrLf
                strXMLCustClrs = strXMLCustClrs + "</a:custClr>" + vbCrLf
                '
            Next
        Next row
        '
        strXMLCustClrs = strXMLCustClrs + "</a:custClrLst>"
        '
        Return strXMLCustClrs

    End Function
    '
    Public Function colr_build_CustClrsVBNET(ByRef lstOfCustomColors As Collection, Optional doSeedColoursAtEnd As Boolean = False) As String
        Dim j, k, rgbColour As Integer
        Dim lst As Collection
        Dim numColumns, numRows As Integer
        Dim strDim, strRGB, strLine, strResult As String
        Dim objCol As cColorObj
        '

        lst = lstOfCustomColors("0")
        strDim = ""
        strLine = ""
        strResult = ""
        '
        'Me.colr_modify_CustomColour(RGB(255, 0, 0), 0, 0, lstOfCustomColors)

        '
        numColumns = lstOfCustomColors.Count
        numRows = lst.Count
        '
        'Setup the dimension statements
        '
        For j = 0 To lstOfCustomColors.Count - 1
            strDim = strDim + "Dim" + " lst" + CStr(j) + " As New Collection()" + vbCrLf
        Next
        '
        strDim = vbCrLf + strDim + "Dim" + " lstofCustomColours" + " As New Collection()" + vbCrLf

        '
        For j = 0 To lstOfCustomColors.Count - 1
            'Get each column, then add the elements to each column.. If doSeed is true, then we add the
            'seed colour to the end
            '
            lst = lstOfCustomColors.Item(CStr(j))
            'Elements of each column
            For k = 0 To lst.Count - 1
                rgbColour = lst.Item(CStr(k))
                objCol = New cColorObj(rgbColour)
                strRGB = "RGB(" + objCol.strColor + ")"
                '
                strLine = strLine + "lst" + CStr(j) + ".Add(" + strRGB + ", " + """" + CStr(k) + """" + ")" + vbCrLf
            Next
            '
            If doSeedColoursAtEnd Then
                rgbColour = Me.lstOfSeedColors.Item(CStr(j))
                objCol = New cColorObj(rgbColour)
                strRGB = "RGB(" + objCol.strColor + ")"
                '
                strLine = strLine + "lst" + CStr(j) + ".Add(" + strRGB + ", " + """" + CStr(lst.Count) + """" + ")" + vbCrLf
                '
            End If
            '
            strLine = strLine + "lstofCustomColours.Add(" + "lst" + CStr(j) + ", " + """" + CStr(j) + """" + ")" + vbCrLf + vbCrLf
        Next
        '
        strResult = strDim + vbCrLf + strLine + vbCrLf + vbCrLf + "Return lstofCustomColours"
        '
        Return strResult
        '
    End Function
    '
    Public Function getRow(rowNum As Integer, totalRows As Integer) As List(Of ToolStripButton)
        Dim btn As ToolStripButton
        'Dim item00, item01, item02, item03 As ToolStripButton
        'Dim item10, item11, item12, item13 As ToolStripButton

        Dim lstOfRowItems As New List(Of ToolStripButton)
        Dim columnNum, btnNumber As Integer
        Dim strColor As String
        '
        'start of row n = n * numColumns, the row has numColumns items
        '
        'row 0
        'item00 = Me.lst_of_CustomColourButtons(CStr(0))
        'item01 = Me.lst_of_CustomColourButtons(CStr(0 + totalRows))
        'item02 = Me.lst_of_CustomColourButtons(CStr(0 + 2 * totalRows))
        'item03 = Me.lst_of_CustomColourButtons(CStr(0 + 3 * totalRows))
        '
        '
        'row 1
        'item10 = Me.lst_of_CustomColourButtons(CStr(1))
        'item11 = Me.lst_of_CustomColourButtons(CStr(1 + totalRows))
        'item12 = Me.lst_of_CustomColourButtons(CStr(1 + 2 * totalRows))
        'item13 = Me.lst_of_CustomColourButtons(CStr(1 + 3 * totalRows))
        '
        For columnNum = 0 To numColumns - 1
            'btn = Me.lst_of_CustomColourButtons(CStr(rowNum + columnNum * totalRows))
            btnNumber = rowNum + (columnNum * totalRows) - 1
            '
            btn = Me.lst_of_CustomColourButtons.Item(btnNumber)
            strColor = Me.getRGB_longForm(btn.BackColor)

            lstOfRowItems.Add(btn)
        Next
        '
        Return lstOfRowItems
    End Function
    '
    '
    ''' <summary>
    ''' This method will return the rgb colour as a Hex String (RRGGBB)
    ''' </summary>
    ''' <param name="colourRgb"></param>
    ''' <returns></returns>
    Public Function getRGB_Hex(colourRgb As Color) As String
        Dim strRslt As String
        '
        strRslt = ""
        strRslt = colourRgb.R.ToString("X2") + colourRgb.G.ToString("X2") + colourRgb.B.ToString("X2")
        '
        Return strRslt
    End Function
    '
    Public Function getRGB_Int(colourRgb As Color) As Integer
        '
        getRGB_Int = RGB(colourRgb.R, colourRgb.G, colourRgb.B)
        '
    End Function
    '
    '
    ''' <summary>
    ''' This version of custom colors was written by colr_build_CustClrsVBNET from a set of
    ''' colours automatically generated from a seed... The seed colours have been placed at the end of
    ''' each column. The idea being that when building the palette from custom colours written as
    ''' a list, the seed colour row is taken from the last item
    ''' </summary>
    ''' <returns></returns>
    Public Function colr_set_CustomColours_AA_02() As Collection
        'This version of custom colors was written by colr_build_CustClrsVBNET

        Dim lst0 As New Collection()
        Dim lst1 As New Collection()
        Dim lst2 As New Collection()
        Dim lst3 As New Collection()
        Dim lst4 As New Collection()
        Dim lst5 As New Collection()
        Dim lst6 As New Collection()
        Dim lst7 As New Collection()
        Dim lst8 As New Collection()
        Dim lst9 As New Collection()
        Dim lst10 As New Collection()
        Dim lst11 As New Collection()
        Dim lst12 As New Collection()
        Dim lst13 As New Collection()
        Dim lst14 As New Collection()
        Dim lst15 As New Collection()
        Dim lst16 As New Collection()
        Dim lst17 As New Collection()
        Dim lst18 As New Collection()
        Dim lstofCustomColours As New Collection()

        lst0.Add(RGB(255, 255, 255), "0")
        lst0.Add(RGB(249, 249, 249), "1")
        lst0.Add(RGB(229, 229, 229), "2")
        lst0.Add(RGB(200, 200, 200), "3")
        lst0.Add(RGB(125, 125, 125), "4")
        lst0.Add(RGB(77, 77, 77), "5")
        lst0.Add(RGB(29, 29, 29), "6")
        lst0.Add(RGB(0, 0, 0), "7")
        lst0.Add(RGB(255, 255, 255), "8")
        lstofCustomColours.Add(lst0, "0")

        lst1.Add(RGB(230, 226, 238), "0")
        lst1.Add(RGB(204, 195, 220), "1")
        lst1.Add(RGB(157, 133, 190), "2")
        lst1.Add(RGB(161, 102, 255), "3")
        lst1.Add(RGB(129, 51, 255), "4")
        lst1.Add(RGB(108, 63, 153), "5")
        lst1.Add(RGB(51, 16, 99), "6")
        lst1.Add(RGB(34, 11, 65), "7")
        lst1.Add(RGB(20, 0, 52), "8")
        lstofCustomColours.Add(lst1, "1")

        lst2.Add(RGB(230, 219, 240), "0")
        lst2.Add(RGB(204, 183, 225), "1")
        lst2.Add(RGB(178, 147, 210), "2")
        lst2.Add(RGB(153, 110, 196), "3")
        lst2.Add(RGB(128, 74, 181), "4")
        lst2.Add(RGB(102, 59, 145), "5")
        lst2.Add(RGB(76, 45, 108), "6")
        lst2.Add(RGB(51, 30, 72), "7")
        lst2.Add(RGB(108, 63, 153), "8")
        lstofCustomColours.Add(lst2, "2")

        lst3.Add(RGB(228, 222, 237), "0")
        lst3.Add(RGB(202, 188, 220), "1")
        lst3.Add(RGB(175, 155, 202), "2")
        lst3.Add(RGB(148, 122, 184), "3")
        lst3.Add(RGB(121, 89, 166), "4")
        lst3.Add(RGB(97, 71, 133), "5")
        lst3.Add(RGB(73, 53, 100), "6")
        lst3.Add(RGB(49, 35, 67), "7")
        lst3.Add(RGB(157, 133, 190), "8")
        lstofCustomColours.Add(lst3, "3")

        lst4.Add(RGB(228, 223, 236), "0")
        lst4.Add(RGB(200, 191, 217), "1")
        lst4.Add(RGB(173, 158, 199), "2")
        lst4.Add(RGB(145, 126, 180), "3")
        lst4.Add(RGB(118, 94, 161), "4")
        lst4.Add(RGB(94, 75, 129), "5")
        lst4.Add(RGB(71, 56, 97), "6")
        lst4.Add(RGB(47, 38, 64), "7")
        lst4.Add(RGB(204, 195, 220), "8")
        lstofCustomColours.Add(lst4, "4")

        lst5.Add(RGB(227, 223, 226), "0")
        lst5.Add(RGB(200, 191, 217), "1")
        lst5.Add(RGB(172, 159, 198), "2")
        lst5.Add(RGB(144, 126, 180), "3")
        lst5.Add(RGB(116, 94, 161), "4")
        lst5.Add(RGB(93, 75, 129), "5")
        lst5.Add(RGB(70, 57, 96), "6")
        lst5.Add(RGB(47, 38, 64), "7")
        lst5.Add(RGB(230, 226, 238), "8")
        lstofCustomColours.Add(lst5, "5")

        lst6.Add(RGB(233, 242, 247), "0")
        lst6.Add(RGB(196, 221, 233), "1")
        lst6.Add(RGB(138, 197, 219), "2")
        lst6.Add(RGB(99, 177, 207), "3")
        lst6.Add(RGB(60, 158, 195), "4")
        lst6.Add(RGB(0, 106, 159), "5")
        lst6.Add(RGB(0, 72, 110), "6")
        lst6.Add(RGB(0, 32, 50), "7")
        lst6.Add(RGB(123, 189, 214), "8")
        lstofCustomColours.Add(lst6, "6")

        lst7.Add(RGB(204, 238, 255), "0")
        lst7.Add(RGB(153, 221, 255), "1")
        lst7.Add(RGB(102, 204, 255), "2")
        lst7.Add(RGB(51, 187, 255), "3")
        lst7.Add(RGB(0, 170, 255), "4")
        lst7.Add(RGB(0, 136, 204), "5")
        lst7.Add(RGB(0, 102, 153), "6")
        lst7.Add(RGB(0, 68, 102), "7")
        lst7.Add(RGB(0, 106, 159), "8")
        lstofCustomColours.Add(lst7, "7")

        lst8.Add(RGB(241, 247, 237), "0")
        lst8.Add(RGB(212, 231, 200), "1")
        lst8.Add(RGB(159, 209, 139), "2")
        lst8.Add(RGB(134, 197, 109), "3")
        lst8.Add(RGB(83, 146, 58), "4")
        lst8.Add(RGB(66, 141, 82), "5")
        lst8.Add(RGB(0, 64, 22), "6")
        lst8.Add(RGB(16, 37, 21), "7")
        lst8.Add(RGB(159, 209, 139), "8")
        lstofCustomColours.Add(lst8, "8")

        lst9.Add(RGB(220, 239, 224), "0")
        lst9.Add(RGB(186, 222, 193), "1")
        lst9.Add(RGB(151, 206, 163), "2")
        lst9.Add(RGB(116, 190, 132), "3")
        lst9.Add(RGB(81, 174, 101), "4")
        lst9.Add(RGB(65, 139, 81), "5")
        lst9.Add(RGB(49, 104, 61), "6")
        lst9.Add(RGB(33, 69, 40), "7")
        lst9.Add(RGB(66, 141, 82), "8")
        lstofCustomColours.Add(lst9, "9")

        lst10.Add(RGB(218, 241, 238), "0")
        lst10.Add(RGB(181, 227, 221), "1")
        lst10.Add(RGB(144, 213, 205), "2")
        lst10.Add(RGB(107, 199, 188), "3")
        lst10.Add(RGB(70, 185, 171), "4")
        lst10.Add(RGB(56, 148, 137), "5")
        lst10.Add(RGB(42, 111, 103), "6")
        lst10.Add(RGB(28, 74, 68), "7")
        lst10.Add(RGB(56, 148, 137), "8")
        lstofCustomColours.Add(lst10, "10")

        lst11.Add(RGB(209, 225, 250), "0")
        lst11.Add(RGB(163, 194, 245), "1")
        lst11.Add(RGB(117, 164, 240), "2")
        lst11.Add(RGB(71, 133, 235), "3")
        lst11.Add(RGB(25, 103, 230), "4")
        lst11.Add(RGB(20, 82, 184), "5")
        lst11.Add(RGB(15, 62, 138), "6")
        lst11.Add(RGB(10, 41, 92), "7")
        lst11.Add(RGB(20, 84, 188), "8")
        lstofCustomColours.Add(lst11, "11")

        lst12.Add(RGB(216, 208, 251), "0")
        lst12.Add(RGB(178, 161, 247), "1")
        lst12.Add(RGB(139, 114, 243), "2")
        lst12.Add(RGB(100, 67, 239), "3")
        lst12.Add(RGB(61, 20, 235), "4")
        lst12.Add(RGB(49, 16, 188), "5")
        lst12.Add(RGB(37, 12, 141), "6")
        lst12.Add(RGB(25, 8, 94), "7")
        lst12.Add(RGB(50, 16, 192), "8")
        lstofCustomColours.Add(lst12, "12")

        lst13.Add(RGB(213, 207, 252), "0")
        lst13.Add(RGB(170, 160, 248), "1")
        lst13.Add(RGB(128, 112, 245), "2")
        lst13.Add(RGB(85, 64, 242), "3")
        lst13.Add(RGB(43, 17, 238), "4")
        lst13.Add(RGB(34, 13, 191), "5")
        lst13.Add(RGB(26, 10, 143), "6")
        lst13.Add(RGB(17, 7, 95), "7")
        lst13.Add(RGB(102, 83, 243), "8")
        lstofCustomColours.Add(lst13, "13")

        lst14.Add(RGB(227, 215, 244), "0")
        lst14.Add(RGB(199, 176, 232), "1")
        lst14.Add(RGB(170, 136, 221), "2")
        lst14.Add(RGB(142, 96, 210), "3")
        lst14.Add(RGB(114, 56, 199), "4")
        lst14.Add(RGB(91, 45, 159), "5")
        lst14.Add(RGB(68, 34, 119), "6")
        lst14.Add(RGB(46, 23, 79), "7")
        lst14.Add(RGB(93, 46, 162), "8")
        lstofCustomColours.Add(lst14, "14")

        lst15.Add(RGB(220, 207, 252), "0")
        lst15.Add(RGB(185, 158, 250), "1")
        lst15.Add(RGB(150, 110, 247), "2")
        lst15.Add(RGB(115, 62, 244), "3")
        lst15.Add(RGB(80, 14, 241), "4")
        lst15.Add(RGB(64, 11, 193), "5")
        lst15.Add(RGB(48, 8, 145), "6")
        lst15.Add(RGB(32, 5, 97), "7")
        lst15.Add(RGB(178, 149, 249), "8")
        lstofCustomColours.Add(lst15, "15")

        lst16.Add(RGB(232, 208, 251), "0")
        lst16.Add(RGB(209, 161, 247), "1")
        lst16.Add(RGB(185, 114, 243), "2")
        lst16.Add(RGB(162, 67, 239), "3")
        lst16.Add(RGB(139, 20, 235), "4")
        lst16.Add(RGB(111, 16, 188), "5")
        lst16.Add(RGB(83, 12, 141), "6")
        lst16.Add(RGB(56, 8, 94), "7")
        lst16.Add(RGB(204, 152, 246), "8")
        lstofCustomColours.Add(lst16, "16")

        lst17.Add(RGB(219, 240, 240), "0")
        lst17.Add(RGB(182, 225, 226), "1")
        lst17.Add(RGB(146, 209, 211), "2")
        lst17.Add(RGB(110, 194, 196), "3")
        lst17.Add(RGB(74, 179, 181), "4")
        lst17.Add(RGB(59, 143, 145), "5")
        lst17.Add(RGB(44, 107, 109), "6")
        lst17.Add(RGB(29, 72, 73), "7")
        lst17.Add(RGB(60, 146, 148), "8")
        lstofCustomColours.Add(lst17, "17")

        lst18.Add(RGB(255, 200, 200), "0")
        lst18.Add(RGB(255, 175, 175), "1")
        lst18.Add(RGB(255, 150, 150), "2")
        lst18.Add(RGB(255, 125, 125), "3")
        lst18.Add(RGB(255, 100, 100), "4")
        lst18.Add(RGB(255, 75, 75), "5")
        lst18.Add(RGB(255, 50, 50), "6")
        lst18.Add(RGB(255, 25, 25), "7")
        lst18.Add(RGB(255, 0, 0), "8")
        lstofCustomColours.Add(lst18, "18")

        'lst18.Add(RGB(217, 242, 237), "0")
        'lst18.Add(RGB(178, 230, 220), "1")
        'lst18.Add(RGB(140, 217, 202), "2")
        'lst18.Add(RGB(102, 204, 184), "3")
        'lst18.Add(RGB(64, 191, 167), "4")
        'lst18.Add(RGB(51, 153, 133), "5")
        'lst18.Add(RGB(38, 115, 100), "6")
        'lst18.Add(RGB(25, 77, 67), "7")
        'lst18.Add(RGB(52, 156, 136), "8")
        'lst18.Add(RGB(255, 0, 0), "8")
        'lstofCustomColours.Add(lst18, "18")



        Return lstofCustomColours
        '
        Return lstofCustomColours

    End Function
    '
    '
    ''' <summary>
    ''' This version of custom colors was written by colr_build_CustClrsVBNET from a set of
    ''' colours automatically generated from a seed... The seed colours have been placed at the end of
    ''' each column. The idea being that when building the palette from custom colours written as
    ''' a list, the seed colour row is taken from the last item
    ''' </summary>
    ''' <returns></returns>
    Public Function colr_set_CustomColours_AA_03() As Collection
        'This version of custom colors was written by colr_build_CustClrsVBNET
        Dim lst0 As New Collection()
        Dim lst1 As New Collection()
        Dim lst2 As New Collection()
        Dim lst3 As New Collection()
        Dim lst4 As New Collection()
        Dim lst5 As New Collection()
        Dim lst6 As New Collection()
        Dim lst7 As New Collection()
        Dim lst8 As New Collection()
        Dim lst9 As New Collection()
        Dim lst10 As New Collection()
        Dim lst11 As New Collection()
        Dim lst12 As New Collection()
        Dim lst13 As New Collection()
        Dim lst14 As New Collection()
        Dim lst15 As New Collection()
        Dim lst16 As New Collection()
        Dim lst17 As New Collection()
        Dim lst18 As New Collection()
        Dim lstofCustomColours As New Collection()

        lst0.Add(RGB(255, 255, 255), "0")
        lst0.Add(RGB(249, 249, 249), "1")
        lst0.Add(RGB(229, 229, 229), "2")
        lst0.Add(RGB(216, 216, 216), "3")
        lst0.Add(RGB(125, 125, 125), "4")
        lst0.Add(RGB(77, 77, 77), "5")
        lst0.Add(RGB(29, 29, 29), "6")
        lst0.Add(RGB(0, 0, 0), "7")
        lst0.Add(RGB(255, 255, 255), "8")
        lstofCustomColours.Add(lst0, "0")

        lst1.Add(RGB(231, 216, 255), "0")
        lst1.Add(RGB(207, 177, 255), "1")
        lst1.Add(RGB(183, 137, 255), "2")
        lst1.Add(RGB(158, 98, 255), "3")
        lst1.Add(RGB(75, 0, 196), "4")
        lst1.Add(RGB(60, 0, 157), "5")
        lst1.Add(RGB(45, 0, 118), "6")
        lst1.Add(RGB(30, 0, 78), "7")
        lst1.Add(RGB(20, 0, 52), "8")
        lstofCustomColours.Add(lst1, "1")

        lst2.Add(RGB(235, 227, 244), "0")
        lst2.Add(RGB(216, 199, 232), "1")
        lst2.Add(RGB(196, 172, 221), "2")
        lst2.Add(RGB(177, 144, 209), "3")
        lst2.Add(RGB(98, 57, 139), "4")
        lst2.Add(RGB(78, 46, 111), "5")
        lst2.Add(RGB(59, 34, 83), "6")
        lst2.Add(RGB(39, 23, 56), "7")
        lst2.Add(RGB(108, 63, 153), "8")
        lstofCustomColours.Add(lst2, "2")

        lst3.Add(RGB(234, 229, 241), "0")
        lst3.Add(RGB(214, 204, 228), "1")
        lst3.Add(RGB(193, 178, 214), "2")
        lst3.Add(RGB(173, 153, 200), "3")
        lst3.Add(RGB(93, 68, 128), "4")
        lst3.Add(RGB(75, 55, 102), "5")
        lst3.Add(RGB(56, 41, 77), "6")
        lst3.Add(RGB(37, 27, 51), "7")
        lst3.Add(RGB(157, 133, 190), "8")
        lstofCustomColours.Add(lst3, "3")

        lst4.Add(RGB(234, 230, 241), "0")
        lst4.Add(RGB(213, 205, 226), "1")
        lst4.Add(RGB(192, 181, 212), "2")
        lst4.Add(RGB(171, 156, 197), "3")
        lst4.Add(RGB(91, 72, 124), "4")
        lst4.Add(RGB(73, 58, 99), "5")
        lst4.Add(RGB(55, 43, 74), "6")
        lst4.Add(RGB(36, 29, 50), "7")
        lst4.Add(RGB(204, 195, 220), "8")
        lstofCustomColours.Add(lst4, "4")

        lst5.Add(RGB(241, 230, 240), "0")
        lst5.Add(RGB(226, 205, 225), "1")
        lst5.Add(RGB(212, 181, 210), "2")
        lst5.Add(RGB(197, 156, 194), "3")
        lst5.Add(RGB(124, 72, 120), "4")
        lst5.Add(RGB(99, 58, 96), "5")
        lst5.Add(RGB(74, 43, 72), "6")
        lst5.Add(RGB(50, 29, 48), "7")
        lst5.Add(RGB(51, 16, 99), "8")
        lstofCustomColours.Add(lst5, "5")

        lst6.Add(RGB(225, 240, 246), "0")
        lst6.Add(RGB(195, 225, 236), "1")
        lst6.Add(RGB(165, 210, 227), "2")
        lst6.Add(RGB(135, 195, 218), "3")
        lst6.Add(RGB(46, 121, 150), "4")
        lst6.Add(RGB(37, 97, 120), "5")
        lst6.Add(RGB(28, 73, 90), "6")
        lst6.Add(RGB(19, 49, 60), "7")
        lst6.Add(RGB(123, 189, 214), "8")
        lstofCustomColours.Add(lst6, "6")

        lst7.Add(RGB(216, 242, 255), "0")
        lst7.Add(RGB(177, 229, 255), "1")
        lst7.Add(RGB(137, 216, 255), "2")
        lst7.Add(RGB(98, 203, 255), "3")
        lst7.Add(RGB(0, 131, 196), "4")
        lst7.Add(RGB(0, 105, 157), "5")
        lst7.Add(RGB(0, 78, 118), "6")
        lst7.Add(RGB(0, 52, 78), "7")
        lst7.Add(RGB(0, 106, 159), "8")
        lstofCustomColours.Add(lst7, "7")

        lst8.Add(RGB(232, 244, 227), "0")
        lst8.Add(RGB(209, 233, 199), "1")
        lst8.Add(RGB(185, 222, 171), "2")
        lst8.Add(RGB(162, 210, 143), "3")
        lst8.Add(RGB(80, 140, 56), "4")
        lst8.Add(RGB(64, 112, 45), "5")
        lst8.Add(RGB(48, 84, 33), "6")
        lst8.Add(RGB(32, 56, 22), "7")
        lst8.Add(RGB(159, 209, 139), "8")
        lstofCustomColours.Add(lst8, "8")

        lst9.Add(RGB(228, 242, 231), "0")
        lst9.Add(RGB(202, 230, 208), "1")
        lst9.Add(RGB(175, 217, 184), "2")
        lst9.Add(RGB(148, 205, 160), "3")
        lst9.Add(RGB(63, 134, 78), "4")
        lst9.Add(RGB(50, 107, 62), "5")
        lst9.Add(RGB(38, 80, 47), "6")
        lst9.Add(RGB(25, 53, 31), "7")
        lst9.Add(RGB(66, 141, 82), "8")
        lstofCustomColours.Add(lst9, "9")

        lst10.Add(RGB(227, 244, 242), "0")
        lst10.Add(RGB(198, 233, 229), "1")
        lst10.Add(RGB(170, 223, 216), "2")
        lst10.Add(RGB(141, 212, 203), "3")
        lst10.Add(RGB(54, 142, 132), "4")
        lst10.Add(RGB(43, 114, 105), "5")
        lst10.Add(RGB(32, 85, 79), "6")
        lst10.Add(RGB(22, 57, 53), "7")
        lst10.Add(RGB(56, 148, 137), "8")
        lstofCustomColours.Add(lst10, "10")

        lst11.Add(RGB(220, 232, 251), "0")
        lst11.Add(RGB(184, 208, 247), "1")
        lst11.Add(RGB(149, 185, 244), "2")
        lst11.Add(RGB(113, 161, 240), "3")
        lst11.Add(RGB(19, 79, 177), "4")
        lst11.Add(RGB(15, 63, 142), "5")
        lst11.Add(RGB(11, 48, 106), "6")
        lst11.Add(RGB(8, 32, 71), "7")
        lst11.Add(RGB(20, 84, 188), "8")
        lstofCustomColours.Add(lst11, "11")

        lst12.Add(RGB(225, 219, 252), "0")
        lst12.Add(RGB(195, 183, 249), "1")
        lst12.Add(RGB(166, 146, 246), "2")
        lst12.Add(RGB(136, 110, 243), "3")
        lst12.Add(RGB(47, 15, 181), "4")
        lst12.Add(RGB(38, 12, 145), "5")
        lst12.Add(RGB(28, 9, 109), "6")
        lst12.Add(RGB(19, 6, 72), "7")
        lst12.Add(RGB(50, 16, 192), "8")
        lstofCustomColours.Add(lst12, "12")
        '
        GoTo finis
        '
        lst13.Add(RGB(222, 218, 252), "0")
        lst13.Add(RGB(190, 182, 250), "1")
        lst13.Add(RGB(157, 145, 247), "2")
        lst13.Add(RGB(125, 108, 245), "3")
        lst13.Add(RGB(33, 13, 183), "4")
        lst13.Add(RGB(26, 10, 147), "5")
        lst13.Add(RGB(20, 8, 110), "6")
        lst13.Add(RGB(13, 5, 73), "7")
        lst13.Add(RGB(102, 83, 243), "8")
        lstofCustomColours.Add(lst13, "13")

        lst14.Add(RGB(233, 224, 246), "0")
        lst14.Add(RGB(212, 194, 238), "1")
        lst14.Add(RGB(190, 163, 229), "2")
        lst14.Add(RGB(168, 133, 220), "3")
        lst14.Add(RGB(88, 43, 153), "4")
        lst14.Add(RGB(70, 35, 122), "5")
        lst14.Add(RGB(53, 26, 92), "6")
        lst14.Add(RGB(35, 17, 61), "7")
        lst14.Add(RGB(93, 46, 162), "8")
        lstofCustomColours.Add(lst14, "14")

        lst15.Add(RGB(228, 218, 253), "0")
        lst15.Add(RGB(201, 181, 251), "1")
        lst15.Add(RGB(174, 144, 249), "2")
        lst15.Add(RGB(147, 106, 247), "3")
        lst15.Add(RGB(61, 11, 186), "4")
        lst15.Add(RGB(49, 8, 149), "5")
        lst15.Add(RGB(37, 6, 111), "6")
        lst15.Add(RGB(25, 4, 74), "7")
        lst15.Add(RGB(178, 149, 249), "8")
        lstofCustomColours.Add(lst15, "15")

        lst16.Add(RGB(237, 219, 252), "0")
        lst16.Add(RGB(219, 183, 249), "1")
        lst16.Add(RGB(201, 147, 246), "2")
        lst16.Add(RGB(184, 111, 242), "3")
        lst16.Add(RGB(107, 16, 180), "4")
        lst16.Add(RGB(85, 13, 144), "5")
        lst16.Add(RGB(64, 9, 108), "6")
        lst16.Add(RGB(43, 6, 72), "7")
        lst16.Add(RGB(204, 152, 246), "8")
        lstofCustomColours.Add(lst16, "16")

        lst17.Add(RGB(227, 243, 244), "0")
        lst17.Add(RGB(199, 232, 232), "1")
        lst17.Add(RGB(171, 220, 221), "2")
        lst17.Add(RGB(143, 208, 210), "3")
        lst17.Add(RGB(57, 138, 140), "4")
        lst17.Add(RGB(45, 110, 112), "5")
        lst17.Add(RGB(34, 83, 84), "6")
        lst17.Add(RGB(23, 55, 56), "7")
        lst17.Add(RGB(60, 146, 148), "8")
        lstofCustomColours.Add(lst17, "17")

        lst18.Add(RGB(226, 245, 241), "0")
        lst18.Add(RGB(196, 235, 228), "1")
        lst18.Add(RGB(167, 226, 214), "2")
        lst18.Add(RGB(137, 216, 201), "3")
        lst18.Add(RGB(49, 147, 128), "4")
        lst18.Add(RGB(39, 118, 103), "5")
        lst18.Add(RGB(29, 88, 77), "6")
        lst18.Add(RGB(20, 59, 51), "7")
        lst18.Add(RGB(52, 156, 136), "8")
        lstofCustomColours.Add(lst18, "18")
        '
finis:
        Return lstofCustomColours

    End Function
    '
    '
    ''' <summary>
    ''' This version of custom colors was written by colr_build_CustClrsVBNET from a set of
    ''' colours automatically generated from a seed... The seed colours have been placed at the end of
    ''' each column. The idea being that when building the palette from custom colours written as
    ''' a list, the seed colour row is taken from the last item
    ''' </summary>
    ''' <returns></returns>
    Public Function colr_set_CustomColours_AA_10x5_20250926() As Collection
        'This version of custom colors was written by colr_build_CustClrsVBNET

        Dim lst0 As New Collection()
        Dim lst1 As New Collection()
        Dim lst2 As New Collection()
        Dim lst3 As New Collection()
        Dim lst4 As New Collection()
        Dim lst5 As New Collection()
        Dim lst6 As New Collection()
        Dim lst7 As New Collection()
        Dim lst8 As New Collection()
        Dim lst9 As New Collection()
        'Dim lst10 As New Collection()
        'Dim lst11 As New Collection()
        'Dim lst12 As New Collection()
        'Dim lst13 As New Collection()
        'Dim lst14 As New Collection()
        'Dim lst15 As New Collection()
        'Dim lst16 As New Collection()
        'Dim lst17 As New Collection()
        'Dim lst18 As New Collection()
        Dim lstofCustomColours As New Collection()
        '

        lst0.Add(RGB(77, 77, 77), "0")              'second row
        lst0.Add(RGB(125, 125, 125), "1")           'third row
        lst0.Add(RGB(200, 200, 200), "2")           'fourth row
        lst0.Add(RGB(229, 229, 229), "3")           'fifth row
        lst0.Add(RGB(0, 0, 0), "4")                 'Top row primary colour
        'lst0.Add(RGB(93, 75, 129), "5")
        'lst0.Add(RGB(70, 57, 96), "6")
        'lst0.Add(RGB(47, 38, 64), "7")
        'lst0.Add(RGB(0, 0, 0), "8")
        lstofCustomColours.Add(lst0, "0")

        lst1.Add(RGB(51, 16, 99), "0")
        lst1.Add(RGB(108, 63, 152), "1")
        lst1.Add(RGB(157, 133, 190), "2")
        lst1.Add(RGB(204, 195, 220), "3")
        lst1.Add(RGB(20, 0, 52), "4")
        'lst1.Add(RGB(77, 77, 77), "5")
        'lst1.Add(RGB(29, 29, 29), "6")
        'lst1.Add(RGB(0, 0, 0), "7")
        'lst1.Add(RGB(255, 255, 255), "8")
        lstofCustomColours.Add(lst1, "1")

        lst2.Add(RGB(0, 106, 159), "0")
        lst2.Add(RGB(123, 189, 214), "1")
        lst2.Add(RGB(185, 209, 229), "2")
        lst2.Add(RGB(228, 238, 244), "3")
        lst2.Add(RGB(0, 72, 110), "4")
        'lst2.Add(RGB(108, 63, 153), "5")
        'lst2.Add(RGB(51, 16, 99), "6")
        'lst2.Add(RGB(34, 11, 65), "7")
        'lst2.Add(RGB(20, 0, 52), "8")
        lstofCustomColours.Add(lst2, "2")

        lst3.Add(RGB(66, 141, 82), "0")
        lst3.Add(RGB(132, 206, 136), "1")
        lst3.Add(RGB(187, 226, 176), "2")
        lst3.Add(RGB(231, 243, 225), "3")
        lst3.Add(RGB(0, 64, 22), "4")
        'lst3.Add(RGB(102, 59, 145), "5")
        'lst3.Add(RGB(76, 45, 108), "6")
        'lst3.Add(RGB(51, 30, 72), "7")
        'lst3.Add(RGB(108, 63, 153), "8")
        lstofCustomColours.Add(lst3, "3")

        lst4.Add(RGB(255, 255, 255), "0")
        lst4.Add(RGB(255, 255, 255), "1")
        lst4.Add(RGB(255, 255, 255), "2")
        lst4.Add(RGB(255, 255, 255), "3")
        lst4.Add(RGB(255, 255, 255), "4")
        'lst4.Add(RGB(97, 71, 133), "5")
        'lst4.Add(RGB(73, 53, 100), "6")
        'lst4.Add(RGB(49, 35, 67), "7")
        'lst4.Add(RGB(157, 133, 190), "8")
        lstofCustomColours.Add(lst4, "4")

        lst5.Add(RGB(255, 255, 255), "0")
        lst5.Add(RGB(255, 255, 255), "1")
        lst5.Add(RGB(255, 255, 255), "2")
        lst5.Add(RGB(255, 255, 255), "3")
        lst5.Add(RGB(255, 255, 255), "4")
        'lst5.Add(RGB(94, 75, 129), "5")
        'lst5.Add(RGB(71, 56, 97), "6")
        'lst5.Add(RGB(47, 38, 64), "7")
        'lst5.Add(RGB(204, 195, 220), "8")
        lstofCustomColours.Add(lst5, "5")

        lst6.Add(RGB(255, 255, 255), "0")
        lst6.Add(RGB(255, 255, 255), "1")
        lst6.Add(RGB(255, 255, 255), "2")
        lst6.Add(RGB(255, 255, 255), "3")
        lst6.Add(RGB(255, 255, 255), "4")
        'lst6.Add(RGB(0, 106, 159), "5")
        'lst6.Add(RGB(0, 72, 110), "6")
        'lst6.Add(RGB(0, 32, 50), "7")
        'lst6.Add(RGB(123, 189, 214), "8")
        lstofCustomColours.Add(lst6, "6")

        lst7.Add(RGB(174, 97, 34), "0")
        lst7.Add(RGB(213, 119, 42), "1")
        lst7.Add(RGB(221, 146, 85), "2")
        lst7.Add(RGB(232, 181, 140), "3")
        lst7.Add(RGB(95, 53, 19), "4")
        'lst7.Add(RGB(0, 136, 204), "5")
        'lst7.Add(RGB(0, 102, 153), "6")
        'lst7.Add(RGB(0, 68, 102), "7")
        'lst7.Add(RGB(0, 106, 159), "8")
        lstofCustomColours.Add(lst7, "7")

        lst8.Add(RGB(34, 108, 108), "0")
        lst8.Add(RGB(63, 153, 153), "1")
        lst8.Add(RGB(112, 196, 196), "2")
        lst8.Add(RGB(196, 231, 231), "3")
        lst8.Add(RGB(22, 55, 55), "4")
        'lst8.Add(RGB(66, 141, 82), "5")
        'lst8.Add(RGB(0, 64, 22), "6")
        'lst8.Add(RGB(16, 37, 21), "7")
        'lst8.Add(RGB(159, 209, 139), "8")
        lstofCustomColours.Add(lst8, "8")

        lst9.Add(RGB(171, 30, 124), "0")
        lst9.Add(RGB(217, 38, 157), "1")
        lst9.Add(RGB(234, 134, 200), "2")
        lst9.Add(RGB(247, 212, 235), "3")
        lst9.Add(RGB(106, 19, 77), "4")
        'lst9.Add(RGB(65, 139, 81), "5")
        'lst9.Add(RGB(49, 104, 61), "6")
        'lst9.Add(RGB(33, 69, 40), "7")
        'lst9.Add(RGB(66, 141, 82), "8")
        lstofCustomColours.Add(lst9, "9")


        '
        Return lstofCustomColours

    End Function
    '
    '
    ''' <summary>
    ''' This version of custom colors was written by colr_build_CustClrsVBNET from a set of
    ''' colours automatically generated from a seed... The seed colours have been placed at the end of
    ''' each column. The idea being that when building the palette from custom colours written as
    ''' a list, the seed colour row is taken from the last item
    ''' </summary>
    ''' <returns></returns>
    Public Function colr_set_CustomColours_AA_7x6_20250926() As Collection
        'This version of custom colors was written by colr_build_CustClrsVBNET

        Dim lst0 As New Collection()
        Dim lst1 As New Collection()
        Dim lst2 As New Collection()
        Dim lst3 As New Collection()
        Dim lst4 As New Collection()
        Dim lst5 As New Collection()
        Dim lst6 As New Collection()
        Dim lst7 As New Collection()
        Dim lst8 As New Collection()
        Dim lst9 As New Collection()
        'Dim lst10 As New Collection()
        'Dim lst11 As New Collection()
        'Dim lst12 As New Collection()
        'Dim lst13 As New Collection()
        'Dim lst14 As New Collection()
        'Dim lst15 As New Collection()
        'Dim lst16 As New Collection()
        'Dim lst17 As New Collection()
        'Dim lst18 As New Collection()
        Dim lstofCustomColours As New Collection()
        '

        lst0.Add(RGB(0, 0, 0), "0")
        lst0.Add(RGB(77, 77, 77), "1")
        lst0.Add(RGB(125, 125, 125), "2")
        lst0.Add(RGB(200, 200, 200), "3")
        lst0.Add(RGB(229, 229, 229), "4")
        lst0.Add(RGB(125, 125, 125), "5")
        'lst0.Add(RGB(125, 125, 125), "6")
        'lst0.Add(RGB(229, 229, 229), "7")
        'lst0.Add(RGB(229, 229, 229), "8")
        lstofCustomColours.Add(lst0, "0")

        lst1.Add(RGB(20, 0, 52), "0")
        lst1.Add(RGB(51, 16, 99), "1")
        lst1.Add(RGB(108, 63, 152), "2")
        lst1.Add(RGB(157, 133, 190), "3")
        lst1.Add(RGB(204, 195, 220), "4")
        lst1.Add(RGB(108, 63, 152), "5")
        'lst1.Add(RGB(29, 29, 29), "6")
        'lst1.Add(RGB(0, 0, 0), "7")
        'lst1.Add(RGB(255, 255, 255), "8")
        lstofCustomColours.Add(lst1, "1")

        lst2.Add(RGB(0, 72, 110), "0")
        lst2.Add(RGB(0, 106, 159), "1")
        lst2.Add(RGB(123, 189, 214), "2")
        lst2.Add(RGB(185, 209, 229), "3")
        lst2.Add(RGB(228, 238, 244), "4")
        lst2.Add(RGB(123, 189, 214), "5")
        'lst2.Add(RGB(51, 16, 99), "6")
        'lst2.Add(RGB(34, 11, 65), "7")
        'lst2.Add(RGB(20, 0, 52), "8")
        lstofCustomColours.Add(lst2, "2")

        lst3.Add(RGB(0, 64, 22), "0")
        lst3.Add(RGB(66, 141, 82), "1")
        lst3.Add(RGB(132, 206, 136), "2")
        lst3.Add(RGB(187, 226, 176), "3")
        lst3.Add(RGB(231, 243, 225), "4")
        lst3.Add(RGB(132, 206, 136), "5")
        'lst3.Add(RGB(76, 45, 108), "6")
        'lst3.Add(RGB(51, 30, 72), "7")
        'lst3.Add(RGB(108, 63, 153), "8")
        lstofCustomColours.Add(lst3, "3")

        'lst4.Add(RGB(255, 255, 255), "0")
        'lst4.Add(RGB(255, 255, 255), "1")
        'lst4.Add(RGB(255, 255, 255), "2")
        'lst4.Add(RGB(255, 255, 255), "3")
        'lst4.Add(RGB(255, 255, 255), "4")
        'lst4.Add(RGB(255, 255, 255), "5")
        'lst4.Add(RGB(73, 53, 100), "6")
        'lst4.Add(RGB(49, 35, 67), "7")
        'lst4.Add(RGB(157, 133, 190), "8")
        'lstofCustomColours.Add(lst4, "4")

        'lst5.Add(RGB(255, 255, 255), "0")
        'lst5.Add(RGB(255, 255, 255), "1")
        'lst5.Add(RGB(255, 255, 255), "2")
        'lst5.Add(RGB(255, 255, 255), "3")
        'lst5.Add(RGB(255, 255, 255), "4")
        'lst5.Add(RGB(255, 255, 255), "5")
        'lst5.Add(RGB(71, 56, 97), "6")
        'lst5.Add(RGB(47, 38, 64), "7")
        'lst5.Add(RGB(204, 195, 220), "8")
        'lstofCustomColours.Add(lst5, "5")

        'lst6.Add(RGB(255, 255, 255), "0")
        'lst6.Add(RGB(255, 255, 255), "1")
        'lst6.Add(RGB(255, 255, 255), "2")
        'lst6.Add(RGB(255, 255, 255), "3")
        'lst6.Add(RGB(255, 255, 255), "4")
        'lst6.Add(RGB(255, 255, 255), "5")
        'lst6.Add(RGB(71, 56, 97), "6")
        'lst6.Add(RGB(47, 38, 64), "7")
        'lst6.Add(RGB(204, 195, 220), "8")
        'lstofCustomColours.Add(lst6, "6")

        lst4.Add(RGB(95, 53, 19), "0")
        lst4.Add(RGB(174, 97, 34), "1")
        lst4.Add(RGB(213, 119, 42), "2")
        lst4.Add(RGB(221, 146, 85), "3")
        lst4.Add(RGB(232, 181, 140), "4")
        lst4.Add(RGB(213, 119, 42), "5")
        'lst4.Add(RGB(0, 102, 153), "6")
        'lst4.Add(RGB(0, 68, 102), "7")
        'lst4.Add(RGB(0, 106, 159), "8")
        lstofCustomColours.Add(lst4, "4")

        lst5.Add(RGB(22, 55, 55), "0")
        lst5.Add(RGB(34, 108, 108), "1")
        lst5.Add(RGB(63, 153, 153), "2")
        lst5.Add(RGB(112, 196, 196), "3")
        lst5.Add(RGB(196, 231, 231), "4")
        lst5.Add(RGB(63, 153, 153), "5")
        'lst8.Add(RGB(0, 64, 22), "6")
        'lst8.Add(RGB(16, 37, 21), "7")
        'lst8.Add(RGB(159, 209, 139), "8")
        lstofCustomColours.Add(lst5, "5")

        lst6.Add(RGB(106, 19, 77), "0")
        lst6.Add(RGB(171, 30, 124), "1")
        lst6.Add(RGB(217, 38, 157), "2")
        lst6.Add(RGB(234, 134, 200), "3")
        lst6.Add(RGB(247, 212, 235), "4")
        lst6.Add(RGB(217, 38, 157), "5")
        'lst9.Add(RGB(49, 104, 61), "6")
        'lst9.Add(RGB(33, 69, 40), "7")
        'lst9.Add(RGB(66, 141, 82), "8")
        lstofCustomColours.Add(lst6, "6")


        '
        Return lstofCustomColours

    End Function
    '
    '
    Public Function frmPicker_get_CustomColours(ByRef lstOfSeedColours As Collection) As Collection
        Dim numSteps As Integer
        Dim objCol As cColorObj
        Dim lstOfColours As New Collection()
        'Dim lstOfSeedColours As New Collection()
        Dim j As Integer
        Dim objGlobals As New cGlobals()
        'Dim myCol As Color
        '
        numSteps = 11
        numSteps = 15
        numSteps = 8

        'lstOfSeedColours = objGlobals.getSeedColours()
        '
        'For each seed color get a Collection of Colours, which are variations of the seed
        'This goes from column 0 to column lstOfSeedColours.Count - 1
        '
        For j = 0 To lstOfSeedColours.Count - 1
            objCol = New cColorObj(CInt(lstOfSeedColours(CStr(j))))
            lstOfColours.Add(objCol.getColoursFromSeed(True, numSteps, objCol.myColour), CStr(j))
        Next
        '
        'myCol = New Color()

        'myCol = lstOfColours.Item(CStr(lstOfSeedColours.Count - 1))
        'myCol.R = 255
        '
        Return lstOfColours
        '
    End Function

    Private Sub rdBtn_Theme2024_CheckedChanged(sender As Object, e As EventArgs)

    End Sub

    Private Sub btn_changeToAATheme_Click(sender As Object, e As EventArgs) Handles btn_changeToAATheme.Click
        Dim objThmMgr As New cThemeMgr()
        Dim objGlobals As New cGlobals()
        Dim myDoc As Word.Document
        Dim strSwitch As String
        '
        strSwitch = "rdBtn_Theme2025"
        myDoc = objGlobals.glb_get_wrdActiveDoc()
        '
        Select Case Me.cmBox_themesToChoose.SelectedItem.ToString()
            Case "Most current"
                If objThmMgr.thm_Set_ThemeToAAStd_fromFile(myDoc) Then
                    Me.frm_refresh_thmPalette()
                    'Me.objColorMgr.refreshThemePalette()
                    'btn.BackColor = Me.objGlobals.btn_On_Color
                    'objFormatMgr.frmt_change_ExcelFont(wb)
                Else
                    MsgBox("You may be missing a theme file",, "Missing Theme File")
                End If
                '
            Case "Most current (manual)"
                objThmMgr.thm_Set_ThemeToAAStd_Manually(myDoc)
                Me.frm_refresh_thmPalette()
                '
            Case "2024 theme from file"
                'If objThmMgr.thm_Set_ThemeToAAStd_20250926_fromFile(myDoc, True) Then
                'Me.frm_refresh_thmPalette()
                'Me.objColorMgr.refreshThemePalette()
                'btn.BackColor = Me.objGlobals.btn_On_Color
                'objFormatMgr.frmt_change_ExcelFont(wb)
                'Else
                MsgBox("You may be missing a theme file",, "Missing Theme File")
                'End If
                '
            Case "2024 theme manual"
                objThmMgr.thm_Set_ThemeToAAStd_20240808_Manually(myDoc)
                Me.frm_refresh_thmPalette()
                '
            Case "Legacy light theme"
                objThmMgr.thm_Set_ThemeToAALightOrange_Manually(myDoc)
                Me.frm_refresh_thmPalette()
        End Select
        '
        '
finis:

    End Sub

    Private Sub frm_colorPicker02_Activated(sender As Object, e As EventArgs) Handles Me.Activated

    End Sub

    Private Sub rdBtn_colourCells_Click(sender As Object, e As EventArgs) Handles rdBtn_colourCells.Click, rdBtn_Grid.Click, rdBtn_colourText.Click
        Dim rdBtn As System.Windows.Forms.RadioButton = TryCast(sender, System.Windows.Forms.RadioButton)
        '
        If IsNothing(rdBtn) Then GoTo finis
        '
        Select Case rdBtn.Name
            Case "rdBtn_colourCells"
                Me.frm_colorPicker_Rename("tbl_Cells")
            Case "rdBtn_colourText"
                Me.frm_colorPicker_Rename("text_Colour")
            Case "rdBtn_Grid"
                Me.frm_colorPicker_Rename("tbl_cellBorders")
        End Select
        '
finis:

    End Sub
    '
    '
End Class
Imports System.Windows.Forms
Imports System.Drawing
Imports System.IO
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Tools.Ribbon
Public Class frm_colorPicker
    '
    Public objColorMgr As cColorMgr
    Public strFormMode As String            'text_Colour, seriesFill, seriesBorder
    'Public rbn As rbnPMTools
    Public objGlobals As cGlobals
    Public objWorkAround As New cWorkArounds()
    Public rgbColor_Selected As Long
    '
    Private _colorMatrix(8, 12) As Integer

    Public Sub New(strFormMode As String)
        Dim btn As ToolStripButton
        Dim btn_getClrsXML As Button
        Dim strCustClrsXML As String

        Dim i As Integer
        Dim paletteButtons As Collection
        Dim titleBarHeight As Integer
        Dim topOffSet, bottomOffSet As Integer
        '
        ' This call is required by the designer.
        InitializeComponent()
        '
        'Dim rbns As ThisRibbonCollection
        '
        'rbns = Globals.Ribbons
        strCustClrsXML = ""

        'Me.rbn = rbns.rbnPMTools
        'strLineWeight = rbn.cmBox_lineWeight.Text
        'strBorderWeight = rbn.cmBox_borderWeight.Text

        '
        Me.strFormMode = strFormMode
        '
        Me.clrPicker_name_Form(Me.strFormMode)
        '
        Me.btnSpin_BorderWidth.SelectedIndex = 0
        '
        'topOffSet = 30
        topOffSet = 30
        bottomOffSet = 40

        Me.objColorMgr = New cColorMgr()
        paletteButtons = objColorMgr.buildColorPalette(New System.Drawing.Point(12, topOffSet), "Theme Colour Palette (locked)", Me)
        'Now we must wire up the palette buttons
        For i = 1 To paletteButtons.Count
            btn = paletteButtons.Item(i)
            AddHandler btn.MouseDown, AddressOf Me.btnHandler_MouseDown
            AddHandler btn.MouseHover, AddressOf Me.btnHandler_Hover
        Next
        '
        paletteButtons = objColorMgr.buildColorPalette_Custom(New System.Drawing.Point(260, topOffSet), "Custom Colours Palette (RGB)", Me)
        'Now we must wire up the palette buttons
        For i = 1 To paletteButtons.Count
            btn = paletteButtons.Item(i)
            AddHandler btn.MouseDown, AddressOf Me.btnHandler_MouseDown
            AddHandler btn.MouseHover, AddressOf Me.btnHandler_Hover
        Next
        '
        '
        Me.btn_changeThemeForThisWorkBook.Left = 12
        Me.btn_changeThemeForThisWorkBook.Top = topOffSet + objColorMgr.extraPaletteHeight - Me.btn_changeThemeForThisWorkBook.Height - 3
        Me.btn_Close.Top = Me.btn_changeThemeForThisWorkBook.Top
        '
        titleBarHeight = RectangleToScreen(Me.ClientRectangle).Top - Me.Top
        'Me.Height = titleBarHeight + topOffSet + objColorMgr.extraPaletteHeight + 15
        Me.Height = titleBarHeight + topOffSet + objColorMgr.extraPaletteHeight + bottomOffSet

        Me.FormBorderStyle = FormBorderStyle.FixedDialog
        Me.MinimizeBox = False
        Me.MaximizeBox = False
        '
        Me.objGlobals = New cGlobals()
        '
        btn_getClrsXML = New Button()
        '
        btn_getClrsXML.AutoSize = False
        btn_getClrsXML.Font = New Drawing.Font("SansSerif", 8.25)
        btn_getClrsXML.Height = 24
        btn_getClrsXML.Width = objColorMgr.extraPaletteWidth
        btn_getClrsXML.Margin = New Padding(0, 0, 0, 2)
        btn_getClrsXML.BackColor = Color.White
        btn_getClrsXML.Visible = False
        btn_getClrsXML.Text = "Get CustClrsXML as File"
        btn_getClrsXML.Name = "btn_getClrsXML"
        '
        'MsgBox("Top is " + objColorMgr.strip.Top.ToString())
        btn_getClrsXML.Location = New System.Drawing.Point(Me.grpBox_Marker2.Left, (objColorMgr.strip.Top + objColorMgr.strip.Height + 0))
        'btn_getClrsXML.Top = titleBarHeight + topOffSet + objColorMgr.extraPaletteHeight + 20
        'btn_getClrsXML.Left = 60
        '
        Me.btn_getColours.Visible = False
        '
        Select Case Me.strFormMode
            Case "testMode"
                'My XML writer
                AddHandler btn_getClrsXML.MouseDown, AddressOf Me.btn_getClrsXML_Handler
                Me.btn_getColours.Visible = True
                btn_getClrsXML.Visible = True
                '
            Case "tbl_Header_Colour_all"
                AddHandler btn_getClrsXML.MouseDown, AddressOf Me.btn_fillHeaders_Handler
                '
            Case Else
                'My XML writer
                AddHandler btn_getClrsXML.MouseDown, AddressOf Me.btn_getClrsXML_Handler
                '
        End Select

        Me.Controls.Add(btn_getClrsXML)
        '
        Me.Width = Me.grpBox_Marker2.Left + objColorMgr.extraPaletteWidth + 40
        '

        '
        'Me.myDoc = Globals.ThisAddIn.Application.ActiveDocument
        'For Each ctrl In Me.Controls
        'If TypeOf (ctrl) Is ToolStrip Then
        'k = 1
        'End If
        'Next ctrl
        '
        '
        'Me.CreateColorMatrix()
        '
        '
        'Me.cColourObj_Test()
        '
        '*** Series fill
        'Globals.ThisAddIn.Application.Selection.
        'Me.toolStripTest()
        ' Add any initialization after the InitializeComponent() call.
        'test()
    End Sub
    '
    ''' <summary>
    ''' This method will set the name of the form depending on which mode it
    ''' was initialised in (i.e. according to strFormMode
    ''' </summary>
    ''' <param name="strFormMode"></param>
    Public Sub clrPicker_name_Form(strFormMode As String)
        '
        Select Case strFormMode
            Case "testMode"
                Me.Text = "Test Mode"
            Case "seriesFill"
                Me.Text = "Colour Fill Mode"
            Case "seriesBorder"
                Me.Text = "Border Colour Mode"
            Case "text_Colour"
                Me.Text = "Text Colour Mode"
            Case "backPanel"
                Me.Text = "Image Back Panel fill Mode"
            Case "tbl_Cells"
                Me.Text = "Table Cell(s) fill Mode"
                Me.btn_noColour.Visible = True
            Case "tbl_Header_Colour_all"
                Me.Text = "Fill all Table Header Rows"
        End Select
        '
    End Sub
    '
    Private Function C(red As Integer, green As Integer, blue As Integer) As Color
        C = Color.FromArgb(red, green, blue)
    End Function
    '
    Public Function getRGB(colourRGB As Integer) As String
        '
        getRGB = CStr(Me.getARGB_R(colourRGB)) & "," & CStr(Me.getARGB_G(colourRGB)) & "," & CStr(Me.getARGB_B(colourRGB))
        '
    End Function

    '
    Public Function getARGB_A(colourARGB As Integer) As Integer
        Dim b As UInt32
        Dim msk As UInt32
        Dim tmp As UInt32
        '
        msk = &HFF000000UI
        b = CUInt(colourARGB)
        tmp = b And msk
        tmp = tmp / (2 ^ 24)
        '
        getARGB_A = CInt(tmp)
        'getARGB_A = CStr(tmp)
    End Function
    '
    Public Function getARGB_B(colourARGB As Integer) As Integer
        Dim b As UInt32
        Dim msk As UInt32
        Dim tmp As UInt32
        '
        msk = &HFF0000UI
        b = CUInt(colourARGB)
        tmp = b And msk
        tmp = tmp / (2 ^ 16)
        '
        getARGB_B = CInt(tmp)
        'getARGB_B = CStr(tmp)
    End Function
    '
    Public Function getARGB_G(colourARGB As Integer) As Integer
        Dim b As UInt32
        Dim msk As UInt32
        Dim tmp As UInt32
        '
        msk = &HFF00UI
        b = CUInt(colourARGB)
        tmp = b And msk
        tmp = tmp / (2 ^ 8)
        '
        getARGB_G = CInt(tmp)
        'getARGB_G = CStr(tmp)
    End Function
    '
    Public Function getARGB_R(colourARGB As Integer) As Integer
        Dim b As UInt32
        Dim msk As UInt32
        Dim tmp As UInt32
        '
        msk = &HFF
        b = CUInt(colourARGB)
        tmp = b And msk
        '
        getARGB_R = CInt(tmp)
        'getARGB_R = CStr(tmp)
    End Function


    Public Function getARGB(alpha As Integer, colourRGB As UInt32) As Integer
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
        getARGB = localColor.ToArgb()
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
        Dim objTblsMgr As New cTablesMgr()
        Dim btn As Button
        '
        btn = sender
        '
        Try
            'MsgBox("Into handler")
            btn.Enabled = False
            objTblsMgr.tbl_colour_set_HeaderRow(Me.rgbColor_Selected, True)
        Catch ex As Exception

        End Try
        '
        btn.Enabled = True
        MsgBox("Standard/Regular table headers fill complete")
        'System.Windows.Forms.FileDialog
        '
        'MsgBox(strResult)
    End Sub
    '
    '
    Public Sub btn_getClrsXML_Handler(sender As Object, e As MouseEventArgs)
        Dim strResult As String
        'Dim dlg As DialogResult
        Dim dlg_SaveFile As New SaveFileDialog()
        Dim strFilePath As String
        Dim myStream As StreamWriter
        '
        'strDocuments = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        '
        strFilePath = ""
        '
        strResult = Me.objColorMgr.colr_build_CustClrsXML()
        '
        dlg_SaveFile.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        dlg_SaveFile.FilterIndex = 1
        dlg_SaveFile.RestoreDirectory = True
        dlg_SaveFile.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        '
        Try
            If dlg_SaveFile.ShowDialog() = DialogResult.OK Then
                strFilePath = dlg_SaveFile.FileName
                myStream = New StreamWriter(strFilePath, False)
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
        Dim objColMgr As New cColorMgr()
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

                    'lstOfShapes = objBckPanelMgr.pnl_getBackPanel_PlaceHolders(objGlobals.glb_get_wrdSect)
                    'If lstOfShapes.Count > 0 Then
            'cShp = lstOfShapes.Item(0)
            'objBckPanelMgr.rgbFill = RGB(btnColor.R, btnColor.G, btnColor.B)
            'objBckPanelMgr.pnl_reset_BackPanelColour(cShp)
            'Else

        'End If
                Case "tbl_Cells"
                    Try
                        objTblsMgr.tbl_colour_set_colourOfCells(rgbColor)
                    Catch ex As Exception
                        MsgBox("Have you selected some table cells to colour?")
                    End Try
                Case "tbl_Header_Colour_all"
                    'Me.btn_doAllTblHeaders.Visible = True
                    'lstOfControls = Me.Controls.Find("btn_getClrsXML", True)
                    Me.Controls.Item("btn_getClrsXML").Text = "Click to fill all Regular/Standard Table Header Rows"
                    '
                    Me.Controls.Item("btn_getClrsXML").Visible = True
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
        Dim btn As ToolStripButton
        Dim btnColor As Color
        Dim strMsg As String
        'Dim objChartMgr As cChartMgr
        'Dim objSeriesMgr As cSeriesMgr

        '
        strMsg = "To change the colour of a Chart Item (e.g. Chart Background, ChartBorder, Series Fill, Series Border, etc...). You'll need to select the chart item that you want to colour"
        '
        btn = sender
        btnColor = btn.BackColor
        'objChartMgr = New cChartMgr()
        'objSeriesMgr = New cSeriesMgr(objChartMgr)
        '
        'If Not objSeriesMgr.doColorofSelectedSeriesOrPoints(drpDown.SuperTip, drpDown.ScreenTip, CSng(Me.cmBox_lineWeight.Text), CSng(Me.cmBox_borderWeight.Text)) Then

        'End If

        '
        Select Case Me.strFormMode
            Case "seriesFill"
                'If the Spin control is visible, then we want to do the borders
                'If objChartMgr.doColour(btnColor, Me.btnSpin_BorderWidth.Visible, CStr(Me.btnSpin_BorderWidth.SelectedItem)) Then
                'All is OK
                ' Else
                'Let's look for selected cells and if there are any we'll fill them
                'with the selected colour
                'Try
                'For Each drCell In Globals.ThisAddIn.Application.Selection
                'drCell.interior.Color = btnColor
                'Next
                'Catch ex As Exception
                'MsgBox(strMsg)
                'End Try
                'End If

            Case "seriesBorder"
                Try
                    'objChartMgr.doBorderColor(btnColor)
                Catch ex As Exception

                End Try
        End Select
        '
        '
    End Sub

    Private Sub btn_changeThemeForThisWorkBook_Click(sender As Object, e As EventArgs) Handles btn_changeThemeForThisWorkBook.Click
        Dim btn As System.Windows.Forms.Button
        Dim objFormatMgr As cFormatMgr
        Dim strMsg As String
        '
        Me.objColorMgr.colr_set_ThemeColours(Me.objGlobals.glb_get_wrdActiveDoc(), "aacBase")

        'Globals.ThisAddIn.Application.CommandBars(CommandBarPopup())

        strMsg = "The Acil Allen Theme File is missing from 'C:\Templates\'." & vbCrLf & "You'll need to contact IT Support to resolve this"
        objFormatMgr = New cFormatMgr
        '
        btn = sender
        'btn.Text = "AA Theme is Active"
        'Me.standardThemeIsOn = True
        'btnIcon = My.Resources.ThemesIcon_AA()
        'btn.Image = btnIcon
        'Globals.ThisAddIn.Application.
        'If objFormatMgr.applyStandardTheme(wb) Then
        'Me.objColorMgr.refreshThemePalette()
        'btn.BackColor = objGlobals.btn_On_Color
        'Else
        'MsgBox(strMsg,, "Missing Theme File")
        'End If
        '
        objColorMgr.refreshThemePalette()

    End Sub

    Private Sub btn_Close_Click(sender As Object, e As EventArgs) Handles btn_Close.Click
        Me.objWorkAround.wrk_fix_forCursorRace()
        Me.Close()
    End Sub

    Private Sub DoSeriesOfChartBordersToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DoSeriesOfChartBordersToolStripMenuItem.Click
        Dim mnuItem As ToolStripMenuItem
        '
        mnuItem = sender
        If mnuItem.Checked Then
            mnuItem.CheckState = CheckState.Unchecked
            Me.btnSpin_BorderWidth.Visible = False
            Me.lbl_borderSize.Visible = False
        Else
            mnuItem.CheckState = CheckState.Checked
            Me.btnSpin_BorderWidth.Visible = True
            Me.lbl_borderSize.Visible = True
        End If
    End Sub

    Private Sub frm_colorPicker_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub btn_GetColorPicker_Click(sender As Object, e As EventArgs) Handles btn_GetColorPicker.Click
        Dim dlgs As Word.Dialogs
        Dim dlg As Word.Dialog

        dlgs = Globals.ThisAddIn.Application.Dialogs
        dlg = dlgs.Item(WdWordDialog.wdDialogBuildingBlockOrganizer)
        'dlg = dlgs.Item(XlBuiltInDialog.xl)

        '
        dlg.Show()
        'Globals.ThisAddIn.Application.Dialogs(XlBuiltInDialog.xlDialogColorPalette)
        'Dim myDialog As System.Windows.Forms.ColorDialog
        '
        'myDialog = New ColorDialog()
        'myDialog.AllowFullOpen = True
        'myDialog.ShowDialog()
        'ColorDialog MyDialog = New ColorDialog()

    End Sub

    Private Sub btn_getColours_Click(sender As Object, e As EventArgs) Handles btn_getColours.Click
        Dim strResult As String
        Dim dlg_SaveFile As New SaveFileDialog()
        Dim strFilePath As String
        Dim myStream As StreamWriter
        'Dim objColMgr As New cColorMgr()
        '
        'strDocuments = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        '
        strFilePath = ""
        '
        'strResult = Me.objColorMgr.colr_build_CustClrsXML()
        strResult = Me.objColorMgr.colr_build_CustClrsVBNET(objColorMgr.lstOfCustomColors, True)
        '
        dlg_SaveFile.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*"
        dlg_SaveFile.FilterIndex = 1
        dlg_SaveFile.RestoreDirectory = True
        dlg_SaveFile.InitialDirectory = My.Computer.FileSystem.SpecialDirectories.MyDocuments
        '
        Try
            If dlg_SaveFile.ShowDialog() = DialogResult.OK Then
                strFilePath = dlg_SaveFile.FileName
                myStream = New StreamWriter(strFilePath, False)
                myStream.Write(strResult)
                myStream.Close()
                MsgBox("Custom Colours XML file successfully written")
            Else
                MsgBox("Action cancelled by the user")
            End If
        Catch ex As Exception
            MsgBox("Failed to write Custom Colours XML")
        End Try
        'System.Windows.Forms.FileDialog
        '
        'MsgBox(strResult)

    End Sub

    Private Sub btn_doAllTblHeaders_Click(sender As Object, e As EventArgs) Handles btn_doAllTblHeaders.Click
        Dim objWCAGMgr As New cWCAGMgr()
        Dim objTblsMgr As New cTablesMgr()
        '
        objTblsMgr.tbl_colour_set_HeaderRow(Me.rgbColor_Selected, True)
        'For Each tbl In Me.objGlobals.glb_get_wrdActiveDoc().Tables
        'Try
        'If objTblsMgr.tbl_is_tblStandard(tbl) And objTblsMgr.glb_tbls_isRegular(tbl) Then
        'dr = tbl.Rows.Item(1)
        'objTblsMgr.tbl_colour_set_colourOfRow(dr, Me.rgbColor_Selected)
        'Else
        'If objTblsMgr.tbl_is_tblGlossary(tbl) And objTblsMgr.glb_tbls_isRegular(tbl) Then
        'dr = tbl.Rows.Item(1)
        'objTblsMgr.tbl_colour_set_colourOfRow(dr, Me.rgbColor_Selected)
        'End If
        'End If
        'Catch ex As Exception

        'End Try
        '
        'Next
        '
        'For Each sect In Me.objGlobals.glb_get_wrdActiveDoc().Sections
        'For Each tbl In sect.Range.Tables
        'Try
        'If objTblsMgr.tbl_is_tblStandard(tbl) And objTblsMgr.glb_tbls_isRegular(tbl) Then
        'dr = tbl.Rows.Item(1)
        'objTblsMgr.tbl_colour_set_colourOfRow(dr, Me.rgbColor_Selected)
        'End If
        'Catch ex As Exception

        'End Try
        '
        'Next
        '
    End Sub

    Private Sub btn_noColour_Click(sender As Object, e As EventArgs) Handles btn_noColour.Click
        Dim objTblsMgr As New cTablesMgr()
        '
        Try
            objTblsMgr.tbl_colour_set_colourOfCellsToNone(objGlobals.glb_get_wrdSel.Cells())
        Catch ex As Exception

        End Try
        '
    End Sub


    '
End Class
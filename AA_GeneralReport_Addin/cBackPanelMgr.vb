
Imports System.Drawing
Imports System.Windows.Forms
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word
Imports Microsoft.VisualBasic.FileIO

Public Class cBackPanelMgr
    Inherits cGlobals

    Public parentImageMgr As cImageMgr
    Public strShapeName As String
    Public strBackPanel_CaseStudy
    Public rgbFill As Integer
    Public strScratchFilePath As String
    Public objTools As cTools
    '
    'Public objTools As cTools
    Public currentSect As Word.Section
    '
    Public strPicturePlaceHolderName As String
    Public strFormLayoutName As String
    '
    Public Sub New()
        MyBase.New()
        Me.strShapeName = "aac_BackColour"
        Me.strBackPanel_CaseStudy = "aac_BackColour_CaseStudy"
        '
        Me.rgbFill = Me._glb_colour_purple_Dark
        Me.strScratchFilePath = My.Computer.FileSystem.SpecialDirectories.MyPictures + "\aac_scratch_file.jpg"
        '
        Me.objTools = New cTools()
        Me.currentSect = Nothing
        Me.strPicturePlaceHolderName = "Logo_Pict_Background"
        Me.strFormLayoutName = "Acil Allen"
        '
        Me.parentImageMgr = Nothing
    End Sub
    '
    ''' <summary>
    ''' This method will return true if shp object has the same name as a standard
    ''' back panel
    ''' </summary>
    ''' <param name="shp"></param>
    ''' <returns></returns>
    Public Function pnl_Shp_IsBackPanel(ByRef shp As Word.Shape) As Boolean
        Dim rslt As Boolean
        '
        rslt = False
        If shp.Name = Me.strShapeName Then rslt = True
        'If shp.Name = Me.strShapeName Or shp.Name = Me.strBackPanel_CaseStudy Then rslt = True

        '
        Return rslt
    End Function
    '
    '
    ''' <summary>
    ''' This method takes as input a lstOfPanels in the form of cShapeMgr. Thnere is only ever one item in
    ''' this list. If this item is not a Shape (i.e. a pasted image) it is replaced with a Shape with all
    ''' of the same characteristics
    ''' </summary>
    ''' <param name="lstOfpanels"></param>
    ''' <returns></returns>
    Public Function pnl_Image_ReplaceWithShape(ByRef lstOfpanels As List(Of cShapeMgr)) As List(Of cShapeMgr)
        Dim objShpMgr As New cShapeMgr()
        Dim objShpMgr_New As New cShapeMgr()
        Dim lstOfBackPanels_FullPage As List(Of cShapeMgr)
        Dim height, width, top, left As Single
        Dim newShape As Word.Shape
        Dim rng As Word.Range
        Dim strNewShapeName As String
        Dim hf As Word.HeaderFooter
        Dim sect As Word.Section
        Dim rgbFill As Integer
        Dim objGlobals As New cGlobals()
        '
        strNewShapeName = strShapeName
        sect = Nothing
        rgbFill = objGlobals._glb_colour_purple_Dark
        '
        '
        Try
            If lstOfpanels.Count > 0 Then
                objShpMgr = lstOfpanels.Item(0)
                sect = objShpMgr.hf.Range.Sections.Item(1)
                '
                If Not objShpMgr.shp.Type = MsoShapeType.msoAutoShape Then
                    strNewShapeName = objShpMgr.shp.Name
                    rng = objShpMgr.anchor
                    width = objShpMgr.width
                    height = objShpMgr.height
                    top = objShpMgr.top
                    left = objShpMgr.left
                    hf = objShpMgr.hf
                    '
                    objShpMgr.shp.Delete()
                    'newShape = Me.pnl_BackPanel_Insert(objShpMgr.hf, rng, RGB(255, 0, 0))
                    newShape = Me.pnl_BackPanel_Insert(objShpMgr.hf, rng, rgbFill)
                    newShape.ZOrder(MsoZOrderCmd.msoBringToFront)
                    newShape.Height = height
                    newShape.Width = width
                    newShape.Left = left
                    newShape.Top = top
                    newShape.LockAnchor = True
                    newShape.Name = strNewShapeName
                    '
                    objShpMgr_New = New cShapeMgr()
                    objShpMgr_New.InitShape(newShape, hf)
                    lstOfpanels.Clear()
                    lstOfpanels.Add(objShpMgr_New)
                    '
                    'Now get any full page back panels (purple) so that we can extract the colour
                    'and set the new shape to that colour, making it effectively invisible
                    lstOfBackPanels_FullPage = Me.pnl_getBackPanel_PlaceHolders(sect)
                    If lstOfBackPanels_FullPage.Count > 0 Then
                        objShpMgr = lstOfBackPanels_FullPage.Item(0)
                        newShape.Fill.ForeColor.RGB = rgbFill
                        newShape.Fill.BackColor.RGB = rgbFill
                    End If


                End If
                '
                '
            Else

            End If
            '
            '***** It works
            'img = Clipboard.GetImage()
            'img.Save(Me.strScratchFilePath)
            'newShape.Fill.UserPicture(Me.strScratchFilePath)

        Catch ex As Exception
            MsgBox("Error in " + "cBackPanelMgr.pnl_Image_ReplaceWithShape")
        End Try
        '
        Return lstOfpanels
    End Function

    '
    Public Sub pnl_fill_withUserImage(ByRef objShpMgr As cShapeMgr)
        Dim img As System.Drawing.Image
        'NewPic.Select()
        'sel = Globals.ThisAddin.Application.Selection
        'sel.CopyAsPicture()
        'NewPic.Delete()
        '
        '***** It works
        'img = Clipboard.GetImage()
        img = System.Windows.Forms.Clipboard.GetImage()
        img.Save(Me.strScratchFilePath)
        objShpMgr.shp.Fill.UserPicture(Me.strScratchFilePath)

    End Sub
    '
    '
    Public Sub pnl_fill_withRawUserImage(ByRef objShpMgr As cShapeMgr)
        Dim strFilePath As String
        Dim picDialog As Word.Dialog
        '
        strFilePath = ""
        picDialog = Globals.ThisAddIn.Application.Dialogs(Word.WdWordDialog.wdDialogInsertPicture)
        strFilePath = ""
        picDialog.name = "*.*"
        '
        Select Case picDialog.Display
            Case 0, -2                              'Close and Cancel button
            Case -1                             'OK button
                strFilePath = picDialog.Name
                If Not (My.Computer.FileSystem.FileExists(strFilePath)) Then
                    strFilePath = ""
                End If
        End Select
        '
        If Not strFilePath = "" Then
            objShpMgr.shp.Fill.UserPicture(strFilePath)
        End If

    End Sub
    '
    '
    Public Sub pnl_reset_BackPanelColour(ByRef sect As Word.Section, Optional strPanelName As String = "aac_BackColour", Optional rgbFill As Long = -1)
        Dim lstOfBackPanels As List(Of cShapeMgr)
        '
        lstOfBackPanels = Me.pnl_getBackPanel_PlaceHolders(sect, strPanelName)                     'To get rid of any existing back panels
        Me.pnl_reset_BackPanelColour(lstOfBackPanels, rgbFill)
        '
    End Sub
    '

    '
    Public Sub pnl_reset_BackPanelColour(ByRef lstOfBackPanels As List(Of cShapeMgr), Optional rgbFill As Long = -1)
        Dim j As Integer
        Dim shp As Word.Shape
        '
        If rgbFill = -1 Then rgbFill = Me.rgbFill
        '
        If lstOfBackPanels.Count <> 0 Then
            For j = lstOfBackPanels.Count - 1 To 0 Step -1
                shp = lstOfBackPanels.Item(j).shp
                shp.Fill.Solid()
                shp.Fill.Transparency = 0.0
                shp.Fill.ForeColor.RGB = rgbFill
                shp.Fill.BackColor.RGB = rgbFill
                '
            Next
        End If

    End Sub
    '
    ''' <summary>
    ''' This method will fill the shape (panel) referenced in  objShpMgr with the standard colour. You can vary the
    ''' colour by set the variable rgbFill from -1 to some other positive number
    ''' </summary>
    ''' <param name="objShpMgr"></param>
    Public Sub pnl_reset_BackPanelColour(ByRef objShpMgr As cShapeMgr, Optional rgbFill As Long = -1)
        '
        If rgbFill = -1 Then rgbFill = Me.rgbFill
        '
        If Me.pnl_Shp_IsBackPanel(objShpMgr.shp) Then
            objShpMgr.shp.Fill.Solid()
            objShpMgr.shp.Fill.Transparency = 0.0
            objShpMgr.shp.Fill.ForeColor.RGB = rgbFill
            objShpMgr.shp.Fill.BackColor.RGB = rgbFill
        End If
        '
    End Sub
    '
    ''' <summary>
    ''' This smthod will set the transparency of the back image panel (if it exists) to some vlaue (transparency)
    ''' between 0.0 (opaque) and 1.0 (clear)
    ''' 
    ''' </summary>
    ''' <param name="transparency"></param>
    ''' <param name="sect"></param>
    Public Sub pnl_reset_BackPanelTransparency(transparency As Single, ByRef sect As Word.Section)
        Dim lstOfPanels As New List(Of cShapeMgr)
        Dim objShpMgr As cShapeMgr
        '
        lstOfPanels = Me.pnl_getBackPanel_PlaceHolders(sect)
        '
        If lstOfPanels.Count <> 0 Then
            If transparency >= 1.0 Then transparency = 1.0
            If transparency <= 0.0 Then transparency = 0.0
            '
            objShpMgr = lstOfPanels.Item(0)
            objShpMgr.shp.Fill.Transparency = transparency
        End If
        '
        '
    End Sub
    '
    ''' <summary>
    ''' This method will return false if there is no image back panel in the section (sect).
    ''' If it returns true, then the variable transparency is set to the back panel transparency.
    ''' That is somewhere between 0.0 and 1.0.  Note that in some methods we may need this value
    ''' as a percentage. If so, make certain that you make the appropriate change. That is
    ''' transparency = Cint(transparency*100)
    ''' </summary>
    ''' <param name="transparency"></param>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function pnl_get_BackPanelTransparency(ByRef transparency As Single, ByRef sect As Word.Section) As Boolean
        Dim rslt As Boolean
        Dim lstOfPanels As New List(Of cShapeMgr)
        Dim objShpMgr As cShapeMgr
        '
        rslt = False
        '
        lstOfPanels = Me.pnl_getBackPanel_PlaceHolders(sect)
        '
        If lstOfPanels.Count <> 0 Then
            '
            rslt = True
            objShpMgr = lstOfPanels.Item(0)
            transparency = objShpMgr.shp.Fill.Transparency
            '
        Else
            transparency = 0.0
            rslt = False
        End If
        '
        Return rslt
        '
    End Function

    '
    Public Function pnl_has_BackPanel(ByRef sect As Word.Section, Optional strShpName As String = "aac_BackColour") As Boolean
        Dim lstOfPanels As New List(Of cShapeMgr)
        Dim rslt As Boolean
        '
        rslt = False
        lstOfPanels = Me.pnl_getBackPanel_PlaceHolders(sect, strShpName)
        If lstOfPanels.Count <> 0 Then rslt = True
        '
        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method returns a collection of the placeholders (typically back panel or small Cover Page
    ''' image placeholders... They are identified by name and returned as cShapeMgr types. The default
    ''' placeholder name is "aac_BackColour" becuase this is used mostly to pick out the purple
    ''' back panels
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function pnl_getBackPanel_PlaceHolders(ByRef sect As Word.Section, Optional strShapeName As String = "aac_BackColour") As List(Of cShapeMgr)
        'This method returns a collection of the picture placeholders
        'in the section.. They are isdentified by name adn returned as
        'cShapeMgr types.. The original shapes are deleted because we have all their
        'details
        '
        Dim hf As Word.HeaderFooter
        Dim objShape As cShapeMgr
        'Dim i As Int16
        Dim shp As Word.Shape
        'Dim secttest_00 As Word.Section
        Dim lstOfShapes As New List(Of cShapeMgr)
        'Dim strShapeName = "aac_BackColour"
        '
        'hf = Nothing
        'Me.currentSect = sect
        objShape = Nothing
        '
        For Each hf In sect.Headers
            If hf.Exists And hf.IsHeader Then
                If hf.Index = WdHeaderFooterIndex.wdHeaderFooterFirstPage Or hf.Index = WdHeaderFooterIndex.wdHeaderFooterPrimary Then
                    For Each shp In hf.Range.ShapeRange
                        If shp.Name = strShapeName Then
                            objShape = New cShapeMgr
                            objShape.InitShape(shp, hf)
                            lstOfShapes.Add(objShape)
                            '
                            'shp.Delete()
                            'GoTo loop00
                        End If
                    Next shp
                End If
            End If
        Next
        '
loop00:
        '
        Return lstOfShapes
    End Function
    '
    ''' <summary>
    ''' This method will insert a 'Brief' first page fractional backpanel into
    ''' the section sect
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function pnl_BackPanelBriefFirstPage_Insert(ByRef sect As Word.Section) As HeaderFooter
        Dim hf As Word.HeaderFooter
        Dim objRptBrief As New cReportBrief()
        Dim objBackPanelMgr As New cBackPanelMgr()
        Dim objBrandMgr As New cBrandMgr()
        Dim shp As Word.Shape
        '
        hf = Nothing
        Try
            If objRptBrief.brf_section_isFromBrief(sect) Then
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                objBackPanelMgr.pnl_BackPanel_Delete(sect)
                '
                shp = objBackPanelMgr.pnl_BackPanel_Insert(hf)
                shp.Height = sect.PageSetup.PageHeight * 0.2
                'shp.ZOrder(MsoZOrderCmd.msoBringInFrontOfText)
                shp.ZOrder(MsoZOrderCmd.msoSendBehindText)
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                objBrandMgr.brnd_recolour_Logo(hf)
            End If
        Catch ex As Exception
            hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        End Try
        '
        Return hf
        '
    End Function
    '
    ''' <summary>
    ''' This method will insert a coloured rectangle into the Shapes collection of the HeaderFooter hf
    ''' and it will lock it to the range rng. Typically the range will be a collapsed version of the
    ''' the hf range.. If you specifiy an RGB colour, then the rectange fill will be set to that colour.
    ''' If you leave that parameter empty, the colour will default to RGB(20, 0, 52). If you specify a strPanelName, then
    ''' the shape will be given that name
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="rgbFill"></param>
    ''' <returns></returns>
    Public Function pnl_BackPanel_Insert(ByRef hf As Word.HeaderFooter, Optional rgbFill As Integer = -1, Optional strPanelName As String = "") As Word.Shape
        Dim shp As Word.Shape
        'Dim strPanelName = "aac_BackColour"
        Dim sect As Word.Section
        Dim objWCAGMgr As New cWCAGMgr()
        Dim rngInsert As Word.Range
        '
        sect = hf.Range.Sections.Item(1)
        shp = Nothing
        '
        'If there are tables in the header footer, then we wnat the insert range (rngInsert) to be after the tables
        'and if there are no tables, then we want the insert rnage to be at the beginning of the header
        '
        rngInsert = hf.Range
        '
        If hf.Range.Tables.Count = 0 Then
            rngInsert.Collapse(WdCollapseDirection.wdCollapseStart)
        Else
            'We'll out the insert range at the beginning of the last paragraph
            rngInsert.Collapse(WdCollapseDirection.wdCollapseEnd)
            'rngInsert.End = rngInsert.End - 1
        End If
        '
        Try
            shp = Me.pnl_BackPanel_Insert(hf, rngInsert, rgbFill, strPanelName)
        Catch ex As Exception
            shp = Nothing
        End Try
        '
        '
        Return shp
        '
    End Function
    '
    ''' <summary>
    ''' This method will find (in the first page and primary page headers) and delete all shapes with
    ''' the name 'strPanelName'
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="strPanelName"></param>
    Public Sub pnl_BackPanel_Delete(ByRef sect As Word.Section, Optional strPanelName As String = "aac_BackColour")
        Dim lstOfPanels As List(Of cShapeMgr)
        '
        lstOfPanels = Me.pnl_getBackPanel_PlaceHolders(sect, strPanelName)
        Me.pnl_BackPanel_Delete(lstOfPanels)
        '
    End Sub
    '

    '
    Public Sub pnl_BackPanel_Delete(ByRef lstOfPanels As List(Of cShapeMgr))
        Dim j As Integer
        Dim shp As Word.Shape
        '
        If lstOfPanels.Count <> 0 Then
            For j = lstOfPanels.Count - 1 To 0 Step -1
                shp = lstOfPanels.Item(j).shp
                shp.Delete()
            Next
            '
        End If
        '
    End Sub
    '
    '
    Public Sub pnl_BackPanel_Delete(ByRef hf As Word.HeaderFooter)
        '
        If hf.Range.ShapeRange.Count <> 0 Then
            For Each shp In hf.Range.ShapeRange
                If shp.Name = Me.strShapeName Or shp.Name = Me.strBackPanel_CaseStudy Or shp.Name Like "Free*" Or shp.Name Like "Rect*" Then
                    shp.Delete()
                    Exit For
                End If
            Next
        End If
        '
    End Sub
    '
    Public Sub pnl_BackPanel_Delete(ByRef hf As Word.HeaderFooter, strPanelName As String)
        '
        If hf.Range.ShapeRange.Count <> 0 Then
            For Each shp In hf.Range.ShapeRange
                If shp.Name = strPanelName Then
                    shp.Delete()
                    Exit For
                End If
            Next
        End If
        '
    End Sub


    ''' <summary>
    ''' This method will insert the back panel into the header at the range rngInsert. If rgbFill is >= 0 then this specified
    ''' value will be used to fill the shape. If rgbFill is less than 0 then the locally specified value for rgbFill in the
    ''' class parameters will be used to fill the shape
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="rng"></param>
    ''' <param name="rgbFill"></param>
    ''' <returns></returns>
    Public Function pnl_BackPanel_Insert(ByRef hf As Word.HeaderFooter, ByRef rng As Word.Range, Optional rgbFill As Integer = -1, Optional strPanelName As String = "") As Word.Shape
        Dim shp As Word.Shape
        Dim sect As Word.Section
        Dim objWCAGMgr As New cWCAGMgr()
        '
        sect = rng.Sections.Item(1)
        shp = Nothing
        '

        '
        'Go to the default fill (dark purple if no fill is specified)
        '
        If rgbFill < 0 Then rgbFill = RGB(20, 0, 52)
        '
        'If the hf already has a shape with this name, then delete it.
        'We don't want multiple shapes.. We also get rid of any FreeForm or Rectangle shapes
        'that might haved snuck in
        '
        Me.pnl_BackPanel_Delete(hf)

        'If hf.Shapes.Count <> 0 Then
        'For Each shp In hf.Shapes
        'If shp.Name = Me.strShapeName Or shp.Name = Me.strBackPanel_CaseStudy Or shp.Name Like "Free*" Or shp.Name Like "Rect*" Then
        'shp.Delete()
        'Exit For
        'End If
        'Next
        'End If
        '
        Try
            shp = hf.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0.0, 0.0, sect.PageSetup.PageWidth, sect.PageSetup.PageHeight, rng)
            shp.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
            shp.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
            shp.Left = 0.0
            shp.Top = 0.0
            shp.ZOrder(MsoZOrderCmd.msoSendBehindText)
            shp.ZOrder(MsoZOrderCmd.msoSendToBack)
            '
            'To make borders invisible
            shp.Line.Visible = False
            '
            'Optional section to allow different panel shape names
            '
            If strPanelName = "" Then
                shp.Name = Me.strShapeName
            Else
                shp.Name = strPanelName
            End If
            'shp.Name = strShapeName
            'shp.Name = strBackPanel_CaseStudy

            '
            If Not objWCAGMgr.wcag_docProps_isAccessible() Then
                shp.Fill.ForeColor.RGB = rgbFill
            Else
                objWCAGMgr.wcag_backColour_BorderAndFill(shp)
            End If
            '
            shp.LockAnchor = True
            '
            shp.AlternativeText = "Decorative back colour. Please ignore"
            '
            objWCAGMgr.wcag_set_decorative(shp, True)
            'Dim objShp As Object
            'objShp = shp
            'objShp.Decorative = 1
            '
        Catch ex As Exception
            'If we get here it is almost certain that we didn't place the image
            'in the HeaderFooter.. That is Set Hf = Selection.HeaderFooter failed because
            'HeaderFooter was nothin
            '
            shp = Nothing
            MsgBox("Error - Unknown in ChptBase_BackPanel_Insert")
        End Try
        '
        Return shp
        '
    End Function
    '
    ''' <summary>
    ''' This function takes as input a list of panels (only ever expect one)... And resize the underlying
    ''' shape to fit the page
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="lstOfPanels"></param>
    ''' <returns></returns>
    Public Function pnl_resize_PanelToFillPage(ByRef sect As Word.Section, ByRef lstOfPanels As List(Of cShapeMgr)) As cShapeMgr
        Dim objShpMgr As cShapeMgr
        '
        objShpMgr = Nothing
        '
        Try
            If lstOfPanels.Count > 0 Then
                objShpMgr = lstOfPanels.Item(0)
                objShpMgr.shp.Width = sect.PageSetup.PageWidth
                objShpMgr.shp.Height = sect.PageSetup.PageHeight
            End If
            '
        Catch ex As Exception
            objShpMgr = Nothing
        End Try
        '
        Return objShpMgr
    End Function
    '
    Public Function pnl_resize_PanelToFillPage(ByRef sect As Word.Section) As cShapeMgr
        Dim listOfPanels As List(Of cShapeMgr)
        Dim objShpMgr As cShapeMgr
        '
        objShpMgr = Nothing
        Try
            listOfPanels = Me.pnl_getBackPanel_PlaceHolders(sect)
            objShpMgr = Me.pnl_resize_PanelToFillPage(sect, listOfPanels)
        Catch ex As Exception
            objShpMgr = Nothing
        End Try

        Return objShpMgr
    End Function

    ''' <summary>
    ''' Version built 20231022
    ''' </summary>
    ''' <param name="InsertOrPaste"></param>
    ''' <param name="lstOfBackPanels"></param>
    ''' <returns></returns>
    Public Function pnl_BackPanel_Replace(InsertOrPaste As String, ByRef lstOfBackPanels As List(Of cShapeMgr), ByRef oldPic As Word.Shape) As Word.Shape
        Dim picDialog As Word.Dialog
        Dim objShp_BackPanel, objShp_NewPic As cShapeMgr
        Dim objSectMgr As New cSectionMgr()
        Dim objImgrMgr As New cImageMgr()
        Dim strFilePath As String
        Dim sect As Word.Section
        Dim hf As Word.HeaderFooter
        Dim rng, rngSel, rngPasted As Word.Range
        Dim NewPic As Word.Shape
        Dim NewPic_InLine As InlineShape
        Dim cropRect As Word.Shape
        Dim frm As frm_pictControl2
        Dim h, v As Single
        Dim oldSel As Selection
        '
        oldSel = Globals.ThisAddIn.Application.Selection
        rngSel = oldSel.Range
        '
        objShp_BackPanel = lstOfBackPanels.Item(0)
        sect = objShp_BackPanel.anchor.Sections.Item(1)
        Me.currentSect = sect
        '
        '** The range has been verified as giving the right answer
        rng = objShp_BackPanel.anchor
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        '
        pnl_BackPanel_Replace = Nothing
        NewPic = Nothing
        oldPic = Nothing
        '
        Try
            picDialog = Globals.ThisAddIn.Application.Dialogs(Word.WdWordDialog.wdDialogInsertPicture)
            Me.strPicturePlaceHolderName = objShp_BackPanel.name
            hf = objShp_BackPanel.hf
            '
            Select Case InsertOrPaste
                Case "pasteImage"
                    'WorkAround.. Directly pasting image into Header generally causes the image to go into the header table
                    'Tested for Word.Options.Advanced.Insert/paste picture as
                    '20231025   behind text             ..OK
                    '20231025   Inline with text        ..OK
                    '
                    'Move to the end of the section, paste, then select
                    Try
                        rng = objSectMgr.sct_Set_RngTo_SectionEndParagraph_Beginning()
                        rng.Select()
                        'rng.PasteSpecial(, , WdOLEPlacement.wdFloatOverText, , WdPasteDataType.wdPasteMetafilePicture)
                        rng.PasteSpecial(, , WdOLEPlacement.wdFloatOverText, , WdPasteDataType.wdPasteBitmap)
                        '
                        'Get the current selection, then extend it by one char to include the image
                        rngPasted = Globals.ThisAddIn.Application.Selection.Range
                        rngPasted.End = rngPasted.End + 1
                        rngPasted.Select()
                        '
                        '****
                        'Dim selection As Word.Selection = Application.Selection
                        'Globals.ThisAddin.Application.Selection.Copy()
                        'Dim file As String = "C:\Temp\to\file.txt"
                        'System.IO.File.WriteAllText(file, Selection)
                        '*****
                        '
                        NewPic = objImgrMgr.img_get_ImageAsShape(rngPasted)
                        'NewPic.Select()
                        'Globals.ThisAddin.Application.Selection.InlineShapes.Item(1).w
                        '
                        '*** Insert a new copy to be placed back on the clipboard later
                        'rng.PasteSpecial(, , WdOLEPlacement.wdFloatOverText, , WdPasteDataType.wdPasteBitmap)
                        'rngOldPic = Globals.ThisAddin.Application.Selection.Range
                        'rngOldPic.End = rngPasted.End + 1
                        'rngOldPic.Select()
                        '
                        'oldPic = objImgrMgr.img_get_ImageAsShape(rngOldPic)
                        'oldPic.AlternativeText = "oldPic"
                        'oldPic.Visible = False
                        '
                        '***
                        '
                        'Set NewPic = HFSHapes.Item(HFSHapes.Count)
                        NewPic.Top = Me.objTools.MillimetersToPoints(297.0#) - NewPic.Height - 40.0#
                        NewPic.Name = objShp_BackPanel.name
                        '
                        'oldPic.Select()
                        '
                    Catch ex As Exception
                        NewPic = Nothing
                        GoTo finis
                    End Try
                    '
                Case "insertImage"
                    strFilePath = ""
                    picDialog.name = "*.*"
                    '
                    Select Case picDialog.Display
                        Case 0                              'Cancel button
                            NewPic = Nothing
                            GoTo finis
                        Case -1                             'OK button
                            strFilePath = picDialog.Name
                            If Not (My.Computer.FileSystem.FileExists(strFilePath)) Then
                                NewPic = Nothing
                                GoTo finis
                            End If
                        Case -2                             'Close button
                            NewPic = Nothing
                            GoTo finis
                        Case Else
                            NewPic = Nothing
                            GoTo finis
                    End Select
                    '
                    Try
                        rng = objSectMgr.sct_Set_RngTo_SectionEndParagraph_Beginning()
                        rng.Select()
                        '
                        '


                        NewPic_InLine = sect.Range.InlineShapes.AddPicture(strFilePath,,, rng)
                        'rngPasted = Globals.ThisAddin.Application.Selection.Range
                        'rngPasted.End = rngPasted.End + 1
                        'rngPasted.Select()


                        'NewPic = hf.Shapes.AddPicture(strFilePath, False, True, 0, 0, , , rng)
                        NewPic_InLine.Select()
                        rngPasted = Globals.ThisAddIn.Application.Selection.Range
                        '
                        Globals.ThisAddIn.Application.ScreenRefresh()

                        rngPasted = Globals.ThisAddIn.Application.Selection.Range
                        NewPic = rngPasted.InlineShapes.Item(rngPasted.InlineShapes.Count).ConvertToShape()
                        'rngPasted = Globals.ThisAddin.Application.Selection.Range
                        'rngPasted.End = rngPasted.End + 1
                        'rngPasted.Select()
                        Globals.ThisAddIn.Application.ScreenRefresh()

                        '
                        '
                        'NewPic = objImgrMgr.img_get_ImageAsShape(rngPasted)

                        NewPic.Top = Me.objTools.MillimetersToPoints(297.0#) - NewPic.Height - 40.0#
                        '
                        Globals.ThisAddIn.Application.ScreenRefresh()

                        'NewPic.name = strShapeName
                        'NewPic.Name = objShp_BackPanel.name
                    Catch ex As Exception
                        NewPic = Nothing
                        GoTo finis
                    End Try
            End Select
            '
            NewPic.AlternativeText = objShp_BackPanel.altText
            objShp_NewPic = New cShapeMgr()
            objShp_NewPic.InitShape(NewPic, hf)
            '
            'Call NewPic.ScaleWidth(1.0#, MsoTriState.msoTrue)  'sort of resets the record of width & height?
            'Call NewPic.ScaleHeight(1.0#, MsoTriState.msoTrue)
            '
            'objShp_BackPanel.width_original = NewPic.Width
            'objShp_BackPanel.height_original = NewPic.Height

            Call Me.setImagePageProperties(NewPic)          'Page positional properties
            '
            Call NewPic.ScaleWidth(1.0#, MsoTriState.msoTrue)  'sort of resets the record of width & height?
            Call NewPic.ScaleHeight(1.0#, MsoTriState.msoTrue)
            '
            objSectMgr.sct_fit_ShapeToPage(objShp_NewPic, sect)
            '
            'Call Me.setImageScale2(NewPic, sect, objShp_NewPic)
            '
            'Test the scale factors.. they are correct
            'Reset  scale
            'Call NewPic.ScaleWidth(1.0#, MsoTriState.msoTrue)  'sort of resets the record of width & height?
            'Call NewPic.ScaleHeight(1.0#, MsoTriState.msoTrue)

            h = objShp_NewPic.scaleFactor_H
            v = objShp_NewPic.scaleFactor_W
            '
            'The following is now done in Me.setImageScale2
            'objShp_NewPic.AdjustScaleFactors(NewPic)
            '
            'objShp_BackPanel.scaleFactor_W = objShp_BackPanel.width_original / NewPic.Width
            Call Me.pnl_Set_ImagePosition(NewPic, sect)
            '
            pnl_BackPanel_Replace = NewPic
            '
            '
            'GoTo loop2
            '
            If Globals.ThisAddIn.Application.ActiveDocument.ProtectionType <> WdProtectionType.wdNoProtection Then Call Globals.ThisAddIn.Application.ActiveDocument.Unprotect("PASSWORD")
            '
            'Set cropRect = CH_ImgMgr.insertCropRect(objShp, shp)
            cropRect = Me.pnl_insert_CropRect(sect, objShp_BackPanel, NewPic)
            'cropRect = Me.insertCropRect(sect, objShp_NewPic, NewPic)
            '
            '
            'We need to adjust cropping rectange position and etc... It is no longer anchored at the
            'table at the top of the page

            frm = New frm_pictControl2()
            frm.shp_ImageToBeClipped = NewPic                           'image to be clipped (as Shape)
            frm.cropRect = cropRect                                     'Cropping Rectangle
            frm.shp_ImageToBeClipped_as_cShapeMgr = objShp_NewPic
            frm.shp_toBeFilled = objShp_BackPanel                       'SHape to be filled
            frm.LayoutName = Me.strFormLayoutName
            '

            frm.Top = 250
            frm.Left = 250
            '
            frm.parentBackPanelMgr = Me
            '
            If frm.ShowDialog() = 2 Then
                'Cancel = 2, OK (FInish) = 1
                NewPic.Delete()
                NewPic = Nothing
            End If
            '
            'Cancel = 2
            'OK = 1
            '
            'NewPic.ZOrder(MsoZOrderCmd.msoSendToBack)
            'NewPic.ZOrder(MsoZOrderCmd.msoSendBehindText)
            '
            'Now re-apply dimensions
loop2:
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.ScreenRefresh()
            '

            GoTo finis

        Catch ex As Exception
            Globals.ThisAddIn.Application.ActiveDocument.Undo(1)
            Globals.ThisAddIn.Application.ScreenUpdating = True
            Globals.ThisAddIn.Application.ScreenRefresh()
            Globals.ThisAddIn.Application.ScreenRefresh()

        End Try
        '
finis:
        Globals.ThisAddIn.Application.ScreenRefresh()

        Return NewPic
        '
        '
        '

    End Function
    '
    '
    ''' <summary>
    ''' This method will lock the aspect ratio of the shape and then resize it
    ''' </summary>
    ''' <param name="shp"></param>
    ''' <param name="sect"></param>
    Public Sub setImageScale2(ByRef shp As Word.Shape, ByRef sect As Word.Section, ByRef objShp_NewPic As cShapeMgr)
        ' ASPECT RATIO CONTEXT = HEIGHT/WIDTH
        'set shape size according to page size
        'shp.
        Dim objSectMgr As New cSectionMgr()
        '
        '**** In its current form this one distorts... Partially fixed 20231027
        objSectMgr.sct_fit_ShapeToPage(objShp_NewPic, sect)
        '
    End Sub
    '
    Public Sub pnl_Set_ImagePosition(ByRef shp As Word.Shape, ByRef sect As Word.Section)
        'This positions the target image in the centre of the page
        Dim AspectRatio_Page As Single
        '
        AspectRatio_Page = sect.PageSetup.PageHeight / sect.PageSetup.PageWidth
        '
        If AspectRatio_Page > 1 Then  'Page is Portrait
            shp.Left = (595 - shp.Width) / 2
            'shp.top = (841 / 2) - shp.height / 2
            shp.Top = (842 - shp.Height) / 2                     ' Sit at the bottom of the page - Allow 20pts for footer
        ElseIf AspectRatio_Page <= 1 Then 'Page is Landscape
            shp.Left = (841 / 2) - shp.Width / 2
            shp.Top = (595 / 2) - shp.Height / 2
        End If
        '
    End Sub
    '
    '
    ''' <summary>
    ''' ObjSrcShp_tobeFilled is the source shape from which the clipping mask gets its shape/aspect ratio. This is the shape that will
    ''' eventually be filled with a new image. The object shp_NewPict_tobeClipped is the
    ''' new picture that the clipping mask is placed over to get a new pciture that will insert into the source shape
    ''' without distortion
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="objSrcShp_tobeFilled"></param>
    ''' <param name="shp_ImageToBeClipped"></param>
    ''' <returns></returns>
    Public Function pnl_insert_CropRect(ByRef sect As Section, ByRef objSrcShp_tobeFilled As cShapeMgr, ByRef shp_ImageToBeClipped As Word.Shape) As Word.Shape
        'Dim aspectRatio_CropRect, aspectRatio_ImageToBeClipped As Single
        'Dim dummy_h, dummy_w, scaleFactor_h As Single
        Dim clipHeight, clipWidth As Single
        Dim strShpAspectRatio_CropRect As String
        Dim strImageToBeClipped As String
        Dim hf As HeaderFooter
        Dim rng As Word.Range
        Dim shpRect As Word.Shape
        Dim irror As Boolean
        Dim objCropRect As New cCropRectMgr()
        Dim lstOfCropRectDimensions As New Collection()
        '
        'hf = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        hf = objSrcShp_tobeFilled.hf
        rng = shp_ImageToBeClipped.Anchor
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        strShpAspectRatio_CropRect = ""
        strImageToBeClipped = ""
        irror = False
        '
        '
        'Get Cropping Rectangle dimensions in pts.... Remember aspect ratio is h/w
        lstOfCropRectDimensions = objCropRect.rct_(objSrcShp_tobeFilled, shp_ImageToBeClipped)
        clipWidth = CSng(lstOfCropRectDimensions("clipWidth"))
        clipHeight = CSng(lstOfCropRectDimensions("clipHeight"))
        '
        '
        'Set inserCtropRect = ActiveDocument.Shapes.AddShape(msoShapeRectangle, shp.left, shp.top, objShp.width, objShp.height, shp.Anchor)
        shpRect = hf.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, shp_ImageToBeClipped.Left, shp_ImageToBeClipped.Top, clipWidth, clipHeight, rng)
        shpRect.LockAspectRatio = MsoTriState.msoTrue
        shpRect.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
        shpRect.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
        shpRect.LockAnchor = True
        shpRect.WrapFormat.AllowOverlap = True
        shpRect.WrapFormat.Type = WdWrapType.wdWrapNone
        '
        'insertCropRect.left = shp.left
        'insertCropRect.top = shp.top
        shpRect.Fill.Transparency = 0.5
        shpRect.Fill.BackColor.RGB = RGB(0, 128, 128)
        '
        'Make sure the cropping rectangle is aligned with the image to be clipped
        shpRect.Top = shp_ImageToBeClipped.Top
        shpRect.Left = shp_ImageToBeClipped.Left
        '
        Return shpRect
    End Function
    '
    '
    Public Sub pnl_UserImage_fill_with_Image(ByRef objShpMgr As cShapeMgr)
        Dim img As System.Drawing.Image
        'NewPic.Select()
        'sel = Globals.ThisAddin.Application.Selection
        'sel.CopyAsPicture()
        'NewPic.Delete()
        '
        '***** It works
        'img = Clipboard.GetImage()
        img = System.Windows.Forms.Clipboard.GetImage()
        img.Save(Me.strScratchFilePath)
        objShpMgr.shp.Fill.UserPicture(Me.strScratchFilePath)

    End Sub
    '
    '    '
    Public Sub setImagePageProperties(ByRef shp As Word.Shape)
        'Call shp.ScaleWidth(1.0#, MsoTriState.msoTrue)  'sort of resets the record of width & height?
        'Call shp.ScaleHeight(1.0#, MsoTriState.msoTrue)
        '
        shp.LockAspectRatio = MsoTriState.msoTrue
        shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
        shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
        shp.LockAnchor = True
        shp.WrapFormat.AllowOverlap = True
        shp.WrapFormat.Type = WdWrapType.wdWrapNone
        '
    End Sub
    '


End Class

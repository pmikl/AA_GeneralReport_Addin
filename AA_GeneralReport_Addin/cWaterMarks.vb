Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
''' <summary>
''' 
'''Originally written in vba, some account taken for conversion, but this
'''was not a priority at the time this class was written
'''
'''Peter Mikelaitis October 2015...http://mikl.com.au
'''Ported to VB.NET Jan 2017 from version 97p21p05
'''
''' </summary>
Public Class cWaterMarks
    Inherits cGlobals

    Public Sub New()
        MyBase.New()
    End Sub
    '
    '
    Public Sub msg_DocumentIsProtected()
        Dim strMsg As String
        Dim dlgResult As Integer
        '
        strMsg = "The current document is protected." & vbCr & vbCr _
        & "In order to modify the Water Marks you'll need to unprotect" & vbCr _
        & "the document. Contact the author's or local Admin for the protection password"
        dlgResult = MsgBox(strMsg)
        '
    End Sub
    '
    '
    Public Sub waterMarks_Remove_All()
        Dim sect As Section
        Dim hf As HeaderFooter
        Dim strFuzzyName As String
        Dim strFuzzyName2 As String
        Dim shp As Word.Shape
        Dim i, j, k As Integer
        '
        strFuzzyName = "waterMark_aa_*"                         'My custom watermarks
        strFuzzyName2 = "*WaterMark*"                           'Word's watermarks
        '
        On Error GoTo finis
        For Each sect In Globals.ThisAddin.Application.ActiveDocument.Sections
            For Each hf In sect.Headers
                If hf.Exists Then
                    For i = hf.Shapes.Count To 1 Step -1
                        shp = hf.Shapes.Item(i)
                        If shp.Name Like "waterMark_aa_*" Then shp.Delete()
                        If shp.Name Like "*_AAC_version" Then shp.Delete()
                        'If Not ((shp.name Like "logo*") Or (shp.name Like "txtBox*") Or (shp.name Like "txtBx*")) Then shp.Delete
                    Next i
                End If
            Next hf
        Next sect
        '
        'Call Me.remove_Security_WaterMarkfromCoverPage()
        '
        Exit Sub
finis:

    End Sub
    '
    Public Sub waterMarks_Remove_VersionMark()
        Me.waterMarks_Remove("*_AAC_version")
    End Sub
    '
    '
    Public Sub waterMarks_Remove_VersionMark(ByRef sect As Word.Section)
        Dim hf As Word.HeaderFooter
        Dim shp As Word.Shape
        '
        For Each hf In sect.Headers
            If hf.Exists Then
                For i = hf.Range.ShapeRange.Count To 1 Step -1
                    shp = hf.Range.ShapeRange.Item(i)
                    If shp.Name Like "*_AAC_version" Then shp.Delete()
                    'If Not ((shp.name Like "logo*") Or (shp.name Like "txtBox*") Or (shp.name Like "txtBx*")) Then shp.Delete
                Next i
            End If
        Next hf
    End Sub

    '
    Public Sub waterMarks_Remove(strFuzzyName As String)
        'This method will remove all instances of Water Marks with
        'the fuzzy name strFuzzyName.. Typically strFuzzyName would be;
        '-  "waterMark_aa_*             All Acil Allen Water Marks
        '-  "waterMark_aa_*_stat"       Status level Water Marks
        '-  "waterMark_aa_*_sec"        Security Level Water Marks
        '
        Dim objCpMgr As cCoverPageMgr
        Dim objStylesMgr As New cStylesManager()
        Dim myDoc As Word.Document
        Dim sect As Section
        Dim hf As HeaderFooter
        Dim shp As Word.Shape
        Dim i As Integer
        '
        objCpMgr = New cCoverPageMgr
        myDoc = Me.glb_get_wrdActiveDoc
        '
        Try
            'If we remove any of the security/status level watermarks, then we wnat the styles
            'put back the way they were... So, when inserting we always start from the smae place
            '
            If strFuzzyName Like "*_aa_sec" Then objStylesMgr.style_getCreateRefresh_waterMark_sec(myDoc)
            If strFuzzyName Like "*_aa_stat" Then objStylesMgr.style_getCreateRefresh_waterMark_stat(myDoc)
            '
            For Each sect In Me.glb_get_wrdActiveDoc.Sections
                If objCpMgr.cp_Bool_IsCoverPage(sect) And strFuzzyName Like "waterMark_aa_*_sec" Then
                    Call Me.remove_Security_WaterMarkfromCoverPage(sect)
                Else
                    For Each hf In sect.Headers
                        If hf.Exists Then
                            For i = hf.Shapes.Count To 1 Step -1
                                shp = hf.Shapes.Item(i)
                                If shp.Name Like strFuzzyName Then shp.Delete()
                                'If Not ((shp.name Like "logo*") Or (shp.name Like "txtBox*") Or (shp.name Like "txtBx*")) Then shp.Delete
                            Next i
                        End If
                    Next hf
                End If
            Next sect
            '
        Catch ex As Exception

        End Try
        '
        '
        '
        Exit Sub
finis:

    End Sub
    '
    Public Sub waterMarks_RemoveFromSection_Stat(ByRef sect As Word.Section)
        'This method will remove the Status water mark from the current
        'section.. Typically used in the letter to remove full report
        'Water Marks
        '
        Dim shpRng As Word.ShapeRange
        Dim hf As HeaderFooter
        Dim shp As Word.Shape
        Dim i As Integer
        Dim strFuzzyName As String
        '
        'strFuzzyName = "*_aa_stat"
        strFuzzyName = "*_stat"                 'Will do current an legacy watermarks
        '
        For Each hf In sect.Headers
            If hf.Exists Then
                shpRng = hf.Range.ShapeRange
                For i = shpRng.Count To 1 Step -1
                    shp = shpRng.Item(i)
                    If shp.Name Like strFuzzyName Then shp.Delete()
                Next i
            End If
        Next hf

    End Sub
    '
    Public Sub waterMarks_RemoveFromSection_Sec(ByRef sect As Word.Section)
        'This method will remove the Security water mark from the current
        'section.. Typically in a letter section to remove remnant water Marks
        '
        Dim shpRng As Word.ShapeRange
        Dim hf As HeaderFooter
        Dim shp As Word.Shape
        Dim i As Integer
        Dim strFuzzyName As String
        Dim objCpMgr As cCoverPageMgr
        '
        objCpMgr = New cCoverPageMgr()
        'strFuzzyName = "*_aa_sec"
        strFuzzyName = "*_sec"                      'WIll do current and laegacy watermarks
        '
        'We are removing from a standard page so its a background removal
        'Need to use the Shape range of the hf in order to restrict the
        'action to that Header
        For Each hf In sect.Headers
            If hf.Exists Then
                shpRng = hf.Range.ShapeRange
                For i = shpRng.Count To 1 Step -1
                    shp = shpRng.Item(i)
                    If shp.Name Like strFuzzyName Then shp.Delete()
                Next i
            End If
        Next hf
        '
    End Sub
    '
    Public Sub remove_Security_WaterMarkfromCoverPage(ByRef sect As Word.Section)
        'Now do the Secuirty Water Mark in the Cover Page.. We must pass sect which
        'is the section that cont... Hold over form the env version where the Watermark
        'was in the COverPage Table... It is no longer there
        'Dim tbl As Word.Table
        'Dim drCell As Word.Cell
        'Dim rng As Word.Range
        '
        'tbl = sect.Range.Tables(1)
        'drCell = tbl.Range.Cells(4)
        'rng = drCell.Range
        'rng.Delete()
        '
finis:
        '
    End Sub
    '
    Public Function isLandscape(ByRef sect As Word.Section) As Boolean
        isLandscape = False
        If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then isLandscape = True
    End Function
    '
    '
    Public Sub waterMark_to_CoverPage(strName As String, ByRef sect As Word.Section, Optional strAlignment As String = "right", Optional txtColour As Long = -1)
        'This method will insert WaterMarks into the general Body
        'sections
        Dim lstOfDimensions As Collection
        Dim hf As HeaderFooter
        Dim rng As Word.Range
        Dim rngAnchor As Word.Range
        Dim shp As Word.Shape
        Dim left, top, width, height As Single
        Dim strCaption As String
        '
        If txtColour = -1 Then txtColour = Me._glb_colour_WaterMark_Grey_sec
        strCaption = ""
        '
        'Do Security Water marks in the Header Table
        '
        If strName Like "*_sec" Then
            For Each hf In sect.Headers
                If hf.Exists Then
                    rng = hf.Range
                    If rng.Tables.Count = 0 Then
                        rng = hf.Range
                        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                        '
                        strCaption = Me.waterMark_get_securityCaption(strName)
                        '
                        left = 150
                        top = 150
                        width = 150
                        height = 40
                        '
                        lstOfDimensions = New Collection()
                        lstOfDimensions.Add(230.25, "left")
                        lstOfDimensions.Add(45, "top")
                        lstOfDimensions.Add(311, "width")
                        lstOfDimensions.Add(34, "height")

                        '
                        shp = Me.waterMark_insertShape_toRange(rng, hf, lstOfDimensions)
                        shp.Name = shp.Name + "_aa_sec"                                                              'Workaround.. The shape must be named outside the above routine.. I don't know why its
                        Me.waterMark_shape_writeText(shp, strCaption)
                        shp.TextFrame.TextRange.Font.Size = 18
                        'Me.waterMark_shape_align(tbl, strAlignment, shp)

                    End If
                End If
            Next hf
        End If
        '
        'Do the Release Status Water marks in the background
        '
        If strName Like "*_stat" Then
            For Each hf In sect.Headers
                If hf.Exists Then
                    rngAnchor = hf.Range
                    rngAnchor.Collapse(WdCollapseDirection.wdCollapseEnd)
                    '
                    strCaption = ""
                    Select Case strName
                        Case "waterMark_aa_draft_stat"
                            strCaption = "DRAFT"
                        Case "waterMark_aa_draftOnly_stat"
                            strCaption = "DRAFT ONLY"
                    End Select
                    '
                    If sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait Or sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
                        height = 40
                        width = 285
                        left = sect.PageSetup.LeftMargin
                        top = 95

                        '
                        lstOfDimensions = New Collection()
                        lstOfDimensions.Add(left, "left")
                        lstOfDimensions.Add(top, "top")
                        lstOfDimensions.Add(width, "width")
                        lstOfDimensions.Add(height, "height")
                        '
                        shp = Me.waterMark_insertShape_toRange(rngAnchor, hf, lstOfDimensions)
                        'shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
                        'shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
                        shp.Name = strName
                        ''Workaround.. The shape must be named outside the above routine.. I don't know why its
                        Me.waterMark_shape_writeText(shp, strCaption, Me.glb_var_style_waterMark_stat)
                        'shp.TextFrame.TextRange
                        shp.TextFrame.TextRange.Font.Size = 30
                        shp.TextFrame.TextRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
                        'shp.Rotation = -90
                        'shp.Left = -sect.PageSetup.PageWidth / 2
                        'shp.Top = sect.PageSetup.PageHeight / 2
                    Else

                    End If
                    '
                    'shp = Me.waterMark_textBoxTo_Header("DRAFT", hf, tbl, rng, strName, 80.0, "centre")
                    'shp = Me.waterMark_textBoxTo_Header("DRAFT", hf, left, top, width,
                    'height, rng, strName, 80.0, RGB(255, 0, 0))
                    '
                    'shp.Rotation = -45
                End If
            Next
        End If
        Exit Sub

    End Sub
    '
    ''' <summary>
    ''' This function will take the old document status object name (i.e. as it was in legacy versions) and
    ''' translate it to the actual caption to be written
    ''' </summary>
    ''' <param name="strName"></param>
    ''' <returns></returns>
    Public Function waterMark_get_statusCaption(ByRef strName As String) As String
        Dim strCaption As String
        '
        strCaption = ""
        Select Case strName
            Case "waterMark_aa_draft_aa_stat"
                strCaption = "DRAFT"
            Case "waterMark_aa_draftOnly_aa_stat"
                strCaption = "DRAFT ONLY"
        End Select
        '
        Return strCaption
    End Function
    '
    ''' <summary>
    ''' This function will take the old security object name (i.e. as it was in legacy versions) and
    ''' translate it to the actual caption to be written
    ''' </summary>
    ''' <param name="strName"></param>
    ''' <returns></returns>
    Public Function waterMark_get_securityCaption(ByRef strName As String) As String
        Dim strCaption As String
        '
        Select Case strName
            Case "waterMark_aa_Cabinet_aa_sec"
                strCaption = "CABINET-IN-CONFIDENCE"
            Case "waterMark_aa_Commercial_aa_sec"
                strCaption = "COMMERCIAL-IN-CONFIDENCE"
            Case "waterMark_aa_Confidential_aa_sec"
                strCaption = "CONFIDENTIAL"
            Case "waterMark_aa_Restricted_aa_sec"
                strCaption = "RESTRICTED CIRCULATION"
                '
            Case "waterMark_aa_atg_UNOFFICIAL_aa_sec"
                strCaption = "UNOFFICIAL"
            Case "waterMark_aa_atg_OFFICIAL_aa_sec"
                strCaption = "OFFICIAL"
            Case "waterMark_aa_atg_OFFICIAL-Sensitive_aa_sec"
                strCaption = "OFFICIAL:Sensitive"

            Case Else
                strCaption = ""
        End Select
        '
        Return strCaption
    End Function


    Public Sub waterMark_to_CoverPage(strName As String, ByRef sect As Word.Section, ByRef objBB As cBBlocksHandler)
        Dim hf As HeaderFooter
        Dim rng As Word.Range
        Dim rngOfBlock As Word.Range
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        Dim shp As Word.Shape
        Dim colourTooDark As Boolean
        '
        On Error GoTo finis
        '
        colourTooDark = False
        '
        If strName Like "*_sec" Then
            '
            'We are installing a Security level Water Mark in the
            'background of the Cover Page
            For Each hf In sect.Headers
                If hf.Exists Then
                    rng = hf.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                    'Call rng.Move(wdParagraph, -1)
                    rngOfBlock = objBB.insertBuildingBlockFromDefaultLibToRange(strName, "waterMarks", rng)
                    If rngOfBlock.ShapeRange.Count <> 0 Then
                        shp = rngOfBlock.ShapeRange.Item(1)
                        shp.Name = strName
                        shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
                        shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
                        shp.Top = 96.65
                        shp.Left = 0.0
                        shp.LockAspectRatio = True
                        'shp.height = 34.15
                        Call shp.ScaleWidth(2.0, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromTopLeft)
                        shp.Rotation = 0#
                        'shp.Left = (sect.PageSetup.PageWidth - shp.Width) / 2
                        '
                        'If shp.Name Like "*_sec" Then shp.Fill.ForeColor.RGB = RGB(147, 147, 147)
                        If shp.Name Like "*_sec" Then shp.Fill.ForeColor.RGB = RGB(255, 0, 0)
                        'shp.Fill.ForeColor.RGB = RGB(255, 0, 0)
                        'shp.Fill.Transparency = 0
                    End If
                End If
            Next hf
            '
            'We are installing a Security level Wate Mark in the
            'text layer of the Cover Page
            'tbl = sect.Range.Tables(1)
            'drCell = tbl.Range.Cells(4)
            'If drCell.Shading.BackgroundPatternColor < 12000000 Then colourTooDark = True
            'If drCell.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic Then colourTooDark = False
            'rng = drCell.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'rngOfBlock = objBB.insertBuildingBlockFromDefaultLibToRange(strName, "waterMarks", rng)
            'If rngOfBlock.ShapeRange.Count <> 0 Then
            'shp = rngOfBlock.ShapeRange.Item(1)
            'shp.Name = strName
            'shp.LockAspectRatio = True
            'If shp.Name Like "*_sec" Then
            'shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
            'shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
            'shp.Height = 9.85
            'shp.Top = 29.9
            'shp.Left = 0#
            'shp.Fill.ForeColor.RGB = RGB(147, 147, 147)
            'If colourTooDark Then shp.Fill.ForeColor.RGB = RGB(255, 254, 255)
            'shp.Fill.Transparency = 0
            'End If
            'End If
        End If
        If strName Like "*_stat" Then
            'We are installing a Status Water Mark in the backgound
            For Each hf In sect.Headers
                If hf.Exists Then
                    rng = hf.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                    'Call rng.Move(wdParagraph, -1)
                    rngOfBlock = objBB.insertBuildingBlockFromDefaultLibToRange(strName, "waterMarks", rng)
                    If rngOfBlock.ShapeRange.Count <> 0 Then
                        shp = rngOfBlock.ShapeRange.Item(1)
                        shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
                        shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage

                        'shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
                        shp.Name = strName
                        shp.LockAspectRatio = True
                        '
                        Call shp.ScaleWidth(0.3, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromTopLeft)
                        '
                        shp.Rotation = 0#
                        'shp.Top = 412.0
                        'shp.Top = 382.0
                        shp.Top = 52.0
                        'shp.Top = 128.0

                        shp.Left = sect.PageSetup.PageWidth - 56 - shp.Width

                        'shp.Left = (sect.PageSetup.PageWidth - shp.Width) / 2
                        shp.ZOrder(MsoZOrderCmd.msoBringToFront)
                        '
                        shp.Fill.ForeColor.RGB = Me._glb_colour_WaterMark_Grey_sec
                        shp.Fill.Transparency = 0.2
                    End If
                End If
            Next hf
        End If
        '

        Exit Sub
finis:
        '
    End Sub

    '
    Public Sub waterMarks_Add(strWaterMarkName As String, Optional strAlignment As String = "", Optional txtColour As Long = -1)
        Dim tbl As Word.Table
        Dim sectType As String
        Dim strSectionTag As String
        Dim drCell As Word.Cell
        Dim sect As Section
        Dim hf As HeaderFooter
        Dim strFuzzyName As String
        Dim strName As String
        Dim shp As Word.Shape
        Dim rng As Word.Range
        Dim rngOfBlock As Word.Range
        Dim objBB As cBBlocksHandler
        Dim objCpMgr As cCoverPageMgr
        Dim objSectMgr As cSectionMgr
        Dim objLetterMgr As New cStationeryLetter()
        Dim isCoverPage As Boolean
        Dim pgIsLandscape As Boolean
        Dim i, j, k As Integer
        Dim scaleFactor As Single
        Dim delta As Single                             'Offset from the bottom of the page
        '
        scaleFactor = 0.5                                   '1.0 means original size
        delta = 100.0#
        sectType = ""
        '
        strName = "waterMark_aa_" & strWaterMarkName
        '
        On Error GoTo finis
        objBB = New cBBlocksHandler
        objCpMgr = New cCoverPageMgr
        objSectMgr = New cSectionMgr
        '

        For Each sect In objSectMgr.objGlobals.glb_get_wrdActiveDoc.Sections
            strSectionTag = objSectMgr.sct_Get_SectionTag(sect)
            'Now because of document structural changes due to the T&G rebranding in 2020
            'wee need to dummy up some of the tags
            '
            If objCpMgr.cp_Bool_IsCoverPage(sect) Then strSectionTag = "tag_coverPage"

            If strSectionTag = "" Then
                'We may be in a letter.. so let's test
                If objLetterMgr.ltr_is_Stationery(sect) Then strSectionTag = "tag_letter"
            End If
            '
            Select Case strSectionTag
                Case "tag_coverPage"
                    shp = waterMark_sec_toBody(strSectionTag, strName, sect, txtColour)
                    shp = waterMark_stat_toBody(strSectionTag, strName, sect, strAlignment, txtColour)
                Case "tag_letter"
                    shp = waterMark_sec_toBody(strSectionTag, strName, sect, strAlignment, txtColour)
                Case "tag_contactsPage-Front", "tag_contactsPage-Back", "tag_partBanner", "tag_appendixPart", "tag_contactsPage-Back"
                    shp = waterMark_sec_toBody(strSectionTag, strName, sect, strAlignment, txtColour)
                    shp = waterMark_stat_toBody(strSectionTag, strName, sect, strAlignment, txtColour)
                Case Else
                    shp = waterMark_sec_toBody(strSectionTag, strName, sect, strAlignment, txtColour)
                    shp = waterMark_stat_toBody(strSectionTag, strName, sect, strAlignment, txtColour)
            End Select
        Next sect
        '
        'For Each sect In ActiveDocument.Sections
        'isCoverPage = objCpMgr.isCoverPage(sect)
        'If (isCoverPage) And (strName Like "*_sec") Then
        'Call Me.waterMarks_Add_CoverPage(strName, sect)
        'GoTo loop1
        'End If
        '
        'pgIsLandscape = Me.isLandscape(sect)
        'For Each hf In sect.Headers
        'If hf.Exists Then
        'Set rng = hf.Range
        'rng.Collapse (wdCollapseDirection.wdCollapseEnd)
        'Set rngOfBlock = objBB.insertBuildingBlockFromDefaultLibToRange(strName, "waterMarks", rng)
        'If rngOfBlock.ShapeRange.Count <> 0 Then
        'Application.ScreenUpdating = False
        'Set shp = rngOfBlock.ShapeRange.Item(1)
        'shp.name = strName
        'If isCoverPage Then
        'If shp.name Like "*_stat" Then
        'shp.top = 625.3
        'shp.LockAspectRatio = True
        'Call shp.ScaleWidth(0.5, msoFalse, msoScaleFromTopLeft)
        'shp.rotation = 0#
        'shp.left = (sect.PageSetup.PageWidth - shp.width) / 2
        'End If
        'End If
        '
        'If pgIsLandscape Then
        'If shp.name Like "*_sec" Then
        'shp.left = (sect.PageSetup.PageWidth - shp.width) / 2
        'End If
        'If shp.name Like "*_stat" Then
        'shp.top = 240.1
        'shp.left = (sect.PageSetup.PageWidth - shp.width) / 2
        'End If
        'End If
        'shp.LockAnchor = True
        'shp.LockAspectRatio = msoCTrue
        'shp.RelativeVerticalPosition = wdRelativeVerticalPositionBottomMarginArea
        'shp.RelativeHorizontalPosition = wdRelativeHorizontalPositionMargin
        'Call shp.ScaleWidth(scaleFactor, msoFalse, msoScaleFromTopLeft)
        'shp.top = 0#
        'shp.left = (Me.spaceBetweenMargins(sect) - shp.width) / 2
        'shp.left = (sect.PageSetup.PageWidth - shp.width) / 2 - sect.PageSetup.leftMargin
        'shp.left = 0#
        'shp.top = sect.PageSetup.PageHeight - delta + shp.height / 2
        'shp.Fill.ForeColor.RGB = RGB(255, 0, 0)
        'Application.ScreenUpdating = True
        'End If
        'End If
        'Next hf
        'loop1:
        'Next sect
        Exit Sub
finis:

    End Sub
    '
    ''' <summary>
    ''' This function will insert a shape (shp) of specific dimensions (lstOfDimensions) into a range (rngAnchor). The
    ''' anchor range must be located in the header(hf). For a lot of applications the anchor is in the first cell of
    ''' the header table
    ''' </summary>
    ''' <param name="rngAnchor"></param>
    ''' <param name="hf"></param>
    ''' <param name="lstOfDimensions"></param>
    ''' <returns></returns>
    Public Function waterMark_insertShape_toRange(ByRef rngAnchor As Range, ByRef hf As Word.HeaderFooter, ByRef lstOfDimensions As Collection) As Word.Shape
        'Cannot name the shape inside this routine. It causes miss alignment to the Header table
        'If I name it out side everthing appears to be OK
        Dim objWcag As New cWCAGMgr()
        Dim shp As Word.Shape
        'Dim txtBox As TextBox
        Dim left, top, width, height As Single
        '
        left = CSng(lstOfDimensions.Item("left"))
        top = CSng(lstOfDimensions.Item("top"))
        width = CSng(lstOfDimensions.Item("width"))
        height = CSng(lstOfDimensions.Item("height"))

        'txtBox = hf.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, left, top, width, height, rngAnchor)
        shp = hf.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 150, 20, rngAnchor)
        shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
        shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
        shp.LockAnchor = True
        '
        'shp.Name = strName
        '
        shp.Left = left
        shp.Top = top
        shp.Width = width
        shp.Height = height
        '
        'shp.TextFrame.Orientation = MsoTextOrientation.msoTextOrientationVertical
        shp.TextFrame.NoTextRotation = False
        'shp.TextFrame.TextRange.
        'shp.r
        '
        shp.TextFrame.MarginTop = 0.0
        shp.TextFrame.MarginBottom = 0.0
        shp.TextFrame.MarginRight = 0.0
        shp.TextFrame.MarginLeft = 0.0
        '
        shp.Fill.ForeColor.RGB = RGB(255, 255, 255)
        shp.Fill.BackColor.RGB = RGB(255, 255, 255)
        'shp.Line.Visible = False
        shp.Fill.Transparency = 1
        shp.Line.Visible = False
        '
        objWcag.wcag_set_decorative(shp, True)
        '
        Return shp
        '
    End Function
    '

    '
    ''' <summary>
    ''' This function will write strTextToWrite into the TextFrame of the shape (shp). It will use the specified
    ''' style strStyleName. The default is 'aa_waterMarkText_sec'. If the specified style cannot be found in the document,
    ''' then the default style is used. Note that strAlignment 'left', 'right' and 'centre' will override the style
    ''' settings for paragraph alignment
    ''' </summary>
    ''' <param name="shp"></param>
    ''' <param name="strTextToWrite"></param>
    ''' <param name="strStyleName"></param>
    ''' <param name="strAlignment"></param>
    Public Sub waterMark_shape_writeText(ByRef shp As Word.Shape, strTextToWrite As String, Optional strStyleName As String = "aa_waterMarkText_sec", Optional strAlignment As String = "")
        Dim myStyle As Word.Style
        'Dim lt As Word.ListTemplate
        Dim lg As Word.ListGallery
        '
        myStyle = Nothing
        '
        '**** Apply a specific ListTemplate from the Numbering Library in 'Numbering and Bullets'
        '**** to circumvent the seemingly random application of list items to the water marks.
        '**** In this case we explicitly specify 'none' (Nothing). Item 1 in the gallery is the
        '**** simple arabic numbers 1., 2., 3. etc. Item 4 is 'A., B., C.'
        '
        lg = glb_get_wrdApp.ListGalleries.Item(WdListGalleryType.wdNumberGallery)
        'lt = lg.ListTemplates.Item(7)
        'lt = lg.ListTemplates("None")
        'lt = Nothing
        '
        '
        Try
            myStyle = glb_get_wrdActiveDoc.Styles.Item(strStyleName)
            shp.TextFrame.TextRange.Style = myStyle
            shp.TextFrame.TextRange.Text = strTextToWrite
            shp.TextFrame.TextRange.Style = myStyle
            Select Case strAlignment
                Case "left"
                    shp.TextFrame.TextRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
                Case "centre"
                    shp.TextFrame.TextRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                Case "right"
                    shp.TextFrame.TextRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
                Case Else

            End Select

            'shp.TextFrame.TextRange.ListFormat.ApplyNumberDefault()
            'shp.TextFrame.TextRange.ListFormat.ApplyListTemplate(lt)
            '
        Catch ex As Exception
            myStyle = glb_get_wrdActiveDoc.Styles.Item("aa_waterMarkText_sec")
            shp.TextFrame.TextRange.Text = strTextToWrite
            shp.TextFrame.TextRange.Style = myStyle
            'shp.TextFrame.TextRange.ListFormat.ApplyListTemplate(lt)
            '
        End Try
        '
    End Sub
    '
    ''' <summary>
    ''' Takes the header table (tbl) and returns a collection containing dimensions and location
    ''' information that can be used to align a shape to the Table. That is, the shape will overlay the Table.
    ''' If the shape shp is included (i.e. not nothing), then the shape will be aligned, 'left', 'right' or 'centre'
    ''' relative to the underlying table tbl
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="strAlignment"></param>
    ''' <param name="shp"></param>
    Public Function waterMark_shape_align(ByRef tbl As Word.Table, strAlignment As String, Optional ByRef shp As Word.Shape = Nothing) As Collection
        Dim tblWidth, shpLeft, shpTop, shpWidth, shpHeight, headerDistance As Single
        Dim tblLeftEdgeActual As Single
        Dim lstOfDimensions As New Collection()
        Dim sect As Word.Section
        '
        sect = tbl.Range.Sections.Item(1)
        headerDistance = sect.PageSetup.HeaderDistance
        '
        tblWidth = Me.glb_tbls_getTableWidth(tbl)
        tblLeftEdgeActual = sect.PageSetup.LeftMargin + tbl.Rows.Item(1).LeftIndent
        '
        shpWidth = tblWidth / 2
        Select Case strAlignment
            Case "right"
                shpLeft = tblLeftEdgeActual + tblWidth - shpWidth
            Case "centre"
                shpLeft = tblLeftEdgeActual + (tblWidth - shpWidth) / 2
            Case "left"
        End Select
        shpTop = headerDistance
        shpHeight = tbl.Rows.Item(1).Height
        '
        'Store the shp dimensions and location information
        lstOfDimensions.Add(shpLeft, "left")
        lstOfDimensions.Add(shpTop, "top")
        lstOfDimensions.Add(shpWidth, "width")
        lstOfDimensions.Add(shpHeight, "height")
        '
        'If a shape is included, then we'll align it and the text in the textframe
        If Not IsNothing(shp) Then
            Select Case strAlignment
                Case "right"
                    shp.TextFrame.TextRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
                Case "centre"
                    shp.TextFrame.TextRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                Case "left"
                    shp.TextFrame.TextRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
            End Select
            shp.Left = shpLeft
            shp.Top = shpTop
            shp.Width = shpWidth
            shp.Height = shpHeight
        End If
        '
        Return lstOfDimensions
        '
    End Function
    '
    Public Function waterMark_shape_align(strAlignment As String, ByRef shp As Word.Shape) As Collection
        Dim lstOfDimensions As New Collection()
        '
        Select Case strAlignment
            Case "left"
                shp.TextFrame.TextRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
            Case "centre"
                shp.TextFrame.TextRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
            Case "right"
                shp.TextFrame.TextRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
        End Select
        '
        lstOfDimensions.Add(shp.Left, "left")
        lstOfDimensions.Add(shp.Top, "top")
        lstOfDimensions.Add(shp.Width, "width")
        lstOfDimensions.Add(shp.Height, "height")
        '
        Return lstOfDimensions
        '
    End Function


    Public Function waterMark_textBoxTo_Header(strCaption As String,
                                          ByRef hf As Word.HeaderFooter, left As Single, top As Single, width As Single, height As Single,
                                          strName As String, fontSize As Single, Optional strAlignment As String = "right",
                                          Optional strCaptionColour As Long = -1, Optional ByRef rngAnchor As Word.Range = Nothing) As Word.Shape
        Dim shp As Word.Shape
        Dim myStyle As Word.Style
        Dim objStylesMgr As New cStylesManager()
        'Dim drCell As Word.Cell
        'Dim para As Word.Paragraph
        'Dim left, top, width, height, tblWidth As Single
        'Dim rngAnchor As Word.Range
        '
        'We'll put the anchor in the first cell of the header table
        'drCell = headerTable.Range.Cells.Item(2)
        'rngAnchor = drCell.Range

        'para = drCell.Range.Paragraphs.Item(1)
        'rngAnchor = para.Range
        'rngAnchor.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        'para = hf.Range.Paragraphs.Last
        'rngAnchor = para.Range
        'rngAnchor.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        'tblWidth = Me.glb_tbls_getTableWidth(headerTable)
        '
        myStyle = glb_get_wrdActiveDoc.Styles.Item("aa_waterMarkText_sec")
        myStyle.Font.TextColor.RGB = Me._glb_colour_WaterMark_Grey_sec
        If strCaptionColour >= 0 Then
            myStyle.Font.TextColor.RGB = strCaptionColour
        End If
        '
        'strAlignment = "centre"
        '
        '
        shp = hf.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0, 0, 150, 20, rngAnchor)
        shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
        shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
        shp.LockAnchor = True
        '
        shp.Left = left
        shp.Top = top
        shp.Width = width
        shp.Height = height
        '
        '
        'shp.Left = left
        'shp.Top = top
        'shp.Width = width
        '
        '
        Select Case strAlignment
            Case "right"
                myStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
            Case "centre"
                myStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
            Case "left"
                myStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
            Case Else
                myStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
        End Select
        '
        'shp.LockAnchor = True
        '
        shp.TextFrame.TextRange.Style = myStyle
        shp.TextFrame.MarginTop = 0.0
        shp.TextFrame.MarginBottom = 0.0
        shp.TextFrame.MarginRight = 0.0
        shp.TextFrame.MarginLeft = 0.0
        shp.TextFrame.TextRange.Text = strCaption
        shp.TextFrame.TextRange.Font.Size = fontSize
        'shp.TextFrame.TextRange.Font.Color = strCaptionColour
        shp.Name = strName
        '
        shp.Fill.ForeColor.RGB = RGB(255, 0, 0)
        shp.Fill.BackColor.RGB = RGB(255, 0, 0)
        'shp.Line.Visible = False
        shp.Fill.Transparency = 0.5
        shp.Line.Visible = False
        '
        '
        Return shp
        '
    End Function
    '
    Public Sub waterMark_to_Body(strName As String, ByRef sect As Word.Section, ByRef objBB As cBBlocksHandler)
        'This method will insert WaterMarks into the general Body
        'sections
        Dim hf As HeaderFooter
        Dim pgIsLandscape As Boolean
        Dim rng As Word.Range
        Dim rngOfBlock As Word.Range
        Dim shp As Word.Shape
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        Dim para As Word.Paragraph
        Dim objCpMgr As cCoverPageMgr
        Dim hMarginWidth, vMarginHeight, m, k As Single
        Dim height, left, top, width As Single
        '
        'Let's define the writeable area of the current section
        hMarginWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
        vMarginHeight = sect.PageSetup.PageHeight - sect.PageSetup.TopMargin - sect.PageSetup.BottomMargin
        m = 1.5
        k = 1.0
        '
        objCpMgr = New cCoverPageMgr()
        pgIsLandscape = Me.isLandscape(sect)
        '
        'Do Security Water marks in the Header Table
        '
        If strName Like "*_sec" Then
            For Each hf In sect.Headers
                If hf.Exists Then
                    rng = hf.Range
                    If rng.Tables.Count <> 0 Then
                        tbl = rng.Tables.Item(1)
                        drCell = tbl.Range.Cells.Item(2)
                        para = drCell.Range.Paragraphs.Item(1)
                        rng = para.Range
                        rng.Collapse(WdCollapseDirection.wdCollapseStart)
                        '
                        '
                        drCell = tbl.Range.Cells.Item(1)
                        para = drCell.Range.Paragraphs.Item(1)
                        rng = para.Range
                        rng.Collapse(WdCollapseDirection.wdCollapseStart)
                        'para = drCell.Range.Paragraphs.Item(1)
                        'para.SpaceBefore = 4.0
                        'para.RightIndent = 6.0
                        'para.Range.Text = "CABINET-IN-CONFIDENCE"
                        'para.Range.Font.Color = RGB(255, 0, 0)
                        'para.Range.Font.Size = 10
                        'txtBx = hf.Shapes.AddTextbox(MsoTextOrientation.msoTextOrientationHorizontal, 20, 20, 20, 40, rng)
                        'leftEdge = objCpMgr.objGlobals.glb_math_MillimetersToPoints(objCpMgr.objGlobals._glb_header_leftEdge)
                        '
                        'shp = Me.waterMark_textBoxTo_Header("OFFICIAL:Sensitive", hf, tbl.PreferredWidth / 4, 0.0, tbl.PreferredWidth / 2,
                        'tbl.Rows.Item(1).Height, rng, strName, 12.0, RGB(255, 0, 0))

                        GoTo loop1
                        '
                        '
                        rngOfBlock = objBB.insertBuildingBlockFromDefaultLibToRange(strName, "waterMarks", rng)
                        If rngOfBlock.ShapeRange.Count <> 0 Then
                            shp = rngOfBlock.ShapeRange.Item(1)
                            shp.Name = strName
                            'If shp.Name Like "*_sec" Then shp.Fill.ForeColor.RGB = RGB(147, 147, 147)
                            If shp.Name Like "*_sec" Then shp.Fill.ForeColor.RGB = RGB(255, 0, 0)

                            '
                            shp.ConvertToInlineShape()
                        End If
loop1:
                    End If
                End If
            Next hf
        End If
        '
        'Do the Release Status Water marks in the background
        '
        If strName Like "*_stat" Then
            For Each hf In sect.Headers
                If hf.Exists Then
                    rng = hf.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                    '
                    If sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait Then
                        left = sect.PageSetup.PageWidth / 8
                        top = sect.PageSetup.PageHeight / 3
                        width = 3 * Me.glb_get_widthBetweenMargins(sect) / 3
                        height = width / 2
                    Else

                    End If
                    '
                    'shp = Me.waterMark_textBoxTo_Header("DRAFT", hf, left, top, width,
                    'height, rng, strName, 80.0)
                    'shp = Me.waterMark_textBoxTo_Header("DRAFT", hf, left, top, width,
                    'height, rng, strName, 80.0, RGB(255, 0, 0))
                    '
                    'shp.Rotation = -45







                    '
                    '
                    GoTo loop2
                    '
                    rngOfBlock = objBB.insertBuildingBlockFromDefaultLibToRange(strName, "waterMarks", rng)
                    If rngOfBlock.ShapeRange.Count <> 0 Then
                        shp = rngOfBlock.ShapeRange.Item(1)
                        shp.Name = strName
                        shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
                        shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
                        '
                        Call shp.ScaleWidth(0.8, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromTopLeft)
                        'shp.Rotation = 0.0
                        'shp.Left = (sect.PageSetup.PageWidth - shp.Width) / 2
                        shp.Left = sect.PageSetup.LeftMargin + ((hMarginWidth - shp.Width) / (k + 1.0))
                        '
                        'Fidn the vertical position
                        'shp.Top = 400.0
                        shp.Top = sect.PageSetup.TopMargin + ((vMarginHeight - shp.Height) / (m + 1))
                        'shp.Left = 87.2
                        '
                        shp.Fill.ForeColor.RGB = RGB(147, 147, 147)
                        'shp.Fill.ForeColor.RGB = RGB(255, 0, 0)

                        shp.Fill.Transparency = 0.2
                        '
                        'If ob
                        If objCpMgr.cp_Bool_IsCoverPage(sect) Then
                            'shp.Top = 625.3
                            'shp.Top = 400.0
                            'shp.LockAspectRatio = True
                            'Call shp.ScaleWidth(0.5, MsoTriState.msoFalse, MsoScaleFrom.msoScaleFromTopLeft)
                            'shp.Rotation = 0#
                            'shp.Left = (sect.PageSetup.PageWidth - shp.Width) / 2
                            'shp.Fill.Transparency = 0.5
                            'shp.Fill.BackColor.RGB = RGB(255, 0, 0)

                        End If
                    End If
loop2:
                End If
            Next
        End If
        Exit Sub

    End Sub
    '
    ''' <summary>
    ''' This method will align all sec shapes in the Header of the specified section to eithe
    ''' 'left', 'centre' or 'right'
    ''' </summary>
    ''' <param name="strAlignment"></param>
    ''' <param name="sect"></param>
    Public Function waterMark_sec_Alignment(strAlignment As String, ByRef sect As Word.Section) As Boolean
        Dim hf As HeaderFooter
        Dim rng As Word.Range
        Dim shp As Word.Shape
        Dim tbl As Word.Table
        Dim rslt As Boolean
        '
        rslt = False
        '
        Try
            For Each hf In sect.Headers
                If hf.Exists Then
                    rng = hf.Range
                    If hf.Range.Tables.Count <> 0 Then
                        tbl = rng.Tables.Item(1)
                        For Each shp In hf.Range.ShapeRange
                            If shp.Name Like "*_aa_sec" Then
                                Me.waterMark_shape_align(tbl, strAlignment, shp)
                                rslt = True
                            End If
                        Next
                    End If
                End If
            Next
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method will align all sec shapes in the Header of the specified document (myDoc) to either
    ''' 'left', 'centre' or 'right'.. Since the header shape covers the entire header table, all we have
    ''' to do is adjust the paragraph a;lignment
    ''' </summary>
    ''' <param name="strAlignment"></param>
    ''' <param name="myDoc"></param>
    Public Function waterMark_sec_Alignment(strAlignment As String, ByRef myDoc As Word.Document) As Boolean
        Dim myStyle As Word.Style
        Dim rslt As Boolean
        '
        rslt = False
        '
        Try
            myStyle = myDoc.Styles.Item(Me.glb_var_style_waterMark_sec)
            Select Case strAlignment
                Case "left"
                    myStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
                Case "centre"
                    myStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                Case "right"
                    myStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight

            End Select
            rslt = True
            '
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method will accept strName (the old waterMark vector shape name and use it to generate an internal
    ''' strCaption which is written to the text of the Shape (shp) that is inserted as an overlay above the
    ''' header tables in all sections. If a section has no header table (e.g. Cover Page) then the Shape
    ''' and its associated text caption is inserted into the header of the section. The shape and text alignment
    ''' is specified by strAlignment ('left', 'right' and 'centre'). text colour is the default 
    ''' glb._glb_colour_WaterMark_Grey_sec, unless overwitten by an actual value placed in txtColour
    ''' </summary>
    ''' <param name="strName"></param>
    ''' <param name="sect"></param>
    ''' <param name="strAlignment"></param>
    ''' <param name="txtColour"></param>
    ''' <returns></returns>
    Public Function waterMark_sec_toBody(strSectionTag As String, strName As String, ByRef sect As Word.Section, Optional strAlignment As String = "right", Optional txtColour As Long = -1) As Word.Shape
        Dim hf As Word.HeaderFooter
        Dim objStylesMgr As New cStylesManager()
        Dim shp As Word.Shape
        Dim myStyle As Word.Style
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim tblWidth, left, top, width, height, tblLeftEdgeActual As Single
        Dim lstOfDimensions As Collection
        Dim strCaption As String
        '
        strCaption = Me.waterMark_get_securityCaption(strName)
        'txtColour = RGB(255, 0, 0)
        '
        shp = Nothing
        '
        'If this is not a security message (i.e. for the header), then don't do this
        If Not strName Like "*_aa_sec" Then GoTo finis
        'myStyle = Me.glb_get_wrdActiveDoc.Styles.Item(Me.glb_var_style_waterMark_sec)
        '
        'This will get the style and/or create the style.. If it gets it, then the
        'style is as is.. If it creates it, then it is setup to its default.. Suitable for
        'an addin approach
        myStyle = objStylesMgr.style_getCreateRefresh_waterMark_sec(sect.Range.Document)
        '
        Select Case strAlignment
            Case "left"
                myStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
            Case "right"
                myStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
            Case "centre"
                myStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
            Case Else
                'Default condition
                myStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
                '
        End Select
        '
        myStyle.Font.Color = _glb_colour_WaterMark_Grey_sec
        If txtColour >= 0 Then myStyle.Font.Color = txtColour
        '
        myStyle.ParagraphFormat.SpaceBefore = 0.0

        '
        For Each hf In sect.Headers
            If hf.Exists Then
                rng = hf.Range
                If hf.Range.Tables.Count <> 0 Then
                    tbl = rng.Tables.Item(1)
                    'tblWidth = Me.glb_hfs_getHeaderTableWidth(sect)
                    tblWidth = Me.glb_tbls_getTableWidth(tbl)
                    tblLeftEdgeActual = sect.PageSetup.LeftMargin + tbl.Rows.Item(1).LeftIndent
                    '
                    'drCell = tbl.Range.Cells.Item(2)
                    'drCell.Range.Text = ""
                    'drCell.Range.Style = myStyle
                    '
                    'lstOfDimensions = Me.waterMark_shape_align(tbl, strAlignment)
                    '
                    'We will create a Security status shape that covers the entire header table. By doing this
                    'we can adjust alignment by chnaging the style paragraph alignment
                    lstOfDimensions = New Collection()
                    lstOfDimensions.Add(tblLeftEdgeActual, "left")
                    lstOfDimensions.Add(sect.PageSetup.HeaderDistance, "top")
                    lstOfDimensions.Add(tblWidth, "width")
                    lstOfDimensions.Add(tbl.Rows.Item(1).Height, "height")

                    '
                    rng = hf.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                    '
                    '
                    shp = Me.waterMark_insertShape_toRange(rng, hf, lstOfDimensions)
                    shp.Name = shp.Name + "_aa_sec"                                                              'Workaround.. The shape must be named outside the above routine.. I don't know why its
                    Me.waterMark_shape_writeText(shp, strCaption)
                    'Me.waterMark_shape_align(strAlignment, shp)
                    '
                    If txtColour >= 0 Then myStyle.Font.Color = txtColour
                    '
                Else
                    'No Header Table, so this is a Cover Page
                    rng = hf.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                    '
                    left = 150
                    top = 150
                    width = 150
                    height = 40
                    '
                    lstOfDimensions = New Collection()
                    lstOfDimensions.Add(sect.PageSetup.LeftMargin, "left")
                    'lstOfDimensions.Add(110, "top")
                    lstOfDimensions.Add(86, "top")

                    lstOfDimensions.Add(311, "width")
                    lstOfDimensions.Add(34, "height")

                    '
                    shp = Me.waterMark_insertShape_toRange(rng, hf, lstOfDimensions)
                    shp.Name = shp.Name + "_aa_sec"                                                              'Workaround.. The shape must be named outside the above routine.. I don't know why its
                    Me.waterMark_shape_writeText(shp, strCaption)
                    shp.TextFrame.TextRange.Font.Size = 18
                    shp.TextFrame.TextRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
                    '
                    If txtColour >= 0 Then myStyle.Font.Color = txtColour
                    '
                    'Me.waterMark_shape_align(tbl, strAlignment, shp)

                End If
            End If
        Next
        '
        Select Case strSectionTag
            Case "tag_coverPage"
                shp.Left = sect.PageSetup.LeftMargin
                If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
                    shp.Top = 87
                End If
            Case "tag_letter"
            Case "tag_contactsPage-Front", "tag_contactsPage-Back", "tag_partBanner", "tag_appendixPart", "tag_contactsPage-Back"
            Case Else
        End Select
        '
finis:
        '
        Return shp
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will accept strName (the old waterMark vector shape name and use it to generate an internal
    ''' strCaption which is written to the text of the Shape (shp) that is inserted as an overlay above the
    ''' header tables in all sections. If a section has no header table (e.g. Cover Page) then the Shape
    ''' and its associated text caption is inserted into the header of the section. The shape and text alignment
    ''' is specified by strAlignment ('left', 'right' and 'centre'). text colour is the default 
    ''' glb._glb_colour_WaterMark_Grey_sec, unless overwitten by an actual value placed in txtColour
    ''' </summary>
    ''' <param name="strName"></param>
    ''' <param name="sect"></param>
    ''' <param name="strAlignment"></param>
    ''' <param name="txtColour"></param>
    ''' <returns></returns>
    Public Function waterMark_stat_toBody(ByRef strSectionTag As String, strName As String, ByRef sect As Word.Section, Optional strAlignment As String = "", Optional txtColour As Long = -1) As Word.Shape
        Dim objCpMgr As cCoverPageMgr
        Dim objStylesMgr As New cStylesManager()
        Dim objWrkAround As New cWorkArounds()
        Dim hf As Word.HeaderFooter
        Dim shp As Word.Shape
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim rightPadding, left, top, width, height, leftEdgeActual As Single
        Dim lstOfDimensions As Collection
        Dim strCaption As String
        Dim strDoRotation As String
        Dim strWrkPos As String
        '
        strWrkPos = "footer"
        '
        strCaption = Me.waterMark_get_statusCaption(strName)
        rightPadding = 0
        strDoRotation = "0"
        'txtColour = RGB(255, 0, 0)
        '
        shp = Nothing
        '
        'If this is not a security message (i.e. for the header), then don't do this
        If Not strName Like "*_stat" Then GoTo finis
        '
        'This will get the style and/or create the style.. If it gets it, then the
        'style is as is.. If it creates it, then it is setup to its default... Suitable for
        'an addin approach
        objStylesMgr.style_getCreateRefresh_waterMark_stat(sect.Range.Document)
        '
        'myStyle = Me.glb_get_wrdActiveDoc.Styles.Item(Me.glb_var_style_waterMark_stat)
        'myStyle.Font.Color = _glb_colour_WaterMark_Grey_stat
        'myStyle.ParagraphFormat.SpaceBefore = 0.0
        '
        For Each hf In sect.Headers
            'hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
            If hf.Exists Then
                rng = hf.Range
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                If hf.Range.Tables.Count <> 0 Then
                    tbl = hf.Range.Tables.Item(1)
                    leftEdgeActual = sect.PageSetup.LeftMargin + tbl.Rows.Item(1).LeftIndent
                    height = 80
                    height = 44

                    Select Case strSectionTag
                        Case "tag_aa_stn_letter", "tag_aa_stn_memo"
                            width = glb_tbls_getTableWidth(tbl)
                            left = leftEdgeActual
                            '
                            top = sect.PageSetup.HeaderDistance
                            'If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then top = 141
                            strDoRotation = "0"
                            '
                        Case "tag_contactsPage-Front", "tag_contactsPage-Back"

                            'Place the status water Mark in an empty place on the page
                            width = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
                            left = sect.PageSetup.LeftMargin
                            top = 250
                            If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then top = 141
                            strDoRotation = "0"
                        Case "tag_partBanner", "tag_appendixPart"
                            '
                            strWrkPos = "footer"
                            'strWrkPos = "body"
                            '
                            Select Case strWrkPos
                                Case "body"
                                    'This code placed the Document Status in the body of the banner
                                    '
                                    left = leftEdgeActual
                                    width = glb_hfs_getHeaderTableWidth(hf.Range.Sections.Item(1))
                                    height = glb_math_MillimetersToPoints(28.22)
                                    '
                                    top = glb_math_MillimetersToPoints(31.7)
                                    If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
                                        top = glb_math_MillimetersToPoints(19.4)
                                    End If
                                    '
                                    strDoRotation = "0"
                                    strAlignment = "left"

                                Case "footer"
                                    'This code placed the Document Status in the bottom left hand corner of the
                                    'footer
                                    '
                                    left = leftEdgeActual
                                    width = glb_hfs_getHeaderTableWidth(hf.Range.Sections.Item(1))

                                    '
                                    'The bottom margin of the banners and appendix part are 56pts instead of 66pts... For reasons lost in the
                                    'mists of time.... So I need to make a 10 pt adjustment for the Portrait page. The adjustment for landscape needs to
                                    'be greater
                                    '
                                    top = sect.PageSetup.PageHeight - sect.PageSetup.BottomMargin + (height - glb_hfs_getFooterTable_Height_Nominal()) / 2 - 10
                                    If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
                                        top = sect.PageSetup.PageHeight - sect.PageSetup.BottomMargin - (height - glb_hfs_getFooterTable_Height_Nominal()) / 4 - 32
                                    End If

                                    strDoRotation = "0"
                                    strAlignment = "left"

                            End Select
                            '
                        Case Else
                            '
                            left = leftEdgeActual
                            width = glb_hfs_getHeaderTableWidth(hf.Range.Sections.Item(1))
                            'height = 44
                            '
                            top = sect.PageSetup.PageHeight - sect.PageSetup.BottomMargin + (height - glb_hfs_getFooterTable_Height_Nominal()) / 2
                            If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
                                top = sect.PageSetup.PageHeight - sect.PageSetup.BottomMargin - (height - glb_hfs_getFooterTable_Height_Nominal()) / 4
                            End If

                            strDoRotation = "0"
                            strAlignment = "left"

                            'Case "tag_glossary_Chpt"


                            '**** The rest of the body of the report. We will right align the status message
                            'and padd from the right hand (top after rotation) side
                            '
                            'width = sect.PageSetup.PageHeight - sect.PageSetup.HeaderDistance - sect.PageSetup.FooterDistance
                            'left = glb_hfs_getHFTableEdge(hf.Range.Sections.Item(1), "header_leftEdge") + height / 2 - width / 2 - 20 + 2
                            'top = sect.PageSetup.HeaderDistance + width / 2 - height / 2
                            'strAlignment = "right"
                            '
                            'Select Case sect.PageSetup.Orientation
                            'Case WdOrientation.wdOrientPortrait
                            'rightPadding = 186
                            'Case WdOrientation.wdOrientLandscape
                            'rightPadding = 156
                            'End Select
                            '
                            'strDoRotation = "-90"
                            'oRotation = "0"

                    End Select
                    '
                    lstOfDimensions = New Collection()
                    lstOfDimensions.Add(left, "left")
                    lstOfDimensions.Add(top, "top")
                    lstOfDimensions.Add(width, "width")
                    lstOfDimensions.Add(height, "height")
                    '
                    shp = Me.waterMark_insertShape_toRange(rng, hf, lstOfDimensions)
                    shp.TextFrame.MarginRight = rightPadding
                    '
                    shp.Name = shp.Name + "_aa_stat"                                                              'Workaround.. The shape must be named outside the above routine.. I don't know why its
                    Me.waterMark_shape_writeText(shp, strCaption, Me.glb_var_style_waterMark_stat, strAlignment)
                    'Me.waterMark_shape_writeText(shp, "Test", Me.glb_var_style_waterMark_stat, strAlignment)

                    'shp.TextFrame.TextRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight

                    '
                    Select Case strDoRotation
                        Case "0"
                        Case "-90"
                            shp.Rotation = -90
                        Case Else
                    End Select
                    '
                Else
                    'No Header Table, so this is a Cover Page
                    objCpMgr = New cCoverPageMgr()
                    rng = hf.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                    '
                    top = 32
                    width = 300
                    height = 60
                    '
                    'Cover page has varying margins and no header table, so we have to use
                    'absolutes held in cCoverPageMgr
                    left = objCpMgr.cp_get_RightEdge(sect) - width
                    'left = sect.PageSetup.PageWidth - sect.PageSetup.RightMargin - width '499pt for landscape

                    '
                    lstOfDimensions = New Collection()
                    lstOfDimensions.Add(left, "left")
                    lstOfDimensions.Add(top, "top")
                    lstOfDimensions.Add(width, "width")
                    lstOfDimensions.Add(height, "height")

                    '
                    shp = Me.waterMark_insertShape_toRange(rng, hf, lstOfDimensions)
                    shp.Name = shp.Name + "_aa_stat"                                                              'Workaround.. The shape must be named outside the above routine.. I don't know why its
                    Me.waterMark_shape_writeText(shp, strCaption, Me.glb_var_style_waterMark_stat)
                    shp.TextFrame.TextRange.Font.Size = 36
                    shp.TextFrame.TextRange.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
                    '
                    'If txtColour >= 0 Then myStyle.Font.Color = txtColour
                    '
                    'Me.waterMark_shape_align(tbl, strAlignment, shp)

                End If
                '
                '
                'shp.Rotation = -45
            End If
        Next
        '
        Try
            Select Case strSectionTag
                Case "tag_coverPage"
                'shp.Left = sect.PageSetup.LeftMargin
                'shp.Width = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
                Case "tag_letter"
                Case "tag_contactsPage-Front", "tag_contactsPage-Back", "tag_partBanner", "tag_appendixPart", "tag_contactsPage-Back"
                Case "tag_contactsPage-Front", "tag_contactsPage-Back"
                    'shp.Rotation = 0
                    'shp.Width = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
                    'shp.Left = sect.PageSetup.LeftMargin + shp.Width / 2
                    'shp.Top = 250
                Case Else
            End Select

        Catch ex As Exception

        End Try
        '
finis:
        objWrkAround.wrk_fix_forCursorRace()
        '
        Return shp
        '
    End Function


    Public Sub waterMark_to_Body(strName As String, ByRef sect As Word.Section, Optional strAlignment As String = "right", Optional txtColour As Long = -1)
        'This method will insert WaterMarks into the general Body
        'sections
        Dim lstOfDimensions As Collection
        Dim myStyle As Word.Style
        Dim hf As HeaderFooter
        Dim rng As Word.Range
        Dim rngAnchor As Word.Range
        Dim shp As Word.Shape
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        Dim para As Word.Paragraph
        'Dim objCpMgr As cCoverPageMgr
        'Dim hMarginWidth, vMarginHeight, m, k As Single
        Dim left, top, width, height, tblWidth As Single
        Dim strCaption As String
        '
        If txtColour = -1 Then txtColour = Me._glb_colour_WaterMark_Grey_sec
        strCaption = ""
        '
        'Do Security Water marks in the Header Table
        '
        If strName Like "*_sec" Then
            For Each hf In sect.Headers
                If hf.Exists Then
                    rng = hf.Range
                    If hf.Range.Tables.Count <> 0 Then
                        tbl = rng.Tables.Item(1)
                        tblWidth = Me.glb_hfs_getHeaderTableWidth(sect)
                        '
                        myStyle = Me.glb_get_wrdActiveDoc.Styles.Item(Me.glb_var_style_waterMark_sec)
                        myStyle.ParagraphFormat.SpaceBefore = 0.0
                        drCell = tbl.Range.Cells.Item(2)
                        drCell.Range.Text = ""
                        drCell.Range.Style = myStyle
                        '
                        lstOfDimensions = Me.waterMark_shape_align(tbl, strAlignment)
                        '
                        rng = hf.Range
                        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                        '
                        strCaption = Me.waterMark_get_securityCaption(strName)
                        '
                        shp = Me.waterMark_insertShape_toRange(rng, hf, lstOfDimensions)
                        shp.Name = shp.Name + "_aa_sec"                                                              'Workaround.. The shape must be named outside the above routine.. I don't know why its
                        Me.waterMark_shape_writeText(shp, strCaption)
                        Me.waterMark_shape_align(tbl, strAlignment, shp)
                        '
                    Else
                        rng = hf.Range
                        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                        '
                        strCaption = Me.waterMark_get_securityCaption(strName)
                        '
                        left = 150
                        top = 150
                        width = 150
                        height = 40
                        '
                        lstOfDimensions = New Collection()
                        lstOfDimensions.Add(230.25, "left")
                        lstOfDimensions.Add(45, "top")
                        lstOfDimensions.Add(311, "width")
                        lstOfDimensions.Add(34, "height")

                        '
                        shp = Me.waterMark_insertShape_toRange(rng, hf, lstOfDimensions)
                        shp.Name = shp.Name + "_aa_sec"                                                              'Workaround.. The shape must be named outside the above routine.. I don't know why its
                        Me.waterMark_shape_writeText(shp, strCaption)
                        shp.TextFrame.TextRange.Font.Size = 18
                        'Me.waterMark_shape_align(tbl, strAlignment, shp)

                    End If
                End If
            Next
        End If
        '
        GoTo loop1
        '
        If strName Like "*_sec" Then
            For Each hf In sect.Headers
                If hf.Exists Then
                    rng = hf.Range
                    If rng.Tables.Count <> 0 Then
                        tbl = rng.Tables.Item(1)
                        tblWidth = Me.glb_hfs_getHeaderTableWidth(sect)
                        'We'll put the anchor in the first cell of the header table
                        drCell = tbl.Range.Cells.Item(2)
                        para = drCell.Range.Paragraphs.Item(1)
                        para.Style = Me.glb_get_wrdActiveDoc.Styles.Item(Me.glb_var_style_waterMark_sec)
                        rngAnchor = para.Range
                        rngAnchor.Collapse(WdCollapseDirection.wdCollapseStart)
                        '
                        'Now get the left, top, width and height measurements (pts) for
                        'the header shape that we will use to carry the security status information
                        'If we include the shape (Optional) it will also align the shape and any
                        'existing text
                        '
                        lstOfDimensions = Me.waterMark_shape_align(tbl, strAlignment)
                        '
                        strCaption = Me.waterMark_get_securityCaption(strName)
                        rngAnchor.Text = strCaption
                        GoTo loop1
                        '
                        If Not strCaption = "" Then
                            shp = Me.waterMark_insertShape_toRange(rngAnchor, hf, lstOfDimensions)
                            shp.Name = shp.Name + "_aa_sec"                                                              'Workaround.. The shape must be named outside the above routine.. I don't know why its
                            Me.waterMark_shape_writeText(shp, strCaption)
                            Me.waterMark_shape_align(tbl, strAlignment, shp)
                            'shp = Me.waterMark_textBoxTo_Header(strCaption, hf, left, top, width, height, strName, 12.0, strAlignment, RGB(255, 0, 0), rngAnchor)
                        End If
                        '

                        '
                    End If
                End If
            Next hf
        End If
        '
loop1:
        '
        'Do the Release Status Water marks in the background
        '
        If strName Like "*_stat" Then
            For Each hf In sect.Headers
                If hf.Exists Then
                    rngAnchor = hf.Range
                    rngAnchor.Collapse(WdCollapseDirection.wdCollapseEnd)
                    '
                    strCaption = ""
                    Select Case strName
                        Case "waterMark_aa_draft_stat"
                            strCaption = "DRAFT"
                        Case "waterMark_aa_draftOnly_stat"
                            strCaption = "DRAFT ONLY"
                    End Select
                    '
                    If sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait Or sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
                        height = 80
                        'width = sect.PageSetup.PageHeight - sect.PageSetup.TopMargin - sect.PageSetup.BottomMargin
                        width = sect.PageSetup.PageHeight - sect.PageSetup.HeaderDistance - sect.PageSetup.FooterDistance

                        left = glb_hfs_getHFTableEdge(hf.Range.Sections.Item(1), "header_leftEdge") + height / 2 - width / 2 - 20 + 2
                        'left = sect.PageSetup.PageWidth / 8
                        'top = sect.PageSetup.PageHeight / 3
                        'top = sect.PageSetup.TopMargin + width / 2 - height / 2
                        top = sect.PageSetup.HeaderDistance + width / 2 - height / 2


                        '
                        lstOfDimensions = New Collection()
                        lstOfDimensions.Add(left, "left")
                        lstOfDimensions.Add(top, "top")
                        lstOfDimensions.Add(width, "width")
                        lstOfDimensions.Add(height, "height")
                        '
                        shp = Me.waterMark_insertShape_toRange(rngAnchor, hf, lstOfDimensions)
                        '
                        Select Case sect.PageSetup.Orientation
                            Case WdOrientation.wdOrientPortrait
                                shp.TextFrame.MarginBottom = 186
                            Case WdOrientation.wdOrientLandscape
                                shp.TextFrame.MarginLeft = 78
                        End Select
                        '
                        shp.Name = shp.Name + "_aa_stat"                                                              'Workaround.. The shape must be named outside the above routine.. I don't know why its
                        Me.waterMark_shape_writeText(shp, strCaption, Me.glb_var_style_waterMark_stat)
                        shp.Rotation = -90
                        'shp.Left = -sect.PageSetup.PageWidth / 2
                        'shp.Top = sect.PageSetup.PageHeight / 2
                    Else

                    End If
                    '
                    'shp = Me.waterMark_textBoxTo_Header("DRAFT", hf, tbl, rng, strName, 80.0, "centre")
                    'shp = Me.waterMark_textBoxTo_Header("DRAFT", hf, left, top, width,
                    'height, rng, strName, 80.0, RGB(255, 0, 0))
                    '
                    'shp.Rotation = -45
                End If
            Next
        End If
        Exit Sub

    End Sub

    Public Sub waterMark_to_Letter(strName As String, ByRef sect As Word.Section, ByRef objBB As cBBlocksHandler)
        'This method will insert WaterMarks into the Letter
        'sections
        Dim hf As HeaderFooter
        Dim rng As Word.Range
        Dim rngOfBlock As Word.Range
        Dim shp As Word.Shape
        '
        '
        hf = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        rng = hf.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
        'Set rngOfBlock = objBB.insertBuildingBlockFromDefaultLibToRange(strName, "waterMarks", rng)

        For Each hf In sect.Headers
            If hf.Exists Then
                rng = hf.Range
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                'Call rng.Move(wdParagraph, -1)
                rngOfBlock = objBB.insertBuildingBlockFromDefaultLibToRange(strName, "waterMarks", rng)
                If rngOfBlock.ShapeRange.Count <> 0 Then
                    shp = rngOfBlock.ShapeRange.Item(1)
                    shp.Name = strName
                    shp.LockAnchor = True
                    shp.LockAspectRatio = MsoTriState.msoCTrue
                    shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
                    shp.LeftRelative = 0#
                    shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
                    shp.Height = 7.5
                    shp.Top = 52.05
                    shp.Top = 52.05     'wasn't sticking
                    'Call shp.ScaleWidth(scaleFactor, msoFalse, msoScaleFromTopLeft)
                    'shp.top = 0#
                    'shp.left = (Me.spaceBetweenMargins(sect) - shp.width) / 2
                    'shp.left = (sect.PageSetup.PageWidth - shp.width) / 2 - sect.PageSetup.leftMargin
                    'shp.left = 0#
                    'shp.top = sect.PageSetup.PageHeight - delta + shp.height / 2
                    shp.Fill.ForeColor.RGB = RGB(0, 1, 0)
                End If
            End If
        Next hf
    End Sub
    '
    '
    Public Function spaceBetweenMargins(ByRef sect As Word.Section) As Single
        'This method will retrieve the measurement between the margins
        'in the current section.. Genherally used by methods that adjust
        'Table Widths
        '
        spaceBetweenMargins = sect.PageSetup.PageWidth - sect.PageSetup.RightMargin - sect.PageSetup.LeftMargin
        '
    End Function

    '
End Class

Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cCoverPageMgr
    Inherits cChptBase
    '
    Public strCoverPictureName As String
    Public strCoverPatternName As String
    Public sect As Word.Section
    Public lst_CoverPage As Collection
    '
    Public Sub New()
        MyBase.New()
        Me.strCoverPictureName = "cp_pict*"
        Me.strCoverPatternName = "cp_Empty_Pattern*"
        Me.strTagStyleName = "tag_coverPage"
        '
        'Will get measurements depending on the mode of the document
        Me.lst_CoverPage = objGlobals.glb_getDimensions_CoverPage_Prt()
        '
    End Sub
    '
    ''' <summary>
    ''' Since the cover page right margins vary and there is no header table to determine
    ''' where it might be relative to the left edge of the page (in pts)
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function cp_get_RightEdge(ByRef sect As Word.Section) As Single
        Dim rightEdge As Single
        '
        If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
            rightEdge = 800
        Else
            rightEdge = sect.PageSetup.PageWidth - 56
        End If
        '
        Return rightEdge
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will set the physical dimensions of the Cover Page. It calls
    ''' a Globals method that will test the Report Mode and will return the dimensions
    ''' for a Portrait or a Landscape page
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub cp_Set_Dimensions(ByRef sect As Word.Section)
        '
        objGlobals.glb_setDimensions(sect, Me.lst_CoverPage)
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will do a section by section check of myDoc. If it finds that
    ''' one of the sections is a CoverPage section then it will return true and
    ''' sect will be set to the section containing the cover page
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function cp_Bool_HasCoverPage(ByRef myDoc As Word.Document, ByRef sect As Word.Section) As Boolean
        Dim rslt As Boolean
        '
        rslt = False
        For Each sect In myDoc.Sections
            If Me.cp_Bool_IsCoverPage(sect) Then
                rslt = True
                Exit For
            End If
        Next
        '
        Return rslt
    End Function
    '
    '
    '
    ''' <summary>
    ''' This method will return nothing if the document does not have a Cover Page section... If it does it
    ''' will return the section
    ''' </summary>
    ''' <returns></returns>
    Public Function cp_get_CPSection() As Word.Section
        Dim objBnrMgr As New cChptBanner()
        Dim rslt As Boolean
        Dim objsectMgr As New cSectionMgr()
        Dim strTag As String
        Dim sect As Word.Section
        '
        sect = Nothing
        'strTag = objHfMgr.hf_tags_getTagStyleName(sect)'
        strTag = objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_coverPage)
        rslt = objsectMgr.sct_has_strTag(objGlobals.glb_get_wrdActiveDoc, strTag, sect)
        '
        Return sect
        '
    End Function


    '
    Public Sub cp_set_SelectionToTitle(ByRef sect As Word.Section)
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        Dim styl As Word.Style
        '
        rng = sect.Range
        '
        Try
            For Each para In rng.Paragraphs
                styl = para.Style
                If styl.NameLocal = "Cp Title" Then
                    rng = para.Range
                    rng.MoveEnd(WdUnits.wdCharacter, -1)
                    rng.Select()
                    Exit For
                End If
            Next
        Catch ex As Exception

        End Try

    End Sub
    '
    Public Function cp_Delete_CoverPage(ByRef myDoc As Word.Document) As String
        Dim sect As Word.Section
        Dim strMsg As String
        Dim objSectMgr As New cSectionMgr()
        Dim rng As Word.Range
        Dim ctrl As ContentControl
        '
        sect = Nothing
        strMsg = ""
        '
        If Me.cp_Bool_HasCoverPage(myDoc, sect) Then
            '.hasCoverPage will set sect to the section that has the cover page
            Me.sect = sect
            Me.cp_set_SelectionToTitle(sect)
            '
            'We first need to delete the Controls from the document
            'Globals.ThisAddin.Controls.Remove("Report Date")
            'Globals.ThisAddIn.Controls.Remove("Report")
            '
            For i As Integer = myDoc.ContentControls.Count To 1 Step -1
                ctrl = myDoc.ContentControls(i)
                If ctrl.Title = "Report Date" Or ctrl.Title = "Report" Then
                    ctrl.Delete()
                End If
            Next
            '
            rng = objGlobals.glb_get_wrdApp.Selection.Range
            sect = objSectMgr.sct_delete_Section(rng.Sections.Item(1))


        Else
            strMsg = "This Document does not have a Cover Page"
        End If
        '
        Return strMsg
        '
    End Function
    '
    Public Sub cp_convert_ToCoverPage(ByRef sect As Word.Section)
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objBrndMgr As New cBrandMgr()
        Dim objBBMgr As New cBBlocksHandler()
        Dim isPortrait As Boolean
        Dim myDoc As Word.Document
        Dim hf As Word.HeaderFooter
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim objGrfxMgr As New cGraphicsMgr()
        '
        tbl = Nothing
        myDoc = sect.Range.Document
        '
        isPortrait = True
        If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then isPortrait = False
        '
        objHFMgr.hf_hfs_deleteAll(sect)
        sect.PageSetup.DifferentFirstPageHeaderFooter = True
        '
        If isPortrait Then
            objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_CoverPage_Prt())
        Else
            objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_CoverPage_Lnd())
        End If

        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        rng = hf.Range
        rng.Style = myDoc.Styles.Item(Me.strTagStyleName)
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        objHFMgr.hf_Insert_BackShape(hf, rng)
        objBrndMgr.brnd_Insert_Logo(hf)
        '
        If isPortrait Then
            rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_Cpg_PictFilled", "CoverPage", rng)
            'Me.cp_rename_Shape(rng, "cp_pict_purplePattern_prt")
            Me.cp_rename_Shape(rng, "cp_pict_large")

            'objGrfxMgr.grfx_insert_ImageCP(objGlobals.glb_get_wrdSelRng.Sections.Item(1))
        Else
            rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_Cpg_PictFilled_Lnd", "CoverPage", rng)
            'Me.cp_rename_Shape(rng, "cp_pict_purplePattern_lnd")
            Me.cp_rename_Shape(rng, "cp_pict_large")

            'objGrfxMgr.grfx_insert_ImageCP(objGlobals.glb_get_wrdSelRng.Sections.Item(1))
        End If

        '
        rng = sect.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        Me.cp_insert_formattedTitleText(rng)
        'tbl = Me.cp_adjust_TitleTableBlockSize(tbl, strReportMode)
        'MyBase.Base_Paragraphs_Delete(rng, 1)

        '

    End Sub
    '
    '    
    Public Function cp_Insert_CoverPage(ByRef myDoc As Word.Document, strCoverPageType As String) As Word.Section
        Dim objRptMgr As New cReport()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim strReportMode As String
        Dim sect, sectNew, sectSelected As Word.Section
        Dim hf As HeaderFooter
        Dim rng As Word.Range
        Dim strOrientation As String
        Dim objWorkAround As New cWorkArounds()
        'Dim hf As Word.HeaderFooter
        'Dim shp As Word.Shape
        '
        '
        objGlobals.glb_get_wrdSel.Collapse(WdCollapseDirection.wdCollapseStart)
        sectNew = Nothing
        strReportMode = ""
        hf = Nothing
        '
        'Get the orientation of the page that contains the current selection
        '
        sectSelected = objGlobals.glb_get_wrdSect()
        '
        'Workaround as we transition to 2024 report which is really independent of modes
        'but the REport Mode has tendrils everywhere
        '
        strOrientation = "prt"
        strReportMode = objRptMgr.rpt_isPrt
        '
        If sectSelected.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
            strOrientation = "lnd"
            strReportMode = objRptMgr.rpt_isLnd
        End If
        '
        'strReportMode = objRptMgr.Rpt_Mode_Get()
        '
        sect = myDoc.Sections.First
        '
        If Not Me.cp_Bool_HasCoverPage(myDoc, sect) Then
            'If the report has a cover page, then sect is set to that cover page, but in this
            'instance, because there is no cover page it's value is set to the last section in the document
            '
            sect = myDoc.Sections.First
            '
            Select Case strOrientation
                Case "prt"
                    Me.lst_CoverPage = objGlobals.glb_getDimensions_CoverPage_Prt()
                    sectNew = MyBase.sct_insert_Section(False, sect, 6, "newPage", False, "prt", Me.lst_CoverPage)
                    Me.sect = sectNew
                    '
                Case "lnd"
                    Me.lst_CoverPage = objGlobals.glb_getDimensions_CoverPage_Lnd()
                    sectNew = MyBase.sct_insert_Section(False, sect, 6, "newPage", False, "lnd", Me.lst_CoverPage)
                    Me.sect = sectNew
                    '
            End Select
            '
            sectNew.PageSetup.DifferentFirstPageHeaderFooter = True
            objHFMgr.hf_hfs_linkUnlinkAll(sectNew, False)
            objHFMgr.hf_hfs_deleteAll(sectNew)
            '
            'Now set the tag style in the header
            hf = sectNew.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
            rng = hf.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Style = myDoc.Styles.Item(Me.strTagStyleName)
            '
            'rng = sect.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'rng.Select()
            '
            'sectNew = MyBase.sct_insert_SectionInFront(rng)
            '
            'Just to make sure we'll reset the page
            'If strReportMode = objRptMgr.modeLong Or strReportMode = objRptMgr.modeShort Then
            'sectNew = MyBase.sct_insert_Section(False, sect, 6, "newPage", False, "prt", Me.lst_CoverPage)
            'sectNew = MyBase.sct_insert_SectionInFront(rng,,,, "prt", Me.lst_CoverPage)
            'Else
            'sectNew = MyBase.sct_insert_Section(False, sect, 6, "newPage", False, "lnd", Me.lst_CoverPage)
            'End If
            '
            'Me.sect = sectNew
            '
            '
            'The following will return the Protrait or Landscape dimensions depending
            'on the Report Mode
            'Me.cp_Set_Dimensions(sectNew)
            rng = sectNew.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            Me.cp_insert_formattedTitleText(rng)
            'tbl = Me.cp_adjust_TitleTableBlockSize(tbl, strReportMode)
            'MyBase.Base_Paragraphs_Delete(rng, 1)
            '
            'Now do the Background
            Me.cp_Build_CoverPageFromComponents(sectNew, strCoverPageType, strReportMode)
            '
        Else
            'Has cover page... Must deal with it is lnd but we have selected from prt and vice versa
            'sect is the cover page section
            '
            If sect.PageSetup.Orientation = sectSelected.PageSetup.Orientation Then
                'Then the cover page section sect has the same orientation as the section that has the
                'current selection. so we don't have to do anything
                Me.cp_Build_CoverPageFromComponents(sect, strCoverPageType, strReportMode)
                sectNew = sect
            Else
                sect.PageSetup.Orientation = sectSelected.PageSetup.Orientation
                '
                If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
                    objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_CoverPage_Lnd())
                    'strReportMode = objRptMgr.modeLongLandscape
                Else
                    objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_CoverPage_Prt())
                    'strReportMode = objRptMgr.modeLong
                End If

                Me.cp_Build_CoverPageFromComponents(sect, strCoverPageType, strReportMode)
                sectNew = sect
            End If
        End If
        '
        '***** ???
        objGlobals.glb_get_wrdApp.ActiveWindow.View.Type = Word.WdViewType.wdPrintView
        '
        objWorkAround.wrk_fix_forCursorRace()

        '
        Return sectNew
        '
    End Function
    '
    '
    'This method will toggle the colour type of a graphic between
    'Colour/GreyScale. It can look in any section, but
    Public Sub cp_picture_changePictColour(ByRef sect As Section, doColour As Boolean)
        Dim shp As Word.Shape
        Dim dlgResult As Integer
        '
        Globals.ThisAddin.Application.ScreenUpdating = True
        'Me.isCoverPage
        '
        On Error GoTo finis
        '
        shp = Nothing                    'This Is an illegal construct In vba
        shp = Me.cp_img_getSmallImagePicture(sect)
        If Not IsNothing(shp) Then
            If shp.Type = MsoShapeType.msoPicture Then
                If doColour Then
                    shp.PictureFormat.ColorType = MsoPictureColorType.msoPictureGrayscale
                    'shp.PictureFormat.Brightness = 0.3
                    'shp.PictureFormat.Contrast = 0.7
                Else
                    shp.PictureFormat.ColorType = MsoPictureColorType.msoPictureAutomatic
                    'shp.PictureFormat.Brightness = 0.3
                    'shp.PictureFormat.Contrast = 0.7
                End If
            End If
        Else
            dlgResult = MsgBox("The colour change function is only supported" & vbCr _
                 & "in cover pages with a picture placeholder", vbOKOnly + vbExclamation, "Template Message")
        End If
        Exit Sub

finis:
        dlgResult = MsgBox("Picture colour change has failed:" & vbCr & vbCr _
        & "The most likely cause is that a target picture placeholder" & vbCr _
        & "(i.e. an existing picture) is missing from either the template" & vbCr _
        & "or your document." & vbCr & vbCr _
        & "Please contact your local IT staff for support", vbOKOnly + vbExclamation, "Template Message")

    End Sub
    '
    '
    'This method will retrieve the picture with the name specified
    'in the class variable Me.strCoverPictureName
    Public Function cp_img_getSmallImagePicture(ByRef sect As Section) As Word.Shape
        Dim hf As HeaderFooter
        Dim shp As Word.Shape
        shp = Nothing
        '
        For Each hf In sect.Headers
            For Each shp In hf.Shapes
                If shp.Name Like Me.strCoverPictureName Then
                    GoTo endloop
                End If
            Next shp
        Next hf
        '
        Return Nothing
        '
        Exit Function
endloop:
        Return shp
    End Function
    '
    '
    'This method will retrieve the picture with the name specified
    'in the class variable Me.strCoverPictureName
    Public Function cp_img_getSmallEmptyPattern(ByRef sect As Section) As Word.Shape
        Dim hf As HeaderFooter
        Dim shp As Word.Shape
        shp = Nothing
        '
        For Each hf In sect.Headers
            For Each shp In hf.Shapes
                If shp.Name Like Me.strCoverPatternName Then
                    GoTo endloop
                End If
            Next shp
        Next hf
        '
        Return Nothing
        '
        Exit Function
endloop:
        Return shp
    End Function
    '

    '
    ''' <summary>
    ''' This method will get and relocate/resize the small cover page image for the Landscape
    ''' version
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function cp_img_setSmallImageForLandscape(ByRef sect As Word.Section) As Word.Shape
        Dim shp As Word.Shape
        '
        shp = Me.cp_img_getSmallImagePicture(sect)
        '
        If Not IsNothing(shp) Then
            'shp.Rotation = 90.0
            shp.LockAspectRatio = True
            'shp.Width = 350.0
            'shp.Height = 276.85
            shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
            shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
            shp.Left = 325.7

            shp.Top = 221.9
        End If
        '
        Return shp
    End Function
    '
    '
    ''' <summary>
    ''' This method will get and relocate/resize the small cover page image for the Landscape
    ''' version
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function cp_img_setSmallEmptyPatternForLandscape(ByRef sect As Word.Section) As Word.Shape
        Dim shp As Word.Shape
        '
        shp = Me.cp_img_getSmallEmptyPattern(sect)
        '
        If Not IsNothing(shp) Then
            'shp.Rotation = 90.0
            shp.LockAspectRatio = True
            'shp.Width = 350.0
            'shp.Height = 276.85
            shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
            shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
            shp.Left = 337.25
            shp.Top = 235.05
        End If
        '
        Return shp
    End Function
    '

    '
    Public Function cp_adjust_TitleTableBlockSize(ByRef tbl As Word.Table, strReportMode As String) As Word.Table
        Dim objRptMgr As New cReport()
        Dim tblWidth As Single
        Dim dr As Word.Row
        '
        Select Case strReportMode
            Case objRptMgr.rpt_isPrt, objRptMgr.modeShort
                tblWidth = tbl.Range.Cells.Item(1).Width
                '
                'tbl.Range.Cells.Item(2).Width = 100.0
                tbl.Range.Cells.Item(2).Width = 60.0
                tbl.Range.Cells.Item(3).Width = tblWidth - tbl.Range.Cells.Item(2).Width
                'tbl.Range.Cells.Item(4).Width = 435.0
                'tbl.Range.Cells.Item(5).Width = tblWidth - tbl.Range.Cells.Item(4).Width
                'tbl.Range.Cells.Item(6).Width = 344.0
                'tbl.Range.Cells.Item(7).Width = tblWidth - tbl.Range.Cells.Item(6).Width
                '
                dr = tbl.Rows.Item(1)
                    '
            Case objRptMgr.rpt_isLnd
                tblWidth = 373.0
                tbl.Range.Cells.Item(1).Width = tblWidth
                '
                'tbl.Range.Cells.Item(2).Width = 100.0
                tbl.Range.Cells.Item(2).Width = 60.0
                tbl.Range.Cells.Item(3).Width = tblWidth - tbl.Range.Cells.Item(2).Width
                tbl.Range.Cells.Item(4).Width = tblWidth
                tbl.Range.Cells.Item(5).Width = tblWidth

                'tbl.Range.Cells.Item(4).Width = 330.0
                'tbl.Range.Cells.Item(5).Width = tblWidth - tbl.Range.Cells.Item(4).Width
                'tbl.Range.Cells.Item(6).Width = 282.0
                'tbl.Range.Cells.Item(7).Width = tblWidth - tbl.Range.Cells.Item(6).Width
                '
                dr = tbl.Rows.Item(1)
                '
        End Select

        Return tbl
        '
    End Function
    '
    ''' <summary>
    ''' This method will insert the title block and example text at the
    ''' specified range rng
    ''' </summary>
    ''' <param name="rng"></param>
    Public Sub cp_insert_formattedTitleText(ByRef rng As Word.Range)
        Dim objRptMgr As New cReport()
        Dim sect As Word.Section
        Dim para As Word.Paragraph
        'Dim strReportMode As String
        'Dim cmBoxCtrl As ComboBoxContentControl
        'Dim datePickerCtrl As DatePickerContentControl
        'Dim rng2 As Word.Range
        Dim myDoc As Word.Document
        'Dim myStyle As Word.Style
        Dim tbl As Word.Table
        'Dim dr As Word.Row
        'Dim drCell As Word.Cell
        Dim dtTime As New DateTime()
        '
        myDoc = Globals.ThisAddin.Application.ActiveDocument
        sect = rng.Sections.Item(1)
        tbl = Nothing
        '
        Try
            '
            rng.Text = "Example Heading" + vbCrLf + "Example subtitle" + vbCrLf + "00 Month yyyy"
            'rng.Style = myDoc.Styles.Item("Cp Title")
            para = rng.Paragraphs.Item(1)
            para.Style = myDoc.Styles.Item("Cp Title")
            para = rng.Paragraphs.Item(2)
            para.Style = myDoc.Styles.Item("Cp SubTitle")
            para = rng.Paragraphs.Item(3)
            para.Style = myDoc.Styles.Item("Cp Report Date")

            GoTo finis

            rng.Style = myDoc.Styles.Item("Cp Title")
            rng.Text = "Example heading"
            '
            rng = sect.Range
            para = rng.Paragraphs.Item(2)
            rng = para.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Style = myDoc.Styles.Item("Cp SubTitle")
            rng.Text = "Example subtitle"
            '
            '
            rng = sect.Range
            para = rng.Paragraphs.Item(3)
            rng = para.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Style = myDoc.Styles.Item("Cp Report Date")
            rng.Text = "00 Month yyyy"
            '
finis:

        Catch ex As Exception
            MsgBox("Unable to create CoverPage title block")
        End Try
        '
    End Sub
    '    

    '
    ''' <summary>
    ''' This method will insert a title table block at the specified range.
    ''' </summary>
    ''' <returns></returns>
    Public Function cp_insert_TitleTableBlockx(ByRef rng As Word.Range) As Word.Table
        Dim objRptMgr As New cReport()
        Dim strReportMode As String
        Dim cmBoxCtrl As ComboBoxContentControl
        Dim datePickerCtrl As DatePickerContentControl
        Dim rng2 As Word.Range
        Dim myDoc As Word.Document
        Dim myStyle As Word.Style
        Dim tbl As Word.Table
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim dtTime As New DateTime()
        '
        myDoc = Globals.ThisAddin.Application.ActiveDocument
        tbl = Nothing
        '
        Try
            '
            tbl = myDoc.Tables.Add(rng, 4, 2)
            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
            tbl.AllowAutoFit = False
            tbl.Borders.Enable = False
            tbl.TopPadding = 0.0
            tbl.BottomPadding = 0.0
            tbl.LeftPadding = 0.0
            tbl.RightPadding = 0.0
            '
            tbl.Rows.Item(1).HeightRule = WdRowHeightRule.wdRowHeightExactly
            tbl.Rows.Item(1).Height = 16.6
            tbl.Rows.Item(2).HeightRule = WdRowHeightRule.wdRowHeightAuto
            tbl.Rows.Item(2).Cells.Item(1).BottomPadding = 21.6
            tbl.Rows.Item(3).HeightRule = WdRowHeightRule.wdRowHeightAuto
            tbl.Rows.Item(4).HeightRule = WdRowHeightRule.wdRowHeightAuto
            '
            tbl.Rows.Item(1).Range.Style = myDoc.Styles.Item("Cp Report Date")
            tbl.Rows.Item(2).Cells.Item(1).Range.Style = myDoc.Styles.Item("Cp Report To")
            tbl.Rows.Item(2).Cells.Item(2).Range.Style = myDoc.Styles.Item("Cp Client Name")
            tbl.Rows.Item(3).Cells.Item(1).Range.Style = myDoc.Styles.Item("Cp Title")
            tbl.Rows.Item(3).Cells.Item(2).Range.Style = myDoc.Styles.Item("Cp Title")
            tbl.Rows.Item(4).Cells.Item(1).Range.Style = myDoc.Styles.Item("Cp SubTitle")
            tbl.Rows.Item(4).Cells.Item(2).Range.Style = myDoc.Styles.Item("Cp SubTitle")
            '
            dr = tbl.Rows.Item(1)
            dr.Cells.Merge()
            dr = tbl.Rows.Item(3)
            dr.Cells.Merge()
            dr = tbl.Rows.Item(4)
            dr.Cells.Merge()

            '
            strReportMode = objRptMgr.Rpt_Mode_Get()
            '
            tbl = Me.cp_adjust_TitleTableBlockSize(tbl, strReportMode)

#Region "Title Block Controls"

            rng2 = tbl.Range.Cells.Item(1).Range
            rng2.Collapse(WdCollapseDirection.wdCollapseStart)
            rng2.Select()
            rng2.Text = "XX "
            rng2.Collapse(WdCollapseDirection.wdCollapseEnd)
            '
            'datePickerCtrl = myDoc.ContentControls.Add(WdContentControlType.wdContentControlDate, rng2)
            Dim lst As ControlCollection
            Dim i As Integer
            '
            Try

                'We ned to get rid of remnant controls that may have been left over from a prior
                'cover page create/delete cycle
                'lst = Globals.ThisAddin.Controls
                'datePickerCtrl = Globals.ThisAddin.Controls.Item("Report Date")
                'datePickerCtrl = Globals.ThisAddin.Controls.AddContentControl(datePickerCtrl, "Report Date")
                'datePickerCtrl.Delete(True)
                'i = 1
            Catch ex1 As Exception

            End Try
            '
            datePickerCtrl = Nothing
            Try
                'datePickerCtrl = Globals.ThisAddIn.Controls.AddDatePickerContentControl(rng2, "Report Date")
                datePickerCtrl = myDoc.ContentControls.Add(WdContentControlType.wdContentControlDate, rng2)
            Catch ex2 As Exception
                'Dim obj As Control
                lst = myDoc.ContentControls
                'Globals.ThisAddIn.Controls.Remove("Report Date")
                'myDoc.ContentControls
                'i = 1
                'datePickerCtrl = Globals.ThisAddin.Controls.AddDatePickerContentControl(rng2, "Report Date")

                'For i = 0 To lst.Count - 1
                'obj = lst.Item(i)
                'If obj.Name = "Report Date" Then obj.Dispose()
                'Next
                'Globals.ThisAddin.Controls.Item("Report Date")
                'datePickerCtrl = Globals.ThisAddin.Controls.Item("Report Date")
                'datePickerCtrl.Cut()
                'rng2.Paste()
                'datePickerCtrl.Delete(True)
                i = 1
            End Try
            '
            datePickerCtrl.DateCalendarType = WdCalendarType.wdCalendarWestern
            datePickerCtrl.DateDisplayFormat() = "MMMM yyyy"
            '
            '
            datePickerCtrl.Range.Text = dtTime.Date()
            '
            myStyle = myDoc.Styles("Cp Report Date")
            datePickerCtrl.Range.Font.Color = myStyle.Font.Color

            '
            'tbl = rng.Tables.Item(1)
            drCell = tbl.Range.Cells.Item(2)
            rng2 = drCell.Range
            rng2.Text = ""
            rng2.Collapse(WdCollapseDirection.wdCollapseStart)
            rng2.Select()
            'GoTo finis
            '
            Try
                'cmBoxCtrl = Globals.ThisAddIn.Controls.Item("Report")
                cmBoxCtrl = objGlobals.glb_get_wrdActiveDoc.ContentControls.Item("Report")
                cmBoxCtrl.Delete(True)
            Catch ex As Exception

            End Try
            '
            'cmBoxCtrl = Globals.ThisAddIn.Controls.AddComboBoxContentControl("Report")
            'cmBoxCtrl = objGlobals.glb_get_wrdActiveDoc.ContentControls.Add(WdContentControlType.wdContentControlComboBox)

            'With cmBoxCtrl
            '.DropDownListEntries.Add("Report to", "Report to", 1)
            '.DropDownListEntries.Add("Proposal to", "Proposal to", 2)
            '.PlaceholderText = "Report to"
            'End With
            '
            myStyle = Globals.ThisAddin.Application.ActiveDocument.Styles("Cp Report To")
            'cmBoxCtrl.Range.Font.Color = myStyle.Font.Color
            '
#End Region

            '
            tbl.Range.Cells.Item(3).Range.Text = "[Name of the organisation]"
            tbl.Range.Cells.Item(4).Range.Text = "Example heading"
            tbl.Range.Cells.Item(5).Range.Text = "Example sub heading"

        Catch ex As Exception
            MsgBox("Unable to create CoverPage title block")
        End Try
finis:
        Return tbl
    End Function
    '    
    '
    '
    '
#Region "00 isCoverPage"
    '
    ''' <summary>
    ''' This function will return truw if the current section is a CoverPage. The
    ''' library methods 'isCoverPage1' and 'isCoverPage2' deal with the Envelope
    ''' and TandG versions directly (respectively). This routine will handle both approaches.. 
    ''' If the tag is in the first table then it will return true... If not it will test for the header style
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function cp_Bool_IsCoverPage(ByRef sect As Section) As Boolean
        '
        cp_Bool_IsCoverPage = Me.isCoverPage1(sect)
        If Not cp_Bool_IsCoverPage Then
            If sect.PageSetup.DifferentFirstPageHeaderFooter Then
                cp_Bool_IsCoverPage = Me.isCoverPage2(sect)
            End If
        End If
        'isCoverPage = Me.isCoverPage2(sect)
    End Function
    '
    'This method will determine if the current section
    'contains a cover page. It does so by looking for the
    'style tag_coverPage in the first cell of the first Table
    Private Function isCoverPage1(ByRef sect As Section) As Boolean
        Dim rng As Range
        Dim tbl As Table
        Dim drCell As Cell
        Dim drCellStyle, cpTagStyle As Word.Style
        Dim hf As Word.HeaderFooter
        Dim myDoc As Word.Document
        '
        isCoverPage1 = False
        rng = sect.Range
        myDoc = rng.Document
        '
        Try
            If rng.Tables.Count = 0 Then Exit Function
            tbl = rng.Tables(1)
            drCell = tbl.Range.Cells(1)
            drCellStyle = drCell.Range.Style
            cpTagStyle = myDoc.Styles(Me.strTagStyleName)
            '
            'If drCell.Range.Style Is Globals.ThisAddin.Application.ActiveDocument.Styles("tag_coverPage") Then
            If drCellStyle.NameLocal = Me.strTagStyleName Then
                isCoverPage1 = True
            End If
            '
        Catch ex As Exception
            'Will get here if we have floating tables in the body.. For some reason they don't
            'react kindly to looking for the styles in their first cell... So we'll check the 
            'Header style
            '
            If sect.PageSetup.DifferentFirstPageHeaderFooter Then
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
            Else
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
            End If
            rng = hf.Range
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            If rng.Style.NameLocal = Me.strTagStyleName Then
                isCoverPage1 = True
            End If
        End Try

    End Function
    '
    '
    ''' <summary>
    ''' This method will determine if the current section contains a cover page. It does so by looking for the
    ''' style tag_coverPage in the Header of the first page
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Private Function isCoverPage2(ByRef sect As Section) As Boolean
        Dim rng As Range
        Dim rngStyle As Word.Style
        '
        isCoverPage2 = False
        rng = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rngStyle = rng.Style
        '
        'MsgBox("Style is = " + rngStyle.NameLocal)
        '
        'If drCell.Range.Style Is Globals.ThisAddin.Application.ActiveDocument.Styles("tag_coverPage") Then
        If rngStyle.NameLocal = "tag_coverPage" Then
            isCoverPage2 = True
        End If
        '
    End Function
    '
#End Region
    '
    Public Sub cp_Build_Background(ByRef sect As Word.Section)
        Dim objBrndMgr As New cBrandMgr()
        Dim hf As Word.HeaderFooter
        Dim rng As Word.Range
        '
        'Empty out the background and reset the header style
        hf = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        rng = hf.Range
        rng.Delete()
        '
        rng.Style = sect.Range.Document.Styles.Item(Me.strTagStyleName)
        '
        'Now rebuild the background
        objBrndMgr.brnd_Rebuild_Background(sect, False, True)
        '
    End Sub
    '
    Public Sub cp_rename_Shape(ByRef rng As Word.Range, strNewName As String)
        Dim shp As Word.Shape
        '
        Try
            If rng.ShapeRange.Count <> 0 Then
                shp = rng.ShapeRange.Item(1)
                shp.Name = strNewName
            End If
        Catch ex As Exception

        End Try

    End Sub
    '
    Public Sub cp_Build_CoverPageFromComponents(ByRef sect As Word.Section, strCoverPageID As String, strReportMode As String)
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objWCAGMgr As New cWCAGMgr()
        Dim hf As Word.HeaderFooter
        Dim objBBMgr As New cBBlocksHandler()
        Dim objRptMgr As New cReport()
        Dim shp_Pattern, shp_Pict As Word.Shape
        '
        Dim rng As Range
        Dim delta As Single
        Dim objSectMgr As New cSectionMgr()
        Dim objGrfxMgr As New cGraphicsMgr()
        '
        'Globals.ThisAddin.Application.ScreenUpdating = False
        delta = 10.5
        'srcTemplate = Me.myDoc.AttachedTemplate
        'HFRange = Me.sect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Range
        '
        Try
            'Empty out and rebuild the background
            '
            Me.cp_Build_Background(sect)
            '
            Select Case strReportMode
                Case objRptMgr.rpt_isPrt, objRptMgr.modeShort
                    Select Case strCoverPageID
                        Case Me.objGlobals._glb_cpType_TGFilledPattern
                            hf = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                            rng = hf.Range
                            rng.Collapse(WdCollapseDirection.wdCollapseStart)
                            rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_Cpg_PictFilled", "CoverPage", rng)                 'prt purple triangles picture, shpName=cp_pct_large
                            '
                            'Me.cp_rename_Shape(rng, "cp_pict_purplePattern_prt")
                            Me.cp_rename_Shape(rng, "cp_pict_large")
                            objWCAGMgr.wcag_set_decorative(rng, True)
                            '
                        Case Me.objGlobals._glb_cpType_TGEmptyPattern
                            hf = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                            rng = hf.Range
                            rng.Collapse(WdCollapseDirection.wdCollapseStart)
                            'rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_Cpg_PictEmptyPattern", "CoverPage", rng)
                            rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_Cpg_PictEmptyPatternSmall", "CoverPage", rng)      'prt lattice pattern to sit inside pictures
                            '
                            objWCAGMgr.wcag_set_decorative(rng, True)
                            '
                        Case Me.objGlobals._glb_cpType_TGPicturePattern
                            hf = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                            rng = hf.Range
                            rng.Collapse(WdCollapseDirection.wdCollapseStart)
                            rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_Cpg_PictPicture", "CoverPage", rng)                'prt seaside picture
                            '
                            'Me.cp_rename_Shape(rng, "cp_pict_seaSide_prt")
                            Me.cp_rename_Shape(rng, "cp_pict_large")
                            objWCAGMgr.wcag_set_decorative(rng, True)
                            '
                            rng.Collapse(WdCollapseDirection.wdCollapseStart)
                            rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_Cpg_PictEmptyPatternSmall", "CoverPage", rng)      'prt lattice pattern to sit inside pictures
                            '
                            objWCAGMgr.wcag_set_decorative(rng, True)
                            '
                    End Select
                    '
                Case objRptMgr.rpt_isLnd
                    Select Case strCoverPageID
                        Case Me.objGlobals._glb_cpType_TGFilledPattern
                            hf = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                            rng = hf.Range
                            rng.Collapse(WdCollapseDirection.wdCollapseStart)
                            rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_Cpg_PictFilled_Lnd", "CoverPage", rng)             'lnd purple triangles picture
                            '
                            'Me.cp_rename_Shape(rng, "cp_pict_purplePattern_lnd")
                            Me.cp_rename_Shape(rng, "cp_pict_large")
                            objWCAGMgr.wcag_set_decorative(rng, True)
                            '
                        Case Me.objGlobals._glb_cpType_TGEmptyPattern
                            hf = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                            rng = hf.Range
                            rng.Collapse(WdCollapseDirection.wdCollapseStart)
                            rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_Cpg_PictEmptyPattern_Lnd", "CoverPage", rng)       'lnd lattice pattern to sit inside pictures
                            '
                            objWCAGMgr.wcag_set_decorative(rng, True)
                            '
                        Case Me.objGlobals._glb_cpType_TGPicturePattern
                            hf = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                            rng = hf.Range
                            rng.Collapse(WdCollapseDirection.wdCollapseStart)
                            rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_Cpg_PictPicture_Lnd", "CoverPage", rng)
                            '
                            shp_Pict = rng.ShapeRange.Item(1)
                            '
                            'Me.cp_rename_Shape(rng, "cp_pict_seaSide_lnd")
                            Me.cp_rename_Shape(rng, "cp_pict_large")
                            objWCAGMgr.wcag_set_decorative(rng, True)
                            '
                            rng.Collapse(WdCollapseDirection.wdCollapseStart)
                            rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_Cpg_PictEmptyPatternSmll_Lnd", "CoverPage", rng)
                            '
                            'Now adjust the relative positions of the open patter to the picture
                            '
                            shp_Pattern = rng.ShapeRange.Item(1)
                            '
                            Me.cp_align_patternToPict(shp_Pattern, shp_Pict)
                            '
                            objWCAGMgr.wcag_set_decorative(rng, True)
                            '
                    End Select
            End Select
            '



        Catch ex As Exception
            MsgBox("The new Cover Page has failed to Build.. The most likely cause is Template Related." & vbCr & vbCr &
            "You might not be using the right template.. or the template AutoText Entries" & vbCr & vbCr &
            "have been corrupted.. In either case you will need to refer to your IT support staff")

        End Try
        '
        objGlobals.glb_screen_update(True)
        '
    End Sub
    '
    ''' <summary>
    ''' This method aligns the open pattern above the picture shape (shpPict) with a boundary
    ''' of delta points. The default for delta is 10.0 pts
    ''' </summary>
    ''' <param name="shpPattern"></param>
    ''' <param name="shpPict"></param>
    ''' <param name="delta"></param>
    Public Sub cp_align_patternToPict(ByRef shpPattern As Word.Shape, ByRef shpPict As Word.Shape, Optional delta As Single = 10.0)
        Try
            If Not IsNothing(shpPict) And Not IsNothing(shpPattern) Then
                shpPattern.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
                shpPattern.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
                '
                shpPattern.Width = shpPict.Width - 2 * delta
                shpPattern.Height = shpPict.Height - 2 * delta
                '
                shpPattern.Left = shpPict.Left + delta
                shpPattern.Top = shpPict.Top + delta
            End If

        Catch ex As Exception

        End Try
    End Sub

    '
    '
    ''' <summary>
    ''' This method adjusts the various Cover Page Styles to meet the requiremenst of 
    ''' each of the CoverPage types
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="strCoverPageType"></param>
    Public Sub cp_changeCoverPageStyles(ByRef myDoc As Word.Document, strCoverPageType As String)

        Select Case strCoverPageType
            Case 1 'Me.objGlobals.cpType_HalfPageImage
                '
                'myDoc.Styles("Cp Report To").Font.TextColor.RGB = RGB(255, 254, 255)
                'myDoc.Styles("Cp Client Name").Font.TextColor.RGB = RGB(255, 254, 255)
                'myDoc.Styles("Cp Report Date").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp Title").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp SubTitle").Font.TextColor.RGB = RGB(255, 254, 255)
                'myDoc.Styles("Cp Bold SubTitle").Font.TextColor.RGB = RGB(255, 254, 255)
                '
            '
            Case 2 'Me.objGlobals.cpType_SmallImage_Coloured
                '
                'myDoc.Styles("Cp Report To").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp Client Name").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp Report Date").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp Title").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp SubTitle").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp Bold SubTitle").Font.TextColor.RGB = RGB(0, 1, 0)
                '
            Case 3 'Me.objGlobals.cpType_SmallImage_BW         'Since we are toggling the colour separately the bw is the same as the Colour
                '
                'myDoc.Styles("Cp Report To").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp Client Name").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp Report Date").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp Title").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp SubTitle").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp Bold SubTitle").Font.TextColor.RGB = RGB(0, 1, 0)
                '
            Case 4 'Me.objGlobals.cpType_ColouredDeltas
                '
                'myDoc.Styles("Cp Report To").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp Client Name").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp Report Date").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp Title").Font.TextColor.RGB = Me.objGlobals.colour_Purple
                'myDoc.Styles("Cp SubTitle").Font.TextColor.RGB = RGB(151, 143, 139)
                'myDoc.Styles("Cp Bold SubTitle").Font.TextColor.RGB = RGB(0, 1, 0)
                '
            Case 5 'Me.objGlobals.cpType_GreyScaleDeltas
                '
                'myDoc.Styles("Cp Report To").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp Client Name").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp Report Date").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp Title").Font.TextColor.RGB = myDoc.Styles("rgb-Purple").Font.TextColor.RGB
                'myDoc.Styles("Cp Title").Font.TextColor.RGB = RGB(0, 1, 0)
                'myDoc.Styles("Cp SubTitle").Font.TextColor.RGB = RGB(151, 143, 139)
                'myDoc.Styles("Cp Bold SubTitle").Font.TextColor.RGB = RGB(0, 1, 0)
                '
            Case Else
        End Select


    End Sub
    '


#Region "Title Cells"
    '
    '*****
    ''' <summary>
    ''' This method will move the selection in the CoverPage (as specified in sect) to the Title cell.
    ''' If selectAll is true, then the contents of the cell is selected
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="selectAll"></param>
    ''' <returns></returns>
    Public Function cp_sel_MoveToTitleCell(ByRef sect As Word.Section, Optional selectAll As Boolean = False) As Word.Range
        Dim rng As Word.Range
        '
        rng = Nothing
        '
        Try
            rng = getCoverPageTitleRange(sect)
            If IsNothing(rng) Then
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
            End If
            '
            If Not selectAll Then rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Select()
            '
            rng = objGlobals.glb_get_wrdSel.Range()

        Catch ex As Exception

        End Try
        '
        Return rng
    End Function
    '
    Public Function cp_sel_MoveToTitle(ByRef sect As Word.Section, Optional selectAll As Boolean = False) As Word.Range
        Dim rng As Word.Range
        Dim titleStyle As Word.Style
        '
        rng = Nothing
        '
        Try
            rng = sect.Range.Paragraphs.First.Range
            titleStyle = rng.Style
            '
            If titleStyle.NameLocal = "Cp Title" Then
                If selectAll Then
                    rng.MoveEnd(WdUnits.wdCharacter, -1)
                    rng.Select()
                Else
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    rng.Select()
                End If

            End If
            '
        Catch ex As Exception

        End Try
        '
        Return rng

    End Function

    '
    ''' <summary>
    ''' This method makes certian that the Empty Pattern is not hidden by the bottom
    ''' picture in the Cover Page
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub cp_EmptyPattern_ToFront(ByRef sect As Word.Section)
        Dim hf As Word.HeaderFooter
        '
        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        For Each shp In hf.Shapes
            If shp.Name Like "cp_Empty_*" Then
                shp.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringInFrontOfText)
                shp.ZOrder(Microsoft.Office.Core.MsoZOrderCmd.msoBringToFront)
            End If
        Next
        '
    End Sub

    Public Function getCoverPageTitleRange(ByRef sect As Word.Section) As Range
        'This method will return the range of the cell that contains the main
        'title on the Cover Page
        Dim drCell As Word.Cell
        Dim tbl As Word.Table
        Dim rng As Word.Range
        '
        On Error GoTo finis
        '
        drCell = sect.Range.Tables(1).Range.Cells(4)             'This cell holds the nested table that contains the main headings
        drCell.Range.Select()
        rng = drCell.Range
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'rng.Select()
        '
        'drCell = sect.Range.Tables(1).Range.Cells(9)             'This cell holds the nested table that contains the main headings
        'drCell.Range.Select()
        'Now get the nested Table and find the title cell
        'tbl = drCell.Tables.Item(1)
        'drCell = tbl.Range.Cells(2)
        'rng = drCell.Range
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'rng.Select()
        '
        getCoverPageTitleRange = rng
        Exit Function
        '
finis:
        getCoverPageTitleRange = Nothing
        '
    End Function
    '
    ''' <summary>
    ''' Returns the Main title text on the cover page, given the main
    ''' table that contains the elemenst of the page
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function getCoverPageTitleText(ByRef tbl As Word.Table) As String
        Dim drCell As Word.Cell
        Dim tblnest As Word.Table
        Dim rng As Word.Range
        Dim strText As String
        '
        strText = ""
        Try
            drCell = tbl.Range.Cells(9)             'This cell holds the nested table that contains the main headings
            'drCell.Range.Select()
            'Now get the nested Table and find the title cell
            tblnest = drCell.Tables.Item(1)
            drCell = tblnest.Range.Cells(2)
            rng = drCell.Range
            strText = rng.Text.Trim()
        Catch ex As Exception
            strText = ""
        End Try
        Return strText
    End Function
#End Region

End Class

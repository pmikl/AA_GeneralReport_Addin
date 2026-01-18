Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cSectionMgr
    Public objGlobals As cGlobals
    Public objWrkAround As cWorkArounds
    Public Sub New()
        Me.objGlobals = New cGlobals()
        Me.objWrkAround = New cWorkArounds()
    End Sub
    '
    '
    ''' <summary>
    ''' ************  NEEDS WORK  ***********************************************************************
    ''' This function will resize a shape (shp) while keeping its aspect ratio and sit it centrally
    ''' in the page
    ''' </summary>
    ''' <param name="objShpMgr"></param>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function sct_fit_ShapeToPage(ByRef objShpMgr As cShapeMgr, ByRef sect As Word.Section) As Boolean
        Dim objWCAGMgr As New cWCAGMgr()
        Dim shp As Word.Shape
        Dim w, h, width_old, height_old As Single
        Dim lstOfScaleFactors As New Collection()
        Dim scaleFactor_w, scaleFactor_h, shpAspectRatio As Single
        Dim dummy_h, dummy_w As Single
        Dim strShpAspectRatio As String
        Dim i As Integer
        Dim irror As Boolean
        '
        shp = objShpMgr.shp
        shp.LockAspectRatio = True
        shpAspectRatio = shp.Height / shp.Width
        dummy_h = shp.Height
        dummy_w = shp.Width
        w = shp.Width
        h = shp.Height
        '
        width_old = shp.Width
        height_old = shp.Height
        '
        strShpAspectRatio = ""
        irror = False

        If shpAspectRatio > 1 Then strShpAspectRatio = "shp_is_portrait"
        If shpAspectRatio = 1 Then strShpAspectRatio = "shp_is square"
        If shpAspectRatio < 1 Then strShpAspectRatio = "shp_is_landscape"
        '
        If sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait Then
            Select Case strShpAspectRatio
                Case "shp_is_portrait"
                    dummy_w = dummy_w * 0.8
                    dummy_h = dummy_w * shpAspectRatio

                    If dummy_h < 0.7 * sect.PageSetup.PageHeight Then
                        'all is OK.. SO we can operate on the real shape
                        shp.Width = sect.PageSetup.PageWidth * 0.8
                        objShpMgr.scaleFactor_W = height_old / shp.Height
                        objShpMgr.scaleFactor_H = width_old / shp.Width
                    Else
                        'Scale it until the height fits
                        scaleFactor_h = 0.8
                        '
                        For i = 1 To 16
                            dummy_h = sect.PageSetup.PageHeight * scaleFactor_h
                            dummy_w = dummy_h / shpAspectRatio
                            '
                            If dummy_h < 0.7 * sect.PageSetup.PageHeight Then
                                shp.Height = dummy_h
                                objShpMgr.scaleFactor_W = height_old / shp.Height
                                objShpMgr.scaleFactor_H = width_old / shp.Width
                                Exit For
                            End If
                            '
                            scaleFactor_h = scaleFactor_h - 0.05
                            If scaleFactor_h < 0.1 Then
                                objShpMgr.scaleFactor_W = 1
                                objShpMgr.scaleFactor_H = 1

                                irror = True
                                Exit For
                            End If
                        Next

                    End If

                Case "shp_is_square"
                    shp.Width = sect.PageSetup.PageWidth * 0.8
                    scaleFactor_h = 0.8
                    scaleFactor_w = 0.8

                Case "shp_is_landscape"
                    shp.Width = sect.PageSetup.PageWidth * 0.8
                    objShpMgr.scaleFactor_W = height_old / shp.Height
                    objShpMgr.scaleFactor_H = width_old / shp.Width

                    scaleFactor_h = 0.8
                    scaleFactor_w = 0.8

            End Select

        End If

        If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
            Select Case strShpAspectRatio
                Case "shp_is_portrait"
                    shp.Height = shp.Height * 0.8
                    scaleFactor_h = 0.8
                    scaleFactor_w = 0.8

                Case "shp_is_square"
                    shp.Height = shp.Height * 0.8
                    scaleFactor_h = 0.8
                    scaleFactor_w = 0.8

                Case "shp_is_landscape"
                    dummy_h = dummy_h * 0.8
                    dummy_w = dummy_h / shpAspectRatio

                    If dummy_w < 0.7 * sect.PageSetup.PageWidth Then
                        'all is OK.. SO we can operate on the real shape
                        shp.Height = shp.Height * 0.8
                        scaleFactor_h = 0.8
                        scaleFactor_w = 0.8
                    Else
                        'Scale it until the height fits
                        scaleFactor_w = 0.8
                        '
                        For i = 1 To 16
                            dummy_w = w * scaleFactor_w
                            dummy_h = dummy_w / shpAspectRatio
                            '
                            If dummy_w < 0.7 * sect.PageSetup.PageWidth Then
                                shp.Height = dummy_h
                                scaleFactor_w = scaleFactor_h
                                Exit For
                            End If
                            '
                            scaleFactor_h = scaleFactor_h - 0.05
                            If scaleFactor_h < 0.1 Then
                                irror = True
                                Exit For
                            End If
                        Next

                    End If

            End Select

        End If
        '
        If irror = True Then
            shp.Top = (sect.PageSetup.PageHeight - shp.Height) / 2
            shp.Left = (sect.PageSetup.PageWidth - shp.Width) / 2
            '
            objShpMgr.scaleFactor_W = height_old / shp.Height
            objShpMgr.scaleFactor_H = width_old / shp.Height
            '
            shp.ZOrder(MsoZOrderCmd.msoSendBehindText)
            shp.ZOrder(MsoZOrderCmd.msoSendToBack)
            '
            'For WCAG purposes
            'Set Decorative property
            objWCAGMgr.wcag_set_decorative(shp, True)
            '
        End If
        '
        Return irror
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will return a range that is set to the beginning of the last paragraph
    ''' in the current section (as defined by the beginning of the current selection.
    ''' The range is collapsed to a single point... Verified 20231025
    ''' </summary>
    ''' <returns></returns>
    Public Function sct_Set_RngTo_SectionEndParagraph_Beginning() As Word.Range
        Dim sect As Word.Section
        Dim rng As Word.Range

        sect = Globals.ThisAddIn.Application.Selection.Sections.Item(1)
        rng = sect.Range
        'rng.Information()
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
        If Me.sct_Is_LastSection() Then
            'End of document, so just move past the end of the last paragraph
            'rng.End = rng.End - 1
        Else
            'Somewhere in the document, so we need to move past the section break and
            'the last paragraph
            rng.End = rng.End - 2
        End If
        '
        Return rng
        '
    End Function
    '
    ''' <summary>
    ''' This method will reurn true if the current selection in the ActiveDocument is
    ''' in the last section, or false if it is not
    ''' </summary>
    ''' <returns></returns>
    Public Function sct_Is_LastSection() As Boolean
        Dim rslt As Boolean
        Dim myDoc As Word.Document
        Dim sel As Word.Selection
        Dim sect, sectLast As Word.Section
        '
        myDoc = Globals.ThisAddIn.Application.ActiveDocument
        sel = Globals.ThisAddIn.Application.Selection
        '
        sect = sel.Sections.Item(1)
        sectLast = myDoc.Sections.Last
        '
        rslt = False
        If sect.Index = sectLast.Index Then rslt = True
        '
        Return rslt
    End Function
    '
    ''' <summary>
    ''' This method will delete the document contents (leaving header/footers
    ''' alone). It will then assign the Body text style to the remaining
    ''' paragraph
    ''' </summary>
    Public Sub sct_delete_allSections()
        '
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim rng As Range
        Dim sect As Word.Section
        Dim myDoc As Word.Document
        '
        myDoc = Me.objGlobals.glb_get_wrdActiveDoc()
        '
        Try
            'Try to remove the controls in the cover page
            'See cChapterReport.Rpt_Section_Delete
            '
            'Globals.ThisAddIn.Controls.Remove("Report Date")
            'Globals.ThisAddIn.Controls.Remove("Report")
        Catch ex As Exception

        End Try
        '
        Try
            Try
                For Each sect In myDoc.Sections
                    'If Not (sect.Index = myDoc.Sections.Last.Index) Then
                    sect.Range.Delete()

                    'End If
                Next
            Catch ex As Exception

            End Try
            '
            sect = myDoc.Sections.Last
            Me.sct_delete_allSectionContents(sect, 1)
            'sect = Me.objGlobals.glb_get_wrdSect()
            sect.PageSetup.DifferentFirstPageHeaderFooter = False
            sect.PageSetup.OddAndEvenPagesHeaderFooter = False
            objHfMgr.hf_hfs_deleteAll(sect)
            rng = sect.Range
            rng.Style = myDoc.Styles("Body text")
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Select()
            '
            'objGlobals.glb_screen_updateLeaveAsItWas()

        Catch ex As Exception
            'MsgBox("Error in cSectionMgr.deleteAll")
        End Try
        '
    End Sub


    '
    '
    ''' <summary>
    ''' This function will shift the left margin to a position measured from the left hand edge of the
    ''' Header Table. The Header and Footer tables are re-adjusted so that their positions dont change.
    ''' -   marginOffSet is  0 pts, then the left margin will align with the left edge of the Header Table.
    ''' -   marginOffSet is  x pts, then the left margin will be x pts to the right of the left edge of the Header Table
    ''' -   marginOffSet is -x pts, then the left margin will be to the left of the left edge of the Header Tablle
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="marginOffSet"></param>
    ''' <returns></returns>
    Public Function sect_Toggle_Width(ByRef sect As Section, marginOffSet As Single) As String
        Dim objChptBase As New cChptBase()
        Dim objChptBanner As New cChptBanner()
        Dim leftMarginCurrent, leftMarginDefault As Single
        Dim lst As Collection
        Dim lstOfHeaderEdges, lstOfFooterEdges As New Collection()
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim lstOfOldSettings As New Collection()
        Dim tbl As Word.Table
        Dim rng As Word.Range
        '
        'First lets get the standard table offset so that we can allow for this when we toggle page widths
        'marginOffSet = objGlobals.glb_Table_Outdent()
        '
        lstOfHeaderEdges = objHfMgr.hf_hfs_getHfTableEdges(sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary))
        lstOfFooterEdges = objHfMgr.hf_hfs_getHfTableEdges(sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary))
        '
        'Let's shift the left margin
        '
        Select Case sect.PageSetup.Orientation
            Case WdOrientation.wdOrientLandscape
                leftMarginCurrent = sect.PageSetup.LeftMargin
                lst = objGlobals.glb_getDimensions_Std_Lnd()
                leftMarginDefault = CSng(lst("leftMargin"))
                If leftMarginCurrent = leftMarginDefault Then
                    'sect.PageSetup.LeftMargin = objGlobals.glb_hfs_getHFTableEdge(sect, "header_leftEdge") + delta
                    sect.PageSetup.LeftMargin = CSng(lstOfHeaderEdges("leftEdge")) + marginOffSet
                Else
                    sect.PageSetup.LeftMargin = leftMarginDefault
                End If

            Case WdOrientation.wdOrientPortrait
                leftMarginCurrent = sect.PageSetup.LeftMargin
                lst = objGlobals.glb_getDimensions_Std_Prt()
                leftMarginDefault = CSng(lst("leftMargin"))
                If leftMarginCurrent = leftMarginDefault Then
                    'sect.PageSetup.LeftMargin = objGlobals.glb_hfs_getHFTableEdge(sect, "header_leftEdge") + delta
                    sect.PageSetup.LeftMargin = objGlobals.glb_hfs_getHFTableEdge(sect, "header_leftEdge") + marginOffSet
                Else
                    sect.PageSetup.LeftMargin = leftMarginDefault
                End If
        End Select
        '
        'Now re-adjust the Headers and Footers so they appear not to be affecte by the left margin shift
        '
        objHfMgr.hf_headers_resize_all(sect)
        objHfMgr.hf_footers_resize_all(sect)
        '
        'Now toggle the width of the chapter banner if there is one
        rng = sect.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        If rng.Tables.Count <> 0 Then
            tbl = rng.Tables.Item(1)
            If objChptBanner.bnr_is_Chapter_Bdy_or_ES_or_AP(tbl) Then
                Me.objGlobals.glb_tbls_AutoFitBanner(tbl, False)
                'objChptBase.chptBase_Banner_Autofit(tbl, sect, False)
            End If
        End If
        '
        Return "hello"
        '
    End Function
    '
    ''' <summary>
    ''' This method will copy the specified section (sect) and paste it either behind (pasteBehind = True)
    ''' or it will paste it in front (pasteBehind = false). It will return (in sect, as well as the return variable)
    ''' the new section. Typically used to duplicate empty "prt" chapters
    ''' </summary>
    ''' <param name="pasteBehind"></param>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function sct_copyAndPaste_Section(pasteBehind As Boolean, ByRef sect As Word.Section, Optional numSections As Integer = 1) As Word.Section
        Dim rng, rng2 As Word.Range
        Dim tbl As Word.Table
        Dim para As Word.Paragraph
        '
        Select Case numSections
            Case 1
                rng = sect.Range
                rng.Copy()
            Case 2
                rng = sect.Range
                rng.MoveStart(WdUnits.wdSection, -1)
                rng.Copy()
            Case Else
                rng = sect.Range
                rng.Copy()
        End Select
        '
        If pasteBehind Then
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            rng.Paste()
        Else
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            If rng.Tables.Count <> 0 Then
                'We have a section with a table at the front.. We need to make room
                tbl = rng.Tables.Item(1)
                tbl.Rows.Add(tbl.Rows.Item(1))
                tbl.Split(tbl.Rows.Item(2))
                tbl = objGlobals.glb_get_wrdSelRng().Tables.Item(1)
                tbl.Delete()
                rng = objGlobals.glb_get_wrdSelRng()
                rng.Style = rng.Document.Styles.Item("Body Text")
                rng.Paste()
                'Now get rid of the extraneous paragraph that was added as a result of the
                'table split
                sect = rng.Document.Sections.Item(rng.Sections.Item(1).Index + 1)
                rng2 = sect.Range
                rng2.Collapse(WdCollapseDirection.wdCollapseStart)

                para = rng2.Paragraphs.Item(1)
                para.Range.Delete()
            End If
        End If
        '
        sect = rng.Sections.Item(1)
        '
        Return sect
    End Function

    '
    Public Function xsct_resize_ToCustom(ByRef sect As Word.Section, lstOfMargins As Collection, Optional strOrientation As String = "prt") As Boolean
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objGlobals As New cGlobals()
        Dim objChptBanner As New cChptBanner()
        Dim objChptBase As New cChptBase()
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim rslt As Boolean
        '
        rslt = False
        '
        Try
            'Call objHfMgr.hf_hfs_deleteAll(sect)
            If strOrientation = "prt" Then
                sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait
            Else
                sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape

            End If
            '
            'Do the Header first
            objGlobals.glb_setDimensions(sect, lstOfMargins)
            objHfMgr.hf_headers_resize_all(sect)
            objHfMgr.hf_footers_resize_all(sect)
            '
            'Now autofit the Banner (if it is there)
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            If rng.Tables.Count <> 0 Then
                tbl = rng.Tables.Item(1)
                If objChptBanner.bnr_is_Chapter_Bdy_or_ES_or_AP(tbl) Then
                    Me.objGlobals.glb_tbls_AutoFitBanner(tbl, False)
                    'objChptBase.chptBase_Banner_Autofit(tbl, sect, False)
                End If
                'objChptBase.chptBase_Banner_Autofit(tbl, sect, False)
            End If

        Catch ex As Exception
            rslt = False
        End Try


        Return rslt
    End Function

    Public Function sct_reset_ToPortrait(ByRef sect As Word.Section, Optional strHeaderStyleName As String = "spacer") As Boolean
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objGlobals As New cGlobals()
        Dim objChptBase As New cChptBase()
        Dim objChptBanner As New cChptBanner()
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim rslt As Boolean
        '
        rslt = False
        '
        Try
            Call objHfMgr.hf_hfs_deleteAll(sect)
            sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait
            '
            'Do the Header first
            objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Std_Prt())
            objHfMgr.hf_headers_insert(sect,,,,, strHeaderStyleName)
            objHfMgr.hf_footers_insert(sect)
            '
            'Now autofit the Banner (if it is there)
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            If rng.Tables.Count <> 0 Then
                tbl = rng.Tables.Item(1)
                If objChptBanner.bnr_is_Chapter_Bdy_or_ES_or_AP(tbl) Then
                    Me.objGlobals.glb_tbls_AutoFitBanner(tbl, False)
                    'objChptBase.chptBase_Banner_Autofit(tbl, sect, False
                End If
                'objChptBase.chptBase_Banner_Autofit(tbl, sect, False)
            End If

        Catch ex As Exception
            rslt = False
        End Try


        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method will reset the section (sect) to Landscape using either default dimensions or, the
    ''' dimensions supplied in lst. Default.. 
    ''' 1.  lst isNothing, then the dimensions are standard landscpae "glb_getDimensions_Landscape()"
    ''' 2.  lst is defined and count is 0, then the dimensions are standard Landscape Report follower page
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="lst"></param>
    ''' <returns></returns>
    Public Function sct_reset_ToLandscape(ByRef sect As Word.Section, Optional ByRef lst As Collection = Nothing, Optional strHeaderStyleName As String = "spacer") As Boolean
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objGlobals As New cGlobals()
        Dim objChptBase As New cChptBase()
        Dim objChptBanner As New cChptBanner()
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim rslt As Boolean
        '
        rslt = False
        '
        If IsNothing(lst) Then
            lst = objGlobals.glb_getDimensions_Std_Lnd()
        Else
            If lst.Count = 0 Then
                lst = objGlobals.glb_getDimensions_Landscape_LndRpt_followerPage()
            End If
        End If
        '
        Try
            Call objHfMgr.hf_hfs_deleteAll(sect)
            sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape
            '
            'Do the Header first
            objGlobals.glb_setDimensions(sect, lst)
            objHfMgr.hf_headers_insert(sect,,,,, strHeaderStyleName)
            objHfMgr.hf_footers_insert(sect)
            '
            'Now autofit the Banner (if it is there)
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            If rng.Tables.Count <> 0 Then
                tbl = rng.Tables.Item(1)
                If objChptBanner.bnr_is_Chapter_Bdy_or_ES_or_AP(tbl) Then
                    Me.objGlobals.glb_tbls_AutoFitBanner(tbl, False)
                    'objChptBase.chptBase_Banner_Autofit(tbl, sect, False)
                End If
                'objChptBase.chptBase_Banner_Autofit(tbl, sect, False)
            End If

        Catch ex As Exception
            rslt = False
        End Try


        Return rslt
        '
    End Function

    '
    Public Function sct_resize_ToLandscape(ByRef sect As Word.Section) As Boolean
        Dim objPlhBase As New cPlHBase()
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objGlobals As New cGlobals()
        Dim objChptBase As New cChptBase()
        Dim objBackPanelMgr As New cBackPanelMgr()
        Dim objChptBanner As New cChptBanner()
        Dim objDivMgr As New cChptDivider()
        Dim objContactMgr As New cContactsMgr()
        Dim objisOK As New cIsOKToDo()
        Dim objTocMgr As New cTOCMgr()
        Dim objCpMgr As New cCoverPageMgr()
        Dim objTables As New cTablesMgr()
        Dim listOfPanels As New List(Of cShapeMgr)
        Dim lstOfEdges_Header, lstOfEdges_Footer As New Collection()
        Dim objShpMgr As New cShapeMgr()
        Dim lstOfDimensions As New Collection()
        Dim dr As Word.Row
        Dim drCell, drCellNested As Word.Cell
        Dim iShp As Word.InlineShape
        Dim hf As HeaderFooter
        Dim objShp As cShapeMgr

        Dim strType As String
        Dim tbl, tblNested As Word.Table
        Dim rslt As Boolean
        '
        rslt = False
        strType = ""
        dr = Nothing
        drCell = Nothing
        tblNested = Nothing
        iShp = Nothing
        '
        Try
            If sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait Then
                '
                'Get the original table edges
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                lstOfEdges_Header = objHfMgr.hf_hfs_getHfTableEdges(hf, True)
                hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                lstOfEdges_Footer = objHfMgr.hf_hfs_getHfTableEdges(hf, True)
                '
                'Reset the orientation. Get any back panels and resize hem to fit the
                'new shape
                sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape
                '
                'We need to set measurements to a portrait type page so that the
                'WdOrientation.wdOrientPortrait doesn't throw a fault. We also adjust any columns that 
                'may have been set for a Landscape page
                'objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Std_Lnd)
                If sect.PageSetup.TextColumns.Count > 1 Then
                    sect.PageSetup.TextColumns.EvenlySpaced = True
                End If
                '
                listOfPanels = objBackPanelMgr.pnl_getBackPanel_PlaceHolders(sect)
                objBackPanelMgr.pnl_resize_PanelToFillPage(sect, listOfPanels)
                '
                'Get the type of section that the current selection is in
                strType = objisOK.isOKto_selection_isIn()
                '
                Select Case strType
                    Case "caseStudy"
                        objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_CaseStudy_Lnd())

                    Case "cp"
                        objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_CoverPage_Lnd())
                        '
                        'Just delete the small pictures the y can be reset from the main menu option
                        objCpMgr.ChptBase_delete_SmallPicturePlaceHolders(sect)

                        'objCpMgr.cp_img_setSmallImageForLandscape(sect)
                        'objCpMgr.cp_img_setSmallEmptyPatternForLandscape(sect)
                        'tbl = objGlobals.glb_tbls_AutoFitRegularTable(sect, False)
                        '
                    Case "contFront"
                        lstOfDimensions = objGlobals.glb_getDimensions_Contacts_Lnd()

                        objGlobals.glb_setDimensions(sect, lstOfDimensions)
                        tbl = objGlobals.glb_tbls_AutoFitRegularTable(sect, False)
                        '
                        lstOfEdges_Header.Clear()
                        lstOfEdges_Header.Add(CSng(lstOfDimensions("leftMargin")), "leftEdge")
                        lstOfEdges_Header.Add(CSng(lstOfDimensions("rightMargin")), "rightEdge")

                        If Not IsNothing(tbl) Then
                            dr = tbl.Rows.Item(2)
                            dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
                            dr.Height = 83
                            '
                            tblNested = dr.Cells.Item(1).Tables.Item(1)
                            tblNested.PreferredWidth = tbl.PreferredWidth - dr.Cells.Item(1).LeftPadding - dr.Cells.Item(1).RightPadding
                            '
                            dr = tbl.Rows.Item(3)
                            dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
                            dr.Height = 400
                            '
                            tblNested = dr.Cells.Item(1).Tables.Item(1)
                            tblNested.PreferredWidth = tbl.PreferredWidth - dr.Cells.Item(1).LeftPadding - dr.Cells.Item(1).RightPadding
                            '
                            drCellNested = tblNested.Rows.Item(3).Cells.Item(1)
                            '
                            If drCellNested.Range.InlineShapes.Count <> 0 Then
                                iShp = drCellNested.Range.InlineShapes.Item(1)
                                iShp.LockAspectRatio = True
                                iShp.Width = tblNested.PreferredWidth
                            End If
                        End If
                    '
                    Case "contBack"
                        objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Contacts_Lnd)
                        tbl = objGlobals.glb_tbls_AutoFitRegularTable(sect, False)
                        '
                        lstOfEdges_Header.Clear()
                        lstOfDimensions = objGlobals.glb_getDimensions_Contacts_Lnd()
                        lstOfEdges_Header.Add(CSng(lstOfDimensions("leftMargin")), "leftEdge")
                        lstOfEdges_Header.Add(CSng(lstOfDimensions("rightMargin")), "rightEdge")
                        '
                        If Not IsNothing(tbl) Then
                            dr = tbl.Rows.Item(2)
                            dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
                            dr.Height = 480
                            '
                            tblNested = dr.Cells.Item(1).Tables.Item(1)
                            tblNested.PreferredWidth = tbl.PreferredWidth - dr.Cells.Item(1).LeftPadding - dr.Cells.Item(1).RightPadding
                            '
                        End If
                    '
                    Case "toc"
                        'lstOfDimensions = objGlobals.glb_getDimensions_toc_Lnd()
                        'lstOfEdges.Add(CSng(lstOfDimensions("leftMargin")), "leftEdge")
                        'lstOfEdges.Add(CSng(lstOfDimensions("rightMargin")), "rightEdge")

                        objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Std_Lnd)
                        objTocMgr.toc_Styles_AdjustForOrientation(sect)
                        objTocMgr.toc_get_ContentsTable(sect, True)
                        '
                        lstOfEdges_Header.Clear()
                        lstOfEdges_Header.Add(objGlobals.glb_hfs_getHFTableEdge(sect, "header_leftEdge"), "leftEdge")
                        lstOfEdges_Header.Add(objGlobals.glb_hfs_getHFTableEdge(sect, "header_rightEdge"), "rightEdge")
                        '
                    '
                    Case "div", "divAp"
                        'WHen we rsize we go back to original measurements, so that we don't get a creeping error when
                        'someone repeatedly toggles between prt and lnd.... For the Dividers we wnat the headers to be flush
                        'with the margins, so we use the left and right margins as the edges
                        '
                        lstOfEdges_Header.Clear()
                        lstOfDimensions = objGlobals.glb_getDimensions_Divider_Lnd
                        lstOfEdges_Header.Add(CSng(lstOfDimensions("leftMargin")), "leftEdge")
                        lstOfEdges_Header.Add(CSng(lstOfDimensions("rightMargin")), "rightEdge")

                        objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Divider_Lnd)
                        objGlobals.glb_tbls_AutoFitRegularTable(sect, False)
                        '
                        'tbl = sect.Range.Tables.Item(1)
                        'tbl.PreferredWidth = objGlobals.glb_get_widthBetweenMargins(sect)
                    Case "glos"
                        objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Std_Lnd)
                        '
                        'Standard is the width between margins
                        For Each tbl In sect.Range.Tables
                            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                            tbl.PreferredWidth = 100
                            'objTables.tbl_setWidth_ToStandard(tbl)
                        Next

                    Case "briefFirstSection"
                        objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Std_Lnd)
                        objGlobals.glb_tbls_AutoFitRegularTable(sect, False)
                        '
                        objShp = objBackPanelMgr.pnl_resize_PanelToFillPage(sect)
                        objTocMgr.toc_Styles_AdjustForOrientation(sect)
                        'objShp.shp.Height = sect.PageSetup.PageHeight * 0.2

                        objShp.shp.Height = 168
                        '
                        objGlobals.glb_screen_updateLeaveAsItWas()
                        '
                    Case "brief"
                        objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Std_Lnd)


                    Case Else
                        objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Std_Lnd)
                        '
                        '
                        'Now resise all Findings, Recommendatiosn and Case Studies (half page) to fit
                        'the orienttation of the section they are in
                        objPlhBase.Plh_setAll_FindingEtc_Width(sect)
                        '
                        'Me.sect_Toggle_Width(sect, 0.0)

                        'lstOfEdges.Add(15.0, "leftEdge")
                        'lstOfEdges.Add(15.0, "rightEdge")

                        'Me.sect_Toggle_Width(sect, 0.0)
                End Select
                '
                'The resize functions will resize according to the left and right edges in lstOfEdges. if this
                'collection isNothing or it has no elements then the resize is done according to the fall back
                'standard of leftEdge = Me.objGlobals.glb_hfs_getHFTableEdge(sect, "header_leftEdge") which derives
                'from objGlobals._glb_header_leftEdge etc
                '
                objHfMgr.hf_headers_resize_all(sect, lstOfEdges_Header)
                objHfMgr.hf_footers_resize_all(sect, lstOfEdges_Footer)
                '
            End If



        Catch ex As Exception
            rslt = False
        End Try


        Return rslt
        '
    End Function
    '
    '
    Public Function sct_resize_ToPortrait(ByRef sect As Word.Section) As Boolean
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objGlobals As New cGlobals()
        Dim objChptBase As New cChptBase()
        Dim objChptBanner As New cChptBanner()
        Dim objPlhBase As New cPlHBase()
        Dim listOfPanels As New List(Of cShapeMgr)
        Dim objShp As cShapeMgr
        Dim objBackPanelMgr As New cBackPanelMgr()
        Dim objCpMgr As cCoverPageMgr
        Dim objTables As New cTablesMgr()
        Dim lstOfDimensions As New Collection()
        Dim lstOfEdges_Header As New Collection()
        Dim lstOfEdges_Footer As New Collection()
        '
        Dim objTocMgr As New cTOCMgr()
        Dim objDiv As New cChptDivider()
        Dim objApp As New cChptApp()
        Dim objisOK As New cIsOKToDo()
        Dim strType As String
        Dim dr As Word.Row
        Dim iShp As Word.InlineShape
        Dim strText As String
        '
        Dim tbl, tblNested As Word.Table
        Dim drCellNested As Word.Cell
        Dim rslt As Boolean
        '
        rslt = False
        strType = ""
        strText = ""
        '
        Try
            'We need to set measurements to a portrait type page so that the
            'WdOrientation.wdOrientPortrait doesn't throw a fault. We also adjust any columns that 
            'may have been set for a Landscape page
            objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Std_Prt)
            If sect.PageSetup.TextColumns.Count > 1 Then
                sect.PageSetup.TextColumns.EvenlySpaced = True
            End If

            sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait
            listOfPanels = objBackPanelMgr.pnl_getBackPanel_PlaceHolders(sect)
            objBackPanelMgr.pnl_resize_PanelToFillPage(sect, listOfPanels)
            '
            'Get the type of section that the current selection is in
            strType = objisOK.isOKto_selection_isIn()
            '
            lstOfEdges_Footer.Clear()
            lstOfEdges_Header.Clear()
            '
            Select Case strType
                Case "caseStudy"
                    objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_CaseStudy_Prt())

                Case "cp"
                    objCpMgr = New cCoverPageMgr()
                    '
                    lstOfDimensions = objGlobals.glb_getDimensions_CoverPage_Prt()
                    objGlobals.glb_setDimensions(sect, lstOfDimensions)
                    '
                    'Just delete the small pictures, they can be reset from the main cover page
                    'menu options
                    objCpMgr.ChptBase_delete_SmallPicturePlaceHolders(sect)

                    'tbl = objGlobals.glb_tbls_AutoFitRegularTable(sect, False)
                    '
                    'lstOfEdges_Header.Clear()
                    'lstOfEdges_Header.Add(CSng(lstOfDimensions("leftMargin")), "leftEdge")
                    'lstOfEdges_Header.Add(CSng(lstOfDimensions("rightMargin")), "rightEdge")

                Case "contFront"
                    lstOfDimensions = objGlobals.glb_getDimensions_Contacts_Prt("front")

                    objGlobals.glb_setDimensions(sect, lstOfDimensions)
                    tbl = objGlobals.glb_tbls_AutoFitRegularTable(sect, False)
                    '
                    lstOfEdges_Header.Clear()
                    lstOfEdges_Header.Add(CSng(lstOfDimensions("leftMargin")), "leftEdge")
                    lstOfEdges_Header.Add(CSng(lstOfDimensions("rightMargin")), "rightEdge")

                    If Not IsNothing(tbl) Then
                        dr = tbl.Rows.Item(2)
                        dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
                        dr.Height = 135
                        '
                        tblNested = dr.Cells.Item(1).Tables.Item(1)
                        tblNested.PreferredWidth = tbl.PreferredWidth - dr.Cells.Item(1).LeftPadding - dr.Cells.Item(1).RightPadding
                        '
                        dr = tbl.Rows.Item(3)
                        dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
                        dr.Height = 595
                        '
                        tblNested = dr.Cells.Item(1).Tables.Item(1)
                        tblNested.PreferredWidth = tbl.PreferredWidth - dr.Cells.Item(1).LeftPadding - dr.Cells.Item(1).RightPadding
                        '
                        drCellNested = tblNested.Rows.Item(3).Cells.Item(1)
                        '
                        If drCellNested.Range.InlineShapes.Count <> 0 Then
                            iShp = drCellNested.Range.InlineShapes.Item(1)
                            iShp.LockAspectRatio = True
                            iShp.Width = tblNested.PreferredWidth
                        End If
                    End If
                    '
                Case "contBack"
                    objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Contacts_Prt("back"))
                    tbl = objGlobals.glb_tbls_AutoFitRegularTable(sect, False)
                    If Not IsNothing(tbl) Then
                        dr = tbl.Rows.Item(2)
                        dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
                        dr.Height = 721.6
                        '
                        tblNested = dr.Cells.Item(1).Tables.Item(1)
                        tblNested.PreferredWidth = tbl.PreferredWidth - dr.Cells.Item(1).LeftPadding - dr.Cells.Item(1).RightPadding
                        '
                    End If
                    '
                Case "toc"
                    objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Std_Prt())
                    objTocMgr.toc_Styles_AdjustForOrientation(sect)
                    objTocMgr.toc_get_ContentsTable(sect, True)
                    '
                    lstOfEdges_Header.Clear()
                    lstOfEdges_Header.Add(objGlobals.glb_hfs_getHFTableEdge(sect, "header_leftEdge"), "leftEdge")
                    lstOfEdges_Header.Add(objGlobals.glb_hfs_getHFTableEdge(sect, "header_rightEdge"), "rightEdge")
                    '
                    '
                Case "div", "divAp"
                    lstOfDimensions = objGlobals.glb_getDimensions_Divider_Prt
                    lstOfEdges_Header.Add(CSng(lstOfDimensions("leftMargin")), "leftEdge")
                    lstOfEdges_Header.Add(CSng(lstOfDimensions("rightMargin")), "rightEdge")

                    objGlobals.glb_setDimensions(sect, lstOfDimensions)
                    objGlobals.glb_tbls_AutoFitRegularTable(sect, False)

                    '
                    'tbl = sect.Range.Tables.Item(1)
                    'tbl.PreferredWidth = objGlobals.glb_get_widthBetweenMargins(sect)
                Case "glos"
                    objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Std_Prt)
                    '
                    'Standard is the width between margins
                    For Each tbl In sect.Range.Tables
                        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                        tbl.PreferredWidth = 100
                        'objTables.tbl_setWidth_ToStandard(tbl)
                    Next

                Case "briefFirstSection"
                    objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Std_Prt)
                    tbl = objGlobals.glb_tbls_AutoFitRegularTable(sect, False)
                    '
                    objShp = objBackPanelMgr.pnl_resize_PanelToFillPage(sect)
                    objTocMgr.toc_Styles_AdjustForOrientation(sect)

                    objShp.shp.Height = sect.PageSetup.PageHeight * 0.2
                    '

                    '
                    objGlobals.glb_screen_updateLeaveAsItWas()
                        '
                Case "brief"
                    objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Std_Prt)


                Case Else
                    objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Std_Prt)
                    '
                    'Now resise all Findings, Recommendatiosn and Case Studies (half page) to fit
                    'the orienttation of the section they are in
                    objPlhBase.Plh_setAll_FindingEtc_Width(sect)
                    '
                    'Me.sect_Toggle_Width(sect, 0.0)
            End Select
            '
            'Do the Header first
            'objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_PortraitStd())
            'objHfMgr.hf_headers_resize_all(sect, lstOfEdges_Header)
            'objHfMgr.hf_footers_resize_all(sect, lstOfEdges_Footer)

            objHfMgr.hf_headers_resize_all(sect, lstOfEdges_Header)
            objHfMgr.hf_footers_resize_all(sect, lstOfEdges_Footer)
            '
            'Now autofit the Banner (if it is there)
            'rng = sect.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'If rng.Tables.Count <> 0 Then
            'tbl = rng.Tables.Item(1)
            'If objChptBanner.bnr_is_Chapter_Bdy_or_ES_or_AP(tbl) Then
            'Me.objGlobals.glb_tbls_AutoFitBanner(tbl, False)
            'objChptBase.chptBase_Banner_Autofit(tbl, sect, False)
            ' If
            'objChptBase.chptBase_Banner_Autofit(tbl, sect, False)
            'End If
            '


        Catch ex As Exception
            rslt = False
        End Try


        Return rslt
    End Function


    '
    Public Function sect_reset_ToLandscape(ByRef sect As Word.Section) As Boolean
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objGlobals As New cGlobals()
        Dim objChptBase As New cChptBase()
        Dim objChptBanner As New cChptBanner()
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim rslt As Boolean
        '
        rslt = False
        '
        Try
            Call objHfMgr.hf_hfs_deleteAll(sect)
            sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape
            '
            'Do the Header first
            objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Std_Lnd)
            objHfMgr.hf_headers_insert(sect)
            objHfMgr.hf_footers_insert(sect)
            'Me.ChptBase_Set_DimensionsPrt_RptPrt(sect)
            'Me.ChptBase_HeaderFooter_InsertHeader(sect, True, False)
            '
            'Now do the Footer
            'rslt = Me.PageNumbering_Set_ForBody(sect, False)
            'Me.ChptBase_HeaderFooter_InsertFooter(sect)
            '
            'Now autofit the Banner
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            If rng.Tables.Count <> 0 Then
                tbl = rng.Tables.Item(1)
                If objChptBanner.bnr_is_Chapter_Bdy_or_ES_or_AP(tbl) Then
                    Me.objGlobals.glb_tbls_AutoFitBanner(tbl, False)
                    'objChptBase.chptBase_Banner_Autofit(tbl, sect, False)
                End If
                'objChptBase.chptBase_Banner_Autofit(tbl, sect, False)
            End If

        Catch ex As Exception
            rslt = False
        End Try


        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method will delete the specified section. Note it will also attempt to remove
    ''' specific controls if the section proves to be a Cover Page. It will return the next
    ''' section or the last section (i.e. if there is only one section left)
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function sct_delete_Section(ByRef sect As Word.Section) As Word.Section
        Dim myDoc As Word.Document
        Dim objCpMgr As New cCoverPageMgr()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim sectLast, sectDest, sectSrc As Word.Section
        Dim rng As Word.Range
        '
        myDoc = sect.Range.Document
        sectLast = myDoc.Sections.Last
        '
        If objCpMgr.cp_Bool_IsCoverPage(sect) Then
            'We need to delete the Controls from the document
            Try
                'Globals.ThisAddIn.Controls.Remove("Report Date")
                'Globals.ThisAddIn.Controls.Remove("Report")
            Catch ex As Exception

            End Try
        End If
        '
        If sect.Index <> sectLast.Index Then
            sect.Range.Delete()
            '
            'The cursor is left in the next section
            Me.objGlobals.glb_get_wrdSel()
            sect = Me.objGlobals.glb_get_wrdSel().Sections.Item(1)
            sectLast = Me.objGlobals.glb_get_wrdSel().Sections.Last

            If myDoc.Sections.Count = 1 Then
                'We have deleted down to the last section.. Just rest this last page to
                'the standard Portrait shape
                '
                Me.sct_reset_ToPortrait(sect)
                'MyBase.ChptBase_Reset_ToPortrait(sect)
                'MyBase.ChptBase_Reset_ToLandScape(sect)
                '
            End If
            '
        Else
            If myDoc.Sections.Count = 1 Then
                'We have deleted down to the last section.. We are trying to delete the last 
                'and only section in the document  Just reset this last page to
                'the standard Portrait shape
                '
                sect.Range.Delete()
                Me.sct_reset_ToPortrait(sect)

                'MyBase.ChptBase_Reset_ToPortrait(sect)
                MsgBox("A document must always have one section, so this section has been cleared and reset to the standard portrait")
                sect.Range.Style = myDoc.Styles.Item("Body Text")
                'MyBase.ChptBase_Reset_ToLandScape(sect)
                '
            Else
                'We are trying to delete the last section, bu there are other sections in the document
                sect.Range.Delete()
                objHFMgr.hf_hfs_deleteAll(sect)
                '
                sectDest = myDoc.Sections.Last
                sectSrc = Globals.ThisAddIn.Application.ActiveDocument.Sections.Item(sectDest.Index - 1)
                '
                Me.cloneSection(sectSrc, sectDest, True)
                '
                rng = sectDest.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.MoveStart(WdUnits.wdParagraph, -2)
                rng.Delete()
                sect = rng.Sections.Item(1)

            End If
        End If
        '

        '
        Return sect

    End Function
    '
    '
    Public Function sct_delete_allSectionContents(ByRef sect As Word.Section, Optional numParasLeft As Integer = 6, Optional strStyleOfParas As String = "Body Text") As Word.Range
        Dim rng As Word.Range
        Dim sectLast As Word.Section
        Dim myDoc As Word.Document
        Dim objParas As New cParas()
        '
        myDoc = sect.Range.Document
        rng = sect.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)

        sectLast = myDoc.Sections.Last
        '
        If sect.Index = sectLast.Index Then
            '
            rng.MoveEnd(WdUnits.wdStory)
            rng.Delete()
            rng.Style = myDoc.Styles("Body Text")
            '
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            objParas.paras_insert_numParas(sect, numParasLeft)

        Else
            'We are in a standard section. So delete paragraphs to the section boundary
            'and then add numParasLeft
            '
            'rng = sect.Range
            rng.MoveEnd(WdUnits.wdSection, 1)
            rng.MoveEnd(WdUnits.wdParagraph, -2)
            rng.Style = myDoc.Styles("Body Text")
            rng.Delete()
            '
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            objParas.paras_insert_numParas(sect, numParasLeft)
            '
            For i = 1 To numParasLeft
                'para = rng.Paragraphs.Add()
            Next
            '
            rng.Style = myDoc.Styles(strStyleOfParas)
            '
        End If
        '
        Return rng
    End Function
    '
    ''' <summary>
    ''' This method will place a section break at the end of the paragraph containing the current selectio
    ''' When finished, the selection is at the top of the new section and the return value Word.section is the
    ''' new section.. The top of the new section contains an empty paragraph (with the selection at the beginning)
    ''' and the style is the same as the original parent style
    ''' </summary>
    ''' <param name="strStyle"></param>
    ''' <param name="strBreakType"></param>
    ''' <param name="linkToPrevious"></param>
    ''' <returns></returns>
    Public Function sct_insert_SectionAtSelection(Optional strStyle As String = "Body Text", Optional strBreakType As String = "newPage", Optional linkToPrevious As Boolean = False) As Word.Section
        Dim objParas As New cParas()
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim sect As Word.Section = Nothing
        Dim rng As Word.Range
        Dim lastPara As Word.Paragraph
        '
        rng = objGlobals.glb_get_wrdSelRng()
        '
        '****
        'objParas.paras_insert_parasAfterSelectedPara(rng, 2)
        lastPara = objParas.paras_insert_parasAtSelection(rng, 2)
        '****
        '
        'Insert paragraphs. The selection is at the beginning of the last of the 2
        'empty paragraphs... The range rng is collapsed to the beginning of the last of the
        '2 paragarphs
        'lastPara = rng.Paragraphs.First
        rng = lastPara.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Move(WdUnits.wdParagraph, -1)
        rng.Style = rng.Document.Styles.Item(strStyle)
        rng = lastPara.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()
        '
        Select Case strBreakType
            Case "newPage"
                objGlobals.glb_get_wrdApp.Selection.InsertBreak(WdBreakType.wdSectionBreakNextPage)
            Case "oddPage"
                objGlobals.glb_get_wrdApp.Selection.InsertBreak(WdBreakType.wdSectionBreakOddPage)
            Case "evenPage"
                objGlobals.glb_get_wrdApp.Selection.InsertBreak(WdBreakType.wdSectionBreakEvenPage)
            Case Else
                objGlobals.glb_get_wrdApp.Selection.InsertBreak(WdBreakType.wdSectionBreakNextPage)
        End Select
        '
        'Set sect to the new section, the selection is at the top of the section
        sect = objGlobals.glb_get_wrdSect()
        '
        If Not linkToPrevious Then
            objHfMgr.hf_hfs_linkUnlinkAll(sect, linkToPrevious)
        End If

        Return sect
    End Function
    '    
    ''' <summary>
    ''' This method will insert a section break at the 'Selection'... If the selection is in the middle
    ''' of a paragraph of text it will insert the break at the end of the paragraph. The section returned is the
    ''' section new section. The selection is at the beginning of the new section.
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="numParas"></param>
    ''' <param name="strBreakType"></param>
    ''' <param name="linkToPrevious"></param>
    ''' <returns></returns>
    Public Function sct_insert_SectionAtSelection(ByRef rng As Word.Range, Optional numParas As Integer = 6, Optional strBreakType As String = "newPage", Optional linkToPrevious As Boolean = False) As Word.Section
        Dim myDoc As Word.Document
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objParas As New cParas()
        Dim objLogos As New cLogosMgr()
        Dim objPgNumMgr As New cPageNumberMgr()
        Dim objRptBrief As New cReportBrief()
        Dim objTimer As New cTimer()
        Dim sect As Word.Section
        Dim idx As Integer
        '
        'newPage, oddPage, evenPage
        '
        Try
            rng = objGlobals.glb_get_wrdApp.Selection.Range
            myDoc = rng.Document
            'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            'para = objParas.paras_insert_paraAfter_num(rng, 2)
            'rng = para.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            Select Case strBreakType
                Case "newPage"
                    objGlobals.glb_get_wrdApp.Selection.InsertBreak(WdBreakType.wdSectionBreakNextPage)
                Case "oddPage"
                    objGlobals.glb_get_wrdApp.Selection.InsertBreak(WdBreakType.wdSectionBreakOddPage)
                Case "evenPage"
                    objGlobals.glb_get_wrdApp.Selection.InsertBreak(WdBreakType.wdSectionBreakEvenPage)
                Case Else
                    objGlobals.glb_get_wrdApp.Selection.InsertBreak(WdBreakType.wdSectionBreakNextPage)
            End Select
            '
            '
            '
            'rng.Select()
            sect = objGlobals.glb_get_wrdSect()                                     'This is the new section
            objParas.paras_delete_Paragraphs(sect.Range, 6)
            'objParas.paras_insert_numParas(sect, 4, True)
            '
            objHfMgr.hf_hfs_linkUnlinkAll(sect, linkToPrevious)
            '
            idx = objGlobals.glb_get_wrdSel.Information(WdInformation.wdActiveEndSectionNumber)
            '
            '***** 20250710...I don't think this addresses the problem
            objWrkAround.wrk_fix_forSectionProblem(sect.Range.Document)
            '
            'Now make sure the header logo is the correct colour.. This is for those situations
            'such as the 'AA Brief' where we have a white logo on the first page. We don't wnat this
            'translating.. Then delete any artefact backitems (that are not a logo) from the section that
            'may have translated across from the first page of a Brief
            objRptBrief.brf_delete_backItems_And_FixLogos(sect)
            '
            objParas.paras_insertAfterDelete_numParas(sect)
            '
            'Me.sct_setSection_toNoDiffFirstPage(sect)
            'sect.PageSetup.DifferentFirstPageHeaderFooter = False
            '
        Catch ex As Exception
            sect = Nothing
            MsgBox("Unknown error in cSectionMgr.sct_insert_SectionAtSelection")
        End Try

        '
        Return sect
        '
    End Function
    '

    '
    Public Function sct_insert_SectionInFront_Lnd(ByRef rng As Word.Range, Optional numParas As Integer = 6, Optional strBreakType As String = "newPage", Optional linkToPrevious As Boolean = False) As Word.Section
        Dim myDoc As Word.Document
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim sect As Word.Section
        Dim rng2 As Word.Range
        Dim i As Integer
        Dim lst As Collection
        '
        'newPage, oddPage, evenPage
        '
        myDoc = rng.Document
        sect = rng.Sections.Item(1)
        '
        rng2 = sect.Range
        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        '
        sect = myDoc.Sections.Add(rng2, WdSectionStart.wdSectionNewPage)
        objHfMgr.hf_hfs_linkUnlinkAll(sect, linkToPrevious)
        '
        sect = myDoc.Sections.Item(sect.Index - 1)
        objHfMgr.hf_hfs_linkUnlinkAll(sect, linkToPrevious)
        '
        rng2 = sect.Range
        rng2.Style = myDoc.Styles.Item("Body Text")
        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        For i = 1 To numParas
            rng2.Paragraphs.Add(rng2)
        Next
        '
        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        '
        Select Case strBreakType
            Case "newPage"
                sect.PageSetup.SectionStart = WdSectionStart.wdSectionNewPage
            Case "oddPage"
                sect.PageSetup.SectionStart = WdSectionStart.wdSectionOddPage
            Case "evenPage"
                sect.PageSetup.SectionStart = WdSectionStart.wdSectionEvenPage
        End Select
        '
        'Make sure its Landscape
        sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape
        lst = objGlobals.glb_getDimensions_Std_Lnd()
        objGlobals.glb_setDimensions(sect, lst)
        '
        objHfMgr.hf_headers_resize_all(sect)
        objHfMgr.hf_footers_resize_all(sect)

        Return sect

    End Function

    '
    ''' <summary>
    ''' This method will insert a new (empty) section in front of an existing (base) section. The
    ''' base section is the one containing the rng that is input to the method... It will also return
    ''' the new (empty) section... without chnaging the input range rng... It is the base for Chapters etc
    ''' 
    ''' The selection is set in the new (empty) section. Note that the method will insert a specified
    ''' number of paragraphs (numParas which defaults to 6) in the new (empty) section. It will also set the
    ''' new (empty) section) to a "newPage", "oddPage" or "evenPage" section depending on the value of strBreakType.
    ''' The default is 'newPage'.
    ''' 
    ''' The following base section (which contains rng) is always 'newPage'
    ''' 
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="numParas"></param>
    ''' <param name="strBreakType"></param>
    ''' <returns></returns>
    Public Function xsct_insert_SectionInFront_Old(ByRef rng As Word.Range, Optional numParas As Integer = 6, Optional strBreakType As String = "newPage", Optional linkToPrevious As Boolean = False) As Word.Section
        Dim myDoc As Word.Document
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim sect As Word.Section
        Dim rng2 As Word.Range
        Dim i As Integer
        '
        'newPage, oddPage, evenPage
        '
        myDoc = rng.Document
        sect = rng.Sections.Item(1)
        '
        rng2 = sect.Range
        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        '
        sect = myDoc.Sections.Add(rng2, WdSectionStart.wdSectionNewPage)
        objHfMgr.hf_hfs_linkUnlinkAll(sect, linkToPrevious)
        '
        sect = myDoc.Sections.Item(sect.Index - 1)
        objHfMgr.hf_hfs_linkUnlinkAll(sect, linkToPrevious)
        '
        '
        rng2 = sect.Range
        rng2.Style = myDoc.Styles.Item("Body Text")
        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        For i = 1 To numParas
            rng2.Paragraphs.Add(rng2)
        Next
        '
        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        '
        Select Case strBreakType
            Case "newPage"
                sect.PageSetup.SectionStart = WdSectionStart.wdSectionNewPage
            Case "oddPage"
                sect.PageSetup.SectionStart = WdSectionStart.wdSectionOddPage
            Case "evenPage"
                sect.PageSetup.SectionStart = WdSectionStart.wdSectionEvenPage
        End Select
        '

        Return sect

    End Function
    '
    ''' <summary>
    ''' This method will return the list of Margin Dimensions depending up the value of
    ''' strOrientation ("flow", "prt", "lnd", "lndRptChpt", "lndRptFollower"). Flow will return nothing
    ''' as we are expecting any new section to take on the characteristics of the parent section
    ''' </summary>
    ''' <param name="strOrientation"></param>
    ''' <returns></returns>
    Public Function sct_get_lstOfMarginDimensions(strOrientation As String) As Collection
        Dim lstOfMarginDimensions As Collection
        '
        lstOfMarginDimensions = Nothing
        '
        Select Case strOrientation
            Case "flow"
            Case "prt"
                lstOfMarginDimensions = Me.objGlobals.glb_getDimensions_Std_Prt()
            Case "prtDiv"
                lstOfMarginDimensions = Me.objGlobals.glb_getDimensions_Divider_Prt()
            Case "lnd"
                lstOfMarginDimensions = Me.objGlobals.glb_getDimensions_Std_Lnd()
            Case "lndDiv"
                lstOfMarginDimensions = Me.objGlobals.glb_getDimensions_Divider_Lnd()
            Case "lndRptChpt"
                lstOfMarginDimensions = Me.objGlobals.glb_getDimensions_Landscape_LndRpt_ChptPage()
            Case "lndRptFollower"
                lstOfMarginDimensions = Me.objGlobals.glb_getDimensions_Landscape_LndRpt_followerPage()
            Case Else

        End Select
        '
        Return lstOfMarginDimensions
        '
    End Function
    '
    ''' <summary>
    ''' This method will insert a new section (not bounded) either behind or in front of the section sect. On return sect will be
    ''' set to the new section (and NOT the initial section). . The section break defaults to "newPage", unless specified
    ''' as "oddPage or "evenPage". The Headers/Footers default to unlinked. The parameter "strOrientation" is multi faceted. It tells
    ''' us whether
    ''' -   "flow"                                      Section inherits parent section setup and Header/Footers, but dimensions are replaced by lstOfMarginDimensions if it is NOT Nothing
    ''' -   "lnd", "lndRptChpt", "lndRptFollower"       Forced to landscape and will have the default margins as found in "sct_get_lstOfMarginDimensions"
    ''' -   "prt", "prtDiv"                             Forced to portrait and will have the default margins as found in "sct_get_lstOfMarginDimensions"
    ''' 
    ''' The parameter "lstOfMarginDimensions" will override the defaults if it is NOT nothing
    ''' </summary>
    ''' <param name="placeBehind"></param>
    ''' <param name="sect"></param>
    ''' <param name="numParas"></param>
    ''' <param name="strBreakType"></param>
    ''' <param name="linkToPrevious"></param>
    ''' <param name="strOrientation"></param>
    ''' <param name="lstOfMarginDimensions"></param>
    ''' <returns></returns>
    Public Function sct_insert_Section(placeBehind As Boolean, ByRef sect As Word.Section, Optional numParas As Integer = 6, Optional strBreakType As String = "newPage", Optional linkToPrevious As Boolean = False, Optional strOrientation As String = "flow", Optional ByRef lstOfMarginDimensions As Collection = Nothing) As Word.Section
        Dim myDoc As Word.Document
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objParas As New cParas()
        Dim rng2 As Word.Range
        Dim objWtrMrks As New cWaterMarks()
        '
        myDoc = sect.Range.Document
        rng2 = Nothing
        '
        If placeBehind Then
            'If placebehind, then sect = new section
            '-  must unlink the new section (i.e. the one behind)
            '-  and maybe for good luck unlink the initial section
            '
            rng2 = sect.Range
            rng2.Collapse(WdCollapseDirection.wdCollapseEnd)
            rng2.Move(WdUnits.wdCharacter, -2)
            '
            rng2.Select()
            sect = sct_insert_SectionAtSelection(, strBreakType, linkToPrevious)                    'The new section behind uses layout and header footers of src section
            '
            objHfMgr.hf_hfs_linkUnlinkAll(sect, linkToPrevious)
            objHfMgr.hf_hfs_linkUnlinkAll(myDoc.Sections.Item(sect.Index - 1), linkToPrevious)
            objParas.paras_insert_numParas(sect, 6)                                                  'Put 6 paras in the section
            '
            '
        Else
            'if placeInFront, then sect = the initial section,
            '-  must unlink initial section
            '-  unlink section in fron for good measure
            '
            rng2 = sect.Range
            rng2.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            sect = objGlobals.glb_add_sectionBreak(rng2, strBreakType)
            '
            objHfMgr.hf_hfs_linkUnlinkAll(sect, linkToPrevious)
            sect = myDoc.Sections.Item(sect.Index - 1)                      'Set sect to the inserted section
            '
            objHfMgr.hf_hfs_linkUnlinkAll(sect, linkToPrevious)             'Unlink this for safety?
            objParas.paras_insert_numParas(sect)                            'Put 6 paras in the section
            '
            '
        End If
        '
        Select Case strOrientation
            Case "flow"
                'This option allows for a "flow" section, but with a custome set of margins
                'and appropriate Header/Footers... Only called in if the method is supplied with a
                'a lstOfMarginDimensions
                If Not IsNothing(lstOfMarginDimensions) Then
                    objGlobals.glb_setDimensions(sect, lstOfMarginDimensions)
                    '
                    objHfMgr.hf_headers_insert(sect)
                    objHfMgr.hf_footers_insert(sect)
                End If

            Case "lnd", "lndRptChpt", "lndRptFollower"
                sect.PageSetup.DifferentFirstPageHeaderFooter = False
                sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape
                If IsNothing(lstOfMarginDimensions) Then lstOfMarginDimensions = Me.sct_get_lstOfMarginDimensions(strOrientation)
                objGlobals.glb_setDimensions(sect, lstOfMarginDimensions)
                '
                objHfMgr.hf_headers_insert(sect)
                objHfMgr.hf_footers_insert(sect)
                    '
                    'objHfMgr.hf_headers_resize_all(sect)
                    'objHfMgr.hf_footers_resize_all(sect)
            Case "prt", "prtDiv"
                sect.PageSetup.DifferentFirstPageHeaderFooter = False
                sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait
                If IsNothing(lstOfMarginDimensions) Then lstOfMarginDimensions = Me.sct_get_lstOfMarginDimensions(strOrientation)
                objGlobals.glb_setDimensions(sect, lstOfMarginDimensions)
                '
                objHfMgr.hf_headers_insert(sect)
                objHfMgr.hf_footers_insert(sect)
                'objHfMgr.hf_headers_resize_all(sect)
                'objHfMgr.hf_footers_resize_all(sect)
            Case "lndDiv"
                sect.PageSetup.DifferentFirstPageHeaderFooter = False
                sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape
                If IsNothing(lstOfMarginDimensions) Then lstOfMarginDimensions = Me.sct_get_lstOfMarginDimensions(strOrientation)
                objGlobals.glb_setDimensions(sect, lstOfMarginDimensions)
                '
                objHfMgr.hf_headers_insert(sect)
                objHfMgr.hf_footers_insert(sect)
                'objHfMgr.hf_headers_resize_all(sect)
                'objHfMgr.hf_footers_resize_all(sect)

            Case Else

        End Select
        '
finis:
        '
        objWtrMrks.waterMarks_Remove_VersionMark(sect)
        '
        Return sect

    End Function

    '
    ''' <summary>
    ''' This method will insert a new section (not bounded) either behind or in front of the section sect. On return sect will be
    ''' set to the new section (and NOT the initial section). . The section break defaults to "newPage", unless specified
    ''' as "oddPage or "evenPage". The Headers/Footers default to unlinked. The parameter "strOrientation" is multi faceted. It tells
    ''' us whether
    ''' -   "flow"                                      Section inherits parent section setup and Header/Footers, but dimensions are replaced by lstOfMarginDimensions if it is NOT Nothing
    ''' -   "lnd", "lndRptChpt", "lndRptFollower"       Forced to landscape and will have the default margins as found in "sct_get_lstOfMarginDimensions"
    ''' -   "prt", "prtDiv"                             Forced to portrait and will have the default margins as found in "sct_get_lstOfMarginDimensions"
    ''' 
    ''' The parameter "lstOfMarginDimensions" will override the defaults if it is NOT nothing
    ''' </summary>
    ''' <param name="placeBehind"></param>
    ''' <param name="sect"></param>
    ''' <param name="numParas"></param>
    ''' <param name="strBreakType"></param>
    ''' <param name="linkToPrevious"></param>
    ''' <param name="strOrientation"></param>
    ''' <param name="lstOfMarginDimensions"></param>
    ''' <returns></returns>
    Public Function xsct_insert_Section(placeBehind As Boolean, ByRef sect As Word.Section, Optional numParas As Integer = 6, Optional strBreakType As String = "newPage", Optional linkToPrevious As Boolean = False, Optional strOrientation As String = "flow", Optional ByRef lstOfMarginDimensions As Collection = Nothing) As Word.Section
        Dim myDoc As Word.Document
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim rng2 As Word.Range
        Dim i As Integer
        Dim objWtrMrks As New cWaterMarks()
        '
        myDoc = sect.Range.Document
        'sect = rng.Sections.Item(1)
        rng2 = Nothing
        '
        If placeBehind Then
            rng2 = sect.Range
            rng2.Collapse(WdCollapseDirection.wdCollapseEnd)
            rng2.Paragraphs.Add(rng2)
            rng2.Paragraphs.Add(rng2)
            rng2.Collapse(WdCollapseDirection.wdCollapseEnd)
            'rng2.Move(WdUnits.wdParagraph, -1)
            'rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        Else
            rng2 = sect.Range
            rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        End If
        '
        'newPage, oddPage, evenPage
        Select Case strBreakType
            Case "newPage"
                sect = myDoc.Sections.Add(rng2, WdSectionStart.wdSectionNewPage)
            Case "oddPage"
                sect = myDoc.Sections.Add(rng2, WdSectionStart.wdSectionOddPage)
            Case "evenPage"
                sect = myDoc.Sections.Add(rng2, WdSectionStart.wdSectionEvenPage)
            Case Else
                sect = myDoc.Sections.Add(rng2, WdSectionStart.wdSectionNewPage)
        End Select
        '
        sect = myDoc.Sections.Item(sect.Index - 1)
        '
        'sect = myDoc.Sections.Item(sect.Index - 1)
        'sect = myDoc.Sections.Add(rng2, WdSectionStart.wdSectionNewPage)
        objHfMgr.hf_hfs_linkUnlinkAll(sect, linkToPrevious)
        '
        '**** Fix is here
        objHfMgr.hf_hfs_linkUnlinkAll(myDoc.Sections.Item(sect.Index + 1), linkToPrevious)
        '
        If Not placeBehind Then
            'objHfMgr.hf_hfs_linkUnlinkAll(myDoc.Sections.Item(sect.Index + 1), linkToPrevious)
        End If
        '
        '****
        '
        rng2 = sect.Range
        rng2.Style = myDoc.Styles.Item("Body Text")
        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        For i = 1 To numParas
            rng2.Paragraphs.Add(rng2)
        Next
        '
        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        Select Case strOrientation
            Case "flow"
                'This option allows for a "flow" section, but with a custome set of margins
                'and appropriate Header/Footers... Only called in if the method is supplied with a
                'a lstOfMarginDimensions
                If Not IsNothing(lstOfMarginDimensions) Then
                    objGlobals.glb_setDimensions(sect, lstOfMarginDimensions)
                    '
                    objHfMgr.hf_headers_insert(sect)
                    objHfMgr.hf_footers_insert(sect)
                End If

            Case "lnd", "lndRptChpt", "lndRptFollower"
                sect.PageSetup.DifferentFirstPageHeaderFooter = False
                sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape
                If IsNothing(lstOfMarginDimensions) Then lstOfMarginDimensions = Me.sct_get_lstOfMarginDimensions(strOrientation)
                objGlobals.glb_setDimensions(sect, lstOfMarginDimensions)
                '
                objHfMgr.hf_headers_insert(sect)
                objHfMgr.hf_footers_insert(sect)
                    '
                    'objHfMgr.hf_headers_resize_all(sect)
                    'objHfMgr.hf_footers_resize_all(sect)
            Case "prt", "prtDiv"
                sect.PageSetup.DifferentFirstPageHeaderFooter = False
                sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait
                If IsNothing(lstOfMarginDimensions) Then lstOfMarginDimensions = Me.sct_get_lstOfMarginDimensions(strOrientation)
                objGlobals.glb_setDimensions(sect, lstOfMarginDimensions)
                '
                objHfMgr.hf_headers_insert(sect)
                objHfMgr.hf_footers_insert(sect)
                'objHfMgr.hf_headers_resize_all(sect)
                'objHfMgr.hf_footers_resize_all(sect)
            Case "lndDiv"
                sect.PageSetup.DifferentFirstPageHeaderFooter = False
                sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape
                If IsNothing(lstOfMarginDimensions) Then lstOfMarginDimensions = Me.sct_get_lstOfMarginDimensions(strOrientation)
                objGlobals.glb_setDimensions(sect, lstOfMarginDimensions)
                '
                objHfMgr.hf_headers_insert(sect)
                objHfMgr.hf_footers_insert(sect)
                'objHfMgr.hf_headers_resize_all(sect)
                'objHfMgr.hf_footers_resize_all(sect)

            Case Else

        End Select
        '
        objWtrMrks.waterMarks_Remove_VersionMark(sect)
        '
        Return sect

    End Function
    '
    '
    ''' <summary>
    ''' This method will reproduce the settings of the source Section (srcSection) in the
    ''' destination Section (destSection
    ''' </summary>
    ''' <param name="srcSection"></param>
    ''' <param name="destSection"></param>
    Public Sub cloneSection(ByRef srcSection As Section, ByRef destSection As Section)
        '
        Dim strOrientation As String
        Dim objHfMgr As New cHeaderFooterMgr()
        '
        strOrientation = "portrait"
        If srcSection.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then strOrientation = "landscape"
        '
        destSection.PageSetup.PaperSize = srcSection.PageSetup.PaperSize
        destSection.PageSetup.Orientation = srcSection.PageSetup.Orientation
        destSection.PageSetup.GutterPos = srcSection.PageSetup.GutterPos
        destSection.PageSetup.DifferentFirstPageHeaderFooter = srcSection.PageSetup.DifferentFirstPageHeaderFooter
        destSection.PageSetup.OddAndEvenPagesHeaderFooter = srcSection.PageSetup.OddAndEvenPagesHeaderFooter
        '
        'Copy page dimensions
        destSection.PageSetup.TopMargin = srcSection.PageSetup.TopMargin
        destSection.PageSetup.LeftMargin = srcSection.PageSetup.LeftMargin
        destSection.PageSetup.BottomMargin = srcSection.PageSetup.BottomMargin
        destSection.PageSetup.RightMargin = srcSection.PageSetup.RightMargin
        destSection.PageSetup.Gutter = srcSection.PageSetup.Gutter
        destSection.PageSetup.HeaderDistance = srcSection.PageSetup.HeaderDistance
        destSection.PageSetup.FooterDistance = srcSection.PageSetup.FooterDistance
        '
        Select Case objHfMgr.hf_get_HeaderFooterType(srcSection)
            Case "DiffFirstPage-Not"
            Case "DiffFirstPage"
            Case "OddAndEven"
            Case "DiffFirstPage+OddAndEven"
        End Select

        objHfMgr.hf_HF_CopyHeaderFooter("header", srcSection, destSection)
        objHfMgr.hf_HF_CopyHeaderFooter("footer", srcSection, destSection)
        '

    End Sub


    Public Function xsct_insert_SectionInFront(ByRef rng As Word.Range, Optional numParas As Integer = 6, Optional strBreakType As String = "newPage", Optional linkToPrevious As Boolean = False, Optional strOrientation As String = "flow", Optional ByRef lstOfMargins As Collection = Nothing) As Word.Section
        Dim myDoc As Word.Document
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim sect As Word.Section
        Dim rng2 As Word.Range
        Dim i As Integer
        Dim objWtrMrks As New cWaterMarks()
        '
        myDoc = rng.Document
        sect = rng.Sections.Item(1)
        '
        'newPage, oddPage, evenPage
        rng2 = sect.Range
        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        Select Case strBreakType
            Case "newPage"
                sect = myDoc.Sections.Add(rng2, WdSectionStart.wdSectionNewPage)
            Case "oddPage"
                sect = myDoc.Sections.Add(rng2, WdSectionStart.wdSectionOddPage)
            Case "evenPage"
                sect = myDoc.Sections.Add(rng2, WdSectionStart.wdSectionEvenPage)
            Case Else
                sect = myDoc.Sections.Add(rng2, WdSectionStart.wdSectionNewPage)
        End Select
        '
        'sect = myDoc.Sections.Add(rng2, WdSectionStart.wdSectionNewPage)
        objHfMgr.hf_hfs_linkUnlinkAll(sect, linkToPrevious)
        '
        sect = myDoc.Sections.Item(sect.Index - 1)
        objHfMgr.hf_hfs_linkUnlinkAll(sect, linkToPrevious)
        '
        '
        rng2 = sect.Range
        rng2.Style = myDoc.Styles.Item("Body Text")
        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        For i = 1 To numParas
            rng2.Paragraphs.Add(rng2)
        Next
        '
        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        Select Case strOrientation
            Case "flow"
            Case "lnd", "lndRptChpt", "lndRptFollower"
                sect.PageSetup.DifferentFirstPageHeaderFooter = False
                sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape
                If Not IsNothing(lstOfMargins) Then
                    objGlobals.glb_setDimensions(sect, lstOfMargins)
                    '
                    objHfMgr.hf_headers_insert(sect)
                    objHfMgr.hf_footers_insert(sect)
                    '
                    'objHfMgr.hf_headers_resize_all(sect)
                    'objHfMgr.hf_footers_resize_all(sect)
                End If
            Case "prt"
                sect.PageSetup.DifferentFirstPageHeaderFooter = False
                sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait
                If IsNothing(lstOfMargins) Then lstOfMargins = objGlobals.glb_getDimensions_Std_Prt()
                objGlobals.glb_setDimensions(sect, lstOfMargins)
                '
                objHfMgr.hf_headers_insert(sect)
                objHfMgr.hf_footers_insert(sect)
                'objHfMgr.hf_headers_resize_all(sect)
                'objHfMgr.hf_footers_resize_all(sect)
            Case Else

        End Select
        '
        objWtrMrks.waterMarks_Remove_VersionMark(sect)
        '
        Return sect

    End Function
    '
    Public Function sct_insert_SectionBehind(ByRef rng As Word.Range, Optional numParas As Integer = 6, Optional strBreakType As String = "newPage", Optional linkToPrevious As Boolean = False, Optional strMode As String = "flow", Optional ByRef lstOfMargins As Collection = Nothing) As Word.Section
        Dim myDoc As Word.Document
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objParas As New cParas()
        Dim sect As Word.Section
        Dim rng2 As Word.Range
        Dim i As Integer
        Dim objWtrMrks As New cWaterMarks()
        '
        'newPage, oddPage, evenPage
        '
        '
        myDoc = rng.Document
        sect = rng.Sections.Item(1)
        '
        'If its the last section don't do anything
        If sect.Index = myDoc.Sections.Last.Index Then GoTo finis
        '
        rng2 = sect.Range
        rng2.Collapse(WdCollapseDirection.wdCollapseEnd)
        'rng2.Move(WdUnits.wdParagraph, -1)
        rng2.Move(WdUnits.wdCharacter, -1)
        '
        'rng2.Paragraphs.Add()
        'rng2.Paragraphs.Add()

        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        'rng2.Move(WdUnits.wdParagraph, 1)
        '
        Select Case strBreakType
            Case "newPage"
                sect = myDoc.Sections.Add(rng2, WdSectionStart.wdSectionNewPage)
            Case "oddPage"
                sect = myDoc.Sections.Add(rng2, WdSectionStart.wdSectionOddPage)
            Case "evenPage"
                sect = myDoc.Sections.Add(rng2, WdSectionStart.wdSectionEvenPage)
            Case Else
                sect = myDoc.Sections.Add(rng2, WdSectionStart.wdSectionNewPage)
        End Select        '
        '
        objHfMgr.hf_hfs_linkUnlinkAll(sect, linkToPrevious)
        '
        'sect = myDoc.Sections.Item(sect.Index - 1)
        'objHfMgr.hf_hfs_linkUnlinkAll(sect, linkToPrevious)

        rng2 = sect.Range
        rng2.Style = myDoc.Styles.Item("Body Text")
        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        objParas.paras_insert_numParas(sect, numParas)

        'sect.Range.Select()
        'GoTo finis
        'objParas.paras_insert_numParas(rng2, numParas)
        'objParas.paras_insert_numParas(rng2, numParas)

        '
        For i = 1 To numParas
            'rng2.Paragraphs.Add(rng2)
        Next
        '
        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        Select Case strMode
            Case "flow"
            Case "lnd"
                sect.PageSetup.DifferentFirstPageHeaderFooter = False
                sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape
                '
                If IsNothing(lstOfMargins) Then lstOfMargins = objGlobals.glb_getDimensions_Std_Lnd()
                objGlobals.glb_setDimensions(sect, lstOfMargins)
                '
                objHfMgr.hf_headers_insert(sect)
                objHfMgr.hf_footers_insert(sect)

            Case "prt"
                sect.PageSetup.DifferentFirstPageHeaderFooter = False
                sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait
                '
                If IsNothing(lstOfMargins) Then lstOfMargins = objGlobals.glb_getDimensions_Std_Prt()
                objGlobals.glb_setDimensions(sect, lstOfMargins)
                '
                objHfMgr.hf_headers_insert(sect)
                objHfMgr.hf_footers_insert(sect)
                '
                'objHfMgr.hf_headers_resize_all(sect)
                'objHfMgr.hf_footers_resize_all(sect)
            Case Else

        End Select
        '
        'objParas.paras_insert_numParas(sect, 12)
        '
        'Select Case strBreakType
        'Case "newPage"
        'sect.PageSetup.SectionStart = WdSectionStart.wdSectionNewPage
        ' Case "oddPage"
        'sect.PageSetup.SectionStart = WdSectionStart.wdSectionOddPage
        'Case "evenPage"
        'sect.PageSetup.SectionStart = WdSectionStart.wdSectionEvenPage
        'End Select
        '
        'objWtrMrks.waterMarks_Remove_VersionMark(sect)

finis:
        Return sect

    End Function


    '
    Public Function xsct_insert_SectionBehind_Old(ByRef rng As Word.Range, Optional numParas As Integer = 6, Optional strBreakType As String = "newPage", Optional linkToPrevious As Boolean = False) As Word.Section
        Dim myDoc As Word.Document
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim sect As Word.Section
        Dim rng2 As Word.Range
        Dim i As Integer
        '
        'newPage, oddPage, evenPage
        '
        '
        myDoc = rng.Document
        sect = rng.Sections.Item(1)
        '
        rng2 = sect.Range
        rng2.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng2.Move(WdUnits.wdParagraph, -1)
        rng2.Paragraphs.Add()
        rng2.Paragraphs.Add()

        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        rng2.Move(WdUnits.wdParagraph, 1)
        '
        'rng2.Select()
        'GoTo finis
        '
        '
        sect = myDoc.Sections.Add(rng2, WdSectionStart.wdSectionNewPage)
        objHfMgr.hf_hfs_linkUnlinkAll(sect, linkToPrevious)
        '
        'sect = myDoc.Sections.Item(sect.Index - 1)
        'objHfMgr.hf_hfs_linkUnlinkAll(sect, linkToPrevious)

        rng2 = sect.Range
        rng2.Style = myDoc.Styles.Item("Body Text")
        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        For i = 1 To numParas
            rng2.Paragraphs.Add(rng2)
        Next
        '
        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        '
        Select Case strBreakType
            Case "newPage"
                sect.PageSetup.SectionStart = WdSectionStart.wdSectionNewPage
            Case "oddPage"
                sect.PageSetup.SectionStart = WdSectionStart.wdSectionOddPage
            Case "evenPage"
                sect.PageSetup.SectionStart = WdSectionStart.wdSectionEvenPage
        End Select
        '
finis:
        Return sect

    End Function
    '
    ''' <summary>
    ''' This method will adjust inserted bound sections to have different first page = false.. Used only when
    ''' we are inserting bound sections into a bried
    ''' of false
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub sct_adjustInsertedSections_forBrief(ByRef sect As Word.Section, Optional doTwoSections As Boolean = True)
        Dim objRptMgr As New cReport()
        Dim myDoc As Word.Document
        Dim strRptMode As String
        '
        myDoc = sect.Range.Document
        '
        strRptMode = objRptMgr.Rpt_Mode_Get()
        '
        Select Case strRptMode
            Case objRptMgr.rpt_isBrief, objRptMgr.modeShort
                'We adjust the inserted section (sect) and the following section which bounds it
                If doTwoSections Then
                    'For bounded sections
                    sect.PageSetup.DifferentFirstPageHeaderFooter = False
                    myDoc.Sections(sect.Index + 1).PageSetup.DifferentFirstPageHeaderFooter = False
                Else
                    'For insert section at selection
                    sect.PageSetup.DifferentFirstPageHeaderFooter = False
                End If
                '
        End Select
        '
    End Sub
    '
    ''' <summary>
    ''' This method will insert a 'Bounded Section' at rng (collapsed to the nearest end of paragraph).
    ''' strOrientation = "std_Prt", "std_Lnd", "lnd_Lnd" which determines orientation and measurements.
    ''' The Start of Section is determined by strBreak = "newPage", "oddPage", "evenPage"...  Returns the
    ''' new bounded section with the selection also in this section
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="strOrientation"></param>
    ''' <param name="numParas"></param>
    ''' <param name="strBreakType"></param>
    ''' <param name="linkToPrevious"></param>
    ''' <returns></returns>
    Public Function sct_insert_SectionBounded(ByRef rng As Word.Range, strOrientation As String, Optional numParas As Integer = 6, Optional strBreakType As String = "newPage", Optional linkToPrevious As Boolean = False) As Word.Section
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objChptBase As New cChptBase()
        Dim lstOfDimensions As New Collection()
        Dim objBackPanel As New cBackPanelMgr()
        Dim objWrkAround As New cWorkArounds()
        Dim objParas As New cParas()
        Dim hf As Word.HeaderFooter
        Dim sect, sectMid, sectEnd, sectStart As Word.Section
        Dim myDoc As Word.Document
        Dim tbl As Word.Table
        Dim tblWidth As Single
        '
        myDoc = rng.Document
        sect = Nothing
        sectStart = Nothing
        sectMid = Nothing
        sectEnd = Nothing
        tbl = Nothing
        tblWidth = 0.0
        '
        Try
            sect = sct_insert_SectionAtSelection()
            rng = objGlobals.glb_get_wrdSelRng()
            objParas.paras_insert_parasAfterSelectedPara(rng, 2)                                     'With the para already there, that leaves 3
            '
            sectMid = objGlobals.glb_get_wrdSect()
            rng = sectMid.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Select()
            rng.MoveEnd(WdUnits.wdParagraph, 2)
            rng.Style = rng.Document.Styles.Item("Body Text")
            '
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            rng.MoveEnd(WdUnits.wdParagraph, 1)
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            rng.Select()
            sectEnd = sct_insert_SectionAtSelection()
            sectStart = rng.Document.Sections.Item(sectMid.Index - 1)
            '
            rng = sectMid.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Select()
            '
            'Now we make certain that the new bounded section is padded out with
            'empty paras
            objParas.paras_insertAfterDelete_numParas(sectMid)
            '
            '
            Select Case strOrientation
                Case "cstudy_Prt", "cstudy_Lnd"
                    Select Case strOrientation
                        Case "cstudy_Prt"
                            sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait
                            objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_CaseStudy_Prt())
                        Case "cstudy_Lnd"
                            sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape
                            objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_CaseStudy_Lnd())
                    End Select
                    '
                    objHFMgr.hf_headers_resize_all(sect)
                    objHFMgr.hf_footers_resize_all(sect)
                    '
                    hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                    If sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Exists Then
                        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                    End If
                    '
                    objBackPanel.pnl_BackPanel_Insert(hf, objGlobals._glb_colour_CaseStudy_Grey)
                    '
                    'Now do the first page if it exists
                    'If sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Exists Then
                    'hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                    'objBackPanel.pnl_BackPanel_Insert(hf, objGlobals._glb_colour_CaseStudy_Grey)
                    'End If
                    '
                    'Workaround for Cursor race condition
                    objWrkAround.wrk_fix_forCursorRace()
                    '
                Case "xxcstudy_Lnd"
                    sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape
                    '
                    objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_CaseStudy_Lnd())
                    objHFMgr.hf_headers_resize_all(sect)
                    objHFMgr.hf_footers_resize_all(sect)
                    '
                    objBackPanel.pnl_BackPanel_Insert(sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary), objGlobals._glb_colour_CaseStudy_Grey)
                    'Now do the first page if it exists
                    If sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Exists Then
                        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                        objBackPanel.pnl_BackPanel_Insert(hf, objGlobals._glb_colour_CaseStudy_Grey)
                    End If
                    '
                    'Workaround for Cursor race condition
                    objWrkAround.wrk_fix_forCursorRace()
                    '
                Case "std_Prt"
                    sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait
                    objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Std_Prt())
                    objHFMgr.hf_headers_resize_all(sect)
                    objHFMgr.hf_footers_resize_all(sect)
                Case "std_Lnd"
                    sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape
                    objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Std_Lnd)
                    objHFMgr.hf_headers_resize_all(sect)
                    objHFMgr.hf_footers_resize_all(sect)
                Case "lnd_Lnd"
                    sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape
                    objGlobals.glb_setDimensions(sect, objGlobals.glb_getDimensions_Landscape_LndRpt_followerPage)
                    objHFMgr.hf_headers_resize_all(sect)
                    objHFMgr.hf_footers_resize_all(sect)
            End Select
            '
            'Now make certain that the bounded section and the section after flow interms of page numbering
            'For example, an 'es' will normally restart at i. We don't want the bounded section ad the following 
            'section to also restart. So we take steps.
            '
            myDoc = objGlobals.glb_get_wrdActiveDoc()
            '
            'Make certain that page numbering is OK
            'Do the bounded section
            objChptBase.chptBase_PageNumbering_Set(sect, False, 1, "flow")
            '
            'Now do the section behind it
            'sectNew = myDoc.Sections.Item(sect.Index + 1)
            objChptBase.chptBase_PageNumbering_Set(myDoc.Sections.Item(sect.Index + 1), False, 1, "flow")
            '
        Catch ex As Exception

        End Try
        '
finis:
        Return sectMid
        '
    End Function
    '
    ''' <summary>
    ''' This method will set the Section Start to NewPage, OddPage or EvenPage depending
    ''' on the value of strBreakStype; 'newPage', 'oddPage' or 'evenPage
    ''' </summary>
    ''' <param name="strBreakType"></param>
    ''' <param name="sect"></param>
    Public Sub sect_set_SectionStart(ByRef sect As Word.Section, strBreakType As String)

        Select Case strBreakType
            Case "newPage"
                sect.PageSetup.SectionStart = WdSectionStart.wdSectionNewPage
            Case "oddPage"
                sect.PageSetup.SectionStart = WdSectionStart.wdSectionOddPage
            Case "evenPage"
                sect.PageSetup.SectionStart = WdSectionStart.wdSectionEvenPage
        End Select

    End Sub
    '
    ''' <summary>
    ''' Will return the section type as a "binary string" according to
    ''' "0 mirror margins, 0 different odd and even, 0 different first page"
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function sct_get_SectionType(ByRef sect As Word.Section) As String
        Dim strSectionType As String
        '
        strSectionType = ""
        '
        If Not sect.PageSetup.MirrorMargins And Not sect.PageSetup.OddAndEvenPagesHeaderFooter And Not sect.PageSetup.DifferentFirstPageHeaderFooter Then strSectionType = "000"
        If Not sect.PageSetup.MirrorMargins And Not sect.PageSetup.OddAndEvenPagesHeaderFooter And sect.PageSetup.DifferentFirstPageHeaderFooter Then strSectionType = "001"
        If Not sect.PageSetup.MirrorMargins And sect.PageSetup.OddAndEvenPagesHeaderFooter And Not sect.PageSetup.DifferentFirstPageHeaderFooter Then strSectionType = "010"
        If Not sect.PageSetup.MirrorMargins And sect.PageSetup.OddAndEvenPagesHeaderFooter And sect.PageSetup.DifferentFirstPageHeaderFooter Then strSectionType = "011"

        If sect.PageSetup.MirrorMargins And Not sect.PageSetup.OddAndEvenPagesHeaderFooter And Not sect.PageSetup.DifferentFirstPageHeaderFooter Then strSectionType = "100"
        If sect.PageSetup.MirrorMargins And Not sect.PageSetup.OddAndEvenPagesHeaderFooter And sect.PageSetup.DifferentFirstPageHeaderFooter Then strSectionType = "101"
        If sect.PageSetup.MirrorMargins And sect.PageSetup.OddAndEvenPagesHeaderFooter And Not sect.PageSetup.DifferentFirstPageHeaderFooter Then strSectionType = "110"
        If sect.PageSetup.MirrorMargins And sect.PageSetup.OddAndEvenPagesHeaderFooter And sect.PageSetup.DifferentFirstPageHeaderFooter Then strSectionType = "111"
        '
        Return strSectionType
        '
    End Function
    '
    Public Sub cloneSection(ByRef srcSection As Section, ByRef destSection As Section, doHeadersFooters As Boolean)
        'This method will reproduce the settings of the source Section (srcSection) in the
        'destination Section (destSection
        '
        Dim strOrientation As String
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim hf As Word.HeaderFooter
        '
        strOrientation = "portrait"
        If srcSection.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then strOrientation = "landscape"
        '
        destSection.PageSetup.PaperSize = srcSection.PageSetup.PaperSize
        destSection.PageSetup.Orientation = srcSection.PageSetup.Orientation
        destSection.PageSetup.GutterPos = srcSection.PageSetup.GutterPos
        destSection.PageSetup.DifferentFirstPageHeaderFooter = srcSection.PageSetup.DifferentFirstPageHeaderFooter
        destSection.PageSetup.OddAndEvenPagesHeaderFooter = srcSection.PageSetup.OddAndEvenPagesHeaderFooter
        destSection.PageSetup.MirrorMargins = srcSection.PageSetup.MirrorMargins
        '
        'Copy page dimensions
        destSection.PageSetup.TopMargin = srcSection.PageSetup.TopMargin
        destSection.PageSetup.LeftMargin = srcSection.PageSetup.LeftMargin
        destSection.PageSetup.BottomMargin = srcSection.PageSetup.BottomMargin
        destSection.PageSetup.RightMargin = srcSection.PageSetup.RightMargin
        destSection.PageSetup.Gutter = srcSection.PageSetup.Gutter
        destSection.PageSetup.HeaderDistance = srcSection.PageSetup.HeaderDistance
        destSection.PageSetup.FooterDistance = srcSection.PageSetup.FooterDistance
        '
        If doHeadersFooters Then
            objHfMgr.hf_hfs_linkUnlinkAll(destSection, False)
            objHfMgr.hf_headers_delete(destSection)

            For Each hf In srcSection.Headers
                objHfMgr.hf_hfs_CopyHeader(hf, srcSection, destSection)
            Next
            '
            objHfMgr.hf_footers_delete(destSection)
            '
            For Each hf In srcSection.Footers
                objHfMgr.hf_hfs_CopyFooter(hf, srcSection, destSection)
            Next
            '
        End If


        'Call objHfMgr.hf_hfs_CopyHeaderFooter("header", srcSection, destSection)
        'Call objHfMgr.hf_hfs_CopyHeaderFooter("footer", srcSection, destSection)
        '

    End Sub
    '
    ''' <summary>
    ''' This method will look for the section tag. That is, the style in the section's
    ''' header row (first cell). If no header row, then the method looks for the tag style
    ''' in the primary or first page header... If it doesn't find a tag style it will reurn null.
    ''' Tag styles definitions can be found in cChptBanner.bnr_get_tagStyle
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function sct_Get_SectionTag(ByRef sect As Section) As String
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim strTag As String
        '
        '*** 20241228 Change to header based tags
        '
        strTag = objHfMgr.hf_tags_getTagStyleName(sect, "primaryOrFirstPage")
        '
        'strTag = ""
        '
        'Try
        'rng = sect.Range
        'If rng.Tables.Count > 0 Then
        'We may be in a tagged section.. Look for th tag
        'tbl = rng.Tables.Item(1)
        'drCell = tbl.Range.Cells(1)
        'tagStyle = drCell.Range.Style
        'strTag = tagStyle.NameLocal
        'End If
        'Catch ex As Exception
        'strTag = ""
        'End Try
        '
        Return strTag
        '
    End Function
    '
    ''' <summary>
    ''' This method will return true if the sepecified section, sect has the sectionType tag strSectionTag
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="strSectionTypeTag"></param>
    ''' <returns></returns>
    Public Function sct_sectHas_strTag(ByRef sect As Word.Section, strSectionTypeTag As String) As Boolean
        Dim rslt As Boolean
        '
        rslt = False
        If Me.sct_Get_SectionTag(sect) = strSectionTypeTag Then rslt = True
        '
        Return rslt
        '
    End Function
    '
    '
    Public Function sct_has_Text(ByRef sect As Word.Section) As Boolean
        'This method will determine if the section where the current
        'cursor position is located has any text
        '
        Dim rng As Range
        Dim para As Paragraph
        '
        rng = sect.Range
        '
        sct_has_Text = False
        '
        'Remember that paragraph char must be included in
        'any consideration
        For Each para In sect.Range.Paragraphs
            If Len(para.Range.Text) > 1 Then
                sct_has_Text = True
                Exit For
            End If
        Next para
        '
    End Function
    '
    ''' <summary>
    ''' This method will return true if a section is found with a sepecific tagStyle.. If it does return true, then
    ''' sect is set to that section, otherwise sect is set to nothing
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="strSectionTag"></param>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function sct_has_strTag(ByRef myDoc As Word.Document, strSectionTag As String, ByRef sect As Word.Section) As Boolean
        Dim rslt As Boolean
        'Dim sect As Word.Section
        Dim strTempTag As String
        '
        rslt = False
        '
        For Each sect In myDoc.Sections
            strTempTag = Me.sct_Get_SectionTag(sect)
            If strTempTag = strSectionTag Then
                rslt = True
                Exit For
            End If
        Next
        '
        If rslt = False Then sect = Nothing

        Return rslt
        '
    End Function

    '
    '
    ''' <summary>
    ''' This method will delet all paragraphs in the specified section. It will put back
    ''' an empty paragraph and leave the selection there. The returned Range is the range
    ''' of the selection 
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function sct_Paragraphs_DeleteAll(ByRef sect As Word.Section) As Word.Range
        Dim rng As Word.Range
        '
        rng = sect.Range()
        If sect.Index = Globals.ThisAddIn.Application.ActiveDocument.Sections.Last.Index Then
            rng.Delete()
            rng.Select()
        Else
            'A section other than the last section
            rng.MoveEnd(WdUnits.wdParagraph, -2)
            rng.Delete()
            rng.Paragraphs.Add(rng)
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Select()
        End If
        'rng.Delete(WdUnits.wdParagraph, rng.Paragraphs.Count)
        '
        Return rng
        '
    End Function
    '
#Region "INsert sections at start or end"
    '
    'This method will insert a section at the beginning of the of the document.
    'It will take on default headerFooters and page characteristics. I have left
    'the ability to format landscape and portrait pages with the strSectionType
    'variable.. But it is unused at present. On exit, the selection will be at
    'the beginning of the section/document and the return value will be the first
    'section in the document
    ''' <summary>
    ''' This method will insert a section at the beginning of the of the document.
    ''' </summary>
    ''' <param name="lstOfMarginDimensions"></param>
    ''' <returns></returns>
    Public Function sct_insert_SectionAtStart(Optional diffFirstPage As Boolean = True, Optional ByRef lstOfMarginDimensions As Collection = Nothing) As Section
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim sect As Section
        Dim placeBehind As Boolean
        Dim myDoc As Word.Document
        '
        myDoc = objGlobals.glb_get_wrdActiveDoc()
        sect = myDoc.Sections.First
        placeBehind = False
        '
        'If IsNothing(lstOfMarginDimensions) Then lstOfMarginDimensions = Me.objGlobals.glb_getDimensions_Letter()
        '
        'rng = sect.Range
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'sect = myDoc.Sections.Add(rng, WdSectionStart.wdSectionNewPage)
        '
        'sect = myDoc.Sections.Item(sect.Index - 1)

        'objHfMgr.hf_hfs_linkUnlinkAll(sect, False)
        'sect = myDoc.Sections.First




        sect = Me.sct_insert_Section(False, sect, 6, "newPage", False, "prt", lstOfMarginDimensions)
        objHfMgr.hf_hfs_deleteAll(sect)
        sect.PageSetup.DifferentFirstPageHeaderFooter = diffFirstPage
        '
        Return sect
        '
    End Function
    '
    ''' <summary>
    ''' This method will ensure that there is an empty last section at the end of the document
    ''' </summary>
    ''' <param name="lstOfDimensions"></param>
    ''' <returns></returns>
    Public Function sct_insert_SectionAtEnd(ByRef lstOfDimensions As Collection) As Section
        Dim objParas As New cParas()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim myDoc As Word.Document
        Dim sect As Section
        Dim strOrientation As String
        Dim rng As Word.Range
        '
        myDoc = Me.objGlobals.glb_get_wrdActiveDoc()
        sect = myDoc.Sections.Last
        '
        strOrientation = "prt"
        If Me.objGlobals.glb_get_wrdSect().PageSetup.Orientation = WdOrientation.wdOrientLandscape Then strOrientation = "lnd"
        '
        'strRptMode = objRptMgr.Rpt_Mode_Get()
        '
        rng = myDoc.Sections.Last.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng = objParas.paras_insert_numParas(rng, 4)
        rng.Move(WdUnits.wdParagraph, 2)
        '
        sect = myDoc.Sections.Add(rng, WdSectionStart.wdSectionNewPage)
        sect = myDoc.Sections.Last
        objHFMgr.hf_hfs_linkUnlinkAll(sect, False)
        objHFMgr.hf_hfs_deleteAll(sect)
        sect.PageSetup.DifferentFirstPageHeaderFooter = True
        '
        'If list of dimensions are not specified then we set the defaults
        '
        If lstOfDimensions.Count = 0 Then
            Select Case strOrientation
                Case "prt"
                    sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait
                    lstOfDimensions = objGlobals.glb_getDimensions_Contacts_Prt("back")
                    'lstOfDimensions = objGlobals.glb_getDimensions_Std_Prt()
                Case "lnd"
                    sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape
                    lstOfDimensions = objGlobals.glb_getDimensions_Contacts_Lnd()
                Case Else
                    sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait
                    lstOfDimensions = objGlobals.glb_getDimensions_Contacts_Prt("back")
                    'lstOfDimensions = objGlobals.glb_getDimensions_Std_Prt()
            End Select
            '
        End If
        '
        objGlobals.glb_setDimensions(sect, lstOfDimensions)
        objHFMgr.hf_headers_insert(sect)
        objHFMgr.hf_footers_insert(sect)
        '
        'tbl = Nothing
        '
        'objHFMgr.hf_headers_insert(sect, -1, False)
        'objHFMgr.hf_footers_delete(sect)

        '
        Return sect

    End Function


    ''' <summary>
    ''' This method will ensure that there is an empty last section at the end of the document
    ''' </summary>
    ''' <param name="lstOfDimensions"></param>
    ''' <returns></returns>
    Public Function sct_insert_SectionAtEndx(ByRef lstOfDimensions As Collection) As Section
        Dim sect As Section
        Dim rng As Range
        Dim para As Paragraph
        Dim oldRng As Word.Range
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objRptMgr As New cReport()
        '
        oldRng = Me.objGlobals.glb_get_wrdSel.Range
        '
        sect = objGlobals.glb_get_wrdActiveDoc.Sections.Last
        rng = sect.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
        objGlobals.glb_setDimensions(sect, lstOfDimensions)
        GoTo finis
        '
        If Me.sct_has_Text(sect) Then
            para = rng.Paragraphs.Add(rng)
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            '
            sect = Me.sct_insert_Section(True, sect, 6,,, "flow", lstOfDimensions)
            '
            'If strRptMode = objRptMgr.modeLongLandscape Then
            'sect = Me.sct_insert_Section(True, sect, 6,,, "flow", lstOfDimensions)
            'sect = Me.sct_insert_Section(True, sect, 6,,, "lnd", objGlobals.glb_getDimensions_ContactsPage("back"))
            'Else
            'sect = Me.sct_insert_Section(True, sect, 6,,,, objGlobals.glb_getDimensions_ContactsPage("back"))
            'End If
            '
            'sect = Me.sct_insert_SectionAtSelection(rng, 25)
        Else
            objGlobals.glb_setDimensions(sect, lstOfDimensions)
            'sect = Me.sct_insert_Section(True, sect, 6,,, "flow", lstOfDimensions)

            '
            'objHFMgr.hf_headers_insert(sect)
            'objHFMgr.hf_footers_insert(sect)
            'objHFMgr.hf_hfs_deleteAll(sect)
        End If
        '
finis:
        sect = objGlobals.glb_get_wrdActiveDoc.Sections.Last
        '
        Return sect

    End Function
    '
#End Region
    '

    '
    'This method will return true if the current selection is in a table or just under the table
    Public Function sct_Sel_IsIn_Or_JustUnderTable() As Boolean
        Dim rng As Range
        Dim rslt As Boolean
        '
        rslt = False
        rng = Me.objGlobals.glb_get_wrdSel.Range
        If rng.Tables.Count <> 0 Then rslt = True
        '
        Return rslt
    End Function
    '

    ''' <summary>
    ''' This method will return true if the selection is inside a table.. If it is just under it
    ''' will return false
    ''' </summary>
    ''' <returns></returns>
    Public Function sct_Sel_IsIn_TableOnly() As Boolean
        Dim rslt As Boolean
        '
        rslt = objGlobals.glb_get_wrdApp.Selection.Information(WdInformation.wdWithInTable)
        '
        Return rslt
    End Function
    '
    ''' <summary>
    ''' This method will return true if the Selection contains the first paragraph of the section
    ''' </summary>
    ''' <returns></returns>
    Public Function sct_Sel_Is_FirstParaInSection() As Boolean
        Dim rng As Range
        Dim rslt As Boolean
        Dim sect As Word.Section
        Dim paraFirst, paraSel As Word.Paragraph
        Dim paraFirstID, paraSelID As Integer
        Dim objParas As New cParas()
        '
        rng = objGlobals.glb_get_wrdSelRng
        sect = rng.Sections.Item(1)
        '
        paraFirst = sect.Range.Paragraphs.Item(1)
        paraSel = rng.Paragraphs.Item(1)
        '
        paraSelID = paraSel.ParaID
        paraFirstID = paraFirst.ParaID
        '
        If paraSelID = paraFirstID Then
            rslt = True
        End If
        '
        '
        Return rslt
    End Function
    '
    ''' <summary>
    ''' If the selection is in the first paragraph of a section, then this method will
    ''' add two additional (empty) paragraphs to the top of the section and set the selection
    ''' to the beginning of the second paragraph. The style for these added paras is 'Body text'.
    ''' The paragraph containing the original selection is left unchanged.
    ''' </summary>
    Public Function sct_set_SelforTableInsert() As Word.Range
        Dim objParas As New cParas()
        Dim sect As Word.Section
        Dim rng As Word.Range
        Dim paraFirst, paraSecond As Word.Paragraph
        '
        rng = objGlobals.glb_get_wrdSelRng
        sect = rng.Sections.Item(1)
        '
        Try
            If Me.sct_Sel_Is_FirstParaInSection() Then
                paraFirst = sect.Range.Paragraphs.Item(1)
                rng = paraFirst.Range
                '
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng = objParas.paras_insert_numParas(rng, 2)
                paraSecond = sect.Range.Paragraphs.Item(2)
                rng = paraSecond.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                '
            End If
        Catch ex As Exception

        End Try
        '
        Return rng
        '
    End Function
    '

End Class

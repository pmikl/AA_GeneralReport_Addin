Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cChptBase
    Inherits cSectionMgr
    '
    Public strTagStyleName As String
    Public objTools As cTools
    Public tbl_Banner As Word.Table                 'Thi is the banner table
    Public rng_Heading As Word.Range                'This is the range of the Heading, when we turn off the banners
    '
    Public Sub New()
        MyBase.New()
        Me.strTagStyleName = ""
        Me.objTools = New cTools()
        Me.tbl_Banner = Nothing
    End Sub
    '
    '
    ''' <summary>
    ''' This method will delete all of the small image place holders, (including the open 
    ''' patterns) on the Cover Page.. It will return true if all is OK and false if it is not
    ''' </summary>
    Public Function ChptBase_delete_SmallPicturePlaceHolders(ByRef sect As Word.Section) As Boolean
        Dim objShpMgr As cShapeMgr
        Dim lstOfBackPanels As New List(Of cShapeMgr)
        Dim objPanelMgr As New cBackPanelMgr()
        Dim rslt As Boolean
        Dim lstOfPictNames As New List(Of String)
        Dim str As String
        '
        rslt = False
        '
        lstOfPictNames.Add("cp_pict_large")
        lstOfPictNames.Add("cp_pict_purplePattern_prt")
        lstOfPictNames.Add("cp_pict_purplePattern_lnd")
        lstOfPictNames.Add("cp_pict_seaSide_prt")
        lstOfPictNames.Add("cp_pict_seaSide_lnd")
        lstOfPictNames.Add("cp_Empty_Pattern_Small")
        lstOfPictNames.Add("cp_Pict_EmptyPattern")
        lstOfPictNames.Add("aac_jigsaw_Wide")
        lstOfPictNames.Add("aac_Cpg_PictEmptyPatternSmall")
        lstOfPictNames.Add("aac_Cpg_PictEmptyPattern")
        '
        Try
            For Each str In lstOfPictNames
                lstOfBackPanels.Clear()
                lstOfBackPanels = objPanelMgr.pnl_getBackPanel_PlaceHolders(sect, str)                     'To get rid of any existing back panels
                If lstOfBackPanels.Count > 0 Then
                    objShpMgr = lstOfBackPanels.Item(0)
                    objShpMgr.shp.Delete()
                    rslt = True
                End If
            Next
            '
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
    End Function
    '
    ''' <summary>
    ''' This method will place a standard section either behind (placeBehind=true) or in front (placeBehind=false), the section sect The banners will
    ''' be set as per "lstOfBannerSettings". The vriable strRptMode indicates standard, short or landscape reports, but it is unused
    ''' at the moment. The variable 'strOrientation' is used in "MyBase.sct_insert_Section" to select the page orientation and dimensions.
    ''' and "numParas" which defaults to 6 is the number of empty paragraphs placed in the chapter.
    ''' 
    ''' This method will return the banner "Table" and sect will contin the new section NOT the old section
    ''' </summary>
    ''' <param name="placeBehind"></param>
    ''' <param name="sect"></param>
    ''' <param name="lstOfBannerSettings"></param>
    ''' <param name="strRptMode"></param>
    ''' <param name="strOrientation"></param>
    ''' <param name="numParas"></param>
    ''' <returns></returns>
    Public Overridable Function chpt_Insert_Std(placeBehind As Boolean, ByRef sect As Word.Section, ByRef lstOfBannerSettings As Collection, strRptMode As String, Optional strOrientation As String = "prt", Optional numParas As Integer = 6, Optional doBannerImage As Boolean = True) As Word.Table
        Dim objFldsMgr As New cFieldsMgr()
        Dim objBnrMgr As New cChptBanner()
        Dim tbl As Word.Table
        Dim rng As Word.Range
        'Dim doBannerImage As Boolean
        Dim myDoc As Word.Document
        Dim strDoImage As String
        '
        tbl = Nothing
        myDoc = sect.Range.Document
        '
        strDoImage = CStr(lstOfBannerSettings("strDoImage"))
        'doBannerImage = True
        If strDoImage = "False" Then doBannerImage = False
        '
        'Insert a section, then insert a banner
        sect = MyBase.sct_insert_Section(placeBehind, sect, numParas, "newPage", False, strOrientation,)
        '
        '***
        'If placeBehind Then
        'sect.Range.Select()
        'MsgBox("Placebehind")
        'End If
        '
        'tbl = objBnrMgr.bnr_insert_BannerBase(sect.Range, doBannerImage, strRptMode, lstOfBannerSettings)
        Me.tbl_Banner = objBnrMgr.bnr_insert_BannerBase(sect.Range, doBannerImage, strRptMode, lstOfBannerSettings)

        '
        rng = sect.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()
        '
loop1:
        Return Me.tbl_Banner
    End Function
    '
    ''' <summary>
    ''' This method will set the style in the first cell of the Header Table to a style with
    ''' the name strTagStyleName.. It currently searches the primary header only
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="strTagStyleName"></param>
    Public Sub chptBase_set_tagStyleInHeaderTable(ByRef sect As Word.Section, strTagStyleName As String)
        Dim objHfMgr As New cHeaderFooterMgr()
        '
        objHfMgr.hf_tags_setTagStyle(sect, strTagStyleName)
        '
        'hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        'rng = hf.Range
        '
        'If rng.Tables.Count <> 0 Then
        'tbl = rng.Tables.Item(1)
        'rng = tbl.Range.Cells.Item(1).Range
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'rng.Style = rng.Document.Styles.Item(strTagStyleName)
        'End If
        '

    End Sub
    '
    ''' <summary>
    ''' This function will return the range of the text in the standard banners.
    ''' That is, in the 3rd cell
    ''' </summary>
    ''' <returns></returns>
    Public Function chptBase_getRange_Heading1(ByRef sect As Word.Section) As Word.Range
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        '
        rng = Nothing
        Try
            If Not IsNothing(Me.tbl_Banner) Then
                drCell = Me.tbl_Banner.Range.Cells.Item(3)
                rng = drCell.Range
                rng.MoveEnd(WdUnits.wdCharacter, -1)
            Else
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng = rng.Paragraphs.Item(1).Range
                rng.MoveEnd(WdUnits.wdCharacter, -1)
            End If
        Catch ex As Exception
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng = rng.Paragraphs.Item(1).Range
        End Try
        '
        Return rng
    End Function
    '
    '
    ''' <summary>
    ''' This function will return the range of the first paragraph in the input range.
    ''' Typically this is 'Heading 1'
    ''' </summary>
    ''' <returns></returns>
    Public Function chptBase_getRange_Heading1(ByRef rng As Word.Range) As Word.Range
        '
        Try
            rng = rng.Paragraphs.Item(1).Range
            rng.MoveEnd(WdUnits.wdCharacter, -1)
            '
        Catch ex As Exception

        End Try
        '
        Return rng
    End Function

    '
    Public Function chptBase_get_TitleRangeCoverPage(ByRef sect As Word.Section) As Word.Range
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        '
        rng = Nothing
        Try
            drCell = sect.Range.Tables.Item(1).Range.Cells(4)
            rng = drCell.Range
            rng.MoveEnd(WdUnits.wdCharacter, -1)
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'rng.Select()
        Catch ex As Exception

        End Try
        '
        Return rng
    End Function
    '
    '
    Public Overridable Function chpt_Insert_Short(ByRef lstOfSettings As Collection, strRptMode As String) As Word.Table
        Dim objFldsMgr As New cFieldsMgr()
        Dim objBnrMgr As New cChptBanner()
        Dim tbl As Word.Table
        Dim rng As Word.Range
        Dim doBannerImage As Boolean
        '
        '
        doBannerImage = True
        '
        '*** Must insert Banner Base at Selection
        rng = MyBase.objGlobals.glb_get_wrdSelRng()
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        'tbl = Me.chptBase_insert_BannerBase(rng, doBannerImage, strRptMode)
        tbl = objBnrMgr.bnr_insert_BannerBase(rng, doBannerImage, strRptMode, lstOfSettings)

        'objFldsMgr.updateSequenceNumbers_Chapters()
        '
        Return tbl
        '
    End Function
    '
    ''' <summary>
    ''' This method will insert two parts of the standard landscape chapter (i.e. the banner page containing the returned Table, and
    ''' the follower page as identified by the variable sect). The banner page will be inserted behind/inFornt of the section sect. Note
    ''' that sect (on exit) will then point to the follower section. Whilst the returned Table is on the banner page
    ''' </summary>
    ''' <param name="placeBehind"></param>
    ''' <param name="sect"></param>
    ''' <param name="lstOfSettings"></param>
    ''' <returns></returns>
    Public Overridable Function chpt_Insert_LandscapeReport(placeBehind As Boolean, ByRef sect As Word.Section, ByRef lstOfSettings As Collection) As Word.Table
        Dim tbl As Word.Table
        'Dim sectNew As Word.Section
        Dim rng As Word.Range
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objBrndMgr As New cBrandMgr()
        Dim objColsHndlr As New cColsHandler()
        '
        '
        tbl = Me.chpt_Insert_Std(placeBehind, sect, lstOfSettings, "", "lndRptChpt", 3, False)
        '
        'Now do follower page
        rng = sect.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        sect = MyBase.sct_insert_Section(True, sect, 6, "newPage", False, "lndRptFollower",)
        '
        objColsHndlr.cols_setup_columnStructure(sect, "3_columns")
        'objWtrMrks.waterMarks_Remove_VersionMark(sectNew)
        '
        'sect = tbl.Range.Sections.Item(1)
        objBrndMgr.brnd_Rebuild_Background(tbl.Range.Sections.Item(1), False, False)
        objBrndMgr.brnd_LndScp_ChptBanner_jigsaw(tbl.Range.Sections.Item(1))
        '
        'objHFMgr.hf_footers_delete(tbl.Range.Sections.Item(1))
        'objHFMgr.hf_footers_delete(sect)
        '
        'Need to do this to work around the fact that the Appendix 2 part seems to stick
        'Me.chptBase_PageNumbering_Set(tbl.Range.Sections.Item(1), False, 1, "2part")
        'Me.chptBase_PageNumbering_Set(sect, False, 1, "2part")
        '
        'Ensure that the base numbering format is 'std' (i.e. 1 part)... Need to do this to 
        'get around the sticking Appenidx page number format
        Me.chptBase_PageNumbering_Set(tbl.Range.Sections.Item(1), False, 1, "std")
        Me.chptBase_PageNumbering_Set(sect, False, 1, "std")

        'objHFMgr.hf_footers_insert(tbl.Range.Sections.Item(1))
        'objHFMgr.hf_footers_insert(sect)

        'objHFMgr.hf_footers_insert(sect)
        '
        Return tbl
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will setup the column layout for the specified section. It does so in
    ''' accordance with strMode ("3_columns", "2_columns_leftWide", "2_columns_rightWide"
    ''' and "1_column"
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="strLayout"></param>
    Public Sub chptBase_Columns_Setup(ByRef sect As Word.Section, strLayout As String)
        Dim pageWidth, drColWidth, colSpacing As Single
        'Dim objRpt As New cChapterReport()
        Dim rng As Word.Range
        '
        pageWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
        colSpacing = Me.objTools.tools_math_MillimetersToPoints(10.0)
        'colSpacing = Me.ChptBase_Math_MillimeterToPoints(10.0)                                              '10 mm spacing
        'colSpacing = Me.ChptBase_Math_MillimeterToPoints(5.0)
        '
        Select Case strLayout
            Case "4_columns"
                drColWidth = (pageWidth - 3 * colSpacing) / 4.0
                'Add three columns to the existing column.. We'll specify the width and
                'specifiy even spacing
                sect.PageSetup.TextColumns.SetCount(3)
                sect.PageSetup.TextColumns.LineBetween = False
                sect.PageSetup.TextColumns.Add(drColWidth, , True)
                '
                'The above funtion chnages the prior section break to "Continuous" (for some unknown reason)
                'So we need to set it back to next page.. Place the cursor just after the Section break, then set
                'then section break to "Next Page"
                '
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Select()
                Globals.ThisAddIn.Application.Selection.PageSetup.SectionStart = WdSectionStart.wdSectionNewPage

            Case "3_columns"
                drColWidth = (pageWidth - 2 * colSpacing) / 3.0
                'Add two columns to the existing column.. We'll specify the width and
                'specifiy even spacing
                sect.PageSetup.TextColumns.SetCount(2)
                sect.PageSetup.TextColumns.LineBetween = False
                sect.PageSetup.TextColumns.Add(drColWidth, , True)
                '
                'The above funtion chnages the prior section break to "Continuous" (for some unknown reason)
                'So we need to set it back to next page.. Place the cursor just after the Section break, then set
                'then section break to "Next Page"
                '
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Select()
                Globals.ThisAddIn.Application.Selection.PageSetup.SectionStart = WdSectionStart.wdSectionNewPage
                '
            Case "2_columns"
                drColWidth = (pageWidth - colSpacing) / 2.0
                '
                'We add an addiitonal column
                sect.PageSetup.TextColumns.SetCount(1)
                sect.PageSetup.TextColumns.LineBetween = False
                sect.PageSetup.TextColumns.Add(drColWidth, , False)
                '
                'The above funtion chnages the prior section break to "Continuous" (for some unknown reason)
                'So we need to set it back to next page.. Place the cursor just after the Section break, then set
                'then section break to "Next Page"
                '
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Select()
                Globals.ThisAddIn.Application.Selection.PageSetup.SectionStart = WdSectionStart.wdSectionNewPage
            Case "2_columns_leftWide"
                drColWidth = (pageWidth - colSpacing) / 3.0
                '
                'We add an addiitonal column
                sect.PageSetup.TextColumns.SetCount(1)
                sect.PageSetup.TextColumns.LineBetween = False
                sect.PageSetup.TextColumns.Add(drColWidth, colSpacing, False)
                '
                'The above funtion chnages the prior section break to "Continuous" (for some unknown reason)
                'So we need to set it back to next page.. Place the cursor just after the Section break, then set
                'then section break to "Next Page"
                '
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Select()
                Globals.ThisAddIn.Application.Selection.PageSetup.SectionStart = WdSectionStart.wdSectionNewPage
                '
            Case "2_columns_rightWide"
                drColWidth = (pageWidth - colSpacing) / 3.0
                '
                sect.PageSetup.TextColumns.SetCount(1)
                sect.PageSetup.TextColumns.LineBetween = False
                sect.PageSetup.TextColumns.Add(2.0 * drColWidth, colSpacing, False)
                '
                'The above funtion chnages the prior section break to "Continuous" (for some unknown reason)
                'So we need to set it back to next page.. Place the cursor just after the Section break, then set
                'then section break to "Next Page"
                '
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Select()
                Globals.ThisAddIn.Application.Selection.PageSetup.SectionStart = WdSectionStart.wdSectionNewPage

            Case "1_column"
                sect.PageSetup.TextColumns.SetCount(1)

        End Select
    End Sub


    '
    ''' <summary>
    ''' If you pass the banner table (tbl) to this method, it will
    ''' select the first paragraph below the table
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub chptBase_select_Chapter(ByRef tbl As Word.Table)
        Dim rng As Word.Range
        '
        Try
            If Not IsNothing(tbl.Range) Then
                rng = tbl.Range
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                '
                rng.Select()

            End If

        Catch ex As Exception

        End Try
        '
    End Sub

    ''' <summary>
    ''' If you pass a section (sect) to this method it will look for a table
    ''' at the beginnign of the section. If one exists, it is assumed to be
    ''' a banner table. The current selection is chnaged to be at the first
    ''' paragraph below the table.. If there is not table, the section is shifted
    ''' to the beginning of the section
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub chptBase_select_Chapter(ByRef sect As Word.Section)
        Dim rng As Word.Range
        Dim tbl As Word.Table
        '
        rng = sect.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        If rng.Tables.Count <> 0 Then
            tbl = rng.Tables.Item(1)
            rng = tbl.Range
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        End If
        '
        rng.Select()
        '
    End Sub
    '

#Region "Reset Size"
    '
    Public Function chptBase_Toggle_Width(ByRef sect As Section) As String
        Dim objSectMgr As New cSectionMgr()
        Dim objRptMgr As New cReport()
        Dim objChpExec As New cChptExec()
        Dim objChpApp As New cChptApp()
        Dim objCpMgr As New cCoverPageMgr()
        Dim objTOCMgr As New cChptTOC()
        Dim objTagsMgr As New cTagsMgr()
        Dim objChptBanner As New cChptBanner()
        Dim strMsg As String
        Dim strResult, strTagSection As String
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim offSet As Single                                                        'This is the number of points the left margin is offset from the left edge of the header table
        '
        strResult = ""
        strMsg = "The 'Toggle Width' function can only be successfully applied" + vbCr + "to the 'ExecSummary, Chapters or Appendix Chapters"
        '
        offSet = objSectMgr.objGlobals.glb_get_TableOutdent()                          'Make the left margin offset equal to the outdent of the standard AA Table
        '
        strTagSection = objTagsMgr.tags_get_tagStyleName(sect)
        'strTagSection = MyBase.ChptBase_TagStyle_GetForSection(sect)
        '
        'If mode is Landscape and page is not coverPage, not TOC, not Contacts, not banner page then we can proceed
        '
        Select Case objRptMgr.Rpt_Mode_Get()
            Case objRptMgr.rpt_isPrt, objRptMgr.modeShort
                Select Case strTagSection
                    Case Me.strTagStyleName, objChpExec.strTagStyleName, objChpApp.strTagStyleName, ""
                        If objCpMgr.cp_Bool_IsCoverPage(sect) Or objTOCMgr.is_TOCPage(sect) Then
                            MsgBox(strMsg)
                            GoTo finis
                        End If
                        '
                        strResult = objSectMgr.sect_Toggle_Width(sect, offSet)      'Left margin offset from the Header table is normally enough to accomodate the standard AA Table outdent
                        '
                        'Now autofit the Banner
                        rng = sect.Range
                        rng.Collapse(WdCollapseDirection.wdCollapseStart)
                        If rng.Tables.Count <> 0 Then
                            tbl = rng.Tables.Item(1)
                            If objChptBanner.bnr_is_Chapter_Bdy_or_ES_or_AP(tbl) Then
                                Me.objGlobals.glb_tbls_AutoFitBanner(tbl, False)
                                'Me.chptBase_Banner_Autofit(tbl, sect, False)
                            End If
                        End If
                        '
                    Case Else
                        MsgBox(strMsg)
                        '
                End Select
                '
            Case objRptMgr.rpt_isLnd
                If Not (strTagSection = "") Then
                    MsgBox("In Landscape reports you can only toggle" + vbCr + "the width of 'non banner' pages")
                    GoTo finis
                Else
                    If objCpMgr.cp_Bool_IsCoverPage(sect) Or objTOCMgr.is_TOCPage(sect) Then
                        MsgBox(strMsg)
                        GoTo finis
                    End If
                    strResult = objSectMgr.sect_Toggle_Width(sect, offSet)
                End If
        End Select
        '
finis:
        Return strResult
        '
    End Function
#End Region
    '
    '
#Region "PageNumbering"
    '

    Public Overridable Function chptBase_PageNumbering_Set(ByRef sect As Word.Section, restartNumbering As Boolean, Optional startAt As Integer = 1, Optional pgNumStyle As String = "std") As Boolean
        Dim pgNums As PageNumbers
        Dim objToolsMgr As New cTools()
        Dim hf As Word.HeaderFooter
        Dim rslt As Boolean
        Dim str As String
        '
        rslt = True
        str = ""
        '
        Try
            If sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Exists Then str = "firstPage"
            If sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary).Exists Then str = str + "+Primary"
            If sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterEvenPages).Exists Then str = str + "+Even"
            '
            'MsgBox(str)
            '
            Select Case str
                Case "firstPage+Primary"
                    hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                    pgNums = hf.PageNumbers
                    Me.chptBase_PageNumbering_SetFormat(pgNumStyle, pgNums)
                    pgNums.RestartNumberingAtSection = restartNumbering
                    pgNums.StartingNumber = startAt
                    '
                    hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                    pgNums = hf.PageNumbers
                    Me.chptBase_PageNumbering_SetFormat(pgNumStyle, pgNums)
                    pgNums.RestartNumberingAtSection = False
                    pgNums.StartingNumber = startAt

                Case "+Primary"
                    hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                    pgNums = hf.PageNumbers
                    Me.chptBase_PageNumbering_SetFormat(pgNumStyle, pgNums)
                    pgNums.RestartNumberingAtSection = restartNumbering
                    pgNums.StartingNumber = startAt
                Case "+Primary+Even"
                    hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                    pgNums = hf.PageNumbers
                    Me.chptBase_PageNumbering_SetFormat(pgNumStyle, pgNums)
                    pgNums.RestartNumberingAtSection = restartNumbering
                    pgNums.StartingNumber = startAt
                    '
                    hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterEvenPages)
                    pgNums = hf.PageNumbers
                    Me.chptBase_PageNumbering_SetFormat(pgNumStyle, pgNums)
                    pgNums.RestartNumberingAtSection = False
                    pgNums.StartingNumber = startAt
                '
                Case "firstPage+Primary+Even"
                    hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                    pgNums = hf.PageNumbers
                    Me.chptBase_PageNumbering_SetFormat(pgNumStyle, pgNums)
                    pgNums.RestartNumberingAtSection = restartNumbering
                    pgNums.StartingNumber = startAt
                    '
                    hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                    pgNums = hf.PageNumbers
                    Me.chptBase_PageNumbering_SetFormat(pgNumStyle, pgNums)
                    pgNums.RestartNumberingAtSection = False
                    pgNums.StartingNumber = startAt
                    '
                    hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterEvenPages)
                    pgNums = hf.PageNumbers
                    Me.chptBase_PageNumbering_SetFormat(pgNumStyle, pgNums)
                    pgNums.RestartNumberingAtSection = False
                    pgNums.StartingNumber = startAt
            End Select

        Catch ex As Exception
            rslt = False
        End Try

        '
        Return rslt
        '
    End Function
    '
    '
    Public Function chptBase_PageNumbering_isES() As Boolean
        Dim rslt As Boolean
        Dim sect As Word.Section
        Dim hf As Word.HeaderFooter
        Dim pgNums As Word.PageNumbers
        '
        rslt = False
        sect = objGlobals.glb_get_wrdSect
        '
        Try
            hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
            pgNums = hf.PageNumbers
            '
            If pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleLowercaseRoman Then rslt = True

        Catch ex As Exception
            rslt = False
        End Try

        Return rslt
    End Function
    '
    Public Function chptBase_PageNumbering_isChapterBody() As Boolean
        Dim rslt As Boolean
        Dim sect As Word.Section
        Dim hf As Word.HeaderFooter
        Dim pgNums As Word.PageNumbers
        '
        rslt = False
        sect = objGlobals.glb_get_wrdSect
        '
        Try
            hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
            pgNums = hf.PageNumbers
            '
            If pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic And pgNums.HeadingLevelForChapter = 0 Then rslt = True

        Catch ex As Exception
            rslt = False
        End Try

        Return rslt
    End Function
    '
    '
    Public Function chptBase_PageNumbering_isAppendixBody() As Boolean
        Dim rslt As Boolean
        Dim sect As Word.Section
        Dim hf As Word.HeaderFooter
        Dim pgNums As Word.PageNumbers
        '
        rslt = False
        sect = objGlobals.glb_get_wrdSect

        Try
            hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
            pgNums = hf.PageNumbers
            '
            If pgNums.HeadingLevelForChapter = 5 Then rslt = True

        Catch ex As Exception
            rslt = False
        End Try

        Return rslt
    End Function
    '

    '

    Public Sub chptBase_PageNumbering_SetFormat(pgNumStyle As String, ByRef pgNums As Word.PageNumbers)
        Dim objPgNumberMgr As New cPageNumberMgr()
        Dim objPropsMgr As New cPropertyMgr()
        Dim strNumFormat, strResult As String
        '
        strNumFormat = objPgNumberMgr.pgNum_get_numFormat_ForDoc()
        '
        Select Case pgNumStyle
            Case "es"
                'Work around to avoid the Appendix 2 part sticking issue
                Me.chptBase_PageNumbering_SetFormat_to_2part_standard(pgNums)
                '
                pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleLowercaseRoman
                pgNums.IncludeChapterNumber = False
                pgNums.HeadingLevelForChapter = 0
                pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen
            Case "flow"
                'This is for sections that flow from the prior sections

            Case "std"
                'Work around to avoid the Appendix 2 part sticking issue
                Me.chptBase_PageNumbering_SetFormat_to_2part_standard(pgNums)
                '
                pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic
                pgNums.IncludeChapterNumber = True
                pgNums.HeadingLevelForChapter = 0
                pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen
                '
                Select Case strNumFormat
                    Case objPgNumberMgr.pgNum_pgNumType_std
                        pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic
                        pgNums.IncludeChapterNumber = False
                        pgNums.HeadingLevelForChapter = 0
                        pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen
                    Case objPgNumberMgr.pgNum_pgNumype_2part
                        pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic
                        pgNums.IncludeChapterNumber = True
                        pgNums.HeadingLevelForChapter = 0
                        pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen
                    Case Else
                        pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic
                        pgNums.IncludeChapterNumber = False
                        pgNums.HeadingLevelForChapter = 0
                        pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen

                End Select
            Case "div"
                'Work around to avoid the Appendix 2 part sticking issue
                'Work around to avoid the Appendix 2 part sticking issue
                Me.chptBase_PageNumbering_SetFormat_to_2part_standard(pgNums)
                strResult = objPropsMgr.prps_CustomProperty_get("pgNumberFormat", objPgNumberMgr.pgNum_pgNumType_std)
                '
                Select Case strResult
                    Case objPgNumberMgr.pgNum_pgNumType_std
                        pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic
                        pgNums.IncludeChapterNumber = False
                        pgNums.HeadingLevelForChapter = 0
                        pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen
                    Case objPgNumberMgr.pgNum_pgNumype_2part
                        pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleUppercaseRoman
                        pgNums.IncludeChapterNumber = False
                        pgNums.HeadingLevelForChapter = 0
                        pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen
                End Select
                '
                'Me.chptBase_PageNumbering_SetFormat_to_2part_standard(pgNums)
                '
                'pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleUppercaseRoman
                'pgNums.IncludeChapterNumber = False
                'pgNums.HeadingLevelForChapter = 0
                'pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen

            Case "2part"
                'Work around to avoid the Appendix 2 part sticking issue
                Me.chptBase_PageNumbering_SetFormat_to_2part_standard(pgNums)
                '
                pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic
                pgNums.IncludeChapterNumber = True
                pgNums.HeadingLevelForChapter = 0
                pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen
            Case "ap"
                'Work around to avoid the Appendix 2 part sticking issue
                '
                If objGlobals._glb_doApp_as_HeadingAP Then
                    Me.chptBase_PageNumbering_SetFormat_to_2part_standard(pgNums)
                    '
                    pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic
                    pgNums.IncludeChapterNumber = False
                    'pgNums.HeadingLevelForChapter = 0
                    'pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen

                Else
                    Me.chptBase_PageNumbering_SetFormat_to_2part_standard(pgNums)
                    pgNums.IncludeChapterNumber = True
                    pgNums.HeadingLevelForChapter = 5
                    pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen
                End If
                '
                '
                '
        End Select

    End Sub
    '
    ''' <summary>
    ''' This method will reset the pgNums in the HeaderFooter to a known state.
    ''' The state being 2 part standard.. We do this to work around the 'stickiness'
    ''' of the Appnedix Footer page number structure
    ''' </summary>
    ''' <param name="pgNums"></param>
    Public Sub chptBase_PageNumbering_SetFormat_to_2part_standard(ByRef pgNums As Word.PageNumbers)
        '
        Try
            'Work around to avoid the Appendix 2 part sticking issue
            pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic
            pgNums.IncludeChapterNumber = True
            pgNums.HeadingLevelForChapter = 0
            pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen

        Catch ex As Exception

        End Try
    End Sub


#End Region

#Region "Paragraph Handling"
    '
    ''' <summary>
    ''' This method will find the last Table in the Range rng, and then delete all bu the
    ''' n paragraphs between this table and the end of the section.. It will chnage rng to the
    ''' beginning of the first paragraph after the Table... If all wnet well.. If not it will
    ''' remain unchnaged
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="numParasLeft"></param>
    Public Function Base_Paragraphs_Delete(ByRef rng As Word.Range, Optional numParasLeft As Integer = 6, Optional strStyleOfParas As String = "Body Text") As Word.Range
        Dim tbl As Word.Table
        Dim i As Integer
        Dim sect, sectLast As Word.Section
        Dim para As Word.Paragraph
        '
        sect = rng.Sections.Item(1)
        sectLast = Globals.ThisAddIn.Application.ActiveDocument.Sections.Last
        '
        If rng.Tables.Count <> 0 Then
            tbl = rng.Tables.Item(rng.Tables.Count)
            'tbl = rng.Tables.Item(1)
            rng = tbl.Range
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            '
            If sect.Index = sectLast.Index Then
                '
                rng.MoveEnd(WdUnits.wdStory)
                rng.Delete()
                rng.Style = Globals.ThisAddIn.Application.ActiveDocument.Styles("Body Text")

            Else
                'We are in a standard section. So delete paragraphs to the section boundary
                'and then add numParasLeft
                '
                rng.MoveEnd(WdUnits.wdSection, 1)
                rng.Style = Globals.ThisAddIn.Application.ActiveDocument.Styles("Body Text")
                rng.MoveEnd(WdUnits.wdParagraph, -2)
                rng.Delete()
                For i = 1 To numParasLeft
                    para = rng.Paragraphs.Add()
                Next
                '
                rng.Style = Globals.ThisAddIn.Application.ActiveDocument.Styles(strStyleOfParas)
                '
                rng = tbl.Range
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            End If
            '
        Else
            'No tables, so we'll delete from the current range to the end of the section
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.MoveEnd(WdUnits.wdSection, 1)
            rng.MoveEnd(WdUnits.wdParagraph, -2)
            rng.Delete()
            '
            For i = 1 To numParasLeft
                rng.Paragraphs.Add(rng)
            Next
            '
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
        End If
        '
        Return rng
        '
    End Function
    '
#End Region'
    '
    ''' <summary>
    ''' This method will insert a table at the specified range (rng). The table will have (numRows) rows, (numColumns) columns
    ''' The row height is rowHeight (exact) and the table will be set with the style strStyleName
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="numRows"></param>
    ''' <param name="numColumns"></param>
    ''' <param name="rowHeight"></param>
    ''' <param name="strStyleName"></param>
    ''' <returns></returns>
    Public Function chptBase_insert_TableAtRange(ByRef rng As Word.Range, numRows As Integer, numColumns As Integer, rowHeight As Single, strStyleName As String) As Word.Table
        Dim tbl As Word.Table
        Dim myDoc As Word.Document
        '
        myDoc = rng.Document
        '
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        tbl = rng.Tables.Add(rng, numRows, numColumns)
        tbl.AllowAutoFit = False
        tbl.LeftPadding = 0.0
        tbl.RightPadding = 0.0
        tbl.TopPadding = 0.0
        tbl.BottomPadding = 0.0
        tbl.Rows.HeightRule = WdRowHeightRule.wdRowHeightExactly
        tbl.Rows.Height = rowHeight
        tbl.Borders.InsideLineStyle = WdLineStyle.wdLineStyleNone
        tbl.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone
        '
        Try
            'Set the default style for the Header Table
            tbl.Range.Style = myDoc.Styles(strStyleName)
        Catch ex As Exception

        End Try
        '
        Return tbl
    End Function
    '
    '
    'This function checks for illegal insertion actions. It first makes certain that
    'the user is not trying to insert anything is the Cover Page, Contacts Page (Front and Back),
    'Table of Contents.. If so, we get an error message and it returns False. It
    'will then check to see if we are trying to insert in a Table. If so, then we get
    'an error message and it will return false
    Public Overridable Function chptBase_is_OKToInsert(ByRef sect As Section, Optional doTableCheck As Boolean = True) As String
        Dim objInsertTestMgr As New cInsertTestMgr()
        Dim strResult As String
        '
        strResult = objInsertTestMgr.ins_is_OKToInsert(sect, doTableCheck)
        '
        Return strResult
        '
    End Function
    '

End Class

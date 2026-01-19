Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.Windows.Forms
Public Class cReport
    Inherits cChptBase

    Public modePropertyNameForReport As String             'Used to determine whether the report is short or long
    Public modeShort As String
    Public rpt_isPrt As String
    Public rpt_isLnd As String
    Public rpt_isBrief As String
    '
    Public shortReportH1FontSize As Single
    Public shortReportNumberFontSize As Single
    '
    Public longReportH1FontSize As Single
    Public longReportNumberFontSize As Single
    '
    'Public objGlobals As cGlobals
    '

    Public Sub New()
        MyBase.New()
        Me.modePropertyNameForReport = "reportMode"
        Me.modeShort = "shortReport"
        Me.rpt_isPrt = "longReport"
        Me.rpt_isLnd = "longReport_Lnd"
        Me.rpt_isBrief = "aaBrief"
        '
        Me.longReportH1FontSize = 26.0
        Me.longReportNumberFontSize = 80

        Me.shortReportH1FontSize = 20
        Me.shortReportNumberFontSize = 40.0             'Size of Panel chapter Number
        '
        Me.objGlobals = New cGlobals()
    End Sub
    '
    ''' <summary>
    ''' Will return True if the current report is in Portrait Mode
    ''' </summary>
    ''' <returns></returns>
    Public Function Rpt_Mode_IsPortraitRpt() As Boolean
        Dim rslt As Boolean
        '
        rslt = False
        If Me.Rpt_Mode_Get() = Me.rpt_isPrt Then rslt = True
        '
        Return rslt
    End Function
    '
    ''' <summary>
    ''' Will return True if the current report is in Landscape Mode
    ''' </summary>
    ''' <returns></returns>
    Public Function Rpt_Mode_IsLandscapeRpt() As Boolean
        Dim rslt As Boolean
        '
        rslt = False
        If Me.Rpt_Mode_Get() = Me.rpt_isLnd Then rslt = True
        '
        Return rslt
    End Function
    '
    Public Function Rpt_Mode_IsBrief() As Boolean
        Dim rslt As Boolean
        '
        rslt = False
        If Me.Rpt_Mode_Get() = Me.rpt_isBrief Then rslt = True
        '
        Return rslt
    End Function

    '
    '
    ''' <summary>
    ''' This method adjusts the styles within the report to suit the requirements of each document type
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="strReportMode"></param>
    Public Sub Rpt_styles_Upgrade_for_ReportType(ByRef myDoc As Word.Document, strReportMode As String)
        Dim table_DeltaLeftIndent As Single
        Dim objlstStyles As New cStyles_ListStyles()
        Dim objHeadingLevels As New cStyles_HeadingLevels()
        Dim objStylesMgr As New cStylesManager()
        Dim myStyle As Word.Style
        'Dim paraFormat As Word.ParagraphFormat
        Dim objTOCMgr As New cTOCMgr()
        '
        table_DeltaLeftIndent = 2.0
        myStyle = Nothing
        '
        'objStylesMgr.style_upgrade_TemplateStyles(myDoc)
        '
        '
        Select Case strReportMode
            Case Me.rpt_isPrt, Me.rpt_isLnd
                'Put styles back the way they were
                'Me.Rpt_Styles_resetStyles_fromTemplate()
                '
                'objTOCMgr.TOC_Styles_AdjustForReportMode()
                objTOCMgr.TOC_Styles_AdjustForReportMode(strReportMode)

                myStyle = myDoc.Styles.Item("Heading 1 (no number)")
                myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
                myStyle.ParagraphFormat.LineSpacing = 34
                myStyle.ParagraphFormat.SpaceBefore = 0
                myStyle.ParagraphFormat.PageBreakBefore = True
                myStyle.Font.Size = 28

                '
                myStyle = myDoc.Styles.Item("Cp Title")
                myStyle.Font.Size = 40
                myStyle.ParagraphFormat.SpaceAfter = 20.0

                myStyle = myDoc.Styles.Item("Cp SubTitle")
                myStyle.Font.Size = 24
                myStyle.ParagraphFormat.SpaceAfter = 20.0

                myStyle = myDoc.Styles.Item("Cp Report Date")


            Case Me.rpt_isBrief
                'Me.Rpt_Styles_resetStyles_fromTemplate()
                'objTOCMgr.TOC_Styles_AdjustForReportMode()
                objTOCMgr.TOC_Styles_AdjustForReportMode(Me.rpt_isBrief)


                'myDoc.Styles.Item("Heading 1").ParagraphFormat.PageBreakBefore = False
                'myDoc.Styles.Item("Heading 1").ParagraphFormat.PageBreakBefore = False
                'myDoc.Styles.Item("Heading 1 (ES)").ParagraphFormat.PageBreakBefore = False
                'myDoc.Styles.Item("Heading 1 (AP)").ParagraphFormat.PageBreakBefore = False
                myStyle = myDoc.Styles.Item("Heading 1 (no number)")
                myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
                myStyle.ParagraphFormat.LineSpacing = 27
                myStyle.ParagraphFormat.SpaceBefore = 14
                myStyle.ParagraphFormat.PageBreakBefore = False
                myStyle.Font.Size = 20


                '
                myStyle = myDoc.Styles.Item("Cp Title")
                myStyle.Font.Size = 20
                myStyle.ParagraphFormat.SpaceAfter = 6.0

                myStyle = myDoc.Styles.Item("Cp SubTitle")
                myStyle.Font.Size = 14
                myStyle.ParagraphFormat.SpaceAfter = 6.0

                myStyle = myDoc.Styles.Item("Cp Report Date")


        End Select
        '
        'If Not objlstStyles.style_adjust_TableStyles(myDoc, table_DeltaLeftIndent) Then MsgBox("Error upgrading Table Styles")
        '
    End Sub
    '
    '
    Public Sub Rpt_build_newAABrief_PrtandLnd()
        Dim objSectMgr As New cSectionMgr()
        Dim objMsgMgr As New cMessageManager()
        Dim objParas As New cParas()
        Dim objViewMgr As New cViewManager()
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objBackPanelMgr As New cBackPanelMgr()
        Dim objBrandMgr As New cBrandMgr()
        Dim objContactsMgr As New cContactsMgr()
        Dim objCpMgr As New cCoverPageMgr()
        Dim objTblsMgr As New cTablesMgr()
        Dim objStylesMgr As New cStylesManager()
        Dim objBBMgr As New cBBlocksHandler()
        Dim objWCAGMgr As New cWCAGMgr()
        Dim objBnrMgr As New cChptBanner()
        Dim objPageNumberMgr As New cPageNumberMgr()
        Dim objPrint As New cPrintAndDisplayServices()

        Dim tbl As Word.Table
        Dim myDoc As Word.Document
        Dim sect As Word.Section
        Dim hf As Word.HeaderFooter
        Dim rng As Word.Range
        Dim placeBehind As Boolean
        Dim shp As Word.Shape
        Dim startTime As Date
        Dim stpWatch As System.Diagnostics.Stopwatch
        Dim Interval As TimeSpan
        Dim strElapsedTime, strRptMode As String
        '
        Me.objGlobals.glb_screen_update(True)
        placeBehind = True
        myDoc = Me.objGlobals.glb_get_wrdActiveDoc()
        'lst = objBnrMgr.bnr_get_BannerSettings(objBnrMgr.sectType_body, True)
        '
        Me.objGlobals.glb_cursors_setToWait()
        '
        'strRptMode = Me.Rpt_Mode_SetAs_Std()                   Set in PIF file
        strRptMode = Me.Rpt_Mode_Get()
        '
        '
        '*** Temporarily attach the template to make the building blocks available
        '
        'tmpl = objGlobals.glb_get_wrdActiveDoc.AttachedTemplate
        'strTemplateFullName = objGlobals.glb_getTmpl_FullName()
        '
        'myDoc.AttachedTemplate = strTemplateFullName
        'myDoc.UpdateStyles()
        '
        '***
        '
        'Mechanism to adjust styles without having to chnage the template
        Me.Rpt_styles_Upgrade_for_ReportType(myDoc, strRptMode)
        '
        stpWatch = System.Diagnostics.Stopwatch.StartNew()
        startTime = TimeOfDay()
        '
        Try
            objSectMgr.sct_delete_allSections()
            '
            'To get rid of prior document 'ghosts'
            Me.objGlobals.glb_screen_update(True)

            'sect = myDoc.Sections.Item(1)
            sect = myDoc.Sections.Last
            sect.PageSetup.DifferentFirstPageHeaderFooter = True
            '
            'Default to standard page number format
            objPageNumberMgr.pgNum_set_numFormat_ForDoc()
            Me.chptBase_PageNumbering_Set(sect, False, 1, "std")
            '
            objSectMgr.sct_reset_ToPortrait(sect)
            objHfMgr.hf_headers_insert(sect)
            objHfMgr.hf_footers_insert(sect)
            '
            'To reinforce the getting rid of prior document 'ghosts
            Me.objGlobals.glb_screen_update(True)
            '
            '
            objHfMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_brief))
            objHfMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_brief), "firstPage")
            '
            For i = 1 To 3
                'rng = sect.Range
                'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                'sect = myDoc.Sections.Add(rng, WdSectionStart.wdSectionNewPage)
                'objParas.paras_insert_numParas(sect, 6)
                'objHfMgr.hf_hfs_linkUnlinkAll(sect, False)
            Next
            '
            sect = myDoc.Sections.First
            '
            'hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
            'shp = objBackPanelMgr.pnl_BackPanel_Insert(hf)
            'shp.Height = sect.PageSetup.PageHeight * 0.2
            'shp.ZOrder(MsoZOrderCmd.msoSendBehindText)
            'hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
            'objBrandMgr.brnd_recolour_Logo(hf)
            '
            '
            hf = objBackPanelMgr.pnl_BackPanelBriefFirstPage_Insert(sect)
            '
            Me.objGlobals.glb_screen_update(False)
            '
            'hf = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
            rng = hf.Range
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            'rng.Move(WdUnits.wdCharacter, -1)
            rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_Cpg_PictFilled_Lnd", "CoverPage", rng)             'lnd purple triangles picture
            shp = rng.ShapeRange.Item(1)
            shp.Name = "cp_pict_large"
            shp.LockAspectRatio = True
            shp.Height = 78.0
            shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
            shp.Left = 35
            shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
            shp.Top = 56
            '
            'Me.objGlobals.glb_screen_update(False)
            '
            'Me.cp_rename_Shape(rng, "cp_pict_purplePattern_lnd")
            'Me.cp_rename_Shape(rng, "cp_pict_large")
            objWCAGMgr.wcag_set_decorative(rng, True)

            '
            objParas.paras_insert_numParas(sect)
            '
            'hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
            Me.chptBase_PageNumbering_Set(sect, True, 1, "std")

            'pgNums = hf.PageNumbers
            'pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleLowercaseRoman
            'pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic
            'pgNums.IncludeChapterNumber = True
            'pgNums.HeadingLevelForChapter = 0
            'pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen
            '
            'objGlobals.glb_screen_update()
            sect = myDoc.Sections.Last
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            sect = myDoc.Sections.Add(rng, WdSectionStart.wdSectionNewPage)
            objParas.paras_insert_numParas(sect, 6)
            objHfMgr.hf_hfs_linkUnlinkAll(sect, False)

            sect = myDoc.Sections.Last
            objContactsMgr.contacts_convert_toBackContacts(sect)
            '
            objGlobals.glb_screen_update(True)
            '
            sect = myDoc.Sections.First
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            '***
            'Add a para at the beginning to allow for the insertion of letter/memo
            Dim para As Word.Paragraph
            para = rng.Paragraphs.Add()
            para.Range.Style = "spacer"
            rng = para.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Move(WdUnits.wdParagraph, 1)
            'para = rng.Paragraphs.First
            'para.Range.Select()
            rng.Select()
            '
            '***
            '
            'GoTo finis2
            '
            tbl = rng.Tables.Add(rng, 1, 1)
            objGlobals.glb_tbl_apply_aacTableNoLinesStyle(tbl)
            tbl.Rows.Item(1).HeightRule = WdRowHeightRule.wdRowHeightExactly
            tbl.Rows.Item(1).Height = 110
            '
            rng = tbl.Range.Cells.Item(1).Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            objCpMgr.cp_insert_formattedTitleText(rng)
            '
            rng = tbl.Range
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            'rng.Move(WdUnits.wdParagraph, 1)
            rng.Select()
            '
            'objStylesMgr.styles_format_styleSetRpt()
            objStylesMgr.styles_insert_StartupText_ReportES(rng)
            '
finis2:
            '
            'rng = sect.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'rng.Text = "First section"
            '
            'objViewMgr.vw_change_ColumnsAndRows(sect)
            'objViewMgr.vw_change_toPageFitBestFit(sect)
            objViewMgr.vw_change_ColumnsAndRows(sect)
            '
            Me.objGlobals.glb_screen_update(True)


            stpWatch.Stop()
            Interval = stpWatch.Elapsed()
            '
            '
            '*** Reattach original template (Normal)
            'myDoc.AttachedTemplate = ""

            'If Not tmpl.FullName Like "*Normal" Then
            'Wnat the Normal on the new target machine to point to that machine's 
            'myDoc.AttachedTemplate = ""
            'Else
            'myDoc.AttachedTemplate = tmpl.FullName
            'End If
            '***
            '
            '
            'endTime = TimeOfDay()
            'Interval = endTime - startTime
            strElapsedTime = Int(Interval.TotalSeconds()) & " Seconds"
            '
            objGlobals.glb_screen_update(True)

            MsgBox("The Report build is complete (" + strElapsedTime + ")")
            '
            'Globals.Ribbons.

        Catch ex As Exception
            '
            '*** Reattach original template (Normal)
            myDoc.AttachedTemplate = ""

            'If Not tmpl.FullName Like "*Normal" Then
            'Wnat the Normal on the new target machine to point to that machine's Normal
            'myDoc.AttachedTemplate = ""
            'Else
            'myDoc.AttachedTemplate = tmpl.FullName
            'End If
            '***
            '
        End Try
        '
finis:
        '
        Me.objGlobals.glb_cursors_setToNormal()
        '
        objPrint.colour_display_ToEasyView(objGlobals.glb_get_wrdActiveDoc())
        '
        objGlobals.glb_screen_update(True)

    End Sub
    '
    ''' <summary>
    ''' This method build a new report (Prt or Lnd) or a new Brief (depending on the
    ''' value of rptMode ('Prt', 'Lnd' or 'Brf').. It does so by using the example
    ''' documents stored in the Resources file. A build is typically 1 second on my machine
    ''' </summary>
    ''' <param name="rptMode"></param>
    Public Sub Rpt_build_fastReportOrBrief_fromTemplate(Optional rptMode As String = "Prt")
        Dim objFileMgr As New cFileHandler()
        Dim objScratchMgr As New cFileScratchMgr()
        Dim objCpMgr As New cCoverPageMgr()
        Dim objThmMgr As New cThemeMgr()
        Dim objRptMgr As New cReport()
        Dim objStylesMgr As New cStylesManager()
        Dim objTOCMgr As New cTOCMgr()
        Dim objViewMgr As New cViewManager()
        Dim myDoc, residualDoc As Word.Document
        Dim rng As Word.Range
        'Dim srcDocFileInfo As System.IO.FileInfo
        Dim docSourceFullName As String
        Dim placeBehind As Boolean
        Dim strRptMode As String
        Dim startTime As Date
        Dim stpWatch As System.Diagnostics.Stopwatch
        Dim Interval As TimeSpan
        Dim strElapsedTime As String
        '
        Me.objGlobals.glb_screen_update(False)
        '
        placeBehind = True
        'myDoc = Me.objGlobals.glb_get_wrdActiveDoc()
        residualDoc = objGlobals.glb_get_wrdActiveDoc()

        '
        Me.objGlobals.glb_cursors_setToWait()
        'Mechanism to adjust styles without having to chnage the template
        '
        stpWatch = System.Diagnostics.Stopwatch.StartNew()
        startTime = TimeOfDay()
        '
        Try
            docSourceFullName = objFileMgr.file_get_RptExampleFromResources(rptMode)
            'srcDocFileInfo = New IO.FileInfo(docSourceFullName)
            '
            myDoc = objGlobals.glb_get_wrdApp.Documents.Add(docSourceFullName)
            myDoc.AttachedTemplate = objGlobals.glb_getTmpl_FullName()
            '
            objScratchMgr.scratch_delete_Directory_Scratch()
            '
            'Get rid of the empty residual
            If objGlobals.glb_doc_isEmptyAndNotSaved(residualDoc) Then
                residualDoc.Saved = True
                residualDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            End If
            '
            Select Case rptMode
                Case "Prt"
                    strRptMode = Me.Rpt_Mode_SetAs_Std()
                    '
                    'Adjust styles
                    objThmMgr.thm_Set_ThemeToAAStd_fromFile(myDoc)
                    'objRptMgr.Rpt_Styles_resetStyles_fromTemplate(True)
                    objStylesMgr.style_extend_TemplateStyles()
                    '
                    objViewMgr.vw_change_ColumnsAndRows(3, 1, 69)
                    'objViewMgr.vw_change_toPageFitBestFit(myDoc)
                    '

                Case "Lnd"
                    'strRptMode = Me.modeLongLandscape
                    strRptMode = Me.Rpt_Mode_SetAsLandScape()
                    '
                    'Adjust Styles
                    objThmMgr.thm_Set_ThemeToAAStd_fromFile(objGlobals.glb_get_wrdActiveDoc)
                    'objRptMgr.Rpt_Styles_resetStyles_fromTemplate(True)
                    objStylesMgr.style_extend_TemplateStyles()                                                                      'Refresh the styles
                    Me.Rpt_styles_Upgrade_for_ReportType(objGlobals.glb_get_wrdActiveDoc, objRptMgr.rpt_isLnd)               'Upgrade/chnage depending on report mode
                    'objTOCMgr.TOC_Styles_AdjustForReportMode(objRptMgr.rpt_isLnd)                                                   'Force the style for the Brief
                    '
                    objViewMgr.vw_change_ColumnsAndRows(2, 1)

                Case "Brf"
                    strRptMode = Me.Rpt_Mode_SetAsAABrief
                    '
                    'Adjust Styles
                    objThmMgr.thm_Set_ThemeToAAStd_fromFile(objGlobals.glb_get_wrdActiveDoc)
                    'objRptMgr.Rpt_Styles_resetStyles_fromTemplate(True)
                    objStylesMgr.style_extend_TemplateStyles()                                                                          'Refresh the styles
                    Me.Rpt_styles_Upgrade_for_ReportType(objGlobals.glb_get_wrdActiveDoc, objRptMgr.rpt_isBrief)                 'Force the styles for the Brief
                    'objTOCMgr.TOC_Styles_AdjustForReportMode(objRptMgr.rpt_isBrief)                                                     'Force the style for the Brief
                    '
                    objViewMgr.vw_change_ColumnsAndRows(2, 1)

                Case Else
                    'strRptMode = Me.modeLong
                    strRptMode = Me.Rpt_Mode_SetAs_Std()
            End Select
            '
            'Me.Rpt_styles_Upgrade_for_ReportType(myDoc, strRptMode)
            '
            stpWatch.Stop()
            Interval = stpWatch.Elapsed()
            '
            myDoc.BuiltInDocumentProperties("Author") = ""
            myDoc.BuiltInDocumentProperties("Company") = "ACIL Allen"
            myDoc.BuiltInDocumentProperties("Comments") = "© ACIL Allen " + Now().Year.ToString("D4")
            '
            objGlobals.glb_view_setToPrintLayout()
            'objCpMgr.cp_sel_MoveToTitle(myDoc.Sections.First, True)
            '
            'objPrint.colour_display_ToEasyView(myDoc)
            'objStyles.style_copy_StylesFromTemplate(objGlobals.glb_get_wrdActiveDoc)
            '
            '
            'endTime = TimeOfDay()
            'Interval = endTime - startTime
            'strElapsedTime = Int(Interval.TotalSeconds()) & " Seconds"
            Interval.TotalSeconds.ToString("f1")
            'strElapsedTime = Int(Interval.TotalMilliseconds()) & " milliseconds"
            strElapsedTime = Interval.TotalSeconds.ToString("f1") & " seconds"

            '
            '
            objGlobals.glb_screen_update(True)
            '
            MsgBox("The Report build is complete (" + strElapsedTime + ")")
            objGlobals.glb_doc_checkDocType_ActivateTab()
            '
            Select Case rptMode
                Case "Brf"
                    rng = myDoc.Sections.First.Range.Tables.Item(1).Range.Cells.Item(1).Range
                    rng = rng.Paragraphs.Item(1).Range
                    rng.MoveEnd(WdUnits.wdCharacter, -1)
                Case Else
                    rng = objCpMgr.cp_sel_MoveToTitle(myDoc.Sections.First, True)
            End Select
            '
            rng.Select()
            '
            objGlobals.glb_screen_update(True)
            '

        Catch ex As Exception
            objGlobals.glb_view_setToPrintLayout()
        End Try
        '
    End Sub

    '
    'This method build a new report (Prt or Lnd) or a new Brief (depending on the
    'value of rptMode

    ''' <summary>
    ''' This method build a new report (Prt or Lnd) or a new Brief (depending on the
    ''' value of rptMode ('Prt', 'Lnd' or 'Brf').. It does so by using the example
    ''' documents stored in the Resources file. A build is typically 1 second on my machine
    ''' </summary>
    ''' <param name="rptMode"></param>
    Public Sub Rpt_build_fastReportOrBrief_byCopy(Optional rptMode As String = "Prt")
        Dim objFileMgr As New cFileHandler()
        Dim objSectMgr As New cSectionMgr()
        Dim objParas As New cParas()
        Dim objCpMgr As New cCoverPageMgr()
        '
        Dim rng As Word.Range
        Dim sect As Word.Section
        Dim myDoc, srcDoc As Word.Document
        Dim srcDocFileInfo As System.IO.FileInfo
        Dim docSourceFullName As String
        'Dim tmpl As Word.Template
        Dim placeBehind As Boolean
        'Dim strTemplateFullName, strRptMode As String

        Dim strRptMode As String
        Dim startTime As Date
        Dim stpWatch As System.Diagnostics.Stopwatch
        '
        Dim rngTarget As Word.Range
        Dim Interval As TimeSpan
        Dim strElapsedTime As String
        '
        Me.objGlobals.glb_screen_update(True)
        '
        placeBehind = True
        myDoc = Me.objGlobals.glb_get_wrdActiveDoc()
        '
        Select Case rptMode
            Case "Prt"
                'strRptMode = Me.modeLong
                strRptMode = Me.Rpt_Mode_SetAs_Std()
            Case "Lnd"
                'strRptMode = Me.modeLongLandscape
                strRptMode = Me.Rpt_Mode_SetAsLandScape()
            Case "Brf"
                strRptMode = Me.Rpt_Mode_SetAsAABrief
                'Mechanism to adjust styles without having to chnage the template
                'Me.Rpt_styles_Upgrade_for_ReportType(myDoc, strRptMode)

            Case Else
                'strRptMode = Me.modeLong
                strRptMode = Me.Rpt_Mode_SetAs_Std()
        End Select
        '
        '
        Me.objGlobals.glb_cursors_setToWait()
        'Mechanism to adjust styles without having to chnage the template
        Me.Rpt_styles_Upgrade_for_ReportType(myDoc, strRptMode)
        '
        stpWatch = System.Diagnostics.Stopwatch.StartNew()
        startTime = TimeOfDay()
        '
        Try
            docSourceFullName = objFileMgr.file_get_RptExampleFromResources(rptMode)
            srcDocFileInfo = New IO.FileInfo(docSourceFullName)
            '
            'Delete sections down to one, then add another to allow for the trouble free copy
            'from the example document to base document
            objSectMgr.sct_delete_allSections()
            myDoc.Sections.Add()
            '
            Select Case rptMode
                Case "Prt"
                    myDoc.PageSetup.Orientation = WdOrientation.wdOrientPortrait
                Case "Lnd"
                    myDoc.PageSetup.Orientation = WdOrientation.wdOrientLandscape
                Case "Brf"
                    myDoc.PageSetup.Orientation = WdOrientation.wdOrientPortrait
                    myDoc.PageSetup.DifferentFirstPageHeaderFooter = True
                Case Else
                    myDoc.PageSetup.Orientation = WdOrientation.wdOrientPortrait
            End Select
            '
            sect = myDoc.Sections.Last
            '
            'GoTo loop1
            '
            srcDoc = objGlobals.glb_get_wrdApp.Documents.Open(docSourceFullName, Visible:=False)
            srcDoc.Content.Copy()
            '
            rngTarget = myDoc.Range
            rngTarget.Collapse(WdCollapseDirection.wdCollapseEnd)
            rngTarget.Paste()
            '
            'To cater for idiocentricities
            objSectMgr.sct_delete_Section(myDoc.Sections.First)
            sect = myDoc.Sections.Last
            objParas.paras_delete_Paragraphs(sect.Range, 1)
            '
            srcDoc.Close(WdSaveOptions.wdDoNotSaveChanges)
            objFileMgr.file_delete_File(docSourceFullName)
            Try
                srcDocFileInfo.Directory.Delete(True)
            Catch ex2 As Exception

            End Try
            '
            '
            myDoc.BuiltInDocumentProperties("Author") = ""
            myDoc.BuiltInDocumentProperties("Company") = "ACIL Allen"
            myDoc.BuiltInDocumentProperties("Comments") = "© ACIL Allen " + Now().Year.ToString("D4")
            '
            objGlobals.glb_view_setToPrintLayout()
            '
            objCpMgr.cp_sel_MoveToTitle(myDoc.Sections.First, True)

            'objPrint.colour_display_ToEasyView(myDoc)
loop1:
            'objStyles.style_copy_StylesFromTemplate(objGlobals.glb_get_wrdActiveDoc)
            stpWatch.Stop()
            Interval = stpWatch.Elapsed()
            '
            'endTime = TimeOfDay()
            'Interval = endTime - startTime
            'strElapsedTime = Int(Interval.TotalSeconds()) & " Seconds"
            strElapsedTime = Int(Interval.TotalMilliseconds()) & " milliseconds"
            '
            'objCpMgr.cp_set_SelectionToTitle(myDoc.Sections.Item(1))
            '
            '*** Reattach original template (Normal)
            'myDoc.AttachedTemplate = tmpl.FullName
            '***
            '
            objGlobals.glb_screen_update(True)
            '

            MsgBox("The Report build is complete (" + strElapsedTime + ")")
            '
            '
            rng = objCpMgr.cp_sel_MoveToTitle(myDoc.Sections.First, True)
            objGlobals.glb_screen_update(True)
            '
            rng.Select()

        Catch ex As Exception
            objGlobals.glb_view_setToPrintLayout()
        End Try
        '

    End Sub


    Public Sub Rpt_build_newReport_PrtandLnd()
        Dim objSectMgr As New cSectionMgr()
        Dim objFldsMgr As New cFieldsMgr()
        Dim objCpMgr As New cCoverPageMgr()
        Dim objMsgMgr As New cMessageManager()
        Dim objExecMgr As New cChptExec()
        Dim objViewMgr As New cViewManager()
        Dim objStylesMgr As New cStylesManager()
        Dim objChptBnr As New cChptBanner()
        Dim objGlossary As New cGlossary()
        Dim objPrint As New cPrintAndDisplayServices()
        Dim objPageNumberMgr As New cPageNumberMgr()

        Dim sect As Word.Section
        Dim myDoc As Word.Document
        Dim strRptMode As String
        Dim placeBehind As Boolean
        Dim objParas As New cParas()
        Dim objBnrMgr As New cChptBanner()
        Dim rng As Word.Range
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objTocMgr As New cTOCMgr()
        Dim objContactsMgr As New cContactsMgr()
        Dim objDivMgr As New cChptDivider()
        Dim hf As Word.HeaderFooter
        Dim pgNums As Word.PageNumbers
        Dim startTime As Date
        Dim stpWatch As System.Diagnostics.Stopwatch
        Dim Interval As TimeSpan
        Dim strElapsedTime As String
        '
        Me.objGlobals.glb_screen_update(True)
        placeBehind = True
        myDoc = Me.objGlobals.glb_get_wrdActiveDoc()
        '
        '*** Temporarily attach the template to make the building blocks available
        '
        'tmpl = objGlobals.glb_get_wrdActiveDoc.AttachedTemplate
        'strTemplateFullName = objGlobals.glb_getTmpl_FullName()
        '
        'myDoc.AttachedTemplate = strTemplateFullName
        'myDoc.UpdateStyles()
        '
        '***


        'lst = objBnrMgr.bnr_get_BannerSettings(objBnrMgr.sectType_body, True)
        '
        'rslt = objMsgMgr.deleteAllMessage
        '
        'If Not rslt Then GoTo finis
        '
        Me.objGlobals.glb_cursors_setToWait()

        'strRptMode = Me.Rpt_Mode_SetAs_Std()                   Set in PIF file
        strRptMode = Me.Rpt_Mode_Get()
        '
        'Mechanism to adjust styles without having to chnage the template
        Me.Rpt_styles_Upgrade_for_ReportType(myDoc, strRptMode)
        '
        stpWatch = System.Diagnostics.Stopwatch.StartNew()
        startTime = TimeOfDay()
        '
        '
        Try
            'Build a document of empty standard sections
            objSectMgr.sct_delete_allSections()
            Me.objGlobals.glb_screen_update(True)
            '
            '
            'sect = myDoc.Sections.Item(1)
            sect = myDoc.Sections.Last
            '
            'Default to standard page number format
            objPageNumberMgr.pgNum_set_numFormat_ForDoc()
            Me.chptBase_PageNumbering_Set(sect, False, 1, "std")
            '
            '
            'objSectMgr.sct_reset_ToLandscape(sect)
            Select Case strRptMode
                Case Me.rpt_isPrt
                    objSectMgr.sct_reset_ToPortrait(sect)
                Case Me.rpt_isLnd
                    objSectMgr.sct_reset_ToLandscape(sect)
            End Select
            '
            Me.objGlobals.glb_screen_update(True)
            '
            objParas.paras_insert_numParas(sect)
            '
            hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
            pgNums = hf.PageNumbers
            'pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleLowercaseRoman
            pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic
            pgNums.IncludeChapterNumber = True
            pgNums.HeadingLevelForChapter = 0
            pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen

            '
            'rng = sect.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'rng.Text = "First section"
            '
            'objViewMgr.vw_change_ColumnsAndRows(sect)
            'objViewMgr.vw_change_toPageFitBestFit(sect)
            objViewMgr.vw_change_ColumnsAndRows(sect)
            '
            Me.objGlobals.glb_screen_update(False)
            '
            For i = 1 To 12                                         '1 to 13 if References page
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                sect = myDoc.Sections.Add(rng, WdSectionStart.wdSectionNewPage)
                'sect = myDoc.Sections.Last
                'objParas.paras_insert_numParas(sect, 8)
                objParas.paras_insert_numParas(sect, 6)
                objHFMgr.hf_hfs_linkUnlinkAll(sect, False)
            Next
            '
            'Cover Page
            sect = myDoc.Sections.First
            objCpMgr.cp_convert_ToCoverPage(sect)
            '
            Me.objGlobals.glb_screen_update(False)
            '
            'Front Contacts
            sect = myDoc.Sections.Item(2)
            objContactsMgr.contacts_convert_toFrontContacts(sect)
            '
            '
            'Glossary
            '
            'sect = myDoc.Sections.Item(myDoc.Sections.Last.Index - 11)
            sect = myDoc.Sections.Item(4)
            objParas.paras_add_textAndStyle(sect, "Glossary", "Heading (glossary)")
            objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_glos))
            '
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Move(WdUnits.wdParagraph, 1)
            objGlossary.glos_Insert_TableForGlossary(rng)
            '
            'objParas.paras_delete_Paragraphs(sect.Range, 2)
            '
            Me.chptBase_PageNumbering_Set(sect, True, 1, "es")
            '
            'Divider between Glossaryry and ES.. Must have es numbering
            '
            sect = myDoc.Sections.Item(5)
            objDivMgr.div_convert_toDivider(sect, objBnrMgr.tag_div, "Summary Report")          'Set Optional to True if we want to see the footer
            'sect = myDoc.Sections.Item(5)
            Me.chptBase_PageNumbering_Set(sect, False, 1, "es")
            '
            'Executive Summary
            '
            'sect = myDoc.Sections.Item(myDoc.Sections.Last.Index - 8)
            sect = myDoc.Sections.Item(6)
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'objParas.paras_add_textAndStyle(sect, "ES Heading", "Heading 1 (ES)")
            objStylesMgr.styles_insert_StartupText_ReportES(rng)
            objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_es))
            Me.chptBase_PageNumbering_Set(sect, False, 1, "es")
            '
            'Divider between ES and Firstt Chapter.. Must have body numbering
            '
            sect = myDoc.Sections.Item(7)
            objDivMgr.div_convert_toDivider(sect, objBnrMgr.tag_div, "Main Report")             'Set Optional to True if we want to see the footer
            'sect = myDoc.Sections.Item(5)
            Me.chptBase_PageNumbering_Set(sect, True, 1, "std")
            '


            'chptBase_PageNumbering_Set(sect, False, 1, "div")

            '
            'hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
            'pgNums = hf.PageNumbers
            'pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleLowercaseRoman
            'pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic
            'pgNums.IncludeChapterNumber = False
            'pgNums.RestartNumberingAtSection = True
            'pgNums.StartingNumber = 1
            'pgNums.HeadingLevelForChapter = 0
            'pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen

            'chptBase_PageNumbering_Set(sect, True, 1, "es"

            sect = myDoc.Sections.Item(myDoc.Sections.Last.Index - 5)           '-6 if References
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'objParas.paras_add_textAndStyle(sect, "Heading 1", "Heading 1")
            objStylesMgr.styles_insert_StartupText_ReportBody(rng)
            objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_chpt_body))
            chptBase_PageNumbering_Set(sect, False, 1, "std")


            sect = myDoc.Sections.Item(myDoc.Sections.Last.Index - 4)           '-5 if References
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'objParas.paras_add_textAndStyle(sect, "Heading 1", "Heading 1")
            objStylesMgr.styles_insert_StartupText_ReportBody(rng)
            objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_chpt_body))
            chptBase_PageNumbering_Set(sect, False, 1, "std")
            '
            '***References here
            'sect = myDoc.Sections.Item(myDoc.Sections.Last.Index - 4)
            'objParas.paras_add_textAndStyle(sect, "References", "Heading (glossary)")
            'rng = sect.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'rng.Move(WdUnits.wdParagraph, 2)
            'sect.Range.Document.Fields.Add(rng, WdFieldType.wdFieldBibliography,, True)
            '
            'Me.chptBase_PageNumbering_Set(tbl.Range.Sections.Item(1), False, 1, "div")
            'Me.chptBase_PageNumbering_Set(sect, False, 1, "std")
            'objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_glos_bib))
            '

            '
            sect = myDoc.Sections.Item(myDoc.Sections.Last.Index - 3)
            objDivMgr.div_convert_toDivider(sect, objBnrMgr.tag_divAP, "Appendices")            'Set Optional to True if we want to see the footer
            Me.chptBase_PageNumbering_Set(sect, False, 1, "std")
            'chptBase_PageNumbering_Set(sect, False, 1, "div")


            'objDivMgr.div_convert_toAPDivider(sect)
            'objBnrMgr.bnr_insert_BannerBase(sect.Range, False, strRptMode, objBnrMgr.bnr_get_BannerSettings(objBnrMgr.tag_divAP, False))
            'chptBase_PageNumbering_Set(sect, False, 1, "div")

            sect = myDoc.Sections.Item(myDoc.Sections.Last.Index - 2)
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            objStylesMgr.styles_insert_StartupText_ReportAP(rng)

            If objGlobals._glb_doApp_as_HeadingAP Then
                'objParas.paras_add_textAndStyle(sect, "Appendix Heading 1", "Heading 1 (AP)")
                chptBase_PageNumbering_Set(sect, True, 1, "std")
            Else
                'objParas.paras_add_textAndStyle(sect, "Appendix Heading 1", "Heading 6")
                chptBase_PageNumbering_Set(sect, True, 1, "ap")
            End If
            objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_chpt_AP))


            sect = myDoc.Sections.Item(myDoc.Sections.Last.Index - 1)
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            objStylesMgr.styles_insert_StartupText_ReportAP(rng)

            If objGlobals._glb_doApp_as_HeadingAP Then
                'objParas.paras_add_textAndStyle(sect, "Appendix Heading 1", "Heading 1 (AP)")
                chptBase_PageNumbering_Set(sect, True, 1, "std")
            Else
                'objParas.paras_add_textAndStyle(sect, "Appendix Heading 1", "Heading 6")
                chptBase_PageNumbering_Set(sect, True, 1, "ap")
            End If
            objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_chpt_AP))
            '

            ' sect = objContactsMgr.contacts_insert_BackPage()
            'sect = myDoc.Sections.Item(10)
            sect = myDoc.Sections.Last

            objContactsMgr.contacts_convert_toBackContacts(sect)

            'Do primary and first page
            'objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_cont_Back))
            'objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_cont_Back), "firstPage")
            '
            '
            sect = myDoc.Sections.Item(3)
            objTocMgr.toc_convert_toTOC(sect)
            objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_toc))

            objTocMgr.toc_replace_TOCField(sect, "aac_TOC_Levels02")
            '
finis0:
            '
            objPrint.colour_display_ToEasyView(myDoc)

            stpWatch.Stop()
            Interval = stpWatch.Elapsed()
            '
            'endTime = TimeOfDay()
            'Interval = endTime - startTime
            strElapsedTime = Int(Interval.TotalSeconds()) & " Seconds"
            '
            objCpMgr.cp_set_SelectionToTitle(myDoc.Sections.Item(1))
            '
            '*** Reattach original template (Normal)
            'myDoc.AttachedTemplate = tmpl.FullName
            '***
            '
            objGlobals.glb_screen_update(True)

            MsgBox("The Report build is complete (" + strElapsedTime + ")")
            '
        Catch ex As Exception
            stpWatch.Stop()
            Interval = stpWatch.Elapsed()
            '
            '*** Reattach original template (Normal)
            'myDoc.AttachedTemplate = tmpl.FullName
            '***
            '
            MsgBox("There has been a problem building the report. If the problem persists please contact your 'Sys Admin'")
            objGlobals.glb_screen_update(True)
        End Try
        '
        Clipboard.Clear()
        '
finis:
        '
        Me.objGlobals.glb_cursors_setToNormal()
        '
        objGlobals.glb_screen_update(True)

    End Sub
    '
    ''' <summary>
    ''' This method will build a standard Short Report
    ''' </summary>
    Public Sub Rpt_build_newReport_Short()
        Dim objSectMgr As New cSectionMgr()
        Dim objFldsMgr As New cFieldsMgr()
        Dim objCpMgr As New cCoverPageMgr()
        Dim objMsgMgr As New cMessageManager()
        Dim objViewMgr As New cViewManager()
        Dim objStylesMgr As New cStylesManager()

        Dim sect As Word.Section
        Dim myDoc As Word.Document
        Dim strRptMode As String
        Dim placeBehind As Boolean
        Dim objParas As New cParas()
        Dim objBnrMgr As New cChptBanner()
        Dim para As Word.Paragraph
        Dim rng As Word.Range
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objTocMgr As New cTOCMgr()
        Dim objContactsMgr As New cContactsMgr()
        Dim objDivMgr As New cChptDivider()
        Dim rslt As Boolean
        Dim hf As Word.HeaderFooter
        Dim pgNums As Word.PageNumbers
        Dim strElapsedTime As String
        Dim stpWatch As System.Diagnostics.Stopwatch
        Dim startTime As Date
        Dim Interval As TimeSpan
        '
        Me.objGlobals.glb_screen_update(True)
        placeBehind = True
        myDoc = Me.objGlobals.glb_get_wrdActiveDoc()
        'lst = objBnrMgr.bnr_get_BannerSettings(objBnrMgr.sectType_body, True)
        '
        rslt = objMsgMgr.deleteAllMessage
        If Not rslt Then GoTo finis
        '
        Me.objGlobals.glb_cursors_setToWait()
        '
        stpWatch = System.Diagnostics.Stopwatch.StartNew()
        startTime = TimeOfDay()
        '
        strRptMode = Me.Rpt_Mode_Get()
        '
        'Mechanism to adjust styles without having to chnage the template
        Me.Rpt_styles_Upgrade_for_ReportType(myDoc, strRptMode)
        '
        Try
            'Build a document of empty standard sections
            objSectMgr.sct_delete_allSections()
            'sect = myDoc.Sections.Item(1)
            sect = myDoc.Sections.Last
            '
            objSectMgr.sct_reset_ToPortrait(sect)
            objParas.paras_insert_numParas(sect)
            '
            hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
            pgNums = hf.PageNumbers
            'pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleLowercaseRoman
            pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic
            pgNums.IncludeChapterNumber = True
            pgNums.HeadingLevelForChapter = 0
            pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen

            '
            'rng = sect.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'rng.Text = "First section"
            '
            'objViewMgr.vw_change_ColumnsAndRows(sect)
            objViewMgr.vw_change_toPageFitBestFit(myDoc)
            '
            Me.objGlobals.glb_screen_update(False)
            '
            For i = 1 To 10
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                sect = myDoc.Sections.Add(rng, WdSectionStart.wdSectionNewPage)
                'sect = myDoc.Sections.Last
                objParas.paras_insert_numParas(sect)
                objHFMgr.hf_hfs_linkUnlinkAll(sect, False)
            Next
            '
            sect = myDoc.Sections.First
            objCpMgr.cp_convert_ToCoverPage(sect)
            '
            objGlobals.glb_screen_update(False)
            '
            sect = myDoc.Sections.Item(2)
            objContactsMgr.contacts_convert_toFrontContacts(sect)
            '
            '
            'sect = myDoc.Sections.Item(myDoc.Sections.Last.Index - 7)
            'objParas.paras_add_textAndStyle(sect, "Glossary", "Heading (glossary)")
            'objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_glos))
            '
            '
            sect = myDoc.Sections.Item(4)
            objParas.paras_add_textAndStyle(sect, "ES Heading", "Heading 1 (ES)")
            objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_es))
            Me.chptBase_PageNumbering_Set(sect, True, 1, "es")

            '
            'hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
            'pgNums = hf.PageNumbers
            'pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleLowercaseRoman
            'pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleArabic
            'pgNums.IncludeChapterNumber = False
            'pgNums.RestartNumberingAtSection = True
            'pgNums.StartingNumber = 1
            'pgNums.HeadingLevelForChapter = 0
            'pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen

            chptBase_PageNumbering_Set(sect, True, 1, "es")

            sect = myDoc.Sections.Item(5)
            objParas.paras_add_textAndStyle(sect, "Heading 1", "Heading 1")
            '
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            rng.Move(WdUnits.wdParagraph, -1)
            para = objParas.paras_add_textAndStyle(rng, "Heading 1", "Heading 1")
            'para.Format.PageBreakBefore = False
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            'objParas.paras_insert_numParas(rng)
            '
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            rng.Move(WdUnits.wdParagraph, -1)
            para = objParas.paras_add_textAndStyle(rng, "Heading 1", "Heading 1")
            'para.Format.PageBreakBefore = False
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            'objParas.paras_insert_numParas(rng)
            objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_chpt_body))
            chptBase_PageNumbering_Set(sect, True, 1, "std")
            '
            '
            sect = myDoc.Sections.Item(6)
            objDivMgr.div_convert_toDivider(sect, objBnrMgr.tag_divAP, "Appendices")
            Me.chptBase_PageNumbering_Set(sect, False, 1, "std")

            'objDivMgr.div_convert_toDivider(sect,)
            'objBnrMgr.bnr_insert_BannerBase(sect.Range, False, strRptMode, objBnrMgr.bnr_get_BannerSettings(objBnrMgr.tag_divAP, False))
            'chptBase_PageNumbering_Set(sect, False, 1, "div")
            '
            sect = myDoc.Sections.Item(7)
            If objGlobals._glb_doApp_as_HeadingAP Then
                objParas.paras_add_textAndStyle(sect, "Appendix Heading 1", "Heading 1 (AP)")
                chptBase_PageNumbering_Set(sect, True, 1, "std")
            Else
                objParas.paras_add_textAndStyle(sect, "Appendix Heading 1", "Heading 6")
                chptBase_PageNumbering_Set(sect, True, 1, "ap")
            End If
            objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_chpt_AP))
            '
            sect = myDoc.Sections.Item(8)
            If objGlobals._glb_doApp_as_HeadingAP Then
                objParas.paras_add_textAndStyle(sect, "Appendix Heading 1", "Heading 1 (AP)")
                chptBase_PageNumbering_Set(sect, True, 1, "std")
            Else
                objParas.paras_add_textAndStyle(sect, "Appendix Heading 1", "Heading 6")
                chptBase_PageNumbering_Set(sect, True, 1, "ap")
            End If
            objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_chpt_AP))
            '
            'sect = objContactsMgr.contacts_insert_BackPage()
            objContactsMgr.contacts_convert_toBackContacts(myDoc.Sections.Item(9))
            'Do primary and first page
            'objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_cont_Back))
            'objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_cont_Back), "firstPage")
            '
            '
            sect = myDoc.Sections.Item(3)
            objTocMgr.toc_convert_toTOC(sect)
            objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_toc))
            objTocMgr.toc_replace_TOCField(sect, "aac_TOC_Levels02")
            '
finis0:
            stpWatch.Stop()
            Interval = stpWatch.Elapsed()
            '
            'endTime = TimeOfDay()
            'Interval = endTime - startTime
            strElapsedTime = Int(Interval.TotalSeconds()) & " Seconds"
            '
            objCpMgr.cp_set_SelectionToTitle(myDoc.Sections.Item(1))
            objGlobals.glb_screen_update(True)

            MsgBox("The Report build is complete (" + strElapsedTime + ")")
            '
        Catch ex As Exception
            MsgBox("There has been a problem building the report. If the problem persists please contact your 'Sys Admin'")
            objGlobals.glb_screen_update(True)
        End Try
        '
finis:
        '
        Clipboard.Clear()

        Me.objGlobals.glb_cursors_setToNormal()
        '
        objGlobals.glb_screen_update(True)

    End Sub
    '
    ''' <summary>
    ''' This method will assume that the first section is a letter. It
    ''' will then delete all sections after the first section..
    ''' </summary>
    Public Sub Rpt_delete_ReportSections()
        Dim objSectMgr As New cSectionMgr()
        Dim objGlobals As New cGlobals
        Dim sect, sectSrc, sectDest As Section
        Dim rng As Range
        '

        '
        For Each sect In objGlobals.glb_get_wrdActiveDoc().Sections
            If sect.Index <> 1 Then sect.Range.Delete()
        Next
        '
        sectSrc = objGlobals.glb_get_wrdActiveDoc.Sections.Item(1)
        sectDest = objGlobals.glb_get_wrdActiveDoc.Sections.Last

        objSectMgr.cloneSection(sectSrc, sectDest)
        '
        rng = sectDest.Range
        rng.Delete()
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'rng = Globals.ThisDocument.Application.Selection.Range
        rng.MoveStart(WdUnits.wdParagraph, -2)
        rng.Delete()
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        '
    End Sub
    '
#Region "Report Mode CHnage"
    '
    Public Sub Rpt_ModeChange_ToLong()
        Dim rng As Word.Range
        Dim sect As Word.Section
        Dim objChptBnr As New cChptBanner()
        '
        'oldRange = Globals.ThisDocument.Application.Selection.Range
        '
        'Set the mode to long and make certain that the styles are back to normal.
        'We do this just in cose someone was toggling...Note that "Me.Rpt_Mode_SetAs_Std()"
        'will call "Me.Rpt_Styles_ChangeStylesForReportMode(Me.modeLong)" which chnages
        'the styles to suit the standard report.... Note that I am trying to avoid a global
        'styles update
        Me.Rpt_Mode_SetAs_Std()
        '
        For Each sect In Me.objGlobals.glb_get_wrdActiveDoc.Sections
            rng = sect.Range
            For Each tbl In rng.Tables
                objChptBnr.bnr_resize_BannerBase(tbl, "toLarge")                        '
            Next
            '
        Next

    End Sub
    '
    '
    Public Sub Rpt_ModeChange_ToShort()
        Dim rng As Word.Range
        Dim sect As Word.Section
        Dim objChptBnr As New cChptBanner()
        '
        Me.Rpt_Mode_SetAsShort()

        '
        'Me.Rpt_StylesChangeSequence_ForReportMode(Me.modeShort)... done in Me.Rpt_Mode_SetAsShort()
        '
        'By calling the setHeading1FontSizeForLongReport method we set H1 so that
        'the document mode is changed to Long form
        '
        For Each sect In Me.objGlobals.glb_get_wrdActiveDoc.Sections
            rng = sect.Range
            For Each tbl In rng.Tables
                objChptBnr.bnr_resize_BannerBase(tbl, "toSmall")                        '
            Next
            '
        Next

    End Sub
    '
#End Region
    '
#Region "Report Mode"
    '
    ''' <summary>
    ''' This method will return the Report Mode, that is;
    ''' -   longReport
    ''' -   shortReport
    ''' -   longReport_Lnd
    ''' 
    ''' </summary>
    ''' <returns></returns>
    Public Function Rpt_Mode_Get() As String
        Dim objGlobals As New cGlobals()
        Dim myDoc As Word.Document
        Dim props As DocumentProperties
        Dim prop As DocumentProperty
        Dim strReportMode As String
        '
        myDoc = objGlobals.glb_get_wrdActiveDoc()
        props = myDoc.CustomDocumentProperties
        strReportMode = ""

        Try
            prop = props.Item(Me.modePropertyNameForReport)
            strReportMode = CStr(prop.Value)
        Catch ex As Exception
            'We are here because there is no property. Since the Long report is the default we'll set
            'the property to Long report
            strReportMode = Me.rpt_isPrt
            Call props.Add(Me.modePropertyNameForReport, False, MsoDocProperties.msoPropertyTypeString, strReportMode)

        End Try
        '
        '*** Test Setting
        'strReportMode = Me.modeLongLandscape
        '
        Return strReportMode
    End Function
    '
    '
    Public Function Rpt_Mode_SetAsLandScape() As String
        Dim objGlobals As New cGlobals()
        Dim myDoc As Word.Document
        Dim props As DocumentProperties
        Dim prop As DocumentProperty
        Dim strRptMode As String
        '
        myDoc = objGlobals.glb_get_wrdActiveDoc()
        props = myDoc.CustomDocumentProperties
        strRptMode = Me.rpt_isLnd
        '
        Try
            prop = props.Item(Me.modePropertyNameForReport)
            prop.Value = Me.rpt_isLnd
            '
            'Set banner styles to those needed for the long report
            Me.Rpt_Styles_ChangeStylesForReportMode(Me.rpt_isLnd)
            '
        Catch ex As Exception
            Call props.Add(Me.modePropertyNameForReport, False, MsoDocProperties.msoPropertyTypeString, Me.rpt_isLnd)
        End Try
        '
        Return strRptMode
        '
    End Function
    '
    Public Function Rpt_Mode_SetAs_Std() As String
        Dim objGlobals As New cGlobals()
        Dim myDoc As Word.Document
        Dim props As DocumentProperties
        Dim prop As DocumentProperty
        Dim strRptMode As String
        '
        myDoc = objGlobals.glb_get_wrdActiveDoc()
        props = myDoc.CustomDocumentProperties
        strRptMode = Me.rpt_isPrt
        '
        Try
            prop = props.Item(Me.modePropertyNameForReport)
            prop.Value = Me.rpt_isPrt
            '
            'Set banner styles to those needed for the long report
            Call Me.Rpt_Styles_ChangeStylesForReportMode(Me.rpt_isPrt)
            '
        Catch ex As Exception
            Call props.Add(Me.modePropertyNameForReport, False, MsoDocProperties.msoPropertyTypeString, Me.rpt_isPrt)
        End Try
        '
        Return strRptMode
    End Function
    '
    Public Function Rpt_Mode_SetAsShort() As String
        Dim objGlobals As New cGlobals()
        Dim myDoc As Word.Document
        Dim props As DocumentProperties
        Dim prop As DocumentProperty
        Dim strRptMode As String
        '
        myDoc = objGlobals.glb_get_wrdActiveDoc()
        props = myDoc.CustomDocumentProperties
        '
        Try
            'Set the stored mode indicator to short
            prop = props.Item(Me.modePropertyNameForReport)
            prop.Value = Me.modeShort
            '
            'Set banner styles to those needed for the Short report
            Call Me.Rpt_Styles_ChangeStylesForReportMode(Me.modeShort)
            strRptMode = Me.modeShort
            '
        Catch ex As Exception
            Call props.Add(Me.modePropertyNameForReport, False, MsoDocProperties.msoPropertyTypeString, Me.modeShort)
            strRptMode = Me.modeShort
        End Try
        '
        Return strRptMode
finis:
    End Function
    '
    '
    Public Function Rpt_Mode_SetAsAABrief() As String
        Dim objGlobals As New cGlobals()
        Dim myDoc As Word.Document
        Dim props As DocumentProperties
        Dim prop As DocumentProperty
        Dim strRptMode As String
        '
        myDoc = objGlobals.glb_get_wrdActiveDoc()
        props = myDoc.CustomDocumentProperties
        '
        Try
            'Set the stored mode indicator to short
            prop = props.Item(Me.modePropertyNameForReport)
            prop.Value = Me.rpt_isBrief
            '
            'Set banner styles to those needed for the Short report
            Call Me.Rpt_Styles_ChangeStylesForReportMode(Me.rpt_isBrief)
            strRptMode = Me.rpt_isBrief
            '
        Catch ex As Exception
            Call props.Add(Me.modePropertyNameForReport, False, MsoDocProperties.msoPropertyTypeString, Me.rpt_isBrief)
            strRptMode = Me.rpt_isBrief
        End Try
        '
        Return strRptMode
finis:
    End Function
    '

    '
    '
    ''' <summary>
    ''' You call this method to refresh the document styles from the template. The default styleset
    ''' is for the standard Portrait Long Report.. Both the Short and Landscape reports use modified
    ''' versions of that styleset. So, once we update the standard styles we then need to put back any
    ''' chnages made for the Short and Landscape reports
    ''' </summary>
    Public Sub Rpt_Styles_resetStyles_fromTemplate(Optional attachAATemplate As Boolean = False)
        Dim objGlobals As New cGlobals()
        Dim oldTmpl As Word.Template
        Dim strTemplateFullName As String
        '
        Try
            oldTmpl = objGlobals.glb_get_wrdActiveDoc.AttachedTemplate
            '
            strTemplateFullName = objGlobals.glb_getTmpl_FullName()
            objGlobals.glb_get_wrdActiveDoc.AttachedTemplate = strTemplateFullName
            objGlobals.glb_get_wrdActiveDoc.UpdateStyles()
            '
            'Remember to adjust for any modified styles (e.g. for Brief)
            '
            '
            '
            If attachAATemplate Then
                'Leave it be
            Else
                If oldTmpl.FullName Like "*Normal*" Then
                    objGlobals.glb_get_wrdActiveDoc.AttachedTemplate = "Normal"
                Else
                    objGlobals.glb_get_wrdActiveDoc.AttachedTemplate = oldTmpl.FullName
                End If
                '
            End If

        Catch ex As Exception

        End Try
        '
    End Sub
    '
    '
    '
    ''' <summary>
    ''' This method holds the style changes for each of the report modes.. Note that the
    ''' Long Report chnage is only ever used to reverse the chnages made for a Short Report.
    ''' Otherwise they are never used
    ''' </summary>
    ''' <param name="strMode"></param>
    Public Sub Rpt_Styles_ChangeStylesForReportMode(strMode As String)
        Dim objTOCMgr As New cTOCMgr()
        'Dim styleChapterNumber As Word.Style
        'Dim styleTagChapter As Word.Style
        'Dim styleTagChapterApp As Word.Style
        'Dim stylePartNumber As Style
        'Dim styleAppendixChapterNumber As Style
        '
        'styleChapterNumber = Globals.ThisDocument.Application.ActiveDocument.Styles("Heading (Chapter)")
        'styleTagChapter = Globals.ThisDocument.Application.ActiveDocument.Styles("tag_chapterBanner")
        'styleTagChapterApp = Globals.ThisDocument.Application.ActiveDocument.Styles("tag_appendixChapter")

        'stylePartNumber = ActiveDocument.Styles("Part - Number")
        'styleAppendixChapterNumber = ActiveDocument.Styles("Heading (Appendix)")
        '
        '.ParagraphFormat.LineSpacingRule = wdLineSpaceSingle
        '
        'objTOCMgr.TOC_Styles_AdjustForReportMode()
        objTOCMgr.TOC_Styles_AdjustForReportMode(strMode)
        '
        Select Case strMode
            Case Me.rpt_isPrt
                'styleChapterNumber.Font.Size = 160
                'styleChapterNumber.Font.Size = 80
                'styleChapterNumber.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
                'styleChapterNumber.ParagraphFormat.LineSpacing = 98.0
                'styleChapterNumber.Font.Position = 0.0
                '
                'styleTagChapter.ParagraphFormat.PageBreakBefore = True
                'styleTagChapterApp.ParagraphFormat.PageBreakBefore = True
                '
                'objTOCMgr.TOC_Styles_AdjustForReportMode()
                '            '
            Case Me.modeShort
                'styleChapterNumber.Font.Size = 55.0
                'styleChapterNumber.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
                'styleChapterNumber.ParagraphFormat.LineSpacing = 80.0
                'styleChapterNumber.Font.Position = 3.0
                '
                'styleTagChapter.ParagraphFormat.PageBreakBefore = False
                'styleTagChapterApp.ParagraphFormat.PageBreakBefore = False
                '
                'objTOCMgr.TOC_Styles_AdjustForReportMode()
                '
            Case Me.rpt_isLnd
                'objTOCMgr.TOC_Styles_AdjustForReportMode()
                '
            Case Else

        End Select
        '
        Exit Sub
finis:
    End Sub
    '
#End Region
End Class

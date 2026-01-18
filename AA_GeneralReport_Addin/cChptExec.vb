Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cChptExec
    Inherits cChptBase
    Public Sub New()
        MyBase.New()
        '
        Me.strTagStyleName = "tag_execBanner"
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will reurn true if the section (sect) has a banner table and if that
    ''' Banner table has the Appendix tag style in hte first cell
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function es_is_execSummary(ByRef sect As Word.Section)
        Dim objTagsMgr As New cTagsMgr()
        '
        Return objTagsMgr.tags_is_ExecutiveSummary(sect)
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will return true if the document has an Executive Summary
    ''' </summary>
    ''' <returns></returns>
    Public Function es_has_ExecSummary() As Boolean
        Dim rslt As Boolean
        Dim objsectMgr As New cSectionMgr()
        Dim objBnrMgr As New cChptBanner()
        Dim sect As Word.Section

        '
        sect = Nothing
        rslt = False
        rslt = objsectMgr.sct_has_strTag(objGlobals.glb_get_wrdActiveDoc, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_es), sect)
        '
        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method will return nothing if the document does not have an Executive Summary. If it does, then it
    ''' will return the Executive Summary section
    ''' </summary>
    ''' <returns></returns>
    Public Function es_get_ExecSummary() As Word.Section
        Dim objsectMgr As New cSectionMgr()
        Dim objBnrMgr As New cChptBanner()
        Dim sect As Word.Section

        '
        sect = Nothing
        objsectMgr.sct_has_strTag(objGlobals.glb_get_wrdActiveDoc, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_es), sect)
        '
        Return sect
        '
    End Function


    '
    Public Function es_insert_section(placeBehind As Boolean, ByRef sect As Word.Section, strRptMode As String) As Word.Table
        'Dim strRptMode As String
        Dim objRptMgr As New cReport()
        Dim objBnrMgr As New cChptBanner()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objStylesMgr As New cStylesManager()
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim strOrientation As String
        Dim lst As New Collection()
        '
        'strRptMode = objRptMgr.Rpt_Mode_Get()
        tbl = Nothing
        '
        strOrientation = "prt"
        '
        Select Case strRptMode
            Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                If strRptMode = objRptMgr.rpt_isLnd Then strOrientation = "lnd"
                '
                lst = objBnrMgr.bnr_get_BannerSettings(objBnrMgr.tag_es)
                tbl = MyBase.chpt_Insert_Std(placeBehind, sect, lst, strRptMode, strOrientation, 8)
                '
                'Must override the orignal setup defined in cChptBnr.bnr_get_BannerSettings which
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Paragraphs.Item(1).Range.Delete()
                rng = objStylesMgr.styles_insert_StartupText_ReportES(rng)
                '
                objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_es))

                Me.chptBase_PageNumbering_Set(sect, True, 1, "es")

            Case objRptMgr.rpt_isBrief
                rng = objGlobals.glb_get_wrdSelRng()
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng = objStylesMgr.styles_insert_StartupText_ReportES(rng)
                '
        End Select
        '
        'strOrientation = "prt"
        'If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then strOrientation = "lnd"
        'If strRptMode = objRptMgr.modeLongLandscape Then strOrientation = "lnd"
        '
        '
        '
        Return tbl
        '
    End Function
    '
    Public Function es_convert_ToES(ByRef sect As Word.Section, strRptMode As String) As Word.Section
        Dim objBnrMgr As New cChptBanner()
        Dim pgNums As Word.PageNumbers
        '
        objBnrMgr.bnr_insert_BannerBase(sect.Range, True, strRptMode, objBnrMgr.bnr_get_BannerSettings(objBnrMgr.tag_es, True))
        pgNums = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary).PageNumbers
        pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleLowercaseRoman
        pgNums.RestartNumberingAtSection = True
        pgNums.StartingNumber = 1
        '
        Return sect
    End Function
    '
    Public Function es_insert_LndBounded() As Word.Section
        Dim objSectMgr As New cSectionMgr()
        Dim objRptMgr As New cReport()
        Dim rng As Word.Range
        Dim sect As Word.Section
        Dim strRptMode As String
        '
        strRptMode = objRptMgr.Rpt_Mode_Get()
        sect = Nothing
        '
        rng = objGlobals.glb_get_wrdSelRng()
        '
        Select Case strRptMode
            Case objRptMgr.rpt_isPrt, objRptMgr.modeShort
                sect = objSectMgr.sct_insert_SectionBounded(rng, "std_Lnd")
            Case objRptMgr.rpt_isLnd
                sect = objSectMgr.sct_insert_SectionBounded(rng, "lnd_Lnd")
        End Select
        '
        Return sect
        '
    End Function
    '
    '
    Public Overridable Function es_pageNumbering_Set(ByRef sect As Word.Section, restartNumbering As Boolean, Optional startAt As Integer = 1) As Boolean
        Dim pgNums As PageNumbers
        Dim objToolsMgr As New cTools()
        Dim hf As Word.HeaderFooter
        Dim myApp As Word.Application
        Dim rslt As Boolean
        '
        myApp = sect.Range.Application
        rslt = False
        '
        hf = sect.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        pgNums = hf.PageNumbers
        pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleLowercaseRoman
        Try
            pgNums.RestartNumberingAtSection = restartNumbering
            pgNums.StartingNumber = 1

            pgNums.IncludeChapterNumber = False
            pgNums.HeadingLevelForChapter = 0
            pgNums.ChapterPageSeparator = WdSeparatorType.wdSeparatorHyphen
            '
            rslt = True
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
        '
    End Function


End Class

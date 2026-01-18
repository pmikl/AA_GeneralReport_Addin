Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cChptApp
    Inherits cChptBase
    Public Sub New()
        MyBase.New()
        '
        Dim objTagsMgr As New cTagsMgr()
        '
        Me.strTagStyleName = objTagsMgr.bnr_get_tagStyles(objTagsMgr.tag_chpt_AP)
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
    Public Function app_is_AppPage(ByRef sect As Word.Section)
        Dim objTagsMgr As New cTagsMgr()
        '
        Return objTagsMgr.tags_is_Chapter_Appendix(sect)
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will reurn true if the section holding the current selection uses Appendix Style PageNumbering. If it
    ''' does, then the current section is an Appendix Section
    ''' </summary>
    ''' <returns></returns>
    Public Function app_is_AppSection() As Boolean
        Dim sect As Word.Section
        Dim pgNums As Word.PageNumbers
        Dim rslt As Boolean
        '
        rslt = False
        sect = objGlobals.glb_get_wrdSect()
        pgNums = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary).PageNumbers
        '
        If pgNums.HeadingLevelForChapter = 8 And pgNums.IncludeChapterNumber Then rslt = True
        '
        Return True
        '
    End Function
    '

    '
    ''' <summary>
    ''' This method will insert a standard Appendix chapter either in front of, or behind the section that
    ''' contains the current selection. It returns the banner table as a Word.Table object
    ''' </summary>
    ''' <param name="placeBehind"></param>
    ''' <returns></returns>
    Public Function app_insert_App(placeBehind As Boolean, ByRef sect As Word.Section, strRptMode As String) As Word.Range
        Dim objRptMgr As New cReport()
        Dim objBnrMgr As New cChptBanner()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objStylesMgr As New cStylesManager()
        Dim objSectMgr As New cSectionMgr()
        Dim tbl As Word.Table
        Dim myDoc As Word.Document
        Dim rng As Word.Range
        Dim lst As New Collection()
        Dim strOrientation As String
        '
        lst = objBnrMgr.bnr_get_BannerSettings(objBnrMgr.tag_chpt_AP)
        myDoc = sect.Range.Document
        tbl = Nothing
        rng = Nothing
        '
        strOrientation = "prt"
        If strRptMode = objRptMgr.rpt_isLnd Then strOrientation = "lnd"
        '
        Select Case strRptMode
            Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                lst = objBnrMgr.bnr_get_BannerSettings(objBnrMgr.tag_chpt_AP)
                'tbl = MyBase.chpt_Insert_Std(placeBehind, sect, lst, strRptMode, "prt", 8)
                tbl = MyBase.chpt_Insert_Std(placeBehind, sect, lst, strRptMode, strOrientation, 7)
                '
                'Must override the orignal setup defined in cChptBnr.bnr_get_BannerSettings which
                rng = objSectMgr.sct_delete_allSectionContents(sect, 6)
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Paragraphs.Item(1).Range.Delete()
                rng = objStylesMgr.styles_insert_StartupText_ReportAP(rng)
                '
                objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_chpt_AP))

                'Me.chptBase_PageNumbering_Set(tbl.Range.Sections.Item(1), True, 1, "ap")
                Me.chptBase_PageNumbering_Set(sect, True, 1, "ap")
                '
            Case objRptMgr.rpt_isBrief
                rng = objGlobals.glb_get_wrdSelRng()
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng = objStylesMgr.styles_insert_StartupText_ReportAP(rng)

        End Select
        '
        '
        Return rng
        '
    End Function
    '
    ''' <summary>
    ''' This method will insert a bounded Landscape section at the current selection
    ''' </summary>
    ''' <returns></returns>
    Public Function app_insert_landscapeSection() As Word.Section
        Dim objSectMgr As New cSectionMgr()
        Dim sect As Word.Section
        Dim rng As Word.Range
        '
        rng = objGlobals.glb_get_wrdSelRng()
        '
        '
        sect = objSectMgr.sct_insert_SectionBounded(rng, "std_Lnd")
        '
        Return sect
        '
    End Function


End Class

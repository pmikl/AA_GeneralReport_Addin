Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cChptBody
    Inherits cChptBase

    Public Sub New()
        '
        MyBase.New()
        '
        Dim objBanner As New cChptBanner()
        Me.strTagStyleName = objBanner.bnr_get_tagStyles(objBanner.tag_chpt_body)
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
    Public Function chpt_is_BodyPage(ByRef sect As Word.Section)
        Dim objTagsMgr As New cTagsMgr()
        '
        Return objTagsMgr.tags_is_Chapter_Body(sect)
        '
    End Function
    '
    ''' <summary>
    ''' This method will insert a standard chapter either in front of, or behind the section that
    ''' contains the current selection. It returns the banner table as a Word.Table object and the varaiable
    ''' sect now contains the new Chapter Section. 
    ''' </summary>
    ''' <param name="placeBehind"></param>
    ''' <returns></returns>
    Public Function chpt_insert_Body(placeBehind As Boolean, ByRef sect As Word.Section, strRptMode As String) As Word.Range
        Dim objRptMgr As New cReport()
        Dim objBnrMgr As New cChptBanner()
        Dim objFldsMgr As New cFieldsMgr()
        Dim objPgNumMgr As New cPageNumberMgr()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objParas As New cParas()
        Dim objStylesMgr As New cStylesManager()
        Dim objSectMgr As New cSectionMgr()
        Dim tbl As Word.Table
        Dim lst As New Collection()
        Dim strPgNumFormat As String
        Dim numRestart As Boolean
        Dim rng As Word.Range
        '
        'strRptMode = objRptMgr.Rpt_Mode_Get()
        tbl = Nothing
        rng = Nothing
        '
        numRestart = False
        strPgNumFormat = objPgNumMgr.pgNum_get_numFormat_ForDoc()
        If strPgNumFormat = objPgNumMgr.pgNum_pgNumype_2part Then numRestart = True
        '
        Select Case strRptMode
            Case objRptMgr.rpt_isPrt
                lst = objBnrMgr.bnr_get_BannerSettings(objBnrMgr.tag_chpt_body, True)
                tbl = MyBase.chpt_Insert_Std(placeBehind, sect, lst, strRptMode, "prt", 5)
                '
                'Override the single heading setup and insert the 'StartupText'
                rng = objSectMgr.sct_delete_allSectionContents(sect, 6)
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Paragraphs.Item(1).Range.Delete()
                rng = objStylesMgr.styles_insert_StartupText_ReportBody(rng)
                '
                objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_chpt_body))
                '
                Me.chptBase_PageNumbering_Set(sect, numRestart, 1, strPgNumFormat)
                '
                '
            Case objRptMgr.modeShort, objRptMgr.rpt_isBrief
                rng = objGlobals.glb_get_wrdSelRng()
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng = objStylesMgr.styles_insert_StartupText_ReportBody(rng)
                '
                Me.chptBase_PageNumbering_Set(rng.Sections.Item(1), numRestart, 1, strPgNumFormat)
                '
            Case objRptMgr.rpt_isLnd

                lst = objBnrMgr.bnr_get_BannerSettings(objBnrMgr.tag_chpt_body, True)
                tbl = MyBase.chpt_Insert_Std(placeBehind, sect, lst, strRptMode, "lnd", 5)
                '
                rng = sect.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Paragraphs.Item(1).Range.Delete()
                rng = objStylesMgr.styles_insert_StartupText_ReportBody(rng)
                '
                objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_chpt_body))
                '
                Me.chptBase_PageNumbering_Set(sect, numRestart, 1, strPgNumFormat)
                '
        End Select
        '
        'objFldsMgr.updateSequenceNumbers_Chapters()
        '
        Me.tbl_Banner = tbl
        '
        Return rng
        '
    End Function
    '
    Private Sub xchpt_insert_PrtStd()

    End Sub
    '
    Private Sub xchpt_insert_Short()

    End Sub
    '
    Private Sub xchpt_insert_LndStd()

    End Sub

End Class

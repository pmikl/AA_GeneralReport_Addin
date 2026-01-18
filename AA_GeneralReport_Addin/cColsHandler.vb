Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cColsHandler
    Public objGlobals As cGlobals
    '
    Public Sub New()
        Me.objGlobals = New cGlobals()
    End Sub
    '
    '
    ''' <summary>
    ''' This method will setup the column layout for the specified section. It does so in
    ''' accordance with strMode ("4_columns", "3_columns", "2_columns". "2_columns_leftWide", "2_columns_rightWide"
    ''' and "1_column"
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="strLayout"></param>
    Public Sub cols_setup_columnStructure(ByRef sect As Word.Section, strLayout As String)
        Dim pageWidth, drColWidth, colSpacing As Single
        Dim rng As Word.Range
        Dim objTools As New cTools()
        '
        pageWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
        colSpacing = objTools.tools_math_MillimetersToPoints(10.0)                                              '10 mm spacing
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
                objGlobals.glb_get_wrdSel().PageSetup.SectionStart = WdSectionStart.wdSectionNewPage
                'Globals.ThisDocument.Application.Selection.PageSetup.SectionStart = WdSectionStart.wdSectionNewPage

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
                objGlobals.glb_get_wrdSel().PageSetup.SectionStart = WdSectionStart.wdSectionNewPage
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
                objGlobals.glb_get_wrdSel().PageSetup.SectionStart = WdSectionStart.wdSectionNewPage
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
                objGlobals.glb_get_wrdSel().PageSetup.SectionStart = WdSectionStart.wdSectionNewPage
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

                objGlobals.glb_get_wrdSel().PageSetup.SectionStart = WdSectionStart.wdSectionNewPage

            Case "1_columns"
                sect.PageSetup.TextColumns.SetCount(1)

        End Select
    End Sub
    '
    '
    '
    'This function checks for illegal insertion actions. It first makes certain that
    'the user is not trying to insert anything is the Cover Page, Contacts Page (Front and Back),
    'Table of Contents.. If so, we get an error message and it returns False. It
    'will then check to see if we are trying to insert in a Table. If so, then we get
    'an error message and it will return false
    Public Overridable Function cols_isOK_ToChangeColumns(ByRef sect As Section) As Boolean
        Dim objCpMgr As New cCoverPageMgr()
        Dim objSectMgr As New cSectionMgr()
        Dim objRptMgr As New cReport()
        Dim objTOCMgr As New cTOCMgr()
        Dim objContactsPageMgr As New cContactsMgr()
        Dim objBnrMgr As New cChptBanner()
        Dim objDivMgr As New cChptDivider()
        Dim strErrorMsg As String
        Dim rslt As Boolean
        '
        strErrorMsg = ""
        rslt = True
        '
        'strRptMode = objRptMgr.Rpt_Mode_Get()

        'If strRptMode <> objRptMgr.modeLongLandscape Then
        'strErrorMsg = "This Function can only be used in a Landscape Report"
        'GoTo finis
        'End If
        '

        '
        If objCpMgr.cp_Bool_IsCoverPage(sect) Then strErrorMsg = "Changing the column structure " + vbCrLf + "of a Cover Page is not supported"
        If objContactsPageMgr.is_ContactsPage_Front(sect) Then strErrorMsg = "Changing the column structure" + vbCrLf + "of the front contacts page is not supported"
        If objTOCMgr.toc_is_TOCSection(sect) Then strErrorMsg = "Changing the column structure" + vbCrLf + "of the TOC is not supported"
        If objDivMgr.is_divider_Chpt(sect) Then strErrorMsg = "Changing the column structure" + vbCrLf + "of a divider is not supported"
        If objDivMgr.is_divider_App(sect) Then strErrorMsg = "Changing the column structure" + vbCrLf + "of a the appendix divider is not supported"

        'If objSectMgr.sct_Get_SectionTag(sect) = "tag_*" Then Then strErrorMsg = "Changing the column structure of this section is not supported"
        'If objSectMgr.sct_Get_SectionTag(sect) Like objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_div) Then strErrorMsg = "Changing the column structure of a divider is not supported"
        'If objSectMgr.sct_Get_SectionTag(sect) Like objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_divAP) Then strErrorMsg = "Changing the column structure of the appendix divider is not supported"

        'If objChpt.is_Chapter(sect) Then strErrorMsg = "Changing the column structure of a chapter banner page is not supported"
        'If objChpt.is_Appendix_Chapter(sect) Then strErrorMsg = "Changing the column structure of an appendix chapter banner page is not supported"

        If objContactsPageMgr.is_ContactsPage_Back(sect) Then strErrorMsg = "Changing the column structure" + vbCrLf + "of the back contacts is not supported"
        '
        'Check to see if in table
        '
        If strErrorMsg <> "" Then
            rslt = False
            MsgBox(strErrorMsg)
        End If
        '
        '
        Return rslt

    End Function
    '

End Class

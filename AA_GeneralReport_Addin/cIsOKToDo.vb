Public Class cIsOKToDo
    Inherits cGlobals
    '
    Public objLetterMgr As cStationeryLetter
    Public objMemoMgr As cStationeryMemo
    '
    Public objCpMgr As cCoverPageMgr
    Public objChptExec As cChptExec
    Public objSectMgr As cSectionMgr
    Public objChpt As cChptBody
    Public objDivMgr As cChptDivider
    '
    Public objChptBanner As New cChptBanner()
    '
    Public objContactsPgMgr As cContactsMgr
    Public objTOCMgr As cTOCMgr
    '
    Public _isOK As String

    Public Sub New()
        MyBase.New()
        '
        Me.objLetterMgr = New cStationeryLetter()
        Me.objMemoMgr = New cStationeryMemo()
        '
        Me.objCpMgr = New cCoverPageMgr()
        Me.objChptExec = New cChptExec()
        Me.objSectMgr = New cSectionMgr()
        Me.objChpt = New cChptBody()
        Me.objDivMgr = New cChptDivider()
        Me.objContactsPgMgr = New cContactsMgr()
        Me.objTOCMgr = New cTOCMgr()

        '
        Me._isOK = "ok"
    End Sub
    '
    Public Function isOKto_Insert_ESummary(ByRef sect As Word.Section) As String
        Dim strErrorMsg As String = ""
        '
        If objChptExec.es_has_ExecSummary() Then
            'strErrorMsg = "This document already has an Executive Summary." + vbCrLf + vbCrLf + "If you want to insert a new one, you'll have to delete the old one."
            strErrorMsg = ""
        End If


        'If Me.objCpMgr.cp_Bool_IsCoverPage(sect) Then strErrorMsg = "A chapter cannot be inserted in front of a Cover Page"
        'If Me.objContactsPgMgr.is_ContactsPage_Front(sect) Then strErrorMsg = "A chapter cannot be inserted in front of the front Contacts Page"
        'If Me.objTOCMgr.toc_is_TOCSection(sect) Then strErrorMsg = "A chapter cannot be inserted in front of a TOC"
        '
        Return strErrorMsg
    End Function
    '
    '
    ''' <summary>
    ''' Placeholder for a test that may never be implemented
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function isOKto_Insert_ChptAPinFront(ByRef sect As Word.Section, strRptMode As String) As String
        Dim objRptMgr As New cReport()
        Dim objChptAP As New cChptApp()
        Dim objContactsPgMgr As New cContactsMgr()
        Dim myDoc As Word.Document
        Dim sectLast As Word.Section
        Dim strErrorMsg As String = ""
        '
        myDoc = sect.Range.Document
        sectLast = myDoc.Sections.Last
        '
        Select Case strRptMode
            Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                If objChptAP.app_is_AppPage(sect) Or objContactsPgMgr.is_ContactsPage_Back(sect) Then
                    strErrorMsg = ""
                Else
                    strErrorMsg = "A new appendix section can only be inserted" + vbCrLf + "in front of an existing appendix section or" + vbCrLf + "in front of the back contacts page."
                End If
                'If sect.Index = sectLast.Index Then strErrorMsg = "Insertion of a chapter behind the last section is not supported"

            Case objRptMgr.rpt_isBrief
                '
                If objRptMgr.sct_Sel_IsIn_TableOnly() Then strErrorMsg = "Your current cursor position needs to be outside of a table"
                If sect.Index = sectLast.Index Then strErrorMsg = "Insertion of a chapter behind the last section is not supported"


                'If Not Me.objChpt.chpt_is_BodyPage(sect) Then strErrorMsg = "A new chapter section can only be inserted" + vbCrLf + "behind an existing chapter section."
                'If Me.objCpMgr.cp_Bool_IsCoverPage(sect) Then strErrorMsg = "A chapter cannot be inserted behind a Cover Page"
                'If Me.objContactsPgMgr.is_ContactsPage_Front(sect) Then strErrorMsg = "A chapter cannot be inserted behind the front Contacts Page"
                'If Me.objDivMgr.is_divider_App(sect) Then strErrorMsg = "A chapter should not be inserted behind the Appendix Divider"

                'If Me.objContactsPgMgr.is_ContactsPage_Back(sect) Then strErrorMsg = "A chapter cannot be inserted behind the back Contacts Page"

        End Select
        '
        Return strErrorMsg
    End Function
    '
    ''' <summary>
    ''' Placeholder for a test that may never be implemented
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function isOKto_Insert_ChptAPBehind(ByRef sect As Word.Section, strRptMode As String) As String
        Dim objRptMgr As New cReport()
        Dim objChptAP As New cChptApp()
        Dim objDivider As New cChptDivider()
        Dim myDoc As Word.Document
        Dim sectLast As Word.Section
        Dim strErrorMsg As String = ""
        '
        myDoc = sect.Range.Document
        sectLast = myDoc.Sections.Last
        '
        Select Case strRptMode
            Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                If objChptAP.app_is_AppPage(sect) Or objDivider.is_divider_App(sect) Then
                    strErrorMsg = ""
                Else
                    strErrorMsg = "A new appendix section can only be inserted" + vbCrLf + "behind the appendix divider, or behind" + vbCrLf + "an existing appendix section."
                End If
                'If sect.Index = sectLast.Index Then strErrorMsg = "Insertion of a chapter behind the last section is not supported"

            Case objRptMgr.rpt_isBrief
                '
                If objRptMgr.sct_Sel_IsIn_TableOnly() Then strErrorMsg = "Your current cursor position needs to be outside of a table"
                'If sect.Index = sectLast.Index Then strErrorMsg = "Insertion of a chapter behind the last section is not supported"


                'If Not Me.objChpt.chpt_is_BodyPage(sect) Then strErrorMsg = "A new chapter section can only be inserted" + vbCrLf + "behind an existing chapter section."
                'If Me.objCpMgr.cp_Bool_IsCoverPage(sect) Then strErrorMsg = "A chapter cannot be inserted behind a Cover Page"
                'If Me.objContactsPgMgr.is_ContactsPage_Front(sect) Then strErrorMsg = "A chapter cannot be inserted behind the front Contacts Page"
                'If Me.objDivMgr.is_divider_App(sect) Then strErrorMsg = "A chapter should not be inserted behind the Appendix Divider"

                'If Me.objContactsPgMgr.is_ContactsPage_Back(sect) Then strErrorMsg = "A chapter cannot be inserted behind the back Contacts Page"

        End Select
        '
        Return strErrorMsg
    End Function
    '

    '
    ''' <summary>
    ''' This method will return a null string if the action proposed (in this case a Chapter insert in front of the current
    ''' section) is a permitted operation. Otherwise it will return an error message. Note that calling methods will test for
    ''' the null string to determined whetehr to proceed or not. If not then the error message may be output
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function isOKto_Insert_ChapterInFront(ByRef sect As Word.Section, strRptMode As String) As String
        Dim objRptMgr As New cReport()
        Dim myDoc As Word.Document
        Dim sectLast As Word.Section
        Dim strErrorMsg As String = ""
        '
        myDoc = sect.Range.Document
        sectLast = myDoc.Sections.Last
        '
        Select Case strRptMode
            Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                If Me.objCpMgr.cp_Bool_IsCoverPage(sect) Then strErrorMsg = "A chapter cannot be inserted in front of a Cover Page"
                If Me.objContactsPgMgr.is_ContactsPage_Front(sect) Then strErrorMsg = "A chapter cannot be inserted in front of the front Contacts Page"
                If Me.objTOCMgr.toc_is_TOCSection(sect) Then strErrorMsg = "A chapter cannot be inserted in front of a TOC"
                '
            Case objRptMgr.rpt_isBrief
                'If objRptMgr.sct_Sel_IsIn_TableOnly() Then strErrorMsg = "Your current cursor position needs to be outside of a table"
                'If sect.Index = sectLast.Index Then strErrorMsg = "Insertion of a chapter behind the last section is not supported"
                '
        End Select
        '
        Return strErrorMsg
    End Function
    '
    '
    Public Function isOKto_Insert_ChapterBehind(ByRef sect As Word.Section, strRptMode As String) As String
        Dim strErrorMsg As String = ""
        Dim myDoc As Word.Document
        Dim sectLast As Word.Section
        Dim objRptMgr As New cReport()
        '
        myDoc = sect.Range.Document
        sectLast = myDoc.Sections.Last
        '
        Select Case strRptMode
            Case objRptMgr.rpt_isPrt, objRptMgr.rpt_isLnd
                'If Not Me.objChpt.chpt_is_BodyPage(sect) Then strErrorMsg = "A new chapter section can only be inserted" + vbCrLf + "behind an existing chapter section."
                If sect.Index = sectLast.Index Then strErrorMsg = "Insertion of a chapter behind the last section is not supported"

            Case objRptMgr.rpt_isBrief
                '
                If objRptMgr.sct_Sel_IsIn_TableOnly() Then strErrorMsg = "Your current cursor position needs to be outside of a table"
                If sect.Index = sectLast.Index Then strErrorMsg = "Insertion of a chapter behind the last section is not supported"


                'If Not Me.objChpt.chpt_is_BodyPage(sect) Then strErrorMsg = "A new chapter section can only be inserted" + vbCrLf + "behind an existing chapter section."
                'If Me.objCpMgr.cp_Bool_IsCoverPage(sect) Then strErrorMsg = "A chapter cannot be inserted behind a Cover Page"
                'If Me.objContactsPgMgr.is_ContactsPage_Front(sect) Then strErrorMsg = "A chapter cannot be inserted behind the front Contacts Page"
                'If Me.objDivMgr.is_divider_App(sect) Then strErrorMsg = "A chapter should not be inserted behind the Appendix Divider"

                'If Me.objContactsPgMgr.is_ContactsPage_Back(sect) Then strErrorMsg = "A chapter cannot be inserted behind the back Contacts Page"

        End Select
        '
        '
        Return strErrorMsg
    End Function

    '
    '
    Public Function isOKto_Insert_BackPanel(ByRef sect As Word.Section) As String
        Dim strErrorMsg As String = ""
        '
        If Me.objCpMgr.cp_Bool_IsCoverPage(sect) Then strErrorMsg = "Insert/delete of a back panel into the cover page is not supported"
        If Me.objContactsPgMgr.is_ContactsPage_Front(sect) Then strErrorMsg = "Insert/delete of a back panel into the front 'contacts' page is not supported"
        If Me.objDivMgr.is_divider_Chpt(sect) Then strErrorMsg = "Insert/delete of a back panel into a report divider is not supported"
        If Me.objDivMgr.is_divider_App(sect) Then strErrorMsg = "Insert/delete of a back panel into the appendices divider is not supported"
        If Me.objContactsPgMgr.is_ContactsPage_Back(sect) Then strErrorMsg = "Insert/delete of a back panel into the back contacts page is not supported"
        '
        Return strErrorMsg
    End Function
    '
    Public Function isOKto_Delete_BackPanel(ByRef sect As Word.Section) As String
        Dim strErrorMsg As String = ""
        '
        If Me.objCpMgr.cp_Bool_IsCoverPage(sect) Then strErrorMsg = "Deleting a back panel from the cover page is not supported"
        If Me.objContactsPgMgr.is_ContactsPage_Front(sect) Then strErrorMsg = "Deleting a back panel from the front 'contacts' page is not supported"
        If Me.objDivMgr.is_divider_Chpt(sect) Then strErrorMsg = "Deleteing a back panel from a report divider is not supported"
        If Me.objDivMgr.is_divider_App(sect) Then strErrorMsg = "Deleting a back panel from the appendices divider is not supported"
        If Me.objContactsPgMgr.is_ContactsPage_Back(sect) Then strErrorMsg = "Deleting a back panel from the back contacts page is not supported"
        '
        Return strErrorMsg
        '
    End Function
    '


    '
    Public Function isOKto_Insert_DividerChpt(ByRef sect As Word.Section) As String
        Dim strErrorMsg As String = ""
        '
        If Me.objCpMgr.cp_Bool_IsCoverPage(sect) Then strErrorMsg = "Inserting a Divider from this location is not supported"
        If Me.objContactsPgMgr.is_ContactsPage_Front(sect) Then strErrorMsg = "Inserting a Divider from this location is not supported"
        If Me.objTOCMgr.toc_is_TOCSection(sect) Then strErrorMsg = "Inserting a Divider from this location is not supported"
        If objContactsPgMgr.is_ContactsPage_Back(sect) Then strErrorMsg = "Inserting a Divider from this location is not supported"
        '
        '
        Return strErrorMsg
    End Function
    '
    Public Function isOKto_toggle_PlhWidth() As String
        Dim strErrorMsg As String = "ok"
        Dim sect As Word.Section
        Dim tbl As Word.Table
        '
        sect = glb_get_wrdSect()
        tbl = Nothing
        '
        If Me.objCpMgr.cp_Bool_IsCoverPage(sect) Then strErrorMsg = "This function is not supported in the Cover Page"
        If Me.objContactsPgMgr.is_ContactsPage_Front(sect) Then strErrorMsg = "This function is not supported in the front Contacts Page"
        If Me.objTOCMgr.toc_is_TOCSection(sect) Then strErrorMsg = "This function is not supported in the TOC"
        If Me.objContactsPgMgr.is_ContactsPage_Back(sect) Then strErrorMsg = "This function is not supported in the back Contacts Page"
        If Me.objDivMgr.is_divider_Chpt(sect) Then strErrorMsg = "This function is not supported in a report Divider"
        If Me.objDivMgr.is_divider_App(sect) Then strErrorMsg = "This function is not supported in the Appendices Divider"
        '
        tbl = glb_get_wrdSelTbl()
        If Not IsNothing(tbl) Then
            If Me.objChptBanner.bnr_is_Chapter_Bdy_or_ES_or_AP(tbl) Then strErrorMsg = "This function is not supported in heading banners. Please make certain that your cursor is at least one paragraph clear of the banner"
        End If

        '
        Return strErrorMsg
        '
    End Function
    '
    ''' <summary>
    ''' This function will return a null string if the referenced section supports modification
    ''' of the footer text. Basically it determines whether the sect is a CoverPage, Contacts Page,
    ''' TOC, or stationery pages (letter or memo).. If it is any of these an error message is returned
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function isOKto_reset_footerText(ByRef sect As Word.Section) As String
        Dim strErrorMsg As String = ""
        Dim tbl As Word.Table
        '
        tbl = Nothing
        '
        If Me.objCpMgr.cp_Bool_IsCoverPage(sect) Then strErrorMsg = "This function is not supported in the Cover Page"
        If Me.objContactsPgMgr.is_ContactsPage_Front(sect) Then strErrorMsg = "This function is not supported in the front Contacts Page"
        If Me.objTOCMgr.toc_is_TOCSection(sect) Then strErrorMsg = "This function is not supported in the TOC"
        If Me.objContactsPgMgr.is_ContactsPage_Back(sect) Then strErrorMsg = "This function is not supported in the back Contacts Page"
        If Me.objLetterMgr.ltr_is_letter(sect) Then strErrorMsg = "This function is not supported in stationary pages"
        If Me.objMemoMgr.memo_is_memo(sect) Then strErrorMsg = "This function is not supported in stationary pages"
        '
        'tbl = glb_get_wrdSelTbl()
        'If Not IsNothing(tbl) Then
        'If Me.objChptBanner.bnr_is_Chapter_Bdy_or_ES_or_AP(tbl) Then strErrorMsg = "This function is not supported in heading banners. Please make certain that your cursor is at least one paragraph clear of the banner"
        'End If
        '
        Return strErrorMsg
        '
    End Function
    '

    '
    Public Function isOKto_doAction_inReportBody() As String
        Dim strErrorMsg As String = Me._isOK
        Dim sect As Word.Section
        Dim tbl As Word.Table
        '
        sect = glb_get_wrdSect()
        tbl = Nothing
        '
        If Me.objCpMgr.cp_Bool_IsCoverPage(sect) Then strErrorMsg = "This function is not supported in the Cover Page"
        If Me.objContactsPgMgr.is_ContactsPage_Front(sect) Then strErrorMsg = "This function is not supported in the front Contacts Page"
        If Me.objTOCMgr.toc_is_TOCSection(sect) Then strErrorMsg = "This function is not supported in the TOC"
        If Me.objContactsPgMgr.is_ContactsPage_Back(sect) Then strErrorMsg = "This function is not supported in the back Contacts Page"
        If Me.objDivMgr.is_divider_Chpt(sect) Then strErrorMsg = "This function is not supported in a report Divider"
        If Me.objDivMgr.is_divider_App(sect) Then strErrorMsg = "This function is not supported in the Appendices Divider"
        '
        'tbl = glb_get_wrdSelTbl()
        'If Not IsNothing(tbl) Then
        'If Me.objChptBanner.bnr_is_Chapter_Bdy_or_ES_or_AP(tbl) Then strErrorMsg = "This function is not supported in heading banners. Please make certain that your cursor is at least one paragraph clear of the banner"
        'End If

        '
        Return strErrorMsg
        '
    End Function
    '
    '
    Public Function isOKto_doAction_inReportBody(ByRef rng As Word.Range) As String
        Dim strErrorMsg As String = Me._isOK
        Dim sect As Word.Section
        Dim tbl As Word.Table
        '
        sect = rng.Sections.Item(1)
        tbl = Nothing
        '
        If Me.objCpMgr.cp_Bool_IsCoverPage(sect) Then strErrorMsg = "This function is not supported in the Cover Page"
        If Me.objContactsPgMgr.is_ContactsPage_Front(sect) Then strErrorMsg = "This function is not supported in the front Contacts Page"
        If Me.objTOCMgr.toc_is_TOCSection(sect) Then strErrorMsg = "This function is not supported in the TOC"
        If Me.objContactsPgMgr.is_ContactsPage_Back(sect) Then strErrorMsg = "This function is not supported in the back Contacts Page"
        If Me.objDivMgr.is_divider_Chpt(sect) Then strErrorMsg = "This function is not supported in a report Divider"
        If Me.objDivMgr.is_divider_App(sect) Then strErrorMsg = "This function is not supported in the Appendices Divider"
        '
        '
        Return strErrorMsg
        '
    End Function
    '

    Public Function isOKto_section_is(ByRef sect As Word.Section) As String
        Dim objDivMgr As New cChptDivider()
        Dim objContactMgr As New cContactsMgr()
        Dim objCpMgr As New cCoverPageMgr()
        Dim objTocMgr As New cTOCMgr()
        Dim strType As String
        '
        strType = "unknown"
        '
        If objCpMgr.cp_Bool_IsCoverPage(sect) Then strType = "cp"
        If objContactMgr.is_ContactsPage_Front(sect) Then strType = "contFront"
        If objTocMgr.toc_is_TOCSection(sect) Then strType = "toc"
        If objDivMgr.is_divider_Chpt(sect) Then strType = "div"
        If objDivMgr.is_divider_App(sect) Then strType = "divAp"
        If objContactMgr.is_ContactsPage_Back(sect) Then strType = "contBack"
        '
        Return strType
        '
    End Function
    '
    Public Function isOKto_insert_Table() As String
        Dim objGlobals As New cGlobals()
        Dim rng As Word.Range
        Dim strErrorMsg As String
        '
        strErrorMsg = ""
        '
        rng = objGlobals.glb_get_wrdSelRng()
        If rng.Tables.Count <> 0 Then
            strErrorMsg = "Your cursor must be located at least one clear paragraph away from any existing tables, otherwise they'll merge in unexpected ways." + vbCrLf + vbCrLf + "Please relocate your insertion point and try again"
        End If
        '
        Return strErrorMsg
        '
    End Function
    '
    '
    Public Function isOKto_selection_isIn() As String
        Dim objDivMgr As New cChptDivider()
        Dim objContactMgr As New cContactsMgr()
        Dim objCpMgr As New cCoverPageMgr()
        Dim objTocMgr As New cTOCMgr()
        Dim objCaseStudy As New cCaseStudyMgr()
        Dim objGlossMgr As New cGlossary()
        Dim sect As Word.Section
        Dim strType As String
        '
        sect = glb_get_wrdSect()
        strType = "unknown"
        '
        If objCpMgr.cp_Bool_IsCoverPage(sect) Then strType = "cp"
        If objContactMgr.is_ContactsPage_Front(sect) Then strType = "contFront"
        If objTocMgr.toc_is_TOCSection(sect) Then strType = "toc"
        If objGlossMgr.glos_is_Glossary(sect) Then strType = "glos"
        If objDivMgr.is_divider_Chpt(sect) Then strType = "div"
        If objDivMgr.is_divider_App(sect) Then strType = "divAp"
        If objContactMgr.is_ContactsPage_Back(sect) Then strType = "contBack"
        If objCaseStudy.cst_is_caseStudySection(sect) Then strType = "caseStudy"
        '
        If Me.isOKto_isIn_Brief(sect) Then
            strType = "brief"
            If sect.Index = 1 Then strType = "briefFirstSection"
        End If
        '
        Return strType
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will return true if the current section is a 'Brief' section
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function isOKto_isIn_Brief(ByRef sect As Word.Section) As Boolean
        Dim rslt As Boolean
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objBnrMgr As New cChptBanner()
        Dim strTagStyle As String
        '
        rslt = False
        '
        strTagStyle = objHFMgr.hf_tags_getTagStyleName(sect, "primary")
        '
        If strTagStyle = objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_brief) Then
            rslt = True
        End If
        '
        Return rslt
        '
    End Function
    '

End Class

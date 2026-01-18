Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cInsertTestMgr
    Inherits cChptBanner
    Public Sub New()
        MyBase.New()
    End Sub
    '
    '
    'This function checks for illegal insertion actions. It first makes certain that
    'the user is not trying to insert anything is the Cover Page, Contacts Page (Front and Back),
    'Table of Contents.. If so, we get an error message and it returns False. It
    'will then check to see if we are trying to insert in a Table. If so, then we get
    'an error message and it will return false
    Public Overridable Function ins_is_OKToInsert(ByRef sect As Section, Optional doTableCheck As Boolean = True) As String
        Dim objCpMgr As New cCoverPageMgr()
        Dim objSectMgr As New cSectionMgr()
        Dim objTOCMgr As New cTOCMgr()
        Dim objRpt As New cReport()
        Dim strTagName As String
        Dim strErrorMsg As String
        Dim strRptMode As String
        Dim objTagsMgr As New cTagsMgr()
        '
        strErrorMsg = ""
        strRptMode = objRpt.Rpt_Mode_Get()
        strTagName = objTagsMgr.tags_get_tagStyleName(sect)
        '
        If objCpMgr.cp_Bool_IsCoverPage(sect) Then
            strErrorMsg = "Insertion In a Cover Page Is Not supported"
            GoTo finis
        End If
        If objTOCMgr.toc_is_TOCSection(sect) Then
            strErrorMsg = "Insertion In the Table Of Contents Is Not supported"
            GoTo finis
        End If
        '
        If objTagsMgr.tags_is_ContactsPage_Front(glb_get_wrdSect()) Then
            strErrorMsg = "Insertion In a Contacts Page Is Not supported"
            GoTo finis
        End If
        '
        If objTagsMgr.tags_is_ContactsPage_Back(glb_get_wrdSect()) Then
            strErrorMsg = "Insertion In a Contacts Page Is Not supported"
            GoTo finis
        End If
        '
        Select Case strRptMode
            Case objRpt.rpt_isPrt, objRpt.rpt_isBrief, objRpt.rpt_isLnd
                If strTagName Like "tag_*" And Not (strTagName = "tag_chapterBanner") Then
                    Select Case strTagName
                        Case "tag_chapterBanner", "tag_execBanner", "tag_appendixChapter", "tag_aaBrief"
                            'Its OK
                        Case Else
                            strErrorMsg = "Insertion of the object in this page is not supported"
                            GoTo finis
                    End Select
                End If
                'Case objRpt.modeLongLandscape
                'If strTagName Like "tag_*" Then
                'strErrorMsg = "Insertion of the object in this page is not supported"
                'GoTo finis
                'End If
        End Select
        '
        'Check to see if in table
        '
        'Now check for attempts to insert in a Table
        If doTableCheck Then
            If objSectMgr.objGlobals.glb_selection_IsInTable() Then
                strErrorMsg = "Your cursor needs to be in a clear part of the document" & vbCr _
            & "for this operation to succeed. (It is probably in a Table, or just under one)" & vbCr & vbCr _
            & "Please relocate the cursor and try again"
            End If
        End If


        If strErrorMsg <> "" Then
            'rslt = False
            'MsgBox(strErrorMsg)
            'Return rslt
            'Exit Function
        End If
        '
finis:
        Return strErrorMsg

    End Function
    '


End Class

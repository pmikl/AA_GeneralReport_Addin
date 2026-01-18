Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cTagsMgr
    Inherits cChptBanner
    Public Sub New()
        MyBase.New()

    End Sub
    '
    '
    Public Function tags_is_OKToUseMultiColumn(ByRef sect As Word.Section) As Boolean
        Dim rslt As Boolean
        '
        rslt = True
        '
        If Me.tags_is_CoverPage(sect) Then
            rslt = False
            GoTo finis
        End If
        If Me.tags_is_ContactsPage_Front(sect) Then
            rslt = False
            GoTo finis
        End If
        If Me.tags_is_TOCPage(sect) Then
            rslt = False
            GoTo finis
        End If
        If Me.tags_is_ExecutiveSummary(sect) Then
            rslt = False
            GoTo finis
        End If
        If Me.tags_is_Divider_Main(sect) Then
            rslt = False
            GoTo finis
        End If
        If Me.tags_is_Chapter_Body(sect) Then
            rslt = False
            GoTo finis
        End If
        If Me.tags_is_Divider_Appendix(sect) Then
            rslt = False
            GoTo finis
        End If
        If Me.tags_is_Chapter_Appendix(sect) Then
            rslt = False
            GoTo finis
        End If
        If Me.tags_is_ContactsPage_Back(sect) Then
            rslt = False
            GoTo finis
        End If
        'If MyBase.is_Bibliography(sect) Then
        'rslt = False
        'GoTo finis
        'End If
        '
finis:
        Return rslt
        '
    End Function

    '
#Region "Section/Page type"

    ''' <summary>
    ''' This method will determine if the current section contains a cover page. It does so by looking for the
    ''' style tag_coverPage in the Header of the first page
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function tags_is_CoverPage(ByRef sect As Section) As Boolean
        Dim rng As Range
        Dim rngStyle As Word.Style
        Dim rslt As Boolean
        '
        rslt = False
        rng = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rngStyle = rng.Style
        '
        'MsgBox("Style is = " + rngStyle.NameLocal)
        '
        'If drCell.Range.Style Is Globals.ThisAddin.Application.ActiveDocument.Styles("tag_coverPage") Then
        If rngStyle.NameLocal = MyBase.bnr_get_tagStyles(Me.tag_coverPage) Then
            rslt = True
        End If
        '
        Return rslt
    End Function
    '
    Public Function tags_is_ContactsPage_Front(ByRef sect As Section) As Boolean
        Dim rslt As Boolean
        '
        rslt = Me.tags_is_MyPage(sect, MyBase.bnr_get_tagStyles(Me.tag_cont_Front))
        '
        Return rslt
    End Function
    '
    ''' <summary>
    ''' This method will return true if the current selection is in the TOC.
    ''' Note that there is a duplicate of this function in cFormatMgr.. It's there
    ''' for startup performance reasons only
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function tags_is_TOCPage(ByRef sect As Section) As Boolean
        Dim hf As HeaderFooter
        Dim tbl As Table
        Dim drCell As Cell
        Dim styl As Style
        Dim rslt As Boolean
        '
        rslt = False
        '
        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        If sect.PageSetup.DifferentFirstPageHeaderFooter Then
            hf = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        End If
        '
        Try
            For Each tbl In hf.Range.Tables
                For Each drCell In tbl.Range.Cells
                    styl = drCell.Range.Style
                    If styl.NameLocal Like "TOC Heading*" Then
                        rslt = True
                        Exit For
                    End If
                Next drCell
            Next tbl
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
    End Function
    '
    '
    Public Function tags_is_Memo(ByRef sect As Section) As Boolean
        Dim rslt As Boolean
        '
        rslt = Me.tags_is_MyPage(sect, MyBase.bnr_get_tagStyles(Me.tag_memo))
        '
        Return rslt
    End Function
    '
    '
    Public Function tags_is_Letter(ByRef sect As Section) As Boolean
        Dim rslt As Boolean
        '
        rslt = Me.tags_is_MyPage(sect, MyBase.bnr_get_tagStyles(Me.tag_letter))
        '
        Return rslt
    End Function
    '
    '
    Public Function tags_is_Brief(ByRef sect As Section) As Boolean
        Dim rslt As Boolean
        '
        rslt = Me.tags_is_MyPage(sect, MyBase.bnr_get_tagStyles(Me.tag_brief))
        '
        Return rslt
    End Function
    '

    '
    '
    Public Function tags_is_ExecutiveSummary(ByRef sect As Section) As Boolean
        Dim rslt As Boolean
        '
        rslt = Me.tags_is_MyPage(sect, MyBase.bnr_get_tagStyles(Me.tag_es))
        '
        Return rslt
    End Function
    '
    Public Function tags_is_Divider_Main(ByRef sect As Section) As Boolean
        Dim rslt As Boolean
        '
        rslt = Me.tags_is_MyPage(sect, MyBase.bnr_get_tagStyles(Me.tag_div))
        '
        Return rslt
    End Function
    '
    Public Function tags_is_Chapter_Body(ByRef sect As Section) As Boolean
        Dim rslt As Boolean
        '
        rslt = Me.tags_is_MyPage(sect, MyBase.bnr_get_tagStyles(Me.tag_chpt_body))
        '
        Return rslt
    End Function
    '
    Public Function tags_is_Divider_Appendix(ByRef sect As Section) As Boolean
        Dim rslt As Boolean
        '
        rslt = Me.tags_is_MyPage(sect, MyBase.bnr_get_tagStyles(tag_divAP))
        '
        Return rslt
    End Function
    '
    Public Function tags_is_Chapter_Appendix(ByRef sect As Section) As Boolean
        Dim rslt As Boolean
        '
        rslt = Me.tags_is_MyPage(sect, MyBase.bnr_get_tagStyles(tag_chpt_AP))
        '
        Return rslt
    End Function
    '
    '
    Public Function tags_is_ContactsPage_Back(ByRef sect As Section) As Boolean
        Dim rslt As Boolean
        '
        rslt = Me.tags_is_MyPage(sect, MyBase.bnr_get_tagStyles(tag_divAP))
        '
        Return rslt
    End Function
    '
    '
    ''' <summary>
    ''' This method will reurn true if the section (sect) has a banner table and if that
    ''' Banner table has a style in the first cell that matches strTagsStyle
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function tags_is_MyPage(ByRef sect As Word.Section, strTagStyle As String) As Boolean
        Dim strResult As String
        Dim rslt As Boolean
        '
        rslt = False
        strResult = Me.tags_get_tagStyleName(sect)
        '
        If strResult = strTagStyle Then rslt = True
        '
        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method will search the primary and/or first page header for a header table. If one exists it will
    ''' get the name of the style in the first cell. This is the tag style. If the header table doesn't exist
    ''' it will search for the style of the header and present that as the tag style
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function tags_get_tagStyleName(ByRef sect As Word.Section) As String
        Dim strTagStyleName As String
        Dim objHfMgr As New cHeaderFooterMgr()
        '
        'If the search of the primary does not generate a result then search the first page
        strTagStyleName = objHfMgr.hf_tags_getTagStyleName(sect)
        'If strTagStyleName = "" Then strTagStyleName = objHfMgr.hf_tags_getTagStyleName(sect, "firstPage")
        '
        'rng = sect.Range
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        'If rng.Tables.Count <> 0 Then
        'myStyle = rng.Style
        'strTagStyleName = myStyle.NameLocal
        'End If

        Return strTagStyleName
        '
    End Function
    '

#End Region
    '
    ''' <summary>
    ''' This method will look in the first cell of a table and returns the style name
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tags_get_tagStyleName(ByRef tbl As Word.Table) As String
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim strTag As String
        '
        Try
            strTag = objHfMgr.hf_tags_getTagStyleName(tbl)
        Catch ex As Exception
            strTag = ""
        End Try
        '
        Return strTag
    End Function
    '
    '
    Public Function getTagType(strTag As String) As String
        Dim tokens() As String
        '
        Try
            getTagType = ""
            If strTag <> "" Then
                tokens = Split(strTag, "_")
                getTagType = tokens(1)
            End If
        Catch ex As Exception
            getTagType = ""                 'Hence getSectionType to null
        End Try

    End Function
    '
    '
    ''' <summary>
    ''' This method will determine if the current selection is in a Heading banner
    ''' </summary>
    ''' <returns></returns>
    Public Function tags_is_inBanner() As Boolean
        Dim objTagsMgr As New cTagsMgr()
        Dim rslt As Boolean
        Dim tbl As Word.Table
        Dim strTagsStyle As String
        '
        rslt = False
        '
        tbl = Me.glb_get_wrdSelTbl()
        '
        Try
            strTagsStyle = objTagsMgr.tags_get_tagStyleName(tbl)
            Select Case strTagsStyle
                Case "tag_execBanner", "tag_partBanner", "tag_chapterBanner", "tag_appendixPart", "tag_appendixChapter"
                    rslt = True
                Case "tag_biblio_Chpt", "tag_glossary_Chpt", "tag_refs_Chpt", "tag_worksCited_Chpt"
                    rslt = True
            End Select

        Catch ex As Exception
            rslt = False
        End Try


        Return rslt
    End Function


End Class

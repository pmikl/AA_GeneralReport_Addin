Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cPageNumberMgr
    Inherits cChptBase
    Public pgNum_pgNumType_std As String
    Public pgNum_pgNumype_2part As String
    Public Sub New()
        MyBase.New()
        Me.pgNum_pgNumType_std = "std"
        Me.pgNum_pgNumype_2part = "2part"
    End Sub
    '
    ''' <summary>
    ''' This method will get the Page Number format seting for the current document. It will
    ''' be 'std' or '2part'
    ''' </summary>
    ''' <returns></returns>
    Public Function pgNum_get_numFormat_ForDoc() As String
        Dim objPropsMgr As New cPropertyMgr()
        Dim strResult As String
        '
        strResult = objPropsMgr.prps_CustomProperty_get("pgNumberFormat", Me.pgNum_pgNumType_std)
        '
        Return strResult
    End Function
    '
    '

    ''' <summary>
    ''' This method will set the Page Number format for the current document. It will overwrite
    ''' whatever is there with either 'std' or '2part'
    ''' </summary>
    ''' <param name="strNumberoFormat"></param>
    Public Sub pgNum_set_numFormat_ForDoc(Optional strNumberoFormat As String = "std")
        Dim objPropsMgr As New cPropertyMgr()
        '
        Select Case strNumberoFormat
            Case "std"
                objPropsMgr.prps_CustomProperty_set(Me.pgNum_pgNumType_std, "pgNumberFormat")
            Case "2part"
                objPropsMgr.prps_CustomProperty_set(Me.pgNum_pgNumype_2part, "pgNumberFormat")
        End Select
        '
        Return
    End Sub
    '

    '
    ''' <summary>
    ''' This method will retrieve the 'pgNumberFormat' for this document (defaulting to 'std' if
    ''' none exists) and will then set the page numbering for the body of this report as per this format
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub pgNum_setBody_numFormat(ByRef myDoc As Word.Document)
        Dim objPropsMgr As New cPropertyMgr()
        Dim strResult As String
        '

        strResult = objPropsMgr.prps_CustomProperty_get("pgNumberFormat", Me.pgNum_pgNumType_std)
        Me.pgNum_setBody_numFormat(myDoc, strResult)
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will set the page number format for the body of the document to the type identified
    ''' in pgNumStyle
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="pgNumStyle"></param>
    Public Sub pgNum_setBody_numFormat(ByRef myDoc As Word.Document, Optional pgNumStyle As String = "std")
        Dim sect, sectPrior As Word.Section
        Dim reachedEnd, resetNumbering, hasGlossary As Boolean
        Dim objTagsMgr As New cTagsMgr()
        Dim objTocMgr As New cTOCMgr()
        Dim strTag, strTag_Last, strTagPrior As String
        Dim j, divCount, chptIndex As Integer
        Dim objPropsMgr As New cPropertyMgr()
        Dim objHfMgr As New cHeaderFooterMgr()
        '
        chptIndex = 0
        divCount = 0
        resetNumbering = False
        hasGlossary = False
        strTag = ""
        strTag_Last = ""
        strTagPrior = ""
        '
        If pgNumStyle = "2part" Then resetNumbering = True
        '
        For j = 1 To myDoc.Sections.Last.Index
            sect = myDoc.Sections.Item(j)
            'strTag = objTagsMgr.tags_get_tagStyleName(sect)
            strTag = objHfMgr.hf_tags_getTagStyleName(sect)
            '
            If Not strTag = "" Then
                strTag_Last = strTag
            End If

            Select Case strTag
                Case "tag_aaBrief"
                    If sect.Index = 1 Then
                        chptBase_PageNumbering_Set(sect, True, 1, "std")
                    Else
                        chptBase_PageNumbering_Set(sect, False, 1, "std")
                    End If

                Case "tag_execBanner"

                    If hasGlossary Then
                        chptBase_PageNumbering_Set(sect, False, 1, "es")
                    Else
                        'No Glossary, so exec summary starts at i
                        chptBase_PageNumbering_Set(sect, True, 1, "es")

                    End If
                    '
                    Try
                        sectPrior = myDoc.Sections.Item(sect.Index - 1)
                        strTagPrior = objHfMgr.hf_tags_getTagStyleName(sectPrior)
                        Select Case pgNumStyle
                            Case Me.pgNum_pgNumType_std
                                If strTagPrior = "tag_partBanner" Then
                                    chptBase_PageNumbering_Set(sectPrior, False, 1, "es")
                                    chptBase_PageNumbering_Set(sect, False, 1, "es")
                                End If
                            Case Me.pgNum_pgNumype_2part
                                If strTagPrior = "tag_partBanner" Then
                                    chptBase_PageNumbering_Set(sectPrior, True, divCount, "div")
                                    chptBase_PageNumbering_Set(sect, False, 1, "es")
                                End If
                        End Select

                        '
                    Catch ex As Exception

                    End Try

                Case "tag_glossary_Chpt"
                    hasGlossary = True
                    chptBase_PageNumbering_Set(sect, True, 1, "es")

                Case "tag_partBanner", "tag_biblio_Chpt", "tag_refs_Chpt", "tag_worksCited_Chpt"
                    divCount = divCount + 1
                    Select Case pgNumStyle
                        Case Me.pgNum_pgNumType_std
                            If divCount = 1 And chptIndex = 0 Then
                                chptBase_PageNumbering_Set(sect, True, 1, pgNumStyle)
                            Else
                                chptBase_PageNumbering_Set(sect, False, 1, pgNumStyle)
                            End If
                        Case Me.pgNum_pgNumype_2part
                            chptBase_PageNumbering_Set(sect, True, divCount, "div")
                    End Select
                Case "tag_chapterBanner"
                    chptIndex = chptIndex + 1
                    Select Case pgNumStyle
                        Case Me.pgNum_pgNumType_std
                            If chptIndex = 1 Then
                                If divCount = 0 Then
                                    chptBase_PageNumbering_Set(sect, True, 1, pgNumStyle)
                                Else
                                    sectPrior = myDoc.Sections.Item(sect.Index - 1)
                                    strTagPrior = objHfMgr.hf_tags_getTagStyleName(sectPrior)
                                    '
                                    If chptIndex = 1 And strTagPrior = "tag_partBanner" Then
                                        chptBase_PageNumbering_Set(sectPrior, True, 1, pgNumStyle)
                                        chptBase_PageNumbering_Set(sect, False, 1, pgNumStyle)
                                    Else
                                        chptBase_PageNumbering_Set(sect, False, 1, pgNumStyle)

                                    End If

                                    'chptBase_PageNumbering_Set(sect, False, 1, pgNumStyle)
                                End If
                            Else
                                chptBase_PageNumbering_Set(sect, resetNumbering, 1, pgNumStyle)
                            End If
                            '
                        Case Me.pgNum_pgNumype_2part
                            chptBase_PageNumbering_Set(sect, resetNumbering, 1, pgNumStyle)
                    End Select
                Case "tag_toc"

                Case "tag_execBannerx"
                    'chptBase_PageNumbering_Set(sect, False, 1, "es")
                    '
                Case "tag_glossary_Chpt"
                    chptBase_PageNumbering_Set(sect, False, 1, "es")
                    '
                Case "tag_partBanner", "tag_biblio_Chpt", "tag_refs_Chpt", "tag_worksCited_Chpt"
                    chptBase_PageNumbering_Set(sect, False, 1, "div")

                Case ""
                    'If objTocMgr.toc_is_TOCSection(sect) Then
                    'Is TOC Section, lts leave it alone
                    'Else
                    'No tag and not toc section, so it must be some other section in the
                    'document... Make sure the numbering is continous
                    'Select Case strTag_Last
                    'Case "tag_execBanner"
                    'chptBase_PageNumbering_Set(sect, False, 1, "es")
                    'Case "tag_glossary_Chpt"
                    'chptBase_PageNumbering_Set(sect, False, 1, "es")
                    'Case "tag_partBanner", "tag_biblio_Chpt", "tag_refs_Chpt", "tag_worksCited_Chpt"
                    'chptBase_PageNumbering_Set(sect, False, 1, "div")
                    'Case Else
                    'chptBase_PageNumbering_Set(sect, False, 1, pgNumStyle)
                    'End Select


                    'End If
                Case "tag_appendixPart"
                    divCount = divCount + 1
                    sectPrior = myDoc.Sections.Item(sect.Index - 1)
                    '
                    Select Case pgNumStyle
                        Case Me.pgNum_pgNumType_std
                            If chptIndex > 0 And divCount = 1 Then
                                chptBase_PageNumbering_Set(sect, False, 1, pgNumStyle)
                            Else
                                chptBase_PageNumbering_Set(sect, False, 1, pgNumStyle)
                            End If
                        Case Me.pgNum_pgNumype_2part
                            chptBase_PageNumbering_Set(sect, True, divCount, "div")
                    End Select
                    '
                    '*** New version
                    If divCount = 1 Then
                        'chptBase_PageNumbering_Set(sect, False, 1, "div")
                    End If
                    If divCount > 1 And resetNumbering Then
                        'chptBase_PageNumbering_Set(sect, True, divCount, "div")
                    End If
                    If divCount > 1 And Not (resetNumbering) Then
                        'chptBase_PageNumbering_Set(sect, True, divCount, "div")
                    End If
                    '
                    '*** Old version
                    If Me.pgNum_is_DivPageNumberFormat(sectPrior) Then
                        'chptBase_PageNumbering_Set(sect, False, 1, "div")
                    Else
                        'chptBase_PageNumbering_Set(sect, True, divCount, "div")
                    End If
                    reachedEnd = True
                    Exit For
                Case "tag_appendixChapter"
                    reachedEnd = True
                    Exit For
                Case Else

            End Select
        Next

        objTocMgr.toc_update_TOCs(myDoc)
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will set the page number format for the body of the document to the type identified
    ''' in pgNumStyle
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="pgNumStyle"></param>
    Public Sub xpgNum_setBody_numFormat(ByRef myDoc As Word.Document, Optional pgNumStyle As String = "std")
        Dim sect, sectPrior As Word.Section
        Dim reachedEnd As Boolean
        Dim objTagsMgr As New cTagsMgr()
        Dim objTocMgr As New cTOCMgr()
        Dim strTag As String
        Dim j, chptStartIndex, divStartIndex, divCount As Integer
        '
        divCount = 0
        reachedEnd = False
        '
        chptStartIndex = Me.pgNum_get_FirstChptIndex(myDoc)
        divStartIndex = Me.pgNum_get_FirstDivIndex(myDoc)
        '
        If divStartIndex = -1 And chptStartIndex = -1 Then
            GoTo finis
        End If
        '
        If divStartIndex = -1 And Not (chptStartIndex = -1) Then
            'Dummy up the divStartIndex so that the code branches to the
            'option that starts at the first chapter
            divStartIndex = chptStartIndex + 1
        End If
        '
        'Dividers Only
        If Not (divStartIndex = -1) And chptStartIndex = -1 Then GoTo finis
        '
        If divStartIndex < chptStartIndex Then
            'A divider before the first chapter
            chptBase_PageNumbering_Set(myDoc.Sections.Item(divStartIndex), True, 1, "div")
            divCount = divCount + 1
            '
            If pgNumStyle = "std" Then
                For j = divStartIndex + 1 To myDoc.Sections.Last.Index
                    sect = myDoc.Sections.Item(j)
                    strTag = objTagsMgr.tags_get_tagStyleName(sect)
                    Select Case strTag
                        Case "tag_execBanner"
                            chptBase_PageNumbering_Set(sect, True, 1, "es")
                        Case "tag_partBanner", "tag_glossary_Chpt", "tag_biblio_Chpt", "tag_refs_Chpt", "tag_worksCited_Chpt"
                            chptBase_PageNumbering_Set(sect, False, 1, "div")
                        Case "tag_chapterBanner"
                            chptBase_PageNumbering_Set(sect, False, 1, "std")
                        Case ""
                            chptBase_PageNumbering_Set(sect, False, 1, "std")
                        Case "tag_appendixPart"
                            sectPrior = myDoc.Sections.Item(sect.Index - 1)
                            If Me.pgNum_is_DivPageNumberFormat(sectPrior) Then
                                chptBase_PageNumbering_Set(sect, False, 1, "div")
                            Else
                                chptBase_PageNumbering_Set(sect, False, 1, "div")
                            End If
                            reachedEnd = True
                            Exit For
                        Case "tag_appendixChapter"
                            reachedEnd = True
                            Exit For
                    End Select
                Next
            End If
            '
            If pgNumStyle = "2part" Then
                For j = divStartIndex + 1 To myDoc.Sections.Last.Index
                    sect = myDoc.Sections.Item(j)
                    strTag = objTagsMgr.tags_get_tagStyleName(sect)
                    Select Case strTag
                        Case "tag_execBanner"
                            chptBase_PageNumbering_Set(sect, True, 1, "es")
                        Case "tag_partBanner", "tag_glossary_Chpt", "tag_biblio_Chpt", "tag_refs_Chpt", "tag_worksCited_Chpt"
                            divCount = divCount + 1
                            chptBase_PageNumbering_Set(sect, True, divCount, "div")
                        Case "tag_chapterBanner"
                            chptBase_PageNumbering_Set(sect, True, 1, "2part")
                        Case ""
                            chptBase_PageNumbering_Set(sect, False, 1, "2part")
                        Case "tag_appendixPart"
                            divCount = divCount + 1
                            sectPrior = myDoc.Sections.Item(sect.Index - 1)
                            If Me.pgNum_is_DivPageNumberFormat(sectPrior) Then
                                chptBase_PageNumbering_Set(sect, False, 1, "div")
                            Else
                                chptBase_PageNumbering_Set(sect, True, divCount, "div")
                            End If
                            reachedEnd = True
                            Exit For
                        Case "tag_appendixChapter"
                            reachedEnd = True
                            Exit For
                    End Select
                Next
            End If
        End If
        '
        '
        If divStartIndex > chptStartIndex Then
            'No divider before first chapter
            'chptBase_PageNumbering_Set(myDoc.Sections.Item(divStartIndex), True, 1, "div")
            'divCount = divCount + 1
            '
            If pgNumStyle = "std" Then
                For j = chptStartIndex To myDoc.Sections.Last.Index
                    sect = myDoc.Sections.Item(j)
                    strTag = objTagsMgr.tags_get_tagStyleName(sect)
                    Select Case strTag
                        Case "tag_execBanner"
                            chptBase_PageNumbering_Set(sect, True, 1, "es")
                        Case "tag_partBanner", "tag_glossary_Chpt", "tag_biblio_Chpt", "tag_refs_Chpt", "tag_worksCited_Chpt"
                            chptBase_PageNumbering_Set(sect, False, 1, "div")
                        Case "tag_chapterBanner"
                            If j = chptStartIndex Then
                                chptBase_PageNumbering_Set(sect, True, 1, "std")
                            Else
                                chptBase_PageNumbering_Set(sect, False, 1, "std")
                            End If
                        Case ""
                            chptBase_PageNumbering_Set(sect, False, 1, "std")
                        Case "tag_appendixPart"
                            sectPrior = myDoc.Sections.Item(sect.Index - 1)
                            If Me.pgNum_is_DivPageNumberFormat(sectPrior) Then
                                chptBase_PageNumbering_Set(sect, False, 1, "div")
                            Else
                                chptBase_PageNumbering_Set(sect, False, 1, "div")
                            End If
                            reachedEnd = True
                            Exit For
                        Case "tag_appendixChapter"
                            reachedEnd = True
                            Exit For
                    End Select
                Next
            End If
            '
            If pgNumStyle = "2part" Then
                divCount = 0
                For j = chptStartIndex To myDoc.Sections.Last.Index
                    sect = myDoc.Sections.Item(j)
                    strTag = objTagsMgr.tags_get_tagStyleName(sect)
                    Select Case strTag
                        Case "tag_execBanner"
                            chptBase_PageNumbering_Set(sect, True, 1, "es")
                        Case "tag_partBanner", "tag_glossary_Chpt", "tag_biblio_Chpt", "tag_refs_Chpt", "tag_worksCited_Chpt"
                            divCount = divCount + 1
                            chptBase_PageNumbering_Set(sect, True, divCount, "div")
                        Case "tag_chapterBanner"
                            chptBase_PageNumbering_Set(sect, True, 1, "2part")
                        Case ""
                            chptBase_PageNumbering_Set(sect, False, 1, "2part")
                        Case "tag_appendixPart"
                            divCount = divCount + 1
                            sectPrior = myDoc.Sections.Item(sect.Index - 1)
                            If Me.pgNum_is_DivPageNumberFormat(sectPrior) Then
                                chptBase_PageNumbering_Set(sect, False, 1, "div")
                            Else
                                chptBase_PageNumbering_Set(sect, True, divCount, "div")
                            End If
                            reachedEnd = True
                            Exit For
                        Case "tag_appendixChapter"
                            reachedEnd = True
                            Exit For
                    End Select
                Next
            End If
        End If
        '
finis:
        objTocMgr.toc_update_TOCs(myDoc)
        '
    End Sub
    '
    Public Function pgNum_is_DivPageNumberFormat(ByRef sect As Word.Section) As Boolean
        Dim rslt As Boolean
        Dim hf As HeaderFooter
        Dim pgNums As Word.PageNumbers
        '
        rslt = False
        hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        pgNums = hf.PageNumbers
        '
        If pgNums.NumberStyle = WdPageNumberStyle.wdPageNumberStyleUppercaseRoman Then rslt = True
        '
        Return rslt
    End Function
    '
    ''' <summary>
    ''' This method will return the section index of the first Chapter Section.
    ''' If no Chapter section is foun
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function pgNum_get_FirstChptIndex(ByRef myDoc As Word.Document) As Integer
        Dim rslt As Integer
        '
        rslt = Me.pgNum_get_FirstIndex(myDoc, "tag_chapterBanner")
        '
        Return rslt
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will return the section index of the first Chapter Section.
    ''' If no Chapter section is foun
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function pgNum_get_FirstDivIndex(ByRef myDoc As Word.Document) As Integer
        Dim rslt As Integer
        '
        rslt = Me.pgNum_get_FirstIndex(myDoc, "tag_partBanner")
        '
        Return rslt
        '
    End Function

    '
    '
    ''' <summary>
    ''' This method will return the section index of the first appearance of the AAC
    ''' report section type as defined in strTagName (examples include "tag_chapterBanner", "tag_partBanner".
    ''' See cTagsMgr for a list of tags.....If no section is found, this method will reurn -1
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="strTagName"></param>
    ''' <returns></returns>
    Public Function pgNum_get_FirstIndex(ByRef myDoc As Word.Document, strTagName As String) As Integer
        Dim objTagsMgr As New cTagsMgr()
        Dim sect As Word.Section
        Dim strTag As String
        Dim j, startIndex As Integer
        '
        startIndex = -1
        '
        'Find the first instance of the sectiopn with tag 'strTagName'
        For j = 1 To myDoc.Sections.Last.Index
            sect = myDoc.Sections.Item(j)
            strTag = objTagsMgr.tags_get_tagStyleName(sect)
            If strTag = strTagName Then
                startIndex = j
                Exit For
            End If
        Next
        '
        Return startIndex
        '
    End Function

End Class

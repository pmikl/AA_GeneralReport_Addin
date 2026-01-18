Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cPlhInfo
    'Public objGlobals As New cGlobals()
    Public objTablesMgr As New cTablesMgr()
    Public lstOfAllPlhs As New List(Of cPlhInfo)
    Public lstOfFloatingPlhs As New List(Of cPlhInfo)
    Public lstOfInLinePlhs As New List(Of cPlhInfo)
    Public tbl As Word.Table
    Public strCaption As String
    Public pageNum As Integer
    Public pageNumAbs As Integer
    Public isFloating As Boolean
    Public isRegular As Boolean
    Public paraCaption As Word.Paragraph
    '
    Public Sub New(ByRef tbl As Word.Table, strObjCaption As String, pgNum As Integer, pgNumAbs As Integer, isFloating As Boolean, ByRef paraCaption As Word.Paragraph)
        Me.strCaption = strObjCaption
        Me.pageNum = pgNum
        Me.pageNumAbs = pgNumAbs
        Me.isFloating = isFloating
        Me.isRegular = Me.PlhDetail_tbl_isRegular(tbl)
        Me.paraCaption = paraCaption
        Me.tbl = tbl
        '
        Me.lstOfAllPlhs.Clear()
        '
    End Sub
    '
    Public Sub New()
        Me.strCaption = ""
        Me.pageNum = 1
        Me.pageNumAbs = 1
        Me.isFloating = False
        Me.isRegular = True
        'Me.isRegular = Me.PlhDetail_tbl_isRegular(tbl)
        Me.paraCaption = Nothing
        Me.tbl = Nothing
        '
        Me.lstOfAllPlhs.Clear()
        '
    End Sub
    '
    Public Function PlhDetail_tbl_isRegular(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        '
        If objTablesMgr.glb_tbls_isRegularByCol(tbl) And objTablesMgr.glb_tbls_isRegularByRow(tbl) Then rslt = True
        '
        Return rslt
        '
    End Function
    '
    Public Sub PlhDetail_goto_object()
        Dim objGlobals As New cGlobals()
        Dim sel As Word.Selection
        Dim rng As Word.Range
        '
        sel = objGlobals.glb_get_wrdSel
        '
        Try
            'rng = Me.paraCaption.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            'rng.Select()
            'sel.GoTo(WdGoToItem.wdGoToPage,, Me.pageNumAbs)
            rng = Me.tbl.Range
            rng.Select()
            '
        Catch ex As Exception
            rng = Me.paraCaption.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            rng.Select()
            '
        End Try
    End Sub
    '

    Public Sub PlhDetail_convert_inline()
        Dim rng As Word.Range
        Dim objTablesMgr As New cTablesMgr()
        '
        'sel = Me.objGlobals.glb_get_wrdSel
        Try
            'sel.GoTo(WdGoToItem.wdGoToPage,, Me.pageNumAbs)
            rng = tbl.Range
            'rng.Select()
            'objTablesMgr.tbl_convert_toInLineAAC(Me.tbl)
            objTablesMgr.tbl_convert_toInLine(Me.tbl)
            Me.isFloating = False
            '
        Catch ex As Exception

        End Try
    End Sub
    '
    ''' <summary>
    ''' This method will return a list of objects (List(Of cPlhInfo)) that describe the characteristics of
    ''' all 'informational' (as opposed to structural) tables in the the document myDoc
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Function PlhDetail_get_allDetails(ByRef myDoc As Word.Document) As List(Of cPlhInfo)
        Dim objPlhInfo As cPlhInfo
        Dim lstOfPlhs As New List(Of cPlhInfo)
        Dim objTools As New cTools()
        Dim styl As Word.Style
        '
        Dim tbl As Word.Table
        Dim rng As Word.Range
        Dim pageNum, pageNumAbs As Integer
        Dim numHeaderRows, leftIndent, cellPadding, leftIndentBody, bodyWidth As Single
        '
        Dim strCaption, strCaptionNew As String
        Dim para As Word.Paragraph
        Dim kount As Integer
        Dim isFloating As Boolean
        Dim strTagStyle As String
        '
        strCaption = ""
        strCaptionNew = ""
        para = Nothing
        kount = 0
        numHeaderRows = 0
        leftIndent = 0
        cellPadding = 0
        leftIndentBody = 0
        bodyWidth = 0
        '
        lstOfPlhs.Clear()
        '
        For Each tbl In myDoc.Tables
            rng = tbl.Range
            pageNum = rng.Information(WdInformation.wdActiveEndAdjustedPageNumber)
            pageNumAbs = rng.Information(WdInformation.wdActiveEndPageNumber)
            '
            'Me.objTablesMgr.tbl_get_tableProperties(tbl, numHeaderRows, leftIndent, cellPadding, leftIndentBody, bodyWidth)
            isFloating = Me.objTablesMgr.tbl_is_Floating(tbl)
            '
            strTagStyle = objTools.tools_tbls_getFirstCellStyleName(tbl)

            strTagStyle = objTablesMgr.tbl_get_tagStyle(tbl)
            objPlhInfo = Nothing
            '
            Select Case strTagStyle
                Case "Cp Report Date", "tag_contactsPage-Front", "tag_execBanner", "tag_chapterBanner", "tag_glossary_Chpt", "tag_biblio_Chpt", "tag_refs_Chpt", "tag_appendixChapter", "tag_contactsPage-Back"
                    '
                Case "tag_partBanner", "tag_appendixPart"

                Case Else
                    '
                    'If the style Caption is in the table then we have a standard AAC placeholder with an
                    'embedded caption in the top row
                    '
                    With tbl.Range.Find()
                        .Style = "Caption"
                        .Execute()
                        If .Found Then
                            'We have a table with an internal caption
                            strCaption = objTablesMgr.tbl_getTblCaption_AACPlaceHolder(tbl, para)
                        Else
                            'We have a table that has an external Caption... Or we just have a table. So we
                            'need to test the style before we get strCaption
                            strCaption = objTablesMgr.tbl_getTblCaption_AACTable(tbl, para)
                            styl = para.Style
                            If Not (styl.NameLocal = "Caption") Then strCaption = "Table - unknown"
                        End If
                    End With
                    '
                    If strTagStyle Like "Emphasis*" Then strCaption = "Emphasis"
                    '
                    If isFloating Then
                        Select Case strCaption
                            Case ""
                                strCaption = "unknown - floating"
                            Case Else
                                strCaption = strCaption + "- floating"
                        End Select
                    End If
                    ' 
                    If strCaption = "" Then strCaption = "unknown"
                    '
                    objPlhInfo = New cPlhInfo(tbl, strCaption, pageNum, pageNumAbs, isFloating, para)
                    lstOfPlhs.Add(objPlhInfo)
                    '
            End Select
            '
nextItem:

        Next
        '
        Return lstOfPlhs
    End Function
    '
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function PlhDetail_get_allFloatingPlhs(ByRef myDoc As Word.Document) As List(Of cPlhInfo)
        Dim objPlhInfo As cPlhInfo
        Dim lstOfFloatingPlhs As New List(Of cPlhInfo)
        '
        lstOfFloatingPlhs.Clear()
        '
        For Each objPlhInfo In Me.lstOfAllPlhs
            If objPlhInfo.isFloating Then
                lstOfFloatingPlhs.Add(objPlhInfo)
            End If
        Next
        '
        Return lstOfFloatingPlhs
        '
    End Function
    '
    '
    Public Function PlhDetail_get_allInLinePlhs(ByRef myDoc As Word.Document) As List(Of cPlhInfo)
        Dim objPlhInfo As cPlhInfo
        Dim lstOfInLinePlhs As New List(Of cPlhInfo)
        '
        lstOfFloatingPlhs.Clear()
        '
        For Each objPlhInfo In Me.lstOfAllPlhs
            If Not objPlhInfo.isFloating Then
                lstOfInLinePlhs.Add(objPlhInfo)
            End If
        Next
        '
        Return lstOfInLinePlhs
        '
    End Function

End Class

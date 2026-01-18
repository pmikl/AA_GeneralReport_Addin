Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word
Public Class cHeaderFooterMgr
    Public objGlobals As cGlobals
    Public objStylesMgr As New cStylesManager()
    Public Sub New()
        Me.objGlobals = New cGlobals()
        '
    End Sub
    '
    ''' <summary>
    ''' This method will return a list of tag styles, ordered by section
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function hf_getTagStyleMap_Allx(ByRef myDoc As Word.Document) As String
        Dim strRslt As String

        strRslt = ""
        For Each sect In myDoc.Sections
            strRslt = strRslt + "Section " + sect.Index.ToString() + "   =" + vbTab + Me.hf_tags_getTagStyleName(sect, "primaryOrFirstPage") + vbCrLf
        Next
        '
        Return strRslt
        '
    End Function
    '
    ''' <summary>
    ''' This method will return a list of tag styles, ordered by section. This overload returns a collection
    ''' of strings
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function hf_getTagStyleMap_All(ByRef myDoc As Word.Document) As Collection
        Dim lstRslt As New Collection()
        Dim sect As Word.Section
        Dim strRslt As String
        Dim j As Integer
        '
        strRslt = ""
        For j = 1 To myDoc.Sections.Count
            sect = myDoc.Sections.Item(j)
            strRslt = "Section " + sect.Index.ToString() + "   =" + vbTab + Me.hf_tags_getTagStyleName(sect, "primaryOrFirstPage")
            lstRslt.Add(strRslt, CStr(j))
        Next
        '
        Return lstRslt
        '
    End Function
    '

    '
    ''' <summary>
    ''' This method will return a collection accessed by 'cpg', 'cnf' and 'toc'.. Locating the
    ''' section which contains the 'cover page', 'contacts page front' and the 'table of contents'
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function hf_getTagStyleMap_CpContactsFrontAndTOC(ByRef myDoc As Word.Document) As Collection
        Dim lst As New Collection()
        Dim objBnr As New cChptBanner()
        Dim sect As Word.Section
        Dim strRslt As String
        Dim strTagStyle As String

        strRslt = ""
        For Each sect In myDoc.Sections
            strTagStyle = Me.hf_tags_getTagStyleName(sect, "primaryOrFirstPage")
            Select Case strTagStyle
                Case objBnr.bnr_get_tagStyles(objBnr.tag_coverPage)
                    lst.Add(sect.Index, "cpg")
                Case objBnr.bnr_get_tagStyles(objBnr.tag_cont_Front)
                    lst.Add(sect.Index, "cnf")
                Case objBnr.bnr_get_tagStyles(objBnr.tag_toc)
                    lst.Add(sect.Index, "toc")
            End Select

        Next
        '
        Return lst
        '
    End Function
    '
    ''' <summary>
    ''' This method will delete all items in the Footer, including any footer tables
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub hf_footers_delete(ByRef sect As Word.Section)
        '
        sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Delete()
        sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Range.Delete()
        sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterEvenPages).Range.Delete()
        '
    End Sub
    '
    ''' <summary>
    ''' This method will delete all items in the Header, including any header tables
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub hf_headers_delete(ByRef sect As Word.Section)
        '
        sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary).Range.Delete()
        sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Range.Delete()
        sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterEvenPages).Range.Delete()
        '
    End Sub
    '
    ''' <summary>
    ''' This method will delete all shapes and text in the Header Table, but it will leave the table
    ''' intact
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub hf_headers_DeleteContents_All(ByRef sect As Section)
        Dim hf As Word.HeaderFooter

        Try
            For Each hf In sect.Headers
                If hf.Exists Then Me.hf_Hfs_DeleteText(hf, True)
            Next
        Catch ex As Exception

        End Try
    End Sub
    '
    Public Sub hf_headers_DeleteContents_Text(ByRef sect As Section)
        Dim hf As Word.HeaderFooter

        Try
            For Each hf In sect.Headers
                If hf.Exists Then Me.hf_Hfs_DeleteText(hf, False)
            Next
        Catch ex As Exception

        End Try
    End Sub
    '
    ''' <summary>
    ''' This method will delte all text from the footer Table, but will leave the Page Number intact
    ''' and the table intact
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub hf_footers_DeleteContents_Text(ByRef sect As Section)
        Dim hf As Word.HeaderFooter

        Try
            For Each hf In sect.Footers
                If hf.Exists Then Me.hf_Hfs_DeleteText(hf, False, False)
            Next
        Catch ex As Exception

        End Try
    End Sub
    '
    ''' <summary>
    ''' This method will delte all text and page number from the footer Table.
    ''' The table will be left intact
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub hf_footers_DeleteContents_All(ByRef sect As Section)
        Dim hf As Word.HeaderFooter

        Try
            For Each hf In sect.Footers
                If hf.Exists Then Me.hf_Hfs_DeleteText(hf, False, True)
            Next
        Catch ex As Exception

        End Try
    End Sub

    Private Sub hf_Hfs_DeleteText(ByRef hf As Word.HeaderFooter, Optional doRemoveHeaderShapes As Boolean = False, Optional doRemoveFooterPageNum As Boolean = False)
        Dim rng As Range
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        '
        Try
            If hf.IsHeader Then
                '
                rng = hf.Range
                If rng.ShapeRange.Count <> 0 And doRemoveHeaderShapes Then rng.ShapeRange.Delete()
                '
                'Do the Header Table
                If hf.Range.Tables.Count <> 0 Then
                    tbl = hf.Range.Tables.Item(1)
                    For Each drCell In tbl.Range.Cells
                        drCell.Range.Text = ""
                    Next
                End If
                '
            Else
                'Do the Footer Table
                If hf.Range.Tables.Count <> 0 Then
                    tbl = hf.Range.Tables.Item(1)
                    If doRemoveFooterPageNum Then
                        For Each drCell In tbl.Range.Cells
                            drCell.Range.Text = ""
                        Next
                    Else
                        For Each drCell In tbl.Range.Cells
                            If drCell.ColumnIndex <> tbl.Columns.Count Then drCell.Range.Text = ""
                        Next
                    End If

                End If
            End If
        Catch ex As Exception

        End Try
    End Sub
    '

    '
    ''' <summary>
    ''' This method will delete the contact text Boxes in the letter
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function hf_footers_LetterContacts_Clear(ByRef sect As Section) As Word.Range
        Dim hf As Word.HeaderFooter
        Dim rng, rngCell As Word.Range
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        Dim cellStyle As Word.Style
        '
        rngCell = Nothing
        Try
            hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
            rng = hf.Range
            If rng.Tables.Count <> 0 Then
                tbl = rng.Tables.Item(1)
                drCell = tbl.Range.Cells.Item(1)
                cellStyle = drCell.Range.Style
                If cellStyle.NameLocal = "Footer (Letter)" Then
                    rngCell = drCell.Range
                    rngCell.Text = ""
                End If
            End If
            '
        Catch ex As Exception
            rngCell = Nothing
        End Try
        '
        Return rngCell
        '
    End Function
    '

#Region "HFs Link Unlink"

    ''' <summary>
    ''' This method will go through the document myDoc section by section (last to first)
    ''' and unlink all Headers and Footers
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub hf_hfs_UnlinkAllSections(ByRef myDoc As Word.Document)
        Dim sect As Word.Section
        Dim j As Integer
        '
        Try
            For j = myDoc.Sections.Last.Index To 1 Step -1
                sect = myDoc.Sections.Item(j)
                Me.hf_hfs_linkUnlinkAll(sect, False)
            Next
            '
        Catch ex As Exception
            MsgBox("Error in cHeaderFooterMgr.hf_hfs_UnlinkAllSections")
        End Try
        '
    End Sub

    Public Sub hf_hfs_deleteAll(ByRef sect As Word.Section)
        Me.hf_headers_delete(sect)
        Me.hf_footers_delete(sect)
    End Sub
    '
    Public Sub hf_hfs_linkUnlinkAll(ByRef sect As Word.Section, linkToPrevious As Boolean)
        Me.hf_headers_linkUnlink(sect, linkToPrevious)
        Me.hf_footers_linkUnlink(sect, linkToPrevious)
    End Sub

    '
    Public Sub hf_footers_linkUnlink(ByRef sect As Word.Section, linkToPrevious As Boolean)
        '
        'Don't want to do this if there are no previous sections
        If Not sect.Index = 1 Then
            sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = linkToPrevious
            sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage).LinkToPrevious = linkToPrevious
            sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterEvenPages).LinkToPrevious = linkToPrevious
        End If
        '
    End Sub

    '
    Public Sub hf_headers_linkUnlink(ByRef sect As Word.Section, linkToPrevious As Boolean)
        '
        'Don't want to do this if there are no previous sections
        Try
            If Not sect.Index = 1 Then
                sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary).LinkToPrevious = linkToPrevious
                sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage).LinkToPrevious = linkToPrevious
                sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterEvenPages).LinkToPrevious = linkToPrevious
            End If

        Catch ex As Exception

        End Try
        '
    End Sub
    '
#End Region

#Region "Clone HeaderFooter"
    '
    ''' <summary>
    ''' This method will copy the Source HeaderFooter (hfSrc) to the Destination headerFooter (hfDest).
    ''' It will also copy the page numbering setting if doPageNumbering is true
    ''' </summary>
    ''' <param name="hfSrc"></param>
    ''' <param name="hfDst"></param>
    ''' <param name="doPageNumbering"></param>
    Public Sub hf_HF_CopySourceToDestination(ByRef hfSrc As Word.HeaderFooter, ByRef hfDst As Word.HeaderFooter, doPageNumbering As Boolean)
        'Get and paste the header or footer... This approach doesn't assume a table
        'and it will copy across shapes etc..?
        Dim rngSrc, rngDest As Word.Range
        '
        rngSrc = hfSrc.Range
        rngSrc.Copy()
        '
        rngDest = hfDst.Range
        rngDest = rngDest.Paragraphs(1).Range
        rngDest.MoveEnd(WdUnits.wdParagraph)
        rngDest.Paste()
        rngDest.Collapse(WdCollapseDirection.wdCollapseEnd)
        rngDest.Paragraphs(1).Range.Delete()
        '
        If doPageNumbering Then Call Me.hf_hfs_Copy_SectionPageNumbering(hfSrc, hfDst)
        '
    End Sub
    '
    ''' <summary>
    ''' This method will copy the header or footer from the source section
    ''' to the destination section.. It is currently limited to sections
    ''' with "same first page"
    ''' </summary>
    ''' <param name="strHeaderOrFooter"></param>
    ''' <param name="srcSection"></param>
    ''' <param name="destSection"></param>
    Public Sub hf_HF_CopyHeaderFooter(strHeaderOrFooter As String, ByRef srcSection As Section, ByRef destSection As Section)
        '
        Dim hfSrc As HeaderFooter
        Dim hfDst As HeaderFooter
        Dim strHFType As String
        '
        Me.hf_hfs_linkUnlinkAll(destSection, False)
        hfDst = Nothing
        hfSrc = Nothing
        strHFType = ""
        '
        'Define the source and destination Header/Footers
        Select Case strHeaderOrFooter
            Case "header"
                Call Me.hf_headers_Delete_Contents_All(destSection)
                '
                strHFType = Me.hf_get_HeaderFooterType(srcSection)
                '
                Select Case strHFType
                    Case "DiffFirstPage-Not"
                        hfSrc = srcSection.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        hfDst = destSection.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        '
                        Me.hf_HF_CopySourceToDestination(hfSrc, hfDst, True)
                        '
                    Case "DiffFirstPage"
                        hfSrc = srcSection.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                        hfDst = destSection.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                        '
                        Me.hf_HF_CopySourceToDestination(hfSrc, hfDst, True)
                        '
                        hfSrc = srcSection.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        hfDst = destSection.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        '
                        Me.hf_HF_CopySourceToDestination(hfSrc, hfDst, True)
                        '
                    Case "OddAndEven"
                    Case "DiffFirstPage+OddAndEven"

                End Select

            Case "footer"
                Call Me.hf_footers_Delete_Contents_All(destSection)
                '
                strHFType = Me.hf_get_HeaderFooterType(srcSection)
                '
                Select Case strHFType
                    Case "DiffFirstPage-Not"
                        hfSrc = srcSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        hfDst = destSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        '
                        Me.hf_HF_CopySourceToDestination(hfSrc, hfDst, True)
                        '
                    Case "DiffFirstPage"
                        hfSrc = srcSection.Footers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                        hfDst = destSection.Footers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                        '
                        Me.hf_HF_CopySourceToDestination(hfSrc, hfDst, True)
                        '
                        hfSrc = srcSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        hfDst = destSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        '
                        Me.hf_HF_CopySourceToDestination(hfSrc, hfDst, True)
                        '
                    Case "OddAndEven"
                    Case "DiffFirstPage+OddAndEven"
                        '

                End Select

                hfSrc = srcSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                hfDst = destSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        End Select
        '
    End Sub
    '
    '
    Public Sub hf_hfs_Copy_SectionPageNumbering(ByRef hfSource As HeaderFooter, ByRef hfDest As HeaderFooter)
        'This method will copy the section numbering scheme associated with
        'the source Footer to the destination footer
        '
        hfDest.PageNumbers.NumberStyle = hfSource.PageNumbers.NumberStyle
        hfDest.PageNumbers.RestartNumberingAtSection = hfSource.PageNumbers.RestartNumberingAtSection
        hfDest.PageNumbers.StartingNumber = hfSource.PageNumbers.StartingNumber
        '
        'If we try to run the methods in the if statement for a section
        'that doesn not include the chapter number we'll get a fault, so these
        'items must be conditional
        If hfSource.PageNumbers.IncludeChapterNumber Then
            hfDest.PageNumbers.IncludeChapterNumber = hfSource.PageNumbers.IncludeChapterNumber
            hfDest.PageNumbers.ChapterPageSeparator = hfSource.PageNumbers.ChapterPageSeparator
            hfDest.PageNumbers.HeadingLevelForChapter = hfSource.PageNumbers.HeadingLevelForChapter
        End If

    End Sub
    '



#End Region
    '
    ''' <summary>
    ''' This method will return the left and right edges of the specified table (strHeaderType = 'primary', 'firstPage', 'evenPage').
    ''' The result is returned in a Collection of type 'Single'.. The Key 'leftEdge' will return the leftEdge of the Table relative
    ''' to the left Edge of the page and the Key 'rightEdge' will return the rightEdge of the Table relative to the right edge
    ''' of the Table... If doFromTable is true then the measuremenst are obtained from the standard set in cGlobals
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="doDefaultSettings"></param>
    ''' <returns></returns>
    Public Function hf_hfs_getHfTableEdges(ByRef hf As Word.HeaderFooter, Optional doDefaultSettings As Boolean = True) As Collection
        Dim tbl As Word.Table
        Dim sect As Word.Section
        Dim rightEdge, leftEdge, tblWidth As Single
        Dim lst As New Collection()
        '
        sect = hf.Range.Sections.Item(1)
        tblWidth = 0.0
        '
        Try
            If Not doDefaultSettings Then
                'Get actual measurements
                tbl = Me.hf_Hfs_getTable(hf, tblWidth)
                If Not IsNothing(tbl) Then
                    leftEdge = sect.PageSetup.LeftMargin + tbl.Rows.Item(1).LeftIndent          'Remember leftIndent is negative if its to the left of the left margin
                    rightEdge = sect.PageSetup.PageWidth - (tblWidth + leftEdge)
                    '
                    lst.Add(leftEdge, "leftEdge")
                    lst.Add(rightEdge, "rightEdge")
                    '
                Else
                    lst.Clear()
                End If
                '
            Else
                If hf.IsHeader Then
                    leftEdge = Me.objGlobals.glb_hfs_getHFTableEdge(sect, "header_leftEdge")
                    rightEdge = Me.objGlobals.glb_hfs_getHFTableEdge(sect, "header_rightEdge")
                Else
                    leftEdge = Me.objGlobals.glb_hfs_getHFTableEdge(sect, "footer_leftEdge")
                    rightEdge = Me.objGlobals.glb_hfs_getHFTableEdge(sect, "footer_rightEdge")
                End If
                '
                lst.Add(leftEdge, "leftEdge")
                lst.Add(rightEdge, "rightEdge")
                '
            End If
        Catch ex As Exception
            lst.Clear()
        End Try
        '
        Return lst
        '
    End Function
    '

    ''' <summary>
    ''' This method will return the left and right edges of the specified table (strHeaderType = 'primary', 'firstPage', 'evenPage').
    ''' The result is returned in a Collection of type 'Single'.. The Key 'leftEdge' will return the leftEdge of the Table relative
    ''' to the left Edge of the page and the Key 'rightEdge' will return the rightEdge of the Table relative to the right edge
    ''' of the Table
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <returns></returns>
    Public Function xhf_headers_getTableEdges(ByRef hf As Word.HeaderFooter) As Collection
        Dim tbl As Word.Table
        Dim sect As Word.Section
        Dim rightEdge, leftEdge, tblWidth As Single
        Dim lst As New Collection()
        '
        sect = hf.Range.Sections.Item(1)
        tblWidth = 0.0
        '
        Try
            tbl = Me.hf_Hfs_getTable(hf, tblWidth)
            '
            leftEdge = sect.PageSetup.LeftMargin + tbl.Rows.Item(1).LeftIndent          'Remember leftIndent is negative if its to the left of the left margin
            rightEdge = sect.PageSetup.PageWidth - (tblWidth + leftEdge)
            '
            lst.Add(leftEdge, "leftEdge")
            lst.Add(rightEdge, "rightEdge")
            '
        Catch ex As Exception
            lst.Clear()
        End Try
        '
        Return lst
        '
    End Function
    '
    ''' <summary>
    ''' This method will return the first table in the HeaderFooter, hf. If there is no table it will
    ''' return 'nothing'. It will also return the width of the table in the refrerenced variable tblWidth
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="tblWidth"></param>
    ''' <returns></returns>
    Public Function hf_Hfs_getTable(ByRef hf As Word.HeaderFooter, ByRef tblWidth As Single) As Word.Table
        Dim tbl As Word.Table
        Dim rng As Word.Range
        Dim drCol As Word.Column

        '
        tbl = Nothing
        tblWidth = 0.0
        '
        Try
            rng = hf.Range
            If rng.Tables.Count <> 0 Then
                tbl = rng.Tables.Item(1)
                For Each drCol In tbl.Columns
                    tblWidth = tblWidth + drCol.Width
                Next
            End If

        Catch ex As Exception
            tbl = Nothing
            tblWidth = 0.0
        End Try
        '
        Return tbl

    End Function
    '
    ''' <summary>
    ''' This method will return the first table in the select header type (strHeaderType). If there is no table it will
    ''' return 'nothing'. The input 'strHeaderType' can be 'primary', 'firstPage' or 'evenPage'
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="strHeaderType"></param>
    ''' <returns></returns>
    Public Function xhf_headers_getTable(ByRef sect As Word.Section, Optional strHeaderType As String = "primary") As Word.Table
        Dim tbl As Word.Table
        Dim hf As Word.HeaderFooter
        Dim rng As Word.Range
        '
        tbl = Nothing
        Try
            Select Case strHeaderType
                Case "primary"
                    hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                    rng = hf.Range
                    If rng.Tables.Count <> 0 Then
                        tbl = rng.Tables.Item(1)
                    End If
                Case "firstPage"
                Case "evenPage"
            End Select

        Catch ex As Exception

        End Try
        '
        Return tbl

    End Function
    '
    ''' <summary>
    ''' This method will resize all existing Header Tables. If 'doDefaultSettings' is true then the position for the resized Header Table
    ''' (left and right edges) is taken from original Header Table. If it is false the left and right edges are taken from cGlobals
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="lstOfEdges"></param>
    Public Sub hf_headers_resize_all(ByRef sect As Word.Section, Optional ByRef lstOfEdges As Collection = Nothing)
        Dim objTools As New cTools()
        Dim tblWidth, tblWidth_new As Single
        Dim hf As Word.HeaderFooter
        '
        'Me.hf_headers_delete(sect)
        tblWidth = 0.0
        tblWidth_new = 0.0
        '
        'Me.hf_headers_delete(sect)
        tblWidth = 0.0
        '
        '***
        '
        For Each hf In sect.Headers
            If hf.Exists() Then
                Me.hf_headers_resize_one(hf, lstOfEdges)
            End If
        Next
        '
    End Sub

    ''' <summary>
    ''' This method will resize the Header Footer table to match the left and right edges in 'lstOfEdges'
    ''' as obtained from me.hf_headers_getTableEdges with 'doDefaultSettings' set to True.. If lstOfEdges
    ''' is nothing or has no elements then lstOfEdges is filled with the Default settings
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="lstOfEdges"></param>
    Private Sub hf_headers_resize_one(ByRef hf As Word.HeaderFooter, ByRef lstOfEdges As Collection)
        'Dim lstOfEdges As Collection
        Dim objTools As New cTools()
        Dim tblWidth, tblWidth_new, deltaWidth, deltaColWidth, runningWidth As Single
        Dim leftEdge, rightEdge, leftIndent As Single
        Dim drColCount As Integer
        Dim tbl As Word.Table
        Dim dr As Word.Row
        Dim drCol As Word.Column
        Dim rng As Word.Range
        Dim shp As Word.Shape
        Dim sect As Word.Section
        '
        sect = hf.Range.Sections.Item(1)
        'Me.hf_headers_delete(sect)
        tblWidth = 0.0
        tblWidth_new = 0.0
        '
        'If isnothing, the second test doesn't work
        '
        'We use the supplied list of Edges (lstOfEdges) if the supplied lstOfEdges collection
        'is either nothing, or has no elements. In which case we default to the standard set
        'as defined in hf_headers_getTableEdges
        If IsNothing(lstOfEdges) Then
            lstOfEdges = Me.hf_hfs_getHfTableEdges(hf, True)
        ElseIf lstOfEdges.Count = 0 Then
            lstOfEdges = Me.hf_hfs_getHfTableEdges(hf, True)
        End If
        '
        tbl = Me.hf_Hfs_getTable(hf, tblWidth)
        '
        If Not IsNothing(tbl) Then
            Select Case hf.Index
                Case WdHeaderFooterIndex.wdHeaderFooterFirstPage
                    If lstOfEdges.Count <> 0 Then
                        '
                        leftEdge = CSng(lstOfEdges.Item("leftEdge"))
                        rightEdge = CSng(lstOfEdges.Item("rightEdge"))
                        '
                        tblWidth_new = sect.PageSetup.PageWidth - leftEdge - rightEdge
                        'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
                        'tbl.PreferredWidth = tblWidth
                        '
                        'tbl.Rows.LeftIndent = -(sect.PageSetup.LeftMargin - leftEdge)

                        'tbl = Me.hf_Hfs_getTable(hf, tblWidth)
                        deltaWidth = tblWidth_new - tblWidth
                        leftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
                        For Each dr In tbl.Rows
                            dr.LeftIndent = leftIndent
                        Next
                        '
                        drColCount = tbl.Columns.Count
                        deltaColWidth = deltaWidth / drColCount
                        '
                        runningWidth = 0.0
                        For Each drCol In tbl.Columns
                            If drCol.Index <> drColCount Then
                                drCol.Width = drCol.Width + deltaColWidth
                                runningWidth = runningWidth + drCol.Width
                            Else
                                drCol.Width = tblWidth_new - runningWidth
                            End If
                        Next
                    End If

                Case WdHeaderFooterIndex.wdHeaderFooterPrimary
                    If lstOfEdges.Count <> 0 Then
                        '
                        leftEdge = CSng(lstOfEdges.Item("leftEdge"))
                        rightEdge = CSng(lstOfEdges.Item("rightEdge"))
                        '
                        '*** Should work, but doesn't
                        'tbl = Me.hf_Hfs_getTable(hf, tblWidth)

                        'tblWidth_new = sect.PageSetup.PageWidth - leftEdge - rightEdge
                        'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
                        'tbl.PreferredWidth = tblWidth_new
                        'tbl.PreferredWidth = tblWidth_new
                        '
                        'tbl.Rows.LeftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
                        '
                        'GoTo loop1
                        '
                        'hf_objGlobals.glb_screen_update(True)

                        tblWidth_new = sect.PageSetup.PageWidth - leftEdge - rightEdge
                        'tbl = Me.hf_Hfs_getTable(hf, tblWidth)
                        deltaWidth = tblWidth_new - tblWidth
                        leftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
                        For Each dr In tbl.Rows
                            dr.LeftIndent = leftIndent
                        Next
                        '
                        drColCount = tbl.Columns.Count
                        deltaColWidth = deltaWidth / drColCount
                        '
                        runningWidth = 0.0
                        For Each drCol In tbl.Columns
                            If drCol.Index <> drColCount Then
                                drCol.Width = drCol.Width + deltaColWidth
                                runningWidth = runningWidth + drCol.Width
                            Else
                                drCol.Width = tblWidth_new - runningWidth
                            End If
                        Next
                        '
                    End If
                        '
                Case WdHeaderFooterIndex.wdHeaderFooterEvenPages
                    If lstOfEdges.Count <> 0 Then
                        '
                        leftEdge = CSng(lstOfEdges.Item("leftEdge"))
                        rightEdge = CSng(lstOfEdges.Item("rightEdge"))
                        '
                        tblWidth_new = sect.PageSetup.PageWidth - leftEdge - rightEdge
                        'tbl = Me.hf_Hfs_getTable(hf, tblWidth)
                        deltaWidth = tblWidth_new - tblWidth
                        leftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
                        For Each dr In tbl.Rows
                            dr.LeftIndent = leftIndent
                        Next
                        '
                        drColCount = tbl.Columns.Count
                        deltaColWidth = deltaWidth / drColCount
                        '
                        runningWidth = 0.0
                        For Each drCol In tbl.Columns
                            If drCol.Index <> drColCount Then
                                drCol.Width = drCol.Width + deltaColWidth
                                runningWidth = runningWidth + drCol.Width
                            Else
                                drCol.Width = tblWidth_new - runningWidth
                            End If
                        Next
                    End If

                    'Now look for a shape in column 2 and adjust its position
                    rng = tbl.Range.Cells.Item(2).Range
                    If rng.ShapeRange.Count <> 0 Then
                        shp = rng.ShapeRange.Item(1)
                        shp.Left = tbl.Range.Cells.Item(2).Width - shp.Width
                    End If

            End Select

        End If
        '
    End Sub
    '
    ''' <summary>
    ''' This method will resize all existing Footer Tables. If 'doFromTable' is true then the position for the resized Header Table
    ''' (left and right edges) is taken from original Header Table. If it is false the left and right edges are taken from cGlobals
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="lstOfEdges"></param>
    Public Sub hf_footers_resize_all(ByRef sect As Word.Section, Optional ByRef lstOfEdges As Collection = Nothing)
        Dim objTools As New cTools()
        Dim tblWidth, tblWidth_new As Single
        Dim hf As Word.HeaderFooter
        '
        'Me.hf_headers_delete(sect)
        tblWidth = 0.0
        tblWidth_new = 0.0
        '
        For Each hf In sect.Footers
            If hf.Exists() Then
                Me.hf_footers_resize_one(hf, lstOfEdges)
            End If
        Next
        '
    End Sub
    '
    Private Sub hf_footers_resize_one(ByRef hf As Word.HeaderFooter, ByRef lstOfEdges As Collection)
        Dim objTools As New cTools()
        Dim tblWidth, tblWidth_new As Single
        Dim leftEdge, rightEdge As Single
        Dim tbl As Word.Table
        Dim drCol As Word.Column
        Dim sect As Word.Section
        '
        sect = hf.Range.Sections.Item(1)
        'Me.hf_headers_delete(sect)
        tblWidth = 0.0
        tblWidth_new = 0.0
        '
        'We use the supplied list of Edges (lstOfEdges) if the supplied lstOfEdges collection
        'is either nothing, or has no elements. In which case we default to the standard set
        'as defined in hf_headers_getTableEdges
        If IsNothing(lstOfEdges) Then
            lstOfEdges = Me.hf_hfs_getHfTableEdges(hf, True)
        ElseIf lstOfEdges.Count = 0 Then
            lstOfEdges = Me.hf_hfs_getHfTableEdges(hf, True)
        End If
        '
        tbl = Me.hf_Hfs_getTable(hf, tblWidth)
        '
        '
        If IsNothing(tbl) Then GoTo finis
        '
        Try
            Select Case hf.Index
                Case WdHeaderFooterIndex.wdHeaderFooterFirstPage
                    If lstOfEdges.Count <> 0 Then
                        leftEdge = CSng(lstOfEdges.Item("leftEdge"))
                        rightEdge = CSng(lstOfEdges.Item("rightEdge"))
                        '
                        'tblWidth_new = sect.PageSetup.PageWidth - leftEdge - rightEdge
                        tblWidth_new = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - rightEdge         'left edge always at left margin
                        '
                        'tbl = Me.hf_Hfs_getTable(hf, tblWidth)
                        'deltaWidth = tblWidth_new - tblWidth
                        'leftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
                        'For Each dr In tbl.Rows
                        'dr.LeftIndent = leftIndent
                        'Next
                        '
                        'drColCount = tbl.Columns.Count
                        'deltaColWidth = deltaWidth / drColCount
                        '
                        'runningWidth = 0.0
                        'For Each drCol In tbl.Columns
                        'If drCol.Index <> drColCount Then
                        'drCol.Width = drCol.Width + deltaColWidth
                        'runningWidth = runningWidth + drCol.Width
                        'Else
                        'drCol.Width = tblWidth_new - runningWidth
                        'End If
                        'Next
                        drCol = tbl.Columns.Item(1)
                        drCol.Width = objGlobals.glb_get_widthBetweenMargins(sect)
                        drCol = tbl.Columns.Item(2)
                        drCol.Width = tblWidth_new - tbl.Columns.Item(1).Width

                    End If

                Case WdHeaderFooterIndex.wdHeaderFooterPrimary
                    'Assumes that the left edge is always aligned with the left margin. Otherwise
                    'I'll spend too much time here
                    If lstOfEdges.Count <> 0 Then
                        leftEdge = CSng(lstOfEdges.Item("leftEdge"))
                        rightEdge = CSng(lstOfEdges.Item("rightEdge"))
                        '
                        tblWidth_new = sect.PageSetup.PageWidth - leftEdge - rightEdge                         'anything allowed for left edge
                        'tblWidth_new = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - rightEdge         'left edge always at left margin
                        '
                        '
                        'tbl = Me.hf_Hfs_getTable(hf, tblWidth)
                        'deltaWidth = tblWidth_new - tblWidth
                        'leftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
                        'For Each dr In tbl.Rows
                        'dr.LeftIndent = leftIndent
                        'Next
                        '
                        'drColCount = tbl.Columns.Count
                        'deltaColWidth = deltaWidth / drColCount
                        '
                        'runningWidth = 0.0
                        'For Each drCol In tbl.Columns
                        'If drCol.Index <> drColCount Then
                        'drCol.Width = drCol.Width + deltaColWidth
                        'runningWidth = runningWidth + drCol.Width
                        'Else
                        'drCol.Width = tblWidth_new - runningWidth
                        'End If
                        'Next
                        '
                        drCol = tbl.Columns.Item(2)
                        drCol.Width = sect.PageSetup.RightMargin - rightEdge
                        tbl.Columns.Item(1).Width = tblWidth_new - drCol.Width
                        tbl.Rows.LeftIndent = -(sect.PageSetup.LeftMargin - leftEdge)

                        '
                        'drCol = tbl.Columns.Item(1)
                        'drCol.Width = tblWidth_new - tbl.Columns.Item(2).Width
                        '
                        'drCol = tbl.Columns.Item(1)
                        'drCol.Width = objGlobals.glb_get_widthBetweenMargins(sect)
                        'drCol.Width = objGlobals.glb_get_widthBetweenMargins(sect)
                        'drCol = tbl.Columns.Item(2)
                        'drCol.Width = tblWidth_new - tbl.Columns.Item(1).Width

                    End If

                        '
                Case WdHeaderFooterIndex.wdHeaderFooterEvenPages
                    If lstOfEdges.Count <> 0 Then
                        leftEdge = CSng(lstOfEdges.Item("leftEdge"))
                        rightEdge = CSng(lstOfEdges.Item("rightEdge"))
                        '
                        '*** 20231209 adjustment to allow for offset page number
                        'tblWidth_new = sect.PageSetup.PageWidth - leftEdge - rightEdge                         'anything allowed for left edge
                        tblWidth_new = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - rightEdge         'left edge always at left margin
                        '
                        'tbl = Me.hf_Hfs_getTable(hf, tblWidth)
                        'deltaWidth = tblWidth_new - tblWidth
                        'leftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
                        'For Each dr In tbl.Rows
                        'dr.LeftIndent = leftIndent
                        'Next
                        '
                        'drColCount = tbl.Columns.Count
                        'deltaColWidth = deltaWidth / drColCount
                        '
                        'runningWidth = 0.0
                        'For Each drCol In tbl.Columns
                        'If drCol.Index <> drColCount Then
                        'drCol.Width = drCol.Width + deltaColWidth
                        'runningWidth = runningWidth + drCol.Width
                        ' Else
                        'drCol.Width = tblWidth_new - runningWidth
                        'End If
                        'Next
                        drCol = tbl.Columns.Item(1)
                        drCol.Width = objGlobals.glb_get_widthBetweenMargins(sect)
                        drCol = tbl.Columns.Item(2)
                        drCol.Width = tblWidth_new - tbl.Columns.Item(1).Width

                    End If

                    '
                    'Now look for a shape in column 2 and adjust its position
                    'rng = tbl.Range.Cells.Item(2).Range
                    'If rng.ShapeRange.Count <> 0 Then
                    'shp = rng.ShapeRange.Item(1)
                    'shp.Left = tbl.Range.Cells.Item(2).Width - shp.Width
                    'End If

            End Select

        Catch ex As Exception

        End Try
finis:
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will delete the content of all of the headers in the section sect. It will
    ''' leave branding shapes/watermarks intact unless the deletBranding option is
    ''' incldued and set to true
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub hf_headers_Delete_Contents_All(ByRef sect As Section)
        Dim hf As HeaderFooter
        Dim rng As Range
        'Dim tbl As Word.Table
        '
        '****** IMPORTANT
        'We you delete the range of a text object, the style will revert to the
        'default style for that object... So, in the case of a Header, it will
        'revert to the Header style
        '******
        '
        Try
            For Each hf In sect.Headers
                If hf.Exists Then
                    rng = hf.Range
                    If rng.ShapeRange.Count <> 0 Then rng.ShapeRange.Delete()
                    rng.Delete()
                    'At this point the style of the range has reverted to the Header style
                End If
            Next hf
        Catch ex As Exception
            MsgBox("failed in cHeaderFooterMgr.deletHeaders")
        End Try
        '
        'myStyle = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("Normal")
        '
    End Sub
    '
    ''' <summary>
    ''' This method will delete the content of all of the footers in the section sect. It will
    ''' leave branding shapes/watermarks intact unless the deletBranding option is
    ''' incldued and set to true
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub hf_footers_Delete_Contents_All(ByRef sect As Section)
        Dim hf As HeaderFooter
        Dim rng As Range
        '
        '****** IMPORTANT
        'We you delete the range of a text object, the style will revert to the
        'default style for that object... So, in the case of a Header, it will
        'revert to the Footer style
        '******
        '
        Try
            For Each hf In sect.Footers
                If hf.Exists Then
                    rng = hf.Range
                    If rng.ShapeRange.Count <> 0 Then rng.ShapeRange.Delete()
                    rng.Delete()
                    'At this point the style of the range has reverted to the Footer style
                End If
            Next hf
        Catch ex As Exception
            MsgBox("Failed in cHeaderFooterMgr.deleteFooters")
        End Try

        '
    End Sub


    ''' <summary>
    ''' This method will delete the existing header(s) in sect and insert a new set. Note that if
    ''' leftEdge is set to a negative number, then the left edge is flush with the left margin.
    ''' If it is not negative it is set to Me.hf_objGlobals.glb_hfs_getHFTableEdge(sect, "header_leftEdge")
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="leftEdge"></param>
    Public Sub hf_headers_insert(ByRef sect As Word.Section, Optional leftEdge As Single = 0.0, Optional doLogo As Boolean = True, Optional reColourLogo As Boolean = False, Optional rightEdge As Single = -1.0, Optional strFirstCellStyleName As String = "spacer")
        Dim tmp As Single
        Dim hf As Word.HeaderFooter
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        Dim stylFirstCell, stylHeaderCompanyName As Word.Style
        '
        'stylFirstCell = objGlobals.glb_get_wrdActiveDoc.Styles.Item("spacer")
        'stylFirstCell = objGlobals.glb_get_wrdActiveDoc.Styles.Item(strFirstCellStyleName)

        'Me.objStylesMgr.styl_build_Style_Normal(Doc)
        'Me.objStylesMgr.styl_build_Style_NormalNoSpace(Doc)
        'Me.objStylesMgr.styl_build_Style_Spacer(Doc)
        'Me.objStylesMgr.styl_build_Style_BodyText(Doc)



        stylFirstCell = objGlobals.glb_get_wrdActiveDoc.Styles.Item(strFirstCellStyleName)
        stylHeaderCompanyName = objGlobals.glb_get_wrdActiveDoc.Styles.Item("Header-Company Name")

        'stylHeaderCompanyName = objGlobals.glb_get_wrdActiveDoc.Styles.Item("Header-Company Name")
        '
        'Just in case the authors have rolled their own sections we make sure that thye are all unlinked
        Me.hf_headers_linkUnlink(sect, False)
        Me.hf_headers_delete(sect)
        '
        tbl = Nothing
        '
        'The Primary Header is the standard. Everything is measured from this. So lets
        'get the primary left and right edges
        '
        If leftEdge < 0.0 Then
            leftEdge = sect.PageSetup.LeftMargin
        Else
            leftEdge = Me.objGlobals.glb_hfs_getHFTableEdge(sect, "header_leftEdge")
        End If
        '
        If rightEdge < 0.0 Then
            rightEdge = sect.PageSetup.RightMargin
        Else
            rightEdge = Me.objGlobals.glb_hfs_getHFTableEdge(sect, "header_rightEdge")
        End If
        '
        For Each hf In sect.Headers
            If hf.Exists Then
                '
                If hf.LinkToPrevious Then hf.LinkToPrevious = False
                '
                Select Case hf.Index
                    Case WdHeaderFooterIndex.wdHeaderFooterEvenPages
                        'Mirror the left and right edges
                        tmp = leftEdge
                        leftEdge = rightEdge
                        rightEdge = tmp
                        '
                        'doLogo = True
                        tbl = Me.hf_insertEvenPages_header(hf, leftEdge, rightEdge, doLogo, reColourLogo)
                        drCell = tbl.Range.Cells.Item(1)
                        drCell.Range.Style = stylHeaderCompanyName
                        '
                        drCell = tbl.Range.Cells.Item(2)
                        drCell.Range.Style = stylFirstCell

                    Case WdHeaderFooterIndex.wdHeaderFooterFirstPage
                        'doLogo = True
                        tbl = Me.hf_insertFirstPage_Header(hf, leftEdge, rightEdge, doLogo, reColourLogo)
                        drCell = tbl.Range.Cells.Item(1)
                        drCell.Range.Style = stylFirstCell
                        '
                        drCell = tbl.Range.Cells.Item(2)
                        drCell.Range.Style = stylHeaderCompanyName

                    Case WdHeaderFooterIndex.wdHeaderFooterPrimary
                        'MsgBox("Got to Insert Primary header")
                        tbl = Me.hf_insertPrimary_header(hf, leftEdge, rightEdge, doLogo, reColourLogo)
                        drCell = tbl.Range.Cells.Item(1)
                        drCell.Range.Style = stylFirstCell
                        '
                        drCell = tbl.Range.Cells.Item(2)
                        drCell.Range.Style = stylHeaderCompanyName
                        '
                End Select
            End If

        Next
    End Sub
    '
    Public Sub hf_footers_resize(ByRef sect As Word.Section)
        Dim lstOfEdges As Collection
        Dim objTools As New cTools()
        Dim tblWidth, tblWidth_new, deltaWidth, deltaColWidth, runningWidth As Single
        Dim leftEdge, rightEdge, tmpEdge, leftIndent As Single
        Dim drColCount As Integer
        Dim hf As Word.HeaderFooter
        Dim tbl As Word.Table
        Dim dr As Word.Row
        Dim drCol As Word.Column
        Dim rng As Word.Range
        Dim shp As Word.Shape
        '
        'Me.hf_headers_delete(sect)
        tblWidth = 0.0
        tblWidth_new = 0.0
        '
        '
        For Each hf In sect.Footers
            If hf.Exists() Then
                Select Case hf.Index
                    Case WdHeaderFooterIndex.wdHeaderFooterFirstPage
                        lstOfEdges = Me.hf_hfs_getHfTableEdges(hf)
                        If lstOfEdges.Count <> 0 Then
                            leftEdge = CSng(lstOfEdges.Item("leftEdge"))
                            rightEdge = CSng(lstOfEdges.Item("rightEdge"))
                            tblWidth_new = sect.PageSetup.PageWidth - leftEdge - rightEdge
                            tbl = Me.hf_Hfs_getTable(hf, tblWidth)
                            deltaWidth = tblWidth_new - tblWidth
                            leftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
                            For Each dr In tbl.Rows
                                dr.LeftIndent = leftIndent
                            Next
                            '
                            drColCount = tbl.Columns.Count
                            deltaColWidth = deltaWidth / drColCount
                            '
                            runningWidth = 0.0
                            For Each drCol In tbl.Columns
                                If drCol.Index <> drColCount Then
                                    drCol.Width = drCol.Width + deltaColWidth
                                    runningWidth = runningWidth + drCol.Width
                                Else
                                    drCol.Width = tblWidth_new - runningWidth
                                End If
                            Next
                        End If

                    Case WdHeaderFooterIndex.wdHeaderFooterPrimary
                        'lstOfEdges = Me.hf_headers_getTableEdges(hf)
                        'If lstOfEdges.Count <> 0 Then
                        leftEdge = objGlobals.glb_hfs_getHFTableEdge(sect, "footer_leftEdge")
                        If leftEdge < 0.0 Then leftEdge = sect.PageSetup.LeftMargin
                        '
                        rightEdge = objGlobals.glb_hfs_getHFTableEdge(sect, "footer_rightEdge")
                        If rightEdge < 0.0 Then rightEdge = sect.PageSetup.RightMargin
                        '
                        tbl = Me.hf_Hfs_getTable(hf, tblWidth)
                        tblWidth_new = sect.PageSetup.PageWidth - leftEdge - rightEdge
                        deltaWidth = tblWidth_new - tblWidth

                        leftIndent = -(sect.PageSetup.LeftMargin - leftEdge)

                        '
                        For Each dr In tbl.Rows
                            dr.LeftIndent = leftIndent
                        Next
                        '
                        drColCount = tbl.Columns.Count
                        deltaColWidth = deltaWidth / drColCount
                        '
                        tbl.Columns.Item(1).Width = tbl.Columns.Item(1).Width + deltaWidth
                        '
                    Case WdHeaderFooterIndex.wdHeaderFooterEvenPages
                        leftEdge = objGlobals.glb_hfs_getHFTableEdge(sect, "footer_leftEdge")
                        If leftEdge < 0.0 Then leftEdge = sect.PageSetup.LeftMargin
                        '
                        rightEdge = objGlobals.glb_hfs_getHFTableEdge(sect, "footer_rightEdge")
                        If rightEdge < 0.0 Then rightEdge = sect.PageSetup.RightMargin
                        '
                        'If Mirror Margins must swap the edges around
                        If sect.PageSetup.MirrorMargins Then
                            tmpEdge = leftEdge
                            leftEdge = rightEdge
                            rightEdge = tmpEdge
                        End If
                        '
                        tbl = Me.hf_Hfs_getTable(hf, tblWidth)
                        tblWidth_new = sect.PageSetup.PageWidth - leftEdge - rightEdge
                        deltaWidth = tblWidth_new - tblWidth

                        leftIndent = -(sect.PageSetup.RightMargin - leftEdge)

                        '
                        For Each dr In tbl.Rows
                            dr.LeftIndent = leftIndent
                        Next
                        '
                        drColCount = tbl.Columns.Count
                        deltaColWidth = deltaWidth / drColCount
                        '
                        tbl.Columns.Item(2).Width = tbl.Columns.Item(2).Width + deltaWidth
                        '
                        'Now look for a shape in column 2 and adjust its position
                        rng = tbl.Range.Cells.Item(2).Range
                        If rng.ShapeRange.Count <> 0 Then
                            shp = rng.ShapeRange.Item(1)
                            shp.Left = tbl.Range.Cells.Item(2).Width - shp.Width
                        End If

                End Select

            End If
        Next

    End Sub
    '
    '
    Public Sub hf_footers_insert(ByRef sect As Word.Section, Optional ByRef doPageNum As Boolean = True, Optional ByRef doStyleRefs As Boolean = True)
        Dim leftEdge, rightEdge, tmp As Single
        Dim hf As Word.HeaderFooter
        'Dim doPageNum, doStyleRefs As Boolean
        '
        'Just in case the authors role their own sections and leave the header/footers linked
        Me.hf_footers_linkUnlink(sect, False)
        '
        Me.hf_footers_delete(sect)
        '
        'The Primary Header is the standard. Everything is measured from this. So lets
        'get the primary left and right edges
        '
        leftEdge = Me.objGlobals.glb_hfs_getHFTableEdge(sect, "footer_leftEdge")
        rightEdge = Me.objGlobals.glb_hfs_getHFTableEdge(sect, "footer_rightEdge")
        '
        If leftEdge < 0.0 Then leftEdge = sect.PageSetup.LeftMargin
        If rightEdge < 0.0 Then rightEdge = sect.PageSetup.RightMargin
        '
        'doPageNum = True
        'doStyleRefs = True
        '
        For Each hf In sect.Footers
            If hf.Exists Then
                Select Case hf.Index
                    Case WdHeaderFooterIndex.wdHeaderFooterEvenPages
                        'Mirror the left and right edges
                        tmp = leftEdge
                        leftEdge = rightEdge
                        rightEdge = tmp
                        '
                        'doPageNum = True
                        Me.hf_insertEvenPages_Footer(hf, leftEdge, rightEdge, doPageNum, doStyleRefs)
                    Case WdHeaderFooterIndex.wdHeaderFooterFirstPage
                        'doPageNum = False
                        'doStyleRefs = False
                        Me.hf_insertFirstPage_Footer(hf, leftEdge, rightEdge, doPageNum, doStyleRefs)
                    Case WdHeaderFooterIndex.wdHeaderFooterPrimary
                        Me.hf_insertPrimary_Footer(hf, leftEdge, rightEdge, doPageNum, doStyleRefs)
                End Select
            End If

        Next
    End Sub
    '
    '
    Public Function hf_footer_insertLetterFollower_AsSWBuild(ByRef hf As Word.HeaderFooter, insertLogo As Boolean, rgbFooter As Integer, strStationeryType As String) As Word.Table
        Dim tbl As Word.Table
        Dim rng As Word.Range
        Dim drCell As Word.Cell
        Dim dr As Word.Row
        Dim shp As Word.Shape
        Dim objBBMgr As New cBBlocksHandler()
        Dim objBB As Word.BuildingBlock
        Dim leftIndent, tblWidth As Single
        Dim para As Word.Paragraph
        Dim fld As Word.Field
        Dim sect As Word.Section
        '
        sect = hf.Range.Sections.Item(1)
        rng = hf.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        tbl = rng.Tables.Add(rng, 1, 2)
        tbl.AllowAutoFit = False
        tbl.Borders.Enable = False
        '
        drCell = tbl.Range.Cells.Item(1)
        drCell.TopPadding = 4.8
        drCell.BottomPadding = 7.2
        drCell.LeftPadding = 5.2
        drCell.RightPadding = 3.6
        '
        drCell = tbl.Range.Cells.Item(2)
        drCell.TopPadding = 0.0
        drCell.BottomPadding = 0.0
        drCell.LeftPadding = 5.4
        drCell.RightPadding = 5.4
        drCell.RightPadding = 0.0

        '
        dr = tbl.Rows.Item(1)
        dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
        dr.Height = 24.8
        '
        tbl.Columns.Item(2).Width = 326.65
        tbl.Columns.Item(1).Width = 201.6
        '
        tblWidth = tbl.Columns.Item(1).Width + tbl.Columns.Item(2).Width
        leftIndent = -(tblWidth - (sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin))
        dr.LeftIndent = leftIndent
        '
        Try
            'Set the default style for the Footer Table
            tbl.Range.Style = Globals.ThisAddIn.Application.ActiveDocument.Styles("Footer (Letter-Logo)")
            drCell = tbl.Range.Cells.Item(2)
            drCell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
            '
            drCell = tbl.Range.Cells.Item(1)
            drCell.Range.Style = Globals.ThisAddIn.Application.ActiveDocument.Styles("Footer (Letter)")
            drCell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
            '
            rng = drCell.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Text = vbCrLf + "Page "
            rng = drCell.Range
            para = rng.Paragraphs.Item(1)
            rng = para.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            Select Case strStationeryType
                Case "letter"
                    fld = rng.Fields.Add(rng, WdFieldType.wdFieldStyleRef, """LetterRef""", True)
                    rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                    '
                Case "memo", "briefingNote"
                    rng.Text = "Reference: "
                    rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                    fld = rng.Fields.Add(rng, WdFieldType.wdFieldStyleRef, """StationeryRef_Memo""", True)
                    rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                    '
            End Select
            '
            '
            para = drCell.Range.Paragraphs.Item(2)
            rng = para.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Move(WdUnits.wdCharacter, 5)
            rng.Fields.Add(rng, WdFieldType.wdFieldPage)
            rng.Move(WdUnits.wdCharacter, 1)
            rng.Text = " of "
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            rng.Fields.Add(rng, WdFieldType.wdFieldNumPages)
            '
            For Each para In drCell.Range.Paragraphs
                rng = para.Range
                rng.Font.Color = rgbFooter
            Next

        Catch ex As Exception

        End Try
        '
        If insertLogo Then
            'rng = hf.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            '
            objBB = objBBMgr.getBuildingBlockFromDefaultLib("aac_HFs_logoFooter_Small", "headerFooters")
            drCell = tbl.Range.Cells.Item(2)
            rng = drCell.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng = objBB.Insert(rng)
            If rng.ShapeRange.Count <> 0 Then
                shp = rng.ShapeRange.Item(1)
                shp.Left = 310.75
                shp.Top = 3.1
                'shpInline = shp.ConvertToInlineShape()
                'shp = shpInline.ConvertToShape()
                'shp.an
                shp.LockAnchor = True
                'shp.a
                'shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
                'shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
                shp.Height = 16.7
                shp.Width = 14
                'shp.LeftRelative = 527
                'shp.Left = 311.9
                'shp.Left = 527.4

                'shp.Top = 7.1
                'shp.Top = 796.0
                'shp.TopRelative = 796.0
            End If

        End If

        Return tbl
    End Function

    '
    Public Sub xxhf_footers_insert(ByRef sect As Word.Section)
        Dim leftEdge, rightEdge As Single
        Dim hf As Word.HeaderFooter
        Dim tbl As Word.Table
        Dim doMirror As Boolean
        '
        Me.hf_footers_delete(sect)
        '
        'The Primary Header is the standard
        leftEdge = -1.0
        rightEdge = -1
        '
        doMirror = False
        '
        For Each hf In sect.Footers
            If hf.Exists() Then
                Select Case hf.Index
                    Case WdHeaderFooterIndex.wdHeaderFooterFirstPage
                        If Not sect.PageSetup.OddAndEvenPagesHeaderFooter Then
                            'Just do the standard FirstPage header
                            doMirror = False
                            tbl = Me.hf_hfs_tableInsert(hf, leftEdge, rightEdge, doMirror, False)
                            'tbl.Range.Cells.Item(2).Range.Text = "Primary - FirstPage"
                        Else
                            Select Case sect.PageSetup.SectionStart
                                Case WdSectionStart.wdSectionNewPage
                                    'We need to know whether the new page is odd or even.. How to tell
                                    'It may be that we just can't do that...If we have different odd/even
                                    'and different first page, then the section break must be odd or even
                                    'and NOT "Next Page"
                                    '
                                    'I will have it default to odd.. which will cover the majority of situations
                                    '
                                    hf.Range.Text = "New Page section break is not supported... You will need to set it to odd or even"
                                    'Do odd FirstPage header
                                    'doMirror = False
                                    'tbl = Me.hf_headers_tableInsert(hf, leftEdge, rightEdge, doMirror)
                                    'tbl.Range.Cells.Item(2).Range.Text = "Test - FirstPage"

                                Case WdSectionStart.wdSectionOddPage
                                    'Do odd FirstPage header
                                    doMirror = False
                                    tbl = Me.hf_hfs_tableInsert(hf, leftEdge, rightEdge, doMirror, False)
                                    'tbl.Range.Cells.Item(2).Range.Text = "Primary - FirstPage"
                                Case WdSectionStart.wdSectionEvenPage
                                    'Do even FirstPage header
                                    doMirror = True
                                    tbl = Me.hf_hfs_tableInsert(hf, leftEdge, rightEdge, doMirror, False)
                                    'tbl.Range.Cells.Item(1).Range.Text = "EvenPage - FirstPage"
                            End Select

                        End If

                    Case WdHeaderFooterIndex.wdHeaderFooterPrimary
                        doMirror = False
                        tbl = Me.hf_hfs_tableInsert(hf, leftEdge, rightEdge, doMirror, False)
                        'rng = tbl.Range.Cells.Item(2).Range
                        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
                        'rng.Text = "Primary"
                        'tbl.Range.Cells.Item(2).Range.Text = "Primary"

                    Case WdHeaderFooterIndex.wdHeaderFooterEvenPages
                        doMirror = True
                        tbl = Me.hf_hfs_tableInsert(hf, leftEdge, rightEdge, doMirror, False)
                        'rng = tbl.Range.Cells.Item(1).Range
                        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
                        'rng.Text = "EvenPage"
                        'tbl.Range.Cells.Item(1).Range.Text = "EvenPage"

                End Select

            End If
        Next

    End Sub

    '
    '
    ''' <summary>
    ''' This method will insert standard footers in the selected section. Note that it will
    ''' first delete any existing footerss so we don't have overlap of any type. Note that 
    ''' alignWithLeftMargin and alignWithRightMargin apply to how it looks on the Primary.
    ''' The even pages are automatically a mirror, so they take care of themselves
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub xhf_footers_insert(ByRef sect As Word.Section)
        Dim objSectMgr As New cSectionMgr()
        Dim hf As Word.HeaderFooter
        Dim doMirror, alignWithLeftMargin, alighWithRightMargin As Boolean
        Dim tbl As Word.Table
        Dim strSectionType As String
        Dim myDoc As Word.Document
        '
        Me.hf_headers_delete(sect)
        strSectionType = ""
        '
        doMirror = False
        alignWithLeftMargin = True
        alighWithRightMargin = True
        'sect.PageSetup.MirrorMargins
        'sect.PageSetup.OddAndEvenPagesHeaderFooter
        'sect.PageSetup.DifferentFirstPageHeaderFooter
        '
        myDoc = sect.Range.Document
        '
        strSectionType = objSectMgr.sct_get_SectionType(sect)
        '
        Select Case strSectionType
            Case "000"
                '0 mirror margins, 0 different odd and even, 0 different first page
                doMirror = False
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                'tbl = Me.hf_headers_tableInsert(hf, False, False, doMirror)
                tbl = Me.hf_hfs_tableInsert(hf, alignWithLeftMargin, alighWithRightMargin, doMirror)
                tbl.Range.Cells.Item(1).Range.Text = "Primary"

            Case "001"
                '0 mirror margins, 0 different odd and even, 1 different first page
                doMirror = False
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                'tbl = Me.hf_headers_tableInsert(hf, False, False, doMirror)
                tbl = Me.hf_hfs_tableInsert(hf, alignWithLeftMargin, alighWithRightMargin, doMirror)
                tbl.Range.Cells.Item(1).Range.Text = "Primary"
                '
                doMirror = False
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                'tbl = Me.hf_headers_tableInsert(hf, False, False, doMirror)
                tbl = Me.hf_hfs_tableInsert(hf, alignWithLeftMargin, alighWithRightMargin, doMirror)
                tbl.Range.Cells.Item(1).Range.Text = "FirstPage"
                '
            Case "010"
                '0 mirror margins, 1 different odd and even, 0 different first page
                doMirror = False
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                tbl = Me.hf_hfs_tableInsert(hf, alignWithLeftMargin, alighWithRightMargin, doMirror)
                'tbl = Me.hf_headers_tableInsert(hf, False, True, doMirror)
                tbl.Range.Cells.Item(2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
                tbl.Range.Cells.Item(2).Range.Text = "Primary-Odd"
                '
                doMirror = True
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterEvenPages)
                tbl = Me.hf_hfs_tableInsert(hf, alignWithLeftMargin, alighWithRightMargin, doMirror)
                'tbl = Me.hf_headers_tableInsert(hf, True, False, doMirror)
                tbl.Range.Cells.Item(1).Range.Text = "EvenPage"

            Case "011"
                '0 mirror margins, 1 different odd and even, 1 different first page
                doMirror = False
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                tbl = Me.hf_hfs_tableInsert(hf, alignWithLeftMargin, alighWithRightMargin, doMirror)
                tbl.Range.Cells.Item(2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
                tbl.Range.Cells.Item(2).Range.Text = "Primary-Odd"
                '
                doMirror = True
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterEvenPages)
                tbl = Me.hf_hfs_tableInsert(hf, alignWithLeftMargin, alighWithRightMargin, doMirror)
                tbl.Range.Cells.Item(1).Range.Text = "EvenPage"
                '
                Select Case sect.PageSetup.SectionStart
                    Case WdSectionStart.wdSectionOddPage, WdSectionStart.wdSectionNewPage
                        doMirror = False
                        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                        tbl = Me.hf_hfs_tableInsert(hf, alignWithLeftMargin, alighWithRightMargin, doMirror)
                        tbl.Range.Cells.Item(2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
                        tbl.Range.Cells.Item(2).Range.Text = "FirstPage-Odd"
                    Case WdSectionStart.wdSectionEvenPage
                        doMirror = True
                        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                        tbl = Me.hf_hfs_tableInsert(hf, alignWithLeftMargin, alighWithRightMargin, doMirror)
                        tbl.Range.Cells.Item(1).Range.Text = "FirstPage-Even"
                End Select
                '
                '
            Case "100"
            Case "101"
            Case "110"
                '1 mirror margins, 1 different odd and even, 0 different first page
                doMirror = False
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                'tbl = Me.hf_headers_tableInsert(hf, alignWithLeftMargin, alighWithRightMargin, doMirror)
                tbl = Me.hf_hfs_tableInsert(hf, alignWithLeftMargin, alighWithRightMargin, doMirror)
                'tbl = Me.hf_headers_tableInsert(hf, False, True, doMirror)
                tbl.Range.Cells.Item(2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
                tbl.Range.Cells.Item(2).Range.Text = "Primary-Odd"
                '
                doMirror = True
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterEvenPages)
                tbl = Me.hf_hfs_tableInsert(hf, alignWithLeftMargin, alighWithRightMargin, doMirror)
                'tbl = Me.hf_headers_tableInsert(hf, True, False, doMirror)
                tbl.Range.Cells.Item(1).Range.Text = "EvenPage"
            Case "111"
                '1 mirror margins, 1 different odd and even, 1 different first page
                doMirror = False
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                'tbl = Me.hf_headers_tableInsert(hf, alignWithLeftMargin, alighWithRightMargin, doMirror)
                tbl = Me.hf_hfs_tableInsert(hf, alignWithLeftMargin, alighWithRightMargin, doMirror)
                'tbl = Me.hf_headers_tableInsert(hf, False, True, doMirror)
                tbl.Range.Cells.Item(2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
                tbl.Range.Cells.Item(2).Range.Text = "Primary-Odd"
                '
                doMirror = True
                hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterEvenPages)
                tbl = Me.hf_hfs_tableInsert(hf, alignWithLeftMargin, alighWithRightMargin, doMirror)
                'tbl = Me.hf_headers_tableInsert(hf, True, False, doMirror)
                tbl.Range.Cells.Item(1).Range.Text = "EvenPage"
                '
                Select Case sect.PageSetup.SectionStart
                    Case WdSectionStart.wdSectionOddPage, WdSectionStart.wdSectionNewPage
                        doMirror = False
                        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                        tbl = Me.hf_hfs_tableInsert(hf, alignWithLeftMargin, alighWithRightMargin, doMirror)
                        tbl.Range.Cells.Item(2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
                        tbl.Range.Cells.Item(2).Range.Text = "FirstPage-Odd"
                    Case WdSectionStart.wdSectionEvenPage
                        doMirror = True
                        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                        tbl = Me.hf_hfs_tableInsert(hf, alignWithLeftMargin, alighWithRightMargin, doMirror)
                        tbl.Range.Cells.Item(1).Range.Text = "FirstPage-Even"
                End Select


        End Select
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will test the section page Layout and return a string indicating what the layout is
    ''' -   'DiffFirstPage-Not'
    ''' -   'OddAndEven'
    ''' -   'DiffFirstPage
    ''' -   'DiffFirstPage+OddAndEven
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function hf_get_HeaderFooterType(ByRef sect As Word.Section) As String
        Dim strType As String
        '
        strType = ""
        '
        If sect.PageSetup.DifferentFirstPageHeaderFooter = False And sect.PageSetup.OddAndEvenPagesHeaderFooter = False Then strType = "DiffFirstPage-Not"
        If sect.PageSetup.DifferentFirstPageHeaderFooter = False And sect.PageSetup.OddAndEvenPagesHeaderFooter = True Then strType = "OddAndEven"
        If sect.PageSetup.DifferentFirstPageHeaderFooter = True And sect.PageSetup.OddAndEvenPagesHeaderFooter = False Then strType = "DiffFirstPage"
        If sect.PageSetup.DifferentFirstPageHeaderFooter = True And sect.PageSetup.OddAndEvenPagesHeaderFooter = True Then strType = "DiffFirstPage+OddAndEven"
        '
        Return strType
    End Function
    '
    ''' <summary>
    ''' This method will return the Header/Footer table in the Header. It defaults to the "primary" (strHeaderType)
    ''' header, but the user can specify the "firstPage" (strHeaderType) header
    ''' </summary>
    ''' <returns></returns>
    Public Function hf_get_HeaderTable(ByRef sect As Word.Section, Optional strHeaderType As String = "primary") As Word.Table
        Dim tbl As Word.Table
        Dim hf As Word.HeaderFooter
        '
        tbl = Nothing
        '
        Try
            Select Case strHeaderType
                Case "firstPage"
                    If sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Exists Then
                        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                        If hf.Range.Tables.Count <> 0 Then tbl = hf.Range.Tables.Item(1)
                    End If
                    '
                Case "primary"
                    hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                    If hf.Range.Tables.Count <> 0 Then tbl = hf.Range.Tables.Item(1)

                Case "primaryOrFirstPage"
                    hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                    If hf.Range.Tables.Count <> 0 Then
                        tbl = hf.Range.Tables.Item(1)
                    Else
                        If sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Exists Then
                            hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                            If hf.Range.Tables.Count <> 0 Then tbl = hf.Range.Tables.Item(1)
                        End If
                    End If
                    '

                Case Else
                    hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
            End Select
            '
        Catch ex As Exception
            hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        End Try
        '
        '
        Return tbl
    End Function
    '
    Public Function hf_get_HeaderTableWidth(ByRef tbl As Word.Table) As Single
        'Dim drCell As Word.Cell
        Dim tblWidth As Single
        '
        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        tblWidth = tbl.PreferredWidth
        '
        Return tblWidth
        '
    End Function

    '
    ''' <summary>
    ''' This method will retrieve the Header Table from the specified Header (default is primary) and will
    ''' then set the first cell to the style strStyleName.. If there is no header table it will set the
    ''' header style to strStyleName. See chptBase_set_tagStyleInHeaderTable for shadow
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="strStyleName"></param>
    Public Sub hf_tags_setTagStyle(ByRef sect As Word.Section, strStyleName As String, Optional strHeaderType As String = "primary")
        Dim hfTbl As Word.Table
        Dim hf As Word.HeaderFooter
        Dim drCell As Word.Cell
        Dim myDoc As Word.Document
        '
        myDoc = sect.Range.Document
        hf = Nothing
        '
        Try
            hfTbl = Me.hf_get_HeaderTable(sect, strHeaderType)
            If Not IsNothing(hfTbl) Then
                drCell = hfTbl.Range.Cells.Item(1)
                drCell.Range.Style = myDoc.Styles.Item(strStyleName)
            Else
                Select Case strHeaderType
                    Case "primary"
                        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                    Case "firstPage"
                        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                    Case "even"
                        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterEvenPages)
                    Case Else
                        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                End Select
                hf.Range.Style = myDoc.Styles.Item(strStyleName)
            End If
            '
        Catch ex As Exception

        End Try
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method allows us to set the colour used for the specified strFooterType ('primary', 'firstPage', 'even', 'primaryOrFirstPage'). 
    ''' The table text colour of the footer table is set to textColour if 'colourAction' is set to 'newColour'.. If it is set to 'resetColour', 
    ''' then the text in each cell is reset to the default colour used by the style in each cell
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="textColour"></param>
    ''' <param name="colourAction"></param>
    ''' <param name="strFooterType"></param>
    Public Sub hf_set_textColourFooter(ByRef sect As Word.Section, textColour As Long, Optional colourAction As String = "newColour", Optional strFooterType As String = "primaryOrFirstPage")
        Dim hf As Word.HeaderFooter
        '
        Select Case strFooterType
            Case "primary"
                hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                Me.hf_set_textColour(hf, textColour, colourAction)
            Case "firstPage"
                hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                Me.hf_set_textColour(hf, textColour, colourAction)
            Case "even"
                hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterEvenPages)
                Me.hf_set_textColour(hf, textColour, colourAction)
            Case "primaryOrFirstPage"
                hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                Me.hf_set_textColour(hf, textColour, colourAction)
                '
                If sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Exists Then
                    hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                    Me.hf_set_textColour(hf, textColour, colourAction)
                End If
                '
            Case Else
                hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                Me.hf_set_textColour(hf, textColour, colourAction)
        End Select


    End Sub
    '
    ''' <summary>
    ''' This method allows us to set the colour used for the hf (either footer or header tables). The table text colour
    ''' is set to textColour if 'colourAction' is set to 'newColour'.. If it is set to 'resetColour', then the text in each cell
    ''' is reset to the default colour used by the style in each cell
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="textColour"></param>
    ''' <param name="colourAction"></param>
    Public Sub hf_set_textColour(ByRef hf As Word.HeaderFooter, textColour As Long, colourAction As String)
        Dim tbl As Word.Table
        Dim rng As Word.Range
        Dim drCell As Word.Cell
        Dim textStyle As Word.Style
        '
        If hf.Range.Tables.Count <> 0 Then
            tbl = hf.Range.Tables.Item(1)
            '
            Select Case colourAction
                Case "newColour"
                    'tbl.Range.Font.Color = WdColor.wdColorWhite
                    tbl.Range.Font.Color = textColour
                Case "resetColour"
                    'Reset both of the cells to the default style colour used in
                    'each cell
                    '
                    For Each drCell In tbl.Range.Cells
                        rng = drCell.Range
                        textStyle = rng.Style
                        rng.Font.Color = textStyle.Font.Color
                    Next
                    '
            End Select
        End If

    End Sub


    '
    ''' <summary>
    ''' This method will search the primary and/or first page header for a header table. If one exists it will
    ''' get the name of the style in the first cell. This is the tag style. If the header table doesn't exist
    ''' it will search for the style of the header and present that as the tag style
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="strHeaderType"></param>
    ''' <returns></returns>
    Public Function hf_tags_getTagStyleName(ByRef sect As Word.Section, Optional strHeaderType As String = "primary") As String
        Dim hfTbl As Word.Table
        Dim hf As Word.HeaderFooter
        Dim drCell As Word.Cell
        Dim myDoc As Word.Document
        Dim strStyleName As String
        Dim tagStyle As Word.Style
        '
        myDoc = sect.Range.Document
        strStyleName = ""
        hf = Nothing
        '
        Try
            hfTbl = Me.hf_get_HeaderTable(sect, strHeaderType)
            If IsNothing(hfTbl) Then
                hfTbl = Me.hf_get_HeaderTable(sect, "primaryOrFirstPage")
            End If
            '
            If Not IsNothing(hfTbl) Then
                drCell = hfTbl.Range.Cells.Item(1)
                tagStyle = drCell.Range.Style
                strStyleName = tagStyle.NameLocal
            Else
                Select Case strHeaderType
                    Case "primary"
                        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        strStyleName = Me.hf_tags_getTagStyleName(hf)
                    Case "firstPage"
                        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                        strStyleName = Me.hf_tags_getTagStyleName(hf)
                    Case "even"
                        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterEvenPages)
                        strStyleName = Me.hf_tags_getTagStyleName(hf)
                    Case "primaryOrFirstPage"
                        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        strStyleName = Me.hf_tags_getTagStyleName(hf)
                        '
                        'If we can't find a tag style in the header of the primary, then we'll
                        'search first page
                        '
                        If strStyleName = "" Then
                            hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                            strStyleName = Me.hf_tags_getTagStyleName(hf)
                        End If
                        '
                    Case Else
                        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        strStyleName = Me.hf_tags_getTagStyleName(hf)
                End Select
                '
            End If
            '
        Catch ex As Exception
            strStyleName = ""
        End Try
        '
        Return strStyleName
    End Function
    '
    Public Function hf_tags_getTagStyleName(ByRef hf As Word.HeaderFooter) As String
        Dim tagStyle As Word.Style
        Dim strStyleName As String
        '
        strStyleName = ""
        '
        If hf.Exists Then
            tagStyle = hf.Range.Style
            strStyleName = tagStyle.NameLocal
            If Not strStyleName Like "tag_*" Then strStyleName = ""
        Else
            strStyleName = ""
        End If
        '
        Return strStyleName
    End Function

    '
    ''' <summary>
    ''' This method will look in the first cell of a table and returns the style name
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function hf_tags_getTagStyleName(ByRef tbl As Word.Table) As String
        Dim drCell As Word.Cell
        Dim tagStyle As Style
        Dim strTag As String
        '
        strTag = ""
        Try
            drCell = tbl.Range.Cells(1)
            tagStyle = drCell.Range.Style
            strTag = tagStyle.NameLocal
        Catch ex As Exception
            strTag = ""

        End Try
        '
        Return strTag
    End Function
    '

    '
    ''' <summary>
    ''' This method will copy the Headers from the srcSection to the destSection.. Existing headers in
    ''' destSection will be unlinked and deleted
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="srcSection"></param>
    ''' <param name="destSection"></param>
    Public Sub hf_hfs_CopyHeader(ByRef hf As Word.HeaderFooter, ByRef srcSection As Word.Section, ByRef destSection As Word.Section)
        Dim hfDst As Word.HeaderFooter
        '
        Try
            Select Case hf.Index
                Case WdHeaderFooterIndex.wdHeaderFooterEvenPages
                    hfDst = destSection.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterEvenPages)
                    Me.hf_HF_CopySourceToDestination(hf, hfDst, True)
                Case WdHeaderFooterIndex.wdHeaderFooterFirstPage
                    hfDst = destSection.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                    Me.hf_HF_CopySourceToDestination(hf, hfDst, True)
                Case WdHeaderFooterIndex.wdHeaderFooterPrimary
                    hfDst = destSection.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                    Me.hf_HF_CopySourceToDestination(hf, hfDst, True)
            End Select

        Catch ex As Exception

        End Try
    End Sub
    '
    ''' <summary>
    ''' This method will copy the Footers from the srcSection to the destSection.. Existing footers in
    ''' destSection will be unlinked and deleted
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="srcSection"></param>
    ''' <param name="destSection"></param>
    Public Sub hf_hfs_CopyFooter(ByRef hf As Word.HeaderFooter, ByRef srcSection As Word.Section, ByRef destSection As Word.Section)
        Dim hfDst As Word.HeaderFooter
        '
        Try
            Select Case hf.Index
                Case WdHeaderFooterIndex.wdHeaderFooterEvenPages
                    hfDst = destSection.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterEvenPages)
                    Me.hf_HF_CopySourceToDestination(hf, hfDst, True)
                Case WdHeaderFooterIndex.wdHeaderFooterFirstPage
                    hfDst = destSection.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                    Me.hf_HF_CopySourceToDestination(hf, hfDst, True)
                Case WdHeaderFooterIndex.wdHeaderFooterPrimary
                    hfDst = destSection.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                    Me.hf_HF_CopySourceToDestination(hf, hfDst, True)
            End Select

        Catch ex As Exception

        End Try
    End Sub
    '
    '
    Public Sub xhf_hfs_CopyHeaderFooter(strHeaderOrFooter As String, ByRef srcSection As Section, ByRef destSection As Section)
        'This method will copy the header or footer from the source section
        'to the destination section.. It is currently limited to sections
        'with "same first page"
        '
        Dim hfSrc As HeaderFooter
        Dim hfDst As HeaderFooter
        Dim strHFType As String
        '
        Me.hf_hfs_linkUnlinkAll(destSection, False)
        hfDst = Nothing
        hfSrc = Nothing
        strHFType = ""
        '
        'Define the source and destination Header/Footers
        Select Case strHeaderOrFooter
            Case "header"
                Me.hf_headers_delete(destSection)
                'Call Me.hf_headers_Delete_Contents_All(destSection)
                '
                strHFType = Me.hf_get_HeaderFooterType(srcSection)
                '
                Select Case strHFType
                    Case "DiffFirstPage-Not"
                        hfSrc = srcSection.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        hfDst = destSection.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        '
                        Me.hf_HF_CopySourceToDestination(hfSrc, hfDst, True)
                        '
                    Case "DiffFirstPage"
                        hfSrc = srcSection.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                        hfDst = destSection.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                        '
                        Me.hf_HF_CopySourceToDestination(hfSrc, hfDst, True)
                        '
                        hfSrc = srcSection.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        hfDst = destSection.Headers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        '
                        Me.hf_HF_CopySourceToDestination(hfSrc, hfDst, True)
                        '
                    Case "OddAndEven"
                    Case "DiffFirstPage+OddAndEven"

                End Select

            Case "footer"
                Me.hf_footers_delete(destSection)
                'Call Me.hf_footers_Delete_Contents_All(destSection)
                '
                strHFType = Me.hf_get_HeaderFooterType(srcSection)
                '
                Select Case strHFType
                    Case "DiffFirstPage-Not"
                        hfSrc = srcSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        hfDst = destSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        '
                        Me.hf_HF_CopySourceToDestination(hfSrc, hfDst, True)
                        '
                    Case "DiffFirstPage"
                        hfSrc = srcSection.Footers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                        hfDst = destSection.Footers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                        '
                        Me.hf_HF_CopySourceToDestination(hfSrc, hfDst, True)
                        '
                        hfSrc = srcSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        hfDst = destSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                        '
                        Me.hf_HF_CopySourceToDestination(hfSrc, hfDst, True)
                        '
                    Case "OddAndEven"
                    Case "DiffFirstPage+OddAndEven"
                        '

                End Select





                hfSrc = srcSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                hfDst = destSection.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        End Select
        '



    End Sub
    '

    '
#Region "table Build"
    '
    Public Function hf_header_buildBasicTable(ByRef hf As Word.HeaderFooter, tblWidth As Single, doMirror As Boolean) As Word.Table
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim drCol As Word.Column
        Dim dr As Word.Row
        Dim para As Word.Paragraph
        Dim sect As Word.Section
        '
        sect = hf.Range.Sections.Item(1)
        rng = hf.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        '*****
        tbl = rng.Tables.Add(rng, 1, 2)

        '
        tbl.Style = hf.Range.Document.Styles.Item("aac Table (no lines)")
        tbl.ApplyStyleHeadingRows = True
        'objTblMgr.tlb_convert_oneRowTableToWCAG(cellPadding, leftIndentBody, tbl)
        '
        tbl.AllowAutoFit = False
        tbl.Borders.Enable = False
        '
        tbl.TopPadding = 0.0
        tbl.BottomPadding = 0.0
        tbl.LeftPadding = 0.0
        tbl.RightPadding = 0.0
        '
        If doMirror Then
            drCol = tbl.Columns.Item(2)
            'drCol.Width = 2.0 * (tblWidth / 3.0)
            drCol.Width = tblWidth / 2.0
            tbl.Columns.Item(1).Width = tblWidth - drCol.Width
        Else
            drCol = tbl.Columns.Item(1)
            'drCol.Width = 2.0 * (tblWidth / 3.0)
            drCol.Width = tblWidth / 2.0
            tbl.Columns.Item(2).Width = tblWidth - drCol.Width
            tbl.Rows.Item(1).LeftIndent = -(sect.PageSetup.LeftMargin - objGlobals.glb_math_MillimetersToPoints(objGlobals._glb_header_leftEdge))
            '
            '****
            'Force the style in the 2nd cell of the header table to be right aligned
            'so that the secturity watermark sits at the edge of the page. COuld place a right
            'aligned style here, but this would make this Template inconsistent with legacy T&G versions
            '
            para = tbl.Range.Cells.Item(2).Range.Paragraphs.Item(1)
            para.Alignment = WdParagraphAlignment.wdAlignParagraphRight
            '
        End If
        '
        dr = tbl.Rows.Item(1)
        dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
        dr.Height = 17.8
        '
        Return tbl
        '
    End Function
    '
    '
    Public Function hf_footer_buildBasicTable(ByRef hf As Word.HeaderFooter, tblWidth As Single, doMirror As Boolean, doPageNumber As Boolean, doStyleRefs As Boolean) As Word.Table
        Dim myDoc As Word.Document
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim drCol As Word.Column
        Dim dr As Word.Row
        Dim sect As Word.Section
        '
        sect = hf.Range.Sections.Item(1)
        myDoc = hf.Range.Document
        rng = hf.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        tbl = rng.Tables.Add(rng, 1, 2)
        '
        tbl.Style = hf.Range.Document.Styles.Item("aac Table (no lines)")
        tbl.ApplyStyleHeadingRows = True
        'objTblMgr.tlb_convert_oneRowTableToWCAG(cellPadding, leftIndentBody, tbl)
        '
        tbl.AllowAutoFit = False
        tbl.Borders.Enable = False
        '
        tbl.TopPadding = 0.0
        tbl.BottomPadding = 0.0
        tbl.LeftPadding = 0.0
        tbl.RightPadding = 0.0
        '
        dr = tbl.Rows.Item(1)
        dr.HeightRule = WdRowHeightRule.wdRowHeightAtLeast
        '
        dr.Height = Me.objGlobals.glb_hfs_getFooterTable_Height_Nominal()
        'dr.Height = 31.0
        'dr.Height = 41.0

        '
        '*** 20231209 changes
        'tbl.Range.Cells.Item(1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
        'tbl.Range.Cells.Item(2).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
        '
        tbl.Range.Cells.Item(1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter
        tbl.Range.Cells.Item(2).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter
        '
        '
        If doMirror Then
            drCol = tbl.Columns.Item(1)
            'drCol.Width = 37.2
            'drCol.Width = objGlobals._glb_footer_PageNumColWidth
            drCol.Width = sect.PageSetup.RightMargin - objGlobals.glb_math_MillimetersToPoints(objGlobals._glb_footer_rightEdge)

            '
            '*** Change made in 20231209 to accomodate AA new 2024 template
            '
            tbl.Columns.Item(2).Width = tblWidth - drCol.Width
            'tbl.Columns.Item(2).Width = objGlobals.glb_get_widthBetweenMargins(sect)

            'tbl.Columns.Item(2).Width = tblWidth
            'tbl.Range.Cells.Item(1).Width = 37.2
            'tbl.Range.Cells.Item(2).Width = tblWidth
            'dr = tbl.Rows.Item(1)
            tbl.Rows.LeftIndent = -tbl.Columns.Item(1).Width

            'dr.Shading.ForegroundPatternColor = WdColor.wdColorAqua
            'dr.LeftIndent = -37.2
            'dr.LeftIndent = dr.LeftIndent - 37.2
            '
            '*** 
            '
            Try
                tbl.Range.Cells.Item(2).Range.Style = myDoc.Styles.Item("Footer Text")
                tbl.Range.Cells.Item(1).Range.Style = myDoc.Styles.Item("pageNumber")
            Catch ex As Exception

            End Try
            '
            'tbl.Range.Cells.Item(1).LeftPadding = 5.4
            tbl.Range.Cells.Item(1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
            tbl.Range.Cells.Item(2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
            '
        Else
            drCol = tbl.Columns.Item(2)
            'drCol.Width = 37.2
            'drCol.Width = objGlobals._glb_footer_PageNumColWidth
            drCol.Width = sect.PageSetup.RightMargin - objGlobals.glb_math_MillimetersToPoints(objGlobals._glb_footer_rightEdge)

            '
            '*** Change made in 20231209 to accomodate AA new 2024 template
            '
            tbl.Columns.Item(1).Width = tblWidth - drCol.Width
            '
            '***** The following will bring the page number column back between the margins
            'tbl.Columns.Item(1).Width = tblWidth - drCol.Width - drCol.Width
            '*****
            '
            tbl.Rows.LeftIndent = -(sect.PageSetup.LeftMargin - objGlobals.glb_math_MillimetersToPoints(objGlobals._glb_footer_leftEdge))
            'tbl.Columns.Item(1).Width = objGlobals.glb_get_widthBetweenMargins(sect)
            'tbl.Columns.Item(1).Width = tblWidth
            '
            Try
                tbl.Range.Cells.Item(1).Range.Style = myDoc.Styles.Item("Footer Text")
                tbl.Range.Cells.Item(2).Range.Style = myDoc.Styles.Item("pageNumber")
            Catch ex As Exception

            End Try
            '
            'tbl.Range.Cells.Item(2).RightPadding = 5.4
            tbl.Range.Cells.Item(1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
            tbl.Range.Cells.Item(2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
            '
        End If
        '
        If doPageNumber Then Me.hf_insert_PageField(tbl, doMirror)
        If doStyleRefs Then Me.hf_insert_coverPageStyleRefs(tbl, doMirror)
        '
        Return tbl
        '
    End Function
    '
    ''' <summary>
    ''' This method will insert the 'EVen Pages' header so that the header table's left and right edge are
    ''' as specified (a negative value for each indicates the header is flush with the left or right
    ''' margin).. It will also (optionally) insert a logo and it returns the header table
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="leftEdge"></param>
    ''' <param name="rightEdge"></param>
    ''' <param name="doLogo"></param>
    ''' <returns></returns>
    Private Function hf_insertEvenPages_header(ByRef hf As Word.HeaderFooter, leftEdge As Single, rightEdge As Single, Optional doLogo As Boolean = True, Optional reColourLogo As Boolean = False) As Word.Table
        Dim sect As Word.Section
        Dim myDoc As Word.Document
        Dim tblWidth As Single
        Dim tbl As Word.Table
        Dim shp As Word.Shape
        Dim doMirror As Boolean
        '
        sect = hf.Range.Sections.Item(1)
        myDoc = sect.Range.Document
        '
        tblWidth = sect.PageSetup.PageWidth - leftEdge - rightEdge
        '
        'DoMirror is true for Even Pages
        doMirror = sect.PageSetup.MirrorMargins
        tbl = Me.hf_header_buildBasicTable(hf, tblWidth, doMirror)
        '
        'Indents are now done in BuildBasicTable
        If sect.PageSetup.MirrorMargins Then
            'tbl.Rows.Item(1).LeftIndent = -(sect.PageSetup.RightMargin - leftEdge)
        Else
            'tbl.Rows.Item(1).LeftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
        End If
        '
        '*** Warning: If Even Page headers are inserted in a section that is not showing the even
        '*** page (at the time of insertion).. Then the logo shp.Left position stays at 0.0. That is
        '*** any setting for shp.Left in placeLogo does NOT stick.... But the same procedures work
        '*** if the even page is showing/visible when insertion of the even header takes place
        '
        If doLogo Then
            shp = Me.hf_logo_placeLogo(tbl, doMirror, reColourLogo,, "even")
        End If
        '
        Return tbl
        '
    End Function
    '
    ''' <summary>
    ''' This method will insert the 'EVen Pages' footerer so that the header table's left and right edge are
    ''' as specified (a negative value for each indicates the footer is flush with the left or right
    ''' margin).. It will also (optionally) insert a logo and it returns the footer table
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="leftEdge"></param>
    ''' <param name="rightEdge"></param>
    ''' <param name="doPageNumber"></param>
    ''' <param name="doStyleRefs"></param>
    ''' <returns></returns>
    Private Function hf_insertEvenPages_Footer(ByRef hf As Word.HeaderFooter, leftEdge As Single, rightEdge As Single, doPageNumber As Boolean, doStyleRefs As Boolean) As Word.Table
        Dim sect As Word.Section
        Dim myDoc As Word.Document
        Dim tblWidth As Single
        Dim tbl As Word.Table
        Dim doMirror As Boolean
        '
        sect = hf.Range.Sections.Item(1)
        myDoc = sect.Range.Document
        '
        tblWidth = sect.PageSetup.PageWidth - leftEdge - rightEdge
        '
        'DoMirror is true for Even Pages
        doMirror = sect.PageSetup.MirrorMargins
        tbl = Me.hf_footer_buildBasicTable(hf, tblWidth, doMirror, doPageNumber, doStyleRefs)
        '
        If sect.PageSetup.MirrorMargins Then
            'tbl.Rows.Item(1).LeftIndent = -(sect.PageSetup.RightMargin - leftEdge)
            'tbl.Rows.Item(1).LeftIndent = -37.2
        Else
            'tbl.Rows.Item(1).LeftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
        End If
        '
        'If doLogo Then shp = Me.hf_logo_placeLogo(tbl, doMirror)
        '
        Return tbl
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will insert the 'Primary' header so that the header table's left and right edge are
    ''' as specified (a negative value for each indicates the header is flush with the left or right
    ''' margin).. It will also (optionally) insert a logo and it returns the header table
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="leftEdge"></param>
    ''' <param name="rightEdge"></param>
    ''' <param name="doLogo"></param>
    ''' <returns></returns>
    Public Function hf_insertPrimary_header(ByRef hf As Word.HeaderFooter, leftEdge As Single, rightEdge As Single, Optional doLogo As Boolean = True, Optional reColourLogo As Boolean = False) As Word.Table
        Dim sect As Word.Section
        Dim myDoc As Word.Document
        Dim tblWidth As Single
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim doMirror As Boolean
        '
        If hf.LinkToPrevious Then hf.LinkToPrevious = False
        '
        sect = hf.Range.Sections.Item(1)
        myDoc = sect.Range.Document
        '
        tblWidth = sect.PageSetup.PageWidth - leftEdge - rightEdge
        '
        'DoMirror is true for Even Pages
        doMirror = sect.PageSetup.MirrorMargins
        tbl = Me.hf_header_buildBasicTable(hf, tblWidth, doMirror)
        '
        '
        tbl.Rows.Item(1).LeftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
        'dr.LeftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
        '
        'hf_objGlobals.glb_screen_update(False)

        '
        Try
            If doLogo Then
                rng = tbl.Range.Cells.Item(1).Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                'Me.hf_logo_placeLogo(rng, reColourLogo)
                Me.hf_logo_placeLogo(tbl, doMirror, reColourLogo, True, "prim")
            End If
        Catch ex As Exception

        End Try
        '
        Return tbl
        '
    End Function
    '
    ''' <summary>
    ''' This method will insert the 'Primary' footer so that the footer table's left and right edge are
    ''' as specified (a negative value for each indicates the footer is flush with the left or right
    ''' margin).. It will also (optionally) insert a logo and it returns the footer table
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="leftEdge"></param>
    ''' <param name="rightEdge"></param>
    ''' <param name="doPageNumber"></param>
    ''' <param name="doStyleRefs"></param>
    ''' <returns></returns>
    Public Function hf_insertPrimary_Footer(ByRef hf As Word.HeaderFooter, leftEdge As Single, rightEdge As Single, doPageNumber As Boolean, dostyleRefs As Boolean) As Word.Table
        Dim sect As Word.Section
        Dim myDoc As Word.Document
        Dim tblWidth As Single
        Dim tbl As Word.Table
        Dim doMirror As Boolean
        '
        sect = hf.Range.Sections.Item(1)
        myDoc = sect.Range.Document
        '
        tblWidth = sect.PageSetup.PageWidth - leftEdge - rightEdge
        'tblWidth = hf_objGlobals.glb_get_widthBetweenMargins(sect)

        '
        'DoMirror is true for Even Pages
        doMirror = False
        tbl = Me.hf_footer_buildBasicTable(hf, tblWidth, doMirror, doPageNumber, dostyleRefs)
        '
        tbl.Rows.Item(1).LeftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
        'dr.LeftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
        '
        Return tbl
        '
    End Function
    '
    ''' <summary>
    ''' This method will insert the 'FirstPage' header so that the header table's left and right edge are
    ''' as specified (a negative value for each indicates the header is flush with the left or right
    ''' margin).. It will also (optionally) insert a logo and it returns the header table
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="leftEdge"></param>
    ''' <param name="rightEdge"></param>
    ''' <param name="doLogo"></param>
    ''' <returns></returns>
    Public Function hf_insertFirstPage_Header(ByRef hf As Word.HeaderFooter, leftEdge As Single, rightEdge As Single, Optional doLogo As Boolean = True, Optional reColourLogo As Boolean = False) As Word.Table
        Dim sect As Word.Section
        Dim myDoc As Word.Document
        Dim tblWidth As Single
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim doMirror As Boolean
        '
        sect = hf.Range.Sections.Item(1)
        myDoc = sect.Range.Document
        '
        tblWidth = sect.PageSetup.PageWidth - leftEdge - rightEdge
        '
        'DoMirror is true for Even Pages
        doMirror = False
        tbl = Me.hf_header_buildBasicTable(hf, tblWidth, doMirror)
        '
        tbl.Rows.Item(1).LeftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
        'dr.LeftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
        '
        If doLogo Then
            rng = tbl.Range.Cells.Item(1).Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            Me.hf_logo_placeLogo(rng, reColourLogo)
        End If
        '
        Return tbl
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will insert the 'First' footer so that the footer table's left and right edge are
    ''' as specified (a negative value for each indicates the footer is flush with the left or right
    ''' margin).. It will also (optionally) insert a logo and it returns the footer table
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="leftEdge"></param>
    ''' <param name="rightEdge"></param>
    ''' <param name="doPageNumber"></param>
    ''' <param name="doStyleRefs"></param>
    ''' <returns></returns>
    Public Function hf_insertFirstPage_Footer(ByRef hf As Word.HeaderFooter, leftEdge As Single, rightEdge As Single, doPageNumber As Boolean, doStyleRefs As Boolean) As Word.Table
        Dim sect As Word.Section
        Dim myDoc As Word.Document
        Dim tblWidth As Single
        Dim tbl As Word.Table
        Dim doMirror As Boolean
        '
        sect = hf.Range.Sections.Item(1)
        myDoc = sect.Range.Document
        '
        tblWidth = sect.PageSetup.PageWidth - leftEdge - rightEdge
        'Table width total... Includes the outdented page number column???
        'tblWidth = sect.PageSetup.PageWidth - leftEdge - rightEdge + objGlobals._glb_footer_PageNumOffset
        '
        'DoMirror is true for Even Pages
        doMirror = False
        tbl = Me.hf_footer_buildBasicTable(hf, tblWidth, doMirror, doPageNumber, doStyleRefs)
        '
        tbl.Rows.Item(1).LeftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
        'dr.LeftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
        '
        Return tbl
        '
    End Function
    ''' <summary>
    ''' Insert header table with the leftedge of the table "leftEdge" mm from the left side of the
    ''' page and the rightedge of the table "rightEdge" mm from the right side of the page
    ''' If leftEdge = -1.0 then the left edge is flush with the left margin
    ''' If rightEdge = -1.0 then the right edge is flush with the right margin
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="leftEdge"></param>
    ''' <param name="rightEdge"></param>
    ''' <param name="doMirror"></param>
    ''' <returns></returns>
    Private Function hf_hfs_tableInsert(ByRef hf As Word.HeaderFooter, leftEdge As Single, rightEdge As Single, doMirror As Boolean, Optional doLogo As Boolean = True, Optional reColourLogo As Boolean = False) As Word.Table
        Dim objTools As New cTools()
        Dim objBBMgr As New cBBlocksHandler()
        Dim myDoc As Word.Document
        Dim tbl As Word.Table
        Dim dr As Word.Row
        Dim drCol As Word.Column
        Dim rng As Word.Range
        Dim tblWidth, oldLeftEdge As Single
        Dim alignWithLeftMargin, alignWithRightMargin As Boolean
        Dim sect As Word.Section
        '
        sect = hf.Range.Sections.Item(1)
        myDoc = sect.Range.Document
        '
        'leftEdge = objTools.MillimetersToPoints(leftEdge)
        'rightEdge = objTools.MillimetersToPoints(rightEdge)

        If leftEdge < 0.0 Then
            alignWithLeftMargin = True
            leftEdge = sect.PageSetup.LeftMargin
        End If
        If rightEdge < 0.0 Then
            alignWithRightMargin = True
            rightEdge = sect.PageSetup.RightMargin
        End If
        '
        tblWidth = sect.PageSetup.PageWidth - leftEdge - rightEdge
        '

        'hf.Range.PageSetup.OddAndEvenPagesHeaderFooter
        '

        'Swap for even pages
        If doMirror Then
            oldLeftEdge = leftEdge
            leftEdge = rightEdge
            rightEdge = oldLeftEdge
        End If
        '
        rng = hf.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        tbl = rng.Tables.Add(rng, 1, 2)
        tbl.AllowAutoFit = False
        tbl.Borders.Enable = False
        '
        tbl.TopPadding = 0.0
        tbl.BottomPadding = 0.0
        tbl.LeftPadding = 0.0
        tbl.RightPadding = 0.0

        If hf.IsHeader Then
            '
            drCol = tbl.Columns.Item(1)
            drCol.Width = 2.0 * (tblWidth / 3.0)
            tbl.Columns.Item(2).Width = tblWidth - drCol.Width
            '
            dr = tbl.Rows.Item(1)
            dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
            dr.Height = 17.8
        Else
            'Then must be a footer
            dr = tbl.Rows.Item(1)
            dr.HeightRule = WdRowHeightRule.wdRowHeightAtLeast
            dr.Height = 31.0
            '
            tbl.Range.Cells.Item(1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
            tbl.Range.Cells.Item(2).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
            '
            If doMirror Then
                drCol = tbl.Columns.Item(1)
                drCol.Width = 37.2
                tbl.Columns.Item(2).Width = tblWidth - drCol.Width
                Try
                    tbl.Range.Cells.Item(2).Range.Style = myDoc.Styles.Item("Footer Text")
                    tbl.Range.Cells.Item(1).Range.Style = myDoc.Styles.Item("pageNumber")
                Catch ex As Exception

                End Try
                '
                tbl.Range.Cells.Item(1).LeftPadding = 10.0
                tbl.Range.Cells.Item(1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
                tbl.Range.Cells.Item(2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
                '
            Else
                drCol = tbl.Columns.Item(2)
                drCol.Width = 37.2
                tbl.Columns.Item(1).Width = tblWidth - drCol.Width
                '
                Try
                    tbl.Range.Cells.Item(1).Range.Style = myDoc.Styles.Item("Footer Text")
                    tbl.Range.Cells.Item(2).Range.Style = myDoc.Styles.Item("pageNumber")
                Catch ex As Exception

                End Try
                '
                tbl.Range.Cells.Item(2).RightPadding = 10.0
                tbl.Range.Cells.Item(1).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
                tbl.Range.Cells.Item(2).Range.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
                '
            End If
            '
        End If
        '
        If doMirror And sect.PageSetup.OddAndEvenPagesHeaderFooter Then
            dr.LeftIndent = -(sect.PageSetup.RightMargin - leftEdge)
        Else
            dr.LeftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
        End If


        If doMirror And sect.PageSetup.MirrorMargins Then
            'dr.LeftIndent = -(sect.PageSetup.RightMargin - leftEdge)
        Else
            'dr.LeftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
        End If
        '
        If doLogo Then
            'rng2 = tbl.Range.Cells.Item(1).Range
            'If doMirror Then
            'rng2 = tbl.Range.Cells.Item(2).Range
            'tbl.Columns.Item(2).Width
            'End If
            'rng2.Collapse(WdCollapseDirection.wdCollapseStart)
            Me.hf_logo_placeLogo(tbl, doMirror, reColourLogo, True, "prim")
        End If

        '
        Return tbl


    End Function
    '
    ''' <summary>
    ''' This method will place the logo at rng
    ''' </summary>
    ''' <param name="rng"></param>
    Public Function hf_logo_placeLogo(ByRef rng As Word.Range, Optional reColourLogo As Boolean = False, Optional setAsDecorative As Boolean = True) As Word.Shape
        Dim shp As Word.Shape
        Dim objWCAGMgr As New cWCAGMgr()
        Dim objBBMgr As New cBBlocksHandler()
        Dim objBrndMgr As New cBrandMgr()
        '
        shp = Nothing
        '
        rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_HFs_logoHeader", "headerFooters", rng)
        If Not IsNothing(rng) Then
            If rng.ShapeRange.Count <> 0 Then
                shp = rng.ShapeRange.Item(1)
                'shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
                'shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
                shp.LockAnchor = True
                shp.LockAspectRatio = True
                shp.Height = 11.25
                shp.Top = 5.5
                'shp.Top = 2.7
                'shp.Left = 0.25
                shp.Left = 0.0
                '
                If setAsDecorative Then objWCAGMgr.wcag_set_decorative(shp, True)
                'shp.Width = 113.1
                If reColourLogo Then objBrndMgr.brnd_recolour_Logo(shp)

            End If
        End If
        '
        Return shp
        '
    End Function

    Public Function hf_logo_placeLogo(ByRef tbl As Word.Table, doMirror As Boolean, Optional reColourLogo As Boolean = False, Optional setAsDecorative As Boolean = True, Optional strEvenOdd As String = "") As Word.Shape
        Dim shp As Word.Shape
        Dim drColWidth As Single
        Dim objBBMgr As New cBBlocksHandler()
        Dim objBrndMgr As New cBrandMgr()
        Dim rng As Word.Range
        '
        'Set default condition for odd/even fold
        'rng = tbl.Range.Cells.Item(1).Range
        '
        Select Case strEvenOdd
            Case "even"
                rng = tbl.Range.Cells.Item(1).Range
                drColWidth = tbl.Range.Columns.Item(1).Width
            Case "prim"
                If doMirror Then
                    rng = tbl.Range.Cells.Item(2).Range
                    drColWidth = tbl.Range.Columns.Item(2).Width
                Else
                    rng = tbl.Range.Cells.Item(1).Range
                    drColWidth = tbl.Range.Columns.Item(1).Width
                End If

            Case Else
                If doMirror Then
                    rng = tbl.Range.Cells.Item(2).Range
                    drColWidth = tbl.Range.Columns.Item(2).Width
                Else
                    rng = tbl.Range.Cells.Item(1).Range
                    drColWidth = tbl.Range.Columns.Item(1).Width
                End If

        End Select

        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        shp = Me.hf_logo_placeLogo(rng, reColourLogo, setAsDecorative)
        If strEvenOdd = "prim" And doMirror Then
            shp.Left = drColWidth - shp.Width
        End If
        '
        Try
            'rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_HFs_logoHeader", "headerFooters", rng)
            'If rng.ShapeRange.Count <> 0 Then
            'shp = rng.ShapeRange.Item(1)
            'shp.LockAnchor = True
            'shp.ZOrder(MsoZOrderCmd.msoBringInFrontOfText)
            '
            'shp.LockAspectRatio = True
            'shp.Height = 11.25
            'shp.Top = 5.5
            'shp.Left = 5.5

            '
            'shpLeft = 0.0
            'If doMirror Then
            ' shpLeft = drColWidth - shp.Width
            'End If
            '
            'If the tbl is not visible when this command is executed, the shp.Left value just doesn't 
            'stick.. So make sure the even header is visible when this is built/applied
            '
            'shp.Left = shpLeft
            'shp.Left = shpLeft
            'End If
            '
            'If reColourLogo Then objBrndMgr.brnd_recolour_Logo(shp)

        Catch ex As Exception

        End Try
        '
        Return shp
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will search the header table(s) of sect (both primary and first page) and
    ''' if it finds a shape in the first cell and if it is the logo it will recolour the name
    ''' as per rgbName and the underscore bar as per rgbBar... If rgbBar is negative it will
    ''' set the colour to the standard purple
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="rgbName"></param>
    ''' <param name="rgbBar"></param>
    Public Sub hf_logo_colourLogo(ByRef sect As Word.Section, rgbName As Long, rgbBar As Long)
        Dim hf As Word.HeaderFooter
        Dim tbl As Word.Table
        Dim rng As Word.Range
        Dim shp As Word.Shape
        Dim grpShp As Word.Shape
        Dim i As Integer
        '
        If rgbBar < 0 Then rgbBar = RGB(108, 63, 153)
        '
        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        If hf.Range.Tables.Count <> 0 Then
            tbl = hf.Range.Tables.Item(1)
            rng = tbl.Range.Cells.Item(1).Range
            If rng.ShapeRange.Count <> 0 Then
                shp = rng.ShapeRange.Item(1)
                If shp.Name = "logo_AAC_TandG" Then
                    For i = 1 To shp.GroupItems.Count
                        grpShp = shp.GroupItems.Item(i)
                        If grpShp.Name = "Freeform 6" Then grpShp.Fill.ForeColor.RGB = rgbName
                        If grpShp.Name = "Rectangle 42" Then grpShp.Fill.ForeColor.RGB = rgbBar

                    Next
                End If
            End If
        End If
        '
        Try
            hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
            If hf.Range.Tables.Count <> 0 Then
                tbl = hf.Range.Tables.Item(1)
                rng = tbl.Range.Cells.Item(1).Range
                If rng.ShapeRange.Count <> 0 Then
                    shp = rng.ShapeRange.Item(1)
                    For i = 1 To shp.GroupItems.Count
                        grpShp = shp.GroupItems.Item(i)
                        If grpShp.Name = "Freeform 6" Then grpShp.Fill.ForeColor.RGB = rgbName
                        If grpShp.Name = "Rectangle 42" Then grpShp.Fill.ForeColor.RGB = rgbBar

                    Next
                End If
            End If
        Catch ex As Exception

        End Try
    End Sub

    ''' <summary>
    ''' This method will insert a standard header table in the specified header
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <returns></returns>
    Private Function hf_footers_tableInsert(ByRef hf As Word.HeaderFooter, alignWithLeftMargin As Boolean, alignWithRightMargin As Boolean, doMirror As Boolean) As Word.Table
        Dim objTools As New cTools()
        Dim tbl As Word.Table
        Dim dr As Word.Row
        Dim drCol As Word.Column
        Dim rng As Word.Range
        Dim defaultLeftEdge, defaultRightEdge As Single
        Dim leftEdge, rightEdge, tblWidth, oldLeftEdge As Single
        Dim sect As Word.Section
        '
        '*** Test
        'hf.Range.PageSetup.
        '
        sect = hf.Range.Sections.Item(1)
        '
        defaultLeftEdge = objTools.tools_math_MillimetersToPoints(20.0)
        defaultRightEdge = objTools.tools_math_MillimetersToPoints(10.0)
        '
        If alignWithLeftMargin Then
            defaultLeftEdge = sect.PageSetup.LeftMargin
        End If
        If alignWithRightMargin Then
            defaultRightEdge = sect.PageSetup.RightMargin
        End If
        '
        leftEdge = defaultLeftEdge
        rightEdge = defaultRightEdge
        '
        tblWidth = sect.PageSetup.PageWidth - leftEdge - rightEdge
        '
        If doMirror Then
            oldLeftEdge = leftEdge
            leftEdge = rightEdge
            rightEdge = oldLeftEdge
        End If
        '
        '
        'pgWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
        '
        'Type xx = hf.GetType()
        '
        rng = hf.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        tbl = rng.Tables.Add(rng, 1, 2)
        tbl.AllowAutoFit = False
        tbl.Borders.Enable = False
        '
        tbl.TopPadding = 0.0
        tbl.BottomPadding = 0.0
        tbl.LeftPadding = 0.0
        tbl.RightPadding = 0.0
        '
        drCol = tbl.Columns.Item(1)
        drCol.Width = tblWidth / 2.0
        tbl.Columns.Item(2).Width = tblWidth - drCol.Width
        '
        dr = tbl.Rows.Item(1)
        dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
        dr.Height = 17.8
        '

        If doMirror And sect.PageSetup.MirrorMargins Then
            dr.LeftIndent = -(sect.PageSetup.RightMargin - leftEdge)
        Else
            dr.LeftIndent = -(sect.PageSetup.LeftMargin - leftEdge)
        End If
        '
        Return tbl

    End Function

#End Region
    '
    '
    ''' <summary>
    ''' This method will take the footer table as input and will delete the
    ''' existing fields in page number Cell (item 2) and re-insert the page field
    ''' </summary>
    ''' <param name="footerTbl"></param>
    Public Function hf_insert_PageField(ByRef footerTbl As Word.Table, doMirror As Boolean) As Word.Field
        Dim rng As Word.Range
        Dim fld As Word.Field
        Dim drCell As Word.Cell
        Dim myDoc As Word.Document
        '
        drCell = footerTbl.Range.Cells.Item(2)
        If doMirror Then drCell = footerTbl.Range.Cells.Item(1)
        '
        rng = drCell.Range
        myDoc = rng.Document
        '
        For Each fld In rng.Fields
            fld.Delete()
        Next
        'rng.Delete()
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Style = myDoc.Styles.Item("pageNumber")
        fld = rng.Fields.Add(rng, WdFieldType.wdFieldPage)
        '
        '
        ' drCell.Range.Style = myDoc.Styles.Item("pageNumber")
        '
        Return fld
        '
    End Function
    '
    ''' <summary>
    ''' We do this by deleting all fields in the footer, then add back the CoverPage Field
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub hf_footer_deleteEmptySubTitleField(ByRef sect As Word.Section)
        Dim hf As Word.HeaderFooter
        Dim ftrTbl As Word.Table
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        Dim fld As Word.Field
        '
        For Each hf In sect.Footers
            If hf.Exists Then
                Try
                    Select Case hf.Index
                        Case WdHeaderFooterIndex.wdHeaderFooterEvenPages
                            If hf.Range.Tables.Count > 0 Then
                                ftrTbl = hf.Range.Tables.Item(1)
                                drCell = ftrTbl.Range.Cells.Item(2)
                                rng = drCell.Range
                                For Each fld In drCell.Range.Fields
                                    'If fld.
                                    fld.Delete()
                                Next
                                drCell.Range.Text = ""
                                rng = drCell.Range
                                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                                fld = rng.Fields.Add(rng, WdFieldType.wdFieldStyleRef, """Cp Title""", True)
                            End If

                        Case WdHeaderFooterIndex.wdHeaderFooterFirstPage, WdHeaderFooterIndex.wdHeaderFooterPrimary
                            If hf.Range.Tables.Count > 0 Then
                                ftrTbl = hf.Range.Tables.Item(1)
                                drCell = ftrTbl.Range.Cells.Item(1)
                                For Each fld In drCell.Range.Fields
                                    fld.Delete()
                                Next
                                drCell.Range.Text = ""
                                rng = drCell.Range
                                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                                fld = rng.Fields.Add(rng, WdFieldType.wdFieldStyleRef, """Cp Title""", True)

                            End If
                    End Select

                Catch ex As Exception

                End Try
            End If

        Next


    End Sub


    '
    Public Sub hf_footer_addFooterFields(ByRef sect As Word.Section)
        Dim hf As Word.HeaderFooter
        Dim ftrTbl As Word.Table
        Dim drCell As Word.Cell
        Dim fld As Word.Field
        '
        For Each hf In sect.Footers
            If hf.Exists Then
                Try
                    Select Case hf.Index
                        Case WdHeaderFooterIndex.wdHeaderFooterEvenPages
                            If hf.Range.Tables.Count > 0 Then
                                ftrTbl = hf.Range.Tables.Item(1)
                                drCell = ftrTbl.Range.Cells.Item(2)
                                For Each fld In drCell.Range.Fields
                                    fld.Delete()
                                Next
                                drCell.Range.Text = ""
                                Me.hf_insert_coverPageStyleRefs(ftrTbl, True)
                            End If

                        Case WdHeaderFooterIndex.wdHeaderFooterFirstPage, WdHeaderFooterIndex.wdHeaderFooterPrimary
                            If hf.Range.Tables.Count > 0 Then
                                ftrTbl = hf.Range.Tables.Item(1)
                                drCell = ftrTbl.Range.Cells.Item(1)
                                For Each fld In drCell.Range.Fields
                                    fld.Delete()
                                Next
                                drCell.Range.Text = ""
                                Me.hf_insert_coverPageStyleRefs(ftrTbl, False)

                            End If
                    End Select

                Catch ex As Exception

                End Try
            End If

        Next


    End Sub
    '
    Public Function hf_insert_coverPageStyleRefs(ByRef footerTbl As Word.Table, doMirror As Boolean) As Word.Field
        Dim rng As Word.Range
        Dim fld As Word.Field
        Dim drCell As Word.Cell
        Dim myDoc As Word.Document
        Dim stylCp, stylCpSub As Word.Style
        '
        fld = Nothing
        myDoc = footerTbl.Range.Document
        '
        Try
            If doMirror Then
                drCell = footerTbl.Range.Cells.Item(2)
                rng = drCell.Range
                rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft

            Else
                drCell = footerTbl.Range.Cells.Item(1)
                rng = drCell.Range
                rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
                '
            End If
            '
            stylCpSub = myDoc.Styles.Item("Cp SubTitle")
            stylCp = myDoc.Styles.Item("Cp Title")

            '
            'Cater for documents without these styles.. Users of the Addin may encounter this problem
            '
            If Not (IsNothing(stylCp) Or IsNothing(stylCpSub)) Then
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Text = vbCrLf
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                rng.Fields.Add(rng, WdFieldType.wdFieldStyleRef, """Cp SubTitle""", True)
                rng = drCell.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                fld = rng.Fields.Add(rng, WdFieldType.wdFieldStyleRef, """Cp Title""", True)

            End If
            '
            '
        Catch ex As Exception

        End Try

        Return fld
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will insert a coloured rectangle into the Shapes collection of the HeaderFooter hf
    ''' and it will lock it to the range rng. Typically the range will be a collapsed version of the
    ''' the hf range.. If you specifiy an RGB colour, then the rectange fill will be set to that colour.
    ''' If you leave that parameter empty, the colour will default to RGB(20, 0, 52)
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="rngInsert"></param>
    ''' <param name="rgbFill"></param>
    ''' <returns></returns>
    Public Function hf_Insert_BackShape(ByRef hf As Word.HeaderFooter, ByRef rngInsert As Word.Range, Optional rgbFill As Integer = -1) As Word.Shape
        Dim shp As Word.Shape
        Dim objWCAGMgr As New cWCAGMgr()
        Dim strShapeName = "aac_BackColour"
        Dim sect As Word.Section
        'Dim rgbFill As Long
        Dim objPnlMgr As New cBackPanelMgr()
        '
        shp = objPnlMgr.pnl_BackPanel_Insert(hf, rngInsert, rgbFill)
        GoTo finis
        '

        sect = rngInsert.Sections.Item(1)
        shp = Nothing

        '
        'Go to the default fill (dark purple if no fill is specified)
        '
        'selectFillType = rgbFill
        'If rgbFill = -1 Then rgbFill = RGB(20, 0, 52)
        '
        'If the hf already has a shape with this name, then delete it.
        'We don't want multiple shapes.. We also get rid of any FreeForm or Rectangle shapes
        'that might haved snuck in
        '
        'Problems when I use hf.Shapes.. Shapes in the prior sections HeadeFooter will also be deleted
        'even though the hfs are not linked
        '
        If hf.Range.ShapeRange.Count <> 0 Then
            For Each shp In hf.Range.ShapeRange
                If shp.Name = strShapeName Or shp.Name Like "Free*" Or shp.Name Like "Rect*" Then
                    shp.Delete()
                    'Exit For
                End If
            Next
        End If
        '
        Try
            '
            shp = hf.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, 0.0, 0.0, sect.PageSetup.PageWidth, sect.PageSetup.PageHeight, rngInsert)
            shp.RelativeHorizontalPosition = Word.WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
            shp.RelativeVerticalPosition = Word.WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
            shp.Left = 0.0
            shp.Top = 0.0
            shp.ZOrder(MsoZOrderCmd.msoSendBehindText)
            shp.ZOrder(MsoZOrderCmd.msoSendToBack)
            shp.Name = strShapeName
            '
            objWCAGMgr.wcag_set_decorative(shp, True)
            '
            'rgbFill = RGB(20, 0, 52)
            rgbFill = RGB(255, 255, 255)

            '
            Select Case rgbFill
                Case -1
                    shp.Fill.ForeColor.RGB = rgbFill
                Case -2
                    objWCAGMgr.wcag_backColour_BorderAndFill(shp)
                Case Else
                    shp.Fill.ForeColor.RGB = rgbFill
            End Select
            'If Not objWCAGMgr.wcag_docProps_isAccessible() Then
            'shp.Fill.ForeColor.RGB = rgbFill
            'Else
            'objWCAGMgr.wcag_backColour_BorderAndFill(shp)
            'End If
            '
            shp.LockAnchor = True
            '
        Catch ex As Exception
            'If we get here it is almost certain that we didn't place the image
            'in the HeaderFooter.. That is Set Hf = Selection.HeaderFooter failed because
            'HeaderFooter was nothin
            '
            MsgBox("Error - Unknown in hf_Insert_BackShape")
        End Try
        '
finis:
        Return shp
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will set the various footer fileds' text to bold or not
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="setToBold"></param>
    ''' <returns></returns>
    Public Function hf_footer_BoldStatus(ByRef myDoc As Word.Document, setToBold As Boolean) As Boolean
        'Dim lst_FooterFlds As New List(Of Word.Field)
        Dim fld As Word.Field
        Dim rng As Word.Range
        Dim rslt As Boolean
        Dim sect As Word.Section
        Dim hf As Word.HeaderFooter
        Dim strCode As String
        Dim setFooter As Boolean
        Dim styl As Word.Style
        '
        rslt = True
        setFooter = False
        '
        'WdFieldType.wdFieldPage
        Try
            For Each sect In myDoc.Sections
                For Each hf In sect.Footers
                    If hf.Exists Then
                        rng = hf.Range
                        For Each fld In rng.Fields
                            strCode = fld.Code.Text
                            If fld.Type = WdFieldType.wdFieldStyleRef Then
                                styl = fld.Result.Style
                                Select Case styl.NameLocal
                                    Case "Footer Text"
                                        setFooter = True
                                    Case "Footer (Letter)"
                                        setFooter = False
                                    Case Else
                                        setFooter = False
                                End Select
                                'lst_FooterFlds.Add(fld)
                                'fld.n
                                If setFooter Then
                                    fld.Result.Font.Bold = setToBold
                                    'fld.Result.Font.Color = WdColor.wdColorRed
                                End If
                            End If
                        Next
                    End If
                Next
            Next
            rslt = True
            '
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
    End Function
    '

    '
    Public Sub hf_hfs_wcagSetTabs(ByRef hf As Word.HeaderFooter)

        If hf.Range.Sections.Item(1).PageSetup.Orientation = WdOrientation.wdOrientPortrait Then
            hf.Range.ParagraphFormat.TabStops.Add(384, WdTabAlignment.wdAlignTabRight)
        Else
            hf.Range.ParagraphFormat.TabStops.Add(630, WdTabAlignment.wdAlignTabRight)
        End If

    End Sub
    '

    '
    ''' <summary>
    ''' This method will convert the table based Header/Footers to a WCAG compliant table or tabbed text, depending 
    ''' on the value of doAsTable (true, complaint table or false, tabbed text
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="strHFType"></param>
    Public Sub hf_hfs_convertToWCAGCompliance(ByRef sect As Word.Section, doAsTable As Boolean, Optional strHFType As String = "footer")
        Dim hf As HeaderFooter
        '
        Select Case strHFType
            Case "footer"
                For Each hf In sect.Footers
                    If hf.Exists Then
                        Try
                            Select Case hf.Index
                                Case WdHeaderFooterIndex.wdHeaderFooterEvenPages
                                Case WdHeaderFooterIndex.wdHeaderFooterFirstPage
                                    'Me.hf_footers_convertToText(hf)
                                    Me.hf_footers_convertToWCAG(hf, doAsTable)
                                Case WdHeaderFooterIndex.wdHeaderFooterPrimary
                                    'Me.hf_footers_convertToText(hf)
                                    Me.hf_footers_convertToWCAG(hf, doAsTable)
                            End Select

                        Catch ex As Exception

                        End Try
                    End If

                Next
            Case "header"
                For Each hf In sect.Headers
                    If hf.Exists Then
                        Try
                            Select Case hf.Index
                                Case WdHeaderFooterIndex.wdHeaderFooterEvenPages
                                Case WdHeaderFooterIndex.wdHeaderFooterFirstPage
                                    Me.hf_headers_convertToWCAG(hf, doAsTable)
                                    'Me.hf_headers_convertToText(hf)
                                Case WdHeaderFooterIndex.wdHeaderFooterPrimary
                                    Me.hf_headers_convertToWCAG(hf, doAsTable)
                                    'Me.hf_headers_convertToText(hf)
                            End Select

                        Catch ex As Exception

                        End Try
                    End If

                Next

        End Select

    End Sub
    '
    ''' <summary>
    ''' This method will convert the header structure in hf into either a WCAG compliant table (doASTable=true)
    ''' or as text (doAsTable = false)
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="doAsTable"></param>
    Public Sub hf_headers_convertToWCAG(ByRef hf As Word.HeaderFooter, Optional doAsTable As Boolean = True)
        '
        If doAsTable Then
            Me.hf_headers_convertToWCAGTbls(hf)
        Else
            Me.hf_headers_convertToText(hf)
        End If
    End Sub
    '
    '
    ''' <summary>
    ''' This method will convert the header structure in hf into either a WCAG compliant table (doASTable=true)
    ''' or as text (doAsTable = false)
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="doAsTable"></param>
    Public Sub hf_footers_convertToWCAG(ByRef hf As Word.HeaderFooter, Optional doAsTable As Boolean = True)
        Dim tbl As Word.Table
        '
        If doAsTable Then
            Me.hf_footers_convertToWCAGTbls(hf)
        Else
            'In the olde version we had to remove the footer tables and just leave tabbed text
            '**Me.hf_footers_convertToText(hf)
            '
            'We now only have to make sure that the footer text is high contrast, so we remove any hand
            'colouring
            If hf.Range.Tables.Count > 0 Then
                For Each tbl In hf.Range.Tables
                    tbl.Range.Font.Color = RGB(0, 0, 0)
                Next
                '
            End If
            '
        End If
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will take the specified header, obtain the header table and then
    ''' ensure that table is 'Accessible'.. It does so by applying a Tbale style
    ''' and turning on the Header row
    ''' </summary>
    ''' <param name="hf"></param>
    Public Function hf_headers_convertToWCAGTbls(ByRef hf As Word.HeaderFooter) As Word.Table
        Dim myDoc As Word.Document
        Dim objWCAGMgr As New cWCAGMgr()
        Dim shp As Word.Shape
        Dim drCell As Word.Cell
        Dim sect As Word.Section
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        Dim objTblMgr As New cTablesMgr(hf.Range.Document)
        Dim tbl As Word.Table
        Dim numHeaderRows As Integer
        Dim tblWidth, leftIndent, cellPadding, leftIndentBody, bodyWidth As Single
        '
        myDoc = hf.Range.Document
        '
        If hf.Range.Tables.Count = 2 Then
            'We have a TOC
            hf.Range.Tables.Item(2).Delete()
            sect = hf.Range.Sections.Item(1)
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            para = rng.Paragraphs.Add(rng)
            rng = para.Range
            para = rng.Paragraphs.Add(rng)
            para.Range.Text = "Contents"
            para.Range.Style = hf.Range.Document.Styles.Item("TOC Heading")
        End If
        '

        tbl = Me.hf_Hfs_getTable(hf, tblWidth)
        If IsNothing(tbl) Then GoTo finis
        '
        objTblMgr.tbl_get_tableProperties(tbl, numHeaderRows, leftIndent, cellPadding, leftIndentBody, bodyWidth)
        '
        '
        'Set any shapes in the first cell of the header table to decorative
        '
        drCell = tbl.Range.Cells.Item(1)
        If drCell.Range.ShapeRange.Count <> 0 Then
            For Each shp In drCell.Range.ShapeRange
                objWCAGMgr.wcag_set_decorative(shp, True)
            Next
        End If

        '
        tbl.Style = myDoc.Styles.Item("aac Table (no lines)")
        tbl.ApplyStyleHeadingRows = True
        '
        objTblMgr.tbl_convert_oneRowTableToWCAG(cellPadding, leftIndentBody, tbl)
        tbl.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalCenter
        '
finis:
        Return tbl
        '
    End Function

    '
    ''' <summary>
    ''' This method will convert the table based Header/Footers to tabbed text
    ''' </summary>
    ''' <param name="hf"></param>
    Public Sub hf_headers_convertToText(ByRef hf As Word.HeaderFooter)
        Dim objWCAGMgr As New cWCAGMgr()
        Dim tbl As Word.Table
        Dim sect As Word.Section
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        Dim tblWidth As Single
        Dim para As Word.Paragraph
        Dim leftIndent As Single
        Dim shp As Word.Shape
        Dim ilShp As Word.InlineShape
        Dim i As Integer
        '
        shp = Nothing

        '
        tbl = Me.hf_Hfs_getTable(hf, tblWidth)
        If IsNothing(tbl) Then GoTo finis
        '
        If hf.Range.Tables.Count = 2 Then
            'We have a TOC
            hf.Range.Tables.Item(2).Delete()
            sect = hf.Range.Sections.Item(1)
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            para = rng.Paragraphs.Add(rng)
            rng = para.Range
            para = rng.Paragraphs.Add(rng)
            para.Range.Text = "Contents"
            para.Range.Style = hf.Range.Document.Styles.Item("TOC Heading")
        End If
        '
        drCell = tbl.Range.Cells.Item(1)
        rng = drCell.Range
        Try
            'If the shape is not present, the Item method will throw an exception
            shp = rng.ShapeRange.Item(("logo_AAC_TandG"))
            objWCAGMgr.wcag_set_decorative(shp, True)
            '
        Catch ex As Exception
            shp = Nothing
        End Try
        '
        If IsNothing(shp) Then
            tbl.Delete()
            GoTo finis
        End If
        '
        'This header is a standard Table based header
        '
        hf.Range.Style = hf.Range.Document.Styles.Item("Header-Company Name")
        leftIndent = tbl.Rows.Item(1).LeftIndent
        '
        Try
            ilShp = shp.ConvertToInlineShape()
            tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByTabs)
            hf.Range.Style = hf.Range.Document.Styles.Item("Header-Company Name")
            '
            'It seems that ShapeRange includes both inline and other shapes
            If hf.Range.Paragraphs.Count = 2 Then
                For i = hf.Range.Paragraphs.Count To 1 Step -1
                    para = hf.Range.Paragraphs.Item(i)
                    If para.Range.ShapeRange.Count = 0 Then para.Range.Delete()
                Next i

            End If
            hf.Range.ParagraphFormat.LeftIndent = leftIndent
            hf.Range.ParagraphFormat.TabStops.ClearAll()
            '
            Me.hf_hfs_wcagSetTabs(hf)

        Catch ex As Exception

        End Try


finis:


    End Sub
    '
    '
    ''' <summary>
    ''' This method will make the footer table Accessible by applying  the style "aac Table (no lines)"
    ''' and then adjusting as appropriate
    ''' </summary>
    ''' <param name="hf"></param>
    Public Sub hf_footers_convertToWCAGTbls(ByRef hf As Word.HeaderFooter)
        Dim tbl As Word.Table
        Dim drCell As Word.Cell

        tbl = Me.hf_headers_convertToWCAGTbls(hf)
        If Not IsNothing(tbl) Then
            tbl.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
            For Each drCell In tbl.Range.Cells
                drCell.BottomPadding = 2
            Next
        End If

    End Sub
    '
    ''' <summary>
    ''' This method expects as input the footer of an aac report. Typically this footer is a
    ''' table with two parts. The first part contains stylerefs, the second the page number
    ''' </summary>
    ''' <param name="hf"></param>
    Public Sub hf_footers_convertToText(ByRef hf As Word.HeaderFooter)
        Dim tbl As Word.Table
        Dim tabCentre As Single
        Dim drCell As Word.Cell
        Dim flds As Word.Fields
        Dim fld As Word.Field
        Dim rng As Word.Range
        Dim styl As Word.Style
        '
        tabCentre = Me.objGlobals.glb_get_tabCenterPos(hf.Range.Sections.Item(1))
        '
        If hf.Range.Tables.Count > 0 Then
            For Each tbl In hf.Range.Tables
                drCell = tbl.Range.Cells.Item(2)
                'If drCell.Range.Fields.Count <> 0 Then
                'fld = drCell.Range.Fields.Item(1)
                'strText = fld.ToString()
                'strText = fld.Result.Text
                'fld.Delete()
                'rng = drCell.Range
                'rng.Collapse(WdCollapseDirection.wdCollapseStart)
                'rng.Text = "text" + strText + " "
                'End If
                '

                drCell = tbl.Range.Cells.Item(1)
                flds = drCell.Range.Fields
                '
                If flds.Count > 0 Then
                    fld = flds.Item(1)
                    If fld.Type = WdFieldType.wdFieldStyleRef Then
                        rng = tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByTabs)
                        styl = rng.Document.Styles.Item("Footer Text")
                        rng.Style = styl
                        hf.Range.ParagraphFormat.TabStops.ClearAll()
                        Me.hf_hfs_wcagSetTabs(hf)
                    End If
                Else
                    'Someone has removed the stylerefs
                    Try
                        rng = tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByTabs)
                        styl = rng.Document.Styles.Item("Footer Text")
                        rng.Style = styl
                        hf.Range.ParagraphFormat.TabStops.ClearAll()
                        Me.hf_hfs_wcagSetTabs(hf)
                    Catch ex As Exception

                    End Try

                End If
            Next
        End If
    End Sub
    '
    ''' <summary>
    ''' This method expects as input the footer of an aac report. Typically this footer is a
    ''' table with two parts. The first part contains stylerefs, the second the page number
    ''' </summary>
    ''' <param name="hf"></param>
    Public Sub xhf_footers_convertToText(ByRef hf As Word.HeaderFooter)
        Dim tbl As Word.Table
        Dim tabCentre As Single
        Dim drCell As Word.Cell
        Dim flds As Word.Fields
        Dim fld As Word.Field
        Dim rng As Word.Range
        Dim styl As Word.Style
        '
        tabCentre = Me.objGlobals.glb_get_tabCenterPos(hf.Range.Sections.Item(1))
        '
        If hf.Range.Tables.Count > 0 Then
            For Each tbl In hf.Range.Tables
                drCell = tbl.Range.Cells.Item(2)
                'If drCell.Range.Fields.Count <> 0 Then
                'fld = drCell.Range.Fields.Item(1)
                'strText = fld.ToString()
                'strText = fld.Result.Text
                'fld.Delete()
                'rng = drCell.Range
                'rng.Collapse(WdCollapseDirection.wdCollapseStart)
                'rng.Text = "text" + strText + " "
                'End If
                '

                drCell = tbl.Range.Cells.Item(1)
                flds = drCell.Range.Fields
                '
                If flds.Count > 0 Then
                    fld = flds.Item(1)
                    If fld.Type = WdFieldType.wdFieldStyleRef Then
                        rng = tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByTabs)
                        styl = rng.Document.Styles.Item("Footer Text")
                        rng.Style = styl
                        hf.Range.ParagraphFormat.TabStops.ClearAll()
                        Me.hf_hfs_wcagSetTabs(hf)
                    End If
                Else
                    'Someone has removed the stylerefs
                    Try
                        rng = tbl.ConvertToText(WdTableFieldSeparator.wdSeparateByTabs)
                        styl = rng.Document.Styles.Item("Footer Text")
                        rng.Style = styl
                        hf.Range.ParagraphFormat.TabStops.ClearAll()
                        Me.hf_hfs_wcagSetTabs(hf)
                    Catch ex As Exception

                    End Try

                End If
            Next
        End If
    End Sub
    '
    '
End Class

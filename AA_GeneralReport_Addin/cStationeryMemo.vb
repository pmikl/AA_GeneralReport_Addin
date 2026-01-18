Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
''' <summary>
''' 
'''This class deals with all things related to letters, memos ect
'''
'''Peter Mikelaitis October 2020...http://mikl.com.au
'''New addition with the 2020 rebrand
'''
''' </summary>
Public Class cStationeryMemo
    Inherits cStationeryLetter
    '
    Public Sub New()
        MyBase.New()
    End Sub
    '
    '
    Public Function memo_is_memo(ByRef sect As Word.Section) As Boolean
        Dim objTagsMgr As New cTagsMgr()
        '
        Return objTagsMgr.tags_is_Memo(sect)
        '
    End Function
    '

    '
    ''' <summary>
    ''' This method will insert a leter at the beginning of the Active Doucment
    ''' </summary>
    ''' <returns></returns>
    Public Function mem_insert_Memo() As Word.Range
        Dim rng As Range
        Dim sect As Section
        Dim objSectMgr As New cSectionMgr()
        Dim objGlobals As New cGlobals()
        Dim lstOfMarginDimensions As Collection
        '
        rng = objGlobals.glb_get_wrdSelRng
        Try
            lstOfMarginDimensions = objGlobals.glb_getDimensions_Letter()
            sect = objGlobals.glb_get_wrdActiveDoc.Sections.Item(1)
            '
            sect = objSectMgr.sct_insert_SectionAtStart(True, lstOfMarginDimensions)
            Me.ltr_hfs_Reset(sect)
            Me.ltr_insert_Contents_Letter(sect)
            '
            sect.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
            '
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
        Catch ex As Exception

        End Try
        '
        Return rng

        '
    End Function
    '
    Public Function Insert_Stationery_Memo(Optional strType As String = "Memorandum") As Word.Range
        Dim rng As Range
        Dim sect As Section
        Dim objSectMgr As New cSectionMgr()
        Dim objGlobals As New cGlobals()
        Dim objBBMgr As cBBlocksHandler
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim lstOfMarginDimensions As Collection
        '
        objBBMgr = New cBBlocksHandler
        lstOfMarginDimensions = objGlobals.glb_getDimensions_Letter()
        sect = objGlobals.glb_get_wrdActiveDoc.Sections.Item(1)
        '
        sect = objSectMgr.sct_insert_SectionAtStart(True, lstOfMarginDimensions)
        '
        objHfMgr.hf_headers_insert(sect)
        'Now do the pages inclding the Headers (adjustments) and the Footers
        '
        Me.do_MemoSetup_FirstPage(sect)
        Me.do_MemoSetup_Follower(sect)
        '
        'lstOfOldSettings = Me.Chpt_Get_HeaderFooterIndents_All(sect)
        'sect.PageSetup.LeftMargin = 55.8
        'Me.Chpt_Reset_HeaderFooterPosition(sect, lstOfOldSettings)
        '
        Me.do_MemoSetup_Contents(sect)
        Me.do_Setup_WriteStationeryType(sect.Range, strType)
        '
        sect.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
        '
        rng = sect.Range
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '        '
        Return rng

        '
    End Function
    '
    '
    Public Sub do_MemoSetup_Follower(ByRef sect As Word.Section)
        Dim objBBMgr As New cBBlocksHandler()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim drLeftIndentDelta As Single
        Dim hf As Word.HeaderFooter
        Dim tbl As Word.Table
        Dim dr As Word.Row
        '
        '
        'The Header of the Letter is slightly different from the standard header.... So we handle
        'this with deltas.. "See Me.HeaderLeftIndentNew"
        '
        'Now Do the Header
        tbl = objHFMgr.hf_get_HeaderTable(sect)
        dr = tbl.Rows.Item(1)
        drLeftIndentDelta = Me.HeaderLeftIndentNew - dr.LeftIndent
        dr.LeftIndent = Me.HeaderLeftIndentNew
        tbl.Columns.Item(1).Width = tbl.Columns.Item(1).Width + Math.Abs(drLeftIndentDelta)
        '
        'Delete any artefact logos in the Header Table
        Me.ltr_insert_Logo(tbl, False)

        '
        hf = sect.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        objHFMgr.hf_footer_insertLetterFollower_AsSWBuild(hf, True, Me.rgbFooterGrey, "memo")

        'hf = sect.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        'rng = hf.Range
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'objBBMgr.insertBuildingBlockFromDefaultLibToRange(strFooterBBlkName, strBBlkCategory, rng)
        '
finis:
    End Sub
    '
    '
    ''' <summary>
    ''' This method will do the Header and Footer of the first page. It assumes that
    ''' the section has a different first page layout
    ''' </summary>
    ''' <param name="sect"></param>
    ''' 
    Public Sub do_MemoSetup_FirstPage(ByRef sect As Word.Section)
        Dim objBBMgr As New cBBlocksHandler()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim rng As Word.Range
        Dim hf As Word.HeaderFooter
        Dim tbl As Word.Table
        Dim dr As Word.Row
        Dim drLeftIndentDelta As Single
        Dim para As Word.Paragraph
        '
        'The Header of the Letter is slightly different from the standard header.... So we handle
        'this with deltas.. "See Me.HeaderLeftIndentNew"
        '
        'Now Do the Header
        tbl = objHFMgr.hf_get_HeaderTable(sect, "firstPage")
        dr = tbl.Rows.Item(1)
        drLeftIndentDelta = Me.HeaderLeftIndentNew - dr.LeftIndent
        dr.LeftIndent = Me.HeaderLeftIndentNew
        tbl.Columns.Item(1).Width = tbl.Columns.Item(1).Width + Math.Abs(drLeftIndentDelta)
        '
        Me.ltr_insert_Logo(tbl)
        '
        'Now do the Footer
        '
        hf = sect.Footers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        rng = hf.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        tbl = rng.Tables.Add(rng, 1, 1)
        'Me.objGlobals.Globals_Table_Fix(tbl, False, False)
        tbl.Rows.Item(1).HeightRule = WdRowHeightRule.wdRowHeightExactly
        tbl.Rows.Item(1).Height = 48.0 - 4.8
        tbl.Range.Cells.Item(1).BottomPadding = 4.8
        '
        dr = tbl.Rows.Item(1)
        drLeftIndentDelta = Me.HeaderLeftIndentNew - dr.LeftIndent
        dr.LeftIndent = Me.HeaderLeftIndentNew
        tbl.Columns.Item(1).Width = tbl.Columns.Item(1).Width + Math.Abs(drLeftIndentDelta)

        '
        tbl.Range.Cells.Item(1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
        tbl.Range.Cells.Item(1).LeftPadding = 0.0
        tbl.Range.Cells.Item(1).RightPadding = 3.6

        tbl.Range.Cells.Item(1).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles("Footer (Letter-Contact)")
        '
        rng = tbl.Range.Cells.Item(1).Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Text = "Melbourne" + vbTab + "Sydney" + vbTab + "Brisbane" + vbTab + "Canberra" + vbTab + "Perth" + vbTab + "Adelaide"

        para = rng.Paragraphs.Item(1)
        para.Format.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
        para.Format.TabStops.Add(106.15, WdAlignmentTabAlignment.wdLeft, WdTabLeader.wdTabLeaderSpaces)
        para.Format.TabStops.Add(205.75, WdAlignmentTabAlignment.wdLeft, WdTabLeader.wdTabLeaderSpaces)
        para.Format.TabStops.Add(305.15, WdAlignmentTabAlignment.wdLeft, WdTabLeader.wdTabLeaderSpaces)
        para.Format.TabStops.Add(404.35, WdAlignmentTabAlignment.wdLeft, WdTabLeader.wdTabLeaderSpaces)
        para.Format.TabStops.Add(492.55, WdAlignmentTabAlignment.wdLeft, WdTabLeader.wdTabLeaderSpaces)

        '
        'Me.Insert_Contact_OfficeDetails(rng, "none")
        '
        'rng = tbl.Range.Cells.Item(2).Range
        'nestedTbl = tbl.Range.Cells.Item(2).Tables.Add(rng, 1, 1)
        'nestedTbl.Borders.Enable = False
        'd 'rCell = nestedTbl.Range.Cells.Item(1)
        'drCell.TopPadding = 0.0
        'drCell.BottomPadding = 2.4
        'drCell.LeftPadding = 0.0
        'drCell.RightPadding = 2.8
        '
        'rng = drCell.Range
        'rng.Text = ""
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        'Me.insert_Contact_OfficesWebAndABN(rng)


        'Call objBBMgr.insertBuildingBlockFromDefaultLibToRange(strFooterBBlkName, strBBlkCategory, rng)

    End Sub

    '
    ''' <summary>
    ''' This method will insert the letter contents at the 
    ''' </summary>
    Public Sub do_MemoSetup_Contents(ByRef sect As Word.Section)
        Dim objBBMgr As New cBBlocksHandler()
        Dim objSectMgr As New cSectionMgr()
        Dim objParas As New cParas()
        Dim tbl As Word.Table
        Dim para As Word.Paragraph
        Dim rng, rng2 As Word.Range
        Dim objTools As New cTools()
        Dim j As Integer
        '
        rng = objParas.paras_Paragraphs_DeleteAll(sect)
        rng.Text = Me.get_Memo_Msg()
        rng.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("Body Text")
        '
        For j = 1 To rng.Paragraphs.Count
            para = rng.Paragraphs.Item(j)
            If j = 1 Then
                para.Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("Stationery Confidentiality")
                para.Format.KeepWithNext = True
            End If
            'If j = 12 Then para.Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("Signature Block")
            If j = 3 Then
                tbl = Me.do_MemoTable_Build(para)
                rng2 = tbl.Range
                rng2.Collapse(WdCollapseDirection.wdCollapseEnd)
                para = rng2.Paragraphs.Item(1)
                para.Range.Delete()
                rng2.MoveEnd(WdUnits.wdParagraph, 8)
                rng2.Collapse(WdCollapseDirection.wdCollapseEnd)
                rng2.Paragraphs.Item(1).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("Signature Block")
                '
            End If
        Next
        '


        'rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange(strBBlkName, strBBlkCategory, rng)
        '
    End Sub
    '
    Public Overrides Function get_Memo_Msg() As String
        Dim strMsg As String
        '
        strMsg = MyBase.get_Memo_Msg()
        '
        'strMsg = strMsg + "Confidential" + vbCrLf + vbCrLf + vbCrLf
        'strMsg = strMsg + "Start body copy here" + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf
        'strMsg = strMsg + "Signature Block" + vbCrLf + vbCrLf
        '
        Return strMsg
    End Function
    '
    Public Overrides Function do_MemoTable_Build(ByRef para As Word.Paragraph) As Word.Table
        Dim rng As Word.Range
        Dim objTools As New cTools()
        Dim pageWidth As Single
        Dim tbl, tblHeader As Word.Table
        Dim i, k As Integer
        Dim drCell As Word.Cell
        Dim drCol As Word.Column
        Dim dr As Word.Row
        Dim sect As Word.Section
        Dim leftIndent As Single
        Dim hf As HeaderFooter
        '
        'Set as default
        leftIndent = -20.0
        '
        sect = para.Range.Sections.Item(1)
        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        pageWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
        '
        rng = para.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        tbl = rng.Tables.Add(rng, 5, 4)
        '
        'MyBase.objGlobals.Globals_Table_Fix(tbl, False, False)
        '
        tbl.Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("StationeryContact")
        '
        tbl.Columns.Item(1).Width = 0.103 * pageWidth           '10.3 % of avaialble width
        tbl.Columns.Item(2).Width = 0.4638 * pageWidth           '10.3 % of avaialble width
        tbl.Columns.Item(3).Width = 0.1173 * pageWidth           '10.3 % of avaialble width
        '
        tbl.Columns.Item(4).Width = pageWidth - tbl.Columns.Item(1).Width - tbl.Columns.Item(2).Width - tbl.Columns.Item(3).Width
        '
        For Each drCell In tbl.Range.Cells
            drCell.TopPadding = 4.0
            drCell.BottomPadding = 3.6
            drCell.LeftPadding = 5.4
            drCell.RightPadding = 5.4
        Next
        '
        For i = 1 To tbl.Columns.Count
            drCol = tbl.Columns.Item(i)
            Select Case i
                Case 1
                    For k = 1 To drCol.Cells.Count
                        drCell = drCol.Cells.Item(k)
                        drCell.Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("StationeryLabel")
                        If k = 2 Then drCell.Range.Text = "To"
                        If k = 3 Then drCell.Range.Text = "From"
                        If k = 4 Then drCell.Range.Text = "Subject"
                    Next
                Case 2
                    For k = 1 To drCol.Cells.Count
                        drCell = drCol.Cells.Item(k)
                        'drCell.Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("StationeryContact")
                        If k = 2 Then drCell.Range.Text = "Recipient"
                        If k = 3 Then drCell.Range.Text = "Author"
                        If k = 4 Then drCell.Range.Text = "Subject"
                    Next
                Case 3
                    For k = 1 To drCol.Cells.Count
                        drCell = drCol.Cells.Item(k)
                        drCell.Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("StationeryLabel")
                        If k = 2 Then drCell.Range.Text = "Date"
                        If k = 3 Then drCell.Range.Text = "Ref"
                    Next
                Case 4
                    For k = 1 To drCol.Cells.Count
                        drCell = drCol.Cells.Item(k)
                        drCell.Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("StationeryContact")
                        If k = 2 Then drCell.Range.Text = "DD.MM.YYYY"
                        If k = 3 Then drCell.Range.Text = "0000"
                    Next

            End Select
        Next
        '
        dr = tbl.Rows.Item(4)
        dr.Cells.Item(3).Merge(dr.Cells.Item(4))
        dr.Cells.Item(2).Merge(dr.Cells.Item(3))
        '
        dr = tbl.Rows.Item(5)
        dr.Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("spacer")
        dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
        dr.Height = 14.4
        '
        dr.Cells.Item(3).Merge(dr.Cells.Item(4))
        dr.Cells.Item(2).Merge(dr.Cells.Item(3))
        dr.Cells.Item(1).Merge(dr.Cells.Item(2))
        '
        rng = tbl.Rows.Item(2).Range
        rng.MoveEnd(WdUnits.wdRow, 3)
        rng.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle
        rng.Borders.InsideLineWidth = WdLineWidth.wdLineWidth050pt
        rng.Borders.InsideColor = RGB(0, 1, 0)
        '
        rng = tbl.Rows.Item(3).Range
        rng.MoveEnd(WdUnits.wdRow, 1)
        rng.Borders.InsideLineWidth = WdLineWidth.wdLineWidth150pt
        '
        tblHeader = hf.Range.Tables.Item(1)
        leftIndent = tblHeader.Rows.Item(1).LeftIndent
        dr = tbl.Rows.Item(1)
        dr.Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("StationeryType")
        dr.LeftIndent = leftIndent
        dr.Cells.Item(1).Width = dr.Cells.Item(1).Width - leftIndent
        '
        For Each drCell In dr.Cells
            drCell.TopPadding = 0.0
            drCell.BottomPadding = 0.0
            drCell.LeftPadding = 0.0
            drCell.RightPadding = 5.4
        Next
        dr.Cells.Merge()
        '
        drCell = tbl.Range.Cells.Item(9)
        drCell.Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("StationeryRef_Memo")
        '
        drCell = tbl.Range.Cells.Item(1)
        drCell.Shading.BackgroundPatternColor = RGB(20, 1, 52)
        drCell.LeftPadding = -leftIndent
        rng = drCell.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Text = "Memorandum"

        Return tbl

    End Function
    '

End Class

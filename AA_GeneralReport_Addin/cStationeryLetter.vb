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
Public Class cStationeryLetter
    Inherits cGlobals

    Public HeaderLeftIndentNew As Single        'The letter has a slightly different Header table left indent

    Public rgbLetterPurple As Integer
    Public rgbFooterGrey As Integer
    '
    Public Sub New()
        MyBase.New()
        '
        'The following are the dimensions for the default Portrait and Landscape dimension.
        'They can be over ridden in derived classes
        '
        'Me.lst_PortraitPage = New Collection()
        '
        'Initialise the Cover Page dimensions
        'topMargin, leftMargin, bottonMargin, rightMargin, gutter, headerDistance, footerDistance
        '
        'Call Me.initPageSettings(Me.lst_PortraitPage, 80.4#, 55.8, 84.0, 49.95#, 0.0#, 23.4, 3.8)
        '
        Me.rgbLetterPurple = RGB(108, 62, 153)
        Me.rgbFooterGrey = RGB(200, 199, 200)
        Me.HeaderLeftIndentNew = -26.15
        '
    End Sub
    '
    ''' <summary>
    ''' For the specified section (sect). This method tests the footer style in Cell 1 of the
    ''' footer table to see if it is 'Footer(Letter)'.. If it is, then it will return True, because
    ''' under normal circumstances only Stationery uses this style in the footer table
    ''' to see 
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function ltr_is_Stationery(ByRef sect As Word.Section) As Boolean
        Dim rslt As Boolean
        Dim objGlobals As New cGlobals()
        Dim hf As HeaderFooter
        Dim tbl As Word.Table
        Dim ftrStyle As Word.Style
        '
        rslt = False
        '
        hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        If hf.Range.Tables.Count <> 0 Then
            tbl = hf.Range.Tables.Item(1)
            ftrStyle = tbl.Range.Cells.Item(1).Range.Style
            If ftrStyle.NameLocal = "Footer (Letter)" Then
                rslt = True
            Else
                rslt = False
            End If
        End If
        '
        Return rslt
        '
    End Function
    '
    Public Function ltr_is_letter(ByRef sect As Word.Section) As Boolean
        Dim objTagsMgr As New cTagsMgr()
        '
        Return objTagsMgr.tags_is_Letter(sect)
        '
    End Function
    '
    ''' <summary>
    ''' This function will return a null string if section (sect) is 'Stationery'. If it is
    ''' not it will return an appropriate error message
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function ltr_isOK_toInsert(ByRef sect As Word.Section) As String
        Dim strMsg As String
        '
        strMsg = ""
        If Me.ltr_is_Stationery(sect) Then
            strMsg = ""
        Else
            strMsg = "This function is only supported in a 'Letter' item."
            strMsg = strMsg + vbCrLf + vbCrLf + "Please relocate your cursor to a Letter"
        End If
        '
        Return strMsg

    End Function
    '
    ''' <summary>
    ''' This method will insert a leter at the beginning of the Active Doucment
    ''' </summary>
    ''' <returns></returns>
    Public Function ltr_insert_Letter_Memo(Optional strStationeryType As String = "letter") As Word.Range
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objBnrMgr As New cChptBanner()
        Dim myDoc As Word.Document
        Dim rng As Range
        Dim sect As Section
        Dim objSectMgr As New cSectionMgr()
        Dim objGlobals As New cGlobals()
        Dim lstOfMarginDimensions As Collection
        '
        rng = objGlobals.glb_get_wrdSelRng
        myDoc = objGlobals.glb_get_wrdActiveDoc()
        '
        Try
            lstOfMarginDimensions = objGlobals.glb_getDimensions_Letter()
            sect = myDoc.Sections.Item(1)
            '
            'sect = objSectMgr.sct_insert_SectionAtStart(True, lstOfMarginDimensions)
            sect = objSectMgr.sct_insert_Section(False, sect,,,,, lstOfMarginDimensions)
            'objHFMgr.hf_headers_delete(sect)
            'objHFMgr.hf_hfs_deleteAll(sect)
            objSectMgr.sct_delete_allSectionContents(sect)
            sect.PageSetup.DifferentFirstPageHeaderFooter = True
            '
            '
            'rng = sect.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'sect = myDoc.Sections.Add(rng, WdSectionStart.wdSectionNewPage)
            '
            'objHFMgr.hf_hfs_linkUnlinkAll(myDoc.Sections.Item(2), False)
            'sect = myDoc.Sections.First

            '
            '
            Me.ltr_hfs_Reset(sect, strStationeryType)
            '
            'GoTo finis
            '
            Select Case strStationeryType
                Case "letter"
                    objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_letter), "primary")
                    objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_letter), "firstPage")
                    rng = Me.ltr_insert_Contents_Letter(sect)
                Case "memo"
                    objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_memo), "primary")
                    objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_memo), "firstPage")
                    Me.ltr_insert_Contents_Memo(sect)
                    Me.do_Setup_WriteStationeryType(sect.Range, "Memo")
                    rng = sect.Range

                Case "briefingNote"
                    objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_brief), "primary")
                    objHFMgr.hf_tags_setTagStyle(sect, objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_brief), "firstPage")
                    Me.ltr_insert_Contents_Memo(sect)
                    Me.do_Setup_WriteStationeryType(sect.Range, "Briefing Note")

            End Select
            '
            sect.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary).PageNumbers.StartingNumber = 1
            '
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
        Catch ex As Exception

        End Try
        '
finis:
        '
        Return rng

        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will delete the headers and footers of th current section and replace it with the
    ''' Headers and Footers for the Letter. Note that the only test it performs is to determine whether
    ''' the section is setup to support DifferentFirstPage. This method returns true if the operation succeeded
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function ltr_hfs_Reset(ByRef sect As Word.Section, Optional strStationeryType As String = "letter") As Boolean
        Dim rslt As Boolean
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim rightEdge, leftEdge As Single
        Dim tbl As Word.Table
        Dim hf As Word.HeaderFooter
        '
        rslt = False
        leftEdge = 28.5
        'rightEdge = Me.glb_hfs_getHFTableEdge(sect, "header_rightEdge")
        If Not sect.PageSetup.DifferentFirstPageHeaderFooter Then GoTo finis
        '
        'Delete existing HFs and General Setup
        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        objHFMgr.hf_hfs_deleteAll(sect)
        rightEdge = sect.PageSetup.RightMargin
        '
        'Headers (First and second page headers are the smae across statioery types
        tbl = objHFMgr.hf_insertFirstPage_Header(hf, leftEdge, rightEdge, True, False)
        Me.ltr_logo_Resize(tbl)
        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        tbl = objHFMgr.hf_insertPrimary_header(hf, leftEdge, rightEdge, False)
        '
        Select Case strStationeryType
            Case "letter"
                hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                tbl = objHFMgr.hf_insertFirstPage_Footer(hf, leftEdge, rightEdge, False, False)
                '
                Me.ltr_hf_doLetterFooterStructure_FirstPage(tbl)
                '
                hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                tbl = objHFMgr.hf_insertPrimary_Footer(hf, leftEdge, rightEdge, False, False)
                'The logo in this instance is the "A" used in the letter footer of the follower
                Me.ltr_hf_doLetterFooterStructure_PrimaryPage(tbl, True, "letter")

            Case "memo", "briefingNote"
                hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
                tbl = objHFMgr.hf_insertFirstPage_Footer(hf, leftEdge, rightEdge, False, False)
                'Adjust to single column table
                tbl.Columns.Item(1).Width = tbl.Columns.Item(1).Width + tbl.Columns.Item(2).Width
                tbl.Columns.Item(2).Delete()
                Me.ltr_hf_doMemoFooterStructure_FirstPage(tbl)
                '
                hf = sect.Footers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
                tbl = objHFMgr.hf_insertPrimary_Footer(hf, leftEdge, rightEdge, False, False)
                Me.ltr_hf_doLetterFooterStructure_PrimaryPage(tbl, True, "memo")

        End Select
        'General setup
        '
        '
finis:
        Return rslt
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will insert the standard AAC header logo into the first cell of the
    ''' header table (headerTable). It will delete any artefact logos before doing this.
    ''' The Logo will then be resized and repositioned to match the letter requirements. Note
    ''' that if doLogo is False, then this method will just delete any existing logos in the
    ''' first cell of the 'headerTable'
    ''' </summary>
    ''' <param name="headerTable"></param>
    ''' <returns></returns>
    Public Function ltr_insert_Logo(ByRef headerTable As Word.Table, Optional doLogo As Boolean = True) As Word.Shape
        Dim objBBMgr As New cBBlocksHandler()
        Dim objBB As Word.BuildingBlock
        Dim rng As Word.Range
        Dim drCell As Word.Cell
        Dim shp As Word.Shape
        '
        shp = Nothing
        objBB = objBBMgr.getBuildingBlockFromDefaultLib("aac_HFs_logoHeader", "headerFooters")
        '
        drCell = headerTable.Range.Cells.Item(1)
        rng = drCell.Range
        '
        'Delete any artefact logos
        For Each shp In rng.ShapeRange
            If shp.Name = "logo_AAC_TandG" Then
                shp.Delete()
            End If
        Next
        '
        If doLogo Then
            drCell = headerTable.Range.Cells.Item(1)
            rng = drCell.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng = objBB.Insert(rng)
            If rng.ShapeRange.Count <> 0 Then
                shp = rng.ShapeRange.Item(1)
                shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
                shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
                shp.Height = 16.5
                shp.Width = 113.1
                shp.Left = 0.25
                shp.Top = 5.2
            End If

        End If
        '
        Return shp
        '
    End Function
    '
    ''' <summary>
    ''' This method wills earch cell 1 of the supplied header table and if it finds a
    ''' AAC Logo it will resize it
    ''' </summary>
    ''' <param name="headerTable"></param>
    Public Sub ltr_logo_Resize(ByRef headerTable As Word.Table)
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        Dim shp As Word.Shape
        '
        Try
            drCell = headerTable.Range.Cells.Item(1)
            rng = drCell.Range
            '
            If rng.ShapeRange.Count <> 0 Then
                '
                For Each shp In rng.ShapeRange
                    If shp.Name = "logo_AAC_TandG" Then
                        shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
                        shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
                        shp.Height = 16.5
                        shp.Width = 113.1
                        shp.Left = 0.25
                        shp.Top = 5.2
                    End If
                Next
            End If

        Catch ex As Exception

        End Try

    End Sub


    '
    Public Sub ltr_insert_OfficesWebAndABN(ByRef rng As Word.Range)
        Dim strMsg As String
        Dim para As Word.Paragraph
        '
        strMsg = "Melbourne   Sydney   Brisbane   Canberra   Perth   Adelaide" + vbCrLf
        'strMsg = strMsg + "acilallen.com.au" + vbCrLf
        strMsg = strMsg + "acilallen.com.au"
        'strMsg = strMsg + "ABN 68 102 652 148"
        '
        rng.Text = strMsg
        For Each para In rng.Paragraphs
            para.Range.Font.Color = Me.rgbLetterPurple
        Next
        'para = rng.Paragraphs.Item(3)
        'para.Range.Font.Color = Me.rgbFooterGrey

    End Sub
    '
    Public Sub ltr_insert_OfficeAddresses_and_ABN(ByRef rng As Word.Range, strOffice As String, Optional doOnlyNameAndABN As Boolean = False)
        Dim strMsg As String
        Dim para As Word.Paragraph
        '
        strMsg = "ACIL Allen Pty Ltd" + vbCrLf + "ABN 68 102 652 148"
        'strMsg = "ACIL Allen Pty Ltd"

        If strOffice = "" Then
            'rng.Text = strMsg
            'rng.Text = strMsg
            For Each para In rng.Paragraphs
                'para.Range.Font.Color = RGB(0, 1, 0)
            Next
            'para = rng.Paragraphs.Item(1)
            'para.Range.Font.Color = Me.rgbLetterPurple
            'para.Range.Font.Bold = True

            'para = rng.Paragraphs.Item(2)
            'para.Range.Font.Color = Me.rgbLetterPurple
            'para.Range.Font.Bold = False

            'GoTo finis
            '
        End If
        '

        '
        If doOnlyNameAndABN Then
            strMsg = "ACIL Allen Pty Ltd" + vbCrLf + "ABN 68 102 652 148"
            '
            Try
                rng.Text = strMsg
                For Each para In rng.Paragraphs
                    para.Range.Font.Color = RGB(0, 1, 0)
                Next
                para = rng.Paragraphs.Item(1)
                para.Range.Font.Color = Me.rgbLetterPurple
                para.Range.Font.Bold = True

                para = rng.Paragraphs.Item(2)
                para.Range.Font.Color = Me.rgbLetterPurple
                para.Range.Font.Bold = False
                'para.SpaceAfter = 2.4

            Catch ex As Exception

            End Try
            '
        Else
            strMsg = strMsg + vbCrLf
            '
            Select Case strOffice
                Case "Melbourne"
                    strMsg = strMsg + "Suite 4, Level 19; North Tower 80 Collins Street Melbourne VIC 3000" + vbCrLf
                    strMsg = strMsg + "P +61 3 8650 6000"
                Case "Sydney"
                    strMsg = strMsg + "Suite 603, Level 6, 309 Kent Street Sydney NSW 2000" + vbCrLf
                    strMsg = strMsg + "P +61 2 8272 5100"
                Case "Brisbane"
                    strMsg = strMsg + "Level 15/127 Creek Street Brisbane QLD 4000" + vbCrLf
                    strMsg = strMsg + "P +61 7 3009 8700"
                Case "Canberra"
                    strMsg = strMsg + "Level 6/54 Marcus Clarke Street Canberra ACT 2601" + vbCrLf
                    strMsg = strMsg + "P +61 2 6103 8200"

                Case "Perth"
                    strMsg = strMsg + "Level 12/28 The Esplanade Perth WA 6000" + vbCrLf
                    strMsg = strMsg + "P +61 8 9449 9600"

                Case "Adelaide"
                    strMsg = strMsg + "167 Flinders Street Adelaide SA 5000" + vbCrLf
                    strMsg = strMsg + "P +61 8 8122 4965"
                '
                Case "none"
                    strMsg = ""
                Case Else
                    strMsg = "ACIL Allen Consulting Pty Ltd"
            End Select
            '
            '
            Try
                rng.Text = strMsg
                For Each para In rng.Paragraphs
                    para.Range.Font.Color = RGB(0, 1, 0)
                Next
                para = rng.Paragraphs.Item(1)
                para.Range.Font.Color = Me.rgbLetterPurple
                para.Range.Font.Bold = True

                para = rng.Paragraphs.Item(2)
                para.Range.Font.Color = Me.rgbLetterPurple
                para.Range.Font.Bold = False
                para.SpaceAfter = 3.0
                '
                para = rng.Paragraphs.Item(4)
                para.Range.Characters.Item(1).Font.Color = Me.rgbLetterPurple
                para.Range.Characters.Item(1).Font.Bold = True

            Catch ex As Exception

            End Try



        End If
        '
finis:

    End Sub
    '
    ''' <summary>
    ''' This method will do the Header and Footer of the first page. It assumes that
    ''' the section has a different first page layout
    ''' </summary>
    ''' <param name="sect"></param>
    ''' 
    Public Sub do_LetterSetup_FirstPage(ByRef sect As Word.Section)
        Dim objBBMgr As New cBBlocksHandler()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim rng As Word.Range
        Dim hf As Word.HeaderFooter
        Dim tbl, nestedTbl As Word.Table
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim leftIndent, tblWidth As Single
        Dim drLeftIndentDelta As Single
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
        hf = sect.Footers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        rng = hf.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        tbl = rng.Tables.Add(rng, 1, 2)
        '
        tbl.Borders.Enable = False
        tbl.TopPadding = 0.0
        tbl.BottomPadding = 0.0
        '
        'Change width of letter footer table here
        '
        'tbl.Columns.Item(1).Width = 201.6
        'tbl.Columns.Item(2).Width = 326.95
        tbl.Columns.Item(1).Width = 250
        tbl.Columns.Item(2).Width = 278.55

        tbl.Rows.Item(1).HeightRule = WdRowHeightRule.wdRowHeightExactly
        tbl.Rows.Item(1).Height = 48.0
        '
        tblWidth = tbl.Columns.Item(1).Width + tbl.Columns.Item(2).Width
        '
        leftIndent = -(tblWidth - (sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin))
        tbl.Rows.Item(1).LeftIndent = leftIndent
        '
        '
        tbl.Range.Cells.Item(1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
        tbl.Range.Cells.Item(1).LeftPadding = 5.4
        tbl.Range.Cells.Item(1).RightPadding = 3.6
        '
        'tbl.Range.Cells.Item(1.BottomPadding = 2.4
        '
        tbl.Range.Cells.Item(2).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
        tbl.Range.Cells.Item(2).LeftPadding = 5.4
        tbl.Range.Cells.Item(2).RightPadding = 0.0
        '
        'tbl.Range.Cells.Item(2).BottomPadding = 2.4
        '

        tbl.Range.Cells.Item(1).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles("Footer (Letter)")
        tbl.Range.Cells.Item(2).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles("Footer (Letter-Contact)")
        '
        rng = tbl.Range.Cells.Item(1).Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        Me.ltr_insert_OfficeAddresses_and_ABN(rng, "", True)
        '
        rng = tbl.Range.Cells.Item(2).Range
        nestedTbl = tbl.Range.Cells.Item(2).Tables.Add(rng, 1, 1)
        nestedTbl.Borders.Enable = False
        drCell = nestedTbl.Range.Cells.Item(1)
        drCell.TopPadding = 0.0
        'drCell.BottomPadding = 2.4
        drCell.BottomPadding = 0.0
        drCell.LeftPadding = 0.0
        drCell.RightPadding = 2.8
        '
        rng = drCell.Range
        rng.Text = ""
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        Me.ltr_insert_OfficesWebAndABN(rng)
        '
finis:

        'Call objBBMgr.insertBuildingBlockFromDefaultLibToRange(strFooterBBlkName, strBBlkCategory, rng)

    End Sub

    '
    Public Sub do_LetterSetup_Follower(ByRef sect As Word.Section)
        Dim objBBMgr As New cBBlocksHandler()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim hf As Word.HeaderFooter
        Dim tbl As Word.Table
        Dim dr As Word.Row
        Dim drLeftIndentDelta As Single
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
        objHFMgr.hf_footer_insertLetterFollower_AsSWBuild(hf, True, Me.rgbFooterGrey, "letter")
        '
finis:
    End Sub
    ''' <summary>
    ''' This method will insert the letter contents at the 
    ''' </summary>
    Public Function ltr_insert_Contents_Letter(ByRef sect As Word.Section) As Word.Range
        Dim objSectMgr As New cSectionMgr()
        Dim objParas As New cParas()
        Dim rng As Word.Range
        Dim strMsg As String
        Dim para As Word.Paragraph
        Dim i As Integer
        Dim myDoc As Word.Document
        '

        myDoc = objSectMgr.objGlobals.glb_get_wrdActiveDoc()
        '
        rng = objParas.paras_Paragraphs_DeleteAll(sect)
        strMsg = ""
        '
        strMsg = "Confidential" + vbCrLf
        strMsg = strMsg + "Date" + vbCrLf
        strMsg = strMsg + "Reference: 0000" + vbCrLf
        '
        strMsg = strMsg + "Name" + vbCrLf
        strMsg = strMsg + "Title" + vbCrLf
        strMsg = strMsg + "Company" + vbCrLf
        strMsg = strMsg + "Address" + vbCrLf
        strMsg = strMsg + "City State Postcode" + vbCrLf
        '
        strMsg = strMsg + "Letter subject – delete this line if not required" + vbCrLf
        '
        strMsg = strMsg + "Dear" + vbCrLf
        strMsg = strMsg + "Insert body copy" + vbCrLf
        '
        strMsg = strMsg + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf
        '
        strMsg = strMsg + "Signature block" + vbCrLf + vbCrLf
        '
        rng.Text = strMsg
        rng.Style = myDoc.Styles.Item("Body Text")
        '
        For i = 1 To rng.Paragraphs.Count
            para = rng.Paragraphs.Item(i)
            If i = 1 Then
                para.Range.Style = myDoc.Styles.Item("Letter Confidentiality")
                para.Format.KeepWithNext = True
            End If
            If i = 2 Then para.Range.Style = myDoc.Styles.Item("LetterDate")
            If i = 3 Then para.Range.Style = myDoc.Styles.Item("LetterRef")
            If i = 4 Then para.Range.Style = myDoc.Styles.Item("LetterName")
            If i = 5 Then para.Range.Style = myDoc.Styles.Item("LetterTitle")
            If i = 6 Then para.Range.Style = myDoc.Styles.Item("LetterCompany")
            If i = 7 Then para.Range.Style = myDoc.Styles.Item("LetterAddress")
            If i = 8 Then para.Range.Style = myDoc.Styles.Item("LetterCityStatePostcode")
            If i = 9 Then para.Range.Style = myDoc.Styles.Item("LetterSubject")
            If i = 20 Then para.Range.Style = myDoc.Styles.Item("Signature Block")

        Next
        '
        para = rng.Paragraphs.Item(2)
        rng = para.Range
        rng.MoveEnd(WdUnits.wdCharacter, -1)
        '
        Return rng
        'rng = objBBMgr.insertBuildingBlockFromDefaultLibToRange(strBBlkName, strBBlkCategory, rng)
        '
    End Function
    '
    Public Sub ltr_insert_Contents_Memo(ByRef sect As Word.Section)
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
    End Sub
    '
    '
    Public Sub ltr_insert_Contents_BriefingNote(ByRef sect As Word.Section)

    End Sub


    ''' <summary>
    ''' This method expects as input a Primary Footer Table (tbl). It will reformat that table, then add a logo
    ''' if doLogo is true. It will use the colour rgbFooter as the colour of the footer text. If rgbfooter is less than 0,
    ''' then the default 'Me.rgbFooterGrey' is used. If greater or equal to 0, then the value of rgbfooter is used.
    ''' The method will introduce variationsand depending on the value of strStationeryType ("letter", "memo" and "briefingNote")
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="doLogo"></param>
    ''' <param name="rgbFooter"></param>
    ''' <param name="strStationeryType"></param>
    ''' <returns></returns>
    Public Function ltr_hf_doLetterFooterStructure_PrimaryPage(ByRef tbl As Word.Table, doLogo As Boolean, Optional strStationeryType As String = "letter", Optional rgbfooter As Integer = -1) As Word.Table
        Dim oBjBBMgr As New cBBlocksHandler()
        Dim objBB As Word.BuildingBlock
        Dim sect As Word.Section
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim shp As Word.Shape
        Dim rng As Word.Range
        Dim drCell As Word.Cell
        Dim dr As Word.Row
        Dim tblWidth, leftIndent As Single
        Dim para As Word.Paragraph
        Dim fld As Word.Field
        '
        sect = tbl.Range.Sections.Item(1)
        If rgbfooter < 0 Then rgbfooter = Me.rgbFooterGrey
        '
        'hf = sect.Footers(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        'objHFMgr.hf_footer_insertLetterFollower_AsSWBuild(hf, doLogo, Me.rgbFooterGrey, "letter")
        '
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
            tbl.Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles("Footer (Letter-Logo)")
            drCell = tbl.Range.Cells.Item(2)
            drCell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
            '
            drCell = tbl.Range.Cells.Item(1)
            drCell.Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles("Footer (Letter)")
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
                rng.Font.Color = rgbfooter
            Next

        Catch ex As Exception

        End Try
        '
        If doLogo Then
            'rng = hf.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            '
            objBB = oBjBBMgr.getBuildingBlockFromDefaultLib("aac_HFs_logoFooter_Small", "headerFooters")
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
        '
    End Function
    '

    Public Sub ltr_hf_doMemoFooterStructure_FirstPage(ByRef tbl As Word.Table)
        Dim dr As Word.Row
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        Dim drLeftIndentDelta As Single
        '
        'hf = sect.Footers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        'rng = hf.Range
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'tbl = rng.Tables.Add(rng, 1, 1)
        'Me.objGlobals.Globals_Table_Fix(tbl, False, False)
        '

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

    End Sub

    '
    Public Sub ltr_hf_doLetterFooterStructure_FirstPage(ByRef tbl As Word.Table)
        Dim sect As Word.Section
        Dim tblWidth, leftIndent As Single
        Dim rng As Word.Range
        Dim nestedTbl As Word.Table
        Dim drCell As Word.Cell
        '
        sect = tbl.Range.Sections.Item(1)
        'Change width of letter footer table here
        '
        'tbl.Columns.Item(1).Width = 201.6
        'tbl.Columns.Item(2).Width = 326.95
        tbl.Columns.Item(1).Width = 250
        tbl.Columns.Item(2).Width = 278.55

        tbl.Rows.Item(1).HeightRule = WdRowHeightRule.wdRowHeightExactly
        tbl.Rows.Item(1).Height = 48.0
        '
        tblWidth = tbl.Columns.Item(1).Width + tbl.Columns.Item(2).Width
        '
        leftIndent = -(tblWidth - (sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin))
        tbl.Rows.Item(1).LeftIndent = leftIndent
        '
        tbl.Range.Cells.Item(1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
        tbl.Range.Cells.Item(1).LeftPadding = 5.4
        tbl.Range.Cells.Item(1).RightPadding = 3.6
        '
        '*** Bottom padding of the host cell(s)
        tbl.Range.Cells.Item(1).BottomPadding = 2.4
        '
        tbl.Range.Cells.Item(2).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
        tbl.Range.Cells.Item(2).LeftPadding = 5.4
        tbl.Range.Cells.Item(2).RightPadding = 0.0
        '
        '*** Bottom padding of the host cell(s)
        tbl.Range.Cells.Item(2).BottomPadding = 2.4
        '
        tbl.Range.Cells.Item(1).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles("Footer (Letter)")
        tbl.Range.Cells.Item(2).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles("Footer (Letter-Contact)")
        '
        rng = tbl.Range.Cells.Item(1).Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        Me.ltr_insert_OfficeAddresses_and_ABN(rng, "", True)
        '
        rng = tbl.Range.Cells.Item(2).Range
        nestedTbl = tbl.Range.Cells.Item(2).Tables.Add(rng, 1, 1)
        nestedTbl.Borders.Enable = False
        drCell = nestedTbl.Range.Cells.Item(1)
        drCell.TopPadding = 0.0
        'drCell.BottomPadding = 2.4
        drCell.LeftPadding = 0.0
        drCell.RightPadding = 2.8
        '
        rng = drCell.Range
        rng.Text = ""
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        Me.ltr_insert_OfficesWebAndABN(rng)
        '
        'rng = tbl.Range.Cells.Item(1).Range
        'Me.ltr_insert_OfficeAddresses_and_ABN(rng, "", False)
        '
    End Sub
    '
    ''' <summary>
    ''' This method will check to see if the current selection is in a Stationery item (e.g. Letter) and
    ''' if it is, it will clear the current office addres and replace it with the address associated
    ''' with the off strOfficeName ("Melbourne", "Sydney", "Brisbane", "Canberra", "Perth", "Hobart", "Adelaide"
    ''' </summary>
    ''' <param name="strOfficeName"></param>
    ''' <param name="doVerbose"></param>
    Public Sub ltr_insert_OfficeAddress_intoFooter(strOfficeName As String, Optional doVerbose As Boolean = True)
        Dim strMsg As String
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim rng As Word.Range

        '
        strMsg = Me.ltr_isOK_toInsert(Me.glb_get_wrdSect())
        If strMsg = "" Then
            rng = objHFMgr.hf_footers_LetterContacts_Clear(Me.glb_get_wrdSect())
            If Not IsNothing(rng) Then
                Me.ltr_insert_OfficeAddresses_and_ABN(rng, strOfficeName)
            End If
        Else
            If doVerbose Then MsgBox(strMsg)
        End If
        '
    End Sub
    '
    Public Overridable Function get_Memo_Msg() As String
        Dim strMsg As String
        '
        strMsg = ""
        '
        strMsg = strMsg + "Confidential" + vbCrLf + vbCrLf + vbCrLf
        strMsg = strMsg + "Start body copy here" + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf + vbCrLf
        strMsg = strMsg + "Signature Block" + vbCrLf + vbCrLf
        '
        Return strMsg
    End Function
    '
    Public Overridable Function do_MemoTable_Build(ByRef para As Word.Paragraph) As Word.Table
        Dim rng As Word.Range
        Dim objTools As New cTools()
        Dim pageWidth As Single
        Dim tbl As Word.Table
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
        leftIndent = 0.0
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
        'tblHeader = hf.Range.Tables.Item(1)
        'leftIndent = tblHeader.Rows.Item(1).LeftIndent
        dr = tbl.Rows.Item(1)
        dr.Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("StationeryType")
        dr.LeftIndent = leftIndent
        dr.Cells.Item(1).Width = dr.Cells.Item(1).Width - leftIndent
        '
        'GoTo finis
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
        'drCell.LeftPadding = -leftIndent
        drCell.LeftPadding = 5.4
        rng = drCell.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        rng.Text = "Memorandum"

finis:
        Return tbl

    End Function
    '
    ''' <summary>
    ''' This method expects a range (rng) that contains a Stationery Block.. It will write the 'strStationeryType'
    ''' into that block
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="strStationeryType"></param>
    Public Sub do_Setup_WriteStationeryType(ByRef rng As Word.Range, strStationeryType As String)
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        '
        If rng.Tables.Count <> 0 Then
            tbl = rng.Tables.Item(1)
            drCell = tbl.Range.Cells.Item(1)
            drCell.Range.Text = strStationeryType
        End If

    End Sub

End Class

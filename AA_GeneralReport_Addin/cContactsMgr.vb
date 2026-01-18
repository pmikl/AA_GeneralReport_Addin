Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.Drawing
Imports System.IO

Public Class cContactsMgr
    Public strTagStyleName_Front As String
    Public strTagStyleName_Back As String

    Public objGlobals As cGlobals
    Public objParas As cParas

    Public Sub New()
        Me.strTagStyleName_Front = "tag_contactsPage-Front"
        Me.strTagStyleName_Back = "tag_contactsPage-Back"
        '
        Me.objGlobals = New cGlobals()
        Me.objParas = New cParas()
        '
    End Sub
    '
    '
    Public Sub contacts_Build_Background(ByRef sect As Word.Section, strTagStyleForHeader As String, Optional deleteHeaderFooters As Boolean = True)
        Dim objBrndMgr As New cBrandMgr()
        '
        Try
            objBrndMgr.brnd_Rebuild_Background(sect, deleteHeaderFooters, False, strTagStyleForHeader)
        Catch ex As Exception
            objBrndMgr.brnd_Rebuild_Background(sect, deleteHeaderFooters, False, "Header")
        End Try
        '
    End Sub
    '
    ''' <summary>
    ''' This method will convert the existing section (sect) to a standard Front Contacts Page
    ''' It assumes that the orientation is correct.
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function contacts_convert_toFrontContacts(ByRef sect As Word.Section, Optional doBottomOfPageImage As Boolean = True) As Word.Table
        Dim lstOfDimensions As Collection
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim hf As Word.HeaderFooter
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        '
        'doBottomOfPageImage = False
        'doBottomOfPageImage = True


        lstOfDimensions = Me.objGlobals.glb_getDimensions_Contacts_Prt("front")
        objGlobals.glb_setDimensions(sect, lstOfDimensions)
        '
        '
        objHfMgr.hf_headers_insert(sect, -1, False)
        'objHfMgr.hf_headers_delete(sect)
        objHfMgr.hf_footers_delete(sect)
        '
        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        rng = hf.Range
        '
        'Now set the Front Contacts Page tagStyle
        Try
            tbl = rng.Tables.Item(1)
            drCell = tbl.Range.Cells.Item(1)
            drCell.Range.Style = objGlobals.glb_get_wrdActiveDoc.Styles.Item(Me.strTagStyleName_Front)
        Catch ex As Exception

        End Try
        '
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        'rng.Move(WdUnits.wdParagraph, -1)
        objHfMgr.hf_Insert_BackShape(hf, rng, RGB(255, 255, 255))
        '
        tbl = Me.insert_Contacts_Table_Front(sect, True, doBottomOfPageImage)

        '
        Return tbl
    End Function
    '
    Public Function contacts_convert_toBackContacts(ByRef sect As Word.Section, Optional doBottomOfPageImage As Boolean = True) As Word.Table
        Dim lstOfDimensions As Collection
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim hf As Word.HeaderFooter
        Dim rng As Word.Range
        Dim tbl As Word.Table
        '
        'doBottomOfPageImage = False
        'doBottomOfPageImage = True
        'The Back Contacts Page is Different First Page... If this section isn't we need to force it
        '
        If Not sect.PageSetup.DifferentFirstPageHeaderFooter Then sect.PageSetup.DifferentFirstPageHeaderFooter = True

        lstOfDimensions = Me.objGlobals.glb_getDimensions_Contacts_Prt("front")
        objGlobals.glb_setDimensions(sect, lstOfDimensions)
        '
        tbl = Nothing
        '
        objHfMgr.hf_headers_insert(sect, -1, False)
        'objHfMgr.hf_headers_delete(sect)
        objHfMgr.hf_footers_delete(sect)
        '
        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
        objHfMgr.hf_tags_setTagStyle(sect, Me.strTagStyleName_Back, "primary")
        rng = hf.Range       '
        '
        If sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage).Exists Then
            hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
            objHfMgr.hf_tags_setTagStyle(sect, Me.strTagStyleName_Back, "firstPage")
            rng = hf.Range       '
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            objHfMgr.hf_Insert_BackShape(hf, rng)
            '
            tbl = Me.insert_Contacts_Table_Back(sect)
            '
        End If
        '
        'Now set the Front Contacts Page tagStyle
        Try
            'tbl = rng.Tables.Item(1)
            'drCell = tbl.Range.Cells.Item(1)
            'drCell.Range.Style = objGlobals.glb_get_wrdActiveDoc.Styles.Item(Me.strTagStyleName_Back)
        Catch ex As Exception

        End Try
        '
        '
        Return tbl
    End Function

    '
    '
    ''' <summary>
    ''' This method will insert a front contacts page in front of the section that contains the
    ''' current selection... If there is a banner, it will delete the banner, insert the section and paste
    ''' the banner back.. the selection is left in its orignal position
    ''' </summary>
    ''' <returns></returns>
    Public Function contacts_insert_FrontPage(placeBehind As Boolean, ByRef sect As Word.Section, strRptMode As String, doBottomOfPageImage As Boolean) As Word.Section
        Dim objSectMgr As New cSectionMgr()
        Dim objRptMgr As New cReport()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim sectNew As Word.Section
        'Dim strRptMode As String
        'Dim rng As Word.Range
        Dim tbl As Word.Table
        '
        'sect = Me.objGlobals.glb_get_wrdSel.Sections.Item(1)
        sectNew = sect
        'strRptMode = objRptMgr.Rpt_Mode_Get()
        'rng = Me.objGlobals.glb_get_wrdSelRng()
        '
        If Not Me.conts_has_ContactsPage_Front() Then
            Select Case strRptMode
                Case objRptMgr.rpt_isPrt, objRptMgr.modeShort
                    sectNew = objSectMgr.sct_insert_Section(placeBehind, sect, 6, "newPage", False, "prt", Me.objGlobals.glb_getDimensions_Contacts_Prt("front"))
                    tbl = contacts_convert_toFrontContacts(sectNew, doBottomOfPageImage)
                    'tbl.Range.Rows.Item(3).Delete()
                    '
                    'objHFMgr.hf_footers_delete(sectNew)
                    'objHFMgr.hf_headers_delete(sectNew)
                    'objHFMgr.hf_headers_insert(sectNew, -1, False)
                    'objHFMgr.hf_headers_delete(sectNew)
                    'tbl = Me.insert_Contacts_Table_Front(sectNew)
                    '

                    'objHFMgr.hf_headers_insert(sectNew, -1.0, False)
                    '
                    'objGlobals.glb_screen_update(False)
                    '



                Case objRptMgr.rpt_isLnd
                    sectNew = objSectMgr.sct_insert_Section(placeBehind, sect, 6, "newPage", False, "lnd", Me.objGlobals.glb_getDimensions_Contacts_Prt("front"))
                    tbl = contacts_convert_toFrontContacts(sectNew, doBottomOfPageImage)


                    'sectNew = objSectMgr.sct_insert_SectionInFront(rng,,,, "lnd", Me.objGlobals.glb_getDimensions_ContactsPage("front"))
                    'tbl = Me.insert_Contacts_Table_Front(sectNew)
                    'Adjust Table height to fit Landscape
                    '
                    '**** This code adjusts (after the fact) the row that contains the Acknowledgent of Country and the
                    'aboriginal art
                    'dr = tbl.Range.Rows.Item(3)
                    'dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
                    'dr.Height = 470.0
                    'dr.Delete()
                    '
            End Select
            '
            'Me.contacts_Build_Background(sectNew, Me.strTagStyleName_Front, False)
            '
        Else
            MsgBox("This document already has a front Contacts Page")
        End If
        '
        Return sectNew
        '
    End Function
    '
    Public Function contacts_insert_BackPage() As Word.Section
        Dim objSectMgr As New cSectionMgr()
        Dim objRptMgr As New cReport()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objParas As New cParas()
        Dim lstOfDimensions As Collection
        Dim strRptMode As String
        Dim sect, sectNew As Word.Section
        Dim myDoc As Word.Document
        Dim tbl As Word.Table
        Dim strOrientation As String
        '
        sectNew = Nothing
        myDoc = Me.objGlobals.glb_get_wrdActiveDoc()
        sect = myDoc.Sections.Last
        '
        strOrientation = "prt"
        If Me.objGlobals.glb_get_wrdSect().PageSetup.Orientation = WdOrientation.wdOrientLandscape Then strOrientation = "lnd"
        '
        strRptMode = objRptMgr.Rpt_Mode_Get()
        '
        lstOfDimensions = New Collection()
        sect = objSectMgr.sct_insert_SectionAtEnd(lstOfDimensions)
        '
        tbl = contacts_convert_toBackContacts(sect)
        '
        objHFMgr.hf_headers_insert(sectNew, -1.0, False)
        'objHFMgr.hf_headers_resize_all(sectNew)
        Me.contacts_Build_Background(sectNew, Me.strTagStyleName_Back, False)
        '
        '**objHFMgr.hf_headers_DeleteContents_All(sectNew)
        'objHFMgr.hf_footers_delete(sectNew)

finis:
        Return sect
        '
    End Function

    '
    ''' <summary>
    ''' This method will return true if a front contacts page is found... The variable sect contains
    ''' the section occupied by that page
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function conts_get_getContactsPageFront(ByRef myDoc As Word.Document, ByRef sect As Word.Section) As Boolean
        Dim rslt As Boolean
        '
        rslt = False
        For Each sect In myDoc.Sections
            If Me.is_ContactsPage_Front(sect) Then
                rslt = True
                Exit For
            End If
        Next
        '
        Return rslt
        '
    End Function
    '
    '
    '
    'This method will insert the Report or Proposal citation into the Front
    'the selected Contacts page
    Public Sub contacts_insert_Citation(strPage As String, strType As String)
        Dim rng As Range
        Dim rng3 As Range
        Dim dr As Word.Row
        Dim drHostCell As Word.Cell
        Dim drCell As Cell
        Dim tbl As Word.Table
        Dim objTools As New cTools()
        Dim sect As Word.Section
        Dim objBBMgr As New cBBlocksHandler()
        Dim objLegal As New cLegalAndAbout()
        '
        sect = Globals.ThisAddin.Application.Selection.Range.Sections.Item(1)
        'objTools.upDateCopyRightNotice()
        '
        Select Case strPage
            Case "front"
                If Me.is_ContactsPage_Front(sect) Then
                    '
                    drHostCell = Me.conts_get_hostCell(sect.Range.Tables.Item(1))
                    '
                    If IsNothing(drHostCell) Then GoTo finis
                    'drHostCell = sect.Range.Tables.Item(1).Range.Cells.Item(2)
                    tbl = drHostCell.Tables.Item(1)
                    '
                    'Counting from the bottom we get the row (cell) with the big
                    'Reliance and disclaimer statement
                    dr = tbl.Rows.Item(tbl.Rows.Last.Index - 1)
                    drCell = dr.Cells.Item(1)
                    rng = drCell.Range
                    '
                    'Now lets get the SUggested citation for this report 
                    dr = tbl.Rows.Item(tbl.Rows.Last.Index - 2)
                    drCell = dr.Cells.Item(1)
                    rng3 = drCell.Range

                    '
                    'drCell = tbl.Range.Cells.Item(4)
                    'rng = drCell.Range
                    'rng3 = tbl.Range.Cells.Item(3).Range

                    '*** Envelope Version
                    'drCell = Globals.ThisAddin.Application.Selection.Range.Tables(1).Range.Cells(6)
                    'drCell4 = Globals.ThisAddin.Application.Selection.Range.Tables(1).Range.Cells(4)
                    'rng = drCell.Range
                    '
                    '***rng3.Delete()
                    'rng.Delete()
                    'rng.Select()

                    Select Case strType
                        Case "report_to", "proposal_to"
                            rng3.Delete()
                            rng = objLegal.insert_SuggestedCitation(rng3, Globals.ThisAddin.Application.ActiveDocument, strType)
                            rng.Select()
                            'objLegal.insert_disclaimer(rng, Globals.ThisAddin.Application.ActiveDocument)
                            'rng = objBBMgr.insertBuildingBlockFromDefaultLib_ReturnRangeOfBlock("aac_Cts_frontCitation_Current", "ContactsPage")
                            'rng3.Select()
                            'rng3 = objBBMgr.insertBuildingBlockFromDefaultLib_ReturnRangeOfBlock("aac_Cts_frontSuggested", "ContactsPage")
                            'rng.Select()
                        Case "xproposal_to"
                            'objLegal.insert_SuggestedCitation(rng3, Globals.ThisAddin.Application.ActiveDocument)
                            'objLegal.insert_disclaimer(rng, Globals.ThisAddin.Application.ActiveDocument)

                        Case "copyrightStatement"
                            rng.Delete()
                            rng.Select()
                            '
                            rng = objLegal.insert_CopyrightStatement(rng, Globals.ThisAddin.Application.ActiveDocument)
                            'rng = objBBMgr.insertBuildingBlockFromDefaultLib_ReturnRangeOfBlock("aac_Cts_copyRightStatement", "ContactsPage")
                            rng.Select()
                        Case "disclaimer"
                            rng.Delete()
                            rng.Select()
                            '
                            rng = objLegal.insert_disclaimer(rng, Globals.ThisAddin.Application.ActiveDocument)
                            '
                        Case Else
                    End Select
                    'MsgBox ("Font Page citation")
                Else
                    MsgBox("Please ensure that your cursor is located somewhere on the Front Contacts Page")
                End If
            Case "back"
            Case Else
        End Select
        '
finis:
    End Sub
    '

    '
    ''' <summary>
    ''' This method will return the first cell in the table (tbl) that contains a 
    ''' nested table. If the cell contains more than one nested table it will return
    ''' the first table. If there are no cells with nested tables it will return nothing
    ''' '
    ''' The modified version returns the last cell.. It is this one that contains the table
    ''' of interest
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function conts_get_hostCell(ByRef tbl As Word.Table) As Word.Cell
        Dim drCell As Word.Cell
        Dim dr As Word.Row
        '
        drCell = Nothing

        Try
            dr = tbl.Rows.Last
            drCell = dr.Cells.Item(1)
            'For Each drCell In tbl.Range.Cells
            'If drCell.Tables.Count <> 0 Then
            'GoTo finis
            'End If
            'Next
            'drCell = Nothing
            '
        Catch ex As Exception
            drCell = Nothing
        End Try
        '
finis:
        Return drCell
    End Function
    '
#Region "Contacts Tables"
    '
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="verticalAlignmentIsBottom"></param>
    ''' <param name="doBottomPageImage"></param>
    ''' <param name="doAccessibleVersion"></param>
    ''' <returns></returns>
    Public Function insert_Contacts_Table_Front(ByRef sect As Word.Section, verticalAlignmentIsBottom As Boolean, doBottomPageImage As Boolean, Optional doAccessibleVersion As Boolean = False) As Word.Table
        Dim myDoc As Word.Document
        Dim objLegal As New cLegalAndAbout()
        Dim objParas As New cParas()
        Dim rng As Word.Range
        Dim tbl, tblNestReliance, tblNestAboutAAC As Word.Table
        '
        '
        rng = sect.Range
        '
        myDoc = rng.Document
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        'tbl = rng.Tables.Add(rng, 2, 1)
        tbl = rng.Tables.Add(rng, 3, 1)
        tbl = Me.conts_frontContacts_tableContainer(tbl, verticalAlignmentIsBottom)
        '
        rng = Me.objParas.paras_delete_Paragraphs(tbl.Range, 0)
        'rng = MyBase.Base_Paragraphs_Delete(tbl.Range, 0)
        rng = tbl.Rows.Item(3).Cells.Item(1).Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        tblNestReliance = rng.Tables.Add(rng, 3, 1)
        '
        conts_frontContacts_tableNestedReliance(tblNestReliance, doBottomPageImage)            'Nested Table with the Image
        '
        rng = tbl.Rows.Item(2).Cells.Item(1).Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        tblNestAboutAAC = rng.Tables.Add(rng, 2, 1)
        '
        conts_frontContacts_tableNestedAboutAAC(tblNestAboutAAC)
        '
        'doAccessibleVersion = True
        If doAccessibleVersion Then
            tblNestReliance.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs,)
            tblNestAboutAAC.ConvertToText(WdTableFieldSeparator.wdSeparateByParagraphs)
            '
            objParas.paras_delete_lastParasInTableCell(tbl.Range.Cells.Item(2))
            objParas.paras_delete_lastParasInTableCell(tbl.Range.Cells.Item(3))
            'Accessible version.. We place the contacts table after a spacing table. The
            'spacing table is left to ensure the contacts page tag is still there
            'rng = MyBase.Base_Paragraphs_Delete(tbl.Range, 0)
            'rng = sect.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            'rng.Move(WdUnits.wdParagraph, -1)
            'rng.Text = "hello"
            '
            'GoTo finis
            'para = rng.Paragraphs.Add(rng)
            'rng = sect.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            'para = rng.Paragraphs.Item(1)
            'rng = para.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            'Must adjust the spacing table height
            'tbl.Rows.Item(2).Height = 364.0
            'If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then tbl.Rows.Item(2).Height = 200.0
            '
        Else
            'Standard version.. We palce the contacts table as a nested table in the first
            'table
            '
            'Must adjust the containing/host table height
            'tbl.Rows.Item(4).Height = 626.6
            'If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then tbl.Rows.Item(4).Height = 356.0
            '
        End If
        '
        '
        '
finis:
        Return tblNestReliance
    End Function
    '
    '
    ''' <summary>
    ''' This method takes as input a table (tbl) and formats it according to the requirements
    ''' of the Front Contacts Page table 01... This is the either the host table for a standard
    ''' format document or the spacing table for an accessible document. The host height or the 
    ''' padding height is set by drHeight. The default is the height for a table 01 hosting a
    ''' nested table, that is 604.6
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function conts_frontContacts_tableContainer(ByRef tbl As Word.Table, verticalAlignmentIsBottom As Boolean) As Word.Table
        Dim objLegal As New cLegalAndAbout()
        Dim myDoc As Word.Document
        Dim sect As Word.Section
        Dim dr As Word.Row
        '
        myDoc = tbl.Range.Document
        sect = tbl.Range.Sections.Item(1)
        '
        tbl.Style = myDoc.Styles.Item("aac Table (no lines)")
        tbl.ApplyStyleHeadingRows = True

        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        tbl.AllowAutoFit = False
        tbl.Borders.Enable = False
        tbl.TopPadding = 0.0
        tbl.BottomPadding = 0.0
        tbl.LeftPadding = 5.4
        tbl.RightPadding = 5.4
        '
        'tagStyle row
        tbl.Rows.Item(1).HeightRule = WdRowHeightRule.wdRowHeightExactly
        tbl.Rows.Item(1).Height = 7.0
        tbl.Rows.Item(1).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item(Me.strTagStyleName_Front)
        '
        'About Acil Allen Row
        dr = tbl.Rows.Item(2)
        'dr.Range.Style = myDoc.Styles.Item("Cp About Acil Allen")
        dr.Range.Style = myDoc.Styles.Item("Normal - no space")
        dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
        dr.Cells.Item(1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop
        dr.Height = 365 - 230
        If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then dr.Height = 93
        'dr.Cells.Item(1).BottomPadding = 6.8
        'rng = dr.Cells.Item(1).Range
        'rng.Font.Color = RGB(255, 255, 255)
        'rng.Font.Bold = True
        'rng.Text = "About ACIL Allen"
        '
        dr = tbl.Rows.Item(3)
        dr.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
        dr.Range.Style = myDoc.Styles.Item("Normal - no space")
        dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
        'dr.Cells.Item(1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
        dr.Height = 365 + 230
        If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then dr.Height = 390



        'ACIL Allen Mission
        'dr = tbl.Rows.Item(3)
        'dr.Range.Style = myDoc.Styles.Item("Cp About Acil Allen")
        'dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
        'dr.Cells.Item(1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop
        'dr.Height = 84
        'dr.Cells.Item(1).BottomPadding = 16
        'rng = dr.Cells.Item(1).Range
        'objLegal.insert_aboutACIlAllen(rng, myDoc)
        '
        'Row that hosts the nested table
        'dr = tbl.Rows.Item(4)
        'dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
        '
        'dr.Height = 604.6
        'If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then dr.Height = 356
        '


        'tbl.Rows.Item(4).HeightRule = WdRowHeightRule.wdRowHeightExactly
        'tbl.Rows.Item(4).Height = 721.6
        'tbl.Rows.Item(4).Height = 604.6

        '
        If verticalAlignmentIsBottom Then
            'tbl.Rows.Item(4).Cells.Item(1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
        Else
            'tbl.Rows.Item(4).Cells.Item(1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop
        End If
        '
        '
        Return tbl
        '
    End Function
    '
    Public Function conts_frontContacts_tableNestedAboutAAC(ByRef tbl As Word.Table) As Word.Table
        Dim objLegal As New cLegalAndAbout()
        Dim myDoc As Word.Document
        Dim dr As Word.Row
        Dim rng As Word.Range
        '
        myDoc = tbl.Range.Document
        tbl.Style = myDoc.Styles.Item("aac Table (no lines)")
        tbl.ApplyStyleHeadingRows = True
        '
        '
        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        tbl.AllowAutoFit = False
        tbl.Borders.Enable = False
        tbl.TopPadding = 0.0
        tbl.BottomPadding = 0.0
        tbl.LeftPadding = 0.0
        tbl.RightPadding = 0.0
        '
        rng = tbl.Range
        rng.Style = myDoc.Styles.Item("Cp About Acil Allen")
        '
        dr = tbl.Rows.Item(1)
        rng = dr.Cells.Item(1).Range
        rng.Text = "About ACIL Allen"
        'rng.Font.Color = RGB(255, 255, 255)
        rng.Font.Color = objGlobals._glb_colour_purple_Dark

        rng.Font.Bold = True
        '
        dr = tbl.Rows.Item(2)
        rng = dr.Cells.Item(1).Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        objLegal.insert_aboutACIlAllen(rng, myDoc)
        '
        Return tbl
    End Function
    '
    ''' <summary>
    ''' This method takes as input a table (tbl) and formats it according to the requirements
    ''' of the Front Contacts Page table 02... This is the 'About Acil Allen' table. It is either
    ''' nested in the second cell of table01 for a standard format document or is placed 
    ''' underneath table01 whne the document is 'Accessible'
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function conts_frontContacts_tableNestedReliance(ByRef tbl As Word.Table, Optional doImage As Boolean = True) As Word.Table
        Dim objLegal As New cLegalAndAbout()
        Dim objFileMgr As New cFileHandler()
        Dim objRemResource As New cRemoteResources()
        Dim objRsrcsMgr As New cResourcesMgr()
        Dim myDoc As Word.Document
        Dim img_For_ContactsPageBottom As Image
        Dim iShp As InlineShape
        Dim strFileName As String
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim para As Word.Paragraph
        Dim rng As Word.Range
        Dim strLocalFilePath, strWebPath As String
        Dim checkFingerPrint As Boolean
        'Dim myFileInfo As FileInfo
        'Dim sect As Section
        'Dim j As Integer
        '
        strLocalFilePath = ""
        '
        myDoc = tbl.Range.Document
        tbl.Style = myDoc.Styles.Item("aac Table (no lines)")
        tbl.ApplyStyleHeadingRows = True
        '
        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        tbl.AllowAutoFit = False
        tbl.Borders.Enable = False
        tbl.TopPadding = 0.0
        tbl.BottomPadding = 0.0
        tbl.LeftPadding = 0.0
        tbl.RightPadding = 0.0
        checkFingerPrint = True
        '
        '
        'tbl.Rows.Item(1).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("Cp About Acil Allen")
        'tbl.Rows.Item(1).Cells.Item(1).BottomPadding = 6.8
        'rng = tbl.Rows.Item(1).Cells.Item(1).Range
        'rng.Font.Color = RGB(255, 254, 255)
        'rng.Font.Bold = True
        'rng.Text = "About ACIL Allen"

        '
        'tbl.Rows.Item(2).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("Cp About Acil Allen")
        'tbl.Rows.Item(2).Cells.Item(1).BottomPadding = 16.0
        'rng = tbl.Rows.Item(2).Cells.Item(1).Range
        '
        'objLegal.insert_aboutACIlAllen(rng, Globals.ThisAddin.Application.ActiveDocument)
        '
        'Acil Allen suggested citation row
        dr = tbl.Rows.Item(1)
        dr.Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("Cp Disclaimer 9pt")
        rng = dr.Cells.Item(1).Range
        objLegal.insert_SuggestedCitation(rng, Globals.ThisAddin.Application.ActiveDocument)
        '
        'Acil Allen Releiance and Disclaimer
        dr = tbl.Rows.Item(2)
        dr.Range.Style = myDoc.Styles.Item("Cp Disclaimer")
        rng = dr.Cells.Item(1).Range
        '
        objLegal.insert_disclaimer(rng, Globals.ThisAddin.Application.ActiveDocument)
        '
        '
        'Acil Allen Acknowledgment of country row
        dr = tbl.Rows.Item(3)
        '
        If Not doImage Then
            dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
            dr.Height = 3.2
            dr.Range.Style = myDoc.Styles.Item("spacer")
        Else
            'strImageName = "aac_pict_indigenous_00.png"
            strWebPath = "http://templates.acilallen.com.au/word/images/"
            strFileName = "artwork_contactsPage_front_release.png"
            checkFingerPrint = True
            'strLocalFilePath = objRemResource.remRsrc_get_fileFromWeb(strWebPath, strFileName, checkFingerPrint)
            strLocalFilePath = "dummy"
            '
            'strHash = objRemResource.cryptSHA_get_SHAasString(New IO.FileInfo(strLocalFilePath))
            '
            '
            If strLocalFilePath = "" Then
                'We didn't get the file from the web, so just do the normal thing
                dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
                dr.Height = 3.2
                dr.Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("spacer")
            Else
                'We did get the file so we will place the indigenous artwork here
                '
                dr.HeightRule = WdRowHeightRule.wdRowHeightAuto

                'tblNest.Rows.Item(5).Cells.Item(1).TopPadding = 50
                drCell = dr.Cells.Item(1)
                'drCell.TopPadding = 12
                'drCell.LeftPadding = 0
                'drCell.RightPadding = 0

                rng = drCell.Range
                rng.Style = myDoc.Styles.Item("Normal - no space")
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Paragraphs.Add()
                'rng.Paragraphs.Add()
                '
                para = rng.Paragraphs.Add()
                rng = para.Range
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                rng.Move(WdUnits.wdParagraph, -1)
                '
                rng.ParagraphFormat.SpaceBefore = 0
                rng.ParagraphFormat.SpaceAfter = 0
                rng.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                '
                '***
                'Two alternatives, one gets the file from the web, the other gets it from a local Resource
                '
                'Extract image from http://templates.acilallen.com.au/word/images/   Or  Let's get it from local Resources (much better)
                'iShp = objFileMgr.file_insert_imageFromFile(rng, drCell.Width - drCell.LeftPadding - drCell.RightPadding, strLocalFilePath)
                '
                img_For_ContactsPageBottom = objRsrcsMgr.rsrcs_get_contactsArtWork()
                Try
                    '*** Problem here iShp seems to come out as nothing, but the image is inserted
                    iShp = objFileMgr.file_insert_imageFromResources(rng, drCell, img_For_ContactsPageBottom)
                    If Not IsNothing(iShp) Then
                        iShp.AlternativeText = "A piece of indigenous artwork called Goomup, by Jarni McGuire"
                    End If
                Catch ex As Exception

                End Try
                '
                '***
                '
                'sect = rng.Sections.Item(1)
                'If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
                'iShp.LockAspectRatio = False
                'iShp.Height = 99.4
                'End If
                'iShp = objFileMgr.file_insert_imageFromWeb(rng, drCell.Width, strImageName, False)
                'iShp.AlternativeText = "Goomup, by Jarni McGuire"
                '
                rng = drCell.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                'para = rng.Paragraphs.Add()
                rng.Style = myDoc.Styles.Item("Cp About Acil Allen")
                'The following paragraph spacing separates the artwork from the acknowledgement text
                objLegal.legal_insert_ackOfCountry(rng)
                rng.ParagraphFormat.SpaceBefore = 12
                rng.ParagraphFormat.SpaceAfter = 12
                '
                rng = drCell.Range
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                rng.Move(WdUnits.wdParagraph, -1)
                rng.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
                rng = objLegal.legal_insert_ackForArtWork(rng, "Goomup, by Jarni McGuire")
                rng.Font.Size = 7
                'rng.Font.Color = RGB(255, 255, 255)
                rng.Font.Color = objGlobals._glb_colour_purple_Dark
                rng.ParagraphFormat.SpaceBefore = 3
                rng.ParagraphFormat.SpaceAfter = 0
                rng.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                '
                '
            End If

            '
        End If
        '
        'Clean out the image from local storage
        If strLocalFilePath <> "" Then
            'myFileInfo = New FileInfo(strLocalFilePath)
            'If myFileInfo.Exists Then
            'myFileInfo.Delete()
            'End If
        End If

        Return tbl
        '
    End Function
    '

    Public Function xx_insert_Contacts_Table_Front(ByRef sect As Word.Section) As Word.Table
        Dim objLegal As New cLegalAndAbout()
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objFileMgr As New cFileHandler()
        Dim rng As Word.Range
        Dim tbl, tblNest As Word.Table
        Dim drCellHost, drCell As Word.Cell
        Dim myDoc As Word.Document
        '
        rng = sect.Range
        myDoc = rng.Document
        'rng.Text = ""
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        tbl = rng.Tables.Add(rng, 2, 1)
        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        tbl.AllowAutoFit = False
        tbl.Borders.Enable = False
        tbl.TopPadding = 0.0
        tbl.BottomPadding = 0.0
        tbl.LeftPadding = 5.4
        tbl.RightPadding = 5.4
        '
        tbl.Rows.Item(1).HeightRule = WdRowHeightRule.wdRowHeightExactly
        tbl.Rows.Item(1).Height = 7.0
        tbl.Rows.Item(1).Range.Style = myDoc.Styles.Item(Me.strTagStyleName_Front)
        '
        tbl.Rows.Item(2).HeightRule = WdRowHeightRule.wdRowHeightExactly
        tbl.Rows.Item(2).Height = 721.6
        tbl.Rows.Item(2).Cells.Item(1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
        '
        rng = Me.objParas.paras_delete_Paragraphs(tbl.Range, 0)

        drCellHost = tbl.Rows.Item(2).Cells.Item(1)
        rng = drCellHost.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        tblNest = drCellHost.Tables.Add(rng, 5, 1)
        tblNest.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        tblNest.AllowAutoFit = False
        tblNest.Borders.Enable = False
        tblNest.TopPadding = 0.0
        tblNest.BottomPadding = 0.0
        tblNest.LeftPadding = 5.4
        tblNest.RightPadding = 5.4
        '
        '
        tblNest.Rows.Item(1).Range.Style = myDoc.Styles.Item("Cp About Acil Allen")
        tblNest.Rows.Item(1).Cells.Item(1).BottomPadding = 6.8
        rng = tblNest.Rows.Item(1).Cells.Item(1).Range
        rng.Font.Color = RGB(255, 254, 255)
        rng.Font.Bold = True
        rng.Text = "About ACIL Allen"

        '
        tblNest.Rows.Item(2).Range.Style = myDoc.Styles.Item("Cp About Acil Allen")
        tblNest.Rows.Item(2).Cells.Item(1).BottomPadding = 16.0
        rng = tblNest.Rows.Item(2).Cells.Item(1).Range
        '
        objLegal.insert_aboutACIlAllen(rng, myDoc)
        '
        tblNest.Rows.Item(3).Range.Style = myDoc.Styles.Item("Cp Disclaimer 9pt")
        rng = tblNest.Rows.Item(3).Cells.Item(1).Range
        objLegal.insert_SuggestedCitation(rng, myDoc)

        tblNest.Rows.Item(4).Range.Style = myDoc.Styles.Item("Cp Disclaimer")
        rng = tblNest.Rows.Item(4).Cells.Item(1).Range
        '
        objLegal.insert_disclaimer(rng, myDoc)
        '
        tblNest.Rows.Item(5).HeightRule = WdRowHeightRule.wdRowHeightExactly
        tblNest.Rows.Item(5).Height = 3.2
        tblNest.Rows.Item(5).Range.Style = myDoc.Styles.Item("spacer")
        '
        '**** test
        Dim para As Word.Paragraph
        '
        myDoc = sect.Range.Document
        tblNest.Rows.Item(5).HeightRule = WdRowHeightRule.wdRowHeightAuto

        'tblNest.Rows.Item(5).Cells.Item(1).TopPadding = 50
        drCell = tblNest.Rows.Item(5).Cells.Item(1)
        drCell.LeftPadding = 0
        drCell.RightPadding = 0

        rng = drCell.Range
        rng.Style = myDoc.Styles.Item("Normal")
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Paragraphs.Add()
        rng.Paragraphs.Add()
        '
        para = rng.Paragraphs.Add()
        rng = para.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
        objFileMgr.file_get_imageFromWeb(rng)

        '
        Return tbl
    End Function
    '
    Public Function insert_Contacts_Table_Back(ByRef sect As Word.Section) As Word.Table
        Dim objLegal As New cLegalAndAbout()
        Dim rng As Word.Range
        Dim tbl, tblNest As Word.Table
        Dim drCell As Word.Cell
        Dim drCellHost As Word.Cell
        Dim myDoc As Word.Document
        '
        rng = sect.Range
        myDoc = rng.Document
        'rng.Text = ""
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        tbl = rng.Tables.Add(rng, 2, 1)
        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        tbl.AllowAutoFit = False
        tbl.Borders.Enable = False
        tbl.TopPadding = 0.0
        tbl.BottomPadding = 0.0
        tbl.LeftPadding = 5.4
        tbl.RightPadding = 5.4
        '
        tbl.Rows.Item(1).HeightRule = WdRowHeightRule.wdRowHeightExactly
        tbl.Rows.Item(1).Height = 7.0
        tbl.Rows.Item(1).Range.Style = myDoc.Styles.Item(Me.strTagStyleName_Back)
        '
        'Adjust back page table depending upon orientation
        '
        If sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait Then
            tbl.Rows.Item(2).HeightRule = WdRowHeightRule.wdRowHeightExactly
            tbl.Rows.Item(2).Height = 721.6
            tbl.Rows.Item(2).Cells.Item(1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
        Else
            tbl.Rows.Item(2).HeightRule = WdRowHeightRule.wdRowHeightExactly
            tbl.Rows.Item(2).Height = 470
            tbl.Rows.Item(2).Cells.Item(1).VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalBottom
        End If
        '
        rng = Me.objParas.paras_delete_Paragraphs(tbl.Range, 0)
        '
        drCellHost = tbl.Rows.Item(2).Cells.Item(1)
        rng = drCellHost.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        tblNest = drCellHost.Tables.Add(rng, 4, 3)
        tblNest.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        tblNest.AllowAutoFit = False
        tblNest.Borders.Enable = False
        tblNest.TopPadding = 0.0
        tblNest.BottomPadding = 0.0
        tblNest.LeftPadding = 5.4
        tblNest.RightPadding = 5.4
        '
        '
        tblNest.Range.Style = myDoc.Styles.Item("Cp Contact Details")

        tblNest.Rows.Item(2).Cells.Item(1).BottomPadding = 23.4
        tblNest.Rows.Item(3).Cells.Item(1).BottomPadding = 18.8
        'tblNest.Rows.Item(4).Range.Font.Color = RGB(157, 133, 190)
        drCell = tblNest.Range.Cells.Item(10)
        objLegal.insert_Back_WebAddress(drCell.Range)
        '
        rng = tblNest.Range.Cells.Item(4).Range
        objLegal.insert_Back_MelbourneAndCanberra(rng)
        rng = tblNest.Range.Cells.Item(5).Range
        objLegal.insert_Back_SydneyAndPerth(rng)
        rng = tblNest.Range.Cells.Item(6).Range
        objLegal.insert_Back_BrisbaneAndAdelaide(rng)

        '
        rng = tblNest.Range.Cells.Item(7).Range
        objLegal.insert_Back_CompanyAndABN(rng)

        '
        'objLegal.insert_aboutACIlAllen(rng, Globals.ThisAddin.Application.ActiveDocument)
        '
        'tblNest.Rows.Item(3).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("Cp Disclaimer 9pt")
        'rng = tblNest.Rows.Item(3).Cells.Item(1).Range
        'objLegal.insert_SuggestedCitation(rng, Globals.ThisAddin.Application.ActiveDocument)

        'tblNest.Rows.Item(4).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("Cp Disclaimer")
        'rng = tblNest.Rows.Item(4).Cells.Item(1).Range
        '
        'objLegal.insert_disclaimer(rng, Globals.ThisAddin.Application.ActiveDocument)
        '
        'tblNest.Rows.Item(5).HeightRule = WdRowHeightRule.wdRowHeightExactly
        'tblNest.Rows.Item(5).Height = 3.2
        'tblNest.Rows.Item(5).Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles.Item("spacer")
        '
        Return tbl
    End Function

#End Region
    '
#Region "Determining if a Back Contacts Page"
    '
    '
    Public Function is_ContactsPage_Front(ByRef sect As Word.Section) As Boolean
        Dim rslt As Boolean
        Dim objsectMgr As New cSectionMgr()
        '
        rslt = False
        rslt = objsectMgr.sct_sectHas_strTag(sect, Me.strTagStyleName_Front)
        '
        Return rslt
    End Function
    '
    ''' <summary>
    ''' This method will return true if the current section is a Back Contacts Page
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Overridable Function is_ContactsPage_Back(ByRef sect As Word.Section) As Boolean
        Dim rslt As Boolean
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objBnrMgr As New cChptBanner()
        Dim strTag As String
        '
        rslt = False
        strTag = objHfMgr.hf_tags_getTagStyleName(sect, "primaryOrFirstPage")
        If strTag = objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_cont_Back) Then
            rslt = True
        End If

        Return rslt
    End Function

    '
    ''' <summary>
    ''' This method will return true if the document has a Back Contacts Page
    ''' </summary>
    ''' <returns></returns>
    Public Function has_ContactsPage_Back() As Boolean
        Dim sect As Word.Section
        Dim objsectMgr As New cSectionMgr()
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objBnrMgr As New cChptBanner()
        Dim strTag As String
        Dim rslt As Boolean
        '
        rslt = False
        '
        sect = objGlobals.glb_get_wrdActiveDoc.Sections.Last
        strTag = objHfMgr.hf_tags_getTagStyleName(sect)
        If strTag = objBnrMgr.bnr_get_tagStyles(objBnrMgr.tag_cont_Back) Then rslt = True
        '
        Return rslt
    End Function

    '
    ''' <summary>
    ''' This method will return true if the document has a Front Contacts Page
    ''' </summary>
    ''' <returns></returns>
    Public Function conts_has_ContactsPage_Front() As Boolean
        Dim rslt As Boolean
        Dim objsectMgr As New cSectionMgr()
        Dim sect As Word.Section

        '
        rslt = False
        sect = Nothing
        rslt = objsectMgr.sct_has_strTag(objGlobals.glb_get_wrdActiveDoc, Me.strTagStyleName_Front, sect)
        '
        Return rslt
    End Function
    '
    '
    ''' <summary>
    ''' This method will return nothing if the document does not have a Front Contacts Page... If it does it
    ''' will return the section
    ''' </summary>
    ''' <returns></returns>
    Public Function conts_getContactsPage_Front() As Word.Section
        Dim rslt As Boolean
        Dim objsectMgr As New cSectionMgr()
        Dim sect As Word.Section
        '
        sect = Nothing
        rslt = objsectMgr.sct_has_strTag(objGlobals.glb_get_wrdActiveDoc, Me.strTagStyleName_Front, sect)
        '
        Return sect
        '
    End Function




#End Region
    '


End Class

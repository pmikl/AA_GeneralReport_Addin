Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cPlHFloatingMgr
    Inherits cPlHBase
    '
    '
    Public Sub New()
        MyBase.New()
    End Sub
    '
    Public Sub PlHFloat_convert_toInLine(ByRef tbl As Word.Table)
        Me.objTblsMgr.tbl_convert_toInLine(tbl)
    End Sub
    '
    Public Sub PlHFloat_convert_toFloating(ByRef tbl As Word.Table)
        Me.objTblsMgr.tbl_convert_toFloating(tbl)
    End Sub
    '
    Public Function plHFloat_is_Floating(ByRef tbl As Word.Table) As Boolean
        Return Me.objTblsMgr.tbl_is_Floating(tbl)
    End Function
    '
    Public Sub plHFLoat_setWidth_toStandard(ByRef tbl As Word.Table)
        Dim sect As Word.Section
        '
        sect = tbl.Range.Sections.Item(1)
        '
        If Me.plHFloat_is_Floating(tbl) Then
            tbl.Rows.WrapAroundText = True

            tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
            tbl.Rows.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph
            'tbl.Rows.RelativeHorizontalPosition = 0.0
            tbl.Rows.HorizontalPosition = 0.0
            tbl.Rows.VerticalPosition = 0.0
            tbl.Rows.Alignment = WdRowAlignment.wdAlignRowLeft

            tbl.PreferredWidth = glb_get_widthBetweenMargins(sect)
        End If

    End Sub
    '
    '
    Public Sub plHFLoat_setWidth_toWide(ByRef tbl As Word.Table)
        Dim sect As Word.Section
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim hdr_tbl As Word.Table
        Dim tblWidth, leftIndent As Single
        '
        sect = tbl.Range.Sections.Item(1)
        hdr_tbl = Nothing
        hdr_tbl = objHfMgr.hf_get_HeaderTable(sect, "primaryOrFirstPage")
        tblWidth = objHfMgr.hf_get_HeaderTableWidth(hdr_tbl)
        '
        'tblWidth = glb_hfs_getHeaderTableWidth(sect, "primaryOrFirstPage")
        leftIndent = hdr_tbl.Rows.Item(1).LeftIndent
        '
        If tblWidth > 0.0 Then
            Try
                '
                If Me.plHFloat_is_Floating(tbl) Then
                    tbl.Rows.WrapAroundText = True

                    tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
                    tbl.Rows.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph
                    'tbl.Rows.RelativeHorizontalPosition = 0.0
                    tbl.Rows.HorizontalPosition = leftIndent
                    tbl.Rows.VerticalPosition = 0.0
                    tbl.Rows.Alignment = WdRowAlignment.wdAlignRowLeft

                    tbl.PreferredWidth = tblWidth
                End If

            Catch ex As Exception

            End Try

        Else

        End If

    End Sub
    '
    '
    ''' <summary>
    ''' This method will float the specified table ad lock it to the top margin
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub PlHFloat_lock_toMarginsTop(ByRef tbl As Word.Table)
        Dim rng As Word.Range
        Dim sect As Word.Section
        Dim marginWidth, leftIndent, selPosH As Single
        '
        sect = tbl.Range.Sections.Item(1)
        glb_get_widthBetweenMargins(sect)
        marginWidth = glb_get_widthBetweenMargins(sect)
        'tblWidth = tbl.Range.Cells.Item(1).Width
        'selPosH = 0.0
        '
        'Now find (in relative terms) the indent of the first row with respect to the rest of the Table
        'drHeader = tbl.Rows.First
        'drLast = tbl.Rows.Last
        'leftIndent = drHeader.LeftIndent - drLast.LeftIndent
        '
        selPosH = leftIndent

        '
        'Now float the table
        tbl.Rows.WrapAroundText = True
        tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
        tbl.Rows.HorizontalPosition = selPosH
        tbl.Rows.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin
        tbl.Rows.VerticalPosition = 0.0
        tbl.Rows.AllowOverlap = False
        rng = tbl.Range
        '
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will float the specified table
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub PlHFloat_lock_toParagraphAndColumn(ByRef tbl As Word.Table, Optional strAlign As String = "left")
        Dim selPosH As Single
        '
        selPosH = 0.0
        '
        'Now float the table
        tbl.Rows.WrapAroundText = True

        tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionColumn
        'tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
        'tbl.Rows.HorizontalPosition = 0.0
        tbl.Rows.HorizontalPosition = selPosH
        tbl.Rows.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph
        tbl.Rows.VerticalPosition = 0.0
        tbl.Rows.AllowOverlap = False
        '
        Select Case strAlign
            Case "left"
                tbl.Rows.Alignment = WdRowAlignment.wdAlignRowLeft
            Case "centre"
                tbl.Rows.Alignment = WdRowAlignment.wdAlignRowCenter
            Case "right"
                tbl.Rows.Alignment = WdRowAlignment.wdAlignRowRight
        End Select
        '
    End Sub

    '
    ''' <summary>
    ''' This method will float the specified table
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub PlHFloat_lock_toParagraphAndMarginLeft(ByRef tbl As Word.Table, Optional strAlign As String = "left")
        Dim selPosH As Single
        '
        selPosH = 0.0
        '
        'Now float the table
        tbl.Rows.WrapAroundText = True

        'tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionColumn
        tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
        'tbl.Rows.HorizontalPosition = 0.0
        tbl.Rows.HorizontalPosition = selPosH
        tbl.Rows.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph
        tbl.Rows.VerticalPosition = 0.0
        tbl.Rows.AllowOverlap = False
        '
        Select Case strAlign
            Case "left"
                tbl.Rows.Alignment = WdRowAlignment.wdAlignRowLeft
            Case "centre"
                tbl.Rows.Alignment = WdRowAlignment.wdAlignRowCenter
            Case "right"
                tbl.Rows.Alignment = WdRowAlignment.wdAlignRowRight
        End Select
        '
    End Sub
    '
    Public Sub xxx(ByRef tbl As Word.Table)
        Dim sect As Word.Section
        '
        sect = tbl.Range.Sections.Item(1)
        '
        tbl.Rows.WrapAroundText = True
        tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
        tbl.Rows.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph
        'tbl.Rows.RelativeHorizontalPosition = 0.0
        tbl.Rows.HorizontalPosition = 0.0
        tbl.Rows.VerticalPosition = 0.0
        tbl.Rows.Alignment = WdRowAlignment.wdAlignRowLeft

        tbl.PreferredWidth = glb_get_widthBetweenMargins(sect)

    End Sub
    '
    '
    Public Sub Plh_Float_LockInPosition_RelativeToLeftPageEdge(ByRef tbl As Word.Table, leftEdge As Single, tblWidth As Single)
        Dim rng, rngOld As Word.Range
        Dim selPosVert As Single
        Dim tblWorkAround As Word.Table
        '
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)

        rngOld = rng

        tbl.Rows.WrapAroundText = True
        tbl.Columns.Item(1).Width = glb_get_wrdSect.PageSetup.PageWidth
        '
        tbl.LeftPadding = glb_get_wrdSect.PageSetup.LeftMargin
        tbl.RightPadding = glb_get_wrdSect.PageSetup.RightMargin

        'selPosH = Globals.ThisDocument.Application.Selection.Range.Information(WdInformation.wdHorizontalPositionRelativeToPage)
        'selPosVert = Globals.ThisDocument.Application.Selection.Range.Information(WdInformation.wdVerticalPositionRelativeToPage)
        '
        'sect = tbl.Range.Sections.Item(1)
        '
        'Just make certain that we are at the first cell
        '
        'para = Globals.ThisDocument.Application.Selection.Paragraphs.Item(1)
        'rng = para.Range
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        'selPosH = rng.Information(WdInformation.wdHorizontalPositionRelativeToPage)
        'selPosVert = rng.Information(WdInformation.wdVerticalPositionRelativeToPage)
        '
        '*** This seems to cause race conditions
        selPosVert = CSng(glb_get_wrdSelRng.Information(WdInformation.wdVerticalPositionRelativeToPage))
        '
        '*** This does not... So it has something to do with how we get selPosVert
        'selPosVert = 400
        'selPosVert = Globals.ThisDocument.vPos
        '
        'Relative to the margins
        'selPosH = 0.0
        'selPosVert = selPosVert - sect.PageSetup.TopPage
        '
        'Now we are locking on the margin, so we just set selPosH to 0.0
        'selPosH = 0.0
        'topPadding = drCell.TopPadding
        'bottomPadding = drCell.BottomPadding
        '
        'Now find (in relative terms) the indent of the first row with respect to the rest of the Table
        'drHeader = tbl.Rows.First
        'drLast = tbl.Rows.Last
        'leftIndent = drHeader.LeftIndent - drLast.LeftIndent
        '
        'selPosH = leftIndent
        'If tbl.Rows.Item(1).LeftIndent < 0.0 Then selPosH = tbl.Rows.Item(1).LeftIndent
        '
        'Now float the table
        '
        tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
        tbl.Rows.HorizontalPosition = leftEdge
        tbl.Rows.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph
        '
        tbl.Rows.VerticalPosition = 0
        'tbl.Rows.VerticalPosition = selPosVert
        'tbl.Range.movew
        tbl.Rows.AllowOverlap = False
        rng = tbl.Range
        '
        tbl.Rows.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
        tbl.Rows.VerticalPosition = selPosVert
        '
        tbl.Rows.WrapAroundText = True
        tbl.Rows.WrapAroundText = True
        '
        '***
        '*** Seems that we have to perfom an immediate inline action to get rid of the cursor race
        '
        'rngOld.Move(WdUnits.wdParagraph, 1)
        rngOld.Collapse(WdCollapseDirection.wdCollapseStart)
        'rngOld.Text = "example"
        tblWorkAround = Me.Plh_Table_Build(rngOld, 2, 1, False, False, False)
        tblWorkAround.Delete()
        '
        '***

        '
    End Sub
    '

    '
    ''' <summary>
    ''' This method will float the specified table... It determines the horizontal and vertical positions
    ''' of the current selection (by checking the Information related to the range of the paragraph) and
    ''' floates it on th page, locking it in position
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub PlHFloat_lock_toMarginsLeftAndTop(ByRef tbl As Word.Table)
        Dim rng As Word.Range
        Dim selPosH, selPosVert, topPadding, bottomPadding, leftIndent As Single
        Dim drCell As Word.Cell
        Dim sect As Word.Section
        Dim myDoc As Word.Document
        Dim drLast, drHeader As Word.Row

        'selPosH = Globals.ThisDocument.Application.Selection.Range.Information(WdInformation.wdHorizontalPositionRelativeToPage)
        'selPosVert = Globals.ThisDocument.Application.Selection.Range.Information(WdInformation.wdVerticalPositionRelativeToPage)
        '
        sect = tbl.Range.Sections.Item(1)
        myDoc = sect.Range.Document
        '
        'Just make certain that we are at the first cell
        drCell = tbl.Range.Cells.Item(1)
        rng = drCell.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        'para = Globals.ThisDocument.Application.Selection.Paragraphs.Item(1)
        'rng = para.Range
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        selPosH = rng.Information(WdInformation.wdHorizontalPositionRelativeToPage)
        selPosVert = rng.Information(WdInformation.wdVerticalPositionRelativeToPage)
        '
        'Relative to the margins
        selPosH = selPosH - sect.PageSetup.LeftMargin
        selPosVert = selPosVert - sect.PageSetup.TopMargin
        '
        'Now we are locking on the margin, so we just set selPosH to 0.0
        selPosH = 0.0
        topPadding = drCell.TopPadding
        bottomPadding = drCell.BottomPadding
        '
        'Now find (in relative terms) the indent of the first row with respect to the rest of the Table
        drHeader = tbl.Rows.First
        drLast = tbl.Rows.Last
        leftIndent = drHeader.LeftIndent - drLast.LeftIndent
        '
        selPosH = leftIndent
        'If tbl.Rows.Item(1).LeftIndent < 0.0 Then selPosH = tbl.Rows.Item(1).LeftIndent
        '
        'Now float the table
        'Now float the table
        tbl.Rows.WrapAroundText = True
        tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
        tbl.Rows.HorizontalPosition = selPosH
        tbl.Rows.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin
        'tbl.Rows.VerticalPosition = selPosVert
        tbl.Rows.VerticalPosition = selPosVert - (topPadding + bottomPadding) - myDoc.Styles.Item("Caption").ParagraphFormat.SpaceBefore
        tbl.Rows.AllowOverlap = False
        rng = tbl.Range
        '
    End Sub
    '

    Public Sub Plh_Float_LockInPosition_RelativeToMargins_RegularTable(ByRef tbl As Word.Table, from_leftPageEdge As Single)
        Dim rng As Word.Range
        Dim selPosH, selPosVert As Single
        Dim drCell As Word.Cell
        Dim sect As Word.Section

        'selPosH = Globals.ThisDocument.Application.Selection.Range.Information(WdInformation.wdHorizontalPositionRelativeToPage)
        'selPosVert = Globals.ThisDocument.Application.Selection.Range.Information(WdInformation.wdVerticalPositionRelativeToPage)
        '
        sect = tbl.Range.Sections.Item(1)
        '
        'Just make certain that we are at the first cell
        drCell = tbl.Range.Cells.Item(1)
        rng = drCell.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        'para = Globals.ThisDocument.Application.Selection.Paragraphs.Item(1)
        'rng = para.Range
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        'selPosH = from_leftPageEdge
        selPosVert = rng.Information(WdInformation.wdVerticalPositionRelativeToPage)
        '
        'Relative to the margins
        selPosH = -(sect.PageSetup.LeftMargin - from_leftPageEdge)
        selPosVert = selPosVert - sect.PageSetup.TopMargin
        '
        'Now we are locking on the margin, so we just set selPosH to 0.0
        'selPosH = 0.0
        'topPadding = drCell.TopPadding
        'bottomPadding = drCell.BottomPadding
        '
        'Now find (in relative terms) the indent of the first row with respect to the rest of the Table
        'drHeader = tbl.Rows.First
        'drLast = tbl.Rows.Last
        'leftIndent = drHeader.LeftIndent - drLast.LeftIndent
        '
        'selPosH = leftIndent
        'If tbl.Rows.Item(1).LeftIndent < 0.0 Then selPosH = tbl.Rows.Item(1).LeftIndent
        '
        'Now float the table
        tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
        tbl.Rows.HorizontalPosition = selPosH
        tbl.Rows.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionMargin
        'tbl.Rows.VerticalPosition = selPosVert
        tbl.Rows.VerticalPosition = selPosVert
        tbl.Rows.AllowOverlap = False
        rng = tbl.Range
        '
    End Sub


    ''' <summary>
    ''' This method will take an existing regular table (not an AAC Table) and float it, locked
    ''' to a paragraph. The Table's left edge will be set relative to the left hand side of the
    ''' page (pts)
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="from_leftPageEdge"></param>
    Public Sub Plh_Float_LockToToParagraph_RegularTable(ByRef tbl As Word.Table, from_leftPageEdge As Single)
        Dim rng As Word.Range
        '
        'Now float the table
        'tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionColumn
        tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
        '
        tbl.Rows.HorizontalPosition = from_leftPageEdge
        tbl.Rows.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph
        tbl.Rows.VerticalPosition = 0.0
        tbl.Rows.AllowOverlap = False
        rng = tbl.Range
        '
    End Sub
    '

    '


    '
    ''' <summary>
    ''' This method will warn if the user is trying a Table Float action in a section
    ''' (e.g. Cover Page) where it is not alloled
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Overridable Function Plh_isOK_ToFloatPlaceHolders(ByRef sect As Section) As String
        Dim objRptMgr As New cReport()
        Dim objTagsMgr As New cTagsMgr()
        Dim strErrorMsg As String
        Dim rslt As Boolean
        '
        strErrorMsg = ""
        rslt = True
        '
        'If objRptMgr.Rpt_Mode_Get() <> objRptMgr.modeLongLandscape Then
        'strErrorMsg = "This Function can only be used in a Landscape Report"
        'GoTo finis
        'End If
        '
        'objTagsMgr.tags_get_tagStyleName()
        'objSectMgr.getSectionTag(sect)
        '

        If objTagsMgr.tags_is_CoverPage(sect) Then strErrorMsg = "This function is not supported in a Cover Page"
        If objTagsMgr.tags_is_TOCPage(sect) Then strErrorMsg = "This function is not supported in the Table of Contents"
        If objTagsMgr.tags_get_tagStyleName(sect) Like "tag_*" And Not objTagsMgr.tags_get_tagStyleName(sect) = "tag_chapterBanner" Then strErrorMsg = "Manipulating placeholders in this section is not supported"
        'If objChpt.is_Chapter(sect) Then strErrorMsg = "Changing the column structure of a chapter banner page is not supported"
        'If objChpt.is_Appendix_Chapter(sect) Then strErrorMsg = "Changing the column structure of an appendix chapter banner page is not supported"

        'If objContactsPageMgr.is_ContactsPage_Back(sect) Then strErrorMsg = "Changing the column structure of the back contacts is not supported"
        'If objContactsPageMgr.is_ContactsPage_Front(sect) Then strErrorMsg = "Changing the column structure of the front contacts page is not supported"
        '
        'Check to see if in table
        '
finis:
        Return strErrorMsg

    End Function
    '
#Region "Floating Table Widths"
    '
    ''' <summary>
    ''' This method will set the table based placeholder to fit edge to edge
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function Plh_Float_EdgToEdge(ByRef tbl As Word.Table) As String
        Dim sect As Word.Section
        Dim drShapeCell As Word.Cell
        Dim shpInline As Word.InlineShape
        Dim pageWidth As Single
        Dim objTblMgr As New cTablesMgr()
        'Dim objHfMgr As New cHeaderFooterMgr()
        Dim objISOK As New cIsOKToDo()
        'Dim tbl_Hdr As Word.Table
        Dim strMsg As String
        '
        sect = tbl.Range.Sections.Item(1)
        'tbl_Hdr = objHfMgr.hf_get_HeaderTable(sect)
        '
        pageWidth = sect.PageSetup.PageWidth
        'tbl.Columns.Item(1).Width = pageWidth
        'tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
        'tbl.Rows.HorizontalPosition = 0
        '
        strMsg = objISOK.isOKto_doAction_inReportBody()
        '
        If strMsg = objISOK._isOK Then
            strMsg = ""
            '
            tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
            tbl.Rows.HorizontalPosition = 0
            tbl.Rows.AllowOverlap = False
            '
            objTblMgr.tbl_width_Change(tbl, pageWidth)
            '
            drShapeCell = tbl.Range.Cells.Item(2)
            '
            If drShapeCell.Range.InlineShapes.Count <> 0 Then
                shpInline = drShapeCell.Range.InlineShapes.Item(1)
                shpInline.LockAspectRatio = True
                shpInline.Width = pageWidth
                shpInline.LockAspectRatio = False
            End If

        End If
        '
        Return strMsg
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will set the table based placeholder to fit edge to edge
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function Plh_Float_Wide(ByRef tbl As Word.Table) As String
        Dim sect As Word.Section
        Dim drShapeCell As Word.Cell
        Dim shpInline As Word.InlineShape
        Dim tblWidth As Single
        Dim objTblMgr As New cTablesMgr()
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objISOK As New cIsOKToDo()
        Dim tbl_Hdr As Word.Table
        Dim leftIndent As Single
        Dim strMsg As String
        '
        sect = tbl.Range.Sections.Item(1)
        tbl_Hdr = objHfMgr.hf_get_HeaderTable(sect)
        leftIndent = tbl_Hdr.Rows.LeftIndent
        '
        tblWidth = objHfMgr.hf_get_HeaderTableWidth(tbl_Hdr)
        'tbl.Columns.Item(1).Width = pageWidth
        'tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
        'tbl.Rows.HorizontalPosition = 0
        '
        strMsg = objISOK.isOKto_doAction_inReportBody()
        '
        If strMsg = objISOK._isOK Then
            strMsg = ""
            '
            tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
            tbl.Rows.HorizontalPosition = sect.PageSetup.LeftMargin + leftIndent                                    'Remember leftIndent is negative
            tbl.Rows.AllowOverlap = False
            '
            objTblMgr.tbl_width_Change(tbl, tblWidth)
            '
            drShapeCell = tbl.Range.Cells.Item(2)
            '
            If drShapeCell.Range.InlineShapes.Count <> 0 Then
                shpInline = drShapeCell.Range.InlineShapes.Item(1)
                shpInline.LockAspectRatio = True
                shpInline.Width = tblWidth
                shpInline.LockAspectRatio = False
            End If

        End If
        '
        Return strMsg
        '
    End Function


    Public Function Plh_Float_MarginToMargin(ByRef tbl As Word.Table) As String
        Dim sect As Word.Section
        Dim drShapeCell As Word.Cell
        Dim textColumnIndex As Integer
        Dim shpInline As Word.InlineShape
        Dim marginWidth As Single
        Dim objTblMgr As New cTablesMgr()
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objisOKtoDo As New cIsOKToDo()
        Dim strMsg As String
        '
        sect = glb_get_wrdSect()
        'sect = Globals.ThisDocument.Application.Selection.Range.Sections.Item(1)
        'marginWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
        marginWidth = glb_get_widthBetweenMargins(sect)
        '
        strMsg = objisOKtoDo.isOKto_doAction_inReportBody()
        '
        If strMsg = objisOKtoDo._isOK Then
            '
            strMsg = ""
            textColumnIndex = Me.Plh_Columnsx2_FindColumnNumber(sect)
            '
            If textColumnIndex = 1 Then
                objTblMgr.tbl_width_Change(tbl, marginWidth)
                'objTblMgr.tbl_set_ToWide(tbl)

                'tbl.Columns.Item(1).Width = marginWidth
                drShapeCell = tbl.Range.Cells.Item(2)
                '
                If drShapeCell.Range.InlineShapes.Count <> 0 Then
                    shpInline = drShapeCell.Range.InlineShapes.Item(1)
                    shpInline.LockAspectRatio = True
                    shpInline.Width = tbl.Columns.Item(1).Width
                    shpInline.LockAspectRatio = False
                End If
            Else
                strMsg = "PlaceHolders can only be made 'full width' if they are in column 1"
            End If
        Else
            strMsg = ""
        End If
        '
        Return strMsg
        '
    End Function
    '
    Public Function Plh_Float_FullColumn_x2(ByRef tbl As Word.Table) As String
        Dim sect As Word.Section
        Dim drCell, drShapeCell As Word.Cell
        Dim rng As Word.Range
        Dim cellStyle As Word.Style
        Dim shpInline As Word.InlineShape
        Dim marginWidth, tblWidth As Single
        Dim textColumnIndex, numTextColumns As Integer
        Dim objTblMgr As New cTablesMgr()
        Dim strMsg As String
        '
        sect = glb_get_wrdSect()
        marginWidth = glb_get_widthBetweenMargins(sect)
        strMsg = ""
        '
        Try
            textColumnIndex = Me.Plh_Columnsx2_FindColumnNumber(sect)
            numTextColumns = sect.PageSetup.TextColumns.Count
            'MsgBox("Column number = " + CStr(textColumnNumber))
            'Exit Sub
            '
            drCell = tbl.Range.Cells.Item(1)
            rng = drCell.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            cellStyle = rng.Style
            '
            'Will only allow this for the following types
            '
            'The width of the PlaceHodler is dependent on its position (i.e. which column it is in) and how many columns in the section
            '
            Select Case numTextColumns
                Case 1
                    tblWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
                    strMsg = "This is a one column section. Two column spanning is not available"

                Case 2
                    tblWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
                    Select Case textColumnIndex
                        Case 1
                            tblWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
                        Case 2
                            tblWidth = sect.PageSetup.TextColumns.Item(2).Width
                            strMsg = "A PlaceHolder in column 2 cannot be expanded to span two columns"
                    End Select

                Case 3
                    Select Case textColumnIndex
                        Case 1
                            tblWidth = sect.PageSetup.TextColumns.Item(1).Width + sect.PageSetup.TextColumns.Spacing + sect.PageSetup.TextColumns.Item(2).Width
                        Case 2
                            tblWidth = sect.PageSetup.TextColumns.Item(2).Width + sect.PageSetup.TextColumns.Spacing + sect.PageSetup.TextColumns.Item(3).Width
                        Case 3
                            tblWidth = sect.PageSetup.TextColumns.Item(3).Width
                            strMsg = "A PlaceHolder in column 3 cannot be expanded to span two columns"

                    End Select

                Case 4
                    Select Case textColumnIndex
                        Case 1
                            tblWidth = sect.PageSetup.TextColumns.Item(1).Width + sect.PageSetup.TextColumns.Spacing + sect.PageSetup.TextColumns.Item(2).Width
                        Case 2
                            tblWidth = sect.PageSetup.TextColumns.Item(2).Width + sect.PageSetup.TextColumns.Spacing + sect.PageSetup.TextColumns.Item(3).Width
                        Case 3
                            tblWidth = sect.PageSetup.TextColumns.Item(3).Width
                            strMsg = "A PlaceHolder in column 3 cannot be expanded to span two columns"
                        Case 4
                            tblWidth = sect.PageSetup.TextColumns.Item(3).Width
                            strMsg = "A PlaceHolder in column 4 cannot be expanded to span two columns"

                    End Select
            End Select
            '
            If strMsg = "" Then
                objTblMgr.tbl_width_Change(tbl, tblWidth)
                '
                'To cater for figure panels
                drShapeCell = tbl.Range.Cells.Item(2)
                '
                If drShapeCell.Range.InlineShapes.Count <> 0 Then
                    shpInline = drShapeCell.Range.InlineShapes.Item(1)
                    shpInline.LockAspectRatio = True
                    shpInline.Width = tblWidth
                    shpInline.LockAspectRatio = False
                End If

            End If
            '
        Catch ex As Exception
            strMsg = "Is the current selection in a Box, Table or Figure?"
        End Try
        '
finis:
        Return strMsg
        '
    End Function
    '
    Public Function Plh_Float_FullColumn(ByRef tbl As Word.Table) As String
        Dim sect As Word.Section
        Dim drShapeCell As Word.Cell
        Dim textColumnIndex As Integer
        Dim shpInline As Word.InlineShape
        Dim marginWidth As Single
        Dim objTblMgr As New cTablesMgr()
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim objisOKtoDo As New cIsOKToDo()
        Dim strMsg As String
        '
        sect = glb_get_wrdSect()
        'sect = Globals.ThisDocument.Application.Selection.Range.Sections.Item(1)
        'marginWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
        marginWidth = glb_get_widthBetweenMargins(sect)
        '
        strMsg = objisOKtoDo.isOKto_doAction_inReportBody()
        '
        If strMsg = objisOKtoDo._isOK Then
            '
            strMsg = ""
            textColumnIndex = Me.Plh_Columnsx2_FindColumnNumber(sect)
            '
            Try
                objTblMgr.tbl_width_Change(tbl, sect.PageSetup.TextColumns.Item(textColumnIndex).Width)
            Catch ex As Exception
                objTblMgr.tbl_width_Change(tbl, sect.PageSetup.TextColumns.Item(1).Width)
            End Try

            drShapeCell = tbl.Range.Cells.Item(2)
            '
            If drShapeCell.Range.InlineShapes.Count <> 0 Then
                shpInline = drShapeCell.Range.InlineShapes.Item(1)
                shpInline.LockAspectRatio = True
                shpInline.Width = tbl.Columns.Item(1).Width
                shpInline.LockAspectRatio = False
            End If
        Else
            strMsg = ""
        End If
        '
        Return strMsg
        '
    End Function

    '
    Public Sub Plh_Float_FullColumnx(ByRef tbl As Word.Table)
        Dim sect As Word.Section
        Dim drCell, drShapeCell As Word.Cell
        Dim dr As Word.Row
        Dim rng As Word.Range
        Dim cellStyle As Word.Style
        Dim shp As Word.Shape
        Dim shpInline As Word.InlineShape
        Dim marginWidth, delta_Width, leftIndent As Single
        Dim textColumnNumber As Integer
        Dim objTblMgr As New cTablesMgr()
        '

        sect = objGlobals.glb_get_wrdApp.Selection.Range.Sections.Item(1)
        marginWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
        '
        Try
            textColumnNumber = Me.Plh_Columnsx2_FindColumnNumber(sect)
            'MsgBox("Column number = " + CStr(textColumnNumber))
            'Exit Sub
            '
            drCell = tbl.Range.Cells.Item(1)
            rng = drCell.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            cellStyle = rng.Style
            '
            'Will only allow this for the following types
            '
            Select Case cellStyle.NameLocal
                Case "tag_execBanner", "tag_chapterBanner", "tag_appendixChapter"
                    'Banners are two column (Portrait Report) or one column (Landscape Report)
                    Select Case tbl.Columns.Count
                        Case 1
                            tbl.Columns.Item(1).Width = sect.PageSetup.TextColumns.Item(textColumnNumber).Width
                            drShapeCell = tbl.Range.Cells.Item(1)
                            '
                            If drShapeCell.Range.ShapeRange.Count <> 0 Then
                                shp = drShapeCell.Range.ShapeRange.Item(1)
                                shp.Width = tbl.Columns.Item(1).Width
                            End If
                            '
                        Case 2
                            'Two columns
                            delta_Width = sect.PageSetup.TextColumns.Item(textColumnNumber).Width - tbl.Columns.Item(1).Width - tbl.Columns.Item(2).Width
                            tbl.Columns.Item(1).Width = tbl.Columns.Item(1).Width + delta_Width
                            drShapeCell = tbl.Range.Cells.Item(1)
                            '
                            If drShapeCell.Range.ShapeRange.Count <> 0 Then
                                shp = drShapeCell.Range.ShapeRange.Item(1)
                                shp.Width = tbl.Columns.Item(1).Width + tbl.Columns.Item(2).Width
                            End If
                            '
                    End Select

                Case "Caption", "Caption Label"
                    'Placeholders are one column, so we'll just widen that column
                    '
                    tbl.Columns.Item(1).Width = sect.PageSetup.TextColumns.Item(textColumnNumber).Width
                    drShapeCell = tbl.Range.Cells.Item(2)
                    '
                    If drShapeCell.Range.InlineShapes.Count <> 0 Then
                        shpInline = drShapeCell.Range.InlineShapes.Item(1)
                        shpInline.Width = tbl.Columns.Item(1).Width
                    End If
                    '
                Case "Table column headings"
                    objTblMgr.tbl_width_Change(tbl, sect.PageSetup.TextColumns.Item(textColumnNumber).Width)
                    dr = tbl.Rows.Last
                    If dr.LeftIndent < 0.0 Then
                        leftIndent = dr.LeftIndent
                        For Each dr In tbl.Rows
                            dr.LeftIndent = dr.LeftIndent - leftIndent
                        Next
                    End If
                    '
            End Select
        Catch ex As Exception
            MsgBox("Is the current selection in a Box, Table or Figure?")
        End Try
        '

    End Sub

#End Region
    '


End Class

Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core

Public Class cPlHBase
    Inherits cTagsMgr
    '
    Public lstOfPlhTypes As List(Of String)                 'Must be set in each derived class
    Public objTblsMgr As New cTablesMgr()
    Public objGlobals As New cGlobals()
    '
    Public colourHeader As Long             'Table Header Colour
    'Public colourUnits As Long              'Units row colour
    'Public colourUnits_2 As Long              'Units row colour


    Public Sub New()
        MyBase.New()
        Me.lstOfPlhTypes = New List(Of String)
        Me.objTblsMgr = New cTablesMgr()
        '
        Me.colourHeader = objGlobals._glb_colour_purple_Dark
        '
    End Sub
    '
    ''' <summary>
    ''' This method will detect and scale any image/shape panels in cellIndex 2 of the table.
    ''' WHich means it will scale the image panel in the Figure Placeholder to the Preferred width of tbl. If figWIdth > 0.0
    ''' then the shape width will be set to figWidth (the aspect ratio is maintained)
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="figWidth"></param>
    Public Sub Plh_scale_FigureImageShape(ByRef tbl As Word.Table, Optional figWidth As Single = 0.0, Optional cellIndex As Integer = 2)
        Dim drShapeCell As Word.Cell
        Dim shpInline As Word.InlineShape
        '
        drShapeCell = tbl.Range.Cells.Item(cellIndex)
        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        '
        If drShapeCell.Range.InlineShapes.Count <> 0 Then
            shpInline = drShapeCell.Range.InlineShapes.Item(1)
            shpInline.LockAspectRatio = True
            '
            If figWidth > 0.0 Then
                shpInline.Width = figWidth
            Else
                shpInline.Width = tbl.PreferredWidth
            End If
            shpInline.LockAspectRatio = False
            '
        End If

    End Sub
    '
    Public Overridable Function Plh_insert_PlaceHolder_WithTest(ByRef rng As Word.Range, strType As String) As Word.Table
        Dim tbl As Word.Table
        Dim numTextColumns As Integer
        Dim strResult As String
        Dim isOKMgr As New cIsOKToDo()
        '
        tbl = Nothing
        '
        numTextColumns = rng.Sections.Item(1).PageSetup.TextColumns.Count
        '
        Try
            strResult = isOKMgr.isOKto_doAction_inReportBody(rng)
            '
            If strResult = isOKMgr._isOK Then
                tbl = Me.Plh_insert_PlaceHolderBasic(rng, strType)
            Else
                MsgBox("This function is only allowed in the body of the report or appendices")
            End If

        Catch ex As Exception

        End Try
        'End If
        '
        Return tbl
    End Function

    '
    ''' <summary>
    ''' This method will insert a PlaceHolder at the current selection with no test. That is
    ''' the user will not be warned if the selection point willr result is a 'troubled' insertion
    ''' </summary>
    ''' <param name="strType"></param>
    ''' <returns></returns>
    Public Overridable Function Plh_insert_PlaceHolder(strType As String) As Word.Table
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim numTextColumns As Integer
        Dim sect As Word.Section
        '
        tbl = Nothing
        rng = glb_get_wrdSelRng()
        sect = rng.Sections.Item(1)
        '
        numTextColumns = sect.PageSetup.TextColumns.Count
        '
        'If Me.Plh_IsOK_ToInsert(sect) Then
        Try
            tbl = Me.Plh_insert_PlaceHolderBasic(rng, strType)
            '
        Catch ex As Exception

        End Try
        'End If
        '
        '
        Return tbl
    End Function
    '
    ''' <summary>
    ''' This method will build a standard Table
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="numRows"></param>
    ''' <param name="numColumns"></param>
    ''' <param name="doBorders"></param>
    ''' <param name="doWideTable"></param>
    ''' <returns></returns>
    Public Function Plh_Table_Build(ByRef rng As Word.Range, numRows As Integer, numColumns As Integer, Optional doBorders As Boolean = True, Optional doWideTable As Boolean = False, Optional doAACTable As Boolean = False) As Word.Table
        Dim tbl As Word.Table
        Dim dr As Word.Row
        Dim drCol As Word.Column
        Dim sect As Word.Section
        Dim leftIndent As Single
        Dim rng2 As Word.Range

        '
        sect = rng.Sections.Item(1)
        '
        tbl = rng.Tables.Add(rng, numRows, numColumns)
        tbl.Style = rng.Document.Styles.Item("aac Table (no lines)")
        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        Me.objTblsMgr.tbl_fix_Table(tbl, doBorders, Me._glb_colour_TableBorders, True)

        'Me.objTblsMgr.fix_Table(tbl, doBorders, Me.colour_TableBorders)
        dr = tbl.Rows.Last
        'dr.Range.Cells.Item(1).Range.Text = "last"
        '
        If doAACTable Then
            tbl.Range.Style = objGlobals.glb_get_wrdStyle(objGlobals.glb_var_style_tblTextStyle)
            dr = tbl.Rows.Add(tbl.Rows.First)
            'Me.Plh_Table_doHeaderRow(dr)
            '
            'To Add rows to the end
            dr = objTblsMgr.tbl_rows_AddToEndOfTable(tbl)
            'rng = tbl.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            'dr = tbl.Rows.Add(rng)
            'dr.Range.Cells.Item(1).Range.Text = "Source"
            dr.Range.Style = objGlobals.glb_get_wrdStyle(objGlobals.glb_var_style_tblSourceStyle)
            dr.Range.Cells.Merge()
            dr.Range.Cells.Item(1).BottomPadding = 2.0
            rng2 = dr.Range
            rng2.Collapse(WdCollapseDirection.wdCollapseStart)
            Me.Plh_Insert_SourceAndNoteText(rng2)

            '
            dr = objTblsMgr.tbl_rows_AddToEndOfTable(tbl)
            'rng = tbl.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            'dr = tbl.Rows.Add(rng)
            'dr.Range.Cells.Item(1).Range.Text = "Spacer"
            dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
            dr.Height = Me.objGlobals.var_glb_tbl_bottomSpacerRowHeight
            dr.Range.Style = objGlobals.glb_get_wrdStyle("spacer")
            dr.Range.Cells.Merge()

            '

            'rng.Text = "hello"
            GoTo finis

            dr = tbl.Rows.Add(tbl.Rows.Last)
            dr = tbl.Rows.Add(tbl.Rows.Last)
            '
            dr = tbl.Rows.Last
            dr.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone
            dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
            dr.Height = Me.objGlobals.var_glb_tbl_bottomSpacerRowHeight
            dr.Range.Style = objGlobals.glb_get_wrdStyle("spacer")
            dr.Range.Cells.Merge()
            '
            dr = tbl.Rows.Item(tbl.Rows.Last.Index - 1)
            dr.Range.Style = objGlobals.glb_get_wrdStyle(objGlobals.glb_var_style_tblSourceStyle)
            '
            'dr.Borders.Item(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleSingle
            'dr.Borders.Item(WdBorderType.wdBorderHorizontal).LineWidth = WdLineWidth.wdLineWidth050pt
            'dr.Borders.Item(WdBorderType.wdBorderHorizontal).Color = Me.colour_TableBorders

            '
            dr.Range.Cells.Merge()
            dr.Range.Cells.Item(1).BottomPadding = 2.0
            rng2 = dr.Range
            rng2.Collapse(WdCollapseDirection.wdCollapseStart)
            Me.Plh_Insert_SourceAndNoteText(rng2)


        End If
        '
        If doWideTable And sect.PageSetup.TextColumns.Count = 1 Then
            leftIndent = sect.PageSetup.LeftMargin - Me.objGlobals.glb_hfs_getHFTableEdge(sect, "header_leftEdge")
            For Each dr In tbl.Rows
                dr.LeftIndent = -leftIndent + objGlobals.glb_get_TableOutdent()
            Next
            For Each drCol In tbl.Columns
                drCol.Width = drCol.Width + leftIndent / tbl.Columns.Count
            Next
        End If

finis:
        Return tbl
        '
    End Function
    '
    Public Function Plh_Table_Convert_ToAATable(ByRef tbl As Word.Table) As Word.Table
        Dim dr As Word.Row
        '
        'Me.objTblsMgr
        'Me.objTblsMgr.fix_Table(tbl, True, Me.colour_TableBorders)
        tbl.Range.Style = objGlobals.glb_get_wrdStyle(objGlobals.glb_var_style_tblTextStyle)
        'tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed)
        '
        dr = tbl.Rows.Add(tbl.Rows.Item(1))
        'Me.Plh_Table_doHeaderRow(dr)
        dr = tbl.Rows.Item(dr.Index + 1)
        dr.Borders.Item(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleNone


        'tbl = rng.Tables.Add(rng, numRows, numColumns)
        'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        '
        '
        Return tbl
        '
    End Function
    '

    '

    '
    '
    '
    Public Function Plh_insert_PlaceHolderBasic(ByRef rng As Word.Range, strType As String) As Word.Table
        Dim objBBMgr As New cBBlocksHandler()
        Dim objGrph As New cGraphicsMgr()
        Dim objCaptMgr As New cCaptionManager()
        Dim objFloatMgr As New cPlHFloatingMgr()
        Dim objHfMgr As New cHeaderFooterMgr()
        Dim lstOfEdges As New Collection()
        'Dim shp As Word.InlineShape
        'Dim shpTest As Word.Shape
        Dim tbl As Word.Table
        Dim sect As Word.Section
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim rng2 As Word.Range
        Dim tblWidth As Single
        Dim borderColour As Long
        Dim para As Word.Paragraph
        Dim myStyle As Word.Style
        '
        tbl = Nothing
        borderColour = RGB(0, 1, 0)
        sect = rng.Sections.Item(1)
        '
        'Find the text column and if in error default to column 1
        'currentTextColumn = Me.Plh_Columnsx2_FindColumnNumber(sect)
        'If currentTextColumn <= 0 Then currentTextColumn = 1
        'tbl = Me.Plh_Table_Build(rng, 3, 1)
        '
        '
        'If Me.Plh_IsOK_ToInsert(sect) Then
        'tbl = rng.Tables.Add(rng, 3, 1)
        'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        'tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed)

        'tblWidth = sect.PageSetup.TextColumns.Item(currentTextColumn).Width
        '
        'tbl.Columns.Item(1).Width = tblWidth
        '
        'Me.Plh_Table_FixPadding(tbl, 0.0, 0.0, 0.0, 0.0)
        'drCell = tbl.Range.Cells.Item(1)
        'Me.Base_Table_FixCellPadding(drCell, 8.0, 0.4, 7.2, 0.0)
        ' Me.Plh_Table_FixCellPadding(drCell, 8.0, 0.4, 0.0, 0.0)

        'dr = tbl.Range.Rows.Item(1)
        'dr.Range.Style = Me.Plh_Style_GetSpecificStyle("Caption")
        '
        '
        '
        '
        Select Case strType
            Case "Box_ES", "Box", "Box_AP", "Box_LT", "Recommendation_LT"
                tbl = Me.Plh_Table_Build(rng, 3, 1, False, False, False)
                '
                tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                tbl.PreferredWidth = 100
                '

                drCell = tbl.Range.Cells.Item(1)
                'Me.Base_Table_FixCellPadding(drCell, 8.0, 0.4, 7.2, 0.0)
                'Me.Plh_Table_FixCellPadding(drCell, 8.0, 0.4, 4.0, 0.0)
                Me.Plh_Table_FixCellPadding(drCell, 0.0, 0.4, 4.0, 0.0)

                dr = tbl.Range.Rows.Item(1)
                'objGlobals.tblCaptionStyle()
                dr.Range.Style = objGlobals.glb_get_wrdStyle(objGlobals.glb_var_style_tblCaptionStyle)


                drCell = tbl.Range.Cells.Item(2)
                Me.Plh_Table_FixCellPadding(drCell, 5.6, 9.2, 4.0, 6.0)
                drCell.Range.Text = "Overtype here"
                dr = tbl.Range.Rows.Item(2)
                '
                dr.Range.Style = objGlobals.glb_get_wrdStyle("Box Text")
                dr.Shading.BackgroundPatternColor = RGB(229, 229, 229)
                dr.Shading.Texture = Word.WdTextureIndex.wdTextureNone
                '
                'Now do the caption according to the strType
                rng = tbl.Range.Cells.Item(1).Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng = Me.Plh_Captions_InsertCaptions(strType, rng, True)
                rng.Select()
                '
                '
                'Now do the Source Cell
                drCell = tbl.Range.Cells.Item(tbl.Range.Cells.Count)
                Me.Plh_Table_FixCellPadding(drCell, 0.6, 5.65, 0.0, 0.0)
                dr = tbl.Range.Rows.Item(tbl.Range.Rows.Last.Index)
                dr.Range.Style = objGlobals.glb_get_wrdStyle("Source")
                rng2 = dr.Range
                rng2.Collapse(WdCollapseDirection.wdCollapseStart)
                Me.Plh_Insert_SourceAndNoteText(rng2)
                '
            Case "Key_Finding", "Key_Finding_ES", "Recommendation", "Recommendation_ES", "CaseStudy_HalfPage"
                'Default condition is inline
                tbl = Me.Plh_Table_Build(rng, 2, 1, False, False, False)
                tbl.Range.Cells.Item(2).BottomPadding = 6.0
                '
                'Now let's finish off the boxes
                Select Case strType
                    Case "Key_Finding", "Key_Finding_ES", "Finding", "Finding_ES"
                        tbl.Shading.ForegroundPatternColor = objGlobals._glb_colour_Finding_Purple
                        'Now do the caption according to the strType
                        rng = tbl.Range.Cells.Item(1).Range
                        rng.Collapse(WdCollapseDirection.wdCollapseStart)
                        rng = Me.Plh_Captions_InsertCaptions(strType, rng, True)
                        rng.Select()

                    Case "Recommendation", "Recommendation_ES"
                        tbl.Shading.ForegroundPatternColor = objGlobals._glb_colour_Recommendation_Purple
                        'Now do the caption according to the strType
                        rng = tbl.Range.Cells.Item(2).Range
                        rng.Text = "Insert recommendation text here"
                        '
                        rng = tbl.Range.Cells.Item(1).Range
                        rng.Collapse(WdCollapseDirection.wdCollapseStart)
                        rng = Me.Plh_Captions_InsertCaptions(strType, rng, True)
                        rng.Select()

                    Case "CaseStudy_HalfPage"
                        tbl.Shading.ForegroundPatternColor = objGlobals._glb_colour_CaseStudy_Grey
                        'Now do the caption according to the strType
                        'myStyle = tbl.Range.Document.Styles.Item("Heading (CaseStudy)")

                        rng = tbl.Range.Cells.Item(1).Range
                        rng.Collapse(WdCollapseDirection.wdCollapseStart)
                        rng = Me.Plh_Captions_InsertCaptions("CaseStudy_HalfPage", rng, True)
                        '
                        'rng.Font.Color = WdColor.wdColorRed
                        'rng = tbl.Range.Cells.Item(1).Range
                        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
                        'rng.Text = "Case Study (Partial Page)"
                        rng.Select()

                End Select
                '
                'For Portrait or Landscape the full page width percentage calculation is the same. Note that
                'The Percentage specs need to be at the end. If we place it at the beginning, the specs for left
                'and right padding (in pts) forces the table back actual measurements for width
                'If sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait And sect.PageSetup.TextColumns.Count = 1 Then
                '
                tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                'tbl.PreferredWidth = 100
                tbl.PreferredWidth = Me.Plh_get_FindingsEtc_Width_As_Percent(tbl)

                If sect.PageSetup.TextColumns.Count = 1 Then
                    '
                    'tbl.Columns.Item(1).Width = glb_get_wrdSect.PageSetup.PageWidth
                    'tbl.LeftPadding = glb_get_wrdSect.PageSetup.LeftMargin
                    'tbl.Rows.LeftIndent = -glb_get_wrdSect.PageSetup.LeftMargin
                    'tbl.RightPadding = glb_get_wrdSect.PageSetup.RightMargin
                    'tbl.Range.Cells.Item(2).Range.Style = glb_get_wrdActiveDoc.Styles.Item("Body Text")
                    'tbl.Range.Cells.Item(2).BottomPadding = 6.0
                    '
                    '***
                    '
                    'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                    'tbl.PreferredWidth = 100
                    'tbl.PreferredWidth = 100 * glb_get_wrdSect.PageSetup.PageWidth / glb_get_widthBetweenMargins(glb_get_wrdSect)
                    '
                    '***

                End If
                '

            Case "1", "2"
                If sect.PageSetup.TextColumns.Count = 1 Then
                    tbl = Me.Plh_Table_Build(rng, 2, 1, False, False, False)
                    tbl.Columns.Item(1).Width = glb_get_wrdSect.PageSetup.PageWidth
                    tbl.LeftPadding = glb_get_wrdSect.PageSetup.LeftMargin
                    tbl.RightPadding = glb_get_wrdSect.PageSetup.RightMargin
                    tbl.Range.Cells.Item(2).Range.Style = glb_get_wrdActiveDoc.Styles.Item("Body Text")
                    '
                    '
                    tbl.Rows.LeftIndent = -glb_get_wrdSect.PageSetup.LeftMargin
                    tbl.Shading.ForegroundPatternColor = objGlobals._glb_colour_Recommendation_Purple
                    '
                    'Now do the caption according to the strType
                    rng = tbl.Range.Cells.Item(1).Range
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    rng = Me.Plh_Captions_InsertCaptions(strType, rng, True)
                    rng.Select()
                    '
                    'Now do the Source Cell
                    'drCell = tbl.Range.Cells.Item(3)
                    'Me.Plh_Table_FixCellPadding(drCell, 0.6, 5.65, drCell.LeftPadding, drCell.RightPadding)
                    'drCell.Range.Style = objGlobals.glb_get_wrdStyle("Source")
                    'rng2 = drCell.Range
                    'rng2.Collapse(WdCollapseDirection.wdCollapseStart)
                    'Me.Plh_Insert_SourceAndNoteText(rng2)
                Else
                    tbl = Me.Plh_Table_Build(rng, 2, 1, False, False, False)
                    tbl.Range.Cells.Item(2).Range.Style = glb_get_wrdActiveDoc.Styles.Item("Body Text")
                    'Now do the caption according to the strType


                    objFloatMgr.Plh_Float_LockInPosition_RelativeToLeftPageEdge(tbl, 0.0, glb_get_wrdSect.PageSetup.PageWidth)
                    tbl.Shading.ForegroundPatternColor = objGlobals._glb_colour_Recommendation_Purple
                    '
                    rng = tbl.Range.Cells.Item(1).Range
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    rng = Me.Plh_Captions_InsertCaptions(strType, rng, True)
                    rng.Select()
                    '
                End If

                '




            Case "Figure_ES", "Figure", "Figure_AP", "Figure_LT"
                'tbl = Me.Plh_Table_Build(rng, 4, 1, False, False, False)
                tbl = Me.Plh_Table_Build(rng, 3, 1, False, False, False)
                glb_tbl_apply_aacTableNoLinesStyle(tbl)
                '
                '
                tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                tbl.PreferredWidth = 100
                '

                'tbl.Style = rng.Document.Styles.Item("aac Table (no lines)")
                '
                tblWidth = tbl.Columns.Item(1).Width
                drCell = tbl.Range.Cells.Item(1)
                'Me.Plh_Table_FixCellPadding(drCell, 8.0, 0.4, 0.0, 0.0)
                '
                dr = tbl.Range.Rows.Item(2)
                Try
                    dr.Range.Style = objGlobals.glb_get_wrdStyle("Figure")
                Catch ex1 As Exception
                    myStyle = glb_get_wrdActiveDoc.Styles.Add("Figure", WdStyleType.wdStyleTypeParagraph)
                    myStyle.Font.Size = 8.0
                    myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                    dr.Range.Style = objGlobals.glb_get_wrdStyle("Figure")

                End Try
                '
                'Now do the picture cell so that it has 04 pt top and bottom padding. This
                'is to ensure that the Figure/Image doesn't partially obscure (during pdf)
                'any top/bottom enclosing lines
                drCell = tbl.Range.Cells.Item(2)
                'Me.Plh_Table_FixCellPadding(drCell, 2.0, 2.0, 0.0, 0.0)

                'brdr = dr.Borders.Item(WdBorderType.wdBorderTop)
                'brdr.Visible = True
                'brdr.LineStyle = WdLineStyle.wdLineStyleSingle
                'brdr.LineWidth = WdLineWidth.wdLineWidth050pt
                'brdr.Color = borderColour
                '

                '
                drCell = tbl.Range.Cells.Item(3)
                Me.Plh_Table_FixCellPadding(drCell, 0.6, 5.65, 0.0, 0.0)
                'dr = tbl.Range.Rows.Item(tbl.Range.Rows.Last.Index - 1)
                dr = tbl.Rows.Last

                Try
                    dr.Range.Style = objGlobals.glb_get_wrdStyle("Source")
                Catch ex1 As Exception
                    myStyle = glb_get_wrdActiveDoc.Styles.Add("Source", WdStyleType.wdStyleTypeParagraph)
                    myStyle.Font.Size = 8.0
                    myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                    myStyle.ParagraphFormat.SpaceBefore = 2.6
                    '
                    dr.Range.Style = objGlobals.glb_get_wrdStyle("Source")
                    '
                End Try
                rng2 = dr.Range
                rng2.Collapse(WdCollapseDirection.wdCollapseStart)
                Me.Plh_Insert_SourceAndNoteText(rng2)
                '
                'dr = tbl.Rows.Last
                'dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
                'dr.Height = 8.0
                '
                'GoTo finis
                '
                'Add a new row on the bottom for spacing purposes
                'drCell = tbl.Range.Cells.Item(3)
                'drCell.Split(2, 1)
                'dr = tbl.Rows.Last
                'dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
                'dr.Height = 8.0
                '
                '
                'Do picture section top border and bottom borders
                '
                'drCell = tbl.Range.Cells.Item(2)
                'drCell.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop
                'brdr = drCell.Borders.Item(WdBorderType.wdBorderTop)
                'brdr.Visible = False
                ''brdr.LineStyle = WdLineStyle.wdLineStyleSingle
                'brdr.LineWidth = WdLineWidth.wdLineWidth050pt
                'brdr.Color = borderColour
                '
                'drCell = tbl.Range.Cells.Item(2)
                'brdr = drCell.Borders.Item(WdBorderType.wdBorderTop)
                'brdr.Visible = True
                'brdr.LineStyle = WdLineStyle.wdLineStyleSingle
                'brdr.LineWidth = WdLineWidth.wdLineWidth050pt
                'brdr.Color = borderColour
                '
                'drCell = tbl.Range.Cells.Item(3)
                'brdr = drCell.Borders.Item(WdBorderType.wdBorderBottom)
                'brdr.Visible = False
                'brdr.LineStyle = WdLineStyle.wdLineStyleSingle
                'brdr.LineWidth = WdLineWidth.wdLineWidth050pt
                'brdr.Color = borderColour
                '
                'GoTo finis
                Try
                    'Insert and format temporary picture placeholder
                    'rng = tbl.Range.Cells.Item(2).Range
                    'rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    '
                    'Globals.ThisAddin.Application.Selection.InlineShapes.AddPicture()
                    'para = rng.Paragraphs.Item(1)
                    'para.a
                    '
                    objGrph.grfx_inline_insertShape(tbl.Range.Cells.Item(2), tblWidth, 108.0, Me._glb_colour_FigureFill)
                    '
                    'rngLocal = objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_Img_FillPict", "Images", rng)
                    'shp = rngLocal.InlineShapes.Item(1)
                    '
                    'shp.LockAspectRatio = False
                    'shp.Width = tblWidth
                    'shp.Height = 108.0
                    '
                    'shp.

                Catch ex As Exception
                    MsgBox("Error in image insert")
                End Try
                '
                'Now do the caption according to the strType
                rng = tbl.Range.Cells.Item(1).Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng = Me.Plh_Captions_InsertCaptions(strType, rng, True)
                rng.Select()
        '

            '

            Case "Table_ES", "Table", "Table_AP", "Table_LT"
                tbl = Me.Plh_Table_Build(rng, 3, 1, True, False, False)
                '
                'Top Row
                dr = tbl.Range.Rows.Item(1)
                dr.Cells.Item(1).BottomPadding = 2.0
                dr.HeightRule = WdRowHeightRule.wdRowHeightAuto
                dr.Range.Style = objGlobals.glb_get_wrdStyle("Table text")
                dr.Range.Text = "Paste table, Or a picture Of a table On this paragraph"
                '
                'Bottom (spacer) row
                dr = tbl.Rows.Last
                dr.Range.Style = objTblsMgr.glb_get_wrdStyle(objTblsMgr.glb_var_style_tblSpacerStyle)
                dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
                dr.Height = 8.0
                '
                'Now do the Source Cell
                drCell = tbl.Range.Cells.Item(tbl.Range.Cells.Count - 1)
                Me.Plh_Table_FixCellPadding(drCell, 0, 2.0, 0.0, 0.0)
                dr = tbl.Range.Rows.Item(tbl.Range.Rows.Last.Index - 1)
                dr.Range.Style = objGlobals.glb_get_wrdStyle("Source")
                rng2 = dr.Range
                rng2.Collapse(WdCollapseDirection.wdCollapseStart)
                rng2 = Me.Plh_Insert_SourceAndNoteText(rng2, "sourceAndNote")
                '
                rng = objTblsMgr.tbl_captions_Insert(tbl, strType)
                para = rng.Paragraphs.Item(1)
                '
                'objTblsMgr.tbl_caption_doIndent(para, tbl)
                '
                'Now do the caption according to the strType
                'rng = tbl.Range.Cells.Item(1).Range
                'rng.Collapse(WdCollapseDirection.wdCollapseStart)
                'rng = Me.Plh_Captions_InsertCaptions(strType, rng, True)
                'rng.Select()


        End Select
        '
finis:
        'dr = tbl.Range.Rows.Item(2)
        Return tbl
        '
    End Function
    '
    ''' <summary>
    ''' This method will search the specified section for all 'Findings, Recommendations and
    ''' Case Study (half page)' tables and will set the width to percentage, and to a value that
    ''' guarantees full with. GThis functiom is used when sections are re-oriented,
    ''' </summary>
    ''' <param name="sect"></param>
    Public Sub Plh_setAll_FindingEtc_Width(ByRef sect As Word.Section)
        Dim tbl As Word.Table
        Dim strText As String
        '
        For Each tbl In sect.Range.Tables
            strText = tbl.Range.Cells.Item(1).Range.Text
            If strText Like "Finding*" Or strText Like "Recommendation*" Or strText Like "Case Study*" Then
                tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                tbl.PreferredWidth = Me.Plh_get_FindingsEtc_Width_As_Percent(tbl)
            End If
        Next

    End Sub
    '
    ''' <summary>
    ''' This method will get the size (as percent) of the full width 'Findings, Recommendations and
    ''' Case Study (half page).
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function Plh_get_FindingsEtc_Width_As_Percent(ByRef tbl As Word.Table) As Single
        Dim sect As Word.Section
        Dim tblWidthAsPercent As Single
        '
        'For Portrait or Landscape the full page width percentage calculation is the same. Note that
        'The Percentage specs need to be at the end. If we place it at the beginning, the specs for left
        'and right padding (in pts) forces the table back actual measurements for width
        'If sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait And sect.PageSetup.TextColumns.Count = 1 Then
        '
        sect = tbl.Range.Sections.Item(1)
        tblWidthAsPercent = 100
        '
        If sect.PageSetup.TextColumns.Count = 1 Then
            '
            'tbl.Columns.Item(1).Width = glb_get_wrdSect.PageSetup.PageWidth
            tbl.LeftPadding = glb_get_wrdSect.PageSetup.LeftMargin
            tbl.Rows.LeftIndent = -glb_get_wrdSect.PageSetup.LeftMargin
            tbl.RightPadding = glb_get_wrdSect.PageSetup.RightMargin
            tbl.Range.Cells.Item(2).Range.Style = glb_get_wrdActiveDoc.Styles.Item("Body Text")
            'tbl.Range.Cells.Item(2).BottomPadding = 6.0
            '
            '***
            '
            'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
            'tbl.PreferredWidth = 100
            tblWidthAsPercent = 100 * sect.PageSetup.PageWidth / glb_get_widthBetweenMargins(sect)
            '
            '***

        End If
        '
        Return tblWidthAsPercent
        '
    End Function
    '
    Public Overridable Function Plh_insert_PlaceHolder_Wide(strType As String) As Word.Table
        Dim tbl, tblheader As Word.Table
        Dim objGlobals As New cGlobals()
        Dim objTools As New cTools()
        Dim objHFMgr As New cHeaderFooterMgr()
        Dim objTblMgr As New cPlHTable()
        Dim objFloatMgr As New cPlHFloatingMgr()
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim sect As Word.Section
        Dim marginWidth, leftIndent As Single
        Dim offSet As Single
        Dim numTextColumns As Integer
        Dim shp As Word.InlineShape
        Dim strColType, strResult As String
        '
        tbl = Nothing
        strColType = ""
        offSet = -40.0
        sect = objGlobals.glb_get_wrdSect
        marginWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
        '
        numTextColumns = sect.PageSetup.TextColumns.Count
        '
        strResult = Me.Plh_is_OKToInsert(sect)
        '
        If strResult = "" Then
            Select Case numTextColumns
                Case 1
                    tbl = Me.Plh_insert_PlaceHolderBasic(objGlobals.glb_get_wrdSelRng, strType)
                    'tbl = MyBase.Plh_PlaceHolder_Insert(strType)
                    tblheader = objHFMgr.hf_get_HeaderTable(sect)
                    '
                    tbl.Columns.Item(1).Width = tbl.Columns.Item(1).Width - tblheader.Rows.Item(1).LeftIndent - objTools.tools_math_MillimetersToPoints(objTblMgr.tbl_OutDent)
                    '
                    leftIndent = tblheader.Rows.Item(1).LeftIndent + objTools.tools_math_MillimetersToPoints(objTblMgr.tbl_OutDent)
                    For Each dr In tbl.Rows
                        dr.LeftIndent = leftIndent
                        drCell = dr.Range.Cells.Item(1)
                        'Problem with values for "textColumns.spacing" when 
                        'drCell.Width = sect.PageSetup.TextColumns.Item(1).Width + sect.PageSetup.TextColumns.Spacing + sect.PageSetup.TextColumns.Item(2).Width
                        'drCell.Width = marginWidth
                        If dr.Index = 2 Then
                            If drCell.Range.InlineShapes.Count <> 0 Then
                                shp = drCell.Range.InlineShapes.Item(1)
                                shp.Width = drCell.Width
                            End If
                        End If

                    Next
                '                    
                'Me.PlhFig_Figure_adjustPlaceHolder(tbl, RGB(0, 1, 0))
                '
                Case 2
                    '
                    If Not (Me.Plh_Columnsx2_FindColumnNumber(sect) = 1) Then
                        MsgBox("For multi-column layouts, wide figures" + vbCrLf + "can only be inserted In column 1" + vbCrLf + vbCrLf + "Please relocate your selection point" + vbCrLf + "And Try again")
                        GoTo finis
                    End If
                    '
                    tbl = Me.Plh_insert_PlaceHolderBasic(objGlobals.glb_get_wrdSelRng, strType)
                    'tbl = MyBase.Plh_PlaceHolder_Insert(strType)
                    'Me.PlhFig_Figure_adjustPlaceHolder(tbl, RGB(0, 1, 0))
                    '
                    For Each dr In tbl.Rows
                        drCell = dr.Range.Cells.Item(1)
                        'Problem with values for "textColumns.spacing" when 
                        'drCell.Width = sect.PageSetup.TextColumns.Item(1).Width + sect.PageSetup.TextColumns.Spacing + sect.PageSetup.TextColumns.Item(2).Width
                        drCell.Width = marginWidth
                        If dr.Index = 2 Then
                            If drCell.Range.InlineShapes.Count <> 0 Then
                                shp = drCell.Range.InlineShapes.Item(1)
                                shp.Width = drCell.Width
                            End If
                        End If
                    Next
                    '
                    'rng = sect.PageSetup.TextColumns.Item(1).
                    'Globals.ThisAddin.Application.Selection.InRange(rng)
                    objFloatMgr.PlHFloat_lock_toParagraphAndMarginLeft(tbl)
                'Me.Plh_Float_LockToToParagraph(tbl)
                    '
                '
                Case 3
                    If Not (Me.Plh_Columnsx2_FindColumnNumber(sect) = 1) Then
                        MsgBox("For multi-column layouts, wide figures" + vbCrLf + "can only be inserted In column 1" + vbCrLf + vbCrLf + "Please relocate your selection point" + vbCrLf + "And Try again")
                        GoTo finis
                    End If
                    '
                    tbl = Me.Plh_insert_PlaceHolderBasic(objGlobals.glb_get_wrdSelRng, strType)
                    '
                    For Each dr In tbl.Rows
                        drCell = dr.Range.Cells.Item(1)
                        drCell.Width = sect.PageSetup.TextColumns.Item(1).Width + sect.PageSetup.TextColumns.Spacing + sect.PageSetup.TextColumns.Item(2).Width + sect.PageSetup.TextColumns.Spacing + sect.PageSetup.TextColumns.Item(3).Width
                        If dr.Index = 2 Then
                            If drCell.Range.InlineShapes.Count <> 0 Then
                                shp = drCell.Range.InlineShapes.Item(1)
                                shp.Width = drCell.Width
                            End If
                        End If
                        '
                    Next
                    '
                    objFloatMgr.PlHFloat_lock_toParagraphAndMarginLeft(tbl)

            End Select
        Else
            MsgBox(strResult)
        End If


finis:

        '
        Return tbl

    End Function

    '
    ''' <summary>
    ''' This method expects the PlaceHolder Type (strType)which is typically the name of the
    ''' Sequence field used in the Captions for that type. It will insert a caption at rng
    ''' and it will apply the Caption C
    ''' </summary>
    ''' <param name="strType"></param>
    ''' <param name="rng"></param>
    ''' <param name="doCaptionStyle"></param>
    ''' <returns></returns>
    Public Function Plh_Captions_InsertCaptions(strType As String, rng As Word.Range, doCaptionStyle As Boolean, Optional strCaptiontext As String = "") As Word.Range
        Dim rngSelection As Word.Range
        Dim objParas As New cParas()
        Dim allIsOK As Boolean
        Dim tokens() As String
        Dim para, paraCaption As Word.Paragraph
        Dim hangingIndent As Single
        Dim j As Integer
        '
        allIsOK = True
        rngSelection = Nothing
        hangingIndent = Me.var_glb_style_tblCaption_Line2_Indent
        '
        '
        Select Case strType
                            '
            Case "CaseStudy_HalfPage", "CaseStudy"
                If strCaptiontext = "" Then strCaptiontext = "Title of Case Study"
                Me.Plh_Captions_InsertNameBeforeFields("Case Study", rng)
                rngSelection = Me.Plh_Captions_Insert("notSpecific", "CaseStudy", rng, strCaptiontext)
                'hangingIndent = 84
                hangingIndent = rng.Sections.Item(1).PageSetup.LeftMargin - glb_math_MillimetersToPoints(_glb_header_leftEdge)
                j = 1
                '
            Case "Box_ES"
                If strCaptiontext = "" Then strCaptiontext = "Title of Box"
                Me.Plh_Captions_InsertNameBeforeFields("Box" & Chr(160) & "ES", rng)
                rngSelection = Me.Plh_Captions_Insert("ES", "Box_ES", rng, strCaptiontext)
                hangingIndent = 60.0

            Case "Box"
                If strCaptiontext = "" Then strCaptiontext = "Title of Box"
                Me.Plh_Captions_InsertNameBeforeFields("Box", rng)
                rngSelection = Me.Plh_Captions_Insert("Report", "Box", rng, strCaptiontext)
                hangingIndent = 55.0
                '
            Case "Box_AP"
                If strCaptiontext = "" Then strCaptiontext = "Title of Box"
                Me.Plh_Captions_InsertNameBeforeFields("Box", rng)
                rngSelection = Me.Plh_Captions_Insert("Appendix", "Box_AP", rng, strCaptiontext)
                hangingIndent = 65.0
                '
            Case "Box_LT"
                If strCaptiontext = "" Then strCaptiontext = "Title of Box"
                Me.Plh_Captions_InsertNameBeforeFields("Box", rng)
                rngSelection = Me.Plh_Captions_Insert("notSpecific", "Box_LT", rng, strCaptiontext)
                hangingIndent = 45.0
                '
            Case "Figure_ES"
                If strCaptiontext = "" Then strCaptiontext = "Title of Figure"
                Me.Plh_Captions_InsertNameBeforeFields("Figure" & Chr(160) & "ES", rng)
                rngSelection = Me.Plh_Captions_Insert("ES", "Figure_ES", rng, strCaptiontext)
                hangingIndent = 75.0

            Case "Figure"
                If strCaptiontext = "" Then strCaptiontext = "Title of Figure"
                Me.Plh_Captions_InsertNameBeforeFields("Figure", rng)
                rngSelection = Me.Plh_Captions_Insert("Report", "Figure", rng, strCaptiontext)
                hangingIndent = 70.0
                '                                '
            Case "Figure_AP"
                If strCaptiontext = "" Then strCaptiontext = "Title of Figure"
                Me.Plh_Captions_InsertNameBeforeFields("Figure", rng)
                rngSelection = Me.Plh_Captions_Insert("Appendix", "Figure_AP", rng, strCaptiontext)
                hangingIndent = 75.0
                '
            Case "Figure_LT"
                If strCaptiontext = "" Then strCaptiontext = "Title of Figure"
                Me.Plh_Captions_InsertNameBeforeFields("Figure", rng)
                rngSelection = Me.Plh_Captions_Insert("notSpecific", "Figure_LT", rng, strCaptiontext)
                hangingIndent = 75.0
                '                                '
            Case "Table_ES"
                If strCaptiontext = "" Then strCaptiontext = "Title of Table"
                Me.Plh_Captions_InsertNameBeforeFields("Table" & Chr(160) & "ES", rng)
                rngSelection = Me.Plh_Captions_Insert("ES", "Table_ES", rng, strCaptiontext)
                hangingIndent = 65.0
                '
            Case "Table"
                If strCaptiontext = "" Then strCaptiontext = "Title of Table"
                Me.Plh_Captions_InsertNameBeforeFields("Table", rng)
                rngSelection = Me.Plh_Captions_Insert("Report", "Table", rng, strCaptiontext)
                hangingIndent = 60.0
                '
            Case "Table_AP"
                If strCaptiontext = "" Then strCaptiontext = "Title of Table"
                Me.Plh_Captions_InsertNameBeforeFields("Table", rng)
                rngSelection = Me.Plh_Captions_Insert("Appendix", "Table_AP", rng, strCaptiontext)
                hangingIndent = 65.0
                '
            Case "Table_LT"
                If strCaptiontext = "" Then strCaptiontext = "Title of Table"
                Me.Plh_Captions_InsertNameBeforeFields("Table", rng)
                rngSelection = Me.Plh_Captions_Insert("notSpecific", "Table_LT", rng, strCaptiontext)
                hangingIndent = 65.0
                '
            Case "Key_Finding"
                'If strCaptiontext = "" Then strCaptiontext = "Title Of Key Finding"
                'Me.Plh_Captions_InsertNameBeforeFields("Key" & Chr(160) & "Finding", rng)
                'rngSelection = Me.Plh_Captions_Insert("notSpecific", "Key_Finding", rng, strCaptiontext)
                '
                If strCaptiontext = "" Then strCaptiontext = "Title of Finding"
                Me.Plh_Captions_InsertNameBeforeFields("Finding", rng)
                rngSelection = Me.Plh_Captions_Insert("notSpecific", "Finding", rng, strCaptiontext)
                hangingIndent = 65

                '
            Case "Key_Finding_ES"
                'If strCaptiontext = "" Then strCaptiontext = "Title Of Key Finding"
                'Me.Plh_Captions_InsertNameBeforeFields("Key" + Chr(160) + "Finding" + Chr(160) + "ES", rng)
                'rngSelection = Me.Plh_Captions_Insert("notSpecific", "Key_Finding_ES", rng, strCaptiontext)
                '
                If strCaptiontext = "" Then strCaptiontext = "Title of Finding"
                Me.Plh_Captions_InsertNameBeforeFields("Finding" + Chr(160) + "ES", rng)
                rngSelection = Me.Plh_Captions_Insert("notSpecific", "Finding_ES", rng, strCaptiontext)
                hangingIndent = 80
                '
            Case "Recommendation"
                If strCaptiontext = "" Then strCaptiontext = "Title of Recommendation"
                Me.Plh_Captions_InsertNameBeforeFields("Recommendation", rng)
                rngSelection = Me.Plh_Captions_Insert("notSpecific", "Recommendation", rng, strCaptiontext)
                hangingIndent = 115
                '
            Case "Recommendation_ES"
                If strCaptiontext = "" Then strCaptiontext = "Title of Recommendation"
                Me.Plh_Captions_InsertNameBeforeFields("Recommendation" + Chr(160) + "ES", rng)
                rngSelection = Me.Plh_Captions_Insert("notSpecific", "Recommendation_ES", rng, strCaptiontext)
                hangingIndent = 118
                '
            Case "Recommendation_LT"
                If strCaptiontext = "" Then strCaptiontext = "Title of Recommendation"
                Me.Plh_Captions_InsertNameBeforeFields("Box", rng)
                rngSelection = Me.Plh_Captions_Insert("notSpecific", "Recommendation_LT", rng, strCaptiontext)
                hangingIndent = 118
                '
            Case Else
                allIsOK = False

        End Select
        '
        '
        Try
            If allIsOK Then
                If doCaptionStyle Then
                    para = rngSelection.Paragraphs.Item(1)
                    rng = para.Range
                    rng.Style = objGlobals.glb_get_wrdStyle(objGlobals.glb_var_style_tblCaptionStyle)
                End If
                '
                paraCaption = objParas.paras_set_HangingIndent(rng, hangingIndent)
                'paraCaption = objParas.paras_set_HangingIndent(rng, 60.1)
                '***
                '
                rng = paraCaption.Range
                '
                'rng.MoveStart(WdUnits.wdParagraph, -1)
                tokens = Split(rng.Text, vbTab)
                '
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.MoveEnd(WdUnits.wdCharacter, tokens(0).Count)
                '
                'rng.MoveEnd(WdUnits.wdCharacter, -(tokens(1).Length + 1))
                glb_get_wrdActiveDoc()
                'rng.Font.Bold = True
                'rng.Style = glb_get_wrdActiveDoc.Styles("Caption Label")
                rng.Font.Bold = True
                '
                'If we are inserting in table cell, we'll assume that the cell has been formatted with
                'the Caption Style... But for free text insertion in the body of he text, we'll need to
                'format the paragraph to Caption Style
                '

            End If
        Catch ex As Exception

        End Try
        '
        Return rngSelection
        '
    End Function
    '
    ''' <summary>
    ''' This method will directly apply the Caption style to the 1st paragraph in rng, which is typically some text
    ''' formatted as caption text, but not in the Caption Style.. This is extracted from the code above
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="hangingIndent"></param>
    Public Sub Plh_Captions_doCaptionStyle(ByRef rng As Word.Range, Optional hangingIndent As Single = 115.0)
        Dim para, paraCaption As Word.Paragraph
        Dim objParas As New cParas()
        Dim rng2 As Word.Range
        Dim tokens() As String

        para = rng.Paragraphs.Item(1)
        para.Range.Style = objGlobals.glb_get_wrdStyle(objGlobals.glb_var_style_tblCaptionStyle)
        '
        paraCaption = objParas.paras_set_HangingIndent(para.Range, hangingIndent)
        'paraCaption = objParas.paras_set_HangingIndent(rng, 60.1)
        '***
        '
        rng2 = paraCaption.Range
        '
        'rng.MoveStart(WdUnits.wdParagraph, -1)
        tokens = Split(rng2.Text, vbTab)
        '
        rng2.Collapse(WdCollapseDirection.wdCollapseStart)
        rng2.MoveEnd(WdUnits.wdCharacter, tokens(0).Count)
        '
        'Either set the label to bold or to the character style 'Caption Label'
        'rng.Font.Bold = True
        rng2.Style = glb_get_wrdActiveDoc.Styles("Caption Label")
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method takes the string strPlhName, adds a special space character "Chr(160)" and
    ''' inserts it at rng.. Typically this is used to insert the Caption label (e.g. Box 1.1)
    ''' before the Capation fields
    ''' </summary>
    ''' <param name="strPlhLabelName"></param>
    ''' <param name="rng"></param>
    Public Sub Plh_Captions_InsertNameBeforeFields(strPlhLabelName As String, ByRef rng As Word.Range)
        Dim strNewPlhLabel As String
        'Dim strSeqUpdate As String

        strNewPlhLabel = strPlhLabelName & Chr(160)
        'strNewPlhLabel = strPlhLabelName & Chr(32)

        'strSeqUpdate = "SEQ " & strSequenceId & "*"
        'strSeqUpdate = StrConv(strSeqUpdate, vbUpperCase)
        '
        rng.Text = strNewPlhLabel                               'The range includes all of the added text
        rng.MoveEnd(WdUnits.wdCharacter, -1)
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng.Move(WdUnits.wdCharacter, 1)
        '
    End Sub
    '

    '
    Public Function Plh_Captions_ConvertCaptions(lstOfTargetFields As List(Of String), strToNewType As String, ByRef rngSrc As Word.Range) As Word.Range
        Dim rngSelection, rng, rng2 As Word.Range
        Dim tokens() As String
        Dim strExistingCaptionText As String
        Dim para As Word.Paragraph
        Dim fld As Word.Field
        Dim strFieldCode As String
        Dim lstParas As New List(Of Word.Paragraph)
        Dim rngOld As Word.Range
        '*****
        Dim lst As New Collection()
        Dim strFieldCodeOfTargetFields As String
        Dim i As Integer
        '
        rngSelection = Nothing
        rngOld = Me.glb_get_wrdSel.Range
        '
        'lstParas = Me.Plh_Captions_GetFieldParagraphs(strFieldCodeOfTargetFields, rngSrc)
        '
        'lst.Add(strFieldCodeOfTargetFields)
        'lst.Add("*SEQ KEY_FINDING*")
        'lst.Add("*SEQ KEY_FINDING_ES*")
        'lst.Add("*SEQ RECOMMENDATION*")
        'lst.Add("*SEQ RECOMMENDATION_ES*")
        '
        For Each fld In rngSrc.Fields
            strFieldCode = fld.Code.Text
            For i = 0 To lstOfTargetFields.Count - 1
                strFieldCodeOfTargetFields = "*" + lstOfTargetFields.Item(i) + "*"
                If strFieldCode Like strFieldCodeOfTargetFields Then
                    fld.Select()
                    rng2 = Me.glb_get_wrdSel.Range
                    para = rng2.Paragraphs.Item(1)
                    lstParas.Add(para)
                End If
            Next
        Next

        '
        For Each para In lstParas
            Try
                rng = para.Range
                rng.MoveEnd(WdUnits.wdCharacter, -1)
                tokens = Split(rng.Text, vbTab)
                strExistingCaptionText = tokens(1)
                '
                strFieldCode = ""
                '
                'This section will delte the field codes in the paragraph, but it will
                'save the last Field.Code in strFieldCode... This last field code is
                'always the sequence field
                '
                fld = rng.Fields.Item(rng.Fields.Count)
                strFieldCode = fld.Code.Text
                '
                For Each fld In rng.Fields
                    fld.Delete()
                Next
                '
                'para.Range.Delete()
                'para.Range.Text = ""
                rng = para.Range
                rng.MoveEnd(WdUnits.wdCharacter, -1)
                rng.Delete()
                'rng.Select()
                'rng.Collapse(WdCollapseDirection.wdCollapseStart)
                'Globals.ThisAddin.Application.ScreenUpdating = True
                'MsgBox("Stop")
                'Globals.ThisAddin.Application.ScreenRefresh()
                '
                If strFieldCode Like "*Box*" And strToNewType = "Box_ES" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)
                If strFieldCode Like "*Key*" And strToNewType = "Box_ES" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)
                If strFieldCode Like "*Recommendation*" And strToNewType = "Box_ES" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)
                '
                If strFieldCode Like "*Box*" And strToNewType = "Box" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)
                If strFieldCode Like "*Key*" And strToNewType = "Box" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)
                If strFieldCode Like "*Recommendation*" And strToNewType = "BOX" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)
                '
                If strFieldCode Like "*Box*" And strToNewType = "Box_AP" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)
                '
                If strFieldCode Like "*Box*" And strToNewType = "Box_LT" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)
                If strFieldCode Like "*Key*" And strToNewType = "Box_LT" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)
                If strFieldCode Like "*Recommendation*" And strToNewType = "Box_LT" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)
                '
                If strFieldCode Like "*Figure*" And strToNewType = "Figure_ES" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)
                If strFieldCode Like "*Figure*" And strToNewType = "Figure" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)
                If strFieldCode Like "*Figure*" And strToNewType = "Figure_AP" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)
                If strFieldCode Like "*Figure*" And strToNewType = "Figure_LT" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)
                '
                If strFieldCode Like "*Table*" And strToNewType = "Table_ES" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)
                If strFieldCode Like "*Table*" And strToNewType = "Table" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)
                If strFieldCode Like "*Table*" And strToNewType = "Table_AP" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)
                If strFieldCode Like "*Table*" And strToNewType = "Table_LT" Then rngSelection = Me.Plh_Captions_InsertCaptions(strToNewType, rng, True, strExistingCaptionText)



            Catch ex As Exception

            End Try
        Next
        '
        rngOld.Select()
        '
        Return rngSelection
        '
    End Function
    '
    ''' <summary>
    ''' This method will write the appropriate table caption to the range rng. If strForceCaptioTo
    ''' is set to "", then this method will gues the caption based on the page numbering of the section
    ''' that the tbale is in. If it is not "", then stForceCaption = 'LT', 'ES', 'BD' and 'AP' will
    ''' force the caption to 'Letter', 'Executive Summary', 'Report Body' and 'Appendix'.
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="strForceCaptionTo"></param>
    ''' <returns></returns>
    Public Function Plh_Captions_getAndWriteCaption(ByRef rng As Word.Range, Optional strForceCaptionTo As String = "") As Word.Range
        Dim objChptBase As New cChptBase()
        Dim objPlhBase As New cPlHBase()
        Dim objTagsMgr As New cTagsMgr()
        Dim sect As Word.Section
        '
        sect = rng.Sections.Item(1)
        '
        If strForceCaptionTo = "" Then
            If objTagsMgr.tags_is_Letter(sect) Then
                rng = Me.Plh_Captions_InsertCaptions("Table_LT", rng, True)
                GoTo loop1
                '
            End If
            '
            If objTagsMgr.tags_is_Brief(sect) Then
                rng = Me.Plh_Captions_InsertCaptions("Table_ES", rng, True)
                GoTo loop1
                '
            End If
            '
            If objChptBase.chptBase_PageNumbering_isChapterBody() Then
                'Use this to determine which Table Caption option to use
                rng = Me.Plh_Captions_InsertCaptions("Table", rng, True)
                '
            End If
            If objChptBase.chptBase_PageNumbering_isES() Then
                'Use this to determine which Table Caption option to use
                rng = Me.Plh_Captions_InsertCaptions("Table_ES", rng, True)
                '
            End If
            If objChptBase.chptBase_PageNumbering_isAppendixBody() Then
                'Use this to determine which Table Caption option to use
                rng = Me.Plh_Captions_InsertCaptions("Table_AP", rng, True)
                '
            End If
loop1:
        Else
            Select Case strForceCaptionTo
                Case "LT"
                    rng = Me.Plh_Captions_InsertCaptions("Table_LT", rng, True)
                Case "ES"
                    rng = Me.Plh_Captions_InsertCaptions("Table_ES", rng, True)
                Case "BD"
                    rng = Me.Plh_Captions_InsertCaptions("Table", rng, True)
                Case "AP"
                    rng = Me.Plh_Captions_InsertCaptions("Table_AP", rng, True)
                Case "none"

            End Select
        End If
        '
        Return rng
        '
    End Function

    '
    Public Function Plh_Captions_Insert(strWhere As String, strSequenceName As String, ByRef rng As Word.Range, strCaptionText As String) As Word.Range
        Dim fld, fld2 As Word.Field
        Dim rngSelection, rng2 As Word.Range
        '
        rngSelection = Nothing
        '
        Select Case strWhere
            Case "ES"
                fld = rng.Fields.Add(rng, WdFieldType.wdFieldSequence, strSequenceName + " \* ARABIC ")
                fld.Select()
                rng = fld.Application.Selection.Range
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                'rng.Move(WdUnits.wdCharacter, 1)
                fld.Update()
                '
                rng.Text = vbTab
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                'rng.Move(WdUnits.wdCharacter, 1)
                rng.Text = strCaptionText
                rngSelection = rng.Duplicate()
                '
            Case "Report"
                '
                rng2 = rng.Duplicate()
                rng2.Text = "."
                rng2.Collapse(WdCollapseDirection.wdCollapseEnd)
                '
                fld2 = rng2.Fields.Add(rng2, WdFieldType.wdFieldSequence, strSequenceName + " \* ARABIC \S 1")
                fld = rng.Fields.Add(rng, WdFieldType.wdFieldStyleRef, "1 \S", True)
                '
                fld2.Update()
                fld.Update()
                '
                fld2.Select()
                rng2 = fld2.Application.Selection.Range
                rng2.Collapse(WdCollapseDirection.wdCollapseEnd)
                '
                rng2.Text = vbTab
                rng2.Collapse(WdCollapseDirection.wdCollapseEnd)
                rng2.Text = strCaptionText
                '
                rngSelection = rng2
                '
            Case "Appendix"
                '
                'Have to fiddle this one... The style ref 9 field will cause a problem if there
                'is no heading Level 9 in the document... The fault will stop the SEQ field
                'from being inserted... So we insert the SEQ field first and then the Styleref. This causes a problem
                'with the Captions Conversion method, because it looks for a SEQ field
                '
                'Need to do this for the Chapet Headings as well
                '
                '***
                '
                rng2 = rng.Duplicate()
                rng2.Text = "."
                rng2.Collapse(WdCollapseDirection.wdCollapseEnd)
                '
                'fld2 = rng2.Fields.Add(rng2, WdFieldType.wdFieldSequence, strSequenceName + " \* ARABIC \S 9")
                'fld = rng.Fields.Add(rng, WdFieldType.wdFieldStyleRef, "9 \S", True)
                '
                fld2 = rng2.Fields.Add(rng2, WdFieldType.wdFieldSequence, strSequenceName + " \* ARABIC \S 6")
                fld = rng.Fields.Add(rng, WdFieldType.wdFieldStyleRef, "6 \S", True)
                '
                fld2.Update()
                fld.Update()
                '
                fld2.Select()
                rng2 = fld2.Application.Selection.Range
                rng2.Collapse(WdCollapseDirection.wdCollapseEnd)
                '
                rng2.Text = vbTab
                rng2.Collapse(WdCollapseDirection.wdCollapseEnd)
                rng2.Text = strCaptionText
                '
                rngSelection = rng2
                '


                '***
                '*** Old version
                '
                'rng.Text = "."
                'rng.Move(WdUnits.wdCharacter, 1)
                'fld2 = rng.Fields.Add(rng, WdFieldType.wdFieldSequence, strSequenceName + " \* ARABIC \S 9")
                'rng.Move(WdUnits.wdCharacter, 1)
                'rng2 = rng.Duplicate
                '
                'rng.Move(WdUnits.wdCharacter, -2)
                'fld = rng.Fields.Add(rng, WdFieldType.wdFieldStyleRef, "9 \S", True)
                '            
                'fld2.Update()
                'fld.Update()
                '
                'rng2.Text = vbTab
                'rng2.Move(WdUnits.wdCharacter, 1)
                'rng2.Text = strCaptionText
                'rngSelection = rng2.Duplicate()
            Case "notSpecific"
                '
                '***
                fld = rng.Fields.Add(rng, WdFieldType.wdFieldSequence, strSequenceName + " \* ARABIC ")
                fld.Select()
                rng = fld.Application.Selection.Range
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                'rng.Move(WdUnits.wdCharacter, 1)
                fld.Update()
                '
                rng.Text = vbTab
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                'rng.Move(WdUnits.wdCharacter, 1)
                rng.Text = strCaptionText
                rngSelection = rng.Duplicate()
                '

                '***
                '*** Olde Version
                'fld = rng.Fields.Add(rng, WdFieldType.wdFieldSequence, strSequenceName + " \* ARABIC ")
                'rng.Move(WdUnits.wdCharacter, 1)
                'fld.Update()
                '
                'rng.Text = vbTab
                'rng.Move(WdUnits.wdCharacter, 1)
                'rng.Text = strCaptionText
                'rngSelection = rng.Duplicate()

        End Select
        '
        Return rngSelection
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method expects the section (sect) to have two or three columns. It will return the column
    ''' number that the selection is in.. If it returns 0 or less than 0, then it could not determine a result
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function Plh_Columnsx2_FindColumnNumber(ByRef sect As Word.Section) As Integer
        Dim selPosH, selPosVert As Single
        Dim columnNumber As Integer
        Dim para As Word.Paragraph
        Dim rng As Word.Range
        Dim col1, col2, col3, col4 As Single
        '
        columnNumber = 0
        '
        Try
            para = Globals.ThisAddin.Application.Selection.Paragraphs.Item(1)
            rng = para.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            '
            'Find the position of the Selection point relative to the page edge (selPos)
            '
            selPosH = rng.Information(WdInformation.wdHorizontalPositionRelativeToPage)
            selPosVert = rng.Information(WdInformation.wdVerticalPositionRelativeToPage)
            '

            If sect.PageSetup.TextColumns.Count = 1 Then
                columnNumber = 1
            End If
            '
            If sect.PageSetup.TextColumns.Count = 2 Then
                'If we are not in column 1, then we must be in column 2
                columnNumber = 2
                '
                'The minus 5.0 s to adjust (emprically) for the effects of the offset Table)
                If selPosH > (sect.PageSetup.LeftMargin - 5) And selPosH <= sect.PageSetup.LeftMargin + sect.PageSetup.TextColumns.Item(1).Width Then columnNumber = 1
                '
                'If selPosH > sect.PageSetup.PageWidth - sect.PageSetup.RightMargin - sect.PageSetup.TextColumns.Item(2).Width And selPosH <= sect.PageSetup.PageWidth - sect.PageSetup.RightMargin Then
                'MsgBox("In column 2")
                'columnNumber = 2
                'End If
                '
            End If
            '
            If sect.PageSetup.TextColumns.Count = 3 Then
                '
                'Find the right hand boundaries of each column
                col1 = sect.PageSetup.LeftMargin + sect.PageSetup.TextColumns.Item(1).Width
                col2 = col1 + sect.PageSetup.TextColumns.Spacing + sect.PageSetup.TextColumns.Item(2).Width
                col3 = col2 + sect.PageSetup.TextColumns.Spacing + sect.PageSetup.TextColumns.Item(3).Width
                '
                columnNumber = 3
                If selPosH < col1 Then columnNumber = 1
                If selPosH > col1 And selPosH <= col2 Then columnNumber = 2

            End If
            '
            '
            If sect.PageSetup.TextColumns.Count = 4 Then
                '
                'Find the right hand boundaries of each column
                col1 = sect.PageSetup.LeftMargin + sect.PageSetup.TextColumns.Item(1).Width
                col2 = col1 + sect.PageSetup.TextColumns.Spacing + sect.PageSetup.TextColumns.Item(2).Width
                col3 = col2 + sect.PageSetup.TextColumns.Spacing + sect.PageSetup.TextColumns.Item(3).Width
                col4 = col3 + sect.PageSetup.TextColumns.Spacing + sect.PageSetup.TextColumns.Item(3).Width
                '
                columnNumber = 4
                If selPosH < col1 Then columnNumber = 1
                If selPosH > col1 And selPosH <= col2 Then columnNumber = 2
                If selPosH > col2 And selPosH <= col3 Then columnNumber = 3

            End If
            '

        Catch ex As Exception

        End Try


        Return columnNumber
    End Function
    '
    '
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub Plh_Table_FixPadding(ByRef tbl As Word.Table, topPadding As Single, bottomPadding As Single, leftPadding As Single, rightPadding As Single)
        tbl.Borders.Enable = False
        tbl.TopPadding = topPadding
        tbl.BottomPadding = bottomPadding
        tbl.LeftPadding = leftPadding
        tbl.RightPadding = rightPadding
        '
        tbl.Range.Style = Globals.ThisAddin.Application.ActiveDocument.Styles("Body Text")
        '
    End Sub
    '
    Public Sub Plh_Table_FixCellPadding(ByRef drCell As Word.Cell, topPadding As Single, bottomPadding As Single, leftPadding As Single, rightPadding As Single)
        drCell.TopPadding = topPadding
        drCell.BottomPadding = bottomPadding
        drCell.LeftPadding = leftPadding
        drCell.RightPadding = rightPadding
    End Sub
    '
    'This function checks for illegal insertion actions. It first makes certain that
    'the user is not trying to insert anything is the Cover Page, Contacts Page (Front and Back),
    'Table of Contents.. If so, we get an error message and it returns False. It
    'will then check to see if we are trying to insert in a Table. If so, then we get
    'an error message and it will return false
    Public Overridable Function Plh_is_OKToInsert(ByRef sect As Section, Optional doTableCheck As Boolean = True) As String
        Dim objInsertTest As New cInsertTestMgr()
        Dim strErrorMsg As String
        '
        strErrorMsg = objInsertTest.ins_is_OKToInsert(sect, doTableCheck)
        '

        Dim objCpMgr As New cCoverPageMgr()
        Dim objSectMgr As New cSectionMgr()
        Dim objTOCMgr As New cTOCMgr()
        Dim objRpt As New cReport()
        Dim strTagName As String
        Dim strRptMode As String
        '
        strErrorMsg = ""
        strRptMode = objRpt.Rpt_Mode_Get()
        strTagName = Me.tags_get_tagStyleName(sect)
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
        If tags_is_ContactsPage_Front(glb_get_wrdSect()) Then
            strErrorMsg = "Insertion In a Contacts Page Is Not supported"
            GoTo finis
        End If
        '
        If tags_is_ContactsPage_Back(glb_get_wrdSect()) Then
            strErrorMsg = "Insertion In a Contacts Page Is Not supported"
            GoTo finis
        End If
        '
        'Select Case strRptMode
        'Case objRpt.modeLong, objRpt.modeShort
        'Case objRpt.modeLongLandscape
        'If strTagName Like "tag_*" Then
        'strErrorMsg = "Insertion Of a 'PlaceHolder' in this page is not supported"
        'GoTo finis
        'End If
        'End Select
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
    '
    ''' <summary>
    ''' This method will return the Style that has the name 'strStyleName'. If there
    ''' is an error it will return the Normal style
    ''' </summary>
    ''' <param name="strStyleName"></param>
    ''' <returns></returns>
    Public Function xPlh_Style_GetSpecificStyle(strStyleName As String) As Word.Style
        Dim rslt As Word.Style
        '
        rslt = Nothing
        Try
            rslt = glb_get_wrdActiveDoc.Styles(strStyleName)
        Catch ex As Exception
            rslt = glb_get_wrdActiveDoc.Styles("Normal")
        End Try
        '
        Return rslt
    End Function
    '
    '
    ''' <summary>
    ''' This method will insert the Source information at the specific range, which
    ''' needs to be collapsed to a single point.. It will return a range that includes
    ''' all of the inserted paragraphs. The parameter strInsertType can take on the values
    ''' 'sourceOnly', 'sourceAndNote', 'note'
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <returns></returns>
    Public Function Plh_Insert_SourceAndNoteText(ByRef rng As Word.Range, Optional strInsertType As String = "sourceOnly") As Word.Range
        '
        rng = Me.objTblsMgr.tbl_insert_SourceAndNoteText(rng, strInsertType)
        '
        Return rng
        '
    End Function
    '
End Class

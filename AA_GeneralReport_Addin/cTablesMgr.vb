Imports Microsoft.Office.Interop.Word
Imports System.Windows.Forms
'Imports Microsoft.Office.Core
'Imports System.Windows.Forms.VisualStyles.VisualStyleElement.Window
Public Class cTablesMgr
    Inherits cGlobals
    '
    Public var_tbl_colourHeader As Long             'Table Header Colour
    Public var_tbl_colourUnits As Long              'Units row colour
    '
    Public var_tbl_tblHeaderStyle As Word.Style
    Public var_tbl_UnitsStyle As Word.Style
    Public var_tbl_TextBoldStyle As Word.Style
    Public var_tbl_TextStyle As Word.Style
    Public var_tbl_CaptionStyle As Word.Style
    '
    Public var_tbl_padding_Top As Single
    Public var_tbl_padding_Bottom As Single
    Public var_tbl_padding_Left As Single
    Public var_tbl_padding_Right As Single
    '
    Public var_tbl_myDoc As Word.Document
    Public var_tbl_TableStyleDefault As Word.Style
    Public var_tbl_TableStyleNoLines As Word.Style

    Public Sub New()
        MyBase.New()
        Dim objStylesMgr As New cStylesManager()
        Dim objTblStyles As New cTableStyles()
        '
        '
        var_tbl_myDoc = glb_get_wrdActiveDoc()
        '
        Me.var_tbl_colourHeader = Me._glb_colour_purple_Dark
        Me.var_tbl_colourUnits = Me._glb_colour_UnitsGrey
        '
        '

        '
        'Me.var_tbl_TextStyle = objStylesMgr.style_txt_getTableTextStyle(glb_get_wrdActiveDoc)
        'Me.var_tbl_TextStyle = var_tbl_myDoc.Styles.Item("Table text")
        'Me.var_tbl_tblHeaderStyle = var_tbl_myDoc.Styles.Item("Table column headings")
        'Me.var_tbl_UnitsStyle = var_tbl_myDoc.Styles.Item("Table units row")
        '
        'Looks for specific styles. Will add them if they are not there
        '
        Me.var_tbl_TextStyle = objStylesMgr.style_txt_getTableTextStyle(var_tbl_myDoc)
        Me.var_tbl_tblHeaderStyle = objStylesMgr.style_txt_getTableHeadingStyle(var_tbl_myDoc)
        Me.var_tbl_UnitsStyle = objStylesMgr.style_txt_getTableUnitsRowStyle(var_tbl_myDoc)
        '
        'objStylesMgr.style_txt_getTableCaptionStyle(var_tbl_myDoc)

        '
        Me.var_tbl_padding_Top = 0#
        Me.var_tbl_padding_Bottom = 2.0#
        Me.var_tbl_padding_Left = 0#
        Me.var_tbl_padding_Right = 0#
        '
        'var_tbl_TableStyleDefault = var_tbl_myDoc.Styles.Item("aac Table (Basic)")
        'var_tbl_TableStyleNoLines = var_tbl_myDoc.Styles.Item("aac Table (no lines)")
        '
        var_tbl_TableStyleDefault = objTblStyles.tblstyl_add_aacTableBasic(var_tbl_myDoc)
        var_tbl_TableStyleNoLines = objTblStyles.tblstyl_add_aacTableNoLines(var_tbl_myDoc)


    End Sub
    '
    Public Sub New(ByRef myDoc As Word.Document)
        MyBase.New()
        '
        var_tbl_myDoc = glb_get_wrdActiveDoc()
        '
        Me.var_tbl_colourHeader = Me._glb_colour_purple_Dark
        Me.var_tbl_colourUnits = Me._glb_colour_UnitsGrey
        '
        Me.var_tbl_tblHeaderStyle = myDoc.Styles.Item("Table column headings")
        Me.var_tbl_UnitsStyle = myDoc.Styles.Item("Table units row")
        Me.var_tbl_TextStyle = myDoc.Styles.Item("Table text")
        '
        Me.var_tbl_padding_Top = 0#
        Me.var_tbl_padding_Bottom = 2.0#
        Me.var_tbl_padding_Left = 0#
        Me.var_tbl_padding_Right = 0#
        '
        var_tbl_TableStyleDefault = var_tbl_myDoc.Styles.Item("aac Table (Basic)")

    End Sub
    '
    Public Sub tbl_apply_TableStyle(ByRef tbl As Word.Table)
        Try
            glb_tbl_apply_aacTableNoLinesStyle(tbl)
        Catch ex As Exception

        End Try
    End Sub
    '
    '
    ''' <summary>
    ''' This method will return the style in the first cell of the table (tbl)
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_get_tagStyle(ByRef tbl As Word.Table) As String
        Dim objTools As New cTools()
        Dim strStyleName As String
        '
        strStyleName = objTools.tools_tbls_getFirstCellStyleName(tbl)
        '
        Return strStyleName
        '
    End Function
    '
    ''' <summary>
    ''' This method will retrieve the caption for the generic AAC placeholders (Box, Figure, Recommendation). That is, for any placeholder that
    ''' has its caption included in the first row of the 'holding' table
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="para"></param>
    ''' <param name="doFullCaption"></param>
    ''' <returns></returns>
    Public Function tbl_getTblCaption_AACPlaceHolder(ByRef tbl As Word.Table, ByRef para As Word.Paragraph, Optional doFullCaption As Boolean = False) As String
        Dim strCaption As String
        Dim rng As Word.Range
        'Dim para As Word.Paragraph
        Dim tokens As String()
        '
        strCaption = ""

        Try
            rng = tbl.Range.Cells.Item(1).Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'rng.Move(WdUnits.wdParagraph, -1)
            para = rng.Paragraphs.Item(1)
            strCaption = para.Range.Text
            strCaption = Trim(strCaption)
            If Not doFullCaption Then
                tokens = strCaption.Split(vbTab)
                strCaption = tokens(0)
            End If
        Catch ex As Exception
            strCaption = ""
        End Try


        Return strCaption
    End Function
    '
    ''' <summary>
    ''' This method will return the caption that is just above the table (tbl). At the momen this means
    ''' AAC tables and not any other placeholder. But it will deal with the case where authors have
    ''' built their own tables and placed a caption above it.
    ''' table
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="para"></param>
    ''' <param name="doFullCaption"></param>
    ''' <returns></returns>
    Public Function tbl_getTblCaption_AACTable(ByRef tbl As Word.Table, ByRef para As Word.Paragraph, Optional doFullCaption As Boolean = False) As String
        Dim strCaption As String
        Dim rng As Word.Range
        'Dim para As Word.Paragraph
        Dim tokens As String()
        '
        strCaption = ""
        Try
            rng = tbl.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Move(WdUnits.wdParagraph, -1)
            para = rng.Paragraphs.Item(1)
            strCaption = para.Range.Text
            strCaption = Trim(strCaption)
            If Not doFullCaption Then
                tokens = strCaption.Split(vbTab)
                strCaption = tokens(0)
                strCaption = strCaption.Replace(vbCr, "")
            End If
        Catch ex As Exception
            strCaption = ""
        End Try


        Return strCaption

    End Function
    '
    ''' <summary>
    ''' This method will insert a caption at the top of the table tbl. It assumes that there is
    ''' an empty paragraph at the top of the Table. it will add an additional paragraph to ensure adequate spacing
    ''' Allowed values for strCaptionType are; 'Table_ES", 'Table', 'Table_AP' and 'Table_LT'
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_captions_Insert(ByRef tbl As Table, strCaptionType As String, Optional doSecondLineIndent As Boolean = True) As Word.Range
        Dim rng As Range
        Dim para, captionParagraph As Word.Paragraph
        Dim objChptPlhMgr As New cPlHBase()
        Dim captionStyle As Word.Style
        Dim indentSizeInPoints As Single
        Dim dr As Word.Row
        '
        captionStyle = tbl.Range.Document.Styles.Item(Me.glb_var_style_tblCaptionStyle)
        '
        '*** For encapsulated Table
        'dr = tbl.Rows.Add(tbl.Rows.Item(1))
        'dr.Range.Style = captionStyle
        'dr.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic
        'dr.Shading.BackgroundPatternColor = Me._glb_colour_UnitsGrey
        'dr.Shading.Texture = Word.WdTextureIndex.wdTextureNone
        'dr.Cells.Merge()
        'dr.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone
        'rng = dr.Range

        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'GoTo loop1
        '
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Move(WdUnits.wdParagraph, -1)
        rng.Paragraphs.Add(rng)
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        para = rng.Paragraphs.Item(1)
        para.Style = captionStyle.NameLocal
        rng = para.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        '***
loop1:
        '
        rng = objChptPlhMgr.Plh_Captions_InsertCaptions(strCaptionType, rng, True)
        '
        captionParagraph = rng.Paragraphs.Item(1)
        indentSizeInPoints = Me.var_glb_style_tblCaption_Line2_Indent
        '
        Try
            dr = tbl.Rows.Item(2)
            '
            captionParagraph.LeftIndent = dr.LeftIndent + indentSizeInPoints
            'captionParagraph.LeftIndent = 50.0

            captionParagraph.FirstLineIndent = -indentSizeInPoints
            '
        Catch ex As Exception

        End Try
        '
        Return rng
    End Function
    '
    ''' <summary>
    ''' This method will place an indented paragraph above the Table (tbl)
    ''' </summary>
    ''' <param name="captionParagraph"></param>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_captions_doIndent(ByRef tbl As Word.Table, ByRef captionParagraph As Word.Paragraph) As Word.Paragraph
        Dim indentSizeInPoints As Single
        Dim dr As Word.Row
        '
        Try
            indentSizeInPoints = Me.var_glb_style_tblCaption_Line2_Indent
            dr = tbl.Rows.Item(2)
            '
            captionParagraph.LeftIndent = dr.LeftIndent + indentSizeInPoints
            captionParagraph.FirstLineIndent = -indentSizeInPoints
            '
        Catch ex As Exception

        End Try
        '
        Return captionParagraph
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will indent the caption paragraph above the table (tbl) to match the
    ''' indent of the table. If there is no Caption Paragraph it will do nothing
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_captions_doIndent(ByRef tbl As Word.Table) As Word.Paragraph
        Dim indentSizeInPoints As Single
        Dim theStyle As Word.Style
        Dim para As Word.Paragraph
        Dim dr As Word.Row
        Dim rng As Word.Range
        '
        para = Nothing
        '
        If Not IsNothing(tbl) Then
            rng = tbl.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Move(WdUnits.wdParagraph, -1)
            para = rng.Paragraphs.Item(1)
            '
            Try
                theStyle = para.Range.Style
                If theStyle.NameLocal = Me.glb_var_style_tblCaptionStyle Then
                    '
                    indentSizeInPoints = Me.var_glb_style_tblCaption_Line2_Indent
                    dr = tbl.Rows.Item(2)
                    '
                    para.LeftIndent = dr.LeftIndent + indentSizeInPoints
                    para.FirstLineIndent = -indentSizeInPoints
                    '
                End If
            Catch ex As Exception

            End Try

        End If
        '
        Return para
        '
    End Function
    '
    '
    Public Sub tbl_colour_set_colourOfColumn(ByRef drCol As Word.Column, fillColour As Long)
        Dim drCell As Cell
        '
        For Each drCell In drCol.Cells
            drCell.Shading.BackgroundPatternColor = fillColour
        Next drCell
        '
    End Sub
    '
    ''' <summary>
    ''' This method will fill the cells drCells with the colour fillColour
    ''' </summary>
    ''' <param name="drCells"></param>
    ''' <param name="fillColour"></param>
    ''' <returns></returns>
    Public Function tbl_colour_set_colourOfCells(ByRef drCells As Word.Cells, fillColour As Long) As Boolean
        Dim drCell As Cell
        Dim rslt As Boolean
        '
        rslt = False
        If drCells.Count <> 0 Then rslt = True
        '
        For Each drCell In drCells
            drCell.Shading.Texture = Word.WdTextureIndex.wdTextureSolid
            drCell.Shading.BackgroundPatternColor = fillColour
            drCell.Shading.ForegroundPatternColor = fillColour
        Next drCell
        '
        Return rslt
        '
    End Function
    '

    '
    ''' <summary>
    ''' This method will fill the selected table cells with the colour fillCOlour
    ''' </summary>
    ''' <param name="fillColour"></param>
    ''' <returns></returns>
    Public Function tbl_colour_set_colourOfCells(fillColour As Long) As Boolean
        Dim rslt As Boolean
        Dim drCells As Word.Cells
        '
        drCells = glb_get_wrdSel.Cells
        rslt = False
        If drCells.Count <> 0 Then rslt = True
        '
        drCells.Shading.Texture = Word.WdTextureIndex.wdTextureNone
        drCells.Shading.BackgroundPatternColor = fillColour
        'drCells.Shading.ForegroundPatternColor = WdColor.
        '
        Return rslt
        '
    End Function
    '

    '
    Public Function tbl_colour_set_colourOfCells(ByRef btn As System.Windows.Forms.ToolStripButton) As Boolean
        Dim rslt As Boolean
        Dim drCells As Word.Cells
        Dim btnColor As System.Drawing.Color
        '
        rslt = False
        drCells = glb_get_wrdSel.Cells
        '
        Try
            btnColor = btn.BackColor
            drCells = glb_get_wrdSel.Cells
            If drCells.Count <> 0 Then rslt = True
            '
            'Transparent settings in frm_colorPicker translate to white.. So what we do here is
            'look for the tooltip text and if it says transparent we handle it differently
            If btn.ToolTipText <> "Transparent" Then
                drCells.Shading.Texture = Word.WdTextureIndex.wdTextureNone
                drCells.Shading.BackgroundPatternColor = RGB(btnColor.R, btnColor.G, btnColor.B)
            Else
                drCells.Shading.Texture = Word.WdTextureIndex.wdTextureNone
                drCells.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic
                drCells.Shading.ForegroundPatternColor = WdColor.wdColorAutomatic
            End If
            rslt = True
            '
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
        '
    End Function

    '
    '
    Public Function tbl_colour_set_colourOfCellsToNone(ByRef drCells As Word.Cells) As Boolean
        Dim drCell As Cell
        Dim rslt As Boolean
        '
        rslt = False
        If drCells.Count <> 0 Then rslt = True
        '
        For Each drCell In drCells
            drCell.Shading.Texture = Word.WdTextureIndex.wdTextureNone
            drCell.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic
            drCell.Shading.ForegroundPatternColor = WdColor.wdColorAutomatic
        Next drCell
        '
        Return rslt
        '
    End Function
    '
    '
    Public Function tbl_colour_set_colourOfCellToNone(ByRef drCell As Word.Cell) As Boolean
        Dim rslt As Boolean
        '
        rslt = False
        '
        Try
            drCell.Shading.Texture = Word.WdTextureIndex.wdTextureNone
            drCell.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic
            drCell.Shading.ForegroundPatternColor = WdColor.wdColorAutomatic
        Catch ex As Exception
            rslt = False
        End Try
        '
        '
        Return rslt
        '
    End Function
    '

    '
    '
    Public Sub tbl_colour_set_colourOfRow(ByRef dr As Word.Row, fillColour As Long)
        '
        Me.tbl_colour_set_colourOfCells(dr.Range.Cells, fillColour)
        'dr.Shading.Texture = Word.WdTextureIndex.wdTextureNone
        'dr.Shading.BackgroundPatternColor = fillColour
        'dr.Shading.ForegroundPatternColor = fillColour
        '
    End Sub
    '
    ''' <summary>
    ''' This method expects the selection to be in one or more cells. It will get the row.index
    ''' of the first cell and colour all cells with the same row index.. This function does NOT depend on the
    ''' table being regular
    ''' </summary>
    ''' <param name="fillColour"></param>
    Public Sub tbl_colour_set_colourOfRow(ByRef tbl As Word.Table, rowIndex As Integer, fillColour As Long)
        Dim drCell As Word.Cell
        '
        Try
            For j = 1 To tbl.Range.Cells.Count
                drCell = tbl.Range.Cells.Item(j)
                If drCell.RowIndex > rowIndex Then Exit For
                If drCell.RowIndex = rowIndex Then
                    drCell.Shading.Texture = Word.WdTextureIndex.wdTextureNone
                    drCell.Shading.BackgroundPatternColor = fillColour
                    drCell.Shading.ForegroundPatternColor = fillColour
                End If
            Next
        Catch ex As Exception

        End Try
        '
    End Sub
    '

    '
    '
    Public Sub tbl_colour_set_colourOfRowToNone(ByRef dr As Word.Row)
        '
        Me.tbl_colour_set_colourOfCellsToNone(dr.Range.Cells)
        'dr.Shading.Texture = Word.WdTextureIndex.wdTextureNone
        'dr.Shading.BackgroundPatternColor = fillColour
        'dr.Shading.ForegroundPatternColor = fillColour
        '
    End Sub
    '

    '
    '
    Public Sub tbl_delete_table()
        'This method will delete the currently selected table
        Dim rng As Range
        Dim tbl As Table
        '
        rng = Globals.ThisAddin.Application.Selection.Range
        If rng.Tables.Count = 0 Then
            MsgBox("Please make certain that you cursor is in a table")
            Exit Sub
        End If
        '
        For Each tbl In rng.Tables
            tbl.Delete()
        Next tbl
    End Sub
    '
    Public Function tbl_convert_aacToNoOutDent_ConvertOrResize(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim strSelect As String
        Dim myDoc As Word.Document
        '
        rslt = True
        strSelect = ""
        myDoc = tbl.Range.Document
        '
        If Me.tbl_is_LegacyAATable(tbl) Then strSelect = "isAACLegacyTable"
        '
        Select Case strSelect
            Case "isAACLegacyTable"
                'We will convert in place
                rslt = Me.tbl_convert_aacToNoOutDent(tbl)
                'rslt = Me.tbl_convert_aacToNoOutDent_TemplateVersion(tbl)
                '

            Case "isSquareEdgeTable"
                'Will widen a narrow table an narrow a wide table
                Try
                    'Assume regular by row 
                    'Resize
                Catch ex As Exception
                    Try
                        'Not regular by row, so treat the tables ina way where we can't acces the rows
                        'Resize
                    Catch ex2 As Exception
                        'MsgBox("Error")

                    End Try
                End Try
        End Select



        Return rslt
    End Function
    '
    ''' <summary>
    ''' This version makes use of the embedded Tabel Style
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_convert_aacToNoOutDent_TemplateVersion(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        'Dim dr As Word.Row
        'Dim drCell As Word.Cell
        Dim tblWidth As Single
        Dim sect As Word.Section
        'Dim hf_tbl As Word.Table
        Dim myDoc As Word.Document

        '
        sect = tbl.Range.Sections.Item(1)
        myDoc = tbl.Range.Document
        rslt = False
        '
        Try

            'We first start by assuming that the table is regular by rows.. If it isn't then we will got to the
            'catch... Which assumes nothing about rows
            'dr = tbl.Range.Rows.Item(1)
            '
            '
            '
            tblWidth = Me.glb_tbls_getTableWidth(tbl)

            '
            'Check the difference, allowing for precision creep.. STandard leftPadding is 22.7
            If tblWidth <= glb_get_widthBetweenMargins(sect) Then
                tbl.Style = myDoc.Styles("aac Table (Basic)")
                tbl.ApplyStyleHeadingRows = False

                Me.tbl_setWidth_ToStandard(tbl)
                'tbl.Rows.SetLeftIndent(0.0, WdRulerStyle.wdAdjustProportional)
                'tbl.PreferredWidth = glb_get_widthBetweenMargins(sect)
                '
            Else
                Me.tbl_setWidth_toWide(tbl)
                tbl.Style = myDoc.Styles("aac Table (Basic)")
                tbl.ApplyStyleHeadingRows = False

                'tbl.Rows.SetLeftIndent(hf_tbl.Rows.Item(1).LeftIndent, WdRulerStyle.wdAdjustProportional)
                'tbl.PreferredWidth = hf_tbl.PreferredWidth
                '
            End If
            rslt = True
        Catch ex As Exception
            rslt = False
        End Try
        '
        '
        Return rslt
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will convert all of the aac outdented Tables in the specified range, rng to the
    ''' new non outdent structure. It successively calls the method 'tbl_convert_aacToNoOutDent(tbl)'.
    ''' The table is checked to see that it is a standard outdented aac table. Typically used to convert
    ''' Tables in a selection
    ''' </summary>
    Public Sub tbl_convert_aacTablesNoOutDent(ByRef rng As Word.Range)
        Dim tbl As Word.Table

        Dim objGlobals As New cGlobals()
        Dim objTools As New cTools()
        '
        '
        Try
            '
            For Each tbl In rng.Tables
                'Me.wcag_convert_aacTableToWCAG(tbl)
                If Me.glb_tbl_isLegacyAATable(tbl) Then
                    Me.tbl_convert_aacToNoOutDent(tbl)
                End If
                '
            Next
        Catch ex As Exception

        End Try
    End Sub
    '


    '
    ''' <summary>
    ''' This method will convert an aac outldented table to a 'non' outdented, or square Table.
    ''' It does so by looking for outdent differences between the first and last row. If there is a
    ''' difference, then the first row is adjusted and then the entire Table is adjusted to the
    ''' outdent of the last row to ensure that the columns all align
    ''' '
    ''' Verified 20231127
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_convert_aacToNoOutDent(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim leftIndent, leftPadding, tblWidth, marginWidth As Single
        Dim sect As Word.Section
        Dim hf_tbl As Word.Table
        Dim myDoc As Word.Document
        Dim objCaptions As New cCaptionManager()
        Dim para As Word.Paragraph
        Dim objTblStyles As New cTableStyles()

        '
        sect = tbl.Range.Sections.Item(1)
        myDoc = tbl.Range.Document
        rslt = True
        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        hf_tbl = Nothing
        '
        Me.glb_hfs_getHeaderTable(hf_tbl)

        leftIndent = 0
        '
        Try
            'We first start by assuming that the table is regular by rows.. If it isn't then we will got to the
            'catch... Which assumes nothing about rows
            dr = tbl.Range.Rows.Item(1)
            leftPadding = tbl.Range.Cells.Item(1).LeftPadding
            '
            tblWidth = Me.tbl_get_tableBodyWidth(tbl)
            '
            'Check the difference, allowing for precision creep.. STandard leftPadding is 22.7
            If leftPadding > 15 Then
                '

                drCell = tbl.Range.Cells.Item(1)
                drCell.LeftPadding = 0.0
                drCell.Width = drCell.Width - leftPadding
                '
                dr.SetLeftIndent(dr.LeftIndent + leftPadding, WdRulerStyle.wdAdjustNone)
                tbl.Rows.SetLeftIndent(dr.LeftIndent, WdRulerStyle.wdAdjustProportional)
                '
                'Shift Left To bring the rows together, then shift back.. This way the table should stay in place
                tbl.Rows.SetLeftIndent(tbl.Rows.LeftIndent - leftPadding, WdRulerStyle.wdAdjustNone)
                tbl.Rows.SetLeftIndent(tbl.Rows.LeftIndent + leftPadding, WdRulerStyle.wdAdjustNone)
                '
                If tblWidth <= glb_get_widthBetweenMargins(sect) Then
                    '
                    '*** These two statements aren't strcitly necessary and may be removed for the Addin version
                    'Me.glb_tbl_apply_aacTableNoLinesStyle(tbl)
                    Me.glb_tbl_apply_aacTableBasicStyle(tbl)
                    'dr = tbl.Rows.Item(1)
                    'dr.Borders.Item(WdBorderType.wdBorderBottom) = WdLineStyle.wdLineStyleNone

                    'tbl.Style = myDoc.Styles("aac Table (Basic)"
                    'tbl.ApplyStyleHeadingRows = True
                    '
                    '
                    Me.tbl_setWidth_ToStandard(tbl)
                    'tbl.Rows.SetLeftIndent(0.0, WdRulerStyle.wdAdjustProportional)
                    'tbl.PreferredWidth = glb_get_widthBetweenMargins(sect)
                    '
                Else
                    '
                    '*** These two statements aren't strcitly necessary and may be removed for the Addin version
                    'Me.glb_tbl_apply_aacTableNoLinesStyle(tbl)
                    Me.glb_tbl_apply_aacTableBasicStyle(tbl)
                    'tbl.Style = myDoc.Styles("aac Table (Basic)"
                    'tbl.ApplyStyleHeadingRows = True
                    '
                    Me.tbl_setWidth_toWide(tbl)
                    '
                    'tbl.Rows.SetLeftIndent(hf_tbl.Rows.Item(1).LeftIndent, WdRulerStyle.wdAdjustProportional)
                    'tbl.PreferredWidth = hf_tbl.PreferredWidth
                    '
                End If
                '
                tbl.ApplyStyleHeadingRows = True
            Else
                If leftPadding >= 0 And leftPadding <= 1 Then
                    'We have a standard table that somehow got here
                    'tblWidth = 0.0
                    'For Each drCell In tbl.Rows.Item(1).Cells
                    'tblWidth = tblWidth + drCell.Width
                    'Next
                    tblWidth = Me.glb_tbls_getTableWidth(tbl)
                    '
                    If tblWidth <= glb_get_widthBetweenMargins(sect) Then
                        tbl.PreferredWidth = glb_get_widthBetweenMargins(sect)
                        tbl.Rows.SetLeftIndent(0.0, WdRulerStyle.wdAdjustProportional)

                    Else
                        tbl.PreferredWidth = hf_tbl.PreferredWidth
                        tbl.Rows.SetLeftIndent(hf_tbl.Rows.Item(1).LeftIndent, WdRulerStyle.wdAdjustProportional)
                    End If

                End If
                '
                tbl.ApplyStyleHeadingRows = True
                '
            End If

        Catch ex As Exception
            Try
                drCell = tbl.Range.Cells.Item(1)
                leftPadding = drCell.LeftPadding
                '
                drCell.LeftPadding = 0.0
                drCell.SetWidth(drCell.Width - leftPadding, WdRulerStyle.wdAdjustNone)
                '
                leftIndent = tbl.Rows.LeftIndent
                '
                tblWidth = glb_tbls_getTableWidth(tbl)
                marginWidth = glb_get_widthBetweenMargins(sect)
                '
                '*** tbl.PreferredWidth is not defined for the multi merged cell table (at this point). It is is et to 9999999
                tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
                '
                If tblWidth <= marginWidth Then
                    '
                    '*** These two statements aren't strcitly necessary and may be removed for the Addin version
                    'Me.glb_tbl_apply_aacTableNoLinesStyle(tbl)
                    'Me.glb_tbl_apply_aacTableNoLinesStyle(tbl)
                    Me.glb_tbl_apply_aacTableBasicStyle(tbl)

                    'tbl.Style = myDoc.Styles("aac Table (Basic)"
                    'tbl.ApplyStyleHeadingRows = True
                    '
                    'Me.tbl_set_ToStandard(tbl)
                    tbl.Rows.SetLeftIndent(0.0, WdRulerStyle.wdAdjustNone)
                    tbl.PreferredWidth = marginWidth
                    para = objCaptions.cpt_indent_CaptionParagraph(tbl, 0.0, True)

                Else
                    '
                    '*** These two statements aren't strcitly necessary and may be removed for the Addin version
                    'Me.glb_tbl_apply_aacTableNoLinesStyle(tbl)
                    'Me.glb_tbl_apply_aacTableNoLinesStyle(tbl)
                    Me.glb_tbl_apply_aacTableBasicStyle(tbl)
                    'tbl.Style = myDoc.Styles("aac Table (Basic)"
                    'tbl.ApplyStyleHeadingRows = True
                    '
                    'Me.tbl_set_ToWide(tbl)

                    tbl.Rows.SetLeftIndent(hf_tbl.Rows.Item(1).LeftIndent, WdRulerStyle.wdAdjustNone)
                    tbl.PreferredWidth = hf_tbl.PreferredWidth
                    para = objCaptions.cpt_indent_CaptionParagraph(tbl, hf_tbl.Rows.LeftIndent, True)

                End If
                '
                tbl.ApplyStyleHeadingRows = True
                '
            Catch ex2 As Exception
                MsgBox("Could not handle the table")
                rslt = False
                '
            End Try
        End Try
        '
        '
        Return rslt
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will evenly chnage the width of all the columns in a Table
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub tbl_width_Changex(ByRef tbl As Word.Table, newTableWidth As Single)
        Dim drHeader, drLastBodyRow As Word.Row
        Dim drCol As Word.Column
        Dim numColumns, j As Integer
        Dim tblSource As Word.Table
        Dim oldRange As Word.Range
        Dim drHeaderCell As Word.Cell
        Dim headerOffSet, leftIndent, tblWidth, deltaWidth, deltatWidthPerColumn As Single
        Dim strFlag As String

        Try
            'Just in case we were passed a Table that was Nothing
            'Get the Header Row, then get the First cell offset... This is always non zero in
            'an AAC Table
            drHeader = Me.tbl_headerRow_Get(tbl)
            Try
                drHeaderCell = drHeader.Cells.Item(1)
                '
                'Make certain that the Table is regular, whilst recording the offsets
                'used in the Table
                '
                headerOffSet = drHeaderCell.LeftPadding
                tbl = Me.HeaderRowAndTable_SetRegularFormatting(headerOffSet, drHeader)
                leftIndent = drHeader.LeftIndent
                strFlag = Me.tbl_get_TableType(tbl)
                '
            Catch ex As Exception
                MsgBox("Table Error. Are you sure your cursor in somewhere in the body of a standard AAC Table?")
                Exit Sub
            End Try
            '
            oldRange = Me.glb_get_wrdSel.Range
            '
            'Find the last Body Row and split off the Caption
            drLastBodyRow = Me.tbl_get_LastBodyRow_ForMCWC(tbl)
            drLastBodyRow.Select()
            tblSource = tbl.Split(drLastBodyRow.Next)
            tblWidth = Me.glb_tbls_getTableWidth(tbl)
            '
            numColumns = tbl.Columns.Count
            '
            deltaWidth = newTableWidth - tblWidth
            deltatWidthPerColumn = deltaWidth / numColumns
            '
            For j = 1 To tbl.Columns.Count
                drCol = tbl.Columns.Item(j)
                drCol.Width = drCol.Width + deltatWidthPerColumn
            Next
            '
            tblSource.Columns.Item(1).Width = tblSource.Columns.Item(1).Width + deltaWidth
            '
            'Undo the Table Split, then reset the Header Row to have the original offset
            Me.tbl_delete_ParaAtEndOfTable(tbl)
            drHeader = tbl.Rows.Item(1)
            '
            Me.tbl_headerRow_Reset(headerOffSet, leftIndent, drHeader)
            '
            oldRange.Select()
            '
        Catch ex As Exception

        End Try
    End Sub
    '
    Public Sub tbl_width_Change(ByRef tbl As Word.Table, newTableWidth As Single, Optional leftIndent As Single = 0.0)
        Dim objPlhBase As New cPlHBase()
        '
        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        tbl.Rows.SetLeftIndent(leftIndent, WdRulerStyle.wdAdjustProportional)
        tbl.PreferredWidth = newTableWidth
        '
        Try
            objPlhBase.Plh_scale_FigureImageShape(tbl)
        Catch ex As Exception

        End Try

    End Sub

    ''' <summary>
    ''' This method will toggle the table width to either sit between the margins (standard size),
    ''' or to sit between the edges of the header table (wide size). it will return true if all
    ''' went according to plan
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_toggle_tblWidth(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean = False
        'Dim objPlhBase As New cPlHBase()
        '
        '
        'If Me.tbl_width_isWide(tbl) Then strToggle = "isWide"
        'If Me.tbl_width_isStandard(tbl) Then strToggle = "isStd"
        '
        'strToggle = Me.tbl_width_isWide(tbl)
        'MsgBox(strToggle)
        Try
            Select Case Me.tbl_width_isWide(tbl)
                Case "isWide"
                    rslt = Me.tbl_setWidth_ToStandard(tbl)
                Case "isStd"
                    rslt = Me.tbl_setWidth_toWide(tbl)
                    'objPlhBase.Plh_scale_FigureImageShape(tbl)
                Case Else
                    rslt = Me.tbl_setWidth_ToStandard(tbl)
                    tbl.Rows.LeftIndent = 0.0
            End Select
            '
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
    End Function
    '
    ''' <summary>
    ''' This method will return true if the body width of the table is greater than the 
    ''' width between the margins.. Because we are comparing numbers dimensioned as Single
    ''' and not integers we need to allow for precision errors. So comparisons are done
    ''' within a specified tolerance
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_width_isWide(ByRef tbl As Word.Table) As String
        Dim rslt As String
        Dim tblWidth, marginWidth, headerWidth As Single
        Dim sect As Word.Section
        Dim tolerance As Single
        '
        tolerance = 0.1
        rslt = ""
        sect = tbl.Range.Sections.Item(1)
        tblWidth = Me.glb_tbls_getTableWidth(tbl)
        marginWidth = Me.glb_get_widthBetweenMargins(sect)
        headerWidth = Me.glb_hfs_getHeaderTableWidth(sect, "primary")
        '
        Try
            If tblWidth >= (marginWidth - tolerance) And tblWidth <= (marginWidth + tolerance) Then rslt = "isStd"
            If tblWidth < (marginWidth - tolerance) Then rslt = "isShort"
            If tblWidth > (marginWidth + tolerance) And tblWidth <= (headerWidth + tolerance) Then rslt = "isWide"

            'rslt = "isStd"
            'Else
            'rslt = "isWide"
            'End If
            '
            'MsgBox("Table is = " + rslt + vbCrLf + "tblWidth = " + tblWidth.ToString() + vbCrLf + "marginWIdth = " + marginWidth.ToString())
        Catch ex As Exception
            rslt = ""
        End Try
        '
        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method will return true if the body width of the table is less than or equal
    ''' to the width between the margins
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_width_isStandard(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim tblWidth, marginWidth As Single
        Dim sect As Word.Section
        '
        rslt = False
        sect = tbl.Range.Sections.Item(1)
        tblWidth = Me.glb_tbls_getTableWidth(tbl)
        marginWidth = Me.glb_get_widthBetweenMargins(sect)
        '
        Try
            If tblWidth <= marginWidth Then
                rslt = True
            End If

        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
    End Function
    '
    '
    ''' <summary>
    ''' This function will set the table (tbl) to its wide setting. Wide is defined as matching the width of the
    ''' Header Table.. If all is OK it will return true, otherwise it will return false.
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_setWidth_toWide(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim sect As Word.Section
        Dim hf_tbl As Word.Table
        Dim objWCAgMgr As New cWCAGMgr()
        Dim objCaptions As New cCaptionManager()
        Dim objPlhBase As New cPlHBase()
        Dim objFloatMgr As New cPlHFloatingMgr()
        Dim para As Word.Paragraph
        Dim paraLeftIndent As Single
        '
        hf_tbl = Nothing
        para = Nothing
        sect = tbl.Range.Sections.Item(1)
        rslt = True
        paraLeftIndent = 0.0
        '
        '
        Try
            Me.glb_hfs_getHeaderTable(hf_tbl)
            '
            If Me.tbl_is_LegacyAATable(tbl) Then Me.tbl_convert_aacToNoOutDent(tbl)
            '
            If Me.tbl_is_Floating(tbl) Then
                objFloatMgr.plHFLoat_setWidth_toWide(tbl)
                paraLeftIndent = 0.0
            Else
                tbl.Rows.SetLeftIndent(hf_tbl.Rows.Item(1).LeftIndent, WdRulerStyle.wdAdjustProportional)
                tbl.PreferredWidth = hf_tbl.PreferredWidth
                paraLeftIndent = hf_tbl.Rows.Item(1).LeftIndent
                '
            End If

            para = objCaptions.cpt_indent_CaptionParagraph(tbl, paraLeftIndent, True)
            para = objCaptions.cpt_indent_SourceNoteParagraph(tbl, paraLeftIndent, True)
            '
            objPlhBase.Plh_scale_FigureImageShape(tbl)
            '
        Catch ex As Exception
            rslt = False
        End Try
        '
        '
        Return rslt
    End Function
    '

    '
    ''' <summary>
    ''' This function will set the table (tbl) to its standard setting. Wide is defined as matching the width of the
    ''' Header Table.. If all is OK it will return true, otherwise it will return false.
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_setWidth_ToStandard(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim sect As Word.Section
        Dim hf_tbl As Word.Table
        Dim objWCAgMgr As New cWCAGMgr()
        Dim objCaptions As New cCaptionManager()
        Dim objPlhBase As New cPlHBase()
        Dim objFloatMgr As New cPlHFloatingMgr()
        Dim para As Word.Paragraph
        Dim paraLeftIndent, cellPadding As Single
        '
        hf_tbl = Nothing
        para = Nothing
        sect = tbl.Range.Sections.Item(1)
        rslt = True
        cellPadding = tbl.Range.Cells.Item(1).LeftPadding
        '
        paraLeftIndent = 0.0
        '
        Try
            Me.glb_hfs_getHeaderTable(hf_tbl)
            '
            If Me.tbl_is_LegacyAATable(tbl) Then Me.tbl_convert_aacToNoOutDent(tbl)
            '
            If Me.tbl_is_Floating(tbl) Then
                objFloatMgr.plHFLoat_setWidth_toStandard(tbl)
                paraLeftIndent = 0.0
            Else
                tbl.Rows.SetLeftIndent(0.0, WdRulerStyle.wdAdjustProportional)
                tbl.PreferredWidth = glb_get_widthBetweenMargins(sect)
                paraLeftIndent = 0.0
            End If

            '
            para = objCaptions.cpt_indent_CaptionParagraph(tbl, paraLeftIndent, True)
            para = objCaptions.cpt_indent_SourceNoteParagraph(tbl, paraLeftIndent, True)
            '
            objPlhBase.Plh_scale_FigureImageShape(tbl)
            '
        Catch ex As Exception
            rslt = False
        End Try
        '
        '
        Return rslt
    End Function
    '



    '
    ''' <summary>
    ''' This function will return the Header Row of the Table. At the moment it assumes
    ''' that this is always the first row
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_headerRow_Get(ByRef tbl As Word.Table) As Word.Row
        Try
            tbl_headerRow_Get = tbl.Rows.Item(1)
        Catch ex As Exception
            tbl_headerRow_Get = Nothing
        End Try
    End Function
    '
    '
    Public Function tbl_delete_ParaAtEndOfTable(ByRef tbl As Word.Table) As Word.Range
        Dim rng As Word.Range
        '
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng.Select()
        rng.Expand(WdUnits.wdParagraph)
        rng.Delete()
        '
        Return rng
        '
    End Function
    '
    ''' <summary>
    ''' This method will delete the range after the table that contains all Source and
    ''' Note styled paragraphs
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub tbl_delete_ParasSourceAndNoteAtEndOfTable(ByRef tbl As Word.Table)
        Dim rng As Word.Range
        '
        rng = Me.tbl_get_RangeOfParasSourceAndNoteAtEndOfTable(tbl)
        rng.Delete()
        '
        Return
        '
    End Sub
    '
    Public Function tbl_get_RangeOfParasSourceAndNoteAtEndOfTable(ByRef tbl As Word.Table) As Word.Range
        Dim strRslt As String
        Dim myDoc As Word.Document
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        Dim stylSrc, stylNote, paraStyle As Word.Style
        Dim j As Integer
        '
        myDoc = tbl.Range.Document()
        strRslt = ""
        stylSrc = myDoc.Styles.Item("Source")
        stylNote = myDoc.Styles.Item("Note")
        '
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
        'Set para to the first paragraph after the table
        para = rng.Paragraphs.Item(1)
        paraStyle = para.Range.Style
        rng = para.Range
        If paraStyle.NameLocal = stylSrc.NameLocal Or paraStyle.NameLocal = stylNote.NameLocal Then
            strRslt = strRslt + para.Range.Text
            'rng.MoveEnd(WdUnits.wdParagraph, 1)
            '
            For j = 2 To 10
                para = para.Next
                paraStyle = para.Range.Style
                If paraStyle.NameLocal = stylSrc.NameLocal Or paraStyle.NameLocal = stylNote.NameLocal Then
                    strRslt = strRslt + para.Range.Text
                    rng.MoveEnd(WdUnits.wdParagraph, 1)
                Else
                    Exit For
                End If
            Next

        Else
            'The first para is not a Source or Caption/Note style
            rng = Nothing
            '
        End If
        '
        Return rng
        '
    End Function


    '
    ''' <summary>
    ''' This method will delete the paragraph at the top of the Table
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_delete_ParaAtTopOfTable(ByRef tbl As Word.Table) As Word.Range
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        '
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Move(WdUnits.wdCharacter, -1)
        para = rng.Paragraphs.Item(1)
        para.Range.Delete()
        '
        Return rng
        '
    End Function



    Public Sub tbl_headerRow_Reset(headerOffSet As Single, LeftIndent As Single, ByRef drheader As Word.Row)
        'Remember headerOffset is a positive value and 
        'LeftIndent Is a negative value
        '
        Dim drCell As Word.Cell
        '
        For i = 1 To drheader.Cells.Count
            drCell = drheader.Cells.Item(i)
            drCell.BottomPadding = 0.0
        Next
        '
        drheader.LeftIndent = LeftIndent - headerOffSet
        drheader.Cells(1).LeftPadding = headerOffSet
        'drheader.LeftIndent = -headerOffSet + LeftIndent
        'drheader.LeftIndent = -LeftIndent

        drheader.Cells(1).Width = drheader.Cells(1).Width + headerOffSet
        '
    End Sub
    '
    ''' <summary>
    ''' This method will add a Row to the end of the Table (tbl). The return object
    ''' is the added row
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_rows_AddToEndOfTable(ByRef tbl As Word.Table) As Word.Row
        Dim rng As Word.Range
        Dim dr As Word.Row
        Dim drCellLast As Word.Cell
        Dim rowIdx, numCells As Integer
        '
        numCells = tbl.Range.Cells.Count
        drCellLast = tbl.Range.Cells.Item(numCells)
        rowIdx = drCellLast.RowIndex
        Me.tbl_get_RowCells(rowIdx, tbl)
        '
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        dr = tbl.Rows.Add(rng)
        '
        Return dr
        '
    End Function
    '
    ''' <summary>
    ''' This method will add a Row to the end of the Table (tbl). The return object
    ''' is the added row
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_rows_AddToEndOfTable_As_Cells(ByRef tbl As Word.Table) As Word.Range
        Dim rngOfCells, rngLastCell As Word.Range
        Dim drCellLast As Word.Cell
        Dim rowIdx, numCells As Integer
        '
        numCells = tbl.Range.Cells.Count
        drCellLast = tbl.Range.Cells.Item(numCells)
        rowIdx = drCellLast.RowIndex
        '
        drCellLast = tbl.Range.Cells.Item(numCells)
        rngLastCell = drCellLast.Range
        rngLastCell.Collapse(WdCollapseDirection.wdCollapseStart)
        rngLastCell.Select()
        '
        glb_get_wrdSel.InsertRowsBelow(1)
        '
        numCells = tbl.Range.Cells.Count
        drCellLast = tbl.Range.Cells.Item(numCells)
        rowIdx = drCellLast.RowIndex
        rngOfCells = Me.tbl_get_RowCells(rowIdx, tbl)
        '
        Return rngOfCells
        '
    End Function
    '

    '
    Public Function tbl_add_Column(strLeftRight As String, CursorPosition As Integer, tblWidth As Single, ByRef tbl As Word.Table) As Word.Column
        Dim drCol As Word.Column
        Dim strMsg As String
        '
        strMsg = "This function will only work on tables that " + vbCrLf
        strMsg = strMsg + "have a consistent column structure from top to bottom "
        '
        drCol = Nothing
        Try
            drCol = tbl.Columns.Item(CursorPosition)
            '
            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
            '
            'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
            Select Case strLeftRight
                Case "left"
                    '
                    drCol = tbl.Columns.Add(tbl.Columns.Item(CursorPosition))
                '
                Case "right"
                    'To insert to the right we move to next column and insert to
                    'it's left. If the selected column is the last column then we
                    'just add a column to the end of the Table
                    '
                    If drCol.Index = tbl.Columns.Last.Index Then
                        drCol = tbl.Columns.Add()
                    Else
                        drCol = drCol.Next
                        drCol = tbl.Columns.Add(drCol)

                    End If
            End Select

            '
            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
            tbl.PreferredWidth = tblWidth
            '
            'tbl.Columns.DistributeWidth()
            '
            'For single column tables the internal vertical border appears for no apparent
            'reason, so we have to ensure that it is not there
            tbl.Borders.Item(WdBorderType.wdBorderVertical).LineStyle = WdLineStyle.wdLineStyleNone
            '
        Catch ex As Exception
            MsgBox(strMsg)
        End Try

        '
        Return drCol
    End Function
    '
    '
    Public Function tbl_get_TableType(ByRef tbl As Word.Table) As String
        Dim numCells As Integer
        Dim strFlag As String
        Dim IsCaptionStyle As Boolean
        Dim dr As Word.Row
        '
        strFlag = ""
        IsCaptionStyle = False
        '
        'Start at the last row and run up the Table. If the number of cells per row
        'chnages to somethog other than 1 then we have found the body of the Table
        'if it gets to row one and its still 1 cell wide then we have a one column Table
        dr = tbl.Rows.Last
        numCells = dr.Cells.Count
        '
        If Me.tbl_rows_RowHasBottomStyle(dr) Then
            'The last Row has a spacer, Source or Note Style Style
            dr = tbl.Rows.First
            If dr.Cells.Count = 1 Then
                strFlag = "SingleColumnWithSource"
            Else
                strFlag = "MultiColumnWithSource"
            End If
        Else
            'The Last Row does Not have a Caption Style
            dr = tbl.Rows.First
            If dr.Cells.Count <> 1 Then
                strFlag = "MultiColumnNoSource"
            Else
                strFlag = "SingleColumnNoSource"
            End If

        End If
        '
        Return strFlag
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will go to the specified cell in the table (the default is the first cell) and
    ''' returns the name of the style
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="cellIndex"></param>
    ''' <returns></returns>
    Public Function tbl_CellStyle_GetFirstCellStyle(ByRef tbl As Word.Table, Optional cellIndex As Integer = 1) As String
        Dim drCell As Word.Cell
        Dim strCellStyleName As String
        Dim myStyle As Word.Style
        Dim rng As Word.Range
        '
        strCellStyleName = ""
        '
        Try
            drCell = tbl.Range.Cells.Item(cellIndex)
            rng = drCell.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            myStyle = rng.Style
            strCellStyleName = myStyle.NameLocal
            '
        Catch ex As Exception
            '
            strCellStyleName = ""
            '
        End Try
        '
        Return strCellStyleName
        '
    End Function
    '
    '
    Public Function tbl_get_LastBodyRow_ForMCWC(ByRef tbl As Word.Table) As Word.Row
        Dim Flag, numCells As Integer
        Dim IsCaptionStyle As Boolean
        Dim dr As Word.Row

        '
        Flag = 0
        IsCaptionStyle = False
        '
        'Start at the last row and run up the Table. If the number of cells per row
        'chnages to somethog other than 1 then we have found the body of the Table
        'if it gets to row one and its still 1 cell wide then we have a one column Table
        dr = tbl.Rows.Last
        numCells = dr.Cells.Count
        '
        For i = tbl.Rows.Count To 1 Step -1
            dr = tbl.Rows.Item(i)
            If dr.Cells.Count <> numCells Then
                Exit For
            End If
        Next
        '
        Return dr
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will return true is the supplied row (dr) contains one
    ''' of the 'bottom' styles; 'spacer', 'spacer_tbl', 'Note', 'Note label' 'Source'
    ''' </summary>
    ''' <param name="dr"></param>
    ''' <returns></returns>
    Public Function tbl_rows_RowHasBottomStyle(ByRef dr As Word.Row) As Boolean
        Dim CellStyle As Word.Style
        Dim drCell As Word.Cell
        Dim rng, OldRng As Word.Range
        Dim IsCaptionStyle As Boolean
        Dim sel As Selection
        '
        sel = Me.glb_get_wrdSel
        OldRng = sel.Range
        '
        drCell = dr.Cells.Item(1)
        rng = drCell.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()
        '
        CellStyle = sel.Style
        Select Case CellStyle.NameLocal
            Case "spacer", "spacer_tbl", "Note", "Note label", "Source"
                IsCaptionStyle = True
            Case Else
                IsCaptionStyle = False
        End Select
        '
        'Re-establish the initial Selection
        OldRng.Select()
        '
        Return IsCaptionStyle
    End Function
    '
    Public Function HeaderRowAndTable_SetRegularFormatting(headerOffSet As Single, ByRef drHeader As Word.Row) As Word.Table
        Dim rng As Word.Range
        Dim tbl As Word.Table
        '
        drHeader.Cells(1).LeftPadding = 0.0
        'drHeader.Range.wi
        drHeader.Cells(1).Width = drHeader.Cells(1).Width - headerOffSet
        drHeader.LeftIndent = drHeader.LeftIndent + headerOffSet
        '
        rng = drHeader.Range

        tbl = rng.Tables.Item(1)
        'tbl.Select()
        '
        tbl.AllowAutoFit = True
        tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed)
        '
        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        'tbl.PreferredWidth = tblWidth
        '
        Return tbl
    End Function
    '
    ''' <summary>
    ''' This method will determine whether the table has the same number of cells 
    ''' per row.. If not we will return false, that is, the table is not regular. This
    ''' test is not definitive, but will cover most circumstances
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_is_RegularByRow(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        '
        rslt = glb_tbls_isRegularByRow(tbl)
        '
        Return rslt
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will determine whether the current Selection is just under a 
    ''' Table
    ''' </summary>
    ''' <returns></returns>
    Public Function tbl_selection_IsUnderTable() As Boolean
        Dim rslt As Boolean
        Dim rng As Word.Range
        Dim k As Integer
        '
        rslt = False
        '
        rng = Me.glb_get_wrdSel().Range
        Try
            k = rng.Cells.Count
        Catch ex As Exception
            k = 0
        End Try
        rng.Move(WdUnits.wdCharacter, -1)
        If rng.Information(WdInformation.wdAtEndOfRowMarker) And k = 0 Then
            rslt = True
        End If
        '
        Return rslt
    End Function
    '
    ''' <summary>
    ''' This funtion will insert a standard AAC table at rng. The table will have numRows, numColumns and borders (if doBorders = true). If tblOutDent is left
    ''' at -1.0 the actual outdent will default to Me.glb_tbl_OutDent (mm). If tblOutdent is specified (in mm) it overrides the default
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="numRows"></param>
    ''' <param name="numColumns"></param>
    ''' <param name="doBorders"></param>
    ''' <param name="tblOutDent"></param>
    ''' <param name="strRowSelection"></param>
    ''' <param name="doWide"></param>
    ''' <param name="strBottomRows"></param>
    ''' <param name="textStyleName"></param>
    ''' <returns></returns>
    Public Function tbl_aacTable_build(ByRef rng As Word.Range, numRows As Integer, numColumns As Integer, doBorders As Boolean, Optional tblOutDent As Single = -1.0, Optional strRowSelection As String = "header", Optional doWide As Boolean = False, Optional strBottomRows As String = "sourceOnly", Optional textStyleName As String = "") As Word.Table
        Dim objTools As New cTools()
        Dim tbl As Word.Table
        Dim sect As Word.Section
        Dim doEnvelope As Boolean
        '
        doEnvelope = True
        '
        sect = rng.Sections.Item(1)
        '
        tbl = Me.tbl_build_Table_Standard(rng, numRows, numColumns, textStyleName)
        '
        'Wide only makes sense in single column layouts.. In multi column, the table has to be set
        'to 'floating'.. This is taken care of in the 'PlaceHolder Mgmnt' menu items in the 
        'PlaceHolder Tab
        '
        If tblOutDent < 0.0 Then
            tblOutDent = Me.glb_get_TableOutdent()
        Else
            tblOutDent = objTools.tools_math_MillimetersToPoints(tblOutDent)
        End If
        '
        '
        'Me.tbl_aacTable_Wide(tbl, tblOutDent, doWide)
        '
        tbl.AllowPageBreaks = True
        tbl.Rows.AllowBreakAcrossPages = False
        '
        '

        Return tbl
    End Function
    '
    ''' <summary>
    ''' This method takes as input a table. It will apply the Acil Allen basic Table, setting the table
    ''' to % preferred width with an actual preferred width of 100. The user can select whether to include the
    ''' Caption and Source. If strForceCaption = "", then the software decides what table caption (LT, ES, BD and AP)
    ''' to insert. Determining this from the page number format. If strForceCaption is set to 'LT', 'ES', 'BD' or 'AP'
    ''' then the Table caption is forced to one of these options.. FOr the source, the user can select by setting
    ''' strDoSource = "none", "sourceOnly", "sourceAndNote", "note"... The return value is the range of the Caption
    ''' paragraph
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="strForceCaptionTo"></param>
    ''' <param name="strDoSource"></param>
    ''' <returns></returns>
    Public Function tbl_format_rapidFormat(ByRef tbl As Word.Table, Optional strForceCaptionTo As String = "", Optional strDoSource As String = "sourceOnly") As Word.Range
        Dim objStylesMgr As New cStylesManager()
        Dim objPlhBase As New cPlHBase()
        Dim objWrkAround As New cWorkArounds()
        Dim myDoc As Word.Document
        Dim sect As Word.Section
        Dim rng, rngSrc, rngStart, rngTblStart As Word.Range
        Dim tblTemp As Word.Table
        Dim strMsg As String
        '
        strMsg = ""
        myDoc = tbl.Range.Document
        sect = tbl.Range.Sections.Item(1)
        '
        '********
        rng = sect.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        rngTblStart = tbl.Range
        rngTblStart.Collapse(WdCollapseDirection.wdCollapseStart)
        tblTemp = sect.Range.Tables.Item(1)
        '
        '
        If False Then
            'tbl is the first table in the section... Now we need to establish whether it is
            'at the top of the section
            rngStart = sect.Range
            rngStart.Collapse(WdCollapseDirection.wdCollapseStart)
            If rngStart.Information(WdInformation.wdWithInTable) Then
                'tbl is hard up against the section boundary
                'rng = objSectMgr.sct_set_SelforTableInsert()
                MsgBox("Table at the top of the section")
                GoTo finis
            End If
        End If
        '
        '********
        '
        rng = Nothing
        '
        Try
            'Set up with table set to WdPreferredWidthType.wdPreferredWidthPercent
            'with preferred width set to 100%
            Me.glb_tbl_apply_aacTableBasicStyle(tbl)
            '
            'Now we need to make space just before the table so that the caption doesn't interfere with
            'prior text.. Let's add a row (so we are working inside th table space). Then split it off,
            'then move up to the split off row and delete it
            '
            rng = Me.tbl_para_addAbove(tbl)
            'rng = Me.tbl_para_addAbove2(tbl)

            '
            'doCaption is set to false if we don't want a Caption and source information.
            'Typically you'd do this if you are going to do this if you are going to create
            'the tbale here and then encapsulate it later
            rng = objPlhBase.Plh_Captions_getAndWriteCaption(rng, strForceCaptionTo)
            '
            rngSrc = tbl.Range
            rngSrc.Collapse(WdCollapseDirection.wdCollapseEnd)
            'If doCaption Then rngSrc = Me.tbl_insert_SourceAndNoteText(rngSrc, strDoSource)
            rngSrc = Me.tbl_insert_SourceAndNoteText(rngSrc, strDoSource)
            '
            rng.Select()
            '
        Catch ex2 As Exception
            strMsg = "Error In Rapid Format.. Is your cursor in a table?" + vbCrLf + vbCrLf _
                + "Or is the table too irregular?" + vbCrLf _
                + "You may need to finish it off by hand."
            MsgBox(strMsg)
        End Try
        '
finis:
        '
        objWrkAround.wrk_fix_forCursorRace()
        objWrkAround.wrk_fix_forCursorRace()
        '
        Return rng
        '
    End Function
    '

    ''' <summary>
    ''' Will take the tab le tbl ans go through all the cells with the Row.Index = rowindex and will
    ''' make the selected border (strBorderType = 'top', 'bottom') visible or invisible. This function 
    ''' does not require the table to be regular
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="rowIndex"></param>
    ''' <param name="strBorderType"></param>
    ''' <param name="isVisible"></param>
    Public Sub tbl_format_rowBorderVisibility(ByRef tbl As Word.Table, rowIndex As Integer, strBorderType As String, isVisible As Boolean)
        Dim drCell As Word.Cell
        '
        For j = 1 To tbl.Range.Cells.Count
            drCell = tbl.Range.Cells.Item(j)
            If drCell.RowIndex > rowIndex Then Exit For
            If drCell.RowIndex = rowIndex Then
                If strBorderType = "top" Then drCell.Borders.Item(WdBorderType.wdBorderTop).Visible = False
                If strBorderType = "bottom" Then drCell.Borders.Item(WdBorderType.wdBorderBottom).Visible = False
            End If
        Next
    End Sub
    '
    Public Function tbl_format_rapidFormat_Encap(ByRef tbl As Word.Table, Optional strForceCaptionTo As String = "") As Word.Range
        Dim objStylesMgr As New cStylesManager()
        Dim objCaptionsMgr As New cCaptionManager()
        Dim tblTop, tblBottom, tblAll As Word.Table
        Dim objChptBase As New cChptBase()
        Dim objPlhBase As New cPlHBase()
        Dim myDoc As Word.Document
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim rng, rngSelection, rngSrc, rngTop As Word.Range
        Dim para As Word.Paragraph
        Dim strInsertType As String
        '
        myDoc = tbl.Range.Document
        rng = Nothing
        strInsertType = ""
        '
        Try
            'Do a rapid table format, but don't add the caption or Source
            'information
            rng = Me.tbl_format_rapidFormat(tbl, "none", "none")
            para = rng.Paragraphs.Item(1)
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            '
            tblTop = rng.Tables.Add(rng, 1, 1)
            Me.glb_tbl_apply_aacTableBasicStyle(tblTop)
            Me.tbl_colour_set_colourOfCellsToNone(tblTop.Range.Cells)
            tbl_format_rowBorderVisibility(tbl, 1, "top", False)

            '
            rngSrc = tbl.Range
            rngSrc.Collapse(WdCollapseDirection.wdCollapseEnd)
            'rngSrc.Move(WdUnits.wdCharacter, 1)
            tblBottom = rngSrc.Tables.Add(rngSrc, 1, 1)
            Me.glb_tbl_apply_aacTableBasicStyle(tblBottom)
            Me.tbl_colour_set_colourOfCellsToNone(tblBottom.Range.Cells)
            tblBottom.Range.Cells.Item(1).Borders.Item(WdBorderType.wdBorderBottom).Visible = False
            '
            rngSrc = tblBottom.Range.Cells.Item(1).Range
            rngSrc.Collapse(WdCollapseDirection.wdCollapseStart)
            Me.tbl_insert_SourceAndNoteText(rngSrc)
            '
            rng = tblTop.Range.Cells.Item(1).Range
            rng.Style = myDoc.Styles.Item("Caption")
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng = objPlhBase.Plh_Captions_getAndWriteCaption(rng, strForceCaptionTo)
            '
            tbl_colour_set_colourOfRow(tbl, 1, Me.var_tbl_colourHeader)
            '
            rngTop = tblTop.Range
            rngTop.Collapse(WdCollapseDirection.wdCollapseEnd)
            rngTop.Delete()
            '
            rngSrc = tbl.Range
            rngSrc.Collapse(WdCollapseDirection.wdCollapseEnd)
            'If doCaption Then rngSrc = Me.tbl_insert_SourceAndNoteText(rngSrc)
            '
            tblAll = rng.Tables.Item(1)
            tblAll.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
            tblAll.PreferredWidth = 100
            '
            'tblAll.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
            'tblAll.PreferredWidth = 

            '

            'Me.glb_tbl_apply_aacTableBasicStyle(tbl)

            'tblTop.Range.Style = 
            'tbl.Rows.Add(tbl.Rows.First)
            'rng = para.Range
            '
            GoTo finis

            tbl.Rows.Add(tbl.Rows.First)
            tbl.Rows.Add(tbl.Rows.First)
            '
            Me.glb_tbl_apply_aacTableBasicStyle(tbl)
            tbl.Range.Style = objStylesMgr.style_txt_getTableTextStyle(myDoc)
            '
            tbl.AllowPageBreaks = True
            tbl.Rows.AllowBreakAcrossPages = False
            '
            '
            dr = tbl.Rows.Item(1)
            dr.Cells.Merge()
            'dr.Cells.Shading.Texture = WdTextureIndex.wdTextureNone
            dr.Shading.ForegroundPatternColor = WdColor.wdColorAutomatic
            dr.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic

            '
            dr = tbl.Rows.Item(2)
            dr.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone
            dr.Range.Style = objStylesMgr.style_txt_getTableHeadingStyle(myDoc)
            '
            dr.Shading.ForegroundPatternColor = Me._glb_colour_purple_Dark
            For Each drCell In dr.Range.Cells
                drCell.Range.Text = "Heading"
            Next
            '
            rng = tbl.Rows.Item(1).Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            '*** Must fix up the selection mechanism
            rngSelection = Nothing
            '
            rng = objPlhBase.Plh_Captions_getAndWriteCaption(rng)
            '
            dr = tbl.Rows.Last
            tbl.Rows.Add(dr)
            dr = tbl.Rows.Last
            dr.Cells.Merge()
            dr.Borders.Item(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleNone
            dr.Cells.Item(1).BottomPadding = 2.0
            '
            strInsertType = "sourceOnly"
            rngSrc = dr.Range
            rngSrc.Collapse(WdCollapseDirection.wdCollapseStart)
            rngSrc = Me.tbl_insert_SourceAndNoteText(rngSrc)

finis:

        Catch ex As Exception

        End Try
        '
        rng.Select()
        '
        Return rng
        '
    End Function
    '
    ''' <summary>
    ''' This method will take a regular table and add a paragraph above it, by adding
    ''' a row, then the table is split. The top 'split off' row is deleted, leaving a 
    ''' paragraph at the top of the table. The range which is returned is the collapsed
    ''' at the beginning of this paragraph. The referenced table is returned as it was
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_para_addAbove(ByRef tbl As Word.Table) As Word.Range
        Dim rng As Word.Range
        'Dim drCell As Word.Cell
        'Dim para As Word.Paragraph

        tbl.Rows.Add(tbl.Rows.Item(1))
        tbl = tbl.Split(2)
        '
        '*****
        'drCell = tbl.Range.Cells.Item(1)
        'rng = drCell.Range
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'rng.Move(WdUnits.wdCharacter, -1)
        '
        'para = rng.Paragraphs.Add(rng)
        'rng = para.Range
        'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        'GoTo finis


        '*****
        '
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Move(WdUnits.wdCharacter, -2)
        'rng.Tables.Item(1).Delete()
        rng.Rows.Item(1).Delete()
        '
        'Now we go back to the paragraph above the table (inserted by the splitting process) and
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Move(WdUnits.wdCharacter, -1)
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
finis:
        '
        Return rng
        '
    End Function
    '
    Public Function tbl_para_addAbove2(ByRef tbl As Word.Table) As Word.Range
        Dim rng As Word.Range

        tbl.Rows.Add(tbl.Rows.Item(1))
        tbl = tbl.Split(2)
        '
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Move(WdUnits.wdCharacter, -2)
        rng.Rows.Item(1).Delete()
        '
        'Now we go back to the paragraph above the table (inserted by the splitting process) and
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Move(WdUnits.wdCharacter, -1)
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
finis:
        '
        Return rng
        '
    End Function
    '

    '
    '
    Public Sub tbl_format_rapidFormat2(ByRef tbl As Word.Table)
        Dim objStylesMgr As New cStylesManager()
        Dim myDoc As Word.Document
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim rng, rngSrc As Word.Range
        Dim numCols As Integer
        Dim tblTemp As Word.Table
        Dim lstOfColumns As New Collection()
        '
        myDoc = tbl.Range.Document
        '
        '
        Try

            'To allow for the header row we must add a new row
            If Me.glb_tbls_isRegular(tbl) Then
                tbl.Rows.Add(tbl.Rows.First)
            Else
                rng = Me.tbl_para_addAbove(tbl)
                '
                numCols = 0
                For j = 1 To tbl.Range.Cells.Count
                    drCell = tbl.Range.Cells.Item(j)
                    If drCell.RowIndex = 2 Then Exit For
                    lstOfColumns.Add(drCell.PreferredWidth,)
                    numCols += 1
                Next
                '
                tblTemp = rng.Tables.Add(rng, 1, numCols,)
                '
                For Each drCell In tbl.Range.Cells
                    If drCell.RowIndex = 2 Then Exit For

                Next
            End If
            '
            Me.glb_tbl_apply_aacTableBasicStyle(tbl)
            tbl.Range.Style = objStylesMgr.style_txt_getTableTextStyle(myDoc)
            '
            tbl.AllowPageBreaks = True
            tbl.Rows.AllowBreakAcrossPages = False
            '
            dr = tbl.Rows.Item(1)
            dr.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone
            dr.Range.Style = objStylesMgr.style_txt_getTableHeadingStyle(myDoc)
            '
            dr.Shading.ForegroundPatternColor = Me._glb_colour_purple_Dark
            For Each drCell In dr.Range.Cells
                drCell.Range.Text = "Heading"
            Next
            '
            rng = tbl.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            '****
            'Now we need to make space just beofre th table so that the caption doesn't interfere with
            'prior text.. Let's add a row (so we are working inside th table space). Then split it off,
            'then move up to the split off row and delete it
            '
            rng = Me.tbl_para_addAbove(tbl)
            '
            'tbl.Rows.Add(tbl.Rows.Item(1))
            'tbl = tbl.Split(2)
            'rng = tbl.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'rng.Move(WdUnits.wdCharacter, -2)
            'rng.Tables.Item(1).Delete()
            '
            'Now we go back to the paragraph above the table (inserted by the splitting process) and
            'rng = tbl.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'rng.Move(WdUnits.wdCharacter, -1)
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            '*****
            '
            '
            Dim strCaptiontext As String
            Dim objPlhBase As New cPlHBase()
            Dim rngSelection As Word.Range
            Dim strInsertType As String
            Dim objChptBase As New cChptBase()
            Dim strCaptionType As String
            '
            strCaptiontext = "Title Of Table"
            strCaptionType = ""
            '
            '
            'See cPlHBase.660 for the info here
            '*** Must fix up the selection mechanism
            rngSelection = Nothing
            '
            rng = objPlhBase.Plh_Captions_getAndWriteCaption(rng)


            'If objChptBase.chptBase_PageNumbering_isChapterBody() Then
            'Use this to determine which Table Caption option to use
            'rng = objPlhBase.Plh_Captions_InsertCaptions("Table", rng, True)

            'End If
            'If objChptBase.chptBase_PageNumbering_isES() Then
            'Use this to determine which Table Caption option to use
            'rng = objPlhBase.Plh_Captions_InsertCaptions("Table_ES", rng, True)


            'End If
            'If objChptBase.chptBase_PageNumbering_isAppendixBody() Then
            'Use this to determine which Table Caption option to use
            'rng = objPlhBase.Plh_Captions_InsertCaptions("Table_AP", rng, True)

            'End If


            'objPlhBase.Plh_Captions_InsertNameBeforeFields("Table", rng)
            'rngSelection = objPlhBase.Plh_Captions_Insert("Report", "Table", rng, strCaptiontext)
            'rngSelection = objPlhBase.Plh_Captions_Insert(strCaptionType, "Table", rng, strCaptiontext)

            'objPlhBase.Plh_Captions_doCaptionStyle(rngSelection)
            '
            strInsertType = "sourceOnly"
            rngSrc = tbl.Range
            rngSrc.Collapse(WdCollapseDirection.wdCollapseEnd)
            rngSrc = Me.tbl_insert_SourceAndNoteText(rngSrc)

            'rng.Paragraphs.Item(1).Style = objGlobals.glb_get_wrdStyle(objGlobals.glb_var_style_tblSourceStyle)

            '
            'drCell = dr.Range.Cells.Item(1)
            'Me.glb_selection_toCellText(drCell)
            '
finis:
            '
            rng.Select()
finis2:
            '
        Catch ex2 As Exception
            MsgBox("Error In Rapid Format.. Is your cursor In a table?")
        End Try

    End Sub
    '
    Public Sub tbl_format_setRowOneToHeadingStyle(ByRef tbl As Word.Table)
        Dim objStylesMgr As New cStylesManager()
        Dim drCell As Word.Cell
        Dim myStyle As Word.Style
        '
        myStyle = objStylesMgr.style_txt_getTableHeadingStyle(Me.glb_get_wrdActiveDoc)
        '
        For j = 1 To tbl.Range.Cells.Count
            drCell = tbl.Range.Cells.Item(j)
            If drCell.RowIndex = 2 Then Exit For
            drCell.Range.Style = myStyle
        Next

    End Sub
    Public Function tbl_aacTable_ConvertTo_AAC(ByRef tbl As Word.Table, doBorders As Boolean, Optional tblOutDent As Single = -1.0, Optional strRowSelection As String = "header", Optional doWide As Boolean = False, Optional strBottomRows As String = "sourceOnly", Optional textStyleName As String = "") As Word.Table
        Dim sect As Word.Section
        Dim leftIndentOriginal As Single
        Dim isAACTable As Boolean
        Dim objTools As New cTools()
        '
        sect = tbl.Range.Sections.Item(1)
        leftIndentOriginal = tbl.Rows.Item(2).LeftIndent
        '
        If tblOutDent < 0.0 Then
            tblOutDent = Me.glb_get_TableOutdent()
        Else
            tblOutDent = objTools.tools_math_MillimetersToPoints(tblOutDent)
        End If
        '
        isAACTable = Me.tbl_aacTable_isAACTable(tbl)
        If Not isAACTable Then
            Me.tbl_fix_Table(tbl, doBorders, Me._glb_colour_TableBorders, True, textStyleName)
            Me.tbl_aacTable_Wide(tbl, tblOutDent, doWide)
            Me.tbl_build_headerUnitsAndCaptionRow(tbl, tblOutDent, strRowSelection)
            Me.tbl_build_sourceAndSpacerRow(tbl, strBottomRows)
        Else
            Me.tbl_aacTable_Wide(tbl, tblOutDent, doWide)
        End If
        '
        Return tbl

    End Function
    '
    ''' <summary>
    ''' This method will return true if the table (tbl) is an AAC Table. Note that the test looks for
    ''' at least two AAC Table characteristics (leftIndent of the first row and the heading style in the
    ''' first cell)
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_aacTable_isAACTable(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim leftIndent As Single
        '
        rslt = False
        '
        Try
            leftIndent = tbl.Rows.Item(1).LeftIndent
            If Me.tbl_CellStyle_GetFirstCellStyle(Me.glb_get_wrdSelTbl) = Me.glb_var_style_tblHeaderStyle And leftIndent < 0.0 Then rslt = True

        Catch ex As Exception
            rslt = False
        End Try
        '
        '
        Return rslt
    End Function
    '
    ''' <summary>
    ''' This function will set the specified table (tbl) to a Wide aac table, taking into account the Table OutDent.
    ''' It will only do this for Single column pages. In addition doWide must be true, otherwise the function
    ''' does nothing. This aspect is here just to reduce lines of code in the calling routing
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="tblOutDent"></param>
    ''' <returns></returns>
    Public Function tbl_aacTable_Wide(ByRef tbl As Word.Table, tblOutDent As Single, doWide As Boolean, Optional isAACTable As Boolean = False) As Word.Table
        Dim objPlhBaseMgr As New cPlHBase()
        Dim sect As Word.Section
        Dim leftIndent, deltaWidth, deltaLeftIndent As Single
        Dim leftIndentOriginal As Single
        Dim dr As Word.Row
        '
        sect = tbl.Range.Sections.Item(1)
        leftIndentOriginal = tbl.Rows.Item(1).LeftIndent
        '
        '
        '
        Try
            If sect.PageSetup.TextColumns.Count = 1 Then
                If doWide And Not isAACTable Then
                    leftIndent = sect.PageSetup.LeftMargin - (glb_hfs_getHFTableEdge(sect, "header_leftEdge") + tblOutDent)
                    deltaLeftIndent = leftIndent + leftIndentOriginal
                    deltaWidth = deltaLeftIndent / tbl.Columns.Count
                    For Each dr In tbl.Rows
                        dr.LeftIndent = -leftIndent
                    Next
                    For Each drCol In tbl.Columns
                        drCol.Width = drCol.Width + deltaWidth
                    Next

                End If
                If doWide And isAACTable Then
                    leftIndent = sect.PageSetup.LeftMargin - (glb_hfs_getHFTableEdge(sect, "header_leftEdge") + tblOutDent)
                    deltaLeftIndent = leftIndent + leftIndentOriginal
                    deltaWidth = deltaLeftIndent / tbl.Columns.Count
                    For Each dr In tbl.Rows
                        If dr.Index <> 1 Then dr.LeftIndent = -leftIndent
                    Next
                    For Each drCol In tbl.Columns
                        'drCol.Width = drCol.Width + deltaWidth
                    Next

                End If

            End If
            '
        Catch ex As Exception

        End Try
        '
loop1:
        Return tbl
        '
    End Function
    '
    ''' <summary>
    ''' This method will insert a header and/or units row at the top of the Table (tbl), depending on the
    ''' value of 'strRowSelection'. This can be; 'header', 'unitsRow', or 'header+UnitsRow'
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="strHeadersAndUnits"></param>
    ''' <returns></returns>
    Public Function tbl_build_headerUnitsAndCaptionRow(ByRef tbl As Word.Table, tblOutDent As Single, Optional strHeadersAndUnits As String = "header", Optional strTextStyleName As String = "") As Word.Row
        Dim objStylesMgr As New cStylesManager()
        Dim myDoc As Word.Document
        Dim sect As Word.Section
        Dim dr As Word.Row
        '
        myDoc = tbl.Range.Document
        '
        'Default to standard table text style
        '
        If strTextStyleName = "" Then strTextStyleName = Me.glb_var_style_tblTextStyle
        '
        sect = tbl.Range.Sections.Item(1)
        dr = Nothing
        '
        'Make certain that we start with no heading rows
        tbl.ApplyStyleHeadingRows = False
        tbl.Range.Style = myDoc.Styles.Item(strTextStyleName)
        '
        Select Case strHeadersAndUnits
            Case "header"
                tbl.ApplyStyleHeadingRows = True
                'tbl.Range.Style = objStylesMgr.style_txt_getTableTextStyle(myDoc)
                tbl.Rows.First.Range.Style = objStylesMgr.style_txt_getTableHeadingStyle(myDoc)
                '
                'Set Heading row repeat
                tbl.Rows.Item(1).HeadingFormat = True
                '
                'Get rid of Units row
                'tbl.Rows.Item(2).Delete()
                'tbl.Rows.First.Borders.Item(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleNone
                'dr = tbl.Rows.Add(tbl.Rows.First)
                'Me.tbl_build_headerRow(dr, tblOutDent)
                '
            Case "unitsRow"
                'tbl.ApplyStyleHeadingRows = False

                dr = tbl.Rows.Item(1)
                'dr = tbl.Rows.Add(tbl.Rows.First)
                dr.Shading.BackgroundPatternColor = Me._glb_colour_UnitsGrey
                dr.Shading.ForegroundPatternColor = Me._glb_colour_UnitsGrey
                dr.Shading.Texture = Word.WdTextureIndex.wdTextureSolid
                dr.Range.Style = Me.glb_get_wrdStyle(Me.glb_var_style_tblUnitsStyle)
                '
            Case "header+UnitsRow"
                tbl.ApplyStyleHeadingRows = True
                tbl.Rows.First.Range.Style = objStylesMgr.style_txt_getTableHeadingStyle(myDoc)

                dr = tbl.Rows.Item(2)
                dr.Shading.BackgroundPatternColor = Me._glb_colour_UnitsGrey
                dr.Shading.ForegroundPatternColor = Me._glb_colour_UnitsGrey
                dr.Shading.Texture = Word.WdTextureIndex.wdTextureSolid
                dr.Range.Style = Me.glb_get_wrdStyle(Me.glb_var_style_tblUnitsStyle)
                '
                '
                'Set Heading row repeat
                tbl.Rows.Item(1).HeadingFormat = True
                '
                'dr = tbl.Rows.Add(tbl.Rows.First)
                'Me.tbl_build_headerRow(dr, tblOutDent)
                '
            Case "caption"
                'dr = tbl.Rows.Add(tbl.Rows.First)
                'dr.Borders.Item(WdBorderType.wdBorderRight).LineStyle = WdLineStyle.wdLineStyleNone
                'dr.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic
                'dr.Shading.Texture = Word.WdTextureIndex.wdTextureNone
                'dr.Range.Style = Me.glb_get_wrdStyle(Me.glb_var_style_tblCaptionStyle)
                'dr.Cells.Merge()
                '
            Case "none"
                'dr = tbl.Rows.Add(tbl.Rows.First)

        End Select
        '
finis:
        Return dr
        '
    End Function
    '
    ''' <summary>
    ''' This method will modify the supplied Row (dr) to be a AAC header row. If the tblOutDent is not supplied, then
    ''' the outdent is the standard 'Me.glb_tbl_OutDent' (which is internally chnaged from mm to points). If it is supplied (in pts),
    ''' then that value is used
    ''' </summary>
    ''' <param name="dr"></param>
    ''' <param name="tblOutDent"></param>
    ''' <returns></returns>
    Public Function tbl_build_headerRow(ByRef dr As Word.Row, Optional tblOutDent As Single = -1.0) As Word.Row
        Dim offSet As Single
        Dim drCell As Word.Cell
        '
        Try
            If tblOutDent < 0.0 Then
                'Default condition
                offSet = -Me.glb_get_TableOutdent()
            Else
                offSet = -tblOutDent
            End If
            '
            If Not tblOutDent = 0.0 Then
                'Don't offset if Table outdent is 0.0
                dr.LeftIndent = dr.LeftIndent + offSet
            End If
            '
            dr.Shading.BackgroundPatternColor = Me.var_tbl_colourHeader
            dr.Shading.Texture = Word.WdTextureIndex.wdTextureNone
            dr.Range.Style = Me.glb_get_wrdStyle(Me.glb_var_style_tblHeaderStyle)
            '
            dr.Borders.Item(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleNone
            dr.Borders.Item(WdBorderType.wdBorderRight).LineStyle = WdLineStyle.wdLineStyleSingle
            dr.Borders.Item(WdBorderType.wdBorderRight).LineWidth = WdLineWidth.wdLineWidth075pt
            dr.Borders.Item(WdBorderType.wdBorderRight).Color = Me.var_tbl_colourHeader
            '
            If Not tblOutDent = 0.0 Then
                drCell = dr.Cells.Item(1)
                drCell.Width = drCell.Width - offSet
                drCell.LeftPadding = Math.Abs(offSet)
            End If
            '
            For i = 1 To dr.Cells.Count
                drCell = dr.Cells.Item(i)
                drCell.Range.Text = "Heading"
            Next i

        Catch ex As Exception

        End Try

        Return dr
    End Function
    '
    ''' <summary>
    ''' This method will insert a 'Data Source/Note' and Spacer Row at the end of the table.
    ''' What type of 'Data Source/Note' is dependent on the value of strRow ("sourceOnly", "note",
    ''' "sourceAndNote")
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="strRowsToInsert"></param>
    ''' <returns></returns>
    Public Function tbl_build_sourceAndSpacerRow(ByRef tbl As Word.Table, strRowsToInsert As String) As Word.Row
        Dim rng As Word.Range
        Dim dr As Word.Row
        '
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
        Select Case strRowsToInsert
            Case "sourceOnly", "note", "sourceAndNote"
                dr = tbl_build_sourceRow(tbl, strRowsToInsert)
                'To Add rows to the end
                'dr = Me.tbl_rows_AddToEndOfTable(tbl)
                'dr.Range.Style = glb_get_wrdStyle(glb_style_tblSourceStyle)
                ' dr.Range.Cells.Merge()
                'dr.Range.Cells.Item(1).BottomPadding = 2.0
                'rng2 = dr.Range
                'rng2.Collapse(WdCollapseDirection.wdCollapseStart)
                'Me.tbl_insert_SourceAndNoteText(rng2)
                'dr.Cells.Item(1).BottomPadding
                '
                'dr = tbl_build_spacerRow(tbl)
                '
                '
            Case Else
                dr = tbl_build_spacerRow(tbl)
        End Select
        '
        Return dr
        '
    End Function
    '
    ''' <summary>
    ''' This method will add and format a 'Source' row at the end of the Table (tbl)
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_build_sourceRow(ByRef tbl As Word.Table, strInsertType As String) As Word.Row
        Dim dr As Word.Row
        Dim rng As Word.Range
        Dim brdr As Word.Border
        '
        '
        dr = Me.tbl_rows_AddToEndOfTable(tbl)
        '
        '
        dr.Range.Style = glb_get_wrdStyle(glb_var_style_tblSourceStyle)
        dr.Range.Cells.Merge()
        dr.Range.Cells.Item(1).BottomPadding = 0.0
        brdr = dr.Borders.Item(WdBorderType.wdBorderBottom)
        brdr.LineStyle = WdLineStyle.wdLineStyleNone
        '
        rng = dr.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        Me.tbl_insert_SourceAndNoteText(rng, strInsertType)
        '
finis:
        '
        Return dr
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will add and format a 'Source' row at the end of the Table (tbl). It does so in a table 
    ''' safe way. It returns the range of the last cell
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_build_sourceRowAsCells(ByRef tbl As Word.Table, strInsertType As String) As Word.Range
        Dim drCellLast As Word.Cell
        Dim rngOfCells As Word.Range
        Dim rng As Word.Range
        Dim brdr As Word.Border
        '
        rngOfCells = tbl_rows_AddToEndOfTable_As_Cells(tbl)
        rngOfCells.Cells.Merge()
        '
        rngOfCells.Style = glb_get_wrdStyle(glb_var_style_tblSourceStyle)
        rngOfCells.Cells.Item(1).BottomPadding = 0.0
        brdr = rngOfCells.Borders.Item(WdBorderType.wdBorderBottom)
        brdr.LineStyle = WdLineStyle.wdLineStyleNone
        '
        Me.tbl_colour_set_colourOfCellsToNone(rngOfCells.Cells)
        '
        drCellLast = tbl.Range.Cells.Item(tbl.Range.Cells.Count)
        '
        rng = drCellLast.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        Me.tbl_insert_SourceAndNoteText(rng, strInsertType)
        '
        '
        Return drCellLast.Range
        '
    End Function
    '
    '
    '
    ''' <summary>
    ''' This method will insert a top row above an existing standard table, encapsulating the existing 
    ''' 'Caption' paragraph
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_build_topRowTableTypeSafe(ByRef tbl As Word.Table) As Word.Table
        Dim objParas As New cParas()
        Dim tblBody As Word.Table
        Dim rng, rngCaption, rngCell, rngParaSrc As Word.Range
        Dim drCell As Word.Cell
        Dim brdr As Word.Border
        Dim strSrc As String
        '
        tblBody = Nothing
        rngParaSrc = Nothing
        strSrc = ""
        '
        Try
            rng = tbl.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Move(WdUnits.wdCharacter, -1)
            rngCaption = rng.Paragraphs.Item(1).Range
            '
            'para = rng.Paragraphs.Item(1)
            rngCaption.Cut()
            '
            tbl.Rows.Add(tbl.Rows.First)
            tbl.Rows.Item(1).Cells.Merge()
            '
            drCell = tbl.Range.Cells.Item(1)
            brdr = drCell.Borders.Item(WdBorderType.wdBorderBottom)
            brdr.LineStyle = WdLineStyle.wdLineStyleNone
            rngCell = drCell.Range
            rngCell.Style = glb_get_wrdActiveDoc.Styles("Caption")
            Me.tbl_colour_set_colourOfCellToNone(drCell)
            rngCell.Collapse(WdCollapseDirection.wdCollapseStart)
            'rngCell.Paragraphs
            rngCell.Paste()
            '
            'Me.tbl_delete_ParaAtEndOfTable(tbl)
            Me.tbl_delete_ParaAtTopOfTable(tbl)
            '
        Catch ex As Exception

        End Try
        '
        Return tbl
    End Function
    '
    '
    ''' <summary>
    ''' This method will insert a top row above an existing standard table, encapsulating the existing 
    ''' 'Caption' paragraph
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_build_topRow(ByRef tbl As Word.Table) As Word.Table
        Dim objParas As New cParas()
        Dim tblBody As Word.Table
        Dim rng, rngCaption, rngCell, rngParaSrc As Word.Range
        Dim drCell As Word.Cell
        Dim brdr As Word.Border
        Dim strSrc As String
        '
        tblBody = Nothing
        rngParaSrc = Nothing
        strSrc = ""
        '
        Try
            rng = tbl.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Move(WdUnits.wdCharacter, -1)
            rngCaption = rng.Paragraphs.Item(1).Range
            '
            'para = rng.Paragraphs.Item(1)
            rngCaption.Cut()
            '
            tbl.Rows.Add(tbl.Rows.First)
            tbl.Rows.Item(1).Cells.Merge()
            '
            drCell = tbl.Range.Cells.Item(1)
            brdr = drCell.Borders.Item(WdBorderType.wdBorderBottom)
            brdr.LineStyle = WdLineStyle.wdLineStyleNone
            rngCell = drCell.Range
            rngCell.Style = glb_get_wrdActiveDoc.Styles("Caption")
            Me.tbl_colour_set_colourOfCellToNone(drCell)
            rngCell.Collapse(WdCollapseDirection.wdCollapseStart)
            'rngCell.Paragraphs
            rngCell.Paste()
            '
            'Me.tbl_delete_ParaAtEndOfTable(tbl)
            Me.tbl_delete_ParaAtTopOfTable(tbl)
            '
        Catch ex As Exception

        End Try
        '
        Return tbl
    End Function
    '

    '
    ''' <summary>
    ''' This method will add and format a spacer row at the end of the Table (tbl)
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_build_spacerRow(ByRef tbl As Word.Table) As Word.Row
        Dim dr As Word.Row
        '
        dr = Me.tbl_rows_AddToEndOfTable(tbl)
        'rng = tbl.Range
        'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        'dr = tbl.Rows.Add(rng)
        'dr.Range.Cells.Item(1).Range.Text = "Spacer"
        dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
        dr.Height = var_glb_tbl_bottomSpacerRowHeight
        dr.Range.Style = glb_get_wrdStyle("spacer")
        dr.Range.Cells.Merge()
        '
        Return dr
        '
    End Function
    '
    ''' <summary>
    ''' This method will insert the Source information at the specific range, which
    ''' needs to be collapsed to a single point.. It will return a range that includes
    ''' all of the inserted paragraphs. The parameter strInsertType can take on the values
    ''' 'none', 'sourceOnly', 'sourceAndNote', 'note'
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <returns></returns>
    Public Function tbl_insert_SourceAndNoteText(ByRef rng As Word.Range, Optional strInsertType As String = "sourceOnly") As Word.Range
        '
        Dim para As Word.Paragraph
        Dim rngpara As Word.Range
        Dim myDoc As Word.Document
        Dim i As Integer
        '
        myDoc = rng.Document

        Select Case strInsertType
            Case "none"

            Case "sourceOnly"
                rng.Text = "Source: ACIL Allen"
                rng.Paragraphs.Item(1).Style = Me.glb_get_wrdStyle(Me.glb_var_style_tblSourceStyle)
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                'rng.MoveStart(WdUnits.wdCharacter, -12)
                rng.MoveStart(WdUnits.wdCharacter, -10)

                '
            Case "note"
                rng.Text = "Note"
                rng.Paragraphs.Item(1).Style = Me.glb_get_wrdStyle(Me.glb_var_style_tblNoteStyle)
                'rng.Collapse(WdCollapseDirection.wdCollapseStart)

            Case "sourceAndNote"
                'rng.Text = "a this is a reference note" + vbCrLf + "Note" + vbCrLf + "Source" + vbCrLf
                rng.Text = "Source" + vbCrLf + "Note" + vbCrLf + "a this is a reference note" + vbCrLf

                For i = 1 To rng.Paragraphs.Count
                    para = rng.Paragraphs.Item(i)
                    Select Case i
                        Case 1
                            para.Style = myDoc.Styles("Source")

                        Case 2
                            para.Style = myDoc.Styles("Note")
                        Case 3
                            para.Style = myDoc.Styles("Note")
                            rngpara = para.Range
                            rngpara.Collapse(WdCollapseDirection.wdCollapseStart)
                            rngpara.MoveEnd(WdUnits.wdCharacter)
                            rngpara.Style = myDoc.Styles("Note label")
                    End Select
                    '
                Next
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                rng.MoveStart(WdUnits.wdCharacter, -2)
                rng.MoveEnd(WdUnits.wdCharacter, -1)
                rng.Delete()

                '
            Case "notex"
                rng.Text = "Note"
                rng.Collapse(WdCollapseDirection.wdCollapseStart)

                para = rng.Paragraphs.Add(rng)
                para.Range.Style = myDoc.Styles("Note")
                rng = para.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Text = "a this is a reference note"
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.MoveEnd(WdUnits.wdCharacter)
                rng.Font.Superscript = True
                '
                rng = para.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.MoveEnd(WdUnits.wdParagraph, 2)

            Case Else
                rng.Text = "Source: ACIL Allen"
                'rng.MoveStart(WdUnits.wdCharacter, -10)
                rng.Paragraphs.Item(1).Style = Me.glb_get_wrdStyle(Me.glb_var_style_tblSourceStyle)
        End Select
        '
        '
        Return rng

    End Function
    '


    ''' <summary>
    ''' This method will build a standard regular table (at rng) with rows and columns as specified by numRows and numColumns.
    ''' The table style will be the AAC style cGlobals.glb_style_tblTextStyle. If doBorders is true, then the Table will have
    ''' inside Horizontal borders of 50pt set to the AAC colour Me.colour_TableBorders. The table style is set to the 'textStyleName'.
    ''' If 'textStyleName' is set to "", then the Table is formatted with the 'Table text' style (glb_style_tblTextStyle)
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="numRows"></param>
    ''' <param name="numColumns"></param>
    ''' <returns></returns>
    Public Function tbl_build_Table_Standard(ByRef rng As Word.Range, numRows As Integer, numColumns As Integer, Optional strTextStyle As String = "") As Word.Table
        Dim tbl As Word.Table
        Dim myDoc As Word.Document
        Dim objStylesMgr As New cStylesManager()
        Dim myStyle As Word.Style
        '
        '
        myDoc = rng.Document
        Try
            myStyle = myDoc.Styles.Item(strTextStyle)
        Catch ex As Exception
            myStyle = myDoc.Styles.Item("Table text")
        End Try
        '
        tbl = rng.Tables.Add(rng, numRows, numColumns)
        Me.glb_tbl_apply_aacTableNoLinesStyle(tbl)
        '
        tbl.Range.Style = myStyle

        tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed)
        '
        tbl.AllowPageBreaks = True
        tbl.Rows.AllowBreakAcrossPages = False
        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
        tbl.PreferredWidth = 100

        'tbl.Style = Me.var_tbl_TableStyleDefault
        'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
        'tbl.PreferredWidth = 100
        'Me.tbl_fix_Table(tbl, doBorders, Me._glb_colour_TableBorders, True, textStyleName, doTopBorder, doBottomBorder)
        '
        Return tbl
    End Function
    '
    ''' <summary>
    ''' This method will insert a PullOut/SideNote 'Picture' Table as a Floating Table locked
    ''' to the paragraph containing the current selection. The 'PullOut' left edge is aligned to the
    ''' left edge of the Header Table. The Pullout Table is the return object
    ''' </summary>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <returns></returns>
    Public Function tbl_build_Table_PullOutPict(Optional width As Single = 102, Optional height As Single = 154) As Word.Table
        Dim objGrfxMgr As New cGraphicsMgr()
        Dim tbl As Word.Table
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        Dim myStyle As Word.Style
        Dim objFloatMgr As New cPlHFloatingMgr()
        'Dim width, height As Single     'Pict graphic dimensions in pts
        '
        'width = 102.2
        'height = 154
        '
        rng = Me.glb_get_wrdSelRng
        para = rng.Paragraphs.Item(1)
        rng = para.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        tbl = Me.tbl_build_Table_Standard(rng, 3, 1, True)
        tbl.Columns.Item(1).Width = width
        '
        rng = tbl.Range.Cells.Item(1).Range
        rng.Style = Me.glb_get_wrdStyle("PullOut Title")
        Try
            myStyle = rng.Style
            myStyle.ParagraphFormat.KeepWithNext = True
        Catch ex As Exception

        End Try
        '
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Text = "Insert title here"
        '
        rng = tbl.Range.Cells.Item(3).Range
        rng.Style = Me.glb_get_wrdStyle("Source")
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Text = "Source"
        rng.ParagraphFormat.LeftIndent = 0.0
        tbl.Range.Cells.Item(3).BottomPadding = 1.6
        '
        rng = tbl.Range.Cells.Item(2).Range
        objGrfxMgr.grfx_inline_insertShape(tbl.Range.Cells.Item(2), width, height, Me._glb_colour_FigureFill)
        '
        'Now float the table
        objFloatMgr.Plh_Float_LockToToParagraph_RegularTable(tbl, Me.glb_hfs_getHFTableEdge(tbl.Range.Sections.Item(1), "header_leftEdge"))
        'objFloatMgr.Plh_Float_LockInPosition_RelativeToMargins_RegularTable(tbl, Me.glb_hfs_getHFTableEdge(tbl.Range.Sections.Item(1), "header_leftEdge"))
        '
        Return tbl
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will insert an Emphasis Box as a Floating Table locked
    ''' to the paragraph containing the current selection. The 'Emphasis Box' is either aligned to
    ''' the left margin (strLeftRight = 'left'), right margin (strLeftRight = 'right') or centred (strLeftRight = 'centre')
    ''' If the width of the box is not specified (i.e. left at -1.0), then the width is set to 
    ''' 1/3 of the width between left/right margins.
    ''' </summary>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <returns></returns>
    Public Function tbl_build_Table_EmphasisBox(Optional strLeftRight As String = "left", Optional width As Single = -1.0, Optional height As Single = 154) As Word.Table
        Dim tbl As Word.Table
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        Dim sect As Word.Section
        'Dim width, height As Single     'Pict graphic dimensions in pts
        '
        'width = 102.2
        'height = 154
        '
        rng = Me.glb_get_wrdSelRng
        sect = glb_get_wrdSect()
        para = rng.Paragraphs.Item(1)
        rng = para.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        tbl = Nothing
        '
        '
        Select Case strLeftRight
            Case "left"
                tbl = Me.tbl_build_Table_Standard(rng, 1, 1, "Emphasis Text (Left)")
                tbl.PreferredWidth = 33
                'tbl.Style = 
                'If width < 0.0 Then width = glb_get_widthBetweenMargins(sect) / 3.0
                'tbl.Columns.Item(1).Width = width
                tbl.Rows.WrapAroundText = True
                tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
                tbl.Rows.Alignment = WdRowAlignment.wdAlignRowLeft
                tbl.Rows.HorizontalPosition = 0.0
                '
                tbl.Rows.DistanceTop = 3.0
                tbl.Rows.DistanceBottom = 3.0
                tbl.Rows.DistanceLeft = 0.0
                tbl.Rows.DistanceRight = 6.0
                '
                rng = tbl.Range.Cells.Item(1).Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Text = "Insert Left Emphasis text here"

            Case "right"
                tbl = Me.tbl_build_Table_Standard(rng, 1, 1, "Emphasis Text (Right)")
                tbl.PreferredWidth = 33
                'If width < 0.0 Then width = glb_get_widthBetweenMargins(sect) / 3.0
                'tbl.Columns.Item(1).Width = width
                tbl.Rows.WrapAroundText = True
                tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
                tbl.Rows.Alignment = WdRowAlignment.wdAlignRowRight
                '
                tbl.Rows.DistanceTop = 3.0
                tbl.Rows.DistanceBottom = 3.0
                tbl.Rows.DistanceLeft = 6.0
                tbl.Rows.DistanceRight = 0.0
                '
                rng = tbl.Range.Cells.Item(1).Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Text = "Insert Right Emphasis text here"
                '
            Case "centre"
                tbl = Me.tbl_build_Table_Standard(rng, 1, 1, "Emphasis Text (Centre)")
                tbl.PreferredWidth = 33

                'tbl.Columns.Item(1).Width = glb_get_widthBetweenMargins(sect) / 3.0
                'tbl.Rows.WrapAroundText = True
                tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
                tbl.Rows.Alignment = WdRowAlignment.wdAlignRowCenter
                '
                tbl.Rows.DistanceTop = 0.0
                tbl.Rows.DistanceBottom = 0.0
                tbl.Rows.DistanceLeft = 3.0
                tbl.Rows.DistanceRight = 3.0
                '
                rng = tbl.Range.Cells.Item(1).Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                rng.Text = "Insert Centre Emphasis text here"

        End Select
        '
        '
        '
        'Now float the table
        'tbl.Rows.WrapAroundText = True


        'tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionColumn
        'tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
        'tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionInnerMarginArea
        'tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
        '
        'tbl.Rows.Alignment = WdRowAlignment.wdAlignRowLeft
        'tbl.Rows.Alignment = WdRowAlignment.wdAlignRowRight
        'tbl.Rows.Alignment = WdRowAlignment.wdAlignRowCenter


        'tbl.Rows.HorizontalPosition = 0.0
        'tbl.Rows.RelativeHorizontalPosition = 20.0
        'tbl.Rows.HorizontalPosition = Me.glb_hfs_getHFTableEdge(tbl.Range.Sections.Item(1), "header_leftEdge")
        tbl.Rows.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionParagraph
        tbl.Rows.VerticalPosition = 0.1
        tbl.Rows.AllowOverlap = False
        '
        Return tbl
        '
    End Function
    '
    ''' <summary>
    ''' This method will take the current table. Normally this is a new table created by using Microsoft standar
    ''' tools. It applies a table style and then makes sure that all other aspects of the table are set to AAC
    ''' requirements
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub tbl_fix_Table(ByRef tbl As Word.Table)
        Dim objStylesMgr As New cStylesManager()
        '
        'Remember to adjust the other areas where we use (Basic)
        glb_tbl_apply_aacTableNoLinesStyle(tbl)
        'objStylesMgr.s

    End Sub

    '
    '
    ''' <summary>
    ''' This method will set the cell margins for the current Table to 0 and
    ''' the Borders to null (depending on the value of leavBorders). The AutoFit behaviour
    ''' is set the wdAutoFitFixed.. Note that settings in AutoFitBehavior will cause
    ''' alloAutoFit to change (see http://msdn.microsoft.com/en-us/library/office/ff820953(v=office.15).aspx)
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="leaveBorders"></param>
    ''' <remarks></remarks>
    Public Sub tbl_fix_Table(ByRef tbl As Table, ByVal leaveBorders As Boolean)
        Dim brdr As Border
        Dim currentSect As Section
        Dim widthBetweenMargins As Single
        '
        Try
            'myDoc = Globals.ThisAddin.Application.ActiveDocument
            currentSect = tbl.Range.Sections.Item(1)
            widthBetweenMargins = Me.glb_get_widthBetweenMargins(currentSect)
            '
            tbl.TopPadding = 0.0#
            tbl.BottomPadding = 0.0#
            tbl.LeftPadding = 0.0#
            tbl.RightPadding = 0.0#
            'tbl.PreferredWidth = widthBetweenMargins
            '
            tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitFixed)
            'tbl.Rows.AllowBreakAcrossPages = False
            'tbl.AllowAutoFit = True
            'Call tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow)    'columns change size o accomodate text.. sets AllowAutoFit = true
            'Call tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed)      'columns sizes don't change to accommodate text, will set AllowAutoFit = false
            'Call tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)   'Table width changes to accommodate text.. Not content, then width=0.. sets AlloAutoFit = true
            '
            tbl.AllowPageBreaks = True                                           'Will allow a row to break across pages
            tbl.Rows.AllowBreakAcrossPages = False
            '
            If leaveBorders Then
            Else
                For Each brdr In tbl.Borders
                    brdr.LineStyle = Word.WdLineStyle.wdLineStyleNone
                Next brdr
            End If

        Catch ex As Exception
            MsgBox("Error-Table_Fix")
        End Try
        '
    End Sub
    '
    Public Function tbl_bordersCell_colourAndVisibility(ByRef btn As System.Windows.Forms.ToolStripButton, ByVal doBorders As Boolean) As Boolean
        Dim rslt As Boolean
        Dim drCells As Word.Cells
        Dim btnColor As System.Drawing.Color
        '
        rslt = False
        '
        Try
            btnColor = btn.BackColor
            drCells = glb_get_wrdSel.Cells
            If doBorders Then
                '
                If drCells.Count > 1 Then
                    drCells.Borders(WdBorderType.wdBorderHorizontal).LineStyle = WdLineStyle.wdLineStyleSingle
                    drCells.Borders(WdBorderType.wdBorderHorizontal).LineWidth = WdLineWidth.wdLineWidth050pt
                    drCells.Borders(WdBorderType.wdBorderHorizontal).Color = RGB(btnColor.R, btnColor.G, btnColor.B)
                    'rng.Rows.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
                    'rng.Rows.Borders(wdBorderTop).LineWidth = wdLineWidth050pt
                    'rng.Rows.Borders(wdBorderTop).Color = borderColour
                    drCells.Borders(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleSingle
                    drCells.Borders(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth050pt
                    drCells.Borders(WdBorderType.wdBorderBottom).Color = RGB(btnColor.R, btnColor.G, btnColor.B)
                Else
                    drCells.Borders(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleSingle
                    drCells.Borders(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth050pt
                    drCells.Borders(WdBorderType.wdBorderBottom).Color = RGB(btnColor.R, btnColor.G, btnColor.B)
                End If
                '
            Else
                drCells.Borders.InsideLineStyle = WdLineStyle.wdLineStyleNone
                drCells.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone

            End If
            '
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
        '
    End Function


    '
    Public Sub tbl_borders_colourAndVisibility(ByRef rng As Range, ByVal doBorders As Boolean, ByVal borderColour As Long)
        'This method will colour or make disappear the horizontal borders of the selection
        Dim dr As Row
        '
        On Error GoTo finis
        If rng.Rows.Count = 0 Then GoTo finis
        '
        If doBorders Then
            If rng.Rows.Count > 1 Then
                rng.Rows.Borders(WdBorderType.wdBorderHorizontal).LineStyle = WdLineStyle.wdLineStyleSingle
                rng.Rows.Borders(WdBorderType.wdBorderHorizontal).LineWidth = WdLineWidth.wdLineWidth050pt
                rng.Rows.Borders(WdBorderType.wdBorderHorizontal).Color = borderColour
                'rng.Rows.Borders(wdBorderTop).LineStyle = wdLineStyleSingle
                'rng.Rows.Borders(wdBorderTop).LineWidth = wdLineWidth050pt
                'rng.Rows.Borders(wdBorderTop).Color = borderColour
                rng.Rows.Borders(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleSingle
                rng.Rows.Borders(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth050pt
                rng.Rows.Borders(WdBorderType.wdBorderBottom).Color = borderColour
            Else
                rng.Rows.Borders(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleSingle
                rng.Rows.Borders(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth050pt
                rng.Rows.Borders(WdBorderType.wdBorderBottom).Color = borderColour
            End If
        Else

            rng.Rows.Borders.InsideLineStyle = WdLineStyle.wdLineStyleNone
            rng.Rows.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone

        End If
        Exit Sub
finis:
        MsgBox("Have you selected some rows? Please note that this Function only works On regular tables")

    End Sub
    '
    Public Sub tbl_rapid_Format(ByRef tbl As Word.Table)

        Try
            Me.tbl_fix_Table(tbl, True, RGB(255, 0, 0), True)

        Catch ex As Exception

        End Try

    End Sub
    '
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="doBorders"></param>
    ''' <param name="borderColour"></param>
    ''' <param name="doAACTableStyle"></param>
    ''' <param name="textStyleName"></param>
    ''' <param name="doTopBorder"></param>
    ''' <param name="doBottomBorder"></param>
    Public Sub tbl_fix_Table(ByRef tbl As Table, ByVal doBorders As Boolean, ByVal borderColour As Long, doAACTableStyle As Boolean, Optional textStyleName As String = "", Optional doTopBorder As Boolean = False, Optional doBottomBorder As Boolean = False)
        'This method will  the cell margins for the current Table to 0 and
        'the Borders on or off depending on the value of doBorders
        Dim brdr As Border
        Dim myDoc As Document

        '
        myDoc = tbl.Range.Document
        '
        Try
            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
            'tbl.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop

            'Call tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow)       'columns change size o accomodate text.. sets AllowAutoFit = true
            tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed)              'columns sizes don't change to accommodate text, will  AllowAutoFit = false
            'Call tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)      'Table width changes to accommodate text.. Not content, then width=0.. sets AlloAutoFit = true

            tbl.TopPadding = 0#
            tbl.BottomPadding = 0#
            tbl.LeftPadding = 0#
            tbl.RightPadding = 0#
            '
            tbl.AllowPageBreaks = True
            tbl.Rows.AllowBreakAcrossPages = False
            '
            tbl.Borders.InsideLineStyle = WdLineStyle.wdLineStyleNone
            tbl.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone
            tbl.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic
            '
            'The Table text style defaults to 'glb_style_tblTextStyle' (Table text), but will take
            'and use any text style name that is entered
            If textStyleName = "" Then
                textStyleName = glb_var_style_tblTextStyle
            End If
            '
            'If the text style doesn't work, then we go back to the default
            Try
                If doAACTableStyle Then tbl.Range.Style = Me.glb_get_wrdStyle(textStyleName)
            Catch ex2 As Exception
                textStyleName = glb_var_style_tblTextStyle
                If doAACTableStyle Then tbl.Range.Style = Me.glb_get_wrdStyle(textStyleName)
            End Try
            '
            If doBorders Then
                Try
                    '
                    'tbl.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone
                    'tbl.Borders.InsideLineStyle = WdLineStyle.wdLineStyleSingle
                    'tbl.Borders.InsideLineWidth = WdLineWidth.wdLineWidth050pt
                    'tbl.Borders.InsideColor = borderColour

                    tbl.Borders.Item(WdBorderType.wdBorderHorizontal).LineStyle = WdLineStyle.wdLineStyleSingle
                    tbl.Borders.Item(WdBorderType.wdBorderHorizontal).LineWidth = WdLineWidth.wdLineWidth050pt
                    tbl.Borders.Item(WdBorderType.wdBorderHorizontal).Color = borderColour
                    'tbl.Borders.Item(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleNone
                    'tbl.Borders.Item(WdBorderType.wdBorderTop).LineWidth = WdLineWidth.wdLineWidth050pt
                    'tbl.Borders.Item(WdBorderType.wdBorderTop).Color = borderColour
                    'tbl.Borders.Item(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleSingle
                    'tbl.Borders.Item(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth050pt
                    'tbl.Borders.Item(WdBorderType.wdBorderBottom).Color = borderColour
                    'tbl.Borders.Item(wdBorderLeft).LineStyle = wdLineStyleSingle
                    'tbl.Borders.Item(wdBorderLeft).LineWidth = wdLineWidth050pt
                    'tbl.Borders.Item(wdBorderLeft).Color = borderColour
                    'tbl.Borders.Item(wdBorderRight).LineStyle = wdLineStyleSingle
                    'tbl.Borders.Item(wdBorderRight).LineWidth = wdLineWidth050pt
                    'tbl.Borders.Item(wdBorderRight).Color = borderColour
                    '
                Catch ex2 As Exception

                End Try
                '
            Else
                For Each brdr In tbl.Borders
                    'brdr.LineStyle = Word.WdLineStyle.wdLineStyleNone
                Next brdr
            End If
            '
            If doTopBorder Then
                tbl.Borders.Item(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleSingle
                tbl.Borders.Item(WdBorderType.wdBorderTop).LineWidth = WdLineWidth.wdLineWidth050pt
                tbl.Borders.Item(WdBorderType.wdBorderTop).Color = borderColour
            End If
            '
            If doBottomBorder Then
                tbl.Borders.Item(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleSingle
                tbl.Borders.Item(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth050pt
                tbl.Borders.Item(WdBorderType.wdBorderBottom).Color = borderColour
            End If

        Catch ex As Exception
            MsgBox("Error - fixTables")

        End Try
        '
    End Sub
    '
    '
    Public Sub tbl_doBorders_MaintainPadding(ByRef tbl As Table, ByVal doBorders As Boolean, ByVal borderColour As Long)
        'This method will  the cell margins for the current Table to 0 and
        'the Borders on or off depending on the value of doBorders
        Dim brdr As Border
        Dim myDoc As Document
        Dim currentSect As Section
        Dim widthBetweenMargins As Single
        '
        On Error GoTo finis
        myDoc = Globals.ThisAddin.Application.ActiveDocument
        currentSect = Globals.ThisAddin.Application.Selection.Sections(1)
        '
        '
        'tbl.TopPadding = 0#
        'tbl.BottomPadding = 0#
        'tbl.leftPadding = 0#
        'tbl.RightPadding = 0#
        'tbl.PreferredWidth = widthBetweenMargins
        '
        'tbl.AllowAutoFit = True
        'Call tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow)    'columns change size o accomodate text.. sets AllowAutoFit = true
        'Call tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed)      'columns sizes don't change to accommodate text, will  AllowAutoFit = false
        'Call tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)   'Table width changes to accommodate text.. Not content, then width=0.. sets AlloAutoFit = true
        '
        tbl.AllowPageBreaks = True                                           'Will allow a row to break across pages
        If doBorders Then
            If tbl.Rows.Count > 1 Then
                'If there is only 1 row, then there are no horizontal borders
                tbl.Borders.Item(WdBorderType.wdBorderHorizontal).LineStyle = WdLineStyle.wdLineStyleSingle
                tbl.Borders.Item(WdBorderType.wdBorderHorizontal).LineWidth = WdLineWidth.wdLineWidth050pt
                tbl.Borders.Item(WdBorderType.wdBorderHorizontal).Color = borderColour
            End If
            tbl.Borders.Item(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleSingle
            tbl.Borders.Item(WdBorderType.wdBorderTop).LineWidth = WdLineWidth.wdLineWidth050pt
            tbl.Borders.Item(WdBorderType.wdBorderTop).Color = borderColour
            tbl.Borders.Item(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleSingle
            tbl.Borders.Item(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth050pt
            tbl.Borders.Item(WdBorderType.wdBorderBottom).Color = borderColour
            'tbl.Borders.Item(wdBorderLeft).LineStyle = wdLineStyleSingle
            'tbl.Borders.Item(wdBorderLeft).LineWidth = wdLineWidth050pt
            'tbl.Borders.Item(wdBorderLeft).Color = borderColour
            'tbl.Borders.Item(wdBorderRight).LineStyle = wdLineStyleSingle
            'tbl.Borders.Item(wdBorderRight).LineWidth = wdLineWidth050pt
            'tbl.Borders.Item(wdBorderRight).Color = borderColour

        Else
            For Each brdr In tbl.Borders
                brdr.LineStyle = Word.WdLineStyle.wdLineStyleNone
            Next brdr
        End If
        Exit Sub
        '
finis:
        MsgBox("Error - fixTables")
    End Sub
    '
    ''' <summary>
    ''' This method will determine whether the table has the same number of cells 
    ''' per row.. If not we will return false, that is, the table is not regular. This
    ''' test is not definitive, but will cover most circumstances
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tableIsRegular(ByRef tbl As Word.Table) As Boolean
        Dim dr As Word.Row
        Dim i, baseCount, cellCount As Integer
        Dim rslt As Boolean
        '

        rslt = True
        '
        'Get the base cell count... If there is a variation from this, the table is not regular
        '
        dr = tbl.Rows.Item(1)
        baseCount = dr.Range.Cells.Count

        For i = 1 To tbl.Rows.Count
            dr = tbl.Rows.Item(i)
            cellCount = dr.Range.Cells.Count
            If cellCount <> baseCount Then
                rslt = False
                Exit For
            End If
        Next
        '
        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method will merge all cells in the table (tbl) that have the specified 
    ''' row index
    ''' </summary>
    ''' <param name="rowIdx"></param>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_merge_cellsWithRowIndex(rowIdx As Integer, ByRef tbl As Word.Table) As Word.Range
        Dim drCell, rsltCell, drCellLast As Word.Cell
        Dim rng As Word.Range
        Dim j As Integer
        '
        drCell = Nothing
        rsltCell = Nothing
        rng = Nothing
        '
        '
        Try
            '**** Work Around to row size error
            'Note if I first set the first row to dr and then operate on it I get the
            'resising problem, buth this time not in the body, but in the first row. If
            'I just refer directly then all is OK
            '
            '
            'tbl = tblBody
            drCellLast = tbl.Range.Cells.Item(tbl.Range.Cells.Count)
            drCell = tbl_get_firstCellWithRowIndex(rowIdx, tbl)
            rng = drCell.Range
            '
            For j = 1 To 1000
                'Test to make sure we are not going beyond the bounds of the table
                If Not IsNothing(drCell.Next) Then
                    drCell = drCell.Next
                    If drCell.RowIndex = rowIdx Then
                        rng.MoveEnd(WdUnits.wdCell, 1)
                    Else
                        Exit For
                    End If
                Else
                    Exit For
                End If
            Next
            '
            If rowIdx = 1 Then
                j = 1
            Else
                j = 2
            End If
            '
            If rng.Cells.Count > 1 Then
                rng.Select()
                glb_get_wrdSel.Cells.Merge()
                '
                Me.tbl_colour_set_colourOfCellsToNone(rng.Cells)
                '
                'tbl.PreferredWidthType = tblPreferredWidthType
                'tbl.PreferredWidth = tblPreferredWidth
                '
                'rng.Cells.Merge()
            End If
            '
        Catch ex As Exception
            rng = Nothing
        End Try
finis:
        Return rng
        '
    End Function
    '
    '
    Public Function tbl_get_firstCellWithColumnIndex(columnIndex As Integer, ByRef tbl As Word.Table) As Word.Cell
        Dim drCell, rsltCell As Word.Cell
        Dim j As Integer
        '
        drCell = Nothing
        rsltCell = Nothing
        '
        If columnIndex <= 0 Then

        End If
        '
        Try
            If columnIndex >= 1 Then
                For j = 1 To tbl.Range.Cells.Count
                    drCell = tbl.Range.Cells.Item(j)
                    If drCell.ColumnIndex = columnIndex Then
                        rsltCell = tbl.Range.Cells.Item(j)
                        Exit For
                    End If
                Next
            Else
                drCell = Nothing
            End If
            '
        Catch ex As Exception
            rsltCell = Nothing
        End Try

        Return rsltCell
        '
    End Function
    '
    Public Function tbl_getSelect_firstCellWithColumnIndex(columnIndex As Integer, ByRef tbl As Word.Table) As Word.Cell
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        '
        drCell = Me.tbl_get_firstCellWithColumnIndex(columnIndex, tbl)
        rng = drCell.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()
        '
        Return drCell
        '
    End Function
    '
    Public Function tbl_get_firstCellWithRowIndex(rowIdx As Integer, ByRef tbl As Word.Table) As Word.Cell
        Dim drCell, rsltCell As Word.Cell
        Dim j As Integer
        '
        drCell = Nothing
        rsltCell = Nothing
        '
        Try
            For j = 1 To tbl.Range.Cells.Count
                drCell = tbl.Range.Cells.Item(j)
                If drCell.RowIndex = rowIdx Then
                    rsltCell = tbl.Range.Cells.Item(j)
                    Exit For
                End If
            Next
            '
        Catch ex As Exception
            rsltCell = Nothing
        End Try

        Return rsltCell
        '
    End Function
    '

    '
    '
    Public Function tbl_get_secondCellWithColumnIndex(columnIndex As Integer, ByRef tbl As Word.Table) As Word.Cell
        Dim drCell, rsltCell As Word.Cell
        Dim j, kount As Integer
        '
        drCell = Nothing
        rsltCell = Nothing
        kount = 0
        '
        Try
            For j = 1 To tbl.Range.Cells.Count
                drCell = tbl.Range.Cells.Item(j)
                If drCell.ColumnIndex = columnIndex Then
                    rsltCell = tbl.Range.Cells.Item(j)
                    If kount = 0 Then
                        kount = kount + 1
                    Else
                        Exit For
                    End If
                End If
            Next
            '
        Catch ex As Exception
            rsltCell = Nothing
        End Try

        Return rsltCell
        '
    End Function
    '
    '
    Public Function tbl_aacTable_ColumnDelete(ByRef tbl As Word.Table) As Integer
        Dim objMsgMgr As New cMessageManager()
        Dim objTblsMgr As New cTablesMgr()
        Dim lstOfMeasurements As New Collection()
        Dim drCell As Word.Cell
        Dim drCol As Word.Column
        Dim columnIndex As Integer
        Dim tblPreferredWidth As Single
        '
        drCol = Nothing
        columnIndex = -1
        '
        Try
            'If strMsg is null then we are not in any forbidden areas. So
            'now we just need to check whether we are in a Table
            'tbl = glb_get_wrdSelTbl()
            '
            'If Me.isInBoxFigureRec() Then
            If Me.tbl_is_EncapsulatedTable(tbl) Then
                'tbl = objTblsMgr.glb_get_wrdSelTbl
                tblPreferredWidth = Me.tbl_Encaps_getWidth(tbl, lstOfMeasurements)
                '
                'Get the currently selected cell. If it exists then operate on the tbale. If it is
                'not then return an error
                '
                drCell = glb_get_wrdSelCell()
                '
                If Not IsNothing(drCell) Then
                    columnIndex = drCell.ColumnIndex
                    columnIndex = objTblsMgr.tbl_Encaps_columnDelete_doSplit(columnIndex, tbl)
                Else
                    columnIndex = -1
                End If
                '
            Else
                'tbl = objTblsMgr.glb_get_wrdSelTbl
                '
                '*** Need to simplify/rewrite in line with new table structures
                '*** It is only used here
                '
                columnIndex = objTblsMgr.tbl_column_Delete(tbl)
                tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                tbl.PreferredWidth = 100
                'objMsgMgr.IsInBoxFigureRec()
                '
            End If
            '
            'tbl.PreferredWidthType = tblPreferredWidthType
            '
            'If Not (tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthAuto) Then
            'tbl.PreferredWidth = tblPreferredWidth
            'End If
            '
        Catch ex As Exception
            columnIndex = -1
        End Try
        '
        Return columnIndex
        '
    End Function
    '
    ''' <summary>
    ''' This method will accept a column action command (strColumnAction = 'left', 'right' and 'delete'), and if
    ''' 'left' will insert a column to the left of the current selection. If 'right' it will insert a column to the
    ''' right of the current selection, and if 'delete' will delete the column containg the current selection. This method
    ''' will determine if the table (tbl) is Encapsulated or not and adjust it's behaviour accordingly.
    ''' </summary>
    ''' <param name="strColumnAction"></param>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_aacTable_ColumnInsertDelete(strColumnAction As String, ByRef tbl As Word.Table) As Integer
        Dim objMsgMgr As New cMessageManager()
        Dim objTblsMgr As New cTablesMgr()
        Dim lstOfMeasurements As New Collection()
        Dim drCol As Word.Column
        Dim drCell As Word.Cell
        Dim columnIndex As Integer
        '
        drCol = Nothing
        columnIndex = -1
        '
        '
        Try
            'If strMsg is null then we are not in any forbidden areas. So
            'now we just need to check whether we are in a Table
            'tbl = glb_get_wrdSelTbl()
            '
            'If Me.isInBoxFigureRec() Then
            If Me.tbl_is_EncapsulatedTable(tbl) Or Me.tbl_is_EncapsulatedFigure(tbl) Or Me.tbl_is_EncapsulatedBox(tbl) Then
                '
                columnIndex = objTblsMgr.tbl_Encaps_columnInsertDelete_doSplit_New(strColumnAction, tbl)
                '*** This older version is left here as a reference
                'columnIndex = objTblsMgr.tbl_Encaps_columnInsert_doSplit(strColumnAction, tbl)
                '
            Else
                'If Me.tbl_is_tblStandard(tbl) Then
                '
                tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                tbl.Columns.PreferredWidth = 100
                tbl.PreferredWidth = 100
                '
                Select Case strColumnAction
                    Case "left"
                        glb_get_wrdSel.InsertColumns()
                        drCell = glb_get_wrdSelCell()
                        If Not IsNothing(drCell) Then
                            columnIndex = drCell.ColumnIndex
                        Else
                            columnIndex = -1
                        End If
                    Case "right"
                        glb_get_wrdSel.InsertColumnsRight()
                        drCell = glb_get_wrdSelCell()
                        If Not IsNothing(drCell) Then
                            columnIndex = drCell.ColumnIndex
                        Else
                            columnIndex = -1
                        End If
                    Case "delete"
                        glb_get_wrdSel.Columns.Delete()
                        drCell = glb_get_wrdSelCell()
                        If Not IsNothing(drCell) Then
                            columnIndex = drCell.ColumnIndex
                        Else
                            columnIndex = -1
                        End If

                End Select
                '
                tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                tbl.PreferredWidth = 100
                '
            End If
            '
        Catch ex As Exception
            columnIndex = -1
        End Try
        '
        '
        Return columnIndex
        '
    End Function
    '
    ''' <summary>
    ''' This method supports the 'Insert Right' and 'Insert Left' functiosn that
    ''' have been added to allow users to easily add table columns to the standard AAC Table
    ''' </summary>
    ''' <returns></returns>
    Public Function IsOKToAddTableColumn() As String
        Dim objCpMgr As cCoverPageMgr
        Dim objContactsMgr As New cContactsMgr()
        Dim objTOCMgr As New cTOCMgr()
        Dim objTagsMgr As New cTagsMgr()
        Dim objDivMgr As New cChptDivider()
        Dim strErrorMsg As String
        Dim sect As Word.Section
        '
        objCpMgr = New cCoverPageMgr
        sect = Me.glb_get_wrdSect()
        strErrorMsg = ""
        '
        If objCpMgr.cp_Bool_IsCoverPage(sect) Then strErrorMsg = "This function is not supported in the Cover Page"
        If objTOCMgr.toc_is_TOCSection(sect) Then strErrorMsg = "This function is not supported in the Table of Contents"
        If objContactsMgr.is_ContactsPage_Front(sect) Then strErrorMsg = "This function is not supported in the Front Contacts page"
        If objContactsMgr.is_ContactsPage_Back(sect) Then strErrorMsg = "This function is not supported in the back Contacts Page"
        If objDivMgr.is_divider_Any(sect) Then strErrorMsg = "This function is not supported in a Divider"
        'If objSectMgr.isInBanner() Then strErrorMsg = "AAC Table 'Insert Right/Left' is not supported in a 'Heading Banner'"
        'If objTagsMgr.tags_is_inBanner() Then strErrorMsg = "AA Table 'Insert Right/Left' is not supported in a 'Heading Banner'"
        '
        '
        Return strErrorMsg

    End Function
    '
    ''' <summary>
    ''' This method is to be called bfore the user is given the option to undo a the Table
    ''' column insert left/right function... We don't want to accidently overwrite a banner
    ''' or some other table based structure
    ''' </summary>
    ''' <returns></returns>
    Public Function tbl_isOK_ToPasteOverTable() As String
        Dim objCpMgr As cCoverPageMgr
        Dim objContactsMgr As New cContactsMgr()
        Dim objTOCMgr As New cTOCMgr()
        Dim objChptBase As New cChptBanner()
        Dim objTagsMgr As New cTagsMgr()
        Dim strErrorMsg As String
        Dim sect As Word.Section
        Dim objGlobals As New cGlobals()
        '
        objCpMgr = New cCoverPageMgr
        sect = objGlobals.glb_get_wrdSect()
        strErrorMsg = ""
        '
        If objCpMgr.cp_Bool_IsCoverPage(sect) Then strErrorMsg = "This function is not supported in a Cover Page"
        If objTOCMgr.toc_is_TOCSection(sect) Then strErrorMsg = "This function is not supported in the TOC"
        If objContactsMgr.is_ContactsPage_Front(sect) Then strErrorMsg = "This function is not supported in a Contacts Page"
        If objContactsMgr.is_ContactsPage_Back(sect) Then strErrorMsg = "This function is not supported in a Contacts Page"
        'If objTagsMgr.tags_is_inBanner() Then strErrorMsg = "This function is not supported in a Headng Banner"
        If Me.isInBoxFigureRec() Then strErrorMsg = "This function is not supported in Figures or Boxes"
        '
        '
        Return strErrorMsg
    End Function
    ''' <summary>
    ''' This method will return a collection of measurements for an Encapsulated Table. Values as 
    ''' Single can be retrieved as (tblWidth_Pts, "widthInPts"), (tblWidthPercent, "widthInPercent")
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_Encaps_getWidth(ByRef tbl As Word.Table, ByRef lstOfMeasurements As Collection) As Single
        Dim sect As Word.Section
        Dim tblWidth_Pts, tblWidthPercent As Single
        '
        sect = tbl.Range.Sections.Item(1)

        tblWidth_Pts = glb_get_wrdSel.Tables.Item(1).Range.Cells.Item(1).PreferredWidth
        '
        Select Case sect.PageSetup.TextColumns.Count
            Case 1
                tblWidthPercent = 100 * tblWidth_Pts / sect.PageSetup.TextColumns.Item(1).Width
            Case 2
                tblWidthPercent = 100
            Case 3
                tblWidthPercent = 100
            Case 4
                tblWidthPercent = 100
            Case Else
                tblWidthPercent = 100
        End Select
        '
        lstOfMeasurements.Add(tblWidth_Pts, "widthInPts")
        lstOfMeasurements.Add(tblWidthPercent, "widthInPercent")
        '
        Return tblWidth_Pts
        '
    End Function
    '
    ''' <summary>
    ''' This method will assume that tbl is an AA Encapsulated TAble and it will split the top
    ''' row/cell away as tblTop and return the paragraph between tblTop and the rest of the 
    ''' Table which is now tbl.. Note that because we use Selection.SplitTable this method
    ''' doesn't care whether the table is regular or not.
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="tblTop"></param>
    ''' <param name="splitParaTop"></param>
    ''' <returns></returns>
    Public Function tbl_Encaps_splitTop(ByRef tbl As Word.Table, ByRef tblTop As Word.Table, ByRef splitParaTop As Word.Paragraph) As Word.Table
        Dim tblBody As Word.Table
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        '
        '
        tblTop = Nothing
        tblBody = Nothing
        '
        splitParaTop = Nothing
        'splitParaBottom = Nothing'
        Try
            'Verified 20240724.. We use the more flexible selection.splitTable
            '
            'First split from a cell in the second row and get the splitParaTop and tblTop
            drCell = tbl.Range.Cells.Item(2)
            rng = drCell.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Select()
            glb_get_wrdSel.SplitTable()
            splitParaTop = glb_get_wrdSel.Paragraphs.Item(1)
            '
            rng = splitParaTop.Range
            rng.Move(WdUnits.wdParagraph, -1)
            tblTop = rng.Tables.Item(1)
            '
            rng = splitParaTop.Range
            rng.Move(WdUnits.wdParagraph, 1)
            tbl = rng.Tables.Item(1)

        Catch ex As Exception
            tbl = Nothing
        End Try
        '
        '
        Return tbl
        '
    End Function
    '
    ''' <summary>
    ''' This method assumes that tbl is an AA Encapsulated Table. It will split the table returning the top
    ''' rwo as tblTop, and the splitParaTop (the paragraph that separates the tblTop and tbBody). It also
    ''' returns tblBody (middle body of the Table), the botom row/cell as tblBottom and the paragraph
    ''' between (splitParaBottom). On exit tbl is set to tblBody.. Note that because we use 
    ''' Selection.SplitTable this method doesn't care whether the table is regular or not... If there is no Source/Note cell
    ''' then tblBottom=Nothing and spliParaBottom=Nothing
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="tblTop"></param>
    ''' <param name="tblBottom"></param>
    ''' <param name="splitParaTop"></param>
    ''' <param name="splitParaBottom"></param>
    ''' <returns></returns>
    Public Function tbl_Encaps_split(ByRef tbl As Word.Table, ByRef tblTop As Word.Table, ByRef tblBottom As Word.Table, ByRef splitParaTop As Word.Paragraph, ByRef splitParaBottom As Word.Paragraph) As Word.Table
        Dim tblBody As Word.Table
        Dim rng As Word.Range
        Dim drCell As Word.Cell
        Dim srcStyle As Word.Style
        '
        '
        tblTop = Nothing
        tblBody = Nothing
        tblBottom = Nothing
        '
        splitParaTop = Nothing
        splitParaBottom = Nothing
        '
        Try
            'Verified 20240724.. We use the more flexible selection.splitTable
            '
            'First split from a cell in the second row and get the splitParaTop and tblTop
            drCell = tbl.Range.Cells.Item(2)
            rng = drCell.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Select()
            glb_get_wrdSel.SplitTable()
            splitParaTop = glb_get_wrdSel.Paragraphs.Item(1)
            '
            rng = splitParaTop.Range
            rng.Move(WdUnits.wdParagraph, -1)
            tblTop = rng.Tables.Item(1)
            '
            '
            'Now split away the bottom row/cell and get splitParaBottom,
            'tblBody and tblBottom
            '
            rng = splitParaTop.Range
            rng.Move(WdUnits.wdParagraph, 1)
            '
            tbl = rng.Tables.Item(1)
            '
            drCell = tbl.Range.Cells.Item(tbl.Range.Cells.Count)
            srcStyle = drCell.Range.Style
            '
            'If no source/note row/cell, then splitParaBottom = nothing
            'and tblBottom = nothing
            '
            If srcStyle.NameLocal = "Source" Or srcStyle.NameLocal = "Note" Then
                rng = drCell.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                '
                rng.Select()
                '
                glb_get_wrdSel.SplitTable()
                splitParaBottom = glb_get_wrdSel.Paragraphs.Item(1)
                '
                rng = glb_get_wrdSelRng()
                rng.Move(WdUnits.wdParagraph, -1)
                rng.Select()
                tblBody = glb_get_wrdSelTbl()
                '
                '
                rng = splitParaBottom.Range
                rng.Move(WdUnits.wdParagraph, 1)
                tblBottom = rng.Tables.Item(1)
                '
                'Make certaint he bottom row/cell has no colour
                Me.tbl_colour_set_colourOfCellsToNone(tblBottom.Range.Cells)
                '
                tbl = tblBody

            Else
                tblBody = tbl
                splitParaBottom = Nothing
                tblBottom = Nothing
            End If
            '
        Catch ex As Exception

        End Try
        '
        '
        '
        Return tblBody
        '
    End Function
    '
    ''' <summary>
    ''' This method will assume that tbl is an AA Encapsulated TAble and it will split the bottom
    ''' row/cell away as tblBottom and return the paragraph between tblBottom and the rest of the 
    ''' Table which is now tbl.. Note that because we use Selection.SplitTable this method
    ''' doesn't care whether the table is regular or not.
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="tblBottom"></param>
    ''' <param name="splitParaBottom"></param>
    ''' <returns></returns>
    Public Function tbl_Encaps_splitBottom(ByRef tbl As Word.Table, ByRef tblBottom As Word.Table, ByRef splitParaBottom As Word.Paragraph) As Word.Table
        Dim tblBody As Word.Table
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        '
        '
        tblBottom = Nothing
        tblBody = Nothing
        '
        splitParaBottom = Nothing
        'splitParaBottom = Nothing'
        Try
            'Verified 20240724.. We use the more flexible selection.splitTable
            drCell = tbl.Range.Cells.Item(tbl.Range.Cells.Count)
            rng = drCell.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            rng.Select()
            '
            glb_get_wrdSel.SplitTable()
            splitParaBottom = glb_get_wrdSel.Paragraphs.Item(1)
            '
            rng = glb_get_wrdSelRng()
            rng.Move(WdUnits.wdParagraph, -1)
            rng.Select()
            tblBody = glb_get_wrdSelTbl()
            '
            '
            rng = splitParaBottom.Range
            rng.Move(WdUnits.wdParagraph, 1)
            tblBottom = rng.Tables.Item(1)
            '
            Me.tbl_colour_set_colourOfCellsToNone(tblBottom.Range.Cells)
            '
            tbl = tblBody

        Catch ex As Exception
            tbl = Nothing
        End Try
        '
        '
        Return tbl
        '
    End Function
    '
    ''' <summary>
    ''' This method will return true if the paragraph after the table is either
    ''' a Source or Note paragraph
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_has_SourceNotePara(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim myDoc As Word.Document
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        Dim stylSrc, stylNote, paraStyle As Word.Style
        '
        rslt = False
        '
        myDoc = tbl.Range.Document()
        stylSrc = myDoc.Styles.Item("Source")
        stylNote = myDoc.Styles.Item("Note")
        '
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
        'Set para to the first paragraph after the table
        para = rng.Paragraphs.Item(1)
        paraStyle = para.Range.Style
        rng = para.Range
        If paraStyle.NameLocal = stylSrc.NameLocal Or paraStyle.NameLocal = stylNote.NameLocal Then rslt = True
        '
        Return rslt
        '
    End Function

    '
    ''' <summary>
    ''' This method will return true if the paragraph above the table (not in the table) is
    ''' a Caption
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_has_topCaption(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim rng As Word.Range
        Dim styl As Word.Style
        '
        rslt = False
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Move(WdUnits.wdCharacter, -1)
        styl = rng.Style
        '
        If styl.NameLocal = "Caption" Then rslt = True
        '
        Return rslt
    End Function
    '
    Public Function tbl_build_topRowAsCells(ByRef tbl As Word.Table) As String
        Dim rngCaption, rngCell, rngOfCells As Word.Range
        Dim drCell As Word.Cell
        Dim brdr As Word.Border
        Dim rng As Word.Range
        Dim strErrorMsg As String
        '
        strErrorMsg = ""
        Try
            'First build the top row and capture the 'Caption' paragraph, then get the range
            'of the following 'Source and Note' paragraphs and then get the text string. This
            'will either by null or some value
            'rng = tbl.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'rng.Move(WdUnits.wdCharacter, -1)
            'styl = rng.Style
            'Build the top row only if the para above the table uses the 'Caption' style

            If Me.tbl_has_topCaption(tbl) Then
                '
                Try
                    'First get the caption, then add the new top row using Selection as it is
                    'tolerant of merged tables
                    '
                    rng = tbl.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    rng.Move(WdUnits.wdCharacter, -1)
                    rngCaption = rng.Paragraphs.Item(1).Range
                    rngCaption.Cut()
                    '
                    rng = tbl.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    rng.Select()
                    glb_get_wrdSel().InsertRowsAbove(1)
                    '
                    'Get the cells of the first row and merge them
                    rngOfCells = Me.tbl_get_RowCells(1, tbl)
                    rngOfCells.Cells.Merge()
                    '
                    drCell = tbl.Range.Cells.Item(1)
                    drCell.BottomPadding = 0
                    drCell.TopPadding = 0
                    drCell.LeftPadding = 0
                    drCell.RightPadding = 0
                    '
                    brdr = drCell.Borders.Item(WdBorderType.wdBorderBottom)
                    brdr.LineStyle = WdLineStyle.wdLineStyleNone
                    rngCell = drCell.Range
                    rngCell.Style = glb_get_wrdActiveDoc.Styles("Caption")
                    Me.tbl_colour_set_colourOfCellToNone(drCell)
                    rngCell.Collapse(WdCollapseDirection.wdCollapseStart)
                    '
                    rngCell.Paste()
                    '
                    'Me.tbl_delete_ParaAtEndOfTable(tbl)
                    Me.tbl_delete_ParaAtTopOfTable(tbl)
                    '
                    Clipboard.Clear()
                    '
                Catch ex2 As Exception

                End Try

            Else
                strErrorMsg = "No Caption at the top of the table."
            End If

        Catch ex As Exception
            strErrorMsg = "Unknown error.. Try making your table more 'regular'"
        End Try
        '
        '
        Return strErrorMsg
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will take a standard AA table with Caption and Source and it will
    ''' encapsulate both Caption and Source/Notes in their own rows, building
    ''' an 'Encapsulated' table.
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_convert_tblStdToEncaps(ByRef tbl As Word.Table) As String
        Dim myDoc As Word.Document
        Dim objParas As New cParas()
        Dim rng As Word.Range
        Dim strSrc, strErrorMsg As String
        '
        strSrc = ""
        strErrorMsg = ""
        myDoc = tbl.Range.Document
        '
        Try
            If Me.tbl_is_tblStandard(tbl) Then
                '
                '****
                'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                'tbl.PreferredWidth = 100
                'tbl.Columns.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                'tbl.Columns.PreferredWidth = 100
                '****
                '
                'First build the top row and capture the 'Caption' paragraph, then get the range
                'of the following 'Source and Note' paragraphs and then get the text string. This
                'will either by null or some value
                '
                strErrorMsg = Me.tbl_build_topRowAsCells(tbl)
                'Me.tbl_build_topRow(tbl)
                '
                'Now deal with the Source and Note, by first getting the Source/Note rnage and the
                'associated text. Then build the row and add the text, delete any extraneous paragraphs
                'in the new row (after text insertion). Finally, delete the Source/Note paragraphs
                'that were at the end of the table
                '
                rng = Me.tbl_get_RangeOfParasSourceAndNoteAtEndOfTable(tbl)
                If Not IsNothing(rng) Then
                    strSrc = rng.Text
                    'Now build the new source row
                    'dr = Me.tbl_build_sourceRow(tbl, "none")
                    rng = Me.tbl_build_sourceRowAsCells(tbl, "none")
                    '
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    rng.Text = strSrc
                    '
                    objParas.paras_delete_lastParasInTableCell(rng.Cells.Item(1))
                    '
                    Me.tbl_delete_ParasSourceAndNoteAtEndOfTable(tbl)
                    rng = tbl.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                    'objParas.paras_insert_numParas(rng, 1)
                    '
                Else
                    If strErrorMsg = "" Then
                        strErrorMsg = "No Source/Note at the end of the table."
                    Else
                        strErrorMsg = strErrorMsg + vbCrLf + "and no Source/Note at the end of the table."
                    End If
                End If
                '
            Else
                'MsgBox("This is not a Standard Table")
            End If
            '
        Catch ex As Exception

        End Try

        Return strErrorMsg
    End Function
    '
    Public Function tbl_convert_tblEncapsToStd(ByRef tbl As Word.Table) As Word.Table
        Dim myDoc As Word.Document
        Dim rng, rngBottom As Word.Range
        Dim styl As Word.Style
        Dim strErrorMsg As String
        Dim tblTop, tblBottom, tblBody As Word.Table
        Dim splitParaTop, splitParaBottom As Word.Paragraph
        Dim drCell As Word.Cell
        '
        '
        myDoc = tbl.Range.Document
        strErrorMsg = ""
        tblTop = Nothing
        tblBody = Nothing
        tblBottom = Nothing
        rngBottom = Nothing
        '
        splitParaTop = Nothing
        splitParaBottom = Nothing
        '
        Try
            drCell = tbl.Range.Cells.Item(1)
            rng = drCell.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            styl = rng.Style
            '
            tblBody = Me.tbl_Encaps_split(tbl, tblTop, tblBottom, splitParaTop, splitParaBottom)
            rng = tblTop.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            '
            If Not IsNothing(tblTop) Then
                If styl.NameLocal = "Caption" Then
                    tblTop.ConvertToText()
                    splitParaTop.Range.Delete()
                Else
                    strErrorMsg = "The first cell does not contain a valid Caption."
                End If
            End If
            '
            If Not IsNothing(tblBottom) Then
                rngBottom = tblBottom.Range
                rngBottom.Collapse(WdCollapseDirection.wdCollapseStart)
                styl = rngBottom.Style
                If styl.NameLocal = "Source" Or styl.NameLocal = "Note" Then
                    tblBottom.ConvertToText()
                    splitParaBottom.Range.Delete()
                Else
                    strErrorMsg = "The last cell/row does not contain a valid Source/Note."
                End If
            Else

            End If
            '

        Catch ex As Exception
            tblBody = Nothing
        End Try
        '
        '
        Return tbl
        '
    End Function
    '
    ''' <summary>
    ''' This method will return a rnage of Cells in the table tbl that have the specified
    ''' rowidx. We do it this way to allow the table to have any merged cell structure
    ''' </summary>
    ''' <param name="rowIdx"></param>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_get_RowCells(rowIdx As Integer, ByRef tbl As Word.Table) As Word.Range
        Dim drCell As Word.Cell
        Dim rngOfCells As Word.Range
        '
        'Set a nominal starting point
        rngOfCells = Nothing
        '
        Try
            For Each drCell In tbl.Range.Cells
                If drCell.RowIndex = rowIdx Then
                    If drCell.ColumnIndex = 1 Then
                        rngOfCells = drCell.Range
                    Else
                        rngOfCells.MoveEnd(WdUnits.wdCell, 1)
                    End If
                End If
            Next
        Catch ex As Exception
            rngOfCells = Nothing
        End Try
        '
        Return rngOfCells
        '
    End Function
    '
    ''' <summary>
    ''' This function will detremine whether the current selection is in the last cell or first cell of
    ''' the table tbl.. If as string ="" is returned then the selection is in the body of the table
    ''' (i.e. not the first or not the last cell
    ''' </summary>
    ''' <param name="drCellSelected"></param>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_Encaps_SelectedIsFirstOrLastCell(ByRef drCellSelected As Word.Cell, ByRef tbl As Word.Table) As String
        Dim strMsg As String
        Dim drCellLast As Word.Cell
        '
        strMsg = ""
        '
        Try
            '
            'drCell can't be the first cell and can't be the last cell
            If Not IsNothing(drCellSelected) Then
                If drCellSelected.ColumnIndex = 1 And drCellSelected.RowIndex = 1 Then
                    If Me.tbl_is_EncapsulatedTable(tbl) Then
                        strMsg = "First cell selected. Move your cursor to the body of the table"
                    Else
                        strMsg = ""
                    End If
                Else
                    drCellLast = tbl.Range.Cells.Item(tbl.Range.Cells.Count)
                    If drCellSelected.ColumnIndex = drCellLast.ColumnIndex And drCellSelected.RowIndex = drCellLast.RowIndex Then
                        If Me.tbl_is_EncapsulatedTable(tbl) Then
                            strMsg = "Last cell selected. Move your cursor to the body of the table"
                        Else
                            strMsg = ""
                        End If
                    Else
                        'Everythig is OK
                        strMsg = ""
                    End If
                End If
                '
            Else
                strMsg = "Error. Try placing your cursor in the body of the table"
            End If
        Catch ex As Exception
            strMsg = "Unknown error in cTablesMgr.tbl_Encaps_SelectedIsFirstOrLastCell"
        End Try

        Return strMsg
    End Function
    '
    '
    Public Function tbl_Encaps_columnDelete_doSplit(currentColumnIdx As Integer, ByRef tbl As Word.Table) As Integer
        Dim tblTop, tblBottom As Word.Table
        Dim splitParaTop, splitParaBottom As Word.Paragraph
        Dim tblPreferredWidth As System.Single
        'Dim drCell As Word.Cell
        Dim drCol As Word.Column
        Dim newColumnIdx As Integer
        'Dim sect As Word.Section
        Dim wasRightAligned As Boolean
        '
        wasRightAligned = False
        '
        drCol = Nothing
        tblTop = Nothing
        tblBottom = Nothing
        splitParaTop = Nothing
        splitParaBottom = Nothing
        '
        If Not IsNothing(tbl) Then
            'sect = tbl.Range.Sections.Item(1)
            '
            'If tbl.Rows.Alignment = WdRowAlignment.wdAlignRowRight Then
            'wasRightAligned = True
            'tbl.Rows.Alignment = WdRowAlignment.wdAlignRowLeft
            'End If
            '
            Try
                '
                '
                Select Case tbl.PreferredWidthType
                    Case WdPreferredWidthType.wdPreferredWidthPercent
                        '
                        tblPreferredWidth = tbl.PreferredWidth
                        tbl = tbl_Encaps_split(tbl, tblTop, tblBottom, splitParaTop, splitParaBottom)
                        '
                        If tblPreferredWidth > 1000 Then
                            tblPreferredWidth = tblTop.PreferredWidth
                        End If
                        '
                        newColumnIdx = Me.tbl_Encaps_columnDelete_ifPercent(currentColumnIdx, tbl, tblTop, tblBottom, splitParaTop, splitParaBottom, tblPreferredWidth)
                        '
                    Case WdPreferredWidthType.wdPreferredWidthPoints
                        '
                        '*** At this point we ould leave the table in points and process it through Me.tbl_Encaps_insertColumn_ifPts.
                        '*** Or we can convert to percentage and process it through Me.tbl_Encaps_insertColumn_ifPercent.. This
                        '*** version converts the table to percent
                        '
                        'tblPreferredWidth = tbl.PreferredWidth
                        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                        'The PreferredWidth may be undefined because of the table structure, but we'll try
                        'anyway and correct it if its undefined (=9999999)
                        '
                        tblPreferredWidth = tbl.PreferredWidth
                        '
                        'tbl.PreferredWidth = 200
                        tbl = tbl_Encaps_split(tbl, tblTop, tblBottom, splitParaTop, splitParaBottom)
                        '
                        'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                        'tbl.PreferredWidth = 200
                        'tblPreferredWidth_Pts = tbl.PreferredWidth
                        'If tblPreferredWidth_Pts > 1000 Then
                        'tblPreferredWidth_Pts = tblTop.PreferredWidth
                        'End If
                        '
                        'If tblPreferredWidth is undefined (=9999999), then st a specific value obtained
                        'from the tblTop (Table Top)
                        If tblPreferredWidth >= 1000 Then tblPreferredWidth = tblTop.PreferredWidth

                        '
                        'If tblPreferredWidth = 9999999 Then
                        'Because the width is undefined (the table is not uniform) we need to get the measuremnt
                        'from the top/bottom row
                        'tblPreferredWidth = tblTop.PreferredWidth
                        'End If

                        '
                        'newColumnIdx = Me.tbl_Encaps_insertColumn_ifPts(strInsertDirection, currentColumnIdx, tbl, tblTop, tblBottom, splitParaTop, splitParaBottom, tblPreferredWidth)
                        newColumnIdx = Me.tbl_Encaps_columnDelete_ifPercent(currentColumnIdx, tbl, tblTop, tblBottom, splitParaTop, splitParaBottom, tblPreferredWidth)
                        '
                        '
                    Case WdPreferredWidthType.wdPreferredWidthAuto
                        '
                    Case Else
                        tblPreferredWidth = 433
                End Select
                '
                '
                'tbl = tbl_convert_tblEncapsToStd(tbl)
                '

            Catch ex As Exception
                currentColumnIdx = -1
                'MsgBox("cTablesMgr.3728")
            End Try


        End If
        '
        '
        Return newColumnIdx
        '
    End Function
    '
    ''' <summary>
    ''' This method will accept a table (tbl) and a 'column action' direction, strColumnAction. This direction
    ''' can be 'left' to insert column left from the current selection, 'right' to insert a column to the right
    ''' of the current selection and it can be 'delete' which will delete the column containing the current selection
    ''' </summary>
    ''' <param name="strColumnAction"></param>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_Encaps_columnInsertDelete_doSplit_New(strColumnAction As String, ByRef tbl As Word.Table) As Integer
        Dim tblTop, tblBottom As Word.Table
        Dim splitParaTop, splitParaBottom As Word.Paragraph
        Dim tblPreferredWidth As System.Single
        Dim drCell As Word.Cell
        Dim drCol As Word.Column
        Dim currentColumnIdx, newColumnIdx As Integer
        Dim sect As Word.Section
        Dim wasRightAligned As Boolean
        Dim rng As Word.Range
        '
        wasRightAligned = False
        '
        drCol = Nothing
        tblTop = Nothing
        tblBottom = Nothing
        splitParaTop = Nothing
        splitParaBottom = Nothing
        '
        If Not IsNothing(tbl) Then
            drCell = glb_get_wrdSelRng.Cells.Item(1)
            currentColumnIdx = drCell.ColumnIndex
            sect = tbl.Range.Sections.Item(1)
            '
            Try
                '
                Select Case tbl.PreferredWidthType
                    Case WdPreferredWidthType.wdPreferredWidthAuto
                        'MsgBox("tbl.PreferredWidthType = " + "Auto")
                    Case WdPreferredWidthType.wdPreferredWidthPercent
                        'MsgBox("tbl.PreferredWidthType = " + "Percent")
                    Case WdPreferredWidthType.wdPreferredWidthPoints
                        'MsgBox("tbl.PreferredWidthType = " + "Points")
                End Select
                '
                tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                '
                tbl = tbl_Encaps_split(tbl, tblTop, tblBottom, splitParaTop, splitParaBottom)
                '
                tblTop.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                tblTop.Columns.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                tblBottom.PreferredWidthType = tblTop.PreferredWidthType
                tblBottom.Columns.PreferredWidthType = tblTop.PreferredWidthType

                '
                tbl.PreferredWidthType = tblTop.PreferredWidthType
                '
                tblPreferredWidth = tblTop.PreferredWidth
                '
                drCell = Me.tbl_get_firstCellWithColumnIndex(currentColumnIdx, tbl)
                '
                If tblPreferredWidth > 1000 Then
                    tblPreferredWidth = tblTop.PreferredWidth
                    'MsgBox("cTablesMgr.Error line 4249")
                    'tbl.PreferredWidthType = tblTop.PreferredWidthType
                    'tbl.PreferredWidth = tblTop.PreferredWidth
                    'tbl.PreferredWidth = 100
                End If
                '
                '
                Select Case strColumnAction
                    Case "left"
                        drCell.Range.Select()
                        glb_get_wrdSel.InsertColumns()
                        newColumnIdx = drCell.ColumnIndex - 1
                        If newColumnIdx = 0 Then newColumnIdx = 1

                    Case "right"
                        drCell.Range.Select()
                        glb_get_wrdSel.InsertColumnsRight()
                        'newColumnIdx = drCell.ColumnIndex
                        newColumnIdx = drCell.ColumnIndex + 1
                        '
                    Case "delete"
                        drCell.Range.Select()
                        glb_get_wrdSel.Columns.Delete()
                        drCell = glb_get_wrdSelCell()
                        If Not IsNothing(drCell) Then
                            newColumnIdx = drCell.ColumnIndex
                        Else
                            newColumnIdx = -1
                        End If
                        '
                End Select
                '
                'Now repair the expansion of the table caused by the column insertion. Then delete the top split
                'paragraph. Note that we cannot delete it as splitParaTop.Range.Delete() as this (for some unknown reason
                'causes a resizing problem of the tbl (body) upon joining... But if we select and delete (i.e. simulate
                'a 'hand' select/delete then all is OK... So, we have to use this work around
                '
                tbl.Columns.PreferredWidthType = tblTop.PreferredWidthType
                tbl.PreferredWidth = tblTop.PreferredWidth
                '
                'Workaround - Simulate a hand selection/delete of the paragraph
                rng = splitParaTop.Range
                rng.Select()
                glb_get_wrdSel.Delete()
                '
                If Not (IsNothing(tblBottom) Or IsNothing(splitParaBottom)) Then
                    'tblBottom is now set to 100%, so now tblBody and tblBottom are compatible. When they
                    'are joined we have no strange surprises, but only if we simulate a 'hand' select and delete
                    'of the paragraph
                    '
                    'Workaround - Simulate a hand selection/delete of the paragraph
                    '
                    rng = splitParaBottom.Range
                    rng.Select()
                    glb_get_wrdSel.Delete()
                    '
                    'splitParaBottom.Range.Delete()
                    '
                End If
                '
            Catch ex As Exception
                newColumnIdx = -1
            End Try

        End If
        '
finis:
        '
        Return newColumnIdx
    End Function


    '
    Public Function tbl_Encaps_columnInsert_doSplit(strInsertDirection As String, ByRef tbl As Word.Table) As Integer
        Dim tblTop, tblBottom As Word.Table
        Dim splitParaTop, splitParaBottom As Word.Paragraph
        Dim tblPreferredWidth As System.Single
        Dim drCell As Word.Cell
        Dim drCol As Word.Column
        Dim currentColumnIdx, newColumnIdx As Integer
        Dim sect As Word.Section
        Dim wasRightAligned As Boolean
        '
        wasRightAligned = False
        '
        drCol = Nothing
        tblTop = Nothing
        tblBottom = Nothing
        splitParaTop = Nothing
        splitParaBottom = Nothing
        '
        If Not IsNothing(tbl) Then
            drCell = glb_get_wrdSelRng.Cells.Item(1)
            currentColumnIdx = drCell.ColumnIndex
            sect = tbl.Range.Sections.Item(1)
            '
            'If tbl.Rows.Alignment = WdRowAlignment.wdAlignRowRight Then
            'wasRightAligned = True
            'tbl.Rows.Alignment = WdRowAlignment.wdAlignRowLeft
            'End If
            '
            Try
                '
                '
                Select Case tbl.PreferredWidthType
                    Case WdPreferredWidthType.wdPreferredWidthPercent
                        '
                        tblPreferredWidth = tbl.PreferredWidth
                        'tbl=tblBody
                        tbl = tbl_Encaps_split(tbl, tblTop, tblBottom, splitParaTop, splitParaBottom)
                        '
                        If tblPreferredWidth > 1000 Then
                            tblPreferredWidth = tblTop.PreferredWidth
                            '
                            '***
                            'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                            'tbl.Columns.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                            '****
                        End If
                        '
                        newColumnIdx = Me.tbl_Encaps_columnInsert_ifPercent(strInsertDirection, currentColumnIdx, tbl, tblTop, tblBottom, splitParaTop, splitParaBottom, tblPreferredWidth)
                        '
                    Case WdPreferredWidthType.wdPreferredWidthPoints
                        '
                        '*** At this point we ould leave the table in points and process it through Me.tbl_Encaps_insertColumn_ifPts.
                        '*** Or we can convert to percentage and process it through Me.tbl_Encaps_insertColumn_ifPercent.. This
                        '*** version converts the table to percent
                        '
                        'tblPreferredWidth = tbl.PreferredWidth
                        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                        'The PreferredWidth may be undefined because of the table structure, but we'll try
                        'anyway and correct it if its undefined (=9999999)
                        '
                        tblPreferredWidth = tbl.PreferredWidth
                        '
                        'tbl.PreferredWidth = 200
                        tbl = tbl_Encaps_split(tbl, tblTop, tblBottom, splitParaTop, splitParaBottom)
                        '
                        'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                        'tbl.PreferredWidth = 200
                        'tblPreferredWidth_Pts = tbl.PreferredWidth
                        'If tblPreferredWidth_Pts > 1000 Then
                        'tblPreferredWidth_Pts = tblTop.PreferredWidth
                        'End If
                        '
                        If tblPreferredWidth = 9999999 Then
                            'Because the width is undefined (the table is not uniform) we need to get the measuremnt
                            'from the top/bottom row
                            tblPreferredWidth = tblTop.PreferredWidth
                        End If

                        '
                        'newColumnIdx = Me.tbl_Encaps_insertColumn_ifPts(strInsertDirection, currentColumnIdx, tbl, tblTop, tblBottom, splitParaTop, splitParaBottom, tblPreferredWidth)
                        newColumnIdx = Me.tbl_Encaps_columnInsert_ifPercent(strInsertDirection, currentColumnIdx, tbl, tblTop, tblBottom, splitParaTop, splitParaBottom, tblPreferredWidth)
                        '
                        '
                    Case WdPreferredWidthType.wdPreferredWidthAuto
                        newColumnIdx = -2
                        '
                    Case Else
                        tblPreferredWidth = 433
                End Select
                '
                '
                'tbl = tbl_convert_tblEncapsToStd(tbl)
                '

            Catch ex As Exception
                newColumnIdx = -1
                'MsgBox("cTablesMgr.3728")
            End Try


        End If
        '
        '
        Return newColumnIdx
        '
    End Function
    '
    '
    Public Function tbl_Encaps_columnDelete_ifPercent(columnIndex As Integer, ByRef tbl As Word.Table, ByRef tblTop As Word.Table,
                                                      ByRef tblBottom As Word.Table, ByRef splitParaTop As Word.Paragraph, ByRef splitParaBottom As Word.Paragraph, tblPreferredWidth As Single) As Integer
        Dim drCell, drCellCol As Word.Cell
        Dim rng As Word.Range
        '
        Try
            'Now make sure that the cell that was originally selected (i.e. the one that
            'identifies the column is still selected
            drCellCol = Me.tbl_getSelect_firstCellWithColumnIndex(columnIndex, tbl)
            rng = drCellCol.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Select()
            '
            glb_get_wrdSel.Columns.Delete()
            drCell = glb_get_wrdSelCell()
            'drCell = Me.tbl_getSelect_firstCellWithColumnIndex(columnIndex, tbl)
            If Not IsNothing(drCell) Then
                columnIndex = drCell.ColumnIndex
            Else
                columnIndex = -1
            End If
            '
            'Need to make certain that the columns are not set to some absolute value
            'tblBody is set so that the columns are percent and add up to 100, then
            'we set the table to 100% of the width of the smallest column in a multi column
            'layout or 100% of the width of the single sect.PageLayout.textColumn in
            'a single column layout
            '
            'Setup for equal sized columns
            tbl.Columns.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
            'tbl.Columns.PreferredWidth = tblPreferredWidth
            '***
            tbl.Columns.PreferredWidth = 100 / tbl.Columns.Count
            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
            '***
            'tbl.PreferredWidth = 100
            'tbl.PreferredWidth = 50
            tbl.PreferredWidth = tblPreferredWidth
            '
            If Not (IsNothing(tblBottom) Or IsNothing(splitParaBottom)) Then
                'tblBottom is now set to 100%, so now tblBody and tblBottom are compatible. When they
                'are joined we have no strange surprises
                tblBottom.Columns.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                'tblBottom.Columns.PreferredWidth = 50
                tblBottom.PreferredWidth = tblPreferredWidth
                splitParaBottom.Range.Delete()
                '
            End If
            '
            tblTop.Columns.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
            'tblTop.Columns.PreferredWidth = 50
            tblTop.PreferredWidth = tblPreferredWidth
            splitParaTop.Range.Delete()
            '
            '
        Catch ex As Exception
            columnIndex = -1
        End Try
        '
        Return columnIndex
    End Function
    '
    '
    Public Function tbl_Encaps_columnInsert_ifPercent(strInsertDirection As String, columnIndex As Integer, ByRef tbl As Word.Table, ByRef tblTop As Word.Table,
                                                      ByRef tblBottom As Word.Table, ByRef splitParaTop As Word.Paragraph, ByRef splitParaBottom As Word.Paragraph, tblPreferredWidth As Single) As Integer
        Dim drCell, drCellCol As Word.Cell
        Dim rng As Word.Range
        '
        Try
            'Now make sure that the cell that was originally selected (i.e. the one that
            'identifies the column is still selected
            drCellCol = Me.tbl_getSelect_firstCellWithColumnIndex(columnIndex, tbl)
            rng = drCellCol.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Select()

            Select Case strInsertDirection
                        '
                Case "left"
                    glb_get_wrdSel.InsertColumns()
                    drCell = glb_get_wrdSelCell()
                    'drCell = Me.tbl_getSelect_firstCellWithColumnIndex(columnIndex, tbl)
                    columnIndex = drCell.ColumnIndex
                            '
                Case "right"
                    glb_get_wrdSel.InsertColumnsRight()
                    drCell = glb_get_wrdSelCell()
                    'drCell = Me.tbl_getSelect_firstCellWithColumnIndex(columnIndex + 1, tbl)
                    columnIndex = drCell.ColumnIndex
                    '
            End Select
            '
            'Need to make certain that the columns are not set to some absolute value
            'tblBody is set so that the columns are percent and add up to 100, then
            'we set the table to 100% of the width of the smallest column in a multi column
            'layout or 100% of the width of the single sect.PageLayout.textColumn in
            'a single column layout
            '
            'Setup for equal sized columns
            tbl.Columns.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
            'tbl.Columns.PreferredWidth = tblPreferredWidth
            '***
            tbl.Columns.PreferredWidth = 100 / tbl.Columns.Count
            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
            '***
            'tbl.PreferredWidth = 100
            'tbl.PreferredWidth = 50
            tbl.PreferredWidth = tblPreferredWidth
            '
            If Not (IsNothing(tblBottom) Or IsNothing(splitParaBottom)) Then
                'tblBottom is now set to 100%, so now tblBody and tblBottom are compatible. When they
                'are joined we have no strange surprises
                tblBottom.Columns.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                'tblBottom.Columns.PreferredWidth = 50
                tblBottom.PreferredWidth = tblPreferredWidth
                splitParaBottom.Range.Delete()
                '
            End If
            '
            tblTop.Columns.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
            'tblTop.Columns.PreferredWidth = 50
            tblTop.PreferredWidth = tblPreferredWidth
            splitParaTop.Range.Delete()
            '
            '
        Catch ex As Exception
            columnIndex = -1
        End Try
        '
        Return columnIndex
    End Function
    '
    '
    Public Sub tbl_insert_Column(strInsertDirection As String, ByRef tbl As Word.Table)
        Dim drHeader, drLastBodyRow As Word.Row
        Dim drCellSelected, drHeaderCell As Word.Cell
        Dim drCol As Word.Column
        Dim headerOffSet As Single
        Dim rng As Word.Range
        Dim columnIndex, rowIndex, InsertedColumnIndex As Integer
        Dim tblWidth, leftIndent As Single
        Dim tblSource As Word.Table
        Dim objTools As New cTools()
        Dim strFlag As String
        Dim rngTableOriginal As Word.Range

        '
        'Setup initial values
        drLastBodyRow = Nothing
        leftIndent = 0
        '
        'Make certain that the Selection is in a Cell
        rng = Globals.ThisAddin.Application.Selection.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()
        '
        'Setup the initial parametrs. Get Table default width and find the selected cell and
        'its containing Table. The find the column and row indices of the selected cell
        'tblWidth = objTools.widthBetweenMargins
        tbl = rng.Tables.Item(1)
        '
        '******************************************************************************
        'Place a copy of the Table on the Clipboard, just in case something goes wrong
        rngTableOriginal = tbl.Range
        rngTableOriginal.Copy()
        '
        '******************************************************************************
        '
        'Now get the co-ordinates of the selected Cell... We'll use the column index later
        drCellSelected = rng.Cells.Item(1)
        columnIndex = drCellSelected.ColumnIndex
        rowIndex = drCellSelected.RowIndex
        '
        'Get the Header Row, then get the First cell offset... This is always non zero in
        'an AAC Table
        drHeader = Me.getHeaderRow(tbl)
        Try
            drHeaderCell = drHeader.Cells.Item(1)
            '
            'Make certain that the Table is regular, whilst recording the offsets
            'used in the Table
            '
            headerOffSet = drHeaderCell.LeftPadding
            tbl = Me.HeaderRowAndTable_SetRegularFormatting(headerOffSet, drHeader)
            leftIndent = drHeader.LeftIndent
            strFlag = Me.Get_TableType(tbl)
            '
        Catch ex As Exception
            MsgBox("Table Error. Are you sure your cursor in somewhere in the body of a standard AA Table?")
            Exit Sub
        End Try
        '
        '
        Try
            Select Case strFlag
                Case "MultiColumnWithSource"
                    'Find the last Body Row and split off the Caption
                    drLastBodyRow = Me.Get_LastBodyRow_ForMCWC(tbl)
                    drLastBodyRow.Select()
                    tblSource = tbl.Split(drLastBodyRow.Next)
                    tblWidth = Me.Get_TableWidth(tbl)
                    '
                    'Now make sure the Table is Regular and then add a column.. 
                    '
                    drCol = Me.Add_Column(strInsertDirection, columnIndex, tblWidth, tbl)
                    InsertedColumnIndex = drCol.Index
                    '
                    '
                    'Undo the Table Split, then reset the Header Row to have the original offset
                    Me.Delete_ParaAtEndOfTable(tbl)
                    drHeader = tbl.Rows.Item(1)
                    '
                    Me.ResetHeaderRow(headerOffSet, leftIndent, drHeader)
                    '
                    'Now make certian that the cursor is in the Header Cell of the new column
                    rng = drHeader.Cells.Item(InsertedColumnIndex).Range
                    'rng = drCol.Cells.Item(1).Range
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    rng.Select()
                Case "MultiColumnNoSource"
                    'Find the last Body Row and split off the Caption
                    drLastBodyRow = tbl.Rows.Last
                    drLastBodyRow.Select()
                    tblWidth = Me.Get_TableWidth(tbl)

                    'tblSource = tbl.Split(drLastBodyRow)
                    'rng = tblSource.Range.Cells.Item(1).Range
                    'rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    'rng.Select()
                    '
                    'tblSource = rng.Tables.Item(1)
                    'tblSource.Select()
                    '
                    'leftIndent = Me.Get_TableLeftIndent(tblWidth, tbl)
                    'MsgBox("left Indent = " + CStr(leftIndent) + vbCrLf + "tblWidth = " + CStr(tblWidth))
                    '
                    'Now make sure the Table is Regular and then add a column
                    drCol = Me.Add_Column(strInsertDirection, columnIndex, tblWidth, tbl)
                    InsertedColumnIndex = drCol.Index

                    '
                    'Undo the Table Split, then reset the Header Row to have the original offset
                    'Me.Delete_ParaAtEndOfTable(tbl)
                    drHeader = tbl.Rows.Item(1)
                    Me.ResetHeaderRow(headerOffSet, leftIndent, drHeader)
                    '
                    'Now make certian that the cursor is in the Header Cell of the new column
                    rng = drHeader.Cells.Item(InsertedColumnIndex).Range
                    'rng = drCol.Cells.Item(1).Range
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    rng.Select()
                '
                Case "SingleColumnWithSource"
                    'Find the last Body Row and split off the Caption
                    ' Msg
                    drLastBodyRow = Me.Get_LastBodyRow_ForSCWC(tbl)
                    drLastBodyRow.Select()
                    tblSource = tbl.Split(drLastBodyRow.Next)
                    tblWidth = Me.Get_TableWidth(tblSource)
                    '
                    'leftIndent = Me.Get_TableLeftIndent(tblWidth, tblSource)
                    'MsgBox("left Indent = " + CStr(leftIndent) + vbCrLf + "tblWidth = " + CStr(tblWidth))
                    '
                    'Now make sure the Table is Regular and then add a column
                    '
                    drCol = Me.Add_Column(strInsertDirection, columnIndex, tblWidth, tbl)
                    InsertedColumnIndex = drCol.Index
                    '
                    '
                    'Undo the Table Split, then reset the Header Row to have the original offset
                    Me.Delete_ParaAtEndOfTable(tbl)
                    drHeader = tbl.Rows.Item(1)
                    '
                    Me.ResetHeaderRow(headerOffSet, leftIndent, drHeader)
                    '
                    'Now make certian that the cursor is in the Header Cell of the new column
                    rng = drHeader.Cells.Item(InsertedColumnIndex).Range
                    'rng = drCol.Cells.Item(1).Range
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    rng.Select()
                '
                Case "SingleColumnNoSource"
                    'Find the last Body Row and split off the Caption
                    drLastBodyRow = Me.Get_LastBodyRow_ForSCWC(tbl)
                    'drLastBodyRow.Select()
                    'tblSource = tbl.Split(drLastBodyRow.Next)
                    'tblSource.Select()
                    '
                    tblWidth = Me.Get_TableWidth(tbl)
                    'tblWidth = tbl.PreferredWidth
                    'leftIndent = Me.Get_TableLeftIndent(tblWidth, tbl)
                    'MsgBox("left Indent = " + CStr(leftIndent) + vbCrLf + "tblWidth = " + CStr(tblWidth))

                    'Now make sure the Table is Regular and then add a column
                    drCol = Me.Add_Column(strInsertDirection, columnIndex, tblWidth, tbl)
                    InsertedColumnIndex = drCol.Index
                    '
                    'Undo the Table Split, then reset the Header Row to have the original offset
                    'Me.Delete_ParaAtEndOfTable(tbl)
                    drHeader = tbl.Rows.Item(1)
                    Me.ResetHeaderRow(headerOffSet, leftIndent, drHeader)
                    'tbl.Rows.SetLeftIndent(119.4, WdRulerStyle.wdAdjustProportional)
                    '
                    'Now make certian that the cursor is in the Header Cell of the new column
                    rng = drHeader.Cells.Item(InsertedColumnIndex).Range
                    'rng = drCol.Cells.Item(1).Range
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    rng.Select()
                    '
            End Select

        Catch ex As Exception

        End Try
        '
        'Globals.ThisAddin.Application.Selection.Paste()
        'shp = Globals.ThisAddin.Application.Selection.ShapeRange.Item(1)
        'shp.Top = topOld
        'shp.Left = leftOld

    End Sub
    '
    '
    Public Function Get_LastBodyRow_ForMCWC(ByRef tbl As Word.Table) As Word.Row
        Dim Flag, numCells As Integer
        Dim IsCaptionStyle As Boolean
        Dim dr As Word.Row

        '
        Flag = 0
        IsCaptionStyle = False
        '
        'Start at the last row and run up the Table. If the number of cells per row
        'chnages to somethog other than 1 then we have found the body of the Table
        'if it gets to row one and its still 1 cell wide then we have a one column Table
        dr = tbl.Rows.Last
        numCells = dr.Cells.Count
        '
        For i = tbl.Rows.Count To 1 Step -1
            dr = tbl.Rows.Item(i)
            If dr.Cells.Count <> numCells Then
                Exit For
            End If
        Next
        '
        Return dr
        '
    End Function
    '
    '
    Public Function Get_LastBodyRow_ForSCWC(ByRef tbl As Word.Table) As Word.Row
        Dim Flag, numCells As Integer
        Dim IsCaptionStyle As Boolean
        Dim dr As Word.Row

        '
        Flag = 0
        IsCaptionStyle = False
        '
        'Start at the last row and run up the Table. If the number of cells per row
        'chnages to somethog other than 1 then we have found the body of the Table
        'if it gets to row one and its still 1 cell wide then we have a one column Table
        dr = tbl.Rows.Last
        numCells = dr.Cells.Count
        '
        For i = tbl.Rows.Count To 1 Step -1
            dr = tbl.Rows.Item(i)
            If Not Me.Row_Has_BottomRowStyle(dr) Then
                Exit For
            End If
        Next
        '
        Return dr
        '
    End Function
    '
    '
    Public Function Row_Has_BottomRowStyle(ByRef dr As Word.Row) As Boolean
        Dim CellStyle As Word.Style
        Dim drCell As Word.Cell
        Dim rng, OldRng As Word.Range
        Dim IsCaptionStyle As Boolean
        Dim sel As Selection
        '
        sel = Me.glb_get_wrdSel
        OldRng = sel.Range
        '
        drCell = dr.Cells.Item(1)
        rng = drCell.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()
        '
        CellStyle = sel.Style
        Select Case CellStyle.NameLocal
            Case "spacer", "spacer_tbl", "Note", "Note label", "Source"
                IsCaptionStyle = True
            Case Else
                IsCaptionStyle = False
        End Select
        '
        'Re-establish the initial Selection
        OldRng.Select()
        '
        Return IsCaptionStyle
    End Function
    '
    ''' <summary>
    ''' Assumes that the indent has been removed
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function Get_TableWidth(ByRef tbl As Word.Table) As Single
        Dim i As Integer
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        Dim tblWidth As Single
        '
        tblWidth = 0.0
        dr = tbl.Rows.Item(1)
        '
        For i = 1 To dr.Range.Cells.Count
            'This does not tkae into account the space between columns
            drCell = dr.Range.Cells.Item(i)
            tblWidth = tblWidth + drCell.Width
        Next
        '
        Return tblWidth
    End Function
    '
    '
    Public Function Get_TableLeftIndent(tblWidth As Single, ByRef tbl As Word.Table) As Single
        Dim objTools As New cTools()
        Dim WidthBetweenMargins, LeftIndent As Single
        Dim tblOffSet As Single
        '
        WidthBetweenMargins = objTools.widthBetweenMargins()
        tblOffSet = WidthBetweenMargins - tblWidth
        '
        If tblOffSet = 0.0 Then
            'Table has been fitted to the margins
            LeftIndent = tbl.Range.Cells.Item(1).LeftPadding
        End If
        '
        If tblOffSet < 0.0 Then
            'Table overhangs the Margins

        End If
        '
        If tblOffSet > 0.0 Then
            'Table sits inside the Margins


        End If
        '
        LeftIndent = -(tblWidth - WidthBetweenMargins)
        '
        'If the LeftIndent is positive, then the Table is smaller than
        'the space between the margins.. The Table will then be hard up against the
        'left margin, leaving a  space between it's right edge and the right margin.
        'So we set the LeftIndent to 0.0
        '
        'If LeftIndent >= 0.0 Then LeftIndent = 0.0
        '
        Return LeftIndent
    End Function
    '
    Public Function Get_TableType(ByRef tbl As Word.Table) As String
        Dim numCells As Integer
        Dim strFlag As String
        Dim IsCaptionStyle As Boolean
        Dim dr As Word.Row
        '
        strFlag = ""
        IsCaptionStyle = False
        '
        'Start at the last row and run up the Table. If the number of cells per row
        'chnages to somethog other than 1 then we have found the body of the Table
        'if it gets to row one and its still 1 cell wide then we have a one column Table
        dr = tbl.Rows.Last
        numCells = dr.Cells.Count
        '
        If Me.Row_Has_BottomRowStyle(dr) Then
            'The last Row has a spacer, Source or Note Style Style
            dr = tbl.Rows.First
            If dr.Cells.Count = 1 Then
                strFlag = "SingleColumnWithSource"
            Else
                strFlag = "MultiColumnWithSource"
            End If
        Else
            'The Last Row does Not have a Caption Style
            dr = tbl.Rows.First
            If dr.Cells.Count <> 1 Then
                strFlag = "MultiColumnNoSource"
            Else
                strFlag = "SingleColumnNoSource"
            End If

        End If
        '
        Return strFlag
        '
    End Function
    '
    '
    Public Function tbl_column_Delete(ByRef tbl As Word.Table) As Integer
        Dim sel As Word.Selection
        Dim drCol As Word.Column
        Dim drCell As Word.Cell
        Dim columnIndex As Integer
        '
        drCol = Nothing
        columnIndex = -1
        '
        Try
            sel = glb_get_wrdSel()
            drCell = glb_get_wrdSelRng.Cells.Item(1)
            columnIndex = drCell.ColumnIndex
            '
            glb_get_wrdSel.Columns.Delete()
            '
        Catch ex As Exception
            columnIndex = -1
        End Try

        Return columnIndex
        '
    End Function
    '
    '
    Public Function tbl_column_Add(strLeftRight As String, ByRef tbl As Word.Table) As Integer
        'Dim sel As Word.Selection
        Dim drCol As Word.Column
        Dim drCell As Word.Cell
        Dim columnIndex As Integer
        Dim tblPreferredWidth As Single
        '
        drCol = Nothing
        columnIndex = -1
        '
        Try
            drCell = glb_get_wrdSelCell()
            '
            Select Case tbl.PreferredWidthType
                Case WdPreferredWidthType.wdPreferredWidthPercent

                Case WdPreferredWidthType.wdPreferredWidthPoints

            End Select
            '
            If Not IsNothing(drCell) Then
                'drCol = glb_get_wrdSelRng.Columns.Item(1)
                Select Case strLeftRight
                    Case "left"
                        '
                        glb_get_wrdSel.InsertColumns()
                        drCell = glb_get_wrdSelCell()
                        If Not IsNothing(drCell) Then
                            columnIndex = drCell.ColumnIndex
                        Else
                            columnIndex = -1
                        End If
                        '
                        'drCol = tbl.Columns.Add(drCol)
                        'columnIndex = drCol.Index

                        '
                    Case "right"
                        '
                        glb_get_wrdSel.InsertColumnsRight()
                        drCell = glb_get_wrdSelRng.Cells.Item(1)
                        If Not IsNothing(drCell) Then
                            columnIndex = drCell.ColumnIndex
                        Else
                            columnIndex = -1
                        End If
                        '
                End Select
                '
                'Setup for equal sized columns
                tbl.Columns.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                'tbl.Columns.PreferredWidth = tblPreferredWidth
                '***
                tbl.Columns.PreferredWidth = 100 / tbl.Columns.Count
                tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                '***
                'tbl.PreferredWidth = 100
                'tbl.PreferredWidth = 50
                tbl.PreferredWidth = tblPreferredWidth
                '
                '
                '
                '
                'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                'tbl.PreferredWidth = 100
                '
                'tbl.Columns.DistributeWidth()
                '
                'For single column tables the internal vertical border appears for no apparent
                'reason, so we have to ensure that it is not there
                tbl.Borders.Item(WdBorderType.wdBorderVertical).LineStyle = WdLineStyle.wdLineStyleNone
                '
            End If
            '
        Catch ex As Exception
            columnIndex = -1
        End Try

        Return columnIndex
        '
    End Function
    '
    Public Function Add_Column(strLeftRight As String, CursorPosition As Integer, tblWidth As Single, ByRef tbl As Word.Table) As Word.Column
        Dim drCol As Word.Column
        Dim strMsg As String
        '
        strMsg = "This function will only work on tables that " + vbCrLf
        strMsg = strMsg + "have a consistent column structure from top to bottom "
        '
        '
        drCol = Nothing
        Try
            drCol = tbl.Columns.Item(CursorPosition)
            '
            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
            '
            'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
            Select Case strLeftRight
                Case "left"
                    '
                    drCol = tbl.Columns.Add(tbl.Columns.Item(CursorPosition))
                '
                Case "right"
                    'To insert to the right we move to next column and insert to
                    'it's left. If the selected column is the last column then we
                    'just add a column to the end of the Table
                    '
                    If drCol.Index = tbl.Columns.Last.Index Then
                        drCol = tbl.Columns.Add()
                    Else
                        drCol = drCol.Next
                        drCol = tbl.Columns.Add(drCol)

                    End If
            End Select

            '
            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
            tbl.PreferredWidth = tblWidth
            '
            'tbl.Columns.DistributeWidth()
            '
            'For single column tables the internal vertical border appears for no apparent
            'reason, so we have to ensure that it is not there
            tbl.Borders.Item(WdBorderType.wdBorderVertical).LineStyle = WdLineStyle.wdLineStyleNone
            '
            '
        Catch ex As Exception
            MsgBox(strMsg)
        End Try
        '
        Return drCol
    End Function
    '
    ''' <summary>
    ''' This function will return the Header Row of the Table. At the moment it assumes
    ''' that this is always the first row
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function getHeaderRow(ByRef tbl As Word.Table) As Word.Row
        Try
            getHeaderRow = tbl.Rows.Item(1)
        Catch ex As Exception
            getHeaderRow = Nothing
        End Try
    End Function
    '
    '
    Public Function Delete_ParaAtEndOfTable(ByRef tbl As Word.Table) As Word.Range
        Dim rng As Word.Range
        '
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng.Select()
        rng.Expand(WdUnits.wdParagraph)
        rng.Delete()
        '
        Return rng
        '
    End Function

    Public Sub ResetHeaderRow(headerOffSet As Single, LeftIndent As Single, ByRef drheader As Word.Row)
        'Remember headerOffset is a positive value and 
        'LeftIndent Is a negative value
        '
        Dim drCell As Word.Cell
        '
        For i = 1 To drheader.Cells.Count
            drCell = drheader.Cells.Item(i)
            drCell.BottomPadding = 0.0
        Next
        '
        drheader.LeftIndent = LeftIndent - headerOffSet
        drheader.Cells(1).LeftPadding = headerOffSet
        'drheader.LeftIndent = -headerOffSet + LeftIndent
        'drheader.LeftIndent = -LeftIndent

        drheader.Cells(1).Width = drheader.Cells(1).Width + headerOffSet
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will return true if the cursor is in a AAC Box or Figure or Recomemndation
    ''' </summary>
    ''' <returns></returns>
    Public Function isInBoxFigureRec() As Boolean
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        Dim drCellStyle As Word.Style
        Dim strCell As String
        '
        rng = glb_get_wrdSelRng()
        isInBoxFigureRec = False
        '
        If rng.Tables.Count > 0 Then
            tbl = rng.Tables.Item(1)
            drCell = tbl.Range.Cells.Item(1)
            rng = drCell.Range
            strCell = rng.Text
            'dr = tbl.Rows.Item(1)
            'rng = dr.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            drCellStyle = rng.Style
            If drCellStyle.NameLocal Like "Caption*" And (strCell Like "Figure*" Or strCell Like "Box*") Then isInBoxFigureRec = True
            'If tbl.Rows.Count > 1 Then
            'dr = tbl.Rows.Item(2)
            'rng = dr.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'drCellStyle = rng.Style
            'If drCellStyle.NameLocal Like "Caption*" Then isInBoxFigureRec = True
            'End If
        Else
            isInBoxFigureRec = False
        End If
    End Function
    '

    '

    ''' <summary>
    ''' This function will return the number of floating tables in myDoc. In strTables you will find the caption (full or partial depending
    ''' on doFullCaption) as well as the page number both relative and absolute. If floatingTables Only is set to False the method will
    ''' do this for all tables
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="strTables"></param>
    ''' <param name="doFullCaption"></param>
    ''' <param name="floatingTablesOnly"></param>
    ''' <returns></returns>
    Public Function tbl_get_floatingTables(ByRef myDoc As Word.Document, ByRef strTables As String, doFullCaption As Boolean, Optional floatingTablesOnly As Boolean = True) As Integer
        Dim tbl As Word.Table
        Dim objTablesMgr As New cTablesMgr(myDoc)
        Dim kount, pageNum, pageNumAbs, numHeaderRows As Integer
        Dim leftIndent, cellPadding, leftIndentBody, bodyWidth As Single
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        '
        leftIndent = 0
        cellPadding = 0
        leftIndentBody = 0
        bodyWidth = 0
        '
        para = Nothing
        '
        numHeaderRows = 0
        kount = 0
        pageNum = 0
        pageNumAbs = 0
        strTables = ""
        '
        For Each tbl In myDoc.Tables
            rng = tbl.Range
            pageNum = rng.Information(WdInformation.wdActiveEndAdjustedPageNumber)
            pageNumAbs = rng.Information(WdInformation.wdActiveEndPageNumber)
            If pageNum Mod 2 = 0 Then
                'pageNum is even
            End If
            '
            If floatingTablesOnly Then
                If objTablesMgr.tbl_is_Floating(tbl) Then
                    kount = kount + 1
                    strTables = strTables + Me.tbl_get_tblCaption(tbl, para, doFullCaption) + vbTab + "pg " + pageNum.ToString() + vbTab + "abs " + pageNumAbs.ToString() + vbCrLf
                End If
            Else
                objTablesMgr.tbl_get_tableProperties(tbl, numHeaderRows, leftIndent, cellPadding, leftIndentBody, bodyWidth)
                If numHeaderRows >= 1 Then
                    'It is an AAC table
                    kount = kount + 1
                    strTables = strTables + Me.tbl_get_tblCaption(tbl, para, doFullCaption) + vbTab + "pg " + pageNum.ToString() + vbTab + "abs " + pageNumAbs.ToString() + vbCrLf
                End If
            End If
        Next
        '
        Return kount
        '
    End Function
    '
    Public Function tbl_get_tblCaption(ByRef tbl As Word.Table, ByRef para As Word.Paragraph, Optional doFullCaption As Boolean = False) As String
        Dim strCaption As String
        Dim rng As Word.Range
        'Dim para As Word.Paragraph
        Dim tokens As String()
        '
        strCaption = ""
        Try
            rng = tbl.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Move(WdUnits.wdParagraph, -1)
            para = rng.Paragraphs.Item(1)
            strCaption = para.Range.Text
            strCaption = Trim(strCaption)
            If Not doFullCaption Then
                tokens = strCaption.Split(vbTab)
                strCaption = tokens(0)
            End If
        Catch ex As Exception
            strCaption = ""
        End Try


        Return strCaption

    End Function
    '
    ''' <summary>
    ''' This method will take a table (typically one row as you might find in a header), 
    ''' apply a table style to it, turn on the heading row, then indent 
    ''' the table. If the table is one row it will add a dummy row to the top because it seems that
    ''' indenting a table with an applied style only works (as of April 4, 2022) if the table has more than
    ''' one row. It will identt the rows, then delete the dummy
    ''' </summary>
    ''' <param name="leftIndent"></param>
    ''' <param name="tbl"></param>
    ''' <param name="strTblStyleName"></param>
    Public Sub tbl_convert_oneRowTableToWCAG(cellPadding As Single, leftIndent As Single, ByRef tbl As Word.Table, Optional strTblStyleName As String = "aac Table (no lines)")
        Dim myDoc As Word.Document
        Dim dr As Word.Row
        '
        myDoc = tbl.Range.Document
        '
        Try
            tbl.Style = myDoc.Styles.Item(strTblStyleName)
            tbl.ApplyStyleHeadingRows = True
            dr = tbl.Rows.Add(tbl.Rows.First)
            For Each dr In tbl.Rows
                dr.LeftIndent = leftIndent
                dr.Cells.Item(1).LeftPadding = cellPadding
            Next
            tbl.Rows.First.Delete()
            '
        Catch ex As Exception

        End Try
        '

    End Sub
    '
    ''' <summary>
    ''' This method will return true if the document (myDoc) has at least one floating table.
    ''' Typically used in WCAG conversion to force the author to look at and handle all floating tables
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function tbl_has_floatingTables(ByRef myDoc As Word.Document) As Boolean
        Dim sect As Word.Section
        Dim tbl As Word.Table
        Dim rslt As Boolean
        '
        rslt = False
        '
        For Each sect In myDoc.Sections
            For Each tbl In sect.Range.Tables
                rslt = Me.tbl_is_Floating(tbl)
                If rslt = True Then
                    GoTo finis
                End If
            Next
        Next
        '
finis:
        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' Hopefully I can develop a reliable test that indicates whether a Table
    ''' is Floating or not
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_is_Floating(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        '
        rslt = tbl.Rows.WrapAroundText
        '
        Return rslt
    End Function
    '
    '
    ''' <summary>
    ''' This method will return true if the table can be manipulated row by row. If it returns true, then
    ''' the various table properties are returned as referenced variables. All measurements are in points
    ''' 
    ''' numHeaderRows       :The number of indented header rows in the table.. Measuring from the top
    ''' leftIndent          :First row left indent relative to the body of the table (leftIndentBody)
    ''' cellPadding         :Cell padding in the  first cell of the first row
    ''' leftIndentBody      :The leftIndent of the body of the table relative to the left margin
    ''' bodyWidth           :The width of the body of the table
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_get_tableProperties(ByRef tbl As Word.Table,
                                            ByRef numHeaderRows As Integer,
                                            ByRef leftIndent As Single,
                                            ByRef cellPadding As Single,
                                            ByRef leftIndentBody As Single,
                                            ByRef bodyWidth As Single) As Boolean
        Dim rslt As Boolean
        Dim lst As New Collection()
        '
        rslt = False
        '
        Try
            numHeaderRows = 0
            '
            cellPadding = tbl.Range.Cells.Item(1).LeftPadding
            leftIndent = -cellPadding
            leftIndentBody = Math.Round(tbl.Rows.Last.LeftIndent, 2)
            bodyWidth = Me.tbl_get_tableBodyWidth(tbl)

            'leftIndentTop = Math.Round(tbl.Rows.First.LeftIndent, 2)
            '
            'leftIndent = leftIndentTop - leftIndentBody
            '
            If glb_tbl_isLegacyAATable(tbl) Then numHeaderRows = 1
            '
            'bodyWidth = 0.0
            'For Each drCell In tbl.Rows.Last.Cells
            'bodyWidth = bodyWidth + drCell.Width
            'Next
            '
            'Determine how many outdented rows we have... Normally only one for AAC Tables, but authors
            'can add more
            'For j = 1 To tbl.Rows.Count
            'dr = tbl.Rows.Item(j)
            'If dr.LeftIndent <> leftIndentBody Then
            'numHeaderRows = numHeaderRows + 1
            'Else
            'Exit For
            'End If
            'Next j
            '
            rslt = True
            'lst.Add(leftIndent, "leftIndent")
            'lst.Add(cellPadding, "cellPadding")
            'lst.Add(leftIndentBody, "leftIndentBody")
            'lst.Add(bodyWidth, "bodyWidth")
            'lst.Add(CSng(kount), "numHeaderRows")
            '
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will return the body width of the table. It does so by getting get the width of 
    ''' the first row (in a table type safe way) by adding the width of the cells in the first row as 
    ''' determined by the rowIndex. Remember in vertically merged tables the rows are not accessible. If the Table
    ''' is an AAC legacy (outdented) table, then the returned table width is internally calculated as the
    ''' width of the first row minus the left padding in the first cell
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_get_tableBodyWidth(ByRef tbl As Word.Table) As Single
        Dim tblWidth As Single
        '
        tblWidth = Me.glb_tbls_getTableWidth(tbl)
        '
        Return tblWidth
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will convert selected table(s) to inline.. I hope
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function tbl_convert_toInLine(ByRef myDoc As Word.Document) As Boolean
        Dim rslt As Boolean
        Dim sect As Word.Section
        Dim rng As Word.Range
        '
        rslt = False
        Try
            For Each sect In myDoc.Sections
                rng = sect.Range
                rslt = Me.tbl_convert_toInLine(rng)
            Next
            rslt = True
        Catch ex As Exception
            rslt = False
        End Try

        Return rslt
    End Function
    '
    '
    ''' <summary>
    ''' This method will convert selected table(s) to inline.. I hope
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <returns></returns>
    Public Function tbl_convert_toInLine(ByRef rng As Word.Range) As Boolean
        Dim rslt As Boolean
        Dim tbl As Word.Table
        '
        rslt = False
        Try
            For Each tbl In rng.Tables
                Me.tbl_convert_toInLine(tbl)
            Next
        Catch ex As Exception

        End Try

        Return rslt
    End Function
    '
    Public Function tbl_convert_toInLineBasic(ByRef tbl As Word.Table) As Boolean
        'tbl.Rows.WrapAroundText = False
        Return False
    End Function
    '
    ''' <summary>
    ''' This method will convert the table to inline and then reposition it
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_convert_toInLineAAC(ByRef tbl As Word.Table) As Boolean
        Dim objGlobals As New cGlobals()
        Dim dr As Word.Row
        Dim rslt As Boolean
        Dim j, numHeaderRows As Integer
        Dim leftIndent, cellPadding, leftIndentBody, bodyWidth, pageWidth, delta As Single
        '
        leftIndent = 0
        leftIndent = 0
        cellPadding = 0
        leftIndentBody = 0
        bodyWidth = 0
        '
        If Me.tbl_get_tableProperties(tbl, numHeaderRows, leftIndent, cellPadding, leftIndentBody, bodyWidth) Then
            pageWidth = objGlobals.glb_get_widthBetweenMargins(tbl.Range.Sections.Item(1))
            delta = pageWidth - bodyWidth
            '
            'All WCAG tables must be inline, then split the table to egt the header row(s) in tblTop
            Me.tbl_convert_toInLine(tbl)
            '
            If bodyWidth <= pageWidth Then
                For Each dr In tbl.Rows
                    dr.LeftIndent = 0.0
                Next
                For j = 1 To numHeaderRows
                    dr = tbl.Rows.Item(j)
                    dr.LeftIndent = leftIndent
                Next
            Else
                For Each dr In tbl.Rows
                    dr.LeftIndent = delta
                    'dr.LeftIndent = leftIndentBody + leftIndent
                Next
                For j = 1 To numHeaderRows
                    dr = tbl.Rows.Item(j)
                    dr.LeftIndent = leftIndent + delta
                Next

            End If

        End If
        '
        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method will convert a table to inline.. I hope
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_convert_toInLine(ByRef tbl As Word.Table) As Boolean
        Dim objGlobals As New cGlobals()
        Dim rslt As Boolean
        '
        rslt = False
        '
        Try
            tbl.Rows.WrapAroundText = False
            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
            tbl.PreferredWidth = 100
            '
            rslt = True
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
    End Function
    '
    '
    ''' <summary>
    ''' This method will convert a table to floating
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub tbl_convert_toFloating(ByRef tbl As Word.Table)
        tbl.Rows.WrapAroundText = True

    End Sub
    '
    ''' <summary>
    ''' This method will convert a table to floating
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Sub tbl_convert_toFloatingRelToMargin(ByRef tbl As Word.Table, strPos As String, Optional boundary As Single = 12.0, Optional tblWidth As Single = -1.0)
        Dim sect As Word.Section
        'Dim drCell As Word.Cell
        '
        sect = tbl.Range.Sections.Item(1)
        If tblWidth < 0.0 Then tblWidth = glb_get_widthBetweenMargins(sect) / 2.0
        '
        '
        Select Case strPos
            Case "left"
                'Me.tbl_width_Change(tbl, tblWidth)
                'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                'tbl.PreferredWidth = 50
                tbl.Rows.WrapAroundText = True

                tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
                tbl.Rows.HorizontalPosition = 0.0
                '
                'tbl.Rows.DistanceTop = boundary / 2
                tbl.Rows.DistanceTop = 0
                tbl.Rows.DistanceBottom = 0.0
                tbl.Rows.DistanceLeft = 0.0
                tbl.Rows.DistanceRight = boundary / 2
                '
                tbl.Rows.Alignment = WdRowAlignment.wdAlignRowLeft
                '
                tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                tbl.PreferredWidth = 50
                'tbl.Rows.WrapAroundText = True


                '
            Case "right"
                'Me.tbl_width_Change(tbl, tblWidth)
                tbl.Rows.WrapAroundText = True
                'tbl.PreferredWidth = tblWidth

                tbl.Rows.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionMargin
                tbl.Rows.HorizontalPosition = 0.0
                '
                'tbl.Rows.DistanceTop = boundary / 2
                tbl.Rows.DistanceTop = 0
                tbl.Rows.DistanceBottom = 0.0
                tbl.Rows.DistanceLeft = boundary / 2
                tbl.Rows.DistanceRight = 0.0
                '
                tbl.Rows.Alignment = WdRowAlignment.wdAlignRowRight
                tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                tbl.PreferredWidth = 50
                '
        End Select
        '
        If Me.tbl_is_EncapsulatedTable(tbl) Then
            Me.tbl_colour_set_colourOfCellToNone(tbl.Range.Cells.Item(1))
        End If
        '

    End Sub
    '

    '

    ''' <summary>
    ''' This method will return true if the table (tbl) is an AAC legacy Table. It looks at the padding in the
    ''' first cell, which should be 22.7 pt (It checks that the padding is in a range 15.0 to 223.o pt). Then
    ''' to make sure it also checks that the padding of the second cell (which should normally be 0.0) is less than
    ''' the padding of the first cell... This test is table type safe. It will work for both row regular, column regular
    ''' and Tables with mixed horizontal and vertically merged cells
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_is_LegacyAATable(ByRef tbl As Word.Table) As Boolean
        '
        Return Me.glb_tbl_isLegacyAATable(tbl)
        '
    End Function
    '
    ''' <summary>
    ''' This method will set the header row colour of all standard and regular tables
    ''' with the colour rgbFillColour. It will also do the Glossary table, providing it
    ''' is regular and the 'doGlossaryTable' option is set to True
    ''' </summary>
    ''' <param name="rgbFillColour"></param>
    Public Sub tbl_colour_set_HeaderRow(rgbFillColour As Long, Optional doGlossaryTable As Boolean = False)
        Dim dr As Word.Row
        '
        For Each tbl In Me.glb_get_wrdActiveDoc().Tables
            Try
                If Me.tbl_is_tblStandard(tbl) And Me.glb_tbls_isRegular(tbl) Then
                    dr = tbl.Rows.Item(1)
                    Me.tbl_colour_set_colourOfRow(dr, rgbFillColour)
                Else
                    If doGlossaryTable Then
                        If Me.tbl_is_tblGlossary(tbl) And Me.glb_tbls_isRegular(tbl) Then
                            dr = tbl.Rows.Item(1)
                            Me.tbl_colour_set_colourOfRow(dr, rgbFillColour)
                        End If
                    End If
                End If
            Catch ex As Exception

            End Try
            '
        Next

    End Sub
    '
    Public Function tbl_is_tblGlossary(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim objTools As New cTools()
        Dim strStyleName As String
        '
        rslt = False
        '
        Try
            strStyleName = objTools.tools_tbls_getFirstCellStyleName(tbl)
            'For a box, the first cell style is either 'Caption Label' (character style)
            'or 'Caption' (paragraph style)
            If strStyleName = "Glossary" Then
                rslt = True
            End If
        Catch ex As Exception
            rslt = False
        End Try
        '
        '
        '
        Return rslt
    End Function
    '
    ''' <summary>
    ''' This method will return true if the table (tbl) is an AAC Box, Key Finding or Recommendation. 
    ''' Note that the test looks for the words 'Box', 'Key' or 'Recommendation' in the first row of the
    ''' table
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_is_tblStandard(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim objTools As New cTools()
        Dim strStyleName As String
        '
        rslt = False
        '
        '
        Try
            strStyleName = objTools.tools_tbls_getFirstCellStyleName(tbl)
            'For a box, the first cell style is either 'Caption Label' (character style)
            'or 'Caption' (paragraph style)



            If strStyleName Like "Table column heading*" Or strStyleName Like "Figur*" Or strStyleName Like "Box*" Then rslt = True
            '

            'drCell = tbl.Range.Cells.Item(1)
            'rng = drCell.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseStart)
            'rng.Move(WdUnits.wdParagraph, -1)
            '
            'para = rng.Paragraphs.Item(1)
            'paraStyle = para.Style
            'If paraStyle.NameLocal Like "Caption*" Then
            'rng = para.Range
            '
            'If rng.Text Like "Tabl*" Then
            'rslt = True
            ''End If
            'End If
        Catch ex As Exception
            rslt = False
        End Try
        '
        '
        '
        Return rslt
    End Function
    '

    '
    ''' <summary>
    ''' This method will return true if the table (tbl) is an AAC Box, Key Finding or Recommendation. 
    ''' Note that the test looks for the words 'Box', 'Key' or 'Recommendation' in the first row of the
    ''' table
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_is_EncapsulatedTable(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim objTools As New cTools()
        Dim strStyleName As String
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        '
        rslt = False
        '
        Try
            strStyleName = objTools.tools_tbls_getFirstCellStyleName(tbl)
            'For a box, the first cell style is either 'Caption Label' (character style)
            'or 'Caption' (paragraph style)
            If strStyleName Like "Caption*" Then
                drCell = tbl.Range.Cells.Item(1)
                rng = drCell.Range
                If rng.Text Like "Tabl*" Then
                    rslt = True
                End If
            End If
        Catch ex As Exception
            rslt = False
        End Try
        '
        '
        Return rslt
    End Function
    '
    '
    ''' <summary>
    ''' This method will return true if the table (tbl) is an encapsulated Figure
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_is_EncapsulatedFigure(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim objTools As New cTools()
        Dim strStyleName As String
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        '
        rslt = False
        '
        Try
            strStyleName = objTools.tools_tbls_getFirstCellStyleName(tbl)
            'For a box, the first cell style is either 'Caption Label' (character style)
            'or 'Caption' (paragraph style)
            If strStyleName Like "Caption*" Then
                drCell = tbl.Range.Cells.Item(1)
                rng = drCell.Range
                If rng.Text Like "Fig*" Then
                    rslt = True
                End If
            End If
        Catch ex As Exception
            rslt = False
        End Try
        '
        '
        Return rslt
    End Function
    '
    '
    ''' <summary>
    ''' This method will return true if the table (tbl) is an encapsulated Figure
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_is_EncapsulatedBox(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim objTools As New cTools()
        Dim strStyleName As String
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        '
        rslt = False
        '
        Try
            strStyleName = objTools.tools_tbls_getFirstCellStyleName(tbl)
            'For a box, the first cell style is either 'Caption Label' (character style)
            'or 'Caption' (paragraph style)
            If strStyleName Like "Caption*" Then
                drCell = tbl.Range.Cells.Item(1)
                rng = drCell.Range
                If rng.Text Like "Bo*" Then
                    rslt = True
                End If
            End If
        Catch ex As Exception
            rslt = False
        End Try
        '
        '
        Return rslt
    End Function
    '

    '
    '
    ''' <summary>
    ''' This method will return true if the table (tbl) is an AAC Figure. 
    ''' Note that the test looks for the words 'Figure' in the first row of the
    ''' table
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_is_AACFigure(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim rng As Word.Range
        '
        rslt = False
        '
        Try
            rng = tbl.Range.Cells.Item(1).Range
            If rng.Text Like "Figure*" Then
                rslt = True
            End If
        Catch ex As Exception
            rslt = False
        End Try
        '
        '
        Return rslt
    End Function
    '
    '
    ''' <summary>
    ''' This method will return true if the table (tbl) is an AAC Box, Key Finding or Recommendation. 
    ''' Note that the test looks for the words 'Box', 'Key' or 'Recommendation' in the first row of the
    ''' table
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_is_AACBox(ByRef tbl As Word.Table) As Boolean
        Dim rslt As Boolean
        Dim objTools As New cTools()
        Dim strStyleName As String
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        '
        rslt = False
        '
        Try
            If tbl.Columns.Count = 1 Then
                strStyleName = objTools.tools_tbls_getFirstCellStyleName(tbl)
                'For a box, the first cell style is either 'Caption Label' (character style)
                'or 'Caption' (paragraph style)
                If strStyleName Like "Caption*" Then
                    drCell = tbl.Range.Cells.Item(1)
                    rng = drCell.Range
                    If rng.Text Like "Box*" Or rng.Text Like "Recommendation*" Or rng.Text Like "Key*" Or rng.Text Like "Find*" Then
                        rslt = True
                    End If
                End If
            End If
        Catch ex As Exception
            rslt = False
        End Try
        '
        '
        Return rslt
    End Function
    '
    '
    ''' <summary>
    ''' This method will split the Table tbl at the row specified by drNumber. If all was well it will return true. If it does
    ''' return true, then tblTopPart contains the top Table, splitPara contains the paragraph between the two tables, and table
    ''' now contains the bottom table (i.e. that part of the original table that contains the row identified by drNumber and all
    ''' rows below it.
    ''' </summary>
    ''' <param name="drNumber"></param>
    ''' <param name="tbl"></param>
    ''' <param name="splitPara"></param>
    ''' <param name="tblTopPart"></param>
    ''' <returns></returns>
    Public Function tbl_split_Table(drNumber As Integer, ByRef tbl As Word.Table, ByRef splitPara As Word.Paragraph, ByRef tblTopPart As Word.Table) As Boolean
        Dim isSuccessFull As Boolean
        Dim dr As Word.Row
        '
        isSuccessFull = False
        '
        dr = tbl.Rows.Item(drNumber)
        isSuccessFull = Me.tbl_split_Table(dr, tbl, splitPara, tblTopPart)
        '
        Return isSuccessFull
        '
    End Function
    '
    Public Function tbl_split_Table(dr As Word.Row, ByRef tbl As Word.Table, ByRef splitPara As Word.Paragraph, ByRef tblTopPart As Word.Table) As Boolean
        Dim isSuccessFull As Boolean
        Dim rng As Word.Range
        '
        isSuccessFull = False
        '
        Try
            tbl = tbl.Split(dr)
            rng = tbl.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            'Go back one paragraph to get the split paragraph, then go back a
            'further one paragraph to get the top table
            '
            rng.Move(WdUnits.wdParagraph, -1)
            splitPara = rng.Paragraphs.Item(1)
            rng.Move(WdUnits.wdParagraph, -1)
            tblTopPart = rng.Tables.Item(1)
            '
            isSuccessFull = True
            '
        Catch ex As Exception
            isSuccessFull = False
        End Try
        '
        Return isSuccessFull
    End Function
    '
    '
    ''' <summary>
    ''' This method will search tbl from the last to first row, looking for a row with the
    ''' styles 'Source' or 'Note'. If the return value is true, then dr is set to the
    ''' 'Source' row
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="dr"></param>
    ''' <returns></returns>
    Public Function tbl_find_SourceRow(ByRef tbl As Word.Table, ByRef dr As Word.Row) As Boolean
        Dim rng As Word.Range
        Dim cellStyle As Word.Style
        Dim rslt As Boolean
        Dim i, kount As Integer
        '
        rslt = False
        '
        Try
            kount = 0
            For i = tbl.Rows.Last.Index To 1 Step -1
                dr = tbl.Rows.Item(i)
                rng = dr.Cells.Item(1).Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                cellStyle = rng.Style
                If cellStyle.NameLocal = "Source" Or cellStyle.NameLocal = "Note" Then
                    kount = i
                    rslt = True
                    Exit For
                End If
                rslt = False
            Next
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will set the top, bottom borders of the row dr. Note that
    ''' </summary>
    ''' <param name="dr"></param>
    ''' <param name="top"></param>
    ''' <param name="bottom"></param>
    Public Sub tbl_set_borders(ByRef dr As Word.Row, top As Boolean, bottom As Boolean)
        Dim brdrs As Word.Borders
        '
        'Set the top and bottom borders
        brdrs = dr.Borders
        If top Then
            brdrs.Item(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleSingle
            brdrs.Item(WdBorderType.wdBorderTop).LineWidth = WdLineWidth.wdLineWidth050pt
            brdrs.Item(WdBorderType.wdBorderTop).Color = RGB(0, 0, 0)
        End If
        '
        If Not top Then
            brdrs.Item(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleNone
        End If
        '
        If bottom Then
            brdrs.Item(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleSingle
            brdrs.Item(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth050pt
            brdrs.Item(WdBorderType.wdBorderBottom).Color = RGB(0, 0, 0)
        End If
        '
        If Not bottom Then
            brdrs.Item(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleNone
        End If
        '
    End Sub
    '
    '
    Public Function tbl_find_rowWithInlineShape(ByRef tbl As Word.Table) As Word.Row
        Dim kount, i As Integer
        Dim dr As Word.Row
        Dim drCell As Word.Cell
        '
        dr = Nothing
        '
        Try
            'Now find the row with the inline Shape
            kount = 0
            For i = tbl.Rows.Last.Index To 1 Step -1
                dr = tbl.Rows.Item(i)
                drCell = dr.Cells.Item(1)
                If drCell.Range.InlineShapes.Count >= 1 Then
                    kount = i
                    Exit For
                End If
                dr = Nothing
            Next
        Catch ex As Exception
            dr = Nothing
        End Try
        '
        Return dr
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will split off the top custom row(s). These rows sit above the row dr, which is the row containing the
    ''' figure inline image. On entry, the table tblBody contains the inline picture and any custom comment rows (at the top).
    ''' On exit, tblBody contains the row with the inline picture... The return value is the actual row that contains the
    ''' inline image
    ''' </summary>
    ''' <param name="leftIndent"></param>
    ''' <param name="dr"></param>
    ''' <param name="tblBody"></param>
    Public Function tbl_splitTopRow_AACFigureCustomRow(leftIndent As Single, ByRef dr As Word.Row, ByRef tblBody As Word.Table) As Word.Row
        Dim objTablesMgr As New cTablesMgr()
        Dim tblTop As Word.Table
        Dim drCustom, drImage As Word.Row
        Dim drCustomColourBack, drCustomColourFore As Integer
        Dim splitPara As Word.Paragraph
        '
        splitPara = Nothing
        tblTop = Nothing
        drImage = Nothing
        '
        drCustom = tblBody.Rows.Item(dr.Index - 1)
        drCustomColourBack = drCustom.Cells.Item(1).Shading.BackgroundPatternColor
        drCustomColourFore = drCustom.Cells.Item(1).Shading.ForegroundPatternColor
        '
        objTablesMgr.tbl_split_Table(dr, tblBody, splitPara, tblTop)
        splitPara.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
        splitPara.SpaceBefore = 0.0
        splitPara.SpaceAfter = 0.0
        splitPara.KeepWithNext = True
        '
        objTablesMgr.tbl_apply_figureTableStyle(leftIndent, tblTop)
        objTablesMgr.tbl_set_borders(tblTop.Rows.Last, False, True)
        objTablesMgr.tbl_set_borders(tblBody.Rows.First, False, True)
        '
        drImage = tblBody.Rows.First
        '
        splitPara.Range.Font.Size = 1.0
        '
        tblTop.Shading.BackgroundPatternColor = drCustomColourBack
        tblTop.Shading.ForegroundPatternColor = drCustomColourFore
        '
        Return drImage
        '
    End Function
    '
    Public Function tbl_splitBottomRow_AACFigureCustomRow(leftIndent As Single, ByRef dr As Word.Row, ByRef tblBody As Word.Table) As Word.Row
        Dim objTablesMgr As New cTablesMgr()
        Dim tblTop As Word.Table
        Dim drCustom, drImage As Word.Row
        Dim drCustomColourBack, drCustomColourFore As Integer
        Dim splitPara As Word.Paragraph
        '
        splitPara = Nothing
        tblTop = Nothing
        drImage = Nothing
        '
        Try
            drCustom = tblBody.Rows.Item(dr.Index + 1)
            drCustomColourBack = drCustom.Cells.Item(1).Shading.BackgroundPatternColor
            drCustomColourFore = drCustom.Cells.Item(1).Shading.ForegroundPatternColor
            '
            objTablesMgr.tbl_split_Table(drCustom, tblBody, splitPara, tblTop)
            objTablesMgr.tbl_apply_figureTableStyle(leftIndent, tblTop)
            objTablesMgr.tbl_set_borders(tblTop.Rows.First, True, False)
            '
            drImage = tblTop.Rows.Last
            '
            'Set the top and bottom borders
            'objTablesMgr.tbl_set_bordersTopBottom(tblTop)
            '
            splitPara.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
            splitPara.SpaceBefore = 0.0
            splitPara.SpaceAfter = 0.0
            splitPara.KeepWithNext = True
            '
            objTablesMgr.tbl_apply_figureTableStyle(leftIndent, tblBody)
            objTablesMgr.tbl_set_borders(tblBody.Rows.First, True, False)
            splitPara.Range.Font.Size = 1.0
            '
            tblBody.Shading.BackgroundPatternColor = drCustomColourBack
            tblBody.Shading.ForegroundPatternColor = drCustomColourFore

        Catch ex As Exception
            drImage = Nothing
        End Try
        '
        Return drImage
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will split off the top custom row of an AA encapsulated Table. It assumes that the checks
    ''' for encapsulation have been done elsewhere. These rows sit above the row dr, which is the row containing the
    ''' figure inline image. On entry, tbl is the wholde table including the top row which is to be split
    ''' off... On exit the return value is the split off top row and tbl is the rtest (or main body of the table).
    ''' Verified 20240718
    ''' </summary>
    ''' <param name="tbl"></param>
    Public Function tbl_splitTopRow_fromTable(ByRef tbl As Word.Table) As Word.Table
        Dim objTablesMgr As New cTablesMgr()
        Dim tblTop As Word.Table
        Dim dr As Word.Row
        'Dim drCustomColourBack, drCustomColourFore As Integer
        Dim splitPara As Word.Paragraph
        '
        splitPara = Nothing
        tblTop = Nothing
        '
        dr = tbl.Rows.Item(tbl.Rows.First.Index + 1)
        'drCustomColourBack = drCustom.Cells.Item(1).Shading.BackgroundPatternColor
        'drCustomColourFore = drCustom.Cells.Item(1).Shading.ForegroundPatternColor
        '
        objTablesMgr.tbl_split_Table(dr, tbl, splitPara, tblTop)

        'splitPara.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
        'splitPara.SpaceBefore = 0.0
        'splitPara.SpaceAfter = 0.0
        'splitPara.KeepWithNext = True
        '
        'objTablesMgr.tbl_apply_figureTableStyle(leftIndent, tblTop)
        'objTablesMgr.tbl_set_borders(tblTop.Rows.Last, False, True)
        'objTablesMgr.tbl_set_borders(tblBody.Rows.First, False, True)
        '
        'drImage = tblBody.Rows.First
        '
        ' splitPara.Range.Font.Size = 1.0
        '
        tblTop.Shading.Texture = WdTextureIndex.wdTextureNone
        tblTop.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic
        tblTop.Shading.ForegroundPatternColor = WdColor.wdColorAutomatic
        '
        Return tblTop
        '
    End Function
    '
    Public Function tbl_splitBottomRow_fromTable(ByRef tbl As Word.Table) As Word.Table
        Dim objTablesMgr As New cTablesMgr()
        Dim tblTop As Word.Table
        Dim dr As Word.Row
        Dim splitPara As Word.Paragraph
        '
        splitPara = Nothing
        tblTop = Nothing
        '
        Try
            dr = tbl.Rows.Last
            'drCustomColourBack = drCustom.Cells.Item(1).Shading.BackgroundPatternColor
            'drCustomColourFore = drCustom.Cells.Item(1).Shading.ForegroundPatternColor
            '
            objTablesMgr.tbl_split_Table(dr, tbl, splitPara, tblTop)
            'objTablesMgr.tbl_set_borders(tblTop.Rows.First, True, False)
            '
            '
            'Set the top and bottom borders
            'objTablesMgr.tbl_set_bordersTopBottom(tblTop)
            '
            'splitPara.Style = glb_get_wrdActiveDoc.stYles
            'splitPara.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
            'splitPara.SpaceBefore = 0.0
            'splitPara.SpaceAfter = 0.0
            'splitPara.KeepWithNext = True
            '
            tblTop.Shading.Texture = WdTextureIndex.wdTextureNone
            tblTop.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic
            tblTop.Shading.ForegroundPatternColor = WdColor.wdColorAutomatic

        Catch ex As Exception
            tblTop = Nothing
        End Try
        '
        Return tblTop
        '
    End Function
    '
    '
    '
    Public Sub tbl_apply_figureTableStyle(originalLeftIndent As Single, ByRef tbl As Word.Table)
        Dim myDoc As Word.Document
        '
        myDoc = tbl.Range.Document
        '
        tbl.ApplyStyleHeadingRows = True
        '
        If originalLeftIndent = 0.0 Then
            tbl.Style = myDoc.Styles.Item("aac Table (Figure)")
        Else
            tbl.Style = myDoc.Styles.Item("aac Table (Figure-Wide)")
            '
        End If
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will run backwards from the last row in the Table. It will return true if it finds the
    ''' body of the table (defined as having more cells than the last row). If it returns true, then dr contains the
    ''' value of the 'Source' row... We can use this to split the Source block away from the table
    ''' </summary>
    ''' <param name="dr"></param>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tbl_find_tableBodyBottom(ByRef dr As Word.Row, ByRef tbl As Word.Table) As Boolean
        Dim foundBody As Boolean
        Dim i, numLastRowCells, numCurrentRowCells As Integer
        '
        foundBody = False
        dr = Nothing
        '
        Try
            For i = tbl.Rows.Last.Index To 1 Step -1
                dr = tbl.Rows.Item(i)
                If i = tbl.Rows.Last.Index Then
                    numLastRowCells = dr.Range.Cells.Count
                    Continue For
                End If
                numCurrentRowCells = dr.Cells.Count
                If numCurrentRowCells <> numLastRowCells Then
                    foundBody = True
                    Exit For
                End If
            Next
            '
            If foundBody Then dr = tbl.Rows.Item(dr.Index + 1)

        Catch ex As Exception
            foundBody = False
        End Try
        '
        Return foundBody
        '
    End Function
    '
    '
End Class

Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class frm_TableBuilder
    '
    Public objTblMgr As cTablesMgr

    Public colourHeader As Long             'Table Header Colour
    Public colourUnits As Long              'Units row colour
    Public colourUnits_2 As Long              'Units row colour

    '
    Public tblCaptionStyle As Word.Style
    Public tblHeaderStyle As Word.Style
    Public tblUnitsStyle As Word.Style
    Public tblTextBoldStyle As Word.Style
    Public tblTextStyle As Word.Style
    '
    Public tblNoteStyle As Word.Style
    Public tblSourceStyle As Word.Style
    Public tblSpacerStyle As Word.Style
    '
    Public objGlobals As cGlobals
    Public widthBetweenMargins As Single
    '
    Public bottomSpacerRowHeight As Single
    '
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        '
        Me.objTblMgr = New cTablesMgr()
        '
        Me.txtBx_Offset.Text = CStr(objTblMgr.glb_get_TableOutdent_mm)

    End Sub
    '
    Private Sub frm_TableBuilder_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        Dim strMsg As String
        Dim objPlHMgr As New cPlHBase()
        Dim sect As Word.Section
        Dim doTableCheck As Boolean
        'Dim rng As Word.Range
        'Dim numRows, numColumns As Integer
        '
        'strMsg = "Form initialisation has failed . The most likely cause is a missing style" & vbCrLf
        'strMsg = strMsg & "Your document or template may be corrupted.. Or your document" & vbCrLf
        'strMsg = strMsg & "hasn't fully taken on board new styles added after a template update" & vbCrLf & vbCrLf
        'strMsg = strMsg & "Please contact your Admin Staff.. Or" & vbCrLf & vbCrLf
        'strMsg = strMsg & "Select File->Options->Addins->Templates (from the Manage Dropdown list)"
        'strMsg = strMsg & "->Click Go Button->Check 'Automatically update document styles'" & vbCrLf & vbCrLf
        'strMsg = strMsg & "Then close and re-open your document and try rebuilding the table"
        '
        strMsg = ""

        sect = objPlHMgr.glb_get_wrdSect()
        '
        Me.chkBx_Envelope.Checked = False
        'If sect.PageSetup.TextColumns.Count <> 1 Then Me.chkBx_Envelope.Checked = True
        '
        doTableCheck = False
        strMsg = objPlHMgr.Plh_is_OKToInsert(sect, doTableCheck)
        If strMsg = "" Then
            'The we are not in  a prohibited area, so now lets check that the selection is in
            'a table (one that is not a banner). If so, then adjust the from so that it is in
            '"Convert" mode
            If objPlHMgr.glb_selection_IsInTable() Then
                'If Not Me.objTblMgr.tbl_CellStyle_GetFirstCellStyle(objPlHMgr.glb_get_wrdSelTbl) Like "tag*" Then
                Me.Text = "Convert existing table To AAC standard"
                Me.txtBx_numBodyColumns.Enabled = False
                Me.txtBx_numBodyRows.Enabled = False
                Me.btn_BuildTable.Text = "Convert Table"
                '
                If Me.objTblMgr.tbl_CellStyle_GetFirstCellStyle(objPlHMgr.glb_get_wrdSelTbl) Like "tag*" Then
                    'We are in a banner, so don't allow table conversion
                    Me.btn_BuildTable.Enabled = False
                End If
            Else
                'MsgBox(strMsg)
            End If
        End If

        Try
            '
            Me.btn_BuildTable.Select()
            '
        Catch ex As Exception
            MsgBox(strMsg)
            Me.Close()
        End Try
        '
    End Sub
    '
    Private Sub btn_BuildTable_Click(sender As Object, e As EventArgs) Handles btn_BuildTable.Click
        Dim myDoc As Word.Document
        Dim rng, rngSrc As Word.Range
        Dim sect As Word.Section
        Dim objSectMgr As New cSectionMgr()
        Dim objFlds As New cFieldsMgr()
        Dim objTblMgr As New cTablesMgr()
        Dim objPlHMgr As New cPlHBase()
        Dim objStylesMgr As New cStylesManager()
        Dim objisOKMgr As New cIsOKToDo()
        Dim strMsg As String
        Dim numRows, numColumns As Integer
        Dim strRowSelection, strBottomRows, textStyleName, strTableCaption, strForceCaptionTo As String
        Dim tblOutDent As Single
        Dim tbl As Word.Table
        Dim drCell As Word.Cell
        Dim dr As Word.Row
        '
        objTblMgr.glb_get_wrdApp.ScreenUpdating = False
        '
        strMsg = "Your cursor needs to be at least one paragraph" + vbCrLf
        strMsg = strMsg + "clear of any tables or placeholders before you" + vbCrLf
        strMsg = strMsg + "can build a custom Table"
        '
        strRowSelection = "header"
        strBottomRows = "none"
        strTableCaption = "Table"
        strForceCaptionTo = ""
        '
        tblOutDent = CSng(Me.txtBx_Offset.Text)
        '
        If Me.rdBtn_ES.Checked Then strTableCaption = "Table_ES"
        If Me.rdBtn_Report.Checked Then strTableCaption = "Table"
        If Me.rdBtn_App.Checked Then strTableCaption = "Table_AP"
        If Me.rdBtn_Letter.Checked Then strTableCaption = "Table_LT"
        '
        If Me.rdBtn_ES.Checked Then strForceCaptionTo = "ES"
        If Me.rdBtn_Report.Checked Then strForceCaptionTo = "BD"
        If Me.rdBtn_App.Checked Then strForceCaptionTo = "AP"
        If Me.rdBtn_Letter.Checked Then strForceCaptionTo = "LT"
        '
        If Not chkBx_Caption.Checked Then strForceCaptionTo = "none"
        '
        If Me.chkBx_HeaderRow.Checked And Not Me.chkBx_UnitsRow.Checked Then strRowSelection = "header"
        If Me.chkBx_HeaderRow.Checked And Me.chkBx_UnitsRow.Checked Then strRowSelection = "header+UnitsRow"
        If Not Me.chkBx_HeaderRow.Checked And Me.chkBx_UnitsRow.Checked Then strRowSelection = "unitsRow"
        If Not Me.chkBx_HeaderRow.Checked And Not Me.chkBx_UnitsRow.Checked Then strRowSelection = "none"
        '
        If Me.chkBx_DataSource.Checked And Not Me.chkBx_Note.Checked Then strBottomRows = "sourceOnly"
        If Not Me.chkBx_DataSource.Checked And Me.chkBx_Note.Checked Then strBottomRows = "note"
        If Me.chkBx_DataSource.Checked And Me.chkBx_Note.Checked Then strBottomRows = "sourceAndNote"
        If Not Me.chkBx_DataSource.Checked And Not Me.chkBx_Note.Checked Then strBottomRows = "none"
        '
        textStyleName = objTblMgr.glb_var_style_tblTextStyle
        If Me.rdBtn_TextSmall.Checked Then textStyleName = objTblMgr.glb_var_style_tblTextStyle_small
        '
        sect = objTblMgr.glb_get_wrdSect()
        rng = objTblMgr.glb_get_wrdSelRng()
        '
        'strMsg = objPlHMgr.Plh_is_OKToInsert(sect, False)
        strMsg = objisOKMgr.isOKto_doAction_inReportBody()



        If strMsg = objisOKMgr._isOK Then
            'numROws is the number of body rows, numRows + 1 is the number of body rows
            'plus a header row
            '
            numRows = CInt(Me.txtBx_numBodyRows.Text)
            numColumns = CInt(Me.txtBx_numBodyColumns.Text)
            '
            '
            'If chkBx_UnitsRow.Checked Then numRows = numRows
            '
            Select Case Me.btn_BuildTable.Text
                Case "Build Table"
                    'Make certian that the selection is not at the top of a section
                    '
                    '*********
                    'Unecessary, since  cTablesMgr.tbl_format_rapidFormat.tbl_para_addAbove(tbl)
                    'has been put back to inserting a para above by adding a row above, then splitting
                    'the table and deleting the split away top row... So the function always works within
                    'its initial boundaries
                    '
                    'rng = objSectMgr.sct_set_SelforTableInsert()
                    '
                    '*********

                    If Not Me.chkBx_Envelope.Checked Then
                        'Produce a standard table
                        myDoc = rng.Document
                        If chkBx_HeaderRow.Checked Then numRows = numRows + 1
                        If chkBx_UnitsRow.Checked Then numRows = numRows + 1
                        '
                        tbl = objTblMgr.tbl_build_Table_Standard(rng, numRows, numColumns, textStyleName)
                        objTblMgr.glb_tbl_apply_aacTableBasicStyle(tbl)
                        tbl.Range.Style = myDoc.Styles.Item(textStyleName)
                        tbl.Rows.First.Range.Style = objStylesMgr.style_txt_getTableHeadingStyle(myDoc)
                        '
                        'objTblMgr.tbl_format_rapidFormat(tbl, strForceCaptionTo, "")
                        objTblMgr.tbl_format_rapidFormat(tbl, strForceCaptionTo, strBottomRows)
                        '
                        dr = objTblMgr.tbl_build_headerUnitsAndCaptionRow(tbl, 0.0, strRowSelection, textStyleName)
                        '
                        '
                    Else
                        'Produce a standard table, but initially with no caption or source
                        myDoc = rng.Document
                        'Add header and units row.. We can delete them afterwards
                        numRows = numRows + 1
                        numRows = numRows + 1
                        'If chkBx_HeaderRow.Checked Then numRows = numRows + 1
                        'If chkBx_UnitsRow.Checked Then numRows = numRows + 1
                        '
                        tbl = objTblMgr.tbl_build_Table_Standard(rng, numRows, numColumns, textStyleName)
                        objTblMgr.glb_tbl_apply_aacTableBasicStyle(tbl)
                        tbl.Range.Style = myDoc.Styles.Item(textStyleName)
                        tbl.Rows.First.Range.Style = objStylesMgr.style_txt_getTableHeadingStyle(myDoc)
                        '
                        'objTblMgr.tbl_format_rapidFormat(tbl, strForceCaptionTo, "")
                        objTblMgr.tbl_format_rapidFormat(tbl, "none", "none")
                        '
                        dr = objTblMgr.tbl_build_headerUnitsAndCaptionRow(tbl, 0.0, strRowSelection, textStyleName)
                        '
                        'The table has been modified. So the original tbl only returns the table
                        'that was sandwiched by the top and bottom cells. To get the full table
                        'look at the table in the returned range
                        rng = objTblMgr.tbl_format_rapidFormat_Encap(tbl, strForceCaptionTo)
                        tbl = rng.Tables.Item(1)
                        '
                        If Not Me.chkBx_Caption.Checked Then
                            drCell = tbl.Range.Cells.Item(1)
                            drCell.Range.Text = ""
                            drCell.Range.Style = myDoc.Styles.Item("Body Text")
                        End If
                        '
                        drCell = tbl.Range.Cells.Item(tbl.Range.Cells.Count)
                        rngSrc = drCell.Range
                        objTblMgr.tbl_insert_SourceAndNoteText(rngSrc, strBottomRows)
                        '
                        'If Not Me.chkBx_HeaderRow.Checked Then tbl.Rows.Item(2).Delete()
                        'Default condition is table with header and and
                        'empty units row
                        Select Case strRowSelection
                            Case "header"
                                'Get rid of Units row
                                dr = tbl.Rows.Item(3)
                                dr.Delete()
                                '
                                'Now make sure that the header row is setup
                                dr = tbl.Rows.Item(2)
                                objTblMgr.tbl_colour_set_colourOfRow(dr, objTblMgr._glb_colour_purple_Dark)
                                dr.Range.Style = myDoc.Styles.Item(objTblMgr.glb_var_style_tblHeaderStyle)

                            Case "unitsRow"
                                dr = tbl.Rows.Item(3)
                                objTblMgr.tbl_colour_set_colourOfRow(dr, objTblMgr._glb_colour_UnitsGrey)
                                '
                                dr = tbl.Rows.Item(2)
                                dr.Delete()
                                '
                            Case "header+UnitsRow"
                                dr = tbl.Rows.Item(3)
                                objTblMgr.tbl_colour_set_colourOfRow(dr, objTblMgr._glb_colour_UnitsGrey)
                                'dr.Shading.BackgroundPatternColor = Me.objGlobals._glb_colour_UnitsGrey
                                'dr.Shading.ForegroundPatternColor = Me.objGlobals._glb_colour_UnitsGrey
                                'dr.Shading.Texture = Word.WdTextureIndex.wdTextureSolid
                                dr.Range.Style = objTblMgr.glb_get_wrdStyle(objTblMgr.glb_var_style_tblUnitsStyle)

                            Case "none"
                                dr = tbl.Rows.Item(3)
                                dr.Delete()
                                dr = tbl.Rows.Item(2)
                                dr.Delete()
                                '
                        End Select

                    End If
                    '
                    '
                Case "Convert Table"
                    tbl = objTblMgr.glb_get_wrdSelTbl()
                    If Not IsNothing(tbl) Then
                        objTblMgr.tbl_aacTable_ConvertTo_AAC(tbl, Me.chkBx_doBorders.Checked, tblOutDent, strRowSelection, Me.chkBx_wideTable.Checked, strBottomRows, textStyleName)
                    End If
            End Select
            '
            If Me.rdBtn_ES.Checked Then objFlds.flds_update_SequenceNumbers_Tables_ES()
            If Me.rdBtn_Report.Checked Then objFlds.flds_update_SequenceNumbers_Tables()
            If Me.rdBtn_App.Checked Then objFlds.flds_update_SequenceNumbers_Tables_AP()
            If Me.rdBtn_Letter.Checked Then objFlds.flds_update_SequenceNumbers_Tables_LT()
            '
            objFlds.flds_update_CrossReferenceFields()
        Else
            MsgBox(strMsg)
        End If
        '
        objTblMgr.glb_get_wrdApp.ScreenUpdating = True
        Me.Close()
        '
    End Sub
    '

    '
    ''' <summary>
    ''' This method will determine whether the table has the same number of cells 
    ''' per row.. If not we will return false, that is, the table is not regular. This
    ''' test is not definitive, but will cover most circumstances
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tableIsRegular(ByRef tbl As Word.Table) As Boolean
        Dim objTblsMgr As New cTablesMgr()
        Dim rslt As Boolean
        '
        rslt = objTblsMgr.tableIsRegular(tbl)
        '
        '
        Return rslt
        '
    End Function
    Public Function headerRowDetected(ByRef tbl As Word.Table) As Boolean
        'Imperfect detection relying on use of correct Purple Colour. Already
        'the colour used in this document is different from that used in the
        'spreadsheets.. Hence the two tests
        '
        Dim dr As Word.Row
        Dim objGlobals As cGlobals
        '
        headerRowDetected = False
        '
        Try
            objGlobals = New cGlobals()
            '
            dr = tbl.Rows.Item(1)
            '
            'If dr.Shading.BackgroundPatternColor = RGB(151, 87, 166) Then headerRowDetected = True
            If dr.Shading.BackgroundPatternColor = Me.colourHeader Then headerRowDetected = True
            If dr.Shading.BackgroundPatternColor = objGlobals._glb_colour_purple_Dark Then headerRowDetected = True
            '
        Catch ex As Exception

        End Try
        '
    End Function

    '
    Public Sub getTableParts(ByRef tblMaster As Word.Table, ByRef tblHeader As Word.Table, ByRef tblBody As Word.Table, ByRef tblFooter As Word.Table)
        Dim rng As Word.Range
        '
        tblBody = tblMaster.Split(tblMaster.Rows(3))
        rng = tblBody.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        Call rng.MoveStart(WdUnits.wdTable, -1)
        tblHeader = rng.Tables.Item(1)
        '
        tblFooter = tblBody.Split(tblBody.Rows.Item(tblBody.Rows.Count - 1))
        '
    End Sub
    '
    Public Sub addEncapsulatingRows(ByRef tbl As Table, foundHeaderRow As Boolean)
        'This method will add rows around the table body
        Dim dr As Row
        '
        If Not foundHeaderRow Then
            dr = tbl.Rows.Item(1)
            dr = tbl.Rows.Add(dr)
            dr = tbl.Rows.Add(dr)
            dr = tbl.Rows.Add(dr)
        Else
            dr = tbl.Rows.Item(2)
            dr = tbl.Rows.Add(dr)
            dr = tbl.Rows.Item(1)
            dr = tbl.Rows.Add(dr)
        End If
        '
        dr = tbl.Rows.Last
        dr = tbl.Rows.Add(dr)
        dr = tbl.Rows.Last
        dr.Select()
        Globals.ThisAddIn.Application.Selection.Cut()
        '
        dr = tbl.Rows.Last
        dr.Select()
        Globals.ThisAddIn.Application.Selection.Paste()
        dr = tbl.Rows.Last
        dr = tbl.Rows.Add(dr)

        tbl.Descr = "table_Std"
        '
    End Sub
    '
    Public Function adjustTable(ByRef tbl As Word.Table, tblWidth As Single) As Single
        'This method adjusts various characterists of the table
        Dim objTools As New cTools()
        Dim dr As Word.Row
        '
        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        tbl.Rows.HeightRule = WdRowHeightRule.wdRowHeightAuto
        '
        'tbl.AllowAutoFit = False
        'Call tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow)
        tbl.TopPadding = 0#
        tbl.BottomPadding = 2.0#
        tbl.LeftPadding = 0#
        tbl.RightPadding = 0#
        'tbl.PreferredWidth = widthBetweenMargins
        'tbl.AllowAutoFit = False
        '

        If rdBtn_TextStandard.Checked Then
            tbl.Range.Style = "Table text"
        Else
            tbl.Range.Style = "Table text (small)"
        End If
        '
        'If we have what looks like a Header row, then lets apply the Header style
        '
        dr = tbl.Rows.Item(1)
        If dr.Shading.BackgroundPatternColor = Me.colourHeader Or dr.Shading.ForegroundPatternColor = Me.colourHeader Then
            dr.Range.Style = Me.tblHeaderStyle
        End If
        '
        tbl.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop
        If Me.chkBx_wideTable.Checked Then
            Call Me.Table_Fix(tbl, tblWidth, True)
        Else
            tblWidth = objTools.tools_math_MillimetersToPoints(CSng(Me.txtBx_TableWidth.Text))                'in millimeters
            Call Me.Table_Fix(tbl, tblWidth, True)
        End If
        '
        Return tblWidth
        '
    End Function

    '
    Public Function adjustTable(ByRef tbl As Word.Table) As Single
        'This method adjusts various characterists of the table
        Dim tblWidth As Single
        Dim doBorders As Boolean
        Dim objTools As New cTools()
        '
        'For some reason the tbl.Range.style expression does NOT
        'work on Tables pasted from Excel, but the Row by Row approach does
        'tbl.Range.style = "Table text"
        '
        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        tbl.Rows.HeightRule = WdRowHeightRule.wdRowHeightAuto
        '
        tbl.AllowAutoFit = False
        Call tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow)
        tbl.TopPadding = 0#
        tbl.BottomPadding = 2.0#
        tbl.LeftPadding = 0#
        tbl.RightPadding = 0#
        'tbl.PreferredWidth = widthBetweenMargins
        tbl.AllowAutoFit = False
        '
        If rdBtn_TextStandard.Checked Then
            tbl.Range.Style = "Table text"
        Else
            tbl.Range.Style = "Table text (small)"
        End If

        tbl.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop
        'tbl.Borders.Item(wdBorderHorizontal).LineStyle = wdLineStyleSingle
        'tbl.Borders.Item(wdBorderHorizontal).LineWidth = wdLineWidth050pt
        'tbl.Borders.Item(wdBorderHorizontal).Color = Me.objGlobals.colour_TableBorders
        'tbl.Borders.Item(wdBorderTop).LineStyle = wdLineStyleSingle
        'tbl.Borders.Item(wdBorderTop).LineWidth = wdLineWidth050pt
        'tbl.Borders.Item(wdBorderTop).Color = Me.objGlobals.colour_TableBorders
        'tbl.Borders.Item(wdBorderBottom).LineStyle = wdLineStyleSingle
        'tbl.Borders.Item(wdBorderBottom).LineWidth = wdLineWidth050pt
        'tbl.Borders.Item(wdBorderBottom).Color = Me.objGlobals.colour_TableBorders


        'For Each dr In tbl.Rows
        'dr.Range.style = "Table text"
        'Next dr
        '
        doBorders = True
        If Me.chkBx_wideTable.Checked Then
            'Don't need to adjust the Table because it has already been
            'set between the margins with the AutoFit
            tblWidth = Me.widthBetweenMargins
            Call Me.Table_Fix(tbl, tblWidth, True)
            'Call Me.Table_doBorders_MaintainPadding(tbl, doBorders, Me.objGlobals.colour_TableBorders)
        Else
            tblWidth = objTools.tools_math_MillimetersToPoints(CSng(Me.txtBx_TableWidth.Text))                'in millimeters
            Call Me.Table_Fix(tbl, tblWidth, True)
            'Call Me.Table_doBorders_MaintainPadding(tbl, doBorders, Me.objGlobals.colour_TableBorders)

        End If
        '
        tbl.AllowAutoFit = False
        '
        adjustTable = tblWidth
    End Function
    '
    Public Function adjustTable(ByRef tbl As Word.Table, ByRef objToolsMgr As cTools) As Single
        'This method adjusts various characterists of the table
        Dim tblWidth As Single
        Dim doBorders As Boolean
        Dim objTools As New cTools()
        '
        'For some reason the tbl.Range.style expression does NOT
        'work on Tables pasted from Excel, but the Row by Row approach does
        'tbl.Range.style = "Table text"
        '
        tbl.Rows.HeightRule = WdRowHeightRule.wdRowHeightAuto
        '
        tbl.AllowAutoFit = False
        Call tbl.AutoFitBehavior(WdAutoFitBehavior.wdAutoFitWindow)
        tbl.TopPadding = 0#
        tbl.BottomPadding = 2.0#
        tbl.LeftPadding = 0#
        tbl.RightPadding = 0#
        'tbl.PreferredWidth = widthBetweenMargins
        tbl.AllowAutoFit = False
        '
        If rdBtn_TextStandard.Checked Then
            tbl.Range.Style = "Table text"
        Else
            tbl.Range.Style = "Table text (small)"
        End If

        tbl.Range.Cells.VerticalAlignment = WdCellVerticalAlignment.wdCellAlignVerticalTop
        'tbl.Borders.Item(wdBorderHorizontal).LineStyle = wdLineStyleSingle
        'tbl.Borders.Item(wdBorderHorizontal).LineWidth = wdLineWidth050pt
        'tbl.Borders.Item(wdBorderHorizontal).Color = Me.objGlobals.colour_TableBorders
        'tbl.Borders.Item(wdBorderTop).LineStyle = wdLineStyleSingle
        'tbl.Borders.Item(wdBorderTop).LineWidth = wdLineWidth050pt
        'tbl.Borders.Item(wdBorderTop).Color = Me.objGlobals.colour_TableBorders
        'tbl.Borders.Item(wdBorderBottom).LineStyle = wdLineStyleSingle
        'tbl.Borders.Item(wdBorderBottom).LineWidth = wdLineWidth050pt
        'tbl.Borders.Item(wdBorderBottom).Color = Me.objGlobals.colour_TableBorders


        'For Each dr In tbl.Rows
        'dr.Range.style = "Table text"
        'Next dr
        '
        doBorders = True
        If Me.chkBx_wideTable.Checked Then
            'Don't need to adjust the Table because it has already been
            'set between the margins with the AutoFit
            '
            '***
            Dim rng As Word.Range
            rng = Globals.ThisAddIn.Application.Selection.Range
            '            
            tblWidth = Me.widthBetweenMargins
            'tblWidth = rng.Sections.Item(1).PageSetup.TextColumns.Item(1).Width
            '
            '
            'tblWidth = Me.widthBetweenMargins
            'Call Me.Table_Fix(tbl, tblWidth, True)
            'Call Me.Table_doBorders_MaintainPadding(tbl, doBorders, Me.objGlobals.colour_TableBorders)
        Else
            tblWidth = objTools.tools_math_MillimetersToPoints(CSng(Me.txtBx_TableWidth.Text))                'in millimeters
            Call Me.Table_Fix(tbl, tblWidth, True)
            'Call Me.Table_doBorders_MaintainPadding(tbl, doBorders, Me.objGlobals.colour_TableBorders)

        End If
        '
        tbl.AllowAutoFit = False
        '
        adjustTable = tblWidth
    End Function
    '
    Public Function buildTable2() As Word.Range
        'Dim objPlhMgr As New cChapterPlaceHolder()
        Dim objToolsMgr As New cTools()
        Dim rng As Word.Range
        Dim tbl As Word.Table
        Dim dlgResult As Integer
        Dim strMsg As String
        '
        strMsg = "The table wrapping function is designed to work on regular tables" + vbCrLf
        strMsg = strMsg + "A regular table is one without merged, or odd sized cells" + vbCrLf + vbCrLf
        strMsg = strMsg + "The more 'irregular' your table, the less predictable the result" + vbCrLf + vbCrLf
        strMsg = strMsg + "Do you still want to proceed?.. This action cannot be easily undone"
        '
        Try
            '

            If objGlobals.glb_get_wrdSel.Range.Tables.Count = 0 Then
                tbl = Me.tblBldr_Tables_BuildStandard()
                tbl.Rows.Item(1).HeadingFormat = True
                rng = tbl.Range
                '
            Else
                'We are in a Table or just under one, so we'll wrap it
                tbl = objGlobals.glb_get_wrdSel.Range.Tables.Item(1)
                If Not Me.tableIsRegular(tbl) Then
                    dlgResult = MsgBox(strMsg, vbYesNo, "Table Warning")
                    If dlgResult = vbNo Then
                        rng = Nothing
                        GoTo finis
                    End If
                End If
                Try
                    tbl = Me.tblBldr_Tables_WrapTable(tbl)
                    tbl.Rows.Item(1).HeadingFormat = True
                    rng = tbl.Range
                Catch ex As Exception
                    rng = Nothing
                    MsgBox("Failed to wrap the Table")
                End Try

            End If

        Catch ex As Exception
            rng = Nothing
        End Try
        '
finis:
        '
        '
        Return rng
    End Function
    '
    ''' <summary>
    ''' This function 
    ''' </summary>
    ''' <returns></returns>
    Public Function tblBldr_Tables_WrapTable(ByRef tbl As Word.Table) As Word.Table
        Dim objPlhMgr As New cPlHBase()
        Dim objTblMgr As New cTablesMgr()
        Dim rng As Word.Range
        Dim currentTextColumn As Integer
        Dim sect As Word.Section
        Dim tblWidth As Single
        Dim para As Word.Paragraph
        Dim dr, drOld, drHeader, drUnits As Word.Row
        Dim tblSplit As Word.Table
        Dim i As Integer
        Dim alreadyHasHeader As Boolean
        '
        alreadyHasHeader = False
        drHeader = Nothing
        '
        Try
            '
            rng = tbl.Range
            sect = rng.Sections.Item(1)
            '
            'Make some room above the table by adding two empty paragraphs
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Move(WdUnits.wdParagraph, -1)
            rng.Paragraphs.Add(rng)
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            '
            currentTextColumn = objPlhMgr.Plh_Columnsx2_FindColumnNumber(sect)
            tblWidth = sect.PageSetup.TextColumns.Item(currentTextColumn).Width
            objTblMgr.tbl_fix_Table(tbl, False)
            'Remove all padding and borders

            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
            '
            tblWidth = Me.adjustTable(tbl, tblWidth)
            tbl.AllowAutoFit = False
            'GoTo finis
            '
            'Note that the table to be wrapped already has a Header
            dr = tbl.Rows.Item(1)
            'dr.Delete()
            If dr.Shading.BackgroundPatternColor = Me.colourHeader Or dr.Shading.ForegroundPatternColor = Me.colourHeader Then
                alreadyHasHeader = True
            End If
            '

            Call Me.Table_doBorders_MaintainPadding(tbl, Me.chkBx_doBorders.Checked, Me.objGlobals._glb_colour_TableBorders)
            '
            'leace the shading alone
            '
            For i = 1 To tbl.Rows.Count
                dr = tbl.Rows.Item(i)
                Select Case i
                    Case 1
                        If dr.Shading.BackgroundPatternColor = Me.colourHeader Or dr.Shading.ForegroundPatternColor = Me.colourHeader Then
                            'leave it alone its going to be our header row
                        End If
                    Case 2
                        If dr.Shading.BackgroundPatternColor = Me.colourUnits Or dr.Shading.ForegroundPatternColor = Me.colourUnits Then
                            'leave it alone its going to be our units row
                        End If
                        If dr.Shading.BackgroundPatternColor = Me.colourUnits_2 Or dr.Shading.ForegroundPatternColor = Me.colourUnits_2 Then
                            'leave it alone its going to be our units row
                        End If

                    Case Else
                        dr.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic
                        dr.Shading.ForegroundPatternColor = WdColor.wdColorAutomatic
                End Select

            Next
            '
            'GoTo finis
            '
            If Me.chkBx_HeaderRow.Checked And Not Me.chkBx_UnitsRow.Checked Then
                '
                If Not alreadyHasHeader Then
                    drHeader = tbl.Rows.Add(tbl.Range.Rows.Item(1))

                Else
                    drHeader = tbl.Rows.Item(1)
                End If

                'Add Header Row on top
                'drOld = tbl.Range.Rows.Item(1)
                '
                'drHeader = tbl.Rows.Add(drOld)
                'drHeader = tbl.Rows.Item(1)

                'GoTo finis

                'tblOld = tbl.Split(drOld)
                'tblOld.Rows.Item(1).Borders.Item(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleNone

                'drHeader.Range.Select()
                Me.doHeaderRow_forWrap(drHeader)
                'Me.doHeaderRow(drHeader)
                '
                'rng = drHeader.Range
                'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                'para = rng.Paragraphs.Item(1)
                'para.Range.Delete()
                '
                '
            End If
            '
            If Me.chkBx_HeaderRow.Checked And Me.chkBx_UnitsRow.Checked Then
                'Add Header and Units Row
                drOld = tbl.Range.Rows.Item(1)
                drUnits = tbl.Rows.Add(drOld)
                drUnits.Shading.BackgroundPatternColor = Me.colourUnits
                drUnits.Range.Style = Globals.ThisAddIn.Application.ActiveDocument.Styles(Me.tblUnitsStyle)
                '
                drUnits.Range.Borders.Item(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleNone
                drUnits.Range.Borders.Item(WdBorderType.wdBorderHorizontal).LineStyle = WdLineStyle.wdLineStyleNone


                drHeader = tbl.Rows.Add(drUnits)
                drHeader.Shading.BackgroundPatternColor = Me.colourHeader
                drHeader.Range.Style = Globals.ThisAddIn.Application.ActiveDocument.Styles(Me.tblHeaderStyle)

                tblSplit = tbl.Split(drUnits)
                drHeader.Range.Select()
                Me.doHeaderRow(drHeader)
                '
                rng = drHeader.Range
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                para = rng.Paragraphs.Item(1)
                para.Range.Delete()
                '
                '
            End If
            '
            If Not Me.chkBx_HeaderRow.Checked And Me.chkBx_UnitsRow.Checked Then
                'Add Units Row on top
                drOld = tbl.Range.Rows.Item(1)
                drUnits = tbl.Rows.Add(drOld)
                drUnits.Shading.BackgroundPatternColor = Me.colourUnits
                drUnits.Range.Style = Globals.ThisAddIn.Application.ActiveDocument.Styles(Me.tblUnitsStyle)
                '
            End If
            '
            tbl = Globals.ThisAddIn.Application.Selection.Range.Tables.Item(1)
            '
            If Me.chkBx_Caption.Checked Then
                para = Me.doCaption_02(tbl, Me.tblCaptionStyle)
                ' 
                'Now adjust the indent of the Caption.. We get a row other than the Header row and set
                'the paragraph's indent to the indent of that row. If the Table has been indented, then indent the caption
                dr = tbl.Rows.Last
                '
                If dr.LeftIndent < 0.0 Then
                    para.LeftIndent = dr.LeftIndent - para.FirstLineIndent
                End If
                '

            End If
            '
            tbl.Rows.Add()
            tbl.Rows.Add()
            'tbl.Rows.Add()

            'GoTo finis
            '
            'rng = tbl.Range
            'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            '
            'dr = tbl.Rows.Last
            'dr.Range.Copy()
            'dr = tbl.Rows.Add(dr)
            '
            'dr.Range.Select()
            'Globals.ThisAddin.Application.Selection.Paste()

            'tbl = rng.Tables.Add(rng, 2, 1)
            '
            'This Is the spacer row
            dr = tbl.Rows.Last
            dr.Range.Cells.Item(1).TopPadding = 0.0
            dr.Range.Cells.Item(1).LeftPadding = 0.0
            dr.Range.Cells.Item(1).BottomPadding = 0.0
            dr.Range.Cells.Item(1).RightPadding = 0.0
            dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
            dr.Height = 8.0
            dr.Borders.Item(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleNone
            dr.Range.Style = Globals.ThisAddIn.Application.ActiveDocument.Styles.Item("spacer")
            '
            dr.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic
            dr.Shading.ForegroundPatternColor = WdColor.wdColorAutomatic

            dr.Cells.Merge()
            dr.Range.Delete()


            If Me.chkBx_DataSource.Checked Or Me.chkBx_Note.Checked Then
                dr = tbl.Rows.Last
                dr = tbl.Rows.Item(dr.Index - 1)
                '
                dr.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic
                dr.Shading.ForegroundPatternColor = WdColor.wdColorAutomatic
                '
                dr.Range.Style = Globals.ThisAddIn.Application.ActiveDocument.Styles.Item(Me.tblSourceStyle)
                dr.Cells.Merge()
                dr.Range.Delete()
                '
                'dr.Range.Borders.Item(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleSingle
                'dr.Range.Borders.Item(WdBorderType.wdBorderTop).LineWidth = WdLineWidth.wdLineWidth050pt
                'dr.Range.Borders.Item(WdBorderType.wdBorderTop).Color = Me.objGlobals.colour_TableBorders
                '
                'dr.Range.Borders.Item(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleSingle
                'dr.Range.Borders.Item(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth050pt
                'dr.Range.Borders.Item(WdBorderType.wdBorderBottom).Color = Me.objGlobals.colour_TableBorders
                '

                '
                rng = dr.Range
                rng.Collapse(WdCollapseDirection.wdCollapseStart)
                '
                'rng = objChptPlhMgr.Plh_Insert_SourceAndNoteText(rng, "sourceAndNote")
                If Me.chkBx_DataSource.Checked And Me.chkBx_Note.Checked Then rng = objPlhMgr.Plh_Insert_SourceAndNoteText(rng, "sourceAndNote")
                If Me.chkBx_DataSource.Checked And Not Me.chkBx_Note.Checked Then rng = objPlhMgr.Plh_Insert_SourceAndNoteText(rng, "sourceOnly")
                If Not Me.chkBx_DataSource.Checked And Me.chkBx_Note.Checked Then rng = objPlhMgr.Plh_Insert_SourceAndNoteText(rng, "note")
                '
            Else
                dr = tbl.Rows.Last
                dr = tbl.Rows.Item(dr.Index - 1)
                dr.Delete()

            End If
            '

            '
        Catch ex As Exception

        End Try
        '
finis:
        '
        Return tbl
    End Function
    '
    ''' <summary>
    ''' This method will build a standard Table that fits within the column that the selection
    ''' lies within. It sets the table up for the AutoFit option
    ''' </summary>
    ''' <returns></returns>
    Public Function tblBldr_Tables_BuildStandard() As Word.Table
        Dim numBodyRows, numColumns, numTableRows, numTextColumns, currentTextColumn As Integer
        Dim objPlhMgr As New cPlHBase()
        Dim tblWidth, marginWidth As Single
        Dim tbl As Word.Table
        Dim dr As Word.Row
        Dim rng As Word.Range
        Dim sect As Word.Section
        Dim para As Word.Paragraph
        '
        tbl = Nothing
        numBodyRows = CInt(Me.txtBx_numBodyRows.Text)
        numColumns = CInt(Me.txtBx_numBodyColumns.Text)
        '
        numTableRows = numBodyRows + 1                                          'spacer row at the bottom
        '
        'Add an extra empty paragraph so the Table Builder has room to operate
        rng = Globals.ThisAddIn.Application.Selection.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Paragraphs.Add(rng)
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng.Select()                '
        '
        sect = rng.Sections.Item(1)
        '
        numTextColumns = sect.PageSetup.TextColumns.Count
        currentTextColumn = objPlhMgr.Plh_Columnsx2_FindColumnNumber(sect)
        '
        If Me.chkBx_HeaderRow.Checked Then numTableRows = numTableRows + 1
        If Me.chkBx_UnitsRow.Checked Then numTableRows = numTableRows + 1
        If Me.chkBx_Note.Checked Or Me.chkBx_DataSource.Checked Then numTableRows = numTableRows + 1
        '
        marginWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
        '
        If Me.chkBx_wideTable.Checked Then
            tbl = rng.Tables.Add(rng, numTableRows, numColumns)
            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
            tblWidth = sect.PageSetup.TextColumns.Item(currentTextColumn).Width
            '
            tblWidth = Me.adjustTable(tbl, tblWidth)
        Else
            'We get the Table Width from the user setting
            tbl = rng.Tables.Add(rng, numTableRows, numColumns)
            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
            tblWidth = CSng(Me.txtBx_TableWidth.Text)
            '
            tblWidth = Me.adjustTable(tbl, tblWidth)
            If tblWidth > marginWidth Then
                For Each dr In tbl.Rows
                    dr.LeftIndent = -(tblWidth - marginWidth)
                Next
            End If

        End If
        '
        Call Me.Table_doBorders_MaintainPadding(tbl, Me.chkBx_doBorders.Checked, Me.objGlobals._glb_colour_TableBorders)
        '
        dr = tbl.Rows.Last
        dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
        dr.Height = 8.0
        dr.Borders.Item(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleNone
        dr.Range.Style = Globals.ThisAddIn.Application.ActiveDocument.Styles.Item("spacer")
        dr.Cells.Merge()
        '
        If Me.chkBx_HeaderRow.Checked And Me.chkBx_UnitsRow.Checked Then
            dr = tbl.Range.Rows.Item(1)
            Me.doHeaderRow(dr)
            '
            'Units Rows
            dr = tbl.Range.Rows.Item(2)
            dr.Range.Borders.Item(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleNone
            dr.Range.Borders.Item(WdBorderType.wdBorderHorizontal).LineStyle = WdLineStyle.wdLineStyleNone
            dr.Shading.BackgroundPatternColor = Me.colourUnits
            dr.Shading.Texture = Word.WdTextureIndex.wdTextureNone
            dr.Range.Style = Globals.ThisAddIn.Application.ActiveDocument.Styles.Item(Me.tblUnitsStyle)
            '
        End If
        '
        If Me.chkBx_HeaderRow.Checked And Not Me.chkBx_UnitsRow.Checked Then
            dr = tbl.Range.Rows.Item(1)
            Me.doHeaderRow(dr)
            '
            dr = tbl.Range.Rows.Item(2)
            dr.Range.Borders.Item(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleNone
            dr.Range.Borders.Item(WdBorderType.wdBorderHorizontal).LineStyle = WdLineStyle.wdLineStyleNone
        End If
        '
        If Me.chkBx_Caption.Checked Then
            para = Me.doCaption_02(tbl, Me.tblCaptionStyle)
            ' 
            'Now adjust the indent of the Caption.. We get a row other than the Header row and set
            'the paragraph's indent to the indent of that row. If the Table has been indented, then indent the caption
            dr = tbl.Rows.Last
            '
            If dr.LeftIndent < 0.0 Then
                para.LeftIndent = dr.LeftIndent - para.FirstLineIndent
            End If
            '
        End If
        '
        If Me.chkBx_DataSource.Checked Or Me.chkBx_Note.Checked Then
            dr = tbl.Rows.Last
            dr = tbl.Rows.Item(dr.Index - 1)
            dr.Range.Style = Me.tblSourceStyle
            dr.Cells.Merge()
            '
            rng = dr.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            'rng = objChptPlhMgr.Plh_Insert_SourceAndNoteText(rng, "sourceAndNote")
            If Me.chkBx_DataSource.Checked And Me.chkBx_Note.Checked Then rng = objPlhMgr.Plh_Insert_SourceAndNoteText(rng, "sourceAndNote")
            If Me.chkBx_DataSource.Checked And Not Me.chkBx_Note.Checked Then rng = objPlhMgr.Plh_Insert_SourceAndNoteText(rng, "sourceOnly")
            If Not Me.chkBx_DataSource.Checked And Me.chkBx_Note.Checked Then rng = objPlhMgr.Plh_Insert_SourceAndNoteText(rng, "note")

        End If
        '
        Return tbl
    End Function
    '
    Public Function MillimetersToPoints(measurementInmm As Single) As Single
        Dim offSet As Single
        '
        Try
            offSet = 72 * (offSet / 25.4)
            '
        Catch ex As Exception
            offSet = 72 * (8.0 / 25.4)
        End Try
        '
        Return offSet

    End Function
    '
    Public Function MillimetersToPoints(measurementInmm As String) As Single
        Dim offSet As Single
        '
        Try
            offSet = Convert.ToSingle(measurementInmm)
            offSet = 72 * (offSet / 25.4)
            '
        Catch ex As Exception
            offSet = 72 * (8.0 / 25.4)
        End Try
        '
        Return offSet
    End Function
    '
    '
    Public Function buildTable() As Word.Range
        'This form is 'mostly' self contained, but if the cBBlocksHandler class
        'is not available, then you'll have to provide the
        'objBBMgr.insertBuildingBlockFromDefaultLibToRange method. Just be careful
        'with regards to the Building Blocks library and category
        Dim objBBMgr As cBBlocksHandler
        Dim objPlhMgr As New cPlHBase()
        Dim objPlhTblMgr As New cPlHTable()
        Dim objToolsMgr As cTools
        Dim objTblMgr As New cTablesMgr()
        Dim objStylesMgr As cStylesManager
        '
        Dim i, j, k As Integer
        Dim widthBetweenMargins As Single
        Dim tblWidth As Single
        Dim existingTable As Boolean
        Dim foundHeaderRow As Boolean
        Dim colour_TableBorders As Long
        Dim numBodyRows As Integer
        Dim numHeaderRows As Integer
        Dim numColumns As Integer
        Dim doSplit As Boolean
        Dim doBorders As Boolean
        '
        Dim para As Paragraph
        Dim sect As Section
        Dim rng As Range
        Dim rngHdr As Word.Range
        Dim rngCaption As Word.Range
        Dim rngTbl As Word.Range
        Dim tbl As Table
        Dim dr As Row
        Dim drCell As Word.Cell
        Dim drSource As Row
        Dim srcRowExists As Boolean
        Dim drCol As Column
        Dim colWidth As Single
        Dim delta As Single
        Dim getWidthBetweenMargins As Single
        Dim tblCaptionStyle As Style
        Dim tblBodyStyle As Style
        Dim doTblAutoFit As Boolean
        Dim numColumnsInTable As Integer
        Dim drColumnWidth As Single
        Dim paras As Word.Paragraphs
        Dim tblHeader As Word.Table
        Dim tblBody As Word.Table
        Dim tblFooter As Word.Table
        Dim myDoc As Word.Document
        Dim strTestMsg As String
        '
        'Application.ScreenUpdating = False
        '
        On Error GoTo finis
        myDoc = objGlobals.glb_get_wrdActiveDoc
        tblHeader = Nothing
        tblBody = Nothing
        tblFooter = Nothing
        doBorders = True
        foundHeaderRow = False
        '
        colour_TableBorders = Me.objGlobals._glb_colour_TableBorders
        sect = Globals.ThisAddIn.Application.Selection.Range.Sections(1)
        objBBMgr = New cBBlocksHandler()
        objPlhMgr = New cPlHBase()
        objToolsMgr = New cTools()
        objStylesMgr = New cStylesManager()
        '

        '
        srcRowExists = True
        '
        'Need to make some room for the table to ensure that it
        'doesn't collide with the prior line, but got a request to remove the
        'spacing... So I've done it and inserted a routine that will
        'stop users from trying to insert a Table too close to an
        'existing Table
        '
        'Check to see if we are too close to a table
        rng = Globals.ThisAddIn.Application.Selection.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        'The following test will return true if the current insertion
        'point is just below an existing Table
        '
        objGlobals.glb_selection_IsInTable()

        'If objToolsMgr.glb_selection_isAtBottomOfTable(rng) Then
        'buildTable = rng
        'GoTo finis2
        'End If
        '
        If Globals.ThisAddIn.Application.Selection.Range.Tables.Count = 0 Then
            rng = Globals.ThisAddIn.Application.Selection.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            para = rng.Paragraphs.Item(1)
            '
            If para.Range.Text = "" Then
                'Set para = rng.Paragraphs.Add(rng)
                'Set para = para.Next
                'Set rng = para.Range
                'rng.Style = "Body Text"
                'rng.Collapse (wdCollapseStart)
                rng.Select()
            Else
                'Set para = rng.Paragraphs.Add(rng)
                'Set para = para.Next
                'Set rng = para.Range
                'rng.Style = "Body Text"
                'rng.Collapse (wdCollapseStart)
                '
                'Set para = rng.Paragraphs.Add(rng)
                'Set para = para.Next
                'Set rng = para.Range
                'rng.Style = "Body Text"
                'rng.Collapse (wdCollapseStart)
                rng.Select()
            End If
            '
            numBodyRows = CInt(Me.txtBx_numBodyRows.Text)
            numColumns = CInt(Me.txtBx_numBodyColumns.Text)
            '
            doSplit = False
            If numBodyRows <= 0 Then
                numBodyRows = 1
                doSplit = True
            End If
            '
            If numColumns <= 0 Then numColumns = 1
            '
        Else
            '
            numBodyRows = CInt(Me.txtBx_numBodyRows.Text)
            numColumns = CInt(Me.txtBx_numBodyColumns.Text)
            '
            doSplit = False
            If numBodyRows <= 0 Then
                numBodyRows = 1
                doSplit = True
            End If
            '
            If numColumns <= 0 Then numColumns = 1
            '
        End If
        '
        '******
        '
        If Globals.ThisAddIn.Application.Selection.Range.Tables.Count = 0 Then
            'Create the base table.. We first create one row with the
            'correct padding and then duplicate it.. Much faster
            'Set tbl = rng.Tables.Add(rng, (1 + 1 + 1 + numBodyRows + 1 + 1), numColumns)
            tbl = rng.Tables.Add(rng, numBodyRows, numColumns)
            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
            '
            tblWidth = Me.adjustTable(tbl, objToolsMgr)
            '
            Call Me.Table_doBorders_MaintainPadding(tbl, Me.chkBx_doBorders.Checked, Me.objGlobals._glb_colour_TableBorders)
            '
            foundHeaderRow = False
            Call Me.addEncapsulatingRows(tbl, foundHeaderRow)
            Call Me.getTableParts(tbl, tblHeader, tblBody, tblFooter)
            '
            '
            Call Me.Table_removeBorders(tblHeader)
            dr = Me.getHeaderRow(tblHeader)
            '
            Call Me.doHeaderRow(dr, Me.getHeaderRowOffset_pts(), colourHeader, foundHeaderRow)
            para = Me.doCaption_02(tbl, Me.tblCaptionStyle, objBBMgr)
            rngCaption = para.Range
            '
            Call Me.doUnitsRow(tblBody, Me.objGlobals._glb_colour_UnitsGrey)
            '
            'Now do the Source and Spacer Rows
            Call Me.Table_ColourFill_Blank(tblFooter)
            'Call Me.doSourceRowAsText(tbl)                         'As text row undr table
            Call Me.doSourceRow(tblFooter, Me.tblSourceStyle, objBBMgr)   'As a row
            Call Me.doSpacerRow(tblFooter, 6.0#, Me.tblSpacerStyle)
            '
            existingTable = False
        Else
            'Set our table to the existing table... we are going to wrap it
            'But first add an empty para after it so that it's behaviour is
            'consistent with the insertion of an empty table.. It too, ends
            'up with a paragraph after it
            tbl = Globals.ThisAddIn.Application.Selection.Tables.Item(1)
            '
            '*** Test
            objTblMgr.tbl_fix_Table(tbl, False)
            '***
            '
            rngTbl = tbl.Range
            rngTbl.Collapse(WdCollapseDirection.wdCollapseEnd)
            para = rngTbl.Paragraphs.Add(rngTbl)
            'para.Style = "Body Text"

            para.Style = myDoc.Styles.Item("Body Text")
            'para.Range.Select
            '********

            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
            '
            tblWidth = Me.adjustTable(tbl, objToolsMgr)
            Call Me.Table_doBorders_MaintainPadding(tbl, Me.chkBx_doBorders.Checked, Me.objGlobals._glb_colour_TableBorders)
            '
            numBodyRows = tbl.Rows.Count
            foundHeaderRow = Me.headerRowDetected(tbl)
            '
            If foundHeaderRow Then numBodyRows = numBodyRows - 1
            Call Me.addEncapsulatingRows(tbl, foundHeaderRow)
            Call Me.getTableParts(tbl, tblHeader, tblBody, tblFooter)
            '
            Call Me.Table_removeBorders(tblHeader)
            dr = Me.getHeaderRow(tblHeader)
            '
            Call Me.doHeaderRow(dr, Me.getHeaderRowOffset_pts(), colourHeader, foundHeaderRow)
            'dr.Range.Style = Me.tblHeaderStyle
            para = Me.doCaption_02(tbl, Me.tblCaptionStyle, objBBMgr)
            rngCaption = para.Range
            '
            Call Me.doUnitsRow(tblBody, Me.objGlobals._glb_colour_UnitsGrey)
            '
            'Now do the Source and Spacer Rows
            Call Me.Table_ColourFill_Blank(tblFooter)
            'Call Me.doSourceRowAsText(tbl)                         'As text row undr table
            Call Me.doSourceRow(tblFooter, Me.tblSourceStyle, objBBMgr)   'As a row
            Call Me.doSpacerRow(tblFooter, 6.0#, Me.tblSpacerStyle)
            '
            'Set tbl = Me.joinTables(tblHeader, tblBody, tblFooter)
            '
            existingTable = True
        End If
        '
        'Now delete extraneous items that can be deleted before the join
        If Not Me.chkBx_UnitsRow.Checked Then tblBody.Rows.Item(1).Delete()
        'If Not Me.chkBx_Caption.Value Then para.Range.Delete
        If Not Me.chkBx_Caption.Checked Then
            rngCaption.Select()
            Globals.ThisAddIn.Application.Selection.Delete()
            Call Globals.ThisAddIn.Application.Selection.MoveEnd(WdUnits.wdCharacter, 1)
            Globals.ThisAddIn.Application.Selection.Delete()
        End If
        If Not (Me.chkBx_DataSource.Checked Or Me.chkBx_Note.Checked) Then tblFooter.Rows.Item(1).Delete()
        'Else
        'If Me.chkBx_DataSource.Checked And Me.chkBx_Note.Checked Then tblFooter.Rows.Item(1).Range.Text = "Source  And Note Checked"
        'If Me.chkBx_DataSource.Checked And Not Me.chkBx_Note.Checked Then

        'tblFooter.Rows.Item(1).Range.Style = myDoc.Styles("Source")
        'tblFooter.Rows.Item(1).Range.Text = "Source Checked"
        'End If
        'If Not Me.chkBx_DataSource.Checked And Me.chkBx_Note.Checked Then
        'tblFooter.Rows.Item(1).Range.Style = myDoc.Styles("Note")
        'tblFooter.Rows.Item(1).Range.Text = "Note Checked"
        'End If
        'End If
        '
        tbl = Me.joinTables(tblHeader, tblBody, tblFooter)
        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        '
        'We need to go to the first row after the Header and turn off the top border,
        'but only if there's a Header Row
        If Me.chkBx_HeaderRow.Checked Then
            dr = tbl.Rows.Item(2)       'The Header Row is Row 1
            dr.Borders(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleNone
        End If
        '
        '
        '
        '
        'Setup the Caption row and then convert its numbering style to
        'whatever is selected
        'Set para = Me.doCaption(tbl, Me.tblCaptionStyle, objBBMgr)
        'Note that the caption indent must be done after the call
        'to doHeaderRow
        If Me.chkBx_Caption.Checked Then
            'If this is true, then the Caption has not been deleted, so its
            'OK to modify it
            Call Me.doCaptionIndent(para, tblWidth, tbl)
            'Call objPlhMgr.convertTablesTo("App")

            'objChptPlhMgr.Plh_Captions_InsertCaptions("TABLE_AP", Globals.ThisAddin.Application.Selection.Range, True, "")
            'objChptPlhMgr.Plh_Captions_ConvertCaptions(objPlhTblMgr.lstOfPlhTypes, "TABLE_ES", para.Range)
            '
            If Me.rdBtn_ES.Checked Then objPlhMgr.Plh_Captions_ConvertCaptions(objPlhTblMgr.lstOfPlhTypes, "Table_ES", para.Range)
            If Me.rdBtn_App.Checked Then objPlhMgr.Plh_Captions_ConvertCaptions(objPlhTblMgr.lstOfPlhTypes, "Table_AP", para.Range)
            If Me.rdBtn_Letter.Checked Then objPlhMgr.Plh_Captions_ConvertCaptions(objPlhTblMgr.lstOfPlhTypes, "Table_LT", para.Range)

            '
            'strTestMsg = "Para formatting" + vbCrLf
            'strTestMsg = strTestMsg + "para.leftIndent = " + Convert.ToString(para.LeftIndent) + vbCrLf
            'strTestMsg = strTestMsg + "para.FirstLineIndent = " + Convert.ToString(para.Format.FirstLineIndent) + vbCrLf
            'MessageBox.Show(strTestMsg)
        End If
        '
        '
        Globals.ThisAddIn.Application.ScreenUpdating = True
        Globals.ThisAddIn.Application.ScreenRefresh()
        'MessageBox.Show("Test Point")
        '
        If Not Me.chkBx_HeaderRow.Checked Then
            'Delete the Header Row in Place
            tbl.Rows.Item(1).Delete()
        End If
        '
        '
        tbl.AllowAutoFit = False
        tbl.Rows.AllowBreakAcrossPages = False
        '
        If Me.chkBx_HeaderRow.Checked Then
            dr = tbl.Rows.First
            dr.HeadingFormat = True
        End If
        '
        rng = tbl.Range
        'Call rng.MoveEnd(wdParagraph, 1)
        'rng.Select
        buildTable = tbl.Range
        '
        '
        '
        'Application.ScreenRefresh
        'Application.ScreenUpdating = True
        '
        'Unload Me
        Exit Function
finis:
        'Application.ScreenRefresh
        'Application.ScreenUpdating = True

        MsgBox("Error in Table Build. Most likely cause Is a non digit character has been entered, Or you might have tried to wrap an irregular table")
        Me.Close()
        Exit Function
finis2:
        '
        Call MsgBox("Your current cursor position Is too close to a Table" & vbCr & vbCr _
            & "Attempting to insert a New Table will cause the two tables to join." & vbCr _
            & "Try moving the insertion point away by one paragraph", vbOKOnly + vbInformation, "Template Message")

        Me.Close()
    End Function
    '
    '
    '
    Public Function getHeaderRow(ByRef tblHeader As Word.Table) As Word.Row
        'This method will get the Header Row from the Header Table
        getHeaderRow = tblHeader.Rows.Item(2)
    End Function
    '
    'This function will obtain the header row offset and return it
    'in points
    Public Function getHeaderRowOffset_pts() As Single
        Dim offset As Single
        Dim objTools As New cTools()
        '
        offset = CSng(Me.txtBx_Offset.Text)
        getHeaderRowOffset_pts = objTools.tools_math_MillimetersToPoints(offset)
        '
    End Function
    '
    'This method will fill the row dr with the colour fillColour
    '(generated fron the RGB function)
    Public Sub Row_ColourFill(ByRef dr As Word.Row, fillColour As Long)
        dr.Shading.BackgroundPatternColor = fillColour
        dr.Shading.Texture = Word.WdTextureIndex.wdTextureNone
    End Sub
    '
    Public Sub Row_ColourFill_Blank(ByRef dr As Word.Row)
        dr.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic
        dr.Shading.Texture = Word.WdTextureIndex.wdTextureNone
    End Sub
    '
    Public Sub Table_ColourFill_Blank(ByRef tbl As Word.Table)
        tbl.Shading.BackgroundPatternColor = WdColor.wdColorAutomatic
        tbl.Shading.Texture = Word.WdTextureIndex.wdTextureNone
    End Sub
    '

    '
    'This method will return tru if the row dr is the target row. Selection
    'criteria is the colour fill
    Public Function rowFillIsTarget(ByRef dr As Word.Row, testColour As Long) As Boolean
        rowFillIsTarget = False
        If dr.Shading.BackgroundPatternColor = testColour Then rowFillIsTarget = True
    End Function
    '
    '
    '
    Public Sub doBodyRows(ByRef tbl As Table, borderColour As Long, numBodyRows As Integer)
        'The numBodyRows is the number of Body Rows in the original
        'body segment of the table
        'Now do the borders of the Body and Source Rows
        Dim rng As Range
        Dim numRowsToMove As Integer
        'Dim borderColour As Long
        'Dim numBodyRows As Integer
        '
        'borderColour = RGB(100, 100, 100)
        'numBodyRows = 3
        '
        rng = tbl.Rows(3).Range
        numRowsToMove = (numBodyRows - 1)
        '
        Call rng.MoveEnd(Word.WdUnits.wdRow, (numBodyRows - 1))
        'rng.Select
        'Now do the Body and Source rows
        If numBodyRows = 1 Then
            'We only need to do the bottom if there is one row
        Else
            rng.Borders.Item(Word.WdBorderType.wdBorderHorizontal).LineStyle = Word.WdLineStyle.wdLineStyleSingle
            rng.Borders.Item(Word.WdBorderType.wdBorderHorizontal).LineWidth = Word.WdLineWidth.wdLineWidth050pt
            rng.Borders.Item(Word.WdBorderType.wdBorderHorizontal).Color = borderColour
        End If
        rng.Borders.Item(Word.WdBorderType.wdBorderBottom).LineStyle = Word.WdLineStyle.wdLineStyleSingle
        rng.Borders.Item(Word.WdBorderType.wdBorderBottom).LineWidth = Word.WdLineWidth.wdLineWidth050pt
        rng.Borders.Item(Word.WdBorderType.wdBorderBottom).Color = borderColour
        '
        'For Each dr In rng.Rows
        'Call Me.doBodyRowPadding(dr)
        'Set drcell = dr.Range.Cells.Item(1)
        'drcell.TopPadding = 2#
        'drcell.BottomPadding = 2#
        'drcell.leftPadding = 0#
        'drcell.RightPadding = 3#
        'Next dr
        '
    End Sub
    '
    Public Sub doBodyRowPadding(ByRef dr As Row)
        Dim drCell As Cell
        For Each drCell In dr.Cells
            drCell.TopPadding = 2.0#
            drCell.BottomPadding = 2.0#
            drCell.LeftPadding = 0#
            drCell.RightPadding = 3.0#
        Next drCell
        '
    End Sub
    '    '
    Public Function doCaption_02(ByRef tbl As Table, strCaptionStyle As String) As Paragraph
        Dim rng As Range
        Dim para As Paragraph
        Dim objChptPlhMgr As New cPlHBase()
        Dim captionStyle As Word.Style
        '
        captionStyle = tbl.Range.Document.Styles.Item(strCaptionStyle)
        '
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        rng.Move(WdUnits.wdParagraph, -1)
        rng.Paragraphs.Add(rng)
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
        para = rng.Paragraphs.Item(1)
        para.Style = captionStyle.NameLocal
        rng = para.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()
        '
        '*** Test
        If Me.rdBtn_ES.Checked Then objChptPlhMgr.Plh_Captions_InsertCaptions("Table_ES", rng, True)
        If Me.rdBtn_Report.Checked Then objChptPlhMgr.Plh_Captions_InsertCaptions("Table", rng, True)
        If Me.rdBtn_App.Checked Then objChptPlhMgr.Plh_Captions_InsertCaptions("Table_AP", rng, True)
        If Me.rdBtn_Letter.Checked Then objChptPlhMgr.Plh_Captions_InsertCaptions("Table_LT", rng, True)
        '
        para = rng.Paragraphs.Item(1)
        'para.SpaceBefore = 8.0#
        '
        Return para
    End Function


    '
    Public Function doCaption_02(ByRef tbl As Table, captionStyle As Style) As Paragraph
        Dim rng As Range
        Dim para As Paragraph
        Dim objChptPlhMgr As New cPlHBase()
        '
        '
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        rng.Move(WdUnits.wdParagraph, -1)
        '
        para = rng.Paragraphs.Item(1)
        para.Style = captionStyle.NameLocal
        rng = para.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()
        '
        '*** Test
        If Me.rdBtn_ES.Checked Then objChptPlhMgr.Plh_Captions_InsertCaptions("Table_ES", rng, True)
        If Me.rdBtn_Report.Checked Then objChptPlhMgr.Plh_Captions_InsertCaptions("Table", rng, True)
        If Me.rdBtn_App.Checked Then objChptPlhMgr.Plh_Captions_InsertCaptions("Table_AP", rng, True)
        If Me.rdBtn_Letter.Checked Then objChptPlhMgr.Plh_Captions_InsertCaptions("Table_LT", rng, True)
        '
        para = rng.Paragraphs.Item(1)
        'para.SpaceBefore = 8.0#
        '
        Return para
    End Function


    Public Function doCaption_02(ByRef tbl As Table, captionStyle As Style, ByRef objBBMgr As cBBlocksHandler) As Paragraph
        Dim rng As Range
        Dim dr As Row
        Dim para As Paragraph
        Dim objPlhMgr As New cPlHBase()
        '
        dr = tbl.Rows.Item(1)
        rng = dr.ConvertToText
        Call rng.MoveEnd(WdUnits.wdCharacter, -1)
        rng.Delete()
        rng.Select()
        '
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        Call rng.Move(WdUnits.wdParagraph, -1)
        '
        '******
        'Set para = rng.Paragraphs.Add(rng)
        'Set rng = para.Next.Range
        'rng.Select
        '*****
        '
        para = rng.Paragraphs.Item(1)
        para.Style = captionStyle.NameLocal
        rng = para.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()
        'tbl.Range.Select
        '

        'Set dr = tbl.Rows.Item(2)
        'Set rng = Me.doSplitTable_Base(tbl, dr, Me.tblCaptionStyle)

        'Set rng = Me.doCaptionDetached(tbl)
        'rng.Collapse (wdCollapseStart)
        'rng.Text = "Hello"
        '
        'Call objBBMgr.insertBuildingBlockFromDefaultLibToRange("table_captionAndHeadingNoPara", "Tables", rng)
        'Call objBBMgr.insertBuildingBlockFromDefaultLibToRange("captionAndHeading", "Tables", rng)
        'rng.Select
        'Set rng = objBBMgr.insertBuildingBlockFromDefaultLib_ReturnRangeOfBlock("captionAndHeading", "Tables")
        '
        '*** Test
        objPlhMgr.Plh_Captions_InsertCaptions("Table", rng, True)
        'rng = objBBMgr.insertBuildingBlockFromDefaultLib_ReturnRangeOfBlock("captionAndHeadingNoFormat", "Tables")
        'rng = objBBMgr.insertBuildingBlockFromDefaultLib_ReturnRangeOfBlock("captionAndHeading", "Tables")

        '***
        para = rng.Paragraphs.Item(1)
        para.SpaceBefore = 8.0#
        doCaption_02 = para
        '
        '** This causes a delay... Its the delete
        'Set para = para.Next
        'para.Range.Delete
        '
        'Try
        'Set rng = para.Range
        'Call rng.Move(WdUnits.wdParagraph, 1)
        'rng.Collapse (WdCollapseDirection.wdCollapseStart)
        'Call rng.MoveStart(WdUnits.wdCharacter, -1)
        'rng.Select
        '
        'rng.Delete
        'Set para = rng.Paragraphs.Item(1)

        'para.Range.Delete

        'doCaption_02.Range.Select
        '
        '
        'Call rng.Move(wdParagraph, -3)
        'Set dr = rng.Rows.Item(1)
        'dr.Delete
        '
        'Set rng = para.Range
        'Set tbl = rng.Tables.Item(1)

    End Function
    '
    Public Sub doCaptionIndent(ByRef para As Paragraph, tblWidth As Single, ByRef tbl As Word.Table)
        'This method will offset the leftindent of (generally) the Caption
        'Paragraph to match any Autofit functions.. tblWidth must be in points
        Dim objTools As cTools
        Dim rng As Word.Range
        Dim dr As Word.Row
        Dim indentSizeInPoints As Single
        '
        '
        indentSizeInPoints = 65.4

        '
        objTools = New cTools()
        rng = para.Range
        '
        If tblWidth <= objTools.widthBetweenMargins Then
            para.LeftIndent = indentSizeInPoints
            para.FirstLineIndent = -indentSizeInPoints
        Else
            dr = tbl.Rows.Item(2)
            para.LeftIndent = dr.LeftIndent + indentSizeInPoints
            para.FirstLineIndent = -indentSizeInPoints
        End If
        '
    End Sub
    '
    Public Function doCaptionDetached(ByRef tbl As Table) As Range
        Dim dr As Row
        '
        doCaptionDetached = Nothing
        Try
            dr = tbl.Rows.Item(2)
            If Me.chkBx_Caption.Checked Then
                doCaptionDetached = Me.doSplitTable_Base(tbl, dr, Me.tblCaptionStyle)
            End If

        Catch ex As Exception
            doCaptionDetached = Nothing
        End Try

    End Function
    '
    ''' <summary>
    ''' 20201118 This method will setup the input row as the Header Row 
    ''' </summary>
    ''' <param name="dr"></param>
    Public Sub doHeaderRow(ByRef dr As Word.Row, Optional existingHeader As Boolean = False)
        Dim offSet As Single
        Dim drCell As Word.Cell
        Dim sect As Word.Section
        '
        sect = dr.Range.Sections.Item(1)
        '
        offSet = -Me.MillimetersToPoints(Me.txtBx_Offset.Text)
        'dr = tbl.Range.Rows.Item(1)
        dr.LeftIndent = dr.LeftIndent + offSet
        dr.Shading.BackgroundPatternColor = Me.colourHeader
        dr.Shading.Texture = Word.WdTextureIndex.wdTextureNone
        dr.Range.Style = Globals.ThisAddIn.Application.ActiveDocument.Styles(Me.tblHeaderStyle)
        '
        dr.Borders.Item(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleNone
        dr.Borders.Item(WdBorderType.wdBorderRight).LineStyle = WdLineStyle.wdLineStyleSingle
        dr.Borders.Item(WdBorderType.wdBorderRight).LineWidth = WdLineWidth.wdLineWidth075pt
        dr.Borders.Item(WdBorderType.wdBorderRight).Color = Me.colourHeader

        '
        drCell = dr.Cells.Item(1)
        drCell.Width = drCell.Width - offSet
        drCell.LeftPadding = Math.Abs(offSet)
        '
        If existingHeader Then
            'Don't do anything if we hav an existin Header

        Else
            For i = 1 To dr.Cells.Count
                drCell = dr.Cells.Item(i)
                drCell.Range.Text = "Heading"
            Next i
        End If
        '
finis:
    End Sub
    '
    '
    ''' <summary>
    ''' 20201118 This method will setup the input row as the Header Row 
    ''' </summary>
    ''' <param name="dr"></param>
    Public Sub doHeaderRow_forWrap(ByRef dr As Word.Row)
        Dim offSet As Single
        Dim drCell As Word.Cell
        Dim sect As Word.Section
        '
        sect = dr.Range.Sections.Item(1)
        '

        offSet = -Me.MillimetersToPoints(Me.txtBx_Offset.Text)
        'drCell = dr.Range.Cells.Item(1)
        'drCell.Width = drCell.Width - offSet
        'GoTo finis

        'dr = tbl.Range.Rows.Item(1)
        dr.LeftIndent = dr.LeftIndent + offSet
        dr.Shading.BackgroundPatternColor = Me.colourHeader
        dr.Shading.Texture = Word.WdTextureIndex.wdTextureNone
        dr.Range.Style = Globals.ThisAddIn.Application.ActiveDocument.Styles(Me.tblHeaderStyle)
        '

        '
        dr.Borders.Item(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleNone
        dr.Borders.Item(WdBorderType.wdBorderRight).LineStyle = WdLineStyle.wdLineStyleSingle
        dr.Borders.Item(WdBorderType.wdBorderRight).LineWidth = WdLineWidth.wdLineWidth075pt
        dr.Borders.Item(WdBorderType.wdBorderRight).Color = Me.colourHeader

        '
        drCell = dr.Cells.Item(1)
        drCell.Width = drCell.Width - offSet
        drCell.LeftPadding = Math.Abs(offSet)
        '

finis:
    End Sub
    '
    Public Sub doHeaderRow(ByRef dr As Row, headerOffSet As Single, colourHeader As Long)
        Dim drCell As Cell
        Dim i As Integer
        '
        For i = 1 To dr.Cells.Count
            drCell = dr.Cells.Item(i)
            drCell.Range.Text = "Heading"
        Next i
        '

        dr.LeftIndent = -Globals.ThisAddIn.Application.MillimetersToPoints(CSng(Me.txtBx_Offset.Text))
        'dr.Cells.Item(1).Width = dr.Cells.Item(1).Width + Globals.ThisAddin.Application.MillimetersToPoints(CSng(Me.txtBx_Offset.Text))
        '
        'drCaption = dr.Previous
        'dr.LeftIndent = drCaption.LeftIndent - headerOffSet
        'dr.Cells(1).Width = dr.Cells(1).Width + headerOffSet
        For Each drCell In dr.Cells
            drCell.TopPadding = 0#
            drCell.BottomPadding = 0#
            drCell.Range.Style = Me.tblHeaderStyle.NameLocal
            'drCell.TopPadding = 0#
            'drCell.BottomPadding = 0#
            'drCell.leftPadding = 0#
            'drCell.RightPadding = 3#
            'If Not foundHeaderRow Then drCell.Range.Text = "Column Heading"
        Next
        'dr.Cells(1).LeftPadding = headerOffSet
        'dr.Range.Style = Me.tblHeaderStyle.NameLocal
        Call Me.Row_ColourFill(dr, colourHeader)                   'Header Row
finis:
        '
    End Sub

    '
    '
    Public Sub doHeaderRow(ByRef dr As Row, headerOffSet As Single, colourHeader As Long, Optional foundHeaderRow As Boolean = False)
        Dim drCaption As Row
        Dim drCell As Cell
        Dim i As Integer
        '
        If Not foundHeaderRow Then
            For i = 1 To dr.Cells.Count
                drCell = dr.Cells.Item(i)
                drCell.Range.Text = "Heading"
            Next i
        End If
        '
        drCaption = dr.Previous
        dr.LeftIndent = drCaption.LeftIndent - headerOffSet
        dr.Cells(1).Width = dr.Cells(1).Width + headerOffSet
        For Each drCell In dr.Cells
            drCell.TopPadding = 0#
            drCell.BottomPadding = 0#
            drCell.Range.Style = Me.tblHeaderStyle.NameLocal
            'drCell.TopPadding = 0#
            'drCell.BottomPadding = 0#
            'drCell.leftPadding = 0#
            'drCell.RightPadding = 3#
            'If Not foundHeaderRow Then drCell.Range.Text = "Column Heading"
        Next
        dr.Cells(1).LeftPadding = headerOffSet
        'dr.Range.Style = Me.tblHeaderStyle.NameLocal
        Call Me.Row_ColourFill(dr, colourHeader)                   'Header Row
finis:
        '
    End Sub
    '
    Public Sub doSourceRowAsText(ByRef tbl As Word.Table)
        'Now do the Source Row... as a text row under the Table
        Dim dr As Word.Row
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        Dim paras As Word.Paragraphs
        Dim objBBMgr As cBBlocksHandler
        '
        objBBMgr = New cBBlocksHandler()
        '
        dr = tbl.Rows(tbl.Rows.Count - 1)
        tbl = tbl.Split(dr)
        tbl.Delete()
        '
        rng = Globals.ThisAddIn.Application.Selection.Range
        Call rng.Move(WdUnits.wdParagraph, 1)
        tbl = rng.Tables.Item(1)
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        para = rng.Paragraphs.Item(1)
        para.Range.Select()
        para.Range.Style = "Body Text"
        rng = para.Range
        Call rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()
        '
        'Now insert Source etc
        rng = objBBMgr.insertBuildingBlockFromDefaultLib_ReturnRange("table_sourceCellContents", "Tables")
        Call rng.MoveEnd(WdUnits.wdParagraph, 3)
        rng.Select()
        '
        paras = Globals.ThisAddIn.Application.Selection.Range.Paragraphs
        para = paras.Item(3)
        If para.Range.Text Like "SOURCE*" Then
            rng = para.Range
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            Call rng.Move(WdUnits.wdCharacter, -1)
            rng.Select()
        End If
        '
    End Sub

    '

    Public Sub doSourceRow(ByRef tbl As Word.Table, tblSourceStyle As Style, ByRef objBBMgr As cBBlocksHandler)
        'This method expects as input the Footer Section of the Table (as a Table, that is the
        'Table has been parsed into three parts, namely, tblHeader, tblBody and tblFooter
        Dim dr As Word.Row
        Dim rng As Range
        Dim para As Word.Paragraph
        Dim paras As Word.Paragraphs
        Dim objPlhMgr As New cPlHBase()
        'Dim i As Integer
        '
        dr = tbl.Rows.Item(1)
        dr.Range.Style = Globals.ThisAddIn.Application.ActiveDocument.Styles(tblSourceStyle.NameLocal)
        dr.Cells.Merge()
        '
        rng = dr.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        rng = objPlhMgr.Plh_Insert_SourceAndNoteText(rng, "sourceAndNote")
        '
        If Me.chkBx_DataSource.Checked And Not Me.chkBx_Note.Checked Then
            paras = rng.Paragraphs
            paras.Item(2).Range.Delete()
            paras.Item(1).Range.Delete()
        End If
        If Not Me.chkBx_DataSource.Checked And Me.chkBx_Note.Checked Then
            paras = rng.Paragraphs
            paras.Item(3).Range.Delete()
        End If
        '
        rng = dr.Range
        rng.MoveEnd(WdUnits.wdCharacter, -2)

        paras = rng.Paragraphs
        para = paras.Last()
        If para.Range.Text Like "SOURCE*" Or para.Range.Text Like "Note*" Then
            rng = para.Range
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            Call rng.MoveStart(WdUnits.wdCharacter, -1)
            rng.Select()
            Globals.ThisAddIn.Application.Selection.Delete()
        End If
        '
        '
    End Sub
    '
    Public Sub doSpacerRow(ByRef tbl As Table, heightInPoints As Single, spacerStyle As Style)
        'This method expects as input the Footer Section of the Table (as a Table, that is the
        'Table has been parsed into three parts, namely, tblHeader, tblBody and tblFooter
        Dim dr As Row
        Dim strMsg As String
        '
        strMsg = "TableBuilder has partially recovered from a Styles Error" & vbCrLf
        strMsg = strMsg & "You are most likely missing the 'spacer_tbl' style" & vbCrLf
        strMsg = strMsg & "This generally occurs after a template update and if your document" & vbCrLf
        strMsg = strMsg & "has not been set to 'take on the styles from the new template.'" & vbCrLf & vbCrLf
        strMsg = strMsg & "Please contact your Admin Staff.. or" & vbCrLf & vbCrLf
        strMsg = strMsg & "Select File->Options->Addins->Templates (from the Manage Dropdown list)"
        strMsg = strMsg & "->Click Go Button->Check 'Automatically update document styles'"

        On Error GoTo finis
        dr = tbl.Rows.Last
        '
        dr.Borders(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleNone
        dr.HeightRule = WdRowHeightRule.wdRowHeightExactly
        dr.Height = heightInPoints
        '
        dr.Cells.Merge()
        dr.Range.Style = spacerStyle.NameLocal
        Exit Sub
        '
finis:
        dr.Range.Style = "spacer"
        MsgBox(strMsg)
    End Sub
    '
    Public Sub doSplitTable(ByRef tbl As Table, doSplit As Boolean, srcRowExists As Boolean)
        Dim dr As Row
        Dim rng As Range
        '
        If doSplit Then
            'Need to split table if the source row exists
            If srcRowExists Then
                'For Each dr In tbl.Range.Rows
                'If dr.Range.Cells.Count = 1 Then
                'Could be a Source Row
                'Set drCell = dr.Range.Cells.Item(1)
                'If drCell.TopPadding = 2# And drCell.BottomPadding = 2# Then Exit For
                'End If
                'Next dr
                dr = tbl.Rows.Last
                dr = dr.Previous
                dr = dr.Previous
                Call dr.Delete()
                '
                dr = tbl.Rows.Last
                dr = dr.Previous

                'We now have the source row
                tbl = tbl.Split(dr)
                tbl.Range.Select()
                Globals.ThisAddIn.Application.Selection.Collapse(WdCollapseDirection.wdCollapseStart)
                'Call Selection.Move(wdTable, 1)
                Call Globals.ThisAddIn.Application.Selection.Move(WdUnits.wdParagraph, -1)
                rng = Globals.ThisAddIn.Application.Selection.Paragraphs(1).Range
                '
                If rdBtn_TextStandard.Checked Then
                    rng.Style = "Table text"
                Else
                    rng.Style = "Table text (small)"
                End If

                'Call rng.Paragraphs.Add(rng)
                'Call rng.Move(wdParagraph, -1)
                'Set buildTable = rng
            End If
        Else
            'Set buildTable = tbl.Range
        End If

    End Sub
    '
    Public Function doSplitTable_Base(ByRef tbl As Table, ByRef dr As Row, paraStyle As Style) As Range
        Dim newTbl As Table
        Dim rng As Range
        '
        'We now have the source row
        newTbl = tbl.Split(dr)
        newTbl.Range.Select()
        Globals.ThisAddIn.Application.Selection.Collapse(WdCollapseDirection.wdCollapseStart)
        'Call Selection.Move(wdTable, 1)
        Call Globals.ThisAddIn.Application.Selection.Move(WdUnits.wdParagraph, -1)
        rng = Globals.ThisAddIn.Application.Selection.Paragraphs(1).Range
        rng.Style = paraStyle.NameLocal
        doSplitTable_Base = rng
    End Function
    '
    Public Function doTrimTable(ByRef tbl As Table, tblNoteStyle As Style) As Boolean
        'This method will delete those elements not required
        Dim dr As Row
        Dim rng As Range
        Dim para As Paragraph
        '
        doTrimTable = True
        '
        If Not Me.chkBx_UnitsRow.Checked Then Call tbl.Rows.Item(2).Delete()
        If Not Me.chkBx_HeaderRow.Checked Then Call tbl.Rows.Item(1).Delete()

        If Not Me.chkBx_Caption.Checked Then
            rng = tbl.Range
            Call rng.Collapse(WdCollapseDirection.wdCollapseStart)
            Call rng.Move(WdUnits.wdParagraph, -1)
            para = rng.Paragraphs.Item(1)
            rng = para.Range
            rng.Delete()
            rng = tbl.Range
            Call rng.Collapse(WdCollapseDirection.wdCollapseStart)
            Call rng.Move(WdUnits.wdParagraph, -1)
            para = rng.Paragraphs.Item(1)
            rng = para.Range
            para.Style = "Body Text"

            'Call tbl.Rows(1).Delete
        End If
        dr = tbl.Rows.Last
        dr = dr.Previous
        If Not Me.chkBx_Note.Checked And Not Me.chkBx_DataSource.Checked Then
            'Source row no longer exists
            Call dr.Delete()
            doTrimTable = False
        End If
        '
        If Me.chkBx_Note.Checked And Not Me.chkBx_DataSource.Checked Then
            dr.Range.Delete()                                             'Delete the Source Para
            dr.Range.Style = tblNoteStyle.NameLocal
            dr.Range.Text = "Note:"
            'Call objBBMgr.insertBuildingBlockFromDefaultLibToRange("table_sourceCellContentsNoteOnly", "Tables", rng)
        End If
        If Not Me.chkBx_Note.Checked And Me.chkBx_DataSource.Checked Then
            dr.Range.Paragraphs(2).Range.Delete()
            dr.Range.Paragraphs(1).Range.Delete()
        End If
        '
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        Call rng.Move(WdUnits.wdParagraph, 1)
        rng.Select()

        '
        'If Me.chkBx_Note.Value And Not Me.chkBx_DataSource.Value Then dr.Range.Text = "Note: "
        'If Not Me.chkBx_Note.Value And Me.chkBx_DataSource.Value Then dr.Range.Text = "Source: "
        'If Not (Me.chkBx_Note.Value) And Not (Me.chkBx_DataSource.Value) Then Call dr.Delete
        '

    End Function

    '
    Public Sub doUnitsRow(ByRef tbl As Table, colourUnits As Long)
        Dim dr As Row
        '
        dr = tbl.Rows(1)
        Call Me.Row_ColourFill(dr, colourUnits)                    'Units Row
        '
    End Sub
    '
    Public Function joinTables(ByRef tblHeader As Word.Table, ByRef tblBody As Word.Table, ByRef tblFooter As Word.Table) As Word.Table
        'This method will join the parsed components of the original Table
        '
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        Dim tbl As Word.Table
        '
        rng = tblFooter.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        Call rng.Move(WdUnits.wdParagraph, -1)
        '
        para = rng.Paragraphs.Item(1)
        para.Range.Select()
        para.Range.Delete()
        '
        tbl = Globals.ThisAddIn.Application.Selection.Range.Tables.Item(1)
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        Call rng.Move(WdUnits.wdParagraph, -1)
        '
        para = rng.Paragraphs.Item(1)
        para.Range.Select()
        para.Range.Delete()
        '
        joinTables = Globals.ThisAddIn.Application.Selection.Range.Tables.Item(1)

    End Function
    '
    '
    Public Sub Table_Fix(ByRef tbl As Table, tblWidth As Single, ByVal doBorders As Boolean)
        'This method will set the cell margins for the current Table to 0 and
        'the Borders to null (depending on the value of leavBorders). The AutoFit behaviour
        'is set the wdAutoFitFixed.. Note that settings in AutoFitBehavior will cause
        'alloAutoFit to change (see http://msdn.microsoft.com/en-us/library/office/ff820953(v=office.15).aspx)
        Dim drCol As Column
        Dim drColWidth As Single
        Dim strMsg As String
        '
        strMsg = "The Custom Table Sizing function (Table_Fix) has failed. Probable causes are;" & vbCr & vbCr &
                    "- You have selected Custom Sizing on the Table Builder for an" & vbCr &
                    "  irregular Table. An irregular Table is one in which one or" & vbCr &
                    "  more rows have merged cells" & vbCr & vbCr &
                    "The quickest solution is not to use the Custom Size Option on the Table" & vbCr &
                    "Table Builder when wrapping irregular Tables.. If you need a" & vbCr &
                    "wider Table, adjust the margins on your page to a wider " & vbCr &
                    "setting (e.g. using Toggle Width). Then, repaste the Table and" & vbCr &
                    "select AutoFit on the Table Builder"
        '
        Try

            tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
            '
            'drColWidth = tblWidth / tbl.Columns.Count
            'For Each drCol In tbl.Columns
            'drCol.width = drColWidth
            'Next
            '
            If tblWidth <= Me.widthBetweenMargins Then
                tbl.Rows.LeftIndent = 0#
            Else
                tbl.Rows.LeftIndent = -(tblWidth - Me.widthBetweenMargins)
            End If
            '
            If Me.chkBx_equalColumns.Checked Then
                drColWidth = tblWidth / tbl.Columns.Count
                For Each drCol In tbl.Columns
                    drCol.Width = drColWidth
                Next
            Else
                tbl.PreferredWidth = tblWidth
            End If

            'tbl.PreferredWidth = widthBetweenMargins
            '
            'tbl.AllowAutoFit = False
            'Call tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow)    'columns change size o accomodate text.. sets AllowAutoFit = true
            'Call tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed)      'columns sizes don't change to accommodate text, will set AllowAutoFit = false
            'Call tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)   'Table width changes to accommodate text.. Not content, then width=0.. sets AlloAutoFit = true
            '
            tbl.AllowPageBreaks = False                                           'Will allow a row to break across pages
            'Call Me.Table_doBorders_MaintainPadding(tbl, doBorders, Me.objGlobals.colour_TableBorders)

        Catch ex As Exception
            Call MsgBox(strMsg, vbInformation, "Template Message")
        End Try
        '
    End Sub
    '
    Public Sub Table_doBorders_MaintainPadding(ByRef tbl As Table, ByVal doBorders As Boolean, ByVal borderColour As Long)
        'This method will set the cell margins for the current Table to 0 and
        'the Borders on or off depending on the value of doBorders
        '
        Try
            'Set myDoc = Application.ActiveDocument
            'Set currentSect = Application.Selection.Sections(1)
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
            'Call tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed)      'columns sizes don't change to accommodate text, will set AllowAutoFit = false
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
                tbl.Borders.Item(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleNone
                'tbl.Borders.Item(wdBorderTop).LineWidth = wdLineWidth050pt
                'tbl.Borders.Item(wdBorderTop).Color = borderColour
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
                'For Each brdr In tbl.Borders
                'brdr.LineStyle = Word.WdLineStyle.wdLineStyleNone
                'Next brdr
            End If
        Catch ex As Exception
            MsgBox("Error - Table_doBorders_MaintainPadding in frmTableBuilder")
        End Try
        '
    End Sub
    '
    Public Sub Table_removeBorders(ByRef tbl As Table)
        'This method will set the cell margins for the current Table to 0 and
        'the Borders on or off depending on the value of doBorders
        Dim brdr As Border
        '
        Try
            tbl.AllowAutoFit = True
            '
            For Each brdr In tbl.Borders
                brdr.LineStyle = Word.WdLineStyle.wdLineStyleNone
            Next brdr
            '
            tbl.AllowAutoFit = False
        Catch ex As Exception
            tbl.AllowAutoFit = False
            MsgBox("Error - Table_doBorders in frmTableBuilder")
        End Try
        '
    End Sub
    '
    Public Sub Table_doBorders(ByRef tbl As Table, ByVal doBorders As Boolean, ByVal borderColour As Long)
        'This method will set the cell margins for the current Table to 0 and
        'the Borders on or off depending on the value of doBorders
        '
        Try
            'tbl.TopPadding = 0#
            'tbl.BottomPadding = 0#
            'tbl.leftPadding = 0#
            'tbl.RightPadding = 0#
            'tbl.PreferredWidth = widthBetweenMargins
            '
            tbl.AllowAutoFit = True
            'Call tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitWindow)    'columns change size o accomodate text.. sets AllowAutoFit = true
            'Call tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitFixed)      'columns sizes don't change to accommodate text, will set AllowAutoFit = false
            'Call tbl.AutoFitBehavior(Word.WdAutoFitBehavior.wdAutoFitContent)   'Table width changes to accommodate text.. Not content, then width=0.. sets AlloAutoFit = true
            '
            tbl.AllowPageBreaks = True                                           'Will allow a row to break across pages
            If doBorders Then
                tbl.Borders.Item(WdBorderType.wdBorderHorizontal).LineStyle = WdLineStyle.wdLineStyleSingle
                tbl.Borders.Item(WdBorderType.wdBorderHorizontal).LineWidth = WdLineWidth.wdLineWidth050pt
                tbl.Borders.Item(WdBorderType.wdBorderHorizontal).Color = borderColour
                tbl.Borders.Item(WdBorderType.wdBorderTop).LineStyle = WdLineStyle.wdLineStyleNone
                'tbl.Borders.Item(wdBorderTop).LineWidth = wdLineWidth050pt
                'tbl.Borders.Item(wdBorderTop).Color = borderColour
                tbl.Borders.Item(WdBorderType.wdBorderBottom).LineStyle = WdLineStyle.wdLineStyleSingle
                tbl.Borders.Item(WdBorderType.wdBorderBottom).LineWidth = WdLineWidth.wdLineWidth050pt
                tbl.Borders.Item(WdBorderType.wdBorderBottom).Color = borderColour
                'tbl.Borders.Item(wdBorderLeft).LineStyle = wdLineStyleSingle
                'tbl.Borders.Item(wdBorderLeft).LineWidth = wdLineWidth050pt
                'tbl.Borders.Item(wdBorderLeft).Color = borderColour
                'tbl.Borders.Item(wdBorderRight).LineStyle = wdLineStyleSingle
                'tbl.Borders.Item(wdBorderRight).LineWidth = wdLineWidth050pt
                'tbl.Borders.Item(wdBorderRight).Color = borderColour
            End If
            '
            tbl.AllowAutoFit = False
            '
        Catch ex As Exception
            tbl.AllowAutoFit = False
            MsgBox("Error - Table_doBorders in frmTableBuilder")
        End Try
        '
    End Sub

    Private Sub btn_Cancel_Click(sender As Object, e As EventArgs) Handles btn_Cancel.Click
        Me.Close()
    End Sub
End Class
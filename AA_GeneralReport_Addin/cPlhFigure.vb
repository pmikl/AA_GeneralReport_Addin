Imports Microsoft.Office.Interop.Word

Public Class cPlhFigure
    Inherits cPlHBase
    Public Sub New()
        MyBase.New()
        '
        'Set up everything that constitues the PlaceHolder Category of 'Figure'... These names
        'are subsets of the Seqnece Fields used in the Captions associated with each Box type
        '
        Me.lstOfPlhTypes.Clear()
        Me.lstOfPlhTypes.Add("Figure")
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will insert a Box at the current selection point. The type of Box is dependent
    ''' on the value of strType ("Figure_ES", "Figure", "Figure_AP", "Figure_LT"
    ''' </summary>
    ''' <param name="strType"></param>
    ''' <returns></returns>
    Public Function PlhFig_insert_Figure(strType As String, Optional doTableCheck As Boolean = True) As Word.Table
        Dim tbl As Word.Table
        Dim objGlobals As New cGlobals()
        Dim objMsgMgr As New cMessageManager()
        Dim numTextColumns As Integer
        Dim marginWidth As Single
        Dim sect As Word.Section
        Dim rng As Word.Range
        '
        tbl = Nothing
        rng = objGlobals.glb_get_wrdSelRng
        sect = objGlobals.glb_get_wrdSect()
        marginWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
        '
        numTextColumns = sect.PageSetup.TextColumns.Count
        '
        Try
            If doTableCheck Then
                If rng.Tables.Count = 0 Then
                    tbl = Me.Plh_insert_PlaceHolder_WithTest(objGlobals.glb_get_wrdSelRng, strType)
                    tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                    tbl.PreferredWidth = 100
                Else
                    If doTableCheck Then
                        MsgBox(objMsgMgr.msgMgr_msg_tooNearATable())
                    End If
                End If
            Else
                'Just do it with no concern reating to any tables etc
                tbl = Me.Plh_insert_PlaceHolder_WithTest(objGlobals.glb_get_wrdSelRng, strType)
                tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                tbl.PreferredWidth = 100
            End If

        Catch ex As Exception
            tbl = Nothing
        End Try
        '
        Return tbl
        '
    End Function
    '
    Public Function PlhFig_insert_FigureWide(strType As String) As Word.Table
        Dim tbl As Word.Table
        '
        tbl = MyBase.Plh_insert_PlaceHolder_Wide(strType)
        '
        Return tbl
    End Function

    '
    Public Function xPlhFig_insert_FigureWide(strType As String) As Word.Table
        '
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
        Dim strColType As String
        '
        '
        tbl = Nothing
        strColType = ""
        offSet = -40.0
        sect = objGlobals.glb_get_wrdSect
        marginWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
        '
        numTextColumns = sect.PageSetup.TextColumns.Count
        '
        Select Case numTextColumns
            Case 1
                tbl = MyBase.Plh_insert_PlaceHolderBasic(objGlobals.glb_get_wrdSelRng, strType)
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
                    MsgBox("For multi-column layouts, wide figures" + vbCrLf + "can only be inserted in column 1" + vbCrLf + vbCrLf + "Please relocate your selection point" + vbCrLf + "and try again")
                    GoTo finis
                End If
                '
                tbl = MyBase.Plh_insert_PlaceHolderBasic(objGlobals.glb_get_wrdSelRng, strType)
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
                'Globals.ThisDocument.Application.Selection.InRange(rng)
                objFloatMgr.PlHFloat_lock_toParagraphAndMarginLeft(tbl)
                'Me.Plh_Float_LockToToParagraph(tbl)
                    '
                '
            Case 3
                If Not (Me.Plh_Columnsx2_FindColumnNumber(sect) = 1) Then
                    MsgBox("For multi-column layouts, wide figures" + vbCrLf + "can only be inserted in column 1" + vbCrLf + vbCrLf + "Please relocate your selection point" + vbCrLf + "and try again")
                    GoTo finis
                End If
                '
                tbl = MyBase.Plh_insert_PlaceHolderBasic(objGlobals.glb_get_wrdSelRng, strType)
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

finis:

        '
        Return tbl

    End Function
    '

    '
#Region "Conversions"
    Public Sub PlhFig_Captions_ConvertBoxCaptionsTo_ES(ByRef rngSrc As Word.Range)
        Me.Plh_Captions_ConvertCaptions(Me.lstOfPlhTypes, "Figure_ES", rngSrc)
    End Sub
    '
    Public Sub PlhFig_Captions_ConvertBoxCaptionsTo_Report(ByRef rngSrc As Word.Range)
        Me.Plh_Captions_ConvertCaptions(Me.lstOfPlhTypes, "Figure", rngSrc)
    End Sub
    '
    Public Sub PlhFig_Captions_ConvertBoxCaptionsTo_Appendix(ByRef rngSrc As Word.Range)
        Me.Plh_Captions_ConvertCaptions(Me.lstOfPlhTypes, "Figure_AP", rngSrc)
    End Sub
    '
    Public Sub PlhFig_Captions_ConvertBoxCaptionsTo_Letter(ByRef rngSrc As Word.Range)
        Me.Plh_Captions_ConvertCaptions(Me.lstOfPlhTypes, "Figure_LT", rngSrc)
    End Sub

#End Region
    '



End Class

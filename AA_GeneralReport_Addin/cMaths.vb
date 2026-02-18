Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cMaths
    Public Sub New()

    End Sub
    '
    ''' <summary>
    ''' Thsi method will insert a table with an embedded maths editor and a 'Equation Number'
    ''' Sequence field called 'Equation'. It will add a paragraph above the table to ensure that
    ''' the tabel is clear of any other tables. The equaltion editor is selected
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <returns></returns>
    Public Function mth_equationEditor_insert(ByRef rng As Word.Range) As Word.Table
        Dim tbl As Word.Table
        Dim objTblsMgr As New cTablesMgr()
        Dim objFldsMgr As New cFieldsMgr()
        Dim fld As Word.Field
        '
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'rng.Paragraphs.Add(rng)
        'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
        'Insert a Table at the range 
        '
        tbl = objTblsMgr.tbl_build_Table_Standard(rng, 1, 2, objTblsMgr.glb_var_style_rptBodyText)
        'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPoints
        tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
        tbl.PreferredWidth = 100
        tbl.Columns.Item(2).PreferredWidth = 15
        tbl.Columns.Item(1).PreferredWidth = 100 - 15


        'tblWidth = tbl.Columns.Item(1).Width + tbl.Columns.Item(2).Width
        'tbl.Columns.Item(2).Width = 70.0
        'tbl.Columns.Item(1).Width = tblWidth - tbl.Columns.Item(2).Width
        '
        'Now do the Equation Number Sequence field
        rng = tbl.Range.Cells.Item(2).Range

        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Text = "("
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
        fld = rng.Fields.Add(rng, WdFieldType.wdFieldSequence, "Equation" + " \* ARABIC")
        fld.Select()
        rng = objTblsMgr.glb_get_wrdSelRngAll()
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng.Text = ")"
        '
        rng = tbl.Range.Cells.Item(2).Range
        rng.Style = rng.Document.Styles("Caption")
        rng.Paragraphs.Alignment = WdParagraphAlignment.wdAlignParagraphRight

        '
        'Now add the Maths editor control
        rng = tbl.Range.Cells.Item(1).Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()
        '
        Try
            rng = objTblsMgr.glb_get_wrdSel.OMaths.Add(rng)
            rng.MoveEnd(WdUnits.wdCharacter, 1)
            '
            rng.Select()
            '
        Catch ex As Exception
            MsgBox("Unknow error.. Try inserting the equation editor manually")
        End Try
        '
        objFldsMgr.flds_update_SequenceNumbers("Equation")
        '
        Return tbl
        '
    End Function



End Class

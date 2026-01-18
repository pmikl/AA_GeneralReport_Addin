Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cParas
    Public objGlobals As cGlobals
    Public Sub New()
        '
        Me.objGlobals = New cGlobals()
        '
    End Sub
    '
    ''' <summary>
    ''' This method will indent the specified paragraph
    ''' </summary>
    ''' <param name="para"></param>
    ''' <param name="leftIndent"></param>
    ''' <returns></returns>
    Public Function paras_set_HangingIndent(leftIndent As Single, ByRef para As Word.Paragraph, Optional indentSizeInPoints As Single = 65.4) As Word.Paragraph
        'Dim indentSizeInPoints As Single
        'Dim dr As Word.Row
        '
        'indentSizeInPoints = 65.4
        Try
            'indentSizeInPoints = Me.var_glb_style_tblCaption_Line2_Indent
            'dr = tbl.Rows.Item(2)
            '
            'captionParagraph.LeftIndent = dr.LeftIndent + indentSizeInPoints
            para.LeftIndent = leftIndent + indentSizeInPoints
            para.FirstLineIndent = -indentSizeInPoints
            '
        Catch ex As Exception

        End Try
        '
        Return para
        '
    End Function
    '
    ''' <summary>
    ''' This method will set the first paragraph in the rnage rng to a hanging indent of hangingIndent
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="hangingIndent"></param>
    ''' <returns></returns>
    Public Function paras_set_HangingIndent(ByRef rng As Word.Range, hangingIndent As Single) As Word.Paragraph
        Dim para As Word.Paragraph
        '
        para = rng.Paragraphs.Item(1)
        para = Me.paras_set_HangingIndent(para, hangingIndent)
        '
        '***
        'para.FirstLineIndent = -hangingIndent
        'para.LeftIndent = hangingIndent
        '
        '***
        Return para
    End Function
    '
    ''' <summary>
    ''' This method will set the specified paragraph to a hanging indent of hangingIndent
    ''' </summary>
    ''' <param name="para"></param>
    ''' <param name="hangingIndent"></param>
    ''' <returns></returns>
    Public Function paras_set_HangingIndent(ByRef para As Word.Paragraph, hangingIndent As Single) As Word.Paragraph
        '
        '***
        para.FirstLineIndent = -hangingIndent
        para.LeftIndent = hangingIndent
        '
        '***
        Return para
    End Function

    '
    '
    Public Function paras_insert_numParas(ByRef sect As Word.Section, Optional numParas As Integer = 6, Optional startFromScratch As Boolean = False) As Word.Range
        Dim rng As Word.Range
        '
        rng = sect.Range
        '
        'rng.Text = vbCrLf
        rng.Style = rng.Document.Styles.Item("Body Text")
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        For i = 1 To numParas
            rng.Paragraphs.Add(rng)
        Next
        '
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        Return rng
        '
    End Function
    '
    ''' <summary>
    ''' This method will place a numParas empty paragraphs after the last paragraph in the specified
    ''' range (rng). The last paragraph is returned and the selection is adjusted to the start of the
    ''' last paragraph. The byref rng is adjusted to the selection range which is now at the beginning
    ''' of the last paragraph
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="numParas"></param>
    ''' <returns></returns>
    Public Function paras_insert_parasAfterSelectedPara(ByRef rng As Word.Range, numParas As Integer) As Word.Paragraph
        Dim para As Word.Paragraph
        '
        para = rng.Paragraphs.Item(1)
        For i = 1 To numParas
            para = paras_insert_paraAfter(rng)
        Next
        '
        rng = para.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()
        '
        Return para
        '
    End Function
    '
    ''' <summary>
    ''' This method will insert numParas at the selection point.. So, when executed there will always be numParas blank lines
    ''' after the initial selection point... The selection is at the beginning of the last paragraph. The return value is the
    ''' last paragraph
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="numParas"></param>
    ''' <param name="strStyleName"></param>
    ''' <returns></returns>
    Public Function paras_insert_parasAtSelection(ByRef rng As Word.Range, Optional numParas As Integer = 2, Optional strStyleName As String = "Body Text") As Word.Paragraph
        Dim para As Word.Paragraph = Nothing
        Dim rslt As Boolean = False
        '
        If Me.paras_selis_atEndOfPara() Then
            'MsgBox("Select is at end of para")
            numParas = numParas - 1
            rng.Paragraphs.Add(rng)
            rng.MoveStart(WdUnits.wdParagraph, 1)
            rng.Select()
            '
            For i = 1 To numParas
                para = paras_insert_paraAfter(rng)
                rng = para.Range
            Next

        Else
            rng.Paragraphs.Add(rng)
            For i = 1 To numParas
                para = paras_insert_paraAfter(rng)
            Next
            '
        End If
        '
        rng = para.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()
        '
        Return para
    End Function
    '
    ''' <summary>
    ''' This method will place a new empty paragraph after the last paragraph in the specified
    ''' range (rng). That paragraph is returned and the selection is adjusted to the start of the
    ''' new paragraph. The byref rng is adjusted to the selection range which is at the beginning
    ''' of the paragraph
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <returns></returns>
    Public Function paras_insert_paraAfter(ByRef rng As Word.Range) As Word.Paragraph
        Dim para As Word.Paragraph

        para = rng.Paragraphs.Last
        rng = para.Range
        '
        'para.Range.Delete()
        'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        rng.InsertParagraphAfter()
        objGlobals.glb_get_wrdSel().Move(WdUnits.wdParagraph, 1)
        objGlobals.glb_get_wrdSel().Collapse(WdCollapseDirection.wdCollapseStart)
        para = objGlobals.glb_get_wrdSel().Paragraphs.Item(1)
        '
        rng = objGlobals.glb_get_wrdSelRng()
        '
        Return para
        '
    End Function
    '
    ''' <summary>
    ''' This method will first delete all paragraphs ina section, then add back numParas of empty
    ''' paragraphs
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="numParas"></param>
    ''' <returns></returns>
    Public Function paras_insertAfterDelete_numParas(ByRef sect As Word.Section, Optional numParas As Integer = 6) As Word.Range
        Dim rng As Word.Range
        '
        rng = Me.paras_Paragraphs_DeleteAll(sect)
        Me.paras_insert_numParas(rng, numParas - 1)
        '
        Return rng
    End Function
    '
    ''' <summary>
    ''' This method will insert numParas in front of the paragraph containing the
    ''' selection. The inserted paragraphs will be set to the 'Body Text' style. The initial
    ''' paragraph will be left unchanged
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="numParas"></param>
    ''' <returns></returns>
    Public Function paras_insert_numParas(ByRef rng As Word.Range, Optional numParas As Integer = 6) As Word.Range
        '
        'rng.Style = rng.Document.Styles.Item("Body Text")          'This was here until 20240716.. Leaving it here hnaged the style of the inital para
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        For i = 1 To numParas
            rng.Paragraphs.Add(rng)
        Next
        '
        rng.Style = rng.Document.Styles.Item("Body Text")           'Moved to here 20240716. Ensures only the added paras are set to 'Body Text'
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        Return rng
        '
    End Function
    '
    ''' <summary>
    ''' This method will retrieve all paragrahs after a Table that are 'styled' with
    ''' 'Source' or 'Note'
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function paras_get_SourceAndNote(ByRef tbl As Word.Table) As String
        Dim strRslt As String
        Dim myDoc As Word.Document
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        Dim stylSrc, stylNote, paraStyle As Word.Style
        '
        myDoc = tbl.Range.Document()
        strRslt = ""
        stylSrc = myDoc.Styles.Item("Source")
        stylNote = myDoc.Styles.Item("Note")
        '
        rng = tbl.Range
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
        para = rng.Paragraphs.Item(1)
        rng = para.Range
        '
loop00:
        paraStyle = para.Range.Style
        '
        If Not (paraStyle.NameLocal = stylSrc.NameLocal Or paraStyle.NameLocal = stylNote.NameLocal) Then
            If strRslt = vbCrLf Then
                rng = Nothing
            Else
                rng.MoveEnd(WdUnits.wdParagraph, -1)
            End If
            GoTo finis
        End If
        strRslt = strRslt + para.Range.Text
        para = para.Next
        rng.MoveEnd(WdUnits.wdParagraph, 1)

        '
        GoTo loop00
        '
finis:
        Return strRslt
        '
    End Function

    '
    '
    ''' <summary>
    ''' This method will delete all of the last empty paragraphs in the specified Cell
    ''' </summary>
    ''' <param name="drCell"></param>
    ''' <returns></returns>
    Public Function paras_delete_lastParasInTableCell(ByRef drCell As Word.Cell) As Word.Range
        Dim para As Word.Paragraph
        Dim rng As Word.Range
        Dim strText As String
        Dim j As Integer
        'Dim tokens As String()
        '
        j = 0
        rng = drCell.Range
        '
        Try
            For j = rng.Paragraphs.Count To 1 Step -1
                para = rng.Paragraphs.Item(j)
                strText = Trim(para.Range.Text)
                If strText = vbCr & ChrW(7) Then
                    para.Range.Text = ChrW(7)
                    Continue For
                End If
                If strText = vbCr Then para.Range.Delete()
                If strText <> vbCr Then Exit For
            Next
            '
            'Move to the end of cell end circle, then backup to the beginning
            'of the end circle
            rng = drCell.Range
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            rng.Move(WdUnits.wdParagraph, -1)
            '
            'We are now at the beginning of the end of cell circle
            'Move the range start back one char and if the range text is vbCr then we  have the
            'end cell situation of vbCR + end of cell circle.. So just delete the vbCr to make sure
            'that the cell text sits right at the end of the cell
            '
            rng.MoveStart(WdUnits.wdCharacter, -1)
            If rng.Text = vbCr Then rng.Delete()
            'tokens = rng.Text.Split(vbCr)
            'strText = tokens(0)
        Catch ex As Exception

        End Try
        '
        Return rng
        '
    End Function

    '
    ''' <summary>
    ''' This method will find the last Table in the Range rng, and then delete all but the
    ''' n paragraphs between this table and the end of the section.. It will change rng to the
    ''' beginning of the first paragraph after the Table... If no Table, then it will delete all
    ''' paragraphs in the section, except for the specified numParasLeft.. 
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="numParasLeft"></param>
    Public Function paras_delete_Paragraphs(ByRef rng As Word.Range, Optional numParasLeft As Integer = 6, Optional strStyleOfParas As String = "Body Text") As Word.Range
        Dim tbl As Word.Table
        Dim i As Integer
        Dim sect, sectLast As Word.Section
        Dim para As Word.Paragraph
        Dim myDoc As Word.Document
        '
        sect = rng.Sections.Item(1)
        myDoc = rng.Document
        '
        sectLast = myDoc.Sections.Last
        '
        If rng.Tables.Count <> 0 Then
            tbl = rng.Tables.Item(rng.Tables.Count)
            'tbl = rng.Tables.Item(1)
            rng = tbl.Range
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            '
            If sect.Index = sectLast.Index Then
                '
                rng.MoveEnd(WdUnits.wdStory)
                rng.Delete()
                rng.Style = myDoc.Styles("Body Text")

            Else
                'We are in a standard section. So delete paragraphs to the section boundary
                'and then add numParasLeft
                '
                'rng = sect.Range
                rng.MoveEnd(WdUnits.wdSection, 1)
                rng.MoveEnd(WdUnits.wdParagraph, -2)
                rng.Style = myDoc.Styles("Body Text")
                rng.Delete()
                For i = 1 To numParasLeft
                    para = rng.Paragraphs.Add()
                Next
                '
                rng.Style = myDoc.Styles(strStyleOfParas)
                '
                rng = tbl.Range
                rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            End If
            '
        Else
            'No tables, so we'll delete from the current range to the end of the section
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.MoveEnd(WdUnits.wdSection, 1)
            rng.MoveEnd(WdUnits.wdParagraph, -2)
            rng.Delete()
            '
            For i = 1 To numParasLeft
                rng.Paragraphs.Add(rng)
            Next
            '
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
        End If
        '
        Return rng
        '
    End Function

    '
    '
    ''' <summary>
    ''' This method will delet all paragraphs in the specified section. It will put back
    ''' an empty paragraph and leave the selection there. The returned Range is the range
    ''' of the selection 
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function paras_Paragraphs_DeleteAll(ByRef sect As Word.Section) As Word.Range
        Dim rng As Word.Range
        '
        rng = sect.Range()
        If sect.Index = Globals.ThisAddin.Application.ActiveDocument.Sections.Last.Index Then
            rng.Delete()
            rng.Select()
        Else
            'A section other than the last section
            rng.MoveEnd(WdUnits.wdParagraph, -2)
            rng.Delete()
            rng.Paragraphs.Add(rng)
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Select()
        End If
        'rng.Delete(WdUnits.wdParagraph, rng.Paragraphs.Count)
        '
        Return rng
        '
    End Function
    '
    ''' <summary>
    ''' This method will insert the text strText at the beginning od the section and it will
    ''' then apply the style strStyleName
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="strText"></param>
    ''' <param name="strStyleName"></param>
    Public Sub paras_add_textAndStyle(ByRef sect As Word.Section, strText As String, strStyleName As String)
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        Dim myDoc As Word.Document
        '
        myDoc = sect.Range.Document
        '
        Try
            rng = sect.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.MoveEnd(WdUnits.wdParagraph, 1)
            para = rng.Paragraphs.Item(1)
            para.Range.Text = strText
            para.Range.Style = myDoc.Styles.Item(strStyleName)
        Catch ex As Exception

        End Try

    End Sub
    '
    ''' <summary>
    ''' This method will insert the text strText at the beginning od the range and it will
    ''' then apply the style strStyleName
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="strText"></param>
    ''' <param name="strStyleName"></param>
    Public Function paras_add_textAndStyle(ByRef rng As Word.Range, strText As String, strStyleName As String) As Word.Paragraph
        Dim para As Word.Paragraph
        Dim myDoc As Word.Document
        '
        myDoc = rng.Document
        '
        Try
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Text = strText
            'rng.MoveEnd(WdUnits.wdParagraph, 1)
            para = rng.Paragraphs.Item(1)
            'para.Range.Text = strText
            para.Range.Style = myDoc.Styles.Item(strStyleName)
            '
            rng = para.Range

        Catch ex As Exception
            para = Nothing
        End Try
        '
        Return para
    End Function
    '
    Public Function paras_selis_atEndOfPara() As Boolean
        Dim para As Word.Paragraph
        Dim rslt As Boolean = False
        '
        para = objGlobals.glb_get_wrdSelRng.Paragraphs.First

        If objGlobals.glb_get_wrdSel.Start = (para.Range.End - 1) Then rslt = True

        Return rslt
    End Function
    '
End Class

Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
''' <summary>
''' 
'''This class deals witht the installation and removal of
'''Caption labels in the parent Application.. These labels must
'''be resident to allow cross referencing of Tables/Figures/Boxes
'''in the ES, Body and Appendices. Current Acil Allen policy causes
'''Word Application on user machines to be renewed via a company log-in
'''script.. Consequently Custom Cpation labels disappear
'''
'''Peter Mikelaitis October 2015...http://mikl.com.au
'''Ported to VB.NET 17th Jan 2017 from version 97p21p05
'''
''' </summary>
Public Class cCaptionManager
    Public Sub New()

    End Sub
    '
    Public Function cpt_indent_Caption(ByRef tbl As Word.Table, captionParagraph As Word.Paragraph) As Word.Paragraph
        'This method will offset the leftindent of (generally) the Caption
        'Paragraph to match any Autofit functions.. tblWidth must be in points
        Dim objTools As cTools
        Dim rng As Word.Range
        Dim dr As Word.Row
        Dim indentSizeInPoints As Single
        Dim tblWidth As Single
        '
        '
        indentSizeInPoints = 65.4

        '
        objTools = New cTools()
        rng = captionParagraph.Range
        '
        If tblWidth <= objTools.widthBetweenMargins Then
            captionParagraph.LeftIndent = indentSizeInPoints
            captionParagraph.FirstLineIndent = -indentSizeInPoints
        Else
            dr = tbl.Rows.Item(2)
            captionParagraph.LeftIndent = dr.LeftIndent + indentSizeInPoints
            captionParagraph.FirstLineIndent = -indentSizeInPoints
        End If
        '
        Return captionParagraph
        '
    End Function
    '
    ''' <summary>
    ''' This method will look at the paragraph just above the table.. If it exists and has either
    ''' 'Caption' or 'Caption Label' style, then this paragraph is deemed to be a Caption paragraph
    ''' and is returned.. If not, this method returns Nothing... If doIndent is true, then the paraLeftIndent
    ''' (which is normally 0 or a negative number) is used to indent the Caption Paragraph
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <param name="paraLeftIndent"></param>
    ''' <param name="doIndent"></param>
    ''' <returns></returns>
    Public Function cpt_indent_CaptionParagraph(ByRef tbl As Word.Table, ByRef paraLeftIndent As Single, doIndent As Boolean) As Word.Paragraph
        Dim para As Word.Paragraph
        Dim rng As Word.Range
        Dim styl As Word.Style
        Dim drCell As Word.Cell
        Dim padOffset As Single
        '
        drCell = tbl.Range.Cells.Item(1)
        padOffset = drCell.LeftPadding                              'Allows us to adjust the Caption to align with the text in the first columnif there is padding
        '
        para = Nothing
        Try
            rng = tbl.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            rng.Move(WdUnits.wdParagraph, -1)
            '
            If rng.Paragraphs.Count > 0 Then
                para = rng.Paragraphs.Item(1)
                styl = para.Style
                'If styl.NameLocal = "Caption" Or styl.NameLocal = "Caption Label" Then
                If para.Range.Text Like "Tabl*" Then

                    para = rng.Paragraphs.Item(1)
                    If doIndent Then
                        para.LeftIndent = paraLeftIndent - para.FirstLineIndent
                        'Now adjust for cell padding just in case we are dealing with a legacy AAC table
                        If drCell.LeftPadding > 1 Then
                            If paraLeftIndent = 0.0 Then
                                'para.LeftIndent = 0.0
                            Else
                                'para.LeftIndent = paraLeftIndent - para.FirstLineIndent + padOffset
                                para.LeftIndent = paraLeftIndent - para.FirstLineIndent + padOffset
                            End If
                        End If
                    End If
                End If
                '
            End If


        Catch ex As Exception
            para = Nothing
        End Try
        '
        Return para
    End Function
    '
    Public Function cpt_indent_SourceNoteParagraph(ByRef tbl As Word.Table, ByRef paraLeftIndent As Single, doIndent As Boolean) As Word.Paragraph
        Dim para As Word.Paragraph
        Dim rng As Word.Range
        Dim styl As Word.Style
        Dim drCell As Word.Cell
        Dim padOffset As Single
        Dim j As Integer
        '
        drCell = tbl.Range.Cells.Item(1)
        padOffset = drCell.LeftPadding                              'Allows us to adjust the Caption to align with the text in the first columnif there is padding
        '
        para = Nothing
        Try
            rng = tbl.Range
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            'rng.Move(WdUnits.wdParagraph, -1)
            '
            If rng.Paragraphs.Count > 0 Then
                para = rng.Paragraphs.Item(1)
                '
                'If at the end of a document we can run out of paragraphs
                '
                Try
                    For j = 1 To 4
                        styl = para.Style
                        If para.Range.Text Like "Source*" Or para.Range.Text Like "Note*" Then
                            'If styl.NameLocal = "Source" Or styl.NameLocal = "Note" Then
                            'para = rng.Paragraphs.Item(1)
                            If doIndent Then
                                para.LeftIndent = paraLeftIndent - para.FirstLineIndent
                                'Now adjust for cell padding just in case we are dealing with a legacy AAC table
                                If drCell.LeftPadding > 1 Then
                                    If paraLeftIndent = 0.0 Then
                                        'para.LeftIndent = 0.0
                                    Else
                                        'para.LeftIndent = paraLeftIndent - para.FirstLineIndent + padOffset
                                        para.LeftIndent = paraLeftIndent - para.FirstLineIndent + padOffset
                                    End If
                                End If
                            End If
                        End If
                        '
                        para = para.Next
                    Next j

                Catch ex2 As Exception

                End Try
            End If
            '
        Catch ex As Exception
            para = Nothing
        End Try
        '
        Return para
    End Function
    '
    ''' <summary>
    ''' This method will set all of the Cross Reference fields to the font used by the
    ''' 'Body Text' style. This is necessary since the 2024 ersion of the template uses
    ''' Yu Gothic Medium for the Caption style. When cross referencing the cross references
    ''' pick up the style of the Caption and embed this in the body of the text
    ''' </summary>
    ''' <param name="setAsBold"></param>
    Public Sub cpt_setCrossRef_FieldsToBodyTextFont(Optional setAsBold As Boolean = True)
        Dim fld As Word.Field
        Dim tmpString As String
        Dim objGlobals As New cGlobals()
        Dim myDoc As Word.Document
        Dim bodyTextStyle As Word.Style
        '
        myDoc = objGlobals.glb_get_wrdActiveDoc()
        Try
            bodyTextStyle = myDoc.Styles.Item("Body Text")

            For Each fld In objGlobals.glb_get_wrdActiveDoc.Fields
                'strFldName = Trim(fld.n
                If fld.Type = WdFieldType.wdFieldRef Then
                    fld.Select()
                    'Remove any Character Style from the Reference
                    'Then make sure that the result is not bold
                    '
                    '*** AlexR fix 20151028 - preserve formatting on refresh
                    If InStr(fld.Code.Text, "\* MERGEFORMAT") = 0 Then
                        tmpString = fld.Code.Text & " \* MERGEFORMAT "
                        fld.Code.Text = tmpString
                    End If
                    '*** end fix
                    '
                    objGlobals.glb_get_wrdSel.ClearFormatting()
                    objGlobals.glb_get_wrdSel.Style = bodyTextStyle
                    objGlobals.glb_get_wrdSel.Font.Bold = setAsBold
                    'objGlobals.glb_get_wrdSel.Font.Name = bodyTextStyle.NameLocal
                    'Selection.Range.Font.Bold = False
                    'strFldText = Selection.Text
                End If
            Next fld
        Catch ex As Exception

        End Try
        '

    End Sub

    '
    Public Sub setFieldsBoldStatus(setAsBold As Boolean)
        Dim fld As Word.Field
        Dim tmpString As String
        Dim objGlobals As New cGlobals()
        Dim myDoc As Word.Document
        Dim bodyTextStyle As Word.Style
        '
        myDoc = objGlobals.glb_get_wrdActiveDoc()
        bodyTextStyle = myDoc.Styles.Item("Body Text")


        For Each fld In objGlobals.glb_get_wrdActiveDoc.Fields
            'strFldName = Trim(fld.n
            If fld.Type = WdFieldType.wdFieldRef Then
                fld.Select()
                'Remove any Character Style from the Reference
                'Then make sure that the result is not bold
                '
                '*** AlexR fix 20151028 - preserve formatting on refresh
                If InStr(fld.Code.Text, "\* MERGEFORMAT") = 0 Then
                    tmpString = fld.Code.Text & " \* MERGEFORMAT "
                    fld.Code.Text = tmpString
                End If
                '*** end fix
                '
                objGlobals.glb_get_wrdSel.ClearFormatting()
                objGlobals.glb_get_wrdSel.Style = bodyTextStyle
                objGlobals.glb_get_wrdSel.Font.Bold = setAsBold
                'Selection.Range.Font.Bold = False
                'strFldText = Selection.Text
            End If
        Next fld
        '
    End Sub
    '
    Public Sub setupCaptions()
        Dim para As Paragraph
        Dim currentDoc As Word.Document
        Dim objGlobals As New cGlobals()
        '
        'Initial selection is necessary other the objCaptionsMgr function
        'fail with a Not in XML block function
        '
        On Error GoTo final
        currentDoc = objGlobals.glb_get_wrdActiveDoc()
        '
        For Each para In currentDoc.Paragraphs
            'If para.Style = currentDoc.Styles("Body Text") Then
            If para.Style.NameLocal = "Body Text" Then
                para.Range.Select()
                objGlobals.glb_get_wrdSel.Collapse(WdCollapseDirection.wdCollapseStart)
                GoTo finis
            End If
        Next para
        '
finis:
        'Now having made a current Selection in the document we'll
        'modify the current captions
        Call Me.deleteAllNotBuiltInCaptions()
        Call Me.installCustomCaptions()
        Exit Sub
        '
final:
        objGlobals.glb_get_wrdSel.WholeStory()
        objGlobals.glb_get_wrdSel.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
        Call Me.deleteAllNotBuiltInCaptions()
        Call Me.installCustomCaptions()
        '
    End Sub
    '
    Public Sub deleteCaption(strCaptionName)
        Dim lbl As Word.CaptionLabel
        Dim objGlobals As New cGlobals()
        '
        If Me.captionExists(strCaptionName) Then
            lbl = objGlobals.glb_get_wrdApp.CaptionLabels(strCaptionName)
            Call lbl.Delete()
        End If
    End Sub
    '
    Public Sub deleteAllNotBuiltInCaptions()
        Dim lbl As Word.CaptionLabel
        Dim objGlobals As New cGlobals()
        '
        For Each lbl In objGlobals.glb_get_wrdApp.CaptionLabels
            If Not lbl.BuiltIn Then lbl.Delete()
        Next lbl
        '
        'This following doesn't work in the prior loop.. the delete
        'causes a problem
        For Each lbl In objGlobals.glb_get_wrdApp.CaptionLabels
            If lbl.Name = "Table" Or lbl.Name = "Figure" Then
                lbl.IncludeChapterNumber = True
                lbl.ChapterStyleLevel = 1
                lbl.Separator = WdSeparatorType.wdSeparatorPeriod
            End If
        Next lbl

    End Sub
    '
    Public Sub deleteCustomCaptions()
        'This method will clear out all Custom Captions.. To be used
        'if you want to start from a clean slate
        'Call Me.deleteCaption("Table ES")
        'Call Me.deleteCaption("Figure ES")
        'Call Me.deleteCaption("Box ES")
        'Call Me.deleteCaption("Key Finding ES")
        'Call Me.deleteCaption("Recommendation")
        'Call Me.deleteCaption("RECOMMENDATION ES")
        '
        'Call Me.deleteCaption("Table AP")
        'Call Me.deleteCaption("Figure AP")
        'Call Me.deleteCaption("Box AP") 
        '
        Try
            Call Me.deleteCaption("CaseStudy")
            Call Me.deleteCaption("Table ES")
            Call Me.deleteCaption("Figure ES")
            Call Me.deleteCaption("Box ES")
            Call Me.deleteCaption("Key Finding ES")
            Call Me.deleteCaption("Finding ES")
            Call Me.deleteCaption("Recommendation ES")
            '
            'LT
            Call Me.deleteCaption("Table LT")
            Call Me.deleteCaption("Figure LT")
            Call Me.deleteCaption("Box LT")
            Call Me.deleteCaption("Key Finding LT")
            Call Me.deleteCaption("Finding LT")
            Call Me.deleteCaption("Recommendation LT")
            '
            Call Me.deleteCaption("Box")
            Call Me.deleteCaption("Recommendation")
            Call Me.deleteCaption("Key Finding")
            Call Me.deleteCaption("Finding")
            '
            'Appendix
            Call Me.deleteCaption("Table AP")
            Call Me.deleteCaption("Figure AP")
            Call Me.deleteCaption("Box AP")

        Catch ex As Exception
            MsgBox("Fault in delete custom captions")
        End Try
        '
    End Sub
    '
    Public Sub installCustomCaptions()
        'It is important to remember that Caption labels are selected
        '(for cross referencing) on the string value of the numbering
        'Sequence tag and NOT on the string value of the Caption. Hence,
        'when you create a Cpation Label Table AP. It's numbering sequence tag
        'is Table_AP... For cross referencing any table using a numbering sequence
        'tag of Table_AP will appear in the list even though its Caption Laebl may be
        'Table and not Table AP
        Dim lbl As CaptionLabel
        Dim objGlobals As New cGlobals()
        Dim Application As Word.Application
        '
        Application = objGlobals.glb_get_wrdApp
        'ES
        Call Application.CaptionLabels.Add("CaseStudy")
        Call Application.CaptionLabels.Add("Table ES")
        Call Application.CaptionLabels.Add("Figure ES")
        Call Application.CaptionLabels.Add("Box ES")
        Call Application.CaptionLabels.Add("Key Finding ES")
        Call Application.CaptionLabels.Add("Finding ES")
        Call Application.CaptionLabels.Add("Recommendation ES")
        '
        'LT
        Call Application.CaptionLabels.Add("Table LT")
        Call Application.CaptionLabels.Add("Figure LT")
        Call Application.CaptionLabels.Add("Box LT")
        Call Application.CaptionLabels.Add("Key Finding LT")
        Call Application.CaptionLabels.Add("Finding LT")
        Call Application.CaptionLabels.Add("Recommendation LT")
        '
        Call Application.CaptionLabels.Add("Box")
        Call Application.CaptionLabels.Add("Recommendation")
        Call Application.CaptionLabels.Add("Key Finding")
        Call Application.CaptionLabels.Add("Finding")
        '
        'Appendix
        Call Application.CaptionLabels.Add("Table AP")
        Call Application.CaptionLabels.Add("Figure AP")
        Call Application.CaptionLabels.Add("Box AP")
        '
        'Can't do the following because this is for 'Caption Style' so to
        'cross reference custom Heading must go to 'Numbered Items'
        '
        'Call Application.CaptionLabels.Add("Heading 1 (AP)")
        'Call Application.CaptionLabels.Add("Heading 2 (AP)")
        'Call Application.CaptionLabels.Add("Heading 3 (AP)")

        '
        'This will select TABLE X.Y so long as Y has a sequence label of
        'TABLE_AP.. Now we'll set the numbering options
        For Each lbl In Application.CaptionLabels
            If lbl.Name = "Table AP" Or lbl.Name = "Figure AP" Or lbl.Name = "Box_AP" Then
                lbl.IncludeChapterNumber = True
                lbl.ChapterStyleLevel = 9
                lbl.Separator = WdSeparatorType.wdSeparatorPeriod
            End If
        Next lbl

    End Sub
    '
    Public Function captionExists(strCaptionName) As Boolean
        Dim lbl As Word.CaptionLabel
        Dim lbls As Word.CaptionLabels
        Dim objGlobals As New cGlobals()
        '
        captionExists = False
        lbls = objGlobals.glb_get_wrdApp.CaptionLabels
        '
        For Each lbl In objGlobals.glb_get_wrdApp.CaptionLabels
            If lbl.Name = strCaptionName Then
                captionExists = True
                GoTo finis
            End If
        Next lbl
finis:
    End Function
    '
    ''' <summary>
    '''This method converts all of the various TABLE Captions
    '''Note that the order of the Like statements matter.. TABLE*
    '''must be last otherwise it gathers in all of the others
    '''
    '''i is the position of the delimeter    ''' 
    '''</summary>
    ''' <param name="i"></param>
    ''' <param name="para"></param>
    Public Sub reformat_Captions(i As Integer, ByRef para As Word.Paragraph)
        Dim rng As Word.Range
        Dim objGlobals As New cGlobals()
        '
        'If Not (para.Range.Text Like "TABLE*") Then Exit Sub
        '
        Call para.Range.Characters.Item(i - 2).Select()
        Call para.Range.Characters.Item(1).Select()
        '
        objGlobals.glb_get_wrdApp.Selection.MoveEnd(WdUnits.wdCharacter, (i - 3))
        objGlobals.glb_get_wrdApp.Selection.Style = objGlobals.glb_get_wrdActiveDoc.Styles("Caption Label")
        '
        para.Range.Characters.Item(i + 1).Delete()
        para.Range.Characters.Item(i).Delete()
        '
        'para.Range.Characters.Item(i - 1) = ChrW(&H9)      'VBA version
        rng = para.Range.Characters.Item(i - 1)             'update to VB.NET
        rng.Text = ChrW(&H9)
    End Sub
    '
    '
    Public Function findDelimiter(ByRef para As Word.Paragraph) As Long
        'This function will search for the first occurance of the delimeter
        'character in the paragraphh text.. If the delimeter is not found it
        'will return -1
        '
        Dim i As Integer
        Dim chrs As Char()
        Dim strPara As String
        'Dim myChar As char
        'Dim chrDelim As char
        'chrDelim = ChrW(&H2013)
        '
        findDelimiter = -1
        '
        For i = 1 To para.Range.Characters.Count
            'If para.Range.Characters.Item(i) = ChrW(&H2013) Then GoTo finis                'VBA original
            '
            'VB.Net change
            strPara = para.Range.Text
            chrs = strPara.ToCharArray()                                'Character Array is zero based??
            If chrs(i - 1) = ChrW(&H2013) Then GoTo finis                'VBA original

            'delim = CLng(para.Range.Characters.Item(i))
            'If delim = &H2013 Then GoTo finis
        Next
        Exit Function
finis:
        findDelimiter = i
        '
    End Function

    '
    Public Sub convertCaptionToTab()
        'This method will convert existing Caption style text that is in
        'the 'Table xxxx - xxxx' format to a tab format (i.e. the space tab space
        'is changed to tab... The template has been chnaged to support the new
        'tab separated Caption, so this procedure is uncluded so that existing
        'Captions in legacy documents can be changed ove
        '
        Dim objGlobals As New cGlobals()
        Dim rng As Range
        Dim tbl As Word.Table
        Dim para As Paragraph
        Dim strReplace As String
        'Dim tokens() As String
        Dim i As Integer
        '
        rng = objGlobals.glb_get_wrdActiveDoc.Range
        'chrDelim = ChrW(&H2013)
        strReplace = CStr(ChrW(&H9))
        '
        'Set para = Selection.Range.Paragraphs.Item(1)
        'para.Range.Select
        'Selection.ClearFormatting
        'Selection.style = "Caption"
        'Call Me.reformat_Captions(para)

        '
        'ActiveDocument.DefaultTabStop = 18#
        For Each para In objGlobals.glb_get_wrdActiveDoc.Paragraphs
            'If para.Style = Globals.ThisDocument.Application.ActiveDocument.Styles("Caption") Then          VBA version
            If para.Style.NameLocal = "Caption" Then         'VB.Net version
                '****
                'Call Me.reConvertCaption(para)
                '
                i = Me.findDelimiter(para)
                'If i is negative we couldn't find the delimeter
                If i > 0 Then
                    para.Range.Select()
                    objGlobals.glb_get_wrdSel.ClearFormatting()
                    objGlobals.glb_get_wrdSel.Style = objGlobals.glb_get_wrdActiveDoc.Styles("Caption")
                    Call Me.reformat_Captions(i, para)
                    'Call Me.reformat_TABLE_Captions(para)
                    'Call Me.reformat_FIGURE_Captions(para)
                    'Call Me.reformat_BOX_Captions(para)
                End If
                If para.Range.Tables.Count > 0 Then
                    If para.Range.Text Like "KEY*" Or para.Range.Text Like "BOX*" Or para.Range.Text Like "RECOMMENDATION*" Then
                        tbl = para.Range.Tables.Item(1)
                        If tbl.Rows.LeftIndent = 0# Then tbl.Rows.LeftIndent = -22.3
                    End If
                    'if tbl.le
                End If
                objGlobals.glb_get_wrdApp.ScreenRefresh()
                'Globals.ThisDocument.Application.ScreenRefresh()
            End If
        Next para
        '
    End Sub
    '
    Public Sub reConvertCaption(ByRef para As Word.Paragraph)
        'We run the standard conversion functions just to make sure that
        'the Cpation is in the right format to be seen by the standard
        'Cross Reference facility
        Dim fld As Word.Field
        Dim strCode As String
        Dim x As Integer
        'Dim objPlhMgr As cPlhManager
        '
        If para.Range.Fields.Count = 0 Then Exit Sub
        '
        'objPlhMgr = New cPlhManager
        fld = para.Range.Fields.Item(para.Range.Fields.Count)
        strCode = fld.Code.Text
        strCode = Trim(strCode)
        strCode = StrConv(strCode, vbUpperCase)         'For consistency
        If strCode Like "SEQ RECOMMENDATION_ES*" Then
            x = 1
            'Call objPlhMgr.convertBoxFigTableBase(para, "box", "ES")
        End If
        '
    End Sub
    '
    '
    '************* Spare Code (may be useful in speeding it up)
    '
    'For Each para In ActiveDocument.Paragraphs
    'If para.style = ActiveDocument.Styles("Caption") Then
    'Call para.Range.Editors.Add(wdEditorEveryone)
    'para.Range.Select
    'Selection.ClearFormatting
    'Selection.style = "Caption"
    'If para.Range.Text Like "TABLE ES*" Then
    'para.Range.Characters.Item(13).Delete
    'para.Range.Characters.Item(12).Delete
    'para.Range.Characters.Item(11) = ChrW(&H9)
    'para.Range.Characters.Item(1).Select
    'Call Selection.MoveEnd(wdCharacter, 9)
    'Selection.style = "Caption Label"
    'GoTo finis
    'End If
    'If para.Range.Text Like "TABLE *" Then

    'End If

    'Set rng = para.Range
    'rng.Collapse (wdCollapseStart)
    'Set newPara = ActiveDocument.Paragraphs.Add(rng)
    'newPara.Range.Text = para.Range.Text
    'If para.Range.Cells.Count <> 0 Then
    'Set drCell = para.Range.Cells.Item(1)
    'Call drCell.Range.Paragraphs.Add(para.Range)
    'End If
    'Set rng = Application.unio
    'para.Range.HighlightColorIndex = wdBrightGreen
    'End If
    'finis:
    'Next para
    '
    'For Each drCell In rng.Cells
    'Set rngCell = drCell.Range
    'With rngCell.Find
    '.Text = " "
    '.Replacement.Text = ""
    '.Forward = True
    '.Wrap = wdFindAsk = False
    'End With
    'Selection.Find.Execute Replace:=wdReplaceAll
    'Next drCell

    '
    'With rng.Find
    '.style = "Caption"
    '.Forward = True
    '.Wrap = wdFindStop
    'Do
    'foundItem = .Execute
    'If foundItem Then
    'If found then add the paragraph to the collection
    'and then collapse the range to the end so that
    'when we continue it will not repeat the current paragraph
    'Set para = rng.Paragraphs.Item(1)
    'Set rngCaption = para.Range
    'Call rngCaption.MoveEnd(wdCharacter, -1)
    'strText = rngCaption.Text
    'tokens = Split(strText, strDelim)
    'tokens = Split(strText, " – ")
    'If UBound(tokens) >= 1 Then
    'para.Range.Font.ColorIndex = wdRed
    'rngCaption.Font.ColorIndex = wdRed
    'para.Range.Text = tokens(0) & ChrW(&H9) & tokens(1) & vbCrLf
    'End If
    'strText = para.Range.Text
    'para.Range.Text = Replace(strText, strDelim, strReplace)
    'Set rngCaption = para.Range
    'With rngCaption.Find
    '.Text = strDelim
    '.Replacement.Text = strReplace
    '.Wrap = wdFindStop
    'End With
    'strText = StrConv(para.Range.Text, vbLowerCase)
    'If strText Like "table*" Then Call lstOfTables.Add(para)
    'If strText Like "figure*" Then Call lstOfFigures.Add(para)
    'If strText Like "box*" Then Call lstOfBoxes.Add(para)
    '
    'rng.Collapse (wdCollapseEnd)
    'Else
    'Exit Do
    'End If
    'Loop
    'End With


End Class

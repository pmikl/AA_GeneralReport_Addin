Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cTools
    Public Sub New()

    End Sub
    '
    Public Function tools_math_MillimetersToPoints(measurementInmm As Single)
        tools_math_MillimetersToPoints = 72.0 * (measurementInmm / 25.4)
    End Function
    '
    '
    Public Function MillimetersToPoints(measurementInmm As Single)
        MillimetersToPoints = 72 * (measurementInmm / 25.4)
    End Function
    '
    '
    '
    ''' <summary>
    ''' This method will return true if myNumber is Odd and false
    ''' if it is even
    ''' </summary>
    ''' <param name="myNumber"></param>
    ''' <returns></returns>
    Public Function tools_math_isOdd(myNumber As Integer) As Boolean
        Dim ca As Integer
        Dim rslt As Boolean
        '
        rslt = True
        '
        ca = myNumber Mod 2
        '
        If ca = 0 Then
            rslt = False
        End If
        '
        Return rslt
    End Function
    '

    '
#Region "Hidden Formatting On and Off"
    '
    Public Sub tools_viewHidden_Toggle()
        'Call Me.view_OnOff(Not (ActiveWindow.view.ShowAll))        'Can do in one line
        If Me.tools_viewHidden_isOn Then
            Call Me.tools_viewHidden_OnOff(False)
        Else
            Call Me.tools_viewHidden_OnOff(True)
        End If
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will return true if hidden formatting is on
    ''' </summary>
    ''' <returns></returns>
    Public Function tools_viewHidden_isOn() As Boolean
        Dim objGlobals As New cGlobals()
        Dim docView As Word.View
        Dim rslt As Boolean
        '
        'myDoc = Globals.ThisAddIn.Application.ActiveDocument
        docView = objGlobals.glb_get_wrdActiveDoc.ActiveWindow.View
        rslt = docView.ShowParagraphs
        '
        Return rslt
    End Function

    Public Sub tools_viewHidden_OnOff(ByVal turnViewOn As Boolean)
        Dim docView As Word.View
        '
        'myDoc = Globals.ThisAddIn.Application.ActiveDocument
        docView = Globals.ThisAddin.Application.ActiveDocument.ActiveWindow.View
        '
        If turnViewOn Then
            docView.ShowAll = True
            docView.ShowTextBoundaries = False
            docView.ShowObjectAnchors = True
            docView.ShowTabs = True
            docView.TableGridlines = True
            docView.ShowBookmarks = True
            docView.ShowSpaces = True
            docView.ShowCropMarks = True
            docView.ShowParagraphs = True
            docView.ShowPicturePlaceHolders = False
            'docView.ShowFormat = True
        Else
            docView.ShowAll = False
            docView.ShowTextBoundaries = False
            docView.ShowObjectAnchors = False
            docView.ShowTabs = False
            docView.TableGridlines = False
            docView.ShowBookmarks = False
            docView.ShowSpaces = False
            docView.ShowCropMarks = False
            docView.ShowParagraphs = False
            docView.ShowPicturePlaceHolders = False
            'docView.ShowFormat = False
        End If
    End Sub
    '
#End Region
    '
    '
    Public Sub tools_paste_AsUnformattedText()
        Dim objGlobals As New cGlobals()
        ' Paste_as_unformatted_text Macro
        ' Macro recorded 07/03/03 by Dale Higgins
        ' Incorporated directly from old template .. Peter Mikelaitis Feb 2015
        '
        Try
            objGlobals.glb_get_wrdSel.PasteSpecial(Link:=False, DataType:=WdPasteDataType.wdPasteText, Placement:=WdOLEPlacement.wdInLine, DisplayAsIcon:=False)
        Catch ex As Exception
            MsgBox("Paste failed. Do you have something on your clipboard?")
        End Try
        '
        '
    End Sub
    '

    '
    Public Function getPageNumber(ByRef para As Paragraph) As Long
        'This method will retrive page number that contains the active
        'end of the range
        Dim rng As Range
        '
        rng = para.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        getPageNumber = rng.Information(WdInformation.wdActiveEndAdjustedPageNumber)
        'pageNumAbsolute = rng.Information(wdActiveEndPageNumber)
    End Function
    '
    Public Function widthBetweenMargins() As Single
        'This method will retrieve the measurement between the margins
        'in the current section.. Genherally used by methods that adjust
        'Table Widths
        Dim currentSect As Section
        currentSect = Globals.ThisAddin.Application.Selection.Sections(1)
        widthBetweenMargins = currentSect.PageSetup.PageWidth - currentSect.PageSetup.RightMargin - currentSect.PageSetup.LeftMargin
        '
    End Function
    '
    ''' <summary>
    ''' This method will remove paragrapgh training spaces
    ''' </summary>
    ''' <param name="myDoc"></param>
    Sub tools_Remove_TrailingSpacesFromParagraphs(ByRef myDoc As Word.Document)
        Dim rng As Word.Range
        '
        rng = myDoc.Content

        With rng.Find
            .ClearFormatting()
            .Text = "([ ^t^32^160]{1,})^13"
            .Replacement.Text = "^p"
            .Forward = True
            .Wrap = WdFindWrap.wdFindContinue
            .Format = False
            .MatchWildcards = True
        End With
        '
        rng.Find.Execute(Replace:=WdReplace.wdReplaceAll)
        '
    End Sub
    '
    ''' <summary>
    ''' 
    ''' </summary>
    ''' <param name="myDoc"></param>
    Sub tools_Remove_TrailingWhitespace_All(ByRef myDoc As Word.Document)
        'CoPilot suggestion..Perfect. Let’s build a fast, wildcard-powered cleanup routine that removes trailing whitespace before:
        'Paragraph Breaks(^ 13)
        'Manual Line breaks (^l)
        'Section Breaks(^ b)
        'This approach avoids slow paragraph iteration And uses Find.Execute Replace:=wdReplaceAll for speed.
        '
        Dim rng As Word.Range
        rng = myDoc.Range
        '
        ' Remove trailing spaces/tabs before paragraph marks (^13)
        With rng.Find
            .ClearFormatting()
            .Text = "([ ^t^32^160]{1,})^13"
            .Replacement.Text = "^p"
            .Forward = True
            .Wrap = WdFindWrap.wdFindContinue
            .Format = False
            .MatchWildcards = True
        End With
        rng.Find.Execute(Replace:=WdReplace.wdReplaceAll)

        ' Remove trailing spaces/tabs before manual line breaks (^l)
        With rng.Find
            .ClearFormatting()
            .Text = "([ ^t^32^160]{1,})^l"
            .Replacement.Text = "^l"
            .Forward = True
            .Wrap = WdFindWrap.wdFindContinue
            .Format = False
            .MatchWildcards = True
        End With
        rng.Find.Execute(Replace:=WdReplace.wdReplaceAll)

        ' Remove trailing spaces/tabs before section breaks (^b)
        With rng.Find
            .ClearFormatting()
            .Text = "([ ^t^32^160]{1,})^b"
            .Replacement.Text = "^b"
            .Forward = True
            .Wrap = WdFindWrap.wdFindContinue
            .Format = False
            .MatchWildcards = True
        End With
        rng.Find.Execute(Replace:=WdReplace.wdReplaceAll)
    End Sub

    '
    '
    ''' <summary>
    ''' This goes through the entire document and ensures that there is ony one space after each sentence
    ''' See https://word.tips.net/T000488_Spacing_After_Sentences.html
    ''' </summary>
    Public Sub spaces_One()
        Dim rng As Word.Range
        Dim objGlobals As New cGlobals()
        '
        rng = objGlobals.glb_get_wrdActiveDoc.Range
        'rng = Globals.ThisAddin.Application.ActiveDocument.Range
        rng.Find.ClearFormatting()
        '
        'Globals.ThisAddin.Application.ActiveDocument.Select()

        With rng.Find
            'With rng.Find

            .Text = "([.\?\!]) {1,}"
            .Replacement.Text = "\1 "
            .Forward = True
            .Wrap = WdFindWrap.wdFindContinue
            .Format = False
            .MatchWildcards = True
        End With
        rng.Find.Execute(Replace:=WdReplace.wdReplaceAll)
        '
        'Need to rest the Find object, so we change the Selection.. In this way the document
        'can be selected again
        'Globals.ThisAddin.Application.Selection.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
    End Sub
    '
    'This method will  the language to English and
    'it allows you to turn proffing on and off
    Public Sub SpellCheckProofing(turnProofingOn As Boolean)
        Dim objGlobals As New cGlobals()
        '
        objGlobals.glb_get_wrdApp.Selection.WholeStory()

        objGlobals.glb_get_wrdApp.Selection.LanguageID = WdLanguageID.wdEnglishAUS
        objGlobals.glb_get_wrdApp.Selection.NoProofing = Not turnProofingOn
        objGlobals.glb_get_wrdApp.CheckLanguage = False
        objGlobals.glb_get_wrdApp.Selection.HomeKey(Unit:=WdUnits.wdLine)
    End Sub
    '

    '
    'Ensures one space between words
    ''' <summary>
    ''' This method will ensures one space between words throughout the document
    ''' </summary>
    Public Sub spaces_OneBetweenWords()
        '
        Dim rng As Word.Range
        Dim objGlobals As New cGlobals()
        '
        'MessageBox.Show("One Space Between Words")
        'Globals.ThisAddin.Application.Selection.Find.ClearFormatting()
        'Globals.ThisAddin.Application.SelectionSelection.Find.Replacement.ClearFormatting
        'With Globals.ThisAddin.Application.Selection.Find
        '.Text = "([A-z\0-9,]) {1,}"
        '.Replacement.Text = "\1 "
        '.Forward = True
        '.Wrap = WdFindWrap.wdFindContinue
        '.Format = False
        '.MatchWildcards = True
        'End With
        'Globals.ThisAddin.Application.Selection.Find.Execute(Replace:=WdReplace.wdReplaceAll)
        '
        'MessageBox.Show("One Space Between Words")
        'Globals.ThisAddin.Application.Selection.Find.ClearFormatting()
        'Globals.ThisAddin.Application.SelectionSelection.Find.Replacement.ClearFormatting
        'Regex()
        's = Regex.Replace(s, " {2,}", " ")
        'Globals.ThisAddin.Application.ActiveDocument.Select()
        '
        rng = objGlobals.glb_get_wrdActiveDoc.Range
        'rng = Globals.ThisAddin.Application.ActiveDocument.Range
        rng.Find.ClearFormatting()

        With rng.Find()
            .Text = "([A-z\0-9,]) {1,}"
            .Replacement.Text = "\1 "
            .Forward = True
            .Wrap = WdFindWrap.wdFindContinue
            .Format = False
            .MatchWildcards = True
        End With
        rng.Find.Execute(Replace:=WdReplace.wdReplaceAll)
        '
        'Need to rest the Find object, so we change the Selection.. In this way the document
        'can be selected again
        'Globals.ThisAddin.Application.Selection.Collapse(WdCollapseDirection.wdCollapseEnd)
        '
    End Sub
    '
    ''' <summary>
    ''' This method will trim (left and right) all entries in the selected Table
    ''' </summary>
    Public Function DeleteExtraSpacesInTable() As Boolean
        ' RemoveSpacesInTable Macro
        ' Works on selection
        Dim objGlobals As New cGlobals()
        Dim tbl As Word.Table
        Dim drCell As Cell
        Dim rng As Range
        Dim myString As String
        Dim i As Integer
        Dim rslt As Boolean
        '
        rslt = True
        '
        Try
            rng = objGlobals.glb_get_wrdSelRngAll
            If rng.Tables.Count <> 0 Then
                tbl = rng.Tables.Item(1)
                rng = tbl.Range
                For i = 1 To rng.Cells.Count
                    drCell = rng.Cells.Item(i)
                    myString = Trim(drCell.Range.Text)
                    myString = Left$(myString, Len(myString) - 1)
                    '
                    If Right$(myString, 1) = vbCrLf Or Right$(myString, 1) = vbNewLine Or Right$(myString, 1) = vbCr Then
                        myString = Left$(myString, Len(myString) - 1)
                    End If
                    '
                    myString = Trim(myString)
                    '
                    drCell.Range.Delete()
                    drCell.Range.Text = myString
                Next
            Else
                rslt = False
            End If

        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
        '
    End Function
    '
    Public Sub RemoveAllStyleAliases(ByRef myDoc As Word.Document)
        Dim sty As Style
        For Each sty In myDoc.Styles
            sty.NameLocal = Split(sty.NameLocal, ",")(0)
        Next sty
    End Sub
    '
    '
    ''' <summary>
    ''' This method will return the style name of the style in the first cell of
    ''' the Table tbl
    ''' </summary>
    ''' <param name="tbl"></param>
    ''' <returns></returns>
    Public Function tools_tbls_getFirstCellStyleName(ByRef tbl As Word.Table) As String
        Dim drCell As Word.Cell
        Dim rng As Word.Range
        Dim myStyle As Word.Style
        Dim strStyleName As String
        '
        Try
            drCell = tbl.Range.Cells.Item(1)
            rng = drCell.Range
            rng.Collapse(WdCollapseDirection.wdCollapseStart)
            '
            myStyle = rng.Style
            strStyleName = myStyle.NameLocal
            '
        Catch ex As Exception
            strStyleName = "error"
        End Try
        '
        Return strStyleName
        '
    End Function
    '
    ''' <summary>
    ''' This method will retrieve the text contents of a celll. It will strip out the special 
    ''' End of cell character(s)... A cell with no text contains 0D,07
    ''' </summary>
    ''' <param name="drCell"></param>
    ''' <param name="stripEndOfCell"></param>
    Public Function tools_cell_GetText(ByRef drCell As Word.Cell, ByVal stripEndOfCell As Boolean) As String
        '
        Dim strEnd As String
        '
        tools_cell_GetText = Trim(drCell.Range.Text)
        If stripEndOfCell Then
            tools_cell_GetText = Mid(drCell.Range.Text, 1, (Len(drCell.Range.Text) - 2))
            'No strip away any and all carriage returns at the send of the string
            Do While True
                strEnd = Right(tools_cell_GetText, 1)
                If strEnd = vbCr Then
                    tools_cell_GetText = Left(tools_cell_GetText, (Len(tools_cell_GetText) - 1))
                Else
                    Exit Do
                End If
            Loop
        End If
    End Function


End Class

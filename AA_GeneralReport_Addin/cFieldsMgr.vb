Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cFieldsMgr
    Public Sub New()

    End Sub
    '
    '
    Public Sub upDateTableOfFigures()
        Dim j As Integer
        Dim myDoc As Word.Document
        Dim objGlobals As New cGlobals()
        '
        myDoc = objGlobals.glb_get_wrdActiveDoc()
        '
        For j = 1 To myDoc.TablesOfFigures.Count
            myDoc.TablesOfFigures(j).Update()
        Next j

    End Sub
    '
#Region "Cross Reference Fields"
    '
    ''' <summary>
    ''' This method generate a list of errored cross references. It will do a field
    ''' update before testing for errors if doFldUpdate is true
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="showErrorCrossRefsOnly"></param>
    ''' <param name="doFldUpdate"></param>
    ''' <returns></returns>
    Public Function CrossReference_List(ByRef myDoc As Word.Document, showErrorCrossRefsOnly As Boolean, doFldUpdate As Boolean) As List(Of Word.Field)
        Dim flds As Word.Fields
        Dim errorFlds As New List(Of Word.Field)
        Dim fld As Word.Field
        Dim rng As Word.Range
        Dim lstOfErrorPages As New Collection()
        '
        'targetFldCode = StrConv(targetFldCode, vbLowerCase)
        'targetFldCode = targetFldCode & "*"
        '
        'myDoc = Globals.ThisDocument.Application.ActiveDocument
        'showErrorCrossRefsOnly = False
        '
        Try
            flds = myDoc.Range.Fields
            For Each fld In flds
                If showErrorCrossRefsOnly Then
                    If fld.Type = WdFieldType.wdFieldRef Then
                        '
                        If doFldUpdate Then fld.Update()
                        '
                        fld.Select()
                        If fld.Result.Text Like "Error*" Then
                            rng = fld.Result
                            errorFlds.Add(fld)
                        End If

                    End If
                Else
                    If fld.Type = WdFieldType.wdFieldRef Then
                        'fld.Result.
                        fld.Select()
                        rng = fld.Result
                        errorFlds.Add(fld)
                    End If
                End If
            Next fld
        Catch ex As Exception
            MsgBox("Error in Cross Reference Change Style")
        End Try
        '
        Return errorFlds
    End Function

    Public Sub CrossReference_changeStyle()
        Dim flds As Word.Fields
        Dim fld As Word.Field
        Dim strFldCode As String
        Dim myDoc As Word.Document
        Dim objGlobals As New cGlobals()
        '
        'targetFldCode = StrConv(targetFldCode, vbLowerCase)
        'targetFldCode = targetFldCode & "*"
        '
        myDoc = objGlobals.glb_get_wrdActiveDoc
        Try
            flds = myDoc.Range.Fields
            For Each fld In flds
                If fld.Type = WdFieldType.wdFieldRef Then
                    'fld.Result.
                    fld.Select()

                    'sel = Globals.ThisDocument.Application.Selection
                    'bkMarks = Globals.ThisDocument.Application.Selection.Bookmarks
                    'bkMark = bkMarks.Item(1)
                    objGlobals.glb_get_wrdSel.ClearFormatting()
                    'bkMark.Range.Style = "Body Text"
                    objGlobals.glb_get_wrdSel.Style = "Cross Reference"
                    strFldCode = Trim(fld.Code.Text)
                    'fld.Code = 
                    'strFldCode = StrConv(strFldCode, vbLowerCase)
                    'If strFldCode Like targetFldCode Then
                    'fld.Update()
                    'End If
                End If
            Next fld
        Catch ex As Exception
            MsgBox("Error in Cross Reference Change Style")
        End Try

    End Sub

#End Region
    '
#Region "Update Fields"
    '
    Public Sub updateCrossReferenceFields()
        'Dim flds As Word.Fields
        Dim objGlobals As New cGlobals()
        Dim myDoc As Word.Document
        '
        'targetFldCode = StrConv(targetFldCode, vbLowerCase)
        'targetFldCode = targetFldCode & "*"
        '

        myDoc = objGlobals.glb_get_wrdActiveDoc
        Me.updateCrossReferenceFields(myDoc)
        'flds = myDoc.Range.Fields
        'For Each fld In flds
        'If fld.Type = WdFieldType.wdFieldRef Then
        'fld.Update()
        'End If
        'Next fld


    End Sub
    '
    '
    Public Sub updateCrossReferenceFields(ByRef myDoc As Word.Document)
        Dim flds As Word.Fields
        Dim fld As Word.Field
        '
        'targetFldCode = StrConv(targetFldCode, vbLowerCase)
        'targetFldCode = targetFldCode & "*"
        '
        'myDoc = Globals.ThisDocument.Application.ActiveDocument
        flds = myDoc.Range.Fields
        For Each fld In flds
            If fld.Type = WdFieldType.wdFieldRef Then
                fld.Update()
            End If
        Next fld


    End Sub

    '
    Public Sub updateSequenceNumbers_Appendix()
        Call Me.updateSequenceNumbers("SEQ AppNum")
    End Sub
    '
    Public Sub updateSequenceNumbers_Appendix(ByRef rng As Word.Range)
        Call Me.updateSequenceNumbers("SEQ AppNum")
    End Sub
    '
    Public Sub updateSequenceNumbers_Boxes()
        Call Me.updateSequenceNumbers("SEQ Box")
        Call Me.updateStyleRefs("STYLEREF 1")
    End Sub
    '
    Public Sub updateSequenceNumbers_Boxes_Ap()
        Call Me.updateSequenceNumbers("SEQ Box_AP")
        Call Me.updateStyleRefs("STYLEREF 9")
    End Sub
    '
    Public Sub updateSequenceNumbers_Boxes_ES()
        Call Me.updateSequenceNumbers("SEQ Box_ES")
    End Sub
    '
    Public Sub updateSequenceNumbers_Boxes_LT()
        Call Me.updateSequenceNumbers("SEQ Box_LT")
    End Sub
    '
    Public Sub updateSequenceNumbers_Boxes_KeyFindings()
        Call Me.updateSequenceNumbers("SEQ Key_Finding")
    End Sub
    '
    Public Sub updateSequenceNumbers_Boxes_KeyFindings_ES()
        Call Me.updateSequenceNumbers("SEQ KeyFinding_ES")
    End Sub
    '
    Public Sub updateSequenceNumbers_Boxes_Recommendation()
        Call Me.updateSequenceNumbers("SEQ Recommendation")
    End Sub
    '
    Public Sub updateSequenceNumbers_Boxes_Recommendation_ES()
        Call Me.updateSequenceNumbers("SEQ Recommendation_ES")
    End Sub
    '

    '
    Public Sub updateSequenceNumbers_Chapters()
        Call Me.updateSequenceNumbers("SEQ ChptNum")
    End Sub
    '
    Public Sub updateSequenceNumbers_Figures()
        Call Me.updateSequenceNumbers("SEQ Figure")
        Call Me.updateStyleRefs("STYLEREF 1")
    End Sub
    '
    Public Sub updateSequenceNumbers_Figures_WorkAround()
        Dim objTools As New cTools()
        '
        'Insert a Figure Caption to force the document into a Figure update mode that
        'does not seem to be accessible any other way. Then delete the inserted caption
        '
        'objTools.Insert_Caption_Figure()
        '
        'Now Update the sequence number of all Figure Types. theinserted caption would
        'have upset the main body sequence numbers
        Me.updateSequenceNumbers("SEQ Figure")
        Call Me.updateStyleRefs("STYLEREF 1")
        '
    End Sub
    '
    Public Sub updateSequenceNumbers_Figures_Ap()
        Call Me.updateSequenceNumbers("SEQ Figure_AP")
        Call Me.updateStyleRefs("STYLEREF 9")
    End Sub
    '
    Public Sub updateSequenceNumbers_Figures_ES()
        Call Me.updateSequenceNumbers("SEQ Figure_ES")
    End Sub
    '
    Public Sub updateSequenceNumbers_Figures_LT()
        Call Me.updateSequenceNumbers("SEQ Figure_LT")
    End Sub
    '
    Public Sub updateSequenceNumbers_Parts()
        Call Me.updateSequenceNumbers("SEQ NumList")
    End Sub
    '
    Public Sub updateSequenceNumbers_Tables()
        Call Me.updateSequenceNumbers("SEQ Table")
        Call Me.updateStyleRefs("STYLEREF 1")
    End Sub
    '
    Public Sub updateSequenceNumbers_Tables_AP()
        Call Me.updateSequenceNumbers("SEQ Table")
        Call Me.updateStyleRefs("STYLEREF 9")
    End Sub
    '
    Public Sub updateSequenceNumbers_Tables_ES()
        Call Me.updateSequenceNumbers("SEQ Table_ES")
    End Sub
    '
    Public Sub updateSequenceNumbers_Tables_LT()
        Call Me.updateSequenceNumbers("SEQ Table_LT")
    End Sub
    '
#End Region
    '
#Region "Update Seq and styleRefs base routines"

    '
    ''' <summary>
    ''' This method will update the sequence numbers identified by 'targetFldCode'
    ''' </summary>
    ''' <param name="targetFldCode"></param>
    Public Sub updateSequenceNumbers(targetFldCode As String)
        Dim flds As Word.Fields
        Dim fld As Word.Field
        Dim strFldCode As String
        Dim objGlobals As New cGlobals()
        Dim myDoc As Word.Document
        '
        myDoc = objGlobals.glb_get_wrdActiveDoc()
        '
        'targetFldCode = StrConv(targetFldCode, vbLowerCase)
        targetFldCode = targetFldCode & "*"
        '
        flds = myDoc.Range.Fields

        For Each fld In flds
            If fld.Type = WdFieldType.wdFieldSequence Then
                strFldCode = Trim(fld.Code.Text)
                'strFldCode = StrConv(strFldCode, vbLowerCase)
                If strFldCode Like targetFldCode Then
                    fld.Update()
                End If
            End If
        Next fld
    End Sub
    '
    ''' <summary>
    ''' This method will get the fields in the range (rng) and will cause the specific
    ''' field identified by targetFldCode to update
    ''' </summary>
    ''' <param name="targetFldCode"></param>
    ''' <param name="rng"></param>
    Public Sub updateSequenceNumbers(targetFldCode As String, ByRef rng As Word.Range)
        Dim flds As Word.Fields
        Dim fld As Word.Field
        Dim strFldCode As String
        '
        flds = rng.Fields
        '
        For Each fld In flds
            If fld.Type = WdFieldType.wdFieldSequence Then
                strFldCode = Trim(fld.Code.Text)
                'strFldCode = StrConv(strFldCode, vbLowerCase)
                If strFldCode Like targetFldCode Then
                    fld.Update()
                End If
            End If
        Next fld
        '
    End Sub

    '
    ''' <summary>
    ''' This method will update all sequence fields in the Active Document
    ''' </summary>
    Public Sub updateSequenceNumbers_All()
        Dim flds As Word.Fields
        Dim fld As Word.Field
        Dim objGlobals As New cGlobals()
        Dim myDoc As Word.Document
        '
        myDoc = objGlobals.glb_get_wrdActiveDoc()
        '
        flds = myDoc.Range.Fields
        For Each fld In flds
            If fld.Type = WdFieldType.wdFieldSequence Then fld.Update()
        Next fld
    End Sub
    '
    '
    ''' <summary>
    ''' This method will update the style refs identified in 'targetFldCode'
    ''' </summary>
    ''' <param name="targetFldCode"></param>
    Public Sub updateStyleRefs(targetFldCode As String)
        Dim flds As Word.Fields
        Dim fld As Word.Field
        Dim strFldCode As String
        Dim objGlobals As New cGlobals()
        Dim myDoc As Word.Document
        '
        myDoc = objGlobals.glb_get_wrdActiveDoc()
        '
        'targetFldCode = StrConv(targetFldCode, vbLowerCase)
        targetFldCode = "*" + targetFldCode + "*"
        '
        flds = myDoc.Range.Fields
        For Each fld In flds
            If fld.Type = WdFieldType.wdFieldStyleRef Then
                strFldCode = Trim(fld.Code.Text)
                'strFldCode = StrConv(strFldCode, vbLowerCase)
                If strFldCode Like targetFldCode Then
                    fld.Update()
                End If
            End If
        Next fld
    End Sub
    '
    ''' <summary>
    ''' This method will update all style refs in the document
    ''' </summary>
    Public Sub updateStyleRefs_All()
        Dim flds As Word.Fields
        Dim fld As Word.Field
        Dim objGlobals As New cGlobals()
        Dim myDoc As Word.Document
        '
        'targetFldCode = StrConv(targetFldCode, vbLowerCase)
        myDoc = objGlobals.glb_get_wrdActiveDoc()
        '
        flds = myDoc.Range.Fields
        For Each fld In flds
            If fld.Type = WdFieldType.wdFieldStyleRef Then fld.Update()
        Next fld
    End Sub
    '
    '
    ''' <summary>
    ''' This method will update all StyleRef fields in the current document's Header/Footers
    ''' </summary>
    Public Sub flds_update_StyleRefs_Hfs()
        Dim fld As Word.Field
        Dim myDoc As Word.Document
        Dim hf As Word.HeaderFooter
        Dim objGlobals As New cGlobals()
        '

        myDoc = objGlobals.glb_get_wrdActiveDoc()
        '
        For Each sect In myDoc.Sections
            Try
                For Each hf In sect.Headers
                    If hf.Exists Then
                        For Each fld In hf.Range.Fields
                            If fld.Type = WdFieldType.wdFieldStyleRef Then fld.Update()
                        Next
                    End If
                Next
            Catch ex As Exception

            End Try
            '
            Try
                For Each hf In sect.Footers
                    If hf.Exists Then
                        For Each fld In hf.Range.Fields
                            If fld.Type = WdFieldType.wdFieldStyleRef Then fld.Update()
                        Next
                    End If
                Next
            Catch ex As Exception

            End Try
        Next
    End Sub
    '
    '
    ''' <summary>
    ''' This method will update all StyleRef fields in the current document's Footers
    ''' </summary>
    Public Sub flds_update_StyleRefs_Footers()
        Dim objGlobals As New cGlobals()
        '
        objGlobals.glb_flds_updateStyleRefsFooters()
        '
    End Sub

    '
    ''' <summary>
    ''' This method will cause the various fields in the document to update
    ''' They'll normally update on Ptint, but sometimes you want to see the
    ''' changes immediately.. So we use this. Note that it will bring you back
    ''' to where your cursor started
    ''' </summary>
    Public Sub flds_updateFields_All()
        'may not do fields in header/footer
        Dim rng As Range
        Dim objGlobals As New cGlobals()
        '
        'Get the current range so that we can re-establish it at the end
        rng = objGlobals.glb_get_wrdSelRngAll()
        '
        objGlobals.glb_get_wrdApp()
        objGlobals.glb_get_wrdApp().ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView
        objGlobals.glb_get_wrdApp().ScreenUpdating = True
        objGlobals.glb_get_wrdApp().ActiveDocument.Fields.Update()
        '
        rng.Select()
        '
    End Sub


#End Region
    '
    Public Sub upDateCommentsField()
        Dim objGlobals As New cGlobals()
        Dim myDoc As Word.Document
        '
        myDoc = objGlobals.glb_get_wrdActiveDoc
        '
        Dim f As Field
        For Each f In myDoc.Fields
            If f.Type = WdFieldType.wdFieldComments Then
                f.Update()
            End If
        Next f
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will strip the fields from the TOC, leaving only the text. Useful when you want
    ''' a TOC that is just text... For example, we use this for 'Accessible' documents. A live TOC
    ''' with Fields is deemed by the Accessibility checked to have 'contrast' problems
    ''' </summary>
    Public Sub flds_tocs_unlink(ByRef myDoc As Word.Document, Optional useHyperLinks As Boolean = False)
        Dim objGlobals As New cGlobals()
        Dim fld As Word.Field
        '
        For Each toc In myDoc.TablesOfContents
            toc.UseHyperlinks = useHyperLinks
        Next
        '
        'This will also get the Table of Tables etc
        For Each fld In myDoc.Fields
            If fld.Type = WdFieldType.wdFieldTOC Then
                fld.Unlink()

            End If
        Next

    End Sub
    '
    ''' <summary>
    ''' This method will unlink all fields in the body of the document myDoc. It should leave
    ''' the headers and Footer alone
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub flds_body_unLink(ByRef myDoc As Word.Document)
        Dim sect As Word.Section
        Dim fld As Word.Field
        '
        For Each sect In myDoc.Sections
            For Each fld In sect.Range.Fields
                fld.Unlink()
            Next
        Next
    End Sub
    '
    Public Sub flds_footer_unlink(ByRef myDoc As Word.Document)
        Dim objGlobals As New cGlobals()
        'Dim myDoc As Word.Document
        Dim sect As Word.Section
        Dim hf As HeaderFooter
        Dim fld As Word.Field

        'myDoc = objGlobals.glb_get_wrdActiveDoc()
        'objGlobals.glb_get_wrdActiveDoc.in

        For Each sect In myDoc.Sections
            For Each hf In sect.Footers
                If hf.Exists Then
                    For Each fld In hf.Range.Fields
                        'fld.
                        'fld.Update()
                        If fld.Type = WdFieldType.wdFieldStyleRef Or fld.Type = WdFieldType.wdFieldPage Then
                            Select Case fld.Type
                                Case WdFieldType.wdFieldStyleRef
                                    fld.Unlink()
                                Case WdFieldType.wdFieldPage
                                    'fld.Update()
                                    'rng = fld.Result

                                    'strText = fld.Result.Text
                                    'rng = hf.Range
                                    'rng.Information(WdInformation.wdActiveEndAdjustedPageNumber)
                                    'rng.Collapse(WdCollapseDirection.wdCollapseEnd)
                                    'rng.Move(WdUnits.wdCharacter, -2)
                                    'fld.Delete()
                                    'rng.Text = "test " + strText + " "
                                    'rng.Text = strText
                                    'fld.Delete()
                            End Select
                        End If
                    Next
                End If
            Next
        Next

        'For Each fld In myDoc.Fields
        'If fld.Type = WdFieldType.wdFieldStyleRef Then
        'fld.Unlink()

        'End If
        'Next

    End Sub
    '
    ''' <summary>
    ''' This method will unlink all fields.. But it does so in a sequenced manner. First the footers, then the toc and
    ''' finally all fields
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub flds_all_unlink(ByRef myDoc As Word.Document)
        Dim objGlobals As New cGlobals()

        flds_footer_unlink(myDoc)
        flds_tocs_unlink(myDoc)

        myDoc.Fields.Unlink()

    End Sub
    '
End Class

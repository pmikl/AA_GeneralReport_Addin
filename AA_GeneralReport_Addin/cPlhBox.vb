Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cPlhBox
    Inherits cPlHBase
    Public Sub New()
        MyBase.New()
        '
        'Set up everything that constitues the PlaceHolder Category of 'BOX'... These names
        'are subsets of the Seqnece Fields used in the Captions associated with each Box type
        '
        Me.lstOfPlhTypes.Clear()
        Me.lstOfPlhTypes.Add("Box")
        Me.lstOfPlhTypes.Add("Key_")
        Me.lstOfPlhTypes.Add("Recommendation")
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will insert a Box at the current selection point. The type of Box is dependent
    ''' on the value of strType. Note that if the current selection point contains a table, or is
    ''' right under a table
    ''' </summary>
    ''' <param name="strType"></param>
    ''' <returns></returns>
    Public Function PlhBox_insert_Box(strType As String, Optional doTableCheck As Boolean = True) As Word.Table
        Dim tbl As Word.Table
        Dim objGlobals As New cGlobals()
        Dim objMsgMgr As New cMessageManager()
        Dim numTextColumns As Integer
        Dim marginWidth As Single
        Dim sect As Word.Section
        Dim rng As Word.Range
        '
        '
        tbl = Nothing
        rng = objGlobals.glb_get_wrdSelRng()
        sect = objGlobals.glb_get_wrdSect()
        marginWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
        '
        numTextColumns = sect.PageSetup.TextColumns.Count
        '
        If doTableCheck Then
            If rng.Tables.Count = 0 Then
                tbl = Me.Plh_insert_PlaceHolder_WithTest(objGlobals.glb_get_wrdSelRng, strType)
                '
                'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
                'tbl.PreferredWidth = 100

            Else
                If doTableCheck Then
                    MsgBox(objMsgMgr.msgMgr_msg_tooNearATable())
                End If
            End If
        Else
            'Just do it with no concern reating to any tables etc
            tbl = Me.Plh_insert_PlaceHolder_WithTest(objGlobals.glb_get_wrdSelRng, strType)
            'tbl.PreferredWidthType = WdPreferredWidthType.wdPreferredWidthPercent
            'tbl.PreferredWidth = 100

        End If

        '
        '
        '
        Return tbl
        '
    End Function
    '
    Public Sub PlhBox_Captions_ConvertBoxCaptionsTo_ES(ByRef rngSrc As Word.Range)
        '
        Me.Plh_Captions_ConvertCaptions(Me.lstOfPlhTypes, "Box_ES", rngSrc)
        '
    End Sub
    '
    Public Sub PlhBox_Captions_ConvertBoxCaptionsTo_Report(ByRef rngSrc As Word.Range)
        '
        Me.Plh_Captions_ConvertCaptions(Me.lstOfPlhTypes, "Box", rngSrc)
        '
    End Sub
    '
    Public Sub PlhBox_Captions_ConvertBoxCaptionsTo_Appendix(ByRef rngSrc As Word.Range)
        '
        Me.Plh_Captions_ConvertCaptions(Me.lstOfPlhTypes, "Box_AP", rngSrc)
        '
    End Sub

    '
    Public Sub PlhBox_Captions_ConvertBoxCaptionsTo_Letter(ByRef rngSrc As Word.Range)
        '
        Me.Plh_Captions_ConvertCaptions(Me.lstOfPlhTypes, "Box_LT", rngSrc)
        '
    End Sub
    '
    '
    Public Sub PlhBox_Exmaples_InsertBoxText(doAsText As Boolean, Optional ByVal strNewContent As String = "Overtype here")
        'This method will delete the contents of the specified cell
        Dim rng As Range
        Dim tbl As Table
        Dim drCell As Word.Cell
        Dim strResult As String
        Dim objStylesMgr As New cStylesManager()
        Dim objBBMgr As New cBBlocksHandler()
        Dim objGlobals As New cGlobals()
        '
        Try
            strResult = ""
            rng = objGlobals.glb_get_wrdSelRngAll
            '
            If rng.Tables.Count = 0 Then
                MsgBox("Please make certain that you have placed your cursor in the text area of a Box")
                Exit Sub
            End If
            '
            If Not (rng.Style.NameLocal Like "Box*") Then
                MsgBox("Please make certain that you have placed your cursor in the text area of the Box. It's probably in one of the spacing rows at the top or bottom of the Box ")
                Exit Sub
            End If
            '
            If rng.Tables.Count <> 0 Then
                tbl = rng.Tables.Item(1)
                drCell = tbl.Range.Cells.Item(2)
                rng = drCell.Range
                rng.Delete()
                rng.Style = objGlobals.glb_get_wrdActiveDoc.Styles("Box text")
                '
                If doAsText Then
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    rng.Text = strNewContent
                    'Globals.ThisDocument.Application.Selection.TypeText(strNewContent)
                    'rng.Find.Text = strNewContent
                    'rng.Find.Execute()
                    rng.Select()
                Else
                    '
                    objStylesMgr.insertStyleSetReport_Box()
                    '
                    'rng.Style = Globals.ThisDocument.Application.ActiveDocument.Styles("Box Quote Source")
                    'Call objBBMgr.insertBuildingBlockFromDefaultLibToRange("aac_Styles_BoxStyles", "styleSets", rng)
                    rng = drCell.Range
                    rng.Collapse(WdCollapseDirection.wdCollapseStart)
                    rng.Select()
                End If
                '
            End If
        Catch ex As Exception
            MsgBox("This function is only supported in Boxes (Standard, Key Findings and Recommendations)")
        End Try
        '
    End Sub

End Class

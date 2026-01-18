Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cPlHTable
    Inherits cPlHBase
    '
    Public tbl_OutDent As Single

    Public Sub New()
        MyBase.New()
        '
        'Set up everything that constitues the PlaceHolder Category of 'TABLE'... These names
        'are subsets of the Sequence Fields used in the Captions associated with each Box type
        '
        Me.lstOfPlhTypes.Clear()
        Me.lstOfPlhTypes.Add("Table")
        '
        Me.tbl_OutDent = 8.0                     'OutDent in mm
        Me.tbl_OutDent = 0.0                     '2024 update OutDent in mm 

        '
    End Sub
    '
    '
    '
    ''' <summary>
    ''' This method will insert a Table at the current selection point. The type of Box is dependent
    ''' on the value of strType ("Table_ES", "Table", "Table_AP", "Table_LT"
    ''' </summary>
    ''' <param name="strType"></param>
    ''' <returns></returns>
    Public Function PlhTbl_insert_Table(strType As String) As Word.Table
        Dim tbl As Word.Table
        Dim objGlobals As New cGlobals()
        Dim sect As Word.Section
        '
        tbl = Nothing
        sect = objGlobals.glb_get_wrdSect()
        'marginWidth = sect.PageSetup.PageWidth - sect.PageSetup.LeftMargin - sect.PageSetup.RightMargin
        '
        'numTextColumns = sect.PageSetup.TextColumns.Count
        '
        tbl = Me.Plh_insert_PlaceHolder_WithTest(objGlobals.glb_get_wrdSelRng, strType)
        '
        Return tbl
        '
    End Function
    '
    '
    Public Function PlhTbl_insert_TableWide(strType As String) As Word.Table
        Dim tbl As Word.Table
        '
        tbl = MyBase.Plh_insert_PlaceHolder_Wide(strType)
        objTblsMgr.tbl_captions_doIndent(tbl)
        '
        Return tbl
    End Function
    '

    '
    '
#Region "Conversions"
    Public Sub PlhTbl_Captions_ConvertCaptionsTo_ES(ByRef rngSrc As Word.Range)
        Me.Plh_Captions_ConvertCaptions(Me.lstOfPlhTypes, "Table_ES", rngSrc)
    End Sub
    '
    Public Sub PlhTbl_Captions_ConvertCaptionsTo_Report(ByRef rngSrc As Word.Range)
        Me.Plh_Captions_ConvertCaptions(Me.lstOfPlhTypes, "Table", rngSrc)
    End Sub
    '
    Public Sub PlhTbl_Captions_ConvertCaptionsTo_Appendix(ByRef rngSrc As Word.Range)
        Me.Plh_Captions_ConvertCaptions(Me.lstOfPlhTypes, "Table_AP", rngSrc)
    End Sub
    '
    Public Sub PlhTbl_Captions_ConvertCaptionsTo_Letter(ByRef rngSrc As Word.Range)
        Me.Plh_Captions_ConvertCaptions(Me.lstOfPlhTypes, "Table_LT", rngSrc)
    End Sub

#End Region
    '


End Class

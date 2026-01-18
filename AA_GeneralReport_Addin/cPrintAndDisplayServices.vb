Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cPrintAndDisplayServices
    Public Sub New()

    End Sub
    '
    '
    Public Function colour_display_ToPrintMode(ByRef myDoc As Word.Document) As Boolean
        Dim isOK As Boolean
        '
        isOK = False
        Try
            If Me.isAcilAllenReport(myDoc) Then
                myDoc.Styles.Item("Normal_rgb").Font.Color = RGB(157, 133, 190)
                Me.colour_Controls(myDoc)
            End If
            isOK = True
        Catch ex As Exception
            isOK = False
        End Try
        '
        Return isOK
    End Function
    '
    Public Function colour_display_ToDefault(ByRef myDoc As Word.Document) As Boolean
        Dim isOK As Boolean
        '
        isOK = False
        Try
            If Me.isAcilAllenReport(myDoc) Then
                myDoc.Styles.Item("Normal_rgb").Font.Color = RGB(108, 63, 153)
                Me.colour_Controls(myDoc)
            End If
            isOK = True
            '
        Catch ex As Exception
            isOK = False
        End Try
        '
        Return isOK
    End Function
    Public Sub print_with_ColourChange(myDoc As Word.Document)
        '
        Me.colour_display_ToDesignView(myDoc)
        '
        myDoc.PrintOut()
        '
        Me.colour_display_ToEasyView(myDoc)

    End Sub
    '
    Public Sub save_with_ColourChange(ByRef myDoc As Word.Document)
        '
        Me.colour_display_ToDesignView(myDoc)
        '
        myDoc.Save()
        myDoc.Saved = True
        '
        Me.colour_display_ToEasyView(myDoc)
        '
    End Sub
    '
    Public Sub open_Colour_ToDefault()

    End Sub
    '
    Public Function colour_display_ToDesignView(ByRef myDoc As Word.Document) As Boolean
        Dim isOK As Boolean
        '
        isOK = False
        Try
            If Me.isAcilAllenReport(myDoc) Then
                myDoc.Styles.Item("Normal_rgb").Font.Color = RGB(157, 133, 190)
                Me.colour_Controls(myDoc)
            End If
            isOK = True
        Catch ex As Exception
            isOK = False
        End Try
        '
        Return isOK
    End Function
    '
    Public Function colour_display_ToEasyView(ByRef myDoc As Word.Document) As Boolean
        Dim isOK As Boolean
        '
        isOK = False
        Try
            'If Me.isAcilAllenReport(myDoc) Then
            myDoc.Styles.Item("Normal_rgb").Font.Color = RGB(108, 63, 153)
                Me.colour_Controls(myDoc)
            'End If
            isOK = True
            '
        Catch ex As Exception
            isOK = False
        End Try
        '
        Return isOK
    End Function
    '
    Public Sub colour_Controls(ByRef myDoc As Word.Document)
        Dim styl As Word.Style
        Dim rgbColour As Long
        Dim ctrl As Word.ContentControl
        Dim j As Integer
        '
        styl = myDoc.Styles.Item("Normal_rgb")
        rgbColour = styl.Font.Color
        '
        Try
            For j = 1 To myDoc.ContentControls.Count
                ctrl = myDoc.ContentControls.Item(j)
                ctrl.Range.Font.Color = rgbColour
            Next
        Catch ex As Exception

        End Try
    End Sub
    '
    ''' <summary>
    ''' This method checks for the style "tag_aa_RptTestStyle_#?_00" to determine
    ''' if the document myDoc is an Acil Allen Report
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function isAcilAllenReport(ByRef myDoc As Word.Document) As Boolean
        Dim rslt As Boolean
        Dim styl As Word.Style
        '
        rslt = False
        Try
            styl = myDoc.Styles.Item("tag_aa_RptTestStyle_#?_00")
            If Not IsNothing(styl) Then rslt = True
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
        '
    End Function


End Class

Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cLogosMgr
    Public Sub New()

    End Sub
    '
    '
    ''' <summary>
    ''' This method will search the header of sect (both primary and first page) and
    ''' if it finds a shape and if it is the logo it will recolour the name
    ''' as per rgbName and the underscore bar as per rgbBar... If rgbBar is negative it will
    ''' set the colour to the standard purple
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="rgbName"></param>
    ''' <param name="rgbBar"></param>
    Public Sub logos_set_colour(ByRef sect As Word.Section, rgbName As Long, rgbBar As Long)
        Dim hf As Word.HeaderFooter
        '
        If rgbBar < 0 Then rgbBar = RGB(108, 63, 153)
        '
        Try
            hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
            Me.logos_set_colour(hf, rgbName, rgbBar)
        Catch ex As Exception

        End Try
        '
        Try
            hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
            Me.logos_set_colour(hf, rgbName, rgbBar)
        Catch ex As Exception

        End Try
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will colour the name and the bar part of the aac logo (shp) according to the
    ''' rgb value of rgbName and rgbBar respectiviely.. By default it will set the decorative paramater
    ''' to true
    ''' </summary>
    ''' <param name="shp"></param>
    ''' <param name="rgbName"></param>
    ''' <param name="rgbBar"></param>
    Public Sub logos_set_colour(ByRef shp As Word.Shape, rgbName As Long, rgbBar As Long, Optional setDecorative As Boolean = True)
        Dim i As Integer
        Dim grpShp As Word.Shape
        Dim objWCAGMgr As New cWCAGMgr()
        '
        For i = 1 To shp.GroupItems.Count
            grpShp = shp.GroupItems.Item(i)
            If grpShp.Name Like "Freeform*" Then grpShp.Fill.ForeColor.RGB = rgbName
            If grpShp.Name Like "Rectangle*" Then grpShp.Fill.ForeColor.RGB = rgbBar
        Next
        '
        objWCAGMgr.wcag_set_decorative(shp, setDecorative)

    End Sub
    '
    ''' <summary>
    ''' This method will search the header of sect (both primary and first page) and
    ''' if it finds a shape and if it is the logo it will recolour the name
    ''' as per rgbName and the underscore bar as per rgbBar... If rgbBar is negative it will
    ''' set the colour to the standard purple
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="rgbName"></param>
    ''' <param name="rgbBar"></param>
    Public Sub logos_set_colour(ByRef hf As Word.HeaderFooter, rgbName As Long, rgbBar As Long, Optional setDecorative As Boolean = True)
        Dim rng As Word.Range
        Dim shp As Word.Shape
        '
        If rgbBar < 0 Then rgbBar = RGB(108, 63, 153)
        '
        rng = hf.Range
        If rng.ShapeRange.Count <> 0 Then
            For Each shp In rng.ShapeRange
                Select Case shp.Name
                    Case "logo_AAC_TandG"
                        Me.logos_set_colour(shp, rgbName, rgbBar, setDecorative)
                        'For i = 1 To shp.GroupItems.Count
                        'grpShp = shp.GroupItems.Item(i)
                        'If grpShp.Name = "Freeform 6" Then grpShp.Fill.ForeColor.RGB = rgbName
                        'If grpShp.Name = "Rectangle 42" Then grpShp.Fill.ForeColor.RGB = rgbBar
                        'Next
                    Case "aac_Cpg_Logo"
                        Me.logos_set_colour(shp, rgbName, rgbBar, setDecorative)
                        'For i = 1 To shp.GroupItems.Count
                        'grpShp = shp.GroupItems.Item(i)
                        'If grpShp.Name = "Freeform 6" Then grpShp.Fill.ForeColor.RGB = rgbName
                        'If grpShp.Name = "Rectangle 33" Then grpShp.Fill.ForeColor.RGB = rgbBar
                        'Next
                End Select

            Next
        End If

    End Sub
    '

End Class

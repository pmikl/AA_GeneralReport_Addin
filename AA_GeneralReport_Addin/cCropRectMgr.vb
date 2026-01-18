
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word
Imports Microsoft.VisualBasic.FileIO
Imports System.Drawing
'
'rev 01.00  20250830
'
Public Class cCropRectMgr
    Public Sub New()

    End Sub
    '
    ''' <summary>
    ''' AspectRatio is defined as Height/Width
    ''' </summary>
    ''' <param name="AspectRatio"></param>
    ''' <returns></returns>
    Public Function rct_get_height(AspectRatio As Single, width As Single) As Single
        Dim height As Single
        '
        height = width * AspectRatio
        '
        Return height
    End Function
    '
    ''' <summary>
    ''' AspectRatio is defined as Height/Width
    ''' </summary>
    ''' <param name="AspectRatio"></param>
    ''' <returns></returns>
    Public Function rct_get_width(AspectRatio As Single, height As Single) As Single
        Dim width As Single
        '
        width = height / AspectRatio
        '
        Return width
    End Function
    '
    ''' <summary>
    ''' This method will take the image to eb clipped as a shape and return an string indicating it s shape
    ''' </summary>
    ''' <param name="shp_ImageToBeClipped"></param>
    ''' <returns></returns>
    Public Function rct_get_ImageToBeClippedShapType(ByRef shp_ImageToBeClipped As Word.Shape) As String
        Dim aspectRatio_ImageToBeClipped As Single
        Dim strImageToBeClipped As String

        strImageToBeClipped = ""
        aspectRatio_ImageToBeClipped = shp_ImageToBeClipped.Height / shp_ImageToBeClipped.Width
        '
        If aspectRatio_ImageToBeClipped > 1 Then strImageToBeClipped = "orginal_image_is_portrait"
        If aspectRatio_ImageToBeClipped = 1 Then strImageToBeClipped = "original_image_is square"
        If aspectRatio_ImageToBeClipped < 1 Then strImageToBeClipped = "original_image_is_landscape"
        '
        Return strImageToBeClipped
    End Function
    '
    ''' <summary>
    ''' This method will return a string indictaing the Shape Type of the underlying
    ''' shape associated with the inout cShapeMgr shpImageToBeFilled
    ''' "cropRect_is_portrait", "cropRect_is square", "cropRect_is_landscape"
    ''' </summary>
    ''' <param name="shpImageToBeFilled"></param>
    ''' <returns></returns>
    Public Function rct_get_ImageToBeFilledShapeType(ByRef shpImageToBeFilled As cShapeMgr) As String
        Dim aspectRatioImageToBeFilled As Single
        Dim strShpAspectRatio_CropRect As String
        '
        strShpAspectRatio_CropRect = ""
        aspectRatioImageToBeFilled = shpImageToBeFilled.aspectRatio       'Me.height / Me.width
        '
        If aspectRatioImageToBeFilled > 1 Then strShpAspectRatio_CropRect = "cropRect_is_portrait"
        If aspectRatioImageToBeFilled = 1 Then strShpAspectRatio_CropRect = "cropRect_is square"
        If aspectRatioImageToBeFilled < 1 Then strShpAspectRatio_CropRect = "cropRect_is_landscape"
        '

        Return strShpAspectRatio_CropRect
    End Function
    '
    Public Function rct_(ByRef objSrcShp_tobeFilled As cShapeMgr, ByRef shp_ImageToBeClipped As Word.Shape) As Collection
        Dim lstOfDimensions As New Collection
        '
        Dim aspectRatio_CropRect, aspectRatio_ImageToBeClipped As Single
        Dim dummy_h, dummy_w, scaleFactor_h, scalefactor_w As Single
        Dim clipHeight, clipWidth As Single
        Dim strShpAspectRatio_CropRect As String
        Dim strImageToBeClipped As String
        Dim hf As HeaderFooter
        Dim rng As Word.Range
        Dim irror As Boolean
        Dim objRectMgr As New cCropRectMgr()
        Dim i As Integer
        '
        'hf = sect.Headers(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        hf = objSrcShp_tobeFilled.hf
        rng = shp_ImageToBeClipped.Anchor
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        strShpAspectRatio_CropRect = ""
        strImageToBeClipped = ""
        irror = False
        '
        'Get the aspect ratio form the shape that will eventually be replaced
        'This is the aspect ratio of the cropping rect
        aspectRatio_CropRect = objSrcShp_tobeFilled.aspectRatio       'Me.height / Me.width
        aspectRatio_ImageToBeClipped = shp_ImageToBeClipped.Height / shp_ImageToBeClipped.Width
        '
        strImageToBeClipped = Me.rct_get_ImageToBeClippedShapType(shp_ImageToBeClipped)
        strShpAspectRatio_CropRect = Me.rct_get_ImageToBeFilledShapeType(objSrcShp_tobeFilled)
        '
        Select Case strImageToBeClipped
            Case "orginal_image_is_portrait", "original_image_is square"
                Select Case strShpAspectRatio_CropRect
                    Case "cropRect_is_portrait", "cropRect_is square"
                        '
                        '***Verified 20231030
                        '
                        clipHeight = shp_ImageToBeClipped.Height
                        clipWidth = clipHeight / aspectRatio_CropRect
                        '
                        If clipWidth > shp_ImageToBeClipped.Width Then
                            'We need to narrow the rectangle
                            scaleFactor_h = 1.0
                            '
                            For i = 1 To 10000
                                dummy_h = shp_ImageToBeClipped.Height * scaleFactor_h
                                dummy_w = dummy_h / aspectRatio_CropRect
                                '
                                If dummy_w < (0.999 * shp_ImageToBeClipped.Width) Then
                                    clipHeight = dummy_h
                                    clipWidth = dummy_w
                                    Exit For
                                End If
                                '
                                scaleFactor_h = scaleFactor_h - 0.001
                                If scaleFactor_h < 0.1 Then
                                    irror = True
                                    Exit For
                                End If
                            Next
                        End If

                    Case "cropRect_is_landscape"
                        '
                        '*** Seems OK 20231030
                        '
                        clipWidth = shp_ImageToBeClipped.Width
                        clipHeight = aspectRatio_CropRect * clipWidth
                        '
                        If clipHeight > shp_ImageToBeClipped.Height Then
                            'We need to narrow the rectangle
                            scalefactor_w = 1.0
                            '
                            For i = 1 To 10000
                                dummy_w = shp_ImageToBeClipped.Width * scalefactor_w
                                dummy_h = dummy_w * aspectRatio_CropRect
                                '
                                If dummy_h <= (0.999 * shp_ImageToBeClipped.Height) Then
                                    clipHeight = dummy_h
                                    clipWidth = dummy_w
                                    Exit For
                                End If
                                '
                                scalefactor_w = scalefactor_w - 0.001
                                If scalefactor_w < 0.1 Then
                                    'Error condition, set to a default
                                    clipHeight = 1
                                    clipWidth = clipHeight * aspectRatio_CropRect
                                    irror = True
                                    Exit For
                                End If
                            Next
                        End If


                End Select

            Case "original_image_is_landscape"
                Select Case strShpAspectRatio_CropRect
                    Case "cropRect_is_portrait", "cropRect_is square"
                        '
                        '**** Verified (not extensively), seems OK 20231030
                        '
                        clipHeight = shp_ImageToBeClipped.Height
                        clipWidth = clipHeight / aspectRatio_CropRect
                        '
                        If clipWidth > shp_ImageToBeClipped.Width Then
                            'We need to narrow the rectangle
                            scaleFactor_h = 1.0
                            '
                            For i = 1 To 10000
                                dummy_h = shp_ImageToBeClipped.Height * scaleFactor_h
                                dummy_w = dummy_h / aspectRatio_CropRect
                                '
                                If dummy_h < (0.999 * shp_ImageToBeClipped.Height) Then
                                    clipHeight = dummy_h
                                    clipWidth = dummy_w
                                    Exit For
                                End If
                                '
                                scaleFactor_h = scaleFactor_h - 0.001
                                If scaleFactor_h < 0.1 Then
                                    irror = True
                                    Exit For
                                End If
                            Next
                        End If

                    Case "cropRect_is_landscape"
                        '
                        '**** Verified 20231030
                        '
                        clipWidth = shp_ImageToBeClipped.Width
                        clipHeight = aspectRatio_CropRect * clipWidth
                        '
                        If clipHeight > shp_ImageToBeClipped.Height Then
                            'We need to narrow the rectangle
                            scalefactor_w = 1.0
                            '
                            For i = 1 To 10000
                                dummy_w = shp_ImageToBeClipped.Width * scalefactor_w
                                dummy_h = dummy_w * aspectRatio_CropRect
                                '
                                If dummy_h <= (0.999 * shp_ImageToBeClipped.Height) Then
                                    clipHeight = dummy_h
                                    clipWidth = dummy_w
                                    Exit For
                                End If
                                '
                                scalefactor_w = scalefactor_w - 0.001
                                If scalefactor_w < 0.1 Then
                                    'Error condition, set to a default
                                    clipHeight = 1
                                    clipWidth = clipHeight * aspectRatio_CropRect
                                    irror = True
                                    Exit For
                                End If
                            Next
                            'MsgBox(i.ToString())
                        End If
loop6:

                End Select
        End Select
        '
        lstOfDimensions.Add(clipHeight, "clipHeight")
        lstOfDimensions.Add(clipWidth, "clipWIdth")
        '

        Return lstOfDimensions
    End Function


End Class

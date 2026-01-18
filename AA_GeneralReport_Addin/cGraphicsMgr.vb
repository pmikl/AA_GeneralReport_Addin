Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.Drawing
Imports System.Windows.Forms
Imports System.IO
Public Class cGraphicsMgr
    Inherits cGlobals
    '
    Public _fillColor As Color                          'Default fill colour (dark background)
    '
    Public Sub New()
        MyBase.New()
        Me._fillColor = Color.FromArgb(28, 3, 61)
    End Sub
    '
    '
    ''' <summary>
    ''' This method will paste a graphic at the range rng. The width and height are specified in points.
    ''' The colour fill is set to the default, but it can be overidden if the rgb values are set. The method will
    ''' return the graphic as an inline shape.
    ''' </summary>
    ''' <param name="drCell"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <returns></returns>
    Public Function grfx_inline_insertShape(ByRef drCell As Word.Cell, width As Single, height As Single, shpFillColor As Color, Optional strCodingType As String = "pngIdx") As Word.InlineShape
        Dim rng As Word.Range
        Dim shpInline As Word.InlineShape
        '
        'width = 600.0
        'height = 250.0
        '
        rng = drCell.Range
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        shpInline = Me.grfx_insertShape_inline(rng, width, height, shpFillColor, strCodingType)
        '
        shpInline.AlternativeText = "Grey blank picture placeholder"
        Return shpInline
        '
    End Function
    '
    ''' <summary>
    ''' This method will paste a graphic at the range rng. The width and height are specified in points.
    ''' The colour fill is set to the default, but it can be overidden if the rgb values are set. The method will
    ''' return the graphic as an inline shape.
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <returns></returns>
    Public Function grfx_insertShape_inline(ByRef rng As Word.Range, width As Single, height As Single, shpFillColor As Color, Optional strCodingType As String = "pngIdx", Optional xDpi As Single = 150, Optional yDpi As Single = 150) As Word.InlineShape
        Dim bmp As Bitmap
        Dim myDoc As Word.Document
        Dim grfx As Graphics
        Dim brush As SolidBrush
        'Dim shp As Word.Shape
        Dim shpInline As Word.InlineShape
        Dim widthPx, heightPx As Single
        Dim objGlobals As New cGlobals()
        Dim oldPictWrapType As WdWrapTypeMerged
        'Dim palette As Imaging.ColorPalette
        '
        widthPx = xDpi * width / 72
        heightPx = yDpi * height / 72
        '
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        myDoc = rng.Document
        '
        bmp = New Bitmap(widthPx, heightPx, Imaging.PixelFormat.Format24bppRgb)
        'bmp = New Bitmap(width, height, Imaging.PixelFormat.Format8bppIndexed)
        bmp.SetResolution(xDpi, yDpi)
        '
        grfx = Graphics.FromImage(bmp)
        '
        'Set the default colour for the fill, but if any of the r,g,b values is negative, then
        'generate the colour from the r,g,b values
        brush = New SolidBrush(shpFillColor)
        '
        grfx.FillRectangle(brush, 0, 0, widthPx, heightPx)
        '
        Select Case strCodingType
            Case "png"
                bmp = Me.grfx_Convert_ToPNG(bmp)
                '
            Case "pngIdx"
                bmp = Me.grfx_convert_ToIndexedColor(bmp, shpFillColor)
                bmp = Me.grfx_Convert_ToPNG(bmp)
                '
        End Select
        '
        '****
        'Get the current insert/paste picture option (inline, behind text etc). Change it
        'to inline htem change it back.. Tested 20210807 and it works
        oldPictWrapType = objGlobals.glb_get_wrdApp.Options.PictureWrapType
        '
        objGlobals.glb_get_wrdApp.Options.PictureWrapType = WdWrapTypeMerged.wdWrapMergeInline
        Clipboard.SetImage(bmp)
        rng.Paste()
        '
        objGlobals.glb_get_wrdApp.Options.PictureWrapType = oldPictWrapType
        '
        '****
        '
        'shp = rng.ShapeRange.Item(1)
        shpInline = rng.InlineShapes.Item(1)

        'shpInline = shp.ConvertToInlineShape()
        '
        'Now stretch the shape to the right size in points
        shpInline.LockAspectRatio = False
        'shpInline.Width = width
        'shpInline.Height = height


        Return shpInline
        '
    End Function
    '
    '
    Public Function grfx_Convert_ToPNG(ByRef bmp As Drawing.Bitmap) As Drawing.Bitmap
        '
        'https://jwcooney.com/2019/03/03/vb-net-reading-a-png-image-modifying-it-then-sending-it-to-the-user/

        Dim result() As Byte
        Dim memStream As New IO.MemoryStream
        Dim numBytes As Long
        'Dim encoder As Imaging.ImageCodecInfo
        'Dim result As Drawing.Image
        '
        bmp.Save(memStream, Imaging.ImageFormat.Png)
        result = memStream.ToArray()
        '
        bmp = Image.FromStream(memStream)
        numBytes = memStream.Length
        '
        'MsgBox("Size is " + CStr(numBytes) + " bytes")
        '
        Return bmp
    End Function
    '
    Public Function grfx_Convert_ToTiff(ByRef bmp As Drawing.Bitmap) As Drawing.Bitmap
        '
        'https://jwcooney.com/2019/03/03/vb-net-reading-a-png-image-modifying-it-then-sending-it-to-the-user/

        Dim result() As Byte
        Dim memStream As New IO.MemoryStream
        Dim numBytes As Long
        'Dim encoder As Imaging.ImageCodecInfo
        'Dim result As Drawing.Image
        '
        bmp.Save(memStream, Imaging.ImageFormat.Tiff)
        result = memStream.ToArray()
        '
        bmp = Image.FromStream(memStream)
        numBytes = memStream.Length
        '
        'MsgBox("Size is " + CStr(numBytes) + " bytes")
        '
        Return bmp
    End Function
    '
    '
    Public Function grfx_Convert_ToJpg(ByRef bmp As Drawing.Bitmap) As Drawing.Bitmap
        '
        'https://jwcooney.com/2019/03/03/vb-net-reading-a-png-image-modifying-it-then-sending-it-to-the-user/

        Dim result() As Byte
        Dim memStream As New IO.MemoryStream
        Dim numBytes As Long
        'Dim encoder As Imaging.ImageCodecInfo
        'Dim result As Drawing.Image
        '
        bmp.Save(memStream, Imaging.ImageFormat.Jpeg)
        result = memStream.ToArray()
        '
        bmp = Image.FromStream(memStream)
        numBytes = memStream.Length
        '
        'MsgBox("Size is " + CStr(numBytes) + " bytes")
        '
        Return bmp
    End Function
    '
    '
    Public Function grfx_convert_ToIndexedColor(ByRef bmp As Drawing.Bitmap, shpFillcolor As Drawing.Color) As Drawing.Bitmap
        'https://stackoverflow.com/questions/9010761/c-sharp-convert-bitmap-to-indexed-colour-format
        '
        Dim result As Bitmap
        Dim palette As Imaging.ColorPalette
        'Dim col As Color
        Dim j As Integer
        'Dim memStream As New IO.MemoryStream()
        '
        'palette = bmp.Palette
        'palette = New Imaging.ColorPalette()
        'result = bmp.Clone(New Drawing.Rectangle(0, 0, bmp.Width, bmp.Height), Imaging.PixelFormat.Format8bppIndexed)
        'bmp.
        'bmp.Save(memStream, Imaging.PixelFormat.Format8bppIndexed)
        result = bmp.Clone(New Drawing.Rectangle(0, 0, bmp.Width, bmp.Height), Imaging.PixelFormat.Format8bppIndexed)
        palette = result.Palette
        '
        For j = 0 To palette.Entries.Count - 1
            palette.Entries(j) = shpFillcolor
        Next
        'For Each col In palette.Entries

        ' Next
        'palette.Entries(0) = Color.FromArgb(255, 255, 255)      'background
        'palette.Entries(1) = Color.FromArgb(233, 233, 233)      'background
        'palette.Entries(215) = Color.FromArgb(233, 233, 233)      'light purple
        'palette.Entries(3) = Color.FromArgb(38, 13, 72)      'mid purple
        'palette.Entries(4) = Color.FromArgb(47, 27, 80)      'mid purple

        '
        'bmp.Palette.Entries.
        'palette.Entries(1) = Color.FromArgb(255, 0, 0)
        'palette.Entries(2) = Color.FromArgb(255, 0, 0)
        '
        'result.Palette = palette
        'result.Palette.Entries
        '
        result.Palette = palette
        Return result
    End Function

    '
    ''' <summary>
    ''' This method is meant to build and insert a CoverPage "filled Image"... It should work, but as of 20210823 it was generating
    ''' intermittent unexplained errors in Visual Studio Debug mode... I suspect the problem is in "PasteAsPNG" but I just don't know
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="widthPts"></param>
    ''' <param name="heightPts"></param>
    ''' <param name="xDpi"></param>
    ''' <param name="yDpi"></param>
    ''' <returns></returns>
    Public Function grfx_insert_ImageCP(ByRef sect As Word.Section, Optional widthPts As Single = 480.8, Optional heightPts As Single = 320.05, Optional xDpi As Single = 300, Optional yDpi As Single = 300) As Word.Shape
        Dim shp As Word.Shape
        Dim bmp As Bitmap
        Dim grfx As Graphics
        Dim brush As SolidBrush
        Dim numSegments As Integer
        Dim segmentWidthPx, segmentWidthPts As Single
        Dim widthPx, heightPx As Single
        Dim hf As Word.HeaderFooter
        Dim lstOfColors As New List(Of Drawing.Color)
        Dim rng As Word.Range
        'Dim trianglePts(4), rectPts(5) As Drawing.PointF
        Dim trianglePts(4) As Drawing.PointF
        '
        shp = Nothing
        '
        lstOfColors.Add(Color.FromArgb(255, 0, 0))
        'lstOfColors.Add(Color.FromArgb(41, 8, 75))
        lstOfColors.Add(Color.FromArgb(65, 33, 103))
        lstOfColors.Add(Color.FromArgb(110, 94, 136))
        lstOfColors.Add(Color.FromArgb(89, 67, 121))
        '
        hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterFirstPage)
        If hf.Exists Then
            rng = hf.Range
        Else
            hf = sect.Headers.Item(WdHeaderFooterIndex.wdHeaderFooterPrimary)
            rng = hf.Range
        End If
        '
        'rng = glb_get_wrdSelRng()
        rng.Collapse(WdCollapseDirection.wdCollapseEnd)
        numSegments = 3
        If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
            numSegments = 2
            widthPts = 350.0
            heightPts = 350.0
        End If
        '
        segmentWidthPts = widthPts / numSegments
        widthPx = xDpi * widthPts / 72
        heightPx = yDpi * heightPts / 72
        segmentWidthPx = xDpi * segmentWidthPts / 72
        '
        bmp = New Bitmap(widthPx, heightPx, Imaging.PixelFormat.Format24bppRgb)
        bmp.SetResolution(xDpi, yDpi)
        grfx = Graphics.FromImage(bmp)
        brush = New SolidBrush(lstOfColors.Item(0))                                 'Set the default fill Colour
        grfx.FillRectangle(brush, 0, 0, widthPx, heightPx)                          'Fill the shape
        '
        Me.grfx_buildSegments_CoverPgPrt(grfx, lstOfColors, numSegments, segmentWidthPx, heightPx)
        '
        Try
            Me.grfx_ChptBanner_pasteAsPNG(hf, bmp, widthPts, heightPts)
        Catch ex As Exception
            MsgBox("Error in 'Paste as PNG'")
        End Try
        '
        Return shp
    End Function

    '
    Public Function grfx_insert_ImageChapterBanner(ByRef rng As Word.Range, Optional heightPts As Single = 157, Optional xDpi As Single = 300, Optional yDpi As Single = 300) As Word.Shape
        Dim bmp As Bitmap
        Dim grfx As Graphics
        Dim brush As SolidBrush
        Dim shp As Word.Shape
        Dim numSegments As Integer
        Dim segmentWidthPx, segmentWidthPts, xLoc As Single
        Dim widthPx, heightPx, widthBetweenMargins As Single
        Dim sect As Word.Section
        '
        shp = Nothing
        numSegments = 5
        sect = rng.Sections.Item(1)
        widthBetweenMargins = sect.PageSetup.PageWidth - (sect.PageSetup.LeftMargin + sect.PageSetup.RightMargin)
        '
        If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then numSegments = 8

        'numSegments is 5 for Portrait and 8 for Landscape
        'For Portrait, the standard segment width is 79.4 pt
        segmentWidthPts = widthBetweenMargins / numSegments
        widthPx = xDpi * widthBetweenMargins / 72
        heightPx = yDpi * heightPts / 72
        '
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        bmp = New Bitmap(widthPx, heightPx, Imaging.PixelFormat.Format24bppRgb)
        bmp.SetResolution(xDpi, yDpi)
        grfx = Graphics.FromImage(bmp)
        brush = New SolidBrush(Me._fillColor)                                       'Set the default fill Colour
        grfx.FillRectangle(brush, 0, 0, widthPx, heightPx)                          'Fill the shape

        xLoc = 0
        segmentWidthPx = xDpi * (segmentWidthPts / 72)
        MsgBox("Segment width = " + CStr(segmentWidthPts) + " pts")
        '
        Me.grfx_buildSegments_ChptBanner(grfx, numSegments, segmentWidthPx, heightPx)
        'Clipboard.SetImage(bmp)
        'rng.Paste()

        Me.grfx_ChptBanner_pasteAsPNG(rng, bmp, widthBetweenMargins, heightPts)
        '
        Return shp

    End Function
    '
#Region "Image Segments"
    Public Sub grfx_buildSegments_CoverPgPrt(ByRef grfx As Graphics, ByRef lstOfColors As List(Of Drawing.Color), numSegments As Integer, segmentWidthPx As Single, heightPx As Single)
        Dim shp As Word.Shape
        Dim brush As SolidBrush
        Dim trianglePts(4) As Drawing.PointF
        Dim xLoc As Single
        Dim pt1, pt2, pt3, pt4 As Drawing.PointF
        '
        shp = Nothing
        brush = New SolidBrush(lstOfColors.Item(0))                                       'Set the default fill Colour
        '
        xLoc = 0
        Try
            For j = 1 To numSegments
                Select Case j
                    Case 1
                        'Panel Top
                        pt1 = New Drawing.Point(xLoc, 0)
                        pt2 = New Drawing.Point(xLoc + segmentWidthPx, 0)
                        pt3 = New Drawing.Point(xLoc + segmentWidthPx, heightPx / 2)
                        pt4 = New Drawing.Point(xLoc, 0)
                        '
                        trianglePts(1) = pt1
                        trianglePts(2) = pt2
                        trianglePts(3) = pt3
                        trianglePts(4) = pt4
                        '
                        brush.Color = lstOfColors.Item(1)
                        grfx.FillPolygon(brush, trianglePts)
                        '
                        'Panel Top left
                        pt1 = New Drawing.Point(xLoc, 0)
                        pt2 = New Drawing.Point(xLoc + segmentWidthPx / 2, heightPx / 4)
                        pt3 = New Drawing.Point(xLoc, heightPx / 2)
                        pt4 = New Drawing.Point(xLoc, 0)
                        '
                        trianglePts(1) = pt1
                        trianglePts(2) = pt2
                        trianglePts(3) = pt3
                        trianglePts(4) = pt4
                        '
                        brush.Color = lstOfColors.Item(2)
                        grfx.FillPolygon(brush, trianglePts)
                        '
                        'Panel Mid Left
                        pt1 = New Drawing.Point(xLoc, heightPx / 2)
                        pt2 = New Drawing.Point(xLoc + segmentWidthPx / 2, heightPx / 4)
                        pt3 = New Drawing.Point(xLoc + segmentWidthPx / 2, 3 * heightPx / 4)
                        pt4 = New Drawing.Point(xLoc, heightPx / 2)
                        '
                        trianglePts(1) = pt1
                        trianglePts(2) = pt2
                        trianglePts(3) = pt3
                        trianglePts(4) = pt4
                        '
                        brush.Color = lstOfColors.Item(3)
                        grfx.FillPolygon(brush, trianglePts)
                        '
                        'Panel 3/4 RightLeft
                        pt1 = New Drawing.Point(xLoc + segmentWidthPx / 2, 3 * heightPx / 4)
                        pt2 = New Drawing.Point(xLoc + segmentWidthPx, heightPx / 2)
                        pt3 = New Drawing.Point(xLoc + 3 * segmentWidthPx / 2, 3 * heightPx / 4)
                        pt4 = New Drawing.Point(xLoc + segmentWidthPx / 2, 3 * heightPx / 4)
                        '
                        trianglePts(1) = pt1
                        trianglePts(2) = pt2
                        trianglePts(3) = pt3
                        trianglePts(4) = pt4
                        '
                        brush.Color = lstOfColors.Item(3)
                        grfx.FillPolygon(brush, trianglePts)
                        '
                        'Panel Bottom Left
                        pt1 = New Drawing.Point(xLoc, heightPx / 2)
                        pt2 = New Drawing.Point(xLoc + segmentWidthPx / 2, 3 * heightPx / 4)
                        pt3 = New Drawing.Point(xLoc, heightPx)
                        pt4 = New Drawing.Point(xLoc, heightPx / 2)
                        '
                        trianglePts(1) = pt1
                        trianglePts(2) = pt2
                        trianglePts(3) = pt3
                        trianglePts(4) = pt4
                        '
                        brush.Color = lstOfColors.Item(1)
                        grfx.FillPolygon(brush, trianglePts)
                        '
                        'Panel Bottom Mid
                        pt1 = New Drawing.Point(xLoc, heightPx)
                        pt2 = New Drawing.Point(xLoc + segmentWidthPx / 2, 3 * heightPx / 4)
                        pt3 = New Drawing.Point(xLoc + segmentWidthPx, heightPx)
                        pt4 = New Drawing.Point(xLoc, heightPx)
                        '
                        trianglePts(1) = pt1
                        trianglePts(2) = pt2
                        trianglePts(3) = pt3
                        trianglePts(4) = pt4
                        '
                        brush.Color = lstOfColors.Item(2)
                        grfx.FillPolygon(brush, trianglePts)
                        '
                        '
                        'Panel Bottom Right
                        pt1 = New Drawing.Point(xLoc + segmentWidthPx, heightPx)
                        pt2 = New Drawing.Point(xLoc + 3 * segmentWidthPx / 2, 3 * heightPx / 4)
                        pt3 = New Drawing.Point(xLoc + 2 * segmentWidthPx, heightPx)
                        pt4 = New Drawing.Point(xLoc + segmentWidthPx, heightPx)
                        '
                        trianglePts(1) = pt1
                        trianglePts(2) = pt2
                        trianglePts(3) = pt3
                        trianglePts(4) = pt4
                        '
                        brush.Color = lstOfColors.Item(3)
                        grfx.FillPolygon(brush, trianglePts)
                    '
                    Case 2
                        '
                        'Panel Top
                        pt1 = New Drawing.Point(xLoc, 0)
                        pt2 = New Drawing.Point(xLoc + segmentWidthPx, 0)
                        pt3 = New Drawing.Point(xLoc, heightPx / 2)
                        pt4 = New Drawing.Point(xLoc, 0)
                        '
                        trianglePts(1) = pt1
                        trianglePts(2) = pt2
                        trianglePts(3) = pt3
                        trianglePts(4) = pt4
                        '
                        brush.Color = lstOfColors.Item(3)
                        grfx.FillPolygon(brush, trianglePts)
                        '
                        '
                        'Panel Mid Right
                        pt1 = New Drawing.Point(xLoc, heightPx / 2)
                        pt2 = New Drawing.Point(xLoc + segmentWidthPx, 0)
                        pt3 = New Drawing.Point(xLoc + segmentWidthPx, heightPx)
                        pt4 = New Drawing.Point(xLoc, heightPx / 2)
                        '
                        trianglePts(1) = pt1
                        trianglePts(2) = pt2
                        trianglePts(3) = pt3
                        trianglePts(4) = pt4
                        '
                        brush.Color = lstOfColors.Item(1)
                        grfx.FillPolygon(brush, trianglePts)
                    '
                    Case 3
                        'Panel Top
                        pt1 = New Drawing.Point(xLoc, 0)
                        pt2 = New Drawing.Point(xLoc + segmentWidthPx, 0)
                        pt3 = New Drawing.Point(xLoc + segmentWidthPx, heightPx / 2)
                        pt4 = New Drawing.Point(xLoc, 0)
                        '
                        trianglePts(1) = pt1
                        trianglePts(2) = pt2
                        trianglePts(3) = pt3
                        trianglePts(4) = pt4
                        '
                        brush.Color = lstOfColors.Item(2)
                        grfx.FillPolygon(brush, trianglePts)
                        '
                        'Panel Top left
                        pt1 = New Drawing.Point(xLoc, 0)
                        pt2 = New Drawing.Point(xLoc + segmentWidthPx / 2, heightPx / 4)
                        pt3 = New Drawing.Point(xLoc, heightPx / 2)
                        pt4 = New Drawing.Point(xLoc, 0)
                        '
                        trianglePts(1) = pt1
                        trianglePts(2) = pt2
                        trianglePts(3) = pt3
                        trianglePts(4) = pt4
                        '
                        brush.Color = lstOfColors.Item(3)
                        grfx.FillPolygon(brush, trianglePts)
                        '
                        'Panel Mid Left
                        pt1 = New Drawing.Point(xLoc, heightPx / 2)
                        pt2 = New Drawing.Point(xLoc + segmentWidthPx / 2, heightPx / 4)
                        pt3 = New Drawing.Point(xLoc + segmentWidthPx / 2, 3 * heightPx / 4)
                        pt4 = New Drawing.Point(xLoc, heightPx / 2)
                        '
                        trianglePts(1) = pt1
                        trianglePts(2) = pt2
                        trianglePts(3) = pt3
                        trianglePts(4) = pt4
                        '
                        brush.Color = lstOfColors.Item(0)
                        grfx.FillPolygon(brush, trianglePts)
                        '
                        '
                        'Panel Bottom Left
                        pt1 = New Drawing.Point(xLoc, heightPx / 2)
                        pt2 = New Drawing.Point(xLoc + segmentWidthPx / 2, 3 * heightPx / 4)
                        pt3 = New Drawing.Point(xLoc, heightPx)
                        pt4 = New Drawing.Point(xLoc, heightPx / 2)
                        '
                        trianglePts(1) = pt1
                        trianglePts(2) = pt2
                        trianglePts(3) = pt3
                        trianglePts(4) = pt4
                        '
                        brush.Color = lstOfColors.Item(2)
                        grfx.FillPolygon(brush, trianglePts)
                        '
                        'Panel Bottom Mid
                        pt1 = New Drawing.Point(xLoc, heightPx)
                        pt2 = New Drawing.Point(xLoc + segmentWidthPx / 2, 3 * heightPx / 4)
                        pt3 = New Drawing.Point(xLoc + segmentWidthPx, heightPx)
                        pt4 = New Drawing.Point(xLoc, heightPx)
                        '
                        trianglePts(1) = pt1
                        trianglePts(2) = pt2
                        trianglePts(3) = pt3
                        trianglePts(4) = pt4
                        '
                        brush.Color = lstOfColors.Item(1)
                        grfx.FillPolygon(brush, trianglePts)
                        '
                        '
                        'Panel Bottom Right
                        pt1 = New Drawing.Point(xLoc + segmentWidthPx / 2, 3 * heightPx / 4)
                        pt2 = New Drawing.Point(xLoc + segmentWidthPx, heightPx / 2)
                        pt3 = New Drawing.Point(xLoc + segmentWidthPx, heightPx)
                        pt4 = New Drawing.Point(xLoc + segmentWidthPx / 2, 3 * heightPx / 4)
                        '
                        trianglePts(1) = pt1
                        trianglePts(2) = pt2
                        trianglePts(3) = pt3
                        trianglePts(4) = pt4
                        '
                        brush.Color = lstOfColors.Item(3)
                        grfx.FillPolygon(brush, trianglePts)
                        '
                End Select
                xLoc = xLoc + segmentWidthPx
            Next

        Catch ex As Exception
            MsgBox("Error in 'Build Segmenst Section'")

        End Try

    End Sub
    Public Sub grfx_buildSegments_ChptBanner(ByRef grfx As Graphics, numSegments As Integer, segmentWidthPx As Single, heightPx As Single)
        Dim brush As SolidBrush
        Dim pt1, pt2, pt3, pt4, pt5 As Drawing.PointF
        Dim trianglePts(4), rectPts(5) As Drawing.PointF
        Dim xLoc As Single
        '
        brush = New SolidBrush(Me._fillColor)                                       'Set the default fill Colour
        xLoc = 0
        '
        For j = 1 To numSegments
            Select Case j
                Case 1
                    '
                    'Panel (Top)
                    pt1 = New Drawing.Point(xLoc, 0)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx, 0)
                    pt3 = New Drawing.Point(xLoc + segmentWidthPx, heightPx / 2)
                    pt4 = New Drawing.Point(xLoc, heightPx / 2)
                    pt5 = New Drawing.Point(xLoc, 0)
                    '
                    rectPts(1) = pt1
                    rectPts(2) = pt2
                    rectPts(3) = pt3
                    rectPts(4) = pt4
                    rectPts(5) = pt5
                    '
                    brush.Color = Color.FromArgb(47, 27, 80)
                    'brush.Color = bmp.Palette.Entries(3)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, rectPts)
                    '
                    '
                    pt1 = New Drawing.Point(xLoc, heightPx / 2)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx / 2, heightPx / 4)
                    pt3 = New Drawing.Point(xLoc + segmentWidthPx / 2, 3 * heightPx / 4)
                    pt4 = New Drawing.Point(xLoc, heightPx / 2)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Me._fillColor
                    'brush = New SolidBrush(bmp.Palette.Entries(2))
                    'brush.Color = bmp.Palette.Entries(2)
                    '
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    '
                    pt1 = New Drawing.Point(xLoc, 0)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx / 2, heightPx / 4)
                    pt3 = New Drawing.Point(xLoc, heightPx / 2)
                    pt4 = New Drawing.Point(xLoc, 0)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(56, 37, 86)
                    'brush = New SolidBrush(bmp.Palette.Entries(2))
                    'brush.Color = bmp.Palette.Entries(2)
                    '
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)


                    'brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    pt1 = New Drawing.Point(xLoc, heightPx / 2)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx / 2, 3 * heightPx / 4)
                    pt3 = New Drawing.Point(xLoc, heightPx)
                    pt4 = New Drawing.Point(xLoc, heightPx / 2)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(47, 27, 80)
                    'brush.Color = bmp.Palette.Entries(3)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    pt1 = New Drawing.Point(xLoc + segmentWidthPx / 2, heightPx / 4)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx, heightPx / 2)
                    pt3 = New Drawing.Point(xLoc + segmentWidthPx / 2, 3 * heightPx / 4)
                    pt4 = New Drawing.Point(xLoc + segmentWidthPx / 2, heightPx / 4)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(38, 13, 72)
                    'brush.Color = bmp.Palette.Entries(4)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '

                    'GoTo finis
                    '
                    'Panel (bottom triangle, top)
                    pt1 = New Drawing.Point(xLoc + segmentWidthPx / 2, 3 * heightPx / 4)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx, heightPx / 2)
                    pt3 = New Drawing.Point(xLoc + 3 * segmentWidthPx / 2, 3 * heightPx / 4)
                    pt4 = New Drawing.Point(xLoc + segmentWidthPx / 2, 3 * heightPx / 4)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    '
                    'Panel (bottom triangle, left)
                    pt1 = New Drawing.Point(xLoc, heightPx)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx / 2, 3 * heightPx / 4)
                    pt3 = New Drawing.Point(xLoc + segmentWidthPx, heightPx)
                    pt4 = New Drawing.Point(xLoc, heightPx)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(38, 13, 72)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    '
                    'Panel (bottom triangle, right)
                    pt1 = New Drawing.Point(xLoc + segmentWidthPx, heightPx)
                    pt2 = New Drawing.Point(xLoc + 3 * segmentWidthPx / 2, 3 * heightPx / 4)
                    pt3 = New Drawing.Point(xLoc + 2 * segmentWidthPx, heightPx)
                    pt4 = New Drawing.Point(xLoc + segmentWidthPx, heightPx)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(47, 27, 80)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                Case 2
                    pt1 = New Drawing.Point(xLoc, heightPx / 2)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx, 0)
                    pt3 = New Drawing.Point(xLoc + segmentWidthPx, heightPx)
                    pt4 = New Drawing.Point(xLoc, heightPx / 2)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(38, 13, 72)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    'brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                Case 3
                    pt1 = New Drawing.Point(xLoc, 0)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx, heightPx / 2)
                    pt3 = New Drawing.Point(xLoc, heightPx)
                    pt4 = New Drawing.Point(xLoc, 0)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    '                    '
                    'Panel (bottom triangle, left)
                    pt1 = New Drawing.Point(xLoc, heightPx)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx, heightPx / 2)
                    pt3 = New Drawing.Point(xLoc + segmentWidthPx, heightPx)
                    pt4 = New Drawing.Point(xLoc, heightPx)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(38, 13, 72)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '

                Case 4
                    pt1 = New Drawing.Point(xLoc, heightPx / 2)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx, 0)
                    pt3 = New Drawing.Point(xLoc + segmentWidthPx, heightPx)
                    pt4 = New Drawing.Point(xLoc, heightPx / 2)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)

                    'brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)

                    'Case 1, 3, 5
                    'brush.Color = Color.FromArgb(r, g, b)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)

                    'Case 2, 4
                    'brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    '
                    '
                    '                    '
                    'Panel (bottom triangle, right)
                    pt1 = New Drawing.Point(xLoc, heightPx / 2)
                    pt2 = New Drawing.Point(xLoc, 0)
                    pt3 = New Drawing.Point(xLoc + segmentWidthPx, 0)
                    pt4 = New Drawing.Point(xLoc, heightPx / 2)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(38, 13, 72)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                Case 5
                    pt1 = New Drawing.Point(xLoc, 0)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx, heightPx / 2)
                    pt3 = New Drawing.Point(xLoc, heightPx)
                    pt4 = New Drawing.Point(xLoc, 0)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(38, 13, 72)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    'Panel (top triangle)
                    pt1 = New Drawing.Point(xLoc, 0)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx, 0)
                    pt3 = New Drawing.Point(xLoc + segmentWidthPx / 2, heightPx / 4)
                    pt4 = New Drawing.Point(xLoc, 0)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(47, 27, 80)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    '
                    'Panel (top right triangle)
                    pt1 = New Drawing.Point(xLoc + segmentWidthPx / 2, heightPx / 4)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx, heightPx / 4)
                    pt3 = New Drawing.Point(xLoc + segmentWidthPx, heightPx / 2)
                    pt4 = New Drawing.Point(xLoc + segmentWidthPx / 2, heightPx / 4)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                Case 6
                    '
                    'top  triangle
                    pt1 = New Drawing.Point(xLoc, 0)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx, 0)
                    pt3 = New Drawing.Point(xLoc + segmentWidthPx / 2, heightPx / 4)
                    pt4 = New Drawing.Point(xLoc, 0)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    '  triangle
                    pt1 = New Drawing.Point(xLoc, heightPx / 4)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx / 2, heightPx / 4)
                    pt3 = New Drawing.Point(xLoc, heightPx / 2)
                    pt4 = New Drawing.Point(xLoc, heightPx / 4)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    pt1 = New Drawing.Point(xLoc + segmentWidthPx, 0)
                    pt2 = New Drawing.Point(xLoc, heightPx / 2)
                    pt3 = New Drawing.Point(xLoc + segmentWidthPx, heightPx)
                    pt4 = New Drawing.Point(xLoc + segmentWidthPx, 0)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(38, 13, 72)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    'Panel (top triangle)
                    pt1 = New Drawing.Point(xLoc, 0)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx / 2, heightPx / 4)
                    pt3 = New Drawing.Point(xLoc, heightPx / 4)
                    pt4 = New Drawing.Point(xLoc, 0)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    'brush.Color = Color.FromArgb(47, 27, 80)
                    brush.Color = Color.FromArgb(28, 3, 61)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    '
                    'Panel (top right triangle)
                    pt1 = New Drawing.Point(xLoc, heightPx / 2)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx, heightPx)
                    pt3 = New Drawing.Point(xLoc, heightPx)
                    pt4 = New Drawing.Point(xLoc, heightPx / 2)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                Case 7
                    pt1 = New Drawing.Point(xLoc, 0)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx, heightPx / 2)
                    pt3 = New Drawing.Point(xLoc, heightPx)
                    pt4 = New Drawing.Point(xLoc, 0)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    '                    '
                    'Panel (bottom triangle, left)
                    pt1 = New Drawing.Point(xLoc, heightPx)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx, heightPx / 2)
                    pt3 = New Drawing.Point(xLoc + segmentWidthPx, heightPx)
                    pt4 = New Drawing.Point(xLoc, heightPx)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(38, 13, 72)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                Case 8
                    pt1 = New Drawing.Point(xLoc, heightPx / 2)
                    pt2 = New Drawing.Point(xLoc + segmentWidthPx, 0)
                    pt3 = New Drawing.Point(xLoc + segmentWidthPx, heightPx)
                    pt4 = New Drawing.Point(xLoc, heightPx / 2)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)

                    'brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)

                    'Case 1, 3, 5
                    'brush.Color = Color.FromArgb(r, g, b)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)

                    'Case 2, 4
                    'brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    '
                    '
                    '                    '
                    'Panel (bottom triangle, right)
                    pt1 = New Drawing.Point(xLoc, heightPx / 2)
                    pt2 = New Drawing.Point(xLoc, 0)
                    pt3 = New Drawing.Point(xLoc + segmentWidthPx, 0)
                    pt4 = New Drawing.Point(xLoc, heightPx / 2)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(38, 13, 72)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
            End Select
            xLoc = xLoc + segmentWidthPx
        Next

    End Sub
    '
#End Region
    Public Function grfx_ChptBanner_pasteAsPNG(ByRef hf As Word.HeaderFooter, ByRef bmp As System.Drawing.Bitmap, widthPts As Single, heightPts As Single, Optional strShpName As String = "cp_pict_large") As Word.Shape
        Dim rng As Word.Range
        Dim oldPictWrapType As WdWrapTypeMerged
        Dim shp As Word.Shape
        Dim sect As Word.Section
        '
        shp = Nothing
        sect = hf.Range.Sections.Item(1)
        '
        'Get the current insert/paste picture option (inline, behind text etc). Change it
        'to inline htem change it back.. Tested 20210807 and it works
        oldPictWrapType = Me.glb_get_wrdApp.Options.PictureWrapType
        '
        Me.glb_get_wrdApp.Options.PictureWrapType = WdWrapTypeMerged.wdWrapMergeBehind
        '
        If hf.Exists Then
            rng = hf.Range
            '
            rng.Collapse(WdCollapseDirection.wdCollapseEnd)
            'rng.Move(WdUnits.wdParagraph, -1)
            '
            'Me.grfx_convert_ToIndexedColor(bmp, shp)
            Me.grfx_Convert_ToPNG(bmp)
            '
            Clipboard.SetImage(bmp)
            rng.Paste()

            'Generally rng.ShapeRange.Item(1) will give you the item that was pasted. The range is expanded to include the
            'Clipboard contents... But  is not working this way at this time (20210821). So we need to search for it in the
            'hf ShapeRange
            '
            rng = hf.Range
            For Each shp In rng.ShapeRange
                If shp.Name Like "Picture*" Then
                    shp.LockAnchor = True
                    shp.Name = strShpName
                    shp.Width = widthPts
                    shp.Height = heightPts
                    shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
                    shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
                    If sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait Then
                        shp.Left = 57.25
                        shp.Top = 464.65
                    Else
                        shp.Left = 449.0
                        shp.Top = 130.15

                    End If
                    Exit For
                End If
                shp = Nothing
            Next

        End If
        '
        Me.glb_get_wrdApp.Options.PictureWrapType = oldPictWrapType
        '
        Return shp
    End Function

    '
    Public Sub grfx_ChptBanner_pasteAsPNG(ByRef rng As Word.Range, ByRef bmp As System.Drawing.Bitmap, widthPts As Single, heightPts As Single, Optional strShpName As String = "chpt_Banner_Body")
        Dim oldPictWrapType As WdWrapTypeMerged
        'Dim shp As Word.Shape

        '****
        'Get the current insert/paste picture option (inline, behind text etc). Change it
        'to inline htem change it back.. Tested 20210807 and it works
        oldPictWrapType = Me.glb_get_wrdApp.Options.PictureWrapType
        '
        Me.glb_get_wrdApp.Options.PictureWrapType = WdWrapTypeMerged.wdWrapMergeBehind
        'Me.grfx_convert_ToIndexedColor(bmp)
        Me.grfx_Convert_ToPNG(bmp)

        Clipboard.SetImage(bmp)
        rng.Paste()
        'rng.PasteSpecial(,,,,WdPasteDataType.)
        '
        Me.glb_get_wrdApp.Options.PictureWrapType = oldPictWrapType
        Me.glb_get_wrdApp.ScreenRefresh()

        '
        '****
        '
        'shp = rng.ShapeRange.Item(1)
        'shp = rng.InlineShapes.Item(1)

        'shpInline = shp.ConvertToInlineShape()
        '
        'Now stretch the shape to the right size in points
        'shp.LockAspectRatio = False
        'shp.LockAnchor = True
        '
        'shp.Width = widthPts
        'shp.Height = heightPts
        'shp.Name = strShpName

        'shpInline.Width = width
        'shpInline.Height = height


        'Return shp
        '
    End Sub

    '
    ''' <summary>
    ''' This method will paste a graphic at the range rng. The width and height are specified in points.
    ''' The colour fill is set to the default, but it can be overidden if the rgb values are set. The method will
    ''' return the graphic as an inline shape.
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="width"></param>
    ''' <param name="height"></param>
    ''' <param name="r"></param>
    ''' <param name="g"></param>
    ''' <param name="b"></param>
    ''' <returns></returns>
    Public Function grfx_insertShape_behind(ByRef rng As Word.Range, width As Single, height As Single, Optional r As Integer = -1, Optional g As Integer = -1, Optional b As Integer = -1, Optional xDpi As Single = 300, Optional yDpi As Single = 300) As Word.Shape
        Dim bmp As Bitmap
        Dim myDoc As Word.Document
        Dim grfx As Graphics
        Dim brush As SolidBrush
        Dim shp As Word.Shape
        'Dim shpInline As Word.InlineShape
        Dim widthPx, heightPx As Single
        Dim objGlobals As New cGlobals()
        Dim oldPictWrapType As WdWrapTypeMerged
        Dim j, numSegments As Integer
        Dim segmentWidth, segmentWidthPts, xLoc As Single
        Dim pt1, pt2, pt3, pt4, pt5 As Drawing.PointF
        Dim trianglePts(4), rectPts(5) As Drawing.PointF
        '
        'Set arbitrary width and height in pixels (remember that pxiels and points are different).
        'We will stretch this shape to the correct point size later
        'widthPx = 600
        'heightPx = 200
        'xDpi = 150
        'yDpi = 150
        '
        segmentWidthPts = 79.4
        '
        widthPx = xDpi * width / 72
        heightPx = yDpi * height / 72
        '
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        myDoc = rng.Document
        '
        bmp = New Bitmap(widthPx, heightPx, Imaging.PixelFormat.Format24bppRgb)
        'bmp = New Bitmap(width, height, Imaging.PixelFormat.Format8bppIndexed)
        bmp.SetResolution(xDpi, yDpi)
        '
        grfx = Graphics.FromImage(bmp)
        'bmp = Me.grfx_convert_ToIndexedColor(bmp)
        'bmp.Palette.Entries(1) = Color.FromArgb(233, 233, 233)
        '
        'Set the default colour for the fill, but if any of the r,g,b values is negative, then
        'generate the colour from the r,g,b values
        brush = New SolidBrush(Me._fillColor)
        'brush = New SolidBrush(bmp.Palette.Entries(3))
        '
        If r >= 0 And g >= 0 And b >= 0 Then
            'brush = New SolidBrush(Me._fillColor)
            brush.Color = Color.FromArgb(r, g, b)
            'brush.Color = bmp.Palette.Entries(2)

            grfx.FillRectangle(brush, 0, 0, widthPx, heightPx)
        End If
        '
        'GoTo finis

        xLoc = 0
        numSegments = 5
        segmentWidth = widthPx / numSegments
        segmentWidthPts = width / numSegments
        MsgBox("Segment width = " + CStr(segmentWidthPts) + " pts")
        '
        For j = 1 To numSegments
            Select Case j
                Case 1
                    '
                    'Panel (Top)
                    pt1 = New Drawing.Point(xLoc, 0)
                    pt2 = New Drawing.Point(xLoc + segmentWidth, 0)
                    pt3 = New Drawing.Point(xLoc + segmentWidth, heightPx / 2)
                    pt4 = New Drawing.Point(xLoc, heightPx / 2)
                    pt5 = New Drawing.Point(xLoc, 0)

                    rectPts(1) = pt1
                    rectPts(2) = pt2
                    rectPts(3) = pt3
                    rectPts(4) = pt4
                    rectPts(5) = pt5

                    '
                    brush.Color = Color.FromArgb(47, 27, 80)
                    'brush.Color = bmp.Palette.Entries(3)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, rectPts)
                    '
                    '
                    pt1 = New Drawing.Point(xLoc, 0)
                    pt2 = New Drawing.Point(xLoc + segmentWidth / 2, heightPx / 4)
                    pt3 = New Drawing.Point(xLoc, heightPx / 2)
                    pt4 = New Drawing.Point(xLoc, 0)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(56, 37, 86)
                    'brush = New SolidBrush(bmp.Palette.Entries(2))
                    'brush.Color = bmp.Palette.Entries(2)
                    '
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)


                    'brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    pt1 = New Drawing.Point(xLoc, heightPx / 2)
                    pt2 = New Drawing.Point(xLoc + segmentWidth / 2, 3 * heightPx / 4)
                    pt3 = New Drawing.Point(xLoc, heightPx)
                    pt4 = New Drawing.Point(xLoc, heightPx / 2)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(47, 27, 80)
                    'brush.Color = bmp.Palette.Entries(3)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    pt1 = New Drawing.Point(xLoc + segmentWidth / 2, heightPx / 4)
                    pt2 = New Drawing.Point(xLoc + segmentWidth, heightPx / 2)
                    pt3 = New Drawing.Point(xLoc + segmentWidth / 2, 3 * heightPx / 4)
                    pt4 = New Drawing.Point(xLoc + segmentWidth / 2, heightPx / 4)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(38, 13, 72)
                    'brush.Color = bmp.Palette.Entries(4)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '

                    'GoTo finis
                    '
                    'Panel (bottom triangle, top)
                    pt1 = New Drawing.Point(xLoc + segmentWidth / 2, 3 * heightPx / 4)
                    pt2 = New Drawing.Point(xLoc + segmentWidth, heightPx / 2)
                    pt3 = New Drawing.Point(xLoc + 3 * segmentWidth / 2, 3 * heightPx / 4)
                    pt4 = New Drawing.Point(xLoc + segmentWidth / 2, 3 * heightPx / 4)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    '
                    'Panel (bottom triangle, left)
                    pt1 = New Drawing.Point(xLoc, heightPx)
                    pt2 = New Drawing.Point(xLoc + segmentWidth / 2, 3 * heightPx / 4)
                    pt3 = New Drawing.Point(xLoc + segmentWidth, heightPx)
                    pt4 = New Drawing.Point(xLoc, heightPx)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(38, 13, 72)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    '
                    'Panel (bottom triangle, right)
                    pt1 = New Drawing.Point(xLoc + segmentWidth, heightPx)
                    pt2 = New Drawing.Point(xLoc + 3 * segmentWidth / 2, 3 * heightPx / 4)
                    pt3 = New Drawing.Point(xLoc + 2 * segmentWidth, heightPx)
                    pt4 = New Drawing.Point(xLoc + segmentWidth, heightPx)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(47, 27, 80)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                Case 2
                    pt1 = New Drawing.Point(xLoc, heightPx / 2)
                    pt2 = New Drawing.Point(xLoc + segmentWidth, 0)
                    pt3 = New Drawing.Point(xLoc + segmentWidth, heightPx)
                    pt4 = New Drawing.Point(xLoc, heightPx / 2)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(38, 13, 72)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    'brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                Case 3
                    pt1 = New Drawing.Point(xLoc, 0)
                    pt2 = New Drawing.Point(xLoc + segmentWidth, heightPx / 2)
                    pt3 = New Drawing.Point(xLoc, heightPx)
                    pt4 = New Drawing.Point(xLoc, 0)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    '                    '
                    'Panel (bottom triangle, left)
                    pt1 = New Drawing.Point(xLoc, heightPx)
                    pt2 = New Drawing.Point(xLoc + segmentWidth, heightPx / 2)
                    pt3 = New Drawing.Point(xLoc + segmentWidth, heightPx)
                    pt4 = New Drawing.Point(xLoc, heightPx)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(38, 13, 72)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '

                Case 4
                    pt1 = New Drawing.Point(xLoc, heightPx / 2)
                    pt2 = New Drawing.Point(xLoc + segmentWidth, 0)
                    pt3 = New Drawing.Point(xLoc + segmentWidth, heightPx)
                    pt4 = New Drawing.Point(xLoc, heightPx / 2)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)

                    'brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)

                    'Case 1, 3, 5
                    'brush.Color = Color.FromArgb(r, g, b)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)

                    'Case 2, 4
                    'brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    '
                    '
                    '                    '
                    'Panel (bottom triangle, right)
                    pt1 = New Drawing.Point(xLoc, heightPx / 2)
                    pt2 = New Drawing.Point(xLoc, 0)
                    pt3 = New Drawing.Point(xLoc + segmentWidth, 0)
                    pt4 = New Drawing.Point(xLoc, heightPx / 2)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(38, 13, 72)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                Case 5
                    pt1 = New Drawing.Point(xLoc, 0)
                    pt2 = New Drawing.Point(xLoc + segmentWidth, heightPx / 2)
                    pt3 = New Drawing.Point(xLoc, heightPx)
                    pt4 = New Drawing.Point(xLoc, 0)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(38, 13, 72)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    'Panel (top triangle)
                    pt1 = New Drawing.Point(xLoc, 0)
                    pt2 = New Drawing.Point(xLoc + segmentWidth, 0)
                    pt3 = New Drawing.Point(xLoc + segmentWidth / 2, heightPx / 4)
                    pt4 = New Drawing.Point(xLoc, 0)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(47, 27, 80)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
                    '
                    '
                    'Panel (top right triangle)
                    pt1 = New Drawing.Point(xLoc + segmentWidth / 2, heightPx / 4)
                    pt2 = New Drawing.Point(xLoc + segmentWidth, heightPx / 4)
                    pt3 = New Drawing.Point(xLoc + segmentWidth, heightPx / 2)
                    pt4 = New Drawing.Point(xLoc + segmentWidth / 2, heightPx / 4)

                    trianglePts(1) = pt1
                    trianglePts(2) = pt2
                    trianglePts(3) = pt3
                    trianglePts(4) = pt4
                    '
                    brush.Color = Color.FromArgb(56, 37, 86)
                    'grfx.FillRectangle(brush, xLoc, 0, segmentWidth, heightPx)
                    grfx.FillPolygon(brush, trianglePts)
            End Select
            xLoc = xLoc + segmentWidth
        Next
        '
        'b'rush.Color = Color.FromArgb(255, 0, 0)

        'pt1 = New Drawing.Point(0, 0)
        'pt2 = New Drawing.Point(segmentWidth, heightPx / 2)
        'pt3 = New Drawing.Point(0, heightPx)
        '
        'curvePts(1) = pt1
        'curvePts(2) = pt2
        'curvePts(3) = pt3
        '
        'grfx.FillPolygon(brush, curvePts)
        'grfx.FillRectangle(brush, 0, 0, widthPx / 5, heightPx)
        '
finis:
        '****
        'Get the current insert/paste picture option (inline, behind text etc). Change it
        'to inline htem change it back.. Tested 20210807 and it works
        oldPictWrapType = objGlobals.glb_get_wrdApp.Options.PictureWrapType
        '
        objGlobals.glb_get_wrdApp.Options.PictureWrapType = WdWrapTypeMerged.wdWrapMergeBehind
        'Me.grfx_convert_ToIndexedColor(bmp)
        Me.grfx_Convert_ToPNG(bmp)

        Clipboard.SetImage(bmp)
        rng.Paste()
        'rng.PasteSpecial(,,,,WdPasteDataType.)
        '
        objGlobals.glb_get_wrdApp.Options.PictureWrapType = oldPictWrapType
        objGlobals.glb_get_wrdApp.ScreenRefresh()

        '
        '****
        '
        shp = rng.ShapeRange.Item(1)
        'shp = rng.InlineShapes.Item(1)

        'shpInline = shp.ConvertToInlineShape()
        '
        'Now stretch the shape to the right size in points
        shp.LockAspectRatio = False
        shp.LockAnchor = True

        'shpInline.Width = width
        'shpInline.Height = height


        Return shp
        '
    End Function


    '
    Public Sub test()
        Dim objGlobals As New cGlobals()
        Dim bmp As Bitmap
        'Dim g As Graphics
        Dim rect As Drawing.Rectangle
        Dim myDoc As Word.Document
        '
        bmp = New Bitmap(600, 250)
        Clipboard.SetImage(bmp)
        rect = New Drawing.Rectangle()
        '
        myDoc = objGlobals.glb_get_wrdActiveDoc
        myDoc.Shapes.AddShape(MsoShapeType.msoPicture, 20, 20, 40, 40, objGlobals.glb_get_wrdSelRng)

    End Sub
    '
    Public Sub test2()
        Dim objGlobals As New cGlobals()
        Dim bmp As Bitmap
        'Dim g As Graphics
        Dim rect As Drawing.Rectangle
        Dim myDoc As Word.Document
        '
        bmp = New Bitmap(397, 157)
        bmp.SetResolution(150, 150)
        'bmp.c

        rect = New Drawing.Rectangle()
        '
        myDoc = objGlobals.glb_get_wrdActiveDoc
        myDoc.Shapes.AddShape(MsoShapeType.msoPicture, 20, 20, 40, 40, objGlobals.glb_get_wrdSelRng)

    End Sub

End Class

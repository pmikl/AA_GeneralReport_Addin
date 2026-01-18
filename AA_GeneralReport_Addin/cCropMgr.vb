Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word
Imports System.Windows.Forms
'
'rev 01.00  20250830
'
Public Class cCropMgr
    Inherits cGlobals
    Public strScratchFilePath As String
    '
    Public cropRect As Word.Shape
    Public imageToBeCropped As Word.Shape
    Public shpToBeFilled As Word.Shape          'eg Back Panel

    Public objShp_NewPic As cShapeMgr
    '
    Public Sub New()
        MyBase.New()
        Me.strScratchFilePath = My.Computer.FileSystem.SpecialDirectories.MyPictures + "\aac_scratch_file.jpg"
        Me.cropRect = Nothing
        Me.objShp_NewPic = New cShapeMgr()
        Me.shpToBeFilled = Nothing
    End Sub
    '
    ''' <summary>
    ''' from imageCrop in frm_pictControl2, line 133
    ''' </summary>
    Public Sub crp_crop_CropImage(cropLeft As Single, cropRight As Single, cropTop As Single, cropBottom As Single)
        Dim scaleFactor_H, scaleFactor_W As Single
        '
        'shp.PictureFormat.CropLeft = Me.cropRect.Left * 1 / Me.objSrcShape.scaleFactor_H
        'shp.PictureFormat.CropLeft = Me.cropRect.Left * Me.objSrcShape.scaleFactor_H

        'shp.PictureFormat.CropLeft = deltaLeft
        '
        'j = objSrcShape.scaleFactor_H
        'k = objSrcShape.scaleFactor_W
        'shp.PictureFormat.CropLeft = Me.cropLeft * 1 / 0.8
        'shp.PictureFormat.CropRight = Me.cropRight * 1 / 0.8
        scaleFactor_H = Me.objShp_NewPic.scaleFactor_H
        scaleFactor_W = Me.objShp_NewPic.scaleFactor_W

        '
        Me.imageToBeCropped.PictureFormat.CropLeft = cropLeft * scaleFactor_W
        Me.imageToBeCropped.PictureFormat.CropRight = cropRight * scaleFactor_W
        Me.imageToBeCropped.PictureFormat.CropTop = cropTop * scaleFactor_H
        Me.imageToBeCropped.PictureFormat.CropBottom = cropBottom * scaleFactor_H
        '
        'shpToBeCropped.PictureFormat.CropLeft = cropLeft * objSrcShape.scaleFactor_W
        'shpToBeCropped.PictureFormat.CropRight = cropRight * objSrcShape.scaleFactor_W
        'shpToBeCropped.PictureFormat.CropTop = cropTop * objSrcShape.scaleFactor_H
        'shpToBeCropped.PictureFormat.CropBottom = cropBottom * objSrcShape.scaleFactor_H
        '

    End Sub
    '
    Public Sub crp_fill_ShapeWithImage()
        Dim img As System.Drawing.Image
        'NewPic.Select()
        'sel = Globals.ThisDocument.Application.Selection
        'sel.CopyAsPicture()
        'NewPic.Delete()
        '
        '***** It works
        img = Clipboard.GetImage()
        img.Save(Me.strScratchFilePath)
        Me.shpToBeFilled.Fill.UserPicture(Me.strScratchFilePath)
        Me.shpToBeFilled.Line.Visible = False
        'objShpMgr.shp.Fill.UserPicture(Me.strScratchFilePath)

    End Sub
    '
    '
    ''' <summary>
    ''' This method will modify rngPasted to the actual range in the header (hf) that the image is to
    ''' be pasted.. The method will return true if all is OK, false if there was a fault
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <param name="rngPasted"></param>
    ''' <returns></returns>
    Public Function crp_get_headerInsertRange(ByRef hf As Word.HeaderFooter, ByRef rngPasted As Word.Range) As Boolean
        Dim OK As Boolean
        '
        OK = True
        '
        Try
            If hf.Range.Tables.Count > 0 Then
                '***
                'para = rngPasted.Paragraphs.Add()
                'paraLast = hf.Range.Paragraphs.Last
                '
                'rngPasted = para.Range
                'rngPasted.Collapse(WdCollapseDirection.wdCollapseStart)
                '
                '***
                'Must use paste at selection, anyhting else such as 
                'rngPasted.PasteSpecial(, , WdOLEPlacement.wdInLine, , WdPasteDataType.wdPasteBitmap)
                'will thwo an inexplicable fault

                'rngPasted.Select()
                'Globals.ThisDocument.Application.Selection.Paste()
                '
                'rngPasted = Globals.ThisDocument.Application.Selection.Range
                '
                'rngPasted = hf.Range

            Else
                'rngPasted.Select()
                '
                'Globals.ThisDocument.Application.Selection.Paste()
                'rngPasted = Globals.ThisDocument.Application.Selection.Range
                '
                'rngPasted = hf.Range
            End If

        Catch ex As Exception
            OK = False
        End Try
        '
        Return OK

    End Function

    ''' <summary>
    ''' This method will place the source image on the current Page. It is placed either from the clipboard
    ''' or from a file at strFilePath
    ''' </summary>
    ''' <param name="objSrcImgShpMgr"></param>
    ''' <param name="sect"></param>
    ''' <param name="strPlaceFrom"></param>
    ''' <param name="strFilePath"></param>
    ''' <returns></returns>
    Public Function crp_SrcImage_Place(ByRef objSrcImgShpMgr As cShapeMgr, ByRef sect As Word.Section, Optional ByRef strPlaceFrom As String = "clipboard", Optional strFilePath As String = "") As Word.Shape
        'Dim dlg_Picture As Word.Dialog
        Dim srcImage, newPic As Word.Shape
        Dim newPicInline As Word.InlineShape
        'Dim objShp_NewPic As cShapeMgr
        Dim rngSrc, rngPasted As Word.Range
        Dim objSectMgr As New cSectionMgr()
        Dim hf As HeaderFooter
        Dim h, v As Single
        Dim paraLast As Word.Paragraph
        '
        srcImage = Nothing
        rngSrc = objSrcImgShpMgr.anchor
        hf = objSrcImgShpMgr.hf
        '
        paraLast = hf.Range.Paragraphs.Last
        rngPasted = paraLast.Range
        rngPasted.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        Select Case strPlaceFrom
            Case "clipboard"
                Me.crp_get_headerInsertRange(hf, rngPasted)
                rngPasted.Select()
                glb_get_wrdApp.Selection.Paste()
                rngPasted = glb_get_wrdApp.Selection.Range
                rngPasted = hf.Range

            Case "file"
                Me.crp_get_headerInsertRange(hf, rngPasted)
                newPicInline = sect.Range.InlineShapes.AddPicture(strFilePath,,, rngPasted)
                rngPasted = hf.Range
        End Select
        '
        '
        newPic = Me.crp_get_ImageAsShape(hf)
        '
        Me.objShp_NewPic = New cShapeMgr()
        Me.objShp_NewPic.InitShape(newPic, objSrcImgShpMgr.hf)
        '
        Call Me.setImagePageProperties(newPic)          'Page positional properties
        '
        Call newPic.ScaleWidth(1.0#, MsoTriState.msoTrue)  'sort of resets the record of width & height?
        Call newPic.ScaleHeight(1.0#, MsoTriState.msoTrue)
        '
        Me.sct_fit_ShapeToPage(Me.objShp_NewPic, sect)
        '
        h = Me.objShp_NewPic.scaleFactor_H
        v = Me.objShp_NewPic.scaleFactor_W
        '
        'The following is now done in Me.setImageScale2
        '
        Call Me.crp_Panel_SetSrcImagePosition(newPic, sect)
        Me.imageToBeCropped = newPic
        '
        '
        Return newPic
    End Function
    '
    Public Sub crp_Panel_SetSrcImagePosition(ByRef shp As Word.Shape, ByRef sect As Word.Section)
        'This positions the target image in the centre of the page
        Dim AspectRatio_Page As Single
        '
        AspectRatio_Page = sect.PageSetup.PageHeight / sect.PageSetup.PageWidth
        '
        If AspectRatio_Page > 1 Then  'Page is Portrait
            shp.Left = (595 - shp.Width) / 2
            'shp.top = (841 / 2) - shp.height / 2
            shp.Top = (842 - shp.Height) / 2                     ' Sit at the bottom of the page - Allow 20pts for footer
        ElseIf AspectRatio_Page <= 1 Then 'Page is Landscape
            shp.Left = (841 / 2) - shp.Width / 2
            shp.Top = (595 / 2) - shp.Height / 2
        End If
        '
    End Sub

    '
    Public Sub setImagePageProperties(ByRef shp As Word.Shape)
        'Call shp.ScaleWidth(1.0#, MsoTriState.msoTrue)  'sort of resets the record of width & height?
        'Call shp.ScaleHeight(1.0#, MsoTriState.msoTrue)
        '
        shp.LockAspectRatio = MsoTriState.msoTrue
        shp.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
        shp.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
        shp.LockAnchor = True
        shp.WrapFormat.AllowOverlap = True
        shp.WrapFormat.Type = WdWrapType.wdWrapNone
        '
    End Sub
    '
    Public Sub crp_setView_toPrintView()
        Dim myDoc As Word.Document
        '
        myDoc = glb_get_wrdApp.ActiveDocument
        myDoc.ActiveWindow.View.SeekView = WdSeekView.wdSeekMainDocument
        myDoc.ActiveWindow.ActivePane.View.Type = WdViewType.wdPrintView
        '
    End Sub

    '
    Public Sub crp_setView_toHeader()
        'This method will set the view to the current Header
        Dim wordApp As Word.Application
        Dim windowPane As Word.Pane

        wordApp = glb_get_wrdApp()
        windowPane = wordApp.ActiveWindow.ActivePane
        If wordApp.ActiveWindow.View.SplitSpecial <> Word.WdSpecialPane.wdPaneNone Then
            wordApp.ActiveWindow.Panes(2).Close()
        End If
        If windowPane.View.Type <> Word.WdViewType.wdPrintView Then windowPane.View.Type = Word.WdViewType.wdPrintView
        windowPane.View.SeekView = Word.WdSeekView.wdSeekCurrentPageHeader
        '
    End Sub
    '
    ''' <summary>
    ''' This method will retrieve the (one) image to crop in the hf (HeaderFooter). If it is an inline shape, it
    ''' will convert it to a Floating SHape.. If not, it will just find and retuen it... Typically
    ''' the range 
    ''' </summary>
    ''' <param name="hf"></param>
    ''' <returns></returns>
    Public Function crp_get_ImageAsShape(ByRef hf As Word.HeaderFooter) As Word.Shape
        Dim newPic As Word.Shape
        Dim rng As Word.Range
        Dim strShapeName As String
        Dim j As Integer
        '
        rng = hf.Range
        newPic = Nothing
        '
        Try
            If rng.InlineShapes.Count > 0 Then
                'Get and convert the inline Shape
                newPic = rng.InlineShapes.Item(rng.InlineShapes.Count).ConvertToShape()
                newPic.ZOrder(MsoZOrderCmd.msoSendBehindText)
                newPic.ZOrder(MsoZOrderCmd.msoBringToFront)
                'newPic.Select()
                '
            Else
                'Get the floating Shape
                If rng.ShapeRange.Count > 0 Then
                    For j = 1 To rng.ShapeRange.Count
                        newPic = rng.ShapeRange.Item(j)
                        strShapeName = newPic.Name
                        If newPic.Name Like "Picture*" Then Exit For
                    Next
                    '
                End If
            End If

        Catch ex As Exception
            newPic = Nothing
        End Try
        '
        Return newPic
    End Function
    '
    ''' <summary>
    ''' This mthod will delete the cropping rectangle
    ''' </summary>
    Public Sub crp_delete_cropRectangle()
        Try
            Me.cropRect.Delete()
            Me.cropRect = Nothing
        Catch ex As Exception

        End Try
    End Sub
    '
    ''' <summary>
    ''' This method will delete the New (cropped) image
    ''' </summary>
    Public Sub crp_delete_SrcImage()
        Try
            Me.objShp_NewPic.shp.Delete()
        Catch ex As Exception

        End Try
    End Sub
    '
    ''' <summary>
    ''' This method will delete both the cropping rectangle and the source image
    ''' </summary>
    Public Sub crp_delete_CropRect_and_SrcImage()
        Me.crp_delete_cropRectangle()
        Me.crp_delete_SrcImage()
    End Sub
    '
    ''' <summary>
    ''' ObjSrcShp_tobeFilled is the source shape from which the clipping mask gets its shape/aspect ratio. This is the shape that will
    ''' eventually be filled with a new image. The object shp_NewPict_tobeClipped is the
    ''' new picture that the clipping mask is placed over to get a new pciture that will insert into the source shape
    ''' without distortion
    ''' </summary>
    ''' <param name="sect"></param>
    ''' <param name="objSrcShp_tobeFilled"></param>
    ''' <param name="shp_ImageToBeClipped"></param>
    ''' <returns></returns>
    Public Function crp_insert_CropRect(ByRef sect As Section, ByRef objSrcShp_tobeFilled As cShapeMgr, ByRef shp_ImageToBeClipped As Word.Shape) As Word.Shape
        'Dim aspectRatio_CropRect, aspectRatio_ImageToBeClipped As Single
        'Dim dummy_h, dummy_w, scaleFactor_h As Single
        Dim clipHeight, clipWidth As Single
        Dim strShpAspectRatio_CropRect As String
        Dim strImageToBeClipped As String
        Dim hf As HeaderFooter
        Dim rng As Word.Range
        Dim shpRect As Word.Shape
        Dim irror As Boolean
        Dim objCropRect As New cCropRectMgr()
        Dim lstOfCropRectDimensions As New Collection()
        '
        Me.shpToBeFilled = objSrcShp_tobeFilled.shp
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
        '
        'Get Cropping Rectangle dimensions in pts.... Remember aspect ratio is h/w
        lstOfCropRectDimensions = objCropRect.rct_(objSrcShp_tobeFilled, shp_ImageToBeClipped)
        clipWidth = CSng(lstOfCropRectDimensions("clipWidth"))
        clipHeight = CSng(lstOfCropRectDimensions("clipHeight"))
        '
        '
        'Set inserCtropRect = ActiveDocument.Shapes.AddShape(msoShapeRectangle, shp.left, shp.top, objShp.width, objShp.height, shp.Anchor)
        shpRect = hf.Shapes.AddShape(MsoAutoShapeType.msoShapeRectangle, shp_ImageToBeClipped.Left, shp_ImageToBeClipped.Top, clipWidth, clipHeight, rng)
        shpRect.LockAspectRatio = MsoTriState.msoTrue
        shpRect.RelativeHorizontalPosition = WdRelativeHorizontalPosition.wdRelativeHorizontalPositionPage
        shpRect.RelativeVerticalPosition = WdRelativeVerticalPosition.wdRelativeVerticalPositionPage
        shpRect.LockAnchor = True
        shpRect.WrapFormat.AllowOverlap = True
        shpRect.WrapFormat.Type = WdWrapType.wdWrapNone
        '
        'insertCropRect.left = shp.left
        'insertCropRect.top = shp.top
        shpRect.Fill.Transparency = 0.5
        shpRect.Fill.BackColor.RGB = RGB(0, 128, 128)
        '
        'Make sure the cropping rectangle is aligned with the image to be clipped
        shpRect.Top = shp_ImageToBeClipped.Top
        shpRect.Left = shp_ImageToBeClipped.Left
        '
        Me.cropRect = shpRect
        '
        Return shpRect
    End Function
    '
    '
    ''' <summary>
    ''' ************  NEEDS WORK  ***********************************************************************
    ''' This function will resize a shape (shp) while keeping its aspect ratio and sit it centrally
    ''' in the page
    ''' </summary>
    ''' <param name="objShpMgr"></param>
    ''' <param name="sect"></param>
    ''' <returns></returns>
    Public Function sct_fit_ShapeToPage(ByRef objShpMgr As cShapeMgr, ByRef sect As Word.Section) As Boolean
        Dim objWCAGMgr As New cWCAGMgr()
        Dim shp As Word.Shape
        Dim w, h, width_old, height_old As Single
        Dim lstOfScaleFactors As New Collection()
        Dim scaleFactor_w, scaleFactor_h, shpAspectRatio As Single
        Dim dummy_h, dummy_w As Single
        Dim strShpAspectRatio As String
        Dim i As Integer
        Dim irror As Boolean
        '
        shp = objShpMgr.shp
        shp.LockAspectRatio = True
        shpAspectRatio = shp.Height / shp.Width
        dummy_h = shp.Height
        dummy_w = shp.Width
        w = shp.Width
        h = shp.Height
        '
        width_old = shp.Width
        height_old = shp.Height
        '
        strShpAspectRatio = ""
        irror = False

        If shpAspectRatio > 1 Then strShpAspectRatio = "shp_is_portrait"
        If shpAspectRatio = 1 Then strShpAspectRatio = "shp_is_square"
        If shpAspectRatio < 1 Then strShpAspectRatio = "shp_is_landscape"
        '
        If sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait Then
            '*** All scenarios for porrait page verified 20231106
            Select Case strShpAspectRatio
                Case "shp_is_portrait"
                    dummy_w = dummy_w * 0.8
                    dummy_h = dummy_w * shpAspectRatio

                    If dummy_h < 0.7 * sect.PageSetup.PageHeight Then
                        'all is OK.. SO we can operate on the real shape
                        shp.Width = sect.PageSetup.PageWidth * 0.8
                        objShpMgr.scaleFactor_W = height_old / shp.Height
                        objShpMgr.scaleFactor_H = width_old / shp.Width
                    Else
                        'Scale it until the height fits
                        scaleFactor_h = 0.8
                        '
                        For i = 1 To 16
                            dummy_h = sect.PageSetup.PageHeight * scaleFactor_h
                            dummy_w = dummy_h / shpAspectRatio
                            '
                            If dummy_h < 0.7 * sect.PageSetup.PageHeight Then
                                shp.Height = dummy_h
                                objShpMgr.scaleFactor_W = height_old / shp.Height
                                objShpMgr.scaleFactor_H = width_old / shp.Width
                                Exit For
                            End If
                            '
                            scaleFactor_h = scaleFactor_h - 0.05
                            If scaleFactor_h < 0.1 Then
                                objShpMgr.scaleFactor_W = 1
                                objShpMgr.scaleFactor_H = 1

                                irror = True
                                Exit For
                            End If
                        Next

                    End If

                Case "shp_is_square"
                    shp.Width = sect.PageSetup.PageWidth * 0.8
                    scaleFactor_h = 0.8
                    scaleFactor_w = 0.8

                Case "shp_is_landscape"
                    shp.Width = sect.PageSetup.PageWidth * 0.8
                    objShpMgr.scaleFactor_W = height_old / shp.Height
                    objShpMgr.scaleFactor_H = width_old / shp.Width

                    scaleFactor_h = 0.8
                    scaleFactor_w = 0.8
                    '
                Case Else
            End Select

        End If

        If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then
            Select Case strShpAspectRatio
                Case "shp_is_portrait"
                    shp.Height = shp.Height * 0.8
                    scaleFactor_h = 0.8
                    scaleFactor_w = 0.8
                    '
                    '*** Verified as correct for portrait image on a landscape page
                    objShpMgr.scaleFactor_W = height_old / shp.Height
                    objShpMgr.scaleFactor_H = width_old / shp.Width


                Case "shp_is_square"
                    shp.Height = shp.Height * 0.8
                    scaleFactor_h = 0.8
                    scaleFactor_w = 0.8
                    '
                    '*** Verified as correct for portrait image on a landscape page
                    '*** Is this OK for square ?? Put in 20240105
                    '
                    objShpMgr.scaleFactor_W = height_old / shp.Height
                    objShpMgr.scaleFactor_H = width_old / shp.Width

                Case "shp_is_landscape"
                    dummy_w = dummy_w * 0.8
                    dummy_h = dummy_w * shpAspectRatio

                    If dummy_h < 0.7 * sect.PageSetup.PageHeight Then
                        'all is OK.. SO we can operate on the real shape
                        shp.Width = sect.PageSetup.PageWidth * 0.8
                        objShpMgr.scaleFactor_W = height_old / shp.Height
                        objShpMgr.scaleFactor_H = width_old / shp.Width
                    Else
                        'Scale it until the height fits
                        scaleFactor_h = 0.8
                        '
                        For i = 1 To 16
                            dummy_h = sect.PageSetup.PageHeight * scaleFactor_h
                            dummy_w = dummy_h / shpAspectRatio
                            '
                            If dummy_h < 0.7 * sect.PageSetup.PageHeight Then
                                shp.Height = dummy_h
                                objShpMgr.scaleFactor_W = height_old / shp.Height
                                objShpMgr.scaleFactor_H = width_old / shp.Width
                                Exit For
                            End If
                            '
                            scaleFactor_h = scaleFactor_h - 0.05
                            If scaleFactor_h < 0.1 Then
                                objShpMgr.scaleFactor_W = 1
                                objShpMgr.scaleFactor_H = 1

                                irror = True
                                Exit For
                            End If
                        Next

                    End If

                Case "xshp_is_landscape"
                    dummy_h = dummy_h * 0.8
                    dummy_w = dummy_h / shpAspectRatio
                    irror = False
                    If dummy_w < 0.7 * sect.PageSetup.PageWidth Then
                        'all is OK.. SO we can operate on the real shape
                        shp.Height = shp.Height * 0.8
                        scaleFactor_h = 0.8
                        scaleFactor_w = 0.8
                    Else
                        'Scale it until the height fits
                        scaleFactor_w = 0.8
                        '
                        For i = 1 To 16
                            dummy_w = w * scaleFactor_w
                            dummy_h = dummy_w / shpAspectRatio
                            '
                            If dummy_w < 0.7 * sect.PageSetup.PageWidth Then
                                shp.Height = dummy_h
                                shp.Width = dummy_w
                                objShpMgr.scaleFactor_W = height_old / shp.Height
                                objShpMgr.scaleFactor_H = width_old / shp.Width

                                Exit For
                            End If
                            '
                            scaleFactor_h = scaleFactor_h - 0.05
                            If scaleFactor_h < 0.1 Then
                                irror = True
                                Exit For
                            End If
                        Next

                    End If

            End Select

        End If
        '
        If irror = True Then
            shp.Top = (sect.PageSetup.PageHeight - shp.Height) / 2
            shp.Left = (sect.PageSetup.PageWidth - shp.Width) / 2
            '
            objShpMgr.scaleFactor_W = height_old / shp.Height
            objShpMgr.scaleFactor_H = width_old / shp.Height
            '
            shp.ZOrder(MsoZOrderCmd.msoSendBehindText)
            shp.ZOrder(MsoZOrderCmd.msoSendToBack)
            '
            'For WCAG purposes
            'Set Decorative property
            objWCAGMgr.wcag_set_decorative(shp, True)
            '
        End If
        '
        Return irror
        '
    End Function
    '
End Class

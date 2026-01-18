Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.Drawing
Public Class frm_pictControl3
    Public objCropRect As Word.Shape
    Public objSrcImageToBeCropped As Word.Shape
    Public objSrcShape As cShapeMgr
    Public objBackPanel As Word.Shape
    Public objCropMgr As cCropMgr
    '
    Public hf As Word.HeaderFooter
    '
    Public cropRect_OriginalHeight As Single
    Public cropRect_OriginalWidth As Single
    Public cropRect_OriginalAspectRatio As Single           'Original aspect ratio h/w
    '
    Public crop_Delta_Left, crop_Delta_Right, crop_Delta_Top, crop_Delta_Bottom As Single
    '
    '
    Public Sub New(ByRef objCropMgr As cCropMgr, ByRef objImageToBeCropped As Word.Shape, ByRef objBackPanel As cShapeMgr)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        '
        Me.objCropMgr = objCropMgr
        Me.objCropRect = objCropMgr.cropRect
        Me.objSrcImageToBeCropped = objImageToBeCropped
        Me.objBackPanel = objBackPanel.shp
        Me.hf = objBackPanel.hf
        '
        Me.cropRect_OriginalWidth = objCropRect.Width
        Me.cropRect_OriginalHeight = objCropRect.Height
        Me.cropRect_OriginalAspectRatio = Me.cropRect_OriginalHeight / Me.cropRect_OriginalWidth
        '
        objSrcShape = New cShapeMgr()
        objSrcShape.InitShape(Me.objSrcImageToBeCropped, objBackPanel.hf)
        '
    End Sub
    '
    ''' <summary>
    ''' from imageCrop in frm_pictControl2, line 133
    ''' </summary>
    Public Sub frm_crop_ImageToBeCropped()
        Dim scaleHeight, scaleWidth As Single
        '
        scaleHeight = objSrcShape.scaleFactor_H
        scaleWidth = objSrcShape.scaleFactor_W

        'Me.objCropMgr.crp_crop_ImageToBeCropped(objSrcImageToBeCropped, Me.crop_Delta_Left,
        'Me.crop_Delta_Bottom, objSrcShape.scaleFactor_W, objSrcShape.scaleFactor_H)
        '
        'shp.PictureFormat.CropLeft = Me.cropRect.Left * 1 / Me.objSrcShape.scaleFactor_H
        'shp.PictureFormat.CropLeft = Me.cropRect.Left * Me.objSrcShape.scaleFactor_H

        'shp.PictureFormat.CropLeft = deltaLeft
        '
        'shp.PictureFormat.CropLeft = Me.cropLeft * 1 / 0.8
        'shp.PictureFormat.CropRight = Me.cropRight * 1 / 0.8
        '
        objSrcImageToBeCropped.PictureFormat.CropLeft = Me.crop_Delta_Left * scaleWidth
        objSrcImageToBeCropped.PictureFormat.CropRight = Me.crop_Delta_Right * scaleWidth
        objSrcImageToBeCropped.PictureFormat.CropTop = Me.crop_Delta_Top * scaleHeight
        objSrcImageToBeCropped.PictureFormat.CropBottom = Me.crop_Delta_Bottom * scaleHeight
        '
    End Sub

    Private Sub frm_pictControl3_Activated(sender As Object, e As EventArgs) Handles Me.Activated
        '
        Me.frm_get_CroppingRectangle_Width_and_Height()
        Me.frm_get_SrcImage_Width_and_Height()
        '
        Me.frm_get_Cropping_DeltaInformation()
        '
        Me.scrl_reSize.Minimum = 0
        Me.scrl_reSize.Maximum = 1000
        Me.scrl_reSize.Value = Me.scrl_reSize.Maximum
        '
        Me.scrl_moveLeftRight.Maximum = 1000
        Me.scrl_moveLeftRight.Minimum = 0
        Me.scrl_moveLeftRight.Value = 0
        '
        Me.scrl_moveUpDown.Maximum = 1000
        Me.scrl_moveUpDown.Minimum = 0
        Me.scrl_moveUpDown.Value = 0
        '
    End Sub

    Private Sub scrl_moveLeftRight_Scroll(sender As Object, e As System.Windows.Forms.ScrollEventArgs) Handles scrl_moveLeftRight.Scroll
        '
        Dim setting_Current As Single
        Dim proportion, setting_Max As Single
        Dim maxRight As Single
        '
        setting_Current = CSng(Me.scrl_moveLeftRight.Value)
        setting_Max = Me.scrl_moveLeftRight.Maximum
        proportion = setting_Current / setting_Max
        '
        maxRight = objSrcImageToBeCropped.Width - Me.objCropRect.Width
        '
        Try
            Me.objCropRect.Left = Me.objSrcImageToBeCropped.Left + Me.objSrcImageToBeCropped.Width * proportion
            'Me.cropRect.Top = Me.shp_ImageToBeClipped.Top + Me.shp_ImageToBeClipped.Height * proportion
            Me.frm_get_Cropping_DeltaInformation()
            'Me.frm_getCropSettings(True)
            'Me.frm_get_CroppingRectangle_Width_and_Height()
            'Me.frm_get_SrcImage_Width_and_Height()
            '
        Catch ex As Exception

        End Try
        '

    End Sub

    Private Sub scrl_reSize_Scroll(sender As Object, e As System.Windows.Forms.ScrollEventArgs) Handles scrl_reSize.Scroll
        Dim setting_Current As Single
        Dim proportion, setting_Max As Single
        '
        setting_Current = CSng(scrl_reSize.Value)
        setting_Max = scrl_reSize.Maximum
        proportion = setting_Current / setting_Max
        '
        Me.objCropRect.Width = Me.cropRect_OriginalWidth * proportion
        '
        If proportion >= 0.99 Then
            Me.objCropRect.Width = Me.cropRect_OriginalWidth
        End If
        '
        Me.frm_get_Cropping_DeltaInformation()
        Me.frm_get_CroppingRectangle_Width_and_Height()
        'Me.frm_getCropSettings(True)
        '
        'Me.frm_get_CroppingRectangle_Width_and_Height()
        'Me.frm_set_Cropping_DeltaInformation()
        'Me.frm_get_SrcImage_Width_and_Height()
        'Me.txtBox_scrollVerticalProportion.Text = proportion.ToString("f2")
        '

        'Globals.ThisDocument.Application.ScreenRefresh()
        Globals.ThisAddIn.Application.ScreenRefresh()
        '
    End Sub

    Private Sub scrl_moveUpDown_Scroll(sender As Object, e As System.Windows.Forms.ScrollEventArgs) Handles scrl_moveUpDown.Scroll
        '
        Dim setting_Current As Single
        Dim proportion, setting_Max As Single
        '
        setting_Current = CSng(Me.scrl_moveUpDown.Value)
        setting_Max = Me.scrl_moveUpDown.Maximum
        proportion = setting_Current / setting_Max
        '
        Try
            Me.objCropRect.Top = Me.objSrcImageToBeCropped.Top + Me.objSrcImageToBeCropped.Height * proportion
            Me.frm_get_Cropping_DeltaInformation()
            'Me.frm_getCropSettings(True)
            'Me.frm_get_CroppingRectangle_Width_and_Height()
            'Me.frm_get_SrcImage_Width_and_Height()
            '
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_finishPicture_Click(sender As Object, e As EventArgs) Handles btn_finishPicture.Click
        Dim objWCAGMgr As New cWCAGMgr()
        'Dim img As Image
        'Dim saveImage As System.Drawing.Bitmap
        'Dim strFilePath As String
        'Dim shp As Word.Shape
        'Dim rng As Word.Range
        'Dim scaleHeight, scaleWidth As Single
        '
        '
        Me.objCropMgr.crp_crop_CropImage(Me.crop_Delta_Left, Me.crop_Delta_Right, Me.crop_Delta_Top, Me.crop_Delta_Bottom)


        '
        'Now call the cropping dialog..After exit Me.shp_ImageToBeClipped has been cropped
        'Me.frm_crop_ImageToBeCropped()
        'Call imageCrop(Me.shp_ImageToBeClipped, Me.shp_ImageToBeClipped_as_cShapeMgr)
        '
        '
        '
        '
        'rng = Me.srcShape.
        'My.Computer.FileSystem.WriteAllBytes(strFilePath, shp.str)
        'shp = Me.srcShape
        'shp.can
        '

        'saveImage = New Bitmap(img)

        '*** This procedure sets the shape parameters as determined
        '*** from the Long name in InitShape
        'Call Me.shp_toBeFilled.SetShapeParameters(Me.shp_ImageToBeClipped)
        '
        'If Me.shp_ImageToBeClipped.Name = "aac_BackColour" Then
        'Me.shp_ImageToBeClipped.ZOrder(MsoZOrderCmd.msoSendBehindText)
        'Me.shp_ImageToBeClipped.ZOrder(MsoZOrderCmd.msoSendToBack)
        'End If
        '
        'For WCAG purposes
        'Set Decorative property
        'objWCAGMgr.wcag_set_decorative(Me.shp_ImageToBeClipped, True)

        'Call CH_InsImg.ShuffleShapes(LayoutName)
        'Me.cropRect.Delete()
        Me.Close()
        'Me.Dispose()
    End Sub
    '
    '
    Public Sub frm_get_CroppingRectangle_Width_and_Height()
        Dim aspectRatio As Single
        '
        Me.txtBox_cropRectWidth.Text = Me.objCropRect.Width.ToString("f3")
        Me.txtBox_cropRectHeight.Text = Me.objCropRect.Height.ToString("f3")
        '
        aspectRatio = Me.objCropRect.Height / Me.objCropRect.Width
        Me.txtBox_AspectRatio_CropRect.Text = aspectRatio.ToString("f3")
        '
    End Sub
    '
    Public Sub frm_get_SrcImage_Width_and_Height()
        '
        Me.txtBox_srcImageWidth.Text = Me.objSrcImageToBeCropped.Width.ToString("f2")
        Me.txtBox_srcImageHeight.Text = Me.objSrcImageToBeCropped.Height.ToString("f2")
        Me.txtBox_AspectRatio_SrcImg.Text = (Me.objSrcImageToBeCropped.Height / Me.objSrcImageToBeCropped.Width).ToString("f2")
        '
    End Sub
    '
    '
    Public Sub frm_get_Cropping_DeltaInformation()
        Me.crop_Delta_Left = Me.objCropRect.Left - Me.objSrcImageToBeCropped.Left
        Me.crop_Delta_Right = Me.objSrcImageToBeCropped.Width - Me.crop_Delta_Left - Me.objCropRect.Width
        Me.crop_Delta_Top = Me.objCropRect.Top - Me.objSrcImageToBeCropped.Top
        Me.crop_Delta_Bottom = Me.objSrcImageToBeCropped.Height - Me.crop_Delta_Top - Me.objCropRect.Height
        '
        Me.txtBox_Delta_CropLeft.Text = Me.crop_Delta_Left.ToString("f2")
        Me.txtBox_Delta_CropRight.Text = Me.crop_Delta_Right.ToString("f2")
        Me.txBox_DeltaCropTop.Text = Me.crop_Delta_Top.ToString("f2")
        Me.txtBox_DeltaCropBottom.Text = Me.crop_Delta_Bottom.ToString("f2")
        '
    End Sub
    '



End Class
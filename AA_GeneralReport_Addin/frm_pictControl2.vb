Public Class frm_pictControl2
    Public cropRect As Word.Shape                           'Cropping Rectangle
    Public cropRect_OriginalHeight As Single
    Public cropRect_OriginalWidth As Single
    Public cropRect_OriginalAspectRatio As Single           'Original aspect ratio h/w
    '
    Public shp_ImageToBeClipped As Word.Shape               'Image to eb clipped as a shape
    Public shp_ImageToBeClipped_as_cShapeMgr As cShapeMgr   'cShapeMgr version
    '
    Public shp_toBeFilled As cShapeMgr
    '
    Public parentBackPanelMgr As cBackPanelMgr
    Public parentImageMgr As cImageMgr
    '
    Public LayoutName As String
    '
    'shp_ImageToBeClipped
    '
    'Cropping nvalues (in pts) relative to the left, right, top and bottom boundaries
    'of the image to be cropped
    Public crop_Delta_Left, crop_Delta_Right, crop_Delta_Top, crop_Delta_Bottom As Single
    '
    Public Hmin As Single
    Public Hmax As Single
    Public Vmin As Single
    Public Vmax As Single
    Public cropRectWidthMax As Single

    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        '
        Me.crop_Delta_Left = 0
        Me.crop_Delta_Right = 0
        Me.crop_Delta_Top = 0
        Me.crop_Delta_Bottom = 0
        '
        '
        ' Add any initialization after the InitializeComponent() call.
    End Sub

    Private Sub frm_pictControl2_Activated(sender As Object, e As EventArgs) Handles Me.Activated

    End Sub
End Class
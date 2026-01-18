Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
''' <summary>
''' 
'''This class deals with all things related to Shapes with memories
'''
'''Peter Mikelaitis October 2015...http://mikl.com.au
'''Ported to VB.NET Jan 2017 from version 97p21p05
'''
''' </summary>
Public Class cShapeMgr
    Public shp As Word.Shape
    Public name As String
    Public altText As String
    Public height As Single
    Public width As Single
    Public left As Single
    Public top As Single
    Public rotation As Single
    Public aspectRatio As Single
    Public anchor As Word.Range
    Public hf As Word.HeaderFooter
    '
    Public width_original As Single
    Public height_original As Single
    Public scaleFactor_W As Single                  'Original Width / New Width
    Public scaleFactor_H As Single                  'Original Height / New Height

    Public DimensionsAvailable As Boolean
    Public PositionAvailable As Boolean
    Public RotationAvailable As Boolean
    '
    Public Sub New()
        Me.shp = Nothing
        Me.name = "nothing"
        Me.altText = "nothing"
        Me.height = 0#
        Me.width = 0#
        Me.width_original = 0#
        Me.height_original = 0#
        '
        Me.left = 0#
        Me.top = 0#
        Me.rotation = 0#
        Me.aspectRatio = 0#
        Me.anchor = Nothing
        Me.hf = Nothing
        '
        Me.scaleFactor_W = 1.0#
        Me.scaleFactor_H = 1.0#

        Me.DimensionsAvailable = False
        Me.PositionAvailable = False
        Me.RotationAvailable = False
    End Sub
    Public Sub InitShape(shp As Word.Shape, ByRef hf As Word.HeaderFooter)
        Me.shp = shp
        Me.name = shp.Name
        Me.altText = shp.AlternativeText
        Me.altText = Nothing
        Me.height = shp.Height
        Me.width = shp.Width
        Me.height_original = shp.Height
        Me.width_original = shp.Width

        Me.top = shp.Top
        Me.left = shp.Left
        Me.rotation = shp.Rotation
        Me.anchor = shp.Anchor
        Me.hf = hf
        '
        Me.scaleFactor_W = 1.0#
        Me.scaleFactor_H = 1.0#
        '
        Me.aspectRatio = Me.height / Me.width
        Me.DimensionsAvailable = True
        Me.PositionAvailable = True
        Me.RotationAvailable = False
        '
    End Sub
    '
    ''' <summary>
    ''' This method takes a modified version of the object that was to create this
    ''' instance of the cShapeMgr and will work out the horizontal and vertical scale
    ''' factors. Where scaleFactor is defined as original height or width / new height or width
    ''' </summary>
    ''' <param name="shp"></param>
    Public Sub AdjustScaleFactors(ByRef shp As Word.Shape)
        Me.scaleFactor_W = Me.width_original / shp.Width
        Me.scaleFactor_H = Me.height_original / shp.Height
        '
    End Sub

    'This class retrieves and holds sepecific information
    'relating to the Shape of strShapeName.. This information
    'is encoded in the name which is defined in GetLayoutItems
    'in CH_InsImg
    '
    Public Sub InitShape_Old(strShapeName As String)
        Dim tokens() As String
        Dim items() As String
        Dim maxIndex As Integer
        Dim objTools As New cTools()

        DimensionsAvailable = False
        PositionAvailable = False
        RotationAvailable = False
        Me.name = ""

        On Error GoTo finis

        tokens = Split(strShapeName, "?")
        Me.name = tokens(0)

        If tokens(1) <> "" Then
            Me.DimensionsAvailable = True
            'Get dimensions
            items = Split(tokens(1), ",")
            Me.height = objTools.MillimetersToPoints(CSng(items(0)))
            Me.width = objTools.MillimetersToPoints(CSng(items(1)))
        End If
        If tokens(2) <> "" Then
            Me.PositionAvailable = True
            items = Split(tokens(2), ",")
            Me.left = objTools.MillimetersToPoints(CSng(items(0)))
            Me.top = objTools.MillimetersToPoints(CSng(items(1)))
        End If
        If tokens(3) <> "" Then
            Me.RotationAvailable = True
            Me.rotation = CSng(tokens(3))
        End If
        Exit Sub

finis:

    End Sub

    Public Sub SetShapeParameters(ByRef NewPic As Word.Shape)
        'This method willthe dimensions, position and rotation
        'of the shape NewPic in accordance with parameters derived
        'from the shape's LongName when the init has been run
        If NewPic.Name <> Me.name Then NewPic.Name = Me.name
        If Me.DimensionsAvailable Then
            NewPic.Height = Me.height
            NewPic.Width = Me.width
        End If
        If Me.PositionAvailable Then
            NewPic.Left = Me.left
            NewPic.Top = Me.top
        End If
        If Me.RotationAvailable Then
            NewPic.Rotation = Me.rotation
        End If
    End Sub
    '


End Class

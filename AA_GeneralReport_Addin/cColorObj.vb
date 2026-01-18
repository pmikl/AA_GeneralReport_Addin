Imports System.Drawing
Imports System.Math
Imports Microsoft.Office.Tools.Ribbon

'
Public Class cColorObj
    'Note that RGB-->HSL-->RGB is not quite
    'right. We end up with a colour drift.
    'But we can use variations in the
    'luminance and get what we expect
    '
    Public myColour As Color               'Colour object
    Public lstOfVariations As Collection    'Colour variations for chnages in saturation
    '
    Public h_Office As Double               '0-360 degrees represented by 0-255.. Therefore h_Office = h*scale
    Public h As Double                      '0-360 degrees
    Public s As Double                      'scaled to between 0 and 1
    Public l As Double                      'scaled to between 0 and 1
    '
    Public r As Integer          '
    Public g As Integer         '
    Public b As Integer          '
    '
    Public rgb As Integer
    '
    Public rTest As Integer          '
    Public gTest As Integer          '
    Public bTest As Integer          '
    '
    Public strColor As String
    '
    Public rgbTest As Integer

    Public Sub New()

    End Sub
    Public Sub New(rgb As Long, Optional numVariations As Integer = 6)
        '
        Me.rgb = rgb
        Me.myColour = ColorTranslator.FromWin32(rgb)
        '
        Me.r = myColour.R
        Me.g = myColour.G
        Me.b = myColour.B
        '
        Me.h = Me.myColour.GetHue
        Me.s = Me.myColour.GetSaturation
        Me.l = Me.myColour.GetBrightness
        '
        Me.lstOfVariations = Me.getListOfVariations(numVariations)
        '
        Me.strColor = CStr(Me.r) + "," + CStr(Me.g) + "," + CStr(Me.b)
        '
        '
        'Me.parseRGB(rgb)
        '
    End Sub
    '
    Public Sub New(red As Integer, green As Integer, blue As Integer, Optional numVariations As Integer = 6)
        '
        Me.myColour = Color.FromArgb(255, red, green, blue)
        Me.rgb = ColorTranslator.ToWin32(Me.myColour)
        '
        Me.r = red
        Me.g = green
        Me.b = blue
        '
        Me.h = Me.myColour.GetHue
        Me.s = Me.myColour.GetSaturation
        Me.l = Me.myColour.GetBrightness
        '
        Me.lstOfVariations = Me.getListOfVariations(numVariations)
        Me.strColor = CStr(Me.r) + "," + CStr(Me.g) + "," + CStr(Me.b)
        '
    End Sub
    '
    ''' <summary>
    ''' strColor is a string organised as 'Red,Green,Blue'
    ''' </summary>
    ''' <param name="strColor"></param>
    Public Sub New(strColor As String, Optional numVariations As Integer = 6)
        Dim tokens() As String
        Dim red, green, blue As Integer
        '
        Me.strColor = strColor
        tokens = strColor.Split(",")
        red = CInt(tokens(0))
        green = CInt(tokens(1))
        blue = CInt(tokens(2))
        '
        Me.myColour = Color.FromArgb(255, red, green, blue)
        Me.rgb = ColorTranslator.ToWin32(Me.myColour)
        '
        Me.r = red
        Me.g = green
        Me.b = blue
        '
        Me.h = Me.myColour.GetHue
        Me.s = Me.myColour.GetSaturation
        Me.l = Me.myColour.GetBrightness
        '
        Me.lstOfVariations = Me.getListOfVariations(numVariations)
        '
    End Sub
    '
    ''' <summary>
    ''' This method will return a 'swatch' of colour in the supplied RibbonDropDownItem.
    ''' Note that the ScreenTip is set to a string that identifies the Colour type 
    ''' (strFillType) and the SuperTip contains a string that identifies the
    ''' color as Red,Green,Blue
    ''' </summary>
    ''' <param name="drpDown"></param>
    ''' <param name="objColour"></param>
    ''' <param name="doAsPattern"></param>
    ''' <returns></returns>
    Public Function _getColoredDropDown(ByRef drpDown As RibbonDropDownItem, ByRef objColour As cColorObj, doAsPattern As Boolean) As Boolean
        Dim myImage As Bitmap
        Dim gfx As Graphics
        Dim brsh As SolidBrush
        Dim hatchForeColour, hatchBackColour As Color
        Dim strFillType As String
        '
        _getColoredDropDown = False
        '
        strFillType = "unknown"
        hatchForeColour = objColour.myColour
        hatchBackColour = Color.White
        '
        'drpDown = Globals.Factory.GetRibbonFactory.CreateRibbonDropDownItem()
        myImage = New Bitmap(32, 32, Imaging.PixelFormat.Format24bppRgb)
        gfx = Graphics.FromImage(myImage)
        '
        If Not doAsPattern Then
            brsh = New SolidBrush(objColour.myColour)
            gfx.FillRectangle(brsh, 0, 0, myImage.Width, myImage.Height)
            drpDown.Image = myImage
            drpDown.Label = "Test"
            drpDown.ScreenTip = "Solid Colour"
            drpDown.SuperTip = CStr(objColour.r) + "," + CStr(objColour.g) + "," + CStr(objColour.b)
            '
            _getColoredDropDown = True
        Else
            'Excel hatch specs and the Drawing NameSpace hatch spects are different.. and so they need to be
            'translated
            'Select Case htchStyle
            'Case XlPattern.xlPatternCrissCross
            'strFillType = "CrissCross"
            'htchStyle = Drawing2D.HatchStyle.Percent30
            'Case XlPattern.xlPatternDown
            'strFillType = "DarkDownwardDiagonal"
            'htchStyle = Drawing2D.HatchStyle.DarkDownwardDiagonal
            'Case XlPattern.xlPatternHorizontal
            'strFillType = "DarkHorizontal"
            'htchStyle = Drawing2D.HatchStyle.DarkHorizontal
            'htchStyle = Drawing2D.HatchStyle.LargeConfetti

            'Case Else
            'strFillType = "unknown"
            'htchStyle = Drawing2D.HatchStyle.Shingle
            'End Select
            '
            'hatchBrsh = New Drawing2D.HatchBrush(htchStyle, hatchForeColour, hatchBackColour)
            'gfx.FillRectangle(hatchBrsh, 0, 0, myImage.Width, myImage.Height)
            'drpDown.Label = strFillType
            'drpDown.ScreenTip = strFillType
            'drpDown.SuperTip = CStr(objColour.r) + "," + CStr(objColour.g) + "," + CStr(objColour.b)
            'drpDown.Image = myImage
            'drpDown.Label = "Test"
            _getColoredDropDown = True
        End If
        '
    End Function
    '

    Public Function getListOfVariations(Optional numVariations As Integer = 6) As Collection
        'Will return a list of RGB values derived from the initial colour
        'of this object.. The variations are obtained by chnaging the saturation
        getListOfVariations = getColoursFromSeed(True, numVariations, Me.myColour)
    End Function
    '
    Public Function getColoursFromSeed(doLightToDark As Boolean, numSteps As Integer, ByRef seedColour As Color) As Collection
        'This function will return a collection of RGB (integer) colours that
        'represent a number of steps up and down the seedColour vector... The Collection
        'can be accessed from 0 to numSteps-1
        '
        Dim delta As Single
        Dim i As Integer
        Dim stepColourRGB As Integer
        Dim stepBrightness As Single
        Dim stepSaturation As Single
        '
        'Make the first entry the seed colour
        getColoursFromSeed = New Collection()
        'getColoursFromSeed.Add(Information.RGB(seedColour.R, seedColour.G, seedColour.B), "0")
        '
        If doLightToDark Then
            For i = 0 To numSteps - 1
                delta = 1.0 / CSng(numSteps + 2)
                stepBrightness = 1.0 - (i + 1) * delta
                stepSaturation = 1.0 - i * delta
                '
                stepColourRGB = Me.convertToRGB_fromHSL(seedColour.GetHue, seedColour.GetSaturation, CDec(stepBrightness))
                'stepColourRGB = Me.convertToRGB_fromHSL(seedColour.GetHue, stepSaturation, seedColour.GetBrightness())
                'tmpColor = Color.FromArgb(step)
                'objHSL = New cHSLColour(seedColour.GetHue(), seedColour.GetSaturation(), CDec(stepBrightness))
                '
                '*** For test purposes these routines generate results that are known to be good
                'tmpColour = Color.FromArgb(255, CInt(objHSL.r), CInt(objHSL.g), CInt(objHSL.b))
                'stepColourRGB = Me.getRGBFromColor(stepColour)
                '****
                '
                getColoursFromSeed.Add(stepColourRGB, CStr(i))
            Next
        Else
            For i = numSteps - 1 To 1 Step -1
                delta = 1.0 / CSng(numSteps + 2)
                stepBrightness = 1.0 - (i + 1) * delta
                stepSaturation = 1.0 - i * delta
                '
                stepColourRGB = Me.convertToRGB_fromHSL(seedColour.GetHue, seedColour.GetSaturation, CDec(stepBrightness))
                'stepColourRGB = Me.convertToRGB_fromHSL(seedColour.GetHue, stepSaturation, seedColour.GetBrightness())
                'tmpColor = Color.FromArgb(step)
                'objHSL = New cHSLColour(seedColour.GetHue(), seedColour.GetSaturation(), CDec(stepBrightness))
                '
                '*** For test purposes these routines generate results that are known to be good
                'tmpColour = Color.FromArgb(255, CInt(objHSL.r), CInt(objHSL.g), CInt(objHSL.b))
                'stepColourRGB = Me.getRGBFromColor(stepColour)
                '****
                '
                getColoursFromSeed.Add(stepColourRGB, CStr(i))
            Next
        End If

        '
    End Function
    '
    '
    Public Function convertToRGB_fromHSL(hue As Double, saturation As Double, luminance As Double) As Integer
        'Formula on http://www.rapidtables.com/convert/color/hsl-to-rgb.htm
        Dim C As Double
        Dim X As Double
        Dim m As Double
        Dim rDash As Double
        Dim gDash As Double
        Dim bDash As Double
        Dim red, green, blue As Double
        Dim rInt, gInt, bInt As Integer
        '
        convertToRGB_fromHSL = 0
        '
        C = (1.0 - Math.Abs(2.0 * luminance - 1.0)) * saturation
        X = C * (1.0 - Math.Abs((hue / 60.0) Mod 2.0 - 1.0))
        m = luminance - C / 2.0
        '
        If hue >= 0.0 And hue < 60.0 Then
            rDash = C
            gDash = X
            bDash = 0.0
        End If
        '
        If hue >= 60.0 And hue < 120.0 Then
            rDash = X
            gDash = C
            bDash = 0.0
        End If
        '
        If hue >= 120.0 And hue < 180.0 Then
            rDash = 0.0
            gDash = C
            bDash = X
        End If
        '
        If hue >= 180.0 And hue < 240.0 Then
            rDash = 0.0
            gDash = X
            bDash = C
        End If
        '
        If hue >= 240.0 And hue < 300.0 Then
            rDash = X
            gDash = 0.0
            bDash = C
        End If
        '
        If hue >= 300.0 And hue < 360.0 Then
            rDash = C
            gDash = 0.0
            bDash = X
        End If
        '
        red = (rDash + m) * 255.0
        green = (gDash + m) * 255.0
        blue = (bDash + m) * 255.0
        '
        rInt = CInt(Math.Round(red))
        gInt = CInt(Math.Round(green))
        bInt = CInt(Math.Round(blue))
        '
        convertToRGB_fromHSL = Me.generateRGB(rInt, gInt, bInt)
        '
        'Me.generateRGB(CInt(Me.r), CInt(Me.g), CInt(Me.b))
        'Because we are going from Single to Integer via UInt there may
        'be small rounding errors that will cause the RGB colour to go abouve
        'its maximum value of &HFFFFFF.. If it does then set it back
        'If Me.rgb > &HFFFFFF Then Me.rgb = &HFFFFFF
    End Function
    '
    Public Function generateRGB(red As Integer, green As Integer, blue As Integer) As Integer
        'This function performs the reverse of parseRGB
        Dim result As Integer
        Dim tmp As Integer
        '
        generateRGB = 0
        result = 0
        '
        'Add the red component
        result = red And &HFF
        '
        'Add the green component
        tmp = green << 8
        tmp = tmp And &HFF00
        result = result Or tmp
        '
        'Add the blue component
        tmp = blue << 16
        tmp = tmp And &HFF0000
        result = result Or tmp
        '
        result = result And &HFFFFFF
        generateRGB = result
    End Function
    '
    Public Function parseRGB(rgb) As Collection
        'This method will breakup a Win32 rgb number into its component
        'parts... Note that the 32 bit integer is organised as b,g,r
        'where left to right is MSBit to LSBit
        Dim red, green, blue As Integer
        Dim msk As Integer
        '
        parseRGB = New Collection()
        '
        msk = &HFF
        red = rgb And msk
        '
        'Green
        green = rgb And &HFF00
        green = green >> 8
        '
        blue = rgb And &HFF0000
        blue = blue >> 16
        '
        Me.r = red
        Me.g = green
        Me.b = blue
        '
        parseRGB.Add(Me.r, "red")
        parseRGB.Add(Me.g, "green")
        parseRGB.Add(Me.b, "blue")
    End Function


End Class

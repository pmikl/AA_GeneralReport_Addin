Imports System.IO
'Imports DocumentFormat.OpenXml.Spreadsheet
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word
'
Public Class cThemeMgr
    Public strThemeFileName As String
    '
    Public Sub New()
        'strThemeFileName = "AA_Theme_for_GeneralReport_with_CustClrs.thmx"
        strThemeFileName = "AA_Theme_Base_with_CustClrs.thmx"
    End Sub
    '
    ''' <summary>
    ''' This method will read the standard/current AA Theme file from the Resources folder,
    ''' then writes it to C:\Templates\Themes. Then the theme file is applied to the current
    ''' 'office' document. Note that it creates the folder path if necessary. The method will return
    ''' true if all is OK, otherwise it will return false.
    ''' </summary>
    ''' <returns></returns>
    Public Function thm_Set_ThemeToAAStd_fromFile(ByRef myDoc As Word.Document) As Boolean
        Dim objGlobals As New cGlobals()
        Dim thmFile As Byte()
        Dim myFileInfo As FileInfo
        Dim rslt, writeOver As Boolean
        Dim directoryInfo As System.IO.DirectoryInfo
        Dim strThemeDirectory, strThemeFilePath As String
        '
        rslt = False
        writeOver = False
        '
        Try
            strThemeDirectory = objGlobals.glb_getDir_inUseforTemplates() + "\" + "Themes\"
            'strThemeDirectory = "C:\Templates\DocumentThemes\"
            'strThemeFileName = "AA_Theme_for_GeneralReport_with_CustClrs_20240808.thmx"
            'strThemeFileName = "AA_Theme_for_GeneralReport_with_CustClrs.thmx"
            strThemeFilePath = strThemeDirectory + strThemeFileName
            '
            'Create the themes directory, then retrieve the theme file from the Resources
            'file and then write it to Themes
            directoryInfo = New DirectoryInfo(strThemeDirectory)
            If Not directoryInfo.Exists Then
                directoryInfo = System.IO.Directory.CreateDirectory(strThemeDirectory)
            End If
            '
            'thmFile = My.Resources.AA_Theme_for_GeneralReport_with_CustClrs_20240808
            'thmFile = My.Resources.AA_Theme_for_GeneralReport_with_CustClrs_20250926
            thmFile = My.Resources.AA_Theme_Base_with_CustClrs_20250926
            myFileInfo = New FileInfo(strThemeFilePath)
            '
            'Write th theme file if it doesn't exist.. We can override this behaviour if writeOver
            'is set to true
            If Not myFileInfo.Exists Then
                'It doesn't exist so let's create it
                File.WriteAllBytes(strThemeFilePath, thmFile)
            Else
                'It does exist, do nothing unless told to writeOver
                If writeOver Then
                    Try
                        'Try to overwrite, but if we fail (becuase of a file lock?) then we recover
                        File.WriteAllBytes(strThemeFilePath, thmFile)
                    Catch ex As Exception

                    End Try
                End If
            End If
            '
            myDoc.ApplyDocumentTheme(strThemeFilePath)
            '
            Me.thm_Set_ThemeToAAStd_Manually(myDoc)                             'To override the file settings. Does not affect the custClrs
            '   
            rslt = True
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will setup all of the Theme elements manually. It is used if the standard theme file
    ''' cannot be found
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub thm_Set_ThemeToAAStd_Manually(ByRef myDoc As Word.Document)
        Dim thm As OfficeTheme
        Dim thmColorScheme As ThemeColorScheme
        Dim thmFontScheme As ThemeFontScheme
        Dim thmFont As ThemeFont
        '
        thm = myDoc.DocumentTheme
        thmColorScheme = thm.ThemeColorScheme
        thmFontScheme = thm.ThemeFontScheme
        'Dim thmColor As ThemeColor
        '
        'thmColor = New ThemeColor()
        'thmColorScheme = ThemeColorScheme
        'thmColor.RGB = RGB(0, 0, 0)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark1).RGB = RGB(0, 0, 0)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight1).RGB = RGB(255, 255, 255)
        'thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark2).RGB = RGB(132, 206, 136)
        'thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight2).RGB = RGB(20, 0, 52)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark2).RGB = RGB(20, 0, 52)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight2).RGB = RGB(132, 206, 136)

        '
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent1).RGB = RGB(108, 63, 152)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent2).RGB = RGB(157, 133, 190)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent3).RGB = RGB(200, 200, 200)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent4).RGB = RGB(125, 125, 125)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent5).RGB = RGB(123, 189, 214)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent6).RGB = RGB(0, 106, 159)
        '
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeHyperlink).RGB = RGB(123, 189, 214)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink).RGB = RGB(159, 209, 139)
        '
        'Default windows values for Hyperlink and Followed Hyperlink
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeHyperlink).RGB = RGB(0, 102, 204)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink).RGB = RGB(128, 0, 128)
        '
        thmFont = thmFontScheme.MinorFont.Item(MsoFontLanguageIndex.msoThemeLatin)
        thmFont.Name = "Arial Narrow"
        'thmFont.Name = "Yu Gothic Medium"
        '
        thmFont = thmFontScheme.MajorFont.Item(MsoFontLanguageIndex.msoThemeLatin)
        thmFont.Name = "Yu Gothic Medium"
        '
    End Sub
    '
    '
    ''' <summary>
    ''' Earlier light orange theme
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub thm_Set_ThemeToAALightOrange_Manually(ByRef myDoc As Word.Document)
        Dim thm As OfficeTheme
        Dim thmColorScheme As ThemeColorScheme
        Dim thmFontScheme As ThemeFontScheme
        Dim thmFont As ThemeFont
        '
        thm = myDoc.DocumentTheme
        thmColorScheme = thm.ThemeColorScheme
        thmFontScheme = thm.ThemeFontScheme
        'Dim thmColor As ThemeColor
        '
        'thmColor = New ThemeColor()
        'thmColorScheme = ThemeColorScheme
        'thmColor.RGB = RGB(0, 0, 0)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark1).RGB = RGB(0, 0, 0)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight1).RGB = RGB(255, 255, 255)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark2).RGB = RGB(20, 0, 52)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight2).RGB = RGB(244, 231, 237)
        '
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent1).RGB = RGB(157, 87, 166)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent2).RGB = RGB(255, 174, 59)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent3).RGB = RGB(165, 165, 165)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent4).RGB = RGB(182, 137, 193)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent5).RGB = RGB(234, 139, 0)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent6).RGB = RGB(191, 191, 191)
        '
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeHyperlink).RGB = RGB(255, 222, 102)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink).RGB = RGB(212, 144, 197)
        '
        '
        thmFont = thmFontScheme.MinorFont.Item(MsoFontLanguageIndex.msoThemeLatin)
        'thmFont.Name = "Arial Narrow"
        thmFont.Name = "Yu Gothic Medium"
        '
        thmFont = thmFontScheme.MajorFont.Item(MsoFontLanguageIndex.msoThemeLatin)
        thmFont.Name = "Yu Gothic Medium"
        '
    End Sub
    '
    '
    ''' <summary>
    ''' This method will setup all of the Theme elements manually. It is used if the standard theme file
    ''' cannot be found
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub thm_Set_ThemeToAAStd_20240808_Manually(ByRef myDoc As Word.Document)
        Dim thm As OfficeTheme
        Dim thmColorScheme As ThemeColorScheme
        Dim thmFontScheme As ThemeFontScheme
        Dim thmFont As ThemeFont
        '
        thm = myDoc.DocumentTheme
        thmColorScheme = thm.ThemeColorScheme
        thmFontScheme = thm.ThemeFontScheme
        'Dim thmColor As ThemeColor
        '
        'thmColor = New ThemeColor()
        'thmColorScheme = ThemeColorScheme
        'thmColor.RGB = RGB(0, 0, 0)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark1).RGB = RGB(0, 0, 0)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight1).RGB = RGB(255, 255, 255)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark2).RGB = RGB(108, 63, 105)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight2).RGB = RGB(20, 0, 52)
        '
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent1).RGB = RGB(157, 133, 190)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent2).RGB = RGB(125, 125, 125)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent3).RGB = RGB(200, 200, 200)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent4).RGB = RGB(0, 106, 159)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent5).RGB = RGB(123, 189, 214)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent6).RGB = RGB(66, 141, 82)
        '
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeHyperlink).RGB = RGB(123, 189, 214)
        thmColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink).RGB = RGB(159, 209, 139)
        '
        '
        thmFont = thmFontScheme.MinorFont.Item(MsoFontLanguageIndex.msoThemeLatin)
        'thmFont.Name = "Arial Narrow"
        thmFont.Name = "Yu Gothic Medium"
        '
        thmFont = thmFontScheme.MajorFont.Item(MsoFontLanguageIndex.msoThemeLatin)
        thmFont.Name = "Yu Gothic Medium"
        '
    End Sub
    '
    '
    '
    'Public 
    '
    '
    ''' <summary>
    ''' This method will return the string to be used as the Font Name for then
    ''' Theme Body font.
    ''' </summary>
    ''' <returns></returns>
    Public Function thm_get_FontNameForBody() As String
        Dim rslt As String
        '
        rslt = "+mn-lt"         'Minor(Body) - latin

        Return rslt
    End Function
    '
    '
    ''' <summary>
    ''' This method will return the string to be used as the Font Name for then
    ''' Theme Heading font.
    ''' </summary>
    ''' <returns></returns>
    Public Function thm_get_FontNameForHeading() As String
        Dim rslt As String
        '
        rslt = "+mj-lt"         'Major(Heading) - latin

        Return rslt
    End Function
    '
    Public Function thm_change_FontBody(ByRef myDoc As Word.Document, Optional strBodyFontName As String = "Yu Gothic Medium") As Boolean
        Dim thm As OfficeTheme
        Dim thmColorScheme As ThemeColorScheme
        Dim thmFontScheme As ThemeFontScheme
        Dim thmFont As ThemeFont
        Dim rslt As Boolean = False
        '
        thm = myDoc.DocumentTheme
        thmColorScheme = thm.ThemeColorScheme
        thmFontScheme = thm.ThemeFontScheme
        '
        Try
            thmFont = thmFontScheme.MinorFont.Item(MsoFontLanguageIndex.msoThemeLatin)
            thmFont.Name = strBodyFontName
            rslt = True
        Catch ex As Exception
            rslt = False
        End Try
        '
        'thmFont = thmFontScheme.MajorFont.Item(MsoFontLanguageIndex.msoThemeLatin)
        'thmFont.Name = "Yu Gothic Medium"
        Return rslt
        '
    End Function
    '
    Public Function thm_change_FontHeading(ByRef myDoc As Word.Document, Optional strBodyFontName As String = "Yu Gothic Medium") As Boolean
        Dim thm As OfficeTheme
        Dim thmColorScheme As ThemeColorScheme
        Dim thmFontScheme As ThemeFontScheme
        Dim thmFont As ThemeFont
        Dim rslt As Boolean = False
        '
        thm = myDoc.DocumentTheme
        thmColorScheme = thm.ThemeColorScheme
        thmFontScheme = thm.ThemeFontScheme
        '
        Try
            thmFont = thmFontScheme.MajorFont.Item(MsoFontLanguageIndex.msoThemeLatin)
            thmFont.Name = strBodyFontName
            rslt = True
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
        '
    End Function
    '
End Class

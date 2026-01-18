Imports System.Windows.Forms
Imports System.Drawing
Imports System.Math
Imports Microsoft.Office.Core
Imports Microsoft.Office.Interop.Word
'
Public Class cColorMgr
    Public lstOfThemeButtons As Collection
    Public lstOfthemeToolStrips As Collection
    '
    Public lst_of_CustomColourButtons As Collection
    Public lstOfCustomColors As Collection
    Public lstOfSeedColors As Collection
    Public numColumns, numRows As Integer               'Custom colours

    '
    Public numColumns_Theme As Integer
    Public numRows_Theme As Integer
    Public extraPaletteHeight As Integer
    Public extraPaletteWidth As Integer

    '
    Public strip As ToolStrip
    '
    'Public strp As ToolStrip

    '
    Public Sub New()
        Me.lstOfThemeButtons = New Collection()
        Me.lstOfthemeToolStrips = New Collection()
        Me.lstOfCustomColors = New Collection()
        Me.lstOfSeedColors = New Collection()
        '
        Me.numColumns_Theme = 0
        Me.numRows_Theme = 0
        Me.numColumns = 0
        Me.numRows = 0
        Me.strip = Nothing
    End Sub
    '
    ''' <summary>
    ''' This method takes in the collection of custom palette buttons and returns an
    ''' XML file that can be inserted into the Theme File to Add a Custom Colors sectio
    ''' </summary>
    Public Function colr_build_CustClrsXML() As String
        Dim btn As ToolStripButton
        Dim lstOfButtons As List(Of ToolStripButton)
        Dim j, row As Integer
        Dim strXMLCustClrs As String
        '
        '
        '
        strXMLCustClrs = "<a:custClrLst>" + vbCrLf
        '
        'MsgBox("NumRows = " + Me.numRows.ToString() + "NumClumns = " + Me.numColumns.ToString)
        'Seems that we can only display 5 rows in Custom Colours
        'For row = 1 to 10 Step 2
        For row = 1 To 5 Step 1
            lstOfButtons = getRow(row, 5)

            For j = 0 To lstOfButtons.Count - 1
                btn = lstOfButtons.Item(j)
                Me.getRGB_Hex(btn.BackColor)
                '
                strXMLCustClrs = strXMLCustClrs + "<a:custClr name=" + """" + btn.ToolTipText + """" + ">" + vbCrLf
                strXMLCustClrs = strXMLCustClrs + "<a:srgbClr val=" + """" + Me.getRGB_Hex(btn.BackColor) + """" + " />" + vbCrLf
                strXMLCustClrs = strXMLCustClrs + "</a:custClr>" + vbCrLf
                '
            Next
        Next row
        '
        strXMLCustClrs = strXMLCustClrs + "</a:custClrLst>"
        '
        Return strXMLCustClrs

    End Function
    '
    Public Function colr_build_CustClrsVBNET(ByRef lstOfCustomColors As Collection, Optional doSeedColoursAtEnd As Boolean = False) As String
        Dim j, k, rgbColour As Integer
        Dim lst As Collection
        Dim numColumns, numRows As Integer
        Dim strDim, strRGB, strLine, strResult As String
        Dim objCol As cColorObj
        '

        lst = lstOfCustomColors("0")
        strDim = ""
        strLine = ""
        strResult = ""
        '
        'Me.colr_modify_CustomColour(RGB(255, 0, 0), 0, 0, lstOfCustomColors)

        '
        numColumns = lstOfCustomColors.Count
        numRows = lst.Count
        '
        'Setup the dimension statements
        '
        For j = 0 To lstOfCustomColors.Count - 1
            strDim = strDim + "Dim" + " lst" + CStr(j) + " As New Collection()" + vbCrLf
        Next
        '
        strDim = vbCrLf + strDim + "Dim" + " lstofCustomColours" + " As New Collection()" + vbCrLf

        '
        For j = 0 To lstOfCustomColors.Count - 1
            'Get each column, then add the elements to each column.. If doSeed is true, then we add the
            'seed colour to the end
            '
            lst = lstOfCustomColors.Item(CStr(j))
            'Elements of each column
            For k = 0 To lst.Count - 1
                rgbColour = lst.Item(CStr(k))
                objCol = New cColorObj(rgbColour)
                strRGB = "RGB(" + objCol.strColor + ")"
                '
                strLine = strLine + "lst" + CStr(j) + ".Add(" + strRGB + ", " + """" + CStr(k) + """" + ")" + vbCrLf
            Next
            '
            If doSeedColoursAtEnd Then
                rgbColour = Me.lstOfSeedColors.Item(CStr(j))
                objCol = New cColorObj(rgbColour)
                strRGB = "RGB(" + objCol.strColor + ")"
                '
                strLine = strLine + "lst" + CStr(j) + ".Add(" + strRGB + ", " + """" + CStr(lst.Count) + """" + ")" + vbCrLf
                '
            End If
            '
            strLine = strLine + "lstofCustomColours.Add(" + "lst" + CStr(j) + ", " + """" + CStr(j) + """" + ")" + vbCrLf + vbCrLf
        Next
        '
        strResult = strDim + vbCrLf + strLine + vbCrLf + vbCrLf + "Return lstofCustomColours"
        '
        Return strResult
        '
    End Function
    '
    Public Function getRow(rowNum As Integer, totalRows As Integer) As List(Of ToolStripButton)
        Dim btn As ToolStripButton
        'Dim item00, item01, item02, item03 As ToolStripButton
        'Dim item10, item11, item12, item13 As ToolStripButton

        Dim lstOfRowItems As New List(Of ToolStripButton)
        Dim columnNum As Integer
        '
        'start of row n = n * numColumns, the row has numColumns items
        '
        'row 0
        'item00 = Me.lst_of_CustomColourButtons(CStr(0))
        'item01 = Me.lst_of_CustomColourButtons(CStr(0 + totalRows))
        'item02 = Me.lst_of_CustomColourButtons(CStr(0 + 2 * totalRows))
        'item03 = Me.lst_of_CustomColourButtons(CStr(0 + 3 * totalRows))
        '
        '
        'row 1
        'item10 = Me.lst_of_CustomColourButtons(CStr(1))
        'item11 = Me.lst_of_CustomColourButtons(CStr(1 + totalRows))
        'item12 = Me.lst_of_CustomColourButtons(CStr(1 + 2 * totalRows))
        'item13 = Me.lst_of_CustomColourButtons(CStr(1 + 3 * totalRows))
        '
        For columnNum = 0 To numColumns - 1
            'btn = Me.lst_of_CustomColourButtons(CStr(rowNum + columnNum * totalRows))
            btn = Me.lst_of_CustomColourButtons.Item(rowNum + columnNum * totalRows)

            lstOfRowItems.Add(btn)
        Next
        '
        Return lstOfRowItems
    End Function
    '
    Public Sub refreshThemePalette()
        'Check out the current list of theme buttons in the palette. If they
        'exist then we can chnage the colour
        Dim lstOfThemeColors As Collection
        Dim objColor As cColorObj
        Dim btn As ToolStripButton
        Dim strp As ToolStrip
        Dim column, row As Integer
        Dim rgbColor As Integer
        Dim baseColor As Color
        Dim variationColor As Color
        '
        '
        lstOfThemeColors = Me.getThemeColours()
        '
        For column = 0 To Me.lstOfthemeToolStrips.Count - 1
            strp = Me.lstOfthemeToolStrips(CStr(column))
            objColor = New cColorObj(CInt(lstOfThemeColors(CStr(column))))
            '
            'Now do first item
            '
            For row = 0 To objColor.lstOfVariations.Count
                btn = strp.Items().Item(row)
                If row = 0 Then
                    btn.BackColor = objColor.myColour
                    btn.ToolTipText = Me.getRGB_longForm(btn.BackColor)
                    btn.ToolTipText = Me.colorPalette_Row0_ToolTips(btn, row, column)


                Else
                    rgbColor = CInt(objColor.lstOfVariations(CStr(row - 1)))
                    baseColor = ColorTranslator.FromWin32(rgbColor)
                    variationColor = Color.FromArgb(255, baseColor)
                    'btn = Me.lstOfThemeButtons.Item(k)
                    btn.BackColor = variationColor
                    btn.ToolTipText = Me.getRGB_longForm(btn.BackColor)
                End If
            Next
        Next
        '
        '
        '

        'If Me.lstOfThemeButtons.Count > 0 Then
        'We have some theme buttons
        'lstOfThemeColors = Me.getThemeColours()
        'k = 0
        'For column = 0 To Me.numColumns_Theme - 1
        'objColor = New cColourObj(CInt(lstOfThemeColors(CStr(column))))
        'For row = 0 To objColor.lstOfVariations.Count - 1
        'strKey = CStr(column) & "," & CStr(row)
        'k = k + 1
        'btn = Me.lstOfThemeButtons.Item(strKey)
        'If row = 0 Then
        'btn.BackColor = objColor.myColour
        'Else
        ' rgbColor = CInt(objColor.lstOfVariations(CStr(row)))
        'baseColor = ColorTranslator.FromWin32(rgbColor)
        'variationColor = Color.FromArgb(255, baseColor)
        'btn = Me.lstOfThemeButtons.Item(k)
        'btn.BackColor = baseColor
        'End If
        'Next
        'Next
        'For i = 1 To Me.lstOfThemeButtons.Count
        'btn = Me.lstOfThemeButtons.Item(i)
        'btn.BackColor = Color.Red
        'Next
        'End If
    End Sub
    '
    Public Function colr_set_SeedColoursx() As Collection
        Dim lstOfSeedColours As New Collection()
        '
        'lstOfSeedColours.Add(RGB(255, 255, 255), "0")           'FFFFFF (5)
        'lstOfSeedColours.Add(RGB(20, 0, 52), "1")               '140034 (5)
        'lstOfSeedColours.Add(RGB(204, 102, 225), "2")             '220B41 (3)
        'lstOfSeedColours.Add(RGB(149, 79, 114), "3")              '331063 (4)
        'lstOfSeedColours.Add(RGB(0, 106, 159), "4")               '002032 (7)
        'lstOfSeedColours.Add(RGB(68, 114, 196), "5")              '102515 (6)

        'lstOfSeedColours.Add(RGB(0, 32, 255), "6")              '102515 (6)
        'lstOfSeedColours.Add(RGB(68, 114, 196), "7")              '102515 (6)
        'lstOfSeedColours.Add(RGB(66, 141, 82), "8")              '4D4D4D (2)
        'lstOfSeedColours.Add(RGB(237, 125, 49), "9")              '331063 (4)
        '
        lstOfSeedColours.Add(RGB(20, 0, 52), "0")                   'FFFFFF (5)
        lstOfSeedColours.Add(RGB(108, 63, 153), "1")                '1D1D1D (1)
        lstOfSeedColours.Add(RGB(157, 133, 190), "2")               '140034 (5)
        lstOfSeedColours.Add(RGB(0, 106, 159), "3")                 '140034 (5)
        lstOfSeedColours.Add(RGB(123, 189, 214), "4")               '220B41 (3)
        lstOfSeedColours.Add(RGB(125, 125, 125), "5")               '002032 (7)
        lstOfSeedColours.Add(RGB(200, 200, 200), "6")               '002032 (7)
        lstOfSeedColours.Add(RGB(66, 141, 82), "7")                 '102515 (6)
        lstOfSeedColours.Add(RGB(159, 209, 139), "8")              '102515 (6)
        lstOfSeedColours.Add(RGB(255, 255, 255), "9")              '4D4D4D (2)



        Return lstOfSeedColours
        '
    End Function
    '
    ''' <summary>
    ''' This method will get the list of Seed Colours used to create the
    ''' Custom COours section of the ColorPicker as per excel
    ''' </summary>
    ''' <returns></returns>
    Public Function colr_set_SeedColours() As Collection
        Dim lstOfSeedColours As New Collection()
        '
        lstOfSeedColours.Add(RGB(255, 255, 255), "0")
        lstOfSeedColours.Add(RGB(20, 0, 52), "1")
        lstOfSeedColours.Add(RGB(108, 63, 153), "2")
        lstOfSeedColours.Add(RGB(157, 133, 190), "3")
        lstOfSeedColours.Add(RGB(204, 195, 220), "4")
        lstOfSeedColours.Add(RGB(230, 226, 238), "5")
        lstOfSeedColours.Add(RGB(123, 189, 214), "6")
        lstOfSeedColours.Add(RGB(0, 106, 159), "7")
        lstOfSeedColours.Add(RGB(159, 209, 139), "8")
        lstOfSeedColours.Add(RGB(66, 141, 82), "9")
        lstOfSeedColours.Add(RGB(56, 148, 137), "10")
        lstOfSeedColours.Add(RGB(20, 84, 188), "11")
        lstOfSeedColours.Add(RGB(50, 16, 192), "12")
        lstOfSeedColours.Add(RGB(102, 83, 243), "13")
        lstOfSeedColours.Add(RGB(93, 46, 162), "14")
        lstOfSeedColours.Add(RGB(178, 149, 249), "15")
        lstOfSeedColours.Add(RGB(204, 152, 246), "16")
        lstOfSeedColours.Add(RGB(60, 146, 148), "17")
        lstOfSeedColours.Add(RGB(52, 156, 136), "18")
        'lstOfSeedColours.Add(RGB(255, 0, 0), "18")

        '    
        Return lstOfSeedColours
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will get the list of Seed Colours used to create the
    ''' Custom COours section of the ColorPicker as per excel
    ''' </summary>
    ''' <returns></returns>
    Public Function colr_set_SeedColours_10x5() As Collection
        Dim lstOfSeedColours As New Collection()
        '
        lstOfSeedColours.Add(RGB(255, 255, 255), "0")
        lstOfSeedColours.Add(RGB(20, 0, 52), "1")
        lstOfSeedColours.Add(RGB(108, 63, 153), "2")
        lstOfSeedColours.Add(RGB(157, 133, 190), "3")
        lstOfSeedColours.Add(RGB(204, 195, 220), "4")
        lstOfSeedColours.Add(RGB(230, 226, 238), "5")
        lstOfSeedColours.Add(RGB(123, 189, 214), "6")
        lstOfSeedColours.Add(RGB(0, 106, 159), "7")
        lstOfSeedColours.Add(RGB(159, 209, 139), "8")
        lstOfSeedColours.Add(RGB(66, 141, 82), "9")
        'lstOfSeedColours.Add(RGB(56, 148, 137), "10")
        'lstOfSeedColours.Add(RGB(20, 84, 188), "11")
        'lstOfSeedColours.Add(RGB(50, 16, 192), "12")
        'lstOfSeedColours.Add(RGB(102, 83, 243), "13")
        'lstOfSeedColours.Add(RGB(93, 46, 162), "14")
        'lstOfSeedColours.Add(RGB(178, 149, 249), "15")
        'lstOfSeedColours.Add(RGB(204, 152, 246), "16")
        'lstOfSeedColours.Add(RGB(60, 146, 148), "17")
        'lstOfSeedColours.Add(RGB(52, 156, 136), "18")
        '    
        Return lstOfSeedColours
        '
    End Function

    '
    '
    ''' <summary>
    ''' This method will get the list of Seed Colours used to create the
    ''' Custom COours section of the ColorPicker as per excel
    ''' </summary>
    ''' <returns></returns>
    Public Function colr_set_SeedColours_02() As Collection
        Dim lstOfSeedColours As New Collection()
        '
        lstOfSeedColours.Add(RGB(255, 255, 255), "0")
        lstOfSeedColours.Add(RGB(20, 0, 52), "1")
        lstOfSeedColours.Add(RGB(108, 63, 153), "2")
        lstOfSeedColours.Add(RGB(157, 133, 190), "3")
        lstOfSeedColours.Add(RGB(204, 195, 220), "4")
        lstOfSeedColours.Add(RGB(230, 226, 238), "5")
        lstOfSeedColours.Add(RGB(123, 189, 214), "6")
        lstOfSeedColours.Add(RGB(0, 106, 159), "7")
        lstOfSeedColours.Add(RGB(159, 209, 139), "8")
        lstOfSeedColours.Add(RGB(66, 141, 82), "9")
        lstOfSeedColours.Add(RGB(56, 148, 137), "10")
        lstOfSeedColours.Add(RGB(20, 84, 188), "11")
        lstOfSeedColours.Add(RGB(50, 16, 192), "12")
        lstOfSeedColours.Add(RGB(102, 83, 243), "13")
        lstOfSeedColours.Add(RGB(93, 46, 162), "14")
        lstOfSeedColours.Add(RGB(178, 149, 249), "15")
        lstOfSeedColours.Add(RGB(204, 152, 246), "16")
        lstOfSeedColours.Add(RGB(60, 146, 148), "17")
        lstOfSeedColours.Add(RGB(52, 156, 136), "18")
        '    
        Return lstOfSeedColours
        '
    End Function
    '

    '
    Public Function colr_get_SeedColours(ByRef lstOfCustomColours As Collection) As Collection
        Dim numROws, numColumns As Integer
        Dim lst As Collection
        Dim lstOfSeedColours As New Collection()
        Dim j As Integer
        '
        lst = lstOfCustomColours("0")
        numColumns = lstOfCustomColours.Count
        numROws = lst.Count
        '
        For j = 0 To numColumns - 1
            lst = lstOfCustomColours(CStr(j))
            lstOfSeedColors.Add(lst.Item(CStr(lst.Count - 1)), CStr(j))
            '
            'Now remove the seed colour from the bottom of the column
            lst.Remove(CStr(lst.Count - 1))
        Next
        '
        Return lstOfSeedColors
        '
    End Function
    '
    ''' <summary>
    ''' This method will get the list of Seed Colours used to create the
    ''' Custom COours section of the ColorPicker
    ''' </summary>
    ''' <returns></returns>
    Public Function colr_set_SeedColoursxx() As Collection
        Dim lstOfSeedColours As New Collection()
        '
        lstOfSeedColours.Add(RGB(255, 255, 255), "0")           'FFFFFF (5)
        lstOfSeedColours.Add(RGB(0, 0, 0), "1")                 '000000 (5)
        lstOfSeedColours.Add(RGB(20, 0, 52), "2")               '140034 (5)
        lstOfSeedColours.Add(RGB(29, 29, 29), "3")              '1D1D1D (1)
        lstOfSeedColours.Add(RGB(34, 11, 65), "4")              '220B41 (3)
        lstOfSeedColours.Add(RGB(0, 32, 50), "5")               '002032 (7)
        lstOfSeedColours.Add(RGB(16, 37, 21), "6")              '102515 (6)
        lstOfSeedColours.Add(RGB(77, 77, 77), "7")              '4D4D4D (2)
        lstOfSeedColours.Add(RGB(51, 16, 99), "8")              '331063 (4)
        lstOfSeedColours.Add(RGB(0, 72, 110), "9")              '00486E (8)

        lstOfSeedColours.Add(RGB(0, 64, 22), "10")              '004016
        lstOfSeedColours.Add(RGB(125, 125, 125), "11")          '7D7D7D
        lstOfSeedColours.Add(RGB(108, 63, 153), "12")           '6C3F99
        lstOfSeedColours.Add(RGB(0, 106, 159), "13")            '006A9F
        lstOfSeedColours.Add(RGB(66, 141, 82), "14")            '428D52
        lstOfSeedColours.Add(RGB(200, 200, 200), "15")          'C8C8C8
        lstOfSeedColours.Add(RGB(157, 133, 190), "16")          '9D85BE
        lstOfSeedColours.Add(RGB(123, 189, 214), "17")          '7BBDD6
        lstOfSeedColours.Add(RGB(159, 209, 139), "18")          '9FD18B
        lstOfSeedColours.Add(RGB(229, 229, 229), "19")          'E5E5E5
        lstOfSeedColours.Add(RGB(204, 195, 220), "20")          'CCC3DC
        '
        lstOfSeedColours.Add(RGB(196, 221, 233), "21")          'C4DDE9
        lstOfSeedColours.Add(RGB(212, 231, 200), "22")          'D4E7C8
        lstOfSeedColours.Add(RGB(249, 249, 249), "23")          'F9F9F9
        lstOfSeedColours.Add(RGB(230, 226, 238), "24")          'E6E2EE
        lstOfSeedColours.Add(RGB(233, 242, 247), "25")          'E9F2F7
        lstOfSeedColours.Add(RGB(241, 247, 237), "26")          'F1F7ED


        'lstOfSeedColours.Add(RGB(125, 125, 125), "2")           '7D7D7D (5)
        'lstOfSeedColours.Add(RGB(108, 63, 153), "3")            '6C3F99 (1)
        'lstOfSeedColours.Add(RGB(0, 106, 159), "4")             '006A9F (3)
        'lstOfSeedColours.Add(RGB(66, 141, 82), "5")             '428D52 (7)
        'lstOfSeedColours.Add(RGB(255, 0, 0), "4")               'FF0000 (x)

        'lstOfSeedColours.Add(RGB(200, 200, 200), "6")           'C8C8C8 (6)
        'lstOfSeedColours.Add(RGB(157, 133, 190), "7")           '9D85BE (2)
        'lstOfSeedColours.Add(RGB(123, 189, 214), "8")           '7BBDD6 (4)
        'lstOfSeedColours.Add(RGB(159, 209, 139), "9")           '9FD18B (8)
        '
        'lstOfSeedColours.Add(RGB(255, 0, 0), "9")
        'lstOfSeedColours.Add(RGB(20, 84, 188), "10")
        'lstOfSeedColours.Add(RGB(50, 16, 192), "11")
        'lstOfSeedColours.Add(RGB(102, 83, 243), "12")
        'lstOfSeedColours.Add(RGB(93, 46, 162), "13")
        'lstOfSeedColours.Add(RGB(178, 149, 249), "14")
        'lstOfSeedColours.Add(RGB(204, 152, 246), "15")
        'lstOfSeedColours.Add(RGB(60, 146, 148), "16")
        'lstOfSeedColours.Add(RGB(52, 156, 136), "17")
        'lstOfSeedColours.Add(RGB(44, 164, 81), "18")
        '
        'We need to set the form to auto size horizontally
        'lstOfSeedColours.Add(RGB(255, 0, 0), "19")
        '
        '
        'lstOfSeedColours.Add(RGB(20, 0, 52), "0")
        'lstOfSeedColours.Add(RGB(200, 200, 200), "1")
        'lstOfSeedColours.Add(RGB(125, 125, 125), "2")
        'lstOfSeedColours.Add(RGB(157, 133, 190), "3")
        'lstOfSeedColours.Add(RGB(108, 63, 105), "4")
        'lstOfSeedColours.Add(RGB(123, 189, 214), "5")
        'lstOfSeedColours.Add(RGB(0, 106, 159), "6")
        'lstOfSeedColours.Add(RGB(159, 209, 139), "7")
        'lstOfSeedColours.Add(RGB(66, 141, 82), "8")
        'lstOfSeedColours.Add(RGB(56, 148, 137), "9")
        'lstOfSeedColours.Add(RGB(20, 84, 188), "10")
        'lstOfSeedColours.Add(RGB(50, 16, 192), "11")
        'lstOfSeedColours.Add(RGB(102, 83, 243), "12")
        'lstOfSeedColours.Add(RGB(93, 46, 162), "13")
        'lstOfSeedColours.Add(RGB(178, 149, 249), "14")
        'lstOfSeedColours.Add(RGB(204, 152, 246), "15")
        'lstOfSeedColours.Add(RGB(60, 146, 148), "16")
        'lstOfSeedColours.Add(RGB(52, 156, 136), "17")
        'lstOfSeedColours.Add(RGB(44, 164, 81), "18")
        '


        Return lstOfSeedColours
        '
    End Function
    '
    ''' <summary>
    ''' The lstOfCustomColors contains a collection made of of N columns (each of which is a collection). The
    ''' colour values are held as rgb Long Integers
    ''' </summary>
    ''' <param name="lstOfCustomColors"></param>
    ''' <param name="column"></param>
    ''' <param name="row"></param>
    Public Sub colr_modify_CustomColour(newColour As Integer, column As Integer, row As Integer, ByRef lstOfCustomColors As Collection)
        Dim lst As Collection
        Dim numColumns, numRows As Integer
        '
        'Get the frist column to determine the number of rows
        lst = lstOfCustomColors(CStr(column))
        '
        numColumns = lstOfCustomColors.Count
        numRows = lst.Count
        '
        lst.Remove(CStr(row))
        lst.Add(newColour, CStr(row))

    End Sub
    '
    Public Sub colr_modify_CustomColour(column As Integer, ByRef newColumnColors As Collection, ByRef lstOfCustomColors As Collection)
        Dim lst As Collection
        Dim numColumns, numRows As Integer
        Dim lstNewColoursSubset As New Collection()
        Dim k, rgbColor As Integer
        '
        'Get the selected column to determine the number of rows and make it available
        lst = lstOfCustomColors(CStr(column))
        '
        numColumns = lstOfCustomColors.Count
        numRows = lst.Count
        '
        For k = 0 To lst.Count - 1
            Try
                rgbColor = newColumnColors.Item(CStr(k))
                lstNewColoursSubset.Add(rgbColor, CStr(k))
            Catch ex As Exception
                rgbColor = RGB(255, 255, 255)
                lstNewColoursSubset.Add(rgbColor, CStr(k))
            End Try
        Next

        '
        lstOfCustomColors.Remove(CStr(column))
        lstOfCustomColors.Add(lstNewColoursSubset, CStr(column))
        '

    End Sub

    '
    ''' <summary>
    ''' This function will build a colour palette based on the current them. It
    ''' will return a collection of buttons so that their event handlers can be
    ''' wired in the Form that called this building function
    ''' </summary>
    ''' <param name="location"></param>
    ''' <param name="frm"></param>
    ''' <returns></returns>
    Public Function buildColorPalette_Custom(location As System.Drawing.Point, title As String, ByRef frm As System.Windows.Forms.Form) As Collection
        Dim lst, lstOfAltColours As Collection
        'Dim strip As ToolStrip
        Dim btn As ToolStripButton
        Dim left, top As Integer
        Dim rgbColor As Integer
        Dim lbl As Label
        Dim objCol As cColorObj
        Dim objGlobals As New cGlobals()
        Dim strActionType As String
        '
        lstOfAltColours = New Collection()
        Me.lstOfCustomColors = New Collection()
        Me.lstOfSeedColors = New Collection()
        '
        'strActionType = "fromSeed"
        strActionType = "customList"
        '
        Select Case strActionType
            Case "fromSeed"
                lstOfSeedColors = Me.colr_set_SeedColours()
                lstOfCustomColors = Me.getCustomColours(lstOfSeedColors)
                '
                'Modify some colours insitu
                lstOfAltColours = Me.colr_set_CustomColours_AA_02()
                'Me.colr_modify_CustomColour(0, lstOfAltColours("0"), lstOfCustomColors)
                'Me.colr_modify_CustomColour(1, lstOfAltColours("1"), lstOfCustomColors)
                '
            Case "customList"
                '
                'Must generate the seed colours from the first element of each column of the custom colours
                'and then must move every colour in each custom colour column up one (i.e. 1 to 0, 2 to 1 etc)
                'lstOfSeedColors = Me.colr_set_SeedColours()
                Me.lstOfCustomColors = Me.colr_set_CustomColours_AA_02()
                'Me.lstOfCustomColors = Me.colr_set_CustomColours_AA_10x5()
                lstOfSeedColors = Me.colr_get_SeedColours(Me.lstOfCustomColors)

        End Select
        'lstOfSeedColors = Me.colr_set_SeedColours()
        'lstOfCustomColors = Me.getCustomColours(lstOfSeedColors)

        '
        '
        lst = lstOfCustomColors("0")
        '
        'Me.colr_modify_CustomColour(RGB(255, 0, 0), 0, 0, lstOfCustomColors)

        '
        numColumns = lstOfCustomColors.Count
        numRows = lst.Count
        '
        left = location.X
        top = location.Y
        buildColorPalette_Custom = New Collection()
        '
        lbl = New Label()
        lbl.Text = title
        lbl.AutoSize = True
        lbl.Location = New System.Drawing.Point(location.X, location.Y - 15)
        frm.Controls.Add(lbl)
        '
        For column = 0 To numColumns - 1
            strip = New ToolStrip()
            strip.Dock = False
            strip.AutoSize = True
            strip.Location = New System.Drawing.Point(left + column * 18, top)
            strip.Width = 18
            strip.Height = 200
            strip.CanOverflow = True
            strip.Margin = New Padding(0, 0, 0, 0)
            '
            strip.LayoutStyle = ToolStripLayoutStyle.VerticalStackWithOverflow
            strip.BackColor = Color.Transparent
            '
            'Get the first column.. which is stored as a collection, in the collection
            lst = lstOfCustomColors(CStr(column))
            Try
                For row = 0 To numRows
                    'objColor = New cColourObj(CInt(_colorMatrix(0, column)))
                    'objColor = New cColourObj(CInt(lst(CStr(column))))

                    btn = New ToolStripButton()
                    btn.AutoSize = False
                    btn.Height = 16
                    btn.Width = 16
                    btn.Margin = New Padding(0, 0, 0, 2)
                    If row = 0 Then btn.Margin = New Padding(0, 0, 0, 8)
                    btn.BackColor = Color.DarkOrange

                    '
                    'rgbColor = CInt(objColor.lstOfVariations(CStr(row - 1)))
                    'baseColor = ColorTranslator.FromWin32(CInt(lst(CStr(row))))
                    'variationColor = Color.FromArgb(0, baseColor)

                    'strHex = Hex(ColorTranslator.ToWin32(variationColor))
                    'btn.BackColor = baseColor
                    '
                    Select Case row
                        Case 0
                            '
                            Select Case strActionType
                                Case "fromSeed", "customList"
                                    rgbColor = CInt(lstOfSeedColors(CStr(column)))
                                Case ""
                                    rgbColor = CInt(lstOfCustomColors(CStr(column))("0"))
                            End Select
                            objCol = New cColorObj(rgbColor)
                            btn.BackColor = objCol.myColour

                        Case Else
                            Select Case strActionType
                                Case "fromSeed"
                                    rgbColor = CInt(lst(CStr(row - 1)))
                                Case "customList"
                                    rgbColor = CInt(lst(CStr(row - 1)))
                            End Select

                            objCol = New cColorObj(rgbColor)
                            btn.BackColor = objCol.myColour

                    End Select
                    '
                    btn.ToolTipText = Me.getRGB_longForm(btn.BackColor)
                    '
                    strip.Items.Add(btn)
                    buildColorPalette_Custom.Add(btn)
                    '
                Next
                frm.Controls.Add(strip)
                Me.extraPaletteHeight = strip.Height
                Me.extraPaletteWidth = strip.Width * (numColumns + 1)
                '
                Me.lst_of_CustomColourButtons = buildColorPalette_Custom
                '
            Catch ex As Exception

            End Try

        Next

    End Function
    '
    Public Function buildColorPalette(location As System.Drawing.Point, title As String, ByRef frm As System.Windows.Forms.Form) As Collection
        Dim numColumns, numRows As Integer
        Dim lstOfThemeColors As Collection
        Dim strip As ToolStrip
        Dim btn As ToolStripButton
        Dim objColor As cColorObj
        Dim left, top As Integer
        Dim rgbColor As Integer
        Dim baseColor As Color
        Dim variationColor As Color
        Dim lbl As Label
        Dim strButtonTip As String
        '
        lstOfThemeColors = Me.getThemeColours()
        '
        numColumns = lstOfThemeColors.Count
        numRows = 7
        'numRows = 15
        '
        'Now save these values for later use... if we want to refresh the palette
        Me.numColumns_Theme = numColumns
        Me.numRows_Theme = numRows
        '
        left = location.X
        top = location.Y
        buildColorPalette = New Collection()
        '
        lbl = New Label()
        lbl.Text = title
        lbl.AutoSize = True
        lbl.Location = New System.Drawing.Point(location.X, location.Y - 15)
        frm.Controls.Add(lbl)
        '
        strButtonTip = ""
        '
        For column = 0 To numColumns - 1
            strip = New ToolStrip()
            'strip.Name = 
            strip.Dock = False
            strip.AutoSize = True
            strip.Location = New System.Drawing.Point(left + column * 18, top)
            strip.Width = 18
            strip.Height = 200
            strip.CanOverflow = True
            strip.Margin = New Padding(0, 0, 0, 0)
            '
            strip.LayoutStyle = ToolStripLayoutStyle.VerticalStackWithOverflow
            strip.BackColor = Color.Transparent
            '
            Me.lstOfthemeToolStrips.Add(strip, CStr(column))
            '
            For row = 0 To numRows - 1
                objColor = New cColorObj(CInt(lstOfThemeColors(CStr(column))))
                'objColor = New cColorObj(RGB(255, 0, 0))

                btn = New ToolStripButton()
                btn.AutoSize = False
                btn.Height = 16
                btn.Width = 16
                btn.Margin = New Padding(0, 0, 0, 2)
                btn.BackColor = Color.DarkOrange
                'AddHandler btn.Click, AddressOf Me.btnHandler_Colour
                'btn.OnClick(New EventArgs())
                strip.Items.Add(btn)
                If row = 0 Then
                    btn.Margin = New Padding(0, 0, 0, 8)
                    btn.BackColor = objColor.myColour
                    '
                    '******
                    strButtonTip = Me.colorPalette_Row0_ToolTips(btn, row, column)
                Else
                    Try
                        rgbColor = CInt(objColor.lstOfVariations(CStr(row - 1)))
                        baseColor = ColorTranslator.FromWin32(rgbColor)
                        variationColor = Color.FromArgb(255, baseColor)
                        btn.BackColor = variationColor
                        '
                        '*****
                        'strButtonTip = btn.BackColor.ToString()
                        strButtonTip = Me.getRGB_longForm(btn.BackColor)

                    Catch ex As Exception

                    End Try
                End If
                '
                '
                '
                'strButtonTip = "(0,0,0)"
                'If row = 0 And column = 0 Then strButtonTip = "Text/Background - Dark 1" + " " + strButtonTip
                ' If row = 0 And column = 1 Then strButtonTip = "Text/Background - Light 1" + " " + strButtonTip
                'If row = 0 And column = 2 Then strButtonTip = "Text/Background - Dark 2" + " " + strButtonTip
                'If row = 0 And column = 3 Then strButtonTip = "Text/Background - Light 2" + " " + strButtonTip
                '
                'If row = 0 And column = 4 Then strButtonTip = "Accent 1" + " " + strButtonTip
                'If row = 0 And column = 5 Then strButtonTip = "Accent 2" + " " + strButtonTip
                'If row = 0 And column = 6 Then strButtonTip = "Accent 3" + " " + strButtonTip
                'If row = 0 And column = 7 Then strButtonTip = "Accent 4" + " " + strButtonTip
                'If row = 0 And column = 8 Then strButtonTip = "Accent 5" + " " + strButtonTip
                'If row = 0 And column = 9 Then strButtonTip = "Accent 6" + " " + strButtonTip
                '
                'If row = 0 And column = 10 Then strButtonTip = "HyperLink" + " " + strButtonTip
                'If row = 0 And column = 11 Then strButtonTip = "Followed HyperLink" + " " + strButtonTip
                '
                btn.ToolTipText = strButtonTip
                buildColorPalette.Add(btn)
                strButtonTip = ""
            Next
            Me.lstOfThemeButtons = buildColorPalette
            frm.Controls.Add(strip)
            Me.extraPaletteHeight = strip.Height
            Me.extraPaletteWidth = strip.Width * (numColumns + 1)
        Next

    End Function
    '
    Public Function colorPalette_Row0_ToolTips(ByRef btn As ToolStripButton, row As Integer, column As Integer) As String
        Dim strButtonTip As String

        'strButtonTip = btn.BackColor.ToString()
        strButtonTip = "(" + Me.getRGB_longForm(btn.BackColor) + ")"
        '
        'strButtonTip = "(0,0,0)"
        If row = 0 And column = 0 Then strButtonTip = "Text/Background - Dark 1" + " " + strButtonTip
        If row = 0 And column = 1 Then strButtonTip = "Text/Background - Light 1" + " " + strButtonTip
        If row = 0 And column = 2 Then strButtonTip = "Text/Background - Dark 2" + " " + strButtonTip
        If row = 0 And column = 3 Then strButtonTip = "Text/Background - Light 2" + " " + strButtonTip
        '
        If row = 0 And column = 4 Then strButtonTip = "Accent 1" + " " + strButtonTip
        If row = 0 And column = 5 Then strButtonTip = "Accent 2" + " " + strButtonTip
        If row = 0 And column = 6 Then strButtonTip = "Accent 3" + " " + strButtonTip
        If row = 0 And column = 7 Then strButtonTip = "Accent 4" + " " + strButtonTip
        If row = 0 And column = 8 Then strButtonTip = "Accent 5" + " " + strButtonTip
        If row = 0 And column = 9 Then strButtonTip = "Accent 6" + " " + strButtonTip
        '
        If row = 0 And column = 10 Then strButtonTip = "HyperLink" + " " + strButtonTip
        If row = 0 And column = 11 Then strButtonTip = "Followed HyperLink" + " " + strButtonTip
        '
        Return strButtonTip

    End Function

    Public Function getCustomColours(ByRef lstOfSeedColours As Collection) As Collection
        Dim numSteps As Integer
        Dim objCol As cColorObj
        Dim lstOfColours As New Collection()
        'Dim lstOfSeedColours As New Collection()
        Dim j As Integer
        Dim objGlobals As New cGlobals()
        'Dim myCol As Color
        '
        numSteps = 11
        numSteps = 15
        numSteps = 8

        'lstOfSeedColours = objGlobals.getSeedColours()
        '
        'For each seed color get a Collection of Colours, which are variations of the seed
        'This goes from column 0 to column lstOfSeedColours.Count - 1
        '
        For j = 0 To lstOfSeedColours.Count - 1
            objCol = New cColorObj(CInt(lstOfSeedColours(CStr(j))))
            lstOfColours.Add(objCol.getColoursFromSeed(True, numSteps, objCol.myColour), CStr(j))
        Next
        '
        'myCol = New Color()

        'myCol = lstOfColours.Item(CStr(lstOfSeedColours.Count - 1))
        'myCol.R = 255
        '
        Return lstOfColours
        '
    End Function
    '
    ''' <summary>
    ''' This version of custom colors was written by colr_build_CustClrsVBNET from a set of
    ''' colours automatically generated from a seed... The seed colours have been placed at the end of
    ''' each column. The idea being that when building the palette from custom colours written as
    ''' a list, the seed colour row is taken from the last item
    ''' </summary>
    ''' <returns></returns>
    Public Function colr_set_CustomColours_AA_02() As Collection
        'This version of custom colors was written by colr_build_CustClrsVBNET

        Dim lst0 As New Collection()
        Dim lst1 As New Collection()
        Dim lst2 As New Collection()
        Dim lst3 As New Collection()
        Dim lst4 As New Collection()
        Dim lst5 As New Collection()
        Dim lst6 As New Collection()
        Dim lst7 As New Collection()
        Dim lst8 As New Collection()
        Dim lst9 As New Collection()
        Dim lst10 As New Collection()
        Dim lst11 As New Collection()
        Dim lst12 As New Collection()
        Dim lst13 As New Collection()
        Dim lst14 As New Collection()
        Dim lst15 As New Collection()
        Dim lst16 As New Collection()
        Dim lst17 As New Collection()
        Dim lst18 As New Collection()
        Dim lstofCustomColours As New Collection()

        lst0.Add(RGB(255, 255, 255), "0")
        lst0.Add(RGB(249, 249, 249), "1")
        lst0.Add(RGB(229, 229, 229), "2")
        lst0.Add(RGB(200, 200, 200), "3")
        lst0.Add(RGB(125, 125, 125), "4")
        lst0.Add(RGB(77, 77, 77), "5")
        lst0.Add(RGB(29, 29, 29), "6")
        lst0.Add(RGB(0, 0, 0), "7")
        lst0.Add(RGB(255, 255, 255), "8")
        lstofCustomColours.Add(lst0, "0")

        lst1.Add(RGB(230, 226, 238), "0")
        lst1.Add(RGB(204, 195, 220), "1")
        lst1.Add(RGB(157, 133, 190), "2")
        lst1.Add(RGB(161, 102, 255), "3")
        lst1.Add(RGB(129, 51, 255), "4")
        lst1.Add(RGB(108, 63, 153), "5")
        lst1.Add(RGB(51, 16, 99), "6")
        lst1.Add(RGB(34, 11, 65), "7")
        lst1.Add(RGB(20, 0, 52), "8")
        lstofCustomColours.Add(lst1, "1")

        lst2.Add(RGB(230, 219, 240), "0")
        lst2.Add(RGB(204, 183, 225), "1")
        lst2.Add(RGB(178, 147, 210), "2")
        lst2.Add(RGB(153, 110, 196), "3")
        lst2.Add(RGB(128, 74, 181), "4")
        lst2.Add(RGB(102, 59, 145), "5")
        lst2.Add(RGB(76, 45, 108), "6")
        lst2.Add(RGB(51, 30, 72), "7")
        lst2.Add(RGB(108, 63, 153), "8")
        lstofCustomColours.Add(lst2, "2")

        lst3.Add(RGB(228, 222, 237), "0")
        lst3.Add(RGB(202, 188, 220), "1")
        lst3.Add(RGB(175, 155, 202), "2")
        lst3.Add(RGB(148, 122, 184), "3")
        lst3.Add(RGB(121, 89, 166), "4")
        lst3.Add(RGB(97, 71, 133), "5")
        lst3.Add(RGB(73, 53, 100), "6")
        lst3.Add(RGB(49, 35, 67), "7")
        lst3.Add(RGB(157, 133, 190), "8")
        lstofCustomColours.Add(lst3, "3")

        lst4.Add(RGB(228, 223, 236), "0")
        lst4.Add(RGB(200, 191, 217), "1")
        lst4.Add(RGB(173, 158, 199), "2")
        lst4.Add(RGB(145, 126, 180), "3")
        lst4.Add(RGB(118, 94, 161), "4")
        lst4.Add(RGB(94, 75, 129), "5")
        lst4.Add(RGB(71, 56, 97), "6")
        lst4.Add(RGB(47, 38, 64), "7")
        lst4.Add(RGB(204, 195, 220), "8")
        lstofCustomColours.Add(lst4, "4")

        lst5.Add(RGB(227, 223, 226), "0")
        lst5.Add(RGB(200, 191, 217), "1")
        lst5.Add(RGB(172, 159, 198), "2")
        lst5.Add(RGB(144, 126, 180), "3")
        lst5.Add(RGB(116, 94, 161), "4")
        lst5.Add(RGB(93, 75, 129), "5")
        lst5.Add(RGB(70, 57, 96), "6")
        lst5.Add(RGB(47, 38, 64), "7")
        lst5.Add(RGB(230, 226, 238), "8")
        lstofCustomColours.Add(lst5, "5")

        lst6.Add(RGB(233, 242, 247), "0")
        lst6.Add(RGB(196, 221, 233), "1")
        lst6.Add(RGB(138, 197, 219), "2")
        lst6.Add(RGB(99, 177, 207), "3")
        lst6.Add(RGB(60, 158, 195), "4")
        lst6.Add(RGB(0, 106, 159), "5")
        lst6.Add(RGB(0, 72, 110), "6")
        lst6.Add(RGB(0, 32, 50), "7")
        lst6.Add(RGB(123, 189, 214), "8")
        lstofCustomColours.Add(lst6, "6")

        lst7.Add(RGB(204, 238, 255), "0")
        lst7.Add(RGB(153, 221, 255), "1")
        lst7.Add(RGB(102, 204, 255), "2")
        lst7.Add(RGB(51, 187, 255), "3")
        lst7.Add(RGB(0, 170, 255), "4")
        lst7.Add(RGB(0, 136, 204), "5")
        lst7.Add(RGB(0, 102, 153), "6")
        lst7.Add(RGB(0, 68, 102), "7")
        lst7.Add(RGB(0, 106, 159), "8")
        lstofCustomColours.Add(lst7, "7")

        lst8.Add(RGB(241, 247, 237), "0")
        lst8.Add(RGB(212, 231, 200), "1")
        lst8.Add(RGB(159, 209, 139), "2")
        lst8.Add(RGB(134, 197, 109), "3")
        lst8.Add(RGB(83, 146, 58), "4")
        lst8.Add(RGB(66, 141, 82), "5")
        lst8.Add(RGB(0, 64, 22), "6")
        lst8.Add(RGB(16, 37, 21), "7")
        lst8.Add(RGB(159, 209, 139), "8")
        lstofCustomColours.Add(lst8, "8")

        lst9.Add(RGB(220, 239, 224), "0")
        lst9.Add(RGB(186, 222, 193), "1")
        lst9.Add(RGB(151, 206, 163), "2")
        lst9.Add(RGB(116, 190, 132), "3")
        lst9.Add(RGB(81, 174, 101), "4")
        lst9.Add(RGB(65, 139, 81), "5")
        lst9.Add(RGB(49, 104, 61), "6")
        lst9.Add(RGB(33, 69, 40), "7")
        lst9.Add(RGB(66, 141, 82), "8")
        lstofCustomColours.Add(lst9, "9")

        lst10.Add(RGB(218, 241, 238), "0")
        lst10.Add(RGB(181, 227, 221), "1")
        lst10.Add(RGB(144, 213, 205), "2")
        lst10.Add(RGB(107, 199, 188), "3")
        lst10.Add(RGB(70, 185, 171), "4")
        lst10.Add(RGB(56, 148, 137), "5")
        lst10.Add(RGB(42, 111, 103), "6")
        lst10.Add(RGB(28, 74, 68), "7")
        lst10.Add(RGB(56, 148, 137), "8")
        lstofCustomColours.Add(lst10, "10")

        lst11.Add(RGB(209, 225, 250), "0")
        lst11.Add(RGB(163, 194, 245), "1")
        lst11.Add(RGB(117, 164, 240), "2")
        lst11.Add(RGB(71, 133, 235), "3")
        lst11.Add(RGB(25, 103, 230), "4")
        lst11.Add(RGB(20, 82, 184), "5")
        lst11.Add(RGB(15, 62, 138), "6")
        lst11.Add(RGB(10, 41, 92), "7")
        lst11.Add(RGB(20, 84, 188), "8")
        lstofCustomColours.Add(lst11, "11")

        lst12.Add(RGB(216, 208, 251), "0")
        lst12.Add(RGB(178, 161, 247), "1")
        lst12.Add(RGB(139, 114, 243), "2")
        lst12.Add(RGB(100, 67, 239), "3")
        lst12.Add(RGB(61, 20, 235), "4")
        lst12.Add(RGB(49, 16, 188), "5")
        lst12.Add(RGB(37, 12, 141), "6")
        lst12.Add(RGB(25, 8, 94), "7")
        lst12.Add(RGB(50, 16, 192), "8")
        lstofCustomColours.Add(lst12, "12")

        lst13.Add(RGB(213, 207, 252), "0")
        lst13.Add(RGB(170, 160, 248), "1")
        lst13.Add(RGB(128, 112, 245), "2")
        lst13.Add(RGB(85, 64, 242), "3")
        lst13.Add(RGB(43, 17, 238), "4")
        lst13.Add(RGB(34, 13, 191), "5")
        lst13.Add(RGB(26, 10, 143), "6")
        lst13.Add(RGB(17, 7, 95), "7")
        lst13.Add(RGB(102, 83, 243), "8")
        lstofCustomColours.Add(lst13, "13")

        lst14.Add(RGB(227, 215, 244), "0")
        lst14.Add(RGB(199, 176, 232), "1")
        lst14.Add(RGB(170, 136, 221), "2")
        lst14.Add(RGB(142, 96, 210), "3")
        lst14.Add(RGB(114, 56, 199), "4")
        lst14.Add(RGB(91, 45, 159), "5")
        lst14.Add(RGB(68, 34, 119), "6")
        lst14.Add(RGB(46, 23, 79), "7")
        lst14.Add(RGB(93, 46, 162), "8")
        lstofCustomColours.Add(lst14, "14")

        lst15.Add(RGB(220, 207, 252), "0")
        lst15.Add(RGB(185, 158, 250), "1")
        lst15.Add(RGB(150, 110, 247), "2")
        lst15.Add(RGB(115, 62, 244), "3")
        lst15.Add(RGB(80, 14, 241), "4")
        lst15.Add(RGB(64, 11, 193), "5")
        lst15.Add(RGB(48, 8, 145), "6")
        lst15.Add(RGB(32, 5, 97), "7")
        lst15.Add(RGB(178, 149, 249), "8")
        lstofCustomColours.Add(lst15, "15")

        lst16.Add(RGB(232, 208, 251), "0")
        lst16.Add(RGB(209, 161, 247), "1")
        lst16.Add(RGB(185, 114, 243), "2")
        lst16.Add(RGB(162, 67, 239), "3")
        lst16.Add(RGB(139, 20, 235), "4")
        lst16.Add(RGB(111, 16, 188), "5")
        lst16.Add(RGB(83, 12, 141), "6")
        lst16.Add(RGB(56, 8, 94), "7")
        lst16.Add(RGB(204, 152, 246), "8")
        lstofCustomColours.Add(lst16, "16")

        lst17.Add(RGB(219, 240, 240), "0")
        lst17.Add(RGB(182, 225, 226), "1")
        lst17.Add(RGB(146, 209, 211), "2")
        lst17.Add(RGB(110, 194, 196), "3")
        lst17.Add(RGB(74, 179, 181), "4")
        lst17.Add(RGB(59, 143, 145), "5")
        lst17.Add(RGB(44, 107, 109), "6")
        lst17.Add(RGB(29, 72, 73), "7")
        lst17.Add(RGB(60, 146, 148), "8")
        lstofCustomColours.Add(lst17, "17")

        lst18.Add(RGB(255, 200, 200), "0")
        lst18.Add(RGB(255, 175, 175), "1")
        lst18.Add(RGB(255, 150, 150), "2")
        lst18.Add(RGB(255, 125, 125), "3")
        lst18.Add(RGB(255, 100, 100), "4")
        lst18.Add(RGB(255, 75, 75), "5")
        lst18.Add(RGB(255, 50, 50), "6")
        lst18.Add(RGB(255, 25, 25), "7")
        lst18.Add(RGB(255, 0, 0), "8")
        lstofCustomColours.Add(lst18, "18")

        'lst18.Add(RGB(217, 242, 237), "0")
        'lst18.Add(RGB(178, 230, 220), "1")
        'lst18.Add(RGB(140, 217, 202), "2")
        'lst18.Add(RGB(102, 204, 184), "3")
        'lst18.Add(RGB(64, 191, 167), "4")
        'lst18.Add(RGB(51, 153, 133), "5")
        'lst18.Add(RGB(38, 115, 100), "6")
        'lst18.Add(RGB(25, 77, 67), "7")
        'lst18.Add(RGB(52, 156, 136), "8")
        'lst18.Add(RGB(255, 0, 0), "8")
        'lstofCustomColours.Add(lst18, "18")



        Return lstofCustomColours
        '
        Return lstOfCustomColours

    End Function
    '
    '
    ''' <summary>
    ''' This version of custom colors was written by colr_build_CustClrsVBNET from a set of
    ''' colours automatically generated from a seed... The seed colours have been placed at the end of
    ''' each column. The idea being that when building the palette from custom colours written as
    ''' a list, the seed colour row is taken from the last item
    ''' </summary>
    ''' <returns></returns>
    Public Function colr_set_CustomColours_AA_03() As Collection
        'This version of custom colors was written by colr_build_CustClrsVBNET
        Dim lst0 As New Collection()
        Dim lst1 As New Collection()
        Dim lst2 As New Collection()
        Dim lst3 As New Collection()
        Dim lst4 As New Collection()
        Dim lst5 As New Collection()
        Dim lst6 As New Collection()
        Dim lst7 As New Collection()
        Dim lst8 As New Collection()
        Dim lst9 As New Collection()
        Dim lst10 As New Collection()
        Dim lst11 As New Collection()
        Dim lst12 As New Collection()
        Dim lst13 As New Collection()
        Dim lst14 As New Collection()
        Dim lst15 As New Collection()
        Dim lst16 As New Collection()
        Dim lst17 As New Collection()
        Dim lst18 As New Collection()
        Dim lstofCustomColours As New Collection()

        lst0.Add(RGB(255, 255, 255), "0")
        lst0.Add(RGB(249, 249, 249), "1")
        lst0.Add(RGB(229, 229, 229), "2")
        lst0.Add(RGB(216, 216, 216), "3")
        lst0.Add(RGB(125, 125, 125), "7")
        lst0.Add(RGB(77, 77, 77), "8")
        lst0.Add(RGB(29, 29, 29), "9")
        lst0.Add(RGB(0, 0, 0), "10")
        lst0.Add(RGB(255, 255, 255), "11")
        lstofCustomColours.Add(lst0, "0")

        lst1.Add(RGB(231, 216, 255), "0")
        lst1.Add(RGB(207, 177, 255), "1")
        lst1.Add(RGB(183, 137, 255), "2")
        lst1.Add(RGB(158, 98, 255), "3")
        lst1.Add(RGB(75, 0, 196), "7")
        lst1.Add(RGB(60, 0, 157), "8")
        lst1.Add(RGB(45, 0, 118), "9")
        lst1.Add(RGB(30, 0, 78), "10")
        lst1.Add(RGB(20, 0, 52), "11")
        lstofCustomColours.Add(lst1, "1")

        lst2.Add(RGB(235, 227, 244), "0")
        lst2.Add(RGB(216, 199, 232), "1")
        lst2.Add(RGB(196, 172, 221), "2")
        lst2.Add(RGB(177, 144, 209), "3")
        lst2.Add(RGB(98, 57, 139), "7")
        lst2.Add(RGB(78, 46, 111), "8")
        lst2.Add(RGB(59, 34, 83), "9")
        lst2.Add(RGB(39, 23, 56), "10")
        lst2.Add(RGB(108, 63, 153), "11")
        lstofCustomColours.Add(lst2, "2")

        lst3.Add(RGB(234, 229, 241), "0")
        lst3.Add(RGB(214, 204, 228), "1")
        lst3.Add(RGB(193, 178, 214), "2")
        lst3.Add(RGB(173, 153, 200), "3")
        lst3.Add(RGB(93, 68, 128), "7")
        lst3.Add(RGB(75, 55, 102), "8")
        lst3.Add(RGB(56, 41, 77), "9")
        lst3.Add(RGB(37, 27, 51), "10")
        lst3.Add(RGB(157, 133, 190), "11")
        lstofCustomColours.Add(lst3, "3")

        lst4.Add(RGB(234, 230, 241), "0")
        lst4.Add(RGB(213, 205, 226), "1")
        lst4.Add(RGB(192, 181, 212), "2")
        lst4.Add(RGB(171, 156, 197), "3")
        lst4.Add(RGB(91, 72, 124), "7")
        lst4.Add(RGB(73, 58, 99), "8")
        lst4.Add(RGB(55, 43, 74), "9")
        lst4.Add(RGB(36, 29, 50), "10")
        lst4.Add(RGB(204, 195, 220), "11")
        lstofCustomColours.Add(lst4, "4")

        lst5.Add(RGB(241, 230, 240), "0")
        lst5.Add(RGB(226, 205, 225), "1")
        lst5.Add(RGB(212, 181, 210), "2")
        lst5.Add(RGB(197, 156, 194), "3")
        lst5.Add(RGB(124, 72, 120), "7")
        lst5.Add(RGB(99, 58, 96), "8")
        lst5.Add(RGB(74, 43, 72), "9")
        lst5.Add(RGB(50, 29, 48), "10")
        lst5.Add(RGB(51, 16, 99), "11")
        lstofCustomColours.Add(lst5, "5")

        lst6.Add(RGB(225, 240, 246), "0")
        lst6.Add(RGB(195, 225, 236), "1")
        lst6.Add(RGB(165, 210, 227), "2")
        lst6.Add(RGB(135, 195, 218), "3")
        lst6.Add(RGB(46, 121, 150), "7")
        lst6.Add(RGB(37, 97, 120), "8")
        lst6.Add(RGB(28, 73, 90), "9")
        lst6.Add(RGB(19, 49, 60), "10")
        lst6.Add(RGB(123, 189, 214), "11")
        lstofCustomColours.Add(lst6, "6")

        lst7.Add(RGB(216, 242, 255), "0")
        lst7.Add(RGB(177, 229, 255), "1")
        lst7.Add(RGB(137, 216, 255), "2")
        lst7.Add(RGB(98, 203, 255), "3")
        lst7.Add(RGB(0, 131, 196), "7")
        lst7.Add(RGB(0, 105, 157), "8")
        lst7.Add(RGB(0, 78, 118), "9")
        lst7.Add(RGB(0, 52, 78), "10")
        lst7.Add(RGB(0, 106, 159), "11")
        lstofCustomColours.Add(lst7, "7")

        lst8.Add(RGB(232, 244, 227), "0")
        lst8.Add(RGB(209, 233, 199), "1")
        lst8.Add(RGB(185, 222, 171), "2")
        lst8.Add(RGB(162, 210, 143), "3")
        lst8.Add(RGB(80, 140, 56), "7")
        lst8.Add(RGB(64, 112, 45), "8")
        lst8.Add(RGB(48, 84, 33), "9")
        lst8.Add(RGB(32, 56, 22), "10")
        lst8.Add(RGB(159, 209, 139), "11")
        lstofCustomColours.Add(lst8, "8")

        lst9.Add(RGB(228, 242, 231), "0")
        lst9.Add(RGB(202, 230, 208), "1")
        lst9.Add(RGB(175, 217, 184), "2")
        lst9.Add(RGB(148, 205, 160), "3")
        lst9.Add(RGB(63, 134, 78), "7")
        lst9.Add(RGB(50, 107, 62), "8")
        lst9.Add(RGB(38, 80, 47), "9")
        lst9.Add(RGB(25, 53, 31), "10")
        lst9.Add(RGB(66, 141, 82), "11")
        lstofCustomColours.Add(lst9, "9")

        lst10.Add(RGB(227, 244, 242), "0")
        lst10.Add(RGB(198, 233, 229), "1")
        lst10.Add(RGB(170, 223, 216), "2")
        lst10.Add(RGB(141, 212, 203), "3")
        lst10.Add(RGB(54, 142, 132), "7")
        lst10.Add(RGB(43, 114, 105), "8")
        lst10.Add(RGB(32, 85, 79), "9")
        lst10.Add(RGB(22, 57, 53), "10")
        lst10.Add(RGB(56, 148, 137), "11")
        lstofCustomColours.Add(lst10, "10")

        lst11.Add(RGB(220, 232, 251), "0")
        lst11.Add(RGB(184, 208, 247), "1")
        lst11.Add(RGB(149, 185, 244), "2")
        lst11.Add(RGB(113, 161, 240), "3")
        lst11.Add(RGB(19, 79, 177), "7")
        lst11.Add(RGB(15, 63, 142), "8")
        lst11.Add(RGB(11, 48, 106), "9")
        lst11.Add(RGB(8, 32, 71), "10")
        lst11.Add(RGB(20, 84, 188), "11")
        lstofCustomColours.Add(lst11, "11")

        lst12.Add(RGB(225, 219, 252), "0")
        lst12.Add(RGB(195, 183, 249), "1")
        lst12.Add(RGB(166, 146, 246), "2")
        lst12.Add(RGB(136, 110, 243), "3")
        lst12.Add(RGB(47, 15, 181), "7")
        lst12.Add(RGB(38, 12, 145), "8")
        lst12.Add(RGB(28, 9, 109), "9")
        lst12.Add(RGB(19, 6, 72), "10")
        lst12.Add(RGB(50, 16, 192), "11")
        lstofCustomColours.Add(lst12, "12")

        lst13.Add(RGB(222, 218, 252), "0")
        lst13.Add(RGB(190, 182, 250), "1")
        lst13.Add(RGB(157, 145, 247), "2")
        lst13.Add(RGB(125, 108, 245), "3")
        lst13.Add(RGB(33, 13, 183), "7")
        lst13.Add(RGB(26, 10, 147), "8")
        lst13.Add(RGB(20, 8, 110), "9")
        lst13.Add(RGB(13, 5, 73), "10")
        lst13.Add(RGB(102, 83, 243), "11")
        lstofCustomColours.Add(lst13, "13")

        lst14.Add(RGB(233, 224, 246), "0")
        lst14.Add(RGB(212, 194, 238), "1")
        lst14.Add(RGB(190, 163, 229), "2")
        lst14.Add(RGB(168, 133, 220), "3")
        lst14.Add(RGB(88, 43, 153), "7")
        lst14.Add(RGB(70, 35, 122), "8")
        lst14.Add(RGB(53, 26, 92), "9")
        lst14.Add(RGB(35, 17, 61), "10")
        lst14.Add(RGB(93, 46, 162), "11")
        lstofCustomColours.Add(lst14, "14")

        lst15.Add(RGB(228, 218, 253), "0")
        lst15.Add(RGB(201, 181, 251), "1")
        lst15.Add(RGB(174, 144, 249), "2")
        lst15.Add(RGB(147, 106, 247), "3")
        lst15.Add(RGB(61, 11, 186), "7")
        lst15.Add(RGB(49, 8, 149), "8")
        lst15.Add(RGB(37, 6, 111), "9")
        lst15.Add(RGB(25, 4, 74), "10")
        lst15.Add(RGB(178, 149, 249), "11")
        lstofCustomColours.Add(lst15, "15")

        lst16.Add(RGB(237, 219, 252), "0")
        lst16.Add(RGB(219, 183, 249), "1")
        lst16.Add(RGB(201, 147, 246), "2")
        lst16.Add(RGB(184, 111, 242), "3")
        lst16.Add(RGB(107, 16, 180), "7")
        lst16.Add(RGB(85, 13, 144), "8")
        lst16.Add(RGB(64, 9, 108), "9")
        lst16.Add(RGB(43, 6, 72), "10")
        lst16.Add(RGB(204, 152, 246), "11")
        lstofCustomColours.Add(lst16, "16")

        lst17.Add(RGB(227, 243, 244), "0")
        lst17.Add(RGB(199, 232, 232), "1")
        lst17.Add(RGB(171, 220, 221), "2")
        lst17.Add(RGB(143, 208, 210), "3")
        lst17.Add(RGB(57, 138, 140), "7")
        lst17.Add(RGB(45, 110, 112), "8")
        lst17.Add(RGB(34, 83, 84), "9")
        lst17.Add(RGB(23, 55, 56), "10")
        lst17.Add(RGB(60, 146, 148), "11")
        lstofCustomColours.Add(lst17, "17")

        lst18.Add(RGB(226, 245, 241), "0")
        lst18.Add(RGB(196, 235, 228), "1")
        lst18.Add(RGB(167, 226, 214), "2")
        lst18.Add(RGB(137, 216, 201), "3")
        lst18.Add(RGB(49, 147, 128), "7")
        lst18.Add(RGB(39, 118, 103), "8")
        lst18.Add(RGB(29, 88, 77), "9")
        lst18.Add(RGB(20, 59, 51), "10")
        lst18.Add(RGB(52, 156, 136), "11")
        lstofCustomColours.Add(lst18, "18")
        '
        Return lstofCustomColours

    End Function
    '
    '
    ''' <summary>
    ''' This version of custom colors was written by colr_build_CustClrsVBNET from a set of
    ''' colours automatically generated from a seed... The seed colours have been placed at the end of
    ''' each column. The idea being that when building the palette from custom colours written as
    ''' a list, the seed colour row is taken from the last item
    ''' </summary>
    ''' <returns></returns>
    Public Function colr_set_CustomColours_AA_10x5() As Collection
        'This version of custom colors was written by colr_build_CustClrsVBNET

        Dim lst0 As New Collection()
        Dim lst1 As New Collection()
        Dim lst2 As New Collection()
        Dim lst3 As New Collection()
        Dim lst4 As New Collection()
        Dim lst5 As New Collection()
        Dim lst6 As New Collection()
        Dim lst7 As New Collection()
        Dim lst8 As New Collection()
        Dim lst9 As New Collection()
        'Dim lst10 As New Collection()
        'Dim lst11 As New Collection()
        'Dim lst12 As New Collection()
        'Dim lst13 As New Collection()
        'Dim lst14 As New Collection()
        'Dim lst15 As New Collection()
        'Dim lst16 As New Collection()
        'Dim lst17 As New Collection()
        'Dim lst18 As New Collection()
        Dim lstofCustomColours As New Collection()
        '

        lst0.Add(RGB(29, 29, 29), "0")
        lst0.Add(RGB(77, 77, 77), "1")
        lst0.Add(RGB(99, 99, 99), "2")
        lst0.Add(RGB(110, 110, 110), "3")
        'lst5.Add(RGB(116, 94, 161), "4")
        'lst5.Add(RGB(93, 75, 129), "5")
        'lst5.Add(RGB(70, 57, 96), "6")
        'lst5.Add(RGB(47, 38, 64), "7")
        lst0.Add(RGB(0, 0, 0), "4")
        lstofCustomColours.Add(lst0, "0")

        lst1.Add(RGB(249, 249, 249), "0")
        lst1.Add(RGB(229, 229, 229), "1")
        lst1.Add(RGB(200, 200, 200), "2")
        lst1.Add(RGB(125, 125, 125), "3")
        'lst0.Add(RGB(125, 125, 125), "4")
        'lst0.Add(RGB(77, 77, 77), "5")
        'lst0.Add(RGB(29, 29, 29), "6")
        'lst0.Add(RGB(0, 0, 0), "7")
        lst1.Add(RGB(255, 255, 255), "4")
        lstofCustomColours.Add(lst1, "1")

        lst2.Add(RGB(230, 226, 238), "0")
        lst2.Add(RGB(204, 195, 220), "1")
        lst2.Add(RGB(157, 133, 190), "2")
        lst2.Add(RGB(51, 16, 99), "3")
        'lst1.Add(RGB(129, 51, 255), "4")
        'lst1.Add(RGB(108, 63, 153), "5")
        'lst1.Add(RGB(51, 16, 99), "6")
        'lst1.Add(RGB(34, 11, 65), "7")
        lst2.Add(RGB(20, 0, 52), "4")
        lstofCustomColours.Add(lst2, "2")

        lst3.Add(RGB(230, 219, 240), "0")
        lst3.Add(RGB(204, 183, 225), "1")
        lst3.Add(RGB(178, 147, 210), "2")
        lst3.Add(RGB(34, 11, 65), "3")
        'lst2.Add(RGB(128, 74, 181), "4")
        'lst2.Add(RGB(102, 59, 145), "5")
        'lst2.Add(RGB(76, 45, 108), "6")
        'lst2.Add(RGB(51, 30, 72), "7")
        lst3.Add(RGB(108, 63, 153), "4")
        lstofCustomColours.Add(lst3, "3")

        lst4.Add(RGB(228, 222, 237), "0")
        lst4.Add(RGB(202, 188, 220), "1")
        lst4.Add(RGB(175, 155, 202), "2")
        lst4.Add(RGB(148, 122, 184), "3")
        'lst3.Add(RGB(121, 89, 166), "4")
        'lst3.Add(RGB(97, 71, 133), "5")
        'lst3.Add(RGB(73, 53, 100), "6")
        'lst3.Add(RGB(49, 35, 67), "7")
        lst4.Add(RGB(157, 133, 190), "4")
        lstofCustomColours.Add(lst4, "4")

        lst5.Add(RGB(228, 223, 236), "0")
        lst5.Add(RGB(200, 191, 217), "1")
        lst5.Add(RGB(173, 158, 199), "2")
        lst5.Add(RGB(145, 126, 180), "3")
        'lst4.Add(RGB(118, 94, 161), "4")
        'lst4.Add(RGB(94, 75, 129), "5")
        'lst4.Add(RGB(71, 56, 97), "6")
        'lst4.Add(RGB(47, 38, 64), "7")
        lst5.Add(RGB(204, 195, 220), "4")
        lstofCustomColours.Add(lst5, "5")

        lst6.Add(RGB(233, 242, 247), "0")
        lst6.Add(RGB(196, 221, 233), "1")
        lst6.Add(RGB(138, 197, 219), "2")
        lst6.Add(RGB(0, 72, 110), "3")
        'lst6.Add(RGB(60, 158, 195), "4")
        'lst6.Add(RGB(0, 106, 159), "5")
        'lst6.Add(RGB(0, 72, 110), "6")
        'lst6.Add(RGB(0, 32, 50), "7")
        lst6.Add(RGB(123, 189, 214), "4")
        lstofCustomColours.Add(lst6, "6")

        lst7.Add(RGB(204, 238, 255), "0")
        lst7.Add(RGB(153, 221, 255), "1")
        lst7.Add(RGB(102, 204, 255), "2")
        lst7.Add(RGB(0, 32, 50), "3")
        'lst7.Add(RGB(0, 170, 255), "4")
        'lst7.Add(RGB(0, 136, 204), "5")
        'lst7.Add(RGB(0, 102, 153), "6")
        'lst7.Add(RGB(0, 68, 102), "7")
        lst7.Add(RGB(0, 106, 159), "4")
        lstofCustomColours.Add(lst7, "7")

        lst8.Add(RGB(241, 247, 237), "0")
        lst8.Add(RGB(212, 231, 200), "1")
        lst8.Add(RGB(134, 197, 109), "2")
        lst8.Add(RGB(0, 64, 22), "3")
        'lst8.Add(RGB(83, 146, 58), "4")
        'lst8.Add(RGB(66, 141, 82), "5")
        'lst8.Add(RGB(0, 64, 22), "6")
        'lst8.Add(RGB(16, 37, 21), "7")
        lst8.Add(RGB(159, 209, 139), "4")
        lstofCustomColours.Add(lst8, "8")

        lst9.Add(RGB(220, 239, 224), "0")
        lst9.Add(RGB(186, 222, 193), "1")
        lst9.Add(RGB(151, 206, 163), "2")
        lst9.Add(RGB(116, 190, 132), "3")
        'lst9.Add(RGB(81, 174, 101), "4")
        'lst9.Add(RGB(65, 139, 81), "5")
        'lst9.Add(RGB(49, 104, 61), "6")
        'lst9.Add(RGB(33, 69, 40), "7")
        lst9.Add(RGB(66, 141, 82), "4")
        lstofCustomColours.Add(lst9, "9")


        '
        Return lstofCustomColours

    End Function
    '

    ''' <summary>
    ''' This method will return an RGB string of the form "R,G,B"
    ''' </summary>
    ''' <param name="colourRgb"></param>
    ''' <returns></returns>
    Public Function getRGB(colourRgb As Color) As String
        Dim strRed, strGreen, strBlue As String
        '
        strRed = colourRgb.R.ToString()
        strGreen = colourRgb.G.ToString()
        strBlue = colourRgb.B.ToString()
        '
        getRGB = strRed + "," + strGreen + "," + strBlue
        '

    End Function
    '

    '
    ''' <summary>
    ''' This method will return an RGB string of the form "R=xxx, G=xxx, B=xxx"
    ''' </summary>
    ''' <param name="colourRgb"></param>
    ''' <returns></returns>
    Public Function getRGB_longForm(colourRgb As Color) As String
        Dim strRed, strGreen, strBlue As String
        '
        strRed = colourRgb.R.ToString()
        strGreen = colourRgb.G.ToString()
        strBlue = colourRgb.B.ToString()
        '
        'getRGB_longForm = "R=" + strRed + ", " + "G=" + strGreen + ", " + "B=" + strBlue
        '
        getRGB_longForm = "RGB = " + strRed + "," + strGreen + "," + strBlue


    End Function
    '
    ''' <summary>
    ''' This method will return the rgb colour as a Hex String (RRGGBB)
    ''' </summary>
    ''' <param name="colourRgb"></param>
    ''' <returns></returns>
    Public Function getRGB_Hex(colourRgb As Color) As String
        Dim strRslt As String
        '
        strRslt = ""
        strRslt = colourRgb.R.ToString("X2") + colourRgb.G.ToString("X2") + colourRgb.B.ToString("X2")
        '
        Return strRslt
    End Function
    '
    Public Function getRGB_Int(colourRgb As Color) As Integer
        '
        getRGB_Int = RGB(colourRgb.R, colourRgb.G, colourRgb.B)
        '
    End Function
    '
    Public Sub colr_set_ThemeColours(ByRef myDoc As Word.Document, strTheme As String)
        'https://learn.microsoft.com/en-us/office/vba/api/Office.ThemeColorScheme.Colors
        '
        '
        Select Case strTheme
            Case "aacBase"
                Me.colr_set_ThemeColours_AACBase(myDoc)
        End Select
        '        
        '
    End Sub
    '
    Public Sub colr_set_ThemeColours_AACBase(ByRef myDoc As Word.Document)
        'https://learn.microsoft.com/en-us/office/vba/api/Office.ThemeColorScheme.Colors
        '
        Dim thm As OfficeTheme
        Dim thm1, thm2, thm3, thm4, thm5, thm6 As ThemeColor
        Dim thm7, thm8, thm9, thm10, thm11, thm12 As ThemeColor
        '
        '
        thm = myDoc.DocumentTheme
        '
        'thm.ThemeColorScheme.GetCustomColor()
        'thm.ThemeColorScheme.GetCustomColor()
        'thm.ThemeColorScheme.Colors()
        '


        thm1 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark1)
        thm2 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight1)
        thm3 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark2)
        thm4 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight2)
        thm5 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent1)
        thm6 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent2)
        thm7 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent3)
        thm8 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent4)
        thm9 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent5)
        thm10 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent6)
        thm11 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeHyperlink)
        thm12 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink)
        '

        thm1.RGB = RGB(0, 0, 0)
        thm2.RGB = RGB(255, 255, 255)
        thm3.RGB = RGB(108, 63, 153)
        thm4.RGB = RGB(20, 0, 52)
        thm5.RGB = RGB(157, 133, 190)
        thm6.RGB = RGB(125, 125, 125)
        thm7.RGB = RGB(200, 200, 200)
        thm8.RGB = RGB(0, 106, 159)
        thm9.RGB = RGB(123, 189, 214)
        thm10.RGB = RGB(66, 141, 82)
        thm11.RGB = RGB(123, 189, 214)
        thm12.RGB = RGB(159, 209, 139)
        '
    End Sub
    '
    Public Sub colr_set_ThemeColours_AACBase_Alt(ByRef myDoc As Word.Document)
        'https://learn.microsoft.com/en-us/office/vba/api/Office.ThemeColorScheme.Colors
        '
        Dim thm As OfficeTheme
        Dim thm1, thm2, thm3, thm4, thm5, thm6 As ThemeColor
        Dim thm7, thm8, thm9, thm10, thm11, thm12 As ThemeColor
        '
        '
        thm = myDoc.DocumentTheme
        '
        'thm.ThemeColorScheme.GetCustomColor()
        'thm.ThemeColorScheme.GetCustomColor()
        'thm.ThemeColorScheme.Colors()
        '


        thm1 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark1)
        thm2 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight1)
        thm3 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark2)
        thm4 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight2)
        thm5 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent1)
        thm6 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent2)
        thm7 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent3)
        thm8 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent4)
        thm9 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent5)
        thm10 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent6)
        thm11 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeHyperlink)
        thm12 = thm.ThemeColorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink)
        '
        'thm1.RGB = Color.FromArgb(&H00, &H00, &H00).

        thm1.RGB = RGB(0, 0, 0)
        thm2.RGB = RGB(255, 255, 255)
        thm3.RGB = RGB(108, 63, 153)
        thm4.RGB = RGB(20, 0, 52)
        thm5.RGB = RGB(157, 133, 190)
        thm6.RGB = RGB(125, 125, 125)
        thm7.RGB = RGB(200, 200, 200)
        thm8.RGB = RGB(0, 106, 159)
        thm9.RGB = RGB(123, 189, 214)
        thm10.RGB = RGB(66, 141, 82)
        thm11.RGB = RGB(123, 189, 214)
        thm12.RGB = RGB(159, 209, 139)
        '
    End Sub


    '
    ''' <summary>
    ''' This method will return a collection that returns the Application's
    ''' current theme colours as RGB (i.e. 32 bit Integer in VB.NET). In VBA,
    ''' RGB Colours have to be held in type Long. The Integer in VBA is 16 bits only.
    ''' 
    ''' The colours can be access by a key that starts at '0' adn extends (generally)
    ''' to 11.. But this method is not limited to 12 items. Hence the Collection
    ''' </summary>
    ''' <returns></returns>
    Public Function getThemeColours() As Collection
        Dim themeColours As Collection
        Dim thm As OfficeTheme
        Dim colorScheme As ThemeColorScheme
        Dim thm1, thm2, thm3, thm4, thm5, thm6 As ThemeColor
        Dim thm7, thm8, thm9, thm10, thm11, thm12 As ThemeColor
        Dim transparency As Integer
        Dim objGlobals As New cGlobals()
        '
        transparency = 0
        themeColours = New Collection
        '
        thm = objGlobals.glb_get_wrdActiveDoc.DocumentTheme
        'thm = Globals.ThisAddIn.Application.ActiveDocument.DocumentTheme
        'thm = Globals.ThisDocument.Application.ActiveDocument.Theme
        colorScheme = thm.ThemeColorScheme
        'For Each thmcolor In colorScheme.Colors
        thm1 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark1)
        thm2 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight1)
        thm3 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeDark2)
        thm4 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeLight2)
        thm5 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent1)
        thm6 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent2)
        thm7 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent3)
        thm8 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent4)
        thm9 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent5)
        thm10 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeAccent6)
        thm11 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeHyperlink)
        thm12 = colorScheme.Colors(MsoThemeColorSchemeIndex.msoThemeFollowedHyperlink)
        '
        '
        'strRGB = Me.getRGB(thm1.RGB)
        'strRGB = Me.getRGB(thm2.RGB)
        'strRGB = Me.getRGB(thm3.RGB)
        'strRGB = Me.getRGB(thm4.RGB)
        'strRGB = Me.getRGB(thm5.RGB)
        'strRGB = Me.getRGB(thm6.RGB)

        '
        'Me.getARGB(255, RGB(28, 50, 76))
        'Me.getARGB(255, thm2.RGB)
        'Me.getARGB(255, thm3.RGB)
        '
        themeColours.Add(thm1.RGB, "0")
        themeColours.Add(thm2.RGB, "1")
        themeColours.Add(thm3.RGB, "2")
        themeColours.Add(thm4.RGB, "3")
        themeColours.Add(thm5.RGB, "4")
        themeColours.Add(thm6.RGB, "5")
        themeColours.Add(thm7.RGB, "6")
        themeColours.Add(thm8.RGB, "7")
        themeColours.Add(thm9.RGB, "8")
        themeColours.Add(thm10.RGB, "9")
        themeColours.Add(thm11.RGB, "10")
        themeColours.Add(thm12.RGB, "11")
        '
        getThemeColours = themeColours

        '
    End Function
    '

    Public Sub btnHandler_Colour(sender As Object, e As EventArgs)
        Dim btn As ToolStripButton
        Dim btnColor As Color
        Dim strMsg As String
        '
        strMsg = "My Colour is "
        '
        btn = sender
        btnColor = btn.BackColor
        'btnColor.ToArgb()
        strMsg = strMsg & CStr(btnColor.R) & "," & CStr(btnColor.G) & "," & CStr(btnColor.B)
        '
        MsgBox(strMsg)
    End Sub
    '
    '
    Public Sub xdoSelectedItem_Fill(strColour As String, strPattern As String, strLineWeight As String, strBorderWeight As String)
        Dim strMsg As String
        Dim objColor As cColorObj
        Dim objGlobals As New cGlobals()
        '
        'rbns = Globals.Ribbons()
        'rbn = rbns.rbn_AAExcel00
        'strLineWeight = rbn.cmBox_lineWeight.Text
        'strBorderWeight = rbn.cmBox_borderWeight.Text
        '
        strMsg = "To change the colour Of a Chart Item (e.g. Chart Background, ChartBorder, Series Fill, Series Border, etc...).You 'll need to select the chart item that you want to colour"
        '
        'Me.doSelectedItem(galry)
        Try
            'objChartMgr = New cChartMgr()
            'objSeriesMgr = New cSeriesMgr(objChartMgr)
            '
            Try
                'If Not (objSeriesMgr._doColorFillofSelectedSeriesOrPoints(strColour, strPattern, CSng(strLineWeight), CSng(strBorderWeight))) Then
                'Fill the selected Cells if can't fill the selected Chart Item
                Try
                    For Each drCell In objGlobals.glb_get_wrdSelRng.Cells
                        objColor = New cColorObj(strColour)
                        drCell.interior.Color = objColor.myColour
                    Next

                Catch ex As Exception

                    End Try
                'End If
            Catch ex As Exception
                MsgBox("The current selection is not supported")
            End Try
        Catch ex As Exception

        End Try
    End Sub
    '


End Class

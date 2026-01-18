Imports Microsoft.Office.Interop.Word

Public Class cStyles_HeadingLevels
    Inherits cStylesManager

    Public Sub New()
        MyBase.New()
    End Sub
    '

    Public Sub style_getCreate_HeadingES()
        Dim docStyle As Word.Style
        Dim myDoc As Word.Document
        '
        myDoc = glb_get_wrdActiveDoc()
        '
        docStyle = Me.glb_styles_getCreate("Heading 1 (ES)", myDoc)
        Try
            docStyle.Font.Name = "Yu Gothic Medium"
        Catch ex As Exception

        End Try
        docStyle.Font.Size = 28
        docStyle.Font.Color = _glb_colour_purple_Dark
        docStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
        docStyle.ParagraphFormat.LineSpacing = 42
        docStyle.ParagraphFormat.SpaceBefore = 0
        docStyle.ParagraphFormat.SpaceAfter = 14
        docStyle.NextParagraphStyle = glb_styles_getCreate("Introduction", myDoc)
        '
        '
        'Heading 2
        docStyle = Me.glb_styles_getCreate("Heading 2 (ES)", myDoc)
        Try
            docStyle.Font.Name = "Yu Gothic Medium"
        Catch ex As Exception

        End Try
        docStyle.Font.Size = 18
        docStyle.Font.Bold = False
        docStyle.Font.Color = _glb_colour_purple_Mid
        docStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
        docStyle.ParagraphFormat.LineSpacing = 27
        docStyle.ParagraphFormat.SpaceBefore = 18
        docStyle.ParagraphFormat.SpaceAfter = 19
        docStyle.ParagraphFormat.Borders.InsideLineStyle = WdLineStyle.wdLineStyleNone
        docStyle.ParagraphFormat.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone
        docStyle.NextParagraphStyle = glb_styles_getCreate("Body Text", myDoc)
        '
        '
        'Heading 3
        docStyle = Me.glb_styles_getCreate("Heading 3 (ES)", myDoc)
        Try
            docStyle.Font.Name = "Yu Gothic Medium"
        Catch ex As Exception

        End Try
        docStyle.Font.Size = 16
        docStyle.Font.Color = WdColor.wdColorAutomatic
        docStyle.Font.Bold = True
        docStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
        docStyle.ParagraphFormat.LineSpacing = 27
        docStyle.ParagraphFormat.SpaceBefore = 18
        docStyle.ParagraphFormat.SpaceAfter = 19
        docStyle.ParagraphFormat.Borders.InsideLineStyle = WdLineStyle.wdLineStyleNone
        docStyle.ParagraphFormat.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone
        docStyle.NextParagraphStyle = glb_styles_getCreate("Body Text", myDoc)
        '
        '
        'Heading 4
        docStyle = Me.glb_styles_getCreate("Heading 4 (ES)", myDoc)
        Try
            docStyle.Font.Name = "Yu Gothic Medium"
        Catch ex As Exception

        End Try
        docStyle.Font.Size = 14
        docStyle.Font.Color = WdColor.wdColorAutomatic
        docStyle.Font.Italic = True
        docStyle.Font.Bold = True
        docStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
        docStyle.ParagraphFormat.LineSpacing = 27
        docStyle.ParagraphFormat.SpaceBefore = 18
        docStyle.ParagraphFormat.SpaceAfter = 19
        docStyle.NextParagraphStyle = glb_styles_getCreate("Body Text", myDoc)
        '
        'docStyle = myDoc.Styles.Item("Heading 1 (ES)")
        ' docStyle.ParagraphFormat.SpaceBefore = 20
        'docStyle.Font.Color = WdColor.wdColorAutomatic
        '
        'docStyle = myDoc.Styles.Item("Heading 1 (AP)")
        'docStyle.ParagraphFormat.SpaceBefore = 20
        'docStyle.Font.Color = WdColor.wdColorAutomatic
        '
    End Sub


    Public Sub style_getCreate_Heading()
        Dim myDoc As Word.Document
        Dim docStyle As Word.Style
        '
        myDoc = glb_get_wrdActiveDoc()
        '
        'Heading 1
        Me.style_format_HeadingLevels_1("Heading 1 (no number)", myDoc)
        Me.style_format_HeadingLevels_1("Heading 1 (ES)", myDoc)
        Me.style_format_HeadingLevels_1("Heading 1", myDoc)
        Me.style_format_HeadingLevels_1("Heading 1 (AP)", myDoc)
        docStyle = Me.style_format_HeadingLevels_1("Heading (glossary)", myDoc)
        docStyle.ParagraphFormat.LeftIndent = 0
        '
        'Heading 2
        Me.style_format_HeadingLevels_2("Heading 2 (no number)", myDoc)
        Me.style_format_HeadingLevels_2("Heading 2 (ES)", myDoc)
        Me.style_format_HeadingLevels_2("Heading 2", myDoc)
        Me.style_format_HeadingLevels_2("Heading 2 (AP)", myDoc)
        '
        'Heading 3
        Me.style_format_HeadingLevels_3("Heading 3 (no number)", myDoc)
        Me.style_format_HeadingLevels_3("Heading 3 (ES)", myDoc)
        Me.style_format_HeadingLevels_3("Heading 3", myDoc)
        Me.style_format_HeadingLevels_3("Heading 3 (AP)", myDoc)
        '
        'Heading 4
        Me.style_format_HeadingLevels_4("Heading 4 (no number)", myDoc)
        Me.style_format_HeadingLevels_4("Heading 4 (ES)", myDoc)
        Me.style_format_HeadingLevels_4("Heading 4", myDoc)
        Me.style_format_HeadingLevels_4("Heading 4 (AP)", myDoc)
        '
        Me.style_format_HeadingLevels_Glossary("Heading (glossary)", myDoc)
        Me.style_format_HeadingLevels_LargeNumber("Heading (Chapter)", myDoc)
    End Sub
    '
    Public Sub style_format_HeadingLevels_Glossary(strStyleName_Heading1 As String, ByRef myDoc As Word.Document)
        Dim docStyle As Word.Style
        '
        docStyle = Me.glb_styles_getCreate(strStyleName_Heading1, myDoc)
        Try
            docStyle.Font.Name = "Yu Gothic Medium"
        Catch ex As Exception

        End Try
        docStyle.Font.Size = 28
        docStyle.Font.Color = _glb_colour_purple_Dark
        docStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
        docStyle.ParagraphFormat.LineSpacing = 42
        docStyle.ParagraphFormat.SpaceBefore = 0
        docStyle.ParagraphFormat.SpaceAfter = 14
        docStyle.NextParagraphStyle = glb_styles_getCreate("Introduction", myDoc)
        '
    End Sub
    '
    Public Sub style_format_HeadingLevels_LargeNumber(strStyleName_Heading1 As String, ByRef myDoc As Word.Document)
        Dim docStyle As Word.Style
        '
        docStyle = Me.glb_styles_getCreate(strStyleName_Heading1, myDoc)
        Try
            docStyle.Font.Name = "Yu Gothic Medium"
        Catch ex As Exception

        End Try
        docStyle.Font.Size = 80
        docStyle.Font.Color = _glb_colour_purple_Dark
        docStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
        docStyle.ParagraphFormat.LineSpacing = 98
        docStyle.ParagraphFormat.SpaceBefore = 0
        docStyle.ParagraphFormat.SpaceAfter = 0
        docStyle.NextParagraphStyle = glb_styles_getCreate("Body Text", myDoc)
        docStyle.ParagraphFormat.LeftIndent = 10.8
        docStyle.ParagraphFormat.RightIndent = 3.1
        '
    End Sub

    ''' <summary>
    ''' This method formats the specified style (strStyleName_Heading1) to the format
    ''' used for Heading Level 1 (e.g. Heading 1 (no number), Heading 1 (ES), Heading 1, Heading 1 (AP)).
    ''' Heading 6 (Appendix replacement for Heading 6 (AP))
    ''' </summary>
    ''' <param name="strStyleName_Heading1"></param>
    ''' <param name="myDoc"></param>
    Public Function style_format_HeadingLevels_1(strStyleName_Heading1 As String, ByRef myDoc As Word.Document) As Word.Style
        Dim docStyle As Word.Style
        '
        docStyle = Me.glb_styles_getCreate(strStyleName_Heading1, myDoc)
        Try
            docStyle.Font.Name = "Yu Gothic Medium"
        Catch ex As Exception

        End Try
        docStyle.Font.Size = 28
        docStyle.Font.Color = _glb_colour_purple_Dark
        docStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
        docStyle.ParagraphFormat.LineSpacing = 42
        docStyle.ParagraphFormat.SpaceBefore = 0
        docStyle.ParagraphFormat.SpaceAfter = 14
        docStyle.ParagraphFormat.PageBreakBefore = True
        docStyle.NextParagraphStyle = glb_styles_getCreate("Introduction", myDoc)
        '
        Return docStyle
    End Function
    '
    ''' <summary>
    ''' This method formats the specified style (strStyleName_Heading2) to the format
    ''' used for Heading Level 2 (e.g. Heading 2 (no number), Heading 2 (ES), Heading 2, Heading 2 (AP)).
    ''' Heading 7 (Appendix replacement for Heading 7 (AP))
    ''' </summary>
    ''' <param name="strStyleName_Heading2"></param>
    ''' <param name="myDoc"></param>
    Public Sub style_format_HeadingLevels_2(strStyleName_Heading2 As String, ByRef myDoc As Word.Document)
        Dim docStyle As Word.Style

        docStyle = Me.glb_styles_getCreate(strStyleName_Heading2, myDoc)
        Try
            docStyle.Font.Name = "Yu Gothic Medium"
        Catch ex As Exception

        End Try
        docStyle.Font.Size = 18
        docStyle.Font.Bold = False
        docStyle.Font.Color = _glb_colour_purple_Mid
        docStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
        docStyle.ParagraphFormat.LineSpacing = 27
        docStyle.ParagraphFormat.SpaceBefore = 18
        docStyle.ParagraphFormat.SpaceAfter = 19
        docStyle.ParagraphFormat.Borders.InsideLineStyle = WdLineStyle.wdLineStyleNone
        docStyle.ParagraphFormat.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone
        docStyle.NextParagraphStyle = glb_styles_getCreate("Body Text", myDoc)
        '
    End Sub
    '
    ''' <summary>
    ''' This method formats the specified style (strStyleName_Heading3) to the format
    ''' used for Heading Level 3 (e.g. Heading 3 (no number), Heading 3 (ES), Heading 3, Heading 3 (AP), 
    ''' Heading 8 (Appendix replacement for Heading 3 (AP))
    ''' </summary>
    ''' <param name="strStyleName_Heading3"></param>
    ''' <param name="myDoc"></param>
    Public Sub style_format_HeadingLevels_3(strStyleName_Heading3 As String, ByRef myDoc As Word.Document)
        Dim docStyle As Word.Style

        docStyle = Me.glb_styles_getCreate(strStyleName_Heading3, myDoc)
        Try
            docStyle.Font.Name = "Yu Gothic Medium"
        Catch ex As Exception

        End Try
        docStyle.Font.Size = 16
        docStyle.Font.Color = WdColor.wdColorAutomatic
        docStyle.Font.Bold = True
        docStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
        docStyle.ParagraphFormat.LineSpacing = 27
        docStyle.ParagraphFormat.SpaceBefore = 18
        docStyle.ParagraphFormat.SpaceAfter = 19
        docStyle.ParagraphFormat.Borders.InsideLineStyle = WdLineStyle.wdLineStyleNone
        docStyle.ParagraphFormat.Borders.OutsideLineStyle = WdLineStyle.wdLineStyleNone
        docStyle.NextParagraphStyle = glb_styles_getCreate("Body Text", myDoc)
        '

    End Sub
    ''' <summary>
    ''' This method formats the specified style (strStyleName_Heading4) to the format
    ''' used for Heading Level 4 (e.g. Heading 4 (no number), Heading 4 (ES), Heading 4, Heading 4 (AP), 
    ''' Heading 9 (Appendix replacement for Heading 4 (AP))
    ''' </summary>
    ''' <param name="strStyleName_Heading4"></param>
    ''' <param name="myDoc"></param>
    Public Sub style_format_HeadingLevels_4(strStyleName_Heading4 As String, ByRef myDoc As Word.Document)
        Dim docStyle As Word.Style
        '
        'docStyle = Me.glb_styles_getCreate("Heading 4", myDoc)
        docStyle = Me.glb_styles_getCreate(strStyleName_Heading4, myDoc)
        Try
            docStyle.Font.Name = "Yu Gothic Medium"
        Catch ex As Exception

        End Try
        '
        docStyle.Font.Size = 14
        docStyle.Font.Color = WdColor.wdColorAutomatic
        docStyle.Font.Italic = True
        docStyle.Font.Bold = True
        docStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceExactly
        docStyle.ParagraphFormat.LineSpacing = 27
        docStyle.ParagraphFormat.SpaceBefore = 18
        docStyle.ParagraphFormat.SpaceAfter = 19
        docStyle.NextParagraphStyle = glb_styles_getCreate("Body Text", myDoc)
        '
    End Sub


End Class

Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
''' <summary>
''' 
'''This class deals with all things related to Styles
'''Originally written in vba, some account taken for conversion, but this
'''was not a priority at the time this class was written
'''
'''Peter Mikelaitis October 2015...http://mikl.com.au
'''Ported to VB.NET Jan 2017 from version 97p21p05
'''
''' </summary>
Public Class cStylesManager
    Inherits cGlobals
    'Public objGlobals As cGlobals
    '
    Public Sub New()
        MyBase.New()
        'Me = New cGlobals()
    End Sub
    '
    ''' <summary>
    ''' This method will return true if a specific style with name stylename exists. The root
    ''' is in globals which is inherited by this class. It is only here so that I don't lose track
    ''' of the root
    ''' in myDoc
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="styleName"></param>
    ''' <returns></returns>
    Function style_style_Exists(myDoc As Word.Document, styleName As String) As Boolean
        Dim rslt As Boolean
        '
        rslt = glb_style_Exists(myDoc, styleName)
        '
        Return rslt
    End Function
    '
    Public Function style_lstLevel_setNumberingHeadings_Main(ByRef myDoc As Word.Document) As Word.ListTemplate
        Dim lstStyle As Word.Style
        Dim lstTemplate As Word.ListTemplate
        Dim lstLevel As Word.ListLevel
        '
        lstTemplate = Nothing

        lstStyle = myDoc.Styles.Item("lstStyle_Heading_Numbered")
        'lstStyle = objGlobals.glb_get_wrdActiveDoc.Styles.Item("aa Heading (no number) List")
        lstTemplate = lstStyle.ListTemplate
        lstLevel = lstTemplate.ListLevels.Item(3)
        lstLevel.NumberPosition = 0
        lstLevel.NumberFormat = "%1.%2.%3"
        lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
        lstLevel.TabPosition = 51
        lstLevel.TextPosition = 50
        lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingTab
        '
        'Relink the headings to make them stick
        Me.style_lstLevel_relinkHeadings_Main(myDoc)

        Return lstTemplate
        '
    End Function
    '
    '
    Public Function style_lstLevel_removeNumberingHeadings_Main(ByRef myDoc As Word.Document) As Word.ListTemplate
        Dim lstStyle As Word.Style
        Dim lstTemplate As Word.ListTemplate
        Dim lstLevel As Word.ListLevel
        '
        lstTemplate = Nothing

        lstStyle = myDoc.Styles.Item("lstStyle_Heading_Numbered")
        'lstStyle = objGlobals.glb_get_wrdActiveDoc.Styles.Item("aa Heading (no number) List")
        lstTemplate = lstStyle.ListTemplate
        lstLevel = lstTemplate.ListLevels.Item(3)
        lstLevel.NumberPosition = 0
        lstLevel.NumberFormat = ""
        lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
        lstLevel.TabPosition = 0
        lstLevel.TextPosition = 0
        lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingNone
        '
        'Relink the headings to make them stick
        Me.style_lstLevel_relinkHeadings_Main(myDoc)

        Return lstTemplate
        '
    End Function
    '

    '
    Public Sub style_lstLevel_relinkHeadings_Main(ByRef myDoc As Word.Document)
        Dim lstStyle As Word.Style
        Dim lstTemplate As Word.ListTemplate
        Dim lstLevel As Word.ListLevel
        Dim j As Integer
        '
        lstStyle = myDoc.Styles.Item("lstStyle_Heading_Numbered")
        'lstStyle = objGlobals.glb_get_wrdActiveDoc.Styles.Item("aa Heading (no number) List")
        lstTemplate = lstStyle.ListTemplate
        For j = 1 To 5
            lstLevel = lstTemplate.ListLevels.Item(j)
            Select Case j
                Case 1
                    lstLevel.LinkedStyle = "Heading 1"
                Case 2
                    lstLevel.LinkedStyle = "Heading 2"
                Case 3
                    lstLevel.LinkedStyle = "Heading 3"
                Case 4
                    lstLevel.LinkedStyle = "Heading 4"
                Case 5
                    lstLevel.LinkedStyle = "Heading 5"
            End Select
        Next


    End Sub
    '
    Public Sub style_extend_TemplateStyles()
        Dim myDoc As Word.Document
        '
        myDoc = glb_get_wrdActiveDoc()
        Me.style_extend_TemplateStyles(myDoc)
        '
    End Sub

    '
    ''' <summary>
    ''' Once a template is released there is ofthen a need to modify styles as a result of user experience.
    ''' This can be done by issuing a new dotx, but this is disruptive, although necessary from time to time.
    ''' During release we use this method to modify styles.. It is is inserted at the Build of a Report
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub style_extend_TemplateStyles(ByRef myDoc As Word.Document)
        Dim table_DeltaLeftIndent As Single
        Dim objlstStyles As New cStyles_ListStyles()
        Dim objHeadingLevels As New cStyles_HeadingLevels()
        Dim objStylesMgr As New cStylesManager()
        Dim myStyle As Word.Style
        '
        table_DeltaLeftIndent = 2.0
        myStyle = Nothing
        '
        '*****
#Region "dotx 12.12.45 - Deployed 20250912 for Addin"
        '
        '*** Confirmed as being incorporated in 12.12.45 
        '
        'myStyle = objlstStyles.lstStyle_modify_Appendices()
        '
        'myStyle = style_getCreateRefresh_Heading5_noNum(myDoc)
        'myStyle = style_getCreateRefresh_Heading5_ES(myDoc)
        'myStyle = style_getCreateRefresh_Heading5_AP(myDoc)
        'myStyle = style_getCreateRefresh_Heading5(myDoc)
        '
        'myStyle = style_getCreateRefresh_FooterText(myDoc)
        'myStyle = style_getCreateRefresh_pageNumber(myDoc)
        '
#End Region
        '
finis:

    End Sub
    '


    '    '
    ''' <summary>
    ''' This method will check for th existence of the 'Table text' style. If it doesn't exist
    ''' it creates it and any style it depends upon. It then adds settings/modifications to that style
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Function style_txt_getTableHeadingStyle(ByRef myDoc As Word.Document) As Word.Style
        Dim txtStyle As Word.Style
        '
        Try
            txtStyle = myDoc.Styles.Item("Table column headings")
        Catch ex As Exception
            txtStyle = myDoc.Styles.Add("Table column headings", WdStyleType.wdStyleTypeParagraph)
            Try
                txtStyle.BaseStyle = myDoc.Styles.Item("Table text")
            Catch ex2 As Exception
                txtStyle.BaseStyle = Me.style_txt_getTableTextStyle(myDoc)
            End Try
        End Try
        '
        txtStyle.Font.Bold = True
        txtStyle.Font.Color = RGB(255, 255, 255)
        txtStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple
        Me.style_set_MultipleLineSpacing(txtStyle, 0.8)
        txtStyle.ParagraphFormat.LeftIndent = 2.0
        txtStyle.ParagraphFormat.RightIndent = 4.0
        txtStyle.ParagraphFormat.SpaceBefore = 2.0
        txtStyle.ParagraphFormat.SpaceAfter = 2.0
        '
        Return txtStyle
        '
    End Function
    '
    Public Function style_txt_getTableUnitsRowStyle(ByRef myDoc As Word.Document) As Word.Style
        Dim txtStyle As Word.Style
        '
        Try
            txtStyle = myDoc.Styles.Item("Table units row")
        Catch ex As Exception
            txtStyle = myDoc.Styles.Add("Table units row", WdStyleType.wdStyleTypeParagraph)
            Try
                txtStyle.BaseStyle = myDoc.Styles.Item("Normal - no space")
            Catch ex2 As Exception
                txtStyle.BaseStyle = Me.style_txt_getNormalNoSpaceStyle(myDoc)
            End Try
        End Try
        '
        txtStyle.Font.Size = 10.0
        txtStyle.Font.Color = WdColor.wdColorAutomatic
        txtStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast
        txtStyle.ParagraphFormat.LineSpacing = 12
        txtStyle.ParagraphFormat.LeftIndent = 2.0
        txtStyle.ParagraphFormat.RightIndent = 5.95
        txtStyle.ParagraphFormat.SpaceBefore = 3.0
        txtStyle.ParagraphFormat.SpaceAfter = 0.0
        '
        txtStyle.NextParagraphStyle = txtStyle
        '
        Return txtStyle
        '
    End Function

    '
    '
    ''' <summary>
    ''' This method will check for th existence of the 'Table text' style. If it doesn't exist
    ''' it creates it and any style it depends upon. It then adds settings/modifications to that style
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Function style_txt_getTableTextStyleSmall(ByRef myDoc As Word.Document) As Word.Style
        Dim txtStyle As Word.Style
        Dim strTxt As String
        '
        Try
            txtStyle = myDoc.Styles.Item("Table text (small)")
            strTxt = txtStyle.NameLocal
        Catch ex As Exception
            txtStyle = myDoc.Styles.Add("Table text (small)", WdStyleType.wdStyleTypeParagraph)
            Try
                txtStyle.BaseStyle = myDoc.Styles.Item("Normal - no space")
            Catch ex2 As Exception
                txtStyle.BaseStyle = Me.style_txt_getNormalNoSpaceStyle(myDoc)

            End Try
            txtStyle.Font.Size = 7.5
            txtStyle.ParagraphFormat.SpaceBefore = 2.0
            txtStyle.ParagraphFormat.SpaceAfter = 0.0
            txtStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast
            txtStyle.Font.Name = "Yu Gothic Medium"
            'myStyle.Font.Size = 10.5
            'txtStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple
            txtStyle.ParagraphFormat.LineSpacing = 9.5
            'myStyle.ParagraphFormat.LineSpacing = objGlobals.glb_get_wrdActiveDoc.LinesToPoints(0.9)
            'txtStyle.ParagraphFormat.LineSpacing = 11
            txtStyle.ParagraphFormat.LeftIndent = 2.0
            txtStyle.ParagraphFormat.RightIndent = 5.95

        End Try
        '
        txtStyle.NextParagraphStyle = txtStyle
        '
        'myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple
        'myStyle.ParagraphFormat.LineSpacing = objGlobals.glb_get_wrdActiveDoc.LinesToPoints(0.9)
        '
        '
        Return txtStyle
        '
    End Function
    '
    ''' <summary>
    ''' Sets a style's line spacing to "Multiple" with the given multiplier.
    ''' Word stores multiple spacing internally as points, so we convert.
    ''' </summary>
    ''' <param name="style">The Word style to modify.</param>
    ''' <param name="multiple">The UI-style multiplier (e.g., 0.8, 1.15, 1.5).</param>
    Public Sub style_set_MultipleLineSpacing(ByRef style As Word.Style, multiple As Single)

        ' Word stores multiple spacing as: multiple × 12 points
        Dim internalPoints As Single = multiple * 12.0F

        style.ParagraphFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple
        style.ParagraphFormat.LineSpacing = internalPoints
        '
    End Sub
    '
    ''' <summary>
    ''' Looks for a Caption style. This is a inbuilt style. So we just return it unaltered.
    ''' But if we do have to add it then we add it as AA standard
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function style_txt_getTableCaptionStyle(ByRef myDoc As Word.Document) As Word.Style
        Dim txtStyle As Word.Style
        Dim strName As String
        '
        Try
            txtStyle = myDoc.Styles.Item("Caption")
            strName = txtStyle.NameLocal                                                'To force  a catch if null
        Catch ex As Exception
            txtStyle = myDoc.Styles.Add("Caption", WdStyleType.wdStyleTypeParagraph)
            Try
                txtStyle.BaseStyle = myDoc.Styles.Item("Normal - no space")
                strName = txtStyle.NameLocal                                                'To force  a catch if null
                '
            Catch ex2 As Exception
                txtStyle.BaseStyle = Me.style_txt_getNormalNoSpaceStyle(myDoc)
            End Try
            txtStyle.Font.Size = 10.5
            txtStyle.ParagraphFormat.SpaceBefore = 12.0
            txtStyle.ParagraphFormat.SpaceAfter = 4.0
            txtStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple
            Me.style_set_MultipleLineSpacing(txtStyle, 0.9)
            txtStyle.ParagraphFormat.LeftIndent = 0.0
            txtStyle.ParagraphFormat.RightIndent = 0.0
            '
        End Try
        '
        '
        txtStyle.NextParagraphStyle = txtStyle
        '
        'myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple
        'myStyle.ParagraphFormat.LineSpacing = objGlobals.glb_get_wrdActiveDoc.LinesToPoints(0.9)
        '
        '
        Return txtStyle
        '
    End Function


    '
    ''' <summary>
    ''' This method will check for th existence of the 'Table text' style. If it doesn't exist
    ''' it creates it and any style it depends upon. It then adds settings/modifications to that style
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Function style_txt_getTableTextStyle(ByRef myDoc As Word.Document) As Word.Style
        Dim txtStyle As Word.Style
        Dim strName As String
        '
        Try
            txtStyle = myDoc.Styles.Item("Table text")
            strName = txtStyle.NameLocal                                                'To force  a catch if null
        Catch ex As Exception
            txtStyle = myDoc.Styles.Add("Table text", WdStyleType.wdStyleTypeParagraph)
            Try
                txtStyle.BaseStyle = myDoc.Styles.Item("Normal - no space")
            Catch ex2 As Exception
                txtStyle.BaseStyle = Me.style_txt_getNormalNoSpaceStyle(myDoc)

            End Try
            txtStyle.Font.Size = 9.5
            txtStyle.ParagraphFormat.SpaceBefore = 3.0
            txtStyle.ParagraphFormat.SpaceAfter = 0.0
            txtStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple
            Me.style_set_MultipleLineSpacing(txtStyle, 0.8)
            'txtStyle.ParagraphFormat.LineSpacing = 0.8F * 12.0F
            'myStyle.Font.Name = "Yu Gothic Medium"
            'myStyle.Font.Size = 10.5
            'myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple
            'myStyle.ParagraphFormat.LineSpacing = objGlobals.glb_get_wrdActiveDoc.LinesToPoints(0.9)
            txtStyle.ParagraphFormat.LeftIndent = 2.0
            txtStyle.ParagraphFormat.RightIndent = 5.95

        End Try
        '
        txtStyle.NextParagraphStyle = txtStyle
        '
        'myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple
        'myStyle.ParagraphFormat.LineSpacing = objGlobals.glb_get_wrdActiveDoc.LinesToPoints(0.9)
        '
        '
        Return txtStyle
        '
    End Function
    '

    ''' <summary>
    ''' This method will check for th existence of the 'Table text' style. If it doesn't exist
    ''' it creates it. Since it depends upon 'Normal' no base styles need be created
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Function style_txt_getNormalNoSpaceStyle(ByRef myDoc As Word.Document) As Word.Style
        Dim txtStyle As Word.Style
        Dim strText As String
        '
        Try
            txtStyle = myDoc.Styles.Item("Normal - no space")
            strText = txtStyle.NameLocal                            'Force a catch if null
        Catch ex As Exception
            txtStyle = myDoc.Styles.Add("Normal - no space", WdStyleType.wdStyleTypeParagraph)
            txtStyle.BaseStyle = myDoc.Styles.Item("Normal")
            txtStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
            txtStyle.ParagraphFormat.SpaceBefore = 0
            txtStyle.ParagraphFormat.SpaceAfter = 0
            '
        End Try
        '
        txtStyle.NextParagraphStyle = txtStyle
        '
        Return txtStyle
    End Function
    '
    ''' <summary>
    ''' This method will adjust the table styles by tableLeftDelta (pts) from their position(s) as defined in the
    ''' template
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="table_DeltaLeftIndent"></param>
    ''' <returns></returns>
    Public Function style_adjust_TableStyles(ByRef myDoc As Word.Document, table_DeltaLeftIndent As Single) As Boolean
        Dim rslt As Boolean
        Dim myStyle As Word.Style
        Dim lstTmpl As ListTemplate
        Dim lstLevel As ListLevel

        '
        rslt = True
        Try
            myStyle = myDoc.Styles.Item("Source")
            myStyle.ParagraphFormat.LeftIndent = table_DeltaLeftIndent
            myStyle.ParagraphFormat.KeepWithNext = False
            '
            myStyle = myDoc.Styles.Item("Table text")
            myStyle.ParagraphFormat.LeftIndent = table_DeltaLeftIndent
            myStyle = myDoc.Styles.Item("Table text (small)")
            myStyle.ParagraphFormat.LeftIndent = table_DeltaLeftIndent
            '
            myStyle = myDoc.Styles.Item("Table Quote")
            myStyle.ParagraphFormat.LeftIndent = table_DeltaLeftIndent
            myStyle = myDoc.Styles.Item("Table Quote (small)")
            myStyle.ParagraphFormat.LeftIndent = table_DeltaLeftIndent
            '
            myStyle = myDoc.Styles.Item("Table Quote Source")
            myStyle.ParagraphFormat.LeftIndent = table_DeltaLeftIndent
            myStyle = myDoc.Styles.Item("Table Quote Source (small)")
            myStyle.ParagraphFormat.LeftIndent = table_DeltaLeftIndent
            '
            '
            myStyle = myDoc.Styles.Item("Table Quote Bullet")
            lstTmpl = myStyle.ListTemplate
            lstLevel = lstTmpl.ListLevels.Item(1)
            'lstLevel.NumberFormat = "x"
            lstLevel.NumberPosition = table_DeltaLeftIndent
            lstLevel.TextPosition = 14 + table_DeltaLeftIndent
            lstLevel.LinkedStyle() = "Table Quote Bullet"

            myStyle = myDoc.Styles.Item("Table Quote Bullet (small)")
            lstTmpl = myStyle.ListTemplate
            lstLevel = lstTmpl.ListLevels.Item(1)
            lstLevel.NumberPosition = table_DeltaLeftIndent
            lstLevel.TextPosition = 14 + table_DeltaLeftIndent
            lstLevel.LinkedStyle() = "Table Quote Bullet (small)"

            ' myStyle.ParagraphFormat.LeftIndent = tableLeftIndent
            'myStyle = myDoc.Styles.Item("Table Quote Bullet (small)")
            'myStyle.ParagraphFormat.LeftIndent = tableLeftIndent

            '
            'Table List Bullets
            myStyle = myDoc.Styles.Item("aac Table List Bullets")
            lstTmpl = myStyle.ListTemplate
            '
            lstLevel = lstTmpl.ListLevels.Item(1)
            'lstLevel.NumberFormat = "x"
            'lstLevel.NumberPosition = tableLeftIndent
            lstLevel.NumberPosition = 0 + table_DeltaLeftIndent
            'lstLevel.TextPosition = 16
            lstLevel.TextPosition = 14 + table_DeltaLeftIndent
            lstLevel.LinkedStyle() = "Table list bullet"
            '
            lstLevel = lstTmpl.ListLevels.Item(2)
            'lstLevel.NumberPosition = 16
            'lstLevel.TextPosition = 28
            lstLevel.NumberPosition = 14 + table_DeltaLeftIndent
            lstLevel.TextPosition = 26 + table_DeltaLeftIndent
            lstLevel.LinkedStyle() = "Table list bullet 2"
            '
            lstLevel = lstTmpl.ListLevels.Item(3)
            'lstLevel.NumberPosition = 28
            'lstLevel.TextPosition = 40
            lstLevel.NumberPosition = 26 + table_DeltaLeftIndent
            lstLevel.TextPosition = 38 + table_DeltaLeftIndent
            lstLevel.LinkedStyle() = "Table list bullet 3"
            '
            'Table List Numbers
            myStyle = myDoc.Styles.Item("aac Table List Numbers")
            lstTmpl = myStyle.ListTemplate
            '
            lstLevel = lstTmpl.ListLevels.Item(1)
            'lstLevel.NumberFormat = "x"
            'lstLevel.NumberPosition = tableLeftIndent
            'lstLevel.TextPosition = 16
            lstLevel.NumberPosition = 0 + table_DeltaLeftIndent
            lstLevel.TextPosition = 17 + table_DeltaLeftIndent
            lstLevel.LinkedStyle() = "Table list number"
            '
            lstLevel = lstTmpl.ListLevels.Item(2)
            'lstLevel.NumberPosition = 16
            'lstLevel.TextPosition = 28
            lstLevel.NumberPosition = 17 + table_DeltaLeftIndent
            lstLevel.TextPosition = 34 + table_DeltaLeftIndent
            lstLevel.LinkedStyle() = "Table list number 2"
            '
            lstLevel = lstTmpl.ListLevels.Item(3)
            'lstLevel.NumberPosition = 28
            'lstLevel.TextPosition = 40
            lstLevel.NumberPosition = 34 + table_DeltaLeftIndent
            lstLevel.TextPosition = 45.35 + table_DeltaLeftIndent
            lstLevel.LinkedStyle() = "Table list number 3"
            '
            '
            'Table (small) List Bullets
            myStyle = myDoc.Styles.Item("aac Table List Bullets (small)")
            lstTmpl = myStyle.ListTemplate
            '
            lstLevel = lstTmpl.ListLevels.Item(1)
            'lstLevel.NumberFormat = "x"
            'lstLevel.NumberPosition = tableLeftIndent
            lstLevel.NumberPosition = 0 + table_DeltaLeftIndent
            'lstLevel.TextPosition = 16
            lstLevel.TextPosition = 14 + table_DeltaLeftIndent
            lstLevel.LinkedStyle() = "Table list bullet (small)"
            '
            lstLevel = lstTmpl.ListLevels.Item(2)
            'lstLevel.NumberPosition = 16
            'lstLevel.TextPosition = 28
            lstLevel.NumberPosition = 14 + table_DeltaLeftIndent
            lstLevel.TextPosition = 26 + table_DeltaLeftIndent
            lstLevel.LinkedStyle() = "Table list bullet 2 (small)"
            '
            lstLevel = lstTmpl.ListLevels.Item(3)
            'lstLevel.NumberPosition = 28
            'lstLevel.TextPosition = 40
            lstLevel.NumberPosition = 26 + table_DeltaLeftIndent
            lstLevel.TextPosition = 38 + table_DeltaLeftIndent
            lstLevel.LinkedStyle() = "Table list bullet 3 (small)"
            '
            'Table List Numbers
            myStyle = myDoc.Styles.Item("aac Table List Numbers (small)")
            lstTmpl = myStyle.ListTemplate
            '
            lstLevel = lstTmpl.ListLevels.Item(1)
            'lstLevel.NumberFormat = "x"
            'lstLevel.NumberPosition = tableLeftIndent
            'lstLevel.TextPosition = 16
            lstLevel.NumberPosition = 0 + table_DeltaLeftIndent
            lstLevel.TextPosition = 17 + table_DeltaLeftIndent
            lstLevel.LinkedStyle() = "Table list number (small)"
            '
            lstLevel = lstTmpl.ListLevels.Item(2)
            'lstLevel.NumberPosition = 16
            'lstLevel.TextPosition = 28
            lstLevel.NumberPosition = 17 + table_DeltaLeftIndent
            lstLevel.TextPosition = 29 + table_DeltaLeftIndent
            lstLevel.LinkedStyle() = "Table list number 2 (small)"
            '
            lstLevel = lstTmpl.ListLevels.Item(3)
            'lstLevel.NumberPosition = 28
            'lstLevel.TextPosition = 40
            lstLevel.NumberPosition = 29 + table_DeltaLeftIndent
            lstLevel.TextPosition = 40 + table_DeltaLeftIndent
            lstLevel.LinkedStyle() = "Table list number 3 (small)"
            '
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
        '
    End Function
    '
    '
    '
    ''' <summary>
    ''' This method will get the style of a specified paragraph in a Cell. 
    ''' -   drCell              (The target cell)
    ''' -   paraNumberInCell    (The target paragraph, 1 or 2 or 3 etc)       
    ''' 
    ''' </summary>
    ''' <param name="drCell"></param>
    ''' <param name="paraNumberInCell"></param>
    Public Function getStyle(ByRef drCell As Word.Cell, paraNumberInCell As Integer) As Word.Style
        Dim para As Word.Paragraph
        Dim paraStyle As Word.Style
        Dim myDoc As Word.Document
        '
        'Get the document that hosts drCell
        '
        myDoc = drCell.Range.Document
        '
        Try
            para = drCell.Range.Paragraphs(paraNumberInCell)
            paraStyle = myDoc.Styles(para.Style.NameLocal)
        Catch ex As Exception
            paraStyle = Nothing
        End Try
        '
        Return paraStyle
    End Function
    '
    ''' <summary>
    ''' This method will find the template attached to myDoc and use the 'CopyStylesFromTemplate'
    ''' function to copy the styles of the template to the document myDoc
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Function style_copy_StylesFromTemplate(ByRef myDoc As Word.Document) As Boolean
        Dim tmpl As Word.Template
        Dim rslt As Boolean = False
        '
        Try
            tmpl = myDoc.AttachedTemplate
            myDoc.CopyStylesFromTemplate(tmpl.FullName)
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
    ''' This method will return the style with the specific name strStyleName (e.g. 'aac Table (Basic)'. If it doesn't
    ''' exits in myDoc the method wil return Nothing
    ''' </summary>
    ''' <param name="strStyleName"></param>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function style_get_style(strStyleName As String, ByRef myDoc As Word.Document) As Word.Style
        Dim styl, targetStyle As Word.Style
        '
        styl = Nothing
        targetStyle = Nothing
        '
        Try
            targetStyle = myDoc.Styles.Item(strStyleName)
        Catch ex As Exception
            targetStyle = Nothing
        End Try
        '
        Return targetStyle
    End Function '
    '
    ''' <summary>
    ''' Returns true if the method can 'get' an existing style.... or False it it had to create one.
    ''' This method will return the style with the specific name strStyleName (e.g. 'aac Table (Basic)'.
    ''' The style is returned byRef in newStyle
    ''' </summary>
    ''' <param name="strStyleName"></param>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function style_getCreate_style(strStyleName As String, ByRef myDoc As Word.Document, ByRef newStyle As Word.Style, Optional strBaseStyle As String = "Normal") As Boolean
        Dim styl, targetStyle, baseStyle As Word.Style
        Dim rslt As Boolean
        '
        rslt = False
        styl = Nothing
        targetStyle = Nothing
        '
        Try
            newStyle = myDoc.Styles.Item(strStyleName)
            rslt = True
        Catch ex As Exception
            'targetStyle = Nothing
            newStyle = myDoc.Styles.Add(strStyleName)
            baseStyle = myDoc.Styles.Item(strBaseStyle)
            newStyle.BaseStyle = myDoc.Styles.Item(strBaseStyle)
            '
            newStyle.Font.Color = baseStyle.Font.Color
            '
            rslt = False
        End Try
        '
        Return rslt
        '
    End Function
    '
    ''' <summary>
    ''' This method will force the style used for the document status (DRAFT etc). It will first delete
    ''' all (stat) water mark shapes. Then it will delet and re-establish the 'stat' style
    ''' myDoc.Styles.Item(glb_var_style_waterMark_stat) back to its default.
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub style_waterMark_stat_ResetToDefault(ByRef myDoc As Word.Document)
        Dim myStyle As Word.Style
        Dim objWaterMarks As New cWaterMarks
        '
        objWaterMarks.waterMarks_Remove("*_stat")
        Try
            myStyle = myDoc.Styles.Item(glb_var_style_waterMark_stat)
            myStyle.Delete()
        Catch ex As Exception

        End Try
        '
        myStyle = Nothing
        '
        myStyle = Me.style_getCreateRefresh_waterMark_stat(myDoc)
        '

    End Sub
    '
    ''' <summary>
    ''' ListGallery or ListTemplate test... Does application of a ListTemplate set the
    ''' ListTemplate of a range that already have a ListTemplate to 'none'.
    ''' https://learn.microsoft.com/en-us/office/vba/api/word.listtemplate
    ''' </summary>
    ''' <param name="rng"></param>
    Public Sub style_apply_ListTemplate(ByRef rng As Word.Range)
        Dim myStyle As Word.Style
        Dim lt As Word.ListTemplate
        Dim lg As Word.ListGallery
        '

        '
        myStyle = Nothing
        lg = glb_get_wrdApp.ListGalleries.Item(WdListGalleryType.wdNumberGallery)
        lt = lg.ListTemplates.Item(1)
        '
        rng.ListFormat.ApplyListTemplate(lt)
        '
    End Sub

    ''' <summary>
    ''' This method will get or create and then refresh the Security header sec style (i.e. the style used
    ''' for 'Commercial-in-Confidence' etc).. If it is createed then the base style is 'Normal'. If it already exists, it is as it is
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function style_getCreateRefresh_waterMark_sec(ByRef myDoc As Word.Document) As Word.Style
        Dim myStyle As Word.Style
        'Dim lt As Word.ListTemplate
        'Dim lg As Word.ListGallery
        '
        myStyle = Nothing
        'lg = glb_get_wrdApp.ListGalleries.Item(WdListGalleryType.wdNumberGallery)
        'lt = lg.ListTemplates.Item(1)
        '
        Try
            If Not Me.style_getCreate_style(glb_var_style_waterMark_sec, myDoc, myStyle, "Normal - no space") Then                            'Confirmed updated in 12.12.41 20240404
                'The style wasn't there, so we created it. Hence it needs to be setup
                'If it was there we don't chnage anything. Especially important when
                'opening an existing document where the user has chnaged the style settings
                'via the water mark functions
                'myStyle.ListTemplate = Nothing'

                myStyle.NextParagraphStyle = myDoc.Styles.Item(myStyle.NameLocal)
                myStyle.Font.Color = Me._glb_colour_WaterMark_Grey_sec
                myStyle.Font.Bold = False
                myStyle.Font.Size = 12.0
                '
                myStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
                myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                myStyle.ParagraphFormat.SpaceBefore = 0.0
                myStyle.ParagraphFormat.SpaceAfter = 0.0

            End If

            '
        Catch ex As Exception
            myStyle = Nothing
        End Try


        Return myStyle
        '
    End Function
    Public Function xstyle_getCreateRefresh_Table_text(ByRef myDoc As Word.Document) As Word.Style
        Dim myStyle As Word.Style
        '
        myStyle = Nothing
        '
        Try
            '
            If Not Me.style_getCreate_style("Table text", myDoc, myStyle, "Normal - no space") Then                            'Confirmed updated in 12.12.41 20240404
                'The style wasn't there, so we created it. Hence it needs to be setup
                'If it was there we don't chnage anything. Especially important when
                'opening an existing document where the user has chnaged the style settings
                'via the water mark functions
                myStyle.NextParagraphStyle = myDoc.Styles.Item(myStyle.NameLocal)
                myStyle.Font.Color = RGB(0, 1, 0)
                myStyle.Font.Size = 9.5
                myStyle.Font.Name = "Yu Gothic medium"
                '
                myStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphLeft
                myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceMultiple
                myStyle.ParagraphFormat.LineSpacing = 0.8F
                myStyle.ParagraphFormat.SpaceBefore = 3.0
                myStyle.ParagraphFormat.SpaceAfter = 0.0
                '
                myStyle.ParagraphFormat.LeftIndent = 2.0
                myStyle.ParagraphFormat.RightIndent = 5.95
            End If
            '
        Catch ex As Exception

        End Try
        '
        Return myStyle
        '
    End Function

    Public Function style_getCreateRefresh_Table_column_headings(ByRef myDoc As Word.Document) As Word.Style
        Dim myStyle As Word.Style
        '
        myStyle = Nothing
        '
        Try
            '
            If Not Me.style_getCreate_style(glb_var_style_waterMark_stat, myDoc, myStyle, "Normal - no space") Then                            'Confirmed updated in 12.12.41 20240404
                'The style wasn't there, so we created it. Hence it needs to be setup
                'If it was there we don't chnage anything. Especially important when
                'opening an existing document where the user has chnaged the style settings
                'via the water mark functions
                myStyle.NextParagraphStyle = myDoc.Styles.Item(myStyle.NameLocal)
                myStyle.Font.Color = Me._glb_colour_WaterMark_Grey_stat
                myStyle.Font.Size = 36.0
                '
                myStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                myStyle.ParagraphFormat.SpaceBefore = 0.0
                myStyle.ParagraphFormat.SpaceAfter = 0.0
            End If
            '
        Catch ex As Exception

        End Try
        '
        Return myStyle
        '
    End Function
    '
    ''' <summary>
    ''' This method will get or create and then refresh the Security stat style (i.e. the style used
    ''' for 'DRAFT' etc).. If it is cretaed then the base style is 'Normal'. If it already exists, it is as it is
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function style_getCreateRefresh_waterMark_stat(ByRef myDoc As Word.Document) As Word.Style
        Dim myStyle As Word.Style
        '
        myStyle = Nothing
        '
        Try
            '
            If Not Me.style_getCreate_style(glb_var_style_waterMark_stat, myDoc, myStyle, "Normal - no space") Then                            'Confirmed updated in 12.12.41 20240404
                'The style wasn't there, so we created it. Hence it needs to be setup
                'If it was there we don't chnage anything. Especially important when
                'opening an existing document where the user has chnaged the style settings
                'via the water mark functions
                myStyle.NextParagraphStyle = myDoc.Styles.Item(myStyle.NameLocal)
                myStyle.Font.Color = Me._glb_colour_WaterMark_Grey_stat
                myStyle.Font.Size = 36.0
                '
                myStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
                myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle
                myStyle.ParagraphFormat.SpaceBefore = 0.0
                myStyle.ParagraphFormat.SpaceAfter = 0.0
            End If
            '
        Catch ex As Exception

        End Try
        '
        Return myStyle
        '
    End Function

    '
    '
    ''' <summary>
    ''' This method will force the style used for the document security level (COMMERCIAL-IN-CONFIDENCE etc)
    ''' back to its default. It will first delete all (*_sec) water mark shapes. Then it will delete and re-establish 
    ''' the 'sec' style myDoc.Styles.Item(glb_var_style_waterMark_sec)
    ''' </summary>
    ''' <param name="myDoc"></param>
    Public Sub style_waterMark_sec_ResetToDefault(ByRef myDoc As Word.Document)
        Dim myStyle As Word.Style
        Dim objWaterMarks As New cWaterMarks
        '
        objWaterMarks.waterMarks_Remove("*_sec")
        '
        Try
            myStyle = myDoc.Styles.Item(glb_var_style_waterMark_sec)
            myStyle.Delete()
        Catch ex As Exception

        End Try
        '
        myStyle = Nothing
        '
        myStyle = Me.style_getCreateRefresh_waterMark_sec(myDoc)
        '
    End Sub
    '
    Public Function style_getCreateRefresh_lstStyle_Appendeics() As Word.Style
        Dim myStyle As Word.Style
        '
        myStyle = Nothing
        '



        Return myStyle
    End Function

    '
    ''' <summary>
    ''' This method will get or create and then refresh the Heading 5 (no number) style.
    ''' If it is createed then the base style is 'Normal'. If it already exists, it is as it is.
    ''' Heading 5, Heading 5 (ES) and Heading 5 (AP) use this as a base style
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function style_getCreateRefresh_Heading5_noNum(ByRef myDoc As Word.Document) As Word.Style
        Dim myStyle As Word.Style
        '
        myStyle = Nothing
        '
        Try
            '
            Me.style_getCreate_style("Heading 5 (no number)", myDoc, myStyle, "Normal - no space")
            'If the style is not there, create it
            If Not IsNothing(myStyle) Then
                myStyle.BaseStyle = "Normal - no space"
                myStyle.Font.Bold = True
                myStyle.Font.Italic = True
                'myStyle.Font.Color = WdColor.wdColorAutomatic
                myStyle.Font.Color = RGB(0, 1, 0)
                myStyle.NextParagraphStyle = myDoc.Styles.Item("Body Text")
                myStyle.Font.Name = "Yu Gothic"
                myStyle.Font.Size = 12
                myStyle.ParagraphFormat.SpaceBefore = 12.0
                myStyle.ParagraphFormat.SpaceAfter = 0.0
                myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast
                myStyle.ParagraphFormat.LineSpacing = 12
                myStyle.ParagraphFormat.WidowControl = True
                myStyle.ParagraphFormat.KeepWithNext = True
                myStyle.ParagraphFormat.KeepTogether = True
            End If


            '
        Catch ex As Exception
            myStyle = Nothing
        End Try

        Return myStyle
        '
    End Function
    '
    ''' <summary>
    ''' This method will get or create and then refresh the Heading 5 (ES) style.
    ''' If it is created then the base style is 'Normal'. If it already exists, it is as it is.
    ''' Heading 5 (ES), uses Heading 5 (no number) as a base style.
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function style_getCreateRefresh_Heading5_ES(ByRef myDoc As Word.Document) As Word.Style
        Dim myStyle As Word.Style
        Dim baseStyle As Word.Style
        '
        myStyle = Nothing
        '
        Try
            '
            'If the style is not there, create it
            Me.style_getCreate_style("Heading 5 (ES)", myDoc, myStyle, "Heading 5 (no number)")
            If Not IsNothing(myStyle) Then
                myStyle.BaseStyle = "Heading 5 (no number)"
                myStyle.Font.Bold = True
                myStyle.Font.Italic = True
                '
                baseStyle = myDoc.Styles.Item("Heading 5 (no number)")
                myStyle.Font.Color = baseStyle.Font.Color
                '
                myStyle.NextParagraphStyle = myDoc.Styles.Item("Body Text")
                myStyle.Font.Name = "Yu Gothic"
                myStyle.Font.Size = 12
                myStyle.ParagraphFormat.SpaceBefore = 12.0
                myStyle.ParagraphFormat.SpaceAfter = 0.0
                myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast
                myStyle.ParagraphFormat.LineSpacing = 12
                myStyle.ParagraphFormat.WidowControl = True
                myStyle.ParagraphFormat.KeepWithNext = True
                myStyle.ParagraphFormat.KeepTogether = True
                '
            End If
            '
        Catch ex As Exception
            myStyle = Nothing
        End Try
        '
        Return myStyle
    End Function
    '
    '
    ''' <summary>
    ''' This method will get or create and then refresh the Heading 5 (AP) style.
    ''' If it is created then the base style is 'Normal'. If it already exists, it is as it is.
    ''' Heading 5 (AP), uses Heading 5 (no number) as a base style.
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function style_getCreateRefresh_Heading5_AP(ByRef myDoc As Word.Document) As Word.Style
        Dim myStyle As Word.Style
        Dim baseStyle As Word.Style
        '
        myStyle = Nothing
        '
        Try
            Me.style_getCreate_style("Heading 5 (AP)", myDoc, myStyle, "Heading 5 (no number)")
            If Not IsNothing(myStyle) Then
                myStyle.BaseStyle = "Heading 5 (no number)"
                baseStyle = myDoc.Styles.Item("Heading 5 (no number)")
                '
                myStyle.Font.Bold = True
                myStyle.Font.Italic = True
                '
                myStyle.Font.Color = baseStyle.Font.Color
                myStyle.NextParagraphStyle = myDoc.Styles.Item("Body Text")
                myStyle.Font.Name = "Yu Gothic"
                myStyle.Font.Size = 12
                myStyle.ParagraphFormat.SpaceBefore = 12.0
                myStyle.ParagraphFormat.SpaceAfter = 0.0
                myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast
                myStyle.ParagraphFormat.LineSpacing = 12
                myStyle.ParagraphFormat.WidowControl = True
                myStyle.ParagraphFormat.KeepWithNext = True
                myStyle.ParagraphFormat.KeepTogether = True
                '
            End If
            '
        Catch ex As Exception
            myStyle = Nothing
        End Try
        '
        Return myStyle
    End Function
    '
    ''' <summary>
    ''' This method will get or create and then refresh the Heading 5 style.
    ''' If it is created then the base style is 'Normal'. If it already exists, it is as it is.
    ''' Heading 5, uses Heading 5 (no number) as a base style.
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function style_getCreateRefresh_Heading5(ByRef myDoc As Word.Document) As Word.Style
        Dim myStyle As Word.Style
        Dim baseStyle As Word.Style
        '
        myStyle = Nothing
        '
        Try
            'Style is built in, so just make sure it is modified
            Me.style_getCreate_style("Heading 5", myDoc, myStyle, "Heading 5 (no number)")
            If Not IsNothing(myStyle) Then
                myStyle.BaseStyle = "Heading 5 (no number)"
                baseStyle = myDoc.Styles.Item("Heading 5 (no number)")
                myStyle.Font.Bold = True
                myStyle.Font.Italic = True
                '
                myStyle.Font.Color = baseStyle.Font.Color
                myStyle.NextParagraphStyle = myDoc.Styles.Item("Body Text")
                myStyle.Font.Name = "Yu Gothic"
                myStyle.Font.Size = 12
                myStyle.ParagraphFormat.SpaceBefore = 12.0
                myStyle.ParagraphFormat.SpaceAfter = 0.0
                myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceAtLeast
                myStyle.ParagraphFormat.LineSpacing = 12
                myStyle.ParagraphFormat.WidowControl = True
                myStyle.ParagraphFormat.KeepWithNext = True
                myStyle.ParagraphFormat.KeepTogether = True
                '
            End If
            '
        Catch ex As Exception
            myStyle = Nothing
        End Try
        '
        Return myStyle
    End Function
    '
    '
    ''' <summary>
    ''' This method will get or create and then refresh the pageNumber style.
    ''' If it is createed then the base style is 'Normal - no space'. If it already exists, it is as it is.
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function style_getCreateRefresh_pageNumber(ByRef myDoc As Word.Document) As Word.Style
        Dim myStyle As Word.Style
        '
        myStyle = Nothing
        '
        Try
            '
            Me.style_getCreate_style("pageNumber", myDoc, myStyle, "Normal - no space")
            'If the style is not there, create it
            'myStyle.Font.Color = WdColor.wdColorAutomatic
            myStyle.BaseStyle = "Normal - no space"
            myStyle.Font.Color = RGB(149, 149, 149)
            '
            myStyle.NextParagraphStyle = myDoc.Styles.Item(myStyle.NameLocal)
            myStyle.Font.Name = "Yu Gothic Medium"
            myStyle.Font.Size = 10
            myStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphCenter
            myStyle.ParagraphFormat.SpaceBefore = 0.0
            myStyle.ParagraphFormat.SpaceAfter = 0.0
            myStyle.ParagraphFormat.RightIndent = 2.0
            '
            myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle

            '
        Catch ex As Exception
            myStyle = Nothing
        End Try

        Return myStyle
        '
    End Function
    '
    ''' <summary>
    ''' This method will get or create and then refresh the pageNumber style.
    ''' If it is createed then the base style is 'Normal - no space'. If it already exists, it is as it is.
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function style_getCreateRefresh_FooterText(ByRef myDoc As Word.Document) As Word.Style
        Dim myStyle As Word.Style
        '
        myStyle = Nothing
        '
        Try
            '
            Me.style_getCreate_style("Footer Text", myDoc, myStyle, "Normal - no space")
            'If the style is not there, create it
            'myStyle.Font.Color = WdColor.wdColorAutomatic
            myStyle.BaseStyle = "Normal - no space"
            myStyle.Font.Color = RGB(0, 1, 0)
            '
            myStyle.NextParagraphStyle = myDoc.Styles.Item(myStyle.NameLocal)
            myStyle.Font.Name = "Yu Gothic Medium"
            myStyle.Font.Size = 6
            myStyle.ParagraphFormat.Alignment = WdParagraphAlignment.wdAlignParagraphRight
            myStyle.ParagraphFormat.SpaceBefore = 0.0
            myStyle.ParagraphFormat.SpaceAfter = 0.0
            myStyle.ParagraphFormat.RightIndent = 10.0
            'myStyle.ParagraphFormat.RightIndent = 4.0
            '
            myStyle.ParagraphFormat.LineSpacingRule = WdLineSpacing.wdLineSpaceSingle



            '
        Catch ex As Exception
            myStyle = Nothing
        End Try

        Return myStyle
        '
    End Function
    '
    '
    '
    '
    Public Sub applyColourToSelection(strColour As String)
        Dim rng As Range
        '
        rng = Me.glb_get_wrdSel().Range
        '
        Call applyColour(strColour, rng)
    End Sub
    '
    Private Sub applyColour(strColour As String, ByRef rng As Range)
        Dim col As Long
        Dim colRng As Range
        Dim para As Paragraph
        Dim paraStyle As Style
        Dim i, k
        '
        'mySel = Me.glb_get_wrdSel()
        '
        Select Case strColour
            Case "purple"
                col = RGB(157, 133, 190)
            Case "purple_Secondary"
                col = RGB(108, 63, 153)
            Case "grey"
                col = RGB(125, 125, 125)
            Case "yellow"
                col = RGB(255, 174, 59)
            Case "white"
                col = RGB(255, 254, 255)
            Case "reset"
                'rng.Collapse (wdCollapseDirection.wdCollapseStart)
                k = rng.Paragraphs.Count
                For i = 1 To rng.Paragraphs.Count
                    para = rng.Paragraphs(i)
                    paraStyle = para.Style
                    colRng = para.Range
                    colRng.Select()
                    Me.glb_get_wrdSel().Range.Font.TextColor.RGB = paraStyle.Font.TextColor.RGB
                    Me.glb_get_wrdSel().Collapse(WdCollapseDirection.wdCollapseStart)
                Next i
                'Have to do this again.. For some reason it doesn't want to stick the
                'first time
                rng = glb_get_wrdApp.Selection.Range
                For i = 1 To rng.Paragraphs.Count
                    para = rng.Paragraphs(i)
                    paraStyle = para.Style
                    colRng = para.Range
                    colRng.Select()
                    Me.glb_get_wrdSel().Range.Font.TextColor.RGB = paraStyle.Font.TextColor.RGB
                    Me.glb_get_wrdSel().Collapse(WdCollapseDirection.wdCollapseStart)
                Next i

        End Select
        '
        glb_get_wrdApp.ScreenRefresh()
        '
        'If Me.canDoAction Then
        rng.Font.TextColor.RGB = col
        'Else
        'dlgResult = MsgBox("Text colour change is not supported in this part of the document:", vbOKOnly, "Template Message")
        'End If
    End Sub
    '
    ''' <summary>
    ''' This method will apply the style, myStyle to all of the paragraphs in
    ''' the current selection
    ''' </summary>
    ''' <param name="myStyle"></param>
    Public Sub applyStyleToSelection(myStyle As Word.Style)
        Dim rngSel As Word.Range
        Dim para As Paragraph
        '
        rngSel = Me.glb_get_wrdSel.Range
        'myDoc = rngSel.Document
        '
        rngSel.Style = myStyle
        Try
            For Each para In rngSel.Paragraphs
                'rng = para.Range
                'rng.Style = myStyle
            Next para
        Catch ex As Exception

        End Try

    End Sub

    '
    'Provide this method with the style name and it will the current
    'selection to that style.. It works at the paragraph level
    Public Sub applyStyleToSelection(strStyleName As String)
        Dim myStyle As Style
        Dim myDoc As Word.Document
        '
        myDoc = Me.glb_get_wrdActiveDoc
        Try
            myStyle = myDoc.Styles(strStyleName)
        Catch ex As Exception
            myStyle = myDoc.Styles.Item("Normal")
        End Try
        '
        Me.applyStyleToSelection(myStyle)
        '
    End Sub
    '
    Public Function applyStyleToSelection(strStyleName As String, rngForStyle As Word.Range) As Word.Range
        Dim myStyle As Style
        Dim rng As Word.Range
        Dim para As Paragraph
        Dim myDoc As Word.Document
        '
        myDoc = rngForStyle.Document
        rng = Nothing
        '
        Try
            myStyle = myDoc.Styles(strStyleName)
        Catch ex As Exception
            myStyle = myDoc.Styles.Item("Normal")
        End Try
        '
        Try
            For Each para In rngForStyle.Paragraphs
                rng = para.Range
                rng.Style = myStyle
            Next para
            '
        Catch ex As Exception

        End Try
        '
        Return rng

    End Function
    '
#Region "applyStyles_Tables"
    '
    Public Sub applyStyle_TableColumnHeadings()
        Call Me.applyStyleToSelection("Table column headings")
    End Sub
    '
    Public Sub applyStyle_TableListBullet()
        Call Me.applyStyleToSelection("Table list bullet")
    End Sub
    '
    Public Sub applyStyle_TableListBullet_small()
        Call Me.applyStyleToSelection("Table list bullet (small)")
    End Sub
    '
    Public Sub applyStyle_TableListBullet2()
        Call Me.applyStyleToSelection("Table list bullet 2")
    End Sub
    '
    Public Sub applyStyle_TableListBullet2_small()
        Call Me.applyStyleToSelection("Table list bullet 2 (small)")
    End Sub
    '
    Public Sub applyStyle_TableListBullet3()
        Call Me.applyStyleToSelection("Table list bullet 3")
    End Sub
    '
    Public Sub applyStyle_TableListBullet3_small()
        Call Me.applyStyleToSelection("Table list bullet 3 (small)")
    End Sub

    Public Sub applyStyle_TableListNumber()
        Call Me.applyStyleToSelection("Table list number")
    End Sub
    '
    Public Sub applyStyle_TableListNumber_small()
        Call Me.applyStyleToSelection("Table list number (small)")
    End Sub

    Public Sub applyStyle_TableListNumber2()
        Call Me.applyStyleToSelection("Table list number 2")
    End Sub
    '
    Public Sub applyStyle_TableListNumber2_small()
        Call Me.applyStyleToSelection("Table list number 2 (small)")
    End Sub

    Public Sub applyStyle_TableListNumber3()
        Call Me.applyStyleToSelection("Table list number 3")
    End Sub
    '
    Public Sub applyStyle_TableListNumber3_small()
        Call Me.applyStyleToSelection("Table list number 3 (small)")
    End Sub
    '
    Public Sub applyStyle_TableQuote()
        Call Me.applyStyleToSelection("Table Quote")
    End Sub
    '
    Public Sub applyStyle_TableQuote_small()
        Call Me.applyStyleToSelection("Table Quote (small)")
    End Sub

    Public Sub applyStyle_TableQuoteBullet()
        Call Me.applyStyleToSelection("Table Quote Bullet")
    End Sub

    Public Sub applyStyle_TableQuoteBullet_small()
        Call Me.applyStyleToSelection("Table Quote Bullet (small)")
    End Sub
    '
    Public Sub applyStyle_TableQuoteSource()
        Call Me.applyStyleToSelection("Table Quote Source")
    End Sub
    '
    Public Sub applyStyle_TableQuoteSource_small()
        Call Me.applyStyleToSelection("Table Quote Source (small)")
    End Sub
    '
    Public Sub applyStyle_TableSideHeading1()
        Call Me.applyStyleToSelection("Table side heading 1")
    End Sub
    '
    Public Sub applyStyle_TableSideHeading1_small()
        Call Me.applyStyleToSelection("Table side heading 1 (small)")
    End Sub
    '
    Public Sub applyStyle_TableSideHeading2()
        Call Me.applyStyleToSelection("Table side heading 2")
    End Sub
    '
    Public Sub applyStyle_TableSideHeading2_small()
        Call Me.applyStyleToSelection("Table side heading 2 (small)")
    End Sub
    '
    Public Sub applyStyle_TableText()
        Call Me.applyStyleToSelection("Table text")
    End Sub
    '
    Public Sub applyStyle_TableText_small()
        Call Me.applyStyleToSelection("Table text (small)")
    End Sub
    '
    Public Sub applyStyle_TableUnitsRow()
        Call Me.applyStyleToSelection("Table units row")
    End Sub
    '
#End Region
    '
#Region "SideNotes"

    ''' <summary>
    ''' This method will insert an 'Emphasis Box' at the current selection. The box will be aligned
    ''' 'left', 'centre' or 'right' relative to the left/right margins. The parameter strType is set
    ''' to 'left', 'centre', right' respectively.. If the 'width' is left at -1.0, then the width of the
    ''' box is 1/3 of the width between the left/right margins. The height is 154 points unless otherwise specified
    ''' </summary>
    ''' <param name="strType"></param>
    ''' <returns></returns>
    Public Function styl_insert_EmphasisBox(strType As String, Optional width As Single = -1.0, Optional height As Single = 154) As String
        'This method will insert sidenotes.. You can specifiy the type, as
        'each may use a different building block
        Dim objGlobals As New cGlobals()
        Dim rng As Word.Range
        Dim objTblMgr As New cTablesMgr()
        Dim tbl As Word.Table
        Dim objGrfxMgr As New cGraphicsMgr()
        Dim strMsg As String
        '
        strMsg = ""
        rng = Nothing
        '
        If objGlobals.glb_get_wrdSel.Tables.Count = 0 Then
            Select Case strType
                Case "left", "centre", "right"
                    'strBBName = "aac_Side_PullOutText"
                    tbl = objTblMgr.tbl_build_Table_EmphasisBox(strType, width, height)
                    rng = tbl.Range
                Case "pict"
                    tbl = objTblMgr.tbl_build_Table_PullOutPict()
                    rng = tbl.Range
            End Select

        Else
            strMsg = "A SideNote cannot be attached to a Cell of a Table or PlaceHolder"
        End If
        '
        Return strMsg
        '
    End Function
    '
    Public Sub insertSideNoteText()
        'This method inserts the standard SideNote. If the user needs to
        'alter this sidenote they can apply the various style options
        'Dim rng As Range
        Dim objGlobals As New cGlobals()
        'Dim shp As Word.Shape
        '
        Try
            'rng = insertSideNote("text")
            'If rng Is Nothing Then Exit Sub
            '
            'shp = rng.ShapeRange.Item(1)
            'rng = shp.TextFrame.TextRange
            'rng.Select()
            'objGlobals.glb_get_wrdSel.MoveEnd(WdUnits.wdCharacter, -1)
            'Globals.ThisDocument.Application.Selection.Collapse (wdCollapseDirection.wdCollapseStart)
            'rng.
            'Call objToolsMgr.findText("Insert pullout text here")
        Catch ex As Exception

        End Try

    End Sub
    '
    Public Sub insertSideNotePict()
        'This method inserts the picture SideNote. If the user needs to
        'alter this sidenote they can apply the various style options
        '
        Try
            'rng = insertSideNote("pict")
            'If rng Is Nothing Then Exit Sub
            '
            'Set shp = rng.ShapeRange.Item(1)
            'Set rng = shp.TextFrame.TextRange
            'Set tbl = rng.Tables(1)
            'Set drCell = tbl.Range.Cells(2)
            'Call drCell.Range.Select

        Catch ex As Exception

        End Try
        '
    End Sub
#End Region
    '
#Region "StyleSets"
    '
    ''' <summary>
    ''' This method will insert the standard Report StyleSet.. it is done in software
    ''' because of issues with storign and retrieving from 'AutoText' 20201121
    ''' </summary>
    ''' <returns></returns>
    Public Function styles_format_styleSetES() As Word.Range
        Dim myDoc As Word.Document
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        Dim sect As Word.Section
        '
        rng = Me.glb_get_wrdSel.Range
        sect = rng.Sections.Item(1)
        myDoc = rng.Document
        '
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'rng.Paragraphs.Add(rng)
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        rng = Me.styles_insert_StyleSetMessage(rng, "es")
        '
        para = rng.Paragraphs.Item(1)
        para.Range.Style = myDoc.Styles.Item("Heading 1 (ES)")

        para = rng.Paragraphs.Item(2)
        para.Range.Style = myDoc.Styles.Item("Introduction")
        '
        para = rng.Paragraphs.Item(3)
        para.Range.Style = myDoc.Styles.Item("Heading 2 (ES)")
        '        
        para = rng.Paragraphs.Item(4)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        para = rng.Paragraphs.Item(5)
        para.Range.Style = myDoc.Styles.Item("Heading 3 (ES)")
        '        
        para = rng.Paragraphs.Item(6)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        para = rng.Paragraphs.Item(7)
        para.Range.Style = myDoc.Styles.Item("List Bullet")
        para = rng.Paragraphs.Item(8)
        para.Range.Style = myDoc.Styles.Item("List Bullet 2")
        para = rng.Paragraphs.Item(9)
        para.Range.Style = myDoc.Styles.Item("List Bullet 2")
        'para = rng.Paragraphs.Item(10)
        'para.Range.Style = myDoc.Styles.Item("List Bullet 2")
        'para = rng.Paragraphs.Item(11)
        'para.Range.Style = myDoc.Styles.Item("List Bullet 2")
        para = rng.Paragraphs.Item(10)
        para.Range.Style = myDoc.Styles.Item("List Bullet")
        '        
        para = rng.Paragraphs.Item(11)
        para.Range.Style = myDoc.Styles.Item("Heading 4 (ES)")
        para = rng.Paragraphs.Item(12)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        para = rng.Paragraphs.Item(13)
        para.Range.Style = myDoc.Styles.Item("List Number")
        para = rng.Paragraphs.Item(14)
        para.Range.Style = myDoc.Styles.Item("List Number 2")
        para = rng.Paragraphs.Item(15)
        para.Range.Style = myDoc.Styles.Item("List Number 2")
        'para = rng.Paragraphs.Item(18)
        'para.Range.Style = myDoc.Styles.Item("List Number 2")
        'para = rng.Paragraphs.Item(19)
        'para.Range.Style = myDoc.Styles.Item("List Number 2")
        para = rng.Paragraphs.Item(16)
        para.Range.Style = myDoc.Styles.Item("List Number")
        '
        para = rng.Paragraphs.Item(17)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '
        para = rng.Paragraphs.Item(18)
        para.Range.Style = myDoc.Styles.Item("Heading 5 (ES)")
        para = rng.Paragraphs.Item(19)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        '        
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()
        '
        Return rng
    End Function
    '
    ''' <summary>
    ''' This method will insert the standard Report StyleSet.. it is done in software
    ''' because of issues with storign and retrieving from 'AutoText' 20201121
    ''' </summary>
    ''' <returns></returns>
    Public Function styles_format_styleSetRpt() As Word.Range
        Dim myDoc As Word.Document
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        '
        rng = Me.glb_get_wrdSel.Range
        myDoc = rng.Document
        '
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'rng.Paragraphs.Add(rng)
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        rng = Me.styles_insert_StyleSetMessage(rng, "rpt")
        para = rng.Paragraphs.Item(1)
        para.Range.Style = myDoc.Styles.Item("Heading 1")

        para = rng.Paragraphs.Item(2)
        para.Range.Style = myDoc.Styles.Item("Introduction")
        '
        para = rng.Paragraphs.Item(3)
        para.Range.Style = myDoc.Styles.Item("Heading 2")
        '        
        para = rng.Paragraphs.Item(4)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        para = rng.Paragraphs.Item(5)
        para.Range.Style = myDoc.Styles.Item("Heading 3")
        '        
        para = rng.Paragraphs.Item(6)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        para = rng.Paragraphs.Item(7)
        para.Range.Style = myDoc.Styles.Item("List Bullet")
        para = rng.Paragraphs.Item(8)
        para.Range.Style = myDoc.Styles.Item("List Bullet 2")
        para = rng.Paragraphs.Item(9)
        para.Range.Style = myDoc.Styles.Item("List Bullet 2")
        'para = rng.Paragraphs.Item(10)
        'para.Range.Style = myDoc.Styles.Item("List Bullet 2")
        'para = rng.Paragraphs.Item(11)
        'para.Range.Style = myDoc.Styles.Item("List Bullet 2")
        para = rng.Paragraphs.Item(10)
        para.Range.Style = myDoc.Styles.Item("List Bullet")
        '        
        para = rng.Paragraphs.Item(11)
        para.Range.Style = myDoc.Styles.Item("Heading 4")
        para = rng.Paragraphs.Item(12)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        para = rng.Paragraphs.Item(13)
        para.Range.Style = myDoc.Styles.Item("List Number")
        para = rng.Paragraphs.Item(14)
        para.Range.Style = myDoc.Styles.Item("List Number 2")
        para = rng.Paragraphs.Item(15)
        para.Range.Style = myDoc.Styles.Item("List Number 2")
        'para = rng.Paragraphs.Item(17)
        'para.Range.Style = myDoc.Styles.Item("List Number 2")
        'para = rng.Paragraphs.Item(19)
        'para.Range.Style = myDoc.Styles.Item("List Number 2")
        para = rng.Paragraphs.Item(16)
        para.Range.Style = myDoc.Styles.Item("List Number")
        '
        para = rng.Paragraphs.Item(17)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '
        '        
        para = rng.Paragraphs.Item(18)
        para.Range.Style = myDoc.Styles.Item("Heading 5")
        para = rng.Paragraphs.Item(19)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        '        
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()

        Return rng
    End Function
    '
    '
    ''' <summary>
    ''' This method will insert the standard Report StyleSet.. it is done in software
    ''' because of issues with storign and retrieving from 'AutoText' 20201121
    ''' </summary>
    ''' <returns></returns>
    Public Function styles_format_styleSet_NoNumber() As Word.Range
        Dim myDoc As Word.Document
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        '
        rng = Me.glb_get_wrdSel.Range
        myDoc = rng.Document
        '
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'rng.Paragraphs.Add(rng)
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        rng = Me.styles_insert_StyleSetMessage(rng, "noNum")
        para = rng.Paragraphs.Item(1)
        para.Range.Style = myDoc.Styles.Item("Heading 1 (no number)")

        para = rng.Paragraphs.Item(2)
        para.Range.Style = myDoc.Styles.Item("Introduction")
        '
        para = rng.Paragraphs.Item(3)
        para.Range.Style = myDoc.Styles.Item("Heading 2 (no number)")
        '        
        para = rng.Paragraphs.Item(4)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        para = rng.Paragraphs.Item(5)
        para.Range.Style = myDoc.Styles.Item("Heading 3 (no number)")
        '        
        para = rng.Paragraphs.Item(6)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        para = rng.Paragraphs.Item(7)
        para.Range.Style = myDoc.Styles.Item("List Bullet")
        para = rng.Paragraphs.Item(8)
        para.Range.Style = myDoc.Styles.Item("List Bullet 2")
        para = rng.Paragraphs.Item(9)
        para.Range.Style = myDoc.Styles.Item("List Bullet 2")
        'para = rng.Paragraphs.Item(10)
        'para.Range.Style = myDoc.Styles.Item("List Bullet 2")
        'para = rng.Paragraphs.Item(11)
        'para.Range.Style = myDoc.Styles.Item("List Bullet 2")
        para = rng.Paragraphs.Item(10)
        para.Range.Style = myDoc.Styles.Item("List Bullet")
        '        
        para = rng.Paragraphs.Item(11)
        para.Range.Style = myDoc.Styles.Item("Heading 4 (no number)")
        para = rng.Paragraphs.Item(12)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        para = rng.Paragraphs.Item(13)
        para.Range.Style = myDoc.Styles.Item("List Number")
        para = rng.Paragraphs.Item(14)
        para.Range.Style = myDoc.Styles.Item("List Number 2")
        para = rng.Paragraphs.Item(15)
        para.Range.Style = myDoc.Styles.Item("List Number 2")
        'para = rng.Paragraphs.Item(17)
        'para.Range.Style = myDoc.Styles.Item("List Number 2")
        'para = rng.Paragraphs.Item(19)
        'para.Range.Style = myDoc.Styles.Item("List Number 2")
        para = rng.Paragraphs.Item(16)
        para.Range.Style = myDoc.Styles.Item("List Number")
        '
        para = rng.Paragraphs.Item(17)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '
        '        
        para = rng.Paragraphs.Item(18)
        para.Range.Style = myDoc.Styles.Item("Heading 5 (no number)")
        para = rng.Paragraphs.Item(19)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        '        
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()

        Return rng
    End Function
    '
    '
    Public Function styles_format_styleSetAP() As Word.Range
        Dim myDoc As Word.Document
        Dim rng As Word.Range
        Dim para As Word.Paragraph
        Dim sect As Word.Section
        Dim strOrientation As String
        '
        sect = glb_get_wrdSect()
        rng = Me.glb_get_wrdSel.Range
        myDoc = rng.Document
        '
        strOrientation = "prt"
        If sect.PageSetup.Orientation = WdOrientation.wdOrientLandscape Then strOrientation = "lnd"
        '
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'rng.Paragraphs.Add(rng)
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        rng = Me.styles_insert_StyleSetMessage(rng, "ap")
        '
        para = rng.Paragraphs.Item(1)
        If _glb_doApp_as_HeadingAP Then
            para.Range.Style = myDoc.Styles.Item("Heading 1 (AP)")
        Else
            para.Range.Style = myDoc.Styles.Item("Heading 6")
        End If
        '        
        '
        para = rng.Paragraphs.Item(2)
        para.Range.Style = myDoc.Styles.Item("Introduction")
        '
        para = rng.Paragraphs.Item(3)
        If _glb_doApp_as_HeadingAP Then
            para.Range.Style = myDoc.Styles.Item("Heading 2 (AP)")
        Else
            para.Range.Style = myDoc.Styles.Item("Heading 7")
        End If
        '        
        para = rng.Paragraphs.Item(4)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        para = rng.Paragraphs.Item(5)
        If _glb_doApp_as_HeadingAP Then
            para.Range.Style = myDoc.Styles.Item("Heading 3 (AP)")
        Else
            para.Range.Style = myDoc.Styles.Item("Heading 8")
        End If
        '        
        para = rng.Paragraphs.Item(6)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        para = rng.Paragraphs.Item(7)
        para.Range.Style = myDoc.Styles.Item("List Bullet")
        para = rng.Paragraphs.Item(8)
        '
        para.Range.Style = myDoc.Styles.Item("List Bullet 2")
        para = rng.Paragraphs.Item(9)
        para.Range.Style = myDoc.Styles.Item("List Bullet 2")
        'para = rng.Paragraphs.Item(10)
        'para.Range.Style = myDoc.Styles.Item("List Bullet 2")
        'para = rng.Paragraphs.Item(11)
        'para.Range.Style = myDoc.Styles.Item("List Bullet 2")
        para = rng.Paragraphs.Item(10)
        para.Range.Style = myDoc.Styles.Item("List Bullet")
        '        
        para = rng.Paragraphs.Item(11)
        If _glb_doApp_as_HeadingAP Then
            para.Range.Style = myDoc.Styles.Item("Heading 4 (AP)")
        Else
            para.Range.Style = myDoc.Styles.Item("Heading 9")
        End If
        para = rng.Paragraphs.Item(12)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        para = rng.Paragraphs.Item(13)
        para.Range.Style = myDoc.Styles.Item("List Number")
        para = rng.Paragraphs.Item(14)
        para.Range.Style = myDoc.Styles.Item("List Number 2")
        para = rng.Paragraphs.Item(15)
        para.Range.Style = myDoc.Styles.Item("List Number 2")
        'para = rng.Paragraphs.Item(18)
        'para.Range.Style = myDoc.Styles.Item("List Number 2")
        'para = rng.Paragraphs.Item(19)
        'para.Range.Style = myDoc.Styles.Item("List Number 2")
        para = rng.Paragraphs.Item(16)
        para.Range.Style = myDoc.Styles.Item("List Number")
        '
        para = rng.Paragraphs.Item(17)
        para.Range.Style = myDoc.Styles.Item("Body Text")

        '        
        para = rng.Paragraphs.Item(18)
        para.Range.Style = myDoc.Styles.Item("Heading 5 (AP)")
        para = rng.Paragraphs.Item(19)
        para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        'para = rng.Paragraphs.Item(23)
        'para.Range.Style = myDoc.Styles.Item("Heading 6 (AP)")
        'para = rng.Paragraphs.Item(24)
        'para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()
        '
        Return rng        '
    End Function
    '
    '
    Public Function styles_insert_StartupText_ReportES(ByRef rng As Word.Range) As Word.Range
        Dim myDoc As Word.Document
        'Dim rng As Word.Range
        Dim para As Word.Paragraph
        '
        'rng = Me.glb_get_wrdSel.Range
        myDoc = rng.Document
        '
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'rng.Paragraphs.Add(rng)
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        rng = Me.styles_insert_StartupText(rng, "es")
        '
        para = rng.Paragraphs.Item(1)
        para.Range.Style = myDoc.Styles.Item("Heading 1 (ES)")

        para = rng.Paragraphs.Item(2)
        para.Range.Style = myDoc.Styles.Item("Introduction")
        '
        para = rng.Paragraphs.Item(3)
        para.Range.Style = myDoc.Styles.Item("Heading 2 (ES)")
        '        
        'para = rng.Paragraphs.Item(4)
        'para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()

        Return rng
        '
    End Function
    '
    Public Function styles_insert_StartupText_ReportBody(ByRef rng As Word.Range) As Word.Range
        Dim myDoc As Word.Document
        'Dim rng As Word.Range
        Dim para As Word.Paragraph
        '
        'rng = Me.glb_get_wrdSel.Range
        myDoc = rng.Document
        '
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'rng.Paragraphs.Add(rng)
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        rng = Me.styles_insert_StartupText(rng, "rpt")
        '
        para = rng.Paragraphs.Item(1)
        para.Range.Style = myDoc.Styles.Item("Heading 1")

        para = rng.Paragraphs.Item(2)
        para.Range.Style = myDoc.Styles.Item("Introduction")
        '
        para = rng.Paragraphs.Item(3)
        para.Range.Style = myDoc.Styles.Item("Heading 2")
        '        
        'para = rng.Paragraphs.Item(4)
        'para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.MoveEnd(WdUnits.wdCharacter, Len("Heading 1"))
        rng.Select()

        Return rng
    End Function
    Public Function styles_insert_StartupText_ReportAP(ByRef rng As Word.Range) As Word.Range
        Dim myDoc As Word.Document
        'Dim rng As Word.Range
        Dim para As Word.Paragraph
        '
        'rng = Me.glb_get_wrdSel.Range
        myDoc = rng.Document
        '
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        'rng.Paragraphs.Add(rng)
        'rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        rng = Me.styles_insert_StartupText(rng, "ap")
        '
        para = rng.Paragraphs.Item(1)
        If _glb_doApp_as_HeadingAP Then
            para.Range.Style = myDoc.Styles.Item("Heading 1 (AP)")
        Else
            para.Range.Style = myDoc.Styles.Item("Heading 6")
        End If
        '        
        '
        para = rng.Paragraphs.Item(2)
        para.Range.Style = myDoc.Styles.Item("Introduction")
        '
        para = rng.Paragraphs.Item(3)
        If _glb_doApp_as_HeadingAP Then
            para.Range.Style = myDoc.Styles.Item("Heading 2 (AP)")
        Else
            para.Range.Style = myDoc.Styles.Item("Heading 7")
        End If
        '        
        'para = rng.Paragraphs.Item(4)
        'para.Range.Style = myDoc.Styles.Item("Body Text")
        '        
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        rng.Select()

        Return rng
    End Function

    '

    Public Function styles_insert_StartupText(ByRef rng As Word.Range, strSetType As String) As Word.Range
        Dim strMsg As String
        Dim strHeading1, strHeading2, strHeading3, strHeading4 As String
        '
        strHeading1 = ""
        strHeading2 = ""
        strHeading3 = ""
        strHeading4 = ""
        '
        Select Case strSetType
            Case "es"
                strHeading1 = "ES Heading"
                strHeading2 = "ES Heading 2"
                'strHeading3 = "ES Heading 3"
                'strHeading4 = "Heading 4"
            Case "rpt"
                strHeading1 = "Heading 1"
                strHeading2 = "Heading 2"
                'strHeading3 = "Heading 3"
                'strHeading4 = "Heading 4"
            Case "ap"
                strHeading1 = "Appendix Heading 1"
                strHeading2 = "Appendix Heading 2"
                'strHeading3 = "Appendix Heading 3"
                'strHeading4 = "Appendix Heading 4"

        End Select
        '
        strMsg = strHeading1 + vbCrLf
        strMsg = strMsg + "Introduction." + vbCrLf
        strMsg = strMsg + strHeading2 + vbCrLf
        'strMsg = strMsg + "Enter body copy here" + vbCrLf
        '
        rng.Text = strMsg
        '
        Return rng
        '
    End Function
    ''' <summary>
    ''' We'll adjust the number of paras in a styles message in order to fit Landscape and or
    ''' portrait pages
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="strSetType"></param>
    ''' <returns></returns>
    Public Function styles_insert_StyleSetMessage(ByRef rng As Word.Range, strSetType As String) As Word.Range
        Dim strMsg As String
        Dim strHeading1, strHeading2, strHeading3, strHeading4, strHeading5 As String
        Dim sect As Word.Section
        '
        sect = rng.Sections.Item(1)
        '
        strHeading1 = ""
        strHeading2 = ""
        strHeading3 = ""
        strHeading4 = ""
        strHeading5 = ""
        '
        Select Case strSetType
            Case "es"
                strHeading1 = "ES Heading"
                strHeading2 = "ES Heading 2"
                strHeading3 = "ES Heading 3"
                strHeading4 = "ES Heading 4"
                strHeading5 = "ES Heading 5"
            Case "rpt"
                strHeading1 = "Heading 1"
                strHeading2 = "Heading 2"
                strHeading3 = "Heading 3"
                strHeading4 = "Heading 4"
                strHeading5 = "Heading 5"
            Case "ap"
                strHeading1 = "Appendix Heading 1"
                strHeading2 = "Appendix Heading 2"
                strHeading3 = "Appendix Heading 3"
                strHeading4 = "Appendix Heading 4"
                strHeading5 = "Appendix Heading 5"
            Case "noNum"
                strHeading1 = "Heading 1 (no number)"
                strHeading2 = "Heading 2 (no number)"
                strHeading3 = "Heading 3 (no number)"
                strHeading4 = "Heading 4 (no number)"
                strHeading5 = "Heading 5 (no number)"

        End Select
        '
        strMsg = strHeading1 + vbCrLf
        strMsg = strMsg + "Insert Introduction here..." + vbCrLf
        strMsg = strMsg + strHeading2 + vbCrLf
        strMsg = strMsg + "Enter body copy here" + vbCrLf
        strMsg = strMsg + strHeading3 + vbCrLf
        strMsg = strMsg + "Body copy" + vbCrLf
        '
        strMsg = strMsg + "List bullet item" + vbCrLf
        strMsg = strMsg + "List bullet item" + vbCrLf
        strMsg = strMsg + "List bullet item" + vbCrLf
        '
        'If sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait Then
        'strMsg = strMsg + "List bullet item" + vbCrLf
        'strMsg = strMsg + "List bullet item" + vbCrLf
        'End If
        '
        strMsg = strMsg + "List bullet item" + vbCrLf

        strMsg = strMsg + strHeading4 + vbCrLf
        strMsg = strMsg + "Enter body copy here" + vbCrLf
        '
        strMsg = strMsg + "List number item" + vbCrLf
        strMsg = strMsg + "List number item" + vbCrLf
        strMsg = strMsg + "List number item" + vbCrLf
        '
        'If sect.PageSetup.Orientation = WdOrientation.wdOrientPortrait Then
        'strMsg = strMsg + "List number item" + vbCrLf
        'strMsg = strMsg + "List number item" + vbCrLf
        'End If
        '
        strMsg = strMsg + "List number item" + vbCrLf
        '
        strMsg = strMsg + "Enter body copy here" + vbCrLf

        strMsg = strMsg + strHeading5 + vbCrLf
        strMsg = strMsg + "Enter body copy here" + vbCrLf
        '
        'strMsg = strMsg + "Heading Level 6" + vbCrLf
        'strMsg = strMsg + "Enter body copy here" + vbCrLf
        '
        rng.Text = strMsg
        '
        Return rng
    End Function
    '
    ''' <summary>
    ''' This method will insert the standard 'Box' message at the range rng. The returned
    ''' range will contain the entire inserted message
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <returns></returns>
    Public Function insertStyleSetMessage_Box(ByRef rng As Word.Range) As Word.Range
        Dim strMsg As String
        '
        '
        strMsg = "Highlight Box Heading (style is ‘Box Side Heading 1’)" + vbCrLf
        strMsg = strMsg + "Basic Box Heading (style is ‘Box Side Heading 2’)" + vbCrLf
        strMsg = strMsg + "Enter text for highlight box here (style is ‘Box Text’) " + vbCrLf
        '
        strMsg = strMsg + "Box List bullet" + vbCrLf
        strMsg = strMsg + "Box List bullet" + vbCrLf
        strMsg = strMsg + "Box List bullet" + vbCrLf
        strMsg = strMsg + "Box List bullet" + vbCrLf
        strMsg = strMsg + "Box List bullet" + vbCrLf
        strMsg = strMsg + "Box List bullet" + vbCrLf
        strMsg = strMsg + "Box List bullet" + vbCrLf


        strMsg = strMsg + "More 'Box Text' here" + vbCrLf
        '
        strMsg = strMsg + "Box List number" + vbCrLf
        strMsg = strMsg + "Box List number" + vbCrLf
        strMsg = strMsg + "Box List number" + vbCrLf
        strMsg = strMsg + "Box List number" + vbCrLf
        strMsg = strMsg + "Box List number" + vbCrLf
        strMsg = strMsg + "Box List number" + vbCrLf
        strMsg = strMsg + "Box List number" + vbCrLf
        '
        strMsg = strMsg + "More 'Box Text' here" + vbCrLf
        '
        strMsg = strMsg + "Box Quote" + vbCrLf
        strMsg = strMsg + "Box quote list bullet" + vbCrLf
        strMsg = strMsg + "Box quote list bullet" + vbCrLf
        strMsg = strMsg + "Box quote list bullet" + vbCrLf
        '
        strMsg = strMsg + "Box Quote Source"
        '
        rng.Text = strMsg
        '
        Return rng
    End Function
    '
    ''' <summary>
    ''' This method will insert the styleset example at the current
    ''' Globals.ThisDocument.Application.Selection.
    ''' </summary>
    Public Sub insertStyleSetReport_Box()
        'This method will insert the styleset example at the current
        'Globals.ThisDocument.Application.Selection.
        Dim rng As Range

        Dim myDoc As Word.Document
        Dim para As Word.Paragraph
        '
        rng = Me.glb_get_wrdSelRngAll
        myDoc = Me.glb_get_wrdActiveDoc
        '
        rng.Delete()
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
        rng = Me.insertStyleSetMessage_Box(rng)
        '
        para = rng.Paragraphs.Item(1)
        para.Range.Style = myDoc.Styles.Item("Box Side Heading 1")
        '
        para = rng.Paragraphs.Item(2)
        para.Range.Style = myDoc.Styles.Item("Box Side Heading 2")
        '        
        para = rng.Paragraphs.Item(3)
        para.Range.Style = myDoc.Styles.Item("Box Text")
        '        
        para = rng.Paragraphs.Item(4)
        para.Range.Style = myDoc.Styles.Item("Box List Bullet")
        para = rng.Paragraphs.Item(5)
        para.Range.Style = myDoc.Styles.Item("Box List Bullet 2")
        para = rng.Paragraphs.Item(6)
        para.Range.Style = myDoc.Styles.Item("Box List Bullet 3")
        para = rng.Paragraphs.Item(7)
        para.Range.Style = myDoc.Styles.Item("Box List Bullet 3")
        para = rng.Paragraphs.Item(8)
        para.Range.Style = myDoc.Styles.Item("Box List Bullet 3")
        para = rng.Paragraphs.Item(9)
        para.Range.Style = myDoc.Styles.Item("Box List Bullet 2")
        para = rng.Paragraphs.Item(10)
        para.Range.Style = myDoc.Styles.Item("Box List Bullet")
        '
        para = rng.Paragraphs.Item(11)
        para.Range.Style = myDoc.Styles.Item("Box Text")
        '        
        para = rng.Paragraphs.Item(12)
        para.Range.Style = myDoc.Styles.Item("Box List Number")
        para = rng.Paragraphs.Item(13)
        para.Range.Style = myDoc.Styles.Item("Box List Number 2")
        para = rng.Paragraphs.Item(14)
        para.Range.Style = myDoc.Styles.Item("Box List Number 3")
        para = rng.Paragraphs.Item(15)
        para.Range.Style = myDoc.Styles.Item("Box List Number 3")
        para = rng.Paragraphs.Item(16)
        para.Range.Style = myDoc.Styles.Item("Box List Number 3")
        para = rng.Paragraphs.Item(17)
        para.Range.Style = myDoc.Styles.Item("Box List Number 2")
        para = rng.Paragraphs.Item(18)
        para.Range.Style = myDoc.Styles.Item("Box List Number")
        '
        para = rng.Paragraphs.Item(19)
        para.Range.Style = myDoc.Styles.Item("Box Text")
        '        
        '        
        para = rng.Paragraphs.Item(20)
        para.Range.Style = myDoc.Styles.Item("Box Quote")
        para = rng.Paragraphs.Item(21)
        para.Range.Style = myDoc.Styles.Item("Box Quote List Bullet")
        para = rng.Paragraphs.Item(22)
        para.Range.Style = myDoc.Styles.Item("Box Quote List Bullet")
        para = rng.Paragraphs.Item(23)
        para.Range.Style = myDoc.Styles.Item("Box Quote List Bullet")
        '        
        para = rng.Paragraphs.Item(24)
        para.Range.Style = myDoc.Styles.Item("Box Quote Source")
        '        
        rng.Collapse(WdCollapseDirection.wdCollapseStart)
        '
    End Sub
    '

    '
    Public Function insert_ExampleTableText_column01(ByRef rng As Word.Range, strType As String) As Word.Range
        Dim para As Word.Paragraph
        Dim myDoc As Word.Document
        Dim i As Integer
        '
        myDoc = Me.glb_get_wrdActiveDoc()
        '
        rng.Text = ""
        rng.Text = Me.get_TblExampleMessage_column01()
        '
        Select Case strType
            Case "normal"
                For i = 1 To rng.Paragraphs.Count
                    para = rng.Paragraphs.Item(i)
                    Select Case i
                        Case 1
                            para.Style = myDoc.Styles("Table side heading 1")
                        Case 2
                            para.Style = myDoc.Styles("Table text")
                        Case 3
                            para.Style = myDoc.Styles("Table side heading 2")
                        Case 4
                            para.Style = myDoc.Styles("Table text")
                        Case 5
                            para.Style = myDoc.Styles("Table Quote")
                        Case 6, 7, 8, 9
                            para.Style = myDoc.Styles("Table Quote Bullet")
                        Case 10
                            para.Style = myDoc.Styles("Table Quote Source")

                    End Select
                Next

            Case "small"
                For i = 1 To rng.Paragraphs.Count
                    para = rng.Paragraphs.Item(i)
                    Select Case i
                        Case 1
                            para.Style = myDoc.Styles("Table side heading 1 (small)")
                        Case 2
                            para.Style = myDoc.Styles("Table text (small)")
                        Case 3
                            para.Style = myDoc.Styles("Table side heading 2 (small)")
                        Case 4
                            para.Style = myDoc.Styles("Table text (small)")
                        Case 5
                            para.Style = myDoc.Styles("Table Quote (small)")
                        Case 6, 7, 8, 9
                            para.Style = myDoc.Styles("Table Quote Bullet (small)")
                        Case 10
                            para.Style = myDoc.Styles("Table Quote Source (small)")

                    End Select
                Next
        End Select
        '
        Return rng
    End Function
    '
    Public Function insert_ExampleTableText_column02(ByRef rng As Word.Range, strType As String) As Word.Range
        Dim para As Word.Paragraph
        Dim myDoc As Word.Document
        Dim i As Integer
        '
        myDoc = Me.glb_get_wrdActiveDoc()
        '
        rng.Text = ""
        rng.Text = Me.get_TblExampleMessage_column02()
        '
        Select Case strType
            Case "normal"
                For i = 1 To rng.Paragraphs.Count
                    para = rng.Paragraphs.Item(i)
                    Select Case i
                        Case 1
                            para.Style = myDoc.Styles("Table text")
                        Case 2, 3
                            para.Style = myDoc.Styles("Table list bullet")
                        Case 4
                            para.Style = myDoc.Styles("Table list bullet 2")
                        Case 5, 6
                            para.Style = myDoc.Styles("Table list bullet 3")
                        Case 7
                            para.Style = myDoc.Styles("Table list bullet 2")
                        Case 8, 9
                            para.Style = myDoc.Styles("Table list bullet")
                        Case 10
                            para.Style = myDoc.Styles("Table text")
                    End Select
                Next
            Case "small"
                For i = 1 To rng.Paragraphs.Count
                    para = rng.Paragraphs.Item(i)
                    Select Case i
                        Case 1
                            para.Style = myDoc.Styles("Table text (small)")
                        Case 2, 3
                            para.Style = myDoc.Styles("Table list bullet (small)")
                        Case 4
                            para.Style = myDoc.Styles("Table list bullet 2 (small)")
                        Case 5, 6
                            para.Style = myDoc.Styles("Table list bullet 3 (small)")
                        Case 7
                            para.Style = myDoc.Styles("Table list bullet 2 (small)")
                        Case 8, 9
                            para.Style = myDoc.Styles("Table list bullet (small)")
                        Case 10
                            para.Style = myDoc.Styles("Table text (small)")
                    End Select
                Next

        End Select
        '
        Return rng
    End Function
    '
    '
    Public Function insert_ExampleTableText_column03(ByRef rng As Word.Range, strType As String) As Word.Range
        Dim para As Word.Paragraph
        Dim myDoc As Word.Document
        Dim i As Integer
        '

        myDoc = Me.glb_get_wrdActiveDoc()
        '
        rng.Text = ""
        rng.Text = Me.get_TblExampleMessage_column03()
        '
        Select Case strType
            Case "normal"
                For i = 1 To rng.Paragraphs.Count
                    para = rng.Paragraphs.Item(i)
                    Select Case i
                        Case 1
                            para.Style = myDoc.Styles("Table text")
                        Case 2, 3
                            para.Style = myDoc.Styles("Table list number")
                        Case 4
                            para.Style = myDoc.Styles("Table list number 2")
                        Case 5, 6
                            para.Style = myDoc.Styles("Table list number 3")
                        Case 7
                            para.Style = myDoc.Styles("Table list number 2")
                        Case 8, 9
                            para.Style = myDoc.Styles("Table list number")
                        Case 10
                            para.Style = myDoc.Styles("Table text")
                    End Select
                Next
            Case "small"
                For i = 1 To rng.Paragraphs.Count
                    para = rng.Paragraphs.Item(i)
                    Select Case i
                        Case 1
                            para.Style = myDoc.Styles("Table text (small)")
                        Case 2, 3
                            para.Style = myDoc.Styles("Table list number (small)")
                        Case 4
                            para.Style = myDoc.Styles("Table list number 2 (small)")
                        Case 5, 6
                            para.Style = myDoc.Styles("Table list number 3 (small)")
                        Case 7
                            para.Style = myDoc.Styles("Table list number 2 (small)")
                        Case 8, 9
                            para.Style = myDoc.Styles("Table list number (small)")
                        Case 10
                            para.Style = myDoc.Styles("Table text (small)")
                    End Select
                Next

        End Select
        '
        Return rng
    End Function
    '
    Public Function get_TblExampleMessage_column01() As String
        Dim strMsg As String
        '
        strMsg = ""
        strMsg = strMsg + "Table side heading 1" + vbCrLf
        strMsg = strMsg + "Table text style ... Or aut molecto eatibust harupti assinis sint, est et ommolum in consequi" + vbCrLf
        strMsg = strMsg + "Table side heading 2" + vbCrLf
        strMsg = strMsg + "aciet im sitatinctem faccuptatas enist, comnient, odiorem quidi te non cuptaqu osapidebis veleculpa a doluptati dis suntisimi, cus aut et ut labo. Itatquam et etur apistiuntiis di dolor maximustiat." + vbCrLf
        '
        strMsg = strMsg + "Table Quote" + vbCrLf

        strMsg = strMsg + "Table quote bullet" + vbCrLf
        strMsg = strMsg + "Table quote bullet" + vbCrLf
        strMsg = strMsg + "Table quote bullet" + vbCrLf
        strMsg = strMsg + "Table quote bullet" + vbCrLf
        '
        strMsg = strMsg + "Table Quote Source style"
        '
        Return strMsg

    End Function
    '
    Public Function get_TblExampleMessage_column02() As String
        Dim strMsg As String
        '
        strMsg = ""
        strMsg = strMsg + "Table text style….Or aut molecto eatibust harupti assinis sint, est et ommolum in consequi" + vbCrLf
        strMsg = strMsg + "Table list bullet" + vbCrLf
        strMsg = strMsg + "Table list bullet" + vbCrLf
        strMsg = strMsg + "Table list bullet" + vbCrLf
        strMsg = strMsg + "Table list bullet" + vbCrLf
        strMsg = strMsg + "Table list bullet" + vbCrLf
        strMsg = strMsg + "Table list bullet" + vbCrLf
        strMsg = strMsg + "Table list bullet" + vbCrLf
        strMsg = strMsg + "Table list bullet" + vbCrLf

        strMsg = strMsg + "Table text style"
        '
        Return strMsg

    End Function
    '
    Public Function get_TblExampleMessage_column03() As String
        Dim strMsg As String
        '
        strMsg = ""
        strMsg = strMsg + "Table text style….Or aut molecto eatibust harupti assinis sint, est et ommolum in consequi" + vbCrLf
        strMsg = strMsg + "Table list number" + vbCrLf
        strMsg = strMsg + "Table list number" + vbCrLf
        strMsg = strMsg + "Table list number" + vbCrLf
        strMsg = strMsg + "Table list number" + vbCrLf
        strMsg = strMsg + "Table list number" + vbCrLf
        strMsg = strMsg + "Table list number" + vbCrLf
        strMsg = strMsg + "Table list number" + vbCrLf
        strMsg = strMsg + "Table list number" + vbCrLf

        strMsg = strMsg + "Table text style"
        '
        Return strMsg

    End Function


#End Region
    '
    ''' <summary>
    ''' This method will accept a document (myDoc) and a Font name (strFontName) and
    ''' will convert all styles withing myDoc to use the specified font. If there is an error
    ''' it will return an error message string. Otherwise the return string will be null
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <param name="strFontName"></param>
    ''' <returns></returns>
    Public Function styl_convert_toNewFont(ByRef myDoc As Word.Document, strFontName As String) As String
        Dim styl As Word.Style
        Dim strMsg As String
        '
        strMsg = ""
        '
        Try
            'Change Normal style to the new style.. Then go through all styles
            'and change them if they are not already changed
            styl = myDoc.Styles.Item("Normal")
            styl.Font.Name = strFontName
            'styl.Font.Size = 9.5
            'styl.Font.Size = 10
            '
            For Each styl In myDoc.Styles
                If styl.Font.Name <> strFontName Then
                    styl.Font.Name = strFontName
                End If
            Next
            '
        Catch ex As Exception
            strMsg = "Error in conversion to '" + strFontName + "'"
        End Try
        '
        Return strMsg
        '
    End Function
    '
#Region "Build Styles in current Doc"
    '
    Public Sub styl_buildStyle_BodyText(strStyleName As String)
        Dim myDoc As Word.Document
        Dim myStyle As Word.Style
        Dim spaceAfter As Single
        '
        myDoc = Me.glb_get_wrdActiveDoc()
        '
        Try
            'See if the style exists
            myStyle = myDoc.Styles.Item(strStyleName)
            spaceAfter = myStyle.ParagraphFormat.SpaceAfter
        Catch ex As Exception
            myStyle = myDoc.Styles.Add(strStyleName)

        End Try
    End Sub
#End Region
    '
End Class

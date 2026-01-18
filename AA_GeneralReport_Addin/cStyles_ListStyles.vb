Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class cStyles_ListStyles
    Inherits cStylesManager
    Public Sub New()
        MyBase.New()
    End Sub
    '
    Public Sub lstStyle_link_HeadingStyles_ES()
        Dim myStyle As Word.Style
        Dim lstTmpl As ListTemplate
        Dim lstLevel As ListLevel
        Dim myDoc As Word.Document
        '
        myDoc = glb_get_wrdActiveDoc()
        '
        Try
            myStyle = myDoc.Styles.Item("aac_lstStyle_HeadingsES")
            lstTmpl = myStyle.ListTemplate
        Catch ex As Exception
            myStyle = myDoc.Styles.Add("aac_lstStyle_HeadingsES", WdStyleType.wdStyleTypeList)
            myStyle.BaseStyle = myDoc.Styles.Item("Normal")
            lstTmpl = myStyle.ListTemplate
        End Try
        '
        lstLevel = lstTmpl.ListLevels.Item(1)
        lstLevel.LinkedStyle() = "Heading 1 (ES)"
        '
        lstLevel = lstTmpl.ListLevels.Item(2)
        lstLevel.LinkedStyle() = "Heading 2 (ES)"
        '
        lstLevel = lstTmpl.ListLevels.Item(3)
        lstLevel.LinkedStyle() = "Heading 3 (ES)"
        '
        lstLevel = lstTmpl.ListLevels.Item(4)
        lstLevel.LinkedStyle() = "Heading 4 (ES)"

    End Sub
    '
    Public Sub lstStyle_link_HeadingStyles_Std()
        Dim myStyle As Word.Style
        Dim lstTmpl As ListTemplate
        Dim lstLevel As ListLevel
        Dim myDoc As Word.Document
        '
        myDoc = glb_get_wrdActiveDoc()
        '
        Try
            myStyle = myDoc.Styles.Item("lstStyle_Heading_Numbered")
            lstTmpl = myStyle.ListTemplate
        Catch ex As Exception
            myStyle = myDoc.Styles.Add("lstStyle_Heading_Numbered", WdStyleType.wdStyleTypeList)
            'myStyle.BaseStyle = myDoc.Styles.Item("Normal")
            lstTmpl = myStyle.ListTemplate
        End Try
        '
        lstLevel = lstTmpl.ListLevels.Item(1)
        lstLevel.LinkedStyle() = "Heading 1"
        lstLevel = lstTmpl.ListLevels.Item(1)
        lstLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
        lstLevel.NumberFormat = "%1"
        'lstLevel.NumberPosition = tableLeftIndent
        'lstLevel.NumberPosition = -18
        lstLevel.NumberPosition = -35
        'lstLevel.NumberPosition
        'lstLevel.TextPosition = 16
        lstLevel.TextPosition = 0
        lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft

        '
        lstLevel = lstTmpl.ListLevels.Item(2)
        lstLevel.LinkedStyle() = "Heading 2"
        lstLevel = lstTmpl.ListLevels.Item(2)
        lstLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
        'lstLevel.NumberFormat = "%1.%2.%3"
        lstLevel.NumberFormat = "%1.%2"

        lstLevel.NumberPosition = -6.8
        lstLevel.TextPosition = 0
        lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignRight
        '
        lstLevel = lstTmpl.ListLevels.Item(3)
        lstLevel.LinkedStyle() = "Heading 3"
        lstLevel = lstTmpl.ListLevels.Item(3)
        'lstLevel.LinkedStyle() = strHeading3
        lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
        lstLevel.NumberPosition = 0
        lstLevel.TextPosition = 0
        'lstLevel.TextPosition = 60
        'lstLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
        'lstLevel.NumberFormat = "%1.%2.%3"
        'lstLevel.NumberFormat = ""
        '
        lstLevel = lstTmpl.ListLevels.Item(4)
        lstLevel.LinkedStyle() = "Heading 4"
        '
        lstLevel = lstTmpl.ListLevels.Item(5)
        lstLevel.LinkedStyle() = "Heading 5"
        '
        '
        '

        '
        lstLevel = lstTmpl.ListLevels.Item(4)
        'lstLevel.LinkedStyle() = strHeading4
        lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
        lstLevel.NumberPosition = 0
        lstLevel.TextPosition = 0


    End Sub


    Public Sub lstStyle_build_Heading(strListStyleType As String)
        Dim myStyle As Word.Style
        Dim lstTmpl As ListTemplate
        Dim lstLevel As ListLevel
        Dim myDoc As Word.Document
        Dim strHeading1, strHeading2, strHeading3, strHeading4 As String
        Dim strLinkedStyle As String
        Dim j As Integer
        '
        strLinkedStyle = ""
        strHeading1 = ""
        strHeading2 = ""
        strHeading3 = ""
        strHeading4 = ""
        '
        myDoc = glb_get_wrdActiveDoc()
        myStyle = Nothing
        '
        Select Case strListStyleType
            Case "es"
                Try
                    myStyle = myDoc.Styles.Item("aac_lstStyle_HeadingsES")
                    j = 1
                    'myStyle.Delete()
                    'myStyle = myDoc.Styles.Add("aac_lstStyle_HeadingsES", WdStyleType.wdStyleTypeList)
                    'myStyle.BaseStyle = myDoc.Styles.Item("Normal")

                Catch ex As Exception
                    'myStyle = myDoc.Styles.Add("aac_lstStyle_HeadingsES", WdStyleType.wdStyleTypeList)
                    'myStyle.BaseStyle = myDoc.Styles.Item("Normal")
                    '
                End Try
                strHeading1 = "Heading 1 (ES)"
                strHeading2 = "Heading 2 (ES)"
                strHeading3 = "Heading 3 (ES)"
                strHeading4 = "Heading 4 (ES)"
                '
                lstTmpl = myStyle.ListTemplate
                lstLevel = lstTmpl.ListLevels.Item(1)
                'lstLevel.NumberFormat = ""
                lstLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleNone
                lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                lstLevel.NumberPosition = 0
                lstLevel.TextPosition = 0
                'lstLevel.LinkedStyle = ""
                lstLevel.LinkedStyle = "Heading 1 (ES)"
                'lstLevel.LinkedStyle = "Heading 1 (ES)"
                'lstLevel.LinkedStyle = "Heading 1 (ES)"

                'lstLevel.NumberPosition = 0
                'lstLevel.TextPosition = 0
                'strLinkedStyle = lstLevel.LinkedStyle()
                'j = 1
                '
            Case "rpt"
                Try
                    myStyle = myDoc.Styles.Item("aac_lstStyle_HeadingNumbered")
                Catch ex As Exception
                    myStyle = myDoc.Styles.Add("aac_lstStyle_HeadingNumbered", WdStyleType.wdStyleTypeList)
                    myStyle.BaseStyle = myDoc.Styles.Item("Normal")
                    '
                End Try
                strHeading1 = "Heading 1"
                strHeading2 = "Heading 2"
                strHeading3 = "Heading 3"
                strHeading4 = "Heading 4"
                '
            Case "ap"
                Try
                    myStyle = myDoc.Styles.Item("aac_lstStyle_HeadingApp")
                Catch ex As Exception
                    myStyle = myDoc.Styles.Add("aac_lstStyle_HeadingApp", WdStyleType.wdStyleTypeList)
                    myStyle.BaseStyle = myDoc.Styles.Item("Normal")
                    '
                End Try
                strHeading1 = "Heading 1 (AP)"
                strHeading2 = "Heading 2 (AP)"
                strHeading3 = "Heading 3 (AP)"
                strHeading4 = "Heading 4 (AP)"
                '
                'strHeading1 = "Heading 6"
                'strHeading2 = "Heading 7"
                'strHeading3 = "Heading 8"
                'strHeading4 = "Heading 9"
                '
        End Select
        '        
        If Not IsNothing(myStyle) Then
            lstTmpl = myStyle.ListTemplate
            '
            Select Case strListStyleType
                Case "rpt", "ap"
                    lstLevel = lstTmpl.ListLevels.Item(1)
                    'lstLevel.LinkedStyle() = strHeading1
                    'lstLevel.NumberFormat = "x"
                    'lstLevel.NumberPosition = tableLeftIndent
                    'lstLevel.NumberPosition = -18
                    lstLevel.NumberPosition = -35
                    'lstLevel.TextPosition = 16
                    lstLevel.TextPosition = -2
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignRight
                    '
                    lstLevel = lstTmpl.ListLevels.Item(2)
                    'lstLevel.LinkedStyle() = strHeading2
                    lstLevel.NumberPosition = -6.8
                    lstLevel.TextPosition = 0
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignRight
                    '
                    lstLevel = lstTmpl.ListLevels.Item(3)
                    'lstLevel.LinkedStyle() = strHeading3
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    lstLevel.NumberPosition = 0
                    lstLevel.TextPosition = 0
                    'lstLevel.TextPosition = 60

                    lstLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
                    'lstLevel.NumberFormat = "%1.%2.%3"
                    lstLevel.NumberFormat = ""
                    '
                    lstLevel = lstTmpl.ListLevels.Item(4)
                    'lstLevel.LinkedStyle() = strHeading4
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    lstLevel.NumberPosition = 0
                    lstLevel.TextPosition = 0

                Case "es"

                Case "nonum"
                    lstLevel = lstTmpl.ListLevels.Item(1)
                    'lstLevel.LinkedStyle() = ""
                    'lstLevel.LinkedStyle() = strHeading1
                    'lstLevel.NumberFormat = "x"
                    'lstLevel.NumberPosition = tableLeftIndent
                    'lstLevel.NumberPosition = -18
                    lstLevel.NumberPosition = 0
                    'lstLevel.TextPosition = 16
                    lstLevel.TextPosition = 0
                    'lstLevel.NumberFormat = ""
                    'lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    'lstLevel.LinkedStyle() = "Heading 1 (ES)"

                    '
                    lstLevel = lstTmpl.ListLevels.Item(2)
                    'lstLevel.LinkedStyle() = strHeading2
                    lstLevel.NumberPosition = 0
                    lstLevel.TextPosition = 0
                    'lstLevel.NumberFormat = ""
                    'lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    'lstLevel.LinkedStyle() = strHeading2
                    'lstLevel.LinkedStyle() = "Heading 2 (ES)"
                    '
                    lstLevel = lstTmpl.ListLevels.Item(3)
                    'lstLevel.LinkedStyle() = strHeading3
                    'lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    lstLevel.NumberPosition = 0
                    lstLevel.TextPosition = 0
                    'lstLevel.LinkedStyle() = strHeading3
                    'lstLevel.LinkedStyle() = "Heading 3 (ES)"


                    'lstLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
                    'lstLevel.NumberFormat = "%1.%2.%3"
                    'lstLevel.NumberFormat = ""
                    '
                    lstLevel = lstTmpl.ListLevels.Item(4)
                    'lstLevel.LinkedStyle() = strHeading4
                    'lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    lstLevel.NumberPosition = 0
                    lstLevel.TextPosition = 0
                    'lstLevel.NumberFormat = ""
                    'lstLevel.LinkedStyle() = "Heading 4 (ES)"

            End Select
            '
        End If

    End Sub
    '
    Public Function lstStyle_modify_Appendices() As Word.Style
        Dim myStyle As Word.Style
        Dim lstTmpl As ListTemplate
        Dim lstLevel As ListLevel
        Dim myDoc As Word.Document
        '
        myDoc = glb_get_wrdActiveDoc()
        myStyle = Nothing
        '
        Try
            myStyle = myDoc.Styles.Item("aac_lstStyle_Appendices")
            lstTmpl = myStyle.ListTemplate
            lstLevel = lstTmpl.ListLevels.Item(2)
            lstLevel.Font.Name = "Yu Gothic Medium"
            'lstLevel.Font.Size = 17

        Catch ex As Exception
            myStyle = Nothing
        End Try
        '
        Return myStyle
        '
    End Function
    '
    Public Sub lstStyle_Heading_Heading1(table_DeltaLeftIndent As Single)
        Dim myStyle As Word.Style
        Dim lstTmpl As ListTemplate
        Dim lstLevel As ListLevel
        Dim myDoc As Word.Document
        '
        myDoc = glb_get_wrdActiveDoc()
        '
        Try
            myStyle = myDoc.Styles.Item("aac_lstStyle_HeadingNumbered")
        Catch ex As Exception
            myStyle = myDoc.Styles.Add("aac_lstStyle_HeadingNumbered", WdStyleType.wdStyleTypeList)
            myStyle.BaseStyle = myDoc.Styles.Item("Normal")
            '
            lstTmpl = myStyle.ListTemplate
            lstLevel = lstTmpl.ListLevels.Item(1)
            'lstLevel.NumberFormat = "x"
            'lstLevel.NumberPosition = tableLeftIndent
            lstLevel.NumberPosition = -6.8
            'lstLevel.TextPosition = 16
            lstLevel.TextPosition = 12.4
            lstLevel.LinkedStyle() = "Table list bullet"
            '
            lstLevel = lstTmpl.ListLevels.Item(2)
            'lstLevel.NumberPosition = 16
            'lstLevel.TextPosition = 28
            lstLevel.NumberPosition = 14 + table_DeltaLeftIndent
            lstLevel.TextPosition = 26 + table_DeltaLeftIndent
            lstLevel.LinkedStyle() = "Table list bullet 2"

        End Try

        '
        'Table List Bullets
        lstTmpl = myStyle.ListTemplate
        '
        lstLevel = lstTmpl.ListLevels.Item(1)
        'lstLevel.NumberFormat = "x"
        'lstLevel.NumberPosition = tableLeftIndent
        lstLevel.NumberPosition = -6.8
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
    End Sub

End Class

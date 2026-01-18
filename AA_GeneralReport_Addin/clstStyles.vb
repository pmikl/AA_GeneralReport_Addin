Imports Microsoft.Office.Interop.Word
Public Class clstStyles
    Inherits cGlobals
    '
    Public lstStyle_BD_Name As String = "lstStyle_Heading_Numbered"
    Public lstStyle_AP_Name As String = "aac_lstStyle_Appendices"
    '
    Public Sub New()
        MyBase.New()
    End Sub
    '
    '
    ''' <summary>
    ''' This method will set the lstStyle harnessing for the Body Heading styles, Heading 1,
    ''' Heading 2, Heading 3, Heading 4, Heading 5.
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function lstStyle_set_HeadingsBD_Numbered(ByRef myDoc As Word.Document) As Word.ListTemplate
        Dim lstStyle As Word.Style
        Dim lstTemplate As Word.ListTemplate
        Dim lstLevel As Word.ListLevel
        Dim j As Integer
        '
        lstTemplate = Nothing
        lstStyle = myDoc.Styles.Item(Me.lstStyle_BD_Name)
        lstTemplate = lstStyle.ListTemplate
        '
        For j = 1 To 5
            lstLevel = lstTemplate.ListLevels.Item(j)
            Select Case j
                Case 1
                    lstLevel.LinkedStyle = "Heading 1"
                    lstLevel.NumberFormat = "%1"
                    lstLevel.NumberPosition = -18
                    lstLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignRight
                    'lstLevel.TabPosition =
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingTab
                Case 2
                    lstLevel.LinkedStyle = "Heading 2"
                    lstLevel.NumberFormat = "%1.%2"
                    lstLevel.NumberPosition = -18
                    lstLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignRight
                    'lstLevel.TabPosition = -18
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingTab
                Case 3
                    lstLevel.LinkedStyle = "Heading 3"
                    lstLevel.NumberFormat = "%1.%2.%3"
                    lstLevel.NumberPosition = 0
                    lstLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    lstLevel.TabPosition = 51
                    lstLevel.TextPosition = 50
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingTab
                Case 4
                    lstLevel.LinkedStyle = "Heading 4"
                    lstLevel.NumberFormat = ""
                    lstLevel.NumberPosition = 0
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    lstLevel.TabPosition = 0
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingNone
                Case 4
                    lstLevel.LinkedStyle = "Heading 5"
                    lstLevel.NumberFormat = ""
                    lstLevel.NumberPosition = 0
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    lstLevel.TabPosition = 0
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingNone

            End Select

        Next
        '
        'Relink the headings to make them stick
        Me.lstStyle_relink_Headings_BD(myDoc)
        '
        Return lstTemplate
        '
    End Function
    '
    '
    Public Function lstStyle_set_HeadingsBD_noNumbered(ByRef myDoc As Word.Document) As Word.ListTemplate
        Dim lstStyle As Word.Style
        Dim lstTemplate As Word.ListTemplate
        Dim lstLevel As Word.ListLevel
        Dim j As Integer
        '
        lstTemplate = Nothing
        lstStyle = myDoc.Styles.Item(Me.lstStyle_BD_Name)
        lstTemplate = lstStyle.ListTemplate
        '
        For j = 1 To 5
            lstLevel = lstTemplate.ListLevels.Item(j)
            Select Case j
                Case 1
                    lstLevel.LinkedStyle = "Heading 1"
                    lstLevel.NumberFormat = "%1"
                    lstLevel.NumberPosition = -18
                    lstLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignRight
                    'lstLevel.TabPosition =
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingTab
                Case 2
                    lstLevel.LinkedStyle = "Heading 2"
                    lstLevel.NumberFormat = "%1.%2"
                    lstLevel.NumberPosition = -18
                    lstLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignRight
                    'lstLevel.TabPosition = -18
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingTab
                Case 3
                    lstLevel.LinkedStyle = "Heading 3"
                    lstLevel.NumberFormat = ""
                    lstLevel.NumberPosition = 0
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    lstLevel.TabPosition = 0
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingNone
                Case 4
                    lstLevel.LinkedStyle = "Heading 4"
                    lstLevel.NumberFormat = ""
                    lstLevel.NumberPosition = 0
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    lstLevel.TabPosition = 0
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingNone
                Case 4
                    lstLevel.LinkedStyle = "Heading 5"
                    lstLevel.NumberFormat = ""
                    lstLevel.NumberPosition = 0
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    lstLevel.TabPosition = 0
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingNone

            End Select

        Next
        '
        'Relink the headings to make them stick
        Me.lstStyle_relink_Headings_BD(myDoc)
        '
        Return lstTemplate
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will set the lstStyle harnessing for the Appendix Heading styles, Heading 6,
    ''' Heading 7, Heading 8, Heading 9, Heading 5 (AP).
    ''' </summary>
    ''' <param name="myDoc"></param>
    ''' <returns></returns>
    Public Function lstStyle_set_HeadingsAP_Numbered(ByRef myDoc As Word.Document) As Word.ListTemplate
        Dim lstStyle As Word.Style
        Dim lstTemplate As Word.ListTemplate
        Dim lstLevel As Word.ListLevel
        Dim j As Integer
        '
        lstTemplate = Nothing
        lstStyle = myDoc.Styles.Item(Me.lstStyle_AP_Name)
        lstTemplate = lstStyle.ListTemplate
        '
        For j = 1 To 5
            lstLevel = lstTemplate.ListLevels.Item(j)
            Select Case j
                Case 1
                    lstLevel.LinkedStyle = "Heading 6"
                    lstLevel.NumberFormat = "%1"
                    lstLevel.NumberPosition = -18
                    lstLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleUppercaseLetter
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignRight
                    'lstLevel.TabPosition =
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingTab
                Case 2
                    lstLevel.LinkedStyle = "Heading 7"
                    lstLevel.NumberFormat = "%1.%2"
                    lstLevel.NumberPosition = -18
                    lstLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignRight
                    'lstLevel.TabPosition = -18
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingTab
                Case 3
                    lstLevel.LinkedStyle = "Heading 8"
                    lstLevel.NumberFormat = "%1.%2.%3"
                    lstLevel.NumberPosition = 0
                    lstLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    lstLevel.TabPosition = 51
                    lstLevel.TextPosition = 50
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingTab
                Case 4
                    lstLevel.LinkedStyle = "Heading 9"
                    lstLevel.NumberFormat = ""
                    lstLevel.NumberPosition = 0
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    lstLevel.TabPosition = 0
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingNone
                Case 4
                    lstLevel.LinkedStyle = "Heading 5 (AP)"
                    lstLevel.NumberFormat = ""
                    lstLevel.NumberPosition = 0
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    lstLevel.TabPosition = 0
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingNone

            End Select

        Next
        '
        'Relink the headings to make them stick
        Me.lstStyle_relink_Headings_AP(myDoc)
        '
        Return lstTemplate
        '
    End Function
    '
    '
    Public Function lstStyle_set_HeadingsAP_noNumbered(ByRef myDoc As Word.Document) As Word.ListTemplate
        Dim lstStyle As Word.Style
        Dim lstTemplate As Word.ListTemplate
        Dim lstLevel As Word.ListLevel
        Dim j As Integer
        '
        lstTemplate = Nothing
        lstStyle = myDoc.Styles.Item(Me.lstStyle_AP_Name)
        lstTemplate = lstStyle.ListTemplate
        '
        For j = 1 To 5
            lstLevel = lstTemplate.ListLevels.Item(j)
            Select Case j
                Case 1
                    lstLevel.LinkedStyle = "Heading 6"
                    lstLevel.NumberFormat = "%1"
                    lstLevel.NumberPosition = -18
                    lstLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleUppercaseLetter
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignRight
                    'lstLevel.TabPosition =
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingTab
                Case 2
                    lstLevel.LinkedStyle = "Heading 7"
                    lstLevel.NumberFormat = "%1.%2"
                    lstLevel.NumberPosition = -18
                    lstLevel.NumberStyle = WdListNumberStyle.wdListNumberStyleArabic
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignRight
                    'lstLevel.TabPosition = -18
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingTab
                Case 3
                    lstLevel.LinkedStyle = "Heading 8"
                    lstLevel.NumberFormat = ""
                    lstLevel.NumberPosition = 0
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    lstLevel.TabPosition = 0
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingNone
                Case 4
                    lstLevel.LinkedStyle = "Heading 9"
                    lstLevel.NumberFormat = ""
                    lstLevel.NumberPosition = 0
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    lstLevel.TabPosition = 0
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingNone
                Case 4
                    lstLevel.LinkedStyle = "Heading 5 (AP)"
                    lstLevel.NumberFormat = ""
                    lstLevel.NumberPosition = 0
                    lstLevel.Alignment = WdListLevelAlignment.wdListLevelAlignLeft
                    lstLevel.TabPosition = 0
                    lstLevel.TextPosition = 0
                    lstLevel.TrailingCharacter = WdTrailingCharacter.wdTrailingNone

            End Select

        Next
        '
        'Relink the headings to make them stick
        Me.lstStyle_relink_Headings_AP(myDoc)
        '
        Return lstTemplate
        '
    End Function
    '
    '
    '
    Public Sub lstStyle_relink_Headings_noNum(ByRef myDoc As Word.Document)
        Dim lstStyle As Word.Style
        Dim lstTemplate As Word.ListTemplate
        Dim lstLevel As Word.ListLevel
        Dim j As Integer
        '
        Try
            lstStyle = myDoc.Styles.Item("aa Heading (no number) List")
            'lstStyle = objGlobals.glb_get_wrdActiveDoc.Styles.Item("aa Heading (no number) List")
            lstTemplate = lstStyle.ListTemplate
            For j = 1 To 5
                lstLevel = lstTemplate.ListLevels.Item(j)
                Select Case j
                    Case 1
                        lstLevel.LinkedStyle = "Heading 1 (no number)"
                    Case 2
                        lstLevel.LinkedStyle = "Heading 2 (no number)"
                    Case 3
                        lstLevel.LinkedStyle = "Heading 3 (no number)"
                    Case 4
                        lstLevel.LinkedStyle = "Heading 4 (no number)"
                    Case 5
                        lstLevel.LinkedStyle = "Heading 5 (no number)"
                End Select
            Next
        Catch ex As Exception

        End Try

    End Sub
    '

    '
    Public Sub lstStyle_relink_Headings_ES(ByRef myDoc As Word.Document)
        Dim lstStyle As Word.Style
        Dim lstTemplate As Word.ListTemplate
        Dim lstLevel As Word.ListLevel
        Dim j As Integer
        '
        Try
            lstStyle = myDoc.Styles.Item("aac_lstStyle_HeadingsES")
            'lstStyle = objGlobals.glb_get_wrdActiveDoc.Styles.Item("aa Heading (no number) List")
            lstTemplate = lstStyle.ListTemplate
            For j = 1 To 5
                lstLevel = lstTemplate.ListLevels.Item(j)
                Select Case j
                    Case 1
                        lstLevel.LinkedStyle = "Heading 1 (ES)"
                    Case 2
                        lstLevel.LinkedStyle = "Heading 2 (ES)"
                    Case 3
                        lstLevel.LinkedStyle = "Heading 3 (ES)"
                    Case 4
                        lstLevel.LinkedStyle = "Heading 4 (ES)"
                    Case 5
                        lstLevel.LinkedStyle = "Heading 5 (ES)"
                End Select
            Next
        Catch ex As Exception

        End Try

    End Sub
    '

    '
    Public Sub lstStyle_relink_Headings_BD(ByRef myDoc As Word.Document)
        Dim lstStyle As Word.Style
        Dim lstTemplate As Word.ListTemplate
        Dim lstLevel As Word.ListLevel
        Dim j As Integer
        '
        Try
            lstStyle = myDoc.Styles.Item(Me.lstStyle_BD_Name)
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

        Catch ex As Exception

        End Try

    End Sub
    '
    '
    Public Sub lstStyle_relink_Headings_AP(ByRef myDoc As Word.Document)
        Dim lstStyle As Word.Style
        Dim lstTemplate As Word.ListTemplate
        Dim lstLevel As Word.ListLevel
        Dim j As Integer
        '
        Try
            lstStyle = myDoc.Styles.Item(Me.lstStyle_AP_Name)
            'lstStyle = objGlobals.glb_get_wrdActiveDoc.Styles.Item("aa Heading (no number) List")
            lstTemplate = lstStyle.ListTemplate
            For j = 1 To 5
                lstLevel = lstTemplate.ListLevels.Item(j)
                Select Case j
                    Case 1
                        lstLevel.LinkedStyle = "Heading 6"
                    Case 2
                        lstLevel.LinkedStyle = "Heading 7"
                    Case 3
                        lstLevel.LinkedStyle = "Heading 8"
                    Case 4
                        lstLevel.LinkedStyle = "Heading 9"
                    Case 5
                        lstLevel.LinkedStyle = "Heading 5 (AP)"
                End Select
            Next

        Catch ex As Exception

        End Try

    End Sub
    '

End Class

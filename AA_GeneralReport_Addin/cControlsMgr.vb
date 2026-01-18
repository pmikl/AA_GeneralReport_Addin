Imports Microsoft.Office.Interop.Word
Public Class cControlsMgr
    Public _strTabId_Styles = "tab_aa_Styles"
    Public _strTabId_Placeholders = "tab_aa_Placeholders"
    Public _strTabId_PagesAndSections = "tab_aa_PagesAndSections"
    Public _strTabId_Finalise = "tab_aa_Finalise"
    Public _strTabId_AAHome = "tab_aa_Home"

    Public Sub New()

    End Sub
    '
    ''' <summary>
    ''' This method will toggle the visibility of the tab as specified by strTabID. 
    ''' Typically you would select strTabId from the class variables '_strTabId_Styles, _strTabId_Placeholders
    ''' _strTabId_PagesAndSections, _strTabId_Finalise
    ''' </summary>
    Public Sub ctrl_tabToggle_Visibility(strTabId As String)
        Dim tabVisibility As Boolean
        '
        Select Case strTabId
            Case _strTabId_Styles
                tabVisibility = Globals.Ribbons.rbn_Styles.tab_aa_PagesAndSections.Visible
                Globals.Ribbons.rbn_Styles.tab_aa_Styles.Visible = Not tabVisibility
                '
            Case _strTabId_Placeholders
                tabVisibility = Globals.Ribbons.rbn_Styles.tab_aa_PagesAndSections.Visible
                Globals.Ribbons.rbn_Styles.tab_aa_Placeholders.Visible = Not tabVisibility
                '
            Case _strTabId_PagesAndSections
                tabVisibility = Globals.Ribbons.rbn_Styles.tab_aa_PagesAndSections.Visible
                Globals.Ribbons.rbn_Styles.tab_aa_PagesAndSections.Visible = Not tabVisibility
                '
            Case _strTabId_Finalise
                tabVisibility = Globals.Ribbons.rbn_Styles.tab_aa_PagesAndSections.Visible
                Globals.Ribbons.rbn_Styles.tab_aa_Finalise.Visible = Not tabVisibility
                '
            Case _strTabId_AAHome
                tabVisibility = Globals.Ribbons.rbn_Styles.tab_aa_Home.Visible
                Globals.Ribbons.rbn_Styles.tab_aa_Home.Visible = Not tabVisibility
                '
        End Select

    End Sub
    '
    ''' <summary>
    ''' This method will set the visibility (variable isVisible) of the tab as specified by strTabID. 
    ''' Typically you would select strTabId from the class variables '_strTabId_Styles, _strTabId_Placeholders
    ''' _strTabId_PagesAndSections, _strTabId_Finalise.. If strTabId is set to 'all' then all of the
    ''' aa tabs' visibility is set to isVisible
    ''' </summary>
    Public Sub ctrl_tabSet_Visibility(strTabId As String, Optional isVisible As Boolean = True)
        '
        '
        Select Case strTabId
            Case _strTabId_Styles
                Globals.Ribbons.rbn_Styles.tab_aa_Styles.Visible = isVisible
                '
            Case _strTabId_Placeholders
                Globals.Ribbons.rbn_Styles.tab_aa_Placeholders.Visible = isVisible
                '
            Case _strTabId_PagesAndSections
                Globals.Ribbons.rbn_Styles.tab_aa_PagesAndSections.Visible = isVisible
                '
            Case _strTabId_Finalise
                Globals.Ribbons.rbn_Styles.tab_aa_Finalise.Visible = isVisible
                '
            Case _strTabId_AAHome
                Globals.Ribbons.rbn_Styles.tab_aa_Home.Visible = isVisible
                '
            Case "all"
                Globals.Ribbons.rbn_Styles.tab_aa_Styles.Visible = isVisible
                Globals.Ribbons.rbn_Styles.tab_aa_Placeholders.Visible = isVisible
                Globals.Ribbons.rbn_Styles.tab_aa_PagesAndSections.Visible = isVisible
                Globals.Ribbons.rbn_Styles.tab_aa_Finalise.Visible = isVisible
                'Globals.Ribbons.rbn_Styles.tab_aa_Home.Visible = Not isVisible
                Globals.Ribbons.rbn_Styles.tab_aa_Home.Visible = True

                '
        End Select

    End Sub
    '

    '
    ''' <summary>
    ''' This method will activate the specified tab. The use strTabId is meant to isolates the actual tab name from
    ''' any intenal uses...Typically you would select strTabId from the class variables '_strTabId_Styles, _strTabId_Placeholders
    ''' _strTabId_PagesAndSections, _strTabId_Finalise, _strTabId_AAHome
    ''' </summary>
    ''' <param name="strTabId"></param>
    Public Sub ctrl_tab_Activate(strTabId As String)
        Select Case strTabId
            Case _strTabId_Styles
                Globals.Ribbons.rbn_Styles.RibbonUI.ActivateTab("tab_aa_Styles")
                '
            Case _strTabId_Placeholders
                Globals.Ribbons.rbn_Styles.RibbonUI.ActivateTab("tab_aa_Placeholders")
                '
            Case _strTabId_PagesAndSections
                Globals.Ribbons.rbn_Styles.RibbonUI.ActivateTab("tab_aa_PagesAndSections")
                '
            Case _strTabId_Finalise
                Globals.Ribbons.rbn_Styles.RibbonUI.ActivateTab("tab_aa_Finalise")
                '
            Case _strTabId_AAHome
                Globals.Ribbons.rbn_Styles.RibbonUI.ActivateTab("tab_aa_Home")
                '
        End Select
    End Sub
    '
End Class

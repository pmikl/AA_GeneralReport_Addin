Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Public Class frm_findTables
    Public numFloatingTables As Integer
    Public objTablesMgr As New cTablesMgr()
    Public listOfFloatingTables As String
    Public lstOfPlhInfo_All As New List(Of cPlhInfo)
    Public lstOfDisplayPlhs As New List(Of cPlhInfo)
    '
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()
        '
        lbl_instruction00.Text = "Left click to select a placeholder, then right" + vbCrLf + "click to call up the context menu."
        lbl_Instruction_AllowedTypes.Text = "Floating and Irregular tables are not allowed" + vbCrLf + "in Accessible documents. Test for both"
        '
        Me.lstOfPlhInfo_All.Clear()
        Me.lstOfDisplayPlhs.Clear()
        ' Add any initialization after the InitializeComponent() call.
        '
        'Me.lbl_numFloatingTables.Text = "Floating placeholders in the document"
        ' Me.listOfFloatingTables = ""
        'Me.rTxtBx_floatingTables_location.SelectAll()
        'Me.rTxtBx_floatingTables_location.SelectionTabs = New Integer() {80, 120}
        'Me.btn_convertToInline.Enabled = False
        Me.frm_refresh()
        Me.frm_display()
        '
    End Sub
    '
    Public Sub frm_refresh()
        Dim objGlobals As New cGlobals()
        '
        objGlobals.glb_cursors_setToWait()
        '
        Try
            Me.lstBx_plhDetail.Items.Clear()
            Me.lstOfPlhInfo_All = Me.frm_get_plhDetails(objGlobals.glb_get_wrdActiveDoc)

        Catch ex As Exception

        End Try
        '
        'Me.Activate()
        'Me.lstBx_plhDetail.Refresh()
        '
        objGlobals.glb_cursors_setToNormal()

    End Sub
    '
    Public Function frm_get_plhDetails(ByRef myDoc As Word.Document) As List(Of cPlhInfo)
        Dim objPlhInfo As New cPlhInfo()
        Dim objGlobals As New cGlobals()
        Dim lstOfPlhs As New List(Of cPlhInfo)
        '
        Try
            myDoc = objGlobals.glb_get_wrdActiveDoc()
            lstOfPlhs = objPlhInfo.PlhDetail_get_allDetails(myDoc)
            '
        Catch ex As Exception
            Me.lstBx_plhDetail.Items.Clear()
        End Try
        '
        Return lstOfPlhs

    End Function
    '
    Public Sub frm_display()
        Dim objPlhInfo As cPlhInfo
        Dim doDisplayType As String
        Dim j As Integer
        '
        Me.lstBx_plhDetail.Items.Clear()
        Me.lstOfDisplayPlhs.Clear()
        '
        doDisplayType = ""
        '
        If Me.rdBtn_All.Checked Then doDisplayType = "All"
        If Me.rdBtn_Floating.Checked Then doDisplayType = "Floating"
        If Me.rdBtn_inLine.Checked Then doDisplayType = "in-Line"
        If Me.rdBtn_Irregular.Checked Then doDisplayType = "irregular"

        '
        'Fill the display list Me.lstOfDisplayPlhs from the list of all Placeholders
        'and then display the display list
        Select Case doDisplayType
            Case "All"
                For j = 0 To Me.lstOfPlhInfo_All.Count - 1
                    objPlhInfo = Me.lstOfPlhInfo_All.Item(j)
                    Me.lstOfDisplayPlhs.Add(objPlhInfo)
                Next
                '
            Case "Floating"
                For j = 0 To Me.lstOfPlhInfo_All.Count - 1
                    objPlhInfo = Me.lstOfPlhInfo_All.Item(j)
                    If objPlhInfo.isFloating Then Me.lstOfDisplayPlhs.Add(objPlhInfo)
                Next
                '
            Case "in-Line"
                For j = 0 To Me.lstOfPlhInfo_All.Count - 1
                    objPlhInfo = Me.lstOfPlhInfo_All.Item(j)
                    If Not objPlhInfo.isFloating Then Me.lstOfDisplayPlhs.Add(objPlhInfo)
                Next
                '
            Case "irregular"
                For j = 0 To Me.lstOfPlhInfo_All.Count - 1
                    objPlhInfo = Me.lstOfPlhInfo_All.Item(j)
                    If Not objPlhInfo.isRegular Then Me.lstOfDisplayPlhs.Add(objPlhInfo)
                Next
                '

        End Select
        '
        For j = 0 To Me.lstOfDisplayPlhs.Count - 1
            objPlhInfo = Me.lstOfDisplayPlhs.Item(j)
            Me.lstBx_plhDetail.Items.Add(objPlhInfo.strCaption)
        Next
        '
    End Sub
    '
    Public Sub frm_goto_selectedPlaceholder()
        Dim objPlhInfo As cPlhInfo
        Dim j As Integer
        '
        j = Me.lstBx_plhDetail.SelectedIndex
        '
        Try
            objPlhInfo = Me.lstOfDisplayPlhs.Item(j)
            objPlhInfo.PlhDetail_goto_object()
        Catch ex As Exception

        End Try
    End Sub
    '
    Public Sub frm_convert_toInline()
        Dim objPlhInfo As cPlhInfo
        Dim j As Integer
        '
        j = Me.lstBx_plhDetail.SelectedIndex
        '
        Try
            objPlhInfo = Me.lstOfDisplayPlhs.Item(j)
            objPlhInfo.PlhDetail_convert_inline()
        Catch ex As Exception

        End Try
    End Sub

    Private Sub btn_refresh_Click(sender As Object, e As EventArgs) Handles btn_refresh.Click
        Me.frm_refresh()
        Me.frm_display()
    End Sub

    Private Sub btn_goTo_Click(sender As Object, e As EventArgs)
        Me.frm_goto_selectedPlaceholder()
        '
    End Sub

    Private Sub GoToTheSelectedPlaceholderToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles GoToTheSelectedPlaceholderToolStripMenuItem.Click
        Me.frm_goto_selectedPlaceholder()
    End Sub

    Private Sub rdBtn_Floating_Click(sender As Object, e As EventArgs) Handles rdBtn_Floating.Click
        '
        Me.frm_refresh()
        Me.frm_display()
        '
    End Sub

    Private Sub rdBtn_All_Click(sender As Object, e As EventArgs) Handles rdBtn_All.Click

        Me.frm_refresh()
        Me.frm_display()
        '
    End Sub

    Private Sub rdBtn_inLine_Click(sender As Object, e As EventArgs) Handles rdBtn_inLine.Click
        '
        Me.frm_refresh()
        Me.frm_display()
        '
    End Sub
    Private Sub rdBtn_Irregular_Click(sender As Object, e As EventArgs) Handles rdBtn_Irregular.Click
        '
        Me.frm_refresh()
        Me.frm_display()
        '
    End Sub

    Private Sub btn_convertToInlineAll_Click(sender As Object, e As EventArgs) Handles btn_convertAlltoInLine.Click
        Dim objPlhInfo As cPlhInfo
        Dim j As Integer

        Try
            For j = 0 To Me.lstOfDisplayPlhs.Count - 1
                objPlhInfo = Me.lstOfDisplayPlhs.Item(j)
                objPlhInfo.PlhDetail_convert_inline()
                '
                '
            Next j
            '
            'Make this slow, so the authjor can watch it happening.. Otherwise put the
            'following outside the loop
            Me.frm_refresh()
            Me.frm_display()
            '
        Catch ex As Exception

        End Try
        '
        '
    End Sub

    Private Sub ConvertToInlineToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ConvertToInlineToolStripMenuItem.Click
        Dim objGlobals As New cGlobals()

        Me.frm_convert_toInline()
        Me.frm_display()
        MsgBox("Complete. The selected placeholder is now inline")
        objGlobals.glb_screen_update(True)
        '
    End Sub
    '
End Class
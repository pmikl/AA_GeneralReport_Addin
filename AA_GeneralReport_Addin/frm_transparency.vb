Public Class frm_transparency
    Public objBackPanelMgr As cBackPanelMgr
    '
    Public Sub New()
        Dim transparency As Single
        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        '
        Me.objBackPanelMgr = New cBackPanelMgr()
        '
        transparency = 0.0
        If Me.objBackPanelMgr.pnl_get_BackPanelTransparency(transparency, Me.objBackPanelMgr.glb_get_wrdSect) Then
            Me.scrl_Transparency.Value = CInt(transparency * 100)
            Me.txtBox_transparencyValue.Text = CStr(Me.scrl_Transparency.Value)
        End If
        '
    End Sub
    '
    Private Sub scrl_Transparency_Scroll(sender As Object, e As Windows.Forms.ScrollEventArgs) Handles scrl_Transparency.Scroll
        Dim setting_Current As Integer
        Dim transparency As Single
        Dim hScrollBar As Windows.Forms.HScrollBar
        '
        hScrollBar = sender
        '
        setting_Current = hScrollBar.Value
        Me.txtBox_transparencyValue.Text = CStr(setting_Current)
        '
        transparency = CSng(setting_Current / 100)
        Me.objBackPanelMgr.pnl_reset_BackPanelTransparency(transparency, Me.objBackPanelMgr.glb_get_wrdSect)
        '
    End Sub

    Private Sub scrl_Transparency_GotFocus(sender As Object, e As EventArgs) Handles scrl_Transparency.GotFocus

    End Sub

    Private Sub frm_transparency_GotFocus(sender As Object, e As EventArgs) Handles Me.GotFocus

    End Sub

    Private Sub frm_transparency_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub scrl_Transparency_Enter(sender As Object, e As EventArgs) Handles scrl_Transparency.Enter

    End Sub
    '

End Class
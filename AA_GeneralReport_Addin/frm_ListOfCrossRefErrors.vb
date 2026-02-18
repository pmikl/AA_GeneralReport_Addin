Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.Drawing
Public Class frm_ListOfCrossRefErrors
    Public myDoc As Word.Document
    Public lstOfCrossRefs As List(Of Field)

    Public Sub New(ByRef theDoc As Word.Document)

        ' This call is required by the designer.
        InitializeComponent()
        Me.myDoc = theDoc
        Me.txtBox_SourceDocument.Text = Me.myDoc.Name
        '
        'objTools.updateCrossReferenceFields()

        ' Add any initialization after the InitializeComponent() call.
        Me.lstOfCrossRefs = New List(Of Field)
        Me.fillList()
        '
        'Me.lbl_ListDescription.Text = "List of Orphaned Cross Reference Fields"
        '
        'rng.Select()
    End Sub
    '
    '
    Public Sub fillList()
        Dim objToolsMgr As New cTools()
        Dim objFlds As New cFieldsMgr()
        Dim fld As Word.Field
        Dim rng As Word.Range
        Dim showErrorCrossRefsOnly As Boolean
        Dim i As Integer
        Dim strFld As String
        Dim tokens() As String
        Dim doFldUpdate As Boolean
        '
        'Me.myDoc
        showErrorCrossRefsOnly = Me.chkBox_ShowAllCrossRefFields.Checked
        Me.lstBox_CrossRefErrors.Items.Clear()
        doFldUpdate = True
        '
        Me.lstOfCrossRefs = objFlds.flds_CrossReference_List(Me.myDoc, showErrorCrossRefsOnly, doFldUpdate)
        '
        If Me.lstOfCrossRefs.Count > 0 Then
            For i = 0 To Me.lstOfCrossRefs.Count - 1
                fld = Me.lstOfCrossRefs.Item(i)
                rng = fld.Result
                strFld = fld.Result.Text
                If strFld Like "*Error*" Then
                    Me.lstBox_CrossRefErrors.Items.Add("Cross Ref Field Error," + vbTab + "Page = " + CStr(rng.Information(WdInformation.wdActiveEndAdjustedPageNumber)))
                Else
                    tokens = Split(strFld, vbTab)
                    Me.lstBox_CrossRefErrors.Items.Add(tokens(0) + ",    " + vbTab + "Page = " + CStr(rng.Information(WdInformation.wdActiveEndAdjustedPageNumber)))

                End If

                'If fld.Result.Text Like "Error*" Then
                'Me.lstBox_CrossRefErrors.Items.Add("Cross Reference Error, Page Number = " + CStr(rng.Information(WdInformation.wdActiveEndAdjustedPageNumber)))
                'Me.lstBox_CrossRefErrors.Items.Add(fld.Result.Text + " Page = " + CStr(rng.Information(WdInformation.wdActiveEndAdjustedPageNumber)))
                'Else
                ' Me.lstBox_CrossRefErrors.Items.Add("Cross Reference, Page Number = " + CStr(rng.Information(WdInformation.wdActiveEndAdjustedPageNumber)))
                'End If
                'Me.lstBox_CrossRefErrors.Refresh()
            Next
            Me.SelectAListItem(0)
        End If
    End Sub
    '
    Public Function SelectAListItem(itemIndex As Integer) As Boolean
        Dim fld As Word.Field
        Dim rng As Word.Range
        Dim allIsOK As Boolean
        '
        allIsOK = False

        Try
            Me.myDoc.Activate()
            Me.lstBox_CrossRefErrors.SelectedIndex = itemIndex
            fld = Me.lstOfCrossRefs.Item(itemIndex)
            '
            rng = fld.Result
            rng.Select()
            '
            allIsOK = True
        Catch ex As Exception
            allIsOK = False
        End Try
        '
        Return allIsOK
    End Function
    '
    '
    Public Sub RefreshTheList()
        Try
            Me.txtBox_SourceDocument.Text = Me.myDoc.Name
            Me.fillList()
        Catch ex As Exception
            MsgBox("The source document is no longer avialable")
            Me.Close()
        End Try
    End Sub
    '
    Public Sub DeleteSelectedField()
        Dim i As Integer
        Dim fld As Word.Field

        i = Me.lstBox_CrossRefErrors.SelectedIndex
        fld = Me.lstOfCrossRefs.Item(i)
        '
        fld.Delete()
        '
    End Sub
    '
    Public Sub DeleteAllFields()
        Dim fld As Word.Field
        Dim i As Integer
        '
        Try
            For i = Me.lstOfCrossRefs.Count - 1 To 0 Step -1
                fld = Me.lstOfCrossRefs.Item(i)
                fld.Delete()
                Me.fillList()
            Next
        Catch ex As Exception
            MsgBox("Could not delete all Fields")
            Try
                Me.fillList()
            Catch ex2 As Exception
                Me.Close()
            End Try
        End Try
    End Sub
    '

    Private Sub btn_Close_Click(sender As Object, e As EventArgs) Handles btn_Close.Click
        Me.Close()
    End Sub

    Private Sub btn_Refresh_Click(sender As Object, e As EventArgs) Handles btn_Refresh.Click
        Me.RefreshTheList()
    End Sub


    Private Sub lstBox_CrossRefErrors_Click(sender As Object, e As EventArgs) Handles lstBox_CrossRefErrors.Click
        Dim i As Integer
        Dim allIsOK As Boolean

        Try
            i = Me.lstBox_CrossRefErrors.SelectedIndex
            allIsOK = Me.SelectAListItem(i)
            If Not allIsOK Then
                MsgBox("Error in the Fields list.. Try refreshing it")
            End If
            '
        Catch ex As Exception
        End Try
    End Sub
    '

    Private Sub DeleteSelectedFieldToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteSelectedFieldToolStripMenuItem.Click
        '
        Try
            Me.DeleteSelectedField()
            Me.RefreshTheList()
        Catch ex As Exception
            MsgBox("Error in field delete... Did you delet this manually. Try refreshing the form")
        End Try
        '
    End Sub
    '
    Public Function testForDocument() As Boolean
        Dim rslt As Boolean
        '
        rslt = False
        Try
            Me.myDoc.Activate()
            rslt = True
        Catch ex As Exception
            rslt = False
        End Try
        '
        Return rslt
    End Function

    Private Sub CloseToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles CloseToolStripMenuItem.Click
        Me.Close()
    End Sub

    Private Sub RefreshToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshToolStripMenuItem.Click
        Me.RefreshTheList()
    End Sub

    Private Sub DeleteALLOfTheFieldsInTheListToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles DeleteALLOfTheFieldsInTheListToolStripMenuItem.Click
        Me.DeleteAllFields()
    End Sub

    Private Sub lbl_Instruction_Click(sender As Object, e As EventArgs) Handles lbl_Instruction.Click

    End Sub

    Private Sub lstBox_CrossRefErrors_DrawItem(sender As Object, e As System.Windows.Forms.DrawItemEventArgs) Handles lstBox_CrossRefErrors.DrawItem
        'Dim lstBox As ListBox
        'Dim fld As Word.Field
        '
        'lstBox = sender
        'e.DrawBackground()
        '
        ' MessageBox.Show("Draw item")
        'fld = Me.lstOfErrors.Item(e.Index)
        'If the current index is on the Faults List then display it as red
        'If fld.Result.Text Like "Error*" Then
        'e.Graphics.DrawString(lstBox.Items(e.Index).ToString(), lstBox.Font, Brushes.Red, e.Bounds)
        'Else
        'e.Graphics.DrawString(lstBox.Items(e.Index).ToString(), lstBox.Font, Brushes.Black, e.Bounds)
        'e.Graphics.DrawString(lstBox.Items(e.Index).ToString(), lstBox.Font, Brushes.Red, e.Bounds)
        'End If
    End Sub

    Private Sub chkBox_ShowAllCrossRefFields_Click(sender As Object, e As EventArgs) Handles chkBox_ShowAllCrossRefFields.Click
        Me.RefreshTheList()
    End Sub


End Class
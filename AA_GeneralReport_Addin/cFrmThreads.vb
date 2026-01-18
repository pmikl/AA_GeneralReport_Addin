Public Class cFrmThreads
    Public Delegate Sub timeOutMessage(ByVal sender As Object, ByVal e As testEventArgs)
    Public Delegate Sub conversionError(ByVal sender As Object, ByVal e As testEventArgs)
    Public Event sendTimeOut As timeOutMessage
    Public Event sendError As conversionError
    '
    Public _saveDirectoryFullName As String
    Public _conversionMode As String
    Public docsArray As String()
    Public Sub New()
        Me._conversionMode = ""
        Me._saveDirectoryFullName = ""
        Me.docsArray = Nothing
    End Sub
    '
    Public Sub convertSelectedDocs()
        Dim objWCAGMgr As cWCAGMgr
        Dim objNETMgr As cDotNetHandler
        Dim strCurrentDocPath As String
        'Dim newDoc As Word.Document
        Dim objFileMgr As New cFileHandler()
        Dim i As Integer
        Dim strConversionType As String
        '
        strConversionType = ""
        '
        RaiseEvent sendTimeOut(Me, New testEventArgs("start", 0))
        '
        For i = 0 To Me.docsArray.Count - 1
            '
            RaiseEvent sendTimeOut(Me, New testEventArgs("loop", i))
            '
            'If Me.frm.flg_killConversion = 1 Then Exit For
            '
            'Me.frm.lstBox_listOfDocs.SelectedIndex = i
            'strCurrentDocPath = Me.frm.docsArray(i)
            '
            strCurrentDocPath = Me.docsArray(i)
            '
            Select Case Me._conversionMode
                Case "wcag"
                    objWCAGMgr = New cWCAGMgr(strCurrentDocPath, Me._saveDirectoryFullName, "WCAG")
                    objWCAGMgr.convertSelectedDoc()

                    If objWCAGMgr._convertError Then
                        'Me.frm.lstOfFaults.Add(i)
                        RaiseEvent sendError(Me, New testEventArgs("loop", i))
                    End If
                    '
                    '
                Case "dotNET"
                    objNETMgr = New cDotNetHandler(strCurrentDocPath, Me._saveDirectoryFullName, "NET")
                    '
                    'strConversionType = objNETMgr._strConvertType_VBA_to_Env
                    strConversionType = objNETMgr._strConvertTag_dotNET_Env_to_TandG
                    'strConversionType = objNETMgr._strConvertType_VBA_to_TandG
                    '
                    objNETMgr.convertSelectedDoc()
                    '
                    If objNETMgr._convertError Then
                        'Me.lstOfFaults.Add(i)
                        RaiseEvent sendError(Me, New testEventArgs("loop", i))
                    End If
            End Select
            '
        Next
        '
        RaiseEvent sendTimeOut(Me, New testEventArgs("end", 0))

        '
    End Sub

    '
    Public Sub convertSelectedDocsToWCAG()
        'Dim objWCAGMgr As cWCAGMgr
        'Dim strCurrentDocPath, strFolder As String
        'Dim newDoc As Word.Document
        Dim objFileMgr As New cFileHandler()
        Dim i As Integer
        'Dim drawEvent As [Delegate]
        'Dim evt As eve
        '

        For i = 0 To 5
            '
            '
            'Select Case Me._conversionMode
            'Case "wcag"
            'objWCAGMgr = New cWCAGMgr(strCurrentDocPath, Me.frm._saveDirectoryFullName, "WCAG")
            'objWCAGMgr.convertSelectedDoc()

            'If objWCAGMgr._convertError Then
            'Me.frm.lstOfFaults.Add(i)
            'End If
            'Case "dotNET"
            'MessageBox.Show("WCAG Conversion")
            'End Select
            '
        Next
        '
        '
    End Sub

End Class
Public Class testEventArgs
    Inherits EventArgs

    Public _eventStatus As String
    Public _loopCounter As Int32
    Public Sub New(strEventStatus As String, loopCounter As Int32)
        MyBase.New()
        Me._eventStatus = strEventStatus
        Me._loopCounter = loopCounter
    End Sub
End Class

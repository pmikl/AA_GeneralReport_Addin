Imports System.IO
Imports System.Windows.Forms
Public Class frm_cryptTools
    Public objCrypt As cCrypt
    '
    Public Sub New()
        ' This call is required by the designer.
        InitializeComponent()
        '
        Me.objCrypt = New cCrypt()
        '
    End Sub
    '
    Public Function frm_get_SHAType() As String
        Dim strType As String
        '
        strType = ""
        '
        If Me.rdBtn_SHA1.Checked Then strType = "SHA1"
        If Me.rdBtn_SHA256.Checked Then strType = "SHA256"
        If Me.rdBtn_SHA384.Checked Then strType = "SHA384"
        If Me.rdBtn_SHA512.Checked Then strType = "SHA512"
        '
        Return strType
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will return the Hash of the selected sourceFile. It will do so on the basis of
    ''' the value of strHashName which can be SHA1, SHA256, SHA384, SHA512
    ''' </summary>
    ''' <param name="strHashType"></param>
    ''' <param name="strFileFullPath"></param>
    ''' <returns></returns>
    Public Function frm_getHash_asString(strHashType As String, strFileFullPath As String) As String
        Dim objCrypt As New cCrypt()
        Dim strHash As String
        '
        strHash = ""
        'MsgBox(rdBtn.Name + " Path = " + Me.txtBox_FilePath.Text)

        Try
            'If it exists on disk
            If Not (Path.GetDirectoryName(strFileFullPath) = "") Then
                Select Case strHashType
                    Case "SHA1"
                        strHash = objCrypt.cryptSHA_get_SHAasString(strFileFullPath, 1)
                    Case "SHA256"
                        strHash = objCrypt.cryptSHA_get_SHAasString(strFileFullPath, 256)
                    Case "SHA384"
                        strHash = objCrypt.cryptSHA_get_SHAasString(strFileFullPath, 384)
                    Case "SHA512"
                        strHash = objCrypt.cryptSHA_get_SHAasString(strFileFullPath, 512)
                    Case Else
                        strHash = ""
                End Select

            End If
        Catch ex As Exception

        End Try
        '
        Return strHash
        '
    End Function
    '
    Public Function frm_getHash_asString(strFileFullPath As String, Optional doInVBNETFormat As Boolean = False) As String
        Dim rdBtn As RadioButton
        Dim strResult, strHashType As String
        Dim strDirectory, strFileName, strExtension As String
        Dim sb As New StringBuilder()
        Dim tokens As String()
        Dim j As Integer
        '
        rdBtn = Nothing
        strResult = ""
        strHashType = Me.frm_get_SHAType()
        '
        strDirectory = Path.GetDirectoryName(strFileFullPath)
        strFileName = Path.GetFileName(strFileFullPath)
        strExtension = Path.GetExtension(strFileFullPath)
        '        
        Me.txtBox_Hash.Text = ""
        '
        If Not Me.txtBox_FilePath.Text = "" Then
            If Not strHashType = "" Then
                'sourceFile = New FileInfo(Me.txtBox_FilePath.Text)
                If Not (strDirectory = "") Then
                    strResult = Me.frm_getHash_asString(strHashType, strFileFullPath)
                    If Not doInVBNETFormat Then
                        Me.txtBox_Hash.Text = strResult
                    Else
                        tokens = strResult.Split(" ")
                        sb.Append("New Byte() {")
                        For j = 0 To tokens.Length - 1
                            If j <> tokens.Length - 1 Then
                                sb.Append("&H" + tokens(j) + ", ")
                            Else
                                sb.Append("&H" + tokens(j) + "}")
                            End If
                        Next
                        Me.txtBox_Hash.Text = Trim(sb.ToString())

                    End If
                End If
            End If
        End If
        '
        Return strResult
        '
    End Function

    '
    Public Function frm_getHash_asString(Optional doInVBNETFormat As Boolean = False) As String
        Dim rdBtn As RadioButton
        'Dim sourceFile As FileInfo
        Dim strResult, strHashType, strFileFullPath As String
        Dim sb As New StringBuilder()
        Dim tokens As String()
        Dim j As Integer
        '
        rdBtn = Nothing
        strResult = ""
        strHashType = Me.frm_get_SHAType()
        '        
        Me.txtBox_Hash.Text = ""
        '
        If Not Me.txtBox_FilePath.Text = "" Then
            strFileFullPath = Me.txtBox_FilePath.Text
            If Not strHashType = "" Then
                'sourceFile = New FileInfo(Me.txtBox_FilePath.Text)
                If Not (Path.GetDirectoryName(strFileFullPath) = "") Then
                    strResult = Me.frm_getHash_asString(strHashType, strFileFullPath)
                    If Not doInVBNETFormat Then
                        Me.txtBox_Hash.Text = strResult
                    Else
                        tokens = strResult.Split(" ")
                        sb.Append("New Byte() {")
                        For j = 0 To tokens.Length - 1
                            If j <> tokens.Length - 1 Then
                                sb.Append("&H" + tokens(j) + ", ")
                            Else
                                sb.Append("&H" + tokens(j) + "}")
                            End If
                        Next
                        Me.txtBox_Hash.Text = Trim(sb.ToString())

                    End If
                End If
            End If
        Else
            MsgBox("Please select a file")
        End If
        '
        Return strResult
        '
    End Function

    Private Sub rdBtn_SHA1_Click(sender As Object, e As EventArgs) Handles rdBtn_SHA1.Click

    End Sub

    Private Sub rdBtn_SHA256_CheckedChanged(sender As Object, e As EventArgs) Handles rdBtn_SHA256.CheckedChanged

    End Sub

    Private Sub rdBtn_SHA384_CheckedChanged(sender As Object, e As EventArgs) Handles rdBtn_SHA384.CheckedChanged

    End Sub

    Private Sub rdBtn_SHA512_CheckedChanged(sender As Object, e As EventArgs) Handles rdBtn_SHA512.CheckedChanged

    End Sub

    Private Sub btn_getFileSHA_Click(sender As Object, e As EventArgs) Handles btn_getFileSHA.Click
        Dim doInVBNETFormat As Boolean = False
        '
        frm_getHash_asString(doInVBNETFormat)
        '
    End Sub

    Private Sub SelectFileToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles SelectFileToolStripMenuItem.Click
        Dim objFileHandler As New cFileHandler()
        Dim strFileNames As String()
        Dim strDlgFilter As String = "all files (*.*)|*.*"
        Dim strStartDir As String

        Me.txtBox_FilePath.Text = ""
        '
        strStartDir = objFileHandler.file_getDir_documentsLocal()
        strFileNames = objFileHandler.file_get_filesFromDlg(strStartDir, "Get Files", strDlgFilter, False)
        '
        If Not IsNothing(strFileNames) Then
            If Not (strFileNames.Length = 0) Then
                Try
                    Me.txtBox_FilePath.Text = strFileNames(0)
                    '
                Catch ex As Exception

                End Try
            Else
                MsgBox("No file selected")
            End If
            '
        End If
        '
        '
    End Sub

    Private Sub btn_getSHA_asVBNET_Click(sender As Object, e As EventArgs) Handles btn_getSHA_asVBNET.Click
        Dim doInVBNETFormat As Boolean = True
        '
        frm_getHash_asString(doInVBNETFormat)
        '

    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click

    End Sub

    Private Sub btn_cipherFile_Click(sender As Object, e As EventArgs) Handles btn_cipherFile.Click

    End Sub

    Private Sub btn_decryptTheFile_Click(sender As Object, e As EventArgs) Handles btn_decryptTheFile.Click

    End Sub

    Private Sub btn_GenerateRandomKey_Click(sender As Object, e As EventArgs) Handles btn_GenerateRandomKey.Click
        Dim objConverter As New cConverter()
        Dim aesKey As Byte()
        Dim sbAesKey As New StringBuilder()
        '
        aesKey = Me.objCrypt.cryptAES_key_createVectorPair(Me.txtBox_AESKey.Text)
        '
        'Me.txtBox_GeneratedKey.Text = objConverter.ConvertByteArrayToHexString(Me.objCrypt.fCurrentKey()).ToString()
        'Me.txtBox_GeneratedIV.Text = objConverter.ConvertByteArrayToHexString(Me.objCrypt.fCurrentIV()).ToString()
        '
    End Sub
End Class
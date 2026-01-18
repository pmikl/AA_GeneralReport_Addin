Imports Microsoft.Office.Interop.Word
Imports Microsoft.Office.Core
Imports System.IO
'
Public Class cRemoteResources
    Inherits cCrypt
    '
    Public Sub New()
        MyBase.New()
    End Sub
    '
    ''' <summary>
    ''' This method will go to the Acil Allen web site and get the image file specified by strImageName
    ''' (e.g. "artwork_contactsPage_front_release.png"). If
    ''' defaultToWeb is set to true it will always go to the web. If it is set to false it will check for
    ''' a file with the name strImageName in Documents\AAC Images and if it is there it will use this. Typically
    ''' this is faster, but may result in out of date information. The location Documents\AAC Images  needs to be
    ''' cleaned out first to ensure the web is the primary source. The method returns the fullPath name of the
    ''' retrieved file. If no file was retrieved it will return an empty string.
    ''' </summary>
    ''' <param name="defaultToWeb"></param>
    ''' <param name="strFileName"></param>
    ''' <returns></returns>
    Public Function remRsrc_get_fileFromWeb(strWebPath As String, strFileName As String, checkFingerPrint As Boolean, Optional defaultToWeb As Boolean = True) As String
        Dim rslt As Boolean
        Dim i As Integer
        Dim objWCAG As New cWCAGMgr()
        Dim objBBMgr As New cBBlocksHandler()
        'Dim strWebPath As String
        Dim strFullWebPath As String
        Dim strLocalFileFullPath As String
        Dim strFolderPath As String
        Dim myFileInfo As FileInfo
        Dim timeOut As Integer
        Dim myHash, myStoredFingerPrint As Byte()
        '
        'strWebPath = "http://templates.acilallen.com.au/word/images/"
        strFullWebPath = strWebPath + strFileName
        '
        timeOut = 2000              'milliseconds
        rslt = False
        strFolderPath = ""
        strLocalFileFullPath = ""
        myHash = Nothing
        '
        '**
        i = 1
        '**
        '
        Try
            strFolderPath = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments) + "\AAC Images"
            strLocalFileFullPath = strFolderPath + "\" + strFileName
            myFileInfo = New FileInfo(strLocalFileFullPath)
            '
            'If the folder does not exist, create it.. If it does exist, then delete the image file if it exists
            If Directory.Exists(strFolderPath) Then
                If myFileInfo.Exists Then
                    myFileInfo.Delete()
                End If
            Else
                Directory.CreateDirectory(strFolderPath)
                '
            End If
            '
            '
            If defaultToWeb Then
                'Don't check for local copy. Go straight to the web and get a new copy
                myFileInfo = New FileInfo(strLocalFileFullPath)
                If myFileInfo.Exists Then
                    myFileInfo.Delete()
                Else
                    My.Computer.Network.DownloadFile(strFullWebPath, strLocalFileFullPath, "", "", False, timeOut, True)
                    'Ensure that its there
                    myFileInfo = New FileInfo(strLocalFileFullPath)
                    If Not myFileInfo.Exists Then
                        strLocalFileFullPath = ""
                        GoTo finis
                    End If
                    '
                End If
            Else
                'Check for a local copy and use that in preference
                myFileInfo = New FileInfo(strLocalFileFullPath)
                '
                If Not myFileInfo.Exists Then
                    My.Computer.Network.DownloadFile(strFullWebPath, strLocalFileFullPath, "", "", False, timeOut, True)
                    myFileInfo = New FileInfo(strLocalFileFullPath)
                    If Not myFileInfo.Exists Then
                        strLocalFileFullPath = ""
                        GoTo finis
                    End If
                End If
            End If
            '
            'iShp = rng.InlineShapes.AddPicture(strLocalFileFullPath)
        Catch ex As Exception
            strLocalFileFullPath = ""
        End Try
        '
        'Dim strHash As String
        'strHash = Me.cryptSHA_get_SHAasString(New IO.FileInfo(strLocalFileFullPath))
        'Dim frm As New frm_getHash()
        'frm.txtBox_hash.Text = strHash
        'frm.Show()
        '
finis:
        Try
            If strLocalFileFullPath <> "" And checkFingerPrint Then
                'We need to check the fingerprint hash
                myHash = Me.cryptSHA_get_SHA(strLocalFileFullPath)
                myStoredFingerPrint = Me.remRsrc_get_fileStoredFingerPrint(strFileName)
                '
                'If the hash of the downloaded file does not match the stored finger print, then delete the
                'local storage and return null for the file path
                If Not Me.crypt_SHA_Compare(myHash, myStoredFingerPrint) Then
                    Directory.Delete(strFolderPath, True)
                    Directory.CreateDirectory(strFolderPath)
                    strLocalFileFullPath = ""
                End If
            End If
            '
        Catch ex As Exception
            strLocalFileFullPath = ""
        End Try
        '
        '
        Return strLocalFileFullPath

        '
    End Function
    '
    ''' <summary>
    ''' This method will return the valid SHA for the specified file name. This is just the name and not
    ''' the path
    ''' </summary>
    ''' <param name="strFileName"></param>
    ''' <returns></returns>
    Public Function remRsrc_get_fileStoredFingerPrint(strFileName As String) As Byte()
        Dim validHash As Byte()
        '
        validHash = Nothing
        '
        Try
            Select Case strFileName
                Case "artwork_contactsPage_front_release.png"
                    '300dpi version
                    'validHash = New Byte() {&H68, &HEC, &HB2, &H63, &HEA, &HE8, &H4C, &H9, &H2, &HD7, &HF6, &HE7, &H53, &H5F, &H50, &HCA, &H55, &H34, &H99, &HC0, &HF3, &HFA, &H8C, &H2, &HE2, &H46, &H9C, &H61, &H7C, &H66, &HA8, &HDA, &HA, &H65, &H8C, &HB4, &H39, &H40, &H92, &H88, &HF, &H7E, &HDD, &H29, &HAE, &HE0, &H39, &H7B, &H24, &H3E, &HAF, &HAA, &HBC, &HD7, &HAA, &HD9, &H18, &H8, &H38, &H64, &H8A, &HF8, &H8, &H6D}

                    '150dpi version
                    validHash = New Byte() {&H87, &H8F, &HEF, &H9B, &H43, &H2F, &H7F, &H40, &H66, &HE4, &H82, &HE0, &HF9, &H2E, &HA6, &H91, &HEC, &H4D, &HD5, &HCC, &HF9, &H28, &H65, &H1, &HA0, &H3, &HE3, &HFD, &H6B, &HDB, &HB9, &HD1, &H11, &HB9, &HD8, &HEB, &HAD, &H78, &H2F, &HD, &H28, &H8F, &H79, &HD0, &H36, &H4F, &HAB, &HB7, &HE5, &HEE, &H38, &H75, &H7C, &H6, &H66, &H9, &HD0, &H37, &HBD, &HF1, &H35, &H8B, &H4E, &H2F}
            End Select
            '
        Catch ex As Exception

        End Try
        '
        Return validHash
        '
    End Function

    ''' <summary>
    ''' This method will retrieve the specified image (strImageName) from either the local store at Documents\AAC Images
    ''' or from the AAC web site. If defaulToWeb is true it will go directly to the web site, and on the way it will delete any
    ''' local copies, therebye ensuring that our local copy is fresh
    ''' </summary>
    ''' <param name="rng"></param>
    ''' <param name="width"></param>
    ''' <param name="strImageName"></param>
    ''' <returns></returns>
    Public Function remRsrc_insert_imageFromWeb(ByRef rng As Word.Range, width As Single, strImageName As String,
                                              Optional strWebPath As String = "http://templates.acilallen.com.au/word/images/") As InlineShape
        Dim iShp As InlineShape
        Dim strLocalFileFullPath As String
        Dim checkFingerPrint As Boolean = True
        '
        'strImageName = "aac_pict_indigenous_00.png"
        strLocalFileFullPath = Me.remRsrc_get_fileFromWeb(strWebPath, strImageName, checkFingerPrint)
        iShp = Nothing
        '
        Try
            If Not (strLocalFileFullPath = "") Then
                iShp = rng.InlineShapes.AddPicture(strLocalFileFullPath)
                iShp.LockAspectRatio = True
                iShp.Width = width
                '
            End If
        Catch ex As Exception
            iShp = Nothing
        End Try
        '
        '
        Return iShp
        '
    End Function
    '
End Class

Imports System
Imports System.Runtime.Serialization
Imports System.Runtime.Serialization.Formatters.Binary

Imports System.Security.Cryptography
Imports System.Collections
Imports System.Data
Imports System.Text
Imports System.IO
Imports System.Xml
Public Class cCrypt
    'Encryption
    Public objRijnd As RijndaelManaged = New RijndaelManaged()
    '
    'Coding
    Public AsciiCoder As ASCIIEncoding = New ASCIIEncoding()
    Public UniCodeCoder As UnicodeEncoding = New UnicodeEncoding()
    '
    Public Property vKeySize As Integer
        'Constructor to get and set the Key size
        Get
            Return Me.objRijnd.KeySize
        End Get
        Set(ByVal value As Integer)
            Me.objRijnd.KeySize = value
        End Set
        '
    End Property
    '
    Public Property fCurrentKey As Byte() 'Gets or set the current Key
        Get
            Return Me.objRijnd.Key
        End Get
        Set(ByVal value As Byte())
            Me.objRijnd.Key = value
        End Set
    End Property
    '
    Public Property fCurrentIV As Byte() 'Gets or set the current Initialisation Vector
        Get
            Return Me.objRijnd.IV
        End Get
        Set(ByVal value As Byte())
            Me.objRijnd.IV = value
        End Set
    End Property
    '
    Public Enum CryptoAction
        'Define the enumeration for CryptoAction.
        ActionEncrypt = 1
        ActionDecrypt = 2
    End Enum
    Public Sub New()
        '
        Me.objRijnd.KeySize = 256
        Me.objRijnd.Padding = PaddingMode.Zeros
        fCurrentKey = New Byte() {&HF5, &HA6, &HF1, &HFA, &HC9, &HF6, &H7, &H92, &H2C, &HFC, &H27, &H32, &HFE, &H6A, &H5B, &H1A, &HB, &HCB, &HB9, &H4E, &HC5, &H12, &HF, &H76, &HBE, &HEC, &H69, &H1C, &HD1, &H53, &HDE, &H36}
        fCurrentIV = New Byte() {&H88, &H71, &H51, &H4A, &HF2, &H4, &HA, &HD3, &HF6, &H96, &H69, &HEC, &HA9, &HF4, &H4F, &HF1}
        '
    End Sub
    '
#Region "AES"
    '
    ''' <summary>
    ''' This method will take the password and create a Hash (SHA512) in this case
    ''' which we treat as the 512 bit (64 byte) Vector Pair (Key and IV). The first
    ''' 32 bytes (256 bits) are extracted and used for the Key... The next 16 bytes (128 bits)
    ''' are extracted and used as the IV vector... The Me.fCurrentKey and Me.fCurrentIV are
    ''' set accordingly, whilst the Me.vKeySize is set to 256 bits (32 * 8).
    ''' 
    ''' </summary>
    ''' <param name="strPassword"></param>
    ''' <returns></returns>
    Public Function cryptAES_key_createVectorPair(strPassword As String) As Byte()
        Dim chrData() As Char = strPassword.ToCharArray                     'Convert strPassword to an array and store in chrData.
        Dim intLength As Integer = chrData.GetUpperBound(0)                 'Use intLength to get strPassword size.
        Dim bytDataToHash(intLength) As Byte                                'Declare bytDataToHash and make it the same size as chrData.
        Dim SHA512 As New System.Security.Cryptography.SHA512Managed        'Declare what hash to use.
        Dim bytResult As Byte()                                             'Declare bytResult, .
        Dim bytKey(31), bytIV(15) As Byte
        '
        Me.vKeySize = 32 * 8
        '
        'Declare bytKey(31).  It will hold 256 bits.
        '
        'Use For Next to convert and store chrData into bytDataToHash.
        For i As Integer = 0 To chrData.GetUpperBound(0)
            bytDataToHash(i) = CByte(Asc(chrData(i)))
        Next
        '
        bytResult = SHA512.ComputeHash(bytDataToHash)                       'Hash bytDataToHash and store it in bytResult
        '
        'Use For Next to put a specific size (256 bits) of 
        'bytResult into bytKey. The 0 To 31 will put the first 256 bits
        'of 512 bits into bytKey...As an added layer of security we can
        'choose any part of the 512 bit set
        '
        For i As Integer = 0 To 31
            'Me.fCurrentKey(i) = bytResult(i)
            bytKey(i) = bytResult(i)
        Next
        '
        Me.fCurrentKey = bytKey
        '
        'Now for the IV vector use the next 16 bytes
        For i As Integer = 32 To 47
            bytIV(i - 32) = bytResult(i)
        Next
        '
        Me.fCurrentIV = bytIV
        '
        'Set cCrypt keysize
        '
        Return Me.fCurrentKey 'Return the key.
        '
    End Function
    '
    Public Function cryptAES_key_createKeyOnly(strPassword As String) As Byte()
        Dim chrData() As Char = strPassword.ToCharArray                     'Convert strPassword to an array and store in chrData.
        Dim intLength As Integer = chrData.GetUpperBound(0)                 'Use intLength to get strPassword size.
        Dim bytDataToHash(intLength) As Byte                                'Declare bytDataToHash and make it the same size as chrData.
        Dim SHA512 As New System.Security.Cryptography.SHA512Managed        'Declare what hash to use.
        Dim bytResult As Byte()                                             'Declare bytResult, .
        Dim bytKey(31), bytIV(15) As Byte
        '
        Me.vKeySize = 32 * 8
        '
        'Declare bytKey(31).  It will hold 256 bits.
        '
        'Use For Next to convert and store chrData into bytDataToHash.
        For i As Integer = 0 To chrData.GetUpperBound(0)
            bytDataToHash(i) = CByte(Asc(chrData(i)))
        Next
        '
        bytResult = SHA512.ComputeHash(bytDataToHash)                       'Hash bytDataToHash and store it in bytResult
        '
        'Use For Next to put a specific size (256 bits) of 
        'bytResult into bytKey. The 0 To 31 will put the first 256 bits
        'of 512 bits into bytKey...As an added layer of security we can
        'choose any part of the 512 bit set
        '
        For i As Integer = 0 To 31
            'Me.fCurrentKey(i) = bytResult(i)
            bytKey(i) = bytResult(i)
        Next
        '
        Me.fCurrentKey = bytKey
        '
        Return Me.fCurrentKey 'Return the key.
        '
    End Function

    '
    ''' <summary>
    ''' This method independently sets the IV vector (16 bytes) without affecting the
    ''' Key vector
    ''' </summary>
    ''' <param name="strPassword"></param>
    ''' <returns></returns>
    Public Function cryptAES_key_createIV(strPassword As String) As Byte()
        Dim chrData() As Char
        Dim intLength As Integer

        'Convert strPassword to an array and store in chrData, then use intLength to get strPassword size.
        chrData = strPassword.ToCharArray
        intLength = chrData.GetUpperBound(0)
        'Declare bytDataToHash and make it the same size as chrData.
        Dim bytDataToHash(intLength) As Byte

        'Use For Next to convert and store chrData into bytDataToHash.
        For i As Integer = 0 To chrData.GetUpperBound(0)
            bytDataToHash(i) = CByte(Asc(chrData(i)))
        Next

        'Declare what hash to use.
        Dim SHA512 As New System.Security.Cryptography.SHA512Managed
        'Declare bytResult, Hash bytDataToHash and store it in bytResult.
        Dim bytResult As Byte() = SHA512.ComputeHash(bytDataToHash)
        'Declare bytIV(15).  It will hold 128 bits.
        Dim bytIV(15) As Byte

        'Use For Next to put a specific size (128 bits) of bytResult into bytIV.
        'The 0 To 30 for bytKey used the first 256 bits of the hashed password.
        'The 32 To 47 will put the next 128 bits into bytIV.
        For i As Integer = 32 To 47
            bytIV(i - 32) = bytResult(i)
        Next
        '
        Me.fCurrentIV = bytIV
        '
        Return Me.fCurrentIV 'Return the IV.

    End Function
    '
    '
    ''' <summary>
    ''' This method will take the password and create a Hash (SHA512) in this case.
    ''' 256 bits of that Hash are returned in the byte array as the Key. Note that
    ''' we can return any 256 block of the 512 bits. Currently we choose to
    ''' return the first 256 bits
    ''' </summary>
    ''' <param name="strPassword"></param>
    ''' <returns></returns>
    Public Function cryptAES_get_Key(strPassword As String) As Byte()
        Dim chrData() As Char = strPassword.ToCharArray                     'Convert strPassword to an array and store in chrData.
        Dim intLength As Integer = chrData.GetUpperBound(0)                 'Use intLength to get strPassword size.
        Dim bytDataToHash(intLength) As Byte                                'Declare bytDataToHash and make it the same size as chrData.
        Dim SHA512 As New System.Security.Cryptography.SHA512Managed        'Declare what hash to use.
        Dim bytResult As Byte()                                             'Declare bytResult, .
        Dim bytKey(31) As Byte                                              'Declare bytKey(31).  It will hold 256 bits.
        '
        'Use For Next to convert and store chrData into bytDataToHash.
        For i As Integer = 0 To chrData.GetUpperBound(0)
            bytDataToHash(i) = CByte(Asc(chrData(i)))
        Next
        '
        bytResult = SHA512.ComputeHash(bytDataToHash)                       'Hash bytDataToHash and store it in bytResult
        '
        'Use For Next to put a specific size (256 bits) of 
        'bytResult into bytKey. The 0 To 31 will put the first 256 bits
        'of 512 bits into bytKey...As an added layer of security we can
        'choose any part of the 512 bit set
        For i As Integer = 0 To 31
            bytKey(i) = bytResult(i)
        Next
        '
        Me.fCurrentKey = bytKey
        '
        Return bytKey 'Return the key.
        '
    End Function
    '
    Public Function cryptAES_get_IV(strPassword As String) As Byte()
        Dim chrData() As Char
        Dim intLength As Integer

        'Convert strPassword to an array and store in chrData, then use intLength to get strPassword size.
        chrData = strPassword.ToCharArray
        intLength = chrData.GetUpperBound(0)
        'Declare bytDataToHash and make it the same size as chrData.
        Dim bytDataToHash(intLength) As Byte

        'Use For Next to convert and store chrData into bytDataToHash.
        For i As Integer = 0 To chrData.GetUpperBound(0)
            bytDataToHash(i) = CByte(Asc(chrData(i)))
        Next

        'Declare what hash to use.
        Dim SHA512 As New System.Security.Cryptography.SHA512Managed
        'Declare bytResult, Hash bytDataToHash and store it in bytResult.
        Dim bytResult As Byte() = SHA512.ComputeHash(bytDataToHash)
        'Declare bytIV(15).  It will hold 128 bits.
        Dim bytIV(15) As Byte

        'Use For Next to put a specific size (128 bits) of bytResult into bytIV.
        'The 0 To 30 for bytKey used the first 256 bits of the hashed password.
        'The 32 To 47 will put the next 128 bits into bytIV.
        For i As Integer = 32 To 47
            bytIV(i - 32) = bytResult(i)
        Next
        '
        Me.fCurrentIV = bytIV

        Return bytIV 'Return the IV.
    End Function
    '
    '
    ''' <summary>
    ''' This method will take a Bytes_In byte array and return a transformed byte array. The transformation
    ''' is in accordance with Rijndael AES with key and IV as input byte arrays. The TransFormType is either
    ''' CryptoAction.ActionEncrypt or Crypto.ActionDecrypt. If the method fails, the tranformed output is et
    ''' to Nothing...This method has been tested 20220617
    ''' </summary>
    ''' <param name="Bytes_In"></param>
    ''' <param name="key"></param>
    ''' <param name="IV"></param>
    ''' <param name="TransformType"></param>
    ''' <returns></returns>
    Public Function cryptAES_transform_BytesToBytes(ByRef Bytes_In As Byte(), ByRef key As Byte(), ByRef IV As Byte(), TransformType As CryptoAction) As Byte()
        Dim transformed As Byte()
        Dim transformBlock As ICryptoTransform
        '
        'https://stackoverflow.com/questions/20603747/encrypting-and-decrypting-a-byte-array
        '
        transformBlock = Nothing
        transformed = Nothing
        '
        Try
            Using rijAlg As Rijndael = Rijndael.Create()
                ' Create an Rijndael object with the specified key and IV.
                '
                rijAlg.Key = key
                rijAlg.IV = IV
                '
                ' Create a transformBlock to perform the stream transform (i.e. encrypt or decrypt)
                Select Case TransformType
                    Case CryptoAction.ActionEncrypt
                        transformBlock = rijAlg.CreateEncryptor(rijAlg.Key, rijAlg.IV)
                    Case CryptoAction.ActionDecrypt
                        transformBlock = rijAlg.CreateDecryptor(rijAlg.Key, rijAlg.IV)
                End Select
                '
                ' Create the streams used for encryption.
                Using msTransformed As MemoryStream = New MemoryStream()
                    Using csEncrypt As CryptoStream = New CryptoStream(msTransformed, transformBlock, CryptoStreamMode.Write)
                        csEncrypt.Write(Bytes_In, 0, Bytes_In.Length)
                        csEncrypt.Close()
                        'Using swEncrypt As StreamWriter = New StreamWriter(csEncrypt)
                        'Write all data to the stream.
                        'swEncrypt.Write(Bytes_In)
                        'End Using
                        transformed = msTransformed.ToArray()
                    End Using
                    msTransformed.Close()
                End Using

            End Using
            '
        Catch ex As Exception
            transformed = Nothing
        End Try
        '
        ' Return the encrypted bytes from the memory stream.
        Return transformed
    End Function

    ''' <summary>
    ''' This method will take as input a plain text string (strPlainText) and produce a byte array of ciphered bytes
    ''' </summary>
    ''' <param name="strPlainText"></param>
    ''' <param name="Key"></param>
    ''' <param name="IV"></param>
    ''' <returns></returns>
    Public Function cryptAES_transform_plainTextStringToCipheredBytes(ByVal strPlainText As String, ByRef Key As Byte(), ByRef IV As Byte()) As Byte()
        Dim encrypted As Byte()
        Dim transformBlock As ICryptoTransform
        '
        encrypted = Nothing
        transformBlock = Nothing
        '
        ' Check arguments.
        If strPlainText = "" Or IsNothing(Key) Or IsNothing(IV) Then
            GoTo finis
        End If
        '
        Try
            Using rijAlg As Rijndael = Rijndael.Create()
                ' Create an Rijndael object with the specified key and IV.
                rijAlg.Key = Key
                rijAlg.IV = IV

                ' Create a transformBlock to perform the stream transform (i.e. encrypt or decrypt)
                transformBlock = rijAlg.CreateEncryptor(rijAlg.Key, rijAlg.IV)
                '
                ' Create the streams used for encryption.
                Using msTransformed As MemoryStream = New MemoryStream()
                    Using csEncrypt As CryptoStream = New CryptoStream(msTransformed, transformBlock, CryptoStreamMode.Write)
                        Using swEncrypt As StreamWriter = New StreamWriter(csEncrypt)

                            'Write all data to the stream.
                            swEncrypt.Write(strPlainText)
                        End Using
                        encrypted = msTransformed.ToArray()
                    End Using
                End Using
            End Using
            '
        Catch ex As Exception

        End Try
        '
finis:
        ' Return the encrypted bytes from the memory stream.
        Return encrypted
    End Function
    '
    '
    ''' <summary>
    ''' This method will take as input a byte array of 'cipheredBytes' and convert them to a plain text string. If the ciphered
    ''' bytes were produced from a string, then this method will recover the string
    ''' </summary>
    ''' <param name="cipheredBytes"></param>
    ''' <param name="Key"></param>
    ''' <param name="IV"></param>
    ''' <returns></returns>
    Public Function cryptAES_transform_cipheredBytesToPlainTextString(ByRef cipheredBytes As Byte(), ByRef Key As Byte(), ByRef IV As Byte()) As String
        '
        Dim transformBlock As ICryptoTransform
        ' Declare the string used to hold
        ' the decrypted text.
        Dim plaintext As String = Nothing
        '
        '
        ' Check arguments.
        If IsNothing(cipheredBytes) Or IsNothing(Key) Or IsNothing(IV) Then
            GoTo finis
        End If
        '

        ' Create an Rijndael object
        ' with the specified key and IV.
        Using rijAlg As Rijndael = Rijndael.Create()
            rijAlg.Key = Key
            rijAlg.IV = IV

            ' Create a decryptor to perform the stream transform.
            transformBlock = rijAlg.CreateDecryptor(rijAlg.Key, rijAlg.IV)

            ' Create the streams used for decryption.
            Using msDecrypt As MemoryStream = New MemoryStream(cipheredBytes)
                Using csDecrypt As CryptoStream = New CryptoStream(msDecrypt, transformBlock, CryptoStreamMode.Read)
                    Using srDecrypt As StreamReader = New StreamReader(csDecrypt)

                        ' Read the decrypted bytes from the decrypting stream
                        ' and place them in a string.
                        plaintext = srDecrypt.ReadToEnd()
                    End Using
                End Using
            End Using
        End Using

finis:
        Return plaintext
    End Function
    '
    '
    ''' <summary>
    ''' This method will take the input string strPlainText and cipher it to a byte array of ciphered bytes, and then return that
    ''' byte array as a string of Hex 2 digit numbers.
    ''' as hex digits
    ''' </summary>
    ''' <param name="strPlainText"></param>
    ''' <param name="Key"></param>
    ''' <param name="IV"></param>
    ''' <returns></returns>
    Public Function cryptAES_transform_plainTextStringToCipheredHexString(ByVal strPlainText As String, ByRef Key As Byte(), ByRef IV As Byte()) As String
        Dim cipherText As Byte()
        Dim objConverter As New cConverter()
        Dim sb As StringBuilder
        '
        cipherText = Me.cryptAES_transform_plainTextStringToCipheredBytes(strPlainText, Key, IV)
        sb = cConverter.ConvertByteArrayToHexString(cipherText)
        '
        'strResult = cryptAES_cipher_ByteArrayToHexString(plainTextAsByteArray, Direction)
        Return sb.ToString()
        '
    End Function
    '
    '
    ''' <summary>
    ''' This method will take the input string strCipheredTextAsHexNumbers (expecting it to be a ciphered representation of a plain text
    ''' string). It will transfor the Hex numbers to intermediate byte array of bytes (still ciphered) and the decipher those bytes
    ''' to a plain text string
    ''' </summary>
    ''' <param name="strCipheredTextAsHexNumbers"></param>
    ''' <param name="Key"></param>
    ''' <param name="IV"></param>
    ''' <returns></returns>
    Public Function cryptAES_transform_cipheredHexStringToPlainTextString(ByVal strCipheredTextAsHexNumbers As String, ByVal Key As Byte(), ByVal IV As Byte()) As String
        Dim cipherText As Byte()
        Dim strResult As String
        Dim objConverter As New cConverter()
        Dim lstOfHexNums = New List(Of String)()
        '
        strResult = ""
        '
        cipherText = cConverter.convert_hexString_ToByteArray(strCipheredTextAsHexNumbers)
        '
        strResult = Me.cryptAES_transform_cipheredBytesToPlainTextString(cipherText, Key, IV)
        '
        Return strResult
        '
    End Function

    '
    '***** Being tested
    '
    ''' <summary>
    ''' This method will use AES to cipher/decipher the input file represented by fileInfo_In to the file
    ''' represented by fileInfo_Out. The transform action is detremined by the value of Direction.
    ''' Tested both ways 20220617 OK (graphic ok, word document a problem but can be recovered)
    ''' </summary>
    ''' <param name="srcFileInFo"></param>
    ''' <param name="Direction"></param>
    Public Function cryptAES_transform_FileToFile(ByRef srcFileInFo As FileInfo, ByRef key As Byte(), ByRef IV As Byte(), Direction As CryptoAction) As FileInfo
        Dim objFileMgr As New cFileHandler()
        Dim destFileInfo As FileInfo
        Dim fsDestination As FileStream
        Dim transformBlock As ICryptoTransform = Nothing
        Dim strFileOut_FullPath As String
        Dim csEncrypt As CryptoStream
        Dim myBuffer(4096) As Byte              'bytes for processing
        Dim lngBytesProcessed As Long = 0       'running count of bytes processed
        Dim inputFileLength As Long             'the input file's length
        Dim intBytesInCurrentBlock As Integer   'current bytes being processed
        Dim strFileType As String
        '
        'https://www.codeproject.com/Articles/12092/Encrypt-Decrypt-Files-in-VB-NET-Using-Rijndael
        '
        '
        strFileType = "aes"
        csEncrypt = Nothing
        destFileInfo = Nothing
        '
        If Not srcFileInFo.Exists Then GoTo finis
        '
        'GoTo finis
        '
        Try
            'Using rijAlg As Rijndael = Rijndael.Create()
            Dim rijAlg As New RijndaelManaged()
            rijAlg.Key = key
            rijAlg.IV = IV
            '
            Select Case Direction
                Case CryptoAction.ActionEncrypt
                    strFileType = "ciphered"
                    transformBlock = rijAlg.CreateEncryptor(Me.fCurrentKey, Me.fCurrentIV)
                    '
                Case CryptoAction.ActionDecrypt
                    strFileType = "deCiphered"
                    transformBlock = rijAlg.CreateDecryptor(Me.fCurrentKey, Me.fCurrentIV)
                    '
            End Select
            '
            strFileOut_FullPath = objFileMgr.file_get_newFileName(srcFileInFo, srcFileInFo.DirectoryName, strFileType)
            destFileInfo = New FileInfo(strFileOut_FullPath)
            '
            'Using fsSource As FileStream = New FileStream(srcFileInFo.FullName, FileMode.Open, FileAccess.Read)
            Dim fsSource As FileStream = New FileStream(srcFileInFo.FullName, FileMode.Open, FileAccess.Read)
            inputFileLength = fsSource.Length
            '
            If Not destFileInfo.Exists Then
                fsDestination = New FileStream(destFileInfo.FullName, FileMode.OpenOrCreate, FileAccess.Write)
            Else
                fsDestination = New FileStream(destFileInfo.FullName, FileMode.Truncate, FileAccess.Write)
            End If
            '
            fsDestination.SetLength(0)
            '
            '
            csEncrypt = New CryptoStream(fsDestination, transformBlock, CryptoStreamMode.Write)
            lngBytesProcessed = 0
            '
            'Use While to loop until all of the file is processed.
            'Read file with the input filestream, then Write output file with the cryptostream and
            'contnue till finished
            While lngBytesProcessed < inputFileLength
                intBytesInCurrentBlock = fsSource.Read(myBuffer, 0, 4096)
                csEncrypt.Write(myBuffer, 0, intBytesInCurrentBlock)
                lngBytesProcessed = lngBytesProcessed + CLng(intBytesInCurrentBlock)
            End While
            '
            fsDestination.Close()
            fsSource.Close()
            'csEncrypt.Close()
            '
            'End Using


            ' End Using

        Catch ex As Exception
            MsgBox("Error in file to file transform")
        End Try

finis:
        Return destFileInfo
        '
    End Function

    '
    ''' <summary>
    ''' This method will use AES to cipher/decipher the input file represented by fileInfo_In to the file
    ''' represented by fileInfo_Out. The transform action is detremined by the value of Direction
    ''' </summary>
    ''' <param name="fileInFo_In"></param>
    ''' <param name="fileInfo_Out"></param>
    ''' <param name="Direction"></param>
    Public Sub cryptAES_cipher_fileTofile_x(ByRef fileInFo_In As FileInfo, ByRef fileInfo_Out As FileInfo, Direction As CryptoAction)
        Dim fs_Out, fs_In As FileStream
        Dim cipherTransform As ICryptoTransform = Nothing
        Dim cipherBlock As CryptoStream
        Dim myBuffer(4096) As Byte              'bytes for processing
        Dim lngBytesProcessed As Long = 0       'running count of bytes processed
        Dim inputFileLength As Long             'the input file's length
        Dim intBytesInCurrentBlock As Integer   'current bytes being processed
        '
        'https://www.codeproject.com/Articles/12092/Encrypt-Decrypt-Files-in-VB-NET-Using-Rijndael
        '
        If Not fileInFo_In.Exists Then GoTo finis
        '
        cipherBlock = Nothing
        '
        fs_In = New FileStream(fileInFo_In.FullName, FileMode.Open, FileAccess.Read)
        inputFileLength = fs_In.Length
        '
        Try
            If Not fileInfo_Out.Exists Then
                fs_Out = New FileStream(fileInfo_Out.FullName, FileMode.OpenOrCreate, FileAccess.Write)
            Else
                fs_Out = New FileStream(fileInfo_Out.FullName, FileMode.Truncate, FileAccess.Write)
            End If
            '
            fs_Out.SetLength(0)
            '
            Select Case Direction
                Case CryptoAction.ActionEncrypt
                    cipherTransform = Me.objRijnd.CreateEncryptor(Me.fCurrentKey, Me.fCurrentIV)
                    cipherBlock = New CryptoStream(fs_In, cipherTransform, CryptoStreamMode.Write)
                    '
                Case CryptoAction.ActionDecrypt
                    cipherTransform = Me.objRijnd.CreateDecryptor(Me.fCurrentKey, Me.fCurrentIV)
                    cipherBlock = New CryptoStream(fs_Out, cipherTransform, CryptoStreamMode.Write)
                    '
            End Select
            '
            lngBytesProcessed = 0
            '
            'Use While to loop until all of the file is processed.
            'Read file with the input filestream, then Write output file with the cryptostream and
            'contnue till finished
            While lngBytesProcessed < inputFileLength
                intBytesInCurrentBlock = fs_In.Read(myBuffer, 0, 4096)
                cipherBlock.Write(myBuffer, 0, intBytesInCurrentBlock)
                lngBytesProcessed = lngBytesProcessed + CLng(intBytesInCurrentBlock)
            End While


            'Close FileStreams and CryptoStream. Do we need to flush
            'fs_In.Flush()
            cipherBlock.Close()
            fs_In.Close()
            fs_Out.Close()
            '
        Catch ex As Exception
            'fs_In.Dispose
        End Try


finis:
    End Sub

    '
    '***** Being tested

    'serialize object to memory stream
    Public Shared Function SerializeToStream(ByVal o As Object) As MemoryStream
        Dim stream As MemoryStream = New MemoryStream()
        Dim formatter As IFormatter = New BinaryFormatter()
        formatter.Serialize(stream, o)
        Return stream
    End Function
    '


    'deserialize object from memory stream
    Public Shared Function DeserializeFromStream(Of T As New)(ByVal memoryStream As MemoryStream) As T
        If memoryStream Is Nothing Then
            Return New T()
        End If
        Dim o As T
        Dim binaryFormatter As BinaryFormatter = New BinaryFormatter()
        Using memoryStream
            memoryStream.Seek(0, SeekOrigin.Begin)
            o = CType(binaryFormatter.Deserialize(memoryStream), T)
        End Using
        Return o
    End Function
    '


    '

    '***** Everything below this line os not tested






    ''' <summary>
    ''' This method will take the input byte array plainText, cipher it and then return the result as a string,
    ''' representing each byte as Hex digits
    ''' </summary>
    ''' <param name="plainText"></param>
    ''' <param name="Direction"></param>
    ''' <returns></returns>
    Public Function cryptAES_cipher_ByteArrayToHexString(ByRef plainText As Byte(), ByRef key As Byte(), ByRef IV As Byte(), Direction As CryptoAction) As String
        Dim cipherText As Byte()
        Dim sb As New StringBuilder()
        Dim j As Integer
        '
        cipherText = Me.cryptAES_transform_BytesToBytes(plainText, key, IV, Direction)
        '
        For j = 0 To cipherText.Length - 1
            sb.Append(cipherText(j).ToString("x").PadLeft(2, "0").ToUpper())
            sb.Append(" ")
        Next
        '
        Return sb.ToString()
    End Function

    '
    ''' <summary>
    ''' This method will take the input string plainText and cipher it using AES. the ciphered bytes
    ''' are returned as a Byte Array
    ''' </summary>
    ''' <param name="plainText"></param>
    ''' <param name="Direction"></param>
    ''' <returns></returns>
    Public Function cryptAES_cipher_stringToByteArray(ByRef plainText As String, ByRef key As Byte(), ByRef IV As Byte(), Direction As CryptoAction) As Byte()
        Dim plainTextAsByteArray As Byte()
        '
        plainTextAsByteArray = Me.UniCodeCoder.GetBytes(plainText)
        '
        Return Me.cryptAES_transform_BytesToBytes(plainTextAsByteArray, key, IV, CryptoAction.ActionEncrypt)
        '
    End Function




#End Region

#Region "SHA"
    '
    ''' <summary>
    ''' This routine will read any file and generate the SHA256 for that file.
    ''' The digits of the digest are returned in the StringBuilder. H0 is the
    ''' first Hex digit. The type of SHA is selected by SHATypeSelect
    ''' 
    ''' SHATypeSelect = 1		(SHA1)
    ''' SHATypeSelect = 256		(SHA256)
    ''' SHATypeSelect = 384		(SHA384)
    ''' SHATypeSelect = 512		(SHA512)
    ''' Test Status:		
    ''' </summary>
    ''' <param name="strFileFullPath"></param>
    ''' <returns></returns>
    Public Function cryptSHA_get_SHA(ByRef strFileFullPath As String, Optional SHATypeSelect As Integer = 512) As Byte()
        'Create a instance of the SHA1 (Managed) type
        Dim result As Byte()
        Dim fs As FileStream
        Dim shaM As SHA1Managed = New SHA1Managed()
        Dim sha256M As SHA256Managed
        Dim sha384M As SHA384Managed
        Dim sha512M As SHA512Managed
        Dim sb As StringBuilder = New StringBuilder()
        '
        result = Nothing
        '
        Try
            fs = New FileStream(strFileFullPath, FileMode.OpenOrCreate, FileAccess.Read)
            '
            Select Case SHATypeSelect
                Case 1
                    shaM = New SHA1Managed()
                    result = shaM.ComputeHash(fs)
                    fs.Close()
                '
                Case 256
                    sha256M = New SHA256Managed
                    result = sha256M.ComputeHash(fs)
                    fs.Close()
                '
                Case 384
                    sha384M = New SHA384Managed()
                    result = sha384M.ComputeHash(fs)
                    fs.Close()
                '
                Case 512
                    sha512M = New SHA512Managed()
                    result = sha512M.ComputeHash(fs)
                    fs.Close()
                    '
            End Select

        Catch ex As Exception
            result = Nothing
        End Try
        '
        Return result
    End Function
    '
    Public Function cryptSHA_get_SHAasString(ByRef strFileFullPath As String, Optional SHATypeSelect As Integer = 512) As String
        Dim strHash As String
        Dim theHash As Byte()
        Dim sb As New StringBuilder()
        Dim i As Integer
        '
        strHash = ""
        theHash = Me.cryptSHA_get_SHA(strFileFullPath, SHATypeSelect)
        '
        If Not IsNothing(theHash) Then
            For i = 0 To theHash.Length - 1
                sb.Append(theHash(i).ToString("x").PadLeft(2, "0").ToUpper())
                sb.Append(" ")
            Next
            '
            strHash = sb.ToString()
            strHash = Trim(strHash)
            '
        End If
        '
        Return strHash
        '
    End Function
    '
    ''' <summary>
    ''' This method will compare the two SHA byte arrays, SHA1 and SHA2. If they are the same
    ''' then the method will return true
    ''' </summary>
    ''' <param name="SHA1"></param>
    ''' <param name="SHA2"></param>
    ''' <returns></returns>
    Public Function crypt_SHA_Compare(ByRef SHA1 As Byte(), ByRef SHA2 As Byte()) As Boolean
        Dim rslt As Boolean
        Dim digit As Byte
        Dim j As Integer
        '
        rslt = False
        digit = Nothing
        '
        Try
            rslt = False
            If SHA1.Length <> SHA2.Length Then GoTo finis
            '
            For j = 0 To SHA1.Length - 1
                If SHA1(j) <> SHA2(j) Then
                    rslt = False
                    GoTo finis
                End If
            Next
            '
            rslt = True
            '
        Catch ex As Exception
            rslt = False
        End Try
        '
finis:
        Return rslt
        '
    End Function

    '
    ''' <summary>
    ''' This method will compare two SHA values. They are input as the strings SHA1
    ''' and SHA2. The input is formated as Hex digits separated by a space
    ''' (i.e. XX XX XX XX ). The method will return true if the two SHA
    ''' values are the same and false if they are not.. Note that the method
    ''' will always return false if the number of digits in each SHA list
    ''' is not the same
    ''' </summary>
    ''' <param name="SHA1"></param>
    ''' <param name="SHA2"></param>
    ''' <returns></returns>
    Public Function crypt_SHA_Compare(ByVal SHA1 As String, ByVal SHA2 As String) As Boolean
        Dim objPars As cParser = New cParser()
        Dim objCvt As New cConverter()
        Dim bytesOfSHA1 As Byte()
        Dim bytesOfSHA2 As Byte()
        Dim delim = New Char(1) {}

        delim(0) = " "
        delim(1) = Convert.ToChar(&H0)

        'If any part of this fails, then return false
        Try
            'Now convert to actual byte values and compare, returning false
            'if they are not the same
            bytesOfSHA1 = cConverter.convert_hexString_ToByteArray(SHA1)
            bytesOfSHA2 = cConverter.convert_hexString_ToByteArray(SHA2)
            '
            If bytesOfSHA1.Length <> bytesOfSHA2.Length Then Return False
            '
            'Now lets see if they are the same
            For i = 0 To bytesOfSHA1.Length - 1
                If bytesOfSHA1(i) <> bytesOfSHA2(i) Then
                    Return False
                End If
            Next
            '
        Catch
            Return False
        End Try

        Return True
    End Function

    '
#End Region

End Class

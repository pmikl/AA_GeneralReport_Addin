<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_cryptTools
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Me.grp_textCipher = New System.Windows.Forms.GroupBox()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.txtBox_cipherText = New System.Windows.Forms.TextBox()
        Me.txtBox_plainText = New System.Windows.Forms.TextBox()
        Me.btn_getSHA_asVBNET = New System.Windows.Forms.Button()
        Me.btn_getFileSHA = New System.Windows.Forms.Button()
        Me.grpBox_SHA = New System.Windows.Forms.GroupBox()
        Me.rdBtn_SHA512 = New System.Windows.Forms.RadioButton()
        Me.rdBtn_SHA384 = New System.Windows.Forms.RadioButton()
        Me.rdBtn_SHA256 = New System.Windows.Forms.RadioButton()
        Me.txtBox_Hash = New System.Windows.Forms.TextBox()
        Me.rdBtn_SHA1 = New System.Windows.Forms.RadioButton()
        Me.lbl_selectedFile = New System.Windows.Forms.Label()
        Me.txtBox_FilePath = New System.Windows.Forms.TextBox()
        Me.ContextMenuStrip1 = New System.Windows.Forms.ContextMenuStrip(Me.components)
        Me.SelectFileToolStripMenuItem = New System.Windows.Forms.ToolStripMenuItem()
        Me.dlg_getFileName = New System.Windows.Forms.OpenFileDialog()
        Me.lbl_AESKey = New System.Windows.Forms.Label()
        Me.txtBox_AESKey = New System.Windows.Forms.TextBox()
        Me.txtBox_AESIV = New System.Windows.Forms.TextBox()
        Me.lbl_AESIV = New System.Windows.Forms.Label()
        Me.btn_getRandomPassword = New System.Windows.Forms.Button()
        Me.txtBox_cipheredFilePath = New System.Windows.Forms.TextBox()
        Me.btn_cipherFile = New System.Windows.Forms.Button()
        Me.lbl_randomPasswordInfo = New System.Windows.Forms.Label()
        Me.lbl_encryptTheCurrentFile = New System.Windows.Forms.Label()
        Me.btn_decryptTheFile = New System.Windows.Forms.Button()
        Me.grpBox_AES = New System.Windows.Forms.GroupBox()
        Me.tbCtrl_Keys = New System.Windows.Forms.TabControl()
        Me.tb_AES = New System.Windows.Forms.TabPage()
        Me.lbl_Key = New System.Windows.Forms.Label()
        Me.txtBox_GeneratedKey = New System.Windows.Forms.TextBox()
        Me.btn_GenerateRandomKey = New System.Windows.Forms.Button()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.txtBox_GeneratedIV = New System.Windows.Forms.TextBox()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.tb_RSA = New System.Windows.Forms.TabPage()
        Me.grp_textCipher.SuspendLayout()
        Me.grpBox_SHA.SuspendLayout()
        Me.ContextMenuStrip1.SuspendLayout()
        Me.grpBox_AES.SuspendLayout()
        Me.tbCtrl_Keys.SuspendLayout()
        Me.tb_AES.SuspendLayout()
        Me.SuspendLayout()
        '
        'grp_textCipher
        '
        Me.grp_textCipher.Controls.Add(Me.Button2)
        Me.grp_textCipher.Controls.Add(Me.Button1)
        Me.grp_textCipher.Controls.Add(Me.txtBox_cipherText)
        Me.grp_textCipher.Controls.Add(Me.txtBox_plainText)
        Me.grp_textCipher.Location = New System.Drawing.Point(359, 22)
        Me.grp_textCipher.Name = "grp_textCipher"
        Me.grp_textCipher.Size = New System.Drawing.Size(554, 274)
        Me.grp_textCipher.TabIndex = 14
        Me.grp_textCipher.TabStop = False
        Me.grp_textCipher.Text = "GroupBox1"
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(244, 207)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(77, 23)
        Me.Button2.TabIndex = 3
        Me.Button2.Text = "<-- De Cipher"
        Me.Button2.UseVisualStyleBackColor = True
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(244, 148)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 2
        Me.Button1.Text = "Cipher -->"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'txtBox_cipherText
        '
        Me.txtBox_cipherText.Location = New System.Drawing.Point(327, 23)
        Me.txtBox_cipherText.Multiline = True
        Me.txtBox_cipherText.Name = "txtBox_cipherText"
        Me.txtBox_cipherText.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtBox_cipherText.Size = New System.Drawing.Size(221, 236)
        Me.txtBox_cipherText.TabIndex = 1
        '
        'txtBox_plainText
        '
        Me.txtBox_plainText.Location = New System.Drawing.Point(14, 23)
        Me.txtBox_plainText.Multiline = True
        Me.txtBox_plainText.Name = "txtBox_plainText"
        Me.txtBox_plainText.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtBox_plainText.Size = New System.Drawing.Size(221, 236)
        Me.txtBox_plainText.TabIndex = 0
        Me.txtBox_plainText.Text = "E1 28 AE E8 2C 0C 70 45 A5 AF 9A 6F DE 73 51 74"
        '
        'btn_getSHA_asVBNET
        '
        Me.btn_getSHA_asVBNET.Location = New System.Drawing.Point(167, 84)
        Me.btn_getSHA_asVBNET.Name = "btn_getSHA_asVBNET"
        Me.btn_getSHA_asVBNET.Size = New System.Drawing.Size(175, 23)
        Me.btn_getSHA_asVBNET.TabIndex = 13
        Me.btn_getSHA_asVBNET.Text = "Get File SHA in vb.NET Format"
        Me.btn_getSHA_asVBNET.UseVisualStyleBackColor = True
        '
        'btn_getFileSHA
        '
        Me.btn_getFileSHA.Location = New System.Drawing.Point(16, 84)
        Me.btn_getFileSHA.Name = "btn_getFileSHA"
        Me.btn_getFileSHA.Size = New System.Drawing.Size(113, 23)
        Me.btn_getFileSHA.TabIndex = 12
        Me.btn_getFileSHA.Text = "Get File SHA"
        Me.btn_getFileSHA.UseVisualStyleBackColor = True
        '
        'grpBox_SHA
        '
        Me.grpBox_SHA.Controls.Add(Me.rdBtn_SHA512)
        Me.grpBox_SHA.Controls.Add(Me.rdBtn_SHA384)
        Me.grpBox_SHA.Controls.Add(Me.rdBtn_SHA256)
        Me.grpBox_SHA.Controls.Add(Me.txtBox_Hash)
        Me.grpBox_SHA.Controls.Add(Me.rdBtn_SHA1)
        Me.grpBox_SHA.Location = New System.Drawing.Point(12, 117)
        Me.grpBox_SHA.Name = "grpBox_SHA"
        Me.grpBox_SHA.Size = New System.Drawing.Size(341, 179)
        Me.grpBox_SHA.TabIndex = 11
        Me.grpBox_SHA.TabStop = False
        Me.grpBox_SHA.Text = "SHA"
        '
        'rdBtn_SHA512
        '
        Me.rdBtn_SHA512.AutoSize = True
        Me.rdBtn_SHA512.Checked = True
        Me.rdBtn_SHA512.Location = New System.Drawing.Point(258, 20)
        Me.rdBtn_SHA512.Name = "rdBtn_SHA512"
        Me.rdBtn_SHA512.Size = New System.Drawing.Size(68, 17)
        Me.rdBtn_SHA512.TabIndex = 3
        Me.rdBtn_SHA512.TabStop = True
        Me.rdBtn_SHA512.Text = "SHA 512"
        Me.rdBtn_SHA512.UseVisualStyleBackColor = True
        '
        'rdBtn_SHA384
        '
        Me.rdBtn_SHA384.AutoSize = True
        Me.rdBtn_SHA384.Location = New System.Drawing.Point(169, 20)
        Me.rdBtn_SHA384.Name = "rdBtn_SHA384"
        Me.rdBtn_SHA384.Size = New System.Drawing.Size(68, 17)
        Me.rdBtn_SHA384.TabIndex = 2
        Me.rdBtn_SHA384.Text = "SHA 384"
        Me.rdBtn_SHA384.UseVisualStyleBackColor = True
        '
        'rdBtn_SHA256
        '
        Me.rdBtn_SHA256.AutoSize = True
        Me.rdBtn_SHA256.Location = New System.Drawing.Point(77, 20)
        Me.rdBtn_SHA256.Name = "rdBtn_SHA256"
        Me.rdBtn_SHA256.Size = New System.Drawing.Size(68, 17)
        Me.rdBtn_SHA256.TabIndex = 1
        Me.rdBtn_SHA256.Text = "SHA 256"
        Me.rdBtn_SHA256.UseVisualStyleBackColor = True
        '
        'txtBox_Hash
        '
        Me.txtBox_Hash.Location = New System.Drawing.Point(7, 55)
        Me.txtBox_Hash.Multiline = True
        Me.txtBox_Hash.Name = "txtBox_Hash"
        Me.txtBox_Hash.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtBox_Hash.Size = New System.Drawing.Size(324, 109)
        Me.txtBox_Hash.TabIndex = 0
        '
        'rdBtn_SHA1
        '
        Me.rdBtn_SHA1.AutoSize = True
        Me.rdBtn_SHA1.Location = New System.Drawing.Point(7, 20)
        Me.rdBtn_SHA1.Name = "rdBtn_SHA1"
        Me.rdBtn_SHA1.Size = New System.Drawing.Size(56, 17)
        Me.rdBtn_SHA1.TabIndex = 0
        Me.rdBtn_SHA1.Text = "SHA 1"
        Me.rdBtn_SHA1.UseVisualStyleBackColor = True
        '
        'lbl_selectedFile
        '
        Me.lbl_selectedFile.AutoSize = True
        Me.lbl_selectedFile.Location = New System.Drawing.Point(13, 17)
        Me.lbl_selectedFile.Name = "lbl_selectedFile"
        Me.lbl_selectedFile.Size = New System.Drawing.Size(107, 13)
        Me.lbl_selectedFile.TabIndex = 10
        Me.lbl_selectedFile.Text = "Current File Selection"
        '
        'txtBox_FilePath
        '
        Me.txtBox_FilePath.ContextMenuStrip = Me.ContextMenuStrip1
        Me.txtBox_FilePath.Location = New System.Drawing.Point(16, 36)
        Me.txtBox_FilePath.Multiline = True
        Me.txtBox_FilePath.Name = "txtBox_FilePath"
        Me.txtBox_FilePath.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtBox_FilePath.Size = New System.Drawing.Size(327, 41)
        Me.txtBox_FilePath.TabIndex = 9
        '
        'ContextMenuStrip1
        '
        Me.ContextMenuStrip1.Items.AddRange(New System.Windows.Forms.ToolStripItem() {Me.SelectFileToolStripMenuItem})
        Me.ContextMenuStrip1.Name = "ContextMenuStrip1"
        Me.ContextMenuStrip1.Size = New System.Drawing.Size(127, 26)
        '
        'SelectFileToolStripMenuItem
        '
        Me.SelectFileToolStripMenuItem.Name = "SelectFileToolStripMenuItem"
        Me.SelectFileToolStripMenuItem.Size = New System.Drawing.Size(126, 22)
        Me.SelectFileToolStripMenuItem.Text = "Select &File"
        '
        'dlg_getFileName
        '
        Me.dlg_getFileName.FileName = "OpenFileDialog1"
        '
        'lbl_AESKey
        '
        Me.lbl_AESKey.AutoSize = True
        Me.lbl_AESKey.Location = New System.Drawing.Point(7, 22)
        Me.lbl_AESKey.Name = "lbl_AESKey"
        Me.lbl_AESKey.Size = New System.Drawing.Size(80, 13)
        Me.lbl_AESKey.TabIndex = 0
        Me.lbl_AESKey.Text = "Password (Key)"
        '
        'txtBox_AESKey
        '
        Me.txtBox_AESKey.Location = New System.Drawing.Point(10, 39)
        Me.txtBox_AESKey.Name = "txtBox_AESKey"
        Me.txtBox_AESKey.Size = New System.Drawing.Size(143, 20)
        Me.txtBox_AESKey.TabIndex = 1
        Me.txtBox_AESKey.Text = "testPassword"
        '
        'txtBox_AESIV
        '
        Me.txtBox_AESIV.Location = New System.Drawing.Point(183, 39)
        Me.txtBox_AESIV.Name = "txtBox_AESIV"
        Me.txtBox_AESIV.Size = New System.Drawing.Size(143, 20)
        Me.txtBox_AESIV.TabIndex = 2
        Me.txtBox_AESIV.Text = "testIV"
        '
        'lbl_AESIV
        '
        Me.lbl_AESIV.AutoSize = True
        Me.lbl_AESIV.Location = New System.Drawing.Point(183, 22)
        Me.lbl_AESIV.Name = "lbl_AESIV"
        Me.lbl_AESIV.Size = New System.Drawing.Size(103, 13)
        Me.lbl_AESIV.TabIndex = 3
        Me.lbl_AESIV.Text = "Password Part 2 (IV)"
        '
        'btn_getRandomPassword
        '
        Me.btn_getRandomPassword.Location = New System.Drawing.Point(155, 39)
        Me.btn_getRandomPassword.Name = "btn_getRandomPassword"
        Me.btn_getRandomPassword.Size = New System.Drawing.Size(27, 20)
        Me.btn_getRandomPassword.TabIndex = 4
        Me.btn_getRandomPassword.Text = "Randomise"
        Me.btn_getRandomPassword.UseVisualStyleBackColor = True
        '
        'txtBox_cipheredFilePath
        '
        Me.txtBox_cipheredFilePath.Location = New System.Drawing.Point(10, 109)
        Me.txtBox_cipheredFilePath.Multiline = True
        Me.txtBox_cipheredFilePath.Name = "txtBox_cipheredFilePath"
        Me.txtBox_cipheredFilePath.ScrollBars = System.Windows.Forms.ScrollBars.Both
        Me.txtBox_cipheredFilePath.Size = New System.Drawing.Size(313, 44)
        Me.txtBox_cipheredFilePath.TabIndex = 5
        '
        'btn_cipherFile
        '
        Me.btn_cipherFile.Location = New System.Drawing.Point(10, 162)
        Me.btn_cipherFile.Name = "btn_cipherFile"
        Me.btn_cipherFile.Size = New System.Drawing.Size(160, 23)
        Me.btn_cipherFile.TabIndex = 6
        Me.btn_cipherFile.Text = "Encrypt the Current File"
        Me.btn_cipherFile.UseVisualStyleBackColor = True
        '
        'lbl_randomPasswordInfo
        '
        Me.lbl_randomPasswordInfo.AutoSize = True
        Me.lbl_randomPasswordInfo.Location = New System.Drawing.Point(53, 64)
        Me.lbl_randomPasswordInfo.Name = "lbl_randomPasswordInfo"
        Me.lbl_randomPasswordInfo.Size = New System.Drawing.Size(259, 13)
        Me.lbl_randomPasswordInfo.TabIndex = 7
        Me.lbl_randomPasswordInfo.Text = "Click the middle button to get a random password pair"
        '
        'lbl_encryptTheCurrentFile
        '
        Me.lbl_encryptTheCurrentFile.AutoSize = True
        Me.lbl_encryptTheCurrentFile.Location = New System.Drawing.Point(10, 90)
        Me.lbl_encryptTheCurrentFile.Name = "lbl_encryptTheCurrentFile"
        Me.lbl_encryptTheCurrentFile.Size = New System.Drawing.Size(74, 13)
        Me.lbl_encryptTheCurrentFile.TabIndex = 8
        Me.lbl_encryptTheCurrentFile.Text = "Encrypted File"
        '
        'btn_decryptTheFile
        '
        Me.btn_decryptTheFile.Location = New System.Drawing.Point(183, 162)
        Me.btn_decryptTheFile.Name = "btn_decryptTheFile"
        Me.btn_decryptTheFile.Size = New System.Drawing.Size(140, 23)
        Me.btn_decryptTheFile.TabIndex = 9
        Me.btn_decryptTheFile.Text = "Decrypt the File"
        Me.btn_decryptTheFile.UseVisualStyleBackColor = True
        '
        'grpBox_AES
        '
        Me.grpBox_AES.Controls.Add(Me.btn_decryptTheFile)
        Me.grpBox_AES.Controls.Add(Me.lbl_encryptTheCurrentFile)
        Me.grpBox_AES.Controls.Add(Me.lbl_randomPasswordInfo)
        Me.grpBox_AES.Controls.Add(Me.btn_cipherFile)
        Me.grpBox_AES.Controls.Add(Me.txtBox_cipheredFilePath)
        Me.grpBox_AES.Controls.Add(Me.btn_getRandomPassword)
        Me.grpBox_AES.Controls.Add(Me.lbl_AESIV)
        Me.grpBox_AES.Controls.Add(Me.txtBox_AESIV)
        Me.grpBox_AES.Controls.Add(Me.txtBox_AESKey)
        Me.grpBox_AES.Controls.Add(Me.lbl_AESKey)
        Me.grpBox_AES.Location = New System.Drawing.Point(572, 302)
        Me.grpBox_AES.Name = "grpBox_AES"
        Me.grpBox_AES.Size = New System.Drawing.Size(341, 194)
        Me.grpBox_AES.TabIndex = 15
        Me.grpBox_AES.TabStop = False
        Me.grpBox_AES.Text = "Advanced Encryption Standard (AES)"
        '
        'tbCtrl_Keys
        '
        Me.tbCtrl_Keys.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.tbCtrl_Keys.Controls.Add(Me.tb_AES)
        Me.tbCtrl_Keys.Controls.Add(Me.tb_RSA)
        Me.tbCtrl_Keys.Location = New System.Drawing.Point(12, 302)
        Me.tbCtrl_Keys.Name = "tbCtrl_Keys"
        Me.tbCtrl_Keys.SelectedIndex = 0
        Me.tbCtrl_Keys.Size = New System.Drawing.Size(341, 295)
        Me.tbCtrl_Keys.TabIndex = 16
        '
        'tb_AES
        '
        Me.tb_AES.Controls.Add(Me.lbl_Key)
        Me.tb_AES.Controls.Add(Me.txtBox_GeneratedKey)
        Me.tb_AES.Controls.Add(Me.btn_GenerateRandomKey)
        Me.tb_AES.Controls.Add(Me.TextBox1)
        Me.tb_AES.Controls.Add(Me.Label1)
        Me.tb_AES.Controls.Add(Me.txtBox_GeneratedIV)
        Me.tb_AES.Controls.Add(Me.Label2)
        Me.tb_AES.Location = New System.Drawing.Point(4, 22)
        Me.tb_AES.Name = "tb_AES"
        Me.tb_AES.Padding = New System.Windows.Forms.Padding(3)
        Me.tb_AES.Size = New System.Drawing.Size(333, 269)
        Me.tb_AES.TabIndex = 0
        Me.tb_AES.Text = "AES"
        Me.tb_AES.UseVisualStyleBackColor = True
        '
        'lbl_Key
        '
        Me.lbl_Key.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.lbl_Key.AutoSize = True
        Me.lbl_Key.Location = New System.Drawing.Point(6, 52)
        Me.lbl_Key.Name = "lbl_Key"
        Me.lbl_Key.Size = New System.Drawing.Size(45, 13)
        Me.lbl_Key.TabIndex = 8
        Me.lbl_Key.Text = "aes Key"
        '
        'txtBox_GeneratedKey
        '
        Me.txtBox_GeneratedKey.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.txtBox_GeneratedKey.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBox_GeneratedKey.Location = New System.Drawing.Point(9, 68)
        Me.txtBox_GeneratedKey.Multiline = True
        Me.txtBox_GeneratedKey.Name = "txtBox_GeneratedKey"
        Me.txtBox_GeneratedKey.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtBox_GeneratedKey.Size = New System.Drawing.Size(200, 63)
        Me.txtBox_GeneratedKey.TabIndex = 7
        '
        'btn_GenerateRandomKey
        '
        Me.btn_GenerateRandomKey.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.btn_GenerateRandomKey.Location = New System.Drawing.Point(8, 223)
        Me.btn_GenerateRandomKey.Name = "btn_GenerateRandomKey"
        Me.btn_GenerateRandomKey.Size = New System.Drawing.Size(200, 23)
        Me.btn_GenerateRandomKey.TabIndex = 6
        Me.btn_GenerateRandomKey.Text = "Get 32 byte Key and IV from Password"
        Me.btn_GenerateRandomKey.UseVisualStyleBackColor = True
        '
        'TextBox1
        '
        Me.TextBox1.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.TextBox1.Location = New System.Drawing.Point(9, 25)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(200, 20)
        Me.TextBox1.TabIndex = 0
        Me.TextBox1.Text = "password"
        '
        'Label1
        '
        Me.Label1.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(7, 7)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(80, 13)
        Me.Label1.TabIndex = 1
        Me.Label1.Text = "Password (Key)"
        '
        'txtBox_GeneratedIV
        '
        Me.txtBox_GeneratedIV.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.txtBox_GeneratedIV.Font = New System.Drawing.Font("Courier New", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBox_GeneratedIV.Location = New System.Drawing.Point(8, 150)
        Me.txtBox_GeneratedIV.Multiline = True
        Me.txtBox_GeneratedIV.Name = "txtBox_GeneratedIV"
        Me.txtBox_GeneratedIV.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtBox_GeneratedIV.Size = New System.Drawing.Size(201, 51)
        Me.txtBox_GeneratedIV.TabIndex = 2
        '
        'Label2
        '
        Me.Label2.Anchor = System.Windows.Forms.AnchorStyles.Left
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(8, 134)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(103, 13)
        Me.Label2.TabIndex = 5
        Me.Label2.Text = "Generated IV vector"
        '
        'tb_RSA
        '
        Me.tb_RSA.Location = New System.Drawing.Point(4, 22)
        Me.tb_RSA.Name = "tb_RSA"
        Me.tb_RSA.Padding = New System.Windows.Forms.Padding(3)
        Me.tb_RSA.Size = New System.Drawing.Size(300, 269)
        Me.tb_RSA.TabIndex = 1
        Me.tb_RSA.Text = "RSA"
        Me.tb_RSA.UseVisualStyleBackColor = True
        '
        'frm_cryptTools
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(925, 661)
        Me.Controls.Add(Me.tbCtrl_Keys)
        Me.Controls.Add(Me.grpBox_AES)
        Me.Controls.Add(Me.grp_textCipher)
        Me.Controls.Add(Me.btn_getSHA_asVBNET)
        Me.Controls.Add(Me.btn_getFileSHA)
        Me.Controls.Add(Me.grpBox_SHA)
        Me.Controls.Add(Me.lbl_selectedFile)
        Me.Controls.Add(Me.txtBox_FilePath)
        Me.Name = "frm_cryptTools"
        Me.Text = "frm_cryptTools"
        Me.grp_textCipher.ResumeLayout(False)
        Me.grp_textCipher.PerformLayout()
        Me.grpBox_SHA.ResumeLayout(False)
        Me.grpBox_SHA.PerformLayout()
        Me.ContextMenuStrip1.ResumeLayout(False)
        Me.grpBox_AES.ResumeLayout(False)
        Me.grpBox_AES.PerformLayout()
        Me.tbCtrl_Keys.ResumeLayout(False)
        Me.tb_AES.ResumeLayout(False)
        Me.tb_AES.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents grp_textCipher As Windows.Forms.GroupBox
    Friend WithEvents Button2 As Windows.Forms.Button
    Friend WithEvents Button1 As Windows.Forms.Button
    Friend WithEvents txtBox_cipherText As Windows.Forms.TextBox
    Friend WithEvents txtBox_plainText As Windows.Forms.TextBox
    Friend WithEvents btn_getSHA_asVBNET As Windows.Forms.Button
    Friend WithEvents btn_getFileSHA As Windows.Forms.Button
    Friend WithEvents grpBox_SHA As Windows.Forms.GroupBox
    Friend WithEvents rdBtn_SHA512 As Windows.Forms.RadioButton
    Friend WithEvents rdBtn_SHA384 As Windows.Forms.RadioButton
    Friend WithEvents rdBtn_SHA256 As Windows.Forms.RadioButton
    Friend WithEvents txtBox_Hash As Windows.Forms.TextBox
    Friend WithEvents rdBtn_SHA1 As Windows.Forms.RadioButton
    Friend WithEvents lbl_selectedFile As Windows.Forms.Label
    Friend WithEvents txtBox_FilePath As Windows.Forms.TextBox
    Friend WithEvents dlg_getFileName As Windows.Forms.OpenFileDialog
    Friend WithEvents ContextMenuStrip1 As Windows.Forms.ContextMenuStrip
    Friend WithEvents SelectFileToolStripMenuItem As Windows.Forms.ToolStripMenuItem
    Friend WithEvents lbl_AESKey As Windows.Forms.Label
    Friend WithEvents txtBox_AESKey As Windows.Forms.TextBox
    Friend WithEvents txtBox_AESIV As Windows.Forms.TextBox
    Friend WithEvents lbl_AESIV As Windows.Forms.Label
    Friend WithEvents btn_getRandomPassword As Windows.Forms.Button
    Friend WithEvents txtBox_cipheredFilePath As Windows.Forms.TextBox
    Friend WithEvents btn_cipherFile As Windows.Forms.Button
    Friend WithEvents lbl_randomPasswordInfo As Windows.Forms.Label
    Friend WithEvents lbl_encryptTheCurrentFile As Windows.Forms.Label
    Friend WithEvents btn_decryptTheFile As Windows.Forms.Button
    Friend WithEvents grpBox_AES As Windows.Forms.GroupBox
    Friend WithEvents tbCtrl_Keys As Windows.Forms.TabControl
    Friend WithEvents tb_AES As Windows.Forms.TabPage
    Friend WithEvents lbl_Key As Windows.Forms.Label
    Friend WithEvents txtBox_GeneratedKey As Windows.Forms.TextBox
    Friend WithEvents btn_GenerateRandomKey As Windows.Forms.Button
    Friend WithEvents TextBox1 As Windows.Forms.TextBox
    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents txtBox_GeneratedIV As Windows.Forms.TextBox
    Friend WithEvents Label2 As Windows.Forms.Label
    Friend WithEvents tb_RSA As Windows.Forms.TabPage
End Class

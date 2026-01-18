<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm_pictControl3
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
        Me.Label1 = New System.Windows.Forms.Label()
        Me.btn_Cancel = New System.Windows.Forms.Button()
        Me.btn_finishPicture = New System.Windows.Forms.Button()
        Me.GroupBox4 = New System.Windows.Forms.GroupBox()
        Me.txtBox_AspectRatio_SrcImg = New System.Windows.Forms.TextBox()
        Me.txtBox_srcImageHeight = New System.Windows.Forms.TextBox()
        Me.lbl_srcImageWidth = New System.Windows.Forms.Label()
        Me.txtBox_srcImageWidth = New System.Windows.Forms.TextBox()
        Me.lbl_srcImageHeight = New System.Windows.Forms.Label()
        Me.lbl_srcImgHdivW = New System.Windows.Forms.Label()
        Me.GroupBox3 = New System.Windows.Forms.GroupBox()
        Me.txtBox_cropRectHeight = New System.Windows.Forms.TextBox()
        Me.lbl_cropRectWidth = New System.Windows.Forms.Label()
        Me.txtBox_cropRectWidth = New System.Windows.Forms.TextBox()
        Me.lbl_AspectRatio_CropRect = New System.Windows.Forms.Label()
        Me.lbl_cropRect_Height = New System.Windows.Forms.Label()
        Me.txtBox_AspectRatio_CropRect = New System.Windows.Forms.TextBox()
        Me.grpBox_Crop_Values = New System.Windows.Forms.GroupBox()
        Me.txtBox_Delta_CropLeft = New System.Windows.Forms.TextBox()
        Me.lbl_cropBottom = New System.Windows.Forms.Label()
        Me.lbl_Delta_CropLeft = New System.Windows.Forms.Label()
        Me.txBox_DeltaCropTop = New System.Windows.Forms.TextBox()
        Me.txtBox_Delta_CropRight = New System.Windows.Forms.TextBox()
        Me.lbl_Delta_CropRight = New System.Windows.Forms.Label()
        Me.lbl_DeltaCropTop = New System.Windows.Forms.Label()
        Me.txtBox_DeltaCropBottom = New System.Windows.Forms.TextBox()
        Me.GroupBox2 = New System.Windows.Forms.GroupBox()
        Me.lbl_moveRight = New System.Windows.Forms.Label()
        Me.lbl_moveLeft = New System.Windows.Forms.Label()
        Me.lbl_moveUpDown = New System.Windows.Forms.Label()
        Me.scrl_moveLeftRight = New System.Windows.Forms.HScrollBar()
        Me.scrl_moveUpDown = New System.Windows.Forms.VScrollBar()
        Me.GroupBox1 = New System.Windows.Forms.GroupBox()
        Me.lbl_Maximum = New System.Windows.Forms.Label()
        Me.lbl_Minimum = New System.Windows.Forms.Label()
        Me.scrl_reSize = New System.Windows.Forms.VScrollBar()
        Me.GroupBox4.SuspendLayout()
        Me.GroupBox3.SuspendLayout()
        Me.grpBox_Crop_Values.SuspendLayout()
        Me.GroupBox2.SuspendLayout()
        Me.GroupBox1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(12, 317)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(83, 13)
        Me.Label1.TabIndex = 47
        Me.Label1.Text = "frm_pictControl3"
        Me.Label1.Visible = False
        '
        'btn_Cancel
        '
        Me.btn_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.btn_Cancel.Location = New System.Drawing.Point(243, 291)
        Me.btn_Cancel.Name = "btn_Cancel"
        Me.btn_Cancel.Size = New System.Drawing.Size(99, 23)
        Me.btn_Cancel.TabIndex = 46
        Me.btn_Cancel.Text = "Cancel"
        Me.btn_Cancel.UseVisualStyleBackColor = True
        '
        'btn_finishPicture
        '
        Me.btn_finishPicture.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.btn_finishPicture.Location = New System.Drawing.Point(12, 291)
        Me.btn_finishPicture.Name = "btn_finishPicture"
        Me.btn_finishPicture.Size = New System.Drawing.Size(225, 23)
        Me.btn_finishPicture.TabIndex = 45
        Me.btn_finishPicture.Text = "Finish the Picture"
        Me.btn_finishPicture.UseVisualStyleBackColor = True
        '
        'GroupBox4
        '
        Me.GroupBox4.Controls.Add(Me.txtBox_AspectRatio_SrcImg)
        Me.GroupBox4.Controls.Add(Me.txtBox_srcImageHeight)
        Me.GroupBox4.Controls.Add(Me.lbl_srcImageWidth)
        Me.GroupBox4.Controls.Add(Me.txtBox_srcImageWidth)
        Me.GroupBox4.Controls.Add(Me.lbl_srcImageHeight)
        Me.GroupBox4.Controls.Add(Me.lbl_srcImgHdivW)
        Me.GroupBox4.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox4.Location = New System.Drawing.Point(374, 12)
        Me.GroupBox4.Name = "GroupBox4"
        Me.GroupBox4.Size = New System.Drawing.Size(275, 64)
        Me.GroupBox4.TabIndex = 44
        Me.GroupBox4.TabStop = False
        Me.GroupBox4.Text = "Source Image Info (pts)"
        '
        'txtBox_AspectRatio_SrcImg
        '
        Me.txtBox_AspectRatio_SrcImg.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBox_AspectRatio_SrcImg.Location = New System.Drawing.Point(187, 34)
        Me.txtBox_AspectRatio_SrcImg.Name = "txtBox_AspectRatio_SrcImg"
        Me.txtBox_AspectRatio_SrcImg.Size = New System.Drawing.Size(71, 20)
        Me.txtBox_AspectRatio_SrcImg.TabIndex = 27
        Me.txtBox_AspectRatio_SrcImg.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtBox_srcImageHeight
        '
        Me.txtBox_srcImageHeight.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBox_srcImageHeight.Location = New System.Drawing.Point(90, 34)
        Me.txtBox_srcImageHeight.Name = "txtBox_srcImageHeight"
        Me.txtBox_srcImageHeight.Size = New System.Drawing.Size(67, 20)
        Me.txtBox_srcImageHeight.TabIndex = 25
        Me.txtBox_srcImageHeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lbl_srcImageWidth
        '
        Me.lbl_srcImageWidth.AutoSize = True
        Me.lbl_srcImageWidth.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_srcImageWidth.Location = New System.Drawing.Point(8, 19)
        Me.lbl_srcImageWidth.Name = "lbl_srcImageWidth"
        Me.lbl_srcImageWidth.Size = New System.Drawing.Size(69, 13)
        Me.lbl_srcImageWidth.TabIndex = 14
        Me.lbl_srcImageWidth.Text = "Src Image W"
        '
        'txtBox_srcImageWidth
        '
        Me.txtBox_srcImageWidth.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBox_srcImageWidth.Location = New System.Drawing.Point(11, 35)
        Me.txtBox_srcImageWidth.Name = "txtBox_srcImageWidth"
        Me.txtBox_srcImageWidth.Size = New System.Drawing.Size(65, 20)
        Me.txtBox_srcImageWidth.TabIndex = 13
        Me.txtBox_srcImageWidth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lbl_srcImageHeight
        '
        Me.lbl_srcImageHeight.AutoSize = True
        Me.lbl_srcImageHeight.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_srcImageHeight.Location = New System.Drawing.Point(87, 18)
        Me.lbl_srcImageHeight.Name = "lbl_srcImageHeight"
        Me.lbl_srcImageHeight.Size = New System.Drawing.Size(66, 13)
        Me.lbl_srcImageHeight.TabIndex = 26
        Me.lbl_srcImageHeight.Text = "Src Image H"
        '
        'lbl_srcImgHdivW
        '
        Me.lbl_srcImgHdivW.AutoSize = True
        Me.lbl_srcImgHdivW.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_srcImgHdivW.Location = New System.Drawing.Point(184, 19)
        Me.lbl_srcImgHdivW.Name = "lbl_srcImgHdivW"
        Me.lbl_srcImgHdivW.Size = New System.Drawing.Size(50, 13)
        Me.lbl_srcImgHdivW.TabIndex = 28
        Me.lbl_srcImgHdivW.Text = "Src H/W"
        '
        'GroupBox3
        '
        Me.GroupBox3.Controls.Add(Me.txtBox_cropRectHeight)
        Me.GroupBox3.Controls.Add(Me.lbl_cropRectWidth)
        Me.GroupBox3.Controls.Add(Me.txtBox_cropRectWidth)
        Me.GroupBox3.Controls.Add(Me.lbl_AspectRatio_CropRect)
        Me.GroupBox3.Controls.Add(Me.lbl_cropRect_Height)
        Me.GroupBox3.Controls.Add(Me.txtBox_AspectRatio_CropRect)
        Me.GroupBox3.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox3.Location = New System.Drawing.Point(374, 77)
        Me.GroupBox3.Name = "GroupBox3"
        Me.GroupBox3.Size = New System.Drawing.Size(275, 70)
        Me.GroupBox3.TabIndex = 43
        Me.GroupBox3.TabStop = False
        Me.GroupBox3.Text = "Crop Rectangle Info (pts)"
        '
        'txtBox_cropRectHeight
        '
        Me.txtBox_cropRectHeight.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBox_cropRectHeight.Location = New System.Drawing.Point(91, 39)
        Me.txtBox_cropRectHeight.Name = "txtBox_cropRectHeight"
        Me.txtBox_cropRectHeight.Size = New System.Drawing.Size(67, 20)
        Me.txtBox_cropRectHeight.TabIndex = 23
        Me.txtBox_cropRectHeight.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lbl_cropRectWidth
        '
        Me.lbl_cropRectWidth.AutoSize = True
        Me.lbl_cropRectWidth.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_cropRectWidth.Location = New System.Drawing.Point(9, 22)
        Me.lbl_cropRectWidth.Name = "lbl_cropRectWidth"
        Me.lbl_cropRectWidth.Size = New System.Drawing.Size(69, 13)
        Me.lbl_cropRectWidth.TabIndex = 12
        Me.lbl_cropRectWidth.Text = "Crop Rect W"
        '
        'txtBox_cropRectWidth
        '
        Me.txtBox_cropRectWidth.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBox_cropRectWidth.Location = New System.Drawing.Point(12, 38)
        Me.txtBox_cropRectWidth.Name = "txtBox_cropRectWidth"
        Me.txtBox_cropRectWidth.Size = New System.Drawing.Size(67, 20)
        Me.txtBox_cropRectWidth.TabIndex = 11
        Me.txtBox_cropRectWidth.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'lbl_AspectRatio_CropRect
        '
        Me.lbl_AspectRatio_CropRect.AutoSize = True
        Me.lbl_AspectRatio_CropRect.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_AspectRatio_CropRect.Location = New System.Drawing.Point(185, 23)
        Me.lbl_AspectRatio_CropRect.Name = "lbl_AspectRatio_CropRect"
        Me.lbl_AspectRatio_CropRect.Size = New System.Drawing.Size(78, 13)
        Me.lbl_AspectRatio_CropRect.TabIndex = 30
        Me.lbl_AspectRatio_CropRect.Text = "Aspect R H/W"
        '
        'lbl_cropRect_Height
        '
        Me.lbl_cropRect_Height.AutoSize = True
        Me.lbl_cropRect_Height.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_cropRect_Height.Location = New System.Drawing.Point(88, 23)
        Me.lbl_cropRect_Height.Name = "lbl_cropRect_Height"
        Me.lbl_cropRect_Height.Size = New System.Drawing.Size(66, 13)
        Me.lbl_cropRect_Height.TabIndex = 24
        Me.lbl_cropRect_Height.Text = "Crop Rect H"
        '
        'txtBox_AspectRatio_CropRect
        '
        Me.txtBox_AspectRatio_CropRect.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBox_AspectRatio_CropRect.Location = New System.Drawing.Point(188, 39)
        Me.txtBox_AspectRatio_CropRect.Name = "txtBox_AspectRatio_CropRect"
        Me.txtBox_AspectRatio_CropRect.Size = New System.Drawing.Size(71, 20)
        Me.txtBox_AspectRatio_CropRect.TabIndex = 29
        Me.txtBox_AspectRatio_CropRect.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'grpBox_Crop_Values
        '
        Me.grpBox_Crop_Values.Controls.Add(Me.txtBox_Delta_CropLeft)
        Me.grpBox_Crop_Values.Controls.Add(Me.lbl_cropBottom)
        Me.grpBox_Crop_Values.Controls.Add(Me.lbl_Delta_CropLeft)
        Me.grpBox_Crop_Values.Controls.Add(Me.txBox_DeltaCropTop)
        Me.grpBox_Crop_Values.Controls.Add(Me.txtBox_Delta_CropRight)
        Me.grpBox_Crop_Values.Controls.Add(Me.lbl_Delta_CropRight)
        Me.grpBox_Crop_Values.Controls.Add(Me.lbl_DeltaCropTop)
        Me.grpBox_Crop_Values.Controls.Add(Me.txtBox_DeltaCropBottom)
        Me.grpBox_Crop_Values.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.grpBox_Crop_Values.Location = New System.Drawing.Point(374, 153)
        Me.grpBox_Crop_Values.Name = "grpBox_Crop_Values"
        Me.grpBox_Crop_Values.Size = New System.Drawing.Size(275, 159)
        Me.grpBox_Crop_Values.TabIndex = 42
        Me.grpBox_Crop_Values.TabStop = False
        Me.grpBox_Crop_Values.Text = "Crop Settings (pts)"
        '
        'txtBox_Delta_CropLeft
        '
        Me.txtBox_Delta_CropLeft.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBox_Delta_CropLeft.Location = New System.Drawing.Point(35, 78)
        Me.txtBox_Delta_CropLeft.Name = "txtBox_Delta_CropLeft"
        Me.txtBox_Delta_CropLeft.Size = New System.Drawing.Size(68, 20)
        Me.txtBox_Delta_CropLeft.TabIndex = 19
        '
        'lbl_cropBottom
        '
        Me.lbl_cropBottom.AutoSize = True
        Me.lbl_cropBottom.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_cropBottom.Location = New System.Drawing.Point(88, 135)
        Me.lbl_cropBottom.Name = "lbl_cropBottom"
        Me.lbl_cropBottom.Size = New System.Drawing.Size(90, 13)
        Me.lbl_cropBottom.TabIndex = 14
        Me.lbl_cropBottom.Text = "Delta Cop Bottom"
        '
        'lbl_Delta_CropLeft
        '
        Me.lbl_Delta_CropLeft.AutoSize = True
        Me.lbl_Delta_CropLeft.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Delta_CropLeft.Location = New System.Drawing.Point(28, 61)
        Me.lbl_Delta_CropLeft.Name = "lbl_Delta_CropLeft"
        Me.lbl_Delta_CropLeft.Size = New System.Drawing.Size(75, 13)
        Me.lbl_Delta_CropLeft.TabIndex = 20
        Me.lbl_Delta_CropLeft.Text = "Delta CropLeft"
        '
        'txBox_DeltaCropTop
        '
        Me.txBox_DeltaCropTop.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txBox_DeltaCropTop.Location = New System.Drawing.Point(101, 35)
        Me.txBox_DeltaCropTop.Name = "txBox_DeltaCropTop"
        Me.txBox_DeltaCropTop.Size = New System.Drawing.Size(67, 20)
        Me.txBox_DeltaCropTop.TabIndex = 11
        Me.txBox_DeltaCropTop.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'txtBox_Delta_CropRight
        '
        Me.txtBox_Delta_CropRight.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBox_Delta_CropRight.Location = New System.Drawing.Point(169, 78)
        Me.txtBox_Delta_CropRight.Name = "txtBox_Delta_CropRight"
        Me.txtBox_Delta_CropRight.Size = New System.Drawing.Size(68, 20)
        Me.txtBox_Delta_CropRight.TabIndex = 21
        Me.txtBox_Delta_CropRight.TextAlign = System.Windows.Forms.HorizontalAlignment.Right
        '
        'lbl_Delta_CropRight
        '
        Me.lbl_Delta_CropRight.AutoSize = True
        Me.lbl_Delta_CropRight.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Delta_CropRight.Location = New System.Drawing.Point(184, 61)
        Me.lbl_Delta_CropRight.Name = "lbl_Delta_CropRight"
        Me.lbl_Delta_CropRight.Size = New System.Drawing.Size(53, 13)
        Me.lbl_Delta_CropRight.TabIndex = 22
        Me.lbl_Delta_CropRight.Text = "cropRight"
        '
        'lbl_DeltaCropTop
        '
        Me.lbl_DeltaCropTop.AutoSize = True
        Me.lbl_DeltaCropTop.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_DeltaCropTop.Location = New System.Drawing.Point(98, 19)
        Me.lbl_DeltaCropTop.Name = "lbl_DeltaCropTop"
        Me.lbl_DeltaCropTop.Size = New System.Drawing.Size(79, 13)
        Me.lbl_DeltaCropTop.TabIndex = 13
        Me.lbl_DeltaCropTop.Text = "Delta Crop Top"
        '
        'txtBox_DeltaCropBottom
        '
        Me.txtBox_DeltaCropBottom.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.txtBox_DeltaCropBottom.Location = New System.Drawing.Point(101, 112)
        Me.txtBox_DeltaCropBottom.Name = "txtBox_DeltaCropBottom"
        Me.txtBox_DeltaCropBottom.Size = New System.Drawing.Size(67, 20)
        Me.txtBox_DeltaCropBottom.TabIndex = 12
        Me.txtBox_DeltaCropBottom.TextAlign = System.Windows.Forms.HorizontalAlignment.Center
        '
        'GroupBox2
        '
        Me.GroupBox2.Controls.Add(Me.lbl_moveRight)
        Me.GroupBox2.Controls.Add(Me.lbl_moveLeft)
        Me.GroupBox2.Controls.Add(Me.lbl_moveUpDown)
        Me.GroupBox2.Controls.Add(Me.scrl_moveLeftRight)
        Me.GroupBox2.Controls.Add(Me.scrl_moveUpDown)
        Me.GroupBox2.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox2.Location = New System.Drawing.Point(126, 12)
        Me.GroupBox2.Name = "GroupBox2"
        Me.GroupBox2.Size = New System.Drawing.Size(231, 273)
        Me.GroupBox2.TabIndex = 41
        Me.GroupBox2.TabStop = False
        Me.GroupBox2.Text = "2. Move cropping rectangle"
        '
        'lbl_moveRight
        '
        Me.lbl_moveRight.AutoSize = True
        Me.lbl_moveRight.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_moveRight.Location = New System.Drawing.Point(162, 238)
        Me.lbl_moveRight.Name = "lbl_moveRight"
        Me.lbl_moveRight.Size = New System.Drawing.Size(62, 13)
        Me.lbl_moveRight.TabIndex = 4
        Me.lbl_moveRight.Text = "Move Right"
        '
        'lbl_moveLeft
        '
        Me.lbl_moveLeft.AutoSize = True
        Me.lbl_moveLeft.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_moveLeft.Location = New System.Drawing.Point(6, 238)
        Me.lbl_moveLeft.Name = "lbl_moveLeft"
        Me.lbl_moveLeft.Size = New System.Drawing.Size(55, 13)
        Me.lbl_moveLeft.TabIndex = 3
        Me.lbl_moveLeft.Text = "Move Left"
        '
        'lbl_moveUpDown
        '
        Me.lbl_moveUpDown.AutoSize = True
        Me.lbl_moveUpDown.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_moveUpDown.Location = New System.Drawing.Point(85, 25)
        Me.lbl_moveUpDown.Name = "lbl_moveUpDown"
        Me.lbl_moveUpDown.Size = New System.Drawing.Size(60, 13)
        Me.lbl_moveUpDown.TabIndex = 2
        Me.lbl_moveUpDown.Text = "Up / Down"
        '
        'scrl_moveLeftRight
        '
        Me.scrl_moveLeftRight.Location = New System.Drawing.Point(3, 216)
        Me.scrl_moveLeftRight.Maximum = 1000
        Me.scrl_moveLeftRight.Name = "scrl_moveLeftRight"
        Me.scrl_moveLeftRight.Size = New System.Drawing.Size(224, 18)
        Me.scrl_moveLeftRight.TabIndex = 1
        '
        'scrl_moveUpDown
        '
        Me.scrl_moveUpDown.Location = New System.Drawing.Point(105, 46)
        Me.scrl_moveUpDown.Maximum = 1000
        Me.scrl_moveUpDown.Name = "scrl_moveUpDown"
        Me.scrl_moveUpDown.Size = New System.Drawing.Size(18, 157)
        Me.scrl_moveUpDown.TabIndex = 0
        '
        'GroupBox1
        '
        Me.GroupBox1.Controls.Add(Me.lbl_Maximum)
        Me.GroupBox1.Controls.Add(Me.lbl_Minimum)
        Me.GroupBox1.Controls.Add(Me.scrl_reSize)
        Me.GroupBox1.Font = New System.Drawing.Font("Tahoma", 8.25!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.GroupBox1.Location = New System.Drawing.Point(12, 12)
        Me.GroupBox1.Name = "GroupBox1"
        Me.GroupBox1.Size = New System.Drawing.Size(93, 273)
        Me.GroupBox1.TabIndex = 40
        Me.GroupBox1.TabStop = False
        Me.GroupBox1.Text = "1. Resize"
        '
        'lbl_Maximum
        '
        Me.lbl_Maximum.AutoSize = True
        Me.lbl_Maximum.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Maximum.Location = New System.Drawing.Point(15, 238)
        Me.lbl_Maximum.Name = "lbl_Maximum"
        Me.lbl_Maximum.Size = New System.Drawing.Size(51, 13)
        Me.lbl_Maximum.TabIndex = 2
        Me.lbl_Maximum.Text = "Maximum"
        '
        'lbl_Minimum
        '
        Me.lbl_Minimum.AutoSize = True
        Me.lbl_Minimum.Font = New System.Drawing.Font("Microsoft Sans Serif", 8.25!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lbl_Minimum.Location = New System.Drawing.Point(15, 25)
        Me.lbl_Minimum.Name = "lbl_Minimum"
        Me.lbl_Minimum.Size = New System.Drawing.Size(48, 13)
        Me.lbl_Minimum.TabIndex = 1
        Me.lbl_Minimum.Text = "Minimum"
        '
        'scrl_reSize
        '
        Me.scrl_reSize.Location = New System.Drawing.Point(31, 46)
        Me.scrl_reSize.Minimum = 50
        Me.scrl_reSize.Name = "scrl_reSize"
        Me.scrl_reSize.Size = New System.Drawing.Size(17, 188)
        Me.scrl_reSize.TabIndex = 0
        Me.scrl_reSize.Value = 50
        '
        'frm_pictControl3
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.BackColor = System.Drawing.Color.Silver
        Me.ClientSize = New System.Drawing.Size(368, 327)
        Me.ControlBox = False
        Me.Controls.Add(Me.Label1)
        Me.Controls.Add(Me.btn_Cancel)
        Me.Controls.Add(Me.btn_finishPicture)
        Me.Controls.Add(Me.GroupBox4)
        Me.Controls.Add(Me.GroupBox3)
        Me.Controls.Add(Me.grpBox_Crop_Values)
        Me.Controls.Add(Me.GroupBox2)
        Me.Controls.Add(Me.GroupBox1)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.Fixed3D
        Me.Name = "frm_pictControl3"
        Me.Text = "Cropping Control"
        Me.GroupBox4.ResumeLayout(False)
        Me.GroupBox4.PerformLayout()
        Me.GroupBox3.ResumeLayout(False)
        Me.GroupBox3.PerformLayout()
        Me.grpBox_Crop_Values.ResumeLayout(False)
        Me.grpBox_Crop_Values.PerformLayout()
        Me.GroupBox2.ResumeLayout(False)
        Me.GroupBox2.PerformLayout()
        Me.GroupBox1.ResumeLayout(False)
        Me.GroupBox1.PerformLayout()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Windows.Forms.Label
    Friend WithEvents btn_Cancel As Windows.Forms.Button
    Friend WithEvents btn_finishPicture As Windows.Forms.Button
    Friend WithEvents GroupBox4 As Windows.Forms.GroupBox
    Friend WithEvents txtBox_AspectRatio_SrcImg As Windows.Forms.TextBox
    Friend WithEvents txtBox_srcImageHeight As Windows.Forms.TextBox
    Friend WithEvents lbl_srcImageWidth As Windows.Forms.Label
    Friend WithEvents txtBox_srcImageWidth As Windows.Forms.TextBox
    Friend WithEvents lbl_srcImageHeight As Windows.Forms.Label
    Friend WithEvents lbl_srcImgHdivW As Windows.Forms.Label
    Friend WithEvents GroupBox3 As Windows.Forms.GroupBox
    Friend WithEvents txtBox_cropRectHeight As Windows.Forms.TextBox
    Friend WithEvents lbl_cropRectWidth As Windows.Forms.Label
    Friend WithEvents txtBox_cropRectWidth As Windows.Forms.TextBox
    Friend WithEvents lbl_AspectRatio_CropRect As Windows.Forms.Label
    Friend WithEvents lbl_cropRect_Height As Windows.Forms.Label
    Friend WithEvents txtBox_AspectRatio_CropRect As Windows.Forms.TextBox
    Friend WithEvents grpBox_Crop_Values As Windows.Forms.GroupBox
    Friend WithEvents txtBox_Delta_CropLeft As Windows.Forms.TextBox
    Friend WithEvents lbl_cropBottom As Windows.Forms.Label
    Friend WithEvents lbl_Delta_CropLeft As Windows.Forms.Label
    Friend WithEvents txBox_DeltaCropTop As Windows.Forms.TextBox
    Friend WithEvents txtBox_Delta_CropRight As Windows.Forms.TextBox
    Friend WithEvents lbl_Delta_CropRight As Windows.Forms.Label
    Friend WithEvents lbl_DeltaCropTop As Windows.Forms.Label
    Friend WithEvents txtBox_DeltaCropBottom As Windows.Forms.TextBox
    Friend WithEvents GroupBox2 As Windows.Forms.GroupBox
    Friend WithEvents lbl_moveRight As Windows.Forms.Label
    Friend WithEvents lbl_moveLeft As Windows.Forms.Label
    Friend WithEvents lbl_moveUpDown As Windows.Forms.Label
    Friend WithEvents scrl_moveLeftRight As Windows.Forms.HScrollBar
    Friend WithEvents scrl_moveUpDown As Windows.Forms.VScrollBar
    Friend WithEvents GroupBox1 As Windows.Forms.GroupBox
    Friend WithEvents lbl_Maximum As Windows.Forms.Label
    Friend WithEvents lbl_Minimum As Windows.Forms.Label
    Friend WithEvents scrl_reSize As Windows.Forms.VScrollBar
End Class

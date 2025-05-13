<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class fMain
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
        Me.SpritePictureBox = New System.Windows.Forms.PictureBox()
        Me.lblMain = New System.Windows.Forms.Label()
        Me.lblLoadedImages = New System.Windows.Forms.Label()
        Me.SpriteUpDown = New System.Windows.Forms.NumericUpDown()
        Me.lblPictures = New System.Windows.Forms.Label()
        Me.lblFrames = New System.Windows.Forms.Label()
        Me.lblGroups = New System.Windows.Forms.Label()
        Me.lblSprites = New System.Windows.Forms.Label()
        Me.lblImageInfo = New System.Windows.Forms.Label()
        CType(Me.SpritePictureBox, System.ComponentModel.ISupportInitialize).BeginInit()
        CType(Me.SpriteUpDown, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'SpritePictureBox
        '
        Me.SpritePictureBox.Location = New System.Drawing.Point(12, 100)
        Me.SpritePictureBox.Name = "SpritePictureBox"
        Me.SpritePictureBox.Size = New System.Drawing.Size(440, 300)
        Me.SpritePictureBox.TabIndex = 0
        Me.SpritePictureBox.TabStop = False
        '
        'lblMain
        '
        Me.lblMain.AutoSize = True
        Me.lblMain.Location = New System.Drawing.Point(12, 40)
        Me.lblMain.Name = "lblMain"
        Me.lblMain.Size = New System.Drawing.Size(127, 13)
        Me.lblMain.TabIndex = 1
        Me.lblMain.Text = "OP2_ART (.BMP / .PRT)"
        '
        'lblLoadedImages
        '
        Me.lblLoadedImages.AutoSize = True
        Me.lblLoadedImages.Location = New System.Drawing.Point(221, 10)
        Me.lblLoadedImages.Name = "lblLoadedImages"
        Me.lblLoadedImages.Size = New System.Drawing.Size(87, 13)
        Me.lblLoadedImages.TabIndex = 2
        Me.lblLoadedImages.Text = "lblLoadedImages"
        '
        'SpriteUpDown
        '
        Me.SpriteUpDown.Location = New System.Drawing.Point(15, 70)
        Me.SpriteUpDown.Name = "SpriteUpDown"
        Me.SpriteUpDown.Size = New System.Drawing.Size(124, 20)
        Me.SpriteUpDown.TabIndex = 3
        '
        'lblPictures
        '
        Me.lblPictures.AutoSize = True
        Me.lblPictures.Location = New System.Drawing.Point(221, 40)
        Me.lblPictures.Name = "lblPictures"
        Me.lblPictures.Size = New System.Drawing.Size(55, 13)
        Me.lblPictures.TabIndex = 4
        Me.lblPictures.Text = "lblPictures"
        '
        'lblFrames
        '
        Me.lblFrames.AutoSize = True
        Me.lblFrames.Location = New System.Drawing.Point(361, 10)
        Me.lblFrames.Name = "lblFrames"
        Me.lblFrames.Size = New System.Drawing.Size(51, 13)
        Me.lblFrames.TabIndex = 5
        Me.lblFrames.Text = "lblFrames"
        '
        'lblGroups
        '
        Me.lblGroups.AutoSize = True
        Me.lblGroups.Location = New System.Drawing.Point(361, 40)
        Me.lblGroups.Name = "lblGroups"
        Me.lblGroups.Size = New System.Drawing.Size(51, 13)
        Me.lblGroups.TabIndex = 6
        Me.lblGroups.Text = "lblGroups"
        '
        'lblSprites
        '
        Me.lblSprites.AutoSize = True
        Me.lblSprites.Location = New System.Drawing.Point(12, 10)
        Me.lblSprites.Name = "lblSprites"
        Me.lblSprites.Size = New System.Drawing.Size(39, 13)
        Me.lblSprites.TabIndex = 7
        Me.lblSprites.Text = "Sprites"
        '
        'lblImageInfo
        '
        Me.lblImageInfo.AutoSize = True
        Me.lblImageInfo.Location = New System.Drawing.Point(221, 72)
        Me.lblImageInfo.Name = "lblImageInfo"
        Me.lblImageInfo.Size = New System.Drawing.Size(64, 13)
        Me.lblImageInfo.TabIndex = 8
        Me.lblImageInfo.Text = "lblImageInfo"
        '
        'fMain
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(465, 414)
        Me.Controls.Add(Me.lblImageInfo)
        Me.Controls.Add(Me.lblSprites)
        Me.Controls.Add(Me.lblGroups)
        Me.Controls.Add(Me.lblFrames)
        Me.Controls.Add(Me.lblPictures)
        Me.Controls.Add(Me.SpriteUpDown)
        Me.Controls.Add(Me.lblLoadedImages)
        Me.Controls.Add(Me.lblMain)
        Me.Controls.Add(Me.SpritePictureBox)
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle
        Me.MaximizeBox = False
        Me.Name = "fMain"
        Me.Text = "Outpost 2 Graphics"
        CType(Me.SpritePictureBox, System.ComponentModel.ISupportInitialize).EndInit()
        CType(Me.SpriteUpDown, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents SpritePictureBox As PictureBox
    Friend WithEvents lblMain As Label
    Friend WithEvents lblLoadedImages As Label
    Friend WithEvents SpriteUpDown As NumericUpDown
    Friend WithEvents lblPictures As Label
    Friend WithEvents lblFrames As Label
    Friend WithEvents lblGroups As Label
    Friend WithEvents lblSprites As Label
    Friend WithEvents lblImageInfo As Label
End Class

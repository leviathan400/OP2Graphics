Imports System.IO
Imports System.Drawing
Imports Microsoft.VisualBasic.ApplicationServices
Imports System.Security
Imports System.Media
Imports System.Runtime.CompilerServices

' OP2Graphics
' https://github.com/leviathan400/OP2Graphics
'
' Simple project to open and view the images within the OP2_ART.
'
' Replication of Hooman's work; an old VB6 project from 2005: https://forum.outpost2.net/index.php?topic=1593
' Does not look at tilesets
'
'
' Outpost 2: Divided Destiny is a real-time strategy video game released in 1997.

Public Class fMain

    Public ApplicationName As String = "OP2Graphics"
    Public Version As String = "0.5.0"
    Public Build As String = "0020"

#Region "Class Declarations"
    ' Art file paths
    'Private bmpFilePath As String = "D:\op2graphics\OP2_ART.BMP"
    'Private prtFilePath As String = "D:\op2graphics\op2_art.prt"
    Private bmpFilePath As String
    Private prtFilePath As String

    ' Class member for the PRT file
    Private prtFile As New CPrtFile()
#End Region

#Region "Initialization"
    ''' <summary>
    ''' Handles the form load event, initializes the application and loads the OP2 art files
    ''' </summary>
    ''' <param name="sender">The source of the event</param>
    ''' <param name="e">The event data</param>
    Private Sub fMain_Load(sender As Object, e As EventArgs) Handles MyBase.Load

        Debug.WriteLine("--- " & ApplicationName & " Started ---")
        Me.Icon = My.Resources.dynamics
        Me.Text = "Outpost 2 Graphics Viewer"

        lblLoadedImages.Text = "Loaded Images: N/A"
        lblPictures.Text = "Pictures: N/A"
        lblFrames.Text = "Frames: N/A"
        lblGroups.Text = "Groups: N/A"

        If String.IsNullOrEmpty(My.Settings.OP2Path) Then
            'First run.. set OP2Path
            Debug.WriteLine("First Run")
            MessageBox.Show("Browse for the Outpost 2 (1.4.1) folder.", "OP2Graphics", MessageBoxButtons.OK, MessageBoxIcon.Information)
            SetOP2Path()
        End If

        If Not Directory.Exists(My.Settings.OP2Path) Then
            ' If user deleted or moved their OP2 path we need to set it again
            MessageBox.Show("Cant find Outpost 2 folder. Browse for the Outpost 2 (1.4.1) folder.", "OP2Graphics", MessageBoxButtons.OK, MessageBoxIcon.Information)
            SetOP2Path()
        End If

        ' Update file paths based on the selected OP2 path
        UpdateFilePaths()

        LoadOP2ART()

        Dim player As New SoundPlayer(My.Resources.dump)
        player.Play()

        lblImageInfo.Visible = False

        'Type 4 Shadow Sprite
        'SpriteUpDown.Value = 5045
        'prtFile.DebugImageInfo(5045)

        'Type 5 Shadow Sprite
        'SpriteUpDown.Value = 4590
        'prtFile.DebugImageInfo(4590)
    End Sub
#End Region

#Region "File Management"
    ''' <summary>
    ''' Opens a folder browser dialog to let the user select the Outpost 2 game directory
    ''' </summary>
    Private Sub SetOP2Path()
        Using folderDialog As New FolderBrowserDialog()
            folderDialog.Description = "Select Outpost 2 Game Directory"
            folderDialog.ShowNewFolderButton = False

            ' Set initial directory with proper validation
            If Not String.IsNullOrWhiteSpace(My.Settings.OP2Path) AndAlso Directory.Exists(My.Settings.OP2Path) Then
                folderDialog.SelectedPath = My.Settings.OP2Path
            Else
                ' Default to Program Files
                folderDialog.SelectedPath = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)
            End If

            If folderDialog.ShowDialog() = DialogResult.OK Then
                Dim selectedPath As String = folderDialog.SelectedPath

                ' Validate the Outpost 2 directory
                If Not File.Exists(Path.Combine(selectedPath, "outpost2.exe")) Then
                    MessageBox.Show("Could not find outpost2.exe in the selected folder. Please select a valid Outpost 2 installation folder.", "Invalid Folder", MessageBoxButtons.OK, MessageBoxIcon.Error)
                    ' Call SetOP2Path recursively to force user to select a valid folder
                    SetOP2Path()
                    Return
                End If

                ' Set the OP2 path in settings
                My.Settings.OP2Path = selectedPath
                My.Settings.Save()

                ' Update file paths based on the selected OP2 path
                'UpdateFilePaths()
            End If
        End Using
    End Sub

    ''' <summary>
    ''' Updates the file paths for the BMP and PRT files based on the OP2 path in settings
    ''' </summary>
    Private Sub UpdateFilePaths()
        ' Set the paths to the BMP and PRT files - We are expecting a 1.4.1 folder
        bmpFilePath = Path.Combine(My.Settings.OP2Path, "OP2_ART.BMP")
        prtFilePath = Path.Combine(My.Settings.OP2Path, "OPU", "base", "sprites", "op2_art.prt")

        'Debug.WriteLine("BMP file path set to: " & bmpFilePath)
        'Debug.WriteLine("PRT file path set to: " & prtFilePath)
    End Sub

    ''' <summary>
    ''' Loads the OP2_ART.BMP and op2_art.prt files, initializes the UI with the loaded data
    ''' </summary>
    ''' <returns>True if files loaded successfully, False otherwise</returns>
    Private Function LoadOP2ART() As Boolean
        'Debug.WriteLine("Starting to load OP2_ART files...")

        ' Try to open the BMP file
        If prtFile.OpenBitmapFile(bmpFilePath) <> 0 Then
            Debug.WriteLine("Error loading BMP file from: " & bmpFilePath)
            MessageBox.Show("Could not load OP2_ART.BMP file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If
        'Debug.WriteLine("Successfully loaded BMP file")

        ' Try to open the PRT file
        If prtFile.OpenPrtFile(prtFilePath) <> 0 Then
            Debug.WriteLine("Error loading PRT file from: " & prtFilePath)
            MessageBox.Show("Could not load op2_art.prt file.", "Error", MessageBoxButtons.OK, MessageBoxIcon.Error)
            Return False
        End If
        'Debug.WriteLine("Successfully loaded PRT file")

        Debug.WriteLine("Loaded OP2_ART files.")

        ' Update the UI with count of loaded images
        Dim numImages As Integer = prtFile.NumImages
        lblLoadedImages.Text = "Loaded Images: " & numImages.ToString()
        'Debug.WriteLine("Total number of images: " & numImages)

        ' Configure the numeric up/down control
        SpriteUpDown.Minimum = 0
        SpriteUpDown.Maximum = numImages - 1
        SpriteUpDown.Value = 0  ' Start with the first image

        ' Display the first image
        DisplaySprite(0)

        ' Update the UI with all sprite information
        UpdateSpriteInfoLabels()

        Return True
    End Function
#End Region

#Region "UI Interaction"
    ''' <summary>
    ''' Updates the sprite information labels with current counts from the PRT file
    ''' </summary>
    Private Sub UpdateSpriteInfoLabels()
        ' Log detailed sprite stats to the debug output
        'prtFile.LogSpriteStats()

        ' Update the labels in the UI
        lblLoadedImages.Text = $"Loaded Images: {prtFile.NumImages}"
        lblPictures.Text = $"Pictures: {prtFile.NumPictures}"
        lblFrames.Text = $"Frames: {prtFile.NumFrames}"
        lblGroups.Text = $"Groups: {prtFile.NumGroups}"

    End Sub

    ''' <summary>
    ''' Handles the value changed event for the sprite numeric up-down control
    ''' </summary>
    ''' <param name="sender">The source of the event</param>
    ''' <param name="e">The event data</param>
    Private Sub SpriteUpDown_ValueChanged(sender As Object, e As EventArgs) Handles SpriteUpDown.ValueChanged
        DisplaySprite(CInt(SpriteUpDown.Value))
    End Sub

    ''' <summary>
    ''' Handles the text changed event for the sprite numeric up-down control
    ''' </summary>
    ''' <param name="sender">The source of the event</param>
    ''' <param name="e">The event data</param>
    Private Sub SpriteUpDown_TextChanged(sender As Object, e As EventArgs) Handles SpriteUpDown.TextChanged
        ' Try to parse the current text
        Dim spriteIndex As Integer
        If Integer.TryParse(SpriteUpDown.Text, spriteIndex) Then
            ' Check if the parsed value is in valid range
            If spriteIndex >= 0 AndAlso spriteIndex < prtFile.NumImages Then
                ' Display the sprite
                DisplaySprite(spriteIndex)
            End If
        End If
    End Sub
#End Region

#Region "Image Display"
    ''' <summary>
    ''' Displays the sprite with the specified index in the PictureBox
    ''' </summary>
    ''' <param name="imageIndex">The index of the image to display</param>
    Private Sub DisplaySprite(imageIndex As Integer)
        If imageIndex < 0 OrElse imageIndex >= prtFile.NumImages Then
            Debug.WriteLine($"DisplaySprite: Invalid image index: {imageIndex}")
            Return
        End If

        'Debug.WriteLine($"DisplaySprite: Showing sprite index: {imageIndex}")

        ' Get image dimensions
        Dim imgWidth As Integer = prtFile.ImageWidth(imageIndex)
        Dim imgHeight As Integer = prtFile.ImageHeight(imageIndex)
        Dim imgType As Integer = prtFile.ImageType(imageIndex)
        Dim imgPalette As Integer = prtFile.ImagePalette(imageIndex)

        ' Update info labels
        lblImageInfo.Text = $"Image: {imageIndex}, Size: {imgWidth}x{imgHeight}, Type: {imgType}, Palette: {imgPalette}"

        ' Clean up previous image
        If SpritePictureBox.Image IsNot Nothing Then
            SpritePictureBox.Image.Dispose()
            SpritePictureBox.Image = Nothing
        End If

        ' Create a bitmap for the picture box with light gray background
        Dim bmp As New Bitmap(SpritePictureBox.Width, SpritePictureBox.Height)

        Using g As Graphics = Graphics.FromImage(bmp)
            ' Clear with light gray background
            g.Clear(System.Drawing.Color.FromArgb(224, 224, 224))

            ' Get the sprite bitmap
            Dim spriteBmp As Bitmap = prtFile.GetImage(imageIndex)

            If spriteBmp IsNot Nothing Then
                ' Calculate position to center the sprite
                Dim x As Integer = (SpritePictureBox.Width - imgWidth) \ 2
                Dim y As Integer = (SpritePictureBox.Height - imgHeight) \ 2

                ' Draw the sprite with nearest neighbor interpolation (for pixel-perfect display)
                g.InterpolationMode = Drawing2D.InterpolationMode.NearestNeighbor
                g.PixelOffsetMode = Drawing2D.PixelOffsetMode.Half

                ' For transparency handling
                If imgType = 4 OrElse imgType = 5 Then
                    ' For shadow types, use alpha blending
                    g.CompositingMode = Drawing2D.CompositingMode.SourceOver
                    g.CompositingQuality = Drawing2D.CompositingQuality.HighQuality
                End If

                ' Draw the sprite
                g.DrawImageUnscaled(spriteBmp, x, y)

                ' Clean up
                spriteBmp.Dispose()
            Else
                g.DrawString("Error loading sprite", New Font("Arial", 10, FontStyle.Bold),
                     New SolidBrush(System.Drawing.Color.Red), 10, 50)
            End If
        End Using

        ' Update the picture box
        SpritePictureBox.Image = bmp
    End Sub
#End Region

End Class
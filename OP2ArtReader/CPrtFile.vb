Imports System.IO
Imports System.Drawing
Imports System.Drawing.Imaging
Imports System.Runtime.InteropServices
Imports OP2Graphics.CPalette

''' <summary>
''' Class for loading and accessing Outpost 2 graphics (op2_art.prt and op2_art.bmp)
''' </summary>
Public Class CPrtFile

#Region "Win32 API Structures and Constants"
    ' Win32 API structures and constants
    Public Const DIB_RGB_COLORS As Integer = 0
    Public Const BI_RGB As Integer = 0

    <StructLayout(LayoutKind.Sequential)>
    Public Structure BITMAPINFOHEADER
        Public biSize As Integer
        Public biWidth As Integer
        Public biHeight As Integer
        Public biPlanes As Short
        Public biBitCount As Short
        Public biCompression As Integer
        Public biSizeImage As Integer
        Public biXPelsPerMeter As Integer
        Public biYPelsPerMeter As Integer
        Public biClrUsed As Integer
        Public biClrImportant As Integer
    End Structure

    <StructLayout(LayoutKind.Sequential)>
    Public Structure RGBQUAD
        Public rgbBlue As Byte
        Public rgbGreen As Byte
        Public rgbRed As Byte
        Public rgbReserved As Byte
    End Structure

    ' Import the Win32 API function
    <DllImport("gdi32.dll")>
    Public Shared Function SetDIBitsToDevice(
        hdc As IntPtr,
        xDest As Integer,
        yDest As Integer,
        dwWidth As Integer,
        dwHeight As Integer,
        xSrc As Integer,
        ySrc As Integer,
        uStartScan As Integer,
        cScanLines As Integer,
        lpvBits As IntPtr,
        lpbmi As IntPtr,
        fuColorUse As Integer) As Integer
    End Function
#End Region

#Region "Private Fields"
    Private DebugMode As Boolean = False

    ' PRT data
    Private numLoadedPalettes As Integer           ' Total number of palettes
    Private palette() As CPalette                  ' Array of all palettes
    Private numLoadedImages As Integer             ' Total number of images (5,390)
    Private images() As ImageInfo                  ' Array of image info structs
    Private numLoadedGroups As Integer             ' Number of groups (2,079) - called "animations" in code
    Private numLoadedFrames As Integer             ' Total number of frames (24,185)
    Private numLoadedPictures As Integer           ' Total number of pictures (160,922) - called "frameComponents" in code
    Private numLoadedFrameOptional As Integer      ' Additional info attached to select frames
    Private groups() As GroupInfo                  ' Array describing all groups - called "animations" in code
    Private frames() As FrameInfo                  ' Array describing all frames
    Private pictures() As PictureInfo              ' Array describing all pictures - called "frameComponents" in code

    ' Bitmap data
    Private bitmapData() As Byte                   ' Raw pixel data from bmp file
#End Region

#Region "Data Structures"
    ' Structures for the image port of the .prt file (second section)
    Private Structure ImageInfo
        Public ScanLineByteWidth As Integer   ' Pitch
        Public DataPtr As Integer             ' Pointer to pixel data in bmp file
        Public Height As Integer              ' Image height
        Public Width As Integer               ' Image width
        Public Type As Short                  ' Image type (menu, game, shadow, etc.)
        Public PaletteNum As Short            ' Palette index to use for this image
    End Structure

    ' The field "Type" has the following values/meanings
    ' 0 - menu graphic
    ' 1 - in game graphic
    ' 2
    ' 3
    ' 4 - Unit shadows - (1 bpp)
    ' 5 - Unit shadows - (1 bpp)
    ' ... more types

    ' Structures for the Animation section of the .prt file (third section)
    Private Structure Rect
        Public Left As Integer
        Public Top As Integer
        Public Right As Integer
        Public Bottom As Integer
    End Structure

    ' Group header (called "Animation" in the code)
    Private Structure GroupHeaderInfo
        Public Unknown As Integer             ' Set to 1 **TODO** Find out use
        Public SelectionBox As Rect           ' Selection rectangle for the group
        Public PixelXDisplacement As Integer  ' X offset for positioning
        Public PixelYDisplacement As Integer  ' Y offset for positioning
        Public Unknown2 As Integer            ' Set to 0 (usually) **TODO** Find out use
        Public NumFrames As Integer           ' Number of frames in this group
    End Structure

    Private Structure GroupAppendixInfo       ' Extra data for groups
        Public Unknown1 As Integer
        Public Unknown2 As Integer
        Public Unknown3 As Integer
        Public Unknown4 As Integer
    End Structure

    ' Frame header info
    Private Structure FrameHeaderInfo
        Public NumPictures As Byte            ' Number of pictures in this frame (with high bit as flag)
        Public Unknown As Byte                ' Unknown with high bit as flag
    End Structure

    ' Optional frame data
    Private Structure FrameOptionalInfo
        Public Unknown1 As Byte
        Public Unknown2 As Byte
        Public Unknown3 As Byte
        Public Unknown4 As Byte
    End Structure

    ' Picture information (called "FrameComponent" in the code)
    Private Structure PictureInfo
        Public ImageIndex As Short            ' Index in the image() array for this picture
        Public Unused As Byte                 ' Set to 0xFF (Unused/Ignored by Outpost2)
        Public PictureOrder As Byte           ' Order to draw pictures within a frame
        Public PixelXOffset As Short          ' X offset within the frame
        Public PixelYOffset As Short          ' Y offset within the frame
    End Structure

    ' Group info (called "Animation" in the code)
    Private Structure GroupInfo
        Public Header As GroupHeaderInfo      ' Group header information
        Public StartingFrameIndex As Integer  ' Index of first frame in the group
        Public NumAppendices As Integer       ' Number of appendix entries
        Public StartingAppendixIndex As Integer ' Index of first appendix
    End Structure

    ' Frame info
    Private Structure FrameInfo
        Public Header As FrameHeaderInfo      ' Frame header information
        Public Optional_ As FrameOptionalInfo ' Optional frame data
        Public StartingPictureIndex As Integer ' Index of first picture in the frame
    End Structure
#End Region

#Region "File Loading"
    ''' <summary>
    ''' Opens and loads the OP2_ART.BMP file
    ''' </summary>
    ''' <param name="strFileName">Path to the BMP file</param>
    ''' <returns>0 on success, -1 on failure</returns>
    Public Function OpenBitmapFile(strFileName As String) As Integer
        Try
            ' Check if the file exists
            If Not File.Exists(strFileName) Then
                Debug.WriteLine($"BMP file not found: {strFileName}")
                Return -1
            End If

            ' Read the bitmap file
            Using fs As New FileStream(strFileName, FileMode.Open, FileAccess.Read)
                Using reader As New BinaryReader(fs)
                    Return ReadBitmapData(reader)
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine("Error opening bitmap file: " & ex.Message)
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Opens and loads the OP2_ART.PRT file
    ''' </summary>
    ''' <param name="strFileName">Path to the PRT file</param>
    ''' <returns>0 on success, -1 on failure</returns>
    Public Function OpenPrtFile(strFileName As String) As Integer
        Try
            ' Check if the file exists
            If Not File.Exists(strFileName) Then
                Debug.WriteLine($"PRT file not found: {strFileName}")
                Return -1
            End If

            ' Open the file for reading
            Using fs As New FileStream(strFileName, FileMode.Open, FileAccess.Read)
                Using reader As New BinaryReader(fs)
                    Return ReadPrtData(reader)
                End Using
            End Using
        Catch ex As Exception
            Debug.WriteLine("Error opening PRT file: " & ex.Message)
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Reads the PRT file data
    ''' </summary>
    Public Function ReadPrtData(reader As BinaryReader) As Integer
        Try
            Dim i As Integer, j As Integer, k As Integer
            Dim strTag As String

            ' Read in the palette section
            ' ***************************
            If DebugMode = True Then
                Debug.WriteLine("Reading palette section...")
            End If

            ' Read the palette section header tag
            Dim tagBytes(3) As Byte
            reader.Read(tagBytes, 0, 4)
            strTag = System.Text.Encoding.ASCII.GetString(tagBytes)

            If strTag <> "CPAL" Then
                Debug.WriteLine("Invalid PRT file: Missing CPAL tag")
                Return -1    ' Return failure
            End If

            ' Read the number of palettes
            numLoadedPalettes = reader.ReadInt32()
            If DebugMode = True Then
                Debug.WriteLine($"Found {numLoadedPalettes} palettes")
            End If

            ' Allocate the palette array
            ReDim palette(numLoadedPalettes - 1)

            ' Read in all the palette info
            For i = 0 To numLoadedPalettes - 1
                ' Create a new palette object
                palette(i) = New CPalette()

                ' Read in the palette data
                If palette(i).ReadPaletteData(reader) <> 0 Then
                    ' Error reading in palette data. Abort.
                    Debug.WriteLine("CPrtFile: Error reading in palette #" & i)
                    Return -1
                End If
            Next

            ' Read in the images info
            ' ***********************

            ' Read in the number of image info structs
            numLoadedImages = reader.ReadInt32()

            If DebugMode = True Then
                Debug.WriteLine("Reading image section...")
                Debug.WriteLine($"Found {numLoadedImages} images")
            End If

            ' Allocate array to hold image structs
            ReDim images(numLoadedImages - 1)

            ' Read in the image info structs
            For i = 0 To numLoadedImages - 1
                images(i).ScanLineByteWidth = reader.ReadInt32()
                images(i).DataPtr = reader.ReadInt32()
                images(i).Height = reader.ReadInt32()
                images(i).Width = reader.ReadInt32()
                images(i).Type = reader.ReadInt16()
                images(i).PaletteNum = reader.ReadInt16()
            Next

            ' Read in the Animation info
            ' *************************

            ' Read in the structure counts
            numLoadedGroups = reader.ReadInt32()          ' Groups (animations)
            numLoadedFrames = reader.ReadInt32()          ' Frames
            numLoadedPictures = reader.ReadInt32()        ' Pictures (frame components)
            numLoadedFrameOptional = reader.ReadInt32()   ' Optional frame data

            If DebugMode = True Then
                Debug.WriteLine("Reading animation data section...")
                Debug.WriteLine($"Found {numLoadedGroups} groups, {numLoadedFrames} frames, {numLoadedPictures} pictures")
            End If

            ' Allocate space for the data
            ReDim groups(numLoadedGroups - 1)
            ReDim frames(numLoadedFrames - 1)
            ReDim pictures(numLoadedPictures - 1)

            Dim currentFrameIndex As Integer = 0
            Dim currentPictureIndex As Integer = 0

            ' Read in the Group data (animations)
            For i = 0 To numLoadedGroups - 1
                ' Read the group header (animation header)
                groups(i).Header.Unknown = reader.ReadInt32()
                groups(i).Header.SelectionBox.Left = reader.ReadInt32()
                groups(i).Header.SelectionBox.Top = reader.ReadInt32()
                groups(i).Header.SelectionBox.Right = reader.ReadInt32()
                groups(i).Header.SelectionBox.Bottom = reader.ReadInt32()
                groups(i).Header.PixelXDisplacement = reader.ReadInt32()
                groups(i).Header.PixelYDisplacement = reader.ReadInt32()
                groups(i).Header.Unknown2 = reader.ReadInt32()
                groups(i).Header.NumFrames = reader.ReadInt32()

                ' Set the starting frame index
                groups(i).StartingFrameIndex = currentFrameIndex

                ' Read in all frame info for this group
                For j = 0 To groups(i).Header.NumFrames - 1
                    ' Read in frame header
                    frames(currentFrameIndex).Header.NumPictures = reader.ReadByte()
                    frames(currentFrameIndex).Header.Unknown = reader.ReadByte()

                    ' Check if first block of optional data needs to be read in
                    If (frames(currentFrameIndex).Header.NumPictures And &H80) = &H80 Then
                        ' Read in the optional data
                        frames(currentFrameIndex).Optional_.Unknown1 = reader.ReadByte()
                        frames(currentFrameIndex).Optional_.Unknown2 = reader.ReadByte()
                    End If

                    ' Check if second block of optional data needs to be read in
                    If (frames(currentFrameIndex).Header.Unknown And &H80) = &H80 Then
                        ' Read in the optional data
                        frames(currentFrameIndex).Optional_.Unknown3 = reader.ReadByte()
                        frames(currentFrameIndex).Optional_.Unknown4 = reader.ReadByte()
                    End If

                    ' Set the starting picture index for this frame
                    frames(currentFrameIndex).StartingPictureIndex = currentPictureIndex

                    ' Calculate actual number of pictures by clearing the high bit
                    Dim numPicsInFrame As Integer = frames(currentFrameIndex).Header.NumPictures And &H7F

                    ' Read all pictures for this frame
                    For k = 0 To numPicsInFrame - 1
                        ' Read the picture data (frame component)
                        pictures(currentPictureIndex).ImageIndex = reader.ReadInt16()
                        pictures(currentPictureIndex).Unused = reader.ReadByte()
                        pictures(currentPictureIndex).PictureOrder = reader.ReadByte()
                        pictures(currentPictureIndex).PixelXOffset = reader.ReadInt16()
                        pictures(currentPictureIndex).PixelYOffset = reader.ReadInt16()

                        currentPictureIndex += 1
                    Next k

                    currentFrameIndex += 1
                Next j

                ' Read number of appendices
                groups(i).NumAppendices = reader.ReadInt32()

                ' Read in all appendix data
                For j = 0 To groups(i).NumAppendices - 1
                    ' Read in the appendix structure (skipping for now)
                    reader.ReadInt32()  ' Unknown1
                    reader.ReadInt32()  ' Unknown2 
                    reader.ReadInt32()  ' Unknown3
                    reader.ReadInt32()  ' Unknown4
                Next j
            Next i

            ' Validate counts
            If numLoadedFrames <> currentFrameIndex Then
                Debug.WriteLine($"Warning: Frame count mismatch. Expected {numLoadedFrames}, got {currentFrameIndex}")
            End If

            If numLoadedPictures <> currentPictureIndex Then
                Debug.WriteLine($"Warning: Picture count mismatch. Expected {numLoadedPictures}, got {currentPictureIndex}")
            End If

            If DebugMode = True Then
                Debug.WriteLine("Successfully loaded PRT data")
            End If

            Return 0

        Catch ex As Exception
            Debug.WriteLine("Error reading PRT data: " & ex.Message)
            Debug.WriteLine(ex.StackTrace)
            Return -1
        End Try
    End Function

    ''' <summary>
    ''' Reads the bitmap data from a BMP file
    ''' </summary>
    Public Function ReadBitmapData(reader As BinaryReader) As Integer
        Try
            'Debug.WriteLine("Reading BMP data...")

            ' Read the bitmap file header
            Dim bfType As UInt16 = reader.ReadUInt16()          ' Should be 'BM' (0x4D42)
            Dim bfSize As UInt32 = reader.ReadUInt32()          ' File size
            Dim bfReserved1 As UInt16 = reader.ReadUInt16()     ' Reserved (0)
            Dim bfReserved2 As UInt16 = reader.ReadUInt16()     ' Reserved (0)
            Dim bfOffBits As UInt32 = reader.ReadUInt32()       ' Offset to pixel data

            ' Read BITMAPINFOHEADER
            Dim biSize As UInt32 = reader.ReadUInt32()          ' Size of this header
            Dim biWidth As Int32 = reader.ReadInt32()           ' Width
            Dim biHeight As Int32 = reader.ReadInt32()          ' Height
            Dim biPlanes As UInt16 = reader.ReadUInt16()        ' Planes (1)
            Dim biBitCount As UInt16 = reader.ReadUInt16()      ' Bits per pixel
            Dim biCompression As UInt32 = reader.ReadUInt32()   ' Compression type
            Dim biSizeImage As UInt32 = reader.ReadUInt32()     ' Image size
            Dim biXPelsPerMeter As Int32 = reader.ReadInt32()   ' X resolution
            Dim biYPelsPerMeter As Int32 = reader.ReadInt32()   ' Y resolution
            Dim biClrUsed As UInt32 = reader.ReadUInt32()       ' Colors used
            Dim biClrImportant As UInt32 = reader.ReadUInt32()  ' Important colors

            If DebugMode = True Then
                Debug.WriteLine($"Bitmap header: Type={bfType:X4}, Size={bfSize}, OffBits={bfOffBits}")
                Debug.WriteLine($"Bitmap info: {biWidth}x{biHeight}, {biBitCount} bpp, compression={biCompression}")
            End If

            ' Skip the color table if present (for OP2_ART.BMP it's zeros)
            Dim colorTableSize As Integer = CInt(bfOffBits - 14 - 40)
            If colorTableSize > 0 Then
                reader.BaseStream.Position = 14 + 40 + colorTableSize

                If DebugMode = True Then
                    Debug.WriteLine($"Skipped color table of {colorTableSize} bytes")
                End If
            End If

            ' Read all bitmap data
            Dim dataLength As Integer = CInt(reader.BaseStream.Length - reader.BaseStream.Position)
            ReDim bitmapData(dataLength - 1)
            bitmapData = reader.ReadBytes(dataLength)

            If DebugMode = True Then
                Debug.WriteLine($"Read {dataLength} bytes of bitmap data")
            End If

            Return 0
        Catch ex As Exception
            Debug.WriteLine($"Error reading bitmap data: {ex.Message}")
            Return -1
        End Try
    End Function
#End Region

#Region "Data Access Properties"
    ''' <summary>
    ''' Returns the total number of images in the file (lowest level)
    ''' </summary>
    Public ReadOnly Property NumImages() As Integer
        Get
            Return numLoadedImages
        End Get
    End Property

    ''' <summary>
    ''' Returns the total number of pictures (positioned images) in the file
    ''' </summary>
    Public ReadOnly Property NumPictures As Integer
        Get
            ' Pictures are stored in the frameComponents array
            Return numLoadedPictures
        End Get
    End Property

    ''' <summary>
    ''' Returns the total number of frames across all groups
    ''' </summary>
    Public ReadOnly Property NumFrames As Integer
        Get
            Return numLoadedFrames
        End Get
    End Property

    ''' <summary>
    ''' Returns the number of groups (animation sequences)
    ''' </summary>
    Public ReadOnly Property NumGroups As Integer
        Get
            Return numLoadedGroups
        End Get
    End Property

    ''' <summary>
    ''' Gets the width of an image
    ''' </summary>
    Public ReadOnly Property ImageWidth(imageNum As Integer) As Integer
        Get
            If imageNum >= 0 AndAlso imageNum < numLoadedImages Then
                Return images(imageNum).Width
            End If
            Return 0
        End Get
    End Property

    ''' <summary>
    ''' Gets the height of an image
    ''' </summary>
    Public ReadOnly Property ImageHeight(imageNum As Integer) As Integer
        Get
            If imageNum >= 0 AndAlso imageNum < numLoadedImages Then
                Return images(imageNum).Height
            End If
            Return 0
        End Get
    End Property

    ''' <summary>
    ''' Gets the type of an image
    ''' </summary>
    Public ReadOnly Property ImageType(imageNum As Integer) As Integer
        Get
            If imageNum >= 0 AndAlso imageNum < numLoadedImages Then
                Return images(imageNum).Type
            End If
            Return 0
        End Get
    End Property

    ''' <summary>
    ''' Gets the palette number for an image
    ''' </summary>
    Public ReadOnly Property ImagePalette(imageNum As Integer) As Integer
        Get
            If imageNum >= 0 AndAlso imageNum < numLoadedImages Then
                Return images(imageNum).PaletteNum
            End If
            Return 0
        End Get
    End Property

    ''' <summary>
    ''' Gets number of frames in a specific group
    ''' </summary>
    Public ReadOnly Property NumFramesInGroup(groupIndex As Integer) As Integer
        Get
            If groupIndex >= 0 AndAlso groupIndex < numLoadedGroups Then
                Return groups(groupIndex).Header.NumFrames
            End If
            Return 0
        End Get
    End Property

    ''' <summary>
    ''' Gets number of pictures in a specific frame
    ''' </summary>
    Public ReadOnly Property NumPicturesInFrame(frameIndex As Integer) As Integer
        Get
            If frameIndex >= 0 AndAlso frameIndex < numLoadedFrames Then
                Return frames(frameIndex).Header.NumPictures And &H7F
            End If
            Return 0
        End Get
    End Property
#End Region

#Region "Navigation Methods"
    ''' <summary>
    ''' Gets the starting frame index for a group
    ''' </summary>
    Public Function GetGroupStartFrame(groupIndex As Integer) As Integer
        If groupIndex >= 0 AndAlso groupIndex < numLoadedGroups Then
            Return groups(groupIndex).StartingFrameIndex
        End If
        Return -1
    End Function

    ''' <summary>
    ''' Gets the frame index for a specific frame within a group
    ''' </summary>
    Public Function GetFrameIndexInGroup(groupIndex As Integer, frameOffset As Integer) As Integer
        If groupIndex >= 0 AndAlso groupIndex < numLoadedGroups AndAlso
           frameOffset >= 0 AndAlso frameOffset < groups(groupIndex).Header.NumFrames Then
            Return groups(groupIndex).StartingFrameIndex + frameOffset
        End If
        Return -1
    End Function

    ''' <summary>
    ''' Gets the starting picture index for a frame
    ''' </summary>
    Public Function GetFrameStartPicture(frameIndex As Integer) As Integer
        If frameIndex >= 0 AndAlso frameIndex < numLoadedFrames Then
            Return frames(frameIndex).StartingPictureIndex
        End If
        Return -1
    End Function

    ''' <summary>
    ''' Gets the picture index for a specific picture within a frame
    ''' </summary>
    Public Function GetPictureIndexInFrame(frameIndex As Integer, pictureOffset As Integer) As Integer
        If frameIndex >= 0 AndAlso frameIndex < numLoadedFrames AndAlso
           pictureOffset >= 0 AndAlso pictureOffset < (frames(frameIndex).Header.NumPictures And &H7F) Then
            Return frames(frameIndex).StartingPictureIndex + pictureOffset
        End If
        Return -1
    End Function

    ''' <summary>
    ''' Gets the image index for a picture
    ''' </summary>
    Public Function GetImageIndexForPicture(pictureIndex As Integer) As Integer
        If pictureIndex >= 0 AndAlso pictureIndex < numLoadedPictures Then
            Return pictures(pictureIndex).ImageIndex
        End If
        Return -1
    End Function
#End Region

#Region "Image Rendering"
    ''' <summary>
    ''' Creates a bitmap for the specified image with proper palette and pixel data
    ''' </summary>
    Public Function GetImage(imageNum As Integer) As Bitmap
        Try
            If imageNum < 0 OrElse imageNum >= numLoadedImages Then
                Debug.WriteLine($"GetImage: Invalid image number: {imageNum}")
                Return Nothing
            End If

            ' Get image information
            Dim width As Integer = images(imageNum).Width
            Dim height As Integer = images(imageNum).Height
            Dim dataOffset As Integer = images(imageNum).DataPtr
            Dim scanLineWidth As Integer = images(imageNum).ScanLineByteWidth
            Dim imageType As Integer = images(imageNum).Type
            Dim paletteNum As Integer = images(imageNum).PaletteNum

            'Debug.WriteLine($"GetImage: Image {imageNum}: {width}x{height}, Type: {imageType}, Palette: {paletteNum}")

            ' Create a compatible bitmap that we can render to
            Dim resultBitmap As New Bitmap(width, height, System.Drawing.Imaging.PixelFormat.Format32bppArgb)
            Dim g As Graphics = Graphics.FromImage(resultBitmap)

            ' Clear with transparent background
            g.Clear(System.Drawing.Color.Transparent)

            ' Create a GDI+ DC for rendering
            Dim hdc As IntPtr = g.GetHdc()

            Try
                ' Determine bit depth based on image type
                Dim bitDepth As Short = If(imageType = 4 OrElse imageType = 5, 1, 8)

                ' Calculate size of BITMAPINFOHEADER + color table
                Dim colorTableSize As Integer = If(bitDepth = 1, 2, 256) * Marshal.SizeOf(GetType(WinGDI.RGBQUAD))
                Dim infoHeaderSize As Integer = Marshal.SizeOf(GetType(WinGDI.BITMAPINFOHEADER))
                Dim totalSize As Integer = infoHeaderSize + colorTableSize

                ' Allocate memory for BITMAPINFO structure (header + color table)
                Dim bmiPtr As IntPtr = Marshal.AllocHGlobal(totalSize)
                Dim bitmapDataPtr As IntPtr = IntPtr.Zero

                Try
                    ' Initialize memory to zeros
                    Dim buffer(totalSize - 1) As Byte
                    Marshal.Copy(buffer, 0, bmiPtr, totalSize)

                    ' Set up the BITMAPINFOHEADER
                    Dim biHeader As New WinGDI.BITMAPINFOHEADER With {
                        .biSize = infoHeaderSize,
                        .biWidth = width,
                        .biHeight = -height,  ' Negative for top-down DIB
                        .biPlanes = 1,
                        .biBitCount = bitDepth,
                        .biCompression = WinGDI.BI_RGB,
                        .biSizeImage = 0,
                        .biXPelsPerMeter = 0,
                        .biYPelsPerMeter = 0,
                        .biClrUsed = 0,
                        .biClrImportant = 0
                    }

                    ' Copy the header to the allocated memory
                    Marshal.StructureToPtr(biHeader, bmiPtr, False)

                    ' Set up the color table
                    Dim colorTablePtr As IntPtr = IntPtr.Add(bmiPtr, infoHeaderSize)

                    If bitDepth = 1 Then
                        ' For shadow sprites (1bpp), we need a 2-color palette
                        Dim shadowPal(1) As WinGDI.RGBQUAD

                        ' Different shadow colors based on image type
                        If imageType = 4 Then
                            ' Type 4: Blue shadow on transparent background
                            shadowPal(0) = New WinGDI.RGBQUAD With {
                                .rgbBlue = 0,
                                .rgbGreen = 0,
                                .rgbRed = 0,
                                .rgbReserved = 0  ' Fully transparent
                            }
                            shadowPal(1) = New WinGDI.RGBQUAD With {
                                .rgbBlue = 255,
                                .rgbGreen = 0,
                                .rgbRed = 0,
                                .rgbReserved = 0  ' Fully opaque blue
                            }
                            'Debug.WriteLine("Using blue shadow color for Type 4")

                        ElseIf imageType = 5 Then

                            ' Type 5: Blue shadow on magenta background
                            shadowPal(0) = New WinGDI.RGBQUAD With {
                                .rgbBlue = 255,
                                .rgbGreen = 0,
                                .rgbRed = 255,  ' Magenta background
                                .rgbReserved = 255  ' Fully opaque
                            }
                            shadowPal(1) = New WinGDI.RGBQUAD With {
                                .rgbBlue = 255,
                                .rgbGreen = 0,
                                .rgbRed = 0,
                                .rgbReserved = 255  ' Blue shadow
                            }
                            'Debug.WriteLine("Using blue shadow on magenta background for Type 5")

                        End If

                        ' Copy the colors to unmanaged memory
                        Dim shadowPalSize As Integer = 2 * Marshal.SizeOf(GetType(WinGDI.RGBQUAD))
                        Dim shadowPalBytes(shadowPalSize - 1) As Byte
                        Dim gch As GCHandle = GCHandle.Alloc(shadowPal, GCHandleType.Pinned)
                        Try
                            Marshal.Copy(gch.AddrOfPinnedObject(), shadowPalBytes, 0, shadowPalSize)
                        Finally
                            gch.Free()
                        End Try
                        Marshal.Copy(shadowPalBytes, 0, colorTablePtr, shadowPalSize)
                    Else
                        ' For normal sprites (8bpp), get the palette from our palette object
                        Dim palColors(255) As WinGDI.RGBQUAD
                        palette(paletteNum).GetPaletteData(palColors)

                        ' Copy all 256 colors to unmanaged memory
                        Dim palSize As Integer = 256 * Marshal.SizeOf(GetType(WinGDI.RGBQUAD))
                        Dim palBytes(palSize - 1) As Byte
                        Dim gch As GCHandle = GCHandle.Alloc(palColors, GCHandleType.Pinned)
                        Try
                            Marshal.Copy(gch.AddrOfPinnedObject(), palBytes, 0, palSize)
                        Finally
                            gch.Free()
                        End Try
                        Marshal.Copy(palBytes, 0, colorTablePtr, palSize)
                    End If

                    ' Allocate memory for bitmap data and copy the data
                    Dim dataSize As Integer = height * scanLineWidth
                    bitmapDataPtr = Marshal.AllocHGlobal(dataSize)
                    Marshal.Copy(bitmapData, dataOffset, bitmapDataPtr, dataSize)

                    ' Call the SetDIBitsToDevice function
                    Dim result As Integer = WinGDI.SetDIBitsToDevice(
                        hdc, 0, 0, width, height, 0, 0, 0, height,
                        bitmapDataPtr, bmiPtr, WinGDI.DIB_RGB_COLORS)

                    If result = 0 Then
                        Debug.WriteLine("SetDIBitsToDevice failed!")
                    End If

                Finally
                    ' Free allocated memory
                    If bitmapDataPtr <> IntPtr.Zero Then
                        Marshal.FreeHGlobal(bitmapDataPtr)
                    End If

                    Marshal.FreeHGlobal(bmiPtr)
                End Try

            Finally
                ' Release the DC
                g.ReleaseHdc(hdc)
                g.Dispose()
            End Try

            Return resultBitmap

        Catch ex As Exception
            Debug.WriteLine($"GetImage: Error creating bitmap: {ex.Message}")
            Debug.WriteLine($"GetImage: Stack trace: {ex.StackTrace}")
            Return Nothing
        End Try
    End Function

    ''' <summary>
    ''' Draws a frame composed of multiple pictures
    ''' </summary>
    Public Function DrawFrame(g As Graphics, destX As Integer, destY As Integer, frameIndex As Integer) As Boolean
        Try
            If frameIndex < 0 OrElse frameIndex >= numLoadedFrames Then
                Debug.WriteLine($"DrawFrame: Invalid frame index: {frameIndex}")
                Return False
            End If

            ' Get number of pictures in this frame (mask off the high bit)
            Dim numPics As Integer = frames(frameIndex).Header.NumPictures And &H7F
            Dim startPic As Integer = frames(frameIndex).StartingPictureIndex

            ' Draw each picture in the frame
            For i As Integer = 0 To numPics - 1
                Dim pictureIndex As Integer = startPic + i
                Dim imageIndex As Integer = pictures(pictureIndex).ImageIndex
                Dim xOffset As Integer = pictures(pictureIndex).PixelXOffset
                Dim yOffset As Integer = pictures(pictureIndex).PixelYOffset

                ' Draw the image at the specified offset
                Dim bmp As Bitmap = GetImage(imageIndex)
                If bmp IsNot Nothing Then
                    g.DrawImage(bmp, destX + xOffset, destY + yOffset)
                    bmp.Dispose()
                End If
            Next

            Return True
        Catch ex As Exception
            Debug.WriteLine($"DrawFrame: Error: {ex.Message}")
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Draw an image to a destination Graphics object
    ''' </summary>
    Public Function DrawImage(g As Graphics, destX As Integer, destY As Integer, imageNum As Integer) As Boolean
        Try
            Dim bmp As Bitmap = GetImage(imageNum)
            If bmp Is Nothing Then
                Return False
            End If

            g.DrawImage(bmp, destX, destY)
            bmp.Dispose()

            Return True
        Catch ex As Exception
            Debug.WriteLine("Error drawing image: " & ex.Message)
            Return False
        End Try
    End Function

    ''' <summary>
    ''' Draw an image directly to a form or control
    ''' </summary>
    Public Function Paste(destDC As IntPtr, destX As Integer, destY As Integer, imageNum As Integer) As Boolean
        Try
            Dim bmp As Bitmap = GetImage(imageNum)
            If bmp Is Nothing Then
                Return False
            End If

            ' Create a Graphics object from the DC
            Using g As Graphics = Graphics.FromHdc(destDC)
                g.DrawImage(bmp, destX, destY)
            End Using

            bmp.Dispose()
            Return True
        Catch ex As Exception
            Debug.WriteLine("Error pasting image: " & ex.Message)
            Return False
        End Try
    End Function
#End Region

#Region "Diagnostics"
    ''' <summary>
    ''' Generates a summary report of the sprite information
    ''' </summary>
    Public Sub LogSpriteStats()
        Debug.WriteLine("---------- OP2_ART Sprite Statistics ----------")
        Debug.WriteLine($"Total Images: {NumImages}")
        Debug.WriteLine($"Total Pictures: {NumPictures}")
        Debug.WriteLine($"Total Frames: {NumFrames}")
        Debug.WriteLine($"Total Groups: {NumGroups}")

        ' Log details about unique image types
        Dim typeCount As New Dictionary(Of Integer, Integer)
        For i As Integer = 0 To numLoadedImages - 1
            Dim imgType As Integer = images(i).Type
            If typeCount.ContainsKey(imgType) Then
                typeCount(imgType) += 1
            Else
                typeCount.Add(imgType, 1)
            End If
        Next

        Debug.WriteLine("Image Types Distribution:")
        For Each kvp As KeyValuePair(Of Integer, Integer) In typeCount
            Debug.WriteLine($"  Type {kvp.Key}: {kvp.Value} images")
        Next

        ' Log details about palettes
        Dim paletteCount As New Dictionary(Of Integer, Integer)
        For i As Integer = 0 To numLoadedImages - 1
            Dim palIdx As Integer = images(i).PaletteNum
            If paletteCount.ContainsKey(palIdx) Then
                paletteCount(palIdx) += 1
            Else
                paletteCount.Add(palIdx, 1)
            End If
        Next

        Debug.WriteLine("Palette Usage:")
        For Each kvp As KeyValuePair(Of Integer, Integer) In paletteCount
            Debug.WriteLine($"  Palette {kvp.Key}: used by {kvp.Value} images")
        Next

        Debug.WriteLine("Group Info:")
        For i As Integer = 0 To Math.Min(20, numLoadedGroups - 1)  ' First 20 groups
            Debug.WriteLine($"  Group {i}: {groups(i).Header.NumFrames} frames")
        Next

        Debug.WriteLine("-----------------------------------------------")
    End Sub

    ''' <summary>
    ''' Outputs detailed debug information about a specific image
    ''' </summary>
    Public Sub DebugImageInfo(imageNum As Integer)
        If imageNum < 0 OrElse imageNum >= numLoadedImages Then
            Debug.WriteLine($"DebugImageInfo: Invalid image number: {imageNum}")
            Return
        End If

        Dim img As ImageInfo = images(imageNum)

        Debug.WriteLine("============================================")
        Debug.WriteLine($"DETAILED DEBUG FOR IMAGE #{imageNum}")
        Debug.WriteLine("============================================")
        Debug.WriteLine($"Dimensions: {img.Width}x{img.Height}")
        Debug.WriteLine($"Type: {img.Type}")
        Debug.WriteLine($"Palette: {img.PaletteNum}")
        Debug.WriteLine($"DataPtr: {img.DataPtr}")
        Debug.WriteLine($"ScanLineByteWidth: {img.ScanLineByteWidth}")

        ' Calculate some useful values
        Dim bytesNeeded As Integer
        If img.Type = 4 OrElse img.Type = 5 Then
            bytesNeeded = (img.Width + 7) \ 8 * img.Height  ' 1bpp
            Debug.WriteLine($"Bit depth: 1bpp (shadow sprite)")
        Else
            bytesNeeded = img.Width * img.Height  ' 8bpp
            Debug.WriteLine($"Bit depth: 8bpp")
        End If

        Debug.WriteLine($"Calculated bytes needed: {bytesNeeded}")
        Debug.WriteLine($"Actual bytes allocated: {img.ScanLineByteWidth * img.Height}")

        ' Debug data pointer validity
        If img.DataPtr < 0 OrElse img.DataPtr >= bitmapData.Length Then
            Debug.WriteLine($"ERROR: DataPtr {img.DataPtr} is outside valid range (0-{bitmapData.Length - 1})")
        Else
            Debug.WriteLine($"DataPtr is valid within bitmap data range")

            ' Show first few bytes of image data
            Dim dataStr As String = "First bytes: "
            For i As Integer = 0 To Math.Min(15, img.ScanLineByteWidth - 1)
                If img.DataPtr + i < bitmapData.Length Then
                    dataStr += $"{bitmapData(img.DataPtr + i):X2} "
                Else
                    dataStr += "?? "
                End If
            Next
            Debug.WriteLine(dataStr)
        End If

        ' If it's a shadow sprite, also dump the first few bytes in binary
        If img.Type = 4 OrElse img.Type = 5 Then
            Debug.WriteLine("Binary representation of first few bytes:")
            For i As Integer = 0 To Math.Min(3, img.ScanLineByteWidth - 1)
                If img.DataPtr + i < bitmapData.Length Then
                    Dim b As Byte = bitmapData(img.DataPtr + i)
                    Debug.WriteLine($"Byte {i}: {Convert.ToString(b, 2).PadLeft(8, "0"c)}")
                End If
            Next
        End If

        Debug.WriteLine("============================================")
    End Sub
#End Region

End Class
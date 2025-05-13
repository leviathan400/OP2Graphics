Imports System.IO
Imports System.Text

''' <summary>
''' Class for handling Outpost 2 palettes used in the PRT file
''' </summary>
Public Class CPalette

#Region "Structures"
    ''' <summary>
    ''' Represents a color in the palette with RGBA components
    ''' </summary>
    Public Structure RGBQUAD
        Public rgbBlue As Byte
        Public rgbGreen As Byte
        Public rgbRed As Byte
        Public rgbReserved As Byte
    End Structure

    ''' <summary>
    ''' Internal structure for section headers in the palette data
    ''' </summary>
    Private Structure SectionHeader
        Public tag As String    ' 4 character tag
        Public sectionSize As Integer
    End Structure
#End Region

#Region "Private Fields"
    ''' <summary>
    ''' The palette data (256 colors)
    ''' </summary>
    Private palData(255) As RGBQUAD
#End Region

#Region "Data Access Methods"
    ''' <summary>
    ''' Reads palette data from a binary file using a BinaryReader
    ''' </summary>
    ''' <param name="reader">BinaryReader positioned at the start of palette data</param>
    ''' <returns>0 on success, -1 on error</returns>
    Public Function ReadPaletteData(reader As BinaryReader) As Integer
        Dim numTagsLeft As Integer
        Dim sectHead As New SectionHeader()
        Dim i As Integer
        Dim temp As Byte

        ' Initialize tag count
        numTagsLeft = -2

        ' Read all needed tags
        Do
            ' Decrease the tag count
            numTagsLeft = numTagsLeft - 1

            ' Read in the section header
            Dim tagBytes(3) As Byte
            reader.Read(tagBytes, 0, 4)
            sectHead.tag = Encoding.ASCII.GetString(tagBytes)
            sectHead.sectionSize = reader.ReadInt32()

            ' Determine which type of section it was
            Select Case sectHead.tag
                Case "PPAL"
                    ' Error check the section size
                    If sectHead.sectionSize <> 1048 Then
                        Debug.WriteLine("CPalette: Warning! 'PPAL' section size is not 1048")
                    End If
                Case "RIFF"
                    Debug.WriteLine("CPalette: Warning! 'RIFF' tag unhandled")
                Case "head"
                    ' Error check the section size
                    If sectHead.sectionSize <> 4 Then
                        Debug.WriteLine("CPalette: Warning! 'head' section size is not 4")
                    End If
                    ' Read numTagsLeft (4 bytes of data)
                    numTagsLeft = reader.ReadInt32()
                Case "data"
                    ' Error check the section size
                    If sectHead.sectionSize <> 1024 Then
                        Debug.WriteLine("CPalette: Warning! 'data' section size is not 1024")
                    End If

                    ' Read the palette data - read each RGBQUAD structure
                    For i = 0 To 255
                        palData(i).rgbBlue = reader.ReadByte()
                        palData(i).rgbGreen = reader.ReadByte()
                        palData(i).rgbRed = reader.ReadByte()
                        palData(i).rgbReserved = reader.ReadByte()
                    Next
                Case "pspl"
                    Debug.WriteLine("CPalette: Warning! 'pspl' tag unhandled")
                Case "ptpl"
                    Debug.WriteLine("CPalette: Warning! 'ptpl' tag unhandled")
                Case Else
                    Debug.WriteLine("CPalette: Warning! Unrecognized section tag: " & sectHead.tag)
                    Return -1    ' Error
            End Select
        Loop While numTagsLeft <> 0

        ' Reverse the color components of the palette data
        ' (For some reason, the custom Outpost 2 format stores
        ' it backwards from the standard Windows format)
        For i = 0 To 255
            temp = palData(i).rgbBlue
            palData(i).rgbBlue = palData(i).rgbRed
            palData(i).rgbRed = temp
        Next

        ' Return success
        Return 0
    End Function

    ''' <summary>
    ''' Copies palette data to the destination array
    ''' </summary>
    ''' <param name="paletteData">Destination array for palette data</param>
    Friend Sub GetPaletteData(ByRef paletteData() As WinGDI.RGBQUAD)
        For i As Integer = 0 To 255
            paletteData(i) = New WinGDI.RGBQUAD() With {
                .rgbBlue = palData(i).rgbBlue,
                .rgbGreen = palData(i).rgbGreen,
                .rgbRed = palData(i).rgbRed,
                .rgbReserved = palData(i).rgbReserved
            }
        Next
    End Sub

    ''' <summary>
    ''' Sets an entry in the palette
    ''' </summary>
    ''' <param name="index">Index of the palette entry to modify</param>
    ''' <param name="palEntry">New palette entry value</param>
    Friend Sub SetPaletteData(index As Integer, palEntry As RGBQUAD)
        If index > 0 AndAlso index <= 255 Then
            palData(index) = palEntry
        End If
    End Sub
#End Region

#Region "Conversion Methods"
    ''' <summary>
    ''' Creates a .NET Color array from the palette
    ''' </summary>
    ''' <returns>Array of 256 Colors</returns>
    Public Function GetDotNetColors() As System.Drawing.Color()
        Dim colors(255) As System.Drawing.Color

        Debug.WriteLine("GetDotNetColors: Converting palette to .NET colors")

        For i As Integer = 0 To 255
            ' Use FromArgb with full alpha channel (255)
            ' Note: The palette swap from BGRA to RGBA should already be handled in ReadPaletteData
            colors(i) = System.Drawing.Color.FromArgb(255,
                                                     palData(i).rgbRed,
                                                     palData(i).rgbGreen,
                                                     palData(i).rgbBlue)

            ' Debug first few palette entries
            If i < 5 Then
                Debug.WriteLine($"Palette[{i}] = R:{palData(i).rgbRed}, G:{palData(i).rgbGreen}, B:{palData(i).rgbBlue}")
            End If
        Next

        Return colors
    End Function
#End Region

End Class
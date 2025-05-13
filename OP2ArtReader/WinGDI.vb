Imports System.Runtime.InteropServices

Public Module WinGDI

    ' Win32 API constants
    Public Const DIB_RGB_COLORS As Integer = 0
    Public Const BI_RGB As Integer = 0

    <StructLayout(LayoutKind.Sequential)>
    Public Structure RGBQUAD
        Public rgbBlue As Byte
        Public rgbGreen As Byte
        Public rgbRed As Byte
        Public rgbReserved As Byte
    End Structure

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

    ' Import the Win32 API function
    <DllImport("gdi32.dll")>
    Public Function SetDIBitsToDevice(
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

End Module
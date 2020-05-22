Attribute VB_Name = "modPictures"
Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(7) As Byte
End Type

Private Type PicBmp
    Size As Long
    Type As Long
    hBmp As Long
    hPal As Long
    Reserved As Long
End Type

Public Enum eColor
    Red
    Green
    Blue
End Enum

Private Type BITMAPINFOHEADER '40 bytes
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Type BITMAP
    bmType As Long
    bmWidth As Long
    bmHeight As Long
    bmWidthBytes As Long
    bmPlanes As Integer
    bmBitsPixel As Integer
    bmBits As Long
End Type

Const ScrCopy = &HCC0020

Private Type BITMAPINFO
    bmiHeader As BITMAPINFOHEADER
    bmiColors As RGBQUAD
End Type

Private Const DIB_RGB_COLORS = 0&
Private Const BI_RGB = 0&

Private Declare Function GetVersionExA Lib "kernel32" _
               (lpVersionInformation As OSVERSIONINFO) As Integer

Private Type OSVERSIONINFO
       dwOSVersionInfoSize As Long
       dwMajorVersion As Long
       dwMinorVersion As Long
       dwBuildNumber As Long
       dwPlatformId As Long
       szCSDVersion As String * 128
End Type



Private Declare Function GetDC Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetObject Lib "gdi32" Alias "GetObjectA" (ByVal hObject As Long, ByVal nCount As Long, lpObject As Any) As Long
Private Declare Function GetBitmapBits Lib "gdi32" (ByVal hBitmap As Long, ByVal dwCount As Long, lpBits As Any) As Long
Declare Function StretchBlt Lib "gdi32" (ByVal hdc As Long, ByVal X As Long, ByVal Y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Declare Function GetSysColor Lib "user32" (ByVal nIndex As Long) As Long
Private Declare Function GetDIBits Lib "gdi32" (ByVal aHDC As Long, ByVal hBitmap As Long, ByVal nStartScan As Long, ByVal nNumScans As Long, lpBits As Any, lpBI As BITMAPINFO, ByVal wUsage As Long) As Long
Public Declare Function CreateCompatibleDC Lib "gdi32" (ByVal hdc As Long) As Long
Public Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Public Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long

Private Declare Function compress Lib "zlib.dll" (dest As Any, destLen As Any, src As Any, ByVal srcLen As Long) As Integer

Dim PicBits()       As Byte
Dim PicInfo         As BITMAP
Dim Cnt             As Long
Dim BytesPerLine    As Long

Public Function PictureWidth(hbitMaps As Long) As Long

    GetObject hbitMaps, Len(PicInfo), PicInfo
    PictureWidth = PicInfo.bmWidth

End Function

Public Function PictureHeight(hbitMaps As Long) As Long

    GetObject hbitMaps, Len(PicInfo), PicInfo
    PictureHeight = PicInfo.bmHeight

End Function

Public Function PDFAdaptor(hbitMaps As Long, Optional GrayScale As Boolean = False, Optional Zipped As Boolean = False) As Byte()



    Dim lGray     As Long
    Dim Red       As Byte
    Dim Green     As Byte
    
    Dim strWidht  As String * 5
    Dim strHeight As String * 5
    Dim intX      As Long
    Dim intY      As Long
    
    Dim BitsPdf()       As Byte
    Dim BitsZip()       As Byte
    Dim ZipSize         As Long
    
    Dim bmInfo As BITMAPINFO
    
    Dim hdcNew As Long
    
        
    GetObject hbitMaps, Len(PicInfo), PicInfo
    BytesPerLine = (PicInfo.bmWidth * 3 + 3) And &HFFFFFFFC
    ReDim PicBits(1 To (BytesPerLine * PicInfo.bmHeight * 3)) As Byte
    ReDim BitsPdf(1 To (BytesPerLine * PicInfo.bmHeight / 4 * 9)) As Byte
    
    If getVersion = 1 Then
        'reallocate storage space
        'En el resto guardamos los Bytes de la imagen
        GetBitmapBits hbitMaps, UBound(PicBits), PicBits(1)
    Else
        With bmInfo.bmiHeader
            .biSize = 40
            .biWidth = PicInfo.bmWidth
            ' Use negative height to scan top-down.
            .biHeight = -PicInfo.bmHeight
            .biPlanes = 1
            .biBitCount = 32
            .biCompression = BI_RGB
            .biSizeImage = BytesPerLine * PicInfo.bmHeight * 3
        End With
        'Creo un DC y le asigno el mapa de bits
        hdcNew = CreateCompatibleDC(0&)
        SelectObject hdcNew, hbitMaps
        GetDIBits hdcNew, hbitMaps, 0, PicInfo.bmHeight, PicBits(1), bmInfo, DIB_RGB_COLORS
        DeleteDC hdcNew
    End If
    


    intY = 1
    For intX = 1 To UBound(PicBits) Step 4
        If Not GrayScale Then
            BitsPdf(intY) = PicBits(intX + 2) 'blue
            BitsPdf(intY + 1) = PicBits(intX + 1) 'green
            BitsPdf(intY + 2) = PicBits(intX) 'red
        Else
            lGray = (222 * CLng(PicBits(intX)) + 707 * CLng(PicBits(intX + 1)) + 71 * CLng(PicBits(intX + 2))) / 1000
            BitsPdf(intY) = lGray 'blue
            BitsPdf(intY + 1) = lGray 'green
            BitsPdf(intY + 2) = lGray 'red
        End If
        intY = intY + 3
    Next

    If Not Zipped Then
        PDFAdaptor = BitsPdf
    Else
        ZipSize = UBound(BitsPdf) * 1.01 + 12
        ReDim BitsZip(1 To ZipSize)
        compress BitsZip(1), ZipSize, BitsPdf(1), UBound(BitsPdf)
        ReDim Preserve BitsZip(1 To ZipSize)
        PDFAdaptor = BitsZip
    End If
    

End Function

Public Function LongToRGB(Color As Long, Componente As eColor) As Integer

    Select Case Componente
    Case Red
        LongToRGB = Color Mod &H100
    Case Green
        LongToRGB = (Color \ &H100) Mod &H100
    Case Blue
        LongToRGB = (Color \ &H10000) Mod &H100
    End Select

End Function

Private Function getVersion() As Integer
   Dim osinfo As OSVERSIONINFO
   Dim retvalue As Integer

   osinfo.dwOSVersionInfoSize = 148
   osinfo.szCSDVersion = Space$(128)
   retvalue = GetVersionExA(osinfo)
   With osinfo
     Select Case .dwPlatformId
        Case 1
            getVersion = 1
        Case 2
            getVersion = 2
        Case Else
             getVersion = 0
        End Select
     End With
End Function





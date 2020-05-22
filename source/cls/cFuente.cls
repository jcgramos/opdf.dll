VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cFont"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public Name     As String
Public BaseFont As String
Public Encoding As eEncoding
Public pdfObj   As Integer
Public fType    As eType
Public fStyle As eStyle

'Apis
Private Declare Function GetTextMetrics Lib "gdi32" Alias "GetTextMetricsA" _
                        (ByVal hdc As Long, lpMetrics As TextMetric) As Long

Private Declare Function GetCharWidth32 Lib "gdi32" Alias "GetCharWidth32A" (ByVal hdc As Long, ByVal iFirstChar As Long, ByVal iLastChar As Long, lpBuffer As Long) As Long

' Estructura TEXTMETRIC
Private Type TextMetric
    tmHeight As Long
    tmAscent As Long
    tmDescent As Long
    tmInternalLeading As Long
    tmExternalLeading As Long
    tmAveCharWidth As Long
    tmMaxCharWidth As Long
    tmWeight As Long
    tmOverhang As Long
    tmDigitizedAspectX As Long
    tmDigitizedAspectY As Long
    tmFirstChar As Byte
    tmLastChar As Byte
    tmDefaultChar As Byte
    tmBreakChar As Byte
    tmItalic As Byte
    tmUnderlined As Byte
    tmStruckOut As Byte
    tmPitchAndFamily As Byte
    tmCharSet As Byte
End Type

Public Enum eStyle
 Normal
 Bold
 Italic
 BoldItalic
End Enum


Public Enum AlignV
    vTop
    vBottom
    vMiddle
End Enum

Public Enum AlignH
    hLeft
    hRight
    hCenter
    hjustify
End Enum

'Codificacion
Enum eEncoding
    MacRomanEncoding
    MacExpertEncoding
    WinAnsiEncoding
End Enum
'Tipo de fuente
Enum eType
    TrueType
    Type1
End Enum

'14 Fuentes estandard (Actualmente esta enumeración no se usa)
Enum eBaseFont
    Times_Roman
    Helvetica
    Courier
    Symbol
    Times_Bold
    Helvetica_Bold
    Courier_Bold
    ZapfDingbats
    Times_Italic
    Helvetica_Oblique
    Courier_Oblique
    Times_BoldItalic
    Helvetica_BoldOblique
    Courier_BoldOblique
End Enum
Private Vararray(1 To 255) 'La defino aqui porque se usa para calcular el ancho del texto en varias funciones
Private OldFont
Friend Sub WFont()
    If pdfObj <> 0 Then
        Exit Sub
    End If
Select Case fType
    Case TrueType
        Select Case fStyle
        Case Normal
           LoadFontTTF Name, BaseFont
        Case Italic
           LoadFontTTF Name, BaseFont, ",Italic"
        Case Bold
           LoadFontTTF Name, BaseFont, ",Bold"
        Case BoldItalic
           LoadFontTTF Name, BaseFont, ",BoldItalic"
        End Select
    Case Type1
    pdfObj = NextObj
    Objetos.Add pdfObj, glngOffSet
    
    outText pdfObj & " 0 obj", True
    outText "<<", True
    outText "/Type /Font", True
    outText "/Subtype /Type1", True
    outText "/Name /" & Name, True
    outText "/BaseFont /" & BaseFont, True
    outText ">>", True
    outText "endobj", True
End Select
End Sub

Private Sub Class_Initialize()
'    BaseFont = "Times New Roman"
    Encoding = WinAnsiEncoding
End Sub

Friend Function ImprimeTexto(Top As Integer, Left As Integer, Height As Integer, Width As Integer, strTexto As String, FontSize As Integer, vAlign As AlignV, hAlign As AlignH, FontName As String)

    Dim intX        As Integer
    Dim intStart    As Integer
    Dim intPalabras As Integer
    Dim intLineas   As Integer
    Dim WordSpacing As Double
    Dim LineWidth   As Integer
    Dim LineCount   As Integer
    Dim tLeft       As Integer
    Dim tTop        As Integer
    Dim AjusteV     As Integer
    Dim sPalabra    As String
    
    
    'Si hay alguna palabra cuyo ancho sea mayor que el ancho, la subdividimos en varias palabras
Reiteracion:
    intPalabras = CuentaPalabras(strTexto)
    For intX = 0 To intPalabras - 1
        sPalabra = DevuelvePalabras(strTexto, intX, intX)
        If TextWidth(sPalabra, FontSize, 1) > Width Then
            Do While TextWidth(sPalabra, FontSize, 1) > Width
                sPalabra = Mid(sPalabra, 1, Len(sPalabra) - 1)
            Loop
            strTexto = DevuelvePalabras(strTexto, 0, intX - 1) & " " & sPalabra & " " & Mid(DevuelvePalabras(strTexto, intX, intX), Len(sPalabra) + 1) & DevuelvePalabras(strTexto, intX + 1, intPalabras - 1)
           GoTo Reiteracion
        End If
    Next
            
            
    If OldFont <> FontName Then
        intLineas = CuentaLineas(strTexto, FontSize, Width)
    Else
        intLineas = CuentaLineas2(strTexto, FontSize, Width)
    End If
    OldFont = FontName
    
    intStart = 0
    LineCount = 0
    Top = gintPageHeight - Top
    
    Do While intStart < intPalabras
        For intX = intStart To intPalabras - 1
            LineWidth = TextWidth(DevuelvePalabras(strTexto, intStart, intX), FontSize, 0)
            If LineWidth > Width Then
                Exit For
            End If
        Next
        
        'Probando suerte
        LineCount = LineCount + 1
        LineWidth = TextWidth(DevuelvePalabras(strTexto, intStart, intX - 1), FontSize, 0)
        WordSpacing = 0
        
        'Ajuste horizontal
        tLeft = Left
        Select Case hAlign
        Case hjustify
            If intX <= intPalabras - 1 Then
                If intX - intStart > 1 Then
                    Do While TextWidth(DevuelvePalabras(strTexto, intStart, intX - 1), FontSize, WordSpacing) < Width
                        WordSpacing = WordSpacing + 0.1
                    Loop
                    WordSpacing = WordSpacing - 0.1
                End If
            End If
        Case hRight
            tLeft = Left + (Width - LineWidth)
        Case hCenter
            tLeft = Left + (Width - LineWidth) / 2
        End Select
        
        'Ajuste Vertical
        Select Case vAlign
        Case vMiddle
            AjusteV = (Height - (intLineas * FontSize)) * -0.5
        Case vBottom
            AjusteV = (Height - (intLineas * FontSize)) * -1
        End Select
        
            
        If intStart = 0 Then
            outText "1 0 0 1 " & CStr(tLeft) & " " & CStr(Top - LineCount * FontSize + AjusteV) & " Tm", True   'Posicionamos
            outText Replace(Format(WordSpacing, "0.00"), ",", ".") & " Tw", True
            outText "<" & StringToHex(DevuelvePalabras(strTexto, intStart, intX - 1)) & "> Tj", True
        ElseIf intX <= intPalabras Then
            outText "1 0 0 1 " & CStr(tLeft) & " " & CStr(Top - LineCount * FontSize + AjusteV) & " Tm", True   'Posicionamos
            outText Replace(Format(WordSpacing, "0.00"), ",", ".") & " Tw", True
            outText "<" & StringToHex(DevuelvePalabras(strTexto, intStart, intX - 1)) & "> Tj", True
        Else
            outText "1 0 0 1 " & CStr(tLeft) & " " & CStr(Top - LineCount * FontSize + AjusteV) & " Tm", True 'Posicionamos
            outText "0 Tw", True
            outText "<" & StringToHex(DevuelvePalabras(strTexto, intStart, intX - 1)) & "> Tj", True
        End If
        intStart = intX
    Loop
    
    
    

End Function

Private Function DevuelvePalabras(strTexto As String, intDesde As Integer, intHasta As Integer) As String

    Dim varAux  As Variant
    Dim strAux  As Variant
    Dim intX    As Integer
    
    
    'Eliminamos espaciado doble
    Do While InStr(strTexto, Space(2)) <> 0
        strTexto = Replace(strTexto, Space(2), Space(1))
    Loop
    
    varAux = Split(strTexto, Space(1))
    
    For intX = intDesde To intHasta
        strAux = strAux + varAux(intX) + Space(1)
    Next
    
    DevuelvePalabras = Trim(strAux)
    
End Function

Private Function CuentaPalabras(strTexto As String) As Integer

    Dim varAux  As Variant
    
    
    'Eliminamos espaciado doble
    Do While InStr(strTexto, Space(2)) <> 0
        strTexto = Replace(strTexto, Space(2), Space(1))
    Loop
    
    varAux = Split(strTexto, Space(1))
    
    
    CuentaPalabras = UBound(varAux) + 1
    
End Function
Friend Function CuentaLineas(strTexto As String, FontSize As Integer, Width As Integer) As Integer

    Dim intPalabras As Integer
    Dim intStart    As Integer
    Dim LineCount   As Integer
    Dim intX        As Integer
    Dim sPalabra    As String
    Dim Virtform As Form
    Set Virtform = New cFrm
    
    'Calculo el array de anchos de fuentes para usarlo después con textwidth
    Dim Buffer As Long
    Dim Tmetrics As TextMetric
        With Virtform
            .FontName = Me.BaseFont
            .FontSize = 750 'Esta es la altura para calcular el width de cada char
            GetTextMetrics .hdc, Tmetrics
        For i = Tmetrics.tmLastChar To Tmetrics.tmFirstChar Step -1
            GetCharWidth32 .hdc, i, i, Buffer
            Vararray(i) = Buffer
        Next i
    End With
    
    
    'Si hay alguna palabra cuyo ancho sea mayor que el ancho, la subdividimos en varias palabras
Reiteracion:
    intPalabras = CuentaPalabras(strTexto)
    For intX = 0 To intPalabras - 1
        sPalabra = DevuelvePalabras(strTexto, intX, intX)
        If TextWidth(sPalabra, FontSize, 1) > Width Then
            Do While TextWidth(sPalabra, FontSize, 1) > Width
                sPalabra = Mid(sPalabra, 1, Len(sPalabra) - 1)
            Loop
            strTexto = DevuelvePalabras(strTexto, 0, intX - 1) & " " & sPalabra & " " & Mid(DevuelvePalabras(strTexto, intX, intX), Len(sPalabra) + 1) & DevuelvePalabras(strTexto, intX + 1, intPalabras - 1)
           GoTo Reiteracion
        End If
    Next
    
    intStart = 0
    LineCount = 0

    Do While intStart < intPalabras
        For intX = intStart To intPalabras - 1
            LineWidth = TextWidth(DevuelvePalabras(strTexto, intStart, intX), FontSize, 0)
            If LineWidth > Width Then
                Exit For
            End If
        Next
        LineCount = LineCount + 1
        intStart = intX
    Loop
    
    CuentaLineas = LineCount
    
    Unload Virtform
    Set Virtform = Nothing

    
End Function
Friend Function TextWidth(Text As String, FontSize As Integer, Optional WordSpacing As Double = 1) As Integer
    Dim intX     As Integer
    Dim ToPoints As Double
    Dim Width    As Integer

    ToPoints = FontSize / 1000
    For intX = 1 To Len(Text)
        Width = Width + Vararray(Asc(Mid(Text, intX, 1))) * ToPoints
        If Mid(Text, intX, 1) = Space(1) Then
            Width = Width + WordSpacing
        End If
    Next
    TextWidth = Width
End Function

Private Function ToStr(Valore As Variant, Optional Dec As Integer = 3) As String
  ToStr = Replace(CStr(Round(Valore, Dec)), ",", ".")
End Function
Public Function LoadFontTTF(Name As String, BaseFont As String, _
                    Optional Style As String)
Dim i As Integer
Dim cadena As String
Dim strWidths As String
Dim Buffer As Long
Dim Tmetrics As TextMetric
Dim FuenteMasAncha As Long 'Compruebo 1 por 1 ya que tmetrics.tmmaxcharwidth no funciona bien
Dim Virtform As Form

FuenteMasAncha = 0
Set Virtform = New cFrm
    With Virtform
        .FontName = BaseFont
        .FontSize = 750
        Select Case Options
          Case Normal
              .FontBold = False
              .FontItalic = False
          Case Bold
              .FontBold = True
          Case Italic
              .FontItalic = True
          Case BoldItalic
              .FontBold = True
              .FontItalic = True
        End Select
        GetTextMetrics .hdc, Tmetrics
    For i = Tmetrics.tmLastChar To Tmetrics.tmFirstChar Step -1
        If Buffer > FuenteMasAncha Then FuenteMasAncha = Buffer
        GetCharWidth32 .hdc, i, i, Buffer
        cadena = i & "   " & Buffer & Chr(13) & Chr(10) & cadena
        strWidths = Buffer & " " & strWidths
    Next i
End With
    pdfObj = NextObj
    Objetos.Add pdfObj, glngOffSet
    outText pdfObj & " 0 obj", True
    outText "<<", True
    outText "/Type /FontDescriptor", True
    outText "/FontName /" & Replace(Virtform.FontName, " ", ""), True
    outText "/Flags 32 /FontBBox [ 0 " & Format(Tmetrics.tmDescent, "-0") & " " & FuenteMasAncha & " " & Tmetrics.tmAscent & " ] ", True
    outText "/StemV 80 ", True
    outText "/ItalicAngle 0 /CapHeight " & Tmetrics.tmAscent, True
    outText "/Ascent " & Tmetrics.tmAscent & " /Descent " & Format(Tmetrics.tmDescent, "-0"), True
    outText ">>", True
    outText "endobj", True
    
    pdfObj = NextObj
    Objetos.Add pdfObj, glngOffSet
    outText pdfObj & " 0 obj", True
    outText "<< /Type /Font ", True
    outText "/Subtype /TrueType", True
    outText "/Name /" & Name, True
    outText "/BaseFont /" & Replace(Virtform.FontName, " ", "") & Style, True
    outText "/FirstChar " & Tmetrics.tmFirstChar, True
    outText "/LastChar " & Tmetrics.tmLastChar, True
    outText "/FontDescriptor " & CStr(pdfObj - 1) & " 0 R", True
    outText "/Encoding /WinAnsiEncoding", True
    outText "/Widths [ " & strWidths & " ]", True
    outText ">>", True
    outText "endobj", True

Unload Virtform
Set Virtform = Nothing

End Function

Friend Function CuentaLineas2(strTexto As String, FontSize As Integer, Width As Integer) As Integer

    Dim intPalabras As Integer
    Dim intStart    As Integer
    Dim LineCount   As Integer
    Dim intX        As Integer
    Dim sPalabra    As String
    
    Dim Virtform As Form
    Set Virtform = New cFrm
    
    'Si hay alguna palabra cuyo ancho sea mayor que el ancho, la subdividimos en varias palabras
Reiteracion:
    intPalabras = CuentaPalabras(strTexto)
    For intX = 0 To intPalabras - 1
        sPalabra = DevuelvePalabras(strTexto, intX, intX)
        If TextWidth(sPalabra, FontSize, 1) > Width Then
            Do While TextWidth(sPalabra, FontSize, 1) > Width
                sPalabra = Mid(sPalabra, 1, Len(sPalabra) - 1)
            Loop
            strTexto = DevuelvePalabras(strTexto, 0, intX - 1) & " " & sPalabra & " " & Mid(DevuelvePalabras(strTexto, intX, intX), Len(sPalabra) + 1) & DevuelvePalabras(strTexto, intX + 1, intPalabras - 1)
           GoTo Reiteracion
        End If
    Next
    
    intStart = 0
    LineCount = 0

    Do While intStart < intPalabras
        For intX = intStart To intPalabras - 1
            LineWidth = TextWidth(DevuelvePalabras(strTexto, intStart, intX), FontSize, 0)
            If LineWidth > Width Then
                Exit For
            End If
        Next
        LineCount = LineCount + 1
        intStart = intX
    Loop
    
    CuentaLineas2 = LineCount
    
    Unload Virtform
    Set Virtform = Nothing
    
End Function


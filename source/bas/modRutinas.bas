Attribute VB_Name = "modRutinas"
Option Explicit

Global glngOffSet       As Long
Global Objetos          As cObjetos
'Para saber si ejecutamos en memoria y guardar los datos
Global gbolMemory       As Boolean
Dim outBytes()          As Byte

'Indicador de si hemos escrito la cabecera
Global gbolWCabecera    As Boolean
Dim mintHandle          As Integer
Global glngPageLength   As Long
Global gintPageHeight   As Integer
Global gintPageWidth    As Integer
Global gintObj          As Integer

'Apis
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetShellWindow Lib "user32.dll" () As Long





'b   closepath, fill,and stroke path.
'B   fill and stroke path.
'b*  closepath, eofill,and stroke path.
'B*  eofill and stroke path.
'BI  begin image.
'BMC     begin marked content.
'BT  begin text object.
'BX  begin section allowing undefined operators.
'c   curveto.
'cm  concat. Concatenates the matrix to the current transform.
'cs  setcolorspace for fill.
'CS  setcolorspace for stroke.
'd   setdash.
'Do  execute the named XObject.
'DP  mark a place in the content stream, with a dictionary.
'EI  end image.
'EMC     end marked content.
'ET  end text object.
'EX  end section that allows undefined operators.
'f   fill path.
'f*  eofill Even/odd fill path.
'g   setgray (fill).
'G   setgray (stroke).
'gs  set parameters in the extended graphics state.
'h   closepath.
'i   setflat.
'ID  begin image data.
'j   setlinejoin.
'J   setlinecap.
'k   setcmykcolor (fill).
'K   setcmykcolor (stroke).
'l   lineto.
'm   moveto.
'M   setmiterlimit.
'n   end path without fill or stroke.
'q   save graphics state.
'Q   restore graphics state.
're  rectangle.
'rg  setrgbcolor (fill).
'RG  setrgbcolor (stroke).
's   closepath and stroke path.
'S   stroke path.
'sc  setcolor (fill).
'SC  setcolor (stroke).
'sh  shfill (shaded fill).
'Tc  set character spacing.
'Td  move text current point.
'TD  move text current point and set leading.
'Tf  set font name and size.
'Tj  show text.
'TJ  show text, allowing individual character positioning.
'TL  set leading.
'Tm  set text matrix.
'Tr  set text rendering mode.
'Ts  set super/subscripting text rise.
'Tw  set word spacing.
'Tz  set horizontal scaling.
'T*  move to start of next line.
'v   curveto.
'w   setlinewidth.
'W   clip.
'y   curveto.



Public Sub OpenPDF(FileName As String)

    ShellExecute GetShellWindow, "open", FileName, vbNullString, App.Path, 1

End Sub

Public Sub PrintPDF(FileName As String)

    ShellExecute GetShellWindow, "print", FileName, vbNullString, App.Path, 0

End Sub
Public Sub CreateFile(FileName)

    'Por si acaso habia un documento en memoria anterior lo borro
    ReDim outBytes(0)
    gbolMemory = False
    
    mintHandle = FreeFile
    If FileExists(FileName) Then
        Kill FileName
    End If
    Open FileName For Output As mintHandle

End Sub

Public Sub MemoryNew()

    'Borro posible documento anterior
    ReDim outBytes(0)
    gbolMemory = True
    
End Sub

Public Function GetMemory() As Byte()

    GetMemory = outBytes
    
End Function

Public Sub FileClose()
    
    Close #mintHandle

End Sub


Public Sub outText(strText As String, Optional CrLf As Boolean = False)

  Dim lngStart As Long
  Dim intX     As Integer
  
  If Not gbolWCabecera Then
'    Call WCabecera
  End If
  
  strText = strText & IIf(CrLf, vbCr, "")
  glngOffSet = glngOffSet + Len(strText)
  If Not gbolMemory Then
    Print #mintHandle, strText;
  Else
     lngStart = UBound(outBytes)
     ReDim Preserve outBytes(lngStart + Len(strText))
     For intX = 1 To Len(strText)
        outBytes(lngStart + intX - 1) = Asc(Mid(strText, intX, 1))
     Next
  End If
    
End Sub

'Devuelve el proximo numero de Objeto
Public Function NextObj() As Integer

    gintObj = gintObj + 1
    'Reservo los numero de Objeto 2,3 y 4 para el grupo Pages y resources
    If gintObj = 2 Or gintObj = 3 Or gintObj = 4 Then
        gintObj = 5
    End If
    
    NextObj = gintObj
    
End Function

Private Function FileExists(ByVal FileName As String) As Boolean
  
  On Error Resume Next
  FileExists = FileLen(FileName) > 0
  Err.Clear
  
End Function

Public Function HexToString(strHex As String) As String

    Dim intX   As Long
    Dim strAux As String
    
    For intX = 1 To Len(strHex) Step 3
        strAux = strAux & Chr(Mid(strHex, intX, 3))
    Next
    
    HexToString = strAux
        
End Function

Public Function StringToHex(strString As String) As String

    Dim intX   As Long
    Dim strAux As String
    Dim strL1  As String
    Dim strL2  As String
    Dim strHex As String
    
    strHex = "0123456789ABCDEF"
    For intX = 1 To Len(strString)
        strL1 = Mid(strHex, Int(Asc(Mid(strString, intX, 1)) / 16) + 1, 1)
        strL2 = Mid(strHex, (Asc(Mid(strString, intX, 1)) Mod 16) + 1, 1)
        strAux = strAux & strL1 & strL2
    Next
    
    StringToHex = strAux
        
End Function
Public Function Max(ParamArray aValores() As Variant) As Variant

    Dim intX As Integer
    
    Max = aValores(LBound(aValores))
    For intX = LBound(aValores) To UBound(aValores)
        If aValores(intX) > Max Then
            Max = aValores(intX)
        End If
    Next
    
End Function


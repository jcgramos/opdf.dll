VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPDF"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
'Colecciones

Public Fonts        As cFonts
Private Pages       As cPaginas
Private Imagenes    As cImagenes
Private Outlines    As cOutlines
Public Bloques      As cBloques
Private Enlaces     As cEnlaces

'Propiedades Publicas
Public Author       As String
Public CreationDate As Variant
Public Creator      As String
Public Producer     As String
Public Title        As String
Public Subject      As String
Public Keywords     As String
Public FileName     As String
Attribute FileName.VB_VarUserMemId = 0
Public TextEncoding As eEncoding
'Variables privadas
Dim bolPageOpen     As Boolean
Dim intPageCount    As Integer

'Para guardar el Portapapeles
Dim mintClipboard   As Integer
Dim mvarClipboard   As Variant

'Modificaciones del Texto
Dim mColor          As Long
Dim mBT             As Boolean

'Añadidos finales
Dim glngOffPrevio   As Long



'BLOQUES
Public Enum pdfObjectType
  pdfNull = &H0
  pdfFirstPage = &H1
  pdfEvenPages = &H2
  pdfOddPages = &H4
  pdfNotFirstPage = &H8
  pdfAllPages = pdfEvenPages Or pdfOddPages Or pdfFirstPage
End Enum

Private LineasArchivo As String
'Del oPDF
Public Enum Page_Size
    A4_Vertical
    A4_Horizontal
    Folio_Vertical
    Folio_Horizontal
    UsarAnchoAlto
End Enum


Public Enum eLineCap
    ButtCap
    RoundCap
    ProyectingSquareCap
End Enum

Public Enum eLineJoin
    MilterJoin
    RoundJoin
    BevelJoin
End Enum

Public Enum eRenderMode
    FillText
    StrokeText
    FillStroke
    Invisible
    FillPath
    StrokePath
    FillStrokePath
    Path
End Enum
    
Public Enum eTextState
     BeginText
     SetFont
     SetColor
     SetPos
     TextShow
     NewLine
     CharSpace
     WordSpace
     HorizontalScale
     TextLeading
     RenderMode
     TextRise
     TransformationMatrix
     EndText

End Enum
Private OldFont As String
Public Property Get PageHeight() As Integer

    PageHeight = gintPageHeight
    
End Property

Public Property Get PageCount() As Integer
    
    PageCount = Pages.Count
    
End Property

Public Function GetPageObj(NumPage As Integer) As Integer


    GetPageObj = Pages(NumPage).pdfObj
    
End Function

Public Property Get PageWidth() As Integer

    PageWidth = gintPageWidth
    
End Property

Public Function GetTextWidth(Text As String, FontName As String, FontSize As Integer) As Double

    Dim oFont As New cFont
    
    Set oFont = Fonts(FontName)
    
    GetTextWidth = oFont.TextWidth(Text, FontSize)
    Set oFont = Nothing
    

End Function


Public Function GetCellHeight(ByVal Text As String, ByVal FontName As String, ByVal FontSize As Integer, ByVal CellWidth As Integer) As Double
    Dim oFont As New cFont
    Set oFont = Fonts(FontName)
    
    If FontName <> OldFont Then
        GetCellHeight = oFont.CuentaLineas(Text, FontSize, CellWidth) * FontSize
    Else
        GetCellHeight = oFont.CuentaLineas2(Text, FontSize, CellWidth) * FontSize
    End If
    OldFont = FontName
    Set oFont = Nothing
End Function

Public Function PDFCreate(Optional cFileName As String = "") As Boolean

    On Error GoTo procerror
    'En caso de no recibir el parametro tomamos el global, en otro caso igualamos el global
    If cFileName = "" Then
        cFileName = FileName
        Call MemoryNew
    Else
        FileName = cFileName
        Call CreateFile(FileName)
    End If
    'Obtenemos manejador de Archivo y borramos anterior
    Call WCabecera
    
    PDFCreate = True
        
Fin:
    Exit Function

procerror:
    If Err = 70 Then
        MsgBox "Error creando " & FileName & vbCrLf & "Es probable que el fichero ya se encuentre abierto", vbCritical, "oPDF.DLL Active X"
        PDFCreate = False
    Else
        MsgBox "Se ha producido el siguiente error: " & Error & " durante la generación del documento PDF", vbCritical, "oPDF.DLL Active X"
        PDFCreate = False
    End If
    Resume Fin
    Resume
    
End Function

Public Sub AddOutline(Name As String, Text As String, Optional Parent As String = 0, Optional Page As Integer = -1, Optional PosY As Integer = 0, Optional PosX As Integer = 0, Optional PercentZoom As Integer = 0)

    If Page = -1 Then
        Page = PageCount
    End If
    
    Outlines.Add Name, Text, Parent, Page, PosX, PosY, PercentZoom
    
End Sub

Public Sub Show()

    If Not gbolMemory Then
        OpenPDF (FileName)
    End If
    
End Sub

Public Function NewPage(ByVal nSize As Page_Size, Optional nWidth As Integer = 0, Optional nHeight As Integer = 0) As Variant

    Dim intObj    As Integer
    Dim strFonts  As String
    Dim strImages As String
    Dim strBloques As String
    Dim intX      As Integer
    
    
    'En el caso que haya una pagina anterior sin terminar la finalizo
    If bolPageOpen Then
        Call ClosePage
    End If
    
    Select Case nSize
    Case 0 'A4 Vertical
        gintPageWidth = 595.2
        gintPageHeight = 842
    Case 1 'A4 Horizontal
        gintPageWidth = 842
        gintPageHeight = 595.2
    Case 2 'Folio Vertical
        gintPageWidth = 612
        gintPageHeight = 792
    Case 3 'Folio Horizontal
        gintPageWidth = 792
        gintPageHeight = 612
    Case Else
        gintPageWidth = nWidth
        gintPageHeight = nHeight
    End Select
    
    
    'Inicializamos Coleccion de Enlaces
    Set Enlaces = New cEnlaces
    
''''''''''''''CADENAS
    intPageCount = intPageCount + 1
    intObj = NextObj
    
    'Guardo el Numero de Objeto de la pagina en el array de Paginas
    Pages.Add gintPageHeight, gintPageWidth
    Pages(Pages.Count).pdfObj = intObj
    Pages(Pages.Count).OffSet = glngOffSet
    'Añado a la tabla de objetos
    Objetos.Add intObj, glngOffSet
    outText intObj & " 0 obj", True
    outText "<<", True
    outText "/Type /Page", True
    outText "/Parent 3 0 R", True ' El Objeto 3 esta reservado para la definicion de Pages al Final
    outText "/MediaBox [ 0 0 " & gintPageWidth & " " & gintPageHeight & "]", True
    outText "/Contents " & intObj + 1 & " 0 R", True
    outText "/Annots " & intObj + 3 & " 0 R", True
'    outText "/Annots [<</Type /Annot /Subtype /Link /Rect [10 542 70 -100] /Border [0 0 0] /A <</S /URI /URI (www.google.es)>>>>]", True
    outText "/Rotate 0", True
    outText ">>", True
    outText "endobj", True
    
    
    
    
    intObj = NextObj
    'Añado a la tabla de objetos
    Objetos.Add intObj, glngOffSet
    
    outText intObj & " 0 obj", True
    outText "<<", True
    outText "/Length " & intObj + 1 & " 0 R", True
    outText ">>", True
    outText "stream", True
 
    
    'Flag de Pagina abierta y contador de tamaño en bytes de pagina
    bolPageOpen = True
    glngPageLength = glngOffSet



End Function


Public Function AddToPage(ByVal PageOriginal As Integer) As Variant

    Dim objPage   As Integer
    Dim intObj    As Integer
    Dim strFonts  As String
    Dim strImages As String
    Dim strBloques As String
    Dim intX      As Integer
    
    
    'En el caso que haya una pagina anterior sin terminar la finalizo
    If bolPageOpen Then
        Call ClosePage
    End If
    
    objPage = Pages(PageOriginal).pdfObj
    gintPageWidth = Pages(PageOriginal).Width
    gintPageHeight = Pages(PageOriginal).Height
   
''''''''''''''CADENAS
    intObj = NextObj
    
    'Añado a la tabla de objetos
    Objetos.Add intObj, glngOffSet
    outText objPage & " 0 obj", True
    outText "<<", True
    outText "/Type /Page", True
    outText "/Parent 3 0 R", True ' El Objeto 3 esta reservado para la definicion de Pages al Final
    outText "/MediaBox [ 0 0 " & gintPageWidth & " " & gintPageHeight & "]", True
    outText "/Contents [" & objPage + 1 & " 0 R " & intObj + 1 & " 0 R]", True
    outText "/Rotate 0", True
    outText ">>", True
    outText "endobj", True
    
    intObj = NextObj
    'Añado a la tabla de objetos
    Objetos.Add intObj, glngOffSet
    
    outText intObj & " 0 obj", True
    outText "<<", True
    outText "/Length " & intObj + 1 & " 0 R", True
    outText ">>", True
    outText "stream", True
 
    
    'Flag de Pagina abierta y contador de tamaño en bytes de pagina
    bolPageOpen = True
    glngPageLength = glngOffSet



End Function

Public Sub ClosePage()

    Dim lngPageLength As Long
    Dim intObj        As Integer
    Dim intX          As Integer
    
'    outText "Q", True
    lngPageLength = glngOffSet - glngPageLength
    
    outText "endstream", True
    outText "endobj", True
    
    intObj = NextObj
    'Añado a la tabla de objetos
    Objetos.Add intObj, glngOffSet
    
    outText intObj & " 0 obj", True
    outText CStr(lngPageLength), True
    outText "endobj", True

    
    'Ahora escribimos los enlaces
    intObj = NextObj
    'Añado a la tabla de objetos
    Objetos.Add intObj, glngOffSet
    outText intObj & " 0 obj", True
    outText "["
    For intX = 1 To Enlaces.Count
        outText "<</Type /Annot /Subtype /Link /Rect ["
        outText CStr(Enlaces.Item(intX).Left) & " " & CStr(Enlaces.Item(intX).Top) & " " & CStr(Enlaces.Item(intX).Width) & " " & CStr(Enlaces.Item(intX).Height * -1)
        outText "] /Border [0 0 0] /A <</S /URI /URI (" & Enlaces.Item(intX).Url & ")>>>>"
    Next
    outText "]", True
    outText "endobj", True


    bolPageOpen = False
    
End Sub

Public Sub PDFClose(Optional NotFileClose As Boolean = False)

    Dim lngPosition As Long
    Dim strPages    As String
    Dim intX        As Integer
    Dim intObj      As Integer
    
    If bolPageOpen Then
        Call ClosePage
    End If
    
'    'Objeto Resources

    'Creamos la cadena con los objetos Fuentes
    For intX = 1 To Fonts.Count
        Fonts(intX).WFont
        strFonts = strFonts & "/" & Fonts(intX).Name & " " & Fonts(intX).pdfObj & " 0 R" & vbCr
    Next
    
    'Creamos la cadena con los objetos Images
    For intX = 1 To Imagenes.Count
        strImages = strImages & "/" & Imagenes(intX).Name & " " & Imagenes(intX).pdfObj & " 0 R" & vbCr
    Next
    
    'Creamos la cadena con los Bloques
    For intX = 1 To Bloques.Count
        strBloques = strBloques & "/" & Bloques(intX).Name & " " & Bloques(intX).pdfObj & " 0 R" & vbCr
    Next

    Objetos.Add 4, glngOffSet
    outText "4 0 obj", True
    outText "<< ", True
    outText "/Font"
    outText "<< ", True
    outText Trim(strFonts), True
    outText ">>", True
    If Imagenes.Count > 0 Then
        outText "/ProcSet[/PDF /ImageB /Text]", True
        outText "/XObject ", True
        outText "<<", True
        outText Trim(strImages), True
        outText ">>", True
    End If
    If Bloques.Count > 0 Then
        outText "/XObject ", True
        outText "<<", True
        outText Trim(strBloques), True
        outText ">>", True
    End If
    outText ">>", True
    outText "endobj", True
    

'    'Objeto Catalogo
    Objetos.Add 2, glngOffSet
    outText "2 0 obj", True
    outText "<<", True
    outText "/Type /Catalog", True
    outText "/Pages 3 0 R", True
    If Outlines.Count > 0 Then
        intObj = NextObj
        outText "/Outlines " & CStr(intObj) & " 0 R", True
        outText "/PageMode /UseOutlines", True
    End If
    outText ">>", True
    outText "endobj", True
    
    If Outlines.Count > 0 Then
        Call WOutlines(intObj)
    End If
    
    
    For intX = 1 To Pages.Count
        strPages = strPages & Pages(intX).pdfObj & " 0 R "
    Next
    'Añado a la tabla de objetos
    Objetos.Add 3, glngOffSet
    
    outText "3 0 obj", True
    outText "<<", True
    outText "/Type /Pages", True
    outText "/Count " & Pages.Count, True
    outText "/Kids [ " & strPages & "]", True
    outText "/Resources 4 0 R", True
    outText ">>", True
    outText "endobj", True
    
    
    lngPosition = glngOffSet
    outText "xref", True
    outText "0 " & NextObj, True
    outText "0000000000 65535 f ", True

    For intX = 1 To Objetos.Count
      outText Format(Objetos("N" & CStr(intX)).OffSet, "0000000000") & " 00000 n ", True
    Next intX

    outText "trailer", True
    outText "<<", True
    outText "/Size " & gintObj, True
    outText "/Root 2 0 R", True
    outText "/Info 1 0 R", True
    outText ">>", True
    outText "startxref", True
    outText CStr(lngPosition), True
    outText "%%EOF", True
    
    'Por si hay añadidos
    glngOffPrevio = lngPosition
    If Not gbolMemory And Not NotFileClose Then
        Call FileClose
    End If
    

    
End Sub


Public Function BeginUpdate()

Call PDFClose(True)


Objetos.RemoveAll

End Function

Public Function EndUpdate()


    lngPosition = glngOffSet
    outText "xref", True
    outText "0 " & NextObj, True
    outText "0000000000 65535 f ", True

    For intX = 1 To Objetos.Count
      outText Format(Objetos(intX).OffSet, "0000000000") & " 00000 n ", True
    Next intX

    outText "trailer", True
    outText "<<", True
    outText "/Size " & gintObj, True
    outText "/Root 2 0 R", True
    outText "/Info 1 0 R", True
    outText "/Prev " & CStr(glngOffPrevio), True
    outText ">>", True
    outText "startxref", True
    outText CStr(lngPosition), True
    outText "%%EOF"
    
    'Por si hay añadidos
    glngOffPrevio = lngPosition
    If Not gbolMemory And Not NotFileClose Then
        Call FileClose
    End If
    


End Function


Public Function GetPDFStream() As Byte()

    If gbolMemory Then
        GetPDFStream = GetMemory()
    End If
    
End Function

Public Function SetTextColor(Optional ByVal Color As ColorConstants = vbBlack)

    mColor = Color
    
End Function

Public Sub WTextCode(Code As eTextState, ParamArray vData() As Variant)

    
    'NO haremos nada si no se ha llamado a un bloque de Texto
    If Not mBT And Code <> BeginText Then
        Exit Sub
    End If
    
    Select Case Code
    Case BeginText
        outText "BT", True
        mBT = True
    Case SetFont
        outText "/" & CStr(vData(0)) & " " & CStr(vData(1)) & " Tf", True
    Case SetColor
        outText Replace(Format(LongToRGB(CLng(vData(0)), Red) / 255, "0.0") & " " & Format(LongToRGB(CLng(vData(0)), Green) / 255, "0.0") & " " & Format(LongToRGB(CLng(vData(0)), Blue) / 255, "0.0"), ",", ".") & " rg", True
    Case SetPos
        outText CStr(vData(0)) & " " & CStr(Me.PageHeight - vData(1)) & " Td", True
    Case TextShow
        outText "<" & StringToHex(CStr(vData(0))) & "> Tj", True
    Case NewLine
        outText "T*", True
    Case CharSpace
        outText CStr(vData(0)) & " Tc", True
    Case WordSpace
        outText CStr(vData(0)) & " Tw", True
    Case HorizontalScale
        outText CStr(vData(0)) & " Tz", True
    Case TextLeading
        outText CStr(vData(0)) & " TL", True
    Case RenderMode
        outText CStr(vData(0)) & " Tr", True
    Case TextRise
        outText CStr(vData(0)) & " Ts", True
    Case TransformationMatrix
        outText Replace(Format(CDbl(vData(0)), "0.0000") & " " & Format(CDbl(vData(1)), "0.0000") & " " & Format(CDbl(vData(2)), "0.0000") & " " & Format(CDbl(vData(3)), "0.0000") & " " & Format(CDbl(vData(4)), "0.0000") & " " & Format(CDbl(vData(5)), "0.0000") & " Tm", ",", "."), True
    Case EndText
        outText "ET", True
        mBT = False
    End Select
    
End Sub

Public Function WCode(Text As String)

    outText Text, True
    
End Function
Public Function WText(ByVal Top As Integer, ByVal Left As Integer, Text As String, FontName As String, FontSize As Integer, Optional ByVal Rotation As Double = 0)

    Dim nTop        As Integer
    Dim oFont       As cFont
    Dim dblCos      As Double
    Dim dblSin      As Double
    
    Set oFont = Fonts(FontName)
  
    
    'Invertimos lsa coordenadas (Acrobat lee de abajo a arriba
    nTop = PageHeight - Top
    
    SaveGraphicState
    If Not mBT Then
        outText "BT", True
    End If
    outText "/" & FontName & " " & CStr(FontSize) & " Tf", True
    Rotation = Rotation * 3.1415 / 180
    dblCos = Cos(Rotation)
    dblSin = Sin(Rotation)
    outText Replace(Format(dblCos, "0.0000") & " " & Format(dblSin, "0.0000") & " " & Format(dblSin * -1, "0.0000") & " " & Format(dblCos, "0.0000") & " " & Format(Left, "0.0000") & " " & Format(nTop, "0.0000") & " Tm", ",", "."), True
    outText Replace(Format(LongToRGB(mColor, Red) / 255, "0.0") & " " & Format(LongToRGB(mColor, Green) / 255, "0.0") & " " & Format(LongToRGB(mColor, Blue) / 255, "0.0"), ",", ".") & " rg", True
    outText "<" & StringToHex(Text) & "> Tj", True
    If Not mBT Then
        outText "ET", True
    End If
    RestoreGraphicState


End Function

Public Function WGrid()

    Dim intX As Integer
    Dim intY As Integer
    
    
    SaveGraphicState
    MoveTo 0, 0
    For intX = 0 To PageHeight Step 10
        WLineTo PageWidth, intX, 0, intX
    Next
    MoveTo 0, 0
    For intX = 0 To PageWidth Step 10
        WLineTo intX, PageHeight, intX, 0
    Next
    LineStroke
    RestoreGraphicState
        
        
End Function
Public Function WTextBox(ByVal Top As Integer, ByVal Left As Integer, ByVal Height As Integer, ByVal Width As Integer, ByVal Text As String, FontName As String, FontSize As Integer, Optional hAlign As AlignH = hLeft, Optional vAlign As AlignV = vTop, Optional TextColor As ColorConstants = vbBlack, Optional Border As Integer = 0, Optional BorderColor As ColorConstants = vbBlack, Optional FillBox As Boolean = False, Optional TextMargin As Integer = 0, Optional Url As String = "")

    Dim oFont As cFont
    
    Set oFont = Fonts(FontName)
    If Border > 0 Then
        WRectangle Top, Left, Height, Width, Border, BorderColor, FillBox
    End If
    
    If Url <> "" Then
        Enlaces.Add Url, PageHeight - Top, Left, Height, Width
    End If
    'Aplicamos el margen para el texto
    If TextMargin > 0 Then
        Top = Top + TextMargin
        Left = Left + TextMargin
        Height = Max(Height - (2 * TextMargin), FontSize)
        Width = Max(Width - (2 * TextMargin), FontSize)
    End If
    
    outText "BT", True
    outText "/" & FontName & " " & CStr(FontSize) & " Tf", True
    outText Replace(Format(LongToRGB(TextColor, Red) / 255, "0.0") & " " & Format(LongToRGB(TextColor, Green) / 255, "0.0") & " " & Format(LongToRGB(TextColor, Blue) / 255, "0.0"), ",", ".") & " rg", True
    Call oFont.ImprimeTexto(Top, Left, Height, Width, Text, FontSize, vAlign, hAlign, FontName)
    outText "ET", True
    
    Set oFont = Nothing
    
End Function
Public Function SetColors(Optional ForeColor As ColorConstants = -1, Optional FillColor As ColorConstants = -1)

    If Not FillColor = -1 Then
        outText Replace(Format(LongToRGB(FillColor, Red) / 255, "0.0") & " " & Format(LongToRGB(FillColor, Green) / 255, "0.0") & " " & Format(LongToRGB(FillColor, Blue) / 255, "0.0"), ",", ".") & " rg", True
    End If
    
    If Not ForeColor = -1 Then
        outText Replace(Format(LongToRGB(ForeColor, Red) / 255, "0.0") & " " & Format(LongToRGB(ForeColor, Green) / 255, "0.0") & " " & Format(LongToRGB(ForeColor, Blue) / 255, "0.0"), ",", ".") & " RG", True
    End If
    
    
End Function


Public Function SetLineFormat(Optional LineWidth As Byte = 1, Optional LineCap As eLineCap = ButtCap, Optional LineJoin As eLineJoin = BevelJoin, Optional DashPhase As Integer = 0, Optional DashArray As String = "[]")

    outText CStr(LineWidth) & " w", True
    outText CStr(LineCap) & " J", True
    outText CStr(LineJoin) & " j", True
    outText DashArray & " " & CStr(DashPhase) & " d", True

End Function

Public Function MoveTo(ByVal X As Double, ByVal Y As Double)

    'Invertimos el eje de coordenadas
    Y = PageHeight - Y
    
    outText Replace(Format(X, "0.00") & " " & Format(Y, "0.00"), ",", ".") & " m", True

End Function

Public Function WLineTo(ByVal X As Double, ByVal Y As Double, Optional ByVal StartX As Double, Optional ByVal StartY As Double)

    'Invertimos el eje de coordenadas
    Y = PageHeight - Y
    
    
    If StartX + StartY <> 0 Then
        MoveTo StartX, StartY
    End If
    outText CStr(X) & " " & CInt(Y) & " l", True
    
End Function

Public Function SaveGraphicState()
    
    outText "q", True

End Function

Public Function RestoreGraphicState()
    
    outText "Q", True

End Function

Public Sub LineStroke()

    outText "S", True
    
End Sub

Public Sub WCurve(ByVal X As Double, ByVal Y As Double, Optional ByVal ControlX1 As Integer, Optional ByVal ControlY1 As Integer, Optional ByVal ControlX2 As Integer, Optional ByVal ControlY2 As Integer, Optional ByVal StartX As Double, Optional ByVal StartY As Double)

    'Invertimos el eje de coordenadas
    Y = PageHeight - Y
    ControlY1 = PageHeight - ControlY1
    ControlY2 = PageHeight - ControlY2
    
    If StartY <> 0 Then
        StartY = PageHeight - StartY
    End If
        
    'Moveto Inicial
    If StartX + StartY <> 0 Then
        MoveTo StartX, StartY
    End If

    'No se especificaron puntos de control
    If ControlX1 + ControlY1 + ControlX2 + ControlY2 = 0 Then
        WLineTo X, Y, StartX, StartY
        Exit Sub
    End If
    
    'Solo Punto de Control 1, Referenciamos con StartPoint
    If ControlX2 + ControlY2 = 0 Then
        outText CStr(ControlX1) & " " & CInt(ControlY1) & " " & CStr(X) & " " & CInt(Y) & " v", True
    End If
    
    'Solo Punto de Control 2, Referenciamos con EndPoint
    If ControlX1 + ControlY1 = 0 Then
        outText CStr(ControlX2) & " " & CInt(ControlY2) & " " & CStr(X) & " " & CInt(Y) & " y", True
    End If
    
    'Dos puntos de Control
    outText CStr(ControlX1) & " " & CInt(ControlY1) & " " & CStr(ControlX2) & " " & CInt(ControlY2) & " " & CStr(X) & " " & CInt(Y) & " c", True
    
        
End Sub

Public Sub WCircle(ByVal X As Double, ByVal Y As Double, Radio As Double)
 
    
    'Invertimos el eje de coordenadas
    Y = PageHeight - Y
    
    outText "q", True
    outText Replace(Format(X + Radio, "0.00") & " " & Format(Y, "0.00"), ",", ".") & " m", True
    outText Replace(Format(X + Radio, "0.00") & " " & Format(Y + Radio * 0.552, "0.00") & " " & Format(X + Radio * 0.552, "0.00") & " " & Format(Y + Radio, "0.00") & " " & Format(X, "0.00") & " " & Format(Y + Radio, "0.00"), ",", ".") & " c", True
    outText Replace(Format(X - Radio * 0.552, "0.00") & " " & Format(Y + Radio, "0.00") & " " & Format(X - Radio, "0.00") & " " & Format(Y + Radio * 0.552, "0.00") & " " & Format(X - Radio, "0.00") & " " & Format(Y, "0.00"), ",", ".") & " c", True
    outText Replace(Format(X - Radio, "0.00") & " " & Format(Y - Radio * 0.552, "0.00") & " " & Format(X - Radio * 0.552, "0.00") & " " & Format(Y - Radio, "0.00") & " " & Format(X, "0.00") & " " & Format(Y - Radio, "0.00"), ",", ".") & " c", True
    outText Replace(Format(X + Radio * 0.552, "0.00") & " " & Format(Y - Radio, "0.00") & " " & Format(X + Radio, "0.00") & " " & Format(Y - Radio * 0.552, "0.00") & " " & Format(X + Radio, "0.00") & " " & Format(Y, "0.00"), ",", ".") & " c", True
    outText "s", True
    outText "Q", True

 End Sub

Public Sub WRectangle(ByVal Top As Integer, ByVal Left As Integer, ByVal Height As Integer, ByVal Width As Integer, Optional ByVal Border As Integer = 1, Optional BorderColor As ColorConstants = vbBlack, Optional Fill As Boolean = False)

    Dim nTop As Integer
    
    'Invertimos lsa coordenadas (Acrobat lee de abajo a arriba
    nTop = PageHeight - Top

    outText "q", True
    outText Border & " w", True
    outText Replace(Format(LongToRGB(BorderColor, Red) / 255, "0.0") & " " & Format(LongToRGB(BorderColor, Green) / 255, "0.0") & " " & Format(LongToRGB(BorderColor, Blue) / 255, "0.0"), ",", ".") & " RG", True
    outText CStr(Left) & " " & CStr(nTop) & " " & CStr(Width) & " " & CStr(Height * -1) & " re", True
    If Fill Then
        outText Replace(Format(LongToRGB(BorderColor, Red) / 255, "0.0") & " " & Format(LongToRGB(BorderColor, Green) / 255, "0.0") & " " & Format(LongToRGB(BorderColor, Blue) / 255, "0.0"), ",", ".") & " rg", True
        outText "f", True
    End If
    outText "S", True
    outText "Q", True


End Sub

Public Sub LoadImageFromFile(FileName As String, Name As String, Optional GrayScale As Boolean = False, Optional Zipped As Boolean = False)

    Dim objPicture As IPictureDisp
    If Dir(FileName) <> "" Then
        Set objPicture = LoadPicture(FileName)
        If Not objPicture Is Nothing Then
            CargarImagen objPicture, Name, GrayScale, Zipped, False
        End If
    End If
    
End Sub

Public Sub LoadImage(Image As Variant, Name As String, Optional GrayScale As Boolean = False, Optional Zipped As Boolean = False)
    
    Call CargarImagen(Image, Name, GrayScale, Zipped, True)

End Sub

Private Sub CargarImagen(Image As Variant, Name As String, Optional GrayScale As Boolean = False, Optional Zipped As Boolean = False, Optional UseClipboard As Boolean = True)
    Dim intObj    As Integer
    Dim bitAux()  As Byte
    Dim intX      As Long
    Dim lngWidth  As Long
    Dim lngHeight As Long
    Dim strAux    As String
    
    'Pasamos por el portapapeles para unificar Pictures e Images (Feo pero no encuentro otra manera)
    If UseClipboard Then
        SaveClipBoard
        Clipboard.SetData Image, 2
        bitAux = PDFAdaptor(Clipboard.GetData(2), GrayScale, Zipped)
        lngWidth = PictureWidth(Clipboard.GetData(2))
        lngHeight = PictureHeight(Clipboard.GetData(2))
        Clipboard.Clear
        RestoreClipboard
    Else
        bitAux = PDFAdaptor(CLng(Image), GrayScale, Zipped)
        lngWidth = PictureWidth(CLng(Image))
        lngHeight = PictureHeight(CLng(Image))
    End If

    'Añado a la tabla de objetos
    intObj = NextObj
    Objetos.Add intObj, glngOffSet
    
    outText intObj & " 0 obj", True
    outText "<<", True
    outText "/Type /XObject", True
    outText "/Subtype /Image", True
    outText "/Name /" & Name, True
    outText "/BitsPerComponent 8", True
    outText "/ColorSpace /DeviceRGB", True
    If Not Zipped Then
        outText "/Filter []", True
    Else
        outText "/Filter /FlateDecode", True
    End If
    outText "/Width " & CStr(lngWidth), True
    outText "/Height " & CStr(lngHeight), True
    outText "/Length " & CStr(UBound(bitAux) - 10), True
    outText ">>", True
    outText "stream", True
    For intX = 1 To UBound(bitAux) - 1
        outText Chr(bitAux(intX))
    Next
    outText Asc(bitAux(UBound(bitAux))), True
    outText "endstream", True
    outText "endobj", True
    
    'Guardo en el array de Objetos Imagenes
    Imagenes.Add
    Imagenes(Imagenes.Count).pdfObj = intObj
    Imagenes(Imagenes.Count).OffSet = glngOffSet
    Imagenes(Imagenes.Count).Name = Name
    
End Sub

Public Sub WImage(ByVal Top As Long, ByVal Left As Long, ByVal Height As Long, ByVal Width As Long, ImageName As String, Optional ByVal Rotation As Double = 0)
    
    Dim nTop As Integer
    
    'Invertimos lsa coordenadas (Acrobat lee de abajo a arriba
    nTop = PageHeight - Top
   
    outText "q", True
    outText "1 0 0 1 " & CStr(Left) & " " & CStr(nTop) & " cm", True 'Posicionamos
    If Rotation <> 0 Then
        Rotation = Rotation * 3.1415 / 180
        outText Replace(Format(Cos(Rotation), "0.0000") & " " & Format(Sin(Rotation), "0.0000") & " " & Format(Sin(Rotation) * -1, "0.0000") & " " & Format(Cos(Rotation), "0.0000") & " 0 0 cm ", ",", "."), True
    End If
    outText CStr(Width) & " 0 0 " & CStr(Height) & " 0 0  cm", True ' Redimensionamos
    
    outText "/" & ImageName & " Do", True
    outText "Q", True

End Sub
'Escribe la Cabecera del PDF
Private Function WCabecera()

    Dim intObj  As Integer
    Dim strDate As String
   
    strDate = "D:" & Format(CreationDate, "YYYYMMDDHHNNSS")
    'Inicio de documento (cadenas fijas)
    outText "%PDF-1.3", True
    outText "%âãÏÓ", True
    
    'Añado a la tabla de objetos
    intObj = NextObj
    Objetos.Add intObj, glngOffSet
    
    outText intObj & " 0 obj", True
    outText "<<", True
    outText "/Author (" & Author & ")", True
    outText "/CreationDate (" & strDate & ")", True
    outText "/Creator (" & Creator & ")", True
    outText "/Producer (" & Producer & ")", True
    outText "/Title (" & Title & ")", True
    outText "/Subject (" & Subject & ")", True
    outText "/Keywords (" & Keywords & ")", True
    outText ">>", True
    outText "endobj", True
    
    gbolWCabecera = True
    
End Function

Private Sub WOutlines(intParent As Integer)

    Dim intObj    As Integer
    Dim obj       As cOutline
    Dim intX      As Integer
    
    'Añado a la tabla de objetos
    Objetos.Add intParent, glngOffSet
    
    outText intParent & " 0 obj", True
    outText "<<", True
    outText "/Type /Outlines", True
    outText "/First " & CStr(intParent + 1) & " 0 R", True
    outText "/Last " & CStr(intParent + Outlines.Count) & " 0 R", True
    outText "/Count " & CStr(Outlines.Count), True
    outText ">>", True
    outText "endobj", True
    
    'Primero asigno el numero de objeto a todos los Outlines
    For intX = 1 To Outlines.Count
        Set obj = Outlines.Item(intX)
        obj.pdfObj = NextObj
    Next
    
    For intX = 1 To Outlines.Count
        Set obj = Outlines.Item(intX)
        intObj = obj.pdfObj
        Objetos.Add intObj, glngOffSet
        outText intObj & " 0 obj", True
        outText "<<", True
        outText "/Title < " & StringToHex(obj.Descripcion) & " > ", True
        'Si tiene hijos que dependan de su rama
        If obj.Children > 0 Then
            outText "/First " & CStr(obj.FirstChilden.pdfObj) & " 0 R", True
            outText "/Last " & CStr(obj.LastChilden.pdfObj) & " 0 R", True
            outText "/Count " & CStr(obj.Children), True
        End If
        If obj.Parent <> "0" Then
            outText "/Parent " & CStr(Outlines.Item(obj.Parent).pdfObj) & " 0 R", True
        Else
            outText "/Parent " & CStr(intParent) & " 0 R", True
        End If
        
        If Not obj.GetPrev Is Nothing Then
            outText "/Prev " & CStr(obj.GetPrev.pdfObj) & " 0 R", True
        End If
        If Not obj.GetNext Is Nothing Then
            outText "/Next " & CStr(obj.GetNext.pdfObj) & " 0 R", True
        End If
        outText "/Dest " & obj.GetDest(Pages(obj.Page)), True
        outText ">>", True
        outText "endobj", True
    Next
    
    
    
    
End Sub

Private Sub SaveClipBoard()

    If Clipboard.GetFormat(1) Then
        mintClipboard = 1
        mvarClipboard = Clipboard.GetText
    Else
        mintClipboard = 0
    End If

End Sub

Private Sub RestoreClipboard()

    If mintClipboard = 1 Then
        Clipboard.SetText CStr(mvarClipboard)
    End If
    
End Sub


Private Sub Class_Initialize()

 Set Fonts = New cFonts
 Set Pages = New cPaginas
 Set Objetos = New cObjetos
 Set Imagenes = New cImagenes
 Set Outlines = New cOutlines
 Set Bloques = New cBloques
 
 bolPageOpen = False
 intPageCount = 0
 gintObj = 0
 glngOffSet = 0
 TextEncoding = WinAnsiEncoding
 
 
 Author = "Desconocido"
 CreationDate = Now
 Creator = "Desconocido"
 Producer = "oPdf Writer"
 Title = "Desconocido"
 Subject = "Desconocido"
 Keywords = "---"
 mColor = vbBlack

End Sub


Private Sub Class_Terminate()

    Fonts.RemoveAll
    Set Fonts = Nothing
    Pages.RemoveAll
    Set Pages = Nothing
    Objetos.RemoveAll
    Set Objetos = Nothing
    Imagenes.RemoveAll
    Set Imagenes = Nothing
    Outlines.RemoveAll
    Set Outlines = Nothing
    Bloques.RemoveAll
    Set Bloques = Nothing
End Sub
Public Sub CreaBloque(Name As String, Archivo As String)
'añade un objeto a la coleccion
    Dim intObj  As Integer
    
    intObj = NextObj
    Objetos.Add intObj, glngOffSet
    
    outText intObj & " 0 obj", True
    outText "<<", True
    outText "/Type /XObject", True
    outText "/Subtype /Form", True
    outText "/FormType 1", True
    outText "/BBox [0 0 595.2 842]", True
    outText "/Matrix [1 0 0 1 0 0]", True
    outText "/Name /" & Name, True
    outText "/ColorSpace /DeviceRGB", True
    outText ">>", True
    outText "stream", True
    
    LeerArchivo Archivo
    outText LineasArchivo, True
    
    outText "endstream", True
    outText "endobj", True

    'Guardo en el array de Bloques
    Bloques.Add
    Bloques(Bloques.Count).pdfObj = intObj
    Bloques(Bloques.Count).OffSet = glngOffSet
    Bloques(Bloques.Count).Name = Name
End Sub
Public Function WBloque(Name As String)
    outText "/" & Name & " Do", True
End Function
Private Function LeerArchivo(ArchivoparaLeer As String)
    Dim obj_FSO As Object
    Dim Archivo As Object
    
    Set obj_FSO = CreateObject("Scripting.FileSystemObject")
    Set Archivo = obj_FSO.OpenTextFile(ArchivoparaLeer, 1)

    LineasArchivo = Archivo.ReadAll
    
    '///Nota: para leer una linea se utiliza ReadLine
    Archivo.Close
    Set obj_FSO = Nothing
    Set Archivo = Nothing
End Function


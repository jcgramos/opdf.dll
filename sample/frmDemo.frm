VERSION 5.00
Begin VB.Form frmDemo 
   Caption         =   "Demostracion de Libreria DLL"
   ClientHeight    =   2640
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   7200
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2640
   ScaleWidth      =   7200
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdGenerar 
      Caption         =   "Generar PDF"
      Default         =   -1  'True
      Height          =   495
      Left            =   4680
      TabIndex        =   0
      Top             =   1800
      Width           =   1815
   End
   Begin VB.Image Image1 
      Height          =   1260
      Left            =   60
      Picture         =   "frmDemo.frx":0000
      Stretch         =   -1  'True
      Top             =   60
      Width           =   3120
   End
End
Attribute VB_Name = "frmDemo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public LineasEscudo As String
Private Sub cmdGenerar_Click()

    Dim oDoc  As cPDF
    Dim dblX  As Double
    Dim dblY  As Double
    Dim Angle As Double
    Dim Pi    As Double
    Dim aPages As Integer
    
    Set oDoc = New cPDF
    
    
    oDoc.Author = "Juan de Los Palotes"
    'Creo el documento
    If Not oDoc.PDFCreate("c:\Prueba.pdf") Then
        Exit Sub
    End If
    

    'Defino los fuentes
    oDoc.Fonts.Add "F1", "Code 128", TrueType
    oDoc.Fonts.Add "F2", "Times New Roman", TrueType
    oDoc.Fonts.Add "F3", "Palace Script MT", TrueType
    oDoc.Fonts.Add "F4", "Courier New", TrueType
'
'    'Cargo las imagenes, para comprimir necesita zlib.dll (Libre)
    oDoc.LoadImage Image1, "Logo", False, False
    oDoc.LoadImage Image1, "LogoGris", True, False
'
    oDoc.NewPage A4_Vertical
    'Muestro la imagen
    oDoc.WImage 100, 200, 61, 323, "Logo"


    'Ejemplo de circulos
    'DashArray = "Puntos On, Puntos Off ...)
    oDoc.SetLineFormat 5, , , 0, "[10 2]"
    For dblX = 10 To 50 Step 10
        oDoc.WCircle 100, 200, dblX
    Next
'
'    'Ejemplo de Lineas ( si no especifico startX y StartY coge la ultima posicion )
    oDoc.SetLineFormat 10, RoundCap, RoundJoin
    oDoc.MoveTo 200, 200
    oDoc.WLineTo 250, 200
    oDoc.WLineTo 250, 250
    oDoc.LineStroke
'
'    'Ejemplo de Lineas ( si no especifico startX y StartY coge la ultima posicion )
    oDoc.SetLineFormat 10, ProyectingSquareCap, BevelJoin
    oDoc.MoveTo 300, 200
    oDoc.WLineTo 350, 200
    oDoc.WLineTo 350, 250
    oDoc.LineStroke
''
    oDoc.AddOutline "Bookmark1", "Pagina 1"
    'Ejemplo de curvas, Canvas en radianes, quien las entienda que las compre ;)
    oDoc.MoveTo 400, 200
    oDoc.WCurve 450, 200, 451, 250, 449, 250
    oDoc.LineStroke


    'Un Código de barras
    oDoc.WTextBox 400, 200, 100, 150, "0123456789", "F1", 25
    
    


'   'Varias cajas de texto
    oDoc.NewPage A4_Vertical
    oDoc.WImage 100, 200, 61, 323, "Logo"
    
    oDoc.AddOutline "Bookmark2", "Pagina 2"
    oDoc.AddOutline "Bookmark21", "Caja 4", "Bookmark2", , 300, 310
    
    
    oDoc.WTextBox 300, 10, 100, 70, "Esta es una caja de texto con borde magenta", "F2", 10, , , , 1, vbMagenta, , , "www.google.es"
    oDoc.WTextBox 300, 110, 100, 70, "Esta es una caja dé texto con borde negro alineada a la derecha", "F2", 10, hRight, , , 1, vbBlack
    oDoc.WTextBox 300, 210, 100, 70, "Esta es una caja de texto sin borde y con el texto justificado", "F2", 10, hjustify
    oDoc.WTextBox 300, 310, 100, 70, "Esta es una caja de texto con borde y doble centrado", "F2", 10, hCenter, vMiddle, vbBlue, 1, vbGreen
'
    dblY = oDoc.GetCellHeight("Esto es un ejemplo de altura adaptada al tamaño del texto", "F2", 10, 70)
'
    oDoc.WTextBox 300, 410, dblY, 70, "Esto es un ejemplo de altura adaptada al tamaño del texto", "F2", 10, hjustify, , , 1
    
    oDoc.ClosePage
    
    oDoc.NewPage A4_Vertical
    oDoc.WImage 100, 200, 61, 323, "LogoGris", -5
    oDoc.AddOutline "Bookmark3", "Pagina 3"
    
    oDoc.WTextBox 300, 210, 100, 70, "Enlace a oPDF", "F2", 10, , , , , , , , "http://www.opdf.tk"
    Call CajaRedondeada(oDoc, 200, 400, 450, 500)
    
    oDoc.ClosePage
    
    oDoc.NewPage A4_Vertical
    oDoc.WImage 100, 200, 61, 323, "Logo"
    oDoc.AddOutline "Bookmark4", "Pagina 4"
    
    
    'Ahora WText con angulo
    Pi = 3.14159265
    For Angle = 0 To 2 * Pi Step Pi / 4
        dblX = 50 * Cos(Angle)
        dblY = 50 * Sin(Angle)

        oDoc.WText 200 + dblX, 300 + dblY, "Angulo", "F2", 8, (Angle * 180 / Pi)
    Next
    
    'Lo utilizo para algo util
    
    oDoc.SetTextColor RGB(195, 195, 195)
    
    oDoc.WText 250, 150, "Marca de agua", "F2", 60, -45
        
    
    oDoc.ClosePage

    oDoc.BeginUpdate
    For intX = 1 To oDoc.PageCount
        oDoc.AddToPage intX
        oDoc.WTextBox 800, 350, 100, 250, "Pagina " & CStr(intX) & " de " & CStr(oDoc.PageCount), "F4", 25
        oDoc.ClosePage
    Next
    oDoc.EndUpdate
    
       

    oDoc.Show
   
    

End Sub


Private Sub CajaRedondeada(ByRef oDoc As Variant, X As Long, Y As Long, X1 As Long, Y1 As Long)



oDoc.MoveTo X1 - (X1 - X) / 2, Y
oDoc.WCurve X1, Y1 - (Y1 - Y) / 2, X1, Y, X1, Y
oDoc.WCurve X1 - (X1 - X) / 2, Y1, X1, Y1, X1, Y1
oDoc.WCurve X, Y1 - (Y1 - Y) / 2, X, Y1, X, Y1
oDoc.WCurve X1 - (X1 - X) / 2, Y, X, Y, X, Y
oDoc.LineStroke


End Sub


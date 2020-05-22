VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cOutline"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Public pdfObj       As Integer
Public OffSet       As Long
Public Name         As String
Public Descripcion  As String
Public Parent       As String
Public Page         As Integer
Public PosX         As Integer
Public PosY         As Integer
Public Zoom         As Integer
Private mOutlines   As cOutlines


Friend Property Let POutlines(vData As cOutlines)

    Set mOutlines = vData
    
End Property

Friend Property Get Children() As Integer

    Dim obj     As cOutline
    Dim Aux     As Integer
    Dim intX    As Integer
    
    For intX = 1 To mOutlines.Count
        Set obj = mOutlines.Item(intX)
        If obj.Parent = Me.Name Then
            Aux = Aux + 1
        End If
    Next
    
    Children = Aux
            
End Property

Friend Property Get FirstChilden() As cOutline

    Dim obj     As cOutline
    Dim intX    As Integer
    
    For intX = 1 To mOutlines.Count
        Set obj = mOutlines.Item(intX)
        If obj.Parent = Me.Name Then
            Set FirstChilden = obj
            Exit For
        End If
    Next
    
End Property

Friend Property Get LastChilden() As cOutline

    Dim obj     As cOutline
    Dim intX    As Integer

    
    For intX = 1 To mOutlines.Count
        Set obj = mOutlines.Item(intX)
        If obj.Parent = Me.Name Then
            Set LastChilden = obj
        End If
    Next

End Property

Friend Property Get GetPrev() As cOutline

    Dim obj     As cOutline
    Dim intX    As Integer
    
    For intX = 1 To mOutlines.Count
        Set obj = mOutlines.Item(intX)
        If obj.Parent = Me.Parent And obj.pdfObj < Me.pdfObj Then
            Set GetPrev = obj
        End If
    Next

End Property

Friend Property Get GetNext() As cOutline

    Dim obj     As cOutline
    Dim intX    As Integer
    
    For intX = 1 To mOutlines.Count
        Set obj = mOutlines.Item(intX)
        If obj.Parent = Me.Parent And obj.pdfObj > Me.pdfObj Then
            Set GetNext = obj
            Exit For
        End If
    Next

End Property

Friend Property Get GetDest(oPage As cObjeto) As String

    Dim strAux As String
    
    strAux = "[" & oPage.pdfObj & " 0 R /XYZ "
    strAux = strAux & CStr(PosX) & " "
    strAux = strAux & CStr(oPage.Height - PosY) & " "
    strAux = strAux & IIf(Zoom <> 0, Replace(CStr(Zoom / 100), ",", "."), "null") & " ]"
    
    GetDest = strAux
    
End Property



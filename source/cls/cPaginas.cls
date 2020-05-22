VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cPaginas"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
Dim mCol As Collection

Public Function Add(Optional Height As Integer = 0, Optional Width As Integer = 0) As cObjeto
    
    On Error GoTo procerror
    
    Dim objPagina As cObjeto
    Set objPagina = New cObjeto
    objPagina.Height = Height
    objPagina.Width = Width
    
    mCol.Add objPagina
    Set Add = objPagina
    Exit Function
    
procerror:
    Resume Fin
    Resume
Fin:
End Function

Public Sub RemoveAll()
    Set mCol = New Collection
End Sub

Private Sub Class_Initialize()
    Set mCol = New Collection
End Sub

Private Sub Class_Terminate()
    
    On Error GoTo procerror

    Set mCol = Nothing
    Exit Sub
    
procerror:
    Resume Fin
    Resume
Fin:
End Sub

Public Property Get Count() As Long
    'se usa al obtener el número de elementos de la
    'colección. Sintaxis: Debug.Print x.Count
    Count = mCol.Count
End Property

Public Property Get NewEnum() As IUnknown
    'esta propiedad permite enumerar
    'esta colección con la sintaxis For...Each
    Set NewEnum = mCol.[_NewEnum]
End Property

Public Property Get Item(vntIndexKey As Variant) As cObjeto
Attribute Item.VB_UserMemId = 0
    'se usa al hacer referencia a un elemento de la colección
    'vntIndexKey contiene el índice o la clave de la colección,
    'por lo que se declara como un Variant
    'Syntax: Set foo = x.Item(xyz) or Set foo = x.Item(5)
  Set Item = mCol(vntIndexKey)
End Property



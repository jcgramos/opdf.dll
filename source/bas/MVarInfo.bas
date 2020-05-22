Attribute VB_Name = "MVarInfo"
Option Explicit

'******************************************************************************
'*
'* Module MVarInfo
'*
'* This module implements one method called gGetVarInfo, which delivers information
'* about variables, especially array variables, without the need of provoking errors.
'*
'* Use it to decide: is it an array var
'*                   is it an undimmed array (aka does not have a SA descriptor)
'*                   is it a zombie array (dimmed with UBound < LBound)
'*                   how many dimensions has the array
'*                   the bounds of the array's dimensions (stored in *reverse*
'*                   order, see below)
'*
'* Note: can not be used with arrays of basetype fixed string
'*       if used with UDTs, the UDT has to be defined in a public object module
'*
'*
'* (c) 2004 Ulrich Korndörfer proSource software development
'*          www.prosource.de
'*          German site with VB articles (in english) and code (comments in english)
'*
'* Precautions: None. May be compiled to native code (which is strongly recommended),
'*              with all extended options selected.
'*
'* External dependencies: some APIs (see below)
'*
'* Version history
'*
'* Version 1.0 from 2004.09.23
'*
'*  Disclaimer:
'*  All code in this class is for demonstration purposes only.
'*  It may be used in a production environment, as it is thoroughly tested,
'*  but do not hold us responsible for any consequences resulting of using it.
'*
'******************************************************************************


'******************************************************************************
'* API declarations and consts
'******************************************************************************

Private Declare Sub CopyMem4 Lib "msvbvm60.dll" Alias "GetMem4" _
       (ByVal FromAddr As Long, ByRef ToAddr As Any)
Private Declare Sub CopyMem8 Lib "msvbvm60.dll" Alias "GetMem8" _
       (ByVal FromAddr As Long, ByRef ToAddr As Any)

Private Const VARIANT_DATA_OFFSET As Long = 8
Private Const SAD_LEN As Long = 16
Private Const SADBOUND_LEN As Long = 8


'******************************************************************************
'* Types
'******************************************************************************

'TVarInfo_SADBOUND is translated to VB from the SAFEARRAYBOUND structure as defined in MSDN
'Note: HighBound =  UBound(Arr) = LowBound + ElementCount - 1

Private Type TVarInfo_SADBOUND
  ElementCount As Long 'How many elements this dimension has
  LowBound As Long     'The minimum index of this dimension (equals LBound(Arr))
End Type

'TVarInfo holds all gathered information about the variable.
'The part starting with SAD_DimCount and ending with SAD_DataAdress is translated
'from the SAFEARRAY structure as defined in MSDN.

Public Type TVarInfo

  'Variable's type
  
  Var_BaseType As VbVarType 'Base type like vbSingle etc. The vbArray part has been removed.
  Var_IsArray As Boolean    'Is it an array var (vbArray).
  
  'If VarIsArray = True and a SA descriptor has been set for it
  '(if the array has been "dimmed"), then Var_SADAdress <> 0.
  
  Var_SADAdress As Long   'Adress of the SA descriptor
  
  'If Var_SADAdress <> 0, the SA descriptors data follow here.
  'The following is a 1:1 translation of the SAFEARRAY struct as defined in the MSDN.
  'It has a length of 16 bytes.
  
  SAD_DimCount As Integer 'How many dimensions the array has
  SAD_Features As Integer 'Features flags
  SAD_ElementSize As Long 'The size of the arrays element type
  SAD_Locks As Long       'How many locks are applied
  SAD_DataAddress As Long 'Address of the memory location, where the data starts
  
  'If SAD_DimCount > 0, the bounds of the dimensions follow here.
  'Note: a VB array with a SA descriptor set (that is with Var_SADAdress <> 0)
  'always has a SAD_DimCount > 0
  'Note: the dims are stored in *reverse* order!
  '      Dim 1 has index SAD_DimCount - 1, Dim n has index 0!
  
  SAD_Bounds() As TVarInfo_SADBOUND

End Type


'******************************************************************************
'* Public methods
'******************************************************************************

'Not allowed array types are:

'- Arrays of fixed strings, eg. Dim FSA() As String * 10.
'  Those can not be wrapped into a Variant!

'Noteworthy array types are:

'- Arrays of object references, eg. Dim OA() As CMyClass. Those can be used,
'  but note: an *undimmed* array of any object type always is (by VB)
'  created as zombie array!

'- Arrays of UDTs, eg. Dim UA() As TMyType. Those can be used only if TMyType is
'  declared as public in a public object module! Also note that, like object arrays,
'  an *undimmed* array of any UDT-type always is (by VB) created as zombie array!

Public Function gGetVarInfo(ByRef theVar As Variant) As TVarInfo
Dim i As Long

With gGetVarInfo
  
  .Var_BaseType = VarType(theVar)
  .Var_IsArray = (.Var_BaseType And vbArray) <> 0
  .Var_BaseType = .Var_BaseType And Not vbArray
  
  If .Var_IsArray Then
    
    'Get the adress of SAD (the array var's safearray descriptor)
    
    CopyMem4 UAdd(VarPtr(theVar), VARIANT_DATA_OFFSET), .Var_SADAdress
    CopyMem4 .Var_SADAdress, .Var_SADAdress
    
    'If it has none, it is "undimmed". Exit then as there is no more data available
    
    If .Var_SADAdress = 0 Then Exit Function

    'Copy the content of the array's SAD into our type
    
    CopyMem8 .Var_SADAdress, .SAD_DimCount
    CopyMem8 UAdd(.Var_SADAdress, 8), .SAD_Locks
    
    'If the array has no dimensions, there also are no bounds. So exit then.
    
    If .SAD_DimCount = 0 Then Exit Function
  
    'Copy the array bounds into our type.
    
    ReDim .SAD_Bounds(0 To .SAD_DimCount - 1)
    For i = 0 To .SAD_DimCount - 1
      CopyMem8 UAdd(.Var_SADAdress, SAD_LEN + i * 8), .SAD_Bounds(i)
    Next i
    
  End If

End With

End Function


'******************************************************************************
'* Private helpers
'******************************************************************************

'UAdd does an unsigned add of Base and Offset.

'Base and Offset are treated as *unsigned* longs and are added unsigned, with:

'&H0 <= Base        <= &HFFFFFFFF and
'&H0 <= Offset      <= &H7FFFFFFF and of course for the unsigned sum
'&H0 <= Base+Offset <= &HFFFFFFFF

'This would be the code for an unrestricted unsigned add:

'&H0 <= Base        <= &HFFFFFFFF and
'&H0 <= Offset      <= &HFFFFFFFF and
'&H0 <= Base+Offset <= &HFFFFFFFF

'UAdd = (((Base Xor &H80000000) + (Offset And &H7FFFFFFF)) Xor &H80000000) Or (Offset And &H80000000)

'In both cases overflow behaviour is complex and has to be investigated.

Private Function UAdd(ByVal Base As Long, ByVal Offset As Long) As Long
UAdd = ((Base Xor &H80000000) + Offset) Xor &H80000000
End Function


Attribute VB_Name = "MCommon"
Option Explicit

'******************************************************************************
'*
'* Module MCommon
'*
'* Assisting module to CCipher, CDataBlock and CHash. Implements checks and error handling
'* helpers used by these classes.
'*
'* (c) 2004 Ulrich Korndörfer proSource software development
'*          www.prosource.de
'*          German site with VB articles (in english) and code (comments in english)
'*
'* Precautions: None. May be compiled to native code (which is strongly recommended),
'*              with all extended options selected.
'*
'* External dependencies: MVarInfo
'*
'* Version history
'*
'* Version 1.1 from 2004.09.23 Now uses MVarInfo. This allows to use the extended
'*                             options when compiling to native code.
'* Version 1.0 from 2004.04.01
'*
'*  Disclaimer:
'*  All code in this class is for demonstration purposes only.
'*  It may be used in a production environment, as it is thoroughly tested,
'*  but do not hold us responsible for any consequences resulting of using it.
'*
'******************************************************************************


'******************************************************************************
'* Private consts
'******************************************************************************

Private Const cRANGEFROM As Long = 1
Private Const cRANGETO As Long = (17 * 17 - 1) '288
'Private Const cRANGETO As Long = (29 * 29 - 1) '840


'******************************************************************************
'* Public array checking methods
'******************************************************************************

'Arr must be initialized as one-dimensional array and have at least one byte (no zombie).
'Arr may be static or dynamic.

'If given, Low or High both must be valid indexes to Arr (including type and magnitude)
'If both Low and High are given, Low must be lower equal High
'Sets L, H and Length

Public Sub gCheckArray(ByRef Arr() As Byte, _
                       ByRef Low As Variant, _
                       ByRef High As Variant, _
                       ByRef l As Long, _
                       ByRef h As Long, _
                       ByRef Length As Long, _
                       ByRef ClassName As String, _
                       ByRef MethodName As String, _
                       ByRef Description As String)
Dim lh As Long

On Error GoTo MethodError

With gGetVarInfo(Arr)
  If .SAD_DimCount <> 1 Then GoTo ErrorExit 'Either not initialized or wrong dimension
  With .SAD_Bounds(0)
    If .ElementCount = 0 Then GoTo ErrorExit 'Is zombie
    lh = .LowBound + .ElementCount - 1
    'Get boundaries. Throws error if Low or High can not be converted to a long value
    If IsMissing(Low) Then l = .LowBound Else l = Low
    If IsMissing(High) Then h = lh Else h = High
    'Check, if L and H are valid indexes
    If l < .LowBound Or l > lh Then GoTo ErrorExit
    If h < .LowBound Or h > lh Then GoTo ErrorExit
    'Check, if H >= L
    Length = h - l + 1
    If (Length < 1) Then GoTo ErrorExit
  End With
End With

Exit Sub

ErrorExit:
  On Error GoTo 0
MethodError:
  gRaiseError ClassName, MethodName, Description, 5

End Sub

'Arr must be initialized as one-dimensional array and have at least one byte (no zombie),
'starting at Index 0.
'Arr may be static or dynamic.

'Pos must be a valid index to Arr.
'Sets Pos and Length

Public Sub gSimpleCheckArray(ByRef Arr() As Byte, _
                             ByRef Length As Long, _
                             ByVal Pos As Long, _
                             ByRef ClassName As String, _
                             ByRef MethodName As String, _
                             ByRef Description As String)

On Error GoTo MethodError

With gGetVarInfo(Arr)
  If .SAD_DimCount <> 1 Then GoTo ErrorExit 'Either not initialized or wrong dimension
  With .SAD_Bounds(0)
    If .ElementCount = 0 Or .LowBound <> 0 Then GoTo ErrorExit 'Is zombie or has wrong base index
    If Pos < 0 Or Pos >= .ElementCount Then GoTo ErrorExit 'Pos is invalid index
    Length = .ElementCount
  End With
End With

Exit Sub

ErrorExit:
  On Error GoTo 0
MethodError:
  gRaiseError ClassName, MethodName, Description, 5

End Sub


'******************************************************************************
'* Public others
'******************************************************************************

'gCheckPrimeInRange checks:

'- Is Value in the range cRANGEFROM to cRANGETO (including the borders)
'- Is Value a prime number from this range.

'gPrintPrimes is a testroutine for gCheckPrimeInRange, also callable directly
'in the debug window, which prints 62 found primes between 1 and 288:

'    1    2    3    5    7   11   13   17   19   23
'   29   31   37   41   43   47   53   59   61   67
'   71   73   79   83   89   97  101  103  107  109
'  113  127  131  137  139  149  151  157  163  167
'  173  179  181  191  193  197  199  211  223  227
'  229  233  239  241  251  257  263  269  271  277
'  281  283

'Further primes are:

'            293  307  311  313  317  331  337  347
'  349  353  359  367  373  379  383  389  397  401
'  409  419  421  431  433  439  443  449  457  461
'  463  467  479  487  491  499  503  509  521  523
'  541  547  557  563  569  571  577  587  593  599
'  601  607  613  617  619  631  641  643  647  653
'  659  661  673  677  683  691  701  709  719  727
'  733  739  743  751  757  761  769  773  787  797
'  809  811  821  823  827  829  839

'Between and including 17 and 251 there are 48 primes.

#If TESTMODE = 1 Then

  Public Sub gPrintPrimes()
  Dim i As Long, s As String, C As Long
  
  For i = cRANGEFROM To cRANGETO
    If gCheckPrimeInRange(i) Then
      s = s & FInt(i, 5): C = C + 1
      If C Mod 10 = 0 Then Debug.Print s: s = vbNullString: C = 0
    End If
  Next i
  If C <> 0 Then Debug.Print s
  End Sub

  Private Function FInt(ByVal Val As String, ByVal PadLen As Long) As String
  Dim l As Long
  l = Len(Val): If l >= PadLen Then FInt = Val Else FInt = Space$(PadLen - l) & Val
  End Function


#End If

Public Function gCheckPrimeInRange(ByVal Value As Long) As Boolean

Select Case True
  Case Value < cRANGEFROM
  Case Value > cRANGETO
  Case (Value Mod 2) = 0:  gCheckPrimeInRange = (Value = 2)
  Case (Value Mod 3) = 0:  gCheckPrimeInRange = (Value = 3)
  Case (Value Mod 5) = 0:  gCheckPrimeInRange = (Value = 5)
  Case (Value Mod 7) = 0:  gCheckPrimeInRange = (Value = 7)
  Case (Value Mod 11) = 0: gCheckPrimeInRange = (Value = 11)
  Case (Value Mod 13) = 0: gCheckPrimeInRange = (Value = 13) 'For cRANGETO = 288
'  Case (Value Mod 17) = 0: gCheckPrimeInRange = (Value = 17)
'  Case (Value Mod 19) = 0: gCheckPrimeInRange = (Value = 19)
'  Case (Value Mod 23) = 0: gCheckPrimeInRange = (Value = 23) 'For cRANGETO = 840
  Case Else:               gCheckPrimeInRange = True
End Select
End Function


'******************************************************************************
'* Public error handling helpers
'******************************************************************************

Public Sub gCheckCond(ByVal ErrCond As Boolean, _
                      ByRef ClassName As String, _
                      ByRef MethodName As String, _
                      ByRef Description As String)
If ErrCond Then gRaiseError ClassName, MethodName, Description, 5
End Sub

Public Sub gRaiseError(ByRef ClassName As String, _
                       ByRef MethodName As String, _
                       Optional ByRef Description As String, _
                       Optional ByVal Number As Long)

If Err.Number = 0 Then
  Err.Raise Number, ClassName & "." & MethodName, Description
Else
  If Len(Description) = 0 Then
    Err.Raise Err.Number, ClassName & "." & MethodName & vbCrLf & Err.Source, Err.Description
  Else
    Err.Raise Err.Number, ClassName & "." & MethodName & vbCrLf & Err.Source, Description & vbCrLf & Err.Description
  End If
End If

End Sub


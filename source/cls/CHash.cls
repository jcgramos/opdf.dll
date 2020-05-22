VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CHash"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'******************************************************************************
'*
'* Class CHash
'*
'* This class implements various hashes (1 is currently implemented):
'*
'* 1. CRC32 hash
'*
'* Adopted from:  PNG_CRC32Algo.c (PNG V1.2 CRC32 algo)
'*
'* Equivalent to: CRC32 as to be used or defined with:
'*                - PNG graphics V1.2
'*                - ISO 3309 or ITU-T V.42
'*                - PKZip
'*
'*
'* (c) 2004 Ulrich Korndörfer proSource software development
'*          www.prosource.de
'*          German site with VB articles (in english) and code (comments in english)
'*
'* Precautions: None. May be compiled to native code (which is strongly recommended),
'*              with all extended options selected.
'*
'* External dependencies: - module MCommon
'*
'* Version history
'*
'* Version 1.0 from 2004.09.18
'*
'*  Disclaimer:
'*  All code in this class is for demonstration purposes only.
'*  It may be used in a production environment, as it is thoroughly tested,
'*  but do not hold us responsible for any consequences resulting of using it.
'*
'******************************************************************************

Private Const mcClassName As String = "CHash"


'******************************************************************************
'* Private consts
'******************************************************************************

'*****
'* For the CRC32 checksum
'*****

'This is the official polynomial used by CRC32 in PKZip.
'It is of reverse type. The unreversed value is &H04C11DB7
'When using a reverse type poly, no overflow handling on calculating the CRC is necessary.
Private Const CRC32_POLYNOMIAL As Long = &HEDB88320
'This is the register initialization value used in PKZIP's CRC32
Private Const CRC32_REGISTERINIT As Long = &HFFFFFFFF


'******************************************************************************
'* Private member vars
'******************************************************************************

'*****
'* For the CRC32 checksum
'*****

Private mCRC32_LUT() As Long


'******************************************************************************
'* Public properties and methods
'******************************************************************************

'*****
'* CRC32 checksum
'*****

'Takes a byte array and calculates the CRC32 for all bytes between
'the indexes Low and High, then returns it as the functions result.

'If Low and/or High is missing, Low is set to LBound(Arr),
'High is set to UBound(Arr).

'Set ModeAdd to True, if you want a "rolling" CRC.
'In this case, the CRC calculated up to then has to be supplied with parameter
'CRC. So at least once CRC32 must be have called with ModeAdd=False.

Public Function CRC32(ByRef Arr() As Byte, _
                      Optional ByVal ModeAdd As Boolean = False, _
                      Optional ByVal CRC As Variant, _
                      Optional ByVal Low As Variant, _
                      Optional ByVal High As Variant) As Long
Const cMETHODNAME As String = "CRC32"
Dim l As Long, h As Long

gCheckArray Arr, Low, High, l, h, (0), _
            mcClassName, cMETHODNAME, "Input params invalid"

If ModeAdd Then
  gCheckCond IsMissing(CRC), mcClassName, cMETHODNAME, "Missing CRC register value"
  CRC32 = Not CalcCRC32(Not CLng(CRC), Arr, l, h)
Else
  CRC32 = Not CalcCRC32(CRC32_REGISTERINIT, Arr, l, h)
End If

End Function

Private Function CalcCRC32(ByVal Register As Long, _
                           ByRef Arr() As Byte, _
                           ByVal Low As Long, _
                           ByVal High As Long) As Long
Dim i As Long, Lookup As Long

For i = Low To High
  Lookup = (Register Xor Arr(i)) And &HFF&
  Register = ((Register And &HFFFFFF00) \ &H100&) And &HFFFFFF  '*unsigned* shr 8 bit
  Register = Register Xor mCRC32_LUT(Lookup)
Next i

CalcCRC32 = Register
End Function

'Fills the Lookuptable used by the CRC32 checksum. Must be called at least once
'before using any other of the CRC32 methods!

Public Sub CRC32_Init()

Dim i As Long, j As Long, Current As Long

ReDim mCRC32_LUT(0 To 255)

For i = 1 To 255
  Current = i
  For j = 1 To 8
    If (Current And 1&) <> 0 Then
      Current = (((Current And &HFFFFFFFE) \ &H2&) And &H7FFFFFFF) Xor CRC32_POLYNOMIAL
    Else
      Current = ((Current And &HFFFFFFFE) \ &H2&) And &H7FFFFFFF '*unsigned* shift right 1 bit
    End If
  Next j
  mCRC32_LUT(i) = Current
Next i

End Sub


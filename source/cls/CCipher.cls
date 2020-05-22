VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CCipher"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'******************************************************************************
'*
'* Class CCipher
'*
'* This class implements the ARCFOUR stream cipher (see 'A Stream Cipher Encryption
'* Algorithm "Arcfour"' from fourth of July 1999, an Internet Engineering Task Force (IETF)
'* draft described in the file "<draft-kaukonen-cipher-arcfour-03.txt>", available
'* elsewhere on the internet). ARCFOUR is a very secure stream cipher, which seems
'* to be never having been broken until now. And ARCFOUR is a very fast encryption
'* algorithm (apparantly due to its ingenious simplicity).
'*
'* ARCFOUR is believed to yield identical results to the RC4 stream cipher, which is
'* intellectual property and trademark of RSA Data Security, Inc. (www.rsa.com) and
'* which never has been made public officially. This implementation follows closely
'* aforementioned IETF draft for the ARCFOUR stream cipher and yields identical results
'* to the test vectors supplied with this draft and to other known test vectors.
'*
'* Due to the nature of this stream cipher (which in fact is a pseudo random generator
'* generating a random sequence of bytes, seeded by an up to 256 byte length key), it
'* also can be used to get uniformly distributed sequences of random bytes, which are
'* well fitted to cryptographic needs. Methods for this purpose are included.
'* Also included are de/encoding methods, which additionally use a one time pad for
'* further enhancing security. For better security also a method is implemented which
'* allows to advance in the stream of generated random bytes before starting to use them
'* for encryption. This fights against special classes of keys, which are assumed to be
'* 'weak'.
'*
'* Notes: - If one uses the same key for encrypting more than one message,
'*          this may lead to possible very easy means (depending on the nature
'*          of the messages) for an attacker to be able to decrypt the messages!
'*          The key will not be cracked, but the messages possibly will! And
'*          further messages using the same key then can be cracked very easy too,
'*          because the random stream of bytes used for encryption has been detected.
'*          Also if by any means the attacker manages to get access to an encrypted
'*          message *and* the original message, he can easily retreive the random sequence
'*          used for encryption (but again not the key) and therefore easily can decode
'*          further messages if they are encrypted with the same key (equal keys generate
'*          equal random byte sequences).
'*
'*        - Use keys of at least 128 bits (16 bytes) length. This should be sufficiently
'*          safe to resist brute force attacks which try to guess the key. Of course
'*          greater key lengths do not hurt and increase security. Remember that key
'*          length does not influence in anyway the performance of the algorithm.
'*          A 1 byte key takes the same processing time as a 256 byte key.
'*
'*        - Preferably do not use "natural text" keys or so called "pass phrases",
'*          instead use random like byte sequences. Those can be derived from
'*          text strings by using some hashing algorithms like SHA1.
'*
'*        - Avoid using "periodic" keys (keys which are build from repeated short
'*          byte sequences like "a1 23 c2 a1 23 c2 ... a1 23 c2"). An attacker which
'*          uses a brute force attack by trying all possible keys, starting with small
'*          ones, will be very early successful, as the "effective" length of those kind
'*          of keys is reduced to the length of the period (above example then collapses
'*          to a key of an effective length of 3 bytes or 24 bit). Using key lengths
'*          which are prime numbers and where not all bytes are the same is sufficient
'*          for avoiding "periodic" keys.
'*
'*        - Whenever possible, advance the state machine before starting encryption.
'*          Advance at least 256 cycles, but the more cycles you advance, the better
'*          is the enhancement in security. Advancing one cycle costs nearly
'*          the same processing time as the encryption of one byte does. But, as
'*          ARCFOUR is very fast, this should not hinder from advancing.
'*
'*        - After seeding and encryption, as soon as possible clear the internal state.
'*          This enhances security if attackers try to break into an application
'*          which uses this class.
'*
'*
'* (c) 2004 Ulrich Korndörfer proSource software development
'*          www.prosource.de
'*          German site with VB articles (in english) and code (comments in english)
'*
'* Precautions: - When compiled to native code (which is strongly recommended),
'*                array bound checking must not be switched off. This is due to
'*                the array checking routines in MCommon.
'*
'* External dependencies: - module MCommon
'*                        - some APIs (see below)
'*
'* Version history
'*
'* Version 1.0.1 from 2004.07.12
'*  Fixed bug in DoDeEncode
'*
'* Version 1.0 from 2004.04.01
'*
'*  Disclaimer:
'*  All code in this class is for demonstration purposes only.
'*  It may be used in a production environment, as it is thoroughly tested,
'*  but do not hold us responsible for any consequences resulting of using it.
'*
'******************************************************************************

Private Const mcClassName As String = "CCipher"


'******************************************************************************
'* API declarations
'******************************************************************************

Private Declare Function CoCreateGuid Lib "OLE32.dll" (ByRef GuidPtr As Any) As Long
Private Const S_OK As Long = &H0


'******************************************************************************
'* Private types
'******************************************************************************

'Type used for the various states
Private Type TState
  X As Long
  Y As Long
  State(0 To 255) As Long
End Type

'Helper type for converting byte arrays to currency values and vice versa
Private Type TCurrency
  Val As Currency
End Type

'Helper type for converting byte arrays to currency values and vice versa
Private Type TCurrencyByte
  Val(0 To 7) As Byte
End Type

'Helper type for converting byte arrays to long values and vice versa
Private Type TLong
  Val As Long
End Type

'Helper type for converting byte arrays to long values and vice versa
Private Type TLongByte
  Val(0 To 3) As Byte
End Type

'Helper type for converting byte arrays to integer values and vice versa
Private Type TInt
  Val As Integer
End Type

'Helper type for converting byte arrays to integer values and vice versa
Private Type TIntByte
  Val(0 To 2) As Byte
End Type


'******************************************************************************
'* Private member vars
'******************************************************************************

Private mCurrent As TState 'Holds the current state
Private mSeed As TState 'Helper holding the initial state (state just after seeding)
Private mTemplate As TState 'Helper holding the unseeded state, used as starting point for seeding
Private mIsNotSeeded As Boolean 'Helper telling if we have been seeded


'******************************************************************************
'* Public methods
'******************************************************************************

'Resets the instance to unseeded state

Public Sub Clear()
DoClear
End Sub

'Resets a seeded instance to its initial state it had just after seeding
'It is equivalent to calling Seed again using the former seed key, but faster.

'The class instance must have been seeded.

Public Sub Reset()

CheckSeeded "Reset"
LSet mCurrent = mSeed

End Sub

'Seeds the state arrays using a byte array as key. Key must be a onedimensional array
'with at least one element. Key() is not altered by the routine.

'Use Low and/or High to select a range from the input array to be used as key
'If given, Low and/or High must be valid indexes to Key(), and Low must be lower
'equal than High
'If Low is missing, LBound(Key) is used.
'If High is missing, UBound(Key) is used.

Public Sub Seed(ByRef Key() As Byte, _
                Optional ByVal Low As Variant, _
                Optional ByVal High As Variant)

Const cMETHODNAME As String = "Seed"
Dim L As Long, KeyLen As Long

DoClear
gCheckArray Key, Low, High, L, (0), KeyLen, _
            mcClassName, cMETHODNAME, "Input params invalid"

DoSeed Key, L, KeyLen
End Sub

Public Sub SeedFromLong(ByVal Key As Long)

Const cMETHODNAME As String = "SeedFromLong"
Dim B As TLongByte, L As TLong

DoClear
L.Val = Key: LSet B = L
DoSeed B.Val, 0, 4
End Sub

Public Sub SeedFromCurrency(ByVal Key As Currency)

Const cMETHODNAME As String = "SeedFromCurrency"
Dim B As TCurrencyByte, C As TCurrency

DoClear
C.Val = Key: LSet B = C
DoSeed B.Val, 0, 7
End Sub

'Seeds the state array using a key which is made by concatenating fresh GUIDs
'(16 byte len each), until the given KeyLen is reached.

'KeyLen must be in the range 1 to 256

Public Sub SeedFromGUID(ByVal KeyLen As Long)

Const cMETHODNAME As String = "SeedFromGUID"
Dim i As Long, Key() As Byte, Count As Long, Res As Long

DoClear
gCheckCond (KeyLen < 1 Or KeyLen > 256), _
           mcClassName, cMETHODNAME, "Invalid key len"

If (KeyLen Mod 16) = 0 Then Count = KeyLen \ 16 Else Count = KeyLen \ 16 + 1
ReDim Key(0 To Count * 16 - 1)

For i = 0 To Count - 1
  Res = CoCreateGuid(Key(i * 16))
  gCheckCond (Res <> S_OK), mcClassName, cMETHODNAME, "Could not create GUID"
Next i

DoSeed Key, 0, KeyLen
End Sub

'Advances a seeded state by the given number of cycles, that is:
'calculate Cycles times a new state, but do not use the state.
'Call it for avoiding the influence of "weak seeding keys" just after seeding
'with a value for Cycles "long enough" (at least >= 256). As default here
'the prime number 521 is used.

'The class instance must have been seeded.
'Cycles must be a positive integer value or 0. If 0, no advance is done.

Public Sub Advance(Optional ByVal Cycles As Long = 521)

CheckSeeded "Advance"
gCheckCond (Cycles < 0), mcClassName, "Advance", "Invalid value"

DoAdvance Cycles
End Sub

'De- or encodes the given byte array Arr(). Arr must be a onedimensional array
'with at least one element. The class instance must have been seeded.

'Use Low and/or High to select a range from the input array to be decoded or encoded.
'If given, Low and/or High must be valid indexes to Arr(), and Low must be lower
'equal than High.
'If Low is missing, LBound(Arr) is used.
'If High is missing, UBound(Arr) is used.

'The elements of Arr() between Low and High are altered as de/encoding is in place,
'all other elements of Arr() are untouched.


Public Sub DeEncode(ByRef Arr() As Byte, _
                    Optional ByVal Low As Variant, _
                    Optional ByVal High As Variant)

Const cMETHODNAME As String = "DeEncode"
Dim L As Long, H As Long

CheckSeeded cMETHODNAME
gCheckArray Arr, Low, High, L, H, (0), _
            mcClassName, cMETHODNAME, "Input params invalid"

DoDeEncode Arr, L, H
End Sub

'De- or encodes the given byte array Arr() additionaly using a one time pad array (OTP).
'Arr() and OTP() must be onedimensional arrays with at least one element. OTP() additionally
'must be zero based. The class instance must have been seeded.

'An additional XOR is made with the bytes from the given OTP starting with the byte at
'index OTPPos. For each byte to de/encode OTPPos is advanced by 1 in a ring counter fashion.

'The elements of Arr() between Low and High are altered as de/encoding is in place,
'all other elements of Arr() are untouched.

'OTP() stays untouched. OTPPos on method entry must be a valid index to OTPPos.
'On method exit OTPPos is set to the last used position + 1.

'Use Low and/or High to select a range from the input array to be decoded or encoded.
'If given, Low and/or High must be valid indexes to Arr(), and Low must be lower
'equal than High.
'If Low is missing, LBound(Arr) is used.
'If High is missing, UBound(Arr) is used.

Public Sub DeEncodeWithOTP(ByRef Arr() As Byte, _
                           ByRef OTP() As Byte, _
                           ByRef OTPPos As Long, _
                           Optional ByVal Low As Variant, _
                           Optional ByVal High As Variant)

Const cMETHODNAME As String = "DeEncodeWithOTP"
Dim OTPLen As Long, L As Long, H As Long

CheckSeeded cMETHODNAME
gCheckArray Arr, Low, High, L, H, (0), _
            mcClassName, cMETHODNAME, "Input array or its params invalid"
gSimpleCheckArray OTP, OTPLen, OTPPos, _
            mcClassName, cMETHODNAME, "OTP array or its params invalid"

DoDeEncodeWithOTP Arr, OTP, OTPPos, OTPLen, L, H
End Sub

'Returns the next 8 random bytes packed into a (signed) currency value.
'The class instance must have been seeded.

Public Function NextRandomCurrency() As Currency
Dim C As TCurrency, B As TCurrencyByte

CheckSeeded "NextRandomCurrency"

DoDeEncode B.Val, 0, 7: LSet C = B

NextRandomCurrency = C.Val
End Function

'Returns the next 4 random bytes packed into a (signed) long value.
'The class instance must have been seeded.

Public Function NextRandomLong() As Long
Dim L As TLong, B As TLongByte

CheckSeeded "NextRandomLong"

DoDeEncode B.Val, 0, 3: LSet L = B

NextRandomLong = L.Val
End Function

'Returns the next 2 random bytes packed into a (signed) integer value.
'The class instance must have been seeded.

Public Function NextRandomInt() As Integer
Dim i As TInt, B As TIntByte

CheckSeeded "NextRandomInt"

DoDeEncode B.Val, 0, 1: LSet i = B

NextRandomInt = i.Val
End Function

'Returns the next random byte.
'The class instance must have been seeded.

Public Function NextRandomByte() As Byte
Static Arr(0 To 0) As Byte

CheckSeeded "NextRandomByte"
DoDeEncode Arr, 0, 0

NextRandomByte = Arr(0)
End Function

'Returns a 0-based array of random currencies of length Count. Consecutive 8 random bytes
'have been packed into one (signed) currency value. Count must be greater 0.
'The class instance must have been seeded.

Public Function NextRandomCurrencyArr(ByVal Count As Long) As Currency()
Const cMETHODNAME As String = "NextRandomCurrencyArr"
Dim Arr() As Currency, L As TCurrency, B As TCurrencyByte, i As Long

CheckSeeded cMETHODNAME
CheckCount Count, cMETHODNAME

ReDim Arr(0 To Count - 1)
For i = 0 To Count - 1
  DoDeEncode B.Val, 0, 7: LSet L = B: Arr(i) = L.Val
Next i
NextRandomCurrencyArr = Arr
End Function

'Returns a 0-based array of random longs of length Count. Consecutive 4 random bytes have
'been packed into one (signed) long value. Count must be greater 0.
'The class instance must have been seeded.

Public Function NextRandomLongArr(ByVal Count As Long) As Long()
Const cMETHODNAME As String = "NextRandomLongArr"
Dim Arr() As Long, L As TLong, B As TLongByte, i As Long

CheckSeeded cMETHODNAME
CheckCount Count, cMETHODNAME

ReDim Arr(0 To Count - 1)
For i = 0 To Count - 1
  DoDeEncode B.Val, 0, 3: LSet L = B: Arr(i) = L.Val
Next i
NextRandomLongArr = Arr
End Function

'Returns a 0-based array of random integers of length Count. Consecutive 2 random bytes have
'been packed into one (signed) integer value. Count must be greater 0.
'The class instance must have been seeded.

Public Function NextRandomIntArr(ByVal Count As Long) As Integer()
Const cMETHODNAME As String = "NextRandomIntArr"
Dim Arr() As Integer, i As TInt, B As TIntByte, k As Long

CheckSeeded cMETHODNAME
CheckCount Count, cMETHODNAME

ReDim Arr(0 To Count - 1)
For k = 0 To Count - 1
  DoDeEncode B.Val, 0, 1: LSet i = B: Arr(k) = i.Val
Next k
NextRandomIntArr = Arr
End Function

'Returns a 0-based array of random bytes of length Count. Count must be greater 0.
'The class instance must have been seeded.

Public Function NextRandomByteArr(ByVal Count As Long) As Byte()
Const cMETHODNAME As String = "NextRandomByteArr"
Dim Arr() As Byte

CheckSeeded cMETHODNAME
CheckCount Count, cMETHODNAME

ReDim Arr(0 To Count - 1)
DoDeEncode Arr, 0, Count - 1

NextRandomByteArr = Arr
End Function


'******************************************************************************
'* Private helpers
'******************************************************************************

Private Sub DoSeed(ByRef Key() As Byte, ByVal Low As Long, ByVal KeyLen As Long)
Dim X As Long, Y As Long, T As Long

With mSeed
  For X = 0 To 255
    T = .State(X)
    Y = (Y + T + Key(X Mod KeyLen + Low)) And &HFF&
    .State(X) = .State(Y): .State(Y) = T
  Next X
End With

LSet mCurrent = mSeed
mIsNotSeeded = False

End Sub

Private Sub DoAdvance(ByVal Cycles As Long)
Dim i As Long, T As Long

With mCurrent
  For i = 1 To Cycles
    .X = (.X + 1) And &HFF&
    T = .State(.X)
    .Y = (.Y + T) And &HFF&
    .State(.X) = .State(.Y): .State(.Y) = T
  Next i
End With

End Sub

Private Sub DoDeEncode(ByRef Arr() As Byte, ByVal Low As Long, ByVal High As Long)
Dim i As Long, T As Long

With mCurrent
  For i = Low To High
    .X = (.X + 1) And &HFF&
    T = .State(.X)
    .Y = (.Y + T) And &HFF&
    .State(.X) = .State(.Y): .State(.Y) = T
    Arr(i) = Arr(i) Xor .State((T + .State(.X)) And &HFF&)
  Next i
End With

End Sub

Private Sub DoDeEncodeWithOTP(ByRef Arr() As Byte, _
                              ByRef OTP() As Byte, _
                              ByRef OTPPos As Long, _
                              ByVal OTPLen As Long, _
                              ByVal Low As Long, _
                              ByVal High As Long)
Dim i As Long, T As Long, Pos As Long

Pos = OTPPos
With mCurrent
  For i = Low To High
    .X = (.X + 1) And &HFF&
    T = .State(.X)
    .Y = (.Y + T) And &HFF&
    .State(.X) = .State(.Y): .State(.Y) = T
    Arr(i) = Arr(i) Xor .State((T + .State(.X)) And &HFF&) Xor OTP(Pos)
    Pos = (Pos + 1) Mod OTPLen
  Next i
End With
OTPPos = Pos
End Sub

Private Sub DoClear()
mIsNotSeeded = True
LSet mSeed = mTemplate
LSet mCurrent = mTemplate
End Sub

Private Sub CheckSeeded(ByRef MethodName As String)
If mIsNotSeeded Then gRaiseError MethodName, "Not seeded", 5
End Sub

Private Sub CheckCount(ByVal Count As Long, ByRef MethodName As String)
If Count < 1 Then gRaiseError MethodName, "Invalid count", 5
End Sub


'******************************************************************************
'* Class de/construction
'******************************************************************************

Private Sub Class_Initialize()
Dim i As Long
With mTemplate
  For i = 1 To 255: .State(i) = i: Next i
End With
DoClear
End Sub

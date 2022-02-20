Attribute VB_Name = "CBOR"
Option Explicit

'
' Copyright (c) 2022 Koki Takeyama
'
' Permission is hereby granted, free of charge, to any person obtaining
' a copy of this software and associated documentation files (the "Software"),
' to deal in the Software without restriction, including without limitation
' the rights to use, copy, modify, merge, publish, distribute, sublicense,
' and/or sell copies of the Software, and to permit persons to whom the
' Software is furnished to do so, subject to the following conditions:
'
' The above copyright notice and this permission notice shall be included
' in all copies or substantial portions of the Software.
'
' THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR
' IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY,
' FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE
' AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER
' LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING
' FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS
' IN THE SOFTWARE.
'

''
'' CBOR for VBA
''

'
' Conditional
'

' Integer
#If Win64 Then
#Const USE_LONGLONG = True
#End If

' Array
#Const USE_COLLECTION = True

'
' Declare
'

#If VBA7 Then
Private Declare PtrSafe Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByRef Dest As Any, ByRef Src As Any, ByVal Length As LongPtr)
#Else
'Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" _
    (ByRef Dest As Any, ByRef Src As Any, ByVal Length As Long)
#End If

'
' Types
'

Private Type IntegerT
    Value As Integer
End Type

Private Type LongT
    Value As Long
End Type

#If Win64 And USE_LONGLONG Then
Private Type LongLongT
    Value As LongLong
End Type
#End If

Private Type SingleT
    Value As Single
End Type

Private Type DoubleT
    Value As Double
End Type

Private Type Bytes2T
    Bytes(0 To 1) As Byte
End Type

Private Type Bytes4T
    Bytes(0 To 3) As Byte
End Type

Private Type Bytes8T
    Bytes(0 To 7) As Byte
End Type

''
'' CBOR for VBA - Encoding
''

Public Function GetCborBytes(Value) As Byte()
    Select Case VarType(Value)
    
    ' 0
    Case vbEmpty
        GetCborBytes = GetCborBytesFromEmpty
        
    ' 1
    Case vbNull
        GetCborBytes = GetCborBytesFromNull
        
    ' 2
    Case vbInteger
        GetCborBytes = GetCborBytesFromInt(Value)
        
    ' 3
    Case vbLong
        GetCborBytes = GetCborBytesFromInt(Value)
        
    ' 4
    Case vbSingle
        GetCborBytes = GetCborBytesFromSingle(Value)
        
    ' 5
    Case vbDouble
        GetCborBytes = GetCborBytesFromDouble(Value)
        
    ' 8
    Case vbString
        GetCborBytes = GetCborBytesFromString((Value))
        
    ' 9
    Case vbObject
        GetCborBytes = GetCborBytesFromObject(Value)
        
    ' 11
    Case vbBoolean
        GetCborBytes = GetCborBytesFromBoolean(Value)
        
    ' 14 - temporaly work around
    Case vbDecimal
        GetCborBytes = GetCborBytesFromInt(Value)
        
    ' 17
    Case vbByte
        GetCborBytes = GetCborBytesFromInt(Value)
        
    ' 20
    #If Win64 And USE_LONGLONG Then
    Case vbLongLong
        GetCborBytes = GetCborBytesFromInt(Value)
    #End If
        
    ' 8209. (17 + 8192)
    Case vbByte + vbArray
        GetCborBytes = GetCborBytesFromByteArray(Value)
        
    Case Else
        GetCborBytes = GetCborBytesFromUnknown(Value)
        
    End Select
End Function

'
' 0. Empty
'

Private Function GetCborBytesFromEmpty() As Byte()
    GetCborBytesFromEmpty = GetCborBytesFromUndefined
End Function

'
' 1. Null
'

'Private Function GetCborBytesFromNull() As Byte()
'    GetCborBytesFromNull = GetCborBytesFromNull
'End Function

'
' 2. Integer
' 3. Long
' 17. Byte
' 20. LongLong
'
Private Function GetCborBytesFromInt(Value) As Byte()
    If Value >= 0 Then
        GetCborBytesFromInt = GetCborBytesFromPosInt(Value)
    Else
        GetCborBytesFromInt = GetCborBytesFromNegInt(Value)
    End If
End Function

Private Function GetCborBytesFromPosInt(Value) As Byte()
    Select Case Value
    
    Case 0 To 23 '&H17
        GetCborBytesFromPosInt = GetCborBytesFromPosFixInt((Value))
        
    Case 24 To 255 '&H18 To &HFF
        GetCborBytesFromPosInt = GetCborBytesFromPosInt8((Value))
        
    Case 256 To 65535 '&H100 To &HFFFF&
        GetCborBytesFromPosInt = GetCborBytesFromPosInt16((Value))
        
    #If Win64 And USE_LONGLONG Then
    Case 65536 To 4294967295^ '&H10000 To &HFFFFFFFF^
        GetCborBytesFromPosInt = GetCborBytesFromPosInt32((Value))
        
    Case Else
        GetCborBytesFromPosInt = GetCborBytesFromPosInt64((Value))
        
    #Else
    Case Else
        GetCborBytesFromPosInt = GetCborBytesFromPosInt32((Value))
        
    #End If
        
    End Select
End Function

Private Function GetCborBytesFromNegInt(Value) As Byte()
    Select Case Value
    
    Case -24 To -1
        GetCborBytesFromNegInt = GetCborBytesFromNegFixInt((Value))
        
    Case -256 To -25
        GetCborBytesFromNegInt = GetCborBytesFromNegInt8((Value))
        
    Case -65536 To -257
        GetCborBytesFromNegInt = GetCborBytesFromNegInt16((Value))
        
    #If Win64 And USE_LONGLONG Then
    Case -4294967296^ To -65537
        GetCborBytesFromNegInt = GetCborBytesFromNegInt32((Value))
        
    Case Else
        GetCborBytesFromNegInt = GetCborBytesFromNegInt64((Value))
        
    #Else
    Case Else
        GetCborBytesFromNegInt = GetCborBytesFromNegInt32((Value))
        
    #End If
        
    End Select
End Function

'
' 4. Single
'
Private Function GetCborBytesFromSingle(Value) As Byte()
    GetCborBytesFromSingle = GetCborBytesFromFloat32((Value))
End Function

'
' 5. Double
'
Private Function GetCborBytesFromDouble(Value) As Byte()
    GetCborBytesFromDouble = GetCborBytesFromFloat64((Value))
End Function

'
' 8. String
'

Private Function GetCborBytesFromString(ByVal Value As String) As Byte()
    If CStr(Value) = "" Then
        GetCborBytesFromString = GetCborBytes0(&H60)
        Exit Function
    End If
    
    Dim StrBytes() As Byte
    StrBytes = GetBytesFromString(Value)
    
    Dim StrLength As Long
    StrLength = UBound(StrBytes) - LBound(StrBytes) + 1
    
    Select Case StrLength
    
    Case 1 To 23 '&H17
        GetCborBytesFromString = GetCborBytesFromFixStr(StrBytes, StrLength)
        
    Case 24 To 255 '&H18 To &HFF
        GetCborBytesFromString = GetCborBytesFromStr8(StrBytes, StrLength)
        
    Case 256 To 65535 '&H100 To &HFFFF&
        GetCborBytesFromString = GetCborBytesFromStr16(StrBytes, StrLength)
        
    '#If Win64 And USE_LONGLONG Then
    'Case 65536 To 4294967295^ '&H10000 To &HFFFFFFFF^
    '    GetCborBytesFromString = GetCborBytesFromStr32(StrBytes, StrLength)
    '
    'Case Else
    '    GetCborBytesFromString = GetCborBytesFromStr64(StrBytes, StrLength)
    '
    '#Else
    Case Else
        GetCborBytesFromString = GetCborBytesFromStr32(StrBytes, StrLength)
        
    '#End If
        
    End Select
End Function

'
' 9. Object
'

Private Function GetCborBytesFromObject(Value) As Byte()
    If Value Is Nothing Then
        GetCborBytesFromObject = GetCborBytesFromNull
        Exit Function
    End If
    
    Select Case TypeName(Value)
    
    Case "Collection"
        GetCborBytesFromObject = GetCborBytesFromCollection(Value)
        
    Case "Dictionary"
        GetCborBytesFromObject = GetCborBytesFromDictionary(Value)
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'
' 9. Object - Collection
'

Private Function GetCborBytesFromCollection(Value) As Byte()
    Select Case Value.Count
    
    Case 0
        GetCborBytesFromCollection = GetCborBytes0(&H80)
        
    Case 1 To 23 ' &H17
        GetCborBytesFromCollection = GetCborBytesFromFixArray(Value)
        
    Case 24 To 255 '&H18 To &HFF
        GetCborBytesFromCollection = GetCborBytesFromArray8(Value)
        
    Case 256 To 65535 '&H100 To &HFFFF&
        GetCborBytesFromCollection = GetCborBytesFromArray16(Value)
        
    '#If Win64 And USE_LONGLONG Then
    'Case 65536 To 4294967295^ '&H10000 To &HFFFFFFFF^
    '    GetCborBytesFromCollection = GetCborBytesFromArray32(Value)
    '
    'Case Else
    '    GetCborBytesFromCollection = GetCborBytesFromArray64(Value)
    '
    '#Else
    Case Else
        GetCborBytesFromCollection = GetCborBytesFromArray32(Value)
        
    '#End If
        
    End Select
End Function

'
' 9. Object - Dictionary
'

Private Function GetCborBytesFromDictionary(Value) As Byte()
    Select Case Value.Count
    
    Case 0
        GetCborBytesFromDictionary = GetCborBytes0(&HA0)
        
    Case 1 To 23 ' &H17
        GetCborBytesFromDictionary = GetCborBytesFromFixMap(Value)
        
    Case 24 To 255 '&H18 To &HFF
        GetCborBytesFromDictionary = GetCborBytesFromMap8(Value)
        
    Case 256 To 65535 '&H100 To &HFFFF&
        GetCborBytesFromDictionary = GetCborBytesFromMap16(Value)
        
    '#If Win64 And USE_LONGLONG Then
    'Case 65536 To 4294967295^ '&H10000 To &HFFFFFFFF^
    '    GetCborBytesFromDictionary = GetCborBytesFromMap32(Value)
    '
    'Case Else
    '    GetCborBytesFromDictionary = GetCborBytesFromMap64(Value)
    '
    '#Else
    Case Else
        GetCborBytesFromDictionary = GetCborBytesFromMap32(Value)
        
    '#End If
        
    End Select
End Function

'
' 11. Boolean
'

Private Function GetCborBytesFromBoolean(Value) As Byte()
    If Value Then
        GetCborBytesFromBoolean = GetCborBytesFromTrue
    Else
        GetCborBytesFromBoolean = GetCborBytesFromFalse
    End If
End Function

'
' 8209. (17 + 8192) Byte Array
'

Private Function GetCborBytesFromByteArray(Value) As Byte()
    Dim Length As Long
    
    On Error Resume Next
    Length = UBound(Value) - LBound(Value) + 1
    On Error GoTo 0
    
    Select Case Length
    
    Case 0
        GetCborBytesFromByteArray = GetCborBytes0(&H40)
        
    Case 1 To 23 '&H17
        GetCborBytesFromByteArray = GetCborBytesFromFixBin(Value, Length)
        
    Case 24 To 255 '&H18 To &HFF
        GetCborBytesFromByteArray = GetCborBytesFromBin8(Value, Length)
        
    Case 256 To 65535 '&H100 To &HFFFF&
        GetCborBytesFromByteArray = GetCborBytesFromBin16(Value, Length)
        
    '#If Win64 And USE_LONGLONG Then
    'Case 65536 To 4294967295^ '&H10000 To &HFFFFFFFF^
    '    GetCborBytesFromByteArray = GetCborBytesFromBin32(Value, Length)
    '
    'Case Else
    '    GetCborBytesFromByteArray = GetCborBytesFromBin64(Value, Length)
    '
    '#Else
    Case Else
        GetCborBytesFromByteArray = GetCborBytesFromBin32(Value, Length)
        
    '#End If
        
    End Select
End Function

'
' X. Unknown
'

Private Function GetCborBytesFromUnknown(Value) As Byte()
    If IsArray(Value) Then
        GetCborBytesFromUnknown = GetCborBytesFromArray(Value)
    Else
        Err.Raise 13 ' unmatched type
    End If
End Function

''
'' CBOR for VBA - Encoding - Core
''

'
' major type 0: positive integer
'

' 0x00..0x17 | unsigned integer 0x00..0x17 (0..23)

Private Function GetCborBytesFromPosFixInt(ByVal Value As Byte) As Byte()
    Debug.Assert (Value <= &H17)
    GetCborBytesFromPosFixInt = GetCborBytes0(Value)
End Function

' 0x18 | unsigned integer (one-byte uint8_t follows)

Private Function GetCborBytesFromPosInt8(ByVal Value As Byte) As Byte()
    GetCborBytesFromPosInt8 = _
        GetCborBytes1(&H18, GetBytesFromUInt8(Value))
End Function

' 0x19 | unsigned integer (two-byte uint16_t follows)

Private Function GetCborBytesFromPosInt16(ByVal Value As Long) As Byte()
    Debug.Assert ((Value >= 0) And (Value <= &HFFFF&))
    GetCborBytesFromPosInt16 = _
        GetCborBytes1(&H19, GetBytesFromUInt16(Value, True))
End Function

' 0x1a | unsigned integer (four-byte uint32_t follows)

Private Function GetCborBytesFromPosInt32(ByVal Value) As Byte()
    #If Win64 Then
    Debug.Assert ((Value >= 0) And (Value <= &HFFFFFFFF^))
    #Else
    Debug.Assert (Value >= 0)
    #End If
    GetCborBytesFromPosInt32 = _
        GetCborBytes1(&H1A, GetBytesFromUInt32(Value, True))
End Function

' 0x1b | unsigned integer (eight-byte uint64_t follows)

Private Function GetCborBytesFromPosInt64(ByVal Value) As Byte()
    Debug.Assert (Value >= 0)
    GetCborBytesFromPosInt64 = _
        GetCborBytes1(&H1B, GetBytesFromUInt64(Value, True))
End Function

'
' major type 1: negative integer
'

' 0x20..0x37 | negative integer -1-0x00..-1-0x17 (-1..-24)

Private Function GetCborBytesFromNegFixInt(ByVal Value As Integer) As Byte()
    Debug.Assert ((Value >= -24) And (Value <= -1))
    GetCborBytesFromNegFixInt = GetCborBytes0(&H20 Or CByte(Abs(Value + 1)))
End Function

' 0x38 | negative integer -1-n (one-byte uint8_t for n follows)

Private Function GetCborBytesFromNegInt8(ByVal Value As Integer) As Byte()
    Debug.Assert ((Value >= -256) And (Value <= -1))
    GetCborBytesFromNegInt8 = _
        GetCborBytes1(&H38, GetBytesFromUInt8(CByte(Abs(Value + 1))))
End Function

' 0x39 | negative integer -1-n (two-byte uint16_t for n follows)

Private Function GetCborBytesFromNegInt16(ByVal Value As Long) As Byte()
    Debug.Assert ((Value >= -65536) And (Value <= -1))
    GetCborBytesFromNegInt16 = _
        GetCborBytes1(&H39, GetBytesFromUInt16(CLng(Abs(Value + 1)), True))
End Function

' 0x3a | negative integer -1-n (four-byte uint32_t for n follows)

#If Win64 And USE_LONGLONG Then

Private Function GetCborBytesFromNegInt32(ByVal Value As LongLong) As Byte()
    Debug.Assert ((Value >= -4294967296^) And (Value <= -1))
    GetCborBytesFromNegInt32 = _
        GetCborBytes1(&H3A, GetBytesFromUInt32(CLngLng(Abs(Value + 1)), True))
End Function

#Else

Private Function GetCborBytesFromNegInt32(ByVal Value) As Byte()
    Debug.Assert (Value <= -1)
    GetCborBytesFromNegInt32 = _
        GetCborBytes1(&H3A, GetBytesFromUInt32(Abs(Value + 1), True))
End Function

#End If

' 0x3b | negative integer -1-n (eight-byte uint64_t for n follows)

Private Function GetCborBytesFromNegInt64(ByVal Value) As Byte()
    Debug.Assert (Value <= -1)
    GetCborBytesFromNegInt64 = _
        GetCborBytes1(&H3B, GetBytesFromUInt64(Abs(Value + 1), True))
End Function

'
' major type 2: byte string
'

' 0x40..0x57 | byte string (0x00..0x17 bytes follow)

Private Function GetCborBytesFromFixBin( _
    BinBytes, ByVal BinLength As Byte) As Byte()
    'BinBytes() As Byte, ByVal BinLength As Byte) As Byte()
    
    Debug.Assert ((BinLength > 0) And (BinLength <= &H17))
    
    GetCborBytesFromFixBin = GetCborBytes1(&H40 Or BinLength, BinBytes)
End Function

' 0x58 | byte string (one-byte uint8_t for n, and then n bytes follow)

Private Function GetCborBytesFromBin8( _
    BinBytes, ByVal BinLength As Byte) As Byte()
    'BinBytes() As Byte, ByVal BinLength As Byte) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetCborBytesFromBin8 = _
        GetCborBytes2(&H58, GetBytesFromUInt8(BinLength), BinBytes)
End Function

' 0x59 | byte string (two-byte uint16_t for n, and then n bytes follow)

Private Function GetCborBytesFromBin16( _
    BinBytes, ByVal BinLength As Long) As Byte()
    'BinBytes() As Byte, ByVal BinLength As Long) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetCborBytesFromBin16 = _
        GetCborBytes2(&H59, GetBytesFromUInt16(BinLength, True), BinBytes)
End Function

' 0x5a | byte string (four-byte uint32_t for n, and then n bytes follow)

Private Function GetCborBytesFromBin32( _
    BinBytes, ByVal BinLength As Long) As Byte()
    'BinBytes() As Byte, ByVal BinLength As Long) As Byte()
    
    Debug.Assert (BinLength > 0)
    
    GetCborBytesFromBin32 = _
        GetCborBytes2(&H5A, GetBytesFromUInt32(BinLength, True), BinBytes)
End Function

' 0x5b | byte string (eight-byte uint64_t for n, and then n bytes follow)

'Private Function GetCborBytesFromBin64( _
'    BinBytes, ByVal BinLength As LongLong) As Byte()
'    'BinBytes() As Byte, ByVal BinLength As LongLong) As Byte()
'
'    Debug.Assert (BinLength > 0)
'
'    GetCborBytesFromBin64 = _
'        GetCborBytes2(&H5B, GetBytesFromUInt64(BinLength, True), BinBytes)
'End Function

' 0x5f | byte string, byte strings follow, terminated by "break"

'Private Function GetCborBytesFromBinBreak( _
'    BinBytes, ByVal BinLength As Long) As Byte()
'    'BinBytes() As Byte, ByVal BinLength As Long) As Byte()
'
'    Debug.Assert (BinLength > 0)
'
'    GetCborBytesFromBinBreak = _
'        GetCborBytes2(&H5F, BinBytes, GetBytesFromUInt8(&HFF))
'End Function

'
' major type 3: text string
'

' 0x60..0x77 | UTF-8 string (0x00..0x17 bytes follow)

Private Function GetCborBytesFromFixStr( _
    StrBytes() As Byte, ByVal StrLength As Byte) As Byte()
    
    Debug.Assert ((StrLength > 0) And (StrLength <= &H17))
    
    GetCborBytesFromFixStr = GetCborBytes1(&H60 Or StrLength, StrBytes)
End Function

' 0x78 | UTF-8 string (one-byte uint8_t for n, and then n bytes follow)

Private Function GetCborBytesFromStr8( _
    StrBytes() As Byte, ByVal StrLength As Byte) As Byte()
    
    Debug.Assert (StrLength > 0)
    
    GetCborBytesFromStr8 = _
        GetCborBytes2(&H78, GetBytesFromUInt8(StrLength), StrBytes)
End Function

' 0x79 | UTF-8 string (two-byte uint16_t for n, and then n bytes follow)

Private Function GetCborBytesFromStr16( _
    StrBytes() As Byte, ByVal StrLength As Long) As Byte()
    
    Debug.Assert (StrLength > 0)
    
    GetCborBytesFromStr16 = _
        GetCborBytes2(&H79, GetBytesFromUInt16(StrLength, True), StrBytes)
End Function

' 0x7a | UTF-8 string (four-byte uint32_t for n, and then n bytes follow)

Private Function GetCborBytesFromStr32( _
    StrBytes() As Byte, ByVal StrLength) As Byte()
    
    Debug.Assert (StrLength > 0)
    
    GetCborBytesFromStr32 = _
        GetCborBytes2(&H7A, GetBytesFromUInt32(StrLength, True), StrBytes)
End Function

' 0x7b | UTF-8 string (eight-byte uint64_t for n, and then n bytes follow)

'Private Function GetCborBytesFromStr64( _
'    StrBytes() As Byte, ByVal StrLength) As Byte()
'
'    Debug.Assert (StrLength > 0)
'
'    GetCborBytesFromStr64 = _
'        GetCborBytes2(&H7B, GetBytesFromUInt64(StrLength, True), StrBytes)
'End Function

' 0x7f | UTF-8 string, UTF-8 strings follow, terminated by "break"

'Private Function GetCborBytesFromStrBreak( _
'    StrBytes() As Byte, ByVal StrLength) As Byte()
'
'    Debug.Assert (StrLength > 0)
'
'    GetCborBytesFromStrBreak = _
'        GetCborBytes2(&H7F, StrBytes, GetBytesFromUInt8(&HFF))
'End Function

'
' major type 4: array
'

' 0x80..0x97 | array (0x00..0x17 data items follow)

Private Function GetCborBytesFromFixArray(Value) As Byte()
    Dim Count As Long
    
    If IsArray(Value) Then
        Count = UBound(Value) - LBound(Value) + 1
    ElseIf TypeName(Value) = "Collection" Then
        Count = Value.Count
    Else
        Err.Raise 13 ' unmatched type
    End If
    
    Debug.Assert ((Count > 0) And (Count <= &H17))
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0)
    CborBytes(0) = &H80 Or Count
    
    AddCborBytesFromArray CborBytes, Value
    
    GetCborBytesFromFixArray = CborBytes
End Function

' 0x98 | array (one-byte uint8_t for n, and then n data items follow)

Private Function GetCborBytesFromArray8(Value) As Byte()
    Dim Count As Long
    
    If IsArray(Value) Then
        Count = UBound(Value) - LBound(Value) + 1
    ElseIf TypeName(Value) = "Collection" Then
        Count = Value.Count
    Else
        Err.Raise 13 ' unmatched type
    End If
    
    Debug.Assert ((Count > 0) And (Count <= &HFF))
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0)
    CborBytes(0) = &H98
    AddBytes CborBytes, GetBytesFromUInt8(Count)
    
    AddCborBytesFromArray CborBytes, Value
    
    GetCborBytesFromArray8 = CborBytes
End Function

' 0x99 | array (two-byte uint16_t for n, and then n data items follow)

Private Function GetCborBytesFromArray16(Value) As Byte()
    Dim Count As Long
    
    If IsArray(Value) Then
        Count = UBound(Value) - LBound(Value) + 1
    ElseIf TypeName(Value) = "Collection" Then
        Count = Value.Count
    Else
        Err.Raise 13 ' unmatched type
    End If
    
    Debug.Assert ((Count > 0) And (Count <= &HFFFF&))
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0)
    CborBytes(0) = &H99
    AddBytes CborBytes, GetBytesFromUInt16(Count, True)
    
    AddCborBytesFromArray CborBytes, Value
    
    GetCborBytesFromArray16 = CborBytes
End Function

' 0x9a | array (four-byte uint32_t for n, and then n data items follow)

Private Function GetCborBytesFromArray32(Value) As Byte()
    Dim Count As Long
    
    If IsArray(Value) Then
        Count = UBound(Value) - LBound(Value) + 1
    ElseIf TypeName(Value) = "Collection" Then
        Count = Value.Count
    Else
        Err.Raise 13 ' unmatched type
    End If
    
    Debug.Assert (Count > 0)
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0)
    CborBytes(0) = &H9A
    AddBytes CborBytes, GetBytesFromUInt32(Count, True)
    
    AddCborBytesFromArray CborBytes, Value
    
    GetCborBytesFromArray32 = CborBytes
End Function

' 0x9b | array (eight-byte uint64_t for n, and then n data items follow)

'Private Function GetCborBytesFromArray64(Value) As Byte()
'    Dim Count As Long
'
'    If IsArray(Value) Then
'        Count = UBound(Value) - LBound(Value) + 1
'    ElseIf TypeName(Value) = "Collection" Then
'        Count = Value.Count
'    Else
'        Err.Raise 13 ' unmatched type
'    End If
'
'    Debug.Assert (Count > 0)
'
'    Dim CborBytes() As Byte
'    ReDim CborBytes(0)
'    CborBytes(0) = &H9B
'    AddBytes CborBytes, GetBytesFromUInt32(Count, True)
'
'    AddCborBytesFromArray CborBytes, Value
'
'    GetCborBytesFromArray64 = CborBytes
'End Function

' 0x9f | array, data items follow, terminated by "break"

'Private Function GetCborBytesFromArrayBreak(Value) As Byte()
'    Dim CborBytes() As Byte
'    ReDim CborBytes(0)
'    CborBytes(0) = &H9F
'    AddCborBytesFromArray CborBytes, Value
'
'    AddBytes CborBytes, GetBytesFromUInt8(&HFF)
'
'    GetCborBytesFromArrayBreak = CborBytes
'End Function

'
' major type 5: map
'

' 0xa0..0xb7 | map (0x00..0x17 pairs of data items follow)

Private Function GetCborBytesFromFixMap(Value) As Byte()
    Debug.Assert (TypeName(Value) = "Dictionary")
    Debug.Assert ((Value.Count > 0) And (Value.Count <= &H17))
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0)
    CborBytes(0) = &HA0 Or Value.Count
    
    AddCborBytesFromMap CborBytes, Value
    
    GetCborBytesFromFixMap = CborBytes
End Function

' 0xb8 | map (one-byte uint8_t for n, and then n pairs of data items follow)

Private Function GetCborBytesFromMap8(Value) As Byte()
    Debug.Assert (TypeName(Value) = "Dictionary")
    Debug.Assert ((Value.Count > 0) And (Value.Count <= &HFF))
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0)
    CborBytes(0) = &HB8
    AddBytes CborBytes, GetBytesFromUInt8(Value.Count)
    
    AddCborBytesFromMap CborBytes, Value
    
    GetCborBytesFromMap8 = CborBytes
End Function

' 0xb9 | map (two-byte uint16_t for n, and then n pairs of data items follow)

Private Function GetCborBytesFromMap16(Value) As Byte()
    Debug.Assert (TypeName(Value) = "Dictionary")
    Debug.Assert ((Value.Count > 0) And (Value.Count <= &HFFFF&))
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0)
    CborBytes(0) = &HB9
    AddBytes CborBytes, GetBytesFromUInt16(Value.Count, True)
    
    AddCborBytesFromMap CborBytes, Value
    
    GetCborBytesFromMap16 = CborBytes
End Function

' 0xba | map (four-byte uint32_t for n, and then n pairs of data items follow)

Private Function GetCborBytesFromMap32(Value) As Byte()
    Debug.Assert (TypeName(Value) = "Dictionary")
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0)
    CborBytes(0) = &HBA
    AddBytes CborBytes, GetBytesFromUInt32(Value.Count, True)
    
    AddCborBytesFromMap CborBytes, Value
    
    GetCborBytesFromMap32 = CborBytes
End Function

' 0xbb | map (eight-byte uint64_t for n, and then n pairs of data items follow)

'Private Function GetCborBytesFromMap64( _
'    Value, Optional Optimize As Boolean) As Byte()
'
'    Dim CborBytes() As Byte
'    ReDim CborBytes(0)
'    CborBytes(0) = &HBB
'    AddBytes CborBytes, GetBytesFromUInt64(Value.Count, True)
'
'    AddCborBytesFromMap CborBytes, Value
'
'    GetCborBytesFromMap64 = CborBytes
'End Function

' 0xbf | map, pairs of data items follow, terminated by "break"

'Private Function GetCborBytesFromMapBreak(Value) As Byte()
'    Debug.Assert (TypeName(Value) = "Dictionary")
'
'    Dim CborBytes() As Byte
'    ReDim CborBytes(0)
'    CborBytes(0) = &HBF
'    AddCborBytesFromMap CborBytes, Value
'
'    AddBytes CborBytes, GetBytesFromUInt8(&HFF)
'
'    GetCborBytesFromMapBreak = CborBytes
'End Function

'
' major type 6: tag
'

' 0xc0 | text-based date/time (data item follows; see Section 3.4.1)
' 0xc1 | epoch-based date/time (data item follows; see Section 3.4.2)
' 0xc2 | unsigned bignum (data item "byte string" follows)
' 0xc3 | negative bignum (data item "byte string" follows)
' 0xc4 | decimal Fraction (data item "array" follows; see Section 3.4.4)
' 0xc5 | bigfloat (data item "array" follows; see Section 3.4.4)
' 0xc6..0xd4 | (tag)
' 0xd5..0xd7 | expected conversion (data item follows; see Section 3.4.5.2)
' 0xd8..0xdb | (more tags; 1/2/4/8 bytes of tag number and then a data item follow)

'
' major type 7: simple/float
'

' 0xf4 | false

Private Function GetCborBytesFromFalse() As Byte()
    GetCborBytesFromFalse = GetCborBytes0(&HF4)
End Function

' 0xf5 | true

Private Function GetCborBytesFromTrue() As Byte()
    GetCborBytesFromTrue = GetCborBytes0(&HF5)
End Function

' 0xf6 | null

Private Function GetCborBytesFromNull() As Byte()
    GetCborBytesFromNull = GetCborBytes0(&HF6)
End Function

' 0xf7 | undefined

Private Function GetCborBytesFromUndefined() As Byte()
    GetCborBytesFromUndefined = GetCborBytes0(&HF7)
End Function

' 0xf9 | half-precision float (two-byte IEEE 754)

'Private Function GetCborBytesFromFloat16(ByVal Value As Single) As Byte()
'    GetCborBytesFromFloat16 = _
'        GetCborBytes1(&HF9, GetBytesFromFloat16(Value, True))
'End Function

' 0xfa | single-precision float (four-byte IEEE 754)

Private Function GetCborBytesFromFloat32(ByVal Value As Single) As Byte()
    GetCborBytesFromFloat32 = _
        GetCborBytes1(&HFA, GetBytesFromFloat32(Value, True))
End Function

' 0xfb | double-precision float (eight-byte IEEE 754)

Private Function GetCborBytesFromFloat64(ByVal Value As Double) As Byte()
    GetCborBytesFromFloat64 = _
        GetCborBytes1(&HFB, GetBytesFromFloat64(Value, True))
End Function

' 0xff | "break" stop code

''
'' CBOR for VBA - Encoding - Array Helper
''

Private Sub AddCborBytesFromArray(CborBytes() As Byte, Value)
    Dim LB As Long
    Dim UB As Long
    
    If IsArray(Value) Then
        LB = LBound(Value)
        UB = UBound(Value)
        
    ElseIf TypeName(Value) = "Collection" Then
        LB = 1
        UB = Value.Count
        
    Else
        Err.Raise 13 ' unmatched type
        
    End If
    
    Dim Index As Long
    For Index = LB To UB
        AddBytes CborBytes, GetCborBytes(Value(Index))
    Next
End Sub

''
'' CBOR for VBA - Encoding - Map Helper
''

Private Sub AddCborBytesFromMap(CborBytes() As Byte, Value)
    Debug.Assert (TypeName(Value) = "Dictionary")
    
    Dim Keys
    Keys = Value.Keys
    
    Dim Index As Long
    For Index = LBound(Keys) To UBound(Keys)
        AddBytes CborBytes, GetCborBytes(Keys(Index))
        AddBytes CborBytes, GetCborBytes(Value.Item(Keys(Index)))
    Next
End Sub

''
'' CBOR for VBA - Encoding - Formatter
''

Private Function GetCborBytes0(HeaderValue As Byte) As Byte()
    Dim CborBytes(0) As Byte
    CborBytes(0) = HeaderValue
    GetCborBytes0 = CborBytes
End Function

Private Function GetCborBytes1( _
    HeaderValue As Byte, SrcBytes) As Byte()
    'HeaderValue As Byte, SrcBytes() As Byte) As Byte()
    
    Dim SrcLB As Long
    Dim SrcUB As Long
    SrcLB = LBound(SrcBytes)
    SrcUB = UBound(SrcBytes)
    
    Dim SrcLen As Long
    SrcLen = SrcUB - SrcLB + 1
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0 To SrcLen)
    CborBytes(0) = HeaderValue
    
    CopyBytes CborBytes, 1, SrcBytes, SrcLB, SrcLen
    
    GetCborBytes1 = CborBytes
End Function

Private Function GetCborBytes2( _
    HeaderValue As Byte, SrcBytes1, SrcBytes2) As Byte()
    'HeaderValue As Byte, SrcBytes1() As Byte, SrcBytes2() As Byte) As Byte()
    
    Dim SrcLB1 As Long
    Dim SrcUB1 As Long
    SrcLB1 = LBound(SrcBytes1)
    SrcUB1 = UBound(SrcBytes1)
    
    Dim SrcLen1 As Long
    SrcLen1 = SrcUB1 - SrcLB1 + 1
    
    Dim SrcLB2 As Long
    Dim SrcUB2 As Long
    SrcLB2 = LBound(SrcBytes2)
    SrcUB2 = UBound(SrcBytes2)
    
    Dim SrcLen2 As Long
    SrcLen2 = SrcUB2 - SrcLB2 + 1
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0 To SrcLen1 + SrcLen2)
    CborBytes(0) = HeaderValue
    
    CopyBytes CborBytes, 1, SrcBytes1, SrcLB1, SrcLen1
    CopyBytes CborBytes, 1 + SrcLen1, SrcBytes2, SrcLB2, SrcLen2
    
    GetCborBytes2 = CborBytes
End Function

Private Function GetCborBytes3(HeaderValue As Byte, _
    SrcBytes1, SrcBytes2, SrcBytes3) As Byte()
    
    Dim SrcLB1 As Long
    Dim SrcUB1 As Long
    SrcLB1 = LBound(SrcBytes1)
    SrcUB1 = UBound(SrcBytes1)
    
    Dim SrcLen1 As Long
    SrcLen1 = SrcUB1 - SrcLB1 + 1
    
    Dim SrcLB2 As Long
    Dim SrcUB2 As Long
    SrcLB2 = LBound(SrcBytes2)
    SrcUB2 = UBound(SrcBytes2)
    
    Dim SrcLen2 As Long
    SrcLen2 = SrcUB2 - SrcLB2 + 1
    
    Dim SrcLB3 As Long
    Dim SrcUB3 As Long
    SrcLB3 = LBound(SrcBytes3)
    SrcUB3 = UBound(SrcBytes3)
    
    Dim SrcLen3 As Long
    SrcLen3 = SrcUB3 - SrcLB3 + 1
    
    Dim CborBytes() As Byte
    ReDim CborBytes(0 To SrcLen1 + SrcLen2 + SrcLen3)
    Bytes(0) = HeaderValue
    
    CopyBytes CborBytes, 1, SrcBytes1, SrcLB1, SrcLen1
    
    CopyBytes CborBytes, 1 + SrcLen1, SrcBytes2, SrcLB2, SrcLen2
    
    CopyBytes CborBytes, 1 + SrcLen1 + SrcLen2, SrcBytes3, SrcLB3, SrcLen3
    
    GetCborBytes3 = CborBytes
End Function

''
'' CBOR for VBA - Encoding - Bytes Operator
''

Private Sub AddBytes(DstBytes() As Byte, SrcBytes() As Byte)
    Dim DstLB As Long
    Dim DstUB As Long
    DstLB = LBound(DstBytes)
    DstUB = UBound(DstBytes)
    
    Dim SrcLB As Long
    Dim SrcUB As Long
    Dim SrcLen As Long
    SrcLB = LBound(SrcBytes)
    SrcUB = UBound(SrcBytes)
    SrcLen = SrcUB - SrcLB + 1
    
    ReDim Preserve DstBytes(DstLB To DstUB + SrcLen)
    CopyBytes DstBytes, DstUB + 1, SrcBytes, SrcLB, SrcLen
End Sub

Private Sub CopyBytes( _
    DstBytes() As Byte, DstIndex As Long, _
    SrcBytes, SrcIndex As Long, ByVal Length As Long)
    'SrcBytes() As Byte, SrcIndex As Long, ByVal Length As Long)
    
    Dim Offset As Long
    For Offset = 0 To Length - 1
        DstBytes(DstIndex + Offset) = SrcBytes(SrcIndex + Offset)
    Next
End Sub

Private Sub ReverseBytes( _
    ByRef Bytes() As Byte, _
    Optional Index As Long, _
    Optional ByVal Length As Long)
    
    Dim UB As Long
    
    If Length = 0 Then
        UB = UBound(Bytes)
        Length = UB - Index + 1
    Else
        UB = Index + Length - 1
    End If
    
    Dim Offset As Long
    For Offset = 0 To (Length \ 2) - 1
        Dim Temp As Byte
        Temp = Bytes(Index + Offset)
        Bytes(Index + Offset) = Bytes(UB - Offset)
        Bytes(UB - Offset) = Temp
    Next
End Sub

''
'' CBOR for VBA - Encoding - Converter
''

'
' 0x18. UInt8 - a 8-bit unsigned integer
'

Private Function GetBytesFromUInt8(ByVal Value As Byte) As Byte()
    Dim Bytes(0) As Byte
    Bytes(0) = Value
    GetBytesFromUInt8 = Bytes
End Function

'
' 0x19. UInt16 - a 16-bit unsigned integer
'

Private Function GetBytesFromUInt16( _
    ByVal Value As Long, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert ((Value >= 0) And (Value <= &HFFFF&))
    
    Dim Bytes4() As Byte
    Bytes4 = GetBytesFromLong(Value, BigEndian)
    
    Dim Bytes(0 To 1) As Byte
    If BigEndian Then
        CopyBytes Bytes, 0, Bytes4, 2, 2
    Else
        CopyBytes Bytes, 0, Bytes4, 0, 2
    End If
    
    GetBytesFromUInt16 = Bytes
End Function

'
' 0x1a. UInt32 - a 32-bit unsigned integer
'

#If Win64 And USE_LONGLONG Then

Private Function GetBytesFromUInt32( _
    ByVal Value As LongLong, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert ((Value >= 0) And (Value <= &HFFFFFFFF^))
    
    Dim Bytes8() As Byte
    Bytes8 = GetBytesFromLongLong(Value, BigEndian)
    
    Dim Bytes(0 To 3) As Byte
    If BigEndian Then
        CopyBytes Bytes, 0, Bytes8, 4, 4
    Else
        CopyBytes Bytes, 0, Bytes8, 0, 4
    End If
    
    GetBytesFromUInt32 = Bytes
End Function

#Else

Private Function GetBytesFromUInt32( _
    ByVal Value As Variant, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert (VarType(Value) = vbDecimal)
    Debug.Assert ((Value >= 0) And (Value <= CDec("4294967295")))
    
    Dim Bytes14() As Byte
    Bytes14 = GetBytesFromDecimal(Value, BigEndian)
    
    Dim Bytes(0 To 3) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        CopyBytes Bytes, 0, Bytes14, 10, 4
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        CopyBytes Bytes, 0, Bytes14, 0, 4
    End If
    
    GetBytesFromUInt32 = Bytes
End Function

#End If

'
' 0x1b. UInt64 - a 64-bit unsigned integer
'

Private Function GetBytesFromUInt64( _
    ByVal Value As Variant, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert (VarType(Value) = vbDecimal)
    Debug.Assert ((Value >= 0) And (Value <= CDec("18446744073709551615")))
    
    Dim Bytes14() As Byte
    Bytes14 = GetBytesFromDecimal(Value, BigEndian)
    
    Dim Bytes(0 To 7) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        CopyBytes Bytes, 0, Bytes14, 6, 8
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        CopyBytes Bytes, 0, Bytes14, 0, 8
    End If
    
    GetBytesFromUInt64 = Bytes
End Function

'
' 0xf9. Float16 - an IEEE 754 half precision floating point number
'

'Private Function GetBytesFromFloat16( _
'    ByVal Value As Single, Optional BigEndian As Boolean) As Byte()
'
'    GetBytesFromFloat16 = GetBytesFromSingle(Value, BigEndian)
'End Function

'
' 0xfa. Float32 - an IEEE 754 single precision floating point number
'

Private Function GetBytesFromFloat32( _
    ByVal Value As Single, Optional BigEndian As Boolean) As Byte()
    
    GetBytesFromFloat32 = GetBytesFromSingle(Value, BigEndian)
End Function

'
' 0xfb. Float64 - an IEEE 754 double precision floating point number
'

Private Function GetBytesFromFloat64( _
    ByVal Value As Double, Optional BigEndian As Boolean) As Byte()
    
    GetBytesFromFloat64 = GetBytesFromDouble(Value, BigEndian)
End Function

'
' 2. Integer - a 16-bit signed integer
'

Private Function GetBytesFromInteger( _
    ByVal Value As Integer, Optional BigEndian As Boolean) As Byte()
    
    Dim I As IntegerT
    I.Value = Value
    
    Dim B2 As Bytes2T
    LSet B2 = I
    
    If BigEndian Then
        ReverseBytes B2.Bytes
    End If
    
    GetBytesFromInteger = B2.Bytes
End Function

'
' 3. Long - a 32-bit signed integer
'

Private Function GetBytesFromLong( _
    ByVal Value As Long, Optional BigEndian As Boolean) As Byte()
    
    Dim L As LongT
    L.Value = Value
    
    Dim B4 As Bytes4T
    LSet B4 = L
    
    If BigEndian Then
        ReverseBytes B4.Bytes
    End If
    
    GetBytesFromLong = B4.Bytes
End Function

'
' 4. Single - an IEEE 754 single precision floating point number
'

Private Function GetBytesFromSingle( _
    ByVal Value As Single, Optional BigEndian As Boolean) As Byte()
    
    Dim S As SingleT
    S.Value = Value
    
    Dim B4 As Bytes4T
    LSet B4 = S
    
    If BigEndian Then
        ReverseBytes B4.Bytes
    End If
    
    GetBytesFromSingle = B4.Bytes
End Function

'
' 5. Double - an IEEE 754 double precision floating point number
'

Private Function GetBytesFromDouble( _
    ByVal Value As Double, Optional BigEndian As Boolean) As Byte()
    
    Dim D As DoubleT
    D.Value = Value
    
    Dim B8 As Bytes8T
    LSet B8 = D
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    GetBytesFromDouble = B8.Bytes
End Function

'
' 8. String
'

Private Function GetBytesFromString( _
    ByVal Value As String, Optional Charset As String = "utf-8") As Byte()
    
    Debug.Assert (Value <> "")
    
    Static ADODBStream As Object
    
    If ADODBStream Is Nothing Then
        Set ADODBStream = CreateObject("ADODB.Stream")
    End If
    
    With ADODBStream
        .Open
        
        .Type = 2 'ADODB.adTypeText
        .Charset = Charset
        .WriteText Value
        
        .Position = 0
        .Type = 1 'ADODB.adTypeBinary
        If Charset = "utf-8" Then
            .Position = 3 ' avoid BOM
        End If
        GetBytesFromString = .Read
        
        .Close
    End With
End Function

'
' 14. Decimal - a 96-bit unsigned integer
'with 8-bit scaling factor and 8-bit sign factor
'

Private Function GetBytesFromDecimal( _
    ByVal Value As Variant, Optional BigEndian As Boolean) As Byte()
    
    Debug.Assert (VarType(Value) = vbDecimal)
    
    ' BytesRaw:
    ' vartype         - 2 bytes
    ' scale           - 1 byte
    ' sign            - 1 byte
    ' data high bytes - 4 bytes
    ' data low bytes  - 8 bytes
    Dim BytesRaw() As Byte
    ReDim BytesRaw(0 To 15)
    CopyMemory ByVal VarPtr(BytesRaw(0)), ByVal VarPtr(Value), 16
    
    ' BytesX:
    ' scale           - 1 byte
    ' sign            - 1 byte
    ' data high bytes - 4 bytes
    ' data low bytes  - 8 bytes
    Dim BytesX(0 To 13) As Byte
    CopyBytes BytesX, 0, BytesRaw, 2, 14
    
    ' Bytes:
    ' data low bytes  - 8 bytes
    ' data high bytes - 4 bytes
    ' scale           - 1 byte
    ' sign            - 1 byte
    Dim Bytes(0 To 13) As Byte
    
    ' data low bytes
    CopyBytes Bytes, 0, BytesX, 6, 8
    
    ' data high bytes
    CopyBytes Bytes, 8, BytesX, 2, 4
    
    ' scale
    Bytes(12) = BytesX(0)
    
    ' sign
    Bytes(13) = BytesX(1)
    
    If BigEndian Then
        ReverseBytes Bytes
    End If
    
    GetBytesFromDecimal = Bytes
End Function

'
' 20. LongLong - a 64-bit signed integer
'

#If Win64 And USE_LONGLONG Then

Private Function GetBytesFromLongLong( _
    ByVal Value As LongLong, Optional BigEndian As Boolean) As Byte()
    
    Dim LL As LongLongT
    LL.Value = Value
    
    Dim B8 As Bytes8T
    LSet B8 = LL
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    GetBytesFromLongLong = B8.Bytes
End Function

#End If

''
'' CBOR for VBA - Decoding
''

Public Function GetCborLength( _
    CborBytes() As Byte, Optional Index As Long) As Long
    
    Dim ItemCount As Long
    Dim ItemLength As Long
    
    Select Case CborBytes(Index)
    
    '
    ' major type 0: positive integer
    '
    
    ' 0x00..0x17 | unsigned integer 0x00..0x17 (0..23)
    Case &H0 To &H17
        GetCborLength = 1
        
    ' 0x18 | unsigned integer (one-byte uint8_t follows)
    Case &H18
        GetCborLength = 1 + 1
        
    ' 0x19 | unsigned integer (two-byte uint16_t follows)
    Case &H19
        GetCborLength = 1 + 2
        
    ' 0x1a | unsigned integer (four-byte uint32_t follows)
    Case &H1A
        GetCborLength = 1 + 4
        
    ' 0x1b | unsigned integer (eight-byte uint64_t follows)
    Case &H1B
        GetCborLength = 1 + 8
        
    '
    ' major type 1: negative integer
    '
    
    ' 0x20..0x37 | negative integer -1-0x00..-1-0x17 (-1..-24)
    Case &H20 To &H37
        GetCborLength = 1
        
    ' 0x38 | negative integer -1-n (one-byte uint8_t for n follows)
    Case &H38
        GetCborLength = 1 + 1
        
    ' 0x39 | negative integer -1-n (two-byte uint16_t for n follows)
    Case &H39
        GetCborLength = 1 + 2
        
    ' 0x3a | negative integer -1-n (four-byte uint32_t for n follows)
    Case &H3A
        GetCborLength = 1 + 4
        
    ' 0x3b | negative integer -1-n (eight-byte uint64_t for n follows)
    Case &H3B
        GetCborLength = 1 + 8
        
    '
    ' major type 2: byte string
    '
    
    ' 0x40..0x57 | byte string (0x00..0x17 bytes follow)
    Case &H40 To &H57
        ItemLength = (CborBytes(Index) And &H1F)
        GetCborLength = 1 + ItemLength
        
    ' 0x58 | byte string (one-byte uint8_t for n, and then n bytes follow)
    Case &H58
        ItemLength = CborBytes(Index + 1)
        GetCborLength = 1 + 1 + ItemLength
        
    ' 0x59 | byte string (two-byte uint16_t for n, and then n bytes follow)
    Case &H59
        ItemLength = GetUInt16FromBytes(CborBytes, Index + 1, True)
        GetCborLength = 1 + 2 + ItemLength
        
    ' 0x5a | byte string (four-byte uint32_t for n, and then n bytes follow)
    Case &H5A
        ItemLength = CLng(GetUInt32FromBytes(CborBytes, Index + 1, True))
        GetCborLength = 1 + 4 + ItemLength
        
    ' 0x5b | byte string (eight-byte uint64_t for n, and then n bytes follow)
    'Case &H5B
    '    ItemLength = CLng(GetUInt64FromBytes(CborBytes, Index + 1, True))
    '    GetCborLength = 1 + 8 + ItemLength
        
    ' 0x5f | byte string, byte strings follow, terminated by "break"
    'Case &H5F
    '    ItemLength = _
    '        GetBreakIndexFromCborBytes(CborBytes, Index + 1) - (Index + 1)
    '    GetCborLength = 1 + ItemLength + 1
        
    '
    ' major type 3: text string
    '
    
    ' 0x60..0x77 | UTF-8 string (0x00..0x17 bytes follow)
    Case &H60 To &H77
        ItemLength = (CborBytes(Index) And &H1F)
        GetCborLength = 1 + ItemLength
        
    ' 0x78 | UTF-8 string (one-byte uint8_t for n, and then n bytes follow)
    Case &H78
        ItemLength = CborBytes(Index + 1)
        GetCborLength = 1 + 1 + ItemLength
        
    ' 0x79 | UTF-8 string (two-byte uint16_t for n, and then n bytes follow)
    Case &H79
        ItemLength = GetUInt16FromBytes(CborBytes, Index + 1, True)
        GetCborLength = 1 + 2 + ItemLength
        
    ' 0x7a | UTF-8 string (four-byte uint32_t for n, and then n bytes follow)
    Case &H7A
        ItemLength = CLng(GetUInt32FromBytes(CborBytes, Index + 1, True))
        GetCborLength = 1 + 4 + ItemLength
        
    ' 0x7b | UTF-8 string (eight-byte uint64_t for n, and then n bytes follow)
    'Case &H7B
    '    ItemLength = CLng(GetUInt64FromBytes(CborBytes, Index + 1, True))
    '    GetCborLength = 1 + 8 + ItemLength
        
    ' 0x7f | UTF-8 string, UTF-8 strings follow, terminated by "break"
    'Case &H7F
    '    ItemLength = _
    '        GetBreakIndexFromCborBytes(CborBytes, Index + 1) - (Index + 1)
    '    GetCborLength = 1 + ItemLength + 1
        
    '
    ' major type 4: array
    '
    
    ' 0x80..0x97 | array (0x00..0x17 data items follow)
    Case &H80 To &H97
        ItemCount = CborBytes(Index) And &H1F
        ItemLength = _
            GetCborLengthFromItemCborBytes(ItemCount, CborBytes, Index + 1)
        GetCborLength = 1 + ItemLength
        
    ' 0x98 | array (one-byte uint8_t for n, and then n data items follow)
    Case &H98
        ItemCount = CborBytes(Index + 1)
        ItemLength = _
            GetCborLengthFromItemCborBytes(ItemCount, CborBytes, Index + 1 + 1)
        GetCborLength = 1 + 1 + ItemLength
        
    ' 0x99 | array (two-byte uint16_t for n, and then n data items follow)
    Case &H99
        ItemCount = GetUInt16FromBytes(CborBytes, Index + 1, True)
        ItemLength = _
            GetCborLengthFromItemCborBytes(ItemCount, CborBytes, Index + 1 + 2)
        GetCborLength = 1 + 2 + ItemLength
        
    ' 0x9a | array (four-byte uint32_t for n, and then n data items follow)
    Case &H9A
        ItemCount = CLng(GetUInt32FromBytes(CborBytes, Index + 1, True))
        ItemLength = _
            GetCborLengthFromItemCborBytes(ItemCount, CborBytes, Index + 1 + 4)
        GetCborLength = 1 + 4 + ItemLength
        
    ' 0x9b | array (eight-byte uint64_t for n, and then n data items follow)
    'Case &H9B
    '    ItemCount = CLng(GetUInt64FromBytes(CborBytes, Index + 1, True))
    '    ItemLength = _
    '        GetCborLengthFromItemCborBytes(ItemCount, CborBytes, Index + 1 + 8)
    '    GetCborLength = 1 + 8 + ItemLength
        
    ' 0x9f | array, data items follow, terminated by "break"
    'Case &H9F
    '    ' to do
    '    ItemLength = 0
    '    GetCborLength = 1 + ItemLength + 1
        
    '
    ' major type 5: map
    '
    
    ' 0xa0..0xb7 | map (0x00..0x17 pairs of data items follow)
    Case &HA0 To &HB7
        ItemCount = CborBytes(Index) And &H1F
        ItemLength = _
            GetCborLengthFromItemCborBytes(ItemCount * 2, CborBytes, Index + 1)
        GetCborLength = 1 + ItemLength
        
    ' 0xb8 | map (one-byte uint8_t for n, and then n pairs of data items follow)
    Case &HB8
        ItemCount = CborBytes(Index + 1)
        ItemLength = _
            GetCborLengthFromItemCborBytes(ItemCount * 2, CborBytes, Index + 1 + 1)
        GetCborLength = 1 + 1 + ItemLength
        
    ' 0xb9 | map (two-byte uint16_t for n, and then n pairs of data items follow)
    Case &HB9
        ItemCount = GetUInt16FromBytes(CborBytes, Index + 1, True)
        ItemLength = _
            GetCborLengthFromItemCborBytes(ItemCount * 2, CborBytes, Index + 1 + 2)
        GetCborLength = 1 + 2 + ItemLength
        
    ' 0xba | map (four-byte uint32_t for n, and then n pairs of data items follow)
    Case &HBA
        ItemCount = CLng(GetUInt32FromBytes(CborBytes, Index + 1, True))
        ItemLength = _
            GetCborLengthFromItemCborBytes(ItemCount * 2, CborBytes, Index + 1 + 4)
        GetCborLength = 1 + 4 + ItemLength
        
    ' 0xbb | map (eight-byte uint64_t for n, and then n pairs of data items follow)
    'Case &HBB
    '    ItemCount = CLng(GetUInt64FromBytes(CborBytes, Index + 1, True))
    '    ItemLength = _
    '        GetCborLengthFromItemCborBytes(ItemCount * 2, CborBytes, Index + 1 + 8)
    '    GetCborLength = 1 + 8 + ItemLength
        
    ' 0xbf | map, pairs of data items follow, terminated by "break"
    'Case &HBF
    '    ' to do
    '    ItemLength = 0
    '    GetCborLength = 1 + ItemLength + 1
        
    '
    ' major type 6: tag
    '
    
    ' 0xc0 | text-based date/time (data item follows; see Section 3.4.1)
    ' 0xc1 | epoch-based date/time (data item follows; see Section 3.4.2)
    ' 0xc2 | unsigned bignum (data item "byte string" follows)
    ' 0xc3 | negative bignum (data item "byte string" follows)
    ' 0xc4 | decimal Fraction (data item "array" follows; see Section 3.4.4)
    ' 0xc5 | bigfloat (data item "array" follows; see Section 3.4.4)
    ' 0xc6..0xd4 | (tag)
    ' 0xd5..0xd7 | expected conversion (data item follows; see Section 3.4.5.2)
    ' 0xd8..0xdb | (more tags; 1/2/4/8 bytes of tag number and then a data item follow)
    
    '
    ' major type 7: simple/float
    '
    
    ' 0xf4 | false
    Case &HF4
        GetCborLength = 1
        
    ' 0xf5 | true
    Case &HF5
        GetCborLength = 1
        
    ' 0xf6 | null
    Case &HF6
        GetCborLength = 1
        
    ' 0xf7 | undefined
    Case &HF7
        GetCborLength = 1
        
    ' 0xf9 | half-precision float (two-byte IEEE 754)
    'Case &HF9
    '    GetCborLength = 1 + 2
        
    ' 0xfa | single-precision float (four-byte IEEE 754)
    Case &HFA
        GetCborLength = 1 + 4
        
    ' 0xfb | double-precision float (eight-byte IEEE 754)
    Case &HFB
        GetCborLength = 1 + 8
        
    ' 0xff | "break" stop code
    'Case &HFF
    '    GetCborLength = 1
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Private Function GetCborLengthFromItemCborBytes(ByVal ItemCount As Long, _
    CborBytes() As Byte, Optional Index As Long) As Long
    
    Dim Length As Long
    
    Dim Count As Long
    For Count = 1 To ItemCount
        Length = Length + GetCborLength(CborBytes, Index + Length)
    Next
    
    GetCborLengthFromItemCborBytes = Length
End Function

Public Function IsCborObject( _
    CborBytes() As Byte, Optional Index As Long) As Boolean
    
    Select Case CborBytes(Index)
    
    '
    ' major type 0: positive integer
    '
    
    ' 0x00..0x17 | unsigned integer 0x00..0x17 (0..23)
    ' 0x18 | unsigned integer (one-byte uint8_t follows)
    ' 0x19 | unsigned integer (two-byte uint16_t follows)
    ' 0x1a | unsigned integer (four-byte uint32_t follows)
    ' 0x1b | unsigned integer (eight-byte uint64_t follows)
    Case &H0 To &H1B
        IsCborObject = False
        
    '
    ' major type 1: negative integer
    '
    
    ' 0x20..0x37 | negative integer -1-0x00..-1-0x17 (-1..-24)
    ' 0x38 | negative integer -1-n (one-byte uint8_t for n follows)
    ' 0x39 | negative integer -1-n (two-byte uint16_t for n follows)
    ' 0x3a | negative integer -1-n (four-byte uint32_t for n follows)
    ' 0x3b | negative integer -1-n (eight-byte uint64_t for n follows)
    Case &H20 To &H3B
        IsCborObject = False
        
    '
    ' major type 2: byte string
    '
    
    ' 0x40..0x57 | byte string (0x00..0x17 bytes follow)
    ' 0x58 | byte string (one-byte uint8_t for n, and then n bytes follow)
    ' 0x59 | byte string (two-byte uint16_t for n, and then n bytes follow)
    ' 0x5a | byte string (four-byte uint32_t for n, and then n bytes follow)
    ' 0x5b | byte string (eight-byte uint64_t for n, and then n bytes follow)
    ' 0x5f | byte string, byte strings follow, terminated by "break"
    Case &H40 To &H5B, &H5F
        IsCborObject = False
        
    '
    ' major type 3: text string
    '
    
    ' 0x60..0x77 | UTF-8 string (0x00..0x17 bytes follow)
    ' 0x78 | UTF-8 string (one-byte uint8_t for n, and then n bytes follow)
    ' 0x79 | UTF-8 string (two-byte uint16_t for n, and then n bytes follow)
    ' 0x7a | UTF-8 string (four-byte uint32_t for n, and then n bytes follow)
    ' 0x7b | UTF-8 string (eight-byte uint64_t for n, and then n bytes follow)
    ' 0x7f | UTF-8 string, UTF-8 strings follow, terminated by "break"
    Case &H60 To &H7B, &H7F
        IsCborObject = False
        
    '
    ' major type 4: array
    '
    
    ' 0x80..0x97 | array (0x00..0x17 data items follow)
    ' 0x98 | array (one-byte uint8_t for n, and then n data items follow)
    ' 0x99 | array (two-byte uint16_t for n, and then n data items follow)
    ' 0x9a | array (four-byte uint32_t for n, and then n data items follow)
    ' 0x9b | array (eight-byte uint64_t for n, and then n data items follow)
    ' 0x9f | array, data items follow, terminated by "break"
    Case &H80 To &H9B ', &H9F
        #If USE_COLLECTION Then
        IsCborObject = True
        #Else
        IsCborObject = False
        #End If
        
    '
    ' major type 5: map
    '
    
    ' 0xa0..0xb7 | map (0x00..0x17 pairs of data items follow)
    ' 0xb8 | map (one-byte uint8_t for n, and then n pairs of data items follow)
    ' 0xb9 | map (two-byte uint16_t for n, and then n pairs of data items follow)
    ' 0xba | map (four-byte uint32_t for n, and then n pairs of data items follow)
    ' 0xbb | map (eight-byte uint64_t for n, and then n pairs of data items follow)
    ' 0xbf | map, pairs of data items follow, terminated by "break"
    Case &HA0 To &HBB ', &HBF
        IsCborObject = True
        
    '
    ' major type 6: tag
    '
    
    ' 0xc0 | text-based date/time (data item follows; see Section 3.4.1)
    ' 0xc1 | epoch-based date/time (data item follows; see Section 3.4.2)
    ' 0xc2 | unsigned bignum (data item "byte string" follows)
    ' 0xc3 | negative bignum (data item "byte string" follows)
    ' 0xc4 | decimal Fraction (data item "array" follows; see Section 3.4.4)
    ' 0xc5 | bigfloat (data item "array" follows; see Section 3.4.4)
    ' 0xc6..0xd4 | (tag)
    ' 0xd5..0xd7 | expected conversion (data item follows; see Section 3.4.5.2)
    ' 0xd8..0xdb | (more tags; 1/2/4/8 bytes of tag number and then a data item follow)
    
    '
    ' major type 7: simple/float
    '
    
    ' 0xf4 | false
    ' 0xf5 | true
    ' 0xf6 | null
    ' 0xf7 | undefined
    Case &HF4 To &HF7 ', &HFF
        IsCborObject = False
        
    ' 0xf9 | half-precision float (two-byte IEEE 754)
    ' 0xfa | single-precision float (four-byte IEEE 754)
    ' 0xfb | double-precision float (eight-byte IEEE 754)
    Case &HF9 To &HFB
        IsCborObject = False
        
    ' 0xff | "break" stop code
    
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

Public Function GetValue( _
    CborBytes() As Byte, Optional Index As Long) As Variant
    
    Select Case CborBytes(Index)
    
    '
    ' major type 0: positive integer
    '
    
    ' 0x00..0x17 | unsigned integer 0x00..0x17 (0..23)
    Case &H0 To &H17
        GetValue = GetPosFixIntFromCborBytes(CborBytes, Index)
        
    ' 0x18 | unsigned integer (one-byte uint8_t follows)
    Case &H18
        GetValue = GetPosInt8FromCborBytes(CborBytes, Index)
        
    ' 0x19 | unsigned integer (two-byte uint16_t follows)
    Case &H19
        GetValue = GetPosInt16FromCborBytes(CborBytes, Index)
        
    ' 0x1a | unsigned integer (four-byte uint32_t follows)
    Case &H1A
        GetValue = GetPosInt32FromCborBytes(CborBytes, Index)
        
    ' 0x1b | unsigned integer (eight-byte uint64_t follows)
    Case &H1B
        GetValue = GetPosInt64FromCborBytes(CborBytes, Index)
        
    '
    ' major type 1: negative integer
    '
    
    ' 0x20..0x37 | negative integer -1-0x00..-1-0x17 (-1..-24)
    Case &H20 To &H37
        GetValue = GetNegFixIntFromCborBytes(CborBytes, Index)
        
    ' 0x38 | negative integer -1-n (one-byte uint8_t for n follows)
    Case &H38
        GetValue = GetNegInt8FromCborBytes(CborBytes, Index)
        
    ' 0x39 | negative integer -1-n (two-byte uint16_t for n follows)
    Case &H39
        GetValue = GetNegInt16FromCborBytes(CborBytes, Index)
        
    ' 0x3a | negative integer -1-n (four-byte uint32_t for n follows)
    Case &H3A
        GetValue = GetNegInt32FromCborBytes(CborBytes, Index)
        
    ' 0x3b | negative integer -1-n (eight-byte uint64_t for n follows)
    Case &H3B
        GetValue = GetNegInt64FromCborBytes(CborBytes, Index)
        
    '
    ' major type 2: byte string
    '
    
    ' 0x40..0x57 | byte string (0x00..0x17 bytes follow)
    Case &H40 To &H57
        GetValue = GetFixBinFromCborBytes(CborBytes, Index)
        
    ' 0x58 | byte string (one-byte uint8_t for n, and then n bytes follow)
    Case &H58
        GetValue = GetBin8FromCborBytes(CborBytes, Index)
        
    ' 0x59 | byte string (two-byte uint16_t for n, and then n bytes follow)
    Case &H59
        GetValue = GetBin16FromCborBytes(CborBytes, Index)
        
    ' 0x5a | byte string (four-byte uint32_t for n, and then n bytes follow)
    Case &H5A
        GetValue = GetBin32FromCborBytes(CborBytes, Index)
        
    ' 0x5b | byte string (eight-byte uint64_t for n, and then n bytes follow)
    'Case &H5B
    '    GetValue = GetBin64FromCborBytes(CborBytes, Index)
        
    ' 0x5f | byte string, byte strings follow, terminated by "break"
    'Case &H5F
    '    GetValue = GetBinBreakFromCborBytes(CborBytes, Index)
        
    '
    ' major type 3: text string
    '
    
    ' 0x60..0x77 | UTF-8 string (0x00..0x17 bytes follow)
    Case &H60 To &H77
        GetValue = GetFixStrFromCborBytes(CborBytes, Index)
        
    ' 0x78 | UTF-8 string (one-byte uint8_t for n, and then n bytes follow)
    Case &H78
        GetValue = GetStr8FromCborBytes(CborBytes, Index)
        
    ' 0x79 | UTF-8 string (two-byte uint16_t for n, and then n bytes follow)
    Case &H79
        GetValue = GetStr16FromCborBytes(CborBytes, Index)
        
    ' 0x7a | UTF-8 string (four-byte uint32_t for n, and then n bytes follow)
    Case &H7A
        GetValue = GetStr32FromCborBytes(CborBytes, Index)
        
    ' 0x7b | UTF-8 string (eight-byte uint64_t for n, and then n bytes follow)
    'Case &H7B
    '    GetValue = GetStr64FromCborBytes(CborBytes, Index)
        
    ' 0x7f | UTF-8 string, UTF-8 strings follow, terminated by "break"
    'Case &H7F
    '    GetValue = GetStrBreakFromCborBytes(CborBytes, Index)
        
    '
    ' major type 4: array
    '
    
    ' 0x80..0x97 | array (0x00..0x17 data items follow)
    Case &H80 To &H97
        #If USE_COLLECTION Then
        Set GetValue = GetFixArrayFromCborBytes(CborBytes, Index)
        #Else
        GetValue = GetFixArrayFromCborBytes(CborBytes, Index)
        #End If
        
    ' 0x98 | array (one-byte uint8_t for n, and then n data items follow)
    Case &H98
        #If USE_COLLECTION Then
        Set GetValue = GetArray8FromCborBytes(CborBytes, Index)
        #Else
        GetValue = GetArray8FromCborBytes(CborBytes, Index)
        #End If
        
    ' 0x99 | array (two-byte uint16_t for n, and then n data items follow)
    Case &H99
        #If USE_COLLECTION Then
        Set GetValue = GetArray16FromCborBytes(CborBytes, Index)
        #Else
        GetValue = GetArray16FromCborBytes(CborBytes, Index)
        #End If
        
    ' 0x9a | array (four-byte uint32_t for n, and then n data items follow)
    Case &H9A
        #If USE_COLLECTION Then
        Set GetValue = GetArray32FromCborBytes(CborBytes, Index)
        #Else
        GetValue = GetArray32FromCborBytes(CborBytes, Index)
        #End If
        
    ' 0x9b | array (eight-byte uint64_t for n, and then n data items follow)
    'Case &H9B
    '    #If USE_COLLECTION Then
    '    Set GetValue = GetArray64FromCborBytes(CborBytes, Index)
    '    #Else
    '    GetValue = GetArray64FromCborBytes(CborBytes, Index)
    '    #End If
        
    ' 0x9f | array, data items follow, terminated by "break"
    'Case &H9F
    '    #If USE_COLLECTION Then
    '    Set GetValue = GetArrayBreakFromCborBytes(CborBytes, Index)
    '    #Else
    '    GetValue = GetArrayBreakFromCborBytes(CborBytes, Index)
    '    #End If
        
    '
    ' major type 5: map
    '
    
    ' 0xa0..0xb7 | map (0x00..0x17 pairs of data items follow)
    Case &HA0 To &HB7
        Set GetValue = GetFixMapFromCborBytes(CborBytes, Index)
        
    ' 0xb8 | map (one-byte uint8_t for n, and then n pairs of data items follow)
    Case &HB8
        Set GetValue = GetMap8FromCborBytes(CborBytes, Index)
        
    ' 0xb9 | map (two-byte uint16_t for n, and then n pairs of data items follow)
    Case &HB9
        Set GetValue = GetMap16FromCborBytes(CborBytes, Index)
        
    ' 0xba | map (four-byte uint32_t for n, and then n pairs of data items follow)
    Case &HBA
        Set GetValue = GetMap32FromCborBytes(CborBytes, Index)
        
    ' 0xbb | map (eight-byte uint64_t for n, and then n pairs of data items follow)
    'Case &HBB
    '    Set GetValue = GetMap64FromCborBytes(CborBytes, Index)
        
    ' 0xbf | map, pairs of data items follow, terminated by "break"
    'Case &HBF
    '    Set GetValue = GetMapBreakFromCborBytes(CborBytes, Index)
        
    '
    ' major type 6: tag
    '
    
    ' 0xc0 | text-based date/time (data item follows; see Section 3.4.1)
    ' 0xc1 | epoch-based date/time (data item follows; see Section 3.4.2)
    ' 0xc2 | unsigned bignum (data item "byte string" follows)
    ' 0xc3 | negative bignum (data item "byte string" follows)
    ' 0xc4 | decimal Fraction (data item "array" follows; see Section 3.4.4)
    ' 0xc5 | bigfloat (data item "array" follows; see Section 3.4.4)
    ' 0xc6..0xd4 | (tag)
    ' 0xd5..0xd7 | expected conversion (data item follows; see Section 3.4.5.2)
    ' 0xd8..0xdb | (more tags; 1/2/4/8 bytes of tag number and then a data item follow)
    
    '
    ' major type 7: simple/float
    '
    
    ' 0xf4 | false
    Case &HF4
        GetValue = GetFalseFromCborBytes(CborBytes, Index)
        
    ' 0xf5 | true
    Case &HF5
        GetValue = GetTrueFromCborBytes(CborBytes, Index)
        
    ' 0xf6 | null
    Case &HF6
        GetValue = GetNullFromCborBytes(CborBytes, Index)
        
    ' 0xf7 | undefined
    Case &HF7
        GetValue = GetUndefinedFromCborBytes(CborBytes, Index)
        
    ' 0xf9 | half-precision float (two-byte IEEE 754)
    'Case &HF9
    '    GetValue = GetFloat16FromCborBytes(CborBytes, Index)
        
    ' 0xfa | single-precision float (four-byte IEEE 754)
    Case &HFA
        GetValue = GetFloat32FromCborBytes(CborBytes, Index)
        
    ' 0xfb | double-precision float (eight-byte IEEE 754)
    Case &HFB
        GetValue = GetFloat64FromCborBytes(CborBytes, Index)
        
    ' 0xff | "break" stop code
    'Case &HFF
    '    ' to do
        
    Case Else
        Err.Raise 13 ' unmatched type
        
    End Select
End Function

'
' major type 0: positive integer
'

' 0x00..0x17 | unsigned integer 0x00..0x17 (0..23)

Private Function GetPosFixIntFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Byte
    
    GetPosFixIntFromCborBytes = CborBytes(Index)
End Function

' 0x18 | unsigned integer (one-byte uint8_t follows)

Private Function GetPosInt8FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Byte
    
    GetPosInt8FromCborBytes = CborBytes(Index + 1)
End Function

' 0x19 | unsigned integer (two-byte uint16_t follows)

Private Function GetPosInt16FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Long
    
    GetPosInt16FromCborBytes = GetUInt16FromBytes(CborBytes, Index + 1, True)
End Function

' 0x1a | unsigned integer (four-byte uint32_t follows)

Private Function GetPosInt32FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long)
    
    GetPosInt32FromCborBytes = GetUInt32FromBytes(CborBytes, Index + 1, True)
End Function

' 0x1b | unsigned integer (eight-byte uint64_t follows)

Private Function GetPosInt64FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long)
    
    GetPosInt64FromCborBytes = GetUInt64FromBytes(CborBytes, Index + 1, True)
End Function

'
' major type 1: negative integer
'

' 0x20..0x37 | negative integer -1-0x00..-1-0x17 (-1..-24)

Private Function GetNegFixIntFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Integer
    
    GetNegFixIntFromCborBytes = -1 - (CborBytes(Index) And &H1F)
End Function

' 0x38 | negative integer -1-n (one-byte uint8_t for n follows)

Private Function GetNegInt8FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Integer
    
    GetNegInt8FromCborBytes = -1 - CborBytes(Index + 1)
End Function

' 0x39 | negative integer -1-n (two-byte uint16_t for n follows)

Private Function GetNegInt16FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Long
    
    GetNegInt16FromCborBytes = _
        -1 - GetUInt16FromBytes(CborBytes, Index + 1, True)
End Function

' 0x3a | negative integer -1-n (four-byte uint32_t for n follows)

#If Win64 And USE_LONGLONG Then

Private Function GetNegInt32FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As LongLong
    
    GetNegInt32FromCborBytes = _
        -1 - GetUInt32FromBytes(CborBytes, Index + 1, True)
End Function

#Else

Private Function GetNegInt32FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long)
    
    GetNegInt32FromCborBytes = _
        -1 - GetUInt32FromBytes(CborBytes, Index + 1, True)
End Function

#End If

' 0x3b | negative integer -1-n (eight-byte uint64_t for n follows)

Private Function GetNegInt64FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long)
    
    GetNegInt64FromCborBytes = _
        -1 - GetUInt64FromBytes(CborBytes, Index + 1, True)
End Function

'
' major type 2: byte string
'

' 0x40..0x57 | byte string (0x00..0x17 bytes follow)

Private Function GetFixBinFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Byte()
    
    Dim BinBytes() As Byte
    
    Dim Length As Byte
    Length = CborBytes(Index) And &H1F
    If Length = 0 Then
        GetFixBinFromCborBytes = BinBytes
        Exit Function
    End If
    
    ReDim BinBytes(0 To Length - 1)
    CopyBytes BinBytes, 0, CborBytes, Index + 1, Length
    
    GetFixBinFromCborBytes = BinBytes
End Function

' 0x58 | byte string (one-byte uint8_t for n, and then n bytes follow)

Private Function GetBin8FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Byte()
    
    Dim BinBytes() As Byte
    
    Dim Length As Byte
    Length = CborBytes(Index + 1)
    If Length = 0 Then
        GetBin8FromCborBytes = BinBytes
        Exit Function
    End If
    
    ReDim BinBytes(0 To Length - 1)
    CopyBytes BinBytes, 0, CborBytes, Index + 1 + 1, Length
    
    GetBin8FromCborBytes = BinBytes
End Function

' 0x59 | byte string (two-byte uint16_t for n, and then n bytes follow)

Private Function GetBin16FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Byte()
    
    Dim BinBytes() As Byte
    
    Dim Length As Long
    Length = GetUInt16FromBytes(CborBytes, Index + 1, True)
    If Length = 0 Then
        GetBin16FromCborBytes = BinBytes
        Exit Function
    End If
    
    ReDim BinBytes(0 To Length - 1)
    CopyBytes BinBytes, 0, CborBytes, Index + 1 + 2, Length
    
    GetBin16FromCborBytes = BinBytes
End Function

' 0x5a | byte string (four-byte uint32_t for n, and then n bytes follow)

Private Function GetBin32FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Byte()
    
    Dim BinBytes() As Byte
    
    Dim Length As Long
    Length = CLng(GetUInt32FromBytes(CborBytes, Index + 1, True))
    If Length = 0 Then
        GetBin32FromCborBytes = BinBytes
        Exit Function
    End If
    
    ReDim BinBytes(0 To Length - 1)
    CopyBytes BinBytes, 0, CborBytes, Index + 1 + 4, Length
    
    GetBin32FromCborBytes = BinBytes
End Function

' 0x5b | byte string (eight-byte uint64_t for n, and then n bytes follow)

'Private Function GetBin64FromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long) As Byte()
'
'    Dim BinBytes() As Byte
'
'    Dim Length As LongLong
'    Length = CLngLng(GetUInt64FromBytes(CborBytes, Index + 1, True))
'    If Length = 0 Then
'        GetBin64FromCborBytes = BinBytes
'        Exit Function
'    End If
'
'    ReDim BinBytes(0 To Length - 1)
'    CopyBytes BinBytes, 0, CborBytes, Index + 1 + 8, Length
'
'    GetBin64FromCborBytes = BinBytes
'End Function

' 0x5f | byte string, byte strings follow, terminated by "break"

'Private Function GetBinBreakFromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long) As Byte()
'
'    Dim BinBytes() As Byte
'
'    Dim Length As Long
'    Length = GetBreakIndexFromCborBytes(CborBytes, Index + 1) - (Index + 1)
'    If Length = 0 Then
'        GetBinBreakFromCborBytes = BinBytes
'        Exit Function
'    End If
'
'    ReDim BinBytes(0 To Length - 1)
'    CopyBytes BinBytes, 0, CborBytes, Index + 1, Length
'
'    GetBinBreakFromCborBytes = BinBytes
'End Function

'
' major type 3: text string
'

' 0x60..0x77 | UTF-8 string (0x00..0x17 bytes follow)

Private Function GetFixStrFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As String
    
    Dim Length As Byte
    Length = (CborBytes(Index) And &H1F)
    If Length = 0 Then
        GetFixStrFromCborBytes = ""
        Exit Function
    End If
    
    GetFixStrFromCborBytes = GetStringFromBytes(CborBytes, Index + 1, Length)
End Function

' 0x78 | UTF-8 string (one-byte uint8_t for n, and then n bytes follow)

Private Function GetStr8FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As String
    
    Dim Length As Byte
    Length = CborBytes(Index + 1)
    If Length = 0 Then
        GetStr8FromCborBytes = ""
        Exit Function
    End If
    
    GetStr8FromCborBytes = _
        GetStringFromBytes(CborBytes, Index + 1 + 1, Length)
End Function

' 0x79 | UTF-8 string (two-byte uint16_t for n, and then n bytes follow)

Private Function GetStr16FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As String
    
    Dim Length As Long
    Length = GetUInt16FromBytes(CborBytes, Index + 1, True)
    If Length = 0 Then
        GetStr16FromCborBytes = ""
        Exit Function
    End If
    
    GetStr16FromCborBytes = _
        GetStringFromBytes(CborBytes, Index + 1 + 2, Length)
End Function

' 0x7a | UTF-8 string (four-byte uint32_t for n, and then n bytes follow)

Private Function GetStr32FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As String
    
    Dim Length As Long
    Length = CLng(GetUInt32FromBytes(CborBytes, Index + 1, True))
    If Length = 0 Then
        GetStr32FromCborBytes = ""
        Exit Function
    End If
    
    GetStr32FromCborBytes = _
        GetStringFromBytes(CborBytes, Index + 1 + 4, Length)
End Function

' 0x7b | UTF-8 string (eight-byte uint64_t for n, and then n bytes follow)

'Private Function GetStr64FromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long) As String
'
'    Dim Length As LongLong
'    Length = CLngLng(GetUInt32FromBytes(CborBytes, Index + 1, True))
'    If Length = 0 Then
'        GetStr64FromCborBytes = ""
'        Exit Function
'    End If
'
'    GetStr64FromCborBytes = _
'        GetStringFromBytes(CborBytes, Index + 1 + 8, Length)
'End Function

' 0x7f | UTF-8 string, UTF-8 strings follow, terminated by "break"

'Private Function GetStrBreakFromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long) As String
'
'    Dim Length As Long
'    Length = GetBreakIndexFromCborBytes(CborBytes, Index + 1) - (Index + 1)
'    If Length = 0 Then
'        GetStrBreakFromCborBytes = ""
'        Exit Function
'    End If
'
'    GetStrBreakFromCborBytes = _
'        GetStringFromBytes(CborBytes, Index + 1, Length)
'End Function

'
' major type 4: array
'

' 0x80..0x97 | array (0x00..0x17 data items follow)

#If USE_COLLECTION Then

Private Function GetFixArrayFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Collection
    
    Dim ItemCount As Long
    ItemCount = CborBytes(Index) And &H1F
    
    Set GetFixArrayFromCborBytes = _
        GetArrayFromCborBytes(CborBytes, Index + 1, ItemCount)
End Function

#Else

Private Function GetFixArrayFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long)
    
    Dim ItemCount As Long
    ItemCount = CborBytes(Index) And &HF
    
    GetFixArrayFromCborBytes = _
        GetArrayFromCborBytes(CborBytes, Index + 1, ItemCount)
End Function

#End If

' 0x98 | array (one-byte uint8_t for n, and then n data items follow)

#If USE_COLLECTION Then

Private Function GetArray8FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Collection
    
    Dim ItemCount As Long
    ItemCount = CborBytes(Index + 1)
    
    Set GetArray8FromCborBytes = _
        GetArrayFromCborBytes(CborBytes, Index + 1 + 1, ItemCount)
End Function

#Else

Private Function GetArray8FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long)
    
    Dim ItemCount As Long
    ItemCount = CborBytes(Index + 1)
    
    GetArray8FromCborBytes = _
        GetArrayFromCborBytes(CborBytes, Index + 1 + 1, ItemCount)
End Function

#End If

' 0x99 | array (two-byte uint16_t for n, and then n data items follow)

#If USE_COLLECTION Then

Private Function GetArray16FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Collection
    
    Dim ItemCount As Long
    ItemCount = GetUInt16FromBytes(CborBytes, Index + 1, True)
    
    Set GetArray16FromCborBytes = _
        GetArrayFromCborBytes(CborBytes, Index + 1 + 2, ItemCount)
End Function

#Else

Private Function GetArray16FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long)
    
    Dim ItemCount As Long
    ItemCount = GetUInt16FromBytes(CborBytes, Index + 1, True)
    
    GetArray16FromCborBytes = _
        GetArrayFromCborBytes(CborBytes, Index + 1 + 2, ItemCount)
End Function

#End If

' 0x9a | array (four-byte uint32_t for n, and then n data items follow)

#If USE_COLLECTION Then

Private Function GetArray32FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Collection
    
    Dim ItemCount As Long
    ItemCount = CLng(GetUInt32FromBytes(CborBytes, Index + 1, True))
    
    Set GetArray32FromCborBytes = _
        GetArrayFromCborBytes(CborBytes, Index + 1 + 4, ItemCount)
End Function

#Else

Private Function GetArray32FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long)
    
    Dim ItemCount As Long
    ItemCount = CLng(GetUInt32FromBytes(CborBytes, Index + 1, True))
    
    GetArray32FromCborBytes = _
        GetArrayFromCborBytes(CborBytes, Index + 1 + 4, ItemCount)
End Function

#End If

' 0x9b | array (eight-byte uint64_t for n, and then n data items follow)

'#If USE_COLLECTION Then
'
'Private Function GetArray64FromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long) As Collection
'
'    Dim ItemCount As Long
'    ItemCount = CLng(GetUInt64FromBytes(CborBytes, Index + 1, True))
'
'    Set GetArray64FromCborBytes = _
'        GetArrayFromCborBytes(CborBytes, Index + 1 + 8, ItemCount)
'End Function
'
'#Else
'
'Private Function GetArray64FromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long)
'
'    Dim ItemCount As Long
'    ItemCount = CLng(GetUInt64FromBytes(CborBytes, Index + 1, True))
'
'    GetArray64FromCborBytes = _
'        GetArrayFromCborBytes(CborBytes, Index + 1 + 8, ItemCount)
'End Function
'
'#End If

' 0x9f | array, data items follow, terminated by "break"

'#If USE_COLLECTION Then
'
'Private Function GetArrayBreakFromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long) As Collection
'
'    Dim ItemCount As Long
'    ItemCount = 0 ' to do
'
'    Set GetArrayBreakFromCborBytes = _
'        GetArrayFromCborBytes(CborBytes, Index + 1, ItemCount)
'End Function
'
'#Else
'
'Private Function GetArrayBreakFromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long)
'
'    Dim ItemCount As Long
'    ItemCount = 0 ' to do
'
'    GetArrayBreakFromCborBytes = _
'        GetArrayFromCborBytes(CborBytes, Index + 1, ItemCount)
'End Function
'
'#End If

'
' major type 5: map
'

' 0xa0..0xb7 | map (0x00..0x17 pairs of data items follow)

Private Function GetFixMapFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Object
    
    Dim ItemCount As Long
    ItemCount = CborBytes(Index) And &H1F
    
    Set GetFixMapFromCborBytes = _
        GetMapFromCborBytes(CborBytes, Index + 1, ItemCount)
End Function

' 0xb8 | map (one-byte uint8_t for n, and then n pairs of data items follow)

Private Function GetMap8FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Object
    
    Dim ItemCount As Long
    ItemCount = CborBytes(Index + 1)
    
    Set GetMap8FromCborBytes = _
        GetMapFromCborBytes(CborBytes, Index + 1 + 1, ItemCount)
End Function

' 0xb9 | map (two-byte uint16_t for n, and then n pairs of data items follow)

Private Function GetMap16FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Object
    
    Dim ItemCount As Long
    ItemCount = GetUInt16FromBytes(CborBytes, Index + 1, True)
    
    Set GetMap16FromCborBytes = _
        GetMapFromCborBytes(CborBytes, Index + 1 + 2, ItemCount)
End Function

' 0xba | map (four-byte uint32_t for n, and then n pairs of data items follow)

Private Function GetMap32FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Object
    
    Dim ItemCount As Long
    ItemCount = CLng(GetUInt32FromBytes(CborBytes, Index + 1, True))
    
    Set GetMap32FromCborBytes = _
        GetMapFromCborBytes(CborBytes, Index + 1 + 4, ItemCount)
End Function

' 0xbb | map (eight-byte uint64_t for n, and then n pairs of data items follow)

'Private Function GetMap64FromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long) As Object
'
'    Dim ItemCount As Long
'    ItemCount = CLng(GetUInt64FromBytes(CborBytes, Index + 1, True))
'
'    Set GetMap64FromCborBytes = _
'        GetMapFromCborBytes(CborBytes, Index + 1 + 8, ItemCount)
'End Function

' 0xbf | map, pairs of data items follow, terminated by "break"

'Private Function GetMapBreakFromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long) As Object
'
'    Dim ItemCount As Long
'    ItemCount = 0 ' to do
'
'    Set GetMapBreakFromCborBytes = _
'        GetMapFromCborBytes(CborBytes, Index + 1, ItemCount)
'End Function

'
' major type 6: tag
'

' 0xc0 | text-based date/time (data item follows; see Section 3.4.1)
' 0xc1 | epoch-based date/time (data item follows; see Section 3.4.2)
' 0xc2 | unsigned bignum (data item "byte string" follows)
' 0xc3 | negative bignum (data item "byte string" follows)
' 0xc4 | decimal Fraction (data item "array" follows; see Section 3.4.4)
' 0xc5 | bigfloat (data item "array" follows; see Section 3.4.4)
' 0xc6..0xd4 | (tag)
' 0xd5..0xd7 | expected conversion (data item follows; see Section 3.4.5.2)
' 0xd8..0xdb | (more tags; 1/2/4/8 bytes of tag number and then a data item follow)

'
' major type 7: simple/float
'

' 0xf4 | false

Private Function GetFalseFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Boolean
    
    GetFalseFromCborBytes = False
End Function

' 0xf5 | true

Private Function GetTrueFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Boolean
    
    GetTrueFromCborBytes = True
End Function

' 0xf6 | null

Private Function GetNullFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long)
    
    GetNullFromCborBytes = Null
End Function

' 0xf7 | undefined

Private Function GetUndefinedFromCborBytes( _
    CborBytes() As Byte, Optional Index As Long)
    
    GetUndefinedFromCborBytes = Empty
End Function

' 0xf9 | half-precision float (two-byte IEEE 754)

'Private Function GetFloat16FromCborBytes( _
'    CborBytes() As Byte, Optional Index As Long) As Single
'
'    GetFloat16FromCborBytes = GetFloat16FromBytes(CborBytes, Index + 1, True)
'End Function

' 0xfa | single-precision float (four-byte IEEE 754)

Private Function GetFloat32FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Single
    
    GetFloat32FromCborBytes = GetFloat32FromBytes(CborBytes, Index + 1, True)
End Function

' 0xfb | double-precision float (eight-byte IEEE 754)

Private Function GetFloat64FromCborBytes( _
    CborBytes() As Byte, Optional Index As Long) As Double
    
    GetFloat64FromCborBytes = GetFloat64FromBytes(CborBytes, Index + 1, True)
End Function

' 0xff | "break" stop code

''
'' CBOR for VBA - Decoding - Array Helper
''

#If USE_COLLECTION Then

Private Function GetArrayFromCborBytes( _
    CborBytes() As Byte, Index As Long, ItemCount As Long) As Collection
    
    Dim Collection_ As Collection
    Set Collection_ = New Collection
    
    Dim Offset As Long
    Dim Count As Long
    For Count = 0 To ItemCount - 1
        Collection_.Add GetValue(CborBytes, Index + Offset)
        
        Offset = Offset + GetCborLength(CborBytes, Index + Offset)
    Next
    
    Set GetArrayFromCborBytes = Collection_
End Function

#Else

Private Function GetArrayFromCborBytes( _
    CborBytes() As Byte, Index As Long, ItemCount As Long)
    
    Dim Array_()
    
    If ItemCount = 0 Then
        GetArrayFromCborBytes = Array_
        Exit Function
    End If
    
    ReDim Array_(0 To ItemCount - 1)
    
    Dim Offset As Long
    Dim Count As Long
    For Count = 0 To ItemCount - 1
        If IsCborObject(CborBytes, Index + Offset) Then
            Set Array_(Count) = GetValue(CborBytes, Index + Offset)
        Else
            Array_(Count) = GetValue(CborBytes, Index + Offset)
        End If
        
        Offset = Offset + GetCborLength(CborBytes, Index + Offset)
    Next
    
    GetArrayFromCborBytes = Array_
End Function

#End If

''
'' CBOR for VBA - Decoding - Map Helper
''

Private Function GetMapFromCborBytes( _
    CborBytes() As Byte, Index As Long, ItemCount As Long) As Object
    
    Dim Map As Object
    Set Map = CreateObject("Scripting.Dictionary")
    
    Dim KeyOffset As Long
    Dim ValueOffset As Long
    
    Dim Count As Long
    For Count = 1 To ItemCount
        ValueOffset = KeyOffset + GetCborLength(CborBytes, Index + KeyOffset)
        
        Map.Add _
            GetValue(CborBytes, Index + KeyOffset), _
            GetValue(CborBytes, Index + ValueOffset)
        
        KeyOffset = ValueOffset + GetCborLength(CborBytes, Index + ValueOffset)
    Next
    
    Set GetMapFromCborBytes = Map
End Function

''
'' CBOR for VBA - Decoding - Converter
''

'
' 0x19. UInt16 - a 16-bit unsigned integer
'

Private Function GetUInt16FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Long
    
    Dim Bytes4(0 To 3) As Byte
    If BigEndian Then
        CopyBytes Bytes4, 2, Bytes, Index, 2
    Else
        CopyBytes Bytes4, 0, Bytes, Index, 2
    End If
    
    GetUInt16FromBytes = GetLongFromBytes(Bytes4, 0, BigEndian)
End Function

'
' 0x1a. UInt32 - a 32-bit unsigned integer
'

#If USE_LONGLONG Then

Private Function GetUInt32FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As LongLong
    
    Dim Bytes8(0 To 7) As Byte
    If BigEndian Then
        CopyBytes Bytes8, 4, Bytes, Index, 4
    Else
        CopyBytes Bytes8, 0, Bytes, Index, 4
    End If
    
    GetUInt32FromBytes = GetLongLongFromBytes(Bytes8, 0, BigEndian)
End Function

#Else

Private Function GetUInt32FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Variant
    
    Dim Bytes14(0 To 13) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        CopyBytes Bytes14, 10, Bytes, Index, 4
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        CopyBytes Bytes14, 0, Bytes, Index, 4
    End If
    
    GetUInt32FromBytes = GetDecimalFromBytes(Bytes14, 0, BigEndian)
End Function

#End If

'
' 0x1b. UInt64 - a 64-bit unsigned integer
'

Private Function GetUInt64FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Variant
    
    Dim Bytes14(0 To 13) As Byte
    If BigEndian Then
        ' sign            - 1 byte
        ' scale           - 1 byte
        ' data high bytes - 4 bytes
        ' data low bytes  - 8 bytes
        CopyBytes Bytes14, 6, Bytes, Index, 8
    Else
        ' data low bytes  - 8 bytes
        ' data high bytes - 4 bytes
        ' scale           - 1 byte
        ' sign            - 1 byte
        CopyBytes Bytes14, 0, Bytes, Index, 8
    End If
    
    GetUInt64FromBytes = GetDecimalFromBytes(Bytes14, 0, BigEndian)
End Function

'
' 0xf9. Float16 - an IEEE 754 half precision floating point number
'

'Private Function GetFloat16FromBytes(Bytes() As Byte, _
'    Optional Index As Long, Optional BigEndian As Boolean) As Single
'
'    GetFloat16FromBytes = GetSingleFromBytes(Bytes, Index, BigEndian)
'End Function

'
' 0xfa. Float32 - an IEEE 754 single precision floating point number
'

Private Function GetFloat32FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Single
    
    GetFloat32FromBytes = GetSingleFromBytes(Bytes, Index, BigEndian)
End Function

'
' 0xfb. Float64 - an IEEE 754 double precision floating point number
'

Private Function GetFloat64FromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Double
    
    GetFloat64FromBytes = GetDoubleFromBytes(Bytes, Index, BigEndian)
End Function

'
' 2. Integer - a 16-bit signed integer
'

Private Function GetIntegerFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Integer
    
    Dim B2 As Bytes2T
    CopyBytes B2.Bytes, 0, Bytes, Index, 2
    
    If BigEndian Then
        ReverseBytes B2.Bytes
    End If
    
    Dim I As IntegerT
    LSet I = B2
    
    GetIntegerFromBytes = I.Value
End Function

'
' 3. Long - a 32-bit signed integer
'

Private Function GetLongFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Long
    
    Dim B4 As Bytes4T
    CopyBytes B4.Bytes, 0, Bytes, Index, 4
    
    If BigEndian Then
        ReverseBytes B4.Bytes
    End If
    
    Dim L As LongT
    LSet L = B4
    
    GetLongFromBytes = L.Value
End Function

'
' 4. Single - an IEEE 754 single precision floating point number
'

Private Function GetSingleFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Single
    
    Dim B4 As Bytes4T
    CopyBytes B4.Bytes, 0, Bytes, Index, 4
    
    If BigEndian Then
        ReverseBytes B4.Bytes
    End If
    
    Dim S As SingleT
    LSet S = B4
    
    GetSingleFromBytes = S.Value
End Function

'
' 5. Double - an IEEE 754 double precision floating point number
'

Private Function GetDoubleFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Double
    
    Dim B8 As Bytes8T
    CopyBytes B8.Bytes, 0, Bytes, Index, 8
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    Dim D As DoubleT
    LSet D = B8
    
    GetDoubleFromBytes = D.Value
End Function

'
' 8. String
'

Private Function GetStringFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional ByVal Length As Long, _
    Optional Charset As String = "utf-8") As String
    
    If Length = 0 Then
        Length = UBound(Bytes) - Index + 1
    End If
    
    Dim Bytes_() As Byte
    ReDim Bytes_(0 To Length - 1)
    CopyBytes Bytes_, 0, Bytes, Index, Length
    
    Static ADODBStream As Object
    
    If ADODBStream Is Nothing Then
        Set ADODBStream = CreateObject("ADODB.Stream")
    End If
    
    With ADODBStream
        .Open
        
        .Type = 1 'ADODB.adTypeBinary
        .Write Bytes_
        
        .Position = 0
        .Type = 2 'ADODB.adTypeText
        .Charset = Charset
        GetStringFromBytes = .ReadText
        
        .Close
    End With
End Function

'
' 14. Decimal - a 96-bit unsigned integer
'with 8-bit scaling factor and 8-bit sign factor
'

Private Function GetDecimalFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As Variant
    
    ' BytesXX = Bytes:
    ' data low bytes  - 8 bytes
    ' data high bytes - 4 bytes
    ' scale           - 1 byte
    ' sign            - 1 byte
    Dim BytesXX(0 To 13) As Byte
    CopyBytes BytesXX, 0, Bytes, Index, 14
    
    If BigEndian Then
        ReverseBytes BytesXX
    End If
    
    ' BytesX:
    ' scale           - 1 byte
    ' sign            - 1 byte
    ' data high bytes - 4 bytes
    ' data low bytes  - 8 bytes
    Dim BytesX(0 To 13) As Byte
    
    ' scale
    BytesX(0) = BytesXX(12)
    
    ' sign
    BytesX(1) = BytesXX(13)
    
    ' data high bytes
    CopyBytes BytesX, 2, BytesXX, 8, 4
    
    ' data low bytes
    CopyBytes BytesX, 6, BytesXX, 0, 8
    
    ' BytesRaw:
    ' vartype         - 2 bytes
    ' scale           - 1 byte
    ' sign            - 1 byte
    ' data high bytes - 4 bytes
    ' data low bytes  - 8 bytes
    Dim BytesRaw() As Byte
    ReDim BytesRaw(0 To 15)
    BytesRaw(0) = 14
    BytesRaw(1) = 0
    CopyBytes BytesRaw, 2, BytesX, Index, 14
    
    Dim Value As Variant
    CopyMemory ByVal VarPtr(Value), ByVal VarPtr(BytesRaw(0)), 16
    
    GetDecimalFromBytes = Value
End Function

'
' 20. LongLong - a 64-bit signed integer
'

#If Win64 And USE_LONGLONG Then

Private Function GetLongLongFromBytes(Bytes() As Byte, _
    Optional Index As Long, Optional BigEndian As Boolean) As LongLong
    
    Dim B8 As Bytes8T
    CopyBytes B8.Bytes, 0, Bytes, Index, 8
    
    If BigEndian Then
        ReverseBytes B8.Bytes
    End If
    
    Dim LL As LongLongT
    LSet LL = B8
    
    GetLongLongFromBytes = LL.Value
End Function

#End If

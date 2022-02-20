Attribute VB_Name = "Test_CBOR"
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
'' CBOR for VBA - Test
''

' Test Counter
Private m_Test_Count As Long
Private m_Test_Success As Long
Private m_Test_Fail As Long

#If Win64 Then
#Const USE_LONGLONG = True
#End If

' Array
#Const USE_COLLECTION = True

Public Sub Test_Cbor()
    Test_Initialize
    
    Test_Cbor_PosFixInt_TestCases
    Test_Cbor_PosInt8_TestCases
    Test_Cbor_PosInt16_TestCases
    Test_Cbor_PosInt32_TestCases
    Test_Cbor_PosInt64_TestCases
    
    Test_Cbor_NegFixInt_TestCases
    Test_Cbor_NegInt8_TestCases
    Test_Cbor_NegInt16_TestCases
    Test_Cbor_NegInt32_TestCases
    Test_Cbor_NegInt64_TestCases
    
    Test_Cbor_FixBin_TestCases
    Test_Cbor_Bin8_TestCases
    Test_Cbor_Bin16_TestCases
    Test_Cbor_Bin32_TestCases
    
    Test_Cbor_FixStr_TestCases
    Test_Cbor_Str8_TestCases
    Test_Cbor_Str16_TestCases
    Test_Cbor_Str32_TestCases
    
    Test_Cbor_FixArray_TestCases
    Test_Cbor_Array8_TestCases
    Test_Cbor_Array16_TestCases
    Test_Cbor_Array32_TestCases
    
    Test_Cbor_FixMap_TestCases
    Test_Cbor_Map8_TestCases
    Test_Cbor_Map16_TestCases
    Test_Cbor_Map32_TestCases
    
    Test_Cbor_False_TestCases
    Test_Cbor_True_TestCases
    Test_Cbor_Null_TestCases
    Test_Cbor_Undefined_TestCases
    
    Test_Cbor_Float32_TestCases
    Test_Cbor_Float64_TestCases
    
    Test_Terminate
End Sub

'
' CBOR for VBA - Test Cases
'

Private Sub Test_Cbor_PosFixInt_TestCases()
    Debug.Print "Target: PosFixInt"
    
    Test_Cbor_Int_Core "00", &H0
    Test_Cbor_Int_Core "17", &H17
End Sub


Private Sub Test_Cbor_PosInt8_TestCases()
    Debug.Print "Target: PosInt8"
    
    Test_Cbor_Int_Core "18 18", &H18
    Test_Cbor_Int_Core "18 7F", &H7F
    Test_Cbor_Int_Core "18 80", &H80
    Test_Cbor_Int_Core "18 FF", &HFF
End Sub

Private Sub Test_Cbor_PosInt16_TestCases()
    Debug.Print "Target: PosInt16"
    
    Test_Cbor_Int_Core "19 01 00", &H100
    Test_Cbor_Int_Core "19 7F FF", &H7FFF
    Test_Cbor_Int_Core "19 80 00", &H8000&
    Test_Cbor_Int_Core "19 FF FF", &HFFFF&
End Sub

Private Sub Test_Cbor_PosInt32_TestCases()
    Debug.Print "Target: PosInt32"
    
    Test_Cbor_Int_Core "1A 00 01 00 00", &H10000
    Test_Cbor_Int_Core "1A 7F FF FF FF", &H7FFFFFFF
#If Win64 And USE_LONGLONG Then
    Test_Cbor_Int_Core "1A 80 00 00 00", &H80000000^
    Test_Cbor_Int_Core "1A FF FF FF FF", &HFFFFFFFF^
#Else
    Test_Cbor_Int_Core "1A 80 00 00 00", CDec("&H80000000")
    Test_Cbor_Int_Core "1A FF FF FF FF", CDec("&HFFFFFFFF")
#End If
End Sub

Private Sub Test_Cbor_PosInt64_TestCases()
    Debug.Print "Target: PosInt64"
#If Win64 And USE_LONGLONG Then
    Test_Cbor_Int_Core "1B 00 00 00 01 00 00 00 00", _
        CLngLng("4294967296")
    Test_Cbor_Int_Core "1B 7F FF FF FF FF FF FF FF", _
        CLngLng("9223372036854775807")
#Else
    Test_Cbor_Int_Core "1B 00 00 00 01 00 00 00 00", _
        CDec("4294967296")
    Test_Cbor_Int_Core "1B 7F FF FF FF FF FF FF FF", _
        CDec("9223372036854775807")
#End If
    Test_Cbor_Int_Core "1B 80 00 00 00 00 00 00 00", _
        CDec("9223372036854775808")
    Test_Cbor_Int_Core "1B FF FF FF FF FF FF FF FF", _
        CDec("18446744073709551615")
End Sub

Private Sub Test_Cbor_NegFixInt_TestCases()
    Debug.Print "Target: NegFixInt"
    
    Test_Cbor_Int_Core "20", -1
    Test_Cbor_Int_Core "37", -24
End Sub

Private Sub Test_Cbor_NegInt8_TestCases()
    Debug.Print "Target: NegInt8"
    
    Test_Cbor_Int_Core "38 18", -25
    Test_Cbor_Int_Core "38 7F", -128
    Test_Cbor_Int_Core "38 80", -129
    Test_Cbor_Int_Core "38 FF", -256
End Sub

Private Sub Test_Cbor_NegInt16_TestCases()
    Debug.Print "Target: NegInt16"
    
    Test_Cbor_Int_Core "39 01 00", -257
    Test_Cbor_Int_Core "39 7F FF", CInt(-32768)
    Test_Cbor_Int_Core "39 80 00", CLng(-32769)
    Test_Cbor_Int_Core "39 FF FF", -65536
End Sub

Private Sub Test_Cbor_NegInt32_TestCases()
    Debug.Print "Target: NegInt32"
    
    Test_Cbor_Int_Core "3A 00 01 00 00", -65537
    Test_Cbor_Int_Core "3A 7F FF FF FF", CLng("-2147483648")
#If Win64 And USE_LONGLONG Then
    Test_Cbor_Int_Core "3A 80 00 00 00", CLngLng("-2147483649")
    Test_Cbor_Int_Core "3A FF FF FF FF", CLngLng("-4294967296")
#Else
    Test_Cbor_Int_Core "3A 80 00 00 00", CDec("-2147483649")
    Test_Cbor_Int_Core "3A FF FF FF FF", CDec("-4294967296")
#End If
End Sub

Private Sub Test_Cbor_NegInt64_TestCases()
    Debug.Print "Target: NegInt64"
    
#If Win64 And USE_LONGLONG Then
    Test_Cbor_Int_Core "3B 00 00 00 01 00 00 00 00", _
        CLngLng("-4294967297")
    Test_Cbor_Int_Core "3B 7F FF FF FF FF FF FF FF", _
        CLngLng("-9223372036854775808")
#Else
    Test_Cbor_Int_Core "3B 00 00 00 01 00 00 00 00", _
        CDec("-4294967297")
    Test_Cbor_Int_Core "3B 7F FF FF FF FF FF FF FF", _
        CDec("-9223372036854775808")
#End If
    Test_Cbor_Int_Core "3B 80 00 00 00 00 00 00 00", _
        CDec("-9223372036854775809")
    Test_Cbor_Int_Core "3B FF FF FF FF FF FF FF FF", _
        CDec("-18446744073709551616")
End Sub

Private Sub Test_Cbor_FixBin_TestCases()
    Debug.Print "Target: FixBin"
    
    Test_Cbor_Bin_Core "40", ""
    Test_Cbor_Bin_Core2 "41", &H1
    Test_Cbor_Bin_Core2 "57", &H17
End Sub

Private Sub Test_Cbor_Bin8_TestCases()
    Debug.Print "Target: Bin8"
    
    Test_Cbor_Bin_Core2 "58 18", &H18
    Test_Cbor_Bin_Core2 "58 FF", &HFF
End Sub

Private Sub Test_Cbor_Bin16_TestCases()
    Debug.Print "Target: Bin16"
    
    Test_Cbor_Bin_Core2 "59 01 00", &H100
    Test_Cbor_Bin_Core2 "59 FF FF", &HFFFF&
End Sub

Private Sub Test_Cbor_Bin32_TestCases()
    Debug.Print "Target: Bin32"
    
    Test_Cbor_Bin_Core2 "5A 00 01 00 00", &H10000
End Sub

Private Sub Test_Cbor_FixStr_TestCases()
    Debug.Print "Target: FixStr"
    
    Test_Cbor_Str_Core "60", ""
    Test_Cbor_Str_Core "61 61", "a"
    Test_Cbor_Str_Core "63 E3 81 82", ChrW(&H3042)
    Test_Cbor_Str_Core _
        "77 61 62 63 64 65 66 67 68 69 6A 6B 6C 6D 6E 6F" & _
        "70 71 72 73 74 75 76 77", _
        "abcdefghijklmnopqrstuvw"
End Sub

Private Sub Test_Cbor_Str8_TestCases()
    Debug.Print "Target: Str8"
    
    Test_Cbor_Str_Core _
        "78 18 61 62 63 64 65 66 67 68 69 6A 6B 6C 6D 6E 6F" & _
        "70 71 72 73 74 75 76 77 78", _
        "abcdefghijklmnopqrstuvwx"
    Test_Cbor_Str_Core2 "78 FF", &HFF
End Sub

Private Sub Test_Cbor_Str16_TestCases()
    Debug.Print "Target: Str16"
    
    Test_Cbor_Str_Core2 "79 01 00", &H100
    Test_Cbor_Str_Core2 "79 FF FF", &HFFFF&
End Sub

Private Sub Test_Cbor_Str32_TestCases()
    Debug.Print "Target: Str32"
    
    Test_Cbor_Str_Core2 "7A 00 01 00 00", &H10000
End Sub

Public Sub Test_Cbor_FixArray_TestCases()
    Debug.Print "Target: FixArray"
    
    Test_Cbor_Array_Core "80"
    Test_Cbor_Array_Core "81 01"
    Test_Cbor_Array_Core2 "97", &H17
End Sub

Public Sub Test_Cbor_Array8_TestCases()
    Debug.Print "Target: Array8"
    
    Test_Cbor_Array_Core2 "98 18", &H18
    Test_Cbor_Array_Core2 "98 FF", &HFF
End Sub

Public Sub Test_Cbor_Array16_TestCases()
    Debug.Print "Target: Array16"
    
    Test_Cbor_Array_Core2 "99 01 00", &H100
    'Test_Cbor_Array_Core2 "99 FF FF", &HFFFF&
End Sub

Public Sub Test_Cbor_Array32_TestCases()
    'Debug.Print "Target: Array32"
    
    'Test_Cbor_Array_Core2 "9A 00 01 00 00", &H10000
End Sub

Public Sub Test_Cbor_FixMap_TestCases()
    Debug.Print "Target: FixMap"
    
    Test_Cbor_Map_Core "A0"
    Test_Cbor_Map_Core "A1 61 61 00"
    Test_Cbor_Map_Core2 "B7", &H17
End Sub

Public Sub Test_Cbor_Map8_TestCases()
    Debug.Print "Target: Map8"
    
    Test_Cbor_Map_Core2 "B8 18", &H18
    Test_Cbor_Map_Core2 "B8 FF", &HFF
End Sub

Public Sub Test_Cbor_Map16_TestCases()
    Debug.Print "Target: Map16"
    
    Test_Cbor_Map_Core2 "B9 01 00", &H100
    'Test_Cbor_Map_Core2 "B9 FF FF", &HFFFF&
End Sub

Public Sub Test_Cbor_Map32_TestCases()
    Debug.Print "Target: Map32"
    
    'Test_Cbor_Map_Core2 "B9 00 01 00 00", &H10000
End Sub

Private Sub Test_Cbor_False_TestCases()
    Debug.Print "Target: False"
    
    Test_Cbor_Bool_Core "F4", False
End Sub

Private Sub Test_Cbor_True_TestCases()
    Debug.Print "Target: True"
    
    Test_Cbor_Bool_Core "F5", True
End Sub

Private Sub Test_Cbor_Null_TestCases()
    Debug.Print "Target: Null"
    Test_Cbor_Null_Core "F6", Null
End Sub

Private Sub Test_Cbor_Undefined_TestCases()
    Debug.Print "Target: Undefined"
    Test_Cbor_Undefined_Core "F7", Empty
End Sub

Private Sub Test_Cbor_Float32_TestCases()
    Debug.Print "Target: Float32"
    
    Test_Cbor_Float_Core "FA 41 46 00 00", 12.375!
    Test_Cbor_Float_Core "FA 3F 80 00 00", 1!
    Test_Cbor_Float_Core "FA 3F 00 00 00", 0.5
    Test_Cbor_Float_Core "FA 3E C0 00 00", 0.375
    Test_Cbor_Float_Core "FA 3E 80 00 00", 0.25
    Test_Cbor_Float_Core "FA BF 80 00 00", -1!
    
    ' Positive Zero
    Test_Cbor_Float_Core "FA 00 00 00 00", 0!
    
    ' Positive SubNormal Minimum
    Test_Cbor_Float_Core "FA 00 00 00 01", 1.401298E-45
    
    ' Positive SubNormal Maximum
    Test_Cbor_Float_Core "FA 00 7F FF FF", 1.175494E-38
    
    ' Positive Normal Minimum
    Test_Cbor_Float_Core "FA 00 80 00 00", 1.175494E-38
    
    ' Positive Normal Maximum
    Test_Cbor_Float_Core "FA 7F 7F FF FF", 3.402823E+38
    
    ' Positive Infinity
    Test_Cbor_Float_Core "FA 7F 80 00 00", "inf"
    
    ' Positive NaN
    Test_Cbor_Float_Core "FA 7F FF FF FF", "nan"
    
    ' Negative Zero
    Test_Cbor_Float_Core "FA 80 00 00 00", -0!
    
    ' Negative SubNormal Minimum
    Test_Cbor_Float_Core "FA 80 00 00 01", -1.401298E-45
    
    ' Negative SubNormal Maximum
    Test_Cbor_Float_Core "FA 80 7F FF FF", -1.175494E-38
    
    ' Negative Normal Minimum
    Test_Cbor_Float_Core "FA 80 80 00 00", -1.175494E-38
    
    ' Negative Normal Maximum
    Test_Cbor_Float_Core "FA FF 7F FF FF", -3.402823E+38
    
    ' Negative Infinity
    Test_Cbor_Float_Core "FA FF 80 00 00", "-inf"
    
    ' Negative NaN
    Test_Cbor_Float_Core "FA FF FF FF FF", "-nan"
End Sub

Private Sub Test_Cbor_Float64_TestCases()
    Debug.Print "Target: Float64"
    
    Test_Cbor_Float_Core "FB 40 28 C0 00 00 00 00 00", 12.375
    Test_Cbor_Float_Core "FB 3F F0 00 00 00 00 00 00", 1#
    Test_Cbor_Float_Core "FB 3F E0 00 00 00 00 00 00", 0.5
    Test_Cbor_Float_Core "FB 3F D8 00 00 00 00 00 00", 0.375
    Test_Cbor_Float_Core "FB 3F D0 00 00 00 00 00 00", 0.25
    Test_Cbor_Float_Core "FB 3F B9 99 99 99 99 99 9A", 0.1
    Test_Cbor_Float_Core "FB 3F D5 55 55 55 55 55 55", 1# / 3#
    Test_Cbor_Float_Core "FB BF F0 00 00 00 00 00 00", -1#
    
    ' Positive Zero
    Test_Cbor_Float_Core "FB 00 00 00 00 00 00 00 00", 0#
    
    ' Positive SubNormal Minimum
    Test_Cbor_Float_Core "FB 00 00 00 00 00 00 00 01", _
        4.94065645841247E-324
    
    ' Positive SubNormal Maximum
    Test_Cbor_Float_Core "FB 00 0F FF FF FF FF FF FF", _
        2.2250738585072E-308
    
    ' Positive Normal Minimum
    Test_Cbor_Float_Core "FB 00 10 00 00 00 00 00 00", _
        2.2250738585072E-308
    
    ' Positive Normal Maximum
    Test_Cbor_Float_Core "FB 7F EF FF FF FF FF FF FF", _
        "1.79769313486232E+308"
    
    ' Positive Infinity
    Test_Cbor_Float_Core "FB 7F F0 00 00 00 00 00 00", "inf"
    
    ' Positive NaN
    Test_Cbor_Float_Core "FB 7F FF FF FF FF FF FF FF", "nan"
    
    ' Negative Zero
    Test_Cbor_Float_Core "FB 80 00 00 00 00 00 00 00", -0#
    
    ' Negative SubNormal Minimum
    Test_Cbor_Float_Core "FB 80 00 00 00 00 00 00 01", _
        -4.94065645841247E-324
    
    ' Negative SubNormal Maximum
    Test_Cbor_Float_Core "FB 80 0F FF FF FF FF FF FF", _
        -2.2250738585072E-308
    
    ' Negative Normal Minimum
    Test_Cbor_Float_Core "FB 80 10 00 00 00 00 00 00", _
        -2.2250738585072E-308
    
    ' Negative Normal Maximum
    Test_Cbor_Float_Core "FB FF EF FF FF FF FF FF FF", _
        "-1.79769313486232E+308"
    
    ' Negative Infinity
    Test_Cbor_Float_Core "FB FF F0 00 00 00 00 00 00", "-inf"
    
    ' Negative NaN
    Test_Cbor_Float_Core "FB FF FF FF FF FF FF FF FF", "-nan"
End Sub

'
' CBOR for VBA - Test Core
'

Private Sub Test_Cbor_Int_Core(HexBE As String, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = CBOR.GetValue(BytesBE)
    
    DebugPrint_Int_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputCBBytesBE() As Byte
    OutputCBBytesBE = CBOR.GetCborBytes(OutputValue)
    
    DebugPrint_Int_GetBytes OutputValue, OutputCBBytesBE, BytesBE
End Sub

Private Sub Test_Cbor_Bin_Core(HexStr As String, ExpectedHexStr As String)
    Dim Bytes() As Byte
    Bytes = GetBytesFromHexString(HexStr)
    
    Dim ExpectedValue() As Byte
    ExpectedValue = GetBytesFromHexString(ExpectedHexStr)
    
    Dim OutputValue() As Byte
    OutputValue = CBOR.GetValue(Bytes)
    
    DebugPrint_Bin_GetValue Bytes, OutputValue, ExpectedValue
    
    Dim OutputCBBytes() As Byte
    OutputCBBytes = CBOR.GetCborBytes(OutputValue)
    
    DebugPrint_Bin_GetBytes OutputValue, OutputCBBytes, Bytes
End Sub

Private Sub Test_Cbor_Str_Core(HexStr As String, ExpectedValue As String)
    Dim Bytes() As Byte
    Bytes = GetBytesFromHexString(HexStr)
    
    Dim OutputValue As String
    OutputValue = CBOR.GetValue(Bytes)
    
    DebugPrint_Str_GetValue Bytes, OutputValue, ExpectedValue
    
    Dim OutputBytes() As Byte
    OutputBytes = CBOR.GetCborBytes(OutputValue)
    
    DebugPrint_Str_GetBytes OutputValue, OutputBytes, Bytes
End Sub

Public Sub Test_Cbor_Array_Core(HexStr As String)
    Dim Bytes() As Byte
    Bytes = GetBytesFromHexString(HexStr)
    
    #If USE_COLLECTION Then
    Dim ExpectedDummy As Collection
    Set ExpectedDummy = New Collection
    
    Dim OutputValue As Collection
    Set OutputValue = CBOR.GetValue(Bytes)
    #Else
    Dim ExpectedDummy
    ExpectedDummy = Array(Empty)
    
    Dim OutputValue
    OutputValue = CBOR.GetValue(Bytes)
    #End If
    
    DebugPrint_Array_GetValue Bytes, OutputValue, ExpectedDummy
    
    Dim OutputCBBytes() As Byte
    OutputCBBytes = CBOR.GetCborBytes(OutputValue)
    
    DebugPrint_Array_GetBytes OutputValue, OutputCBBytes, Bytes
End Sub

Public Sub Test_Cbor_Map_Core(HexStr As String)
    Dim Bytes() As Byte
    Bytes = GetBytesFromHexString(HexStr)
    
    Dim ExpectedDummy As Object
    Set ExpectedDummy = CreateObject("Scripting.Dictionary")
    
    Dim OutputValue As Object
    Set OutputValue = CBOR.GetValue(Bytes)
    
    DebugPrint_Map_GetValue Bytes, OutputValue, ExpectedDummy
    
    Dim OutputCBBytes() As Byte
    OutputCBBytes = CBOR.GetCborBytes(OutputValue)
    
    DebugPrint_Map_GetBytes OutputValue, OutputCBBytes, Bytes
End Sub

Private Sub Test_Cbor_Bool_Core(HexBE As String, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = CBOR.GetValue(BytesBE)
    
    DebugPrint_Bool_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputCBBytesBE() As Byte
    OutputCBBytesBE = CBOR.GetCborBytes(OutputValue)
    
    DebugPrint_Bool_GetBytes OutputValue, OutputCBBytesBE, BytesBE
End Sub

Private Sub Test_Cbor_Null_Core(HexBE As String, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = CBOR.GetValue(BytesBE)
    
    DebugPrint_Null_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputCBBytesBE() As Byte
    OutputCBBytesBE = CBOR.GetCborBytes(OutputValue)
    
    DebugPrint_Null_GetBytes OutputValue, OutputCBBytesBE, BytesBE
End Sub

Private Sub Test_Cbor_Undefined_Core(HexBE As String, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = CBOR.GetValue(BytesBE)
    
    DebugPrint_Undefined_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputCBBytesBE() As Byte
    OutputCBBytesBE = CBOR.GetCborBytes(OutputValue)
    
    DebugPrint_Undefined_GetBytes OutputValue, OutputCBBytesBE, BytesBE
End Sub

Private Sub Test_Cbor_Float_Core(HexBE As String, ExpectedValue)
    Dim BytesBE() As Byte
    BytesBE = GetBytesFromHexString(HexBE)
    
    Dim OutputValue
    OutputValue = CBOR.GetValue(BytesBE)
    
    DebugPrint_Float_GetValue BytesBE, OutputValue, ExpectedValue
    
    Dim OutputCBBytesBE() As Byte
    OutputCBBytesBE = CBOR.GetCborBytes(OutputValue)
    
    DebugPrint_Float_GetBytes OutputValue, OutputCBBytesBE, BytesBE
End Sub

'
' CBOR for VBA - Test Core - Binary
'

Private Sub Test_Cbor_Bin_Core2(HeadHex As String, DataLength As Long)
    Dim HeadBytes() As Byte
    HeadBytes = GetBytesFromHexString(HeadHex)
    
    Dim CBBytes() As Byte
    CBBytes = GetTestBinBytes(HeadBytes, DataLength)
    
    Dim ExpectedValue() As Byte
    ExpectedValue = GetTestBinValue(DataLength)
    
    Dim OutputValue() As Byte
    OutputValue = CBOR.GetValue(CBBytes)
    
    DebugPrint_Bin_GetValue CBBytes, OutputValue, ExpectedValue
    
    Dim OutputBytes() As Byte
    OutputBytes = CBOR.GetCborBytes(OutputValue)
    
    DebugPrint_Bin_GetBytes OutputValue, OutputBytes, CBBytes
End Sub

Private Function GetTestBinValue(Length As Long) As Byte()
    Dim TestValue() As Byte
    ReDim TestValue(0 To Length - 1)
    
    Dim Index As Long
    For Index = 1 To Length
        TestValue(Index - 1) = Index Mod 256
    Next
    
    GetTestBinValue = TestValue
End Function

Private Function GetTestBinBytes( _
    HeadBytes() As Byte, BodyLength As Long) As Byte()
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(HeadBytes)
    UB = UBound(HeadBytes)
    
    Dim HeadLength As Long
    HeadLength = UB - LB + 1
    
    Dim TestBytes() As Byte
    ReDim TestBytes(HeadLength + BodyLength - 1)
    
    Dim Index As Long
    For Index = 0 To HeadLength - 1
        TestBytes(Index) = HeadBytes(LB + Index)
    Next
    For Index = 1 To BodyLength
        TestBytes(HeadLength + Index - 1) = Index Mod 256
    Next
    
    GetTestBinBytes = TestBytes
End Function

'
' CBOR for VBA - Test Core - String
'

Private Sub Test_Cbor_Str_Core2(HeadHex As String, DataLength As Long)
    Dim HeadBytes() As Byte
    HeadBytes = GetBytesFromHexString(HeadHex)
    
    Dim Bytes() As Byte
    Bytes = GetTestStrBytes(HeadBytes, DataLength)
    
    Dim ExpectedValue As String
    ExpectedValue = GetTestStr(DataLength)
    
    Dim OutputValue As String
    OutputValue = CBOR.GetValue(Bytes)
    
    DebugPrint_Str_GetValue Bytes, OutputValue, ExpectedValue
    
    Dim OutputBytes() As Byte
    OutputBytes = CBOR.GetCborBytes(OutputValue)
    
    DebugPrint_Str_GetBytes OutputValue, OutputBytes, Bytes
End Sub

Private Function GetTestStr(Length As Long) As String
    Dim TestStr As String
    
    Dim Index As Long
    For Index = 1 To Length
        TestStr = TestStr & Hex(Index Mod 16)
    Next
    
    GetTestStr = TestStr
End Function

Private Function GetTestStrBytes( _
    HeadBytes() As Byte, BodyLength As Long) As Byte()
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(HeadBytes)
    UB = UBound(HeadBytes)
    
    Dim HeadLength As Long
    HeadLength = UB - LB + 1
    
    Dim TestBytes() As Byte
    ReDim TestBytes(HeadLength + BodyLength - 1)
    
    Dim Index As Long
    For Index = 0 To HeadLength - 1
        TestBytes(Index) = HeadBytes(LB + Index)
    Next
    For Index = 1 To BodyLength
        TestBytes(HeadLength + Index - 1) = Asc(Hex(Index Mod 16))
    Next
    
    GetTestStrBytes = TestBytes
End Function

'
' CBOR for VBA - Test Core - Array
'

Public Sub Test_Cbor_Array_Core2(HeadHex As String, ElementCount As Long)
    Dim HeadBytes() As Byte
    HeadBytes = GetBytesFromHexString(HeadHex)
    
    Dim Bytes() As Byte
    Bytes = GetTestArrayBytes(HeadBytes, ElementCount)
    
    #If USE_COLLECTION Then
    Dim ExpectedDummy As Collection
    Set ExpectedDummy = New Collection
    
    Dim OutputValue As Collection
    Set OutputValue = CBOR.GetValue(Bytes)
    #Else
    Dim ExpectedDummy
    ExpectedDummy = Array(Empty)
    
    Dim OutputValue
    OutputValue = CBOR.GetValue(Bytes)
    #End If
    
    DebugPrint_Array_GetValue Bytes, OutputValue, ExpectedDummy
    
    Dim OutputCBBytes() As Byte
    OutputCBBytes = CBOR.GetCborBytes(OutputValue)
    
    DebugPrint_Array_GetBytes OutputValue, OutputCBBytes, Bytes
End Sub

Private Function GetTestArrayBytes( _
    HeadBytes() As Byte, ElementCount As Long) As Byte()
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(HeadBytes)
    UB = UBound(HeadBytes)
    
    Dim HeadLength As Long
    HeadLength = UB - LB + 1
    
    Dim TestBytes() As Byte
    ReDim TestBytes(0 To HeadLength - 1)
    
    Dim Index As Long
    For Index = 0 To HeadLength - 1
        TestBytes(Index) = HeadBytes(LB + Index)
    Next
    For Index = 1 To ElementCount
        AddBytes TestBytes, CBOR.GetCborBytes(Index)
    Next
    
    GetTestArrayBytes = TestBytes
End Function

'
' CBOR for VBA - Test Core - Map
'

Public Sub Test_Cbor_Map_Core2(HeadHex As String, ElementCount As Long)
    Dim HeadBytes() As Byte
    HeadBytes = GetBytesFromHexString(HeadHex)
    
    Dim Bytes() As Byte
    Bytes = GetTestMapBytes(HeadBytes, ElementCount)
    
    Dim ExpectedDummy As Object
    Set ExpectedDummy = CreateObject("Scripting.Dictionary")
    
    Dim OutputValue As Object
    Set OutputValue = CBOR.GetValue(Bytes)
    
    DebugPrint_Map_GetValue Bytes, OutputValue, ExpectedDummy
    
    Dim OutputCBBytes() As Byte
    OutputCBBytes = CBOR.GetCborBytes(OutputValue)
    
    DebugPrint_Map_GetBytes OutputValue, OutputCBBytes, Bytes
End Sub

Private Function GetTestMapBytes( _
    HeadBytes() As Byte, ElementCount As Long) As Byte()
    
    Dim LB As Long
    Dim UB As Long
    LB = LBound(HeadBytes)
    UB = UBound(HeadBytes)
    
    Dim HeadLength As Long
    HeadLength = UB - LB + 1
    
    Dim TestBytes() As Byte
    ReDim TestBytes(0 To HeadLength - 1)
    
    Dim Index As Long
    For Index = 0 To HeadLength - 1
        TestBytes(Index) = HeadBytes(LB + Index)
    Next
    For Index = 1 To ElementCount
        AddBytes TestBytes, CBOR.GetCborBytes("key-" & CStr(Index))
        AddBytes TestBytes, CBOR.GetCborBytes("value-" & CStr(Index))
    Next
    
    GetTestMapBytes = TestBytes
End Function

'
' CBOR for VBA - Test - Debug.Print - Integer
'

Private Sub DebugPrint_Int_GetBytes( _
    Value, OutputCBBytes() As Byte, ExpectedCBBytes() As Byte)
    
    If VarType(Value) = vbDecimal Then
        DebugPrint_GetCborBytes CStr(Value), OutputCBBytes, ExpectedCBBytes
    Else
        DebugPrint_GetCborBytes _
            CStr(Value) & " (" & Hex(Value) & ")", _
            OutputCBBytes, ExpectedCBBytes
    End If
End Sub

Private Sub DebugPrint_Int_GetValue( _
    CBBytes() As Byte, OutputValue, ExpectedValue)
    
    If (VarType(OutputValue) = vbDecimal) Or _
        (VarType(ExpectedValue) = vbDecimal) Then
        
        DebugPrint_GetValue CBBytes, OutputValue, ExpectedValue, _
            CStr(OutputValue), CStr(ExpectedValue)
    Else
        DebugPrint_GetValue CBBytes, OutputValue, ExpectedValue, _
            CStr(OutputValue) & " (" & Hex(OutputValue) & ")", _
            CStr(ExpectedValue) & " (" & Hex(ExpectedValue) & ")"
    End If
End Sub

'
' CBOR for VBA - Test - Debug.Print - Binary
'

Private Sub DebugPrint_Bin_GetBytes( _
    Value, OutputCBBytes() As Byte, ExpectedCBBytes() As Byte)
    
    Dim HexString As String
    HexString = GetHexStringFromBytes(Value, , , " ")
    
    DebugPrint_GetCborBytes HexString, OutputCBBytes, ExpectedCBBytes
End Sub

Private Sub DebugPrint_Bin_GetValue( _
    CBBytes() As Byte, OutputValue() As Byte, ExpectedValue() As Byte)
    
    Dim OutputHexString As String
    OutputHexString = GetHexStringFromBytes(OutputValue, , , " ")
    
    Dim ExpectedHexString As String
    ExpectedHexString = GetHexStringFromBytes(ExpectedValue, , , " ")
    
    DebugPrint_GetValue CBBytes, OutputHexString, ExpectedHexString, _
        OutputHexString, ExpectedHexString
End Sub

'
' CBOR for VBA - Test - Debug.Print - String
'

Private Sub DebugPrint_Str_GetBytes( _
    Value, OutputCBBytes() As Byte, ExpectedCBBytes() As Byte)
    
    DebugPrint_GetCborBytes Value, OutputCBBytes, ExpectedCBBytes
End Sub

Private Sub DebugPrint_Str_GetValue( _
    CBBytes() As Byte, OutputValue, ExpectedValue)
    
    DebugPrint_GetValue CBBytes, OutputValue, ExpectedValue, _
        OutputValue, ExpectedValue
End Sub

'
' CBOR for VBA - Test - Debug.Print - Array
'

Private Sub DebugPrint_Array_GetBytes( _
    Value, OutputCBBytes() As Byte, ExpectedCBBytes() As Byte)
    
    DebugPrint_GetCborBytes _
        "(" & TypeName(Value) & ")", OutputCBBytes, ExpectedCBBytes
End Sub

Private Sub DebugPrint_Array_GetValue( _
    CBBytes() As Byte, OutputValue, ExpectedValue)
    
    Dim OutputDummy As String
    OutputDummy = "(" & TypeName(OutputValue) & ")"
    
    Dim ExpectedDummy As String
    ExpectedDummy = "(" & TypeName(ExpectedValue) & ")"
    
    DebugPrint_GetValue CBBytes, _
        OutputDummy, ExpectedDummy, _
        OutputDummy, ExpectedDummy
End Sub

'
' CBOR for VBA - Test - Debug.Print - Map
'

Private Sub DebugPrint_Map_GetBytes( _
    Value, OutputCBBytes() As Byte, ExpectedCBBytes() As Byte)
    
    DebugPrint_GetCborBytes _
        "(" & TypeName(Value) & ")", OutputCBBytes, ExpectedCBBytes
End Sub

Private Sub DebugPrint_Map_GetValue( _
    CBBytes() As Byte, OutputValue, ExpectedValue)
    
    Dim OutputDummy As String
    OutputDummy = "(" & TypeName(OutputValue) & ")"
    
    Dim ExpectedDummy As String
    ExpectedDummy = "(" & TypeName(ExpectedValue) & ")"
    
    DebugPrint_GetValue CBBytes, _
        OutputDummy, ExpectedDummy, _
        OutputDummy, ExpectedDummy
End Sub

'
' CBOR for VBA - Test - Debug.Print - Boolean
'

Private Sub DebugPrint_Bool_GetBytes( _
    Value, OutputCBBytes() As Byte, ExpectedCBBytes() As Byte)
    
    DebugPrint_GetCborBytes CStr(Value), OutputCBBytes, ExpectedCBBytes
End Sub

Private Sub DebugPrint_Bool_GetValue( _
    CBBytes() As Byte, OutputValue, ExpectedValue)
    
    DebugPrint_GetValue CBBytes, OutputValue, ExpectedValue, _
        CStr(OutputValue), CStr(ExpectedValue)
End Sub

'
' CBOR for VBA - Test - Debug.Print - Null
'

Private Sub DebugPrint_Null_GetBytes( _
    Value, OutputCBBytes() As Byte, ExpectedCBBytes() As Byte)
    
    DebugPrint_GetCborBytes _
        IIf(IsNull(Value), "Null", "not Null"), OutputCBBytes, ExpectedCBBytes
End Sub

Private Sub DebugPrint_Null_GetValue( _
    CBBytes() As Byte, OutputValue, ExpectedValue)
    
    DebugPrint_GetValue CBBytes, _
        IIf(IsNull(OutputValue), "Null", "not Null"), _
        IIf(IsNull(ExpectedValue), "Null", "not Null"), _
        IIf(IsNull(OutputValue), "Null", "not Null"), _
        IIf(IsNull(ExpectedValue), "Null", "not Null")
End Sub

'
' CBOR for VBA - Test - Debug.Print - Undefined
'

Private Sub DebugPrint_Undefined_GetBytes( _
    Value, OutputCBBytes() As Byte, ExpectedCBBytes() As Byte)
    
    DebugPrint_GetCborBytes _
        IIf(IsEmpty(Value), "Empty", "not Empty"), _
        OutputCBBytes, ExpectedCBBytes
End Sub

Private Sub DebugPrint_Undefined_GetValue( _
    CBBytes() As Byte, OutputValue, ExpectedValue)
    
    DebugPrint_GetValue CBBytes, _
        IIf(IsEmpty(OutputValue), "Empty", "not Empty"), _
        IIf(IsEmpty(ExpectedValue), "Empty", "not Empty"), _
        IIf(IsEmpty(OutputValue), "Empty", "not Empty"), _
        IIf(IsEmpty(ExpectedValue), "Empty", "not Empty")
End Sub

'
' CBOR for VBA - Test - Debug.Print - Float
'

Private Sub DebugPrint_Float_GetBytes( _
    Value, OutputCBBytes() As Byte, ExpectedCBBytes() As Byte)
    
    DebugPrint_GetCborBytes CStr(Value), OutputCBBytes, ExpectedCBBytes
End Sub

Private Sub DebugPrint_Float_GetValue( _
    CBBytes() As Byte, OutputValue, ExpectedValue)
    
    DebugPrint_GetValue CBBytes, _
        CStr(OutputValue), CStr(ExpectedValue), _
        CStr(OutputValue), CStr(ExpectedValue)
End Sub

''
'' CBOR for VBA - Test Counter
''

Private Property Get Test_Count() As Long
    Test_Count = m_Test_Count
End Property

Private Sub Test_Initialize()
    m_Test_Count = 0
    m_Test_Success = 0
    m_Test_Fail = 0
End Sub

Private Sub Test_Countup(bSuccess As Boolean)
    m_Test_Count = m_Test_Count + 1
    If bSuccess Then
        m_Test_Success = m_Test_Success + 1
    Else
        m_Test_Fail = m_Test_Fail + 1
    End If
End Sub

Private Sub Test_Terminate()
    Debug.Print _
        "Count: " & CStr(m_Test_Count) & ", " & _
        "Success: " & CStr(m_Test_Success) & ", " & _
        "Fail: " & CStr(m_Test_Fail)
End Sub

''
'' CBOR for VBA - Test - Debug.Print
''

Private Sub DebugPrint_GetCborBytes( _
    Source, OutputCBBytes() As Byte, ExpectedCBBytes() As Byte)
    
    Dim bSuccess As Boolean
    bSuccess = CompareBytes(OutputCBBytes, ExpectedCBBytes)
    
    Test_Countup bSuccess
    
    Dim OutputCBBytesStr As String
    OutputCBBytesStr = GetHexStringFromBytes(OutputCBBytes, , , " ")
    
    Dim ExpectedCBBytesStr As String
    ExpectedCBBytesStr = GetHexStringFromBytes(ExpectedCBBytes, , , " ")
    
    Debug.Print "No." & CStr(Test_Count) & _
        " Result: " & IIf(bSuccess, "OK", "NG") & _
        " Source: " & Source & _
        " Output: " & OutputCBBytesStr & _
        " Expect: " & ExpectedCBBytesStr
End Sub

Private Sub DebugPrint_GetValue( _
    CBBytes() As Byte, OutputValue, ExpectedValue, Output, Expect)
    
    Dim bSuccess As Boolean
    bSuccess = (OutputValue = ExpectedValue)
    
    Test_Countup bSuccess
    
    Dim CBBytesStr As String
    CBBytesStr = GetHexStringFromBytes(CBBytes, , , " ")
    
    Debug.Print "No." & CStr(Test_Count) & _
        " Result: " & IIf(bSuccess, "OK", "NG") & _
        " Source: " & CBBytesStr & _
        " Output: " & Output & _
        " Expect: " & Expect
End Sub

''
'' CBOR for VBA - Test - Byte Array Helper
''

Private Function CompareBytes(Bytes1() As Byte, Bytes2() As Byte) As Boolean
    Dim LB1 As Long
    Dim UB1 As Long
    LB1 = LBound(Bytes1)
    UB1 = UBound(Bytes1)
    
    Dim LB2 As Long
    Dim UB2 As Long
    LB2 = LBound(Bytes2)
    UB2 = UBound(Bytes2)
    
    If (UB1 - LB1 + 1) <> (UB2 - LB2 + 1) Then Exit Function
    
    Dim Index As Long
    For Index = 0 To UB1 - LB1
        If Bytes1(LB1 + Index) <> Bytes2(LB2 + Index) Then Exit Function
    Next
    
    CompareBytes = True
End Function

''
'' CBOR for VBA - Test - Hex String
''

Private Function GetBytesFromHexString(ByVal Value As String) As Byte()
    Dim Value_ As String
    Dim Index As Long
    For Index = 1 To Len(Value)
        Select Case Mid(Value, Index, 1)
        Case "0" To "9", "A" To "F", "a" To "f"
            Value_ = Value_ & Mid(Value, Index, 1)
        End Select
    Next
    
    Dim Length As Long
    Length = Len(Value_) \ 2
    
    Dim Bytes() As Byte
    
    If Length = 0 Then
        GetBytesFromHexString = Bytes
        Exit Function
    End If
    
    ReDim Bytes(0 To Length - 1)
    
    'Dim Index As Long
    For Index = 0 To Length - 1
        Bytes(Index) = CByte("&H" & Mid(Value_, 1 + Index * 2, 2))
    Next
    
    GetBytesFromHexString = Bytes
End Function

'Private Function GetHexStringFromBytes(Bytes() As Byte,
Private Function GetHexStringFromBytes(Bytes, _
    Optional Index As Long, Optional Length As Long, _
    Optional Separator As String) As String
    
    If Length = 0 Then
        On Error Resume Next
        Length = UBound(Bytes) - Index + 1
        On Error GoTo 0
    End If
    If Length = 0 Then
        GetHexStringFromBytes = ""
        Exit Function
    End If
    
    Dim HexString As String
    HexString = Right("0" & Hex(Bytes(Index)), 2)
    
    Dim Offset As Long
    For Offset = 1 To Length - 1
        HexString = _
            HexString & Separator & Right("0" & Hex(Bytes(Index + Offset)), 2)
    Next
    
    GetHexStringFromBytes = HexString
End Function

''
'' CBOR for VBA - Test - Bytes Operator
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

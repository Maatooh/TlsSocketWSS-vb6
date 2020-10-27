Attribute VB_Name = "mUTF8"
Option Explicit

'UTF-8
Private Declare Function WideCharToMultiByte Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long, ByRef lpMultiByteStr As Any, ByVal cchMultiByte As Long, ByVal lpDefaultChar As String, ByVal lpUsedDefaultChar As Long) As Long
Private Declare Function MultiByteToWideChar Lib "kernel32" (ByVal CodePage As Long, ByVal dwFlags As Long, ByVal lpMultiByteStr As Long, ByVal cchMultiByte As Long, ByVal lpWideCharStr As Long, ByVal cchWideChar As Long) As Long
Private Const CP_UTF8 = 65001
'utf8תunicode
Public Function ToUnicodeData(ByRef Utf() As Byte) As Byte()
    Dim lret As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    Dim BT() As Byte
    lLength = UBound(Utf) + 1
    If lLength <= 0 Then Exit Function
    lBufferSize = lLength * 2 - 1
    ReDim BT(lBufferSize)
    lret = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf(0)), lLength, VarPtr(BT(0)), lBufferSize + 1)
    If lret <> 0 Then
        ReDim Preserve BT(lret - 1)
        ToUnicodeData = BT
    End If
End Function
'utf8תunicode
Public Function ToUnicodeString(ByRef Utf() As Byte) As String
    Dim lret As Long
    Dim lLength As Long
    Dim lBufferSize As Long
    On Error GoTo errline:
    lLength = UBound(Utf) + 1
    If lLength <= 0 Then Exit Function
    lBufferSize = lLength * 2
    ToUnicodeString = String$(lBufferSize, Chr(0))
    lret = MultiByteToWideChar(CP_UTF8, 0, VarPtr(Utf(0)), lLength, StrPtr(ToUnicodeString), lBufferSize)
    If lret <> 0 Then
        ToUnicodeString = Left(ToUnicodeString, lret)
    End If
    Exit Function
errline:
    ToUnicodeString = ""
End Function
'unicodeתutf8
Public Function Encoding(ByVal UCS As String) As Byte()
    Dim lLength As Long
    Dim lBufferSize As Long
    Dim lResult As Long
    Dim abUTF8() As Byte
    lLength = Len(UCS)
    If lLength = 0 Then Exit Function
    lBufferSize = lLength * 3 + 1
    ReDim abUTF8(lBufferSize - 1)
    lResult = WideCharToMultiByte(CP_UTF8, 0, StrPtr(UCS), lLength, abUTF8(0), lBufferSize, vbNullString, 0)
    If lResult <> 0 Then
    lResult = lResult - 1
    ReDim Preserve abUTF8(lResult)
    Encoding = abUTF8
    End If
End Function

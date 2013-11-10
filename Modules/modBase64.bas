Attribute VB_Name = "modBase64"

' **********************************************************************
'  描  述：21种加密54种压缩 算法模块 海阔天空收集整理
'  Play78.com : 网站导航，源码之家，绝对开源
'  海阔天空整理,有问题请上www.paly78.com 提
'  网址：http://www.play78.com/
'  QQ：13355575
'  e-mail:hglai@eyou.com
' **********************************************************************


Option Explicit

Private Const Base64 = "ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz0123456789+/"

Public Sub Base64Array_Encode(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim InpPos As Long
    Dim c1, c2, c3 As Integer
    ReDim OutStream(500)
    'If OutPos = 0 Then Exit Sub
    InpPos = 0
    OutPos = 0
    Do While InpPos <= UBound(ByteArray)
        c1 = ReadValue(ByteArray, InpPos)
        c2 = ReadValue(ByteArray, InpPos)
        c3 = ReadValue(ByteArray, InpPos)
        Call AddValueToStream(OutStream, OutPos, mimeencode(Int(c1 / 4)))
        Call AddValueToStream(OutStream, OutPos, mimeencode((c1 And 3) * 16 + Int(c2 / 16)))
        If InpPos - 2 <= UBound(ByteArray) Then
            Call AddValueToStream(OutStream, OutPos, mimeencode((c2 And 15) * 4 + Int(c3 / 64)))
        End If
        If InpPos - 1 <= UBound(ByteArray) Then
            Call AddValueToStream(OutStream, OutPos, mimeencode(c3 And 63))
        End If
    Loop
    ReDim ByteArray(OutPos - 1)
    Call CopyMemory(VarPtr(ByteArray(0)), VarPtr(OutStream(0)), OutPos)
End Sub

Public Sub Base64Array_Decode(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim InpPos As Long
    Dim w1, w2, w3, w4 As Integer
    ReDim OutStream(500)
    InpPos = 0
    OutPos = 0
    Do While InpPos < UBound(ByteArray)
        w1 = mimedecode(ReadValue(ByteArray, InpPos))
        w2 = mimedecode(ReadValue(ByteArray, InpPos))
        w3 = mimedecode(ReadValue(ByteArray, InpPos))
        w4 = mimedecode(ReadValue(ByteArray, InpPos))
        If w2 >= 0 Then Call AddValueToStream(OutStream, OutPos, (w1 * 4 + Int(w2 / 16)) And 255)
        If w3 >= 0 Then Call AddValueToStream(OutStream, OutPos, (w2 * 16 + Int(w3 / 4)) And 255)
        If w4 >= 0 Then Call AddValueToStream(OutStream, OutPos, (w3 * 64 + w4) And 255)
    Loop
    ReDim ByteArray(OutPos - 1)
    Call CopyMemory(VarPtr(ByteArray(0)), VarPtr(OutStream(0)), OutPos)
End Sub

Private Function mimeencode(w As Integer) As Byte
   If w >= 0 Then mimeencode = Asc(Mid$(Base64, w + 1, 1)) Else mimeencode = 0
End Function

Private Function mimedecode(A As Integer) As Integer
   If A = 0 Then mimedecode = -1: Exit Function
   mimedecode = InStr(Base64, Chr(A)) - 1
End Function

Private Function ReadValue(FromArray() As Byte, FromPos As Long) As Integer
    If FromPos <= UBound(FromArray) Then
        ReadValue = FromArray(FromPos)
    Else
        ReadValue = 0
    End If
    FromPos = FromPos + 1
End Function

Private Sub AddValueToStream(ToStream() As Byte, ToPos As Long, Number As Byte)
    If ToPos > UBound(ToStream) Then ReDim Preserve ToStream(ToPos + 100)
    ToStream(ToPos) = Number
    ToPos = ToPos + 1
End Sub

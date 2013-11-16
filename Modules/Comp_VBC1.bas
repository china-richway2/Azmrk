Attribute VB_Name = "Comp_VBC1"
Option Explicit
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Type BytePos
    Data() As Byte
    Position As Long
    Buffer As Integer
    BitPos As Integer
End Type
Private Stream(1) As BytePos    '0=vbc-code 1=bitstreams

Private ExtraLengthBits(15) As Integer
Private StartValLength(15) As Integer
Private Dictionary(255) As Byte
Private CharCount(256) As Long

Public Sub Compress_VBC_Dynamic(ByteArray() As Byte)
    Dim Char As Integer
    Dim NewFileLen As Long
    Dim x As Long
    Dim y As Long
    Call init_Dynamic_VBC
    For x = 0 To UBound(ByteArray)
        Char = ByteArray(x)
        Call Store_Char(Char)
        Call update_Model(Char)
    Next
'send EOF character
    Call Store_Char(256)
'lets fill the leftovers
    For x = 0 To 1
        Do While Stream(x).BitPos > 0
            Call AddBitsToStream(Stream(x), 0, 1)
        Loop
    Next
'Lets restore the bounderies
    For x = 0 To 1
        ReDim Preserve Stream(x).Data(Stream(x).Position - 1)
    Next
'whe calculate the new length of the new data
    NewFileLen = 0
    For x = 0 To 1
        NewFileLen = NewFileLen + UBound(Stream(x).Data) + 1
    Next
    ReDim ByteArray(NewFileLen + 3)
'here we store the compressed data
    NewFileLen = 0
    For x = 0 To 0
        ByteArray(NewFileLen) = Int(UBound(Stream(x).Data) / &H10000) And &HFF
        NewFileLen = NewFileLen + 1
        ByteArray(NewFileLen) = Int(UBound(Stream(x).Data) / &H100) And &HFF
        NewFileLen = NewFileLen + 1
        ByteArray(NewFileLen) = UBound(Stream(x).Data) And &HFF
        NewFileLen = NewFileLen + 1
    Next
    For x = 0 To 1
        For y = 0 To UBound(Stream(x).Data)
            ByteArray(NewFileLen) = Stream(x).Data(y)
            NewFileLen = NewFileLen + 1
        Next
    Next
End Sub

Public Sub DeCompress_VBC_Dynamic(ByteArray() As Byte)
    Dim OutStream() As Byte
    Dim OutPos As Long
    Dim InposCont As Long
    Dim InContBit As Integer
    Dim InposData As Long
    Dim InDataBit As Integer
    Dim Char As Integer
    Dim VBC_Code As Integer
    Dim x As Long
    ReDim OutStream(500)
    Call init_Dynamic_VBC
    InposCont = 0
    InposData = 0
    For x = 0 To 2
        InposData = CLng(InposData) * 256 + ByteArray(InposCont)
        InposCont = InposCont + 1
    Next
    InposData = InposData + InposCont + 1
    InContBit = 0
    InDataBit = 0
    OutPos = 0
    Do
        VBC_Code = ReadBitsFromArray(ByteArray, InposCont, InContBit, 4)
        Char = StartValLength(VBC_Code) + ReadBitsFromArray(ByteArray, InposData, InDataBit, ExtraLengthBits(VBC_Code))
        If Char = 256 Then Exit Do
    If Char = 255 Then
        Char = 255
    End If
        Char = Dictionary(Char)
        Call AddCharToArray(OutStream, OutPos, CByte(Char))
        Call update_Model(Char)
    Loop
    ReDim ByteArray(OutPos - 1)
    For x = 0 To OutPos - 1
        ByteArray(x) = OutStream(x)
    Next
End Sub

Private Sub Store_Char(Char As Integer)
    Dim VBC_Code As Integer         '0-15
    Dim ByteValue As Integer
    If Char = 256 Then
        ByteValue = Char
    Else
        For ByteValue = 0 To 255
            If Dictionary(ByteValue) = Char Then Exit For
        Next
    End If
    Dim x As Integer
    For VBC_Code = 1 To 15
        If StartValLength(VBC_Code) > ByteValue Then Exit For
    Next
    VBC_Code = VBC_Code - 1
    ByteValue = ByteValue - StartValLength(VBC_Code)
    Call AddBitsToStream(Stream(0), VBC_Code, 4)
    Call AddBitsToStream(Stream(1), ByteValue, ExtraLengthBits(VBC_Code))
End Sub

Private Sub init_Dynamic_VBC()
    Dim x As Integer
'                    bitsNeeded    from to char     gain/loss
    ExtraLengthBits(0) = 0         '0                  +4
    ExtraLengthBits(1) = 0         '1                  +4
    ExtraLengthBits(2) = 0         '2                  +4
    ExtraLengthBits(3) = 0         '3                  +4
    ExtraLengthBits(4) = 2         '4-7                +2
    ExtraLengthBits(5) = 2         '8-11               +2
    ExtraLengthBits(6) = 2         '12-15              +2
    ExtraLengthBits(7) = 2         '16-19              +2
    ExtraLengthBits(8) = 2         '20-23              +2
    ExtraLengthBits(9) = 2         '24-27              +2
    ExtraLengthBits(10) = 2        '28-31              +2
    ExtraLengthBits(11) = 2        '32-35              +2
    ExtraLengthBits(12) = 5        '36-67              -1
    ExtraLengthBits(13) = 6        '68-131             -2
    ExtraLengthBits(14) = 6        '132-195            -2
    ExtraLengthBits(15) = 6        '196-259            -2
    StartValLength(0) = 0
    For x = 1 To 15
        StartValLength(x) = StartValLength(x - 1) + (2 ^ ExtraLengthBits(x - 1))
    Next
    'StartValLength = 0 1 2 3 4 8 12 16 20 24 28 32 36 68 132 196
    For x = 0 To 255
        Dictionary(x) = x
        CharCount(x) = 0
    Next
    CharCount(256) = 0
    For x = 0 To 1
        With Stream(x)
            ReDim .Data(500)
            .BitPos = 0
            .Buffer = 0
            .Position = 0
        End With
    Next
End Sub

Private Sub update_Model(Char As Integer)
    Dim DictPos As Integer
    Dim OldPos As Integer
    Dim Temp As Long
    For DictPos = 0 To 255
        If Dictionary(DictPos) = Char Then Exit For
    Next
    DictPos = DictPos
    OldPos = DictPos
    If OldPos = 256 Then Stop
    CharCount(DictPos) = CharCount(DictPos) + 1
    Do While DictPos > 0
        If CharCount(DictPos) > CharCount(DictPos - 1) Then
            Temp = CharCount(DictPos - 1)
            CharCount(DictPos - 1) = CharCount(DictPos)
            CharCount(DictPos) = Temp
            DictPos = DictPos - 1
        Else
            Exit Do
        End If
    Loop
    If OldPos = DictPos Then Exit Sub
    Dim NewDict(255) As Byte
    CopyMemory NewDict(0), Dictionary(0), OldPos
    CopyMemory NewDict(OldPos), Dictionary(OldPos + 1), 255 - OldPos
    CopyMemory Dictionary(0), NewDict(0), DictPos
    Dictionary(DictPos) = Char
    CopyMemory Dictionary(DictPos + 1), NewDict(DictPos), 255 - DictPos
    'Dictionary = Left(Dictionary, DictPos - 1) & Chr(Char) & _
    Mid(Dictionary, DictPos, OldPos - DictPos) & Mid(Dictionary, OldPos + 1)
End Sub

'this sub will add an amount of bits to a sertain stream
Private Sub AddBitsToStream(Toarray As BytePos, Number As Integer, Numbits As Integer)
    Dim x As Long
    If Numbits = 8 And Toarray.BitPos = 0 Then
        If Toarray.Position > UBound(Toarray.Data) Then ReDim Preserve Toarray.Data(Toarray.Position + 500)
        Toarray.Data(Toarray.Position) = Number And &HFF
        Toarray.Position = Toarray.Position + 1
        Exit Sub
    End If
    For x = Numbits - 1 To 0 Step -1
        Toarray.Buffer = Toarray.Buffer * 2 + (-1 * ((Number And 2 ^ x) > 0))
        Toarray.BitPos = Toarray.BitPos + 1
        If Toarray.BitPos = 8 Then
            If Toarray.Position > UBound(Toarray.Data) Then ReDim Preserve Toarray.Data(Toarray.Position + 500)
            Toarray.Data(Toarray.Position) = Toarray.Buffer
            Toarray.BitPos = 0
            Toarray.Buffer = 0
            Toarray.Position = Toarray.Position + 1
        End If
    Next
End Sub

'this function will return a value out of the amaunt of bits you asked for
Private Function ReadBitsFromArray(FromArray() As Byte, FromPos As Long, FromBit As Integer, Numbits As Integer) As Long
    Dim x As Integer
    Dim Temp As Long
    For x = 1 To Numbits
        Temp = Temp * 2 + (-1 * ((FromArray(FromPos) And 2 ^ (7 - FromBit)) > 0))
        FromBit = FromBit + 1
        If FromBit = 8 Then
            If FromPos + 1 > UBound(FromArray) Then
                Do While x < Numbits
                    Temp = Temp * 2
                    x = x + 1
                Loop
                FromPos = FromPos + 1
                Exit For
            End If
            FromPos = FromPos + 1
            FromBit = 0
        End If
    Next
    ReadBitsFromArray = Temp
End Function

'this sub will add a char into the outputstream
Private Sub AddCharToArray(Toarray() As Byte, ToPos As Long, Char As Byte)
    If ToPos > UBound(Toarray) Then ReDim Preserve Toarray(ToPos + 500)
    Toarray(ToPos) = Char
    ToPos = ToPos + 1
End Sub



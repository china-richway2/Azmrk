Attribute VB_Name = "Test"
Public Type UnhookStruct
    szFunName As String * 36
    bNum As Byte
    wBytes(1 To 128) As Byte
End Type

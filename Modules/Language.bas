Attribute VB_Name = "Language"
Public Sub LoadLanguage(Name As String)
    Open Name For Binary Access Read As #1
    Dim Languages() As Byte
    ReDim Languages(LOF(1) - 1)
    Get #1, , Languages
    DeCompress_VBC_Dynamic Languages
    Close #1
    Dim s As String
    s = StrConv(Languages, vbUnicode)
End Sub

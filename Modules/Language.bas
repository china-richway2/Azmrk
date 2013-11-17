Attribute VB_Name = "Language"
Public Type LangItem
    Name As String
    Value As String
End Type
Public Type LangWindow
    SubItemCount As Long
    WindowName As String
    SubItems() As LangItem
End Type
Public Type Lang
    WindowCount As Long
    LangName As String
    LangWindows() As LangWindow
End Type
Dim Languages() As Lang, LangCount As Long
Public DefaultLang As Long

Public Sub LoadLanguage(Name As String)
    Open Name For Binary Access Read As #1
    If LangCount = 0 Then
        ReDim Languages(0)
    Else
        ReDim Preserve Languages(LangCount)
    End If
    Dim Bin() As Byte
    ReDim Bin(LOF(1) - 1)
    Get #1, , Bin
    DeCompress_VBC_Dynamic Bin
    Close #1
    Dim s As String
    s = StrConv(Bin, vbUnicode)
    
    Dim Langs() As String
    Langs = Split(s, vbCrLf & vbCrLf)
    Dim i As Long
    With Languages(LangCount)
        ReDim .LangWindows(UBound(Langs) - 1)
        For i = 0 To UBound(Langs) - 1
            LoadLang Langs(i), .LangWindows(i)
        Next
    End With
    LangCount = LangCount + 1
End Sub

Public Sub ApplyLang(Window As Form, Optional Lang As Long = -1)
    Dim i As Long
    If Lang = -1 Then Lang = DefaultLang
    With Languages(DefaultLang)
        For i = 0 To UBound(.LangWindows)
            If .LangWindows(i).WindowName = Window.Name Then Exit For
        Next
        If i > UBound(.LangWindows) Then Exit Sub '没有对应的语言项
        With .LangWindows(i)
            For i = 0 To .SubItemCount - 1
                Dim pName As String
                pName = .SubItems(i).Name
                If pName = .WindowName Then
                    Window.Caption = .SubItems(i).Value
                ElseIf right(pName, 1) = ")" Then
                    Dim Index As Long
                    Index = InStr(pName, "(") + 1
                    Dim k As String
                    k = left(pName, Index - 2)
                    Index = Val(Mid(pName, Index, Len(pName) - Index))
                    pName = k
                    Window.Controls(pName)(Index).Caption = .SubItems(i).Value
                Else
                    Window.Controls(pName).Caption = .SubItems(i).Value
                End If
            Next
        End With
    End With
End Sub

Public Function FindString(ByVal Name As String, Optional ByVal Lang As Long = -1) As String
    Dim i As Long
    If Lang = -1 Then Lang = DefaultLang
    With Languages(DefaultLang)
        For i = 0 To UBound(.LangWindows)
            If .LangWindows(i).WindowName = "Default" Then Exit For
        Next
        With .LangWindows(i)
            For i = 0 To UBound(.SubItems)
                If .SubItems(i).Name = Name Then
                    FindString = .SubItems(i).Value
                    Exit Function
                End If
            Next
        End With
    End With
End Function

Private Sub LoadLang(ByVal s As String, Target As LangWindow)
    Dim p() As String
    p = Split(s, vbCrLf)
    Dim k, Title As String, ItemCount As Long
    ItemCount = UBound(p)
    Target.SubItemCount = ItemCount
    ReDim Target.SubItems(ItemCount - 1)
    ItemCount = 0
    For Each k In p
        If left(k, 1) = "[" And right(k, 1) = "]" Then
            Target.WindowName = Mid(k, 2, Len(k) - 2)
        ElseIf k <> "" Then
            Dim i As Long
            i = InStr(k, "=")
            With Target.SubItems(ItemCount)
                .Name = left(k, i - 1)
                If .Name = "" Then Stop
                .Value = Mid(k, i + 1)
            End With
            ItemCount = ItemCount + 1
        End If
    Next
End Sub

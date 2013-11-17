VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "打包"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Sub Form_Load()
    Dim p As String
    p = Dir(App.Path & "\", vbDirectory)
    Do Until p = ""
        If p <> "." And p <> ".." And GetAttr(App.Path & "\" & p) And vbDirectory Then s = s & App.Path & "\" & p & "|"
        p = Dir
    Loop
    Dim i
    For Each i In Split(s, "|")
        If i <> "" Then Pack CStr(i)
    Next
    End
End Sub

Private Sub Pack(Path As String)
    Dim p As String
    p = Dir(Path & "\")
    Do Until p = ""
        Dim s As String
        Dim k As String
        Open Path & "\" & p For Binary As #1
        k = Space(LOF(1))
        Get #1, , k
        k = Trim(k)
        k = Replace(k, Chr(0), "")
        Close #1
        s = s & "[" & Left(p, Len(p) - 4) & "]" & vbCrLf & k
        If Right(s, 2) <> vbCrLf Then s = s & vbCrLf
        s = s & vbCrLf
        p = Dir
    Loop
    Dim Buffer1() As Byte, Buffer2() As Byte, Length1 As Long, Length2 As Long
    Buffer1 = StrConv(s, vbFromUnicode)
    Open Path & ".txt" For Binary As #1
    Put #1, , Buffer1
    Close #1
    Length1 = UBound(Buffer1) + 1
    Length2 = Length1
    Buffer2 = Buffer1
    Compress_VBC_Dynamic Buffer2
    Length2 = UBound(Buffer2) + 1
    Open Path & ".lang" For Output As #1
    Close #1
    Open Path & ".lang" For Binary As #1
    Put #1, , Buffer2
    Close #1
End Sub

VERSION 5.00
Begin VB.Form FormTextBox 
   Caption         =   "文本框浏览与修改"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   100
      Left            =   1800
      Top             =   1320
   End
   Begin VB.TextBox Text1 
      Height          =   3135
      Left            =   0
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   0
      Width           =   4695
   End
End
Attribute VB_Name = "FormTextBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public nWnd As Long
Dim Catching As Boolean
Dim LastSetText() As Byte
Public Sub CatchText()
    Dim nTextLength As Long, d() As Byte, Ret As Long
    ReDim d(1 To 1024)
    Catching = True
    Ret = SendMessageTimeout(nWnd, WM_GETTEXT, 1024, VarPtr(d(1)), 0, 5, nTextLength)
    If Ret Then
        If nTextLength > 1024 Then
            ReDim d(1 To nTextLength)
            Ret = SendMessageTimeout(nWnd, WM_GETTEXT, nTextLength, VarPtr(d(1)), 0, 5, nTextLength)
            If Ret Then
                Text1.Text = StrConv(d, vbUnicode)
            End If
        Else
            Text1.Text = StrConv(d, vbUnicode)
        End If
        Caption = "文本框浏览与修改"
    Else
        Caption = "文本框浏览与修改 - 目标窗口未响应或已经关闭"
    End If
    Catching = False
End Sub

Private Sub Form_Load()
    ApplyLang Me
End Sub

Private Sub Form_Resize()
    Text1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub Text1_Change()
    If Catching Then Exit Sub
    Static LastSetText As String
    LastSetText = Text1.Text
    Dim Ret As Long, Ret2 As Long
    Ret = SendMessageTimeout(nWnd, WM_SETTEXT, 1024, StrPtr(StrConv(Text1.Text, vbFromUnicode)), 0, 5, Ret2)
    If Ret Then
        If Ret2 Then
            Caption = FindString("FormTextBox")
        Else
            Caption = FindString("FormTextBox.Error")
        End If
    Else
        Caption = FindString("FormTextBox.Timeout")
    End If
End Sub

Private Sub Timer1_Timer()
    CatchText
End Sub

VERSION 5.00
Begin VB.Form DGEditDWord 
   Caption         =   "修改值"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   495
      Left            =   3000
      TabIndex        =   7
      Top             =   1320
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   480
      TabIndex        =   6
      Top             =   1320
      Width           =   1215
   End
   Begin VB.TextBox Text3 
      Height          =   270
      Left            =   1080
      TabIndex        =   5
      Top             =   840
      Width           =   3495
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1080
      TabIndex        =   3
      Top             =   480
      Width           =   3495
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1080
      TabIndex        =   1
      Top             =   120
      Width           =   3495
   End
   Begin VB.Label Label3 
      Caption         =   "带符号数："
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   840
      Width           =   975
   End
   Begin VB.Label Label2 
      Caption         =   "无符号数："
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "16进制值："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "DGEditDWord"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim bType As Integer, hKey As Long, ValName As String, Class As String, ClassType As Long
Dim dwData As Long
Private Sub SetValue(ByVal DWord As Long)
    If bType <> 1 Then Text1 = right("0000000" & Hex(DWord), 8)
    If bType <> 2 Then Text2 = IIf(DWord > 0, DWord, CDbl(DWord) + 2 ^ 32)
    If bType <> 3 Then Text3 = DWord
    dwData = DWord
End Sub

Public Sub Init(ByVal Key As Long, ByVal ValueName As String, ByVal szClass As String, ByVal dwClassType As Long)
    Caption = FindString("Edit1") & szClass & FindString("Value1")
    hKey = Key
    ValName = ValueName
    Class = szClass
    ClassType = dwClassType
End Sub

Private Sub Command1_Click()
    If RegSetValueEx(hKey, ValName, 0, ClassType, dwData, 4) <> ERROR_SUCCESS Then
        MsgBox FindString("EditDword.Fail"), vbCritical
    Else
        Unload Me
    End If
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Form_Load()
    ApplyLang Me
End Sub

Private Sub Text1_Change()
    If bType Then Exit Sub
    bType = 1
    If IsNumeric("&H" & Text1) Then
        SetValue Val("&H" & Text1)
        Text1.Tag = Text1
    Else
        Text1 = Text1.Tag
    End If
    bType = 0
End Sub

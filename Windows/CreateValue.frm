VERSION 5.00
Begin VB.Form CreateValue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "创建值"
   ClientHeight    =   3060
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4710
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   4710
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "取消"
      Height          =   495
      Left            =   2880
      TabIndex        =   12
      Top             =   2520
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "确定"
      Height          =   495
      Left            =   720
      TabIndex        =   11
      Top             =   2520
      Width           =   1215
   End
   Begin VB.OptionButton Opt 
      Caption         =   "REG_MULTI_SZ"
      Height          =   255
      Index           =   7
      Left            =   960
      TabIndex        =   10
      Top             =   2160
      Width           =   1455
   End
   Begin VB.OptionButton Opt 
      Caption         =   "REG_DWORD_BIG_ENDIAN"
      Height          =   255
      Index           =   6
      Left            =   960
      TabIndex        =   9
      Top             =   1920
      Width           =   2175
   End
   Begin VB.OptionButton Opt 
      Caption         =   "REG_DWORD_LITTLE_ENDIAN"
      Height          =   255
      Index           =   5
      Left            =   960
      TabIndex        =   8
      Top             =   1680
      Width           =   2415
   End
   Begin VB.OptionButton Opt 
      Caption         =   "REG_DWORD"
      Height          =   255
      Index           =   4
      Left            =   960
      TabIndex        =   7
      Top             =   1440
      Width           =   1215
   End
   Begin VB.OptionButton Opt 
      Caption         =   "REG_BINARY"
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.OptionButton Opt 
      Caption         =   "REG_EXPAND_SZ"
      Height          =   255
      Index           =   2
      Left            =   960
      TabIndex        =   5
      Top             =   960
      Width           =   1575
   End
   Begin VB.OptionButton Opt 
      Caption         =   "REG_SZ"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   4
      Top             =   720
      Width           =   855
   End
   Begin VB.OptionButton Opt 
      Caption         =   "REG_NONE"
      Height          =   255
      Index           =   0
      Left            =   960
      TabIndex        =   3
      Top             =   480
      Value           =   -1  'True
      Width           =   1095
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   840
      TabIndex        =   1
      Top             =   120
      Width           =   3735
   End
   Begin VB.Label Label2 
      Caption         =   "值类型："
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   735
   End
   Begin VB.Label Label1 
      Caption         =   "值名称："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   855
   End
End
Attribute VB_Name = "CreateValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim hKey As String, hCheck As Integer
Public Sub Init(ByVal Key As String)
    hKey = Key
End Sub

Private Sub Command1_Click()
    If Text1 = "" Then
        MsgBox "值名称不能为空！", vbInformation
        Text1.SetFocus
    End If
    Registry.CreateValue hKey, Text1, hCheck, True
    Call EnumValue(Menu.tvwKeys.SelectedItem)
    Unload Me
    Menu.SetFocus
End Sub

Private Sub Command2_Click()
    Unload Me
End Sub

Private Sub Opt_Click(Index As Integer)
    hCheck = Index
End Sub

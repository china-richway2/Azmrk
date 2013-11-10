VERSION 5.00
Begin VB.Form ShutdownWindow 
   Caption         =   "关机"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "关机"
      Height          =   495
      Left            =   720
      TabIndex        =   3
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CheckBox Check4 
      Caption         =   "重新启动"
      Height          =   255
      Left            =   720
      TabIndex        =   2
      Tag             =   "2"
      Top             =   840
      Width           =   1695
   End
   Begin VB.CheckBox Check3 
      Caption         =   "强制性的关机"
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Tag             =   "4"
      Top             =   600
      Width           =   1575
   End
   Begin VB.CheckBox Check2 
      Caption         =   "关闭系统并关闭电源"
      Height          =   255
      Left            =   720
      TabIndex        =   0
      Tag             =   "1"
      Top             =   360
      Width           =   2055
   End
End
Attribute VB_Name = "ShutdownWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
    EnablePrivilege SE_SHUTDOWN
    Dim mValue As Long
    mValue = Check2.Value * Check2.Tag + Check3.Value * Check3.Tag + Check4.Value * Check4.Tag
    ExitWindowsEx mValue, 0
End Sub

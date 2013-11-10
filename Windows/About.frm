VERSION 5.00
Begin VB.Form About 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "关于本程序"
   ClientHeight    =   2865
   ClientLeft      =   2340
   ClientTop       =   1935
   ClientWidth     =   5730
   ClipControls    =   0   'False
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1977.474
   ScaleMode       =   0  'User
   ScaleWidth      =   5380.766
   ShowInTaskbar   =   0   'False
   Begin VB.PictureBox picIcon 
      AutoSize        =   -1  'True
      ClipControls    =   0   'False
      Height          =   540
      Left            =   240
      ScaleHeight     =   337.12
      ScaleMode       =   0  'User
      ScaleWidth      =   337.12
      TabIndex        =   1
      Top             =   240
      Width           =   540
   End
   Begin VB.CommandButton cmdOK 
      Cancel          =   -1  'True
      Caption         =   "确定"
      Default         =   -1  'True
      Height          =   705
      Left            =   4125
      TabIndex        =   0
      Top             =   2025
      Width           =   1500
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "GAM2046   http://hi.baidu.com/dazzles"
      ForeColor       =   &H00000000&
      Height          =   180
      Left            =   840
      TabIndex        =   5
      Top             =   1320
      Width           =   3345
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Naylon   http://hi.baidu.com/naylonslain"
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   840
      TabIndex        =   4
      Top             =   960
      Width           =   3615
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00808080&
      BorderStyle     =   6  'Inside Solid
      Index           =   1
      X1              =   84.515
      X2              =   5309.398
      Y1              =   1273.452
      Y2              =   1273.452
   End
   Begin VB.Label lblTitle 
      Caption         =   "Azmrk"
      ForeColor       =   &H00000000&
      Height          =   480
      Left            =   1050
      TabIndex        =   3
      Top             =   240
      Width           =   3885
   End
   Begin VB.Line Line1 
      BorderColor     =   &H00FFFFFF&
      BorderWidth     =   2
      Index           =   0
      X1              =   98.6
      X2              =   5309.398
      Y1              =   1283.805
      Y2              =   1283.805
   End
   Begin VB.Label lblDisclaimer 
      Caption         =   "警告: 本程序乃Naylon无聊时胡乱编写，可随意复制传播，翻版不究，盗版无妨。"
      ForeColor       =   &H00000000&
      Height          =   825
      Left            =   255
      TabIndex        =   2
      Top             =   2025
      Width           =   3630
   End
End
Attribute VB_Name = "About"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdOK_Click()
  Unload Me
End Sub

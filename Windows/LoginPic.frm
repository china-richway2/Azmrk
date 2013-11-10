VERSION 5.00
Begin VB.Form LoginPic 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "Form1"
   ClientHeight    =   4185
   ClientLeft      =   4725
   ClientTop       =   3315
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "LoginPic.frx":0000
   ScaleHeight     =   4185
   ScaleWidth      =   6030
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.Image Image1 
      Height          =   4155
      Left            =   0
      Picture         =   "LoginPic.frx":0152
      Stretch         =   -1  'True
      Top             =   0
      Width           =   6015
   End
End
Attribute VB_Name = "LoginPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub Form_Load()
    App.TaskVisible = False
    'SkinH_Attach  'skin
    'SkinH_SetAero 1 'skin
'SetIcon Me.hwnd, "IDR_MAINFRAME", True 'icon
'    With Image1
'        .top = 0
'        .left = 0
'        .Height = LoginPic.Height
'        .Width = LoginPic.Width
'    End With

End Sub


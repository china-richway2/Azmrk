VERSION 5.00
Begin VB.Form DownloadWindow 
   Caption         =   "ÏÂÔØ"
   ClientHeight    =   945
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4935
   LinkTopic       =   "Form1"
   ScaleHeight     =   945
   ScaleWidth      =   4935
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin Azmrk.Download Download1 
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   120
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   1296
   End
End
Attribute VB_Name = "DownloadWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Status As Boolean

Public Sub Download(url As String, Target As String)
    Download1.BeginDownload url, Target
    Status = False
    Network.Execute url
End Sub

Private Sub Download1_DownloadComplete(MaxBytes As Long, SaveFile As String)
    Status = True
    Unload Me
End Sub


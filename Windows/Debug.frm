VERSION 5.00
Begin VB.Form DebugWindow 
   Caption         =   "调试进程"
   ClientHeight    =   5520
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   8220
   LinkTopic       =   "Form1"
   ScaleHeight     =   5520
   ScaleWidth      =   8220
   StartUpPosition =   3  '窗口缺省
End
Attribute VB_Name = "DebugWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim hProcess As Long, hDebugObj As Long
Public Sub Initialize(ByVal pHandle As Long, ByVal hDebug As Long)
    hProcess = pHandle
    hDebugObj = hDebug
End Sub
Private Sub Form_Load()

End Sub

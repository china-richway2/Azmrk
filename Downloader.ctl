VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.UserControl Download 
   AutoRedraw      =   -1  'True
   ClientHeight    =   225
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   5490
   ScaleHeight     =   225
   ScaleWidth      =   5490
   Begin MSComctlLib.ProgressBar pbar 
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   5535
      _ExtentX        =   9763
      _ExtentY        =   450
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "Download"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Event DownloadProgress(CurBytes As Long, MaxBytes As Long, SaveFile As String)
Event DownloadError(SaveFile As String)
Event DownloadComplete(MaxBytes As Long, SaveFile As String)
Public downStat As Boolean


Public Function CancelAsyncRead() As Boolean
    On Error Resume Next
    UserControl.CancelAsyncRead
End Function

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
    On Error Resume Next
    Dim f() As Byte, fn As Long
    If AsyncProp.BytesMax <> 0 Then
        fn = FreeFile
        f = AsyncProp.Value
        Open AsyncProp.PropertyName For Binary Access Write As #fn
        Put #fn, , f
        Close #fn
    Else
        RaiseEvent DownloadError(AsyncProp.PropertyName)
        Exit Sub
    End If
    RaiseEvent DownloadComplete(CLng(AsyncProp.BytesMax), AsyncProp.PropertyName)
    downStat = False
End Sub
Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
    On Error Resume Next
    If AsyncProp.BytesMax <> 0 Then
        RaiseEvent DownloadProgress(CLng(AsyncProp.BytesRead), CLng(AsyncProp.BytesMax), AsyncProp.PropertyName)
        downStat = True
    End If
End Sub

Public Sub BeginDownload(url As String, SaveFile As String)
    On Error GoTo ErrorBeginDownload
    downStat = True
    UserControl.AsyncRead url, vbAsyncTypeByteArray, SaveFile, vbAsyncReadForceUpdate
    Exit Sub
ErrorBeginDownload:
    downStat = False
    RaiseEvent DownloadError(SaveFile)
End Sub



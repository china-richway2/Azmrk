VERSION 5.00
Object = "{248DD890-BB45-11CF-9ABC-0080C7E7B78D}#1.0#0"; "MSWINSCK.OCX"
Begin VB.Form MainWindow 
   Caption         =   "服务端"
   ClientHeight    =   3090
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command1 
      Caption         =   "启动"
      Height          =   495
      Left            =   1800
      TabIndex        =   4
      Top             =   1560
      Width           =   1215
   End
   Begin VB.TextBox Text2 
      Height          =   270
      Left            =   1680
      TabIndex        =   3
      Top             =   1080
      Width           =   1935
   End
   Begin VB.TextBox Text1 
      Height          =   270
      Left            =   1680
      TabIndex        =   1
      Text            =   "15536"
      Top             =   600
      Width           =   1935
   End
   Begin MSWinsockLib.Winsock wsk 
      Left            =   3720
      Top             =   360
      _ExtentX        =   741
      _ExtentY        =   741
      _Version        =   393216
      Protocol        =   1
   End
   Begin VB.Label Label2 
      Caption         =   "密码："
      Height          =   255
      Left            =   600
      TabIndex        =   2
      Top             =   1080
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "本地端口："
      Height          =   255
      Left            =   600
      TabIndex        =   0
      Top             =   600
      Width           =   1815
   End
End
Attribute VB_Name = "MainWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public Function ReadByte() As Byte
    Dim wByte As Byte
    wsk.GetData wByte
    ReadByte = wByte
End Function

Public Sub SendByte(ByVal wByte As Byte)
    wsk.SendData wByte
End Sub

Public Function ReadLong() As Long
    Dim wLong As Long
    wsk.GetData wLong
    ReadLong = wLong
End Function

Public Sub SendLong(ByVal wLong As Long)
    wsk.SendData wLong
End Sub

Public Function ReadInt() As Integer
    Dim wInt As Integer
    wsk.GetData wInt
    ReadInt = wInt
End Function

Public Sub SendInt(ByVal wInt As result)
    wsk.SendData CInt(wInt)
End Sub

Public Function ReadStr() As String
    Dim sStr() As Byte, i As Long
    i = -1
    ReDim sStr(0)
    Do
        i = i + 1
        sStr(i) = ReadByte
        ReDim Preserve sStr(i)
    Loop Until sStr(i) = 0
    ReadStr = StrConv(sStr, vbUnicode)
End Function

Public Sub SendStr(ByVal sStr As String)
    Dim S() As Byte
    S = StrConv(sStr, vbFromUnicode)
    wsk.SendData S
    wsk.SendData CByte(0)
End Sub

Private Sub Command1_Click()
    wsk.Bind Text1.Text
End Sub

Private Sub wsk_DataArrival(ByVal bytesTotal As Long)
    Dim S As String
    S = ReadStr
    Dim nCmd As Cmd
    nCmd = ReadInt
    If nCmd = CHECK_SERVER Then '只有此命令可以绕过密码检测函数
        SendInt SUCCESS
        Send
        Exit Sub
    End If
    If S <> Text2.Text Then
        SendInt SERVER_PASSWORD_INCORRECT
        Send
    End If
    Select Case nCmd
    Case CMD_LOGIN
    Case PROCESS_REFRESH
        Call PNNew
    Case CMD_PROCESS_TERMINATE
    Case PROCESS_SUSPEND
    Case PROCESS_RESUME
    Case PROCESS_SET_PRIORITY
    
    Case THREAD_REFRESH
    Case CMD_THREAD_TERMINATE
    Case THREAD_SUSPEND
    Case THREAD_RESUME
    Case THREAD_SET_PRIORITY
    
    Case WINDOW_REFRESH
    Case WINDOW_FROMPOINT
    Case WINDOW_UPDATE
    Case WINDOW_CLOSE
    End Select
End Sub


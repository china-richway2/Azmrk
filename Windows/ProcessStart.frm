VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form ProcessStart 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "    "
   ClientHeight    =   3990
   ClientLeft      =   150
   ClientTop       =   540
   ClientWidth     =   7230
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3990
   ScaleWidth      =   7230
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame fSInfo 
      Caption         =   "启动参数"
      Height          =   2175
      Left            =   120
      TabIndex        =   24
      Top             =   1200
      Width           =   2655
      Begin VB.Frame Frame7 
         Caption         =   "显示模式"
         Height          =   1815
         Left            =   120
         TabIndex        =   25
         Top             =   240
         Width           =   2415
         Begin VB.OptionButton optShow 
            Caption         =   "不激活"
            Height          =   255
            Index           =   8
            Left            =   1320
            TabIndex        =   34
            Top             =   480
            Width           =   855
         End
         Begin VB.OptionButton optShow 
            Caption         =   "不激活+最小化"
            Height          =   255
            Index           =   7
            Left            =   720
            TabIndex        =   33
            Top             =   240
            Value           =   -1  'True
            Width           =   1575
         End
         Begin VB.OptionButton optShow 
            Caption         =   "最小化"
            Height          =   255
            Index           =   6
            Left            =   1320
            TabIndex        =   32
            Top             =   1440
            Width           =   975
         End
         Begin VB.OptionButton optShow 
            Caption         =   "原尺寸显示"
            Height          =   255
            Index           =   5
            Left            =   120
            TabIndex        =   31
            Top             =   1440
            Width           =   1215
         End
         Begin VB.OptionButton optShow 
            Caption         =   "不激活+显示"
            Height          =   255
            Index           =   4
            Left            =   120
            TabIndex        =   30
            Top             =   1200
            Width           =   1335
         End
         Begin VB.OptionButton optShow 
            Caption         =   "不激活+最大化"
            Height          =   255
            Index           =   3
            Left            =   120
            TabIndex        =   29
            Top             =   960
            Width           =   1575
         End
         Begin VB.OptionButton optShow 
            Caption         =   "激活+最小化"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   28
            Top             =   720
            Width           =   1575
         End
         Begin VB.OptionButton optShow 
            Caption         =   "激活+显示"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   27
            Top             =   480
            Width           =   1215
         End
         Begin VB.OptionButton optShow 
            Caption         =   "隐藏"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   26
            Top             =   240
            Width           =   735
         End
      End
   End
   Begin VB.Frame Frame6 
      Caption         =   "结果"
      Height          =   1095
      Left            =   3600
      TabIndex        =   15
      Top             =   2760
      Width           =   3615
      Begin VB.TextBox LastErr 
         Height          =   270
         Left            =   2400
         TabIndex        =   22
         Top             =   720
         Width           =   1095
      End
      Begin VB.TextBox TID 
         Height          =   270
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   19
         Top             =   480
         Width           =   975
      End
      Begin VB.TextBox PID 
         Height          =   270
         Left            =   2520
         Locked          =   -1  'True
         TabIndex        =   17
         Top             =   240
         Width           =   975
      End
      Begin VB.Label Label5 
         Caption         =   "LastErr"
         Height          =   255
         Left            =   1680
         TabIndex        =   21
         Top             =   720
         Width           =   735
      End
      Begin VB.Label m_TID 
         Caption         =   "TID"
         Height          =   255
         Left            =   2160
         TabIndex        =   18
         Top             =   480
         Width           =   375
      End
      Begin VB.Label m_PID 
         Caption         =   "PID"
         Height          =   255
         Left            =   2160
         TabIndex        =   16
         Top             =   240
         Width           =   375
      End
   End
   Begin VB.Frame Frame3 
      Caption         =   "参数"
      Height          =   1095
      Left            =   3600
      TabIndex        =   8
      Top             =   1560
      Width           =   3615
      Begin VB.PictureBox Picture1 
         Height          =   255
         Left            =   1680
         ScaleHeight     =   195
         ScaleWidth      =   195
         TabIndex        =   20
         Top             =   720
         Width           =   255
      End
      Begin VB.TextBox dwCreationFlags 
         Height          =   270
         Left            =   2040
         TabIndex        =   14
         Tag             =   "0"
         Top             =   720
         Width           =   1455
      End
      Begin VB.TextBox bInheritHandles 
         Height          =   270
         Left            =   2040
         TabIndex        =   12
         Tag             =   "0"
         Top             =   480
         Width           =   1455
      End
      Begin VB.TextBox lpApplicationName 
         Height          =   270
         Left            =   2040
         TabIndex        =   10
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label m_dwCreationFlags 
         Caption         =   "dwCreationFlags"
         Height          =   255
         Left            =   120
         TabIndex        =   13
         Top             =   720
         Width           =   1575
      End
      Begin VB.Label m_bInheritHandles 
         Caption         =   "bInheritHandles"
         Height          =   255
         Left            =   120
         TabIndex        =   11
         Top             =   480
         Width           =   1695
      End
      Begin VB.Label m_lpApplicationName 
         Caption         =   "lpApplicationName"
         Height          =   255
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1575
      End
   End
   Begin VB.CommandButton cmdSelect 
      Caption         =   "选择文件"
      Height          =   375
      Left            =   6000
      TabIndex        =   7
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "重置"
      Height          =   375
      Left            =   3600
      TabIndex        =   6
      Top             =   1200
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "启动"
      Height          =   375
      Left            =   4800
      TabIndex        =   5
      Top             =   1200
      Width           =   1215
   End
   Begin VB.Frame Frame1 
      Caption         =   "基本参数"
      Height          =   975
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7095
      Begin VB.CheckBox Check1 
         Caption         =   "启动并调试"
         Height          =   255
         Left            =   5760
         TabIndex        =   23
         Top             =   600
         Width           =   1215
      End
      Begin VB.TextBox lpCurrentDirectory 
         Height          =   270
         Left            =   1080
         TabIndex        =   4
         Top             =   600
         Width           =   4695
      End
      Begin VB.TextBox lpCommandLine 
         Height          =   270
         Left            =   1080
         TabIndex        =   2
         Top             =   240
         Width           =   5895
      End
      Begin VB.Label Label2 
         Caption         =   "工作目录："
         Height          =   255
         Left            =   120
         TabIndex        =   3
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label1 
         Caption         =   "命令行："
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   855
      End
   End
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   3240
      Top             =   1080
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Menu m1 
      Caption         =   "dwCreationFlags内容"
      Visible         =   0   'False
      Begin VB.Menu mnuFlags 
         Caption         =   "CREATE_NEW"
         Index           =   0
      End
      Begin VB.Menu mnuFlags 
         Caption         =   "CREATE_ALWAYS"
         Index           =   1
      End
      Begin VB.Menu mnuFlags 
         Caption         =   "CREATE_SUSPENDED"
         Index           =   2
      End
      Begin VB.Menu mnuFlags 
         Caption         =   "CREATE_NEW_CONSOLE"
         Index           =   4
      End
      Begin VB.Menu mnuFlags 
         Caption         =   "CREATE_NEW_PROCESS_GROUP"
         Index           =   9
      End
      Begin VB.Menu mnuFlag 
         Caption         =   "CREATE_NO_WINDOW"
         Index           =   27
      End
   End
End
Attribute VB_Name = "ProcessStart"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim pAttr As SECURITY_ATTRIBUTES, tAttr As SECURITY_ATTRIBUTES, sInfo As STARTUPINFO, pInfo As PROCESS_INFORMATION

Private Sub cmdSelect_Click()
    On Error GoTo Cancel
    With Dialog1
        .CancelError = True
        .InitDir = App.Path
        .Filter = "可执行文件(*.exe)|*.exe"
        .ShowOpen
        lpCommandLine = """" & .FileName & """"
        lpCurrentDirectory = left(.FileName, InStrRev(.FileName, "\"))
        lpApplicationName = Mid(.FileName, InStrRev(.FileName, "\") + 1)
    End With
Cancel:
End Sub

Private Sub Check(ByRef dwData As Long, ByRef szData As TextBox)
    If IsNumeric(szData.Text) Then
        dwData = Val(szData.Text)
    Else
        MsgBox szData.Name & "一栏必须是数值！", vbInformation
        Err.Raise 1
    End If
End Sub
Private Sub Command1_Click()
    On Error GoTo E
    With sInfo
        Dim i As Long
        For i = 0 To optShow.UBound
            If optShow(i).Value Then Exit For
        Next
        sInfo.wShowWindow = i
    End With
    Dim A As Long, b As Long, Ret As Long
    Check A, bInheritHandles
    Check b, dwCreationFlags
    Ret = CreateProcess(lpApplicationName.Text, lpCommandLine.Text, _
    pAttr, tAttr, A, b, ByVal 0, lpCurrentDirectory.Text, sInfo, pInfo)
    Dim s() As String
    If Ret Then
        ReDim s(2)
        s(0) = "创建进程成功！"
        s(1) = "进程PID:" & pInfo.dwProcessId
        s(2) = "主线程TID:" & pInfo.dwThreadId
        MsgBox Join(s, vbCrLf), vbInformation
        ZwClose pInfo.hProcess
        ZwClose pInfo.hThread
    Else
        ReDim s(1)
        s(0) = "创建进程失败！"
        s(1) = "GetLastError返回值：" & Err.LastDllError
        MsgBox Join(s, vbCrLf), vbCritical
    End If
    Call PNNew
    Exit Sub
E:
    If Err.Number <> 1 Then
        MsgBox "错误" & Err.Number & "：" & Err.Description, vbCritical
    End If
End Sub

Private Sub Command2_Click()
    Dim obj As Object
    For Each obj In Me.Controls
        If TypeOf obj Is TextBox Then
            If obj.Name <> "sFileName" And obj.Name <> "sPath" Then
                obj.Text = obj.Tag
            End If
        End If
    Next
End Sub

Private Sub m_hThread_Click()
End Sub

Private Sub Picture1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    On Error Resume Next
    If IsNumeric(dwCreationFlags.Text) Then
        Dim i As Long
        For i = 0 To 31
            If mnuFlags(i) Is Nothing Then
            Else
                If Val(dwCreationFlags.Text) And (2 ^ i) Then
                    mnuFlags(i).Checked = True
                Else
                    mnuFlags(i).Checked = False
                End If
            End If
        Next
        PopupMenu m1
    End If
End Sub

Private Sub Form_Load()
    lpCurrentDirectory.Text = App.Path
    Command2_Click
End Sub

Private Sub mnuFlags_Click(Index As Integer)
    mnuFlags(Index).Checked = Not mnuFlags(Index).Checked
    If left(dwCreationFlags.Text, 2) = "&H" Then
        dwCreationFlags.Text = "&H" & Hex(Val(dwCreationFlags.Text) Xor (2 ^ Index))
    Else
        dwCreationFlags.Text = Val(dwCreationFlags.Text) Xor (2 ^ Index)
    End If
End Sub


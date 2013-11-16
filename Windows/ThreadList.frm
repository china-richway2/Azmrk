VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form ThreadList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "进程线程"
   ClientHeight    =   5460
   ClientLeft      =   3645
   ClientTop       =   3555
   ClientWidth     =   10395
   Icon            =   "ThreadList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ThreadList.frx":0CCA
   ScaleHeight     =   5460
   ScaleWidth      =   10395
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ListView ListView1 
      Height          =   5475
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10395
      _ExtentX        =   18336
      _ExtentY        =   9657
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Menu tMenu 
      Caption         =   "线程查看"
      Visible         =   0   'False
      Begin VB.Menu tNew 
         Caption         =   "刷新列表"
      End
      Begin VB.Menu t01 
         Caption         =   "-"
      End
      Begin VB.Menu tSuspend 
         Caption         =   "挂起线程"
      End
      Begin VB.Menu tResume 
         Caption         =   "恢复线程"
      End
      Begin VB.Menu tMenuTerminate 
         Caption         =   "结束线程"
         Begin VB.Menu tTerminate 
            Caption         =   "ZwTerminateThread"
         End
         Begin VB.Menu tTerminateByThreadMessage 
            Caption         =   "PostThreadMessage"
         End
         Begin VB.Menu tTerminateByDestroyThreadContext 
            Caption         =   "Developing..."
         End
      End
      Begin VB.Menu t02 
         Caption         =   "-"
      End
      Begin VB.Menu tShowSubWindows 
         Caption         =   "显示窗口"
      End
   End
End
Attribute VB_Name = "ThreadList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mPid As Long

Private Sub Form_Load()
    ApplyLang Me
    With ListView1.ColumnHeaders
        .Add , , "线程ID", 920
        .Add , , "TEB", 1300
        .Add , , "ETHREAD", 1300
        .Add , , "优先级", 920
        .Add , , "线程入口", 1300
        .Add , , "线程状态", 1000
        .Add , , "线程模块", 3540
    End With
    
    If nsItem = 0 Then Exit Sub

    mPid = nsItem
    
    ListAllThreads mPid, Me
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LVAutoOrder ListView1, ColumnHeader
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu tMenu
    End If
End Sub

Private Sub tNew_Click()
    If ListView1.Sorted = True Then ListView1.Sorted = False
    ListAllThreads mPid
End Sub

Private Sub tResume_Click()
    Dim hThread As Long
    
    hThread = FxNormalOpenThread(THREAD_SUSPEND_RESUME, CLng(ListView1.SelectedItem.Text))

    If hThread = 0 Then
        MsgBox "拒绝访问!", 0, "失败"
        Exit Sub
    End If
    
    ResumeThread hThread
    ZwClose hThread
    
    Call tNew_Click
End Sub

Private Sub tShowSubWindows_Click()
    If ListView1.SelectedItem Is Nothing Then Exit Sub
    mWindowFilterMethod = (mWindowFilterMethod And (Not 12)) Or MethodListByTID
    mWindowFilterArg = ListView1.SelectedItem.Text
    SetTab = True
    Menu.lLabels_Click 0
    SetTab = False
    Call CNNew
    Menu.SetFocus
End Sub

Private Sub tSuspend_Click()
    Dim hThread As Long
    
    hThread = FxNormalOpenThread(THREAD_SUSPEND_RESUME, CLng(ListView1.SelectedItem.Text))

    If hThread = 0 Then
        MsgBox "拒绝访问!", 0, "失败"
        Exit Sub
    End If
    
    SuspendThread hThread
    ZwClose hThread
    
    Call tNew_Click
End Sub

Private Sub tTerminate_Click()
    Dim hThread As Long

    hThread = FxNormalOpenThread(THREAD_TERMINATE, CLng(ListView1.SelectedItem.Text))
    
    If hThread = 0 Then
        MsgBox "拒绝访问!", 0, "失败"
        Exit Sub
    End If
    
    ZwTerminateThread hThread, 0&
    WaitForSingleObject hThread, INFINITE
    
    ZwClose hThread
    
    Call tNew_Click
End Sub

Private Sub tTerminateByDestroyThreadContext_Click()
    Dim hThread As Long
    
    hThread = FxNormalOpenThread(THREAD_ALL_ACCESS, CLng(ListView1.SelectedItem.Text))
    FxDestroyThreadContext hThread
    
    Call tNew_Click
End Sub

Private Sub tTerminateByThreadMessage_Click()
    PostThreadMessage CLng(Replace(ListView1.SelectedItem.SubItems(1), "0x", "&H")), WM_STOP, 0, 0
    
    Call tNew_Click
End Sub

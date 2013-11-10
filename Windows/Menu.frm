VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Object = "{BDC217C8-ED16-11CD-956C-0000C04E4C0A}#1.1#0"; "tabctl32.ocx"
Begin VB.Form Menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Azmrk"
   ClientHeight    =   7350
   ClientLeft      =   2010
   ClientTop       =   2490
   ClientWidth     =   12585
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Azmrk|Menu"
   MaxButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   12585
   StartUpPosition =   2  '屏幕中心
   Begin TabDlg.SSTab SSTab1 
      Height          =   7155
      Left            =   120
      TabIndex        =   0
      Top             =   60
      Width           =   12345
      _ExtentX        =   21775
      _ExtentY        =   12621
      _Version        =   393216
      Tabs            =   6
      TabsPerRow      =   6
      TabHeight       =   520
      WordWrap        =   0   'False
      ShowFocusRect   =   0   'False
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      TabCaption(0)   =   "窗口"
      TabPicture(0)   =   "Menu.frx":20082
      Tab(0).ControlEnabled=   -1  'True
      Tab(0).Control(0)=   "Label1"
      Tab(0).Control(0).Enabled=   0   'False
      Tab(0).Control(1)=   "ListView1"
      Tab(0).Control(1).Enabled=   0   'False
      Tab(0).Control(2)=   "Text1"
      Tab(0).Control(2).Enabled=   0   'False
      Tab(0).Control(3)=   "Check3"
      Tab(0).Control(3).Enabled=   0   'False
      Tab(0).Control(4)=   "Check2"
      Tab(0).Control(4).Enabled=   0   'False
      Tab(0).Control(5)=   "Check1"
      Tab(0).Control(5).Enabled=   0   'False
      Tab(0).ControlCount=   6
      TabCaption(1)   =   "进程"
      TabPicture(1)   =   "Menu.frx":2009E
      Tab(1).ControlEnabled=   0   'False
      Tab(1).Control(0)=   "Label3"
      Tab(1).Control(1)=   "ListView2"
      Tab(1).Control(2)=   "pcNewTask"
      Tab(1).Control(3)=   "pcSearchModules"
      Tab(1).Control(4)=   "ImageList1"
      Tab(1).ControlCount=   5
      TabCaption(2)   =   "服务"
      TabPicture(2)   =   "Menu.frx":200BA
      Tab(2).ControlEnabled=   0   'False
      Tab(2).Control(0)=   "Label5"
      Tab(2).Control(1)=   "LVServer"
      Tab(2).ControlCount=   2
      TabCaption(3)   =   "网络"
      TabPicture(3)   =   "Menu.frx":200D6
      Tab(3).ControlEnabled=   0   'False
      Tab(3).ControlCount=   0
      TabCaption(4)   =   "文件"
      TabPicture(4)   =   "Menu.frx":200F2
      Tab(4).ControlEnabled=   0   'False
      Tab(4).ControlCount=   0
      TabCaption(5)   =   "注册表"
      TabPicture(5)   =   "Menu.frx":2010E
      Tab(5).ControlEnabled=   0   'False
      Tab(5).Control(0)=   "imlIcons"
      Tab(5).Control(1)=   "tvwKeys"
      Tab(5).Control(2)=   "lvwData"
      Tab(5).ControlCount=   3
      Begin VB.CheckBox Check1 
         Caption         =   "总在最上"
         Height          =   255
         Left            =   8520
         TabIndex        =   14
         Top             =   6960
         Width           =   1095
      End
      Begin VB.CheckBox Check2 
         Caption         =   "显示全部窗口"
         Height          =   255
         Left            =   10680
         TabIndex        =   13
         Top             =   6960
         Width           =   1575
      End
      Begin VB.CheckBox Check3 
         Caption         =   "跟随鼠标"
         Height          =   255
         Left            =   9600
         TabIndex        =   12
         Top             =   6960
         Width           =   1035
      End
      Begin MSComctlLib.ImageList ImageList1 
         Left            =   -68040
         Top             =   6300
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         MaskColor       =   12632256
         _Version        =   393216
      End
      Begin VB.CommandButton pcSearchModules 
         Caption         =   "搜索模块"
         Height          =   435
         Left            =   -71520
         TabIndex        =   9
         Top             =   6900
         Visible         =   0   'False
         Width           =   1395
      End
      Begin VB.CommandButton pcNewTask 
         Caption         =   "新建任务"
         Height          =   435
         Left            =   -73200
         TabIndex        =   6
         Top             =   6900
         Width           =   1395
      End
      Begin VB.TextBox Text1 
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   2160
         TabIndex        =   3
         Text            =   "输入标题或类名或句柄查找"
         Top             =   6900
         Width           =   6135
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6045
         Left            =   120
         TabIndex        =   1
         Top             =   750
         Width           =   12120
         _ExtentX        =   21378
         _ExtentY        =   10663
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
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
      Begin MSComctlLib.ListView ListView2 
         Height          =   6135
         Left            =   -74880
         TabIndex        =   4
         Top             =   720
         Width           =   12105
         _ExtentX        =   21352
         _ExtentY        =   10821
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   0
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
      Begin MSComctlLib.ListView LVServer 
         Height          =   6135
         Left            =   -74880
         TabIndex        =   7
         Top             =   720
         Width           =   12105
         _ExtentX        =   21352
         _ExtentY        =   10821
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
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
      Begin MSComctlLib.ListView lvwData 
         Height          =   6045
         Left            =   -72300
         TabIndex        =   10
         Top             =   720
         Width           =   9495
         _ExtentX        =   16748
         _ExtentY        =   10663
         View            =   3
         LabelEdit       =   1
         Sorted          =   -1  'True
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "imlIcons"
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "名称"
            Object.Width           =   3969
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "值"
            Object.Width           =   6615
         EndProperty
      End
      Begin MSComctlLib.TreeView tvwKeys 
         Height          =   6105
         Left            =   -74880
         TabIndex        =   11
         Top             =   720
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   10769
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         HotTracking     =   -1  'True
         ImageList       =   "imlIcons"
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
      End
      Begin MSComctlLib.ImageList imlIcons 
         Left            =   -70140
         Top             =   1320
         _ExtentX        =   1005
         _ExtentY        =   1005
         BackColor       =   -2147483643
         ImageWidth      =   16
         ImageHeight     =   16
         MaskColor       =   12632256
         _Version        =   393216
         BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
            NumListImages   =   5
            BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Menu.frx":2012A
               Key             =   ""
            EndProperty
            BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Menu.frx":206C6
               Key             =   ""
            EndProperty
            BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Menu.frx":20C62
               Key             =   ""
            EndProperty
            BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Menu.frx":20DBE
               Key             =   ""
            EndProperty
            BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
               Picture         =   "Menu.frx":20F1A
               Key             =   ""
            EndProperty
         EndProperty
      End
      Begin VB.Label Label5 
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   8
         Top             =   7020
         Width           =   3135
      End
      Begin VB.Label Label3 
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   -74880
         TabIndex        =   5
         Top             =   7020
         Width           =   3135
      End
      Begin VB.Label Label1 
         ForeColor       =   &H00000000&
         Height          =   255
         Left            =   120
         TabIndex        =   2
         Top             =   6900
         Width           =   3135
      End
   End
   Begin VB.Menu mainMenu 
      Caption         =   "功能"
      Begin VB.Menu mainSetting 
         Caption         =   "设置"
      End
      Begin VB.Menu mainReadme 
         Caption         =   "说明"
      End
   End
   Begin VB.Menu FatherMenus 
      Caption         =   "父窗菜单"
      Visible         =   0   'False
      Begin VB.Menu nNewMenu 
         Caption         =   "刷新列表"
         Begin VB.Menu nNew 
            Caption         =   "EnumWindows"
         End
         Begin VB.Menu nFxNew 
            Caption         =   "Parent[Naylon]"
         End
         Begin VB.Menu nFdNewByMessage 
            Caption         =   "PostMessage[gam2046]"
         End
      End
      Begin VB.Menu nChildNewMenu 
         Caption         =   "刷新列表"
      End
      Begin VB.Menu nViewChildWindows 
         Caption         =   "查看子窗口"
      End
      Begin VB.Menu nViewFatherWindows 
         Caption         =   "查看父窗口"
      End
      Begin VB.Menu n01 
         Caption         =   "-"
      End
      Begin VB.Menu nJumpToParent 
         Caption         =   "转到父窗口"
      End
      Begin VB.Menu nJumpToTasklist 
         Caption         =   "转到对应进程"
      End
      Begin VB.Menu n02 
         Caption         =   "-"
      End
      Begin VB.Menu nWindowMax 
         Caption         =   "窗口最大化"
      End
      Begin VB.Menu nWindowMin 
         Caption         =   "窗口最小化"
      End
      Begin VB.Menu n03 
         Caption         =   "-"
      End
      Begin VB.Menu nHide 
         Caption         =   "隐藏窗口"
      End
      Begin VB.Menu nShow 
         Caption         =   "显示窗口"
      End
      Begin VB.Menu n04 
         Caption         =   "-"
      End
      Begin VB.Menu nEnableF 
         Caption         =   "冻结窗口"
      End
      Begin VB.Menu nEnableT 
         Caption         =   "激活窗口"
      End
      Begin VB.Menu n05 
         Caption         =   "-"
      End
      Begin VB.Menu nAmend 
         Caption         =   "修改窗口标题"
      End
      Begin VB.Menu n06 
         Caption         =   "-"
      End
      Begin VB.Menu nMove 
         Caption         =   "移动窗口"
      End
      Begin VB.Menu n09 
         Caption         =   "-"
      End
      Begin VB.Menu nTop 
         Caption         =   "置顶窗口"
      End
      Begin VB.Menu nNoTop 
         Caption         =   "取消置顶"
      End
      Begin VB.Menu n07 
         Caption         =   "-"
      End
      Begin VB.Menu nCopyItems 
         Caption         =   "复制选定项信息"
         Begin VB.Menu nCopyName 
            Caption         =   "窗口名称"
         End
         Begin VB.Menu nCopyClass 
            Caption         =   "窗口类名"
         End
         Begin VB.Menu nCopyHandle 
            Caption         =   "窗口句柄"
         End
      End
      Begin VB.Menu n08 
         Caption         =   "-"
      End
      Begin VB.Menu nCloseMenu 
         Caption         =   "关闭窗口"
         Begin VB.Menu nClose 
            Caption         =   "WM_CLOSE"
         End
         Begin VB.Menu nCloseWindowByMessage 
            Caption         =   "BombWindow"
         End
         Begin VB.Menu nCloseWindowByParent 
            Caption         =   "ReplaceParentWindow"
         End
         Begin VB.Menu nCloseWindowByEndTask 
            Caption         =   "EndTask"
         End
         Begin VB.Menu nCloseWindowByWndProc 
            Caption         =   "Developing..."
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu pMenu 
      Caption         =   "进程菜单"
      Visible         =   0   'False
      Begin VB.Menu pNew 
         Caption         =   "刷新列表"
         Begin VB.Menu pNewSh 
            Caption         =   "Toolhelp32"
         End
         Begin VB.Menu pNewBySession 
            Caption         =   "SessionProcessLinks"
         End
         Begin VB.Menu pNewByHandle 
            Caption         =   "Developing..."
            Enabled         =   0   'False
         End
      End
      Begin VB.Menu p01 
         Caption         =   "-"
      End
      Begin VB.Menu pListThread 
         Caption         =   "查看进程线程"
      End
      Begin VB.Menu pListModule 
         Caption         =   "查看进程模块"
      End
      Begin VB.Menu pListWindows 
         Caption         =   "查看进程窗口"
      End
      Begin VB.Menu p02 
         Caption         =   "-"
      End
      Begin VB.Menu pJumpToParent 
         Caption         =   "转到父进程"
      End
      Begin VB.Menu p03 
         Caption         =   "-"
      End
      Begin VB.Menu pSetPriority 
         Caption         =   "设置优先级"
         Begin VB.Menu pPriorityHigh 
            Caption         =   "较高"
         End
         Begin VB.Menu pPriorityNormal 
            Caption         =   "标准"
         End
         Begin VB.Menu pPriorityLow 
            Caption         =   "较低"
         End
      End
      Begin VB.Menu p04 
         Caption         =   "-"
      End
      Begin VB.Menu pSuspendProcess 
         Caption         =   "挂起进程"
      End
      Begin VB.Menu pResumeProcess 
         Caption         =   "恢复进程"
      End
      Begin VB.Menu p05 
         Caption         =   "-"
      End
      Begin VB.Menu pMoreInformation 
         Caption         =   "详细信息"
      End
      Begin VB.Menu p06 
         Caption         =   "-"
      End
      Begin VB.Menu pCopyInfo 
         Caption         =   "复制选定项信息"
         Begin VB.Menu pCopyPid 
            Caption         =   "PID"
         End
         Begin VB.Menu pCopyPEB 
            Caption         =   "PEB"
         End
         Begin VB.Menu pCopyEPROCESS 
            Caption         =   "EPROCESS"
         End
         Begin VB.Menu pCopyName 
            Caption         =   "进程名称"
         End
         Begin VB.Menu pCopyPath 
            Caption         =   "进程路径"
         End
         Begin VB.Menu pCopyCommandLine 
            Caption         =   "命令行"
         End
      End
      Begin VB.Menu p07 
         Caption         =   "-"
      End
      Begin VB.Menu pMenuTerminateProcess 
         Caption         =   "结束进程"
         Begin VB.Menu pTerminateProcessNormal 
            Caption         =   "ZwTerminateProcess"
         End
         Begin VB.Menu pTerminateProcessByRemoteThread 
            Caption         =   "CreateRemoteThread->ExitProcess"
         End
         Begin VB.Menu pTerminateProcessByDebugProcess 
            Caption         =   "ZwDebugActiveProcess"
         End
         Begin VB.Menu pWinStationTerminateProcess 
            Caption         =   "WinStationTerminateProcess"
         End
      End
   End
   Begin VB.Menu sMenu 
      Caption         =   "服务菜单"
      Visible         =   0   'False
      Begin VB.Menu sNew 
         Caption         =   "刷新列表"
      End
      Begin VB.Menu s01 
         Caption         =   "-"
      End
      Begin VB.Menu MenuStartServer 
         Caption         =   "启动服务"
      End
      Begin VB.Menu MenuPauseServer 
         Caption         =   "暂停服务"
         Visible         =   0   'False
      End
      Begin VB.Menu MenuStopServer 
         Caption         =   "停止服务"
      End
      Begin VB.Menu s02 
         Caption         =   "-"
      End
      Begin VB.Menu MenuS 
         Caption         =   "设置启动类型"
         Begin VB.Menu MenuSetAuto 
            Caption         =   "自动启动"
         End
         Begin VB.Menu MenuSetUser 
            Caption         =   "手动启动"
         End
         Begin VB.Menu MenuSetCant 
            Caption         =   "禁止启动"
         End
      End
      Begin VB.Menu s03 
         Caption         =   "-"
      End
      Begin VB.Menu sCopyInfo 
         Caption         =   "复制选定项信息"
         Begin VB.Menu sCopyServiceName 
            Caption         =   "服务名称"
         End
         Begin VB.Menu sCopyServiceExePath 
            Caption         =   "映像路径"
         End
         Begin VB.Menu sCopyServiceDllPath 
            Caption         =   "DLL 路径"
         End
         Begin VB.Menu sCopyServiceDescribe 
            Caption         =   "服务描述"
         End
      End
      Begin VB.Menu s04 
         Caption         =   "-"
      End
      Begin VB.Menu sMoreInformation 
         Caption         =   "详细信息"
      End
      Begin VB.Menu s05 
         Caption         =   "-"
      End
      Begin VB.Menu sSelectExe 
         Caption         =   "定位文件"
      End
      Begin VB.Menu sExeNature 
         Caption         =   "文件属性"
      End
      Begin VB.Menu s06 
         Caption         =   "-"
      End
      Begin VB.Menu sSelectDll 
         Caption         =   "定位 DLL"
      End
      Begin VB.Menu sDllNature 
         Caption         =   "DLL 属性"
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FirstFocus As Boolean

Private Sub Check1_Click()
    If Check1.Value = 0 Then
        SetWindowPos Me.hwnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    ElseIf Check1.Value = 1 Then
        SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub

Private Sub Check2_Click()
    If ListView1.Tag = 0 Then
        Call CNNew
    Else
        nChildNewEx ListView1.Tag
    End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        Load State
        State.Show
        If Check1.Value = 1 Then SetWindowPos State.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        Unload State
    End If
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    If CmdStr = Date Then AutoUpdate Me
    Cancel = 0
End Sub

Private Sub mainReadme_Click()
    Load About
    About.Show
End Sub

Private Sub mainSetting_Click()
    Load Setting
    Setting.Show
End Sub

Private Sub nChildNewMenu_Click()
    Call CNNew
End Sub

Private Sub nCloseWindowByEndTask_Click()
    Dim ret As Long

    ret = EndTask(CLng(ListView1.SelectedItem.SubItems(2)), 0, 1)
    
    If ret <> 1 Then MsgBox "关闭窗口失败!", 0, "失败": Exit Sub
    
    Call CNNew
End Sub

Private Sub nCloseWindowByMessage_Click()
    FxBombWindow CLng(ListView1.SelectedItem.SubItems(2))
    
    Call CNNew
End Sub

Private Sub nFdNewByMessage_click()
    '函数名：FdEnumWindowsByMessage
    Dim hwnd As Long
    Dim tmp As Long
    'hWnd = 2 ^ 31 - 1
    'Stop
    If MsgBox("此操作需要耗费较长的时间、可能会造成系统资源使用率的短时间的增高、确定要继续此操作？", vbQuestion + vbYesNo, "提示") = vbYes Then
        'MsgBox "这里你在列表框添加信息、、另外TID、PID我不知道怎么获得"
        'ListView1.ColumnHeaders.Clear
        ListView1.ListItems.Clear
        'For hwnd = 0 To 2 ^ 31 - 1
        For hwnd = 0 To 10000000
            If PostMessage(hwnd, 0, 0, 0) = 0 Then
                tmp = tmp + 1
                If tmp > 9999999 Then Exit For
            Else
                Call EnumWindowsProc(hwnd, 0)
                'DoEvents '在像listview添加信息时用
            End If
            DoEvents
        Next
    End If
End Sub

Private Sub Form_Load()
    'SetIcon Me.hwnd, "IDR_MAINFRAME", True 'icon
    
    With ListView1.ColumnHeaders
        .Add , , "窗口", 3711
        .Add , , "类名", 3711
        .Add , , "句柄", 930
        .Add , , "父窗句柄", 910
        .Add , , "PID", 660
        .Add , , "TID", 660
        .Add , , "状态", 1200
    End With
    
    With ListView2.ColumnHeaders
        .Add , , "进程名", 1500
        .Add , , "进程ID", 920
        .Add , , "父进程ID", 960
        .Add , , "PEB", 1300
        .Add , , "EPROCESS", 1300
        .Add , , "优先级", 920
        .Add , , "内存使用", 1800
        .Add , , "映像路径", 4200
        .Add , , "命令行", 4500
        .Add , , "文件厂商", 3400
    End With
       
    With LVServer.ColumnHeaders
        .Add , , "名称", 2600
        .Add , , "状态", 1000
        .Add , , "启动类型", 1000
        .Add , , "路径", 3500   '2000
        .Add , , "描述", 4000
        .Add , , "登录身份", 1400
        .Add , , "动态链接库路径", 3500   '1400
    
    End With
    'Dim strComputer, strNameSpace, strClass As String
    'Dim objSWbemLocator, objSWbemServices As Object
    
'    strComputer = "."           '计算机名，.为本机
'    strNameSpace = "root\cimv2" '指定命名空间为root\cimv2
'    strClass = "Win32_Service"  '指定类为Win32_Service
'    Set objSWbemLocator = CreateObject("WbemScripting.SWbemLocator")    '建立1个WBEM对象的引用指针
'    Set objSWbemServices = objSWbemLocator.ConnectServer(strComputer, strNameSpace)  '连接到指定计算机、命名空间的WMI，返回一个对 SWbemServices 对象的引用
    'RefreshList '刷新服务列表
    

    FirstFocus = True
    ListView1.Tag = 0
    
    If Check1.Value = 1 Then SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

    'ListViewColor Me, ListView1
    'ListViewColor Me, ListView2
    'ListViewColor Me, LVServer
    ListView1.Sorted = True
    SetTextColor Me
    'Label2.DragIcon = 15
    Me.Caption = "Azmrk WindowsXP Edition v" & App.Major & "." & App.Minor & "." & App.Revision
    
    'SetWindowPos Me.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    '2010-02-24  界面布置
    Label1.top = Text1.top + 50
    '注册表相关
    On Error GoTo errTrap ' just in case ;)

    'cSplit.Initialise picSplitter, Me  '分隔条
    Dim NodX As Object
    Set NodX = tvwKeys.Nodes.Add(, , "COMP", "我的电脑", 5) '添加我的电脑节点
    NodX.Expanded = True '
    
    'Set lastNode = NodX
    
    '主键
    Set NodX = tvwKeys.Nodes.Add("COMP", tvwChild, "HKEY_CLASSES_ROOT", "HKEY_CLASSES_ROOT", 1)
    Set NodX = tvwKeys.Nodes.Add("COMP", tvwChild, "HKEY_CURRENT_USER", "HKEY_CURRENT_USER", 1)
    Set NodX = tvwKeys.Nodes.Add("COMP", tvwChild, "HKEY_LOCAL_MACHINE", "HKEY_LOCAL_MACHINE", 1)
    Set NodX = tvwKeys.Nodes.Add("COMP", tvwChild, "HKEY_USERS", "HKEY_USERS", 1)
    Set NodX = tvwKeys.Nodes.Add("COMP", tvwChild, "HKEY_CURRENT_CONFIG", "HKEY_CURRENT_CONFIG", 1)
    Set NodX = tvwKeys.Nodes.Add("COMP", tvwChild, "HKEY_DYN_DATA", "HKEY_DYN_DATA", 1)
    '
    Set NodX = tvwKeys.Nodes.Add("HKEY_CLASSES_ROOT", tvwChild)
    Set NodX = tvwKeys.Nodes.Add("HKEY_CURRENT_USER", tvwChild)
    Set NodX = tvwKeys.Nodes.Add("HKEY_LOCAL_MACHINE", tvwChild)
    Set NodX = tvwKeys.Nodes.Add("HKEY_USERS", tvwChild)
    Set NodX = tvwKeys.Nodes.Add("HKEY_CURRENT_CONFIG", tvwChild)
    Set NodX = tvwKeys.Nodes.Add("HKEY_DYN_DATA", tvwChild)
    
    Set NodX = Nothing
    
    If SoftValue(0) = "1" Then Check1.Value = 1: Call Check1_Click
    If SoftValue(1) = "1" Then Check2.Value = 1: Call Check2_Click
    If SoftValue(2) = "1" Then Check3.Value = 1: Call Check3_Click
    
    Exit Sub
errTrap:
    Dim msg As String
    msg = "未知错误!" & vbCrLf
    msg = msg & "描述: " & Err.Description & String(2, vbCrLf)
    MsgBox msg, vbExclamation, "错误: " & Err.Number
End Sub
    


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload State
    Unload LoginPic
    Unload ModuleList
    Unload ThreadList
    Unload Setting
    'End
End Sub

Private Sub Label2_Click()
    ShellExecute Me.hwnd, "open", "http://hi.baidu.com/dazzles", vbNullString, vbNullString, SW_HIDE
    'http://sighttp.qq.com/cgi-bin/check?sigkey=10e2f1de4f3638083759f062e8997cd18e83e614ff27ed6511e0665cc7ab711b
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.MousePointer = 99
    Me.MouseIcon = VB.LoadResPicture(101, vbResCursor)
End Sub

Private Sub Label4_Click()
    ShellExecute Me.hwnd, "open", "http://hi.baidu.com/naylonslain", vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.MousePointer = 99
    Me.MouseIcon = VB.LoadResPicture(101, vbResCursor)
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LVAutoOrder ListView1, ColumnHeader
End Sub

Private Sub ListView1_DblClick()
    Dim news As Long
    Dim nIndex As Long
    
    Text1.Text = ""
    
    With ListView1
        If (.Tag = 0) Or (.Tag <> 0 And .SelectedItem.Text <> "..") Then   '父窗浏览转入子窗浏览
            nSelectedItemIndex(.Tag) = .SelectedItem.Index   '记录当前选择项的序号
            .Tag = .Tag + 1
            nSelectedItem(.Tag) = CLng(.SelectedItem.SubItems(2))   '记录当前选择项的句柄
            .ListItems.Clear
            EnumAllChildWindows nSelectedItem(.Tag), ""
        ElseIf .Tag = 1 And .SelectedItem.Text = ".." Then   '子窗浏览转入父窗浏览
            .Tag = 0
            Text1.Text = ""
            Call CNNew
            FxSetListviewNowLine ListView1, nSelectedItemIndex(.Tag)
        ElseIf .Tag > 1 And .SelectedItem.Text = ".." Then   '孙窗浏览转入子窗浏览
            .Tag = .Tag - 1
            .ListItems.Clear
            EnumAllChildWindows nSelectedItem(.Tag), ""
            FxSetListviewNowLine ListView1, nSelectedItemIndex(.Tag)
        End If
    End With
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim SubText As String

    With ListView1
        If Button = 2 Then
            SubText = .SelectedItem.SubItems(6)
            
            If .Tag > 0 And .SelectedItem.SubItems(1) = "" Then Exit Sub

            '设置菜单选项
            If InStr(SubText, "隐藏") > 0 Then 'show and hide
                nHide.Enabled = False
                nShow.Enabled = True
            ElseIf InStr(SubText, "可见") > 0 Then
                nHide.Enabled = True
                nShow.Enabled = False
            End If
            
            If InStr(SubText, "激活") > 0 Then 'enabled or not
                nEnableF.Enabled = True
                nEnableT.Enabled = False
            ElseIf InStr(SubText, "冻结") > 0 Then
                nEnableF.Enabled = False
                nEnableT.Enabled = True
            End If
            
            If InStr(SubText, "最大") > 0 Or InStr(SubText, "激活") > 0 Then
                nWindowMax.Enabled = False
                nWindowMin.Enabled = True
            ElseIf InStr(SubText, "最小") > 0 Or InStr(SubText, "激活") > 0 Then
                nWindowMax.Enabled = True
                nWindowMin.Enabled = False
            End If
            
            If .Tag = 0 Then
                nJumpToParent.Visible = True
                nViewFatherWindows.Visible = False
                n09.Visible = True
                nTop.Visible = True
                nNoTop.Visible = True
                nNewMenu.Visible = True
                nChildNewMenu.Visible = False
            ElseIf .Tag > 0 Then
                nJumpToParent.Visible = False
                nViewFatherWindows.Visible = True
                n09.Visible = False
                nTop.Visible = False
                nNoTop.Visible = False
                nNewMenu.Visible = False
                nChildNewMenu.Visible = True
            End If
            
            PopupMenu FatherMenus
        End If
    End With
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LVAutoOrder ListView2, ColumnHeader
End Sub

Private Sub ListView2_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)

    If Button = 2 Then
        '设置菜单状态
        If ListView2.Tag = 0 Then
            Dim SubText As String
            SubText = ListView2.SelectedItem.SubItems(4)
            If InStr(SubText, "较高") > 0 Then
                pPriorityHigh.Enabled = False
                pPriorityNormal.Enabled = True
                pPriorityLow.Enabled = True
            ElseIf InStr(SubText, "标准") > 0 Then
                pPriorityHigh.Enabled = True
                pPriorityNormal.Enabled = False
                pPriorityLow.Enabled = True
            ElseIf InStr(SubText, "较低") > 0 Then
                pPriorityHigh.Enabled = True
                pPriorityNormal.Enabled = True
                pPriorityLow.Enabled = False
            End If
        Else
            pSetPriority.Enabled = False
        End If
        
        PopupMenu pMenu
        pSetPriority.Enabled = True
    End If

    'ListViewColor Me, ListView2
End Sub

Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LVAutoOrder LVServer, ColumnHeader
End Sub

Private Sub ListView3_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu sMenu
    End If
End Sub

Private Sub LVServer_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             x As Single, _
                             y As Single)

    If Button = 2 Then

        Select Case LVServer.SelectedItem.SubItems(1)

            Case "已启动"
                MenuStartServer.Enabled = False
                MenuPauseServer.Enabled = True
                MenuStopServer.Enabled = True

            Case "已停止"
                MenuStartServer.Enabled = True
                MenuPauseServer.Enabled = False
                MenuStopServer.Enabled = False

        End Select
        
        Select Case LVServer.SelectedItem.SubItems(2)
        
            Case "自动"
                MenuSetAuto.Enabled = False
                MenuSetUser.Enabled = True
                MenuSetCant.Enabled = True

            Case "手动"
                MenuSetAuto.Enabled = True
                MenuSetUser.Enabled = False
                MenuSetCant.Enabled = True

            Case "禁用"
                MenuSetAuto.Enabled = True
                MenuSetUser.Enabled = True
                MenuSetUser.Enabled = False
        
        End Select

        Me.PopupMenu sMenu
    End If

End Sub

Private Sub MenuPauseServer_Click()
    Dim Registry As clsRegistry
    Set Registry = New clsRegistry
    Dim r_initial   As String
    Dim rv_value    As String
    Dim serv_status As String
    
    r_initial = LVServer.SelectedItem

    If r_initial = "" Then Exit Sub
    rv_value = Registry.GetValue(eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\" & r_initial, "Description")

    If rv_value = "" Then
        ServicePause "", r_initial
        msNew_Click
        Exit Sub
    End If

    ServicePause "", r_initial
    msNew_Click
End Sub

Private Sub MenuSetAuto_Click()
    SetServerBootType LVServer.SelectedItem, 2
End Sub

Private Sub MenuSetCant_Click()
    SetServerBootType LVServer.SelectedItem, 4
End Sub

Private Sub MenuSetUser_Click()
    SetServerBootType LVServer.SelectedItem, 3
End Sub

Private Sub MenuStartServer_Click()
    Dim Registry As clsRegistry
    Set Registry = New clsRegistry
    Dim r_initial   As String
    Dim rv_value    As String
    Dim serv_status As String
    
    r_initial = LVServer.SelectedItem

    If r_initial = "" Then Exit Sub
    rv_value = Registry.GetValue(eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\" & r_initial, "Description")

    If rv_value = "" Then
        ServiceStart "", r_initial
        msNew_Click
        Exit Sub
    End If
    ServiceStart "", r_initial
    msNew_Click
End Sub

Private Sub MenuStopServer_Click()
    Dim Registry As clsRegistry
    Set Registry = New clsRegistry
    Dim r_initial   As String
    Dim rv_value    As String
    Dim serv_status As String
    
    r_initial = LVServer.SelectedItem

    If r_initial = "" Then Exit Sub
    rv_value = Registry.GetValue(eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\" & r_initial, "Description")

    If rv_value = "" Then
        ServiceStop "", r_initial
        msNew_Click
        Exit Sub
    End If

    ServiceStop "", r_initial
    msNew_Click
End Sub

Private Sub nAmend_Click()
    Dim newText As String
   
    newText = InputBox("请键入新的内容：", "修改标题", ListView1.SelectedItem.Text)
    If newText = "" Or newText = ListView1.SelectedItem.Text Then
        MsgBox "请修改内容！且值不能为空！", vbOKOnly + vbInformation, "警告"
        Exit Sub
    End If
    SetWindowText ListView1.SelectedItem.SubItems(2), newText   'ListView1.SelectedItem.SubItems(2)
    Call CNNew
End Sub

Private Sub nChildNewEx(ByVal Index As Long)
    Dim nIndex As Long
    
    nIndex = FxGetListviewNowLine(ListView1)
    
    ListView1.ListItems.Clear
    EnumAllChildWindows nSelectedItem(Index), ""

    FxSetListviewNowLine ListView1, nIndex
End Sub

Private Sub nClose_Click()
    'EnableWindow CLng(ListView1.SelectedItem.SubItems(2)), 0
    PostMessage CLng(ListView1.SelectedItem.SubItems(2)), WM_CLOSE, 0, ByVal 0
    PostMessage CLng(ListView1.SelectedItem.SubItems(2)), WM_DESTROY, 0, ByVal 0
    
    Call CNNew
End Sub

Private Sub nCloseWindowByParent_Click()
    FxCloseWindowByParent CLng(ListView1.SelectedItem.SubItems(2))
    
    Call CNNew
End Sub

Private Sub nCloseWindowByWndProc_Click()
    FxCloseWindowByWndProc CLng(ListView1.SelectedItem.SubItems(2))
    
    Call CNNew
End Sub

Private Sub nCopyClass_Click()
    Clipboard.Clear
    Clipboard.SetText ListView1.SelectedItem.SubItems(1), 1
End Sub

Private Sub nCopyHandle_Click()
    Clipboard.Clear
    Clipboard.SetText ListView1.SelectedItem.SubItems(2), 1
End Sub

Private Sub nCopyName_Click()
    Clipboard.Clear
    Clipboard.SetText ListView1.SelectedItem.Text, 1
End Sub

Private Sub nEnableF_Click()
    EnableWindow CLng(ListView1.SelectedItem.SubItems(2)), 0
    
    Call CNNew
End Sub

Private Sub nEnableT_Click()
    EnableWindow CLng(ListView1.SelectedItem.SubItems(2)), 1
    
    Call CNNew
End Sub

Private Sub nFxNew_Click()
    Label1.Tag = 1
    
    Call CNNew
End Sub

Private Sub nHide_Click()
    ShowWindow CLng(ListView1.SelectedItem.SubItems(2)), SW_HIDE  'SW_HIDE=0
    
    Call CNNew
End Sub

Private Sub nJumpToParent_Click()
    Dim myId As Long
    Dim i As Long
    
    myId = CLng(ListView1.SelectedItem.SubItems(3))
    
    For i = 1 To ListView1.ListItems.Count
        If CLng(ListView1.ListItems(i).SubItems(2)) = myId Then
            ListView1.ListItems(i).Selected = True
            ListView1.ListItems(i).EnsureVisible
            Exit For
            Exit Sub
        End If
    Next i
End Sub

Private Sub nJumpToTasklist_Click()
    Dim i As Long
    Dim jmpid As Long
    
    jmpid = CLng(ListView1.SelectedItem.SubItems(4))
    
    For i = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(i).SubItems(1) = jmpid Then
            FxSetListviewNowLine ListView2, i
            SSTab1.Tab = 1
        End If
    Next i
End Sub

Private Sub nMove_Click()
    Dim MoveTo As String
    
    MoveTo = InputBox("请输入新的坐标，逗号隔开。例：303,505", "设置坐标")
    If CLng(InStr(1, MoveTo, ",")) = 0 Or MoveTo = "" Then Exit Sub
    SetWindowPos CLng(ListView1.SelectedItem.SubItems(2)), 0, CLng(Mid(MoveTo, 1, InStr(1, MoveTo, ",") - 1)), CLng(Mid(MoveTo, InStr(1, MoveTo, ",") + 1, Len(MoveTo))), 0, 0, SWP_NOSIZE
End Sub

Private Sub nNew_Click()
    Label1.Tag = 0
    
    Call CNNew
End Sub

Private Sub nNoTop_Click()
    SetWindowPos CLng(ListView1.SelectedItem.SubItems(2)), HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    Call CNNew
End Sub

Private Sub nShow_Click()
    ShowWindow CLng(ListView1.SelectedItem.SubItems(2)), SW_SHOW  'SW_SHOW=1;SW_SHOWNOACTIVATE=4
    
    Call CNNew
End Sub

Private Sub nTop_Click()
    SetWindowPos CLng(ListView1.SelectedItem.SubItems(2)), HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    ShowWindow CLng(ListView1.SelectedItem.SubItems(2)), SW_SHOWNOACTIVATE
    
    Call CNNew
End Sub

Private Sub nViewChildWindows_Click()
    Call ListView1_DblClick
End Sub

Private Sub nViewFatherWindows_Click()
    Dim myHwnd As Long
    
    With ListView1
        myHwnd = .SelectedItem.SubItems(3)

        .ListItems(1).Selected = True
        
        Call ListView1_DblClick
        
        Dim i As Long
        
        For i = 1 To .ListItems.Count
            If .ListItems(i).SubItems(2) = myHwnd Then
                FxSetListviewNowLine ListView1, i
                Exit For
            End If
        Next i
    End With
End Sub

Private Sub nWindowMax_Click()
    ShowWindow CLng(ListView1.SelectedItem.SubItems(2)), SW_MAXIMIZE
    
    Call CNNew
End Sub

Private Sub nWindowMin_Click()
    ShowWindow CLng(ListView1.SelectedItem.SubItems(2)), SW_MINIMIZE
    
    Call CNNew
End Sub

Private Sub pcNewTask_Click()
    MsgBox Hex(PROCESS_ALL_ACCESS)
End Sub

Private Sub pCopyCommandLine_Click()
    With Clipboard
        .Clear
        .SetText ListView2.SelectedItem.SubItems(8)
    End With
End Sub

Private Sub pCopyEPROCESS_Click()
    With Clipboard
        .Clear
        .SetText ListView2.SelectedItem.SubItems(4)
    End With
End Sub

Private Sub pCopyName_Click()
    With Clipboard
        .Clear
        .SetText ListView2.SelectedItem.Text
    End With
End Sub

Private Sub pCopyPath_Click()
    With Clipboard
        .Clear
        .SetText ListView2.SelectedItem.SubItems(7)
    End With
End Sub

Private Sub pCopyPEB_Click()
    With Clipboard
        .Clear
        .SetText ListView2.SelectedItem.SubItems(3)
    End With
End Sub

Private Sub pCopyPid_Click()
    With Clipboard
        .Clear
        .SetText CStr(ListView2.SelectedItem.SubItems(1))
    End With
End Sub

Private Sub pJumpToParent_Click()
    Dim myId As Long
    Dim i As Long
    
    myId = CLng(ListView2.SelectedItem.SubItems(2))
    
    For i = 1 To ListView2.ListItems.Count
        If ListView2.ListItems(i).SubItems(1) = myId Then
            ListView2.ListItems(i).Selected = True
            ListView2.ListItems(i).EnsureVisible
            Exit For
            Exit Sub
        End If
    Next i
End Sub

Private Sub pListModule_Click()
    On Error Resume Next
       
    nsItem = CLng(ListView2.SelectedItem.SubItems(1))
    
    Unload ModuleList
    Load ModuleList
    
    If Check1.Value = 1 Then SetWindowPos ModuleList.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    ModuleList.Show
End Sub

Private Sub pListThread_Click()
    On Error Resume Next

    nsItem = CLng(ListView2.SelectedItem.SubItems(1))

    Unload ThreadList
    Load ThreadList
    
    If Check1.Value = 1 Then SetWindowPos ThreadList.hwnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    ThreadList.Show
End Sub

Private Sub pListWindows_Click()
    On Error Resume Next
    
    viewProcessWindows = CLng(ListView2.SelectedItem.SubItems(1))
    SSTab1.Tab = 0
    ListView1.Tag = 0
    'Text1.Text = "输入标题或类名或句柄查找"
    Call CNNew
    viewProcessWindows = 0
End Sub

Private Sub pMoreInformation_Click()
    ShellExecute Me.hwnd, "open", "http://www.baidu.com/s?wd=" & (ListView2.SelectedItem.Text), vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub pNewHf_Click()
    ListView2.Tag = 1
    Call PNNew
End Sub

Private Sub pNewBySession_Click()
    ListView2.Tag = 1
    Call PNNew
End Sub

Private Sub pNewSh_Click()
    ListView2.Tag = 0
    Call PNNew
End Sub

Private Sub pNewWmi_Click()
    ListView2.Tag = 3
    Call PNNew
End Sub

Private Sub pPriorityHigh_Click()
    Dim hProcess As Long
    
    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, CLng(ListView2.SelectedItem.SubItems(1)))
    SetPriorityClass hProcess, HIGH_PRIORITY_CLASS
       
    ZwClose hProcess
    Call PNNew
End Sub

Private Sub pPriorityLow_Click()
    Dim hProcess As Long
    
    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, CLng(ListView2.SelectedItem.SubItems(1)))
    SetPriorityClass hProcess, IDLE_PRIORITY_CLASS
    
    ZwClose hProcess
    Call PNNew
End Sub

Private Sub pPriorityNormal_Click()
    Dim hProcess As Long
    
    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, CLng(ListView2.SelectedItem.SubItems(1)))
    SetPriorityClass hProcess, NORMAL_PRIORITY_CLASS
    
    ZwClose hProcess
    Call PNNew
End Sub

Private Sub pResumeProcess_Click()
    'SusResProcess ListView2.SelectedItem.SubItems(1), False
    'DoEvents
    '---以上是通过恢复进程的所有线程---
    
    Dim hProcess As Long
    
    hProcess = FxNormalOpenProcess(PROCESS_SUSPEND_RESUME, CLng(ListView2.SelectedItem.SubItems(1)))
    ZwResumeProcess hProcess
    
    ZwClose hProcess
    Call PNNew
End Sub

Private Sub pSuspendProcess_Click()
    'SusResProcess ListView2.SelectedItem.SubItems(1), True
    'DoEvents
    '---以上是通过挂起进程的所有线程---
    
    Dim hProcess As Long
    
    hProcess = FxNormalOpenProcess(PROCESS_SUSPEND_RESUME, CLng(ListView2.SelectedItem.SubItems(1)))
    ZwSuspendProcess hProcess
    
    ZwClose hProcess
    Call PNNew
End Sub

Private Sub pTerminateProcessByDebugProcess_Click()
    FxTerminateProcessByDebugProcess CLng(ListView2.SelectedItem.SubItems(1))
    Call PNNew
End Sub

Private Sub pTerminateProcessByRemoteThread_Click()
    Dim hProcess, hThread, hFunction As Long
    Dim lpThreadAttributes As SECURITY_ATTRIBUTES
    
    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, CLng(ListView2.SelectedItem.SubItems(1)))
    If hProcess = 0 Then
        MsgBox "拒绝访问!", 0, "失败"
        Exit Sub
    End If
    
    hFunction = GetModuleHandle("kernel32.dll")
    hFunction = GetProcAddress(hFunction, "ExitProcess")
    
    hThread = Err_CreateRemoteThread(hProcess, lpThreadAttributes, 0, hFunction&, 0, 0, 0)
    'hThread = CreateRemoteThread(hProcess, lpThreadAttributes, 0, hFunction, hModule, 0, 0)
    If hThread = 0 Then
        MsgBox "创建线程失败!", 0, "失败"
        ZwClose hProcess
        Exit Sub
    End If
    
    WaitForSingleObject hThread, INFINITE
    
    ZwClose hThread
    ZwClose hProcess
    
    Call PNNew
End Sub

Private Sub pTerminateProcessNormal_Click()
    Dim hProcess As Long
    
    hProcess = FxNormalOpenProcess(PROCESS_TERMINATE, CLng(ListView2.SelectedItem.SubItems(1)))

    If hProcess = 0 Then
        MsgBox "拒绝访问!", 0, "失败"
        Exit Sub
    End If
    
    ZwTerminateProcess hProcess, 0
    
    WaitForSingleObject hProcess, INFINITE
    
    ZwClose hProcess
    
    Call PNNew
End Sub

Private Sub pWinStationTerminateProcess_Click()
    Dim ret As Long
    
    ret = WinStationTerminateProcess(WTS_CURRENT_SERVER_HANDLE, CLng(ListView2.SelectedItem.SubItems(1)), 0)
    
    If ret <> 1 Then MsgBox "结束进程失败!", 0, "失败": Exit Sub
    
    Call PNNew
End Sub

Private Sub sCopyServiceDescribe_Click()
    Clipboard.Clear
    Clipboard.SetText LVServer.SelectedItem.SubItems(4), 1
End Sub

Private Sub sCopyServiceDllPath_Click()
    Clipboard.Clear
    Clipboard.SetText LVServer.SelectedItem.SubItems(6), 1
End Sub

Private Sub sCopyServiceExePath_Click()
    Clipboard.Clear
    Clipboard.SetText LVServer.SelectedItem.SubItems(3), 1
End Sub

Private Sub sCopyServiceName_Click()
    Clipboard.Clear
    Clipboard.SetText LVServer.SelectedItem.Text, 1
End Sub

Private Sub sDllNature_Click()
    If Not LVServer.SelectedItem.SubItems(6) = "" Then
        ShowFileProperties LVServer.SelectedItem.SubItems(6)
    Else
        MsgBox "没有找到文件！", vbOKOnly + vbInformation, "警告"
    End If
End Sub

Private Sub sExeNature_Click()
    If Not LVServer.SelectedItem.SubItems(3) = "" Then
        ShowFileProperties LVServer.SelectedItem.SubItems(3)
    Else
        MsgBox "没有找到文件！", vbOKOnly + vbInformation, "警告"
    End If
End Sub

Private Sub sMoreInformation_Click()
    ShellExecute Me.hwnd, "open", "http://www.baidu.com/s?wd=" & (LVServer.SelectedItem.Text), vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub sNew_Click()
    Call msNew_Click
End Sub

Private Sub sSelectDll_Click()
    If Not LVServer.SelectedItem.SubItems(6) = "" Then
        FindFiles LVServer.SelectedItem.SubItems(6)
    Else
        MsgBox "没有找到文件！", vbOKOnly + vbInformation, "警告"
    End If
End Sub

Private Sub sSelectExe_Click()
    If Not LVServer.SelectedItem.SubItems(3) = "" Then
        FindFiles LVServer.SelectedItem.SubItems(3)
    Else
        MsgBox "没有找到文件！", vbOKOnly + vbInformation, "警告"
    End If
End Sub

Private Sub SSTab1_Click(PreviousTab As Integer)
    If SSTab1.Tab = 0 Then
        Call CNNew
    ElseIf SSTab1.Tab = 1 Then
        Call PNNew
    ElseIf SSTab1.Tab = 2 Then
        Call msNew_Click
    ElseIf SSTab1.Tab >= 3 Then
        MsgBox "暂时不支持。", vbInformation
        If PreviousTab >= 3 Then
            SSTab1.Tab = 0
        Else
            SSTab1.Tab = PreviousTab
        End If
    End If
End Sub

Private Sub SSTab1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Me.MousePointer = 0
End Sub

Private Sub Text1_Change()
    If ListView1.Tag = 0 Then
        Call CNNew
    ElseIf ListView1.Tag = 1 Then
        EnumAllChildWindows nSelectedItem(ListView1.Tag), Text1.Text
    End If
End Sub

Private Sub Text1_GotFocus()
    If FirstFocus = True Then
        Text1.Text = ""
        FirstFocus = False
    End If
End Sub

Public Function SetVisual(ByRef Visuals() As String, ByRef Soft() As String) '设置外观
    On Error GoTo 0
    
    Dim i    As Long ', ok As Boolean
    Dim temp As String
    ReadINI "Visual settings", "Skin", temp
    Setting.chk.Value = Val(temp)

    If Setting.chk.Value = 0 Then

        For i = 0 To Me.Count - 1

            If InStr(Me.Controls(i).Name, "Slider") > 0 Then Me.Controls(i).Enabled = False
        Next

        Exit Function
    End If

    SkinH_Attach  'skin
    SkinH_SetAero 1 'skin

    For i = 0 To 2

        If IsNumeric(Visuals(i)) Then
            'ok = False
            'Else:
            'ok = True
            Controls("Slider" & i + 1).Value = Val(Visuals(i))
        End If

        If IsNumeric(Soft(i)) Then Controls("check" & i + 1).Value = Val(Soft(i))
    Next

    SkinH_AdjustHSV Setting.Slider1.Value, Setting.Slider2.Value, Setting.Slider3.Value

    For i = 3 To 9

        If IsNumeric(Visuals(i)) Then
            'ok = False
            'Else
            'ok = True
            Controls("Slider" & i + 1).Value = Val(Visuals(i))
        End If

    Next

    SkinH_AdjustAero Setting.Slider4.Value, Setting.Slider7.Value, Setting.Slider6.Value, Setting.Slider5.Value, 0, 0, Setting.Slider8.Value, Setting.Slider9.Value, Setting.Slider10.Value

    If IsNumeric(Visuals(10)) Then Setting.Slider11.Value = Visuals(10)
    SkinH_SetMenuAlpha Setting.Slider11.Value
End Function

Private Sub ServiceStart(ComputerName As String, ServiceName As String) 'start server
    Dim ServiceStatus As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim res As Long

    hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
        If hService <> 0 Then
            res = StartService(hService, 0, 0)
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If
End Sub
Private Sub ServiceStop(ComputerName As String, ServiceName As String) 'stop
    Dim ServiceStatus As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim res As Long

    hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
        If hService <> 0 Then
            res = ControlService(hService, SERVICE_CONTROL_STOP, ServiceStatus)
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If
End Sub
Private Sub ServicePause(ComputerName As String, ServiceName As String) 'pause
    Dim ServiceStatus As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim res As Long

    hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
        If hService <> 0 Then
            res = ControlService(hService, SERVICE_CONTROL_PAUSE, ServiceStatus)
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If
End Sub

Private Sub SetServerBootType(ByVal SubText As String, BootType As Long)
    Dim Reg As clsRegistry
    
    Set Reg = New clsRegistry
    'SubText = LVServer.SelectedItem

    '    If Not Reg.SetValue(eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\" & SubText, "Start", BootType) Then
    '        MsgBox "尝试修改失败、", vbInformation, "提示"
    '    Else
    '        sNew_Click
    '    End If
    Reg.DeleteValue eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\" & SubText, "Start"

    If Not Reg.SetValue(eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\" & SubText, "Start", BootType) Then
        MsgBox "尝试修改失败、", vbInformation, "提示"
    Else
        sNew_Click
    End If
End Sub

Private Sub tvwKeys_NodeCheck(ByVal Node As MSComctlLib.Node)
    MsgBox Node
End Sub

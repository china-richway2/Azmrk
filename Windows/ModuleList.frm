VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "msComDlg32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form ModuleList 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "����ģ��"
   ClientHeight    =   5460
   ClientLeft      =   3645
   ClientTop       =   3885
   ClientWidth     =   10395
   Icon            =   "ModuleList.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "ModuleList.frx":0CCA
   ScaleHeight     =   5460
   ScaleWidth      =   10395
   StartUpPosition =   2  '��Ļ����
   Begin MSComDlg.CommonDialog Dialog1 
      Left            =   5040
      Top             =   1320
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
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
         Name            =   "����"
         Size            =   9
         Charset         =   134
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   0
   End
   Begin VB.Menu mMenu 
      Caption         =   "ģ��˵�"
      Visible         =   0   'False
      Begin VB.Menu mMenuNew 
         Caption         =   "ˢ���б�"
         Begin VB.Menu mNew 
            Caption         =   "ToolHelp32"
         End
         Begin VB.Menu mNewByVirtualMemory 
            Caption         =   "VirtualQueryEx"
         End
         Begin VB.Menu mNewByRead 
            Caption         =   "���ڴ�"
         End
      End
      Begin VB.Menu m01 
         Caption         =   "-"
      End
      Begin VB.Menu mLoadDll 
         Caption         =   "����ģ��"
      End
      Begin VB.Menu MenuUnloadDll 
         Caption         =   "ж��ģ��"
         Begin VB.Menu mUnloadDllByRemoteThread 
            Caption         =   "�߳��ͷ�"
         End
         Begin VB.Menu mUnloadDllByUnmapView 
            Caption         =   "ȡ��Sectionӳ��"
         End
      End
      Begin VB.Menu m02 
         Caption         =   "-"
      End
      Begin VB.Menu mLocationFile 
         Caption         =   "��λ�ļ�"
      End
      Begin VB.Menu mLookNature 
         Caption         =   "�鿴����"
      End
      Begin VB.Menu m03 
         Caption         =   "-"
      End
      Begin VB.Menu MenuCopy 
         Caption         =   "������Ŀ"
         Begin VB.Menu mCopyName 
            Caption         =   "��������"
         End
         Begin VB.Menu mCopyHandle 
            Caption         =   "���ƾ��"
         End
         Begin VB.Menu mCopyPath 
            Caption         =   "����·��"
         End
      End
   End
End
Attribute VB_Name = "ModuleList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim mPid As Long

Private Sub Form_Load()
    ListView1.ColumnHeaders.Add , , "ģ������", 1500
    ListView1.ColumnHeaders.Add , , "ģ����", 1300
    ListView1.ColumnHeaders.Add , , "ģ��·��", 4200
    ListView1.ColumnHeaders.Add , , "�������", 1300
    ListView1.ColumnHeaders.Add , , "ģ���С", 1200
    ListView1.Tag = 0
    
    mPid = nsItem
    
    'ListViewColor Me, ListView1
    'SetTextColor Me
    'SetIcon ModuleList.hwnd, "IDR_MAINFRAME", True
    Dialog1.Filter = "��̬���ӿ��ļ�(*.dll,*.ocx)|*.dll;*.ocx"
    
    Call MNNew(mPid, Me)
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LVAutoOrder ListView1, ColumnHeader
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
    If Button = 2 Then
        PopupMenu mMenu
    End If
End Sub

Private Sub mCopyHandle_Click()
    Clipboard.Clear
    Clipboard.SetText ListView1.SelectedItem.SubItems(1), 1
End Sub

Private Sub mCopyName_Click()
    Clipboard.Clear
    Clipboard.SetText ListView1.SelectedItem.Text
End Sub

Private Sub mCopyPath_Click()
    Clipboard.Clear
    Clipboard.SetText ListView1.SelectedItem.SubItems(2), 1
End Sub

Private Sub mLoadDll_Click()
    Dialog1.InitDir = App.Path
    Dialog1.ShowOpen
    If Dialog1.FileName = "" Then Exit Sub

    If FxRemoteProcessInsertDll(mPid, Dialog1.FileName, True) Then Call MNNew(mPid, Me)
    Dialog1.FileName = ""
End Sub

Private Sub mLocationFile_Click()
    'MsgBox ListView1.SelectedItem.SubItems(2)
    'Shell "explorer.exe /select," & (ListView1.SelectedItem.SubItems(2)), vbNormalFocus
    FindFiles ListView1.SelectedItem.SubItems(2)
End Sub

Private Sub mLookNature_Click()
    ShowFileProperties ListView1.SelectedItem.SubItems(2)
End Sub

Private Sub mNew_Click()
    ListAllModules mPid, Me
    ListView1.Tag = 0
End Sub

Private Sub mNewByRead_Click()
    ListView1.Tag = 2
    MNNew mPid, Me
End Sub

Private Sub mNewByVirtualMemory_Click()
    FxEnumModulesByVirtualMemory mPid, Me
    ListView1.Tag = 1
End Sub

Private Sub mUnloadDllByRemoteThread_Click()
    FxRemoteProcessFreeDll mPid, UnFormatHex(ListView1.SelectedItem.SubItems(1))
    Call MNNew(mPid, Me)
End Sub

Private Sub mUnloadDllByUnmapView_Click()
    'FxUnloadDllByUnmapView mPid, 0, ListView1.SelectedItem.Text
    FxUnloadDllByUnmapView mPid, UnFormatHex(ListView1.SelectedItem.SubItems(1)), 0
    Call MNNew(mPid, Me)
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form HandleList 
   Caption         =   "目标进程打开的句柄"
   ClientHeight    =   5130
   ClientLeft      =   165
   ClientTop       =   555
   ClientWidth     =   8400
   LinkTopic       =   "Form1"
   ScaleHeight     =   5130
   ScaleWidth      =   8400
   StartUpPosition =   3  '窗口缺省
   Begin MSComctlLib.ListView ListView1 
      Height          =   630
      Left            =   1320
      TabIndex        =   0
      Top             =   600
      Width           =   1110
      _ExtentX        =   1958
      _ExtentY        =   1111
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "类型"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "pObject"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "句柄属性"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "句柄权限"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "句柄"
         Object.Width           =   1940
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "CreatorBackTraceIndex"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Menu hMenu 
      Caption         =   "菜单"
      Visible         =   0   'False
      Begin VB.Menu hRefresh 
         Caption         =   "刷新"
      End
      Begin VB.Menu hClose 
         Caption         =   "释放"
      End
      Begin VB.Menu pCloseAll 
         Caption         =   "尝试全部释放"
      End
   End
End
Attribute VB_Name = "HandleList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mPid As Long
Dim mEProcess As Long
Private Sub Form_Load()
    If nsItem = 0 Then Exit Sub '防止一些错误
    mEProcess = nsItem
    mPid = PsGetPidByEProcess(nsItem)
    Call HandleTarget
End Sub

Private Sub HandleTarget()
    Dim bytBuf() As Byte
    Dim arySize As Long
    Dim nLine As Long
    nLine = FxGetListviewNowLine(ListView1)
    ListView1.ListItems.Clear
    arySize = 1
    Do
        ReDim bytBuf(arySize)
        st = ZwQuerySystemInformation(SystemHandleInformation, VarPtr(bytBuf(0)), arySize, 0&)
        If (Not NT_SUCCESS(st)) Then
            If (st <> STATUS_INFO_LENGTH_MISMATCH) Then
                Erase bytBuf
                Exit Sub
            End If
        Else
            Exit Do
        End If
        arySize = arySize * 2
        ReDim bytBuf(arySize)
    Loop
    
    Dim NumOfHandle As Long
    NumOfHandle = 0
    CopyMemory ByVal VarPtr(NumOfHandle), ByVal VarPtr(bytBuf(0)), Len(NumOfHandle)
    Dim h_info() As SYSTEM_HANDLE_TABLE_ENTRY_INFO
    ReDim h_info(NumOfHandle)
    CopyMemory ByVal VarPtr(h_info(0)), ByVal VarPtr(bytBuf(0)) + Len(NumOfHandle), Len(h_info(0)) * NumOfHandle
    
    Dim i As Long
    For i = 0 To NumOfHandle - 1
        If h_info(i).UniqueProcessId = mPid Then
            Dim S As String
            S = ObjectTypeNames(h_info(i).ObjectTypeIndex)
            If S = "" Then S = "OB_TYPE_UNKNOWN (" & h_info(i).ObjectTypeIndex & ")"
            With ListView1.ListItems.Add(, , S)
                .SubItems(1) = FormatHex(h_info(i).pObject)
                .SubItems(2) = FormatHex(h_info(i).HandleAttributes)
                .SubItems(3) = FormatHex(h_info(i).GrantedAccess)
                .SubItems(4) = FormatHex(h_info(i).HandleValue)
                .SubItems(5) = FormatHex(h_info(i).CreatorBackTraceIndex)
            End With
        End If
    Next
    FxSetListviewNowLine ListView1, nLine
    Caption = "共有 " & ListView1.ListItems.count & " 个句柄"
End Sub

Private Sub Form_Resize()
    ListView1.Move 0, 0, ScaleWidth, ScaleHeight
End Sub

Private Sub hClose_Click()
    CloseRemoteHandle UnFormatHex(ListView1.SelectedItem.SubItems(4)), mPid, True
    HandleTarget
End Sub

Private Sub hRefresh_Click()
    HandleTarget
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LVAutoOrder ListView1, ColumnHeader
End Sub

Private Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu hMenu
End Sub

Private Sub pCloseAll_Click()
    Dim i As Long, j As Long, k As Control
    On Error Resume Next
    For Each k In Me.Controls
        k.Visible = False
    Next
    For i = 1 To ListView1.ListItems.count
        j = j + CloseRemoteHandle(UnFormatHex(ListView1.ListItems(i).SubItems(4)), mPid, False)
    Next
    For Each k In Me.Controls
        k.Visible = True
    Next
    MsgBox "成功释放 " & j & " 个对象", vbInformation
    HandleTarget
End Sub

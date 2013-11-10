VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form WaitWindow 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Azmrk - 请稍等"
   ClientHeight    =   1185
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   1185
   ScaleWidth      =   4680
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command2 
      Caption         =   "停止"
      Height          =   375
      Left            =   1680
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin MSComctlLib.ProgressBar pbBar 
      Height          =   375
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   661
      _Version        =   393216
      Appearance      =   1
   End
End
Attribute VB_Name = "WaitWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim fStop As Boolean
Public Sub BeginReleaseAll()
    Dim i As Long
    Dim bytBuf() As Byte
    Dim arySize As Long
    Dim st As Long
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
    CopyMemory VarPtr(NumOfHandle), VarPtr(bytBuf(0)), Len(NumOfHandle)
    Dim h_info() As SYSTEM_HANDLE_TABLE_ENTRY_INFO
    ReDim h_info(NumOfHandle)
    CopyMemory VarPtr(h_info(0)), VarPtr(bytBuf(0)) + Len(NumOfHandle), Len(h_info(0)) * NumOfHandle
    Dim j As Long
    pbBar.max = UBound(h_info)
    
    Show
    For i = LBound(h_info) To UBound(h_info)
        With h_info(i)
            j = j + CloseRemoteHandle(.HandleValue, .UniqueProcessId, False)
            'If (i And 31) = 0 Then
                pbBar.Value = i
                DoEvents
                If fStop Then GoTo S
            'End If
        End With
    Next i
S:
    'Erase h_info
    MsgBox "成功释放 " & j & " 个句柄", vbInformation
End Sub

Private Sub Command2_Click()
    fStop = True
    Unload Me
End Sub

Private Sub Form_Load()
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    BeginReleaseAll
End Sub

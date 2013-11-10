VERSION 5.00
Begin VB.Form DebugThreadWindow 
   Caption         =   "Azmrk 调试器 - 线程 "
   ClientHeight    =   4020
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6375
   LinkTopic       =   "Form1"
   ScaleHeight     =   4020
   ScaleWidth      =   6375
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton Command6 
      Caption         =   "单步步过"
      Height          =   495
      Left            =   2400
      TabIndex        =   10
      Top             =   2280
      Width           =   1215
   End
   Begin VB.CommandButton Command5 
      Caption         =   "恢复"
      Height          =   495
      Left            =   2400
      TabIndex        =   9
      Top             =   1680
      Width           =   1215
   End
   Begin VB.CommandButton Command4 
      Caption         =   "挂起"
      Height          =   495
      Left            =   2400
      TabIndex        =   8
      Top             =   1080
      Width           =   1215
   End
   Begin VB.CommandButton Command3 
      Caption         =   "单步步入"
      Height          =   495
      Left            =   2400
      TabIndex        =   6
      Top             =   120
      Width           =   1215
   End
   Begin VB.CommandButton Command2 
      Caption         =   "继续调试"
      Height          =   495
      Left            =   1320
      TabIndex        =   4
      Top             =   2880
      Width           =   1215
   End
   Begin VB.CommandButton Command1 
      Caption         =   "反汇编EIP"
      Height          =   495
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   1215
   End
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   4920
      Top             =   1080
   End
   Begin VB.Frame Frame1 
      Caption         =   "寄存器"
      Height          =   3495
      Left            =   3840
      TabIndex        =   0
      Top             =   120
      Width           =   2175
      Begin VB.PictureBox pbContext 
         Height          =   3135
         Left            =   120
         ScaleHeight     =   3075
         ScaleWidth      =   1755
         TabIndex        =   1
         Top             =   240
         Width           =   1815
      End
   End
   Begin VB.Label Label2 
      Caption         =   "挂起计数：0"
      Height          =   255
      Left            =   2400
      TabIndex        =   7
      Top             =   720
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "未发生异常."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   3480
      Width           =   2415
   End
   Begin VB.Label TimeLabel 
      Caption         =   "运行时间："
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   2520
      Width           =   2175
   End
End
Attribute VB_Name = "DebugThreadWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public hThread As Long, hProcess As Long, dwTid As Long, dwPid As Long
Public Changed As Long
Public lX As Long, lY As Long
Public Ys As Long
Public Parentw As DebugWindow
Public RunTime As Currency
Dim Breakpoints() As Long
Dim nTime As Times, aFlag As FLGSTRUCT, aContext As CONTEXT, rctFlag As RECT, rctRegister As RECT

Public Sub BreakThreadAt(ByVal lpBreakPoint As Long)
    On Error GoTo Es
    Dim i As Long: i = UBound(Breakpoints)
    ReDim Preserve Breakpoints(i + 1)
    Breakpoints(i + 1) = lpBreakPoint
    WriteProcessMemory hProcess, ByVal lpBreakPoint, CByte(&HCC), 1, 0
Es:
    ReDim Breakpoints(0)
    i = -1
    Resume Next
End Sub

Public Sub SetThread(ByVal A As Long, ByVal b As Long, ByVal c As Long, ByVal d As Long, ByVal Owner As DebugWindow)
    dwPid = A
    dwTid = b
    hProcess = c
    hThread = d
    aContext.ContextFlags = CONTEXT_FULL
    Dim i As Long
    Ys = pbContext.TextHeight("")
    rctFlag.left = pbContext.TextWidth("EAX 0x00000000 FF ")
    rctFlag.right = pbContext.TextWidth("EAX 0x00000000 FF 0")
    rctRegister.left = pbContext.TextWidth("EAX ")
    rctRegister.right = pbContext.TextWidth("EAX 0x00000000")
    If Owner.MainTID = dwTid Then
        Caption = "Azmrk 调试器 - 主线程 " & Hex(dwTid)
    Else
        Caption = "Azmrk 调试器 - 线程 " & Hex(dwTid)
    End If
    Set Parentw = Owner
End Sub

Public Sub ChangeContext()
    Dim bContext As CONTEXT, bFlag As FLGSTRUCT, nLength As Long
    ZwSuspendThread hThread, 0
    'Stop
    Call ZwQueryInformationThread(hThread, ThreadTimes, nTime, Len(nTime), nLength)
    RunTime = nTime.KernelTime + nTime.UserTime
    Dim strTimes As String
    strTimes = RunTime
    If Len(strTimes) > 3 Then
        strTimes = left(strTimes, Len(strTimes) - 3) & "." & right(strTimes, 3)
    Else
        strTimes = "0." & strTimes
    End If
    TimeLabel = "运行时间：" & strTimes & "秒"
    bContext.ContextFlags = CONTEXT_FULL
    If ZwGetContextThread(hThread, bContext) < 0 Then Exit Sub
    GetFlags bContext.EFlags, bFlag
    
    If aContext.EAX <> bContext.EAX Then
        Changed = Changed Or &H1
    Else
        Changed = Changed And &HFFFFFFFE
    End If
    If aContext.ECX <> bContext.ECX Then
        Changed = Changed Or &H2
    Else
        Changed = Changed And &HFFFFFFFD
    End If
    If aContext.EDX <> bContext.EDX Then
        Changed = Changed Or &H4
    Else
        Changed = Changed And &HFFFFFFFB
    End If
    If aContext.EBX <> bContext.EBX Then
        Changed = Changed Or &H8
    Else
        Changed = Changed And &HFFFFFFF7
    End If
    If aContext.Esp <> bContext.Esp Then
        Changed = Changed Or &H10
    Else
        Changed = Changed And &HFFFFFFEF
    End If
    If aContext.Ebp <> bContext.Ebp Then
        Changed = Changed Or &H20
    Else
        Changed = Changed And &HFFFFFFDF
    End If
    If aContext.Esi <> bContext.Esi Then
        Changed = Changed Or &H40
    Else
        Changed = Changed And &HFFFFFFBF
    End If
    If aContext.Edi <> bContext.Edi Then
        Changed = Changed Or &H80
    Else
        Changed = Changed And &HFFFFFF8F
    End If
    If aContext.EFlags <> bContext.EFlags Then
        Changed = Changed Or &H100
    Else
        Changed = Changed And &HFFFFFEFF
    End If
    If aContext.Edi <> bContext.Edi Then
        Changed = Changed Or &H200
    Else
        Changed = Changed And &HFFFFFDFF
    End If
    
    If aFlag.fCF <> bFlag.fCF Then
        Changed = Changed Or &H400
    Else
        Changed = Changed And &HFFFFBFF
    End If
    If aFlag.fPF <> bFlag.fPF Then
        Changed = Changed Or &H800
    Else
        Changed = Changed And &HFFFF7FF
    End If
    If aFlag.fAF <> bFlag.fAF Then
        Changed = Changed Or &H1000
    Else
        Changed = Changed And &HFFFEFFF
    End If
    If aFlag.fZF <> bFlag.fZF Then
        Changed = Changed Or &H2000
    Else
        Changed = Changed And &HFFFDFFF
    End If
    If aFlag.fSF <> bFlag.fSF Then
        Changed = Changed Or &H4000
    Else
        Changed = Changed And &HFFFBFFF
    End If
    If aFlag.fTF <> bFlag.fTF Then
        Changed = Changed Or &H8000
    Else
        Changed = Changed And &HFFF7FFF
    End If
    If aFlag.fIF <> bFlag.fIF Then
        Changed = Changed Or &H10000
    Else
        Changed = Changed And &HFFEFFFF
    End If
    If aFlag.fDF <> bFlag.fDF Then
        Changed = Changed Or &H20000
    Else
        Changed = Changed And &HFFDFFFF
    End If
    If aFlag.fOF <> bFlag.fOF Then
        Changed = Changed Or &H40000
    Else
        Changed = Changed And &HFFBFFFF
    End If
    If aFlag.fRF <> bFlag.fRF Then
        Changed = Changed Or &H80000
    Else
        Changed = Changed And &HFF7FFFF
    End If
    
    If aContext.SegEs <> bContext.SegEs Then
        Changed = Changed Or &H100000
    Else
        Changed = Changed And &HFFEFFFFF
    End If
    If aContext.SegCs <> bContext.SegCs Then
        Changed = Changed Or &H200000
    Else
        Changed = Changed And &HFFDFFFFF
    End If
    If aContext.SegSs <> bContext.SegSs Then
        Changed = Changed Or &H400000
    Else
        Changed = Changed And &HFFBFFFFF
    End If
    If aContext.SegDs <> bContext.SegDs Then
        Changed = Changed Or &H800000
    Else
        Changed = Changed And &HFF7FFFFF
    End If
    If aContext.SegFs <> bContext.SegFs Then
        Changed = Changed Or &H1000000
    Else
        Changed = Changed And &HFEFFFFFF
    End If
    If aContext.SegGs <> bContext.SegGs Then
        Changed = Changed Or &H2000000
    Else
        Changed = Changed And &HFDFFFFFF
    End If
    
    aContext = bContext
    aFlag = bFlag
    ZwResumeThread hThread, 0
    'Debug.Assert aContext.Eip = &H7C92E514
End Sub

Public Sub Command1_Click()
    Parentw.hDisasm.JumpTo aContext.EIP
End Sub

Private Sub Command2_Click()
    Parentw.CheckContinue
    Command2.Enabled = False
    Label1 = "未发生异常."
End Sub

Private Sub Command3_Click()
    aContext.EFlags = aContext.EFlags Or FlagTF
    ZwSetContextThread hThread, aContext
    Command2_Click
End Sub

Private Sub Command4_Click()
    Dim nCount As Long
    ZwSuspendThread hThread, nCount
    Label2 = "挂起计数：" & nCount + 1
End Sub

Private Sub Command5_Click()
    Dim nCount As Long
    ZwResumeThread hThread, nCount
    Label2 = "挂起计数：" & nCount - 1
End Sub

Private Sub Command6_Click()
    Dim nOpcode(10) As Byte
    'Stop
    aContext.EFlags = aContext.EFlags And (Not FlagTF)
    ZwSetContextThread hThread, aContext
    ZwReadVirtualMemory hProcess, ByVal aContext.EIP, nOpcode(0), 11, 0
    Dim dis As Disasm
    dis.EIP = VarPtr(nOpcode(0))
    Dim nLength As Long
    nLength = Disasm(dis)
    Dim Back As Byte
    Dim Status As Long
    If dis.Instruction.BranchType = CallType Then
        ZwReadVirtualMemory hProcess, ByVal aContext.EIP + nLength, Back, 1, 0
        WriteProcessMemory hProcess, ByVal aContext.EIP + nLength, CByte(&HCC), 1, 0
        Parentw.Timer1.Enabled = False
        Dim NtEvent As NT_DEBUG_EVENT
        Dim Timeout As INT64
        Timeout.dwHigh = &HFFFFFFFF
        Timeout.dwLow = -10000& * 10
        Parentw.CheckContinue
        Do
            Status = ZwWaitForDebugEvent(Parentw.hDebug, 1, Timeout, NtEvent)
            Dim Buffer(199) As Long
            If Status = 258 Then
                WriteProcessMemory hProcess, ByVal aContext.EIP + nLength, Back, 1, 0
                Exit Do
            End If
            DbgUiConvertStateChangeStructure NtEvent, Buffer(0)
            If Buffer(0) <> EXCEPTION_DEBUG_EVENT Then
                Parentw.DispatchEvent VarPtr(NtEvent)
                Parentw.CheckContinue
            Else
                Dim ex As EXCEPTION_DEBUG_INFO
                If Buffer(2) <> dwTid Then
                    Parentw.DispatchEvent VarPtr(NtEvent)
                    Parentw.CheckContinue
                Else
                    UnionToType ex, Buffer(3), Len(ex)
                    If ex.ExceptionRecord.ExceptionCode <> EXCEPTION_BREAKPOINT Then
                        Parentw.DispatchEvent VarPtr(NtEvent)
                        Parentw.CheckContinue
                    Else
                        WriteProcessMemory hProcess, ByVal aContext.EIP + nLength, Back, 1, 0
                        ZwGetContextThread hThread, aContext
                        aContext.EIP = aContext.EIP - 1
                        ZwSetContextThread hThread, aContext
                        Parentw.DispatchEvent VarPtr(NtEvent)
                        Exit Sub
                    End If
                End If
            End If
        Loop
    Else
        Command3_Click
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Cancel = 1
    Hide
    ZwClose hThread
End Sub

Private Sub pbContext_DblClick()
    '判断点击范围：在Flags上点击
    If lX >= rctFlag.left And lX <= rctFlag.right Then
        Dim i As Long
        i = lY \ Ys
        If i < 9 Then
            Select Case i
            Case 0: aContext.EFlags = aContext.EFlags Xor FlagCF
            Case 1: aContext.EFlags = aContext.EFlags Xor FlagPF
            Case 2: aContext.EFlags = aContext.EFlags Xor FlagAF
            Case 3: aContext.EFlags = aContext.EFlags Xor FlagZF
            Case 4: aContext.EFlags = aContext.EFlags Xor FlagSF
            Case 5: aContext.EFlags = aContext.EFlags Xor FlagTF
            Case 6: aContext.EFlags = aContext.EFlags Xor FlagIF
            Case 7: aContext.EFlags = aContext.EFlags Xor FlagDF
            Case 8: aContext.EFlags = aContext.EFlags Xor FlagOF
            Case 9: aContext.EFlags = aContext.EFlags Xor FlagRF
            Case Else: Exit Sub
            End Select
            ZwSetContextThread hThread, aContext
            Call ChangeContext
            Call pbContext_Paint
            Exit Sub
        End If
    End If
    '如果是在数字上点击...
    If lX >= rctRegister.left And lX <= rctRegister.right Then
        i = lY \ Ys
        If i <= 9 Then Call EditRegister(i)
    End If
End Sub

Private Sub EditRegister(ByVal Num As Long)
    Dim sFpu As String
    Select Case Num
    Case 0: sFpu = "EAX"
    Case 1: sFpu = "ECX"
    Case 2: sFpu = "EDX"
    Case 3: sFpu = "EBX"
    Case 4: sFpu = "ESP"
    Case 5: sFpu = "EBP"
    Case 6: sFpu = "ESI"
    Case 7: sFpu = "EDI"
    Case 8: sFpu = "EFL"
    Case 9: sFpu = "EIP"
    End Select
    Dim nNum As Long
    Select Case Num
    Case 0: nNum = aContext.EAX
    Case 1: nNum = aContext.ECX
    Case 2: nNum = aContext.EDX
    Case 3: nNum = aContext.EBX
    Case 4: nNum = aContext.Esp
    Case 5: nNum = aContext.Ebp
    Case 6: nNum = aContext.Esi
    Case 7: nNum = aContext.Edi
    Case 8: nNum = aContext.EFlags
    Case 9: nNum = aContext.EIP
    End Select
    sFpu = InputBox("请输入新的 " & sFpu & " 值：", "修改 " & sFpu & " 值", nNum)
    If IsNumeric(sFpu) Then
        sNum = CLng(sFpu)
    Else
        Exit Sub
    End If
    Select Case Num
    Case 0: aContext.EAX = nNum
    Case 1: aContext.ECX = nNum
    Case 2: aContext.EDX = nNum
    Case 3: aContext.EBX = nNum
    Case 4: aContext.Esp = nNum
    Case 5: aContext.Ebp = nNum
    Case 6: aContext.Esi = nNum
    Case 7: aContext.Edi = nNum
    Case 8: aContext.EFlags = nNum
    Case 9: aContext.EIP = nNum
    End Select
    ZwSetContextThread hThread, aContext
    MsgBox "OK"
End Sub

Private Sub pbContext_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    lX = X
    lY = Y
End Sub

Private Sub pbContext_Paint()
    Dim X As Long
    X = pbContext.TextWidth("EAX 0x00000000 ")
    With pbContext
        .Cls
        pbContext.Print "EAX ";
        .ForeColor = (Changed And 1&) * vbRed
        pbContext.Print FormatHex(aContext.EAX)
        .ForeColor = vbBlack
        pbContext.Print "ECX ";
        .ForeColor = (Changed And 2&) \ 2 * vbRed
        pbContext.Print FormatHex(aContext.ECX)
        .ForeColor = vbBlack
        pbContext.Print "EDX ";
        .ForeColor = (Changed And 4&) \ 4 * vbRed
        pbContext.Print FormatHex(aContext.EDX)
        .ForeColor = vbBlack
        pbContext.Print "EBX ";
        .ForeColor = (Changed And 8&) \ 8 * vbRed
        pbContext.Print FormatHex(aContext.EBX)
        .ForeColor = vbBlack
        pbContext.Print "ESP ";
        .ForeColor = (Changed And &H10&) \ &H10& * vbRed
        pbContext.Print FormatHex(aContext.Esp)
        .ForeColor = vbBlack
        pbContext.Print "EBP ";
        .ForeColor = (Changed And &H20&) \ &H20& * vbRed
        pbContext.Print FormatHex(aContext.Ebp)
        .ForeColor = vbBlack
        pbContext.Print "ESI ";
        .ForeColor = (Changed And &H40&) \ &H40& * vbRed
        pbContext.Print FormatHex(aContext.Esi)
        .ForeColor = vbBlack
        pbContext.Print "EDI ";
        .ForeColor = (Changed And &H80&) \ &H80& * vbRed
        pbContext.Print FormatHex(aContext.Edi)
        .ForeColor = vbBlack
        pbContext.Print "EFL ";
        .ForeColor = (Changed And &H100&) \ &H100& * vbRed
        pbContext.Print FormatHex(aContext.EFlags)
        .ForeColor = vbBlack
        pbContext.Print "EIP ";
        .ForeColor = (Changed And &H200&) \ &H200& * vbRed
        pbContext.Print FormatHex(aContext.EIP)
        .ForeColor = vbBlack
        
        .CurrentY = 0
        
        .CurrentX = X
        .ForeColor = (Changed And &H400&) \ &H400& * vbRed
        pbContext.Print "CF"; aFlag.fCF
        .ForeColor = vbBlack
        .CurrentX = X
        .ForeColor = (Changed And &H800&) \ &H800& * vbRed
        pbContext.Print "PF"; aFlag.fPF
        .ForeColor = vbBlack
        .CurrentX = X
        .ForeColor = (Changed And &H1000&) \ &H1000& * vbRed
        pbContext.Print "AF"; aFlag.fAF
        .ForeColor = vbBlack
        .CurrentX = X
        .ForeColor = (Changed And &H2000&) \ &H2000& * vbRed
        pbContext.Print "ZF"; aFlag.fZF
        .ForeColor = vbBlack
        .CurrentX = X
        .ForeColor = (Changed And &H4000&) \ &H4000& * vbRed
        pbContext.Print "SF"; aFlag.fSF
        .ForeColor = vbBlack
        .CurrentX = X
        .ForeColor = (Changed And &H8000&) \ &H8000& * vbRed
        pbContext.Print "TF"; aFlag.fTF
        .ForeColor = vbBlack
        .CurrentX = X
        .ForeColor = (Changed And &H10000) \ &H10000 * vbRed
        pbContext.Print "IF"; aFlag.fIF
        .ForeColor = vbBlack
        .CurrentX = X
        .ForeColor = (Changed And &H20000) \ &H20000 * vbRed
        pbContext.Print "DF"; aFlag.fDF
        .ForeColor = vbBlack
        .CurrentX = X
        .ForeColor = (Changed And &H40000) \ &H40000 * vbRed
        pbContext.Print "OF"; aFlag.fOF
        .ForeColor = vbBlack
        .CurrentX = X
        .ForeColor = (Changed And &H80000) \ &H80000 * vbRed
        pbContext.Print "RF"; aFlag.fRF
        
        .CurrentX = 0
        PrintSegment "ES", aContext.SegEs, Changed And &H100000
        PrintSegment "CS", aContext.SegCs, Changed And &H200000
        PrintSegment "SS", aContext.SegSs, Changed And &H400000
        PrintSegment "DS", aContext.SegDs, Changed And &H800000
        PrintSegment "FS", aContext.SegFs, Changed And &H1000000
        PrintSegment "GS", aContext.SegGs, Changed And &H2000000
        
        .ForeColor = vbBlack
    End With
End Sub

Private Sub PrintSegment(ByVal strName As String, ByVal wSegment As Long, ByVal nChanged As Long)
    Dim ldt As LDT_ENTRY
    GetThreadSelectorEntry hThread, wSegment, ldt
    If nChanged Then pbContext.ForeColor = vbRed
    pbContext.Print strName; " "; Hex2(wSegment, 4); " "; Hex2(SegmentToNum(ldt), 8)
    If nChanged Then pbContext.ForeColor = vbBlack
End Sub

Private Sub Timer1_Timer()
    Call ChangeContext
    pbContext_Paint
End Sub

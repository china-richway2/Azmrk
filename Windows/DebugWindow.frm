VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form DebugWindow 
   Caption         =   "Azmrk 调试器 - 进程 "
   ClientHeight    =   3795
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   6030
   LinkTopic       =   "Form1"
   ScaleHeight     =   3795
   ScaleWidth      =   6030
   StartUpPosition =   3  '窗口缺省
   Begin VB.Timer Timer1 
      Interval        =   10
      Left            =   2400
      Top             =   1680
   End
   Begin MSComctlLib.StatusBar Bar 
      Align           =   2  'Align Bottom
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   3420
      Width           =   6030
      _ExtentX        =   10636
      _ExtentY        =   661
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   1
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            Object.Width           =   5080
            MinWidth        =   5080
         EndProperty
      EndProperty
   End
   Begin VB.Menu dMenu 
      Caption         =   "菜单"
      Begin VB.Menu dBreak 
         Caption         =   "Break目标进程"
      End
      Begin VB.Menu dSuspendImmediatly 
         Caption         =   "瞬间挂起"
      End
      Begin VB.Menu dResume 
         Caption         =   "恢复"
      End
   End
End
Attribute VB_Name = "DebugWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Threads() As DebugThread
Dim Modules() As DbgModuleInfo
Public tNum As Long
Public dNum As Long
Public hProcess As Long, hDebug As Long
Public PID As Long, hDisasm As DisasmWindow
Public mNum As Long, MainTID As Long
Public Status As Long, EventNum As Byte
'Status
'0 正在运行
'1 发生异常
'2
Dim de1 As EXCEPTION_DEBUG_INFO
Dim de2 As CREATE_THREAD_DEBUG_INFO
Dim de3 As CREATE_PROCESS_DEBUG_INFO
Dim de4 As EXIT_THREAD_DEBUG_INFO
Dim de5 As EXIT_PROCESS_DEBUG_INFO
Dim de6 As LOAD_DLL_DEBUG_INFO
Dim de7 As UNLOAD_DLL_DEBUG_INFO
Dim de8 As OUTPUT_DEBUG_STRING_INFO
Dim de9 As RIP_INFO
Dim EventCid As CLIENT_ID

Public Sub Attach(ByVal dwProcessId As Long)
    Dim Status As Long, oa As OBJECT_ATTRIBUTES, NtEvent As NT_DEBUG_EVENT
    ReDim Threads(0), Dlls(0), Modules(0)
    Caption = "Azmrk 调试器 - 进程 " & Hex(dwProcessId)
    PID = dwProcessId
    oa.Length = 24
    Status = ZwCreateDebugObject(hDebug, &H1F000F, VarPtr(oa), 0)
    If Not NT_SUCCESS(Status) Then
        MsgBox "创建调试对象失败！", vbCritical
        Exit Sub
    End If
    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, dwProcessId)
    If hProcess = 0 Then Exit Sub
    Set hDisasm = New DisasmWindow
    SetParent hDisasm.hWnd, Me.hWnd
    hDisasm.Show
    hDisasm.SetProcess hProcess, Me
    Status = ZwDebugActiveProcess(hProcess, hDebug)
    If Not NT_SUCCESS(Status) Then
        If Status = STATUS_PORT_ALREADY_SET Then
            MsgBox "无法调试进程，可能是因为其他调试器已经正在调试此进程！", vbCritical
        Else
            MsgBox "调试进程失败！", vbCritical
        End If
        ZwClose hDebug
        Exit Sub
    End If
    Caption = "Azmrk 调试器 - " & GetProcessName(GetProcessPath(hProcess))
    'dEvent.dwProcessId = dwProcessId
    'ListProcessThreads
    Dim Buffer(199) As Long
    Dim dwSize As Long
    Dim tbi As THREAD_BASIC_INFORMATION
    Dim nTimeout As INT64
    nTimeout.dwHigh = &HFFFFFFFF
    nTimeout.dwLow = -10& * 10000&
    '下面是防止发现调试器的部分
    '1.免IsDebuggerPresent
    Dim pbi As PROCESS_BASIC_INFORMATION
    'ZwQueryInformationProcess hProcess, ProcessBasicInformation, pbi, Len(pbi), 0
    'ZwWriteVirtualMemory hProcess, ByVal pbi.PebBaseAddress + 2, CByte(0), 1, 0
    Show
End Sub

Private Sub dBreak_Click()
    Dim Cid As CLIENT_ID, hThread As Long, fn As Long
    'Stop
    fn = GetProcAddress(GetModuleHandle("ntdll"), "DbgBreakPoint")
    RtlCreateUserThread hProcess, ByVal 0, 1, 0, 0, 0, fn, 0, hThread, Cid
    fn = GetProcAddress(GetModuleHandle("ntdll"), "RtlExitUserThread")
    Dim Ctxt As CONTEXT
    Timer1_Timer
    Ctxt.ContextFlags = CONTEXT_FULL
    ZwGetContextThread hThread, Ctxt
    WriteProcessMemory hProcess, ByVal Ctxt.Esp, fn, 4, 0 '把返回地址修改到RtlExitUserThread
    ZwResumeThread hThread, fn
End Sub

Private Sub dResume_Click()
    ZwResumeProcess hProcess
End Sub

Private Sub dSuspendImmediatly_Click()
    ZwSuspendProcess hProcess
End Sub

Private Sub Form_Unload(Cancel As Integer)
    ZwClose hDebug
    ZwClose hProcess
    Dim i As Long
    For i = 0 To tNum - 1
        Dim A As DebugThreadWindow
        Set A = Threads(i).hWindow
        ZwClose A.hThread
        DestroyWindow A.hWnd
    Next
End Sub

Public Sub DispatchEvent(ByVal pNtEvent As Long)
    Dim tbi As THREAD_BASIC_INFORMATION, Buffer(199) As Long
    Dim NtEvent As NT_DEBUG_EVENT, Status As Long
    UnionToType NtEvent, ByVal pNtEvent, Len(NtEvent)
    DbgUiConvertStateChangeStructure NtEvent, Buffer(0)
    EventNum = Buffer(0)
    Select Case EventNum
    Case EXCEPTION_DEBUG_EVENT '= 2
        UnionToType de1, Buffer(3), Len(de1)
        With de1
            Select Case .ExceptionRecord.ExceptionCode
            Case EXCEPTION_ACCESS_VIOLATION
                Bar.Panels(1).Text = "访问违规："
            Case EXCEPTION_ARRAY_BOUNDS_EXCEEDED
                Bar.Panels(1).Text = "数组出界"
            Case EXCEPTION_BREAKPOINT
            Case EXCEPTION_CONTINUABLE
            Case EXCEPTION_CONTINUE_EXECUTION
            Case EXCEPTION_CONTINUE_SEARCH
            Case EXCEPTION_DATATYPE_MISALIGNMENT
            Case EXCEPTION_EXECUTE_HANDLER
            Case EXCEPTION_FLT_DENORMAL_OPERAND
            Case EXCEPTION_FLT_DIVIDE_BY_ZERO
            Case EXCEPTION_FLT_INEXACT_RESULT
            Case EXCEPTION_FLT_INVALID_OPERATION
            Case EXCEPTION_FLT_OVERFLOW
            Case EXCEPTION_FLT_STACK_CHECK
            Case EXCEPTION_FLT_UNDERFLOW
            Case EXCEPTION_INT_DIVIDE_BY_ZERO
            Case EXCEPTION_INT_OVERFLOW
            Case EXCEPTION_IN_PAGE_ERROR
            Case EXCEPTION_NONCONTINUABLE
            Case EXCEPTION_PRIV_INSTRUCTION
            Case EXCEPTION_SINGLE_STEP
            End Select
        End With
        'Stop
        With GetThreadWindow(Buffer(2))
            .Command2.Enabled = True
            .Label1.Caption = "异常 " & Hex(de1.ExceptionRecord.ExceptionCode)
            .ChangeContext
            .Command1_Click
        End With
        EventCid.UniqueProcess = Buffer(1)
        EventCid.UniqueThread = Buffer(2)
        Status = 1
        Exit Sub
    Case CREATE_THREAD_DEBUG_EVENT '= 3
        UnionToType de2, Buffer(3), Len(de2)
        With de2
            ReDim Preserve Threads(tNum)
            Threads(tNum).dwThreadHandle = .hThread
            Status = ZwQueryInformationThread(Threads(tNum).dwThreadHandle, ThreadBasicInformation, tbi, Len(tbi), 0)
            Threads(tNum).dwThreadId = tbi.ClientId.UniqueThread
            Set Threads(tNum).hWindow = New DebugThreadWindow
            Threads(tNum).hWindow.SetThread PID, tbi.ClientId.UniqueThread, hProcess, .hThread, Me
            Threads(tNum).hWindow.Show
            SetParent Threads(tNum).hWindow.hWnd, Me.hWnd
            tNum = tNum + 1
        End With
    Case CREATE_PROCESS_DEBUG_EVENT '= 4
        UnionToType de3, Buffer(3), Len(de3)
        With de3
            ZwClose .hProcess
            ZwQueryInformationThread .hThread, ThreadBasicInformation, tbi, Len(tbi), 0
            ZwClose .hThread
            MainTID = tbi.ClientId.UniqueThread
            ZwClose .hFile
        End With
    Case EXIT_THREAD_DEBUG_EVENT '= 5
        UnionToType de4, Buffer(3), Len(de4)
        Dim A As DebugThreadWindow
        With de4
            Debug.Print de4.dwExitCode
            Set A = GetThreadWindow(Buffer(2))
            ZwClose A.hThread
            DestroyWindow A.hWnd
        End With
    Case EXIT_PROCESS_DEBUG_EVENT '= 6
        UnionToType de5, Buffer(3), Len(de5)
        MsgBox "进程已经终止，退出代码：" & de5.dwExitCode, vbInformation
        ZwClose hProcess
        Exit Sub
    Case LOAD_DLL_DEBUG_EVENT '= 7
        UnionToType de6, Buffer(3), Len(de6)
        With de6
            Dim sName As String
            sName = Space(260)
            sName = left(sName, GetModuleFileNameEx(hProcess, de6.lpBaseOfDll, sName, 260))
            Debug.Print sName
        End With
    Case UNLOAD_DLL_DEBUG_EVENT '= 8
        UnionToType de7, Buffer(3), Len(de7)
        With de7
            ReDim Dlls(dNum)
            Dlls(dNum).Modinfo.lpBaseOfDll = .lpBaseOfDll
        End With
    Case OUTPUT_DEBUG_STRING_EVENT '= 9
        UnionToType de8, Buffer(3), Len(de8)
        With de8
        End With
    Case RIP_EVENT '= 10
        UnionToType de9, Buffer(3), Len(de9)
        With de9
        End With
    Case 0
    End Select
    Debug.Print ZwDebugContinue(hDebug, NtEvent.dwProcessId, DBG_CONTINUE)
End Sub

Public Sub Timer1_Timer()
    Dim nTimeout As INT64, NtEvent As NT_DEBUG_EVENT
    nTimeout.dwLow = &HFFFFFFFF: nTimeout.dwHigh = &HFFFFFFFF
    Dim Status As Long
    'Do
        Status = ZwWaitForDebugEvent(hDebug, 1, nTimeout, NtEvent)
        If Not NT_SUCCESS(Status) Then Exit Sub
        If Status = 258 Then Exit Sub
        DispatchEvent VarPtr(NtEvent)
    'Loop
End Sub

Public Function ExprToPtr(ByVal szExp As String) As Long
    Dim i As Long, j As Long, mHandle As Long
    For i = 0 To mNum - 1
        With Modules(i)
            Dim S As String
            S = GetProcessName(.ModuleName)
            S = left(S, InStr(S, "."))
            If left(szExp, Len(S)) = S Then
                szExp = Mid(szExp, Len(S) + 1)
                mHandle = .ModuleHandle
                For j = 0 To UBound(.Procs)
                    With .Procs(j)
                        If .FunName = szExp Then
                            ExprToPtr = mHandle + .ProcOffset
                            Exit Function
                        End If
                    End With
                Next
            End If
        End With
    Next
    If IsNumeric("&H" & szExp) Then
        ExprToPtr = Val("&H" & szExp)
    End If
End Function

Public Function PtrToExpr(ByVal lPtr As Long) As String
    Dim i As Long, j As Long, mHandle As Long
    On Error Resume Next
    For i = 0 To mNum - 1
        With Modules(i)
            Dim S As String
            S = GetProcessName(.ModuleName)
            S = left(S, InStr(S, "."))
            mHandle = .ModuleHandle
            j = -1
            j = UBound(.Procs)
            If j = -1 Then GoTo NextModule
            For j = 0 To UBound(.Procs)
                With .Procs(j)
                    If lPtr = mHandle + .ProcOffset Then
                        PtrToExpr = S & .FunName
                        Exit Function
                    End If
                End With
            Next
        End With
NextModule:
    Next
    i = 0
    ZwReadVirtualMemory hProcess, ByVal lPtr, i, 2, 0
    If i = &H25FF Then 'JMP DWORD PTR
        ZwReadVirtualMemory hProcess, ByVal lPtr + 2, i, 4, 0
        PtrToExpr = "JMP.&" & PtrToExpr(i)
        Exit Function
    End If
    PtrToExpr = Space(256)
    ZwReadVirtualMemory hProcess, ByVal lPtr, ByVal PtrToExpr, 256, 0
    If CheckIsAscii(PtrToExpr) > 10 Then
        PtrToExpr = "ASCII>" & left(PtrToExpr, CheckIsAscii(PtrToExpr))
        Exit Function
    End If
    Dim Buffer(255) As Byte
    ZwReadVirtualMemory hProcess, ByVal lPtr, Buffer(0), 256, 0
    If CheckIsAscii(CStr(Buffer)) > 10 Then
        PtrToExpr = "UNICODE>" & left(CStr(Buffer), CheckIsAscii(CStr(Buffer)))
    End If
    PtrToExpr = ""
End Function

Private Function CheckIsAscii(ByVal nAsciiString As String) As Long
    Dim i As Long
    For i = 1 To Len(nAsciiString)
        Dim j As Integer
        j = Asc(Mid(nAsciiString, i, 1))
        If j < 33 Or j > 127 Then
            CheckIsAscii = i
            Exit Function
        End If
    Next
End Function

Public Function GetThreadWindow(ByVal dwTid As Long) As DebugThreadWindow
    Dim i As Long
    For i = 0 To tNum - 1
        With Threads(i)
            If .hWindow.dwTid = dwTid Then
                Set GetThreadWindow = .hWindow
                Exit Function
            End If
        End With
    Next
    ReDim Preserve Threads(tNum)
    With Threads(tNum)
        Set .hWindow = New DebugThreadWindow
        SetParent Threads(tNum).hWindow.hWnd, Me.hWnd
        .dwThreadId = dwTid
        .dwThreadHandle = FxNormalOpenThread(THREAD_ALL_ACCESS, dwTid)
        .hWindow.SetThread PID, dwTid, hProcess, .dwThreadHandle, Me
        Set GetThreadWindow = .hWindow
        .hWindow.Show
    End With
    tNum = tNum + 1
End Function

Public Sub CheckContinue()
    If Status <> 1 Then Exit Sub
    Status = 0 'Running
    If EventNum = EXCEPTION_DEBUG_EVENT Then
        If de1.ExceptionRecord.ExceptionCode = EXCEPTION_SINGLE_STEP Then
            Debug.Print "Continue Exception Single Step"
            Call ZwDebugContinue(hDebug, EventCid, DBG_CONTINUE)
            Exit Sub
        ElseIf de1.ExceptionRecord.ExceptionCode = EXCEPTION_BREAKPOINT Then
            Debug.Print "Continue Exception Break Point"
            Call ZwDebugContinue(hDebug, EventCid, DBG_CONTINUE)
            Exit Sub
        End If
    End If
    Debug.Print ZwDebugContinue(hDebug, EventCid, DBG_EXCEPTION_NOT_HANDLED)
End Sub

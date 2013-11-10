Attribute VB_Name = "Thread"
Option Explicit
Public Declare Function Thread32First Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef lpte As THREADENTRY32) As Boolean
Public Declare Function Thread32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, ByRef lpte As THREADENTRY32) As Boolean
Public Declare Function OpenThread Lib "kernel32.dll" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Boolean, ByVal dwThreadId As Long) As Long
Public Declare Function SuspendThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function ResumeThread Lib "kernel32.dll" (ByVal hThread As Long) As Long
Public Declare Function TerminateThread Lib "kernel32.dll" (ByVal hThread As Long, ByVal dwExitCode As Long) As Boolean
Public Declare Function CreateThread Lib "kernel32" (lpThreadAttributes As Long, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Declare Function Err_CreateRemoteThread Lib "kernel32" Alias "CreateRemoteThread" (ByVal hProcess As Long, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Any, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Public Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long
Public Declare Function PostThreadMessage Lib "user32" Alias "PostThreadMessageA" (ByVal idThread As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function ZwQueryInformationThread Lib "NTDLL.DLL" (ByVal hThread As Long, ByVal ThreadInformationClass As THREADINFOCLASS, ByRef ThreadInformation As Any, ByVal ThreadInformationLength As Long, ByRef ReturnLength As Long) As Long
Public Declare Function ZwGetContextThread Lib "NTDLL.DLL" (ByVal hThread As Long, ByRef pContext As CONTEXT) As Long
Public Declare Function ZwSetContextThread Lib "NTDLL.DLL" (ByVal hThread As Long, ByRef pContext As CONTEXT) As Long
Public Declare Function ZwOpenThread Lib "NTDLL.DLL" (ByRef ThreadHandle As Long, ByVal AccessMask As Long, ByRef ObjectAttributes As OBJECT_ATTRIBUTES, ByRef ClientId As CLIENT_ID) As Long
Public Declare Function ZwTerminateThread Lib "NTDLL.DLL" (ByVal ThreadHandle As Long, ByVal ExitStatus As Long) As Long
Public Declare Function RtlCreateUserThread Lib "NTDLL.DLL" (ByVal hProcess As Long, ByRef ThreadSecurityDescriptor As Any, ByVal CreateSuspended As Long, ByVal ZeroBits As Long, ByVal MaximumStackSize As Long, ByVal CommittedStackSize As Long, ByVal StartAddress As Long, ByVal Parameter As Long, ByRef hThread As Long, ByRef ClientId As CLIENT_ID) As Long


Public Const THREAD_TERMINATE = &H1
Public Const THREAD_SUSPEND_RESUME = &H2
Public Const THREAD_GET_CONTEXT = &H8
Public Const THREAD_SET_CONTEXT = &H10
Public Const THREAD_SET_INFORMATION = &H20
Public Const THREAD_QUERY_INFORMATION = &H40
Public Const THREAD_SET_THREAD_TOKEN = &H80
Public Const THREAD_IMPERSONATE = &H100
Public Const THREAD_DIRECT_IMPERSONATION = &H200
Public Const THREAD_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H3FF)

Public Const CONTEXT_ALPHA = &H20000
Public Const CONTEXT_CONTROL = (CONTEXT_ALPHA Or &H1)
Public Const CONTEXT_FLOATING_POINT = (CONTEXT_ALPHA Or &H2)
Public Const CONTEXT_INTEGER = (CONTEXT_ALPHA Or &H4)

Public Const MAXIMUM_SUPPORTED_EXTENSION = 512
Public Const SIZE_OF_80387_REGISTERS = 80


Public Enum THREADINFOCLASS
    ThreadBasicInformation = 0
    ThreadTimes = 1
    ThreadPriority = 2
    ThreadBasePriority = 3
    ThreadAffinityMask = 4
    ThreadImpersonationToken = 5
    ThreadDescriptorTableEntry = 6
    ThreadEnableAlignmentFaultFixup = 7
    ThreadEventPair = 8
    ThreadQuerySetWin32StartAddress = 9
    ThreadZeroTlsCell = 10
    ThreadPerformanceCount = 11
    ThreadAmILastThread = 12
    ThreadIdealProcessor = 13
    ThreadPriorityBoost = 14
    ThreadSetTlsArrayAddress = 15
    ThreadIsIoPending = 16
    ThreadHideFromDebugger = 17
End Enum


Public Type FLOATING_SAVE_AREA
    ControlWord As Long
    StatusWord As Long
    TagWord As Long
    ErrorOffset As Long
    ErrorSelector As Long
    DataOffset As Long
    DataSelector As Long
    RegisterArea(SIZE_OF_80387_REGISTERS) As Byte
    Cr0NpxState As Long
End Type

Public Type THREAD_CONTEXT
    FltF0 As Long
    FltF1 As Long
    FltF2 As Long
    FltF3 As Long
    FltF4 As Long
    FltF5 As Long
    FltF6 As Long
    FltF7 As Long
    FltF8 As Long
    FltF9 As Long
    FltF10 As Long
    FltF11 As Long
    FltF12 As Long
    FltF13 As Long
    FltF14 As Long
    FltF15 As Long
    FltF16 As Long
    FltF17 As Long
    FltF18 As Long
    FltF19 As Long
    FltF20 As Long
    FltF21 As Long
    FltF22 As Long
    FltF23 As Long
    FltF24 As Long
    FltF25 As Long
    FltF26 As Long
    FltF27 As Long
    FltF28 As Long
    FltF29 As Long
    FltF30 As Long
    FltF31 As Long

    IntV0 As Long    '  $0: return value register, v0
    IntT0 As Long    '  $1: temporary registers, t0 - t7
    IntT1 As Long    '  $2:
    IntT2 As Long    '  $3:
    IntT3 As Long    '  $4:
    IntT4 As Long    '  $5:
    IntT5 As Long    '  $6:
    IntT6 As Long    '  $7:
    IntT7 As Long    '  $8:
    IntS0 As Long    '  $9: nonvolatile registers, s0 - s5
    IntS1 As Long    ' $10:
    IntS2 As Long    ' $11:
    IntS3 As Long    ' $12:
    IntS4 As Long    ' $13:
    IntS5 As Long    ' $14:
    IntFp As Long    ' $15: frame pointer register, fp/s6
    IntA0 As Long    ' $16: argument registers, a0 - a5
    IntA1 As Long    ' $17:
    IntA2 As Long    ' $18:
    IntA3 As Long    ' $19:
    IntA4 As Long    ' $20:
    IntA5 As Long    ' $21:
    IntT8 As Long    ' $22: temporary registers, t8 - t11
    IntT9 As Long    ' $23:
    IntT10 As Long   ' $24:
    IntT11 As Long   ' $25:
    IntRa As Long    ' $26: return address register, ra
    IntT12 As Long   ' $27: temporary register, t12
    IntAt As Long    ' $28: assembler temp register, at
    IntGp As Long    ' $29: global pointer register, gp
    IntSp As Long    ' $30: stack pointer register, sp
    IntZero As Long  ' $31: zero register, zero

    Fpcr As Long     ' floating point control register
    SoftFpcr As Long ' software extension to FPCR

    Fir As Long      ' (fault instruction) continuation address
    Psr As Long          ' processor status

    ContextFlags As Long
    Fill(4) As Long      ' padding for 16-byte stack frame alignment
End Type

Public Type CONTEXT
    ContextFlags As Long

    Dr0 As Long
    Dr1 As Long
    Dr2 As Long
    Dr3 As Long
    Dr6 As Long
    Dr7 As Long

    FloatSave As FLOATING_SAVE_AREA

    SegGs As Long
    SegFs As Long
    SegEs As Long
    SegDs As Long

    Edi As Long
    Esi As Long
    Ebx As Long
    Edx As Long
    Ecx As Long
    Eax As Long

    Ebp As Long
    Eip As Long
    SegCs As Long              ' MUST BE SANITIZED
    EFlags As Long             ' MUST BE SANITIZED
    Esp As Long
    SegSs As Long

    ExtendedRegisters(MAXIMUM_SUPPORTED_EXTENSION) As Byte
End Type

Public Type THREAD_BASIC_INFORMATION
    ExitStatus As Long
    TebBaseAddress As Long
    ClientId As CLIENT_ID
    AffinityMask As Long
    Priority As Long
    BasePriority As Long
End Type

Public Type THREADENTRY32
    dwSize As Long
    cntUsage As Long
    th32ThreadID As Long
    th32OwnerProcessID As Long
    tpBasePri As Long
    tpDeltaPri As Long
    dwFlags As Long
    'th32CurrentProcessID As Long
End Type

Public Type USER_STACK
     FixedStackBase As Long
     FixedStackLimit As Long
     ExpandableStackBase As Long
     ExpandableStackLimit As Long
     ExpandableStackBottom As Long
End Type


Public Sub ListAllThreads(ByVal pid As Long)
    Dim ThreadInfo As THREADENTRY32
    Dim tbi As THREAD_BASIC_INFORMATION
    Dim cne As Long
    Dim msh As Long
    Dim mPath As String
    Dim hProcess As Long
    Dim hThread As Long
    Dim StartAddr As Long

    ThreadList.ListView1.ListItems.Clear
    
    msh = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, pid)
    ThreadInfo.dwSize = LenB(ThreadInfo)
    
    hProcess = FxNormalOpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, pid)
    mPath = GetProcessName(GetProcessPath(hProcess))
    ThreadList.Caption = "[" & (mPath) & "]中的线程"
    
    cne = Thread32First(msh, ThreadInfo)
    Do While cne
        If ThreadInfo.th32OwnerProcessID = pid Then
            hThread = FxNormalOpenThread(THREAD_ALL_ACCESS, ThreadInfo.th32ThreadID)
            ZwQueryInformationThread hThread, ThreadBasicInformation, tbi, Len(tbi), 0
            StartAddr = FxGetThreadStartAddress(hThread)
            
            'ThreadList.ListView1.ListItems.Add , ,
            ThreadList.ListView1.ListItems.Add , , CStr(ThreadInfo.th32ThreadID)
            With ThreadList.ListView1.ListItems(ThreadList.ListView1.ListItems.Count)
                .SubItems(1) = FormatHex(tbi.TebBaseAddress)
                '.SubItems(2) = ETHREAD
                .SubItems(3) = PriorityCheck(ThreadInfo.tpBasePri)
                .SubItems(4) = FormatHex(StartAddr)
                '.SubItems(5) = ThreadStatus
                .SubItems(6) = FxGetThreadModuleFileName(hProcess, hThread, StartAddr)
            End With
            'GetExitCodeThread hThread, ByVal ec
            'ListView1.ListItems(ListView1.ListItems.Count).SubItems(3) = WaitForSingleObject(hThread, 10)
            ZwClose hThread
        End If
        cne = Thread32Next(msh, ThreadInfo)
    Loop
    
    ZwClose hProcess: hProcess = 0
    ZwClose msh: msh = 0
    
    FxGetThreadEThread
    
    ThreadList.Caption = (ThreadList.Caption) & " (" & ThreadList.ListView1.ListItems.Count & ")"
End Sub

Public Sub SusResProcess(pid As Long, dType As Boolean)
    Dim ThreadInfo As THREADENTRY32
    Dim cne As Integer
    Dim msh As Long
    Dim hThread As Long
    
    msh = CreateToolhelp32Snapshot(TH32CS_SNAPTHREAD, pid)
    ThreadInfo.dwSize = LenB(ThreadInfo)
    
    cne = Thread32First(msh, ThreadInfo)
    Do While cne
        If ThreadInfo.th32OwnerProcessID = pid Then
            hThread = OpenThread(THREAD_SUSPEND_RESUME, False, ThreadInfo.th32ThreadID)
            If dType = True Then
                SuspendThread hThread
            Else
                ResumeThread hThread
            End If
        End If
        cne = Thread32Next(msh, ThreadInfo)
    Loop
    CloseHandle msh
End Sub

Public Function FxNormalOpenThread(ByVal dwDesiredAccess As Long, ByVal tid As Long) As Long
    Dim oa As OBJECT_ATTRIBUTES
    Dim Cid As CLIENT_ID
    Dim hThread As Long
    Dim st As Long
    
    oa.Length = LenB(oa)
    Cid.UniqueThread = tid

    st = ZwOpenThread(hThread, dwDesiredAccess, oa, Cid)
    If Not NT_SUCCESS(st) Then
        hThread = LzOpenThread(dwDesiredAccess, tid)
    End If
    
    FxNormalOpenThread = hThread
End Function

Public Function LzOpenThread(ByVal dwDesiredAccess As Long, ByVal ThreadID As Long) As Long
    '/**函数功能:通过复制句柄表里的句柄来“打开”线程**/
    Dim st As Long
    Dim Cid As CLIENT_ID
    Dim oa As OBJECT_ATTRIBUTES
    Dim NumOfHandle As Long
    Dim tbi As THREAD_BASIC_INFORMATION
    Dim i As Long
    Dim hProcessToDup As Long, hThreadCur As Long, hThreadToRet As Long
    
    oa.Length = Len(oa)
    '首先尝试ZwOpenThread
    Cid.UniqueThread = ThreadID
    st = ZwOpenThread(hThreadToRet, dwDesiredAccess, oa, Cid)
    If (NT_SUCCESS(st)) Then LzOpenThread = hThreadToRet: Exit Function
    st = 0
    
    Dim bytBuf() As Byte
    Dim arySize As Long: arySize = 1
    Do
        ReDim bytBuf(arySize)
        st = ZwQuerySystemInformation(SystemHandleInformation, VarPtr(bytBuf(0)), arySize, 0&)
        If (Not NT_SUCCESS(st)) Then
            If (st <> STATUS_INFO_LENGTH_MISMATCH) Then
                Erase bytBuf
                Exit Function
            End If
        Else
            Exit Do
        End If
        arySize = arySize * 2
        ReDim bytBuf(arySize)
    Loop
    
    NumOfHandle = 0
    CopyMemory VarPtr(NumOfHandle), VarPtr(bytBuf(0)), Len(NumOfHandle)
    Dim h_info() As SYSTEM_HANDLE_TABLE_ENTRY_INFO
    ReDim h_info(NumOfHandle)
    CopyMemory VarPtr(h_info(0)), VarPtr(bytBuf(0)) + Len(NumOfHandle), Len(h_info(0)) * NumOfHandle
    
    '//枚举句柄完成，下来开始测试句柄
    For i = LBound(h_info) To UBound(h_info)
        With h_info(i)
            If (.ObjectTypeIndex = OB_TYPE_PROCESS + 1) Then
                'oa.Length = Len(oa)
                'cid.UniqueProcess = .UniqueProcessId
                'st = ZwOpenProcess(hProcessToDup, PROCESS_ALL_ACCESS, oa, cid)
                hProcessToDup = FxNormalOpenProcess(PROCESS_DUP_HANDLE, .UniqueProcessId)
                If hProcessToDup <> 0 Then
                    st = ZwDuplicateObject(hProcessToDup, .HandleValue, ZwGetCurrentProcess, hThreadCur, THREAD_ALL_ACCESS, 0, DUPLICATE_SAME_ATTRIBUTES)
                    If (NT_SUCCESS(st)) Then
                        st = ZwQueryInformationThread(hThreadCur, ThreadBasicInformation, tbi, Len(tbi), 0)
                        If (NT_SUCCESS(st)) Then
                            If (tbi.ClientId.UniqueThread = ThreadID) Then
                                st = ZwDuplicateObject(hProcessToDup, .HandleValue, ZwGetCurrentProcess, hThreadToRet, dwDesiredAccess, 0, DUPLICATE_SAME_ATTRIBUTES)
                                If (NT_SUCCESS(st)) Then
                                    If hThreadToRet <> 0 Then
                                        LzOpenThread = hThreadToRet
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                    st = ZwClose(hThreadCur)
                End If
                st = ZwClose(hProcessToDup)
            End If
        End With
    Next i
    
    Erase h_info
End Function

Public Function FxGetThreadId(ByVal hThread As Long) As Long
    Dim tbi As THREAD_BASIC_INFORMATION
    Dim st As Long
    
    st = ZwQueryInformationThread(hThread, ThreadBasicInformation, tbi, Len(tbi), 0)
    If (NT_SUCCESS(st)) Then
        FxGetThreadId = tbi.ClientId.UniqueThread
    End If
End Function

Public Function FxGetThreadModuleFileName(ByVal hProcess As Long, ByVal hThread As Long, Optional ByVal StartAddr As Long = 0) As String
    Dim lPtr As Long
    Dim pbi As PROCESS_BASIC_INFORMATION
    Dim tPEB_LDR_DATA As PEB_LDR_DATA
    Dim tLDR_MODULE As LDR_MODULE
    Dim tBLDR_MODULE As LDR_MODULE
    Dim tFLDR_MODULE As LDR_MODULE
    Dim modPath As String * MAX_PATH

    '获得PEB
    ZwQueryInformationProcess hProcess, ProcessBasicInformation, pbi, Len(pbi), 0
    '获得线程的入口地址
    If Not StartAddr Then
        ZwQueryInformationThread hThread, ThreadQuerySetWin32StartAddress, ByVal VarPtr(StartAddr), Len(StartAddr), 0
    End If
    'PEB指针
    lPtr = pbi.PebBaseAddress

    '如果地址无误
    If lPtr Then
        '如果成功读取到数据
        If Not ReadProcessMemory(hProcess, ByVal lPtr + 12, lPtr, &H4, 0&) = 0 Then
            '找到链表头
            ReadProcessMemory hProcess, ByVal lPtr, ByVal VarPtr(tPEB_LDR_DATA), Len(tPEB_LDR_DATA), 0
            ReadProcessMemory hProcess, ByVal tPEB_LDR_DATA.InLoadOrderModuleList.Flink, ByVal VarPtr(tLDR_MODULE), Len(tLDR_MODULE), 0
            '继续读取数据直到DLL基址为0
            Do While tLDR_MODULE.BaseAddress <> 0
                If StartAddr > tLDR_MODULE.BaseAddress And StartAddr < tLDR_MODULE.BaseAddress + tLDR_MODULE.SizeOfImage Then
                    GetModuleFileNameEx hProcess, tLDR_MODULE.BaseAddress, modPath, MAX_PATH
                    FxGetThreadModuleFileName = modPath
                    Exit Function
                End If
                ReadProcessMemory hProcess, ByVal tLDR_MODULE.InLoadOrderModuleList.Flink, ByVal VarPtr(tLDR_MODULE), Len(tLDR_MODULE), 0
            Loop
        End If
    End If
End Function

Public Sub FxGetThreadEThread()
    '/**函数功能:填充Lsitview中的ETHREAD项**/
    
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
    
    Dim i, j As Long
    Dim nowTid As Long
    
    For i = LBound(h_info) To UBound(h_info) / 4
        With h_info(i)
            If .ObjectTypeIndex = OB_TYPE_PROCESS + 1 Then
                nowTid = PsGetTidByEThread(.pObject)
                For j = 1 To ThreadList.ListView1.ListItems.Count
                    If ThreadList.ListView1.ListItems(j).Text = nowTid And ThreadList.ListView1.ListItems(j).SubItems(2) = "" Then
                        ThreadList.ListView1.ListItems(j).SubItems(2) = FormatHex(.pObject)
                        Exit For
                    End If
                Next j
            End If
        End With
    Next i
    
    Erase h_info
End Sub

Public Function FxGetThreadStartAddress(ByVal hThread As Long) As Long
    Dim StartAddr As Long
    
    ZwQueryInformationThread hThread, ThreadQuerySetWin32StartAddress, ByVal VarPtr(StartAddr), Len(StartAddr), 0
    FxGetThreadStartAddress = StartAddr
End Function

Public Function FxDestroyThreadContext(ByVal hThread As Long) As Long
    Dim old_context As CONTEXT
    Dim new_context As CONTEXT
    Dim errNum As Long
    
    old_context.ContextFlags = CONTEXT_CONTROL
    If NT_SUCCESS(ZwGetContextThread(hThread, old_context)) Then
        'With old_context
            'MsgBox .Ebp & .Esp & .ContextFlags
            'MsgBox "GetContextThread Succeed!"
            old_context.Ebp = 100
            If NT_SUCCESS(ZwSetContextThread(hThread, old_context)) Then
                'MsgBox "SetContextThread Succeed!"
                If NT_SUCCESS(ZwGetContextThread(hThread, new_context)) Then MsgBox new_context.Ebp
            End If
            errNum = GetLastError
            'MsgBox errNum
        'End With
    End If
End Function

Public Function PsGetThreadStartAddressByEThread(ByVal ETHREAD As Long) As Long
    ReadKernelMemory ETHREAD + &H228, PsGetThreadStartAddressByEThread, 4, 0
End Function

Public Function PsGetTidByEThread(ByVal ETHREAD As Long) As Long
    '/**函数功能:由ETHREAD获取TID**/

    Dim mc As MEMORY_CHUNKS
    Dim retl As CLIENT_ID
    Dim Cid As CLIENT_ID
    
    With mc
        .Address = ETHREAD + &H1EC
        .Length = Len(Cid)
        .pData = VarPtr(Cid)
    End With
    
    Dim st As Long
    st = ZwSystemDebugControl(SysDbgCopyMemoryChunks_0, VarPtr(mc), Len(mc), 0&, 0&, VarPtr(retl))
    PsGetTidByEThread = Cid.UniqueThread
    If (Not NT_SUCCESS(st)) Then PsGetTidByEThread = 0
End Function

Public Function FxCreateRemoteThread(ByVal pid As Long, ByVal lpStartAddress As Long, ByVal lpParameter As Long)
'OsCreateRemoteThread(DWORD dwpid,StartAddress:pointer)
'{
    Dim hProcess As Long
    hProcess = OpenProcess(PROCESS_ALL_ACCESS, 0, pid) '//dwpid就是某些个系统进程ID
    
    Dim stack As USER_STACK
    
    'DWORD ret;
    Dim ret As Long
    
    'ULONG n = 1024*1024;//1MB
    Dim n As Long
    n = 1024 * 1024
    
    Dim rAddress As Long
    
    ret = ZwAllocateVirtualMemory(hProcess, rAddress, 0, n, MEM_RESERVE, PAGE_READWRITE)
    
    'stack.ExpandableStackBase = PCHAR(stack.ExpandableStackBottom) +1024*1024;
    'stack.ExpandableStackLimit = PCHAR(stack.ExpandableStackBase) - 4096;
    'n = 4096 + PAGE_SIZE;
    
    'PVOID p = PCHAR(stack.ExpandableStackBase) - n;
    'ret=ZwAllocateVirtualMemory(hProcess, &p, 0, &n, MEM_COMMIT, PAGE_READWRITE);
    
    'ULONG x; n = PAGE_SIZE;
    'ret=ZwProtectVirtualMemory(hProcess, &p, &n, PAGE_READWRITE | PAGE_GUARD, &x);
    
    'CONTEXT context = {CONTEXT_FULL};
    
    'ret=ZwGetContextThread(GetCurrentThread(),&context);
    
    'context.Esp = ULONG(stack.ExpandableStackBase) - 2048;
    'context.Eip = ULONG(startaddress);
    
    'CLIENT_ID cid;
    
    'ret=ZwCreateThread(&hThread, THREAD_ALL_ACCESS, 0, hProcess, &cid, &context, &stack, TRUE);
    'if(ret) MessageBox(0,"ZwCreateThread","",0);
    
    'ret=ZwGetContextThread(hThread,&context);
    'ret=RtlNtStatusToDosError(ret);
    'ret=ZwResumeThread(hThread,0);
    'ret=RtlNtStatusToDosError(ret);
    
    'CloseHandle(hProcess);
    'CloseHandle(hThread);
    
    'return true;
'}
End Function

Public Function ChCreateRemoteThread(ByVal hProcess As Long, ByVal StartAddress As Long, ByVal Parameter As Long, ByRef Cid As CLIENT_ID) As Long
    Dim hThread As Long
    Dim ntStatus As Long
    
    ntStatus = RtlCreateUserThread(hProcess, ByVal 0&, 0, 0, 0, 0, StartAddress, Parameter, hThread, Cid)
    
    WaitForSingleObject hThread, INFINITE
    
    ChCreateRemoteThread = hThread
End Function

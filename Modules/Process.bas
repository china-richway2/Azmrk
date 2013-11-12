Attribute VB_Name = "Process"
Option Explicit
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetProcessImageFileName Lib "psapi.dll" Alias "GetProcessImageFileNameA" (ByVal hProcess As Long, ByVal lpImageFileName As String, ByVal nSize As Long) As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function EndTask Lib "user32.dll" (ByVal hWnd As Long, ByVal fShutDown As Long, ByVal fForce As Long) As Long
Public Declare Function WinStationTerminateProcess Lib "winsta.dll" (ByVal hServer As Long, ByVal ProcessID As Long, ByVal ExitCode As Long) As Long
Public Declare Function ZwQueryInformationProcess Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, ByVal InformationClass As PROCESSINFOCLASS, ByRef ProcessInformation As Any, ByVal ProcessInformationLength As Long, ByRef ReturnLenght As Long) As Long
Public Declare Function ZwSetInformationProcess Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, ByVal InformationClass As PROCESSINFOCLASS, ByRef ProcessInformation As Any, ByVal ProcessInformationLength As Long) As Long
Public Declare Function ZwOpenProcess Lib "NTDLL.DLL" (ByRef ProcessHandle As Long, ByVal AccessMask As Long, ByRef ObjectAttributes As OBJECT_ATTRIBUTES, ByRef ClientId As CLIENT_ID) As Long
Public Declare Function ZwTerminateProcess Lib "NTDLL.DLL" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function ZwSuspendProcess Lib "NTDLL.DLL" (ByVal hProcess As Long) As Long
Public Declare Function ZwResumeProcess Lib "NTDLL.DLL" (ByVal hProcess As Long) As Long
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)
Public Declare Function CreateProcess Lib "kernel32" Alias "CreateProcessA" (ByVal lpApplicationName As String, ByVal lpCommandLine As String, lpProcessAttributes As SECURITY_ATTRIBUTES, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal bInheritHandles As Long, ByVal dwCreationFlags As Long, lpEnvironment As Any, ByVal lpCurrentDriectory As String, lpStartupInfo As STARTUPINFO, lpProcessInformation As PROCESS_INFORMATION) As Long
Public Declare Function GetProcessTimes Lib "kernel32" (ByVal hProcess As Long, lpCreationTime As Any, lpExitTime As Any, lpKernelTime As Any, lpUserTime As Any) As Long
Public Type STARTUPINFO
        cb As Long
        lpReserved As String
        lpDesktop As String
        lpTitle As String
        dwX As Long
        dwY As Long
        dwXSize As Long
        dwYSize As Long
        dwXCountChars As Long
        dwYCountChars As Long
        dwFillAttribute As Long
        dwFlags As Long
        wShowWindow As Integer
        cbReserved2 As Integer
        lpReserved2 As Long
        hStdInput As Long
        hStdOutput As Long
        hStdError As Long
End Type
Public Type PROCESS_INFORMATION
        hProcess As Long
        hThread As Long
        dwProcessId As Long
        dwThreadId As Long
End Type
Public Const CREATE_NEW = &H1
Public Const CREATE_ALWAYS = &H2
Public Const CREATE_SUSPENDED = &H4
Public Const CREATE_NEW_CONSOLE = &H10
Public Const CREATE_NEW_PROCESS_GROUP = &H200
Public Const CREATE_NO_WINDOW = &H8000000
'Public Const CREATE_PROCESS_DEBUG_EVENT = 3
'Public Const CREATE_THREAD_DEBUG_EVENT = 2


Public Const PROCESS_TERMINATE = &H1
Public Const PROCESS_CREATE_THREAD = &H2
Public Const PROCESS_SET_SESSIONID = &H4
Public Const PROCESS_VM_OPERATION = &H8
Public Const PROCESS_VM_READ = &H10
Public Const PROCESS_VM_WRITE = &H20
Public Const PROCESS_DUP_HANDLE = &H40
Public Const PROCESS_CREATE_PROCESS = &H80
Public Const PROCESS_SET_QUOTA = &H100
Public Const PROCESS_SET_INFORMATION = &H200
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const PROCESS_SUSPEND_RESUME = &H800
Public Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
'Public Const PROCESS_ALL_ACCESS As Long = &H1F0FFF '所有权限

Public Const SMTO_ABORTIFHUNG = &H2
Public Const IDLE_PRIORITY_CLASS = &H40 '新进程应该有非常低的优先级——只有在系统空闲的时候才能运行。基本值是4
Public Const HIGH_PRIORITY_CLASS = &H80  '新进程有非常高的优先级，它优先于大多数应用程序。基本值是13。注意尽量避免采用这个优先级
Public Const NORMAL_PRIORITY_CLASS = &H20 '标准优先级。如进程位于前台，则基本值是9；如在后台，则优先值是7

Public Const DUPLICATE_CLOSE_SOURCE = &H1              '// winnt
Public Const DUPLICATE_SAME_ACCESS = &H2                  '// winnt
Public Const DUPLICATE_SAME_ATTRIBUTES = &H4

Public Const WTS_CURRENT_SERVER_HANDLE = 0

Public Const ZwGetCurrentProcess As Long = -1 '//0xFFFFFFFF


Public Enum PROCESSINFOCLASS
      ProcessBasicInformation
      ProcessQuotaLimits
      ProcessIoCounters
      ProcessVmCounters
      ProcessTimes
      ProcessBasePriority
      ProcessRaisePriority
      ProcessDebugPort
      ProcessExceptionPort
      ProcessAccessToken
      ProcessLdtInformation
      ProcessLdtSize
      ProcessDefaultHardErrorMode
      ProcessIoPortHandlers         '// Note: this is kernel mode only
      ProcessPooledUsageAndLimits
      ProcessWorkingSetWatch
      ProcessUserModeIOPL
      ProcessEnableAlignmentFaultFixup
      ProcessPriorityClass
      ProcessWx86Information
      ProcessHandleCount
      ProcessAffinityMask
      ProcessPriorityBoost
      ProcessDeviceMap
      ProcessSessionInformation
      ProcessForegroundInformation
      ProcessWow64Information
      ProcessImageFileName
      ProcessLUIDDeviceMapsEnabled
      ProcessBreakOnTermination
      ProcessDebugObjectHandle
      ProcessDebugFlags
      ProcessHandleTracing
      ProcessIoPriority
      ProcessExecuteFlags
      ProcessResourceManagement
      ProcessCookie
      ProcessImageInformation
      MaxProcessInfoClass           '// MaxProcessInfoClass should always be the last enum
End Enum


Public Type PROCESS_BASIC_INFORMATION
    ExitStatus As Long ' 接收进程终止状态
    PebBaseAddress As Long '接收进程环境块地址
    AffinityMask As Long ' 接收进程关联掩码；每个标志位表示一个处理器
    BasePriority As Long ' 接收进程的优先级类
    UniqueProcessId As Long ' 接收进程ID
    InheritedFromUniqueProcessId As Long '接收父进程ID
End Type

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 1024
End Type

Public Type SYSTEM_PROCESSORS
    NextEntryDelta As Long
    ThreadCount As Long
    Reserved(1 To 6) As Long
    CreateTime As FILETIME
    UserTime As FILETIME
    KernelTime As FILETIME
    ProcessName As UNICODE_STRING
    BasePriority As Long
    ProcessID As Long
    InheritedFromProcessId As Long
    HandleCount As Long
    Reserved2(1 To 2) As Long
    PrivatePageCount As Long
    VmCounters As VM_COUNTERS
    IoCounters As IO_COUNTERS
    '从这里开始第一个SYSTEM_THREAD_INFORMATION
End Type

Public Type SYSTEM_THREAD_INFORMATION
    KernelTime As FILETIME
    UserTime As FILETIME
    CreateTime As FILETIME
    WaitTime As Long
    StartAddress As Long
    ClientId As CLIENT_ID
    Priority As Long
    BasePriority As Long
    ContextSwitchCount As Long
    State As THREAD_STATE
    WaitReason As Long
End Type

Public Type AzmrkThread
    Basic As SYSTEM_THREAD_INFORMATION
End Type

Public Type AzmrkProcess
    Basic As PROCESS_BASIC_INFORMATION
    EPROCESS As Long
    ExePath As String
    ImageName As String
    CmdLine As String
    CreateTime As FILETIME
    ExitTime As FILETIME
    KernelTime As FILETIME
    UserTime As FILETIME
    LastKernelTime As FILETIME
    LastUserTime As FILETIME
    ListViewIndex As Long
    State As Boolean
    Handle As Long
    FirstUpdate As Boolean
    ThreadCount As Long
    HandleCount As Long
    VmCounters As VM_COUNTERS
    IoCounters As IO_COUNTERS
    Threads() As AzmrkThread
End Type

Public Type PROCESS_ENVIRONMENT_BLOCK
    InheritedAddressSpace As Byte
    ReadImageFileExecOptions As Byte
    BeingDebugged As Byte
    SparePool As Byte
    Mutant As Long
    ImageBaseAddress As Long
    Ldr As Long 'PPEB_LDR_DATA
    ProcessParameters As Long 'PRTL_USER_PROCESS_PARAMETERS
    SubSystemData As Long
    ProcessHeap As Long
    FastPebLock As Long 'PRTL_CRITICAL_SECTION
End Type

Public Type TEB_LDR_DATA
    SehListPtr As Long '+0x00
    StackTop As Long '+0x04
    StackBottom As Long '+0x08
    SubSystemTib As Long '+0x0C
    FiberData As Long '+0x10
    ArbitraryUserPointer As Long '+0x14
    FsImageAddr As Long '+0x18
    PID As Long '+0x20
    TID As Long '+0x24
    ActiveRpcInfo As Long '+0x28
    ThreadLocalSaveAreaPtr As Long '+0x2C
    Peb As Long '+0x30
    LastErr As Long '+0x34
End Type

Public Type PEB_LDR_DATA
    Length                          As Long
    initialized                     As Long
    SsHandle                        As Long
    InLoadOrderModuleList           As LIST_ENTRY
    InMemoryOrderModuleList         As LIST_ENTRY
    InInitializationOrderModuleList As LIST_ENTRY
End Type

Public Enum EnumProcessMethod
    MethodSnapshot
    MethodSession
    MethodEnumProcesses
    MethodTest
    MethodHandleList
    MethodQuery
End Enum

Public nsItem As Long, MainState As Boolean, Processes() As AzmrkProcess

Public Function FileTime2String(lpFiletime As FILETIME) As String
    Dim l2(1 To 8) As Byte, l3(1 To 8) As Byte
    CopyMemory VarPtr(l2(1)), VarPtr(lpFiletime), 8
    Dim i As Integer, is0 As Boolean
    Do
        Dim T As Long
        T = 0: is0 = True
        For i = 8 To 1 Step -1
            T = T * 256 + l2(i)
            If T <> 0 Then is0 = False
            l2(i) = T \ 10
            T = T Mod 10
        Next
        If is0 = True Then Exit Function
        FileTime2String = T & FileTime2String
    Loop
End Function

Public Sub FileTimeSub(lp1 As FILETIME, lp2 As FILETIME)
    Dim p1(1 To 8) As Byte, P2(1 To 8) As Byte
    CopyMemory VarPtr(p1(1)), VarPtr(lp1), 8
    CopyMemory VarPtr(P2(1)), VarPtr(lp2), 8
    Dim i As Byte
    For i = 1 To 8
        If p1(i) < P2(i) Then
            p1(i - 1) = p1(i - 1) - 1
            p1(i) = CInt(p1(i)) + 256 - P2(i)
        Else
            p1(i) = p1(i) - P2(i)
        End If
    Next
    CopyMemory VarPtr(lp1), VarPtr(p1(1)), 8
End Sub

Public Sub FileTimeAdd(lp1 As FILETIME, lp2 As FILETIME)
    Dim p1(1 To 8) As Byte, P2(1 To 8) As Byte, n As Integer
    CopyMemory VarPtr(p1(1)), VarPtr(lp1), 8
    CopyMemory VarPtr(P2(1)), VarPtr(lp2), 8
    Dim i As Byte
    For i = 1 To 8
        n = n + p1(i)
        n = n + P2(i)
        p1(i) = n And 255
        n = n \ 256
    Next
    CopyMemory VarPtr(lp1), VarPtr(p1(1)), 8
End Sub

Public Function NewProcess(ByVal dwPid As Long) As Long
    Dim i As Long
    On Error GoTo ErrHand
    For i = 0 To UBound(Processes)
        If Processes(i).Basic.UniqueProcessId = dwPid Then
            Processes(i).State = MainState
            NewProcess = i
            Exit Function
        End If
    Next
    For i = 0 To UBound(Processes)
        If Processes(i).ListViewIndex = 0 Then
            NewProcess = i
            With Processes(NewProcess)
                .State = MainState
                .ListViewIndex = Menu.ListView2.ListItems.Add.Index
                .Basic.UniqueProcessId = dwPid
            End With
            Exit Function
        End If
    Next
    ReDim Preserve Processes(i)
NewP:
    With Processes(i)
        .Basic.UniqueProcessId = dwPid
        NewProcess = i
        .ListViewIndex = Menu.ListView2.ListItems.Add.Index
        .State = MainState
    End With
    Exit Function
ErrHand:
    ReDim Processes(0)
    GoTo NewP
End Function

Public Sub ProcessAntiFill(ByVal nItem As Long)
    Dim st As Long
    With Processes(nItem)
        .LastKernelTime = .KernelTime
        .LastUserTime = .UserTime
        ZwQueryInformationProcess .Handle, ProcessVmCounters, .VmCounters, Len(.VmCounters), 0
        st = ZwQueryInformationProcess(.Handle, ProcessIoCounters, .IoCounters, Len(.IoCounters), 0)
        .LastKernelTime = .KernelTime
        .LastUserTime = .UserTime
        ZwQueryInformationProcess .Handle, ProcessTimes, .CreateTime, 32, 0
        ZwQueryInformationProcess .Handle, ProcessHandleCount, .HandleCount, 4, 0
    End With
End Sub

Public Sub ProcessFillByEProcess(ByVal nItem As Long)
    With Processes(nItem)
        If .Basic.BasePriority = 0 Then
            ReadKernelMemory .EPROCESS + &H62, VarPtr(.Basic.BasePriority), 1, 0
        End If
        If .VmCounters.VirtualSize = 0 Then
            ReadKernelMemory .EPROCESS + &H1B8, VarPtr(.VmCounters), 48, 0
        End If
        If .ImageName = "" Then
            .ImageName = FxGetProcessName(.EPROCESS)
        End If
    End With
End Sub

Public Sub FillProcessByHandle(ByVal nItem As Long)
    Dim nHandle As Long
    With Processes(nItem)
        If .Handle = 0 Then
            Debug.Assert .Basic.UniqueProcessId <> 948
            nHandle = FxNormalOpenProcess(PROCESS_ALL_ACCESS, .Basic.UniqueProcessId)
            If nHandle = 0 Then
                .Handle = FxNormalOpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, .Basic.UniqueProcessId)
                If .Handle = 0 Then Exit Sub
            Else
                .Handle = nHandle
            End If
        End If
        .CmdLine = GetProcessCommandLine(.Handle)
        .ExePath = GetProcessPath(.Handle)
        .ImageName = GetProcessName(.ExePath)
        ZwQueryInformationProcess .Handle, ProcessBasicInformation, .Basic, Len(.Basic), 0
    End With
End Sub


Public Sub FxListProcessBySession()
    Dim dwReturnLen As Long
    Dim etStart As Long
    Dim etLast As Long
    Dim etNow As Long
    Dim etNext As Long
    Dim tListProcess As LIST_ENTRY
    Dim tBListProcess As LIST_ENTRY
    Dim tFListProcess As LIST_ENTRY
    Dim nItem As Long
    Dim pbi As PROCESS_BASIC_INFORMATION
    Dim PID As Long
    Dim EPROCESS As Long
    Dim pPath As String
    Dim pName As String
    Dim loopMax As Long

    etStart = FxAddSystemProcess
    etNext = etStart
    loopMax = 0
    Do
        PID = 0
        '获取PID
        ReadKernelMemory etNext + &H84, VarPtr(PID), Len(PID), dwReturnLen
        '如果PID无误就添加项目
        If PID > 0 And PID < 65535 Then
            nItem = NewProcess(PID)
            With Processes(nItem)
                FillProcessByHandle nItem
                .EPROCESS = etNext
                FillProcessByHandle nItem
            End With
        End If
        '获取本节的LIST_ENTRY
        ReadKernelMemory etNext + &HB4, VarPtr(tListProcess), Len(tListProcess), dwReturnLen
        'MsgBox CStr(tListProcess.Blink) & "," & CStr(tListProcess.Flink)
        '本节
        etNow = etNext
        '上一个结
        etLast = tListProcess.Flink - &HB4
        '下一个结
        etNext = tListProcess.Blink - &HB4

        loopMax = loopMax + 1
    Loop While loopMax < 65535 And (etNext <> etStart)
    
End Sub

Public Sub mpNew_Click()
    Dim ProcessInfo As PROCESSENTRY32
    Dim pbi As PROCESS_BASIC_INFORMATION
    Dim pc As Long
    Dim pm As Long
    Dim nItem As Long
    Dim i As Long
    Dim hInfo As SYSTEM_HANDLE_TABLE_ENTRY_INFO
    
    '开始遍历
    pc = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    RdQueryHandleInformation pc, hInfo, -1
    Dim l As SYSTEM_HANDLE_TABLE_ENTRY_INFO
    ProcessInfo.dwSize = Len(ProcessInfo)

    pm = Process32First(pc, ProcessInfo)
    Do While pm
        nItem = NewProcess(ProcessInfo.th32ProcessID)
        With Processes(nItem)
            .Basic.BasePriority = ProcessInfo.pcPriClassBase
            .ImageName = left(ProcessInfo.szExeFile, InStr(ProcessInfo.szExeFile, Chr(0)))
            FillProcessByHandle nItem
        End With
        
        pm = Process32Next(pc, ProcessInfo)
    Loop
    
    ZwClose pc
    FillEProcess
End Sub

Public Sub ListProcessByQuery()
    Dim buffer() As Byte
    Dim nLength As Long
S:
    ZwQuerySystemInformation SystemProcessInformation, 0, 0, nLength
    If nLength = 0 Then Exit Sub
    ReDim buffer(1 To nLength)
    If Not NT_SUCCESS(ZwQuerySystemInformation(SystemProcessInformation, VarPtr(buffer(1)), nLength, nLength)) Then
        GoTo S
    End If
    Dim inf As SYSTEM_PROCESSORS, E As Long, n As Long, StrNameBuffer() As Byte
    E = 1 '
    Do
        CopyMemory VarPtr(inf), VarPtr(buffer(E)), Len(inf)
        E = E + inf.NextEntryDelta
        n = NewProcess(inf.ProcessID)
        With Processes(n)
            If inf.ProcessName.Length <> 0 Then
                ReDim StrNameBuffer(inf.ProcessName.Length - 1)
                CopyMemory VarPtr(StrNameBuffer(0)), inf.ProcessName.buffer, inf.ProcessName.Length
                .ImageName = StrNameBuffer
            End If
            .ThreadCount = inf.ThreadCount
            .VmCounters = inf.VmCounters
            .IoCounters = inf.IoCounters
            .CreateTime = inf.CreateTime
            .LastKernelTime = .KernelTime
            .LastUserTime = .UserTime
            .KernelTime = inf.KernelTime
            .UserTime = inf.UserTime
            .Basic.BasePriority = inf.BasePriority
            .Basic.InheritedFromUniqueProcessId = inf.InheritedFromProcessId
            CopyMemory VarPtr(.VmCounters), VarPtr(inf.VmCounters), LenB(inf.VmCounters) + LenB(inf.IoCounters)
            Call FillProcessByHandle(n)
            ProcessAntiFill n
        End With
    Loop Until inf.NextEntryDelta = 0
    FillEProcess
End Sub

Public Sub ListProcessByWmi()
    Dim objSWbemLocator As New SWbemLocator
    Dim objSWbemServices As SWbemServices
    Dim objSWbemObjectSet As SWbemObjectSet
    Dim objSWbemObject As SWbemObject
    Dim i As Long
    Dim pIndex As Long
    
    pIndex = 1
    
    '清空表
    pIndex = FxGetListviewNowLine(Menu.ListView2)
    
    Menu.ListView2.Tag = 2
    
    Menu.ListView2.ListItems.Clear '清空ListView
    Set objSWbemServices = objSWbemLocator.ConnectServer()  '连接到本机的WMI，返回一个对 SWbemServices 对象的引用
    Set objSWbemObjectSet = objSWbemServices.InstancesOf("Win32_Process")   '返回Win32_Process类名标识的所有实例
    i = 0
    For Each objSWbemObject In objSWbemObjectSet  '枚举每一个Win32_Process的实例
        Menu.ListView2.ListItems.Add , "a" & i, objSWbemObject.Name '将进程ID添加到ListView1第一列
        With Menu.ListView2.ListItems("a" & i)
            .SubItems(1) = objSWbemObject.Handle '将进程名添加到ListView1第二列
            .SubItems(2) = FxGetProcessMemoryInformation(objSWbemObject.Handle)
        End With
        If Not IsNull(objSWbemObject.ExecutablePath) Then Menu.ListView2.ListItems("a" & i).SubItems(3) = objSWbemObject.ExecutablePath '将进程路径添加到ListView1第三列
        i = i + 1
    Next
    Set objSWbemObjectSet = Nothing
End Sub

Public Sub ListProcessHf()
    '通过PSAPI.DLL里的EnumProcesses来遍历进程,效果同Toolhelp32系列,保留,不使用
    Dim PID(1024) As Long
    Dim prCount As Long
    Dim i As Integer
    Dim pIndex As Integer
    
    pIndex = 1
    
    If Menu.ListView2.ListItems.count > 0 And Menu.ListView2.Tag = 1 Then
        pIndex = Menu.ListView2.SelectedItem.Index
    End If
    If Menu.ListView2.Sorted = True Then Menu.ListView2.Sorted = False
    
    Menu.ListView2.Tag = 1

    Menu.ListView2.ListItems.Clear
    EnumProcesses PID(0), 1024, prCount
    For i = 0 To prCount / 4 - 1
        'ListView2.ListItems.Add , , pID(i)
        
    Next i
End Sub

Public Function FxAddSystemProcess() As Long
    Dim EPROCESS As Long
    Dim Ret() As Long
    'Dim hModule As Long
    Dim PsInitialSystemProcess As Long
    Dim lngSList As Long
    Dim lngAList As Long
    Dim etStart As Long
    Dim i As Integer
    
    Menu.ListView2.ListItems.Add , , "Idle"
    With Menu.ListView2.ListItems(1)
        .SubItems(1) = 0
        .SubItems(2) = 0
    End With

    lngSList = 180: lngAList = 136 'XP硬编码
    
    'hModule = LoadLibraryEx(GetDeviceDriver(BaseName), 0, 1)
    PsInitialSystemProcess = GetProcAddress(pKernel, "PsInitialSystemProcess")
    PsInitialSystemProcess = PsInitialSystemProcess + GetDeviceDriver(BaseAddress) - pKernel
    'FreeLibrary hModule
    
    'System
    ReadKernelMemory ByVal PsInitialSystemProcess, ByVal VarPtr(EPROCESS), 4, 0
    ReDim Preserve Ret(0)
    Ret(0) = EPROCESS
    'MsgBox "System EPROCESS:" & FormatHex(EPROCESS)
    
    'smss.exe
    ReadKernelMemory ByVal (EPROCESS + lngAList), ByVal VarPtr(EPROCESS), 4, 0
    EPROCESS = EPROCESS - lngAList
    ReDim Preserve Ret(1)
    Ret(1) = EPROCESS
    'MsgBox "smss.exe EPROCESS:" & FormatHex(EPROCESS)
    
    Dim PID As Long
    Dim hProcess As Long
    Dim pbi As PROCESS_BASIC_INFORMATION
    Dim pPath As String
    Dim pName As String
    
    
    For i = 0 To 1
        ReadKernelMemory ByVal Ret(i) + &H84, ByVal VarPtr(PID), 4, 0
        hProcess = FxNormalOpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, PID)
        ZwQueryInformationProcess hProcess, ProcessBasicInformation, pbi, Len(pbi), 0
        pPath = GetProcessPath(hProcess)
        pName = GetProcessName(pPath)
            
        Menu.ListView2.ListItems.Add , , pName
        With Menu.ListView2.ListItems(Menu.ListView2.ListItems.count)
            .SubItems(1) = CStr(PID)
            .SubItems(2) = CStr(pbi.InheritedFromUniqueProcessId)
            .SubItems(3) = FormatHex(pbi.PebBaseAddress)
            .SubItems(4) = FormatHex(Ret(i))
            .SubItems(5) = PriorityCheck(pbi.BasePriority)
            .SubItems(6) = FxGetProcessMemoryInformation(hProcess)
            .SubItems(7) = pPath
            .SubItems(8) = GetProcessCommandLine(hProcess)
        End With
            
        ZwClose hProcess: hProcess = 0
    Next i
    
    ReadKernelMemory ByVal (EPROCESS + lngAList), ByVal VarPtr(etStart), 4, 0
    FxAddSystemProcess = etStart - lngAList
    'MsgBox "etStart:" & FormatHex(etStart)
End Function

Public Sub RdNewByHandleList()
    Dim st As Long
    Dim i As Long, j As Long, k As Long
    Dim PID(65536) As Boolean
    Call RefreshHandleTable
    For i = 1 To NumOfHandle
        With HandleTable(i)
            If .ObjectTypeIndex = OB_TYPE_PROCESS Then
                j = .pObject
                k = PsGetPidByEProcess(j)
                If PID(k \ 4) = False Then
                    PID(k \ 4) = True
                    Dim nItem As Long
                    nItem = NewProcess(k)
                    With Processes(nItem)
                        .EPROCESS = j
                        ProcessFillByEProcess nItem
                        FillProcessByHandle nItem
                    End With
                End If
            End If
        End With
    Next i
End Sub

Public Sub FillEProcess()
    Dim nItem As Long
    
    RefreshHandleTable
    For nItem = 0 To UBound(Processes)
        With Processes(nItem)
            If .Handle <> 0 Then
                Dim inf As SYSTEM_HANDLE_TABLE_ENTRY_INFO
                RdQueryHandleInformation .Handle, inf, -1
                .EPROCESS = inf.pObject
                ProcessFillByEProcess nItem
            End If
        End With
    Next
End Sub

Public Function PriorityCheck(ByVal Pcb As Long) As String
    '/**函数功能:判断进程优先级，返回字符串**/
    Select Case Pcb
    Case Is > 9
        PriorityCheck = "较高" & "[" & (Pcb) & "]"
    Case Is >= 7
        PriorityCheck = "标准" & "[" & (Pcb) & "]"
    Case Is >= 4
        PriorityCheck = "较低" & "[" & (Pcb) & "]"
    Case Else
        PriorityCheck = "特殊" & "[" & (Pcb) & "]"
    End Select
End Function

Public Function GetProcessState(ByVal frmhWnd As Long, Optional Timeout As Long = 20) As String
    Dim Results As Long

    If Not SendMessageTimeout(frmhWnd, ByVal 0, ByVal 0, ByVal 0, SMTO_ABORTIFHUNG, Timeout, Results) = 1 Then
        'If Results = 0 Then GetState = True
        GetProcessState = "正常"
    Else
        GetProcessState = "挂起"
    End If
End Function

Public Function GetProcessPath(ByVal hProcess As Long) As String
    '/**函数功能:由进程句柄获取进程路径**/
    On Error Resume Next

    Dim hModule As Long
    Dim Ret As Long
    Dim szPathName As String

    If hProcess <> 0 Then
        Ret = EnumProcessModules(hProcess, hModule, 4, 0)
        If Ret <> 0 Then
            szPathName = Space(260)
            Ret = GetModuleFileNameEx(hProcess, hModule, szPathName, 260)
            GetProcessPath = left(szPathName, Ret)
        End If
    End If

    If GetProcessPath = "" Then
        GetProcessPath = "System"
    End If
End Function

Public Function GetProcessCommandLine(ByVal hProcess As Long) As String
    '/**函数功能:由PID获取进程命令行**/
    Dim NTSTATUS As Long
    Dim objBasic As PROCESS_BASIC_INFORMATION
    Dim objBaseAddress As Long
    Dim bytName() As Byte
    Dim strModuleName As String
    Dim obj As Long
    Dim dwSize As Long
    
    If hProcess = 0 Then
        GetProcessCommandLine = ""
        Exit Function
    End If
           
    Dim lngRet As Long, lngReturn As Long
    
    NTSTATUS = ZwQueryInformationProcess(hProcess, ProcessBasicInformation, objBasic, Len(objBasic), dwSize)
    If (NT_SUCCESS(NTSTATUS)) Then
        '获取PEB指针
        '获取_RTL_USER_PROCESS_PARAMETERS结构指针
        ZwReadVirtualMemory hProcess, ByVal objBasic.PebBaseAddress + &H10, obj, 4, lngRet
        If lngRet <> 4 Then Exit Function
        '获取路径长度
        ZwReadVirtualMemory hProcess, ByVal obj + &H40, dwSize, 2, lngRet
        If lngRet <> 2 Then Exit Function
        '获取路径指针
        ZwReadVirtualMemory hProcess, ByVal obj + &H44, obj, 4, lngRet
        If lngRet <> 4 Then Exit Function
        ReDim bytName(dwSize - 1)
        '获取路径
        ZwReadVirtualMemory hProcess, ByVal obj, bytName(0), dwSize, lngRet
        If lngRet <> dwSize Then Exit Function
        GetProcessCommandLine = bytName
     End If
End Function

Public Function GetProcessName(ByVal Path As String, Optional ByVal FindText As String = "\") As String
    '/**函数功能:由进程路径获取进程名**/
    GetProcessName = Mid$(Path, InStrRev(Path, FindText) + 1)
End Function

Public Function FxGetProcessName(ByVal EPROCESS As Long) As String
    Dim proName As String * 16 'richway2修改：MAX_PATH改为16
    Dim byBuff(MAX_PATH) As Byte
    
    ReadKernelMemory EPROCESS + &H174, VarPtr(byBuff(0)), 16, 0
    FxGetProcessName = Replace(StrConv(byBuff(), vbUnicode), Chr(0), "")
End Function

Public Function FxNormalOpenProcess(ByVal dwDesiredAccess As Long, ByVal PID As Long) As Long
    '/**函数功能:打开一个进程，失败则调用LzOpenProcess**/
    Dim oa As OBJECT_ATTRIBUTES
    Dim Cid As CLIENT_ID
    Dim st As Long
    Dim hProcess As Long
    
    oa.Length = LenB(oa)

    Cid.UniqueProcess = PID

    st = ZwOpenProcess(hProcess, dwDesiredAccess, oa, Cid)
    If Not NT_SUCCESS(st) Then
        hProcess = LzOpenProcess(dwDesiredAccess, PID)
    End If

    FxNormalOpenProcess = hProcess
End Function

Public Function LzOpenProcess(ByVal dwDesiredAccess As Long, ByVal ProcessID As Long) As Long
    '/**函数功能:通过复制句柄表里的句柄来“打开”进程**/
    Dim st As Long
    Dim Cid As CLIENT_ID
    Dim oa As OBJECT_ATTRIBUTES
    Dim NumOfHandle As Long
    Dim pbi As PROCESS_BASIC_INFORMATION
    Dim i As Long
    Dim hProcessToDup As Long, hProcessCur As Long, hProcessToRet As Long
    
    'oa.Length = Len(oa)
    ''首先尝试ZwOpenProcess
    'Cid.UniqueProcess = ProcessID
    'st = ZwOpenProcess(hProcessToRet, dwDesiredAccess, oa, Cid)
    'If (NT_SUCCESS(st)) Then LzOpenProcess = hProcessToRet: Exit Function
    st = 0
    
    Dim bytBuf() As Byte
    Dim arySize As Long: arySize = 1
    st = ZwQuerySystemInformation(SystemHandleInformation, 0, 0, arySize)
    If st <> STATUS_INFO_LENGTH_MISMATCH Or arySize = 0 Then
        Exit Function
    End If
Again:
    ReDim bytBuf(arySize)
    st = ZwQuerySystemInformation(SystemHandleInformation, VarPtr(bytBuf(0)), arySize, arySize)
    If Not NT_SUCCESS(st) Then GoTo Again
    
    NumOfHandle = 0
    CopyMemory VarPtr(NumOfHandle), VarPtr(bytBuf(0)), Len(NumOfHandle)
    Dim h_info() As SYSTEM_HANDLE_TABLE_ENTRY_INFO
    ReDim h_info(NumOfHandle)
    CopyMemory VarPtr(h_info(0)), VarPtr(bytBuf(0)) + Len(NumOfHandle), Len(h_info(0)) * NumOfHandle
    
    '//枚举句柄完成，下来开始测试句柄
    For i = LBound(h_info) To UBound(h_info)
        With h_info(i)
            If (.ObjectTypeIndex = OB_TYPE_PROCESS) Then
                Cid.UniqueProcess = .UniqueProcessId
                st = ZwOpenProcess(hProcessToDup, PROCESS_DUP_HANDLE, oa, Cid)
                If (NT_SUCCESS(st)) Then
                    st = ZwDuplicateObject(hProcessToDup, .HandleValue, ZwGetCurrentProcess, hProcessCur, dwDesiredAccess Or PROCESS_QUERY_INFORMATION, 0, DUPLICATE_SAME_ATTRIBUTES)
                    If (NT_SUCCESS(st)) Then
                        st = ZwQueryInformationProcess(hProcessCur, ProcessBasicInformation, pbi, Len(pbi), 0)
                        If (NT_SUCCESS(st)) Then
                            If (pbi.UniqueProcessId = ProcessID) Then
                                st = ZwDuplicateObject(hProcessToDup, .HandleValue, ZwGetCurrentProcess, hProcessToRet, dwDesiredAccess, 0, DUPLICATE_SAME_ATTRIBUTES)
                                If (NT_SUCCESS(st)) Then
                                    If hProcessToRet <> 0 Then
                                        LzOpenProcess = hProcessToRet
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                    st = ZwClose(hProcessCur)
                End If
                st = ZwClose(hProcessToDup)
            End If
        End With
    Next i
    
    Erase h_info
End Function

Public Function RdOpenProcess(ByVal mPid As Long) As Long
    '直接修改本进程的句柄表来打开进程
    Dim mHandle As Long
    Dim dwPid As Long
    Dim st As Long
    mHandle = OpenProcess(PROCESS_QUERY_INFORMATION, False, GetCurrentProcessId)
    Dim mEProcess As Long, mBuffer As SYSTEM_HANDLE_TABLE_ENTRY_INFO
    RdQueryHandleInformation mHandle, mBuffer
    mEProcess = mBuffer.pObject
    ZwClose mHandle
    Dim bytBuf() As Byte
    Dim arySize As Long
    arySize = 1
    Do
        ReDim bytBuf(arySize)
        st = ZwQuerySystemInformation(16, VarPtr(bytBuf(0)), arySize, 0&)
        If (Not NT_SUCCESS(st)) Then
            If (st <> &HC0000004) Then
                Erase bytBuf
                Exit Function
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
    Dim i As Long, oTarget As Long
    For i = 0 To NumOfHandle - 1
        If h_info(i).ObjectTypeIndex = OB_TYPE_PROCESS Then
            If PsGetPidByEProcess(h_info(i).pObject) = mPid Then
                oTarget = h_info(i).pObject
                Exit For
            End If
        End If
    Next
    '取得本进程的EProcess和目标进程的EProcess
    
    '读取句柄表地址
    Dim mHandleTable As Long
    ReadKernelMemory mEProcess + &HC4, VarPtr(mHandleTable), 4, 0
    Dim mHandleNum As Long
    ReadKernelMemory mHandleTable + &H3C, VarPtr(mHandleNum), 4, 0
    Dim TableCode As Long '读取句柄表标志
    ReadKernelMemory mHandleTable, VarPtr(TableCode), 4, 0
    TableCode = TableCode And 3 '读取句柄表级数
    If TableCode >= 2 Then Exit Function '三级表
    'If TableCode = 1 Then
        'Call Table1_Enum(mHandleTable
End Function

Public Function FxGetProcessId(ByVal hProcess As Long) As Long
    '/**函数功能:由进程句柄获取PID**/
    Dim pbi As PROCESS_BASIC_INFORMATION
    Dim st As Long
    
    st = ZwQueryInformationProcess(hProcess, ProcessBasicInformation, pbi, Len(pbi), 0)
    If (NT_SUCCESS(st)) Then
        FxGetProcessId = pbi.UniqueProcessId
    End If
End Function

Public Function FxGetObjectTypeProcess() As Long
    '/**函数功能:获取进程的句柄类型的索引**/
    Dim mHandle, mPid As Long
    Dim st As Long
       
    mPid = GetCurrentProcessId
    
    st = ZwDuplicateObject(GetCurrentProcess, GetCurrentProcess, GetCurrentProcess, mHandle, PROCESS_ALL_ACCESS, 0, DUPLICATE_SAME_ATTRIBUTES)
    
    If NT_SUCCESS(st) Then
        Dim bytBuf() As Byte
        Dim arySize As Long
        
        arySize = 1
        Do
            ReDim bytBuf(arySize)
            st = ZwQuerySystemInformation(SystemHandleInformation, VarPtr(bytBuf(0)), arySize, 0)
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
        
        Dim NumOfHandle As Long
        NumOfHandle = 0
        CopyMemory VarPtr(NumOfHandle), VarPtr(bytBuf(0)), Len(NumOfHandle)
        Dim h_info() As SYSTEM_HANDLE_TABLE_ENTRY_INFO
        ReDim h_info(NumOfHandle)
        CopyMemory VarPtr(h_info(0)), VarPtr(bytBuf(0)) + Len(NumOfHandle), Len(h_info(0)) * NumOfHandle
        
        Dim i As Long
        For i = 1 To NumOfHandle
            If h_info(i).HandleValue = mHandle And h_info(i).UniqueProcessId = mPid Then
                ZwClose mHandle
                FxGetObjectTypeProcess = h_info(i).ObjectTypeIndex
                Exit For
            End If
        Next i
    End If
End Function

Public Sub FxGetProcessEProcess(ByRef Listview As Object, ByVal pidColumn As Long, ByVal epColumn As Long)
    '/**函数功能:填充Lsitview中的EPROCESS项**/
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
    Dim nowPid As Long
    
    For i = LBound(h_info) To UBound(h_info) / 4
        With h_info(i)
            If .ObjectTypeIndex = OB_TYPE_PROCESS Then
                nowPid = PsGetPidByEProcess(.pObject)
                For j = 0 To UBound(Processes)
                    If Processes(j).Basic.UniqueProcessId = nowPid Then
                        Processes(j).EPROCESS = .pObject
                        Exit For
                    End If
                Next j
            End If
        End With
    Next i
    
    Erase h_info
End Sub

Public Function PsGetImageFileNameByEProcess(ByVal EPROCESS As Long) As String
    '/**函数功能:由EPROCESS获取进程名**/
    ReadKernelMemory EPROCESS + &H174, VarPtr(PsGetImageFileNameByEProcess), 4, 0
End Function

Public Function PsGetPidByEProcess(ByVal EPROCESS As Long) As Long
    '/**函数功能:由EPROCESS获取PID**/
    ReadKernelMemory EPROCESS + &H84, VarPtr(PsGetPidByEProcess), 4, 0
End Function

Public Function FxGetCurrentEProcess() As Long
    '/**函数功能:获取自身的EPROCESS**/
    Dim mHandle As Long
    Dim dwPid As Long
    Dim st As Long
       
    dwPid = GetCurrentProcessId
    mHandle = OpenProcess(PROCESS_QUERY_INFORMATION, False, dwPid)
    
    If NT_SUCCESS(st) Then
        Dim bytBuf() As Byte
        Dim arySize As Long
        arySize = 1
        Do
            ReDim bytBuf(arySize)
            st = ZwQuerySystemInformation(16, VarPtr(bytBuf(0)), arySize, 0&)
            If (Not NT_SUCCESS(st)) Then
                If (st <> &HC0000004) Then
                    Erase bytBuf
                    Exit Function
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
        For i = 0 To NumOfHandle
            If h_info(i).HandleValue = mHandle And h_info(i).UniqueProcessId = dwPid Then
                FxGetCurrentEProcess = h_info(i).pObject
                Exit For
            End If
        Next i
    End If
End Function

Public Sub RdUnlockProcess(ByVal EPROCESS As Long)
    '尝试解锁进程来方便打开进程
    Dim fpProtect As Long
    ReadKernelMemory EPROCESS + &H80, VarPtr(fpProtect), 4, 0
    If fpProtect <> 0 Then
        Debug.Print "检测到保护模式：" & fpProtect
        WriteKernelMemory EPROCESS + &H80, VarPtr(CLng(0)), 4
    End If
    ReadKernelMemory EPROCESS, VarPtr(fpProtect), 4, 0
    If fpProtect = 0 Then
        MsgBox "解锁失败", vbCritical
    End If
    Dim fpProtect2 As Long, fTest As Long
    ReadKernelMemory fpProtect + &H58, VarPtr(fpProtect2), 4, 0
    If fpProtect <> 0 Then
        Debug.Print "检测到保护模式：" & fpProtect2
        If WriteKernelMemory(fpProtect2 + &H58, VarPtr(CLng(0)), 4) Then
            MsgBox "完成", vbInformation
        Else
            MsgBox "写入内核内存时出错！", vbCritical
        End If
    Else
        MsgBox "没必要执行解锁或使用了其他的锁定方式！", vbInformation
    End If
End Sub

Public Sub FxTerminateProcessByDebugProcess(ByVal PID As Long)
    '/**函数功能:通过调试进程的方法结束进程**/
    Dim hDebug As Long
    Dim hProcess As Long
    Dim Status As Long
    Dim errStr As String
       
    '建立调试对象
    If Not NT_SUCCESS(ZwCreateDebugObject(hDebug, &H1F000F, 0&, 1&)) Then errStr = "建立调试对象失败!": GoTo errors

    '获得调试句柄
    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, PID)
    If hProcess <= 0 Then ZwClose hDebug: errStr = "拒绝访问!": GoTo errors
    
    '接管调试进程然后退出
    Status = ZwDebugActiveProcess(hProcess, hDebug)
    ZwResumeProcess hProcess
    ZwClose hProcess
    ZwClose hDebug
    
    '判断是否成功
    If Not NT_SUCCESS(Status) Then errStr = "调试进程失败!": GoTo errors
Exit Sub
errors:
    MsgBox errStr, 0, "失败"
End Sub

Public Sub PNNew()
    '/**函数功能:智能判断遍历进程方法并刷新Lsitview(刷新列表时请使用此函数)**/
    Dim pIndex As Long
    
    If Menu.ListView2.Sorted = True Then Menu.ListView2.Sorted = False
    
    MainState = Not MainState
    
    'LockWindowUpdate Menu.ListView2.hwnd
    
    Select Case Menu.ListView2.Tag
    Case MethodSnapshot
        Call mpNew_Click
    Case MethodSession
        Call FxListProcessBySession
    Case MethodEnumProcesses
    Case MethodTest
        'Call RwNewByTest
    Case MethodHandleList
        Call RdNewByHandleList
    Case MethodQuery
        Call ListProcessByQuery
    End Select
    
    
    Dim nItem As Long, i As Long
    For nItem = 0 To UBound(Processes)
        Dim etNow As AzmrkProcess, j As Long
        If Processes(nItem).Handle <> 0 Then ProcessAntiFill nItem
        etNow = Processes(nItem)
        Processes(nItem).ListViewIndex = etNow.ListViewIndex - j
        etNow.ListViewIndex = etNow.ListViewIndex - j
        If etNow.State <> MainState And etNow.ListViewIndex <> 0 Then
            ZwClose etNow.Handle
            Menu.ListView2.ListItems.Remove etNow.ListViewIndex
            Processes(nItem).ListViewIndex = 0
            j = j + 1
        ElseIf etNow.ListViewIndex <> 0 Then
            With Menu.ListView2.ListItems(etNow.ListViewIndex)
                .ListSubItems.Clear
                .Text = etNow.ImageName
                For i = 1 To 14
                    If Menu.pColumnSelect(i).Checked Then .ListSubItems.Add , , FillSubItem(etNow, i)
                Next
            End With
        End If
    Next
    
    Menu.Label3.Caption = "共有" & (Menu.ListView2.ListItems.count) & "个进程"
End Sub

Public Function FillSubItem(etNow As AzmrkProcess, ByVal nNum As Long) As String
    Select Case RealProcessColumnNames(nNum)
    Case "进程ID": FillSubItem = etNow.Basic.UniqueProcessId
    Case "父进程ID": FillSubItem = etNow.Basic.InheritedFromUniqueProcessId
    Case "PEB": FillSubItem = FormatHex(etNow.Basic.PebBaseAddress)
    Case "EPROCESS": FillSubItem = FormatHex(etNow.EPROCESS)
    Case "优先级": FillSubItem = PriorityCheck(etNow.Basic.BasePriority)
    Case "内存使用": FillSubItem = ByteToKMG(etNow.VmCounters.WorkingSetSize) & " - " & ByteToKMG(etNow.VmCounters.PeakWorkingSetSize)
    Case "IO读取次数": FillSubItem = FileTime2String(etNow.IoCounters.ReadOperationCount)
    Case "IO写入次数": FillSubItem = FileTime2String(etNow.IoCounters.WriteOperationCount)
    Case "IO其他次数": FillSubItem = FileTime2String(etNow.IoCounters.OtherOperationCount)
    Case "IO读取字节": FillSubItem = FileTime2String(etNow.IoCounters.ReadTransferCount)
    Case "IO写入字节": FillSubItem = FileTime2String(etNow.IoCounters.WriteTransferCount)
    Case "IO其他字节": FillSubItem = FileTime2String(etNow.IoCounters.OtherTransferCount)
    Case "映像路径": FillSubItem = etNow.ExePath
    Case "命令行": FillSubItem = etNow.CmdLine
    End Select
End Function

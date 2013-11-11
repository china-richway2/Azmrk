Attribute VB_Name = "Debugger"
Public Declare Function WaitForDebugEvent Lib "kernel32.dll" (lpde As DEBUG_EVENT, ByVal dwTimeout As Long) As Long
Public Declare Function ZwDebugActiveProcess Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, ByVal DebugObjectHandle As Long) As Long
Public Declare Function ZwCreateDebugObject Lib "NTDLL.DLL" (ByRef pDebugObjectHandle As Long, ByVal DesiredAccess As Long, ByVal pObjectAttributes As Long, ByVal Flags As Long) As Long
Public Declare Function ZwWaitForDebugEvent Lib "NTDLL.DLL" (ByVal hDebug As Long, ByVal Flags As Long, lpTimeout As INT64, lpEvent As Any) As Long
'Public Declare Function DebugActiveProcess Lib "kernel32" (ByVal dwProcessId As Long) As Long
'Public Declare Function ContinueDebugEvent Lib "kernel32" (ByVal dwProcessId As Long, ByVal dwThreadId As Long, ByVal dwContinueStatus As Long) As Long
Public Declare Function ZwDebugContinue Lib "NTDLL.DLL" (ByVal hDebug As Long, cid As Any, ByVal ContinueStatus As Long) As Long
Public Declare Function DebugActiveProcessStop Lib "kernel32" (ByVal dwProcessId As Long) As Long
Public Declare Function DbgUiConvertStateChangeStructure Lib "NTDLL.DLL" (lpDebugEvent As Any, lpNtEvent As Any) As Long
Public Declare Function DbgUiIssueRemoteBreakin Lib "NTDLL.DLL" (ByVal hProcess As Long) As Long
Public Declare Sub UnionToType Lib "NTDLL.DLL" Alias "RtlMoveMemory" (dest As Any, Src As Any, ByVal Length As Long)
Public Declare Function ZwGetContextThread Lib "NTDLL.DLL" (ByVal hThread As Long, ByRef pContext As Any) As Long
Public Declare Function ZwSetContextThread Lib "NTDLL.DLL" (ByVal hThread As Long, ByRef pContext As Any) As Long
Public Declare Function ZwRemoveProcessDebug Lib "NTDLL.DLL" (ByVal hDProcess As Long, ByVal hDebug As Long) As Long
Public Declare Function ZwSetInformationDebugObject Lib "NTDLL.DLL" (ByVal hDebug As Long, ByVal dwInfoClass As Long, lpInformation As Any, ByVal dwLength As Long, lpLength As Long) As Long
Public Declare Function GetThreadSelectorEntry Lib "kernel32" (ByVal hThread As Long, ByVal dwSelector As Long, lpSelectorEntry As LDT_ENTRY) As Long
Public Const DBG_CONTINUE = &H10002
Public Const DBG_CONTROL_BREAK = &H40010008
Public Const DBG_CONTROL_C = &H40010005
Public Const DBG_EXCEPTION_NOT_HANDLED = &H80010001
Public Const DBG_TERMINATE_PROCESS = &H40010004
Public Const DBG_TERMINATE_THREAD = &H40010003
Public Const EXCEPTION_MAXIMUM_PARAMETERS = 15
Public Const STATUS_ACCESS_VIOLATION = &HC0000005
Public Const STATUS_ARRAY_BOUNDS_EXCEEDED = &HC000008C
Public Const STATUS_BREAKPOINT = &H80000003
Public Const STATUS_DATATYPE_MISALIGNMENT = &H80000002
Public Const STATUS_FLOAT_DENORMAL_OPERAND = &HC000008D
Public Const STATUS_FLOAT_DIVIDE_BY_ZERO = &HC000008E
Public Const STATUS_FLOAT_INEXACT_RESULT = &HC000008F
Public Const STATUS_FLOAT_INVALID_OPERATION = &HC0000090
Public Const STATUS_FLOAT_OVERFLOW = &HC0000091
Public Const STATUS_FLOAT_STACK_CHECK = &HC0000092
Public Const STATUS_FLOAT_UNDERFLOW = &HC0000093
Public Const STATUS_INTEGER_DIVIDE_BY_ZERO = &HC0000094
Public Const STATUS_INTEGER_OVERFLOW = &HC0000095
Public Const STATUS_IN_PAGE_ERROR = &HC0000006
Public Const STATUS_PRIVILEGED_INSTRUCTION = &HC0000096
Public Const STATUS_SINGLE_STEP = &H80000004

Public Const EXCEPTION_ACCESS_VIOLATION = STATUS_ACCESS_VIOLATION
Public Const EXCEPTION_ARRAY_BOUNDS_EXCEEDED = STATUS_ARRAY_BOUNDS_EXCEEDED
Public Const EXCEPTION_BREAKPOINT = STATUS_BREAKPOINT
Public Const EXCEPTION_CONTINUABLE = 0
Public Const EXCEPTION_CONTINUE_EXECUTION = -1
Public Const EXCEPTION_CONTINUE_SEARCH = 0
Public Const EXCEPTION_DATATYPE_MISALIGNMENT = STATUS_DATATYPE_MISALIGNMENT
Public Const EXCEPTION_EXECUTE_HANDLER = 1
Public Const EXCEPTION_FLT_DENORMAL_OPERAND = STATUS_FLOAT_DENORMAL_OPERAND
Public Const EXCEPTION_FLT_DIVIDE_BY_ZERO = STATUS_FLOAT_DIVIDE_BY_ZERO
Public Const EXCEPTION_FLT_INEXACT_RESULT = STATUS_FLOAT_INEXACT_RESULT
Public Const EXCEPTION_FLT_INVALID_OPERATION = STATUS_FLOAT_INVALID_OPERATION
Public Const EXCEPTION_FLT_OVERFLOW = STATUS_FLOAT_OVERFLOW
Public Const EXCEPTION_FLT_STACK_CHECK = STATUS_FLOAT_STACK_CHECK
Public Const EXCEPTION_FLT_UNDERFLOW = STATUS_FLOAT_UNDERFLOW
Public Const EXCEPTION_INT_DIVIDE_BY_ZERO = STATUS_INTEGER_DIVIDE_BY_ZERO
Public Const EXCEPTION_INT_OVERFLOW = STATUS_INTEGER_OVERFLOW
Public Const EXCEPTION_IN_PAGE_ERROR = STATUS_IN_PAGE_ERROR
Public Const EXCEPTION_NONCONTINUABLE = &H1
Public Const EXCEPTION_PRIV_INSTRUCTION = STATUS_PRIVILEGED_INSTRUCTION
Public Const EXCEPTION_SINGLE_STEP = STATUS_SINGLE_STEP
Public Const CONTEXT_i386 = &H10000
Public Const CONTEXT_CONTROL = CONTEXT_i386 Or &H1&
Public Const CONTEXT_INTEGER = CONTEXT_i386 Or &H2&
Public Const CONTEXT_SEGMENTS = CONTEXT_i386 Or &H4&
Public Const CONTEXT_FLOATING_POINT = CONTEXT_i386 Or &H8&
Public Const CONTEXT_DEBUG_REGISTERS = CONTEXT_i386 Or &H10&
Public Const CONTEXT_FULL = CONTEXT_CONTROL Or CONTEXT_i386 Or &H1F
Public Enum EFLFlags
    FlagCF = &H1
    FlagPF = &H4
    FlagAF = &H8
    FlagZF = &H10
    FlagSF = &H20
    FlagTF = &H100
    FlagIF = &H200
    FlagDF = &H400
    FlagOF = &H800
    FlagRF = &H10000
End Enum
Public Type EXCEPTION_RECORD
    ExceptionCode As Long
    ExceptionFlags As Long
    pExceptionRecord As Long '指向EXCEPTION_RECORD
    ExceptionAddress As Long
    NumberParameters As Long
    ExceptionInformation(0 To EXCEPTION_MAXIMUM_PARAMETERS - 1) As Long
End Type
Public Enum DebugEventCode
    EXCEPTION_DEBUG_EVENT = 1
    CREATE_THREAD_DEBUG_EVENT = 2
    CREATE_PROCESS_DEBUG_EVENT = 3
    EXIT_THREAD_DEBUG_EVENT = 4
    EXIT_PROCESS_DEBUG_EVENT = 5
    LOAD_DLL_DEBUG_EVENT = 6
    UNLOAD_DLL_DEBUG_EVENT = 7
    OUTPUT_DEBUG_STRING_EVENT = 8
    RIP_EVENT = 9
End Enum
Public Type CREATE_THREAD_DEBUG_INFO
    hThread As Long
    lpThreadLocalBase As Long
    lpStartAddress As Long
End Type
Public Type NT_DEBUG_EVENT
    EventCode As Long
    dwProcessId As Long
    dwThreadId As Long
    R(20) As Long
End Type

Public Type CREATE_PROCESS_DEBUG_INFO
    hFile As Long
    hProcess As Long
    hThread As Long
    lpBaseOfImage As Long
    dwDebugInfoFileOffset As Long
    nDebugInfoSize As Long
    lpThreadLocalBase As Long
    lpStartAddress As Long
    lpImageName As Long
    fUnicode As Integer
End Type

Public Type EXIT_THREAD_DEBUG_INFO
    dwExitCode As Long
End Type

Public Type EXIT_PROCESS_DEBUG_INFO
    dwExitCode As Long
End Type

Public Type OUTPUT_DEBUG_STRING_INFO
    lpDebugStringData As Long 'ASCIIZ PTR
    fUnicode As Integer
    nDebugStringLength As Integer
End Type

Public Type RIP_INFO
    dwError As Long
    dwType As Long
End Type

Public Type EXCEPTION_DEBUG_INFO
    ExceptionRecord As EXCEPTION_RECORD
    dwFirstChange As Long
End Type

Public Type UNLOAD_DLL_DEBUG_INFO
     lpBaseOfDll As Long
End Type

Public Type LOAD_DLL_DEBUG_INFO
    hFile As Long
    lpBaseOfDll As Long
    dwDebugInfoFileOffset As Long
    nDebugInfoSize As Long
    lpImageName As Long '指向文件名
    fUnicode As Integer
End Type
'Public Type DEBUGEVENTUNION
'    deuException As EXCEPTION_DEBUG_INFO
'    deuCreateThread As CREATE_THREAD_DEBUG_INFO
'    deuCreateProcessInfo As CREATE_PROCESS_DEBUG_INFO
'    deuExitThread As EXIT_THREAD_DEBUG_INFO
'    deuExitProcess As EXIT_PROCESS_DEBUG_INFO
'    deuLoadDll As LOAD_DLL_DEBUG_INFO
'    deuUnloadDll As UNLOAD_DLL_DEBUG_INFO
'    deuDebugString As OUTPUT_DEBUG_STRING_INFO
'    deuRipInfo As RIP_INFO
'End Type

Public Type DEBUG_EVENT
    dwDebugEventCode As Long
    dwProcessId As Long
    dwThreadId As Long
    'u As DEBUGEVENTUNION
End Type

Public Type DebugThread
    dwThreadId As Long
    dwThreadHandle As Long
    hWindow As DebugThreadWindow
End Type

Public Type Times
    CreationTime As FILETIME
    ExitTime As FILETIME
    UserTime As Currency
    KernelTime As Currency
End Type

Public Type FLOATING_SAVE_AREA
    ControlWord As Long
    StatusWord As Long
    TagWord As Long
    ErrorOffset As Long
    ErrorSelector As Long
    DataOffset As Long
    DataSelector As Long
    RegisterArea(80 - 1) As Byte
    Cr0NpxState As Long
End Type

Public Type FLGSTRUCT
    fCF As Byte
    fPF As Byte
    fAF As Byte
    fZF As Byte
    fSF As Byte
    fTF As Byte
    fIF As Byte
    fDF As Byte
    fOF As Byte
    fRF As Byte
End Type

Public Type DbgFunctionInfo
    ProcOffset As Long
    FunName As String
End Type

Public Type DbgModuleInfo
    ModuleName As String
    ModuleHandle As Long
    ModuleSize As Long
    Procs() As DbgFunctionInfo
End Type

Public Type IMAGE_EXPORT_DIRECTORY           '导出表，40个字节
    Unused(11) As Byte
    nName As Long                            '指向文件名的RVA
    nBase As Long                            '指向导出函数的起始序号
    NumberOfFunctions As Long                '文件函数的总数
    NumberOfNames As Long                    '以名称导出的函数总数
    AddressOfFunctions As Long               '指向导出函数地址表的RVA
    AddressOfNames As Long                   '指向函数名地址表的RVA
    AddressOfNameOrdinals As Long            '指向函数名序号表的RVA
End Type

Public Type LDT_BYTES
    BaseMid As Byte
    Flags1 As Byte
    Flags2 As Byte
    BaseHi As Byte
End Type

Public Type LDT_ENTRY
    LimitLow As Integer
    BaseLow As Integer
    HighWord As LDT_BYTES
End Type

Function SegmentToNum(n As LDT_ENTRY) As Long
    Dim A(3) As Byte
    UnionToType A(0), n.BaseLow, 2
    A(2) = n.HighWord.BaseMid
    A(3) = n.HighWord.BaseHi
    UnionToType SegmentToNum, A(0), 4
End Function

Public Sub AttachDebugger(ByVal dwPid As Long)
    Load DebugWindow
    DebugWindow.Attach dwPid
End Sub

Public Sub DbgFillModuleInfo(ByVal hProcess As Long, lpInfo As DbgModuleInfo)
    Dim tmp As Long, Export As IMAGE_EXPORT_DIRECTORY
    On Error GoTo Ends
    With lpInfo
        Dim szModuleName As String * 260
        tmp = GetModuleFileNameEx(hProcess, lpInfo.ModuleHandle, szModuleName, 260)
        lpInfo.ModuleName = left(szModuleName, tmp)
        ZwReadVirtualMemory hProcess, ByVal lpInfo.ModuleHandle + &H3C, tmp, 4, 0 '获取e_lfanew
        ZwReadVirtualMemory hProcess, ByVal lpInfo.ModuleHandle + tmp + &H78, tmp, 4, 0 '获取导出表地址
        ZwReadVirtualMemory hProcess, ByVal lpInfo.ModuleHandle + tmp, Export, 40, 0
        Dim Functions() As Long, Names() As Long, NameOrdinals() As Integer
        ReDim Functions(Export.NumberOfFunctions - 1)
        ReDim Names(Export.NumberOfNames - 1)
        ReDim NameOrdinals(Export.NumberOfNames - 1)
        ZwReadVirtualMemory hProcess, ByVal lpInfo.ModuleHandle + Export.AddressOfNames, Names(0), Export.NumberOfNames * 4, 0
        ZwReadVirtualMemory hProcess, ByVal lpInfo.ModuleHandle + Export.AddressOfFunctions, Functions(0), Export.NumberOfFunctions * 4, 0
        ZwReadVirtualMemory hProcess, ByVal lpInfo.ModuleHandle + Export.AddressOfNameOrdinals, NameOrdinals(0), Export.NumberOfNames * 2, 0
        ReDim .Procs(UBound(Functions))
        For tmp = Export.nBase To UBound(.Procs)
            With .Procs(tmp)
                ZwReadVirtualMemory hProcess, ByVal lpInfo.ModuleHandle + Names(tmp - Export.nBase), ByVal szModuleName, 260, 0
                .FunName = left(szModuleName, InStr(szModuleName, vbNullChar) - 1)
                If NameOrdinals(tmp) = 0 Then GoTo FindNext
                .ProcOffset = Functions(NameOrdinals(tmp) - Export.nBase)
            End With
FindNext:
        Next
    End With
Ends:
    'Stop
    'Resume
End Sub

Public Sub GetFlags(ByVal EFL As EFLFlags, fStruct As FLGSTRUCT)
    With fStruct
        .fCF = (EFL And FlagCF) \ FlagCF
        .fPF = (EFL And FlagPF) \ FlagPF
        .fAF = (EFL And FlagAF) \ FlagAF
        .fZF = (EFL And FlagZF) \ FlagZF
        .fSF = (EFL And FlagSF) \ FlagSF
        .fTF = (EFL And FlagTF) \ FlagTF
        .fIF = (EFL And FlagIF) \ FlagIF
        .fDF = (EFL And FlagDF) \ FlagDF
        .fOF = (EFL And FlagOF) \ FlagOF
        .fRF = (EFL And FlagRF) \ FlagRF
    End With
End Sub

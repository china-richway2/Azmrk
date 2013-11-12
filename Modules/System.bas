Attribute VB_Name = "System"
Option Explicit
'[winio]
Public Type tagPhysStruct
    dwPhysMemSizeInBytes As Long
    Reserved1 As Long 'INT64的高位
    pvPhysAddress As Long
    Reserved2 As Long 'INT64的高位
    PhysicalMemoryHandle As Long
    Reserved3 As Long 'INT64的高位
    pvPhysMemLin As Long
    Reserved4 As Long 'INT64的高位
    pvPhysSection As Long
    Reserved5 As Long 'INT64的高位
End Type
    
Public Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Public Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Public Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Public Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long

Public Declare Function InitializeWinIo Lib "winio32" () As Long
Public Declare Function ShutdownWinIo Lib "winio32" () As Long

Public Declare Function InstallWinIoDriver Lib "winio32" (ByVal strWinIoDriverPath As Long, ByVal bIsDemendLoaded As Long) As Long
Public Declare Function RemoveWinIoDriver Lib "winio32" () As Long

Public Declare Function GetPortVal Lib "winio32" (ByVal wPortAddr As Integer, pdwPortVal As Long, ByVal bSize As Byte) As Long
Public Declare Function SetPortVal Lib "winio32" (ByVal wPortAddr As Integer, ByVal dwPortVal As Long, ByVal bSize As Byte) As Long
'bSize只能为1 2或4

Public Declare Function MapPhysToLin Lib "winio32" (Data As tagPhysStruct) As Long
Public Declare Function UnmapPhysicalMemory Lib "winio32" (Data As tagPhysStruct) As Long
Public Declare Function GetPhysLong Lib "winio32" (ByVal pbPhysAddr As Long, pdwPhysVal As Long) As Long
'根据物理地址获取物理地址指向的DWORD值。
Public Declare Function SetPhysLong Lib "winio32" (ByVal pbPhysAddr As Long, ByVal dwPhysVal As Long) As Long
'设置物理地址指向的DWORD值。

Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function ZwQuerySystemInformation Lib "NTDLL.DLL" (ByVal SystemInformationClass As SYSTEM_INFORMATION_CLASS, ByVal pSystemInformation As Long, ByVal SystemInformationLength As Long, ByRef ReturnLength As Long) As Long
Public Declare Function ZwSystemDebugControl Lib "NTDLL.DLL" (ByVal scCommand As SYSDBG_COMMAND, ByVal pInputBuffer As Long, ByVal InputBufferLength As Long, ByVal pOutputBuffer As Long, ByVal OutputBufferLength As Long, ByRef pReturnLength As Long) As Long
Public Declare Function RtlInitUnicodeString Lib "ntdll" (UnicodeString As UNICODE_STRING, ByVal StringPtr As Long) As Long
'Public Declare Function SetSecurityInfo Lib "advapi32.dll" (ByVal Handle As Long, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As Long, ppsidOwner As Long, ppsidGroup As Long, ppDacl As Any, ppSacl As Any) As Long
'Public Declare Function GetSecurityInfo Lib "advapi32.dll" (ByVal Handle As Long, ByVal ObjectType As SE_OBJECT_TYPE, ByVal SecurityInfo As Long, ppsidOwner As Long, ppsidGroup As Long, ppDacl As Any, ppSacl As Any, ppSecurityDescriptor As Long) As Long
'Public Declare Function SetEntriesInAcl Lib "advapi32.dll" Alias "SetEntriesInAclA" (ByVal cCountOfExplicitEntries As Long, pListOfExplicitEntries As EXPLICIT_ACCESS, ByVal OldAcl As Long, NewAcl As Long) As Long
Public Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (LpVersionInformation As OSVERSIONINFO) As Long
'Dim PhysicalMemory As Long, nMemoryManagementMap As Long
'Public Declare Sub BuildExplicitAccessWithName _
Lib "advapi32.dll" _
Alias "BuildExplicitAccessWithNameA" (pExplicitAccess As EXPLICIT_ACCESS, _
ByVal pTrusteeName As String, _
ByVal AccessPermissions As Long, _
ByVal AccessMode As ACCESS_MODE, _
ByVal Inheritance As Long)

'Public Const ERROR_SUCCESS = 0&
Public Const SECTION_MAP_WRITE = &H2
Public Const SECTION_MAP_READ = &H4
Public Const READ_CONTROL = &H20000
Public Const WRITE_DAC = &H40000
Public Const NO_INHERITANCE = 0
Public Const DACL_SECURITY_INFORMATION = &H4

Public Enum ACCESS_MODE
    NOT_USED_ACCESS
    GRANT_ACCESS
    SET_ACCESS
    DENY_ACCESS
    REVOKE_ACCESS
    SET_AUDIT_SUCCESS
    SET_AUDIT_FAILURE
End Enum

Public Enum MULTIPLE_TRUSTEE_OPERATION
    NO_MULTIPLE_TRUSTEE
    TRUSTEE_IS_IMPERSONATE
End Enum

Public Enum TRUSTEE_FORM
    TRUSTEE_IS_SID
    TRUSTEE_IS_NAME
End Enum

Public Enum TRUSTEE_TYPE
    TRUSTEE_IS_UNKNOWN
    TRUSTEE_IS_USER
    TRUSTEE_IS_GROUP
End Enum

Public Type TRUSTEE
    pMultipleTrustee As Long
    MultipleTrusteeOperation As MULTIPLE_TRUSTEE_OPERATION
    TrusteeForm As TRUSTEE_FORM
    TrusteeType As TRUSTEE_TYPE
    ptstrName As String
End Type

Public Type EXPLICIT_ACCESS
    grfAccessPermissions As Long
    grfAccessMode As ACCESS_MODE
    grfInheritance As Long
    TRUSTEE As TRUSTEE
End Type

Public Declare Function ZwOpenSection Lib "NTDLL.DLL" (SectionHandle As Long, ByVal DesiredAccess As Long, ObjectAttributes As Any) As Long
'Public Declare Function LocalFree Lib "kernel32" (ByVal hMem As Any) As Long
Public Declare Function MapViewOfFile Lib "kernel32" (ByVal hFileMappingObject As Long, ByVal dwDesiredAccess As Long, ByVal dwFileOffsetHigh As Long, ByVal dwFileOffsetLow As Long, ByVal dwNumberOfBytesToMap As Long) As Long
Public Declare Function UnmapViewOfFile Lib "kernel32" (lpBaseAddress As Any) As Long
Public Type LIST_ENTRY
    Blink                           As Long
    Flink                           As Long
End Type

Public Type OSVERSIONINFO
    dwOSVersionInfoSize             As Long
    dwMajorVersion                  As Long
    dwMinorVersion                  As Long
    dwBuildNumber                   As Long
    dwPlatformId                    As Long
    szCSDVersion                    As String * 128
End Type

Public Type UNICODE_STRING
    Length                          As Integer
    MaximumLength                   As Integer
    buffer                          As Long
End Type

Public Type INT64
    dwLow                           As Long
    dwHigh                          As Long
End Type
Public Const TH32CS_SNAPHEAPLIST = &H1
Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPTHREAD = &H4
Public Const TH32CS_SNAPMODULE = &H8
Public Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Public Const TH32CS_INHERIT = &H80000000

Public Enum SE_OBJECT_TYPE
    SE_UNKNOWN_OBJECT_TYPE = 0
    SE_FILE_OBJECT
    SE_SERVICE
    SE_PRINTER
    SE_REGISTRY_KEY
    SE_LMSHARE
    SE_KERNEL_OBJECT
    SE_WINDOW_OBJECT
    SE_DS_OBJECT
    SE_DS_OBJECT_ALL
    SE_PROVIDER_DEFINED_OBJECT
    SE_WMIGUID_OBJECT
End Enum


Public Enum SYSTEM_INFORMATION_CLASS
      SystemBasicInformation
      SystemProcessorInformation           '// obsolete...delete
      SystemPerformanceInformation
      SystemTimeOfDayInformation
      SystemPathInformation
      SystemProcessInformation
      SystemCallCountInformation
      SystemDeviceInformation
      SystemProcessorPerformanceInformation
      SystemFlagsInformation
      SystemCallTimeInformation
      SystemModuleInformation
      SystemLocksInformation
      SystemStackTraceInformation
      SystemPagedPoolInformation
      SystemNonPagedPoolInformation
      SystemHandleInformation
      SystemObjectInformation
      SystemPagefileInformation
      SystemVdmInstemulInformation
      SystemVdmBopInformation
      SystemFileCacheInformation
      SystemPoolTagInformation
      SystemInterruptInformation
      SystemDpcBehaviorInformation
      SystemFullMemoryInformation
      SystemLoadGdiDriverInformation
      SystemUnloadGdiDriverInformation
      SystemTimeAdjustmentInformation
      SystemSummaryMemoryInformation
      SystemMirrorMemoryInformation
      SystemPerformanceTraceInformation
      SystemObsolete0
      SystemExceptionInformation
      SystemCrashDumpStateInformation
      SystemKernelDebuggerInformation
      SystemContextSwitchInformation
      SystemRegistryQuotaInformation
      SystemExtendServiceTableInformation
      SystemPrioritySeperation
      SystemVerifierAddDriverInformation
      SystemVerifierRemoveDriverInformation
      SystemProcessorIdleInformation
      SystemLegacyDriverInformation
      SystemCurrentTimeZoneInformation
      SystemLookasideInformation
      SystemTimeSlipNotification
      SystemSessionCreate
      SystemSessionDetach
      SystemSessionInformation
      SystemRangeStartInformation
      SystemVerifierInformation
      SystemVerifierThunkExtend
      SystemSessionProcessInformation
      SystemLoadGdiDriverInSystemSpace
      SystemNumaProcessorMap
      SystemPrefetcherInformation
      SystemExtendedProcessInformation
      SystemRecommendedSharedDataAlignment
      SystemComPlusPackage
      SystemNumaAvailableMemory
      SystemProcessorPowerInformation
      SystemEmulationBasicInformation
      SystemEmulationProcessorInformation
      SystemExtendedHandleInformation
      SystemLostDelayedWriteInformation
      SystemBigPoolInformation
      SystemSessionPoolTagInformation
      SystemSessionMappedViewInformation
      SystemHotpatchInformation
      SystemObjectSecurityMode
      SystemWatchdogTimerHandler
      SystemWatchdogTimerInformation
      SystemLogicalProcessorInformation
      SystemWow64SharedInformation
      SystemRegisterFirmwareTableInformationHandler
      SystemFirmwareTableInformation
      SystemModuleInformationEx
      SystemVerifierTriageInformation
      SystemSuperfetchInformation
      SystemMemoryListInformation
      SystemFileCacheInformationEx
      MaxSystemInfoClass    '// MaxSystemInfoClass should always be the last enum
End Enum

Public Type SYSTEM_BASIC_INFORMATION
    Reserved1(1 To 24) As Byte
    Reserved2(1 To 4) As Long 'PVOID
    NumberOfProcessors As Byte
End Type

Public Enum SYSDBG_COMMAND
    '//以下5个在Windows NT各个版本上都有
    SysDbgGetTraceInformation = 1
    SysDbgSetInternalBreakpoint = 2
    SysDbgSetSpecialCall = 3
    SysDbgClearSpecialCalls = 4
    SysDbgQuerySpecialCalls = 5
    '// 以下是NT 5.1 新增的
    SysDbgDbgBreakPointWithStatus = 6
    '//获取KdVersionBlock
    SysDbgSysGetVersion = 7
    '//从内核空间拷贝到用户空间或者从用户空间拷贝到用户空间
    '//但是不能从用户空间拷贝到内核空间
    SysDbgCopyMemoryChunks_0 = 8
    '//SysDbgReadVirtualMemory = 8
    '//从用户空间拷贝到内核空间或者从用户空间拷贝到用户空间
    '//但是不能从内核空间拷贝到用户空间
    SysDbgCopyMemoryChunks_1 = 9
    '//SysDbgWriteVirtualMemory = 9
    '//从物理地址拷贝到用户空间 不能写到内核空间
    SysDbgCopyMemoryChunks_2 = 10
    '//SysDbgReadVirtualMemory = 10
    '//从用户空间拷贝到物理地址 不能读取内核空间
    SysDbgCopyMemoryChunks_3 = 11
    '//SysDbgWriteVirtualMemory = 11
    '//读写处理器相关控制块
    SysDbgSysReadControlSpace = 12
    SysDbgSysWriteControlSpace = 13
    '//读写端口
    SysDbgSysReadIoSpace = 14
    SysDbgSysWriteIoSpace = 15
    '//分别调用RDMSR@4和_WRMSR@12
    SysDbgSysReadMsr = 16
    SysDbgSysWriteMsr = 17
    '//读写总线数据
    SysDbgSysReadBusData = 18
    SysDbgSysWriteBusData = 19
    SysDbgSysCheckLowMemory = 20
    '// 以下是NT 5.2 新增的
    '//分别调用_KdEnableDebugger@0和_KdDisableDebugger@0
    SysDbgEnableDebugger = 21
    SysDbgDisableDebugger = 22
    '//获取和设置一些调试相关的变量
    SysDbgGetAutoEnableOnEvent = 23
    SysDbgSetAutoEnableOnEvent = 24
    SysDbgGetPitchDebugger = 25
    SysDbgSetDbgPrintBufferSize = 26
    SysDbgGetIgnoreUmExceptions = 27
    SysDbgSetIgnoreUmExceptions = 28
End Enum


Public Type CLIENT_ID
    UniqueProcess As Long
    UniqueThread  As Long
End Type

Public Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type
Public Function ReadKernelMemory(ByVal Addr As Long, ByVal buffer As Long, ByVal size As Long, ReturnLength As Long) As Boolean
    Dim QueryBuff As MEMORY_CHUNKS
    
    With QueryBuff
        .address = Addr
        .pData = buffer
        .Length = size
    End With
    
    ZwSystemDebugControl 8, VarPtr(QueryBuff), Len(QueryBuff), 0, 0, ReturnLength
    ReadKernelMemory = ReturnLength = size
End Function

Public Function WriteKernelMemory(ByVal Addr As Long, ByVal buffer As Long, ByVal size As Long, ReturnLength As Long) As Boolean
    Dim QueryBuff As MEMORY_CHUNKS
    
    With QueryBuff
        .address = Addr
        .pData = buffer
        .Length = size
    End With
    
    ZwSystemDebugControl 9, VarPtr(QueryBuff), Len(QueryBuff), 0, 0, ReturnLength
    WriteKernelMemory = ReturnLength = size
End Function

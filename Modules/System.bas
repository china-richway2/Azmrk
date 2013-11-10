Attribute VB_Name = "System"
Option Explicit
Public Declare Function CreateToolhelp32Snapshot Lib "kernel32" (ByVal dwFlags As Long, ByVal th32ProcessID As Long) As Long
Public Declare Function ZwQuerySystemInformation Lib "NTDLL.DLL" (ByVal SystemInformationClass As SYSTEM_INFORMATION_CLASS, ByVal pSystemInformation As Long, ByVal SystemInformationLength As Long, ByRef ReturnLength As Long) As Long
Public Declare Function ZwSystemDebugControl Lib "NTDLL.DLL" (ByVal scCommand As SYSDBG_COMMAND, ByVal pInputBuffer As Long, ByVal InputBufferLength As Long, ByVal pOutputBuffer As Long, ByVal OutputBufferLength As Long, ByRef pReturnLength As Long) As Long


Public Const STATUS_INFO_LENGTH_MISMATCH = &HC0000004

Public Const TH32CS_SNAPHEAPLIST = &H1
Public Const TH32CS_SNAPPROCESS = &H2
Public Const TH32CS_SNAPTHREAD = &H4
Public Const TH32CS_SNAPMODULE = &H8
Public Const TH32CS_SNAPALL = (TH32CS_SNAPHEAPLIST Or TH32CS_SNAPPROCESS Or TH32CS_SNAPTHREAD Or TH32CS_SNAPMODULE)
Public Const TH32CS_INHERIT = &H80000000


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
      SystemPageFileInformation
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
        .Address = Addr
        .pData = buffer
        .Length = size
    End With
    
    ZwSystemDebugControl 8, VarPtr(QueryBuff), Len(QueryBuff), 0, 0, ReturnLength
    ReadKernelMemory = ReturnLength = size
End Function

Public Function WriteKernelMemory(ByVal Addr As Long, ByVal buffer As Long, ByVal size As Long, ReturnLength As Long) As Boolean
    Dim QueryBuff As MEMORY_CHUNKS
    
    With QueryBuff
        .Address = Addr
        .pData = buffer
        .Length = size
    End With
    
    ZwSystemDebugControl 9, VarPtr(QueryBuff), Len(QueryBuff), 0, 0, ReturnLength
    WriteKernelMemory = ReturnLength = size
End Function

Attribute VB_Name = "Memory"
Option Explicit
Public Declare Function GetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, lpMinimumWorkingSetSize As Long, lpMaximumWorkingSetSize As Long) As Long
Public Declare Function GetProcessMemoryInfo Lib "psapi.dll" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Public Declare Function GetMappedFileName Lib "psapi.dll" Alias "GetMappedFileNameA" (ByVal hProcess As Long, ByVal lpv As Long, lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function VirtualQuery Lib "kernel32" (ByVal lpAddress As Long, ByRef lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Public Declare Function VirtualQueryEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByRef lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Public Declare Function ZwAllocateVirtualMemory Lib "ntdll" (ByVal ProcessHandle As Long, ByRef BaseAddress As Long, ByVal ZeroBits As Long, ByRef RegionSize As Long, ByVal AllocationType As Long, ByVal Protect As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReadProcessLongMemory Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function WriteProcessLongMemory Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)


Public Const MEM_FREE = &H10000
Public Const MEM_PRIVATE = &H20000
Public Const MEM_COMMIT = 4096
Public Const MEM_RESERVE = &H2000
Public Const MEM_DECOMMIT = &H4000
Public Const MEM_RELEASE = &H8000

Public Const PAGE_READONLY = &H2
Public Const PAGE_READWRITE = &H4
Public Const PAGE_EXECUTE_READWRITE = &H40


Public Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type

Public Type MEMORY_CHUNKS
    Address As Long
    pData As Long
    Length As Long
End Type

Public Type MEMORY_BASIC_INFORMATION
    BaseAddress As Long       '// 区域基地址。
    AllocationBase As Long    '// 分配基地址。
    AllocationProtect As Long '// 区域被初次保留时赋予的保护属性。
    RegionSize As Long        '// 区域大小（以字节为计量单位）。
    State As Long             '// 状态（MEM_FREE、MEM_RESERVE或 MEM_COMMIT）。
    Protect As Long           '// 保护属性。
    Type As Long              '// 类型。
End Type

Public Type MEMORYSTATUS
    dwLength As Long
    dwMemoryLoad As Long
    dwTotalPhys As Long
    dwAvailPhys As Long
    dwTotalPageFile As Long
    dwAvailPageFile As Long
    dwTotalVirtual As Long
    dwAvailVirtual As Long
End Type

Public Function FxGetProcessMemoryInformation(ByVal hProcess As Long) As String
    Dim pmc As PROCESS_MEMORY_COUNTERS
        
    GetProcessMemoryInfo hProcess, pmc, LenB(pmc)
    'pmc.WorkingSetSize = ByteToMillionByte(pmc.WorkingSetSize)
    'pmc.PeakWorkingSetSize = ByteToMillionByte(pmc.PeakWorkingSetSize)
    FxGetProcessMemoryInformation = ByteToKMG(pmc.WorkingSetSize) & " - " & ByteToKMG(pmc.PeakWorkingSetSize)
End Function


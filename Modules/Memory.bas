Attribute VB_Name = "Memory"
Option Explicit
Public Declare Function GetProcessWorkingSetSize Lib "kernel32" (ByVal hProcess As Long, lpMinimumWorkingSetSize As Long, lpMaximumWorkingSetSize As Long) As Long
Public Declare Function GetProcessMemoryInfo Lib "psapi.dll" (ByVal hProcess As Long, ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long
Public Declare Function GetMappedFileName Lib "psapi.dll" Alias "GetMappedFileNameA" (ByVal hProcess As Long, ByVal lpv As Long, lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Public Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Public Declare Function VirtualQuery Lib "kernel32" (ByVal lpAddress As Long, ByRef lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Public Declare Function VirtualQueryEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByRef lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Public Declare Function VirtualProtect Lib "kernel32" (lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Public Declare Function VirtualProtectEx Lib "kernel32" (ByVal hProcess As Long, lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, lpflOldProtect As Long) As Long
Public Declare Function ZwAllocateVirtualMemory Lib "ntdll" (ByVal ProcessHandle As Long, ByRef BaseAddress As Long, ByVal ZeroBits As Long, ByRef RegionSize As Long, ByVal AllocationType As Long, ByVal Protect As Long) As Long
Public Declare Function ZwReadVirtualMemory Lib "ntdll" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function ZwWriteVirtualMemory Lib "ntdll" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function ReadProcessLongMemory Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Function WriteProcessLongMemory Lib "kernel32" Alias "WriteProcessMemory" (ByVal hProcess As Long, ByVal lpBaseAddress As Long, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Public Declare Sub CopyMemory Lib "KERNEL32.DLL" Alias "RtlMoveMemory" (ByVal Destination As Long, ByVal Source As Long, ByVal Length As Long)
Public Declare Sub ZeroMemory Lib "kernel32" Alias "RtlZeroMemory" (dest As Any, ByVal numBytes As Long)
Public Declare Function LocalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function IsBadReadPtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Public Declare Function IsBadWritePtr Lib "kernel32" (ByVal lp As Long, ByVal ucb As Long) As Long
Public Declare Function GlobalAddAtom Lib "kernel32" Alias "GlobalAddAtomA" (ByVal lpString As String) As Integer
Public Declare Function GlobalAlloc Lib "kernel32" (ByVal wFlags As Long, ByVal dwBytes As Long) As Long
Public Declare Function GlobalCompact Lib "kernel32" (ByVal dwMinFree As Long) As Long
Public Declare Function GlobalDeleteAtom Lib "kernel32" (ByVal nAtom As Integer) As Integer
Public Declare Function GlobalFindAtom Lib "kernel32" Alias "GlobalFindAtomA" (ByVal lpString As String) As Integer

Public Declare Function GetProcessHeap Lib "KERNEL32.DLL" () As Long
Public Declare Function RtlAllocateHeap Lib "ntdll" (ByVal HeapHandle As Long, ByVal Flags As Long, ByVal Size As Long) As Long
Public Declare Function RtlFreeHeap Lib "ntdll" (ByVal HeapHandle As Long, ByVal Flags As Long, ByVal HeapBase As Long) As Boolean

Public Declare Sub GlobalMemoryStatus Lib "kernel32" (lpBuffer As MEMORYSTATUS)
Public Declare Function GlobalGetAtomName Lib "kernel32" Alias "GlobalGetAtomNameA" (ByVal nAtom As Integer, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Sub GlobalFix Lib "kernel32" (ByVal hMem As Long)
Public Declare Function GlobalFlags Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalFree Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalHandle Lib "kernel32" (wMem As Any) As Long
Public Declare Function GlobalLock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalReAlloc Lib "kernel32" (ByVal hMem As Long, ByVal dwBytes As Long, ByVal wFlags As Long) As Long
Public Declare Function GlobalSize Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Sub GlobalUnfix Lib "kernel32" (ByVal hMem As Long)
Public Declare Function GlobalUnlock Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalUnWire Lib "kernel32" (ByVal hMem As Long) As Long
Public Declare Function GlobalWire Lib "kernel32" (ByVal hMem As Long) As Long

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

Public Type VM_COUNTERS
    PeakVirtualSize As Long
    VirtualSize As Long
    'cb
    '以下和PROCESS_MEMORY_COUNTERS一样
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

Public Type IO_COUNTERS
    ReadOperationCount As FILETIME
    WriteOperationCount As FILETIME
    OtherOperationCount As FILETIME
    ReadTransferCount As FILETIME
    WriteTransferCount As FILETIME
    OtherTransferCount As FILETIME
End Type


Public Type MEMORY_CHUNKS
    address As Long
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
    FxGetProcessMemoryInformation = ByteToKMG(pmc.WorkingSetSize) & " - " & ByteToKMG(pmc.PeakWorkingSetSize)
End Function

Public Function AllocMemory(ByVal dwSize As Long, Optional ByVal wFlags As Long) As Long
    'CopyMemory VarPtr(AllocMemory), &H7C8853A4, 4
    'AllocMemory = RtlAllocateHeap(AllocMemory, wFlags, dwSize)
    '以上代码和以下代码一样；但是GetProcessHeap并不是读取&H7C8853A4处的内容
    AllocMemory = RtlAllocateHeap(GetProcessHeap, wFlags, dwSize)
End Function

Public Sub FreeMemory(ByVal lpMem As Long, Optional ByVal wFlags As Long)
    Call RtlFreeHeap(GetProcessHeap, wFlags, lpMem)
End Sub

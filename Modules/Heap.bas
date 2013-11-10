Attribute VB_Name = "Heap"
Option Explicit
Public Declare Function Heap32ListFirst Lib "kernel32" (ByVal hSnapshot As Long, ByRef lphe As HEAPLIST32) As Long
Public Declare Function Heap32ListNext Lib "kernel32" (ByVal hSnapshot As Long, ByRef lphe As HEAPLIST32) As Long
Public Declare Function Heap32First Lib "kernel32" (ByVal hSnapshot As Long, ByRef lphe As HEAPENTRY32, ByVal th32ProcessID As Long, ByVal th32HeapID As Long) As Long
Public Declare Function Heap32Next Lib "kernel32" (ByVal hSnapshot As Long, ByRef lphe As HEAPENTRY32) As Long


Public Type HEAPLIST32
    dwSize As Long
    th32ProcessID As Long   '// owning process
    th32HeapID As Long      '// heap (in owning process's context!)
    dwFlags As Long
End Type

Public Type HEAPENTRY32
    dwSize As Long
    hHandle As Long     '// Handle of this heap block
    dwAddress As Long   '// Linear address of start of block
    dwBlockSize As Long '// Size of block in bytes
    dwFlags As Long
    dwLockCount As Long
    dwResvd As Long
    th32ProcessID As Long   '// owning process
    th32HeapID As Long      '// heap block is in
End Type

Public Sub HNNew(ByVal pid As Long)
    Call Toolhelp32HeapListNew(pid)
End Sub

Public Sub Toolhelp32HeapListNew(ByVal pid As Long)
    Dim hSnapshot As Long
    Dim hSnapshot2 As Long
    Dim st1, st2 As Long
    Dim HeapListInfo As HEAPLIST32
    Dim HeapInfo As HEAPENTRY32
    
    hSnapshot = CreateToolhelp32Snapshot(TH32CS_SNAPHEAPLIST, pid)
    HeapListInfo.dwSize = LenB(HeapListInfo)
    
    st1 = Heap32ListFirst(hSnapshot, HeapListInfo)
    
    Do While st1
        hSnapshot2 = CreateToolhelp32Snapshot(TH32CS_SNAPALL, pid)
        HeapInfo.dwSize = LenB(HeapInfo)
        st2 = Heap32First(hSnapshot2, HeapInfo, 0, HeapListInfo.th32HeapID)
        'Do While st2
            HeapList.ListView1.ListItems.Add , , HeapInfo.hHandle
            'st2 = Heap32Next(hSnapshot2, HeapInfo)
        'Loop
        
      
        st1 = Heap32ListNext(hSnapshot, HeapListInfo)
    Loop
End Sub

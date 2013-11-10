Attribute VB_Name = "Object"
Option Explicit
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function WaitForMultipleObjectsEx Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long, ByValbAlertable As Long) As Long
Public Declare Function CloseHandle Lib "KERNEL32.DLL" (ByVal Handle As Long) As Long
Public Declare Function ZwDuplicateObject Lib "NTDLL.DLL" (ByVal SourceProcessHandle As Long, ByVal SourceHandle As Long, ByVal TargetProcessHandle As Long, ByRef TargetHandle As Long, ByVal DesiredAccess As Long, ByVal HandleAttributes As Long, ByVal Options As Long) As Long
Public Declare Function ZwClose Lib "NTDLL.DLL" (ByVal ObjectHandle As Long) As Long
Public Declare Sub DebugBreak Lib "kernel32" ()

Public Const OBJ_CASE_INSENSITIVE = &H40

Public Const WAIT_TIMEOUT = &H102

Public Const INFINITE = &HFFFF      '  Infinite timeout

Public Enum SYSTEM_HANDLE_TYPE
    OB_TYPE_UNKNOWN = 0
    OB_TYPE_TYPE = 1
    OB_TYPE_DIRECTORY = 2
    OB_TYPE_SYMBOLIC_LINK = 3
    OB_TYPE_TOKEN = 4
    OB_TYPE_PROCESS = 5
    OB_TYPE_THREAD = 6
    OB_TYPE_JOB = 7
    OB_TYPE_DEBUG_OBJECT = 8
    OB_TYPE_EVENT = 9
    OB_TYPE_EVENT_PAIR = 10
    OB_TYPE_MUTANT = 11
    OB_TYPE_CALLBACK = 12
    OB_TYPE_SEMAPHORE = 13
    OB_TYPE_TIMER = 14
    OB_TYPE_PROFILE = 15
    OB_TYPE_KEYED_EVENT = 16
    OB_TYPE_WINDOWS_STATION = 17
    OB_TYPE_DESKTOP = 18
    OB_TYPE_SECTION = 19
    OB_TYPE_KEY = 20
    OB_TYPE_PORT = 21
    OB_TYPE_WAITABLE_PORT = 22
    OB_TYPE_ADAPTER = 23
    OB_TYPE_CONTROLLER = 24
    OB_TYPE_DEVICE = 25
    OB_TYPE_DRIVER = 26
    OB_TYPE_IOCOMPLETION = 27
    OB_TYPE_FILE = 28
    OB_TYPE_WMIGUID = 29
End Enum

Public Type SYSTEM_HANDLE_TABLE_ENTRY_INFO
    UniqueProcessId As Integer
    CreatorBackTraceIndex As Integer
    ObjectTypeIndex As Byte
    HandleAttributes As Byte
    HandleValue As Integer
    pObject As Long
    GrantedAccess As Long
End Type

Public Type OBJECT_ATTRIBUTES
    Length As Long
    RootDirectory As Long
    ObjectName As Long
    Attributes As Long
    SecurityDescriptor As Long
    SecurityQualityOfService As Long
End Type

Public NumOfHandle As Long
Public HandleTable() As SYSTEM_HANDLE_TABLE_ENTRY_INFO

Public Sub RdQueryHandleInformation(ByVal mHandle As Long, lpBuffer As SYSTEM_HANDLE_TABLE_ENTRY_INFO, Optional ByVal dwPid As Long)
    Dim mPid As Long, i As Long
    If dwPid = -1 Then
        mPid = GetCurrentProcessId
    ElseIf dwPid > 0 Then
        mPid = dwPid
    End If
    For i = 1 To NumOfHandle
        If HandleTable(i).HandleValue = mHandle Then
            If mPid = 0 Or HandleTable(i).UniqueProcessId = mPid Then
                lpBuffer = HandleTable(i)
                Exit For
            End If
        End If
    Next i
End Sub

Public Sub RefreshHandleTable()
    Dim arySize As Long, bytBuf() As Byte, i As Long, st As Long
    ReDim bytBuf(19)
    st = ZwQuerySystemInformation(SystemHandleInformation, VarPtr(bytBuf(0)), 20, arySize)
    If st <> STATUS_INFO_LENGTH_MISMATCH Then Exit Sub
    ReDim bytBuf(arySize - 1)
Again:
    st = ZwQuerySystemInformation(SystemHandleInformation, VarPtr(bytBuf(0)), arySize, arySize)
    If Not NT_SUCCESS(st) Then
        If st <> STATUS_INFO_LENGTH_MISMATCH Then Exit Sub
        GoTo Again
    End If
    CopyMemory VarPtr(NumOfHandle), VarPtr(bytBuf(0)), 4
    ReDim HandleTable(1 To NumOfHandle)
    CopyMemory VarPtr(HandleTable(1)), VarPtr(bytBuf(4)), NumOfHandle * 16
End Sub


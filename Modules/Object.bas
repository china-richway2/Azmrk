Attribute VB_Name = "Object"
Option Explicit
Public Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function WaitForMultipleObjects Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long) As Long
Public Declare Function WaitForMultipleObjectsEx Lib "kernel32" (ByVal nCount As Long, lpHandles As Long, ByVal bWaitAll As Long, ByVal dwMilliseconds As Long, ByValbAlertable As Long) As Long
Public Declare Function CloseHandle Lib "kernel32.dll" (ByVal Handle As Long) As Long
Public Declare Function ZwDuplicateObject Lib "NTDLL.DLL" (ByVal SourceProcessHandle As Long, ByVal SourceHandle As Long, ByVal TargetProcessHandle As Long, ByRef TargetHandle As Long, ByVal DesiredAccess As Long, ByVal HandleAttributes As Long, ByVal Options As Long) As Long
Public Declare Function ZwCreateDebugObject Lib "NTDLL.DLL" (ByRef pDebugObjectHandle As Long, ByVal DesiredAccess As Long, ByVal pObjectAttributes As Long, ByVal flags As Long) As Long
Public Declare Function ZwClose Lib "NTDLL.DLL" (ByVal ObjectHandle As Long) As Long


Public Const WAIT_TIMEOUT = &H102

Public Const INFINITE = &HFFFF      '  Infinite timeout


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



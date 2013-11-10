Attribute VB_Name = "Privilege"
Option Explicit
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As TOKEN_PRIVILEGES, ByRef pReturnLength As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long


Private Type MEMORY_CHUNKS
        Address As Long
        pData As Long
        Length As Long
End Type

Private Type LUID
        UsedPart As Long
        IgnoredForNowHigh32BitPart As Long
End Type '

Private Type TOKEN_PRIVILEGES
        PrivilegeCount As Long
        TheLuid As LUID
        Attributes As Long
End Type


Public Const SE_CREATE_TOKEN = "SeCreateTokenPrivilege"
Public Const SE_ASSIGNPRIMARYTOKEN = "SeAssignPrimaryTokenPrivilege"
Public Const SE_LOCK_MEMORY = "SeLockMemoryPrivilege"
Public Const SE_INCREASE_QUOTA = "SeIncreaseQuotaPrivilege"
Public Const SE_UNSOLICITED_INPUT = "SeUnsolicitedInputPrivilege"
Public Const SE_MACHINE_ACCOUNT = "SeMachineAccountPrivilege"
Public Const SE_TCB = "SeTcbPrivilege"
Public Const SE_SECURITY = "SeSecurityPrivilege"
Public Const SE_TAKE_OWNERSHIP = "SeTakeOwnershipPrivilege"
Public Const SE_LOAD_DRIVER = "SeLoadDriverPrivilege"
Public Const SE_SYSTEM_PROFILE = "SeSystemProfilePrivilege"
Public Const SE_SYSTEMTIME = "SeSystemtimePrivilege"
Public Const SE_PROF_SINGLE_PROCESS = "SeProfileSingleProcessPrivilege"
Public Const SE_INC_BASE_PRIORITY = "SeIncreaseBasePriorityPrivilege"
Public Const SE_CREATE_PAGEFILE = "SeCreatePagefilePrivilege"
Public Const SE_CREATE_PERMANENT = "SeCreatePermanentPrivilege"
Public Const SE_BACKUP = "SeBackupPrivilege"
Public Const SE_RESTORE = "SeRestorePrivilege"
Public Const SE_SHUTDOWN = "SeShutdownPrivilege"
Public Const SE_DEBUG = "SeDebugPrivilege"
Public Const SE_AUDIT = "SeAuditPrivilege"
Public Const SE_SYSTEM_ENVIRONMENT = "SeSystemEnvironmentPrivilege"
Public Const SE_CHANGE_NOTIFY = "SeChangeNotifyPrivilege"
Public Const SE_REMOTE_SHUTDOWN = "SeRemoteShutdownPrivilege"

Private Const SE_PRIVILEGE_ENABLED As Long = &H2
Private Const TOKEN_QUERY As Long = &H8
Private Const TOKEN_ADJUST_PRIVILEGES As Long = &H20

Public Function EnablePrivilege(ByVal seName As String) As Boolean
        On Error Resume Next
        Dim p_lngRtn As Long
        Dim p_lngToken As Long
        Dim p_lngBufferLen As Long
        Dim p_typLUID As LUID
        Dim p_typTokenPriv As TOKEN_PRIVILEGES
        Dim p_typPrevTokenPriv As TOKEN_PRIVILEGES
        p_lngRtn = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, p_lngToken)

        If p_lngRtn = 0 Then
                EnablePrivilege = False
                Exit Function
        End If

        If Err.LastDllError <> 0 Then
                EnablePrivilege = False
                Exit Function
        End If

        p_lngRtn = LookupPrivilegeValue(0&, seName, p_typLUID)

        If p_lngRtn = 0 Then
                EnablePrivilege = False
                Exit Function
        End If

        p_typTokenPriv.PrivilegeCount = 1
        p_typTokenPriv.Attributes = SE_PRIVILEGE_ENABLED
        p_typTokenPriv.TheLuid = p_typLUID
        EnablePrivilege = (AdjustTokenPrivileges(p_lngToken, False, p_typTokenPriv, Len(p_typPrevTokenPriv), p_typPrevTokenPriv, p_lngBufferLen) <> 0)
End Function


Attribute VB_Name = "Token"
Option Explicit
Public Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, ByRef NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, ByRef PreviousState As TOKEN_PRIVILEGES, ByRef pReturnLength As Long) As Long
Public Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpLuid As LUID) As Long
Public Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Public Declare Function ConvertSidToStringSidW Lib "advapi32.dll" (ByVal Sid As Long, StringSid As Any) As Long
Public Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As TokenInformationClass, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Public Const TOKEN_QUERY As Long = &H8
Public Enum TokenInformationClass
    TokenUser = 1
    TokenGroups = 2
    TokenPrivileges = 3
    TokenOwner = 4
    TokenPrimaryGroup = 5
    TokenDefaultDacl = 6
    TokenSource = 7
    TokenType = 8
    TokenImpersonationLevel = 9
    TokenStatistics = 10
End Enum
Public Type SID_AND_ATTRIBUTES
    Sid As Long
    Attributes As Long
End Type
Public Type TOKEN_USER
    SidAndAttributes As SID_AND_ATTRIBUTES
    Fill(1 To 28) As Byte
End Type

Private Type MEMORY_CHUNKS
        address As Long
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

Public Function GetUserName() As String
    Dim hToken As Long
    If OpenProcessToken(GetCurrentProcess, TOKEN_QUERY, hToken) = 0 Or hToken = 0 Then
        CloseHandle hToken
        Exit Function
    End If
    Dim tkInf As TOKEN_USER
    If GetTokenInformation(hToken, TokenUser, tkInf, Len(tkInf), 0) = 0 Then
        CloseHandle hToken
        Exit Function
    End If
    Dim S As Long
    ConvertSidToStringSidW tkInf.SidAndAttributes.Sid, S
    Dim pLen As Long
    pLen = lstrlenW(S) * 2
    Dim pPtr() As Byte
    ReDim pPtr(1 To pLen)
    CopyMemory VarPtr(pPtr(1)), ByVal S, pLen
    LocalFree S
    GetUserName = pPtr
    CloseHandle hToken
End Function


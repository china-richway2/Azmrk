VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsRegistry"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private Const ERROR_SUCCESS = 0
Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_OPTION_NON_VOLATILE = 0
Private Const REG_SZ = 1
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_ALL = &H1F0000
Private Const KEY_CREATE_LINK = &H20
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const KEY_NOTIFY = &H10
Private Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

Private Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Public Enum HKEYs
    eHKEY_CLASSES_ROOT = &H80000000
    eHKEY_CURRENT_USER = &H80000001
    eHKEY_LOCAL_MACHINE = &H80000002
    eHKEY_USERS = &H80000003
    eHKEY_PERFORMANCE_DATA = &H80000004
    eHKEY_CURRENT_CONFIG = &H80000005
    eHKEY_DYN_DATA = &H80000006
End Enum

Private Declare Function RegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" _
        (ByVal hKey As HKEYs, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, _
        phkResult As Long) As Long
Private Declare Function RegConnectRegistry Lib "advapi32.dll" _
        Alias "RegConnectRegistryA" _
        (ByVal lpMachineName As String, _
        ByVal hKey As Long, _
        phkResult As Long) As Long
Private Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long

'APIs to get/set values in the registry
Private Declare Function RegQueryValueEx Lib "advapi32.dll" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, _
        ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As Any, _
        lpcbData As Long) As Long
Private Declare Function RegQueryValueExString Lib "advapi32.dll" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, _
        ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As String, _
        lpcbData As Long) As Long
Private Declare Function RegSetValueEx Lib "advapi32.dll" _
        Alias "RegSetValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, _
        lpData As Long, ByVal cbData As Long) As Long
Private Declare Function RegSetValueExString Lib "advapi32.dll" _
        Alias "RegSetValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, _
        ByVal lpData As String, ByVal cbData As Long) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" _
        Alias "RegDeleteValueA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String) As Long

Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" ( _
    ByVal hKey As Long, _
    ByVal dwIndex As Long, _
    ByVal lpName As String, _
    lpcbName As Long, _
    ByVal lpReserved As Long, _
    ByVal lpClass As String, _
    lpcbClass As Long, _
    lpftLastWriteTime As FILETIME) As Long
Private Declare Function RegEnumValue Lib "advapi32.dll" _
        Alias "RegEnumValueA" _
        (ByVal hKey As Long, _
        ByVal dwIndex As Long, _
        ByVal lpValueName As String, _
        lpcbValueName As Long, _
        ByVal lpReserved As Long, _
        lpType As Long, _
        lpData As Byte, _
        lpcbData As Long) As Long
        
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" _
        Alias "RegCreateKeyExA" _
        (ByVal hKey As Long, _
        ByVal lpSubKey As String, _
        ByVal Reserved As Long, _
        ByVal lpClass As String, _
        ByVal dwOptions As Long, _
        ByVal samDesired As Long, _
        lpSecurityAttributes As SECURITY_ATTRIBUTES, _
        phkResult As Long, _
        lpdwDisposition As Long) As Long
Private Declare Function RegDeleteKey Lib "advapi32.dll" _
        Alias "RegDeleteKeyA" _
        (ByVal hKey As Long, ByVal lpSubKey As String) As Long






'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'public wrappers for registry APIs
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Public Function SetValue(ByVal PredefinedKey As HKEYs, ByVal KeyName As String, ByVal ValueName As String, ByVal Value As Variant, Optional ComputerName As String) As Boolean
    On Error GoTo ErrHand
    
    Dim GetHandle As Long
    Dim hKey As Long
    Dim lpType As Long
    Dim lpcbData As Long
    Dim lpData As String
    Dim lReturnCode As Long
    Dim lhRemoteRegistry As Long
  
    If left$(KeyName, 1) = "\" Then
        KeyName = right$(KeyName, Len(KeyName) - 1)
    End If
    
    If ComputerName = "" Then
        GetHandle = RegOpenKeyEx(PredefinedKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
    Else
        lReturnCode = RegConnectRegistry(ComputerName, PredefinedKey, lhRemoteRegistry)
        GetHandle = RegOpenKeyEx(lhRemoteRegistry, KeyName, 0, KEY_ALL_ACCESS, hKey)
    End If
    
    If GetHandle = ERROR_SUCCESS Then
        lpcbData = 255
        lpData = String(lpcbData, Chr(0))
        
        GetHandle = RegQueryValueEx(hKey, ValueName, _
            0, lpType, ByVal lpData, lpcbData)
        
        If GetHandle = ERROR_SUCCESS Then
            Select Case lpType
                Case REG_SZ
                    GetHandle = RegSetValueExString(hKey, ValueName, 0, REG_SZ, CStr(Value), Len(Value) + 1)
                Case REG_DWORD
                    GetHandle = RegSetValueEx(hKey, ValueName, 0, REG_DWORD, ByVal CLng(Value), lpcbData)
                Case REG_BINARY
                    GetHandle = RegSetValueExString(hKey, ValueName, 0, REG_BINARY, CByte(Value), Len(Value) + 1)
            End Select
        Else
            GetHandle = RegSetValueExString(hKey, ValueName, 0, REG_SZ, ByVal CStr(Value), Len(CStr(Value)) + 1)
        End If
    Else:
        GoTo ErrHand
        RegCloseKey hKey
    End If
    
    SetValue = True
    
    Exit Function

ErrHand:
    Err.Raise "11001", "clsRegistry", "SetValue"
End Function

Public Function GetValue(ByVal PredefinedKey As HKEYs, _
                         ByVal KeyName As String, _
                         ByVal ValueName As String, _
                         Optional ComputerName As String) As Variant
    On Error GoTo ErrHand
    
    Dim GetHandle        As Long
    Dim hKey             As Long
    Dim lpData           As String
    Dim lpDataDWORD      As Long
    Dim lpcbData         As Long
    Dim lpType           As Long
    Dim lReturnCode      As Long
    Dim lhRemoteRegistry As Long
    
    If left$(KeyName, 1) = "\" Then
        KeyName = right$(KeyName, Len(KeyName) - 1)
    End If
    
    If ComputerName = "" Then
        GetHandle = RegOpenKeyEx(PredefinedKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
    Else
        lReturnCode = RegConnectRegistry(ComputerName, PredefinedKey, lhRemoteRegistry)
        GetHandle = RegOpenKeyEx(lhRemoteRegistry, KeyName, 0, KEY_ALL_ACCESS, hKey)
    End If
            
    If GetHandle = ERROR_SUCCESS Then
        lpcbData = 255
        lpData = String(lpcbData, Chr(0))
        
        GetHandle = RegQueryValueEx(hKey, ValueName, 0, lpType, ByVal lpData, lpcbData)
            
        If GetHandle = ERROR_SUCCESS Then

            Select Case lpType

                Case REG_SZ
                    GetHandle = RegQueryValueExString(hKey, ValueName, 0, lpType, ByVal lpData, lpcbData)
                
                    If GetHandle = 0 Then
                        GetValue = left$(lpData, lpcbData - 1)
                    Else
                        GetValue = ""
                    End If
                    
                Case REG_DWORD
                    GetHandle = RegQueryValueEx(hKey, ValueName, 0, lpType, lpDataDWORD, lpcbData)
                    
                    If GetHandle = 0 Then
                        GetValue = CLng(lpDataDWORD)
                    Else
                        GetValue = 0
                    End If

                Case REG_BINARY
                    GetHandle = RegQueryValueEx(hKey, ValueName, 0, lpType, lpDataDWORD, lpcbData)
                    
                    If GetHandle = 0 Then
                        GetValue = CByte(lpDataDWORD)
                    Else
                        GetValue = 0
                    End If

                Case 2 'Ì«¼ÙÁË
                    GetHandle = RegQueryValueExString(hKey, ValueName, 0, lpType, ByVal lpData, lpcbData)
                
                    If GetHandle = 0 Then
                        GetValue = left$(lpData, lpcbData - 1)
                    Else
                        GetValue = ""
                    End If

            End Select

        End If
        
        RegCloseKey hKey
    End If
    
    Exit Function

ErrHand:
    Err.Raise "11002", "clsRegistry", "GetValue"
End Function

Public Function ListSubKey(ByVal PredefinedKey As HKEYs, ByVal KeyName As String, ByVal Index As Long, Optional ComputerName As String) As String
    On Error GoTo ErrHand
    
    Dim GetHandle As Long
    Dim hKey As Long
    Dim dwIndex As Long
    Dim lpName As String
    Dim lpcbName As Long
    Dim lpReserved As Long
    Dim lpftLastWriteTime As FILETIME
    Dim i As Integer
    Dim lReturnCode As Long
    Dim lhRemoteRegistry As Long
    If left$(KeyName, 1) = "\" Then
        KeyName = right$(KeyName, Len(KeyName) - 1)
    End If
    
    If ComputerName = "" Then
        GetHandle = RegOpenKeyEx(PredefinedKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
    Else
        lReturnCode = RegConnectRegistry(ComputerName, PredefinedKey, lhRemoteRegistry)
        GetHandle = RegOpenKeyEx(lhRemoteRegistry, KeyName, 0, KEY_ALL_ACCESS, hKey)
    End If
            
    If GetHandle = ERROR_SUCCESS Then
        lpcbName = 255: lpName = String$(lpcbName, Chr(0))
                    
        GetHandle = RegEnumKeyEx(hKey, Index, lpName, lpcbName, lpReserved, vbNullString, 0, lpftLastWriteTime)
                
        If GetHandle = ERROR_SUCCESS Then
            ListSubKey = left$(lpName, lpcbName)
        Else
            ListSubKey = ""
        End If
            
        RegCloseKey hKey
    End If
    
    Exit Function

ErrHand:
    Err.Raise "11003", "clsRegistry", "ListSubKey"
End Function


Public Function DeleteValue(ByVal PredefinedKey As HKEYs, ByVal KeyName As String, ByVal ValueName As String, Optional ComputerName As String) As Boolean
    On Error GoTo ErrHand
    
    Dim GetHandle As Long
    Dim hKey As Long
    Dim lReturnCode As Long
    Dim lhRemoteRegistry As Long
    
    If left$(KeyName, 1) = "\" Then
        KeyName = right$(KeyName, Len(KeyName) - 1)
    End If
    
    If ComputerName = "" Then
        GetHandle = RegOpenKeyEx(PredefinedKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
    Else
        lReturnCode = RegConnectRegistry(ComputerName, PredefinedKey, lhRemoteRegistry)
        GetHandle = RegOpenKeyEx(lhRemoteRegistry, KeyName, 0, KEY_ALL_ACCESS, hKey)
    End If

    If GetHandle = ERROR_SUCCESS Then
        GetHandle = RegDeleteValue(hKey, ValueName)
        If GetHandle = ERROR_SUCCESS Then
            DeleteValue = True
        Else:
            DeleteValue = False
        End If
    End If
    
    Exit Function
ErrHand:
    Err.Raise "11004", "clsRegistry", "DeleteValue"
End Function

Public Function CreateKey(ByVal PredefinedKey As HKEYs, ByVal KeyName As String, Optional ComputerName As String) As Boolean
    On Error GoTo ErrHand
    
    Dim hNewKey As Long
    Dim lpSecurityAttributes As SECURITY_ATTRIBUTES
    Dim GetHandle As Long
    Dim lReturnCode As Long
    Dim lhRemoteRegistry As Long

    If left$(KeyName, 1) = "\" Then
        KeyName = right$(KeyName, Len(KeyName) - 1)
    End If
    
    If ComputerName = "" Then
        GetHandle = RegCreateKeyEx(PredefinedKey, KeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpSecurityAttributes, hNewKey, GetHandle)
    Else
        lReturnCode = RegConnectRegistry(ComputerName, PredefinedKey, lhRemoteRegistry)
        GetHandle = RegCreateKeyEx(lhRemoteRegistry, KeyName, 0&, vbNullString, REG_OPTION_NON_VOLATILE, KEY_ALL_ACCESS, lpSecurityAttributes, hNewKey, GetHandle)
    End If

    If GetHandle = ERROR_SUCCESS Then
        GetHandle = RegCloseKey(hNewKey)
    
        CreateKey = True
    Else
        CreateKey = False
    End If
    
    Exit Function

ErrHand:
    Err.Raise "11005", "clsRegistry", "CreateKey"
End Function

Public Function DeleteKey(ByVal PredefinedKey As HKEYs, ByVal KeyName As String, Optional ComputerName As String) As Boolean
    On Error GoTo ErrHand
    
    Dim GetHandle As Long
    Dim lReturnCode As Long
    Dim lhRemoteRegistry As Long
    
    If left$(KeyName, 1) = "\" Then
        KeyName = right$(KeyName, Len(KeyName) - 1)
    End If
    
    If ComputerName = "" Then
        GetHandle = RegDeleteKey(PredefinedKey, KeyName)
    Else
        lReturnCode = RegConnectRegistry(ComputerName, PredefinedKey, lhRemoteRegistry)
        GetHandle = RegDeleteKey(lhRemoteRegistry, KeyName)
    End If

    If GetHandle = ERROR_SUCCESS Then
        DeleteKey = True
    Else
        DeleteKey = False
    End If
    
    Exit Function

ErrHand:
    Err.Raise "11006", "clsRegistry", "DeleteKey"
End Function

Public Function ListSubValue(ByVal PredefinedKey As HKEYs, ByVal KeyName As String, ByVal Index As Long, Optional ComputerName As String) As String
    On Error GoTo ErrHand

    Dim GetHandle As Long
    Dim hKey As Long
    Dim dwIndex As Long
    Dim lpName As String
    Dim lpcbName As Long
    Dim lpReserved As Long
    Dim lpftLastWriteTime As FILETIME
    Dim i As Integer
    Dim lhRemoteRegistry As Long
    Dim lReturnCode As Long

    If left$(KeyName, 1) = "\" Then
        KeyName = right$(KeyName, Len(KeyName) - 1)
    End If
   
    If ComputerName = "" Then
        GetHandle = RegOpenKeyEx(PredefinedKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
    Else
        lReturnCode = RegConnectRegistry(ComputerName, PredefinedKey, lhRemoteRegistry)
        GetHandle = RegOpenKeyEx(lhRemoteRegistry, KeyName, 0, KEY_ALL_ACCESS, hKey)
    End If
    
            
    If GetHandle = ERROR_SUCCESS Then
        lpcbName = 255: lpName = String$(lpcbName, Chr(0))
        
        GetHandle = RegEnumValue(hKey, Index, lpName, lpcbName, 0, ByVal 0&, ByVal 0&, ByVal 0&)

        If GetHandle = ERROR_SUCCESS Then
            ListSubValue = left(lpName, lpcbName)
        Else
            ListSubValue = ""
        End If

        RegCloseKey hKey
    End If

    Exit Function

ErrHand:
    Err.Raise "11007", "clsRegistry", "ListSubValue"
End Function

Public Function KeyExists(ByVal PredefinedKey As HKEYs, ByVal KeyName As String, Optional ComputerName As String) As Boolean
    On Error GoTo ErrHand

    Dim hKey As Long
    Dim GetHandle As Long
    Dim lhRemoteRegistry As Long
    Dim lReturnCode As Long
    
    If left$(KeyName, 1) = "\" Then
        KeyName = right$(KeyName, Len(KeyName) - 1)
    End If
    
    If ComputerName = "" Then
        GetHandle = RegOpenKeyEx(PredefinedKey, KeyName, 0, KEY_ALL_ACCESS, hKey)
    Else
        lReturnCode = RegConnectRegistry(ComputerName, PredefinedKey, lhRemoteRegistry)
        GetHandle = RegOpenKeyEx(lhRemoteRegistry, KeyName, 0, KEY_ALL_ACCESS, hKey)
    End If
    
    If GetHandle = ERROR_SUCCESS Then
        KeyExists = True
    Else
        KeyExists = False
    End If
    
    Exit Function
ErrHand:
    Err.Raise "11008", "clsRegistry", "KeyExists"
End Function

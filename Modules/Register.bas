Attribute VB_Name = "Registry"
Option Explicit
Public Const ERROR_SUCCESS = 0
Public Const ERROR_MORE_DATA = 234&
Public Const ERROR_NO_MORE_ITEMS = 259&
Public Const KEY_CREATE_LINK = &H20
Public Const KEY_CREATE_SUB_KEY = &H4
Public Const KEY_ENUMERATE_SUB_KEYS = &H8
Public Const KEY_NOTIFY = &H10
Public Const KEY_QUERY_VALUE = &H1
Public Const KEY_SET_VALUE = &H2
Public Const KEY_ALL_ACCESS = ((STANDARD_RIGHTS_ALL Or KEY_QUERY_VALUE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY Or KEY_ENUMERATE_SUB_KEYS Or KEY_NOTIFY Or KEY_CREATE_LINK) And (Not SYNCHRONIZE))

#If True Then '以下内容被淘汰；改为调用ntdll函数
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" _
        Alias "RegOpenKeyExA" _
        (ByVal hKey As HKEYs, ByVal lpSubKey As String, _
        ByVal ulOptions As Long, ByVal samDesired As Long, _
        phkResult As Long) As Long
Public Declare Function RegConnectRegistry Lib "advapi32.dll" _
        Alias "RegConnectRegistryA" _
        (ByVal lpMachineName As String, _
        ByVal hKey As Long, _
        phkResult As Long) As Long
Public Declare Function RegCloseKey Lib "advapi32.dll" _
        (ByVal hKey As Long) As Long

'APIs to get/set values in the registry
Public Declare Function RegQueryValueEx Lib "advapi32.dll" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, _
        ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, lpData As Any, _
        lpcbData As Long) As Long
Public Declare Function RegQueryValueExString Lib "advapi32.dll" _
        Alias "RegQueryValueExA" (ByVal hKey As Long, _
        ByVal lpValueName As String, ByVal lpReserved As Long, _
        lpType As Long, ByVal lpData As String, _
        lpcbData As Long) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" _
        Alias "RegSetValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, _
        lpData As Any, ByVal cbData As Long) As Long
Public Declare Function RegSetValueExString Lib "advapi32.dll" _
        Alias "RegSetValueExA" _
        (ByVal hKey As Long, ByVal lpValueName As String, _
        ByVal Reserved As Long, ByVal dwType As Long, _
        ByVal lpData As String, ByVal cbData As Long) As Long
Public Declare Function RegDeleteValue Lib "advapi32.dll" _
        Alias "RegDeleteValueA" _
        (ByVal hKey As Long, _
        ByVal lpValueName As String) As Long

Public Declare Function RegEnumKeyEx Lib "advapi32.dll" _
        Alias "RegEnumKeyExA" _
        (ByVal hKey As Long, ByVal dwIndex As Long, _
        ByVal lpName As String, lpcbName As Long, _
        ByVal lpReserved As Long, ByVal lpClass As String, _
        lpcbClass As Long, _
        lpftLastWriteTime As Any) As Long
Public Declare Function RegEnumValue Lib "advapi32.dll" _
        Alias "RegEnumValueA" _
        (ByVal hKey As Long, _
        ByVal dwIndex As Long, _
        ByVal lpValueName As String, _
        lpcbValueName As Long, _
        ByVal lpReserved As Long, _
        lpType As Long, _
        lpData As Byte, _
        lpcbData As Long) As Long

Public Declare Function RegCreateKeyEx Lib "advapi32.dll" _
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
Public Declare Function RegDeleteKey Lib "advapi32.dll" _
        Alias "RegDeleteKeyA" _
        (ByVal hKey As Long, ByVal lpSubKey As String) As Long
#End If
Public Declare Function ZwOpenKey Lib "ntdll" (KeyHandle As Long, ByVal DesiredAccess As Long, ObjectAttributes As OBJECT_ATTRIBUTES) As Long
Public Declare Function ZwEnumerateKey Lib "ntdll" (ByVal KeyHandle As Long, ByVal Index As Long, ByVal InformationClass As KEY_INFORMATION_CLASS, Information As Any, ByVal Length As Long, ReturnLength As Long) As Long
Public Declare Function ZwEnumerateValueKey Lib "ntdll" (ByVal KeyHandle As Long, ByVal Index As Long, ByVal InformationClass As KEY_VALUE_INFORMATION_CLASS, Information As Any, ByVal Length As Long, ReturnLength As Long) As Long
Public Declare Function ZwQueryKey Lib "ntdll" (ByVal KeyHandle As Long, ByVal KeyInformationClass As KEY_INFORMATION_CLASS, KeyInformation As Any, ByVal Length As Long, ReturnLength As Long) As Long
Public Declare Function ZwQueryValueKey Lib "ntdll" (ByVal KeyHandle As Long, ValueName As UNICODE_STRING, ByVal KeyValueInformationClass As KEY_VALUE_INFORMATION_CLASS, KeyValueInformation As Any, ByVal Length As Long, RetLength As Long) As Long
Public Declare Function ZwCreateKey Lib "ntdll" (KeyHandle As Long, ByVal DesiredAccess As Long, ObjectAttributes As OBJECT_ATTRIBUTES, ByVal TitleIndex As Long, Class As Any, ByVal CreateOptions As Long, Disposition As Long)
Public Enum KEY_INFORMATION_CLASS
    KeyBasicInformation
    KeyNodeInformation
    KeyFullInformation
    KeyNameInformation
    KeyCachedInformation
    KeyFlagsInformation
    KeyVirtualizationInformation
    KeyHandleTagsInformation
    MaxKeyInfoClass
End Enum
Public Type KEY_BASIC_INFORMATION
    LastWriteTime As FILETIME
    TitleIndex As Long
    NameLength As Long
    'Name As String * 1024
End Type
Public Type KEY_NODE_INFORMATION
    LastWriteTime As FILETIME
    TitleIndex As Long
    ClassOffset As Long
    ClassLength As Long
    NameLength As Long
    'Name[] as String
End Type
Public Type KEY_CACHED_INFORMATION
    LastWriteTime As FILETIME
    TitleIndex As Long
    SubKeys As Long
    MaxNameLen As Long
    Values As Long
    MaxValueNameLen As Long
    MaxValueDataLen As Long
    NameLength As Long
End Type
Public Type KEY_FULL_INFORMATION
    LastWriteTime As FILETIME
    TitleIndex As Long
    ClassOffset As Long
    ClassLength As Long
    SubKeys As Long
    MaxNameLen As Long
    MaxClassLen As Long
    Values As Long
    MaxValueNameLen As Long
    MaxValueDataLen As Long
    'Class As String * 1024
End Type
Public Type KEY_VIRTUALIZATION_INFORMATION
    VirtualizationCandidate As Long '1
    VirtualizationEnabled As Long '1
    VirtualTarget As Long '1
    VirtualStore As Long '1
    VirtualSource As Long '1
    Reserved As Long '27
End Type
Public Enum KEY_VALUE_INFORMATION_CLASS
    KeyValueBasicInformation
    KeyValueFullInformation
    KeyValuePartialInformation
    KeyValueFullInformationAlign64
    KeyValuePartialInformationAlign64
    MaxKeyValueInfoClass
End Enum
Public Type KEY_VALUE_FULL_INFORMATION
    TitleIndex As Long
    Type As Long
    DataOffset As Long
    DataLength As Long
    NameLength As Long
End Type

Public Const REG_NONE = 0
Public Const REG_SZ = 1
Public Const REG_EXPAND_SZ = 2
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const REG_DWORD_LITTLE_ENDIAN = 4
Public Const REG_DWORD_BIG_ENDIAN = 5
Public Const REG_LINK = 6
Public Const REG_MULTI_SZ = 7
Public Const oHKEY_LOCAL_MACHINE = "\Registry\Machine"
Public Const oHKEY_CLASSES_ROOT = "\Registry\Machine\SOFTWARE\Classes"
Public Const oHKEY_USERS = "\Registry\User"
Public Const oHKEY_CURRENT_CONFIG = "\Registry\Machine\SYSTEM\CURRENTCONTROLSET\HARDWARE PROFILES\CURRENT"
Public Function oHKEY_CURRENT_USER() As String
    oHKEY_CURRENT_USER = oHKEY_USERS & "\" & GetUserName
End Function
Public Function OpenRegKey(ByVal szPath As String, ByVal Access As Long, ByVal NotQuiet As Boolean) As Long
    Dim s As String
    Dim root As String
    Dim subkey As String
    Dim rootkey As String
    Dim ComputerName As String
    ComputerName = left(szPath, InStr(szPath, "\") - 1)
    s = Mid(szPath, InStr(szPath, "\") + 1)
    If InStr(s, "\") > 0 Then
        root = left(s, InStr(s, "\") - 1)
        subkey = Mid(s, InStr(s, "\") + 1)
    Else
        root = s
    End If
    Select Case root
    Case "HKEY_LOCAL_MACHINE"
        rootkey = oHKEY_LOCAL_MACHINE
    Case "HKEY_CLASSES_ROOT"
        rootkey = oHKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
        rootkey = oHKEY_CURRENT_USER
    Case "HKEY_CURRENT_CONFIG"
        rootkey = oHKEY_CURRENT_CONFIG
    Case "HKEY_USERS"
        rootkey = oHKEY_USERS
    End Select
    'If ComputerName <> "我的电脑" Then
    '    result = RegConnectRegistry(ComputerName, rootkey, Key)
    '    If result = ERROR_SUCCESS Then
    '        result = RegOpenKeyEx(Key, subkey, 0, KEY_ALL_ACCESS, Key)
    '        If result <> ERROR_SUCCESS Then
    '            If NotQuiet Then MsgBox "打开远程注册表项时出错！", vbCritical
    '            Exit Function
    '        End If
    '    Else
    '        If NotQuiet Then MsgBox "打开远程根注册表项时出错！", vbCritical
    '        Exit Function
    '    End If
    'Else
    '    If RegOpenKeyEx(rootkey, subkey, 0, KEY_ALL_ACCESS, Key) <> ERROR_SUCCESS Then
    '        If NotQuiet Then MsgBox "打开注册表项时出错！", vbCritical
    '        Exit Function
    '    End If
    'End If
    'OpenRegKey = Key
    'Exit Function
    Dim oa As OBJECT_ATTRIBUTES, us As UNICODE_STRING
    If subkey = "" Then
        szPath = rootkey
    Else
        szPath = rootkey & "\" & subkey
    End If
    RtlInitUnicodeString us, StrPtr(szPath)
    oa.Length = Len(oa)
    oa.ObjectName = VarPtr(us)
    oa.Attributes = OBJ_CASE_INSENSITIVE
    Call ZwOpenKey(OpenRegKey, Access, oa)
End Function

Public Function GetRegKey(ByVal hKey As Long, Index As Long, szName As String, szClass As String) As Boolean
    Dim st As Long, Buffer() As Byte, Length As Long
    ReDim Buffer(1023)
    st = ZwEnumerateKey(hKey, Index, KeyNodeInformation, Buffer(0), 1024, Length)
    If st = &H8000001A Then
        Exit Function
    End If
    If Length > 1024 Then
        ReDim Buffer(Length - 1)
        st = ZwEnumerateKey(hKey, Index, KeyNodeInformation, Buffer(0), Length, Length)
    End If
    If Not NT_SUCCESS(st) Then
        MsgBox "读取注册表失败！", vbCritical
        Exit Function
    End If
    Dim nInf As KEY_NODE_INFORMATION
    CopyMemory VarPtr(nInf), VarPtr(Buffer(0)), 24
    szName = Space(nInf.NameLength \ 2)
    CopyMemory StrPtr(szName), VarPtr(Buffer(24)), LenB(szName)
    szClass = Space(nInf.ClassLength \ 2)
    If szClass <> "" Then CopyMemory StrPtr(szClass), VarPtr(Buffer(nInf.ClassOffset)), LenB(szClass)
    GetRegKey = True
    Index = Index + 1
End Function

Public Function QueryValueKey(ByVal hKey As Long, Name As String, dwType As Long, lpDataSize As Long) As Long
    Dim Full As KEY_VALUE_FULL_INFORMATION, us As UNICODE_STRING
    Dim st As Long, Length As Long
    RtlInitUnicodeString us, StrPtr(Name)
    QueryValueKey = AllocMemory(1024)
    st = ZwQueryValueKey(hKey, us, KeyValueFullInformation, ByVal QueryValueKey, 1024, Length)
    If st = STATUS_OBJECT_NAME_NOT_FOUND Then
        dwType = -1
        lpDataSize = 0
        FreeMemory QueryValueKey
        QueryValueKey = 0
        Exit Function
    End If
    If Length > 1024 Then
        FreeMemory QueryValueKey
        QueryValueKey = AllocMemory(Length)
        st = ZwQueryValueKey(hKey, us, KeyValueFullInformation, ByVal QueryValueKey, Length, Length)
        If Not NT_SUCCESS(st) Then
            QueryValueKey = 0
            MsgBox "获取数据失败！", vbCritical
            Exit Function
        End If
    End If
    CopyMemory VarPtr(Full), QueryValueKey, 20
    Name = Space(Full.NameLength \ 2)
    CopyMemory StrPtr(Name), QueryValueKey + 20, Full.NameLength
    dwType = Full.Type
    Dim Buff As Long
    Buff = AllocMemory(Full.DataLength)
    CopyMemory Buff, QueryValueKey + Full.DataOffset, Full.DataLength
    FreeMemory QueryValueKey
    QueryValueKey = Buff
    lpDataSize = Full.DataLength
End Function

Public Function QueryValueKeyString(ByVal hKey As Long, Name As String, HasNode As Boolean) As String
    Dim lPtr As Long, d As Long
    lPtr = QueryValueKey(hKey, Name, 0, d)
    If lPtr = 0 Then
        HasNode = False
        Exit Function
    End If
    QueryValueKeyString = Space(d \ 2)
    CopyMemory StrPtr(QueryValueKeyString), lPtr, d
    FreeMemory lPtr
    HasNode = True
End Function

Public Function QueryValueKeyDWord(ByVal hKey As Long, Name As String, HasNode As Boolean) As Long
    Dim lPtr As Long
    lPtr = QueryValueKey(hKey, Name, 0, 0)
    If lPtr = 0 Then Exit Function
    CopyMemory VarPtr(QueryValueKeyDWord), lPtr, 4
    FreeMemory lPtr
    HasNode = True
End Function

Public Function EnumReg(ByVal hNode As Node, Optional ByVal NotQuiet As Boolean = False) As Boolean
    Dim Index As Long, xx As Long
    Dim Length As Long, Key As Long
    Dim Name As String, Cache As KEY_CACHED_INFORMATION
    'Exit Function
    Key = OpenRegKey(hNode.FullPath, KEY_ALL_ACCESS, NotQuiet)
    Dim NameInfo() As Byte
    If Key = 0 Then Exit Function
    Do
        ReDim NameInfo(1023)
        ZwEnumerateKey Key, Index, KeyNameInformation, ByVal VarPtr(NameInfo(0)), 1024, Length
        If Length = 0 Then GoTo LoopNext
        If Length > 1024 Then
            ReDim NameInfo(Length - 1)
            ZwEnumerateKey Key, Index, KeyNameInformation, ByVal NameInfo(0), Length, Length
            If Length = 0 Then GoTo LoopNext
        End If
        CopyMemory VarPtr(Length), VarPtr(NameInfo(0)), 4
        Name = Space(Length \ 2)
        CopyMemory StrPtr(Name), VarPtr(NameInfo(4)), Length
        Debug.Print ZwEnumerateKey(Key, Index - 1, KeyFullInformation, ByVal VarPtr(Cache), Len(Cache), Length)
        Menu.tvwKeys.Nodes.Add hNode.Key, tvwChild, hNode.FullPath & "\" & Name, Name
        If Cache.SubKeys > 0 Then
            Menu.tvwKeys.Nodes.Add hNode.FullPath & "\" & Name, tvwChild
        End If
        If Index And 255 = 0 Then DoEvents
LoopNext:
    Loop
    ZwClose Key
    EnumReg = True
End Function

Public Sub EnumValue(ByVal hNode As Node, Optional ByVal NotQuiet As Boolean = False)
    Dim Key As Long
    Key = OpenRegKey(hNode.FullPath, KEY_ALL_ACCESS, NotQuiet)
    Dim Index As Long, Name As String, cbName As Long, lpType As Long, lpcbData As Long
    Dim lData() As Byte, result As Long, kkObj As ListItem
    Menu.lvwData.ListItems.Clear
    With Menu.lvwData.ListItems.Add(, , "(默认)")
        .SubItems(1) = "REG_SZ"
        .SubItems(2) = "(未设置)"
    End With
    Dim KeyInf As KEY_FULL_INFORMATION
    Do
        ReDim lData(255)
        cbName = 256
        Name = Space(cbName)
        lpcbData = 256
        result = RegEnumValue(Key, Index, Name, cbName, 0, lpType, lData(0), lpcbData)
        If result = 234 Then
            ReDim lData(lpcbData - 1)
            Name = Space(cbName)
            result = RegEnumValue(Key, Index, Name, cbName, 0, lpType, lData(0), lpcbData)
        End If
        If result = ERROR_NO_MORE_ITEMS Then '枚举完毕
            Exit Do
        End If
        If result <> ERROR_SUCCESS Then
            If NotQuiet Then Err.Raise 17 '不是长度错误
            RegCloseKey Key
            Exit Sub
        End If
        If cbName = 0 Then
            Set kkObj = Menu.lvwData.ListItems(1)
        Else
            Set kkObj = Menu.lvwData.ListItems.Add(, , Name)
        End If
        With kkObj
            Select Case lpType
            Case REG_NONE
                .SubItems(1) = "REG_NONE"
            Case REG_SZ
                .SubItems(1) = "REG_SZ"
                .SubItems(2) = StrConv(lData, vbUnicode)
            Case REG_EXPAND_SZ
                .SubItems(1) = "REG_EXPAND_SZ"
                .SubItems(2) = StrConv(lData, vbUnicode)
            Case REG_BINARY
                .SubItems(1) = "REG_BINARY"
                Call Base64Array_Encode(lData)
                .SubItems(2) = StrConv(lData, vbUnicode)
            Case REG_DWORD
                .SubItems(1) = "REG_DWORD"
                Dim DWord As Long
                CopyMemory VarPtr(DWord), VarPtr(lData(0)), 4
                .SubItems(2) = FormatHex(DWord) & " (" & DWord & ")"
            Case REG_DWORD_BIG_ENDIAN
                .SubItems(1) = "REG_DWORD_BIG_ENDIAN"
                CopyMemory VarPtr(DWord), VarPtr(lData(0)), 4
                .SubItems(2) = DWord
            Case REG_LINK
                .SubItems(1) = "REG_LINK"
                .SubItems(2) = StrConv(lData, vbUnicode)
            Case REG_MULTI_SZ
                .SubItems(1) = "REG_MULTI_SZ"
                .SubItems(2) = StrConv(lData, vbUnicode)
            End Select
        End With
        Index = Index + 1
    Loop
    RegCloseKey Key
    Set kkObj = Nothing
End Sub

Public Function DeleteKey(ByVal szPath As String, Optional ByVal NotQuiet As Boolean = False) As Boolean
    Dim s As String
    Dim root As String
    Dim subkey As String
    Dim rootkey As HKEYs
    Dim ComputerName As String
    ComputerName = left(szPath, InStr(szPath, "\") - 1)
    szPath = Mid(szPath, InStr(szPath, "\") + 1)
    If InStr(szPath, "\") > 0 Then
        root = left(szPath, InStr(szPath, "\") - 1)
        subkey = Mid(szPath, InStr(szPath, "\") + 1)
    Else
        root = szPath
    End If
    Select Case root
    Case "HKEY_LOCAL_MACHINE"
        rootkey = eHKEY_LOCAL_MACHINE
    Case "HKEY_CLASSES_ROOT"
        rootkey = eHKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
        rootkey = eHKEY_CURRENT_USER
    Case "HKEY_CURRENT_CONFIG"
        rootkey = eHKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA"
        rootkey = eHKEY_DYN_DATA
    Case "HKEY_PERFORMANCE_DATA"
        rootkey = eHKEY_PERFORMANCE_DATA
    Case "HKEY_USERS"
        rootkey = eHKEY_USERS
    End Select
    Dim Key As Long, result As Long
    If ComputerName <> "我的电脑" Then
        result = RegConnectRegistry(ComputerName, rootkey, Key)
        If result = ERROR_SUCCESS Then
            If RegDeleteKey(result, subkey) <> ERROR_SUCCESS Then
                MsgBox "打开远程根注册表项时出错！", vbCritical
                DeleteKey = False
            Else
                DeleteKey = True
            End If
            Exit Function
        Else
            If NotQuiet Then MsgBox "打开远程根注册表项时出错！", vbCritical
            DeleteKey = False
            Exit Function
        End If
    Else
        If RegDeleteKey(rootkey, subkey) <> ERROR_SUCCESS Then
            MsgBox "打开远程根注册表项时出错！", vbCritical
            DeleteKey = False
        Else
            DeleteKey = True
        End If
    End If
End Function

Public Function CreateKey(ByVal szPath As String, ByVal szKeyName As String, Optional ByVal NotQuiet As Boolean) As Boolean
    Dim root As String
    Dim subkey As String
    Dim rootkey As HKEYs
    Dim ComputerName As String
    ComputerName = left(szPath, InStr(szPath, "\") - 1)
    szPath = Mid(szPath, InStr(szPath, "\") + 1)
    If InStr(szPath, "\") > 0 Then
        root = left(szPath, InStr(szPath, "\") - 1)
        subkey = Mid(szPath, InStr(szPath, "\") + 1)
    Else
        root = szPath
    End If
    Select Case root
    Case "HKEY_LOCAL_MACHINE"
        rootkey = eHKEY_LOCAL_MACHINE
    Case "HKEY_CLASSES_ROOT"
        rootkey = eHKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
        rootkey = eHKEY_CURRENT_USER
    Case "HKEY_CURRENT_CONFIG"
        rootkey = eHKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA"
        rootkey = eHKEY_DYN_DATA
    Case "HKEY_PERFORMANCE_DATA"
        rootkey = eHKEY_PERFORMANCE_DATA
    Case "HKEY_USERS"
        rootkey = eHKEY_USERS
    End Select
    
    Dim s As String
    Dim subkey As String
    Dim ComputerName As String
    ComputerName = left(szPath, InStr(szPath, "\") - 1)
    s = Mid(szPath, InStr(szPath, "\") + 1)
    If InStr(s, "\") > 0 Then
        root = left(s, InStr(s, "\") - 1)
        subkey = Mid(s, InStr(s, "\") + 1)
    Else
        root = s
    End If
    Select Case root
    Case "HKEY_LOCAL_MACHINE"
        s = oHKEY_LOCAL_MACHINE
    Case "HKEY_CLASSES_ROOT"
        s = oHKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
        s = oHKEY_CURRENT_USER
    Case "HKEY_CURRENT_CONFIG"
        s = oHKEY_CURRENT_CONFIG
    Case "HKEY_USERS"
        s = oHKEY_USERS
    End Select
    Dim oa As OBJECT_ATTRIBUTES, us As UNICODE_STRING
    If subkey <> "" Then
        s = s & "\" & subkey
    End If
    RtlInitUnicodeString us, StrPtr(szPath)
    oa.Length = Len(oa)
    oa.ObjectName = VarPtr(us)
    oa.Attributes = OBJ_CASE_INSENSITIVE
    Call ZwCreateKey(Key, Access, oa, 0, ByVal 0, 0, ByVal 0)
    
    If Key = 0 Then
        If NotQuiet Then MsgBox "创建子项时出错！", vbCritical
        CreateKey = False
        Exit Function
    End If
    ZwClose Key
End Function

Public Function CreateValue(ByVal szPath As String, ByVal ValueName As String, ByVal ValueType As Long, ByVal NotQuiet As Boolean) As Boolean
    Dim root As String
    Dim subkey As String
    Dim rootkey As HKEYs
    Dim ComputerName As String
    ComputerName = left(szPath, InStr(szPath, "\") - 1)
    szPath = Mid(szPath, InStr(szPath, "\") + 1)
    If InStr(szPath, "\") > 0 Then
        root = left(szPath, InStr(szPath, "\") - 1)
        subkey = Mid(szPath, InStr(szPath, "\") + 1)
    Else
        root = szPath
    End If
    Select Case root
    Case "HKEY_LOCAL_MACHINE"
        rootkey = eHKEY_LOCAL_MACHINE
    Case "HKEY_CLASSES_ROOT"
        rootkey = eHKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
        rootkey = eHKEY_CURRENT_USER
    Case "HKEY_CURRENT_CONFIG"
        rootkey = eHKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA"
        rootkey = eHKEY_DYN_DATA
    Case "HKEY_PERFORMANCE_DATA"
        rootkey = eHKEY_PERFORMANCE_DATA
    Case "HKEY_USERS"
        rootkey = eHKEY_USERS
    End Select
    Dim Key As Long, result As Long, s As Long, d As Long, sAttr As SECURITY_ATTRIBUTES
    If ComputerName <> "我的电脑" Then
        result = RegConnectRegistry(ComputerName, rootkey, Key)
        If result = ERROR_SUCCESS Then
            If RegOpenKeyEx(result, subkey, 0, KEY_ALL_ACCESS, Key) <> ERROR_SUCCESS Then
                MsgBox "打开远程注册表项时出错！", vbCritical
                Exit Function
            End If
        Else
            If NotQuiet Then MsgBox "打开远程根注册表项时出错！", vbCritical
            Exit Function
        End If
    Else
        If RegOpenKeyEx(rootkey, subkey, 0, KEY_ALL_ACCESS, Key) <> ERROR_SUCCESS Then
            If NotQuiet Then MsgBox "打开远程根注册表项时出错！", vbCritical
            Exit Function
        End If
    End If
    If ValueType = REG_DWORD Or ValueType = REG_DWORD_LITTLE_ENDIAN Or ValueType = REG_DWORD_BIG_ENDIAN Then
        result = RegSetValueEx(Key, ValueName, 0, ValueType, 0, 4)
    Else
        result = RegSetValueEx(Key, ValueName, 0, ValueType, 0, 0)
    End If
    If result <> ERROR_SUCCESS Then
        If NotQuiet Then MsgBox "创建注册表值时出错！", vbCritical
        Exit Function
    End If
    RegCloseKey Key
    CreateValue = True
End Function

Public Function DeleteValue(ByVal szPath As String, ByVal ValueName As String, ByVal NotQuiet As Boolean) As Boolean
    Dim Key As Long, result As Long, s As Long, d As Long, sAttr As SECURITY_ATTRIBUTES
    Key = OpenRegKey(szPath, KEY_ALL_ACCESS, NotQuiet)
    result = RegDeleteValue(Key, ValueName)
    If result <> ERROR_SUCCESS Then
        If NotQuiet Then MsgBox "删除注册表值时出错！", vbCritical
        Exit Function
    End If
    ZwClose Key
    DeleteValue = True
End Function

Public Function SetReg(ByVal szPath As String, ByVal ValueName As String, ByVal szClassName As String, ByVal NotQuiet As Boolean, ByVal IsFirst As Boolean)
    Dim root As String
    Dim subkey As String
    Dim rootkey As HKEYs
    Dim ComputerName As String, i As Long
    ComputerName = left(szPath, InStr(szPath, "\") - 1)
    szPath = Mid(szPath, InStr(szPath, "\") + 1)
    If InStr(szPath, "\") > 0 Then
        root = left(szPath, InStr(szPath, "\") - 1)
        subkey = Mid(szPath, InStr(szPath, "\") + 1)
    Else
        root = szPath
    End If
    Select Case root
    Case "HKEY_LOCAL_MACHINE"
        rootkey = eHKEY_LOCAL_MACHINE
    Case "HKEY_CLASSES_ROOT"
        rootkey = eHKEY_CLASSES_ROOT
    Case "HKEY_CURRENT_USER"
        rootkey = eHKEY_CURRENT_USER
    Case "HKEY_CURRENT_CONFIG"
        rootkey = eHKEY_CURRENT_CONFIG
    Case "HKEY_DYN_DATA"
        rootkey = eHKEY_DYN_DATA
    Case "HKEY_PERFORMANCE_DATA"
        rootkey = eHKEY_PERFORMANCE_DATA
    Case "HKEY_USERS"
        rootkey = eHKEY_USERS
    End Select
    Dim Key As Long, result As Long, s As Long, d As Long, sAttr As SECURITY_ATTRIBUTES
    If ComputerName <> "我的电脑" Then
        result = RegConnectRegistry(ComputerName, rootkey, Key)
        If result = ERROR_SUCCESS Then
            If RegOpenKeyEx(result, subkey, 0, KEY_ALL_ACCESS, Key) <> ERROR_SUCCESS Then
                If NotQuiet Then MsgBox "打开远程注册表项时出错！", vbCritical
                SetReg = False
                Exit Function
            End If
        Else
            If NotQuiet Then MsgBox "打开远程根注册表项时出错！", vbCritical
            SetReg = False
            Exit Function
        End If
    Else
        If RegOpenKeyEx(rootkey, subkey, 0, KEY_ALL_ACCESS, Key) <> ERROR_SUCCESS Then
            If NotQuiet Then MsgBox "打开远程根注册表项时出错！", vbCritical
            SetReg = False
            Exit Function
        End If
    End If
    Dim Classs As Long
    If IsFirst Then
        Classs = REG_SZ
    Else
        result = RegQueryValueEx(Key, ValueName, 0, s, CLng(0), 4)
        If result <> 0 And result <> 234 Then
            If NotQuiet Then MsgBox "获取数据类型时出错！", vbCritical
            Exit Function
        End If
    End If
    If s = REG_DWORD Or s = REG_DWORD_BIG_ENDIAN Then
        With DGEditDWord
            .Init Key, ValueName, szClassName, s
            .Show vbModal, Menu
            RegCloseKey Key
        End With
    Else
        With DGEditValue
            .Init szClassName, Key, ValueName
            .Show vbModal, Menu
            RegCloseKey Key
        End With
    End If
End Function

Attribute VB_Name = "Module"
Option Explicit
Public Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Public Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Public Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Public Declare Function Module32First Lib "kernel32.dll" (ByVal hSnapshot As Long, lppe As MODULEENTRY32) As Long
Public Declare Function Module32Next Lib "kernel32.dll" (ByVal hSnapshot As Long, lppe As MODULEENTRY32) As Long
Public Declare Function EnumProcessModules Lib "psapi.dll" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetModuleFileName Lib "kernel32.dll" Alias "GetModuleFileNameA" (ByVal hModule As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long
Public Declare Function GetModuleFileNameEx Lib "psapi.dll" Alias "GetModuleFileNameExA" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Public Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Public Declare Function GetModuleInformation Lib "psapi.dll" (ByVal hProcess As Long, ByVal hModule As Long, ByVal lpModuleInfo As Long, ByVal cb As Long) As Long
Public Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Public Declare Function ZwUnmapViewOfSection Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, ByVal BaseAddress As Long) As Long


Public Type MODULEENTRY32
    dwSize As Long             'ָ���ṹ�Ĵ�С���ڵ���Module32Firstǰ��Ҫ���ã����򽫻�ʧ��
    th32ModuleID As Long       'ģ���
    th32ProcessID As Long      '������ģ��Ľ��̺�
    GlblcntUsage As Long       '��ģ���ȫ�����ü���
    ProccntUsage As Long       '����ģ��Ľ����������е�ģ�����ü���
    modBaseAddr As Byte        'ģ�����ַ
    modBaseSize As Long        'ģ���С���ֽ�����
    hModule As Long            '����ģ��Ľ����������е�hModule���
    szModule As String * 256   'ģ����
    szExePath As String * 1024 'ģ���Ӧ���ļ�����·��
End Type

Public Type MODULEINFO
    lpBaseOfDll As Long
    SizeOfImage As Long
    EntryPoint As Long
End Type

Public Type UNICODE_STRING
    Length                          As Long
    MaximumLength                   As Long
    buffer                          As Long
End Type

Public Type LDR_MODULE 'LDR_DATA_TABLE_ENTRY
    InLoadOrderModuleList           As LIST_ENTRY
    InMemoryOrderModuleList         As LIST_ENTRY
    InInitializationOrderModuleList As LIST_ENTRY
    BaseAddress                     As Long
    EntryPoint                      As Long
    SizeOfImage                     As Long
    FullDllName                     As UNICODE_STRING
    BaseDllName                     As UNICODE_STRING
    flags                           As Long
    LoadCount                       As Integer
    TlsIndex                        As Integer
    HashTableEntry                  As LIST_ENTRY
    TimeDateStamp                   As Long
End Type

Public Type PEB_LDR_DATA
    Length                          As Long
    Initialized                     As Long
    SsHandle                        As Long
    InLoadOrderModuleList           As LIST_ENTRY
    InMemoryOrderModuleList         As LIST_ENTRY
    InInitializationOrderModuleList As LIST_ENTRY
End Type


Public Sub ListAllModules(ByVal pid As Long)
    Dim MODULEINFO As MODULEENTRY32
    Dim cne As Long
    Dim msh As Long
    Dim mPath As String
    Dim mNature As String
    Dim nIndex As Long
    Dim hProcess As Long
    
    If ModuleList.ListView1.Sorted = True Then ModuleList.ListView1.Sorted = False
    nIndex = 1
    
    If ModuleList.ListView1.ListItems.Count > 0 Then
        nIndex = ModuleList.ListView1.SelectedItem.Index
    End If
    
    hProcess = FxNormalOpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, pid)
    mPath = GetProcessName(GetProcessPath(hProcess))
    ModuleList.Caption = "[" & (mPath) & "]�е�ģ��"
    
    ModuleList.ListView1.ListItems.Clear
    
    msh = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, pid)
    MODULEINFO.dwSize = LenB(MODULEINFO)
    
    cne = Module32First(msh, MODULEINFO)
    Do While cne
        mNature = "Normal"
        If MODULEINFO.ProccntUsage = 65535 Then mNature = "System"
        ModuleList.ListView1.ListItems.Add , , MODULEINFO.szModule
        With ModuleList.ListView1.ListItems(ModuleList.ListView1.ListItems.Count)
            .SubItems(1) = FormatHex(MODULEINFO.hModule) 'FormatHex(ModuleInfo.hModule)
            .SubItems(2) = MODULEINFO.szExePath
            .SubItems(3) = FormatHex(FxGetModuleEntryFuncAddr(hProcess, MODULEINFO.hModule))
            .SubItems(4) = FxGetModuleSize(hProcess, MODULEINFO.hModule)
        End With
        cne = Module32Next(msh, MODULEINFO)
    Loop
    
    ZwClose msh
    ZwClose hProcess
    
    DoEvents
    
    If ModuleList.ListView1.ListItems.Count > nIndex Then
        ModuleList.ListView1.ListItems(nIndex).Selected = True
        ModuleList.ListView1.ListItems(nIndex).EnsureVisible
    End If
    
    ModuleList.Caption = ModuleList.Caption & " (" & ModuleList.ListView1.ListItems.Count & ")"
End Sub

Public Sub FxEnumModulesByVirtualMemory(ByVal pid As Long)
    Dim Memory As MEMORY_BASIC_INFORMATION
    Dim pFind As Long
    Dim hProcess As Long
    Dim hAppHandle As Long
    Dim szModPath As String * 256
    Dim szModName As String
    Dim errStr As String
    Dim mPath As String
    Dim nIndex As Long
    
    If ModuleList.ListView1.Sorted = True Then ModuleList.ListView1.Sorted = False
    nIndex = 1
    
    If ModuleList.ListView1.ListItems.Count > 0 Then
        nIndex = ModuleList.ListView1.SelectedItem.Index
    End If
    
    hProcess = FxNormalOpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, pid)
    mPath = GetProcessName(GetProcessPath(hProcess))
    ModuleList.Caption = "[" & (mPath) & "]�е�ģ��"
    
    If hProcess = 0 Then errStr = "�򿪽���ʧ��!": GoTo errors
        
    ModuleList.ListView1.ListItems.Clear
    
    Do While VirtualQueryEx(hProcess, pFind, Memory, LenB(Memory)) = LenB(Memory)
        If Memory.State = MEM_FREE Then
            Memory.AllocationBase = Memory.BaseAddress
        End If
    
        If Not (Memory.BaseAddress = hAppHandle Or Memory.AllocationBase <> Memory.BaseAddress Or Memory.AllocationBase = 0) Then
            szModPath = ""
            If GetModuleFileNameEx(hProcess, Memory.AllocationBase, szModPath, 256) Then
                szModName = GetProcessName(szModPath)
                    
                ModuleList.ListView1.ListItems.Add , , szModName
                With ModuleList.ListView1.ListItems(ModuleList.ListView1.ListItems.Count)
                    .SubItems(1) = FormatHex(Memory.AllocationBase) 'FormatHex(ModuleInfo.hModule)
                    .SubItems(2) = szModPath
                    .SubItems(3) = FormatHex(FxGetModuleEntryFuncAddr(hProcess, Memory.AllocationBase))
                    .SubItems(4) = FxGetModuleSize(hProcess, Memory.AllocationBase)
                End With
            End If
        End If
            
        pFind = pFind + Memory.RegionSize
        With Memory
            .AllocationBase = 0
            .AllocationProtect = 0
            .BaseAddress = 0
            .Protect = 0
            .RegionSize = 0
            .State = 0
            .Type = 0
        End With
    Loop
    
    DoEvents
    
    If ModuleList.ListView1.ListItems.Count > nIndex Then
        ModuleList.ListView1.ListItems(nIndex).Selected = True
        ModuleList.ListView1.ListItems(nIndex).EnsureVisible
    End If
    
    ModuleList.Caption = ModuleList.Caption & " (" & ModuleList.ListView1.ListItems.Count & ")"
    
    Exit Sub
errors:
    MsgBox errStr, 0, "����"
End Sub

Public Function FxUnloadDllByUnmapView(ByVal pid As Long, ByVal hModule As Long, Optional ByVal DllName As String = "") As Long
    Dim hProcess As Long
    Dim errStr As String
    Dim temp As String * 260

    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, pid)
    If hProcess = 0 Then errStr = "�򿪽���ʧ��!": GoTo errors
   
    If hModule = 0 Then
        hModule = GetModuleHandle(DllName)
        If hModule = 0 Then errStr = "��ȡģ����ʧ��!": GoTo errors
    End If
    
    ZwUnmapViewOfSection hProcess, hModule
    
    If GetModuleFileNameEx(hProcess, hModule, temp, 260) Then   '������ܻ�ȡ��ģ������˵����ģ���Դ���
        If hModule = 0 Then errStr = "ж��ģ��ʧ��!": GoTo errors
    Else
        FxUnloadDllByUnmapView = 1
    End If
    
    ZwClose hProcess
    
    Exit Function
errors:
    MsgBox errStr, 0, "����"
    FxUnloadDllByUnmapView = 0
End Function

Public Function FxRemoteProcessInsertDll(ByVal pid As Long, ByVal DllPath As String) As Long
    Dim lpThreadAttributes As SECURITY_ATTRIBUTES
    Dim hProcess As Long
    Dim hThread As Long
    Dim hModule As Long
    Dim DllBuffer As Long
    Dim DllFileSize As Long
    Dim rReturn As Long
    Dim fAddr As Long
    Dim Cid As CLIENT_ID
    Dim errNum As Long
    Dim errStr As String
    
    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, pid)
    '����Ȩ��PROCESS_QUERY_INFORMATION Or PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE
    If hProcess = 0 Then errStr = "�򿪽���ʧ��!": GoTo errors

    DllFileSize = LenB(StrConv(DllPath, vbFromUnicode)) + 1
    DllBuffer = VirtualAllocEx(hProcess, 0, DllFileSize, MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
    If DllBuffer = 0 Then errStr = "�����ڴ�ռ�ʧ��!": GoTo errors

    rReturn = WriteProcessLongMemory(hProcess, DllBuffer, ByVal DllPath, DllFileSize, 0)
    If rReturn = 0 Then errStr = "д���ڴ�ʧ��!": GoTo errors
    
    hModule = GetModuleHandle("kernel32.dll")
    If hModule = 0 Then errStr = "��ȡģ���ַʧ��!": GoTo errors
    
    fAddr = GetProcAddress(hModule, "LoadLibraryA")
    If fAddr = 0 Then errStr = "��ȡ������ַʧ��!": GoTo errors

    hThread = CreateRemoteThread(hProcess, lpThreadAttributes, 0, fAddr, DllBuffer, 0, ByVal 0&)
    'hThread = ChCreateRemoteThread(hProcess, fAddr, DllBuffer, Cid)
    If hThread = 0 Then errStr = "����Dllʧ��!": GoTo errors
    
    WaitForSingleObject hThread, INFINITE
    
    
    '<�������������Ϣ����
    errNum = GetLastError
    Debug.Print "DllFileSize:" & DllFileSize & ",DllBuffer:" & FormatHex(DllBuffer) & ",hThread:" & hThread & ",errNum:" & errNum
    '��������������������>
    
    VirtualFreeEx hProcess, DllBuffer, DllFileSize, MEM_DECOMMIT
    
    ZwClose hProcess
    ZwClose hThread
    
    FxRemoteProcessInsertDll = 1
    Exit Function
errors:
    MsgBox errStr, 0, "����"
    FxRemoteProcessInsertDll = 0
End Function

Public Function FxRemoteProcessFreeDll(ByVal pid As Long, ByVal hModule As Long, Optional ByVal DllName As String = "") As Long
    Dim lpThreadAttributes As SECURITY_ATTRIBUTES
    Dim hProcess As Long
    Dim hThread As Long
    Dim hFunction As Long
    Dim mName As String * 256
    Dim i As Long
    Dim uSucceed As Long
    Dim uMax As Long
    Dim Cid As CLIENT_ID
    Dim errNum As Long
    Dim errStr As String
    
    If hModule = 0 Then
        hModule = FxGetRemoteModuleName(pid, DllName)
        If hModule = 0 Then errStr = "��ȡģ����ʧ��!": GoTo errors
    End If

    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, pid)
    If hProcess = 0 Then errStr = "�򿪽���ʧ��!": GoTo errors
    
    hFunction = GetModuleHandle("kernel32.dll")
    hFunction = GetProcAddress(hFunction, "FreeLibrary")
    
    Do
        hThread = CreateRemoteThread(hProcess, lpThreadAttributes, 0, hFunction, hModule, 0, 0)
        'hThread = ChCreateRemoteThread(hProcess, hFunction, hModule, Cid)
        'WaitForSingleObject hThread, INFINITE
        uSucceed = 0
        GetExitCodeThread hThread, uSucceed
        uMax = uMax + 1
    Loop While uSucceed And uMax < 200
    
    If uSucceed = 1 Then errStr = "ж��ģ��ʧ��!": GoTo errors  '��������Է���1(�ɹ�)����˵��ģ��û�б�ж��
    
    '<�������������Ϣ����
    errNum = GetLastError
    Debug.Print "hFunction:" & hFunction & ",hModule:" & hModule & ",hThread:" & hThread & ",DllName:" & DllName & ",errNum:" & errNum
    '��������������������>
    '���֪ʶ������ͬһ��DLL��ÿ����һ��LoadLibrary���Ὣ��DLL�����ü�������1����FreeLibrary����ʱ����Ӧ�ļ�ȥ1��ֱ������Ϊ0ʱϵͳ������Free����DLL��

    ZwClose hProcess
    ZwClose hThread
    
    FxRemoteProcessFreeDll = 1
    Exit Function
errors:
    MsgBox errStr, 0, "����"
    FxRemoteProcessFreeDll = 0
End Function

Public Function FxGetRemoteModuleName(ByVal pid As Long, ByVal ModuleName As String) As Long
    Dim lpThreadAttributes As SECURITY_ATTRIBUTES
    Dim hThread As Long
    Dim hProcess As Long
    Dim DllPath As String
    Dim DllBuffer As Long
    Dim DllFileSize As Long
    Dim fAddr As Long
    Dim rReturn As Long
    Dim errNum As Long
    Dim hModule As Long
    Dim errStr As String
    
    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, pid)
    '����Ȩ��PROCESS_QUERY_INFORMATION Or PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE
    If hProcess = 0 Then errStr = "�򿪽���ʧ��!": GoTo errors
    
    DllPath = ModuleName
    DllFileSize = LenB(StrConv(DllPath, vbFromUnicode)) + 1
    DllBuffer = VirtualAllocEx(hProcess, 0, DllFileSize, MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
    If DllBuffer = 0 Then errStr = "�����ڴ�ռ�ʧ��!": GoTo errors

    rReturn = WriteProcessLongMemory(hProcess, DllBuffer, ByVal DllPath, DllFileSize, 0)
    If DllBuffer = 0 Then errStr = "д���ڴ�ʧ��!": GoTo errors
    
    fAddr = GetProcAddress(GetModuleHandle("kernel32.dll"), "GetModuleHandleA")
    hThread = CreateRemoteThread(hProcess, lpThreadAttributes, 0, fAddr, DllBuffer, 0, ByVal 0&)
    WaitForSingleObject hThread, INFINITE
    GetExitCodeThread hThread, hModule
    If DllBuffer = 0 Then errStr = "��ȡģ����ʧ��!": GoTo errors
    
    '<�������������Ϣ����
    errNum = GetLastError
    Debug.Print "DllFileSize:" & DllFileSize & ",DllBuffer:" & FormatHex(DllBuffer) & ",hThread:" & hThread & ",hModule:" & hModule & ",errNum:" & errNum
    '��������������������>
    
    VirtualFreeEx hProcess, DllBuffer, DllFileSize, MEM_DECOMMIT
    ZwClose hProcess
    ZwClose hThread
    
    FxGetRemoteModuleName = hModule
errors:
    MsgBox errStr, 0, "����"
    FxGetRemoteModuleName = 0
End Function

Public Function FxGetModuleSize(ByVal hProcess As Long, ByVal hModule As Long) As Long
    Dim mi As MODULEINFO
    
    GetModuleInformation hProcess, hModule, VarPtr(mi), Len(mi)
    
    FxGetModuleSize = mi.SizeOfImage
End Function

Public Function FxGetModuleEntryFuncAddr(ByVal hProcess As Long, ByVal hModule As Long) As Long
    Dim mi As MODULEINFO
    
    GetModuleInformation hProcess, hModule, VarPtr(mi), Len(mi)
    
    FxGetModuleEntryFuncAddr = mi.EntryPoint
End Function

Public Sub MNNew(ByVal pid As Long)
    With ModuleList.ListView1
        If .Tag = 0 Then
            Call ListAllModules(pid)
        ElseIf .Tag = 1 Then
            Call FxEnumModulesByVirtualMemory(pid)
        End If
    End With
End Sub

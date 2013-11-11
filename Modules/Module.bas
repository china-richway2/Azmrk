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

Public Type LDR_MODULE 'LDR_DATA_TABLE_ENTRY
    InLoadOrderModuleList           As LIST_ENTRY
    InMemoryOrderModuleList         As LIST_ENTRY
    InInitializationOrderModuleList As LIST_ENTRY
    BaseAddress                     As Long
    EntryPoint                      As Long
    SizeOfImage                     As Long
    FullDllName                     As UNICODE_STRING
    BaseDllName                     As UNICODE_STRING
    Flags                           As Long
    Loadcount                       As Integer
    TlsIndex                        As Integer
    HashTableEntry                  As LIST_ENTRY
    TimeDateStamp                   As Long
End Type


Public Sub ListAllModules(ByVal PID As Long, ByVal OwnerForm As ModuleList)
    Dim MODULEINFO As MODULEENTRY32
    Dim cne As Long
    Dim msh As Long
    Dim mPath As String
    Dim mNature As String
    Dim nIndex As Long
    Dim hProcess As Long
    
    If OwnerForm.ListView1.Sorted = True Then OwnerForm.ListView1.Sorted = False
    nIndex = 1
    
    If OwnerForm.ListView1.ListItems.count > 0 Then
        nIndex = OwnerForm.ListView1.SelectedItem.Index
    End If
    
    hProcess = FxNormalOpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, PID)
    
    OwnerForm.ListView1.ListItems.Clear
    
    msh = CreateToolhelp32Snapshot(TH32CS_SNAPMODULE, PID)
    MODULEINFO.dwSize = LenB(MODULEINFO)
    
    cne = Module32First(msh, MODULEINFO)
    Do While cne
        mNature = "Normal"
        If MODULEINFO.ProccntUsage = 65535 Then mNature = "System"
        With OwnerForm.ListView1.ListItems.Add(, , MODULEINFO.szModule)
        'With OwnerForm.ListView1.ListItems(OwnerForm.ListView1.ListItems.Count)
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
    
    If OwnerForm.ListView1.ListItems.count > nIndex Then
        OwnerForm.ListView1.ListItems(nIndex).Selected = True
        OwnerForm.ListView1.ListItems(nIndex).EnsureVisible
    End If
    
    OwnerForm.Caption = "[" & OwnerForm.ListView1.ListItems(1).Text & "]�е�ģ�� (" & OwnerForm.ListView1.ListItems.count & "��)"
End Sub

Public Sub FxEnumModulesByVirtualMemory(ByVal PID As Long, ByVal OwnerForm As ModuleList)
    Dim Memory As MEMORY_BASIC_INFORMATION
    Dim pFind As Long
    Dim hProcess As Long
    Dim hAppHandle As Long
    Dim szModPath As String * 256
    Dim szModName As String
    Dim errStr As String
    Dim mPath As String
    Dim nIndex As Long
    
    If OwnerForm.ListView1.Sorted = True Then OwnerForm.ListView1.Sorted = False
    nIndex = 1
    
    If OwnerForm.ListView1.ListItems.count > 0 Then
        nIndex = OwnerForm.ListView1.SelectedItem.Index
    End If
    
    hProcess = FxNormalOpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, PID)
    
    If hProcess = 0 Then errStr = "�򿪽���ʧ��!": GoTo errors
        
    OwnerForm.ListView1.ListItems.Clear
    
    Do While VirtualQueryEx(hProcess, pFind, Memory, LenB(Memory)) = LenB(Memory)
        If Memory.State = MEM_FREE Then
            Memory.AllocationBase = Memory.BaseAddress
        End If
    
        If Not (Memory.BaseAddress = hAppHandle Or Memory.AllocationBase <> Memory.BaseAddress Or Memory.AllocationBase = 0) Then
            szModPath = ""
            If GetModuleFileNameEx(hProcess, Memory.AllocationBase, szModPath, 256) Then
                szModName = GetProcessName(szModPath)
                    
                With OwnerForm.ListView1.ListItems.Add(, , szModName)
                'With OwnerForm.ListView1.ListItems(OwnerForm.ListView1.ListItems.Count)
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
    
    If OwnerForm.ListView1.ListItems.count > nIndex Then
        OwnerForm.ListView1.ListItems(nIndex).Selected = True
        OwnerForm.ListView1.ListItems(nIndex).EnsureVisible
    End If
    
    OwnerForm.Caption = "[" & OwnerForm.ListView1.ListItems(1).Text & "]�е�ģ�� (" & OwnerForm.ListView1.ListItems.count & "��)"
    
    Exit Sub
errors:
    MsgBox errStr, 0, "����"
End Sub

Public Sub RdNewByReadMemory(ByVal PID As Long, ByVal OwnerForm As ModuleList)
    Dim hProcess As Long
    Dim hAppHandle As Long
    Dim szModPath As String
    Dim szModName As String
    Dim mPath As String
    Dim etStart As Long
    Dim lPtr As Long
    Dim etCur As LDR_MODULE
    Dim Basic As PROCESS_BASIC_INFORMATION
    Dim Peb As PEB_LDR_DATA
    Dim nIndex As Long
    Dim errStr As String
    nIndex = 1
    
    If OwnerForm.ListView1.Sorted = True Then OwnerForm.ListView1.Sorted = False
    
    If OwnerForm.ListView1.ListItems.count > 0 Then
        nIndex = OwnerForm.ListView1.SelectedItem.Index
    End If
    
    hProcess = FxNormalOpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, PID)
    
    If hProcess = 0 Then errStr = "�򿪽���ʧ��!": GoTo errors
        
    OwnerForm.ListView1.ListItems.Clear
    '��ȡPEBָ��
    If Not NT_SUCCESS(ZwQueryInformationProcess(hProcess, ProcessBasicInformation, Basic, Len(Basic), 0)) Then
        errStr = "��ȡPEBָ��ʧ�ܣ�"
        GoTo errors
    End If
    '��ȡPEB_LDR_DATA�ṹָ��
    If Not NT_SUCCESS(ZwReadVirtualMemory(hProcess, ByVal Basic.PebBaseAddress + 12, etStart, 4, 0)) Then
        errStr = "��ȡ�ڴ�ʧ�ܣ�"
        GoTo errors
    End If
    '��ȡPEB_LDR_DATA
    Debug.Print ZwReadVirtualMemory(hProcess, ByVal etStart, Peb, Len(Peb), 0)
    '��ȡ��һ��
    etStart = Peb.InLoadOrderModuleList.Blink
    Debug.Print ZwReadVirtualMemory(hProcess, ByVal etStart, etCur, Len(etCur), 0)
    
    Do
        szModPath = Space(etCur.FullDllName.Length \ 2)
        ZwReadVirtualMemory hProcess, ByVal etCur.FullDllName.Buffer, ByVal StrPtr(szModPath), LenB(szModPath), 0
        szModName = Space(etCur.BaseDllName.Length \ 2)
        ZwReadVirtualMemory hProcess, ByVal etCur.BaseDllName.Buffer, ByVal StrPtr(szModName), LenB(szModName), 0
        With OwnerForm.ListView1.ListItems.Add(, , szModName)
            .SubItems(1) = FormatHex(etCur.BaseAddress)
            .SubItems(2) = szModPath
            .SubItems(3) = FormatHex(etCur.EntryPoint)
            .SubItems(4) = etCur.SizeOfImage
        End With
        ZwReadVirtualMemory hProcess, ByVal etCur.InLoadOrderModuleList.Blink, etCur, Len(etCur), 0
    Loop Until etCur.InLoadOrderModuleList.Blink = etStart
    
    OwnerForm.Caption = "[" & OwnerForm.ListView1.ListItems(1).Text & "]�е�ģ�� (" & OwnerForm.ListView1.ListItems.count & "��)"
    ZwClose hProcess
    Exit Sub
errors:
    MsgBox errStr, vbCritical
    ZwClose hProcess
End Sub

Public Function FxUnloadDllByUnmapView(ByVal PID As Long, ByVal hModule As Long, Optional ByVal DllName As String = "") As Long
    Dim hProcess As Long
    Dim errStr As String
    Dim temp As String * 260

    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, PID)
    If hProcess = 0 Then errStr = "�򿪽���ʧ��!": GoTo errors
   
    If hModule = 0 Then
        hModule = GetModuleHandle(DllName)
        If hModule = 0 Then ZwClose hProcess: errStr = "��ȡģ����ʧ��!": GoTo errors
    End If
    
    ZwUnmapViewOfSection hProcess, hModule
    
    If GetModuleFileNameEx(hProcess, hModule, temp, 260) Then   '������ܻ�ȡ��ģ������˵����ģ���Դ���
        If hModule = 0 Then ZwClose hProcess: errStr = "ж��ģ��ʧ��!": GoTo errors
    Else
        FxUnloadDllByUnmapView = 1
    End If
    
    ZwClose hProcess
    
    Exit Function
errors:
    MsgBox errStr, 0, "����"
    FxUnloadDllByUnmapView = 0
End Function

Public Function FxRemoteProcessInsertDll(ByVal PID As Long, ByVal DllPath As String, ByVal IsLoadLibrary As Boolean) As Long
    Dim lpThreadAttributes As SECURITY_ATTRIBUTES
    Dim hProcess As Long
    Dim hThread As Long
    Dim hModule As Long
    Dim DllBuffer As Long
    Dim DllFileSize As Long
    Dim rReturn As Long
    Dim fAddr As Long
    Dim cid As CLIENT_ID
    Dim errNum As Long
    Dim errStr As String
    
    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, PID)
    '����Ȩ��PROCESS_QUERY_INFORMATION Or PROCESS_VM_OPERATION Or PROCESS_VM_READ Or PROCESS_VM_WRITE
    If hProcess = 0 Then errStr = "�򿪽���ʧ��!": GoTo errors

    DllFileSize = LenB(StrConv(DllPath, vbFromUnicode)) + 1
    DllBuffer = VirtualAllocEx(hProcess, 0, DllFileSize, MEM_RESERVE Or MEM_COMMIT, PAGE_READWRITE)
    If DllBuffer = 0 Then errStr = "�����ڴ�ռ�ʧ��!": GoTo errors

    rReturn = WriteProcessLongMemory(hProcess, DllBuffer, ByVal DllPath, DllFileSize, 0)
    If rReturn = 0 Then errStr = "д���ڴ�ʧ��!": GoTo errors
    
    hModule = GetModuleHandle("kernel32.dll")
    If hModule = 0 Then errStr = "��ȡģ���ַʧ��!": GoTo errors
    
    If IsLoadLibrary Then
        fAddr = GetProcAddress(hModule, "LoadLibraryA")
    Else
        fAddr = GetProcAddress(hModule, "GetModuleHandleA")
    End If
    If fAddr = 0 Then errStr = "��ȡ������ַʧ��!": GoTo errors

    hThread = CreateRemoteThread(hProcess, lpThreadAttributes, 0, fAddr, DllBuffer, 0, ByVal 0&)
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

Public Function FxRemoteProcessFreeDll(ByVal PID As Long, ByVal hModule As Long, Optional ByVal DllName As String = "") As Long
    Dim lpThreadAttributes As SECURITY_ATTRIBUTES
    Dim hProcess As Long
    Dim hThread As Long
    Dim hFunction As Long
    Dim mName As String * 256
    Dim i As Long
    Dim uSucceed As Long
    Dim uMax As Long
    Dim cid As CLIENT_ID
    Dim errNum As Long
    Dim errStr As String
    
    If hModule = 0 Then
        hModule = FxGetRemoteModuleName(PID, DllName)
        If hModule = 0 Then errStr = "��ȡģ����ʧ��!": GoTo errors
    End If

    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, PID)
    If hProcess = 0 Then errStr = "�򿪽���ʧ��!": GoTo errors
    
    hFunction = GetModuleHandle("kernel32.dll")
    hFunction = GetProcAddress(hFunction, "FreeLibrary")
    
    Do
        hThread = CreateRemoteThread(hProcess, lpThreadAttributes, 0, hFunction, hModule, 0, 0)
        uSucceed = 0
        GetExitCodeThread hThread, uSucceed
        uMax = uMax + 1
    Loop While (uSucceed = 1) And uMax < 256
    
    If uSucceed = 1 Or uSucceed = 259 Then errStr = "ж��ģ��ʧ��!": GoTo errors  '��������Է���1(�ɹ�)����˵��ģ��û�б�ж��
    
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

Public Function FxGetRemoteModuleName(ByVal PID As Long, ByVal ModuleName As String) As Long
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
    
    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, PID)
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

Public Sub MNNew(ByVal PID As Long, ByVal OwnerForm As ModuleList)
    With OwnerForm.ListView1
        If .Tag = 0 Then
            Call ListAllModules(PID, OwnerForm)
        ElseIf .Tag = 1 Then
            Call FxEnumModulesByVirtualMemory(PID, OwnerForm)
        ElseIf .Tag = 2 Then
            Call RdNewByReadMemory(PID, OwnerForm)
        End If
    End With
End Sub

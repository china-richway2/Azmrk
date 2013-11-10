Attribute VB_Name = "Process"
Option Explicit
Public Declare Function GetCurrentProcess Lib "kernel32" () As Long
Public Declare Function GetCurrentProcessId Lib "kernel32" () As Long
Public Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Public Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function Process32First Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function Process32Next Lib "kernel32" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Public Declare Function EnumProcesses Lib "psapi.dll" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Public Declare Function GetProcessImageFileName Lib "psapi.dll" Alias "GetProcessImageFileNameA" (ByVal hProcess As Long, ByVal lpImageFileName As String, ByVal nSize As Long) As Long
Public Declare Function SetPriorityClass Lib "kernel32" (ByVal hProcess As Long, ByVal dwPriorityClass As Long) As Long
Public Declare Function EndTask Lib "user32.dll" (ByVal hwnd As Long, ByVal fShutDown As Long, ByVal fForce As Long) As Long
Public Declare Function WinStationTerminateProcess Lib "winsta.dll" (ByVal hServer As Long, ByVal ProcessID As Long, ByVal ExitCode As Long) As Long
Public Declare Function ZwQueryInformationProcess Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, ByVal InformationClass As Long, ByRef ProcessInformation As Any, ByVal ProcessInformationLength As Long, ByRef ReturnLenght As Long) As Long
Public Declare Function ZwOpenProcess Lib "NTDLL.DLL" (ByRef ProcessHandle As Long, ByVal AccessMask As Long, ByRef ObjectAttributes As OBJECT_ATTRIBUTES, ByRef ClientId As CLIENT_ID) As Long
Public Declare Function ZwTerminateProcess Lib "NTDLL.DLL" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Public Declare Function ZwSuspendProcess Lib "NTDLL.DLL" (ByVal hProcess As Long) As Long
Public Declare Function ZwResumeProcess Lib "NTDLL.DLL" (ByVal hProcess As Long) As Long
Public Declare Function ZwDebugActiveProcess Lib "NTDLL.DLL" (ByVal ProcessHandle As Long, ByVal DebugObjectHandle As Long) As Long
Public Declare Sub ExitProcess Lib "kernel32" (ByVal uExitCode As Long)


Public Const STANDARD_RIGHTS_REQUIRED = &HF0000
Public Const SYNCHRONIZE = &H100000

Public Const PROCESS_TERMINATE = &H1
Public Const PROCESS_CREATE_THREAD = &H2
Public Const PROCESS_SET_SESSIONID = &H4
Public Const PROCESS_VM_OPERATION = &H8
Public Const PROCESS_VM_READ = &H10
Public Const PROCESS_VM_WRITE = &H20
Public Const PROCESS_DUP_HANDLE = &H40
Public Const PROCESS_CREATE_PROCESS = &H80
Public Const PROCESS_SET_QUOTA = &H100
Public Const PROCESS_SET_INFORMATION = &H200
Public Const PROCESS_QUERY_INFORMATION = &H400
Public Const PROCESS_SUSPEND_RESUME = &H800
Public Const PROCESS_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &HFFF)
'Public Const PROCESS_ALL_ACCESS As Long = &H1F0FFF '����Ȩ��

Public Const SMTO_ABORTIFHUNG = &H2
Public Const IDLE_PRIORITY_CLASS = &H40 '�½���Ӧ���зǳ��͵����ȼ�����ֻ����ϵͳ���е�ʱ��������С�����ֵ��4
Public Const HIGH_PRIORITY_CLASS = &H80  '�½����зǳ��ߵ����ȼ����������ڴ����Ӧ�ó��򡣻���ֵ��13��ע�⾡���������������ȼ�
Public Const NORMAL_PRIORITY_CLASS = &H20 '��׼���ȼ��������λ��ǰ̨�������ֵ��9�����ں�̨��������ֵ��7

Public Const DUPLICATE_CLOSE_SOURCE = &H1              '// winnt
Public Const DUPLICATE_SAME_ACCESS = &H2                  '// winnt
Public Const DUPLICATE_SAME_ATTRIBUTES = &H4

Public Const WTS_CURRENT_SERVER_HANDLE = 0

Public Const ZwGetCurrentProcess As Long = -1 '//0xFFFFFFFF


Public Enum PROCESSINFOCLASS
      ProcessBasicInformation
      ProcessQuotaLimits
      ProcessIoCounters
      ProcessVmCounters
      ProcessTimes
      ProcessBasePriority
      ProcessRaisePriority
      ProcessDebugPort
      ProcessExceptionPort
      ProcessAccessToken
      ProcessLdtInformation
      ProcessLdtSize
      ProcessDefaultHardErrorMode
      ProcessIoPortHandlers         '// Note: this is kernel mode only
      ProcessPooledUsageAndLimits
      ProcessWorkingSetWatch
      ProcessUserModeIOPL
      ProcessEnableAlignmentFaultFixup
      ProcessPriorityClass
      ProcessWx86Information
      ProcessHandleCount
      ProcessAffinityMask
      ProcessPriorityBoost
      ProcessDeviceMap
      ProcessSessionInformation
      ProcessForegroundInformation
      ProcessWow64Information
      ProcessImageFileName
      ProcessLUIDDeviceMapsEnabled
      ProcessBreakOnTermination
      ProcessDebugObjectHandle
      ProcessDebugFlags
      ProcessHandleTracing
      ProcessIoPriority
      ProcessExecuteFlags
      ProcessResourceManagement
      ProcessCookie
      ProcessImageInformation
      MaxProcessInfoClass           '// MaxProcessInfoClass should always be the last enum
End Enum


Public Type PROCESS_BASIC_INFORMATION
    ExitStatus As Long ' ���ս�����ֹ״̬
    PebBaseAddress As Long '���ս��̻������ַ
    AffinityMask As Long ' ���ս��̹�������
    BasePriority As Long ' ���ս��̵����ȼ���
    UniqueProcessId As Long ' ���ս���ID
    InheritedFromUniqueProcessId As Long '���ո�����ID
End Type

Public Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 1024
End Type

Public Type FX_PROCESS_INFORMATION
    ProcessID As Long
    ParentProcessId As Long
    peb As Long
    EPROCESS As Long
    Priority As Long
    MemoryUse As Long
    HighMemoryUse As Long
    ExePath As Long
End Type


Public nsItem As Long
Public OB_TYPE_PROCESS As Long


Public Sub FxListProcessBySession()
    Dim dwReturnLen As Long
    Dim etStart As Long
    Dim etLast As Long
    Dim etNow As Long
    Dim etNext As Long
    Dim tListProcess As LIST_ENTRY
    Dim tBListProcess As LIST_ENTRY
    Dim tFListProcess As LIST_ENTRY
    Dim hProcess As Long
    Dim pbi As PROCESS_BASIC_INFORMATION
    Dim pid As Long
    Dim EPROCESS As Long
    Dim pPath As String
    Dim pName As String
    Dim loopMax As Long

    etStart = FxAddSystemProcess
    etNext = etStart
    loopMax = 0
    Do
        pid = 0
        '��ȡPID
        ReadKernelMemory etNext + &H84, VarPtr(pid), Len(pid), dwReturnLen
        '���PID����������Ŀ
        If pid > 0 And pid < 65535 Then
            hProcess = FxNormalOpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, pid)
            ZwQueryInformationProcess hProcess, ProcessBasicInformation, pbi, Len(pbi), 0
            pPath = GetProcessPath(hProcess)
            pName = GetProcessName(pPath)
            
            Menu.ListView2.ListItems.Add , , FxGetProcessName(etNext)
            With Menu.ListView2.ListItems(Menu.ListView2.ListItems.Count)
                .SubItems(1) = CStr(pid)
                .SubItems(2) = CStr(pbi.InheritedFromUniqueProcessId)
                .SubItems(3) = FormatHex(pbi.PebBaseAddress)
                .SubItems(4) = FormatHex(etNext)
                .SubItems(5) = PriorityCheck(pbi.BasePriority)
                .SubItems(6) = FxGetProcessMemoryInformation(hProcess)
                .SubItems(7) = pPath
                .SubItems(8) = GetProcessCommandLine(hProcess)
            End With
            
            ZwClose hProcess: hProcess = 0
        End If
        '��ȡ���ڵ�LIST_ENTRY
        ReadKernelMemory etNext + &HB4, VarPtr(tListProcess), Len(tListProcess), dwReturnLen
        'MsgBox CStr(tListProcess.Blink) & "," & CStr(tListProcess.Flink)
        '����
        etNow = etNext
        '��һ����
        etLast = tListProcess.Flink - &HB4
        '��һ����
        etNext = tListProcess.Blink - &HB4

        loopMax = loopMax + 1
    Loop While loopMax < 65535 And (etNext <> etStart)
    
    'Menu.ListView2.ListItems (Menu.ListView2.ListItems.Count)
End Sub

Public Sub mpNew_Click()
    Dim ProcessInfo As PROCESSENTRY32
    Dim pbi As PROCESS_BASIC_INFORMATION
    Dim pc As Long
    Dim pm As Long
    Dim hProcess As Long
    Dim i As Long
    
    '��ʼ����
    pc = CreateToolhelp32Snapshot(TH32CS_SNAPPROCESS, 0)
    ProcessInfo.dwSize = Len(ProcessInfo)

    pm = Process32First(pc, ProcessInfo)
    Do While pm
        Menu.ListView2.ListItems.Add , , ProcessInfo.szExeFile   '���б��������
        
        hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, ProcessInfo.th32ProcessID)
        ZwQueryInformationProcess hProcess, ProcessBasicInformation, pbi, Len(pbi), 0
        
        With Menu.ListView2.ListItems(Menu.ListView2.ListItems.Count)
            .SubItems(1) = ProcessInfo.th32ProcessID
            .SubItems(2) = ProcessInfo.th32ParentProcessID
            .SubItems(3) = FormatHex(pbi.PebBaseAddress)
            '.SubItems(4) = EPROCESS   '�ɺ����FxGetProcessEProcess����������
            .SubItems(5) = PriorityCheck(ProcessInfo.pcPriClassBase)
            .SubItems(6) = FxGetProcessMemoryInformation(hProcess)
            .SubItems(7) = GetProcessPath(hProcess)
            .SubItems(8) = GetProcessCommandLine(hProcess)
        End With
        
        ZwClose hProcess: hProcess = 0
        
        pm = Process32Next(pc, ProcessInfo)
    Loop
    
    ZwClose pc
    
    FxGetProcessEProcess Menu.ListView2, 1, 4   '��FxGetProcessEProcess��������б��е�EPROCESS��
End Sub

Public Sub ListProcessByWmi()
    Dim objSWbemLocator As New SWbemLocator
    Dim objSWbemServices As SWbemServices
    Dim objSWbemObjectSet As SWbemObjectSet
    Dim objSWbemObject As SWbemObject
    Dim i As Long
    Dim pIndex As Long
    
    pIndex = 1
    
    '��ձ�
    pIndex = FxGetListviewNowLine(Menu.ListView2)
    
    Menu.ListView2.Tag = 2
    
    Menu.ListView2.ListItems.Clear '���ListView
    Set objSWbemServices = objSWbemLocator.ConnectServer()  '���ӵ�������WMI������һ���� SWbemServices ���������
    Set objSWbemObjectSet = objSWbemServices.InstancesOf("Win32_Process")   '����Win32_Process������ʶ������ʵ��
    i = 0
    For Each objSWbemObject In objSWbemObjectSet  'ö��ÿһ��Win32_Process��ʵ��
        Menu.ListView2.ListItems.Add , "a" & i, objSWbemObject.Name '������ID��ӵ�ListView1��һ��
        With Menu.ListView2.ListItems("a" & i)
            .SubItems(1) = objSWbemObject.Handle '����������ӵ�ListView1�ڶ���
            .SubItems(2) = FxGetProcessMemoryInformation(objSWbemObject.Handle)
        End With
        If Not IsNull(objSWbemObject.ExecutablePath) Then Menu.ListView2.ListItems("a" & i).SubItems(3) = objSWbemObject.ExecutablePath '������·����ӵ�ListView1������
        i = i + 1
    Next
    Set objSWbemObjectSet = Nothing
End Sub

Public Sub ListProcessHf()
    'ͨ��PSAPI.DLL���EnumProcesses����������,Ч��ͬToolhelp32ϵ��,����,��ʹ��
    Dim pid(1024) As Long
    Dim prCount As Long
    Dim i As Integer
    Dim pIndex As Integer
    
    pIndex = 1
    
    If Menu.ListView2.ListItems.Count > 0 And Menu.ListView2.Tag = 1 Then
        pIndex = Menu.ListView2.SelectedItem.Index
    End If
    If Menu.ListView2.Sorted = True Then Menu.ListView2.Sorted = False
    
    Menu.ListView2.Tag = 1

    Menu.ListView2.ListItems.Clear
    EnumProcesses pid(0), 1024, prCount
    For i = 0 To prCount / 4 - 1
        'ListView2.ListItems.Add , , pID(i)
        
    Next i
End Sub

Public Function FxAddSystemProcess()
    Dim EPROCESS As Long
    Dim ret() As Long
    Dim hModule As Long
    Dim PsInitialSystemProcess As Long
    Dim lngSList As Long
    Dim lngAList As Long
    Dim etStart As Long
    Dim i As Integer
    
    Menu.ListView2.ListItems.Add , , "Idle"
    With Menu.ListView2.ListItems(1)
        .SubItems(1) = 0
        .SubItems(2) = 0
    End With

    lngSList = 180: lngAList = 136 'XPӲ����
    
    hModule = LoadLibraryEx(GetDeviceDriver(BaseName), 0, 1)
    PsInitialSystemProcess = GetProcAddress(hModule, "PsInitialSystemProcess")
    PsInitialSystemProcess = PsInitialSystemProcess + GetDeviceDriver(BaseAddress) - hModule
    FreeLibrary hModule
    
    'System
    ReadKernelMemory ByVal PsInitialSystemProcess, ByVal VarPtr(EPROCESS), 4, 0
    ReDim Preserve ret(0)
    ret(0) = EPROCESS
    'MsgBox "System EPROCESS:" & FormatHex(EPROCESS)
    
    'smss.exe
    ReadKernelMemory ByVal (EPROCESS + lngAList), ByVal VarPtr(EPROCESS), 4, 0
    EPROCESS = EPROCESS - lngAList
    ReDim Preserve ret(1)
    ret(1) = EPROCESS
    'MsgBox "smss.exe EPROCESS:" & FormatHex(EPROCESS)
    
    Dim pid As Long
    Dim hProcess As Long
    Dim pbi As PROCESS_BASIC_INFORMATION
    Dim pPath As String
    Dim pName As String
    
    
    For i = 0 To 1
        ReadKernelMemory ByVal ret(i) + &H84, ByVal VarPtr(pid), 4, 0
        hProcess = FxNormalOpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ, pid)
        ZwQueryInformationProcess hProcess, ProcessBasicInformation, pbi, Len(pbi), 0
        pPath = GetProcessPath(hProcess)
        pName = GetProcessName(pPath)
            
        Menu.ListView2.ListItems.Add , , pName
        With Menu.ListView2.ListItems(Menu.ListView2.ListItems.Count)
            .SubItems(1) = CStr(pid)
            .SubItems(2) = CStr(pbi.InheritedFromUniqueProcessId)
            .SubItems(3) = FormatHex(pbi.PebBaseAddress)
            .SubItems(4) = FormatHex(ret(i))
            .SubItems(5) = PriorityCheck(pbi.BasePriority)
            .SubItems(6) = FxGetProcessMemoryInformation(hProcess)
            .SubItems(7) = pPath
            .SubItems(8) = GetProcessCommandLine(hProcess)
        End With
            
        ZwClose hProcess: hProcess = 0
    Next i
    
    ReadKernelMemory ByVal (EPROCESS + lngAList), ByVal VarPtr(etStart), 4, 0
    FxAddSystemProcess = etStart - lngAList
    'MsgBox "etStart:" & FormatHex(etStart)
End Function

Public Function PriorityCheck(ByVal pcb As Long) As String
    '/**��������:�жϽ������ȼ��������ַ���**/
    Select Case pcb
    Case Is > 9
        PriorityCheck = "�ϸ�" & "[" & (pcb) & "]"
    Case Is >= 7
        PriorityCheck = "��׼" & "[" & (pcb) & "]"
    Case Is >= 4
        PriorityCheck = "�ϵ�" & "[" & (pcb) & "]"
    Case Else
        PriorityCheck = "����" & "[" & (pcb) & "]"
    End Select
End Function

Public Function GetProcessState(ByVal frmhWnd As Long, Optional Timeout As Long = 20) As String
    Dim Results As Long

    If Not SendMessageTimeout(frmhWnd, ByVal 0, ByVal 0, ByVal 0, SMTO_ABORTIFHUNG, Timeout, Results) = 1 Then
        'If Results = 0 Then GetState = True
        GetProcessState = "����"
    Else
        GetProcessState = "����"
    End If
End Function

Public Function GetProcessPath(ByVal hProcess As Long) As String
    '/**��������:�ɽ��̾����ȡ����·��**/
    On Error Resume Next

    Dim hModule As Long
    Dim ret As Long
    Dim szPathName As String

    If hProcess <> 0 Then
        ret = EnumProcessModules(hProcess, hModule, 4, 0)
        If ret <> 0 Then
            szPathName = Space(260)
            ret = GetModuleFileNameEx(hProcess, hModule, szPathName, 260)
            GetProcessPath = left(szPathName, ret)
        End If
    End If

    If GetProcessPath = "" Then
        GetProcessPath = "System"
    End If
End Function

Public Function GetProcessCommandLine(ByVal hProcess As Long) As String
    '/**��������:��PID��ȡ����������**/
    Dim ntStatus As Long
    Dim objBasic As PROCESS_BASIC_INFORMATION
    Dim objFlink As Long
    Dim objPEB As Long, objLdr As Long
    Dim objBaseAddress As Long
    Dim bytName(260 * 2 - 1) As Byte
    Dim strModuleName As String
    Dim objName As Long
    
    If hProcess = 0 Then
        GetProcessCommandLine = ""
        Exit Function
    End If
           
    Dim lngRet As Long, lngReturn As Long
    
    ntStatus = ZwQueryInformationProcess(hProcess, ProcessBasicInformation, objBasic, Len(objBasic), ByVal 0&)
    If (NT_SUCCESS(ntStatus)) Then
        '��ȡPEBָ��
        objPEB = objBasic.PebBaseAddress
        '��ȡ_RTL_USER_PROCESS_PARAMETERS�ṹָ��
        lngRet = ReadProcessMemory(hProcess, ByVal objPEB + &H10, objLdr, 4, ByVal 0&)
        If lngRet <> 1 Then Exit Function
        '��ȡ·��ָ��
        lngRet = ReadProcessMemory(hProcess, ByVal objLdr + &H44, objName, 4, ByVal 0&)
        If lngRet <> 1 Then Exit Function
        '��ȡ·��
        lngRet = ReadProcessMemory(hProcess, ByVal objName, bytName(0), 260 * 2, ByVal 0&)
        If lngRet <> 1 Then Exit Function
        strModuleName = bytName
        If InStr(strModuleName, """") = 0 Then
            strModuleName = Mid(strModuleName, InStr(strModuleName, Chr(0)) + 1, Len(strModuleName) - InStr(strModuleName, Chr(0)))
            'strModuleName = SetPath(strModuleName)
        Else
            strModuleName = Mid(strModuleName, InStr(strModuleName, """"), Len(strModuleName) - InStr(strModuleName, """"))
        End If
        strModuleName = left(strModuleName & Chr(0), InStr(strModuleName & Chr(0), Chr(0)) - 1)
        GetProcessCommandLine = strModuleName
    End If
End Function

Public Function GetProcessName(ByVal Path As String, Optional ByVal FindText As String = "\") As String
    '/**��������:�ɽ���·����ȡ������**/
    GetProcessName = Mid$(Path, InStrRev(Path, FindText) + 1)
End Function

Public Function FxGetProcessName(ByVal EPROCESS As Long) As String
    Dim proName As String * MAX_PATH
    Dim byBuff(MAX_PATH) As Byte
    
    ReadKernelMemory EPROCESS + &H174, VarPtr(byBuff(0)), MAX_PATH, 0
    FxGetProcessName = StrConv(byBuff(), vbUnicode)
End Function

Public Function FxNormalOpenProcess(ByVal dwDesiredAccess As Long, ByVal pid As Long) As Long
    '/**��������:��һ�����̣�ʧ�������LzOpenProcess**/
    Dim oa As OBJECT_ATTRIBUTES
    Dim Cid As CLIENT_ID
    Dim st As Long
    Dim hProcess As Long
    
    oa.Length = LenB(oa)

    Cid.UniqueProcess = pid

    st = ZwOpenProcess(hProcess, dwDesiredAccess, oa, Cid)
    'hProcess = OpenProcess(dwDesiredAccess, False, pid)
    If Not NT_SUCCESS(st) Then
        hProcess = LzOpenProcess(dwDesiredAccess, pid)
    End If

    FxNormalOpenProcess = hProcess
End Function

Public Function LzOpenProcess(ByVal dwDesiredAccess As Long, ByVal ProcessID As Long) As Long
    '/**��������:ͨ�����ƾ������ľ�������򿪡�����**/
    Dim st As Long
    Dim Cid As CLIENT_ID
    Dim oa As OBJECT_ATTRIBUTES
    Dim NumOfHandle As Long
    Dim pbi As PROCESS_BASIC_INFORMATION
    Dim i As Long
    Dim hProcessToDup As Long, hProcessCur As Long, hProcessToRet As Long
    
    oa.Length = Len(oa)
    '���ȳ���ZwOpenProcess
    Cid.UniqueProcess = ProcessID
    st = ZwOpenProcess(hProcessToRet, dwDesiredAccess, oa, Cid)
    If (NT_SUCCESS(st)) Then LzOpenProcess = hProcessToRet: Exit Function
    st = 0
    
    Dim bytBuf() As Byte
    Dim arySize As Long: arySize = 1
    Do
        ReDim bytBuf(arySize)
        st = ZwQuerySystemInformation(SystemHandleInformation, VarPtr(bytBuf(0)), arySize, 0&)
        If (Not NT_SUCCESS(st)) Then
            If (st <> STATUS_INFO_LENGTH_MISMATCH) Then
                Erase bytBuf
                Exit Function
            End If
        Else
            Exit Do
        End If
        arySize = arySize * 2
        ReDim bytBuf(arySize)
    Loop
    
    NumOfHandle = 0
    CopyMemory VarPtr(NumOfHandle), VarPtr(bytBuf(0)), Len(NumOfHandle)
    Dim h_info() As SYSTEM_HANDLE_TABLE_ENTRY_INFO
    ReDim h_info(NumOfHandle)
    CopyMemory VarPtr(h_info(0)), VarPtr(bytBuf(0)) + Len(NumOfHandle), Len(h_info(0)) * NumOfHandle
    
    '//ö�پ����ɣ�������ʼ���Ծ��
    For i = LBound(h_info) To UBound(h_info)
        With h_info(i)
            If (.ObjectTypeIndex = OB_TYPE_PROCESS) Then
                Cid.UniqueProcess = .UniqueProcessId
                st = ZwOpenProcess(hProcessToDup, PROCESS_DUP_HANDLE, oa, Cid)
                If (NT_SUCCESS(st)) Then
                    st = ZwDuplicateObject(hProcessToDup, .HandleValue, ZwGetCurrentProcess, hProcessCur, PROCESS_ALL_ACCESS, 0, DUPLICATE_SAME_ATTRIBUTES)
                    If (NT_SUCCESS(st)) Then
                        st = ZwQueryInformationProcess(hProcessCur, ProcessBasicInformation, pbi, Len(pbi), 0)
                        If (NT_SUCCESS(st)) Then
                            If (pbi.UniqueProcessId = ProcessID) Then
                                st = ZwDuplicateObject(hProcessToDup, .HandleValue, ZwGetCurrentProcess, hProcessToRet, dwDesiredAccess, 0, DUPLICATE_SAME_ATTRIBUTES)
                                If (NT_SUCCESS(st)) Then
                                    If hProcessToRet <> 0 Then
                                        LzOpenProcess = hProcessToRet
                                        Exit Function
                                    End If
                                End If
                            End If
                        End If
                    End If
                    st = ZwClose(hProcessCur)
                End If
                st = ZwClose(hProcessToDup)
            End If
        End With
    Next i
    
    Erase h_info
End Function

Public Function FxGetProcessId(ByVal hProcess As Long) As Long
    '/**��������:�ɽ��̾����ȡPID**/
    Dim pbi As PROCESS_BASIC_INFORMATION
    Dim st As Long
    
    st = ZwQueryInformationProcess(hProcess, ProcessBasicInformation, pbi, Len(pbi), 0)
    If (NT_SUCCESS(st)) Then
        FxGetProcessId = pbi.UniqueProcessId
    End If
End Function

Public Function FxGetObjectTypeProcess() As Long
    '/**��������:��ȡ���̵ľ�����͵�����**/
    Dim mHandle, mPid As Long
    Dim st As Long
       
    mPid = GetCurrentProcessId
    
    st = ZwDuplicateObject(GetCurrentProcess, GetCurrentProcess, GetCurrentProcess, mHandle, PROCESS_ALL_ACCESS, 0, DUPLICATE_SAME_ATTRIBUTES)
    
    If NT_SUCCESS(st) Then
        Dim bytBuf() As Byte
        Dim arySize As Long
        
        arySize = 1
        Do
            ReDim bytBuf(arySize)
            st = ZwQuerySystemInformation(SystemHandleInformation, VarPtr(bytBuf(0)), arySize, 0&)
            If (Not NT_SUCCESS(st)) Then
                If (st <> STATUS_INFO_LENGTH_MISMATCH) Then
                    Erase bytBuf
                    Exit Function
                End If
            Else
                Exit Do
            End If
            arySize = arySize * 2
            ReDim bytBuf(arySize)
        Loop
        
        Dim NumOfHandle As Long
        NumOfHandle = 0
        CopyMemory VarPtr(NumOfHandle), VarPtr(bytBuf(0)), Len(NumOfHandle)
        Dim h_info() As SYSTEM_HANDLE_TABLE_ENTRY_INFO
        ReDim h_info(NumOfHandle)
        CopyMemory VarPtr(h_info(0)), VarPtr(bytBuf(0)) + Len(NumOfHandle), Len(h_info(0)) * NumOfHandle
        
        Dim i As Long
        For i = 1 To NumOfHandle
            If h_info(i).HandleValue = mHandle And h_info(i).UniqueProcessId = mPid Then
                FxGetObjectTypeProcess = h_info(i).ObjectTypeIndex
                Exit For
            End If
        Next i
    End If
End Function

Public Sub FxGetProcessEProcess(ByRef Listview As Object, ByVal pidColumn As Long, ByVal epColumn As Long)
    '/**��������:���Lsitview�е�EPROCESS��**/
    Dim bytBuf() As Byte
    Dim arySize As Long
        Dim st As Long
        
    arySize = 1
    Do
        ReDim bytBuf(arySize)
        st = ZwQuerySystemInformation(SystemHandleInformation, VarPtr(bytBuf(0)), arySize, 0&)
        If (Not NT_SUCCESS(st)) Then
            If (st <> STATUS_INFO_LENGTH_MISMATCH) Then
                Erase bytBuf
                Exit Sub
            End If
        Else
            Exit Do
        End If
        arySize = arySize * 2
        ReDim bytBuf(arySize)
    Loop
        
    Dim NumOfHandle As Long
    NumOfHandle = 0
    CopyMemory VarPtr(NumOfHandle), VarPtr(bytBuf(0)), Len(NumOfHandle)
    Dim h_info() As SYSTEM_HANDLE_TABLE_ENTRY_INFO
    ReDim h_info(NumOfHandle)
    CopyMemory VarPtr(h_info(0)), VarPtr(bytBuf(0)) + Len(NumOfHandle), Len(h_info(0)) * NumOfHandle
    
    Dim i, j As Long
    Dim nowPid As Long
    
    For i = LBound(h_info) To UBound(h_info) / 4
        With h_info(i)
            If .ObjectTypeIndex = OB_TYPE_PROCESS Then
                nowPid = PsGetPidByEProcess(.pObject)
                For j = 1 To Listview.ListItems.Count
                    If Listview.ListItems(j).SubItems(pidColumn) = nowPid And Listview.ListItems(j).SubItems(epColumn) = "" Then
                        Listview.ListItems(j).SubItems(epColumn) = FormatHex(.pObject)
                        Exit For
                    End If
                Next j
            End If
        End With
    Next i
    
    Erase h_info
End Sub

Public Function PsGetImageFileNameByEProcess(ByVal EPROCESS As Long) As Long
    '/**��������:��EPROCESS��ȡ������**/
    ReadKernelMemory EPROCESS + &H174, VarPtr(PsGetImageFileNameByEProcess), 4, 0
End Function

Public Function PsGetPidByEProcess(ByVal EPROCESS As Long) As Long
    '/**��������:��EPROCESS��ȡPID**/
    ReadKernelMemory EPROCESS + &H84, VarPtr(PsGetPidByEProcess), 4, 0
End Function

Public Function FxGetCurrentEProcess() As Long
    '/**��������:��ȡ�����EPROCESS**/
    Dim mHandle As Long
    Dim dwPid As Long
    Dim st As Long
       
    dwPid = GetCurrentProcessId
    mHandle = OpenProcess(PROCESS_QUERY_INFORMATION, False, dwPid)
    
    If NT_SUCCESS(st) Then
        Dim bytBuf() As Byte
        Dim arySize As Long
        arySize = 1
        Do
            ReDim bytBuf(arySize)
            st = ZwQuerySystemInformation(16, VarPtr(bytBuf(0)), arySize, 0&)
            If (Not NT_SUCCESS(st)) Then
                If (st <> &HC0000004) Then
                    Erase bytBuf
                    Exit Function
                End If
            Else
                Exit Do
            End If
            arySize = arySize * 2
            ReDim bytBuf(arySize)
        Loop
        
        Dim NumOfHandle As Long
        NumOfHandle = 0
        CopyMemory ByVal VarPtr(NumOfHandle), ByVal VarPtr(bytBuf(0)), Len(NumOfHandle)
        Dim h_info() As SYSTEM_HANDLE_TABLE_ENTRY_INFO
        ReDim h_info(NumOfHandle)
        CopyMemory ByVal VarPtr(h_info(0)), ByVal VarPtr(bytBuf(0)) + Len(NumOfHandle), Len(h_info(0)) * NumOfHandle

        Dim i As Long
        For i = 0 To NumOfHandle
            If h_info(i).HandleValue = mHandle And h_info(i).UniqueProcessId = dwPid Then
                FxGetCurrentEProcess = h_info(i).pObject
                Exit For
            End If
        Next i
    End If
End Function

Public Sub FxTerminateProcessByDebugProcess(ByVal pid As Long)
    '/**��������:ͨ�����Խ��̵ķ�����������**/
    Dim hDebug As Long
    Dim hProcess As Long
    Dim status As Long
    Dim errStr As String
       
    '�������Զ���
    If Not NT_SUCCESS(ZwCreateDebugObject(hDebug, &H1F000F, 0&, 1&)) Then errStr = "�������Զ���ʧ��!": GoTo errors

    '��õ��Ծ��
    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, pid)
    If hProcess <= 0 Then ZwClose hDebug: errStr = "�ܾ�����!": GoTo errors
    
    '�ӹܵ��Խ���Ȼ���˳�
    status = ZwDebugActiveProcess(hProcess, hDebug)
    ZwClose hProcess
    ZwClose hDebug
    
    '�ж��Ƿ�ɹ�
    If Not NT_SUCCESS(status) Then errStr = "���Խ���ʧ��!": GoTo errors
Exit Sub
errors:
    MsgBox errStr, 0, "ʧ��"
End Sub

Public Sub PNNew()
    '/**��������:�����жϱ������̷�����ˢ��Lsitview(ˢ���б�ʱ��ʹ�ô˺���)**/
    Dim pIndex As Long
    
    If Menu.ListView2.Sorted = True Then Menu.ListView2.Sorted = False
    
    pIndex = FxGetListviewNowLine(Menu.ListView2)
    
    Menu.ListView2.ListItems.Clear
    
    'LockWindowUpdate Menu.ListView2.hwnd
    
    If Menu.ListView2.Tag = 0 Then
        Call mpNew_Click
    ElseIf Menu.ListView2.Tag = 1 Then
        Call FxListProcessBySession
    ElseIf Menu.ListView2.Tag = 2 Then
        Call ListProcessByWmi
    End If
    
    FxSetListviewNowLine Menu.ListView2, pIndex
    
    'LockWindowUpdate 0
    Menu.Label3.Caption = "����" & (Menu.ListView2.ListItems.Count) & "������"
End Sub

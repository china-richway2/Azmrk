Attribute VB_Name = "FxExecRemoteFunc"
Option Explicit

Private Declare Function VirtualAllocEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal flAllocationType As Long, ByVal flProtect As Long) As Long
Private Declare Function VirtualFreeEx Lib "kernel32" (ByVal hProcess As Long, ByVal lpAddress As Long, ByVal dwSize As Long, ByVal dwFreeType As Long) As Long
Private Declare Function ReadProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Function WriteProcessMemory Lib "kernel32" (ByVal hProcess As Long, lpBaseAddress As Any, lpBuffer As Any, ByVal nSize As Long, lpNumberOfBytesWritten As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (Destination As Long, Source As Long, ByVal Length As Long)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (Destination As Long, ByVal Length As Long)

Private Declare Function LoadLibrary Lib "kernel32" Alias "LoadLibraryA" (ByVal lpLibFileName As String) As Long
Private Declare Function LoadLibraryEx Lib "kernel32.dll" Alias "LoadLibraryExA" (ByVal lpLibFileName As String, ByVal hFile As Long, ByVal dwFlags As Long) As Long
Private Declare Function FreeLibrary Lib "kernel32.dll" (ByVal hLibModule As Long) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long

Private Declare Function NtQueryInformationProcess Lib "NTDLL.dll" (ByVal ProcessHandle As Long, ByVal ProcessInformationClass As Long, ByVal ProcessInformation As Long, ByVal ProcessInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function CreateRemoteThread Lib "kernel32" (ByVal hProcess As Long, lpThreadAttributes As SECURITY_ATTRIBUTES, ByVal dwStackSize As Long, lpStartAddress As Long, lpParameter As Long, ByVal dwCreationFlags As Long, lpThreadId As Long) As Long
Private Declare Function GetExitCodeThread Lib "kernel32" (ByVal hThread As Long, lpExitCode As Long) As Long

Private Declare Function WaitForSingleObject Lib "kernel32" (ByVal hHandle As Long, ByVal dwMilliseconds As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal Handle As Long) As Long


Private Const MEM_FREE = &H10000
Private Const MEM_Private = &H20000
Private Const MEM_COMMIT = 4096
Private Const MEM_RESERVE = &H2000
Private Const MEM_DECOMMIT = &H4000
Private Const MEM_RELEASE = &H8000

Private Const PAGE_READONLY = &H2
Private Const PAGE_READWRITE = &H4
Private Const PAGE_EXECUTE_READWRITE = &H40
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SYNCHRONIZE = &H100000

Private Type LIST_ENTRY
    Blink                           As Long
    Flink                           As Long
End Type

Private Type UNICODE_STRING
    Length                          As Integer
    MaximumLength                   As Integer
    Buffer                          As Long
End Type

Private Type LDR_MODULE 'LDR_DATA_TABLE_ENTRY
    InLoadOrderModuleList           As LIST_ENTRY
    InMemoryOrderModuleList         As LIST_ENTRY
    InInitializationOrderModuleList As LIST_ENTRY
    BaseAddress                     As Long
    EntryPoint                      As Long
    SizeOfImage                     As Long
    FullDllName                     As UNICODE_STRING
    BaseDllName                     As UNICODE_STRING
    Flags                           As Long
    LoadCount                       As Integer
    TlsIndex                        As Integer
    HashTableEntry                  As LIST_ENTRY
    TimeDateStamp                   As Long
End Type

Private Type PEB_LDR_DATA
    Length                          As Long
    Initialized                     As Long
    SsHandle                        As Long
    InLoadOrderModuleList           As LIST_ENTRY
    InMemoryOrderModuleList         As LIST_ENTRY
    InInitializationOrderModuleList As LIST_ENTRY
End Type

Private Type PROCESS_BASIC_INFORMATION
    ExitStatus                      As Long
    PEBBaseAddress                  As Long
    AffinityMask                    As Long
    BasePriority                    As Long
    UniqueProcessId                 As Long
    InheritedFromUniqueProcessId    As Long
End Type

Private Type SECURITY_ATTRIBUTES
    nLength As Long
    lpSecurityDescriptor As Long
    bInheritHandle As Long
End Type

Private Function FxGetModuleHandle(ByVal hProcess As Long, ByRef modPath As String) As Long
    Dim lPtr            As Long
    Dim tPEB_LDR_DATA   As PEB_LDR_DATA
    Dim tLDR_MODULE     As LDR_MODULE
    Dim tPBI    As PROCESS_BASIC_INFORMATION
    Dim bytDllName(259)    As Byte

    FxGetModuleHandle = 0
    If hProcess Then
        If NtQueryInformationProcess(hProcess, 0, VarPtr(tPBI), Len(tPBI), 0) = 0 Then
            lPtr = tPBI.PEBBaseAddress
        End If

        If Not ReadProcessMemory(hProcess, ByVal lPtr + 12, lPtr, &H4, 0&) = 0 Then
            ReadProcessMemory hProcess, ByVal lPtr, ByVal VarPtr(tPEB_LDR_DATA), Len(tPEB_LDR_DATA), 0
            ReadProcessMemory hProcess, ByVal tPEB_LDR_DATA.InLoadOrderModuleList.Flink, ByVal VarPtr(tLDR_MODULE), Len(tLDR_MODULE), 0
            Do While tLDR_MODULE.BaseAddress <> 0
                ZeroMemory ByVal VarPtr(bytDllName(0)), 260
                ReadProcessMemory hProcess, ByVal tLDR_MODULE.FullDllName.Buffer, ByVal VarPtr(bytDllName(0)), tLDR_MODULE.FullDllName.Length, 0
                'Debug.Print Hex(tLDR_MODULE.BaseAddress) & "," & StrConv(bytDllName, vbNarrow) & "," & CStr(tLDR_MODULE.BaseDllName.MaximumLength)
                If modPath = left(StrConv(bytDllName, vbNarrow), Len(modPath)) Then
                    FxGetModuleHandle = tLDR_MODULE.BaseAddress
                    Exit Do
                End If
                ReadProcessMemory hProcess, ByVal tLDR_MODULE.InLoadOrderModuleList.Flink, ByVal VarPtr(tLDR_MODULE), Len(tLDR_MODULE), 0
            Loop
        End If
    End If
End Function

Public Function FxExecuteRemoteFunction(ByVal hProcess As Long, ByRef modPath As String, ByRef funName As String, ParamArray Params()) As Long
'/*函数功能：远程进程内执行任意模块内的任意函数*/
'/*使用方法：参数1传远程进程的句柄，参数2传可执行代码所在模块的绝对路径，参数3传要执行的函数，参数4-n传函数参数*/
'/*模块作者：Naylon [http://hi.baidu.com/naylonslain]，转载请注明原作者信息*/
'/*修改时间：2010-09-10*/

'--||初始化||--
    Dim errStr As String
    If hProcess = 0 Then errStr = "参数不正确": GoTo errors
    
    Dim i As Long
    Dim dwRet As Long   '代表参数是否处理成功
    Dim pamCount As Long   '代表参数总数-1的值（因为数组下标是0）
    Dim pamAddr() As Long   '用来记录每个参数的值（String记录地址）
    Dim pamType As Long
    
    pamCount = UBound(Params)
    ReDim pamAddr(pamCount)
    
'--||从右至左（stdcall）将参数分类写入目标进程的地址空间，pamAddr数组记录每个参数的值（对于String是地址）||--
    For i = pamCount To 0 Step -1
        pamType = VarType(Params(i))
        If pamType = vbString Then
            Dim strData As String
            strData = CStr(Params(i))
            If strData = "" Then
                pamAddr(i) = 0
                dwRet = 1
            Else
                Dim strSize As Long
                strSize = LenB(StrConv(strData, vbFromUnicode)) + 1
                pamAddr(i) = VirtualAllocEx(hProcess, 0, strSize, MEM_COMMIT, PAGE_READWRITE)
                WriteProcessMemory hProcess, ByVal pamAddr(i), ByVal strData, strSize, dwRet
            End If
        ElseIf pamType = vbBoolean Or pamType = vbByte Or pamType = vbInteger Or pamType = vbLong Then
            pamAddr(i) = CLng(Params(i))
            dwRet = 1
        Else
            errStr = "参数" & CStr(pamCount - i + 1) & "不支持的类型": GoTo errors
        End If
    
        If dwRet = 0 Then
            errStr = "参数" & CStr(pamCount - i + 1) & "写入失败": GoTo errors
        Else
            'Debug.Print "参数" & CStr(pamCount - i + 1) & "成功写入，地址" & FormatHex(pamAddr(i))
        End If
    Next i
    
'--||准备工作||--
    '--计算shellcode大小（占用字节）--
    Dim scSize As Long
    scSize = (pamCount + 1) * 5   '每个参数都要push 0x00000000，占5字节
    scSize = scSize + 5 + 1   '调用函数用call 0x0000000，占5字节;call之后要ret，占1字节
    Dim sc() As Byte
    ReDim sc(scSize - 1)
    '--push参数入栈--
    Dim j As Long
    i = 0: j = 0
    For i = pamCount To 0 Step -1
        sc(j) = &H68   'push
        CopyMemory ByVal VarPtr(sc(j + 1)), ByVal VarPtr(pamAddr(i)), 4
        j = j + 5
    Next i

    '--获取函数信息--
    '获取模块基址
    Dim hLocalModule As Long
    hLocalModule = LoadLibrary(modPath)
    If hLocalModule = 0 Then errStr = "加载模块失败（Local）": GoTo errors
    '获取函数地址
    Dim funcAddr As Long
    funcAddr = GetProcAddress(hLocalModule, funName)
    '计算函数偏移
    Dim funcOffset As Long
    funcOffset = funcAddr - hLocalModule
    '卸载模块
    FreeLibrary hLocalModule
    '获取远程模块地址
    Dim hRemoteModule As Long
    hRemoteModule = FxGetModuleHandle(hProcess, modPath)
    If hRemoteModule = 0 Then
        '递归,注入DLL
        hRemoteModule = FxExecuteRemoteFunction(hProcess, Environ("windir") & "\system32\kernel32.dll", "LoadLibraryA", modPath)
        If hRemoteModule = 0 Then errStr = "加载模块失败（Remote）": GoTo errors
    End If
    '模块基址 + 函数偏移 = 函数地址
    funcAddr = hRemoteModule + funcOffset
    
'--||构造shellcode||--
    '为shellcode分配内存空间
    Dim codeAddr As Long
    codeAddr = VirtualAllocEx(hProcess, 0, scSize, MEM_COMMIT, PAGE_EXECUTE_READWRITE)
    '计算call的偏移
    sc(j) = &HE8   'call
    Dim callOffset As Long
    callOffset = funcAddr - codeAddr - (scSize - 1)
    CopyMemory ByVal VarPtr(sc(j + 1)), ByVal VarPtr(callOffset), 4
    sc(j + 5) = &HC3   'ret
    '--写入shellcode--
    WriteProcessMemory hProcess, ByVal codeAddr, ByVal VarPtr(sc(0)), scSize, dwRet
    'Debug.Print "shellcode地址" & FormatHex(codeAddr)
    If dwRet = 0 Then errStr = "写入shellcode失败": GoTo errors
    
'--||创建线程执行shellcode||--
    Dim sa As SECURITY_ATTRIBUTES
    Dim hThreadRet As Long
    hThreadRet = CreateRemoteThread(hProcess, sa, 0, ByVal codeAddr, ByVal 0, 0, 0)
    If hThreadRet = 0 Then errStr = "执行shellcode失败": GoTo errors
    'WaitForSingleObject hThreadRet, INFINITE   '等待线程执行结束
    'GetExitCodeThread hThreadRet, FxExecuteRemoteFunction   '获取函数的返回值
    
    VirtualFreeEx hProcess, codeAddr, scSize, MEM_DECOMMIT
Exit Function

errors:
    Debug.Print errStr
    FxExecuteRemoteFunction = 0
End Function


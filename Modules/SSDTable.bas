Attribute VB_Name = "SSDTable"
Option Explicit
Public pKernel As Long   '用户态的ntkrnlpa.exe基址
Dim pNtDLL As Long       '用户态的ntdll.dll基址
Dim pBaseKernel As Long  '内核态的ntkrnlpa.exe基址
Dim nWin32K() As Byte    'win32k.sys内容
Dim pWin32K As Long      '指向nWin32K
Dim Win32KBase As Long   '内核态的win32k.sys基址
Public Type SSDT_ENTRY
    strName As String
    dwCurrAddress As Long
    dwRealAddress As Long
End Type
Public SSDTData() As SSDT_ENTRY
Public ShadowSSDTData() As SSDT_ENTRY
'typedef struct _SYSTEM_SERVICE_TABLE
'{
'  PVOID   ServiceTableBase;      // SSDT (System Service Dispatch Table)的基地址
'  PULONG  ServiceCounterTableBase;  // 用于checked builds, 包含SSDT中每个服务被调用的次数
'  ULONG   NumberOfService;      // 服务函数的个数, NumberOfService*4 就是整个地址表的大小
'  ULONG   ParamTableBase;        // SSPT (System Service Parameter Table)的基地址
'} SYSTEM_SERVICE_TABLE, *PSYSTEM_SERVICE_TABLE;
'
'typedef struct _SERVICE_DESCRIPTOR_TABLE
'{
'  SYSTEM_SERVICE_TABLE   ntoskrnl;  // ntoskrnl.exe的服务函数
'  SYSTEM_SERVICE_TABLE   win32k;    // win32k.sys的服务函数,(gdi.dll/user.dll的内核支持)
'  SYSTEM_SERVICE_TABLE   NotUsed1;
'  SYSTEM_SERVICE_TABLE   NotUsed2;
'} SYSTEM_DESCRIPTOR_TABLE, *PSYSTEM_DESCRIPTOR_TABLE;

Public Sub InitSSDTableModule()
    pNtDLL = GetModuleHandle("ntdll.dll")
    Dim i As Long
    Dim A() As Byte
    ZwQuerySystemInformation SystemModuleInformation, 0, 0, ByVal VarPtr(i)
    ReDim A(i - 1)
    ZwQuerySystemInformation SystemModuleInformation, VarPtr(A(0)), i, i
    Dim m() As ModuleInformation
    CopyMemory VarPtr(i), VarPtr(A(0)), 4
    ReDim m(i - 1)
    CopyMemory VarPtr(m(0)), VarPtr(A(4)), Len(m(0)) * i
    Erase A
    pKernel = LoadLibrary("C:" & StrConv(m(0).ImageName, vbUnicode))
    Open "C:\WINDOWS\system32\win32k.sys" For Binary As #1
    ReDim nWin32K(LOF(1) - 1)
    Get #1, , nWin32K
    Close #1
    pWin32K = VarPtr(nWin32K(0))
    pBaseKernel = m(0).Base
    Dim j As Long
    For j = 0 To i - 1
        If Replace(Mid(StrConv(m(j).ImageName, vbUnicode), m(j).ModuleNameOffset + 1), vbNullChar, "") = "win32k.sys" Then
            Win32KBase = m(j).Base
        End If
    Next
    GetSSDT
End Sub

Public Sub GetSSDT()
    Dim i As Long, j As Long, k As Long, KSDT As Long, SSDT As Long
    Dim A() As Byte
    Dim Export As IMAGE_EXPORT_DIRECTORY
    
    KSDT = GetProcAddress(pKernel, "KeServiceDescriptorTable") - pKernel + pBaseKernel
    'SSDT表的获取
    ReadKernelMemory KSDT + 8, VarPtr(i), 4, 0 '获取SSDT表项数量
    ReDim SSDTData(i - 1)
    ReadKernelMemory KSDT, VarPtr(SSDT), 4, 0 '获取SSDT表地址
    CopyMemory VarPtr(i), pNtDLL + &H3C, 4
    CopyMemory VarPtr(i), pNtDLL + i + &H78, 4
    CopyMemory VarPtr(Export), pNtDLL + i, Len(Export)
    For i = Export.nBase To Export.NumberOfFunctions - 1
        Dim strFunName As String
        CopyMemory VarPtr(k), pNtDLL + Export.AddressOfNames + i * 4, 4
        j = lstrlenA(pNtDLL + k)
        ReDim A(j - 1) As Byte
        CopyMemory VarPtr(A(0)), pNtDLL + k, j
        strFunName = StrConv(A, vbUnicode)
        If left(strFunName, 2) = "Nt" Then 'And strFunName <> "NtCurrentTeb" Then
            k = 0
            CopyMemory VarPtr(k), pNtDLL + Export.AddressOfNameOrdinals + i * 2, 2
            CopyMemory VarPtr(k), pNtDLL + Export.AddressOfFunctions + k * 4, 4
            CopyMemory VarPtr(k), pNtDLL + k + 1, 4
            If UBound(SSDTData) >= k Then
                'ReDim Preserve SSDTData(k)
                SSDTData(k).strName = strFunName
                ReadKernelMemory SSDT + k * 4, VarPtr(SSDTData(k).dwCurrAddress), 4, 0
            End If
        End If
    Next
    SSDT = SSDT - pBaseKernel + pKernel
    For i = 0 To UBound(SSDTData)
        CopyMemory VarPtr(SSDTData(i).dwRealAddress), SSDT + i * 4, 4
        SSDTData(i).dwRealAddress = SSDTData(i).dwRealAddress - &H400000 + pBaseKernel
    Next
    
    'ShadowSSDT表的获取
    '获取KeAddSystemServiceTable
    KSDT = GetProcAddress(pKernel, "KeAddSystemServiceTable") - pKernel + pBaseKernel
    '从汇编代码中获取KeSystemServiceTableShadow
    ReadKernelMemory KSDT + 28, VarPtr(KSDT), 4, 0
    
    ReadKernelMemory KSDT + 24, VarPtr(i), 4, 0 '获取Shadow SSDT表项数量
    ReDim ShadowSSDTData(i - 1)
    ReadKernelMemory KSDT + 16, VarPtr(SSDT), 4, 0 '获取Shadow SSDT表地址
    If i <> 667 Then
        MsgBox "警告！不支持此版本的ShadowSSDT的函数名获取！", vbCritical
    Else
        Open App.Path & "\ShadowSSDT函数表.txt" For Input As #1
    End If
    For j = 0 To UBound(ShadowSSDTData)
        ReadKernelMemory SSDT + j * 4, VarPtr(ShadowSSDTData(j).dwCurrAddress), 4, 0
        CopyMemory VarPtr(ShadowSSDTData(j).dwRealAddress), SSDT - Win32KBase + pWin32K + j * 4, 4
        If i = 667 Then
            Line Input #1, ShadowSSDTData(j).strName
        End If
    Next
    Close #1
End Sub

Public Sub RecoverSSDTSingle(ByVal nIndex As Long)
    Dim KSDT As Long
    KSDT = GetProcAddress(pKernel, "KeServiceDescriptorTable") - pKernel + pBaseKernel
    'SSDT表的获取
    ReadKernelMemory KSDT, VarPtr(KSDT), 4, 0 '获取SSDT表地址
    WriteKernelMemory KSDT + nIndex * 4, VarPtr(SSDTData(nIndex).dwRealAddress), 4, 0
End Sub

Public Sub RecoverSSDTAll()
    Dim i As Long
    For i = 0 To UBound(SSDTData)
        RecoverSSDTSingle i
    Next
End Sub

Public Sub RecoverShadowSSDTSingle(ByVal nIndex As Long)
    Dim KSDT As Long
    '获取KeAddSystemServiceTable
    KSDT = GetProcAddress(pKernel, "KeAddSystemServiceTable") - pKernel + pBaseKernel
    '从汇编代码中获取KeSystemServiceTableShadow
    ReadKernelMemory KSDT + 28, VarPtr(KSDT), 4, 0
    'SSDT表的获取
    ReadKernelMemory KSDT, VarPtr(KSDT), 4, 0 '获取SSDT表地址
    WriteKernelMemory KSDT + nIndex * 4, VarPtr(ShadowSSDTData(nIndex).dwRealAddress), 4, 0
End Sub

Public Sub RecoverShadowSSDTAll()
    Dim i As Long
    For i = 0 To UBound(ShadowSSDTData)
        RecoverShadowSSDTSingle i
    Next
End Sub

Public Sub ShutdownSSDT()
    FreeLibrary pKernel
End Sub

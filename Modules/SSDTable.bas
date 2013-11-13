Attribute VB_Name = "SSDTable"
Option Explicit
Public pKernel As Long   '�û�̬��ntkrnlpa.exe��ַ
Dim pNtDLL As Long       '�û�̬��ntdll.dll��ַ
Dim pBaseKernel As Long  '�ں�̬��ntkrnlpa.exe��ַ
Dim nWin32K() As Byte    'win32k.sys����
Dim pWin32K As Long      'ָ��nWin32K
Dim Win32KBase As Long   '�ں�̬��win32k.sys��ַ
Public Type SSDT_ENTRY
    strName As String
    dwCurrAddress As Long
    dwRealAddress As Long
End Type
Public SSDTData() As SSDT_ENTRY
Public ShadowSSDTData() As SSDT_ENTRY
'typedef struct _SYSTEM_SERVICE_TABLE
'{
'  PVOID   ServiceTableBase;      // SSDT (System Service Dispatch Table)�Ļ���ַ
'  PULONG  ServiceCounterTableBase;  // ����checked builds, ����SSDT��ÿ�����񱻵��õĴ���
'  ULONG   NumberOfService;      // �������ĸ���, NumberOfService*4 ����������ַ��Ĵ�С
'  ULONG   ParamTableBase;        // SSPT (System Service Parameter Table)�Ļ���ַ
'} SYSTEM_SERVICE_TABLE, *PSYSTEM_SERVICE_TABLE;
'
'typedef struct _SERVICE_DESCRIPTOR_TABLE
'{
'  SYSTEM_SERVICE_TABLE   ntoskrnl;  // ntoskrnl.exe�ķ�����
'  SYSTEM_SERVICE_TABLE   win32k;    // win32k.sys�ķ�����,(gdi.dll/user.dll���ں�֧��)
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
    'SSDT��Ļ�ȡ
    ReadKernelMemory KSDT + 8, VarPtr(i), 4, 0 '��ȡSSDT��������
    ReDim SSDTData(i - 1)
    ReadKernelMemory KSDT, VarPtr(SSDT), 4, 0 '��ȡSSDT���ַ
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
    
    'ShadowSSDT��Ļ�ȡ
    '��ȡKeAddSystemServiceTable
    KSDT = GetProcAddress(pKernel, "KeAddSystemServiceTable") - pKernel + pBaseKernel
    '�ӻ������л�ȡKeSystemServiceTableShadow
    ReadKernelMemory KSDT + 28, VarPtr(KSDT), 4, 0
    
    ReadKernelMemory KSDT + 24, VarPtr(i), 4, 0 '��ȡShadow SSDT��������
    ReDim ShadowSSDTData(i - 1)
    ReadKernelMemory KSDT + 16, VarPtr(SSDT), 4, 0 '��ȡShadow SSDT���ַ
    If i <> 667 Then
        MsgBox "���棡��֧�ִ˰汾��ShadowSSDT�ĺ�������ȡ��", vbCritical
    Else
        Open App.Path & "\ShadowSSDT������.txt" For Input As #1
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
    'SSDT��Ļ�ȡ
    ReadKernelMemory KSDT, VarPtr(KSDT), 4, 0 '��ȡSSDT���ַ
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
    '��ȡKeAddSystemServiceTable
    KSDT = GetProcAddress(pKernel, "KeAddSystemServiceTable") - pKernel + pBaseKernel
    '�ӻ������л�ȡKeSystemServiceTableShadow
    ReadKernelMemory KSDT + 28, VarPtr(KSDT), 4, 0
    'SSDT��Ļ�ȡ
    ReadKernelMemory KSDT, VarPtr(KSDT), 4, 0 '��ȡSSDT���ַ
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

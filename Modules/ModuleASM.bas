Attribute VB_Name = "ModuleASM"
Dim Code() As Byte
Private Declare Function IDECallProc Lib "user32" Alias "CallWindowProcA" (lpPrevWndFunc As Any, A As Any, b As Any, c As Any, d As Any) As Long
Public Enum ASMServiceID
    ASMReserved '���ڵ�һ�αȽ���dec eax������test eax,eax���Բ�ʹ��
    ASMDWordFromPtr '��ȡ����Ĳ���Bָ���DWord
    ASMReadFS '��ȡTEB+bָ���DWord
    ASMReadPEB '��ȡPEB+bָ���DWord
    ASMWriteDWord 'д��DWord
    ASMCallProc '����Bָ����������CΪ��������������Dָ���һ������
    ASMCPUID '����CPUID��Cָ��16�ֽڵ��ڴ�ֱ���EBX EDX ECX EAX��BΪ����ʱEAX��ֵ
    ASMLStrLen 'ASCII�ַ�����ȡ���ȣ�b�����ַ�������
    ASMWStrLen 'Unicode�ַ�����ȡ���ȣ�b�����ַ�������
    ASMMoveMemory '��CopyMemory�����memmove������B C D�ֱ�ΪĿ�ꡢԴ������
    ASMLeftShift '��������
    ASMRightShift '��������
    ASMGetCPUShortName '��ȡCPU�����ƣ�Bָ������ƣ�12���ֽڣ�������CPU֧�ֵ����ֵ
    ASMGetCPULongName '��ȡCPU�����ƣ�Bָ�����ƣ�48���ֽڣ����������Ƶ�ʵ�ʳ���
    ASMSetCallBack '���ûص�������BΪ1ʱ����RtlAllocateHeap�ĵ�ַ��BΪ2ʱ����RtlFreeHeap�ĵ�ַ
    ASMUnicodeStringToVbString '��UNICODE_STRINGת��ΪVBʹ�õ��ַ�����Bָ��UNICODE_STRING
    'Cָ��һ��VB�ַ��������C����Ϊ�ַ���������ֵҲ���ַ���
    ASMAnsiStringToUnicodeString '��Ansi�ַ���ת��ΪUnicode�ַ�����BΪAnsi�ַ���ָ��
    'CΪUnicode�ַ�����������DΪAnsi�ַ������ȣ�Unicode�������ĳ���ΪAnsi�ַ������ȵ�2��
    '����Unicode�ַ�����ʵ�ʳ��ȡ�
End Enum
Public Type CPUID_FPU
    EBX As Long
    EDX As Long
    ECX As Long
    EAX As Long
End Type
Public Type CacheInfo
    Level As Byte
    Size As Integer
    Way As Byte
    LineSize As Byte
End Type
Public mCacheInfo(255) As CacheInfo
Private Sub FillCacheInfo(ByVal n As Integer, ByVal Level As Integer, ByVal Size As Integer, ByVal Key As Integer, ByVal LineSize As Integer)
    With mCacheInfo(n)
        .Level = Level
        .Size = Size
        .Way = Way
        .LineSize = LineSize
    End With
End Sub
Public Sub InitASMModule()
    Code = LoadResData(106, "Bin")
    ASMCall ASMSetCallBack, 1, GetProcAddress(GetModuleHandle("ntdll"), "RtlAllocateHeap"), 0
    FillCacheInfo &H6, 1, 8, 4, 32
    FillCacheInfo &H8, 1, 16, 4, 32
    FillCacheInfo &HA, 1, 8, 2, 32
    FillCacheInfo &HC, 1, 16, 4, 32
    FillCacheInfo &H2C, 1, 32, 8, 64
    FillCacheInfo &H30, 1, 32, 8, 64
    FillCacheInfo &H60, 1, 16, 8, 64
    FillCacheInfo &H66, 1, 8, 4, 64
    FillCacheInfo &H67, 1, 16, 4, 64
    FillCacheInfo &H68, 1, 32, 4, 64
    FillCacheInfo &H39, 2, 128, 4, 64
    FillCacheInfo &H3B, 2, 128, 2, 64
    FillCacheInfo &H3C, 2, 256, 4, 64
    FillCacheInfo &H41, 2, 128, 4, 32
    FillCacheInfo &H42, 2, 256, 4, 32
    FillCacheInfo &H43, 2, 512, 4, 32
    FillCacheInfo &H44, 2, 1024, 4, 32
    FillCacheInfo &H45, 2, 2048, 4, 32
    FillCacheInfo &H79, 2, 128, 8, 64
    FillCacheInfo &H7A, 2, 256, 8, 64
    FillCacheInfo &H7B, 2, 512, 8, 64
    FillCacheInfo &H7C, 2, 1024, 8, 64
    FillCacheInfo &H82, 2, 256, 8, 32
    FillCacheInfo &H83, 2, 512, 8, 32
    FillCacheInfo &H84, 2, 1024, 8, 32
    FillCacheInfo &H85, 2, 2048, 8, 32
    FillCacheInfo &H86, 2, 512, 4, 64
    FillCacheInfo &H87, 2, 1024, 8, 64
    FillCacheInfo &H22, 3, 512, 4, 64
    FillCacheInfo &H23, 3, 1024, 8, 64
    FillCacheInfo &H25, 3, 2048, 8, 64
    FillCacheInfo &H29, 3, 4096, 8, 64
End Sub

Public Function ASMCall(ByVal A As ASMServiceID, ByVal b As Long, ByVal c As Long, ByVal d As Long) As Long
    ASMCall = IDECallProc(Code(0), ByVal A, ByVal b, ByVal c, ByVal d)
End Function

Public Function CPULongName() As String
    Dim n As Long, b(47) As Byte
    n = ASMCall(ASMGetCPULongName, VarPtr(b(0)), 0, 0)
    Dim c() As Byte
    ReDim c(n - 1)
    CopyMemory VarPtr(c(0)), VarPtr(b(0)), n
    CPULongName = StrConv(c, vbUnicode)
End Function

Public Function CPUShortName() As String
    Dim b(11) As Byte
    ASMCall ASMGetCPUShortName, VarPtr(b(0)), 0, 0
    CPUShortName = StrConv(b, vbUnicode)
End Function

Public Function CPUSerialNumber() As String
    Dim cid As CPUID_FPU
    ASMCall ASMCPUID, 1, VarPtr(cid), 0
    Dim c2 As CPUID_FPU
    ASMCall ASMCPUID, 2, VarPtr(c2), 0
    If (cid.EDX And &H40000) = 0 Then
        CPUSerialNumber = ""
        Exit Function
    End If
    Dim ints(5) As Integer
    CopyMemory VarPtr(ints(0)), VarPtr(cid.EAX), 4
    CopyMemory VarPtr(ints(2)), VarPtr(cid.ECX), 4
    CopyMemory VarPtr(ints(4)), VarPtr(cid.EDX), 4
    Dim S(5) As String
    S(0) = right("0000" & Hex(ints(0)), 4)
    S(1) = right("0000" & Hex(ints(1)), 4)
    S(2) = right("0000" & Hex(ints(2)), 4)
    S(3) = right("000" & Hex(ints(3)), 3)
    S(4) = right("0000" & Hex(ints(4)), 4)
    S(5) = right("0000" & Hex(ints(5)), 4)
    CPUSerialNumber = Join(S, "-")
End Function

Public Function UnicodeStringToStringASM(u As UNICODE_STRING) As String
    Dim p As String
    ASMCall ASMUnicodeStringToVbString, VarPtr(u), VarPtr(p), 0
    UnicodeStringToStringASM = p
End Function

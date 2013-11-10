Attribute VB_Name = "BeaEngine"
Public Declare Function BeaEngineRevision Lib "BeaEngine.dll" Alias "_BeaEngineRevision@0" () As Long
Public Declare Function BeaEngineVersion Lib "BeaEngine.dll" Alias "_BeaEngineVersion@0" () As Long
Public Declare Function Disasm Lib "BeaEngine.dll" Alias "_Disasm@4" (lpAsm As Disasm) As Long
 
Public Type REX_Struct
    W_      As Byte
    R_      As Byte
    X_      As Byte
    B_      As Byte
    State   As Byte
End Type
Public Type PREFIXINFO
    Number          As Long
    NbUndefined     As Long
    LockPrefix      As Byte
    OperandSize     As Byte
    AddressSize     As Byte
    RepnePrefix     As Byte
    RepPrefix       As Byte
    FSPrefix        As Byte
    SSPrefix        As Byte
    GSPrefix        As Byte
    ESPrefix        As Byte
    CSPrefix        As Byte
    DSPrefix        As Byte
    BranchTaken     As Byte
    BranchNotTaken  As Byte
    REX             As REX_Struct
End Type

Public Type EFL_Struct
    OF_     As Byte
    SF_     As Byte
    ZF_     As Byte
    AF_     As Byte
    PF_     As Byte
    CF_     As Byte
    TF_     As Byte
    IF_     As Byte
    DF_     As Byte
    NT_     As Byte
    RF_     As Byte
    Reserve As Byte          ' alignment
End Type

Public Type MemoryType
    BaseRegister    As Long   '左边（如0=EAX）
    IndexRegister   As Long   '右边（如1=ECX）
    Scale           As Long
    Displacement    As INT64
End Type

Public Type INSTRTYPE
    Category                As Long
    Opcode                  As Long
    Mnemonic(1 To 16)       As Byte           '命令名称(如mov)
    BranchType              As Long
    Flags                   As EFL_Struct
    AddrValue               As INT64
    Immediat                As INT64
    ImplicitModifiedRegs    As Long
End Type

Public Type ArgType
    ArgMnemonic(1 To 32) As Byte         '参数本身(如eax*4+ecx+0x00000004)
    ArgType              As Long
    ArgSize              As Long
    ArgPosition          As Long
    AccessMode           As Long
    Memory               As MemoryType
    SegmentReg           As Long
End Type

Public Type Disasm
    EIP             As Long
    VirtualAddr     As INT64             '假设EIP指向的内存的地址
    SecurityBlock   As Long              '最多还能读取的字节数
    CompleteInStr(1 To 64) As Byte
    Archi           As Long
    Options         As INT64
    Instruction     As INSTRTYPE
    Argument1       As ArgType
    Argument2       As ArgType
    Argument3       As ArgType
    Prefix          As PREFIXINFO
    Reserved_(1 To 40) As Long
End Type


Public Const LowPosition = 0
Public Const HighPosition = 1

Public Const ESReg = 1
Public Const DSReg = 2
Public Const FSReg = 3
Public Const GSReg = 4
Public Const CSReg = 5
Public Const SSReg = 6


 ' ********** Prefixes

Public Const InvalidPrefix = 4
Public Const InUsePrefix = 1
Public Const SuperfluousPrefix = 2
Public Const NotUsedPrefix = 0
Public Const MandatoryPrefix = 8

 ' ********** EFLAGS states
 
Public Const TE_ = 1        ' test
Public Const MO_ = 2        ' modify
Public Const RE_ = 4        ' reset
Public Const SE_ = 8        ' set
Public Const UN_ = &H10         ' undefined
Public Const PR_ = &H20         ' restore prior value

' __________________________________________________________________________________________________________
'
'                                       INSTRUCTION_TYPE
' __________________________________________________________________________________________________________

Public Const GENERAL_PURPOSE_INSTRUCTION = &H10000
Public Const FPU_INSTRUCTION = &H20000
Public Const MMX_INSTRUCTION = &H40000
Public Const SSE_INSTRUCTION = &H80000
Public Const SSE2_INSTRUCTION = &H100000
Public Const SSE3_INSTRUCTION = &H200000
Public Const SSSE3_INSTRUCTION = &H400000
Public Const SSE41_INSTRUCTION = &H800000
Public Const SSE42_INSTRUCTION = &H1000000
Public Const SYSTEM_INSTRUCTION = &H2000000
Public Const VM_INSTRUCTION = &H4000000
Public Const UNDOCUMENTED_INSTRUCTION = &H8000000
Public Const AMD_INSTRUCTION = &H10000000
Public Const ILLEGAL_INSTRUCTION = &H20000000
Public Const AES_INSTRUCTION = &H40000000
Public Const CLMUL_INSTRUCTION = &H80000000
    

Public Const DATA_TRANSFER = 1
Public Const ARITHMETIC_INSTRUCTION = 2
Public Const LOGICAL_INSTRUCTION = 3
Public Const SHIFT_ROTATE = 4
Public Const BIT_BYTE = 5
Public Const CONTROL_TRANSFER = 6
Public Const STRING_INSTRUCTION = 7
Public Const InOutINSTRUCTION = 8
Public Const ENTER_LEAVE_INSTRUCTION = 9
Public Const FLAG_CONTROL_INSTRUCTION = 10
Public Const SEGMENT_REGISTER = 11
Public Const MISCELLANEOUS_INSTRUCTION = 12

Public Const COMPARISON_INSTRUCTION = 13
Public Const LOGARITHMIC_INSTRUCTION = 14
Public Const TRIGONOMETRIC_INSTRUCTION = 15
Public Const UNSUPPORTED_INSTRUCTION = 16
    
Public Const LOAD_CONSTANTS = 17
Public Const FPUCONTROL = 18
Public Const STATE_MANAGEMENT = 19

Public Const CONVERSION_INSTRUCTION = 20

Public Const SHUFFLE_UNPACK = 21
Public Const PACKED_SINGLE_PRECISION = 22
Public Const SIMD128bits = 23
Public Const SIMD64bits = 24
Public Const CACHEABILITY_CONTROL = 25
    
Public Const FP_INTEGER_CONVERSION = 26
Public Const SPECIALIZED_128bits = 27
Public Const SIMD_FP_PACKED = 28
Public Const SIMD_FP_HORIZONTAL = 29
Public Const AGENT_SYNCHRONISATION = 30

Public Const PACKED_ALIGN_RIGHT = 31
Public Const PACKED_SIGN = 32

    ' ****************************************** SSE4
    
Public Const PACKED_BLENDING_INSTRUCTION = 33
Public Const PACKED_TEST = 34
    
    ' CONVERSION_INSTRUCTION -> Packed Integer Format Conversions et Dword Packing With Unsigned Saturation
    ' COMPARISON -> Packed Comparison SIMD Integer Instruction
    ' ARITHMETIC_INSTRUCTION -> Dword Multiply Instruction
    ' DATA_TRANSFER -> POPCNT

Public Const PACKED_MINMAX = 35
Public Const HORIZONTAL_SEARCH = 36
Public Const PACKED_EQUALITY = 37
Public Const STREAMING_LOAD = 38
Public Const INSERTION_EXTRACTION = 39
Public Const DOT_PRODUCT = 40
Public Const SAD_INSTRUCTION = 41
Public Const ACCELERATOR_INSTRUCTION = 42
Public Const ROUND_INSTRUCTION = 43

' __________________________________________________________________________________________________________
'
'                                       BranchTYPE
' __________________________________________________________________________________________________________
Public Const Jo_ = 1
Public Const Jno_ = -1
Public Const Jc_ = 2
Public Const Jnc_ = -2
Public Const Je_ = 3
Public Const Jne_ = -3
Public Const Ja_ = 4
Public Const Jna_ = -4
Public Const Js_ = 5
Public Const Jns_ = -5
Public Const Jp_ = 6
Public Const Jnp_ = -6
Public Const Jl_ = 7
Public Const Jnl_ = -7
Public Const Jg_ = 8
Public Const Jng_ = -8
Public Const Jb_ = 9
Public Const Jnb_ = -9
Public Const Jecxz_ = 10
Public Const JmpType = 11
Public Const CallType = 12
Public Const RetType = 13

' __________________________________________________________________________________________________________
'
'                                       ARGUMENTS_TYPE
' __________________________________________________________________________________________________________

Public Const NO_ARGUMENT = &H10000000
Public Const REGISTER_TYPE = &H20000000
Public Const MEMORY_TYPE = &H40000000
Public Const CONSTANT_TYPE = &H80000000

Public Const MMX_REG = &H10000
Public Const GENERAL_REG = &H20000
Public Const FPU_REG = &H40000
Public Const SSE_REG = &H80000
Public Const CR_REG = &H100000
Public Const DR_REG = &H200000
Public Const SPECIAL_REG = &H400000
Public Const MEMORY_MANAGEMENT_REG = &H800000                   ' GDTR (REG0), LDTR (REG1), IDTR (REG2), TR (REG3)
Public Const SEGMENT_REG = &H1000000                            ' ES (REG0), CS (REG1), SS (REG2), DS (REG3), FS (REG4), GS (REG5)

Public Const RELATIVE_ = &H4000000
Public Const ABSOLUTE_ = &H8000000

Public Const READ_ = 1
Public Const WRITE_ = 2
    ' ************ Regs

Public Const REG0 = 1                              ' 30h
Public Const REG1 = 2                              ' 31h
Public Const REG2 = 4                              ' 32h
Public Const REG3 = 8                              ' 33h
Public Const REG4 = &H10                            ' 34h
Public Const REG5 = &H20                            ' 35h
Public Const REG6 = &H40                           ' 36h
Public Const REG7 = &H80                            ' 37h
Public Const REG8 = &H100                           ' 38h
Public Const REG9 = &H200                           ' 39h
Public Const REG10 = &H400                              ' 3Ah
Public Const REG11 = &H800                              ' 3Bh
Public Const REG12 = &H1000                             ' 3Ch
Public Const REG13 = &H2000                             ' 3Dh
Public Const REG14 = &H4000                             ' 3Eh
Public Const REG15 = &H8000                             ' 3Fh

    ' ************ SPECIAL_INFO

Public Const UNKNOWN_OPCODE = -1
Public Const OUT_OF_BLOCK = 0
Public Const NoTabulation = 0
Public Const Tabulation = 1
Public Const MasmSyntax = 0
Public Const GoAsmSyntax = &H100
Public Const NasmSyntax = &H200
Public Const ATSyntax = &H400
Public Const PrefixedNumeral = &H10000
Public Const SuffixedNumeral = 0
Public Const ShowSegmentRegs = &H1000000

Public Sub Test()
    Dim A(5999) As Byte
    Open "C:\Disasm.dat" For Binary As #1
    Open "C:\Disasm.out.txt" For Output As #2
    Dim n As Disasm
    Get #1, , A
    Close #1
    n.EIP = VarPtr(A(0))
    n.SecurityBlock = 6000
    n.VirtualAddr.dwLow = 0
    Dim i As Long, j As Long, S As String
    Do
        j = Disasm(n)
        If j <= 0 Then Stop
        i = i + j
        If i = 6000 Then Exit Do
        n.EIP = VarPtr(A(i))
        n.SecurityBlock = 6000 - i
        n.VirtualAddr.dwLow = i
        S = StrConv(n.CompleteInStr, vbUnicode)
        S = left(S, InStr(S, vbNullChar) - 1)
        S = Trim(S): If right(S, 1) = "h" Then S = left(S, Len(S) - 1)
        If S = "int3" Then Exit Do
        Print #2, S
    Loop
    Close #2
End Sub

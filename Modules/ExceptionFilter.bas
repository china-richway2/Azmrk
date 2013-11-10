Attribute VB_Name = "ExceptionFilter"
Public Declare Function ZwContinue Lib "ntdll" (pContext As Any, ByVal bAlertable As Long) As Long
Public Type AzmrkHook
    Back(4) As Byte
    NewAddr As Long
    OldAddr As Long
    Status As Boolean
End Type
Public Type EXCEPTION_RECORD
    ExceptionCode As Long
    ExceptionFlags As Long
    Record As Long
    ExceptionAddress As Long
    NumberParameters As Long
    ExceptionInformation As Long
End Type

Public ExceptionFilter As AzmrkHook
Sub HookApi(Hook As AzmrkHook)
    Dim Old As Long
    VirtualProtectEx GetCurrentProcess, ByVal Hook.OldAddr, 5, &H40, Old
    If Hook.Status Then
        CopyMemory VarPtr(Hook.Back(0)), Hook.OldAddr, 5
        CopyMemory Hook.OldAddr, VarPtr(CByte(&HE9)), 1
        CopyMemory Hook.OldAddr + 1, VarPtr(CLng(Hook.NewAddr - Hook.OldAddr - 5)), 4
    Else
        CopyMemory Hook.OldAddr, VarPtr(Hook.Back(0)), 5
    End If
    VirtualProtectEx GetCurrentProcess, ByVal Hook.OldAddr, 5, Old, Old
End Sub
Function NewFilter(ExceptionRecord As EXCEPTION_RECORD, pContext As CONTEXT) As Long
    MsgBox "Right"
    NewFilter = 1
    ZwContinue pContext, 0
End Function
Sub HookExceptionFilter()
    Dim Ptr As Long
    Ptr = GetModuleHandle("ntdll")
    Ptr = GetProcAddress(Ptr, "KiUserExceptionDispatcher") + 10
    Dim P2 As Long
    CopyMemory VarPtr(P2), Ptr, 4
    Ptr = AddUnsigned(Ptr + 4, P2)
    ExceptionFilter.OldAddr = Ptr
    ExceptionFilter.NewAddr = ReturnPtr(AddressOf NewFilter)
    ExceptionFilter.Status = True
    HookApi ExceptionFilter
End Sub

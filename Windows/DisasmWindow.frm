VERSION 5.00
Begin VB.Form DisasmWindow 
   Caption         =   "反汇编窗口"
   ClientHeight    =   3090
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3090
   ScaleWidth      =   4680
   StartUpPosition =   3  '窗口缺省
   Begin VB.PictureBox MyDisasmList 
      Height          =   735
      Left            =   840
      ScaleHeight     =   675
      ScaleWidth      =   1035
      TabIndex        =   1
      Top             =   120
      Width           =   1095
   End
   Begin VB.Timer Timer1 
      Interval        =   10000
      Left            =   1800
      Top             =   1320
   End
   Begin VB.VScrollBar vs 
      Height          =   1575
      Left            =   3960
      TabIndex        =   0
      Top             =   480
      Width           =   255
   End
   Begin VB.Menu dMenu 
      Caption         =   "菜单"
      Begin VB.Menu dJumpToExp 
         Caption         =   "转到表达式"
      End
   End
End
Attribute VB_Name = "DisasmWindow"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Process As Long, ParentWindow As DebugWindow
Dim vsMin As Long, vsSize As Long
Dim arrItems() As MYLISTITEM
Dim numItem As Long
Dim leftAddr As Single, leftHex As Single, leftDisasm As Single
Dim Selected As Integer
Dim Regs(9) As Long, HaveContext As Boolean
Dim lIsScroll As Long, Changing As Boolean
Private Type MYLISTITEM
    Addr As String * 8
    ASMCode As String
    Tip As String
    HexData As String
End Type
Public Sub SetProcess(ByVal hProcess As Long, Window As DebugWindow)
    leftAddr = TextWidth("00000000") + 45
    leftHex = leftAddr + TextWidth("0000000000000")
    leftDisasm = leftHex + TextWidth("                                ")
    Process = hProcess
    Set ParentWindow = Window
End Sub

Public Sub SetContext(ByVal pContext As Long)
    Dim lpContext As CONTEXT
    CopyMemory VarPtr(lpContext), pContext, Len(lpContext)
    With lpContext
        Regs(0) = .EAX
        Regs(1) = .ECX
        Regs(2) = .EDX
        Regs(3) = .EBX
        Regs(4) = .Esp
        Regs(5) = .Ebp
        Regs(6) = .Esi
        Regs(7) = .Edi
        Regs(8) = .EFlags
        Regs(9) = .EIP
        HaveContext = True
    End With
End Sub

Private Function ArgTypeToNum(ByVal ArgType As Long) As Long
    If ArgType And REG0 Then
        ArgTypeToNum = 0
    ElseIf ArgType And REG0 Then
        ArgTypeToNum = 0
    ElseIf ArgType And REG1 Then
        ArgTypeToNum = 1
    ElseIf ArgType And REG2 Then
        ArgTypeToNum = 2
    ElseIf ArgType And REG3 Then
        ArgTypeToNum = 3
    ElseIf ArgType And REG4 Then
        ArgTypeToNum = 4
    ElseIf ArgType And REG5 Then
        ArgTypeToNum = 5
    ElseIf ArgType And REG6 Then
        ArgTypeToNum = 6
    ElseIf ArgType And REG7 Then
        ArgTypeToNum = 7
    ElseIf ArgType And REG8 Then
        ArgTypeToNum = 8
    ElseIf ArgType And REG9 Then
        ArgTypeToNum = 9
    ElseIf ArgType And REG10 Then
        ArgTypeToNum = 10
    ElseIf ArgType And REG11 Then
        ArgTypeToNum = 11
    ElseIf ArgType And REG12 Then
        ArgTypeToNum = 12
    ElseIf ArgType And REG13 Then
        ArgTypeToNum = 13
    ElseIf ArgType And REG14 Then
        ArgTypeToNum = 14
    ElseIf ArgType And REG15 Then
        ArgTypeToNum = 15
    End If
End Function

Private Function ArgumentToTip(lpArg As ArgType) As String
    Dim dwValue As Long
    If Not HaveContext Then Exit Function
    With lpArg
        'NO_ARGUMENT
        'REGISTER_TYPE
        'MEMORY_TYPE
        'CONSTANT_TYPE
        If .ArgType And NO_ARGUMENT Then
            Exit Function
        ElseIf .ArgType And REGISTER_TYPE Then
            Dim RegNum As Long
            If .ArgType And MMX_REG Then
            ElseIf .ArgType And GENERAL_REG Then: dwValue = Regs(ArgTypeToNum(.ArgType))
            ElseIf .ArgType And FPU_REG Then
            ElseIf .ArgType And SSE_REG Then
            ElseIf .ArgType And CR_REG Then
            ElseIf .ArgType And DR_REG Then: ArgumentToTip = "特权指令"
            ElseIf .ArgType And SPECIAL_REG Then
            ElseIf .ArgType And MEMORY_MANAGEMENT_REG Then
            ElseIf .ArgType And SEGMENT_REG Then
                If .AccessMode = WRITE_ Then
                    ArgumentToTip = "修改段寄存器"
                End If
            End If
        ElseIf .ArgType And CONSTANT_TYPE Then
            dwValue = .ArgPosition
        ElseIf .ArgType And MEMORY_TYPE Then
            With .Memory
                If .BaseRegister <> 0 Then
                    If .IndexRegister <> 0 Then
                        dwValue = Regs(ArgTypeToNum(.BaseRegister) + ArgTypeToNum(.IndexRegister) * .Scale)
                    Else
                        dwValue = Regs(ArgTypeToNum(.BaseRegister))
                    End If
                Else
                    If .IndexRegister <> 0 Then
                        dwValue = Regs(ArgTypeToNum(.IndexRegister) * .Scale)
                    End If
                End If
                dwValue = dwValue + .Displacement.dwLow
            End With
            Dim Buffer() As Byte
            ReDim Buffer(.ArgSize)
            If .AccessMode = READ_ Then
                ZwReadVirtualMemory Process, ByVal dwValue, Buffer(0), .ArgSize, RegNum
                If RegNum <> .ArgSize Then
                    ArgumentToTip = "目标内存无法访问或权限不足"
                    Exit Function
                End If
            ElseIf .AccessMode And WRITE_ Then
                ZwReadVirtualMemory Process, ByVal dwValue, Buffer(0), .ArgSize, RegNum
                If RegNum <> .ArgSize Then
                    ArgumentToTip = "目标内存无法访问或权限不足"
                    Exit Function
                End If
                ZwWriteVirtualMemory Process, ByVal dwValue, Buffer(0), .ArgSize, RegNum
                '把原来的数据再次写到目标内存，判断是否能写
                If RegNum <> .ArgSize Then
                    ArgumentToTip = "目标内存无法写入"
                    Exit Function
                End If
            End If
            If .ArgSize = 4 Then 'DWORD
                ArgumentToTip = "[" & Hex2(dwValue, 8) & "]="
                ZwReadVirtualMemory Process, ByVal dwValue, dwValue, 4, 0
                ArgumentToTip = ArgumentToTip & Hex2(dwValue, 8)
                Dim strExpr As String
                strExpr = ParentWindow.PtrToExpr(dwValue)
                If strExpr <> "" Then ArgumentToTip = ArgumentToTip & "=" & strExpr
            End If
        End If
    End With
End Function

Public Sub JumpTo(ByVal nAddr As Long)
    'Stop
    Dim Memory As MEMORY_BASIC_INFORMATION
    ReDim arrItems(0): numItem = 0
    If VirtualQueryEx(Process, nAddr, Memory, 28) = 28 Then
        vsMin = Memory.BaseAddress
        vsSize = Memory.RegionSize
        Changing = True
        vs.Value = (nAddr - vsMin) * 32768 / vsSize
        Changing = False
        If Memory.State = MEM_FREE Then
            Memory.AllocationBase = Memory.BaseAddress
        End If
        Dim i As Long, j As Long, first As Boolean
        Dim bufHex As String * 100
        i = nAddr - Memory.BaseAddress
        Dim MyDisasm As Disasm
        With MyDisasm
            Dim Buffer() As Byte
            ReDim Buffer(Memory.RegionSize - 1)
            .EIP = VarPtr(Buffer(0))
            Dim nSize As Long
            Call ZwReadVirtualMemory(Process, ByVal Memory.BaseAddress, Buffer(0), Memory.RegionSize, nSize)
            If nSize = 0 Then Exit Sub
            .VirtualAddr.dwLow = Memory.BaseAddress
        End With
        Do
            With MyDisasm
                .SecurityBlock = Memory.RegionSize - i
                If .SecurityBlock <= 0 Then Exit Do
                .VirtualAddr.dwLow = Memory.BaseAddress + i
                .EIP = VarPtr(Buffer(i))
            End With
            ReDim Preserve arrItems(numItem)
            With arrItems(numItem)
                Dim nLength As Long
                nLength = Disasm(MyDisasm)
                .Addr = Hex2(MyDisasm.VirtualAddr.dwLow, 8)
                bufHex = vbNullString
                .HexData = bufHex
                If nLength = UNKNOWN_OPCODE Then
                    .HexData = Hex2(Buffer(i), 2)
                    i = i + 1
                    .ASMCode = "???"
                    .Tip = "未知指令"
                ElseIf nLength = OUT_OF_BLOCK Or nLength > MyDisasm.SecurityBlock Then
                    .HexData = Hex2(Buffer(i), 2)
                    i = i + 1
                    .ASMCode = "???"
                    .Tip = "指令在内存块尾"
                    Exit Do
                Else
                    For j = i To i + nLength - 1
                        Mid(bufHex, (j - i) * 2 + 1, 2) = Hex2(Buffer(j), 2)
                    Next
                    i = j
                    .HexData = bufHex
                    .ASMCode = StrConv(MyDisasm.CompleteInStr, vbUnicode)
                    .ASMCode = left(.ASMCode, InStr(.ASMCode, vbNullChar) - 1)
                    'MyDisasm.Argument1
                End If
                .Tip = ParentWindow.PtrToExpr(MyDisasm.VirtualAddr.dwLow) & vbCrLf
                'Stop
                .Tip = .Tip & ArgumentToTip(MyDisasm.Argument1) & vbCrLf
                .Tip = .Tip & ArgumentToTip(MyDisasm.Argument2) & vbCrLf
                .Tip = .Tip & ArgumentToTip(MyDisasm.Argument3) & vbCrLf
                j = MyDisasm.Instruction.AddrValue.dwLow
                If ParentWindow.PtrToExpr(j) <> "" Then .Tip = .Tip & Hex(j) & "=" & ParentWindow.PtrToExpr(j) & vbCrLf
                Do While InStr(.Tip, vbCrLf & vbCrLf) > 0
                    .Tip = Replace(.Tip, vbCrLf & vbCrLf, vbCrLf)
                Loop
                If right(.Tip, 2) = vbCrLf Then .Tip = left(.Tip, Len(.Tip) - 2)
            End With
            numItem = numItem + 1
            If i >= nAddr - Memory.BaseAddress + 256 Then Exit Do
        Loop
    End If
    Call MyDisasmList_Paint
End Sub

Private Sub dJumpToExp_Click()
    JumpTo ParentWindow.ExprToPtr(InputBox("请输入表达式"))
End Sub

Private Sub Form_Load()
    ApplyLang Me
End Sub

Private Sub Form_Resize()
    vs.Move ScaleWidth - vs.Width, 0
    vs.Height = ScaleHeight
    MyDisasmList.Move 0, 0, ScaleWidth - vs.Width, ScaleHeight
End Sub

Private Sub MyDisasmList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim hCursor As Long
    hCursor = LoadCursor(0, IDC_SIZEWE)
    If (x \ 15 > leftAddr \ 15 - 3) And (x \ 15 < leftAddr \ 15 + 3) Then
        SetCursor hCursor
        If Button = 1 Then
            leftAddr = x
            Call MyDisasmList_Paint
        End If
    ElseIf (x \ 15 > leftHex \ 15 - 3) And (x \ 15 < leftHex \ 15 + 3) Then
        SetCursor hCursor
        If Button = 1 Then
            If x > leftAddr + 45 Then leftHex = x
            Call MyDisasmList_Paint
        End If
    ElseIf (x \ 15 > leftDisasm \ 15 - 3) And (x \ 15 < leftDisasm \ 15 + 3) Then
        SetCursor hCursor
        If Button = 1 Then
            If x > leftHex + 45 Then leftDisasm = x
            Call MyDisasmList_Paint
        End If
    End If
End Sub

Private Sub MyDisasmList_Paint()
    With MyDisasmList
        Dim i As Long
        .Cls
        '写列首
        MyDisasmList.Print FindString("Disasm.Address");
        MyDisasmList.Line (leftAddr - 45, 0)-(leftAddr - 45, .ScaleHeight)
        .CurrentX = leftAddr: .CurrentY = 0
        MyDisasmList.Print FindString("Disasm.HexData");
        MyDisasmList.Line (leftHex - 45, 0)-(leftHex - 45, .ScaleHeight)
        .CurrentX = leftHex: .CurrentY = 0
        MyDisasmList.Print FindString("Disasm");
        MyDisasmList.Line (leftDisasm - 45, 0)-(leftDisasm - 45, .ScaleHeight)
        .CurrentX = leftDisasm: .CurrentY = 0
        MyDisasmList.Print FindString("Disasm.Notes")
        '然后所有ListItem画上去
        For i = 0 To numItem - 1
            If .CurrentY > .ScaleHeight Then Exit Sub
            MyDisasmList.Print arrItems(i).Addr;
            .CurrentX = leftAddr
            MyDisasmList.Print arrItems(i).HexData;
            .CurrentX = leftHex
            MyDisasmList.Print arrItems(i).ASMCode;
            .CurrentX = leftDisasm
            MyDisasmList.Print arrItems(i).Tip
        Next
    End With
End Sub

Private Sub Timer1_Timer()
    'FindAll
End Sub

Private Sub vs_Change()
    On Error GoTo E
    If Changing Then Exit Sub
    If Not lIsScroll Then
        If vs.Value > lIsScroll Then
            JumpTo Val("&H" & arrItems(2).Addr)
        Else
            JumpTo Val("&H" & arrItems(0).Addr) - 1
        End If
        lIsScroll = 0
        Exit Sub
    End If
    JumpTo vs.Value / 32768 * vsSize + vsMin
E:
    lIsScroll = 0
End Sub

Private Sub vs_Scroll()
    If Changing Then Exit Sub
    lIsScroll = vs.Value
    JumpTo vs.Value / 32768 * vsSize + vsMin
End Sub

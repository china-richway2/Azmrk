Attribute VB_Name = "GlobalModule"
Public Type ModuleInformation
    Reserved(7) As Byte
    Base As Long
    Size As Long
    Flags As Long
    Index As Integer
    Unknown As Integer
    Loadcount As Integer
    ModuleNameOffset As Integer
    ImageName(255) As Byte
End Type

Private Function StringFromPtr(ByVal pString As Long) As String
    Dim Buff() As Byte, Length As Long
    Length = lstrlenA(pString)
    If Length = 0 Then Exit Function
    ReDim Buff(Length - 1)
    CopyMemory VarPtr(Buff(0)), pString, Length
    StringFromPtr = StrConv(Buff, vbUnicode)
End Function

Public Sub GMNew()
    Dim nLength As Long, nInf() As ModuleInformation, Buffer() As Byte, st As Long
    Menu.LVModules.ListItems.Clear
    st = ZwQuerySystemInformation(SystemModuleInformation, 0, 0, nLength)
    If st < 0 Then
        If st <> STATUS_INFO_LENGTH_MISMATCH Then
            Exit Sub
        End If
    End If
    ReDim Buffer(1 To nLength)
    st = ZwQuerySystemInformation(SystemModuleInformation, VarPtr(Buffer(1)), nLength, nLength)
    If st < 0 Then Exit Sub
    CopyMemory VarPtr(nLength), VarPtr(Buffer(1)), 4
    ReDim nInf(1 To nLength)
    CopyMemory VarPtr(nInf(1)), VarPtr(Buffer(5)), nLength * Len(nInf(1))
    For st = 1 To nLength
        With Menu.LVModules.ListItems.Add(, , StringFromPtr(VarPtr(nInf(st).ImageName(nInf(st).ModuleNameOffset))))
            .SubItems(1) = StringFromPtr(VarPtr(nInf(st).ImageName(0)))
            .SubItems(2) = FormatHex(nInf(st).Base)
            .SubItems(3) = FormatHex(nInf(st).Size)
            .SubItems(4) = nInf(st).Loadcount
        End With
    Next
End Sub

Public Function AddrToModuleName(ByVal lpAddress As Long) As String
    Dim nLength As Long, nInf() As ModuleInformation, Buffer() As Byte, st As Long
    Menu.LVModules.ListItems.Clear
    st = ZwQuerySystemInformation(SystemModuleInformation, 0, 0, nLength)
    If st < 0 Then
        If st <> STATUS_INFO_LENGTH_MISMATCH Then
            Exit Function
        End If
    End If
    ReDim Buffer(1 To nLength)
    st = ZwQuerySystemInformation(SystemModuleInformation, VarPtr(Buffer(1)), nLength, nLength)
    If st < 0 Then Exit Function
    CopyMemory VarPtr(nLength), VarPtr(Buffer(1)), 4
    ReDim nInf(1 To nLength)
    CopyMemory VarPtr(nInf(1)), VarPtr(Buffer(5)), nLength * Len(nInf(1))
    For st = 1 To nLength
        If lpAddress >= nInf(st).Base And lpAddress < nInf(st).Base + nInf(st).Size Then
            AddrToModuleName = StringFromPtr(VarPtr(nInf(st).ImageName(0)))
            Exit Function
        End If
    Next
    AddrToModuleName = "Î´ÖªÄ£¿é"
End Function

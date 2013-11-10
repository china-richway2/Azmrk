Attribute VB_Name = "Handle"
Public Function CloseRemoteHandle(ByVal hHandle As Long, ByVal mPid As Long, Optional ByVal NotQuiet As Boolean) As Long
    Dim hThread As Long
    Dim zModule As Long, zProc As Long
    Dim hProcess As Long, dwTid As Long
    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, mPid)
    DoEvents
    If hProcess = 0 Then
        If NotQuiet Then MsgBox "目标进程无法打开！", vbCritical
        Exit Function
    End If
    If NT_SUCCESS(ZwDuplicateObject(hProcess, hHandle, GetCurrentProcess, hThread, 1, 0, DUPLICATE_CLOSE_SOURCE)) Then
        ZwClose hThread
        CloseRemoteHandle = 1
    End If
    zModule = GetModuleHandle("ntdll")
    'DoEvents
    zProc = GetProcAddress(zModule, "ZwClose")
    'DoEvents
    Dim ta As SECURITY_ATTRIBUTES
    'DoEvents
    hThread = CreateRemoteThread(hProcess, ta, 0, zProc, hHandle, 0, dwTid)
    'DoEvents
    If WaitForSingleObject(hThread, 3) <> 0 Then
        If NotQuiet Then MsgBox "未知错误！", vbCritical
    Else
        zProc = -1
        GetExitCodeThread hThread, zProc
        'DoEvents
        If zProc = -1 Then
            If NotQuiet Then MsgBox "无法获取返回值！", vbCritical
        ElseIf Not NT_SUCCESS(zProc) Then
            If NotQuiet Then MsgBox "NT 错误：" & FormatHex(zProc), vbCritical
        Else
            If NotQuiet Then MsgBox "成功", vbInformation
            CloseRemoteHandle = 1
        End If
        'DoEvents
    End If
    ZwClose hThread
    ZwClose hProcess
    'DoEvents
End Function

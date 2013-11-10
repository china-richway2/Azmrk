Attribute VB_Name = "Window"
Option Explicit
Public Declare Function IsWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hWnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hWnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function RealGetWindowClass Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hWnd As Long, ByVal Msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hWnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function InvalidateRectBynum Lib "user32" Alias "InvalidateRect" (ByVal hWnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hWnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hWnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hWnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hWnd As Long, ByVal fEnable As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function RegisterClass Lib "user32" Alias "RegisterClassA" (Class As WNDCLASS) As Long
Public Declare Function UnregisterClass Lib "user32" Alias "UnregisterClassA" (ByVal lpClassName As String, ByVal hInstance As Long) As Long
Public Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long


Public Const SWP_NOMOVE = &H2
Public Const SWP_NOSIZE = &H1

Public Const HWND_TOPMOST = -1
Public Const HWND_NOTOPMOST = -2

Public Const GW_CHILD = 5
Public Const GW_HWNDNEXT = 2

Public Const WM_CLOSE = 16
Public Const WM_QUIT = &H12
Public Const WM_STOP = 18
Public Const WM_DESTROY = &H2
Public Const WM_NCDESTROY = &H82
Public Const WM_SETICON = &H80
Public Const WM_GETTEXT = &HD
Public Const WM_SETTEXT = &HC

Public Const SW_SHOWNOACTIVATE = 4
Public Const SW_SHOW = 1
Public Const SW_HIDE = 0
Public Const SW_MINIMIZE = 6
Public Const SW_MAXIMIZE = 3
Public Const SW_MAX = 10

Public Const GWL_WNDPROC = (-4)
Public Const GWL_EXSTYLE = (-20)

Public Const LWA_COLORKEY = &H1
Public Const LWA_ALPHA = &H2

Public Const WS_EX_LAYERED = &H80000
Public Const WS_EX_TRANSPARENT   As Long = &H20&


Public Type myWindow
    Handle As Long
    Text As String
    Class As String * 255
    Parent As Long
    State As String
    ProcessID  As Long
    ThreadID As Long
End Type

Public Type WINDOWPLACEMENT
    Length As Long
    Flags As Long
    showCmd As Long
    ptMinPosition As POINTAPI
    ptMaxPosition As POINTAPI
    rcNormalPosition As RECT
End Type

Public Type WNDCLASS
    Style As Long
    lpfnWndProc As Long
    cbClsExtra As Long
    cbWndExtra2 As Long
    hInstance As Long
    hIcon As Long
    hCursor As Long
    hbrBackground As Long
    lpszMenuName As String
    lpszClassName As String
End Type

Public Enum EnumWindowsMethod
    MethodEnumWindows
    MethodParent
    '以下四个枚举所有窗口
    MethodPostMessage
    MethodEnumChildWindows
    MethodIsWindow
    MethodGetThread
End Enum

Public Enum CloseWindowMethod
    MethodPostCloseMsg
    MethodBombWindow
    MethodReplaceParentWindow
    MethodEndTask
End Enum

Public Enum WindowFilterMethod
    '标志位第1位：是否显示所有窗口
    MethodListAll = 1
    '标志位第2位：查找模式
    MethodSearch = 2
    '标志位第3、4位：过滤模式
    MethodListByPID = 4
    MethodListByTID = 8
    MethodListByParent = 12
End Enum

Public nWindow As myWindow
Public nChildWindow As myWindow
'Public viewProcessWindows As Long
Public nSelectedItem(256) As Long
Public nSelectedItemIndex(256) As Long

Public mWindowFilterMethod As WindowFilterMethod
Public mWindowFilterArg As Long
Public SetTab As Boolean

Private FindStr, cFindStr As String
Public Function WindowFilter() As Boolean
    WindowFilter = True
    If mWindowFilterMethod And MethodListAll Then
        Exit Function
    End If
    If mWindowFilterMethod And MethodSearch Then
        If Not (InStr(nWindow.Text, Menu.Text1.Text) > 0 Or _
        InStr(nWindow.Handle, Menu.Text1.Text) > 0 Or _
        InStr(nWindow.Class, Menu.Text1.Text) > 0) Then
            WindowFilter = False
            Exit Function
        End If
    End If
    Select Case mWindowFilterMethod And 12
    Case MethodListByPID
        If nWindow.ProcessID <> mWindowFilterArg Then
            WindowFilter = False
            Exit Function
        End If
    Case MethodListByTID
        If nWindow.ThreadID <> mWindowFilterArg Then
            WindowFilter = False
            Exit Function
        End If
    End Select
    If Menu.ListView1.Tag <> 0 Then
        WindowFilter = nWindow.Parent = nSelectedItem(Menu.ListView1.Tag)
    Else
        WindowFilter = (nWindow.Parent = 0)
    End If
End Function

Public Function CloseWindow(ByVal hMethod As CloseWindowMethod, ByVal hWnd As Long) As Boolean
    Select Case hMethod
    Case MethodPostCloseMsg
        PostMessage CLng(hWnd), WM_CLOSE, 0, ByVal 0
        PostMessage CLng(hWnd), WM_QUIT, 0, ByVal 0
        PostMessage CLng(hWnd), WM_DESTROY, 0, ByVal 0
        CloseWindow = True
    Case MethodBombWindow
        Call FxBombWindow(hWnd)
        CloseWindow = True
    Case MethodReplaceParentWindow
        Call FxCloseWindowByParent(hWnd)
        CloseWindow = True
    Case MethodEndTask
        CloseWindow = EndTask(CLng(hWnd), 0, 1) = 1
    End Select
End Function

Public Function EnumWindowsByMethod(ByVal hMethod As EnumWindowsMethod)
    If Menu.ListView1.Tag > 0 Then
        Menu.ListView1.ListItems.Add , , ".."
    End If
    Select Case hMethod
    Case MethodEnumWindows
        Call EnumWindowsProc(GetDesktopWindow, 0)
        Call EnumAllWindows(vbNullString)
    Case MethodPostMessage
        Dim hWnd As Long, tmp As Long
        For hWnd = 0 To 10000000
            If PostMessage(hWnd, 0, 0, 0) = 0 Then
                tmp = tmp + 2
                If tmp > 300000 Then Exit For
            Else
                tmp = 0
                Call EnumWindowsProc(hWnd, 0)
            End If
        Next
    Case MethodParent
        Call mnFxNew(Menu.Text1.Text)
    Case MethodIsWindow
        Dim i As Long, j As Long
        Do
            i = i + 2
            If IsWindow(i) Then
                If Menu.ListView1.Tag > 0 Then
                    If GetParent(i) = nSelectedItem(Menu.ListView1.Tag) Then
                        Call EnumWindowsProc(i, 0)
                    End If
                Else
                    If GetParent(i) = 0 Or GetParent(i) = GetDesktopWindow Then
                        Call EnumWindowsProc(i, 0)
                    End If
                End If
                j = 0
            Else
                j = j + 1
            End If
        Loop Until j = 300000
    Case MethodGetThread
        Call FxEnumWindowsByThread("")
    End Select
End Function

Public Sub EnumAllWindows(ByVal mfind As String)
    FindStr = mfind
    EnumWindows AddressOf EnumWindowsProc, ByVal 0&
    Menu.Label1.Caption = "共有：" & (Menu.ListView1.ListItems.Count) & "个活动窗体"
End Sub

Public Sub EnumAllChildWindows(ByVal pHwnd As Long, ByVal mfind As String)
    FindStr = mfind
    'Menu.ListView1.ListItems.Clear   '写出CNNew后取消这句
    Menu.ListView1.ListItems.Add , , ".."
    EnumChildWindows pHwnd, AddressOf EnumWindowsProc, ByVal 0&
    Menu.Label1.Caption = "共有：" & (Menu.ListView1.ListItems.Count) - 1 & "个子窗体"
End Sub

Public Function EnumWindowsProc(ByVal hWnd As Long, ByVal lParam As Long) As Boolean
    On Error Resume Next
    Dim text_len As Long
    
    With nWindow
        .Handle = hWnd
        .Text = ""
        .Class = ""
        .Parent = 0
        .ProcessID = 0
        .ThreadID = 0
    End With
    
    text_len = GetWindowTextLength(nWindow.Handle)
    nWindow.Text = Space(text_len + 1)
    GetWindowText nWindow.Handle, nWindow.Text, text_len + 1
    text_len = RealGetWindowClass(nWindow.Handle, nWindow.Class, 255)
    If LenB(StrConv(nWindow.Class, vbFromUnicode)) < text_len Then
        nWindow.Class = Space(text_len + 1)
        RealGetWindowClass nWindow.Handle, nWindow.Class, text_len
    End If
    nWindow.State = GetWindowState(nWindow.Handle)
    nWindow.Parent = GetParent(nWindow.Handle)
    nWindow.ThreadID = GetWindowThreadProcessId(nWindow.Handle, nWindow.ProcessID)
        
    If WindowFilter Then
        With Menu.ListView1.ListItems
        .Add , , left(nWindow.Text, InStr(1, nWindow.Text, vbNullChar))
        .Item(.Count).SubItems(1) = left(nWindow.Class, InStr(1, nWindow.Class, vbNullChar))
        .Item(.Count).SubItems(2) = nWindow.Handle
        .Item(.Count).SubItems(3) = nWindow.Parent
        .Item(.Count).SubItems(4) = nWindow.ProcessID
        .Item(.Count).SubItems(5) = nWindow.ThreadID
        .Item(.Count).SubItems(6) = nWindow.State
        End With
    End If
    
    EnumWindowsProc = True
End Function

Public Sub mnNew_Click()
    Dim nIndex As Long
    
    nIndex = FxGetListviewNowLine(Menu.ListView1)
    
    If Menu.ListView1.ListItems.Count > 0 Then
        nIndex = Menu.ListView1.SelectedItem.Index
    End If
    
    Menu.ListView1.ListItems.Clear
    If Menu.Text1.Text <> "输入标题或类名或句柄查找" And Menu.Text1.Text <> "" Then
        EnumAllWindows Menu.Text1.Text
    Else
        EnumAllWindows ""
    End If
    DoEvents
    
    FxSetListviewNowLine Menu.ListView1, nIndex
End Sub

'/**线程方法枚举窗口,保留,不使用**/
Public Sub mnFxtNew(mfind As String)
    Dim nIndex As Long
    
    nIndex = FxGetListviewNowLine(Menu.ListView1)
    
    Menu.ListView1.ListItems.Clear
    If Menu.Text1.Text <> "输入标题或类名或句柄查找" And Menu.Text1.Text <> "" Then
        FindStr = mfind
    Else
        FindStr = ""
    End If
    'FxEnumWindowsByThread 1, 600000
    'Call FEWBT_MultiThreading
       
    FxSetListviewNowLine Menu.ListView1, nIndex
End Sub

Public Function FxEnumWindowsByThread2(ByRef pa As String) As Long
    '函数原版
    Dim i, nHwnd As Long
    Dim nWindow As myWindow
    Dim TID, PID As Long
    Dim min, max As Long
    Dim except2 As String
    
    except2 = ""
    
    min = CLng(Mid(pa, 1, InStr(1, pa, ",")))
    max = CLng(Mid(pa, InStr(1, pa, ","), Len(pa)))
    
    For i = min To max
        If GetWindowThreadProcessId(i, PID) Then
            nHwnd = FxGetGrandParent(i)
            If InStr(1, except2, (nHwnd) & "$") = 0 Then
                EnumWindowsProc nHwnd, 0
                Menu.Label1.Caption = "共有：" & (Menu.ListView1.ListItems.Count) - 1 & "个活动窗体"
                except2 = (except2) & (nHwnd) & "$"
            End If
        End If
    Next i
    
    If Menu.ListView1.ListItems.Count > 0 Then
        FxEnumWindowsByThread2 = 1
    Else
        FxEnumWindowsByThread2 = 0
    End If
End Function

Public Function FxEnumWindowsByThread(ByRef pa As String) As Long
    '由richway2改写
    Dim hWnd As Long, i As Long
    Do
        hWnd = hWnd + 2
        If GetWindowThreadProcessId(hWnd, 0) Then
            EnumWindowsProc hWnd, 0
        Else
            i = i + 1
        End If
    Loop Until i = 600000
    FxEnumWindowsByThread = IIf(Menu.ListView1.ListItems.Count = 0, 0, 1)
End Function
'/**—————————————————
'/**————————————————————————**/

Public Sub mnFxNew(mfind As String)
    Dim nIndex As Long
    
    nIndex = FxGetListviewNowLine(Menu.ListView1)
    
    Menu.ListView1.ListItems.Clear
    If Menu.Text1.Text <> "输入标题或类名或句柄查找" And Menu.Text1.Text <> "" Then
        FindStr = mfind
    Else
        FindStr = ""
    End If
    FxEnumWindowsByParent
    
    DoEvents
    
    FxSetListviewNowLine Menu.ListView1, nIndex
End Sub

Public Function FxEnumWindowsByParent() As Long
    Dim i, nHwnd As Long
    Dim except As String
    
    except = ""
    
    For i = 1 To 1500000
        If GetParent(i) Then
            nHwnd = FxGetGrandParent(i)
            If InStr(1, except, "$" & (nHwnd) & "$") = 0 Then
                EnumWindowsProc nHwnd, 0
                Menu.Label1.Caption = "共有：" & (Menu.ListView1.ListItems.Count) - 1 & "个活动窗体"
                except = (except) & "$" & (nHwnd) & "$"
            End If
        End If
        'DoEvents
    Next i
    
    Menu.ListView1.Refresh
    
    If Menu.ListView1.ListItems.Count > 0 Then
        FxEnumWindowsByParent = 1
    Else
        FxEnumWindowsByParent = 0
    End If
End Function

Public Function FxGetGrandParent(ByVal hWindow As Long) As Long
    Dim hDesk, hLast As Long
    
    hDesk = GetDesktopWindow()
    
    Do
        hLast = hWindow
        hWindow = GetParent(hWindow)
    Loop Until hWindow = hDesk Or hWindow = 0
    
    FxGetGrandParent = hLast
End Function

Public Function GetWindowState(hWnd As Long) As String
    Dim mw As WINDOWPLACEMENT
    Dim we As String
    
    If IsWindowEnabled(hWnd) = 0 Then
        we = "冻结"
    Else
        we = "激活"
    End If
    
    If IsWindowVisible(hWnd) = 0 Then
        GetWindowState = "隐藏" & "/" & we
    Else
        GetWindowPlacement hWnd, mw
        If mw.showCmd = 1 Then
            GetWindowState = "可见" & "/" & we
        ElseIf mw.showCmd = 2 Then
            GetWindowState = "最小" & "/" & we
        ElseIf mw.showCmd = 3 Then
            GetWindowState = "最大" & "/" & we
        End If
    End If
End Function

Public Function FxCloseWindowByParent(ByVal aHwnd As Long) As Long
    Load FormKiller
    Call SetParent(aHwnd, FormKiller.hWnd) '返回原先
    FxCloseWindowByParent = GetLastError
    
    Unload FormKiller
End Function

Public Function FxCloseWindowByWndProc(ByVal aHwnd As Long) As Long
    Dim hFunction As Long

    hFunction = GetModuleHandle("kernel32.dll")
    hFunction = GetProcAddress(hFunction, "ExitProcess")
    
    MsgBox SetWindowLong(aHwnd, GWL_WNDPROC, hFunction) & "," & aHwnd

    FxCloseWindowByWndProc = 1
End Function

Public Sub nChildNewEx(ByVal Index As Long)
    Dim nIndex As Long
    
    nIndex = FxGetListviewNowLine(Menu.ListView1)
    
    Menu.ListView1.ListItems.Clear
    EnumAllChildWindows nSelectedItem(Index), ""

    FxSetListviewNowLine Menu.ListView1, nIndex
End Sub

Public Sub FxBombWindow(ByVal hWnd As Long)
    Dim i As Long
    
    ShowWindow hWnd, SW_HIDE   '先让窗口隐藏,防止轰炸后屏幕上有残象
    
    For i = 1 To 65535
        PostMessage hWnd, i, 0, ByVal 0
    Next i
End Sub

Public Sub SetWindowMethod(ByVal NewMethod As EnumWindowsMethod)
    Select Case Menu.Label1.Tag
    Case MethodEnumWindows
        Menu.nNew.Checked = False
    Case MethodParent
        Menu.nFxNew.Checked = False
    Case MethodPostMessage
        Menu.nFdNewByMessage.Checked = False
    Case MethodIsWindow
        Menu.nRwNewByIsWindow.Checked = False
    Case MethodGetThread
        Menu.nFxNewByTID.Checked = False
    End Select
    Select Case NewMethod
    Case MethodEnumWindows
        Menu.nNew.Checked = True
    Case MethodParent
        Menu.nFxNew.Checked = True
    Case MethodPostMessage
        Menu.nFdNewByMessage.Checked = True
    Case MethodIsWindow
        Menu.nRwNewByIsWindow.Checked = True
    Case MethodGetThread
        Menu.nFxNewByTID.Checked = True
    End Select
    Menu.Label1.Tag = NewMethod
End Sub

Public Sub CNNew()
    Dim nIndex As Long
    
    If Menu.ListView1.Sorted = True Then Menu.ListView1.Sorted = False
    
    nIndex = FxGetListviewNowLine(Menu.ListView1)
        
    Menu.ListView1.ListItems.Clear
    
    'LockWindowUpdate Menu.ListView1.hwnd
    
    If Menu.ListView1.Tag = 0 Then
        Call EnumWindowsByMethod(Menu.Label1.Tag)
    Else
        Call EnumWindowsByMethod(Menu.Label1.Tag)
    End If
    
    'If Menu.ListView1.Tag = 0 Then
    '    If Menu.Label1.Tag = 0 Then
    '        Call mnNew_Click
    '    ElseIf Menu.Label1.Tag = 1 Then
    '        Call mnFxNew(Menu.Text1.Text)
    '    End If
    'ElseIf Menu.ListView1.Tag = 1 Then
    '    EnumAllChildWindows nSelectedItem(Menu.ListView1.Tag), Menu.Text1.Text
    'End If

    FxSetListviewNowLine Menu.ListView1, nIndex
    Menu.Label1.Caption = "共有：" & Menu.ListView1.ListItems.Count & "个" & IIf(Menu.ListView1.Tag > 0, "子", "活动") & "窗体"
    
    'LockWindowUpdate 0
End Sub

Public Function GetText(ByVal hWnd As Long) As String
    Dim sStr As String, j As Long
    sStr = Space(1024)
    j = SendMessage(hWnd, WM_GETTEXT, Len(sStr), ByVal sStr)
    GetText = left(sStr, j)
End Function

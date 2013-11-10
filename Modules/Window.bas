Attribute VB_Name = "Window"
Option Explicit
Public Declare Function IsWindow Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetWindow Lib "user32" (ByVal hwnd As Long, ByVal wCmd As Long) As Long
Public Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function GetWindowRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Public Declare Function GetWindowPlacement Lib "user32" (ByVal hwnd As Long, lpwndpl As WINDOWPLACEMENT) As Long
Public Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function RealGetWindowClass Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Public Declare Function GetParent Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function GetDesktopWindow Lib "user32" () As Long
Public Declare Function SetWindowPos Lib "user32" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Public Declare Function SetParent Lib "user32" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Public Declare Function SetWindowText Lib "user32" Alias "SetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String) As Long
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function PostMessage Lib "user32" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Public Declare Function InvalidateRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT, ByVal bErase As Long) As Long
Public Declare Function InvalidateRectBynum Lib "user32" Alias "InvalidateRect" (ByVal hwnd As Long, ByVal lpRect As Long, ByVal bErase As Long) As Long
Public Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Boolean
Public Declare Function EnumChildWindows Lib "user32" (ByVal hWndParent As Long, ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Public Declare Function GetWindowTextLength Lib "user32" Alias "GetWindowTextLengthA" (ByVal hwnd As Long) As Long
Public Declare Function ShowWindow Lib "user32" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Public Declare Function SetLayeredWindowAttributes Lib "user32" (ByVal hwnd As Long, ByVal crKey As Long, ByVal bAlpha As Byte, ByVal dwFlags As Long) As Long
Public Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hwnd As Long, lpdwProcessId As Long) As Long
Public Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function IsWindowEnabled Lib "user32" (ByVal hwnd As Long) As Long
Public Declare Function EnableWindow Lib "user32" (ByVal hwnd As Long, ByVal fEnable As Long) As Long
Public Declare Function WindowFromPoint Lib "user32" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Public Declare Function DestroyWindow Lib "user32" (ByVal hwnd As Long) As Long
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
    flags As Long
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


Public nWindow As myWindow
Public nChildWindow As myWindow
Public viewProcessWindows As Long
Public nSelectedItem(256) As Long
Public nSelectedItemIndex(256) As Long

Private FindStr, cFindStr As String

Public Sub EnumAllWindows(ByVal mfind As String)
    FindStr = mfind
    EnumWindows AddressOf EnumWindowsProc, ByVal 0&
    Menu.Label1.Caption = "共有：" & (Menu.ListView1.ListItems.Count) & "个活动窗体"
End Sub

Public Sub EnumAllChildWindows(ByVal pHwnd As Long, ByVal mfind As String)
    cFindStr = mfind
    Menu.ListView1.ListItems.Clear   '写出CNNew后取消这句
    Menu.ListView1.ListItems.Add , , ".."
    EnumChildWindows pHwnd, AddressOf EnumWindowsProc, ByVal 0&
    Menu.Label1.Caption = "共有：" & (Menu.ListView1.ListItems.Count) - 1 & "个子窗体"
End Sub

Public Function EnumWindowsProc(ByVal hwnd As Long, ByVal lParam As Long) As Boolean
    On Error Resume Next
    Dim text_len As Long
    
    With nWindow
        .Handle = hwnd
        .Text = ""
        .Class = ""
        .Parent = 0
        .ProcessID = 0
        .ThreadID = 0
    End With
    
    text_len = GetWindowTextLength(nWindow.Handle)
    nWindow.Text = Space(text_len + 1)
    GetWindowText nWindow.Handle, nWindow.Text, text_len + 1
    RealGetWindowClass nWindow.Handle, nWindow.Class, 255
    nWindow.State = GetWindowState(nWindow.Handle)
    nWindow.Parent = GetParent(nWindow.Handle)
    nWindow.ThreadID = GetWindowThreadProcessId(nWindow.Handle, nWindow.ProcessID)
    
    Dim o, k As Long
    
    o = (left(nWindow.Text, 1) <> vbNullChar Or Menu.Check2.Value = 1) And lParam <> 3708

    If viewProcessWindows = 0 Then
        If o Then
            k = 1
            o = (InStr(1, LCase(nWindow.Text), LCase(FindStr)) > 0)
            o = o Or (InStr(1, LCase(nWindow.Class), LCase(FindStr)) > 0)
            o = o Or (InStr(1, LCase(CStr(nWindow.Handle)), LCase(CStr(FindStr))) > 0)
        End If
    ElseIf viewProcessWindows > 0 Then
        o = 1
        k = IIf(viewProcessWindows = nWindow.ProcessID, 1, 0)
    End If
        
    If o And k Then
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

Public Function FxEnumWindowsByThread(ByRef pa As String) As Long
    Dim i, nHwnd As Long
    Dim nWindow As myWindow
    Dim tid, pid As Long
    Dim min, max As Long
    Dim except2 As String
    
    except2 = ""
    
    min = CLng(Mid(pa, 1, InStr(1, pa, ",")))
    max = CLng(Mid(pa, InStr(1, pa, ","), Len(pa)))
    
    For i = min To max
        If GetWindowThreadProcessId(i, pid) Then
            nHwnd = FxGetGrandParent(i)
            If InStr(1, except2, (nHwnd) & "$") = 0 Then
                EnumWindowsProc nHwnd, 0
                Menu.Label1.Caption = "共有：" & (Menu.ListView1.ListItems.Count) - 1 & "个活动窗体"
                except2 = (except2) & (nHwnd) & "$"
            End If
        End If
    Next i
    
    If Menu.ListView1.ListItems.Count > 0 Then
        FxEnumWindowsByThread = 1
    Else
        FxEnumWindowsByThread = 0
    End If
End Function
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

Public Function GetWindowState(hwnd As Long) As String
    Dim mw As WINDOWPLACEMENT
    Dim we As String
    
    If IsWindowEnabled(hwnd) = 0 Then
        we = "冻结"
    Else
        we = "激活"
    End If
    
    If IsWindowVisible(hwnd) = 0 Then
        GetWindowState = "隐藏" & "/" & we
    Else
        GetWindowPlacement hwnd, mw
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
    SetParent aHwnd, FormKiller.hwnd
    Unload FormKiller
    
    FxCloseWindowByParent = 1
End Function

Public Function FxCloseWindowByWndProc(ByVal aHwnd As Long) As Long
    Dim hFunction As Long

    hFunction = GetModuleHandle("kernel32.dll")
    hFunction = GetProcAddress(hFunction, "ExitProcess")
    
    MsgBox SetWindowLong(aHwnd, GWL_WNDPROC, hFunction) & "," & aHwnd

    FxCloseWindowByWndProc = 1
End Function

Public Sub FxBombWindow(ByVal hwnd As Long)
    Dim i As Long
    
    ShowWindow hwnd, SW_HIDE   '先让窗口隐藏,防止轰炸后屏幕上有残象
    
    For i = 1 To 65535
        PostMessage hwnd, i, 0, ByVal 0
    Next i
End Sub

Public Sub CNNew()
    Dim nIndex As Long
    
    If Menu.ListView1.Sorted = True Then Menu.ListView1.Sorted = False
    
    nIndex = FxGetListviewNowLine(Menu.ListView1)
        
    Menu.ListView1.ListItems.Clear
    
    'LockWindowUpdate Menu.ListView1.hwnd
    
    If Menu.ListView1.Tag = 0 Then
        If Menu.Label1.Tag = 0 Then
            Call mnNew_Click
        ElseIf Menu.Label1.Tag = 1 Then
            Call mnFxNew(Menu.Text1.Text)
        End If
    ElseIf Menu.ListView1.Tag = 1 Then
        EnumAllChildWindows nSelectedItem(Menu.ListView1.Tag), Menu.Text1.Text
    End If

    FxSetListviewNowLine Menu.ListView1, nIndex
    
    'LockWindowUpdate 0
End Sub

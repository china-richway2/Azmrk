Attribute VB_Name = "Basic"
Option Explicit
'--------------set text color--------------------
Public Declare Function ChooseColorAPI Lib "comdlg32.dll" Alias "ChooseColorA" (pChoosecolor As CHOOSECOLOR) As Long

Public Type CHOOSECOLOR
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    rgbResult As Long
    lpCustColors As String
    Flags As Long
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type
'--------------set text color--------------------
'-------------reg ocx---------------
Private Declare Function RegComCtl32 Lib "msComCtl32.OCX" Alias "DllRegisterServer" () As Long
Private Declare Function RegTabCtl32 Lib "msTabCtl32.OCX" Alias "DllRegisterServer" () As Long
Private Declare Function RegComDlg32 Lib "msComDlg32.OCX" Alias "DllRegisterServer" () As Long
'------------reg ocx ---------2010.02.20-----------
Private Declare Function CreateMutex Lib "kernel32" Alias "CreateMutexA" (ByVal lpMutexAttributes As Long, ByVal bInitialOwner As Long, ByVal lpName As String) As Long
Private Declare Function ReleaseMutex Lib "kernel32" (ByVal hMutex As Long) As Long
'Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long
'-----------------skin api---------------------
Public Declare Function SkinH_Attach Lib "SkinH_VB6.dll" () As Long
Public Declare Function SkinH_AttachEx Lib "SkinH_VB6.dll" (ByVal lpSkinFile As String, ByVal lpPasswd As String) As Long
Public Declare Function SkinH_AttachExt Lib "SkinH_VB6.dll" (ByVal lpSkinFile As String, ByVal lpPasswd As String, ByVal nHue As Integer, ByVal nSat As Integer, ByVal nBri As Integer) As Long
Public Declare Function SkinH_AttachRes Lib "SkinH_VB6.dll" (lpRes As Any, ByVal nSize As Long, ByVal lpPasswd As String, ByVal nHue As Integer, ByVal nSat As Integer, ByVal nBri As Integer) As Long
Public Declare Function SkinH_AdjustHSV Lib "SkinH_VB6.dll" (ByVal nHue As Integer, ByVal nSat As Integer, ByVal nBri As Integer) As Long
Public Declare Function SkinH_Detach Lib "SkinH_VB6.dll" () As Long
Public Declare Function SkinH_DetachEx Lib "SkinH_VB6.dll" (ByVal hWnd As Long) As Long
Public Declare Function SkinH_SetAero Lib "SkinH_VB6.dll" (ByVal hWnd As Long) As Long
Public Declare Function SkinH_SetWindowAlpha Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal nAlpha As Integer) As Long
Public Declare Function SkinH_SetMenuAlpha Lib "SkinH_VB6.dll" (ByVal nAlpha As Integer) As Long
Public Declare Function SkinH_GetColor Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal nPosX As Integer, ByVal nPosY As Integer) As Long
Public Declare Function SkinH_Map Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal nType As Integer) As Long
Public Declare Function SkinH_LockUpdate Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal nLocked As Integer) As Long
Public Declare Function SkinH_SetBackColor Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal nRed As Integer, ByVal nGreen As Integer, ByVal nBlue As Integer) As Long
Public Declare Function SkinH_SetForeColor Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal nRed As Integer, ByVal nGreen As Integer, ByVal nBlue As Integer) As Long
Public Declare Function SkinH_SetWindowMovable Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal bMove As Integer) As Long
Public Declare Function SkinH_AdjustAero Lib "SkinH_VB6.dll" (ByVal nAlpha As Integer, ByVal nShwDark As Integer, ByVal nShwSharp As Integer, ByVal nShwSize As Integer, ByVal nX As Integer, ByVal nY As Integer, ByVal nRed As Integer, ByVal nGreen As Integer, ByVal nBlue As Integer) As Long
Public Declare Function SkinH_NineBlt Lib "SkinH_VB6.dll" (ByVal hDtDC As Long, ByVal left As Integer, ByVal top As Integer, ByVal right As Integer, ByVal bottom As Integer, ByVal nMRect As Integer) As Long
Public Declare Function SkinH_SetTitleMenuBar Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal bEnable As Integer, ByVal nMenuY As Integer, ByVal nTopOffs As Integer, ByVal nRightOffs As Integer) As Long
Public Declare Function SkinH_SetFont Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal hFont As Long) As Long
Public Declare Function SkinH_SetFontEx Lib "SkinH_VB6.dll" (ByVal hWnd As Long, ByVal szFace As String, ByVal nHeight As Integer, ByVal nWidth As Integer, ByVal nWeight As Integer, ByVal nItalic As Integer, ByVal nUnderline As Integer, ByVal nStrikeOut As Integer) As Long
Public Declare Function SkinH_VerifySign Lib "SkinH_VB6.dll" () As Long
'-----------------------------------------------
Public Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long
Public Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function ClipCursor Lib "user32" (lpRect As Any) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShellExecuteEx Lib "shell32" (lpSEI As ShellEexeCuteInfo) As Long
Public Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Public Declare Function ZwQueryObject Lib "NTDLL.DLL" (ByVal ObjectHandle As Long, ByVal ObjectInformationClass As String, ByRef ObjectInformation As Long, ByVal ObjectInformationLength As Long, ByRef ReturnLength As Long) As Long
Public Declare Function HEREWHERE Lib "AzmrkHelper" () As Long
'----------------read ini files------------------------
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'-------------set the icon---------------------
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal nCursorType As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
'-------------skin const---------------------
Public Const SM_CXICON = 11
Public Const SM_CYICON = 12
Public Const SM_CXSMICON = 49
Public Const SM_CYSMICON = 50

Public Const LR_DEFAULTCOLOR = &H0
Public Const LR_MONOCHROME = &H1
Public Const LR_COLOR = &H2
Public Const LR_COPYRETURNORG = &H4
Public Const LR_COPYDELETEORG = &H8
Public Const LR_LOADFROMFILE = &H10
Public Const LR_LOADTRANSPARENT = &H20
Public Const LR_DEFAULTSIZE = &H40
Public Const LR_VGACOLOR = &H80
Public Const LR_LOADMAP3DCOLORS = &H1000
Public Const LR_CREATEDIBSECTION = &H2000
Public Const LR_COPYFROMRESOURCE = &H4000
Public Const LR_SHARED = &H8000&

Public Const IMAGE_ICON = 1

Public Const ICON_SMALL = 0
Public Const ICON_BIG = 1

Public Const GW_OWNER = 4

Public Const COLOR_WINDOW = 5

Public Const CS_VREDRAW = &H1
Public Const CS_HREDRAW = &H2

Public Const IDC_ARROW = 32512&
Public Const IDC_IBEAM = 32513&
Public Const IDC_WAIT = 32514&
Public Const IDC_CROSS = 32515&
Public Const IDC_UPARROW = 32516&
Public Const IDC_SIZE = 32640&
Public Const IDC_ICON = 32641&
Public Const IDC_SIZENWSE = 32642&
Public Const IDC_SIZENESW = 32643&
Public Const IDC_SIZEWE = 32644&
Public Const IDC_SIZENS = 32645&
Public Const IDC_SIZEALL = 32646&
Public Const IDC_NO = 32648&
Public Const IDC_HAND = 32649&
Public Const IDC_APPSTARTING = 32650&
Public Const IDC_HELP = 32651&
'----------------------------------------------
Public Const SEE_MASK_INVOKEIDLIST = &HC

Public Const LVM_FIRST As Long = &H1000
Public Const LVM_GETSELECTIONMARK As Long = (LVM_FIRST + 66)
'----------------------------------------------
Public Const STANDARD_RIGHTS_ALL = &H1F0000
Public Const SYNCHRONIZE = &H100000
Public Const STANDARD_RIGHTS_REQUIRED = &HF0000

Public Type ShellEexeCuteInfo
    cbSize As Long
    fMask As Long
    hWnd As Long
    lpVerb As String
    lpFile As String
    lpParameters As String
    lpDirectory As String
    nShow As Long
    hInstApp As Long
    lpIDList As Long
    lpClass As String
    hkeyClass As Long
    dwHotKey As Long
    hIcon As Long
    hProcess As Long
End Type

Public Type POINTAPI
    x As Long
    y As Long
End Type

Public Type RECT
    left As Long
    top As Long
    right As Long
    bottom As Long
End Type

Public Type FILETIME
    dwLowDateTime As Long
    dwHighDateTime As Long
End Type

Private Ret As Long

Public AllColor As Long '全局颜色
Public VisualValue(11) As String
Public SoftValue(2) As String
Public NonLoading As Boolean
Public ObjectTypeNames(100) As String
Public Const ProcessColumnCount As Long = 14
Public ProcessColumnSetting(ProcessColumnCount) As String '选择列设置
Public ProcessColumnNames() As String '显示的列名
Public ProcessColumnWidth(ProcessColumnCount) As Long '选择列宽度
Public RealProcessColumnNames() As String '列名顺序

Public Function Output(ByVal Num As Long, ByVal Types As String, ByVal FileName As String) As Long '输出资源文件
    Dim TempFile() As Byte
    Dim FileNum    As Integer
    Dim TempDir    As String
    TempDir = App.Path
    If right(TempDir, 1) <> "\" Then TempDir = TempDir & "\"

    If Dir(TempDir & FileName) = "" Then
        TempFile = LoadResData(Num, Types)
        FileNum = FreeFile
        Open TempDir & FileName For Binary Access Write As #FileNum
            Put #FileNum, , TempFile
        Close #FileNum
        Output = True
    Else
        Output = False
    End If

End Function

Public Sub ShowFileProperties(ByVal aFile As String)
    Dim Info As ShellEexeCuteInfo

    With Info
        .hWnd = ModuleList.hWnd
        .lpVerb = "properties"
        .lpFile = aFile
        .fMask = SEE_MASK_INVOKEIDLIST
        .cbSize = Len(Info)
    End With
    ShellExecuteEx Info
End Sub

Public Sub LoginTransport(ByVal hWnd As Long)
    SetWindowPos hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE

    Dim rtn As Long, i As Long
    rtn = GetWindowLong(hWnd, -20)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong hWnd, -20, rtn

    SetLayeredWindowAttributes hWnd, 0, 0, LWA_ALPHA
    ShowWindow hWnd, SW_SHOW
    For i = 55 To 255 Step 3
        SetLayeredWindowAttributes hWnd, 0, i, LWA_ALPHA
        Sleep 15
    Next
    
    ZwTerminateThread -2, 0
End Sub

Public Sub SetStatus(ByRef lpText As String)
    LoginPic.Status = lpText
End Sub

Public Function CheckFor(FileName As String, DownloadPath As String) As Integer
    On Error GoTo E
    Open FileName For Input As #1
    Close #1
    CheckFor = 0
    Exit Function
E:
    If Err.Number = 53 Then
        If DownloadPath <> "" Then
            If MsgBox(FindString("CheckFor.NotFound1") & FileName & FindString("CheckFor.NotFound2"), vbYesNo + vbQuestion) = vbYes Then
                ShellExecute 0, "open", DownloadPath, vbNullString, vbNullString, 0
                End
            End If
        End If
        CheckFor = 2
    Else
        CheckFor = 1
    End If
End Function

Public Sub ExistCheck(FileName As String, Path As String)
    If CheckFor(FileName, Path) = 2 Then
        MsgBox FileName & FindString("ExistCheck.NotFound"), vbCritical
        End
    End If
End Sub

Private Sub OutputStrings(p As Form)
    Open p.Name & ".txt" For Output As #1
    Dim i
    On Error Resume Next
    Print #1, p.Name & "=" & p.Caption
    For Each i In p.Controls
        Dim s As String
        s = i.Caption
        If s <> "" Then
            Print #1, i.Name & "=" & s
            s = ""
        End If
    Next
    Close #1
End Sub

Public Sub Main()
    On Error Resume Next
    
    
    ExistCheck "mscomdlg32.ocx", "http://url.cn/WSPgQ5"
    ExistCheck "tabctl32.ocx", "http://url.cn/JDI5yv"
    ExistCheck "mscomctl32.ocx", "http://url.cn/OoIsU1"
    If CheckFor("config.ini", "") = 2 Then
        Open "config.ini" For Output As #1
        Print #1, "[Visual settings]"
        Print #1, "Hue=0"
        Print #1, "Saturation=0"
        Print #1, "Brightness=0"
        Print #1, "Alpha=120"
        Print #1, "Shadow Size=8"
        Print #1, "Shadow Sharpness=10"
        Print #1, "Shadow Darkness=120"
        Print #1, "Shadow Color R=0"
        Print #1, "Shadow Color G=0"
        Print #1, "Shadow Color B=0"
        Print #1, "Menu Alpha=255"
        Print #1, "Skin=0"
        Print #1, "Text Color=0"
        Print #1, "[Soft Settings]"
        Print #1, "Always on top=0"
        Print #1, "Show all windows=0"
        Print #1, "Follow Mouse=0"
        Print #1, "Enum windows method=0"
        Print #1, "[Process Column]"
        Print #1, ";以下项名（=左边的内容）可以修改，会修改Azmrk显示的ColumnHeader名称 =右边的内容是是否显示项，-1代表显示，0代表不显示"
        Print #1, ";不可以删除"
        Print #1, "进程名=-1"
        Print #1, "进程ID=-1"
        Print #1, "父进程ID=-1"
        Print #1, "Peb=-1"
        Print #1, "EPROCESS=-1"
        Print #1, "优先级=-1"
        Print #1, "内存使用=-1"
        Print #1, "IO读取次数=-1"
        Print #1, "IO写入次数=-1"
        Print #1, "IO其他次数=-1"
        Print #1, "IO读取字节=-1"
        Print #1, "IO写入字节=-1"
        Print #1, "IO其他字节=-1"
        Print #1, "映像路径=-1"
        Print #1, "命令行=-1"
        Print #1, "[Process Column Width]"
        Print #1, ";以下项名不可以修改，但是可以交换顺序来交换项的顺序，等号右边为宽度 不可以删除；交换时注意同时交换上面的内容"
        Print #1, ";如交换以下的EPROCESS和PEB后如果不交换上面的EPROCESS和PEB会导致EPROCESS项显示PEB，PEB项显示EPROCESS"
        Print #1, "进程名=1500"
        Print #1, "进程ID=1005"
        Print #1, "父进程ID=1005"
        Print #1, "Peb=1125"
        Print #1, "EPROCESS=1125"
        Print #1, "优先级=1080"
        Print #1, "内存使用=1905"
        Print #1, "IO读取次数=1440"
        Print #1, "IO写入次数=1440"
        Print #1, "IO其他次数=1440"
        Print #1, "IO读取字节=1440"
        Print #1, "IO写入字节=1440"
        Print #1, "IO其他字节=1440"
        Print #1, "映像路径=1500"
        Print #1, "命令行=1500"
        Print #1, "[Language]"
        Print #1, "LanguagePack=中文"
        Close #1
    End If
    '语言设置
    Dim s As String
    ReadINI "Language", "LanguagePack", s, "language.lang"
    s = left(s, InStr(s, Chr(0)) - 1)
    If right(s, 5) <> ".lang" Then
        s = s & ".lang"
    End If
    Open s For Input As #1
    Close #1
    LoadLanguage s
        
    Dim Thread As Long
    SetStatus FindString("Main.Initialize")
    EnablePrivilege SE_DEBUG
    InitSSDTableModule
    SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
    'viewProcessWindows = False
    
    
    Dim rtn As Long
    Dim i As Long
 
    Load LoginPic

    rtn = GetWindowLong(LoginPic.hWnd, -20)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong LoginPic.hWnd, -20, rtn

    SetLayeredWindowAttributes LoginPic.hWnd, 0, 0, LWA_ALPHA
    LoginPic.Show
    rtn = begin_thread(AddressOf LoginTransport, VarPtr(LoginPic.hWnd), 1)
    
    SetStatus FindString("Main.ReadConfig")
    Dim TempStr       As String
    Dim VisualTitle() As String '视觉设置
    Dim SoftSetting() As String '软件设置
    Call ReadINI("Visual settings", vbNullString, TempStr)
    
    VisualTitle = Split(TempStr, Chr(0)) '获得skin记录信息
    Call ReadINI("Soft Settings", vbNullString, TempStr, "")
    SoftSetting = Split(TempStr, Chr(0))
    Call ReadINI("Process Column", vbNullString, TempStr, "")
    ProcessColumnNames = Split(TempStr, Chr(0))
    Call ReadINI("Process Column Width", vbNullString, TempStr, "")
    RealProcessColumnNames = Split(TempStr, Chr(0))
    
    For i = 0 To UBound(VisualValue)
        ReadINI "Visual settings", VisualTitle(i), VisualValue(i)
    Next

    ReadINI "visual settings", "Text Color", VisualValue(11)
    AllColor = VisualValue(11)

    For i = 0 To 2
        ReadINI "Soft Settings", SoftSetting(i), SoftValue(i)
    Next
    
    For i = 0 To 14
        ReadINI "Process Column", ProcessColumnNames(i), ProcessColumnSetting(i), ""
        ReadINI "Process Column Width", RealProcessColumnNames(i), TempStr, ""
        ProcessColumnWidth(i) = Val(TempStr)
    Next
    

    'Dim wBuffer As SYSTEM_HANDLE_TABLE_ENTRY_INFO
    'ZwDuplicateObject ZwGetCurrentProcess, ZwGetCurrentProcess, ZwGetCurrentProcess, i, PROCESS_ALL_ACCESS, 0, DUPLICATE_SAME_ATTRIBUTES
    'RdQueryHandleInformation i, wBuffer, -1
    'OB_TYPE_PROCESS = wBuffer.ObjectTypeIndex
    'ZwClose i
    'ZwDuplicateObject ZwGetCurrentProcess, ZwGetCurrentThread, ZwGetCurrentProcess, i, THREAD_ALL_ACCESS, 0, DUPLICATE_SAME_ATTRIBUTES
    'RdQueryHandleInformation i, wBuffer, -1
    'OB_TYPE_THREAD = wBuffer.ObjectTypeIndexs
    'ZwClose i
    ObjectTypeNames(OB_TYPE_UNKNOWN) = "OB_TYPE_UNKNOWN"
    ObjectTypeNames(OB_TYPE_TYPE) = "OB_TYPE_TYPE"
    ObjectTypeNames(OB_TYPE_DIRECTORY) = "OB_TYPE_DIRECTORY"
    ObjectTypeNames(OB_TYPE_SYMBOLIC_LINK) = "OB_TYPE_SYMBOLIC_LINK"
    ObjectTypeNames(OB_TYPE_TOKEN) = "OB_TYPE_TOKEN"
    ObjectTypeNames(OB_TYPE_PROCESS) = "OB_TYPE_PROCESS"
    ObjectTypeNames(OB_TYPE_THREAD) = "OB_TYPE_THREAD"
    ObjectTypeNames(OB_TYPE_JOB) = "OB_TYPE_JOB"
    ObjectTypeNames(OB_TYPE_DEBUG_OBJECT) = "OB_TYPE_DEBUG_OBJECT"
    ObjectTypeNames(OB_TYPE_EVENT) = "OB_TYPE_EVENT"
    ObjectTypeNames(OB_TYPE_EVENT_PAIR) = "OB_TYPE_EVENT_PAIR"
    ObjectTypeNames(OB_TYPE_MUTANT) = "OB_TYPE_MUTANT"
    ObjectTypeNames(OB_TYPE_CALLBACK) = "OB_TYPE_CALLBACK"
    ObjectTypeNames(OB_TYPE_SEMAPHORE) = "OB_TYPE_SEMAPHORE"
    ObjectTypeNames(OB_TYPE_TIMER) = "OB_TYPE_TIMER"
    ObjectTypeNames(OB_TYPE_PROFILE) = "OB_TYPE_PROFILE"
    ObjectTypeNames(OB_TYPE_KEYED_EVENT) = "OB_TYPE_KEYED_EVENT"
    ObjectTypeNames(OB_TYPE_WINDOWS_STATION) = "OB_TYPE_WINDOWS_STATION"
    ObjectTypeNames(OB_TYPE_DESKTOP) = "OB_TYPE_DESKTOP"
    ObjectTypeNames(OB_TYPE_SECTION) = "OB_TYPE_SECTION"
    ObjectTypeNames(OB_TYPE_KEY) = "OB_TYPE_KEY"
    ObjectTypeNames(OB_TYPE_PORT) = "OB_TYPE_PORT"
    ObjectTypeNames(OB_TYPE_WAITABLE_PORT) = "OB_TYPE_WAITABLE_PORT"
    ObjectTypeNames(OB_TYPE_ADAPTER) = "OB_TYPE_ADAPTER"
    ObjectTypeNames(OB_TYPE_CONTROLLER) = "OB_TYPE_CONTROLLER"
    ObjectTypeNames(OB_TYPE_DEVICE) = "OB_TYPE_DEVICE"
    ObjectTypeNames(OB_TYPE_DRIVER) = "OB_TYPE_DRIVER"
    ObjectTypeNames(OB_TYPE_IOCOMPLETION) = "OB_TYPE_IOCOMPLETION"
    ObjectTypeNames(OB_TYPE_FILE) = "OB_TYPE_FILE"
    ObjectTypeNames(OB_TYPE_WMIGUID) = "OB_TYPE_WMIGUID"
    
    '/**图标设置
    'SetIcon Menu.hwnd, "IDR_MAINFRAME", True
    'SetIcon State.hwnd, "IDR_MAINFRAME", True
    '**/图标设置
    
    Load Menu
    With Menu
        SetStatus FindString("Main.LoadSkin")
        .SetVisual VisualValue, SoftValue
        DoEvents
        .Label1.Tag = "0"
        'ListViewColor Menu, Menu.ListView2
        SetStatus FindString("Main.LoadWindow")
        Call CNNew
        DoEvents
        SetStatus FindString("Main.EnumProcess")
        Call PNNew
        DoEvents
        SetStatus FindString("Main.EnumService")
        Call msNew_Click
        DoEvents
        SetStatus FindString("Main.EnumDrivers")
        Call GMNew
        DoEvents
        nSelectedItem(0) = .ListView1.ListItems(1).SubItems(2)
        SetStatus FindString("Main.Starting")
        WaitForSingleObject rtn, 100000
        ZwClose rtn
        Unload LoginPic
        .Show
        .Refresh
    End With
End Sub

'---------------------set the icon----------------
Public Sub SetIcon(ByVal hWnd As Long, _
                   ByVal sIconResName As String, _
                   Optional ByVal bSetAsAppIcon As Boolean = True)
    Dim lhWndTop   As Long
    Dim lhwnd      As Long
    Dim cx         As Long
    Dim cy         As Long
    Dim hIconLarge As Long
    Dim hIconSmall As Long
      
    If (bSetAsAppIcon) Then
        ' 查找VB隐藏的父窗体:
        lhwnd = hWnd
        lhWndTop = lhwnd

        Do While Not (lhwnd = 0)
            lhwnd = GetWindow(lhwnd, GW_OWNER)

            If Not (lhwnd = 0) Then
                lhWndTop = lhwnd
            End If

        Loop

    End If
   
    cx = GetSystemMetrics(SM_CXICON)
    cy = GetSystemMetrics(SM_CYICON)
   
    hIconLarge = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)
   
End Sub

'------------set the icon-----------------------
Public Function ListViewColor(ByRef FrmName As Form, _
                              ByRef ListName As Listview, _
                              Optional ByVal objName As String, _
                              Optional N1Color As Long = &HFFFFFF, _
                              Optional N2Color As Long = &HFFC0C0) As Long
    On Error GoTo errorHand
    Dim iHeight As Double
    Dim iNull   As Boolean
    Dim itmX    As ListItem
    Dim PicName As PictureBox
    Dim i       As Long
    
    Const WordList = "qwertyuiopasdfghjklzxcvbnmQWERTYUIOPLKJHGFDSAZXCVBNM"

    If objName = "" Then

        For i = 1 To 10
            objName = objName & Mid$(WordList, Int(Rnd * 52) + 1, 1)
        Next

    End If

    Set PicName = FrmName.Controls.Add("vb.PictureBox", objName, FrmName)

    If ListName.ListItems.count <= 0 Then
        Set itmX = ListName.ListItems.Add()
        itmX.Text = "test........"
        iNull = True
    End If

    With PicName
        .AutoRedraw = True
        .ScaleMode = vbTwips
        .Font = ListName.Font
        .BorderStyle = 1
        .Appearance = 0
        iHeight = ListName.ListItems(1).Height
        .Height = iHeight * 2 + 30
    End With

    With ListName
        PicName.Line (0, 0)-(.Width, iHeight), N1Color, BF
        PicName.Line (0, iHeight)-(.Width, iHeight * 2), N2Color, BF
        .PictureAlignment = lvwTile
        .Picture = PicName.Image

        If iNull = True Then
            .ListItems.Clear
        End If

    End With

    FrmName.ScaleMode = vbTwips
errorHand:
    Exit Function
End Function

Public Function WriteINI(ByVal BasicTitle As String, ByVal SelectionTitle As String, ByVal ValueTitle As String, Optional ByVal IniPath As String) As Long
    If IniPath = "" Then IniPath = App.Path & "\config.ini"
    WriteINI = WritePrivateProfileString(BasicTitle, SelectionTitle, ValueTitle, IniPath)
End Function

Public Function ReadINI(ByVal BasicTitle As String, ByVal SelectionTitle As String, ByRef ValueTitle As String, Optional ByVal DefaultValue As String = "", Optional ByVal IniPath As String)
    If IniPath = "" Then IniPath = App.Path & "\config.ini"
    ValueTitle = Space$(256)
    ReadINI = GetPrivateProfileString(BasicTitle, SelectionTitle, DefaultValue, ValueTitle, 256, IniPath)
End Function

Public Sub FindFiles(ByVal Path As String, Optional ByVal OpenStyle As Long = vbNormalFocus) '定位文件
    Shell "explorer.exe /select," & Path, OpenStyle
End Sub

Public Function ByteToKMG(ByVal b As String) As String
    'B = 1024 ^ 3
    Select Case Len(b)
    
        Case Is >= 10 'G
            ByteToKMG = Format(Val(b) / 1024 ^ 3, "#0.00G")
        Case Is >= 7 'M
            ByteToKMG = Format(Val(b) / 1024 ^ 2, "##0.00M")
        Case Is >= 4 'K
            ByteToKMG = Format(Val(b) / 1024, "##0.00K")
        Case Else 'B
            ByteToKMG = Format(b, "##0.00B")
    
    End Select
End Function

Public Sub LVAutoOrder(LV As Listview, ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    If LV.Sorted = False Then LV.Sorted = True
    LV.SortKey = ColumnHeader.Index - 1
    '实现点击后可转换排序是从A-Z还是Z-A
    LV.SortOrder = 1 - LV.SortOrder
End Sub

Public Function SetTextColor(ByRef Frm As Form, Optional ByVal TColor As Long) As Long
    Dim i As Long
    On Error Resume Next
    If TColor = 0 Then TColor = AllColor
    If TColor = 0 Then Exit Function

    For i = 0 To Frm.count - 1
        Frm.Controls(i).ForeColor = TColor
    Next

    AllColor = TColor
    WriteINI "Visual settings", "Text Color", TColor
End Function

Public Function AutoUpdate(ByRef txt As TextBox)
    Dim T As Long

    txt.LinkMode = 0

    txt.LinkTopic = "AutoUpdate|frmUpdate"

    txt.LinkMode = 2

    txt.LinkExecute (Val(Format(Now, "yymmddhhmm")) Xor 903100000 And 175564877)

    T = txt.LinkTimeout

    txt.LinkTimeout = 1

    txt.LinkMode = 0

    txt.LinkTimeout = T

End Function

Public Function FormatHex(ByVal Num As Long) As String
    '使用方法：
    '传递10进制数：Format16H(16)=0x00000010
    '传递16进制数：Format16H(&HF)=0x00000010
    Dim i As Long
    Dim Temp As String

    Temp = Hex(Num)
    i = 8 - Len(Temp)

    FormatHex = "0x" & String$(i, "0") & Temp
End Function

Public Function UnFormatHex(ByVal Num As String) As Long
    On Error Resume Next
    UnFormatHex = Val("&H" & Mid(Num, 3))
End Function

Public Function FxGetListviewNowLine(ByRef Listview As Object)
    Dim nIndex As Long
    
    nIndex = 1
    If Listview.ListItems.count > 0 Then
        nIndex = Listview.SelectedItem.Index
    End If
    
    FxGetListviewNowLine = nIndex
End Function

Public Sub FxSetListviewNowLine(ByRef Listview As Object, ByVal nIndex As Long)
    If Listview.ListItems.count >= nIndex Then
        Listview.ListItems(nIndex).Selected = True
        Listview.ListItems(nIndex).EnsureVisible
    End If
End Sub

Public Function UnicodeStringToString(ByRef us As UNICODE_STRING) As String
    UnicodeStringToString = Space(us.Length \ 2)
    CopyMemory StrPtr(UnicodeStringToString), us.Buffer, us.Length
End Function

Public Function StringFromPtr(ByVal Ptr As Long) As String
    Dim us As UNICODE_STRING
    RtlInitUnicodeString us, Ptr
    StringFromPtr = UnicodeStringToString(us)
End Function

Public Function AnsiStringFromPtr(ByVal Ptr As Long) As String
    Dim buf() As Byte, n As Long
    n = lstrlenA(Ptr)
    If n = 0 Then Exit Function
    ReDim buf(n - 1)
    CopyMemory VarPtr(buf(0)), Ptr, n
    AnsiStringFromPtr = StrConv(buf, vbUnicode)
End Function

Public Function NT_SUCCESS(ByVal Status As Long) As Boolean
    NT_SUCCESS = (Status >= 0)
End Function

Public Function AddUnsigned(lX As Long, lY As Long) As Long
    Dim lX4 As Long, lY4 As Long, lX8 As Long, lY8 As Long, lResult As Long
    lX8 = lX And &H80000000
    lY8 = lY And &H80000000
    lX4 = lX And &H40000000
    lY4 = lY And &H40000000

    lResult = (lX And &H3FFFFFFF) + (lY And &H3FFFFFFF)

    If lX4 And lY4 Then
        lResult = lResult Xor &H80000000 Xor lX8 Xor lY8
    ElseIf lX4 Or lY4 Then
        If lResult And &H40000000 Then
            lResult = lResult Xor &HC0000000 Xor lX8 Xor lY8
        Else
            lResult = lResult Xor &H40000000 Xor lX8 Xor lY8
        End If
    Else
        lResult = lResult Xor lX8 Xor lY8
    End If

    AddUnsigned = lResult
End Function

Public Function ReturnPtr(ByVal n As Long) As Long
    ReturnPtr = n
End Function

Public Function IsIDE() As Boolean
    Debug.Assert GetTrue(IsIDE)
End Function

Private Function GetTrue(A As Boolean) As Boolean
    GetTrue = True
    A = True
End Function

Public Function Assert(ByVal bBool As Boolean, ByVal sString As String, ByVal Quiet As Boolean) As Boolean
    Assert = Not bBool
    If Not bBool Then
        MsgBox sString, vbCritical
        Debug.Assert False
        ExitProcess 0
    End If
End Function

Public Function Int2Long(ByVal n As Integer) As Long
    If n < 0 Then Int2Long = n + 65536 Else Int2Long = n
End Function

Public Function Hex2(ByVal nHex As Long, ByVal A As Long) As String
    Hex2 = right(String(A, "0") & Hex(nHex), A)
End Function

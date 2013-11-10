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
    flags As Long
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
Public Declare Function SkinH_DetachEx Lib "SkinH_VB6.dll" (ByVal hwnd As Long) As Long
Public Declare Function SkinH_SetAero Lib "SkinH_VB6.dll" (ByVal hwnd As Long) As Long
Public Declare Function SkinH_SetWindowAlpha Lib "SkinH_VB6.dll" (ByVal hwnd As Long, ByVal nAlpha As Integer) As Long
Public Declare Function SkinH_SetMenuAlpha Lib "SkinH_VB6.dll" (ByVal nAlpha As Integer) As Long
Public Declare Function SkinH_GetColor Lib "SkinH_VB6.dll" (ByVal hwnd As Long, ByVal nPosX As Integer, ByVal nPosY As Integer) As Long
Public Declare Function SkinH_Map Lib "SkinH_VB6.dll" (ByVal hwnd As Long, ByVal nType As Integer) As Long
Public Declare Function SkinH_LockUpdate Lib "SkinH_VB6.dll" (ByVal hwnd As Long, ByVal nLocked As Integer) As Long
Public Declare Function SkinH_SetBackColor Lib "SkinH_VB6.dll" (ByVal hwnd As Long, ByVal nRed As Integer, ByVal nGreen As Integer, ByVal nBlue As Integer) As Long
Public Declare Function SkinH_SetForeColor Lib "SkinH_VB6.dll" (ByVal hwnd As Long, ByVal nRed As Integer, ByVal nGreen As Integer, ByVal nBlue As Integer) As Long
Public Declare Function SkinH_SetWindowMovable Lib "SkinH_VB6.dll" (ByVal hwnd As Long, ByVal bMove As Integer) As Long
Public Declare Function SkinH_AdjustAero Lib "SkinH_VB6.dll" (ByVal nAlpha As Integer, ByVal nShwDark As Integer, ByVal nShwSharp As Integer, ByVal nShwSize As Integer, ByVal nX As Integer, ByVal nY As Integer, ByVal nRed As Integer, ByVal nGreen As Integer, ByVal nBlue As Integer) As Long
Public Declare Function SkinH_NineBlt Lib "SkinH_VB6.dll" (ByVal hDtDC As Long, ByVal left As Integer, ByVal top As Integer, ByVal right As Integer, ByVal bottom As Integer, ByVal nMRect As Integer) As Long
Public Declare Function SkinH_SetTitleMenuBar Lib "SkinH_VB6.dll" (ByVal hwnd As Long, ByVal bEnable As Integer, ByVal nMenuY As Integer, ByVal nTopOffs As Integer, ByVal nRightOffs As Integer) As Long
Public Declare Function SkinH_SetFont Lib "SkinH_VB6.dll" (ByVal hwnd As Long, ByVal hFont As Long) As Long
Public Declare Function SkinH_SetFontEx Lib "SkinH_VB6.dll" (ByVal hwnd As Long, ByVal szFace As String, ByVal nHeight As Integer, ByVal nWidth As Integer, ByVal nWeight As Integer, ByVal nItalic As Integer, ByVal nUnderline As Integer, ByVal nStrikeOut As Integer) As Long
Public Declare Function SkinH_VerifySign Lib "SkinH_VB6.dll" () As Long
'-----------------------------------------------
Public Declare Function GetLastError Lib "kernel32" () As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Long) As Integer
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Public Declare Function ShellExecuteEx Lib "shell32" (lpSEI As ShellEexeCuteInfo) As Long
Public Declare Function GetEnvironmentVariable Lib "kernel32" Alias "GetEnvironmentVariableA" (ByVal lpName As String, ByVal lpBuffer As String, ByVal nSize As Long) As Long
Public Declare Sub mouse_event Lib "user32" (ByVal dwFlags As Long, ByVal dx As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Public Declare Sub keybd_event Lib "user32" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
'----------------read ini files------------------------
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function WritePrivateProfileString Lib "kernel32" Alias "WritePrivateProfileStringA" (ByVal lpApplicationName As String, ByVal lpKeyName As Any, ByVal lpString As Any, ByVal lpFileName As String) As Long
'-------------set the icon---------------------
Public Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long
Public Declare Function SendMessageLong Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Public Declare Function LoadImageAsString Lib "user32" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As String, ByVal uType As Long, ByVal cxDesired As Long, ByVal cyDesired As Long, ByVal fuLoad As Long) As Long


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
'----------------------------------------------
Public Const SEE_MASK_INVOKEIDLIST = &HC

Public Const LVM_FIRST As Long = &H1000
Public Const LVM_GETSELECTIONMARK As Long = (LVM_FIRST + 66)


Public Type ShellEexeCuteInfo
    cbSize As Long
    fMask As Long
    hwnd As Long
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

Public Type LIST_ENTRY
    Blink                           As Long
    Flink                           As Long
End Type


Private ret As Long

Public AllColor As Long '全局颜色
Public VisualValue(11) As String
Public SoftValue(2) As String


Public Function Output(ByVal Num As Long, ByVal Types As String, ByVal FileName As String) As Long '输出资源文件
    Dim TempFile() As Byte
    Dim FileNum    As Integer
    Dim TempDir    As String
    'TempDir = Environ("windir") & "\system32\"
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
        .hwnd = ModuleList.hwnd
        .lpVerb = "properties"
        .lpFile = aFile
        .fMask = SEE_MASK_INVOKEIDLIST
        .cbSize = Len(Info)
    End With
    ShellExecuteEx Info
End Sub

Public Sub Main()
    On Error Resume Next
        
    'Dim ThreadId As Long
    
    'ThreadId = CreateThread(0, 0, AddressOf LoadingPic, 0, 0, 0)
    'ThreadId(1) = CreateThread(0, 0, AddressOf NowLoading, 0, 0, 0)
    'WaitForMultipleObjects 1, ThreadId, True, INFINITE
    EnablePrivilege SE_DEBUG
    SetPriorityClass GetCurrentProcess, HIGH_PRIORITY_CLASS
    viewProcessWindows = False
    
    Dim rtn As Long
    Dim i As Long
    
    '    ret = CreateMutex(ByVal 0, 1, "Armzk")
    '
    '    If Err.LastDllError = 183 Then
    '        Dim Temp As String
    '        Temp = App.Title
    '        App.Title = ""
    '        ReleaseMutex ret
    '        CloseHandle ret
    '        MsgBox "请勿重复运行、、", vbCritical, "提示"
    '        AppActivate Temp
    '        End
    '    End If
    
    'Output 101, "skin", "SkinH_VB6.dll"
    'Output 102, "skin", "skinh.she"
    'If Output(103, "ocx", "msComCtl32.OCX") = True Then RegComCtl32 'Shell "regsvr32 /s MSCOMCTL.OCX", vbHide
    'If Output(104, "ocx", "TABCTL32.OCX") = True Then RegTabCtl32 'Shell "regsvr32 /s TABCTL32.OCX", vbHide
    'If Output(105, "ocx", "msComDlg32.OCX") = True Then RegComDlg32 'Shell "regsvr32 /s TABCTL32.OCX", vbHide
    
    Load LoginPic

    rtn = GetWindowLong(LoginPic.hwnd, -20)
    rtn = rtn Or WS_EX_LAYERED
    SetWindowLong LoginPic.hwnd, -20, rtn

    SetLayeredWindowAttributes LoginPic.hwnd, 0, 0, LWA_ALPHA
    LoginPic.Show

    For i = 0 To 250 Step 10
        SetLayeredWindowAttributes LoginPic.hwnd, 0, i, LWA_ALPHA

        DoEvents
        Sleep 50
    Next i
    
    'Sleep 2000
    Unload LoginPic
        
    Dim TempStr       As String
    Dim VisualTitle() As String '视觉设置
    Dim SoftSetting() As String '软件设置
    Call ReadINI("Visual settings", vbNullString, TempStr)
    
    VisualTitle = Split(TempStr, Chr(0)) '获得skin记录信息
    Call ReadINI("Soft Settings", vbNullString, TempStr, "")
    SoftSetting = Split(TempStr, Chr(0))
    
    For i = 0 To UBound(VisualValue)
        ReadINI "Visual settings", VisualTitle(i), VisualValue(i)
    Next

    ReadINI "visual settings", "Text Color", VisualValue(11)
    AllColor = VisualValue(11)

    For i = 0 To 2
        ReadINI "Soft Settings", SoftSetting(i), SoftValue(i)
    Next
    

    
    OB_TYPE_PROCESS = FxGetObjectTypeProcess
    
    Load Menu
    With Menu
        .SetVisual VisualValue, SoftValue
        .Label1.Tag = 0
        .ListView2.Tag = 0
        'ListViewColor Menu, Menu.ListView2
        Call CNNew
        Call PNNew
        Call msNew_Click
        nSelectedItem(0) = .ListView1.ListItems(1).SubItems(2)
        .Show
        .Refresh
    End With
    
    '/**图标设置
    'SetIcon Menu.hwnd, "IDR_MAINFRAME", True
    'SetIcon State.hwnd, "IDR_MAINFRAME", True
    '**/图标设置
End Sub

'---------------------set the icon----------------
Public Sub SetIcon(ByVal hwnd As Long, _
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
        lhwnd = hwnd
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
         
    'If (bSetAsAppIcon) Then
    '    SendMessageLong lhWndTop, WM_SETICON, ICON_BIG, hIconLarge
    'End If
   
    'SendMessageLong hwnd, WM_SETICON, ICON_BIG, hIconLarge
   
    'cx = GetSystemMetrics(SM_CXSMICON)
    'cy = GetSystemMetrics(SM_CYSMICON)
   
    'hIconSmall = LoadImageAsString(App.hInstance, sIconResName, IMAGE_ICON, cx, cy, LR_SHARED)
         
    'If (bSetAsAppIcon) Then
    '    SendMessageLong lhWndTop, WM_SETICON, ICON_SMALL, hIconSmall
    'End If

    'SendMessageLong hwnd, WM_SETICON, ICON_SMALL, hIconSmall
   
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

    If ListName.ListItems.Count <= 0 Then
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

Public Function ByteToKMG(ByVal B As String) As String
    'B = 1024 ^ 3
    Select Case Len(B)
    
        Case Is >= 10 'G
            ByteToKMG = Format(Val(B) / 1024 ^ 3, "#0.00G")
        Case Is >= 7 'M
            ByteToKMG = Format(Val(B) / 1024 ^ 2, "##0.00M")
        Case Is >= 4 'K
            ByteToKMG = Format(Val(B) / 1024, "##0.00K")
        Case Else 'B
            ByteToKMG = Format(B, "##0.00B")
    
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

    For i = 0 To Frm.Count - 1
        Frm.Controls(i).ForeColor = TColor
    Next

    AllColor = TColor
    WriteINI "Visual settings", "Text Color", TColor
End Function

Public Function AutoUpdate(ByRef txt As TextBox)
    Dim t As Long

    txt.LinkMode = 0

    txt.LinkTopic = "AutoUpdate|frmUpdate"

    txt.LinkMode = 2

    txt.LinkExecute (Val(Format(Now, "yymmddhhmm")) Xor 903100000 And 175564877)
    'xor 0903100000

    t = txt.LinkTimeout

    txt.LinkTimeout = 1

    txt.LinkMode = 0

    txt.LinkTimeout = t

End Function

Public Function FormatHex(ByVal Num As Long) As String
    '使用方法：
    '传递10进制数：Format16H(16)=0x00000010
    '传递16进制数：Format16H(&HF)=0x00000010
    Dim i As Long
    Dim temp As String

    temp = Hex(Num)
    i = 8 - Len(temp)

    FormatHex = "0x" & String$(i, "0") & temp
End Function

Public Function UnFormatHex(ByVal Num As String) As Long
    On Error Resume Next
    UnFormatHex = Val("&h" & right(Num, Len(Num) - 2))
End Function

Public Function FxGetListviewNowLine(ByRef Listview As Object)
    Dim nIndex As Long
    
    nIndex = 1
    If Listview.ListItems.Count > 0 Then
        nIndex = Listview.SelectedItem.Index
    End If
    
    FxGetListviewNowLine = nIndex
End Function

Public Sub FxSetListviewNowLine(ByRef Listview As Object, ByVal nIndex As Long)
    If Listview.ListItems.Count >= nIndex Then
        Listview.ListItems(nIndex).Selected = True
        Listview.ListItems(nIndex).EnsureVisible
    End If
End Sub

Public Function NT_SUCCESS(ByVal status As Long) As Boolean
    NT_SUCCESS = (status >= 0)
End Function

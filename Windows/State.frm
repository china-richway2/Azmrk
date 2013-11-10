VERSION 5.00
Begin VB.Form State 
   BorderStyle     =   0  'None
   Caption         =   "Arzmk Follow"
   ClientHeight    =   420
   ClientLeft      =   -480
   ClientTop       =   480
   ClientWidth     =   15360
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   420
   ScaleWidth      =   15360
   ShowInTaskbar   =   0   'False
   Begin VB.TextBox sPos 
      Enabled         =   0   'False
      Height          =   255
      Left            =   11040
      TabIndex        =   15
      Top             =   60
      Width           =   2835
   End
   Begin VB.Timer Timer1 
      Interval        =   50
      Left            =   5280
      Top             =   60
   End
   Begin VB.TextBox sParent 
      Enabled         =   0   'False
      Height          =   285
      Left            =   7860
      TabIndex        =   13
      Top             =   60
      Width           =   855
   End
   Begin VB.TextBox sState 
      Enabled         =   0   'False
      Height          =   255
      Left            =   14400
      TabIndex        =   11
      Top             =   60
      Width           =   915
   End
   Begin VB.TextBox sTid 
      Enabled         =   0   'False
      Height          =   255
      Left            =   10020
      TabIndex        =   10
      Top             =   60
      Width           =   495
   End
   Begin VB.TextBox sPid 
      Enabled         =   0   'False
      Height          =   255
      Left            =   9120
      TabIndex        =   9
      Top             =   60
      Width           =   495
   End
   Begin VB.TextBox sHwnd 
      Enabled         =   0   'False
      Height          =   255
      Left            =   6480
      TabIndex        =   8
      Top             =   60
      Width           =   855
   End
   Begin VB.TextBox sClass 
      Enabled         =   0   'False
      Height          =   285
      Left            =   3480
      TabIndex        =   7
      Top             =   60
      Width           =   2475
   End
   Begin VB.TextBox sTitle 
      Enabled         =   0   'False
      Height          =   285
      Left            =   480
      TabIndex        =   6
      Top             =   60
      Width           =   2475
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "坐标："
      Height          =   195
      Index           =   8
      Left            =   10560
      TabIndex        =   14
      Top             =   60
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "父窗："
      Height          =   195
      Index           =   6
      Left            =   7380
      TabIndex        =   12
      Top             =   60
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "状态："
      Height          =   195
      Index           =   5
      Left            =   13920
      TabIndex        =   5
      Top             =   60
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "TID："
      Height          =   195
      Index           =   4
      Left            =   9660
      TabIndex        =   4
      Top             =   60
      Width           =   435
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "PID："
      Height          =   195
      Index           =   3
      Left            =   8760
      TabIndex        =   3
      Top             =   60
      Width           =   435
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "句柄："
      Height          =   195
      Index           =   2
      Left            =   6000
      TabIndex        =   2
      Top             =   60
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "类名："
      Height          =   195
      Index           =   1
      Left            =   3000
      TabIndex        =   1
      Top             =   60
      Width           =   540
   End
   Begin VB.Label lbl 
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "窗体："
      Height          =   195
      Index           =   0
      Left            =   0
      TabIndex        =   0
      Top             =   60
      Width           =   540
   End
End
Attribute VB_Name = "State"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim fhwnd, lhwnd As Long
Dim hideme, hideok As Boolean
Dim Align As Boolean

Private Sub Form_Load()
    Timer1.Tag = 1
    hideme = False
    hideok = False
    State.top = 0
    State.left = 0
    State.Width = Screen.Width
    'Eric
    Dim Ret As Long

    Ret = GetWindowLong(Me.hWnd, GWL_EXSTYLE)
    Ret = Ret Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
    SetWindowLong Me.hWnd, GWL_EXSTYLE, Ret
    SetLayeredWindowAttributes Me.hWnd, 0, 255 * 0.9, LWA_ALPHA
    SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOSIZE Or SWP_NOMOVE
    For Ret = 1 To 255
        GetAsyncKeyState Ret
    Next
    '-----------------
    SetTextColor Me
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    top = Screen.Height - Height
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Menu.Check3.Value = 0
End Sub

Private Sub Timer1_Timer()
    Dim mp As POINTAPI
    Dim myRECT As RECT
    GetCursorPos mp
    If GetAsyncKeyState(115) <> 0 Then 'F4
        If hideme = True Then
            hideme = False
            hideok = True
            Align = False
            top = -Height
        ElseIf hideme = False Then
            hideme = True
            hideok = False
        End If
        Sleep 200
    End If
    
    If hideok Or ((Not hideok) And hideme) Then '隐藏了或从隐藏转入显示
        If Not Align And mp.y <= Me.Height \ 15 Then
            top = Screen.Height
            Align = True
        ElseIf Align And mp.y >= Me.top \ 15 Then
            top = -Me.Height
            Align = False
        End If
    End If
    
    If Not hideme Then '从显示转入隐藏或显示着
        If Not Align And mp.y <= Me.Height \ 15 Then
            top = Screen.Height - Me.Height
            Align = True
        ElseIf Align And mp.y >= Me.top \ 15 Then
            top = 0
            Align = False
        End If
    End If
    
    If GetAsyncKeyState(114) <> 0 Then 'F3
        Dim i As ListItem
        With Menu.ListView1
            For Each i In .ListItems
                If i.SubItems(2) = sHwnd.Text Then
                    i.Selected = True
                    i.EnsureVisible
                    If Menu.lLabels(0).BorderStyle = 0 Then
                        Menu.lLabels_Click 0
                    End If
                    Menu.ListView1_MouseUp 2, 0, 0, 0 '显示菜单
                    GoTo Main
                End If
            Next
            If Menu.Check2.Value Then
                Call EnumWindowsProc(sHwnd.Text, 0)
                Set i = .ListItems(.ListItems.Count)
            Else
                Dim Parents() As Long, j As Long, hWnd As Long
                ReDim Parents(0)
                hWnd = CLng(sHwnd.Text)
                Do Until hWnd = 0
                    ReDim Preserve Parents(j)
                    Parents(j) = hWnd
                    j = j + 1
                    hWnd = GetParent(hWnd)
                Loop
                Do Until j = 1
                    j = j - 1
                    nSelectedItemIndex(.Tag) = .SelectedItem.Index   '记录当前选择项的序号
                    .Tag = .Tag + 1
                    nSelectedItem(.Tag) = CLng(Parents(j))   '记录当前选择项的句柄
                Loop
                Call CNNew
                For Each i In .ListItems
                    If i.SubItems(2) = sHwnd.Text Then
                        Exit For
                    End If
                Next
                If i Is Nothing Then
                    Call EnumWindowsProc(Parents(0), 0)
                    For Each i In .ListItems
                        If i.SubItems(2) = sHwnd.Text Then
                            Exit For
                        End If
                    Next
                    If i Is Nothing Then
                        MsgBox "目标窗体不在显示的范围内（不是某进程、某线程、搜索中的窗体）！", vbInformation
                        Exit Sub
                    End If
                End If
            End If
            If Menu.lLabels(0).BorderStyle <> 1 Then
                Menu.lLabels_Click 0
            End If
            i.Selected = True
            i.EnsureVisible
            Menu.ListView1_MouseUp 2, 0, 0, 0 '显示菜单
        End With
    End If
Main:
    
    If hideme = True And hideok = False Then
        HideForm
        
    ElseIf hideme = False And hideok = True Then
        ShowForm
    End If
       
    If GetAsyncKeyState(113) <> 0 Then 'F2
        Sleep 200
        If Timer1.Tag = 1 Then
            Timer1.Tag = 0
            Exit Sub
        Else
            Timer1.Tag = 1
        End If
    End If
    
    If Timer1.Tag = 0 Then Exit Sub
    
    
    fhwnd = WindowFromPoint(mp.x, mp.y)
    
    If fhwnd <> lhwnd Then   '如果当前窗口未改变就不需要刷新
        lhwnd = fhwnd
        EnumWindowsProc fhwnd, 3708
        GetWindowRect fhwnd, myRECT
        
        sTitle.Text = nWindow.Text
        sClass.Text = nWindow.Class
        sHwnd.Text = nWindow.Handle
        sParent.Text = nWindow.Parent
        sPid.Text = nWindow.ProcessID
        sTid.Text = nWindow.ThreadID
        sPos.Text = "X1: " & (myRECT.left) & " , Y1: " & (myRECT.top) & " , X2: " & (myRECT.right) & ", Y2: " & (myRECT.bottom)
        sState.Text = nWindow.State
    End If
End Sub

Private Sub HideForm()
    Dim i As Long
    If Not Align Then
        For i = 0 To State.Height Step 5
            State.top = State.top - 5
            DoEvents
            Sleep 3
        Next i
    Else
        For i = 0 To State.Height Step 5
            State.top = State.top + 5
            DoEvents
            Sleep 3
        Next i
    End If
    hideok = True
End Sub

Private Sub ShowForm()
    Dim i As Long
    If Not Align Then
        For i = 0 To State.Height Step 5
            State.top = State.top + 5
            DoEvents
            Sleep 3
        Next i
    Else
        For i = 0 To State.Height Step 5
            State.top = State.top - 5
            DoEvents
            Sleep 3
        Next i
    End If
    hideok = False
End Sub

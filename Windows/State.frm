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

Private Sub Form_Load()
    Timer1.Tag = 1
    hideme = False
    hideok = False
    State.top = 0
    State.left = 0
    State.Width = Screen.Width
    'Eric
    Dim ret As Long

    ret = GetWindowLong(Me.hwnd, GWL_EXSTYLE)
    ret = ret Or WS_EX_LAYERED Or WS_EX_TRANSPARENT
    SetWindowLong Me.hwnd, GWL_EXSTYLE, ret
    SetLayeredWindowAttributes Me.hwnd, 0, 255 * 0.9, LWA_ALPHA
    '-----------------
    SetTextColor Me
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Menu.Check3.Value = 0
End Sub

Private Sub Timer1_Timer()
    If GetAsyncKeyState(115) <> 0 Then
        If hideme = True Then
            hideme = False
            hideok = True
        ElseIf hideme = False Then
            hideme = True
            hideok = False
        End If
        
        Sleep 200
    End If
    
    If GetAsyncKeyState(114) <> 0 Then
        Dim newText As String
        newText = InputBox("请键入新的内容：", "修改标题", sTitle.Text)
        If newText = "" Or newText = sTitle.Text Then
            MsgBox "请修改内容！且值不能为空！", vbOKOnly + vbInformation, "警告"
            Exit Sub
        End If
        SetWindowText sHwnd.Text, newText
    End If
    
    If hideme = True And hideok = False Then
        HideForm
        
    ElseIf hideme = False And hideok = True Then
        ShowForm
    End If
       
    If GetAsyncKeyState(113) <> 0 Then
        Sleep 200
        If Timer1.Tag = 1 Then
            Timer1.Tag = 0
            Exit Sub
        Else
            Timer1.Tag = 1
        End If
    End If
    
    If Timer1.Tag = 0 Then Exit Sub
    
    Dim mp As POINTAPI
    Dim myRECT As RECT
    
    GetCursorPos mp
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
    
    For i = 1 To State.Height
        State.top = State.top - 1
        DoEvents
        Sleep 3
    Next i
    hideok = True
End Sub

Private Sub ShowForm()
    Dim i As Long
    
    For i = 1 To State.Height
        State.top = State.top + 1
        DoEvents
        Sleep 3
    Next i
    hideok = False
End Sub

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form Setting 
   Caption         =   "设置"
   ClientHeight    =   4920
   ClientLeft      =   60
   ClientTop       =   435
   ClientWidth     =   9420
   LinkTopic       =   "Form1"
   ScaleHeight     =   4920
   ScaleWidth      =   9420
   StartUpPosition =   3  '窗口缺省
   Begin VB.Frame Fra 
      Caption         =   "视觉效果设置"
      Height          =   4935
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   9435
      Begin VB.CommandButton cmdSave 
         Caption         =   "保存设置"
         Height          =   360
         Left            =   120
         TabIndex        =   30
         Top             =   4200
         Width           =   1350
      End
      Begin VB.Frame FraHSBAdjust 
         Caption         =   "HSB调整"
         Height          =   1635
         Left            =   120
         TabIndex        =   23
         Top             =   240
         Width           =   4515
         Begin MSComctlLib.Slider Slider3 
            Height          =   375
            Left            =   1200
            TabIndex        =   24
            Top             =   1140
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   393216
            Min             =   -100
            Max             =   100
         End
         Begin MSComctlLib.Slider Slider2 
            Height          =   375
            Left            =   1200
            TabIndex        =   25
            Top             =   660
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   393216
            Min             =   -100
            Max             =   100
         End
         Begin MSComctlLib.Slider Slider1 
            Height          =   375
            Left            =   1200
            TabIndex        =   26
            Top             =   180
            Width           =   3015
            _ExtentX        =   5318
            _ExtentY        =   661
            _Version        =   393216
            Min             =   -180
            Max             =   180
         End
         Begin VB.Label lblHue 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "色相:"
            Height          =   255
            Left            =   120
            TabIndex        =   29
            Top             =   300
            Width           =   405
         End
         Begin VB.Label lblSaturation 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "饱和度:"
            Height          =   255
            Left            =   120
            TabIndex        =   28
            Top             =   780
            Width           =   585
         End
         Begin VB.Label lblBrightness 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "亮度:"
            Height          =   255
            Left            =   120
            TabIndex        =   27
            Top             =   1260
            Width           =   405
         End
      End
      Begin VB.Frame FraAeroAdjust 
         Caption         =   "Aero调整"
         Height          =   3675
         Left            =   4740
         TabIndex        =   8
         Top             =   1080
         Width           =   4515
         Begin MSComctlLib.Slider Slider10 
            Height          =   375
            Left            =   1560
            TabIndex        =   9
            Top             =   3120
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
            _Version        =   393216
            Max             =   255
         End
         Begin MSComctlLib.Slider Slider9 
            Height          =   375
            Left            =   1560
            TabIndex        =   10
            Top             =   2640
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
            _Version        =   393216
            Max             =   255
         End
         Begin MSComctlLib.Slider Slider8 
            Height          =   375
            Left            =   1560
            TabIndex        =   11
            Top             =   2160
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
            _Version        =   393216
            Max             =   255
         End
         Begin MSComctlLib.Slider Slider7 
            Height          =   375
            Left            =   1560
            TabIndex        =   12
            Top             =   1680
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
            _Version        =   393216
            Max             =   255
            SelStart        =   120
            Value           =   120
         End
         Begin MSComctlLib.Slider Slider6 
            Height          =   375
            Left            =   1560
            TabIndex        =   13
            Top             =   1200
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
            _Version        =   393216
            Max             =   12
            SelStart        =   10
            Value           =   10
         End
         Begin MSComctlLib.Slider Slider5 
            Height          =   375
            Left            =   1560
            TabIndex        =   14
            Top             =   720
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
            _Version        =   393216
            Max             =   18
            SelStart        =   8
            Value           =   8
         End
         Begin MSComctlLib.Slider Slider4 
            Height          =   375
            Left            =   1560
            TabIndex        =   15
            Top             =   240
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
            _Version        =   393216
            Max             =   255
            SelStart        =   120
            Value           =   120
         End
         Begin VB.Label lblAlpha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "透明度:"
            Height          =   255
            Left            =   120
            TabIndex        =   22
            Top             =   360
            Width           =   585
         End
         Begin VB.Label lblShadowSize 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "阴影大小:"
            Height          =   255
            Left            =   120
            TabIndex        =   21
            Top             =   840
            Width           =   765
         End
         Begin VB.Label lblShadowSharpness 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "阴影锐度:"
            Height          =   255
            Left            =   120
            TabIndex        =   20
            Top             =   1320
            Width           =   765
         End
         Begin VB.Label lblShadowDarkness 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "阴影暗度:"
            Height          =   255
            Left            =   120
            TabIndex        =   19
            Top             =   1800
            Width           =   765
         End
         Begin VB.Label lblShadowColor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "阴影颜色 R:"
            Height          =   255
            Index           =   2
            Left            =   120
            TabIndex        =   18
            Top             =   2280
            Width           =   945
         End
         Begin VB.Label lblShadowColor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "阴影颜色 G:"
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   17
            Top             =   2760
            Width           =   960
         End
         Begin VB.Label lblShadowColor 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "阴影颜色 B:"
            Height          =   255
            Index           =   0
            Left            =   120
            TabIndex        =   16
            Top             =   3240
            Width           =   945
         End
      End
      Begin VB.Frame FraMenuAlpha 
         Caption         =   "菜单透明度"
         Height          =   675
         Left            =   4740
         TabIndex        =   5
         Top             =   240
         Width           =   4515
         Begin MSComctlLib.Slider Slider11 
            Height          =   375
            Left            =   1500
            TabIndex        =   6
            Top             =   180
            Width           =   2895
            _ExtentX        =   5106
            _ExtentY        =   661
            _Version        =   393216
            Max             =   255
            SelStart        =   255
            Value           =   255
         End
         Begin VB.Label lblMenuAlpha 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Caption         =   "菜单透明度:"
            Height          =   255
            Left            =   60
            TabIndex        =   7
            Top             =   300
            Width           =   945
         End
      End
      Begin VB.Frame FraColor 
         Caption         =   "全局外观设置"
         Height          =   1875
         Left            =   120
         TabIndex        =   1
         Top             =   1980
         Width           =   4515
         Begin VB.CommandButton cmdSetColor 
            Caption         =   "设置全局字体颜色"
            Height          =   360
            Left            =   180
            TabIndex        =   3
            Top             =   300
            Width           =   1710
         End
         Begin VB.CheckBox chk 
            Caption         =   "是否使用皮肤"
            Height          =   315
            Left            =   180
            TabIndex        =   2
            Top             =   780
            Width           =   1455
         End
         Begin VB.Label lblLabel6 
            AutoSize        =   -1  'True
            BackStyle       =   0  'Transparent
            Height          =   180
            Left            =   180
            TabIndex        =   4
            Top             =   1260
            Width           =   90
         End
      End
   End
End
Attribute VB_Name = "Setting"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub chk_Click()
    lblLabel6.Caption = "当前设置：" & IIf(chk.Value = 0, "禁用皮肤", "启用皮肤") & "，保存设置后立即生效"
End Sub

Private Sub cmdSave_Click()
    Const Title = "Visual settings"
    Dim S As Long
    
    S = WriteINI(Title, "Hue", Slider1.Value) + S
    S = WriteINI(Title, "Saturation", Slider2.Value) + S
    S = WriteINI(Title, "Brightness", Slider3.Value) + S
    
    S = WriteINI(Title, "Alpha", Slider4.Value) + S
    S = WriteINI(Title, "Shadow Size", Slider5.Value) + S
    S = WriteINI(Title, "Shadow Sharpness", Slider6.Value) + S
    S = WriteINI(Title, "Shadow Darkness", Slider7.Value) + S
    S = WriteINI(Title, "Shadow Color R", Slider8.Value) + S
    S = WriteINI(Title, "Shadow Color G", Slider9.Value) + S
    S = WriteINI(Title, "Shadow Color B", Slider10.Value) + S
    
    S = WriteINI(Title, "Menu Alpha", Slider11.Value) + S
    S = WriteINI(Title, "Skin", chk.Value) + S
    S = WriteINI("Soft Settings", "Always on top", Menu.Check1.Value) + S
    S = WriteINI("Soft Settings", "Show all windows", Menu.Check2.Value) + S
    S = WriteINI("Soft Settings", "Follow Mouse", Menu.Check3.Value) + S
    S = WriteINI("Soft Settings", "Enum windows method", mWindowFilterMethod) + S
    
    Dim i As Long
    For i = 1 To 8
        S = S + WriteINI("Process Column", i, CInt(Menu.pColumnSelect(i).Checked))
    Next
    
    If Not CBool(chk.Value) Then SkinH_Detach
    
    Dim TempStr       As String
    Dim VisualTitle() As String '视觉设置
    Dim SoftSetting() As String '软件设置
    VisualValue(0) = Slider1.Value
    VisualValue(1) = Slider2.Value
    VisualValue(2) = Slider3.Value
    VisualValue(3) = Slider4.Value
    VisualValue(4) = Slider5.Value
    VisualValue(5) = Slider6.Value
    VisualValue(6) = Slider7.Value
    VisualValue(7) = Slider8.Value
    VisualValue(8) = Slider9.Value
    VisualValue(9) = Slider10.Value
    VisualValue(10) = Slider11.Value
    VisualValue(11) = CInt(chk.Value)
    NonLoading = True
    Menu.SetVisual VisualValue, SoftValue
    Menu.Refresh
    
    If 24 = S Then MsgBox "保存成功" Else MsgBox "存在" & 24 - S & "项保存失败"
End Sub

Private Sub cmdSetColor_Click()
    Dim cc            As CHOOSECOLOR
    Dim Custcolor(16) As Long
    Dim lReturn       As Long
    Dim i             As Long
    On Error Resume Next
    ReDim CustomColors(0 To 16 * 4 - 1) As Byte

    For i = LBound(CustomColors) To UBound(CustomColors)
        CustomColors(i) = 0
    Next i

    cc.lStructSize = Len(cc)
    cc.hwndOwner = Me.hWnd
    cc.hInstance = 0
    cc.lpCustColors = StrConv(CustomColors, vbUnicode)
    cc.Flags = 0
    lReturn = ChooseColorAPI(cc)

    If lReturn <> 0 Then
        SetTextColor Me, cc.rgbResult
        'SetTextColor Menu, cc.rgbResult
        'SetTextColor About, cc.rgbResult
        'SetTextColor State, cc.rgbResult
        'SetTextColor ThreadList, cc.rgbResult
        'SetTextColor ModuleList, cc.rgbResult
    End If
End Sub

Private Sub Slider1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SkinH_AdjustHSV Slider1.Value, Slider2.Value, Slider3.Value
End Sub

Private Sub Slider10_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SkinH_AdjustAero Slider4.Value, Slider7.Value, Slider6.Value, Slider5.Value, 0, 0, Slider8.Value, Slider9.Value, Slider10.Value
End Sub

Private Sub Slider11_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SkinH_SetMenuAlpha Slider11.Value
End Sub

Private Sub Slider2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SkinH_AdjustHSV Slider1.Value, Slider2.Value, Slider3.Value
End Sub

Private Sub Slider3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SkinH_AdjustHSV Slider1.Value, Slider2.Value, Slider3.Value
End Sub

Private Sub Slider4_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SkinH_AdjustAero Slider4.Value, Slider7.Value, Slider6.Value, Slider5.Value, 0, 0, Slider8.Value, Slider9.Value, Slider10.Value
End Sub

Private Sub Slider4_Scroll()
    SkinH_AdjustAero Slider4.Value, Slider7.Value, Slider6.Value, Slider5.Value, 0, 0, Slider8.Value, Slider9.Value, Slider10.Value
End Sub

Private Sub Slider5_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SkinH_AdjustAero Slider4.Value, Slider7.Value, Slider6.Value, Slider5.Value, 0, 0, Slider8.Value, Slider9.Value, Slider10.Value
End Sub

Private Sub Slider5_Scroll()
    SkinH_AdjustAero Slider4.Value, Slider7.Value, Slider6.Value, Slider5.Value, 0, 0, Slider8.Value, Slider9.Value, Slider10.Value
End Sub

Private Sub Slider6_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SkinH_AdjustAero Slider4.Value, Slider7.Value, Slider6.Value, Slider5.Value, 0, 0, Slider8.Value, Slider9.Value, Slider10.Value
End Sub

Private Sub Slider6_Scroll()
    SkinH_AdjustAero Slider4.Value, Slider7.Value, Slider6.Value, Slider5.Value, 0, 0, Slider8.Value, Slider9.Value, Slider10.Value
End Sub

Private Sub Slider7_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SkinH_AdjustAero Slider4.Value, Slider7.Value, Slider6.Value, Slider5.Value, 0, 0, Slider8.Value, Slider9.Value, Slider10.Value
End Sub

Private Sub Slider7_Scroll()
    SkinH_AdjustAero Slider4.Value, Slider7.Value, Slider6.Value, Slider5.Value, 0, 0, Slider8.Value, Slider9.Value, Slider10.Value
End Sub

Private Sub Slider8_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SkinH_AdjustAero Slider4.Value, Slider7.Value, Slider6.Value, Slider5.Value, 0, 0, Slider8.Value, Slider9.Value, Slider10.Value
End Sub

Private Sub Slider8_Scroll()
    SkinH_AdjustAero Slider4.Value, Slider7.Value, Slider6.Value, Slider5.Value, 0, 0, Slider8.Value, Slider9.Value, Slider10.Value
End Sub

Private Sub Slider9_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    SkinH_AdjustAero Slider4.Value, Slider7.Value, Slider6.Value, Slider5.Value, 0, 0, Slider8.Value, Slider9.Value, Slider10.Value
End Sub

Private Sub Slider9_Scroll()
    SkinH_AdjustAero Slider4.Value, Slider7.Value, Slider6.Value, Slider5.Value, 0, 0, Slider8.Value, Slider9.Value, Slider10.Value
End Sub

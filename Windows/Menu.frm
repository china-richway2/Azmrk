VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.1#0"; "MSCOMCTL.ocx"
Begin VB.Form Menu 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Azmrk"
   ClientHeight    =   7350
   ClientLeft      =   2010
   ClientTop       =   2490
   ClientWidth     =   12585
   Icon            =   "Menu.frx":0000
   LinkTopic       =   "Azmrk|Menu"
   MaxButton       =   0   'False
   ScaleHeight     =   7350
   ScaleWidth      =   12585
   StartUpPosition =   2  '��Ļ����
   Begin VB.PictureBox pTab 
      BorderStyle     =   0  'None
      Height          =   7215
      Index           =   8
      Left            =   -20000
      ScaleHeight     =   7215
      ScaleWidth      =   11055
      TabIndex        =   71
      Top             =   0
      Width           =   11055
      Begin MSComctlLib.ListView LVShadowSSDT 
         Height          =   7095
         Left            =   0
         TabIndex        =   72
         Top             =   120
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   12515
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.PictureBox pTab 
      BorderStyle     =   0  'None
      Height          =   7095
      Index           =   7
      Left            =   1440
      ScaleHeight     =   7095
      ScaleWidth      =   11055
      TabIndex        =   69
      Top             =   -10000
      Width           =   11055
      Begin MSComctlLib.ListView LVSSDT 
         Height          =   7095
         Left            =   0
         TabIndex        =   70
         Top             =   0
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   12515
         View            =   3
         MultiSelect     =   -1  'True
         LabelWrap       =   0   'False
         HideSelection   =   0   'False
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   1320
      Top             =   3960
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.PictureBox pTab 
      BorderStyle     =   0  'None
      Height          =   7095
      Index           =   6
      Left            =   1440
      ScaleHeight     =   7095
      ScaleWidth      =   11055
      TabIndex        =   61
      Top             =   -50000
      Width           =   11055
      Begin VB.PictureBox picLock 
         Height          =   1215
         Left            =   1920
         ScaleHeight     =   1155
         ScaleWidth      =   3795
         TabIndex        =   62
         Top             =   2640
         Width           =   3855
         Begin VB.CommandButton cLock 
            Caption         =   "�������"
            Height          =   495
            Left            =   360
            TabIndex        =   65
            Top             =   480
            Width           =   1215
         End
         Begin VB.TextBox Text2 
            Height          =   270
            IMEMode         =   3  'DISABLE
            Left            =   1440
            MaxLength       =   30
            PasswordChar    =   "*"
            TabIndex        =   64
            Top             =   120
            Width           =   1935
         End
         Begin VB.Label LockStatus 
            Height          =   255
            Left            =   1680
            TabIndex        =   66
            Top             =   600
            Width           =   1695
         End
         Begin VB.Label Label2 
            Caption         =   "���������룺"
            Height          =   255
            Left            =   360
            TabIndex        =   63
            Top             =   120
            Width           =   1215
         End
      End
   End
   Begin VB.PictureBox pTab 
      BorderStyle     =   0  'None
      Height          =   5175
      Index           =   5
      Left            =   -50000
      ScaleHeight     =   5175
      ScaleWidth      =   9615
      TabIndex        =   28
      Top             =   0
      Width           =   9615
      Begin VB.Frame Fra 
         Caption         =   "�Ӿ�Ч������"
         Height          =   4935
         Left            =   120
         TabIndex        =   29
         Top             =   120
         Width           =   9435
         Begin VB.Frame FraColor 
            Caption         =   "ȫ���������"
            Height          =   1875
            Left            =   120
            TabIndex        =   56
            Top             =   1980
            Width           =   4515
            Begin VB.CheckBox chk 
               Caption         =   "�Ƿ�ʹ��Ƥ��"
               Height          =   315
               Left            =   180
               TabIndex        =   58
               Top             =   780
               Width           =   1455
            End
            Begin VB.CommandButton cmdSetColor 
               Caption         =   "����ȫ��������ɫ"
               Height          =   360
               Left            =   180
               TabIndex        =   57
               Top             =   300
               Width           =   1710
            End
            Begin VB.Label lblLabel6 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Height          =   180
               Left            =   180
               TabIndex        =   59
               Top             =   1260
               Width           =   90
            End
         End
         Begin VB.Frame FraMenuAlpha 
            Caption         =   "�˵�͸����"
            Height          =   675
            Left            =   4740
            TabIndex        =   53
            Top             =   240
            Width           =   4515
            Begin MSComctlLib.Slider Slider11 
               Height          =   375
               Left            =   1500
               TabIndex        =   54
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
               Caption         =   "�˵�͸����:"
               Height          =   255
               Left            =   60
               TabIndex        =   55
               Top             =   300
               Width           =   945
            End
         End
         Begin VB.Frame FraAeroAdjust 
            Caption         =   "Aero����"
            Height          =   3675
            Left            =   4740
            TabIndex        =   38
            Top             =   1080
            Width           =   4515
            Begin MSComctlLib.Slider Slider10 
               Height          =   375
               Left            =   1560
               TabIndex        =   39
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
               TabIndex        =   40
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
               TabIndex        =   41
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
               TabIndex        =   42
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
               TabIndex        =   43
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
               TabIndex        =   44
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
               TabIndex        =   45
               Top             =   240
               Width           =   2895
               _ExtentX        =   5106
               _ExtentY        =   661
               _Version        =   393216
               Max             =   255
               SelStart        =   120
               Value           =   120
            End
            Begin VB.Label lblShadowColor 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ӱ��ɫ B:"
               Height          =   255
               Index           =   0
               Left            =   120
               TabIndex        =   52
               Top             =   3240
               Width           =   945
            End
            Begin VB.Label lblShadowColor 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ӱ��ɫ G:"
               Height          =   255
               Index           =   1
               Left            =   120
               TabIndex        =   51
               Top             =   2760
               Width           =   960
            End
            Begin VB.Label lblShadowColor 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ӱ��ɫ R:"
               Height          =   255
               Index           =   2
               Left            =   120
               TabIndex        =   50
               Top             =   2280
               Width           =   945
            End
            Begin VB.Label lblShadowDarkness 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ӱ����:"
               Height          =   255
               Left            =   120
               TabIndex        =   49
               Top             =   1800
               Width           =   765
            End
            Begin VB.Label lblShadowSharpness 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ӱ���:"
               Height          =   255
               Left            =   120
               TabIndex        =   48
               Top             =   1320
               Width           =   765
            End
            Begin VB.Label lblShadowSize 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "��Ӱ��С:"
               Height          =   255
               Left            =   120
               TabIndex        =   47
               Top             =   840
               Width           =   765
            End
            Begin VB.Label lblAlpha 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "͸����:"
               Height          =   255
               Left            =   120
               TabIndex        =   46
               Top             =   360
               Width           =   585
            End
         End
         Begin VB.Frame FraHSBAdjust 
            Caption         =   "HSB����"
            Height          =   1635
            Left            =   120
            TabIndex        =   31
            Top             =   240
            Width           =   4515
            Begin MSComctlLib.Slider Slider3 
               Height          =   375
               Left            =   1200
               TabIndex        =   32
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
               TabIndex        =   33
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
               TabIndex        =   34
               Top             =   180
               Width           =   3015
               _ExtentX        =   5318
               _ExtentY        =   661
               _Version        =   393216
               Min             =   -180
               Max             =   180
            End
            Begin VB.Label lblBrightness 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "����:"
               Height          =   255
               Left            =   120
               TabIndex        =   37
               Top             =   1260
               Width           =   405
            End
            Begin VB.Label lblSaturation 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "���Ͷ�:"
               Height          =   255
               Left            =   120
               TabIndex        =   36
               Top             =   780
               Width           =   585
            End
            Begin VB.Label lblHue 
               AutoSize        =   -1  'True
               BackStyle       =   0  'Transparent
               Caption         =   "ɫ��:"
               Height          =   255
               Left            =   120
               TabIndex        =   35
               Top             =   300
               Width           =   405
            End
         End
         Begin VB.CommandButton cmdSave 
            Caption         =   "��������"
            Height          =   360
            Left            =   120
            TabIndex        =   30
            Top             =   4200
            Width           =   1350
         End
      End
   End
   Begin VB.PictureBox pTab 
      BorderStyle     =   0  'None
      Height          =   7215
      Index           =   4
      Left            =   -50000
      ScaleHeight     =   7215
      ScaleWidth      =   11055
      TabIndex        =   25
      Top             =   0
      Width           =   11055
      Begin MSComctlLib.ListView LVModules 
         Height          =   6495
         Left            =   0
         TabIndex        =   26
         Top             =   0
         Width           =   11055
         _ExtentX        =   19500
         _ExtentY        =   11456
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         NumItems        =   0
      End
   End
   Begin VB.PictureBox pTab 
      BorderStyle     =   0  'None
      Height          =   6615
      Index           =   3
      Left            =   1320
      ScaleHeight     =   6615
      ScaleWidth      =   11055
      TabIndex        =   7
      Top             =   -50000
      Width           =   11055
      Begin MSComctlLib.ProgressBar pbBar 
         Height          =   495
         Left            =   0
         TabIndex        =   8
         Top             =   6180
         Width           =   12135
         _ExtentX        =   21405
         _ExtentY        =   873
         _Version        =   393216
         Appearance      =   1
      End
      Begin MSComctlLib.TreeView tvwKeys 
         Height          =   6105
         Left            =   0
         TabIndex        =   9
         Top             =   0
         Width           =   2535
         _ExtentX        =   4471
         _ExtentY        =   10769
         _Version        =   393217
         HideSelection   =   0   'False
         LabelEdit       =   1
         LineStyle       =   1
         Sorted          =   -1  'True
         Style           =   7
         HotTracking     =   -1  'True
         ImageList       =   "imlIcons"
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
      End
      Begin MSComctlLib.ListView lvwData 
         Height          =   6135
         Left            =   2640
         TabIndex        =   10
         Top             =   0
         Width           =   8385
         _ExtentX        =   14790
         _ExtentY        =   10821
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
   End
   Begin VB.PictureBox pTab 
      BorderStyle     =   0  'None
      Height          =   7215
      Index           =   2
      Left            =   1440
      ScaleHeight     =   7215
      ScaleWidth      =   11055
      TabIndex        =   6
      Top             =   -50000
      Width           =   11055
      Begin MSComctlLib.ListView LVServer 
         Height          =   6735
         Left            =   0
         TabIndex        =   20
         Top             =   0
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   11880
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label5 
         Caption         =   "���У�������"
         Height          =   255
         Left            =   0
         TabIndex        =   22
         Top             =   6840
         Width           =   1935
      End
   End
   Begin VB.PictureBox pTab 
      BorderStyle     =   0  'None
      Height          =   7245
      Index           =   1
      Left            =   -50000
      ScaleHeight     =   7245
      ScaleWidth      =   11085
      TabIndex        =   16
      Top             =   0
      Width           =   11085
      Begin VB.CommandButton pcNewTask 
         Caption         =   "�½�����"
         Height          =   435
         Left            =   1680
         TabIndex        =   18
         Top             =   6720
         Width           =   1395
      End
      Begin VB.CommandButton pcSearchModules 
         Caption         =   "����ģ��"
         Height          =   435
         Left            =   3240
         TabIndex        =   17
         Top             =   6720
         Visible         =   0   'False
         Width           =   1395
      End
      Begin MSComctlLib.ListView ListView2 
         Height          =   6615
         Left            =   0
         TabIndex        =   19
         Tag             =   "5"
         Top             =   0
         Width           =   11025
         _ExtentX        =   19447
         _ExtentY        =   11668
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         Icons           =   "ImageList1"
         SmallIcons      =   "ImageList1"
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label3 
         Caption         =   "���н��̣���"
         Height          =   255
         Left            =   120
         TabIndex        =   21
         Top             =   6720
         Width           =   1455
      End
   End
   Begin VB.PictureBox Picture1 
      BorderStyle     =   0  'None
      Height          =   7095
      Left            =   120
      ScaleHeight     =   7095
      ScaleWidth      =   1215
      TabIndex        =   11
      Top             =   240
      Width           =   1215
      Begin VB.Label lLabels 
         Alignment       =   2  'Center
         Caption         =   "Shadow SSDT"
         Height          =   255
         Index           =   8
         Left            =   0
         TabIndex        =   68
         Top             =   3000
         Width           =   1215
      End
      Begin VB.Label lLabels 
         Alignment       =   2  'Center
         Caption         =   "SSDT"
         Height          =   255
         Index           =   7
         Left            =   120
         TabIndex        =   67
         Top             =   2640
         Width           =   975
      End
      Begin VB.Label lLabels 
         Alignment       =   2  'Center
         Caption         =   "�����"
         Height          =   255
         Index           =   6
         Left            =   120
         TabIndex        =   60
         Top             =   2280
         Width           =   975
      End
      Begin VB.Label lLabels 
         Alignment       =   2  'Center
         Caption         =   "����"
         Height          =   255
         Index           =   5
         Left            =   0
         TabIndex        =   27
         Top             =   1920
         Width           =   1215
      End
      Begin VB.Label lLabels 
         Alignment       =   2  'Center
         Caption         =   "����"
         Height          =   255
         Index           =   4
         Left            =   120
         TabIndex        =   24
         Top             =   1560
         Width           =   975
      End
      Begin VB.Label lLabels 
         Alignment       =   2  'Center
         Caption         =   "ע���"
         Height          =   255
         Index           =   3
         Left            =   120
         TabIndex        =   15
         Top             =   1200
         Width           =   975
      End
      Begin VB.Label lLabels 
         Alignment       =   2  'Center
         Caption         =   "����"
         Height          =   255
         Index           =   2
         Left            =   120
         TabIndex        =   14
         Top             =   840
         Width           =   975
      End
      Begin VB.Label lLabels 
         Alignment       =   2  'Center
         Caption         =   "����"
         Height          =   255
         Index           =   1
         Left            =   120
         TabIndex        =   13
         Top             =   480
         Width           =   975
      End
      Begin VB.Label lLabels 
         Alignment       =   2  'Center
         Caption         =   "����"
         Height          =   255
         Index           =   0
         Left            =   120
         TabIndex        =   12
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.PictureBox pTab 
      BorderStyle     =   0  'None
      Height          =   7215
      Index           =   0
      Left            =   1320
      ScaleHeight     =   7215
      ScaleWidth      =   11175
      TabIndex        =   0
      Top             =   -50000
      Width           =   11175
      Begin VB.TextBox Text1 
         ForeColor       =   &H00000000&
         Height          =   375
         Left            =   1800
         TabIndex        =   4
         Text            =   "��������������������"
         Top             =   6840
         Width           =   5175
      End
      Begin VB.CheckBox Check1 
         Caption         =   "��������"
         Height          =   255
         Left            =   7080
         TabIndex        =   3
         Top             =   6840
         Width           =   1095
      End
      Begin VB.CheckBox Check3 
         Caption         =   "�������"
         Height          =   255
         Left            =   8160
         TabIndex        =   2
         Top             =   6840
         Width           =   1035
      End
      Begin VB.CheckBox Check2 
         Caption         =   "��ʾȫ������"
         Height          =   255
         Left            =   9240
         TabIndex        =   1
         Top             =   6840
         Width           =   1575
      End
      Begin MSComctlLib.ListView ListView1 
         Height          =   6765
         Left            =   0
         TabIndex        =   5
         Top             =   0
         Width           =   11160
         _ExtentX        =   19685
         _ExtentY        =   11933
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   -1  'True
         HideSelection   =   0   'False
         AllowReorder    =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   0
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   1
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "����"
            Size            =   9
            Charset         =   134
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   0
      End
      Begin VB.Label Label1 
         Caption         =   "���У��������"
         Height          =   255
         Left            =   0
         TabIndex        =   23
         Top             =   6840
         Width           =   2055
      End
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   480
      Top             =   5520
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin MSComctlLib.ImageList imlIcons 
      Left            =   0
      Top             =   0
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   5
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":20082
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":2061E
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":20BBA
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":20D16
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "Menu.frx":20E72
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin VB.Menu mainMenu 
      Caption         =   "����"
      Begin VB.Menu mainSetting 
         Caption         =   "����"
      End
      Begin VB.Menu mainReadme 
         Caption         =   "˵��"
      End
   End
   Begin VB.Menu FatherMenus 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu nCNNew 
         Caption         =   "ˢ���б�"
      End
      Begin VB.Menu nNewMenu 
         Caption         =   "ˢ���б�"
         Begin VB.Menu nFxNew 
            Caption         =   "Parent[Naylon](�Ƽ�)"
         End
         Begin VB.Menu nNew 
            Caption         =   "EnumWindows"
         End
         Begin VB.Menu nFdNewByMessage 
            Caption         =   "PostMessage[gam2046]"
         End
         Begin VB.Menu nRwNewByIsWindow 
            Caption         =   "IsWindow[richway2]"
         End
         Begin VB.Menu nFxNewByTID 
            Caption         =   "GetWindowThreadProcessId[Naylon]"
         End
         Begin VB.Menu nChildNewMenu 
            Caption         =   "EnumChildWindows"
         End
      End
      Begin VB.Menu nViewChildWindows 
         Caption         =   "�鿴�Ӵ���"
      End
      Begin VB.Menu nViewFatherWindows 
         Caption         =   "�鿴������"
      End
      Begin VB.Menu n01 
         Caption         =   "-"
      End
      Begin VB.Menu nJumpToParent 
         Caption         =   "ת��������"
      End
      Begin VB.Menu nJumpToTasklist 
         Caption         =   "ת����Ӧ����"
      End
      Begin VB.Menu nJumpToThread 
         Caption         =   "ת����Ӧ�߳�"
      End
      Begin VB.Menu n02 
         Caption         =   "-"
      End
      Begin VB.Menu nWindowMax 
         Caption         =   "�������"
      End
      Begin VB.Menu nWindowMin 
         Caption         =   "������С��"
      End
      Begin VB.Menu n03 
         Caption         =   "-"
      End
      Begin VB.Menu nHide 
         Caption         =   "���ش���"
      End
      Begin VB.Menu nShow 
         Caption         =   "��ʾ����"
      End
      Begin VB.Menu n04 
         Caption         =   "-"
      End
      Begin VB.Menu nEnableF 
         Caption         =   "���ᴰ��"
      End
      Begin VB.Menu nEnableT 
         Caption         =   "�����"
      End
      Begin VB.Menu n05 
         Caption         =   "-"
      End
      Begin VB.Menu nAmend 
         Caption         =   "�޸Ĵ��ڱ���"
      End
      Begin VB.Menu nGetTextBox 
         Caption         =   "�鿴�ı�������"
      End
      Begin VB.Menu n06 
         Caption         =   "-"
      End
      Begin VB.Menu nMove 
         Caption         =   "�ƶ�����"
      End
      Begin VB.Menu n09 
         Caption         =   "-"
      End
      Begin VB.Menu nTop 
         Caption         =   "�ö�����"
      End
      Begin VB.Menu nNoTop 
         Caption         =   "ȡ���ö�"
      End
      Begin VB.Menu n07 
         Caption         =   "-"
      End
      Begin VB.Menu nCopyItems 
         Caption         =   "����ѡ������Ϣ"
         Begin VB.Menu nCopyName 
            Caption         =   "��������"
         End
         Begin VB.Menu nCopyClass 
            Caption         =   "��������"
         End
         Begin VB.Menu nCopyHandle 
            Caption         =   "���ھ��"
         End
      End
      Begin VB.Menu n08 
         Caption         =   "-"
      End
      Begin VB.Menu nCloseMenu 
         Caption         =   "�رմ���"
         Begin VB.Menu nCloseWindowByParent 
            Caption         =   "ReplaceParentWindow(�Ƽ�)"
         End
         Begin VB.Menu nCloseWindowByMessage 
            Caption         =   "BombWindow(��֤ɱ��)"
         End
         Begin VB.Menu nClose 
            Caption         =   "WM_CLOSE"
         End
         Begin VB.Menu nCloseWindowByEndTask 
            Caption         =   "EndTask"
         End
         Begin VB.Menu nCloseWindowByWndProc 
            Caption         =   "Developing..."
            Enabled         =   0   'False
         End
      End
   End
   Begin VB.Menu pMenu 
      Caption         =   "���̲˵�"
      Visible         =   0   'False
      Begin VB.Menu pNewMenu 
         Caption         =   "ˢ���б�"
      End
      Begin VB.Menu pNew 
         Caption         =   "ˢ���б�"
         Begin VB.Menu pNewBySession 
            Caption         =   "SessionProcessLinks(�Ͽ�)"
         End
         Begin VB.Menu pRdNewByHandleList 
            Caption         =   "�����(����ʾ���ؽ���)"
         End
         Begin VB.Menu pNewSh 
            Caption         =   "Toolhelp32"
         End
         Begin VB.Menu pNewByTest 
            Caption         =   "���Դ����н���"
            Visible         =   0   'False
         End
         Begin VB.Menu pNewByQuery 
            Caption         =   "ZwQuerySystemInformation"
         End
         Begin VB.Menu pNewByHandle 
            Caption         =   "Developing..."
            Visible         =   0   'False
         End
      End
      Begin VB.Menu p01 
         Caption         =   "-"
      End
      Begin VB.Menu pListThread 
         Caption         =   "�鿴�����߳�"
      End
      Begin VB.Menu pListModule 
         Caption         =   "�鿴����ģ��"
      End
      Begin VB.Menu pListWindows 
         Caption         =   "�鿴���̴���"
      End
      Begin VB.Menu pListHandles 
         Caption         =   "�鿴���̾��"
      End
      Begin VB.Menu p02 
         Caption         =   "-"
      End
      Begin VB.Menu pJumpToParent 
         Caption         =   "ת��������"
      End
      Begin VB.Menu p03 
         Caption         =   "-"
      End
      Begin VB.Menu pSetPriority 
         Caption         =   "�������ȼ�"
         Begin VB.Menu pPriorityHigh 
            Caption         =   "�ϸ�"
         End
         Begin VB.Menu pPriorityNormal 
            Caption         =   "��׼"
         End
         Begin VB.Menu pPriorityLow 
            Caption         =   "�ϵ�"
         End
      End
      Begin VB.Menu p04 
         Caption         =   "-"
      End
      Begin VB.Menu pSuspendProcess 
         Caption         =   "�������"
      End
      Begin VB.Menu pResumeProcess 
         Caption         =   "�ָ�����"
      End
      Begin VB.Menu pAttach 
         Caption         =   "���ӵ�����"
      End
      Begin VB.Menu pUnlockProcess 
         Caption         =   "���Խ�������"
         Visible         =   0   'False
      End
      Begin VB.Menu p05 
         Caption         =   "-"
      End
      Begin VB.Menu pMoreInformation 
         Caption         =   "��ϸ��Ϣ"
      End
      Begin VB.Menu p06 
         Caption         =   "-"
      End
      Begin VB.Menu pCopyInfo 
         Caption         =   "����ѡ������Ϣ"
         Begin VB.Menu pCopyPid 
            Caption         =   "PID"
         End
         Begin VB.Menu pCopyPEB 
            Caption         =   "PEB"
         End
         Begin VB.Menu pCopyEPROCESS 
            Caption         =   "EPROCESS"
         End
         Begin VB.Menu pCopyName 
            Caption         =   "��������"
         End
         Begin VB.Menu pCopyPath 
            Caption         =   "����·��"
         End
         Begin VB.Menu pCopyCommandLine 
            Caption         =   "������"
         End
      End
      Begin VB.Menu p07 
         Caption         =   "-"
      End
      Begin VB.Menu pMenuTerminateProcess 
         Caption         =   "��������"
         Begin VB.Menu pTerminateProcessNormal 
            Caption         =   "ZwTerminateProcess"
         End
         Begin VB.Menu pTerminateProcessByRemoteThread 
            Caption         =   "CreateRemoteThread->ExitProcess"
         End
         Begin VB.Menu pTerminateProcessByDebugProcess 
            Caption         =   "ZwDebugActiveProcess"
         End
         Begin VB.Menu pWinStationTerminateProcess 
            Caption         =   "WinStationTerminateProcess"
         End
      End
      Begin VB.Menu p08 
         Caption         =   "-"
      End
      Begin VB.Menu pColumnS 
         Caption         =   "��ʾ��"
         Begin VB.Menu pColumnSelect 
            Caption         =   "������"
            Checked         =   -1  'True
            Enabled         =   0   'False
            Index           =   0
         End
      End
      Begin VB.Menu pRestart 
         Caption         =   "ʹ����ͬ����������"
      End
      Begin VB.Menu pReleaseAll 
         Caption         =   "�ͷž��"
         Visible         =   0   'False
      End
   End
   Begin VB.Menu sMenu 
      Caption         =   "����˵�"
      Visible         =   0   'False
      Begin VB.Menu sNew 
         Caption         =   "ˢ���б�"
      End
      Begin VB.Menu s01 
         Caption         =   "-"
      End
      Begin VB.Menu MenuStartServer 
         Caption         =   "��������"
      End
      Begin VB.Menu MenuPauseServer 
         Caption         =   "��ͣ����"
         Visible         =   0   'False
      End
      Begin VB.Menu MenuStopServer 
         Caption         =   "ֹͣ����"
      End
      Begin VB.Menu MenuDeleteServer 
         Caption         =   "ɾ������"
      End
      Begin VB.Menu s02 
         Caption         =   "-"
      End
      Begin VB.Menu MenuS 
         Caption         =   "������������"
         Begin VB.Menu MenuSetAuto 
            Caption         =   "�Զ�����"
         End
         Begin VB.Menu MenuSetUser 
            Caption         =   "�ֶ�����"
         End
         Begin VB.Menu MenuSetCant 
            Caption         =   "��ֹ����"
         End
      End
      Begin VB.Menu s03 
         Caption         =   "-"
      End
      Begin VB.Menu sCopyInfo 
         Caption         =   "����ѡ������Ϣ"
         Begin VB.Menu sCopyServiceName 
            Caption         =   "��������"
         End
         Begin VB.Menu sCopyServiceExePath 
            Caption         =   "ӳ��·��"
         End
         Begin VB.Menu sCopyServiceDllPath 
            Caption         =   "DLL ·��"
         End
         Begin VB.Menu sCopyServiceDescribe 
            Caption         =   "��������"
         End
      End
      Begin VB.Menu s04 
         Caption         =   "-"
      End
      Begin VB.Menu sMoreInformation 
         Caption         =   "��ϸ��Ϣ"
      End
      Begin VB.Menu s05 
         Caption         =   "-"
      End
      Begin VB.Menu sSelectExe 
         Caption         =   "��λ�ļ�"
      End
      Begin VB.Menu sExeNature 
         Caption         =   "�ļ�����"
      End
      Begin VB.Menu s06 
         Caption         =   "-"
      End
      Begin VB.Menu sSelectDll 
         Caption         =   "��λ DLL"
      End
      Begin VB.Menu sDllNature 
         Caption         =   "DLL ����"
      End
   End
   Begin VB.Menu mnuReg 
      Caption         =   "ע���Ŀ¼"
      Visible         =   0   'False
      Begin VB.Menu mnuConnectRemoteReg 
         Caption         =   "����Զ�̵���"
      End
      Begin VB.Menu r01 
         Caption         =   "-"
      End
      Begin VB.Menu rDelete 
         Caption         =   "ɾ����"
      End
      Begin VB.Menu rRename 
         Caption         =   "�޸�����"
         Visible         =   0   'False
      End
      Begin VB.Menu rCreateSubKey 
         Caption         =   "��������"
      End
   End
   Begin VB.Menu mnuRegValue 
      Caption         =   "ֵĿ¼"
      Visible         =   0   'False
      Begin VB.Menu rRefresh 
         Caption         =   "ˢ��"
      End
      Begin VB.Menu r02 
         Caption         =   "-"
      End
      Begin VB.Menu rCreateValue 
         Caption         =   "����ֵ"
      End
      Begin VB.Menu rDeleteKey 
         Caption         =   "ɾ��ֵ"
      End
      Begin VB.Menu rEditValue 
         Caption         =   "�༭ֵ"
      End
   End
   Begin VB.Menu dMenu 
      Caption         =   "�����˵�"
      Visible         =   0   'False
      Begin VB.Menu dNew 
         Caption         =   "ˢ��"
      End
      Begin VB.Menu dDump 
         Caption         =   "ת��"
      End
   End
   Begin VB.Menu sSSDT 
      Caption         =   "SSDT�˵�"
      Visible         =   0   'False
      Begin VB.Menu sRecover 
         Caption         =   "�ָ�"
      End
      Begin VB.Menu sRecoverAll 
         Caption         =   "�ָ�ȫ��"
      End
   End
End
Attribute VB_Name = "Menu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private FirstFocus As Boolean
Private SetText As Boolean

Private Sub Check1_Click()
    If Check1.Value = 0 Then
        SetWindowPos Me.hWnd, HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    ElseIf Check1.Value = 1 Then
        SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    End If
End Sub

Private Sub Check2_Click()
    If Check2.Value Then
        mWindowFilterMethod = mWindowFilterMethod Or MethodListAll
    Else
        mWindowFilterMethod = mWindowFilterMethod And (Not MethodListAll)
    End If
    'If ListView1.Tag = 0 Then
        Call CNNew
    'Else
    '    nChildNewEx ListView1.Tag
    'End If
End Sub

Private Sub Check3_Click()
    If Check3.Value = 1 Then
        Load State
        State.Show
        If Check1.Value = 1 Then SetWindowPos State.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    Else
        Unload State
    End If
End Sub

Private Sub cLock_Click()
    Static Password As String
    Dim A As RECT
    If Text2 = "" Then
        LockStatus = "���벻��Ϊ��."
        Exit Sub
    End If
    If Password <> "" Then
        If Password = Text2 Then
            A.right = Screen.Width
            A.bottom = Screen.Height
            ClipCursor A
            LockStatus = "�����ɹ�."
            cLock.Caption = "�������"
            Password = ""
        Else
            LockStatus = "�������."
        End If
        Exit Sub
    End If
    Password = Text2
    Text2.Text = ""
    GetWindowRect picLock.hWnd, A
    ClipCursor A
    LockStatus = "�����������."
    cLock.Caption = "�������"
    A.left = 0
    A.top = 0
    A.bottom = 99999
    A.right = 99999
End Sub

Private Sub dDump_Click()
    If LVModules.SelectedItem Is Nothing Then Exit Sub
    With LVModules.SelectedItem
        Dim nAddr As Long, nSize As Long, Data() As Byte, pSize As Long
        nAddr = UnFormatHex(.SubItems(2))
        nSize = UnFormatHex(.SubItems(3))
        ReDim Data(1 To nSize)
         ReadKernelMemory nAddr, VarPtr(Data(1)), nSize, pSize
        If pSize = 0 Then
            MsgBox "ת��ʧ�ܣ�", vbCritical
        End If
        ReDim Preserve Data(1 To pSize)
        cd.Filter = "�����ļ�(*.*)|*.*|�����ļ�(*.sys)|*.sys"
        On Error GoTo e
        cd.CancelError = True
        cd.ShowSave
        Open cd.FileName For Binary As #1
        Put #1, , Data
        Close #1
    End With
e:
End Sub

Private Sub Form_LinkExecute(CmdStr As String, Cancel As Integer)
    If CmdStr = Date Then AutoUpdate Me
    Cancel = 0
End Sub

Private Sub Form_Resize()
    If Me.WindowState = 1 Then Exit Sub '��С��������
    #If False Then
    '����Ҫ�����ţ�
    SSTab1.Width = ScaleWidth - 240
    SSTab1.Height = ScaleHeight - 195
    '�Դ����������ţ�
    Check2.Move SSTab1.Width - 1665, SSTab1.Height - 495
    Check3.Move SSTab1.Width - 2745, SSTab1.Height - 495
    Check1.Move SSTab1.Width - 3825, SSTab1.Height - 495
    Text1.Move SSTab1.Width - 10185, SSTab1.Height - 495
    ListView1.Height = SSTab1.Height - 1110
    ListView1.Width = SSTab1.Width - 225
    '�Խ����������ţ�
    ListView2.Height = SSTab1.Height - 1020
    ListView2.Width = SSTab1.Width - 240
    pcNewTask.Move 1800, SSTab1.Height - 555
    '�Է����������ţ�
    LVServer.Height = SSTab1.Height - 1020
    LVServer.Width = SSTab1.Width - 240
    '��ע�������ţ�
    tvwKeys.Width = SSTab1.Width / 12345 * 2535
    lvwData.Width = SSTab1.Width / 12345 * 9465
    lvwData.left = tvwKeys.Width + 225
    With lvwData.ColumnHeaders
        .Clear
        .Add , , "����", lvwData.Width / 6240 * 1440
        .Add , , "����", lvwData.Width / 6240 * 1000
        .Add , , "ֵ", lvwData.Width / 6240 * 3750
    End With
    #End If
End Sub

Public Sub lLabels_Click(Index As Integer)
    Dim i As Long
    On Error Resume Next
    For i = 0 To pTab.UBound
        pTab(i).Visible = False
        lLabels(i).BorderStyle = 0
    Next
    pTab(Index).Visible = True
    lLabels(Index).BorderStyle = 1
    pTab(Index).Move 1440, 120
    
    If SetTab Then Exit Sub
    Select Case Index
    Case 0
        SetText = True
        Text1 = "��������������������"
        mWindowFilterMethod = 0
        Call CNNew
        SetText = False
        FirstFocus = True
    Case 1
        Call PNNew
    Case 2
        Call msNew_Click
    Case 3
        mnuReg.Visible = True
    Case 4
        Call GMNew
    Case 7
        Call GetSSDT
        Call ListSSDT
    Case 8
        Call GetSSDT
        Call ListShadowSSDT
    End Select
    If Index <> 3 Then
        mnuReg.Visible = False
    End If
End Sub

Private Sub LVModules_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then PopupMenu dMenu
End Sub

Private Sub LVSSDT_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        If LVSSDT.SelectedItem Is Nothing Then Exit Sub
        PopupMenu sSSDT
    End If
End Sub

Private Sub lvwData_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If tvwKeys.SelectedItem Is Nothing Then Exit Sub
    If Button = 2 Then
        If lvwData.SelectedItem Is Nothing Then
            rDeleteKey.Enabled = False
            rEditValue.Enabled = False
        Else
            rDeleteKey.Enabled = True
            rEditValue.Enabled = True
        End If
        PopupMenu mnuRegValue
    End If
End Sub

Private Sub mainReadme_Click()
    Load About
    About.Show
End Sub

Private Sub mnuRegRefresh_Click()
    tvwKeys.Nodes.Clear
End Sub

Private Sub mainShutdown_Click()
    ShutdownWindow.Show vbModal, Me
End Sub

Private Sub MenuDeleteServer_Click()
    If MsgBox("��ȷ��Ҫɾ���˷�����", vbQuestion + vbYesNo) Then
        Dim hKey As Long
        hKey = OpenRegKey("�ҵĵ���\HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Service\" & LVServer.SelectedItem.Text, KEY_ALL_ACCESS, True)
        If hKey = 0 Then Exit Sub
        ZwClose hKey
    End If
End Sub

Private Sub mnuConnectRemoteReg_Click()
    Exit Sub
    tvwKeys.Nodes.Clear
    Call InitRegTree(InputBox("������Զ�̵�����"))
    If tvwKeys.Nodes.count = 1 Then
        tvwKeys.Nodes.Clear
        MsgBox "����ʧ��", vbCritical
        Call InitRegTree
    End If
End Sub

Private Sub nChildNewMenu_Click()
    Call CNNew
End Sub

Private Sub nCloseWindowByEndTask_Click()
    Dim Ret As Long

    Ret = EndTask(CLng(ListView1.SelectedItem.SubItems(2)), 0, 1)
    
    If Ret <> 1 Then MsgBox "�رմ���ʧ��!", 0, "ʧ��": Exit Sub
    
    Call CNNew
End Sub

Private Sub nCloseWindowByMessage_Click()
    FxBombWindow CLng(ListView1.SelectedItem.SubItems(2))
    
    Call CNNew
End Sub

Private Sub nCNNew_Click()
    Call CNNew
End Sub

Private Sub nFdNewByMessage_click()
    '��������FdEnumWindowsByMessage
    SetWindowMethod MethodPostMessage
    Call CNNew
End Sub

Private Sub Form_Load()
    'SetIcon Me.hWnd, "IDR_MAINFRAME", True 'icon
    Dim i As Long
    
    SetStatus "���ش���..."
    
    With ListView1.ColumnHeaders
        .Add , , "����", 3200
        .Add , , "����", 3200
        .Add , , "���", 930
        .Add , , "�������", 910
        .Add , , "PID", 660
        .Add , , "TID", 660
        .Add , , "״̬", 1200
    End With
    
    With ListView2.ColumnHeaders
        For i = 0 To ProcessColumnCount
            If i <> 0 Then
                Load pColumnSelect(i)
                With pColumnSelect(i)
                    .Enabled = True
                    .Caption = ProcessColumnNames(i)
                End With
            End If
            If left(ProcessColumnSetting(i), 2) = "-1" Then
                .Add , , ProcessColumnNames(i), ProcessColumnWidth(i)
                pColumnSelect(i).Checked = True
            Else
                pColumnSelect(i).Checked = False
            End If
        Next
    End With
       
    With LVServer.ColumnHeaders
        .Add , , "����", 2600
        .Add , , "״̬", 1000
        .Add , , "��������", 1000
        .Add , , "·��", 3500   '2000
        .Add , , "����", 4000
        .Add , , "��¼���", 1400
        .Add , , "��̬���ӿ�·��", 3500   '1400
    
    End With
    
    With LVModules.ColumnHeaders
        .Add , , "ģ����", 1440
        .Add , , "ģ��·��", 2880
        .Add , , "ģ���ַ", 1100
        .Add , , "ģ���С", 1100
        .Add , , "���ش���", 1440
    End With
    
    With lvwData.ColumnHeaders
        .Add , , "����", 1440
        .Add , , "����", 1440
        .Add , , "����", 2880
    End With
    
    With LVSSDT.ColumnHeaders
        .Add , , "���", 800
        .Add , , "������", 3000
        .Add , , "ԭʼ��ַ", 1200
        .Add , , "���ڵ�ַ", 1200
        .Add , , "����ģ��", 3600
    End With
    
    With LVShadowSSDT.ColumnHeaders
        .Add , , "���", 800
        .Add , , "������", 3000
        .Add , , "ԭʼ��ַ", 1200
        .Add , , "���ڵ�ַ", 1200
        .Add , , "����ģ��", 3600
    End With
    

    FirstFocus = True
    ListView1.Tag = 0
    Dim S As String
    S = Chr(0)
    ReadINI "Soft Settings", vbNullString, "Enum windows method", S
    SetWindowMethod Val(Split(S, Chr(0))(0))
    
    If Check1.Value = 1 Then SetWindowPos Me.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE

    ListView1.Sorted = True
    SetTextColor Me
    Me.Caption = "Azmrk WindowsXP Edition v" & App.Major & "." & App.Minor & "." & App.Revision
    
    '2010-02-24  ���沼��
    'Label1.top = Text1.top + 50
    'ע������
    On Error GoTo errTrap ' just in case ;)
    InitRegTree
    
    If SoftValue(0) = "1" Then Check1.Value = 1: Call Check1_Click
    If SoftValue(1) = "1" Then Check2.Value = 1: Call Check2_Click
    If SoftValue(2) = "1" Then Check3.Value = 1: Call Check3_Click
    
    lLabels_Click 0
    Exit Sub
errTrap:
    Dim Msg As String
    Msg = "δ֪����!" & vbCrLf
    Msg = Msg & "����: " & Err.Description & String(2, vbCrLf)
    MsgBox Msg, vbExclamation, "����: " & Err.Number
End Sub
    
Private Function InitRegTree(Optional ByVal szComputerName As String = "�ҵĵ���") As Boolean
    Dim NodX As Object
    Set NodX = tvwKeys.Nodes.Add(, , szComputerName, szComputerName, 5) '����ҵĵ��Խڵ�
    NodX.Expanded = True '
    
    'Set lastNode = NodX
    
    '����
    On Error Resume Next
    Dim rootname
    For Each rootname In Array("HKEY_LOCAL_MACHINE", _
                               "HKEY_CLASSES_ROOT", _
                               "HKEY_CURRENT_USER", _
                               "HKEY_CURRENT_CONFIG", _
                               "HKEY_USERS")
        tvwKeys.Nodes.Add szComputerName, tvwChild, szComputerName & "\" & rootname, rootname
        tvwKeys.Nodes.Add szComputerName & "\" & rootname, tvwChild
    Next
ErrHand:
    Resume Next
End Function


Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    Unload State
    Unload LoginPic
    Unload ModuleList
    Unload ThreadList
    Unload HandleList
    Unload ProcessStart
    Unload ShutdownWindow
    Unload State
    Unload WaitWindow
    Unload About
    Unload CreateValue
    Unload DGEditDWord
    Unload DGEditValue
    Dim i As Long
    For i = 0 To UBound(Processes)
        ZwClose Processes(i).Handle
    Next
    ShutdownSSDT
    
    'End
    'If IsIDE Then
        End
    'Else
    '    ExitProcess 0
    'End If
End Sub

Private Sub Label2_Click()
    ShellExecute Me.hWnd, "open", "http://hi.baidu.com/dazzles", vbNullString, vbNullString, SW_HIDE
    'http://sighttp.qq.com/cgi-bin/check?sigkey=10e2f1de4f3638083759f062e8997cd18e83e614ff27ed6511e0665cc7ab711b
End Sub

Private Sub Label2_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 99
    Me.MouseIcon = VB.LoadResPicture(101, vbResCursor)
End Sub

Private Sub Label4_Click()
    ShellExecute Me.hWnd, "open", "http://hi.baidu.com/naylonslain", vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub Label4_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Me.MousePointer = 99
    Me.MouseIcon = VB.LoadResPicture(101, vbResCursor)
End Sub

Private Sub ListView1_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LVAutoOrder ListView1, ColumnHeader
End Sub

Private Sub ListView1_DblClick()
    Dim news As Long
    Dim nIndex As Long
    If Check2.Value Then Exit Sub '��ʾ���д���ʱ��֧��ֻ��ʾĳ���ڵ��Ӵ�����
    
    Text1.Text = ""
    
    With ListView1
        If (.Tag = 0) Or (.Tag <> 0 And .SelectedItem.Text <> "..") Then   '�������ת���Ӵ����
            nSelectedItemIndex(.Tag) = .SelectedItem.Index   '��¼��ǰѡ��������
            .Tag = .Tag + 1
            nSelectedItem(.Tag) = CLng(.SelectedItem.SubItems(2))   '��¼��ǰѡ����ľ��
            .ListItems.Clear
            EnumAllChildWindows nSelectedItem(.Tag), ""
        ElseIf .Tag = 1 And .SelectedItem.Text = ".." Then   '�Ӵ����ת�븸�����
            .Tag = 0
            Text1.Text = ""
            Call CNNew
            FxSetListviewNowLine ListView1, nSelectedItemIndex(.Tag)
        ElseIf .Tag > 1 And .SelectedItem.Text = ".." Then   '�ﴰ���ת���Ӵ����
            .Tag = .Tag - 1
            .ListItems.Clear
            EnumAllChildWindows nSelectedItem(.Tag), ""
            FxSetListviewNowLine ListView1, nSelectedItemIndex(.Tag)
        End If
    End With
End Sub

Public Sub ListView1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SubText As String

    With ListView1
        If Button = 2 Then
            If .SelectedItem Is Nothing Then
                Exit Sub
            End If
            SubText = .SelectedItem.SubItems(6)
            
            If .Tag > 0 And .SelectedItem.SubItems(1) = "" Then Exit Sub

            '���ò˵�ѡ��
            If InStr(SubText, "����") > 0 Then 'show and hide
                nHide.Enabled = False
                nShow.Enabled = True
            ElseIf InStr(SubText, "�ɼ�") > 0 Then
                nHide.Enabled = True
                nShow.Enabled = False
            End If
            
            If InStr(SubText, "����") > 0 Then 'enabled or not
                nEnableF.Enabled = True
                nEnableT.Enabled = False
            ElseIf InStr(SubText, "����") > 0 Then
                nEnableF.Enabled = False
                nEnableT.Enabled = True
            End If
            
            If InStr(SubText, "���") > 0 Or InStr(SubText, "����") > 0 Then
                nWindowMax.Enabled = False
                nWindowMin.Enabled = True
            ElseIf InStr(SubText, "��С") > 0 Or InStr(SubText, "����") > 0 Then
                nWindowMax.Enabled = True
                nWindowMin.Enabled = False
            End If
            nJumpToParent.Visible = True
            nViewFatherWindows.Visible = False
            n09.Visible = True
            nTop.Visible = True
            nNoTop.Visible = True
            nNewMenu.Visible = True
            If .Tag = 0 Then
                nChildNewMenu.Enabled = False
                nJumpToParent.Enabled = CBool(Check2.Value)
                nViewFatherWindows.Enabled = False
            End If
            
            If Check2.Value = 1 Then
                nNew.Enabled = False
                nFxNew.Enabled = False
            Else
                nNew.Enabled = True
                nFxNew.Enabled = True
            End If
            
            PopupMenu FatherMenus
        End If
    End With
End Sub

Private Sub ListView2_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LVAutoOrder ListView2, ColumnHeader
End Sub

Private Sub ListView2_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If Button = 2 Then
        '���ò˵�״̬
        If ListView2.Tag = 0 Then
            Dim SubText As String
            SubText = ListView2.SelectedItem.SubItems(4)
            If InStr(SubText, "�ϸ�") > 0 Then
                pPriorityHigh.Enabled = False
                pPriorityNormal.Enabled = True
                pPriorityLow.Enabled = True
            ElseIf InStr(SubText, "��׼") > 0 Then
                pPriorityHigh.Enabled = True
                pPriorityNormal.Enabled = False
                pPriorityLow.Enabled = True
            ElseIf InStr(SubText, "�ϵ�") > 0 Then
                pPriorityHigh.Enabled = True
                pPriorityNormal.Enabled = True
                pPriorityLow.Enabled = False
            End If
        Else
            pSetPriority.Enabled = False
        End If
        
        PopupMenu pMenu
        pSetPriority.Enabled = True
    End If

    'ListViewColor Me, ListView2
End Sub

Private Sub ListView3_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    LVAutoOrder LVServer, ColumnHeader
End Sub

Private Sub ListView3_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 Then
        PopupMenu dMenu
    End If
End Sub

Private Sub LVServer_MouseUp(Button As Integer, _
                             Shift As Integer, _
                             X As Single, _
                             Y As Single)

    If Button = 2 Then

        Select Case LVServer.SelectedItem.SubItems(1)

            Case "������"
                MenuStartServer.Enabled = False
                MenuPauseServer.Enabled = True
                MenuStopServer.Enabled = True

            Case "��ֹͣ"
                MenuStartServer.Enabled = True
                MenuPauseServer.Enabled = False
                MenuStopServer.Enabled = False

        End Select
        
        Select Case LVServer.SelectedItem.SubItems(2)
        
            Case "�Զ�"
                MenuSetAuto.Enabled = False
                MenuSetUser.Enabled = True
                MenuSetCant.Enabled = True

            Case "�ֶ�"
                MenuSetAuto.Enabled = True
                MenuSetUser.Enabled = False
                MenuSetCant.Enabled = True

            Case "����"
                MenuSetAuto.Enabled = True
                MenuSetUser.Enabled = True
                MenuSetUser.Enabled = False
        
        End Select

        Me.PopupMenu sMenu
    End If

End Sub

Private Sub MenuPauseServer_Click()
    Dim Registry As clsRegistry
    Set Registry = New clsRegistry
    Dim r_initial   As String
    Dim rv_value    As String
    Dim serv_status As String
    
    r_initial = LVServer.SelectedItem

    If r_initial = "" Then Exit Sub
    rv_value = Registry.GetValue(eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\" & r_initial, "Description")

    If rv_value = "" Then
        ServicePause "", r_initial
        msNew_Click
        Exit Sub
    End If

    ServicePause "", r_initial
    msNew_Click
End Sub

Private Sub MenuSetAuto_Click()
    SetServerBootType LVServer.SelectedItem, 2
End Sub

Private Sub MenuSetCant_Click()
    SetServerBootType LVServer.SelectedItem, 4
End Sub

Private Sub MenuSetUser_Click()
    SetServerBootType LVServer.SelectedItem, 3
End Sub

Private Sub MenuStartServer_Click()
    Dim Registry As clsRegistry
    Set Registry = New clsRegistry
    Dim r_initial   As String
    Dim rv_value    As String
    Dim serv_status As String
    
    r_initial = LVServer.SelectedItem

    If r_initial = "" Then Exit Sub
    rv_value = Registry.GetValue(eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\" & r_initial, "Description")

    If rv_value = "" Then
        ServiceStart "", r_initial
        msNew_Click
        Exit Sub
    End If
    ServiceStart "", r_initial
    msNew_Click
End Sub

Private Sub MenuStopServer_Click()
    Dim Registry As clsRegistry
    Set Registry = New clsRegistry
    Dim r_initial   As String
    Dim rv_value    As String
    Dim serv_status As String
    
    r_initial = LVServer.SelectedItem

    If r_initial = "" Then Exit Sub
    rv_value = Registry.GetValue(eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\" & r_initial, "Description")

    If rv_value = "" Then
        ServiceStop "", r_initial
        msNew_Click
        Exit Sub
    End If

    ServiceStop "", r_initial
    msNew_Click
End Sub

Private Sub nAmend_Click()
    Dim newText As String
   
    newText = InputBox("������µ����ݣ�", "�޸ı���", ListView1.SelectedItem.Text)
    'If newText = "" Or newText = ListView1.SelectedItem.Text Then
    '    MsgBox "���޸����ݣ���ֵ����Ϊ�գ�", vbOKOnly + vbInformation, "����"
    '    Exit Sub
    'End If
    SetWindowText ListView1.SelectedItem.SubItems(2), newText   'ListView1.SelectedItem.SubItems(2)
    Call CNNew
End Sub

Private Sub nChildNewEx(ByVal Index As Long)
End Sub

Private Sub nClose_Click()
    PostMessage CLng(ListView1.SelectedItem.SubItems(2)), WM_CLOSE, 0, ByVal 0
    PostMessage CLng(ListView1.SelectedItem.SubItems(2)), WM_DESTROY, 0, ByVal 0
    
    Call CNNew
End Sub

Private Sub nCloseWindowByParent_Click()
    FxCloseWindowByParent CLng(ListView1.SelectedItem.SubItems(2))
    
    Call CNNew
End Sub

Private Sub nCloseWindowByWndProc_Click()
    FxCloseWindowByWndProc CLng(ListView1.SelectedItem.SubItems(2))
    
    Call CNNew
End Sub

Private Sub nCopyClass_Click()
    Clipboard.Clear
    Clipboard.SetText ListView1.SelectedItem.SubItems(1), 1
End Sub

Private Sub nCopyHandle_Click()
    Clipboard.Clear
    Clipboard.SetText ListView1.SelectedItem.SubItems(2), 1
End Sub

Private Sub nCopyName_Click()
    Clipboard.Clear
    Clipboard.SetText ListView1.SelectedItem.Text, 1
End Sub

Private Sub nEnableF_Click()
    EnableWindow CLng(ListView1.SelectedItem.SubItems(2)), 0
    
    Call CNNew
End Sub

Private Sub nEnableT_Click()
    EnableWindow CLng(ListView1.SelectedItem.SubItems(2)), 1
    
    Call CNNew
End Sub

Private Sub nFxNew_Click()
    SetWindowMethod MethodParent
    
    Call CNNew
End Sub

Private Sub nFxNewByTID_Click()
    SetWindowMethod MethodGetThread
    
    Call CNNew
End Sub

Private Sub nGetTextBox_Click()
    Dim X As New FormTextBox
    X.nWnd = CLng(ListView1.SelectedItem.SubItems(2))
    X.CatchText
    X.Show
End Sub

Private Sub nHide_Click()
    ShowWindow CLng(ListView1.SelectedItem.SubItems(2)), SW_HIDE  'SW_HIDE=0
    
    Call CNNew
End Sub

Private Sub nJumpToParent_Click()
    Dim myId As Long
    Dim i As Long
    
    myId = CLng(ListView1.SelectedItem.SubItems(3))
    
    For i = 1 To ListView1.ListItems.count
        If CLng(ListView1.ListItems(i).SubItems(2)) = myId Then
            ListView1.ListItems(i).Selected = True
            ListView1.ListItems(i).EnsureVisible
            Exit For
            Exit Sub
        End If
    Next i
End Sub

Private Sub nJumpToTasklist_Click()
    Dim i As Long
    Dim jmpid As Long
    
    jmpid = CLng(ListView1.SelectedItem.SubItems(4))
    
    For i = 1 To ListView2.ListItems.count
        If ListView2.ListItems(i).SubItems(1) = jmpid Then
            FxSetListviewNowLine ListView2, i
            lLabels_Click 1
        End If
    Next i
End Sub

Private Sub nJumpToThread_Click()
    On Error Resume Next
    nsItem = CLng(ListView1.SelectedItem.SubItems(4))
    Dim wList As ThreadList
    Set wList = New ThreadList
    If Check1.Value = 1 Then SetWindowPos wList.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    wList.Show
    Dim i As ListItem
    For Each i In wList.ListView1.ListItems
        If i.Text = ListView1.SelectedItem.SubItems(5) Then
            i.Selected = True
            i.EnsureVisible
            wList.SetFocus
            Exit Sub
        End If
    Next
End Sub

Private Sub nMove_Click()
    Dim MoveTo As String
    
    MoveTo = InputBox("�������µ����꣬���Ÿ���������303,505", "��������")
    If CLng(InStr(1, MoveTo, ",")) = 0 Or MoveTo = "" Then Exit Sub
    SetWindowPos CLng(ListView1.SelectedItem.SubItems(2)), 0, CLng(Mid(MoveTo, 1, InStr(1, MoveTo, ",") - 1)), CLng(Mid(MoveTo, InStr(1, MoveTo, ",") + 1, Len(MoveTo))), 0, 0, SWP_NOSIZE
End Sub

Private Sub nNew_Click()
    SetWindowMethod MethodEnumWindows
    
    Call CNNew
End Sub

Private Sub nNoTop_Click()
    SetWindowPos CLng(ListView1.SelectedItem.SubItems(2)), HWND_NOTOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    Call CNNew
End Sub

Private Sub nRwNewByIsWindow_Click()
    SetWindowMethod MethodIsWindow
    
    Call CNNew
End Sub

Private Sub nShow_Click()
    ShowWindow CLng(ListView1.SelectedItem.SubItems(2)), SW_SHOW  'SW_SHOW=1;SW_SHOWNOACTIVATE=4
    
    Call CNNew
End Sub

Private Sub nTop_Click()
    'SetWindowPos CLng(ListView1.SelectedItem.SubItems(2)), HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    SetWindowPos CLng(ListView1.SelectedItem.SubItems(2)), 1, 0, 0, 0, 0, -3
    ShowWindow CLng(ListView1.SelectedItem.SubItems(2)), SW_SHOWNOACTIVATE
    
    Call CNNew
End Sub

Private Sub nViewChildWindows_Click()
    Call ListView1_DblClick
End Sub

Private Sub nViewFatherWindows_Click()
    Dim myHwnd As Long
    
    With ListView1
        myHwnd = .SelectedItem.SubItems(3)

        .ListItems(1).Selected = True
        
        Call ListView1_DblClick
        
        Dim i As Long
        
        For i = 1 To .ListItems.count
            If .ListItems(i).SubItems(2) = myHwnd Then
                FxSetListviewNowLine ListView1, i
                Exit For
            End If
        Next i
    End With
End Sub

Private Sub nWindowMax_Click()
    ShowWindow CLng(ListView1.SelectedItem.SubItems(2)), SW_MAXIMIZE
    
    Call CNNew
End Sub

Private Sub nWindowMin_Click()
    ShowWindow CLng(ListView1.SelectedItem.SubItems(2)), SW_MINIMIZE
    
    Call CNNew
End Sub

Private Sub pAttach_Click()
    Dim nSel As ListItem
    Set nSel = ListView2.SelectedItem
    If nSel Is Nothing Then Exit Sub
    Dim n As Long
    If CheckFor("BeaEngine.dll", "") = False Then
        MsgBox "û���ҵ�BeaEngine.dll���޷����е��ԡ�", vbCritical
        pAttach.Enabled = False
        Exit Sub
    End If
    For n = 0 To UBound(Processes)
        With Processes(n)
            If .ListViewIndex = nSel.Index Then
                If MsgBox("ȷ��Ҫ�ڴ˽����ϸ��ӵ�������", vbQuestion Or vbYesNo) = vbNo Then Exit Sub
                AttachDebugger .Basic.UniqueProcessId
                Exit Sub
            End If
        End With
    Next
End Sub

Private Sub pcNewTask_Click()
    ProcessStart.Show vbModal, Me
End Sub

Private Sub pColumnSelect_Click(Index As Integer)
    Dim j As Long, i As Long
    pColumnSelect(Index).Checked = Not pColumnSelect(Index).Checked
    On Error Resume Next
    For j = 1 To Index
        If ListView2.ColumnHeaders(j - i + 1).Text <> pColumnSelect(j).Caption Then
            i = i + 1
        End If
    Next
    If pColumnSelect(Index).Checked Then
        ListView2.ColumnHeaders.Add Index + 2 - i, , ProcessColumnNames(Index), ProcessColumnWidth(Index)
        For j = 0 To UBound(Processes)
            Dim n As ListSubItem
            With Processes(j)
                ListView2.ListItems(.ListViewIndex).ListSubItems.Add(Index).Text = FillSubItem(Processes(j), Index)
            End With
        Next
    Else
        ListView2.ColumnHeaders.Remove Index + 1 - i
        For j = 1 To ListView2.ListItems.count
            ListView2.ListItems(j).ListSubItems.Remove Index
        Next
    End If
    ListView2.Refresh
End Sub

Private Sub pCopyCommandLine_Click()
    With Clipboard
        .Clear
        .SetText ListView2.SelectedItem.SubItems(8)
    End With
End Sub

Private Sub pCopyEPROCESS_Click()
    With Clipboard
        .Clear
        .SetText ListView2.SelectedItem.SubItems(4)
    End With
End Sub

Private Sub pCopyName_Click()
    With Clipboard
        .Clear
        .SetText ListView2.SelectedItem.Text
    End With
End Sub

Private Sub pCopyPath_Click()
    With Clipboard
        .Clear
        .SetText ListView2.SelectedItem.SubItems(7)
    End With
End Sub

Private Sub pCopyPEB_Click()
    With Clipboard
        .Clear
        .SetText ListView2.SelectedItem.SubItems(3)
    End With
End Sub

Private Sub pCopyPid_Click()
    With Clipboard
        .Clear
        .SetText CStr(ListView2.SelectedItem.SubItems(1))
    End With
End Sub

Private Sub pJumpToParent_Click()
    Dim myId As Long
    Dim i As Long
    
    myId = CLng(ListView2.SelectedItem.SubItems(2))
    
    For i = 1 To ListView2.ListItems.count
        If ListView2.ListItems(i).SubItems(1) = myId Then
            ListView2.ListItems(i).Selected = True
            ListView2.ListItems(i).EnsureVisible
            Exit For
            Exit Sub
        End If
    Next i
End Sub

Private Sub pListHandles_Click()
    On Error Resume Next
    
    nsItem = UnFormatHex(ListView2.SelectedItem.SubItems(4))
    Dim A As HandleList
    Set A = New HandleList
    
    If Check1.Value = 1 Then SetWindowPos A.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    A.Show
End Sub

Private Sub pListModule_Click()
    On Error Resume Next
       
    nsItem = CLng(ListView2.SelectedItem.SubItems(1))
    Dim A As ModuleList
    Set A = New ModuleList
    
    If Check1.Value = 1 Then SetWindowPos A.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    A.Show
End Sub

Private Sub pListThread_Click()
    On Error Resume Next

    nsItem = CLng(ListView2.SelectedItem.SubItems(1))

    Dim wList As ThreadList
    Set wList = New ThreadList
    
    If Check1.Value = 1 Then SetWindowPos wList.hWnd, HWND_TOPMOST, 0, 0, 0, 0, SWP_NOMOVE Or SWP_NOSIZE
    
    wList.Show
End Sub

Private Sub pListWindows_Click()
    On Error Resume Next
    
    'viewProcessWindows = CLng(ListView2.SelectedItem.SubItems(1))
    mWindowFilterMethod = mWindowFilterMethod Or MethodListByPID
    mWindowFilterArg = CLng(ListView2.SelectedItem.SubItems(1))
    SetTab = True
    lLabels_Click 0
    SetTab = False
    ListView1.Tag = 0
    Call CNNew
End Sub

Private Sub pMoreInformation_Click()
    ShellExecute Me.hWnd, "open", "http://www.baidu.com/s?wd=" & (ListView2.SelectedItem.Text), vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub pNewByHandle_Click()
    ListView2.Tag = MethodHandleList
    Call PNNew
End Sub

Private Sub pNewByQuery_Click()
    ListView2.Tag = MethodQuery
    Call PNNew
End Sub

Private Sub pNewBySession_Click()
    ListView2.Tag = MethodSession
    Call PNNew
End Sub

Private Sub pNewByTest_Click()
    ListView2.Tag = MethodTest
    Call PNNew
End Sub

Private Sub pNewMenu_Click()
    Call PNNew
End Sub

Private Sub pNewSh_Click()
    ListView2.Tag = MethodSnapshot
    Call PNNew
End Sub

Private Sub pNewWmi_Click()
    'ListView2.Tag = MethodWmi
    'Call PNNew
End Sub

Private Sub pPriorityHigh_Click()
    Dim hProcess As Long
    
    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, CLng(ListView2.SelectedItem.SubItems(1)))
    SetPriorityClass hProcess, HIGH_PRIORITY_CLASS
       
    ZwClose hProcess
    Call PNNew
End Sub

Private Sub pPriorityLow_Click()
    Dim hProcess As Long
    
    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, CLng(ListView2.SelectedItem.SubItems(1)))
    SetPriorityClass hProcess, IDLE_PRIORITY_CLASS
    
    ZwClose hProcess
    Call PNNew
End Sub

Private Sub pPriorityNormal_Click()
    Dim hProcess As Long
    
    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, CLng(ListView2.SelectedItem.SubItems(1)))
    SetPriorityClass hProcess, NORMAL_PRIORITY_CLASS
    
    ZwClose hProcess
    Call PNNew
End Sub

Private Sub pRdNewByHandleList_Click()
    ListView2.Tag = MethodHandleList
    Call PNNew
End Sub

Private Sub pReleaseAll_Click()
    Load WaitWindow
    'WaitWindow.BeginReleaseAll
End Sub

Private Sub pRestart_Click()
    Dim hProcess As Long
    
    hProcess = FxNormalOpenProcess(PROCESS_TERMINATE, CLng(ListView2.SelectedItem.SubItems(1)))
    ZwTerminateProcess hProcess, 0
    ZwClose hProcess
    Shell ListView2.SelectedItem.SubItems(8)
End Sub

Private Sub pResumeProcess_Click()
    'SusResProcess ListView2.SelectedItem.SubItems(1), False
    'DoEvents
    '---������ͨ���ָ����̵������߳�---
    
    Dim hProcess As Long
    
    hProcess = FxNormalOpenProcess(PROCESS_SUSPEND_RESUME, CLng(ListView2.SelectedItem.SubItems(1)))
    ZwResumeProcess hProcess
    
    ZwClose hProcess
    Call PNNew
End Sub

Private Sub pSuspendProcess_Click()
    'SusResProcess ListView2.SelectedItem.SubItems(1), True
    'DoEvents
    '---������ͨ��������̵������߳�---
    
    Dim hProcess As Long
    
    hProcess = FxNormalOpenProcess(PROCESS_SUSPEND_RESUME, CLng(ListView2.SelectedItem.SubItems(1)))
    ZwSuspendProcess hProcess
    
    ZwClose hProcess
    Call PNNew
End Sub

Private Sub pTerminateProcessByDebugProcess_Click()
    FxTerminateProcessByDebugProcess CLng(ListView2.SelectedItem.SubItems(1))
    Call PNNew
End Sub

Private Sub pTerminateProcessByRemoteThread_Click()
    Dim hProcess As Long, hThread As Long, hFunction As Long
    Dim lpThreadAttributes As SECURITY_ATTRIBUTES
    
    hProcess = FxNormalOpenProcess(PROCESS_ALL_ACCESS, CLng(ListView2.SelectedItem.SubItems(1)))
    If hProcess = 0 Then
        MsgBox "�ܾ�����!", 0, "ʧ��"
        Exit Sub
    End If
    
    hFunction = GetModuleHandle("kernel32.dll")
    hFunction = GetProcAddress(hFunction, "ExitProcess")
    
    hThread = CreateRemoteThread(hProcess, lpThreadAttributes, 0, hFunction, 0, 0, 0)
    If hThread = 0 Then
        MsgBox "�����߳�ʧ��!", 0, "ʧ��"
        ZwClose hProcess
        Exit Sub
    End If
    
    'WaitForSingleObject hThread, INFINITE
    
    ZwClose hThread
    ZwClose hProcess
    
    Call PNNew
End Sub

Private Sub pTerminateProcessNormal_Click()
    Dim hProcess As Long
    
    hProcess = FxNormalOpenProcess(PROCESS_TERMINATE, CLng(ListView2.SelectedItem.SubItems(1)))

    If hProcess = 0 Then
        MsgBox "�ܾ�����!", 0, "ʧ��"
        Exit Sub
    End If
    
    ZwTerminateProcess hProcess, 0
    
    WaitForSingleObject hProcess, INFINITE
    
    ZwClose hProcess
    
    Call PNNew
End Sub

Private Sub pUnlockProcess_Click()
    Call RdUnlockProcess(UnFormatHex(ListView2.SelectedItem.SubItems(4)))
End Sub

Private Sub pWinStationTerminateProcess_Click()
    Dim Ret As Long
    
    Ret = WinStationTerminateProcess(WTS_CURRENT_SERVER_HANDLE, CLng(ListView2.SelectedItem.SubItems(1)), 0)
    
    If Ret <> 1 Then MsgBox "��������ʧ��!", 0, "ʧ��": Exit Sub
    
    Call PNNew
End Sub

Private Sub rCreateSubKey_Click()
    With tvwKeys
        If .SelectedItem Is Nothing Then Exit Sub
        Dim sSubKey As String
        sSubKey = InputBox("��������������")
        If sSubKey = "" Then Exit Sub
        Call CreateKey(.SelectedItem.FullPath, sSubKey, True)
        Call EnumReg(.SelectedItem.Parent)
    End With
End Sub

Private Sub rCreateValue_Click()
    Dim wKey As String
    wKey = tvwKeys.SelectedItem.FullPath
    CreateValue.Init wKey
    CreateValue.Show vbModal, Me
End Sub

Private Sub rDelete_Click()
    With tvwKeys
        If .SelectedItem Is Nothing Then Exit Sub
        DeleteKey .SelectedItem.FullPath, True
        Call EnumReg(.SelectedItem.Parent)
    End With
End Sub

Private Sub rDeleteKey_Click()
    With tvwKeys
        If .SelectedItem Is Nothing Then Exit Sub
        DeleteValue .SelectedItem.FullPath, lvwData.SelectedItem.Text, True
        Call EnumValue(.SelectedItem)
    End With
End Sub

Private Sub rEditValue_Click()
    On Error Resume Next
    SetReg tvwKeys.SelectedItem.FullPath, lvwData.SelectedItem.Text, lvwData.SelectedItem.SubItems(1), True, lvwData.SelectedItem.Index = 0
End Sub

Private Sub rRefresh_Click()
    Call EnumValue(tvwKeys.SelectedItem, True)
End Sub

Private Sub sCopyServiceDescribe_Click()
    Clipboard.Clear
    Clipboard.SetText LVServer.SelectedItem.SubItems(4), 1
End Sub

Private Sub sCopyServiceDllPath_Click()
    Clipboard.Clear
    Clipboard.SetText LVServer.SelectedItem.SubItems(6), 1
End Sub

Private Sub sCopyServiceExePath_Click()
    Clipboard.Clear
    Clipboard.SetText LVServer.SelectedItem.SubItems(3), 1
End Sub

Private Sub sCopyServiceName_Click()
    Clipboard.Clear
    Clipboard.SetText LVServer.SelectedItem.Text, 1
End Sub

Private Sub sDllNature_Click()
    If Not LVServer.SelectedItem.SubItems(6) = "" Then
        ShowFileProperties LVServer.SelectedItem.SubItems(6)
    Else
        MsgBox "û���ҵ��ļ���", vbOKOnly + vbInformation, "����"
    End If
End Sub

Private Sub sExeNature_Click()
    If Not LVServer.SelectedItem.SubItems(3) = "" Then
        ShowFileProperties LVServer.SelectedItem.SubItems(3)
    Else
        MsgBox "û���ҵ��ļ���", vbOKOnly + vbInformation, "����"
    End If
End Sub

Private Sub sMoreInformation_Click()
    ShellExecute Me.hWnd, "open", "http://www.baidu.com/s?wd=" & (LVServer.SelectedItem.Text), vbNullString, vbNullString, SW_SHOW
End Sub

Private Sub sNew_Click()
    Call msNew_Click
End Sub

Private Sub sRecover_Click()
    RecoverSSDTSingle Val(LVSSDT.SelectedItem.Text)
End Sub

Private Sub sSelectDll_Click()
    If Not LVServer.SelectedItem.SubItems(6) = "" Then
        FindFiles LVServer.SelectedItem.SubItems(6)
    Else
        MsgBox "û���ҵ��ļ���", vbOKOnly + vbInformation, "����"
    End If
End Sub

Private Sub sSelectExe_Click()
    If Not LVServer.SelectedItem.SubItems(3) = "" Then
        FindFiles LVServer.SelectedItem.SubItems(3)
    Else
        MsgBox "û���ҵ��ļ���", vbOKOnly + vbInformation, "����"
    End If
End Sub

Private Sub Text1_Change()
    If SetText Then Exit Sub
    If mWindowFilterMethod And MethodSearch Then
        Call CNNew
    ElseIf Text1 <> "" Then
        mWindowFilterMethod = mWindowFilterMethod Or MethodSearch
        Call CNNew
    End If
    'If ListView1.Tag = 0 Then
    '    Call CNNew
    'ElseIf ListView1.Tag >= 1 Then
    '    EnumAllChildWindows nSelectedItem(ListView1.Tag), Text1.Text
    'End If
End Sub

Private Sub Text1_GotFocus()
    If FirstFocus = True Then
        Text1.Text = ""
        FirstFocus = False
    End If
End Sub

Public Function SetVisual(ByRef Visuals() As String, ByRef Soft() As String) '�������
    On Error GoTo 0
    Static SkinH_VB6 As Boolean
    
    If SkinH_VB6 Then Exit Function
    
    Dim i    As Long ', ok As Boolean
    Dim temp As String
    ReadINI "Visual settings", "Skin", temp
    Menu.chk.Value = Val(temp)

    If Menu.chk.Value = 0 Then

        For i = 0 To Me.count - 1

            If InStr(Me.Controls(i).Name, "Slider") > 0 Then Me.Controls(i).Enabled = False
        Next

        Exit Function
    End If
    
    If CheckFor("SkinH_VB6.dll", "") = False Then
        MsgBox "û���ҵ�SkinH_VB6.dll���޷�����Ƥ����", vbCritical
        SkinH_VB6 = True
        Exit Function
    End If

    SkinH_Attach  'skin
    SkinH_SetAero 1 'skin

    For i = 0 To 2

        If IsNumeric(Visuals(i)) Then
            'ok = False
            'Else:
            'ok = True
            Menu.Controls("Slider" & i + 1).Value = Val(Visuals(i))
        End If

        If IsNumeric(Soft(i)) Then Controls("check" & i + 1).Value = Val(Soft(i))
    Next

    SkinH_AdjustHSV Menu.Slider1.Value, Menu.Slider2.Value, Menu.Slider3.Value

    For i = 3 To 9

        If IsNumeric(Visuals(i)) Then
            'ok = False
            'Else
            'ok = True
            Menu.Controls("Slider" & i + 1).Value = Val(Visuals(i))
        End If

    Next

    SkinH_AdjustAero Menu.Slider4.Value, Menu.Slider7.Value, Menu.Slider6.Value, Menu.Slider5.Value, 0, 0, Menu.Slider8.Value, Menu.Slider9.Value, Menu.Slider10.Value

    If IsNumeric(Visuals(10)) Then Menu.Slider11.Value = Visuals(10)
    SkinH_SetMenuAlpha Menu.Slider11.Value
End Function

Private Sub ServiceStart(ComputerName As String, ServiceName As String) 'start server
    Dim ServiceStatus As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim res As Long

    hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
        If hService <> 0 Then
            res = StartService(hService, 0, 0)
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If
End Sub
Private Sub ServiceStop(ComputerName As String, ServiceName As String) 'stop
    Dim ServiceStatus As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim res As Long

    hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
        If hService <> 0 Then
            res = ControlService(hService, SERVICE_CONTROL_STOP, ServiceStatus)
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If
End Sub
Private Sub ServicePause(ComputerName As String, ServiceName As String) 'pause
    Dim ServiceStatus As SERVICE_STATUS
    Dim hSManager As Long
    Dim hService As Long
    Dim res As Long

    hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)
    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)
        If hService <> 0 Then
            res = ControlService(hService, SERVICE_CONTROL_PAUSE, ServiceStatus)
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSManager
    End If
End Sub

Private Sub SetServerBootType(ByVal SubText As String, BootType As Long)
    Dim Reg As clsRegistry
    
    Set Reg = New clsRegistry
    Reg.DeleteValue eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\" & SubText, "Start"

    If Not Reg.SetValue(eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\" & SubText, "Start", BootType) Then
        MsgBox "�����޸�ʧ�ܡ�", vbInformation, "��ʾ"
    Else
        sNew_Click
    End If
End Sub

Private Sub tvwKeys_Expand(ByVal Node As MSComctlLib.Node)
    Dim S As String
    Dim Y As Node
    S = Node.FullPath
    If Node.Tag <> 1 Then
        tvwKeys.Nodes.Remove Node.Child.Index
        EnumReg Node, True
        Node.Tag = 1
    End If
    Call EnumValue(Node)
End Sub

Private Sub tvwKeys_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button <> 2 Then Exit Sub
    rDelete.Enabled = (UBound(Split(tvwKeys.SelectedItem.Key, "\")) > 1)
    rRename.Enabled = rDelete.Enabled
    PopupMenu mnuReg
End Sub

Private Sub tvwKeys_NodeCheck(ByVal Node As MSComctlLib.Node)
    MsgBox Node
End Sub

Private Sub tvwKeys_NodeClick(ByVal Node As MSComctlLib.Node)
    Call EnumValue(Node)
End Sub

Private Sub ListSSDT()
    With LVSSDT.ListItems
        .Clear
        Dim i As Long
        For i = 0 To UBound(SSDTData)
            With .Add(, , i)
                Dim j As Long
                j = 0
                If SSDTData(i).dwRealAddress <> SSDTData(i).dwCurrAddress Then j = vbRed
                .ForeColor = j
                With .ListSubItems.Add
                    .Text = SSDTData(i).strName
                    .ForeColor = j
                End With
                With .ListSubItems.Add
                    .Text = FormatHex(SSDTData(i).dwRealAddress)
                    .ForeColor = j
                End With
                With .ListSubItems.Add
                    .Text = FormatHex(SSDTData(i).dwCurrAddress)
                    .ForeColor = j
                End With
                With .ListSubItems.Add
                    .Text = AddrToModuleName(SSDTData(i).dwCurrAddress)
                    .ForeColor = j
                End With
            End With
        Next
    End With
End Sub

Private Sub ListShadowSSDT()
    With LVShadowSSDT.ListItems
        .Clear
        Dim i As Long
        For i = 0 To UBound(ShadowSSDTData)
            With .Add(, , i)
                Dim j As Long
                j = 0
                If ShadowSSDTData(i).dwRealAddress <> ShadowSSDTData(i).dwCurrAddress Then j = vbRed
                .ForeColor = j
                With .ListSubItems.Add
                    .Text = ShadowSSDTData(i).strName
                    .ForeColor = j
                End With
                With .ListSubItems.Add
                    .Text = FormatHex(ShadowSSDTData(i).dwRealAddress)
                    .ForeColor = j
                End With
                With .ListSubItems.Add
                    .Text = FormatHex(ShadowSSDTData(i).dwCurrAddress)
                    .ForeColor = j
                End With
                With .ListSubItems.Add
                    .Text = AddrToModuleName(ShadowSSDTData(i).dwCurrAddress)
                    .ForeColor = j
                End With
            End With
        Next
    End With
End Sub

'<Setting>
Private Sub chk_Click()
    lblLabel6.Caption = "��ǰ���ã�" & IIf(chk.Value = 0, "����Ƥ��", "����Ƥ��") & "���������ú�������Ч"
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
    
    Dim i As Long, j As Boolean, k As Long
    For i = 1 To 14
        j = pColumnSelect(i).Checked
        S = S + WriteINI("Process Column", ProcessColumnNames(i), CInt(j))
        If j Then
            S = S + WriteINI("Process Column Width", RealProcessColumnNames(i), Round(ListView2.ColumnHeaders(i - k + 1).Width))
        Else
            k = k + 1
        End If
    Next
    
    If Not CBool(chk.Value) Then SkinH_Detach
    
    Dim TempStr       As String
    Dim VisualTitle() As String '�Ӿ�����
    Dim SoftSetting() As String '�������
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
    
    If 38 = S Then MsgBox "����ɹ�" Else MsgBox "����" & 38 - S & "���ʧ��"
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

VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "msComCtl32.OCX"
Begin VB.Form HeapList 
   Caption         =   "进程堆查看"
   ClientHeight    =   5430
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10365
   BeginProperty Font 
      Name            =   "宋体"
      Size            =   9
      Charset         =   134
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   5430
   ScaleWidth      =   10365
   StartUpPosition =   2  '屏幕中心
   Begin MSComctlLib.ListView ListView1 
      Height          =   5415
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   10335
      _ExtentX        =   18230
      _ExtentY        =   9551
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "宋体"
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
Attribute VB_Name = "HeapList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim mPid As Long


Private Sub Form_Load()
    ListView1.ColumnHeaders.Add , , "堆ID", 1500
    ListView1.ColumnHeaders.Add , , "堆大小", 1300
    ListView1.ColumnHeaders.Add , , "堆地址", 4200
    ListView1.ColumnHeaders.Add , , "块大小", 1200

    ListView1.Tag = 0
    
    mPid = nsItem
    
    'ListViewColor Me, ListView1
    'SetTextColor Me
    
    Call HNNew(mPid)
End Sub

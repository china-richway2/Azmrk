VERSION 5.00
Begin VB.Form DGEditValue 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "修改值"
   ClientHeight    =   4005
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5685
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   4005
   ScaleWidth      =   5685
   StartUpPosition =   3  '窗口缺省
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "取消"
      Height          =   495
      Left            =   4320
      TabIndex        =   10
      Top             =   3480
      Width           =   1215
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "确定"
      Height          =   495
      Left            =   4320
      TabIndex        =   9
      Top             =   2520
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      Caption         =   "Base64值"
      Height          =   1455
      Left            =   120
      TabIndex        =   7
      Top             =   2520
      Width           =   4095
      Begin VB.TextBox txtBase64 
         Height          =   1095
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   8
         Top             =   240
         Width           =   3855
      End
   End
   Begin VB.Frame Frame2 
      Caption         =   "Unicode"
      Height          =   975
      Left            =   2760
      TabIndex        =   4
      Top             =   1440
      Width           =   2895
      Begin VB.TextBox txtUnicode 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   6
         Top             =   240
         Width           =   2655
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "ASCII"
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   1440
      Width           =   2535
      Begin VB.TextBox txtASCII 
         Height          =   615
         Left            =   120
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   5
         Top             =   240
         Width           =   2295
      End
   End
   Begin VB.TextBox txtHex 
      Height          =   735
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   1
      Top             =   360
      Width           =   5535
   End
   Begin VB.Label Label2 
      Caption         =   "字符串："
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   1200
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "十六进制值："
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
End
Attribute VB_Name = "DGEditValue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Key As Long, dwClasss As Long, szValName As String, dData() As Byte, pType As Boolean

Private Sub SetString(ByRef szString() As Byte, ByVal iType As Long)
    Dim j() As Byte
    pType = False
    dData = szString
    If iType <> 0 Then txtASCII.Text = szString
    If iType <> 1 Then txtUnicode.Text = StrConv(szString, vbUnicode)
    If iType <> 2 Then
        Dim s As String
        s = Space(UBound(szString) * 2 + 2)
        Dim i As Long
        For i = 1 To Len(s) Step 2
            Mid(s, i, 2) = right("0" & Hex(szString(i \ 2)), 2)
        Next
        txtHex.Text = s
    End If
    If iType <> 3 Then
        j = szString
        Call Base64Array_Encode(j)
        txtBase64.Text = StrConv(j, vbUnicode)
    End If
    pType = True
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Sub Init(ByVal szClassName As String, ByVal hKey As Long, ByVal szValueName As String)
    Caption = "修改 " & szClassName & " 值"
    Key = hKey
    szValName = szValueName
    Dim dDataLen As Long
    dDataLen = 256
    ReDim dData(255)
    Dim Ret As Long
    Ret = RegQueryValueEx(Key, szValName, 0, dwClasss, dData(0), dDataLen)
    If Ret = 234 Then
        ReDim dData(dDataLen - 1)
        Ret = RegQueryValueEx(Key, szValName, 0, dwClasss, dData(0), dDataLen)
    End If
    If Ret <> 0 Then
        MsgBox "注册表未知错误！错误号：" & Ret, vbCritical
        Unload Me
    End If
End Sub

Private Sub cmdOK_Click()
    If RegSetValueEx(Key, szValName, 0, dwClasss, dData(0), UBound(dData) + 1) <> ERROR_SUCCESS Then
        MsgBox "失败！", vbCritical
    Else
        Unload Me
    End If
End Sub

Private Sub Form_Load()
    pType = True
End Sub

Private Sub Form_Terminate()
    RegCloseKey Key
End Sub

Private Sub Text1_KeyPress(KeyAscii As Integer)
    If KeyAscii = 8 Then Exit Sub
    If InStr("0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz+/", Chr(KeyAscii)) <= 0 Then
        KeyAscii = 0
    End If
End Sub

Private Sub txtASCII_Change()
    If pType Then SetString txtASCII.Text, 0
End Sub

Private Sub txtBase64_Change()
    If Not pType Then Exit Sub
    On Error Resume Next
    Dim s() As Byte
    s = StrConv(txtBase64.Text, vbFromUnicode)
    Call Base64Array_Decode(s)
    SetString s, -1
    'SetString s, 4
End Sub

Private Sub txtHex_Change()
    If Not pType Then Exit Sub
    Dim d() As Byte
    Dim i As Long
    If Len(txtHex.Text) <= 1 Then Exit Sub
    ReDim d(1 To Len(txtHex.Text) \ 2)
    For i = 1 To Len(txtHex.Text) Step 2
        d(i \ 2 + 1) = Val("&H" & Mid(txtHex.Text, i, 2))
    Next
    SetString d, -1
    'SetString d, 2
End Sub

Private Sub txtUnicode_Change()
    If pType Then SetString StrConv(txtUnicode.Text, vbFromUnicode), 1
End Sub

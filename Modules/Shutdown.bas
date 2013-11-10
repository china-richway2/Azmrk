Attribute VB_Name = "Shutdown"
Public Const EWX_LOGOFF As Long = &H0
Public Const EWX_SHUTDOWN As Long = &H1
Public Const EWX_REBOOT As Long = &H2
Public Const EWX_FORCE As Long = &H4
Public Const EWX_POWEROFF As Long = &H8
Public Const EWX_FORCEIFHUNG As Long = &H10
Public Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long
Public Declare Function ExitWindows Lib "user32" (ByVal dwReserved As Long, ByVal uReturnCode As Long) As Long


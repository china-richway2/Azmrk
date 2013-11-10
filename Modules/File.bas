Attribute VB_Name = "File"
Option Explicit
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long


Public Const OFS_MAXPATHNAME = 128
Public Const OF_CREATE = &H1000
Public Const OF_READ = &H0
Public Const OF_WRITE = &H1

Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1

Public Const MAX_PATH = 260


Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

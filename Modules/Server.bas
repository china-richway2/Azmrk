Attribute VB_Name = "Server"
Public Enum result
    SUCCESS
    PASSWORD_INCORRECT
    PROCESS_CPU_MEM_LIST
    PROCESS_TERMINATED
    PROCESS_CREATED
    
    WINDOW_DESTROY
    WINODW_CREATE
    WINDOW_FIRSTEDIT
    WINDOW_EDITDMC
End Enum

Public Enum Cmd
    CHECK_SERVER
    
    PROCESS_REFRESH
    CMD_PROCESS_TERMINATE
    PROCESS_SUSPEND
    PROCESS_RESUME
    PROCESS_SET_PRIORITY
    
    THREAD_REFRESH
    CMD_THREAD_TERMINATE
    THREAD_SUSPEND
    THREAD_RESUME
    THREAD_SET_PRIORITY
    
    WINDOW_REFRESH
    WINDOW_FROMPOINT
    WINDOW_UPDATE
    WINDOW_CLOSE
End Enum
Dim Buffer() As Byte, i As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)

Private Function AddSize(ByVal nSize As Long) As Long
    On Error GoTo Init
    ReDim Preserve Buffer(i + nSize)
    AddSize = i
    i = i + nSize
    Exit Function
Init:
    ReDim Buffer(nSize - 1)
    i = -1
    AddSize = 0
End Function

Public Sub Send()
    MainWindow.wsk.SendData Buffer
    i = -1
    Erase Buffer()
End Sub

Public Sub SendByte(ByVal nByte As Byte)
    Buffer(AddSize(1)) = nByte
End Sub

Public Sub SendInt(ByVal nInt As Integer)
    CopyMemory Buffer(AddSize(2)), nInt, 2
End Sub

Public Sub SendLng(ByVal nLng As Long)
    CopyMemory Buffer(AddSize(4)), nLng, 4
End Sub

Public Sub SendStr(ByVal sStr As String)
    Dim nStr As String
    nStr = StrConv(sStr, vbFromUnicode)
    CopyMemory Buffer(AddSize(LenB(nStr))), ByVal StrPtr(nStr), LenB(nStr)
    Buffer(AddSize(1)) = 0
End Sub

Public Sub Read(ByRef nBuffer() As Byte)
    Erase Buffer()
    nBuffer = Buffer
    i = -1
End Sub

Public Function ReadByte() As Byte
    ReadByte = Buffer(AddSize(1))
End Function

Public Function ReadInt() As Integer
    CopyMemory ReadInt, Buffer(AddSize(2)), 2
End Function

Public Function ReadLng() As Long
    CopyMemory ReadLng, Buffer(AddSize(4)), 4
End Function

Public Function ReadStr() As String
    Dim bBuf() As Byte, j As Long
    j = i
    Do Until Buffer(AddSize(1)) = 0
    Loop
    ReDim bBuf(i - j - 1)
    CopyMemory bBuf(0), Buffer(j), i - j
    ReadStr = StrConv(bBuf, vbUnicode)
End Function

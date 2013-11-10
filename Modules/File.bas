Attribute VB_Name = "File"
Option Explicit
Public Declare Function CreateFile Lib "kernel32" Alias "CreateFileA" (ByVal lpFileName As String, ByVal dwDesiredAccess As Long, ByVal dwShareMode As Long, lpSecurityAttributes As SECURITY_ATTRIBUTES, ByVal dwCreationDisposition As Long, ByVal dwFlagsAndAttributes As Long, ByVal hTemplateFile As Long) As Long
Public Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, lpReOpenBuff As OFSTRUCT, ByVal wStyle As Long) As Long
Public Declare Function ReadFile Lib "kernel32" (ByVal hFile As Long, lpBuffer As Any, ByVal nNumberOfBytesToRead As Long, lpNumberOfBytesRead As Long, ByVal lpOverlapped As Long) As Long
Public Declare Function GetFileSize Lib "kernel32" (ByVal hFile As Long, lpFileSizeHigh As Long) As Long


Public Const OFS_MAXPATHNAME = 128
Public Const OF_CREATE = &H1000
Public Const OF_READ = &H0
Public Const OF_WRITE = &H1

Public Const FILE_BEGIN = 0
Public Const FILE_CURRENT = 1
Public Const FILE_END = 2

Public Const FILE_ATTRIBUTE_HIDDEN = &H2
Public Const FILE_ATTRIBUTE_SYSTEM = &H4
Public Const FILE_ATTRIBUTE_DIRECTORY = &H10
Public Const FILE_ATTRIBUTE_ARCHIVE = &H20
Public Const FILE_ATTRIBUTE_NORMAL = &H80
Public Const FILE_ATTRIBUTE_TEMPORARY = &H100
Public Const FILE_ATTRIBUTE_COMPRESSED = &H800

Public Const FILE_TYPE_UNKNOWN = &H0
Public Const FILE_TYPE_DISK = &H1
Public Const FILE_TYPE_CHAR = &H2
Public Const FILE_TYPE_PIPE = &H3
Public Const FILE_TYPE_REMOTE = &H8000

Public Const FILE_READ_DATA = &H1
Public Const FILE_WRITE_DATA = &H2
Public Const FILE_APPEND_DATA = &H4
Public Const FILE_READ_EA = &H8
Public Const FILE_WRITE_EA = &H10
Public Const FILE_EXECUTE = &H20
Public Const FILE_DELETE_CHILD = &H40
Public Const FILE_READ_ATTRIBUTES = &H80
Public Const FILE_WRITE_ATTRIBUTES = &H100
Public Const FILE_READ_PROPERTIES = FILE_READ_EA
Public Const FILE_WRITE_PROPERTIES = FILE_WRITE_EA

Public Const READ_CONTROL = &H20000
Public Const STANDARD_RIGHTS_READ = (READ_CONTROL)
Public Const STANDARD_RIGHTS_EXECUTE = (READ_CONTROL)
Public Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)

Public Const FILE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SYNCHRONIZE Or &H1FF)
Public Const FILE_GENERIC_WRITE = (STANDARD_RIGHTS_WRITE Or FILE_WRITE_DATA Or FILE_WRITE_ATTRIBUTES Or FILE_WRITE_EA Or FILE_APPEND_DATA Or SYNCHRONIZE)
Public Const FILE_GENERIC_EXECUTE = (STANDARD_RIGHTS_EXECUTE Or FILE_READ_ATTRIBUTES Or FILE_EXECUTE Or SYNCHRONIZE)
Public Const FILE_GENERIC_READ = (STANDARD_RIGHTS_READ Or FILE_READ_DATA Or FILE_READ_ATTRIBUTES Or FILE_READ_EA Or SYNCHRONIZE)

Public Const FILE_SHARE_READ = &H1
Public Const FILE_SHARE_WRITE = &H2
Public Const FILE_UNICODE_ON_DISK = &H4
Public Const FILE_TRAVERSE = &H20
Public Const FILE_VOLUME_IS_COMPRESSED = &H8000

Public Const FILE_CASE_SENSITIVE_SEARCH = &H1
Public Const FILE_ADD_FILE = &H2
Public Const FILE_ADD_SUBDIRECTORY = &H4
Public Const FILE_CASE_PRESERVED_NAMES = &H2
Public Const FILE_CREATE_PIPE_INSTANCE = (&H4)
Public Const FILE_FILE_COMPRESSION = &H10
Public Const FILE_FLAG_BACKUP_SEMANTICS = &H2000000
Public Const FILE_FLAG_DELETE_ON_CLOSE = &H4000000
Public Const FILE_FLAG_NO_BUFFERING = &H20000000
Public Const FILE_FLAG_OVERLAPPED = &H40000000
Public Const FILE_FLAG_POSIX_SEMANTICS = &H1000000
Public Const FILE_FLAG_RANDOM_ACCESS = &H10000000
Public Const FILE_FLAG_SEQUENTIAL_SCAN = &H8000000
Public Const FILE_FLAG_WRITE_THROUGH = &H80000000
Public Const FILE_LIST_DIRECTORY = (&H1)

Public Const FILE_NOTIFY_CHANGE_FILE_NAME = &H1
Public Const FILE_NOTIFY_CHANGE_DIR_NAME = &H2
Public Const FILE_NOTIFY_CHANGE_ATTRIBUTES = &H4
Public Const FILE_NOTIFY_CHANGE_SIZE = &H8
Public Const FILE_NOTIFY_CHANGE_LAST_WRITE = &H10
Public Const FILE_NOTIFY_CHANGE_SECURITY = &H100
Public Const FILE_PERSISTENT_ACLS = &H8

Public Const MAX_PATH = 260

Public Type OFSTRUCT
    cBytes As Byte
    fFixedDisk As Byte
    nErrCode As Integer
    Reserved1 As Integer
    Reserved2 As Integer
    szPathName(OFS_MAXPATHNAME) As Byte
End Type

Attribute VB_Name = "Driver"
Option Explicit

Private Declare Function EnumDeviceDrivers Lib "psapi.dll" (ByRef lpImageBase As Long, ByVal cb As Long, ByRef lpcbNeeded As Long) As Long
Private Declare Function GetDeviceDriverBaseName Lib "psapi.dll" Alias "GetDeviceDriverBaseNameA" (ByVal ImageBase As Long, ByVal lpBaseName As String, ByVal nSize As Long) As Long
Private Declare Function GetDeviceDriverFileName Lib "psapi.dll" Alias "GetDeviceDriverFileNameA" (ByVal ImageBase As Long, ByVal lpFileName As String, ByVal nSize As Long) As Long


Public Enum enmDeviceDriver
    BaseAddress = 0
    BaseName = 1
End Enum


'http://msdn.microsoft.com/en-us/library/ms682617(VS.85).aspx
Public Function GetDeviceDriver(ByVal nWhat As enmDeviceDriver, Optional ByVal currentDriver As Long = 0)
    'Code By zhouweizhu@126.com
    Dim loadAddresses() As Long, bytesNeeded As Long, driverCount As Integer, outputBuffer As String  ', currentDriver As Integer
    Dim nLength As Long
    Let outputBuffer = Space$(255)
    ReDim loadAddresses(0)
    If EnumDeviceDrivers(loadAddresses(0), 4 * (UBound(loadAddresses) + 1), bytesNeeded) <> 0 Then
        Let driverCount = bytesNeeded / 4
        If currentDriver < driverCount Then
            ReDim loadAddresses(driverCount - 1)
            Call EnumDeviceDrivers(loadAddresses(0), 4 * (UBound(loadAddresses) + 1), bytesNeeded)
            If nWhat = BaseAddress Then
                Let GetDeviceDriver = loadAddresses(currentDriver)
            ElseIf nWhat = BaseName Then
                Let nLength = GetDeviceDriverBaseName(loadAddresses(currentDriver), outputBuffer, Len(outputBuffer))
                Let GetDeviceDriver = Mid$(outputBuffer, 1, nLength)
            End If
        End If
    End If
End Function

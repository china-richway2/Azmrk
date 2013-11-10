Attribute VB_Name = "LocalThread"
Option Explicit


Public Type FEWBT_PARAM
    min As Long
    max As Long
End Type

Public Function FEWBT_MultiThreading() As Long()
    Dim t_min, t_max As Long
    Dim hThread() As Long
    Dim i As Integer
    Dim pa As FEWBT_PARAM
    'all_max=9999999  '7λ,1*10^8-1
    
    ReDim hThread(9)
    
    With pa
        .min = 0
        .max = 999999
    End With
    
    LockWindowUpdate Menu.ListView1.hwnd
    
    For i = 0 To 9
        hThread(i) = CreateThread(0, 0, AddressOf FxEnumWindowsByThread, VarPtr(pa.min & "," & pa.max), 0, 0)
        With pa
            .min = .max + 1
            .max = .min + 999999
        End With
    Next i
    
    If WaitForMultipleObjects(10, hThread(0), 0, INFINITE) Then
        LockWindowUpdate 0
        Menu.ListView1.Refresh
        
        FEWBT_MultiThreading = hThread()
    End If
End Function

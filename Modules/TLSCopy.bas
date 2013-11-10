Attribute VB_Name = "TLSCopy"
Option Explicit

Private Declare Function CreateThread Lib "kernel32.dll" ( _
    ByVal lpsa As Long, _
    ByVal cbStack As Long, _
    ByVal lpStartAddr As Long, _
    ByVal lpvThreadParam As Long, _
    ByVal fdwCreate As Long, _
    ByVal lpIDThread As Long _
    ) As Long

Private Declare Sub ExitThread Lib "kernel32.dll" ( _
    ByVal dwExitCode As Long _
    )

Private Declare Function TerminateThread Lib "kernel32.dll" ( _
    ByVal hThread As Long, _
    ByVal dwExitCode As Long _
    ) As Long

Private Declare Function ResumeThread Lib "kernel32.dll" ( _
    ByVal hThread As Long _
    ) As Long

Private Declare Function GetModuleHandleW Lib "kernel32.dll" ( _
    ByVal lpModuleName As Long _
    ) As Long

Private Declare Function GetProcAddress Lib "kernel32.dll" ( _
    ByVal hModule As Long, _
    ByVal lpProcName As Long _
    ) As Long

Private Declare Sub RtlMoveMemory Lib "kernel32.dll" ( _
    ByVal Destination As Long, _
    ByVal Source As Long, _
    ByVal Length As Long _
    )

Private Declare Function VirtualProtect Lib "kernel32.dll" ( _
    ByVal lpAddress As Long, _
    ByVal dwSize As Long, _
    ByVal flNewProtect As Long, _
    ByVal lpflOldProtect As Long _
    ) As Long

Private Declare Function CallWindowProcW Lib "user32.dll" ( _
    ByVal lpPrevWndFunc As Long, _
    ByVal hWnd As Long, _
    ByVal Msg As Long, _
    ByVal wParam As Long, _
    ByVal lParam As Long _
    ) As Long

Private Declare Function MessageBoxW Lib "user32.dll" ( _
    ByVal hWnd As Long, _
    ByVal lpText As Long, _
    ByVal lpCaption As Long, _
    ByVal uType As Long _
    ) As Long
    
Private Declare Function GetProcessHeap Lib "kernel32.dll" ( _
    ) As Long

Private Declare Function HeapAlloc Lib "kernel32.dll" ( _
    ByVal hHeap As Long, _
    ByVal dwFlags As Long, _
    ByVal dwBytes As Long _
    ) As Long

Private Declare Function HeapFree Lib "kernel32.dll" ( _
    ByVal hHeap As Long, _
    ByVal dwFlags As Long, _
    ByVal lpMem As Long _
    ) As Long

Private Declare Function CloseHandle Lib "kernel32.dll" ( _
    ByVal hObject As Long _
    ) As Long

Private Declare Function InterlockedDecrement Lib "kernel32.dll" ( _
    ByVal lpAddend As Long _
    ) As Long

Private Declare Function SetPixel Lib "gdi32.dll" ( _
    ByVal hdc As Long, _
    ByVal X As Long, _
    ByVal Y As Long, _
    ByVal crColor As Long _
    ) As Long

Private Declare Sub Sleep Lib "kernel32.dll" ( _
    ByVal dwMilliseconds As Long _
    )

Private Const PAGE_EXECUTE_READWRITE As Long = &H40&
Private Const CREATE_SUSPENDED As Long = &H4&
Private Const STILL_ACTIVE As Long = &H103&

Dim initialized As Long
Dim teb_address As Long
Dim base_kernel32 As Long
Dim base_terminate_thread As Long
Dim base_close_handle As Long
Dim thread_delegate(0 To 14) As Long

' description: call a specified address
' arguments:
'   address - specify the calling address
'   size - specify the code size
'   param_0 - the first parameter
'   param_1 - the second parameter
'   param_2 - the third parameter
'   param_3 - the fourth parameter
' return value:
'   the value returned by the calling function
Private Function tiny_call_addr( _
    ByVal address As Long, _
    ByVal size As Long, _
    ByVal param_0 As Long, _
    ByVal param_1 As Long, _
    ByVal param_2 As Long, _
    ByVal param_3 As Long _
    ) As Long

    Dim result As Long
    Dim old_protect As Long
    
    ' make the page executable
    result = VirtualProtect(address, size, PAGE_EXECUTE_READWRITE, VarPtr(old_protect))
    
    If result <> 0 Then
        ' use CallWindowProcW to call our address
        result = CallWindowProcW(address, param_0, param_1, param_2, param_3)
        ' restore the page protection
        VirtualProtect address, size, old_protect, VarPtr(old_protect)
    End If
    
    tiny_call_addr = result
End Function

' description: dereference a delegate block
' arguments:
'   delegate_block_ptr - a pointer to the delegate block
' return value:
'   if the thread is terminated, the return value is the exit code
'   if the thread is still active, the return value is STILL_ACTIVE
Public Function dereference_delegate_block(ByVal delegate_block_ptr As Long) As Long
    Dim delegate_buffer(0 To 7) As Long
    Dim heap_handle As Long
    RtlMoveMemory VarPtr(delegate_buffer(0)), delegate_block_ptr, 32
    If InterlockedDecrement(delegate_block_ptr + &HC) = 0 Then
        CloseHandle delegate_buffer(2)
        heap_handle = GetProcessHeap()
        HeapFree heap_handle, 0, delegate_block_ptr
    End If
    dereference_delegate_block = delegate_buffer(4)
End Function

' description: create a thread with delegation and tls copied
' arguments:
'   start_address - represents the starting address of the thread
'   argument_list - argument list to be passed to a new thread or NULL
'   argument_count - count of arguments in the argument list, can be zero
'   stack_size - stack size for a new thread or zero
Public Function begin_thread( _
    ByVal start_address As Long, _
    Optional ByVal argument_list As Long = 0, _
    Optional ByVal argument_count As Long = 0, _
    Optional ByVal stack_size As Long = 0 _
    ) As Long
    
    Dim retrieve_teb(0 To 2) As Long
    Dim old_protect As Long
    Dim heap_handle As Long
    Dim delegate_buffer(0 To 7) As Long
    Dim delegate_block_ptr As Long
    Dim thread_handle As Long
    Dim result As Long

    result = 0

    ' initialize data in the first time
    If initialized = 0 Then
        retrieve_teb(0) = &HA164FF8B    ' mov edi, edi
        retrieve_teb(1) = &H18&         ' mov eax, dword ptr fs:[18]
        retrieve_teb(2) = &H10C2&       ' retn 0x10
        
        teb_address = tiny_call_addr(VarPtr(retrieve_teb(0)), _
            12, 0, 0, 0, 0)
        
        If teb_address <> 0 Then
            
            base_kernel32 = GetModuleHandleW(StrPtr("kernel32.dll"))
            base_terminate_thread = GetProcAddress(base_kernel32, _
                StrPtr(StrConv("TerminateThread", vbFromUnicode)))
            base_close_handle = GetProcAddress(base_kernel32, _
                StrPtr(StrConv("CloseHandle", vbFromUnicode)))
            
            If base_terminate_thread <> 0 Then
                thread_delegate(0) = &H448BFF8B     ' mov edi, edi
                                                    ' mov eax, [esp+0x4]
                thread_delegate(1) = &H308B0424     ' mov esi, [eax]
                thread_delegate(2) = &H183D8B64     ' mov edi, fs:[0x18]
                thread_delegate(3) = &H81000000     ' add edi, 0xe10
                thread_delegate(4) = &HE10C7
                thread_delegate(5) = &H59406A00     ' push 0x40
                                                    ' pop ecx
                thread_delegate(6) = &H488BA5F3     ' rep movsd
                                                    ' mov ecx, [eax+0x1c]
                thread_delegate(7) = &HC1D18B1C     ' mov edx, ecx
                                                    ' shl edx, 0x2
                thread_delegate(8) = &HE22B02E2     ' sub esp, edx
                thread_delegate(9) = &H8B20708D     ' lea esi, [eax+0x20]
                                                    ' mov edi, esp
                thread_delegate(10) = &HFFA5F3FC    ' rep movsd
                                                    ' call [eax+0x18]
                thread_delegate(11) = &H4C8B1850    ' mov ecx, [esp+0x4]
                thread_delegate(12) = &H41890424    ' mov [ecx+0x10], eax
                thread_delegate(13) = &H51FF5110    ' push ecx
                                                    ' call [ecx+0x4]
                thread_delegate(14) = &HCCCCCC04    ' int3 (never be here)
        
                If VirtualProtect(VarPtr(thread_delegate(0)), 15 * 4, _
                    PAGE_EXECUTE_READWRITE, VarPtr(old_protect)) <> 0 Then
                    initialized = 1
                End If
            End If
        End If
    End If
    
    If initialized <> 0 Then
        heap_handle = GetProcessHeap()
        If heap_handle <> 0 Then
            delegate_block_ptr = HeapAlloc(heap_handle, 0, 32 + argument_count * 4)
            If delegate_block_ptr <> 0 Then
                thread_handle = CreateThread(0, stack_size, VarPtr(thread_delegate(0)), _
                    delegate_block_ptr, CREATE_SUSPENDED, 0)
                If thread_handle <> 0 Then
                    delegate_buffer(0) = teb_address + &HE10    ' TlsSlots base
                    delegate_buffer(1) = return_long(AddressOf exit_thread_callback)
                    delegate_buffer(2) = thread_handle          ' handle
                    delegate_buffer(3) = 2                      ' reference count
                    delegate_buffer(4) = STILL_ACTIVE           ' exit code
                    delegate_buffer(5) = 0                      ' reserved
                    delegate_buffer(6) = start_address          ' start address
                    delegate_buffer(7) = argument_count         ' argument count
                    RtlMoveMemory delegate_block_ptr, VarPtr(delegate_buffer(0)), 32
                    If argument_count <> 0 Then
                        RtlMoveMemory delegate_block_ptr + 32, argument_list, argument_count * 4
                    End If
                    If ResumeThread(thread_handle) = -1 Then
                        TerminateThread thread_handle, -1
                        CloseHandle thread_handle
                        HeapFree heap_handle, 0, delegate_block_ptr
                    Else
                        result = delegate_block_ptr
                    End If
                End If
            End If
        End If
    End If
    
    begin_thread = thread_handle
End Function

Private Function return_long(ByVal long_ As Long) As Long
    return_long = long_
End Function

Private Sub exit_thread_callback(ByVal delegate_block_ptr As Long)
    Dim exit_code As Long
    exit_code = dereference_delegate_block(delegate_block_ptr)
    TerminateThread -2, exit_code
End Sub



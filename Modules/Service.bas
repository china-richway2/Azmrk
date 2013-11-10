Attribute VB_Name = "Service"
Option Explicit
Public Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long
Public Declare Function ControlService Lib "advapi32.dll" (ByVal hService As Long, ByVal dwControl As Long, lpServiceStatus As SERVICE_STATUS) As Long
Public Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Public Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Public Declare Function QueryServiceStatus Lib "advapi32.dll" (ByVal hService As Long, lpServiceStatus As SERVICE_STATUS) As Long
Public Declare Function StartService Lib "advapi32.dll" Alias "StartServiceA" (ByVal hService As Long, ByVal dwNumServiceArgs As Long, ByVal lpServiceArgVectors As Long) As Long
Public Declare Function GetServiceKeyName Lib "advapi32.dll" Alias "GetServiceKeyNameA" (ByVal hSCManager As Long, ByVal lpDisplayName As String, ByVal lpServiceName As String, lpcchBuffer As Long) As Long


'API Constants
Public Const SERVICES_ACTIVE_DATABASE = "ServicesActive"
'Service Control
Public Const SERVICE_CONTROL_STOP = &H1
Public Const SERVICE_CONTROL_PAUSE = &H2
'Service State - for CurrentState
Public Const SERVICE_STOPPED = &H1
Public Const SERVICE_START_PENDING = &H2
Public Const SERVICE_STOP_PENDING = &H3
Public Const SERVICE_RUNNING = &H4
Public Const SERVICE_CONTINUE_PENDING = &H5
Public Const SERVICE_PAUSE_PENDING = &H6
Public Const SERVICE_PAUSED = &H7
'Service Control Manager object specific access types
Public Const SC_MANAGER_CONNECT = &H1
Public Const SC_MANAGER_CREATE_SERVICE = &H2
Public Const SC_MANAGER_ENUMERATE_SERVICE = &H4
Public Const SC_MANAGER_LOCK = &H8
Public Const SC_MANAGER_QUERY_LOCK_STATUS = &H10
Public Const SC_MANAGER_MODIFY_BOOT_CONFIG = &H20
Public Const SC_MANAGER_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SC_MANAGER_CONNECT Or SC_MANAGER_CREATE_SERVICE Or SC_MANAGER_ENUMERATE_SERVICE Or SC_MANAGER_LOCK Or SC_MANAGER_QUERY_LOCK_STATUS Or SC_MANAGER_MODIFY_BOOT_CONFIG)
'Service object specific access types
Public Const SERVICE_QUERY_CONFIG = &H1
Public Const SERVICE_CHANGE_CONFIG = &H2
Public Const SERVICE_QUERY_STATUS = &H4
Public Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Public Const SERVICE_START = &H10
Public Const SERVICE_STOP = &H20
Public Const SERVICE_PAUSE_CONTINUE = &H40
Public Const SERVICE_INTERROGATE = &H80
Public Const SERVICE_USER_DEFINED_CONTROL = &H100
Public Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)


Public Type SERVICE_STATUS
    dwServiceType As Long
    dwCurrentState As Long
    dwControlsAccepted As Long
    dwWin32ExitCode As Long
    dwServiceSpecificExitCode As Long
    dwCheckPoint As Long
    dwWaitHint As Long
End Type


Public Sub msNew_Click()
    '    Dim i As Long
    '
    '    ListView3.ListItems.Clear
    '    i = 1
    '    Set objSWbemObjectSet = objSWbemServices.ExecQuery("SELECT * FROM " & strClass)  '通过WQL查询，返回指定类的所有
    '
    '    For Each objSWbemObject In objSWbemObjectSet
    '        ListView3.ListItems.Add , , objSWbemObject.DisplayName '将服务名称添加到ListView1第一列
    '        ListView3.ListItems(i).SubItems(1) = objSWbemObject.State '将服务的状态添加到ListView1第二列
    '        ListView3.ListItems(i).SubItems(2) = objSWbemObject.StartMode '将服务的启动方式添加到ListView1第三列
    '        ListView3.ListItems(i).SubItems(3) = objSWbemObject.PathName '将服务程序的路径添加到ListView1第四列
    '        ListView3.ListItems(i).SubItems(4) = objSWbemObject.StartName '将服务的登录身份添加到ListView1第五列
    '        ListView3.ListItems(i).SubItems(5) = objSWbemObject.ProcessID '将服务的进程ID添加到ListView1第六列
    '        i = i + 1
    '    Next
    '-----------------------------2010.02.21 08:28-----------------------
    Dim Registry As clsRegistry
    Set Registry = New clsRegistry
    Dim r_initial As String
    Dim rv_value As String
    Dim index_count As Long
    Dim num_count As Long
    Dim sIndex As Integer
    
    num_count = 0
    index_count = 0
    sIndex = 1

    If Menu.LVServer.ListItems.Count > 0 Then
        sIndex = Menu.LVServer.SelectedItem.Index
    End If
        
    DoEvents:
    Menu.LVServer.ListItems.Clear

    Do While r_initial <> " "
        r_initial = Registry.ListSubKey(eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services", index_count)

        If r_initial = "" Then
            Exit Sub
        End If
        rv_value = Registry.GetValue(eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\" & r_initial, "ObjectName")

        If rv_value <> "" Then
            num_count = num_count + 1
            'lstServices.AddItem r_initial 'num_count & ".) " &
            Menu.LVServer.ListItems.Add , , r_initial
            GetServerInfo r_initial, num_count
        End If
        
        index_count = index_count + 1
        'LVServer.ListItems(sIndex).Selected = True
        'LVServer.ListItems(sIndex).EnsureVisible
        Menu.Label5.Caption = "共有" & (num_count) & "个服务"
        'lblServiceCount = "当前 " & index_count & " 个服务"
        'If index_count > 15 Then Exit Sub
    Loop
    
    
    'lblServiceCount = index_count & " services"
    'r_initial = Registry.GetValue(eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\beep", "Group")

    'Set ListView1.SelectedItem = ListView1.ListItems(1)
    'ListView1_ItemClick ListView1.ListItems(1)
End Sub

Public Function GetServerInfo(ByVal ServerNames As String, ByVal CCount As Long)
    Dim Registry As clsRegistry
    Set Registry = New clsRegistry
    Dim r_initial   As String
    Dim rv_value    As String
    Dim serv_status As String
    Dim StartType   As String
    Dim StartPath   As String
    Dim LoginUser   As String
    Dim DllPath     As String
    r_initial = ServerNames

    If r_initial = "" Then Exit Function
    rv_value = Registry.GetValue(eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\" & r_initial, "Description")
    StartType = Registry.GetValue(eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\" & r_initial, "start")
    StartPath = Registry.GetValue(eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\" & r_initial, "imagepath")
    LoginUser = Registry.GetValue(eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\" & r_initial, "ObjectName")
    DllPath = Registry.GetValue(eHKEY_LOCAL_MACHINE, "System\currentcontrolset\services\" & r_initial & "\Parameters", "ServiceDll")

    Select Case StartType
    
        Case 2: StartType = "自动"

        Case 3: StartType = "手动"

        Case 4: StartType = "禁用"
    
    End Select

    Menu.LVServer.ListItems(CCount).SubItems(2) = StartType '启动类型
    Menu.LVServer.ListItems(CCount).SubItems(3) = SetupStartPath(Replace(StartPath, """", ""))
    Menu.LVServer.ListItems(CCount).SubItems(5) = LoginUser
    Menu.LVServer.ListItems(CCount).SubItems(6) = SetupStartPath(Replace(DllPath, """", ""))

    If rv_value = "" Then
        Menu.LVServer.ListItems(CCount).SubItems(4) = "无法获取描述."
        serv_status = ServiceStatus("", r_initial)
        Menu.LVServer.ListItems(CCount).SubItems(1) = serv_status
        Exit Function
    End If

    Menu.LVServer.ListItems(CCount).SubItems(4) = rv_value
    serv_status = ServiceStatus("", r_initial)
    Menu.LVServer.ListItems(CCount).SubItems(1) = serv_status
End Function

Public Function SetupStartPath(ByVal Path As String) As String
    Dim temp As String
    
    If Len(Path) = 0 Then Exit Function
    temp = left$(Path, InStr(Path, "\") - 1)

    If InStr(temp, "%") > 0 Then
        If LCase$(temp) = "%systemroot%" Then
            SetupStartPath = Environ("windir") & Mid(Path, InStr(Path, "\"))
        Else
            SetupStartPath = Environ(temp) & Mid(Path, InStr(Path, "\"))
        End If

    Else
        SetupStartPath = Path
    End If
End Function

Public Function ServiceStatus(ComputerName As String, ServiceName As String) As String
    Dim ServiceStat    As SERVICE_STATUS
    Dim hSManager      As Long
    Dim hService       As Long
    Dim hServiceStatus As Long

    ServiceStatus = ""
    hSManager = OpenSCManager(ComputerName, SERVICES_ACTIVE_DATABASE, SC_MANAGER_ALL_ACCESS)

    If hSManager <> 0 Then
        hService = OpenService(hSManager, ServiceName, SERVICE_ALL_ACCESS)

        If hService <> 0 Then
            hServiceStatus = QueryServiceStatus(hService, ServiceStat)

            If hServiceStatus <> 0 Then

                Select Case ServiceStat.dwCurrentState

                    Case SERVICE_STOPPED
                        ServiceStatus = "已停止"

                    Case SERVICE_START_PENDING
                        ServiceStatus = "开始"

                    Case SERVICE_STOP_PENDING
                        ServiceStatus = "停止"

                    Case SERVICE_RUNNING
                        ServiceStatus = "已启动"

                    Case SERVICE_CONTINUE_PENDING
                        ServiceStatus = "继续"

                    Case SERVICE_PAUSE_PENDING
                        ServiceStatus = "暂停"

                    Case SERVICE_PAUSED
                        ServiceStatus = "暂停"
                End Select

            End If

            CloseServiceHandle hService
        End If

        CloseServiceHandle hSManager
    End If

End Function

Attribute VB_Name = "modDestroy"
Option Explicit
Private Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegDeleteKey Lib "advapi32.dll" Alias "RegDeleteKeyA" (ByVal hKey As Long, ByVal lpSubKey As String) As Long
Private Declare Function RegDeleteValue Lib "advapi32.dll" Alias "RegDeleteValueA" (ByVal hKey As Long, ByVal lpValueName As String) As Long
Private Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Private Declare Function RegEnumKeyEx Lib "advapi32.dll" Alias "RegEnumKeyExA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpName As String, lpcbName As Long, ByVal lpReserved As Long, ByVal lpClass As String, lpcbClass As Long, lpftLastWriteTime As Any) As Long
Public Declare Function RegSetValueEx Lib "advapi32.dll" Alias "RegSetValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal Reserved As Long, ByVal dwType As Long, lpData As Any, ByVal cbData As Long) As Long

Private Declare Function SHDeleteKeyA Lib "Shlwapi.dll" (ByVal lRootKey As Long, ByVal szKeyToDelete As String) As Long

Private Declare Function SHFileOperation Lib "shell32.dll" Alias "SHFileOperationA" (lpFileOp As SHFILEOPSTRUCT) As Long

Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst32 Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext32 Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Sub CloseHandle Lib "kernel32" (ByVal hPass As Long)

Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)

Private Declare Function GetCurrentProcess Lib "kernel32" () As Long

Private Declare Function OpenSCManager Lib "advapi32.dll" Alias "OpenSCManagerA" (ByVal lpMachineName As String, ByVal lpDatabaseName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function OpenService Lib "advapi32.dll" Alias "OpenServiceA" (ByVal hSCManager As Long, ByVal lpServiceName As String, ByVal dwDesiredAccess As Long) As Long
Private Declare Function DeleteService Lib "advapi32.dll" (ByVal hService As Long) As Long
Private Declare Function CloseServiceHandle Lib "advapi32.dll" (ByVal hSCObject As Long) As Long

Private Type PROCESSENTRY32
    dwSize As Long
    cntUsage As Long
    th32ProcessID As Long
    th32DefaultHeapID As Long
    th32ModuleID As Long
    cntThreads As Long
    th32ParentProcessID As Long
    pcPriClassBase As Long
    dwFlags As Long
    szExeFile As String * 260
End Type

Private Type SHFILEOPSTRUCT
    hwnd As Long
    wFunc As Long
    pFrom As String
    pTo As String
    fFlags As Integer
    fAborted As Boolean
    hNameMaps As Long
    sProgress As String
End Type

Private Const SC_MANAGER_CREATE_SERVICE = &H2
Private Const SERVICE_QUERY_CONFIG = &H1
Private Const SERVICE_CHANGE_CONFIG = &H2
Private Const SERVICE_QUERY_STATUS = &H4
Private Const SERVICE_ENUMERATE_DEPENDENTS = &H8
Private Const SERVICE_START = &H10
Private Const SERVICE_STOP = &H20
Private Const SERVICE_PAUSE_CONTINUE = &H40
Private Const SERVICE_INTERROGATE = &H80
Private Const SERVICE_USER_DEFINED_CONTROL = &H100
Private Const STANDARD_RIGHTS_REQUIRED = &HF0000
Private Const SERVICE_ALL_ACCESS = (STANDARD_RIGHTS_REQUIRED Or SERVICE_QUERY_CONFIG Or SERVICE_CHANGE_CONFIG Or SERVICE_QUERY_STATUS Or SERVICE_ENUMERATE_DEPENDENTS Or SERVICE_START Or SERVICE_STOP Or SERVICE_PAUSE_CONTINUE Or SERVICE_INTERROGATE Or SERVICE_USER_DEFINED_CONTROL)

Private Const TH32CS_SNAPPROCESS As Long = 2&
Private Const WM_CLOSE = &H10
Private Const WM_QUIT = &H12

Private Const FO_DELETE = &H3
Private Const FOF_SILENT = &H4
Private Const FOF_NOCONFIRMATION = &H10
Private Const FOF_ALLOWUNDO = &H40

Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const KEY_ENUMERATE_SUB_KEYS = &H8
Private Const READ_CONTROL = &H20000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Private Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))

Private Const REG_BINARY = 3
Private Const REG_DWORD = 4
Private Const REG_SZ = 1

Public Sub KillRegKey(lHive&, sKey$)
    'KillRegSubKeys lHive, sKey
    
    SHDeleteKeyA lHive, sKey
End Sub

'Private Sub KillRegSubKeys(lHive&, sKey$)
'    Dim hKey&, i&, j&, sName$, sSubKeys$()
'    If RegOpenKeyEx(lHive, sKey, 0, KEY_ENUMERATE_SUB_KEYS, hKey) <> 0 Then Exit Sub
'    i = 0
'    sName = String(255, 0)
'    If RegEnumKeyEx(hKey, i, sName, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0 Then
'        RegCloseKey hKey
'        Exit Sub
'    End If
'    Do
'        sName = Left(sName, InStr(sName, Chr(0)) - 1)
'        ReDim Preserve sSubKeys(i)
'        sSubKeys(i) = sName
'
'        i = i + 1
'        sName = String(255, 0)
'    Loop Until RegEnumKeyEx(hKey, i, sName, 255, 0, vbNullString, ByVal 0, ByVal 0) <> 0
'    RegCloseKey hKey
'
'    For j = 0 To UBound(sSubKeys)
'        KillRegSubKeys lHive, sKey & "\" & sSubKeys(j)
'        RegDeleteKey lHive, sKey & "\" & sSubKeys(j)
'    Next j
'End Sub

Public Sub KillRegVal(lHive&, sKey$, sVal$)
    Dim hKey&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_WRITE, hKey) <> 0 Then Exit Sub
    RegDeleteValue hKey, sVal
    RegCloseKey hKey
End Sub

Public Sub SetRegValStr(lHive&, sKey$, sVal$, sData$)
    Dim hKey&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_SET_VALUE, hKey) <> 0 Then Exit Sub
    RegSetValueEx hKey, sVal, 0, REG_SZ, ByVal sData, Len(sData)
    RegCloseKey hKey
End Sub

Public Sub SetRegValDword(lHive&, sKey$, sVal$, lData&)
    Dim hKey&
    If RegOpenKeyEx(lHive, sKey, 0, KEY_SET_VALUE, hKey) <> 0 Then Exit Sub
    RegSetValueEx hKey, sVal, 0, REG_DWORD, lData, 4
    RegCloseKey hKey
End Sub

Public Sub KillFile(sFile$)
    Dim uSFIO As SHFILEOPSTRUCT
    With uSFIO
        .wFunc = FO_DELETE
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT
        If bUseRecycleBin Then .fFlags = .fFlags Or FOF_ALLOWUNDO
        .pFrom = sFile
    End With
    SHFileOperation uSFIO
End Sub

Public Sub KillFolder(sFolder$)
    Dim uSFIO As SHFILEOPSTRUCT
    With uSFIO
        .wFunc = FO_DELETE
        .fFlags = FOF_NOCONFIRMATION Or FOF_SILENT
        If bUseRecycleBin Then .fFlags = .fFlags Or FOF_ALLOWUNDO
        .pFrom = sFolder
    End With
    SHFileOperation uSFIO
End Sub

Public Sub ServiceDelete(sServiceName$)
    If Not bIsWinNT Then Exit Sub
    Dim hSCManager&, hService&
    If sServiceName = vbNullString Then Exit Sub
    
    hSCManager = OpenSCManager(vbNullString, vbNullString, SC_MANAGER_CREATE_SERVICE)
    If hSCManager > 0 Then
        hService = OpenService(hSCManager, sServiceName, SERVICE_ALL_ACCESS)
        If hService > 0 Then
            If DeleteService(hService) = 0 Then
                'Logg "Failed: ServiceDelete " & sCmd & " (operation failed)"
            Else
                bRebootNeeded = True
            End If
            CloseServiceHandle hService
        End If
        CloseServiceHandle hSCManager
    End If
End Sub



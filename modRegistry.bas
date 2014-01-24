Attribute VB_Name = "modMain"
'TOTEST: P2P Networking + Altnet files
'TOTEST: Kazaa 2.5x (?) uninstall regkey
'TOTEST: Kazaa 2.52 (+?) and Kazaa 2.6

'v1.00: original release
'v1.01: exclude SPORDER.DLL (other.def),
'       improve detection of nonstardard Kazaa folder
'v1.10: changed removal LSP hooks from registry hack to
'       WSA uninstall APIs ^_^,
'       use SHDeleteKey instead of RegDelKey API
'       added warning for removal of files in Shared Folder,
'       added RunOnce regval for BullGuard install
'       added 'delete to recycle bin' checkbox (default is enabled)
'       added 'delete only selected items' removal method
'       (never released)
'v1.20: added stuff and bundles from kazaa 2.5 and newer up to 3.0
'--------------------------------------

Option Explicit

Public Declare Function RegCloseKey Lib "advapi32.dll" (ByVal hKey As Long) As Long
Public Declare Function RegOpenKeyEx Lib "advapi32.dll" Alias "RegOpenKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal ulOptions As Long, ByVal samDesired As Long, phkResult As Long) As Long
Public Declare Function RegQueryValueEx Lib "advapi32.dll" Alias "RegQueryValueExA" (ByVal hKey As Long, ByVal lpValueName As String, ByVal lpReserved As Long, lpType As Long, lpData As Any, lpcbData As Long) As Long

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetSystemDirectory Lib "kernel32" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function GetTempPath Lib "kernel32" Alias "GetTempPathA" (ByVal nBufferLength As Long, ByVal lpBuffer As String) As Long
Private Declare Function GetUserName Lib "advapi32.dll" Alias "GetUserNameA" (ByVal lpBuffer As String, nSize As Long) As Long

Private Declare Function SHRestartSystemMB Lib "shell32" Alias "#59" (ByVal hOwner As Long, ByVal sExtraPrompt As String, ByVal uFlags As Long) As Long
Private Declare Function SHFileExists Lib "shell32" Alias "#45" (ByVal szPath As String) As Long

Private Declare Function FindWindow Lib "user32" Alias "FindWindowA" (ByVal lpClassName As String, ByVal lpWindowName As String) As Long
Private Declare Function IsWindowVisible Lib "user32" (ByVal hwnd As Long) As Long

Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
Private Declare Function GetSaveFileName Lib "comdlg32.dll" Alias "GetSaveFileNameA" (pOpenfilename As OPENFILENAME) As Long

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128
End Type

Private Type OPENFILENAME
    lStructSize As Long
    hwndOwner As Long
    hInstance As Long
    lpstrFilter As String
    lpstrCustomFilter As String
    nMaxCustFilter As Long
    nFilterIndex As Long
    lpstrFile As String
    nMaxFile As Long
    lpstrFileTitle As String
    nMaxFileTitle As Long
    lpstrInitialDir As String
    lpstrTitle As String
    flags As Long
    nFileOffset As Integer
    nFileExtension As Integer
    lpstrDefExt As String
    lCustData As Long
    lpfnHook As Long
    lpTemplateName As String
End Type

Private Const OFN_OVERWRITEPROMPT = &H2
Private Const OFN_HIDEREADONLY = &H4
Private Const OFN_PATHMUSTEXIST = &H800
Private Const OFN_NONETWORKBUTTON = &H20000

Private Const VER_PLATFORM_WIN32s = 0
Private Const VER_PLATFORM_WIN32_WINDOWS = 1
Private Const VER_PLATFORM_WIN32_NT = 2

Private Const HKEY_CLASSES_ROOT = &H80000000
Public Const HKEY_CURRENT_USER = &H80000001
Public Const HKEY_LOCAL_MACHINE = &H80000002
Private Const HKEY_USERS = &H80000003

Public Const KEY_QUERY_VALUE = &H1
Private Const KEY_SET_VALUE = &H2
Private Const KEY_CREATE_SUB_KEY = &H4
Private Const READ_CONTROL = &H20000
Private Const SYNCHRONIZE = &H100000
Private Const STANDARD_RIGHTS_WRITE = (READ_CONTROL)
Public Const KEY_WRITE = ((STANDARD_RIGHTS_WRITE Or KEY_SET_VALUE Or KEY_CREATE_SUB_KEY) And (Not SYNCHRONIZE))
Public Const REG_BINARY = 3
Public Const REG_DWORD = 4
Public Const REG_SZ = 1    ' string

Public sRegKeysKazaa$(), sFilesKazaa$(), sFoldersKazaa$(), sRegvalsKazaa$(), sProcessesKazaa$()
Public sRegKeysBDE$(), sFilesBDE$(), sFoldersBDE$(), sRegvalsBDE$(), sProcessesBDE$()
Public sRegKeysCyDoor$(), sFilesCyDoor$(), sFoldersCyDoor$(), sRegvalsCyDoor$(), sProcessesCyDoor$()
Public sRegKeysCommonName$(), sFilesCommonName$(), sFoldersCommonName$(), sRegvalsCommonName$(), sProcessesCommonName$()
Public sRegKeysNewDotNet$(), sFilesNewDotNet$(), sFoldersNewDotNet$(), sRegvalsNewDotNet$(), sProcessesNewDotNet$()
Public sRegKeysWebHancer$(), sFilesWebHancer$(), sFoldersWebHancer$(), sRegvalsWebHancer$(), sProcessesWebHancer$()
Public sRegKeysMediaLoads$(), sFilesMediaLoads$(), sFoldersMediaLoads$(), sRegvalsMediaLoads$(), sProcessesMediaLoads$()
Public sRegKeysSaveNow$(), sFilesSaveNow$(), sFoldersSaveNow$(), sRegvalsSaveNow$(), sProcessesSaveNow$()
Public sRegKeysDelfin$(), sFilesDelfin$(), sFoldersDelfin$(), sRegvalsDelfin$(), sProcessesDelfin$()
Public sRegKeysOnFlow$(), sFilesOnFlow$(), sFoldersOnFlow$(), sRegvalsOnFlow$(), sProcessesOnFlow$()
Public sRegKeysAltnet$(), sFilesAltnet$(), sFoldersAltnet$(), sRegvalsAltnet$(), sProcessesAltnet$()
Public sRegKeysBullguard$(), sFilesBullguard$(), sFoldersBullguard$(), sRegValsBullguard$(), sProcessesBullguard$()
Public sRegKeysGator$(), sFilesGator$(), sFoldersGator$(), sRegvalsGator$(), sProcessesGator$()
Public sRegKeysMyWay$(), sFilesMyWay$(), sFoldersMyWay$(), sRegvalsMyWay$(), sProcessesMyWay$()
Public sRegKeysP2P$(), sFilesP2P$(), sFoldersP2P$(), sRegvalsP2P$(), sProcessesP2P$()
Public sRegKeysPerfectNav$(), sFilesPerfectNav$(), sFoldersPerfectNav$(), sRegvalsPerfectNav$(), sProcessesPerfectNav$()
Public sRegKeysOther$(), sFilesOther$(), sFoldersOther$(), sRegvalsOther$(), sProcessesOther$()

Public sFoundRegKeys$(), sFoundFiles$(), sFoundFolders$(), sFoundRegvals$(), sFoundProcesses$()

Public sWinDir$, sWinSysDir$, sTempDir$
Public bDontRemoveBullguard As Boolean
Public bUseRecycleBin As Boolean, bIsWinNT As Boolean

Public Sub Scan()
    frmMain.fraFrame.Enabled = False
    EnumKazaaComponents
    EnumBDEComponents
    EnumCyDoorComponents
    EnumCommonNameComponents
    EnumNewDotNetComponents
    EnumWebHancerComponents
    EnumMediaLoadsComponents
    EnumSaveNowComponents
    EnumDelfinComponents
    EnumOnFlowComponents
    EnumOtherComponents
    
    EnumAltnetComponents
    EnumBullguardComponents
    EnumGatorComponents
    EnumMywayComponents
    EnumP2PComponents
    EnumPerfectnavComponents
    
    CheckWinsockLSP
    
    Dim i%
    ReDim sFoundRegKeys(0)
    ReDim sFoundFiles(0)
    ReDim sFoundFolders(0)
    ReDim sFoundRegvals(0)
    ReDim sFoundProcesses(0)
    With frmMain.lstLog
        For i = 0 To .ListCount - 1
            If Left(.List(i), 6) = "REGKEY" Then
                ReDim Preserve sFoundRegKeys(UBound(sFoundRegKeys) + 1)
                sFoundRegKeys(UBound(sFoundRegKeys)) = .List(i)
            End If
            If Left(.List(i), 6) = "REGVAL" Then
                ReDim Preserve sFoundRegvals(UBound(sFoundRegvals) + 1)
                sFoundRegvals(UBound(sFoundRegvals)) = .List(i)
            End If
            If Left(.List(i), 4) = "FILE" Then
                ReDim Preserve sFoundFiles(UBound(sFoundFiles) + 1)
                sFoundFiles(UBound(sFoundFiles)) = .List(i)
            End If
            If Left(.List(i), 6) = "FOLDER" Then
                ReDim Preserve sFoundFolders(UBound(sFoundFolders) + 1)
                sFoundFolders(UBound(sFoundFolders)) = .List(i)
            End If
            If Left(.List(i), 7) = "PROCESS" Then
                ReDim Preserve sFoundProcesses(UBound(sFoundProcesses) + 1)
                sFoundProcesses(UBound(sFoundProcesses)) = .List(i)
            End If
        Next i
        
        .Clear
        
        For i = 1 To UBound(sFoundProcesses)
            .AddItem sFoundProcesses(i)
        Next i
        For i = 1 To UBound(sFoundRegKeys)
            .AddItem sFoundRegKeys(i)
        Next i
        For i = 1 To UBound(sFoundRegvals)
            .AddItem sFoundRegvals(i)
        Next i
        For i = 1 To UBound(sFoundFolders)
            .AddItem sFoundFolders(i)
        Next i
        For i = 1 To UBound(sFoundFiles)
            .AddItem sFoundFiles(i)
        Next i
    End With
    With frmMain.lstLogSel
        .Clear
        
        For i = 1 To UBound(sFoundProcesses)
            .AddItem sFoundProcesses(i)
        Next i
        For i = 1 To UBound(sFoundRegKeys)
            .AddItem sFoundRegKeys(i)
        Next i
        For i = 1 To UBound(sFoundRegvals)
            .AddItem sFoundRegvals(i)
        Next i
        For i = 1 To UBound(sFoundFolders)
            .AddItem sFoundFolders(i)
        Next i
        For i = 1 To UBound(sFoundFiles)
            .AddItem sFoundFiles(i)
        Next i
    End With
    
    frmMain.fraFrame.Enabled = True
    Status "Scan complete, " & frmMain.lstLog.ListCount & " items found"
End Sub

Public Sub Destroy()
    Dim i&, sItem$, lHive&, sKey$, sVal$, bFixLSPNeeded As Boolean
    Dim sLockedFiles$(), sLockedFolders$(), sMsg$, sService$
    On Error GoTo Error:
    
    'special case for bullguard
    Dim sProgramFiles$
    sProgramFiles = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProgramFilesDir", "C:\Program Files")
    If FileExists(sProgramFiles & "\Bullguard Software\Bullguard\uninst.exe") Then
        If MsgBox("KazaaBegone has detected a full install of BullGuard " & _
                   "on your system. BullGuard is a free package, based " & _
                   "on the BitDefender antivirus program and the Sygate " & _
                   "firewall." & vbCrLf & vbCrLf & _
                   "KazaaBegone cannot safely remove these." & _
                   vbCrLf & "Do you want to run BullGuard's own " & _
                   "uninstaller?" & vbCrLf & vbCrLf & _
                   "If you choose 'No', KazaaBegone will not attempt " & _
                   "to clean up any BullGuard components." & vbCrLf & _
                   "If you choose 'Yes', KazaaBegone will clean up " & _
                   "any leftovers that were not removed by the " & _
                   "BullGuard uninstaller.", vbQuestion + vbYesNo) = vbYes Then
            ShellRun sProgramFiles & "\BullGuard Software\Bullguard\uninst.exe"
            If MsgBox("The BullGuard uninstaller has been started." & _
                   vbCrLf & "Click OK to continue after the " & _
                   "uninstaller finished, or Cancel to stop " & _
                   "removal.", vbInformation + vbOKCancel) = vbCancel Then
                Exit Sub
            End If
        Else
            bDontRemoveBullguard = True
        End If
    End If
    
    ReDim sLockedFiles(0)
    ReDim sLockedFolders(0)
    If frmMain.lstLog.Visible Then
        With frmMain.lstLog
            If .ListCount = 0 Then
                MsgBox "Nothing to delete!", vbExclamation
                Exit Sub
            End If
            
            For i = .ListCount - 1 To 0 Step -1
                sItem = .List(i)
                
                If bDontRemoveBullguard And InStr(1, sItem, "bullguard", vbTextCompare) > 0 Then
                    GoTo NextItem
                End If
                
                If Left(sItem, 6) = "REGKEY" Then
                    'REGKEY: [Kazaa] HKLM\Software\Kazaa
                    
                    'If InStr(sItem, "\Catalog_Entries\") > 0 Or _
                    '   InStr(sItem, "Broken Winsock stack") > 0 Or _
                    '   InStr(sItem, "Hijacked Internet access by") > 0 Or _
                    '   InStr(sItem, "Broken Internet access") > 0 Then
                    '    bFixLSPNeeded = True
                    '    .RemoveItem i
                    '    GoTo NextItem
                    'End If
                    
                    sItem = Mid(sItem, InStr(sItem, "]") + 2)
                    Select Case Left(sItem, 4)
                        Case "HKCR": lHive = HKEY_CLASSES_ROOT
                        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
                        Case "HKCU": lHive = HKEY_CURRENT_USER
                        Case "HKUS": lHive = HKEY_USERS
                    End Select
                    sKey = Mid(sItem, 6)
                    If InStr(1, sKey, "\Services\", vbTextCompare) > 0 Then
                        'it's a (bullguard) service so kill it first
                        'so the LEGACY_ clone goes down with it
                        sService = Mid(sKey, InStrRev(sKey, "\") + 1)
                        ServiceDelete sService
                    End If
                    KillRegKey lHive, sKey
                    .RemoveItem i
                    GoTo NextItem
                End If
                If Left(sItem, 6) = "REGVAL" Then
                    'REGVAL: [Other] HKLM\System\CurrentControlSet\Control\Shutdown,SetupProgramRan
                    
                    sItem = Mid(sItem, InStr(sItem, "]") + 2)
                    Select Case Left(sItem, 4)
                        Case "HKCR": lHive = HKEY_CLASSES_ROOT
                        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
                        Case "HKCU": lHive = HKEY_CURRENT_USER
                        Case "HKUS": lHive = HKEY_USERS
                    End Select
                    sKey = Mid(sItem, 6)
                    sVal = Mid(sKey, InStrRev(sKey, ",") + 1)
                    sKey = Left(sKey, InStrRev(sKey, ",") - 1)
                    KillRegVal lHive, sKey, sVal
                    .RemoveItem i
                    GoTo NextItem
                End If
                If Left(sItem, 4) = "FILE" Then
                    'FILE: [Kazaa] C:\Program Files\Kazaa\Kazaa.exe
                    sItem = Mid(sItem, InStr(sItem, "]") + 2)
                    KillFile sItem
                    If FileExists(sItem) Then
                        ReDim Preserve sLockedFiles(UBound(sLockedFiles) + 1)
                        sLockedFiles(UBound(sLockedFiles)) = sItem
                    Else
                        .RemoveItem i
                    End If
                    GoTo NextItem
                End If
                If Left(sItem, 6) = "FOLDER" Then
                    'FOLDER: [Kazaa] C:\Program Files\Kazaa
                    sItem = Mid(sItem, InStr(sItem, "]") + 2)
                    KillFolder sItem
                    If FolderExists(sItem) Then
                        ReDim Preserve sLockedFolders(UBound(sLockedFolders) + 1)
                        sLockedFolders(UBound(sLockedFolders)) = sItem
                    Else
                        .RemoveItem i
                    End If
                End If
                If Left(sItem, 7) = "WINSOCK" Then
                    'WINSOCK: [New.Net] New.Net Namespace Provider
                    'WINSOCK: [webHancer] webHancer [UDP/IP]
                    sItem = Mid(sItem, InStr(sItem, "]") + 2)
                    FixWinsockLSP sItem
                    .RemoveItem i
                End If
                If Left(sItem, 7) = "PROCESS" Then
                    'PROCESSS: [BDE] C:\WINDOWS\system32\bdeinstall.exe
                    sItem = Mid(sItem, InStr(sItem, "]") + 2)
                    KillProcess sItem
                    .RemoveItem i
                End If
NextItem:
                If .ListCount > 0 Then
                    Status "Uninstalling... " & CStr(100 - 100 * Int(CDbl(i) / .ListCount)) & " %"
                Else
                    Status "Uninstalling... 100%"
                End If
            Next i
        End With
    Else
        With frmMain.lstLogSel
            If .ListCount = 0 Or .SelCount = 0 Then
                MsgBox "Nothing to delete!", vbExclamation
                Exit Sub
            End If
            
            For i = .ListCount - 1 To 0 Step -1
                If .Selected(i) Then
                    sItem = .List(i)
                    
                    If bDontRemoveBullguard And InStr(1, sItem, "bullguard", vbTextCompare) > 0 Then
                        GoTo NextItemSel
                    End If
                    
                    If Left(sItem, 6) = "REGKEY" Then
                        'REGKEY: [Kazaa] HKLM\Software\Kazaa
                        
                        'If InStr(sItem, "\Catalog_Entries\") > 0 Or _
                        '   InStr(sItem, "Broken Winsock stack") > 0 Or _
                        '   InStr(sItem, "Hijacked Internet access by") > 0 Or _
                        '   InStr(sItem, "Broken Internet access") > 0 Then
                        '    bFixLSPNeeded = True
                        '    .RemoveItem i
                        '    GoTo NextItem
                        'End If
                        
                        sItem = Mid(sItem, InStr(sItem, "]") + 2)
                        Select Case Left(sItem, 4)
                            Case "HKCR": lHive = HKEY_CLASSES_ROOT
                            Case "HKLM": lHive = HKEY_LOCAL_MACHINE
                            Case "HKCU": lHive = HKEY_CURRENT_USER
                            Case "HKUS": lHive = HKEY_USERS
                        End Select
                        sKey = Mid(sItem, 6)
                        If InStr(1, sKey, "\Services\", vbTextCompare) > 0 Then
                            'it's a (bullguard) service so kill it first
                            'so the LEGACY_ clone goes down with it
                            sService = Mid(sKey, InStrRev(sKey, "\") + 1)
                            ServiceDelete sService
                        End If
                        KillRegKey lHive, sKey
                        .RemoveItem i
                        GoTo NextItemSel
                    End If
                    If Left(sItem, 6) = "REGVAL" Then
                        'REGVAL: [Other] HKLM\System\CurrentControlSet\Control\Shutdown,SetupProgramRan
                        
                        sItem = Mid(sItem, InStr(sItem, "]") + 2)
                        Select Case Left(sItem, 4)
                            Case "HKCR": lHive = HKEY_CLASSES_ROOT
                            Case "HKLM": lHive = HKEY_LOCAL_MACHINE
                            Case "HKCU": lHive = HKEY_CURRENT_USER
                            Case "HKUS": lHive = HKEY_USERS
                        End Select
                        sKey = Mid(sItem, 6)
                        sVal = Mid(sKey, InStrRev(sKey, ",") + 1)
                        sKey = Left(sKey, InStrRev(sKey, ",") - 1)
                        KillRegVal lHive, sKey, sVal
                        .RemoveItem i
                        GoTo NextItemSel
                    End If
                    If Left(sItem, 4) = "FILE" Then
                        'FILE: [Kazaa] C:\Program Files\Kazaa\Kazaa.exe
                        sItem = Mid(sItem, InStr(sItem, "]") + 2)
                        KillFile sItem
                        If FileExists(sItem) Then
                            ReDim Preserve sLockedFiles(UBound(sLockedFiles) + 1)
                            sLockedFiles(UBound(sLockedFiles)) = sItem
                        Else
                            .RemoveItem i
                        End If
                        GoTo NextItemSel
                    End If
                    If Left(sItem, 6) = "FOLDER" Then
                        'FOLDER: [Kazaa] C:\Program Files\Kazaa
                        sItem = Mid(sItem, InStr(sItem, "]") + 2)
                        KillFolder sItem
                        If FolderExists(sItem) Then
                            ReDim Preserve sLockedFolders(UBound(sLockedFolders) + 1)
                            sLockedFolders(UBound(sLockedFolders)) = sItem
                        Else
                            .RemoveItem i
                        End If
                    End If
                    If Left(sItem, 7) = "WINSOCK" Then
                        'WINSOCK: [New.Net] New.Net Namespace Provider
                        'WINSOCK: [webHancer] webHancer [UDP/IP]
                        sItem = Mid(sItem, InStr(sItem, "]") + 2)
                        FixWinsockLSP sItem
                        .RemoveItem i
                    End If
                    If Left(sItem, 7) = "PROCESS" Then
                        'PROCESSS: [BDE] C:\WINDOWS\system32\bdeinstall.exe
                        sItem = Mid(sItem, InStr(sItem, "]") + 2)
                        KillProcess sItem
                        .RemoveItem i
                    End If
NextItemSel:
                End If
                If .ListCount > 0 Then
                    Status "Uninstalling... " & CStr(100 - 100 * Int(CDbl(i) / .ListCount)) & " %"
                Else
                    Status "Uninstalling... 100%"
                End If
            Next i
        End With
    End If
    Status "Done!"
    'If bFixLSPNeeded Then
    '    bRebootNeeded = False
    '    FixLSP
    'End If
    
    If UBound(sLockedFiles) <> 0 Or UBound(sLockedFolders) <> 0 Then
        'not all files/folders could be deleted
        sMsg = "The following files or folders could not be " & _
               "deleted because they were in use:" & _
               vbCrLf & Join(sLockedFiles, vbCrLf) & _
               vbCrLf & Join(sLockedFolders, vbCrLf) & vbCrLf & _
               "You should restart the system and try again."
        
        MsgBox sMsg, vbExclamation
    Else
        sMsg = "The uninstall of Kazaa and all bundled " & _
               "programs has been completed successfully!"
        MsgBox sMsg, vbInformation
    End If
    If bRebootNeeded Then
        If bIsWinNT Then
            SHRestartSystemMB frmMain.hwnd, StrConv("Your Winsock stack has been restored." & vbCrLf & _
            "However, if your Internet connection is broken on reboot, " & _
            "download LSPFix from www.cexx.org/lspfix.htm to restore it." & vbCrLf & vbCrLf, vbUnicode), 2
        Else
            SHRestartSystemMB frmMain.hwnd, "Your Winsock stack has been restored." & vbCrLf & _
            "However, if your Internet connection is broken on reboot, " & _
            "download LSPFix from www.cexx.org/lspfix.htm to restore it." & vbCrLf & vbCrLf, 2
        End If
    End If
    Exit Sub
    
Error:
    MsgBox "Unexpected error occurred at modMain_Destroy:" & vbCrLf & _
           "Error#" & Err.Number & ": " & Err.Description & "." & vbCrLf & _
           vbCrLf & "The uninstall may not have completed!", vbExclamation
End Sub
    
Public Function CmnDlgSaveFile$(sFilter$, sTitle$, Optional sDefFilename$, Optional sInitialDir$, Optional sDefExt$)
    Dim uOFN As OPENFILENAME
    With uOFN
        .lStructSize = Len(uOFN)
        .lpstrFilter = Replace(sFilter, "|", Chr(0)) & Chr(0) & Chr(0)
        .lpstrFile = sDefFilename & String(260 - Len(sDefFilename), 0)
        .nMaxFile = 260
        .lpstrInitialDir = sInitialDir
        .lpstrTitle = sTitle
        .flags = OFN_HIDEREADONLY Or OFN_NONETWORKBUTTON Or _
                 OFN_OVERWRITEPROMPT Or OFN_PATHMUSTEXIST
        .lpstrDefExt = sDefExt
        If GetSaveFileName(uOFN) <> 0 Then
            CmnDlgSaveFile = TrimNull(.lpstrFile)
        End If
    End With
End Function

Public Sub LoadDefs()
    'use .def files to load var arrays
    Dim sPath$, sLine$
        
    sPath = App.Path & IIf(Right(App.Path, 1) = "\", "", "\")
        
    If Not FileExists(sPath & "kazaa.def") Then GoTo MissingDef:
    If Not FileExists(sPath & "bde.def") Then GoTo MissingDef:
    If Not FileExists(sPath & "cydoor.def") Then GoTo MissingDef:
    If Not FileExists(sPath & "commonname.def") Then GoTo MissingDef:
    If Not FileExists(sPath & "newdotnet.def") Then GoTo MissingDef:
    If Not FileExists(sPath & "webhancer.def") Then GoTo MissingDef:
    If Not FileExists(sPath & "medialoads.def") Then GoTo MissingDef:
    If Not FileExists(sPath & "savenow.def") Then GoTo MissingDef:
    If Not FileExists(sPath & "delfin.def") Then GoTo MissingDef:
    If Not FileExists(sPath & "onflow.def") Then GoTo MissingDef:
    
    If Not FileExists(sPath & "altnet.def") Then GoTo MissingDef:
    If Not FileExists(sPath & "bullguard.def") Then GoTo MissingDef:
    If Not FileExists(sPath & "gator.def") Then GoTo MissingDef:
    If Not FileExists(sPath & "myway.def") Then GoTo MissingDef:
    If Not FileExists(sPath & "p2pnetworking.def") Then GoTo MissingDef:
    If Not FileExists(sPath & "perfectnav.def") Then GoTo MissingDef:
    If Not FileExists(sPath & "other.def") Then GoTo MissingDef:
    
    ReDim sFilesKazaa(0)
    ReDim sFoldersKazaa(0)
    ReDim sRegKeysKazaa(0)
    ReDim sRegvalsKazaa(0)
    ReDim sProcessesKazaa(0)
    Open sPath & "kazaa.def" For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> vbNullString And InStr(sLine, "//") <> 1 Then
                Select Case Left(sLine, InStr(sLine, "=") - 1)
                    Case "KazaaFile"
                        ReDim Preserve sFilesKazaa(UBound(sFilesKazaa) + 1)
                        sFilesKazaa(UBound(sFilesKazaa)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "KazaaFolder"
                        ReDim Preserve sFoldersKazaa(UBound(sFoldersKazaa) + 1)
                        sFoldersKazaa(UBound(sFoldersKazaa)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "KazaaRegkey"
                        ReDim Preserve sRegKeysKazaa(UBound(sRegKeysKazaa) + 1)
                        sRegKeysKazaa(UBound(sRegKeysKazaa)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "KazaaRegval"
                        ReDim Preserve sRegvalsKazaa(UBound(sRegvalsKazaa) + 1)
                        sRegvalsKazaa(UBound(sRegvalsKazaa)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "KazaaProcess"
                        ReDim Preserve sProcessesKazaa(UBound(sProcessesKazaa) + 1)
                        sProcessesKazaa(UBound(sProcessesKazaa)) = Mid(sLine, InStr(sLine, "=") + 1)
                End Select
            End If
        Loop Until EOF(1)
    Close #1
    
    ReDim sFilesBDE(0)
    ReDim sFoldersBDE(0)
    ReDim sRegKeysBDE(0)
    ReDim sRegvalsBDE(0)
    ReDim sProcessesBDE(0)
    Open sPath & "bde.def" For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> vbNullString And InStr(sLine, "//") <> 1 Then
                Select Case Left(sLine, InStr(sLine, "=") - 1)
                    Case "BDEFile"
                        ReDim Preserve sFilesBDE(UBound(sFilesBDE) + 1)
                        sFilesBDE(UBound(sFilesBDE)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "BDEFolder"
                        ReDim Preserve sFoldersBDE(UBound(sFoldersBDE) + 1)
                        sFoldersBDE(UBound(sFoldersBDE)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "BDERegkey"
                        ReDim Preserve sRegKeysBDE(UBound(sRegKeysBDE) + 1)
                        sRegKeysBDE(UBound(sRegKeysBDE)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "BDERegval"
                        ReDim Preserve sRegvalsBDE(UBound(sRegvalsBDE) + 1)
                        sRegvalsBDE(UBound(sRegvalsBDE)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "BDEProcess"
                        ReDim Preserve sProcessesBDE(UBound(sProcessesBDE) + 1)
                        sProcessesBDE(UBound(sProcessesBDE)) = Mid(sLine, InStr(sLine, "=") + 1)
                End Select
            End If
        Loop Until EOF(1)
    Close #1
    
    ReDim sFilesCyDoor(0)
    ReDim sFoldersCyDoor(0)
    ReDim sRegKeysCyDoor(0)
    ReDim sRegvalsCyDoor(0)
    ReDim sProcessesCyDoor(0)
    Open sPath & "cydoor.def" For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> vbNullString And InStr(sLine, "//") <> 1 Then
                Select Case Left(sLine, InStr(sLine, "=") - 1)
                    Case "CydoorFile"
                        ReDim Preserve sFilesCyDoor(UBound(sFilesCyDoor) + 1)
                        sFilesCyDoor(UBound(sFilesCyDoor)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "CydoorFolder"
                        ReDim Preserve sFoldersCyDoor(UBound(sFoldersCyDoor) + 1)
                        sFoldersCyDoor(UBound(sFoldersCyDoor)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "CydoorRegkey"
                        ReDim Preserve sRegKeysCyDoor(UBound(sRegKeysCyDoor) + 1)
                        sRegKeysCyDoor(UBound(sRegKeysCyDoor)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "CydoorRegval"
                        ReDim Preserve sRegvalsCyDoor(UBound(sRegvalsCyDoor) + 1)
                        sRegvalsCyDoor(UBound(sRegvalsCyDoor)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "CydoorProcess"
                        ReDim Preserve sProcessesCyDoor(UBound(sProcessesCyDoor) + 1)
                        sProcessesCyDoor(UBound(sProcessesCyDoor)) = Mid(sLine, InStr(sLine, "=") + 1)
                End Select
            End If
        Loop Until EOF(1)
    Close #1
    
    ReDim sFilesCommonName(0)
    ReDim sFoldersCommonName(0)
    ReDim sRegKeysCommonName(0)
    ReDim sRegvalsCommonName(0)
    ReDim sProcessesCommonName(0)
    Open sPath & "commonname.def" For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> vbNullString And InStr(sLine, "//") <> 1 Then
                Select Case Left(sLine, InStr(sLine, "=") - 1)
                    Case "CommonnameFile"
                        ReDim Preserve sFilesCommonName(UBound(sFilesCommonName) + 1)
                        sFilesCommonName(UBound(sFilesCommonName)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "CommonnameFolder"
                        ReDim Preserve sFoldersCommonName(UBound(sFoldersCommonName) + 1)
                        sFoldersCommonName(UBound(sFoldersCommonName)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "CommonnameRegkey"
                        ReDim Preserve sRegKeysCommonName(UBound(sRegKeysCommonName) + 1)
                        sRegKeysCommonName(UBound(sRegKeysCommonName)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "CommonnameRegval"
                        ReDim Preserve sRegvalsCommonName(UBound(sRegvalsCommonName) + 1)
                        sRegvalsCommonName(UBound(sRegvalsCommonName)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "CommonnameProcess"
                        ReDim Preserve sProcessesCommonName(UBound(sProcessesCommonName) + 1)
                        sProcessesCommonName(UBound(sProcessesCommonName)) = Mid(sLine, InStr(sLine, "=") + 1)
                End Select
            End If
        Loop Until EOF(1)
    Close #1
    
    ReDim sFilesNewDotNet(0)
    ReDim sFoldersNewDotNet(0)
    ReDim sRegKeysNewDotNet(0)
    ReDim sRegvalsNewDotNet(0)
    ReDim sProcessesNewDotNet(0)
    Open sPath & "newdotnet.def" For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> vbNullString And InStr(sLine, "//") <> 1 Then
                Select Case Left(sLine, InStr(sLine, "=") - 1)
                    Case "NewdotnetFile"
                        ReDim Preserve sFilesNewDotNet(UBound(sFilesNewDotNet) + 1)
                        sFilesNewDotNet(UBound(sFilesNewDotNet)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "NewdotnetFolder"
                        ReDim Preserve sFoldersNewDotNet(UBound(sFoldersNewDotNet) + 1)
                        sFoldersNewDotNet(UBound(sFoldersNewDotNet)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "NewdotnetRegkey"
                        ReDim Preserve sRegKeysNewDotNet(UBound(sRegKeysNewDotNet) + 1)
                        sRegKeysNewDotNet(UBound(sRegKeysNewDotNet)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "NewdotnetRegval"
                        ReDim Preserve sRegvalsNewDotNet(UBound(sRegvalsNewDotNet) + 1)
                        sRegvalsNewDotNet(UBound(sRegvalsNewDotNet)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "NewdotnetProcess"
                        ReDim Preserve sProcessesNewDotNet(UBound(sProcessesNewDotNet) + 1)
                        sProcessesNewDotNet(UBound(sProcessesNewDotNet)) = Mid(sLine, InStr(sLine, "=") + 1)
                End Select
            End If
        Loop Until EOF(1)
    Close #1
    
    ReDim sFilesWebHancer(0)
    ReDim sFoldersWebHancer(0)
    ReDim sRegKeysWebHancer(0)
    ReDim sRegvalsWebHancer(0)
    ReDim sProcessesWebHancer(0)
    Open sPath & "webhancer.def" For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> vbNullString And InStr(sLine, "//") <> 1 Then
                Select Case Left(sLine, InStr(sLine, "=") - 1)
                    Case "WebhancerFile"
                        ReDim Preserve sFilesWebHancer(UBound(sFilesWebHancer) + 1)
                        sFilesWebHancer(UBound(sFilesWebHancer)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "WebhancerFolder"
                        ReDim Preserve sFoldersWebHancer(UBound(sFoldersWebHancer) + 1)
                        sFoldersWebHancer(UBound(sFoldersWebHancer)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "WebhancerRegkey"
                        ReDim Preserve sRegKeysWebHancer(UBound(sRegKeysWebHancer) + 1)
                        sRegKeysWebHancer(UBound(sRegKeysWebHancer)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "WebhancerRegval"
                        ReDim Preserve sRegvalsWebHancer(UBound(sRegvalsWebHancer) + 1)
                        sRegvalsWebHancer(UBound(sRegvalsWebHancer)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "WebhancerProcess"
                        ReDim Preserve sProcessesWebHancer(UBound(sProcessesWebHancer) + 1)
                        sProcessesWebHancer(UBound(sProcessesWebHancer)) = Mid(sLine, InStr(sLine, "=") + 1)
                End Select
            End If
        Loop Until EOF(1)
    Close #1
    
    ReDim sFilesMediaLoads(0)
    ReDim sFoldersMediaLoads(0)
    ReDim sRegKeysMediaLoads(0)
    ReDim sRegvalsMediaLoads(0)
    ReDim sProcessesMediaLoads(0)
    Open sPath & "medialoads.def" For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> vbNullString And InStr(sLine, "//") <> 1 Then
                Select Case Left(sLine, InStr(sLine, "=") - 1)
                    Case "MedialoadsFile"
                        ReDim Preserve sFilesMediaLoads(UBound(sFilesMediaLoads) + 1)
                        sFilesMediaLoads(UBound(sFilesMediaLoads)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "MedialoadsFolder"
                        ReDim Preserve sFoldersMediaLoads(UBound(sFoldersMediaLoads) + 1)
                        sFoldersMediaLoads(UBound(sFoldersMediaLoads)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "MedialoadsRegkey"
                        ReDim Preserve sRegKeysMediaLoads(UBound(sRegKeysMediaLoads) + 1)
                        sRegKeysMediaLoads(UBound(sRegKeysMediaLoads)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "MedialoadsRegval"
                        ReDim Preserve sRegvalsMediaLoads(UBound(sRegvalsMediaLoads) + 1)
                        sRegvalsMediaLoads(UBound(sRegvalsMediaLoads)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "MedialoadsProcess"
                        ReDim Preserve sProcessesMediaLoads(UBound(sProcessesMediaLoads) + 1)
                        sProcessesMediaLoads(UBound(sProcessesMediaLoads)) = Mid(sLine, InStr(sLine, "=") + 1)
                End Select
            End If
        Loop Until EOF(1)
    Close #1
    
    ReDim sFilesSaveNow(0)
    ReDim sFoldersSaveNow(0)
    ReDim sRegKeysSaveNow(0)
    ReDim sRegvalsSaveNow(0)
    ReDim sProcessesSaveNow(0)
    Open sPath & "Savenow.def" For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> vbNullString And InStr(sLine, "//") <> 1 Then
                Select Case Left(sLine, InStr(sLine, "=") - 1)
                    Case "SavenowFile"
                        ReDim Preserve sFilesSaveNow(UBound(sFilesSaveNow) + 1)
                        sFilesSaveNow(UBound(sFilesSaveNow)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "SavenowFolder"
                        ReDim Preserve sFoldersSaveNow(UBound(sFoldersSaveNow) + 1)
                        sFoldersSaveNow(UBound(sFoldersSaveNow)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "SavenowRegkey"
                        ReDim Preserve sRegKeysSaveNow(UBound(sRegKeysSaveNow) + 1)
                        sRegKeysSaveNow(UBound(sRegKeysSaveNow)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "SavenowRegval"
                        ReDim Preserve sRegvalsSaveNow(UBound(sRegvalsSaveNow) + 1)
                        sRegvalsSaveNow(UBound(sRegvalsSaveNow)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "SavenowProcess"
                        ReDim Preserve sProcessesSaveNow(UBound(sProcessesSaveNow) + 1)
                        sProcessesSaveNow(UBound(sProcessesSaveNow)) = Mid(sLine, InStr(sLine, "=") + 1)
                End Select
            End If
        Loop Until EOF(1)
    Close #1
    
    ReDim sFilesDelfin(0)
    ReDim sFoldersDelfin(0)
    ReDim sRegKeysDelfin(0)
    ReDim sRegvalsDelfin(0)
    ReDim sProcessesDelfin(0)
    Open sPath & "Delfin.def" For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> vbNullString And InStr(sLine, "//") <> 1 Then
                Select Case Left(sLine, InStr(sLine, "=") - 1)
                    Case "DelfinFile"
                        ReDim Preserve sFilesDelfin(UBound(sFilesDelfin) + 1)
                        sFilesDelfin(UBound(sFilesDelfin)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "DelfinFolder"
                        ReDim Preserve sFoldersDelfin(UBound(sFoldersDelfin) + 1)
                        sFoldersDelfin(UBound(sFoldersDelfin)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "DelfinRegkey"
                        ReDim Preserve sRegKeysDelfin(UBound(sRegKeysDelfin) + 1)
                        sRegKeysDelfin(UBound(sRegKeysDelfin)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "DelfinRegval"
                        ReDim Preserve sRegvalsDelfin(UBound(sRegvalsDelfin) + 1)
                        sRegvalsDelfin(UBound(sRegvalsDelfin)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "DelfinProcess"
                        ReDim Preserve sProcessesDelfin(UBound(sProcessesDelfin) + 1)
                        sProcessesDelfin(UBound(sProcessesDelfin)) = Mid(sLine, InStr(sLine, "=") + 1)
                End Select
            End If
        Loop Until EOF(1)
    Close #1
    
    ReDim sFilesOnFlow(0)
    ReDim sFoldersOnFlow(0)
    ReDim sRegKeysOnFlow(0)
    ReDim sRegvalsOnFlow(0)
    ReDim sProcessesOnFlow(0)
    Open sPath & "Onflow.def" For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> vbNullString And InStr(sLine, "//") <> 1 Then
                Select Case Left(sLine, InStr(sLine, "=") - 1)
                    Case "OnflowFile"
                        ReDim Preserve sFilesOnFlow(UBound(sFilesOnFlow) + 1)
                        sFilesOnFlow(UBound(sFilesOnFlow)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "OnflowFolder"
                        ReDim Preserve sFoldersOnFlow(UBound(sFoldersOnFlow) + 1)
                        sFoldersOnFlow(UBound(sFoldersOnFlow)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "OnflowRegkey"
                        ReDim Preserve sRegKeysOnFlow(UBound(sRegKeysOnFlow) + 1)
                        sRegKeysOnFlow(UBound(sRegKeysOnFlow)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "OnflowRegval"
                        ReDim Preserve sRegvalsOnFlow(UBound(sRegvalsOnFlow) + 1)
                        sRegvalsOnFlow(UBound(sRegvalsOnFlow)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "OnflowProcess"
                        ReDim Preserve sProcessesOnFlow(UBound(sProcessesOnFlow) + 1)
                        sProcessesOnFlow(UBound(sProcessesOnFlow)) = Mid(sLine, InStr(sLine, "=") + 1)
                End Select
            End If
        Loop Until EOF(1)
    Close #1
    
    ReDim sFilesOther(0)
    ReDim sFoldersOther(0)
    ReDim sRegKeysOther(0)
    ReDim sRegvalsOther(0)
    ReDim sProcessesOther(0)
    Open sPath & "Other.def" For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> vbNullString And InStr(sLine, "//") <> 1 Then
                Select Case Left(sLine, InStr(sLine, "=") - 1)
                    Case "OtherFile"
                        ReDim Preserve sFilesOther(UBound(sFilesOther) + 1)
                        sFilesOther(UBound(sFilesOther)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "OtherFolder"
                        ReDim Preserve sFoldersOther(UBound(sFoldersOther) + 1)
                        sFoldersOther(UBound(sFoldersOther)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "OtherRegkey"
                        ReDim Preserve sRegKeysOther(UBound(sRegKeysOther) + 1)
                        sRegKeysOther(UBound(sRegKeysOther)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "OtherRegval"
                        ReDim Preserve sRegvalsOther(UBound(sRegvalsOther) + 1)
                        sRegvalsOther(UBound(sRegvalsOther)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "OtherProcess"
                        ReDim Preserve sProcessesOther(UBound(sProcessesOther) + 1)
                        sProcessesOther(UBound(sProcessesOther)) = Mid(sLine, InStr(sLine, "=") + 1)
                End Select
            End If
        Loop Until EOF(1)
    Close #1
    
    '-----------
    ReDim sFilesAltnet(0)
    ReDim sFoldersAltnet(0)
    ReDim sRegKeysAltnet(0)
    ReDim sRegvalsAltnet(0)
    ReDim sProcessesAltnet(0)
    Open sPath & "Altnet.def" For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> vbNullString And InStr(sLine, "//") <> 1 Then
                Select Case Left(sLine, InStr(sLine, "=") - 1)
                    Case "AltnetFile"
                        ReDim Preserve sFilesAltnet(UBound(sFilesAltnet) + 1)
                        sFilesAltnet(UBound(sFilesAltnet)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "AltnetFolder"
                        ReDim Preserve sFoldersAltnet(UBound(sFoldersAltnet) + 1)
                        sFoldersAltnet(UBound(sFoldersAltnet)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "AltnetRegkey"
                        ReDim Preserve sRegKeysAltnet(UBound(sRegKeysAltnet) + 1)
                        sRegKeysAltnet(UBound(sRegKeysAltnet)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "AltnetRegval"
                        ReDim Preserve sRegvalsAltnet(UBound(sRegvalsAltnet) + 1)
                        sRegvalsAltnet(UBound(sRegvalsAltnet)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "AltnetProcess"
                        ReDim Preserve sProcessesAltnet(UBound(sProcessesAltnet) + 1)
                        sProcessesAltnet(UBound(sProcessesAltnet)) = Mid(sLine, InStr(sLine, "=") + 1)
                End Select
            End If
        Loop Until EOF(1)
    Close #1
    
    ReDim sFilesBullguard(0)
    ReDim sFoldersBullguard(0)
    ReDim sRegKeysBullguard(0)
    ReDim sRegValsBullguard(0)
    ReDim sProcessesBullguard(0)
    Open sPath & "Bullguard.def" For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> vbNullString And InStr(sLine, "//") <> 1 Then
                Select Case Left(sLine, InStr(sLine, "=") - 1)
                    Case "BullguardFile"
                        ReDim Preserve sFilesBullguard(UBound(sFilesBullguard) + 1)
                        sFilesBullguard(UBound(sFilesBullguard)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "BullguardFolder"
                        ReDim Preserve sFoldersBullguard(UBound(sFoldersBullguard) + 1)
                        sFoldersBullguard(UBound(sFoldersBullguard)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "BullguardRegkey"
                        ReDim Preserve sRegKeysBullguard(UBound(sRegKeysBullguard) + 1)
                        sRegKeysBullguard(UBound(sRegKeysBullguard)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "BullguardRegval"
                        ReDim Preserve sRegValsBullguard(UBound(sRegValsBullguard) + 1)
                        sRegValsBullguard(UBound(sRegValsBullguard)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "BullguardProcess"
                        ReDim Preserve sProcessesBullguard(UBound(sProcessesBullguard) + 1)
                        sProcessesBullguard(UBound(sProcessesBullguard)) = Mid(sLine, InStr(sLine, "=") + 1)
                End Select
            End If
        Loop Until EOF(1)
    Close #1
    
    ReDim sFilesGator(0)
    ReDim sFoldersGator(0)
    ReDim sRegKeysGator(0)
    ReDim sRegvalsGator(0)
    ReDim sProcessesGator(0)
    Open sPath & "Gator.def" For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> vbNullString And InStr(sLine, "//") <> 1 Then
                Select Case Left(sLine, InStr(sLine, "=") - 1)
                    Case "GatorFile"
                        ReDim Preserve sFilesGator(UBound(sFilesGator) + 1)
                        sFilesGator(UBound(sFilesGator)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "GatorFolder"
                        ReDim Preserve sFoldersGator(UBound(sFoldersGator) + 1)
                        sFoldersGator(UBound(sFoldersGator)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "GatorRegkey"
                        ReDim Preserve sRegKeysGator(UBound(sRegKeysGator) + 1)
                        sRegKeysGator(UBound(sRegKeysGator)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "GatorRegval"
                        ReDim Preserve sRegvalsGator(UBound(sRegvalsGator) + 1)
                        sRegvalsGator(UBound(sRegvalsGator)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "GatorProcess"
                        ReDim Preserve sProcessesGator(UBound(sProcessesGator) + 1)
                        sProcessesGator(UBound(sProcessesGator)) = Mid(sLine, InStr(sLine, "=") + 1)
                End Select
            End If
        Loop Until EOF(1)
    Close #1
    
    ReDim sFilesMyWay(0)
    ReDim sFoldersMyWay(0)
    ReDim sRegKeysMyWay(0)
    ReDim sRegvalsMyWay(0)
    ReDim sProcessesMyWay(0)
    Open sPath & "Myway.def" For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> vbNullString And InStr(sLine, "//") <> 1 Then
                Select Case Left(sLine, InStr(sLine, "=") - 1)
                    Case "MywayFile"
                        ReDim Preserve sFilesMyWay(UBound(sFilesMyWay) + 1)
                        sFilesMyWay(UBound(sFilesMyWay)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "MywayFolder"
                        ReDim Preserve sFoldersMyWay(UBound(sFoldersMyWay) + 1)
                        sFoldersMyWay(UBound(sFoldersMyWay)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "MywayRegkey"
                        ReDim Preserve sRegKeysMyWay(UBound(sRegKeysMyWay) + 1)
                        sRegKeysMyWay(UBound(sRegKeysMyWay)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "MywayRegval"
                        ReDim Preserve sRegvalsMyWay(UBound(sRegvalsMyWay) + 1)
                        sRegvalsMyWay(UBound(sRegvalsMyWay)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "MywayProcess"
                        ReDim Preserve sProcessesMyWay(UBound(sProcessesMyWay) + 1)
                        sProcessesMyWay(UBound(sProcessesMyWay)) = Mid(sLine, InStr(sLine, "=") + 1)
                End Select
            End If
        Loop Until EOF(1)
    Close #1
    
    ReDim sFilesP2P(0)
    ReDim sFoldersP2P(0)
    ReDim sRegKeysP2P(0)
    ReDim sRegvalsP2P(0)
    ReDim sProcessesP2P(0)
    Open sPath & "P2pnetworking.def" For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> vbNullString And InStr(sLine, "//") <> 1 Then
                Select Case Left(sLine, InStr(sLine, "=") - 1)
                    Case "P2pFile"
                        ReDim Preserve sFilesP2P(UBound(sFilesP2P) + 1)
                        sFilesP2P(UBound(sFilesP2P)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "P2pFolder"
                        ReDim Preserve sFoldersP2P(UBound(sFoldersP2P) + 1)
                        sFoldersP2P(UBound(sFoldersP2P)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "P2pRegkey"
                        ReDim Preserve sRegKeysP2P(UBound(sRegKeysP2P) + 1)
                        sRegKeysP2P(UBound(sRegKeysP2P)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "P2pRegval"
                        ReDim Preserve sRegvalsP2P(UBound(sRegvalsP2P) + 1)
                        sRegvalsP2P(UBound(sRegvalsP2P)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "P2pProcess"
                        ReDim Preserve sProcessesP2P(UBound(sProcessesP2P) + 1)
                        sProcessesP2P(UBound(sProcessesP2P)) = Mid(sLine, InStr(sLine, "=") + 1)
                End Select
            End If
        Loop Until EOF(1)
    Close #1
    
    ReDim sFilesPerfectNav(0)
    ReDim sFoldersPerfectNav(0)
    ReDim sRegKeysPerfectNav(0)
    ReDim sRegvalsPerfectNav(0)
    ReDim sProcessesPerfectNav(0)
    Open sPath & "Perfectnav.def" For Input As #1
        Do
            Line Input #1, sLine
            If Trim(sLine) <> vbNullString And InStr(sLine, "//") <> 1 Then
                Select Case Left(sLine, InStr(sLine, "=") - 1)
                    Case "PerfectnavFile"
                        ReDim Preserve sFilesPerfectNav(UBound(sFilesPerfectNav) + 1)
                        sFilesPerfectNav(UBound(sFilesPerfectNav)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "PerfectnavFolder"
                        ReDim Preserve sFoldersPerfectNav(UBound(sFoldersPerfectNav) + 1)
                        sFoldersPerfectNav(UBound(sFoldersPerfectNav)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "PerfectnavRegkey"
                        ReDim Preserve sRegKeysPerfectNav(UBound(sRegKeysPerfectNav) + 1)
                        sRegKeysPerfectNav(UBound(sRegKeysPerfectNav)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "PerfectnavRegval"
                        ReDim Preserve sRegvalsPerfectNav(UBound(sRegvalsPerfectNav) + 1)
                        sRegvalsPerfectNav(UBound(sRegvalsPerfectNav)) = Mid(sLine, InStr(sLine, "=") + 1)
                    Case "PerfectnavProcess"
                        ReDim Preserve sProcessesPerfectNav(UBound(sProcessesPerfectNav) + 1)
                        sProcessesPerfectNav(UBound(sProcessesPerfectNav)) = Mid(sLine, InStr(sLine, "=") + 1)
                End Select
            End If
        Loop Until EOF(1)
    Close #1
    Exit Sub
    
MissingDef:
    MsgBox "One or more of the KazaaBG definition " & _
           "files is missing. Please make sure you " & _
           "copy all files in the zip file into " & _
           "the same new folder before running KazaaBG.", vbCritical
    End
End Sub

Public Sub GetWindowsInfo()
    Dim uOVI As OSVERSIONINFO
    sWinDir = String(255, 0)
    sWinDir = Left(sWinDir, GetWindowsDirectory(sWinDir, 255))
    sWinSysDir = String(260, 0)
    sWinSysDir = Left(sWinSysDir, GetSystemDirectory(sWinSysDir, 260))
    sTempDir = String(260, 0)
    sTempDir = Left(sTempDir, GetTempPath(Len(sTempDir), sTempDir) - 1)
    
    uOVI.dwOSVersionInfoSize = Len(uOVI)
    GetVersionEx uOVI
    If uOVI.dwPlatformId = VER_PLATFORM_WIN32_NT Then bIsWinNT = True
End Sub

Public Sub ExpandVars()
    Dim sStartMenu$, sProgramFiles$, sKazaaFolder$, sSysDrive$
    Dim sDesktop$, sUserName$, sApplicData$, sPrograms$
    Dim sNameSpaceCatalog$, sProtocolCatalog$, i%
    Dim sAllUsersAppData$, sAllUsersPrograms$, sQuickLaunch$, sAllUsersDesktop$
    Const sRegShellFolder$ = "Software\Microsoft\Windows\CurrentVersion\explorer\Shell Folders"
    Status "Expanding variables in regkeys/files/folders..."
    
    sStartMenu = RegGetString(HKEY_CURRENT_USER, sRegShellFolder, "Programs", "C:\Windows\Start Menu\Programs")
    sDesktop = RegGetString(HKEY_CURRENT_USER, sRegShellFolder, "Desktop", "C:\Windows\Desktop")
    sProgramFiles = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion", "ProgramFilesDir", "C:\Program Files")
    sApplicData = RegGetString(HKEY_CURRENT_USER, sRegShellFolder, "AppData", "C:\Windows\Application Data")
    sPrograms = RegGetString(HKEY_CURRENT_USER, sRegShellFolder, "Programs")
    sQuickLaunch = sApplicData & "\Microsoft\Internet Explorer\Quick Launch"
    
    sAllUsersAppData = RegGetString(HKEY_LOCAL_MACHINE, sRegShellFolder, "AppData", "C:\Windows\All Users\Application Data")
    sAllUsersPrograms = RegGetString(HKEY_LOCAL_MACHINE, sRegShellFolder, "Programs")
    sAllUsersDesktop = RegGetString(HKEY_LOCAL_MACHINE, sRegShellFolder, "Desktop")
    
    sKazaaFolder = RegGetString(HKEY_LOCAL_MACHINE, "Software\KAZAA\CloudLoad", "ExeDir", sProgramFiles & "\Kazaa")
    If InStr(sKazaaFolder, ".exe") > 0 Then
        sKazaaFolder = Left(sKazaaFolder, InStrRev(sKazaaFolder, "\") - 1)
    End If
    If sKazaaFolder = sProgramFiles & "\Kazaa" Then
        sKazaaFolder = RegGetString(HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "KAZAA", sProgramFiles & "\Kazaa")
        If sKazaaFolder <> sProgramFiles & "\Kazaa" Then
            'C:\PROGRAM FILES\KAZAA\KAZAA.EXE /SYSTRAY
            sKazaaFolder = Left(sKazaaFolder, InStrRev(sKazaaFolder, "\") - 1)
        End If
    End If
    
    sNameSpaceCatalog = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\WinSock2\Parameters", "Current_NameSpace_Catalog", "NameSpace_Catalog5")
    sProtocolCatalog = RegGetString(HKEY_LOCAL_MACHINE, "System\CurrentControlSet\Services\WinSock2\Parameters", "Current_Protocol_Catalog", "Protocol_Catalog9")
        
    sUserName = String(255, 0)
    GetUserName sUserName, 255
    sUserName = Left(sUserName, InStr(sUserName, Chr(0)) - 1)
    sSysDrive = Left(sWinDir, 2)
    
    ' --------- KAZAA ------------------------
    For i = 1 To UBound(sFoldersKazaa)
        If sFoldersKazaa(i) = vbNullString Then Exit For
        sFoldersKazaa(i) = Replace(sFoldersKazaa(i), "%PROGRAMFILES%", sProgramFiles)
        sFoldersKazaa(i) = Replace(sFoldersKazaa(i), "%KAZAAFOLDER%", sKazaaFolder)
        sFoldersKazaa(i) = Replace(sFoldersKazaa(i), "%STARTMENU%", sStartMenu)
        sFoldersKazaa(i) = Replace(sFoldersKazaa(i), "%WINDIR%", sWinDir)
        sFoldersKazaa(i) = Replace(sFoldersKazaa(i), "%WINSYSDIR%", sWinSysDir)
        sFoldersKazaa(i) = Replace(sFoldersKazaa(i), "%DESKTOP%", sDesktop)
        sFoldersKazaa(i) = Replace(sFoldersKazaa(i), "%USERNAME%", sUserName)
        sFoldersKazaa(i) = Replace(sFoldersKazaa(i), "%APPLICDATA%", sApplicData)
        sFoldersKazaa(i) = Replace(sFoldersKazaa(i), "%PROGRAMS%", sPrograms)
        sFoldersKazaa(i) = Replace(sFoldersKazaa(i), "%QUICKLAUNCH%", sQuickLaunch)
    Next i
    For i = 1 To UBound(sFilesKazaa)
        If sFilesKazaa(i) = vbNullString Then Exit For
        sFilesKazaa(i) = Replace(sFilesKazaa(i), "%PROGRAMFILES%", sProgramFiles)
        sFilesKazaa(i) = Replace(sFilesKazaa(i), "%KAZAAFOLDER%", sKazaaFolder)
        sFilesKazaa(i) = Replace(sFilesKazaa(i), "%STARTMENU%", sStartMenu)
        sFilesKazaa(i) = Replace(sFilesKazaa(i), "%WINDIR%", sWinDir)
        sFilesKazaa(i) = Replace(sFilesKazaa(i), "%WINSYSDIR%", sWinSysDir)
        sFilesKazaa(i) = Replace(sFilesKazaa(i), "%DESKTOP%", sDesktop)
        sFilesKazaa(i) = Replace(sFilesKazaa(i), "%USERNAME%", sUserName)
        sFilesKazaa(i) = Replace(sFilesKazaa(i), "%APPLICDATA%", sApplicData)
        sFilesKazaa(i) = Replace(sFilesKazaa(i), "%PROGRAMS%", sPrograms)
        sFilesKazaa(i) = Replace(sFilesKazaa(i), "%QUICKLAUNCH%", sQuickLaunch)
    Next i
    For i = 1 To UBound(sProcessesKazaa)
        If sProcessesKazaa(i) = vbNullString Then Exit For
        If InStr(sProcessesKazaa(i), "%PROGRAMFILES%") > 0 Then sProcessesKazaa(i) = Replace(sProcessesKazaa(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sProcessesKazaa(i), "%KAZAAFOLDER%") > 0 Then sProcessesKazaa(i) = Replace(sProcessesKazaa(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sProcessesKazaa(i), "%STARTMENU%") > 0 Then sProcessesKazaa(i) = Replace(sProcessesKazaa(i), "%STARTMENU%", sStartMenu)
        If InStr(sProcessesKazaa(i), "%WINDIR%") > 0 Then sProcessesKazaa(i) = Replace(sProcessesKazaa(i), "%WINDIR%", sWinDir)
        If InStr(sProcessesKazaa(i), "%WINSYSDIR%") > 0 Then sProcessesKazaa(i) = Replace(sProcessesKazaa(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sProcessesKazaa(i), "%DESKTOP%") > 0 Then sProcessesKazaa(i) = Replace(sProcessesKazaa(i), "%DESKTOP%", sDesktop)
        If InStr(sProcessesKazaa(i), "%USERNAME%") > 0 Then sProcessesKazaa(i) = Replace(sProcessesKazaa(i), "%USERNAME%", sUserName)
        If InStr(sProcessesKazaa(i), "%APPLICDATA%") > 0 Then sProcessesKazaa(i) = Replace(sProcessesKazaa(i), "%APPLICDATA%", sApplicData)
    Next i
    
    ' --------- BDE ------------------------
    For i = 1 To UBound(sFoldersBDE)
        If sFoldersBDE(i) = vbNullString Then Exit For
        If InStr(sFoldersBDE(i), "%PROGRAMFILES%") > 0 Then sFoldersBDE(i) = Replace(sFoldersBDE(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFoldersBDE(i), "%KAZAAFOLDER%") > 0 Then sFoldersBDE(i) = Replace(sFoldersBDE(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFoldersBDE(i), "%STARTMENU%") > 0 Then sFoldersBDE(i) = Replace(sFoldersBDE(i), "%STARTMENU%", sStartMenu)
        If InStr(sFoldersBDE(i), "%WINDIR%") > 0 Then sFoldersBDE(i) = Replace(sFoldersBDE(i), "%WINDIR%", sWinDir)
        If InStr(sFoldersBDE(i), "%WINSYSDIR%") > 0 Then sFoldersBDE(i) = Replace(sFoldersBDE(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFoldersBDE(i), "%DESKTOP%") > 0 Then sFoldersBDE(i) = Replace(sFoldersBDE(i), "%DESKTOP%", sDesktop)
        If InStr(sFoldersBDE(i), "%USERNAME%") > 0 Then sFoldersBDE(i) = Replace(sFoldersBDE(i), "%USERNAME%", sUserName)
        If InStr(sFoldersBDE(i), "%APPLICDATA%") > 0 Then sFoldersBDE(i) = Replace(sFoldersBDE(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sFilesBDE)
        If sFilesBDE(i) = vbNullString Then Exit For
        If InStr(sFilesBDE(i), "%PROGRAMFILES%") > 0 Then sFilesBDE(i) = Replace(sFilesBDE(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFilesBDE(i), "%KAZAAFOLDER%") > 0 Then sFilesBDE(i) = Replace(sFilesBDE(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFilesBDE(i), "%STARTMENU%") > 0 Then sFilesBDE(i) = Replace(sFilesBDE(i), "%STARTMENU%", sStartMenu)
        If InStr(sFilesBDE(i), "%WINDIR%") > 0 Then sFilesBDE(i) = Replace(sFilesBDE(i), "%WINDIR%", sWinDir)
        If InStr(sFilesBDE(i), "%WINSYSDIR%") > 0 Then sFilesBDE(i) = Replace(sFilesBDE(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFilesBDE(i), "%DESKTOP%") > 0 Then sFilesBDE(i) = Replace(sFilesBDE(i), "%DESKTOP%", sDesktop)
        If InStr(sFilesBDE(i), "%USERNAME%") > 0 Then sFilesBDE(i) = Replace(sFilesBDE(i), "%USERNAME%", sUserName)
        If InStr(sFilesBDE(i), "%APPLICDATA%") > 0 Then sFilesBDE(i) = Replace(sFilesBDE(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sProcessesBDE)
        If sProcessesBDE(i) = vbNullString Then Exit For
        If InStr(sProcessesBDE(i), "%PROGRAMFILES%") > 0 Then sProcessesBDE(i) = Replace(sProcessesBDE(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sProcessesBDE(i), "%KAZAAFOLDER%") > 0 Then sProcessesBDE(i) = Replace(sProcessesBDE(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sProcessesBDE(i), "%STARTMENU%") > 0 Then sProcessesBDE(i) = Replace(sProcessesBDE(i), "%STARTMENU%", sStartMenu)
        If InStr(sProcessesBDE(i), "%WINDIR%") > 0 Then sProcessesBDE(i) = Replace(sProcessesBDE(i), "%WINDIR%", sWinDir)
        If InStr(sProcessesBDE(i), "%WINSYSDIR%") > 0 Then sProcessesBDE(i) = Replace(sProcessesBDE(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sProcessesBDE(i), "%DESKTOP%") > 0 Then sProcessesBDE(i) = Replace(sProcessesBDE(i), "%DESKTOP%", sDesktop)
        If InStr(sProcessesBDE(i), "%USERNAME%") > 0 Then sProcessesBDE(i) = Replace(sProcessesBDE(i), "%USERNAME%", sUserName)
        If InStr(sProcessesBDE(i), "%APPLICDATA%") > 0 Then sProcessesBDE(i) = Replace(sProcessesBDE(i), "%APPLICDATA%", sApplicData)
    Next i
    
    ' --------- CYDOOR ------------------------
    For i = 1 To UBound(sFoldersCyDoor)
        If sFoldersCyDoor(i) = vbNullString Then Exit For
        If InStr(sFoldersCyDoor(i), "%PROGRAMFILES%") > 0 Then sFoldersCyDoor(i) = Replace(sFoldersCyDoor(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFoldersCyDoor(i), "%KAZAAFOLDER%") > 0 Then sFoldersCyDoor(i) = Replace(sFoldersCyDoor(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFoldersCyDoor(i), "%STARTMENU%") > 0 Then sFoldersCyDoor(i) = Replace(sFoldersCyDoor(i), "%STARTMENU%", sStartMenu)
        If InStr(sFoldersCyDoor(i), "%WINDIR%") > 0 Then sFoldersCyDoor(i) = Replace(sFoldersCyDoor(i), "%WINDIR%", sWinDir)
        If InStr(sFoldersCyDoor(i), "%WINSYSDIR%") > 0 Then sFoldersCyDoor(i) = Replace(sFoldersCyDoor(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFoldersCyDoor(i), "%DESKTOP%") > 0 Then sFoldersCyDoor(i) = Replace(sFoldersCyDoor(i), "%DESKTOP%", sDesktop)
        If InStr(sFoldersCyDoor(i), "%USERNAME%") > 0 Then sFoldersCyDoor(i) = Replace(sFoldersCyDoor(i), "%USERNAME%", sUserName)
        If InStr(sFoldersCyDoor(i), "%APPLICDATA%") > 0 Then sFoldersCyDoor(i) = Replace(sFoldersCyDoor(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sFilesCyDoor)
        If sFilesCyDoor(i) = vbNullString Then Exit For
        If InStr(sFilesCyDoor(i), "%PROGRAMFILES%") > 0 Then sFilesCyDoor(i) = Replace(sFilesCyDoor(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFilesCyDoor(i), "%KAZAAFOLDER%") > 0 Then sFilesCyDoor(i) = Replace(sFilesCyDoor(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFilesCyDoor(i), "%STARTMENU%") > 0 Then sFilesCyDoor(i) = Replace(sFilesCyDoor(i), "%STARTMENU%", sStartMenu)
        If InStr(sFilesCyDoor(i), "%WINDIR%") > 0 Then sFilesCyDoor(i) = Replace(sFilesCyDoor(i), "%WINDIR%", sWinDir)
        If InStr(sFilesCyDoor(i), "%WINSYSDIR%") > 0 Then sFilesCyDoor(i) = Replace(sFilesCyDoor(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFilesCyDoor(i), "%DESKTOP%") > 0 Then sFilesCyDoor(i) = Replace(sFilesCyDoor(i), "%DESKTOP%", sDesktop)
        If InStr(sFilesCyDoor(i), "%USERNAME%") > 0 Then sFilesCyDoor(i) = Replace(sFilesCyDoor(i), "%USERNAME%", sUserName)
        If InStr(sFilesCyDoor(i), "%APPLICDATA%") > 0 Then sFilesCyDoor(i) = Replace(sFilesCyDoor(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sProcessesCyDoor)
        If sProcessesCyDoor(i) = vbNullString Then Exit For
        If InStr(sProcessesCyDoor(i), "%PROGRAMFILES%") > 0 Then sProcessesCyDoor(i) = Replace(sProcessesCyDoor(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sProcessesCyDoor(i), "%KAZAAFOLDER%") > 0 Then sProcessesCyDoor(i) = Replace(sProcessesCyDoor(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sProcessesCyDoor(i), "%STARTMENU%") > 0 Then sProcessesCyDoor(i) = Replace(sProcessesCyDoor(i), "%STARTMENU%", sStartMenu)
        If InStr(sProcessesCyDoor(i), "%WINDIR%") > 0 Then sProcessesCyDoor(i) = Replace(sProcessesCyDoor(i), "%WINDIR%", sWinDir)
        If InStr(sProcessesCyDoor(i), "%WINSYSDIR%") > 0 Then sProcessesCyDoor(i) = Replace(sProcessesCyDoor(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sProcessesCyDoor(i), "%DESKTOP%") > 0 Then sProcessesCyDoor(i) = Replace(sProcessesCyDoor(i), "%DESKTOP%", sDesktop)
        If InStr(sProcessesCyDoor(i), "%USERNAME%") > 0 Then sProcessesCyDoor(i) = Replace(sProcessesCyDoor(i), "%USERNAME%", sUserName)
        If InStr(sProcessesCyDoor(i), "%APPLICDATA%") > 0 Then sProcessesCyDoor(i) = Replace(sProcessesCyDoor(i), "%APPLICDATA%", sApplicData)
    Next i
    
    ' --------- COMMONNAME ------------------------
    For i = 1 To UBound(sFoldersCommonName)
        If sFoldersCommonName(i) = vbNullString Then Exit For
        If InStr(sFoldersCommonName(i), "%PROGRAMFILES%") > 0 Then sFoldersCommonName(i) = Replace(sFoldersCommonName(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFoldersCommonName(i), "%KAZAAFOLDER%") > 0 Then sFoldersCommonName(i) = Replace(sFoldersCommonName(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFoldersCommonName(i), "%STARTMENU%") > 0 Then sFoldersCommonName(i) = Replace(sFoldersCommonName(i), "%STARTMENU%", sStartMenu)
        If InStr(sFoldersCommonName(i), "%WINDIR%") > 0 Then sFoldersCommonName(i) = Replace(sFoldersCommonName(i), "%WINDIR%", sWinDir)
        If InStr(sFoldersCommonName(i), "%WINSYSDIR%") > 0 Then sFoldersCommonName(i) = Replace(sFoldersCommonName(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFoldersCommonName(i), "%DESKTOP%") > 0 Then sFoldersCommonName(i) = Replace(sFoldersCommonName(i), "%DESKTOP%", sDesktop)
        If InStr(sFoldersCommonName(i), "%USERNAME%") > 0 Then sFoldersCommonName(i) = Replace(sFoldersCommonName(i), "%USERNAME%", sUserName)
        If InStr(sFoldersCommonName(i), "%APPLICDATA%") > 0 Then sFoldersCommonName(i) = Replace(sFoldersCommonName(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sFilesCommonName)
        If sFilesCommonName(i) = vbNullString Then Exit For
        If InStr(sFilesCommonName(i), "%PROGRAMFILES%") > 0 Then sFilesCommonName(i) = Replace(sFilesCommonName(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFilesCommonName(i), "%KAZAAFOLDER%") > 0 Then sFilesCommonName(i) = Replace(sFilesCommonName(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFilesCommonName(i), "%STARTMENU%") > 0 Then sFilesCommonName(i) = Replace(sFilesCommonName(i), "%STARTMENU%", sStartMenu)
        If InStr(sFilesCommonName(i), "%WINDIR%") > 0 Then sFilesCommonName(i) = Replace(sFilesCommonName(i), "%WINDIR%", sWinDir)
        If InStr(sFilesCommonName(i), "%WINSYSDIR%") > 0 Then sFilesCommonName(i) = Replace(sFilesCommonName(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFilesCommonName(i), "%DESKTOP%") > 0 Then sFilesCommonName(i) = Replace(sFilesCommonName(i), "%DESKTOP%", sDesktop)
        If InStr(sFilesCommonName(i), "%USERNAME%") > 0 Then sFilesCommonName(i) = Replace(sFilesCommonName(i), "%USERNAME%", sUserName)
        If InStr(sFilesCommonName(i), "%APPLICDATA%") > 0 Then sFilesCommonName(i) = Replace(sFilesCommonName(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sProcessesCommonName)
        If sProcessesCommonName(i) = vbNullString Then Exit For
        If InStr(sProcessesCommonName(i), "%PROGRAMFILES%") > 0 Then sProcessesCommonName(i) = Replace(sProcessesCommonName(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sProcessesCommonName(i), "%KAZAAFOLDER%") > 0 Then sProcessesCommonName(i) = Replace(sProcessesCommonName(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sProcessesCommonName(i), "%STARTMENU%") > 0 Then sProcessesCommonName(i) = Replace(sProcessesCommonName(i), "%STARTMENU%", sStartMenu)
        If InStr(sProcessesCommonName(i), "%WINDIR%") > 0 Then sProcessesCommonName(i) = Replace(sProcessesCommonName(i), "%WINDIR%", sWinDir)
        If InStr(sProcessesCommonName(i), "%WINSYSDIR%") > 0 Then sProcessesCommonName(i) = Replace(sProcessesCommonName(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sProcessesCommonName(i), "%DESKTOP%") > 0 Then sProcessesCommonName(i) = Replace(sProcessesCommonName(i), "%DESKTOP%", sDesktop)
        If InStr(sProcessesCommonName(i), "%USERNAME%") > 0 Then sProcessesCommonName(i) = Replace(sProcessesCommonName(i), "%USERNAME%", sUserName)
        If InStr(sProcessesCommonName(i), "%APPLICDATA%") > 0 Then sProcessesCommonName(i) = Replace(sProcessesCommonName(i), "%APPLICDATA%", sApplicData)
    Next i
    
    ' --------- NEWDOTNET ------------------------
    For i = 1 To UBound(sFoldersNewDotNet)
        If sFoldersNewDotNet(i) = vbNullString Then Exit For
        If InStr(sFoldersNewDotNet(i), "%PROGRAMFILES%") > 0 Then sFoldersNewDotNet(i) = Replace(sFoldersNewDotNet(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFoldersNewDotNet(i), "%KAZAAFOLDER%") > 0 Then sFoldersNewDotNet(i) = Replace(sFoldersNewDotNet(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFoldersNewDotNet(i), "%STARTMENU%") > 0 Then sFoldersNewDotNet(i) = Replace(sFoldersNewDotNet(i), "%STARTMENU%", sStartMenu)
        If InStr(sFoldersNewDotNet(i), "%WINDIR%") > 0 Then sFoldersNewDotNet(i) = Replace(sFoldersNewDotNet(i), "%WINDIR%", sWinDir)
        If InStr(sFoldersNewDotNet(i), "%WINSYSDIR%") > 0 Then sFoldersNewDotNet(i) = Replace(sFoldersNewDotNet(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFoldersNewDotNet(i), "%DESKTOP%") > 0 Then sFoldersNewDotNet(i) = Replace(sFoldersNewDotNet(i), "%DESKTOP%", sDesktop)
        If InStr(sFoldersNewDotNet(i), "%USERNAME%") > 0 Then sFoldersNewDotNet(i) = Replace(sFoldersNewDotNet(i), "%USERNAME%", sUserName)
        If InStr(sFoldersNewDotNet(i), "%APPLICDATA%") > 0 Then sFoldersNewDotNet(i) = Replace(sFoldersNewDotNet(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sFilesNewDotNet)
        If sFilesNewDotNet(i) = vbNullString Then Exit For
        If InStr(sFilesNewDotNet(i), "%PROGRAMFILES%") > 0 Then sFilesNewDotNet(i) = Replace(sFilesNewDotNet(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFilesNewDotNet(i), "%KAZAAFOLDER%") > 0 Then sFilesNewDotNet(i) = Replace(sFilesNewDotNet(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFilesNewDotNet(i), "%STARTMENU%") > 0 Then sFilesNewDotNet(i) = Replace(sFilesNewDotNet(i), "%STARTMENU%", sStartMenu)
        If InStr(sFilesNewDotNet(i), "%WINDIR%") > 0 Then sFilesNewDotNet(i) = Replace(sFilesNewDotNet(i), "%WINDIR%", sWinDir)
        If InStr(sFilesNewDotNet(i), "%WINSYSDIR%") > 0 Then sFilesNewDotNet(i) = Replace(sFilesNewDotNet(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFilesNewDotNet(i), "%DESKTOP%") > 0 Then sFilesNewDotNet(i) = Replace(sFilesNewDotNet(i), "%DESKTOP%", sDesktop)
        If InStr(sFilesNewDotNet(i), "%USERNAME%") > 0 Then sFilesNewDotNet(i) = Replace(sFilesNewDotNet(i), "%USERNAME%", sUserName)
        If InStr(sFilesNewDotNet(i), "%APPLICDATA%") > 0 Then sFilesNewDotNet(i) = Replace(sFilesNewDotNet(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sProcessesNewDotNet)
        If sProcessesNewDotNet(i) = vbNullString Then Exit For
        If InStr(sProcessesNewDotNet(i), "%PROGRAMFILES%") > 0 Then sProcessesNewDotNet(i) = Replace(sProcessesNewDotNet(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sProcessesNewDotNet(i), "%KAZAAFOLDER%") > 0 Then sProcessesNewDotNet(i) = Replace(sProcessesNewDotNet(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sProcessesNewDotNet(i), "%STARTMENU%") > 0 Then sProcessesNewDotNet(i) = Replace(sProcessesNewDotNet(i), "%STARTMENU%", sStartMenu)
        If InStr(sProcessesNewDotNet(i), "%WINDIR%") > 0 Then sProcessesNewDotNet(i) = Replace(sProcessesNewDotNet(i), "%WINDIR%", sWinDir)
        If InStr(sProcessesNewDotNet(i), "%WINSYSDIR%") > 0 Then sProcessesNewDotNet(i) = Replace(sProcessesNewDotNet(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sProcessesNewDotNet(i), "%DESKTOP%") > 0 Then sProcessesNewDotNet(i) = Replace(sProcessesNewDotNet(i), "%DESKTOP%", sDesktop)
        If InStr(sProcessesNewDotNet(i), "%USERNAME%") > 0 Then sProcessesNewDotNet(i) = Replace(sProcessesNewDotNet(i), "%USERNAME%", sUserName)
        If InStr(sProcessesNewDotNet(i), "%APPLICDATA%") > 0 Then sProcessesNewDotNet(i) = Replace(sProcessesNewDotNet(i), "%APPLICDATA%", sApplicData)
    Next i
    
    ' --------- WEBHANCER ------------------------
    For i = 1 To UBound(sFoldersWebHancer)
        If sFoldersWebHancer(i) = vbNullString Then Exit For
        If InStr(sFoldersWebHancer(i), "%PROGRAMFILES%") > 0 Then sFoldersWebHancer(i) = Replace(sFoldersWebHancer(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFoldersWebHancer(i), "%KAZAAFOLDER%") > 0 Then sFoldersWebHancer(i) = Replace(sFoldersWebHancer(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFoldersWebHancer(i), "%STARTMENU%") > 0 Then sFoldersWebHancer(i) = Replace(sFoldersWebHancer(i), "%STARTMENU%", sStartMenu)
        If InStr(sFoldersWebHancer(i), "%WINDIR%") > 0 Then sFoldersWebHancer(i) = Replace(sFoldersWebHancer(i), "%WINDIR%", sWinDir)
        If InStr(sFoldersWebHancer(i), "%WINSYSDIR%") > 0 Then sFoldersWebHancer(i) = Replace(sFoldersWebHancer(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFoldersWebHancer(i), "%DESKTOP%") > 0 Then sFoldersWebHancer(i) = Replace(sFoldersWebHancer(i), "%DESKTOP%", sDesktop)
        If InStr(sFoldersWebHancer(i), "%USERNAME%") > 0 Then sFoldersWebHancer(i) = Replace(sFoldersWebHancer(i), "%USERNAME%", sUserName)
        If InStr(sFoldersWebHancer(i), "%APPLICDATA%") > 0 Then sFoldersWebHancer(i) = Replace(sFoldersWebHancer(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sFilesWebHancer)
        If sFilesWebHancer(i) = vbNullString Then Exit For
        If InStr(sFilesWebHancer(i), "%PROGRAMFILES%") > 0 Then sFilesWebHancer(i) = Replace(sFilesWebHancer(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFilesWebHancer(i), "%KAZAAFOLDER%") > 0 Then sFilesWebHancer(i) = Replace(sFilesWebHancer(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFilesWebHancer(i), "%STARTMENU%") > 0 Then sFilesWebHancer(i) = Replace(sFilesWebHancer(i), "%STARTMENU%", sStartMenu)
        If InStr(sFilesWebHancer(i), "%WINDIR%") > 0 Then sFilesWebHancer(i) = Replace(sFilesWebHancer(i), "%WINDIR%", sWinDir)
        If InStr(sFilesWebHancer(i), "%WINSYSDIR%") > 0 Then sFilesWebHancer(i) = Replace(sFilesWebHancer(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFilesWebHancer(i), "%DESKTOP%") > 0 Then sFilesWebHancer(i) = Replace(sFilesWebHancer(i), "%DESKTOP%", sDesktop)
        If InStr(sFilesWebHancer(i), "%USERNAME%") > 0 Then sFilesWebHancer(i) = Replace(sFilesWebHancer(i), "%USERNAME%", sUserName)
        If InStr(sFilesWebHancer(i), "%APPLICDATA%") > 0 Then sFilesWebHancer(i) = Replace(sFilesWebHancer(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sProcessesWebHancer)
        If sProcessesWebHancer(i) = vbNullString Then Exit For
        If InStr(sProcessesWebHancer(i), "%PROGRAMFILES%") > 0 Then sProcessesWebHancer(i) = Replace(sProcessesWebHancer(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sProcessesWebHancer(i), "%KAZAAFOLDER%") > 0 Then sProcessesWebHancer(i) = Replace(sProcessesWebHancer(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sProcessesWebHancer(i), "%STARTMENU%") > 0 Then sProcessesWebHancer(i) = Replace(sProcessesWebHancer(i), "%STARTMENU%", sStartMenu)
        If InStr(sProcessesWebHancer(i), "%WINDIR%") > 0 Then sProcessesWebHancer(i) = Replace(sProcessesWebHancer(i), "%WINDIR%", sWinDir)
        If InStr(sProcessesWebHancer(i), "%WINSYSDIR%") > 0 Then sProcessesWebHancer(i) = Replace(sProcessesWebHancer(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sProcessesWebHancer(i), "%DESKTOP%") > 0 Then sProcessesWebHancer(i) = Replace(sProcessesWebHancer(i), "%DESKTOP%", sDesktop)
        If InStr(sProcessesWebHancer(i), "%USERNAME%") > 0 Then sProcessesWebHancer(i) = Replace(sProcessesWebHancer(i), "%USERNAME%", sUserName)
        If InStr(sProcessesWebHancer(i), "%APPLICDATA%") > 0 Then sProcessesWebHancer(i) = Replace(sProcessesWebHancer(i), "%APPLICDATA%", sApplicData)
    Next i
    
    ' --------- MEDIALOADS ------------------------
    For i = 1 To UBound(sFoldersMediaLoads)
        If sFoldersMediaLoads(i) = vbNullString Then Exit For
        If InStr(sFoldersMediaLoads(i), "%PROGRAMFILES%") > 0 Then sFoldersMediaLoads(i) = Replace(sFoldersMediaLoads(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFoldersMediaLoads(i), "%KAZAAFOLDER%") > 0 Then sFoldersMediaLoads(i) = Replace(sFoldersMediaLoads(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFoldersMediaLoads(i), "%STARTMENU%") > 0 Then sFoldersMediaLoads(i) = Replace(sFoldersMediaLoads(i), "%STARTMENU%", sStartMenu)
        If InStr(sFoldersMediaLoads(i), "%WINDIR%") > 0 Then sFoldersMediaLoads(i) = Replace(sFoldersMediaLoads(i), "%WINDIR%", sWinDir)
        If InStr(sFoldersMediaLoads(i), "%WINSYSDIR%") > 0 Then sFoldersMediaLoads(i) = Replace(sFoldersMediaLoads(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFoldersMediaLoads(i), "%DESKTOP%") > 0 Then sFoldersMediaLoads(i) = Replace(sFoldersMediaLoads(i), "%DESKTOP%", sDesktop)
        If InStr(sFoldersMediaLoads(i), "%USERNAME%") > 0 Then sFoldersMediaLoads(i) = Replace(sFoldersMediaLoads(i), "%USERNAME%", sUserName)
        If InStr(sFoldersMediaLoads(i), "%APPLICDATA%") > 0 Then sFoldersMediaLoads(i) = Replace(sFoldersMediaLoads(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sFilesMediaLoads)
        If sFilesMediaLoads(i) = vbNullString Then Exit For
        If InStr(sFilesMediaLoads(i), "%PROGRAMFILES%") > 0 Then sFilesMediaLoads(i) = Replace(sFilesMediaLoads(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFilesMediaLoads(i), "%KAZAAFOLDER%") > 0 Then sFilesMediaLoads(i) = Replace(sFilesMediaLoads(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFilesMediaLoads(i), "%STARTMENU%") > 0 Then sFilesMediaLoads(i) = Replace(sFilesMediaLoads(i), "%STARTMENU%", sStartMenu)
        If InStr(sFilesMediaLoads(i), "%WINDIR%") > 0 Then sFilesMediaLoads(i) = Replace(sFilesMediaLoads(i), "%WINDIR%", sWinDir)
        If InStr(sFilesMediaLoads(i), "%WINSYSDIR%") > 0 Then sFilesMediaLoads(i) = Replace(sFilesMediaLoads(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFilesMediaLoads(i), "%DESKTOP%") > 0 Then sFilesMediaLoads(i) = Replace(sFilesMediaLoads(i), "%DESKTOP%", sDesktop)
        If InStr(sFilesMediaLoads(i), "%USERNAME%") > 0 Then sFilesMediaLoads(i) = Replace(sFilesMediaLoads(i), "%USERNAME%", sUserName)
        If InStr(sFilesMediaLoads(i), "%APPLICDATA%") > 0 Then sFilesMediaLoads(i) = Replace(sFilesMediaLoads(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sProcessesMediaLoads)
        If sProcessesMediaLoads(i) = vbNullString Then Exit For
        If InStr(sProcessesMediaLoads(i), "%PROGRAMFILES%") > 0 Then sProcessesMediaLoads(i) = Replace(sProcessesMediaLoads(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sProcessesMediaLoads(i), "%KAZAAFOLDER%") > 0 Then sProcessesMediaLoads(i) = Replace(sProcessesMediaLoads(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sProcessesMediaLoads(i), "%STARTMENU%") > 0 Then sProcessesMediaLoads(i) = Replace(sProcessesMediaLoads(i), "%STARTMENU%", sStartMenu)
        If InStr(sProcessesMediaLoads(i), "%WINDIR%") > 0 Then sProcessesMediaLoads(i) = Replace(sProcessesMediaLoads(i), "%WINDIR%", sWinDir)
        If InStr(sProcessesMediaLoads(i), "%WINSYSDIR%") > 0 Then sProcessesMediaLoads(i) = Replace(sProcessesMediaLoads(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sProcessesMediaLoads(i), "%DESKTOP%") > 0 Then sProcessesMediaLoads(i) = Replace(sProcessesMediaLoads(i), "%DESKTOP%", sDesktop)
        If InStr(sProcessesMediaLoads(i), "%USERNAME%") > 0 Then sProcessesMediaLoads(i) = Replace(sProcessesMediaLoads(i), "%USERNAME%", sUserName)
        If InStr(sProcessesMediaLoads(i), "%APPLICDATA%") > 0 Then sProcessesMediaLoads(i) = Replace(sProcessesMediaLoads(i), "%APPLICDATA%", sApplicData)
    Next i
    
    ' --------- SAVENOW ------------------------
    For i = 1 To UBound(sFoldersSaveNow)
        If sFoldersSaveNow(i) = vbNullString Then Exit For
        If InStr(sFoldersSaveNow(i), "%PROGRAMFILES%") > 0 Then sFoldersSaveNow(i) = Replace(sFoldersSaveNow(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFoldersSaveNow(i), "%KAZAAFOLDER%") > 0 Then sFoldersSaveNow(i) = Replace(sFoldersSaveNow(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFoldersSaveNow(i), "%STARTMENU%") > 0 Then sFoldersSaveNow(i) = Replace(sFoldersSaveNow(i), "%STARTMENU%", sStartMenu)
        If InStr(sFoldersSaveNow(i), "%WINDIR%") > 0 Then sFoldersSaveNow(i) = Replace(sFoldersSaveNow(i), "%WINDIR%", sWinDir)
        If InStr(sFoldersSaveNow(i), "%WINSYSDIR%") > 0 Then sFoldersSaveNow(i) = Replace(sFoldersSaveNow(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFoldersSaveNow(i), "%DESKTOP%") > 0 Then sFoldersSaveNow(i) = Replace(sFoldersSaveNow(i), "%DESKTOP%", sDesktop)
        If InStr(sFoldersSaveNow(i), "%USERNAME%") > 0 Then sFoldersSaveNow(i) = Replace(sFoldersSaveNow(i), "%USERNAME%", sUserName)
        If InStr(sFoldersSaveNow(i), "%APPLICDATA%") > 0 Then sFoldersSaveNow(i) = Replace(sFoldersSaveNow(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sFilesSaveNow)
        If sFilesSaveNow(i) = vbNullString Then Exit For
        If InStr(sFilesSaveNow(i), "%PROGRAMFILES%") > 0 Then sFilesSaveNow(i) = Replace(sFilesSaveNow(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFilesSaveNow(i), "%KAZAAFOLDER%") > 0 Then sFilesSaveNow(i) = Replace(sFilesSaveNow(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFilesSaveNow(i), "%STARTMENU%") > 0 Then sFilesSaveNow(i) = Replace(sFilesSaveNow(i), "%STARTMENU%", sStartMenu)
        If InStr(sFilesSaveNow(i), "%WINDIR%") > 0 Then sFilesSaveNow(i) = Replace(sFilesSaveNow(i), "%WINDIR%", sWinDir)
        If InStr(sFilesSaveNow(i), "%WINSYSDIR%") > 0 Then sFilesSaveNow(i) = Replace(sFilesSaveNow(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFilesSaveNow(i), "%DESKTOP%") > 0 Then sFilesSaveNow(i) = Replace(sFilesSaveNow(i), "%DESKTOP%", sDesktop)
        If InStr(sFilesSaveNow(i), "%USERNAME%") > 0 Then sFilesSaveNow(i) = Replace(sFilesSaveNow(i), "%USERNAME%", sUserName)
        If InStr(sFilesSaveNow(i), "%APPLICDATA%") > 0 Then sFilesSaveNow(i) = Replace(sFilesSaveNow(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sProcessesSaveNow)
        If sProcessesSaveNow(i) = vbNullString Then Exit For
        If InStr(sProcessesSaveNow(i), "%PROGRAMFILES%") > 0 Then sProcessesSaveNow(i) = Replace(sProcessesSaveNow(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sProcessesSaveNow(i), "%KAZAAFOLDER%") > 0 Then sProcessesSaveNow(i) = Replace(sProcessesSaveNow(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sProcessesSaveNow(i), "%STARTMENU%") > 0 Then sProcessesSaveNow(i) = Replace(sProcessesSaveNow(i), "%STARTMENU%", sStartMenu)
        If InStr(sProcessesSaveNow(i), "%WINDIR%") > 0 Then sProcessesSaveNow(i) = Replace(sProcessesSaveNow(i), "%WINDIR%", sWinDir)
        If InStr(sProcessesSaveNow(i), "%WINSYSDIR%") > 0 Then sProcessesSaveNow(i) = Replace(sProcessesSaveNow(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sProcessesSaveNow(i), "%DESKTOP%") > 0 Then sProcessesSaveNow(i) = Replace(sProcessesSaveNow(i), "%DESKTOP%", sDesktop)
        If InStr(sProcessesSaveNow(i), "%USERNAME%") > 0 Then sProcessesSaveNow(i) = Replace(sProcessesSaveNow(i), "%USERNAME%", sUserName)
        If InStr(sProcessesSaveNow(i), "%APPLICDATA%") > 0 Then sProcessesSaveNow(i) = Replace(sProcessesSaveNow(i), "%APPLICDATA%", sApplicData)
    Next i
    
    ' --------- DELFIN ------------------------
    For i = 1 To UBound(sFoldersDelfin)
        If sFoldersDelfin(i) = vbNullString Then Exit For
        If InStr(sFoldersDelfin(i), "%PROGRAMFILES%") > 0 Then sFoldersDelfin(i) = Replace(sFoldersDelfin(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFoldersDelfin(i), "%KAZAAFOLDER%") > 0 Then sFoldersDelfin(i) = Replace(sFoldersDelfin(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFoldersDelfin(i), "%STARTMENU%") > 0 Then sFoldersDelfin(i) = Replace(sFoldersDelfin(i), "%STARTMENU%", sStartMenu)
        If InStr(sFoldersDelfin(i), "%WINDIR%") > 0 Then sFoldersDelfin(i) = Replace(sFoldersDelfin(i), "%WINDIR%", sWinDir)
        If InStr(sFoldersDelfin(i), "%WINSYSDIR%") > 0 Then sFoldersDelfin(i) = Replace(sFoldersDelfin(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFoldersDelfin(i), "%DESKTOP%") > 0 Then sFoldersDelfin(i) = Replace(sFoldersDelfin(i), "%DESKTOP%", sDesktop)
        If InStr(sFoldersDelfin(i), "%USERNAME%") > 0 Then sFoldersDelfin(i) = Replace(sFoldersDelfin(i), "%USERNAME%", sUserName)
        If InStr(sFoldersDelfin(i), "%APPLICDATA%") > 0 Then sFoldersDelfin(i) = Replace(sFoldersDelfin(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sFilesDelfin)
        If sFilesDelfin(i) = vbNullString Then Exit For
        If InStr(sFilesDelfin(i), "%PROGRAMFILES%") > 0 Then sFilesDelfin(i) = Replace(sFilesDelfin(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFilesDelfin(i), "%KAZAAFOLDER%") > 0 Then sFilesDelfin(i) = Replace(sFilesDelfin(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFilesDelfin(i), "%STARTMENU%") > 0 Then sFilesDelfin(i) = Replace(sFilesDelfin(i), "%STARTMENU%", sStartMenu)
        If InStr(sFilesDelfin(i), "%WINDIR%") > 0 Then sFilesDelfin(i) = Replace(sFilesDelfin(i), "%WINDIR%", sWinDir)
        If InStr(sFilesDelfin(i), "%WINSYSDIR%") > 0 Then sFilesDelfin(i) = Replace(sFilesDelfin(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFilesDelfin(i), "%DESKTOP%") > 0 Then sFilesDelfin(i) = Replace(sFilesDelfin(i), "%DESKTOP%", sDesktop)
        If InStr(sFilesDelfin(i), "%USERNAME%") > 0 Then sFilesDelfin(i) = Replace(sFilesDelfin(i), "%USERNAME%", sUserName)
        If InStr(sFilesDelfin(i), "%APPLICDATA%") > 0 Then sFilesDelfin(i) = Replace(sFilesDelfin(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sProcessesDelfin)
        If sProcessesDelfin(i) = vbNullString Then Exit For
        If InStr(sProcessesDelfin(i), "%PROGRAMFILES%") > 0 Then sProcessesDelfin(i) = Replace(sProcessesDelfin(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sProcessesDelfin(i), "%KAZAAFOLDER%") > 0 Then sProcessesDelfin(i) = Replace(sProcessesDelfin(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sProcessesDelfin(i), "%STARTMENU%") > 0 Then sProcessesDelfin(i) = Replace(sProcessesDelfin(i), "%STARTMENU%", sStartMenu)
        If InStr(sProcessesDelfin(i), "%WINDIR%") > 0 Then sProcessesDelfin(i) = Replace(sProcessesDelfin(i), "%WINDIR%", sWinDir)
        If InStr(sProcessesDelfin(i), "%WINSYSDIR%") > 0 Then sProcessesDelfin(i) = Replace(sProcessesDelfin(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sProcessesDelfin(i), "%DESKTOP%") > 0 Then sProcessesDelfin(i) = Replace(sProcessesDelfin(i), "%DESKTOP%", sDesktop)
        If InStr(sProcessesDelfin(i), "%USERNAME%") > 0 Then sProcessesDelfin(i) = Replace(sProcessesDelfin(i), "%USERNAME%", sUserName)
        If InStr(sProcessesDelfin(i), "%APPLICDATA%") > 0 Then sProcessesDelfin(i) = Replace(sProcessesDelfin(i), "%APPLICDATA%", sApplicData)
    Next i
    
    ' --------- ONFLOW ------------------------
    For i = 1 To UBound(sFoldersOnFlow)
        If sFoldersOnFlow(i) = vbNullString Then Exit For
        If InStr(sFoldersOnFlow(i), "%PROGRAMFILES%") > 0 Then sFoldersOnFlow(i) = Replace(sFoldersOnFlow(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFoldersOnFlow(i), "%KAZAAFOLDER%") > 0 Then sFoldersOnFlow(i) = Replace(sFoldersOnFlow(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFoldersOnFlow(i), "%STARTMENU%") > 0 Then sFoldersOnFlow(i) = Replace(sFoldersOnFlow(i), "%STARTMENU%", sStartMenu)
        If InStr(sFoldersOnFlow(i), "%WINDIR%") > 0 Then sFoldersOnFlow(i) = Replace(sFoldersOnFlow(i), "%WINDIR%", sWinDir)
        If InStr(sFoldersOnFlow(i), "%WINSYSDIR%") > 0 Then sFoldersOnFlow(i) = Replace(sFoldersOnFlow(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFoldersOnFlow(i), "%DESKTOP%") > 0 Then sFoldersOnFlow(i) = Replace(sFoldersOnFlow(i), "%DESKTOP%", sDesktop)
        If InStr(sFoldersOnFlow(i), "%USERNAME%") > 0 Then sFoldersOnFlow(i) = Replace(sFoldersOnFlow(i), "%USERNAME%", sUserName)
        If InStr(sFoldersOnFlow(i), "%APPLICDATA%") > 0 Then sFoldersOnFlow(i) = Replace(sFoldersOnFlow(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sFilesOnFlow)
        If sFilesOnFlow(i) = vbNullString Then Exit For
        If InStr(sFilesOnFlow(i), "%PROGRAMFILES%") > 0 Then sFilesOnFlow(i) = Replace(sFilesOnFlow(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFilesOnFlow(i), "%KAZAAFOLDER%") > 0 Then sFilesOnFlow(i) = Replace(sFilesOnFlow(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFilesOnFlow(i), "%STARTMENU%") > 0 Then sFilesOnFlow(i) = Replace(sFilesOnFlow(i), "%STARTMENU%", sStartMenu)
        If InStr(sFilesOnFlow(i), "%WINDIR%") > 0 Then sFilesOnFlow(i) = Replace(sFilesOnFlow(i), "%WINDIR%", sWinDir)
        If InStr(sFilesOnFlow(i), "%WINSYSDIR%") > 0 Then sFilesOnFlow(i) = Replace(sFilesOnFlow(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFilesOnFlow(i), "%DESKTOP%") > 0 Then sFilesOnFlow(i) = Replace(sFilesOnFlow(i), "%DESKTOP%", sDesktop)
        If InStr(sFilesOnFlow(i), "%USERNAME%") > 0 Then sFilesOnFlow(i) = Replace(sFilesOnFlow(i), "%USERNAME%", sUserName)
        If InStr(sFilesOnFlow(i), "%APPLICDATA%") > 0 Then sFilesOnFlow(i) = Replace(sFilesOnFlow(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sProcessesOnFlow)
        If sProcessesOnFlow(i) = vbNullString Then Exit For
        If InStr(sProcessesOnFlow(i), "%PROGRAMFILES%") > 0 Then sProcessesOnFlow(i) = Replace(sProcessesOnFlow(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sProcessesOnFlow(i), "%KAZAAFOLDER%") > 0 Then sProcessesOnFlow(i) = Replace(sProcessesOnFlow(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sProcessesOnFlow(i), "%STARTMENU%") > 0 Then sProcessesOnFlow(i) = Replace(sProcessesOnFlow(i), "%STARTMENU%", sStartMenu)
        If InStr(sProcessesOnFlow(i), "%WINDIR%") > 0 Then sProcessesOnFlow(i) = Replace(sProcessesOnFlow(i), "%WINDIR%", sWinDir)
        If InStr(sProcessesOnFlow(i), "%WINSYSDIR%") > 0 Then sProcessesOnFlow(i) = Replace(sProcessesOnFlow(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sProcessesOnFlow(i), "%DESKTOP%") > 0 Then sProcessesOnFlow(i) = Replace(sProcessesOnFlow(i), "%DESKTOP%", sDesktop)
        If InStr(sProcessesOnFlow(i), "%USERNAME%") > 0 Then sProcessesOnFlow(i) = Replace(sProcessesOnFlow(i), "%USERNAME%", sUserName)
        If InStr(sProcessesOnFlow(i), "%APPLICDATA%") > 0 Then sProcessesOnFlow(i) = Replace(sProcessesOnFlow(i), "%APPLICDATA%", sApplicData)
    Next i
    
    ' --------- OTHER ------------------------
    For i = 1 To UBound(sFoldersOther)
        If sFoldersOther(i) = vbNullString Then Exit For
        If InStr(sFoldersOther(i), "%PROGRAMFILES%") > 0 Then sFoldersOther(i) = Replace(sFoldersOther(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFoldersOther(i), "%KAZAAFOLDER%") > 0 Then sFoldersOther(i) = Replace(sFoldersOther(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFoldersOther(i), "%STARTMENU%") > 0 Then sFoldersOther(i) = Replace(sFoldersOther(i), "%STARTMENU%", sStartMenu)
        If InStr(sFoldersOther(i), "%WINDIR%") > 0 Then sFoldersOther(i) = Replace(sFoldersOther(i), "%WINDIR%", sWinDir)
        If InStr(sFoldersOther(i), "%WINSYSDIR%") > 0 Then sFoldersOther(i) = Replace(sFoldersOther(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFoldersOther(i), "%DESKTOP%") > 0 Then sFoldersOther(i) = Replace(sFoldersOther(i), "%DESKTOP%", sDesktop)
        If InStr(sFoldersOther(i), "%USERNAME%") > 0 Then sFoldersOther(i) = Replace(sFoldersOther(i), "%USERNAME%", sUserName)
        If InStr(sFoldersOther(i), "%APPLICDATA%") > 0 Then sFoldersOther(i) = Replace(sFoldersOther(i), "%APPLICDATA%", sApplicData)
        If InStr(sFoldersOther(i), "%TEMPDIR%") > 0 Then sFoldersOther(i) = Replace(sFoldersOther(i), "%TEMPDIR%", sTempDir)
    Next i
    For i = 1 To UBound(sFilesOther)
        If sFilesOther(i) = vbNullString Then Exit For
        If InStr(sFilesOther(i), "%PROGRAMFILES%") > 0 Then sFilesOther(i) = Replace(sFilesOther(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFilesOther(i), "%KAZAAFOLDER%") > 0 Then sFilesOther(i) = Replace(sFilesOther(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFilesOther(i), "%STARTMENU%") > 0 Then sFilesOther(i) = Replace(sFilesOther(i), "%STARTMENU%", sStartMenu)
        If InStr(sFilesOther(i), "%WINDIR%") > 0 Then sFilesOther(i) = Replace(sFilesOther(i), "%WINDIR%", sWinDir)
        If InStr(sFilesOther(i), "%WINSYSDIR%") > 0 Then sFilesOther(i) = Replace(sFilesOther(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFilesOther(i), "%DESKTOP%") > 0 Then sFilesOther(i) = Replace(sFilesOther(i), "%DESKTOP%", sDesktop)
        If InStr(sFilesOther(i), "%USERNAME%") > 0 Then sFilesOther(i) = Replace(sFilesOther(i), "%USERNAME%", sUserName)
        If InStr(sFilesOther(i), "%APPLICDATA%") > 0 Then sFilesOther(i) = Replace(sFilesOther(i), "%APPLICDATA%", sApplicData)
        If InStr(sFilesOther(i), "%TEMPDIR%") > 0 Then sFilesOther(i) = Replace(sFilesOther(i), "%TEMPDIR%", sTempDir)
    Next i
    For i = 1 To UBound(sProcessesOther)
        If sProcessesOther(i) = vbNullString Then Exit For
        If InStr(sProcessesOther(i), "%PROGRAMFILES%") > 0 Then sProcessesOther(i) = Replace(sProcessesOther(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sProcessesOther(i), "%KAZAAFOLDER%") > 0 Then sProcessesOther(i) = Replace(sProcessesOther(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sProcessesOther(i), "%STARTMENU%") > 0 Then sProcessesOther(i) = Replace(sProcessesOther(i), "%STARTMENU%", sStartMenu)
        If InStr(sProcessesOther(i), "%WINDIR%") > 0 Then sProcessesOther(i) = Replace(sProcessesOther(i), "%WINDIR%", sWinDir)
        If InStr(sProcessesOther(i), "%WINSYSDIR%") > 0 Then sProcessesOther(i) = Replace(sProcessesOther(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sProcessesOther(i), "%DESKTOP%") > 0 Then sProcessesOther(i) = Replace(sProcessesOther(i), "%DESKTOP%", sDesktop)
        If InStr(sProcessesOther(i), "%USERNAME%") > 0 Then sProcessesOther(i) = Replace(sProcessesOther(i), "%USERNAME%", sUserName)
        If InStr(sProcessesOther(i), "%APPLICDATA%") > 0 Then sProcessesOther(i) = Replace(sProcessesOther(i), "%APPLICDATA%", sApplicData)
        If InStr(sProcessesOther(i), "%TEMPDIR%") > 0 Then sProcessesOther(i) = Replace(sProcessesOther(i), "%TEMPDIR%", sTempDir)
    Next i
    
    ' --------- ALTNET ------------------------
    For i = 1 To UBound(sFoldersAltnet)
        If sFoldersAltnet(i) = vbNullString Then Exit For
        If InStr(sFoldersAltnet(i), "%PROGRAMFILES%") > 0 Then sFoldersAltnet(i) = Replace(sFoldersAltnet(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFoldersAltnet(i), "%KAZAAFOLDER%") > 0 Then sFoldersAltnet(i) = Replace(sFoldersAltnet(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFoldersAltnet(i), "%STARTMENU%") > 0 Then sFoldersAltnet(i) = Replace(sFoldersAltnet(i), "%STARTMENU%", sStartMenu)
        If InStr(sFoldersAltnet(i), "%WINDIR%") > 0 Then sFoldersAltnet(i) = Replace(sFoldersAltnet(i), "%WINDIR%", sWinDir)
        If InStr(sFoldersAltnet(i), "%WINSYSDIR%") > 0 Then sFoldersAltnet(i) = Replace(sFoldersAltnet(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFoldersAltnet(i), "%DESKTOP%") > 0 Then sFoldersAltnet(i) = Replace(sFoldersAltnet(i), "%DESKTOP%", sDesktop)
        If InStr(sFoldersAltnet(i), "%USERNAME%") > 0 Then sFoldersAltnet(i) = Replace(sFoldersAltnet(i), "%USERNAME%", sUserName)
        If InStr(sFoldersAltnet(i), "%APPLICDATA%") > 0 Then sFoldersAltnet(i) = Replace(sFoldersAltnet(i), "%APPLICDATA%", sApplicData)
        If InStr(sFoldersAltnet(i), "%PROGRAMS%") > 0 Then sFoldersAltnet(i) = Replace(sFoldersAltnet(i), "%PROGRAMS%", sPrograms)
        If InStr(sFoldersAltnet(i), "%TEMPDIR%") > 0 Then sFoldersAltnet(i) = Replace(sFoldersAltnet(i), "%TEMPDIR%", sTempDir)
    Next i
    For i = 1 To UBound(sFilesAltnet)
        If sFilesAltnet(i) = vbNullString Then Exit For
        If InStr(sFilesAltnet(i), "%PROGRAMFILES%") > 0 Then sFilesAltnet(i) = Replace(sFilesAltnet(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFilesAltnet(i), "%KAZAAFOLDER%") > 0 Then sFilesAltnet(i) = Replace(sFilesAltnet(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFilesAltnet(i), "%STARTMENU%") > 0 Then sFilesAltnet(i) = Replace(sFilesAltnet(i), "%STARTMENU%", sStartMenu)
        If InStr(sFilesAltnet(i), "%WINDIR%") > 0 Then sFilesAltnet(i) = Replace(sFilesAltnet(i), "%WINDIR%", sWinDir)
        If InStr(sFilesAltnet(i), "%WINSYSDIR%") > 0 Then sFilesAltnet(i) = Replace(sFilesAltnet(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFilesAltnet(i), "%DESKTOP%") > 0 Then sFilesAltnet(i) = Replace(sFilesAltnet(i), "%DESKTOP%", sDesktop)
        If InStr(sFilesAltnet(i), "%USERNAME%") > 0 Then sFilesAltnet(i) = Replace(sFilesAltnet(i), "%USERNAME%", sUserName)
        If InStr(sFilesAltnet(i), "%APPLICDATA%") > 0 Then sFilesAltnet(i) = Replace(sFilesAltnet(i), "%APPLICDATA%", sApplicData)
        If InStr(sFilesAltnet(i), "%PROGRAMS%") > 0 Then sFilesAltnet(i) = Replace(sFilesAltnet(i), "%PROGRAMS%", sPrograms)
        If InStr(sFilesAltnet(i), "%TEMPDIR%") > 0 Then sFilesAltnet(i) = Replace(sFilesAltnet(i), "%TEMPDIR%", sTempDir)
    Next i
    For i = 1 To UBound(sProcessesAltnet)
        If sProcessesAltnet(i) = vbNullString Then Exit For
        If InStr(sProcessesAltnet(i), "%PROGRAMFILES%") > 0 Then sProcessesAltnet(i) = Replace(sProcessesAltnet(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sProcessesAltnet(i), "%KAZAAFOLDER%") > 0 Then sProcessesAltnet(i) = Replace(sProcessesAltnet(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sProcessesAltnet(i), "%STARTMENU%") > 0 Then sProcessesAltnet(i) = Replace(sProcessesAltnet(i), "%STARTMENU%", sStartMenu)
        If InStr(sProcessesAltnet(i), "%WINDIR%") > 0 Then sProcessesAltnet(i) = Replace(sProcessesAltnet(i), "%WINDIR%", sWinDir)
        If InStr(sProcessesAltnet(i), "%WINSYSDIR%") > 0 Then sProcessesAltnet(i) = Replace(sProcessesAltnet(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sProcessesAltnet(i), "%DESKTOP%") > 0 Then sProcessesAltnet(i) = Replace(sProcessesAltnet(i), "%DESKTOP%", sDesktop)
        If InStr(sProcessesAltnet(i), "%USERNAME%") > 0 Then sProcessesAltnet(i) = Replace(sProcessesAltnet(i), "%USERNAME%", sUserName)
        If InStr(sProcessesAltnet(i), "%APPLICDATA%") > 0 Then sProcessesAltnet(i) = Replace(sProcessesAltnet(i), "%APPLICDATA%", sApplicData)
        If InStr(sProcessesAltnet(i), "%TEMPDIR%") > 0 Then sProcessesAltnet(i) = Replace(sProcessesAltnet(i), "%TEMPDIR%", sTempDir)
    Next i
    ' --------- BULLGUARD ------------------------
    For i = 1 To UBound(sFoldersBullguard)
        If sFoldersBullguard(i) = vbNullString Then Exit For
        If InStr(sFoldersBullguard(i), "%PROGRAMFILES%") > 0 Then sFoldersBullguard(i) = Replace(sFoldersBullguard(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFoldersBullguard(i), "%KAZAAFOLDER%") > 0 Then sFoldersBullguard(i) = Replace(sFoldersBullguard(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFoldersBullguard(i), "%STARTMENU%") > 0 Then sFoldersBullguard(i) = Replace(sFoldersBullguard(i), "%STARTMENU%", sStartMenu)
        If InStr(sFoldersBullguard(i), "%WINDIR%") > 0 Then sFoldersBullguard(i) = Replace(sFoldersBullguard(i), "%WINDIR%", sWinDir)
        If InStr(sFoldersBullguard(i), "%WINSYSDIR%") > 0 Then sFoldersBullguard(i) = Replace(sFoldersBullguard(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFoldersBullguard(i), "%DESKTOP%") > 0 Then sFoldersBullguard(i) = Replace(sFoldersBullguard(i), "%DESKTOP%", sDesktop)
        If InStr(sFoldersBullguard(i), "%USERNAME%") > 0 Then sFoldersBullguard(i) = Replace(sFoldersBullguard(i), "%USERNAME%", sUserName)
        If InStr(sFoldersBullguard(i), "%APPLICDATA%") > 0 Then sFoldersBullguard(i) = Replace(sFoldersBullguard(i), "%APPLICDATA%", sApplicData)
        If InStr(sFoldersBullguard(i), "%TEMPDIR%") > 0 Then sFoldersBullguard(i) = Replace(sFoldersBullguard(i), "%TEMPDIR%", sTempDir)
        If InStr(sFoldersBullguard(i), "%ALLUSERSDESKTOP%") > 0 Then sFoldersBullguard(i) = Replace(sFoldersBullguard(i), "%ALLUSERSDESKTOP%", sAllUsersDesktop)
    Next i
    For i = 1 To UBound(sFilesBullguard)
        If sFilesBullguard(i) = vbNullString Then Exit For
        If InStr(sFilesBullguard(i), "%PROGRAMFILES%") > 0 Then sFilesBullguard(i) = Replace(sFilesBullguard(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFilesBullguard(i), "%KAZAAFOLDER%") > 0 Then sFilesBullguard(i) = Replace(sFilesBullguard(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFilesBullguard(i), "%STARTMENU%") > 0 Then sFilesBullguard(i) = Replace(sFilesBullguard(i), "%STARTMENU%", sStartMenu)
        If InStr(sFilesBullguard(i), "%WINDIR%") > 0 Then sFilesBullguard(i) = Replace(sFilesBullguard(i), "%WINDIR%", sWinDir)
        If InStr(sFilesBullguard(i), "%WINSYSDIR%") > 0 Then sFilesBullguard(i) = Replace(sFilesBullguard(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFilesBullguard(i), "%DESKTOP%") > 0 Then sFilesBullguard(i) = Replace(sFilesBullguard(i), "%DESKTOP%", sDesktop)
        If InStr(sFilesBullguard(i), "%USERNAME%") > 0 Then sFilesBullguard(i) = Replace(sFilesBullguard(i), "%USERNAME%", sUserName)
        If InStr(sFilesBullguard(i), "%APPLICDATA%") > 0 Then sFilesBullguard(i) = Replace(sFilesBullguard(i), "%APPLICDATA%", sApplicData)
        If InStr(sFilesBullguard(i), "%TEMPDIR%") > 0 Then sFilesBullguard(i) = Replace(sFilesBullguard(i), "%TEMPDIR%", sTempDir)
        If InStr(sFilesBullguard(i), "%ALLUSERSDESKTOP%") > 0 Then sFilesBullguard(i) = Replace(sFilesBullguard(i), "%ALLUSERSDESKTOP%", sAllUsersDesktop)
    Next i
    For i = 1 To UBound(sProcessesBullguard)
        If sProcessesBullguard(i) = vbNullString Then Exit For
        If InStr(sProcessesBullguard(i), "%PROGRAMFILES%") > 0 Then sProcessesBullguard(i) = Replace(sProcessesBullguard(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sProcessesBullguard(i), "%KAZAAFOLDER%") > 0 Then sProcessesBullguard(i) = Replace(sProcessesBullguard(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sProcessesBullguard(i), "%STARTMENU%") > 0 Then sProcessesBullguard(i) = Replace(sProcessesBullguard(i), "%STARTMENU%", sStartMenu)
        If InStr(sProcessesBullguard(i), "%WINDIR%") > 0 Then sProcessesBullguard(i) = Replace(sProcessesBullguard(i), "%WINDIR%", sWinDir)
        If InStr(sProcessesBullguard(i), "%WINSYSDIR%") > 0 Then sProcessesBullguard(i) = Replace(sProcessesBullguard(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sProcessesBullguard(i), "%DESKTOP%") > 0 Then sProcessesBullguard(i) = Replace(sProcessesBullguard(i), "%DESKTOP%", sDesktop)
        If InStr(sProcessesBullguard(i), "%USERNAME%") > 0 Then sProcessesBullguard(i) = Replace(sProcessesBullguard(i), "%USERNAME%", sUserName)
        If InStr(sProcessesBullguard(i), "%APPLICDATA%") > 0 Then sProcessesBullguard(i) = Replace(sProcessesBullguard(i), "%APPLICDATA%", sApplicData)
        If InStr(sProcessesBullguard(i), "%TEMPDIR%") > 0 Then sProcessesBullguard(i) = Replace(sProcessesBullguard(i), "%TEMPDIR%", sTempDir)
    Next i
    ' --------- GATOR ------------------------
    For i = 1 To UBound(sFoldersGator)
        If sFoldersGator(i) = vbNullString Then Exit For
        If InStr(sFoldersGator(i), "%PROGRAMFILES%") > 0 Then sFoldersGator(i) = Replace(sFoldersGator(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFoldersGator(i), "%KAZAAFOLDER%") > 0 Then sFoldersGator(i) = Replace(sFoldersGator(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFoldersGator(i), "%STARTMENU%") > 0 Then sFoldersGator(i) = Replace(sFoldersGator(i), "%STARTMENU%", sStartMenu)
        If InStr(sFoldersGator(i), "%WINDIR%") > 0 Then sFoldersGator(i) = Replace(sFoldersGator(i), "%WINDIR%", sWinDir)
        If InStr(sFoldersGator(i), "%WINSYSDIR%") > 0 Then sFoldersGator(i) = Replace(sFoldersGator(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFoldersGator(i), "%DESKTOP%") > 0 Then sFoldersGator(i) = Replace(sFoldersGator(i), "%DESKTOP%", sDesktop)
        If InStr(sFoldersGator(i), "%USERNAME%") > 0 Then sFoldersGator(i) = Replace(sFoldersGator(i), "%USERNAME%", sUserName)
        If InStr(sFoldersGator(i), "%APPLICDATA%") > 0 Then sFoldersGator(i) = Replace(sFoldersGator(i), "%APPLICDATA%", sApplicData)
        If InStr(sFoldersGator(i), "%PROGRAMS%") > 0 Then sFoldersGator(i) = Replace(sFoldersGator(i), "%PROGRAMS%", sPrograms)
        If InStr(sFoldersGator(i), "%TEMPDIR%") > 0 Then sFoldersGator(i) = Replace(sFoldersGator(i), "%TEMPDIR%", sTempDir)
    Next i
    For i = 1 To UBound(sFilesGator)
        If sFilesGator(i) = vbNullString Then Exit For
        If InStr(sFilesGator(i), "%PROGRAMFILES%") > 0 Then sFilesGator(i) = Replace(sFilesGator(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFilesGator(i), "%KAZAAFOLDER%") > 0 Then sFilesGator(i) = Replace(sFilesGator(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFilesGator(i), "%STARTMENU%") > 0 Then sFilesGator(i) = Replace(sFilesGator(i), "%STARTMENU%", sStartMenu)
        If InStr(sFilesGator(i), "%WINDIR%") > 0 Then sFilesGator(i) = Replace(sFilesGator(i), "%WINDIR%", sWinDir)
        If InStr(sFilesGator(i), "%WINSYSDIR%") > 0 Then sFilesGator(i) = Replace(sFilesGator(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFilesGator(i), "%DESKTOP%") > 0 Then sFilesGator(i) = Replace(sFilesGator(i), "%DESKTOP%", sDesktop)
        If InStr(sFilesGator(i), "%USERNAME%") > 0 Then sFilesGator(i) = Replace(sFilesGator(i), "%USERNAME%", sUserName)
        If InStr(sFilesGator(i), "%APPLICDATA%") > 0 Then sFilesGator(i) = Replace(sFilesGator(i), "%APPLICDATA%", sApplicData)
        If InStr(sFilesGator(i), "%PROGRAMS%") > 0 Then sFilesGator(i) = Replace(sFilesGator(i), "%PROGRAMS%", sPrograms)
        If InStr(sFilesGator(i), "%TEMPDIR%") > 0 Then sFilesGator(i) = Replace(sFilesGator(i), "%TEMPDIR%", sTempDir)
    Next i
    For i = 1 To UBound(sProcessesGator)
        If sProcessesGator(i) = vbNullString Then Exit For
        If InStr(sProcessesGator(i), "%PROGRAMFILES%") > 0 Then sProcessesGator(i) = Replace(sProcessesGator(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sProcessesGator(i), "%KAZAAFOLDER%") > 0 Then sProcessesGator(i) = Replace(sProcessesGator(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sProcessesGator(i), "%STARTMENU%") > 0 Then sProcessesGator(i) = Replace(sProcessesGator(i), "%STARTMENU%", sStartMenu)
        If InStr(sProcessesGator(i), "%WINDIR%") > 0 Then sProcessesGator(i) = Replace(sProcessesGator(i), "%WINDIR%", sWinDir)
        If InStr(sProcessesGator(i), "%WINSYSDIR%") > 0 Then sProcessesGator(i) = Replace(sProcessesGator(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sProcessesGator(i), "%DESKTOP%") > 0 Then sProcessesGator(i) = Replace(sProcessesGator(i), "%DESKTOP%", sDesktop)
        If InStr(sProcessesGator(i), "%USERNAME%") > 0 Then sProcessesGator(i) = Replace(sProcessesGator(i), "%USERNAME%", sUserName)
        If InStr(sProcessesGator(i), "%APPLICDATA%") > 0 Then sProcessesGator(i) = Replace(sProcessesGator(i), "%APPLICDATA%", sApplicData)
        If InStr(sProcessesGator(i), "%TEMPDIR%") > 0 Then sProcessesGator(i) = Replace(sProcessesGator(i), "%TEMPDIR%", sTempDir)
    Next i
    ' --------- MYWAY ------------------------
    For i = 1 To UBound(sFoldersMyWay)
        If sFoldersMyWay(i) = vbNullString Then Exit For
        If InStr(sFoldersMyWay(i), "%PROGRAMFILES%") > 0 Then sFoldersMyWay(i) = Replace(sFoldersMyWay(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFoldersMyWay(i), "%KAZAAFOLDER%") > 0 Then sFoldersMyWay(i) = Replace(sFoldersMyWay(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFoldersMyWay(i), "%STARTMENU%") > 0 Then sFoldersMyWay(i) = Replace(sFoldersMyWay(i), "%STARTMENU%", sStartMenu)
        If InStr(sFoldersMyWay(i), "%WINDIR%") > 0 Then sFoldersMyWay(i) = Replace(sFoldersMyWay(i), "%WINDIR%", sWinDir)
        If InStr(sFoldersMyWay(i), "%WINSYSDIR%") > 0 Then sFoldersMyWay(i) = Replace(sFoldersMyWay(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFoldersMyWay(i), "%DESKTOP%") > 0 Then sFoldersMyWay(i) = Replace(sFoldersMyWay(i), "%DESKTOP%", sDesktop)
        If InStr(sFoldersMyWay(i), "%USERNAME%") > 0 Then sFoldersMyWay(i) = Replace(sFoldersMyWay(i), "%USERNAME%", sUserName)
        If InStr(sFoldersMyWay(i), "%APPLICDATA%") > 0 Then sFoldersMyWay(i) = Replace(sFoldersMyWay(i), "%APPLICDATA%", sApplicData)
        If InStr(sFoldersMyWay(i), "%TEMPDIR%") > 0 Then sFoldersMyWay(i) = Replace(sFoldersMyWay(i), "%TEMPDIR%", sTempDir)
    Next i
    For i = 1 To UBound(sFilesMyWay)
        If sFilesMyWay(i) = vbNullString Then Exit For
        If InStr(sFilesMyWay(i), "%PROGRAMFILES%") > 0 Then sFilesMyWay(i) = Replace(sFilesMyWay(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFilesMyWay(i), "%KAZAAFOLDER%") > 0 Then sFilesMyWay(i) = Replace(sFilesMyWay(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFilesMyWay(i), "%STARTMENU%") > 0 Then sFilesMyWay(i) = Replace(sFilesMyWay(i), "%STARTMENU%", sStartMenu)
        If InStr(sFilesMyWay(i), "%WINDIR%") > 0 Then sFilesMyWay(i) = Replace(sFilesMyWay(i), "%WINDIR%", sWinDir)
        If InStr(sFilesMyWay(i), "%WINSYSDIR%") > 0 Then sFilesMyWay(i) = Replace(sFilesMyWay(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFilesMyWay(i), "%DESKTOP%") > 0 Then sFilesMyWay(i) = Replace(sFilesMyWay(i), "%DESKTOP%", sDesktop)
        If InStr(sFilesMyWay(i), "%USERNAME%") > 0 Then sFilesMyWay(i) = Replace(sFilesMyWay(i), "%USERNAME%", sUserName)
        If InStr(sFilesMyWay(i), "%APPLICDATA%") > 0 Then sFilesMyWay(i) = Replace(sFilesMyWay(i), "%APPLICDATA%", sApplicData)
        If InStr(sFilesMyWay(i), "%TEMPDIR%") > 0 Then sFilesMyWay(i) = Replace(sFilesMyWay(i), "%TEMPDIR%", sTempDir)
    Next i
    For i = 1 To UBound(sProcessesMyWay)
        If sProcessesMyWay(i) = vbNullString Then Exit For
        If InStr(sProcessesMyWay(i), "%PROGRAMFILES%") > 0 Then sProcessesMyWay(i) = Replace(sProcessesMyWay(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sProcessesMyWay(i), "%KAZAAFOLDER%") > 0 Then sProcessesMyWay(i) = Replace(sProcessesMyWay(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sProcessesMyWay(i), "%STARTMENU%") > 0 Then sProcessesMyWay(i) = Replace(sProcessesMyWay(i), "%STARTMENU%", sStartMenu)
        If InStr(sProcessesMyWay(i), "%WINDIR%") > 0 Then sProcessesMyWay(i) = Replace(sProcessesMyWay(i), "%WINDIR%", sWinDir)
        If InStr(sProcessesMyWay(i), "%WINSYSDIR%") > 0 Then sProcessesMyWay(i) = Replace(sProcessesMyWay(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sProcessesMyWay(i), "%DESKTOP%") > 0 Then sProcessesMyWay(i) = Replace(sProcessesMyWay(i), "%DESKTOP%", sDesktop)
        If InStr(sProcessesMyWay(i), "%USERNAME%") > 0 Then sProcessesMyWay(i) = Replace(sProcessesMyWay(i), "%USERNAME%", sUserName)
        If InStr(sProcessesMyWay(i), "%APPLICDATA%") > 0 Then sProcessesMyWay(i) = Replace(sProcessesMyWay(i), "%APPLICDATA%", sApplicData)
        If InStr(sProcessesMyWay(i), "%TEMPDIR%") > 0 Then sProcessesMyWay(i) = Replace(sProcessesMyWay(i), "%TEMPDIR%", sTempDir)
    Next i
    ' --------- P2PNETWORKING ------------------------
    For i = 1 To UBound(sFoldersP2P)
        If sFoldersP2P(i) = vbNullString Then Exit For
        If InStr(sFoldersP2P(i), "%PROGRAMFILES%") > 0 Then sFoldersP2P(i) = Replace(sFoldersP2P(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFoldersP2P(i), "%KAZAAFOLDER%") > 0 Then sFoldersP2P(i) = Replace(sFoldersP2P(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFoldersP2P(i), "%STARTMENU%") > 0 Then sFoldersP2P(i) = Replace(sFoldersP2P(i), "%STARTMENU%", sStartMenu)
        If InStr(sFoldersP2P(i), "%WINDIR%") > 0 Then sFoldersP2P(i) = Replace(sFoldersP2P(i), "%WINDIR%", sWinDir)
        If InStr(sFoldersP2P(i), "%WINSYSDIR%") > 0 Then sFoldersP2P(i) = Replace(sFoldersP2P(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFoldersP2P(i), "%DESKTOP%") > 0 Then sFoldersP2P(i) = Replace(sFoldersP2P(i), "%DESKTOP%", sDesktop)
        If InStr(sFoldersP2P(i), "%USERNAME%") > 0 Then sFoldersP2P(i) = Replace(sFoldersP2P(i), "%USERNAME%", sUserName)
        If InStr(sFoldersP2P(i), "%APPLICDATA%") > 0 Then sFoldersP2P(i) = Replace(sFoldersP2P(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sFilesP2P)
        If sFilesP2P(i) = vbNullString Then Exit For
        If InStr(sFilesP2P(i), "%PROGRAMFILES%") > 0 Then sFilesP2P(i) = Replace(sFilesP2P(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFilesP2P(i), "%KAZAAFOLDER%") > 0 Then sFilesP2P(i) = Replace(sFilesP2P(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFilesP2P(i), "%STARTMENU%") > 0 Then sFilesP2P(i) = Replace(sFilesP2P(i), "%STARTMENU%", sStartMenu)
        If InStr(sFilesP2P(i), "%WINDIR%") > 0 Then sFilesP2P(i) = Replace(sFilesP2P(i), "%WINDIR%", sWinDir)
        If InStr(sFilesP2P(i), "%WINSYSDIR%") > 0 Then sFilesP2P(i) = Replace(sFilesP2P(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFilesP2P(i), "%DESKTOP%") > 0 Then sFilesP2P(i) = Replace(sFilesP2P(i), "%DESKTOP%", sDesktop)
        If InStr(sFilesP2P(i), "%USERNAME%") > 0 Then sFilesP2P(i) = Replace(sFilesP2P(i), "%USERNAME%", sUserName)
        If InStr(sFilesP2P(i), "%APPLICDATA%") > 0 Then sFilesP2P(i) = Replace(sFilesP2P(i), "%APPLICDATA%", sApplicData)
    Next i
    For i = 1 To UBound(sProcessesP2P)
        If sProcessesP2P(i) = vbNullString Then Exit For
        If InStr(sProcessesP2P(i), "%PROGRAMFILES%") > 0 Then sProcessesP2P(i) = Replace(sProcessesP2P(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sProcessesP2P(i), "%KAZAAFOLDER%") > 0 Then sProcessesP2P(i) = Replace(sProcessesP2P(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sProcessesP2P(i), "%STARTMENU%") > 0 Then sProcessesP2P(i) = Replace(sProcessesP2P(i), "%STARTMENU%", sStartMenu)
        If InStr(sProcessesP2P(i), "%WINDIR%") > 0 Then sProcessesP2P(i) = Replace(sProcessesP2P(i), "%WINDIR%", sWinDir)
        If InStr(sProcessesP2P(i), "%WINSYSDIR%") > 0 Then sProcessesP2P(i) = Replace(sProcessesP2P(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sProcessesP2P(i), "%DESKTOP%") > 0 Then sProcessesP2P(i) = Replace(sProcessesP2P(i), "%DESKTOP%", sDesktop)
        If InStr(sProcessesP2P(i), "%USERNAME%") > 0 Then sProcessesP2P(i) = Replace(sProcessesP2P(i), "%USERNAME%", sUserName)
        If InStr(sProcessesP2P(i), "%APPLICDATA%") > 0 Then sProcessesP2P(i) = Replace(sProcessesP2P(i), "%APPLICDATA%", sApplicData)
    Next i
    ' --------- PERFECTNAV ------------------------
    For i = 1 To UBound(sFoldersPerfectNav)
        If sFoldersPerfectNav(i) = vbNullString Then Exit For
        If InStr(sFoldersPerfectNav(i), "%PROGRAMFILES%") > 0 Then sFoldersPerfectNav(i) = Replace(sFoldersPerfectNav(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFoldersPerfectNav(i), "%KAZAAFOLDER%") > 0 Then sFoldersPerfectNav(i) = Replace(sFoldersPerfectNav(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFoldersPerfectNav(i), "%STARTMENU%") > 0 Then sFoldersPerfectNav(i) = Replace(sFoldersPerfectNav(i), "%STARTMENU%", sStartMenu)
        If InStr(sFoldersPerfectNav(i), "%WINDIR%") > 0 Then sFoldersPerfectNav(i) = Replace(sFoldersPerfectNav(i), "%WINDIR%", sWinDir)
        If InStr(sFoldersPerfectNav(i), "%WINSYSDIR%") > 0 Then sFoldersPerfectNav(i) = Replace(sFoldersPerfectNav(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFoldersPerfectNav(i), "%DESKTOP%") > 0 Then sFoldersPerfectNav(i) = Replace(sFoldersPerfectNav(i), "%DESKTOP%", sDesktop)
        If InStr(sFoldersPerfectNav(i), "%USERNAME%") > 0 Then sFoldersPerfectNav(i) = Replace(sFoldersPerfectNav(i), "%USERNAME%", sUserName)
        If InStr(sFoldersPerfectNav(i), "%APPLICDATA%") > 0 Then sFoldersPerfectNav(i) = Replace(sFoldersPerfectNav(i), "%APPLICDATA%", sApplicData)
        If InStr(sFoldersPerfectNav(i), "%TEMPDIR%") > 0 Then sFoldersPerfectNav(i) = Replace(sFoldersPerfectNav(i), "%TEMPDIR%", sTempDir)
    Next i
    For i = 1 To UBound(sFilesPerfectNav)
        If sFilesPerfectNav(i) = vbNullString Then Exit For
        If InStr(sFilesPerfectNav(i), "%PROGRAMFILES%") > 0 Then sFilesPerfectNav(i) = Replace(sFilesPerfectNav(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sFilesPerfectNav(i), "%KAZAAFOLDER%") > 0 Then sFilesPerfectNav(i) = Replace(sFilesPerfectNav(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sFilesPerfectNav(i), "%STARTMENU%") > 0 Then sFilesPerfectNav(i) = Replace(sFilesPerfectNav(i), "%STARTMENU%", sStartMenu)
        If InStr(sFilesPerfectNav(i), "%WINDIR%") > 0 Then sFilesPerfectNav(i) = Replace(sFilesPerfectNav(i), "%WINDIR%", sWinDir)
        If InStr(sFilesPerfectNav(i), "%WINSYSDIR%") > 0 Then sFilesPerfectNav(i) = Replace(sFilesPerfectNav(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sFilesPerfectNav(i), "%DESKTOP%") > 0 Then sFilesPerfectNav(i) = Replace(sFilesPerfectNav(i), "%DESKTOP%", sDesktop)
        If InStr(sFilesPerfectNav(i), "%USERNAME%") > 0 Then sFilesPerfectNav(i) = Replace(sFilesPerfectNav(i), "%USERNAME%", sUserName)
        If InStr(sFilesPerfectNav(i), "%APPLICDATA%") > 0 Then sFilesPerfectNav(i) = Replace(sFilesPerfectNav(i), "%APPLICDATA%", sApplicData)
        If InStr(sFilesPerfectNav(i), "%TEMPDIR%") > 0 Then sFilesPerfectNav(i) = Replace(sFilesPerfectNav(i), "%TEMPDIR%", sTempDir)
    Next i
    For i = 1 To UBound(sProcessesPerfectNav)
        If sProcessesPerfectNav(i) = vbNullString Then Exit For
        If InStr(sProcessesPerfectNav(i), "%PROGRAMFILES%") > 0 Then sProcessesPerfectNav(i) = Replace(sProcessesPerfectNav(i), "%PROGRAMFILES%", sProgramFiles)
        If InStr(sProcessesPerfectNav(i), "%KAZAAFOLDER%") > 0 Then sProcessesPerfectNav(i) = Replace(sProcessesPerfectNav(i), "%KAZAAFOLDER%", sKazaaFolder)
        If InStr(sProcessesPerfectNav(i), "%STARTMENU%") > 0 Then sProcessesPerfectNav(i) = Replace(sProcessesPerfectNav(i), "%STARTMENU%", sStartMenu)
        If InStr(sProcessesPerfectNav(i), "%WINDIR%") > 0 Then sProcessesPerfectNav(i) = Replace(sProcessesPerfectNav(i), "%WINDIR%", sWinDir)
        If InStr(sProcessesPerfectNav(i), "%WINSYSDIR%") > 0 Then sProcessesPerfectNav(i) = Replace(sProcessesPerfectNav(i), "%WINSYSDIR%", sWinSysDir)
        If InStr(sProcessesPerfectNav(i), "%DESKTOP%") > 0 Then sProcessesPerfectNav(i) = Replace(sProcessesPerfectNav(i), "%DESKTOP%", sDesktop)
        If InStr(sProcessesPerfectNav(i), "%USERNAME%") > 0 Then sProcessesPerfectNav(i) = Replace(sProcessesPerfectNav(i), "%USERNAME%", sUserName)
        If InStr(sProcessesPerfectNav(i), "%APPLICDATA%") > 0 Then sProcessesPerfectNav(i) = Replace(sProcessesPerfectNav(i), "%APPLICDATA%", sApplicData)
        If InStr(sProcessesPerfectNav(i), "%TEMPDIR%") > 0 Then sProcessesPerfectNav(i) = Replace(sProcessesPerfectNav(i), "%TEMPDIR%", sTempDir)
    Next i
    
    sRegvalsKazaa(2) = Replace(sRegvalsKazaa(2), "%PROGRAMFILES%", sProgramFiles)
    sRegvalsKazaa(3) = Replace(sRegvalsKazaa(3), "%PROGRAMFILES%", sProgramFiles)
    sRegKeysOther(1) = Replace(sRegKeysOther(1), "%USERNAME%", sUserName)
    sRegKeysOther(2) = Replace(sRegKeysOther(2), "%USERNAME%", sUserName)
    sRegKeysOther(3) = Replace(sRegKeysOther(3), "%USERNAME%", sUserName)
    sRegvalsOther(2) = Replace(sRegvalsOther(2), "%PROGRAMFILES%", sProgramFiles)
    sFoldersOther(2) = Replace(sFoldersOther(2), "%APPDATA%", sApplicData)
    sFoldersOther(3) = Replace(sFoldersOther(3), "%APPDATA%", sApplicData)
    sFilesOther(39) = Replace(sFilesOther(39), "%APPDATA%", sApplicData)
    sRegvalsCommonName(2) = Replace(sRegvalsCommonName(2), "%PROGRAMFILES%", sProgramFiles)
    sRegvalsCommonName(3) = Replace(sRegvalsCommonName(3), "%PROGRAMFILES%", sProgramFiles)
    sRegvalsCommonName(4) = Replace(sRegvalsCommonName(4), "%PROGRAMFILES%", sProgramFiles)
    sFoldersBullguard(2) = Replace(sFoldersBullguard(2), "%SYSDRIVE%", sSysDrive)
    sFoldersBullguard(3) = Replace(sFoldersBullguard(3), "%ALLUSERSAPPDATA%", sAllUsersAppData)
    sFoldersBullguard(4) = Replace(sFoldersBullguard(4), "%ALLUSERSPROGRAMS%", sAllUsersPrograms)
    sRegvalsP2P(2) = Replace(sRegvalsP2P(2), "%WINSYSDIR%", sWinSysDir)
    sRegvalsP2P(3) = Replace(sRegvalsP2P(3), "%WINSYSDIR%", sWinSysDir)
End Sub

Public Function RegGetString$(lHive&, sKey$, sValue$, Optional sDef$ = vbNullString)
    Dim hKey&, sData$
    If RunningInIDE And (InStr(sKey, "%") > 0 Or InStr(sValue, "%") > 0) Then
        MsgBox "Found a % in sKey or sValue, that sucks!" & _
                vbCrLf & sKey & "\" & sValue, , "RegGetString"
        Exit Function
    End If
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        sData = String(255, 0)
        RegQueryValueEx hKey, sValue, 0, REG_SZ, ByVal sData, 255
        sData = Left(sData, InStr(sData, Chr(0)) - 1)
        RegCloseKey hKey
    End If
    RegGetString = IIf(sData = vbNullString, sDef, sData)
End Function

Public Function RegGetDword&(lHive&, sKey$, sValue$, Optional lDef& = 0)
    Dim hKey&, lData&
    If RunningInIDE And (InStr(sKey, "%") > 0 Or InStr(sValue, "%") > 0) Then
        MsgBox "Found a % in sKey or sValue, that sucks!" & _
                vbCrLf & sKey & "\" & sValue, , "RegGetDword"
        Exit Function
    End If
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        If RegQueryValueEx(hKey, sValue, 0, REG_DWORD, lData, 4) = 0 Then
            RegGetDword = lData
        Else
            RegGetDword = lDef
        End If
        RegCloseKey hKey
    End If
End Function

Public Sub Logg(s$)
    With frmMain.lstLog
        .AddItem s
        '.ListIndex = .ListCount - 1
    End With
End Sub

Public Sub Status(s$)
    frmMain.lblStatus.Caption = s
    DoEvents
End Sub

Public Function RegKeyExists(sRegKey$) As Boolean
    'HKLM\Software\Kazaa
    Dim hKey&, lHive&, sKey$
    If RunningInIDE And InStr(sRegKey, "%") > 0 Then
        MsgBox "Found a % in sRegKey, that sucks!" & _
                vbCrLf & sRegKey, , "RegKeyExists"
        Exit Function
    End If
    Select Case Left(sRegKey, 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case "HKUS": lHive = HKEY_USERS
        Case Else: MsgBox "dork! you misspelled '" & sRegKey & "'!", vbExclamation: Exit Function
    End Select
    sKey = Mid(sRegKey, 6)
    
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        RegKeyExists = True
    Else
        RegKeyExists = False
    End If
    RegCloseKey hKey
End Function

Public Function RegvalExists(sRegval$) As Boolean
    Dim lHive&, sKey$, sVal$, hKey&
    If RunningInIDE And InStr(sRegval, "%") > 0 And InStr(sRegval, "%SystemRoot%") = 0 Then
        MsgBox "Found a % in sRegval, that sucks!" & _
                vbCrLf & sRegval, , "RegvalExists"
        Exit Function
    End If
    Select Case Left(sRegval, 4)
        Case "HKCR": lHive = HKEY_CLASSES_ROOT
        Case "HKCU": lHive = HKEY_CURRENT_USER
        Case "HKLM": lHive = HKEY_LOCAL_MACHINE
        Case Else: MsgBox "dork! you misspelled '" & sRegval & "'!", vbExclamation: Exit Function
    End Select
    sKey = Mid(sRegval, 6)
    sVal = Mid(sKey, InStr(sKey, ",") + 1)
    sKey = Left(sKey, InStr(sKey, ",") - 1)
    
    If Not RegKeyExists(Left(sRegval, InStr(sRegval, ",") - 1)) Then
        RegvalExists = False
        Exit Function
    End If
    
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        If RegQueryValueEx(hKey, sVal, 0, 0, ByVal 0, 0) = 0 Then
            RegvalExists = True
        Else
            RegvalExists = False
        End If
        RegCloseKey hKey
    Else
        RegvalExists = False
    End If
End Function

Public Function FileExists(sFile$) As Boolean
    If RunningInIDE And InStr(sFile, "%") > 0 Then
        MsgBox "found a % in sFile, that sucks!" & _
               vbCrLf & sFile, , "FileExists"
        Exit Function
    End If
    
    FileExists = False
    If bIsWinNT Then
        If SHFileExists(StrConv(sFile, vbUnicode)) Then FileExists = True
    Else
        If SHFileExists(sFile) Then FileExists = True
    End If
End Function

Public Function FolderExists(sFolder$) As Boolean
    If RunningInIDE And InStr(sFolder, "%") > 0 Then
        MsgBox "found a % in sFolder, that sucks!" & _
               vbCrLf & sFolder, , "FolderExists"
        Exit Function
    End If
    
    FolderExists = False
    If bIsWinNT Then
        If SHFileExists(StrConv(sFolder, vbUnicode)) Then FolderExists = True
    Else
        If SHFileExists(sFolder) Then FolderExists = True
    End If
End Function

Public Function RunningInIDE() As Boolean
    On Error Resume Next
    Debug.Print 1 / 0
    If Err Then
        RunningInIDE = True
    Else
        RunningInIDE = False
    End If
End Function

Public Function CheckKazaaRunning() As Boolean
    Dim hwndKazaa&, sMsg$
    CheckKazaaRunning = False
    
    hwndKazaa = FindWindow("Kazaa", vbNullString)
    If hwndKazaa = 0 Then Exit Function
    CheckKazaaRunning = True
    
    If IsWindowVisible(hwndKazaa) = 1 Then
        'kazaa window is visible
        'MsgBox "visible"
    Else
        'kazaa is minimized to system tray
        'MsgBox "not visible"
    End If
    
    sMsg = "Kazaa is still running. Please " & _
           "close it by right-clicking the " & _
           "Kazaa icon in the bottom-right " & _
           "corner of the screen and " & _
           "selecting 'Close Kazaa'."
    
    MsgBox sMsg, vbExclamation
    
    hwndKazaa = 0
    hwndKazaa = FindWindow("Kazaa", vbNullString)
    If hwndKazaa = 0 Then
        CheckKazaaRunning = False
    Else
        CheckKazaaRunning = True
    End If
End Function

Public Sub EnumKazaaComponents()
    Dim i%
    Status "Searching for Kazaa processes..."
    For i = 1 To UBound(sProcessesKazaa)
        If sProcessesKazaa(i) = vbNullString Then Exit For
        If ProcessExists(sProcessesKazaa(i)) Then Logg "PROCESS: [KAZAA] " & sProcessesKazaa(i)
    Next i
    Status "Searching for Kazaa regkeys..."
    For i = 1 To UBound(sRegKeysKazaa)
        If sRegKeysKazaa(i) = vbNullString Then Exit For
        If RegKeyExists(sRegKeysKazaa(i)) Then Logg "REGKEY: [Kazaa] " & sRegKeysKazaa(i)
    Next i
    Status "Searching for Kazaa folders..."
    For i = 1 To UBound(sFoldersKazaa)
        If sFoldersKazaa(i) = vbNullString Then Exit For
        If FolderExists(sFoldersKazaa(i)) Then Logg "FOLDER: [Kazaa] " & sFoldersKazaa(i)
    Next i
    Status "Searching for Kazaa files..."
    For i = 1 To UBound(sFilesKazaa)
        If sFilesKazaa(i) = vbNullString Then Exit For
        If FileExists(sFilesKazaa(i)) Then Logg "FILE: [Kazaa] " & sFilesKazaa(i)
    Next i
    For i = 1 To UBound(sRegvalsKazaa)
        If sRegvalsKazaa(i) = vbNullString Then Exit For
        If RegvalExists(sRegvalsKazaa(i)) Then Logg "REGVAL: [Kazaa] " & sRegvalsKazaa(i)
    Next i
End Sub

Public Sub EnumBDEComponents()
    Dim i%
    Status "Searching for BDE/B3D processes..."
    For i = 1 To UBound(sProcessesBDE)
        If sProcessesBDE(i) = vbNullString Then Exit For
        If ProcessExists(sProcessesBDE(i)) Then Logg "PROCESS: [BDE] " & sProcessesBDE(i)
    Next i
    Status "Searching for BDE/B3D regkeys..."
    For i = 1 To UBound(sRegKeysBDE)
        If sRegKeysBDE(i) = vbNullString Then Exit For
        If RegKeyExists(sRegKeysBDE(i)) Then Logg "REGKEY: [BDE] " & sRegKeysBDE(i)
    Next i
    Status "Searching for BDE/B3D folders..."
    For i = 1 To UBound(sFoldersBDE)
        If sFoldersBDE(i) = vbNullString Then Exit For
        If FolderExists(sFoldersBDE(i)) Then Logg "FOLDER: [BDE] " & sFoldersBDE(i)
    Next i
    Status "Searching for BDE/B3D files..."
    For i = 1 To UBound(sFilesBDE)
        If sFilesBDE(i) = vbNullString Then Exit For
        If FileExists(sFilesBDE(i)) Then Logg "FILE: [BDE] " & sFilesBDE(i)
    Next i
    Status "Searching for BDE/B3D regvals..."
    For i = 1 To UBound(sRegvalsBDE)
        If sRegvalsBDE(i) = vbNullString Then Exit For
        If RegvalExists(sRegvalsBDE(i)) Then Logg "REGVAL: [BDE] " & sRegvalsBDE(i)
    Next i
End Sub

Public Sub EnumCyDoorComponents()
    Dim i%
    Status "Searching for CyDoor processes..."
    For i = 1 To UBound(sProcessesCyDoor)
        If sProcessesCyDoor(i) = vbNullString Then Exit For
        If ProcessExists(sProcessesCyDoor(i)) Then Logg "PROCESS: [CyDoor] " & sProcessesCyDoor(i)
    Next i
    Status "Searching for CyDoor regkeys..."
    For i = 1 To UBound(sRegKeysCyDoor)
        If sRegKeysCyDoor(i) = vbNullString Then Exit For
        If RegKeyExists(sRegKeysCyDoor(i)) Then Logg "REGKEY: [CyDoor] " & sRegKeysCyDoor(i)
    Next i
    Status "Searching for CyDoor folders..."
    For i = 1 To UBound(sFoldersCyDoor)
        If sFoldersCyDoor(i) = vbNullString Then Exit For
        If FolderExists(sFoldersCyDoor(i)) Then Logg "FOLDER: [CyDoor] " & sFoldersCyDoor(i)
    Next i
    Status "Searching for CyDoor files..."
    For i = 1 To UBound(sFilesCyDoor)
        If sFilesCyDoor(i) = vbNullString Then Exit For
        If FileExists(sFilesCyDoor(i)) Then Logg "FILE: [CyDoor] " & sFilesCyDoor(i)
    Next i
    For i = 1 To UBound(sRegvalsCyDoor)
        If sRegvalsCyDoor(i) = vbNullString Then Exit For
        If RegvalExists(sRegvalsCyDoor(i)) Then Logg "REGVAL: [CyDoor] " & sRegvalsCyDoor(i)
    Next i
End Sub

Public Sub EnumCommonNameComponents()
    Dim i%
    Status "Searching for CommonName processes..."
    For i = 1 To UBound(sProcessesCommonName)
        If sProcessesCommonName(i) = vbNullString Then Exit For
        If ProcessExists(sProcessesCommonName(i)) Then Logg "PROCESS: [CommonName] " & sProcessesCommonName(i)
    Next i
    Status "Searching for CommonName regkeys..."
    For i = 1 To UBound(sRegKeysCommonName)
        If sRegKeysCommonName(i) = vbNullString Then Exit For
        If RegKeyExists(sRegKeysCommonName(i)) Then Logg "REGKEY: [CommonName] " & sRegKeysCommonName(i)
    Next i
    Status "Searching for CommonName folders..."
    For i = 1 To UBound(sFoldersCommonName)
        If sFoldersCommonName(i) = vbNullString Then Exit For
        If FolderExists(sFoldersCommonName(i)) Then Logg "FOLDER: [CommonName] " & sFoldersCommonName(i)
    Next i
    Status "Searching for CommonName files..."
    For i = 1 To UBound(sFilesCommonName)
        If sFilesCommonName(i) = vbNullString Then Exit For
        If FileExists(sFilesCommonName(i)) Then Logg "FILE: [CommonName] " & sFilesCommonName(i)
    Next i
    For i = 1 To UBound(sRegvalsCommonName)
        If sRegvalsCommonName(i) = vbNullString Then Exit For
        If RegvalExists(sRegvalsCommonName(i)) Then Logg "REGVAL: [CommonName] " & sRegvalsCommonName(i)
    Next i
End Sub

Public Sub EnumNewDotNetComponents()
    Dim i%
    Status "Searching for NewDotNet processes..."
    For i = 1 To UBound(sProcessesNewDotNet)
        If sProcessesNewDotNet(i) = vbNullString Then Exit For
        If ProcessExists(sProcessesNewDotNet(i)) Then Logg "PROCESS: [NewDotNet] " & sProcessesNewDotNet(i)
    Next i
    Status "Searching for NewDotNet regkeys..."
    For i = 1 To UBound(sRegKeysNewDotNet)
        If sRegKeysNewDotNet(i) = vbNullString Then Exit For
        If RegKeyExists(sRegKeysNewDotNet(i)) Then Logg "REGKEY: [NewDotNet] " & sRegKeysNewDotNet(i)
    Next i
    Status "Searching for NewDotNet folders..."
    For i = 1 To UBound(sFoldersNewDotNet)
        If sFoldersNewDotNet(i) = vbNullString Then Exit For
        If FolderExists(sFoldersNewDotNet(i)) Then Logg "FOLDER: [NewDotNet] " & sFoldersNewDotNet(i)
    Next i
    Status "Searching for NewDotNet files..."
    For i = 1 To UBound(sFilesNewDotNet)
        If sFilesNewDotNet(i) = vbNullString Then Exit For
        If FileExists(sFilesNewDotNet(i)) Then Logg "FILE: [NewDotNet] " & sFilesNewDotNet(i)
    Next i
    For i = 1 To UBound(sRegvalsNewDotNet)
        If sRegvalsNewDotNet(i) = vbNullString Then Exit For
        If RegvalExists(sRegvalsNewDotNet(i)) Then Logg "REGVAL: [NewDotNet] " & sRegvalsNewDotNet(i)
    Next i
End Sub

Public Sub EnumWebHancerComponents()
    Dim i%
    Status "Searching for WebHancer processes..."
    For i = 1 To UBound(sProcessesWebHancer)
        If sProcessesWebHancer(i) = vbNullString Then Exit For
        If ProcessExists(sProcessesWebHancer(i)) Then Logg "PROCESS: [WebHancer] " & sProcessesWebHancer(i)
    Next i
    Status "Searching for WebHancer regkeys..."
    For i = 1 To UBound(sRegKeysWebHancer)
        If sRegKeysWebHancer(i) = vbNullString Then Exit For
        If RegKeyExists(sRegKeysWebHancer(i)) Then Logg "REGKEY: [WebHancer] " & sRegKeysWebHancer(i)
    Next i
    Status "Searching for WebHancer folders..."
    For i = 1 To UBound(sFoldersWebHancer)
        If sFoldersWebHancer(i) = vbNullString Then Exit For
        If FolderExists(sFoldersWebHancer(i)) Then Logg "FOLDER: [WebHancer] " & sFoldersWebHancer(i)
    Next i
    Status "Searching for WebHancer files..."
    For i = 1 To UBound(sFilesWebHancer)
        If sFilesWebHancer(i) = vbNullString Then Exit For
        If FileExists(sFilesWebHancer(i)) Then Logg "FILE: [WebHancer] " & sFilesWebHancer(i)
    Next i
    For i = 1 To UBound(sRegvalsWebHancer)
        If sRegvalsWebHancer(i) = vbNullString Then Exit For
        If RegvalExists(sRegvalsWebHancer(i)) Then Logg "REGVAL: [WebHancer] " & sRegvalsWebHancer(i)
    Next i
End Sub

Public Sub EnumMediaLoadsComponents()
    Dim i%
    Status "Searching for MediaLoads processes..."
    For i = 1 To UBound(sProcessesMediaLoads)
        If sProcessesMediaLoads(i) = vbNullString Then Exit For
        If ProcessExists(sProcessesMediaLoads(i)) Then Logg "PROCESS: [MediaLoads] " & sProcessesMediaLoads(i)
    Next i
    Status "Searching for MediaLoads regkeys..."
    For i = 1 To UBound(sRegKeysMediaLoads)
        If sRegKeysMediaLoads(i) = vbNullString Then Exit For
        If RegKeyExists(sRegKeysMediaLoads(i)) Then Logg "REGKEY: [MediaLoads] " & sRegKeysMediaLoads(i)
    Next i
    Status "Searching for MediaLoads folders..."
    For i = 1 To UBound(sFoldersMediaLoads)
        If sFoldersMediaLoads(i) = vbNullString Then Exit For
        If FolderExists(sFoldersMediaLoads(i)) Then Logg "FOLDER: [MediaLoads] " & sFoldersMediaLoads(i)
    Next i
    Status "Searching for MediaLoads files..."
    For i = 1 To UBound(sFilesMediaLoads)
        If sFilesMediaLoads(i) = vbNullString Then Exit For
        If FileExists(sFilesMediaLoads(i)) Then Logg "FILE: [MediaLoads] " & sFilesMediaLoads(i)
    Next i
    For i = 1 To UBound(sRegvalsMediaLoads)
        If sRegvalsMediaLoads(i) = vbNullString Then Exit For
        If RegvalExists(sRegvalsMediaLoads(i)) Then Logg "REGVAL: [MediaLoads] " & sRegvalsMediaLoads(i)
    Next i
End Sub

Public Sub EnumSaveNowComponents()
    Dim i%
    Status "Searching for SaveNow processes..."
    For i = 1 To UBound(sProcessesSaveNow)
        If sProcessesSaveNow(i) = vbNullString Then Exit For
        If ProcessExists(sProcessesSaveNow(i)) Then Logg "PROCESS: [SaveNow] " & sProcessesSaveNow(i)
    Next i
    Status "Searching for SaveNow regkeys..."
    For i = 1 To UBound(sRegKeysSaveNow)
        If sRegKeysSaveNow(i) = vbNullString Then Exit For
        If RegKeyExists(sRegKeysSaveNow(i)) Then Logg "REGKEY: [SaveNow] " & sRegKeysSaveNow(i)
    Next i
    Status "Searching for SaveNow folders..."
    For i = 1 To UBound(sFoldersSaveNow)
        If sFoldersSaveNow(i) = vbNullString Then Exit For
        If FolderExists(sFoldersSaveNow(i)) Then Logg "FOLDER: [SaveNow] " & sFoldersSaveNow(i)
    Next i
    Status "Searching for SaveNow files..."
    For i = 1 To UBound(sFilesSaveNow)
        If sFilesSaveNow(i) = vbNullString Then Exit For
        If FileExists(sFilesSaveNow(i)) Then Logg "FILE: [SaveNow] " & sFilesSaveNow(i)
    Next i
    For i = 1 To UBound(sRegvalsSaveNow)
        If sRegvalsSaveNow(i) = vbNullString Then Exit For
        If RegvalExists(sRegvalsSaveNow(i)) Then Logg "REGVAL: [SaveNow] " & sRegvalsSaveNow(i)
    Next i
End Sub

Public Sub EnumDelfinComponents()
    Dim i%
    Status "Searching for Delfin processes..."
    For i = 1 To UBound(sProcessesDelfin)
        If sProcessesDelfin(i) = vbNullString Then Exit For
        If ProcessExists(sProcessesDelfin(i)) Then Logg "PROCESS: [Delfin] " & sProcessesDelfin(i)
    Next i
    Status "Searching for Delfin regkeys..."
    For i = 1 To UBound(sRegKeysDelfin)
        If sRegKeysDelfin(i) = vbNullString Then Exit For
        If RegKeyExists(sRegKeysDelfin(i)) Then Logg "REGKEY: [Delfin] " & sRegKeysDelfin(i)
    Next i
    Status "Searching for Delfin folders..."
    For i = 1 To UBound(sFoldersDelfin)
        If sFoldersDelfin(i) = vbNullString Then Exit For
        If FolderExists(sFoldersDelfin(i)) Then Logg "FOLDER: [Delfin] " & sFoldersDelfin(i)
    Next i
    Status "Searching for Delfin files..."
    For i = 1 To UBound(sFilesDelfin)
        If sFilesDelfin(i) = vbNullString Then Exit For
        If FileExists(sFilesDelfin(i)) Then Logg "FILE: [Delfin] " & sFilesDelfin(i)
    Next i
    For i = 1 To UBound(sRegvalsDelfin)
        If sRegvalsDelfin(i) = vbNullString Then Exit For
        If RegvalExists(sRegvalsDelfin(i)) Then Logg "REGVAL: [Delfin] " & sRegvalsDelfin(i)
    Next i
End Sub

Public Sub EnumOnFlowComponents()
    Dim i%
    Status "Searching for OnFlow processes..."
    For i = 1 To UBound(sProcessesOnFlow)
        If sProcessesOnFlow(i) = vbNullString Then Exit For
        If ProcessExists(sProcessesOnFlow(i)) Then Logg "PROCESS: [OnFlow] " & sProcessesOnFlow(i)
    Next i
    Status "Searching for OnFlow regkeys..."
    For i = 1 To UBound(sRegKeysOnFlow)
        If sRegKeysOnFlow(i) = vbNullString Then Exit For
        If RegKeyExists(sRegKeysOnFlow(i)) Then Logg "REGKEY: [OnFlow] " & sRegKeysOnFlow(i)
    Next i
    Status "Searching for OnFlow folders..."
    For i = 1 To UBound(sFoldersOnFlow)
        If sFoldersOnFlow(i) = vbNullString Then Exit For
        If FolderExists(sFoldersOnFlow(i)) Then Logg "FOLDER: [OnFlow] " & sFoldersOnFlow(i)
    Next i
    Status "Searching for OnFlow files..."
    For i = 1 To UBound(sFilesOnFlow)
        If sFilesOnFlow(i) = vbNullString Then Exit For
        If FileExists(sFilesOnFlow(i)) Then Logg "FILE: [OnFlow] " & sFilesOnFlow(i)
    Next i
    For i = 1 To UBound(sRegvalsOnFlow)
        If sRegvalsOnFlow(i) = vbNullString Then Exit For
        If RegvalExists(sRegvalsOnFlow(i)) Then Logg "REGVAL: [OnFlow] " & sRegvalsOnFlow(i)
    Next i
End Sub

Public Sub EnumOtherComponents()
    Dim i%
    Status "Searching for Other processes..."
    For i = 1 To UBound(sProcessesOther)
        If sProcessesOther(i) = vbNullString Then Exit For
        If ProcessExists(sProcessesOther(i)) Then Logg "PROCESS: [Other] " & sProcessesOther(i)
    Next i
    Status "Searching for Other regkeys..."
    For i = 1 To UBound(sRegKeysOther)
        If sRegKeysOther(i) = vbNullString Then Exit For
        If RegKeyExists(sRegKeysOther(i)) Then Logg "REGKEY: [Other] " & sRegKeysOther(i)
    Next i
    Status "Searching for Other folders..."
    For i = 1 To UBound(sFoldersOther)
        If sFoldersOther(i) = vbNullString Then Exit For
        If FolderExists(sFoldersOther(i)) Then Logg "FOLDER: [Other] " & sFoldersOther(i)
    Next i
    Status "Searching for Other files..."
    For i = 1 To UBound(sFilesOther)
        If sFilesOther(i) = vbNullString Then Exit For
        If FileExists(sFilesOther(i)) Then Logg "FILE: [Other] " & sFilesOther(i)
    Next i
    For i = 1 To UBound(sRegvalsOther)
        If sRegvalsOther(i) = vbNullString Then Exit For
        If RegvalExists(sRegvalsOther(i)) Then Logg "REGVAL: [Other] " & sRegvalsOther(i)
    Next i
End Sub

Public Sub EnumAltnetComponents()
    Dim i%
    Status "Searching for Altnet processes..."
    For i = 1 To UBound(sProcessesAltnet)
        If sProcessesAltnet(i) = vbNullString Then Exit For
        If ProcessExists(sProcessesAltnet(i)) Then Logg "PROCESS: [Altnet] " & sProcessesAltnet(i)
    Next i
    Status "Searching for Altnet regkeys..."
    For i = 1 To UBound(sRegKeysAltnet)
        If sRegKeysAltnet(i) = vbNullString Then Exit For
        If RegKeyExists(sRegKeysAltnet(i)) Then Logg "REGKEY: [Altnet] " & sRegKeysAltnet(i)
    Next i
    Status "Searching for Altnet folders..."
    For i = 1 To UBound(sFoldersAltnet)
        If sFoldersAltnet(i) = vbNullString Then Exit For
        If FolderExists(sFoldersAltnet(i)) Then Logg "FOLDER: [Altnet] " & sFoldersAltnet(i)
    Next i
    Status "Searching for Altnet files..."
    For i = 1 To UBound(sFilesAltnet)
        If sFilesAltnet(i) = vbNullString Then Exit For
        If FileExists(sFilesAltnet(i)) Then Logg "FILE: [Altnet] " & sFilesAltnet(i)
    Next i
    For i = 1 To UBound(sRegvalsAltnet)
        If sRegvalsAltnet(i) = vbNullString Then Exit For
        If RegvalExists(sRegvalsAltnet(i)) Then Logg "REGVAL: [Altnet] " & sRegvalsAltnet(i)
    Next i
End Sub

Public Sub EnumBullguardComponents()
    Dim i%
    Status "Searching for Bullguard processes..."
    For i = 1 To UBound(sProcessesBullguard)
        If sProcessesBullguard(i) = vbNullString Then Exit For
        If ProcessExists(sProcessesBullguard(i)) Then Logg "PROCESS: [Bullguard] " & sProcessesBullguard(i)
    Next i
    Status "Searching for Bullguard regkeys..."
    For i = 1 To UBound(sRegKeysBullguard)
        If sRegKeysBullguard(i) = vbNullString Then Exit For
        If RegKeyExists(sRegKeysBullguard(i)) Then Logg "REGKEY: [Bullguard] " & sRegKeysBullguard(i)
    Next i
    Status "Searching for Bullguard folders..."
    For i = 1 To UBound(sFoldersBullguard)
        If sFoldersBullguard(i) = vbNullString Then Exit For
        If FolderExists(sFoldersBullguard(i)) Then Logg "FOLDER: [Bullguard] " & sFoldersBullguard(i)
    Next i
    Status "Searching for Bullguard files..."
    For i = 1 To UBound(sFilesBullguard)
        If sFilesBullguard(i) = vbNullString Then Exit For
        If FileExists(sFilesBullguard(i)) Then Logg "FILE: [Bullguard] " & sFilesBullguard(i)
    Next i
    For i = 1 To UBound(sRegValsBullguard)
        If sRegValsBullguard(i) = vbNullString Then Exit For
        If RegvalExists(sRegValsBullguard(i)) Then Logg "REGVAL: [Bullguard] " & sRegValsBullguard(i)
    Next i
End Sub

Public Sub EnumGatorComponents()
    Dim i%
    Status "Searching for Gator processes..."
    For i = 1 To UBound(sProcessesGator)
        If sProcessesGator(i) = vbNullString Then Exit For
        If ProcessExists(sProcessesGator(i)) Then Logg "PROCESS: [Gator] " & sProcessesGator(i)
    Next i
    Status "Searching for Gator regkeys..."
    For i = 1 To UBound(sRegKeysGator)
        If sRegKeysGator(i) = vbNullString Then Exit For
        If RegKeyExists(sRegKeysGator(i)) Then Logg "REGKEY: [Gator] " & sRegKeysGator(i)
    Next i
    Status "Searching for Gator folders..."
    For i = 1 To UBound(sFoldersGator)
        If sFoldersGator(i) = vbNullString Then Exit For
        If FolderExists(sFoldersGator(i)) Then Logg "FOLDER: [Gator] " & sFoldersGator(i)
    Next i
    Status "Searching for Gator files..."
    For i = 1 To UBound(sFilesGator)
        If sFilesGator(i) = vbNullString Then Exit For
        If FileExists(sFilesGator(i)) Then Logg "FILE: [Gator] " & sFilesGator(i)
    Next i
    For i = 1 To UBound(sRegvalsGator)
        If sRegvalsGator(i) = vbNullString Then Exit For
        If RegvalExists(sRegvalsGator(i)) Then Logg "REGVAL: [Gator] " & sRegvalsGator(i)
    Next i
End Sub

Public Sub EnumMywayComponents()
    Dim i%
    Status "Searching for Myway processes..."
    For i = 1 To UBound(sProcessesMyWay)
        If sProcessesMyWay(i) = vbNullString Then Exit For
        If ProcessExists(sProcessesMyWay(i)) Then Logg "PROCESS: [Myway] " & sProcessesMyWay(i)
    Next i
    Status "Searching for Myway regkeys..."
    For i = 1 To UBound(sRegKeysMyWay)
        If sRegKeysMyWay(i) = vbNullString Then Exit For
        If RegKeyExists(sRegKeysMyWay(i)) Then Logg "REGKEY: [Myway] " & sRegKeysMyWay(i)
    Next i
    Status "Searching for Myway folders..."
    For i = 1 To UBound(sFoldersMyWay)
        If sFoldersMyWay(i) = vbNullString Then Exit For
        If FolderExists(sFoldersMyWay(i)) Then Logg "FOLDER: [Myway] " & sFoldersMyWay(i)
    Next i
    Status "Searching for Myway files..."
    For i = 1 To UBound(sFilesMyWay)
        If sFilesMyWay(i) = vbNullString Then Exit For
        If FileExists(sFilesMyWay(i)) Then Logg "FILE: [Myway] " & sFilesMyWay(i)
    Next i
    For i = 1 To UBound(sRegvalsMyWay)
        If sRegvalsMyWay(i) = vbNullString Then Exit For
        If RegvalExists(sRegvalsMyWay(i)) Then Logg "REGVAL: [Myway] " & sRegvalsMyWay(i)
    Next i
End Sub

Public Sub EnumP2PComponents()
    Dim i%
    Status "Searching for P2Pnetworking processes..."
    For i = 1 To UBound(sProcessesP2P)
        If sProcessesP2P(i) = vbNullString Then Exit For
        If ProcessExists(sProcessesP2P(i)) Then Logg "PROCESS: [P2Pnetworking] " & sProcessesP2P(i)
    Next i
    Status "Searching for P2Pnetworking regkeys..."
    For i = 1 To UBound(sRegKeysP2P)
        If sRegKeysP2P(i) = vbNullString Then Exit For
        If RegKeyExists(sRegKeysP2P(i)) Then Logg "REGKEY: [P2Pnetworking] " & sRegKeysP2P(i)
    Next i
    Status "Searching for P2Pnetworking folders..."
    For i = 1 To UBound(sFoldersP2P)
        If sFoldersP2P(i) = vbNullString Then Exit For
        If FolderExists(sFoldersP2P(i)) Then Logg "FOLDER: [P2Pnetworking] " & sFoldersP2P(i)
    Next i
    Status "Searching for P2Pnetworking files..."
    For i = 1 To UBound(sFilesP2P)
        If sFilesP2P(i) = vbNullString Then Exit For
        If FileExists(sFilesP2P(i)) Then Logg "FILE: [P2Pnetworking] " & sFilesP2P(i)
    Next i
    For i = 1 To UBound(sRegvalsP2P)
        If sRegvalsP2P(i) = vbNullString Then Exit For
        If RegvalExists(sRegvalsP2P(i)) Then Logg "REGVAL: [P2Pnetworking] " & sRegvalsP2P(i)
    Next i
End Sub

Public Sub EnumPerfectnavComponents()
    Dim i%
    Status "Searching for Perfectnav processes..."
    For i = 1 To UBound(sProcessesPerfectNav)
        If sProcessesPerfectNav(i) = vbNullString Then Exit For
        If ProcessExists(sProcessesPerfectNav(i)) Then Logg "PROCESS: [Perfectnav] " & sProcessesPerfectNav(i)
    Next i
    Status "Searching for Perfectnav regkeys..."
    For i = 1 To UBound(sRegKeysPerfectNav)
        If sRegKeysPerfectNav(i) = vbNullString Then Exit For
        If RegKeyExists(sRegKeysPerfectNav(i)) Then Logg "REGKEY: [Perfectnav] " & sRegKeysPerfectNav(i)
    Next i
    Status "Searching for Perfectnav folders..."
    For i = 1 To UBound(sFoldersPerfectNav)
        If sFoldersPerfectNav(i) = vbNullString Then Exit For
        If FolderExists(sFoldersPerfectNav(i)) Then Logg "FOLDER: [Perfectnav] " & sFoldersPerfectNav(i)
    Next i
    Status "Searching for Perfectnav files..."
    For i = 1 To UBound(sFilesPerfectNav)
        If sFilesPerfectNav(i) = vbNullString Then Exit For
        If FileExists(sFilesPerfectNav(i)) Then Logg "FILE: [Perfectnav] " & sFilesPerfectNav(i)
    Next i
    For i = 1 To UBound(sRegvalsPerfectNav)
        If sRegvalsPerfectNav(i) = vbNullString Then Exit For
        If RegvalExists(sRegvalsPerfectNav(i)) Then Logg "REGVAL: [Perfectnav] " & sRegvalsPerfectNav(i)
    Next i
End Sub

Public Function SharedFolderDeleteWarning() As Boolean
    Dim sMsg$
    sMsg = " ** WARNING ** " & vbCrLf & "If any files are left in your " & _
           "Shared Folder, they will be deleted!" & vbCrLf & "If you " & _
           "want to keep your files, move them to another folder " & _
           "before continuing." & vbCrLf & vbCrLf & "Continue " & _
           "with Kazaa uninstall?"
    If MsgBox(sMsg, vbExclamation + vbYesNo) = vbNo Then SharedFolderDeleteWarning = True
End Function

Public Function TrimNull$(s$)
    If InStr(s, Chr(0)) > 0 Then
        TrimNull = Left(s, InStr(s, Chr(0)) - 1)
    Else
        TrimNull = s
    End If
End Function

Public Sub ShellRun(sFilePath$)
    If FileExists(sFilePath) Then
        ShellExecute 0, "open", sFilePath, vbNullString, vbNullString, 1
    End If
End Sub

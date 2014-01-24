Attribute VB_Name = "modLSP"
Option Explicit
Private Declare Function RegEnumValue Lib "advapi32.dll" Alias "RegEnumValueA" (ByVal hKey As Long, ByVal dwIndex As Long, ByVal lpValueName As String, lpcbValueName As Long, ByVal lpReserved As Long, lpType As Long, lpData As Byte, lpcbData As Long) As Long
Private Declare Function RegCreateKeyEx Lib "advapi32.dll" Alias "RegCreateKeyExA" (ByVal hKey As Long, ByVal lpSubKey As String, ByVal Reserved As Long, ByVal lpClass As String, ByVal dwOptions As Long, ByVal samDesired As Long, lpSecurityAttributes As Any, phkResult As Long, lpdwDisposition As Long) As Long

Private Const REG_OPTION_NON_VOLATILE = 0

Private sKeyNameSpace$
Private sKeyProtocol$
Public bRebootNeeded As Boolean

Public Sub GetLSPCatalogNames()
    sKeyNameSpace = "System\CurrentControlSet\Services\WinSock2\Parameters"
    sKeyProtocol = "System\CurrentControlSet\Services\WinSock2\Parameters"
    
    sKeyNameSpace = sKeyNameSpace & "\" & RegGetString(HKEY_LOCAL_MACHINE, sKeyNameSpace, "Current_NameSpace_Catalog", "NameSpace_Catalog5")
    sKeyProtocol = sKeyProtocol & "\" & RegGetString(HKEY_LOCAL_MACHINE, sKeyProtocol, "Current_Protocol_Catalog", "Protocol_Catalog9")
End Sub

Public Sub CheckLSP()
    Dim lNumNameSpace&, lNumProtocol&, i&, j&
    Dim sFile$, uData() As Byte, hKey&
    lNumNameSpace = RegGetDword(HKEY_LOCAL_MACHINE, sKeyNameSpace, "Num_Catalog_Entries", 0)
    lNumProtocol = RegGetDword(HKEY_LOCAL_MACHINE, sKeyProtocol, "Num_Catalog_Entries", 0)
    
    'check for gaps in LSP chain
    For i = 1 To lNumNameSpace
        If RegKeyExists("HKLM\" & sKeyNameSpace & "\Catalog_Entries\" & Format(i, "000000000000")) Then
            'all fine & peachy
        Else
            'broken LSP detected!
            frmMain.lstLog.AddItem "REGKEY: [NewDotNet/Webhancer] Broken Winsock stack"
            Exit Sub
        End If
    Next i
    For i = 1 To lNumProtocol
        If RegKeyExists("HKLM\" & sKeyProtocol & "\Catalog_Entries\" & Format(i, "000000000000")) Then
            'all fine & dandy
        Else
            'shit, not again!
            frmMain.lstLog.AddItem "REGKEY: [NewDotNet/Webhancer] Broken Winsock stack"
            Exit Sub
        End If
    Next i
    
    'check all LSP providers are present
    For i = 1 To lNumNameSpace
        sFile = RegGetString(HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & Format(i, "000000000000"), "LibraryPath", "")
        sFile = Replace(sFile, "%SYSTEMROOT%", sWinDir, , , vbTextCompare)
        If sFile <> vbNullString And Dir(sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            'file ok
            If InStr(1, sFile, "webhdll.dll", vbTextCompare) > 0 Then
                frmMain.lstLog.AddItem "REGKEY: [NewDotNet] HKLM" & sKeyNameSpace & "\Catalog_Entries\" & Format(i, "000000000000")
            ElseIf InStr(1, sFile, "newdot", vbTextCompare) > 0 Then
                frmMain.lstLog.AddItem "REGKEY: [Webhancer] HKLM" & sKeyNameSpace & "\Catalog_Entries\" & Format(i, "000000000000")
            End If
        Else
            'damn, file is gone
            frmMain.lstLog.AddItem "REGKEY: [NewDotNet/Webhancer] Broken Internet access because of LSP provider '" & sFile & "' missing"
            Exit Sub
        End If
    Next i
    
    For i = 1 To lNumProtocol
        sFile = RegGetFileFromBinary(HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & Format(i, "000000000000"), "PackedCatalogItem")
        sFile = Replace(sFile, "%SYSTEMROOT%", sWinDir, , , vbTextCompare)
        If sFile <> vbNullString And Dir(sFile, vbArchive + vbReadOnly + vbHidden + vbSystem) <> vbNullString Then
            'file ok
            If InStr(sFile, "webhdll.dll") > 0 Then
                frmMain.lstLog.AddItem "REGKEY: [Webhancer] Hijacked Internet access by WebHancer"
            ElseIf InStr(sFile, "newdotnet") > 0 Then
                frmMain.lstLog.AddItem "REGKEY: [NewDotNet] Hijacked Internet access by New.Net"
            End If
        Else
            'damn - crossed again!
            frmMain.lstLog.AddItem "REGKEY: [NewDotNet/Webhancer] Broken Internet access because of LSP provider '" & sFile & "' missing"
            Exit Sub
        End If
    Next i
End Sub

Public Sub FixLSP()
    Dim lNumNameSpace&, lNumProtocol&
    Dim i&, j&, sFile$, hKey&, uData() As Byte
    lNumNameSpace = RegGetDword(HKEY_LOCAL_MACHINE, sKeyNameSpace, "Num_Catalog_Entries", 0)
    lNumProtocol = RegGetDword(HKEY_LOCAL_MACHINE, sKeyProtocol, "Num_Catalog_Entries", 0)
    
    '================================
    'check for missing/spyware files,
    'delete keys with those
    '================================
    For i = 1 To lNumNameSpace
        sFile = RegGetString(HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i), "LibraryPath", "")
        sFile = Replace(sFile, "%SYSTEMROOT%", sWinDir, , , vbTextCompare)
        If sFile <> vbNullString And Dir(sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            'file ok
            If InStr(1, sFile, "webhdll.dll", vbTextCompare) > 0 Or _
               InStr(1, sFile, "newdot", vbTextCompare) > 0 Then
                'it's New.Net/WebHancer! Kill it!
                On Error Resume Next
                Kill sFile  ' error 53 = file not found
                On Error GoTo 0:
                
                KillRegKey HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)
                lNumNameSpace = lNumNameSpace - 1
                
                'delete New.Net startup Reg entry
                KillRegVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "New.Net Startup"
                'delete WebHancer startup Reg entry
                KillRegVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "webHancer Agent"
                
                bRebootNeeded = True
            End If
        Else
            If RegKeyExists("HKLM\" & sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)) Then
                lNumNameSpace = lNumNameSpace - 1
                KillRegKey HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)
                bRebootNeeded = True
            End If
        End If
    Next i
    
    For i = 1 To lNumProtocol
        sFile = RegGetFileFromBinary(HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i), "PackedCatalogItem")
        sFile = Replace(sFile, "%SYSTEMROOT%", sWinDir, , , vbTextCompare)
        If sFile <> vbNullString And Dir(sFile, vbArchive + vbHidden + vbReadOnly + vbSystem) <> vbNullString Then
            'file ok
            If InStr(1, sFile, "webhdll.dll", vbTextCompare) > 0 Or _
               InStr(1, sFile, "newdot", vbTextCompare) > 0 Then
                'it's New.Net/WebHancer! Kill it!
                On Error Resume Next
                Kill sFile  ' error 53 = file not found
                On Error GoTo 0:
                
                KillRegKey HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)
                lNumNameSpace = lNumNameSpace - 1
                'delete New.Net startup Reg entry
                
                KillRegVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "New.Net Startup"
                'delete WebHancer startup Reg entry
                KillRegVal HKEY_LOCAL_MACHINE, "Software\Microsoft\Windows\CurrentVersion\Run", "webHancer Agent"
                bRebootNeeded = True
            End If
        Else
            If RegKeyExists("HKLM\" & sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)) Then
                lNumProtocol = lNumProtocol - 1
                KillRegKey HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)
                bRebootNeeded = True
            End If
        End If
    Next i
    
    '=====================================
    'check LSP chain, fix gaps where found
    '=====================================
    i = 1 'current LSP #
    j = 1 'correct LSP #
    Do
        If RegKeyExists("HKLM\" & sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)) Then
            If i > j Then
                RegRenameKey HKEY_LOCAL_MACHINE, sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i), sKeyNameSpace & "\Catalog_Entries\" & String(12 - Len(CStr(j)), "0") & CStr(j)
                bRebootNeeded = True
            End If
            j = j + 1
        Else
            'nothing, j stays the same
        End If
        i = i + 1
        'check to prevent infinite loop when
        'lNumNameSpace is wrong
        If i = 100 Then
            lNumNameSpace = j - 1
            Exit Do
        End If
    Loop Until j = lNumNameSpace + 1
    
    i = 1
    j = 1
    Do
        If RegKeyExists("HKLM\" & sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i)) Then
            If i > j Then
                RegRenameKey HKEY_LOCAL_MACHINE, sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(i)), "0") & CStr(i), sKeyProtocol & "\Catalog_Entries\" & String(12 - Len(CStr(j)), "0") & CStr(j)
                bRebootNeeded = True
            End If
            j = j + 1
        Else
            'nothing, j stays the same
        End If
        i = i + 1
        If i = 100 Then
            lNumProtocol = j - 1
            Exit Do
        End If
    Loop Until j = lNumProtocol + 1
    
    SetRegValDword HKEY_LOCAL_MACHINE, sKeyNameSpace, "Num_Catalog_Entries", lNumNameSpace
    SetRegValDword HKEY_LOCAL_MACHINE, sKeyProtocol, "Num_Catalog_Entries", lNumProtocol
End Sub

Private Sub RegRenameKey(lHive&, sKeyOldName$, sKeyNewName$)
    Dim hKey&, hKey2&, i&, j&, sName$, lType&, lDataLen&
    Dim sData$, lData&, uData() As Byte
    
    If RegOpenKeyEx(lHive, sKeyOldName, 0, KEY_QUERY_VALUE Or KEY_WRITE, hKey) <> 0 Then Exit Sub
    If RegOpenKeyEx(lHive, sKeyNewName, 0, KEY_QUERY_VALUE, hKey2) = 0 Then
        RegCloseKey hKey2
        RegDeleteKey lHive, sKeyNewName
    End If
    If RegCreateKeyEx(lHive, sKeyNewName, 0, vbNullString, REG_OPTION_NON_VOLATILE, KEY_WRITE, ByVal 0, hKey2, ByVal 0) <> 0 Then Exit Sub
    
    'assume key has no subkeys (which it does not have
    'where we use it for)
    
    i = 0
    sName = String(255, 0)
    ReDim uData(1024)
    lDataLen = 1024
    lType = 0
    If RegEnumValue(hKey, i, sName, 255, 0, lType, uData(0), lDataLen) <> 0 Then
        'no values to transfer
        RegCloseKey hKey
        RegCloseKey hKey2
        RegDeleteKey lHive, sKeyOldName
        Exit Sub
    End If
    
    Do
        sName = Left(sName, InStr(sName, Chr(0)) - 1)
        Select Case lType
            Case REG_SZ
                'reconstruct string
                sData = ""
                For j = 0 To lDataLen - 1
                    If uData(j) = 0 Then Exit For
                    sData = sData & Chr(uData(j))
                Next j
                RegSetValueEx hKey2, sName, 0, REG_SZ, ByVal sData, Len(sData)
                'RegDeleteValue hKey, sName
            Case REG_DWORD
                'reconstruct dword
                lData = 0
                lData = CLng(Val("&H" & _
                                 String(2 - Len(Hex(uData(3))), "0") & Hex(uData(3)) & _
                                 String(2 - Len(Hex(uData(2))), "0") & Hex(uData(2)) & _
                                 String(2 - Len(Hex(uData(1))), "0") & Hex(uData(1)) & _
                                 String(2 - Len(Hex(uData(0))), "0") & Hex(uData(0))))
                RegSetValueEx hKey2, sName, 0, REG_DWORD, lData, 4
                'RegDeleteValue hKey, sName
            Case REG_BINARY
                'at ease, soldier
                ReDim Preserve uData(lDataLen)
                RegSetValueEx hKey2, sName, 0, REG_BINARY, uData(0), UBound(uData)
                'RegDeleteValue hKey, sName
            Case Else
                'wtf?
        End Select
        
        i = i + 1
        sName = String(255, 0)
        ReDim uData(1024)
        lDataLen = 1024
        lType = 0
    Loop Until RegEnumValue(hKey, i, sName, 255, 0, lType, uData(0), lDataLen) <> 0
    RegCloseKey hKey
    RegCloseKey hKey2
    RegDeleteKey lHive, sKeyOldName
End Sub

Private Function RegGetFileFromBinary$(lHive&, sKey$, sValue$)
    Dim hKey&, uData() As Byte, sFile$
    Dim i&
    
    If RegOpenKeyEx(lHive, sKey, 0, KEY_QUERY_VALUE, hKey) = 0 Then
        ReDim uData(1024)
        If RegQueryValueEx(hKey, sValue, 0, 0, uData(0), 1024) = 0 Then
            sFile = ""
            For i = 0 To 1024
                If uData(i) = 0 Then Exit For
                sFile = sFile & Chr(uData(i))
            Next i
        End If
        RegCloseKey hKey
    End If
    RegGetFileFromBinary = sFile
End Function

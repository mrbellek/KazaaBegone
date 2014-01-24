Attribute VB_Name = "modProcess"
Option Explicit
Private Declare Function CreateToolhelpSnapshot Lib "kernel32" Alias "CreateToolhelp32Snapshot" (ByVal lFlags As Long, ByVal lProcessID As Long) As Long
Private Declare Function ProcessFirst Lib "kernel32" Alias "Process32First" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function ProcessNext Lib "kernel32" Alias "Process32Next" (ByVal hSnapshot As Long, uProcess As PROCESSENTRY32) As Long
Private Declare Function CloseHandle Lib "kernel32" (ByVal hObject As Long) As Long

Private Declare Function TerminateProcess Lib "kernel32" (ByVal hProcess As Long, ByVal uExitCode As Long) As Long

Private Declare Function EnumProcesses Lib "PSAPI.DLL" (ByRef lpidProcess As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long
Private Declare Function GetModuleFileNameExA Lib "PSAPI.DLL" (ByVal hProcess As Long, ByVal hModule As Long, ByVal ModuleName As String, ByVal nSize As Long) As Long
Private Declare Function OpenProcess Lib "kernel32" (ByVal dwDesiredAccess As Long, ByVal bInheritHandle As Long, ByVal dwProcessId As Long) As Long
Private Declare Function EnumProcessModules Lib "PSAPI.DLL" (ByVal hProcess As Long, ByRef lphModule As Long, ByVal cb As Long, ByRef cbNeeded As Long) As Long

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

Private Const TH32CS_SNAPPROCESS = &H2
Private Const PROCESS_TERMINATE = &H1
Private Const PROCESS_QUERY_INFORMATION = 1024
Private Const PROCESS_VM_READ = 16

Public Function ProcessExists(sFilePath$) As Boolean
    Dim sList$, i&, hProc&
    Dim hSnap&, uPE32 As PROCESSENTRY32, sExeFile$
    Dim lProcesses&(1 To 1024), lNeeded&, lNumProcesses&, sProcessName$, lModules&(1 To 1024)
    
    If InStr(sFilePath, "%") > 0 Then
        MsgBox "Found a % in sFilePath, that sucks!" & _
                vbCrLf & sFilePath, , "ProcessExists"
        Exit Function
    End If
    
    If Not bIsWinNT Then
        'windows 9x/me method
        hSnap = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)
        If hSnap > 0 Then
            uPE32.dwSize = Len(uPE32)
            If ProcessFirst(hSnap, uPE32) = 0 Then
                CloseHandle hSnap
                Exit Function
            End If
            
            Do
                sExeFile = TrimNull(uPE32.szExeFile)
                If InStr(1, sExeFile, sFilePath, vbTextCompare) > 0 Then
                    ProcessExists = True
                    CloseHandle hSnap
                    Exit Function
                End If
            Loop Until ProcessNext(hSnap, uPE32) = 0
            CloseHandle hSnap
        End If
    Else
        'windows nt/2k/xp/2003/etc method
        On Error Resume Next
        If EnumProcesses(lProcesses(1), CLng(1024) * 4, lNeeded) = 0 Then
            Exit Function
        End If
        lNumProcesses = lNeeded / 4
        For i = 1 To lNumProcesses
            hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ Or PROCESS_TERMINATE, 0, lProcesses(i))
            If hProc <> 0 Then
                lNeeded = 0
                sProcessName = String(260, 0)
                If EnumProcessModules(hProc, lModules(i), CLng(1024) * 4, lNeeded) <> 0 Then
                    GetModuleFileNameExA hProc, lModules(1), sProcessName, Len(sProcessName)
                    sProcessName = TrimNull(sProcessName)
                    If sProcessName <> vbNullString Then
                        If Left(sProcessName, 1) = "\" Then sProcessName = Mid(sProcessName, 2)
                        If Left(sProcessName, 3) = "??\" Then sProcessName = Mid(sProcessName, 4)
                        If InStr(1, sProcessName, "%SystemRoot%", vbTextCompare) > 0 Then sProcessName = Replace(sProcessName, "%SystemRoot%", sWinDir, , , vbTextCompare)
                        If InStr(1, sProcessName, "SystemRoot", vbTextCompare) > 0 Then sProcessName = Replace(sProcessName, "SystemRoot", sWinDir, , , vbTextCompare)
                        
                        If InStr(1, sProcessName, sFilePath, vbTextCompare) > 0 Then
                            ProcessExists = True
                            CloseHandle hProc
                            Exit Function
                        End If
                    End If
                End If
                CloseHandle hProc
            End If
        Next i
    End If
End Function

Public Sub KillProcess(sFilePath$)
    Dim sList$, i&, hProc&
    Dim hSnap&, uPE32 As PROCESSENTRY32, sExeFile$
    Dim lProcesses&(1 To 1024), lNeeded&, lNumProcesses&, sProcessName$, lModules&(1 To 1024)
    
    If InStr(sFilePath, "%") > 0 Then
        MsgBox "Found a % in sFilePath, that sucks!" & _
                vbCrLf & sFilePath, , "KillProcess"
        Exit Sub
    End If
    
    If Not bIsWinNT Then
        'windows 9x/me method
        hSnap = CreateToolhelpSnapshot(TH32CS_SNAPPROCESS, 0)
        If hSnap > 0 Then
            uPE32.dwSize = Len(uPE32)
            If ProcessFirst(hSnap, uPE32) = 0 Then
                CloseHandle hSnap
                Exit Sub
            End If
            
            Do
                sExeFile = TrimNull(uPE32.szExeFile)
                If InStr(1, sExeFile, sFilePath, vbTextCompare) > 0 Then
                    hProc = OpenProcess(PROCESS_TERMINATE, 0, uPE32.th32ProcessID)
                    If hProc <> 0 Then
                        TerminateProcess hProc, 0
                        CloseHandle hProc
                    End If
                    Exit Sub
                End If
            Loop Until ProcessNext(hSnap, uPE32) = 0
            CloseHandle hSnap
        End If
    Else
        'windows nt/2k/xp/2003/etc method
        On Error Resume Next
        If EnumProcesses(lProcesses(1), CLng(1024) * 4, lNeeded) = 0 Then
            Exit Sub
        End If
        lNumProcesses = lNeeded / 4
        For i = 1 To lNumProcesses
            hProc = OpenProcess(PROCESS_QUERY_INFORMATION Or PROCESS_VM_READ Or PROCESS_TERMINATE, 0, lProcesses(i))
            If hProc <> 0 Then
                lNeeded = 0
                sProcessName = String(260, 0)
                If EnumProcessModules(hProc, lModules(i), CLng(1024) * 4, lNeeded) <> 0 Then
                    GetModuleFileNameExA hProc, lModules(1), sProcessName, Len(sProcessName)
                    sProcessName = TrimNull(sProcessName)
                    If sProcessName <> vbNullString Then
                        If Left(sProcessName, 1) = "\" Then sProcessName = Mid(sProcessName, 2)
                        If Left(sProcessName, 3) = "??\" Then sProcessName = Mid(sProcessName, 4)
                        If InStr(1, sProcessName, "%SystemRoot%", vbTextCompare) > 0 Then sProcessName = Replace(sProcessName, "%SystemRoot%", sWinDir, , , vbTextCompare)
                        If InStr(1, sProcessName, "SystemRoot", vbTextCompare) > 0 Then sProcessName = Replace(sProcessName, "SystemRoot", sWinDir, , , vbTextCompare)
                        
                        If InStr(1, sProcessName, sFilePath, vbTextCompare) > 0 Then
                            TerminateProcess hProc, 0
                            CloseHandle hProc
                            Exit Sub
                        End If
                    End If
                End If
                CloseHandle hProc
            End If
        Next i
    End If
End Sub

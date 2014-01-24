Attribute VB_Name = "modLSP"
Option Explicit
Private Declare Function WSAStartup Lib "ws2_32.dll" (ByVal wVR As Long, lpWSAD As WSAData) As Long
Private Declare Function WSACleanup Lib "ws2_32.dll" () As Long

Private Declare Function WSAEnumProtocols Lib "ws2_32.dll" Alias "WSAEnumProtocolsA" (ByVal lpiProtocols As Long, lpProtocolBuffer As Any, lpdwBufferLength As Long) As Long
Private Declare Function WSAEnumNameSpaceProviders Lib "ws2_32.dll" Alias "WSAEnumNameSpaceProvidersA" (lpdwBufferLength As Long, lpnspBuffer As Any) As Long
Private Declare Function WSCDeinstallProvider Lib "ws2_32.dll" (ByVal lpProviderId As Long, ByRef lpErrno As Long) As Long
Private Declare Function WSCUnInstallNameSpace Lib "ws2_32.dll" (ByVal lpProviderId As Long) As Long

Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (ByRef pDest As Any, ByRef pSource As Any, ByVal Length As Long)
Private Declare Function lstrcpy Lib "kernel32.dll" Alias "lstrcpyA" (ByVal lpString1 As String, lpString2 As Any) As String
Private Declare Function StringFromGUID2 Lib "ole32.dll" (rguid As Any, ByVal lpsz As String, ByVal cchMax As Long) As Long

Private Type GUID
    Data1 As Long
    Data2 As Integer
    Data3 As Integer
    Data4(0 To 7) As Byte
End Type

Private Type WSAData
    wVersion As Integer
    wHighVersion As Integer
    szDescription As String * 257
    szSystemStatus As String * 129
    iMaxSockets As Integer
    iMaxUdpDg As Integer
    lpVendorInfo As Long
End Type

Private Type WSANAMESPACE_INFO
    NSProviderId As GUID
    dwNameSpace As Long
    fActive As Long
    dwVersion As Long
    lpszIdentifier As Long
End Type

Private Type WSAPROTOCOLCHAIN
    ChainLen As Long
    ChainEntries(6) As Long
End Type

Private Type WSAPROTOCOL_INFO
    dwServiceFlags1 As Long
    dwServiceFlags2 As Long
    dwServiceFlags3 As Long
    dwServiceFlags4 As Long
    dwProviderFlags As Long
    ProviderId As GUID
    dwCatalogEntryId As Long
    ProtocolChain As WSAPROTOCOLCHAIN
    iVersion As Long
    iAddressFamily As Long
    iMaxSockAddr As Long
    iMinSockAddr As Long
    iSocketType As Long
    iProtocol As Long
    iProtocolMaxOffset As Long
    iNetworkByteOrder As Long
    iSecurityScheme As Long
    dwMessageSize As Long
    dwProviderReserved As Long
    szProtocol As String * 256
End Type

Public bRebootNeeded As Boolean
Public sLSPBlacklist$()

Public Function CheckWinsockLSP() As Boolean
    If RunningInIDE Then Exit Function
    CheckWinsockProtocols
    CheckWinsockNameSpaces
End Function

Private Sub CheckWinsockProtocols()
    Dim uWSAData As WSAData, i%, j%
    Dim uWSAProtInfo As WSAPROTOCOL_INFO
    Dim uBuffer() As Byte, lBufferSize&
    Dim lNumProtocols&, sName$
    
    If WSAStartup(&H202, uWSAData) <> 0 Then
        Exit Sub
    End If
    
    ReDim uBuffer(1)
    WSAEnumProtocols 0, uBuffer(0), lBufferSize
    ReDim uBuffer(lBufferSize - 1)
    
    lNumProtocols = WSAEnumProtocols(0, uBuffer(0), lBufferSize)
    If lNumProtocols <> -1 Then
        For i = 0 To lNumProtocols - 1
            CopyMemory uWSAProtInfo, uBuffer(i * Len(uWSAProtInfo)), Len(uWSAProtInfo)
            sName = TrimNull(uWSAProtInfo.szProtocol)
            'sCLSID = GuidToString(uWSAProtInfo.ProviderId)
            
            For j = 0 To UBound(sLSPBlacklist)
                If j > UBound(sLSPBlacklist) Then Exit For
                If InStr(1, sName, sLSPBlacklist(j), vbTextCompare) > 0 Then
                    'lsp protocol is on blacklist
                    Logg "WINSOCK: [" & sLSPBlacklist(j) & "] " & sName & " (protocol)"
                End If
            Next j
        Next i
    End If
                        
    Do
    Loop Until WSACleanup() = -1
End Sub

Private Sub CheckWinsockNameSpaces()
    Dim uWSAData As WSAData, i%, j%
    Dim uWSANameSpaceInfo As WSANAMESPACE_INFO
    Dim lNumNameSpaces&, uBuffer() As Byte
    Dim lBufferSize&, sName$
    
    If WSAStartup(&H202, uWSAData) <> 0 Then
        Exit Sub
    End If
    
    ReDim uBuffer(1)
    WSAEnumNameSpaceProviders lBufferSize, uBuffer(0)
    ReDim uBuffer(lBufferSize - 1)
    
    lNumNameSpaces = WSAEnumNameSpaceProviders(lBufferSize, uBuffer(0))
    If lNumNameSpaces <> -1 Then
        For i = 0 To lNumNameSpaces - 1
            CopyMemory uWSANameSpaceInfo, uBuffer(i * Len(uWSANameSpaceInfo)), Len(uWSANameSpaceInfo)
            sName = String(255, 0)
            lstrcpy sName, ByVal uWSANameSpaceInfo.lpszIdentifier
            sName = TrimNull(sName)
            'sCLSID = GuidToString(uWSANameSpaceInfo.NSProviderId)
            
            For j = 0 To UBound(sLSPBlacklist)
                'this line is actually needed O_o
                If j > UBound(sLSPBlacklist) Then Exit For
                If InStr(1, sName, sLSPBlacklist(j), vbTextCompare) Then
                    'lsp namespace is on blacklist
                    Logg "WINSOCK: [" & sLSPBlacklist(j) & "] " & sName & " (namespace)"
                End If
            Next j
        Next i
    End If
    
    Do
    Loop Until WSACleanup() = -1
End Sub

Public Sub FixWinsockLSP(sItem$)
    If InStr(sItem, "(namespace)") > 0 Then
        KillLSPNameSpace sItem
    ElseIf InStr(sItem, "(protocol)") > 0 Then
        KillLSPProtocol sItem
    End If
End Sub

Private Sub KillLSPProtocol(sItem$)
    Dim uWSAData As WSAData, i%
    Dim uWSAProtInfo As WSAPROTOCOL_INFO
    Dim uBuffer() As Byte, lBufferSize&, lErr&
    Dim lNumProtocols&, sName$
    
    If WSAStartup(&H202, uWSAData) <> 0 Then
        'failed to start winsock
        Exit Sub
    End If
    
    ReDim uBuffer(1)
    WSAEnumProtocols 0, uBuffer(0), lBufferSize
    ReDim uBuffer(lBufferSize - 1)
    
    lNumProtocols = WSAEnumProtocols(0, uBuffer(0), lBufferSize)
    If lNumProtocols <> -1 Then
        For i = 0 To lNumProtocols - 1
            CopyMemory uWSAProtInfo, uBuffer(i * Len(uWSAProtInfo)), Len(uWSAProtInfo)
            sName = TrimNull(uWSAProtInfo.szProtocol)
            
            If InStr(1, sItem, sName, vbTextCompare) > 0 Then
                'nail the sucker
                WSCDeinstallProvider VarPtr(uWSAProtInfo.ProviderId), lErr
                bRebootNeeded = True
            End If
        Next i
    End If
    
    Do
    Loop Until WSACleanup() = -1
End Sub

Private Sub KillLSPNameSpace(sItem$)
    Dim uWSAData As WSAData, i%
    Dim uWSANameSpaceInfo As WSANAMESPACE_INFO
    Dim lNumNameSpaces&, uBuffer() As Byte
    Dim lBufferSize&, sName$
    
    If WSAStartup(&H202, uWSAData) <> 0 Then
        'failed to start winsock
        Exit Sub
    End If
    
    ReDim uBuffer(1)
    WSAEnumNameSpaceProviders lBufferSize, uBuffer(0)
    ReDim uBuffer(lBufferSize - 1)
    
    lNumNameSpaces = WSAEnumNameSpaceProviders(lBufferSize, uBuffer(0))
    If lNumNameSpaces <> -1 Then
        For i = 0 To lNumNameSpaces - 1
            CopyMemory uWSANameSpaceInfo, uBuffer(i * Len(uWSANameSpaceInfo)), Len(uWSANameSpaceInfo)
            sName = String(255, 0)
            lstrcpy sName, ByVal uWSANameSpaceInfo.lpszIdentifier
            sName = TrimNull(sName)
            
            If InStr(1, sItem, sName, vbTextCompare) > 0 Then
                'nail the sucker
                WSCUnInstallNameSpace VarPtr(uWSANameSpaceInfo)
                bRebootNeeded = True
            End If
        Next i
    End If
    
    Do
    Loop Until WSACleanup() = -1
End Sub

Private Function GuidToString$(uGuid As GUID)
    Dim sGUID$
    sGUID = String(80, 0)
    If StringFromGUID2(uGuid, sGUID, Len(sGUID)) > 0 Then
        GuidToString = StrConv(sGUID, vbFromUnicode)
        GuidToString = TrimNull(GuidToString)
    End If
End Function

Private Function TrimNull$(s$)
    If InStr(s, Chr(0)) = 0 Then
        TrimNull = s
        Exit Function
    Else
        TrimNull = Left(s, InStr(s, Chr(0)) - 1)
    End If
End Function


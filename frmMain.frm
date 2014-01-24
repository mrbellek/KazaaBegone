VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "KazaaBegone"
   ClientHeight    =   5790
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5535
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5790
   ScaleWidth      =   5535
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtNothing 
      Alignment       =   2  'Center
      BorderStyle     =   0  'None
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Left            =   1920
      Locked          =   -1  'True
      TabIndex        =   17
      Text            =   "Nothing found!"
      Top             =   1680
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.Frame fraFrame 
      Height          =   1815
      Left            =   120
      TabIndex        =   1
      Top             =   3600
      Width           =   5295
      Begin VB.OptionButton optScan 
         Caption         =   "Destroy only checked components"
         Enabled         =   0   'False
         Height          =   255
         Index           =   2
         Left            =   240
         TabIndex        =   19
         Top             =   600
         Width           =   3375
      End
      Begin VB.CheckBox chkAllowUndo 
         Caption         =   "Delete files to Recycle Bin"
         Enabled         =   0   'False
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   1440
         Value           =   1  'Checked
         Width           =   2655
      End
      Begin VB.CommandButton cmdGO 
         Caption         =   "GO"
         Default         =   -1  'True
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   735
         Left            =   3840
         TabIndex        =   4
         Top             =   360
         Width           =   1215
      End
      Begin VB.OptionButton optScan 
         Caption         =   "Search && destroy all installed components"
         Height          =   255
         Index           =   1
         Left            =   240
         TabIndex        =   3
         Top             =   960
         Width           =   3375
      End
      Begin VB.OptionButton optScan 
         Caption         =   "Search for installed components only"
         Height          =   255
         Index           =   0
         Left            =   240
         TabIndex        =   2
         Top             =   240
         Value           =   -1  'True
         Width           =   3375
      End
   End
   Begin VB.ListBox lstLog 
      Height          =   3420
      IntegralHeight  =   0   'False
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   5295
   End
   Begin VB.Frame fraDebug 
      Caption         =   "Add lines to definition files"
      Height          =   4815
      Left            =   120
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   5295
      Begin VB.TextBox txtDebugAdd 
         Height          =   735
         Left            =   240
         MultiLine       =   -1  'True
         TabIndex        =   16
         Top             =   600
         Width           =   4815
      End
      Begin VB.CommandButton cmdDebugCheck 
         Caption         =   "Check if above lines are new"
         Height          =   375
         Left            =   1200
         TabIndex        =   15
         Top             =   2280
         Width           =   3015
      End
      Begin VB.CommandButton cmdDebugWrite 
         Caption         =   "Copy to clipboard"
         Height          =   375
         Left            =   1200
         TabIndex        =   14
         Top             =   4320
         Width           =   3015
      End
      Begin VB.ListBox lstDebugNew 
         Height          =   1065
         IntegralHeight  =   0   'False
         Left            =   240
         TabIndex        =   13
         Top             =   3120
         Width           =   4815
      End
      Begin VB.OptionButton optDebugAdd 
         Caption         =   "Regkeys"
         Height          =   255
         Index           =   2
         Left            =   3480
         TabIndex        =   11
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton optDebugAdd 
         Caption         =   "Folders"
         Height          =   255
         Index           =   1
         Left            =   2040
         TabIndex        =   10
         Top             =   1800
         Width           =   1215
      End
      Begin VB.OptionButton optDebugAdd 
         Caption         =   "Files"
         Height          =   255
         Index           =   0
         Left            =   480
         TabIndex        =   8
         Top             =   1800
         Value           =   -1  'True
         Width           =   1215
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "New lines found:"
         Height          =   195
         Index           =   2
         Left            =   240
         TabIndex        =   12
         Top             =   2880
         Width           =   1200
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Above lines are all:"
         Height          =   195
         Index           =   1
         Left            =   240
         TabIndex        =   9
         Top             =   1560
         Width           =   1365
      End
      Begin VB.Label lblInfo 
         AutoSize        =   -1  'True
         Caption         =   "Paste lines here to check (multiline allowed):"
         Height          =   195
         Index           =   0
         Left            =   240
         TabIndex        =   7
         Top             =   360
         Width           =   3165
      End
   End
   Begin VB.ListBox lstLogSel 
      Height          =   3405
      IntegralHeight  =   0   'False
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   20
      Top             =   120
      Visible         =   0   'False
      Width           =   5295
   End
   Begin VB.Label lblStatus 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   0
      TabIndex        =   5
      Top             =   5520
      Width           =   5535
   End
   Begin VB.Menu mnuPopup 
      Caption         =   "Popup"
      Visible         =   0   'False
      Begin VB.Menu mnuPopupSelAll 
         Caption         =   "Select all"
      End
      Begin VB.Menu mnuPopupSelNone 
         Caption         =   "Select none"
      End
      Begin VB.Menu mnuPopupSelInv 
         Caption         =   "Invert selection"
      End
      Begin VB.Menu mnuPopupStr1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuPopupCopy 
         Caption         =   "Copy to clipboard.."
      End
      Begin VB.Menu mnuPopupSave 
         Caption         =   "Save to disk..."
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub cmdDebugCheck_Click()
    Dim i%, j%, vNewLines As Variant
    Dim bIsNewLine As Boolean
    
    lstDebugNew.Clear
    txtDebugAdd.Text = Replace(txtDebugAdd.Text, "HKEY_LOCAL_MACHINE\", "HKLM\")
    txtDebugAdd.Text = Replace(txtDebugAdd.Text, "HKEY_CURRENT_USER\", "HKCU\")
    txtDebugAdd.Text = Replace(txtDebugAdd.Text, "HKEY_CLASSES_ROOT\", "HKCR\")
    txtDebugAdd.Text = Replace(txtDebugAdd.Text, vbTab, vbNullString)
    'txtDebugAdd.Text = Replace(txtDebugAdd.Text, "c:\windows\", "c:\win98\", , , vbTextCompare)
    vNewLines = Split(txtDebugAdd.Text, vbCrLf)
    txtDebugAdd.Text = vbNullString
    
    If optDebugAdd(0).Value Then
        GoTo CheckFiles:
    ElseIf optDebugAdd(1).Value Then
        GoTo CheckFolders:
    ElseIf optDebugAdd(2).Value Then
        GoTo CheckRegkeys:
    Else
        Exit Sub
    End If
    
CheckFiles:
    For j = 0 To UBound(vNewLines)
        bIsNewLine = True
        For i = 1 To UBound(sFilesKazaa)
            If LCase(sFilesKazaa(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFilesBDE)
            If LCase(sFilesBDE(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFilesCyDoor)
            If LCase(sFilesCyDoor(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFilesCommonName)
            If LCase(sFilesCommonName(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFilesNewDotNet)
            If LCase(sFilesNewDotNet(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFilesWebHancer)
            If LCase(sFilesWebHancer(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFilesMediaLoads)
            If LCase(sFilesMediaLoads(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFilesSaveNow)
            If LCase(sFilesSaveNow(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFilesDelfin)
            If LCase(sFilesDelfin(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFilesOnFlow)
            If LCase(sFilesOnFlow(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFilesOther)
            If LCase(sFilesOther(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sFilesAltnet)
            If LCase(sFilesAltnet(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sFilesBullguard)
            If LCase(sFilesBullguard(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sFilesGator)
            If LCase(sFilesGator(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sFilesMyWay)
            If LCase(sFilesMyWay(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sFilesP2P)
            If LCase(sFilesP2P(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sFilesPerfectNav)
            If LCase(sFilesPerfectNav(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        
        
        If bIsNewLine Then lstDebugNew.AddItem vNewLines(j)
        
        If j Mod 10 = 0 Then
            Status "Progress: " & CStr(Int(100 * CLng(j) / UBound(vNewLines))) & " %"
        End If
    Next j
    GoTo EndOfSub
    
CheckFolders:
    For j = 0 To UBound(vNewLines)
        bIsNewLine = True
        For i = 1 To UBound(sFoldersKazaa)
            If LCase(sFoldersKazaa(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFoldersBDE)
            If LCase(sFoldersBDE(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFoldersCyDoor)
            If LCase(sFoldersCyDoor(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFoldersCommonName)
            If LCase(sFoldersCommonName(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFoldersNewDotNet)
            If LCase(sFoldersNewDotNet(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFoldersWebHancer)
            If LCase(sFoldersWebHancer(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFoldersMediaLoads)
            If LCase(sFoldersMediaLoads(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFoldersSaveNow)
            If LCase(sFoldersSaveNow(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFoldersDelfin)
            If LCase(sFoldersDelfin(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFoldersOnFlow)
            If LCase(sFoldersOnFlow(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sFoldersOther)
            If LCase(sFoldersOther(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sFoldersAltnet)
            If LCase(sFoldersAltnet(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sFoldersBullguard)
            If LCase(sFoldersBullguard(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sFoldersGator)
            If LCase(sFoldersGator(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sFoldersMyWay)
            If LCase(sFoldersMyWay(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sFoldersP2P)
            If LCase(sFoldersP2P(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sFoldersPerfectNav)
            If LCase(sFoldersPerfectNav(i)) = LCase(CStr(vNewLines(j))) Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        If bIsNewLine Then lstDebugNew.AddItem vNewLines(j)
    
        If j Mod 10 = 0 Then
            Status "Progress: " & CStr(Int(100 * j / UBound(vNewLines))) & " %"
        End If
    Next j
    GoTo EndOfSub

CheckRegkeys:
    For j = 0 To UBound(vNewLines)
        bIsNewLine = True
        For i = 1 To UBound(sRegKeysKazaa)
            If InStr(1, CStr(vNewLines(j)), sRegKeysKazaa(i), vbTextCompare) > 0 Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sRegKeysBDE)
            If InStr(1, CStr(vNewLines(j)), sRegKeysBDE(i), vbTextCompare) > 0 Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sRegKeysCyDoor)
            If InStr(1, CStr(vNewLines(j)), sRegKeysCyDoor(i), vbTextCompare) > 0 Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sRegKeysCommonName)
            If InStr(1, CStr(vNewLines(j)), sRegKeysCommonName(i), vbTextCompare) > 0 Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sRegKeysNewDotNet)
            If InStr(1, CStr(vNewLines(j)), sRegKeysNewDotNet(i), vbTextCompare) > 0 Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sRegKeysWebHancer)
            If InStr(1, CStr(vNewLines(j)), sRegKeysWebHancer(i), vbTextCompare) > 0 Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sRegKeysMediaLoads)
            If InStr(1, CStr(vNewLines(j)), sRegKeysMediaLoads(i), vbTextCompare) > 0 Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sRegKeysSaveNow)
            If InStr(1, CStr(vNewLines(j)), sRegKeysSaveNow(i), vbTextCompare) > 0 Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sRegKeysDelfin)
            If InStr(1, CStr(vNewLines(j)), sRegKeysDelfin(i), vbTextCompare) > 0 Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sRegKeysOnFlow)
            If InStr(1, CStr(vNewLines(j)), sRegKeysOnFlow(i), vbTextCompare) > 0 Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        For i = 1 To UBound(sRegKeysOther)
            If InStr(1, CStr(vNewLines(j)), sRegKeysOther(i), vbTextCompare) > 0 Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sRegKeysAltnet)
            If InStr(1, CStr(vNewLines(j)), sRegKeysAltnet(i), vbTextCompare) > 0 Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sRegKeysBullguard)
            If InStr(1, CStr(vNewLines(j)), sRegKeysBullguard(i), vbTextCompare) > 0 Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sRegKeysGator)
            If InStr(1, CStr(vNewLines(j)), sRegKeysGator(i), vbTextCompare) > 0 Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sRegKeysMyWay)
            If InStr(1, CStr(vNewLines(j)), sRegKeysMyWay(i), vbTextCompare) > 0 Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sRegKeysP2P)
            If InStr(1, CStr(vNewLines(j)), sRegKeysP2P(i), vbTextCompare) > 0 Then
                bIsNewLine = False
                Exit For
            End If
        Next i
        
        For i = 1 To UBound(sRegKeysPerfectNav)
            If InStr(1, CStr(vNewLines(j)), sRegKeysPerfectNav(i), vbTextCompare) > 0 Then
                bIsNewLine = False
                Exit For
            End If
        Next i

        
        If bIsNewLine Then lstDebugNew.AddItem vNewLines(j)
        
        If j Mod 10 = 0 Then
            Status "Progress: " & CStr(Int(100 * CLng(j) / UBound(vNewLines))) & " %"
            DoEvents
        End If
    Next j
    
EndOfSub:
    lblInfo(0).Visible = False
    txtDebugAdd.Visible = False
    lblInfo(1).Visible = False
    optDebugAdd(0).Visible = False
    optDebugAdd(1).Visible = False
    optDebugAdd(2).Visible = False
    cmdDebugCheck.Visible = False
    lblInfo(2).Top = 360
    lstDebugNew.Top = 600
    lstDebugNew.Height = 3585
    Status "Found " & lstDebugNew.ListCount & " new lines"
End Sub

Private Sub cmdDebugWrite_Click()
    Dim sBlah$, i%
    For i = 0 To lstDebugNew.ListCount - 1
        sBlah = sBlah & lstDebugNew.List(i) & vbCrLf
    Next i
    Clipboard.Clear
    Clipboard.SetText sBlah
    lstDebugNew.Clear
    
    
    lblInfo(2).Top = 2880
    lstDebugNew.Height = 1065
    lstDebugNew.Top = 3120
    lblInfo(0).Visible = True
    txtDebugAdd.Visible = True
    lblInfo(1).Visible = True
    optDebugAdd(0).Visible = True
    optDebugAdd(1).Visible = True
    optDebugAdd(2).Visible = True
    cmdDebugCheck.Visible = True
End Sub

Private Sub cmdGO_Click()
    lstLog.Clear
    txtNothing.Visible = False
        
    If optScan(0).Value Then 'scan only
        Scan
        
        If lstLog.ListCount = 0 Then
            txtNothing.Visible = True
            optScan(1).Enabled = False
        Else
            optScan(1).Enabled = True
            optScan(2).Enabled = True
        End If
    ElseIf optScan(2).Value Then 'delete selected
        If CheckKazaaRunning Then Exit Sub
        If SharedFolderDeleteWarning Then Exit Sub
        
        bUseRecycleBin = CBool(chkAllowUndo.Value)
        Destroy
    Else 'scan and destroy
        If CheckKazaaRunning Then Exit Sub
        If SharedFolderDeleteWarning Then Exit Sub
        
        Scan
        
        If lstLog.ListCount = 0 Then
            txtNothing.Visible = True
            optScan(1).Enabled = False
        Else
            bUseRecycleBin = CBool(chkAllowUndo.Value)
            Destroy
        End If
    End If
End Sub

Private Sub Form_Load()
    Me.Show
    Status "Loading..."
    cmdGO.Enabled = False
    DoEvents
    GetWindowsInfo
    LoadDefs
    ExpandVars
    Status "KazaaBG v" & App.Major & "." & _
           Format(App.Minor, "00") & "." & _
           App.Revision & " - written by Merijn" & _
           " - merijn.bellekom@gmail.com"
    cmdGO.Enabled = True
    
    ReDim sLSPBlacklist(1)
    sLSPBlacklist(0) = "New.Net"
    sLSPBlacklist(1) = "webHancer"
    
    'If 1 Then
    If InStr(Command, "/add") > 0 Then
        lstLog.Visible = False
        fraFrame.Visible = False
        fraDebug.Visible = True
        Status "Activated debug mode"
    End If
    Me.Show
    
    lstLog.AddItem ""
    lstLog.AddItem "   Welcome to KazaaBegone."
    lstLog.AddItem ""
    lstLog.AddItem "   This tool can remove all Kazaa versions up to Kazaa 2.5.1,"
    lstLog.AddItem "   and all the bundled software that comes with them."
    lstLog.AddItem ""
    lstLog.AddItem "   It will NOT remove KazaaLite, KazaaLite K++, Grokster, iMesh"
    lstLog.AddItem "   or any of its components, unless they use the same filenames,"
    lstLog.AddItem "   folder names or Registry keys as the original Kazaa."
    lstLog.AddItem ""
    lstLog.AddItem "   Kazaa versions newer than the versions supported by"
    lstLog.AddItem "   KazaaBegone in this versions may not be fully removed."
    lstLog.AddItem ""
    lstLog.AddItem "   Please choose to scan only or scan and remove anything found,"
    lstLog.AddItem "   and click 'GO'."
End Sub

Private Sub Form_Resize()
    On Error Resume Next
    If Me.WindowState = vbMinimized Then Exit Sub
    If Me.ScaleHeight < 5790 Then
        lstLog.Height = 3420
        lstLogSel.Height = 3420
        fraFrame.Top = 3600
        lblStatus.Top = 5520
        Exit Sub
    End If
    If Me.ScaleWidth < 5535 Then
        lstLog.Width = 5295
        lstLogSel.Width = 5295
        fraFrame.Width = 5295
        lblStatus.Width = 5535
        Exit Sub
    End If
    
    lstLog.Height = Me.ScaleHeight - 2775 + 405
    lstLogSel.Height = Me.ScaleHeight - 2775 + 405
    fraFrame.Top = Me.ScaleHeight - 2595 + 405
    lblStatus.Top = Me.ScaleHeight - 660 + 405
    txtNothing.Top = lstLog.Top + (lstLog.Height - txtNothing.Height) / 2
    
    lstLog.Width = Me.ScaleWidth - 360 + 120
    lstLogSel.Width = Me.ScaleWidth - 360 + 120
    fraFrame.Width = Me.ScaleWidth - 360 + 120
    cmdGO.Left = Me.ScaleWidth - 1815 + 120
    lblStatus.Width = Me.ScaleWidth
    txtNothing.Left = (Me.ScaleWidth - txtNothing.Width) / 2
End Sub

Private Sub lstLog_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And lstLog.ListCount > 0 Then
        mnuPopupSelAll.Visible = False
        mnuPopupSelNone.Visible = False
        mnuPopupSelInv.Visible = False
        mnuPopupStr1.Visible = False
        PopupMenu mnuPopup
    End If
End Sub

Private Sub lstLogSel_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    If Button = 2 And lstLogSel.ListCount > 0 Then
        mnuPopupSelAll.Visible = True
        mnuPopupSelNone.Visible = True
        mnuPopupSelInv.Visible = True
        mnuPopupStr1.Visible = True
        PopupMenu mnuPopup
    End If
End Sub

Private Sub mnuPopupCopy_Click()
    Dim i&, sLog$
    Status "Copying to clipboard..."
    If lstLog.Visible Then
        With lstLog
            For i = 0 To .ListCount - 1
                sLog = sLog & .List(i) & vbCrLf
            Next i
        End With
    Else
        With lstLogSel
            For i = 0 To .ListCount - 1
                sLog = sLog & .List(i) & vbCrLf
            Next i
        End With
    End If
    sLog = "KazaaBegone v" & App.Major & "." & App.Minor & vbCrLf & vbCrLf & sLog
    Clipboard.Clear
    Clipboard.SetText sLog
    Status "Done."
End Sub

Private Sub mnuPopupSave_Click()
    Dim sFile$, i&, sLog$
    sFile = CmnDlgSaveFile("Text files (*.txt)|*.txt|All files (*.*)|*.*", "Save results...", , App.Path)
    If sFile = vbNullString Then Exit Sub
    Status "Saving results..."
    If lstLog.Visible Then
        With lstLog
            For i = 0 To .ListCount - 1
                sLog = sLog & .List(i) & vbCrLf
            Next i
        End With
    Else
        With lstLogSel
            For i = 0 To .ListCount - 1
                sLog = sLog & .List(i) & vbCrLf
            Next i
        End With
    End If
    sLog = "KazaaBegone v" & App.Major & "." & App.Minor & vbCrLf & vbCrLf & sLog
    Open sFile For Output As #1
        Print #1, sLog
    Close #1
    Status "Results saved to " & sFile
End Sub

Private Sub mnuPopupSelAll_Click()
    Dim i&
    For i = 0 To lstLogSel.ListCount - 1
        lstLogSel.Selected(i) = True
    Next i
End Sub

Private Sub mnuPopupSelInv_Click()
    Dim i&
    For i = 0 To lstLogSel.ListCount - 1
        lstLogSel.Selected(i) = Not lstLogSel.Selected(i)
    Next i
End Sub

Private Sub mnuPopupSelNone_Click()
    Dim i&
    For i = 0 To lstLogSel.ListCount - 1
        lstLogSel.Selected(i) = False
    Next i
End Sub

Private Sub optScan_Click(Index As Integer)
    If Index = 0 Then
        chkAllowUndo.Enabled = False
    Else
        chkAllowUndo.Enabled = True
    End If
    
    If Index = 2 Then
        lstLog.Visible = False
        lstLogSel.Visible = True
    Else
        lstLog.Visible = True
        lstLogSel.Visible = False
    End If
End Sub

Attribute VB_Name = "modGlobal"
Option Explicit

'//Declares for displaying the hand cursor
Public Const IDC_HAND               As Long = 32649&    '//ID for hand cursor
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long

'//Declares for displaying opening file
Public Const HELP_CONTEXT           As Long = &H1
Public Const HELP_QUIT              As Long = &H2
Public Const HELP_CONTENTS          As Long = &H3&
Public Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hwnd As Long, ByVal lpFileName As String, ByVal wCommand As Long, ByVal dwData As Any) As Long

'//Declares for manipulating the richtextbox
Public Const EM_LINESCROLL          As Long = &HB6
Public Const EM_SETSEL              As Long = &HB1
Public Const EM_GETMODIFY           As Long = &HB8
Public Const EM_SETMODIFY           As Long = &HB9
Public Const EM_GETLINE             As Long = &HC4
Public Const EM_GETLINECOUNT        As Long = &HBA
Public Const EM_LINEINDEX           As Long = &HBB
Public Const EM_LINELENGTH          As Long = &HC1
Public Const EM_LINEFROMCHAR        As Long = &HC9
Public Const EM_GETFIRSTVISIBLELINE As Long = &HCE

'//Calls for controlling form resizing
Public OldWindowProc                As Long  ' Original window proc
Public Const WM_GETMINMAXINFO       As Long = &H24
Public Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Public Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (pDest As Any, pSource As Any, ByVal ByteLen As Long)
Private Declare Function CallWindowProc Lib "user32" Alias "CallWindowProcA" (ByVal lpPrevWndFunc As Long, ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Type POINTAPI
    px As Long
    py As Long
End Type
Private Type MINMAXINFO
    ptReserved                      As POINTAPI
    ptMaxSize                       As POINTAPI
    ptMaxPosition                   As POINTAPI
    ptMinTrackSize                  As POINTAPI
    ptMaxTrackSize                  As POINTAPI
End Type

'//Misc and shared declares
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32" (ByVal hwndLock As Long) As Long

Public Enum LockedOptions '------------- Enumerator options for LockWindowUpdate
    bLocked = 1
    bUnLocked = 0
End Enum

Public Type enumSetup '----------------- Script editor setup options
    Script              As String
    AppShortName        As String
    TestForFileNames    As Byte
    SecTagColor         As Long
    KeyTagColor         As Long
    ValTagColor         As Long
    DefaultWeb          As String
    DefaultConst        As String
    RegRisFiles         As Byte
End Type

Public rtHwnd           As Long '------- Hwnd of rtBox
Public bChanged         As Boolean '---- True when script has been edited
Public Setup            As enumSetup

Public Sub Main()
Dim args As String
    args = Trim$(Command$)
    If Len(args) > 4 Then
        If Left(args, 1) = Chr(34) Then args = Right(args, Len(args) - 1)
        If Right(args, 1) = Chr(34) Then args = Left(args, Len(args) - 1)
        Setup.Script = args '----------- Assign passed script file
    End If
    With Setup '------------------------ Get Editor settings from registry
        .TestForFileNames = GetSetting("ReVive Script Editor", "Config", "TestFileNames", 1)
        .SecTagColor = GetSetting("ReVive Script Editor", "Config", "SecTagColor", 8388608)
        .KeyTagColor = GetSetting("ReVive Script Editor", "Config", "KeyTagColor", 16711680)
        .ValTagColor = GetSetting("ReVive Script Editor", "Config", "ValTagColor", 13705184)
        .DefaultWeb = GetSetting("ReVive Script Editor", "Config", "DefaultWeb", "http://")
        .DefaultConst = GetSetting("ReVive Script Editor", "Config", "DefaultConst", "<sp>")
        .RegRisFiles = GetSetting("ReVive Script Editor", "Config", "RegRisFiles", "1")
    End With
    Load frmMain
End Sub

Public Sub SetWindowState(ByVal bState As LockedOptions)
'---------------------------------------------------------------------------------
' Purpose   : Toggles LockWindowUpdate, but only when call is either first request
'             to lock a window, or final request to unlock it.
' Note      : Written for locking updates on one expected hWnd.
'---------------------------------------------------------------------------------
Static lRequests As Long
    lRequests = IIf(bState, lRequests + 1, lRequests - 1)
    If lRequests < 1 Then
        Call LockWindowUpdate(0): lRequests = 0
    Else
        If lRequests = 1 Then Call LockWindowUpdate(rtHwnd)
    End If
End Sub

Public Function IsVersionValid(ByVal sVersion As String) As Boolean
'--------------------------------------------------------------------------
' Purpose:  Verifies version number meets format '0.0.0.0'. Only numbers
'           0 - 9 can be used in each version number segment, and there
'           must be 4 version number segments, like '10.0.3.99'.
'--------------------------------------------------------------------------
On Error GoTo Errs
Dim sVer()  As String
Dim i       As Byte
Dim x       As Byte
    sVer() = Split(sVersion, ".", , vbTextCompare)
    If UBound(sVer) = 3 Then '------------------------------- Verify exactly 4 segments exist
        For i = 0 To 3
            If Len(sVer(i)) = 0 Then Exit Function '--------- Verify segment is not 0 length
            For x = 1 To Len(sVer(i)) '---------------------- Verify each segment character is numeric
                If Not IsNumeric(Mid$(sVer(i), x, 1)) Then Exit Function
            Next x
        Next i
        IsVersionValid = True
    End If
Errs:
    Exit Function
End Function

Public Function SubClass_WndMessage(ByVal hwnd As Long, ByVal msg As Long, ByVal wp As Long, ByVal lp As Long) As Long
Dim MinMax As MINMAXINFO
    If msg = WM_GETMINMAXINFO Then
        CopyMemory MinMax, ByVal lp, Len(MinMax)
        MinMax.ptMinTrackSize.px = 8475 \ Screen.TwipsPerPixelX
        MinMax.ptMinTrackSize.py = 5925 \ Screen.TwipsPerPixelY
        CopyMemory ByVal lp, MinMax, Len(MinMax)
        SubClass_WndMessage = 1
        Exit Function
    End If
    SubClass_WndMessage = CallWindowProc(OldWindowProc, hwnd, msg, wp, lp)
End Function

Public Sub DisplayTip(ByVal Key As String, ByVal f As Form)
'---------------------------------------------------------------------------------
' Purpose   : Displays the tip text for the selected key on all necessary forms
'---------------------------------------------------------------------------------
Dim l As Byte
    l = Len(Key)
    With f.rtTip
        .SelBold = False
        .Text = Key & " - " & LoadResString(KeyToResource(Key))
        .SelStart = 0
        .SelLength = l
        .SelBold = True
    End With
End Sub

Public Function KeyToResource(ByVal Key As String) As Byte
'---------------------------------------------------------------------------------
' Purpose   : Returns a byte representing the Keys string index in the resource
'             files string table. Used to determine what help string to display.
'             This function makes it easier to setup new keys throughout app.
'---------------------------------------------------------------------------------
    Select Case Key
        Case "[Setup]":             KeyToResource = 1
        Case "AdminRequired":       KeyToResource = 2
        Case "AppShortName":        KeyToResource = 3
        Case "AppLongName":         KeyToResource = 4
        Case "ForceReboots":        KeyToResource = 5
        Case "LaunchIfKilled":      KeyToResource = 6
        Case "NotifyIcon":          KeyToResource = 7
        Case "RegRISFiles":         KeyToResource = 8
        Case "ScriptURLAlt":        KeyToResource = 9
        Case "ScriptURLPrim":       KeyToResource = 10
        Case "ShowFileIcons":       KeyToResource = 11
        Case "UpdateAppClass":      KeyToResource = 12
        Case "UpdateAppKill":       KeyToResource = 13
        Case "UpdateAppTitle":      KeyToResource = 14
        Case "[Files]":             KeyToResource = 15
        Case "Description":         KeyToResource = 16
        Case "DownloadURL":         KeyToResource = 17
        Case "FileSize":            KeyToResource = 18
        Case "InstallPath":         KeyToResource = 19
        Case "MustExist":           KeyToResource = 20
        Case "MustUpdate":          KeyToResource = 21
        Case "UpdateVersion":       KeyToResource = 22
        Case "UpdateMessage":       KeyToResource = 23
        Case "Directory Constants": KeyToResource = 24
        Case "<ap>":                KeyToResource = 25
        Case "<cf>":                KeyToResource = 26
        Case "<commondesktop>":     KeyToResource = 27
        Case "<commonstartmenu>":   KeyToResource = 28
        Case "<pf>":                KeyToResource = 29
        Case "<sp>":                KeyToResource = 30
        Case "<sys>":               KeyToResource = 31
        Case "<temp>":              KeyToResource = 32
        Case "<userdesktop>":       KeyToResource = 33
        Case "<userstartmenu>":     KeyToResource = 34
        Case "<win>":               KeyToResource = 35
    End Select
End Function

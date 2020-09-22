Attribute VB_Name = "modGlobal"
'***********************************************************************
'Chris Cochran          cwc.software@gmail.com        Updated: 11 Sep 05
'***********************************************************************
Option Explicit

'//Declares for GetFolderPath Routine
Public Enum CSIDL_VALUES
    CSIDL_STARTMENU = &HB '------------------ Values currently in use
    CSIDL_DESKTOPDIRECTORY = &H10
    CSIDL_COMMON_STARTMENU = &H16
    CSIDL_COMMON_DESKTOPDIRECTORY = &H19
    CSIDL_WINDOWS = &H24
    CSIDL_SYSTEM = &H25
    CSIDL_PROGRAM_FILES = &H26
    CSIDL_PROGRAM_FILES_COMMON = &H2B
    CSIDL_FLAG_PER_USER_INIT = &H800
'    CSIDL_DESKTOP = &H0 '------------------- Values available for future expansion
'    CSIDL_INTERNET = &H1
'    CSIDL_PROGRAMS = &H2
'    CSIDL_CONTROLS = &H3
'    CSIDL_PRINTERS = &H4
'    CSIDL_PERSONAL = &H5
'    CSIDL_FAVORITES = &H6
'    CSIDL_STARTUP = &H7
'    CSIDL_RECENT = &H8
'    CSIDL_SENDTO = &H9
'    CSIDL_BITBUCKET = &HA
'    CSIDL_MYDOCUMENTS = &HC
'    CSIDL_MYMUSIC = &HD
'    CSIDL_MYVIDEO = &HE
'    CSIDL_DRIVES = &H11
'    CSIDL_NETWORK = &H12
'    CSIDL_NETHOOD = &H13
'    CSIDL_FONTS = &H14
'    CSIDL_TEMPLATES = &H15
'    CSIDL_COMMON_PROGRAMS = &H17
'    CSIDL_COMMON_STARTUP = &H18
'    CSIDL_APPDATA = &H1A
'    CSIDL_PRINTHOOD = &H1B
'    CSIDL_LOCAL_APPDATA = &H1C
'    CSIDL_ALTSTARTUP = &H1D
'    CSIDL_COMMON_ALTSTARTUP = &H1E
'    CSIDL_COMMON_FAVORITES = &H1F
'    CSIDL_INTERNET_CACHE = &H20
'    CSIDL_COOKIES = &H21
'    CSIDL_HISTORY = &H22
'    CSIDL_COMMON_APPDATA = &H23
'    CSIDL_MYPICTURES = &H27
'    CSIDL_PROFILE = &H28
'    CSIDL_SYSTEMX86 = &H29
'    CSIDL_PROGRAM_FILESX86 = &H2A
'    CSIDL_PROGRAM_FILES_COMMONX86 = &H2C
'    CSIDL_COMMON_TEMPLATES = &H2D
'    CSIDL_COMMON_DOCUMENTS = &H2E
'    CSIDL_COMMON_ADMINTOOLS = &H2F
'    CSIDL_ADMINTOOLS = &H30
'    CSIDL_CONNECTIONS = &H31
'    CSIDL_COMMON_MUSIC = &H35
'    CSIDL_COMMON_PICTURES = &H36
'    CSIDL_COMMON_VIDEO = &H37
'    CSIDL_RESOURCES = &H38
'    CSIDL_RESOURCES_LOCALIZED = &H39
'    CSIDL_COMMON_OEM_LINKS = &H3A
'    CSIDL_CDBURN_AREA = &H3B
'    CSIDL_COMPUTERSNEARME = &H3D
'    CSIDL_FLAG_NO_ALIAS = &H1000
'    CSIDL_FLAG_DONT_VERIFY = &H4000
'    CSIDL_FLAG_CREATE = &H8000
'    CSIDL_FLAG_MASK = &HFF00
End Enum
Private Const S_OK                      As Long = 0
Private Const SHGFP_TYPE_CURRENT        As Long = &H0
Private Declare Function SHGetFolderPath Lib "shfolder.dll" Alias "SHGetFolderPathA" (ByVal hwndOwner As Long, ByVal nFolder As Long, ByVal hToken As Long, ByVal dwReserved As Long, ByVal lpszPath As String) As Long
Private Declare Function lstrlenW Lib "kernel32" (ByVal lpString As Long) As Long

'//Running App declares
Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hWnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hWnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Declare Function GetParent Lib "user32" (ByVal hWnd As Long) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32" (ByVal hWnd As Long, lpdwProcessId As Long) As Long
Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long)
Private Const WM_CLOSE                  As Long = &H10
Private pClassName                      As String
Private pSearchTitle                    As String
Private pAppFound                       As Boolean
Public Type typeRunningApps
    lWndHwnd()                          As Long
    lWndProcessID()                     As Long
End Type
Public tRunningApps                     As typeRunningApps

'//Windows Shutdown declares
Private Const ERROR_NOT_ALL_ASSIGNED    As Long = 1300
Private Const SE_PRIVILEGE_ENABLED      As Long = 2
Private Const TOKEN_QUERY               As Long = &H8
Private Const TOKEN_ADJUST_PRIVILEGES   As Long = &H20
Private Const EWX_REBOOT                As Long = 2
Private Type LUID
    lowpart                             As Long
    highpart                            As Long
End Type
Private Type LUID_AND_ATTRIBUTES
    pLuid                               As LUID
    Attributes                          As Long
End Type
Private Type TOKEN_PRIVILEGES
    PrivilegeCount                      As Long
    Privileges                          As LUID_AND_ATTRIBUTES
End Type
Private Declare Function AdjustTokenPrivileges Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal DisableAllPriv As Long, NewState As TOKEN_PRIVILEGES, ByVal BufferLength As Long, PreviousState As TOKEN_PRIVILEGES, ReturnLength As Long) As Long
Private Declare Function LookupPrivilegeValue Lib "advapi32.dll" Alias "LookupPrivilegeValueA" (ByVal lpSystemName As Any, ByVal lpName As String, lpUid As LUID) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Function ExitWindowsEx Lib "user32" (ByVal uFlags As Long, ByVal dwReserved As Long) As Long

'//DrawText declares
Private Const DT_WORD_ELLIPSIS      As Long = &H40000
Private Const DT_WORDBREAK          As Long = &H10
Public Const DT_FLAGS               As Long = DT_WORD_ELLIPSIS + DT_WORDBREAK '--- Used by all
Public Const DT_LEFT                As Long = &H0
Public Const DT_CENTER              As Long = &H1
Public Const DT_NOPREFIX            As Long = &H800
Public Declare Function DrawText Lib "user32" Alias "DrawTextA" (ByVal hdc As Long, ByVal lpStr As String, ByVal nCount As Long, lpRect As RECT, ByVal wFormat As Long) As Long

'//DrawBorder declares
Private Const BF_LEFT               As Long = &H1
Private Const BF_RIGHT              As Long = &H4
Private Const BF_TOP                As Long = &H2
Private Const BF_BOTTOM             As Long = &H8
Private Const BF_RECT               As Long = (BF_LEFT Or BF_TOP Or BF_RIGHT Or BF_BOTTOM)
Public Enum mBorderStyles
    RaiseShallow = &H4
    SunkenShallow = &H2
    RaisedHigh = &H5
    SunkenDeep = &HA
    Etched = &H6
    Bump = &H9
    FocusRect = &H99
End Enum
Private Declare Function DrawEdge Lib "user32" (ByVal hdc As Long, qrc As RECT, ByVal edge As Long, ByVal grfFlags As Long) As Long
Private Declare Function DrawFocusRect Lib "user32" (ByVal hdc As Long, lpRect As RECT) As Long

'//Gradient Fill Declares
Public Enum ePlane
    VERTICAL = 0
    HORIZONTAL = 1
End Enum
Private Type RGBColor
    R                   As Single
    G                   As Single
    B                   As Single
End Type
Public Type POINTAPI
    x                   As Long
    y                   As Long
End Type
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function MoveToEx Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, lpPoint As POINTAPI) As Long
Private Declare Function LineTo Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function CreatePen Lib "gdi32" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long

'//Finds "HotSpots" within a form (title bars, icons, text)
Public Declare Function PtInRect Lib "user32" (lpRect As RECT, ByVal ptx As Long, ByVal pty As Long) As Long

'//Moves form when titlebar "HotSpot" is selected
Public Declare Function ReleaseCapture Lib "user32" () As Long

'//Show cursor declares
Public Declare Function LoadCursor Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Public Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Public Const IDC_HAND           As Long = 32649&
Public Const IDC_SIZEALL        As Long = 32646&

'//Remove title bar declares
Private Declare Function CreateRectRgn Lib "gdi32" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32" (ByVal hWnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function GetWindowLong Lib "user32" Alias "GetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hWnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
'Already declared in this module - Private Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long
Private Const GWL_STYLE         As Long = (-16)
Private Const WS_CAPTION        As Long = &HC00000
Private Const SWP_FRAMECHANGED  As Long = &H20
Private Const SWP_NOZORDER      As Long = &H4
Public Const SWP_NOMOVE         As Long = &H2
Public Const SWP_NOSIZE         As Long = &H1
Private Const SWP_FLAGS         As Long = SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOZORDER Or SWP_NOSIZE

'//Remove menu items declares
Private Const MF_BYPOSITION     As Long = &H400&
Private Const WM_GETSYSMENU     As Long = &H313
Public Const HTCAPTION          As Long = 2
Public Const WM_NCLBUTTONDOWN   As Long = &HA1
Private Declare Function GetSystemMenu Lib "user32" (ByVal hWnd As Long, ByVal bRevert As Long) As Long
Private Declare Function RemoveMenu Lib "user32" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DrawMenuBar Lib "user32" (ByVal hWnd As Long) As Long

'//Paint specific hDC area by RECT declares (far more efficient than Me.Refresh)
Private Const RDW_INVALIDATE    As Long = &H1
Private Const RDW_UPDATENOW     As Long = &H100
Public Const RDW_FLAGS          As Long = RDW_INVALIDATE + RDW_UPDATENOW
Public Declare Function RedrawWindow Lib "user32" (ByVal hWnd As Long, lprcUpdate As RECT, ByVal hrgnUpdate As Long, ByVal fuRedraw As Long) As Long

'//ShellExecute Declares
Public Const SW_HIDE            As Long = 0 '--- Used for executing regsvr32.exe
Public Const SW_NORMAL          As Long = 1 '--- Used for restarting applications
Public Const SW_MAXIMIZE        As Long = 3 '--- Used when displaying HTML report
Public Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long

'//SetWindowPos declares
Public Const HWND_TOPMOST      As Long = -1
Public Const SWP_NOACTIVATE    As Long = &H10
Public Const SWP_SHOWWINDOW    As Long = &H40
Public Declare Function SetWindowPos Lib "user32" (ByVal hWnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cX As Long, ByVal cY As Long, ByVal wFlags As Long) As Long

'//Detect if in Terminal Services mode declares (***NOT YET IMPLEMENTED***)
'!Private Const SM_REMOTESESSION  As Long = &H1000
'!Private Declare Function GetSystemMetrics Lib "user32" (ByVal nIndex As Long) As Long

'//Misc and shared declares
Public Type RECT
    lLeft                       As Long
    lTop                        As Long
    lRight                      As Long
    lBottom                     As Long
End Type
Public Const MAX_PATH           As Long = 260
Public Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hWnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Public Declare Function SetRect Lib "user32" (lpRect As RECT, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Public Declare Function GetForegroundWindow Lib "user32" () As Long
Public Declare Function SetForegroundWindow Lib "user32" (ByVal hWnd As Long) As Long
Public Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Public Declare Function DrawIconEx Lib "user32" (ByVal lHdc As Long, ByVal xLeft As Long, ByVal yTop As Long, ByVal hIcon As Long, ByVal cxWidth As Long, ByVal cyWidth As Long, ByVal istepIfAniCur As Long, ByVal hbrFlickerFreeDraw As Long, ByVal diFlags As Long) As Long
Public Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
Public Declare Function RemoveDirectory Lib "kernel32" Alias "RemoveDirectoryA" (ByVal lpPathName As String) As Long
Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long

'//Private module level variables
Private lProcessToKill          As Long

'//Public application level variables
Public bADMIN                   As Boolean
Public bOS                      As Byte
Public sTEMPDIR                 As String   '//Temp folder location for downloaded files not yet proccessed
Public lPREVWINDOW              As Long
Public sUpdateMessage           As String   '//Collection of update messages for display on exit
'!Public bREMOTESESSION           As Boolean  '//True when running in a remote Terminal Services environment (***NOT YET IMPLEMENTED***)

Public Sub Main()
'***************************************************************************************************
'Explanation of three possible arguments: (NONE ARE REQUIRED)
'
'Usage: ReVive.exe /n /a scriptpath
'   /n              Notify: Check for updates and notify when they are available.
'
'   /a              Auto: Check for and install updates without notice.
'                   Does notify user when updates are complete. (is case reboot is required)
'                   Overrides /n argument.
'
'   scriptpath      Path of ReVive initilization script to use for updating.
'                   If not passed or not found, will assume App.Path & "\update.ris" of this exe.
'
'   EXAMPLE:        To run LiveUpdate in Notify mode execute:
'                       lResult = "C:\Progra~1\LiveUp~1\ReVive.exe /n C:\Progra~1\MyProg\myapp.ris"
'****************************************************************************************************
On Error GoTo Errs
Dim args    As String
Dim i       As Integer
Dim path    As String
Dim x       As Long
Dim sRISDir As String
Dim bRemote As Boolean

    If App.PrevInstance Then Exit Sub '----------------- Not a good idea to have two LiveUpdates running simultaneously
    lPREVWINDOW = GetForegroundWindow '----------------- Get previously active window
    args = Replace(LCase$(Command$), Chr(34), "") '----- Remove quotation marks if exist
    '!bREMOTESESSION = CBool(GetSystemMetrics(SM_REMOTESESSION)) (***NOT YET IMPLEMENTED***)
    
    If Len(args) Then
        '//Check for Auto mode
        If InStr(1, args, "/a", vbTextCompare) Then
            Setup.RunMode = eAUTO '--------------------- Auto mode
            args = Trim$(Replace(args, "/a", ""))
            args = Trim$(Replace(args, "/n", ""))
        '//Check for Notify mode if Auto wasn't specified
        ElseIf InStr(1, args, "/n", vbTextCompare) Then
            Setup.RunMode = eNOTIFY '------------------- Notify mode
            args = Trim$(Replace(args, "/n", ""))
        End If
        '//Check if setup script path was passed and if it long enough to constitute a script location
        If Len(args) > 6 Then
            '//We need to validate that the passed folder path exist. If the file doesn't
            '..we are OK because we will create it on exit for the next LiveUpdate run,
            '..but the folder path must exist to ensure integrity of the <sp> constant.
            args = Replace(args, Chr(34), "")
            If Dir$(Left$(args, InStrRev(args, "\") - 1), vbDirectory) <> "" Then
                Setup.SetupScriptPath = GetLongPath(args)
            Else
                '//The directory was not found so we must abort LiveUpdate
                If Setup.RunMode = eNORMAL Then '-- Inform the user if in Normal mode, else just exit
                    MsgBox "LiveUpdate could not locate the '" & Left$(args, InStrRev(args, "\") - 1) & _
                           "' directory as expected and must now exit.   " & vbNewLine & vbNewLine & _
                           "If this problem persist please contact the vendor of this software.   ", vbCritical, "LiveUpdate Initialization Script Directory Not Found"
                    
                End If
                Exit Sub
            End If
        End If
    End If
    
    '//If SetupScriptPath was not passed as an argument then default to update.ris in
    '..the App.Path directory. If this file is not found, we will be using ucDownload
    '..settings, and will create the update.ris file when exiting ReVive. If the
    '..ucDownload control does not specify the web script location, ReVive will return
    '..an "Unable to download update script" error.
    '//IMPORTANT: If you always distribute a ris file with your app and pass it, you're good.
    If Len(Setup.SetupScriptPath) = 0 Then
        Setup.SetupScriptPath = App.path & "\update.ris"
    End If
    '//Select a temp directory where the update.ris file is stored or App.Path when no update.ris
    '..file is found or specified.. Doing this ensures updated files do not need moved across volumes,
    '..which would cause us to lose the security descriptor attached to the file.
    sRISDir = Left$(Setup.SetupScriptPath, InStrRev(Setup.SetupScriptPath, "\", , vbTextCompare) - 1)
    sTEMPDIR = sRISDir & "\Temp\ReVive_0000"
    '//Create Temp folder if it does not exist
    If Dir$(sRISDir & "\Temp", vbDirectory) = "" Then
        MkDir sRISDir & "\Temp"
    End If
    '//Select and create a unique temp Revive directory that does not already exist
    Do While Dir$(sTEMPDIR, vbDirectory) <> ""
        x = x + 1
        sTEMPDIR = sRISDir & "\Temp\ReVive_" & Format(x, "0000")
    Loop
    MkDir sTEMPDIR
    bADMIN = IsAdministrator
    bOS = WindowsVersion
    Load frmMain
Errs_Exit:
    Exit Sub
Errs:
    MsgBox "LiveUpdate experienced the following unrecoverable error in Sub Main:" & vbNewLine & vbNewLine & _
            Err.Description & vbNewLine & vbNewLine & _
            "Contact your software vendor if this problem persist.", vbCritical, "LiveUpdate"
        Resume Errs_Exit
End Sub

Public Sub DrawForm(ByVal fForm As Form)
'-------------------------------------------------------------------------
' Purpose   : Central Sub to hide default title bar and draw new one, clip
'             old form region, and draw new border. Used for all forms.
'-------------------------------------------------------------------------
Dim h           As Long
Dim w           As Long
Dim nStyle      As Long
Dim hRgn        As Long
Dim lMenu       As Long
    With fForm
        h = .ScaleHeight
        w = .ScaleWidth
        nStyle = GetWindowLong(.hWnd, GWL_STYLE) '-------- Hide title bar
        nStyle = nStyle And Not WS_CAPTION
        Call SetWindowLong(.hWnd, GWL_STYLE, nStyle)
        SetWindowPos .hWnd, 0, 0, 0, 0, 0, SWP_FLAGS
        hRgn = CreateRectRgn(3, h + 3, w + 3, 3)
        Call SetWindowRgn(.hWnd, hRgn, True)
        DeleteObject hRgn
        Call DrawBorder(.hdc, 0, w, 0, h, RaisedHigh) '--- Draw new form border
        lMenu = GetSystemMenu(.hWnd, False) '------------- Remove Size menu item
        RemoveMenu lMenu, 2, MF_BYPOSITION
        DrawMenuBar .hWnd '------------------------------- Refresh system menu
    End With
End Sub

Public Sub DrawTitleBar(ByVal fForm As Form, ByVal State As WindowState, ByVal sCaption As String, Optional ByVal Buttons As Boolean = False)
Dim r1      As RECT
Dim r2      As RECT
Dim w       As Long
Dim lHdc    As Long
Dim pIcon   As IPictureDisp
    With fForm
        w = .ScaleWidth
        lHdc = .hdc
        '//Draw gradient title bar
        Call SetRect(r1, 2, 2, w - 5, 24)
        Call DrawGradient(lHdc, r1, IIf(State = Active, 3087635, 8487297), 14407116, HORIZONTAL)
        '//Draw title bar text with white text Active or grayed text InActive
        Call SetRect(r2, 6, 6, w - 45, 20)
        .ForeColor = IIf(State = Active, vbWhite, 13423575)
        .FontBold = True
        Call DrawText(lHdc, sCaption, -1, r2, DT_FLAGS + DT_LEFT + DT_NOPREFIX)
        '//Draw Close and Minimize buttons if requested
        If Buttons Then '---- Draw the Close and Minimize buttons from one icon
            Set pIcon = LoadResPicture(206, vbResIcon)
            Call DrawIconEx(fForm.hdc, w - 41, 5, pIcon.Handle, 34, 16, 0, 0, &H3)
            Set pIcon = Nothing
        End If
        '//Repaint only the title bar rect, NOT the entire window.
        '..To illustrate selective repainting change r1 to r2 below.
        Call RedrawWindow(.hWnd, r1, 0&, RDW_FLAGS)
    End With
End Sub

Public Sub Reboot()
'---------------------------------------------------------------------------------------
' Procedure : Reboot
' Author    : Dave Scarmozzino, "The Scarms", http://www.thescarms.com/vbasic/chgres.asp
' Purpose   : Reboots the computer when required.
'---------------------------------------------------------------------------------------
Dim tLuid          As LUID
Dim tTokenPriv     As TOKEN_PRIVILEGES
Dim tPrevTokenPriv As TOKEN_PRIVILEGES
Dim lResult        As Long
Dim lToken         As Long
Dim lLenBuffer     As Long

If bOS < 2 Then '--------- Forget all the AccessToken stuff below if Win9X or ME
    Call ExitWindowsEx(EWX_REBOOT, 0)
Else
    '
    ' Get the access token of the current process.  Get it
    ' with the privileges of querying the access token and
    ' adjusting its privileges.
    '
    lResult = OpenProcessToken(GetCurrentProcess(), TOKEN_ADJUST_PRIVILEGES Or TOKEN_QUERY, lToken)
    If lResult = 0 Then
        Exit Sub 'Failed
    End If
    '
    ' Get the locally unique identifier (LUID) which
    ' represents the shutdown privilege.
    '
    lResult = LookupPrivilegeValue(0&, "SeShutdownPrivilege", tLuid)
    If lResult = 0 Then Exit Sub 'Failed
    '
    ' Populate the new TOKEN_PRIVILEGES values with the LUID
    ' and allow your current process to shutdown the computer.
    '
    With tTokenPriv
        .PrivilegeCount = 1
        .Privileges.Attributes = SE_PRIVILEGE_ENABLED
        .Privileges.pLuid = tLuid
    lResult = AdjustTokenPrivileges(lToken, False, tTokenPriv, Len(tPrevTokenPriv), tPrevTokenPriv, lLenBuffer)
    End With
    
    If lResult = 0 Then
        Exit Sub 'Failed
    Else
        If Err.LastDllError = ERROR_NOT_ALL_ASSIGNED Then Exit Sub 'Failed
    End If
    '
    '  Shutdown Windows.
    '
    Call ExitWindowsEx(EWX_REBOOT, 0)
End If

End Sub

Public Function GetLongPath(ByVal sPath As String) As String
'//If called running Win95 this function will return the string passed to it.
'..This call is only used for consistency when displaying the update report.
On Error Resume Next
Dim lLength  As Long
Dim sBuff    As String
    sBuff = String$(MAX_PATH, 0)
    lLength = GetLongPathName(sPath, sBuff, Len(sBuff))
    If lLength And Err = 0 Then
        GetLongPath = Left$(sBuff, lLength)
    Else
        GetLongPath = sPath
    End If
    If Err Then Err.Clear
End Function

Public Function ProfileGetItem(ByVal lpSectionName As String, _
                               ByVal lpKeyName As String, _
                               ByVal defaultValue As String, _
                               ByVal inifile As String) As String
'************************************************************
'Written by Randy Birch, http://vbnet.mvps.org
'"Using INI Files to Save Application Data - The Basics"
'http://vbnet.mvps.org/index.html?code/file/pprofilebasic.htm
'************************************************************

'Retrieves a value from an ini file corresponding
'to the section and key name passed.
        
   Dim success As Long
   Dim nSize As Long
   Dim ret As String
  
  'call the API with the parameters passed.
  'The return value is the length of the string
  'in ret, including the terminating null. If a
  'default value was passed, and the section or
  'key name are not in the file, that value is
  'returned. If no default value was passed (""),
  'then success will = 0 if not found.

  'Pad a string large enough to hold the data.
   ret = Space$(2048)
   nSize = Len(ret)
   success = GetPrivateProfileString(lpSectionName, _
                                     lpKeyName, _
                                     defaultValue, _
                                     ret, _
                                     nSize, _
                                     inifile)
   
   If success Then
      ProfileGetItem = Left$(ret, success)
   End If
   
End Function

Public Function GetFolderPath(ByVal csidl As CSIDL_VALUES) As String
'****************************************************************************
'Adapted from code written by Randy Birch
'"Using SHGetFolderPath to Find Popular Shell Folders", http://vbnet.mvps.org
'Full post is here: http://vbnet.mvps.org/index.html?code/browse/csidl.htm
'****************************************************************************
Dim buff As String
    buff = Space$(MAX_PATH)
    If SHGetFolderPath(0, csidl Or CSIDL_FLAG_PER_USER_INIT, -1, _
       SHGFP_TYPE_CURRENT, buff) = S_OK Then
           GetFolderPath = Left$(buff, lstrlenW(StrPtr(buff)))
    End If
End Function

Public Sub DrawBorder(ByVal hdc As Long, ByVal LeftX As Long, _
    ByVal RightX As Long, ByVal TopY As Long, _
    ByVal BottomY As Long, _
    Optional BStyle As mBorderStyles = &H6)
'---------------------------------------------------------------------------------------
' Purpose   : Draws border as defined by mBorderStyles
'---------------------------------------------------------------------------------------
Dim R As RECT
    SetRect R, LeftX, TopY, RightX, BottomY '-- Set the rectangle's perimeter values
    If BStyle = FocusRect Then
        DrawFocusRect hdc, R
    Else
        DrawEdge hdc, R, BStyle, BF_RECT
    End If
End Sub

Public Sub DrawGradient(ByVal lHdc As Long, R As RECT, ByVal StartColor As Long, ByVal EndColor As Long, ByVal Direction As ePlane)
Dim s       As RGBColor   'Start RGB colors
Dim e       As RGBColor   'End RBG colors
Dim i       As RGBColor   'Increment RGB colors
Dim x       As Long
Dim lSteps  As Long
    lSteps = IIf(Direction, R.lRight - R.lLeft, R.lBottom - R.lTop)
    s.R = (StartColor And &HFF)
    s.G = (StartColor \ &H100) And &HFF
    s.B = (StartColor And &HFF0000) / &H10000
    e.R = (EndColor And &HFF)
    e.G = (EndColor \ &H100) And &HFF
    e.B = (EndColor And &HFF0000) / &H10000
    With i
        .R = (s.R - e.R) / lSteps
        .G = (s.G - e.G) / lSteps
        .B = (s.B - e.B) / lSteps
        If Direction Then  '-------- HORIZONTAL
            For x = 1 To lSteps
                Call LineApi(lHdc, (lSteps - x) + R.lLeft, R.lTop, (lSteps - x) + R.lLeft, R.lBottom, RGB(e.R + (x * .R), e.G + (x * .G), e.B + (x * .B)))
            Next x
        Else               '-------- VERTICAL
            For x = 1 To lSteps
                Call LineApi(lHdc, R.lLeft, (lSteps - x) + R.lTop, R.lRight, (lSteps - x) + R.lTop, RGB(e.R + (x * .R), e.G + (x * .G), e.B + (x * .B)))
            Next x
        End If
    End With
End Sub

Public Sub LineApi(ByVal lHdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long, ByVal Color As Long)
Dim pt      As POINTAPI
Dim hPen    As Long
Dim hPenOld As Long
    hPen = CreatePen(0, 1, Color)
    hPenOld = SelectObject(lHdc, hPen)
    MoveToEx lHdc, X1, Y1, pt
    LineTo lHdc, X2, Y2
    SelectObject lHdc, hPenOld
    DeleteObject hPen
End Sub

Public Sub ShowSysMenu(ByVal lHwnd As Long, ByVal x As Long, ByVal y As Long)
   '//NOTE: Must be in screen coordinates.
   Call SendMessage(lHwnd, WM_GETSYSMENU, 0, ByVal GetLong(y, x))
End Sub

Public Function IsAppRunning(ByVal sWindowCaptionElement As String, Optional ByVal sClassName As String = "") As Boolean
Dim lHandle As Long
    pAppFound = False
    pSearchTitle = UCase$(sWindowCaptionElement)
    pClassName = UCase$(sClassName)
    Call EnumWindows(AddressOf EnumWindowsCallBack, lHandle) '---- Enumerate windows to see if ours is running
    IsAppRunning = pAppFound '------------------------------------ Set function return
    pClassName = vbNull '----------------------------------------- Cleanup
    pSearchTitle = vbNull
End Function

Public Sub KillApp()
'--------------------------------------------------------------------------------------
' Purpose   : Kills all parent window processes associated with a running ProcessID
'--------------------------------------------------------------------------------------
On Error GoTo Errs
Dim x   As Integer
Dim ub  As Integer
    If Setup.RunMode = eNORMAL Then Screen.MousePointer = vbHourglass
    With tRunningApps
        ub = UBound(.lWndHwnd)
        For x = 0 To ub '---- Cycle through processes looking for our application to close.
            If .lWndProcessID(x) = lProcessToKill Then '---------- Look for ProcessID matches.
                If GetParent(.lWndHwnd(x)) = 0 Then '------------- Filter out non-parent windows.
                    SendMessage .lWndHwnd(x), WM_CLOSE, 0&, 0& '-- Send close message.
                End If
            End If
        Next x
    End With
    DoEvents: Sleep 1000 '--- Provide 1 Sec for processes to end and file locks to release.
                         '... This worked best with a variety of large and DB connected apps.
                         '... 1 Sec might be high, but it provided 100% success on all tests.
Errs:
    Screen.MousePointer = vbDefault
    If Err Then Err.Clear
End Sub

Public Sub CreateReport()
'---------------------------------------------------------------------------------------
' Purpose   : Generates and displays LiveUpdate HTML Report
'---------------------------------------------------------------------------------------
On Error Resume Next
Dim s               As String
Dim sFile           As String
Dim sFileList       As String
Dim sUpdatedList    As String
Dim sAvailList      As String
Dim sErrorList      As String
Dim sMissingList    As String
Dim sStatus()       As String
Dim f               As Integer
Dim x               As Integer
Dim iFileCount      As Integer
Dim iAvailCount     As Integer
Dim iUpdatedCount   As Integer
Dim iErrorCount     As Integer
Dim iMissingFiles   As Integer

sFileList = "&nbsp;&nbsp;-&nbsp;&nbsp;"
sUpdatedList = "&nbsp;&nbsp;-&nbsp;&nbsp;"
sAvailList = "&nbsp;&nbsp;-&nbsp;&nbsp;"
sErrorList = "&nbsp;&nbsp;-&nbsp;&nbsp;"
sMissingList = "&nbsp;&nbsp;-&nbsp;&nbsp;"

'//Do the math and create lists
With FileList
        ReDim sStatus(1 To UBound(.Description))
        For x = 1 To UBound(.Description)
        Select Case .Status(x)
            Case UPDATEREQ
                iAvailCount = iAvailCount + 1
                sAvailList = sAvailList & .FileName(x) & ", "
                sStatus(x) = "Update Available"
            Case NOUPDATEREQ
                sStatus(x) = "Installed Version Current - No Update Available"
            Case UPDATECOMP
                iUpdatedCount = iUpdatedCount + 1
                sUpdatedList = sUpdatedList & .FileName(x) & ", "
                sStatus(x) = "Success - Update Complete"
            Case UPDATECOMPREBOOT
                iUpdatedCount = iUpdatedCount + 1
                sUpdatedList = sUpdatedList & .FileName(x) & ", "
                sStatus(x) = "Success - Reboot Required to Complete Update"
            Case ERRCONNECTING
                iErrorCount = iErrorCount + 1
                sErrorList = sErrorList & .FileName(x) & ", "
                sStatus(x) = "Failed - Error Connecting to File"
            Case ERRTRANSFERRING
                iErrorCount = iErrorCount + 1
                sErrorList = sErrorList & .FileName(x) & ", "
                sStatus(x) = "Failed - Error During File Transfer"
            Case INSUFFPRIVILEGE
                iErrorCount = iErrorCount + 1
                sErrorList = sErrorList & .FileName(x) & ", "
                sStatus(x) = "Failed - Insufficient Privilege to Update"
            Case ERRUPDATING
                iErrorCount = iErrorCount + 1
                sErrorList = sErrorList & .FileName(x) & ", "
                sStatus(x) = "Failed - Reason Unknown"
            Case FILENOTINSTALLED
                iMissingFiles = iMissingFiles + 1
                sMissingList = sMissingList & .FileName(x) & ", "
                iErrorCount = iErrorCount + 1
                sErrorList = sErrorList & .FileName(x) & ", "
                sStatus(x) = "Failed - File Must Be Present on Client to Update"
            Case Else
                sStatus(x) = "Update Available - Pending"
        End Select
        sFileList = sFileList & .FileName(x) & ", "
    Next x
    sFileList = IIf(Len(sFileList) > 25, LCase$(Left$(sFileList, Len(sFileList) - 2)), "")
    sUpdatedList = IIf(Len(sUpdatedList) > 25, LCase$(Left$(sUpdatedList, Len(sUpdatedList) - 2)), "")
    sAvailList = IIf(Len(sAvailList) > 25, LCase$(Left$(sAvailList, Len(sAvailList) - 2)), "")
    sErrorList = IIf(Len(sErrorList) > 25, LCase$(Left$(sErrorList, Len(sErrorList) - 2)), "")
    sMissingList = IIf(Len(sMissingList) > 25, LCase$(Left$(sMissingList, Len(sMissingList) - 2)), "")
    iFileCount = x - 1
End With

'//Begin Creating Report
s = "<html>" & vbNewLine

'//Open head tag
s = s & "<head><META http-equiv=Content-Type content=text/html;charset=UTF-16>" & vbNewLine
    '//Create styles
    s = s & "       <style>BODY{FONT-SIZE:10pt;FONT-FAMILY:MS Sans Serif,Arial}" & vbNewLine
    s = s & "               .headers{FONT-SIZE: larger;COLOR: white;BACKGROUND-COLOR: #2C78A0}" & vbNewLine
    s = s & "               .tables{FONT-SIZE: 10pt;COLOR: black;BACKGROUND-COLOR: #EBF2FA}" & vbNewLine
    s = s & "       </style>" & vbNewLine
    '//////////////////////////////////////////////////////////////////////////////////////////
    '//Following lines removed to avoid script warning in browser for XP users.
    's = s & "       <script>" & vbNewLine
    's = s & "               window.status='Report Generated by ReVive LiveUpdate'" & vbNewLine
    's = s & "       </script>" & vbNewLine
    '//////////////////////////////////////////////////////////////////////////////////////////
'//Close head tag
s = s & "</head>" & vbNewLine
'//Assign browser title
s = s & "<title>" & Setup.AppShortName & " LiveUpdate Report</title>" & vbNewLine
'//Create overview table
s = s & "<BR>" & vbNewLine
s = s & "<table width=750 border=0 align=center cellpadding=0 cellspacing=0>" & vbNewLine
s = s & "<tr class=headers>" & vbNewLine
s = s & "       <td width=750 colspan=2 align=center height=30><b>" & Setup.AppShortName & " LiveUpdate Report</b></td>" & vbNewLine
s = s & "</tr>" & vbNewLine
s = s & "<tr class=tables>" & vbNewLine
s = s & "       <td width=210 height=24>&nbsp;&nbsp;Application Title:</td>" & vbNewLine
s = s & "       <td width=540 height=24>" & Setup.AppLongName & "</td>" & vbNewLine
s = s & "</tr>" & vbNewLine
s = s & "<tr class=tables>" & vbNewLine
s = s & "       <td width=210 height=24>&nbsp;&nbsp;Files Checked:</td>" & vbNewLine
s = s & "       <td width=540 height=24>" & iFileCount & sFileList & "</td>" & vbNewLine
s = s & "</tr>" & vbNewLine
s = s & "<tr class=tables>" & vbNewLine
s = s & "       <td width=210 height=24>&nbsp;&nbsp;Updates Available:</td>" & vbNewLine
s = s & "       <td width=540 height=24>" & iAvailCount & sAvailList & "</td>" & vbNewLine
s = s & "</tr>" & vbNewLine
s = s & "<tr class=tables>" & vbNewLine
s = s & "       <td width=210 height=24>&nbsp;&nbsp;Updates Completed:</td>" & vbNewLine
s = s & "       <td width=540 height=24>" & iUpdatedCount & sUpdatedList & "</td>" & vbNewLine
s = s & "</tr>" & vbNewLine
s = s & "<tr class=tables>" & vbNewLine
s = s & "       <td width=210 height=24>&nbsp;&nbsp;Required Files Missing:</td>" & vbNewLine
s = s & "       <td width=540 height=24>" & iMissingFiles & sMissingList & "</td>" & vbNewLine
s = s & "</tr>" & vbNewLine
s = s & "<tr class=tables>" & vbNewLine
s = s & "       <td width=210 height=24>&nbsp;&nbsp;File Update Errors:</td>" & vbNewLine
s = s & "       <td width=540 height=24>" & iErrorCount & sErrorList & "</td>" & vbNewLine
s = s & "</tr>" & vbNewLine
s = s & "<tr class=tables>" & vbNewLine
s = s & "       <td width=210 height=24>&nbsp;&nbsp;Requires Admin to Update:</td>" & vbNewLine
s = s & "       <td width=540 height=24>" & Setup.AdminRequired & "</td>" & vbNewLine
s = s & "</tr>" & vbNewLine
s = s & "<tr class=tables>" & vbNewLine
s = s & "       <td width=210 height=24>&nbsp;&nbsp;Date Created:</td>" & vbNewLine
s = s & "       <td width=540 height=24>" & Format(Now, "dd mmm yy - h:mm AM/PM") & "</td>" & vbNewLine
s = s & "</tr>" & vbNewLine
s = s & "<tr class=headers>" & vbNewLine
s = s & "       <td width=750 colspan=2 height=2></td>" & vbNewLine
s = s & "</tr>" & vbNewLine
s = s & "<tr>" & vbNewLine
s = s & "       <td width=750 colspan=2 height=20></td>" & vbNewLine
s = s & "</tr>" & vbNewLine
'//Display file information
With FileList
    For x = 1 To UBound(.Description)
        s = s & "<tr class=headers>" & vbNewLine
        s = s & "       <td width=210 align=left height=24><font size=3><b>&nbsp;File " & x & "</b></font></td>" & vbNewLine
        s = s & "       <td width=540 align=right height=24><font size=3><b>" & .Description(x) & "&nbsp;</b></font></td>" & vbNewLine
        s = s & "</tr>" & vbNewLine
        s = s & "<tr class=tables>" & vbNewLine
        s = s & "       <td width=210 height=24>&nbsp;&nbsp;Install Path:</td>" & vbNewLine
        s = s & "       <td width=540 height=24>" & LCase$(.InstallPath(x)) & "</td>" & vbNewLine
        s = s & "</tr>" & vbNewLine
        s = s & "<tr class=tables>" & vbNewLine
        s = s & "       <td width=210 height=24>&nbsp;&nbsp;Installed Version:</td>" & vbNewLine
        s = s & "       <td width=540 height=24>" & IIf(.CurrentVersion(x) = "0" Or .CurrentVersion(x) = "0.0.0.0", "Not Installed", .CurrentVersion(x)) & "</td>" & vbNewLine
        s = s & "</tr>" & vbNewLine
        s = s & "<tr class=tables>" & vbNewLine
        s = s & "       <td width=210 height=24>&nbsp;&nbsp;Updated Version:</td>" & vbNewLine
        s = s & "       <td width=540 height=24>" & .UpdateVersion(x) & "</td>" & vbNewLine
        s = s & "</tr>" & vbNewLine
        s = s & "<tr class=tables>" & vbNewLine
        s = s & "       <td width=210 height=24>&nbsp;&nbsp;File Size:</td>" & vbNewLine
        s = s & "       <td width=540 height=24>" & Format(.FileSize(x), "0,000 Bytes") & "</td>" & vbNewLine
        s = s & "</tr>" & vbNewLine
        s = s & "<tr class=tables>" & vbNewLine
        s = s & "       <td width=210 height=24>&nbsp;&nbsp;Must Succeed to Update App:</td>" & vbNewLine
        s = s & "       <td width=540 height=24>" & .MustUpdate(x) & "</td>" & vbNewLine
        s = s & "</tr>" & vbNewLine
        s = s & "<tr class=tables>" & vbNewLine
        s = s & "       <td width=210 height=24>&nbsp;&nbsp;Must be Installed to Update:</td>" & vbNewLine
        s = s & "       <td width=540 height=24>" & .MustExist(x) & "</td>" & vbNewLine
        s = s & "</tr>" & vbNewLine
        s = s & "<tr class=tables>" & vbNewLine
        s = s & "       <td width=210 height=24>&nbsp;&nbsp;Status of Update:</td>" & vbNewLine
        s = s & "       <td width=540 height=24>" & sStatus(x) & "</td>" & vbNewLine
        s = s & "</tr>" & vbNewLine
        s = IIf(x = UBound(.Description), s & "<tr class=headers>", s & "<tr class=tables>") & vbNewLine
        s = s & "       <td width=750 colspan=2 height=4></td>" & vbNewLine
        s = s & "</tr>" & vbNewLine
    Next x
End With
s = s & "<tr>" & vbNewLine
s = s & "       <td width=750 colspan=2 height=20></td>" & vbNewLine
s = s & "</tr>" & vbNewLine
s = s & "</table>" & vbNewLine
s = s & "</html>" & vbNewLine
'//Save and open report
sFile = Environ("TEMP") & "\UpdateReport.htm"
f = FreeFile
Open sFile For Output As #f
    Print #f, s
Close #f
ShellExecute 0&, "open", sFile, vbNullString, vbNullString, SW_MAXIMIZE

End Sub


'****************************************************************************************
'                               Private Helper Functions
'****************************************************************************************

Private Function EnumWindowsCallBack(ByVal lHandle As Long, ByVal lpData As Long) As Long
'-----------------------------------------------------------------------------------------
' Author    : Chris Cochran (Using sample provided by Dave Scarmozzino, www.thescarms.com)
' Purpose   : Call back function from IsAppRunning. Enumerates all parent windows and
'             searches for a specific window title and optional classname.
'-----------------------------------------------------------------------------------------
On Error GoTo Errs
Dim lResult         As Long
Dim lThreadId       As Long
Dim lProcessId      As Long
Dim sWindowTitle    As String
Dim sClassName      As String
Static lCount       As Integer
    EnumWindowsCallBack = 1
    lThreadId = GetWindowThreadProcessId(lHandle, lProcessId)
    If lThreadId = App.ThreadID Then Exit Function '---------------- Skip if ReVive ThreadID
    If Setup.UpdateAppKill Then '----------------------------------- Skip if we are not killing the app
        With tRunningApps
            ReDim Preserve .lWndHwnd(0 To lCount)
            ReDim Preserve .lWndProcessID(0 To lCount)
            .lWndHwnd(lCount) = lHandle
            .lWndProcessID(lCount) = lProcessId
        End With
        lCount = lCount + 1
    End If
    If Not pAppFound Then '----------------------------------------- Skip below code once app is found
        sWindowTitle = Space$(MAX_PATH)
        lResult = GetWindowText(lHandle, sWindowTitle, MAX_PATH) '-- Get window title
        sWindowTitle = UCase$(Left$(sWindowTitle, lResult))
        sClassName = Space$(MAX_PATH)
        lResult = GetClassName(lHandle, sClassName, MAX_PATH) '----- Get window classname
        sClassName = UCase$(Left$(sClassName, lResult))
        If InStr(1, sWindowTitle, pSearchTitle) Then '-------------- Search for our title
            If Len(pClassName) Then
                If sClassName = pClassName Then '------------------- Check for matching classname if requested
                    pAppFound = True
                    lProcessToKill = lProcessId
                End If
            Else
                pAppFound = True
                lProcessToKill = lProcessId
            End If
        End If
    End If
Errs:
    If Err Then Err.Clear
End Function

Private Function GetLong(ByVal WordHi As Integer, ByVal WordLo As Integer) As Long
    GetLong = (CLng(WordHi) * &H10000) Or (WordLo And &HFFFF&)
End Function

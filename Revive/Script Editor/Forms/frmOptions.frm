VERSION 5.00
Begin VB.Form frmOptions 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Script Editor Options"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   6060
   Icon            =   "frmOptions.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   6060
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame4 
      Caption         =   "Automation Options"
      Height          =   1785
      Left            =   180
      TabIndex        =   12
      Top             =   2640
      Width           =   5685
      Begin VB.ComboBox cmbDefConst 
         Height          =   315
         ItemData        =   "frmOptions.frx":058A
         Left            =   210
         List            =   "frmOptions.frx":05AC
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   1290
         Width           =   2475
      End
      Begin VB.TextBox txtDefWeb 
         Height          =   315
         Left            =   210
         TabIndex        =   3
         Top             =   600
         Width           =   5235
      End
      Begin VB.Label lblConst 
         Caption         =   "ScriptFile Path"
         Height          =   225
         Left            =   2850
         TabIndex        =   15
         Top             =   1350
         Width           =   2625
      End
      Begin VB.Label Label2 
         Caption         =   "Default Folder Constant:"
         Height          =   195
         Left            =   210
         TabIndex        =   14
         Top             =   1050
         Width           =   2535
      End
      Begin VB.Label Label1 
         Caption         =   "Default Address:"
         Height          =   195
         Left            =   210
         TabIndex        =   13
         Top             =   360
         Width           =   1635
      End
   End
   Begin VB.CommandButton cmd 
      Cancel          =   -1  'True
      Caption         =   "&Cancel"
      Height          =   375
      Index           =   1
      Left            =   3840
      TabIndex        =   6
      Top             =   4650
      Width           =   975
   End
   Begin VB.CommandButton cmd 
      Caption         =   "&OK"
      Height          =   375
      Index           =   0
      Left            =   4890
      TabIndex        =   5
      Top             =   4650
      Width           =   975
   End
   Begin VB.Frame Frame2 
      Caption         =   "Script Editor Colors"
      Height          =   975
      Left            =   180
      TabIndex        =   8
      Top             =   1500
      Width           =   5685
      Begin VB.CommandButton cmdDef 
         Caption         =   "&Defaults"
         Height          =   375
         Left            =   4590
         TabIndex        =   2
         TabStop         =   0   'False
         Top             =   360
         Width           =   975
      End
      Begin VB.Label lblTag 
         Caption         =   "Value Tags"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   2
         Left            =   3600
         TabIndex        =   11
         Top             =   480
         Width           =   855
      End
      Begin VB.Shape shpTag 
         FillColor       =   &H00D11FE0&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   2
         Left            =   3330
         Top             =   450
         Width           =   225
      End
      Begin VB.Label lblTag 
         Caption         =   "Key Tags"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   1
         Left            =   2460
         TabIndex        =   10
         Top             =   480
         Width           =   705
      End
      Begin VB.Shape shpTag 
         FillColor       =   &H00FF0000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   1
         Left            =   2190
         Top             =   450
         Width           =   225
      End
      Begin VB.Label lblTag 
         Caption         =   "Section Header Tags"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   -1  'True
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Index           =   0
         Left            =   510
         TabIndex        =   9
         Top             =   480
         Width           =   1515
      End
      Begin VB.Shape shpTag 
         FillColor       =   &H00800000&
         FillStyle       =   0  'Solid
         Height          =   225
         Index           =   0
         Left            =   240
         Top             =   450
         Width           =   225
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Miscellaneous Options"
      Height          =   1125
      Left            =   180
      TabIndex        =   7
      Top             =   210
      Width           =   5685
      Begin VB.CheckBox chkAssociate 
         Caption         =   "Associate .rus files with the ReVive LiveUpdate Script Editor"
         Height          =   195
         Left            =   210
         TabIndex        =   1
         Top             =   750
         Width           =   4575
      End
      Begin VB.CheckBox chkFileNames 
         Caption         =   "Test for Win9X - ME filename compatibility. (DOS 8.3 format filenames)"
         Height          =   195
         Left            =   210
         TabIndex        =   0
         Top             =   390
         Width           =   5295
      End
   End
End
Attribute VB_Name = "frmOptions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//START IsAdministrator Declares
Private Const TOKEN_READ                    As Long = &H20008
Private Const SECURITY_BUILTIN_DOMAIN_RID   As Long = &H20&
Private Const DOMAIN_ALIAS_RID_ADMINS       As Long = &H220&
Private Const SECURITY_NT_AUTHORITY         As Long = &H5
Private Const TokenGroups                   As Long = 2
Private Type SID_IDENTIFIER_AUTHORITY
   Value(6) As Byte
End Type
Private Type SID_AND_ATTRIBUTES
   Sid As Long
   Attributes As Long
End Type
Private Type TOKEN_GROUPS
   GroupCount As Long
   Groups(500) As SID_AND_ATTRIBUTES
End Type
Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal Sid As Long, ByVal name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Long) As Long
Private Declare Function AllocateAndInitializeSid Lib "advapi32.dll" (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Sub FreeSid Lib "advapi32.dll" (pSid As Any)
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

'//START WindowsVersion Declarations
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long       '1 = Windows 95/98.    '2 = Windows NT
    szCSDVersion As String * 128
End Type

'//Misc and shared declares
Private Declare Function SendMessageTimeout Lib "user32" Alias "SendMessageTimeoutA" (ByVal hwnd As Long, ByVal msg As Long, ByVal wParam As Long, ByVal lParam As Any, ByVal fuFlags As Long, ByVal uTimeout As Long, lpdwResult As Long) As Long
Private Declare Function SendMessage Lib "user32" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, lParam As Any) As Long
Private Const EM_GETLINECOUNT   As Long = &HBA

Private bAssocValue             As Long

Private Sub cmbDefConst_Click()
Dim s As String
    Select Case Me.cmbDefConst.Text
        Case "<sp>"
            s = "Clients Script Path"
        Case "<ap>"
            s = "Clients Application Path"
        Case "<win>"
            s = "Clients Windows Directory"
        Case "<sys>"
            s = "Clients System Directory"
        Case "<temp>"
            s = "Clients Temporary Folder"
        Case "<pf>"
            s = "Clients Program Files Folder"
        Case "<cf>"
            s = "Clients Common Files Folder"
        Case "<userdesktop>"
            s = "Clients Desktop Folder"
        Case "<commondesktop>"
            s = "Clients Common Desktop Folder"
        Case "<commonstartmenu>"
            s = "Clients Common StartMenu Folder"
    End Select
    Me.lblConst.Caption = s
End Sub

Private Sub cmd_Click(Index As Integer)
Dim x               As Long
Dim lLineCount      As Long
Dim bWasItChanged   As Boolean
Dim s               As String
    If Index = 0 Then
        With Me.txtDefWeb '-------------------------- Verify default address ends with a / or a \
            If Len(.Text) Then
                s = Right$(.Text, 1)
                If s <> "/" And s <> "\" Then
                    MsgBox "Please end the Default Address with either a \ or a /. If the address     " & vbNewLine & _
                    "is a local or network path end if with a \. If it's a URL end it with a /.    ", vbExclamation, "Script Editor"
                    .SelStart = Len(.Text)
                    .SetFocus
                    Exit Sub
                End If
            End If
        End With
        With Setup
            .TestForFileNames = Abs(Me.chkFileNames.Value)
            .DefaultWeb = Me.txtDefWeb.Text
            .DefaultConst = Me.cmbDefConst.Text
            If Abs(Me.chkAssociate.Value) <> bAssocValue Then
                Call CreateAssociation(Abs(Me.chkAssociate.Value))
            End If
            If .SecTagColor <> Me.shpTag(0).FillColor Or _
                    .KeyTagColor <> Me.shpTag(1).FillColor Or _
                    .ValTagColor <> Me.shpTag(2).FillColor Then
                bWasItChanged = bChanged
                lLineCount = SendMessage(frmMain.rtBox.hwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
                .SecTagColor = Me.shpTag(0).FillColor
                .KeyTagColor = Me.shpTag(1).FillColor
                .ValTagColor = Me.shpTag(2).FillColor
                For x = -lLineCount To -1
                    frmMain.ColorLine Abs(x)
                Next x
                bChanged = bWasItChanged
            End If
            SaveSetting "ReVive Script Editor", "Config", "TestFileNames", .TestForFileNames
            SaveSetting "ReVive Script Editor", "Config", "DefaultWeb", .DefaultWeb
            SaveSetting "ReVive Script Editor", "Config", "DefaultConst", .DefaultConst
            SaveSetting "ReVive Script Editor", "Config", "SecTagColor", .SecTagColor
            SaveSetting "ReVive Script Editor", "Config", "KeyTagColor", .KeyTagColor
            SaveSetting "ReVive Script Editor", "Config", "ValTagColor", .ValTagColor
        End With
    End If
    Unload Me
End Sub

Private Sub cmdDef_Click()
    Me.shpTag(0).FillColor = &H800000
    Me.shpTag(1).FillColor = &HFF0000
    Me.shpTag(2).FillColor = 13705184
End Sub

Private Sub Form_Load()
Dim cR As New cReg
    Me.chkAssociate.Enabled = IsAdministrator
    With Setup '-------------------------------- Display current settings
        Me.chkFileNames.Value = Abs(.TestForFileNames)
        Me.shpTag(0).FillColor = .SecTagColor
        Me.shpTag(1).FillColor = .KeyTagColor
        Me.shpTag(2).FillColor = .ValTagColor
        Me.txtDefWeb.Text = .DefaultWeb
        Me.cmbDefConst.Text = .DefaultConst
    End With
    With cR '----------------------------------- Get and display association settings
        .ClassKey = HKEY_CLASSES_ROOT
        .ValueType = REG_SZ
        .SectionKey = ".rus"
        bAssocValue = IIf(Len(.Value), 1, 0)
    End With
    Me.chkAssociate.Value = bAssocValue
End Sub

Private Function WindowsVersion() As Long
'--------------------------------------------------------------
' Purpose   : Returns 1 if 95/98/ME and 2 or > for NT based OS.
'--------------------------------------------------------------
Dim osinfo   As OSVERSIONINFO
Dim retvalue As Integer
    osinfo.dwOSVersionInfoSize = 148
    osinfo.szCSDVersion = Space$(128)
    retvalue = GetVersionExA(osinfo)
    WindowsVersion = osinfo.dwPlatformId
End Function

Private Function IsAdministrator() As Boolean
'*******************************************************************************************
'Adapted from code written by Randy Birch, http://vbnet.mvps.org
'"How to Determine if the Current User is a Member of Administrators"
'Full post is here: http://vbnet.mvps.org/index.html?code/network/isadministrator.htm
'*******************************************************************************************

 Dim hProcessID     As Long
 Dim hToken         As Long
 Dim res            As Long
 Dim cbBuff         As Long
 Dim tiLen          As Long
 Dim TG             As TOKEN_GROUPS
 Dim SIA            As SID_IDENTIFIER_AUTHORITY
 Dim lSid           As Long
 Dim cnt            As Long
 Dim sAcctName1     As String
 Dim sAcctName2     As String
 Dim cbAcctName     As Long
 Dim sDomainName    As String
 Dim cbDomainName   As Long
 Dim peUse          As Long
 
 '//See if Win9X or ME.
 '..This is the only thing I changed or added in this routine.
 If WindowsVersion = 1 Then
    IsAdministrator = True
    Exit Function
 End If
 
 tiLen = 0
 
'obtain handle to process. 0 indicates failure;
'may return -1 for current process (and is valid)
 hProcessID = GetCurrentProcess()
 
 If hProcessID <> 0 Then

   'obtain a handle to the access
   'token associated with the process
    If OpenProcessToken(hProcessID, TOKEN_READ, hToken) = 1 Then

      'retrieve specified information
      'about an access token. The first
      'call to GetTokenInformation fails
      'since the buffer size is unspecified.
      'On failure the correct buffer size
      'is returned (cbBuff), and a subsequent call
      'is made to return the data.
       res = GetTokenInformation(hToken, _
                                 TokenGroups, _
                                 TG, _
                                 tiLen, _
                                 cbBuff)
       
       If res = 0 And cbBuff > 0 Then

          tiLen = cbBuff
          res = GetTokenInformation(hToken, _
                                    TokenGroups, _
                                    TG, _
                                    tiLen, _
                                    cbBuff)
          
          If res = 1 And tiLen > 0 Then
          
            'The SID_IDENTIFIER_AUTHORITY (SIA) structure
            'represents the top-level authority of a
            'security identifier (SID). By specifying
            'we want admins (by setting the value of
            'the fifth item to SECURITY_NT_AUTHORITY),
            'and passing the relative identifiers (RID)
            'DOMAIN_ALIAS_RID_ADMINS  and
            'SECURITY_BUILTIN_DOMAIN_RID, we obtain
            'the SID for the administrators account
            'in lSid
             SIA.Value(5) = SECURITY_NT_AUTHORITY
 
             res = AllocateAndInitializeSid(SIA, 2, _
                                            SECURITY_BUILTIN_DOMAIN_RID, _
                                            DOMAIN_ALIAS_RID_ADMINS, _
                                            0, 0, 0, 0, 0, 0, _
                                            lSid)
                                            
             If res = 1 Then
             
               'Now obtain the name of the account
               'pointed to by lSid above (ie
               '"Administrators"). Note vbNullString
               'is passed as lpSystemName indicating
               'the SID is looked up on the local computer.
               '
               'Re sDomainName: On Win NT+ systems, the
               'domain name returned for most accounts in
               'the local computer's security database is
               'the computer's name as of the last start
               'of the system (backslashes excluded). If
               'the computer's name changes, the old name
               'continues to be returned as the domain
               'name until the system is restarted.
               '
               'On Win NT+ Server systems, the domain name
               'returned for most accounts in the local
               'computer's security database is the
               'name of the domain for which the server is
               'a domain controller.
               '
               'Some accounts are predefined by the system.
               'The domain name returned for these accounts
               'is BUILTIN.
               '
               'sAcctName is the value of interest in this
               'exercise.
                sAcctName1 = Space$(255)
                sDomainName = Space$(255)
                cbAcctName = 255
                cbDomainName = 255
                res = LookupAccountSid(vbNullString, _
                                       lSid, _
                                       sAcctName1, _
                                       cbAcctName, _
                                       sDomainName, _
                                       cbDomainName, _
                                       peUse)
                                       
                If res = 1 Then
                   
                  'In the call to GetTokenInformation above,
                  'the TOKEN_GROUP member was filled with
                  'the SIDs of the defined groups.
                  '
                  'Here we take each SID from the token
                  'group and retrieve the name of the account
                  'corresponding to the SID. If a SID returns
                  'the same name retrieved above, the user
                  'is a member of the admin group.
                   For cnt = 0 To TG.GroupCount - 1
                   
                      sAcctName2 = Space$(255)
                      sDomainName = Space$(255)
                      cbAcctName = 255
                      cbDomainName = 255
                      
                      res = LookupAccountSid(vbNullString, _
                                             TG.Groups(cnt).Sid, _
                                             sAcctName2, _
                                             cbAcctName, _
                                             sDomainName, _
                                             cbDomainName, _
                                             peUse)
                       
                      If sAcctName1 = sAcctName2 Then
                         IsAdministrator = True
                         Exit For
                      End If   'if sAcctName1 = sAcctName2
                      
                   Next
                   
                End If  'if res = 1 (LookupAccountSid)

                FreeSid ByVal lSid
       
             End If  'if res = 1 (AllocateAndInitializeSid)
             
             CloseHandle hToken
       
          End If  'if res = 1
       
       End If  'if res = 0  (GetTokenInformation)
    
    End If  'if OpenProcessToken
    
    CloseHandle hProcessID

 End If  'if hProcessID  (GetCurrentProcess)
 
End Function

Private Sub CreateAssociation(ByVal bAdd As Byte)
'---------------------------------------------------------------------------------------
' Purpose   : Creates file extension associations and refreshes system icons
'---------------------------------------------------------------------------------------
On Error Resume Next
Dim cR As cReg
    Screen.MousePointer = 11
    Set cR = New cReg
    With cR
        .ClassKey = HKEY_CLASSES_ROOT
        .ValueType = REG_SZ
        If bAdd Then
            .SectionKey = ".rus"
            .Value = "ReViveLiveUpdateScript"
            .CreateKey
            .SectionKey = "ReViveLiveUpdateScript"
            .Value = "ReVive LiveUpdate Script File"
            .CreateKey
            .SectionKey = "ReViveLiveUpdateScript\DefaultIcon"
            .Value = App.path & "\" & App.EXEName & ".exe,1"
            .CreateKey
            .SectionKey = "ReViveLiveUpdateScript\shell\open\command"
            .Value = Chr(34) & App.path & "\" & App.EXEName & ".exe" & Chr(34) & " %1"
            .CreateKey
        Else
            .SectionKey = ".rus"
            .DeleteKey
            .SectionKey = "ReViveLiveUpdateScript\DefaultIcon"
            .DeleteKey
            .SectionKey = "ReViveLiveUpdateScript\shell\open\command"
            .DeleteKey
            .SectionKey = "ReViveLiveUpdateScript\shell\open"
            .DeleteKey
            .SectionKey = "ReViveLiveUpdateScript\shell"
            .DeleteKey
            .SectionKey = "ReViveLiveUpdateScript"
            .DeleteKey
        End If
    End With
    Call RefreshSystemIcons
    Set cR = Nothing
    Screen.MousePointer = 0
End Sub

Private Sub lblTag_Click(Index As Integer)
On Error GoTo Errs
    With frmMain.cd
        .ShowColor
        Me.shpTag(Index).FillColor = .Color
    End With
Errs:
    If Err Then Exit Sub
End Sub

Private Sub lblTag_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    SetCursor LoadCursor(0, IDC_HAND)
End Sub

Private Sub RefreshSystemIcons()
On Error Resume Next
Dim WSH                         As Object
Dim CurIconSize                 As String
Dim cR                          As New cReg
Dim Result                      As Long
Const HWND_BROADCAST            As Long = &HFFFF&
Const WM_SETTINGCHANGE          As Long = &H1A
Const SPI_SETNONCLIENTMETRICS   As Long = 42
Const SMTO_ABORTIFHUNG          As Long = &H2
Const REG_ICONSIZE_KEY          As String = "HKCU\Control Panel\Desktop\WindowMetrics\Shell Icon Size"
    
    '//Get current icon size
    With cR
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Control Panel\Desktop\WindowMetrics"
        .ValueKey = "Shell Icon Size"
        CurIconSize = .Value
    End With
    '//If no default size, assume 32.
    If CurIconSize = "" Then CurIconSize = 32
    '//Change the icon size to 1 pixel less.
    With cR
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Control Panel\Desktop\WindowMetrics"
        .ValueKey = "Shell Icon Size"
        .ValueType = REG_SZ
        .Value = CurIconSize - 1
    End With
    '//Broadcast change to all running apps
    Call SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, SPI_SETNONCLIENTMETRICS, 0&, SMTO_ABORTIFHUNG, 10000&, Result)
    '//Restore the original icon size.
    With cR
        .ClassKey = HKEY_CURRENT_USER
        .SectionKey = "Control Panel\Desktop\WindowMetrics"
        .ValueKey = "Shell Icon Size"
        .ValueType = REG_SZ
        .Value = CurIconSize
    End With
    '//Broadcast change to all running apps
    Call SendMessageTimeout(HWND_BROADCAST, WM_SETTINGCHANGE, SPI_SETNONCLIENTMETRICS, 0&, SMTO_ABORTIFHUNG, 10000&, Result)
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmOptions = Nothing
End Sub

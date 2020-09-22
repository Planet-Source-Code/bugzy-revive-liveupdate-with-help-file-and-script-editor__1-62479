Attribute VB_Name = "modUpdate"
'***********************************************************************
'
'Chris Cochran          cwc.software@gmail.com        Updated: 18 Aug 05
'
'THIS MODULE IS WRITTEN FOR THE SOLE PURPOSE OF SYSTEMATICALLY REPLACING
'OR INSTALLING FILES FOR PERFORMING APPLICATION UPDATES OR INSTALLS. ALL
'OF THE BELOW ROUTINES ARE WINDOWS 95 THROUGH XP COMPATIBLE.
'
'MODULE MUST ACCOMPANY cREG CLASS (Written by Steve McMahon). ALL OTHER
'FUNCTIONS IN THIS MODULE ARE INDEPENDENT OF ALL OTHER CODE IN THIS
'PROJECT. THIS WAS DONE FOR REUSABILITY. THE GetFileExt, WindowsVersion,
'IsAdministrator, IsLocalPathValid, and WindowsVersion FUNCTIONS ARE
'MADE PUBLIC HERE TO ELIMINATE SOME REDUNDANCY IN THIS PROJECT.
'***********************************************************************

Option Explicit

'//Returns from Update routine
Public Enum eupdResults
    eupdSUCCESSCOMP = 0 '-------- Success/Complete
    eupdSUCCESSREBOOT = 1 '------ Success/Reboot required
    eupdSOURCENOTFOUND = 2 '----- Source file not found
    eupdDESTINVALID = 3 '-------- Destination path invalid
    eupdINSUFFPRIV = 4 '--------- Insufficient privilege
    eupdUNKNOWNERR = 5 '--------- Unknown update error
End Enum

'//Establishes Windows directory. A must have for Win9X/ME when manipulating wininit.ini
Private Declare Function GetWindowsDirectory Lib "kernel32" Alias "GetWindowsDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long

'//WindowsVersion Declarations
Private Declare Function GetVersionExA Lib "kernel32" (lpVersionInformation As OSVERSIONINFO) As Integer
Private Type OSVERSIONINFO
    dwOSVersionInfoSize                     As Long
    dwMajorVersion                          As Long
    dwMinorVersion                          As Long
    dwBuildNumber                           As Long
    dwPlatformId                            As Long '1 = Windows 95/98. '2 = Windows NT and Up
    szCSDVersion                            As String * 128
End Type

'//DECLARES For FileInUse
Private Const OFS_MAXPATHNAME               As Long = 128
Private Const OF_SHARE_EXCLUSIVE            As Long = &H10
Private Type OFSTRUCT
    cBytes                                  As Byte
    fFixedDisk                              As Byte
    nErrCode                                As Integer
    Reserved1                               As Integer
    Reserved2                               As Integer
    szPathName(OFS_MAXPATHNAME)             As Byte
End Type
Private Declare Function OpenFile Lib "kernel32" (ByVal lpFileName As String, ByRef lpReOpenBuff As OFSTRUCT, ByVal uStyle As Long) As Long
Private Declare Function CloseHandle Lib "kernel32.dll" (ByVal hObject As Long) As Long

'//MoveFile/DeleteFile Declares
Private Const MOVEFILE_REPLACE_EXISTING     As Long = &H1
Private Const MOVEFILE_DELAY_UNTIL_REBOOT   As Long = &H4
Private Const MOVEFILE_COPY_ALLOWED         As Long = &H2
Private Declare Function MoveFileEx Lib "kernel32" Alias "MoveFileExA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String, ByVal dwFlags As Long) As Long
Private Declare Function MoveFile Lib "kernel32" Alias "MoveFileA" (ByVal lpExistingFileName As String, ByVal lpNewFileName As String) As Long
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long


'//IsAdministrator Declares
Private Const TOKEN_READ                    As Long = &H20008
Private Const SECURITY_BUILTIN_DOMAIN_RID   As Long = &H20&
Private Const DOMAIN_ALIAS_RID_ADMINS       As Long = &H220&
Private Const SECURITY_NT_AUTHORITY         As Long = &H5
Private Const TokenGroups                   As Long = 2
Private Type SID_IDENTIFIER_AUTHORITY
    Value(6)                                As Byte
End Type
Private Type SID_AND_ATTRIBUTES
    Sid                                     As Long
    Attributes                              As Long
End Type
Private Type TOKEN_GROUPS
    GroupCount                              As Long
    Groups(500)                             As SID_AND_ATTRIBUTES
End Type
Private Declare Function LookupAccountSid Lib "advapi32.dll" Alias "LookupAccountSidA" (ByVal lpSystemName As String, ByVal Sid As Long, ByVal name As String, cbName As Long, ByVal ReferencedDomainName As String, cbReferencedDomainName As Long, peUse As Long) As Long
Private Declare Function AllocateAndInitializeSid Lib "advapi32.dll" (pIdentifierAuthority As SID_IDENTIFIER_AUTHORITY, ByVal nSubAuthorityCount As Byte, ByVal nSubAuthority0 As Long, ByVal nSubAuthority1 As Long, ByVal nSubAuthority2 As Long, ByVal nSubAuthority3 As Long, ByVal nSubAuthority4 As Long, ByVal nSubAuthority5 As Long, ByVal nSubAuthority6 As Long, ByVal nSubAuthority7 As Long, lpPSid As Long) As Long
Private Declare Function OpenProcessToken Lib "advapi32.dll" (ByVal ProcessHandle As Long, ByVal DesiredAccess As Long, TokenHandle As Long) As Long
Private Declare Function GetTokenInformation Lib "advapi32.dll" (ByVal TokenHandle As Long, ByVal TokenInformationClass As Long, TokenInformation As Any, ByVal TokenInformationLength As Long, ReturnLength As Long) As Long
Private Declare Function GetCurrentProcess Lib "kernel32" () As Long
Private Declare Sub FreeSid Lib "advapi32.dll" (pSid As Any)

'//For Win 9X/ME (wininit.ini only processes short file name paths)
Private Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

Public Function TestUpdateSuccess(ByVal sSourceFile As String, _
    ByVal sDestinationFile As String) As eupdResults
'--------------------------------------------------------------------------------------
' Purpose   : This routine evaluates if the update operation can complete without error
'             prior to attempting UpdateFile operation. This is used for validating a
'             group of files will succeed before committing an update to any one file.
'
' Checks For:
'       - A valid source file
'       - Valid destination path meeting Windows naming conventions
'       - The ability to create destination path
'       - Admin privilege when in-use files must be processed after reboot
'       - Admin privilege for DLL/OCX files that require regsvr32
'       - Write permission to the destination path
'
' VerifySuccess Returns:
'       0 = Success/Complete
'       2 = sSourceFile not found
'       3 = Destination invalid/Destination write error
'       4 = Insufficient privilege
'--------------------------------------------------------------------------------------
On Error GoTo Errs

Dim INUSE   As Boolean  '--- File In-Use status
Dim FILEEXT As String   '--- File extension of sDestinationFile
Dim path    As String   '--- Update files destination path
    INUSE = FileInUse(sDestinationFile)
    FILEEXT = GetFileExt(sDestinationFile)
    path = Left$(sDestinationFile, InStrRev(sDestinationFile, "\", , vbTextCompare))
    'Ensure source file is downloaded and available
    If Dir$(sSourceFile, vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem) = "" Then
        TestUpdateSuccess = eupdSOURCENOTFOUND
    'See if file is in use and if user has the rights to update those that are
    ElseIf INUSE And Not bADMIN Then
        TestUpdateSuccess = eupdINSUFFPRIV
    'See if user can register OCX and DLL files when required
    ElseIf (FILEEXT = "OCX" Or FILEEXT = "DLL") And Not bADMIN Then
        TestUpdateSuccess = eupdINSUFFPRIV
    'Verify install path meets Windows naming conventions and can be created
    ElseIf Not CreatePath(path) Then
        TestUpdateSuccess = eupdDESTINVALID
    'Verify user has write access in install path
    ElseIf Not CanWriteToPath(path) Then
        TestUpdateSuccess = eupdINSUFFPRIV
    Else
        TestUpdateSuccess = eupdSUCCESSCOMP
    End If
Errs_Exit:
    Exit Function
Errs:
    TestUpdateSuccess = eupdUNKNOWNERR
    Resume Errs_Exit
End Function

Public Function UpdateFile(ByVal sSourceFile As String, ByVal sDestinationFile As String) As eupdResults
'--------------------------------------------------------------------------------------
' Purpose   : This routine evaluates and executes a strategy for updating or installing
'             the passed sDestinationFile with the passed sSourceFile.
'
' Returns   : 0 = Success/Complete
'             1 = Success/Reboot Required
'             2 = sSourceFile not found
'             3 = Destination invalid/Destination write error
'             4 = Insufficient privilege
'             5 = Unknown Error
'
' IMPORTANT : 1. sDestinationFile names should be in DOS 8.3 format for Windows 95/98
'                and ME compatibility. (See Included README.rtf)
'             2. UNC paths for sDestinationFile are not supported. (i.e. "\\LANComp")
'
' REFERENCE : "How To Move Files That Are Currently in Use", Microsoft
'             http://support.microsoft.com/default.aspx?scid=kb;EN-US;140570
'--------------------------------------------------------------------------------------
On Error GoTo Errs
Dim INUSE       As Boolean      'Preliminary file In-Use status
Dim REGREQ      As Boolean      'File will require registerin with regsvr32
Dim EXT         As String       'Destination files extension
Dim lResult     As Long
Dim c           As New cReg

'--- WHERE WE ARE NOW? SO FAR WE HAVE DONE THE FOLLOWING (FROM SUB TESTUPDATESUCCESS):
'       - Ensured source file is valid
'       - Validated destination path met Windows naming conventions
'       - Successfully created the destination path
'       - Verified Admin privilege for in-use files that must be processed after reboot
'       - Verified Admin privilege for DLL & OCX files that require regsvr32
'       - Verified we have write permission to the destination path
'--- NOW LETS GET FREAKY...

EXT = GetFileExt(sDestinationFile)
INUSE = FileInUse(sDestinationFile)
REGREQ = IIf(EXT = "OCX" Or EXT = "DLL", True, False)

'******************************************************************************************
'***   NOTE: THE FOLLOWING MUST BE AS BULLETPROOF AS OUR COLLECTIVE MINDS CAN MAKE IT   ***
'*** I WELCOME ALL SUGGESTIONS, PLEASE ADVISE WITH IMPROVEMENTS. cwc.software@gmail.com ***
'******************************************************************************************

Restart:
If INUSE Then
    If bOS = 1 Then '------- Win9X/ME
        sSourceFile = GetShortName(sSourceFile)
        '//Once again verify sourcefile is valid once converted to a short path
        If Dir$(sSourceFile, vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem) = "" Then
            UpdateFile = eupdSOURCENOTFOUND
            Exit Function
        End If
        sDestinationFile = GetShortName(sDestinationFile)
        '//Once again verify destinationfile path exist once converted to a short path
        If Dir$(Left$(sDestinationFile, InStrRev(sDestinationFile, "\", , vbTextCompare)), vbDirectory) = "" Then
            UpdateFile = eupdDESTINVALID
            Exit Function
        End If
        '//First setup Wininit to delete current out-of-date file on reboot
        If AddToWininit(sDestinationFile) Then
            '//Now setup Wininit to move up-to-date file to DestinationPath on reboot
            If AddToWininit(sSourceFile, sDestinationFile) Then
                '//Setup Registry to register up-to-date file on reboot once it is moved
                If REGREQ Then
                    With c
                        .ClassKey = HKEY_LOCAL_MACHINE
                        .SectionKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
                        .ValueType = REG_SZ
                        .ValueKey = "Reg " & sDestinationFile
                        .Value = "regsvr32 /s " & sDestinationFile
                        .CreateKey
                    End With
                End If
                UpdateFile = eupdSUCCESSREBOOT
            Else
                UpdateFile = eupdUNKNOWNERR
            End If
        Else
            UpdateFile = eupdUNKNOWNERR
        End If
    Else '------------------ Win NT based
        '//User MUST be an Admin on NT based bOS's for INUSE files in all scenarios. We check this first.
        If Not bADMIN Then UpdateFile = eupdINSUFFPRIV: Exit Function
        '//Make registry entry to move file after reboot
        lResult = MoveFileEx(sSourceFile & Chr(0), sDestinationFile & Chr(0), MOVEFILE_DELAY_UNTIL_REBOOT + MOVEFILE_REPLACE_EXISTING)
        If lResult Then
            '//Make registry entry to delete temp directory after reboot
            lResult = MoveFileEx(Left$(sSourceFile, InStrRev(sSourceFile, "\", -1, vbTextCompare) - 1) & Chr(0), vbNullString, MOVEFILE_DELAY_UNTIL_REBOOT)
            '//Make registry entry to register new file after reboot once it is moved
            If REGREQ Then
                With c
                    .ClassKey = HKEY_LOCAL_MACHINE
                    .SectionKey = "SOFTWARE\Microsoft\Windows\CurrentVersion\RunOnce"
                    .ValueType = REG_SZ
                    .ValueKey = "Reg " & sDestinationFile
                    .Value = "regsvr32 /s " & sDestinationFile
                    .CreateKey
                End With
            End If
            UpdateFile = eupdSUCCESSREBOOT
        Else
            UpdateFile = eupdINSUFFPRIV '------------ Failed to make registry entry
        End If
    End If
Else
    If bOS = 1 Then '--------- Win9X/ME
        sSourceFile = GetShortName(sSourceFile)
        sDestinationFile = GetShortName(sDestinationFile)
        '//Attempt to move the file
        lResult = MoveFile(sSourceFile & Chr(0), sDestinationFile & Chr(0))
        If lResult Then
            If REGREQ Then
                ShellExecute frmMain.hWnd, "open", "regsvr32.exe", "/s " & sDestinationFile, vbNullString, SW_HIDE
            End If
            UpdateFile = eupdSUCCESSCOMP
        Else
            '//Failed to move file (either it is in use or already exist)
            '..Lets try deleting the existing file first.
            lResult = DeleteFile(sDestinationFile)
            If lResult Then
                '//Success deleting existing file, now we are set
                lResult = MoveFile(sSourceFile & Chr(0), sDestinationFile & Chr(0))
                If REGREQ Then
                    ShellExecute frmMain.hWnd, "open", "regsvr32.exe", "/s " & sDestinationFile, vbNullString, SW_HIDE
                End If
                UpdateFile = eupdSUCCESSCOMP
            Else
                '//Failed to move file, probably because is has not been "locked" by
                '..a process, but it is infact in-use (MS calls it a  "memory-mapped file").
                '..When this occurs we cannot directly copy the new file over the existing
                '..one, but we can generally move or rename the existing file, then
                '..attempt to move the new file to the update location for next execution.
                lResult = MoveFile(sDestinationFile & Chr(0), sSourceFile & ".tmp" & Chr(0))
                If lResult Then
                    '//It was moved and renamed, now lets finish this non-sense
                    lResult = MoveFile(sSourceFile & Chr(0), sDestinationFile & Chr(0))
                    If REGREQ Then
                        ShellExecute frmMain.hWnd, "open", "regsvr32.exe", "/s " & sDestinationFile, vbNullString, SW_HIDE
                    End If
                    '//Setup to delete in-use file on reboot
                    Call AddToWininit(sSourceFile & ".tmp" & Chr(0))
                    UpdateFile = eupdSUCCESSCOMP
                Else
                    '//This attempt failed, so lets try again as an INUSE file
                    INUSE = True
                    GoTo Restart
                End If
            End If
        End If
    Else '--------------- NT Based bOS
        '//Attempt to move the file
        lResult = MoveFileEx(sSourceFile & Chr(0), sDestinationFile & Chr(0), MOVEFILE_REPLACE_EXISTING + MOVEFILE_COPY_ALLOWED)
        If lResult Then
            '//Success, now see if the new file needs registered
            If REGREQ Then
                ShellExecute frmMain.hWnd, "open", "regsvr32.exe", "/s " & sDestinationFile, vbNullString, SW_HIDE
            End If
            UpdateFile = eupdSUCCESSCOMP
        Else
            '//Failed to move file, probably because is has not been "locked" by
            '..a process, but it is infact in-use (MS calls it a  "memory-mapped file").
            '..When this occurs we cannot directly copy the new file over the existing
            '..one, but we can generally move or rename the existing file, then
            '..attempt to move the new file to the update location. This method
            '..always works on VB EXE's, so the next file execution will be up-to-date.
            lResult = MoveFileEx(sDestinationFile & Chr(0), sSourceFile & ".tmp" & Chr(0), MOVEFILE_REPLACE_EXISTING + MOVEFILE_COPY_ALLOWED)
            If lResult Then
                '//Moving the in-use file succeeded, now move the new file to the DestinationPath
                lResult = MoveFileEx(sSourceFile & Chr(0), sDestinationFile & Chr(0), MOVEFILE_REPLACE_EXISTING + MOVEFILE_COPY_ALLOWED)
                If REGREQ Then
                    ShellExecute frmMain.hWnd, "open", "regsvr32.exe", "/s " & sDestinationFile, vbNullString, SW_HIDE
                End If
                '//Setup registry to delete in-use file and temp folder on reboot. If the user is not
                '..an Admin the file will be left behind. BUT, we did coherse the update to succeed
                '..and if the user is an Admin we will also be able to clean up our mess.
                If bADMIN Then
                    '//Delete file
                    MoveFileEx sSourceFile & ".tmp" & Chr(0), vbNullString, MOVEFILE_DELAY_UNTIL_REBOOT
                    '//Delete temp directory
                    MoveFileEx Left$(sSourceFile, InStrRev(sSourceFile, "\", -1, vbTextCompare) - 1) & Chr(0), vbNullString, MOVEFILE_DELAY_UNTIL_REBOOT
                End If
                UpdateFile = eupdSUCCESSCOMP
            Else
                '//This attempt failed, so lets try again as an INUSE file
                INUSE = True
                GoTo Restart
            End If
        End If
    End If
End If

Errs_Exit:
    Exit Function

Errs:
    UpdateFile = eupdUNKNOWNERR
    Resume Errs_Exit

End Function

Public Function IsLocalPathValid(ByVal sPath As String, _
    Optional ByVal VerifyDriveExist As Boolean = False) As Boolean
'---------------------------------------------------------------------
' Purpose   : Checks if sPath will pass Windows file and folder naming
'             convention rules.
'---------------------------------------------------------------------
On Error GoTo Errs
Dim sFolders()  As String
Dim sBadChars() As String
Dim sResWords() As String
Dim sDrive      As String
Dim x           As Byte
Dim y           As Byte
    '//Exit if \\ is anywhere in path (UNC Paths NOT Supported)
    If InStr(1, sPath, "\\", vbTextCompare) Then Exit Function
    '//Fill invalid character and reserved word arrays
    sBadChars = Split("\ / : * ? < > | " & Chr(34), " ")
    sResWords = Split("COM1 COM2 COM3 COM4 COM5 COM6 COM7 COM8 COM9 LPT1 LPT2 LPT3 LPT4 LPT5 LPT6 LPT7 " & _
                        "LPT8 LPT9 AUX CLOCK$ CON NUL PRN", " ")
    sFolders = Split(sPath, "\") '------------- Fill array with drive and folders
    sDrive = LCase$(sFolders(0)) '-------------- Extract drive and check if valid
    If VerifyDriveExist Then
        If Dir$(sDrive, 63) = "" Then Exit Function
    Else
        For x = 97 To 122 '-------------------- Check to ensure drive letter is a - z
            If sDrive = Chr(x) & ":" Then Exit For
        Next x
        If x = 123 Then Exit Function '-------- Drive letter was not a through z
    End If
    
    For y = 1 To UBound(sFolders)
        For x = 0 To 7 '----------------------- Check for invalid folder characters
            If InStr(1, sFolders(y), sBadChars(x)) Then Exit Function
        Next x
        For x = 0 To 22 '---------------------- Check for reserved words
            If UCase$(sFolders(y)) = sResWords(x) Then Exit Function
        Next x
    Next y
    IsLocalPathValid = True
Errs:
    If Err Then Exit Function
End Function

Public Function GetFileExt(ByVal sFile As String) As String
'---------------------------------------------------------
' Purpose   : Returns a files file extension if one exist.
'---------------------------------------------------------
On Error GoTo Errs
Dim x   As Long
Dim y   As Long
    x = InStrRev(sFile, ".")
    If x Then '---------------------- Skip if a "." is not found
        y = InStrRev(sFile, "\")
        If y Then
            If y < x Then '---------- Be sure "." is to the right of last "\"
                GetFileExt = UCase$(Mid$(sFile, x + 1))
            End If
        Else '----------------------- For passing only a filename without a path
            GetFileExt = UCase$(Mid$(sFile, x + 1))
        End If
    End If
Errs:
    If Err Then GetFileExt = ""
End Function

Public Function WindowsVersion() As Long
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

Public Function CanWriteToPath(ByVal sPath As String) As Boolean
'-------------------------------------------------------------------
' Purpose   : Checks for write permission to a given path. sPath must
'             exist to return True. Supports local and UNC paths.
'--------------------------------------------------------------------
On Error GoTo Errs
Dim f  As Integer
    If Len(sPath) < 2 Then Exit Function
    If Right$(sPath, 1) <> "\" Then sPath = sPath & "\"
    sPath = sPath & "WriteCheck.tmp" '------------------- Create temp file
    f = FreeFile
    Open sPath For Output As #f '------------------------ Attempt write
        Print #f, "Test"
    Close #f
    CanWriteToPath = True
    On Error Resume Next
    Call DeleteFile(sPath) '----------------------------- Delete temp file
Errs:
    If Err Then Exit Function
End Function

Public Function IsAdministrator() As Boolean
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
    '//See if OS is Win9X or ME.
    '..This is the only thing I edited in this routine.
    If bOS = 1 Then
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


'****************************************************************************************
'                               Private Helper Functions
'****************************************************************************************

Private Function WindowsDirectory() As String
'--------------------------------------------------------------
' Purpose   : Returns Windows directory
'--------------------------------------------------------------
Dim path            As String * 255
Dim ReturnLength    As Long
    ReturnLength = GetWindowsDirectory(path, Len(path))
    WindowsDirectory = Left$(path, ReturnLength)
End Function

Private Function FileInUse(ByVal sFile As String) As Boolean
'---------------------------------------------------------
' Purpose   : Attempts to open sFile in EXCLUSIVE mode.
'             Returns True if fails, False if succeeds.
'---------------------------------------------------------
Dim hf  As Long
Dim fo  As OFSTRUCT
    sFile = Trim$(sFile)
    If Len(sFile) = 0 Or Dir$(sFile, 39) = "" Then Exit Function
    If Right$(sFile, 1) <> Chr(0) Then sFile = sFile & Chr(0)
    fo.cBytes = Len(fo)
    hf = OpenFile(sFile, fo, OF_SHARE_EXCLUSIVE) '---- Attempt EXCLUSIVE File Open
    If hf = -1 And Err.LastDllError = 32 Then
        FileInUse = True '---------------------------- File failed to open (In Use)
    Else
        CloseHandle hf '------------------------------ File was opened (Not In Use)
    End If
End Function

Private Function GetShortName(ByVal sFile As String) As String
'--------------------------------------------------------------------------
' Purpose   : Get Windows short file name (works if file does not exist).
'             This is needed when writing to wininit.ini for Win9X/ME only.
'--------------------------------------------------------------------------
On Error Resume Next
Dim sString     As String * 255
Dim lResult     As Long
Dim bCreated    As Boolean
Dim f           As Integer
    f = FreeFile
    If Dir$(sFile, vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem) = "" Then
        Open sFile For Output As #f
        Close #f
        bCreated = True
    End If
    lResult = GetShortPathName(sFile, sString, 255)
    GetShortName = Left$(sString, lResult)
    If bCreated Then Call DeleteFile(sFile)
End Function

Private Function AddToWininit(ByVal sSourceFile As String, _
    Optional ByVal sDestinationFile As String = "Nul") As Boolean
'------------------------------------------------------------------------
' Purpose   : Writes files that need replaced or deleted at reboot to the
'             wininit.ini file. ***FOR Win9X AND ME ONLY***
'------------------------------------------------------------------------
On Error GoTo Errs
Dim sFile       As String '-------- Path for wininit.ini
Dim f           As Long '---------- Freefile assignment
Dim line        As String '-------- Current line from wininit.ini
Static bFound   As Boolean '------- True if [Rename] section was located
    sFile = WindowsDirectory & "\wininit.ini"
    '//Skip this block if previously done for another entry
    If Not bFound Then
        '//See if wininit.ini file exist in the Windows directory, create it if not
        If Dir$(sFile, vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem) = "" Then
            f = FreeFile
            Open sFile For Random As #f
            Close #f
        End If
        '//Scan wininit.ini for a rename section
        f = FreeFile
        Open sFile For Input As #f
        Do While Not EOF(f)
            Line Input #f, line
            If InStr(1, UCase$(line), "[RENAME]", vbTextCompare) Then
                bFound = True
                Exit Do
            End If
        Loop
        Close #f
    End If
    '//Append rename section with passed file to delete or replace at reboot.
    f = FreeFile
    Open sFile For Append As #f
    If Not bFound Then Print #f, "[Rename]"
    Print #f, sDestinationFile & "=" & sSourceFile
    bFound = True
    Close #f
    AddToWininit = True
Errs:
    If Err Then Exit Function
End Function

Private Function CreatePath(ByVal sPath As String) As Boolean
'---------------------------------------------------------------------
' Purpose   : Checks if sPath will pass Windows file and folder naming
'             convention rules, and if so creates the path.
'---------------------------------------------------------------------
On Error GoTo Errs
Dim sFolders()  As String
Dim sNewPath    As String
Dim x           As Long
Dim ub          As Long
    If IsLocalPathValid(sPath, True) Then '-------------- Check naming conventions, etc.
        sFolders = Split(sPath, "\") '------------------- Parse drive and folders
        ub = UBound(sFolders)
        sNewPath = sFolders(0) '------------------------- Extract drive and check for existence
        If Dir$(sNewPath, 63) <> "" Then
            If ub Then
                For x = 1 To ub '------------------------ Create path one folder at a time
                    sNewPath = sNewPath & "\" & sFolders(x)
                    If Dir$(sNewPath, vbDirectory) = "" Then MkDir$ sNewPath
                Next x
                CreatePath = True
            End If
        End If
    End If
Errs:
    If Err Then Exit Function
End Function

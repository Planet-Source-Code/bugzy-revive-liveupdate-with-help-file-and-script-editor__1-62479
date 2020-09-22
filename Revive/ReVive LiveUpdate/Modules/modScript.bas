Attribute VB_Name = "modScript"
'***********************************************************************
'Chris Cochran          cwc.software@gmail.com        Updated: 18 Oct 05
'***********************************************************************
Option Explicit

'//Status of update file progress
Public Enum eStatus
    NOUPDATEREQ = 0                     '//Update not required
    UPDATEREQ = 1                       '//Update required
    DOWNLOADING = 2                     '//Currently downloading
    DOWNLOADED = 3                      '//Finished download
    ERRCONNECTING = 4                   '//Could not connect to host or file
    ERRTRANSFERRING = 5                 '//Transfer started but failed
    INSUFFPRIVILEGE = 6                 '//User must be an admin and is not
    ERRUPDATING = 7                     '//TestUpdateSuccess or UpdateFile returned non-zero value
    UPDATECOMP = 8                      '//Update complete
    UPDATECOMPREBOOT = 9                '//Update complete, reboot required
    UPDATEREADY = 10                    '//TestUpdateSuccess successful, ready for update
    FILENOTINSTALLED = 11               '//Will not be transferred, script has MustBeInstalled=1 and file is not
End Enum

'//Run mode enumerators
Public Enum eRunMode
    eNORMAL = 0
    eNOTIFY = 1
    eAUTO = 2
End Enum

'//Update file list array (Most variables are filled in ParseUpdateScript routine)
Public Type tFileList
    Description()           As String   '//Short File Description
    UpdateVersion()         As String   '//Update File Version
    CurrentVersion()        As String   '//Installed File Version
    DownloadURL()           As String   '//Web download path (URL)
    FileSize()              As Long     '//Size of download (To show total download progress)
    InstallPath()           As String   '//Location of client file to update (Use directory constants, i.e. {APPPATH}\MyProg.exe = App.Path & "\MyProg.exe" on client
    FileName()              As String   '//Filename of update file
    TempPath()              As String   '//Temporary storage path for downloaded while waiting processing
    MustExist()             As Boolean  '//File must exist on users machine to be transferred
    MustUpdate()            As Boolean  '//1 if file must update if required to continue LiveUpdate
    UpdateMessage()         As String   '//String displayed when ReVive exits if file is updated on client
    Status()                As eStatus  '//Progress state of file from eStatus Enumerator
End Type

'***START DECLARES FOR GetFileVersion***
Private Declare Sub CopyMem Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, Source As Any, ByVal Length As Long)
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Byte, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Private Type VS_FIXEDFILEINFO
    dwSignature             As Long
    dwStrucVersion          As Long
    dwFileVersionMS         As Long
    dwFileVersionLS         As Long
    dwProductVersionMS      As Long
    dwProductVersionLS      As Long
    dwFileFlagsMask         As Long
    dwFileFlags             As Long
    dwFileOS                As Long
    dwFileType              As Long
    dwFileSubtype           As Long
    dwFileDateMS            As Long
    dwFileDateLS            As Long
End Type
'***END DECLARES FOR GetFileVersion***

'//Directory constants available for use in web update script.
'//Want to add more? See ReplaceConstants routine below.
'//CAUTION: BE SURE ADDITIONS ARE WINDOWS VERSION FRIENDLY OR
'//IMPLEMENT A PLAN B (LIKE IIf) TO DEFAULT TO A COMMON VALUE.
Private ap                  As String   '//App.Path from where this exe is executed
Private sp                   As String   '//Path of setup script from arguments in Sub Main (App.Path if script not specific)
Private win                 As String   '//Windows directory (Like C:\Winnt)
Private sys                 As String   '//System directory (Like C:\Winnt\System32)
Private temp                As String   '//Windows Temp directory (Like C:\Winnt\Temp)
Private pf                  As String   '//Program files directory (Like C:\Program Files)
Private cf                  As String   '//Common files path (Like C:\Program Files\Common Files)
Private userdesktop         As String   '//Current users desktop
Private userstartmenu       As String   '//Current users start menu
Private commondesktop       As String   '//All users Desktop directory (Like C:\Documents and Settings\All Users\Desktop)
Private commonstartmenu     As String   '//All users start menu

'//LiveUpdate Settings declares
Public Type tSetup
    SetupScriptPath         As String
    AppShortName            As String
    AppLongName             As String
    AdminRequired           As Boolean
    ForceReboots            As Boolean
    ScriptURLPrim           As String
    ScriptURLAlt            As String
    LastChecked             As String
    RunMode                 As eRunMode
    NotifyIcon              As String
    UpdateAppTitle          As String
    UpdateAppClass          As String
    UpdateAppKill           As Boolean
    LaunchIfKilled          As String
    ShowFileIcons           As Boolean
    RegRISFiles             As Boolean
    HideAbsentFiles         As Boolean
End Type

Public FileList             As tFileList    '//Makes update file list available to entire app
Public Setup                As tSetup       '//Stored LiveUpdate settings for active script
Public bREBOOT              As Boolean      '//Set to True when any update file requires a reboot to update (regsvr32)

Public Function ParseUpdateScript(ByVal sFile As String) As Byte
'------------------------------------------------------------------------
' Purpose   : Opens downloaded update file, reads and parses update list,
'             and fills FileList array with update file list entries.
'             This routine also ensures script info is valid, admin
'             priviledge exist when required, and MustExist files are
'             present on users computer. All this before downloads begin.
'
' Returns   : 0 if success
'             1 if client must be an Administrator to continue and is not
'             2 if an error was encountered processing script
'             3 if there are no updates listed in the update script
'------------------------------------------------------------------------
On Error GoTo Errs
Dim f           As Integer '--- Freefile assignment
Dim s           As String '---- Misc string uses
Dim x           As Long '------ Misc long uses
Dim y           As Long
Dim sEXT        As String '---- File extension
Dim reg         As Boolean '--- True if any update files are OCX or DLL
Dim sec         As String '---- Current [File XX] section we are searching in

    If Dir$(sFile, vbNormal + vbReadOnly + vbArchive + vbHidden + vbSystem) = "" Then
        ParseUpdateScript = 2 '--- sFile not found, return script error
        Exit Function
    End If
    f = FreeFile
    Open sFile For Input As #f '-- Open sFile for parsing setup information
    Call AssignContants '--------- Assign script contants to client folders
    With Setup '------------------ Get Script Setup settings
        .AdminRequired = CBool(ProfileGetItem("SETUP", "AdminRequired", 0, sFile))
        .ForceReboots = CBool(ProfileGetItem("SETUP", "ForceReboots", 0, sFile))
        .ScriptURLPrim = ProfileGetItem("SETUP", "ScriptURLPrim", frmMain.ucDL.dScriptURLPrim, sFile)
        .ScriptURLAlt = ProfileGetItem("SETUP", "ScriptURLAlt", frmMain.ucDL.dScriptURLAlt, sFile)
        .AppShortName = ProfileGetItem("SETUP", "AppShortName", frmMain.ucDL.dAppShortName, sFile)
        .AppLongName = ProfileGetItem("SETUP", "AppLongName", frmMain.ucDL.dAppLongName, sFile)
        .NotifyIcon = ProfileGetItem("SETUP", "NotifyIcon", "", sFile)
        .UpdateAppTitle = ProfileGetItem("SETUP", "UpdateAppTitle", "", sFile)
        .UpdateAppClass = ProfileGetItem("SETUP", "UpdateAppClass", "", sFile)
        .UpdateAppKill = ProfileGetItem("SETUP", "UpdateAppKill", False, sFile)
        .LaunchIfKilled = ReplaceConstants(ProfileGetItem("SETUP", "LaunchIfKilled", "", sFile))
        .ShowFileIcons = ProfileGetItem("SETUP", "ShowFileIcons", True, sFile)
        .RegRISFiles = ProfileGetItem("SETUP", "RegRISFiles", False, sFile)
        .HideAbsentFiles = ProfileGetItem("SETUP", "HideAbsentFiles", False, sFile)
        .HideAbsentFiles = True
    End With
    With FileList '--------------- Locate all update file entries in script
        Do
            x = x + 1
            sec = "File " & Format(x, "00")
            s = ProfileGetItem(sec, "Description", "", sFile)
            If Len(s) Then
                ReDim Preserve .Description(1 To x)
                ReDim Preserve .UpdateVersion(1 To x)
                ReDim Preserve .CurrentVersion(1 To x)
                ReDim Preserve .DownloadURL(1 To x)
                ReDim Preserve .InstallPath(1 To x)
                ReDim Preserve .FileName(1 To x)
                ReDim Preserve .TempPath(1 To x)
                ReDim Preserve .MustExist(1 To x)
                ReDim Preserve .FileSize(1 To x)
                ReDim Preserve .MustUpdate(1 To x)
                ReDim Preserve .UpdateMessage(1 To x)
                ReDim Preserve .Status(1 To x)
                .Description(x) = s
                .UpdateVersion(x) = ProfileGetItem(sec, "UpdateVersion", "0.0.0.0", sFile)
                .DownloadURL(x) = ProfileGetItem(sec, "DownloadURL", "", sFile)
                .InstallPath(x) = ReplaceConstants(ProfileGetItem(sec, "InstallPath", "", sFile))
                .FileName(x) = Right$(.InstallPath(x), Len(.InstallPath(x)) - InStrRev(.InstallPath(x), "\"))
                .FileSize(x) = ProfileGetItem(sec, "FileSize", 0, sFile)
                .MustExist(x) = CBool(ProfileGetItem(sec, "MustExist", False, sFile))
                .MustUpdate(x) = CBool(ProfileGetItem(sec, "MustUpdate", False, sFile))
                .UpdateMessage(x) = Trim$(ProfileGetItem(sec, "UpdateMessage", "", sFile))
                .TempPath(x) = sTEMPDIR & "\" & GetFileName(.InstallPath(x))
                sEXT = GetFileExt(.InstallPath(x))
                '//Check for files that must be registered using Regsvr32
                If sEXT = "OCX" Or sEXT = "DLL" Then
                    reg = True
                End If
                '******************** START PRELIMINARY SCRIPT INFO VALIDATION ***********************
                '//Values not validated below are good because they met data type criteria above
                '..or will be validated later when testing update success before updating.
                
                '//Validate that supplied version number is valid
                If IsVersionValid(.UpdateVersion(x)) = 0 Then
                    ParseUpdateScript = 2 '------------- Return script error, invalid version number
                    Exit Do
                End If
                '//Verify InstallPath is valid when MustUpdate = True
                If .MustUpdate(x) Then
                    If Not IsLocalPathValid(.InstallPath(x), True) Then
                        ParseUpdateScript = 2 '--------- Return script error, install path invalid
                        Exit Do
                    End If
                End If
                '//Verify file exist on clients machine when MustBeInstalled = True
                If .MustExist(x) Or (.MustExist(x) And Setup.HideAbsentFiles) Then
                    If Dir$(.InstallPath(x), 39) = "" Then
                        '//Not found and required. This will cause an abort in frmMain's ListScanResults
                        '..routine if file is flagged as MustUpdate.
                        .Status(x) = FILENOTINSTALLED
                    End If
                End If
                '******************** END PRELIMINARY SCRIPT INFO VALIDATION *************************
            Else
                Exit Do
            End If
        Loop
        Close #f '-------------------------------- Close downloaded web update script
    End With
    If x = 1 And ParseUpdateScript = 0 Then '----- No update files found in script
        ParseUpdateScript = 3
    End If
    If ParseUpdateScript = 0 Then
        '//Prepare listview for vertical scrollbar if file count exceeds visible capability
        y = IIf(Setup.ShowFileIcons, 11, 13)
        If x > y Then frmMain.lvFiles.ColumnHeaders(1).Width = 3000
        '//Check for sufficient privilege
        If Setup.AdminRequired Or reg Then '----- Check both script flag and DLL/OCX registration requirements
            If Not bADMIN Then '----------------- Check if client is an Administrator
                ParseUpdateScript = 1 '---------- Return permission insufficient
                GoTo Errs_Exit '----------------- Be done with this nonsense - the users a peon
            End If
        End If
        If bADMIN And Setup.RegRISFiles Then '--- Register .ris files to open with ReVive
            Call RegRISFile
        End If
        Call CompareFileVersions '--------------- Script read fine and all files passed above tests
    End If
Errs_Exit:
    On Error Resume Next
    Call DeleteFile(sFile) '--------------------- Delete downloaded script file
    Exit Function
Errs:
    ParseUpdateScript = 2 '---------------------- Return script error
    Resume Errs_Exit
End Function

'****************************************************************************************
'                               Private Helper Functions
'****************************************************************************************

Private Sub AssignContants()
'------------------------------------------------------------------------
' Purpose   : Assigns local folders to available update script constants.
'------------------------------------------------------------------------
On Error Resume Next
    ap = App.path
    '//Remove right most '\' if ap is located on a root drive
    If Right$(ap, 1) = "\" Then ap = Left$(ap, InStrRev(ap, "\") - 1)
    sp = Left$(Setup.SetupScriptPath, InStrRev(Setup.SetupScriptPath, "\") - 1)
    win = GetFolderPath(CSIDL_WINDOWS)
    sys = GetFolderPath(CSIDL_SYSTEM)
    temp = win & "\Temp"
    pf = GetFolderPath(CSIDL_PROGRAM_FILES)
    cf = GetFolderPath(CSIDL_PROGRAM_FILES_COMMON)
    userdesktop = GetFolderPath(CSIDL_DESKTOPDIRECTORY)
    commondesktop = IIf(Len(GetFolderPath(CSIDL_COMMON_DESKTOPDIRECTORY)) = 0, userdesktop, GetFolderPath(CSIDL_COMMON_DESKTOPDIRECTORY))
    userstartmenu = GetFolderPath(CSIDL_STARTMENU)
    commonstartmenu = IIf(Len(GetFolderPath(CSIDL_COMMON_STARTMENU)) = 0, userstartmenu, GetFolderPath(CSIDL_COMMON_STARTMENU))
End Sub

Private Function ReplaceConstants(ByVal sString As String) As String
'*******************************************************************
'Called from ParseUpdateScript routine for InstallPath contants.
'
'If an unrecognized constant was used, an empty string is returned,
'which will cause a script error in ParseUpdateScript routine.
'
'WANT TO ADD MORE? (Checkout modFolders for more possibilites)
'   Step 1: Update list in declaration section of this module.
'   Step 2: Add to AssignContants sub above and assign a value.
'   Step 3: Insert an InStr for each addition in this procedure.
'*******************************************************************
    '//Verify a constant was used before continuing
    If InStr(1, sString, "<", vbTextCompare) = 0 And _
            InStr(1, sString, ">", vbTextCompare) = 0 Then
        ReplaceConstants = sString
        Exit Function
    End If
    '-------------ONLY ONE CONSTANT PER PATH PROCESSED--------------
    '-----------MOST COMMONLY USED AT TOP FOR EFFICIENCY------------
    If InStr(1, sString, "<sp>", vbTextCompare) Then
        sString = Replace(sString, "<sp>", sp)
    ElseIf InStr(1, sString, "<ap>", vbTextCompare) Then
        sString = Replace(sString, "<ap>", ap)
    ElseIf InStr(1, sString, "<sys>", vbTextCompare) Then
        sString = Replace(sString, "<sys>", sys)
    ElseIf InStr(1, sString, "<win>", vbTextCompare) Then
        sString = Replace(sString, "<win>", win)
    ElseIf InStr(1, sString, "<temp>", vbTextCompare) Then
        sString = Replace(sString, "<temp>", temp)
    ElseIf InStr(1, sString, "<userdesktop>", vbTextCompare) Then
        sString = Replace(sString, "<userdesktop>", userdesktop)
    ElseIf InStr(1, sString, "<userstartmenu>", vbTextCompare) Then
        sString = Replace(sString, "<userstartmenu>", userstartmenu)
    ElseIf InStr(1, sString, "<pf>", vbTextCompare) Then
        sString = Replace(sString, "<pf>", pf)
    ElseIf InStr(1, sString, "<cf>", vbTextCompare) Then
        sString = Replace(sString, "<cf>", cf)
    Else
        '//An unrecognized constant was used (will be caught by ParseUpdateScript)
        Exit Function
    End If
    '//Verify no more than one constant was used
    If InStr(1, sString, "<", vbTextCompare) Or _
            InStr(1, sString, ">", vbTextCompare) Then
        sString = ""
    End If
    ReplaceConstants = sString
End Function

Private Sub CompareFileVersions()
'----------------------------------------------------------------------------------
' Purpose   : Called from ParseUpdateScript routine. Determines what files from
'             the FileList array require updating by LiveUpdate.
'
'             NOTE: Existing OCX, DLL and EXE file versions are gained from the
'             file itself, while all others are extracted from the local RIS file.
'----------------------------------------------------------------------------------
Dim sNewVer()   As String
Dim sExistVer() As String
Dim i           As Byte
Dim x           As Long
Dim s           As String
    With FileList
        For x = 1 To UBound(.Description)
            '//First get the existing file version
            s = GetFileExt(.InstallPath(x))
            If InStr(1, "|EXE|OCX|DLL|", s) Then
                '//Pull version info from file
                .CurrentVersion(x) = GetFileVersion(.InstallPath(x))
                '//If file did not contain version info, check the setup script
                If .CurrentVersion(x) = "0.0.0.0" Then
                    '//First see if file exist on client before getting ver info from clients setup script
                    If Dir$(.InstallPath(x), vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem) <> "" Then
                        .CurrentVersion(x) = ProfileGetItem("Files", .Description(x), "0.0.0.0", Setup.SetupScriptPath)
                    End If
                End If
            Else
                '//First see if file exist on client before getting ver info from clients setup script
                If Dir$(.InstallPath(x), vbArchive + vbHidden + vbNormal + vbReadOnly + vbSystem) = "" Then
                    .CurrentVersion(x) = "0.0.0.0"
                Else
                    .CurrentVersion(x) = ProfileGetItem("Files", .Description(x), "0.0.0.0", Setup.SetupScriptPath)
                End If
            End If
            '//Break version numbers down to 4 different segments for comparing each one
            sNewVer() = Split(.UpdateVersion(x), ".", , vbTextCompare)
            sExistVer() = Split(.CurrentVersion(x), ".", , vbTextCompare)
            If .Status(x) <> FILENOTINSTALLED Then
                '//Compare each segment until we hit a newer version number or try all four segments
                For i = 0 To 3
                    If CLng(sNewVer(i)) > CLng(sExistVer(i)) Then
                        .Status(x) = UPDATEREQ
                        Exit For
                    End If
                Next i
            End If
        Next x
    End With
End Sub

Public Function GetFileVersion(ByVal sFileName As String) As String
'***********************************************************************
' Purpose   : Get the file version number from DLL, EXE, or OCX files.
'
' Adapted from code posted by Eric D. Burdo, http://www.rlisolutions.com
' "Retrieve the version number of a DLL"
'
' Full Post : http://programmers-corner.com/viewSource.php/71
'***********************************************************************
Dim lFreeSize   As Long
Dim tVerBuf()   As Byte
Dim sVerInfo    As VS_FIXEDFILEINFO
Dim lFreeHandle As Long
Dim lBuff       As Long
Dim iMajor      As Integer
Dim iMinor      As Integer
Dim sMajor      As String
Dim sMinor      As String
    lFreeSize = GetFileVersionInfoSize(sFileName, lFreeHandle)
    If lFreeSize Then
        If lFreeSize > 64000 Then lFreeSize = 64000
        ReDim tVerBuf(lFreeSize)
        GetFileVersionInfo sFileName, 0&, lFreeSize, tVerBuf(0)
        VerQueryValue tVerBuf(0), "\" & "", lBuff, lFreeSize
        CopyMem sVerInfo, ByVal lBuff, lFreeSize
    End If
    iMajor = CInt(sVerInfo.dwFileVersionMS \ &H10000)
    iMinor = CInt(sVerInfo.dwFileVersionMS And &HFFFF&)
    sMajor = CStr(iMajor) & "." & LTrim$(CStr(iMinor))
    iMajor = CInt(sVerInfo.dwFileVersionLS \ &H10000)
    iMinor = CInt(sVerInfo.dwFileVersionLS And &HFFFF&)
    sMinor = CStr(iMajor) & "." & LTrim$(CStr(iMinor))
    GetFileVersion = sMajor & "." & sMinor
End Function

Private Function IsVersionValid(ByVal sVersion As String) As Boolean
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

Private Function GetFileName(ByVal sFilePath As String) As String
'--------------------------------------------------
' Purpose   : Returns filename from sFilePath path
'--------------------------------------------------
On Error GoTo Errs
Dim x As Long
    x = InStrRev(sFilePath, "\")
    If x Then sFilePath = Mid$(sFilePath, x + 1)
    GetFileName = sFilePath
Errs:
    If Err Then GetFileName = ""
End Function

Private Sub RegRISFile()
'--------------------------------------------------------------------
' Purpose   : Registers .ris files to open with the ReVive executable
'--------------------------------------------------------------------
On Error GoTo Errs
Dim c       As New cReg
    If Dir(App.path & "\" & App.EXEName & ".exe", 39) = "" Then Exit Sub
    With c
        .ClassKey = HKEY_CLASSES_ROOT
        .SectionKey = ".ris"
        .ValueType = REG_SZ
        .ValueKey = ""
        .Value = "ReVive.Initialization.Script"
        .CreateKey
        .SectionKey = "ReVive.Initialization.Script"
        .ValueKey = ""
        .Value = "ReVive LiveUpdate Initialization Script"
        .CreateKey
        .SectionKey = "ReVive.Initialization.Script\shell\open\command"
        .ValueKey = ""
        .Value = Chr(34) & App.path & "\" & App.EXEName & ".exe" & Chr(34) & " " & Chr(34) & "%1" & Chr(34)
        .CreateKey
    End With
Errs_Exit:
    Set c = Nothing
    Exit Sub
Errs:
    Resume Errs_Exit
End Sub

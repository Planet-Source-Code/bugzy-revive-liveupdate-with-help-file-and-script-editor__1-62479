Attribute VB_Name = "modDistribute"
Option Explicit

Private Declare Function GetPrivateProfileString Lib "kernel32" Alias "GetPrivateProfileStringA" (ByVal lpSectionName As String, ByVal lpKeyName As Any, ByVal lpDefault As String, ByVal lpReturnedString As String, ByVal nSize As Long, ByVal lpFileName As String) As Long

Private Type tDistSetup '---------------------- ReVive Setup Type declares
    AppShortName            As String
    AppLongName             As String
    ScriptURLPrim           As String
    ScriptURLAlt            As String
End Type

Private Type tFileList '----------------------- Update file list array
    Description()           As String '-------- Update File Description
    UpdateVersion()         As String '-------- Update File Version
    DownloadURL()           As String '-------- Web download path (URL)
End Type

Private Const CRYPTPWD      As String = "ReVive" '--- Password for encrypting web URL's in .ris files
'>*>*>*>*>*>*>*>*>*>*>*>*>*>*>*>*>*>*>*>*>*>*'... ***ONCE ESTABLISHED FOR AN APP DO NOT CHANGE***

Private DistSetup           As tDistSetup
Private DistFileList        As tFileList

Public Sub CreateDistributable()
On Error GoTo Errs
Dim s       As String
Dim x       As Long
Dim Y       As Long
Dim sItem   As String
Dim sExt    As String
Dim f       As Integer
Dim sFile   As String
    With DistSetup '----------- Extract Setup Options
        .AppShortName = ProfileGetItem("Setup", "AppShortName", "", Setup.Script)
        .AppLongName = ProfileGetItem("Setup", "AppLongName", "", Setup.Script)
        .ScriptURLPrim = ProfileGetItem("Setup", "ScriptURLPrim", "", Setup.Script)
        .ScriptURLAlt = ProfileGetItem("Setup", "ScriptURLAlt", "", Setup.Script)
    End With
    x = 1
    sItem = "File 01"
    With DistFileList '-------- Extract files that need versions stored in distributable
        Do While ProfileGetItem(sItem, "Description", "", Setup.Script) <> ""
            sExt = GetFileExt(ProfileGetItem(sItem, "InstallPath", "", Setup.Script))
            If sExt <> "OCX" And sExt <> "DLL" And sExt <> "EXE" Then
                ReDim Preserve .Description(0 To Y)
                ReDim Preserve .DownloadURL(0 To Y)
                ReDim Preserve .UpdateVersion(0 To Y)
                .Description(Y) = ProfileGetItem(sItem, "Description", "", Setup.Script)
                .DownloadURL(Y) = ProfileGetItem(sItem, "DownloadURL", "", Setup.Script)
                .UpdateVersion(Y) = ProfileGetItem(sItem, "UpdateVersion", "", Setup.Script)
                Y = Y + 1
            End If
            x = x + 1
            sItem = "File " & Format(x, "00")
        Loop
    End With
    If x = 1 Then '------------ Bolt if no update files are found
        MsgBox "The Script you are trying to distribute does not contain any valid update files.  " & vbNewLine & _
               "Please add update files and try this operation again.  ", vbExclamation, "Error Creating Distributable"
        Exit Sub
    End If
    sFile = frmDistribute.SelectDistFile
    If Len(sFile) = 0 Then Exit Sub
    s = vbNewLine
    With DistSetup '----------- Build [Setup] setion
        s = s & .AppShortName & " ReVive Initialization Script"
        s = s & vbNewLine & vbNewLine
        s = s & "--------------------------------------------------------------------------------"
        s = s & vbNewLine
        s = s & "***WARNING: MODIFYING THIS SCRIPT MAY CAUSE LIVEUPDATE TO FUNCTION IMPROPERLY***"
        s = s & vbNewLine & vbNewLine
        s = s & "[Setup]" & vbNewLine
        s = s & "AppShortName=" & .AppShortName & vbNewLine
        s = s & "AppLongName=" & .AppLongName & vbNewLine
        s = s & "ScriptURLPrim=" & EncryptString(.ScriptURLPrim) & vbNewLine
        s = s & "ScriptURLAlt=" & EncryptString(.ScriptURLAlt) & vbNewLine & vbNewLine
    End With
    If Y Then
        s = s & "[Files]" & vbNewLine
        With DistFileList '---- Build [Files] section
            For x = 0 To Y - 1
                s = s & .Description(x) & "=" & .UpdateVersion(x) & vbNewLine
            Next x
        End With
        s = s & vbNewLine
    End If
    s = s & "***WARNING: MODIFYING THIS SCRIPT MAY CAUSE LIVEUPDATE TO FUNCTION IMPROPERLY***"
    s = s & vbNewLine
    s = s & "--------------------------------------------------------------------------------"
    f = FreeFile
    Open sFile For Output As #f
        Print #f, s
    Close #f
    MsgBox "The " & DistSetup.AppShortName & " ReVive initialization script was created successfully. " & vbNewLine & vbNewLine & _
           "Please distribute this file in your applications installation directory.  ", vbInformation + vbOKOnly, "ReVive"
Errs_Exit:
    Exit Sub
Errs:
    MsgBox "The following error was returned while attempting to create your distributable:  " & vbNewLine & vbNewLine & _
           Err.Description, vbExclamation, "Error Creating Distributable"
    Resume Errs_Exit
End Sub

Private Function ProfileGetItem(lpSectionName As String, _
                               lpKeyName As String, _
                               defaultValue As String, _
                               inifile As String) As String
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

Private Function GetFileExt(sFile As String) As String
'---------------------------------------------------------
' Purpose   : Returns a files file extension if one exist.
'---------------------------------------------------------
On Error GoTo Errs
Dim x As Long
Dim Y As Long
    x = InStrRev(sFile, ".")
    If x Then '---------------------- Skip if a "." is not found
        Y = InStrRev(sFile, "\")
        If Y Then
            If Y < x Then '---------- Be sure "." is to the right of last "\"
                GetFileExt = UCase$(Mid$(sFile, x + 1))
            End If
        Else '----------------------- For passing only a filename without a path
            GetFileExt = UCase$(Mid$(sFile, x + 1))
        End If
    End If
Errs:
    If Err Then GetFileExt = ""
End Function

Private Function EncryptString(ByVal sString As String) As String
'---------------------------------------------------------------------------------------
' Purpose   : Encrypts web URL strings prior to entry into the initialization script.
'             Only called from the CreateDistributable sub.
'---------------------------------------------------------------------------------------
Dim x       As Integer
Dim Y       As Integer
Dim sBuffer As String
    If Len(CRYPTPWD) Then
        For x = 1 To Len(sString)
            Y = Asc(Mid$(sString, x, 1))
            Y = Y + Asc(Mid$(CRYPTPWD, (x Mod Len(CRYPTPWD)) + 1, 1))
            sBuffer = sBuffer & Chr$(Y And &HFF)
        Next x
        EncryptString = sBuffer
    Else
        EncryptString = sString
    End If
End Function

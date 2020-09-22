VERSION 5.00
Begin VB.UserControl ccDownload 
   Appearance      =   0  'Flat
   BackColor       =   &H80000005&
   CanGetFocus     =   0   'False
   ClientHeight    =   420
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   435
   ClipControls    =   0   'False
   BeginProperty Font 
      Name            =   "Verdana"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   HitBehavior     =   0  'None
   InvisibleAtRuntime=   -1  'True
   Picture         =   "ccDownload.ctx":0000
   ScaleHeight     =   420
   ScaleWidth      =   435
   ToolboxBitmap   =   "ccDownload.ctx":0242
End
Attribute VB_Name = "ccDownload"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
' Module    : ccDownload
' Updated   : 13 Sep 2005
' Author    : Chris Cochran
' Purpose   : To handle downloads and report progress for ReVive project.
' Info      : The usercontrol AsynRead method uses local Internet Explorer
'             connection settings and supports both FTP and HTTP downloads.
'---------------------------------------------------------------------------------------
Option Explicit

Public Enum eDownloadResults
    eSUCCESS = 0
    eCONNECTERROR = 1
    eTRANSFERERROR = 2
    eWRITEERROR = 3
End Enum

Private sCurrIdentifier             As String '----- Identifier for file downoad request
Private sAppShortName               As String '----- Application short name
Private sAppLongName                As String '----- Application long name
Private sScriptURLPrim              As String '----- Primary script download URL
Private sScriptURLAlt               As String '----- Alternate script download URL

Public Event DownloadProgress(Identifier As String, RecvdBytes As Long, CurBytes As Long, MaxBytes As Long)
Public Event DownloadComplete(Identifier As String, Result As eDownloadResults)

Public Sub Download(ByVal sURL As String, ByVal Identifier As String)
'******************************************************************
'Initiate file downloads here using the AsyncRead method.
'The Identifier will be the [File XX] number of your script entry.
'******************************************************************
On Error GoTo Errs
    If Setup.RunMode = eNORMAL Then Screen.MousePointer = vbHourglass
    sCurrIdentifier = Identifier
    UserControl.AsyncRead sURL, vbAsyncTypeByteArray, Identifier, vbAsyncReadForceUpdate
Errs_Exit:
    Exit Sub
Errs:
    Screen.MousePointer = vbDefault
    RaiseEvent DownloadComplete(Identifier, eCONNECTERROR)
    Resume Errs_Exit
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
    With UserControl
        .Width = 415
        .Height = 415
    End With
End Sub

Private Sub UserControl_AsyncReadComplete(AsyncProp As AsyncProperty)
On Error GoTo Errs
Dim f()     As Byte
Dim fn      As Long
Dim sFile   As String
Dim s       As String
'Large file tests ran on a 1400 Athlon, Win2K SP4, 256K RAM With Cable Connection
'______________________________________________________________________________
'//Sep 10, 2004 - > 272M file download tested from Microsoft server    (IN IDE)
'               - Windows XP Service Pack 2 for IT Professionals and Developers
'               - Errored with 'Out of memory' at 'f = AsyncProp.Value'
'//Sep 11, 2004 - > 272M file download tested from Microsoft server    (IN IDE)
'               - Windows XP Service Pack 2 for IT Professionals and Developers
'               - Errored with 'Out of memory' at 'f = AsyncProp.Value'
'//Sep 11, 2004 - > 73M file download tested from Microsoft server     (IN IDE)
'               - Office 2003 Service Pack 1
'               - Processed flawlessly
'//Sep 11, 2004 - > 132M file download tested from Microsoft server    (IN IDE)
'               - Windows 2000 Service Pack 4
'               - Processed flawlessly
'//Sep 11, 2004 - > 272M file download tested from Microsoft server  (COMPILED)
'               - Windows XP Service Pack 2 for IT Professionals and Developers
'               - Processed but wasn't pretty, Windows adjusted page file size
'
'//All other tests with smaller files ( < 150Meg ) on multiple systems error free

    With AsyncProp
        s = .PropertyName '----------------------------------------- Get Identifier
        If s = "0" Then
            sFile = sTEMPDIR & "\version.rus" '--------------------- Assign download location for script file
        ElseIf s = "9999" Then
             sFile = sTEMPDIR & "\Notify.ico" '--------------------- Assign download location for notify icon
        Else
            sFile = FileList.TempPath(CLng(s)) '-------------------- Assign download location for update files
        End If
        If .BytesMax <> 0 Then
            If .BytesMax <> .BytesRead Then
                On Error Resume Next
                UserControl.CancelAsyncRead s
                RaiseEvent DownloadComplete(s, eTRANSFERERROR) '---- Lost connection to download
            Else
                '//Call to DoEvents updates transfer status on frmMain
                '..for large files before file processing begins.
                DoEvents
                fn = FreeFile
                f = .Value
                Open sFile For Binary Access Write As #fn '--------- Open file for writing
                Put #fn, , f '-------------------------------------- Put downloaded data to file
                Close #fn
                Erase f '------------------------------------------- Purge Array
                fn = FileLen(sFile)
                If Dir$(sFile, 39) <> "" Then '--------------------- Verify write completed
                    RaiseEvent DownloadComplete(s, eSUCCESS)
                Else
                    On Error Resume Next
                    UserControl.CancelAsyncRead s
                    RaiseEvent DownloadComplete(s, eWRITEERROR) '--- Completed but failed to write to disk
                End If
            End If
        Else
            On Error Resume Next '---------------------------------- ReVive never connected to the file or server
            UserControl.CancelAsyncRead s
            RaiseEvent DownloadComplete(s, eCONNECTERROR)
        End If
    End With
Errs_Exit:
    Exit Sub
Errs:
    If Err.Number = 7 Then '7 = Out of memory
        MsgBox "Windows in running critically low on memory. LiveUpdate will now terminate.  ", vbCritical, "System Error"
    End If
    RaiseEvent DownloadComplete(s, eCONNECTERROR)
    Resume Errs_Exit
End Sub

Private Sub UserControl_AsyncReadProgress(AsyncProp As AsyncProperty)
On Error Resume Next
Static lBytesRead   As Long '------------------------ Stores total bytes in this downlaod for next lBytesRecv calculation
Static lBytesRecv   As Long '------------------------ Total bytes received this pass
    With AsyncProp
        If lBytesRead > .BytesRead Then '------------ Reset statics for new download
            lBytesRead = 0: lBytesRecv = 0
        End If
        If .BytesMax <> 0 Then
            lBytesRecv = .BytesRead - lBytesRead
            lBytesRead = .BytesRead
            If Setup.RunMode = eNORMAL Then Screen.MousePointer = vbHourglass '-- Putting this here avoids it everywhere else
            RaiseEvent DownloadProgress(.PropertyName, lBytesRecv, CLng(.BytesRead), CLng(.BytesMax))
        End If
    End With
End Sub

Public Property Get dAppShortName() As String
Attribute dAppShortName.VB_Description = "Short name of application to update. (Value overwritten when a 'lusetup.ini' is assigned)"
    dAppShortName = sAppShortName
End Property

Public Property Let dAppShortName(sShortName As String)
    If Len(sShortName) > 15 Then
        MsgBox "The AppShortName value cannot exceed 15 characters.  ", vbExclamation, "Invalid Value"
    Else
        sAppShortName = sShortName
        PropertyChanged "dAppShortName"
    End If
End Property

Public Property Get dAppLongName() As String
Attribute dAppLongName.VB_Description = "Long name of application to update. (Value overwritten when a 'lusetup.ini' is assigned)"
    dAppLongName = sAppLongName
End Property

Public Property Let dAppLongName(sLongName As String)
    sAppLongName = sLongName
    PropertyChanged "dAppLongName"
End Property

Public Property Get dScriptURLPrim() As String
Attribute dScriptURLPrim.VB_Description = "Primary update script URL. (Value overwritten when a 'lusetup.ini' is assigned)"
    dScriptURLPrim = sScriptURLPrim
End Property

Public Property Let dScriptURLPrim(sFile As String)
    sScriptURLPrim = sFile
    PropertyChanged "dScriptURLPrim"
End Property

Public Property Get dScriptURLAlt() As String
Attribute dScriptURLAlt.VB_Description = "Alternate update script URL. This script URL is attempted if dScriptURLPrim fails to connect within dTimeOutConnection. (Value overwritten when a 'lusetup.ini' is assigned)"
    dScriptURLAlt = sScriptURLAlt
End Property

Public Property Let dScriptURLAlt(sFile As String)
    sScriptURLAlt = sFile
    PropertyChanged "dScriptURLAlt"
End Property

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        .WriteProperty "dAppShortName", sAppShortName, ""
        .WriteProperty "dAppLongName", sAppLongName, ""
        .WriteProperty "dScriptURLPrim", sScriptURLPrim, ""
        .WriteProperty "dScriptURLAlt", sScriptURLAlt, ""
    End With
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
    On Error Resume Next
    With PropBag
        sScriptURLPrim = .ReadProperty("dScriptURLPrim", "")
        sScriptURLAlt = .ReadProperty("dScriptURLAlt", "")
        sAppShortName = .ReadProperty("dAppShortName", "")
        sAppLongName = .ReadProperty("dAppLongName", "")
    End With
End Sub

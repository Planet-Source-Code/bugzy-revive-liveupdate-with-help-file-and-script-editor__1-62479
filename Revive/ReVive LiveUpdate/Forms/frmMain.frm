VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMain 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EBEDED&
   Caption         =   "LiveUpdate"
   ClientHeight    =   5010
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7605
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   8.25
      Charset         =   0
      Weight          =   700
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   334
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   507
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin MSComctlLib.ImageList iml 
      Left            =   0
      Top             =   4440
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
   End
   Begin ReVive.ccXPButton cmdReport 
      Height          =   375
      Left            =   3600
      TabIndex        =   2
      Top             =   4440
      Visible         =   0   'False
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "&Details"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   0   'False
   End
   Begin ReVive.ccDownload ucDL 
      Left            =   600
      Top             =   4590
      _ExtentX        =   741
      _ExtentY        =   741
      dAppShortName   =   "Sample Project"
      dAppLongName    =   "ReVive PSC Demonstration Project"
      dScriptURLPrim  =   "http://members.cox.net/software.updates/PSC/PSCSample.rus"
      dScriptURLAlt   =   "http://members.cox.net/software.updates/PSC/PSCSample.rus"
   End
   Begin ReVive.ccXPButton cmdNext 
      Default         =   -1  'True
      Height          =   375
      Left            =   6180
      TabIndex        =   0
      Top             =   4440
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "&Next"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   0   'False
   End
   Begin ReVive.ccXPButton cmdCancel 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4890
      TabIndex        =   1
      Top             =   4440
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "&Cancel"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      FocusRect       =   0   'False
   End
   Begin VB.Frame frm 
      BackColor       =   &H00EBEDED&
      Caption         =   "Welcome to LiveUpdate"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Index           =   0
      Left            =   1800
      TabIndex        =   3
      Top             =   510
      Width           =   5565
      Begin VB.Label lblLastUpdated 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EBEDED&
         Caption         =   "Last Updated: Unknown"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   2940
         TabIndex        =   13
         Top             =   3390
         Width           =   2355
      End
      Begin VB.Label lblOpening 
         BackColor       =   &H00EBEDED&
         Caption         =   "OPENING MESSAGE"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2805
         Left            =   330
         TabIndex        =   4
         Top             =   450
         Width           =   4890
      End
   End
   Begin VB.Frame frmError 
      BackColor       =   &H00EBEDED&
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   1800
      TabIndex        =   8
      Top             =   510
      Visible         =   0   'False
      Width           =   5565
      Begin VB.Label lblErrTitle 
         BackStyle       =   0  'Transparent
         Caption         =   "Error Title"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   960
         TabIndex        =   12
         Top             =   510
         Width           =   4215
      End
      Begin VB.Image imgWarnIcon 
         Height          =   480
         Left            =   330
         Top             =   390
         Width           =   480
      End
      Begin VB.Label lblErrExplain 
         BackColor       =   &H00EBEDED&
         Caption         =   "Error Explanation"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   2115
         Left            =   360
         TabIndex        =   10
         Top             =   1080
         Width           =   4815
      End
      Begin VB.Label Label1 
         BackColor       =   &H00EBEDED&
         Caption         =   "Click E&xit to close LiveUpdate."
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   330
         TabIndex        =   9
         Top             =   3390
         Width           =   2385
      End
   End
   Begin VB.Frame frmList 
      BackColor       =   &H00EBEDED&
      Caption         =   "The following components were checked for updates. "
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3765
      Left            =   1800
      TabIndex        =   5
      Top             =   510
      Visible         =   0   'False
      Width           =   5580
      Begin MSComctlLib.ProgressBar pbDownload 
         Height          =   225
         Left            =   230
         TabIndex        =   14
         Top             =   3390
         Visible         =   0   'False
         Width           =   5110
         _ExtentX        =   9022
         _ExtentY        =   397
         _Version        =   393216
         BorderStyle     =   1
         Appearance      =   0
         Scrolling       =   1
      End
      Begin MSComctlLib.ListView lvFiles 
         Height          =   2805
         Left            =   270
         TabIndex        =   6
         TabStop         =   0   'False
         Top             =   390
         Width           =   5030
         _ExtentX        =   8864
         _ExtentY        =   4948
         View            =   3
         LabelEdit       =   1
         LabelWrap       =   0   'False
         HideSelection   =   -1  'True
         HideColumnHeaders=   -1  'True
         FullRowSelect   =   -1  'True
         _Version        =   393217
         SmallIcons      =   "iml"
         ForeColor       =   -2147483640
         BackColor       =   16777215
         Appearance      =   0
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "File Description"
            Object.Width           =   5679
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Alignment       =   1
            SubItemIndex    =   1
            Text            =   "Status"
            Object.Width           =   3158
         EndProperty
      End
      Begin VB.Shape Shape1 
         BackColor       =   &H00C00000&
         BorderColor     =   &H008F837A&
         FillColor       =   &H00C00000&
         Height          =   2895
         Left            =   225
         Top             =   345
         Width           =   5115
      End
      Begin VB.Label lblTotal 
         Alignment       =   1  'Right Justify
         BackColor       =   &H00EBEDED&
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   225
         Left            =   3475
         TabIndex        =   11
         Top             =   3390
         Width           =   1845
      End
      Begin VB.Label lblContinue 
         AutoSize        =   -1  'True
         BackColor       =   &H00EBEDED&
         Caption         =   "Label1"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   195
         Left            =   230
         TabIndex        =   7
         Top             =   3390
         Width           =   480
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'Chris Cochran          cwc.software@gmail.com        Updated: 13 Sep 05
'***********************************************************************
'------------------------------------------------------------------------------------------
' Module    : frmMain
' Purpose   : GUI for navigating LiveUpdate sequence.
' Image     : The frmMain picture is the work of Kornél Pál. Thanks Kornél for offering it.
'------------------------------------------------------------------------------------------
Option Explicit

'//Icon declares for displaying in ListView
Private Type TypeIcon
    cbSize                      As Long
    picType                     As PictureTypeConstants
    hIcon                       As Long
End Type
Private Type CLSID
    id(16)                      As Byte
End Type
Private Type SHFILEINFO
    hIcon                       As Long
    iIcon                       As Long
    dwAttributes                As Long
    szDisplayName               As String * MAX_PATH
    szTypeName                  As String * 80
End Type
Private Const SHGFI_ICON        As Long = &H100
Private Const SHGFI_SMALLICON   As Long = &H1
Private Declare Function OleCreatePictureIndirect Lib "oleaut32.dll" (pDicDesc As TypeIcon, riid As CLSID, ByVal fown As Long, lpUnk As Object) As Long
Private Declare Function SHGetFileInfo Lib "shell32.dll" Alias "SHGetFileInfoA" (ByVal pszPath As String, ByVal dwFileAttributes As Long, psfi As SHFILEINFO, ByVal cbFileInfo As Long, ByVal uFlags As Long) As Long

'//Errors encountered during update process
Private Enum eluErrors
    eluSCRIPTDLERR = 0 '------------------------- Error downloading update script
    eluSCRIPTPROCERR = 1 '----------------------- Error processing update script
    eluPERMISSIONERR = 2 '----------------------- Insufficient privilege to run LiveUpdate
    eluFILEDOWNLOADERR = 3 '--------------------- Error downloading updated file
    eluFILEPROCESSERR = 4 '---------------------- Error processing update file after download
    eluSCRIPTEMPTY = 5 '------------------------- Update script did not contain any entries
End Enum

Private Const oleLiteGray As Long = 16054263 '--- Alternating listview color (See AltLVBackColors sub)
Private Const CRYPTPWD  As String = "ReVive" '--- Password for encrypting web URL's in .ris files
'>*>*>*>*>*>*>*>*>*>*>*>*>*>*>*>*>*>*>*>*>*>*'... ***ONCE ESTABLISHED FOR AN APP DO NOT CHANGE***

Private bStep           As Byte '---------------- Sets up Next button for next click
Private lTotalDlSize    As Long '---------------- The combined size of all needed download files
Private lRecvdBytes     As Long '---------------- Stores total bytes received for all update downloads
Private bWriteScript    As Boolean '------------- True when script is read successfully, when true a new local script is written
Private bJustExit       As Boolean '------------- True when form should not prompt to unload
Private bAppKilled      As Boolean '------------- True when running app is killed by ReVive
Private bUpdateIcons    As Boolean '------------- True when icons need updated after succesful update

Private Sub cmdReport_Click()
    Call CreateReport '-------------------------- Create and display HTML update report
End Sub

Private Sub cmdCancel_Click()
On Error Resume Next
    Select Case cmdCancel.Caption
        Case "&Cancel"
            Unload Me '-------------------------- Exit with Yes/No option
        Case "Restart &Later"
            bREBOOT = False
            bJustExit = True
            Unload Me '-------------------------- Exit without Yes/No option
    End Select
End Sub

Private Sub cmdNext_Click()
'-----------------------------------------------------------------------------------------
' Purpose   : Ground zero for app navigation.
'-----------------------------------------------------------------------------------------
Dim x       As Long
    Select Case bStep
        Case 0 '--------------------------------- Download web update script
            Me.cmdNext.Enabled = False
            Me.ucDL.Download Setup.ScriptURLPrim, CStr(0)
            
        Case 1 '--------------------------------- Prepair screen and begin to cycle downloads
            Me.cmdNext.Enabled = False
            Me.cmdReport.Visible = False
            Me.frmList.Caption = "Downloading and installing available updates... "
            Me.pbDownload.Visible = True
            With FileList '---------------------- Display file sizes and mark them as queued in the listview
                For x = 1 To UBound(.Description)
                    If .Status(x) = UPDATEREQ Then
                        Me.lvFiles.ListItems(x).SubItems(1) = IIf(.FileSize(x) > 1000, _
                            Format(.FileSize(x) / 1000, "#,##0") & "K Queued", Format(.FileSize(x), _
                            "##0") & "Byes Queued")
                    End If
                Next x
            End With
            Call SequenceDownloads '------------- Begin downloading
        Case 2 '--------------------------------- Exit Revive
            Unload Me
    End Select
End Sub

Private Sub ucDL_DownloadProgress(Identifier As String, RecvdBytes As Long, CurBytes As Long, MaxBytes As Long)
'-------------------------------------------------------------------------------
' Purpose   : Keeps frmMain file download progress in sync as downloads progress
'-------------------------------------------------------------------------------
On Error Resume Next
Dim x               As Long
Static ident        As String
    x = CLng(Identifier) '----------------------- Get download identifier
    If x <> 0 And x <> 9999 Then '--------------- Skip if 0 (update script) or 9999 (notification icon)
        With FileList
            If ident <> Identifier Then
                '****************************************************************************
                'This is solely to ensure filesize is correct on script to avoid generating
                'invalid property value error updating progress bar. It also ensures accuracy
                'of download progression if script filesize is incorrect. DO NOT REMOVE.
                If .FileSize(x) <> MaxBytes Then
                    .FileSize(x) = MaxBytes
                    Call ResetProgressBar
                End If
                '****************************************************************************
                ident = CStr(Identifier) '------- Remember identifier to skip above next pass
            End If
            lRecvdBytes = lRecvdBytes + RecvdBytes
            Me.pbDownload.Value = (lRecvdBytes / lTotalDlSize) * 100
            Me.lvFiles.ListItems(x).EnsureVisible
            If MaxBytes = CurBytes Then
                '//Only display message if file is greater than 5Meg (smaller files write to quick)
                If MaxBytes > 5000000 Then Me.lvFiles.ListItems(x).SubItems(1) = "Writing to Disk"
            Else
                '******SELECT 1 OF 3 DIFFERENT FILE DOWNLOAD PROGRESS DISPLAY OPTIONS******
                '//Show progress by bytes counting down
                'Me.lvFiles.ListItems(x).SubItems(1) = IIf((MaxBytes - CurBytes) > 1000, _
                    Format((MaxBytes - CurBytes) / 1000, "#,##0K"), Format(MaxBytes - CurBytes, _
                    "#,##0 Bytes"))
    
                '//Show progress by % complete
                'Me.lvFiles.ListItems(x).SubItems(1) = Format(((CurBytes / MaxBytes) * 100), _
                    "#") & "% Complete"
    
                '//Show progress by bytes counting down and % complete
                Me.lvFiles.ListItems(x).SubItems(1) = IIf((MaxBytes - CurBytes) > 1000, _
                    Format((MaxBytes - CurBytes) / 1000, "#,##0K"), Format(MaxBytes - CurBytes, _
                    "#,##0 Bytes")) & "       " & Format((CurBytes / MaxBytes) * 100, _
                    IIf(((CurBytes / MaxBytes) * 100) < 10, "  0", "00")) & "%"
                '**************************************************************************
            End If
        End With
    End If
End Sub

Private Sub ucDL_DownloadComplete(Identifier As String, Result As eDownloadResults)
'--------------------------------------------------------------------------------------
' Purpose   : Called when any download completes, good or bad, from ucDownload control.
'             This function takes appropriate action based on download result.
'--------------------------------------------------------------------------------------
Dim lResult             As Byte
Dim x                   As Long
Static bAltAttempted    As Boolean
    x = CLng(Identifier)
    If x = 0 Then '-------------- 0 is our update script file
        Select Case Result
            Case eSUCCESS
                lResult = ParseUpdateScript(sTEMPDIR & "\version.rus")
                '//Exit if script parsing fails and we are in auto or notify mode
                If lResult <> 0 And Setup.RunMode <> eNORMAL Then GoTo UnloadNow
                '//Display error for Normal mode if lResult > 0
                If lResult = 1 Then
                    Call ShowLiveUpdateError(eluPERMISSIONERR)
                ElseIf lResult = 2 Then
                    Call ShowLiveUpdateError(eluSCRIPTPROCERR)
                ElseIf lResult = 3 Then
                    Call ShowLiveUpdateError(eluSCRIPTEMPTY)
                Else
                    bWriteScript = True '------ Script read, set to write local scipt on exit
                    '//Download notify icon if available.
                    'NOTE: This was the only file (for whatever reason) that required
                    '      the same filename case in both the script and web file. Wiered...
                    If Len(Setup.NotifyIcon) And Setup.RunMode <> eNORMAL Then
                        Me.ucDL.Download Setup.NotifyIcon, "9999"
                    Else
                        '//Show file, icons, and their update status in the listview
                        Call ListScanResults
                    End If
                    Exit Sub
                End If
            Case Else
                '//Script download failed. Try alternate if available.
                If Not bAltAttempted And Len(Setup.ScriptURLAlt) > 0 Then
                    bAltAttempted = True
                    Me.ucDL.Download Setup.ScriptURLAlt, CStr(0)
                Else
                    '//Exit if script download fails and we are
                    '..in auto or notify mode
                    If Setup.RunMode <> eNORMAL Then GoTo UnloadNow
                    Call ShowLiveUpdateError(eluSCRIPTDLERR)
                End If
        End Select
        Exit Sub
    ElseIf x = 9999 Then '---------------------- 9999 is our notification icon
        Call ListScanResults '------------------ Press-on with success or failure
        Exit Sub
    Else '-------------------------------------- Else represents all update files
        '//First update the listview icon if required
        If bUpdateIcons Then Call UpdateIcon(CInt(Identifier))
        With FileList
            Select Case Result
                Case eSUCCESS
                    .Status(x) = DOWNLOADED
                    Me.lvFiles.ListItems(x).SubItems(1) = "Update Queued"
                Case eCONNECTERROR
                    Me.lvFiles.ListItems(x).SubItems(1) = "Connection Error"
                    If Setup.RunMode <> eNORMAL Then
                        If .MustUpdate(x) Then
                            GoTo UnloadNow
                        Else
                            .Status(x) = ERRCONNECTING
                            Call ResetProgressBar '------ Adjust total download bytes for progress bar
                        End If
                    Else
                        '//Advise user of error and request a course of action
                        Select Case frmDLError.RequestAction(CStr(x), CONNECTERROR)
                            Case 0: '-------------------- Retry
                                .Status(x) = UPDATEREQ
                                Call ResetProgressBar '-- Adjust total download bytes for progress bar
                            Case 1: '-------------------- Continue
                                .Status(x) = ERRCONNECTING
                                With Me.lvFiles.ListItems(x)
                                    .ForeColor = 255
                                    .ListSubItems(1).ForeColor = 255
                                End With
                                Me.lvFiles.Refresh '----- Repaint lvFiles to reflect red text (found this was needed)
                                Call ResetProgressBar '-- Adjust total download bytes for progress bar
                            Case 2: '-------------------- Abort
                                .Status(x) = ERRCONNECTING
                                Call ShowLiveUpdateError(eluFILEDOWNLOADERR)
                                Exit Sub
                        End Select
                    End If
                    
                Case eTRANSFERERROR, eWRITEERROR
                    Me.lvFiles.ListItems(x).SubItems(1) = "Transfer Error"
                    If Setup.RunMode <> eNORMAL Then
                        If .MustUpdate(x) Then
                            GoTo UnloadNow
                        Else
                            .Status(x) = ERRTRANSFERRING
                            Call ResetProgressBar '-- Adjust total download bytes for progress bar
                        End If
                    Else
                        '//Advise user of error and request a course of action
                        Select Case frmDLError.RequestAction(CStr(x), TRANSFERERROR)
                            Case 0: '-------------------- Retry
                                FileList.Status(x) = UPDATEREQ
                                Call ResetProgressBar '-- Adjust total download bytes for progress bar
                            Case 1: '--------------------Abort
                                .Status(x) = ERRTRANSFERRING
                                Call ShowLiveUpdateError(eluFILEDOWNLOADERR)
                                Exit Sub
                            Case 2: '-------------------- Continue
                                .Status(x) = ERRTRANSFERRING
                                With Me.lvFiles.ListItems(x)
                                    .ForeColor = 255
                                    .ListSubItems(1).ForeColor = 255
                                End With
                                Me.lvFiles.Refresh '----- Repaint lvFiles to reflect red text (found this was needed)
                                Call ResetProgressBar '-- Adjust total download bytes for progress bar
                        End Select
                    End If
            End Select
        End With
    End If
    Call SequenceDownloads '----------------------------- Check for more downloads
    Exit Sub
UnloadNow:
        bJustExit = True
        Unload Me
End Sub

Private Sub ListScanResults()
'-----------------------------------------------------------------------------------------
' Purpose   : Post update list and prepare frmMain for downloading files when available.
'             Only called from ucDL_DownloadComplete when the update script is downloaded.
'-----------------------------------------------------------------------------------------
On Error GoTo Errs
Dim x               As Integer
Dim iCurrents       As Integer
Dim bProceed        As Boolean
Dim bAbort          As Boolean
    With FileList
        For x = 1 To UBound(.Description)
            '//Get an icon, hopefully from an existing file, and load into the imagelist
            If Setup.ShowFileIcons Then
                Me.iml.ListImages.Add x, x & "|", GetIcon(.InstallPath(x), SHGFI_SMALLICON)
                Me.lvFiles.ListItems.Add x, .Description(x), .Description(x), , x
            Else
                Me.lvFiles.ListItems.Add x, .Description(x), .Description(x)
            End If
            Me.lvFiles.ListItems(x).ToolTipText = .Description(x)
            If .Status(x) = UPDATEREQ Then
                lTotalDlSize = lTotalDlSize + .FileSize(x)
                Me.lvFiles.ListItems(x).SubItems(1) = "Update Available"
                bProceed = True
            ElseIf .Status(x) = FILENOTINSTALLED Then
                With Me.lvFiles.ListItems(x)
                    .SubItems(1) = "Installation Not Found"
                    .ForeColor = 255
                    .ListSubItems(1).ForeColor = 255
                End With
                If .MustUpdate(x) Then bAbort = True
            Else
                Me.lvFiles.ListItems(x).SubItems(1) = "Installation Current"
                iCurrents = iCurrents + 1
            End If
        Next x
    End With
    '//We are here because we have successfully downloaded the script and scanned for updates.
    '..All user privileges are good and files exist on client that must to perform updates.
    Setup.LastChecked = Date '--------------------------------- Update the "Last Checked" date.
    Call AltLVBackground(Me.lvFiles, vbWhite, oleLiteGray) '--- Draw alternating lvList colors
    If bAbort Then
        '//At least one required file was not found on client
        If Setup.RunMode <> eNORMAL Then GoTo UnloadNow '------ Bolt in NOTIFY and AUTO modes
        With Me.cmdReport
            .Caption = "&Report"
            .Move Me.cmdCancel.Left, Me.cmdCancel.Top
            .Visible = True
        End With
        Me.cmdCancel.TabStop = False
        Me.lblTotal.Visible = False
        Me.lblContinue.Caption = "Select Report for details or Exit to close LiveUpdate."
        With Me.frmList
            .Caption = "LiveUpdate failed to download any available updates. "
            .ZOrder
            .Visible = True
        End With
        With Me.cmdNext
            .Caption = "&Exit"
            .Enabled = True
            .SetFocus
        End With
        bJustExit = True '------------------------ Skip Yes/No exit dialog
        bStep = 2
    Else
        If bProceed Then
            '//There are updates available
            Me.lblTotal.Caption = "Total Download: " & IIf(lTotalDlSize > 1000, Format(lTotalDlSize / 1000, "#,##0K"), Format(lTotalDlSize, "##0 Bytes"))
            With Me.cmdNext
                .Enabled = True
                If Me.Visible Then .SetFocus
            End With
            Me.lblContinue.Caption = "Select Next to download and install updates..."
            With Me.cmdReport
                .Visible = True
            End With
            With Me.frmList
                .ZOrder
                .Visible = True
            End With
            bStep = 1
            If Setup.RunMode = eNOTIFY Then
                Call frmNotify.Notify
            ElseIf Setup.RunMode = eAUTO Then
                Call cmdNext_Click
            End If
        Else
            '//There are no updates available
            If Setup.RunMode <> eNORMAL Then GoTo UnloadNow
            With Me.cmdReport
                .Caption = "&Report"
                .Move Me.cmdCancel.Left, Me.cmdCancel.Top
                .Visible = True
            End With
            Me.cmdCancel.TabStop = False
            With Me.cmdNext
                .Caption = "&Exit"
                .Enabled = True
                .SetFocus
            End With
            Me.lblTotal.Visible = False
            Me.lblContinue.Caption = "Select Report for details or Exit to close LiveUpdate."
            With Me.frmList
                If iCurrents = UBound(FileList.Description) Then '-- See if all files were current
                    .Caption = "There are no updates available at this time. "
                Else
                    .Caption = "LiveUpdate failed to download any available updates. "
                End If
                .ZOrder
                .Visible = True
            End With
            bJustExit = True '----------------------- Skip Yes/No exit dialog
            bStep = 2
        End If
    End If
Errs_Exit:
    Screen.MousePointer = vbDefault
    Exit Sub
Errs:
    If Setup.RunMode = eNORMAL Then '---------------- Exit app if not in Normal mode
        Call ShowLiveUpdateError(eluFILEPROCESSERR)
        Resume Errs_Exit
    End If
UnloadNow:
    bJustExit = True
    Screen.MousePointer = vbDefault
    Unload Me
End Sub

Private Sub SequenceDownloads()
'-----------------------------------------------------------------------------------
' Purpose   : Recursively called from ucDL_DownloadComplete when update files
'             finish downloading (Good or Bad). Update file downloads are started
'             here, and LiveUpdate continues when there are no more UPDATEREQ flags.
'-----------------------------------------------------------------------------------
Dim x   As Long
Dim y   As Long
    With FileList '------------------------ See if there is anything left to download
        y = UBound(.Status)
        For x = 1 To y
            If .Status(x) = UPDATEREQ Then
                .Status(x) = DOWNLOADING
                Me.lvFiles.ListItems(x).SubItems(1) = "Connecting"
                Me.ucDL.Download FileList.DownloadURL(x), CStr(x)
                Exit Sub '----------------- Bolt if we begin a download
            End If
        Next x
        Me.lvFiles.ListItems(1).EnsureVisible
        '//All downloads have now been attempted. Now see if any are ready to install.
        '..If none are ready it is because they failed to download successfully.
        For x = 1 To y
            If .Status(x) = DOWNLOADED Then Exit For
        Next x
    End With
    If x <= y Then
        '//Atleast one file is UPDATEREADY so lets begin installing
        Me.lblTotal.Visible = False
        Me.pbDownload.Visible = False
        Call InstallUpdates '---------------- Begin installing updates
    Else
        '//There are no updates ready to install so prepare frmMain and standby to exit
        If Setup.RunMode <> eNORMAL Then GoTo UnloadNow '-- Bolt if in auto or notify mode
        Me.frmList.Caption = "There are no updates Ready To Install at this time..."
        Me.lblContinue.Caption = "Select Exit to close LiveUpdate."
        With Me.cmdNext
            .Caption = "&Exit"
            .Enabled = True
        End With
        With Me.cmdReport
            .Caption = "&Report"
            .Move Me.cmdCancel.Left, Me.cmdCancel.Top
        End With
        Me.cmdCancel.TabStop = False
        bJustExit = True '------------------- Skip Yes/No exit dialog
        bStep = 2
        Me.lblTotal.Visible = False
        Me.pbDownload.Visible = False
    End If
Errs:
    Screen.MousePointer = vbDefault
    Exit Sub
UnloadNow:
    bJustExit = True
    Unload Me
End Sub

Private Sub InstallUpdates()
'---------------------------------------------------------------------------------------
' Purpose   : Check for the running app and begin installing updates if they all pass
'             the VerifyInstallReady test. Only called from the SequenceDownloads sub.
'---------------------------------------------------------------------------------------
On Error GoTo Errs
Dim x       As Long
Dim bErrors As Boolean
Dim lResult As eupdResults
    With Setup '-------------------- Check to see if program is running and act accordingly.
        If Len(.UpdateAppTitle) And .RunMode = eNORMAL Then
            If IsAppRunning(.UpdateAppTitle, .UpdateAppClass) Then
                If .UpdateAppKill Then
                    Call KillApp '------------------------ Kill app without notification
                Else
                    Call frmCloseApp.WarnAppIsRunning '--- Warn user and recommend killing app
                End If
                If Len(Setup.LaunchIfKilled) Then '------- Mark if app was killed for restarting when complete
                    Erase tRunningApps.lWndHwnd: Erase tRunningApps.lWndProcessID
                    bAppKilled = Not IsAppRunning(.UpdateAppTitle, .UpdateAppClass)
                    Erase tRunningApps.lWndHwnd: Erase tRunningApps.lWndProcessID
                End If
            End If
        End If
    End With
    If VerifyInstallReady Then '------ Make sure *ALL* updates pass test before installing any.
        With FileList '--------------- Atleast one update is ready and none caused an abort.
            For x = 1 To UBound(.Description)
                If .Status(x) = UPDATEREADY Then
                    lResult = UpdateFile(.TempPath(x), .InstallPath(x)) '- Initiate update
                    Select Case lResult
                        Case 0 '------- Success
                            .Status(x) = UPDATECOMP
                            Me.lvFiles.ListItems(x).ListSubItems(1) = "Update Complete"
                            If Len(.UpdateMessage(x)) Then '-------------- Store Update Message
                                sUpdateMessage = sUpdateMessage & .Description(x) & " - " & .UpdateMessage(x) & vbNewLine & vbNewLine
                            End If
                        Case 1 '------- Success, reboot required
                            .Status(x) = UPDATECOMPREBOOT
                            Me.lvFiles.ListItems(x).ListSubItems(1) = "Reboot Required"
                            If Len(.UpdateMessage(x)) Then '-------------- Store Update Message
                                sUpdateMessage = sUpdateMessage & .Description(x) & " - " & .UpdateMessage(x) & vbNewLine & vbNewLine
                            End If
                            bREBOOT = True
                        Case 4 '------- Insufficient privilege
                            .Status(x) = INSUFFPRIVILEGE
                            Me.lvFiles.ListItems(x).ListSubItems(1) = "Permission Denied"
                        Case Else '---- Misc errors
                            .Status(x) = ERRUPDATING
                            Me.lvFiles.ListItems(x).ListSubItems(1) = "Update Failed"
                    End Select
                End If
                '//All updates have now been attempted, let's display results...
                If Not bErrors Then '-- Prepare to display caption when done
                    If .Status(x) = ERRCONNECTING Or _
                            .Status(x) = ERRTRANSFERRING Or _
                            .Status(x) = ERRUPDATING Or _
                            .Status(x) = FILENOTINSTALLED Or _
                            .Status(x) = INSUFFPRIVILEGE Then
                       bErrors = True
                    End If
                End If
            Next x
        End With
        If bErrors Then
            Me.frmList.Caption = "LiveUpdate completed but experienced errors."
        Else
            Me.frmList.Caption = "LiveUpdate completed successfully."
        End If
        If bREBOOT Then
            Me.cmdNext.Caption = "Restart &Now"
            With Me.cmdCancel
                If Setup.ForceReboots = 0 Then
                    .Caption = "Restart &Later"
                    .Visible = True
                Else
                    .Visible = False
                End If
            End With
            Me.cmdReport.Caption = "&Report"
            Me.lblContinue.Caption = "Select Report for details or Restart Now to complete LiveUpdate."
        Else
            Me.cmdNext.Caption = "&Exit"
            With Me.cmdReport
                .Caption = "&Report"
                .Move Me.cmdCancel.Left, Me.cmdCancel.Top
                .Visible = True
            End With
            Me.cmdCancel.TabStop = False
            Me.lblContinue.Caption = "Select Report for details or Exit to close LiveUpdate."
        End If
        With Me.cmdNext
            .Enabled = True
            If Me.Visible Then .SetFocus
        End With
        If Setup.RunMode = eAUTO Then '--- Popup updates completed notify form above taskbar
            frmNotify.Notify True
        End If
        bJustExit = True '---------------- Skip Yes/No exit dialog
        bStep = 2
    End If
    If bAppKilled And Not bREBOOT Then '-- Launch killed app now if requested
        If Len(Setup.LaunchIfKilled) Then
            If Setup.RunMode = eNORMAL Then Screen.MousePointer = vbHourglass
            Me.cmdNext.Enabled = False '-- Disable Exit until timer expires
            ShellExecute 0&, "open", Setup.LaunchIfKilled, vbNullString, vbNullString, SW_NORMAL
            Sleep 2000 '------------------ Allow the restarted app time to load before shifting focus to ReVive
            With Me.cmdNext
                .Enabled = True
                .SetFocus
            End With
        End If
    End If
Errs_Exit:
    SetForegroundWindow Me.hWnd '--------- Return ReVive to the foremost window
    Screen.MousePointer = vbDefault
    Exit Sub
Errs:
    bJustExit = True
    Unload Me
End Sub

Private Function VerifyInstallReady() As Boolean
'--------------------------------------------------------------------------------------------
' Purpose   : Validates files are ready for install by testing for obvious show-stoppers
'             that often occur during the update process. Notice all files are tested before
'             ReVive installs any one update. This ensures integrity of your MustUpdate flags.
'             Due to the nature of this app, thorough update testing is the highest priority.
'             Only called from the InstallUpdates sub just prior to updating files.
'
' Returns   : True = Atleast one file can be updated and no MustUpdate files caused an abort.
'             False = Atleast one file caused an abort or no files passed update test.
'
' Note      : Updating the listview seems uneccessary here, and will be if all files pass
'             and the updates follow (because it will happen so fast). BUT...if no files
'             cause an abort but none remain updatable, the listview will provide why.
'--------------------------------------------------------------------------------------------
On Error GoTo Errs
Dim x           As Long
Dim bContinue   As Boolean
    With FileList
        For x = 1 To UBound(.Description)
            If .Status(x) = DOWNLOADED Then '------------- Only test successful downloads
                Select Case TestUpdateSuccess(.TempPath(x), .InstallPath(x))
                    Case eupdINSUFFPRIV
                        .Status(x) = INSUFFPRIVILEGE
                        If .MustUpdate(x) Then '---------- Required-abort
                            If Setup.RunMode <> eNORMAL Then GoTo UnloadNow
                            Call ShowLiveUpdateError(eluPERMISSIONERR)
                            Exit Function
                        Else '---------------------------- File not required - continue
                            With Me.lvFiles.ListItems(x)
                                .SubItems(1) = "Requires Admin Priv"
                                .ForeColor = vbRed
                                .ListSubItems(1).ForeColor = 255
                            End With
                            DoEvents '-------------------- Required to visually update ListView
                        End If
                    Case eupdDESTINVALID, eupdSOURCENOTFOUND, eupdUNKNOWNERR
                        If .MustUpdate(x) Then '---------- Required-abort
                            If Setup.RunMode <> eNORMAL Then GoTo UnloadNow
                            Call ShowLiveUpdateError(eluFILEPROCESSERR)
                            Exit Function
                        Else
                            .Status(x) = ERRUPDATING '---- File not required - continue
                            With Me.lvFiles.ListItems(x)
                                .SubItems(1) = "Update Test Failed"
                                .ForeColor = vbRed
                                .ListSubItems(1).ForeColor = 255
                            End With
                            DoEvents '-------------------- Required to visually update ListView
                        End If
                    Case Else
                        bContinue = True
                        .Status(x) = UPDATEREADY
                End Select
            End If
        Next x
    End With
    '//If there are any UPDATEREADY files continue, otherwise prepare to exit ReVive.
    If bContinue Then
        '//No single update caused an abort and atleast one file is ready to install.
        VerifyInstallReady = True
    Else
        '//There are no updates ready after tests above, but none caused an abort.
        If Setup.RunMode <> eNORMAL Then GoTo UnloadNow '-- Bolt if in auto or notify mode
        Me.frmList.Caption = "There are no updates Ready To Install at this time..."
        Me.lblContinue.Caption = "Select Exit to close LiveUpdate."
        Me.cmdCancel.Caption = "&Report"
        bJustExit = True '--------------------------------- Skip Yes/No exit dialog
        bStep = 2
        With Me.cmdNext
            .Caption = "&Exit"
            .Enabled = True
            If Me.Visible Then .SetFocus
        End With
    End If
Errs_Exit:
    Exit Function
Errs:
    If Setup.RunMode <> eNORMAL Then GoTo UnloadNow '------ Bolt if in Notify or Auto mode
    Call ShowLiveUpdateError(eluFILEPROCESSERR)
    Resume Errs_Exit
UnloadNow:
    bJustExit = True
    Unload Me
End Function

Private Sub ResetProgressBar()
'-----------------------------------------------------------------------
' Purpose   : Adjusts total update download size and bytes received
'             to reflect accurate progress after a file transfer fails
'             or incorrect file size is specified in web update script.
'             Called from ucDL_DownloadComplete or ucDL_DownloadProgress
'-----------------------------------------------------------------------
On Error Resume Next
Dim x As Long
    lTotalDlSize = 0
    lRecvdBytes = 0
    With FileList
        For x = 1 To UBound(.Description)
            If .Status(x) = DOWNLOADED Or .Status(x) = UPDATEREQ Or .Status(x) = DOWNLOADING Then
                lTotalDlSize = lTotalDlSize + .FileSize(x)
                If .Status(x) = DOWNLOADED Then
                    lRecvdBytes = lRecvdBytes + .FileSize(x)
                End If
            End If
        Next x
    End With
    Me.pbDownload.Value = (lRecvdBytes / lTotalDlSize) * 100 '--- Correct progress bar
End Sub

Private Sub ShowLiveUpdateError(ByVal ErrType As eluErrors)
'-------------------------------------------------------------------
' Purpose   : Central procedure to display LiveUpdate errors to user
'-------------------------------------------------------------------
Dim sT  As String
Dim sE  As String
    Set Me.imgWarnIcon.Picture = LoadResPicture(201, 1) '----- Load warning icon from resource file
    Select Case ErrType
        Case eluSCRIPTDLERR
            sT = "LiveUpdate was unable to download the update script."
            sE = LoadResString(400) '--- Connection error explanation
        Case eluSCRIPTPROCERR
            sT = "LiveUpdate was unable to process the update script."
            sE = LoadResString(401) '--- Script process error explanation
        Case eluPERMISSIONERR
            sT = "This update requires administrator privilege."
            sE = LoadResString(402) '--- Premission error explanation
        Case eluFILEDOWNLOADERR
            sT = "LiveUpdate could not download all required update files."
            sE = LoadResString(403) '--- Required file download error explanation
        Case eluFILEPROCESSERR
            sT = "LiveUpdate was unable to process the update files."
            sE = LoadResString(404) '--- File processing error explanation
        Case eluSCRIPTEMPTY
            sT = "LiveUpdate was unable to extract a valid update list."
            sE = LoadResString(405) '--- Script empty error explanation
    End Select
    Me.lblErrTitle.Caption = sT
    Me.lblErrExplain.Caption = sE
    Me.cmdCancel.Visible = False
    bJustExit = True '------------------ Exit without Yes/No dialog
    With Me.cmdNext
        .Caption = "&Exit"
        .Enabled = True
        .SetFocus
    End With
    With Me.frmError
        .ZOrder 0
        .Visible = True
    End With
    Screen.MousePointer = vbDefault
    bStep = 2
End Sub

Private Sub cmdNext_FormActivate(State As WindowState)
'-----------------------------------------------------------------------------------------
' Purpose   : Draw title bar and text Active and Deactive as form changes Windows foremost.
'             Sub is raised from cmbNext (XPButton) subclassing, "FormActivate" event.
'------------------------------------------------------------------------------------------
On Error Resume Next
    If Me.WindowState <> vbMinimized And Me.Visible Then
        Call DrawTitleBar(Me, State, Setup.AppShortName & " LiveUpdate", Not bREBOOT)
    End If
End Sub

Private Sub Form_Load()
Dim R   As RECT
Dim w   As Long
    w = Me.ScaleWidth
    With Setup '------------------------------- Get LiveUpdate settings from .ris file or ucDownload
        .AppShortName = Left$(ProfileGetItem("Setup", "AppShortName", Me.ucDL.dAppShortName, .SetupScriptPath), 15)
        .AppLongName = ProfileGetItem("Setup", "AppLongName", Me.ucDL.dAppLongName, .SetupScriptPath)
        .NotifyIcon = ProfileGetItem("Setup", "NotifyIcon", "", .SetupScriptPath)
        .ScriptURLPrim = DecryptString(ProfileGetItem("Setup", "ScriptURLPrim", "", .SetupScriptPath))
        .ScriptURLAlt = DecryptString(ProfileGetItem("Setup", "ScriptURLAlt", "", .SetupScriptPath))
        .LastChecked = ProfileGetItem("Setup", "LastChecked", "Unknown", .SetupScriptPath)
        If Len(.ScriptURLPrim) = 0 Then .ScriptURLPrim = Me.ucDL.dScriptURLPrim
        If Len(.ScriptURLAlt) = 0 Then .ScriptURLAlt = Me.ucDL.dScriptURLAlt
        If Len(.LastChecked) Then Me.lblLastUpdated.Caption = "Last Checked: " & .LastChecked
        '//Verify we have an app to update before continuing any further. This situation will only
        '..occur if the ccDownload values are not entered and the ris file cannot be found.
        '..To avoid this scenerio, provide App short and long names, and ScriptURLPrim values
        '..in ucDownload control on frmMain.
        If Len(.AppShortName) = 0 Or Len(.AppLongName) = 0 Or Len(.ScriptURLPrim) = 0 Then
            Call CleanUp '--------------------- Delete all temp folders
            If Setup.RunMode = eNORMAL Then
                MsgBox "Live Update was unable to determine update application settings and will now terminate.  ", vbCritical, "Live Update Error"
            End If
            End '------------------------------ Terminate ReVive
        End If
    End With
    Set Me.Picture = LoadResPicture(300, 0) '-- Load form graphic from resource file
    Call DrawForm(Me) '------------------------ Clip the form and draw new border
    If Setup.RunMode = eNORMAL Then '---------- Skip in eNOTIFY and eAUTO modes
        Me.lblOpening.Caption = "LiveUpdate checks for updates available for your " & _
            Setup.AppLongName & " installation." & vbNewLine & vbNewLine & _
            "Please ensure " & Setup.AppShortName & " is closed and a connection " & _
            "to the internet exist before continuing." & Chr(13) & Chr(13) & _
            "Select Next when you are ready to continue or Cancel to exit."
        Me.Show '------------------------------ Display form for Normal mode
        Call cmdNext_FormActivate(Active) '---- Draw title bar and text (force because skipped while invisible)
    Else
        SetForegroundWindow lPREVWINDOW '------ Return focus to previously active window
        cmdNext_Click '------------------------ Proceed for eAUTO or eNOTIFY modes
    End If
    SetRect R, 108, 285, 504, 331 '------------ Draw gradient behind XP buttons and refresh RECT
    Call DrawGradient(Me.hdc, R, &HEBEDED, 11447982, VERTICAL)
    If Me.Visible Then Call RedrawWindow(Me.hWnd, R, 0&, RDW_FLAGS)
    Call SetRect(R, 2, 313, 108, 325) '------ Draw version info on bottom left of form
    Call DrawText(Me.hdc, "version " & App.Major & "." & App.Minor, -1, R, DT_FLAGS + DT_LEFT + DT_NOPREFIX + DT_CENTER)
    '//Repaint only the version info rect, NOT the entire window
    If Me.Visible Then Call RedrawWindow(Me.hWnd, R, 0&, RDW_FLAGS)
End Sub

Private Sub Form_Activate()
On Error Resume Next
    If Me.Visible Then Me.cmdNext.SetFocus '--- Return focus to cmdNext when returning from other forms
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'----------------------------------------------------------------------------------
' Purpose   : Either minimizes, closes, or move the form (depending on x and y)
'----------------------------------------------------------------------------------
On Error Resume Next
Dim R As RECT
    If Button = vbLeftButton Then
        Call SetRect(R, 486, 5, 502, 21) '------------ Minimize button
        If PtInRect(R, CLng(x), CLng(y)) Then
            Unload Me
            Exit Sub
        End If
        Call SetRect(R, 468, 5, 484, 21) '------------ Close button
        If PtInRect(R, CLng(x), CLng(y)) Then
            Me.WindowState = 1
            Exit Sub
        End If
        Call SetRect(R, 0, 0, Me.ScaleWidth, 24) '---- All other titlebar clicks
        If PtInRect(R, CLng(x), CLng(y)) Then
            SetCursor LoadCursor(0, IDC_SIZEALL)
            Call ReleaseCapture
            Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
            Exit Sub
        End If
    End If
End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
'---------------------------------------------------------------------------------
' Purpose   : Display system menu when user clicks our title bar RECT only
'---------------------------------------------------------------------------------
Dim pt  As POINTAPI
Dim R   As RECT
   If Button = vbRightButton Then
        Call SetRect(R, 0, 0, Me.ScaleWidth, 24)
        If PtInRect(R, CLng(x), CLng(y)) Then '-------- See if we are in the title bar area
            Call GetCursorPos(pt) '-------------------- Must use screen coordinates
            Call ShowSysMenu(Me.hWnd, pt.x, pt.y) '---- Popup forms system menu
        End If
   End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
On Error Resume Next
Dim sFile As String
    If Len(sUpdateMessage) Then
        Me.Hide
        frmUpdateMessage.Show 1 '--------------------------- Display update messages
    End If
    If Not bJustExit And Setup.RunMode = eNORMAL Then
        If frmExit.ShowYesNo = Yes Then
            sFile = Dir$(sTEMPDIR & "\*.*")
            Do While sFile <> ""
                Call DeleteFile(sTEMPDIR & "\" & sFile) '--- Delete ALL files that may have been downloaded so far
                sFile = Dir$(sTEMPDIR & "\*.*")
            Loop
        Else
            Cancel = 1
            Exit Sub
        End If
    End If
    If bREBOOT Then Call Reboot  '------------ Reboot computer
End Sub

Private Sub Form_Unload(Cancel As Integer)
On Error Resume Next
Dim f As Form
    If bWriteScript Then WriteSetupScript '--- Write local ReVive initialization script
    For Each f In Forms '--------------------- Ensure all forms are unloaded
        Unload f
    Next f
    Call CleanUp '---------------------------- Delete un-needed files and temp directory
    Set frmMain = Nothing
    End '------------------------------------- Required only when running in Auto or Notify
        '...---------------------------------- mode to exit form load event following call to
        '...---------------------------------- cmdNext_Click. If ommitted, error will occurr
        '...---------------------------------- if update script is not found on server and
        '...---------------------------------- we are in Auto or Notify modes only.
End Sub

Private Sub CleanUp()
'--------------------------------------------------------------------------------
' Purpose   : Deletes all downloaded files not reserved for post reboot updating.
'--------------------------------------------------------------------------------
On Error Resume Next
Dim x   As Long
Dim s   As String
Dim u   As Long
    With FileList '----------------------- Delete all files not reserved for post reboot updating
        u = UBound(.TempPath)
        For x = 1 To u
            If .Status(x) <> UPDATECOMPREBOOT Then
                Call DeleteFile(.TempPath(x) & ".tmp")
            End If
        Next x
    End With
    Call DeleteFile(sTEMPDIR & "\Notify.ico") '-- Delete Notification icon
    '//Temp folder deletes will only succeed if no files are remaining for post reboot update.
    '..All remaining files will be deleted on reboot after update via registry or WinInit.ini.
    s = Dir$(CurDir$) '------------------- Remove VB's directory lock on sTempDir
    Call RemoveDirectory(sTEMPDIR) '------ Remove our temp directories if empty
    Call RemoveDirectory(Left$(sTEMPDIR, InStrRev(sTEMPDIR, "\", , vbTextCompare) - 1))
End Sub

Private Sub AltLVBackground(lv As ListView, _
    Optional ByVal BackColorOne As OLE_COLOR = vbWhite, _
    Optional ByVal BackColorTwo As OLE_COLOR = 16054263) '--- 16054263 = Lite Gray
'-----------------------------------------------------------------------------------------------
' Purpose   : Alternates colors in a ListView
' PSC Post  : http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=51229&lngWId=1
'-----------------------------------------------------------------------------------------------
Dim h               As Single
Dim sw              As Single
Dim lScaleMode      As Integer
Dim picAlt          As PictureBox
    lScaleMode = lv.Parent.ScaleMode
    lv.Parent.ScaleMode = vbTwips
    Set picAlt = Me.Controls.Add("VB.picturebox", "picAlt")
    '//Draws the desired backcolor scheme in the picAlt picturebox
    '..then loads the image in the passed listviews picture.
    With lv
        If .View = lvwReport Then
            If .ListItems.Count Then
                .PictureAlignment = lvwTile
                h = .ListItems(1).Height
                With picAlt
                    .BackColor = BackColorOne
                    .BorderStyle = 0
                    .AutoRedraw = True
                    .Height = h * 2
                    .Width = 10 * Screen.TwipsPerPixelX
                    sw = .ScaleWidth
                    picAlt.Line (0, h)-Step(sw, h), BackColorTwo, BF
                    Set lv.Picture = .Image
                    Set .Picture = Nothing
                End With
            End If
        End If
    End With
    Set picAlt = Nothing
    Me.Controls.Remove "picAlt"
    lv.Parent.ScaleMode = lScaleMode
End Sub

Private Function EncryptString(ByVal sString As String) As String
'---------------------------------------------------------------------------------------
' Purpose   : Encrypts web URL strings prior to entry into the initialization script.
'             Only called from the WriteSetupScript sub.
'---------------------------------------------------------------------------------------
Dim x       As Integer
Dim y       As Integer
Dim sBuffer As String
    If Len(CRYPTPWD) Then
        For x = 1 To Len(sString)
            y = Asc(Mid$(sString, x, 1))
            y = y + Asc(Mid$(CRYPTPWD, (x Mod Len(CRYPTPWD)) + 1, 1))
            sBuffer = sBuffer & Chr$(y And &HFF)
        Next x
        EncryptString = sBuffer
    Else
        EncryptString = sString
    End If
End Function

Private Function DecryptString(ByVal sString As String) As String
'---------------------------------------------------------------------------------------
' Purpose   : Decrypts web URL strings when reading from the initialization script.
'             Only called from the Form_Load event.
'---------------------------------------------------------------------------------------
Dim x       As Integer
Dim y       As Integer
Dim sBuffer As String
    If Len(CRYPTPWD) Then
        For x = 1 To Len(sString)
            y = Asc(Mid$(sString, x, 1))
            y = y - Asc(Mid$(CRYPTPWD, (x Mod Len(CRYPTPWD)) + 1, 1))
            sBuffer = sBuffer & Chr$(y And &HFF)
        Next x
        DecryptString = sBuffer
    Else
        DecryptString = sString
    End If
End Function

Private Sub WriteSetupScript()
'----------------------------------------------------------------------------------------
' Purpose   : Creates local LiveUpdate initialization file for this app that will be
'             used next time to provide ScriptURL's, file version numbers, and app title.
'             This functionality allows the app developer to alternate script URL's
'             for planned server changes, or to change the LiveUpdate app title display.
'             Values provided in the on-line script are written locally for next time.
'----------------------------------------------------------------------------------------
On Error Resume Next
Dim x       As Long
Dim f       As Integer
Dim s       As String
Dim sEXT    As String
    s = vbNewLine
    With Setup
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
        s = s & "ScriptURLAlt=" & EncryptString(.ScriptURLAlt) & vbNewLine
        s = s & "LastChecked=" & Format(.LastChecked, "dd mmm yyyy") & vbNewLine & vbNewLine
    End With
    s = s & "[Files]" & vbNewLine
    With FileList
        For x = 1 To UBound(.Description)
            '//For non EXE, OCX and DLL files, write version info to setup script.
            '..Version info for EXE, OCX and DLL files is extracted from the files
            '..and will not be needed in the script.
            sEXT = "|" & GetFileExt(.InstallPath(x)) & "|"
            If InStr(1, "|EXE|OCX|DLL|", sEXT) = 0 Then
                If .Status(x) = UPDATECOMP Then
                    s = s & .Description(x) & "=" & .UpdateVersion(x) & vbNewLine
                Else
                    s = s & .Description(x) & "=" & .CurrentVersion(x) & vbNewLine
                End If
            End If
        Next x
    End With
    s = s & vbNewLine
    s = s & "***WARNING: MODIFYING THIS SCRIPT MAY CAUSE LIVEUPDATE TO FUNCTION IMPROPERLY***"
    s = s & vbNewLine
    s = s & "--------------------------------------------------------------------------------"
    f = FreeFile
    Open Setup.SetupScriptPath For Output As #f
        Print #f, s
    Close #f
End Sub

Private Sub UpdateIcon(ByVal FileNumber As Integer)
'---------------------------------------------------------------------------------------
' Purpose   : Updates a files listview icon when the file is downloaded.
'             Only called from ucDownloadComplete and only when icons need updated.
'---------------------------------------------------------------------------------------
On Error Resume Next
Static l As Integer
    If l = 0 Then l = UBound(FileList.Description)
    l = l + 1
    With Me.iml.ListImages
        .Add l, l & "|", GetIcon(FileList.TempPath(FileNumber), SHGFI_SMALLICON)
        Me.lvFiles.ListItems(FileNumber).SmallIcon = l
    End With
End Sub

Private Function GetIcon(sFile As String, Size As Long) As IPictureDisp
'---------------------------------------------------------------------------------------
' Purpose   : Gets an icon from the passed file.
'             Only called from the ListScanResults and UpdateIcon subs.
'---------------------------------------------------------------------------------------
On Error Resume Next
Dim hIcon       As Long
Dim pIcon       As IPictureDisp
Dim shINFO      As SHFILEINFO
Dim lResult     As Long
    lResult = SHGetFileInfo(sFile, 0, shINFO, Len(shINFO), SHGFI_ICON + Size)
    If lResult Then '------------------ If the call fails default to our res icon
        hIcon = shINFO.hIcon
        Set pIcon = IconToPicture(hIcon)
        Set GetIcon = pIcon
        Set pIcon = Nothing
    Else
        Set GetIcon = LoadResPicture(205, vbResIcon)
        bUpdateIcons = True '---------- Plan to update icons once download is complete
    End If
End Function

Private Function IconToPicture(hIcon As Long) As IPictureDisp
'---------------------------------------------------------------------------------------
' Purpose   : Creates a picture from the passed icon handle.
'             Only called from the GetIcon function.
'---------------------------------------------------------------------------------------
On Error Resume Next
Dim lResult         As Long
Dim tCLSID          As CLSID
Dim tIcon           As TypeIcon
Dim pIcon           As IPictureDisp
    With tIcon
        .cbSize = Len(tIcon)
        .picType = vbPicTypeIcon
        .hIcon = hIcon
    End With
    With tCLSID
        .id(8) = &HC0
        .id(15) = &H46
    End With
    lResult = OleCreatePictureIndirect(tIcon, tCLSID, 1, pIcon)
    If lResult = 0 Then '------------- Picture was successfully created
        Set IconToPicture = pIcon
        Set pIcon = Nothing
    Else '---------------------------- If the call fails default to our res icon
        Set IconToPicture = LoadResPicture(205, vbResIcon)
        bUpdateIcons = True '--------- Prepare to update icons as downloads are completed
    End If
End Function

VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmInsert 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Insert Update File"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8400
   HelpContextID   =   3
   Icon            =   "frmInsert.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5760
   ScaleWidth      =   8400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtUpdateMessage 
      Height          =   555
      Left            =   2550
      MultiLine       =   -1  'True
      TabIndex        =   8
      Tag             =   "UpdateMessage"
      Top             =   2940
      Width           =   5595
   End
   Begin VB.CommandButton cmdHelp 
      Cancel          =   -1  'True
      Caption         =   "Help"
      Height          =   375
      Left            =   5250
      TabIndex        =   20
      Top             =   5190
      Width           =   915
   End
   Begin VB.ComboBox cmbDefConst 
      Height          =   315
      ItemData        =   "frmInsert.frx":058A
      Left            =   2550
      List            =   "frmInsert.frx":05AC
      Style           =   2  'Dropdown List
      TabIndex        =   3
      Tag             =   "InstallPath"
      Top             =   1560
      Width           =   1845
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   4380
      TabIndex        =   4
      Tag             =   "InstallPath"
      Text            =   "\"
      Top             =   1560
      Width           =   3765
   End
   Begin VB.TextBox txtFileSize 
      Height          =   315
      Left            =   2550
      TabIndex        =   5
      Tag             =   "FileSize"
      Top             =   1920
      Width           =   1575
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   5190
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7230
      TabIndex        =   9
      Top             =   5190
      Width           =   915
   End
   Begin VB.TextBox txtURL 
      Height          =   315
      Left            =   2550
      TabIndex        =   2
      Tag             =   "DownloadURL"
      Top             =   1200
      Width           =   5595
   End
   Begin VB.TextBox txtDesc 
      Height          =   315
      Left            =   2550
      TabIndex        =   0
      Tag             =   "Description"
      Top             =   480
      Width           =   5595
   End
   Begin VB.TextBox txtVer 
      Height          =   315
      Left            =   2550
      TabIndex        =   1
      Tag             =   "UpdateVersion"
      Top             =   840
      Width           =   1575
   End
   Begin VB.CheckBox chkMustExist 
      Height          =   195
      Left            =   2550
      TabIndex        =   7
      Tag             =   "MustExist"
      Top             =   2625
      Value           =   1  'Checked
      Width           =   195
   End
   Begin VB.CheckBox chkMustUpdate 
      Height          =   195
      Left            =   2550
      TabIndex        =   6
      Tag             =   "MustUpdate"
      Top             =   2295
      Width           =   195
   End
   Begin RichTextLib.RichTextBox rtTip 
      Height          =   1275
      Left            =   330
      TabIndex        =   22
      TabStop         =   0   'False
      Top             =   3720
      Width           =   7830
      _ExtentX        =   13811
      _ExtentY        =   2249
      _Version        =   393217
      BackColor       =   14286847
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmInsert.frx":0611
   End
   Begin VB.Label lbl 
      Caption         =   "UpdateMessage="
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   7
      Left            =   780
      TabIndex        =   21
      Top             =   2970
      Width           =   1725
   End
   Begin VB.Label lbl 
      Caption         =   "FileSize="
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   6
      Left            =   780
      TabIndex        =   19
      Top             =   1950
      Width           =   1725
   End
   Begin VB.Label lblSelFile 
      Caption         =   "Select File"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00C00000&
      Height          =   195
      Left            =   7410
      TabIndex        =   18
      Top             =   150
      Width           =   735
   End
   Begin VB.Label lbl 
      Caption         =   "InstallPath="
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   5
      Left            =   780
      TabIndex        =   17
      Top             =   1590
      Width           =   1725
   End
   Begin VB.Label lbl 
      Caption         =   "DownloadURL="
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   4
      Left            =   780
      TabIndex        =   16
      Top             =   1230
      Width           =   1725
   End
   Begin VB.Label lbl 
      Caption         =   "Description="
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   3
      Left            =   780
      TabIndex        =   15
      Top             =   510
      Width           =   1725
   End
   Begin VB.Label lbl 
      Caption         =   "UpdateVersion="
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   2
      Left            =   780
      TabIndex        =   14
      Top             =   870
      Width           =   1725
   End
   Begin VB.Label lbl 
      Caption         =   "MustExist="
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   1
      Left            =   780
      TabIndex        =   13
      Top             =   2610
      Width           =   1725
   End
   Begin VB.Label lbl 
      Caption         =   "MustUpdate="
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Index           =   0
      Left            =   780
      TabIndex        =   12
      Top             =   2280
      Width           =   1725
   End
   Begin VB.Label lblSec 
      Caption         =   "[File XX]"
      BeginProperty Font 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800000&
      Height          =   240
      Left            =   150
      TabIndex        =   11
      Top             =   150
      Width           =   1080
   End
End
Attribute VB_Name = "frmInsert"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//DECLARES FOR GetFileVersion
Private Declare Function GetFileVersionInfo Lib "Version.dll" Alias "GetFileVersionInfoA" (ByVal lptstrFilename As String, ByVal dwhandle As Long, ByVal dwlen As Long, lpData As Any) As Long
Private Declare Function GetFileVersionInfoSize Lib "Version.dll" Alias "GetFileVersionInfoSizeA" (ByVal lptstrFilename As String, lpdwHandle As Long) As Long
Private Declare Function VerQueryValue Lib "Version.dll" Alias "VerQueryValueA" (pBlock As Any, ByVal lpSubBlock As String, lplpBuffer As Long, puLen As Long) As Long
Private Declare Function lstrlenA Lib "kernel32" (ByVal lpString As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32" Alias "RtlMoveMemory" (Destination As Any, source As Any, ByVal Length As Long)
Private Type VS_FIXEDFILEINFO
    dwSignature         As Long
    dwStrucVersion      As Long
    dwFileVersionMS     As Long
    dwFileVersionLS     As Long
    dwProductVersionMS  As Long
    dwProductVersionLS  As Long
    dwFileFlagsMask     As Long
    dwFileFlags         As Long
    dwFileOS            As Long
    dwFileType          As Long
    dwFileSubtype       As Long
    dwFileDateMS        As Long
    dwFileDateLS        As Long
End Type

Private lLastLine       As Long '------------ Stores line value for inserting next file

Private Sub chkMustExist_GotFocus()
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub chkMustUpdate_GotFocus()
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub cmbDefConst_GotFocus()
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub Form_Load()
Dim x As Byte
    For x = 0 To 7 '----------------------------------------------- Color labels
        Me.lbl(x).ForeColor = Setup.KeyTagColor
    Next x
    Me.lblSec.ForeColor = Setup.SecTagColor
    With Me.cmbDefConst
        .Text = Setup.DefaultConst '------------------------------- Set default script constant
        .ForeColor = Setup.ValTagColor
    End With
    Me.lblSec.Caption = "[File " & Format(FindNextFile, "00]") '--- Display next file assignment
    Me.txtURL.Text = Setup.DefaultWeb '---------------------------- Set deafult web URL
End Sub

Private Sub cmdHelp_Click()
    '//Open help file and display contents
    WinHelp Me.hwnd, App.path & "\ReVive.hlp", HELP_CONTEXT, CLng(3)
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
'//Create file entry and add it to the richtextbox.
Dim s As String
Dim x As Long
    If Not IsVersionValid(Me.txtVer.Text) Then
        MsgBox "Version number is not valid. Please see tip window for details.   ", vbExclamation, "Version Number Error"
        Me.txtVer.SetFocus
        Exit Sub
    End If
    s = vbNewLine
    s = s & Me.lblSec.Caption
    s = s & vbNewLine
    s = s & vbTab & "Description=" & vbTab & Me.txtDesc.Text
    s = s & vbNewLine
    s = s & vbTab & "UpdateVersion=" & vbTab & Me.txtVer.Text
    s = s & vbNewLine
    s = s & vbTab & "DownloadURL=" & vbTab & Me.txtURL.Text
    s = s & vbNewLine
    s = s & vbTab & "InstallPath=" & vbTab & Me.cmbDefConst.Text & Me.txtPath.Text
    s = s & vbNewLine
    s = s & vbTab & "FileSize=" & vbTab & vbTab & Me.txtFileSize.Text
    s = s & vbNewLine
    s = s & vbTab & "MustUpdate=" & vbTab & vbTab & CBool(Me.chkMustUpdate.Value)
    s = s & vbNewLine
    s = s & vbTab & "MustExist=" & vbTab & vbTab & CBool(Me.chkMustExist.Value)
    s = s & vbNewLine
    If Len(Trim$(Me.txtUpdateMessage.Text)) Then
        s = s & vbTab & "UpdateMessage=" & vbTab & Me.txtUpdateMessage.Text
        s = s & vbNewLine
    End If
    SetWindowState bLocked
    With frmMain.rtBox
        .Text = .Text & s
        For x = 1 To lLastLine + 10
            frmMain.ColorLine x
        Next x
        .SelStart = Len(.Text)
    End With
    SetWindowState bUnLocked
    Unload Me
End Sub

Private Function FindNextFile() As Byte
'//Discovers next available file entry number.
Dim x           As Long
Dim Y           As Long
Dim sUsed       As String
Dim lLineCount  As Long
Dim sText       As String * 255 '--------- Buffer for EM_GETLINE call
Dim sLineText   As String '--------------- Line text before trimming
Dim lStartChar  As Long
Dim lLineLength As Long
    lLineCount = SendMessage(rtHwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
    lLastLine = lLineCount      '//Save line number for inserting new file
    sUsed = "|"
    For x = 0 To lLineCount - 1
        lStartChar = SendMessage(rtHwnd, EM_LINEINDEX, x, ByVal 0&)
        lLineLength = SendMessage(rtHwnd, EM_LINELENGTH, lStartChar, ByVal 0&)
        sText = Space(255): Call SendMessage(rtHwnd, EM_GETLINE, x, ByVal sText)
        sLineText = Trim$(StripNulls(Left$(sText, lLineLength))): sText = ""
        If Left$(sLineText, 1) = "[" Then
            For Y = 1 To 99
                If Left$(sLineText, 9) = "[File " & Format(Y, "00]") Then
                    sUsed = sUsed & Y & "|"
                    Exit For
                End If
            Next Y
        End If
    Next x
    For x = 1 To 99
        If InStr(1, sUsed, "|" & x & "|") = 0 Then
            Exit For
        End If
    Next x
    FindNextFile = x
End Function

Private Function StripNulls(ByVal sString As String) As String
    sString = Replace(sString, vbTab, " ")
    sString = Replace(sString, vbLf, " ")
    sString = Replace(sString, vbCr, " ")
    StripNulls = sString
End Function

Private Sub SelectAllText()
On Error Resume Next
    With Screen.ActiveControl
        .SelStart = 0
        .SelLength = Len(.Text)
    End With
    If Err Then Err.Clear
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmInsert = Nothing
End Sub

Private Sub lblSelFile_Click()
'//Opens file open dialog and fills file textboxes with selected file information
On Error GoTo Errs
Dim sFile       As String
Dim sFileName   As String
Dim x           As Long
    With frmMain.cd
        .Filter = "Executable files (*.exe;*.dll;*.ocx)|*.exe;*.dll;*.ocx|" & _
                    "Drivers (*.sys;*.drv;*.fnt)|*.sys;*.drv;*.fnt|" & _
                    "All files (*.*)|*.*"
        .DefaultExt = "*.exe"
        .DialogTitle = "Select an update file..."
        .Flags = cdlOFNFileMustExist Or cdlOFNExplorer _
            Or cdlOFNHideReadOnly Or cdlOFNPathMustExist _
            Or cdlOFNShareAware Or cdlOFNNoDereferenceLinks
        .ShowSave
        sFile = .FileName
    End With
    Me.txtDesc.Text = GetFileDescription(sFile)
    Me.txtVer.Text = GetFileVersion(sFile)
    Me.txtFileSize.Text = FileLen(sFile)
    sFileName = Right$(sFile, Len(sFile) - InStrRev(sFile, "\"))
    Me.txtPath.Text = "\" & sFileName
    Me.txtURL.Text = Setup.DefaultWeb & sFileName
    Me.txtDesc.SetFocus
    Me.txtDesc.SelLength = Len(Me.txtDesc.Text)
Errs:
    If Err Then Exit Sub
End Sub

Private Function GetFileVersion(ByVal sFileName As String) As String
'**********************************************************************
'Adapted from code posted by Eric D. Burdo, http://www.rlisolutions.com
'"Retrieve the version number of a DLL"
'See full post here: http://programmers-corner.com/viewSource.php/71
'**********************************************************************
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
        CopyMemory sVerInfo, ByVal lBuff, lFreeSize
    End If
    iMajor = CInt(sVerInfo.dwFileVersionMS \ &H10000)
    iMinor = CInt(sVerInfo.dwFileVersionMS And &HFFFF&)
    sMajor = CStr(iMajor) & "." & LTrim$(CStr(iMinor))
    iMajor = CInt(sVerInfo.dwFileVersionLS \ &H10000)
    iMinor = CInt(sVerInfo.dwFileVersionLS And &HFFFF&)
    sMinor = CStr(iMajor) & "." & LTrim$(CStr(iMinor))
    GetFileVersion = sMajor & "." & sMinor
End Function

Private Function GetFileDescription(ByVal sFileName As String) As String
Dim nDummy          As Long
Dim nRet            As Long
Dim sBuffer()       As Byte
Dim nBufferLen      As Long
Dim lplpBuffer      As Long
Dim udtVerBuffer    As VS_FIXEDFILEINFO
Dim puLen           As Long
Dim nLanguage       As Integer
Dim nCodePage       As Integer
Dim sSubBlock       As String
    nBufferLen = GetFileVersionInfoSize(sFileName, nDummy)
    If nBufferLen = 0 Then Exit Function
    ReDim sBuffer(nBufferLen) As Byte
    Call GetFileVersionInfo(sFileName, 0&, nBufferLen, sBuffer(0))
    Call VerQueryValue(sBuffer(0), "\", lplpBuffer, puLen)
    Call CopyMemory(udtVerBuffer, ByVal lplpBuffer, Len(udtVerBuffer))
    If VerQueryValue(sBuffer(0), "\VarFileInfo\Translation", lplpBuffer, puLen) Then
        If puLen Then
            nRet = PointerToDWord(lplpBuffer)
            nLanguage = LoWord(nRet)
            nCodePage = HiWord(nRet)
            sSubBlock = "\StringFileInfo\" & FmtHex(&H409, 4) & FmtHex(nCodePage, 4) & "\"
            GetFileDescription = GetStdValue(VarPtr(sBuffer(0)), sSubBlock & "FileDescription")
        End If
    End If
End Function

Private Function GetStdValue(ByVal lpBlock As Long, ByVal Value As String) As String
Dim lplpBuffer       As Long
Dim puLen     As Long
    If VerQueryValue(ByVal lpBlock, Value, lplpBuffer, puLen) Then
        If puLen Then
            GetStdValue = PointerToString(lplpBuffer)
        End If
    End If
End Function

Private Function PointerToString(lpString As Long) As String
Dim Buffer As String
Dim nLen As Long
    If lpString Then
        nLen = lstrlenA(lpString)
        If nLen Then
            Buffer = Space(nLen)
            CopyMemory ByVal Buffer, ByVal lpString, nLen
            PointerToString = Buffer
        End If
    End If
End Function

Private Function FmtHex(ByVal InVal As Long, ByVal OutLen As Integer) As String
    FmtHex = Right$(String$(OutLen, "0") & Hex$(InVal), OutLen)
End Function

Private Function PointerToDWord(ByVal lpDWord As Long) As Long
Dim nRet As Long
    If lpDWord Then
        CopyMemory nRet, ByVal lpDWord, 4
        PointerToDWord = nRet
    End If
End Function

Private Function LoWord(ByVal LongIn As Long) As Integer
    Call CopyMemory(LoWord, LongIn, 2)
End Function

Private Function HiWord(ByVal LongIn As Long) As Integer
    Call CopyMemory(HiWord, ByVal (VarPtr(LongIn) + 2), 2)
End Function

Private Sub lblSelFile_MouseMove(Button As Integer, Shift As Integer, x As Single, Y As Single)
On Error Resume Next
    If x >= 0 And Y >= 0 And x < Me.ActiveControl.Width And Y < Me.ActiveControl.Height Then
        SetCursor LoadCursor(0, IDC_HAND)
    End If
    Err.Clear
End Sub

Private Function LimitTextInput(ByVal source As Integer, ByVal bAllowSeperators As Boolean) As String
'//Limits character input into a control.
Dim Numbers As String
    Numbers = IIf(bAllowSeperators, "0123456789.", "0123456789")
    If source <> 8 Then
        If InStr(Numbers, Chr(source)) = 0 Then
            LimitTextInput = 0
            Beep
            Exit Function
        End If
    End If
    LimitTextInput = source
End Function

Private Sub txtDesc_GotFocus()
    Call SelectAllText
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub txtFileSize_GotFocus()
    SelectAllText
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub txtPath_GotFocus()
    Call SelectAllText
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub txtUpdateMessage_GotFocus()
    Call SelectAllText
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub txtURL_GotFocus()
    Call SelectAllText
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub txtVer_GotFocus()
    Call SelectAllText
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub txtVer_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitTextInput(KeyAscii, True)
End Sub

Private Sub txtFileSize_KeyPress(KeyAscii As Integer)
    KeyAscii = LimitTextInput(KeyAscii, False)
End Sub

VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{38911DA0-E448-11D0-84A3-00DD01104159}#1.1#0"; "COMCT332.OCX"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmMain 
   Caption         =   "ReVive Update Script Editor"
   ClientHeight    =   5235
   ClientLeft      =   3975
   ClientTop       =   4035
   ClientWidth     =   8355
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   5235
   ScaleWidth      =   8355
   Visible         =   0   'False
   Begin RichTextLib.RichTextBox rtTip 
      Height          =   1485
      Left            =   5910
      TabIndex        =   5
      TabStop         =   0   'False
      Top             =   3720
      Width           =   2415
      _ExtentX        =   4260
      _ExtentY        =   2619
      _Version        =   393217
      BackColor       =   14286847
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmMain.frx":08CA
   End
   Begin MSComctlLib.TreeView tv 
      Height          =   2760
      Left            =   5910
      TabIndex        =   3
      TabStop         =   0   'False
      Top             =   960
      Width           =   2430
      _ExtentX        =   4286
      _ExtentY        =   4868
      _Version        =   393217
      HideSelection   =   0   'False
      Indentation     =   573
      LabelEdit       =   1
      Style           =   7
      Appearance      =   1
   End
   Begin MSComctlLib.ImageList iml 
      Left            =   60
      Top             =   4590
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   14
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":094C
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":0EE6
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1480
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":181A
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":1DB4
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":214E
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":26E8
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":2C82
            Key             =   ""
         EndProperty
         BeginProperty ListImage9 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":321C
            Key             =   ""
         EndProperty
         BeginProperty ListImage10 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":37B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage11 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":3D50
            Key             =   ""
         EndProperty
         BeginProperty ListImage12 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":42EA
            Key             =   ""
         EndProperty
         BeginProperty ListImage13 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4684
            Key             =   ""
         EndProperty
         BeginProperty ListImage14 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmMain.frx":4A1E
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   630
      Top             =   4680
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
      CancelError     =   -1  'True
   End
   Begin RichTextLib.RichTextBox rtBox 
      Height          =   4635
      Left            =   0
      TabIndex        =   0
      Top             =   600
      Visible         =   0   'False
      Width           =   5895
      _ExtentX        =   10398
      _ExtentY        =   8176
      _Version        =   393217
      BackColor       =   16777215
      BorderStyle     =   0
      Enabled         =   -1  'True
      HideSelection   =   0   'False
      ScrollBars      =   3
      AutoVerbMenu    =   -1  'True
      OLEDragMode     =   0
      OLEDropMode     =   0
      TextRTF         =   $"frmMain.frx":4B78
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Courier New"
         Size            =   9.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin ComCtl3.CoolBar CoolBar1 
      Align           =   1  'Align Top
      Height          =   600
      Left            =   0
      TabIndex        =   1
      Top             =   0
      Width           =   8355
      _ExtentX        =   14737
      _ExtentY        =   1058
      BandCount       =   1
      _CBWidth        =   8355
      _CBHeight       =   600
      _Version        =   "6.7.9782"
      Child1          =   "tbMain"
      MinHeight1      =   540
      Width1          =   1635
      NewRow1         =   0   'False
      AllowVertical1  =   0   'False
      Begin MSComctlLib.Toolbar tbMain 
         Height          =   540
         Left            =   30
         Negotiate       =   -1  'True
         TabIndex        =   2
         Top             =   30
         Width           =   8235
         _ExtentX        =   14526
         _ExtentY        =   953
         ButtonWidth     =   1244
         ButtonHeight    =   953
         Style           =   1
         ImageList       =   "iml"
         _Version        =   393216
         BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
            NumButtons      =   15
            BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "New"
               Object.ToolTipText     =   "Create New Script File"
               ImageIndex      =   1
            EndProperty
            BeginProperty Button2 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Open"
               Object.ToolTipText     =   "Open Script File"
               ImageIndex      =   2
            EndProperty
            BeginProperty Button3 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Save"
               Object.ToolTipText     =   "Save Current File"
               ImageIndex      =   3
            EndProperty
            BeginProperty Button4 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button5 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Print"
               Object.ToolTipText     =   "Print Current File"
               ImageIndex      =   4
            EndProperty
            BeginProperty Button6 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Test"
               Object.ToolTipText     =   "Test Script File"
               ImageIndex      =   5
            EndProperty
            BeginProperty Button7 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Dist"
               Object.ToolTipText     =   "Create Distribution File"
               ImageIndex      =   6
            EndProperty
            BeginProperty Button8 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button9 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Insert"
               Object.ToolTipText     =   "Insert Update File"
               ImageIndex      =   7
            EndProperty
            BeginProperty Button10 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Comm"
               Object.ToolTipText     =   "Comment Selection"
               ImageIndex      =   8
            EndProperty
            BeginProperty Button11 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Uncomm"
               Object.ToolTipText     =   "Uncomment Selection"
               ImageIndex      =   9
            EndProperty
            BeginProperty Button12 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button13 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Options"
               Object.ToolTipText     =   "Program Options"
               ImageIndex      =   10
            EndProperty
            BeginProperty Button14 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Style           =   3
            EndProperty
            BeginProperty Button15 {66833FEA-8583-11D1-B16A-00C0F0283628} 
               Caption         =   "Help"
               Object.ToolTipText     =   "View Help"
               ImageIndex      =   11
            EndProperty
         EndProperty
         OLEDropMode     =   1
      End
   End
   Begin VB.Label lblAdd 
      Alignment       =   2  'Center
      Caption         =   "Double Click Sub-Keys to Insert"
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   5940
      TabIndex        =   4
      Top             =   675
      Width           =   2325
   End
   Begin VB.Menu menFile 
      Caption         =   "&File"
      Begin VB.Menu menNew 
         Caption         =   "&New..."
         Shortcut        =   ^N
      End
      Begin VB.Menu menOpen 
         Caption         =   "&Open..."
         Shortcut        =   ^O
      End
      Begin VB.Menu menSep1 
         Caption         =   "-"
      End
      Begin VB.Menu menSave 
         Caption         =   "&Save"
         Shortcut        =   ^S
      End
      Begin VB.Menu menSaveAs 
         Caption         =   "Save &As..."
         Shortcut        =   ^A
      End
      Begin VB.Menu menSep2 
         Caption         =   "-"
      End
      Begin VB.Menu menPrint 
         Caption         =   "&Print"
         Shortcut        =   ^P
      End
      Begin VB.Menu menSep3 
         Caption         =   "-"
      End
      Begin VB.Menu menExit 
         Caption         =   "E&xit Script Editor         "
         Shortcut        =   ^X
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'//Enumerators to identify acceptible data types. Each Key is assigned
'..an expected data type and are tested to ensure values meet criteria.
Private Enum enumDataTypes
    eBOOLEAN = 0
    eINTEGER = 1
    eLONG = 2
    eSTRING = 3
    eWEB = 4
    ePATH = 5
    eVERSION = 6
End Enum

'//Enumerators to identify header (section) for determining what keys may be used under them.
'..Used when Testing script for proper syntax and Value datatypes.
Private Enum enumSections
    eNONE = 0
    eSETUP = 1
    eFILES = 2
End Enum

Private Enum enumColors '------------------ Enumerators for tag colors
    eERROR = &HFF&
    eCOMMENT = &H8000&
End Enum

Private Type typeSectionTags '------------- Type declares for Section Header Tags
    Tag()       As String
    Length()    As Long
    Section()   As enumSections
End Type

Private Type typeKeyTags '----------------- Type declares for Key Tags
    Tag()       As String
    Length()    As Long
    Section()   As enumSections
    DataType()  As enumDataTypes
    HasValTag() As Boolean
End Type

Private Type typeValueTags '--------------- Type declares for Value Tags
    Tag()       As String
    Length()    As Long
    DataType()  As enumDataTypes
    TextAfter() As Boolean
End Type

Private Declare Function SetWindowLong Lib "user32" Alias "SetWindowLongA" (ByVal hwnd _
    As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const WM_USER               As Long = &H400 '---- Disable wordwrap declares
Private Const EM_SETTARGETDEVICE    As Long = WM_USER + 72

Private bSkipColoring               As Boolean '------- True when actively coloring a line
Private sExpValue                   As String '-------- Filled with expected datatype when testing

Private CURRLINENUMBER              As Long '---------- Variables that are refreshed in the rtBox_SelChange sub when the selected line number changes
Private CURRLINELENGTH              As Long
Private PREVLINENUMBER              As Long

Private Const GWL_WNDPROC           As Long = (-4)

Private SecTags                     As typeSectionTags
Private KeyTags                     As typeKeyTags
Private ValTags                     As typeValueTags

Private Sub Form_Load()
    rtHwnd = Me.rtBox.hwnd
    Call SendMessage(rtHwnd, EM_SETTARGETDEVICE, 0, 0) '----- Disable wordwrap (ESSENTIAL IN THIS APP)
    Me.rtBox.SelIndent = 100 '------------------------------- Set indent at left
    Me.WindowState = 2 '------------------------------------- Go maximized
    Call LoadTags '------------------------------------------ Assign script tags
    bChanged = False
    If Len(Setup.Script) Then Call FileOpen(Setup.Script) '-- Open passed file
    bChanged = False '--------------------------------------- Reset changed flag
    Call FillTreeView
    Call tv_NodeClick(Me.tv.Nodes(1))
    If Not InIDE Then '-------------------------------------- Skip all this in IDE to avoid crashes
        OldWindowProc = GetWindowLong(Me.hwnd, GWL_WNDPROC)
        Call SetWindowLong(Me.hwnd, GWL_WNDPROC, AddressOf SubClass_WndMessage)
    End If
    Call SetTVColor
    Me.Show
    Me.rtBox.Visible = True
    Me.rtBox.SetFocus
End Sub

Private Sub tbMain_ButtonClick(ByVal Button As MSComctlLib.Button)
On Error GoTo Errs
    Select Case Button.Caption
        Case "New"
            Call FileNew
        Case "Save"
            Call FileSave
        Case "Print"
            Call FilePrint
        Case "Comm"
            Call Comment
        Case "Uncomm"
            Call UnComment
        Case "Open"
            Call FileOpen
        Case "Test"
            Call TestScript
        Case "Options"
            frmOptions.Show 1
        Case "Insert"
            frmInsert.Show 1
        Case "Dist"
            Call CreateDistributable
        Case "Help"
            WinHelp Me.hwnd, App.path & "\ReVive.hlp", HELP_CONTENTS, 0&
    End Select
Errs:
    If Err Then MsgBox Err.Description
    Exit Sub
End Sub

Private Sub LoadTags()
Dim x As Long

'******** IF YOU WANT TO USE MORE HIGHLIGHTED TAGS, JUST ADD THEM HERE *********

'*************************** DECRIPTION OF TAGS ********************************
'1. SectionTags are headers for each section, such as [Setup].
'2. KeyTags are keys used in the script, such as "AdminRequired=".
'3. ValueTags are values deemed appropriate for a datatype assigned to a KeyTag.
'
' Syntax: AddKeyTag "TAGNAME=" As String, SECTIONNAME As enumSections, _
'               VALUETYPE as enumDataTypes, HASVALUETAG as Boolean
'
'         AddValueTag "TAGNAME"=, DATATYPE as enumDataTypes, _
'               ALLOWTEXTAFTER as Boolean
'
'         AddSectionTag "TAGNAME" As String, Section As enumSections
'*******************************************************************************

    AddKeyTag "AdminRequired=", eSETUP, eBOOLEAN, True '------ Add Setup section tags
    AddKeyTag "ForceReboots=", eSETUP, eBOOLEAN, True
    AddKeyTag "ScriptURLPrim=", eSETUP, eWEB, False
    AddKeyTag "ScriptURLAlt=", eSETUP, eWEB, False
    AddKeyTag "AppShortName=", eSETUP, eSTRING, False
    AddKeyTag "AppLongName=", eSETUP, eSTRING, False
    AddKeyTag "NotifyIcon=", eSETUP, eWEB, False
    AddKeyTag "UpdateAppTitle=", eSETUP, eSTRING, False
    AddKeyTag "UpdateAppClass=", eSETUP, eSTRING, False
    AddKeyTag "UpdateAppKill=", eSETUP, eBOOLEAN, True
    AddKeyTag "LaunchIfKilled=", eSETUP, ePATH, True
    AddKeyTag "ShowFileIcons=", eSETUP, eBOOLEAN, True
    AddKeyTag "RegRISFiles=", eSETUP, eBOOLEAN, True
    AddKeyTag "Description=", eFILES, eSTRING, False '-------- Add File section tags
    AddKeyTag "UpdateVersion=", eFILES, eVERSION, False
    AddKeyTag "DownloadURL=", eFILES, eWEB, False
    AddKeyTag "InstallPath=", eFILES, ePATH, True
    AddKeyTag "FileSize=", eFILES, eLONG, False
    AddKeyTag "MustUpdate=", eFILES, eBOOLEAN, True
    AddKeyTag "MustExist=", eFILES, eBOOLEAN, True
    AddKeyTag "UpdateMessage=", eFILES, eSTRING, False
    AddValueTag "<ap>", ePATH, True '------------------------- Add Value tags
    AddValueTag "<sp>", ePATH, True
    AddValueTag "<win>", ePATH, True
    AddValueTag "<sys>", ePATH, True
    AddValueTag "<temp>", ePATH, True
    AddValueTag "<pf>", ePATH, True
    AddValueTag "<cf>", ePATH, True
    AddValueTag "<userdesktop>", ePATH, True
    AddValueTag "<userstartmenu>", ePATH, True
    AddValueTag "<commondesktop>", ePATH, True
    AddValueTag "<commonstartmenu>", ePATH, True
    AddValueTag "True", eBOOLEAN, False
    AddValueTag "False", eBOOLEAN, False
    AddValueTag "0", eBOOLEAN, False
    AddValueTag "1", eBOOLEAN, False
    AddSectionTag "[Setup]", eSETUP '------------------------- Add Section header tags
    For x = 1 To 99
        AddSectionTag "[File " & Format(x, "00]"), eFILES
    Next x
End Sub

Private Sub Form_Resize()
On Error Resume Next
Dim fH As Long
Dim fW As Long
    fH = Me.Height
    fW = Me.Width
    Me.rtBox.Move 0, 600, fW - 2785, fH - 1270
    Me.tv.Move fW - 2755, 960, 2635, fH - 3775
    Me.rtTip.Move fW - 2765, fH - 2805, 2645, 2130
    Me.lblAdd.Move fW - 2600
End Sub

Private Sub menExit_Click()
    Unload Me
End Sub

Private Sub menNew_Click()
    Call FileNew
End Sub

Private Sub menOpen_Click()
    Call FileOpen
End Sub

Private Sub menPrint_Click()
    Call FilePrint
End Sub

Private Sub menSave_Click()
    Call FileSave
End Sub

Private Sub menSaveAs_Click()
On Error GoTo Errs
    FileSave True
Errs:
    Exit Sub
End Sub

Private Sub rtBox_Change()
'---------------------------------------------------------------------------
' Purpose   : Colors current line when text is altered if CURRLINELENGTH > 0
'---------------------------------------------------------------------------
    bChanged = True
    If Not bSkipColoring Then
        If CURRLINELENGTH Then Call ColorLine(CURRLINENUMBER)
    End If
End Sub

Private Sub rtBox_SelChange()
'----------------------------------------------------------------------------
' Purpose   : Checks if user transversed lines, and if they did see if text
'             was modified on last line. If modified, recolor PREVLINENUMBER.
'----------------------------------------------------------------------------
Dim lSelStart As Long
Dim lResult As Long
    lSelStart = Me.rtBox.SelStart
    CURRLINENUMBER = 1 + SendMessage(rtHwnd, EM_LINEFROMCHAR, lSelStart, ByVal 0&)
    CURRLINELENGTH = SendMessage(rtHwnd, EM_LINELENGTH, lSelStart, ByVal 0&)
    If Not bSkipColoring And (PREVLINENUMBER <> CURRLINENUMBER) Then
        lResult = SendMessage(rtHwnd, EM_GETMODIFY, ByVal 0&, ByVal 0&)
        If lResult <> 0 Then
            Call ColorLine(PREVLINENUMBER)
            Call SendMessage(rtHwnd, EM_SETMODIFY, False, ByVal 0&)
        End If
        PREVLINENUMBER = CURRLINENUMBER
    End If
End Sub

Public Sub ColorLine(ByVal lLine As Long)
'------------------------------------------------------------------------------------------------------
' Purpose   : Colors text of passed line. This is alot of code, but it is thorough and fairly quick.
'             It allows formatting with tabs and spaces, and is Tag and Key location and
'             data type aware. For example:
'             1. Does not recognize section headers not identified in LoadTag procedure.
'             2. Value tags are only recognized where Key and Value data types match.
'             3. Value tags are only recognized following an =, and allow spaces and tabs between.
'             4. Tags are colored and altered as typed to match case specified in LoadTag procedure.
'             5. Ignored text after identified section headers is colored as commented.
'             6. Illegal text after Values (when not allowed per LoadTag) is colored red.
'             7. Previous lines are always recolored when exiting them. (See rtBox_SelChange)
'             8. RichTextBox autoscroll is effectively disabled when coloring previous lines.
'             9. Users caret point and selection length is always restored when procedure completes.
'------------------------------------------------------------------------------------------------------
Dim x               As Long         '//Perennial favorite, gotta have it!
Dim Y               As Long         '//Perennial favorite, gotta have it!
Dim z               As Long         '//Perennial favorite, gotta have it!
Dim lSelStart       As Long         '//Starting cursor position
Dim lSelStartLen    As Long         '//Starting selection length
Dim lStartChar      As Long         '//First character in current line
Dim sTrimmedLine    As String       '//Line trimmed of Nulls and spaces
Dim sTrimmedLine2   As String       '//Misc uses
Dim bTagFound       As Boolean      '//True when a tag is found
Dim lConstStart     As Long         '//Beginned char of tag or directory contants
Dim lConstEnd       As Long         '//Ending char of tag or directory contants
Dim lLineLength     As Long         '//Length of untrimmed line
Dim sLineText       As String       '//Line before trimming
Dim sValueText      As String       '//Value text after = sign
Dim sText As String * 255           '//Buffer for EM_GETLINE call
Dim lStartTopLine   As Long         '//Starting top visible line
Dim lEndTopLine   As Long           '//Ending top visible line


lLine = lLine - 1                   '//Adjust since API calls are 0 based
lStartChar = SendMessage(rtHwnd, EM_LINEINDEX, lLine, ByVal 0&)

'//Cancel if line was deleted (this may occur only when coloring non-selected lines)
If lStartChar = -1 Then Exit Sub

SetWindowState bLocked              '//Lock the RTB from processing updates
bSkipColoring = True                '//Skip coloring routines while processing
lSelStart = Me.rtBox.SelStart       '//Get current caret position
lSelStartLen = Me.rtBox.SelLength   '//Get current selected text length

'//Fill variables upfront to be used during processing
lStartTopLine = SendMessage(rtHwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
lLineLength = SendMessage(rtHwnd, EM_LINELENGTH, lStartChar, ByVal 0&)
sText = Space(255): Call SendMessage(rtHwnd, EM_GETLINE, lLine, ByVal sText)
sLineText = UCase$(StripNulls(Left$(sText, lLineLength))): sText = ""
sTrimmedLine = Trim$(sLineText)

'//Except for assigning the RTB's selected text color and bolded state, the remaining
'..code works with assigned variables, so processing is still fairly quick.
With Me.rtBox
    Select Case Left$(sTrimmedLine, 1)
        Case ";"
            '//Commented line, so only make it eCOMMENT color
            SendMessage rtHwnd, EM_SETSEL, ByVal lStartChar, ByVal lStartChar + lLineLength
            .SelColor = eCOMMENT
            .SelBold = False

        Case "["
            '//Section Header, see if it matches Header list from LoadTag procedure
            lConstStart = InStr(1, sLineText, "[", vbTextCompare)
            lConstEnd = InStr(1, sLineText, "]", vbTextCompare)
            If lConstEnd > lConstStart Then
                With SecTags
                    For x = 0 To UBound(.Tag)
                        sTrimmedLine2 = Left$(sTrimmedLine, .Length(x))
                        If sTrimmedLine2 = UCase$(.Tag(x)) Then
                            bTagFound = True
                            Exit For
                        End If
                    Next x
                End With
                If bTagFound Then
                    Y = InStr(1, sLineText, SecTags.Tag(x), vbTextCompare)
                    '//Bold and color section header text
                    SendMessage rtHwnd, EM_SETSEL, ByVal lStartChar + Y - 1, ByVal lStartChar + Y - 1 + SecTags.Length(x)
                    .SelBold = True
                    .SelColor = Setup.SecTagColor
                    .SelText = SecTags.Tag(x)
                    '//Comment any remaining text after section header closes
                    If (lLineLength - lConstEnd) > 0 Then
                        SendMessage rtHwnd, EM_SETSEL, ByVal lStartChar + lConstEnd, ByVal lStartChar + lLineLength
                        .SelBold = False
                        .SelColor = eCOMMENT
                    End If
                Else
                    SendMessage rtHwnd, EM_SETSEL, ByVal lStartChar, ByVal lStartChar + lLineLength
                    .SelColor = vbBlack
                    .SelBold = False
                End If
            Else
                SendMessage rtHwnd, EM_SETSEL, ByVal lStartChar, ByVal lStartChar + lLineLength
                .SelColor = vbBlack
                .SelBold = False
            End If
            
        Case Else
            '//Keys
            With KeyTags
                For x = 0 To UBound(.Tag)
                    sTrimmedLine2 = Left$(sTrimmedLine, .Length(x))
                    If sTrimmedLine2 = UCase$(.Tag(x)) Then
                        bTagFound = True
                        Exit For
                    End If
                Next x
            End With
            If bTagFound Then
                '//Tag was found, lets process it
                Y = InStr(1, sLineText, KeyTags.Tag(x), vbTextCompare)
                '//Color tag
                SendMessage rtHwnd, EM_SETSEL, ByVal lStartChar + Y - 1, ByVal lStartChar + Y - 1 + KeyTags.Length(x)
                .SelBold = False
                .SelColor = Setup.KeyTagColor
                .SelText = KeyTags.Tag(x)      '//Makes tag the case you specifed in LoadTags
                '//Color any remaining text vbBlack
                If (lLineLength - Y + 1 - KeyTags.Length(x)) > 0 Then
                    SendMessage rtHwnd, EM_SETSEL, ByVal lStartChar + Y - 1 + KeyTags.Length(x), ByVal lStartChar + lLineLength
                    .SelBold = False
                    .SelColor = vbBlack
                    '//Now check for value tags. This also ensures tags are in the right place
                    '..after the = sign, but allows for spaces and tabs between them for formatting.
                    If KeyTags.HasValTag(x) Then
                        sTrimmedLine = Replace(sLineText, vbTab, "")
                        sTrimmedLine = Replace(sLineText, " ", "")
                        Y = InStr(1, sTrimmedLine, "=", vbTextCompare)
                        If Len(sTrimmedLine) > Y Then
                            sValueText = Right$(sTrimmedLine, Len(sTrimmedLine) - Y)
                            For z = 0 To UBound(ValTags.Tag)
                                If ValTags.DataType(z) = KeyTags.DataType(x) Then
                                    If Left$(sValueText, ValTags.Length(z)) = UCase$(ValTags.Tag(z)) Then
                                        lConstStart = InStr(1, sLineText, UCase$(ValTags.Tag(z)))
                                        '//Color value tag
                                        SendMessage rtHwnd, EM_SETSEL, ByVal lStartChar + lConstStart - 1, ByVal lStartChar + lConstStart - 1 + ValTags.Length(z)
                                        .SelColor = Setup.ValTagColor
                                        .SelText = ValTags.Tag(z)   '//Makes tag the case you specifed in LoadTags
                                        If Not ValTags.TextAfter(z) Then
                                            '//Don't allow text after so color it red
                                            SendMessage rtHwnd, EM_SETSEL, ByVal lStartChar + lConstStart + ValTags.Length(z) - 1, ByVal lStartChar + lLineLength
                                            .SelColor = eERROR
                                        End If
                                        Exit For
                                    End If
                                End If
                            Next z
                        End If
                    End If
                End If
            Else
                '//No tags found so color line vbBlack
                SendMessage rtHwnd, EM_SETSEL, ByVal lStartChar, ByVal lStartChar + lLineLength
                .SelBold = False
                .SelColor = vbBlack
            End If
    End Select
    '//Return cursor to starting selection
    SendMessage rtHwnd, EM_SETSEL, ByVal lSelStart, ByVal lSelStart + lSelStartLen
End With

'//This mess seems to have whipped (or sudo-disabled) the richtextboxes autoscroll
'..into submission when coloring a PREVLINENUMBER that is out of view. It basically
'..checks the difference between where we are now and where we started, and returns
'..the scrolling position to where we began if necessary.
lEndTopLine = SendMessage(rtHwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
If lStartTopLine <> lEndTopLine Then
    x = IIf(lStartTopLine < lEndTopLine, -(lEndTopLine - lStartTopLine), lStartTopLine - lEndTopLine)
    Call SendMessage(rtHwnd, EM_LINESCROLL, ByVal 0&, ByVal x)
End If

bSkipColoring = False       '//Continue coloring
SetWindowState bUnLocked    '//Refresh RTB

End Sub

Private Sub Comment()
'-----------------------------------------------------------------------------------
' Purpose   : Comments selected block in a RTB. Results mimic VB IDE comment button.
'-----------------------------------------------------------------------------------
On Error Resume Next            '//Can't break it, but better safe than sorry!
Dim x               As Long     '//For loop, then first selected char when complete
Dim Y               As Long     '//Final selected char when complete
Dim lStartChar      As Long     '//First char position in selection
Dim lEndChar        As Long     '//Last char position in selection
Dim lStartLine      As Long     '//First selected line
Dim lEndLine        As Long     '//Last selected line
Dim lLineStartChar  As Long     '//First char position in processing line
Dim lLineLen        As Long     '//Length of processing line
Dim bBeginOffLeft   As Byte     '//Set to 1 if cursor started off left margin
    SetWindowState bLocked
    With Me.rtBox
        lStartChar = .SelStart
        lEndChar = lStartChar + .SelLength
        lStartLine = SendMessage(rtHwnd, EM_LINEFROMCHAR, lStartChar, ByVal 0&)
        lLineStartChar = SendMessage(rtHwnd, EM_LINEINDEX, lStartLine, ByVal 0&)
        If InStr(1, .SelText, vbCrLf, vbTextCompare) Then
            lEndLine = SendMessage(rtHwnd, EM_LINEFROMCHAR, lEndChar - 1, ByVal 0&)
        Else
            lEndLine = SendMessage(rtHwnd, EM_LINEFROMCHAR, lEndChar + 1, ByVal 0&)
        End If
        If lStartChar <> lLineStartChar Then bBeginOffLeft = 1
        '//Comment all lines in the selected block
        For x = lStartLine To lEndLine
            lLineStartChar = SendMessage(rtHwnd, EM_LINEINDEX, x, ByVal 0&)
            lLineLen = SendMessage(rtHwnd, EM_LINELENGTH, lLineStartChar, ByVal 0&)
            SendMessage rtHwnd, EM_SETSEL, ByVal lLineStartChar, ByVal lLineStartChar + lLineLen
            .SelText = ";" & .SelText
        Next x
        '//Now return the caret and select length to where it began (The Hard Part!).
        If bBeginOffLeft Or (lStartChar <> lEndChar) Then
            Y = lEndChar + (lEndLine - lStartLine) + 1
        Else
            Y = lEndChar + (lEndLine - lStartLine)
        End If
        x = lStartChar + bBeginOffLeft
        '//Stop coloring process and set adjusted selection
        bSkipColoring = True
        SendMessage rtHwnd, EM_SETSEL, ByVal x, ByVal Y
        bSkipColoring = False
        .SetFocus
    End With
    SetWindowState bUnLocked
End Sub

Private Sub UnComment()
'--------------------------------------------------------------------------------------
' Purpose   : UnComments selected block in an RTB. Results mimic VB IDE comment button.
'--------------------------------------------------------------------------------------
On Error Resume Next            '//Can't break it, but better safe than sorry!
Dim x                   As Long
Dim Y                   As Long
Dim lStartChar          As Long
Dim lEndChar            As Long
Dim lStartLine          As Long
Dim lEndLine            As Long
Dim lLineStartChar      As Long
Dim lLineLen            As Long
Dim sSelLineText        As String
Dim lCharsRemoved       As Long
Dim bCaretAfterComm     As Byte
    SetWindowState bLocked
    With Me.rtBox
        lStartChar = .SelStart
        lEndChar = lStartChar + .SelLength
        lStartLine = SendMessage(rtHwnd, EM_LINEFROMCHAR, lStartChar, ByVal 0&)
        lLineStartChar = SendMessage(rtHwnd, EM_LINEINDEX, lStartLine, ByVal 0&)
        If InStr(1, .SelText, vbCrLf, vbTextCompare) Then
            lEndLine = SendMessage(rtHwnd, EM_LINEFROMCHAR, lEndChar - 1, ByVal 0&)
        Else
            lEndLine = SendMessage(rtHwnd, EM_LINEFROMCHAR, lEndChar + 1, ByVal 0&)
        End If
        '//UnComment all affected lines in the selected block
        For x = lStartLine To lEndLine
            lLineStartChar = SendMessage(rtHwnd, EM_LINEINDEX, x, ByVal 0&)
            lLineLen = SendMessage(rtHwnd, EM_LINELENGTH, lLineStartChar, ByVal 0&)
            SendMessage rtHwnd, EM_SETSEL, ByVal lLineStartChar, ByVal lLineStartChar + lLineLen
            sSelLineText = Replace$(Replace$(.SelText, vbTab, ""), " ", "")
            If Left$(sSelLineText, 1) = ";" Then
                '//See if the 1st selected char is after the 1st removed comment.
                '..We need this info to determine new select start position.
                If lStartLine = x Then
                    Y = InStr(1, .SelText, ";", vbTextCompare) + lLineStartChar - 1
                    If lStartChar > Y Then
                        bCaretAfterComm = 1
                    Else
                        '//Increment lCharsRemoved if comment is in the 1st lines selection
                        If lStartChar < lEndChar Then lCharsRemoved = lCharsRemoved + 1
                    End If
                ElseIf lEndLine = x Then
                    '//Increment lCharsRemoved if comment is in the last lines selection
                    Y = InStr(1, .SelText, ";", vbTextCompare)
                    If lEndChar - lLineStartChar - 1 > Y Then lCharsRemoved = lCharsRemoved + 1
                Else
                    '//Increment lCharsRemoved for all lines in the middle with leading comments
                    lCharsRemoved = lCharsRemoved + 1
                End If
                '//Set new text for line with comment removed
                sSelLineText = .SelText
                Y = InStr(1, sSelLineText, ";", vbTextCompare)
                .SelText = Left$(sSelLineText, Y - 1) & Right(sSelLineText, lLineLen - Y)
            End If
        Next x
        '//Now return the caret and select length to the corrected positions
        x = lStartChar - bCaretAfterComm
        bSkipColoring = True
        SendMessage rtHwnd, EM_SETSEL, ByVal x, ByVal lEndChar - lCharsRemoved - bCaretAfterComm
        bSkipColoring = False
        .SetFocus
    End With
    SetWindowState bUnLocked
End Sub
Private Function StripNulls(ByVal sString As String) As String
    sString = Replace(sString, vbTab, " ")
    sString = Replace(sString, vbLf, " ")
    sString = Replace(sString, vbCr, " ")
    StripNulls = sString
End Function

Private Function FileSave(Optional ByVal SaveAs = False) As Boolean
On Error GoTo Errs
Dim sFile As String
    If Len(Setup.Script) And Not SaveAs Then
        sFile = Setup.Script
        Me.rtBox.SaveFile Setup.Script, rtfText
        bChanged = False
    Else
        With Me.cd
            .Filter = "ReVive Update Script (*.rus)|*.rus" & _
                      "|Text Files (*.txt;*.rtf;*.doc)|*.txt;*.rtf;*.doc" & _
                      "|All Files (*.*)|*.*"
            .FileName = IIf(Len(Setup.AppShortName) = 0, "update.rus", Setup.AppShortName & ".rus")
            .DefaultExt = ".rus"
            .DialogTitle = "Select a script file name and location..."
            .Flags = cdlOFNFileMustExist Or cdlOFNExplorer _
                Or cdlOFNHideReadOnly Or cdlOFNPathMustExist _
                    Or cdlOFNShareAware
            .ShowSave
            sFile = .FileName
            If Dir(sFile, 39) <> "" And sFile <> Setup.Script Then
                If MsgBox("The file you have selected already exist. Do you want to overwrite it?    ", vbYesNo + vbQuestion, "File Exist") = vbNo Then
                    Exit Function
                End If
            End If
            Me.rtBox.SaveFile .FileName, rtfText
            bChanged = False
        End With
    End If
    Me.Caption = "ReVive Script Editor - " & sFile
    Setup.Script = sFile
    FileSave = True
Errs_Exit:
    Exit Function
Errs:
    If Err.Number <> 32755 Then '------------------------- 32755 = "Cancel was selected."
        MsgBox "The following error occured while attempting to save file  " & vbNewLine & _
                "'" & sFile & "'.   " & vbNewLine & vbNewLine & _
                Err.Description, vbExclamation, "Error Saving File"
    End If
    Resume Errs_Exit
End Function

Private Sub FilePrint()
On Error GoTo Errs
    With cd
        .Flags = cdlPDReturnDC + cdlPDNoPageNums
        If rtBox.SelLength = 0 Then
           .Flags = .Flags + cdlPDAllPages
        Else
           .Flags = .Flags + cdlPDSelection
        End If
        .ShowPrinter
        rtBox.SelPrint .hDC
    End With
Errs:
    If Err Then
        If Err.Number <> 32755 Then  '//32755 = "Cancel was selected."
            MsgBox Err.Number & ":  " & Err.Description, vbExclamation, "Error Printing Script"
        End If
        Exit Sub
    End If
End Sub

Private Sub FileOpen(Optional ByVal sFile As String)
On Error GoTo Errs
Dim m           As Integer
Dim lLineCount  As Long
Dim x           As Long
    If bChanged Then
        m = MsgBox("Do you wish to save the changes made to this script?   ", vbYesNoCancel + vbQuestion, "Save Changes?")
        If m = 6 Then
            If Not FileSave Then Exit Sub
        ElseIf m = 2 Then
            Exit Sub
        End If
    End If
    If Len(sFile) = 0 Then
        With Me.cd
            .HelpFile = App.path & "\ReVive.hlp"
            .HelpContext = 5
            .HelpCommand = cdlHelpContext
            .Filter = "ReVive Update Script (*.rus)|*.rus" & _
                      "|Text Files (*.txt;*.rtf;*.doc)|*.txt;*.rtf;*.doc" & _
                      "|All Files (*.*)|*.*"
            .FileName = ""
            .DefaultExt = ".rus"
            .DialogTitle = "Select the script file you wish to open..."
            .Flags = cdlOFNFileMustExist Or cdlOFNExplorer _
                Or cdlOFNHideReadOnly Or cdlOFNPathMustExist _
                    Or cdlOFNShareAware Or cdlOFNHelpButton
            .ShowOpen
            sFile = .FileName
        End With
        Setup.Script = sFile
    End If
    Screen.MousePointer = 11
    SetWindowState bLocked
    Me.rtBox.LoadFile sFile
    Me.Caption = "ReVive Script Editor - " & sFile
    lLineCount = SendMessage(rtHwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
    For x = -lLineCount To -1
        ColorLine Abs(x)
    Next x
Errs_Exit:
    bChanged = False
    PREVLINENUMBER = 1
    SetWindowState bUnLocked
    Screen.MousePointer = 0
    Exit Sub
Errs:
    If Err.Number <> 32755 Then '------------------------- 32755 = "Cancel was selected."
        MsgBox "The Script Editor could not open the selected file.  ", vbExclamation, "Error Opening File."
    End If
    Resume Errs_Exit
End Sub

Private Sub FileNew()
Dim m As Integer
On Error GoTo Errs
    If bChanged Then
        m = MsgBox("Do you wish to save the changes made to this script?   ", vbYesNoCancel + vbQuestion, "Save Changes?")
        If m = 6 Then
            If Not FileSave Then Exit Sub
        ElseIf m = 2 Then
            Exit Sub
        End If
    End If
    frmNew.Show 1
Errs:
    If Err Then Exit Sub
End Sub

Private Sub AddSectionTag(ByVal Tag As String, _
    Section As enumSections)
Static s As Long
    With SecTags
        ReDim Preserve .Tag(0 To s)
        ReDim Preserve .Length(0 To s)
        ReDim Preserve .Section(0 To s)
        .Tag(s) = Tag                   '//Actual tag
        .Length(s) = Len(.Tag(s))
        .Section(s) = Section
    End With
    s = s + 1
End Sub

Private Sub AddKeyTag(ByVal Tag As String, _
        ByVal Section As enumSections, ByVal DataType As enumDataTypes, _
        ByVal HasValTag As Boolean)
Static k As Long
    With KeyTags
        ReDim Preserve .Tag(0 To k)
        ReDim Preserve .Length(0 To k)
        ReDim Preserve .DataType(0 To k)
        ReDim Preserve .Section(0 To k)
        ReDim Preserve .HasValTag(0 To k)
        .Tag(k) = Tag                   '//Actual tag
        .DataType(k) = DataType         '//Datatype accepted for tag
        .Section(k) = Section           '//Section tag belongs in
        .HasValTag(k) = HasValTag
        .Length(k) = Len(.Tag(k))
    End With
    k = k + 1
End Sub

Private Sub AddValueTag(ByVal Tag As String, _
        ByVal DataType As enumDataTypes, ByVal AllowTextAfter As Boolean)
Static v As Long
    With ValTags
        ReDim Preserve .Tag(0 To v)
        ReDim Preserve .Length(0 To v)
        ReDim Preserve .DataType(0 To v)
        ReDim Preserve .TextAfter(0 To v)
        .Tag(v) = Tag                   '//Actual tag
        .Length(v) = Len(.Tag(v))
        .DataType(v) = DataType
        .TextAfter(v) = AllowTextAfter
    End With
    v = v + 1
End Sub

Private Sub TestScript()
'---------------------------------------------------------------------------------------------
' Purpose   : Tests script for the following errors: (More or less in this order)
'             1. Invalid Section Headers not identified in LoadTag procedure.
'             2. Duplication of Section Headers within script.
'             3. Invalid Section Keys not identified in LoadTag procedure.
'             4. Duplication of Section Keys within same Section.
'             5. Placement of Section Keys in wrong Section as defined in LoadTag procedure.
'             6. Zero length Values for any Section Keys.
'             7. Invalid datatype for each Section Key Value as defined in LoadTag procedure.
'             8. Duplication of Description= values (each must be unique).
'             9. When selected, checks to ensure filenames meet DOS 8.3 format for pre-NT OS's
'             10. Unrecognized text throughout script that is not commented.
'---------------------------------------------------------------------------------------------
Dim line            As Long
Dim x               As Long
Dim sLineText       As String
Dim sKey            As String
Dim sKeyList        As String
Dim sValue          As String
Dim lTotalLines     As Long
Dim lLineLength     As Long
Dim lStartChar      As Long
Dim Section         As enumSections
Dim sSection        As String
Dim sSectionList    As String
Dim sDescriptList   As String
Dim sInstallList    As String
Dim sMessage        As String
Dim bError          As Boolean
Dim lSelStart       As Long
Dim lSelStartLen    As Long
Dim lStartTopLine   As Long
Dim lEndTopLine     As Long

SetWindowState bLocked
bSkipColoring = True

lStartTopLine = SendMessage(rtHwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
lTotalLines = SendMessage(rtHwnd, EM_GETLINECOUNT, ByVal 0&, ByVal 0&)
Section = eNONE
sSection = "an unknown"
lSelStart = Me.rtBox.SelStart
lSelStartLen = Me.rtBox.SelLength

sSectionList = "|"
sKeyList = "|"
sDescriptList = "|"
sInstallList = "|"

For line = 1 To lTotalLines
    lStartChar = SendMessage(rtHwnd, EM_LINEINDEX, line - 1, ByVal 0&)
    lLineLength = SendMessage(rtHwnd, EM_LINELENGTH, lStartChar, ByVal 0&)
    SendMessage rtHwnd, EM_SETSEL, ByVal lStartChar, ByVal lStartChar + lLineLength
    sLineText = Me.rtBox.SelText
    sLineText = Trim$(StripNulls(Left$(sLineText, lLineLength)))
    If Len(sLineText) = 0 Then GoTo SkipLine
    Select Case Left$(sLineText, 1)
        Case ";", "'"
            GoTo SkipLine
        
        Case "["
            For x = 0 To UBound(SecTags.Tag)
                If Left$(sLineText, SecTags.Length(x)) = SecTags.Tag(x) Then
                    Section = SecTags.Section(x)
                    sSection = SecTags.Tag(x)
                    '//Verify section header is not duplicated
                    If InStr(1, sSectionList, "|" & sSection & "|") Then
                        bError = True
                        sMessage = "'" & sSection & "' section header is duplicated. Only first occurrence recognized.   "
                        Exit For
                    End If
                    sKeyList = ""
                    sSectionList = sSectionList & sSection & "|"
                    GoTo SkipLine
                End If
            Next x
            If bError Then Exit For '//Get out if duplicate section found.
            bError = True
            sMessage = "'" & Trim$(sLineText) & "' is not a recognized section header.   "
            Exit For
            
        Case Else
            sLineText = Replace(sLineText, vbTab, "")
            sLineText = Trim$(sLineText)
            x = InStr(1, sLineText, "=", vbTextCompare)
            If x = 0 Then
                bError = True
                sMessage = "Unrecognized and un-commented text.   "
                Exit For
            Else
                sKey = Left$(sLineText, x)
                sValue = Trim$(Right$(sLineText, Len(sLineText) - x))
                For x = 0 To UBound(KeyTags.Tag)
                    If sKey = KeyTags.Tag(x) Then
                        If KeyTags.Section(x) = Section Then
                            '//Verify Key is not duplicated in current section.
                            If InStr(1, sKeyList, "|" & sKey & "|") Then
                                bError = True
                                sMessage = "'" & Trim$(sKey) & "' key duplicated in '" & sSection & _
                                "' section. Only first occurrence recognized.   "
                                Exit For
                            Else
                                sKeyList = sKeyList & sKey & "|"
                            End If
                            '//Verify file Description= keys do not have duplicate values anywhere in script
                            If sKey = "Description=" Then
                                If InStr(1, sDescriptList, "|" & sValue & "|") Then
                                    bError = True
                                    sMessage = "'" & sValue & "' duplicated. Each '" & sKey & "' key " & _
                                    "must have a unique value.  "
                                    Exit For
                                Else
                                    sDescriptList = sDescriptList & sValue & "|"
                                End If
                            End If
                            '//Verify InstallPath= keys do not have duplicate values anywhere in script
                            If sKey = "InstallPath=" Then
                                If InStr(1, sInstallList, "|" & sValue & "|") Then
                                    bError = True
                                    sMessage = "'" & sValue & "' duplicated. Each '" & sKey & "' key " & _
                                    "must have a unique value.  "
                                    Exit For
                                Else
                                    sInstallList = sInstallList & sValue & "|"
                                End If
                            End If
                            Exit For
                        End If
                    End If
                Next x
                If bError Then Exit For '//Get out if duplicate key found.
                If x = UBound(KeyTags.Tag) + 1 Then
                    bError = True
                    sMessage = "'" & Trim$(sKey) & "' is not a recognized key in " & sSection & " section.   "
                    Exit For
                Else
                    If Section <> KeyTags.Section(x) Then
                        bError = True
                        sMessage = "'" & Trim$(sKey) & "' is not a recognized key in " & sSection & " section.   "
                        Exit For
                    ElseIf Len(sValue) = 0 Then
                        bError = True
                        sMessage = "Key '" & Trim$(sKey) & "' has no assigned value.   "
                        Exit For
                    ElseIf Not IsValueValid(sValue, KeyTags.DataType(x)) Then
                        bError = True
                        sMessage = "'" & Trim$(sValue) & "' is not a valid datatype for '" & sKey & "' key.   " & _
                                vbNewLine & vbNewLine & sExpValue
                        Exit For
                    End If
                End If
            End If

    End Select
    
SkipLine:
Next line

If bError Then
    SetWindowState bUnLocked
    MsgBox sMessage & vbNewLine & vbNewLine & "If you wish to ignore this " & _
            "line while testing please comment it.   ", vbExclamation, "Script Error"
Else
    '//Return cursor to starting selection
    SendMessage rtHwnd, EM_SETSEL, ByVal lSelStart, ByVal lSelStart + lSelStartLen
    '//This mess seems to have whipped (or sudo-disabled) the richtextboxes autoscroll
    '..into submission when coloring a PREVLINENUMBER that is out of view. It basically
    '..checks the difference between where we are now and where we started, and returns
    '..the scrolling position to where we began if necessary.
    lEndTopLine = SendMessage(rtHwnd, EM_GETFIRSTVISIBLELINE, ByVal 0&, ByVal 0&)
    If lStartTopLine <> lEndTopLine Then
        x = IIf(lStartTopLine < lEndTopLine, -(lEndTopLine - lStartTopLine), lStartTopLine - lEndTopLine)
        Call SendMessage(rtHwnd, EM_LINESCROLL, ByVal 0&, ByVal x)
    End If
    MsgBox "No apparent errors were detected in script.   ", vbInformation, "Test Complete"
    SetWindowState bUnLocked
End If
bSkipColoring = False
End Sub

Private Function IsValueValid(ByVal sValue As String, ByVal DataType As enumDataTypes) As Boolean
On Error GoTo Errs
Dim b As Boolean
Dim l As Long
Dim i As Integer
Dim s As String
    sValue = LCase$(sValue)
    Select Case DataType
        Case eBOOLEAN
            sExpValue = "Expected value: True, False, 0 or 1.  "
            '//I am looking for specific characters in this app, but a cBOOL will normally work.
            If sValue <> "true" And sValue <> "false" And sValue <> "0" And sValue <> "1" Then
                GoTo Errs
            End If
        Case eINTEGER
            sExpValue = "Expected value: -32,768 to 32,767.  "
            i = CInt(sValue)
        Case eLONG
            sExpValue = "Expected value: -2,147,483,648 to 2,147,483,647.  "
            l = CLng(sValue)
        Case eWEB
            sExpValue = "Expected value: Must begin with http://, ftp://, \\ or a drive letter followed by :.  "
            If Left$(sValue, 7) <> "http://" And Left$(sValue, 6) <> "ftp://" And _
                Left$(sValue, 2) <> "\\" And Mid$(sValue, 2, 1) <> ":" Then
                GoTo Errs
            End If
        Case eVERSION
            sExpValue = "Expected value: A valid version number formatted like '0.0.0.0', containing only characters 0 - 9 and 3 decimal points(.)."
            If Not IsVersionValid(sValue) Then GoTo Errs
        Case ePATH
            sExpValue = "Expected value: A valid path, including filename, meeting Windows naming convention rules.  "
            If Not IsLocalPathValid(sValue, False) Then GoTo Errs
    End Select
    IsValueValid = True
Errs:
    Exit Function
End Function

Private Function IsLocalPathValid(ByVal sPath As String, _
    Optional ByVal VerifyDriveExist As Boolean = False) As Boolean
'---------------------------------------------------------------------
' Purpose   : Checks if sPath will pass Windows file and folder naming
'             convention rules before attempting to create it.
'---------------------------------------------------------------------
On Error GoTo Errs
Dim folder()    As String
Dim badchars()  As String
Dim reswords()  As String
Dim path        As String
Dim i           As Byte
Dim x           As Long
Dim file        As String

'//Exit if \\ is anywhere in path
If InStr(1, sPath, "\\", vbTextCompare) Then Exit Function

'//Build invalid char and reserved word lists
badchars = Split("| < > * ? / : " & Chr(34), " ")
reswords = Split("CON PRN AUX CLOCK$ NUL LPT1 LPT2 LPT3 LPT4 LPT5 LPT6 LPT7 " & _
                    "LPT8 LPT9 COM1 COM2 COM3 COM4 COM5 COM6 COM7 COM8 COM9", " ")

'//Replace contants with C: for this application
If InStr(1, sPath, "<", vbTextCompare) Then
    With ValTags
        For x = 0 To UBound(.Tag)
            If .DataType(x) = ePATH Then
                If InStr(1, sPath, .Tag(x), vbTextCompare) Then
                    sPath = Replace(sPath, .Tag(x), "C:")
                    Exit For
                End If
            End If
        Next x
    End With
End If

path = Right$(sPath, Len(sPath) - InStrRev(sPath, "\"))
x = InStrRev(path, ".")
If x = 0 Then
    '//Verify destination filename is included in InstallPath
    Exit Function
Else
    file = Left$(path, x - 1)
End If

'//Parse drive and folders
folder = Split(sPath, "\")

'//Extract drive and check if valid
path = LCase$(folder(0))
If VerifyDriveExist Then
    If Dir$(path, 63) = "" Then Exit Function
Else
    For x = 97 To 122
        If path = Chr(x) & ":" Then Exit For
    Next x
    If x = 123 Then Exit Function
End If

'//Check for invalid folder characters and reserved words
For i = 1 To UBound(folder)
    For x = 0 To 7
        If InStr(1, folder(i), badchars(x)) Then Exit Function
    Next x
    For x = 0 To 22
        If UCase$(folder(i)) = reswords(x) Then Exit Function
    Next x
Next i

'//If checking for Win95 - ME compatibility, check for filenames
'..longer than 8 characters (DOS 8.3 format)
If Setup.TestForFileNames Then
    If Len(file) > 8 Then
        MsgBox "InstallPath file name " & UCase$(file) & " is longer than 8 characters. When    " & Chr(13) & _
               "updating in-use files on clients running pre-NT based operating systems,    " & Chr(13) & _
               "Windows will rename this file " & Left$(UCase$(file), 6) & "~1, and will therefore inaccurately    " & Chr(13) & _
               "update the clients system. Please consider shortening the file name to    " & Chr(13) & _
               "no more than 8 characters to eliminate this possible compatibility issue.    ", vbInformation, "Compatibility Issue"
    End If
End If
        
IsLocalPathValid = True

Errs:
    If Err Then MsgBox Err.Description: Exit Function
    
End Function

Private Sub Form_Unload(Cancel As Integer)
On Error GoTo Errs
Dim m As Long
    If bChanged Then
        m = MsgBox("Do you wish to save the changes made to this script?   ", vbYesNoCancel + vbQuestion, "Save Changes?")
        If m = 6 Then
            If Not FileSave Then Cancel = 1: Exit Sub
        ElseIf m = 2 Then
            Cancel = True
            Exit Sub
        End If
    End If
    WinHelp Me.hwnd, App.path & "ReVive.hlp", HELP_QUIT, "" '--- Close any open help topics
Errs:
    Call SetWindowLong(Me.hwnd, GWL_WNDPROC, OldWindowProc)
    Set frmMain = Nothing
End Sub

Private Function FillTreeView()

With Me.tv
    Set .ImageList = Me.iml
    .Nodes.Add , , "[Setup]", "[Setup]", 6
    .Nodes.Add "[Setup]", 4, "AdminRequired=", "AdminRequired", 12
    .Nodes.Add "[Setup]", 4, "AppShortName=", "AppShortName", 12
    .Nodes.Add "[Setup]", 4, "AppLongName=", "AppLongName", 12
    .Nodes.Add "[Setup]", 4, "ForceReboots=", "ForceReboots", 12
    .Nodes.Add "[Setup]", 4, "LaunchIfKilled=", "LaunchIfKilled", 12
    .Nodes.Add "[Setup]", 4, "NotifyIcon=", "NotifyIcon", 12
    .Nodes.Add "[Setup]", 4, "RegRISFiles=", "RegRISFiles", 12
    .Nodes.Add "[Setup]", 4, "ScriptURLAlt=", "ScriptURLAlt", 12
    .Nodes.Add "[Setup]", 4, "ScriptURLPrim=", "ScriptURLPrim", 12
    .Nodes.Add "[Setup]", 4, "ShowFileIcons=", "ShowFileIcons", 12
    .Nodes.Add "[Setup]", 4, "UpdateAppClass=", "UpdateAppClass", 12
    .Nodes.Add "[Setup]", 4, "UpdateAppKill=", "UpdateAppKill", 12
    .Nodes.Add "[Setup]", 4, "UpdateAppTitle=", "UpdateAppTitle", 12
    
    .Nodes.Add , , "[Files]", "[Files]", 13
    .Nodes.Add "[Files]", 4, "Description=", "Description", 12
    .Nodes.Add "[Files]", 4, "DownloadURL=", "DownloadURL", 12
    .Nodes.Add "[Files]", 4, "FileSize=", "FileSize", 12
    .Nodes.Add "[Files]", 4, "InstallPath=", "InstallPath", 12
    .Nodes.Add "[Files]", 4, "MustExist=", "MustExist", 12
    .Nodes.Add "[Files]", 4, "MustUpdate=", "MustUpdate", 12
    .Nodes.Add "[Files]", 4, "UpdateVersion=", "UpdateVersion", 12
    .Nodes.Add "[Files]", 4, "UpdateMessage=", "UpdateMessage", 12
    
    .Nodes.Add , , "Directory Constants", "Directory Constants", 14
    .Nodes.Add "Directory Constants", 4, "<ap>", "<ap>", 12
    .Nodes.Add "Directory Constants", 4, "<cf>", "<cf>", 12
    .Nodes.Add "Directory Constants", 4, "<commondesktop>", "<commondesktop>", 12
    .Nodes.Add "Directory Constants", 4, "<commonstartmenu>", "<commonstartmenu>", 12
    .Nodes.Add "Directory Constants", 4, "<pf>", "<pf>", 12
    .Nodes.Add "Directory Constants", 4, "<sp>", "<sp>", 12
    .Nodes.Add "Directory Constants", 4, "<sys>", "<sys>", 12
    .Nodes.Add "Directory Constants", 4, "<temp>", "<temp>", 12
    .Nodes.Add "Directory Constants", 4, "<userdesktop>", "<userdesktop>", 12
    .Nodes.Add "Directory Constants", 4, "<userstartmenu>", "<userstartmenu>", 12
    .Nodes.Add "Directory Constants", 4, "<win>", "<win>", 12
    .Nodes(1).Expanded = True
    .Nodes(15).Expanded = True
    .Nodes(24).Expanded = True
    .Nodes(1).Bold = True
    .Nodes(15).Bold = True
    .Nodes(24).Bold = True
    .Nodes(1).Selected = True
End With
End Function

Private Sub tv_DblClick()
Dim s As String
    s = Me.tv.SelectedItem.Key
    If Left(s, 1) = "[" Or s = "Directory Constants" Then Exit Sub
    Me.rtBox.SetFocus
    Me.Font.name = "Courier New"
    SendKeys s
    If Left(s, 1) = "<" Then
        SendKeys "\"
    Else
        SendKeys vbTab
        If Me.TextWidth(s) <= 1160 Then SendKeys vbTab
    End If
    Me.Font.name = "MS Sans Serif"
End Sub

Private Sub tv_NodeClick(ByVal Node As MSComctlLib.Node)
    Call DisplayTip(Node.Text, Me)
End Sub

Private Sub SetTVColor()
Dim lStyle              As Long
Dim TVNode              As Node
Const GWL_STYLE         As Long = -16&
Const TVM_SETBKCOLOR    As Long = 4381&
Const TVS_HASLINES      As Long = 2&
    For Each TVNode In tv.Nodes
        TVNode.BackColor = 15790320
    Next
    Call SendMessage(tv.hwnd, TVM_SETBKCOLOR, 0, ByVal 15790320)
    lStyle = GetWindowLong(tv.hwnd, GWL_STYLE)
    Call SetWindowLong(tv.hwnd, GWL_STYLE, lStyle And (Not TVS_HASLINES))
    Call SetWindowLong(tv.hwnd, GWL_STYLE, lStyle)
End Sub

Private Function InIDE() As Boolean
On Error GoTo Errs
    Debug.Print 1 / 0
    Exit Function
Errs:
    InIDE = True
End Function


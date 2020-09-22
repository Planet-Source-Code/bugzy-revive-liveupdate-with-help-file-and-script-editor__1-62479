VERSION 5.00
Object = "{3B7C8863-D78F-101B-B9B5-04021C009402}#1.2#0"; "RICHTX32.OCX"
Begin VB.Form frmNew 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "New Script File"
   ClientHeight    =   7065
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   8580
   HelpContextID   =   6
   Icon            =   "frmNew.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   7065
   ScaleWidth      =   8580
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Caption         =   "Application Shutdown Settings"
      Height          =   1995
      Left            =   180
      TabIndex        =   25
      Top             =   3270
      Width           =   8205
      Begin VB.ComboBox cmbDefConst 
         Height          =   315
         ItemData        =   "frmNew.frx":058A
         Left            =   2370
         List            =   "frmNew.frx":05AC
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Tag             =   "LaunchIfKilled"
         Top             =   1440
         Width           =   1845
      End
      Begin VB.TextBox txtAppTitle 
         Height          =   315
         Left            =   2370
         TabIndex        =   9
         Tag             =   "UpdateAppTitle"
         Top             =   720
         Width           =   5595
      End
      Begin VB.TextBox txtAppClass 
         Height          =   315
         Left            =   2370
         TabIndex        =   10
         Tag             =   "UpdateAppClass"
         Text            =   "ThunderRT6Main"
         Top             =   1080
         Width           =   5595
      End
      Begin VB.TextBox txtLaunch 
         Height          =   315
         Left            =   4230
         TabIndex        =   12
         Tag             =   "LaunchIfKilled"
         Text            =   "\"
         Top             =   1440
         Width           =   3735
      End
      Begin VB.CheckBox chkUpdateAppKill 
         Height          =   195
         Left            =   2370
         TabIndex        =   8
         Tag             =   "UpdateAppKill"
         Top             =   420
         Width           =   195
      End
      Begin VB.Label lbl 
         Caption         =   "UpdateAppTitle="
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
         Left            =   240
         TabIndex        =   29
         Top             =   770
         Width           =   2025
      End
      Begin VB.Label lbl 
         Caption         =   "UpdateAppClass="
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
         Index           =   8
         Left            =   240
         TabIndex        =   28
         Top             =   1130
         Width           =   2025
      End
      Begin VB.Label lbl 
         Caption         =   "LaunchIfKilled="
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
         Index           =   9
         Left            =   240
         TabIndex        =   27
         Top             =   1490
         Width           =   2025
      End
      Begin VB.Label lbl 
         Caption         =   "UpdateAppKill="
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
         Index           =   10
         Left            =   240
         TabIndex        =   26
         Top             =   410
         Width           =   1725
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Script Basics"
      Height          =   2955
      Left            =   180
      TabIndex        =   14
      Top             =   180
      Width           =   8205
      Begin VB.CheckBox chkRegRISFiles 
         Height          =   195
         Left            =   7740
         TabIndex        =   30
         Tag             =   "RegRISFiles"
         Top             =   730
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.CheckBox chkAdmin 
         Height          =   195
         Left            =   2370
         TabIndex        =   0
         Tag             =   "AdminRequired"
         Top             =   405
         Width           =   195
      End
      Begin VB.CheckBox chkReboot 
         Height          =   195
         Left            =   2370
         TabIndex        =   1
         Tag             =   "ForceReboots"
         Top             =   730
         Width           =   195
      End
      Begin VB.TextBox txtShort 
         Height          =   315
         Left            =   2370
         MaxLength       =   15
         TabIndex        =   3
         Tag             =   "AppShortName"
         Top             =   990
         Width           =   1755
      End
      Begin VB.TextBox txtLong 
         Height          =   315
         Left            =   2370
         TabIndex        =   4
         Tag             =   "AppLongName"
         Top             =   1350
         Width           =   5595
      End
      Begin VB.TextBox txtPrim 
         Height          =   315
         Left            =   2370
         TabIndex        =   5
         Tag             =   "ScriptURLPrim"
         Top             =   1710
         Width           =   5595
      End
      Begin VB.TextBox txtAlt 
         Height          =   315
         Left            =   2370
         TabIndex        =   6
         Tag             =   "ScriptURLAlt"
         Top             =   2070
         Width           =   5595
      End
      Begin VB.TextBox txtNotifyIcon 
         Height          =   315
         Left            =   2370
         TabIndex        =   7
         Tag             =   "NotifyIcon"
         Top             =   2430
         Width           =   5595
      End
      Begin VB.CheckBox chkShowIcons 
         Height          =   195
         Left            =   7740
         TabIndex        =   2
         Tag             =   "ShowFileIcons"
         Top             =   405
         Value           =   1  'Checked
         Width           =   195
      End
      Begin VB.Label lbl 
         Caption         =   "RegRISFiles="
         BeginProperty Font 
            Name            =   "Courier New"
            Size            =   9.75
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00000000&
         Height          =   225
         Index           =   12
         Left            =   5640
         TabIndex        =   31
         Top             =   737
         Width           =   1725
      End
      Begin VB.Label lbl 
         Caption         =   "AdminRequired="
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
         Left            =   240
         TabIndex        =   24
         Top             =   390
         Width           =   1725
      End
      Begin VB.Label lbl 
         Caption         =   "ForceReboots="
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
         Left            =   240
         TabIndex        =   23
         Top             =   737
         Width           =   1725
      End
      Begin VB.Label lbl 
         Caption         =   "AppShortName="
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
         Left            =   240
         TabIndex        =   22
         Top             =   1050
         Width           =   1725
      End
      Begin VB.Label lbl 
         Caption         =   "AppLongName="
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
         Left            =   240
         TabIndex        =   21
         Top             =   1425
         Width           =   1725
      End
      Begin VB.Label lbl 
         Caption         =   "ScriptURLPrim="
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
         Left            =   240
         TabIndex        =   20
         Top             =   1770
         Width           =   1725
      End
      Begin VB.Label lbl 
         Caption         =   "ScriptURLAlt="
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
         Left            =   240
         TabIndex        =   19
         Top             =   2115
         Width           =   1725
      End
      Begin VB.Label lbl 
         Caption         =   "NotifyIcon="
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
         Left            =   240
         TabIndex        =   18
         Top             =   2460
         Width           =   2025
      End
      Begin VB.Label lbl 
         Caption         =   "ShowFileIcons="
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
         Index           =   11
         Left            =   5640
         TabIndex        =   16
         Top             =   390
         Width           =   1725
      End
   End
   Begin VB.CommandButton cmdHelp 
      Cancel          =   -1  'True
      Caption         =   "Help"
      Height          =   375
      Left            =   5490
      TabIndex        =   17
      Top             =   6570
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6480
      TabIndex        =   15
      Top             =   6570
      Width           =   915
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   7470
      TabIndex        =   13
      Top             =   6570
      Width           =   915
   End
   Begin RichTextLib.RichTextBox rtTip 
      Height          =   945
      Left            =   180
      TabIndex        =   32
      TabStop         =   0   'False
      Top             =   5460
      Width           =   8215
      _ExtentX        =   14499
      _ExtentY        =   1667
      _Version        =   393217
      BackColor       =   14286847
      Enabled         =   -1  'True
      ReadOnly        =   -1  'True
      ScrollBars      =   2
      AutoVerbMenu    =   -1  'True
      TextRTF         =   $"frmNew.frx":0611
   End
End
Attribute VB_Name = "frmNew"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function WNetGetUser Lib "mpr.dll" Alias "WNetGetUserA" (ByVal lpName As String, ByVal lpUserName As String, lpnLength As Long) As Long

Private Sub chkAdmin_GotFocus()
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub chkReboot_GotFocus()
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub


Private Sub chkRegRISFiles_GotFocus()
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub chkShowIcons_GotFocus()
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub chkUpdateAppKill_GotFocus()
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub cmbDefConst_GotFocus()
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub cmdHelp_Click()
    WinHelp Me.hwnd, App.path & "\ReVive.hlp", HELP_CONTEXT, CLng(6)
End Sub

Private Sub Form_Load()
Dim x As Byte
    For x = 0 To 12
        Me.lbl(x).ForeColor = Setup.KeyTagColor
    Next x
    Me.txtPrim.Text = Setup.DefaultWeb
    Me.txtAlt.Text = Setup.DefaultWeb
    Me.txtNotifyIcon.Text = Setup.DefaultWeb
    With Me.cmbDefConst
        .Text = Setup.DefaultConst '--- Set default script constant
        .ForeColor = Setup.ValTagColor
    End With
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub cmdOK_Click()
Dim s As String
Dim x As Long
    s = vbNewLine
    s = s & ";Created by " & UserName & " on " & Format(Date, "dd mmm yyyy")
    s = s & vbNewLine
    s = s & ";" & Me.txtShort.Text & " Remote ReVive Update Script"
    s = s & vbNewLine & vbNewLine
    s = s & "[Setup]"
    s = s & vbNewLine
    s = s & vbTab & "AdminRequired=" & vbTab & CBool(Me.chkAdmin.Value)
    s = s & vbNewLine
    s = s & vbTab & "ForceReboots=" & vbTab & CBool(Me.chkReboot.Value)
    s = s & vbNewLine
    s = s & vbTab & "ShowFileIcons=" & vbTab & CBool(Me.chkShowIcons.Value)
    s = s & vbNewLine
    s = s & vbTab & "RegRISFiles=" & vbTab & CBool(Me.chkRegRISFiles.Value)
    s = s & vbNewLine
    s = s & vbTab & "AppShortName=" & vbTab & Me.txtShort.Text
    s = s & vbNewLine
    s = s & vbTab & "AppLongName=" & vbTab & Me.txtLong.Text
    s = s & vbNewLine
    s = s & vbTab & "ScriptURLPrim=" & vbTab & Me.txtPrim.Text
    s = s & vbNewLine
    s = s & vbTab & "ScriptURLAlt=" & vbTab & Me.txtAlt.Text
    s = s & vbNewLine
    s = s & vbTab & "NotifyIcon=" & vbTab & vbTab & Me.txtNotifyIcon.Text
    s = s & vbNewLine
    s = s & vbTab & "UpdateAppKill=" & vbTab & CBool(Me.chkUpdateAppKill.Value)
    s = s & vbNewLine
    s = s & vbTab & "UpdateAppTitle=" & vbTab & Me.txtAppTitle.Text
    s = s & vbNewLine
    s = s & vbTab & "UpdateAppClass=" & vbTab & Me.txtAppClass.Text
    s = s & vbNewLine
    s = s & vbTab & "LaunchIfKilled=" & vbTab & Me.cmbDefConst.Text & Me.txtLaunch.Text
    s = s & vbNewLine
    Setup.Script = ""
    Setup.AppShortName = LCase$(Me.txtShort.Text)
    With frmMain
        .Caption = "ReVive Script Editor"
        .rtBox.Text = s
        For x = 1 To 18
            .ColorLine x
        Next x
        .rtBox.SelStart = Len(.rtBox.TextRTF)
    End With
    Me.Hide
    If MsgBox("Would you like to begin adding update files to this script?   ", vbYesNo + vbQuestion, "Script Editor") = vbYes Then
        frmInsert.Show 1
    End If
    Unload Me
End Sub

Private Sub SelectAllText()
On Error Resume Next
    Screen.ActiveControl.SelStart = 0
    Screen.ActiveControl.SelLength = Len(Screen.ActiveControl.Text)
    If Err Then Exit Sub
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set frmNew = Nothing
End Sub

Private Sub txtAlt_GotFocus()
    SelectAllText
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub txtAppClass_GotFocus()
    SelectAllText
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub txtAppTitle_GotFocus()
    SelectAllText
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub txtLaunch_GotFocus()
    SelectAllText
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub txtLong_GotFocus()
    SelectAllText
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub txtNotifyIcon_GotFocus()
    SelectAllText
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub txtPrim_GotFocus()
    SelectAllText
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Sub txtShort_GotFocus()
    Call SelectAllText
    Call DisplayTip(Me.ActiveControl.Tag, Me)
End Sub

Private Function UserName() As String
On Error Resume Next
Dim status      As Integer
Dim lpName      As String
Dim lpUserName  As String
Const lpnLength As Integer = 255
    lpUserName = Space$(lpnLength + 1)
    status = WNetGetUser(lpName, lpUserName, lpnLength)
    If status = 0 Then
        UserName = Left$(lpUserName, InStr(lpUserName, Chr(0)) - 1)
    Else
        UserName = "UNKNOWN"
    End If
End Function

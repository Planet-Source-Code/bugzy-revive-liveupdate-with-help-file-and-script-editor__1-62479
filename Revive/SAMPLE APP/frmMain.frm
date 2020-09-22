VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "ReVive LiveUpdate Sample Project"
   ClientHeight    =   5280
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   Icon            =   "frmMain.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5280
   ScaleWidth      =   4695
   StartUpPosition =   3  'Windows Default
   Begin VB.Frame Frame2 
      Caption         =   "Clean Up"
      Height          =   3315
      Left            =   180
      TabIndex        =   5
      Top             =   1200
      Width           =   4335
      Begin VB.CheckBox chk 
         Caption         =   "ReVive Demonstration Executable               (26K)"
         Height          =   225
         Index           =   2
         Left            =   240
         TabIndex        =   10
         Top             =   2550
         Width           =   3800
      End
      Begin VB.CheckBox chk 
         Caption         =   "ccXPButton User Control                             (46K)"
         Height          =   225
         Index           =   3
         Left            =   240
         TabIndex        =   9
         Top             =   2850
         Width           =   3800
      End
      Begin VB.CheckBox chk 
         Caption         =   "ReVive LiveUpdate Script Editor                (229K)"
         Height          =   225
         Index           =   1
         Left            =   240
         TabIndex        =   8
         Top             =   2250
         Width           =   3800
      End
      Begin VB.TextBox Text1 
         Height          =   1365
         Left            =   240
         Locked          =   -1  'True
         MultiLine       =   -1  'True
         ScrollBars      =   2  'Vertical
         TabIndex        =   7
         Text            =   "frmMain.frx":038A
         Top             =   360
         Width           =   3825
      End
      Begin VB.CheckBox chk 
         Caption         =   "ReVive Version Information File                      (8K)"
         Height          =   225
         Index           =   0
         Left            =   240
         TabIndex        =   6
         Top             =   1950
         Width           =   3800
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Select Demonstration Mode"
      Height          =   885
      Left            =   180
      TabIndex        =   3
      Top             =   180
      Width           =   4335
      Begin VB.ComboBox cmbSelect 
         Height          =   315
         ItemData        =   "frmMain.frx":0558
         Left            =   210
         List            =   "frmMain.frx":056E
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   330
         Width           =   3915
      End
   End
   Begin VB.CommandButton cmdHelp 
      Caption         =   "Help"
      Height          =   375
      Left            =   2460
      TabIndex        =   2
      Top             =   4740
      Width           =   1005
   End
   Begin VB.CommandButton cmdVote 
      Caption         =   "PSC Post"
      Height          =   375
      Left            =   210
      TabIndex        =   1
      Top             =   4740
      Width           =   1005
   End
   Begin VB.CommandButton cmdGo 
      Caption         =   "GO"
      Height          =   375
      Left            =   3540
      TabIndex        =   0
      Top             =   4740
      Width           =   1005
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
'**********************************************************
'//Helper Declares added only to demonstrate functionality.

'//Declares for displaying Help file
Private Const HELP_CONTEXT           As Long = &H1
Private Const HELP_QUIT              As Long = &H2
Private Const HELP_CONTENTS          As Long = &H3&
Private Declare Function WinHelp Lib "user32" Alias "WinHelpA" (ByVal hWnd As Long, ByVal lpFileName As String, ByVal wCommand As Long, ByVal dwData As Any) As Long
'//ShellExecute Declares
Private Const SW_HIDE            As Long = 0 '--- Used for executing regsvr32.exe
Private Const SW_NORMAL          As Long = 1 '--- Used for restarting applications
Private Const SW_MAXIMIZE        As Long = 3 '--- Used when displaying HTML report
Private Declare Function ShellExecute Lib "shell32.dll" Alias "ShellExecuteA" (ByVal hWnd As Long, ByVal lpOperation As String, ByVal lpFile As String, ByVal lpParameters As String, ByVal lpDirectory As String, ByVal nShowCmd As Long) As Long
'//Used to delete files preapring for subsequent demonstrations
Private Declare Function DeleteFile Lib "kernel32" Alias "DeleteFileA" (ByVal lpFileName As String) As Long
'**********************************************************

Private Sub cmdGo_Click()
Dim sReVive As String
Dim x       As Byte

sReVive = App.Path & "\ReVive.exe"

Call CleanPreviousDemo '--------------------- Remove selected files to force new download

Select Case Me.cmbSelect.ListIndex
    Case 0
        Shell sReVive & " " & App.Path & "\RIS Files\nna.ris"
    Case 1
        DeleteFile (App.Path & "\docVersion.rtf") '------------------- Force a file download
        Shell sReVive & " /n " & App.Path & "\RIS Files\nna.ris"
    Case 2
        DeleteFile (App.Path & "\docVersion.rtf") '------------------- Force a file download
        Shell sReVive & " /a " & App.Path & "\RIS Files\nna.ris"
    Case 3
        DeleteFile (App.Path & "\docVersion.rtf") '------------------- Force a file download
        Shell sReVive & " " & App.Path & "\RIS Files\updatemessage.ris"
    Case 4
        MsgBox "Notice during this demonstration ReVive will ask you to shutdown the    " & vbNewLine & _
               "ReViveSampleApp.exe app before continuing. If you choose No, the     " & vbNewLine & _
               "app will still be updated, but the client won't see the update until the    " & vbNewLine & _
               "next time it is executed. If you do close the app, it will be restarted " & vbNewLine & _
               "automatically. If the client is an Admin, remaining files will be cleaned" & vbNewLine & _
               "during the next Windows restart.", vbInformation, "NOTICE"
        Shell sReVive & " " & App.Path & "\RIS Files\requestshutdown.ris"
    Case 5
        MsgBox "Notice during this demonstration that the ReViveSampleApp.exe file you are " & vbNewLine & _
               "currently running will be shutdown and replaced with the updated file.", vbInformation, "NOTICE"
        Shell sReVive & " " & App.Path & "\RIS Files\autoshutdown.ris"
End Select

For x = 0 To 3
    Me.chk(x).Value = 0
Next x

End Sub




'******************************************************
'//Helper Subs added only to demonstrate functionality.
'******************************************************
Private Sub Form_Load()
    If Dir(App.Path & "\ReVive.exe") = "" Then
        MsgBox "The ReVive sample application requires the ReVive LiveUpdate project to be compiled" & vbNewLine & _
               "before it can be used to demonstrate functionality. Please compile the ReVive project " & vbNewLine & _
               "to the 'SAMPLE APP' folder before running this project.", vbExclamation, "NOTICE"
        Unload Me
        Exit Sub
    End If
    Me.cmbSelect.Text = "Demonstrate Normal Mode"
    Me.Move 0, 0
End Sub

Private Sub cmbSelect_Click()
Dim i As Byte
    i = Me.cmbSelect.ListIndex
    If (i = 5 Or i = 6) And InIDE Then
        Me.cmdGo.Enabled = False
        MsgBox "To demonstrate this mode you must be running from a compiled EXE.", vbInformation, "Notice"
    Else
        Me.cmdGo.Enabled = True
    End If
End Sub

Private Sub CleanPreviousDemo()
    If Me.chk(0).Value Then DeleteFile (App.Path & "\docVersion.rtf")
    If Me.chk(1).Value Then DeleteFile (App.Path & "\ScriptEditor.exe")
    If Me.chk(2).Value Then DeleteFile (App.Path & "\ReViveSampleApp.exe")
    If Me.chk(3).Value Then DeleteFile (App.Path & "\ccXPButton.ctl")
End Sub
Private Sub cmdHelp_Click()
    WinHelp Me.hWnd, App.Path & "\ReVive.hlp", HELP_CONTENTS, 0&
End Sub

Private Sub cmdVote_Click()
On Error Resume Next
    ShellExecute Me.hWnd, "open", "http://www.planet-source-code.com/vb/scripts/ShowCode.asp?txtCodeId=62479&lngWId=1", vbNullString, vbNullString, SW_MAXIMIZE
End Sub

Private Function InIDE() As Boolean
On Error GoTo Errs
    Debug.Print 1 / 0
    Exit Function
Errs:
    InIDE = True
End Function

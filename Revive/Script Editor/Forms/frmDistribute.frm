VERSION 5.00
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmDistribute 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Create ReVive Distributable"
   ClientHeight    =   2415
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   7140
   Icon            =   "frmDistribute.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2415
   ScaleWidth      =   7140
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdHelp 
      Cancel          =   -1  'True
      Caption         =   "Help"
      Height          =   345
      Left            =   4170
      TabIndex        =   5
      Top             =   1950
      Width           =   915
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "&Cancel"
      Height          =   345
      Left            =   5160
      TabIndex        =   4
      Top             =   1950
      Width           =   855
   End
   Begin VB.CommandButton cmdCreate 
      Caption         =   "C&reate"
      Enabled         =   0   'False
      Height          =   345
      Left            =   6090
      TabIndex        =   3
      Top             =   1950
      Width           =   855
   End
   Begin VB.TextBox txtPath 
      Height          =   315
      Left            =   960
      Locked          =   -1  'True
      TabIndex        =   2
      Top             =   1410
      Width           =   5955
   End
   Begin VB.CommandButton cmdBrowse 
      Caption         =   "&Browse"
      Height          =   335
      Left            =   150
      TabIndex        =   1
      Top             =   1410
      Width           =   795
   End
   Begin MSComDlg.CommonDialog cd 
      Left            =   0
      Top             =   2040
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   $"frmDistribute.frx":058A
      Height          =   1035
      Left            =   150
      TabIndex        =   0
      Top             =   180
      Width           =   6765
   End
End
Attribute VB_Name = "frmDistribute"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim sFile As String

Private Sub cmdBrowse_Click()
On Error GoTo Errs

'//Opens file open dialog to request filename for distributable
With Me.cd
    .HelpFile = App.path & "\ReVive.hlp"
    .HelpContext = 5
    .HelpCommand = cdlHelpContext
    .CancelError = True
    .Filter = "ReVive Initialization Script (*.ris)|*.ris"
    .DefaultExt = "*.ris"
    .FileName = "update.ris"
    .DialogTitle = "Select a filename for this distributable..."
    .Flags = cdlOFNFileMustExist Or cdlOFNExplorer _
        Or cdlOFNHideReadOnly Or cdlOFNPathMustExist _
        Or cdlOFNShareAware Or cdlOFNNoDereferenceLinks _
        Or cdlOFNOverwritePrompt Or cdlOFNHelpButton
    .ShowSave
    sFile = .FileName
End With

Me.txtPath.Text = sFile
Me.cmdCreate.SetFocus

Errs:
    If Err Then
        Me.cmdCancel.SetFocus
        Exit Sub
    End If
End Sub

Private Sub cmdCancel_Click()
    sFile = ""
    Unload Me
End Sub

Private Sub cmdCreate_Click()
    Unload Me
End Sub

Private Sub cmdHelp_Click()
    '//Open help file and display contents
    WinHelp Me.hwnd, App.path & "\ReVive.hlp", HELP_CONTEXT, CLng(9)
End Sub

Private Sub txtPath_Change()
    Me.cmdCreate.Enabled = Len(Me.txtPath.Text)
End Sub

Public Function SelectDistFile() As String
    Me.Show 1
    SelectDistFile = sFile
End Function

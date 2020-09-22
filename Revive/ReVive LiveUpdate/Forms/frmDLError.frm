VERSION 5.00
Begin VB.Form frmDLError 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EBEDED&
   Caption         =   "ReViveDownloadErrorForm"
   ClientHeight    =   2805
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   6435
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
   Icon            =   "frmDLError.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   187
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   429
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin ReVive.ccXPButton cmd 
      Cancel          =   -1  'True
      Height          =   375
      Index           =   2
      Left            =   2430
      TabIndex        =   2
      ToolTipText     =   "Cancel and exit LiveUpdate"
      Top             =   2220
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
   Begin ReVive.ccXPButton cmd 
      Height          =   375
      Index           =   1
      Left            =   3720
      TabIndex        =   1
      ToolTipText     =   "Continue LiveUpdate without this file"
      Top             =   2220
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "C&ontinue"
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
   Begin ReVive.ccXPButton cmd 
      Default         =   -1  'True
      Height          =   375
      Index           =   0
      Left            =   5010
      TabIndex        =   0
      ToolTipText     =   "Attempt this download again"
      Top             =   2220
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "&Retry"
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
End
Attribute VB_Name = "frmDLError"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'Chris Cochran          cwc.software@gmail.com        Updated: 20 Nov 04
'***********************************************************************
Option Explicit

Public Enum eDownloadResult
    CONNECTERROR = 1
    TRANSFERERROR = 2
End Enum

Private bAction As Byte

Public Function RequestAction(ByVal Identifier As String, ByVal ErrorType As eDownloadResult) As Byte
Dim pIcon   As StdPicture
Dim s       As String
Dim R       As RECT
    Set pIcon = LoadResPicture(201, 1) '---------------------------- Load icon
    Call DrawIconEx(Me.hdc, 14, 36, pIcon, 32, 32, 0, 0, &H3)
    Set pIcon = Nothing
    Call DrawForm(Me) '--------------------------------------------- Clip form and draw border
    Me.ForeColor = 0 '---------------------------------------------- Draw error message
    Me.FontBold = False
    s = "LiveUpdate experienced a " & _
        IIf(ErrorType = CONNECTERROR, "connection error", "transfer error") & _
        " while attempting to download update component '" & _
        FileList.Description(CLng(Identifier)) & " '."
    SetRect R, 56, 38, 410, 73
    Call DrawText(Me.hdc, s, -1, R, DT_FLAGS)
    If FileList.MustUpdate(CLng(Identifier)) Then '----------------- Draw options text
        Me.cmd(2).Move Me.cmd(1).Left '----------------------------- Shift Cancel button over Continue
        Me.cmd(1).TabStop = False
        s = LoadResString(406) '------------------------------------ Warn file is required
    Else
        s = LoadResString(407) '------------------------------------ Warn file IS NOT required
    End If
    SetRect R, 56, 80, 410, 133
    Call DrawText(Me.hdc, s, -1, R, DT_FLAGS + DT_NOPREFIX)
    Screen.MousePointer = vbDefault
    Beep
    Me.Show 1
    RequestAction = bAction
End Function

Private Sub cmd_Click(Index As Integer)
    bAction = Index
    Unload Me
End Sub

Private Sub cmd_FormActivate(Index As Integer, State As WindowState)
'----------------------------------------------------------------------------------
'Draws gradient title as form Activates and Deactivates
'----------------------------------------------------------------------------------
On Error Resume Next
    If Me.WindowState <> vbMinimized And Me.Visible And Index = 0 Then
        Call DrawTitleBar(Me, State, "Component Download Error")
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Me.cmd(0).SetFocus
    cmd_FormActivate 0, Active
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'----------------------------------------------------------------------------------
'Set the move cursor and move the form when mouse is down in title bar
'----------------------------------------------------------------------------------
On Error Resume Next
Dim R As RECT
    If Button = vbLeftButton Then
        Call SetRect(R, 0, 0, Me.ScaleWidth, 24)
        If PtInRect(R, CLng(x), CLng(y)) Then
            SetCursor LoadCursor(0, IDC_SIZEALL)
            Call ReleaseCapture
            Call SendMessage(hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SetForegroundWindow(frmMain.hWnd) '-- Return focus to frmMain
    Set frmDLError = Nothing
End Sub

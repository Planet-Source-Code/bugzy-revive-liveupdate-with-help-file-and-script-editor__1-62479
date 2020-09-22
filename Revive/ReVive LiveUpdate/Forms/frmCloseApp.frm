VERSION 5.00
Begin VB.Form frmCloseApp 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EBEDED&
   Caption         =   "ReViveAppWarningForm"
   ClientHeight    =   2400
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5820
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmCloseApp.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   160
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   388
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin ReVive.ccXPButton cmdContinue 
      Height          =   375
      Left            =   4350
      TabIndex        =   0
      Top             =   1800
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
End
Attribute VB_Name = "frmCloseApp"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'Chris Cochran          cwc.software@gmail.com        Updated: 18 Nov 04
'***********************************************************************
Option Explicit

Private bContinue   As Boolean

Public Function WarnAppIsRunning() As Boolean
Dim pIcon   As StdPicture
Dim R       As RECT
Dim s       As String
    Set pIcon = LoadResPicture(201, 1) '---------------------------- Load icon
    Call DrawIconEx(Me.hdc, 14, 36, pIcon, 32, 32, 0, 0, &H3)
    Set pIcon = Nothing
    Call DrawForm(Me) '--------------------------------------------- Clip form and draw border
    Me.ForeColor = 0: Me.FontBold = False '------------------------- Draw app running warning text
    SetRect R, 58, 36, 370, 105
    s = "LiveUpdate has found an instance of " & Setup.AppShortName & LoadResString(409) & Setup.AppShortName & LoadResString(410)
    DrawText Me.hdc, s, -1, R, DT_FLAGS + DT_LEFT
    Beep
    Screen.MousePointer = vbDefault
    Me.Show 1
    If Setup.RunMode = eNORMAL Then Screen.MousePointer = vbHourglass
    WarnAppIsRunning = bContinue
End Function

Private Sub cmdContinue_Click()
    bContinue = True
    Unload Me
End Sub

Private Sub cmdContinue_FormActivate(State As WindowState)
'----------------------------------------------------------------------------------
'Draws gradient title bar as form Activates and Deactivates
'----------------------------------------------------------------------------------
On Error Resume Next
    If Me.WindowState <> vbMinimized And Me.Visible Then
        Call DrawTitleBar(Me, State, "Warning - Update Application Running")
    End If
End Sub

Private Sub Form_Activate()
    Me.cmdContinue.SetFocus
    cmdContinue_FormActivate Active
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
            Call SendMessage(Me.hWnd, WM_NCLBUTTONDOWN, HTCAPTION, 0&)
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call SetForegroundWindow(frmMain.hWnd) '-- Return focus to frmMain
    Set frmCloseApp = Nothing
End Sub

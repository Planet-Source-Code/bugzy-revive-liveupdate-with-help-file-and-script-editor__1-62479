VERSION 5.00
Begin VB.Form frmExit 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EBEDED&
   Caption         =   "ReviveAppExitForm"
   ClientHeight    =   2025
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5505
   ForeColor       =   &H00FFFFFF&
   Icon            =   "frmExit.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   135
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   367
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Visible         =   0   'False
   Begin ReVive.ccXPButton cmdNo 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4050
      TabIndex        =   0
      Top             =   1410
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "&No"
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
   Begin ReVive.ccXPButton cmdYes 
      Height          =   375
      Left            =   2760
      TabIndex        =   1
      Top             =   1410
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
      Caption         =   "&Yes"
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
Attribute VB_Name = "frmExit"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'Chris Cochran          cwc.software@gmail.com        Updated: 25 Mar 05
'***********************************************************************
Option Explicit

Public Enum enumAnswer
    No = 0
    Yes = 1
End Enum

Private Result As enumAnswer

Public Function ShowYesNo() As enumAnswer
Dim pIcon   As StdPicture
Dim s       As String
Dim R       As RECT
    Call DrawForm(Me) '--------------------- Clip form and draw border
    Set pIcon = LoadResPicture(202, 1) '---- Draw icon
    Call DrawIconEx(Me.hdc, 16, 38, pIcon, 32, 32, 0, 0, &H3)
    Set pIcon = Nothing
    Me.ForeColor = 0
    Me.FontBold = False
    s = LoadResString(408) '---------------- Text to ask if user wants to exit
    SetRect R, 58, 46, 410, 73
    Call DrawText(Me.hdc, s, -1, R, DT_FLAGS)
    Beep
    Me.Show 1
    ShowYesNo = Result
End Function

Private Sub cmdNo_Click()
    Result = No
    Unload Me
End Sub

Private Sub cmdYes_Click()
    Result = Yes
    Unload Me
End Sub

Private Sub cmdYes_FormActivate(State As WindowState)
'----------------------------------------------------------------------------------
'Draws gradient title as form Activates and Deactivates
'----------------------------------------------------------------------------------
On Error Resume Next
    If Me.WindowState <> vbMinimized And Me.Visible Then
        Call DrawTitleBar(Me, State, Setup.AppShortName & " LiveUpdate")
    End If
End Sub

Private Sub Form_Activate()
    Me.cmdNo.SetFocus
    cmdYes_FormActivate Active
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
    Set frmExit = Nothing
End Sub

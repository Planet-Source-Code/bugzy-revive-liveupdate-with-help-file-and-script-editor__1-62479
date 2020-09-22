VERSION 5.00
Begin VB.Form frmUpdateMessage 
   AutoRedraw      =   -1  'True
   BackColor       =   &H00EBEDED&
   ClientHeight    =   3870
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   5850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   258
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   390
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txtInfo 
      Appearance      =   0  'Flat
      BorderStyle     =   0  'None
      Height          =   1815
      Left            =   300
      Locked          =   -1  'True
      MultiLine       =   -1  'True
      ScrollBars      =   2  'Vertical
      TabIndex        =   0
      TabStop         =   0   'False
      Top             =   1290
      Width           =   5205
   End
   Begin ReVive.ccXPButton cmdOK 
      Cancel          =   -1  'True
      Height          =   375
      Left            =   4350
      TabIndex        =   1
      Top             =   3300
      Width           =   1185
      _ExtentX        =   2090
      _ExtentY        =   661
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
   Begin VB.Shape Shape1 
      BackColor       =   &H00C00000&
      BorderColor     =   &H008F837A&
      FillColor       =   &H00C00000&
      Height          =   1875
      Left            =   270
      Top             =   1260
      Width           =   5265
   End
End
Attribute VB_Name = "frmUpdateMessage"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'Chris Cochran          cwc.software@gmail.com        Updated: 06 Sep 05
'***********************************************************************
Option Explicit

Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub cmdOK_FormActivate(State As WindowState)
'----------------------------------------------------------------------------------
'Draws gradient title as form Activates and Deactivates
'----------------------------------------------------------------------------------
On Error Resume Next
    If Me.WindowState <> vbMinimized Then
        Call DrawTitleBar(Me, State, Setup.AppLongName)
    End If
End Sub

Private Sub Form_Activate()
    On Error Resume Next
    Me.cmdOK.SetFocus
    cmdOK_FormActivate Active
End Sub

Private Sub Form_Load()
Dim pIcon   As StdPicture
Dim s       As String
Dim R       As RECT
    Call DrawForm(Me) '--------------------- Clip form and draw border
    Set pIcon = LoadResPicture(207, 1) '---- Draw icon
    Call DrawIconEx(Me.hdc, 16, 38, pIcon, 32, 32, 0, 0, &H3)
    Set pIcon = Nothing
    Me.ForeColor = 0
    Me.FontBold = False
    s = LoadResString(411) '---------------- Load text from resource file
    SetRect R, 58, 46, 410, 73
    Call DrawText(Me.hdc, s, -1, R, DT_FLAGS)
    Me.txtInfo.Text = sUpdateMessage
    Beep
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
    Set frmUpdateMessage = Nothing
End Sub

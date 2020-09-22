VERSION 5.00
Begin VB.Form frmNotify 
   AutoRedraw      =   -1  'True
   BorderStyle     =   0  'None
   Caption         =   "ReViveNotifyForm"
   ClientHeight    =   1230
   ClientLeft      =   12765
   ClientTop       =   9930
   ClientWidth     =   2940
   ControlBox      =   0   'False
   ForeColor       =   &H00000000&
   Icon            =   "frmNotify.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   82
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   196
   ShowInTaskbar   =   0   'False
End
Attribute VB_Name = "frmNotify"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'***********************************************************************
'Chris Cochran          cwc.software@gmail.com        Updated: 06 Jan 05
'***********************************************************************
Option Explicit

'//Play Wave from Resource File
Private Declare Function PlaySound Lib "winmm.dll" Alias "PlaySoundA" (ByVal lpszName As Long, ByVal hModule As Long, ByVal dwFlags As Long) As Long

'//Get WorkArea Declares
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As RECT, ByVal fuWinIni As Long) As Long
Private Const SPI_GETWORKAREA   As Long = 48

Private WithEvents tmrShowForm  As Timer '----- Timer used to show form
Attribute tmrShowForm.VB_VarHelpID = -1
Private WithEvents tmrHideForm  As Timer '----- Timer used to hide form
Attribute tmrHideForm.VB_VarHelpID = -1

Private lHwnd                   As Long
Private FORMLEFT                As Long  '----- Stores screen coordinates for left side of notify form
Private TASKBARTOP              As Long  '----- Stores screen coordinates for top of taskbar
Private FORMPIXWIDTH            As Long
Private FORMPIXHEIGHT           As Long
Private MOUSEX                  As Long  '----- Stores the forms X cursor coordinate
Private MOUSEY                  As Long  '----- Stores the forms Y cursor coordinate
Private bReview                 As Boolean '--- True when seleting "Review..." text

Public Sub Notify(Optional ByVal bComplete As Boolean = False)
On Error GoTo Errs
Dim R       As RECT
Dim sIcon   As String
Dim pIcon   As StdPicture
Dim sTitle  As String
Dim sReview As String
Dim lHdc    As Long
    lHdc = Me.hdc '------------------------------------------------- Assign forms private variables
    lHwnd = Me.hWnd
    FORMPIXWIDTH = Me.ScaleWidth
    FORMPIXHEIGHT = Me.ScaleHeight
    FORMLEFT = (GetWorkArea("WIDTH") / Screen.TwipsPerPixelX) - FORMPIXWIDTH
    TASKBARTOP = GetWorkArea("HEIGHT") / Screen.TwipsPerPixelY
    Set tmrShowForm = Me.Controls.Add("VB.timer", "tmrShowForm") '-- Load timers used to show and hide Notify form
    Set tmrHideForm = Me.Controls.Add("VB.timer", "tmrHideForm")
    tmrShowForm.Enabled = False: tmrShowForm.Interval = 10
    tmrHideForm.Enabled = False: tmrHideForm.Interval = 10
    If bComplete Then '----------------------- Prepare to draw form dependent on update status
        sTitle = Setup.AppShortName & " Update Complete"
        sReview = "Review Installed Updates"
    Else
        sTitle = Setup.AppShortName & " Updates Available"
        sReview = "Review and Install Updates"
    End If
    Me.PaintPicture LoadResPicture(301, 0), 4, 4
    '//Display designated app icon if it was download, otherwise default to ReVive icon
    sIcon = sTEMPDIR & "\Notify.ico"
    Set pIcon = IIf(Len(Dir$(sIcon, 39)), LoadPicture(sIcon), LoadResPicture(200, 1))
    Call DrawIconEx(lHdc, 8, 8, pIcon.Handle, 32, 32, 0, 0, &H3)
    Set pIcon = LoadResPicture(203, 1) '------- Load "Idle" Close icon from resource file
    Call DrawIconEx(lHdc, 171, 9, pIcon.Handle, 14, 13, 0, 0, &H3)
    SetRect R, 44, 10, 163, 49 '--------------- Draw title text
    Call DrawText(lHdc, sTitle, -1, R, DT_FLAGS + DT_CENTER + DT_NOPREFIX)
    Me.FontUnderline = True '------------------ Draw review text
    SetRect R, 30, 54, 164, 67
    Call DrawText(lHdc, sReview, -1, R, DT_FLAGS + DT_CENTER + DT_NOPREFIX)
    tmrShowForm.Enabled = True '--------------- Activate form show timer
    Set pIcon = Nothing '---------------------- Cleanup
Errs_Exit:
    Exit Sub
Errs:
    Unload frmMain
End Sub

Private Sub Form_Click()
'//Responds to users mouseclick on Close pic or Review text
Dim R As RECT
    SetRect R, 171, 9, 185, 22
    If PtInRect(R, MOUSEX, MOUSEY) Then '------ Close button
        SetForegroundWindow lPREVWINDOW '------ Return focus to previously active window
        Call CloseForm
    Else
        SetRect R, 28, 54, 162, 67
        If PtInRect(R, MOUSEX, MOUSEY) Then '-- Review text
            Setup.RunMode = eNORMAL
            bReview = True
            frmMain.Show
            tmrHideForm.Enabled = True '------- Begin hiding form
        End If
    End If
End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
'//Simulate mouseover event for our form hotspots to toggle hand
'..cursor and close button graphic (hot or idle)
Dim r1              As RECT
Dim r2              As RECT
Dim pIcon           As StdPicture
Static bInButton    As Boolean
    MOUSEX = CLng(x)
    MOUSEY = CLng(y)
    SetRect r1, 171, 9, 185, 22
    If PtInRect(r1, MOUSEX, MOUSEY) Then '--------- Close button
        SetCursor LoadCursor(0, IDC_HAND)
        If Not bInButton Then
            lPREVWINDOW = GetForegroundWindow '---- Get previously active window
            Set pIcon = LoadResPicture(204, 1) '--- Switch to the "Hot" button
            Call DrawIconEx(Me.hdc, 171, 9, pIcon.Handle, 14, 13, 0, 0, &H3)
            '//Repaint only the close button, NOT the entire window
            Call RedrawWindow(Me.hWnd, r1, 0&, RDW_FLAGS)
            Set pIcon = Nothing
            bInButton = True
        End If
    Else
        SetRect r2, 30, 54, 164, 67
        If PtInRect(r2, MOUSEX, MOUSEY) Then '----- Review text
            SetCursor LoadCursor(0, IDC_HAND)
        End If
        If bInButton Then
            Set pIcon = LoadResPicture(203, 1) '--- Switch to the "Idle" button
            Call DrawIconEx(Me.hdc, 171, 9, pIcon.Handle, 14, 13, 0, 0, &H3)
            '//Repaint only the close button, NOT the entire window
            Call RedrawWindow(Me.hWnd, r1, 0&, RDW_FLAGS)
            Set pIcon = Nothing
            bInButton = False
        End If
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set tmrShowForm = Nothing
    Set tmrHideForm = Nothing
    'Skip "Set frmNotify = Nothing" because pointer to tmrHideForm is still in use
    'since Unload event was called from tmrHideForm. If added, memory read error WILL occur.
End Sub

Private Sub CloseForm()
Dim m   As Long
    If bREBOOT Then
        If Setup.ForceReboots Then
            MsgBox "Updates to your " & Setup.AppShortName & " application require your computer to be     " & Chr(13) & _
                   "restarted. Save your work and click OK to restart your computer now.   ", vbExclamation, "Restart Required"
            Call Reboot
        Else
            m = MsgBox("Updates to your " & Setup.AppShortName & " application require your computer to be     " & Chr(13) & _
                   "restarted. Would you like to restart your computer now?   ", vbYesNoCancel + vbQuestion, "Restart Required")
            If m = 6 Then
                Call Reboot
            ElseIf m = 2 Then
                Exit Sub
            End If
        End If
    End If
    tmrHideForm.Enabled = True
End Sub

Private Sub tmrShowForm_Timer()
'------------------------------------------------------------
' Purpose   : Show notification form as a popup above taskbar
'------------------------------------------------------------
Static x            As Long
Dim lResult         As Long
Dim lHdc            As Long
Const SND_NODEFAULT As Long = &H2
Const SND_ASYNC     As Long = &H1
Const SND_RESOURCE  As Long = &H40004
Const SND_FLAGS     As Long = SND_RESOURCE + SND_ASYNC + SND_NODEFAULT
    If x = 0 Then
        x = 4
        'NOTE: SOUND ONLY PLAYS FROM COMPILED EXE
        lResult = PlaySound(100, App.hInstance, SND_FLAGS)
        lHdc = Me.hdc
        Call DrawBorder(lHdc, 0, FORMPIXWIDTH, 0, FORMPIXHEIGHT, RaisedHigh)
        Call DrawBorder(lHdc, 3, FORMPIXWIDTH - 3, 3, FORMPIXHEIGHT - 3, SunkenShallow)
    End If
    Call SetWindowPos(lHwnd, HWND_TOPMOST, FORMLEFT, TASKBARTOP - x, FORMPIXWIDTH, x, SWP_NOACTIVATE + SWP_SHOWWINDOW)
    If x < FORMPIXHEIGHT Then
        x = x + 2
    Else
        tmrShowForm.Enabled = False
    End If
End Sub

Private Sub tmrHideForm_Timer()
'------------------------------------------------------------
' Purpose   : Hide notification form. Timer cycles until form
'             is hidden, then form unloads.
'------------------------------------------------------------
On Error Resume Next
Static x As Long
    If x = 0 Then x = FORMPIXHEIGHT
    Call SetWindowPos(lHwnd, HWND_TOPMOST, FORMLEFT, TASKBARTOP - x, FORMPIXWIDTH, x, SWP_NOACTIVATE + SWP_SHOWWINDOW)
    If x > 3 Then x = x - 3: Exit Sub
    tmrHideForm.Enabled = False
    If Not bReview Then Unload frmMain
    Unload Me
End Sub

Private Function GetWorkArea(ByVal HEIGHTorWIDTH As String) As Long
'-----------------------------------------------------------------------------
' Purpose   : Get client work area (Height or Width) minus the Windows taskbar
'-----------------------------------------------------------------------------
On Error GoTo Err
Dim wa_info As RECT
    If UCase$(HEIGHTorWIDTH) = "HEIGHT" Then
        If SystemParametersInfo(SPI_GETWORKAREA, 0, wa_info, 0) <> 0 Then
            GetWorkArea = wa_info.lBottom * Screen.TwipsPerPixelY
        Else
            GetWorkArea = Screen.Height
        End If
    Else
        If SystemParametersInfo(SPI_GETWORKAREA, 0, wa_info, 0) <> 0 Then
            GetWorkArea = wa_info.lRight * Screen.TwipsPerPixelX
        Else
            GetWorkArea = Screen.Width
        End If
    End If
Err_Exit:
    Exit Function
Err:
    Resume Err_Exit
End Function

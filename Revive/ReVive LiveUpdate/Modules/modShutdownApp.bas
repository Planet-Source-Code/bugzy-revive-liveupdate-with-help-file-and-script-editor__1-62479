Attribute VB_Name = "modShutdownApp"
Option Explicit

Private Declare Function EnumWindows Lib "user32" (ByVal lpEnumFunc As Long, ByVal lParam As Long) As Long
Private Declare Function GetWindowText Lib "user32" Alias "GetWindowTextA" (ByVal hwnd As Long, ByVal lpString As String, ByVal cch As Long) As Long
Private Declare Function GetClassName Lib "user32" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long
Private Const MAX_PATH  As Long = 260

Private pClassName      As String
Private pSearchTitle    As String
Private pAppFound       As Boolean

Private Function EnumWindowsCallBack(ByVal lHandle As Long, ByVal lpData As Long) As Long

Dim lResult     As Long
Dim sWindowTitle   As String
Dim sClassName  As String

EnumWindowsCallBack = 1

If Not pAppFound Then
    sWindowTitle = Space$(MAX_PATH)
    lResult = GetWindowText(lHandle, sWindowTitle, MAX_PATH)
    sWindowTitle = UCase$(Left$(sWindowTitle, lResult))
    sClassName = Space$(MAX_PATH)
    lResult = GetClassName(lHandle, sClassName, MAX_PATH)
    sClassName = UCase$(Left$(sClassName, lResult))
    If InStr(1, sWindowTitle, pSearchTitle) Then
        If Len(pClassName) Then
            If sClassName = pClassName Then
                pAppFound = True
            End If
        Else
            pAppFound = True
        End If
    End If
End If

End Function

Public Function IsAppRunning(ByVal sWindowCaptionElement As String, Optional ByVal sClassName As String = "") As Boolean
Dim lHandle As Long
    pSearchTitle = UCase$(sWindowCaptionElement)
    pClassName = UCase$(sClassName)
    Call EnumWindows(AddressOf EnumWindowsCallBack, lHandle) '---- Enumerate windows to see if ours is running
    IsAppRunning = pAppFound '------------------------------------ Set function return
    pClassName = vbNull '----------------------------------------- Cleanup
    pSearchTitle = vbNull
End Function

Public Sub Main()

MsgBox IsAppRunning("ReVive LiveUpdate")

End Sub

Attribute VB_Name = "modGlobal"
'****************************************************************************
'
'Chris Cochran            cwc.software@gmail.com           Updated: 22 Aug 05
'
'I added this project for the sole purpose of making it easier to register
'Windows to open .ris files with the ReVive update program. You will want
'to incorporate some variation of this code to do this for clients computers
'if you will be using one ReVive installation to update numerous applications.
'
'This can also be accomplished on client computers by checking RegRISFiles in
'the new script form when creating a script. ReVive will accomplish the .ris
'file registration for you each time it is executed on the client providing
'the logged in user is an admin.
'****************************************************************************

Option Explicit

Public Sub Main()
On Error GoTo Errs
Dim c       As New cReg
Dim sPath   As String
    sPath = InputBox("Please enter the full path, including filename," & vbNewLine & _
                     "to your ReVive executable file, such as:  " & vbNewLine & vbNewLine & _
                     "D:\Projects\My Projects\Update\ReVive\ReVive.exe", "ReVive")
    If Len(Trim(sPath)) = 0 Then Exit Sub
    With c
        .ClassKey = HKEY_CLASSES_ROOT
        .SectionKey = ".ris"
        .ValueType = REG_SZ
        .ValueKey = ""
        .Value = "ReVive.Initialization.Script"
        .CreateKey
        .SectionKey = "ReVive.Initialization.Script"
        .ValueKey = ""
        .Value = "ReVive LiveUpdate Initialization Script"
        .CreateKey
        .SectionKey = "ReVive.Initialization.Script\shell\open\command"
        .ValueKey = ""
        .Value = Chr(34) & sPath & Chr(34) & " " & Chr(34) & "%1" & Chr(34)
        .CreateKey
    End With
Errs_Exit:
    Set c = Nothing
    Exit Sub
Errs:
    Resume Errs_Exit

End Sub

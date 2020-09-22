VERSION 5.00
Begin VB.UserControl ccFileList 
   ClientHeight    =   1230
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   2490
   ScaleHeight     =   1230
   ScaleWidth      =   2490
   Begin VB.VScrollBar vs 
      Height          =   1125
      Left            =   2220
      Max             =   0
      TabIndex        =   1
      Top             =   60
      Width           =   255
   End
   Begin VB.PictureBox Picture1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      BorderStyle     =   0  'None
      ForeColor       =   &H80000008&
      Height          =   1095
      Left            =   60
      ScaleHeight     =   1095
      ScaleWidth      =   2115
      TabIndex        =   0
      Top             =   60
      Width           =   2115
   End
End
Attribute VB_Name = "ccFileList"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type typeColDimensions
    tLeft           As Long
    tWidth          As Long
End Type
Private Type typeColumns
    tDescription    As typeColDimensions
    tStatus         As typeColDimensions
End Type

Private pCOUNT      As Long

Public Sub Add(ByVal Description As String, ByVal StatusText As String)

pCOUNT = pCOUNT + 1

End Sub

Public Sub StatusText(ByVal Index As Long, ByVal Status As String)

End Sub

Public Property Get Count() As Long
    Count = pCOUNT
End Property

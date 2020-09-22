VERSION 5.00
Begin VB.UserControl XPProgressBar 
   AutoRedraw      =   -1  'True
   ClientHeight    =   540
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   630
   ScaleHeight     =   36
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   42
End
Attribute VB_Name = "XPProgressBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private MaxVal As Double
Private MinVal As Double
Private PValue As Double

Private StepLength As Integer
Private SeperatorWidth As Integer

Private Back_Color As OLE_COLOR
Private Bar_Color As OLE_COLOR

Private FullItems As Boolean

Private Declare Function TranslateColor Lib "olepro32.dll" Alias "OleTranslateColor" (ByVal clr As OLE_COLOR, ByVal palet As Long, Col As Long) As Long

Public Property Let Max(temp As Double)
MaxVal = temp
UserControl_Resize
End Property

Public Property Get Max() As Double
Max = MaxVal
End Property

Public Property Let Min(temp As Double)
MinVal = temp
End Property

Public Property Get Min() As Double
Min = MinVal
End Property

Public Property Let Value(temp As Double)
PValue = temp
UserControl_Resize
End Property

Public Property Get Value() As Double
Value = PValue
End Property

Public Property Let Step_Length(temp As Integer)
StepLength = temp
UserControl_Resize
End Property

Public Property Get Step_Length() As Integer
Step_Length = StepLength
End Property

Public Property Let Seperator_Width(temp As Integer)
SeperatorWidth = temp
UserControl_Resize
End Property

Public Property Get Seperator_Width() As Integer
Seperator_Width = SeperatorWidth
End Property

Public Property Let DrawOnlyFullItems(temp As Boolean)
FullItems = temp
UserControl_Resize
End Property

Public Property Get DrawOnlyFullItems() As Boolean
DrawOnlyFullItems = FullItems
End Property

Public Property Let BackColor(newValue As OLE_COLOR)
Back_Color = newValue
UserControl.BackColor = Back_Color
UserControl_Resize
End Property

Public Property Get BackColor() As OLE_COLOR
BackColor = Back_Color
End Property

Public Property Let BarColor(newValue As OLE_COLOR)
Bar_Color = newValue
UserControl_Resize
End Property

Public Property Get BarColor() As OLE_COLOR
BarColor = Bar_Color
End Property

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
DrawOnlyFullItems = PropBag.ReadProperty("DrawOnlyFullItems", False)
Max = PropBag.ReadProperty("Max", 100)
Min = PropBag.ReadProperty("Min", 0)
Value = PropBag.ReadProperty("Value", 0)
Step_Length = PropBag.ReadProperty("Step_Length", 5)
Seperator_Width = PropBag.ReadProperty("Seperator_Width", 2)
Back_Color = PropBag.ReadProperty("BackColor", vbWhite)
Bar_Color = PropBag.ReadProperty("BarColor", 3724597)

UserControl.BackColor = Back_Color
UserControl_Resize
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
    .WriteProperty "DrawOnlyFullItems", FullItems, False
    .WriteProperty "Max", MaxVal, 100
    .WriteProperty "Min", MinVal, 0
    .WriteProperty "Value", PValue, 0
    .WriteProperty "Step_Length", StepLength, 5
    .WriteProperty "Seperator_Width", SeperatorWidth, 2
    .WriteProperty "BackColor", Back_Color
    .WriteProperty "BarColor", Bar_Color
End With
End Sub

Private Sub DrawTop(PosX As Long, PosY As Long, Width As Integer)
UserControl.Line (PosX, PosY)-(PosX + Width, PosY), RGB(104, 104, 104), BF
UserControl.Line (PosX, PosY + 1)-(PosX + Width, PosY + 1), RGB(190, 190, 190), BF
UserControl.Line (PosX, PosY + 2)-(PosX + Width, PosY + 2), RGB(239, 239, 239), BF
End Sub

Private Sub DrawBottom(PosX As Long, PosY As Long, Width As Integer)
UserControl.Line (PosX, PosY)-(PosX + Width, PosY), RGB(255, 255, 255), BF
UserControl.Line (PosX, PosY + 1)-(PosX + Width, PosY + 1), RGB(239, 239, 239), BF
UserControl.Line (PosX, PosY + 2)-(PosX + Width, PosY + 2), RGB(104, 104, 104), BF
End Sub

Private Sub DrawSingleStep(PosX As Long, PosY As Long, PosYEnd As Long, Width As Integer)
UserControl.Line (PosX, PosY)-(PosX + Width, PosY), BlendColors(Bar_Color, vbWhite, 50), BF
UserControl.Line (PosX, PosY + 1)-(PosX + Width, PosY + 1), BlendColors(Bar_Color, vbWhite, 40), BF
UserControl.Line (PosX, PosY + 2)-(PosX + Width, PosY + 2), BlendColors(Bar_Color, vbWhite, 30), BF
UserControl.Line (PosX, PosY + 3)-(PosX + Width, PosY + 3), BlendColors(Bar_Color, vbWhite, 20), BF
UserControl.Line (PosX, PosY + 4)-(PosX + Width, PosY + 4), BlendColors(Bar_Color, vbWhite, 10), BF

UserControl.Line (PosX, PosY + 5)-(PosX + Width, PosYEnd - 4), Bar_Color, BF

UserControl.Line (PosX, PosYEnd - 3)-(PosX + Width, PosYEnd - 3), BlendColors(Bar_Color, &H808080, 5), BF
UserControl.Line (PosX, PosYEnd - 2)-(PosX + Width, PosYEnd - 2), BlendColors(Bar_Color, &H808080, 10), BF
UserControl.Line (PosX, PosYEnd - 1)-(PosX + Width, PosYEnd - 1), BlendColors(Bar_Color, &H808080, 15), BF
UserControl.Line (PosX, PosYEnd)-(PosX + Width, PosYEnd), BlendColors(Bar_Color, &H808080, 20), BF
End Sub

Private Sub DrawLeftSide(PosX As Long, PosY As Long, PosYEnd As Long)
UserControl.Line (PosX, PosY)-(PosX, PosY), RGB(255, 255, 255), BF
UserControl.Line (PosX, PosY + 1)-(PosX, PosY + 1), RGB(172, 171, 166), BF
UserControl.Line (PosX, PosY + 2)-(PosX, PosY + 2), RGB(127, 126, 125), BF
UserControl.Line (PosX, PosY + 3)-(PosX, PosYEnd - 3), RGB(104, 104, 104), BF
UserControl.Line (PosX, PosYEnd - 2)-(PosX, PosYEnd - 2), RGB(127, 126, 125), BF
UserControl.Line (PosX, PosYEnd - 1)-(PosX, PosYEnd - 1), RGB(172, 171, 166), BF
UserControl.Line (PosX, PosYEnd)-(PosX, PosYEnd), RGB(255, 255, 255), BF

UserControl.Line (PosX + 1, PosY)-(PosX + 1, PosY), RGB(172, 171, 167), BF
UserControl.Line (PosX + 1, PosY + 1)-(PosX + 1, PosY + 1), RGB(119, 119, 119), BF
UserControl.Line (PosX + 1, PosY + 2)-(PosX + 1, PosYEnd - 2), RGB(190, 190, 190), BF
UserControl.Line (PosX + 1, PosYEnd - 1)-(PosX + 1, PosYEnd - 1), RGB(119, 119, 119), BF
UserControl.Line (PosX + 1, PosYEnd)-(PosX + 1, PosYEnd), RGB(172, 171, 167), BF

UserControl.Line (PosX + 2, PosY)-(PosX + 2, PosY), RGB(127, 126, 125), BF
UserControl.Line (PosX + 2, PosY + 1)-(PosX + 2, PosY + 2), RGB(190, 190, 190), BF
UserControl.Line (PosX + 2, PosY + 3)-(PosX + 2, PosYEnd - 1), RGB(239, 239, 239), BF
UserControl.Line (PosX + 2, PosYEnd)-(PosX + 2, PosYEnd), RGB(127, 126, 125), BF

UserControl.Line (PosX + 3, PosY)-(PosX + 3, PosY), RGB(104, 104, 104), BF
UserControl.Line (PosX + 3, PosY + 1)-(PosX + 3, PosY + 1), RGB(190, 190, 190), BF
UserControl.Line (PosX + 3, PosY + 2)-(PosX + 3, PosY + 3), RGB(239, 239, 239), BF
UserControl.Line (PosX + 3, PosY + 4)-(PosX + 3, PosYEnd - 3), RGB(255, 255, 255), BF
UserControl.Line (PosX + 3, PosYEnd - 2)-(PosX + 3, PosYEnd - 1), RGB(239, 239, 239), BF
UserControl.Line (PosX + 3, PosYEnd)-(PosX + 3, PosYEnd), RGB(104, 104, 104), BF
End Sub

Private Sub DrawRightSide(PosX As Long, PosY As Long, PosYEnd As Long)
UserControl.Line (PosX + 3, PosY)-(PosX + 3, PosY), RGB(255, 255, 255), BF
UserControl.Line (PosX + 3, PosY + 1)-(PosX + 3, PosY + 1), RGB(172, 171, 166), BF
UserControl.Line (PosX + 3, PosY + 2)-(PosX + 3, PosY + 2), RGB(127, 126, 125), BF
UserControl.Line (PosX + 3, PosY + 3)-(PosX + 3, PosYEnd - 3), RGB(104, 104, 104), BF
UserControl.Line (PosX + 3, PosYEnd - 2)-(PosX + 3, PosYEnd - 2), RGB(127, 126, 125), BF
UserControl.Line (PosX + 3, PosYEnd - 1)-(PosX + 3, PosYEnd - 1), RGB(172, 171, 166), BF
UserControl.Line (PosX + 3, PosYEnd)-(PosX + 3, PosYEnd), RGB(255, 255, 255), BF

UserControl.Line (PosX + 2, PosY)-(PosX + 2, PosY), RGB(172, 171, 167), BF
UserControl.Line (PosX + 2, PosY + 1)-(PosX + 2, PosY + 1), RGB(119, 119, 119), BF
UserControl.Line (PosX + 2, PosY + 2)-(PosX + 2, PosYEnd - 2), RGB(190, 190, 190), BF
UserControl.Line (PosX + 2, PosYEnd - 1)-(PosX + 2, PosYEnd - 1), RGB(119, 119, 119), BF
UserControl.Line (PosX + 2, PosYEnd)-(PosX + 2, PosYEnd), RGB(172, 171, 167), BF

UserControl.Line (PosX + 1, PosY)-(PosX + 1, PosY), RGB(127, 126, 125), BF
UserControl.Line (PosX + 1, PosY + 1)-(PosX + 1, PosY + 2), RGB(190, 190, 190), BF
UserControl.Line (PosX + 1, PosY + 3)-(PosX + 1, PosYEnd - 1), RGB(239, 239, 239), BF
UserControl.Line (PosX + 1, PosYEnd)-(PosX + 1, PosYEnd), RGB(127, 126, 125), BF

UserControl.Line (PosX, PosY)-(PosX, PosY), RGB(104, 104, 104), BF
UserControl.Line (PosX, PosY + 1)-(PosX, PosY + 1), RGB(190, 190, 190), BF
UserControl.Line (PosX, PosY + 2)-(PosX, PosY + 3), RGB(239, 239, 239), BF
UserControl.Line (PosX, PosY + 4)-(PosX, PosYEnd - 3), RGB(255, 255, 255), BF
UserControl.Line (PosX, PosYEnd - 2)-(PosX, PosYEnd - 1), RGB(239, 239, 239), BF
UserControl.Line (PosX, PosYEnd)-(PosX, PosYEnd), RGB(104, 104, 104), BF
End Sub

Private Sub UserControl_Resize()
Dim i As Integer, FullItemsCount As Integer

If UserControl.ScaleWidth < 19 Then UserControl.Width = 19 * Screen.TwipsPerPixelY     ' --> Size usercontrol
If UserControl.ScaleHeight < 17 Then UserControl.Height = 17 * Screen.TwipsPerPixelX

UserControl.Cls
DrawLeftSide 0, 0, UserControl.ScaleHeight - 1
DrawRightSide UserControl.ScaleWidth - 4, 0, UserControl.ScaleHeight - 1
DrawTop 4, 0, UserControl.ScaleWidth - 8
DrawBottom 4, UserControl.ScaleHeight - 3, UserControl.ScaleWidth - 8
    
If PValue > MaxVal Then PValue = MaxVal
If MaxVal <= MinVal Or PValue < MinVal Or PValue > MaxVal Or StepLength < 1 Or SeperatorWidth < 0 Then Exit Sub
    
For i = 1 To Int((UserControl.ScaleWidth - 8) * (PValue - MinVal) / (MaxVal - MinVal) / (StepLength + SeperatorWidth)) '--> Draw full items
    DrawSingleStep ((i - 1) * StepLength + (i - 1) * SeperatorWidth) + 4, 3, UserControl.ScaleHeight - 4, StepLength
Next i

If Not FullItems Then DrawSingleStep (((Int((UserControl.ScaleWidth - 8) * (PValue - MinVal) / (MaxVal - MinVal) / (SeperatorWidth + StepLength)) + 1) - 1) * StepLength + ((Int((UserControl.ScaleWidth - 8) * (PValue - MinVal) / (MaxVal - MinVal) / (SeperatorWidth + StepLength)) + 1) - 1) * SeperatorWidth) + 4, 3, UserControl.ScaleHeight - 4, ((UserControl.ScaleWidth - 8) * (PValue - MinVal) / (MaxVal - MinVal)) - Int((UserControl.ScaleWidth - 8) * (PValue - MinVal) / (MaxVal - MinVal) / (SeperatorWidth + StepLength)) * (SeperatorWidth + StepLength)
End Sub

Public Function BlendColors(ByVal Color1 As Long, ByVal Color2 As Long, ByVal Percentage As Single) As Long
On Error Resume Next
    
Dim R(2) As Integer, G(2) As Integer, B(2) As Integer
Dim fPercentage(2) As Single
Dim DAmt(2) As Single
    
Percentage = SetBound(Percentage, 0, 100)
    
GetRGB R(0), G(0), B(0), Color1
GetRGB R(1), G(1), B(1), Color2
    
DAmt(0) = R(1) - R(0): fPercentage(0) = (DAmt(0) / 100) * Percentage
DAmt(1) = G(1) - G(0): fPercentage(1) = (DAmt(1) / 100) * Percentage
DAmt(2) = B(1) - B(0): fPercentage(2) = (DAmt(2) / 100) * Percentage
    
R(2) = R(0) + fPercentage(0)
G(2) = G(0) + fPercentage(1)
B(2) = B(0) + fPercentage(2)
    
BlendColors = RGB(R(2), G(2), B(2))
End Function

Private Function SetBound(ByVal Num As Single, ByVal MinNum As Single, ByVal MaxNum As Single) As Single
If Num < MinNum Then
    SetBound = MinNum
ElseIf Num > MaxNum Then
    SetBound = MaxNum
Else
    SetBound = Num
End If
End Function

Public Sub GetRGB(R As Integer, G As Integer, B As Integer, ByVal Color As Long)
Dim TempValue As Long
    
TranslateColor Color, 0, TempValue
    
R = TempValue And &HFF&
G = (TempValue And &HFF00&) / 2 ^ 8
B = (TempValue And &HFF0000) / 2 ^ 16
End Sub

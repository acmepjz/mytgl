VERSION 5.00
Begin VB.UserControl LeftRight 
   BackColor       =   &H8000000C&
   CanGetFocus     =   0   'False
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   MousePointer    =   9  'Size W E
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
End
Attribute VB_Name = "LeftRight"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////
'This file is public domain.
'////////////////////////////////

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function SetCursorPos Lib "user32.dll" (ByVal x As Long, ByVal y As Long) As Long

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private nDragX As Long

Public Event Change(ByVal iDelta As Long, ByVal Button As Long, ByVal Shift As Long, ByRef bCancel As Boolean)
Public Event MouseUp(ByVal Button As Long, ByVal Shift As Long)

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim p As POINTAPI
GetCursorPos p
nDragX = p.x
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim p As POINTAPI
Dim iDelta As Long, fDelta As Single
Dim b As Boolean
Dim w As Long
If Button = 0 Then Exit Sub
GetCursorPos p
w = Screen.Width \ Screen.TwipsPerPixelX
iDelta = p.x - nDragX
If Shift And vbAltMask Then
 iDelta = iDelta \ 8&
ElseIf Shift And vbCtrlMask Then
 iDelta = iDelta * 10&
End If
'////new:cursor is out of screen? looks like 3DSMax
If p.x <= 0 Then
 nDragX = nDragX + w '??
 p.x = w - 2
 SetCursorPos p.x, p.y
ElseIf p.x >= w - 1 Then
 nDragX = nDragX - w '??
 p.x = 1
 SetCursorPos p.x, p.y
End If
'////
If iDelta = 0 Then Exit Sub
RaiseEvent Change(iDelta, Button, Shift, b)
If Not b Then nDragX = p.x
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
RaiseEvent MouseUp(Button, Shift)
End Sub

Private Sub UserControl_Paint()
GradientFillRect hdc, 0, 0, ScaleWidth, ScaleHeight, d_Title1, d_Title2, GRADIENT_FILL_RECT_V
End Sub

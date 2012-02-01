VERSION 5.00
Begin VB.UserControl FakeComboBox 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   MousePointer    =   1  'Arrow
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer t2 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1440
      Top             =   2160
   End
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   20
      Left            =   960
      Top             =   2160
   End
   Begin VB.ComboBox cmb0 
      Height          =   315
      Left            =   960
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   1560
      Visible         =   0   'False
      Width           =   2055
   End
End
Attribute VB_Name = "FakeComboBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////
'This file is public domain.
'////////////////////////////////

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function MoveWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal bRepaint As Long) As Long
Private Const WM_SETFONT As Long = &H30
Private Const CB_SHOWDROPDOWN As Long = &H14F
Private Const CB_GETDROPPEDSTATE As Long = &H157
Private Const CB_GETITEMHEIGHT As Long = &H154

Private Declare Function GetFocus Lib "user32.dll" () As Long
Private Declare Function SetFocusAPI Lib "user32.dll" Alias "SetFocus" (ByVal hwnd As Long) As Long

Private bd As Boolean
Private dw As Long, dh As Long
Private pw As Long

Private st As Long

Private cFnt As New CLogFont

Private bm As New cDIBSection

Public Event Click()
Public Event MyGotFocus()
Public Event MyLostFocus()
Public Event KeyDown(ByVal KeyCode As Long, ByVal Shift As Long)
Public Event KeyUp(ByVal KeyCode As Long, ByVal Shift As Long)

Public Property Get PanelWidth() As Long
PanelWidth = pw
End Property

Public Property Let PanelWidth(ByVal n As Long)
If pw <> n Then
 pw = n
 pRedraw
End If
End Property

Public Property Get DropdownWidth() As Long
DropdownWidth = dw
End Property

Public Property Let DropdownWidth(ByVal n As Long)
dw = n
End Property

Public Property Get DropdownHeight() As Long
DropdownHeight = dh
End Property

Public Property Let DropdownHeight(ByVal n As Long)
dh = n
End Property

Public Property Get ListCount() As Long
ListCount = cmb0.ListCount
End Property

Public Property Get ListIndex() As Long
ListIndex = cmb0.ListIndex
End Property

Public Property Let ListIndex(ByVal n As Long)
On Error Resume Next
cmb0.ListIndex = n
pRedraw
End Property

Public Property Get Text() As String
Text = cmb0.Text
End Property

Public Property Let Text(ByVal s As String)
On Error Resume Next
cmb0.List(cmb0.ListIndex) = s
pRedraw
End Property

Public Property Get List(ByVal Index As Long) As String
On Error Resume Next
List = cmb0.List(Index)
End Property

Public Property Let List(ByVal Index As Long, ByVal s As String)
On Error Resume Next
cmb0.List(Index) = s
If Index = cmb0.ListIndex Then pRedraw '??
End Property

Public Property Get itemData(ByVal Index As Long) As Long
On Error Resume Next
itemData = cmb0.itemData(Index)
End Property

Public Property Let itemData(ByVal Index As Long, ByVal n As Long)
On Error Resume Next
cmb0.itemData(Index) = n
End Property

Public Sub AddItem(ByVal s As String, Optional ByVal Index As Long = -1)
On Error Resume Next
If Index < 0 Then
 cmb0.AddItem s
Else
 cmb0.AddItem s, Index
End If
End Sub

Public Sub RemoveItem(ByVal Index As Long)
On Error Resume Next
cmb0.RemoveItem Index
pRedraw
End Sub

Public Sub Clear()
cmb0.Clear
pRedraw
End Sub

Public Property Get BackColor() As OLE_COLOR
BackColor = cmb0.BackColor
End Property

Public Property Let BackColor(ByVal n As OLE_COLOR)
cmb0.BackColor = n
pRedraw
End Property

Public Property Get ForeColor() As OLE_COLOR
ForeColor = cmb0.ForeColor
End Property

Public Property Let ForeColor(ByVal n As OLE_COLOR)
cmb0.ForeColor = n
pRedraw
End Property

Public Property Get Font() As StdFont
Set Font = UserControl.Font
End Property

Public Property Set Font(fnt As StdFont)
Set UserControl.Font = fnt
cFnt.HighQuality = True
Set cFnt.LogFont = fnt
SendMessage cmb0.hwnd, WM_SETFONT, cFnt.Handle, ByVal 0
pRedraw
End Property

Public Property Get BorderStyle() As Boolean
BorderStyle = bd
End Property

Public Property Let BorderStyle(ByVal b As Boolean)
If bd <> b Then
 bd = b
 pRedraw
End If
End Property

Private Sub cmb0_Click()
'???
SetFocusAPI hwnd
RaiseEvent Click
End Sub

Private Sub cmb0_GotFocus()
'???
If SendMessage(cmb0.hwnd, CB_GETDROPPEDSTATE, 0, ByVal 0) = 0 Then
 SetFocusAPI hwnd
End If
End Sub

Private Sub t1_Timer()
Dim p As POINTAPI
If st = 2 Then
 If SendMessage(cmb0.hwnd, CB_GETDROPPEDSTATE, 0, ByVal 0) = 0 Then
  st = 1 '???
  pRedraw
  If GetFocus = cmb0.hwnd Then SetFocusAPI hwnd
 End If
Else
 GetCursorPos p
 ScreenToClient hwnd, p
 If p.x < 0 Or p.x >= ScaleWidth Or p.y < 0 Or p.y >= ScaleHeight Then
  If st <> 0 Then
   st = 0
   pRedraw
  End If
  t1.Enabled = False
  If GetFocus = cmb0.hwnd Then SetFocusAPI hwnd
 End If
End If
End Sub

Private Sub t2_Timer()
Dim h As Long
h = GetFocus
If h <> hwnd And h <> cmb0.hwnd Then
 RaiseEvent MyLostFocus
 t2.Enabled = False
End If
End Sub

Private Sub UserControl_DblClick()
If cmb0.ListIndex >= 0 And cmb0.ListCount > 1 Then
 cmb0.ListIndex = (cmb0.ListIndex + 1) Mod cmb0.ListCount
 pRedraw
 'RaiseEvent Click
End If
End Sub

Private Sub UserControl_GotFocus()
RaiseEvent MyGotFocus
t2.Enabled = True
End Sub

Private Sub UserControl_Initialize()
'
End Sub

Private Sub UserControl_InitProperties()
pw = 11
pInit
End Sub

Private Sub UserControl_KeyDown(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyDown(KeyCode, Shift)
End Sub

Private Sub UserControl_KeyUp(KeyCode As Integer, Shift As Integer)
RaiseEvent KeyUp(KeyCode, Shift)
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
 If x >= ScaleWidth - pw Then
  ShowDropdown
 End If
End If
End Sub

Public Sub ShowDropdown()
Dim w As Long, h As Long
  st = 2
  pRedraw
  'show dropdown
  w = ScaleWidth
  If w < dw Then w = dw
  h = dh
  If h < 1 Then h = 1
  h = h * SendMessage(cmb0.hwnd, CB_GETITEMHEIGHT, -1, ByVal 0)
  MoveWindow cmb0.hwnd, 0, ScaleHeight - cmb0.Height, w, h, 0
  SendMessage cmb0.hwnd, CB_SHOWDROPDOWN, 1, ByVal 0
  SetFocusAPI cmb0.hwnd
  t1.Enabled = True
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 0 Then
 If st = 0 Then
  st = 1
  pRedraw
  t1.Enabled = True
 End If
'Else
End If
End Sub

Private Sub UserControl_Paint()
bm.PaintPicture hdc
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
 Set UserControl.Font = .ReadProperty("Font", UserControl.Font)
 bd = .ReadProperty("BorderStyle", False)
 cmb0.BackColor = .ReadProperty("BackColor", cmb0.BackColor)
 cmb0.ForeColor = .ReadProperty("ForeColor", cmb0.ForeColor)
 pw = .ReadProperty("PanelWidth", 11)
 dw = .ReadProperty("DropdownWidth", 0)
 dh = .ReadProperty("DropdownHeight", 0)
End With
pInit
End Sub

Private Sub UserControl_Resize()
On Error Resume Next
bm.Create ScaleWidth, ScaleHeight
pRedraw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
 .WriteProperty "Font", UserControl.Font
 .WriteProperty "BorderStyle", bd, False
 .WriteProperty "BackColor", cmb0.BackColor
 .WriteProperty "ForeColor", cmb0.ForeColor
 .WriteProperty "PanelWidth", pw, 11
 .WriteProperty "DropdownWidth", dw, 0
 .WriteProperty "DropdownHeight", dh, 0
End With
End Sub

Private Sub pInit()
cFnt.HighQuality = True
Set cFnt.LogFont = UserControl.Font
SendMessage cmb0.hwnd, WM_SETFONT, cFnt.Handle, ByVal 0
UserControl_Resize
End Sub

Private Sub pRedraw()
Dim hbr As Long, r As RECT
Dim clr1 As Long, clr2 As Long
Dim i As Long
Dim w As Long, h As Long
w = bm.Width
h = bm.Height
If bd Then
 clr1 = &H800000
 clr2 = clr1 '&HFF8080 '?
Else
 clr1 = vbWhite
 clr2 = &H800000
End If
r.Right = w
r.Bottom = h
'background
hbr = CreateSolidBrush(TranslateColor(cmb0.BackColor))
FillRect bm.hdc, r, hbr
DeleteObject hbr
'border
r.Right = w
r.Bottom = h
If st Then
 hbr = CreateSolidBrush(clr2)
Else
 hbr = CreateSolidBrush(clr1)
End If
FrameRect bm.hdc, r, hbr
r.Left = w - pw
r.Right = r.Left + 1
r.Bottom = h
FillRect bm.hdc, r, hbr
DeleteObject hbr
'dropdown
Select Case st
Case 0
 clr1 = d_Bar1
 clr2 = d_Bar2
Case 1
 clr1 = d_Hl1
 clr2 = d_Hl2
Case Else
 clr1 = d_Pressed1
 clr2 = d_Pressed2
End Select
GradientFillRect bm.hdc, w - pw + 1, 1, w - 1, h - 1, clr1, clr2, GRADIENT_FILL_RECT_V
r.Left = w - pw \ 2 - 3
r.Right = r.Left + 5
r.Bottom = h \ 2
r.Top = r.Bottom - 1
hbr = CreateSolidBrush(vbBlack)
For i = 1 To 3
 FillRect bm.hdc, r, hbr
 r.Left = r.Left + 1
 r.Top = r.Top + 1
 r.Right = r.Right - 1
 r.Bottom = r.Bottom + 1
Next i
DeleteObject hbr
'text
If Ambient.UserMode Then
 cFnt.DrawTextXP bm.hdc, cmb0.Text, 2, 0, w - pw, h, _
 DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX Or DT_EXPANDTABS Or DT_END_ELLIPSIS, _
 TranslateColor(cmb0.ForeColor), , True
Else
 cFnt.DrawTextXP bm.hdc, Extender.Name, 2, 0, w - pw, h, _
 DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX Or DT_EXPANDTABS Or DT_END_ELLIPSIS, _
 TranslateColor(cmb0.ForeColor), , True
End If
'paint
UserControl_Paint
End Sub

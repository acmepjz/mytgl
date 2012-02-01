VERSION 5.00
Begin VB.UserControl FakeWindowXP0 
   ClientHeight    =   3600
   ClientLeft      =   0
   ClientTop       =   0
   ClientWidth     =   4800
   ControlContainer=   -1  'True
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   FontTransparent =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer tmrMisc 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2880
      Top             =   1920
   End
   Begin VB.Timer tmrSize 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   3360
      Top             =   1440
   End
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   2880
      Top             =   1440
   End
   Begin VB.Image i0 
      Height          =   480
      Left            =   480
      Picture         =   "FakeWindowXP.ctx":0000
      Top             =   480
      Visible         =   0   'False
      Width           =   480
   End
End
Attribute VB_Name = "FakeWindowXP0"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////
'This file is public domain.
'////////////////////////////////

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST As Long = -1
Private Const HWND_TOP As Long = 0
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_HIDEWINDOW As Long = &H80
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOREDRAW As Long = &H8
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_SHOWWINDOW As Long = &H40
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Const GWL_HWNDPARENT As Long = -8
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_TOOLWINDOW As Long = &H80&
Private Const WS_EX_TOPMOST As Long = &H8&

Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOW As Long = 5
Private Const SW_MINIMIZE As Long = 6
Private Const SW_NORMAL As Long = 1

#Const UseFakeMenu = 1
#Const UseFakeTB = 1
#Const UseHookGetWindow = 0

Private Declare Function GetClientRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long

Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long

Private Const TheBackColor = &HC9662A 'outer border color
Private Const TheLineColor = &HF9D4B9 'inner border color
Private Const TheCaptionColor = &HFCE8D8 'caption color
Private Const TheCaptionColor_2 = &HE0A47B 'caption color 2
Private Const TheBackColor2 = &HFEECDD 'start color of back
Private Const TheBackColor2_2 = &HCA7D4E  'end color of back

Private sCaption As String, bCaption As Boolean

Private btnHl As Long, btnHlOld As Long, btnHlMenu As Long
'&h80000001 = close
'&h80000002 = min
'&h80000003 = custom
'&h80000004 = dropdown tab select
'&h80000005 = gripper
'&h80000006 = tab move left
'&h80000007 = tab move right
'&h80000010 = client area
'&h81000000+index = tab

Private bVisible As Boolean

Private bCloseButton As Boolean, pr As Boolean
Private bDrag As Boolean, bFloat As Boolean, OldActiveWindow As Long

Private bMinButton As Boolean, bmIn As Boolean

#If UseFakeMenu Then
Private bCustomButton As Boolean
'dropdown menu position
Private rcCustomButton As RECT, rcDropdown As RECT
#End If

Public Enum enumFakeWindowTabMode
 enumFakeWindowTabModeNormal = 0
 enumFakeWindowTabModeHeader = 1 '???
 #If UseFakeMenu Then
 enumFakeWindowTabModeDropdown = 99
 #End If
End Enum

Private tabMode As Long

Private cFnt As New CLogFont, cFntB As New CLogFont, objFntB As StdFont

Private bm0 As New cDIBSection
Private bm As New cDIBSection

Private xDelta As Long, yDelta As Long

Private ow As Long, oh As Long

Private Type typeFakeWindowTab
 sCaption As String
 sKey As String
 w As Integer
 nFlags As Integer
 '1=enabled
 '2=add separator
 'custom client size
 ww As Long
 hh As Long
End Type

Private tabs() As typeFakeWindowTab, tabc As Long
Private tsld As Long, tleft As Long, tScroll As Long
Private tabCanMoveRight As Boolean

#If UseFakeMenu Then
Private WithEvents objMenu As FakeMenu
Attribute objMenu.VB_VarHelpID = -1
#End If

#If UseFakeTB Then
Private btns() As typeFakeButton, btnc As Long
Private objTB As New IFakeToolbarDraw
Implements IFakeToolbarDraw
#End If

Public Event Click()
Public Event TabClick(ByVal TabIndex As Long, ByVal Key As String)
Public Event MouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long, ByVal ClientWidth As Long, ByVal ClientHeight As Long)
Public Event MouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long, ByVal ClientWidth As Long, ByVal ClientHeight As Long)
Public Event MouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long, ByVal ClientWidth As Long, ByVal ClientHeight As Long)
Public Event CloseButtonClick()
Public Event MinButtonClick()
Public Event Paint(ByVal hdc As Long, ByVal ClientLeft As Long, ByVal ClientTop As Long, ByVal ClientWidth As Long, ByVal ClientHeight As Long)

#If UseFakeMenu Then
Public Sub SetMenu(obj As FakeMenu, Optional ByVal bGetBitmap As Boolean)
Set objMenu = obj
#If UseFakeTB Then
Dim bm1 As cDIBSection, bm2 As cDIBSection
If bGetBitmap Then
 objTB.TransparentColor = obj.fGetBitmap(bm1, bm2)
 objTB.SetBitmap bm1, bm2
 If bm1 Is Nothing Then objTB.PicSize = 16 Else objTB.PicSize = bm1.Height
End If
#End If
End Sub
#End If

#If UseFakeTB Then
Public Sub SetPicture(pic As StdPicture, Optional ByVal TransparentColor As Long = vbGreen)
Dim bm1 As New cDIBSection, bm2 As New cDIBSection
objTB.TransparentColor = TransparentColor
If pic Is Nothing Then
 objTB.SetBitmap Nothing, Nothing
 objTB.PicSize = 16
Else
 bm1.CreateFromPicture pic
 GrayscaleBitmap bm1, bm2, d_Icon_Grayscale, TransparentColor
 objTB.SetBitmap bm1, bm2
 objTB.PicSize = bm1.Height
End If
End Sub
#End If

Public Property Get FakeWindowTabMode() As enumFakeWindowTabMode
FakeWindowTabMode = tabMode
End Property

Public Property Let FakeWindowTabMode(ByVal n As enumFakeWindowTabMode)
If tabMode <> n Then
 tabMode = n
 pRefresh
End If
End Property

Public Property Get Font() As StdFont
Set Font = cFnt.LogFont
End Property

Public Property Set Font(obj As StdFont)
cFnt.HighQuality = True
Set cFnt.LogFont = obj
If objFntB Is Nothing Then pBoldFont
End Property

Public Property Get BoldFont() As StdFont
Set BoldFont = objFntB
End Property

Public Property Set BoldFont(obj As StdFont)
Set objFntB = obj
pBoldFont
End Property

Private Sub pBoldFont()
Dim obj As IFont, obj2 As StdFont
If objFntB Is Nothing Then
 Set obj = cFnt.LogFont
 obj.Clone obj2
 obj2.Bold = True
 cFntB.HighQuality = True
 Set cFntB.LogFont = obj2
Else
 cFntB.HighQuality = True
 Set cFntB.LogFont = objFntB
End If
End Sub

Public Property Get Minimized() As Boolean
Minimized = bmIn
End Property

Public Property Let Minimized(ByVal b As Boolean)
If b <> bmIn Then
 bmIn = b
 'save old size
 ow = ScaleWidth
 If b Then oh = ScaleHeight
 'resize
 If Ambient.UserMode Then
  MoveEx , , , oh
  RaiseEvent MinButtonClick
 Else
  pRedraw
 End If
End If
End Property

Public Property Get AutoDrag() As Boolean
AutoDrag = bDrag
End Property

Public Property Let AutoDrag(ByVal b As Boolean)
bDrag = b
End Property

Public Property Get AutoChangeCaption() As Boolean
AutoChangeCaption = bCaption
End Property

Public Property Let AutoChangeCaption(ByVal b As Boolean)
bCaption = b
End Property

Public Property Get Float() As Boolean
Float = bFloat
End Property

Public Property Let Float(ByVal b As Boolean)
If bFloat <> b Then
 bFloat = b
 pFloat
End If
End Property

Public Sub pFloat()
Dim i As Long, r As RECT
If Not Ambient.UserMode Then Exit Sub
If bFloat Then
 GetWindowRect hwnd, r
 i = GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_TOOLWINDOW
 SetWindowLong hwnd, GWL_EXSTYLE, i
 SetWindowLong hwnd, GWL_HWNDPARENT, Parent.hwnd
 SetParent hwnd, 0
 SetWindowPos hwnd, 0, r.Left, r.Top, 0, 0, SWP_FRAMECHANGED Or SWP_NOSIZE Or SWP_NOZORDER
 SetWindowPos hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE
 OldActiveWindow = GetActiveWindow
 #If UseHookGetWindow Then
 Hook_GetWindow True
 Hook_GetWindow_AddWindow hwnd, Parent.hwnd
 #End If
Else
 GetWindowRect hwnd, r
 SetParent hwnd, ContainerHwnd
 MoveEx r.Left, r.Top, , , True
 #If UseHookGetWindow Then
 Hook_GetWindow_RemoveWindow hwnd
 #End If
End If
End Sub

Private Sub pRefresh()
'recalc size
If tsld > 0 And tsld <= tabc And Not bmIn Then
 pTabClick
End If
'redraw
pRedraw
End Sub

Public Property Get CloseButton() As Boolean
CloseButton = bCloseButton
End Property

Public Property Let CloseButton(ByVal b As Boolean)
If bCloseButton <> b Then
 bCloseButton = b
 pRedraw
End If
End Property

#If UseFakeMenu Then

Public Property Get CustomButton() As Boolean
CustomButton = bCustomButton
End Property

Public Property Let CustomButton(ByVal b As Boolean)
If bCustomButton <> b Then
 bCustomButton = b
 pRedraw
End If
End Property

#End If

Private Sub pMouseDownEvent(ByVal Button As Long, ByVal Shift As Long)
Dim p As POINTAPI, i As Long, w As Long, h As Long
If bmIn Or btnHl <> &H80000010 Then Exit Sub
GetCursorPos p
ScreenToClient hwnd, p
If tabMode = 0 And tabc > 1 Then i = 36 Else i = 19
RaiseEvent MouseDown(Button, Shift, p.x - 3, p.y - i, ScaleWidth - 6, ScaleHeight - i - 3)
End Sub

Private Sub pMouseMoveEvent(ByVal Button As Long, ByVal Shift As Long)
Dim p As POINTAPI, i As Long, w As Long, h As Long
If bmIn Or btnHl <> &H80000010 Then Exit Sub
GetCursorPos p
ScreenToClient hwnd, p
If tabMode = 0 And tabc > 1 Then i = 36 Else i = 19
RaiseEvent MouseMove(Button, Shift, p.x - 3, p.y - i, ScaleWidth - 6, ScaleHeight - i - 3)
End Sub

Private Sub pMouseUpEvent(ByVal Button As Long, ByVal Shift As Long)
Dim p As POINTAPI, i As Long, w As Long, h As Long
If bmIn Or btnHl <> &H80000010 Then Exit Sub
GetCursorPos p
ScreenToClient hwnd, p
If tabMode = 0 And tabc > 1 Then i = 36 Else i = 19
RaiseEvent MouseUp(Button, Shift, p.x - 3, p.y - i, ScaleWidth - 6, ScaleHeight - i - 3)
End Sub

#If UseFakeTB Then

Private Sub IFakeToolbarDraw_Click(ByVal btnIndex As Long, ByVal btnKey As String)
Dim i As Long
If bmIn Then Exit Sub
#If UseFakeMenu Then
If Not objMenu Is Nothing Then
 If tsld > 0 And tsld <= tabc Then
  i = objMenu.FindMenu(tabs(tsld).sKey)
  If i > 0 Then
   objMenu.ClickByIndex i, btnIndex
  End If
 End If
End If
#Else
'TODO:
#End If
End Sub

Private Sub IFakeToolbarDraw_GetButtonSafeArrayData(lpSafeArray As Long, btnc As Long)
'unsupported!
End Sub

Private Sub IFakeToolbarDraw_Paint()
Dim i As Long
If bmIn Then Exit Sub
'calc client area
If tabMode = 0 And tabc > 1 Then i = 36 Else i = 19
'toolbar
#If UseFakeMenu Then
If btnc > 0 Then
 objTB.TheBitmap.PaintPicture bm.hdc, 3, i + 1
End If
#Else
'TODO:
#End If
UserControl_Paint
End Sub

Private Sub IFakeToolbarDraw_SetToolTipText(ByVal s As String)
On Error Resume Next
Extender.ToolTipText = s
End Sub

#End If

#If UseFakeMenu Then
Private Sub objMenu_Click(ByVal idxMenu As Long, ByVal Key As String, ByVal idxButton As Long, ByVal ButtonKey As String, Value As Long)
Dim i As Long
If Key = "____FakeWindow____" + CStr(ObjPtr(Me)) Then
 If Left(ButtonKey, 21) = "____FakeWindowTab____" Then
  i = Val(Mid(ButtonKey, 22))
  If i > 0 And i <= tabc Then
   If (tabs(i).nFlags And 1&) <> 0 And (tsld <> i Or bmIn) Then
    tsld = i
    pTabClick
   End If
  End If
 End If
End If
End Sub
#End If

Private Sub t1_Timer()
Dim p As POINTAPI
If btnHl = &H80000005 And bDrag Then Exit Sub
#If UseFakeMenu Then
If btnHlMenu Then
 If Not objMenu Is Nothing Then
  If objMenu.MenuWindowCount <= 0 Or objMenu.UserData <> ObjPtr(Me) Then
   btnHlMenu = 0
   pRedraw
  End If
 End If
End If
#End If
GetCursorPos p
ScreenToClient hwnd, p
If p.x < 0 Or p.y < 0 Or p.x >= ScaleWidth Or p.y >= ScaleHeight Then
 If btnHl Then
  btnHl = 0
  pRedraw
 End If
 #If UseFakeMenu Then
 If btnHlMenu = 0 Then
 #End If
 t1.Enabled = False
 #If UseFakeMenu Then
 End If
 #End If
End If
End Sub

Private Sub tmrSize_Timer()
If bmIn Then
 bmIn = False
 Minimized = True
End If
If bFloat Then pFloat
tmrSize.Enabled = False
End Sub

Private Sub tmrMisc_Timer()
'fake z-order :-3
If bFloat Then
 If OldActiveWindow = GetActiveWindow And bVisible Then
  SetWindowPos hwnd, HWND_TOP, 0, 0, 0, 0, SWP_NOACTIVATE Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_SHOWWINDOW
 Else
  ShowWindow hwnd, 0
 End If
Else
 If bVisible Then
  ShowWindow hwnd, SW_SHOW
 Else
  ShowWindow hwnd, 0
 End If
End If
'scroll
If (btnHl = &H80000006 Or btnHl = &H80000007) And pr Then
 If tScroll = 0 Or tScroll > 5 Then
  If btnHl = &H80000006 Then
   If tleft > 1 Then
    tleft = tleft - 1
    pRedraw
   End If
  Else
   If tabCanMoveRight And tleft < tabc Then
    tleft = tleft + 1
    pRedraw
   End If
  End If
 End If
 tScroll = tScroll + 1
End If
'fake toolbar
#If UseFakeTB Then
If btnc > 0 Then objTB.OnTimer btns, btnc
#End If
'fake visible
On Error Resume Next
Extender.Visible = bVisible
End Sub

Private Sub UserControl_Click()
Dim i As Long
If btnHl = &H80000001 And bCloseButton Then
 RaiseEvent CloseButtonClick
ElseIf btnHl = &H80000002 And bMinButton Then
 Minimized = Not bmIn
ElseIf btnHl > &H81000000 And btnHl < &H82000000 Then
 i = btnHl - &H81000000
 If i <= tabc Then
  If (tabs(i).nFlags And 1&) <> 0 And (tsld <> i Or bmIn) Then
   tsld = i
   pTabClick
  End If
 End If
Else
 #If UseFakeTB Then
 If btnc > 0 And Not bmIn Then objTB.OnClick btns, btnc
 #End If
 RaiseEvent Click
End If
End Sub

Public Property Get SelectedTab() As Long
If Ambient.UserMode Then
 SelectedTab = tsld
End If
End Property

Public Property Let SelectedTab(ByVal n As Long)
If Ambient.UserMode Then
 If n > 0 And n <= tabc And (n <> tsld Or bmIn) Then
  tsld = n
  pTabClick
 End If
End If
End Property

Private Sub pTabClick()
Dim i As Long
If tsld < 1 Or tsld > tabc Then Exit Sub
If bCaption Then sCaption = tabs(tsld).sCaption
'show toolbar?
#If UseFakeTB Then
#If UseFakeMenu Then
If Not objMenu Is Nothing Then
 objMenu.fGetMenuData objMenu.FindMenu(tabs(tsld).sKey), btns, btnc
 If btnc > 0 Then
  '///???
  Set objTB.MenuObject = objMenu
  objTB.SetCallback Me
  With objTB.TheBitmap
   If .hdc = 0 Then .Create 8, 8
  End With
  '///
  Set objTB.Font = cFnt.LogFont
  i = objTB.GetWidth(btns, btnc)
  If i < 96 Then i = 96
  tabs(tsld).ww = i
  tabs(tsld).hh = 21
  If tabMode = 0 And tabc > 1 Then i = 36 Else i = 19
  objTB.Resize 3, i + 1, tabs(tsld).ww, 20, hwnd, btns, btnc, False
 End If
End If
#Else
'TODO:
#End If
#End If
'event
RaiseEvent TabClick(tsld, tabs(tsld).sKey)
'resize
If bmIn Then Minimized = False Else pRedraw
If tabs(tsld).ww > 0 Or tabs(tsld).hh > 0 Then
 MoveEx , , tabs(tsld).ww, tabs(tsld).hh, , True
End If
End Sub

Public Function FindTab(ByVal Key As String) As Long
Dim i As Long
For i = 1 To tabc
 If tabs(i).sKey = Key Then
  FindTab = i
  Exit Function
 End If
Next i
End Function

Private Sub UserControl_DblClick()
If btnHl = &H80000005 And bMinButton Then
 Minimized = Not bmIn
#If UseFakeTB Then
ElseIf btnc > 0 And Not bmIn Then
 objTB.OnDblClick btns, btnc
#End If
End If
End Sub

Public Property Get IsVisible() As Boolean
IsVisible = bVisible
End Property

Public Property Let IsVisible(ByVal b As Boolean)
bVisible = b
On Error Resume Next
If Ambient.UserMode Then Extender.ZOrder
Extender.Visible = b
End Property

Private Sub UserControl_Initialize()
bm0.CreateFromPicture i0.Picture
tleft = 1
End Sub

Private Sub UserControl_InitProperties()
sCaption = Extender.Name
cFnt.HighQuality = True
Set cFnt.LogFont = UserControl.Font
bVisible = True
bCaption = True
pBoldFont
pInit
End Sub

Public Sub MoveEx(Optional ByVal Left As Long = &H80000000, Optional ByVal Top As Long = &H80000000, Optional ByVal Width As Long, Optional ByVal Height As Long, Optional ByVal IsScreenPos As Boolean, Optional ByVal IsClientSize As Boolean)
Dim r As RECT, r2 As RECT
Dim p As POINTAPI
'get rect
If bFloat Then
 r.Right = Screen.Width / Screen.TwipsPerPixelX
 r.Bottom = Screen.Height / Screen.TwipsPerPixelY
Else
 GetClientRect ContainerHwnd, r
 ClientToScreen ContainerHwnd, p
End If
GetWindowRect hwnd, r2
'calc new size
If Width > 0 Then
 r2.Right = Width
 If IsClientSize Then r2.Right = r2.Right + 6
Else
 r2.Right = r2.Right - r2.Left
End If
If bmIn Then
 r2.Bottom = 22
ElseIf Height > 0 Then
 r2.Bottom = Height
 If IsClientSize Then _
 If tabMode = 0 And tabc > 1 Then r2.Bottom = r2.Bottom + 39 Else r2.Bottom = r2.Bottom + 22
Else
 r2.Bottom = r2.Bottom - r2.Top
End If
'get pos
If Left = &H80000000 Then
 r2.Left = r2.Left - p.x
ElseIf IsScreenPos Then
 r2.Left = Left - p.x
Else
 r2.Left = Left
End If
If Top = &H80000000 Then
 r2.Top = r2.Top - p.y
ElseIf IsScreenPos Then
 r2.Top = Top - p.y
Else
 r2.Top = Top
End If
'calc container size
r.Right = r.Right - r2.Right
r.Bottom = r.Bottom - r2.Bottom
'calc new pos
If r2.Left < 0 Then r2.Left = 0 Else If r2.Left > r.Right Then r2.Left = r.Right
If r2.Top < 0 Then r2.Top = 0 Else If r2.Top > r.Bottom Then r2.Top = r.Bottom
'move
SetWindowPos hwnd, 0, r2.Left, r2.Top, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
'size
Size r2.Right * Screen.TwipsPerPixelX, r2.Bottom * Screen.TwipsPerPixelY
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
btnHlOld = btnHl
pr = True
tScroll = 0
't1.Enabled = False
On Error Resume Next
Extender.ZOrder
On Error GoTo 0
If btnHl Then pRedraw
If btnHl = &H80000005 And bDrag Then
 pDragStart
#If UseFakeMenu Then
ElseIf btnHl = &H80000003 And bCustomButton Then
 'TODO:add or remove button
ElseIf btnHl = &H80000004 And tabMode = 99 And tabc > 1 Then
 'dropdown menu
 If Not objMenu Is Nothing Then
  If btnHlMenu = &H80000004 Then
   objMenu.UnpopupMenu
  Else
   pDropdownTab
  End If
  Exit Sub
 End If
#End If
#If UseFakeTB Then
ElseIf btnc > 0 And Not bmIn Then
 objTB.OnMouseDown btns, btnc, Button, Shift, x, y
#End If
End If
'///
pMouseDownEvent Button, Shift
End Sub

Private Sub pDropdownTab()
Const sKey2 As String = "____FakeWindowTab____"
Dim i As Long, j As Long
Dim p As POINTAPI
Dim sKey As String
'redraw
btnHlMenu = &H80000004
pRedraw
'start
sKey = "____FakeWindow____" + CStr(ObjPtr(Me))
i = objMenu.FindMenu(sKey)
If i = 0 Then
 i = objMenu.AddMenu(sKey)
Else
 objMenu.DestroyMenuButtons i
End If
For j = 1 To tabc
 If (tabs(j).nFlags And 2&) And j > 1 Then _
 objMenu.AddButtonByIndex i, , , , , fbttSeparator
 objMenu.AddButtonByIndex i, , sKey2 + CStr(j), tabs(j).sCaption, , , fbtfDisabled And (tabs(j).nFlags And 1&) = 0, , , , , , j = tsld
Next j
ClientToScreen hwnd, p
objMenu.PopupMenuEx sKey, rcDropdown.Left + p.x, rcDropdown.Top + p.y, _
rcDropdown.Right - rcDropdown.Left, rcDropdown.Bottom - rcDropdown.Top, , , , , ObjPtr(Me)
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, idxHl As Long
Dim w As Long, h As Long, r As RECT
w = ScaleWidth
h = ScaleHeight
'dragging?
If btnHl = &H80000005 And bDrag And Button = 1 Then
 pDragging
 Exit Sub
End If
'hit test
If y >= 4 And y < 18 Then
 r.Right = w - 4
 r.Left = r.Right - 14
 If bCloseButton Then
  If x >= r.Left And x < r.Right And (btnHlOld = &H80000001 Or Button = 0) Then idxHl = &H80000001
  r.Right = r.Left
  r.Left = r.Right - 14
 End If
 If bMinButton Then
  If x >= r.Left And x < r.Right And (btnHlOld = &H80000002 Or Button = 0) Then idxHl = &H80000002
  r.Right = r.Left
  r.Left = r.Right - 14
 End If
 #If UseFakeMenu Then
 If bCustomButton Then
  If x >= r.Left And x < r.Right And (btnHlOld = &H80000003 Or Button = 0) Then idxHl = &H80000003
  r.Right = r.Left
  r.Left = r.Right - 14
 End If
 'gripper & dropdown menu
 If tabMode = 99 And tabc > 1 Then
  If x >= 8 And x < r.Right And (btnHlOld = &H80000004 Or Button = 0) Then idxHl = &H80000004
  If x >= 2 And x < 8 And (btnHlOld = &H80000005 Or Button = 0) Then idxHl = &H80000005
 Else
 #End If
  If tabMode = 1 And tabc > 1 Then
   If x >= 2 And x < 8 Then
    If btnHlOld = &H80000005 Or Button = 0 Then idxHl = &H80000005
   Else
    i = pTabHitTest(8, 3, r.Right, 19, x, y)
    If btnHlOld <> i And Button <> 0 Then i = 0
    If i <> 0 Then idxHl = i
   End If
  Else
   If x >= 2 And x < r.Right And (btnHlOld = &H80000005 Or Button = 0) Then idxHl = &H80000005
  End If
 #If UseFakeMenu Then
 End If
 #End If
End If
'tab hit test
If tabMode = 0 And tabc > 1 And Not bmIn Then
 r.Bottom = 36
 If r.Bottom > h - 3 Then r.Bottom = h - 3
 i = pTabHitTest(4, 19, w - 4, r.Bottom, x, y)
 If btnHlOld <> i And Button <> 0 Then i = 0
 If i <> 0 Then idxHl = i
End If
'tool bar hit test
#If UseFakeTB Then
If btnc > 0 And Not bmIn Then
 objTB.OnMouseMove btns, btnc, Button, Shift, x, y
Else
#End If
'client area
 If tabMode = 0 And tabc > 1 Then i = 36 Else i = 19
 If (x >= 3 And y >= i And x < w - 3 And y < h - 3 And Button = 0) Or _
 (btnHlOld = &H80000010 And Button <> 0) Then idxHl = &H80000010
#If UseFakeTB Then
End If
#End If
'highlight changed?
If idxHl <> btnHl Then
 If idxHl = &H80000005 And bDrag Then
  MousePointer = vbSizeAll
 Else
  MousePointer = vbDefault
 End If
 #If UseFakeMenu Then
 t1.Enabled = idxHl <> 0 Or btnHlMenu <> 0
 #Else
 t1.Enabled = idxHl
 #End If
 '///
 btnHl = idxHl
 pRedraw
End If
'///
pMouseMoveEvent Button, Shift
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
btnHlOld = 0
pr = False
If btnHl Then pRedraw
'///
#If UseFakeTB Then
If btnc > 0 And Not bmIn Then objTB.OnMouseUp btns, btnc, Button, Shift, x, y
#End If
'///
pMouseUpEvent Button, Shift
End Sub

Private Sub UserControl_Paint()
bm.PaintPicture hdc
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
 sCaption = .ReadProperty("Caption", Extender.Name)
 bCloseButton = .ReadProperty("CloseButton", False)
 bMinButton = .ReadProperty("MinButton", False)
 bmIn = .ReadProperty("Minimized", False)
 bDrag = .ReadProperty("AutoDrag", False)
 cFnt.HighQuality = True
 Set cFnt.LogFont = .ReadProperty("Font", UserControl.Font)
 Set objFntB = .ReadProperty("BoldFont", Nothing)
 #If UseFakeMenu Then
 bCustomButton = .ReadProperty("CustomButton", False)
 #End If
 bFloat = .ReadProperty("Float", False)
 tabMode = .ReadProperty("FakeWindowTabMode", 0)
 bVisible = .ReadProperty("IsVisible", True)
 bCaption = .ReadProperty("AutoChangeCaption", True)
End With
pBoldFont
pInit
End Sub

Public Property Get Caption() As String
Caption = sCaption
End Property

Public Property Let Caption(ByVal s As String)
sCaption = s
pRedraw
End Property

Private Sub pInit()
tmrSize.Enabled = Ambient.UserMode
tmrMisc.Enabled = Ambient.UserMode
BackColor = TheBackColor
UserControl_Resize
End Sub

Private Sub UserControl_Resize()
Dim w As Long, h As Long
'////////
w = ScaleWidth
h = ScaleHeight
If w <> bm.Width Or h <> bm.Height Then bm.Create w, h
'////////redraw!!!
pRedraw
End Sub

Private Function pTabHitTest(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal x As Long, ByVal y As Long) As Long
Dim i As Long, w As Long, ww As Long
Dim b As Boolean
'rect test
If x < Left Or y < Top Or x >= Right Or y >= Bottom Then Exit Function
'move left button
If tleft > 1 Then
 If x >= Left And y > Top And x < Left + 14 And y <= Top + 14 Then
  pTabHitTest = &H80000006
  Exit Function
 End If
 Left = Left + 14
End If
ww = Right - 15
If tabMode = 0 Then
 If tleft > 1 Then Left = Left + 1 Else Left = Left + 4
 ww = ww - 1
End If
'tab
b = tleft <= tabc
For i = tleft To tabc
 w = tabs(i).w
 If Left + w > ww Then
  If b Then
   w = ww - Left
   If i = tabc Then If tabMode Then w = w + 14 Else w = w + 11
  ElseIf Left + w > ww + 14 Or i < tabc Then
   b = True
   Exit For
  End If
 End If
 b = False
 'back
 If x >= Left And y >= Top And x <= Left + w And y < Bottom Then
  pTabHitTest = &H81000000 + i
  Exit Function
 End If
 'next
 Left = Left + w
Next i
'new!!! gripper
If tabMode = 1 Then
 If x >= Left And y > Top And y <= Top + 14 Then
  ww = Right
  If b Then ww = ww - 14
  If x < ww Then
   pTabHitTest = &H80000005
   Exit Function
  End If
 End If
End If
'move right button
If b Then
 If x >= Right - 14 And y > Top And x < Right And y <= Top + 14 Then
  pTabHitTest = &H80000007
  Exit Function
 End If
End If
End Function

Private Sub pDrawTab(ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
Dim i As Long, w As Long, ww As Long
Dim r As RECT
Dim hbr As Long, hbrBorder As Long
Dim b As Boolean
tabCanMoveRight = False
If tabMode Then i = TheCaptionColor_2 Else i = TheBackColor
hbrBorder = CreateSolidBrush(i)
'move left button
If tleft > 1 Then
 r.Left = Left
 r.Top = Top + 1
 If btnHl = &H80000006 Then
  r.Right = r.Left + 14
  r.Bottom = r.Top + 14
  If pr Then
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
  Else
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
  End If
  hbr = CreateSolidBrush(d_Border)
  FrameRect bm.hdc, r, hbr
  DeleteObject hbr
 End If
 TransparentBlt bm.hdc, r.Left, r.Top, 14, 14, bm0.hdc, 1, 1, 14, 14, vbGreen
 Left = Left + 14
End If
ww = Right - 15
If tabMode = 0 Then
 If tleft > 1 Then Left = Left + 1 Else Left = Left + 4
 ww = ww - 1
 r.Left = 2
 r.Top = Bottom - 1
 r.Right = Left
 r.Bottom = Bottom
 FillRect bm.hdc, r, hbrBorder
End If
'draw tab
b = True
For i = tleft To tabc
 w = tabs(i).w
 If Left + w > ww Then
  If b Then
   w = ww - Left
   If i = tabc Then If tabMode Then w = w + 14 Else w = w + 11
  ElseIf Left + w > ww + 14 Or i < tabc Then
   tabCanMoveRight = True
   Exit For
  End If
 End If
 b = False
 'back
 If i = tsld Then
  If tabMode Then
   r.Left = Left + 1
   r.Top = Top + 1
   r.Right = Left + w
   r.Bottom = Bottom
   hbr = CreateSolidBrush(TheBackColor2)
   FillRect bm.hdc, r, hbr
   DeleteObject hbr
  End If
  If &H81000000 + i = btnHl And (tabs(i).nFlags And 1&) <> 0 Then
   GradientFillRect bm.hdc, Left + 1, Top + 1, Left + w, Bottom - 1, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
  End If
 Else
  If &H81000000 + i = btnHl And (tabs(i).nFlags And 1&) <> 0 Then
   If pr Then
    GradientFillRect bm.hdc, Left + 1, Top + 1, Left + w, Bottom - 1, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
   Else
    GradientFillRect bm.hdc, Left + 1, Top + 1, Left + w, Bottom - 1, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
   End If
  Else
   GradientFillRect bm.hdc, Left + 1, Top + 1, Left + w, Bottom - 1, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_V
  End If
 End If
 'border
 If i = tsld Or tabMode Then
  r.Left = Left
  r.Top = Top
  r.Right = Left + 1
  r.Bottom = Bottom
  FillRect bm.hdc, r, hbrBorder
  r.Right = Left + w + 1
  r.Bottom = Top + 1
  FillRect bm.hdc, r, hbrBorder
  r.Left = r.Right - 1
  r.Bottom = Bottom
  FillRect bm.hdc, r, hbrBorder
 Else
  r.Left = Left
  r.Top = Top
  r.Right = Left + w + 1
  r.Bottom = Bottom
  FrameRect bm.hdc, r, hbrBorder
 End If
 'caption
 cFnt.DrawTextXP bm.hdc, tabs(i).sCaption, Left, Top, w, Bottom - Top, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS, _
 d_TextDis And (tabs(i).nFlags And 1&) = 0, , True
 'next
 Left = Left + w
Next i
If tabMode = 0 Then
 r.Left = Left
 r.Top = Bottom - 1
 r.Right = Right + 2
 r.Bottom = Bottom
 FillRect bm.hdc, r, hbrBorder
End If
'move right button
If tabCanMoveRight Then
 r.Left = Right - 14
 r.Top = Top + 1
 If btnHl = &H80000007 Then
  r.Right = Right
  r.Bottom = r.Top + 14
  If pr Then
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
  Else
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
  End If
  hbr = CreateSolidBrush(d_Border)
  FrameRect bm.hdc, r, hbr
  DeleteObject hbr
 End If
 TransparentBlt bm.hdc, r.Left, r.Top, 14, 14, bm0.hdc, 17, 1, 14, 14, vbGreen
End If
'over
DeleteObject hbrBorder
End Sub

Private Sub pRedraw(Optional ByVal bPaint As Boolean = True)
Dim w As Long, h As Long
Dim hbr As Long, r As RECT
w = bm.Width
h = bm.Height
r.Right = w
r.Bottom = h
'back
hbr = CreateSolidBrush(TheBackColor)
FillRect bm.hdc, r, hbr
DeleteObject hbr
'border
hbr = CreateSolidBrush(TheLineColor)
r.Left = 2
r.Top = 2
r.Right = w - 2
r.Bottom = h - 2
FrameRect bm.hdc, r, hbr
DeleteObject hbr
SetPixelV bm.hdc, 2, 2, TheBackColor
SetPixelV bm.hdc, w - 3, 2, TheBackColor
'caption
GradientFillRect bm.hdc, 3, 3, w - 3, 19, TheCaptionColor, TheCaptionColor_2, GRADIENT_FILL_RECT_V
'caption button
r.Top = 4
r.Bottom = 18
r.Right = w - 4
r.Left = r.Right - 14
hbr = CreateSolidBrush(d_Border)
If bCloseButton Then
 If btnHl = &H80000001 Then
  If pr Then
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
  Else
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
  End If
  FrameRect bm.hdc, r, hbr
 End If
 TransparentBlt bm.hdc, r.Left + 3, r.Top + 3, 8, 8, bm0.hdc, 16, 24, 8, 8, vbGreen
 r.Right = r.Left
 r.Left = r.Right - 14
End If
If bMinButton Then
 If btnHl = &H80000002 Then
  If pr Then
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
  Else
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
  End If
  FrameRect bm.hdc, r, hbr
 End If
 TransparentBlt bm.hdc, r.Left + 3, r.Top + 3, 8, 8, bm0.hdc, 16 + (bmIn And 8&), 16, 8, 8, vbGreen
 r.Right = r.Left
 r.Left = r.Right - 14
End If
#If UseFakeMenu Then
If bCustomButton Then
 rcCustomButton = r
 If btnHl = &H80000003 Or btnHlMenu = &H80000003 Then
  If btnHlMenu = &H80000003 Then
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_V
  ElseIf pr Then
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
  Else
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
  End If
  FrameRect bm.hdc, r, hbr
 End If
 TransparentBlt bm.hdc, r.Left + 2, r.Top + 5, 9, 5, bm0.hdc, 23, 25, 9, 5, vbGreen
 r.Right = r.Left
 r.Left = r.Right - 14
End If
'caption text & dropdown menu
If tabMode = 99 And tabc > 1 Then
 r.Left = 8
 rcDropdown = r
 If btnHl = &H80000004 Or btnHlMenu = &H80000004 Then
  If btnHlMenu = &H80000004 Then
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_V
  ElseIf pr Then
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
  Else
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
  End If
  FrameRect bm.hdc, r, hbr
 End If
 TransparentBlt bm.hdc, r.Right - 12, r.Top + 5, 9, 5, bm0.hdc, 23, 25, 9, 5, vbGreen
 'gripper
 TransparentBlt bm.hdc, 4, 6, 3, 11, bm0.hdc, 0, 21, 3, 11, vbGreen
 cFntB.DrawTextXP bm.hdc, sCaption, 10, 3, r.Right - 24, 16, DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS, vbBlack, , True
Else
#End If
 If tabMode = 1 And tabc > 1 Then
  'header-tab mode
  pDrawTab 8, 3, r.Right, 19
  'gripper
  TransparentBlt bm.hdc, 4, 6, 3, 11, bm0.hdc, 0, 21, 3, 11, vbGreen
 Else
  cFntB.DrawTextXP bm.hdc, sCaption, 4, 3, r.Left, 16, DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS, vbBlack, , True
 End If
#If UseFakeMenu Then
End If
#End If
DeleteObject hbr
If Not bmIn Then
 'form back
 r.Left = 3
 r.Top = 19
 r.Right = w - 3
 r.Bottom = 35 '19+16
 If r.Bottom < h - 3 Then
  GradientFillRect bm.hdc, 3, r.Bottom, w - 3, h - 3, TheBackColor2, TheBackColor2_2, GRADIENT_FILL_RECT_V
 Else
  r.Bottom = h - 3
 End If
 hbr = CreateSolidBrush(TheBackColor2)
 FillRect bm.hdc, r, hbr
 DeleteObject hbr
 #If UseFakeTB Then
 If (tabMode = 0 And tabc > 1) Or btnc = 0 Then
 #End If
 TransparentBlt bm.hdc, r.Left, r.Top, 5, 5, bm0.hdc, 0, 16, 5, 5, vbGreen
 TransparentBlt bm.hdc, r.Right - 5, r.Top, 5, 5, bm0.hdc, 4, 16, 5, 5, vbGreen
 #If UseFakeTB Then
 End If
 #End If
 'tab
 If tabMode = 0 And tabc > 1 Then
  r.Bottom = r.Bottom + 1
  If r.Bottom > h - 3 Then r.Bottom = h - 3
  pDrawTab 4, 19, w - 4, r.Bottom
 End If
 'calc client area
 If tabMode = 0 And tabc > 1 Then r.Top = 36 Else r.Top = 19
 'toolbar
 #If UseFakeTB Then
 #If UseFakeMenu Then
 If btnc > 0 Then
  objTB.TheBitmap.PaintPicture bm.hdc, 3, r.Top + 1
 End If
 #Else
 'TODO:
 #End If
 #End If
 'owner draw TODO:
 RaiseEvent Paint(bm.hdc, 3, r.Top, w - 6, h - r.Top - 3)
End If
'over
If bPaint Then UserControl_Paint
End Sub

Private Sub UserControl_Terminate()
#If UseFakeMenu Then
Set objMenu = Nothing
#End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
 .WriteProperty "Caption", sCaption, Extender.Name
 .WriteProperty "CloseButton", bCloseButton, False
 .WriteProperty "MinButton", bMinButton, False
 .WriteProperty "AutoDrag", bDrag, False
 .WriteProperty "Minimized", bmIn, False
 'new
 .WriteProperty "Font", cFnt.LogFont, UserControl.Font
 .WriteProperty "BoldFont", objFntB, Nothing
 #If UseFakeMenu Then
 .WriteProperty "CustomButton", bCustomButton, False
 #End If
 .WriteProperty "Float", bFloat, False
 .WriteProperty "FakeWindowTabMode", tabMode, 0
 .WriteProperty "IsVisible", bVisible, True
 .WriteProperty "AutoChangeCaption", bCaption, True
End With
End Sub

Public Property Get MinButton() As Boolean
MinButton = bMinButton
End Property

Public Property Let MinButton(ByVal b As Boolean)
If bMinButton <> b Then
 bMinButton = b
 pRedraw
End If
End Property

Public Sub Refresh()
pRefresh
End Sub

Private Sub pDragStart()
Dim r As RECT, p As POINTAPI
GetWindowRect hwnd, r
p.x = r.Left
p.y = r.Top
If Not bFloat Then ScreenToClient ContainerHwnd, p
r.Left = p.x
r.Top = p.y
GetCursorPos p
xDelta = r.Left - p.x
yDelta = r.Top - p.y
End Sub

Public Sub DragStart()
If bDrag Then
 MousePointer = vbSizeAll
 btnHl = &H80000005
 pDragStart
 SetCapture hwnd
End If
End Sub

Private Sub pDragging()
Dim p As POINTAPI
Dim r As RECT
GetCursorPos p
p.x = p.x + xDelta
p.y = p.y + yDelta
If bFloat Then
 r.Right = Screen.Width / Screen.TwipsPerPixelX
 r.Bottom = Screen.Height / Screen.TwipsPerPixelY
Else
 GetClientRect ContainerHwnd, r
End If
r.Right = r.Right - ScaleWidth
r.Bottom = r.Bottom - ScaleHeight
If p.x < 0 Then p.x = 0 Else If p.x >= r.Right Then p.x = r.Right
If p.y < 0 Then p.y = 0 Else If p.y >= r.Bottom Then p.y = r.Bottom
SetWindowPos hwnd, 0, p.x, p.y, 0, 0, SWP_NOSIZE Or SWP_NOZORDER Or SWP_NOACTIVATE
End Sub

Private Sub pCalcTabWidth(d As typeFakeWindowTab)
Dim w As Long
cFnt.DrawTextXP bm.hdc, d.sCaption, 0, 0, w, 0, DT_SINGLELINE Or DT_CALCRECT
d.w = w + 6
End Sub

Public Sub AddTab(ByVal Caption As String, Optional ByVal Key As String, Optional ByVal Index As Long, Optional ByVal Enabled As Boolean = True, Optional ByVal AddSeparator As Boolean, Optional ByVal ClientWidth As Long, Optional ByVal ClientHeight As Long)
Dim i As Long
tabc = tabc + 1
ReDim Preserve tabs(1 To tabc)
If Index = 0 Or Index >= tabc Then
 i = tabc
Else
 For i = tabc - 1 To Index Step -1
  tabs(i + 1) = tabs(i)
 Next i
 i = Index
End If
If tsld >= i Then tsld = tsld + 1
With tabs(i)
 .sCaption = Caption
 .sKey = Key
 .nFlags = (Enabled And 1&) Or (AddSeparator And 2&)
 .ww = ClientWidth
 .hh = ClientHeight
End With
pCalcTabWidth tabs(i)
pRedraw
End Sub

Public Sub RemoveTab(ByVal Index As Long)
Dim i As Long
If Index <= 0 Or Index > tabc Then Exit Sub
If tabc = 1 Then
 Erase tabs
 tabc = 0
 tsld = 0
 tleft = 1
Else
 For i = Index To tabc - 1
  tabs(i) = tabs(i + 1)
 Next i
 tabc = tabc - 1
 ReDim Preserve tabs(1 To tabc)
 If tleft >= Index And tleft > 1 Then tleft = tleft - 1
 If tsld = Index Then
  If tsld > 1 Then tsld = tsld - 1
  pTabClick
 ElseIf tsld > Index Then
  tsld = tsld - 1
 End If
End If
pRedraw
End Sub

Public Sub ClearTab()
Erase tabs
tabc = 0
tsld = 0
tleft = 1
pRedraw
End Sub

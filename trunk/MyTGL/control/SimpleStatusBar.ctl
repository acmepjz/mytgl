VERSION 5.00
Begin VB.UserControl SimpleStatusBar 
   Alignable       =   -1  'True
   CanGetFocus     =   0   'False
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
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Windowless      =   -1  'True
   Begin VB.Timer tmrRefresh 
      Enabled         =   0   'False
      Interval        =   1000
      Left            =   2640
      Top             =   1560
   End
   Begin VB.Image i1 
      Height          =   180
      Left            =   1440
      Picture         =   "SimpleStatusBar.ctx":0000
      Top             =   2400
      Visible         =   0   'False
      Width           =   180
   End
End
Attribute VB_Name = "SimpleStatusBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////
'This file is public domain.
'////////////////////////////////

Private hwdParent As Long
Private Const WM_NCHITTEST As Long = &H84
Private Const HTCAPTION As Long = 2
Private Const HTBOTTOMRIGHT As Long = 17

Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private bm As New cDIBSection
Private bmd As New cDIBSection

Private fnt As StdFont
Private cFnt As New CLogFont

Public Enum enumSBStyle
 sbNormal = 0
 sbProgressBar = 1
 sbTime = 2
 sbDate = 3
 sbCustomInfo = 98
 sbOwnerDraw = 99
End Enum

Private Type typeSBPanel
 Caption As String
 minw As Long
 x As Long
 w As Long
 en As Boolean
 vs As Boolean
 st As Long
 'progressbar!!
 'mn As Long '=0
 mx As Long
 v As Long
End Type

Private ps() As typeSBPanel, pc As Long
Private pRun As Boolean, ox As Long, oy As Long
Private grip As Boolean, simp As String

Public Event PanelClick(ByVal PanelIndex As Long)
Public Event GetInfo(ByVal PanelIndex As Long, ByRef s As String)

Implements iSubclass
Private cSub As New cSubclass

Private Sub iSubclass_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)

Dim p As POINTAPI, i As Long
If grip Then
 'get pos
 p.y = lParam
 p.x = p.y And &HFFFF&
 p.y = (p.y And &HFFFF0000) \ &H10000
 ScreenToClient hwdParent, p
 'check pos
 On Error Resume Next
 Err.Clear
 i = Parent.ScaleMode
 If Err.Number <> 0 Or i = vbTwips Then 'twip
  p.x = p.x - Extender.Left \ 15
  p.y = p.y - Extender.Top \ 15
 ElseIf i = vbPixels Then 'pixel
  p.x = p.x - Extender.Left
  p.y = p.y - Extender.Top
 Else 'unknown!!!
  Exit Sub
 End If
 On Error GoTo 0
 'check area
 If p.x >= bm.Width - 16 And p.x < bm.Width And p.y >= bm.Height - 16 And p.y < bm.Height Then
  lReturn = HTBOTTOMRIGHT
 End If
End If

End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)
'
End Sub

Private Sub tmrRefresh_Timer()
Dim i As Long, s As String
For i = 1 To pc
 With ps(i)
  If .vs Then
   Select Case .st
   Case 2
    .Caption = Format(Now, "hh:mm:ss")
    pRedrawOne i, True
   Case 3
    .Caption = Format(Now, "yyyy-m-d")
    pRedrawOne i, True
   Case 98
    RaiseEvent GetInfo(i, s)
    .Caption = s
    pRedrawOne i, True
   End Select
  End If
 End With
Next i
End Sub

Private Sub UserControl_Click()
Dim i As Long
If simp <> "" Then
 RaiseEvent PanelClick(-1)
Else
 For i = 1 To pc
  With ps(i)
   If .en And .vs Then
    If ox >= .x And ox < .x + .w + 4 And oy >= 0 And oy < bm.Height Then
     RaiseEvent PanelClick(i)
    End If
   End If
  End With
 Next i
End If
End Sub

Private Sub UserControl_Initialize()
'
End Sub

Private Sub UserControl_InitProperties()
Set fnt = UserControl.Font
grip = True
Extender.Align = vbAlignBottom
pInit
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
ox = x
oy = y
End Sub

Private Sub UserControl_Paint()
On Error Resume Next
bm.PaintPicture hdc
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
 Set fnt = .ReadProperty("Font", UserControl.Font)
 grip = .ReadProperty("ShowResizeGripper", True)
 simp = .ReadProperty("SimpleText")
End With
pInit
End Sub

Private Sub UserControl_Resize()
pResize
End Sub

Private Sub UserControl_Terminate()
If pRun Then
 cSub.DelMsg WM_NCHITTEST, MSG_AFTER
 cSub.UnSubclass
End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
 .WriteProperty "Font", fnt, UserControl.Font
 .WriteProperty "ShowResizeGripper", grip, True
 .WriteProperty "SimpleText", simp
End With
End Sub

Private Sub pInit()
cFnt.HighQuality = True
Set cFnt.LogFont = fnt
bmd.CreateFromPicture i1.Picture
pRun = Ambient.UserMode
If pRun Then
 On Error Resume Next
 hwdParent = Parent.hwnd
 If Err.Number <> 0 Then hwdParent = 0
 On Error GoTo 0
 If hwdParent <> 0 Then
  cSub.AddMsg WM_NCHITTEST, MSG_AFTER
  cSub.Subclass hwdParent, Me
 End If
 tmrRefresh.Enabled = True
End If
pResize
End Sub

Private Sub pResize()
bm.Create ScaleWidth, ScaleHeight
pRedrawAll
End Sub

Private Sub pRedrawAll()
Dim i As Long, c As Long
Dim w As Long, lst As Long
Dim r As RECT
If simp = "" And pRun Then
 'calc panel
 For i = 1 To pc
  With ps(i)
   If .vs Then
    w = w + .minw + 6
    .w = .minw
    lst = 0
    If .minw = 0 Then c = c + 1
   End If
  End With
 Next i
 If lst > 0 Then
  If ps(lst).minw = 0 Then w = w - 2 'the last one is sizable??
 End If
 'calc sizable panel
 w = bm.Width - w
 If grip Then w = w - 16
 For i = 1 To pc
  With ps(i)
   If .vs And .minw = 0 Then
    .w = w \ c
    w = w - .w
    c = c - 1
   End If
  End With
 Next i
 'calc left
 w = 0
 For i = 1 To pc
  With ps(i)
   If .vs Then
    .x = w
    w = w + .w + 6
   End If
  End With
 Next i
End If
'draw background
r.Right = bm.Width
r.Bottom = bm.Height
GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_V
If Not pRun Then
 'design time!!!
 cFnt.DrawTextXP bm.hdc, Extender.Name, 2, 0, bm.Width - 20, bm.Height, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
ElseIf simp <> "" Then
 'simple text
 cFnt.DrawTextXP bm.hdc, simp, 2, 0, bm.Width - 20, bm.Height, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
Else
 'draw panel
 For i = 1 To pc
  If ps(i).vs Then pRedrawOne i
 Next i
End If
'gripper
pRedrawOne -1
'over
UserControl_Paint
End Sub

Private Sub pRedrawOne(ByVal i As Long, Optional ByVal DrawBackground As Boolean)
Dim r As RECT, r1 As RECT, j As Long
Dim lst As Boolean
If i = -1 Then 'resize gripper
 If DrawBackground Then
  r.Right = bm.Width
  r.Left = r.Right - 16
  r.Bottom = bm.Height
  GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_V
 End If
 If grip Then
  TransparentBlt bm.hdc, bm.Width - 12, bm.Height - 12, 12, 12, bmd.hdc, 0, 0, 12, 12, vbGreen
 End If
 'over
 If DrawBackground Then bm.PaintPicture hdc, r.Left, 0, r.Right - r.Left, r.Bottom, r.Left, 0
ElseIf i > 0 Then 'panel
 With ps(i)
  If Not .vs Then Exit Sub
  'the last one and sizable??
  If .minw = 0 Then
   lst = True
   For j = i + 1 To pc
    If ps(j).vs Then
     lst = False
     Exit For
    End If
   Next j
  End If
  r.Left = .x
  r.Right = .x + .w + IIf(lst, 4, 6)
  r.Bottom = bm.Height
  If DrawBackground Then GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_V
  Select Case .st
  Case 1 'progress bar
   r1.Left = .x + 2
   r1.Top = 1
   r1.Bottom = bm.Height - 1
   If .mx > 0 Then
    r1.Right = r1.Left + (.w * .v) \ .mx
    GradientFillRect bm.hdc, r1.Left, r1.Top, r1.Right, r1.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
   End If
   r1.Right = r1.Left + .w
   j = CreateSolidBrush(d_Border)
   FrameRect bm.hdc, r1, j
   DeleteObject j
   cFnt.DrawTextXP bm.hdc, .Caption, .x + 2, 0, .w, bm.Height, DT_VCENTER Or DT_SINGLELINE Or DT_CENTER, IIf(.en, d_Text, d_TextDis), , True
  Case Else
   cFnt.DrawTextXP bm.hdc, .Caption, .x + 2, 0, .w, bm.Height, DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS, IIf(.en, d_Text, d_TextDis), , True
  End Select
  If Not lst Then 'seperator
   r1.Left = .x + .w + 4
   r1.Right = r1.Left + 1
   r1.Top = 2
   r1.Bottom = bm.Height - 3
   j = CreateSolidBrush(d_Sprt1)
   FillRect bm.hdc, r1, j
   DeleteObject j
   OffsetRect r1, 1, 1
   j = CreateSolidBrush(d_Sprt2)
   FillRect bm.hdc, r1, j
   DeleteObject j
  End If
  'over
  If DrawBackground Then bm.PaintPicture hdc, r.Left, 0, r.Right - r.Left, r.Bottom, r.Left, 0
 End With
End If
End Sub

Public Property Get ShowResizeGripper() As Boolean
ShowResizeGripper = grip
End Property

Public Property Let ShowResizeGripper(ByVal b As Boolean)
If grip Xor b Then
 grip = b
 If pRun Then pRedrawAll
End If
End Property

Public Property Get SimpleText() As String
SimpleText = simp
End Property

Public Property Let SimpleText(ByVal s As String)
simp = s
If pRun Then pRedrawAll
End Property

Public Property Get Font() As StdFont
Set Font = fnt
End Property

Public Property Set Font(obj As StdFont)
Set fnt = obj
cFnt.HighQuality = True
Set cFnt.LogFont = obj
If pRun Then pRedrawAll
End Property

Public Property Get PanelCaption(ByVal Index As Long) As String
PanelCaption = ps(Index).Caption
End Property

Public Property Let PanelCaption(ByVal Index As Long, ByVal s As String)
ps(Index).Caption = s
If ps(Index).vs Then pRedrawOne Index, True
End Property

Public Property Get PanelEnable(ByVal Index As Long) As Boolean
PanelEnable = ps(Index).en
End Property

Public Property Let PanelEnable(ByVal Index As Long, ByVal b As Boolean)
If ps(Index).en Xor b Then
 ps(Index).en = b
 If ps(Index).vs Then pRedrawOne Index, True
End If
End Property

Public Property Get PanelVisible(ByVal Index As Long) As Boolean
PanelVisible = ps(Index).vs
End Property

Public Property Let PanelVisible(ByVal Index As Long, ByVal b As Boolean)
If ps(Index).vs Xor b Then
 ps(Index).vs = b
 pRedrawAll
End If
End Property

Public Property Get PanelWidth(ByVal Index As Long) As Long
PanelWidth = ps(Index).minw
End Property

Public Property Let PanelWidth(ByVal Index As Long, ByVal n As Long)
If ps(Index).minw <> n Then
 ps(Index).minw = n
 If ps(Index).vs Then pRedrawAll
End If
End Property

Public Property Get PanelStyle(ByVal Index As Long) As enumSBStyle
PanelStyle = ps(Index).st
End Property

Public Property Let PanelStyle(ByVal Index As Long, ByVal n As enumSBStyle)
If ps(Index).st <> n Then
 ps(Index).st = n
 If ps(Index).vs Then pRedrawOne Index, True
End If
End Property

Public Property Get ProgressBarMax(ByVal Index As Long) As Long
ProgressBarMax = ps(Index).mx
End Property

Public Property Let ProgressBarMax(ByVal Index As Long, ByVal n As Long)
If ps(Index).st = 1 And ps(Index).mx <> n And n > 0 Then
 ps(Index).mx = n
 ps(Index).v = 0
 If ps(Index).vs Then pRedrawOne Index, True
End If
End Property

Public Property Get ProgressBarValue(ByVal Index As Long) As Long
ProgressBarValue = ps(Index).v
End Property

Public Property Let ProgressBarValue(ByVal Index As Long, ByVal n As Long)
If ps(Index).st = 1 And ps(Index).v <> n Then
 If n < 0 Then n = 0
 If n > ps(Index).mx Then n = ps(Index).mx
 ps(Index).v = n
 If ps(Index).vs Then pRedrawOne Index, True
End If
End Property

Public Property Get PanelCount() As Long
PanelCount = pc
End Property

Public Sub AddPanel(Optional ByVal Caption As String, Optional ByVal Index As Long, Optional ByVal Enabled As Boolean = True, Optional ByVal Visible As Boolean = True, Optional ByVal Width As Long, Optional ByVal Style As enumSBStyle)
Dim i As Long
pc = pc + 1
ReDim Preserve ps(1 To pc)
If Index = 0 Or Index >= pc Then
 i = pc
Else
 For i = pc - 1 To Index Step -1
  ps(i + 1) = ps(i)
 Next i
 i = Index
End If
With ps(i)
 .Caption = Caption
 .en = Enabled
 .vs = Visible
 .minw = Width
 .st = Style
 '////
 .v = 0
 .mx = 100
End With
If Visible Then pRedrawAll
End Sub

Public Sub RemovePanel(ByVal Index As Long)
Dim i As Long
If Index <= 0 Or Index > pc Then Exit Sub
If pc = 1 Then
 Erase ps
 pc = 0
Else
 If Index < pc Then
  For i = Index To pc - 1
   ps(i) = ps(i + 1)
  Next i
 End If
 pc = pc - 1
 ReDim Preserve ps(1 To pc)
End If
pRedrawAll
End Sub

VERSION 5.00
Begin VB.UserControl TabHeader2 
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
   FontTransparent =   0   'False
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   2880
      Top             =   1920
   End
   Begin VB.Image i1 
      Height          =   240
      Left            =   2160
      Picture         =   "TabHeader2.ctx":0000
      Top             =   1440
      Visible         =   0   'False
      Width           =   720
   End
End
Attribute VB_Name = "TabHeader2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long

Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long

Private bm As New cDIBSection
Private bmd As New cDIBSection, bmGray As New cDIBSection
Private cFnt As New CLogFont, bk As Long

Private Type typeTab
 Caption As String
 Tag As String
 nFlags As Long
 w As Long
End Type

Private bShowClose As Boolean

Private tabs() As typeTab, tabc As Long
Private tsld As Long, thl As Long, tleft As Long, pr As Boolean
Private thl2 As Boolean
Private t As Long

Private btn1 As Boolean, btn2 As Boolean, bhl As Long
'btn1 is visible!!!
Private btn3 As Boolean 'close button

Private thl_old As Long, thl2_old As Boolean, bhl_old As Long

Public Event TabClick(ByVal TabIndex As Long)
Public Event ContextMenu(ByVal TabIndex As Long)
Public Event TabClose(ByVal TabIndex As Long, ByRef Cancel As Boolean)

Public Property Get ShowCloseButtonOnTab() As Boolean
ShowCloseButtonOnTab = bShowClose
End Property

Public Property Let ShowCloseButtonOnTab(ByVal b As Boolean)
If bShowClose <> b Then
 bShowClose = b
 pRefresh
End If
End Property

Private Sub Timer1_Timer()
If t = 0 Or t > 5 Then
 If btn1 And tleft > 1 And bhl = 1 Then
  tleft = tleft - 1
  pRedraw
 End If
 If btn1 And btn2 And bhl = 2 Then
  tleft = tleft + 1
  pRedraw
 End If
End If
t = t + 1
End Sub

Private Sub UserControl_Initialize()
'
End Sub

Private Sub UserControl_InitProperties()
cFnt.HighQuality = True
Set cFnt.LOGFONT = UserControl.Font
bk = vbButtonFace
pInit
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If thl > 0 Or bhl > 0 Then
 '///
 thl_old = thl
 thl2_old = thl2
 bhl_old = bhl
 '///
 t = 0
 Timer1.Enabled = bhl > 0
 pr = True
 pRedraw
End If
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, j As Long, xx As Long
Dim b As Boolean
'tab highlight?
xx = 16
i = tleft
Do
 If i > tabc Or i <= 0 Then Exit Do
 xx = xx + tabs(i).w
 If xx > bm.Width - 16 Then Exit Do
 If x >= xx - tabs(i).w And x < xx And y >= 0 And y < bm.Height Then
  j = i
  If (tabs(i).nFlags And 3&) = 3& And bShowClose Then
   If x >= xx - 15 And x < xx - 2 Then
    b = y >= bm.Height \ 2 - 7 And y < bm.Height \ 2 + 6
   End If
  End If
  Exit Do
 End If
 i = i + 1
Loop
If j > 0 Then
 If Button Then
  If thl_old <> j Or thl2_old <> b Then j = 0
 End If
 If thl <> j Or thl2 <> b Then
  thl = j
  thl2 = b
  bhl = 0
  pRedraw
 End If
Else
 If x >= 0 And x < 16 And y >= 0 And y < 16 And btn1 And tleft > 1 Then
  j = 1
 ElseIf x >= bm.Width - 16 And x < bm.Width And y >= 0 And y < 16 And ((btn1 And btn2) Or btn3) Then
  If btn3 Then j = 3 Else j = 2
 ElseIf x >= bm.Width - 32 And x < bm.Width - 16 And y >= 0 And y < 16 And btn1 And btn2 And btn3 Then
  j = 2
 End If
 If Button Then
  If bhl_old <> j Then j = 0
 End If
 If thl > 0 Or bhl <> j Then
  thl = 0
  bhl = j
  pRedraw
 End If
End If
'capture?
If x >= 0 And x < bm.Width And y >= 0 And y <= bm.Height Then
 SetCapture hwnd
Else
 If Button = 0 Then ReleaseCapture
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim b As Boolean
If pr Then
 pr = False
 Timer1.Enabled = False
 If thl > 0 Then
  If tabs(thl).nFlags And 1& Then
   If Button = 1 Then
    If thl2 Then
     RaiseEvent TabClose(thl, b)
     If Not b Then RemoveTab thl
    Else
     tsld = thl
     RaiseEvent TabClick(tsld)
    End If
   ElseIf Button = 2 Then
    RaiseEvent ContextMenu(thl)
   End If
  End If
 ElseIf tsld > 0 And tsld <= tabc And bhl = 3 Then
  RaiseEvent TabClose(tsld, b)
  If Not b Then RemoveTab tsld
 End If
 thl = 0
 bhl = 0
 pRedraw
End If
End Sub

Private Sub UserControl_Paint()
bm.PaintPicture hdc
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
 cFnt.HighQuality = True
 Set cFnt.LOGFONT = .ReadProperty("Font", UserControl.Font)
 bk = .ReadProperty("BackColor", vbButtonFace)
 bShowClose = .ReadProperty("ShowCloseButtonOnTab", False)
End With
pInit
End Sub

Private Sub pInit()
tleft = 1
bmd.CreateFromPicture i1.Picture
GrayscaleBitmap bmd, bmGray, d_Icon_Grayscale, vbGreen
UserControl_Resize
End Sub

Private Sub UserControl_Resize()
bm.Create ScaleWidth, ScaleHeight
pRedraw
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
 .WriteProperty "Font", cFnt.LOGFONT, UserControl.Font
 .WriteProperty "BackColor", bk, vbButtonFace
 .WriteProperty "ShowCloseButtonOnTab", bShowClose, False
End With
End Sub

Private Sub pRefresh()
Dim i As Long
For i = 1 To tabc
 pCalcTabWidth tabs(i)
Next i
'///
pRedraw
End Sub

Private Sub pRedraw()
Dim i As Long, x As Long
Dim r As RECT, hbr As Long, hbr2 As Long
Dim r2 As RECT
'check btn1 and btn3
For i = 1 To tabc
 x = x + tabs(i).w
Next i
btn3 = False
If tsld > 0 And tsld <= tabc And Not bShowClose Then
 If (tabs(tsld).nFlags And 3&) = 3& Then
  btn3 = True
  x = x + 16
 End If
End If
btn1 = x > bm.Width - 32
If Not btn1 Then tleft = 1
hbr = CreateSolidBrush(d_Border)
hbr2 = CreateSolidBrush(TranslateColor(bk))
'background
r.Right = bm.Width
r.Bottom = bm.Height
GradientFillRect bm.hdc, 0, 0, r.Right, r.Bottom, d_Bar2, d_Bar1, GRADIENT_FILL_RECT_H
'tab
x = 16
i = tleft
btn2 = False
Do
 If i > tabc Or i <= 0 Then Exit Do
 r.Left = x
 x = x + tabs(i).w
 r.Right = x + 1
 If x > bm.Width - 16 Then
  x = x - tabs(i).w
  btn2 = True
  Exit Do
 End If
 r.Top = 0
 r.Bottom = bm.Height
 'bg
 If thl = i And (tabs(i).nFlags And 1&) <> 0 And Not thl2 Then
  If pr Or tsld = i Then
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
  Else
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
  End If
 ElseIf tsld = i Then
  GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Checked1, d_Checked2, GRADIENT_FILL_RECT_V
 Else
  GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_V
 End If
 'text
 cFnt.DrawTextXP bm.hdc, tabs(i).Caption, r.Left + 4, r.Top, tabs(i).w, bm.Height, DT_VCENTER Or DT_SINGLELINE, d_TextDis And (tabs(i).nFlags And 1&) = 0, , True
 'close button
 If (tabs(i).nFlags And 2&) <> 0 And bShowClose Then
  r2.Left = r.Right - 16
  r2.Top = r.Bottom \ 2 - 7
  r2.Right = r2.Left + 13
  r2.Bottom = r2.Top + 13
  If thl = i And (tabs(i).nFlags And 1&) <> 0 And thl2 Then
   If pr Then
    GradientFillRect bm.hdc, r2.Left, r2.Top, r2.Right, r2.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
   Else
    GradientFillRect bm.hdc, r2.Left, r2.Top, r2.Right, r2.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
   End If
   FrameRect bm.hdc, r2, hbr
  End If
  TransparentBlt bm.hdc, r2.Left + 2, r2.Top + 2, 9, 9, bmd.hdc, 32, 0, 9, 9, vbGreen
 End If
 'border
 FrameRect bm.hdc, r, hbr
 If tsld = i Then
  r.Left = r.Left + 1
  r.Top = r.Bottom - 1
  FillRect bm.hdc, r, hbr2
 End If
 i = i + 1
Loop
'tab line
r.Left = 0
r.Right = 16
r.Bottom = bm.Height
r.Top = r.Bottom - 1
FillRect bm.hdc, r, hbr
r.Left = x
r.Right = bm.Width
FillRect bm.hdc, r, hbr
'button
If btn1 Then
 If tleft > 1 Then
  If bhl = 1 Then
   r.Left = 0
   r.Top = 0
   r.Right = 16
   r.Bottom = 16
   If pr Then
    GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
   Else
    GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
   End If
   FrameRect bm.hdc, r, hbr
  End If
  TransparentBlt bm.hdc, 0, 0, 16, 16, bmd.hdc, 0, 0, 16, 16, vbGreen
 Else
  TransparentBlt bm.hdc, 0, 0, 16, 16, bmGray.hdc, 0, 0, 16, 16, vbGreen
 End If
 x = bm.Width - 16
 If btn3 Then x = x - 16
 If btn2 Then
  If bhl = 2 Then
   r.Left = x
   r.Top = 0
   r.Right = x + 16
   r.Bottom = 16
   If pr Then
    GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
   Else
    GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
   End If
   FrameRect bm.hdc, r, hbr
  End If
  TransparentBlt bm.hdc, x, 0, 16, 16, bmd.hdc, 16, 0, 16, 16, vbGreen
 Else
  TransparentBlt bm.hdc, x, 0, 16, 16, bmGray.hdc, 16, 0, 16, 16, vbGreen
 End If
End If
'close button
If btn3 Then
 x = bm.Width - 16
 If bhl = 3 Then
  r.Left = x + 1
  r.Top = 1
  r.Right = x + 16
  r.Bottom = 16
  If pr Then
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
  Else
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
  End If
  FrameRect bm.hdc, r, hbr
 End If
 TransparentBlt bm.hdc, x + 4, 4, 9, 9, bmd.hdc, 32, 0, 9, 9, vbGreen
End If
'design time!!!
If Not Ambient.UserMode Then
 cFnt.DrawTextXP bm.hdc, Extender.Name, 0, 0, bm.Width, bm.Height, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
End If
'///
DeleteObject hbr
DeleteObject hbr2
'over
UserControl_Paint
End Sub

Public Property Get Font() As StdFont
Set Font = cFnt.LOGFONT
End Property

Public Property Set Font(ByVal obj As StdFont)
cFnt.HighQuality = True
Set cFnt.LOGFONT = obj
pRefresh
End Property

Public Function AddTab(ByVal Caption As String, Optional ByVal Index As Long, Optional ByVal Enabled As Boolean = True, Optional ByVal Tag As String, Optional ByVal Closable As Boolean = False) As Long
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
 .Caption = Caption
 .Tag = Tag
 .nFlags = (Enabled And 1&) Or (Closable And 2&)
End With
pCalcTabWidth tabs(i)
pRedraw
AddTab = i
End Function

Public Sub RemoveTab(ByVal Index As Long)
Dim i As Long
If Index <= 0 Or Index > tabc Then Exit Sub
If tabc = 1 Then
 Erase tabs
 tabc = 0
 tsld = 0
Else
 If Index < tabc Then
  For i = Index To tabc - 1
   tabs(i) = tabs(i + 1)
  Next i
 End If
 i = tabc
 tabc = tabc - 1
 ReDim Preserve tabs(1 To tabc)
 If tsld = Index Then
  If tsld > tabc Then tsld = tabc
  RaiseEvent TabClick(tsld)
 ElseIf tsld > Index Then
  tsld = tsld - 1
 End If
End If
pRedraw
End Sub

Public Property Get TabEnable(ByVal Index As Long) As Boolean
TabEnable = tabs(Index).nFlags And 1&
End Property

Public Property Let TabEnable(ByVal Index As Long, ByVal b As Boolean)
If (tabs(Index).nFlags Xor b) And 1& Then
 tabs(Index).nFlags = (tabs(Index).nFlags And Not 1&) Or (b And 1&)
 pRedraw
End If
End Property

Public Property Get TabCaption(ByVal Index As Long) As String
TabCaption = tabs(Index).Caption
End Property

Public Property Let TabCaption(ByVal Index As Long, ByVal s As String)
If tabs(Index).Caption <> s Then
 tabs(Index).Caption = s
 pCalcTabWidth tabs(Index)
 pRedraw
End If
End Property

Public Property Get TabTag(ByVal Index As Long) As String
TabTag = tabs(Index).Tag
End Property

Public Property Let TabTag(ByVal Index As Long, ByVal s As String)
tabs(Index).Tag = s
End Property

Private Sub pCalcTabWidth(d As typeTab)
Dim w As Long
cFnt.DrawTextXP bm.hdc, d.Caption, 0, 0, w, 0, DT_SINGLELINE Or DT_CALCRECT
If (d.nFlags And 2&) <> 0 And bShowClose Then w = w + 14
d.w = w + 6
End Sub

Public Property Get BackColor() As OLE_COLOR
BackColor = bk
End Property

Public Property Let BackColor(ByVal n As OLE_COLOR)
bk = n
pRedraw
End Property

Public Property Get SelectedTab() As Long
Attribute SelectedTab.VB_MemberFlags = "400"
SelectedTab = tsld
End Property

Public Property Let SelectedTab(ByVal n As Long)
If n > 0 And n <= tabc And n <> tsld Then
 tsld = n
 pRedraw
 RaiseEvent TabClick(tsld)
End If
End Property

Public Property Get TabCount() As Long
TabCount = tabc
End Property

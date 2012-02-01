VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IFakeToolbarDraw"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////
'This file is public domain.
'////////////////////////////////

'================================
' !!! internal-purpose class !!!
'================================

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long

'////////////////////////control data

Private bm As New cDIBSection

Private bmPic As cDIBSection, bmGray As cDIBSection, transClr As Long
Private ps As Long

#Const UseFakeMenu = 1

Private tmr As Timer, tmrEn As Boolean
Private objCallback As IFakeToolbarDraw

Private xx As Long, yy As Long, ww As Long, hh As Long
Private m_hWnd As Long ', m_hDC As Long

Private FS As Long

Private btnHl As Long, btnPressed As Boolean, btnHlOld As Long  '??
'&H80000001=chevron menu
'&H80000002=gripper?

Private cFnt As New CLogFont, cFntб As New CLogFont

'chevron color
'Private Const clr_1_00 As Long = &HFCCEAE
'Private Const clr_1_01 As Long = &HFAB17F
'Private Const clr_1_11 As Long = clr_1_00
'Private Const clr_2_00 As Long = &HB96D3E
'Private Const clr_2_01 As Long = d_Chevron2
'Private Const clr_2_11 As Long = &HB96D3E
Private Const clr_1_00 As Long = (((d_Bar1 And &HFF&) + (d_Chevron1 And &HFF&)) \ 2&) Or _
((((d_Bar1 And &HFF00&) + (d_Chevron1 And &HFF00&)) \ 2&) And &HFF00&) Or _
((((d_Bar1 And &HFF0000) + (d_Chevron1 And &HFF0000)) \ 2&) And &HFF0000)
Private Const clr_2_00 As Long = (((d_Bar2 And &HFF&) + (d_Chevron2 And &HFF&)) \ 2&) Or _
((((d_Bar2 And &HFF00&) + (d_Chevron2 And &HFF00&)) \ 2&) And &HFF00&) Or _
((((d_Bar2 And &HFF0000) + (d_Chevron2 And &HFF0000)) \ 2&) And &HFF0000)
Private Const clr_2_01 As Long = (((d_Bar1 And &HFF&) + (d_Chevron2 And &HFF&)) \ 2&) Or _
((((d_Bar1 And &HFF00&) + (d_Chevron2 And &HFF00&)) \ 2&) And &HFF00&) Or _
((((d_Bar1 And &HFF0000) + (d_Chevron2 And &HFF0000)) \ 2&) And &HFF0000)
Private Const clr_3_00 As Long = (((d_Bar1 And &HFF&) + (d_Hl1 And &HFF&)) \ 2&) Or _
((((d_Bar1 And &HFF00&) + (d_Hl1 And &HFF00&)) \ 2&) And &HFF00&) Or _
((((d_Bar1 And &HFF0000) + (d_Hl1 And &HFF0000)) \ 2&) And &HFF0000)
Private Const clr_4_00 As Long = (((d_Bar2 And &HFF&) + (d_Hl2 And &HFF&)) \ 2&) Or _
((((d_Bar2 And &HFF00&) + (d_Hl2 And &HFF00&)) \ 2&) And &HFF00&) Or _
((((d_Bar2 And &HFF0000) + (d_Hl2 And &HFF0000)) \ 2&) And &HFF0000)
Private Const clr_4_01 As Long = (((d_Bar1 And &HFF&) + (d_Hl2 And &HFF&)) \ 2&) Or _
((((d_Bar1 And &HFF00&) + (d_Hl2 And &HFF00&)) \ 2&) And &HFF00&) Or _
((((d_Bar1 And &HFF0000) + (d_Hl2 And &HFF0000)) \ 2&) And &HFF0000)
Private Const clr_5_00 As Long = (((d_Bar1 And &HFF&) + (d_Pressed1 And &HFF&)) \ 2&) Or _
((((d_Bar1 And &HFF00&) + (d_Pressed1 And &HFF00&)) \ 2&) And &HFF00&) Or _
((((d_Bar1 And &HFF0000) + (d_Pressed1 And &HFF0000)) \ 2&) And &HFF0000)
Private Const clr_6_00 As Long = (((d_Bar2 And &HFF&) + (d_Pressed2 And &HFF&)) \ 2&) Or _
((((d_Bar2 And &HFF00&) + (d_Pressed2 And &HFF00&)) \ 2&) And &HFF00&) Or _
((((d_Bar2 And &HFF0000) + (d_Pressed2 And &HFF0000)) \ 2&) And &HFF0000)
Private Const clr_6_01 As Long = (((d_Bar1 And &HFF&) + (d_Pressed2 And &HFF&)) \ 2&) Or _
((((d_Bar1 And &HFF00&) + (d_Pressed2 And &HFF00&)) \ 2&) And &HFF00&) Or _
((((d_Bar1 And &HFF0000) + (d_Pressed2 And &HFF0000)) \ 2&) And &HFF0000)
'////////fake menu only

#If UseFakeMenu Then

Private btnDisplay As Long
' -1 : show all
'>=0 : show chevron

Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long

Private idxMenu As Long, idxMenuOld As Long '??
Private WithEvents objMenu As FakeMenu
Attribute objMenu.VB_VarHelpID = -1
Private bMainMenu As Boolean

#End If

'////////public callback function

Public Sub Paint()
'
End Sub

Public Sub Click(ByVal btnIndex As Long, ByVal btnKey As String)
'
End Sub

Public Sub SetToolTipText(ByVal s As String)
'
End Sub

Public Sub GetButtonSafeArrayData(ByRef lpSafeArray As Long, ByRef btnCount As Long)
'
End Sub

'////////

Friend Property Get PicSize() As Long
PicSize = ps
End Property

Friend Property Let PicSize(ByVal n As Long)
ps = n
End Property

Private Sub pDrawDropdown(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal clr As Long)
Dim i As Long, j As Long
Select Case FS
Case 0, 4
 For j = 0 To 2
  For i = -j To j
   SetPixelV hdc, x + i, y + 1 - j, clr
  Next i
 Next j
Case 1
 For j = 0 To 2
  For i = -j To j
   SetPixelV hdc, x + i, y - 1 + j, clr
  Next i
 Next j
Case 2
 For i = 0 To 2
  For j = -i To i
   SetPixelV hdc, x + 1 - i, y + j, clr
  Next j
 Next i
Case 3
 For i = 0 To 2
  For j = -i To i
   SetPixelV hdc, x - 1 + i, y + j, clr
  Next j
 Next i
End Select
End Sub

Friend Property Get TheBitmap() As cDIBSection
Set TheBitmap = bm
End Property

#If UseFakeMenu Then

Friend Property Get MainMenu() As Boolean
MainMenu = bMainMenu
End Property

Friend Property Let MainMenu(ByVal b As Boolean)
bMainMenu = b
End Property

Friend Property Get MenuObject() As FakeMenu
Set MenuObject = objMenu
End Property

Friend Property Set MenuObject(obj As FakeMenu)
Set objMenu = obj
End Property

#End If

Friend Sub SetBitmap(bm1 As cDIBSection, bm2 As cDIBSection)
Set bmPic = bm1
Set bmGray = bm2
End Sub

Friend Sub SetCallback(obj As IFakeToolbarDraw)
Set objCallback = obj
End Sub

Friend Property Get TheTimer() As Timer
Set TheTimer = tmr
End Property

Friend Property Set TheTimer(obj As Timer)
Set tmr = obj
End Property

Friend Property Get Orientation() As Long
Orientation = FS
End Property

Friend Property Let Orientation(ByVal n As Long)
If FS <> n Then
 FS = n
 pRotate
End If
End Property

Friend Property Get TransparentColor() As Long
TransparentColor = transClr
End Property

Friend Property Let TransparentColor(ByVal clr As Long)
transClr = clr
End Property

Friend Property Get Font() As StdFont
Set Font = cFnt.LOGFONT
End Property

Friend Property Set Font(obj As StdFont)
Dim fnt As IFont, fnt2 As New StdFont
cFnt.HighQuality = True
Set cFnt.LOGFONT = obj
Set fnt = obj
fnt.Clone fnt2
With fnt2
 .Name = "@" + .Name ':-3
End With
cFntб.HighQuality = True
cFntб.Rotation = -90
Set cFntб.LOGFONT = fnt2
End Property

Private Sub pRotate()
Select Case FS
Case 1
 cFnt.Rotation = 180
Case 2
 cFnt.Rotation = 90
Case 3
 cFnt.Rotation = -90
Case Else
 cFnt.Rotation = 0
End Select
End Sub

Private Sub pCalcBtnSize(btn As typeFakeButton)
Dim x As Long
With btn
 If .nFlags And 1& Then
  .Width = 0
 Else
  Select Case .nType
  Case 1  'separator
   .Width = 4
  Case 6
   .Width = 0
  Case Else
   .Width = 2
   If .PicLeft >= 0 Then .Width = .Width + ps + 2
   If .s <> "" And (.nFlags And 128&) = 0 Then
    cFnt.DrawTextXP bm.hdc, .s, 0, 0, x, , DT_SINGLELINE Or DT_CALCRECT
    .Width = .Width + x + 2
   End If
   If .nType = 5 Then
    .Width = .Width + 9
   ElseIf .nFlags And 4& Then
    .Width = .Width + 7
   End If
  End Select
 End If
End With
End Sub

Friend Sub Refresh(btns() As typeFakeButton, ByVal btnc As Long, Optional ByVal bPaint As Boolean = True)
're-calc button size
GetWidth btns, btnc
'over
Redraw btns, btnc, bPaint
End Sub

Friend Function GetWidth(btns() As typeFakeButton, ByVal btnc As Long) As Long
Dim i As Long, x As Long
#If UseFakeMenu Then
Dim w As Long
If FS < 2 Then w = ww Else w = hh
btnDisplay = -1
#End If
For i = 1 To btnc
 With btns(i)
  .Left = x
  pCalcBtnSize btns(i)
  x = x + .Width
  #If UseFakeMenu Then
  If x > w And btnDisplay < 0 Then
   If .Left > w - 11 And i > 1 Then
    btnDisplay = i - 2
   Else
    btnDisplay = i - 1
   End If
  End If
  #End If
 End With
Next i
GetWidth = x
End Function

Friend Sub Redraw(btns() As typeFakeButton, ByVal btnc As Long, Optional ByVal bPaint As Boolean = True)
Dim i As Long, j As Long
Dim r As RECT, r2 As RECT, hbr As Long
Dim n As Long
#If UseFakeMenu Then
Dim m As Long
m = btnDisplay
If m < 0 Or m > btnc Or objMenu Is Nothing Then m = btnc
#End If
If FS < 2 Then
 'draw background
 #If UseFakeMenu Then
 If bMainMenu Then
  GradientFillRect bm.hdc, 0, 0, ww, hh, d_Bar2, d_Bar1, GRADIENT_FILL_RECT_H
 Else
 #End If
  GradientFillRect bm.hdc, 0, 0, ww, hh, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_V
 #If UseFakeMenu Then
 End If
 #End If
 'draw button
 #If UseFakeMenu Then
 For i = 1 To m
 #Else
 For i = 1 To btnc
 #End If
  With btns(i)
   If (.nFlags And 1&) = 0 Then
    Select Case .nType
    Case 1 'separator
     r.Left = .Left + 1
     r.Top = 1
     r.Right = r.Left + 1
     r.Bottom = hh - 2
     hbr = CreateSolidBrush(d_Sprt1)
     FillRect bm.hdc, r, hbr
     DeleteObject hbr
     OffsetRect r, 1, 1
     hbr = CreateSolidBrush(d_Sprt2)
     FillRect bm.hdc, r, hbr
     DeleteObject hbr
    Case 6
    Case Else 'button
     r.Left = .Left
     r.Top = 0
     r.Right = .Left + .Width
     r.Bottom = hh
     'bg
     If (.nFlags And 2&) = 0 Then
      #If UseFakeMenu Then
      If idxMenu = i Then
       GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_V
      Else
      #End If
       If .Value <> 0 And btnHl <> i Then
        GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Checked1, d_Checked2, GRADIENT_FILL_RECT_V
       ElseIf btnHl = i Then
        If btnPressed Or .Value <> 0 Then
         GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
        Else
         GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
        End If
       End If
      #If UseFakeMenu Then
      End If
      #End If
     End If
     'picture
     If .PicLeft >= 0 And ps > 0 Then
      If .nFlags And 1024& Then
       r2.Left = .Left + 4
       r2.Top = 4
       r2.Right = r2.Left + ps - 4
       r2.Bottom = r2.Top + ps - 4
       hbr = CreateSolidBrush(.PicLeft)
       FillRect bm.hdc, r2, hbr
       DeleteObject hbr
       hbr = CreateSolidBrush(d_Border)
       FrameRect bm.hdc, r2, hbr
       DeleteObject hbr
      Else
       If Not bmPic Is Nothing And Not bmGray Is Nothing Then
        If .nFlags And 2& Then n = bmGray.hdc Else n = bmPic.hdc
        TransparentBlt bm.hdc, .Left + 2, (r.Bottom - ps) \ 2, ps, ps, n, .PicLeft, 0, ps, ps, transClr
       End If
      End If
      j = r.Left + ps + 4
     Else
      j = r.Left + 2
     End If
     'border
     #If UseFakeMenu Then
     If (btnHl = i And (.nFlags And 2&) = 0) Or .Value <> 0 Or idxMenu = i Then
     #Else
     If (btnHl = i And (.nFlags And 2&) = 0) Or .Value <> 0 Then
     #End If
      hbr = CreateSolidBrush(d_Border)
      FrameRect bm.hdc, r, hbr
      DeleteObject hbr
     End If
     'dropdown
     If .nType = 5 Or (.nFlags And 4&) Then
      If .nFlags And 2& Then n = d_TextDis Else n = d_Text
      pDrawDropdown bm.hdc, r.Right - 5, r.Bottom \ 2, n
     End If
     'text
     If .s <> "" And (.nFlags And 128&) = 0 Then
      If (.nFlags And 2050&) = 2& Then n = d_TextDis Else n = d_Text
      If FS = 1 Then
       j = r.Right - 2
       If .nType = 5 Then j = j - 10 Else If .nFlags And 4& Then j = j - 8
       cFnt.TextOutXP bm.hdc, j, (r.Bottom + ps) \ 2, .s, n, , True
      Else
       cFnt.DrawTextXP bm.hdc, .s, j, 0, .Width - 4, hh, DT_VCENTER Or DT_SINGLELINE, n, , True
      End If
     End If
     'split
     #If UseFakeMenu Then
     If idxMenu <> i Then
     #End If
      If btnHl = i And (.nFlags And 2&) = 0 And .nType = 5 Then
       hbr = CreateSolidBrush(d_Border)
       r.Left = r.Right - 9
       r.Right = r.Left + 1
       FrameRect bm.hdc, r, hbr
       DeleteObject hbr
      End If
     #If UseFakeMenu Then
     End If
     #End If
    End Select
   End If
  End With
 Next i
 #If UseFakeMenu Then
 'chevron
 If btnDisplay >= 0 And Not objMenu Is Nothing Then
  If bMainMenu Then r.Right = d_Bar1 Else r.Right = d_Bar2
  If btnHl = &H80000001 Or idxMenu = &H80000001 Then
   'check dropdown
   If idxMenu = &H80000001 Then
    i = d_Pressed1
    j = d_Pressed2
    r.Left = clr_5_00
    If bMainMenu Then r.Top = clr_6_01 Else r.Top = clr_6_00
   Else
    i = d_Hl1
    j = d_Hl2
    r.Left = clr_3_00
    If bMainMenu Then r.Top = clr_4_01 Else r.Top = clr_4_00
   End If
  Else
   i = d_Chevron1
   j = d_Chevron2
   r.Left = clr_1_00
   If bMainMenu Then r.Top = clr_2_01 Else r.Top = clr_2_00
  End If
  hbr = bm.hdc
  GradientFillRect hbr, ww - 9, 0, ww, hh, i, j, GRADIENT_FILL_RECT_V
  SetPixelV hbr, ww - 11, 0, r.Left
  SetPixelV hbr, ww - 10, 0, i
  SetPixelV hbr, ww - 10, 1, r.Left
  SetPixelV hbr, ww - 11, hh - 1, r.Top
  SetPixelV hbr, ww - 10, hh - 1, j
  SetPixelV hbr, ww - 10, hh - 2, r.Top
  SetPixelV hbr, ww - 2, 0, r.Left
  SetPixelV hbr, ww - 1, 0, d_Bar1
  SetPixelV hbr, ww - 1, 1, r.Left
  SetPixelV hbr, ww - 2, hh - 1, r.Top
  SetPixelV hbr, ww - 1, hh - 1, r.Right
  SetPixelV hbr, ww - 1, hh - 2, r.Top
  'dropdown icon
  SetPixelV hbr, ww - 8, 5, 0
  SetPixelV hbr, ww - 8, 6, 0
  SetPixelV hbr, ww - 8, 7, 0
  SetPixelV hbr, ww - 7, 6, 0
  SetPixelV hbr, ww - 7, 7, vbWhite
  SetPixelV hbr, ww - 7, 8, vbWhite
  SetPixelV hbr, ww - 6, 7, vbWhite
  bm.PaintPicture hbr, ww - 4, 5, 3, 4, ww - 8, 5
 End If
 #End If
Else 'vertical
 'draw background
 GradientFillRect bm.hdc, 0, 0, ww, hh, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_H
 'draw button
 #If UseFakeMenu Then
 For i = 1 To m
 #Else
 For i = 1 To btnc
 #End If
  With btns(i)
   If (.nFlags And 1&) = 0 Then
    Select Case .nType
    Case 1 'separator
     r.Left = 1
     r.Top = .Left + 1
     r.Right = ww - 1
     r.Bottom = r.Top + 1
     hbr = CreateSolidBrush(d_Sprt1)
     FillRect bm.hdc, r, hbr
     DeleteObject hbr
     OffsetRect r, 1, 1
     hbr = CreateSolidBrush(d_Sprt2)
     FillRect bm.hdc, r, hbr
     DeleteObject hbr
    Case 6
    Case Else 'button
     r.Left = 0
     r.Top = .Left
     r.Right = ww
     r.Bottom = .Left + .Width
     'bg
     If (.nFlags And 2&) = 0 Then
      #If UseFakeMenu Then
      If idxMenu = i Then
       GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_H
      Else
      #End If
       If .Value <> 0 And btnHl <> i Then
        GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Checked1, d_Checked2, GRADIENT_FILL_RECT_H
       ElseIf btnHl = i Then
        If btnPressed Or .Value <> 0 Then
         GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_H
        Else
         GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_H
        End If
       End If
      #If UseFakeMenu Then
      End If
      #End If
     End If
     'picture
     If .PicLeft >= 0 And ps > 0 Then
      If .nFlags And 1024& Then
       r2.Left = 4
       r2.Top = .Left + 4
       r2.Right = r2.Left + ps - 4
       r2.Bottom = r2.Top + ps - 4
       hbr = CreateSolidBrush(.PicLeft)
       FillRect bm.hdc, r2, hbr
       DeleteObject hbr
       hbr = CreateSolidBrush(d_Border)
       FrameRect bm.hdc, r2, hbr
       DeleteObject hbr
      Else
       If Not bmPic Is Nothing And Not bmGray Is Nothing Then
        If .nFlags And 2& Then n = bmGray.hdc Else n = bmPic.hdc
        TransparentBlt bm.hdc, (r.Right - ps) \ 2, .Left + 2, ps, ps, n, .PicLeft, 0, ps, ps, transClr
       End If
      End If
      j = r.Top + ps + 4
     Else
      j = r.Top + 2
     End If
     'border
     #If UseFakeMenu Then
     If (btnHl = i And (.nFlags And 2&) = 0) Or .Value <> 0 Or idxMenu = i Then
     #Else
     If (btnHl = i And (.nFlags And 2&) = 0) Or .Value <> 0 Then
     #End If
      hbr = CreateSolidBrush(d_Border)
      FrameRect bm.hdc, r, hbr
      DeleteObject hbr
     End If
     'dropdown
     If .nType = 5 Or (.nFlags And 4&) Then
      If .nFlags And 2& Then n = d_TextDis Else n = d_Text
      pDrawDropdown bm.hdc, r.Right \ 2, r.Bottom - 5, n
     End If
     'text
     If .s <> "" And (.nFlags And 128&) = 0 Then
      If (.nFlags And 2050&) = 2& Then n = d_TextDis Else n = d_Text
      Select Case FS
      Case 2
       j = r.Bottom - 2
       If .nType = 5 Then j = j - 10 Else If .nFlags And 4& Then j = j - 8
       cFnt.TextOutXP bm.hdc, (r.Right - ps) \ 2, j, .s, n, , True
      Case 3
       cFnt.TextOutXP bm.hdc, (r.Right + ps) \ 2, j, .s, n, , True
      Case 4
       '????????
       cFntб.TextOutXP bm.hdc, (r.Right + ps) \ 2, j, .s, n, , True
      End Select
     End If
     'split
     #If UseFakeMenu Then
     If idxMenu <> i Then
     #End If
      If btnHl = i And (.nFlags And 2&) = 0 And .nType = 5 Then
       hbr = CreateSolidBrush(d_Border)
       r.Top = r.Bottom - 9
       r.Bottom = r.Top + 1
       FrameRect bm.hdc, r, hbr
       DeleteObject hbr
      End If
     #If UseFakeMenu Then
     End If
     #End If
    End Select
   End If
  End With
 Next i
 #If UseFakeMenu Then
 'chevron
 If btnDisplay >= 0 And Not objMenu Is Nothing Then
  If bMainMenu Then r.Right = d_Bar1 Else r.Right = d_Bar2
  If btnHl = &H80000001 Or idxMenu = &H80000001 Then
   'check dropdown
   If idxMenu = &H80000001 Then
    i = d_Pressed1
    j = d_Pressed2
    r.Left = clr_5_00
    If bMainMenu Then r.Top = clr_6_01 Else r.Top = clr_6_00
   Else
    i = d_Hl1
    j = d_Hl2
    r.Left = clr_3_00
    If bMainMenu Then r.Top = clr_4_01 Else r.Top = clr_4_00
   End If
  Else
   i = d_Chevron1
   j = d_Chevron2
   r.Left = clr_1_00
   If bMainMenu Then r.Top = clr_2_01 Else r.Top = clr_2_00
  End If
  hbr = bm.hdc
  GradientFillRect hbr, 0, hh - 9, ww, hh, i, j, GRADIENT_FILL_RECT_H
  SetPixelV hbr, 0, hh - 11, r.Left
  SetPixelV hbr, 0, hh - 10, i
  SetPixelV hbr, 1, hh - 10, r.Left
  SetPixelV hbr, ww - 1, hh - 11, r.Top
  SetPixelV hbr, ww - 1, hh - 10, j
  SetPixelV hbr, ww - 2, hh - 10, r.Top
  SetPixelV hbr, 0, hh - 2, r.Left
  SetPixelV hbr, 0, hh - 1, d_Bar1
  SetPixelV hbr, 1, hh - 1, r.Left
  SetPixelV hbr, ww - 1, hh - 2, r.Top
  SetPixelV hbr, ww - 1, hh - 1, r.Right
  SetPixelV hbr, ww - 2, hh - 1, r.Top
  'dropdown icon
  SetPixelV hbr, 5, hh - 8, 0
  SetPixelV hbr, 6, hh - 8, 0
  SetPixelV hbr, 7, hh - 8, 0
  SetPixelV hbr, 6, hh - 7, 0
  SetPixelV hbr, 7, hh - 7, vbWhite
  SetPixelV hbr, 8, hh - 7, vbWhite
  SetPixelV hbr, 7, hh - 6, vbWhite
  bm.PaintPicture hbr, 5, hh - 4, 4, 3, 5, hh - 8
 End If
 #End If
End If
If bPaint Then
 If Not objCallback Is Nothing Then objCallback.Paint
End If
End Sub

Private Sub Class_Initialize()
transClr = vbGreen
End Sub

Friend Sub OnTimer(btns() As typeFakeButton, ByVal btnc As Long)
Dim p As POINTAPI
If Not tmrEn Then Exit Sub
#If UseFakeMenu Then
If idxMenu Then
 If Not objMenu Is Nothing Then
  If objMenu.MenuWindowCount <= 0 Or objMenu.UserData <> ObjPtr(Me) Then
   idxMenu = 0
   idxMenuOld = 0
   Redraw btns, btnc
  End If
 End If
End If
#End If
GetCursorPos p
ScreenToClient m_hWnd, p
If p.x < xx Or p.y < yy Or p.x >= xx + ww Or p.y >= yy + hh Then
 If btnHl <> 0 Then
  If Not objCallback Is Nothing Then objCallback.SetToolTipText ""
  btnHl = 0
  Redraw btns, btnc
 End If
 #If UseFakeMenu Then
 If idxMenu = 0 Then
 #End If
  tmrEn = False
  If Not tmr Is Nothing Then tmr.Enabled = False
 #If UseFakeMenu Then
 End If
 #End If
End If
End Sub

Friend Property Get TimerEnabled() As Boolean
TimerEnabled = tmrEn
End Property

Friend Property Let TimerEnabled(ByVal b As Boolean)
tmrEn = b
If Not tmr Is Nothing Then tmr.Enabled = b
End Property

Friend Sub OnClick(btns() As typeFakeButton, ByVal btnc As Long)
#If UseFakeMenu Then
If idxMenu Then Exit Sub
If idxMenuOld = btnHl Then
 idxMenuOld = 0
 Exit Sub
End If
idxMenuOld = 0
#End If
If btnHl > 0 And btnHl <= btnc Then
 'With btns(btnHl) 'fix a bug
  If (btns(btnHl).nFlags And 3&) = 0 Then
   Select Case btns(btnHl).nType
   Case 1, 6
   Case Else
    If Not objCallback Is Nothing Then objCallback.Click btnHl, btns(btnHl).sKey
   End Select
  End If
 'End With
End If
End Sub

Friend Sub OnDblClick(btns() As typeFakeButton, ByVal btnc As Long)
#If UseFakeMenu Then '???
If idxMenu Then
 idxMenuOld = idxMenu
 idxMenu = 0
 If Not objMenu Is Nothing Then
  objMenu.UnpopupMenu
 End If
End If
#End If
OnClick btns, btnc
End Sub

Friend Sub OnMouseMove(btns() As typeFakeButton, ByVal btnc As Long, ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
Dim i As Long
Dim b As Boolean
#If UseFakeMenu Then
Dim m As Long
m = btnDisplay
If m < 0 Or m > btnc Or objMenu Is Nothing Then m = btnc
#End If
x = x - xx
y = y - yy
#If UseFakeMenu Then
'chevron hit test
If btnDisplay >= 0 And Not objMenu Is Nothing Then
 If FS < 2 Then
  If x >= ww - 11 And x < ww Then b = True
 Else
  If y >= hh - 11 And y < hh Then b = True
 End If
 If b Then
  i = &H80000001
  'popup chevron menu?
  If idxMenu <> 0 And idxMenu <> i Then ShowChevron btns, btnc
  If i <> btnHl Then
   If Not objCallback Is Nothing Then objCallback.SetToolTipText "More buttons"
   btnHl = i
   Redraw btns, btnc
  End If
  tmrEn = True
  If Not tmr Is Nothing Then tmr.Enabled = True
  Exit Sub
 End If
End If
'hit test
For i = 1 To m
#Else
For i = 1 To btnc
#End If
 If Button = 0 Or i = btnHlOld Then '????????
  If FS < 2 Then
   b = x >= btns(i).Left And x < btns(i).Left + btns(i).Width And y >= 0 And y < hh
  Else
   b = y >= btns(i).Left And y < btns(i).Left + btns(i).Width And x >= 0 And x < ww
  End If
 End If
 If b Then
  'check sub-menu?
  #If UseFakeMenu Then
  If idxMenu <> 0 And idxMenu <> i Then
   With btns(i)
    b = (.nFlags And 3&) = 0 And .sSubMenu <> "" And .nType <> 1 And .nType <> 6
    If b Then b = objMenu.HasMenu(.sSubMenu)
   End With
   If b Then PopupMenu btns(i), i
  End If
  #End If
  If i <> btnHl Then
   If Not objCallback Is Nothing Then objCallback.SetToolTipText btns(i).s2
   btnHl = i
   Redraw btns, btnc
  End If
  tmrEn = True
  If Not tmr Is Nothing Then tmr.Enabled = True
  Exit Sub
 End If
Next i
If btnHl <> 0 Then
 If Not objCallback Is Nothing Then objCallback.SetToolTipText ""
 btnHl = 0
 Redraw btns, btnc
End If
#If UseFakeMenu Then
If idxMenu = 0 Then
#End If
 tmrEn = False
 If Not tmr Is Nothing Then tmr.Enabled = False
#If UseFakeMenu Then
End If
#End If
End Sub

Friend Sub OnMouseDown(btns() As typeFakeButton, ByVal btnc As Long, ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
#If UseFakeMenu Then
Dim b As Boolean
x = x - xx
y = y - yy
'check sub-menu?
If idxMenu Then
 idxMenuOld = idxMenu
 idxMenu = 0
 If Not objMenu Is Nothing Then
  objMenu.UnpopupMenu
 End If
 Redraw btns, btnc
Else
 idxMenuOld = 0
 If Not objMenu Is Nothing Then
  'chevron menu?
  If btnHl = &H80000001 And btnDisplay >= 0 Then
   ShowChevron btns, btnc
   btnPressed = False
   Redraw btns, btnc
   Exit Sub
  End If
  If btnHl > 0 Then
   With btns(btnHl)
    If (.nFlags And 3&) = 0 And .sSubMenu <> "" And .nType <> 1 And .nType <> 6 Then
     If FS < 2 Then
      If .nType = 5 Then b = x >= .Left + .Width - 9 Else b = x >= .Left
      b = b And x < .Left + .Width And y >= 0 And y < hh
     Else
      If .nType = 5 Then b = y >= .Left + .Width - 9 Else b = y >= .Left
      b = b And y < .Left + .Width And x >= 0 And x < ww
     End If
     If b Then b = objMenu.HasMenu(.sSubMenu)
    End If
   End With
   'popup menu?
   If b Then
    PopupMenu btns(btnHl), btnHl
    'TODO:other
    btnPressed = False
    Redraw btns, btnc
    Exit Sub
   End If
  End If
 End If
End If
#End If
btnHlOld = btnHl
btnPressed = True
If btnHl > 0 Then Redraw btns, btnc
End Sub

Friend Sub OnMouseUp(btns() As typeFakeButton, ByVal btnc As Long, ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
x = x - xx
y = y - yy
#If UseFakeMenu Then
If idxMenu Then Exit Sub
#End If
btnHlOld = 0
btnPressed = False
pClick btns, btnc, btnHl
End Sub

Private Sub pClick(btns() As typeFakeButton, ByVal btnc As Long, ByVal btnHl As Long)
Dim i As Long, j As Long
If btnHl > 0 And btnHl <= btnc Then
 With btns(btnHl)
  If (.nFlags And 3&) = 0 Then
   Select Case .nType
   Case 2 'check
    .Value = (.Value = 0) And 1&
   Case 3 'option
    j = .GroupIndex
    For i = 1 To btnc
     With btns(i)
      If .nType = 3 And .GroupIndex = j Then
       .Value = (i = btnHl) And 1&
      End If
     End With
    Next i
   Case 4 'optnull
    If .Value Then
     .Value = 0
    Else
     j = .GroupIndex
     For i = 1 To btnc
      With btns(i)
       If .nType = 4 And .GroupIndex = j Then
        .Value = (i = btnHl) And 1&
       End If
      End With
     Next i
    End If
   End Select
  End If
 End With
 Redraw btns, btnc
End If
End Sub

Friend Sub Resize(ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByVal hwnd As Long, btns() As typeFakeButton, ByVal btnc As Long, Optional ByVal bPaint As Boolean = True)
m_hWnd = hwnd
xx = Left
yy = Top
If ww <> Width Or hh <> Height Then
 ww = Width
 hh = Height
 bm.Create Width, Height
End If
#If UseFakeMenu Then
Refresh btns, btnc, bPaint
#Else
Redraw btns, btnc, bPaint
#End If
End Sub

Private Sub Class_Terminate()
#If UseFakeMenu Then
Set objMenu = Nothing
#End If
Set objCallback = Nothing
Set bmPic = Nothing
Set bmGray = Nothing
Set tmr = Nothing
End Sub

#If UseFakeMenu Then
Friend Sub PopupMenu(d As typeFakeButton, ByVal idxButton As Long)
Dim r As RECT, b As Boolean
If objMenu Is Nothing Then Exit Sub
'get pos
GetWindowRect m_hWnd, r
r.Left = r.Left + xx
r.Top = r.Top + yy
'popup from chevron?
If idxButton > btnDisplay And btnDisplay >= 0 Then
 If FS < 2 Then
  r.Right = 4
  r.Bottom = hh
 Else
  r.Right = ww
  r.Bottom = 4
 End If
 b = True
Else
 With d
  If FS < 2 Then
   r.Left = r.Left + .Left
   r.Right = .Width
   r.Bottom = hh
  Else
   r.Top = r.Top + .Left
   r.Right = ww
   r.Bottom = .Width
  End If
 End With
 b = False
End If
'popup
objMenu.PopupMenuEx d.sSubMenu, r.Left, r.Top, r.Right, r.Bottom, , , FS, b, ObjPtr(Me)
idxMenu = idxButton
End Sub
#End If

#If UseFakeMenu Then
Private Sub objMenu_Click(ByVal idxMenu As Long, ByVal Key As String, ByVal idxButton As Long, ByVal ButtonKey As String, Value As Long)
Dim btns() As typeFakeButton, btnc As Long
Dim i As Long
If Key = "____FakeTBChevron____" + CStr(ObjPtr(Me)) Then
 'get safearray :-3
 If Not objCallback Is Nothing Then
  objCallback.GetButtonSafeArrayData i, btnc
  If i = 0 Or btnc = 0 Then
   btnc = 0
  Else
   CopyMemory ByVal VarPtrArray(btns), ByVal i, 4&
  End If
 End If
 'check
 i = idxButton + btnDisplay
 If i > 0 And i <= btnc Then
  pClick btns, btnc, i
  If Not objCallback Is Nothing Then objCallback.Click i, btns(i).sKey
 End If
End If
CopyMemory ByVal VarPtrArray(btns), 0&, 4&
End Sub
#End If

#If UseFakeMenu Then
Friend Sub ShowChevron(btns() As typeFakeButton, ByVal btnc As Long)
Dim sKey As String
Dim i As String, j As Long, k As Long
Dim r As RECT
If objMenu Is Nothing Then Exit Sub
sKey = "____FakeTBChevron____" + CStr(ObjPtr(Me))
'///NO CHEVRON DIRTY
i = objMenu.FindMenu(sKey)
If i = 0 Then
 i = objMenu.AddMenu(sKey, , 2)
Else
 objMenu.DestroyMenuButtons i
End If
'add button
For j = btnDisplay + 1 To btnc
 With btns(j)
  objMenu.AddButtonByIndex i, , , .s, .s2, .nType, .nFlags Or 48&, .GroupIndex, .PicLeft, .sTab, .sDesc, .sSubMenu, .Value
 End With
Next j
'///over
'get pos
GetWindowRect m_hWnd, r
r.Left = r.Left + xx
r.Top = r.Top + yy
If FS < 2 Then
 r.Left = r.Left + ww - 11
 r.Right = 11
 r.Bottom = hh
Else
 r.Top = r.Top + hh - 11
 r.Right = ww
 r.Bottom = 11
End If
'popup
objMenu.PopupMenuEx sKey, r.Left, r.Top, r.Right, r.Bottom, , , FS, True, ObjPtr(Me)
idxMenu = &H80000001
End Sub
#End If

#If UseFakeMenu Then
Private Sub objMenu_DrawItem(ByVal idxMenu As Long, ByVal Key As String, ByVal idxButton As Long, ByVal ButtonKey As String, ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal nType As enumFakeButtonOwnerDrawType, nFlags As enumFakeButtonOwnerDrawFlags)
Dim PicLeft As Long, n As Long
If Key = "____FakeTBChevron____" + CStr(ObjPtr(Me)) Then
 If nType = fbtoBefore Then
  nFlags = &HFFFF&
  If ps > 0 Then
   If Not bmPic Is Nothing And Not bmGray Is Nothing Then
    nFlags = nFlags Or fbtoCheckIconFromPicLeft
   End If
  End If
 ElseIf nType = fbtoAfter Then
  If objMenu.ButtonFlags(idxMenu, idxButton) And 1024& Then 'color :-3
   nFlags = &HFFFF&
  Else
   PicLeft = objMenu.ButtonPicLeft(idxMenu, idxButton)
   If PicLeft >= 0 And ps > 0 Then
    If Not bmPic Is Nothing And Not bmGray Is Nothing Then
     If objMenu.ButtonFlags(idxMenu, idxButton) And 2& Then n = bmGray.hdc Else n = bmPic.hdc
     TransparentBlt hdc, Left + 2, Top + 2, ps, ps, n, PicLeft, 0, ps, ps, transClr
    End If
   End If
   nFlags = &HFFFF& And Not fbtoDoDefaultIcon
  End If
 End If
End If
End Sub
#End If

VERSION 5.00
Begin VB.UserControl ctlNCPaint 
   BackColor       =   &H8000000C&
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
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "NCPaint - paint borders"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1935
   End
End
Attribute VB_Name = "ctlNCPaint"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////
'This file is public domain.
'////////////////////////////////

Private Declare Function GetDCEx Lib "user32.dll" (ByVal hwnd As Long, ByVal hrgnclip As Long, ByVal fdwOptions As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Const DCX_WINDOW As Long = &H1&
Private Const DCX_INTERSECTRGN As Long = &H80&
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
'Private Type NCCALCSIZE_PARAMS
'    r(1 To 3) As RECT
'    wp As Long
'End Type
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Const WM_NCCALCSIZE As Long = &H83
Private Const WM_NCPAINT As Long = &H85
Private Const WM_MOUSEWHEEL As Long = &H20A

Private hwdParent As Long

Implements iSubclass

Private cSub As New cSubclass

Public Enum enumNCCalcSizeMode
 enumNCCalcSizeDefault = 0
End Enum

Public Enum enumNCPaintMode
 enumNCPaintDefault = 0
 enumNCPaintSolid = 1
 enumNCPaintHorizontal = 2
 enumNCPaintVertical = 3
 enumNCPaintCustom = 99
End Enum

Private mode1 As Long, mode2 As Long, bdrWidth As Long
Private clr1 As Long, clr2 As Long

Public Event Paint(ByVal hdc As Long, ByVal w As Long, ByVal h As Long)
Public Event MouseWheel(ByVal Button As Long, ByVal Shift As Long, ByVal lAmount As Long)

Public Property Get CalcSizeMode() As enumNCCalcSizeMode
CalcSizeMode = mode1
End Property

Public Property Let CalcSizeMode(ByVal n As enumNCCalcSizeMode)
mode1 = n
End Property

Public Property Get BorderWidth() As Long
BorderWidth = bdrWidth
End Property

Public Property Let BorderWidth(ByVal n As Long)
bdrWidth = n
End Property

Public Property Get PaintMode() As enumNCPaintMode
PaintMode = mode2
End Property

Public Property Let PaintMode(ByVal n As enumNCPaintMode)
mode2 = n
End Property

Public Property Get Color1() As OLE_COLOR
Color1 = clr1
End Property

Public Property Get Color2() As OLE_COLOR
Color2 = clr2
End Property

Public Property Let Color1(ByVal clr As OLE_COLOR)
clr1 = clr
End Property

Public Property Let Color2(ByVal clr As OLE_COLOR)
clr2 = clr
End Property

Private Sub pSolid(ByVal hdc As Long, ByVal w As Long, ByVal h As Long)
Dim hbr As Long
Dim r As RECT
hbr = CreateSolidBrush(TranslateColor(clr1))
If bdrWidth > 0 Then
 'top
 r.Right = w
 r.Bottom = bdrWidth
 FillRect hdc, r, hbr
 'bottom
 r.Top = h - bdrWidth
 r.Bottom = h
 FillRect hdc, r, hbr
 'left
 r.Bottom = r.Top
 r.Top = bdrWidth
 r.Right = bdrWidth
 FillRect hdc, r, hbr
 'right
 r.Left = w - bdrWidth
 r.Right = w
 FillRect hdc, r, hbr
Else
 r.Right = w
 r.Bottom = h
 FillRect hdc, r, hbr
End If
DeleteObject hbr
End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
Dim hd As Long
Dim r As RECT
Select Case uMsg
Case WM_NCPAINT
 hd = GetDCEx(hwnd, wParam, DCX_WINDOW Or DCX_INTERSECTRGN)
 GetWindowRect hwnd, r
 r.Right = r.Right - r.Left
 r.Bottom = r.Bottom - r.Top
 Select Case mode2
 Case 1
  pSolid hd, r.Right, r.Bottom
 Case 2
  If clr1 = clr2 Then
   pSolid hd, r.Right, r.Bottom
  Else
   If bdrWidth > 0 Then
    'top
    GradientFillRect hd, 0, 0, r.Right, bdrWidth, clr1, clr2, GRADIENT_FILL_RECT_H
    'bottom
    GradientFillRect hd, 0, r.Bottom - bdrWidth, r.Right, r.Bottom, clr1, clr2, GRADIENT_FILL_RECT_H
    'left
    StretchBlt hd, 0, 0, bdrWidth, r.Bottom, hd, 0, 0, bdrWidth, bdrWidth, vbSrcCopy
    'right
    StretchBlt hd, r.Right - bdrWidth, 0, bdrWidth, r.Bottom, hd, r.Right - bdrWidth, 0, bdrWidth, bdrWidth, vbSrcCopy
   Else
    GradientFillRect hd, 0, 0, r.Right, r.Bottom, clr1, clr2, GRADIENT_FILL_RECT_H
   End If
  End If
 Case 3
  If clr1 = clr2 Then
   pSolid hd, r.Right, r.Bottom
  Else
   If bdrWidth > 0 Then
    'left
    GradientFillRect hd, 0, 0, bdrWidth, r.Bottom, clr1, clr2, GRADIENT_FILL_RECT_V
    'right
    GradientFillRect hd, r.Right - bdrWidth, 0, r.Right, r.Bottom, clr1, clr2, GRADIENT_FILL_RECT_V
    'top
    StretchBlt hd, 0, 0, r.Right, bdrWidth, hd, 0, 0, bdrWidth, bdrWidth, vbSrcCopy
    'bottom
    StretchBlt hd, 0, r.Bottom - bdrWidth, r.Right, bdrWidth, hd, 0, r.Bottom - bdrWidth, bdrWidth, bdrWidth, vbSrcCopy
   Else
    GradientFillRect hd, 0, 0, r.Right, r.Bottom, clr1, clr2, GRADIENT_FILL_RECT_V
   End If
  End If
 Case 99
  RaiseEvent Paint(hd, r.Right, r.Bottom)
 End Select
 ReleaseDC hwnd, hd
 lReturn = 0
Case WM_MOUSEWHEEL
 ' Low-word of wParam indicates whether virtual keys are down
 r.Left = (wParam And &H3&) Or (((wParam And &H10&) <> 0) And &H4&)
 r.Top = ((wParam And &HC&) \ &H4&) Or (((wParam And &H20&) <> 0) And &H4&)
 ' High order word is the distance the wheel has been rotated, in multiples of WHEEL_DELTA:
 If (wParam And &H8000000) Then
  ' Towards the user:
  hd = &H8000& - (wParam And &H7FFF0000) \ &H10000
 Else
  ' Away from the user:
  hd = -((wParam And &H7FFF0000) \ &H10000)
 End If
 hd = hd \ 120&
 If hd Then RaiseEvent MouseWheel(r.Left, r.Top, hd)
End Select
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)
Select Case uMsg
Case WM_NCCALCSIZE
 'TODO:
End Select
End Sub

Private Sub UserControl_Initialize()
'
End Sub

Private Sub UserControl_InitProperties()
clr1 = vbApplicationWorkspace
clr2 = vbApplicationWorkspace
pInit
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
 mode1 = .ReadProperty("CalcSizeMode", 0)
 mode2 = .ReadProperty("PaintMode", 0)
 clr1 = .ReadProperty("Color1", vbApplicationWorkspace)
 clr2 = .ReadProperty("Color2", vbApplicationWorkspace)
 bdrWidth = .ReadProperty("BorderWidth", 0)
End With
pInit
End Sub

Private Sub pInit()
If Ambient.UserMode Then
 hwdParent = ContainerHwnd
 If mode1 > 0 Then
  cSub.AddMsg WM_NCCALCSIZE, MSG_BEFORE
 End If
 If mode2 > 0 Then
  cSub.AddMsg WM_NCPAINT, MSG_AFTER
 End If
 cSub.AddMsg WM_MOUSEWHEEL, MSG_AFTER
 cSub.Subclass hwdParent, Me
End If
End Sub

Private Sub UserControl_Terminate()
If hwdParent <> 0 Then
 cSub.DelMsg -1, MSG_BEFORE
 cSub.DelMsg -1, MSG_AFTER
 cSub.UnSubclass
End If
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
 .WriteProperty "CalcSizeMode", mode1, 0
 .WriteProperty "PaintMode", mode2, 0
 .WriteProperty "Color1", clr1, vbApplicationWorkspace
 .WriteProperty "Color2", clr2, vbApplicationWorkspace
 .WriteProperty "BorderWidth", bdrWidth, 0
End With
End Sub

VERSION 5.00
Begin VB.Form frmColorPicker 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Color Picker"
   ClientHeight    =   5880
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   6.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   392
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox t1 
      Height          =   240
      Index           =   8
      Left            =   5520
      TabIndex        =   23
      Top             =   3720
      Width           =   945
   End
   Begin VB.TextBox t1 
      Height          =   240
      Index           =   7
      Left            =   5520
      TabIndex        =   21
      Top             =   3480
      Width           =   945
   End
   Begin MyTGL.LeftRight lr1 
      Height          =   240
      Index           =   0
      Left            =   6480
      Top             =   1560
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   423
   End
   Begin VB.TextBox t1 
      Height          =   240
      Index           =   6
      Left            =   5520
      TabIndex        =   18
      Top             =   3120
      Width           =   945
   End
   Begin VB.TextBox t1 
      Height          =   240
      Index           =   5
      Left            =   5520
      TabIndex        =   17
      Top             =   2880
      Width           =   945
   End
   Begin VB.TextBox t1 
      Height          =   240
      Index           =   4
      Left            =   5520
      TabIndex        =   16
      Top             =   2640
      Width           =   945
   End
   Begin VB.TextBox t1 
      Height          =   240
      Index           =   3
      Left            =   5520
      TabIndex        =   15
      Top             =   2400
      Width           =   945
   End
   Begin VB.TextBox t1 
      Height          =   240
      Index           =   2
      Left            =   5520
      TabIndex        =   14
      Top             =   2040
      Width           =   945
   End
   Begin VB.TextBox t1 
      Height          =   240
      Index           =   1
      Left            =   5520
      TabIndex        =   13
      Top             =   1800
      Width           =   945
   End
   Begin VB.TextBox t1 
      Height          =   240
      Index           =   0
      Left            =   5520
      TabIndex        =   12
      Top             =   1560
      Width           =   945
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   11
      Top             =   600
      Width           =   855
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   6240
      TabIndex        =   10
      Top             =   120
      Width           =   855
   End
   Begin VB.OptionButton opt1 
      Caption         =   "G:"
      Height          =   240
      Index           =   5
      Left            =   5040
      TabIndex        =   9
      Top             =   2880
      Width           =   495
   End
   Begin VB.OptionButton opt1 
      Caption         =   "R:"
      Height          =   240
      Index           =   4
      Left            =   5040
      TabIndex        =   8
      Top             =   2640
      Width           =   495
   End
   Begin VB.OptionButton opt1 
      Caption         =   "A:"
      Height          =   240
      Index           =   3
      Left            =   5040
      TabIndex        =   7
      Top             =   2400
      Width           =   495
   End
   Begin VB.OptionButton opt1 
      Caption         =   "B:"
      Height          =   240
      Index           =   2
      Left            =   5040
      TabIndex        =   6
      Top             =   2040
      Width           =   495
   End
   Begin VB.OptionButton opt1 
      Caption         =   "S:"
      Height          =   240
      Index           =   1
      Left            =   5040
      TabIndex        =   5
      Top             =   1800
      Width           =   495
   End
   Begin VB.OptionButton opt1 
      Caption         =   "H:"
      Height          =   240
      Index           =   0
      Left            =   5040
      TabIndex        =   4
      Top             =   1560
      Value           =   -1  'True
      Width           =   495
   End
   Begin VB.PictureBox p2 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   990
      Left            =   5040
      ScaleHeight     =   66
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   66
      TabIndex        =   2
      Top             =   240
      Width           =   990
   End
   Begin VB.PictureBox p1 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   3960
      Left            =   120
      MousePointer    =   2  'Cross
      ScaleHeight     =   264
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   320
      TabIndex        =   0
      Top             =   120
      Width           =   4800
   End
   Begin VB.OptionButton opt1 
      Caption         =   "B:"
      Height          =   240
      Index           =   6
      Left            =   5040
      TabIndex        =   19
      Top             =   3120
      Width           =   495
   End
   Begin MyTGL.LeftRight lr1 
      Height          =   240
      Index           =   1
      Left            =   6480
      Top             =   1800
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   423
   End
   Begin MyTGL.LeftRight lr1 
      Height          =   240
      Index           =   2
      Left            =   6480
      Top             =   2040
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   423
   End
   Begin MyTGL.LeftRight lr1 
      Height          =   240
      Index           =   3
      Left            =   6480
      Top             =   2400
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   423
   End
   Begin MyTGL.LeftRight lr1 
      Height          =   240
      Index           =   4
      Left            =   6480
      Top             =   2640
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   423
   End
   Begin MyTGL.LeftRight lr1 
      Height          =   240
      Index           =   5
      Left            =   6480
      Top             =   2880
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   423
   End
   Begin MyTGL.LeftRight lr1 
      Height          =   240
      Index           =   6
      Left            =   6480
      Top             =   3120
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   423
   End
   Begin VB.Shape shp2 
      Height          =   1485
      Index           =   5
      Left            =   3330
      Top             =   4290
      Width           =   2325
   End
   Begin VB.Shape shp2 
      Height          =   405
      Index           =   4
      Left            =   1170
      Top             =   5370
      Width           =   2085
   End
   Begin VB.Shape shp2 
      Height          =   285
      Index           =   3
      Left            =   90
      Top             =   5370
      Width           =   1005
   End
   Begin VB.Shape shp2 
      Height          =   645
      Index           =   2
      Left            =   2250
      Top             =   4290
      Width           =   1005
   End
   Begin VB.Shape shp2 
      Height          =   765
      Index           =   1
      Left            =   1170
      Top             =   4290
      Width           =   1005
   End
   Begin VB.Shape shp2 
      Height          =   765
      Index           =   0
      Left            =   90
      Top             =   4290
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "System Color"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   9
      Left            =   1170
      TabIndex        =   29
      Top             =   5175
      Width           =   2085
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "QB Color"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   8
      Left            =   90
      TabIndex        =   28
      Top             =   5175
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Web Color"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   7
      Left            =   3330
      TabIndex        =   27
      Top             =   4095
      Width           =   2325
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Word Color"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   6
      Left            =   2250
      TabIndex        =   26
      Top             =   4095
      Width           =   1005
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "Basic Color"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   5
      Left            =   1170
      TabIndex        =   25
      Top             =   4095
      Width           =   1005
   End
   Begin VB.Shape shp1 
      BorderColor     =   &H00000000&
      FillColor       =   &H00FFFFFF&
      FillStyle       =   0  'Solid
      Height          =   105
      Index           =   0
      Left            =   120
      Top             =   4320
      Width           =   105
   End
   Begin VB.Image i0 
      Height          =   300
      Left            =   6240
      Picture         =   "frmColorPicker.frx":0000
      Top             =   1080
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackColor       =   &H8000000C&
      Caption         =   "VB Color"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   180
      Index           =   4
      Left            =   90
      TabIndex        =   24
      Top             =   4095
      Width           =   1005
   End
   Begin VB.Label Label1 
      Caption         =   "VB"
      Height          =   255
      Index           =   3
      Left            =   5040
      TabIndex        =   22
      Top             =   3720
      Width           =   375
   End
   Begin VB.Label Label1 
      Caption         =   "#"
      Height          =   255
      Index           =   2
      Left            =   5040
      TabIndex        =   20
      Top             =   3480
      Width           =   375
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Current"
      Height          =   255
      Index           =   1
      Left            =   5040
      TabIndex        =   3
      Top             =   1260
      Width           =   975
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "New"
      Height          =   255
      Index           =   0
      Left            =   5040
      TabIndex        =   1
      Top             =   60
      Width           =   975
   End
End
Attribute VB_Name = "frmColorPicker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////
'Fake Photoshop color picker by acme_pjz
'
'This file is public domain.
'////////////////////////////////

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long

Private Declare Function CreateRectRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private bm0 As New cDIBSection
Private bm As New cDIBSection
Private bm2 As New cDIBSection

Private m_rOld As Single, m_gOld As Single, m_bOld As Single, m_aOld As Single
Private m_r As Single, m_g As Single, m_b As Single, m_a As Single
Private m_HSB_H As Single, m_HSB_S As Single, m_HSB_B As Single

'properties
Private bChanged As Boolean, bInteger As Boolean, bClamp As Boolean

Private nSelType As Long
Private idxHl As Long
'1=left
'2=right

Private d() As Long

Private bChanging As Boolean

Public Property Get ClampBorder() As Boolean
ClampBorder = bClamp
End Property

Public Property Let ClampBorder(ByVal b As Boolean)
bClamp = b
End Property

Public Property Get UseInteger() As Boolean
UseInteger = bInteger
End Property

Public Property Let UseInteger(ByVal b As Boolean)
bInteger = b
End Property

Private Sub cmdCancel_Click()
Unload Me
End Sub

Public Function GetColorData() As Long
On Error Resume Next
Dim rgbRed As Long, rgbGreen As Long, rgbBlue As Long, rgbReserved As Long
rgbRed = m_r
rgbGreen = m_g
rgbBlue = m_b
rgbReserved = m_a
If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
If rgbReserved < 0 Then rgbReserved = 0 Else If rgbReserved > 255 Then rgbReserved = 255
If rgbReserved > 127 Then rgbReserved = rgbReserved - 256
GetColorData = rgbRed Or _
(rgbGreen * &H100&) Or _
(rgbBlue * &H10000) Or _
(rgbReserved * &H1000000)
End Function

Public Sub SetColorData(ByVal clr As Long)
m_r = clr And &HFF&
m_g = (clr And &HFF00&) \ &H100&
m_b = (clr And &HFF0000) \ &H10000
m_a = (clr And &HFF000000) \ &H1000000
If m_a < 0 Then m_a = m_a + 256
m_rOld = m_r
m_gOld = m_g
m_bOld = m_b
m_aOld = m_a
bChanged = False
pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
End Sub

Public Sub GetColor(ByRef rgbRed As Single, ByRef rgbGreen As Single, ByRef rgbBlue As Single, ByRef rgbReserved As Single)
rgbRed = m_r
rgbGreen = m_g
rgbBlue = m_b
rgbReserved = m_a
End Sub

Public Sub SetColor(ByVal rgbRed As Single, ByVal rgbGreen As Single, ByVal rgbBlue As Single, ByVal rgbReserved As Single)
m_r = rgbRed
m_g = rgbGreen
m_b = rgbBlue
m_a = rgbReserved
m_rOld = rgbRed
m_gOld = rgbGreen
m_bOld = rgbBlue
m_aOld = rgbReserved
bChanged = False
pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
End Sub

Public Property Get Changed() As Boolean
Changed = bChanged
End Property

Private Sub cmdOK_Click()
bChanged = True
Unload Me
End Sub

Private Sub Form_Initialize()
m_a = 255
m_aOld = 255
End Sub

Private Sub Form_Load()
Dim i As Long
bm0.CreateFromPicture i0.Picture
bm.Create p1.ScaleWidth, p1.ScaleHeight
bm2.Create p2.ScaleWidth, p2.ScaleHeight
pAddColors
pRedraw
bChanging = True
pChange -1
bChanging = False
For i = 0 To shp2.UBound
 shp2(i).BorderColor = d_CtrlBorder
Next i
End Sub

Private Sub pAddColor(ByVal clr As Long, ByVal i As Long, ByVal j As Long)
Dim k As Long
k = shp1.UBound + 1
Load shp1(k)
With shp1(k)
 .FillColor = clr
 .Move i * 8, j * 8 + 280, 7, 7
 .Visible = True
End With
End Sub

Private Sub pAddColors()
Dim i As Long, j As Long, k As Long
Dim v As Variant
'////////VB color
v = Array(&HE0E0E0, &HC0C0C0, &H808080, &H404040, 0&, _
&HC0C0FF, &H8080FF, &HFF&, &HC0&, &H80&, &H40&, _
&HC0E0FF, &H80C0FF, &H80FF&, &H40C0&, &H4080&, &H404080, _
&HC0FFFF, &H80FFFF, &HFFFF&, &HC0C0&, &H8080&, &H4040&, _
&HC0FFC0, &H80FF80, &HFF00&, &HC000&, &H8000&, &H4000&, _
&HFFFFC0, &HFFFF80, &HFFFF00, &HC0C000, &H808000, &H404000, _
&HFFC0C0, &HFF8080, &HFF0000, &HC00000, &H800000, &H400000, _
&HFFC0FF, &HFF80FF, &HFF00FF, &HC000C0, &H800080, &H400040)
j = 1
k = 2
For i = 0 To UBound(v)
 pAddColor v(i), j, k
 k = k + 1
 If k > 6 Then
  k = 1
  j = j + 1
 End If
Next i
'////////color dialog color
v = Array(&H8080FF, &HFF&, &H404080, &H80&, &H40&, &H0&, _
&H80FFFF, &HFFFF&, &H4080FF, &H80FF&, &H4080&, &H8080&, _
&H80FF80, &HFF80&, &HFF00&, &H8000&, &H4000&, &H408080, _
&H80FF00, &H40FF00, &H808000, &H408000, &H404000, &H808080, _
&HFFFF80, &HFFFF00, &H804000, &HFF0000, &H800000, &H808040, _
&HFF8000, &HC08000, &HFF8080, &HA00000, &H400000, &HC0C0C0, _
&HC080FF, &HC08080, &H400080, &H800080, &H400040, &H400040, _
&HFF80FF, &HFF00FF, &H8000FF, &HFF0080, &H800040, &HFFFFFF)
j = 10
k = 1
For i = 0 To UBound(v)
 pAddColor v(i), j, k
 k = k + 1
 If k > 6 Then
  k = 1
  j = j + 1
 End If
Next i
'////////word color
v = Array(&H0&, &H80&, &HFF&, &HFF00FF, &HCC99FF, _
&H3399&, &H66FF&, &H99FF&, &HCCFF&, &H99CCFF, _
&H3333&, &H8080&, &HCC99&, &HFFFF&, &H99FFFF, _
&H3300&, &H8000&, &H669933, &HFF00&, &HCCFFCC, _
&H663300, &H808000, &HCCCC33, &HFFFF00, &HFFFFCC, _
&H800000, &HFF0000, &HFF6633, &HFFCC00, &HFFCC99, _
&H993333, &H996666, &H800080, &H663399, &HFF99CC, _
&H333333, &H808080, &H999999, &HC0C0C0, &HFFFFFF)
j = 19
k = 1
For i = 0 To UBound(v)
 pAddColor v(i), j, k
 k = k + 1
 If k > 5 Then
  k = 1
  j = j + 1
 End If
Next i
'////////flash color
For i = 0 To 5
 pAddColor &H333333 * i, 28, i + 1
Next i
pAddColor vbRed, 28, 7
pAddColor vbGreen, 28, 8
pAddColor vbBlue, 28, 9
pAddColor vbYellow, 28, 10
pAddColor vbCyan, 28, 11
pAddColor vbMagenta, 28, 12
For i = 0 To 2
 For j = 0 To 5
  For k = 0 To 5
   pAddColor (i * &H33&) Or (j * &H3300&) Or (k * &H330000), 29 + i * 6 + j, k + 1
   pAddColor ((i + 3) * &H33&) Or (j * &H3300&) Or (k * &H330000), 29 + i * 6 + j, k + 7
  Next k
 Next j
Next i
'////////QBColor
For i = 0 To 7
 pAddColor QBColor(i), i + 1, 10
 pAddColor QBColor(i + 8), i + 1, 11
Next i
'////////system color
For i = 0 To 16
 pAddColor TranslateColor(&H80000000 + i), i + 10, 10
 pAddColor TranslateColor(&H80000011 + i), i + 10, 11
Next i
'////////XXX color
pAddColor d_Title1, 10, 12
pAddColor d_Title2, 11, 12
pAddColor d_Bar1, 12, 12
pAddColor d_Bar2, 13, 12
pAddColor d_Hl1, 14, 12
pAddColor d_Hl2, 15, 12
pAddColor d_Checked1, 16, 12
pAddColor d_Checked2, 17, 12
pAddColor d_Pressed1, 18, 12
pAddColor d_Pressed2, 19, 12
pAddColor d_Chevron1, 20, 12
pAddColor d_Chevron2, 21, 12
pAddColor d_BorderP, 22, 12
pAddColor d_SprtP, 23, 12
pAddColor d_HlP, 24, 12
pAddColor d_CheckedP, 25, 12
pAddColor d_PressedP, 26, 12
'finally
ReDim d(shp1.UBound)
For i = 0 To shp1.UBound
 With shp1(i)
  d(i) = (.Left \ 8&) + (.Top \ 8&) * &H10000
 End With
Next i
End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, j As Long
If x >= 8 And y >= 288 And Button = 1 Then
 j = (x \ 8&) + (y \ 8&) * &H10000
 For i = 0 To shp1.UBound
  If d(i) = j Then
   j = shp1(i).FillColor
   m_r = (j And &HFF&)
   m_g = (j And &HFF00&) \ &H100&
   m_b = (j And &HFF0000) \ &H10000
   pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
   pRedraw
   bChanging = True
   pChange -1
   bChanging = False
   Exit Sub
  End If
 Next i
End If
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
If Not bChanged Then
 m_r = m_rOld
 m_g = m_gOld
 m_b = m_bOld
 m_a = m_aOld
End If
End Sub

Private Sub lr1_Change(Index As Integer, ByVal iDelta As Long, ByVal Button As Long, ByVal Shift As Long, bCancel As Boolean)
On Error Resume Next
Dim f As Single
f = Val(t1(Index).Text)
f = f + iDelta
t1(Index).Text = CStr(f)
End Sub

Private Sub opt1_Click(Index As Integer)
nSelType = Index
pRedraw
End Sub

Private Sub p1_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If x < 268 Then idxHl = 1 Else idxHl = 2
If Button = 1 Then p1_MouseMove Button, Shift, x, y
End Sub

Private Sub p1_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
 If idxHl = 1 Then 'left
  pClickLeft x - 4, y - 4
  pRedraw
  bChanging = True
  pChange -1
  bChanging = False
 ElseIf idxHl = 2 Then 'right
  pClickRight y - 4
  pRedraw
  bChanging = True
  pChange -1
  bChanging = False
 End If
End If
End Sub

Private Sub p1_Paint()
bm.PaintPicture p1.hdc
End Sub

Private Sub p2_Paint()
bm2.PaintPicture p2.hdc
End Sub

Private Sub t1_Change(Index As Integer)
On Error Resume Next
Dim s As String
Dim f As Single, b As Long
If bChanging Then Exit Sub
bChanging = True
Select Case Index
Case 0 'H
 f = Val(t1(Index).Text)
 If f < 0 Then f = 0: b = -1 Else If f > 360 Then f = 360: b = -1
 m_HSB_H = f
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 1 'S
 f = Val(t1(Index).Text)
 If f < 0 Then f = 0: b = -1 Else If f > 100 Then f = 100: b = -1
 m_HSB_S = f
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 2 'B
 f = Val(t1(Index).Text)
 If f < 0 Then f = 0: b = -1 Else If f > 100 Then f = 100: b = -1
 m_HSB_B = f
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 3 'A
 f = Val(t1(Index).Text)
 If bClamp Then
  If f < 0 Then f = 0: b = -1 Else If f > 255 Then f = 255: b = -1
 End If
 m_a = f
Case 4 'R
 f = Val(t1(Index).Text)
 If bClamp Then
  If f < 0 Then f = 0: b = -1 Else If f > 255 Then f = 255: b = -1
 End If
 m_r = f
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
Case 5 'G
 f = Val(t1(Index).Text)
 If bClamp Then
  If f < 0 Then f = 0: b = -1 Else If f > 255 Then f = 255: b = -1
 End If
 m_g = f
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
Case 6 'B
 f = Val(t1(Index).Text)
 If bClamp Then
  If f < 0 Then f = 0: b = -1 Else If f > 255 Then f = 255: b = -1
 End If
 m_b = f
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
Case 7 'web
 s = Replace(t1(Index).Text, "#", "")
 s = Replace(s, " ", "")
 m_r = Val("&H" + Mid(s, 1, 2))
 m_g = Val("&H" + Mid(s, 3, 2))
 m_b = Val("&H" + Mid(s, 5, 2))
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
Case 8 'VB
 b = Val(t1(Index).Text)
 m_r = (b And &HFF&)
 m_g = (b And &HFF00&) \ &H100&
 m_b = (b And &HFF0000) \ &H10000
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
 b = 0
End Select
pChange Index Or b
pRedraw
bChanging = False
End Sub

Private Sub pChange(ByVal Index As Long)
On Error Resume Next
Dim s As String, i As Long
i = m_r
If i < 0 Then i = 0 Else If i > 255 Then i = 255
If i < 16 Then s = s + "0" + Hex(i) Else s = s + Hex(i)
i = m_g
If i < 0 Then i = 0 Else If i > 255 Then i = 255
If i < 16 Then s = s + "0" + Hex(i) Else s = s + Hex(i)
i = m_b
If i < 0 Then i = 0 Else If i > 255 Then i = 255
If i < 16 Then s = s + "0" + Hex(i) Else s = s + Hex(i)
If Index <> 0 Then If bInteger Then t1(0).Text = CStr(Round(m_HSB_H)) Else t1(0).Text = CStr(m_HSB_H)
If Index <> 1 Then If bInteger Then t1(1).Text = CStr(Round(m_HSB_S)) Else t1(1).Text = CStr(m_HSB_S)
If Index <> 2 Then If bInteger Then t1(2).Text = CStr(Round(m_HSB_B)) Else t1(2).Text = CStr(m_HSB_B)
If Index <> 3 Then If bInteger Then t1(3).Text = CStr(Round(m_a)) Else t1(3).Text = CStr(m_a)
If Index <> 4 Then If bInteger Then t1(4).Text = CStr(Round(m_r)) Else t1(4).Text = CStr(m_r)
If Index <> 5 Then If bInteger Then t1(5).Text = CStr(Round(m_g)) Else t1(5).Text = CStr(m_g)
If Index <> 6 Then If bInteger Then t1(6).Text = CStr(Round(m_b)) Else t1(6).Text = CStr(m_b)
If Index <> 7 Then t1(7).Text = s
If Index <> 8 Then t1(8).Text = "&H" + Mid(s, 5, 2) + Mid(s, 3, 2) + Mid(s, 1, 2)
End Sub

Private Sub t1_GotFocus(Index As Integer)
On Error Resume Next
With t1(Index)
 .SelStart = 0
 .SelLength = Len(.Text)
End With
End Sub

Private Sub pRedraw()
On Error Resume Next
Dim r As RECT, hbr As Long
'draw back
hbr = CreateSolidBrush(TranslateColor(vbButtonFace))
r.Right = bm.Width
r.Bottom = bm.Height
FillRect bm.hdc, r, hbr
DeleteObject hbr
'draw border
hbr = CreateSolidBrush(d_CtrlBorder)
r.Right = bm2.Width
r.Bottom = bm2.Height
FrameRect bm2.hdc, r, hbr
r.Left = 3
r.Top = 3
r.Right = 261
r.Bottom = 261
FrameRect bm.hdc, r, hbr
r.Left = 279
r.Right = 305
FrameRect bm.hdc, r, hbr
DeleteObject hbr
'draw
pDrawLeft
pDrawRight
pDrawColor
pDrawPredefined
'over
p1_Paint
p2_Paint
End Sub

Private Sub pDrawPredefined()
On Error Resume Next
Dim i As Long
Dim clr As Long, clr2 As Long
Dim rgbRed As Single, rgbGreen As Single, rgbBlue As Single
rgbRed = m_r
rgbGreen = m_g
rgbBlue = m_b
If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
clr = (rgbRed And &HFF&) Or _
((rgbGreen And &HFF&) * &H100&) Or _
((rgbBlue And &HFF&) * &H10000)
For i = 0 To shp1.UBound
 With shp1(i)
  clr2 = .FillColor
  .BorderColor = clr2 Xor ((clr = clr2) And &H808080)
 End With
Next i
End Sub

Private Sub pDrawLeft()
On Error Resume Next
Dim i As Long, clr As Long
Dim rgbRed As Single, rgbGreen As Single, rgbBlue As Single
Dim nSelX As Long, nSelY As Long
'draw left
Select Case nSelType
Case 0 'hsbH
 pHSB2RGB m_HSB_H, 100, 100, rgbRed, rgbGreen, rgbBlue
 For i = 0 To 255
  clr = ((rgbRed * i / 255) And &HFF&) Or _
  (((rgbGreen * i / 255) And &HFF&) * &H100&) Or _
  (((rgbBlue * i / 255) And &HFF&) * &H10000)
  GradientFillRect bm.hdc, 4, 259 - i, 260, 260 - i, i * &H10101, clr, GRADIENT_FILL_RECT_H
 Next i
 nSelX = m_HSB_S / 100 * 255
 nSelY = (1 - m_HSB_B / 100) * 255
Case 1 'hsbS
 For i = 0 To 255
  pHSB2RGB i / 255 * 360, m_HSB_S, 100, rgbRed, rgbGreen, rgbBlue
  clr = (rgbRed And &HFF&) Or _
  ((rgbGreen And &HFF&) * &H100&) Or _
  ((rgbBlue And &HFF&) * &H10000)
  GradientFillRect bm.hdc, 4 + i, 4, 5 + i, 260, clr, vbBlack, GRADIENT_FILL_RECT_V
 Next i
 nSelX = m_HSB_H / 360 * 255
 nSelY = (1 - m_HSB_B / 100) * 255
Case 2, 3 'hsbB
 For i = 0 To 255
  pHSB2RGB i / 255 * 360, 100, m_HSB_B, rgbRed, rgbGreen, rgbBlue
  clr = (rgbRed And &HFF&) Or _
  ((rgbGreen And &HFF&) * &H100&) Or _
  ((rgbBlue And &HFF&) * &H10000)
  GradientFillRect bm.hdc, 4 + i, 4, 5 + i, 260, clr, CLng(m_HSB_B / 100 * 255) * &H10101, GRADIENT_FILL_RECT_V
 Next i
 nSelX = m_HSB_H / 360 * 255
 nSelY = (1 - m_HSB_S / 100) * 255
Case 4 'r
 rgbRed = m_r
 If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
 clr = (rgbRed And &HFF&)
 For i = 0 To 255
  GradientFillRect bm.hdc, 4, 259 - i, 260, 260 - i, clr, clr Or &HFF0000, GRADIENT_FILL_RECT_H
  clr = clr + &H100&
 Next i
 nSelX = m_b
 nSelY = 255 - m_g
Case 5 'g
 rgbGreen = m_g
 If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
 clr = (rgbGreen And &HFF&) * &H100&
 For i = 0 To 255
  GradientFillRect bm.hdc, 4, 259 - i, 260, 260 - i, clr, clr Or &HFF0000, GRADIENT_FILL_RECT_H
  clr = clr + &H1&
 Next i
 nSelX = m_b
 nSelY = 255 - m_r
Case 6 'b
 rgbBlue = m_b
 If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
 clr = (rgbBlue And &HFF&) * &H10000
 For i = 0 To 255
  GradientFillRect bm.hdc, 4, 259 - i, 260, 260 - i, clr, clr Or &HFF&, GRADIENT_FILL_RECT_H
  clr = clr + &H100&
 Next i
 nSelX = m_r
 nSelY = 255 - m_g
End Select
'draw selected
clr = CreateRectRgn(4, 4, 260, 260)
SelectClipRgn bm.hdc, clr
bm0.PaintPicture bm.hdc, nSelX - 1, nSelY - 1, 11, 11, 0, 9, vbSrcInvert
SelectClipRgn bm.hdc, 0
End Sub

Private Sub pClickLeft(ByVal nSelX As Long, ByVal nSelY As Long)
If nSelX < 0 Then nSelX = 0 Else If nSelX > 255 Then nSelX = 255
If nSelY < 0 Then nSelY = 0 Else If nSelY > 255 Then nSelY = 255
Select Case nSelType
Case 0 'hsbH
 m_HSB_S = nSelX / 255 * 100
 m_HSB_B = (255 - nSelY) / 255 * 100
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 1 'hsbS
 m_HSB_H = nSelX / 255 * 360
 m_HSB_B = (255 - nSelY) / 255 * 100
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 2, 3 'hsbB
 m_HSB_H = nSelX / 255 * 360
 m_HSB_S = (255 - nSelY) / 255 * 100
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 4 'r
 m_b = nSelX
 m_g = 255 - nSelY
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
Case 5 'g
 m_b = nSelX
 m_r = 255 - nSelY
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
Case 6 'b
 m_r = nSelX
 m_g = 255 - nSelY
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
End Select
End Sub

Private Sub pClickRight(ByVal nSelY As Long)
If nSelY < 0 Then nSelY = 0 Else If nSelY > 255 Then nSelY = 255
Select Case nSelType
Case 0 'hsbH
 m_HSB_H = (255 - nSelY) / 255 * 360
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 1 'hsbS
 m_HSB_S = (255 - nSelY) / 255 * 100
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 2 'hsbB
 m_HSB_B = (255 - nSelY) / 255 * 100
 pHSB2RGB m_HSB_H, m_HSB_S, m_HSB_B, m_r, m_g, m_b
Case 3 'alpha!!!
 m_a = nSelY
Case 4 'r
 m_r = 255 - nSelY
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
Case 5 'g
 m_g = 255 - nSelY
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
Case 6 'b
 m_b = 255 - nSelY
 pRGB2HSB m_r, m_g, m_b, m_HSB_H, m_HSB_S, m_HSB_B
End Select
End Sub

Private Sub pDrawRight()
On Error Resume Next
Dim r As RECT
Dim i As Long, clr As Long, clr2 As Long
Dim rgbRed As Single, rgbGreen As Single, rgbBlue As Single
'draw right
Select Case nSelType
Case 0 'hsbH
 r.Left = 280
 r.Right = 304
 For i = 0 To 255
  pHSB2RGB i / 255 * 360, 100, 100, rgbRed, rgbGreen, rgbBlue
  clr = (rgbRed And &HFF&) Or _
  ((rgbGreen And &HFF&) * &H100&) Or _
  ((rgbBlue And &HFF&) * &H10000)
  clr = CreateSolidBrush(clr)
  r.Top = 259 - i
  r.Bottom = r.Top + 1
  FillRect bm.hdc, r, clr
  DeleteObject clr
 Next i
 i = (1 - m_HSB_H / 360) * 255
Case 1 'hsbS
 rgbRed = m_HSB_B
 If rgbRed < 20 Then rgbRed = 20
 i = CLng(rgbRed / 100 * 255) * &H10101
 pHSB2RGB m_HSB_H, 100, rgbRed, rgbRed, rgbGreen, rgbBlue
 clr = (rgbRed And &HFF&) Or _
 ((rgbGreen And &HFF&) * &H100&) Or _
 ((rgbBlue And &HFF&) * &H10000)
 GradientFillRect bm.hdc, 280, 4, 304, 260, clr, i, GRADIENT_FILL_RECT_V
 i = (1 - m_HSB_S / 100) * 255
Case 2 'hsbB
 pHSB2RGB m_HSB_H, m_HSB_S, 100, rgbRed, rgbGreen, rgbBlue
 clr = (rgbRed And &HFF&) Or _
 ((rgbGreen And &HFF&) * &H100&) Or _
 ((rgbBlue And &HFF&) * &H10000)
 GradientFillRect bm.hdc, 280, 4, 304, 260, clr, vbBlack, GRADIENT_FILL_RECT_V
 i = (1 - m_HSB_B / 100) * 255
Case 3 'alpha!!!
 '////////
 For i = 0 To 255
  r.Top = 259 - i
  r.Bottom = r.Top + 1
  'calc blend
  rgbRed = i + m_r
  rgbGreen = i + m_g
  rgbBlue = i + m_b
  If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
  If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
  If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
  clr = (rgbRed And &HFF&) Or _
  ((rgbGreen And &HFF&) * &H100&) Or _
  ((rgbBlue And &HFF&) * &H10000)
  clr = CreateSolidBrush(clr)
  rgbBlue = i / 2
  rgbRed = rgbBlue + m_r
  rgbGreen = rgbBlue + m_g
  rgbBlue = rgbBlue + m_b
  If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
  If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
  If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
  clr2 = (rgbRed And &HFF&) Or _
  ((rgbGreen And &HFF&) * &H100&) Or _
  ((rgbBlue And &HFF&) * &H10000)
  clr2 = CreateSolidBrush(clr2)
  If i And 8& Then
   clr = clr Xor clr2
   clr2 = clr2 Xor clr
   clr = clr Xor clr2
  End If
  r.Left = 280
  r.Right = 288
  FillRect bm.hdc, r, clr2
  r.Left = 288
  r.Right = 296
  FillRect bm.hdc, r, clr
  r.Left = 296
  r.Right = 304
  FillRect bm.hdc, r, clr2
  DeleteObject clr
  DeleteObject clr2
 Next i
 '////////
 i = m_a
Case 4 'r
 rgbGreen = m_g
 rgbBlue = m_b
 If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
 If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
 clr = ((rgbGreen And &HFF&) * &H100&) Or _
 ((rgbBlue And &HFF&) * &H10000)
 GradientFillRect bm.hdc, 280, 4, 304, 260, clr + &HFF&, clr, GRADIENT_FILL_RECT_V
 i = 255 - m_r
Case 5 'g
 rgbRed = m_r
 rgbBlue = m_b
 If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
 If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
 clr = (rgbRed And &HFF&) Or _
 ((rgbBlue And &HFF&) * &H10000)
 GradientFillRect bm.hdc, 280, 4, 304, 260, clr + &HFF00&, clr, GRADIENT_FILL_RECT_V
 i = 255 - m_g
Case 6 'b
 rgbRed = m_r
 rgbGreen = m_g
 If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
 If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
 clr = (rgbRed And &HFF&) Or _
 ((rgbGreen And &HFF&) * &H100&)
 GradientFillRect bm.hdc, 280, 4, 304, 260, clr + &HFF0000, clr, GRADIENT_FILL_RECT_V
 i = 255 - m_b
End Select
'draw selected
TransparentBlt bm.hdc, 269, i, 9, 9, bm0.hdc, 0, 0, 9, 9, vbGreen
TransparentBlt bm.hdc, 306, i, 9, 9, bm0.hdc, 8, 0, 9, 9, vbGreen
End Sub

Private Sub pDrawColor()
On Error Resume Next
Dim r As RECT
Dim i As Long, j As Long, clr As Long
Dim f As Single
Dim rgbRed As Single, rgbGreen As Single, rgbBlue As Single
'calc new color
f = 255 - m_a
rgbRed = f + m_r
rgbGreen = f + m_g
rgbBlue = f + m_b
If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
clr = (rgbRed And &HFF&) Or _
((rgbGreen And &HFF&) * &H100&) Or _
((rgbBlue And &HFF&) * &H10000)
clr = CreateSolidBrush(clr)
For j = 1 To 25 Step 8
 r.Top = j
 r.Bottom = r.Top + 8
 For i = (j And 15&) To 57 Step 16
  r.Left = i
  r.Right = r.Left + 8
  FillRect bm2.hdc, r, clr
 Next i
Next j
DeleteObject clr
f = f / 2
rgbRed = f + m_r
rgbGreen = f + m_g
rgbBlue = f + m_b
If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
clr = (rgbRed And &HFF&) Or _
((rgbGreen And &HFF&) * &H100&) Or _
((rgbBlue And &HFF&) * &H10000)
clr = CreateSolidBrush(clr)
For j = 1 To 25 Step 8
 r.Top = j
 r.Bottom = r.Top + 8
 For i = ((j + 8&) And 15&) To 57 Step 16
  r.Left = i
  r.Right = r.Left + 8
  FillRect bm2.hdc, r, clr
 Next i
Next j
DeleteObject clr
'calc old color
f = 255 - m_aOld
rgbRed = f + m_rOld
rgbGreen = f + m_gOld
rgbBlue = f + m_bOld
If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
clr = (rgbRed And &HFF&) Or _
((rgbGreen And &HFF&) * &H100&) Or _
((rgbBlue And &HFF&) * &H10000)
clr = CreateSolidBrush(clr)
For j = 1 To 25 Step 8
 r.Top = j + 32
 r.Bottom = r.Top + 8
 For i = (j And 15&) To 57 Step 16
  r.Left = i
  r.Right = r.Left + 8
  FillRect bm2.hdc, r, clr
 Next i
Next j
DeleteObject clr
f = f / 2
rgbRed = f + m_rOld
rgbGreen = f + m_gOld
rgbBlue = f + m_bOld
If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
clr = (rgbRed And &HFF&) Or _
((rgbGreen And &HFF&) * &H100&) Or _
((rgbBlue And &HFF&) * &H10000)
clr = CreateSolidBrush(clr)
For j = 1 To 25 Step 8
 r.Top = j + 32
 r.Bottom = r.Top + 8
 For i = ((j + 8&) And 15&) To 57 Step 16
  r.Left = i
  r.Right = r.Left + 8
  FillRect bm2.hdc, r, clr
 Next i
Next j
DeleteObject clr
End Sub

Private Sub pRGB2HSB(ByVal rgbRed As Single, ByVal rgbGreen As Single, ByVal rgbBlue As Single, ByRef hsbH As Single, ByRef hsbS As Single, ByRef hsbB As Single)
Dim fMax As Single, nMax As Long, fMin As Single
If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
If rgbRed > rgbGreen Then
 If rgbRed > rgbBlue Then
  fMax = rgbRed
  nMax = 1
  If rgbGreen > rgbBlue Then fMin = rgbBlue Else fMin = rgbGreen
 Else
  fMax = rgbBlue
  nMax = 3
  fMin = rgbGreen
 End If
Else
 If rgbGreen > rgbBlue Then
  fMax = rgbGreen
  nMax = 2
  If rgbRed > rgbBlue Then fMin = rgbBlue Else fMin = rgbRed
 Else
  fMax = rgbBlue
  nMax = 3
  fMin = rgbRed
 End If
End If
hsbB = fMax * 100 / 255
If fMax = fMin Then
 hsbS = 0
Else
 fMin = fMax - fMin
 hsbS = 100 * fMin / fMax
 Select Case nMax
 Case 1
  fMax = (rgbGreen - rgbBlue) * 60 / fMin
 Case 2
  fMax = 120 + (rgbBlue - rgbRed) * 60 / fMin
 Case Else
  fMax = 240 + (rgbRed - rgbGreen) * 60 / fMin
 End Select
 If fMax > 360 Then fMax = fMax - 360 Else If fMax < 0 Then fMax = fMax + 360
 hsbH = fMax
End If
End Sub

Private Sub pHSB2RGB(ByVal hsbH As Single, ByVal hsbS As Single, ByVal hsbB As Single, ByRef rgbRed As Single, ByRef rgbGreen As Single, ByRef rgbBlue As Single)
Dim nHue As Long, fMin As Single
hsbH = hsbH - Int(hsbH / 360) * 360
If hsbS < 0 Then hsbS = 0 Else If hsbS > 100 Then hsbS = 100
If hsbB < 0 Then hsbB = 0 Else If hsbB > 100 Then hsbB = 100
hsbB = hsbB * 255 / 100
If hsbS = 0 Then
 rgbRed = hsbB
 rgbGreen = hsbB
 rgbBlue = hsbB
Else
 hsbH = hsbH / 60
 nHue = Int(hsbH)
 hsbH = hsbH - nHue
 hsbS = hsbS / 100
 fMin = hsbB * (1 - hsbS)
 If nHue And 1& Then
  hsbS = hsbB * (1 - hsbS * hsbH)
  hsbH = hsbB
  hsbB = hsbS
 Else
  hsbH = hsbB * (1 - hsbS * (1 - hsbH))
 End If
 If nHue < 2 Then
  rgbRed = hsbB
  rgbGreen = hsbH
  rgbBlue = fMin
 ElseIf nHue < 4 Then
  rgbGreen = hsbB
  rgbBlue = hsbH
  rgbRed = fMin
 Else
  rgbBlue = hsbB
  rgbRed = hsbH
  rgbGreen = fMin
 End If
End If
End Sub


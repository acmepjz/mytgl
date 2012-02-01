VERSION 5.00
Begin VB.Form frmLClr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Properties"
   ClientHeight    =   3360
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   2640
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   224
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   176
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton cmdClr 
      Caption         =   "6"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   5
      Left            =   1800
      TabIndex        =   11
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdClr 
      Caption         =   "5"
      BeginProperty Font 
         Name            =   "Webdings"
         Size            =   11.25
         Charset         =   2
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Index           =   4
      Left            =   1800
      TabIndex        =   10
      Top             =   1440
      Width           =   735
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   840
      TabIndex        =   8
      Top             =   2520
      Width           =   1695
   End
   Begin VB.PictureBox picClr 
      Height          =   1275
      Left            =   120
      ScaleHeight     =   81
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   157
      TabIndex        =   6
      Top             =   120
      Width           =   2415
      Begin MyTGL.ctlWndScroll sbClr 
         Left            =   120
         Top             =   120
         _ExtentX        =   1931
         _ExtentY        =   661
         Orientation     =   1
         NCPaintMode     =   1
      End
   End
   Begin VB.CommandButton cmdClr 
      Caption         =   "Clear"
      Height          =   255
      Index           =   3
      Left            =   960
      TabIndex        =   5
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdClr 
      Caption         =   "Delete"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   1800
      Width           =   735
   End
   Begin VB.CommandButton cmdClr 
      Caption         =   "Copy"
      Height          =   255
      Index           =   1
      Left            =   960
      TabIndex        =   3
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdClr 
      Caption         =   "Add"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   1440
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   960
      TabIndex        =   1
      Top             =   2880
      Width           =   735
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   1800
      TabIndex        =   0
      Top             =   2880
      Width           =   735
   End
   Begin VB.Label Label1 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Index           =   2
      Left            =   840
      MousePointer    =   10  'Up Arrow
      TabIndex        =   12
      Top             =   2160
      Width           =   1695
   End
   Begin VB.Label Label1 
      Caption         =   "Count"
      Height          =   255
      Index           =   1
      Left            =   120
      TabIndex        =   9
      Top             =   2520
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Color"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   7
      Top             =   2160
      Width           =   615
   End
End
Attribute VB_Name = "frmLClr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private bm As New cDIBSection
Private cFnt As New CLogFont

'(1 to count,0 to 1)
Private clrs() As Long, clrc As Long
Private sel As Long

Private b As Boolean

Public Property Get IsChanged() As Boolean
IsChanged = b
End Property

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdClr_Click(Index As Integer)
Dim i As Long
Select Case Index
Case 0 'add
 clrc = clrc + 1
 ReDim Preserve clrs(0 To 1, 1 To clrc)
 clrs(0, clrc) = &HFF000000
 pRefresh
Case 1 'copy
 If sel > 0 And sel <= clrc Then
  clrc = clrc + 1
  ReDim Preserve clrs(0 To 1, 1 To clrc)
  CopyMemory clrs(0, sel + 1), clrs(0, sel), (clrc - sel) * 8&
  pRefresh
 End If
Case 2 'delete
 If sel > 0 And sel <= clrc Then
  If clrc <= 1 Then
   Erase clrs
   clrc = 0
  Else
   If sel < clrc Then CopyMemory clrs(0, sel), clrs(0, sel + 1), (clrc - sel) * 8&
   clrc = clrc - 1
   ReDim Preserve clrs(0 To 1, 1 To clrc)
  End If
  sel = 0
  pRefresh
 End If
Case 3 'clear
 If MsgBox("Are you sure?", vbExclamation + vbYesNo) = vbYes Then
  Erase clrs
  clrc = 0
  sel = 0
  pRefresh
 End If
Case 4 'move up
 If sel > 1 And sel <= clrc Then
  i = clrs(0, sel - 1)
  clrs(0, sel - 1) = clrs(0, sel)
  clrs(0, sel) = i
  i = clrs(1, sel - 1)
  clrs(1, sel - 1) = clrs(1, sel)
  clrs(1, sel) = i
  sel = sel - 1
  pRedraw
 End If
Case 5 'move down
 If sel > 0 And sel < clrc Then
  i = clrs(0, sel + 1)
  clrs(0, sel + 1) = clrs(0, sel)
  clrs(0, sel) = i
  i = clrs(1, sel + 1)
  clrs(1, sel + 1) = clrs(1, sel)
  clrs(1, sel) = i
  sel = sel + 1
  pRedraw
 End If
End Select
End Sub

Private Sub cmdOK_Click()
b = True
Unload Me
End Sub

Private Sub Form_Load()
cFnt.HighQuality = True
Set cFnt.LogFont = Me.Font
bm.Create picClr.Width, picClr.Height
sbClr.NCPaintColor1 = d_CtrlBorder
sbClr.LargeChange(efsVertical) = picClr.ScaleHeight
sbClr.SmallChange(efsVertical) = 16
pRefresh
End Sub

Private Sub pRefresh()
Dim h As Long
h = clrc * 16& - picClr.ScaleHeight
If h > 0 Then
 sbClr.Max(efsVertical) = h
 sbClr.Enabled(efsVertical) = True
Else
 sbClr.Max(efsVertical) = 0
 sbClr.Value(efsVertical) = 0
 sbClr.Enabled(efsVertical) = False
End If
pRedraw
End Sub

Private Sub pClick()
Dim n As Long
If sel > 0 And sel <= clrc Then
 n = clrs(0, sel)
 Label1(2).BackColor = (n And &HFF&) * &H10000 + (n And &HFF00&) + (n And &HFF0000) \ &H10000
 Text1.Text = CStr(clrs(1, sel) + 1)
 pRedraw
End If
End Sub

Private Sub Label1_Click(Index As Integer)
Dim clr As Long, n As Long
If Index = 2 And sel > 0 And sel <= clrc Then
 clr = clrs(0, sel)
 n = ColorPicker(clr)
 If clr <> n Then
  Label1(2).BackColor = (n And &HFF&) * &H10000 + (n And &HFF00&) + (n And &HFF0000) \ &H10000
  clrs(0, sel) = n
  pRedraw
 End If
End If
End Sub

Public Sub GetData(ByRef s As String)
Dim b() As Byte
Dim i As Long
If clrc > 0 Then
 ReDim b(1 To clrc * 5&)
 For i = 1 To clrc
  CopyMemory b(i * 5& - 4), clrs(0, i), 4&
  b(i * 5&) = clrs(1, i) 'error?
 Next i
 s = b
Else
 s = ""
End If
End Sub

Public Sub SetData(ByRef s As String)
Dim i As Long
clrc = LenB(s) \ 5
If clrc > 0 Then
 ReDim clrs(0 To 1, 1 To clrc)
 For i = 1 To clrc
  CopyMemory clrs(0, i), ByVal StrPtr(s) + i * 5 - 5, 5& 'error?
 Next i
Else
 Erase clrs
End If
End Sub

Private Sub picClr_DblClick()
Label1_Click 2
End Sub

Private Sub picClr_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyDelete Then cmdClr_Click 2
End Sub

Private Sub picClr_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
i = (CLng(y) + 16& + sbClr.Value(efsVertical)) \ 16&
If i > 0 And i <= clrc And i <> sel Then
 sel = i
 pClick
End If
End Sub

Private Sub picClr_Paint()
bm.PaintPicture picClr.hdc
End Sub

Private Sub pRedraw()
Dim hbr As Long
Dim r As RECT
Dim i As Long, n As Long, m As Long
Dim j As Long
Dim clr1 As Long, clr2 As Long
'background
r.Right = bm.Width
r.Bottom = bm.Height
hbr = CreateSolidBrush(TranslateColor(vbButtonFace))
FillRect bm.hdc, r, hbr
DeleteObject hbr
'color
n = sbClr.Value(efsVertical) \ 16&
m = picClr.ScaleHeight \ 16& + 1
r.Top = n * 16& - sbClr.Value(efsVertical)
r.Right = picClr.ScaleWidth
For i = 1 To n + m
 If i > clrc Then Exit For
 If i > n Then
  r.Bottom = r.Top + 16
  If sel = i Then GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Checked1, d_Checked2, GRADIENT_FILL_RECT_V
  'text
  cFnt.DrawTextXP bm.hdc, CStr(i), 4, r.Top, 32, 16, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
  'color
  r.Left = 32
  r.Right = 48
  r.Bottom = r.Bottom - 1
  clr1 = clrs(0, i)
  clr1 = (clr1 And &HFF&) * &H10000 + (clr1 And &HFF00&) + (clr1 And &HFF0000) \ &H10000
  If i = clrc Or clrs(1, i) = 0 Then
   hbr = CreateSolidBrush(clr1)
   FillRect bm.hdc, r, hbr
   DeleteObject hbr
  Else
   clr2 = clrs(0, i + 1)
   clr2 = (clr2 And &HFF&) * &H10000 + (clr2 And &HFF00&) + (clr2 And &HFF0000) \ &H10000
   GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, clr1, clr2, GRADIENT_FILL_RECT_V
  End If
  If clrs(1, i) = 0 Then
   cFnt.DrawTextXP bm.hdc, CStr(j), 52, r.Top, 64, 15, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
  Else
   cFnt.DrawTextXP bm.hdc, CStr(j) + "-" + CStr(j + clrs(1, i)), 52, r.Top, 64, 15, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
  End If
  'border
  r.Left = 0
  r.Right = picClr.ScaleWidth
  r.Bottom = r.Bottom + 1
  If sel = i Then
   hbr = CreateSolidBrush(d_Border)
   FrameRect bm.hdc, r, hbr
   DeleteObject hbr
  End If
  r.Top = r.Bottom
 End If
 j = j + clrs(1, i) + 1
Next i
'over
picClr_Paint
End Sub

Private Sub sbClr_Change(eBar As EFSScrollBarConstants)
pRedraw
End Sub

Private Sub sbClr_MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long)
pRedraw
End Sub

Private Sub sbClr_Scroll(eBar As EFSScrollBarConstants)
pRedraw
End Sub

Private Sub Text1_Change()
Dim n As Long
If sel > 0 And sel <= clrc Then
 n = Val(Text1.Text) - 1
 If n < 0 Then n = 0
 If n > 255 Then n = 255
 If n <> clrs(1, sel) Then
  clrs(1, sel) = n
  pRedraw
 End If
End If
End Sub

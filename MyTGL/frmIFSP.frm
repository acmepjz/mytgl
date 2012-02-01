VERSION 5.00
Begin VB.Form frmIFSP 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "IFSP"
   ClientHeight    =   6750
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8400
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
   ScaleHeight     =   450
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   560
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin MyTGL.LeftRight objDrag 
      Height          =   285
      Index           =   6
      Left            =   4080
      Top             =   6360
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   503
   End
   Begin MyTGL.LeftRight objDrag 
      Height          =   285
      Index           =   5
      Left            =   2610
      Top             =   6360
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   503
   End
   Begin MyTGL.LeftRight objDrag 
      Height          =   285
      Index           =   4
      Left            =   1320
      Top             =   6360
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   503
   End
   Begin MyTGL.LeftRight objDrag 
      Height          =   285
      Index           =   3
      Left            =   2610
      Top             =   6060
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   503
   End
   Begin MyTGL.LeftRight objDrag 
      Height          =   285
      Index           =   2
      Left            =   1320
      Top             =   6060
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   503
   End
   Begin MyTGL.LeftRight objDrag 
      Height          =   285
      Index           =   1
      Left            =   2610
      Top             =   5760
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   503
   End
   Begin MyTGL.LeftRight objDrag 
      Height          =   285
      Index           =   0
      Left            =   1320
      Top             =   5760
      Width           =   60
      _ExtentX        =   106
      _ExtentY        =   503
   End
   Begin VB.CheckBox chkP 
      Caption         =   "Solid mapping"
      Height          =   255
      Left            =   2760
      TabIndex        =   24
      Top             =   5760
      Width           =   1335
   End
   Begin MyTGL.FakeToolBar FakeToolBar1 
      Height          =   300
      Left            =   240
      TabIndex        =   22
      Top             =   4320
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   529
      Picture         =   "frmIFSP.frx":0000
      TheString       =   $"frmIFSP.frx":0898
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   6
      Left            =   3150
      TabIndex        =   21
      Top             =   6360
      Width           =   900
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   5
      Left            =   1680
      TabIndex        =   16
      Top             =   6360
      Width           =   900
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   4
      Left            =   390
      TabIndex        =   15
      Top             =   6360
      Width           =   900
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   3
      Left            =   1680
      TabIndex        =   14
      Top             =   6060
      Width           =   900
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   390
      TabIndex        =   13
      Top             =   6060
      Width           =   900
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1680
      TabIndex        =   10
      Top             =   5760
      Width           =   900
   End
   Begin VB.TextBox Text1 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   390
      TabIndex        =   8
      Top             =   5760
      Width           =   900
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   6000
      TabIndex        =   6
      Top             =   6240
      Width           =   1095
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   7200
      TabIndex        =   5
      Top             =   6240
      Width           =   1095
   End
   Begin VB.ListBox lstTrans 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   6.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1050
      Left            =   240
      TabIndex        =   4
      Top             =   4680
      Width           =   3855
   End
   Begin VB.PictureBox picPreview 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   3840
      Left            =   4380
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   3
      Top             =   360
      Width           =   3840
   End
   Begin VB.PictureBox picTrans 
      BackColor       =   &H8000000C&
      BorderStyle     =   0  'None
      Height          =   3840
      Left            =   240
      ScaleHeight     =   256
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   256
      TabIndex        =   2
      Top             =   360
      Width           =   3840
   End
   Begin VB.Frame Frame1 
      Caption         =   "Preview"
      Height          =   4455
      Index           =   1
      Left            =   4260
      TabIndex        =   1
      Top             =   120
      Width           =   4095
      Begin VB.Label Label1 
         Height          =   255
         Index           =   8
         Left            =   120
         TabIndex        =   23
         Top             =   4140
         Width           =   1095
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Transform"
      Height          =   6600
      Index           =   0
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4095
      Begin VB.Label Label1 
         Caption         =   "Color"
         Height          =   255
         Index           =   9
         Left            =   2640
         TabIndex        =   25
         Top             =   5940
         Width           =   375
      End
      Begin VB.Label Label1 
         Caption         =   "Mix"
         Height          =   255
         Index           =   7
         Left            =   2640
         TabIndex        =   20
         Top             =   6240
         Width           =   255
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         BorderStyle     =   1  'Fixed Single
         Height          =   255
         Index           =   6
         Left            =   3030
         MousePointer    =   10  'Up Arrow
         TabIndex        =   19
         Top             =   5940
         Width           =   900
      End
      Begin VB.Label Label1 
         Caption         =   "y2"
         Height          =   255
         Index           =   5
         Left            =   1320
         TabIndex        =   18
         Top             =   6240
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "x2"
         Height          =   255
         Index           =   4
         Left            =   60
         TabIndex        =   17
         Top             =   6240
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "y1"
         Height          =   255
         Index           =   3
         Left            =   1320
         TabIndex        =   12
         Top             =   5940
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "x1"
         Height          =   255
         Index           =   2
         Left            =   60
         TabIndex        =   11
         Top             =   5940
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "y0"
         Height          =   255
         Index           =   1
         Left            =   1320
         TabIndex        =   9
         Top             =   5640
         Width           =   255
      End
      Begin VB.Label Label1 
         Caption         =   "x0"
         Height          =   255
         Index           =   0
         Left            =   60
         TabIndex        =   7
         Top             =   5640
         Width           =   255
      End
   End
End
Attribute VB_Name = "frmIFSP"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Sub FillMemory Lib "kernel32.dll" Alias "RtlFillMemory" (ByRef Destination As Any, ByVal Length As Long, ByVal Fill As Byte)
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hDC As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function Polyline Lib "gdi32.dll" (ByVal hDC As Long, ByRef lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hDC As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
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

Private Declare Function QueryPerformanceCounter Lib "kernel32.dll" (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32.dll" (ByRef lpFrequency As Currency) As Long

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long

Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function SetStretchBltMode Lib "gdi32.dll" (ByVal hDC As Long, ByVal nStretchMode As Long) As Long
Private Const COLORONCOLOR As Long = 3

'////////////////
'TODO:don't use this API
Private Declare Function PlgBlt Lib "gdi32.dll" (ByVal hdcDest As Long, ByRef lpPoint As POINTAPI, ByVal hdcSrc As Long, ByVal nXSrc As Long, ByVal nYSrc As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hbmMask As Long, ByVal xMask As Long, ByVal yMask As Long) As Long

'reference bitmap
Private bmRef As New cAlphaDibSection
Private bm3 As New cAlphaDibSection
Private bHasRef As Boolean
'////////////////

Private Type SAFEARRAYBOUND
    cElements As Long
    lLbound As Long
End Type

Private Type SAFEARRAY2D
    cDims As Integer
    fFeatures As Integer
    cbElements As Long
    cLocks As Long
    pvData As Long
    Bounds(0 To 1) As SAFEARRAYBOUND
End Type

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private transOld As typeIFSPTransform, ox As Long, oy As Long

Private trans() As typeIFSPTransform, m As Long
Private nIsSolid() As Byte

Public bCancel As Boolean

Private bm As New cAlphaDibSection, bm2 As New cAlphaDibSection

Private nSelIndex As Long, nSelPropIndex As Long

Private nMode As Long
Private bDisabled As Boolean

Private Const PI As Double = 3.14159265358979
Private Const TWO_PI As Double = PI * 2
Private Const HALF_PI As Double = PI / 2
Private Const ONE_AND_A_HALF_PI As Double = 3 * PI / 2

Public Property Get TheValue() As String
Dim i As Long, j As Long, k As Long
Dim s As String
If m > 0 Then
 s = Space(m * 16&)
 CopyMemory ByVal StrPtr(s), trans(1), m * 32&
 '&h80000000-1
 '&h7f800000-8
 '&h007fffff-23
 'process floating-point
 For i = 1 To m
  CopyMemory k, trans(i).f2, 4&
  j = (k And &H7F800000) \ &H800000
  If j < 85 Then k = 0 Else If j > 168 Then j = 168
  If nIsSolid(i) Then
   If k = 0 Then
    k = 1
   Else
    j = j - 84
    k = (k And &H807FFFFF) Or (j * &H800000)
   End If
  End If
  CopyMemory ByVal (StrPtr(s) + i * 32& - 4&), k, 4&
 Next i
 TheValue = s
End If
End Property

Public Property Let TheValue(s As String)
Dim i As Long, j As Long, k As Long
m = LenB(s) \ 32&
If m > 0 Then
 ReDim trans(1 To m)
 ReDim nIsSolid(1 To m)
 CopyMemory trans(1), ByVal StrPtr(s), m * 32&
 '&h80000000-1
 '&h7f800000-8
 '&h007fffff-23
 'process floating-point
 For i = 1 To m
  CopyMemory k, trans(i).f2, 4&
  j = (k And &H7F800000) \ &H800000
  If j = 0 Then
   If k And &H7FFFFF Then
    nIsSolid(i) = 1
    trans(i).f2 = 0
   End If
  ElseIf j >= 1 And j <= 84 Then
   nIsSolid(i) = 1
   j = j + 84
   k = (k And &H807FFFFF) Or (j * &H800000)
   CopyMemory trans(i).f2, k, 4&
  End If
 Next i
Else
 m = 0
 Erase trans, nIsSolid
End If
End Property

Private Sub chkP_Click()
Dim i As Long
i = lstTrans.ListIndex + 1
If i > 0 And i <= m Then
 If nIsSolid(i) <> chkP.Value Then
  nIsSolid(i) = chkP.Value
  pRedrawPreview
 End If
End If
End Sub

Private Sub Command1_Click()
bCancel = False
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub pAdd()
m = m + 1
ReDim Preserve trans(1 To m)
ReDim Preserve nIsSolid(1 To m)
With trans(m)
 .f(2) = 1
 .f(5) = 1
 .f2 = 1
 .nClr = -1
End With
lstTrans.AddItem pGetStr(m)
pRedraw
End Sub

Private Sub pCopy()
Dim i As Long
i = lstTrans.ListIndex + 1
If i > 0 And i <= m Then
 m = m + 1
 ReDim Preserve trans(1 To m)
 ReDim Preserve nIsSolid(1 To m)
 trans(m) = trans(i)
 nIsSolid(m) = nIsSolid(i)
 lstTrans.AddItem pGetStr(m)
 pRedraw
End If
End Sub

Private Sub pDelete()
Dim i As Long
i = lstTrans.ListIndex + 1
If i > 0 And i <= m Then
 If m <= 1 Then
  Erase trans, nIsSolid
  m = 0
  lstTrans.Clear
 Else
  If i < m Then
   CopyMemory trans(i), trans(i + 1), (m - i) * 32&
   CopyMemory nIsSolid(i), nIsSolid(i + 1), (m - i) '* 1&
  End If
  m = m - 1
  ReDim Preserve trans(1 To m)
  ReDim Preserve nIsSolid(1 To m)
  lstTrans.RemoveItem i - 1
 End If
 pRedraw
End If
End Sub

Private Sub pClear()
If MsgBox("Are you sure?", vbExclamation + vbYesNo) = vbYes Then
 Erase trans, nIsSolid
 m = 0
 lstTrans.Clear
 pRedraw
End If
End Sub

Private Sub FakeToolBar1_Click(ByVal btnIndex As Long, ByVal btnKey As String)
Dim s As String
Select Case btnIndex
Case 1
 pAdd
Case 2
 pCopy
Case 3
 pDelete
Case 4
 pClear
Case 6 To 9
 nMode = btnIndex - 6
Case 11 'load
 If cd.VBGetOpenFileName(s, , , , , True, "Picture|*.bmp;*.jpg;*.gif", , CStr(App.Path), , , Me.hWnd) Then
  Dim bm0 As New cDIBSection
  bm0.CreateFromFile s
  bmRef.Create 256, 256
  FillMemory ByVal bmRef.DIBSectionBitsPtr, 262144, 255
  SetStretchBltMode bmRef.hDC, COLORONCOLOR
  StretchBlt bmRef.hDC, 0, 0, 256, 256, bm0.hDC, 0, 0, bm0.Width, bm0.Height, vbSrcCopy
  bHasRef = True
  pRedrawTrans
 End If
Case 12 'delete
 bmRef.ClearUp
 bHasRef = False
 pRedrawTrans
End Select
End Sub

Private Sub Form_Load()
cUnk.InitASM
bm.Create 256, 256
bm3.Create 256, 256
bm2.Create 256, -256
pList
pRedraw
bCancel = True
End Sub

Private Sub pList()
Dim i As Long
lstTrans.Clear
For i = 1 To m
 lstTrans.AddItem pGetStr(i)
Next i
End Sub

Private Function pGetStr(ByVal i As Long) As String
Dim s As String
Dim j As Long
s = Format(trans(i).f(0), ".000")
For j = 1 To 5
 s = s + vbTab + Format(trans(i).f(j), ".000")
Next j
pGetStr = s
End Function

Private Sub Label1_Click(Index As Integer)
Dim i As Long, j As Long
i = lstTrans.ListIndex + 1
If i > 0 And i <= m And Index = 6 Then
 With trans(i)
  j = ColorPicker(.nClr)
  If .nClr <> j Then
   .nClr = j
   j = (j And &HFF0000) \ &H10000 + (j And &HFF00&) + (j And &HFF&) * &H10000
   Label1(6).BackColor = j
   pRedrawPreview
  End If
 End With
End If
End Sub

Private Sub lstTrans_Click()
Dim i As Long, j As Long
i = lstTrans.ListIndex + 1
If i > 0 And i <= m Then
 bDisabled = True
 With trans(i)
  For j = 0 To 5
   Text1(j).Text = Val(.f(j))
  Next j
  Text1(6).Text = Val(.f2)
  j = (.nClr And &HFF0000) \ &H10000 + (.nClr And &HFF00&) + (.nClr And &HFF&) * &H10000
 End With
 Label1(6).BackColor = j
 chkP.Value = nIsSolid(i)
 pRedrawTrans
 bDisabled = False
End If
End Sub

Private Sub objDrag_Change(Index As Integer, ByVal iDelta As Long, ByVal Button As Long, ByVal Shift As Long, bCancel As Boolean)
Dim i As Long
Dim f As Single
i = lstTrans.ListIndex + 1
If i > 0 And i <= m Then
 f = iDelta / 1000#
 With trans(i)
  If Index = 6 Then
   f = .f2 + f
  Else
   f = .f(Index) + f
  End If
 End With
 Text1(Index).Text = CStr(f) '??????
End If
End Sub

Private Sub picPreview_Paint()
bm2.PaintPicture picPreview.hDC
End Sub

Private Sub picTrans_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, j As Long
'get pos
ox = x
oy = y
Select Case Button
Case 1
 nSelIndex = 0
 i = lstTrans.ListIndex + 1
 If i > 0 And i <= m Then
  transOld = trans(i)
  j = pHitTest(i, ox, oy)
  If j > 0 Then
   nSelIndex = i
   nSelPropIndex = j
   Exit Sub
  End If
 End If
 For i = 1 To m
  j = pHitTest(i, ox, oy)
  If j > 0 Then
   If i <> lstTrans.ListIndex + 1 Then
    lstTrans.ListIndex = i - 1
   End If
   transOld = trans(i)
   nSelIndex = i
   nSelPropIndex = j
   Exit Sub
  End If
 Next i
 '???
 i = lstTrans.ListIndex + 1
 If i > 0 And i <= m And (nMode = 1 Or nMode = 3) Then nSelIndex = i
End Select
End Sub

Private Sub picTrans_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim xx As Long, yy As Long
Dim fX As Single, fY As Single
Dim xxx As Single, yyy As Single
Dim i As Long
'get pos
xx = x
yy = y
Select Case Button
Case 0
 For i = 1 To m
  If pHitTest(i, xx, yy) > 0 Then
   Select Case nMode
   Case 0, 2
    picTrans.MousePointer = vbSizePointer
   Case Else
    picTrans.MousePointer = vbUpArrow
   End Select
   Exit Sub
  End If
 Next i
 Select Case nMode
 Case 0, 2
  picTrans.MousePointer = vbDefault
 Case Else
  picTrans.MousePointer = vbSizePointer
 End Select
Case 1
 If nSelIndex > 0 Then
  With trans(nSelIndex)
   Select Case nMode
   Case 0
    If nSelPropIndex > 0 Then
     i = nSelPropIndex * 2& - 2
     .f(i) = transOld.f(i) + (xx - ox) / 256#
     .f(i + 1) = transOld.f(i + 1) + (yy - oy) / 256#
     pRedraw True
    End If
   Case 1
    For i = 0 To 4 Step 2
     .f(i) = transOld.f(i) + (xx - ox) / 256#
     .f(i + 1) = transOld.f(i + 1) + (yy - oy) / 256#
    Next i
    pRedraw True
   Case 2 'zoom
    If nSelPropIndex > 0 Then
     i = nSelPropIndex * 2& - 2
     fX = (transOld.f(2) + transOld.f(4)) / 2
     fY = (transOld.f(3) + transOld.f(5)) / 2
     xxx = transOld.f(i) - fX
     yyy = transOld.f(i + 1) - fY
     If Abs(xxx) > Abs(yyy) Then
      If Abs(xxx) < 0.001 Then 'error!!
       xxx = 1
      Else
       xxx = (xx / 256# - fX) / xxx
      End If
     Else
      If Abs(yyy) < 0.001 Then 'error!!
       xxx = 1
      Else
       xxx = (yy / 256# - fY) / yyy
      End If
     End If
     For i = 0 To 4 Step 2
      .f(i) = fX + xxx * (transOld.f(i) - fX)
      .f(i + 1) = fY + xxx * (transOld.f(i + 1) - fY)
     Next i
     pRedraw True
    End If
   Case 3 'rotate
    fX = (transOld.f(2) + transOld.f(4)) / 2
    fY = (transOld.f(3) + transOld.f(5)) / 2
    xxx = ox / 256# - fX
    yyy = oy / 256# - fY
    xxx = Atan2(xxx, yyy) 'old angle
    yyy = Atan2(xx / 256# - fX, yy / 256# - fY) 'new angle
    yyy = xxx - yyy 'delta ???
    For i = 0 To 4 Step 2
     xxx = transOld.f(i) - fX
     .f(i) = fX + Cos(yyy) * xxx + Sin(yyy) * (transOld.f(i + 1) - fY)
     .f(i + 1) = fY - Sin(yyy) * xxx + Cos(yyy) * (transOld.f(i + 1) - fY)
    Next i
    pRedraw True
   End Select
  End With
 End If
End Select
End Sub

Private Function Atan2(ByVal x As Double, ByVal y As Double) As Double
If x = 0 Then
 If y < 0 Then
  Atan2 = ONE_AND_A_HALF_PI
 ElseIf y = 0 Then
  Atan2 = 0
 Else
  Atan2 = HALF_PI
 End If
ElseIf x > 0 Then
 If y < 0 Then
  Atan2 = TWO_PI + Atn(y / x)
 Else
  Atan2 = Atn(y / x)
 End If
Else
 Atan2 = PI + Atn(y / x)
End If
End Function


Private Function pHitTest(ByVal i As Long, ByVal x As Long, ByVal y As Long) As Long
Dim j As Long
Dim ii As Long, jj As Long
With trans(i)
 For j = 0 To 4 Step 2
  ii = .f(j) * 256#
  jj = .f(j + 1) * 256#
  If x >= ii - 2 And y >= jj - 2 And x <= ii + 2 And y <= jj + 2 Then
   pHitTest = j \ 2 + 1
   Exit For
  End If
 Next j
End With
End Function

Private Sub picTrans_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 And nSelIndex > 0 Then
 lstTrans.List(nSelIndex - 1) = pGetStr(nSelIndex)
 lstTrans_Click
End If
End Sub

Private Sub picTrans_Paint()
bm.PaintPicture picTrans.hDC
End Sub

Private Sub pRedraw(Optional ByVal bNoRedrawPreviewInIDE As Boolean)
On Error GoTo a
pRedrawTrans
If bNoRedrawPreviewInIDE Then
 Debug.Print 1 / 0 'debug.assert :-3
End If
pRedrawPreview
a:
End Sub

Private Sub pRedrawPreview()
On Error Resume Next
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim i As Long, j As Long, k As Long, kk As Long
Dim x As Single, y As Single
Dim f() As Single 'probability
Dim TheClr() As RGBQUAD 'color
Dim clr1 As RGBQUAD
Dim TheTable() As Long 'alpha
Dim f2 As Single
Dim t1 As Currency, t2 As Currency
Dim TheSeed As Long
TheSeed = 145
bm2.Cls
If m > 0 Then
 'get time
 QueryPerformanceCounter t1
 'init array
 With tSA
  .cbElements = 4
  .cDims = 1
  .Bounds(0).cElements = 65536
  .pvData = bm2.DIBSectionBitsPtr
 End With
 CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
 'calc probability
 ReDim f(1 To m)
 ReDim TheClr(1 To m)
 ReDim TheTable(1 To m)
 For i = 1 To m
  With trans(i)
   'S=x(0)y(1)-x(1)y(0)
   y = Abs((.f(2) - .f(0)) * (.f(5) - .f(1)) - (.f(4) - .f(0)) * (.f(3) - .f(1)))
   If y < 0.001 Then y = 0.001 'too small?
   CopyMemory TheClr(i), .nClr, 4&
   TheTable(i) = 1024# * .f2
  End With
  x = x + y
  f(i) = x
 Next i
 'normalize
 For i = 1 To m
  f(i) = 2# * f(i) / x - 1
 Next i
 'calc
 x = 0.01
 y = 0.01
 clr1 = TheClr(1)
 For i = 1 To 100020
  f2 = cUnk.fRndFloat(TheSeed)
  For j = 1 To m - 1
   If f2 <= f(j) Then Exit For
  Next j
  Debug.Assert j <= m
  With trans(j)
   If nIsSolid(j) Then
    f2 = cUnk.fRnd(TheSeed) / 32768#
    y = cUnk.fRnd(TheSeed) / 32768#
   Else
    f2 = x
   End If
   x = .f(0) + (.f(2) - .f(0)) * f2 + (.f(4) - .f(0)) * y
   y = .f(1) + (.f(3) - .f(1)) * f2 + (.f(5) - .f(1)) * y
  End With
  'calc color
  kk = TheTable(j)
  With TheClr(j)
   k = clr1.rgbBlue + ((-clr1.rgbBlue + .rgbBlue) * kk) \ 1024&
   If k < 0 Then k = 0 Else If k > 255 Then k = 255
   clr1.rgbBlue = k
   k = clr1.rgbGreen + ((-clr1.rgbGreen + .rgbGreen) * kk) \ 1024&
   If k < 0 Then k = 0 Else If k > 255 Then k = 255
   clr1.rgbGreen = k
   k = clr1.rgbRed + ((-clr1.rgbRed + .rgbRed) * kk) \ 1024&
   If k < 0 Then k = 0 Else If k > 255 Then k = 255
   clr1.rgbRed = k
   k = clr1.rgbReserved + ((-clr1.rgbReserved + .rgbReserved) * kk) \ 1024&
   If k < 0 Then k = 0 Else If k > 255 Then k = 255
   clr1.rgbReserved = k
  End With
  If i > 20 Then
   'calc pos
   k = (CLng(x * 256#) And 255&) + (CLng(y * 65536#) And &HFF00&)
   'draw
   bDib(k) = clr1
  End If
 Next i
 'destroy array
 CopyMemory ByVal VarPtrArray(bDib()), 0&, 4&
 'calc time
 QueryPerformanceCounter t2
 t1 = t2 - t1
 QueryPerformanceFrequency t2
 Label1(8).Caption = "Calc " + Format(t1 / t2 * 1000, "0.00") + "ms"
End If
picPreview_Paint
End Sub

Private Sub pRedrawTrans()
Dim i As Long, idx As Long
Dim hbr As Long
Dim hpn As Long
Dim p(4) As POINTAPI, r As RECT
idx = lstTrans.ListIndex + 1
bm.Cls
'/////draw bitmap?
If bHasRef Then
 bmRef.PaintPicture bm.hDC
 If idx > 0 And idx <= m Then
  With trans(idx)
   p(0).x = .f(0) * 256#
   p(0).y = .f(1) * 256#
   p(1).x = .f(2) * 256#
   p(1).y = .f(3) * 256#
   p(2).x = .f(4) * 256#
   p(2).y = .f(5) * 256#
  End With
  ZeroMemory ByVal bm3.DIBSectionBitsPtr, 262144
  SetStretchBltMode bm3.hDC, COLORONCOLOR
  'TODO:don't use this API
  PlgBlt bm3.hDC, p(0), bmRef.hDC, 0, 0, 256, 256, 0, 0, 0
  bm3.AlphaPaintPicture bm.hDC, , , , , , , 128, False 'True
 End If
End If
'/////
hpn = CreatePen(0, 1, vbWhite)
hpn = SelectObject(bm.hDC, hpn)
For i = 1 To m
 If i = idx Then
  hbr = CreateSolidBrush(&H8080FF)
 Else
  hbr = CreateSolidBrush(&HFF8080)
 End If
 With trans(i)
  p(0).x = .f(0) * 256#
  p(0).y = .f(1) * 256#
  p(1).x = .f(2) * 256#
  p(1).y = .f(3) * 256#
  p(2).x = (.f(2) + .f(4) - .f(0)) * 256#
  p(2).y = (.f(3) + .f(5) - .f(1)) * 256#
  p(3).x = .f(4) * 256#
  p(3).y = .f(5) * 256#
  p(4) = p(0)
  Polyline bm.hDC, p(0), 5
  r.Left = p(0).x - 3
  r.Top = p(0).y - 3
  r.Right = r.Left + 7
  r.Bottom = r.Top + 7
  FillRect bm.hDC, r, hbr
  r.Left = p(1).x - 3
  r.Top = p(1).y - 3
  r.Right = r.Left + 7
  r.Bottom = r.Top + 7
  FillRect bm.hDC, r, hbr
  r.Left = p(3).x - 2
  r.Top = p(3).y - 2
  r.Right = r.Left + 5
  r.Bottom = r.Top + 5
  FillRect bm.hDC, r, hbr
 End With
 DeleteObject hbr
Next i
DeleteObject SelectObject(bm.hDC, hpn)
picTrans_Paint
End Sub

Private Sub Text1_Change(Index As Integer)
Dim i As Long
Dim f As Single
i = lstTrans.ListIndex + 1
If i > 0 And i <= m And Not bDisabled Then
 f = Val(Text1(Index).Text)
 With trans(i)
  If Index = 6 Then
   If f <> .f2 Then
    .f2 = f
    pRedrawPreview
   End If
  ElseIf f <> .f(Index) Then
   .f(Index) = f
   lstTrans.List(i - 1) = pGetStr(i)
   pRedraw
  End If
 End With
End If
End Sub

VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "libMyTGL 1.0 Viewer"
   ClientHeight    =   6000
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   9600
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form1"
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CommandButton Command3 
      Caption         =   "Save as"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4320
      TabIndex        =   7
      Top             =   0
      Width           =   975
   End
   Begin VB.PictureBox p1 
      Appearance      =   0  'Flat
      BackColor       =   &H80000005&
      ForeColor       =   &H80000008&
      Height          =   2895
      Left            =   1440
      ScaleHeight     =   191
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   319
      TabIndex        =   3
      Top             =   1440
      Visible         =   0   'False
      Width           =   4815
      Begin VB.PictureBox p2 
         BackColor       =   &H00800000&
         BorderStyle     =   0  'None
         Height          =   855
         Left            =   0
         ScaleHeight     =   57
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   81
         TabIndex        =   4
         Top             =   0
         Width           =   1215
         Begin VB.Label Label1 
            Alignment       =   2  'Center
            BackStyle       =   0  'Transparent
            Caption         =   "Calculating..."
            ForeColor       =   &H8000000E&
            Height          =   255
            Index           =   1
            Left            =   120
            TabIndex        =   6
            Top             =   120
            Width           =   1335
         End
      End
      Begin VB.Label Label1 
         Alignment       =   2  'Center
         Caption         =   "Calculating..."
         Height          =   255
         Index           =   0
         Left            =   1560
         TabIndex        =   5
         Top             =   120
         Width           =   1335
      End
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Calculate"
      Enabled         =   0   'False
      Height          =   375
      Left            =   3240
      TabIndex        =   2
      Top             =   0
      Width           =   975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Load file"
      Height          =   375
      Left            =   2160
      TabIndex        =   1
      Top             =   0
      Width           =   975
   End
   Begin VB.ComboBox Combo1 
      Height          =   315
      Left            =   0
      Style           =   2  'Dropdown List
      TabIndex        =   0
      Top             =   0
      Width           =   1215
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Function StretchDIBits Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal dx As Long, ByVal dy As Long, ByVal SrcX As Long, ByVal SrcY As Long, ByVal wSrcWidth As Long, ByVal wSrcHeight As Long, ByRef lpBits As Any, ByRef lpBitsInfo As BITMAPINFOHEADER, ByVal wUsage As Long, ByVal dwRop As Long) As Long

Private Type BITMAPFILEHEADER
    bfType As Integer
    bfSize As Long
    bfReserved1 As Integer
    bfReserved2 As Integer
    bfOffBits As Long
End Type

Private Type BITMAPINFOHEADER
    biSize As Long
    biWidth As Long
    biHeight As Long
    biPlanes As Integer
    biBitCount As Integer
    biCompression As Long
    biSizeImage As Long
    biXPelsPerMeter As Long
    biYPelsPerMeter As Long
    biClrUsed As Long
    biClrImportant As Long
End Type

Private tPrj As typeProjectExport

Private Sub Combo1_Click()
Form_Paint
End Sub

Private Sub Command1_Click()
Dim s As String
Dim i As Long
'///
With New cCommonDialog
 If Not .VBGetOpenFileName(s, , , , , True, "Data file|*.dat|All files|*.*", , App.Path, , , Me.hWnd) Then Exit Sub
End With
'///
Combo1.Clear
If Not LibMyTGLLoadFile(tPrj, s) Then
 MsgBox "Can't load file " + s, vbCritical
 Exit Sub
End If
'///
For i = 1 To tPrj.nExportCount
 Combo1.AddItem "[" + CStr(i) + "] " + tPrj.sExportName(i)
Next i
Command2.Enabled = True
Command3.Enabled = False
MsgBox "OK!"
End Sub

Private Sub Command2_Click()
Dim s As String, i As Long
'///
Combo1.Enabled = False
Command1.Enabled = False
Command2.Enabled = False
Command3.Enabled = False
p1.Visible = True
p2.Width = 0
'///
Do
 s = "Calculating..." + CStr(tPrj.nCurrentIndex) + "/" + CStr(tPrj.nOperatorCount)
 Label1(0).Caption = s
 Label1(1).Caption = s
 If tPrj.nOperatorCount > 0 Then p2.Width = p1.Width * tPrj.nCurrentIndex / tPrj.nOperatorCount
 DoEvents
 i = LibMyTGLCalc(tPrj, True)
 If i < 0 Then MsgBox "Calculate error!", vbCritical
Loop While i > 0
If i = 0 Then MsgBox "OK!"
'///
Combo1.Enabled = True
Command1.Enabled = True
Command3.Enabled = True
p1.Visible = False
End Sub

Private Sub Command3_Click()
Dim s As String
Dim idx As Long
Dim t1 As BITMAPFILEHEADER
Dim t2 As BITMAPINFOHEADER
Dim lp As Long
'///
idx = Combo1.ListIndex + 1
If idx <= 0 Or idx > tPrj.nExportCount Then Exit Sub
idx = tPrj.nExportIndex(idx)
If idx <= 0 Or idx > tPrj.nOperatorCount Then Exit Sub
lp = tPrj.tBitmap(idx).DIBSectionBitsPtr
If lp = 0 Then Exit Sub
'///
With t2
 .biSize = Len(t2)
 .biWidth = tPrj.tBitmap(idx).Width
 .biHeight = tPrj.tBitmap(idx).Height
 If .biWidth <= 0 Or .biHeight <= 0 Then Exit Sub
 .biPlanes = 1
 .biBitCount = 32
 .biSizeImage = .biWidth * .biHeight * 4&
 .biHeight = -.biHeight
End With
With t1
 .bfType = &H4D42&
 .bfSize = t2.biSizeImage + &H36&
 .bfOffBits = &H36&
End With
'///
With New cCommonDialog
 If Not .VBGetSaveFileName(s, , , "Bitmap|*.bmp", , App.Path, , "bmp", Me.hWnd) Then Exit Sub
End With
'///
Open s For Output As #1
Close
Open s For Binary As #1
Put #1, 1, t1
Put #1, 15, t2
Put #1, &H37&, tPrj.tBitmap(idx).b
Close
'///
MsgBox "OK!"
End Sub

Private Sub Form_Load()
Form_Resize
End Sub

Private Sub Form_Paint()
Dim idx As Long
Dim t2 As BITMAPINFOHEADER
Dim lp As Long
'///
Me.Cls
'///
idx = Combo1.ListIndex + 1
If idx <= 0 Or idx > tPrj.nExportCount Then Exit Sub
idx = tPrj.nExportIndex(idx)
If idx <= 0 Or idx > tPrj.nOperatorCount Then Exit Sub
lp = tPrj.tBitmap(idx).DIBSectionBitsPtr
If lp = 0 Then Exit Sub
'///
With t2
 .biSize = Len(t2)
 .biWidth = tPrj.tBitmap(idx).Width
 .biHeight = tPrj.tBitmap(idx).Height
 If .biWidth <= 0 Or .biHeight <= 0 Then Exit Sub
 .biPlanes = 1
 .biBitCount = 32
 .biSizeImage = .biWidth * .biHeight * 4&
 .biHeight = -.biHeight
End With
StretchDIBits Me.hdc, 0, Combo1.Height, t2.biWidth, -t2.biHeight, 0, 0, t2.biWidth, -t2.biHeight, ByVal lp, t2, 0, vbSrcCopy
End Sub

Private Sub Form_Resize()
On Error Resume Next
Combo1.Width = Me.ScaleWidth - 192
Command1.Move Me.ScaleWidth - 192, 0, 64, Combo1.Height
Command2.Move Me.ScaleWidth - 128, 0, 64, Combo1.Height
Command3.Move Me.ScaleWidth - 64, 0, 64, Combo1.Height
p1.Move 32, (Me.ScaleHeight + Combo1.Height - 18) / 2, Me.ScaleWidth - 64, 18
Label1(0).Move 0, 0, p1.ScaleWidth, p1.ScaleHeight
Label1(1).Move 0, 0, p1.ScaleWidth, p1.ScaleHeight
End Sub

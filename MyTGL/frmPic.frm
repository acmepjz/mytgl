VERSION 5.00
Begin VB.Form frmPic 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Import"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   7200
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
   ScaleHeight     =   400
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   480
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command1 
      Caption         =   "Export"
      Height          =   255
      Left            =   6420
      TabIndex        =   19
      Top             =   2880
      Width           =   675
   End
   Begin VB.PictureBox p2 
      Height          =   375
      Left            =   240
      ScaleHeight     =   21
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   445
      TabIndex        =   15
      Top             =   4920
      Visible         =   0   'False
      Width           =   6735
      Begin VB.Label Label3 
         Alignment       =   2  'Center
         BackStyle       =   0  'Transparent
         Caption         =   "Label3"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Index           =   2
         Left            =   0
         TabIndex        =   18
         Top             =   60
         Width           =   6675
      End
      Begin VB.Label Label3 
         BackColor       =   &H008080FF&
         Height          =   375
         Index           =   1
         Left            =   0
         TabIndex        =   17
         Top             =   0
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackColor       =   &H00FFC0C0&
         Height          =   375
         Index           =   0
         Left            =   0
         TabIndex        =   16
         Top             =   0
         Width           =   2415
      End
   End
   Begin VB.OptionButton optType 
      Caption         =   "Compressed"
      Height          =   300
      Index           =   1
      Left            =   1440
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   120
      Width           =   1215
   End
   Begin VB.OptionButton optType 
      Caption         =   "Original"
      Height          =   300
      Index           =   0
      Left            =   120
      Style           =   1  'Graphical
      TabIndex        =   13
      Top             =   120
      Value           =   -1  'True
      Width           =   1215
   End
   Begin VB.PictureBox p0 
      BorderStyle     =   0  'None
      Height          =   735
      Index           =   1
      Left            =   240
      ScaleHeight     =   49
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   193
      TabIndex        =   10
      Top             =   3480
      Width           =   2895
      Begin VB.ComboBox cmbComp 
         Height          =   315
         Left            =   0
         Style           =   2  'Dropdown List
         TabIndex        =   11
         Top             =   0
         Width           =   1455
      End
   End
   Begin VB.PictureBox p0 
      BorderStyle     =   0  'None
      Height          =   495
      Index           =   0
      Left            =   3600
      ScaleHeight     =   33
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   49
      TabIndex        =   9
      Top             =   120
      Width           =   735
      Begin VB.Label Label2 
         Caption         =   "Label2"
         Height          =   255
         Left            =   0
         TabIndex        =   12
         Top             =   0
         Width           =   1335
      End
   End
   Begin VB.Frame Frame1 
      Caption         =   "Picture Info"
      Height          =   2175
      Left            =   120
      TabIndex        =   6
      Top             =   3240
      Width           =   6975
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Load"
      Height          =   255
      Left            =   5640
      TabIndex        =   5
      Top             =   2880
      Width           =   675
   End
   Begin VB.TextBox txtFile 
      Height          =   285
      Left            =   480
      Locked          =   -1  'True
      TabIndex        =   4
      Text            =   "(Original)"
      Top             =   2880
      Width           =   5055
   End
   Begin VB.CommandButton cmdCancel 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   6360
      TabIndex        =   2
      Top             =   5520
      Width           =   735
   End
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   375
      Left            =   5520
      TabIndex        =   1
      Top             =   5520
      Width           =   735
   End
   Begin VB.PictureBox p1 
      Height          =   2295
      Left            =   120
      ScaleHeight     =   149
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   461
      TabIndex        =   0
      Top             =   480
      Width           =   6975
      Begin MyTGL.ctlWndScroll sb1 
         Left            =   480
         Top             =   360
         _ExtentX        =   4471
         _ExtentY        =   661
      End
   End
   Begin VB.Label Label1 
      Caption         =   "TODO:Comprssed"
      Height          =   255
      Index           =   2
      Left            =   5760
      TabIndex        =   8
      Top             =   120
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "TODO:Original"
      Height          =   255
      Index           =   1
      Left            =   4560
      TabIndex        =   7
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   "File"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   2880
      Width           =   375
   End
End
Attribute VB_Name = "frmPic"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private bData() As Byte '0-based
Private nDataSize As Long
Private nDataType As Long
'-1=empty
'
'
'255=file :-3

Private bChanged As Boolean, bChanged2 As Boolean

Private bm As New cAlphaDibSection

Public Property Get IsChanged() As Boolean
IsChanged = bChanged And bChanged2
End Property

Public Sub SetData(ByRef sData As String)
If LenB(sData) > 0 Then
 bData = sData
 nDataSize = UBound(bData) + 1
 'data type
 nDataType = bData(0)
Else
 Erase bData
 nDataSize = 0
 nDataType = -1
End If
End Sub

Public Sub GetData(ByRef sData As String)
sData = bData
End Sub

Private Sub cmdCancel_Click()
Unload Me
End Sub

Private Sub cmdOK_Click()
bChanged2 = True
Unload Me
End Sub

Private Sub Command1_Click()
Dim s As String
If bm.Width > 0 And bm.Height > 0 Then
 If cd.VBGetSaveFileName(s, , , "Bitmap|*.bmp", , CStr(App.Path), , "bmp", Me.hWnd) Then
  'TODO:
  bm.SavePicture s
 End If
End If
End Sub

Private Sub Command3_Click()
On Error GoTo a
Dim s As String
Dim b() As Byte, m As Long
If cd.VBGetOpenFileName(s, , , , , True, "Picture|*.bmp;*.jpg;*.gif", , CStr(App.Path), , , Me.hWnd) Then
 'load file
 Open s For Binary As #1
 nDataSize = LOF(1)
 ReDim b(nDataSize - 1)
 Get #1, 1, b
 Close
 nDataType = 255
 ReDim bData(nDataSize)
 CopyMemory bData(1), b(0), nDataSize
 Erase b
 bData(0) = nDataType
 nDataSize = nDataSize + 1
 pLoadData
 pRefresh
 'over
 txtFile.Text = s
 bChanged = True
 optType(1).Enabled = True
End If
Exit Sub
a:
Close
MsgBox "Error!", vbCritical
Erase bData
nDataSize = 0
nDataType = -1
pLoadData
End Sub

Private Sub Form_Load()
With cmbComp
 .AddItem "Original" ':-3
End With
'resize
optType_Click 0
Label2.Move 0, 0, p0(0).ScaleWidth, p0(0).ScaleHeight
'load data
pLoadData
'refresh
pRefresh
End Sub

Private Sub pLoadData()
Dim m As Long
Dim f As Single, f2 As Single
If nDataType < 0 Then 'empty
 bm.ClearUp
 txtFile.Text = "(Nothing)"
 Label2.Caption = "(Nothing)"
 optType(1).Enabled = False
 p2.Visible = False
Else
 Select Case nDataType
 Case 255 'file
  pLoadFile
 Case Else
  'TODO:
 End Select
 Label2.Caption = Label2.Caption + vbCr + CStr(bm.Width) + "x" + CStr(bm.Height)
 m = bm.Width * bm.Height * 4&
 p2.Visible = m > 0
 If m > 0 Then
  f = nDataSize / m
  Label3(0).Width = p2.ScaleWidth * f
  Label3(2).Caption = CStr(nDataSize) + "/" + CStr(m) + "," + Format(f, "0.000%")
  '///test!!! entropy
  f2 = pCalcEntropy(bData) / 8
  Label3(1).Width = p2.ScaleWidth * f * f2
  Label3(2).Caption = Label3(2).Caption + ",Entropy coding:" + Format(nDataSize * f2, "0")
  '///
 End If
End If
End Sub

Private Sub pLoadFile()
':-3333
Dim s As String
Dim b() As Byte
ReDim b(nDataSize - 2)
CopyMemory b(0), bData(1), nDataSize - 1
s = Environ("Temp") + "\MyTGLTemp.tmp"
Open s For Binary As #1
Put #1, 1, b
Close
Erase b
bm.CreateFromFile s
'file format
If bData(1) = 66 And bData(2) = 77 Then 'bitmap
 Label2.Caption = "BMP file"
ElseIf bData(1) = 71 And bData(2) = 73 And bData(3) = 70 And bData(4) = 56 Then 'gif
 Label2.Caption = "GIF file"
Else 'jpg??
 Label2.Caption = "JPG file"
End If
End Sub

Private Sub pRefresh()
'err??
If bm.Width > p1.ScaleWidth Then
 sb1.Enabled(efsHorizontal) = True
 sb1.Max(efsHorizontal) = bm.Width - p1.ScaleWidth
 sb1.LargeChange(efsHorizontal) = p1.ScaleWidth
 sb1.SmallChange(efsHorizontal) = 10
Else
 sb1.Enabled(efsHorizontal) = False
 sb1.Value(efsHorizontal) = 0
End If
If bm.Height > p1.ScaleHeight Then
 sb1.Enabled(efsVertical) = True
 sb1.Max(efsVertical) = bm.Height - p1.ScaleHeight
 sb1.LargeChange(efsVertical) = p1.ScaleHeight
 sb1.SmallChange(efsVertical) = 10
Else
 sb1.Enabled(efsVertical) = False
 sb1.Value(efsVertical) = 0
End If
p1_Paint
End Sub

Private Sub optType_Click(Index As Integer)
Dim i As Long
If Index = 0 Then Frame1.Caption = "Picture Info" Else Frame1.Caption = "Compression"
For i = 0 To p0.UBound
 With p0(i)
  .Move Frame1.Left + 8, Frame1.Top + 16, Frame1.Width - 16, Frame1.Height - 24
  .Visible = i = Index
 End With
Next i
End Sub

Private Sub p1_Paint()
'TODO:
bm.PaintPicture p1.hdc, 0, 0, p1.ScaleWidth, p1.ScaleHeight, sb1.Value(efsHorizontal), sb1.Value(efsVertical)
End Sub

Private Sub sb1_Change(eBar As EFSScrollBarConstants)
p1_Paint
End Sub

Private Sub sb1_MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long)
p1_Paint
End Sub

Private Sub sb1_Scroll(eBar As EFSScrollBarConstants)
p1_Paint
End Sub

Private Function pCalcEntropy(b() As Byte) As Double
Dim nCount(255) As Long
Dim i As Long, j As Long, m As Long
m = LBound(b)
For i = m To UBound(b)
 j = b(i)
 nCount(j) = nCount(j) + 1
Next i
m = i - m
For i = 0 To 255
 If nCount(i) > 0 Then
  pCalcEntropy = pCalcEntropy - nCount(i) / m * Log(nCount(i) / m)
 End If
Next i
pCalcEntropy = pCalcEntropy / Log(2)
End Function

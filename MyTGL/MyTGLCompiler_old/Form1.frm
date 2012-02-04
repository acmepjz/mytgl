VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "MyTGL 1.0 Compiler"
   ClientHeight    =   4500
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   5400
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
   MaxButton       =   0   'False
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   360
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin VB.CheckBox Check1 
      Caption         =   "Export name"
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   840
      Value           =   1  'Checked
      Width           =   1335
   End
   Begin VB.ListBox List1 
      Height          =   3210
      Left            =   120
      Style           =   1  'Checkbox
      TabIndex        =   5
      Top             =   1200
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Deselect all"
      Height          =   375
      Index           =   3
      Left            =   4200
      TabIndex        =   4
      Top             =   1560
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Select all"
      Height          =   375
      Index           =   2
      Left            =   4200
      TabIndex        =   3
      Top             =   1080
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Export"
      Height          =   375
      Index           =   1
      Left            =   4200
      TabIndex        =   2
      Top             =   600
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Open file"
      Height          =   375
      Index           =   0
      Left            =   4200
      TabIndex        =   0
      Top             =   120
      Width           =   1095
   End
   Begin VB.Label Label1 
      Height          =   735
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3975
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private tPrj As typeProject

Private m_sOperatorEnabled(255) As String
Private m_objOperators As New Collection

Private Sub Command1_Click(Index As Integer)
Dim s As String, i As Long, m As Long
Select Case Index
Case 0 'open
 With New cCommonDialog
  If Not .VBGetOpenFileName(s, , , , , True, "MyTGL Texture File|*.myt", , App.Path, , , Me.hwnd) Then Exit Sub
 End With
 i = InStr(1, s, vbNullChar)
 If i > 0 Then s = Left(s, i - 1)
 Label1.Caption = "Loading " + s + "..."
 List1.Clear
 '///
 ChDrive AppPath
 ChDir AppPath
 '///
 If Not LoadPrjFile(tPrj, s) Then
  Label1.Caption = Label1.Caption + "Failed!"
  MsgBox "Can't load " + s, vbCritical
  Exit Sub
 End If
 '///
 Label1.Caption = Label1.Caption + "OK!" + vbCrLf + "Operators:" + CStr(tPrj.nOpCount)
 For i = 1 To tPrj.nOpCount
  If tPrj.Operators(i).nType = int_OpType_Export Then
   m = m + 1
   List1.AddItem tPrj.Operators(i).sProps(0)
   List1.ItemData(m - 1) = i
   List1.Selected(m - 1) = True
  End If
 Next i
 Label1.Caption = Label1.Caption + ",Exports:" + CStr(m)
 '///
Case 1 'export
 pExport Check1.Value
Case 2 'select all
 For i = 0 To List1.ListCount - 1
  List1.Selected(i) = True
 Next i
Case 3 'deselect all
 For i = 0 To List1.ListCount - 1
  List1.Selected(i) = False
 Next i
End Select
End Sub

Private Sub pExport(ByVal bExportName As Boolean)
On Error Resume Next
'///
Dim objDAG As New clsDAG2
Dim i As Long, ii As Long, j As Long, k As Long
Dim m As Long
Dim p As Long
Dim x As Long, y As Long, w As Long
Dim objStore As New Collection
Dim s As String, s1 As String
'///
Dim nOrder() As Long
Dim nSort() As Long, nSortCount As Long
Dim nNewIndex() As Long
Dim nOldIndex() As Long
Dim nLastOccurence() As Long
Dim bOperatorUsed(255) As Boolean
'///
Dim b1() As Byte, b1c As Long, b1m As Long
Dim b2() As Byte, b2c As Long, b2m As Long
'///
If tPrj.nOpCount <= 0 Then Exit Sub
'///
For i = 0 To List1.ListCount - 1
 If List1.Selected(i) Then m = m + 1
Next i
If m = 0 Then Exit Sub
'///init DAG
objDAG.NodeCount = tPrj.nOpCount
'///get all store
For i = 1 To tPrj.nOpCount
 If tPrj.Operators(i).nType = int_OpType_Store Then
  Err.Clear
  objStore.Add i, StringToHex(tPrj.Operators(i).Name)
  If Err.Number Then
   MsgBox "Bad store name: " + tPrj.Operators(i).Name, vbCritical
   Exit Sub
  End If
 End If
Next i
'///get all loads
For i = 1 To tPrj.nOpCount
 If tPrj.Operators(i).nType = int_OpType_Load Then
  Err.Clear
  j = objStore.Item(StringToHex(tPrj.Operators(i).sProps(0)))
  If Err.Number Then
   MsgBox "Bad load name: " + tPrj.Operators(i).sProps(0), vbCritical
   Exit Sub
  End If
  'add edge
  objDAG.AddEdge j, i
 End If
Next i
'///scan for parent nodes
For i = 1 To tPrj.nOpCount
 If tPrj.Operators(i).nType > 0 Then
  y = tPrj.Operators(i).Top
  If y > 0 Then
   x = tPrj.Operators(i).Left
   w = tPrj.Operators(i).Width
   p = tPrj.Operators(i).nPage
   For j = 1 To tPrj.Pages(p).Rows(y - 1).nOpCount
    k = tPrj.Pages(p).Rows(y - 1).idxOp(j)
    If tPrj.Operators(k).nType > 0 Then
     If tPrj.Operators(k).Left < x + w And tPrj.Operators(k).Left + tPrj.Operators(k).Width > x Then
      objDAG.AddEdge k, i
     End If
    End If
   Next j
  End If
 End If
Next i
'///sort it
objDAG.RunTopologicalSort
'///test it
For i = 0 To List1.ListCount - 1
 If List1.Selected(i) Then
  j = List1.ItemData(i)
  If Not objDAG.IsNodeSorted(j) Then
   MsgBox List1.List(i) + " is not in DAG", vbCritical
   Exit Sub
  End If
  objDAG.MarkNodeAndParent j
 End If
Next i
'///get some temp variables
ReDim nNewIndex(1 To tPrj.nOpCount)
ReDim nOldIndex(1 To tPrj.nOpCount)
ReDim nLastOccurence(1 To tPrj.nOpCount)
'///
ReDim nOrder(1 To tPrj.nOpCount)
m = 0
For k = 1 To objDAG.SortedNodeCount
 i = objDAG.SortedNode(k)
 If objDAG.IsNodeMarked(i) Then
  Select Case tPrj.Operators(i).nType
  Case int_OpType_Store, int_OpType_Load, int_OpType_Nop, int_OpType_Export
   If objDAG.DegreeIn(i) <> 1 Then
    MsgBox "Bad input node count", vbCritical
    Exit Sub
   End If
   j = objDAG.InputNode(i, 1)
   If nOldIndex(j) > 0 Then j = nOldIndex(j)
   nOldIndex(i) = j
  Case Is > 0
   m = m + 1
   nNewIndex(i) = m
   nOrder(m) = i
   '///
   For ii = 1 To objDAG.DegreeIn(i)
    j = objDAG.InputNode(i, ii)
    If nOldIndex(j) > 0 Then j = nOldIndex(j)
    If nLastOccurence(j) >= 0 Then nLastOccurence(j) = m
   Next ii
   '///
  End Select
 End If
Next k
'///preserve the export operator
For i = 0 To List1.ListCount - 1
 If List1.Selected(i) Then
  j = List1.ItemData(i)
  If nOldIndex(j) > 0 Then j = nOldIndex(j)
  nLastOccurence(j) = -1
 End If
Next i
'///get the remove order
ReDim nSort(1 To tPrj.nOpCount)
For k = 1 To m
 i = nOrder(k)
 j = nLastOccurence(i)
 If j >= 0 Then nSort(k) = (j * &H10000) Or i
Next k
With New ISort2
 .QuickSort nSort, 1, m
End With
'///start export
b1m = 65536
ReDim b1(b1m - 1)
CopyMemory b1(0), &H4754794D, 4&
CopyMemory b1(4), &H3078454C, 4&
CopyMemory b1(8), m, 4&
b1c = 12
j = 1
For k = 1 To m
 i = nOrder(k)
 If b1c + 256& > b1m Then
  b1m = b1m + 65536
  ReDim Preserve b1(b1m - 1)
 End If
 CopyMemory b1(b1c), tPrj.Operators(i).nType, 4&
 b1c = b1c + 4&
 CopyMemory b1(b1c), objDAG.DegreeIn(i), 4&
 b1c = b1c + 4&
 For ii = 1 To objDAG.DegreeIn(i)
  If b1c + 256& > b1m Then
   b1m = b1m + 65536
   ReDim Preserve b1(b1m - 1)
  End If
  x = objDAG.InputNode(i, ii)
  If nOldIndex(x) > 0 Then x = nOldIndex(x)
  CopyMemory b1(b1c), nNewIndex(x), 4&
  b1c = b1c + 4&
 Next ii
 '///
 If k = m Then Exit For
 '///
 Do
  If j > m Then Exit Do
  If k < (nSort(j) And &HFFFF0000) \ &H10000 Then Exit Do
  '///
  ii = nSort(j) And &HFFFF&
  If nSort(j) And &HFFFF0000 Then
   nSort(j) = nNewIndex(ii) Or &H80000000
   If b1c + 256& > b1m Then
    b1m = b1m + 65536
    ReDim Preserve b1(b1m - 1)
   End If
   CopyMemory b1(b1c), nSort(j), 4&
   b1c = b1c + 4&
  End If
  '///
  j = j + 1
 Loop
Next k
'///
If b1c + List1.ListCount * 8& + 256& > b1m Then
 b1m = b1m + List1.ListCount * 8& + 65536
 ReDim Preserve b1(b1m - 1)
End If
ii = 0
For i = 0 To List1.ListCount - 1
 If List1.Selected(i) Then
  j = List1.ItemData(i)
  If nOldIndex(j) > 0 Then j = nOldIndex(j)
  CopyMemory b1(b1c + ii * 8& + 4), nNewIndex(j), 4&
  If bExportName Then j = LenB(List1.List(i)) Else j = 0
  '///
  If j > 0 Then
   If b2c + j + 256& > b2m Then
    b2m = b2c + j + 65536
    ReDim Preserve b2(b2m - 1)
   End If
   CopyMemory b2(b2c), ByVal StrPtr(List1.List(i)), j
   b2c = b2c + j
  End If
  '///
  CopyMemory b1(b1c + ii * 8& + 8), j, 4&
  ii = ii + 1
 End If
Next i
CopyMemory b1(b1c), ii, 4&
b1c = b1c + ii * 8& + 4
'///operator data
For k = 1 To m
 i = nOrder(k)
 Err.Clear
 j = UBound(tPrj.Operators(i).bProps) + 1
 If Err.Number Then j = 0
 If b1c + 256& > b1m Then
  b1m = b1m + 65536
  ReDim Preserve b1(b1m - 1)
 End If
 CopyMemory b1(b1c), j, 4&
 b1c = b1c + 4
 '///
 If j > 0 Then
  If b2c + j + 256& > b2m Then
   b2m = b2c + j + 65536
   ReDim Preserve b2(b2m - 1)
  End If
  CopyMemory b2(b2c), tPrj.Operators(i).bProps(0), j
  b2c = b2c + j
 End If
 '///
 Err.Clear
 ii = UBound(tPrj.Operators(i).sProps) + 1
 If Err.Number Then ii = 0
 If b1c + ii * 4& + 256& > b1m Then
  b1m = b1m + ii * 4& + 65536
  ReDim Preserve b1(b1m - 1)
 End If
 CopyMemory b1(b1c), ii, 4&
 b1c = b1c + 4
 '///
 For ii = 0 To ii - 1
  j = LenB(tPrj.Operators(i).sProps(ii))
  CopyMemory b1(b1c), j, 4&
  b1c = b1c + 4
  '///
  If j > 0 Then
   If b2c + j + 256& > b2m Then
    b2m = b2c + j + 65536
    ReDim Preserve b2(b2m - 1)
   End If
   CopyMemory b2(b2c), ByVal StrPtr(tPrj.Operators(i).sProps(ii)), j
   b2c = b2c + j
  End If
  '///
 Next ii
Next k
'///
For k = 1 To m
 i = tPrj.Operators(nOrder(k)).nType
 If i >= 0 And i <= 255 Then
  bOperatorUsed(i) = True
  If m_sOperatorEnabled(i) = vbNullString Then
   Label1.Caption = Label1.Caption + vbCrLf + "Warning: Unknown operator " + CStr(i)
  ElseIf i = 9 Then
   Label1.Caption = Label1.Caption + vbCrLf + "Currently Import unsupported"
  ElseIf i = 12 Then
   Label1.Caption = Label1.Caption + vbCrLf + "Currently LSystem unsupported"
  End If
 Else
  Label1.Caption = Label1.Caption + vbCrLf + "Warning: Unknown operator " + CStr(i)
 End If
Next k
'///
If SrcFile <> vbNullString Then
 Open SrcFile For Input As #1
 Open App.Path + "\mdlCalc.bas.out.bas" For Output As #2
 Do Until EOF(1)
  Line Input #1, s
  If LCase(Left(s, 6)) = "#const" Then
   i = InStr(1, s, " ")
   If i > 0 Then
    s1 = Trim(Mid(s, i + 1))
    i = InStr(1, s1, " ")
    If i > 0 Then
     s1 = Trim(Left(s1, i - 1))
     If LCase(s1) = "release" Then
      s = "#Const Release = True"
     Else
      Err.Clear
      i = m_objOperators.Item(s1)
      If Err.Number = 0 Then
       s = "#Const " + m_sOperatorEnabled(i) + " = " + CStr(bOperatorUsed(i))
      End If
     End If
    End If
   End If
  End If
  Print #2, s
 Loop
 Close
End If
'///over
If b1c > 0 Then ReDim Preserve b1(b1c - 1)
If b2c > 0 Then ReDim Preserve b2(b2c - 1)
Open App.Path + "\out.dat" For Output As #1
Close
Open App.Path + "\out.dat" For Binary As #1
If b1c > 0 Then Put #1, , b1
If b2c > 0 Then Put #1, , b2
Close
'///
MsgBox "OK! Nodes exported:" + CStr(m)
End Sub

Private Sub Form_Load()
On Error Resume Next
Dim i As Long
'///
Err.Clear
AppPath = App.Path + "\"
i = GetAttr(AppPath + "Prop.def")
If Err.Number = 0 And (i And vbDirectory) = 0 Then
Else
 Err.Clear
 AppPath = App.Path + "\..\"
 i = GetAttr(AppPath + "Prop.def")
 If Err.Number = 0 And (i And vbDirectory) = 0 Then
 Else
  MsgBox "Can't find file Prop.def", vbCritical
  End
 End If
End If
'///
LoadOperationDef
'///
Err.Clear
SrcFile = AppPath + "libMyTGL_old\mdlCalc.bas"
i = GetAttr(SrcFile)
If Err.Number = 0 And (i And vbDirectory) = 0 Then
Else
 MsgBox "Can't find file libMyTGL_old\mdlCalc.bas", vbCritical
 SrcFile = vbNullString
End If
'///
m_sOperatorEnabled(1) = "FlatOperatorEnabled"
m_sOperatorEnabled(2) = "CloudOperatorEnabled"
m_sOperatorEnabled(3) = "GradientOperatorEnabled"
m_sOperatorEnabled(4) = "Gradient2OperatorEnabled"
m_sOperatorEnabled(5) = "CellOperatorEnabled"
m_sOperatorEnabled(6) = "NoiseOperatorEnabled"
m_sOperatorEnabled(7) = "BrickOperatorEnabled"
m_sOperatorEnabled(8) = "PerlinOperatorEnabled"
m_sOperatorEnabled(9) = "ImportOperatorEnabled"
m_sOperatorEnabled(13) = "IFSPOperatorEnabled"
m_sOperatorEnabled(12) = "LSystemOperatorEnabled"
m_sOperatorEnabled(11) = "SlowGrowOperatorEnabled"
m_sOperatorEnabled(14) = "RectOperatorEnabled"
m_sOperatorEnabled(15) = "PixelsOperatorEnabled"
m_sOperatorEnabled(16) = "GlowRectOperatorEnabled"
m_sOperatorEnabled(17) = "CrackOperatorEnabled"
m_sOperatorEnabled(20) = "BlurOperatorEnabled"
m_sOperatorEnabled(21) = "ColorOperatorEnabled"
m_sOperatorEnabled(22) = "RangeOperatorEnabled"
m_sOperatorEnabled(23) = "HSCBOperatorEnabled"
m_sOperatorEnabled(24) = "NormalsOperatorEnabled"
m_sOperatorEnabled(25) = "ColorBalanceOperatorEnabled"
m_sOperatorEnabled(26) = "RotZoomOperatorEnabled"
m_sOperatorEnabled(27) = "RotateMulOperatorEnabled"
m_sOperatorEnabled(28) = "SharpenOperatorEnabled"
m_sOperatorEnabled(29) = "DialectOperatorEnabled"
m_sOperatorEnabled(34) = "DistortOperatorEnabled"
m_sOperatorEnabled(35) = "BumpOperatorEnabled"
m_sOperatorEnabled(36) = "AddOperatorEnabled"
m_sOperatorEnabled(37) = "MaskOperatorEnabled"
m_sOperatorEnabled(38) = "ParticleOperatorEnabled"
m_sOperatorEnabled(39) = "SegmentOperatorEnabled"
m_sOperatorEnabled(40) = "BulgeOperatorEnabled"
m_sOperatorEnabled(41) = "TwirlOperatorEnabled"
m_sOperatorEnabled(42) = "UnwrapOperatorEnabled"
m_sOperatorEnabled(43) = "AbnormalsOperatorEnabled"
'///
For i = 0 To 255
 If m_sOperatorEnabled(i) <> vbNullString Then m_objOperators.Add i, m_sOperatorEnabled(i)
Next i
End Sub


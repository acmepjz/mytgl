VERSION 5.00
Begin VB.UserControl ctlOpPropGrid 
   BorderStyle     =   1  'Fixed Single
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
   Begin VB.Timer tmr1 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1800
      Top             =   1920
   End
   Begin VB.TextBox t1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   285
      Left            =   720
      TabIndex        =   0
      Top             =   1920
      Visible         =   0   'False
      Width           =   735
   End
   Begin MyTGL.LeftRight lr1 
      Height          =   255
      Left            =   1560
      Top             =   1920
      Visible         =   0   'False
      Width           =   135
      _ExtentX        =   238
      _ExtentY        =   450
   End
   Begin MyTGL.FakeComboBox cmb1 
      Height          =   255
      Left            =   720
      TabIndex        =   1
      Top             =   2280
      Visible         =   0   'False
      Width           =   735
      _ExtentX        =   1296
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BorderStyle     =   -1  'True
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
      DropdownHeight  =   16
   End
   Begin VB.Image i0 
      Height          =   480
      Left            =   720
      Picture         =   "ctlOpPropGrid.ctx":0000
      Top             =   1320
      Visible         =   0   'False
      Width           =   480
   End
   Begin MyTGL.ctlWndScroll sb1 
      Left            =   720
      Top             =   960
      _ExtentX        =   5318
      _ExtentY        =   450
      Orientation     =   1
      NCPaintMode     =   1
   End
End
Attribute VB_Name = "ctlOpPropGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////
'This file is part of MyTGL, an opensource procedural media creation tool and library.
'Copyright (C) 2008,2009  acme_pjz
'
'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <http://www.gnu.org/licenses/>.
'////////////////////////////////

Private Const TheBorderColor = &H80511C

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As Long) As Long
Private Declare Function ScreenToClient Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As Long) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Const WM_SETFONT As Long = &H30

'////////fake pointer

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (Ptr() As Any) As Long
Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

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

Private m_lpData As Long, m_Data() As typeMyTGL11OperatorProperties, m_tData As SAFEARRAY2D
Private m_lpDef As Long, m_Def() As typeMyTGL11OperatorPropDefs, m_tDef As SAFEARRAY2D

'////////

Private Type typeOperatorPropDisplay
 nPropDefIndex As Long
 nDataIndex As Long
 nFlags As Long
 '0 - &HFF = level
 '&H100=invisible
 '&H200=group
 '&H400=collapsed
 y As Long
 h As Long
End Type

Private ps() As typeOperatorPropDisplay, pc As Long

'////////

Private m_Height As Long 'total height
Private idxHl As Long, idxSelected As Long
'&Hyyxxxxxx
'xxxxxx=index of OperatorPropDisplay >0
'yy=sub index 0-xx=index &HFE=left &HFF=custom
'&HFF000000 = splitter
Private nDelta As Long
Private bChanging As Boolean

Private cFnt As New CLogFont
Private bm As New cDIBSection, bm0 As New cDIBSection
Private nCaptionWidth As Long

Private nLastY As Long 'last scroll pos
Private objLast As IOperatorPropCallback 'last edited custom callback

Public Event Change(ByVal idxProp As Long, ByVal idxPropDef As Long, ByVal sKey As String, Data() As Byte, ByRef nDataLength As Long)

Private WithEvents objMenu As FakeMenu
Attribute objMenu.VB_VarHelpID = -1

Private Sub pScroll()
Dim i As Long, j As Long, h As Long
pEditEnd
i = sb1.Visible(efsVertical) And sb1.Value(efsVertical)  ':-3
If i = nLastY Then Exit Sub
h = bm.Height
If i > nLastY Then 'down
 j = i - nLastY
 If j < h Then bm.PaintPicture bm.hdc, 0, 0, , h - j, 0, j, vbSrcCopy
 j = nLastY + h
 nLastY = i
 For i = pc To 1 Step -1
  If ps(i).y + ps(i).h > j Then pRedrawOne i Else Exit For
 Next i
Else 'up
 j = nLastY - i
 If j < h Then bm.PaintPicture bm.hdc, 0, j, , h - j, 0, 0, vbSrcCopy
 j = nLastY
 nLastY = i
 For i = 1 To pc
  If ps(i).y < j Then pRedrawOne i Else Exit For
 Next i
End If
UserControl_Paint
End Sub

Private Sub cmb1_Click()
Dim i As Long, j As Long
If bChanging Then Exit Sub
If pc = 0 Then Exit Sub
i = idxSelected And &HFFFFFF
j = idxSelected \ &H1000000
If i > 0 And j >= 0 Then pChange i, j, CStr(cmb1.ListIndex)
End Sub

Private Sub cmb1_KeyDown(ByVal KeyCode As Long, ByVal Shift As Long)
On Error Resume Next
If KeyCode = vbKeyEscape Then
 cmb1.ListIndex = Val(cmb1.Tag)
 pEditEnd
ElseIf KeyCode = vbKeyReturn Then
 pEditEnd
End If
End Sub

Private Sub cmb1_MyLostFocus()
pEditEnd
End Sub

Private Sub lr1_Change(ByVal iDelta As Long, ByVal Button As Long, ByVal Shift As Long, bCancel As Boolean)
Dim i As Long, j As Long
If bChanging Then Exit Sub
If pc = 0 Then Exit Sub
i = idxSelected And &HFFFFFF
j = idxSelected \ &H1000000
If i > 0 And j >= 0 And Button = 1 Then pDelta i, j, iDelta, Shift, bCancel
End Sub

Private Sub objMenu_Click(ByVal idxMenu As Long, ByVal Key As String, ByVal idxButton As Long, ByVal ButtonKey As String, Value As Long)
Dim i As Long
If Key = "____reset_prop" Then
 Select Case idxButton
 Case 1 'reset
  i = idxSelected And &HFFFFFF
  If i > 0 Then pReset i, idxSelected \ &H1000000
 Case 2 'reset all
  i = idxSelected And &HFFFFFF
  If i > 0 Then pReset i, -1
 End Select
End If
End Sub

Private Sub pReset(ByVal idx As Long, ByVal nSubPropIndex As Long)
On Error Resume Next
Dim obj As IOperatorPropCallback
Dim b As Boolean, lp As Long, sKey As String
Dim i As Long, j As Long, n As Long
Dim f As Single
With ps(idx)
 lp = m_Def(0).d(.nPropDefIndex).nType
 With g_PropTypeDefs(lp)
  Set obj = .objCallback
  sKey = .sKey
 End With
 lp = VarPtr(g_PropTypeDefs(lp))
 'check callback
 If obj Is Nothing Then
  '////////////////////////////////default process
  Select Case sKey
  Case "int"
   lp = VarPtr(m_Data(0).d(.nDataIndex).d(0))
   With m_Def(0).d(.nPropDefIndex)
    For i = 0 To .nElementCount - 1
     If nSubPropIndex = i Or nSubPropIndex < 0 Then
      'get default
      j = 0
      Err.Clear
      n = UBound(.datDefault) + 1
      If Err.Number Then n = 0
      If n >= i * 4& + 4& Then
       CopyMemory j, .datDefault(i * 4&), 4&
      ElseIf n = 4& Then
       CopyMemory j, .datDefault(0), 4&
      End If
      'change
      CopyMemory ByVal lp + i * 4&, j, 4&
     End If
    Next i
   End With
   b = True
  Case "boolean", "size", "resize"
   lp = VarPtr(m_Data(0).d(.nDataIndex).d(0))
   With m_Def(0).d(.nPropDefIndex)
    For i = 0 To .nElementCount - 1
     If nSubPropIndex = i Or nSubPropIndex < 0 Then
      'get default
      j = 0
      Err.Clear
      n = UBound(.datDefault) + 1
      If Err.Number Then n = 0
      If n >= i + 1& Then
       j = .datDefault(i)
      ElseIf n = 1& Then
       j = .datDefault(0)
      End If
      'change
      CopyMemory ByVal lp + i, j, 1&
     End If
    Next i
   End With
   b = True
  Case "float", "color"
   lp = VarPtr(m_Data(0).d(.nDataIndex).d(0))
   If sKey = "color" Then
    If nSubPropIndex >= 0 Then nSubPropIndex = (nSubPropIndex + 3&) And 3&
   End If
   With m_Def(0).d(.nPropDefIndex)
    For i = 0 To .nElementCount - 1
     If nSubPropIndex = i Or nSubPropIndex < 0 Then
      'get default
      f = 0
      Err.Clear
      n = UBound(.datDefault) + 1
      If Err.Number Then n = 0
      If n >= i * 4& + 4& Then
       CopyMemory f, .datDefault(i * 4&), 4&
      ElseIf n = 4& Then
       CopyMemory f, .datDefault(0), 4&
      End If
      'change
      CopyMemory ByVal lp + i * 4&, f, 4&
     End If
    Next i
   End With
   b = True
  End Select
  '////////////////////////////////
 Else
  b = obj.Reset(Me, _
  m_Data(0).d(.nDataIndex).d, VarPtr(m_Def(0).d(.nPropDefIndex)), lp, nSubPropIndex)
 End If
 'changed?
 If b Then
  lp = .nDataIndex
  i = .nPropDefIndex
  With m_Data(0).d(lp)
   Err.Clear
   .nSize = UBound(.d) + 1
   If Err.Number Then .nSize = 0
   RaiseEvent Change(lp, i, .sKey, .d, .nSize)
  End With
  pRedrawOne idx
  UserControl_Paint
 End If
End With
End Sub

Private Sub sb1_Change(eBar As EFSScrollBarConstants)
pScroll
End Sub

Private Sub sb1_MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long)
pScroll
End Sub

Private Sub sb1_Scroll(eBar As EFSScrollBarConstants)
pScroll
End Sub

Private Sub t1_Change()
Dim i As Long, j As Long
If bChanging Then Exit Sub
If pc = 0 Then Exit Sub
i = idxSelected And &HFFFFFF
j = idxSelected \ &H1000000
If i > 0 And j >= 0 Then pChange i, j, t1.Text
End Sub

Private Sub t1_KeyDown(KeyCode As Integer, Shift As Integer)
If KeyCode = vbKeyEscape Then
 t1.Text = t1.Tag
 pEditEnd
ElseIf KeyCode = vbKeyReturn Then
 pEditEnd
End If
End Sub

Private Sub t1_KeyPress(KeyAscii As Integer)
If KeyAscii = vbKeyEscape Or KeyAscii = vbKeyReturn Then KeyAscii = 0
End Sub

Private Sub t1_LostFocus()
pEditEnd
End Sub

Private Sub tmr1_Timer()
Dim p(1) As Long
GetCursorPos p(0)
ScreenToClient hwnd, p(0)
If p(0) < 0 Or p(0) >= ScaleWidth Or p(1) < 0 Or p(1) >= ScaleHeight Then
 If idxHl <> 0 Then
  p(0) = idxHl And &HFFFFFF
  idxHl = 0
  If p(0) > 0 Then
   pRedrawOne p(0)
   UserControl_Paint
  End If
 End If
 tmr1.Enabled = False
End If
End Sub

Private Sub UserControl_DblClick()
Dim p(1) As Long
GetCursorPos p(0)
ScreenToClient hwnd, p(0)
UserControl_MouseDown 1, 0, (p(0)), (p(1))
End Sub

Private Sub UserControl_InitProperties()
pInit
End Sub

Public Sub ExpandAll()
Dim i As Long
For i = 1 To pc
 With ps(i)
  .nFlags = .nFlags And Not &H500&
 End With
Next i
pRefresh
End Sub

Public Sub CollapseAll()
Dim i As Long
For i = 1 To pc
 With ps(i)
  .nFlags = .nFlags Or ((.nFlags And &H200&) * 2&) Or (((.nFlags And &HFF&) > 0) And &H100&)
 End With
Next i
pRefresh
End Sub

Private Sub pExpand(ByVal idx As Long)
Dim i As Long
Dim j As Long, jj As Long
Dim k As Long, kk As Long
j = ps(idx).nFlags
jj = -1
If (j And &H700&) = &H600& Then
 ps(idx).nFlags = j And Not &H400&
 j = j And &HFF&
 For i = idx + 1 To pc
  k = ps(i).nFlags
  kk = k And &HFF&
  'exit tree?
  If kk <= j Then Exit For
  'exit collapsed tree?
  If kk <= jj Then jj = -1
  'set visible
  If jj < 0 Then k = k And Not &H100& Else k = k Or &H100&
  'collapsed tree?
  If (k And &H600&) = &H600& And jj < 0 Then jj = kk
  ps(i).nFlags = k
 Next i
End If
End Sub

Private Sub pCollapse(ByVal idx As Long)
Dim i As Long, j As Long, k As Long
j = ps(idx).nFlags
If (j And &H700&) = &H200& Then
 ps(idx).nFlags = j Or &H400&
 j = j And &HFF&
 For i = idx + 1 To pc
  k = ps(i).nFlags
  'exit tree?
  If (k And &HFF&) <= j Then Exit For
  ps(i).nFlags = k Or &H100&
 Next i
End If
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
Dim idx As Long, idxOld As Long
Dim bDblClick As Boolean
pEditEnd
If pc = 0 Then Exit Sub
'splitter
If idxHl = &HFF000000 And Button = 1 Then
 nDelta = nCaptionWidth - x
 UserControl_Paint
 bm.PaintPicture hdc, nCaptionWidth - 2, 0, 5, , , , vbDstInvert
 MousePointer = vbSizeWE
 tmr1.Enabled = False
 Exit Sub
End If
'check clicked item TODO:
idx = idxHl And &HFFFFFF
idxOld = idxSelected And &HFFFFFF
bDblClick = idxHl = idxSelected
idxSelected = idxHl
If idx > 0 Then
 i = ps(idx).nFlags
 If i And &H200& Then 'group
  If i And &H400& Then pExpand idx Else pCollapse idx
  pRefresh
  Exit Sub
 Else
  i = idxHl And &HFF000000
  If i = &HFF000000 Then
   'custom button
   pCustom idx
   If idxOld > 0 Then pRedrawOne idxOld
   If idx > 0 And idx <> idxOld Then pRedrawOne idx
   UserControl_Paint
   Exit Sub
  ElseIf i = &HFE000000 Then
   'all item
   If Not bDblClick Then
    If idxOld > 0 Then pRedrawOne idxOld
    If idx > 0 And idx <> idxOld Then pRedrawOne idx
    UserControl_Paint
   End If
   If Button = 2 And idx > 0 Then
    If Not objMenu Is Nothing Then
     i = ps(idx).nPropDefIndex
     If (g_PropTypeDefs(m_Def(0).d(i).nType).nFlags And 2&) = 0 Then
      i = objMenu.FindMenu("____reset_prop")
      objMenu.ButtonFlags(i, 1) = fbtfDisabled
      objMenu.PopupMenu "____reset_prop"
     End If
    End If
   End If
   Exit Sub
  Else
   If Button = 1 Then
    pEdit idx, i \ &H1000000, bDblClick
   End If
   If idxOld > 0 Then pRedrawOne idxOld
   If idx > 0 And idx <> idxOld Then pRedrawOne idx
   If Button = 2 And idx > 0 Then
    If Not objMenu Is Nothing Then
     i = ps(idx).nPropDefIndex
     If (g_PropTypeDefs(m_Def(0).d(i).nType).nFlags And 2&) = 0 Then
      i = objMenu.FindMenu("____reset_prop")
      objMenu.ButtonFlags(i, 1) = 0
      objMenu.PopupMenu "____reset_prop"
     End If
    End If
   End If
   UserControl_Paint
   Exit Sub
  End If
 End If
End If
End Sub

Private Sub pEdit(ByVal idx As Long, ByVal nSubPropIndex As Long, ByVal bDblClick As Boolean)
On Error Resume Next
Dim obj As IOperatorPropCallback
Dim b As Boolean, lp As Long, sKey As String
Dim r As RECT, nTop As Long, w As Long
Dim m As Long, n As Long, f As Single
w = bm.Width
With ps(idx)
 nTop = .y - nLastY
 lp = m_Def(0).d(.nPropDefIndex).nType
 With g_PropTypeDefs(lp)
  Set obj = .objCallback
  sKey = .sKey
  If .nFlags And 1& Then w = w - 18
 End With
 lp = VarPtr(g_PropTypeDefs(lp))
 'calc size
 m = m_Def(0).d(.nPropDefIndex).nElementCount
 If m = 0 Then m = 1
 r.Top = nTop + (nSubPropIndex And &HFFFFFFFC) * 4&
 r.Bottom = r.Top + 16
 If nSubPropIndex < (m And &HFFFFFFFC) Then n = 4 Else n = m And 3&
 r.Left = nCaptionWidth + ((w - nCaptionWidth) * (nSubPropIndex And 3&)) \ n
 r.Right = nCaptionWidth + ((w - nCaptionWidth) * ((nSubPropIndex And 3&) + 1&)) \ n
 'set dirty flag
 bChanging = True
 'check callback
 If obj Is Nothing Then
  '////////////////////////////////default process
  Select Case sKey
  Case "int"
   CopyMemory n, m_Data(0).d(.nDataIndex).d(nSubPropIndex * 4&), 4&
   'check enum
   With m_Def(0).d(.nPropDefIndex)
    If nSubPropIndex < .nEnumCount Then
     If bDblClick Then
      cmb1.Clear
      With .datEnum(nSubPropIndex)
       For m = 1 To .nCount
        cmb1.AddItem .d(m).sCaption
        If .d(m).nValue = n Then cmb1.ListIndex = m - 1
       Next m
      End With
      If r.Right >= w Then r.Right = w - 1
      cmb1.Move r.Left, r.Top, r.Right - r.Left + 1, r.Bottom - r.Top + 1
      cmb1.Visible = True
      cmb1.SetFocus
      cmb1.ShowDropdown
     End If
    Else
     If bDblClick Then
      t1.Move r.Left + 1, r.Top + 1, r.Right - r.Left - 1, r.Bottom - r.Top - 1
      t1.Text = CStr(n)
      t1.Visible = True
      t1.SetFocus
     Else
      lr1.Move r.Right - 4, r.Top + 1, 4, r.Bottom - r.Top - 1
      lr1.Visible = True
     End If
    End If
   End With
  Case "size", "resize"
   If bDblClick Then
    cmb1.Clear
    If sKey = "resize" Then cmb1.AddItem "(Current)"
    For m = 0 To int_Size_Max
     cmb1.AddItem CStr(nBitFieldMask1(m))
    Next m
    n = m_Data(0).d(.nDataIndex).d(nSubPropIndex)
    If n < cmb1.ListCount Then cmb1.ListIndex = n
    If r.Right >= w Then r.Right = w - 1
    cmb1.Move r.Left, r.Top, r.Right - r.Left + 1, r.Bottom - r.Top + 1
    cmb1.Visible = True
    cmb1.SetFocus
    cmb1.ShowDropdown
   End If
  Case "float"
   If bDblClick Then
    CopyMemory f, m_Data(0).d(.nDataIndex).d(nSubPropIndex * 4&), 4&
    t1.Move r.Left + 1, r.Top + 1, r.Right - r.Left - 1, r.Bottom - r.Top - 1
    t1.Text = CStr(f)
    t1.Visible = True
    t1.SetFocus
   Else
    lr1.Move r.Right - 4, r.Top + 1, 4, r.Bottom - r.Top - 1
    lr1.Visible = True
   End If
  Case "color"
   If bDblClick Then
    CopyMemory f, m_Data(0).d(.nDataIndex).d((nSubPropIndex * 4& + 12&) And &HF&), 4&
    t1.Move r.Left + 1, r.Top + 1, r.Right - r.Left - 1, r.Bottom - r.Top - 1
    t1.Text = CStr(f)
    t1.Visible = True
    t1.SetFocus
   Else
    lr1.Move r.Right - 4, r.Top + 1, 4, r.Bottom - r.Top - 1
    lr1.Visible = True
   End If
  Case "boolean"
   If bDblClick Then
    With m_Data(0).d(.nDataIndex)
     .d(nSubPropIndex) = (.d(nSubPropIndex) = 0) And 1&
    End With
    b = True
   End If
  Case "string"
   If bDblClick Then
    t1.Move r.Left + 1, r.Top + 1, r.Right - r.Left - 1, r.Bottom - r.Top - 1
    t1.Text = m_Data(0).d(.nDataIndex).d
    t1.Visible = True
    t1.SetFocus
   End If
  End Select
  '////////////////////////////////
 Else
  Set objLast = obj
  b = obj.EditBegin(Me, hwnd, r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top, _
  m_Data(0).d(.nDataIndex).d, VarPtr(m_Def(0).d(.nPropDefIndex)), lp, nSubPropIndex, t1, cmb1, lr1, Nothing, bDblClick)
 End If
 t1.Tag = t1.Text
 cmb1.Tag = CStr(cmb1.ListIndex)
 'reset dirty flag
 bChanging = False
 'changed?
 If b Then
  lp = .nDataIndex
  n = .nPropDefIndex
  With m_Data(0).d(lp)
   Err.Clear
   .nSize = UBound(.d) + 1
   If Err.Number Then .nSize = 0
   RaiseEvent Change(lp, n, .sKey, .d, .nSize)
  End With
 End If
End With
End Sub

Private Sub pEditEnd()
t1.Visible = False
cmb1.Visible = False
lr1.Visible = False
If Not objLast Is Nothing Then
 objLast.EditEnd Me
 Set objLast = Nothing
End If
DoEvents '???
End Sub

Private Sub pChange(ByVal idx As Long, ByVal nSubPropIndex As Long, ByVal sText As String)
On Error Resume Next
Dim obj As IOperatorPropCallback
Dim b As Boolean, lp As Long, sKey As String
Dim n As Long, f As Single, f2 As Single
With ps(idx)
 lp = m_Def(0).d(.nPropDefIndex).nType
 With g_PropTypeDefs(lp)
  Set obj = .objCallback
  sKey = .sKey
 End With
 lp = VarPtr(g_PropTypeDefs(lp))
 'check callback
 If obj Is Nothing Then
  '////////////////////////////////default process
  Select Case sKey
  Case "int"
   n = Val(sText)
   'check enum
   With m_Def(0).d(.nPropDefIndex)
    If nSubPropIndex < .nEnumCount Then
     n = n + 1
     With .datEnum(nSubPropIndex)
      b = n > 0 And n <= .nCount
      If b Then n = .d(n).nValue
     End With
    Else
     b = True
    End If
   End With
   If b Then
    'clamp
    With m_Def(0).d(.nPropDefIndex)
     Err.Clear
     lp = UBound(.datMin) + 1
     If Err.Number Then lp = 0
     If lp >= nSubPropIndex * 4& + 4& Then
      CopyMemory lp, .datMin(nSubPropIndex * 4&), 4&
      If n < lp Then n = lp
     ElseIf lp = 4& Then
      CopyMemory lp, .datMin(0), 4&
      If n < lp Then n = lp
     End If
     Err.Clear
     lp = UBound(.datMax) + 1
     If Err.Number Then lp = 0
     If lp >= nSubPropIndex * 4& + 4& Then
      CopyMemory lp, .datMax(nSubPropIndex * 4&), 4&
      If n > lp Then n = lp
     ElseIf lp = 4& Then
      CopyMemory lp, .datMax(0), 4&
      If n > lp Then n = lp
     End If
    End With
    CopyMemory m_Data(0).d(.nDataIndex).d(nSubPropIndex * 4&), n, 4&
   End If
  Case "size", "resize"
   m_Data(0).d(.nDataIndex).d(nSubPropIndex) = Val(sText)
   b = True
  Case "float", "color"
   f = Val(sText)
   If sKey = "color" Then
    n = (nSubPropIndex * 4& + 12&) And &HF&
   Else
    n = nSubPropIndex * 4&
   End If
   'clamp
   With m_Def(0).d(.nPropDefIndex)
    Err.Clear
    lp = UBound(.datMin) + 1
    If Err.Number Then lp = 0
    If lp >= n + 4& Then
     CopyMemory f2, .datMin(n), 4&
     If f < f2 Then f = f2
    ElseIf lp = 4& Then
     CopyMemory f2, .datMin(0), 4&
     If f < f2 Then f = f2
    End If
    Err.Clear
    lp = UBound(.datMax) + 1
    If Err.Number Then lp = 0
    If lp >= n + 4& Then
     CopyMemory f2, .datMax(n), 4&
     If f > f2 Then f = f2
    ElseIf lp = 4& Then
     CopyMemory f2, .datMax(0), 4&
     If f > f2 Then f = f2
    End If
   End With
   CopyMemory m_Data(0).d(.nDataIndex).d(n), f, 4&
   b = True
  Case "string"
   m_Data(0).d(.nDataIndex).d = sText
   b = True
  End Select
  '////////////////////////////////
 Else
  b = obj.OnChange(Me, _
  m_Data(0).d(.nDataIndex).d, VarPtr(m_Def(0).d(.nPropDefIndex)), lp, nSubPropIndex, sText)
 End If
 'changed?
 If b Then
  lp = .nDataIndex
  n = .nPropDefIndex
  With m_Data(0).d(lp)
   Err.Clear
   .nSize = UBound(.d) + 1
   If Err.Number Then .nSize = 0
   RaiseEvent Change(lp, n, .sKey, .d, .nSize)
  End With
  pRedrawOne idx
  UserControl_Paint
 End If
End With
End Sub

Private Sub pDelta(ByVal idx As Long, ByVal nSubPropIndex As Long, ByVal iDelta As Long, ByVal Shift As Long, ByRef bCancel As Boolean)
On Error Resume Next
Dim obj As IOperatorPropCallback
Dim b As Boolean, lp As Long, sKey As String
Dim i As Long, j As Long, n As Long
Dim f As Single, f2 As Single, f3 As Single
Dim nMin As Long, nMax As Long
With ps(idx)
 lp = m_Def(0).d(.nPropDefIndex).nType
 With g_PropTypeDefs(lp)
  Set obj = .objCallback
  sKey = .sKey
 End With
 lp = VarPtr(g_PropTypeDefs(lp))
 If Shift And vbShiftMask Then nSubPropIndex = -1
 'check callback
 If obj Is Nothing Then
  '////////////////////////////////default process
  Select Case sKey
  Case "int"
   bCancel = True
   lp = VarPtr(m_Data(0).d(.nDataIndex).d(0))
   With m_Def(0).d(.nPropDefIndex)
    For i = 0 To .nElementCount - 1
     If nSubPropIndex = i Or nSubPropIndex < 0 Then
      'get min/max
      nMin = &H80000000
      nMax = &H7FFFFFFF
      Err.Clear
      n = UBound(.datMin) + 1
      If Err.Number Then n = 0
      If n >= i * 4& + 4& Then
       CopyMemory nMin, .datMin(i * 4&), 4&
      ElseIf n = 4& Then
       CopyMemory nMin, .datMin(0), 4&
      End If
      Err.Clear
      n = UBound(.datMax) + 1
      If Err.Number Then n = 0
      If n >= i * 4& + 4& Then
       CopyMemory nMax, .datMax(i * 4&), 4&
      ElseIf n = 4& Then
       CopyMemory nMax, .datMax(0), 4&
      End If
      'auto determine ??
      j = iDelta
      If CDbl(nMax) - CDbl(nMin) < 32 Then j = j \ 8&
      'change
      If j Then
       bCancel = False
       CopyMemory n, ByVal lp + i * 4&, 4&
       j = j + n
       If j < nMin Then j = nMin
       If j > nMax Then j = nMax
       If j <> n Then
        CopyMemory ByVal lp + i * 4&, j, 4&
        b = True
       End If
      End If
     End If
    Next i
   End With
  Case "float", "color"
   lp = VarPtr(m_Data(0).d(.nDataIndex).d(0))
   If sKey = "color" Then
    iDelta = iDelta * 1000&
    If nSubPropIndex >= 0 Then nSubPropIndex = (nSubPropIndex + 3&) And 3&
   End If
   With m_Def(0).d(.nPropDefIndex)
    For i = 0 To .nElementCount - 1
     If nSubPropIndex = i Or nSubPropIndex < 0 Then
      CopyMemory f, ByVal lp + i * 4&, 4&
      f3 = f
      f = f + iDelta / 1000
      'clamp
      Err.Clear
      n = UBound(.datMin) + 1
      If Err.Number Then n = 0
      If n >= i * 4& + 4& Then
       CopyMemory f2, .datMin(i * 4&), 4&
       If f < f2 Then f = f2
      ElseIf n = 4& Then
       CopyMemory f2, .datMin(0), 4&
       If f < f2 Then f = f2
      End If
      Err.Clear
      n = UBound(.datMax) + 1
      If Err.Number Then n = 0
      If n >= i * 4& + 4& Then
       CopyMemory f2, .datMax(i * 4&), 4&
       If f > f2 Then f = f2
      ElseIf n = 4& Then
       CopyMemory f2, .datMax(0), 4&
       If f > f2 Then f = f2
      End If
      'change
      If f <> f3 Then
       b = True
       CopyMemory ByVal lp + i * 4&, f, 4&
      End If
     End If
    Next i
   End With
  End Select
  '////////////////////////////////
 Else
  b = obj.OnDelta(Me, _
  m_Data(0).d(.nDataIndex).d, VarPtr(m_Def(0).d(.nPropDefIndex)), lp, nSubPropIndex, iDelta, bCancel)
 End If
 'changed?
 If b Then
  lp = .nDataIndex
  n = .nPropDefIndex
  With m_Data(0).d(lp)
   Err.Clear
   .nSize = UBound(.d) + 1
   If Err.Number Then .nSize = 0
   RaiseEvent Change(lp, n, .sKey, .d, .nSize)
  End With
  pRedrawOne idx
  UserControl_Paint
 End If
End With
End Sub

Private Sub pCustom(ByVal idx As Long)
On Error Resume Next
Dim obj As IOperatorPropCallback
Dim b As Boolean, lp As Long, i As Long, sKey As String
With ps(idx)
 lp = m_Def(0).d(.nPropDefIndex).nType
 With g_PropTypeDefs(lp)
  Set obj = .objCallback
  b = .nFlags And 1&
  sKey = .sKey
 End With
 lp = VarPtr(g_PropTypeDefs(lp))
 If b Then
  b = False
  If Not obj Is Nothing Then
   Set objLast = obj
   b = obj.EditBegin(Me, 0, 0, 0, 0, 0, _
   m_Data(0).d(.nDataIndex).d, VarPtr(m_Def(0).d(.nPropDefIndex)), lp, -1, Nothing, Nothing, Nothing, Nothing, True)
  End If
  If Not b Then
   '////////////////////////////////default process
   Select Case sKey
   Case "color"
    b = pCustomColor(m_Data(0).d(.nDataIndex).d)
   Case "string"
    b = pCustomString(m_Data(0).d(.nDataIndex).d)
   End Select
   '////////////////////////////////
  End If
  'changed?
  If b Then
   lp = .nDataIndex
   i = .nPropDefIndex
   With m_Data(0).d(lp)
    Err.Clear
    .nSize = UBound(.d) + 1
    If Err.Number Then .nSize = 0
    RaiseEvent Change(lp, i, .sKey, .d, .nSize)
   End With
  End If
 End If
End With
End Sub

Private Function pCustomColor(b() As Byte) As Boolean
On Error Resume Next
Dim rgbRed As Single, rgbGreen As Single, rgbBlue As Single, rgbReserved As Single
Dim frm As New frmColorPicker
CopyMemory rgbRed, b(0), 4&
CopyMemory rgbGreen, b(4), 4&
CopyMemory rgbBlue, b(8), 4&
CopyMemory rgbReserved, b(12), 4&
With frm
 .SetColor rgbRed, rgbGreen, rgbBlue, rgbReserved
 .Show 1
 If .Changed Then
  .GetColor rgbRed, rgbGreen, rgbBlue, rgbReserved
  CopyMemory b(0), rgbRed, 4&
  CopyMemory b(4), rgbGreen, 4&
  CopyMemory b(8), rgbBlue, 4&
  CopyMemory b(12), rgbReserved, 4&
  pCustomColor = True
 End If
End With
Unload frm
End Function

Private Function pCustomString(b() As Byte) As Boolean
Dim s As String
Dim frm As New frmStr
Dim i As Long
s = b
Load frm
With frm.txtStr
 .Visible = True
 .Text = s
End With
frm.oString = s
frm.Show 1
If s <> frm.oString Then
 b = frm.oString
 pCustomString = True
End If
Unload frm
End Function

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
Dim idx As Long
If pc = 0 Then Exit Sub
'splitter
If idxHl = &HFF000000 And (Button And 1) Then
 i = nDelta + x
 If i > ScaleWidth - 64 Then i = ScaleWidth - 64
 If i < 32 Then i = 32
 UserControl_Paint
 bm.PaintPicture hdc, i - 2, 0, 5, , , , vbDstInvert
 MousePointer = vbSizeWE
 Exit Sub
End If
'hit test
For i = 1 To pc
 idx = pHitTestOne(i, x, y)
 If idx Then Exit For
Next i
'splitter
If idx = &HFF000000 Then
 MousePointer = vbSizeWE
Else
 MousePointer = vbDefault
End If
'check redraw
If idx <> idxHl Then
 tmr1.Enabled = idx
 i = idxHl And &HFFFFFF
 idxHl = idx
 idx = idx And &HFFFFFF
 If i > 0 Then pRedrawOne i
 If idx > 0 And idx <> i Then pRedrawOne idx
 UserControl_Paint
End If
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long
'Dim idx As Long
If pc = 0 Then Exit Sub
'splitter
If idxHl = &HFF000000 And Button = 1 Then
 i = nDelta + x
 If i > ScaleWidth - 64 Then i = ScaleWidth - 64
 If i < 32 Then i = 32
 If i <> nCaptionWidth Then
  nCaptionWidth = i
  pRedraw
 Else
  UserControl_Paint
 End If
 Exit Sub
End If
'TODO:
End Sub

Private Sub UserControl_Paint()
bm.PaintPicture hdc
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
pInit
End Sub

Private Sub pInit()
bm0.CreateFromPicture i0.Picture
cFnt.HighQuality = True
Set cFnt.LogFont = Font
SendMessage t1.hwnd, WM_SETFONT, cFnt.Handle, ByVal 0
nCaptionWidth = 64
sb1.NCPaintColor1 = d_CtrlBorder
pRefresh
End Sub

Public Sub SetMenu(obj As FakeMenu)
Set objMenu = obj
If Not obj Is Nothing Then
 If obj.FindMenu("____reset_prop") = 0 Then
  obj.AddMenuFromString "____reset_prop", ";&Reset,;Reset &All"
 End If
End If
End Sub

Public Sub SetData(ByVal lpOperatorProperties As Long, ByVal lpOperatorPropDefs As Long)
idxHl = 0
idxSelected = 0
pEditEnd
'///
ZeroMemory ByVal VarPtrArray(m_Data), 4&
m_lpData = lpOperatorProperties
If lpOperatorProperties <> 0 Then
 With m_tData
  .cDims = 1
  .cbElements = 0
  .pvData = lpOperatorProperties
  .Bounds(0).cElements = 1
 End With
 CopyMemory ByVal VarPtrArray(m_Data), VarPtr(m_tData), 4&
End If
ZeroMemory ByVal VarPtrArray(m_Def), 4&
m_lpDef = lpOperatorPropDefs
If lpOperatorPropDefs <> 0 Then
 With m_tDef
  .cDims = 1
  .cbElements = 0
  .pvData = lpOperatorPropDefs
  .Bounds(0).cElements = 1
 End With
 CopyMemory ByVal VarPtrArray(m_Def), VarPtr(m_tDef), 4&
End If
pLoadData
pRefresh
End Sub

Private Sub pLoadData()
Dim i As Long, j As Long, k As Long, l As Long
Dim s As String
Dim b As Boolean
Erase ps
pc = 0
If m_lpData = 0 Or m_lpDef = 0 Then Exit Sub
With m_Def(0)
 For i = 1 To .nCount
  k = .d(i).nType
  If k >= &H80000000 And k <= &H800000FE Then 'group
   l = k And &HFF&
   pc = pc + 1
   ReDim Preserve ps(1 To pc)
   With ps(pc)
    .nDataIndex = 0
    .nPropDefIndex = i
    .nFlags = l Or &H200&
    .h = 16&
   End With
   l = l + 1
  ElseIf k > 0 Then
   s = .d(i).sKey
   b = False
   With m_Data(0)
    For j = 1 To .nCount
     If .d(j).sKey = s Then
      '///????????
      .d(j).nIndex = i
      '///
      b = True
      Exit For
     End If
    Next j
   End With
   If b Then
    k = .d(i).nElementCount
    pc = pc + 1
    ReDim Preserve ps(1 To pc)
    With ps(pc)
     .nDataIndex = j
     .nPropDefIndex = i
     .nFlags = l
     .h = 16&
     If k > 4 Then .h = ((k + 3) And &HFFFFFFFC) * 4&
    End With
   End If
  End If
 Next i
End With
End Sub

Private Sub pRefresh()
Dim i As Long
'calc size
m_Height = 0
For i = 1 To pc
 With ps(i)
  .y = m_Height
  If (.nFlags And &H100&) = 0 Then m_Height = m_Height + .h
 End With
Next i
m_Height = m_Height + 1
'update scrollbar
i = ScaleHeight
With sb1
 If m_Height > i Then
  .Enabled(efsVertical) = True
  .Visible(efsVertical) = True
  .Max(efsVertical) = m_Height - i
  .SmallChange(efsVertical) = 16&
  .LargeChange(efsVertical) = i
 Else
  .Enabled(efsVertical) = False
  .Visible(efsVertical) = False
  .Value(efsVertical) = 0
 End If
End With
'check caption width
If nCaptionWidth > ScaleWidth - 64 Then nCaptionWidth = ScaleWidth - 64
If nCaptionWidth < 32 Then nCaptionWidth = 32
'resize bitmap
bm.Create ScaleWidth, i
'redraw
pRedraw
End Sub

Private Sub pRedraw()
Dim i As Long
pEditEnd '???
'draw back
GradientFillRect bm.hdc, 0, 0, bm.Width, bm.Height, d_Bar2, d_Bar1, GRADIENT_FILL_RECT_H
nLastY = sb1.Visible(efsVertical) And sb1.Value(efsVertical)  ':-3
'draw item (???)
For i = 1 To pc
 pRedrawOne i
Next i
'TODO:draw other
UserControl_Paint
End Sub

Private Sub pRedrawOne(ByVal idx As Long)
Dim obj As IOperatorPropCallback
Dim sKey As String, s As String
Dim i As Long, j As Long
Dim m As Long, n As Long, f As Single
Dim r As RECT, hbr As Long, hbrBorder As Long
Dim nTop As Long
Dim w As Long, h As Long
w = bm.Width
h = bm.Height
With ps(idx)
 nTop = .y - nLastY
 If (.nFlags And &H100&) = 0 And nTop < h And nTop + .h >= 0 Then 'visible
  hbrBorder = CreateSolidBrush(d_Border)
  If .nFlags And &H200& Then
   'draw group
   r.Bottom = nTop + .h
   If (idxHl And &HFFFFFF) = idx Then
    GradientFillRect bm.hdc, 0, nTop, w, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
   Else
    GradientFillRect bm.hdc, 0, nTop, w, r.Bottom, d_Title1, d_Title2, GRADIENT_FILL_RECT_V
   End If
   'TODO:indent?
   i = (.nFlags And &HFF&) * 16&
   TransparentBlt bm.hdc, i + 4, nTop + 4, 9, 9, bm0.hdc, ((.nFlags And &H400&) = 0) And 9, 0, 9, 9, vbGreen
   i = i + 16&
   cFnt.DrawTextXP bm.hdc, m_Def(0).d(.nPropDefIndex).sCaption, _
   i, nTop, w - i, .h, DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX Or DT_END_ELLIPSIS, vbWhite, , True
   r.Left = -1
   r.Top = nTop
   r.Right = w + 1
   r.Bottom = r.Bottom + 1
   FrameRect bm.hdc, r, hbrBorder
  Else
   'draw caption
   r.Right = nCaptionWidth
   r.Bottom = nTop + .h
   If idxHl = (&HFE000000 Or idx) Then
    If (idxSelected And &HFFFFFF) = idx Then
     GradientFillRect bm.hdc, 0, nTop, r.Right, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
    Else
     GradientFillRect bm.hdc, 0, nTop, r.Right, r.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
    End If
   Else
    If (idxSelected And &HFFFFFF) = idx Then
     GradientFillRect bm.hdc, 0, nTop, r.Right, r.Bottom, d_Checked1, d_Checked2, GRADIENT_FILL_RECT_V
    Else
     GradientFillRect bm.hdc, 0, nTop, r.Right, r.Bottom, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_V
    End If
   End If
   cFnt.DrawTextXP bm.hdc, m_Def(0).d(.nPropDefIndex).sCaption, _
   4, nTop, r.Right - 4, .h, DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX Or DT_END_ELLIPSIS, vbBlack, , True
   r.Left = -1
   r.Top = nTop
   r.Right = r.Right + 1
   r.Bottom = r.Bottom + 1
   FrameRect bm.hdc, r, hbrBorder
   'draw custom
   i = m_Def(0).d(.nPropDefIndex).nType
   h = VarPtr(g_PropTypeDefs(i))
   With g_PropTypeDefs(i)
    Set obj = .objCallback
    sKey = .sKey
    n = .nFlags
   End With
   If n And 1& Then
    r.Right = w
    w = w - 18
    r.Left = w
    If idxHl = (&HFF000000 Or idx) Then
     GradientFillRect bm.hdc, r.Left, nTop, r.Right, r.Bottom - 1, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
    Else
     GradientFillRect bm.hdc, r.Left, nTop, r.Right, r.Bottom - 1, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_V
    End If
    TransparentBlt bm.hdc, r.Left + 2, nTop + .h \ 2 - 6, 14, 13, bm0.hdc, 0, 16, 14, 13, vbGreen
    FrameRect bm.hdc, r, hbrBorder
   End If
   'draw sub item
   m = m_Def(0).d(.nPropDefIndex).nElementCount
   If m = 0 Then m = 1
   For i = 0 To m - 1
    r.Top = nTop + (i And &HFFFFFFFC) * 4&
    r.Bottom = r.Top + 16
    If i < (m And &HFFFFFFFC) Then n = 4 Else n = m And 3&
    r.Left = nCaptionWidth + ((w - nCaptionWidth) * (i And 3&)) \ n
    r.Right = nCaptionWidth + ((w - nCaptionWidth) * ((i And 3&) + 1&)) \ n
    n = 0
    j = (i * &H1000000) Or idx
    j = ((idxHl = j) And 1&) Or ((idxSelected = j) And 2&) ' Or idxSelected = (&HFE000000 Or idx)
    If Not obj Is Nothing Then
     n = obj.Draw(Me, bm.hdc, r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top, _
     m_Data(0).d(.nDataIndex).d, VarPtr(m_Def(0).d(.nPropDefIndex)), h, i, j)
    End If
    If n = 0 Then
     '////////////////////////////////default process
     'back
     If j And 1& Then
      If j And 2& Then
       GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
      Else
       GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
      End If
     Else
      If j And 2& Then
       GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Checked1, d_Checked2, GRADIENT_FILL_RECT_V
      Else
       If sKey = "color" Then
        n = pGetColor(m_Data(0).d(.nDataIndex).d)
        hbr = CreateSolidBrush(n)
        n = n Xor &H808080
       Else
        hbr = CreateSolidBrush(vbWhite)
       End If
       FillRect bm.hdc, r, hbr
       DeleteObject hbr
      End If
     End If
     'border
     r.Right = r.Right + 1
     r.Bottom = r.Bottom + 1
     FrameRect bm.hdc, r, hbrBorder
     r.Right = r.Right - 1
     r.Bottom = r.Bottom - 1
     'caption
     j = r.Right - r.Left - 4
     Select Case sKey
     Case "int"
      CopyMemory n, m_Data(0).d(.nDataIndex).d(i * 4&), 4&
      'check enum
      If pGetEnum(n, m_Def(0).d(.nPropDefIndex), i, s) Then
       cFnt.DrawTextXP bm.hdc, s, r.Left + 4, r.Top, j, r.Bottom - r.Top, _
       DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX Or DT_END_ELLIPSIS, vbBlack, , True
      Else
       cFnt.DrawTextXP bm.hdc, CStr(n), r.Left + 4, r.Top, j, r.Bottom - r.Top, _
       DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS, vbBlack, , True
      End If
     Case "float"
      CopyMemory f, m_Data(0).d(.nDataIndex).d(i * 4&), 4&
      cFnt.DrawTextXP bm.hdc, Format(f, "0.000"), r.Left + 4, r.Top, j, r.Bottom - r.Top, _
      DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS, vbBlack, , True
     Case "color"
      CopyMemory f, m_Data(0).d(.nDataIndex).d((i * 4& + 12&) And &HF&), 4&
      cFnt.DrawTextXP bm.hdc, Format(f, "0.0"), r.Left + 4, r.Top, j, r.Bottom - r.Top, _
      DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS, n, , True
     Case "size", "resize"
      n = m_Data(0).d(.nDataIndex).d(i)
      If sKey = "resize" Then n = n - 1
      If n < 0 Then
       s = "(Current)"
      ElseIf n > int_Size_Max Then
       s = "(2^" + CStr(n) + ")"
      Else
       s = CStr(nBitFieldMask1(n))
      End If
      cFnt.DrawTextXP bm.hdc, s, r.Left + 4, r.Top, j, r.Bottom - r.Top, _
      DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS, n, , True
     Case "boolean"
      hbr = 1
      'check enum
      With m_Def(0).d(.nPropDefIndex)
       If i < .nEnumCount Then
        With .datEnum(i)
         If .nCount > 0 Then
          s = .d(1).sCaption
          hbr = 0
         End If
        End With
       End If
      End With
      If hbr Then
       If m_Data(0).d(.nDataIndex).d(i) Then s = "True" Else s = "False"
       cFnt.DrawTextXP bm.hdc, s, r.Left + 4, r.Top, j, r.Bottom - r.Top, _
       DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX Or DT_END_ELLIPSIS, vbBlack, , True
      Else
       cFnt.DrawTextXP bm.hdc, s, r.Left + 16, r.Top, j - 12, r.Bottom - r.Top, _
       DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX Or DT_END_ELLIPSIS, vbBlack, , True
       'checkbox
       bm0.PaintPicture bm.hdc, r.Left + 3, r.Top + 3, 11, 11, 21, _
       10 + ((m_Data(0).d(.nDataIndex).d(i) <> 0) And 11&), vbSrcCopy
       'border
       r.Left = r.Left + 2
       r.Top = r.Top + 2
       r.Right = r.Left + 13
       r.Bottom = r.Top + 13
       hbr = CreateSolidBrush(TheBorderColor)
       FrameRect bm.hdc, r, hbr
       DeleteObject hbr
      End If
     Case "string", "load"
      s = m_Data(0).d(.nDataIndex).d
      cFnt.DrawTextXP bm.hdc, s, r.Left + 4, r.Top, j, r.Bottom - r.Top, _
      DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX Or DT_END_ELLIPSIS, vbBlack, , True
     Case Else '????????
      s = "(" + sKey + ")"
      cFnt.DrawTextXP bm.hdc, s, r.Left + 4, r.Top, j, r.Bottom - r.Top, _
      DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX Or DT_END_ELLIPSIS, vbBlack, , True
     End Select
     '////////////////////////////////
    End If
   Next i
   'TODO:draw other
  End If
  DeleteObject hbrBorder
 End If
End With
End Sub

Private Function pHitTestOne(ByVal idx As Long, ByVal x As Long, ByVal y As Long) As Long
Dim i As Long, m As Long, n As Long
Dim r As RECT
Dim nTop As Long
Dim w As Long, h As Long
w = bm.Width
h = bm.Height
With ps(idx)
 nTop = .y - nLastY
 i = nTop + .h
 If (.nFlags And &H100&) = 0 And nTop < h And i >= 0 And y >= nTop And y <= i Then
  If .nFlags And &H200& Then
   pHitTestOne = idx And x >= 0 And x < w
  Else
   i = m_Def(0).d(.nPropDefIndex).nType
   i = g_PropTypeDefs(i).nFlags And 1&
   If i Then w = w - 18
   If x >= nCaptionWidth - 2 And x <= nCaptionWidth + 2 Then
    pHitTestOne = &HFF000000
   ElseIf x >= 0 And x < nCaptionWidth Then
    pHitTestOne = idx Or &HFE000000
   ElseIf i And x >= w And x < w + 18 Then
    pHitTestOne = idx Or &HFF000000
   Else
    m = m_Def(0).d(.nPropDefIndex).nElementCount
    If m = 0 Then m = 1
    For i = 0 To m - 1
     r.Top = nTop + (i And &HFFFFFFFC) * 4&
     r.Bottom = r.Top + 16
     If i < (m And &HFFFFFFFC) Then n = 4 Else n = m And 3&
     r.Left = nCaptionWidth + ((w - nCaptionWidth) * (i And 3&)) \ n
     r.Right = nCaptionWidth + ((w - nCaptionWidth) * ((i And 3&) + 1&)) \ n
     If x >= r.Left And x <= r.Right And y >= r.Top And y <= r.Bottom Then
      pHitTestOne = idx Or (i * &H1000000)
      Exit For
     End If
    Next i
   End If
  End If
 End If
End With
End Function

Private Function pGetEnum(ByVal n As Long, d As typeMyTGL11OperatorPropDef, ByVal i As Long, s As String) As Boolean
Dim j As Long
If i >= 0 And i < d.nEnumCount Then
 With d.datEnum(i)
  For j = 1 To .nCount
   If n = .d(j).nValue Then
    s = .d(j).sCaption
    pGetEnum = True
    Exit For
   End If
  Next j
 End With
 If Not pGetEnum Then
  s = "(" + CStr(n) + ")"
  pGetEnum = True
 End If
End If
End Function

Private Function pGetColor(b() As Byte) As Long
On Error Resume Next
Dim rgbRed As Single, rgbGreen As Single, rgbBlue As Single
CopyMemory rgbRed, b(0), 4&
CopyMemory rgbGreen, b(4), 4&
CopyMemory rgbBlue, b(8), 4&
If rgbRed < 0 Then rgbRed = 0 Else If rgbRed > 255 Then rgbRed = 255
If rgbGreen < 0 Then rgbGreen = 0 Else If rgbGreen > 255 Then rgbGreen = 255
If rgbBlue < 0 Then rgbBlue = 0 Else If rgbBlue > 255 Then rgbBlue = 255
pGetColor = (rgbRed And &HFF&) Or _
((rgbGreen And &HFF&) * &H100&) Or _
((rgbBlue And &HFF&) * &H10000)
End Function

Private Sub UserControl_Resize()
pRefresh
End Sub

Private Sub UserControl_Terminate()
Set objLast = Nothing
SetData 0, 0
End Sub

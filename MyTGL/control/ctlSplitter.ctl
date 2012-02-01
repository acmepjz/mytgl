VERSION 5.00
Begin VB.UserControl ctlSplitter 
   BackColor       =   &H8000000C&
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
   InvisibleAtRuntime=   -1  'True
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   Windowless      =   -1  'True
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "vbAccelerator Splitter Control"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "ctlSplitter"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////
'This file is under vbAccelerator Software License,
'based on the Apache Software Foundation Software Licence.
'See <http://www.vbaccelerator.com/home/The_Site/Usage_Policy/article.asp>.
'////////////////////////////////

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
Private Type BITMAP '24 bytes
   bmType As Long
   bmWidth As Long
   bmHeight As Long
   bmWidthBytes As Long
   bmPlanes As Integer
   bmBitsPixel As Integer
   bmBits As Long
End Type

Private Const IDC_SIZENS = 32645&
Private Const IDC_SIZEWE = 32644&
Private Const IDC_NO = 32648&

Private Const R2_NOTXORPEN = 10  '  DPxn

Private Const PATINVERT = &H5A0049       ' (DWORD) dest = pattern XOR dest
Private Const DSTINVERT = &H550009       ' (DWORD) dest = (NOT dest)

Private Declare Function GetClientRect Lib "user32" (ByVal hwnd As Long, lpRect As RECT) As Long
Private Declare Function ScreenToClient Lib "user32" (ByVal hwnd As Long, lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32" (lpPoint As POINTAPI) As Long
Private Declare Sub ClipCursorRect Lib "user32" Alias "ClipCursor" (lpRect As RECT)
Private Declare Sub ClipCursorClear Lib "user32" Alias "ClipCursor" (ByVal lpRect As Long)
Private Declare Function DeleteDC Lib "gdi32" (ByVal hdc As Long) As Long
Private Declare Function CreateDCAsNull Lib "gdi32" Alias "CreateDCA" (ByVal lpDriverName As String, lpDeviceName As Any, lpOutput As Any, lpInitData As Any) As Long
Private Declare Function SetROP2 Lib "gdi32" (ByVal hdc As Long, ByVal nDrawMode As Long) As Long
Private Declare Function CreateBitmapIndirect Lib "gdi32" (lpBitmap As BITMAP) As Long
Private Declare Function CreatePatternBrush Lib "gdi32" (ByVal hBitmap As Long) As Long
Private Declare Function SelectObject Lib "gdi32" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32" (ByVal hdc As Long, lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function PatBlt Lib "gdi32" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function LoadCursorLong Lib "user32" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function DestroyCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function SetCursor Lib "user32" (ByVal hCursor As Long) As Long
Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function OffsetRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long

Private Const WM_LBUTTONDOWN = &H201
Private Const WM_LBUTTONUP = &H202
Private Const WM_MOUSEMOVE = &H200
Private Const WM_SIZE = &H5

Public Enum ESPLTOrientationConstants
    cSPLTOrientationHorizontal = 1
    cSPLTOrientationVertical = 2
End Enum

Public Enum ESPLTPanelConstants
   cSPLTLeftOrTopPanel = 1
   cSPLTRightOrBottomPanel = 2
End Enum

Private m_bKeepProportionsWhenResizing As Boolean
Private m_fProportion As Single
Private m_lSplitPos As Long
Private m_lSplitSize As Long
Private m_lMinSize(1 To 2) As Long
Private m_lMaxSize(1 To 2) As Long
Private m_bFullDrag As Boolean
Private m_bInDrag As Boolean
Private m_tPInitial As POINTAPI
Private m_lSplitInitial  As Long
Private m_hBrush As Long
Private m_lPattern(0 To 3) As Long
Private m_tSplitR As RECT
Private m_hCursor As Long

Private m_oObjects(1 To 2) As typeFakeSplitObjects

Private m_eOrientation As ESPLTOrientationConstants

Public Event Split(ByVal x As Long, ByVal y As Long, ByRef bCancel As Boolean)

Implements iSubclass
Private cSub As New cSubclass

'////////////////new
Private Type typeFakeSplitSubObject
 obj As Object
 nDir As Long
End Type
Private Type typeFakeSplitSubObjects
 nCount As Long
 d() As typeFakeSplitSubObject
End Type
Private Type typeFakeSplitObjects
 nCount As Long
 d() As Object
End Type
Private m_oSubObj(0 To 2) As typeFakeSplitSubObjects
Private bEnabled As Boolean
'////////////////

Public Property Get Enabled() As Boolean
Enabled = bEnabled
End Property

Public Property Let Enabled(ByVal b As Boolean)
bEnabled = b
End Property

Public Sub SetSubObject(ByVal ePanel As ESPLTPanelConstants, Optional ByVal Objects As Variant, Optional ByVal Directions As Variant)
Dim i As Long, j As Long, k As Long, m As Long
If IsMissing(Objects) Or IsMissing(Directions) Then
 With m_oSubObj(ePanel)
  Erase .d
  .nCount = 0
 End With
ElseIf VarType(Objects) And vbArray Then
 If VarType(Directions) And vbArray Then
  With m_oSubObj(ePanel)
   j = LBound(Objects) - 1
   k = LBound(Directions) - 1
   m = UBound(Objects) - j
   ReDim .d(1 To m)
   For i = 1 To m
    With .d(i)
     Set .obj = Objects(j + i)
     .nDir = Directions(k + i)
    End With
   Next i
   .nCount = m
  End With
 Else
  With m_oSubObj(ePanel)
   j = LBound(Objects) - 1
   m = UBound(Objects) - j
   ReDim .d(1 To m)
   For i = 1 To m
    With .d(i)
     Set .obj = Objects(j + i)
     .nDir = Directions
    End With
   Next i
   .nCount = m
  End With
 End If
Else
 With m_oSubObj(ePanel)
  ReDim .d(1 To 1)
  With .d(1)
   Set .obj = Objects
   .nDir = Directions
  End With
  .nCount = 1
 End With
End If
Resize
End Sub

Private Sub MyGetWindowRect(ByVal hwnd As Long, lpRect As RECT)
Dim p As POINTAPI
GetClientRect hwnd, lpRect
ClientToScreen hwnd, p
OffsetRect lpRect, p.x, p.y
End Sub

Public Property Get FullDrag() As Boolean
   FullDrag = m_bFullDrag
End Property
Public Property Let FullDrag(ByVal bState As Boolean)
   If Not (m_bFullDrag = bState) Then
      m_bFullDrag = bState
      If Not m_bFullDrag Then
         CreateBrush
      Else
         DestroyBrush
      End If
   End If
End Property

Public Property Get Orientation() As ESPLTOrientationConstants
   Orientation = m_eOrientation
End Property
Public Property Let Orientation(ByVal eOrientation As ESPLTOrientationConstants)
   If Not (m_eOrientation = eOrientation) Then
      m_eOrientation = eOrientation
      If Not (m_hCursor = 0) Then
         DestroyCursor m_hCursor
      End If
      If (m_eOrientation = cSPLTOrientationHorizontal) Then
         m_hCursor = LoadCursorLong(0, IDC_SIZENS)
      Else
         m_hCursor = LoadCursorLong(0, IDC_SIZEWE)
      End If
      Resize
   End If
End Property

Public Property Get Proportion() As Single
Attribute Proportion.VB_MemberFlags = "400"
   If (m_fProportion > 1) Then
      m_fProportion = 1
   End If
   Proportion = m_fProportion * 100
End Property
Public Property Let Proportion(ByVal fProportion As Single)
   If (fProportion > 100#) Or (fProportion < 0#) Then
      Err.Raise 380, App.EXEName & ".cSplitter"
   Else
      m_fProportion = fProportion / 100#
      Resize
   End If
End Property

Public Property Get Position() As Long
Attribute Position.VB_MemberFlags = "400"
   Position = m_lSplitPos
End Property
Public Property Let Position(ByVal lPosition As Long)
   If (lPosition <> m_lSplitPos) Then
      m_lSplitPos = lPosition
      pValidatePosition
      pSetProportion
      Resize
   End If
End Property

Public Property Get KeepProportion() As Boolean
   KeepProportion = m_bKeepProportionsWhenResizing
End Property
Public Property Let KeepProportion(ByVal bState As Boolean)
   m_bKeepProportionsWhenResizing = bState
End Property

Private Sub LetContainer()
   
On Error Resume Next

   If Not Ambient.UserMode Then Exit Sub
cSub.AddMsg WM_LBUTTONDOWN, MSG_AFTER
cSub.AddMsg WM_LBUTTONUP, MSG_AFTER
cSub.AddMsg WM_MOUSEMOVE, MSG_AFTER
cSub.AddMsg WM_SIZE, MSG_AFTER
cSub.Subclass UserControl.ContainerHwnd, Me
End Sub

Public Property Get SplitterSize() As Long
   SplitterSize = m_lSplitSize
End Property

Public Property Let SplitterSize(ByVal lSize As Long)
   If Not (m_lSplitSize = lSize) Then
      If (lSize < 0) Then
         Err.Raise 380, App.EXEName & ".cSplitter"
      Else
         m_lSplitSize = lSize
         Resize
      End If
   End If
End Property

Public Property Get MinimumSize( _
      ByVal ePanel As ESPLTPanelConstants _
   ) As Long
   MinimumSize = m_lMinSize(ePanel)
End Property

Public Property Let MinimumSize( _
      ByVal ePanel As ESPLTPanelConstants, _
      ByVal lSize As Long _
   )
   If Not (m_lMinSize(ePanel) = lSize) Then
      m_lMinSize(ePanel) = lSize
      Resize
   End If
End Property

Public Property Get MaximumSize( _
      ByVal ePanel As ESPLTPanelConstants _
   ) As Long
   MaximumSize = m_lMaxSize(ePanel)
End Property

Public Property Let MaximumSize( _
      ByVal ePanel As ESPLTPanelConstants, _
      ByVal lSize As Long _
   )
   If Not (m_lMaxSize(ePanel) = lSize) Then
      m_lMaxSize(ePanel) = lSize
   End If
End Property

Public Sub Bind(Optional ByVal oLeftTop As Variant, Optional ByVal oRightBottom As Variant)
'modified!!!
Dim i As Long, j As Long, m As Long
If IsMissing(oLeftTop) Then
 With m_oObjects(1)
  Erase .d
  .nCount = 0
 End With
ElseIf VarType(oLeftTop) And vbArray Then
 With m_oObjects(1)
  j = LBound(oLeftTop) - 1
  m = UBound(oLeftTop) - j
  ReDim .d(1 To m)
  For i = 1 To m
   Set .d(i) = oLeftTop(j + i)
  Next i
  .nCount = m
 End With
Else
 With m_oObjects(1)
  ReDim .d(1 To 1)
  Set .d(1) = oLeftTop
  .nCount = 1
 End With
End If
If IsMissing(oRightBottom) Then
 With m_oObjects(2)
  Erase .d
  .nCount = 0
 End With
ElseIf VarType(oRightBottom) And vbArray Then
 With m_oObjects(2)
  j = LBound(oRightBottom) - 1
  m = UBound(oRightBottom) - j
  ReDim .d(1 To m)
  For i = 1 To m
   Set .d(i) = oRightBottom(j + i)
  Next i
  .nCount = m
 End With
Else
 With m_oObjects(2)
  ReDim .d(1 To 1)
  Set .d(1) = oRightBottom
  .nCount = 1
 End With
End If
'over
Resize
End Sub

Private Function pbConfigured() As Boolean
   'add!!
   If Not Ambient.UserMode Then Exit Function
   'end
'   If Not m_oContainer Is Nothing Then
      If m_oObjects(1).nCount > 0 Then
         If m_oObjects(2).nCount > 0 Then
            pbConfigured = True
         End If
      End If
'   End If
End Function

Private Function pHitTest(ByVal x As Long, ByVal y As Long) As Boolean
If (m_eOrientation = cSPLTOrientationVertical) Then
 pHitTest = x >= m_lSplitPos And x < m_lSplitPos + m_lSplitSize
Else
 pHitTest = y >= m_lSplitPos And y < m_lSplitPos + m_lSplitSize
End If
End Function

Private Sub MouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
   If (Button = vbLeftButton) Then
      Dim bCancel As Boolean
      If Not pHitTest(x, y) Then Exit Sub
      RaiseEvent Split(x, y, bCancel)
      If Not bCancel Then
         m_bInDrag = True
      
         Dim tP As POINTAPI
         GetCursorPos tP
         LSet m_tPInitial = tP
         m_lSplitInitial = m_lSplitPos
            
         Dim tR As RECT
         MyGetWindowRect ContainerHwnd, tR
         ClipCursorRect tR
         
         If Not (m_bFullDrag) Then
            If (m_eOrientation = cSPLTOrientationVertical) Then
               m_tSplitR.Left = tR.Left + m_lSplitPos
               m_tSplitR.Right = m_tSplitR.Left + m_lSplitSize
               m_tSplitR.Top = tR.Top
               m_tSplitR.Bottom = tR.Bottom
            Else
               m_tSplitR.Left = tR.Left
               m_tSplitR.Right = tR.Right
               m_tSplitR.Top = tR.Top + m_lSplitPos
               m_tSplitR.Bottom = m_tSplitR.Top + m_lSplitSize
            End If
            
            pDrawSplitter
            
         End If
         
      End If
   End If
End Sub

Private Sub MouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
      
   If (pbConfigured) Then
   
      If (m_bInDrag) Then
         SetCursor m_hCursor
         
         Dim tP As POINTAPI
         GetCursorPos tP
         
         If Not (m_bFullDrag) Then
            pDrawSplitter
         End If
         
         If (m_eOrientation = cSPLTOrientationVertical) Then
            m_lSplitPos = m_lSplitInitial + (tP.x - m_tPInitial.x)
         Else
            m_lSplitPos = m_lSplitInitial + (tP.y - m_tPInitial.y)
         End If
         pValidatePosition
         
         If (m_bFullDrag) Then
            pResizePanels
         Else
            Dim tR As RECT
            MyGetWindowRect ContainerHwnd, tR
            
            If (m_eOrientation = cSPLTOrientationVertical) Then
               m_tSplitR.Left = tR.Left + m_lSplitPos
               m_tSplitR.Right = m_tSplitR.Left + m_lSplitSize
               m_tSplitR.Top = tR.Top
               m_tSplitR.Bottom = tR.Bottom
            Else
               m_tSplitR.Left = tR.Left
               m_tSplitR.Right = tR.Right
               m_tSplitR.Top = tR.Top + m_lSplitPos
               m_tSplitR.Bottom = m_tSplitR.Top + m_lSplitSize
            End If
               
            pDrawSplitter
   
         End If
         
      Else
         If pHitTest(x, y) Then SetCursor m_hCursor
      End If
   End If
End Sub

Private Sub MouseUp(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long)
   If (pbConfigured()) Then
      
      If (m_bInDrag) Then
         ClipCursorClear 0&
         
         Dim tP As POINTAPI
         GetCursorPos tP
         
         If Not m_bFullDrag Then
            pDrawSplitter
         End If
         
         If (m_eOrientation = cSPLTOrientationVertical) Then
            m_lSplitPos = m_lSplitInitial + (tP.x - m_tPInitial.x)
         Else
            m_lSplitPos = m_lSplitInitial + (tP.y - m_tPInitial.y)
         End If
         pValidatePosition
            
         pResizePanels
         
         pSetProportion
         m_bInDrag = False
      End If
   End If
End Sub

Private Sub pDrawSplitter()
Dim lHDC As Long
Dim hOldBrush As Long
   lHDC = CreateDCAsNull("DISPLAY", ByVal 0&, ByVal 0&, ByVal 0&)
   hOldBrush = SelectObject(lHDC, m_hBrush)
   PatBlt lHDC, m_tSplitR.Left, m_tSplitR.Top, m_tSplitR.Right - m_tSplitR.Left, m_tSplitR.Bottom - m_tSplitR.Top, PATINVERT
   SelectObject lHDC, hOldBrush
   DeleteDC lHDC
End Sub

Private Sub pSetProportion()
Dim r As RECT
GetClientRect ContainerHwnd, r
   If (m_eOrientation = cSPLTOrientationVertical) Then
      m_fProportion = (m_lSplitPos * 1#) / r.Right
   Else
      m_fProportion = (m_lSplitPos * 1#) / r.Bottom
   End If
End Sub

Private Sub pValidatePosition()
   
   Dim tR As RECT
   GetClientRect ContainerHwnd, tR
   
   If (m_eOrientation = cSPLTOrientationVertical) Then
      ' Check right too big:
      If (m_lMaxSize(2) > 0) Then
         If ((tR.Right - m_lSplitPos - m_lSplitSize) > m_lMaxSize(2)) Then
            m_lSplitPos = tR.Right - m_lMaxSize(2) - m_lSplitSize
         End If
      End If
      ' Check left too big:
      If (m_lMaxSize(1) > 0) Then
         If (m_lSplitPos > m_lMaxSize(1)) Then
            m_lSplitPos = m_lMaxSize(1)
         End If
      End If
      ' Check right too small:
      If (m_lMinSize(2) > 0) Then
         If ((tR.Right - m_lSplitPos - m_lSplitSize) < m_lMinSize(2)) Then
            m_lSplitPos = tR.Right - m_lMinSize(2) - m_lSplitSize
         End If
      End If
      ' Check left too small:
      If (m_lMinSize(1) > 0) Then
         If (m_lSplitPos < m_lMinSize(1)) Then
            m_lSplitPos = m_lMinSize(1)
         End If
      End If
   Else
      ' Check bottom too big:
      If (m_lMaxSize(2) > 0) Then
         If ((tR.Bottom - m_lSplitPos - m_lSplitSize) > m_lMaxSize(2)) Then
            m_lSplitPos = tR.Bottom - m_lMaxSize(2) - m_lSplitSize
         End If
      End If
      ' Check top too big:
      If (m_lMaxSize(1) > 0) Then
         If (m_lSplitPos > m_lMaxSize(1)) Then
            m_lSplitPos = m_lMaxSize(1)
         End If
      End If
      ' Bottom too small:
      If (m_lMinSize(2) > 0) Then
         If ((tR.Bottom - m_lSplitPos - m_lSplitSize) < m_lMinSize(2)) Then
            m_lSplitPos = tR.Bottom - m_lMinSize(2) - m_lSplitSize
         End If
      End If
      ' Top too small:
      If (m_lMinSize(1) > 0) Then
         If (m_lSplitPos < m_lMinSize(1)) Then
            m_lSplitPos = m_lMinSize(1)
         End If
      End If
   End If
End Sub

Public Sub Resize()
   If pbConfigured() Then
            
      ' Get the container's size:
      Dim tR As RECT
      GetClientRect ContainerHwnd, tR
      
      If (m_bKeepProportionsWhenResizing) Then
         ' attempt to keep the proportions of the two parts:
         If (m_eOrientation = cSPLTOrientationVertical) Then
            m_lSplitPos = (tR.Right - tR.Left) * m_fProportion
         Else
            m_lSplitPos = (tR.Bottom - tR.Top) * m_fProportion
         End If
         pValidatePosition
      End If
            
      pResizePanels
      
   End If
End Sub

Private Sub pResizePanels()
   Dim i As Long, b As Boolean
   Dim r As RECT
   On Error Resume Next
   GetClientRect ContainerHwnd, r
   '////////add!!!
   pResizePanels00 m_oSubObj(0), r
   'check panel2 visible
   b = False
   With m_oObjects(2)
    For i = 1 To .nCount
     If .d(i).Visible Then
      b = True
      Exit For
     End If
    Next i
   End With
   If Not b Then
      pResizePanels0 m_oObjects(1), m_oSubObj(1), r.Left, r.Top, r.Right, r.Bottom
      Exit Sub
   End If
   'check panel1 visible
   b = False
   With m_oObjects(1)
    For i = 1 To .nCount
     If .d(i).Visible Then
      b = True
      Exit For
     End If
    Next i
   End With
   If Not b Then
      pResizePanels0 m_oObjects(2), m_oSubObj(2), r.Left, r.Top, r.Right, r.Bottom
      Exit Sub
   End If
   '////////
   If (m_eOrientation = cSPLTOrientationHorizontal) Then
      i = m_lSplitPos
      pResizePanels0 m_oObjects(1), m_oSubObj(1), r.Left, r.Top, r.Right, i
      pResizePanels0 m_oObjects(2), m_oSubObj(2), r.Left, i + m_lSplitSize, r.Right, r.Bottom
   Else
      i = m_lSplitPos
      pResizePanels0 m_oObjects(1), m_oSubObj(1), r.Left, r.Top, i, r.Bottom
      pResizePanels0 m_oObjects(2), m_oSubObj(2), i + m_lSplitSize, r.Top, r.Right, r.Bottom
   End If

End Sub

Private Sub pResizePanels0(obj As typeFakeSplitObjects, objs As typeFakeSplitSubObjects, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long)
On Error Resume Next
Dim i As Long, j As Long
With objs
 For i = 1 To .nCount
  With .d(i)
   If Not .obj Is Nothing Then
    If .obj.Visible Then
     Select Case .nDir
     Case 2
      j = .obj.Height
      Bottom = Bottom - j
      .obj.Move Left, Bottom, Right - Left, j
     Case 3
      j = .obj.Width
      .obj.Move Left, Top, j, Bottom - Top
      Left = Left + j
     Case 4
      j = .obj.Width
      Right = Right - j
      .obj.Move Right, Top, j, Bottom - Top
     Case Else
      j = .obj.Height
      .obj.Move Left, Top, Right - Left, j
      Top = Top + j
     End Select
    End If
   End If
  End With
 Next i
End With
Right = Right - Left
Bottom = Bottom - Top
For i = 1 To obj.nCount
 obj.d(i).Move Left, Top, Right, Bottom
Next i
End Sub

Private Sub pResizePanels00(objs As typeFakeSplitSubObjects, ByRef r As RECT)
On Error Resume Next
Dim i As Long, j As Long
With objs
 For i = 1 To .nCount
  With .d(i)
   If Not .obj Is Nothing Then
    If .obj.Visible Then
     Select Case .nDir
     Case 2
      j = .obj.Height
      r.Bottom = r.Bottom - j
      .obj.Move r.Left, r.Bottom, r.Right - r.Left, j
     Case 3
      j = .obj.Width
      .obj.Move r.Left, r.Top, j, r.Bottom - r.Top
      r.Left = r.Left + j
     Case 4
      j = .obj.Width
      r.Right = r.Right - j
      .obj.Move r.Right, r.Top, j, r.Bottom - r.Top
     Case Else
      j = .obj.Height
      .obj.Move r.Left, r.Top, r.Right - r.Left, j
      r.Top = r.Top + j
     End Select
    End If
   End If
  End With
 Next i
End With
End Sub

Private Function CreateBrush() As Boolean
Dim tbm As BITMAP
Dim hbm As Long

   DestroyBrush
      
   ' Create a monochrome bitmap containing the desired pattern:
   tbm.bmType = 0
   tbm.bmWidth = 16
   tbm.bmHeight = 8
   tbm.bmWidthBytes = 2
   tbm.bmPlanes = 1
   tbm.bmBitsPixel = 1
   tbm.bmBits = VarPtr(m_lPattern(0))
   hbm = CreateBitmapIndirect(tbm)

   ' Make a brush from the bitmap bits
   m_hBrush = CreatePatternBrush(hbm)

   '// Delete the useless bitmap
   DeleteObject hbm

End Function
Private Sub DestroyBrush()
   If Not (m_hBrush = 0) Then
      DeleteObject m_hBrush
      m_hBrush = 0
   End If
End Sub

Private Sub UserControl_Initialize()
   
   m_fProportion = 0.5
   m_eOrientation = cSPLTOrientationHorizontal
      m_hCursor = LoadCursorLong(0, IDC_SIZENS)
   m_lSplitSize = 4
   m_lMinSize(1) = 8
   m_lMaxSize(1) = -1
   m_lMinSize(2) = 8
   m_lMaxSize(2) = -1
   m_bFullDrag = True
   m_lSplitPos = 128
   
   Dim i As Long
   For i = 0 To 3
      m_lPattern(i) = &HAAAA5555
   Next i
   
End Sub

Private Sub UserControl_InitProperties()
bEnabled = True
pInit
End Sub

Private Sub pInit()
LetContainer
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
 Orientation = .ReadProperty("Orientation", 1)
 SplitterSize = .ReadProperty("SplitterSize", 4)
 FullDrag = .ReadProperty("FullDrag", True)
 KeepProportion = .ReadProperty("KeepProportion", False)
 bEnabled = .ReadProperty("Enabled", True)
End With
pInit
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
 .WriteProperty "Orientation", Orientation, 1
 .WriteProperty "SplitterSize", SplitterSize, 4
 .WriteProperty "FullDrag", FullDrag, True
 .WriteProperty "KeepProportion", KeepProportion, False
 .WriteProperty "Enabled", bEnabled, True
End With
End Sub

Private Sub UserControl_Terminate()

On Error Resume Next
cSub.DelMsg -1, MSG_AFTER
cSub.UnSubclass

   DestroyBrush
   If Not (m_hCursor = 0) Then
      DestroyCursor m_hCursor
   End If
End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
If bEnabled Then lReturn = pWindowProc(hwnd, uMsg, wParam, lParam)
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)
'
End Sub

Private Function pWindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim x As Long, y As Long
x = lParam And &HFFFF&
y = lParam \ &H10000
Select Case iMsg
Case WM_SIZE
 Resize
Case WM_MOUSEMOVE
 MouseMove (wParam), 0, (x), (y)
Case WM_LBUTTONDOWN
 MouseDown 1, 0, (x), (y)
Case WM_LBUTTONUP
 MouseUp 1, 0, (x), (y)
End Select
End Function

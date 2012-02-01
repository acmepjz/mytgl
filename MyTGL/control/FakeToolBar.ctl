VERSION 5.00
Begin VB.UserControl FakeToolBar 
   Alignable       =   -1  'True
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
   Begin VB.Timer tmrKey 
      Enabled         =   0   'False
      Interval        =   50
      Left            =   1680
      Top             =   1200
   End
   Begin VB.Timer tmrShadow 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   1200
      Top             =   1680
   End
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   1200
      Top             =   1200
   End
End
Attribute VB_Name = "FakeToolBar"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////
'This file is public domain.
'////////////////////////////////

#Const UseFakeMenu = 1

Private Declare Function VarPtrArray Lib "msvbvm60.dll" Alias "VarPtr" (ByRef Ptr() As Any) As Long
Private Declare Function LockWindowUpdate Lib "user32.dll" (ByVal hwndLock As Long) As Long

Public Enum enumFakeTBOrientation
 ftboTop = 0
 ftboBottom = 1
 ftboLeft = 2
 ftboRight = 3
 ftboVertical = 4 '???
End Enum

Private btns() As typeFakeButton
Private btnc As Long

Private bmPic As cDIBSection, bmGray As cDIBSection
Private pic As StdPicture

Private theStr As String ':-3

Private objTB As New IFakeToolbarDraw
Implements IFakeToolbarDraw

Public Event Click(ByVal btnIndex As Long, ByVal btnKey As String)

#If UseFakeMenu Then

Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Private Type typeFakeMenuHotKey
 MenuKey As String
 ButtonKey As String
 KeyCode As Long
 Shift As Long
End Type

Private hwdParent As Long

Private btnHotKey() As Long
Private ks() As typeFakeMenuHotKey, kc As Long

Public Property Get Count() As Long
Count = btnc
End Property

Public Property Get MainMenu() As Boolean
MainMenu = objTB.MainMenu
End Property

Public Property Let MainMenu(ByVal b As Boolean)
objTB.MainMenu = b
tmrKey.Enabled = b And Not objTB.MenuObject Is Nothing
objTB.Redraw btns, btnc
End Property

Public Sub SetMenu(obj As FakeMenu, Optional ByVal bBindBitmap As Boolean)
Set objTB.MenuObject = obj
tmrKey.Enabled = objTB.MainMenu And Not obj Is Nothing
If bBindBitmap And Not pic Is Nothing Then
 obj.fBindBitmap bmPic, bmGray, objTB.TransparentColor
End If
objTB.Refresh btns, btnc
End Sub

Public Function AddShortcutKey(ByVal MenuKey As String, ByVal ButtonKey As String, ByVal KeyCode As KeyCodeConstants, ByVal Shift As ShiftConstants) As Long
If Not objTB.MainMenu Then Exit Function
kc = kc + 1
ReDim Preserve ks(1 To kc)
With ks(kc)
 .MenuKey = MenuKey
 .ButtonKey = ButtonKey
 .KeyCode = KeyCode
 .Shift = Shift
End With
AddShortcutKey = kc
End Function

Public Sub RemoveShortcutKey(ByVal Index As Long)
Dim i As Long
If Not objTB.MainMenu Then Exit Sub
If Index > 0 And Index <= kc Then
 If kc <= 1 Then
  Erase ks
  kc = 0
 Else
  For i = Index + 1 To kc
   ks(i - 1) = ks(i)
  Next i
  kc = kc - 1
  ReDim Preserve ks(1 To kc)
 End If
End If
End Sub

Public Sub ClearShortcutKey()
Erase ks
kc = 0
End Sub

Public Property Get ShortcutKeyMenuKey(ByVal Index As Long) As String
ShortcutKeyMenuKey = ks(Index).MenuKey
End Property

Public Property Let ShortcutKeyMenuKey(ByVal Index As Long, ByVal s As String)
ks(Index).MenuKey = s
End Property

Public Property Get ShortcutKeyButtonKey(ByVal Index As Long) As String
ShortcutKeyButtonKey = ks(Index).ButtonKey
End Property

Public Property Let ShortcutKeyButtonKey(ByVal Index As Long, ByVal s As String)
ks(Index).ButtonKey = s
End Property

Public Property Get ShortcutKeyCode(ByVal Index As Long) As KeyCodeConstants
ShortcutKeyCode = ks(Index).KeyCode
End Property

Public Property Let ShortcutKeyCode(ByVal Index As Long, ByVal n As KeyCodeConstants)
ks(Index).KeyCode = n
End Property

Public Property Get ShortcutKeyShift(ByVal Index As Long) As ShiftConstants
ShortcutKeyShift = ks(Index).Shift
End Property

Public Property Let ShortcutKeyShift(ByVal Index As Long, ByVal n As ShiftConstants)
ks(Index).Shift = n
End Property

Private Sub pHotKey(ByVal Index As Long)
Dim s As String
Dim i As Long
s = Replace(btns(Index).s, "&&", "")
i = InStr(1, s, "&")
If i > 0 And i < Len(s) Then
 i = Asc(UCase(Mid(s, i + 1, 1)))
Else
 i = 0
End If
btnHotKey(Index) = i
End Sub

#End If

'////////button properties

Public Function FindButton(ByVal Key As String) As Long
Dim i As Long
For i = 1 To btnc
 If btns(i).sKey = Key Then
  FindButton = i
  Exit Function
 End If
Next i
End Function

Public Property Get ButtonValue(ByVal Index As Long) As Long
ButtonValue = btns(Index).Value
End Property

Public Property Let ButtonValue(ByVal Index As Long, ByVal n As Long)
n = n And &HFF&
If btns(Index).Value <> n Then
 btns(Index).Value = n
 objTB.Redraw btns, btnc
End If
End Property

Public Property Get ButtonFlags(ByVal Index As Long) As enumFakeButtonFlags
ButtonFlags = btns(Index).nFlags
End Property

Public Property Let ButtonFlags(ByVal Index As Long, ByVal n As enumFakeButtonFlags)
If n <> btns(Index).nFlags Then
 btns(Index).nFlags = n
 objTB.Refresh btns, btnc
End If
End Property

Public Property Get ButtonType(ByVal Index As Long) As enumFakeButtonType
ButtonType = btns(Index).nType
End Property

Public Property Let ButtonType(ByVal Index As Long, ByVal n As enumFakeButtonType)
If n <> btns(Index).nType Then
 btns(Index).nType = n
 objTB.Refresh btns, btnc
End If
End Property

Public Property Get ButtonGroupIndex(ByVal Index As Long) As Long
ButtonGroupIndex = btns(Index).GroupIndex
End Property

Public Property Let ButtonGroupIndex(ByVal Index As Long, ByVal n As Long)
btns(Index).GroupIndex = n
End Property

Public Property Get ButtonPicLeft(ByVal Index As Long) As Long
ButtonPicLeft = btns(Index).PicLeft
End Property

Public Property Let ButtonPicLeft(ByVal Index As Long, ByVal n As Long)
If btns(Index).PicLeft <> n Then
 btns(Index).PicLeft = n
 objTB.Refresh btns, btnc
End If
End Property

Public Property Get ButtonCaption(ByVal Index As Long) As String
ButtonCaption = btns(Index).s
End Property

Public Property Let ButtonCaption(ByVal Index As Long, ByVal s As String)
If btns(Index).s <> s Then
 btns(Index).s = s
 #If UseFakeMenu Then
 If objTB.MainMenu Then pHotKey Index
 #End If
 objTB.Refresh btns, btnc
End If
End Property

Public Property Get ButtonToolTipText(ByVal Index As Long) As String
ButtonToolTipText = btns(Index).s2
End Property

Public Property Let ButtonToolTipText(ByVal Index As Long, ByVal s As String)
btns(Index).s2 = s
End Property

Public Property Get ButtonDescription(ByVal Index As Long) As String
ButtonDescription = btns(Index).sDesc
End Property

Public Property Let ButtonDescription(ByVal Index As Long, ByVal s As String)
btns(Index).sDesc = s
End Property

Public Property Get ButtonSubMenu(ByVal Index As Long) As String
ButtonSubMenu = btns(Index).sSubMenu
End Property

Public Property Let ButtonSubMenu(ByVal Index As Long, ByVal s As String)
If btns(Index).sSubMenu <> s Then
 btns(Index).sSubMenu = s
 objTB.Refresh btns, btnc
End If
End Property

Public Property Get ButtonKey(ByVal Index As Long) As String
ButtonKey = btns(Index).sKey
End Property

Public Property Let ButtonKey(ByVal Index As Long, ByVal s As String)
btns(Index).sKey = s
End Property

Public Function AddButton(Optional ByVal Index As Long, Optional ByVal Key As String, Optional ByVal Caption As String, Optional ByVal ToolTipText As String, Optional ByVal nType As enumFakeButtonType, Optional ByVal nFlags As enumFakeButtonFlags, _
Optional ByVal GroupIndex As Long, Optional ByVal PicLeft As Long = -1, Optional ByVal Description As String, Optional ByVal SubMenuKey As String, Optional ByVal Checked As Boolean) As Long
Dim i As Long
i = InStr(1, Caption, vbTab)
If i > 0 Then Caption = Left(Caption, i - 1)
If Index <= 0 Or Index > btnc Then Index = btnc + 1
btnc = btnc + 1
ReDim Preserve btns(1 To btnc)
#If UseFakeMenu Then
ReDim Preserve btnHotKey(1 To btnc)
#End If
For i = btnc To Index + 1 Step -1
 btns(i) = btns(i - 1)
 #If UseFakeMenu Then
 btnHotKey(i) = btnHotKey(i - 1)
 #End If
Next i
With btns(Index)
 .sKey = Key
 .s = Caption
 .s2 = ToolTipText
 .sTab = ""
 .sDesc = Description
 .sSubMenu = SubMenuKey
 .nType = nType
 .nFlags = nFlags
 .Value = Checked And 1&
 .GroupIndex = GroupIndex
 .PicLeft = PicLeft
End With
#If UseFakeMenu Then
If objTB.MainMenu Then pHotKey Index
#End If
AddButton = Index
objTB.Refresh btns, btnc
End Function

Public Sub RemoveButton(ByVal Index As Long)
Dim i As Long
If Index > 0 And Index <= btnc Then
 If btnc <= 1 Then
  Erase btns
  #If UseFakeMenu Then
  Erase btnHotKey
  #End If
  btnc = 0
  objTB.Redraw btns, btnc
 Else
  For i = Index + 1 To btnc
   btns(i - 1) = btns(i)
   #If UseFakeMenu Then
   btnHotKey(i - 1) = btnHotKey(i)
   #End If
  Next i
  btnc = btnc - 1
  ReDim Preserve btns(1 To btnc)
  #If UseFakeMenu Then
  ReDim Preserve btnHotKey(1 To btnc)
  #End If
  objTB.Refresh btns, btnc
 End If
End If
End Sub

Public Sub RemoveButtonEx(Optional ByVal idxStart As Long = 1, Optional ByVal idxEnd As Long)
Dim i As Long, j As Long
If idxStart > 0 And idxStart <= btnc Then
 If idxEnd <= 0 Or idxEnd > btnc Then idxEnd = btnc
 If idxEnd < idxStart Then idxEnd = idxStart
 If idxStart = 1 And idxEnd >= btnc Then
  Erase btns
  #If UseFakeMenu Then
  Erase btnHotKey
  #End If
  btnc = 0
  objTB.Redraw btns, btnc
 Else
  j = idxEnd - idxStart + 1
  For i = idxEnd + 1 To btnc
   btns(i - j) = btns(i)
   #If UseFakeMenu Then
   btnHotKey(i - j) = btnHotKey(i)
   #End If
  Next i
  btnc = btnc - j
  ReDim Preserve btns(1 To btnc)
  #If UseFakeMenu Then
  ReDim Preserve btnHotKey(1 To btnc)
  #End If
  objTB.Refresh btns, btnc
 End If
End If
End Sub

'////////general properties

Public Property Get Orientation() As enumFakeTBOrientation
Orientation = objTB.Orientation
End Property

Public Property Let Orientation(ByVal n As enumFakeTBOrientation)
If objTB.Orientation <> n Then
 objTB.Orientation = n
 objTB.Redraw btns, btnc
End If
End Property

Public Property Get Picture() As StdPicture
Set Picture = pic
End Property

Public Property Set Picture(obj As StdPicture)
Set pic = obj
pChangePic
objTB.Refresh btns, btnc
End Property

Public Property Get Font() As StdFont
Set Font = objTB.Font
End Property

Public Property Set Font(obj As StdFont)
Set objTB.Font = obj
objTB.Refresh btns, btnc
End Property

Public Property Get TransparentColor() As OLE_COLOR
TransparentColor = objTB.TransparentColor
End Property

Public Property Let TransparentColor(ByVal clr As OLE_COLOR)
objTB.TransparentColor = clr
objTB.Redraw btns, btnc
End Property

Public Property Get TheString() As String
TheString = theStr
End Property

Public Property Let TheString(s As String)
theStr = s
pButtonFromStr
objTB.Redraw btns, btnc
End Property

Private Sub IFakeToolbarDraw_Click(ByVal btnIndex As Long, ByVal btnKey As String)
RaiseEvent Click(btnIndex, btnKey)
End Sub

Private Sub IFakeToolbarDraw_GetButtonSafeArrayData(lpSafeArray As Long, btnCount As Long)
lpSafeArray = VarPtrArray(btns)
btnCount = btnc
End Sub

Private Sub IFakeToolbarDraw_Paint()
UserControl_Paint
End Sub

Private Sub IFakeToolbarDraw_SetToolTipText(ByVal s As String)
On Error Resume Next
Extender.ToolTipText = s
End Sub

Private Sub t1_Timer()
objTB.OnTimer btns, btnc
End Sub

#If UseFakeMenu Then
Private Sub tmrKey_Timer()
Dim i As Long, j As Long, k As Long
Dim bShift As Boolean, bCtrl As Boolean, bAlt As Boolean
Dim nKey(255) As Long
Dim b As Boolean, bMenu As Boolean, s As String
'////////
If hwdParent <> 0 Then
 If GetActiveWindow <> hwdParent Then Exit Sub '????????
End If
'////////
Dim objMenu As FakeMenu
Set objMenu = objTB.MenuObject
If objMenu.MenuWindowCount > 0 Then Exit Sub
'get shift constants
i = vbKeyShift
j = (GetAsyncKeyState(i) And &HFFFF&) Or &H80000000
nKey(i) = j
bShift = j And &H8000&
i = vbKeyControl
j = (GetAsyncKeyState(i) And &HFFFF&) Or &H80000000
nKey(i) = j
bCtrl = j And &H8000&
i = vbKeyMenu
j = (GetAsyncKeyState(i) And &HFFFF&) Or &H80000000
nKey(i) = j
bAlt = j And &H8000&
'check hot-key
If bAlt And Not bCtrl And Not bShift Then
 For i = 1 To btnc
  b = False
  bMenu = False
  With btns(i)
   If (.nFlags And 3&) = 0 And .nType <> 1 And .nType <> 6 Then
    j = btnHotKey(i)
    If j > 0 And j <= 255 Then
     If nKey(j) = 0 Then nKey(j) = (GetAsyncKeyState(j) And &HFFFF&) Or &H80000000
     If (nKey(j) And &H8001&) = &H8001& Then
      b = True
      If .sSubMenu <> "" Then
       If objMenu.HasMenu(.sSubMenu) Then bMenu = True
      End If
     End If
    End If
   End If
  End With
  If b Then
   If bMenu Then
    objTB.PopupMenu btns(i), i
    objTB.TimerEnabled = True
   Else
    With btns(i)
     Select Case .nType
     Case 2 'check
      .Value = (.Value = 0) And 1&
     Case 3 'option
      j = .GroupIndex
      For k = 1 To btnc
       With btns(k)
        If .nType = 3 And .GroupIndex = j Then
         .Value = (k = i) And 1&
       End If
       End With
      Next k
     Case 4 'optnull
      If .Value Then
       .Value = 0
      Else
       j = .GroupIndex
       For k = 1 To btnc
        With btns(k)
         If .nType = 4 And .GroupIndex = j Then
          .Value = (k = i) And 1&
         End If
        End With
       Next k
      End If
     End Select
     RaiseEvent Click(i, .sKey)
    End With
   End If
   objTB.Redraw btns, btnc
   Exit Sub
  End If
 Next i
End If
'check shortcut-key
For i = 1 To kc
 If ((bShift And vbShiftMask) Or (bCtrl And vbCtrlMask) Or (bAlt And vbAltMask)) = ks(i).Shift Then
  j = ks(i).KeyCode
  If j > 0 And j <= 255 Then
   If nKey(j) = 0 Then nKey(j) = (GetAsyncKeyState(j) And &HFFFF&) Or &H80000000
   If (nKey(j) And &H8001&) = &H8001& Then
    objMenu.Click ks(i).MenuKey, ks(i).ButtonKey
    Exit Sub
   End If
  End If
 End If
Next i
End Sub
#End If

Private Sub tmrShadow_Timer()
If EnableTooltipDropShadow Then tmrShadow.Enabled = False
End Sub

Private Sub UserControl_Click()
objTB.OnClick btns, btnc
End Sub

Private Sub UserControl_DblClick()
objTB.OnDblClick btns, btnc
End Sub

Private Sub UserControl_InitProperties()
Set objTB.Font = UserControl.Font
pButtonFromStr
pInit
End Sub

Private Sub UserControl_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
objTB.OnMouseDown btns, btnc, Button, Shift, x, y
End Sub

Private Sub UserControl_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
objTB.OnMouseMove btns, btnc, Button, Shift, x, y
End Sub

Private Sub UserControl_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
objTB.OnMouseUp btns, btnc, Button, Shift, x, y
End Sub

Private Sub UserControl_Paint()
objTB.TheBitmap.PaintPicture hdc
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
 Set pic = .ReadProperty("Picture", Nothing)
 objTB.TransparentColor = .ReadProperty("TransparentColor", vbGreen)
 theStr = .ReadProperty("TheString", "")
 Set objTB.Font = .ReadProperty("Font", UserControl.Font)
 objTB.Orientation = .ReadProperty("Orientation", 0)
 #If UseFakeMenu Then
 objTB.MainMenu = .ReadProperty("MainMenu", False)
 #End If
End With
pChangePic
pInit
pButtonFromStr
End Sub

Private Sub UserControl_Resize()
objTB.Resize 0, 0, ScaleWidth, ScaleHeight, hwnd, btns, btnc
End Sub

Private Sub UserControl_Terminate()
Set bmPic = Nothing
Set bmGray = Nothing
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
 .WriteProperty "Picture", pic, Nothing
 .WriteProperty "TransparentColor", objTB.TransparentColor, vbGreen
 .WriteProperty "TheString", theStr
 .WriteProperty "Font", objTB.Font, UserControl.Font
 .WriteProperty "Orientation", objTB.Orientation, 0
 #If UseFakeMenu Then
 .WriteProperty "MainMenu", objTB.MainMenu, False
 #End If
End With
End Sub

Private Sub pChangePic()
If pic Is Nothing Then
 Set bmPic = Nothing
 Set bmGray = Nothing
 objTB.PicSize = 16
Else
 Set bmPic = New cDIBSection
 Set bmGray = New cDIBSection
 bmPic.CreateFromPicture pic
 GrayscaleBitmap bmPic, bmGray, d_Icon_Grayscale, objTB.TransparentColor
 objTB.PicSize = bmPic.Height
End If
objTB.SetBitmap bmPic, bmGray
End Sub

Private Sub pButtonFromStr()
Dim d As typeFakeCommandBar
FakeCommandBarFromString theStr, d, objTB.PicSize
CreateIndirect d
End Sub

Public Sub Clear()
Erase btns
#If UseFakeMenu Then
Erase btnHotKey
#End If
btnc = 0
objTB.Redraw btns, btnc
End Sub

Private Sub pInit()
tmrShadow.Enabled = Ambient.UserMode
Set objTB.TheTimer = t1
objTB.SetCallback Me
objTB.Resize 0, 0, ScaleWidth, ScaleHeight, hwnd, btns, btnc
#If UseFakeMenu Then
On Error Resume Next
hwdParent = Parent.hwnd
#End If
End Sub

Public Function FakeBeginPaint(Optional nNewLeft As Long, Optional nWidth As Long, Optional nHeight As Long) As Long
LockWindowUpdate hwnd
objTB.Redraw btns, btnc, False
nNewLeft = 0
If btnc > 0 Then
 With btns(btnc)
  nNewLeft = .Left + .Width
 End With
End If
nWidth = ScaleWidth
nHeight = ScaleHeight
FakeBeginPaint = objTB.TheBitmap.hdc
End Function

Public Sub FakeEndPaint()
LockWindowUpdate 0
UserControl_Paint
End Sub

Public Sub Refresh()
objTB.Refresh btns, btnc
End Sub

Friend Sub CreateIndirect(d As typeFakeCommandBar)
btns = d.d
btnc = d.nCount
#If UseFakeMenu Then
Dim i As Long
If btnc > 0 Then ReDim btnHotKey(1 To btnc) Else Erase btnHotKey
If objTB.MainMenu Then
 For i = 1 To btnc
  pHotKey i
 Next i
End If
#End If
objTB.Refresh btns, btnc
End Sub

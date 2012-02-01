VERSION 5.00
Begin VB.UserControl FakeMenu 
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
   Begin VB.Timer tmrShadow 
      Enabled         =   0   'False
      Interval        =   100
      Left            =   240
      Top             =   840
   End
   Begin VB.Timer t1 
      Enabled         =   0   'False
      Interval        =   10
      Left            =   240
      Top             =   360
   End
   Begin VB.PictureBox p1 
      BorderStyle     =   0  'None
      Height          =   1095
      Index           =   0
      Left            =   960
      ScaleHeight     =   73
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   113
      TabIndex        =   1
      Top             =   840
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "A really fake menu"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   1575
   End
End
Attribute VB_Name = "FakeMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////
'This file is public domain.
'////////////////////////////////

Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long
Private Const HWND_TOPMOST As Long = -1
Private Const SWP_FRAMECHANGED As Long = &H20
Private Const SWP_HIDEWINDOW As Long = &H80
Private Const SWP_NOACTIVATE As Long = &H10
Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOOWNERZORDER As Long = &H200
Private Const SWP_NOREDRAW As Long = &H8
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_SHOWWINDOW As Long = &H40
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Const GWL_HWNDPARENT As Long = -8
Private Const GWL_EXSTYLE As Long = -20
Private Const WS_EX_TOOLWINDOW As Long = &H80&
Private Const WS_EX_TOPMOST As Long = &H8&
'Private Const GWL_STYLE As Long = -16
'Private Const WS_VISIBLE As Long = &H10000000

Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Const SW_SHOW As Long = 5
Private Const SW_HIDE As Long = 0
Private Declare Function SetFocusAPI Lib "user32.dll" Alias "SetFocus" (ByVal hwnd As Long) As Long
'Private Declare Function SetActiveWindow Lib "user32.dll" (ByVal hwnd As Long) As Long '????????
'Private Declare Function SetForegroundWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
'Private Declare Function LockSetForegroundWindow Lib "user32.dll" (ByVal uLockCode As Long) As Long
Private Declare Function SetCursorPos Lib "user32.dll" (ByVal x As Long, ByVal y As Long) As Long
Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_ABSOLUTE As Long = &H8000
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2
Private Const MOUSEEVENTF_LEFTUP As Long = &H4
Private Const MOUSEEVENTF_MOVE As Long = &H1

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function SetPixelV Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal crColor As Long) As Long
Private Declare Function GetWindowDC Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long
Private Declare Function AlphaBlend Lib "msimg32.dll" (ByVal hdcDest As Long, ByVal xDest As Long, ByVal yDest As Long, ByVal WidthDest As Long, ByVal HeightDest As Long, ByVal hdcSrc As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal WidthSrc As Long, ByVal HeightSrc As Long, ByVal Blendfunc As Long) As Long
Private Declare Function SetBkMode Lib "gdi32.dll" (ByVal hdc As Long, ByVal nBkMode As Long) As Long
Private Const TRANSPARENT As Long = 1
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function GetSystemMetrics Lib "user32.dll" (ByVal nIndex As Long) As Long
Private Const SM_CXMENUCHECK As Long = 71
Private Const SM_CYMENUCHECK As Long = 72
Private Type MEASUREITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemWidth As Long
    itemHeight As Long
    itemData As Long
End Type
Private Type DRAWITEMSTRUCT
    CtlType As Long
    CtlID As Long
    itemID As Long
    itemAction As Long
    itemState As Long
    hwndItem As Long
    hdc As Long
    rcItem As RECT
    itemData As Long
End Type
Private Const ODT_MENU As Long = 1
Private Const ODA_SELECT As Long = &H2
Private Const ODA_DRAWENTIRE As Long = &H1
Private Const ODS_CHECKED As Long = &H8
Private Const ODS_DEFAULT As Long = &H20
Private Const ODS_SELECTED As Long = &H1
Private Const WM_MEASUREITEM As Long = &H2C
Private Const WM_DRAWITEM As Long = &H2B
Private Const WM_MENUSELECT As Long = &H11F
Private Const WM_COMMAND As Long = &H111
Private Const WM_MENUCOMMAND As Long = &H126
Private Const WM_ACTIVATE As Long = &H6
Private Declare Function GetMenuInfo Lib "user32.dll" (ByVal hMenu As Long, ByRef lpMenuInfo As MENUINFO) As Long
Private Type MENUINFO
 cbSize As Long
 fMask As Long
 dwStyle As Long
 cyMax As Long
 hbrBack As Long
 dwContextHelpID As Long
 dwMenuData As Long
End Type
Private Const MIM_STYLE As Long = &H10
Private Const MNS_NOTIFYBYPOS As Long = &H8000000
Private Declare Function HiliteMenuItem Lib "user32.dll" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long
Private Const MF_BYPOSITION As Long = &H400&
Private Const MF_HILITE As Long = &H80&

Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function WindowFromPoint Lib "user32.dll" (ByVal xPoint As Long, ByVal yPoint As Long) As Long
Private Declare Function GetActiveWindow Lib "user32.dll" () As Long
Private Declare Function GetKeyState Lib "user32.dll" (ByVal nVirtKey As Long) As Integer
Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function GetMessage Lib "user32.dll" Alias "GetMessageA" (ByRef lpMsg As typeMSG, ByVal hwnd As Long, ByVal wMsgFilterMin As Long, ByVal wMsgFilterMax As Long) As Long
Private Declare Function TranslateMessage Lib "user32.dll" (ByRef lpMsg As typeMSG) As Long
Private Declare Function DispatchMessage Lib "user32.dll" Alias "DispatchMessageA" (ByRef lpMsg As typeMSG) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type
Private Type typeMSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private hwndOwner As Long
Private cFnt As New CLogFont, cFntB As New CLogFont
Private objFntB As StdFont

'////////menu data

Private Type typeFakeMenuWindow
 nDef As Long 'default menu
 idx As Long 'index of menu
 idxHl As Long '0=none >0=menu index &H80000000=gripper
 nAlpha As Long
 x As Long
 y As Long
 w As Long
 h As Long
 FS As Long 'erased border -1=null,0,1,2,3
 X1 As Long 'erased border start pos
 X2 As Long 'erased border end pos
 bm As cDIBSection
 bmBack As cDIBSection '???
 bmHook As cDIBSection '??????
End Type

Private bmPic As cDIBSection, bmGray As cDIBSection, transClr As Long

'p0(0) can't use!!! will cause error
Private wnds() As typeFakeMenuWindow, wndc As Long '1-based
Private mnus As typeFakeCommandBars

'////////other data

Private Const PicSize As Long = 16 '??????
Private Const AnimationStep As Long = 24

Private HlWndIndex As Long, HlIndex As Long, HlKey As String
Private OldActiveWindow As Long

'////////properties

Public Enum enumFakeMenuCalcSizeMethod
 enumFakeMenuCalcSizeTwoRow = 0
 enumFakeMenuCalcSizeOneRow = 1
 enumFakeMenuCalcSizeOneRowRight = 2
End Enum

Private mode1 As Long, bShadow As Boolean, bAnim As Boolean, mnuAlpha As Long
Private bNC As Boolean

'////////friend properties

Private nUserData As Long, bHook As Boolean, bNotify As Boolean, m_hWnd As Long

'////////events

Public Event Click(ByVal idxMenu As Long, ByVal Key As String, ByVal idxButton As Long, ByVal ButtonKey As String, ByRef Value As Long)
Public Event HookedClick(ByVal hMenu As Long, ByVal nID As Long, ByVal idxButton As Long, ByRef bProcessed As Boolean) '?
Public Event MouseMove(ByVal idxMenu As Long, ByVal Key As String, ByVal idxButton As Long, ByVal ButtonKey As String, ByVal Description As String)
Public Event MeasureItem(ByVal idxMenu As Long, ByVal Key As String, ByVal idxButton As Long, ByVal ButtonKey As String, ByRef w As Long, ByRef h As Long, ByRef bDoDefault As Boolean)
Public Event MakeMenuFloat(ByVal idxMenu As Long, ByVal Key As String, ByVal x As Long, ByVal y As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
Public Event DrawItem(ByVal idxMenu As Long, ByVal Key As String, ByVal idxButton As Long, ByVal ButtonKey As String, ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, ByVal nType As enumFakeButtonOwnerDrawType, ByRef nFlags As enumFakeButtonOwnerDrawFlags)
'new
Public Event BeforeShowMenu(ByVal idxMenu As Long)

Friend Property Get UserData() As Long
UserData = nUserData
End Property

Friend Property Let UserData(ByVal n As Long)
nUserData = n
End Property

Friend Sub Hook(ByVal hwnd As Long, ByVal Notify As Boolean)
bHook = True
bNotify = Notify
m_hWnd = hwnd
End Sub

Friend Sub Unhook()
bHook = False
End Sub

Public Property Get FakeMenuNoConnect() As Boolean
FakeMenuNoConnect = bNC
End Property

Public Property Let FakeMenuNoConnect(ByVal b As Boolean)
bNC = b
End Property

Friend Sub fBindBitmap(bm1 As cDIBSection, bm2 As cDIBSection, ByVal TransparentColor As Long)
Set bmPic = bm1
Set bmGray = bm2
transClr = TransparentColor
End Sub

Friend Function fGetBitmap(bm1 As cDIBSection, bm2 As cDIBSection) As Long
Set bm1 = bmPic
Set bm2 = bmGray
fGetBitmap = transClr
End Function

Friend Sub fGetMenuData(ByVal idxMenu As Long, btns() As typeFakeButton, btnc As Long)
If idxMenu > 0 And idxMenu < mnus.nCount Then
 btns = mnus.d(idxMenu).d
 btnc = mnus.d(idxMenu).nCount
Else
 Erase btns
 btnc = 0
End If
End Sub

Public Sub SetPicture(pic As StdPicture, Optional ByVal TransparentColor As Long = vbGreen)
transClr = TransparentColor
If pic Is Nothing Then
 Set bmPic = Nothing
 Set bmGray = Nothing
Else
 Set bmPic = New cDIBSection
 Set bmGray = New cDIBSection
 bmPic.CreateFromPicture pic
 GrayscaleBitmap bmPic, bmGray, d_Icon_Grayscale, transClr
End If
End Sub

'////////button properties

Public Function FindMenu(ByVal Key As String) As Long
FindMenu = FakeCommandBarGetMenuIndex(mnus, Key)
End Function

Public Function FindButton(ByVal idxMenu As Long, ByVal Key As String) As Long
Dim i As Long
If idxMenu > 0 And idxMenu <= mnus.nCount Then
 For i = 1 To mnus.d(idxMenu).nCount
  If mnus.d(idxMenu).d(i).sKey = Key Then
   FindButton = i
   Exit Function
  End If
 Next i
End If
End Function

Public Property Get ButtonValue(ByVal idxMenu As Long, ByVal idxButton As Long) As Long
ButtonValue = mnus.d(idxMenu).d(idxButton).Value
End Property

Public Property Let ButtonValue(ByVal idxMenu As Long, ByVal idxButton As Long, ByVal n As Long)
mnus.d(idxMenu).d(idxButton).Value = n And &HFF&
End Property

Public Property Get ButtonFlags(ByVal idxMenu As Long, ByVal idxButton As Long) As enumFakeButtonFlags
ButtonFlags = mnus.d(idxMenu).d(idxButton).nFlags
End Property

Public Property Let ButtonFlags(ByVal idxMenu As Long, ByVal idxButton As Long, ByVal n As enumFakeButtonFlags)
With mnus.d(idxMenu).d(idxButton)
 If n <> .nFlags Then
  .nFlags = n
  mnus.d(idxMenu).nFlags2 = 1
 End If
End With
End Property

Public Property Get ButtonType(ByVal idxMenu As Long, ByVal idxButton As Long) As enumFakeButtonType
ButtonType = mnus.d(idxMenu).d(idxButton).nType
End Property

Public Property Let ButtonType(ByVal idxMenu As Long, ByVal idxButton As Long, ByVal n As enumFakeButtonType)
With mnus.d(idxMenu).d(idxButton)
 If n <> .nType Then
  .nType = n
  mnus.d(idxMenu).nFlags2 = 1
 End If
End With
End Property

Public Property Get ButtonGroupIndex(ByVal idxMenu As Long, ByVal idxButton As Long) As Long
ButtonGroupIndex = mnus.d(idxMenu).d(idxButton).GroupIndex
End Property

Public Property Let ButtonGroupIndex(ByVal idxMenu As Long, ByVal idxButton As Long, ByVal n As Long)
mnus.d(idxMenu).d(idxButton).GroupIndex = n
End Property

Public Property Get ButtonPicLeft(ByVal idxMenu As Long, ByVal idxButton As Long) As Long
ButtonPicLeft = mnus.d(idxMenu).d(idxButton).PicLeft
End Property

Public Property Let ButtonPicLeft(ByVal idxMenu As Long, ByVal idxButton As Long, ByVal n As Long)
mnus.d(idxMenu).d(idxButton).PicLeft = n
End Property

Public Property Get ButtonCaption(ByVal idxMenu As Long, ByVal idxButton As Long) As String
ButtonCaption = mnus.d(idxMenu).d(idxButton).s
End Property

Public Property Let ButtonCaption(ByVal idxMenu As Long, ByVal idxButton As Long, ByVal s As String)
With mnus.d(idxMenu).d(idxButton)
 If .s <> s Then
  .s = s
  mnus.d(idxMenu).nFlags2 = 1
 End If
End With
End Property

Public Property Get ButtonCaption2(ByVal idxMenu As Long, ByVal idxButton As Long) As String
ButtonCaption2 = mnus.d(idxMenu).d(idxButton).sTab
End Property

Public Property Let ButtonCaption2(ByVal idxMenu As Long, ByVal idxButton As Long, ByVal s As String)
With mnus.d(idxMenu).d(idxButton)
 If .sTab <> s Then
  .sTab = s
  mnus.d(idxMenu).nFlags2 = 1
 End If
End With
End Property

Public Property Get ButtonToolTipText(ByVal idxMenu As Long, ByVal idxButton As Long) As String
ButtonToolTipText = mnus.d(idxMenu).d(idxButton).s2
End Property

Public Property Let ButtonToolTipText(ByVal idxMenu As Long, ByVal idxButton As Long, ByVal s As String)
mnus.d(idxMenu).d(idxButton).s2 = s
End Property

Public Property Get ButtonDescription(ByVal idxMenu As Long, ByVal idxButton As Long) As String
ButtonDescription = mnus.d(idxMenu).d(idxButton).sDesc
End Property

Public Property Let ButtonDescription(ByVal idxMenu As Long, ByVal idxButton As Long, ByVal s As String)
mnus.d(idxMenu).d(idxButton).sDesc = s
End Property

Public Property Get ButtonSubMenu(ByVal idxMenu As Long, ByVal idxButton As Long) As String
ButtonSubMenu = mnus.d(idxMenu).d(idxButton).sSubMenu
End Property

Public Property Let ButtonSubMenu(ByVal idxMenu As Long, ByVal idxButton As Long, ByVal s As String)
With mnus.d(idxMenu).d(idxButton)
 If .sSubMenu <> s Then
  .sSubMenu = s
  mnus.d(idxMenu).nFlags2 = 1
 End If
End With
End Property

Public Property Get ButtonKey(ByVal idxMenu As Long, ByVal idxButton As Long) As String
ButtonKey = mnus.d(idxMenu).d(idxButton).sKey
End Property

Public Property Let ButtonKey(ByVal idxMenu As Long, ByVal idxButton As Long, ByVal s As String)
mnus.d(idxMenu).d(idxButton).sKey = s
End Property

Public Property Get MenuFlags(ByVal Index As Long) As enumFakeCommandBarFlags
MenuFlags = mnus.d(Index).nFlags
End Property

Public Property Let MenuFlags(ByVal Index As Long, ByVal n As enumFakeCommandBarFlags)
With mnus.d(Index)
 If .nFlags <> n Then
  .nFlags = n
  .nFlags2 = 1
 End If
End With
End Property

Public Property Get MenuKey(ByVal Index As Long) As String
MenuKey = mnus.d(Index).sKey
End Property

Public Property Get MenuCount(ByVal Index As Long) As Long
MenuCount = mnus.nCount
End Property

Public Property Get ButtonCount(ByVal Index As Long) As Long
ButtonCount = mnus.d(Index).nCount
End Property

Public Function AddMenu(ByVal Key As String, Optional ByVal Caption As String, Optional ByVal nFlags As enumFakeCommandBarFlags) As Long
AddMenu = FakeCommandBarAddCommandBar(mnus, Key, Caption, nFlags)
End Function

Public Function AddMenuFromString(ByVal Key As String, ByVal theStr As String, Optional ByVal Caption As String, Optional ByVal nFlags As enumFakeCommandBarFlags) As Long
Dim d As typeFakeCommandBar
FakeCommandBarFromString theStr, d, PicSize
With d
 .sKey = Key
 .sCaption = Caption
 .nFlags = nFlags
End With
AddMenuFromString = FakeCommandBarAddCommandBarIndirect(mnus, d)
End Function

Public Function AddButton(ByVal MenuKey As String, Optional ByVal Key As String, Optional ByVal Caption As String, Optional ByVal ToolTipText As String, Optional ByVal nType As enumFakeButtonType, Optional ByVal nFlags As enumFakeButtonFlags, _
Optional ByVal GroupIndex As Long, Optional ByVal PicLeft As Long = -1, Optional ByVal Caption2 As String, Optional ByVal Description As String, Optional ByVal SubMenuKey As String, Optional ByVal Checked As Boolean) As Long
Dim i As Long
i = FakeCommandBarGetMenuIndex(mnus, MenuKey)
If i > 0 Then
 AddButton = FakeCommandBarAddButton(mnus.d(i), Key, Caption, ToolTipText, nType, nFlags, GroupIndex, PicLeft, Caption2, Description, SubMenuKey, Checked)
End If
End Function

Public Function AddButtonByIndex(ByVal idxMenu As Long, Optional ByVal idxButton As Long, Optional ByVal Key As String, Optional ByVal Caption As String, Optional ByVal ToolTipText As String, Optional ByVal nType As enumFakeButtonType, Optional ByVal nFlags As enumFakeButtonFlags, _
Optional ByVal GroupIndex As Long, Optional ByVal PicLeft As Long = -1, Optional ByVal Caption2 As String, Optional ByVal Description As String, Optional ByVal SubMenuKey As String, Optional ByVal Checked As Boolean) As Long
Dim i As Long
If idxMenu > 0 And idxMenu <= mnus.nCount Then
 i = InStr(1, Caption, vbTab)
 If i > 0 Then
  If Caption2 = "" Then
   Caption2 = Mid(Caption, i + 1)
   Caption = Left(Caption, i - 1)
  End If
 End If
 With mnus.d(idxMenu)
  If idxButton <= 0 Or idxButton > .nCount Then idxButton = .nCount + 1
  .nFlags2 = 1
  .nCount = .nCount + 1
  ReDim Preserve .d(1 To .nCount)
  For i = .nCount To idxButton + 1 Step -1
   .d(i) = .d(i - 1)
  Next i
  With .d(idxButton)
   .sKey = Key
   .s = Caption
   .s2 = ToolTipText
   .sTab = Caption2
   .sDesc = Description
   .sSubMenu = SubMenuKey
   .nType = nType
   .nFlags = nFlags
   .Value = Checked And 1&
   .GroupIndex = GroupIndex
   .PicLeft = PicLeft
  End With
  AddButtonByIndex = idxButton
 End With
End If
End Function

Public Sub RemoveMenu(ByVal idxMenu As Long)
Dim i As Long
If idxMenu > 0 And idxMenu <= mnus.nCount Then
 If mnus.nCount <= 1 Then
  Erase mnus.d
  mnus.nCount = 0
 Else
  For i = idxMenu + 1 To mnus.nCount
   mnus.d(i - 1) = mnus.d(i)
  Next i
  mnus.nCount = mnus.nCount - 1
  ReDim Preserve mnus.d(1 To mnus.nCount)
 End If
End If
End Sub

Public Sub DestroyMenuButtons(ByVal idxMenu As Long)
Dim i As Long
If idxMenu > 0 And idxMenu <= mnus.nCount Then
 With mnus.d(idxMenu)
  Erase .d
  .nCount = 0
  .nFlags2 = 1
 End With
End If
End Sub

Public Sub RemoveButton(ByVal idxMenu As Long, ByVal idxButton As Long)
Dim i As Long
If idxMenu > 0 And idxMenu <= mnus.nCount Then
 With mnus.d(idxMenu)
  If idxButton > 0 And idxButton <= .nCount Then
   .nFlags2 = 1
   If .nCount <= 1 Then
    Erase .d
    .nCount = 0
   Else
    For i = idxButton + 1 To .nCount
     .d(i - 1) = .d(i)
    Next i
    .nCount = .nCount - 1
    ReDim Preserve .d(1 To .nCount)
   End If
  End If
 End With
End If
End Sub

Public Sub RemoveButtonEx(ByVal idxMenu As Long, Optional ByVal idxStart As Long = 1, Optional ByVal idxEnd As Long)
Dim i As Long, j As Long
If idxMenu > 0 And idxMenu <= mnus.nCount Then
 With mnus.d(idxMenu)
  If idxStart > 0 And idxStart <= .nCount Then
   If idxEnd <= 0 Or idxEnd > .nCount Then idxEnd = .nCount
   If idxEnd < idxStart Then idxEnd = idxStart
   .nFlags2 = 1
   If idxStart = 1 And idxEnd >= .nCount Then
    Erase .d
    .nCount = 0
   Else
    j = idxEnd - idxStart + 1
    For i = idxEnd + 1 To .nCount
     .d(i - j) = .d(i)
    Next i
    .nCount = .nCount - j
    ReDim Preserve .d(1 To .nCount)
   End If
  End If
 End With
End If
End Sub

'////////general properties

Public Property Get Font() As StdFont
Set Font = cFnt.LOGFONT
End Property

Public Property Set Font(obj As StdFont)
cFnt.HighQuality = True
Set cFnt.LOGFONT = obj
If objFntB Is Nothing Then pBoldFont
End Property

Public Property Get BoldFont() As StdFont
Set BoldFont = objFntB
End Property

Public Property Set BoldFont(obj As StdFont)
Set objFntB = obj
pBoldFont
End Property

Public Property Get DropShadow() As Boolean
DropShadow = bShadow
End Property

Public Property Let DropShadow(ByVal b As Boolean)
bShadow = b
End Property

Public Property Get FakeAlpha() As Long
FakeAlpha = mnuAlpha
End Property

Public Property Let FakeAlpha(ByVal n As Long)
mnuAlpha = n
End Property

Public Property Get FakeAnimation() As Boolean
FakeAnimation = bAnim
End Property

Public Property Let FakeAnimation(ByVal b As Boolean)
bAnim = b
End Property

Private Sub pCalcBtnSize(ByVal idxMenu As Long, ByVal idxButton As Long, btn As typeFakeButton)
Dim t As MEASUREITEMSTRUCT
Dim b As Boolean
Dim fnt As CLogFont
With btn
 If .nFlags And 1& Then
  .mnuWidth = &H80000000
  .mnuHeight = &H80000000
 Else
  b = True
  If .nFlags And 8& Then
   b = False
   If bHook Then
    With t
     .CtlType = ODT_MENU
     .itemID = Val(btn.sKey)
     .itemData = InStr(1, btn.sKey, "?")
     If .itemData > 0 Then .itemData = Val(Mid(btn.sKey, .itemData + 1))
    End With
    SendMessage m_hWnd, WM_MEASUREITEM, ODT_MENU, t
    .mnuWidth = t.itemWidth
    .mnuHeight = t.itemHeight
   Else
    RaiseEvent MeasureItem(idxMenu, mnus.d(idxMenu).sKey, idxButton, btn.sKey, .mnuWidth, .mnuHeight, b)
   End If
  End If
  If b Then
   Select Case .nType
   Case 1  'separator
    .mnuWidth = -1
    .mnuWidth2 = -1
    .mnuHeight = 3
   Case 6 'column
    .mnuWidth = 3
    .mnuWidth2 = 0
    .mnuHeight = -1
   Case Else
    .mnuWidth = 28 '24 '?
    .mnuHeight = PicSize + 4
    If .nFlags And 64& Then Set fnt = cFntB Else Set fnt = cFnt
    If .s <> "" Then
     fnt.DrawTextXP hdc, .s, 0, 0, t.itemWidth, , DT_SINGLELINE Or DT_CALCRECT
     .mnuWidth = .mnuWidth + t.itemWidth + 4 '?
    End If
    If .sTab <> "" Then
     fnt.DrawTextXP hdc, .sTab, 0, 0, t.itemWidth, , DT_SINGLELINE Or DT_CALCRECT
     .mnuWidth2 = t.itemWidth + 4 '?
    Else
     .mnuWidth2 = 0
    End If
    If .nType = 5 Or (.nFlags And 4&) <> 0 Or .sSubMenu <> "" Then
     .mnuWidth2 = .mnuWidth2 + 8 '12?
    End If
   End Select
  Else
   .mnuWidth2 = 0 '????????
  End If
 End If
End With
End Sub

'toolbar mode
Private Sub pCalcBtnSize_TB(ByVal idxMenu As Long, ByVal idxButton As Long, btn As typeFakeButton)
Dim x As Long
Dim b As Boolean
Dim fnt As CLogFont
With btn
 If .nFlags And 1& Then
  .mnuWidth = &H80000000
  .mnuHeight = &H80000000
  .mnuWidth2 = &H80000000
 Else
  b = True
  If .nFlags And 8& Then
   b = False
   RaiseEvent MeasureItem(idxMenu, mnus.d(idxMenu).sKey, idxButton, btn.sKey, .mnuWidth, .mnuHeight, b)
  End If
  If b Then
   Select Case .nType
   Case 1  'separator
    If .nFlags And 256& Then 'start a row?
     .mnuWidth = -1
     .mnuWidth2 = -1
     .mnuHeight = 3
    Else
     .mnuWidth = 3
     .mnuWidth2 = 0
     .mnuHeight = -1
    End If
   Case 6 'column
    .mnuWidth = 3
    .mnuWidth2 = 0
    .mnuHeight = -1
   Case Else
    .mnuWidth = 2
    .mnuWidth2 = 0
    If .nFlags And 512& Then .mnuWidth2 = -1 'full row
    .mnuHeight = PicSize + 4
    If .PicLeft >= 0 Then .mnuWidth = .mnuWidth + PicSize + 2
    If .s <> "" And (.nFlags And 128&) = 0 Then
     If .nFlags And 64& Then Set fnt = cFntB Else Set fnt = cFnt
     fnt.DrawTextXP hdc, .s, 0, 0, x, , DT_SINGLELINE Or DT_CALCRECT
     .mnuWidth = .mnuWidth + x + 2
    End If
    If .nType = 5 Then
     .mnuWidth = .mnuWidth + 9
    ElseIf .nFlags And 4& Then
     .mnuWidth = .mnuWidth + 7
    End If
   End Select
  End If
 End If
End With
End Sub

Private Sub pCalcMenuSize(ByVal idxMenu As Long, d As typeFakeCommandBar)
Dim i As Long
Dim iStart As Long, iEnd As Long
Dim w As Long 'total menu width
Dim h As Long 'total menu height
Dim y As Long
Dim w1 As Long, w2 As Long, w3 As Long, h1 As Long
Dim b As Boolean
'check dirty
If (d.nFlags2 And 1&) = 0 Then Exit Sub
'start calc
y = 2 'y-offset
If d.nFlags And 1& Then
 'drag to make menu float
 y = y + 8 '?
End If
iStart = 1
Do Until iStart > d.nCount
 'for each column
 For i = iStart To d.nCount
  If d.d(i).nType = 6 And (d.d(i).nFlags And 1&) = 0 Then Exit For
 Next i
 iEnd = i - 1
 If d.nFlags And 2& Then 'toolbar mode
  w1 = 0 'current row width
  w2 = 0 'max column width
  w3 = 0 'current row height
  h1 = 0 'total column height
  'calc size needed
  For i = iStart To iEnd
   pCalcBtnSize_TB idxMenu, i, d.d(i)
   With d.d(i)
    If .mnuWidth >= -1 Then
     b = True
     If .nFlags And &H300& Then 'new row/full row
      'update state
      If w2 < w1 Then w2 = w1
      For w1 = iStart To i - 1
       With d.d(w1)
        If .mnuHeight = -1 Then .mnuHeight = w3
       End With
      Next w1
      h1 = h1 + w3
      w1 = 0
      w3 = 0
      'full row?
      If .mnuWidth2 = -1 Then
       .mnuLeft = w + 2
       .mnuTop = h1 + y
       If w2 < .mnuWidth Then w2 = .mnuWidth
       h1 = h1 + .mnuHeight
       b = False
      End If
     End If
     If b Then
      .mnuLeft = w + w1 + 2
      .mnuTop = h1 + y
      'update state
      w1 = w1 + .mnuWidth
      If w3 < .mnuHeight Then w3 = .mnuHeight
     End If
    End If
   End With
  Next i
  'update state
  If w2 < w1 Then w2 = w1
  For w1 = iStart To iEnd
   With d.d(w1)
    If .mnuHeight = -1 Then .mnuHeight = w3
   End With
  Next w1
  h1 = h1 + w3
  If h1 > 0 Then
   w = w + w2
   If h1 > h Then h = h1
   'update size
   For i = iStart To iEnd
    With d.d(i)
     If .mnuWidth2 = -1 Then .mnuWidth = w2
    End With
   Next i
  End If
 Else 'normal mode
  w1 = 28 '24 '? 'left width
  w2 = 0 '4 '? 'right width
  w3 = 28 '? 'total width
  h1 = 0
  'calc size needed
  For i = iStart To iEnd
   pCalcBtnSize idxMenu, i, d.d(i)
   With d.d(i)
    If .mnuWidth >= -1 Then
     .mnuLeft = w + 2
     .mnuTop = h1 + y
     h1 = h1 + .mnuHeight
     Select Case mode1
     Case 0
      If .mnuWidth > w1 Then w1 = .mnuWidth
      If .mnuWidth2 > w2 Then w2 = .mnuWidth2
     Case 1
      If .mnuWidth2 > 0 Then
       If .mnuWidth > w1 Then w1 = .mnuWidth
       If .mnuWidth2 > w2 Then w2 = .mnuWidth2
      Else
       If .mnuWidth > w3 Then w3 = .mnuWidth
      End If
     Case 2
      w1 = .mnuWidth + .mnuWidth2
      If w1 > w3 Then w3 = w1
      w1 = 0
     End Select
    End If
   End With
  Next i
  If h1 > 0 Then
   If w1 + w2 > w3 Then w3 = w1 + w2 Else w1 = w3 - w2
   w = w + w3
   If h1 > h Then h = h1
   'update size
   For i = iStart To iEnd
    With d.d(i)
     If .mnuWidth >= -1 Then
      .mnuWidth = w3
      .mnuWidth2 = w1
     End If
    End With
   Next i
  End If
 End If
 'next
 iEnd = iEnd + 1
 iStart = iEnd + 1
 If iEnd <= d.nCount Then 'v-separator
  pCalcBtnSize idxMenu, iEnd, d.d(iEnd)
  With d.d(iEnd)
   .mnuLeft = w + 2
   .mnuTop = y
   w = w + .mnuWidth
  End With
 End If
Loop
'over
If h > 0 Then
 d.w = w + 4
 d.h = h + 4
 d.nFlags2 = 0&
Else
 d.w = 48
 d.h = 24
 d.nFlags2 = 2&
End If
If d.nFlags And 1& Then d.h = d.h + 8
End Sub

Public Property Get CalcSizeMethod() As enumFakeMenuCalcSizeMethod
CalcSizeMethod = mode1
End Property

Public Property Let CalcSizeMethod(ByVal n As enumFakeMenuCalcSizeMethod)
mode1 = n
pDirty
End Property

Public Sub SetOwner(ByVal hwnd As Long)
hwndOwner = hwnd
End Sub

Private Function pHitTest(d As typeFakeCommandBar, ByVal x As Long, ByVal y As Long) As Long
Dim i As Long
Dim xx As Long, yy As Long
For i = 1 To d.nCount
 With d.d(i)
  If (.nFlags And 1&) = 0 And .nType <> 1 And .nType <> 6 Then
   xx = x - .mnuLeft
   yy = y - .mnuTop
   If xx >= 0 And yy >= 0 And xx < .mnuWidth And yy < .mnuHeight Then
    pHitTest = i
    Exit For
   End If
  End If
 End With
Next i
End Function

Public Function Click(ByVal MenuKey As String, ByVal ButtonKey As String) As Boolean
Dim i As Long, j As Long
i = FakeCommandBarGetMenuIndex(mnus, MenuKey)
If i > 0 Then
 j = FindButton(i, ButtonKey)
 If j > 0 Then Click = ClickByIndex(i, j)
End If
End Function

Public Function ClickByIndex(ByVal idxMenu As Long, ByVal idxButton As Long) As Boolean
Dim i As Long, j As Long, m As Long
Dim b As Boolean
Dim tMI As MENUINFO
With mnus.d(idxMenu)
 m = .nCount
 With .d(idxButton)
  If (.nFlags And 3&) = 0 And (.sSubMenu = "" Or .nType = 5) Then
   Select Case .nType
   Case 0, 5 'button, split
    b = True
   Case 2 'check
    .Value = (.Value = 0) And 1&
    b = True
   Case 3 'option
    j = .GroupIndex
    With mnus.d(idxMenu)
     For i = 1 To m
      With .d(i)
       If .nType = 3 And .GroupIndex = j Then
        .Value = (i = idxButton) And 1&
       End If
      End With
     Next i
    End With
    b = True
   Case 4 'optnull
    If .Value Then
     .Value = 0
    Else
     j = .GroupIndex
     With mnus.d(idxMenu)
      For i = 1 To m
       With .d(i)
        If .nType = 4 And .GroupIndex = j Then
         .Value = (i = idxButton) And 1&
        End If
       End With
      Next i
     End With
    End If
    b = True
   End Select
   i = .Value
  End If
 End With
End With
If b Then
 'click a button successful!
 HlKey = mnus.d(idxMenu).d(idxButton).sKey
 If wndc > 0 Then pDestroyAllWindow
 'raise event
 If bHook Then
  b = False
  i = Val(mnus.d(idxMenu).sKey)
  j = Val(HlKey)
  RaiseEvent HookedClick(i, j, idxButton, b)
  If bNotify And Not b Then
   '///update state ????????
   'HiliteMenuItem m_hWnd, i, idxButton - 1, MF_BYPOSITION Or MF_HILITE
   '///
   tMI.cbSize = Len(tMI)
   tMI.fMask = MIM_STYLE
   GetMenuInfo j, tMI
   If tMI.dwStyle And MNS_NOTIFYBYPOS Then
    'wParam
    'Specifies the zero-based index of the item selected.
    'Windows 98/Me: The high word is the zero-based index of the item selected. The low word is the item ID.
    SendMessage m_hWnd, WM_MENUCOMMAND, idxButton - 1, ByVal i
   Else
    SendMessage m_hWnd, WM_COMMAND, j, ByVal 0
   End If
  End If
  b = True
 Else
  j = i
  RaiseEvent Click(idxMenu, mnus.d(idxMenu).sKey, idxButton, HlKey, i)
  i = i And &HFF&
  If i <> j Then
   If idxMenu > 0 And idxMenu <= mnus.nCount Then
    With mnus.d(idxMenu)
     If idxButton > 0 And idxButton <= .nCount Then .d(idxButton).Value = i
    End With
   End If
  End If
 End If
End If
ClickByIndex = b
End Function

Private Sub p1_Click(Index As Integer)
Dim idx As Long
Dim idxHl As Long
Dim i As Long, j As Long, m As Long
Dim b As Boolean
If Index > 0 And Index <= wndc Then
 With wnds(Index)
  If .idx > 0 And .idx <= mnus.nCount Then
   idx = .idx
   idxHl = .idxHl
  End If
 End With
 If idx > 0 Then
  With mnus.d(idx)
   If idxHl > 0 And idxHl <= .nCount Then b = True
  End With
  If b Then ClickByIndex idx, idxHl
 End If
End If
End Sub

Private Sub p1_MouseDown(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim idxHl As Long
If Index > 0 And Index <= wndc Then
 With wnds(Index)
  If .idx > 0 And .idx <= mnus.nCount Then
   'hit test
   idxHl = pHitTest(mnus.d(.idx), x, y)
  End If
 End With
 'unpopup sub menu
 If idxHl = 0 And Index < wndc Then
  pUnpopupSubMenu Index
 End If
End If
End Sub

Private Sub p1_MouseMove(Index As Integer, Button As Integer, Shift As Integer, x As Single, y As Single)
Dim r As RECT
Dim idx As Long
Dim idxHl As Long
Dim s As String, mp As Long
Dim bChanged As Boolean, idxMenu As Long
Dim bFloat As Boolean
If Index > 0 And Index <= wndc Then
 With wnds(Index)
  If .idx > 0 And .idx <= mnus.nCount Then
   idx = .idx
   idxHl = .idxHl
   'hit test
   r.Left = x
   r.Top = y
   If (mnus.d(.idx).nFlags And 1&) <> 0 And r.Left >= 2 And r.Left < .w - 2 And r.Top >= 2 And r.Top < 9 Then
    idxHl = &H80000000
    'make this menu float?
    If Button = 1 Then bFloat = True
   Else
    idxHl = pHitTest(mnus.d(.idx), r.Left, r.Top)
   End If
   '????????
   If HlWndIndex <> Index Then
    If HlWndIndex > 0 And HlWndIndex <= wndc Then
     p1_MouseMove (HlWndIndex), (Button), (Shift), (-1), (-1) '???
    End If
    HlWndIndex = Index
   End If
   HlIndex = idxHl
   'changed?
   If idxHl <> .idxHl Then
    'mousemove event
    If bHook Then
'     If bNotify Then
'      r.Left = Val(mnus.d(.idx).sKey)
'      If idxHl > 0 Then
'       r.Top = Val(mnus.d(.idx).d(idxHl).sKey)
'       SendMessage m_hWnd, WM_MENUSELECT, r.Left, ByVal r.Left
'      Else
'       SendMessage m_hWnd, WM_MENUSELECT, 0, ByVal r.Left
'      End If
'     End If
    Else
     If idxHl > 0 Then
      RaiseEvent MouseMove(.idx, mnus.d(.idx).sKey, idxHl, mnus.d(.idx).d(idxHl).sKey, mnus.d(.idx).d(idxHl).sDesc)
     Else
      RaiseEvent MouseMove(0, "", 0, "", "")
     End If
    End If
    'tooltiptext
    If idxHl = &H80000000 Then
     s = "Drag to make this menu float"
'     s = "拖动可使此菜单浮动" ':-3
     mp = vbSizeAll
    Else
     If idxHl > 0 Then s = mnus.d(.idx).d(idxHl).s2
     mp = vbDefault
    End If
    With p1(Index)
     .ToolTipText = s
     .MousePointer = mp
    End With
    '???
    If Index = wndc Or idxHl <> 0 Then
     bChanged = True
     .idxHl = idxHl
     pRedraw Index, True
     'check sub menu
     If idxHl > 0 Then
      With mnus.d(.idx).d(idxHl)
       If (.nFlags And 3&) = 0 And .nType <> 1 And .nType <> 6 Then
        If .sSubMenu <> "" Then
         idxMenu = FakeCommandBarGetMenuIndex(mnus, .sSubMenu)
         If idxMenu > 0 Then 'found!
          GetWindowRect p1(Index).hwnd, r
          r.Left = r.Left + .mnuLeft
          r.Top = r.Top + .mnuTop
          r.Right = .mnuWidth
          r.Bottom = .mnuHeight
         End If
        End If
       End If
      End With
     End If
    End If
    'TODO:other
   End If
  End If
 End With
 'popup/unpopup sub menu TODO:delay
 If bChanged Then
  If Index < wndc Then 'unpopup sub menu
   pUnpopupSubMenu Index
  End If
  If idxMenu > 0 Then 'show sub menu
   pPopupSubMenu idxMenu, r.Left, r.Top, r.Right, r.Bottom, mnus.d(idx).nFlags And 2&
  End If
 End If
End If
'make menu float
If bFloat Then
 GetWindowRect p1(Index).hwnd, r
 pDestroyAllWindow
 RaiseEvent MakeMenuFloat(idx, mnus.d(idx).sKey, r.Left + x, r.Top + y, r.Left, r.Top, r.Right - r.Left, r.Bottom - r.Top)
End If
End Sub

Private Sub p1_Paint(Index As Integer)
If Index > 0 And Index <= wndc Then
 If Not wnds(Index).bm Is Nothing Then
  wnds(Index).bm.PaintPicture p1(Index).hdc
 End If
End If
End Sub

Private Sub t1_Timer()
Static k1 As Long, k2 As Long
Dim i As Long, m As Long
Dim p As POINTAPI
Dim b As Boolean
If wndc > 0 Then
 GetCursorPos p
 'check close menu
 If g_FakeMenuUserData <> ObjPtr(Me) Then b = True
 If Not b Then
  If bHook Then
   If k1 = &H80000000 Then
    k1 = GetKeyState(1)
    k2 = GetKeyState(2)
   Else
    i = GetKeyState(1)
    m = GetKeyState(2)
    If i <> k1 Or m <> k2 Then
     k1 = i
     k2 = m
     i = 1
    Else
     i = 0
    End If
   End If
  End If
  i = i Or ((GetAsyncKeyState(1) Or GetAsyncKeyState(2)) And &H1&)
  If i Then
   m = WindowFromPoint(p.x, p.y)
   b = True
   For i = 1 To p1.UBound
    If m = p1(i).hwnd Then
     b = False
     Exit For
    End If
   Next i
  End If
 End If
 If Not b Then
  If GetAsyncKeyState(vbKeyMenu) = &H8001 Then b = True
  If GetAsyncKeyState(vbKeyEscape) = &H8001 Then b = True
 End If
 If Not b Then
  If GetActiveWindow <> OldActiveWindow Then b = True
 End If
 If b Then
  pDestroyAllWindow
  Exit Sub
 End If
 'check highlight
 If HlWndIndex > 0 And HlWndIndex <= wndc Then
  With wnds(HlWndIndex)
   b = p.x >= .x And p.x < .x + .w And p.y >= .y And p.y < .y + .h
  End With
  If Not b Then
   p1_MouseMove (HlWndIndex), (0), (0), (-1), (-1)
   HlWndIndex = -1
   HlIndex = -1
  End If
 End If
 'menu animation
 If bAnim Then
  If mnuAlpha > 0 And mnuAlpha < 255 Then m = mnuAlpha Else m = 256&
  For i = 1 To wndc
   With wnds(i)
    If .nAlpha < m Then
     If p.x > .x And p.x < .x + .w - 1 And p.y > .y And p.y < .y + .h - 1 Then
      .nAlpha = m
     Else
      .nAlpha = .nAlpha + AnimationStep
      If .nAlpha > m Then .nAlpha = m
     End If
     If .nAlpha >= 255 Then Set .bmBack = Nothing
     pRedraw i, True
    End If
   End With
  Next i
 End If
Else
 k1 = &H80000000
End If
End Sub

Private Sub tmrShadow_Timer()
If EnableTooltipDropShadow Then tmrShadow.Enabled = False
End Sub

Private Sub UserControl_InitProperties()
cFnt.HighQuality = True
Set cFnt.LOGFONT = UserControl.Font
pBoldFont
bShadow = True
bAnim = True
bNC = True
pInit
End Sub

Private Sub pBoldFont()
Dim obj As IFont, obj2 As StdFont
If objFntB Is Nothing Then
 Set obj = cFnt.LOGFONT
 obj.Clone obj2
 obj2.Bold = True
 cFntB.HighQuality = True
 Set cFntB.LOGFONT = obj2
Else
 cFntB.HighQuality = True
 Set cFntB.LOGFONT = objFntB
End If
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
 mode1 = .ReadProperty("CalcSizeMethod", 0)
 cFnt.HighQuality = True
 Set cFnt.LOGFONT = .ReadProperty("Font", UserControl.Font)
 Set objFntB = .ReadProperty("BoldFont", Nothing)
 bShadow = .ReadProperty("DropShadow", True)
 bAnim = .ReadProperty("FakeAnimation", True)
 mnuAlpha = .ReadProperty("FakeAlpha", 0)
 bNC = .ReadProperty("FakeMenuNoConnect", True)
End With
pBoldFont
pInit
End Sub

Private Sub pInit()
hwndOwner = ContainerHwnd
t1.Enabled = Ambient.UserMode
tmrShadow.Enabled = Ambient.UserMode
'other things
End Sub

Private Sub pInitWindow(ByVal hwnd As Long)
Dim i As Long
i = GetWindowLong(hwnd, GWL_EXSTYLE) Or WS_EX_TOOLWINDOW Or WS_EX_TOPMOST
SetWindowLong hwnd, GWL_EXSTYLE, i
SetWindowPos hwnd, 0, 0, 0, 0, 0, SWP_FRAMECHANGED Or SWP_NOMOVE Or SWP_NOSIZE Or SWP_NOZORDER
SetWindowLong hwnd, GWL_HWNDPARENT, hwndOwner
SetParent hwnd, 0
If bShadow Then
 EnableDropShadow hwnd
Else
 DisbleDropShadow hwnd
End If
End Sub

Friend Sub CreateIndirect(d As typeFakeCommandBars)
pDestroyAllWindow
mnus = d
pDirty
End Sub

Private Sub pDirty()
Dim i As Long
For i = 1 To mnus.nCount
 mnus.d(i).nFlags2 = 1
Next i
End Sub

Private Sub pDestroyAllWindow()
Dim i As Long
For i = p1.UBound To 1 Step -1
 p1(i).Visible = False
 Unload p1(i)
Next i
Erase wnds
wndc = 0
HlWndIndex = -1
HlIndex = -1
End Sub

Public Sub Destroy()
Dim d As typeFakeCommandBars
pDestroyAllWindow
mnus = d
End Sub

Private Sub UserControl_Terminate()
Destroy
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
 .WriteProperty "CalcSizeMethod", mode1, 0
 .WriteProperty "Font", cFnt.LOGFONT, UserControl.Font
 .WriteProperty "BoldFont", objFntB, Nothing
 .WriteProperty "DropShadow", bShadow, True
 .WriteProperty "FakeAnimation", bAnim, True
 .WriteProperty "FakeAlpha", mnuAlpha, 0
 .WriteProperty "FakeMenuNoConnect", bNC, True
End With
End Sub

Private Sub pMenuPos_CalcCY(ByRef x As Long, ByVal w As Long, ByVal ww As Long, ByVal w1 As Long, ByRef X1 As Long, ByRef X2 As Long, Optional ByVal bReverse As Boolean)
X1 = x
If bReverse Then
 If x - w + w1 >= 0 Then x = x - w + w1 Else If x > ww - w Then x = 0
Else
 If x > ww - w Then
  x = x - w + w1
  If x < 0 Then x = ww - w
 End If
End If
X1 = X1 - x + 1
X2 = X1 + w1 - 2
If X1 < 1 Then X1 = 1
If X2 > w - 1 Then X2 = w - 1
End Sub

Private Function pMenuPos_CheckDown(ByRef x As Long, ByRef y As Long, ByVal w As Long, ByVal h As Long, ByVal ww As Long, ByVal hh As Long, ByVal w1 As Long, ByVal h1 As Long, ByRef FS As Long, ByRef X1 As Long, ByRef X2 As Long, Optional ByVal nPopupDirection As Long, Optional ByVal bNoConnect As Boolean) As Boolean
If y + h1 - 1 <= hh - h Then
 y = y + h1 - 1
 pMenuPos_CalcCY x, w, ww, w1, X1, X2, nPopupDirection = 3
 If bNoConnect Then FS = -1 Else FS = 0
 pMenuPos_CheckDown = True
End If
End Function

Private Function pMenuPos_CheckUp(ByRef x As Long, ByRef y As Long, ByVal w As Long, ByVal h As Long, ByVal ww As Long, ByVal hh As Long, ByVal w1 As Long, ByVal h1 As Long, ByRef FS As Long, ByRef X1 As Long, ByRef X2 As Long, Optional ByVal nPopupDirection As Long, Optional ByVal bNoConnect As Boolean) As Boolean
If y - h + 1 >= 0 Then
 y = y - h + 1
 pMenuPos_CalcCY x, w, ww, w1, X1, X2, nPopupDirection = 3
 If bNoConnect Then FS = -1 Else FS = 1
 pMenuPos_CheckUp = True
End If
End Function

Private Function pMenuPos_CheckRight(ByRef x As Long, ByRef y As Long, ByVal w As Long, ByVal h As Long, ByVal ww As Long, ByVal hh As Long, ByVal w1 As Long, ByVal h1 As Long, ByRef FS As Long, ByRef X1 As Long, ByRef X2 As Long, Optional ByVal nPopupDirection As Long, Optional ByVal bNoConnect As Boolean) As Boolean
If x + w1 - 1 <= ww - w Then
 x = x + w1 - 1
 pMenuPos_CalcCY y, h, hh, h1, X1, X2, nPopupDirection = 1
 If bNoConnect Then FS = -1 Else FS = 2
 pMenuPos_CheckRight = True
End If
End Function

Private Function pMenuPos_CheckLeft(ByRef x As Long, ByRef y As Long, ByVal w As Long, ByVal h As Long, ByVal ww As Long, ByVal hh As Long, ByVal w1 As Long, ByVal h1 As Long, ByRef FS As Long, ByRef X1 As Long, ByRef X2 As Long, Optional ByVal nPopupDirection As Long, Optional ByVal bNoConnect As Boolean) As Boolean
If x - w + 1 >= 0 Then
 x = x - w + 1
 pMenuPos_CalcCY y, h, hh, h1, X1, X2, nPopupDirection = 1
 If bNoConnect Then FS = -1 Else FS = 3
 pMenuPos_CheckLeft = True
End If
End Function

Private Sub pMenuPos_Calc(ByRef x As Long, ByRef y As Long, ByVal w As Long, ByVal h As Long, ByVal ww As Long, ByVal hh As Long, ByRef FS As Long, Optional ByVal nPopupDirection As Long)
FS = -1
'calc x
If nPopupDirection = 3 Then 'left
 If x - w + 1 >= 0 Then x = x - w + 1 Else If x > ww - w Then x = 0
Else 'right
 If x > ww - w Then
  x = x - w + 1
  If x < 0 Then x = ww - w
 End If
End If
'calc y
If nPopupDirection = 1 Then 'up
 If y - h + 1 >= 0 Then y = y - h + 1 Else If y > hh - h Then y = 0
Else 'down
 If y > hh - h Then
  y = y - h + 1
  If y < 0 Then y = hh - h
 End If
End If
End Sub

Public Function HasMenu(ByVal Key As String) As Boolean
Dim i As Long
For i = 1 To mnus.nCount
 If Key = mnus.d(i).sKey Then Exit For
Next i
HasMenu = i <= mnus.nCount
End Function

Private Sub pUnpopupSubMenu(ByVal Index As Long)
Dim i As Long
If Index < wndc Then
 For i = p1.UBound To Index + 1 Step -1
  p1(i).Visible = False
  Unload p1(i)
 Next i
 wndc = Index
 ReDim Preserve wnds(1 To wndc)
 'Debug.Print "Destroy"; Index + 1
End If
End Sub

Private Sub pCreateBitmap(d As typeFakeMenuWindow)
Dim hd As Long
With d
 .nAlpha = AnimationStep
 Set .bm = New cDIBSection
 .bm.Create .w, .h
 If bAnim Or (mnuAlpha > 0 And mnuAlpha < 255) Then
  If mnuAlpha > 0 And mnuAlpha < 255 And (.nAlpha > mnuAlpha Or Not bAnim) Then .nAlpha = mnuAlpha
  Set .bmBack = New cDIBSection
  .bmBack.Create .w, .h
  hd = GetWindowDC(0)
  BitBlt .bmBack.hdc, 0, 0, .w, .h, hd, .x, .y, vbSrcCopy
  ReleaseDC 0, hd
 Else
  .nAlpha = 256&
 End If
End With
End Sub

Private Sub pPopupSubMenu(ByVal idxMenu As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal bToolBar As Boolean)
Dim FS As Long, X1 As Long, X2 As Long
Dim ww As Long, hh As Long
Dim b As Boolean
'new:raise event
RaiseEvent BeforeShowMenu(idxMenu)
'calc size
pCalcMenuSize idxMenu, mnus.d(idxMenu)
'calc position
ww = Screen.Width / Screen.TwipsPerPixelX
hh = Screen.Height / Screen.TwipsPerPixelY
With mnus.d(idxMenu)
 If bToolBar Then 'toolbar
  'direction=down
  If pMenuPos_CheckDown(x, y, .w, .h, ww, hh, w, h, FS, X1, X2) Then
  ElseIf pMenuPos_CheckUp(x, y, .w, .h, ww, hh, w, h, FS, X1, X2) Then
  ElseIf pMenuPos_CheckRight(x, y, .w, .h, ww, hh, w, h, FS, X1, X2) Then
  ElseIf pMenuPos_CheckLeft(x, y, .w, .h, ww, hh, w, h, FS, X1, X2) Then
  Else
   pMenuPos_Calc x, y, .w, .h, ww, hh, FS
  End If
  w = .w
  h = .h
 Else
  If bNC Then
   x = x - 2
   w = w + 4
   b = True
  End If
  'direction=right
  If pMenuPos_CheckRight(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, , b) Then
  ElseIf pMenuPos_CheckLeft(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, , b) Then
  Else
   If b Then
    x = x + 2
    w = w - 4
    y = y - 2
    h = h + 4
   End If
   If pMenuPos_CheckDown(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, , b) Then
   ElseIf pMenuPos_CheckUp(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, , b) Then
   Else
    If b Then
     y = y + 2
     h = h - 4
    End If
    pMenuPos_Calc x, y, .w, .h, ww, hh, FS
   End If
  End If
  w = .w
  h = .h
 End If
End With
'allocate menu data
wndc = wndc + 1
ReDim Preserve wnds(1 To wndc)
With wnds(wndc)
 .idx = idxMenu
 .x = x
 .y = y
 .w = w
 .h = h
 .FS = FS
 .X1 = X1
 .X2 = X2
 .nDef = 0
End With
pCreateBitmap wnds(wndc)
pRedraw wndc
'create menu window
pCreateWindow wndc
'show menu TODO:
p1(wndc).Visible = True
End Sub

Public Sub UnpopupMenu()
If wndc > 0 Then pDestroyAllWindow
End Sub

Public Sub PopupMenu(ByVal Key As String, Optional ByVal x As Long = &H80000000, Optional ByVal y As Long = &H80000000, Optional ByVal DefaultMenu As Long = -1)
PopupMenuEx Key, x, y, , , DefaultMenu
End Sub

Public Sub PopupMenuEx(ByVal Key As String, Optional ByVal x As Long = &H80000000, Optional ByVal y As Long = &H80000000, Optional ByVal w As Long, Optional ByVal h As Long, Optional ByVal DefaultMenu As Long = -1, Optional ByVal bReturnCmd As Boolean, Optional ByVal nPopupDirection As Long, Optional ByVal bNoConnect As Boolean, Optional ByVal UserData As Long, Optional ByRef sReturnCmd As String, Optional ByVal bSetFocus As Boolean)
Dim p As POINTAPI
Dim FS As Long, X1 As Long, X2 As Long
Dim ww As Long, hh As Long
Dim i As Long
'////////????
GetAsyncKeyState 1
GetAsyncKeyState 2
'////////
If bReturnCmd Then
 sReturnCmd = ""
 HlKey = ""
End If
'get cursor pos
GetCursorPos p
If x = &H80000000 Or y = &H80000000 Then
 x = p.x
 y = p.y
End If
'destroy menu
pDestroyAllWindow
'find menu
For i = 1 To mnus.nCount
 If Key = mnus.d(i).sKey Then Exit For
Next i
If i > mnus.nCount Then Exit Sub
'set user data
nUserData = UserData
g_FakeMenuUserData = ObjPtr(Me)
OldActiveWindow = GetActiveWindow
'new:raise event
RaiseEvent BeforeShowMenu(i)
'calc size
pCalcMenuSize i, mnus.d(i)
'calc position
ww = Screen.Width / Screen.TwipsPerPixelX
hh = Screen.Height / Screen.TwipsPerPixelY
With mnus.d(i)
 If w > 1 And h > 1 Then 'has button size
  If bNoConnect Then
   x = x - 1
   y = y - 1
   w = w + 2
   h = h + 2
  End If
  Select Case nPopupDirection
  Case 1 'up
   If pMenuPos_CheckUp(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, nPopupDirection, bNoConnect) Then
   ElseIf pMenuPos_CheckDown(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, nPopupDirection, bNoConnect) Then
   ElseIf pMenuPos_CheckRight(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, nPopupDirection, bNoConnect) Then
   ElseIf pMenuPos_CheckLeft(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, nPopupDirection, bNoConnect) Then
   Else
    pMenuPos_Calc x, y, .w, .h, ww, hh, FS, nPopupDirection
   End If
  Case 2 'right
   If pMenuPos_CheckRight(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, nPopupDirection, bNoConnect) Then
   ElseIf pMenuPos_CheckLeft(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, nPopupDirection, bNoConnect) Then
   ElseIf pMenuPos_CheckDown(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, nPopupDirection, bNoConnect) Then
   ElseIf pMenuPos_CheckUp(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, nPopupDirection, bNoConnect) Then
   Else
    pMenuPos_Calc x, y, .w, .h, ww, hh, FS, nPopupDirection
   End If
  Case 3 'left
   If pMenuPos_CheckLeft(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, nPopupDirection, bNoConnect) Then
   ElseIf pMenuPos_CheckRight(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, nPopupDirection, bNoConnect) Then
   ElseIf pMenuPos_CheckDown(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, nPopupDirection, bNoConnect) Then
   ElseIf pMenuPos_CheckUp(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, nPopupDirection, bNoConnect) Then
   Else
    pMenuPos_Calc x, y, .w, .h, ww, hh, FS, nPopupDirection
   End If
'  Case 4 'RTL down TODO:
'  Case 5 'RTL up TODO:
'  Case 6 'RTL right TODO:
'  Case 7 'RTL left TODO:
  Case Else 'down
   If pMenuPos_CheckDown(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, nPopupDirection, bNoConnect) Then
   ElseIf pMenuPos_CheckUp(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, nPopupDirection, bNoConnect) Then
   ElseIf pMenuPos_CheckRight(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, nPopupDirection, bNoConnect) Then
   ElseIf pMenuPos_CheckLeft(x, y, .w, .h, ww, hh, w, h, FS, X1, X2, nPopupDirection, bNoConnect) Then
   Else
    pMenuPos_Calc x, y, .w, .h, ww, hh, FS, nPopupDirection
   End If
  End Select
 Else 'no button size
  pMenuPos_Calc x, y, .w, .h, ww, hh, FS, nPopupDirection
 End If
 w = .w
 h = .h
End With
'allocate menu data
ReDim wnds(1 To 1)
wndc = 1
With wnds(1)
 .idx = i
 .x = x
 .y = y
 .w = w
 .h = h
 .FS = FS
 .X1 = X1
 .X2 = X2
 .nDef = DefaultMenu
End With
pCreateBitmap wnds(1)
pRedraw 1
'create menu window
pCreateWindow 1
'show menu TODO:
p1(1).Visible = True
'////////???????? dreadful code
If bSetFocus Then
 SetCursorPos x, y
 mouse_event MOUSEEVENTF_ABSOLUTE Or MOUSEEVENTF_LEFTDOWN Or MOUSEEVENTF_LEFTUP, x, y, 0, 0
 GetAsyncKeyState 1
 GetAsyncKeyState 2
 SetCursorPos p.x, p.y
End If
'////////
'TODO:
If bReturnCmd Then
 'custom message loop !!!
 Dim tMsg As typeMSG
 Do While GetMessage(tMsg, 0, 0, 0)
  TranslateMessage tMsg
  DispatchMessage tMsg
  'end of menu?
  If wndc <= 0 Then Exit Do
 Loop
 sReturnCmd = HlKey
End If
'TODO:
End Sub

Public Property Get MenuWindowCount() As Long
MenuWindowCount = wndc
End Property

Private Sub pCreateWindow(ByVal Index As Long)
Dim i As Long
Dim hwd As Long
If Index > 0 And Index <= wndc Then
 Do While Index > p1.UBound
  i = p1.UBound + 1
  Load p1(i)
  p1(i).Visible = False
 Loop
 With p1(Index)
  hwd = .hwnd
  pInitWindow hwd
  With wnds(Index)
   SetWindowPos hwd, HWND_TOPMOST, .x, .y, .w, .h, SWP_NOACTIVATE
  End With
  .MousePointer = vbDefault
  .ToolTipText = ""
 End With
 'TODO:
 'Debug.Print "Create"; Index
End If
End Sub

Private Sub pDrawDropdown(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal clr As Long)
Dim i As Long, j As Long
For i = 0 To 4
 For j = -i To i
  SetPixelV hdc, x - i, y + j, clr
 Next j
Next i
End Sub

Private Sub pDrawDropdown_TB(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal clr As Long)
Dim i As Long, j As Long
For j = 0 To 2
 For i = -j To j
  SetPixelV hdc, x + i, y + 1 - j, clr
 Next i
Next j
End Sub

Private Sub pDrawCheck(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal clr As Long, ByVal nType As Long)
Dim i As Long, j As Long
If nType = 3 Or nType = 4 Then
 For i = 0 To 2
  For j = 0 To 2
   If i + j < 4 Then
    SetPixelV hdc, x + i, y + j, clr
    SetPixelV hdc, x + i, y - 1 - j, clr
    SetPixelV hdc, x - 1 - i, y + j, clr
    SetPixelV hdc, x - 1 - i, y - 1 - j, clr
   End If
  Next j
 Next i
Else
 y = y + 1
 For i = -3 To 3
  j = i + 1
  If j < 0 Then j = -j
  SetPixelV hdc, x + i, y - j, clr
  SetPixelV hdc, x + i, y + 1 - j, clr
  SetPixelV hdc, x + i, y + 2 - j, clr
 Next i
End If
End Sub

Private Sub pDrawCloseIcon(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal clr As Long)
Dim r As RECT, hbr As Long
r.Left = x - 1
r.Top = y - 1
r.Right = x + 2
r.Bottom = y + 2
hbr = CreateSolidBrush(clr)
FillRect hdc, r, hbr
DeleteObject hbr
For hbr = 2 To 4
 SetPixelV hdc, x + hbr, y + hbr, clr
 SetPixelV hdc, x + hbr - 1, y + hbr, clr
 SetPixelV hdc, x + hbr, y + hbr - 1, clr
 SetPixelV hdc, x - hbr, y + hbr, clr
 SetPixelV hdc, x - hbr + 1, y + hbr, clr
 SetPixelV hdc, x - hbr, y + hbr - 1, clr
 SetPixelV hdc, x + hbr, y - hbr, clr
 SetPixelV hdc, x + hbr - 1, y - hbr, clr
 SetPixelV hdc, x + hbr, y - hbr + 1, clr
 SetPixelV hdc, x - hbr, y - hbr, clr
 SetPixelV hdc, x - hbr + 1, y - hbr, clr
 SetPixelV hdc, x - hbr, y - hbr + 1, clr
Next hbr
End Sub

Private Sub pDrawMinIcon(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal clr As Long)
Dim r As RECT, hbr As Long
r.Left = x - 4
r.Top = y + 5
r.Right = x + 3
r.Bottom = y + 7
hbr = CreateSolidBrush(clr)
FillRect hdc, r, hbr
DeleteObject hbr
End Sub

Private Sub pDrawMaxIcon(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal clr As Long)
Dim r As RECT, hbr As Long
r.Left = x - 5
r.Top = y - 5
r.Right = x + 5
r.Bottom = y + 5
hbr = CreateSolidBrush(clr)
FrameRect hdc, r, hbr
r.Bottom = y - 3
FillRect hdc, r, hbr
DeleteObject hbr
End Sub

Private Sub pDrawRestoreIcon(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal clr As Long)
Dim r As RECT, hbr As Long
r.Left = x - 5
r.Top = y - 2
r.Right = x + 3
r.Bottom = y + 5
hbr = CreateSolidBrush(clr)
FrameRect hdc, r, hbr
r.Bottom = y
r.Right = x + 3
FillRect hdc, r, hbr
r.Left = x - 3
r.Top = y - 5
r.Right = x + 5
r.Bottom = y - 3
FillRect hdc, r, hbr
DeleteObject hbr
SetPixelV hdc, x - 3, y - 3, clr
SetPixelV hdc, x + 4, y - 3, clr
SetPixelV hdc, x + 4, y - 2, clr
SetPixelV hdc, x + 4, y - 1, clr
SetPixelV hdc, x + 4, y, clr
SetPixelV hdc, x + 3, y + 1, clr
SetPixelV hdc, x + 4, y + 1, clr
End Sub

Private Sub pDrawGripper(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal hbrBorder As Long, Optional ByVal bHighlight As Boolean)
Dim r As RECT
r.Left = x
r.Top = y
r.Right = x + w
r.Bottom = y + 7
If bHighlight Then
 GradientFillRect hdc, x, y, r.Right, r.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
 If hbrBorder <> 0 Then FrameRect hdc, r, hbrBorder
Else
 GradientFillRect hdc, x, y, r.Right, r.Bottom, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_V
End If
'draw gripper
r.Right = ((w \ 2&) And &HFFFFFFFC) - 1
If r.Right < 15 Then r.Right = 15
x = x + (w - r.Right) \ 2
y = y + 3
Do
 '////stupid code
 SetPixelV hdc, x - 1, y - 1, d_Gripper
 SetPixelV hdc, x, y - 1, d_Gripper
 SetPixelV hdc, x - 1, y, d_Gripper
 SetPixelV hdc, x, y, d_Gripper
 SetPixelV hdc, x + 1, y, vbWhite
 SetPixelV hdc, x, y + 1, vbWhite
 SetPixelV hdc, x + 1, y + 1, vbWhite
 '////
 x = x + 4
 r.Right = r.Right - 4
Loop While r.Right > 0
End Sub

Private Sub pRedraw(ByVal Index As Long, Optional ByVal bPaint As Boolean)
Dim bm As cDIBSection
Dim t As DRAWITEMSTRUCT, hbm As Long, hMenu As Long
Dim i As Long
Dim hbr As Long, hbrBorder As Long
Dim r As RECT
Dim w As Long, h As Long, idxHl As Long, nDef As Long, bDef As Boolean
Dim bToolBar As Boolean
Dim bDrawRow As Boolean, bDrawDropdown As Boolean, bSelected As Boolean
Dim bDrawPic As Boolean, nOwnerDraw As Long
Dim n As Long
Dim fnt As CLogFont
If Index > 0 And Index <= wndc Then
 With wnds(Index)
  If .idx > 0 And .idx <= mnus.nCount Then
   w = .w
   h = .h
   idxHl = .idxHl
   nDef = .nDef
   If .bm Is Nothing Then
    Set .bm = New cDIBSection
    .bm.Create w, h
   End If
   Set bm = .bm
   'draw back
   hbrBorder = CreateSolidBrush(d_Border)
   hbr = CreateSolidBrush(d_Bg)
   r.Right = w
   r.Bottom = h
   FillRect bm.hdc, r, hbr
   FrameRect bm.hdc, r, hbrBorder
   If .X1 < .X2 And .FS >= 0 And .FS <= 3 Then
    If .FS < 2 Then
     r.Left = .X1
     r.Right = .X2
     If .FS = 0 Then
      r.Top = 0
      r.Bottom = 1
     Else
      r.Top = r.Bottom - 1
     End If
    Else
     r.Top = .X1
     r.Bottom = .X2
     If .FS = 2 Then
      r.Left = 0
      r.Right = 1
     Else
      r.Left = r.Right - 1
     End If
    End If
    FillRect bm.hdc, r, hbr
   End If
   DeleteObject hbr
   With mnus.d(.idx)
    If bHook Then hMenu = Val(.sKey)
    bToolBar = .nFlags And 2&
    'draw gripper
    If .nFlags And 1& Then
     pDrawGripper bm.hdc, 2, 2, w - 4, hbrBorder, idxHl = &H80000000
    End If
    If .nFlags2 And 2& Then 'nothing
     If .nFlags And 1& Then n = 8 Else n = 0
     cFnt.DrawTextXP bm.hdc, "(None)", 4, n, w, 24, DT_VCENTER Or DT_SINGLELINE, d_TextDis, , True
    Else
     'draw menu item
     bDrawRow = Not bToolBar
     For i = 1 To .nCount
      With .d(i)
       If .mnuWidth > 0 Then
        If .nType = 6 Then 'v-separator
         hbr = CreateSolidBrush(d_Sprt1)
         r.Left = .mnuLeft + 1
         r.Top = .mnuTop
         r.Right = r.Left + 1
         r.Bottom = h - 2
         FillRect bm.hdc, r, hbr
         DeleteObject hbr
         bDrawRow = Not bToolBar
        Else
         'sidebar
         If bDrawRow Then
          GradientFillRect bm.hdc, .mnuLeft, .mnuTop, .mnuLeft + 20, h - 2, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_H '20?
          bDrawRow = False
         End If
         If .nType = 1 Then 'separator
          hbr = CreateSolidBrush(d_Sprt1)
          r.Top = .mnuTop + 1
          r.Right = .mnuLeft + .mnuWidth - 1
          r.Bottom = r.Top + 1
          If bToolBar Then
           r.Left = .mnuLeft + 1
           If (.nFlags And 256&) = 0 Then
            r.Right = r.Left + 1
            r.Bottom = .mnuTop + .mnuHeight - 1
           End If
          Else
           r.Left = .mnuLeft + 22
          End If
          FillRect bm.hdc, r, hbr
          DeleteObject hbr
         Else
          '///start
          r.Left = .mnuLeft
          r.Top = .mnuTop
          r.Right = r.Left + .mnuWidth
          r.Bottom = r.Top + .mnuHeight
          bDrawDropdown = .nType = 5 Or (.nFlags And 4&) <> 0 Or (.sSubMenu <> "" And Not bToolBar)
          If bHook Then
           hbm = InStr(1, .sKey, ";")
           bDrawPic = hbm > 0
           If bDrawPic Then
            hbm = Val(Mid(.sKey, hbm + 1))
            bDrawPic = hbm <> 0
           End If
          Else
           bDrawPic = .PicLeft >= 0 And (Not bmPic Is Nothing Or bToolBar Or (.nType And 1024&) <> 0)
          End If
          bSelected = (i = idxHl Or (Index = HlWndIndex And i = HlIndex)) And (.nFlags And 2&) = 0
          bDef = i = nDef Or (.nFlags And 64&) <> 0
          If bDef Then Set fnt = cFntB Else Set fnt = cFnt
          nOwnerDraw = &HFFFF&
          '///owner draw???
          If .nFlags And 16& Then
           nOwnerDraw = 0
           If bHook Then
            t.CtlType = ODT_MENU
            t.itemID = Val(.sKey)
            t.itemAction = ODA_DRAWENTIRE
            t.itemState = (ODS_CHECKED And (.Value <> 0)) _
            Or (ODS_DEFAULT And bDef) _
            Or (&H6& And ((.nFlags And 2&) <> 0))
            t.hwndItem = hMenu
            t.rcItem.Left = 0
            t.rcItem.Top = 0
            t.rcItem.Right = .mnuWidth
            t.rcItem.Bottom = .mnuHeight
            t.itemData = InStr(1, .sKey, "?")
            If t.itemData > 0 Then t.itemData = Val(Mid(.sKey, t.itemData + 1))
            '///fix the select bug ... but slow
            With wnds(Index)
             If .bmHook Is Nothing Then
              Set .bmHook = New cDIBSection
              .bmHook.Create .w, .h
             End If
             t.hdc = .bmHook.hdc
             BitBlt t.hdc, 0, 0, .w, .h, 0, 0, 0, &HFF0062 'whiteness
            End With
            SetBkMode t.hdc, TRANSPARENT
            hbr = SelectObject(t.hdc, fnt.Handle)
            SendMessage m_hWnd, WM_DRAWITEM, 0, t
            SelectObject t.hdc, hbr
            '///draw border
            If bSelected Then
             GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
             FrameRect bm.hdc, r, hbrBorder
            End If
            TransparentBlt bm.hdc, r.Left, r.Top, t.rcItem.Right - 4, t.rcItem.Bottom, _
            t.hdc, 4, 0, t.rcItem.Right - 4, t.rcItem.Bottom, GetPixel(t.hdc, 0, 0)
           Else
            hbr = wnds(Index).idx
            RaiseEvent DrawItem(hbr, mnus.d(hbr).sKey, i, .sKey, bm.hdc, r.Left, r.Top, r.Right, r.Bottom, fbtoBefore, nOwnerDraw)
            If nOwnerDraw And 65536 Then
             bDrawPic = True
            ElseIf nOwnerDraw And 131072 Then
             bDrawPic = .PicLeft >= 0
            End If
           End If
          End If
          If nOwnerDraw And 1& Then
           'draw border
           If bSelected Then
            GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Hl1, d_Hl2, GRADIENT_FILL_RECT_V
            FrameRect bm.hdc, r, hbrBorder
            If .nType = 5 Then 'split
             hbm = r.Left
             If bToolBar Then hbr = 9 Else hbr = 12
             r.Left = r.Right - hbr
             'GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
             FrameRect bm.hdc, r, hbrBorder
             r.Left = hbm
            End If
           ElseIf .nType = 5 And Not bToolBar Then 'split
            hbm = r.Left
            r.Left = r.Right - 12
            r.Right = r.Left + 1
            hbr = CreateSolidBrush(d_Sprt1)
            FillRect bm.hdc, r, hbr
            DeleteObject hbr
            r.Right = r.Left + 12
            r.Left = hbm
           End If
          End If
          'draw checked
          If (.nFlags And 2050&) = 2& Then n = d_TextDis Else n = d_Text
          If .Value And (nOwnerDraw And 16&) <> 0 Then
           'size???
           If Not bToolBar Then
            r.Left = r.Left + 1
            r.Top = r.Top + 1
            r.Right = r.Left + 18
            r.Bottom = r.Bottom - 1
            If (.nFlags And 2&) = 0 Then
             If bSelected Then
              GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
             Else
              GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Checked1, d_Checked2, GRADIENT_FILL_RECT_V
             End If
            End If
            'icon
            If Not bDrawPic Then pDrawCheck bm.hdc, r.Left + 9, .mnuTop + .mnuHeight \ 2, n, .nType
            'border
            FrameRect bm.hdc, r, hbrBorder
           Else
            If (.nFlags And 2&) = 0 Then
             If bSelected Then
              GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Pressed1, d_Pressed2, GRADIENT_FILL_RECT_V
             Else
              GradientFillRect bm.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Checked1, d_Checked2, GRADIENT_FILL_RECT_V
             End If
            End If
            'border
            FrameRect bm.hdc, r, hbrBorder
            If .nType = 5 Then
             r.Right = r.Left - 9
             FrameRect bm.hdc, r, hbrBorder
            End If
           End If
          End If
          'do OwnerDrawAfter
          If .nFlags And 32& Then
           nOwnerDraw = 0
           hbr = wnds(Index).idx
           RaiseEvent DrawItem(hbr, mnus.d(hbr).sKey, i, .sKey, bm.hdc, r.Left, r.Top, r.Right, r.Bottom, fbtoAfter, nOwnerDraw)
          End If
          'draw text
          If nOwnerDraw And 4& Then
           If bToolBar Then
            r.Left = .mnuLeft + 2
            If bDrawPic Then r.Left = r.Left + 18 '?
            fnt.DrawTextXP bm.hdc, .s, r.Left, .mnuTop, .mnuWidth, .mnuHeight, DT_VCENTER Or DT_SINGLELINE, n, , True
           Else
            fnt.DrawTextXP bm.hdc, .s, .mnuLeft + 24, .mnuTop, .mnuWidth, .mnuHeight, DT_VCENTER Or DT_SINGLELINE, n, , True
            If mode1 = 2 Then
             r.Left = .mnuLeft - 2
             If bDrawDropdown Then r.Left = r.Left - 8 '12?
             fnt.DrawTextXP bm.hdc, .sTab, r.Left, .mnuTop, .mnuWidth, .mnuHeight, DT_RIGHT Or DT_VCENTER Or DT_SINGLELINE, n, , True
            Else
             fnt.DrawTextXP bm.hdc, .sTab, .mnuLeft + .mnuWidth2, .mnuTop, .mnuWidth, .mnuHeight, DT_VCENTER Or DT_SINGLELINE, n, , True
            End If
           End If
          End If
          If nOwnerDraw And 8& Then
           If bDrawDropdown Then
            If bToolBar Then
             pDrawDropdown_TB bm.hdc, .mnuLeft + .mnuWidth - 5, .mnuTop + .mnuHeight \ 2, n
            Else
             pDrawDropdown bm.hdc, .mnuLeft + .mnuWidth - 4, .mnuTop + .mnuHeight \ 2, n
            End If
           End If
          End If
          'draw picture
          If bDrawPic And (nOwnerDraw And 2&) <> 0 Then
           If bHook Then
            Select Case hbm
            Case 5, 6, 8 'close
             pDrawCloseIcon bm.hdc, .mnuLeft + 10, .mnuTop + .mnuHeight \ 2, n
            Case 3, 7, 11 'minimize
             pDrawMinIcon bm.hdc, .mnuLeft + 10, .mnuTop + .mnuHeight \ 2, n
            Case 10 'maximize
             pDrawMaxIcon bm.hdc, .mnuLeft + 10, .mnuTop + .mnuHeight \ 2, n
            Case 2, 9 'restore
             pDrawRestoreIcon bm.hdc, .mnuLeft + 10, .mnuTop + .mnuHeight \ 2, n
            Case Else
             r.Right = GetSystemMetrics(SM_CXMENUCHECK) And &HFFFF&
             r.Bottom = GetSystemMetrics(SM_CYMENUCHECK) And &HFFFF&
             If r.Right > PicSize Then r.Right = PicSize
             If r.Bottom > PicSize Then r.Bottom = PicSize
             r.Left = .mnuLeft + 2 + (PicSize - r.Right) \ 2
             r.Top = .mnuTop + 2 + (PicSize - r.Bottom) \ 2
             n = CreateCompatibleDC(0)
             hbm = SelectObject(n, hbm)
             'TODO:transparent
             TransparentBlt bm.hdc, r.Left, r.Top, r.Right, r.Bottom, n, 0, 0, r.Right, r.Bottom, GetPixel(n, 0, 0)
             SelectObject n, hbm
             DeleteDC n
            End Select
           Else
            If .nFlags And 1024& Then
             r.Left = .mnuLeft + 4
             r.Top = .mnuTop + 4
             r.Right = r.Left + 12
             r.Bottom = r.Top + 12
             hbr = CreateSolidBrush(.PicLeft)
             FillRect bm.hdc, r, hbr
             DeleteObject hbr
             FrameRect bm.hdc, r, hbrBorder
            Else
             If Not bmPic Is Nothing Then
              If .nFlags And 2& Then n = bmGray.hdc Else n = bmPic.hdc
              TransparentBlt bm.hdc, .mnuLeft + 2, .mnuTop + 2, PicSize, PicSize, n, .PicLeft, 0, PicSize, PicSize, transClr
             End If
            End If
           End If
          End If
          'TODO:draw other
         End If
        End If
       End If
      End With
     Next i
    End If
   End With
   'over
   DeleteObject hbrBorder
   'alphablend?
   If .nAlpha > 0 And .nAlpha < 255 Then
    If Not .bmBack Is Nothing Then
     AlphaBlend bm.hdc, 0, 0, w, h, .bmBack.hdc, 0, 0, w, h, (255 - .nAlpha) * &H10000
    End If
   End If
   If bPaint Then bm.PaintPicture p1(Index).hdc
  End If
 End With
End If
End Sub

Public Sub Refresh()
pDirty
End Sub

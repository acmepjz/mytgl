VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "MyTGL"
   ClientHeight    =   7200
   ClientLeft      =   165
   ClientTop       =   555
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
   ScaleHeight     =   480
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   640
   StartUpPosition =   3  '´°¿ÚÈ±Ê¡
   Begin MyTGL.FakeWindowXP0 frmBird 
      Height          =   2370
      Left            =   4920
      TabIndex        =   0
      Top             =   720
      Visible         =   0   'False
      Width           =   4050
      _ExtentX        =   7144
      _ExtentY        =   4180
      Caption         =   "Bird's Eye Tool"
      CloseButton     =   -1  'True
      MinButton       =   -1  'True
      AutoDrag        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      IsVisible       =   0   'False
   End
   Begin VB.PictureBox p0 
      BorderStyle     =   0  'None
      Height          =   1695
      Index           =   0
      Left            =   120
      ScaleHeight     =   113
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   297
      TabIndex        =   8
      Top             =   840
      Width           =   4455
      Begin VB.PictureBox picView 
         Height          =   495
         Left            =   0
         ScaleHeight     =   29
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   101
         TabIndex        =   12
         Top             =   0
         Width           =   1575
         Begin MyTGL.ctlNCPaint nc1 
            Left            =   120
            Top             =   120
            _ExtentX        =   2355
            _ExtentY        =   450
            PaintMode       =   1
            BorderWidth     =   1
         End
      End
      Begin VB.PictureBox picProp 
         Height          =   1335
         Left            =   1680
         ScaleHeight     =   85
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   133
         TabIndex        =   9
         Top             =   0
         Width           =   2055
         Begin VB.TextBox txtProp 
            BorderStyle     =   0  'None
            Height          =   225
            Left            =   120
            TabIndex        =   10
            Top             =   840
            Visible         =   0   'False
            Width           =   1215
         End
         Begin MyTGL.LeftRight objDrag 
            Height          =   375
            Left            =   1440
            Top             =   840
            Width           =   135
            _ExtentX        =   238
            _ExtentY        =   661
         End
         Begin MyTGL.FakeComboBox cmbProp 
            Height          =   255
            Left            =   120
            TabIndex        =   11
            Top             =   480
            Visible         =   0   'False
            Width           =   1215
            _ExtentX        =   2143
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
            PanelWidth      =   13
            DropdownHeight  =   16
         End
         Begin MyTGL.ctlWndScroll sb2 
            Left            =   120
            Top             =   120
            _ExtentX        =   2990
            _ExtentY        =   450
            Orientation     =   1
            NCPaintMode     =   1
         End
      End
      Begin MyTGL.ctlSplitter sp1 
         Index           =   1
         Left            =   0
         Top             =   720
         _ExtentX        =   2778
         _ExtentY        =   450
         Orientation     =   2
         FullDrag        =   0   'False
         KeepProportion  =   -1  'True
      End
   End
   Begin VB.PictureBox p0 
      BorderStyle     =   0  'None
      Height          =   1815
      Index           =   1
      Left            =   120
      ScaleHeight     =   121
      ScaleMode       =   3  'Pixel
      ScaleWidth      =   257
      TabIndex        =   4
      Top             =   3720
      Width           =   3855
      Begin VB.PictureBox picObj 
         Height          =   1095
         Left            =   2640
         ScaleHeight     =   69
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   133
         TabIndex        =   6
         Top             =   0
         Width           =   2055
         Begin MyTGL.ctlWndScroll sb1 
            Left            =   120
            Top             =   120
            _ExtentX        =   2990
            _ExtentY        =   450
            NCPaintMode     =   1
         End
      End
      Begin VB.PictureBox picList 
         Height          =   1695
         Left            =   480
         Picture         =   "Form1.frx":0000
         ScaleHeight     =   109
         ScaleMode       =   3  'Pixel
         ScaleWidth      =   133
         TabIndex        =   5
         Top             =   0
         Width           =   2055
         Begin MyTGL.ctlWndScroll sbList 
            Left            =   480
            Top             =   120
            _ExtentX        =   2990
            _ExtentY        =   450
            Orientation     =   1
            NCPaintMode     =   1
         End
      End
      Begin MyTGL.FakeToolBar tb2 
         Height          =   1140
         Left            =   0
         TabIndex        =   7
         Top             =   0
         Width           =   300
         _ExtentX        =   529
         _ExtentY        =   2011
         Picture         =   "Form1.frx":00FF
         TheString       =   $"Form1.frx":0782
         BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
            Name            =   "Tahoma"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Orientation     =   2
      End
      Begin MyTGL.ctlSplitter sp1 
         Index           =   2
         Left            =   2640
         Top             =   1200
         _ExtentX        =   2778
         _ExtentY        =   450
         Orientation     =   2
         FullDrag        =   0   'False
         KeepProportion  =   -1  'True
      End
   End
   Begin MyTGL.FakeWindowXP0 frmAddOp 
      Height          =   735
      Left            =   4920
      TabIndex        =   1
      Top             =   3600
      Visible         =   0   'False
      Width           =   4095
      _ExtentX        =   7223
      _ExtentY        =   1296
      Caption         =   ""
      CloseButton     =   -1  'True
      MinButton       =   -1  'True
      AutoDrag        =   -1  'True
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Float           =   -1  'True
      FakeWindowTabMode=   99
      IsVisible       =   0   'False
   End
   Begin MyTGL.FakeMenu fm0 
      Left            =   8160
      Top             =   120
      _ExtentX        =   2355
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
   End
   Begin MyTGL.FakeMenu fm1 
      Left            =   8160
      Top             =   3120
      _ExtentX        =   2355
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
   End
   Begin MyTGL.FakeToolBar tbMenu 
      Height          =   300
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3615
      _ExtentX        =   6376
      _ExtentY        =   529
      TheString       =   ";&File;;;;;_file,;&Edit;;;;;_edit,;&View;;;;;_view,;&Tools;;;;;_tools,;&Help;;;;;_help"
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      MainMenu        =   -1  'True
   End
   Begin MyTGL.FakeToolBar tb1 
      Height          =   300
      Left            =   120
      TabIndex        =   3
      Top             =   480
      Width           =   4335
      _ExtentX        =   7646
      _ExtentY        =   529
      Picture         =   "Form1.frx":0819
      TheString       =   $"Form1.frx":1252
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
   Begin MyTGL.ctlSplitter sp1 
      Index           =   0
      Left            =   120
      Top             =   3360
      _ExtentX        =   2778
      _ExtentY        =   450
      FullDrag        =   0   'False
      KeepProportion  =   -1  'True
   End
   Begin MyTGL.SimpleStatusBar stb1 
      Height          =   240
      Left            =   120
      Top             =   6840
      Width           =   4320
      _ExtentX        =   7620
      _ExtentY        =   423
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      SimpleText      =   ""
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

#Const Office2003 = 1

Const int_Caption_Width As Long = 80&
Const int_Prop_Height As Long = 16&

Const int_Caption_Width_1 As Long = int_Caption_Width \ 4
Const int_Caption_Width_2 As Long = int_Caption_Width \ 2
Const int_Caption_Width_3 As Long = int_Caption_Width - int_Caption_Width_1

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Const CB_SETITEMHEIGHT As Long = &H153
Private Const WM_SETFONT As Long = &H30

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function FrameRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Declare Function InflateRect Lib "user32.dll" (ByRef lpRect As RECT, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function TransparentBlt Lib "msimg32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal crTransparent As Long) As Long

Private Const WS_BORDER As Long = &H800000
Private Const LBS_NOINTEGRALHEIGHT As Long = &H100&
Private Const CBS_NOINTEGRALHEIGHT As Long = &H400&
Private Const GWL_STYLE As Long = -16
Private Declare Function GetWindowLong Lib "user32.dll" Alias "GetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetWindowLong Lib "user32.dll" Alias "SetWindowLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long

Private Const SWP_NOMOVE As Long = &H2
Private Const SWP_NOSIZE As Long = &H1
Private Const SWP_NOZORDER As Long = &H4
Private Const SWP_NOACTIVATE As Long = &H10
Private Declare Function SetWindowPos Lib "user32.dll" (ByVal hwnd As Long, ByVal hWndInsertAfter As Long, ByVal x As Long, ByVal y As Long, ByVal cx As Long, ByVal cy As Long, ByVal wFlags As Long) As Long

'get memory usage
Private Type PROCESS_MEMORY_COUNTERS
    cb As Long
    PageFaultCount As Long
    PeakWorkingSetSize As Long
    WorkingSetSize As Long
    QuotaPeakPagedPoolUsage As Long
    QuotaPagedPoolUsage As Long
    QuotaPeakNonPagedPoolUsage As Long
    QuotaNonPagedPoolUsage As Long
    PagefileUsage As Long
    PeakPagefileUsage As Long
End Type
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long
Private Declare Function GetProcessMemoryInfo Lib "psapi.dll" (ByVal Process As Long, ByRef ppsmemCounters As PROCESS_MEMORY_COUNTERS, ByVal cb As Long) As Long

'get CPU usage
Private cCPU1 As clsCPU

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Private Const WHITE_BRUSH As Long = 0

Private Declare Function CreatePolygonRgn Lib "gdi32.dll" (ByRef lpPoint As POINTAPI, ByVal nCount As Long, ByVal nPolyFillMode As Long) As Long
Private Declare Function Polyline Lib "gdi32.dll" (ByVal hdc As Long, ByRef lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function SelectClipRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long) As Long
Private Const ALTERNATE As Long = 1
Private Declare Function FrameRgn Lib "gdi32.dll" (ByVal hdc As Long, ByVal hRgn As Long, ByVal hBrush As Long, ByVal nWidth As Long, ByVal nHeight As Long) As Long

Private Declare Sub mouse_event Lib "user32.dll" (ByVal dwFlags As Long, ByVal dX As Long, ByVal dy As Long, ByVal cButtons As Long, ByVal dwExtraInfo As Long)
Private Const MOUSEEVENTF_LEFTDOWN As Long = &H2

Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const KEYEVENTF_KEYUP As Long = &H2

Private Declare Function GetAsyncKeyState Lib "user32.dll" (ByVal vKey As Long) As Integer

Private Declare Function PathRelativePathTo Lib "shlwapi.dll" Alias "PathRelativePathToA" (ByVal pszPath As String, ByVal pszFrom As String, ByVal dwAttrFrom As Long, ByVal pszTo As String, ByVal dwAttrTo As Long) As Long
Private Declare Function PathCombine Lib "shlwapi.dll" Alias "PathCombineA" (ByVal szDest As String, ByVal lpszDir As String, ByVal lpszFile As String) As Long
'Private Declare Function PathCompactPathEx Lib "shlwapi.dll" Alias "PathCompactPathExA" (ByVal pszOut As String, ByVal pszSrc As String, ByVal cchMax As Long, ByVal dwFlags As Long) As Long
Private Declare Function PathCompactPathExW Lib "shlwapi.dll" (ByRef pszOut As Any, ByRef pszSrc As Any, ByVal cchMax As Long, ByVal dwFlags As Long) As Long

'calc time
Private Declare Function GetTickCount Lib "kernel32.dll" () As Long
Private nOldTime As Long

Implements IOperatorCalc

'some bitmap
Private bmProp As New cDIBSection
Private bmObj As New cDIBSection
Private bmView As New cDIBSection
Private bmList As New cDIBSection

Private bm0 As New cDIBSection

'font object
Private cFnt As New CLogFont

'selected operator
Private nSelIndex As Long
Private nSelType As Long
Private nSelPropIndex As Long
Private nSelPropSubIndex As Long '1-based

'showed operator
Private nShowIndex As Long

'cursor
Private nXCur As Long, nYCur As Long
Private nXEnd As Long, nYEnd As Long
Private nCurFlags As Long
'0  nothing
'1  selection
'2  move
'3  resize left
'4  resize right
'----
'5  resize top
'6  resize bottom
'7  resize top-left
'8  resize top-right
'9  resize bottom-left
'10 resize bottom-right
'----

'show menu?
Private bShowMenu As Boolean '????

'view box
Private bViewTile As Boolean
Private nViewX As Long, nViewY As Long
Private nViewXCur As Long, nViewYCur As Long

'project
Private TheFileName As String 'current file
Private ThePrj As typeProject
Private cObj As New clsOperators

'clipboard
Private TheClipboard() As typeOperator_DesignTime
Private TheClipboardItemCount As Long
Private nClipboardLeft As Long, nClipboardTop As Long, nClipboardRight As Long, nClipboardBottom As Long

'list box
Private ThePageState() As Long '??
'0=expand
'1=collapsed
'2=hide expand
'3=hide collapsed
Private VisiblePageCount As Long '??
Private TheIndex As Long '0 to count-1

Private Sub cmbProp_Click()
pChangeValue CStr(cmbProp.ListIndex)
End Sub

Private Sub cmbProp_KeyDown(ByVal KeyCode As Long, ByVal Shift As Long)
Select Case KeyCode
Case vbKeyEscape
 cmbProp.ListIndex = Val(cmbProp.Tag)
 cmbProp.Visible = False
Case vbKeyReturn
 cmbProp.Visible = False
End Select
End Sub

Private Sub cmbProp_MyLostFocus()
cmbProp.Visible = False
End Sub

Private Sub fm1_Click(ByVal idxMenu As Long, ByVal Key As String, ByVal idxButton As Long, ByVal ButtonKey As String, Value As Long)
Dim i As Long
Dim s As String, s2 As String, s3 As String
Select Case Left(ButtonKey, 4)
Case "idx:"
 i = Val(Mid(ButtonKey, 5))
 If i <= 0 Then Exit Sub
 Select Case i
 Case IDM_ADDCOMMENT
  mnuAddC_Click
 Case IDM_SHOWOP
  mnuShowOp_Click
 Case IDM_BRINGTOFRONT
  mnuCom_Click 1
 Case IDM_SENDTOBACK
  mnuCom_Click 2
 Case Else
  'TODO:
  mnuAdd_Click (i)
 End Select
Case "new"
 mnuNew_Click
Case "open"
 mnuOpen_Click
Case "save"
 mnuSave_Click
Case "s_as"
 mnuSA_Click
Case "mru_"
 i = Val(Mid(ButtonKey, 5))
 If i > 0 Then
  s = cSet.GetSettings("MRU" + CStr(i))
  If s = "" Then Exit Sub
  s2 = CStr(App.Path)
  If Right(s2, 1) <> "\" Then s2 = s2 + "\"
  s3 = String(1024, vbNullChar)
  PathCombine s3, s2, s
  s3 = Replace(s3, vbNullChar, "")
  pOpen s3
 End If
Case "clrR"
 For i = 1 To 8
  cSet.Remove "MRU" + CStr(i)
 Next i
 pMRU
Case "edit"
 i = Val(Mid(ButtonKey, 5))
 mnuEdit1_Click (i)
Case "exit"
 Unload Me
Case "Vres"
 nViewX = 0
 nViewY = 0
 pViewRedraw
Case "tile"
 bViewTile = Value
 pViewRedraw
Case "sbmp"
 mnuSaveBmp_Click
Case "beye"
 frmBird.IsVisible = Value
 If frmBird.IsVisible Then frmBird.Refresh
Case "opts"
 frmOpt.Show 1
Case "addp"
 s = "Page " + CStr(ThePrj.nPageCount + 1)
 cObj.AddPage ThePrj, s
 'AddItem
 VisiblePageCount = VisiblePageCount + 1
 ReDim Preserve ThePageState(1 To ThePrj.nPageCount)
 pListRefresh
Case "renp"
 If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
 s = InputBox("Input new name:", "Rename", ThePrj.Pages(TheIndex).Name)
 If s <> "" Then
  ThePrj.Pages(TheIndex).Name = s
  pListRedraw
 End If
Case "p2_u"
 mnuPMoveUp_Click
Case "p2_d"
 mnuPMoveDown_Click
Case "p2_l"
 If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
 'TODO:fix some bug
 With ThePrj.Pages(TheIndex)
  If .nIndent > 0 Then .nIndent = .nIndent - 1
  pListRedraw
 End With
Case "p2_r"
 If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
 'TODO:fix some bug
 With ThePrj.Pages(TheIndex)
  If .nIndent < 255 Then .nIndent = .nIndent + 1
  pListRedraw
 End With
Case "delp"
 mnuDelP_Click
Case "p3_r"
 mnuResetP_Click
Case "abou"
 frmAbout.Show 1
End Select
End Sub

Private Sub fm1_MakeMenuFloat(ByVal idxMenu As Long, ByVal Key As String, ByVal x As Long, ByVal y As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long)
idxMenu = frmAddOp.FindTab(Key)
If idxMenu > 0 Then
 frmAddOp.SelectedTab = idxMenu
 frmAddOp.MoveEx Left, Top, , , True
 frmAddOp.IsVisible = True
 frmAddOp.DragStart
End If
End Sub

Private Sub Form_Initialize()
'init data
Set cCPU1 = New clsCPU
cObj.AddPage ThePrj, "Page 1"
ReDim ThePageState(1 To 1)
VisiblePageCount = 1
TheIndex = -1
End Sub

Private Sub Form_Load()
Dim i As Long, j As Long, k As Long, s As String
Dim b As Boolean
'init bitmap
bm0.CreateFromPicture picList.Picture
Set picList.Picture = Nothing
'init font
cFnt.HighQuality = True
Set cFnt.LogFont = Me.Font
'///new!!! init menu
FakeCommandBarAddButton mnu.d(1), , , , fbttSeparator
FakeCommandBarAddButton mnu.d(1), "idx:" + CStr(IDM_SHOWOP), "&Show Operator" + vbTab + "S"
FakeCommandBarAddButton mnu.d(1), , , , fbttSeparator
'edit menu
k = FakeCommandBarAddButton(mnu.d(1), "edit0", "C&ut", , , , , 48, "X")
FakeCommandBarAddButton mnu.d(1), "edit1", "&Copy", , , , , 64, "C"
FakeCommandBarAddButton mnu.d(1), "edit2", "&Paste", , , , , 80, "V"
FakeCommandBarAddButton mnu.d(1), "edit3", "&Delete", , , , , 96, "Del"
FakeCommandBarAddButton mnu.d(1), "edit4", "Select &All", , , , , , "Ctrl+A"
j = FakeCommandBarAddCommandBar(mnu, "_edit")
For i = k To k + 4
 FakeCommandBarAddButtonIndirect mnu.d(j), mnu.d(1).d(i)
Next i
'continue
FakeCommandBarAddButton mnu.d(1), , , , fbttSeparator
FakeCommandBarAddButton mnu.d(1), "idx:" + CStr(IDM_BRINGTOFRONT), "&Bring to front"
FakeCommandBarAddButton mnu.d(1), "idx:" + CStr(IDM_SENDTOBACK), "&Send to back"
FakeCommandBarApplyMenu fm1, mnu
'clear up
Erase mnu.d
mnu.nCount = 0
With fm1
 s = ";;;hidden;;;;mru_1,;;;hidden;;;;mru_2,;;;hidden;;;;mru_3,;;;hidden;;;;mru_4" + _
 ",;;;hidden;;;;mru_5,;;;hidden;;;;mru_6,;;;hidden;;;;mru_7,;;;hidden;;;;mru_8"
 'file menu
 .AddMenuFromString "_file", "1;&New;;;Ctrl+N;;;new,2;&Open;;;Ctrl+O;;;open,,3;&Save;;;Ctrl+S;;;save,;Save &As;;;;;;s_as,;;;separator:hidden;;;;mru_0," + _
 s + ",,;E&xit;;;Alt+F4;;;exit"
 'view menu
 .AddMenuFromString "_view", ";&Reset;;;;;;Vres,;&Tile;;check;;;;tile,,8;&Save as Bitmap;;;;;;sbmp"
 'tools menu
 .AddMenuFromString "_tools", ";&Bird's Eye;;check;Y;;;beye,,;&Options;;;Ctrl+K;;;opts"
 'MRU menu
 .AddMenuFromString "_mru", s + ",;;;separator:hidden;;;;mru_0,;&Clear Recent;;hidden;;;;clrR"
 'popup 2
 .AddMenuFromString "[P2]", ";&Add Page;;;A;;;addp,;&Rename;;;R;;;renp,;&Delete;;;Del;;;delp,,;Move &Up;;;Shift+Up;;;p2_u," + _
 ";Move D&own;;;Shift+Down;;;p2_d,,;Move &Left;;;Shift+Left;;;p2_l,;Move R&ight;;;Shift+Right;;;p2_r"
 'popup 3
 .AddMenuFromString "[P3]", ";&Reset;;;;;;p3_r"
 'help menu
 .AddMenuFromString "_help", ";&About;;;;;;about"
End With
'///
'init MRU
pMRU
'///complete
tb1.SetMenu fm1, True
tbMenu.SetMenu fm1
'///set hot key
With tbMenu
 .AddShortcutKey "_file", "new", vbKeyN, vbCtrlMask
 .AddShortcutKey "_file", "open", vbKeyO, vbCtrlMask
 .AddShortcutKey "_file", "save", vbKeyS, vbCtrlMask
 .AddShortcutKey "_edit", "edit4", vbKeyA, vbCtrlMask
 .AddShortcutKey "_tools", "opts", vbKeyK, vbCtrlMask
End With
'///add tab
For i = 1 To m_nAddOpTabCount
 frmAddOp.AddTab m_sAddOpCaption(i), m_sAddOpKey(i)
Next i
Erase m_sAddOpCaption, m_sAddOpKey
m_nAddOpTabCount = 0
frmAddOp.SetMenu fm1
'load settings
bViewTile = CBool(cSet.GetSettings("TileTexture", "False"))
With fm1
 i = .FindMenu("_view")
 j = .FindButton(i, "tile")
 .ButtonValue(i, j) = bViewTile And 1&
End With
'init splitter
With sp1(0)
 .Proportion = 45
 .Bind p0(0), p0(1)
 .SetSubObject cSPLTLeftOrTopPanel, Array(tbMenu, tb1), 1
 .SetSubObject cSPLTRightOrBottomPanel, stb1, 2
End With
With sp1(1)
 .Proportion = 60
 .Bind picView, picProp
End With
With sp1(2)
 .Proportion = 20
 .Bind picList, picObj
 .SetSubObject cSPLTLeftOrTopPanel, tb2, 3
End With
'init border color
nc1.Color1 = d_CtrlBorder
sb1.NCPaintColor1 = d_CtrlBorder
sb2.NCPaintColor1 = d_CtrlBorder
sbList.NCPaintColor1 = d_CtrlBorder
'init status
With stb1
 .AddPanel "Ready"
 .AddPanel "Memory:0.00MB", , , , 96, sbCustomInfo
 .AddPanel "CPU:0.0%", , , , 64, sbCustomInfo
 .AddPanel "(" + CStr(nXCur) + "," + CStr(nYCur) + ")", , , , 96
 .AddPanel "Time", , , , 64, sbTime
End With
'init selection
nSelIndex = 0
nSelType = 0
nSelPropIndex = 0
nCurFlags = 0
nViewX = 0
nViewY = 0
nViewXCur = &H80000000
bShowMenu = False
'test
SendMessage txtProp.hwnd, WM_SETFONT, cFnt.Handle, ByVal 0
'resize
picList_Resize
picObj_Resize
picProp_Resize
picView_Resize
'///new!!! hook menu
FakeMenuPopupHook fm0
End Sub

Private Sub Form_Terminate()
Set cCPU1 = Nothing
End Sub

Private Sub Form_Unload(Cancel As Integer)
FakeMenuPopupUnhook
'save settings
cSet.SetSettings "TileTexture", CStr(bViewTile)
cSet.SaveFile
End Sub

Private Sub frmAddOp_CloseButtonClick()
frmAddOp.IsVisible = False
End Sub

Private Sub frmBird_CloseButtonClick()
Dim i As Long, j As Long
frmBird.IsVisible = False
With fm1
 i = .FindMenu("_tools")
 j = .FindButton(i, "beye")
 .ButtonValue(i, j) = 0
End With
End Sub

Private Sub frmBird_MouseDown(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long, ByVal ClientWidth As Long, ByVal ClientHeight As Long)
If Button = 1 Then frmBird_MouseMove Button, Shift, x, y, ClientWidth, ClientHeight
End Sub

Private Sub frmBird_MouseMove(ByVal Button As Long, ByVal Shift As Long, ByVal x As Long, ByVal y As Long, ByVal ClientWidth As Long, ByVal ClientHeight As Long)
If Button = 1 Then
 x = x - 4 - bmObj.Width \ 32&
 y = y - 4 - bmObj.Height \ 32&
 If x < 0 Then x = 0
 If y < 0 Then y = 0
 If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
 With sb1
  .Value(efsHorizontal) = x * 16&
  .Value(efsVertical) = y * 16&
 End With
End If
End Sub

Private Sub frmBird_Paint(ByVal hdc As Long, ByVal ClientLeft As Long, ByVal ClientTop As Long, ByVal ClientWidth As Long, ByVal ClientHeight As Long)
Dim r As RECT, hbr As Long, hbr2 As Long
Dim i As Long, j As Long
If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
'draw comments
With ThePrj.Pages(TheIndex)
 For i = 1 To .nCommentCount
  With .Comments(i)
   hbr = CreateSolidBrush(.Color)
   r.Left = .Left + ClientLeft + 4
   r.Top = .Top + ClientTop + 4
   r.Right = r.Left + .Width
   r.Bottom = r.Top + .Height
   FillRect hdc, r, hbr
   DeleteObject hbr
  End With
 Next i
End With
'draw operator
hbr = CreateSolidBrush(&H800000)
hbr2 = CreateSolidBrush(d_Pressed1)
For j = 0 To int_Page_Height - 1
 r.Top = j + ClientTop + 4
 r.Bottom = r.Top + 1
 With ThePrj.Pages(TheIndex).Rows(j)
  For i = 1 To .nOpCount
   With ThePrj.Operators(.idxOp(i))
    r.Left = .Left + ClientLeft + 4
    r.Right = r.Left + .Width
    If .Flags And int_OpFlags_Selected Then
     FillRect hdc, r, hbr2
    Else
     FillRect hdc, r, hbr
    End If
   End With
  Next i
 End With
Next j
DeleteObject hbr
DeleteObject hbr2
'draw box
hbr = CreateSolidBrush(vbRed)
r.Left = sb1.Value(efsHorizontal) \ 16& + ClientLeft + 4
r.Top = sb1.Value(efsVertical) \ 16& + ClientTop + 4
r.Right = r.Left + bmObj.Width \ 16& + 1
r.Bottom = r.Top + bmObj.Height \ 16& + 1
FrameRect hdc, r, hbr
DeleteObject hbr
End Sub

Private Sub IOperatorCalc_OnProgress(ByVal nOpNow As Long, ByVal nOpCount As Long, ByVal nCalcNow As Long, ByVal nCalcMax As Long, bAbort As Boolean)
On Error Resume Next
If nOpCount > 0 And GetTickCount - nOldTime > 500 Then
 If nCalcMax > 0 Then nCalcNow = (nCalcNow * 256&) \ nCalcMax Else nCalcNow = 0
 With stb1
  .PanelStyle(1) = sbProgressBar
  .ProgressBarMax(1) = nOpCount * 256&
  .ProgressBarValue(1) = (nOpNow - 1) * 256& + nCalcNow
  .PanelCaption(1) = "Processing...Press 'ESC' to cancel..." + CStr(nOpNow) + "/" + CStr(nOpCount)
 End With
 'cancel?
 If GetAsyncKeyState(vbKeyEscape) And &H8000 Then
  nShowIndex = 0
  bAbort = True
 End If
End If
End Sub

Private Sub lstPage_Click()
nSelIndex = 0
nSelType = 0
nSelPropIndex = 0
nCurFlags = 0
cObj.ClearFlags ThePrj
pPropRefresh
pObjRedraw
End Sub

Private Sub mnuAdd_Click(Index As Integer)
Dim idx As Long
If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
pClearLockCount 'fix the unknown bug
idx = cObj.AddOperator(ThePrj, Index, TheIndex, nXCur, nYCur)
If idx Then
 'validate operator
 cObj.ValidateOps ThePrj, idx
 If Index = int_OpType_Store Then
  cObj.ValidateAllLoadOps ThePrj, ""
 End If
 'calc bitmap
 pViewRedraw
 pObjRedraw
End If
End Sub

Private Sub mnuAddC_Click()
Dim n As Long
If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
With ThePrj.Pages(TheIndex)
 n = .nCommentCount + 1
 .nCommentCount = n
 ReDim Preserve .Comments(1 To n)
 With .Comments(n)
  .Left = nXCur
  .Top = nYCur
  .Width = 8
  .Height = 4
  .Color = d_Bar1
  .Name = "Comment" + CStr(n)
 End With
End With
pObjRedraw
End Sub

Private Sub mnuCom_Click(ByVal Index As Integer)
Dim idx As Long, i As Long
Dim d As typeComment
If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
With ThePrj.Pages(TheIndex)
 idx = -nSelIndex
 If idx > 0 And idx <= .nCommentCount Then
  Select Case Index
  Case 1 'bring to front
   If idx < .nCommentCount Then
    d = .Comments(idx)
    For i = idx + 1 To .nCommentCount
     .Comments(i - 1) = .Comments(i)
    Next i
    .Comments(.nCommentCount) = d
    nSelIndex = -.nCommentCount
    pObjRedraw
   End If
  Case 2 'send to back
   If idx > 1 Then
    d = .Comments(idx)
    For i = idx - 1 To 1 Step -1
     .Comments(i + 1) = .Comments(i)
    Next i
    .Comments(1) = d
    nSelIndex = -1
    pObjRedraw
   End If
  End Select
 End If
End With
End Sub

Private Sub mnuDelP_Click()
Dim sto() As typeStoreOp_DesignTime, m As Long
Dim i As Long, j As Long
If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
If MsgBox("Do you really want to delete?", vbExclamation Or vbYesNo Or vbDefaultButton2, "Warning") = vbYes Then
 'get start and end
 m = ThePrj.Pages(TheIndex).nIndent
 For i = TheIndex + 1 To ThePrj.nPageCount
  If ThePrj.Pages(i).nIndent <= m Then Exit For
 Next i
 j = i - 1
 'get store object
 m = cObj.GetStoreObjects(ThePrj, sto)
 'delete
 pClearLockCount 'fix the unknown bug
 cObj.DeletePageEx ThePrj, TheIndex, j
 'validate operator
 For i = 1 To m
  With sto(i)
   If .Index > 0 And .Index <= ThePrj.nOpCount Then
    If ThePrj.Operators(.Index).Flags < 0 Then 'this is deleted!
     cObj.ValidateAllLoadOps ThePrj, .Name
    End If
   End If
  End With
 Next i
 'treeview.RemoveItem
 If ThePrj.nPageCount > 0 Then
  m = 0
  For i = TheIndex To j
   If ThePageState(i) < 2 Then m = m + 1
  Next i
  VisiblePageCount = VisiblePageCount - m
  If TheIndex <= ThePrj.nPageCount Then _
  CopyMemory ThePageState(j + 1), ThePageState(TheIndex), (ThePrj.nPageCount - TheIndex + 1) * 4& _
  Else TheIndex = ThePrj.nPageCount
  ReDim Preserve ThePageState(1 To ThePrj.nPageCount)
 Else
  TheIndex = -1
  VisiblePageCount = 0
  Erase ThePageState
 End If
 'redraw
 nSelIndex = 0
 pListRefresh
 pPropRedraw
 pViewRedraw
 pObjRedraw
End If
End Sub

Private Sub mnuEdit1_Click(Index As Integer)
Dim i As Long, j As Long
Dim bStore As Boolean, s As String
Dim tmp() As Long, m As Long
If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
Select Case Index
Case 0, 1, 3 'cut,copy,delete
 If Index = 3 And nSelIndex < 0 Then
  'delete comment
  pDeleteComment
  nSelIndex = 0
 End If
 TheClipboardItemCount = -TheClipboardItemCount ':-3
 For i = 1 To ThePrj.nOpCount
  With ThePrj.Operators(i)
   If .Flags >= 0 And (.Flags And int_OpFlags_Selected) Then
    If Index = 0 Or Index = 1 Then 'cut,copy
     If TheClipboardItemCount <= 0 Then TheClipboardItemCount = 0
     TheClipboardItemCount = TheClipboardItemCount + 1
     ReDim Preserve TheClipboard(1 To TheClipboardItemCount)
     TheClipboard(TheClipboardItemCount) = ThePrj.Operators(i)
    End If
    If Index = 0 Or Index = 3 Then 'cut,delete
     'get refresh area
     bStore = .nType = int_OpType_Store
     If bStore Then s = .Name
     m = cObj.PageHitTestEx(ThePrj, TheIndex, .Left, .Top + 1, .Left + .Width - 1, .Top + 1, tmp)
     'delete it
     cObj.DeleteOperatorByIndex ThePrj, i
     'validate operator
     For j = 1 To m
      cObj.ValidateOps ThePrj, tmp(j)
     Next j
     If bStore Then cObj.ValidateAllLoadOps ThePrj, s
     'clear select
     If i = nSelIndex Then
      nSelIndex = 0
      pPropRefresh
     End If
     'clear show
     If i = nShowIndex Then
      nShowIndex = 0
      pViewRedraw
     End If
    End If
   End If
  End With
 Next i
 If TheClipboardItemCount > 0 Then
  'check size
  nClipboardLeft = &H7FFFFFFF
  nClipboardTop = &H7FFFFFFF
  nClipboardRight = &H80000000
  nClipboardBottom = &H80000000
  For i = 1 To TheClipboardItemCount
   With TheClipboard(i)
    If nClipboardLeft > .Left Then nClipboardLeft = .Left
    If nClipboardTop > .Top Then nClipboardTop = .Top
    If nClipboardRight < .Left + .Width - 1 Then nClipboardRight = .Left + .Width - 1
    If nClipboardBottom < .Top Then nClipboardBottom = .Top
   End With
  Next i
 Else
  TheClipboardItemCount = -TheClipboardItemCount ':-3
 End If
 'calc bitmap
 pViewRedraw
 pObjRedraw
Case 2 'paste
 If TheClipboardItemCount > 0 Then
  'check size
  If nXCur >= 0 And nXCur < int_Page_Width + nClipboardLeft - nClipboardRight And _
  nYCur >= 0 And nYCur < int_Page_Height + nClipboardTop - nClipboardBottom Then
   'hit test
   bStore = True 'bValid
   For i = 1 To TheClipboardItemCount
    j = TheClipboard(i).Left - nClipboardLeft + nXCur
    m = TheClipboard(i).Top - nClipboardTop + nYCur
    If cObj.PageHitTestEx(ThePrj, TheIndex, j, m, j + TheClipboard(i).Width - 1, m, tmp) > 0 Then
     bStore = False
     Exit For
    End If
   Next i
   'paste!
   If bStore Then
    'add object
    For i = 1 To TheClipboardItemCount
     pClearLockCount 'fix the unknown bug
     With TheClipboard(i)
      j = cObj.AddOperator(ThePrj, .nType, TheIndex, .Left - nClipboardLeft + nXCur, .Top - nClipboardTop + nYCur, .Width)
      Debug.Assert j > 0
     End With
     'insert
     With ThePrj.Operators(j)
      .Name = TheClipboard(i).Name
      For m = 0 To tDef(.nType).StringCount - 1
       .sProps(m) = TheClipboard(i).sProps(m)
      Next m
      m = tDef(.nType).PropSize
      If m > 0 Then CopyMemory .bProps(0), TheClipboard(i).bProps(0), m 'fix the bug!!!!
     End With
     'validate
     cObj.ValidateOps ThePrj, j
    Next i
    'calc bitmap
    pViewRedraw
    pObjRedraw
   End If
  End If
  If Not bStore Then MsgBox "Not enough rooms.", vbExclamation
 End If
Case 4 'select all
 For i = 0 To int_Page_Height - 1
  With ThePrj.Pages(TheIndex).Rows(i)
   For j = 1 To .nOpCount
    With ThePrj.Operators(.idxOp(j))
     Debug.Assert .Flags >= 0
     .Flags = .Flags Or int_OpFlags_Selected
    End With
   Next j
  End With
 Next i
 pObjRedraw
End Select
End Sub

Private Sub pTitle()
If TheFileName = "" Then
 Me.Caption = "MyTGL"
Else
 Me.Caption = "MyTGL - " + TheFileName
End If
End Sub

Private Sub mnuNew_Click()
Dim i As Long
TheFileName = ""
pTitle
pClearLockCount 'fix the unknown bug
cObj.Clear ThePrj
cObj.AddPage ThePrj, "Page 1"
nSelIndex = 0
nSelType = 0
nShowIndex = 0
ReDim ThePageState(1 To 1)
VisiblePageCount = 1
TheIndex = -1
pListRefresh
pPropRedraw
pViewRedraw
pObjRedraw
End Sub

Private Sub mnuOpen_Click()
Dim s As String
If cd.VBGetOpenFileName(s, , , , , True, "MyTexture|*.myt", , CStr(App.Path), , , Me.hwnd) Then pOpen s
End Sub

Private Sub mnuPMoveDown_Click()
Dim i As Long, j As Long, m As Long
If TheIndex <= 0 Or TheIndex >= ThePrj.nPageCount Then Exit Sub
'collapsed?
If ThePageState(TheIndex) = 1 Then
 m = ThePrj.Pages(TheIndex).nIndent
 For j = TheIndex + 1 To ThePrj.nPageCount
  If ThePrj.Pages(j).nIndent <= m Then Exit For
 Next j
 j = j - 1
Else
 j = TheIndex
End If
'find next
i = j
Do
 i = i + 1
 If i > ThePrj.nPageCount Then Exit Sub
Loop While ThePageState(i) >= 2
'TODO:fix some bug
'move it
cObj.MovePageEx ThePrj, TheIndex, j, i, VarPtr(ThePageState(1))
TheIndex = i
pListRedraw
End Sub

Private Sub mnuPMoveUp_Click()
Dim i As Long, j As Long, m As Long
If TheIndex <= 1 Or TheIndex > ThePrj.nPageCount Then Exit Sub
'find previous
i = TheIndex
Do
 i = i - 1
Loop While i > 1 And ThePageState(i) >= 2
'collapsed?
If ThePageState(TheIndex) = 1 Then
 m = ThePrj.Pages(TheIndex).nIndent
 For j = TheIndex + 1 To ThePrj.nPageCount
  If ThePrj.Pages(j).nIndent <= m Then Exit For
 Next j
 j = j - 1
Else
 j = TheIndex
End If
'TODO:fix some bug
'move it
cObj.MovePageEx ThePrj, TheIndex, j, i, VarPtr(ThePageState(1))
TheIndex = i
pListRedraw
End Sub

Private Sub mnuResetP_Click()
Dim pp As typeOperatorProp_DesignTime
Dim i As Long, j As Long
Dim s As String
If nSelType > 0 And nSelIndex > 0 And nSelPropIndex > 0 Then
 Select Case tDef(nSelType).props(nSelPropIndex).nType
 Case eOPT_Color
  If nSelPropSubIndex > 0 And nSelPropSubIndex <= 4 Then
   'get sub properties
   i = Val(tDef(nSelType).props(nSelPropIndex).sDefault)
   CopyMemory j, ByVal (VarPtr(i) + 4 - nSelPropSubIndex), 1&
   s = CStr(j)
  Else 'reset all
   i = Val(tDef(nSelType).props(nSelPropIndex).sDefault)
   CopyMemory ThePrj.Operators(nSelIndex).bProps(tDef(nSelType).props(nSelPropIndex).nOffset), i, 4&
   'calc bitmap
   pPropRedraw
   cObj.SetNotInMemoryFlags ThePrj, nSelIndex
   pViewRedraw
   j = -1
  End If
 Case eOPT_Pt, eOPT_PtInt, eOPT_PtByte, eOPT_Rect, eOPT_RectInt, eOPT_RectByte, _
 eOPT_PtFloat, eOPT_PtHalf, eOPT_RectFloat, eOPT_RectHalf
  If nSelPropSubIndex > 0 And nSelPropSubIndex <= 4 Then
   s = tDef(nSelType).props(nSelPropIndex).sDefault
   'get sub properties
   For j = 1 To nSelPropSubIndex - 1
    i = InStr(i + 1, s, ";")
    If i = 0 Then
     j = -1
     Exit For
    End If
   Next j
   If j >= 0 Then
    j = InStr(i + 1, s, ";")
    If j > 0 Then
     s = Mid(s, i + 1, j - i - 1)
    Else
     s = Mid(s, i + 1)
    End If
   End If
  Else 'reset all
   PropFromString tDef(nSelType).props(nSelPropIndex).sDefault, tDef(nSelType).props(nSelPropIndex), pp
   PropWrite ThePrj.Operators(nSelIndex), tDef(nSelType).props(nSelPropIndex), pp
   'calc bitmap
   pPropRedraw
   cObj.SetNotInMemoryFlags ThePrj, nSelIndex
   pViewRedraw
   j = -1
  End If
 Case Else
  'get default value
  s = tDef(nSelType).props(nSelPropIndex).sDefault
 End Select
 If j >= 0 Then pChangeValue s
End If
End Sub

Private Sub mnuSave_Click()
If TheFileName = "" Then mnuSA_Click Else pSave TheFileName
End Sub

Private Sub mnuSA_Click()
Dim s As String
If cd.VBGetSaveFileName(s, , , "MyTexture|*.myt", , CStr(App.Path), , "myt", Me.hwnd) Then pSave s
End Sub

Private Sub pOpen(ByVal s As String)
Dim i As Long
pClearLockCount 'fix the unknown bug
cObj.Clear ThePrj
nSelIndex = 0
nSelType = 0
nShowIndex = 0
If Not LoadPrjFile(ThePrj, s) Then
 MsgBox "Failed!", vbCritical
 TheFileName = ""
 pDelMRU s
 stb1.PanelCaption(1) = "Failed to load " + s
Else
 TheFileName = s
 pAddMRU s
 stb1.PanelCaption(1) = "Done."
End If
VisiblePageCount = ThePrj.nPageCount
If VisiblePageCount > 0 Then
 ReDim ThePageState(1 To VisiblePageCount)
Else
 Erase ThePageState
End If
TheIndex = -1
pTitle
pListRefresh
pPropRedraw
pViewRedraw
pObjRedraw
pMRU
End Sub

Private Sub pAddMRU(ByVal s As String)
Dim i As Long, j As Long
Dim s2 As String, s3 As String
s2 = CStr(App.Path)
If Right(s2, 1) <> "\" Then s2 = s2 + "\"
s3 = String(1024, vbNullChar)
PathRelativePathTo s3, s2, 0, s, 0
s3 = Replace(s3, vbNullChar, "")
If s3 = "" Then s3 = s
For i = 1 To 8
 s2 = cSet.GetSettings("MRU" + CStr(i))
 If s2 = "" Then Exit For
 If StrComp(s2, s3, vbTextCompare) = 0 Then
  If i > 1 Then
   For j = i - 1 To 1 Step -1
    s2 = cSet.GetSettings("MRU" + CStr(j))
    cSet.SetSettings "MRU" + CStr(j + 1), s2
   Next j
   cSet.SetSettings "MRU1", s3
  End If
  Exit Sub
 End If
Next i
For i = 7 To 1 Step -1
 s2 = cSet.GetSettings("MRU" + CStr(i))
 cSet.SetSettings "MRU" + CStr(i + 1), s2
Next i
cSet.SetSettings "MRU1", s3
End Sub

Private Sub pDelMRU(ByVal s As String)
Dim i As Long, j As Long
Dim s2 As String, s3 As String
s2 = CStr(App.Path)
If Right(s2, 1) <> "\" Then s2 = s2 + "\"
s3 = String(1024, vbNullChar)
PathRelativePathTo s3, s2, 0, s, 0
s3 = Replace(s3, vbNullChar, "")
If s3 = "" Then s3 = s
For i = 1 To 8
 s2 = cSet.GetSettings("MRU" + CStr(i))
 If s2 = "" Then Exit For
 If StrComp(s2, s3, vbTextCompare) = 0 Then
  For j = i To 7
   s = cSet.GetSettings("MRU" + CStr(j + 1))
   cSet.SetSettings "MRU" + CStr(j), s
  Next j
  cSet.SetSettings "MRU8", ""
  Exit Sub
 End If
Next i
End Sub

Friend Sub pMRU()
Dim i As Long, ii As Long, j As Long, k As Long
Dim s As String
Dim s2 As String
For i = 1 To 8
 s = cSet.GetSettings("MRU" + CStr(i))
 If s = "" Then Exit For
 s2 = String(1024, vbNullChar)
 PathCompactPathExW ByVal StrPtr(s2), ByVal StrPtr(s), 32, 0
 s = "&" + CStr(i) + " " + Replace(s2, vbNullChar, "")
 With fm1
  j = .FindMenu("_file")
  k = .FindButton(j, "mru_" + CStr(i))
  .ButtonFlags(j, k) = 0
  .ButtonCaption(j, k) = s
  j = .FindMenu("_mru")
  k = .FindButton(j, "mru_" + CStr(i))
  .ButtonFlags(j, k) = 0
  .ButtonCaption(j, k) = s
 End With
Next i
For ii = i To 8
 With fm1
  j = .FindMenu("_file")
  k = .FindButton(j, "mru_" + CStr(ii))
  .ButtonFlags(j, k) = fbtfHidden
  j = .FindMenu("_mru")
  k = .FindButton(j, "mru_" + CStr(ii))
  .ButtonFlags(j, k) = fbtfHidden
 End With
Next ii
i = (i <= 1) And 1&
With fm1
 j = .FindMenu("_file")
 k = .FindButton(j, "mru_0")
 .ButtonFlags(j, k) = i
 j = .FindMenu("_mru")
 k = .FindButton(j, "mru_0")
 .ButtonFlags(j, k) = i
 k = .FindButton(j, "clrR")
 .ButtonFlags(j, k) = i
End With
End Sub

Private Sub pSave(ByVal s As String)
Dim i As Long
Me.MousePointer = vbHourglass
DoEvents
i = Val(cSet.GetSettings("CompressMode", "1"))
If Not SavePrjFile_1(ThePrj, s, i) Then
'If Not SavePrjFile(ThePrj, s, i) Then
 MsgBox "Failed!", vbCritical
 stb1.PanelCaption(1) = "Failed to save " + TheFileName
Else
 TheFileName = s
 pAddMRU s
 stb1.PanelCaption(1) = "Done."
 pTitle
 pMRU
End If
Me.MousePointer = vbDefault
End Sub

Private Sub mnuSaveBmp_Click()
Dim s As String
Dim idx As Long
With ThePrj
 If nShowIndex > 0 And nShowIndex <= .nOpCount Then
  With .Operators(nShowIndex)
   If .Flags >= 0 Then
    If (.Flags And int_OpFlags_Error) = 0 And (.Flags And int_OpFlags_InMemory) Then
     idx = .nBmIndex
     If idx < 0 Then idx = -idx
     Debug.Assert idx > 0 And idx <= cObj.BitmapCount
     If cd.VBGetSaveFileName(s, , , "Bitmap|*.bmp", , CStr(App.Path), , "bmp", Me.hwnd) Then
      If Not cObj.TheBitmap(idx).SavePicture(s) Then MsgBox "Failed!", vbCritical
     End If
    End If
   End If
  End With
 End If
End With
End Sub

Private Sub mnuShowOp_Click()
Dim i As Long
If TheIndex > 0 And TheIndex <= ThePrj.nPageCount Then
 i = cObj.PageHitTest(ThePrj, TheIndex, nXCur, nYCur)
 If i > 0 Then
  i = ThePrj.Pages(TheIndex).Rows(nYCur).idxOp(i)
  With ThePrj.Operators(i)
   If nXCur >= .Left + .Width Then i = 0
  End With
 End If
Else
 i = 0
End If
If i <> nShowIndex Then
 nShowIndex = i
 pViewRedraw
 pObjRedraw
End If
End Sub

Private Sub objDrag_Change(ByVal iDelta As Long, ByVal Button As Long, ByVal Shift As Long, bCancel As Boolean)
Dim i As Long, j As Long
Dim pp As typeOperatorProp_DesignTime
Dim fDelta As Single
Dim nOffset As Long
If nSelType > 0 And nSelIndex > 0 And nSelPropIndex > 1 And Button = 1 Then
 With tDef(nSelType)
  If nSelPropIndex <= .PropCount Then
   PropRead ThePrj.Operators(nSelIndex), tDef(nSelType).props(nSelPropIndex), pp
   With .props(nSelPropIndex)
    If .ListCount > 0 Then Exit Sub
    fDelta = iDelta / 1000 '??
    Select Case .nType
    Case eOPT_Byte, eOPT_Integer, eOPT_Long
     'auto determine ?? :-3
     If .sMin <> "" And .sMax <> "" Then
      If Val(.sMax) - Val(.sMin) < 32 Then
       iDelta = iDelta \ 8&
       If iDelta = 0 Then
        bCancel = True
        Exit Sub
       End If
      End If
     End If
     pp.iValue(0) = pp.iValue(0) + iDelta
    Case eOPT_Half, eOPT_Single
     pp.fValue(0) = pp.fValue(0) + fDelta
    Case eOPT_Color
     nOffset = .nOffset
     With ThePrj.Operators(nSelIndex)
      If Shift And vbShiftMask Then
       For i = nOffset To nOffset + 3
        j = .bProps(i) + iDelta
        If j < 0 Then j = 0
        If j > 255 Then j = 255
        .bProps(i) = j
       Next i
      ElseIf nSelPropSubIndex > 0 Then
       j = .bProps(nOffset + 4 - nSelPropSubIndex) + iDelta
       If j < 0 Then j = 0
       If j > 255 Then j = 255
       .bProps(nOffset + 4 - nSelPropSubIndex) = j
      End If
      CopyMemory pp.iValue(0), .bProps(nOffset), 4&
     End With
    Case eOPT_Pt, eOPT_PtInt, eOPT_PtByte, eOPT_Rect, eOPT_RectInt, eOPT_RectByte
     If Shift And vbShiftMask Then
      For i = 0 To 3
       pp.iValue(i) = pp.iValue(i) + iDelta
      Next i
     ElseIf nSelPropSubIndex > 0 Then
      pp.iValue(nSelPropSubIndex - 1) = pp.iValue(nSelPropSubIndex - 1) + iDelta
     End If
    Case eOPT_PtFloat, eOPT_PtHalf, eOPT_RectFloat, eOPT_RectHalf
     If Shift And vbShiftMask Then
      For i = 0 To 3
       pp.fValue(i) = pp.fValue(i) + fDelta
      Next i
     ElseIf nSelPropSubIndex > 0 Then
      pp.fValue(nSelPropSubIndex - 1) = pp.fValue(nSelPropSubIndex - 1) + fDelta
     End If
    Case Else 'string,etc.
     Exit Sub
    End Select
   End With
   PropWrite ThePrj.Operators(nSelIndex), tDef(nSelType).props(nSelPropIndex), pp
   pPropRedraw
   'calc bitmap
   cObj.SetNotInMemoryFlags ThePrj, nSelIndex
   pViewRedraw
  End If
 End With
End If
End Sub

Private Sub pListExpand(Optional ByVal nIndex As Long = -1)
Dim i As Long, j As Long, jj As Long
Dim k As Long
Dim m As Long
If nIndex < 0 Then nIndex = TheIndex
If nIndex <= 0 Or nIndex >= ThePrj.nPageCount Then Exit Sub
jj = &H7FFFFFFF
If ThePageState(nIndex) = 1 Then
 j = ThePrj.Pages(nIndex).nIndent
 For i = nIndex + 1 To ThePrj.nPageCount
  k = ThePrj.Pages(i).nIndent
  If k <= j Then Exit For
  If ThePageState(i) = 3 And jj >= k Then 'hide and collapsed
   ThePageState(i) = 1
   m = m + 1
   jj = k
  ElseIf ThePageState(i) >= 2 And k <= jj Then
   ThePageState(i) = ThePageState(i) - 2
   m = m + 1
   jj = &H7FFFFFFF
  End If
 Next i
 If m > 0 Then
  ThePageState(nIndex) = 0
  VisiblePageCount = VisiblePageCount + m
  pListRefresh
 End If
End If
End Sub

Private Sub pListCollapse(Optional ByVal nIndex As Long = -1)
Dim i As Long, j As Long
Dim m As Long
If nIndex < 0 Then nIndex = TheIndex
If nIndex <= 0 Or nIndex >= ThePrj.nPageCount Then Exit Sub
If ThePageState(nIndex) = 0 Then
 j = ThePrj.Pages(nIndex).nIndent
 For i = nIndex + 1 To ThePrj.nPageCount
  If ThePrj.Pages(i).nIndent <= j Then Exit For
  If ThePageState(i) < 2 Then
   ThePageState(i) = ThePageState(i) + 2
   m = m + 1
  End If
 Next i
 If m > 0 Then
  ThePageState(nIndex) = 1
  VisiblePageCount = VisiblePageCount - m
  pListRefresh
 End If
End If
End Sub

Private Sub picList_KeyDown(KeyCode As Integer, Shift As Integer)
Dim i As Long
Select Case Shift
Case 0
 Select Case KeyCode
 Case vbKeyDelete
  mnuDelP_Click
 Case vbKeyUp
  If TheIndex <= 1 Or TheIndex > ThePrj.nPageCount Then Exit Sub
  Do
   TheIndex = TheIndex - 1
  Loop While TheIndex > 1 And ThePageState(TheIndex) >= 2
  lstPage_Click
  pListRedraw
 Case vbKeyDown
  If TheIndex <= 0 Or TheIndex >= ThePrj.nPageCount Then Exit Sub
  i = TheIndex
  Do
   TheIndex = TheIndex + 1
   If TheIndex > ThePrj.nPageCount Then
    TheIndex = i
    Exit Sub
   End If
  Loop While ThePageState(TheIndex) >= 2
  lstPage_Click
  pListRedraw
 Case vbKeyLeft
  If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
  If TheIndex < ThePrj.nPageCount Then
   If ThePageState(TheIndex) = 0 And _
   ThePrj.Pages(TheIndex).nIndent < ThePrj.Pages(TheIndex + 1).nIndent Then
    pListCollapse
    Exit Sub
   End If
  End If
  'move left
  i = ThePrj.Pages(TheIndex).nIndent
  Do
   TheIndex = TheIndex - 1
   If ThePageState(TheIndex) < 2 And ThePrj.Pages(TheIndex).nIndent < i Then Exit Do
  Loop While TheIndex > 1
  lstPage_Click
  pListRedraw
 Case vbKeyRight
  If TheIndex <= 0 Or TheIndex >= ThePrj.nPageCount Then Exit Sub
  If ThePageState(TheIndex) = 1 And _
  ThePrj.Pages(TheIndex).nIndent < ThePrj.Pages(TheIndex + 1).nIndent Then
   pListExpand
   Exit Sub
  End If
  'move right
  If ThePrj.Pages(TheIndex).nIndent < ThePrj.Pages(TheIndex + 1).nIndent And ThePageState(TheIndex + 1) < 2 Then
   TheIndex = TheIndex + 1
   lstPage_Click
   pListRedraw
  End If
 Case vbKeyA
  fm1.Click "[P2]", "addp"
 Case vbKeyR
  fm1.Click "[P2]", "renp"
 End Select
Case vbShiftMask
 Select Case KeyCode
 Case vbKeyUp
  mnuPMoveUp_Click
 Case vbKeyDown
  mnuPMoveDown_Click
 Case vbKeyLeft
  fm1.Click "[P2]", "p2_l"
 Case vbKeyRight
  fm1.Click "[P2]", "p2_r"
 End Select
End Select
End Sub

Private Sub picList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, j As Long
Dim xx As Long
Select Case Button
Case 1
 j = y
 j = (j + sbList.Value(efsVertical)) \ 16&
 xx = x
 For i = 1 To ThePrj.nPageCount
  If j < 0 Then Exit For
  If ThePageState(i) < 2 Then
   If j = 0 Then
    'expand??
    If i < ThePrj.nPageCount Then
     j = ThePrj.Pages(i).nIndent
     If j < ThePrj.Pages(i + 1).nIndent And xx >= j * 8& And xx < j * 8& + 16& Then
      If ThePageState(i) = 0 Then
       pListCollapse i
      Else
       pListExpand i
      End If
     End If
    End If
    If TheIndex <> i Then
     TheIndex = i
     pListRedraw
     lstPage_Click
    End If
    Exit For
   End If
   j = j - 1
  End If
 Next i
Case 2
 fm1.PopupMenu "[P2]"
End Select
End Sub

Private Sub picList_Paint()
bmList.PaintPicture picList.hdc
End Sub

Private Sub picList_Resize()
On Error Resume Next
bmList.Create picList.ScaleWidth, picList.ScaleHeight
pListRefresh
End Sub

Private Sub pListRefresh()
Dim h As Long, hh As Long
hh = picList.ScaleHeight
h = VisiblePageCount * 16& - hh
With sbList
 If h > 0 Then
  .Enabled(efsVertical) = True
  .Max(efsVertical) = h
  .LargeChange(efsVertical) = hh
  .SmallChange(efsVertical) = 16&
 Else
  .Enabled(efsVertical) = False
  .Max(efsVertical) = 0
  .Value(efsVertical) = 0
 End If
End With
pListRedraw
End Sub

Private Sub picObj_DblClick()
If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
If nSelIndex < 0 Then pEditComment
End Sub

Private Sub pEditComment()
Dim idx As Long
If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
idx = -nSelIndex
With ThePrj.Pages(TheIndex)
 If idx > 0 And idx <= .nCommentCount Then
  Dim frm As New frmComment
  frm.fSetComment .Comments(idx)
  frm.Show 1
  frm.fGetComment .Comments(idx)
  pObjRedraw
 End If
End With
End Sub

Private Sub pDeleteComment(Optional ByVal nPage As Long, Optional ByVal nIndex As Long)
Dim i As Long
If nPage = 0 Then nPage = TheIndex
If nIndex = 0 Then nIndex = -nSelIndex
If nPage > 0 And nPage <= ThePrj.nPageCount Then
 With ThePrj.Pages(nPage)
  If nIndex > 0 And nIndex <= .nCommentCount Then
   If .nCommentCount <= 1 Then
    Erase .Comments
    .nCommentCount = 0
   Else
    .nCommentCount = .nCommentCount - 1
    For i = nIndex To .nCommentCount
     .Comments(i) = .Comments(i + 1)
    Next i
    ReDim Preserve .Comments(1 To .nCommentCount)
   End If
  End If
 End With
End If
End Sub

Private Sub picObj_KeyDown(KeyCode As Integer, Shift As Integer)
Dim idx As Long
If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
If nCurFlags Then Exit Sub
Select Case Shift
Case 0
 Select Case KeyCode
 Case vbKeyA
  bShowMenu = True
 Case vbKeyX
  mnuEdit1_Click 0
 Case vbKeyC
  mnuEdit1_Click 1
 Case vbKeyV
  mnuEdit1_Click 2
 Case vbKeyDelete
  mnuEdit1_Click 3
 Case vbKeyS
  mnuShowOp_Click
 Case vbKeyY
  fm1.Click "_tools", "beye"
 Case vbKeyLeft
  If nXCur > 0 Then
   nXCur = nXCur - 1
   pShowPos
   idx = nXCur * 16& - sb1.Value(efsHorizontal)
   If idx >= 0 And idx < sb1.LargeChange(efsHorizontal) Then
    pObjRedraw
   Else
    sb1.Value(efsHorizontal) = nXCur * 16& - sb1.LargeChange(efsHorizontal)
   End If
  End If
 Case vbKeyRight
  If nXCur < int_Page_Width - 1 Then
   nXCur = nXCur + 1
   pShowPos
   idx = nXCur * 16& - sb1.Value(efsHorizontal)
   If idx >= 0 And idx < sb1.LargeChange(efsHorizontal) Then
    pObjRedraw
   Else
    sb1.Value(efsHorizontal) = nXCur * 16&
   End If
  End If
 Case vbKeyUp
  If nYCur > 0 Then
   nYCur = nYCur - 1
   pShowPos
   idx = nYCur * 16& - sb1.Value(efsVertical)
   If idx >= 0 And idx < sb1.LargeChange(efsVertical) Then
    pObjRedraw
   Else
    sb1.Value(efsVertical) = nYCur * 16& - sb1.LargeChange(efsVertical)
   End If
  End If
 Case vbKeyDown
  If nYCur < int_Page_Height - 1 Then
   nYCur = nYCur + 1
   pShowPos
   idx = nYCur * 16& - sb1.Value(efsVertical)
   If idx >= 0 And idx < sb1.LargeChange(efsVertical) Then
    pObjRedraw
   Else
    sb1.Value(efsVertical) = nYCur * 16&
   End If
  End If
 Case vbKeyPageUp
  sb1.Value(efsVertical) = sb1.Value(efsVertical) - sb1.LargeChange(efsVertical)
 Case vbKeyPageDown
  sb1.Value(efsVertical) = sb1.Value(efsVertical) + sb1.LargeChange(efsVertical)
 Case vbKeyHome
  sb1.Value(efsVertical) = 0
 Case vbKeyEnd
  sb1.Value(efsVertical) = sb1.Max(efsVertical)
 End Select
Case vbCtrlMask
 Select Case KeyCode
 Case vbKeyA
  mnuEdit1_Click 4
 End Select
Case vbShiftMask
 Select Case KeyCode
 Case vbKeyA
  mnuAddC_Click
 Case vbKeyPageUp
  sb1.Value(efsHorizontal) = sb1.Value(efsHorizontal) - sb1.LargeChange(efsHorizontal)
 Case vbKeyPageDown
  sb1.Value(efsHorizontal) = sb1.Value(efsHorizontal) + sb1.LargeChange(efsHorizontal)
 Case vbKeyHome
  sb1.Value(efsHorizontal) = 0
 Case vbKeyEnd
  sb1.Value(efsHorizontal) = sb1.Max(efsHorizontal)
 End Select
End Select
End Sub

Private Sub picObj_KeyUp(KeyCode As Integer, Shift As Integer)
If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
If nCurFlags Then Exit Sub
Select Case Shift
Case 0
 Select Case KeyCode
 Case vbKeyA
  If bShowMenu Then
   bShowMenu = False
   'Me.PopupMenu mnuAddPopup
   'new!!
   fm1.PopupMenu "\"
  End If
 End Select
End Select
End Sub

Private Sub picObj_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, j As Long
Dim X1 As Long, X2 As Long, x3 As Long
Dim Y1 As Long, Y2 As Long, y3 As Long
Dim dX As Long, dy As Long
'dim idx As Long
If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
'select/deselect item
If (Shift And vbCtrlMask) = 0 Then cObj.ClearFlags ThePrj
dX = sb1.Value(efsHorizontal)
dy = sb1.Value(efsVertical)
nXCur = (x + dX) \ 16&
nYCur = (y + dy) \ 16&
'show position
pShowPos
'hit test
i = cObj.PageHitTest(ThePrj, TheIndex, nXCur, nYCur)
If i > 0 Then
 i = ThePrj.Pages(TheIndex).Rows(nYCur).idxOp(i)
 With ThePrj.Operators(i)
  If nXCur < .Left + .Width Then
   .Flags = .Flags Or int_OpFlags_Selected
   j = .Left * 16& - dX
   If x < j + 4 Then 'resize left
    nCurFlags = 3
   ElseIf x < j + .Width * 16& - 4 Then 'move
    'TODO:multiple move
    nXEnd = x - .Left * 16& - 8&
    nYEnd = y - .Top * 16& - 8&
    nCurFlags = 2
   Else 'resize right
    nCurFlags = 4
   End If
  Else
   i = 0
  End If
 End With
End If
If i = 0 Then
 'comments hit test
 j = 0
 With ThePrj.Pages(TheIndex)
  For i = .nCommentCount To 1 Step -1
   With .Comments(i)
    X1 = .Left * 16& - dX - 4&
    X2 = X1 + .Width * 8&
    x3 = X1 + .Width * 16&
    Y1 = .Top * 16& - dy - 4&
    Y2 = Y1 + .Height * 8&
    y3 = Y1 + .Height * 16&
    If x >= X1 And x <= x3 + 8 And y >= Y1 And y <= y3 + 8 Then
     If x <= X1 + 8 Or x >= x3 Or y <= Y1 + 8 Or y >= y3 Then
      'TODO:move?
      If x >= X1 And x <= X1 + 8 And y >= Y2 And y <= Y2 + 8 Then
       nCurFlags = 3 'left
       picObj.MousePointer = vbSizeWE
      ElseIf x >= x3 And x <= x3 + 8 And y >= Y2 And y <= Y2 + 8 Then
       nCurFlags = 4 'right
       picObj.MousePointer = vbSizeWE
      ElseIf x >= X2 And x <= X2 + 8 And y >= Y1 And y <= Y1 + 8 Then
       nCurFlags = 5 'top
       picObj.MousePointer = vbSizeNS
      ElseIf x >= X2 And x <= X2 + 8 And y >= y3 And y <= y3 + 8 Then
       nCurFlags = 6 'bottom
       picObj.MousePointer = vbSizeNS
      ElseIf x >= X1 And x <= X1 + 8 And y >= Y1 And y <= Y1 + 8 Then
       nCurFlags = 7 'top-left
       picObj.MousePointer = vbSizeNWSE
      ElseIf x >= x3 And x <= x3 + 8 And y >= Y1 And y <= Y1 + 8 Then
       nCurFlags = 8 'top-right
       picObj.MousePointer = vbSizeNESW
      ElseIf x >= X1 And x <= X1 + 8 And y >= y3 And y <= y3 + 8 Then
       nCurFlags = 9 'bottom-left
       picObj.MousePointer = vbSizeNESW
      ElseIf x >= x3 And x <= x3 + 8 And y >= y3 And y <= y3 + 8 Then
       nCurFlags = 10 'bottom-right
       picObj.MousePointer = vbSizeNWSE
      Else
       nXEnd = x - .Left * 16& - 8&
       nYEnd = y - .Top * 16& - 8&
       nCurFlags = 2 'move
       picObj.MousePointer = vbSizePointer
      End If
      'over
      j = 1
     End If
    End If
   End With
   If j Then Exit For
  Next i
 End With
 If j Then i = -i Else i = 0
End If
'show properties
If i <> 0 Then
 If nSelIndex <> i Then
  nSelIndex = i
  If i > 0 Then
   nSelType = ThePrj.Operators(i).nType
   nSelPropIndex = 0
  End If
  pPropRefresh
 End If
Else
 If nSelIndex <> 0 Then
  nSelIndex = 0
  nSelType = 0
  nSelPropIndex = 0
  pPropRefresh
 End If
 If Button = 1 Then
  nCurFlags = 1
  nXEnd = nXCur
  nYEnd = nYCur
 End If
End If
'redraw
pObjRedraw
'TODO:
If Button = 2 Then
 'new menu test
 fm1.PopupMenu "[P1]"
End If
End Sub

Private Sub picObj_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, j As Long, m As Long
Dim dX As Long, dy As Long
'Dim idx As Long
Dim tmp() As Long
If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then
 picObj.MousePointer = vbDefault
 Exit Sub
End If
dX = sb1.Value(efsHorizontal)
dy = sb1.Value(efsVertical)
i = (x + dX) \ 16&
j = (y + dy) \ 16&
If Button = 0 Then
 'check cursor
 picObj.MousePointer = vbDefault
 If i >= 0 And j >= 0 And i < int_Page_Width And j < int_Page_Height Then
  m = cObj.PageHitTest(ThePrj, TheIndex, i, j)
  If m > 0 Then
   With ThePrj.Operators(ThePrj.Pages(TheIndex).Rows(j).idxOp(m))
    i = .Left * 16& - dX
    If x < i + 4 Then 'resize left
     picObj.MousePointer = vbSizeWE
    ElseIf x < i + .Width * 16& - 4 Then 'move
     picObj.MousePointer = vbSizePointer
    ElseIf x < i + .Width * 16& Then 'resize right
     picObj.MousePointer = vbSizeWE
    End If
   End With
  End If
 End If
End If
If Button = 1 Then
 Select Case nCurFlags
 Case 1 'select area?
  If i >= 0 And j >= 0 And i < int_Page_Width And j < int_Page_Height And (i <> nXEnd Or j <> nYEnd) Then
   nXEnd = i
   nYEnd = j
   pShowPos
   pObjRedraw
  End If
 Case 2 'move
  'TODO:multiple move
  If nSelIndex > 0 Then
   i = (x - nXEnd) \ 16&
   j = (y - nYEnd) \ 16&
   With ThePrj.Operators(nSelIndex)
    If i < 0 Then i = 0
    If i > int_Page_Width - .Width Then i = int_Page_Width - .Width
    If j < 0 Then j = 0
    If j >= int_Page_Height Then i = int_Page_Height - 1
    If i <> .Left Or j <> .Top Then
     'get refresh area
     If j <> .Top Then 'move up/down
      m = cObj.PageHitTestEx(ThePrj, .nPage, .Left, .Top + 1, .Left + .Width - 1, .Top + 1, tmp)
     ElseIf i < .Left Then 'move left, refresh right
      m = cObj.PageHitTestEx(ThePrj, .nPage, i + .Width, .Top + 1, .Left + .Width - 1, .Top + 1, tmp)
     Else 'move right, refresh left
      m = cObj.PageHitTestEx(ThePrj, .nPage, .Left, .Top + 1, i - 1, .Top + 1, tmp)
     End If
     'move it
     If cObj.MoveOperatorByIndex(ThePrj, nSelIndex, i, j) Then
      'validate operator
      For i = 1 To m
       cObj.ValidateOps ThePrj, tmp(i)
      Next i
      cObj.ValidateOps ThePrj, nSelIndex
      'calc bitmap
      pViewRedraw
      pObjRedraw
     End If
    End If
   End With
  ElseIf nSelIndex < 0 Then 'comment
   i = (x - nXEnd) \ 16&
   j = (y - nYEnd) \ 16&
   With ThePrj.Pages(TheIndex).Comments(-nSelIndex)
    If i < 0 Then i = 0
    If i > int_Page_Width - .Width Then i = int_Page_Width - .Width
    If j < 0 Then j = 0
    If j >= int_Page_Height - .Height Then j = int_Page_Height - .Height
    If i <> .Left Or j <> .Top Then
     .Left = i
     .Top = j
     pObjRedraw
    End If
   End With
  End If
 Case 3 'resize left
  If nSelIndex > 0 Then
   With ThePrj.Operators(nSelIndex)
    i = (x + dX + 8) \ 16&
    If i < .Left + .Width And i <> .Left Then
     'get refresh area
     m = cObj.PageHitTestEx(ThePrj, .nPage, .Left, .Top + 1, i - 1, .Top + 1, tmp) 'only when get smaller
     'move it
     If cObj.MoveOperatorByIndex(ThePrj, nSelIndex, i, .Top, .Left + .Width - i) Then
      'validate operator
      For i = 1 To m
       cObj.ValidateOps ThePrj, tmp(i)
      Next i
      cObj.ValidateOps ThePrj, nSelIndex
      'calc bitmap
      pViewRedraw
      pObjRedraw
     End If
    End If
   End With
  ElseIf nSelIndex < 0 Then 'comment
   With ThePrj.Pages(TheIndex).Comments(-nSelIndex)
    i = (x + dX + 8) \ 16&
    If i >= 0 And i < .Left + .Width And i <> .Left Then
     .Width = .Left + .Width - i
     .Left = i
     pObjRedraw
    End If
   End With
  End If
 Case 4 'resize right
  If nSelIndex > 0 Then
   With ThePrj.Operators(nSelIndex)
    i = (x + dX + 8) \ 16& - .Left
    If i > 0 And i <> .Width Then
     'get refresh area
     m = cObj.PageHitTestEx(ThePrj, .nPage, .Left + i, .Top + 1, .Left + .Width - 1, .Top + 1, tmp) 'only when get smaller
     'move it
     If cObj.MoveOperatorByIndex(ThePrj, nSelIndex, .Left, .Top, i) Then
      'validate operator
      For i = 1 To m
       cObj.ValidateOps ThePrj, tmp(i)
      Next i
      cObj.ValidateOps ThePrj, nSelIndex
      'calc bitmap
      pViewRedraw
      pObjRedraw
     End If
    End If
   End With
  ElseIf nSelIndex < 0 Then 'comment
   With ThePrj.Pages(TheIndex).Comments(-nSelIndex)
    i = (x + dX + 8) \ 16&
    If i > .Left And i <= int_Page_Width And i <> .Left + .Width Then
     .Width = i - .Left
     pObjRedraw
    End If
   End With
  End If
 Case 5 'resize top
  If nSelIndex < 0 Then
   With ThePrj.Pages(TheIndex).Comments(-nSelIndex)
    i = (y + dy + 8) \ 16&
    If i >= 0 And i < .Top + .Height And i <> .Top Then
     .Height = .Top + .Height - i
     .Top = i
     pObjRedraw
    End If
   End With
  End If
 Case 6 'resize bottom
  If nSelIndex < 0 Then
   With ThePrj.Pages(TheIndex).Comments(-nSelIndex)
    i = (y + dy + 8) \ 16&
    If i > .Top And i <= int_Page_Height And i <> .Top + .Height Then
     .Height = i - .Top
     pObjRedraw
    End If
   End With
  End If
 Case 7 To 10 'resize top-left
  If nSelIndex < 0 Then
   With ThePrj.Pages(TheIndex).Comments(-nSelIndex)
    i = (x + dX + 8) \ 16&
    j = (y + dy + 8) \ 16&
    If nCurFlags And 1 Then 'left
     m = (i >= 0 And i < .Left + .Width And i <> .Left)
    Else 'right
     m = (i > .Left And i <= int_Page_Width And i <> .Left + .Width)
    End If
    If (nCurFlags - 1) And 2 Then 'top
     m = m + 2& * (j >= 0 And j < .Top + .Height And j <> .Top)
    Else 'bottom
     m = m + 2& * (j > .Top And j <= int_Page_Height And j <> .Top + .Height)
    End If
    m = -m
    If m Then
     If m And 1 Then
      If nCurFlags And 1 Then 'left
       .Width = .Left + .Width - i
       .Left = i
      Else 'right
       .Width = i - .Left
      End If
     End If
     If m And 2 Then
      If (nCurFlags - 1) And 2 Then 'top
       .Height = .Top + .Height - j
       .Top = j
      Else 'bottom
       .Height = j - .Top
      End If
     End If
     pObjRedraw
    End If
   End With
  End If
 End Select
End If
End Sub

Private Sub picObj_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, j As Long
Dim tmp() As Long
'Dim idx As Long
If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
'TODO:
'select area?
If Button = 1 And nCurFlags = 1 Then
 If nXEnd > nXCur Then
  i = nXCur
 Else
  i = nXEnd
  nXEnd = nXCur
 End If
 If nYEnd > nYCur Then
  j = nYCur
 Else
  j = nYEnd
  nYEnd = nYCur
 End If
 cObj.PageHitTestEx ThePrj, TheIndex, i, j, nXEnd, nYEnd, tmp, True
 nCurFlags = 0
 pShowPos
 pObjRedraw
End If
'size?
If Button = 1 And nCurFlags >= 2 And nCurFlags <= 4 Then
 nCurFlags = 0
 picObj.MousePointer = vbDefault
End If
End Sub

Private Sub picObj_Paint()
bmObj.PaintPicture picObj.hdc
End Sub

Private Sub picObj_Resize()
On Error Resume Next
Dim i As Long
With sb1
 i = int_Page_WidthPixels - picObj.ScaleWidth
 If i > 0 Then
  .Max(efsHorizontal) = i
  .LargeChange(efsHorizontal) = picObj.ScaleWidth
  .Enabled(efsHorizontal) = True
 Else
  .Max(efsHorizontal) = 0
  .Value(efsHorizontal) = 0
  .Enabled(efsHorizontal) = False
 End If
 .SmallChange(efsHorizontal) = 16&
 i = int_Page_HeightPixels - picObj.ScaleHeight
 If i > 0 Then
  .Max(efsVertical) = i
  .LargeChange(efsVertical) = picObj.ScaleHeight
  .Enabled(efsVertical) = True
 Else
  .Max(efsVertical) = 0
  .Value(efsVertical) = 0
  .Enabled(efsVertical) = False
 End If
 .SmallChange(efsVertical) = 16&
End With
bmObj.Create picObj.ScaleWidth, picObj.ScaleHeight
pObjRedraw
End Sub

Private Sub picProp_DblClick()
mouse_event MOUSEEVENTF_LEFTDOWN, 0, 0, 0, 0
End Sub

Private Sub picProp_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
Dim i As Long, j As Long
Dim dy As Long, w As Long
Dim pp As typeOperatorProp_DesignTime
Dim bRedraw As Boolean, bPick As Boolean
Dim nLeft(4) As Long
If nSelType > 0 And nSelIndex > 0 Then
 dy = sb2.Value(efsVertical)
 i = (y + dy) \ int_Prop_Height
 If i <> nSelPropIndex Then
  nSelPropIndex = i
  nSelPropSubIndex = 0
  bRedraw = True
 End If
 'hide the drag box
 objDrag.Visible = False
 'show edit box?
 With tDef(nSelType)
  w = picProp.ScaleWidth
  If x > int_Caption_Width And x < w And i > 0 And i <= .PropCount Then
   dy = i * int_Prop_Height - dy - 1&
   With .props(i)
    'select sub-item? show drag box?
    Select Case .nType
    Case eOPT_String, eOPT_Custom
     If x >= w - 20 Then
      bRedraw = True
      bPick = True
     End If
    Case eOPT_Color
     nLeft(0) = int_Caption_Width
     nLeft(1) = w \ 4 + int_Caption_Width_3 - 5
     nLeft(2) = w \ 2 + int_Caption_Width_2 - 10
     nLeft(3) = w - w \ 4 + int_Caption_Width_1 - 15
     nLeft(4) = w - 20
     For j = 1 To 4
      If x < nLeft(j) Then Exit For
     Next j
     If j > 4 Then
      j = nSelPropSubIndex
      bRedraw = True
      If Button = 1 Then
       'show color picker
       CopyMemory pp.iValue(1), ThePrj.Operators(nSelIndex).bProps(.nOffset), 4
       pp.iValue(0) = ColorPicker(pp.iValue(1))
       'save data
       If pp.iValue(0) <> pp.iValue(1) Then
        CopyMemory ThePrj.Operators(nSelIndex).bProps(.nOffset), pp.iValue(0), 4
        'calc bitmap
        cObj.SetNotInMemoryFlags ThePrj, nSelIndex
        pViewRedraw
       End If
      End If
     End If
     If nSelPropSubIndex <> j Then
      nSelPropSubIndex = j
      bRedraw = True
     End If
     If bRedraw And j > 0 Then
      objDrag.Move nLeft(j) - 4, dy + 1, 4, int_Prop_Height - 1
      objDrag.Visible = True
     End If
    Case eOPT_Pt, eOPT_PtInt, eOPT_PtByte, eOPT_PtFloat, eOPT_PtHalf
     nLeft(0) = int_Caption_Width
     nLeft(1) = w \ 2 + int_Caption_Width_2
     nLeft(2) = w
     If x < nLeft(1) Then j = 1 Else j = 2
     If nSelPropSubIndex <> j Then
      nSelPropSubIndex = j
      bRedraw = True
      If j > 0 Then
       objDrag.Move nLeft(j) - 4, dy + 1, 4, int_Prop_Height - 1
       objDrag.Visible = True
      End If
     End If
    Case eOPT_Rect, eOPT_RectInt, eOPT_RectByte, eOPT_RectFloat, eOPT_RectHalf
     nLeft(0) = int_Caption_Width
     nLeft(1) = w \ 4 + int_Caption_Width_3
     nLeft(2) = w \ 2 + int_Caption_Width_2
     nLeft(3) = w - w \ 4 + int_Caption_Width_1
     nLeft(4) = w
     For j = 1 To 4
      If x >= nLeft(j - 1) And x < nLeft(j) Then Exit For
     Next j
     If nSelPropSubIndex <> j Then
      nSelPropSubIndex = j
      bRedraw = True
      If j > 0 Then
       objDrag.Move nLeft(j) - 4, dy + 1, 4, int_Prop_Height - 1
       objDrag.Visible = True
      End If
     End If
    Case eOPT_Byte, eOPT_Integer, eOPT_Long, eOPT_Half, eOPT_Single
     If .ListCount = 0 And bRedraw Then
      objDrag.Move w - 4, dy + 1, 4, int_Prop_Height - 1
      objDrag.Visible = True
     End If
    End Select
    'get value?
    If bPick Or Not bRedraw Or .ListCount > 0 Then
     PropRead ThePrj.Operators(nSelIndex), tDef(nSelType).props(i), pp
    End If
    'show value?
    If (Not bRedraw Or .ListCount > 0) And Button = 1 Then
     If .ListCount > 0 Then
      cmbProp.Clear
      For j = 0 To .ListCount - 1
       cmbProp.AddItem .List(j)
      Next j
      cmbProp.Tag = CStr(pp.iValue(0))
      cmbProp.ListIndex = pp.iValue(0)
      pPropShow cmbProp, int_Caption_Width, dy, w - int_Caption_Width, int_Prop_Height + 1
     Else
      Select Case .nType
      Case eOPT_Custom
       bPick = True
      Case eOPT_Name
       txtProp.Text = pp.sValue
       txtProp.Tag = txtProp.Text
       pPropShow txtProp, int_Caption_Width + 1, dy + 1, w - int_Caption_Width - 2
      Case eOPT_String
       If nSelType = int_OpType_Load Then
        bPick = True
       Else
        txtProp.Text = pp.sValue
        txtProp.Tag = txtProp.Text
        pPropShow txtProp, int_Caption_Width + 1, dy + 1, w - int_Caption_Width - 21&
       End If
      Case eOPT_Byte, eOPT_Integer, eOPT_Long
       txtProp.Text = CStr(pp.iValue(0))
       txtProp.Tag = txtProp.Text
       pPropShow txtProp, int_Caption_Width + 1, dy + 1, w - int_Caption_Width - 2
      Case eOPT_Half, eOPT_Single
       txtProp.Text = CStr(pp.fValue(0))
       txtProp.Tag = txtProp.Text
       pPropShow txtProp, int_Caption_Width + 1, dy + 1, w - int_Caption_Width - 2
      Case eOPT_Color
       If nSelPropSubIndex > 0 Then
        txtProp.Text = CStr(ThePrj.Operators(nSelIndex).bProps(.nOffset + 4 - nSelPropSubIndex))
        txtProp.Tag = txtProp.Text
        pPropShow txtProp, nLeft(nSelPropSubIndex - 1) + 1, dy + 1, nLeft(nSelPropSubIndex) - nLeft(nSelPropSubIndex - 1) - 1
       End If
      Case eOPT_Pt, eOPT_PtInt, eOPT_PtByte, eOPT_Rect, eOPT_RectInt, eOPT_RectByte
       If nSelPropSubIndex > 0 Then
        txtProp.Text = CStr(pp.iValue(nSelPropSubIndex - 1))
        txtProp.Tag = txtProp.Text
        pPropShow txtProp, nLeft(nSelPropSubIndex - 1) + 1, dy + 1, nLeft(nSelPropSubIndex) - nLeft(nSelPropSubIndex - 1) - 1
       End If
      Case eOPT_PtFloat, eOPT_PtHalf, eOPT_RectFloat, eOPT_RectHalf
       If nSelPropSubIndex > 0 Then
        txtProp.Text = CStr(pp.fValue(nSelPropSubIndex - 1))
        txtProp.Tag = txtProp.Text
        pPropShow txtProp, nLeft(nSelPropSubIndex - 1) + 1, dy + 1, nLeft(nSelPropSubIndex) - nLeft(nSelPropSubIndex - 1) - 1
       End If
      End Select
     End If
    End If
    'edit string?
    If bPick And Button = 1 Then
     Select Case .nType
     Case eOPT_Custom
      If pCustom(pp) Then
       'save
       PropWrite ThePrj.Operators(nSelIndex), tDef(nSelType).props(i), pp
       'calc bitmap
       cObj.SetNotInMemoryFlags ThePrj, nSelIndex
       pViewRedraw
       bRedraw = True
      End If
     Case Else
      If StringPicker(pp.sValue, ThePrj, cObj, nSelType = int_OpType_Load) Then
       PropWrite ThePrj.Operators(nSelIndex), tDef(nSelType).props(i), pp
       If nSelType = int_OpType_Load Then
        'validate operator
        cObj.ValidateOps ThePrj, nSelIndex
        pObjRedraw
       End If
       'calc bitmap
       cObj.SetNotInMemoryFlags ThePrj, nSelIndex
       pViewRedraw
       bRedraw = True
      End If
     End Select
    End If
   End With
  End If
 End With
 If bRedraw Then pPropRedraw
 'popup menu?
 If nSelPropIndex > 1 And Button = 2 Then
  With tDef(nSelType)
   If nSelPropIndex <= .PropCount Then
    Select Case .props(nSelPropIndex).nType
    Case eOPT_Group, eOPT_Custom, eOPT_String
    Case Else
     fm1.PopupMenu "[P3]"
    End Select
   End If
  End With
 End If
End If
End Sub

Private Function pCustom(pp As typeOperatorProp_DesignTime) As Boolean
Dim s As String
Select Case nSelType
Case 13 'IFSP
 Dim frm1 As frmIFSP
 Set frm1 = New frmIFSP
 frm1.TheValue = pp.sValue
 frm1.Show 1
 If Not frm1.bCancel Then
  s = frm1.TheValue
  If s <> pp.sValue Then
   pp.sValue = s
   pCustom = True
  End If
 End If
 Set frm1 = Nothing
Case 9 'import
 Dim frm2 As frmPic
 Set frm2 = New frmPic
 frm2.SetData pp.sValue
 frm2.Show 1
 If frm2.IsChanged Then
  frm2.GetData pp.sValue
  pCustom = True
 End If
 Set frm2 = Nothing
Case 12 'L-system(color only :-3)
 Dim frm3 As frmLClr
 Set frm3 = New frmLClr
 frm3.SetData pp.sValue
 frm3.Show 1
 If frm3.IsChanged Then
  frm3.GetData pp.sValue
  pCustom = True
 End If
 Set frm3 = Nothing
Case Else 'TODO:
End Select
End Function

Private Sub picProp_Paint()
bmProp.PaintPicture picProp.hdc, , , picProp.ScaleWidth, picProp.ScaleHeight
End Sub

Private Sub picProp_Resize()
On Error Resume Next
Dim w As Long, h As Long
'make the bitmap a little bigger :-3 so use Width instead of ScaleWidth
w = picProp.Width
h = picProp.Height
If w < bmProp.Width Then w = bmProp.Width
If h < bmProp.Height Then h = bmProp.Height
If w > bmProp.Width Or h > bmProp.Height Then
 bmProp.Create w, h
End If
pPropRefresh
End Sub

Private Sub pPropRefresh()
Dim bNone As Boolean
Dim i As Long
'hide selection
pPropHide
'check selected
If nSelType > 0 And nSelIndex > 0 Then
 'type,(name=prop1 stupid...)
 i = (1 + tDef(nSelType).PropCount) * int_Prop_Height
 i = i - picProp.ScaleHeight
End If
With sb2
 If i > 0 Then
  .Max(efsVertical) = i
  .LargeChange(efsVertical) = picProp.ScaleHeight
  .Enabled(efsVertical) = True
 Else
  .Max(efsVertical) = 0
  .Value(efsVertical) = 0
  .Enabled(efsVertical) = False
 End If
 .SmallChange(efsVertical) = int_Prop_Height
End With
pPropRedraw
End Sub

Private Sub pPropShow(obj As Object, ByVal x As Long, ByVal y As Long, ByVal w As Long, Optional ByVal h As Long)
On Error Resume Next
If h <= 0 Then h = int_Prop_Height - 1
With obj
 .Move x, y, w, h
 .Visible = True
 .SetFocus
End With
End Sub

Private Sub pPropHide()
cmbProp.Visible = False
txtProp.Visible = False
objDrag.Visible = False
End Sub

Private Sub pPropRedraw()
#If Office2003 = 0 Then
Dim hbr As Long
#End If
Dim r As RECT
Dim i As Long, m As Long, n As Long
Dim m2 As Long
Dim dy As Long
r.Right = picProp.ScaleWidth
r.Bottom = picProp.ScaleHeight
#If Office2003 = 0 Then
hbr = CreateSolidBrush(TranslateColor(vbButtonFace))
FillRect bmProp.hdc, r, hbr
DeleteObject hbr
#End If
If nSelType > 0 And nSelIndex > 0 Then
 dy = sb2.Value(efsVertical)
 m = r.Bottom \ int_Prop_Height + 1&
 n = dy \ int_Prop_Height
 m2 = tDef(nSelType).PropCount
 'draw contents
 For i = n To n + m
  r.Top = i * int_Prop_Height - dy
  If i > m2 Then Exit For
  If i <= 0 Then 'type
   #If Office2003 Then
   GradientFillRect bmProp.hdc, 0, r.Top, r.Right, int_Prop_Height, d_Title1, d_Title2, GRADIENT_FILL_RECT_V
   cFnt.DrawTextXP bmProp.hdc, "Type:" + tDef(nSelType).Name, 4, r.Top, r.Right, int_Prop_Height, DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS, vbWhite, , True
   #Else
   cFnt.DrawTextXP bmProp.hdc, "Type:" + tDef(nSelType).Name, 4, r.Top, r.Right, int_Prop_Height, DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS, vbBlack, , True
   #End If
   pPropRedrawGridline bmProp.hdc, 0, r.Top, r.Right, int_Prop_Height, 0
  Else
   pPropRedrawItem bmProp.hdc, 0, r.Top, r.Right, int_Prop_Height, i
  End If
 Next i
 #If Office2003 Then
 'calc size
 r.Top = m2 * int_Prop_Height + int_Prop_Height - dy
 #End If
End If
#If Office2003 Then
'office2k3 style draw background
If r.Top < r.Bottom Then
 GradientFillRect bmProp.hdc, 0, r.Top, r.Right, r.Bottom, d_Bar2, d_Bar1, GRADIENT_FILL_RECT_H
End If
#End If
picProp_Paint
End Sub

Private Sub pPropRedrawItem(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal i As Long)
Dim hbr As Long
Dim r As RECT
Dim j As Long
Dim pp As typeOperatorProp_DesignTime
Dim nLeft(4) As Long
With tDef(nSelType).props(i)
 If .nType = eOPT_Group Then
  'Group box!!!
  #If Office2003 Then
  GradientFillRect hdc, x, y, x + w, y + h, d_Title1, d_Title2, GRADIENT_FILL_RECT_V
  #Else
  hbr = CreateSolidBrush(TranslateColor(vbApplicationWorkspace))
  r.Left = x
  r.Top = y
  r.Right = x + w
  r.Bottom = y + h
  FillRect hdc, r, hbr
  DeleteObject hbr
  #End If
  cFnt.DrawTextXP hdc, tDef(nSelType).props(i).Name, x, y, w, h, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS, vbWhite, , True
 Else
  'normal
  #If Office2003 Then
  'selected?
  If nSelPropIndex = i Then
   GradientFillRect hdc, x, y, x + int_Caption_Width, y + h, d_Checked1, d_Checked2, GRADIENT_FILL_RECT_V
  Else
   GradientFillRect hdc, x, y, x + int_Caption_Width, y + h, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_V
  End If
  #End If
  'draw caption
  cFnt.DrawTextXP hdc, tDef(nSelType).props(i).Name, x + 4, y, int_Caption_Width, h, DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS, vbBlack, , True
  'erase background
  r.Left = x + int_Caption_Width
  r.Top = y
  r.Right = x + w
  r.Bottom = y + h
  #If Office2003 = 0 Then
  If .nType = eOPT_Color Or .nType = eOPT_String Then
   r.Right = r.Right - 20
  End If
  #End If
  hbr = GetStockObject(WHITE_BRUSH)
  FillRect hdc, r, hbr
 End If
 'selected subitem?
 Select Case .nType
 Case eOPT_Color
  nLeft(0) = x + int_Caption_Width
  nLeft(1) = x + w \ 4 + int_Caption_Width_3 - 5
  nLeft(2) = x + w \ 2 + int_Caption_Width_2 - 10
  nLeft(3) = x + w - w \ 4 + int_Caption_Width_1 - 15
  nLeft(4) = x + w - 20
 Case eOPT_Pt, eOPT_PtInt, eOPT_PtByte, eOPT_PtFloat, eOPT_PtHalf
  nLeft(0) = x + int_Caption_Width
  nLeft(1) = x + w \ 2 + int_Caption_Width_2
  nLeft(2) = x + w
  If nSelPropSubIndex > 0 And nSelPropIndex = i Then
   GradientFillRect hdc, nLeft(nSelPropSubIndex - 1), y, nLeft(nSelPropSubIndex), y + h, d_Checked1, d_Checked2, GRADIENT_FILL_RECT_V
  End If
 Case eOPT_Rect, eOPT_RectInt, eOPT_RectByte, eOPT_RectFloat, eOPT_RectHalf
  nLeft(0) = x + int_Caption_Width
  nLeft(1) = x + w \ 4 + int_Caption_Width_3
  nLeft(2) = x + w \ 2 + int_Caption_Width_2
  nLeft(3) = x + w - w \ 4 + int_Caption_Width_1
  nLeft(4) = x + w
  If nSelPropSubIndex > 0 And nSelPropIndex = i Then
   GradientFillRect hdc, nLeft(nSelPropSubIndex - 1), y, nLeft(nSelPropSubIndex), y + h, d_Checked1, d_Checked2, GRADIENT_FILL_RECT_V
  End If
 End Select
 'draw value
 If .nType = eOPT_Custom Then
  pp.sValue = "(Custom)" ':-3
 Else
  PropRead ThePrj.Operators(nSelIndex), tDef(nSelType).props(i), pp
 End If
 If .ListCount > 0 Then
  cFnt.DrawTextXP hdc, .List(pp.iValue(0)), x + int_Caption_Width + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
 Else
  Select Case .nType
  Case eOPT_Name, eOPT_String, eOPT_Custom
   cFnt.DrawTextXP hdc, pp.sValue, x + int_Caption_Width + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
  Case eOPT_Byte, eOPT_Integer, eOPT_Long
   cFnt.DrawTextXP hdc, CStr((pp.iValue(0))), x + int_Caption_Width + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
  Case eOPT_Half, eOPT_Single
   cFnt.DrawTextXP hdc, Format((pp.fValue(0)), ".000"), x + int_Caption_Width + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
  Case eOPT_Color
   j = pp.iValue(0)
   j = (j And &HFF0000) \ &H10000 + (j And &HFF00&) + (j And &HFF&) * &H10000
   hbr = CreateSolidBrush(j)
   FillRect hdc, r, hbr
   DeleteObject hbr
   If nSelPropSubIndex > 0 And nSelPropIndex = i Then
    GradientFillRect hdc, nLeft(nSelPropSubIndex - 1), y, nLeft(nSelPropSubIndex), y + h, d_Checked1, d_Checked2, GRADIENT_FILL_RECT_V
   End If
   j = j Xor &H808080
   cFnt.DrawTextXP hdc, CStr(ThePrj.Operators(nSelIndex).bProps(.nOffset + 3)), _
   nLeft(0) + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, j, , True
   cFnt.DrawTextXP hdc, CStr(ThePrj.Operators(nSelIndex).bProps(.nOffset + 2)), _
   nLeft(1) + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, j, , True
   cFnt.DrawTextXP hdc, CStr(ThePrj.Operators(nSelIndex).bProps(.nOffset + 1)), _
   nLeft(2) + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, j, , True
   cFnt.DrawTextXP hdc, CStr(ThePrj.Operators(nSelIndex).bProps(.nOffset)), _
   nLeft(3) + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, j, , True
  Case eOPT_Pt, eOPT_PtInt, eOPT_PtByte
   cFnt.DrawTextXP hdc, CStr((pp.iValue(0))), nLeft(0) + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
   cFnt.DrawTextXP hdc, CStr((pp.iValue(1))), nLeft(1) + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
  Case eOPT_PtFloat, eOPT_PtHalf
   cFnt.DrawTextXP hdc, Format((pp.fValue(0)), ".000"), nLeft(0) + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
   cFnt.DrawTextXP hdc, Format((pp.fValue(1)), ".000"), nLeft(1) + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
  Case eOPT_Rect, eOPT_RectInt, eOPT_RectByte
   cFnt.DrawTextXP hdc, CStr((pp.iValue(0))), nLeft(0) + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
   cFnt.DrawTextXP hdc, CStr((pp.iValue(1))), nLeft(1) + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
   cFnt.DrawTextXP hdc, CStr((pp.iValue(2))), nLeft(2) + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
   cFnt.DrawTextXP hdc, CStr((pp.iValue(3))), nLeft(3) + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
  Case eOPT_RectFloat, eOPT_RectHalf
   cFnt.DrawTextXP hdc, Format((pp.fValue(0)), ".000"), nLeft(0) + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
   cFnt.DrawTextXP hdc, Format((pp.fValue(1)), ".000"), nLeft(1) + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
   cFnt.DrawTextXP hdc, Format((pp.fValue(2)), ".000"), nLeft(2) + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
   cFnt.DrawTextXP hdc, Format((pp.fValue(3)), ".000"), nLeft(3) + 4, y, w, h, DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
  End Select
 End If
 'draw gridline
 Select Case .nType
 Case eOPT_String, eOPT_Custom
  j = 5
 Case eOPT_Group
  j = 0
 Case eOPT_Color
  j = 4
 Case eOPT_Pt, eOPT_PtFloat, eOPT_PtHalf, eOPT_PtInt, eOPT_PtByte
  j = 2
 Case eOPT_Rect, eOPT_RectFloat, eOPT_RectHalf, eOPT_RectInt, eOPT_RectByte
  j = 3
 Case Else
  j = 1
 End Select
 'draw button
 If j > 3 Then
  'TODO:highlight?
  #If Office2003 Then
  GradientFillRect hdc, x + w - 20, y, x + w, y + h, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_V
  #End If
  cFnt.DrawTextXP hdc, "...", x + w - 20, y, 20, h, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
 End If
 'draw gridline
 pPropRedrawGridline hdc, x, y, w, h, j
End With
End Sub

Private Sub pPropRedrawGridline(ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal w As Long, ByVal h As Long, ByVal i As Long)
Dim hbr As Long
Dim r As RECT
#If Office2003 Then
hbr = CreateSolidBrush(d_Border)
#Else
hbr = CreateSolidBrush(&H404040)
#End If
r.Left = x
r.Bottom = y + h
r.Top = r.Bottom - 1
r.Right = x + w
FillRect bmProp.hdc, r, hbr
r.Top = y
r.Bottom = y + h
If i > 0 Then
 r.Left = x + int_Caption_Width
 r.Right = r.Left + 1
 FillRect bmProp.hdc, r, hbr
End If
If i = 4 Or i = 5 Then 'color,string
 w = w - 20&
 r.Left = x + w
 r.Right = r.Left + 1
 FillRect bmProp.hdc, r, hbr
End If
Select Case i
Case 2 'point
 r.Left = x + w \ 2 + int_Caption_Width_2
 r.Right = r.Left + 1
 FillRect bmProp.hdc, r, hbr
Case 3, 4 'rect,color
 r.Left = x + w \ 4 + int_Caption_Width_3
 r.Right = r.Left + 1
 FillRect bmProp.hdc, r, hbr
 r.Left = x + w \ 2 + int_Caption_Width_2
 r.Right = r.Left + 1
 FillRect bmProp.hdc, r, hbr
 r.Left = x + w - w \ 4 + int_Caption_Width_1
 r.Right = r.Left + 1
 FillRect bmProp.hdc, r, hbr
End Select
DeleteObject hbr
End Sub

Private Sub pObjRedraw()
Dim hbr As Long, hRgn As Long
Dim r As RECT
Dim i As Long, m As Long, n As Long
Dim dX As Long, dy As Long
Dim j As Long, nIndex As Long
Dim s As String
Dim p(5) As POINTAPI
r.Right = bmObj.Width
r.Bottom = bmObj.Height
hbr = CreateSolidBrush(TranslateColor(vbButtonFace))
FillRect bmObj.hdc, r, hbr
DeleteObject hbr
If TheIndex > 0 And TheIndex <= ThePrj.nPageCount Then
 hbr = CreateSolidBrush(d_Border)
 dX = sb1.Value(efsHorizontal)
 dy = sb1.Value(efsVertical)
 m = r.Bottom \ 16& + 1&
 n = dy \ 16&
 'draw comments
 With ThePrj.Pages(TheIndex)
  For i = 1 To .nCommentCount
   With .Comments(i)
    If .Top < n + m And .Top + .Height > n Then
     If .Left * 16& < dX + bmObj.Width And (.Left + .Width) * 16& > dX Then
      'draw it
      r.Left = .Left * 16& - dX
      r.Top = .Top * 16& - dy
      r.Right = r.Left + .Width * 16& + 1
      r.Bottom = r.Top + .Height * 16& + 1
      j = CreateSolidBrush(.Color)
      FillRect bmObj.hdc, r, j
      FrameRect bmObj.hdc, r, hbr
      DeleteObject j
      'title
      cFnt.DrawTextXP bmObj.hdc, .Name, r.Left + 4, r.Top, r.Right - r.Left - 8, 16, _
      DT_VCENTER Or DT_NOPREFIX Or DT_SINGLELINE Or DT_END_ELLIPSIS, vbBlack, , True
      'text
      cFnt.DrawTextXP bmObj.hdc, .Value, r.Left + 4, r.Top + 16, r.Right - r.Left - 8, r.Bottom - r.Top - 16, _
      DT_NOPREFIX, vbBlack, , True
     End If
    End If
   End With
  Next i
 End With
 'draw contents
 For i = n To n + m
  r.Top = i * 16& - dy
  r.Bottom = r.Top + 17&
  If i >= int_Page_Height Then Exit For
  If i >= 0 Then
   With ThePrj.Pages(TheIndex).Rows(i)
    For j = 1 To .nOpCount
     nIndex = .idxOp(j)
     With ThePrj.Operators(nIndex)
      r.Left = .Left * 16& - dX
      r.Right = r.Left + .Width * 16&
      If r.Left < bmObj.Width And r.Right >= 0 And .Flags >= 0 Then
       'save or load??
       If .nType = int_OpType_Store Then
        p(0).x = r.Left
        p(0).y = r.Top
        p(1).x = r.Left
        p(1).y = r.Top + 9
        p(2).x = r.Left + 8
        p(2).y = r.Top + 17
        p(3).x = r.Right - 8
        p(3).y = r.Top + 17
        p(4).x = r.Right
        p(4).y = r.Top + 9
        p(5).x = r.Right
        p(5).y = r.Top
        hRgn = CreatePolygonRgn(p(0), 6, ALTERNATE)
        SelectClipRgn bmObj.hdc, hRgn
       ElseIf .nType = int_OpType_Load Then
        p(0).x = r.Left
        p(0).y = r.Top + 17
        p(1).x = r.Left
        p(1).y = r.Top + 8
        p(2).x = r.Left + 8
        p(2).y = r.Top
        p(3).x = r.Right - 8
        p(3).y = r.Top
        p(4).x = r.Right
        p(4).y = r.Top + 8
        p(5).x = r.Right
        p(5).y = r.Top + 17
        hRgn = CreatePolygonRgn(p(0), 6, ALTERNATE)
        SelectClipRgn bmObj.hdc, hRgn
       End If
       'draw background
       If .Flags And int_OpFlags_Selected Then
        GradientFillRect bmObj.hdc, r.Left, r.Top + 1, r.Right, r.Top + 16, d_Checked1, d_Checked2, GRADIENT_FILL_RECT_V
       Else
        GradientFillRect bmObj.hdc, r.Left, r.Top + 1, r.Right, r.Top + 16, d_Bar1, d_Bar2, GRADIENT_FILL_RECT_V
       End If
       'selected?
       If nIndex = nShowIndex Then
        GradientFillRect bmObj.hdc, r.Right - 4, r.Top + 1, r.Right, r.Top + 16, d_Title1, d_Title2, GRADIENT_FILL_RECT_V
       End If
       'draw name
       If .nType = int_OpType_Load Then
        s = """" + .sProps(0) + """"
       Else
        If .Name = "" Then s = tDef(.nType).Name Else s = """" + .Name + """"
       End If
       cFnt.DrawTextXP bmObj.hdc, s, r.Left, r.Top, r.Right - r.Left, 16&, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE Or DT_END_ELLIPSIS, _
       IIf(.Flags And int_OpFlags_Error, vbRed, vbBlack), , True
       'draw border
       If .nType = int_OpType_Store Or .nType = int_OpType_Load Then
        SelectClipRgn bmObj.hdc, 0
        FrameRgn bmObj.hdc, hRgn, hbr, 1, 1
        DeleteObject hRgn
       Else
        FrameRect bmObj.hdc, r, hbr
       End If
      End If
     End With
    Next j
   End With
  End If
  'draw cursor :-3
  If i = nYCur Then
   r.Left = nXCur * 16& - dX
   pDrawCursor bmObj.hdc, r.Left, r.Top
  End If
 Next i
 'draw selected comment
 With ThePrj.Pages(TheIndex)
  i = -nSelIndex
  If i > 0 And i <= .nCommentCount Then
   With .Comments(i)
    r.Left = .Left * 16& - 4 - dX
    r.Top = .Top * 16& - 4 - dy
    TransparentBlt bmObj.hdc, r.Left, r.Top, 9, 9, bm0.hdc, 18, 0, 9, 9, vbGreen
    TransparentBlt bmObj.hdc, r.Left + .Width * 8&, r.Top, 9, 9, bm0.hdc, 18, 0, 9, 9, vbGreen
    TransparentBlt bmObj.hdc, r.Left + .Width * 16&, r.Top, 9, 9, bm0.hdc, 18, 0, 9, 9, vbGreen
    TransparentBlt bmObj.hdc, r.Left, r.Top + .Height * 8&, 9, 9, bm0.hdc, 18, 0, 9, 9, vbGreen
    TransparentBlt bmObj.hdc, r.Left + .Width * 16&, r.Top + .Height * 8&, 9, 9, bm0.hdc, 18, 0, 9, 9, vbGreen
    TransparentBlt bmObj.hdc, r.Left, r.Top + .Height * 16&, 9, 9, bm0.hdc, 18, 0, 9, 9, vbGreen
    TransparentBlt bmObj.hdc, r.Left + .Width * 8&, r.Top + .Height * 16&, 9, 9, bm0.hdc, 18, 0, 9, 9, vbGreen
    TransparentBlt bmObj.hdc, r.Left + .Width * 16&, r.Top + .Height * 16&, 9, 9, bm0.hdc, 18, 0, 9, 9, vbGreen
   End With
  End If
 End With
 'draw selection
 If nCurFlags = 1 Then
  If nXEnd > nXCur Then
   i = nXCur * 16& - dX
   r.Right = (nXEnd - nXCur + 1) * 16&
  Else
   i = nXEnd * 16& - dX
   r.Right = (nXCur - nXEnd + 1) * 16&
  End If
  If nYEnd > nYCur Then
   j = nYCur * 16& - dy
   r.Bottom = (nYEnd - nYCur + 1) * 16&
  Else
   j = nYEnd * 16& - dy
   r.Bottom = (nYCur - nYEnd + 1) * 16&
  End If
  bmObj.PaintPicture bmObj.hdc, i, j, r.Right, r.Bottom, 0, 0, vbDstInvert
 End If
 'over
 DeleteObject hbr
Else
 cFnt.DrawTextXP bmObj.hdc, "Please select a page", 0, 0, r.Right, r.Bottom, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
End If
picObj_Paint
'draw preview
If frmBird.IsVisible Then frmBird.Refresh
End Sub

Private Sub pViewRedraw()
Dim hbr As Long
Dim r As RECT
Dim i As Long, j As Long, k As Long
r.Right = bmView.Width
r.Bottom = bmView.Height
hbr = CreateSolidBrush(TranslateColor(vbButtonFace))
FillRect bmView.hdc, r, hbr
DeleteObject hbr
With ThePrj
 If nShowIndex > 0 And nShowIndex <= .nOpCount Then
  With .Operators(nShowIndex)
   If .Flags >= 0 Then
    If .Flags And int_OpFlags_Error Then
     cFnt.DrawTextXP bmView.hdc, "Unknown Type", 0, 0, r.Right, r.Bottom, DT_CENTER Or DT_VCENTER Or DT_SINGLELINE, vbBlack, , True
    Else
     If (.Flags And int_OpFlags_InMemory) = 0 Then
      '//////////////////////////////////////////////redraw operator!!!
      Dim t As Double
      Me.MousePointer = vbHourglass
      nOldTime = GetTickCount
      t = cObj.ShowOperator(ThePrj, nShowIndex, Me)
      With stb1
       If nShowIndex > 0 Then
        .PanelCaption(1) = "Calc " + Format(t, "0.00") + "ms"
       Else
        .PanelCaption(1) = "Ready"
       End If
       .PanelStyle(1) = sbNormal
      End With
      Me.MousePointer = vbDefault
      '//////////////////////////////////////////////redraw!!!
     End If
     If nShowIndex > 0 Then
      hbr = .nBmIndex
      If hbr < 0 Then hbr = -hbr
      Debug.Assert hbr > 0 And hbr <= cObj.BitmapCount
      With cObj.TheBitmap(hbr)
       If bViewTile Then
        i = nViewX And (.Width - 1)
        j = nViewY And (.Height - 1)
        .AlphaPaintPicture bmView.hdc, i, j, , , , , , True
        If i > 0 Then
         .AlphaPaintPicture bmView.hdc, i - .Width, j, , , , , , True
         If j > 0 Then
         .AlphaPaintPicture bmView.hdc, i - .Width, j - .Height, , , , , , True
         End If
        End If
        If j > 0 Then
         .AlphaPaintPicture bmView.hdc, i, j - .Height, , , , , , True
        End If
        If i + .Width < r.Right Then 'tile x
         For k = .Width To r.Right - 1 Step .Width
          bmView.PaintPicture bmView.hdc, k, 0, .Width, j + .Height
         Next k
        End If
        If j + .Height < r.Bottom Then 'tile y
         For k = .Height To r.Bottom - 1 Step .Height
          bmView.PaintPicture bmView.hdc, 0, k, r.Right, .Height
         Next k
        End If
       Else
        .AlphaPaintPicture bmView.hdc, nViewX, nViewY, , , , , , True
       End If
      End With
     End If
    End If
   End If
  End With
 End If
End With
picView_Paint
End Sub

Private Sub picView_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
 With ThePrj
  If nShowIndex > 0 And nShowIndex <= .nOpCount Then
   With .Operators(nShowIndex)
    If .Flags >= 0 Then
     If (.Flags And int_OpFlags_Error) = 0 Then
      Debug.Assert .Flags And int_OpFlags_InMemory
      Debug.Assert .nBmWidth > 0 And .nBmHeight > 0
      picView.MousePointer = vbSizePointer
      nViewXCur = nViewX - x
      nViewYCur = nViewY - y
     End If
    End If
   End With
  End If
 End With
ElseIf Button = 2 Then
 fm1.PopupMenu "_view"
End If
End Sub

Private Sub picView_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 And nViewXCur > &H80000000 Then
 With ThePrj
  If nShowIndex > 0 And nShowIndex <= .nOpCount Then
   With .Operators(nShowIndex)
    If .Flags >= 0 Then
     If (.Flags And int_OpFlags_Error) = 0 Then
      Debug.Assert .Flags And int_OpFlags_InMemory
      Debug.Assert .nBmWidth > 0 And .nBmHeight > 0
      nViewX = x + nViewXCur
      nViewY = y + nViewYCur
      If bViewTile Then
       nViewX = nViewX And (.nBmWidth - 1)
       nViewY = nViewY And (.nBmHeight - 1)
      End If
      pViewRedraw
     End If
    End If
   End With
  End If
 End With
End If
End Sub

Private Sub picView_MouseUp(Button As Integer, Shift As Integer, x As Single, y As Single)
If Button = 1 Then
 nViewXCur = &H80000000
 picView.MousePointer = vbDefault
End If
End Sub

Private Sub picView_Paint()
bmView.PaintPicture picView.hdc
End Sub

Private Sub picView_Resize()
On Error Resume Next
bmView.Create picView.Width, picView.Height
pViewRedraw
End Sub

Private Sub sb1_Change(eBar As EFSScrollBarConstants)
pObjRedraw
End Sub

Private Sub sb1_MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long)
pObjRedraw
End Sub

Private Sub sb1_Scroll(eBar As EFSScrollBarConstants)
pObjRedraw
End Sub

Private Sub sb2_Change(eBar As EFSScrollBarConstants)
pPropHide
pPropRedraw
End Sub

Private Sub sb2_MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long)
pPropHide
pPropRedraw
End Sub

Private Sub sb2_Scroll(eBar As EFSScrollBarConstants)
pPropHide
pPropRedraw
End Sub

Private Sub pDrawCursor(ByVal hdc As Long, ByVal x As Long, ByVal y As Long) 'stupid function
Dim p(3) As POINTAPI
Dim hpn As Long
p(0).x = x
p(0).y = y
p(1).x = x + 15
p(1).y = y + 8
p(2).x = x
p(2).y = y + 16
p(3) = p(0)
hpn = CreatePen(0, 1, vbBlack)
hpn = SelectObject(hdc, hpn)
Polyline hdc, p(0), 4
DeleteObject SelectObject(hdc, hpn)
End Sub

Private Sub sbList_Change(eBar As EFSScrollBarConstants)
pListRedraw
End Sub

Private Sub sbList_MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long)
pListRedraw
End Sub

Private Sub sbList_Scroll(eBar As EFSScrollBarConstants)
pListRedraw
End Sub

Private Sub stb1_GetInfo(ByVal PanelIndex As Long, s As String)
Dim t As PROCESS_MEMORY_COUNTERS
Select Case stb1.PanelCount - PanelIndex
Case 3
 t.cb = Len(t)
 GetProcessMemoryInfo GetCurrentProcess, t, t.cb
 s = "Memory:" + Format((t.PagefileUsage + t.WorkingSetSize) / 1048576#, "0.00") + "MB"
Case 2
 s = "CPU:" + Format(cCPU1.GetCPUUsage, "0.0%")
End Select
End Sub

Private Sub tb1_Click(ByVal btnIndex As Long, ByVal btnKey As String)
Select Case btnIndex
Case 1 'new
 mnuNew_Click
Case 2 'open
 mnuOpen_Click
Case 3 'save
 mnuSave_Click
Case 5 'cut
 mnuEdit1_Click 0
Case 6 'copy
 mnuEdit1_Click 1
Case 7 'paste
 mnuEdit1_Click 2
Case 9 'delete
 mnuEdit1_Click 3
Case 11 'save as bitmap
 mnuSaveBmp_Click
Case 12 'save all
Case 13 'compile
Case 15 'release :-3
 If nShowIndex > 0 Then
  nShowIndex = 0
  pObjRedraw
  pViewRedraw
 End If
 cObj.ClearMemory ThePrj
End Select
End Sub

Private Sub tb2_Click(ByVal btnIndex As Long, ByVal btnKey As String)
Dim i As Long, j As Long, k As Long
Select Case btnIndex
Case 1 'add
 fm1.Click "[P2]", "addp"
Case 2 'rename
 fm1.Click "[P2]", "renp"
Case 3 'del
 mnuDelP_Click
Case 5 'up
 mnuPMoveUp_Click
Case 6 'down
 mnuPMoveDown_Click
Case 8 'left
 fm1.Click "[P2]", "p2_l"
Case 9 'right
 fm1.Click "[P2]", "p2_r"
Case 11 'collapse
 If TheIndex <= 0 Or TheIndex > ThePrj.nPageCount Then Exit Sub
 If TheIndex < ThePrj.nPageCount Then
  If ThePageState(TheIndex) = 0 And _
  ThePrj.Pages(TheIndex).nIndent < ThePrj.Pages(TheIndex + 1).nIndent Then pListCollapse
 End If
Case 12 'expand
 If TheIndex <= 0 Or TheIndex >= ThePrj.nPageCount Then Exit Sub
 If ThePageState(TheIndex) = 1 And _
 ThePrj.Pages(TheIndex).nIndent < ThePrj.Pages(TheIndex + 1).nIndent Then pListExpand
Case 13 'collapse all
 VisiblePageCount = 0
 j = &H7FFFFFFF
 For i = 1 To ThePrj.nPageCount - 1
  k = ThePrj.Pages(i).nIndent
  If k <= j Then
   j = k
   If ThePrj.Pages(i + 1).nIndent > k Then ThePageState(i) = 1 Else ThePageState(i) = 0
   VisiblePageCount = VisiblePageCount + 1
  Else
   If ThePrj.Pages(i + 1).nIndent > k Then ThePageState(i) = 3 Else ThePageState(i) = 2
  End If
 Next i
 If ThePrj.nPageCount > 0 Then
  If ThePrj.Pages(i).nIndent <= j Then
   ThePageState(i) = 0
   VisiblePageCount = VisiblePageCount + 1
  Else
   ThePageState(i) = 2
  End If
 End If
 If TheIndex > 0 And TheIndex <= ThePrj.nPageCount Then
  If ThePageState(TheIndex) >= 2 Then
   TheIndex = 0
   lstPage_Click
  End If
 End If
 pListRefresh
Case 14 'expand all
 VisiblePageCount = ThePrj.nPageCount
 ReDim ThePageState(1 To ThePrj.nPageCount) ':-3
 pListRefresh
End Select
End Sub

Private Sub txtProp_Change()
pChangeValue txtProp.Text
End Sub

Private Sub txtProp_KeyDown(KeyCode As Integer, Shift As Integer)
Select Case KeyCode
Case vbKeyEscape
 txtProp.Text = txtProp.Tag
 txtProp.Visible = False
Case vbKeyReturn
 txtProp.Visible = False
End Select
End Sub

Private Sub txtProp_LostFocus()
txtProp.Visible = False
End Sub

Private Sub pChangeValue(ByVal sValue As String)
On Error Resume Next
Dim nChange As Long
'&H0 -don't change
'&H1 -redraw properties
'&H2 -redraw operator
'&H4 -redraw bitmap
'&H8 -must validate operator
Dim pp As typeOperatorProp_DesignTime
Dim i As Long, f As Single
If nSelType > 0 And nSelIndex > 0 And nSelPropIndex > 0 Then
 'get value
 PropRead ThePrj.Operators(nSelIndex), tDef(nSelType).props(nSelPropIndex), pp
 'check type
 With tDef(nSelType).props(nSelPropIndex)
  Select Case .nType
  Case eOPT_Name, eOPT_String
   If sValue <> pp.sValue Then
    If nSelType = int_OpType_Load Then
     If .nType = eOPT_String Then nChange = &HF& Else nChange = &H1&
    Else
     If .nType = eOPT_String Then nChange = &H5& Else nChange = &H3&
     If nSelType = int_OpType_Store Then
      nChange = &H7&
      'rename!
      ThePrj.Operators(nSelIndex).Name = sValue
      'validate all
      cObj.ValidateAllLoadOps ThePrj, pp.sValue
      cObj.ValidateAllLoadOps ThePrj, sValue
     End If
    End If
    pp.sValue = sValue
   End If
  Case eOPT_Byte, eOPT_Integer, eOPT_Long, eOPT_Size, eOPT_Bool, eOPT_ChangeSize
   i = Val(sValue)
   If i <> pp.iValue(0) Then
    pp.iValue(0) = i
    If .nType = eOPT_Size Or .nType = eOPT_ChangeSize Then nChange = &HF& Else nChange = &H5&
   End If
  Case eOPT_Half, eOPT_Single
   f = Val(sValue)
   If f <> pp.fValue(0) Then
    pp.fValue(0) = f
    nChange = &H5&
   End If
  Case eOPT_Color
   If nSelPropSubIndex > 0 And nSelPropSubIndex <= 4 Then
    i = Val(sValue)
    If i < 0 Then i = 0
    If i > 255 Then i = 255
    If i <> ThePrj.Operators(nSelIndex).bProps(.nOffset + 4 - nSelPropSubIndex) Then
     ThePrj.Operators(nSelIndex).bProps(.nOffset + 4 - nSelPropSubIndex) = i
     CopyMemory pp.iValue(0), ThePrj.Operators(nSelIndex).bProps(.nOffset), 4&
     nChange = &H5&
    End If
   End If
  Case eOPT_Pt, eOPT_PtInt, eOPT_PtByte, eOPT_Rect, eOPT_RectInt, eOPT_RectByte
   If nSelPropSubIndex > 0 And nSelPropSubIndex <= 4 Then
    i = Val(sValue)
    If i <> pp.iValue(nSelPropSubIndex - 1) Then
     pp.iValue(nSelPropSubIndex - 1) = i
     nChange = &H5&
    End If
   End If
  Case eOPT_PtFloat, eOPT_PtHalf, eOPT_RectFloat, eOPT_RectHalf
   If nSelPropSubIndex > 0 And nSelPropSubIndex <= 4 Then
    f = Val(sValue)
    If f <> pp.fValue(nSelPropSubIndex - 1) Then
     pp.fValue(nSelPropSubIndex - 1) = f
     nChange = &H5&
    End If
   End If
  End Select
 End With
 'changed?
 If nChange <> 0 Then
  PropWrite ThePrj.Operators(nSelIndex), tDef(nSelType).props(nSelPropIndex), pp
  If nChange And &H1& Then
   pPropRedraw
  End If
  If nChange And &H8& Then
   'validate operator
   cObj.ValidateOps ThePrj, nSelIndex
  End If
  If nChange And &H4& Then
   'calc bitmap
   cObj.SetNotInMemoryFlags ThePrj, nSelIndex
   pViewRedraw
  End If
  If nChange And &H2& Then
   pObjRedraw
  End If
 End If
End If
End Sub

'fix the unknown bug
Private Sub pClearLockCount()
Dim i As Long
'get address of SafeArray
CopyMemory i, ByVal VarPtr(ThePrj) + 12&, 4&
If i <> 0 Then
 'get address of cLocks
 i = i + 8&
 'set to zero (!!!)
 CopyMemory ByVal i, 0&, 4&
End If
End Sub

'show position
Private Sub pShowPos()
Dim s As String
Select Case nCurFlags
Case 1 'selection
 s = "(" + CStr(nXCur) + "," + CStr(nYCur) + ")-(" + CStr(nXEnd) + "," + CStr(nYEnd) + ")"
Case Else
 s = "(" + CStr(nXCur) + "," + CStr(nYCur) + ")"
End Select
stb1.PanelCaption(4) = s
End Sub

Private Sub pListRedraw()
Dim i As Long, y As Long
Dim r As RECT, hbr As Long
'erase background
hbr = CreateSolidBrush(TranslateColor(vbButtonFace))
r.Right = bmList.Width
r.Bottom = bmList.Height
FillRect bmList.hdc, r, hbr
DeleteObject hbr
'find top :-3
y = -sbList.Value(efsVertical)
For i = 1 To ThePrj.nPageCount
 If y > -16 Then Exit For
 If ThePageState(i) < 2 Then y = y + 16
Next i
'start draw
hbr = CreateSolidBrush(&H800000)
r.Right = picList.ScaleWidth
For i = i To ThePrj.nPageCount
 If ThePageState(i) < 2 Then
  r.Top = y
  r.Bottom = y + 16
  If i = TheIndex Then
   r.Left = 0
   GradientFillRect bmList.hdc, r.Left, r.Top, r.Right, r.Bottom, d_Checked1, d_Checked2, GRADIENT_FILL_RECT_V
   FrameRect bmList.hdc, r, hbr
  End If
  r.Left = ThePrj.Pages(i).nIndent * 8&
  If i < ThePrj.nPageCount Then
   If ThePrj.Pages(i + 1).nIndent * 8& > r.Left Then
    If ThePageState(i) = 0 Then
     TransparentBlt bmList.hdc, r.Left + 4, r.Top + 4, 9, 9, bm0.hdc, 9, 0, 9, 9, vbGreen
    Else
     TransparentBlt bmList.hdc, r.Left + 4, r.Top + 4, 9, 9, bm0.hdc, 0, 0, 9, 9, vbGreen
    End If
   End If
  End If
  r.Left = r.Left + 16
  cFnt.DrawTextXP bmList.hdc, ThePrj.Pages(i).Name, r.Left, r.Top, r.Right - r.Left, 16, _
  DT_VCENTER Or DT_SINGLELINE Or DT_NOPREFIX Or DT_EXPANDTABS Or DT_END_ELLIPSIS, vbBlack, , True
  'get next
  y = r.Bottom
  If y >= bmList.Height Then Exit For
 End If
Next i
DeleteObject hbr
picList_Paint
End Sub

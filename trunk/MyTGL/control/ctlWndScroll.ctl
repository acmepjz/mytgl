VERSION 5.00
Begin VB.UserControl ctlWndScroll 
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
      Caption         =   "vbAccelerator Scrollbar Control"
      ForeColor       =   &H00FFFFFF&
      Height          =   255
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   2535
   End
End
Attribute VB_Name = "ctlWndScroll"
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

' ===========================================================================
' Name:     cScrollBars
' Author:   Steve McMahon (steve@vbaccelerator.com)
' Date:     24 December 1998
' 'Requires: SSUBTMR.DLL
'
' ---------------------------------------------------------------------------
' Copyright  1998 Steve McMahon (steve@vbaccelerator.com)
' Visit vbAccelerator - free, advanced source code for VB programmers.
'     http://vbaccelerator.com
' ---------------------------------------------------------------------------
'
' Description:
' A class which can add scroll bars to VB Forms, Picture Boxes and
' UserControls.
' Features:
'  * True API scroll bars, which don't flash or draw badly like
'    the VB ones
'  * Scroll bar values are long integers, i.e. >2 billion values
'  * Set Flat or Encarta scroll bar modes if your COMCTL32.DLL version
'    supports it (>4.72)
'
' Updates:
' 2003-07-02
'  * Added Mouse Wheel Support.  Thanks to Chris Eastwood for
'    the suggestion and starter code.
'    Visit his site at http://vbcodelibrary.co.uk/
'  * Scroll bar now goes to bottom when SB_BOTTOM fired
'    (e.g. right click on scroll bar with mouse)
'  * New ScrollClick events to enable focus
'  * Removed a large quantity of redundant declares which
'    had found their way into this class somehow...
' ===========================================================================


' ---------------------------------------------------------------------
' vbAccelerator Software License
' Version 1.0
' Copyright (c) 2002 vbAccelerator.com
'
' Redistribution and use in source and binary forms, with or
' without modification, are permitted provided that the following
' conditions are met:
'
' 1. Redistributions of source code must retain the above copyright
'    notice, this list of conditions and the following disclaimer.
'
' 2. Redistributions in binary form must reproduce the above copyright
'    notice, this list of conditions and the following disclaimer in
'    the documentation and/or other materials provided with the distribution.
'
' 3. The end-user documentation included with the redistribution, if any,
'    must include the following acknowledgment:
'
'  "This product includes software developed by vbAccelerator (http://vbaccelerator.com/)."
'
' Alternately, this acknowledgment may appear in the software itself, if
' and wherever such third-party acknowledgments normally appear.
'
' 4. The name "vbAccelerator" must not be used to endorse or promote products
'    derived from this software without prior written permission. For written
'    permission, please contact vbAccelerator through steve@vbaccelerator.com.
'
' 5. Products derived from this software may not be called "vbAccelerator",
'    nor may "vbAccelerator" appear in their name, without prior written
'    permission of vbAccelerator.
'
' THIS SOFTWARE IS PROVIDED "AS IS" AND ANY EXPRESSED OR IMPLIED WARRANTIES,
' INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES OF MERCHANTABILITY
' AND FITNESS FOR A PARTICULAR PURPOSE ARE DISCLAIMED. IN NO EVENT SHALL
' VBACCELERATOR OR ITS CONTRIBUTORS BE LIABLE FOR ANY DIRECT, INDIRECT,
' INCIDENTAL, SPECIAL, EXEMPLARY, OR CONSEQUENTIAL DAMAGES (INCLUDING,
' BUT NOT LIMITED TO, PROCUREMENT OF SUBSTITUTE GOODS OR SERVICES; LOSS OF
' USE, DATA, OR PROFITS; OR BUSINESS INTERRUPTION) HOWEVER CAUSED AND ON ANY
' THEORY OF LIABILITY, WHETHER IN CONTRACT, STRICT LIABILITY, OR TORT
' (INCLUDING NEGLIGENCE OR OTHERWISE) ARISING IN ANY WAY OUT OF THE USE OF
' THIS SOFTWARE, EVEN IF ADVISED OF THE POSSIBILITY OF SUCH DAMAGE.
'
' ---------------------------------------------------------------------

Private Type OSVERSIONINFO
    dwOSVersionInfoSize As Long
    dwMajorVersion As Long
    dwMinorVersion As Long
    dwBuildNumber As Long
    dwPlatformId As Long
    szCSDVersion As String * 128      '  Maintenance string for PSS usage
End Type

Private Declare Function GetVersionEx Lib "kernel32" Alias "GetVersionExA" (lpVersionInformation As OSVERSIONINFO) As Long
Private Declare Function RegisterWindowMessage Lib "user32" Alias "RegisterWindowMessageA" (ByVal lpString As String) As Long
Private Declare Function SystemParametersInfo Lib "user32" Alias "SystemParametersInfoA" (ByVal uAction As Long, ByVal uParam As Long, ByRef lpvParam As Long, ByVal fuWinIni As Long) As Long

'private declare function InitializeFlatSB(HWND) as long
Private Declare Function InitialiseFlatSB Lib "comctl32.dll" Alias "InitializeFlatSB" (ByVal lhWnd As Long) As Long

' Scroll bar:
Private Type SCROLLINFO
    cbSize As Long
    fMask As Long
    nMin As Long
    nMax As Long
    nPage As Long
    nPos As Long
    nTrackPos As Long
End Type
Private Declare Function SetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, lpcScrollInfo As SCROLLINFO, ByVal BOOL As Boolean) As Long
Private Declare Function GetScrollInfo Lib "user32" (ByVal hwnd As Long, ByVal n As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function GetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long) As Long
Private Declare Function GetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, lpMinPos As Long, lpMaxPos As Long) As Long
Private Declare Function SetScrollPos Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nPos As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetScrollRange Lib "user32" (ByVal hwnd As Long, ByVal nBar As Long, ByVal nMinPos As Long, ByVal nMaxPos As Long, ByVal bRedraw As Long) As Long
    Private Const SB_BOTH = 3
    Private Const SB_BOTTOM = 7
    Private Const SB_CTL = 2
    Private Const SB_ENDSCROLL = 8
    Private Const SB_HORZ = 0
    Private Const SB_LEFT = 6
    Private Const SB_LINEDOWN = 1
    Private Const SB_LINELEFT = 0
    Private Const SB_LINERIGHT = 1
    Private Const SB_LINEUP = 0
    Private Const SB_PAGEDOWN = 3
    Private Const SB_PAGELEFT = 2
    Private Const SB_PAGERIGHT = 3
    Private Const SB_PAGEUP = 2
    Private Const SB_RIGHT = 7
    Private Const SB_THUMBPOSITION = 4
    Private Const SB_THUMBTRACK = 5
    Private Const SB_TOP = 6
    Private Const SB_VERT = 1

    Private Const SIF_RANGE = &H1
    Private Const SIF_PAGE = &H2
    Private Const SIF_POS = &H4
    Private Const SIF_DISABLENOSCROLL = &H8
    Private Const SIF_TRACKPOS = &H10
    Private Const SIF_ALL = (SIF_RANGE Or SIF_PAGE Or SIF_POS Or SIF_TRACKPOS)

   Private Const ESB_DISABLE_BOTH = &H3
   Private Const ESB_ENABLE_BOTH = &H0
   
   Private Const SBS_SIZEGRIP = &H10&
   
Private Declare Function EnableScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wSBflags As Long, ByVal wArrows As Long) As Long
Private Declare Function ShowScrollBar Lib "user32" (ByVal hwnd As Long, ByVal wBar As Long, ByVal bShow As Long) As Long

' Non-client messages:
Private Const WM_NCLBUTTONDOWN = &HA1
Private Const WM_NCRBUTTONDOWN = &HA4
'Private Const WM_NCMBUTTONDOWN = &HA7

' Hit test codes for scroll bars:
Private Const HTHSCROLL = 6
Private Const HTVSCROLL = 7

' Scroll bar messages:
Private Const WM_VSCROLL = &H115
Private Const WM_HSCROLL = &H114
Private Const WM_MOUSEWHEEL = &H20A

' Mouse wheel stuff:
Private Const WHEEL_DELTA = 120
Private Const WHEEL_PAGESCROLL = -1
Private Const SPI_GETWHEELSCROLLLINES = &H68

' Old school Wheel Mouse is not supported in this class.
' NT3.51 or Win95 only
'// Class name for MSWHEEL.EXE's invisible window
'// use FindWindow to get hwnd to MSWHEEL
Private Const MSH_MOUSEWHEEL = "MSWHEEL_ROLLMSG"
Private Const MSH_WHEELMODULE_CLASS = "MouseZ"
Private Const MSH_WHEELMODULE_TITLE = "Magellan MSWHEEL"
'// Apps need to call RegisterWindowMessage using the #defines
'// below to get the message numbers for:
'// 1) the message that can be sent to the MSWHEEL window to
'//    query if wheel support is active (MSH_WHEELSUPPORT)>
'// 2) the message to query for the number of scroll lines
'//    (MSH_SCROLL_LINES)
'//
'// To send a message to MSWheel window, use FindWindow with the #defines
'// for CLASS and TITLE above.  If FindWindow fails to find the MSWHEEL
'// window or the return from SendMessage is false, then Wheel support
'// is not currently available.
Private Const MSH_WHEELSUPPORT = "MSH_WHEELSUPPORT_MSG"
Private Const MSH_SCROLL_LINES = "MSH_SCROLL_LINES_MSG"

' Flat scroll bars:
Private Const WSB_PROP_CYVSCROLL = &H1&
Private Const WSB_PROP_CXHSCROLL = &H2&
Private Const WSB_PROP_CYHSCROLL = &H4&
Private Const WSB_PROP_CXVSCROLL = &H8&
Private Const WSB_PROP_CXHTHUMB = &H10&
Private Const WSB_PROP_CYVTHUMB = &H20&
Private Const WSB_PROP_VBKGCOLOR = &H40&
Private Const WSB_PROP_HBKGCOLOR = &H80&
Private Const WSB_PROP_VSTYLE = &H100&
Private Const WSB_PROP_HSTYLE = &H200&
Private Const WSB_PROP_WINSTYLE = &H400&
Private Const WSB_PROP_PALETTE = &H800&
Private Const WSB_PROP_MASK = &HFFF&

Private Const FSB_FLAT_MODE = 2&
Private Const FSB_ENCARTA_MODE = 1&
Private Const FSB_REGULAR_MODE = 0&

Private Declare Function FlatSB_EnableScrollBar Lib "comctl32.dll" (ByVal hwnd As Long, ByVal int2 As Long, ByVal UINT3 As Long) As Long
Private Declare Function FlatSB_ShowScrollBar Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, ByVal fRedraw As Boolean) As Long

Private Declare Function FlatSB_GetScrollRange Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, ByVal LPINT1 As Long, ByVal LPINT2 As Long) As Long
Private Declare Function FlatSB_GetScrollInfo Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO) As Long
Private Declare Function FlatSB_GetScrollPos Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long) As Long
Private Declare Function FlatSB_GetScrollProp Lib "comctl32.dll" (ByVal hwnd As Long, ByVal propIndex As Long, ByVal LPINT As Long) As Long

Private Declare Function FlatSB_SetScrollPos Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, ByVal pos As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollInfo Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, LPSCROLLINFO As SCROLLINFO, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollRange Lib "comctl32.dll" (ByVal hwnd As Long, ByVal code As Long, ByVal Min As Long, ByVal Max As Long, ByVal fRedraw As Boolean) As Long
Private Declare Function FlatSB_SetScrollProp Lib "comctl32.dll" (ByVal hwnd As Long, ByVal Index As Long, ByVal newValue As Long, ByVal fRedraw As Boolean) As Long

Private Declare Function InitializeFlatSB Lib "comctl32.dll" (ByVal hwnd As Long) As Long
Private Declare Function UninitializeFlatSB Lib "comctl32.dll" (ByVal hwnd As Long) As Long

' Message response:
Implements iSubclass

' Initialisation state:
Private m_bInitialised As Boolean

' Orientation
Public Enum EFSOrientationConstants
    efsoHorizontal
    efsoVertical
    efsoBoth
End Enum
Private m_eOrientation As EFSOrientationConstants

' Style
Public Enum EFSStyleConstants
    efsRegular = FSB_REGULAR_MODE
    efsEncarta = FSB_ENCARTA_MODE
    efsFlat = FSB_FLAT_MODE
End Enum
Private m_eStyle As EFSStyleConstants
' Bars:
Public Enum EFSScrollBarConstants
   efsHorizontal = SB_HORZ
   efsVertical = SB_VERT
End Enum

' Can we have flat scroll bars?
Private m_bNoFlatScrollBars As Boolean

' hWnd we're adding scroll bars too:
Private m_hWnd As Long

' Small change amount
Private m_lSmallChangeHorz As Long
Private m_lSmallChangeVert As Long
' Enabled:
Private m_bEnabledHorz As Boolean
Private m_bEnabledVert As Boolean
' Visible
Private m_bVisibleHorz As Boolean
Private m_bVisibleVert As Boolean

' Number of lines to scroll for each wheel click:
Private m_lWheelScrollLines  As Long

Public Event ScrollClick(eBar As EFSScrollBarConstants, eButton As MouseButtonConstants)
Public Event Scroll(eBar As EFSScrollBarConstants)
Public Event Change(eBar As EFSScrollBarConstants)
Public Event MouseWheel(eBar As EFSScrollBarConstants, lAmount As Long)

'///////////////Add!!
Private cSub As New cSubclass
Private mode2 As Long, bdrWidth As Long, clr1 As Long, clr2 As Long

Private Declare Function GetDCEx Lib "user32.dll" (ByVal hwnd As Long, ByVal hrgnclip As Long, ByVal fdwOptions As Long) As Long
Private Declare Function StretchBlt Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, ByVal dwRop As Long) As Long
Private Declare Function ReleaseDC Lib "user32.dll" (ByVal hwnd As Long, ByVal hdc As Long) As Long
Private Const DCX_WINDOW As Long = &H1&
Private Const DCX_INTERSECTRGN As Long = &H80&
Private Declare Function GetWindowRect Lib "user32.dll" (ByVal hwnd As Long, ByRef lpRect As RECT) As Long
Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type
Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function FillRect Lib "user32.dll" (ByVal hdc As Long, ByRef lpRect As RECT, ByVal hBrush As Long) As Long
Private Const WM_NCPAINT As Long = &H85
Private Const WM_SETCURSOR As Long = &H20
Private Const WM_TIMER As Long = &H113
Private Const WM_LBUTTONDOWN As Long = &H201
Private Const WM_MBUTTONDOWN As Long = &H207
Private Const WM_RBUTTONDOWN As Long = &H204
Private Const WM_LBUTTONUP As Long = &H202
Private Const WM_MBUTTONUP As Long = &H208
Private Const WM_RBUTTONUP As Long = &H205
Private Const WM_MOUSEMOVE As Long = &H200
Private Const WM_PAINT As Long = &HF&

Private Declare Function BeginPaint Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long
Private Declare Function EndPaint Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPaint As PAINTSTRUCT) As Long

Private Type PAINTSTRUCT
    hdc As Long
    fErase As Long
    rcPaint As RECT
    fRestore As Long
    fIncUpdate As Long
    rgbReserved As Byte
End Type

Private Declare Function SetCapture Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function ReleaseCapture Lib "user32.dll" () As Long

Private Declare Function CreateCompatibleDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function DeleteDC Lib "gdi32.dll" (ByVal hdc As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function BitBlt Lib "gdi32.dll" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
Private Declare Function LoadImage Lib "user32.dll" Alias "LoadImageA" (ByVal hInst As Long, ByVal lpsz As Long, ByVal un1 As Long, ByVal n1 As Long, ByVal n2 As Long, ByVal un2 As Long) As Long
Private Declare Function LoadCursor Lib "user32.dll" Alias "LoadCursorA" (ByVal hInstance As Long, ByVal lpCursorName As Long) As Long
Private Declare Function SetCursor Lib "user32.dll" (ByVal hCursor As Long) As Long
Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function GetPixel Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function Ellipse Lib "gdi32.dll" (ByVal hdc As Long, ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Const NULL_BRUSH As Long = 5

Private Declare Function CreateWindowEx Lib "user32.dll" Alias "CreateWindowExA" (ByVal dwExStyle As Long, ByVal lpClassName As String, ByVal lpWindowName As String, ByVal dwStyle As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hWndParent As Long, ByVal hMenu As Long, ByVal hInstance As Long, ByRef lpParam As Any) As Long
Private Declare Function ShowWindow Lib "user32.dll" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long
Private Declare Function SetParent Lib "user32.dll" (ByVal hWndChild As Long, ByVal hWndNewParent As Long) As Long
Private Declare Function DestroyWindow Lib "user32.dll" (ByVal hwnd As Long) As Long
Private Declare Function CreateEllipticRgn Lib "gdi32.dll" (ByVal X1 As Long, ByVal Y1 As Long, ByVal X2 As Long, ByVal Y2 As Long) As Long
Private Declare Function SetWindowRgn Lib "user32.dll" (ByVal hwnd As Long, ByVal hRgn As Long, ByVal bRedraw As Long) As Long
Private Declare Function SetTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long, ByVal uElapse As Long, ByVal lpTimerFunc As Long) As Long
Private Declare Function KillTimer Lib "user32.dll" (ByVal hwnd As Long, ByVal nIDEvent As Long) As Long

Private Const WS_VISIBLE As Long = &H10000000
Private Const WS_CHILD As Long = &H40000000
Private Const WS_EX_TOPMOST As Long = &H8&
Private Const WS_EX_TOOLWINDOW As Long = &H80&
Private Const IMAGE_BITMAP As Long = 0
Private Const LR_SHARED As Long = &H8000

'user32.dll
'bitmap
'32559 32660 32661
' |           |
' *    -*-   -*-
' |           |
Private m_hMemDC As Long, m_hbm1 As Long, m_hbm2 As Long, m_hbm3 As Long
Private m_hbmOld As Long, m_nclrBmp As Long
'cursor
'32652 32653 32654 32655 32656 32657 32658 32659 32660 32661 32662
' |           |     |                      \       /
' *    -*-   -*-    *     *    -*     *-    *     *     *     *
' |           |           |                            /       \

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Type POINTAPI
    x As Long
    y As Long
End Type

Private m_bInMidButton As Boolean
Private m_tMidBtnPos As POINTAPI, m_tMidBtnPos_Old As POINTAPI
Private m_hWndMidBtn As Long, m_hRgnWnd As Long
Private cSub2 As New cSubclass

'Public Enum enumNCPaintMode
' enumNCPaintDefault = 0
' enumNCPaintSolid = 1
' enumNCPaintHorizontal = 2
' enumNCPaintVertical = 3
' enumNCPaintCustom = 99
'End Enum

Public Event Paint(ByVal hdc As Long, ByVal w As Long, ByVal h As Long)

Public Property Get NCBorderWidth() As Long
NCBorderWidth = bdrWidth
End Property

Public Property Let NCBorderWidth(ByVal n As Long)
bdrWidth = n
End Property

Public Property Get NCPaintMode() As enumNCPaintMode
NCPaintMode = mode2
End Property

Public Property Let NCPaintMode(ByVal n As enumNCPaintMode)
mode2 = n
End Property

Public Property Get NCPaintColor1() As OLE_COLOR
NCPaintColor1 = clr1
End Property

Public Property Get NCPaintColor2() As OLE_COLOR
NCPaintColor2 = clr2
End Property

Public Property Let NCPaintColor1(ByVal clr As OLE_COLOR)
clr1 = clr
End Property

Public Property Let NCPaintColor2(ByVal clr As OLE_COLOR)
clr2 = clr
End Property

'///////////////

Public Property Get Visible(ByVal eBar As EFSScrollBarConstants) As Boolean
   If (eBar = efsHorizontal) Then
      Visible = m_bVisibleHorz
   Else
      Visible = m_bVisibleVert
   End If
End Property
Public Property Let Visible(ByVal eBar As EFSScrollBarConstants, ByVal bState As Boolean)
   If (eBar = efsHorizontal) Then
      m_bVisibleHorz = bState
   Else
      m_bVisibleVert = bState
   End If
   If (m_bNoFlatScrollBars) Then
      ShowScrollBar m_hWnd, eBar, 1& And (bState)
   Else
      FlatSB_ShowScrollBar m_hWnd, eBar, 1& And (bState)
   End If
End Property

Public Property Get Orientation() As EFSOrientationConstants
   Orientation = m_eOrientation
End Property

Public Property Let Orientation(ByVal eOrientation As EFSOrientationConstants)
   m_eOrientation = eOrientation
   pSetOrientation
End Property

Private Sub pSetOrientation()
   ShowScrollBar m_hWnd, SB_HORZ, 1& And ((m_eOrientation = efsoBoth) Or (m_eOrientation = efsoHorizontal))
   ShowScrollBar m_hWnd, SB_VERT, 1& And ((m_eOrientation = efsoBoth) Or (m_eOrientation = efsoVertical))
End Sub

Private Sub pGetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
Dim lO As Long
    
    lO = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)
    
    If (m_bNoFlatScrollBars) Then
        GetScrollInfo m_hWnd, lO, tSI
    Else
        FlatSB_GetScrollInfo m_hWnd, lO, tSI
    End If

End Sub
Private Sub pLetSI(ByVal eBar As EFSScrollBarConstants, ByRef tSI As SCROLLINFO, ByVal fMask As Long)
Dim lO As Long
        
    lO = eBar
    tSI.fMask = fMask
    tSI.cbSize = LenB(tSI)
    
    If (m_bNoFlatScrollBars) Then
        SetScrollInfo m_hWnd, lO, tSI, True
    Else
        FlatSB_SetScrollInfo m_hWnd, lO, tSI, True
    End If
    
End Sub

Public Property Get Style() As EFSStyleConstants
   Style = m_eStyle
End Property
Public Property Let Style(ByVal eStyle As EFSStyleConstants)
Dim lR As Long
   If (eStyle <> efsRegular) Then
      If (m_bNoFlatScrollBars) Then
         ' can't do it..
         Debug.Print "Can't set non-regular style mode on this system - COMCTL32.DLL version < 4.71."
         Exit Property
      End If
   End If
   If (m_eOrientation = efsoHorizontal) Or (m_eOrientation = efsoBoth) Then
      lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_HSTYLE, eStyle, True)
   End If
   If (m_eOrientation = efsoVertical) Or (m_eOrientation = efsoBoth) Then
      lR = FlatSB_SetScrollProp(m_hWnd, WSB_PROP_VSTYLE, eStyle, True)
   End If
   m_eStyle = eStyle
End Property

Public Property Get SmallChange(ByVal eBar As EFSScrollBarConstants) As Long
   If (eBar = efsHorizontal) Then
      SmallChange = m_lSmallChangeHorz
   Else
      SmallChange = m_lSmallChangeVert
   End If
End Property
Public Property Let SmallChange(ByVal eBar As EFSScrollBarConstants, ByVal lSmallChange As Long)
   If (eBar = efsHorizontal) Then
      m_lSmallChangeHorz = lSmallChange
   Else
      m_lSmallChangeVert = lSmallChange
   End If
End Property
Public Property Get Enabled(ByVal eBar As EFSScrollBarConstants) As Boolean
   If (eBar = efsHorizontal) Then
      Enabled = m_bEnabledHorz
   Else
      Enabled = m_bEnabledVert
   End If
End Property
Public Property Let Enabled(ByVal eBar As EFSScrollBarConstants, ByVal bEnabled As Boolean)
Dim lO As Long
Dim lf As Long
        
   lO = eBar
   If (bEnabled) Then
      lf = ESB_ENABLE_BOTH
   Else
      lf = ESB_DISABLE_BOTH
   End If
   If (m_bNoFlatScrollBars) Then
      EnableScrollBar m_hWnd, lO, lf
   Else
      FlatSB_EnableScrollBar m_hWnd, lO, lf
   End If
   'add:fix the bug!
   If (eBar = efsHorizontal) Then
      m_bEnabledHorz = bEnabled
   Else
      m_bEnabledVert = bEnabled
   End If
End Property
Public Property Get Min(ByVal eBar As EFSScrollBarConstants) As Long
Dim tSI As SCROLLINFO
    pGetSI eBar, tSI, SIF_RANGE
    Min = tSI.nMin
End Property
Public Property Get Max(ByVal eBar As EFSScrollBarConstants) As Long
Dim tSI As SCROLLINFO
    pGetSI eBar, tSI, SIF_RANGE Or SIF_PAGE
    Max = tSI.nMax - tSI.nPage + 1 'modified!!
End Property
Public Property Get Value(ByVal eBar As EFSScrollBarConstants) As Long
Dim tSI As SCROLLINFO
    pGetSI eBar, tSI, SIF_POS
    Value = tSI.nPos
End Property
Public Property Get LargeChange(ByVal eBar As EFSScrollBarConstants) As Long
Dim tSI As SCROLLINFO
    pGetSI eBar, tSI, SIF_PAGE
    LargeChange = tSI.nPage
End Property
Public Property Let Min(ByVal eBar As EFSScrollBarConstants, ByVal iMin As Long)
Dim tSI As SCROLLINFO
    tSI.nMin = iMin
    tSI.nMax = Max(eBar) + LargeChange(eBar) - 1 'modified!!
    pLetSI eBar, tSI, SIF_RANGE
End Property
Public Property Let Max(ByVal eBar As EFSScrollBarConstants, ByVal iMax As Long)
Dim tSI As SCROLLINFO
    tSI.nMax = iMax + LargeChange(eBar) - 1 'modified!!
    tSI.nMin = Min(eBar)
    pLetSI eBar, tSI, SIF_RANGE
    pRaiseEvent eBar, False
End Property
Public Property Let Value(ByVal eBar As EFSScrollBarConstants, ByVal iValue As Long)
Dim tSI As SCROLLINFO
    '///////////////Add!!
    If iValue > Max(eBar) Then
        iValue = Max(eBar)
    ElseIf iValue < Min(eBar) Then
        iValue = Min(eBar)
    End If
    '///////////////
    If (iValue <> Value(eBar)) Then
        tSI.nPos = iValue
        pLetSI eBar, tSI, SIF_POS
        pRaiseEvent eBar, False
    End If
End Property
Public Property Let LargeChange(ByVal eBar As EFSScrollBarConstants, ByVal iLargeChange As Long)
Dim tSI As SCROLLINFO
Dim lCurMax As Long
Dim lCurLargeChange As Long
    
   pGetSI eBar, tSI, SIF_ALL
   tSI.nMax = tSI.nMax - tSI.nPage + iLargeChange
   tSI.nPage = iLargeChange
   pLetSI eBar, tSI, SIF_PAGE Or SIF_RANGE
End Property
Public Property Get CanBeFlat() As Boolean
   CanBeFlat = Not (m_bNoFlatScrollBars)
End Property
Private Sub pCreateScrollBar()
Dim lR As Long
Dim lStyle As Long
Dim hParent As Long

   ' Just checks for flag scroll bars...
   On Error Resume Next
   lR = InitialiseFlatSB(m_hWnd)
   If (Err.Number <> 0) Then
       'Can't find DLL entry point InitializeFlatSB in COMCTL32.DLL
       ' Means we have version prior to 4.71
       ' We get standard scroll bars.
       m_bNoFlatScrollBars = True
   Else
      Style = m_eStyle
   End If
   
End Sub

Private Sub pCreate()
Dim hd As Long
   pClearUp
   If Not Ambient.UserMode Then Exit Sub
   '////////init graphics
   hd = GetDCEx(0, 0, 0)
   m_hMemDC = CreateCompatibleDC(hd)
   ReleaseDC 0, hd
   m_hbm1 = LoadImage(0, 32559&, IMAGE_BITMAP, 0, 0, LR_SHARED)
   m_hbm2 = LoadImage(0, 32660&, IMAGE_BITMAP, 0, 0, LR_SHARED)
   m_hbm3 = LoadImage(0, 32661&, IMAGE_BITMAP, 0, 0, LR_SHARED)
   If m_hbm1 Then
    m_hbmOld = SelectObject(m_hMemDC, m_hbm1)
   ElseIf m_hbm2 Then
    m_hbmOld = SelectObject(m_hMemDC, m_hbm2)
   ElseIf m_hbm3 Then
    m_hbmOld = SelectObject(m_hMemDC, m_hbm3)
   Else
    m_hbmOld = 0
   End If
   If m_hbmOld Then m_nclrBmp = GetPixel(m_hMemDC, 14, 14)
   '////////
   m_hWnd = ContainerHwnd
   pCreateScrollBar
   pAttachMessages
End Sub

Private Sub pCreateWindow()
pDestroyWindow
With m_tMidBtnPos_Old
 m_hWndMidBtn = CreateWindowEx(WS_EX_TOPMOST Or WS_EX_TOOLWINDOW, "static", "", WS_CHILD, _
 .x - 15, .y - 15, 32, 32, m_hWnd, 0, App.hInstance, ByVal 0)
 '///
 m_hRgnWnd = CreateEllipticRgn(0, 0, 32, 32)
 SetWindowRgn m_hWndMidBtn, m_hRgnWnd, 0
End With
'///
SetParent m_hWndMidBtn, 0
'///
cSub2.AddMsg WM_PAINT, MSG_BEFORE
cSub2.Subclass m_hWndMidBtn, Me
'///
ShowWindow m_hWndMidBtn, 5
End Sub

Private Sub pDestroyWindow()
If m_hWndMidBtn Then
 cSub2.UnSubclass
 DestroyWindow m_hWndMidBtn
 If m_hRgnWnd Then DeleteObject m_hRgnWnd
 m_hWndMidBtn = 0
 m_hRgnWnd = 0
End If
End Sub

Private Sub pClearUp()
   If m_hWnd <> 0 Then
      On Error Resume Next
      ' Stop flat scroll bar if we have it:
      If Not (m_bNoFlatScrollBars) Then
         UninitializeFlatSB m_hWnd
      End If
    
      On Error GoTo 0
      ' Remove subclass:
'      cSub.DelMsg -1, MSG_BEFORE
'      cSub.DelMsg -1, MSG_AFTER
      cSub.UnSubclass
   End If
   m_hWnd = 0
   m_bInitialised = False
   '////////destroy graphics
   If m_hMemDC Then
    If m_hbmOld Then SelectObject m_hMemDC, m_hbmOld
    DeleteDC m_hMemDC
    m_hMemDC = 0
   End If
   '////////destroy window
   pDestroyWindow
End Sub

Private Sub pAttachMessages()
   If (m_hWnd <> 0) Then
      cSub.AddMsg WM_HSCROLL, MSG_AFTER
      cSub.AddMsg WM_VSCROLL, MSG_AFTER
      cSub.AddMsg WM_MOUSEWHEEL, MSG_AFTER
      cSub.AddMsg WM_NCLBUTTONDOWN, MSG_AFTER
      'cSub.AddMsg WM_NCMBUTTONDOWN, MSG_AFTER
      cSub.AddMsg WM_NCRBUTTONDOWN, MSG_AFTER
      '////////Add!!!
      If mode2 > 0 Then cSub.AddMsg WM_NCPAINT, MSG_AFTER
      '////////Add!!! mid button scroll
      cSub.AddMsg WM_TIMER, MSG_BEFORE
      cSub.AddMsg WM_LBUTTONDOWN, MSG_BEFORE
      cSub.AddMsg WM_MBUTTONDOWN, MSG_BEFORE
      cSub.AddMsg WM_RBUTTONDOWN, MSG_BEFORE
      cSub.AddMsg WM_LBUTTONUP, MSG_BEFORE
      cSub.AddMsg WM_MBUTTONUP, MSG_BEFORE
      cSub.AddMsg WM_RBUTTONUP, MSG_BEFORE
      cSub.AddMsg WM_MOUSEMOVE, MSG_BEFORE
      '////////
      cSub.Subclass m_hWnd, Me
      SystemParametersInfo SPI_GETWHEELSCROLLLINES, _
            0, m_lWheelScrollLines, 0
      If (m_lWheelScrollLines <= 0) Then
         m_lWheelScrollLines = 3
      End If
      m_bInitialised = True
   End If
End Sub

Private Sub iSubclass_Before(bHandled As Boolean, lReturn As Long, hwnd As Long, uMsg As Long, wParam As Long, lParam As Long)
Dim i As Long, j As Long, k As Long
Static nTimeX As Long, nTimeY As Long
Select Case uMsg
Case WM_LBUTTONDOWN, WM_MBUTTONDOWN, WM_RBUTTONDOWN
 If m_bInMidButton Then
  KillTimer hwnd, &HDEADBEEF
  '///
  pDestroyWindow
  '///
  ReleaseCapture
  m_bInMidButton = False
  lReturn = 0
  bHandled = True
 ElseIf uMsg = WM_MBUTTONDOWN Then
  If ((m_eOrientation = efsoHorizontal Or m_eOrientation = efsoBoth) And m_bEnabledHorz) _
  Or ((m_eOrientation = efsoVertical Or m_eOrientation = efsoBoth) And m_bEnabledVert) Then
   SetTimer hwnd, &HDEADBEEF, 20, 0
   nTimeX = 0
   nTimeY = 0
   '///
   GetCursorPos m_tMidBtnPos
   m_tMidBtnPos_Old = m_tMidBtnPos
   '///show window
   If m_hbm1 <> 0 And m_hbm2 <> 0 And m_hbm3 <> 0 Then pCreateWindow
   '///
   SetCapture hwnd
   m_bInMidButton = True
   lReturn = 0
   bHandled = True
  End If
 End If
Case WM_MOUSEMOVE, WM_LBUTTONUP, WM_MBUTTONUP, WM_RBUTTONUP
 If m_bInMidButton Then
  '///get cursor pos
  GetCursorPos m_tMidBtnPos
  If ((m_eOrientation = efsoHorizontal Or m_eOrientation = efsoBoth) And m_bEnabledHorz) Then
   If ((m_eOrientation = efsoVertical Or m_eOrientation = efsoBoth) And m_bEnabledVert) Then
    If m_tMidBtnPos.y - m_tMidBtnPos_Old.y > 16 Then
     If m_tMidBtnPos.x - m_tMidBtnPos_Old.x > 16 Then i = 32662& Else If m_tMidBtnPos.x - m_tMidBtnPos_Old.x < -16 Then i = 32661& Else i = 32656&
    ElseIf m_tMidBtnPos.y - m_tMidBtnPos_Old.y < -16 Then
     If m_tMidBtnPos.x - m_tMidBtnPos_Old.x > 16 Then i = 32660& Else If m_tMidBtnPos.x - m_tMidBtnPos_Old.x < -16 Then i = 32659& Else i = 32655&
    Else
     If m_tMidBtnPos.x - m_tMidBtnPos_Old.x > 16 Then i = 32658& Else If m_tMidBtnPos.x - m_tMidBtnPos_Old.x < -16 Then i = 32657& Else i = 32654&
    End If
   Else
    If m_tMidBtnPos.x - m_tMidBtnPos_Old.x > 16 Then i = 32658& Else If m_tMidBtnPos.x - m_tMidBtnPos_Old.x < -16 Then i = 32657& Else i = 32653&
   End If
  Else
   If m_tMidBtnPos.y - m_tMidBtnPos_Old.y > 16 Then i = 32656& Else If m_tMidBtnPos.y - m_tMidBtnPos_Old.y < -16 Then i = 32655& Else i = 32652&
  End If
  '///
  i = LoadCursor(0, i)
  If i Then
   SetCursor i
   lReturn = 0
   bHandled = True
  End If
 End If
Case WM_PAINT 'paint sub window
 pPaint hwnd
 lReturn = 0
 bHandled = True
Case WM_TIMER
 If wParam = &HDEADBEEF Then
  'TODO:smallchange,etc.
  If ((m_eOrientation = efsoHorizontal Or m_eOrientation = efsoBoth) And m_bEnabledHorz) Then
   i = m_tMidBtnPos.x - m_tMidBtnPos_Old.x
   If i < 0 Then i = -i
   If i > 16 Then i = i \ 16 Else i = 0
  Else
   i = 0
  End If
  If ((m_eOrientation = efsoVertical Or m_eOrientation = efsoBoth) And m_bEnabledVert) Then
   j = m_tMidBtnPos.y - m_tMidBtnPos_Old.y
   If j < 0 Then j = -j
   If j > 16 Then j = j \ 16 Else j = 0
  Else
   j = 0
  End If
  If i = 1 Then nTimeX = nTimeX + 1 Else If i = 2 Then nTimeX = nTimeX + 2 Else If i > 2 Then nTimeX = nTimeX + 4
  If j = 1 Then nTimeY = nTimeY + 1 Else If j = 2 Then nTimeY = nTimeY + 2 Else If j > 2 Then nTimeY = nTimeY + 4
  '///
  If nTimeX >= 4 Then
   nTimeX = nTimeX - 4
   i = i - 2
   If i > 1 Then k = i Else k = 1
   i = i - 12
   If i > 0 Then k = k + i * i
   If m_tMidBtnPos.x < m_tMidBtnPos_Old.x Then k = -k
   Value(efsHorizontal) = Value(efsHorizontal) + k
  End If
  If nTimeY >= 4 Then
   nTimeY = nTimeY - 4
   j = j - 2
   If j > 1 Then k = j Else k = 1
   j = j - 12
   If j > 0 Then k = k + j * j
   If m_tMidBtnPos.y < m_tMidBtnPos_Old.y Then k = -k
   Value(efsVertical) = Value(efsVertical) + k
  End If
  '///
  lReturn = 0
  bHandled = True
 End If
End Select
End Sub

Private Sub pPaint(ByVal hwnd As Long)
Dim tPaint As PAINTSTRUCT
Dim i As Long, j As Long
BeginPaint hwnd, tPaint
'///
If (m_eOrientation = efsoHorizontal Or m_eOrientation = efsoBoth) And m_bEnabledHorz Then
 If (m_eOrientation = efsoVertical Or m_eOrientation = efsoBoth) And m_bEnabledVert Then i = m_hbm3 Else i = m_hbm2
Else
 i = m_hbm1
End If
SelectObject m_hMemDC, i
BitBlt tPaint.hdc, 2, 9, 28, 14, m_hMemDC, 0, 7, vbSrcCopy
BitBlt tPaint.hdc, 5, 5, 22, 22, m_hMemDC, 3, 3, vbSrcCopy
BitBlt tPaint.hdc, 9, 2, 14, 28, m_hMemDC, 7, 0, vbSrcCopy
i = SelectObject(tPaint.hdc, GetStockObject(NULL_BRUSH))
j = SelectObject(tPaint.hdc, CreatePen(0, 2, m_nclrBmp))
Ellipse tPaint.hdc, 1, 1, 30, 30
SelectObject tPaint.hdc, i
DeleteObject SelectObject(tPaint.hdc, j)
'///
EndPaint hwnd, tPaint
End Sub

Private Sub UserControl_Terminate()
   pClearUp
End Sub

Private Sub UserControl_Initialize()
   m_lSmallChangeHorz = 1
   m_lSmallChangeVert = 1
   m_eStyle = efsRegular
   m_eOrientation = efsoBoth
End Sub

Private Sub UserControl_InitProperties()
clr1 = vbApplicationWorkspace
clr2 = vbApplicationWorkspace
bdrWidth = 1
pCreate
End Sub

Private Sub UserControl_ReadProperties(PropBag As PropertyBag)
With PropBag
 Orientation = .ReadProperty("Orientation", m_eOrientation)
 Style = .ReadProperty("Style", m_eStyle)
 mode2 = .ReadProperty("NCPaintMode", 0)
 clr1 = .ReadProperty("NCPaintColor1", vbApplicationWorkspace)
 clr2 = .ReadProperty("NCPaintColor2", vbApplicationWorkspace)
 bdrWidth = .ReadProperty("NCBorderWidth", 1)
End With
pCreate
End Sub

Private Sub UserControl_WriteProperties(PropBag As PropertyBag)
With PropBag
 .WriteProperty "Orientation", m_eOrientation, efsoBoth
 .WriteProperty "Style", m_eStyle, efsRegular
 .WriteProperty "NCPaintMode", mode2, 0
 .WriteProperty "NCPaintColor1", clr1, vbApplicationWorkspace
 .WriteProperty "NCPaintColor2", clr2, vbApplicationWorkspace
 .WriteProperty "NCBorderWidth", bdrWidth, 1
End With
End Sub

Private Sub iSubclass_After(lReturn As Long, ByVal hwnd As Long, ByVal uMsg As Long, ByVal wParam As Long, ByVal lParam As Long)
If uMsg = WM_NCPAINT Then
 pNCPaint hwnd, wParam
 lReturn = 0
Else
 lReturn = pWindowProc(hwnd, uMsg, wParam, lParam)
End If
End Sub

Private Sub pSolid(ByVal hdc As Long, ByVal w As Long, ByVal h As Long)
Dim hbr As Long
Dim r As RECT
hbr = CreateSolidBrush(TranslateColor(clr1))
If bdrWidth > 0 Then
 'top
 r.Right = w
 r.Bottom = bdrWidth
 FillRect hdc, r, hbr
 'bottom
 r.Top = h - bdrWidth
 r.Bottom = h
 FillRect hdc, r, hbr
 'left
 r.Bottom = r.Top
 r.Top = bdrWidth
 r.Right = bdrWidth
 FillRect hdc, r, hbr
 'right
 r.Left = w - bdrWidth
 r.Right = w
 FillRect hdc, r, hbr
Else
 r.Right = w
 r.Bottom = h
 FillRect hdc, r, hbr
End If
DeleteObject hbr
End Sub

Private Sub pNCPaint(ByVal hwnd As Long, ByVal wParam As Long)
Dim hd As Long, hbr As Long
Dim r As RECT
 hd = GetDCEx(hwnd, wParam, DCX_WINDOW Or DCX_INTERSECTRGN)
 '///
 GetWindowRect hwnd, r
 r.Right = r.Right - r.Left
 r.Bottom = r.Bottom - r.Top
 Select Case mode2
 Case 1
  pSolid hd, r.Right, r.Bottom
 Case 2
  If clr1 = clr2 Then
   pSolid hd, r.Right, r.Bottom
  Else
   If bdrWidth > 0 Then
    'top
    GradientFillRect hd, 0, 0, r.Right, bdrWidth, clr1, clr2, GRADIENT_FILL_RECT_H
    'bottom
    GradientFillRect hd, 0, r.Bottom - bdrWidth, r.Right, r.Bottom, clr1, clr2, GRADIENT_FILL_RECT_H
    'left
    StretchBlt hd, 0, 0, bdrWidth, r.Bottom, hd, 0, 0, bdrWidth, bdrWidth, vbSrcCopy
    'right
    StretchBlt hd, r.Right - bdrWidth, 0, bdrWidth, r.Bottom, hd, r.Right - bdrWidth, 0, bdrWidth, bdrWidth, vbSrcCopy
   Else
    GradientFillRect hd, 0, 0, r.Right, r.Bottom, clr1, clr2, GRADIENT_FILL_RECT_H
   End If
  End If
 Case 3
  If clr1 = clr2 Then
   pSolid hd, r.Right, r.Bottom
  Else
   If bdrWidth > 0 Then
    'left
    GradientFillRect hd, 0, 0, bdrWidth, r.Bottom, clr1, clr2, GRADIENT_FILL_RECT_V
    'right
    GradientFillRect hd, r.Right - bdrWidth, 0, r.Right, r.Bottom, clr1, clr2, GRADIENT_FILL_RECT_V
    'top
    StretchBlt hd, 0, 0, r.Right, bdrWidth, hd, 0, 0, bdrWidth, bdrWidth, vbSrcCopy
    'bottom
    StretchBlt hd, 0, r.Bottom - bdrWidth, r.Right, bdrWidth, hd, 0, r.Bottom - bdrWidth, bdrWidth, bdrWidth, vbSrcCopy
   Else
    GradientFillRect hd, 0, 0, r.Right, r.Bottom, clr1, clr2, GRADIENT_FILL_RECT_V
   End If
  End If
 Case 99
  RaiseEvent Paint(hd, r.Right, r.Bottom)
 End Select
 ReleaseDC hwnd, hd
End Sub

Private Function pWindowProc(ByVal hwnd As Long, ByVal iMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Dim lScrollCode As Long
Dim tSI As SCROLLINFO
Dim lV As Long, lSC As Long
Dim eBar As EFSScrollBarConstants
Dim zDelta As Long
Dim lDelta As Long
Dim wMKeyFlags As Long

   Select Case iMsg
   Case WM_MOUSEWHEEL
      ' Low-word of wParam indicates whether virtual keys
      ' are down
      wMKeyFlags = wParam And &HFFFF&
      ' High order word is the distance the wheel has been rotated,
      ' in multiples of WHEEL_DELTA:
      If (wParam And &H8000000) Then
         ' Towards the user:
         zDelta = &H8000& - (wParam And &H7FFF0000) \ &H10000
      Else
         ' Away from the user:
         zDelta = -((wParam And &H7FFF0000) \ &H10000)
      End If
      '////////////Add!!!
      If wMKeyFlags And &H4& Then
         eBar = efsHorizontal
      Else
         eBar = efsVertical
      End If
      '////////////
      lDelta = (zDelta \ WHEEL_DELTA) * SmallChange(eBar) * m_lWheelScrollLines
      RaiseEvent MouseWheel(eBar, lDelta)
      If Not (lDelta = 0) Then
         Value(eBar) = Value(eBar) + lDelta
         pWindowProc = 1
      End If
   
   Case WM_VSCROLL, WM_HSCROLL
      If (iMsg = WM_HSCROLL) Then
         eBar = efsHorizontal
      Else
         eBar = efsVertical
      End If
      lScrollCode = (wParam And &HFFFF&)
      Select Case lScrollCode
      Case SB_THUMBTRACK
         ' Is vertical/horizontal?
         pGetSI eBar, tSI, SIF_TRACKPOS
         Value(eBar) = tSI.nTrackPos
         pRaiseEvent eBar, True
         
      Case SB_LEFT, SB_TOP
         Value(eBar) = Min(eBar)
         pRaiseEvent eBar, False
         
      Case SB_RIGHT, SB_BOTTOM
         Value(eBar) = Max(eBar)
         pRaiseEvent eBar, False
          
      Case SB_LINELEFT, SB_LINEUP
         'Debug.Print "Line"
         lV = Value(eBar)
         If (eBar = efsHorizontal) Then
            lSC = m_lSmallChangeHorz
         Else
            lSC = m_lSmallChangeVert
         End If
         If (lV - lSC < Min(eBar)) Then
            Value(eBar) = Min(eBar)
         Else
            Value(eBar) = lV - lSC
         End If
         pRaiseEvent eBar, False
         
      Case SB_LINERIGHT, SB_LINEDOWN
          'Debug.Print "Line"
         lV = Value(eBar)
         If (eBar = efsHorizontal) Then
            lSC = m_lSmallChangeHorz
         Else
            lSC = m_lSmallChangeVert
         End If
         If (lV + lSC > Max(eBar)) Then
            Value(eBar) = Max(eBar)
         Else
            Value(eBar) = lV + lSC
         End If
         pRaiseEvent eBar, False
          
      Case SB_PAGELEFT, SB_PAGEUP
         Value(eBar) = Value(eBar) - LargeChange(eBar)
         pRaiseEvent eBar, False
         
      Case SB_PAGERIGHT, SB_PAGEDOWN
         Value(eBar) = Value(eBar) + LargeChange(eBar)
         pRaiseEvent eBar, False
         
      Case SB_ENDSCROLL
         pRaiseEvent eBar, False
         
      End Select
      
   Case WM_NCLBUTTONDOWN, WM_NCRBUTTONDOWN
      Dim eBtn As MouseButtonConstants
      eBtn = IIf(iMsg = WM_NCLBUTTONDOWN, vbLeftButton, vbRightButton)
      If wParam = HTVSCROLL Then
         RaiseEvent ScrollClick(efsHorizontal, eBtn)
      ElseIf wParam = HTHSCROLL Then
         RaiseEvent ScrollClick(efsVertical, eBtn)
      End If
      
   End Select

End Function

Private Function pRaiseEvent(ByVal eBar As EFSScrollBarConstants, ByVal bScroll As Boolean)
Static s_lLastValue(0 To 1) As Long
   If (Value(eBar) <> s_lLastValue(eBar)) Then
      If (bScroll) Then
         RaiseEvent Scroll(eBar)
      Else
         RaiseEvent Change(eBar)
      End If
      s_lLastValue(eBar) = Value(eBar)
   End If
   
End Function






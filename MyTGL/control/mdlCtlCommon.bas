Attribute VB_Name = "mdlCtlCommon"
Option Explicit

'////////////////////////////////
'This file is public domain.
'////////////////////////////////

'///////////////////////////////some functions

Private Declare Function GetSysColor Lib "user32.dll" (ByVal nIndex As Long) As Long

Private Type TRIVERTEX
   x As Long
   y As Long
   Red0 As Byte
   Red1 As Byte
   Green0 As Byte
   Green1 As Byte
   Blue0 As Byte
   Blue1 As Byte
   Alpha0 As Byte
   Alpha1 As Byte
End Type

Private Type GRADIENT_RECT
   UpperLeft As Long
   LowerRight As Long
End Type

Private Declare Function GradientFill Lib "msimg32" ( _
   ByVal hdc As Long, _
   pVertex As TRIVERTEX, _
   ByVal dwNumVertex As Long, _
   pMesh As Any, _
   ByVal dwNumMesh As Long, _
   ByVal dwMode As Long) As Long

Public Enum GradientFillStyle
 GRADIENT_FILL_RECT_H = 0&
 GRADIENT_FILL_RECT_V = 1&
' GRADIENT_FILL_TRIANGLE = &H2&
End Enum

'///////////////////////////////the color constants

Public Const d_Bg = &HF6F6F6
Public Const d_Border = &H800000
Public Const d_Title1 = &HD68759
Public Const d_Title2 = &H9A400C
Public Const d_Bar1 = &HFAE2D0
Public Const d_Bar2 = &HE2A981
Public Const d_Hl1 = &HD0FCFD
Public Const d_Hl2 = &H9DDFFD
Public Const d_Checked1 = &H7DDDFA
Public Const d_Checked2 = &H4EBCF5
Public Const d_Pressed1 = &H5586F8
Public Const d_Pressed2 = &HA37D2
Public Const d_Sprt1 = &HCB8C6A
Public Const d_Sprt2 = vbWhite
Public Const d_Text = vbBlack
Public Const d_TextHl = vbBlack
Public Const d_TextDis = &HCB8C6A

Public Const d_Chevron1 = &HF1A675
Public Const d_Chevron2 = &H913500

Public Const d_BorderP = &HC56A31
Public Const d_SprtP = &H99A8AC
Public Const d_HlP = &HEDD2C1
Public Const d_CheckedP = &HD7DDFA
Public Const d_PressedP = &HE2B598

Public Const d_Gripper = &H764127

Public Const d_CtrlBorder = &HB99D7F
Public Const d_Icon_Grayscale = &HFFE7D5

'///////////////////////////////some new functions

Private Const GCL_STYLE As Long = -26
Private Const CS_DROPSHADOW = &H20000
Private Declare Function GetClassLong Lib "user32.dll" Alias "GetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long) As Long
Private Declare Function SetClassLong Lib "user32.dll" Alias "SetClassLongA" (ByVal hwnd As Long, ByVal nIndex As Long, ByVal dwNewLong As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Declare Function FindWindowEx Lib "user32.dll" Alias "FindWindowExA" (ByVal hWnd1 As Long, ByVal hWnd2 As Long, ByVal lpsz1 As String, ByVal lpsz2 As String) As Long
Private Declare Function GetWindowThreadProcessId Lib "user32.dll" (ByVal hwnd As Long, ByRef lpdwProcessId As Long) As Long
Private Declare Function GetCurrentProcessId Lib "kernel32.dll" () As Long

Private Declare Sub InitCommonControls Lib "comctl32.dll" ()
Private Declare Function GetSystemDirectory Lib "kernel32.dll" Alias "GetSystemDirectoryA" (ByVal lpBuffer As String, ByVal nSize As Long) As Long
Private Declare Function CreateActCtxW Lib "kernel32.dll" (ByRef pActCtx As ACTCTXW) As Long
'Private Declare Sub ReleaseActCtx Lib "kernel32.dll" (ByVal hActCtx As Long)
Private Declare Function ActivateActCtx Lib "kernel32.dll" (ByVal hActCtx As Long, ByRef lplpCookie As Long) As Long
'Private Declare Function DeactivateActCtx Lib "kernel32.dll" (ByVal dwFlags As Long, ByVal lpCookie As Long) As Long

Private Type ACTCTXW
 cbSize As Long
 dwFlags As Long
 lpcwstrSource As Long
 wProcessorArchitecture As Integer
 wLangId As Integer
 lpcwstrAssemblyDirectory As Long
 lpcwstrResourceName As Long
 lpcwstrApplicationName As Long
 hModule As Long
End Type

Private Const ACTCTX_FLAG_PROCESSOR_ARCHITECTURE_VALID As Long = 1
Private Const ACTCTX_FLAG_LANGID_VALID As Long = 2
Private Const ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID As Long = 4
Private Const ACTCTX_FLAG_RESOURCE_NAME_VALID As Long = 8
Private Const ACTCTX_FLAG_SET_PROCESS_DEFAULT As Long = 16
Private Const ACTCTX_FLAG_APPLICATION_NAME_VALID As Long = 32
Private Const ACTCTX_FLAG_HMODULE_VALID As Long = 128

Public Sub NewLoadManifest()
On Error GoTo a
Dim t As ACTCTXW, s As String
Dim h As Long, i As Long
Debug.Print 1 \ 0
InitCommonControls
t.cbSize = Len(t)
's = Environ("path")
'i = InStr(s, ";")
'If i > 0 Then s = Left(s, i - 1)
'///
s = Space(1024)
GetSystemDirectory s, 1024
i = InStr(s, vbNullChar)
If i > 0 Then s = Left(s, i - 1)
'///
s = s + "\shell32.dll"
t.lpcwstrSource = StrPtr(s)
t.lpcwstrResourceName = 124
t.dwFlags = ACTCTX_FLAG_RESOURCE_NAME_VALID
h = CreateActCtxW(t)
If h <> -1 And h <> 0 Then ActivateActCtx h, i
a:
End Sub

Public Sub EnableDropShadow(ByVal hwnd As Long)
    SetClassLong hwnd, GCL_STYLE, GetClassLong(hwnd, GCL_STYLE) Or CS_DROPSHADOW
End Sub

Public Sub DisbleDropShadow(ByVal hwnd As Long)
    SetClassLong hwnd, GCL_STYLE, GetClassLong(hwnd, GCL_STYLE) And Not CS_DROPSHADOW
End Sub

Private Function pFindWindowByClassName(ByVal sClassName As String) As Long
Dim hwd As Long, pid As Long
Dim p As Long
pid = GetCurrentProcessId
Do
 hwd = FindWindowEx(0, hwd, sClassName, vbNullString)
 If hwd <> 0 Then
  GetWindowThreadProcessId hwd, p
  If p = pid Then
   pFindWindowByClassName = hwd
   Exit Function
  End If
 End If
Loop Until hwd = 0
End Function

Public Function EnableTooltipDropShadow() As Boolean
Dim hwd As Long
Static b As Boolean
If b Then
 EnableTooltipDropShadow = True
 Exit Function
End If
hwd = pFindWindowByClassName("VBBubble")
If hwd <> 0 Then
 EnableDropShadow hwd
 b = True
End If
hwd = pFindWindowByClassName("VBBubbleRT5")
If hwd <> 0 Then
 EnableDropShadow hwd
 b = True
End If
hwd = pFindWindowByClassName("VBBubbleRT6")
If hwd <> 0 Then
 EnableDropShadow hwd
 b = True
End If
EnableTooltipDropShadow = b
End Function

'///////////////////////////////some functions

Public Function TranslateColor(ByVal clr As Long) As Long
If clr < 0 Then
 TranslateColor = GetSysColor(clr And &HFFFFFF)
Else
 TranslateColor = clr
End If
End Function

Public Sub GradientFillRect( _
      ByVal lHDC As Long, _
      ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, _
      ByVal lStartColor As Long, _
      ByVal lEndColor As Long, _
      ByVal eDir As GradientFillStyle _
   )

    lStartColor = TranslateColor(lStartColor)
    lEndColor = TranslateColor(lEndColor)

    Dim tTV(0 To 1) As TRIVERTEX
    Dim tGR As GRADIENT_RECT

    pSetTriVertexColor tTV(0), lStartColor
    tTV(0).x = Left
    tTV(0).y = Top
    pSetTriVertexColor tTV(1), lEndColor
    tTV(1).x = Right
    tTV(1).y = Bottom

    tGR.UpperLeft = 0
    tGR.LowerRight = 1

    GradientFill lHDC, tTV(0), 2, tGR, 1, eDir

End Sub

Private Sub pSetTriVertexColor(tTV As TRIVERTEX, lColor As Long)
   tTV.Red1 = (lColor And &HFF&)
   tTV.Green1 = (lColor And &HFF00&) \ &H100&
   tTV.Blue1 = (lColor And &HFF0000) \ &H10000
End Sub

'Public Function BlendColor(ByVal clr1 As Long, ByVal clr2 As Long, ByVal p As Long) As Long
'clr1 = TranslateColor(clr1)
'clr2 = TranslateColor(clr2)
'BlendColor = ((clr1 + (((clr2 And &HFF&) - (clr1 And &HFF&)) * p) \ &H100&) And &HFF&) Or _
'(((clr1 And &HFF00&) + (((clr2 And &HFF00&) - (clr1 And &HFF00&)) * p) \ &H100&) And &HFF00&) Or _
'(((clr1 And &HFF0000) + (((clr2 And &HFF0000) - (clr1 And &HFF0000)) \ &H100&) * p) And &HFF0000)
'End Function

Public Sub GrayscaleBitmap(bm As cDIBSection, bmOut As cDIBSection, ByVal clr As Long, ByVal clrTrans As Long)
Dim i As Long, j As Long, ii As Long, m As Long
Dim nClrBlue As Long, nClrGreen As Long, nClrRed As Long
Dim nTransBlue As Long, nTransGreen As Long, nTransRed As Long
Dim k As Long
Dim b() As Byte
If bm.Width <= 0 Or bm.Height <= 0 Then Exit Sub
bmOut.Create bm.Width, bm.Height
m = bm.BytesPerScanLine
ReDim b(m - 1)
nClrBlue = (clr And &HFF0000) \ &H10000
nClrGreen = (clr And &HFF00&) \ &H100&
nClrRed = clr And &HFF&
nTransBlue = (clrTrans And &HFF0000) \ &H10000
nTransGreen = (clrTrans And &HFF00&) \ &H100&
nTransRed = clrTrans And &HFF&
For j = 0 To bm.Height - 1
 CopyMemory b(0), ByVal bm.DIBSectionBitsPtr + j * m, m
 ii = 0
 For i = 0 To bm.Width - 1
  If b(ii) <> nTransBlue Or b(ii + 1) <> nTransGreen Or b(ii + 2) <> nTransRed Then
'   k = b(ii) * 146& + b(ii + 1) * 1454& + b(ii + 2) * 456& + 512&
   k = (CLng(b(ii)) + b(ii + 1) + b(ii + 2)) * 685& + 512&
   b(ii) = (nClrBlue * k) \ 524288
   b(ii + 1) = (nClrGreen * k) \ 524288
   b(ii + 2) = (nClrRed * k) \ 524288
  End If
  ii = ii + 3
 Next i
 CopyMemory ByVal bmOut.DIBSectionBitsPtr + j * m, b(0), m
Next j
End Sub

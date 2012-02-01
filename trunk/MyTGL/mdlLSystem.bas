Attribute VB_Name = "mdlLSystem"
Option Explicit

'///////////////////////////
'
' 2D L-system test
'
'///////////////////////////

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

Private Declare Function CreatePen Lib "gdi32.dll" (ByVal nPenStyle As Long, ByVal nWidth As Long, ByVal crColor As Long) As Long
Private Declare Function DeleteObject Lib "gdi32.dll" (ByVal hObject As Long) As Long
Private Declare Function SelectObject Lib "gdi32.dll" (ByVal hdc As Long, ByVal hObject As Long) As Long
Private Declare Function LineTo Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long) As Long
Private Declare Function MoveToEx Lib "gdi32.dll" (ByVal hdc As Long, ByVal x As Long, ByVal y As Long, ByRef lpPoint As Any) As Long
Private Declare Function LineDDA Lib "gdi32.dll" (ByVal nXStart As Long, ByVal nYStart As Long, ByVal nXEnd As Long, ByVal nYEnd As Long, ByVal lpLineFunc As Long, ByVal lpData As Long) As Long 'stupid...
'LINEDDAPROC=sub(byval x as long,byval y as long,byval data as long)

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

Private Declare Function CreateSolidBrush Lib "gdi32.dll" (ByVal crColor As Long) As Long
Private Declare Function Polygon Lib "gdi32.dll" (ByVal hdc As Long, ByRef lpPoint As POINTAPI, ByVal nCount As Long) As Long
Private Declare Function GetStockObject Lib "gdi32.dll" (ByVal nIndex As Long) As Long
Private Const NULL_PEN As Long = 8
Private Type POINTAPI
    x As Long
    y As Long
End Type

'compile?
Private m_sExp As String, m_lps As Long, m_lpm As Long

Private Type typeZeroL_Operation '2 bytes
 nType As Byte
 nIndex As Byte
End Type

'pre-defined type:
'name asc  meaning
'  F   70  Move forward and draw a line.
'  f  102  Move forward without drawing a line.
'  G   71  Move forward and draw a line. Do not record a vertex.
'  g  103  Move forward without drawing a line. Do not record a vertex.
'  .   46  Record a vertex in the current polygon.
'  +   43  Turn left.
'  -   45  Turn right.
'* ^   94  Pitch up.
'* &   38  Pitch down.
'* \   92  Roll left.
'* /   47  Roll right.
'  |  124  Turn around.
'* $   36  Rotate the turtle to vertical.
'  [   91  Start a branch.
'  ]   93  Complete a branch.
'  {  123  Start a polygon.
'  }  125  Complete a polygon.
'* ~  126  ??? Incorporate a predefined surface. ???
'  !   33  ??? Decrement the diameter of segments. ???
'  '   39  Increment the current color index.
'  %   37  ??? Cut off the remainder of the branch. ???

Private Type typeZeroL_Rule
 nCount As Long
 nFreq As Long
 idxNext As Long 'linked-list
 d() As typeZeroL_Operation
End Type

Private Type typeZeroL
 nCount As Long
 d() As typeZeroL_Rule
 nTable() As Long '256x256 (x4=256kb)
End Type

Private Type typeArgZeroL_Operation '32 bytes
 nType As Byte
 nIndex As Byte
 nArgCount As Integer
 fArg(6) As Single 'count=7
End Type

Private Type typeArgZeroL_Expression
 nCount As Long
 d() As Byte 'p-code!
End Type

'p-code: stupid!!!
'00 xx push arg
'01 xx push const
'02    +
'03    -
'04    *
'05    /
'06    ^
'07    <
'08    >
'09    =
'0A    &(and)
'0B    |(or)
'0C    negative
'0D    !(not)
'0E    <=
'0F    >=
'10    <>

Private Type typeArgZeroL_OperationD
 nType As Byte
 nIndex As Byte
 nArgCount As Integer
 tExp() As typeArgZeroL_Expression '0-based
End Type

Private Type typeArgZeroL_Rule
 nArgCount As Long
 tExp As typeArgZeroL_Expression
 nCount As Long
' nFreq As Long
 idxNext As Long
 d() As typeArgZeroL_OperationD
End Type

Private Type typeArgZeroL
 nCount As Long
 d() As typeArgZeroL_Rule
 fConst(255) As Single 'max=256!
 nTable() As Long
End Type

Private Type typeTwoL_Rule
 nCount As Long
' nFreq As Long
 idxNext As Long
 LeftContextCount As Long
 LeftContext() As typeZeroL_Operation 'must be a path
 RightContextCount As Long
 RightContext() As typeZeroL_Operation 'may be an axial tree TODO:process "[" and "]"
 d() As typeZeroL_Operation
End Type

Private Type typeTwoL
 nCount As Long
 d() As typeTwoL_Rule
 nTable() As Long '256x256
 bIsIgnored() As Byte '256x256
End Type

Private Type typeArgTwoL_OperationContext
 nType As Byte
 nIndex As Byte
 nArgCount As Byte
 nReserved As Byte
End Type

Private Type typeArgTwoL_Rule
 nArgCount As Long
 tExp As typeArgZeroL_Expression
 nCount As Long
' nFreq As Long
 idxNext As Long
 LeftContextCount As Long
 LeftContext() As typeArgTwoL_OperationContext
 RightContextCount As Long
 RightContext() As typeArgTwoL_OperationContext
 d() As typeArgZeroL_OperationD
End Type

Private Type typeArgTwoL
 nCount As Long
 d() As typeArgTwoL_Rule
 fConst(255) As Single
 nTable() As Long
 bIsIgnored() As Byte
End Type

Private Type typeTwoL_String
 nCount As Long
 nMax As Long
 d() As typeZeroL_Operation '0-based
End Type

Private Type typeArgTwoL_String
 nCount As Long
 nMax As Long
 d() As typeArgZeroL_Operation '0-based
End Type

'double????
Private Type typeLSystemState '32 bytes (16 bytes??)
 x As Single
 y As Single
 FS As Single
 idxClr As Integer
 idxWidth As Integer
 nReserved As Long
 nReserved2 As Long
 nReserved3 As Long
 nReserved4 As Long
End Type

Private Type typeLSystemRecursive 'call stack
 nIndex As Long
 nPos As Long
 nReserved As Long
 nReserved2 As Long
End Type

Private Type typeLSystemPoint 'for draw polygon
 x As Single
 y As Single
End Type

Private Type typeSymbolTable '?????
 nCount As Long
 s() As String '0-based
End Type

Private Type typeSymbolTables '?????
 nCount As Long
 d() As typeSymbolTable
End Type

Private Const π As Single = 3.14159265358979
Private Const 二π As Single = π * 2
Private Const TheConst As Single = π / 180

#Const UseLineDDA = True

Private TheBitmap As typeAlphaDibSectionDescriptor
Private TheSA As SAFEARRAY2D
Private TheArray() As RGBQUAD
Private TheClr As RGBQUAD 'current color
Private TheClrIndex As Long 'current index

Private Type typeDrawPolygonEdge '32 bytes :-3
 idx1 As Long
 idx2 As Long
 idxNext As Long 'linked list
 dy As Long
 BigDelta As Long 'may be < 0
 SmallDelta As Long 'must >= 0
 x As Long
 eps As Long
End Type

'calc

Private TheCalcStack(1023) As Single

#If UseLineDDA Then
'MUST check subscript!!!
Private Sub pDrawLine(ByVal x As Long, ByVal y As Long, ByVal d As Long)
TheArray(x And (TheBitmap.Width - 1), y And (TheBitmap.Height - 1)) = TheClr
End Sub

Private Sub pDrawLineAlpha(ByVal x As Long, ByVal y As Long, ByVal d As Long)
Dim k As Long
k = (TheClr.rgbReserved * 1024&) \ 255&
With TheArray(x And (TheBitmap.Width - 1), y And (TheBitmap.Height - 1))
 .rgbBlue = .rgbBlue + ((-.rgbBlue + TheClr.rgbBlue) * k) \ 1024&
 .rgbGreen = .rgbGreen + ((-.rgbGreen + TheClr.rgbGreen) * k) \ 1024&
 .rgbRed = .rgbRed + ((-.rgbRed + TheClr.rgbRed) * k) \ 1024&
 .rgbReserved = .rgbReserved + ((-.rgbReserved + TheClr.rgbReserved) * k) \ 1024&
End With
End Sub
#End If

Public Sub CalcLSystemTest(bmOut As typeAlphaDibSectionDescriptor, bProps() As Byte, sProps() As String)
Dim bHasArg As Boolean, bIsZeroL As Boolean
Dim s As String, v As Variant
Dim i As Long, j As Long, k As Long, m As Long
'init array
s = Replace(sProps(1), vbCrLf, vbLf)
s = Replace(s, vbCr, vbLf)
v = Split(s, vbLf)
m = UBound(v)
'check type
bHasArg = InStr(1, s, "(") > 0 And InStr(1, s, ")") > 0
bIsZeroL = True
For i = 0 To m
 s = v(i)
 j = InStr(1, s, ":")
 If j = 0 Then j = InStr(1, s, "=>")
 If j > 0 Then
  k = InStr(1, s, "<")
  If k < j And k > 0 Then
   bIsZeroL = False
   Exit For
  End If
  k = InStr(1, s, ">")
  If k < j And k > 0 Then
   bIsZeroL = False
   Exit For
  End If
 End If
Next i
'calc
If bIsZeroL Then
 If bHasArg Then
  pCalcArgZeroLTest bmOut, bProps, sProps(0), v
 Else
  pCalcZeroLTest bmOut, bProps, sProps(0), v
 End If
Else
 If bHasArg Then
  pCalcArgTwoLTest bmOut, bProps, sProps(0), v
 Else
  pCalcTwoLTest bmOut, bProps, sProps(0), v
 End If
End If
End Sub

'simplest
Private Sub pCalcZeroLTest(bmOut As typeAlphaDibSectionDescriptor, bProps() As Byte, sProps As String, v As Variant)
Dim TheClrTable() As RGBQUAD, TheClrCount As Long
Dim nCount As Long
Dim fX As Single, fY As Single, FS As Single
Dim d As Single, l As Single
Dim bHQ As Boolean
Dim nXCount As Long, nYCount As Long
Dim nXNow As Long, nYNow As Long
Dim TheSeed As Long
Dim TheSeed2(255) As Long
Dim ls As typeZeroL
'///
Dim nCur As Long 'current generation
Dim nBranch As Long 'current branch
Dim nDeletedBranch As Long
Dim tStack() As typeLSystemRecursive 'call stack
Dim tState() As typeLSystemState 'state stack
Dim curState As typeLSystemState
Dim tPolyPt() As typeLSystemPoint 'polygon point stack
Dim tPoly() As Long 'polygon point count stack
Dim nPolyPt As Long 'polygon point
Dim nPoly As Long 'polygon index
'///
Dim i As Long, j As Long
Dim op As typeZeroL_Operation
Dim idx As Long, lp As Long
Dim f As Long
Dim hbr As Long, hpn As Long
'///
'get data
nXCount = bProps(0)
nYCount = (nXCount And 224&) \ 32&
bHQ = (nXCount And 1&)
nXCount = (nXCount And 28&) \ 4&
pGetSeed bProps, 1, TheSeed
CopyMemory fX, bProps(3), 4&
CopyMemory fY, bProps(7), 4&
CopyMemory FS, bProps(11), 4&
CopyMemory d, bProps(15), 4&
CopyMemory l, bProps(19), 4&
nCount = bProps(23)
'convert data
d = d * 二π
FS = FS * 二π
l = l * bmOut.Width  '???
'get color
pGetColor sProps, TheClrTable, TheClrCount
'//////////////////////stupid algorithm version 0.00
'test input is:"F"+vbcrlf+"F=>F+F--F+F" = koch
'compile?
pCompile_ZeroL v, ls
If ls.nCount = 0 Or nXCount = 0 Or nYCount = 0 Then Exit Sub
'init array
Dim ox As Long, oy As Long
TheBitmap = bmOut
With TheSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = bmOut.Height
 .Bounds(1).cElements = bmOut.Width
 .pvData = bmOut.lpbm
End With
CopyMemory ByVal VarPtrArray(TheArray()), VarPtr(TheSA), 4&
'init array
ReDim tStack(255), tState(255), tPolyPt(4095), tPoly(255)
'init color cache
TheClrIndex = &H80000000
i = -1
CopyMemory TheClr, i, 4&
'start draw
For nXNow = 0 To nXCount - 1
 For nYNow = 0 To nYCount - 1
  nCur = 0
  nBranch = 0
  nDeletedBranch = &H7FFFFFFF
  nPolyPt = 0
  nPoly = -1
  With curState
   .x = fX + nXNow / nXCount
   If .x > 1 Then .x = .x - 1
   .x = .x * bmOut.Width
   .y = fY + nYNow / nYCount
   If .y > 1 Then .y = .y - 1
   .y = .y * bmOut.Height
   .FS = FS
   .idxClr = 0
   .idxWidth = 0
  End With
  idx = 1
  lp = 1
  'run
  Do
   If lp > ls.d(idx).nCount Then
    nCur = nCur - 1
    If nCur < 0 Then Exit Do
    'pop
    With tStack(nCur)
     idx = .nIndex
     lp = .nPos
    End With
   Else
    'find expand
    j = 0
    op = ls.d(idx).d(lp)
    If nCur < nCount Then
     f = -1
     With ls
      i = .nTable(op.nType, op.nIndex)
      Do Until i = 0
       With .d(i)
        If .nFreq < 32768 And f < 0 Then
         'get random
         TheSeed2(nCur) = TheSeed2(nCur) + 1
         f = cUnk.fRnd2(nCur, TheSeed2(nCur), TheSeed) And &H7FFF&
        End If
        f = f - .nFreq
       End With
       If f < 0 Then
        j = i
        Exit Do
       End If
       i = .d(i).idxNext
      Loop
     End With
    End If
    lp = lp + 1
    If j > 0 Then
     'push
     With tStack(nCur)
      .nIndex = idx
      .nPos = lp
     End With
     idx = j
     lp = 1
     nCur = nCur + 1
    Else 'draw
     Select Case op.nType
     Case 70, 71, 102, 103 'F,G,f,g
      If nBranch < nDeletedBranch Then
       With curState
        If op.nType < 80 Then
         '////////get color!!!
         If .idxClr <> TheClrIndex And TheClrCount > 0 Then
          TheClrIndex = .idxClr
          If TheClrIndex < 0 Then
           TheClr = TheClrTable(0)
          ElseIf TheClrIndex >= TheClrCount Then
           TheClr = TheClrTable(TheClrCount - 1)
          Else
           TheClr = TheClrTable(TheClrIndex)
          End If
         End If
         #If UseLineDDA Then
         #Else
         If TheClrCount > 0 Then
          With TheClr
           i = .rgbRed + .rgbGreen * &H100& + .rgbBlue * &H10000
          End With
         Else
          i = vbWhite
         End If
         #End If
         '////////
         #If UseLineDDA Then
         ox = .x
         oy = .y
         #Else
         MoveToEx bmOut.hdc, .x, .y, ByVal 0
         #End If
        End If
        .x = .x + l * Cos(.FS)
        .y = .y + l * Sin(.FS)
        If op.nType < 80 Then
         #If UseLineDDA Then
         If TheClr.rgbReserved > 0 Then
          If TheClr.rgbReserved = 255 Then
           LineDDA ox, oy, .x, .y, AddressOf pDrawLine, 0
          Else
           LineDDA ox, oy, .x, .y, AddressOf pDrawLineAlpha, 0
          End If
         End If
         #Else
         'ERROR!!! API didn't support 32-bit bitmap
         hpn = SelectObject(bmOut.hdc, CreatePen(0, 1, i))
         LineTo bmOut.hdc, .x, .y
         DeleteObject SelectObject(bmOut.hdc, hpn)
         #End If
        End If
        'record vertex
        If nPoly >= 0 And (op.nType And 1) = 0 Then
         tPoly(nPoly) = tPoly(nPoly) + 1
         With tPolyPt(nPolyPt)
          .x = curState.x
          .y = curState.y
         End With
         nPolyPt = nPolyPt + 1
        End If
       End With
      End If
     Case 46 '.
      'record vertex
      If nBranch < nDeletedBranch And nPoly >= 0 Then
       tPoly(nPoly) = tPoly(nPoly) + 1
       With tPolyPt(nPolyPt)
        .x = curState.x
        .y = curState.y
       End With
       nPolyPt = nPolyPt + 1
      End If
     Case 91 '[
      tState(nBranch) = curState
      nBranch = nBranch + 1
     Case 93 ']
      If nBranch > 0 Then nBranch = nBranch - 1
      If nBranch < nDeletedBranch Then nDeletedBranch = &H7FFFFFFF
      curState = tState(nBranch)
     Case 123 '{
      nPoly = nPoly + 1
      tPoly(nPoly) = 0
     Case 125 '}
      If nPoly >= 0 Then
       j = tPoly(nPoly)
       nPoly = nPoly - 1
       If j > 0 Then
        With curState
         '////////get color!!!
         If .idxClr <> TheClrIndex And TheClrCount > 0 Then
          TheClrIndex = .idxClr
          If TheClrIndex < 0 Then
           TheClr = TheClrTable(0)
          ElseIf TheClrIndex >= TheClrCount Then
           TheClr = TheClrTable(TheClrCount - 1)
          Else
           TheClr = TheClrTable(TheClrIndex)
          End If
         End If
         #If UseLineDDA Then
         #Else
         If TheClrCount > 0 Then
          With TheClr
           i = .rgbRed + .rgbGreen * &H100& + .rgbBlue * &H10000
          End With
         Else
          i = vbWhite
         End If
         #End If
         '////////
        End With
        nPolyPt = nPolyPt - j
        If nBranch < nDeletedBranch Then
         #If UseLineDDA Then
         pDrawPolygon tPolyPt, nPolyPt, j
         #Else
         pDrawPolygonTest tPolyPt, nPolyPt, j, i
         #End If
        End If
       End If
      End If
     Case 43 '+
      If nBranch < nDeletedBranch Then
       With curState
        .FS = .FS - d
        If .FS < 0 Then .FS = .FS + 二π
       End With
      End If
     Case 45 '-
      If nBranch < nDeletedBranch Then
       With curState
        .FS = .FS + d
        If .FS > 二π Then .FS = .FS - 二π
       End With
      End If
     Case 124 '|
      If nBranch < nDeletedBranch Then
       With curState
        If .FS < π Then .FS = .FS + π Else .FS = .FS - π
       End With
      End If
     Case 33 '!
      If nBranch < nDeletedBranch Then
       With curState
        .idxWidth = .idxWidth + 1 '??? TODO:
       End With
      End If
     Case 39 ''
      If nBranch < nDeletedBranch Then
       With curState
        .idxClr = .idxClr + 1 '??? TODO:
       End With
      End If
     Case 37 '%
      nDeletedBranch = nBranch
     End Select
    End If
   End If
  Loop
 Next nYNow
Next nXNow
'destroy
ZeroMemory ByVal VarPtrArray(TheArray()), 4&
End Sub

Private Sub pCalcArgZeroLTest(bmOut As typeAlphaDibSectionDescriptor, bProps() As Byte, sProps As String, v As Variant)
Dim TheClrTable() As RGBQUAD, TheClrCount As Long
Dim nCount As Long
Dim fX As Single, fY As Single, FS As Single
Dim bHQ As Boolean
Dim nXCount As Long, nYCount As Long
Dim nXNow As Long, nYNow As Long
Dim TheSeed As Long
'Dim TheSeed2(255) As Long
Dim ls As typeArgZeroL
'///
Dim nCur As Long 'current generation
Dim nBranch As Long 'current branch
Dim nDeletedBranch As Long
Dim tStack() As typeLSystemRecursive 'call stack
Dim tStack2() As typeArgZeroL_Operation 'instance of operation
Dim tState() As typeLSystemState 'state stack
Dim curState As typeLSystemState
Dim tPolyPt() As typeLSystemPoint 'polygon point stack
Dim tPoly() As Long 'polygon point count stack
Dim nPolyPt As Long 'polygon point
Dim nPoly As Long 'polygon index
'///
Dim i As Long, j As Long
Dim op As typeArgZeroL_Operation
Dim idx As Long, lp As Long, lpe As Long
Dim f As Long
Dim hbr As Long, hpn As Long
'///
'get data
nXCount = bProps(0)
nYCount = (nXCount And 224&) \ 32&
bHQ = (nXCount And 1&)
nXCount = (nXCount And 28&) \ 4&
pGetSeed bProps, 1, TheSeed
CopyMemory fX, bProps(3), 4&
CopyMemory fY, bProps(7), 4&
CopyMemory FS, bProps(11), 4&
CopyMemory ls.fConst(0), bProps(15), 8&
nCount = bProps(23)
'convert data
FS = FS * 二π
ls.fConst(0) = ls.fConst(0) * 二π
ls.fConst(1) = ls.fConst(1) * bmOut.Width
ls.fConst(2) = π
'get color
pGetColor sProps, TheClrTable, TheClrCount
'//////////////////////stupid algorithm version 0.00
'compile?
pCompile_ArgZeroL v, ls
If ls.nCount = 0 Or nXCount = 0 Or nYCount = 0 Then Exit Sub
'init array
Dim ox As Long, oy As Long
TheBitmap = bmOut
With TheSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = bmOut.Height
 .Bounds(1).cElements = bmOut.Width
 .pvData = bmOut.lpbm
End With
CopyMemory ByVal VarPtrArray(TheArray()), VarPtr(TheSA), 4&
'init array
ReDim tStack(255), tStack2(4095), tState(255), tPolyPt(4095), tPoly(255)
'init color cache
TheClrIndex = &H80000000
i = -1
CopyMemory TheClr, i, 4&
'start draw
For nXNow = 0 To nXCount - 1
 For nYNow = 0 To nYCount - 1
  nCur = 0
  nBranch = 0
  nDeletedBranch = &H7FFFFFFF
  nPolyPt = 0
  nPoly = -1
  With curState
   .x = fX + nXNow / nXCount
   If .x > 1 Then .x = .x - 1
   .x = .x * bmOut.Width
   .y = fY + nYNow / nYCount
   If .y > 1 Then .y = .y - 1
   .y = .y * bmOut.Height
   .FS = FS
   .idxClr = 0
   .idxWidth = 0
  End With
  idx = 1
  lp = 0
  lpe = ls.d(1).nCount
  'init axiom
  Erase op.fArg
  pArgZeroL_Apply tStack2, 0, ls.d(1).nCount, ls.d(1).d, ls.fConst, op.fArg
  'run
  Do
   If lp >= lpe Then
    nCur = nCur - 1
    If nCur < 0 Then Exit Do
    'pop
    With tStack(nCur)
     idx = .nIndex
     lp = .nPos
     lpe = .nReserved
    End With
   Else
    'find expand
    j = 0
    op = tStack2(lp)
    If nCur < nCount Then
     'TODO:probability
     With ls
      i = .nTable(op.nType, op.nIndex)
      Do Until i = 0
       With .d(i)
        If .nArgCount = op.nArgCount Then
         If .tExp.nCount > 0 Then
          If pCalc(.tExp, ls.fConst, op.fArg) <> 0 Then j = i
         Else
          j = i
         End If
        End If
       End With
       If j > 0 Then Exit Do
       i = .d(i).idxNext
      Loop
     End With
    End If
    lp = lp + 1
    If j > 0 Then
     'push
     With tStack(nCur)
      .nIndex = idx
      .nPos = lp
      .nReserved = lpe
     End With
     idx = j
     lp = lpe
     With ls.d(j)
      pArgZeroL_Apply tStack2, lpe, .nCount, .d, ls.fConst, op.fArg
      lpe = lpe + .nCount
     End With
     nCur = nCur + 1
    Else 'draw
     Select Case op.nType
     Case 70, 71, 102, 103 'F,G,f,g
      If nBranch < nDeletedBranch Then
       With curState
        If op.nType < 80 Then
         '////////get color!!!
         If .idxClr <> TheClrIndex And TheClrCount > 0 Then
          TheClrIndex = .idxClr
          If TheClrIndex < 0 Then
           TheClr = TheClrTable(0)
          ElseIf TheClrIndex >= TheClrCount Then
           TheClr = TheClrTable(TheClrCount - 1)
          Else
           TheClr = TheClrTable(TheClrIndex)
          End If
         End If
         #If UseLineDDA Then
         #Else
         If TheClrCount > 0 Then
          With TheClr
           i = .rgbRed + .rgbGreen * &H100& + .rgbBlue * &H10000
          End With
         Else
          i = vbWhite
         End If
         #End If
         '////////
         #If UseLineDDA Then
         ox = .x
         oy = .y
         #Else
         MoveToEx bmOut.hdc, .x, .y, ByVal 0
         #End If
        End If
        If op.nArgCount > 0 Then
         .x = .x + op.fArg(0) * Cos(.FS)
         .y = .y + op.fArg(0) * Sin(.FS)
        Else
         .x = .x + ls.fConst(1) * Cos(.FS)
         .y = .y + ls.fConst(1) * Sin(.FS)
        End If
        If op.nType < 80 Then
         #If UseLineDDA Then
         If TheClr.rgbReserved > 0 Then
          If TheClr.rgbReserved = 255 Then
           LineDDA ox, oy, .x, .y, AddressOf pDrawLine, 0
          Else
           LineDDA ox, oy, .x, .y, AddressOf pDrawLineAlpha, 0
          End If
         End If
         #Else
         'ERROR!!! API didn't support 32-bit bitmap
         hpn = SelectObject(bmOut.hdc, CreatePen(0, 1, i))
         LineTo bmOut.hdc, .x, .y
         DeleteObject SelectObject(bmOut.hdc, hpn)
         #End If
        End If
        'record vertex
        If nPoly >= 0 And (op.nType And 1) = 0 Then
         tPoly(nPoly) = tPoly(nPoly) + 1
         With tPolyPt(nPolyPt)
          .x = curState.x
          .y = curState.y
         End With
         nPolyPt = nPolyPt + 1
        End If
       End With
      End If
     Case 46 '.
      'record vertex
      If nBranch < nDeletedBranch And nPoly >= 0 Then
       tPoly(nPoly) = tPoly(nPoly) + 1
       With tPolyPt(nPolyPt)
        .x = curState.x
        .y = curState.y
       End With
       nPolyPt = nPolyPt + 1
      End If
     Case 91 '[
      tState(nBranch) = curState
      nBranch = nBranch + 1
     Case 93 ']
      If nBranch > 0 Then nBranch = nBranch - 1
      If nBranch < nDeletedBranch Then nDeletedBranch = &H7FFFFFFF
      curState = tState(nBranch)
     Case 123 '{
      nPoly = nPoly + 1
      tPoly(nPoly) = 0
     Case 125 '}
      If nPoly >= 0 Then
       j = tPoly(nPoly)
       nPoly = nPoly - 1
       If j > 0 Then
        With curState
         '////////get color!!!
         If .idxClr <> TheClrIndex And TheClrCount > 0 Then
          TheClrIndex = .idxClr
          If TheClrIndex < 0 Then
           TheClr = TheClrTable(0)
          ElseIf TheClrIndex >= TheClrCount Then
           TheClr = TheClrTable(TheClrCount - 1)
          Else
           TheClr = TheClrTable(TheClrIndex)
          End If
         End If
         #If UseLineDDA Then
         #Else
         If TheClrCount > 0 Then
          With TheClr
           i = .rgbRed + .rgbGreen * &H100& + .rgbBlue * &H10000
          End With
         Else
          i = vbWhite
         End If
         #End If
         '////////
        End With
        nPolyPt = nPolyPt - j
        If nBranch < nDeletedBranch Then
         #If UseLineDDA Then
         pDrawPolygon tPolyPt, nPolyPt, j
         #Else
         pDrawPolygonTest tPolyPt, nPolyPt, j, i
         #End If
        End If
       End If
      End If
     Case 43 '+
      If nBranch < nDeletedBranch Then
       With curState
        If op.nArgCount > 0 Then
         .FS = .FS - op.fArg(0) '??? rad?? deg??
        Else
         .FS = .FS - ls.fConst(0)
        End If
        If .FS < 0 Then .FS = .FS + 二π
       End With
      End If
     Case 45 '-
      If nBranch < nDeletedBranch Then
       With curState
        If op.nArgCount > 0 Then
         .FS = .FS + op.fArg(0) '??? rad?? deg??
        Else
         .FS = .FS + ls.fConst(0)
        End If
        If .FS > 二π Then .FS = .FS - 二π
       End With
      End If
     Case 124 '|
      If nBranch < nDeletedBranch Then
       With curState
        If .FS < π Then .FS = .FS + π Else .FS = .FS - π
       End With
      End If
     Case 33 '!
      If nBranch < nDeletedBranch Then
       With curState
        If op.nArgCount > 0 Then
         .idxWidth = .idxWidth + op.fArg(0)
        Else
         .idxWidth = .idxWidth + 1 '??? TODO:
        End If
       End With
      End If
     Case 39 ''
      If nBranch < nDeletedBranch Then
       With curState
        If op.nArgCount > 0 Then
         .idxClr = .idxClr + op.fArg(0)
        Else
         .idxClr = .idxClr + 1 '??? TODO:
        End If
       End With
      End If
     Case 37 '%
      nDeletedBranch = nBranch
     End Select
    End If
   End If
  Loop
 Next nYNow
Next nXNow
'destroy
ZeroMemory ByVal VarPtrArray(TheArray()), 4&
End Sub

#If UseLineDDA Then

'no API!!!
'TODO:use edge-chain linked-list may be faster
Private Sub pDrawPolygon(tPolyPt() As typeLSystemPoint, ByVal nStart As Long, ByVal nCount As Long)
Dim p() As POINTAPI '0-based
Dim TheEdge() As typeDrawPolygonEdge '1-based!!
Dim TheEdgeCount As Long
Dim idxHead As Long 'linked-list
Dim tmp As typeDrawPolygonEdge
Dim i As Long, j As Long, k As Long
Dim ii As Long, jj As Long
Dim nTop As Long, nBottom As Long
Dim x As Long, y As Long
Dim xx As Long, yy As Long
If TheClr.rgbReserved = 0 Then Exit Sub 'don't draw :-3
'init point
nTop = &H7FFFFFFF
nBottom = &H80000000
ReDim p(nCount - 1)
For i = 0 To nCount - 1
 With tPolyPt(nStart + i)
  p(i).x = .x
  j = .y
  p(i).y = j
  If j < nTop Then nTop = j
  If j > nBottom Then nBottom = j
 End With
Next i
'init edge
ReDim TheEdge(1 To nCount)
For i = 0 To nCount - 2
 With p(i + 1)
  k = .y - p(i).y
  If k > 0 Then
   TheEdgeCount = TheEdgeCount + 1
   With TheEdge(TheEdgeCount)
    .idx1 = i
    .idx2 = i + 1
    .dy = k
   End With
  ElseIf k < 0 Then
   TheEdgeCount = TheEdgeCount + 1
   With TheEdge(TheEdgeCount)
    .idx1 = i + 1
    .idx2 = i
    .dy = -k
   End With
  End If
 End With
Next i
With p(0)
 k = .y - p(i).y
 If k > 0 Then
  TheEdgeCount = TheEdgeCount + 1
  With TheEdge(TheEdgeCount)
   .idx1 = i
   .idx2 = 0
   .dy = k
  End With
 ElseIf k < 0 Then
  TheEdgeCount = TheEdgeCount + 1
  With TheEdge(TheEdgeCount)
   .idx1 = 0
   .idx2 = i
   .dy = -k
  End With
 End If
End With
'sort using insert-sort
For i = 2 To TheEdgeCount
 k = p(TheEdge(i).idx1).y
 If k < p(TheEdge(i - 1).idx1).y Then
  tmp = TheEdge(i)
  TheEdge(i) = TheEdge(i - 1)
  j = i - 2
  Do Until j < 1
   If k >= p(TheEdge(j).idx1).y Then Exit Do
   TheEdge(j + 1) = TheEdge(j)
   j = j - 1
  Loop
  TheEdge(j + 1) = tmp
 End If
Next i
'start draw!!!
k = (TheClr.rgbReserved * 1024&) \ 255& 'transparency
yy = 1 'next edge
For y = nTop To nBottom
 '////////remove item?
 i = idxHead
 Do Until i = 0
  If p(TheEdge(i).idx2).y <= y Then 'remove
   If i = idxHead Then
    idxHead = TheEdge(i).idxNext
   Else
    TheEdge(j).idxNext = TheEdge(i).idxNext
   End If
  Else
   j = i
  End If
  i = TheEdge(i).idxNext
 Loop
 '////////update position
 i = idxHead
 Do Until i = 0
  With TheEdge(i)
   .x = .x + .BigDelta
   .eps = .eps + .SmallDelta
   If .eps >= .dy Then
    .x = .x + 1
    .eps = .eps - .dy
   End If
  End With
  i = TheEdge(i).idxNext
 Loop
 '////////sort again
 If idxHead > 0 Then
  i = idxHead
  Do
   j = TheEdge(i).idxNext
   If j = 0 Then Exit Do
   xx = TheEdge(j).x
   If TheEdge(i).x > xx Then 'move it
    ii = idxHead
    Do Until TheEdge(ii).x >= xx
     jj = ii
     ii = TheEdge(ii).idxNext
     Debug.Assert ii > 0
    Loop
    TheEdge(i).idxNext = TheEdge(j).idxNext
    If ii = idxHead Then 'change header
     TheEdge(j).idxNext = idxHead
     idxHead = j
    Else
     TheEdge(j).idxNext = TheEdge(jj).idxNext
     TheEdge(jj).idxNext = j
    End If
   Else
    i = j
   End If
  Loop
 End If
 '////////add item?
 Do Until yy > TheEdgeCount
  If p(TheEdge(yy).idx1).y > y Then Exit Do
  With TheEdge(yy)
   'init
   .x = p(.idx1).x
   .eps = .dy \ 2
   i = p(.idx2).x - .x
   .BigDelta = i \ .dy
   .SmallDelta = i Mod .dy
   If .SmallDelta < 0 Then
    .BigDelta = .BigDelta - 1
    .SmallDelta = .SmallDelta + .dy
   End If
   'insert and sort
   If idxHead = 0 Then
    idxHead = yy
   ElseIf .x <= TheEdge(idxHead).x Then
    .idxNext = idxHead
    idxHead = yy
   Else
    jj = idxHead
    Do
     ii = TheEdge(jj).idxNext
     If ii = 0 Then Exit Do
     If .x <= TheEdge(ii).x Then Exit Do
     jj = ii
    Loop
    .idxNext = ii
    TheEdge(jj).idxNext = yy
   End If
  End With
  'next
  yy = yy + 1
 Loop
 '////////draw scanline
 i = idxHead
 Do Until i = 0
  j = TheEdge(i).idxNext
  If j = 0 Then
   'error!!!
   Debug.Assert False
   Exit Do
  Else
   ii = TheEdge(i).x
   jj = TheEdge(j).x
   'draw!!
   For x = ii To jj
    If TheClr.rgbReserved = 255 Then
     TheArray(x And (TheBitmap.Width - 1), y And (TheBitmap.Height - 1)) = TheClr
    Else
     With TheArray(x And (TheBitmap.Width - 1), y And (TheBitmap.Height - 1))
      .rgbBlue = .rgbBlue + ((-.rgbBlue + TheClr.rgbBlue) * k) \ 1024&
      .rgbGreen = .rgbGreen + ((-.rgbGreen + TheClr.rgbGreen) * k) \ 1024&
      .rgbRed = .rgbRed + ((-.rgbRed + TheClr.rgbRed) * k) \ 1024&
      .rgbReserved = .rgbReserved + ((-.rgbReserved + TheClr.rgbReserved) * k) \ 1024&
     End With
    End If
   Next x
   i = TheEdge(j).idxNext
  End If
 Loop
Next y
End Sub

#Else

'uses API :-3
Private Sub pDrawPolygonTest(tPolyPt() As typeLSystemPoint, ByVal nStart As Long, ByVal nCount As Long, ByVal clr As Long)
Dim p() As POINTAPI
Dim i As Long, hbr As Long
'start
ReDim p(nCount - 1)
For i = 0 To nCount - 1
 With tPolyPt(nStart + i)
  p(i).x = .x
  p(i).y = .y
 End With
Next i
'draw
i = SelectObject(TheBitmap.hdc, GetStockObject(NULL_PEN))
hbr = SelectObject(TheBitmap.hdc, CreateSolidBrush(clr))
Polygon TheBitmap.hdc, p(0), nCount
Call SelectObject(TheBitmap.hdc, i) 'don't delete!
DeleteObject SelectObject(TheBitmap.hdc, hbr)
End Sub

#End If

Private Sub pGetColor(ByRef s As String, TheClrTable() As RGBQUAD, TheClrCount As Long)
Dim i As Long, j As Long, k As Long, m As Long
Dim a1 As Long, r1 As Long, g1 As Long, b1 As Long
Dim a2 As Long, r2 As Long, g2 As Long, b2 As Long
Dim m2 As Long
Dim clrs() As RGBQUAD
m = LenB(s) \ 5
If m > 0 Then
 ReDim clrs(0 To 1, 1 To m)
 For i = 1 To m
  CopyMemory clrs(0, i), ByVal StrPtr(s) + i * 5 - 5, 5&
  j = j + 1 + clrs(1, i).rgbBlue
 Next i
 TheClrCount = j
 ReDim TheClrTable(j - 1)
 j = 0
 For i = 1 To m
  m2 = clrs(1, i).rgbBlue + 1&
  If m2 <= 1 Or i = m Then
   For k = j To j + m2 - 1
    TheClrTable(k) = clrs(0, i)
   Next k
  Else
   TheClrTable(j) = clrs(0, i)
   With clrs(0, i)
    b1 = .rgbBlue
    g1 = .rgbGreen
    r1 = .rgbRed
    a1 = .rgbReserved
   End With
   With clrs(0, i + 1)
    b2 = .rgbBlue - b1
    g2 = .rgbGreen - g1
    r2 = .rgbRed - r1
    a2 = .rgbReserved - a1
   End With
   For k = 1 To m2 - 1
    With TheClrTable(j + k)
     .rgbBlue = b1 + (b2 * k) \ m2
     .rgbGreen = g1 + (g2 * k) \ m2
     .rgbRed = r1 + (r2 * k) \ m2
     .rgbReserved = a1 + (a2 * k) \ m2
    End With
   Next k
  End If
  j = j + m2
 Next i
Else
 Erase TheClrTable
 TheClrCount = 0
End If
End Sub

Private Sub pCompile_ZeroL(v As Variant, ls As typeZeroL)
On Error GoTo a
Dim s As String, m As Long
Dim nIdx As Long
Dim i As Long, j As Long, lps As Long
Dim f As Long
'init
With ls
 .nCount = 0
 Erase .d
 ReDim .nTable(255, 255)
 'get
 m = UBound(v)
 For nIdx = 0 To m
  s = Trim(v(nIdx))
  If s = "" Then
   'comment only
  ElseIf Left(s, 1) = "#" Then
   'TODO:#include
  Else
   .nCount = .nCount + 1
   ReDim Preserve .d(1 To .nCount)
   'is rule?
   s = Replace(s, " ", "")
   lps = InStr(1, s, "=>")
   If lps > 0 Then
    'get freq
    i = InStr(2, s, "@")
    If i > 0 And i < lps Then
     .d(.nCount).nFreq = Val(Mid(s, i + 1)) * 32768#
    Else
     .d(.nCount).nFreq = 32768
    End If
    'get type
    i = AscB(Left(s, 1))
    If Mid(s, 2, 1) = "_" Then
     j = AscB(Mid(s, 3, 1))
    Else
     j = 0
    End If
    'process linked-list
    .d(.nCount).idxNext = .nTable(i, j)
    .nTable(i, j) = .nCount
    'over
    s = Mid(s, lps + 2)
   End If
   pCompile_ZeroL_Rule s, .d(.nCount).d, .d(.nCount).nCount
  End If
 Next nIdx
 'process freq
 For i = 0 To 255
  For j = 0 To 255
   lps = .nTable(i, j)
   If lps > 0 Then
    f = 0
    Do Until lps = 0
     f = f + .d(lps).nFreq
     lps = .d(lps).idxNext
    Loop
    If f > 32768 Then
     lps = .nTable(i, j)
     Do Until lps = 0
      .d(lps).nFreq = (.d(lps).nFreq * 32768 + 16383&) \ f
      lps = .d(lps).idxNext
     Loop
    End If
   End If
  Next j
 Next i
End With
Exit Sub
a:
ls.nCount = 0
Erase ls.d, ls.nTable
End Sub

Private Sub pCompile_ZeroL_Rule(ByVal s As String, d() As typeZeroL_Operation, nCount As Long)
Dim lps As Long, m As Long
Dim i As Long, j As Long
m = Len(s)
lps = 1
Do While lps <= m
 i = AscB(Mid(s, lps, 1))
 If Mid(s, lps + 1, 1) = "_" And lps + 2 <= m Then
  j = AscB(Mid(s, lps + 2, 1))
  lps = lps + 3
 Else
  j = 0
  lps = lps + 1
 End If
 nCount = nCount + 1
 ReDim Preserve d(1 To nCount)
 With d(nCount)
  .nType = i
  .nIndex = j
 End With
Loop
End Sub

Private Sub pCompile_ArgZeroL(v As Variant, ls As typeArgZeroL)
On Error GoTo a
Dim s As String, m As Long
Dim nIdx As Long
Dim i As Long, j As Long, lps As Long
Dim lp As Long
Dim f As Long
Dim tConst As typeSymbolTable
Dim tArg As typeSymbolTables
'init
With tConst
 .nCount = 3 'stupid??
 ReDim .s(2)
 .s(0) = "_a"
 .s(1) = "_l"
 .s(2) = "pi" 'stupid??
End With
tArg.nCount = 1
ReDim tArg.d(0)
With ls
 .nCount = 0
 Erase .d
 ReDim .nTable(255, 255)
 'get
 m = UBound(v)
 For nIdx = 0 To m
  s = Trim(v(nIdx))
  If s = "" Then
   'comment only
  ElseIf Left(s, 1) = "#" Then
   'delete space :-3
   i = Len(s)
   Do
    s = Replace(s, "  ", " ")
    j = Len(s)
    If i = j Then Exit Do
    i = j
   Loop
   'find space
   i = InStr(1, s, " ")
   If i > 0 Then
    Select Case Mid(s, 2, i - 2)
    Case "define"
     'find space again
     j = InStr(i + 1, s, " ")
     If j > 0 Then
      'parse expression
      m_sExp = Replace(Mid(s, j + 1), " ", "")
      m_lps = 1
      m_lpm = Len(m_sExp)
      .fConst(tConst.nCount) = pInterprete_ArgZeroL_Expression(tConst, .fConst)
      With tConst
       ReDim Preserve .s(.nCount)
       .s(.nCount) = Mid(s, i + 1, j - i - 1)
       .nCount = .nCount + 1
      End With
     End If
    Case Else
     'TODO:#include
    End Select
   End If
  Else
   .nCount = .nCount + 1
   ReDim Preserve .d(1 To .nCount)
   'is rule?
   s = Replace(s, " ", "")
   lps = InStr(1, s, "=>")
   If lps > 0 Then
    'get type
    i = AscB(Left(s, 1))
    If Mid(s, 2, 1) = "_" Then
     j = AscB(Mid(s, 3, 1))
     lp = 4
    Else
     j = 0
     lp = 2
    End If
    'process linked-list
    .d(.nCount).idxNext = .nTable(i, j)
    .nTable(i, j) = .nCount
    'process arguments
    tArg.nCount = 1
    With tArg.d(0)
     .nCount = 0
     Erase .s
    End With
    If Mid(s, lp, 1) = "(" Then
     i = InStr(lp, s, ")")
     Do
      j = InStr(lp + 1, s, ",")
      If j = 0 Or j > i Then j = i
      With tArg.d(0)
       ReDim Preserve .s(.nCount)
       .s(.nCount) = Mid(s, lp + 1, j - lp - 1)
       .nCount = .nCount + 1
      End With
      lp = j
     Loop While j < i
    End If
    With .d(.nCount)
     .nArgCount = tArg.d(0).nCount
     .tExp.nCount = 0
     Erase .tExp.d
     If .nArgCount > 7 Then Err.Raise 9
    End With
    'process expression
    i = InStr(1, s, ":")
    If i > 0 Then
     m_sExp = Mid(s, i + 1, lps - i - 1)
     m_lps = 1
     m_lpm = Len(m_sExp)
     pCompile_ArgTwoL_Expression .d(.nCount).tExp, tConst, tArg, .fConst
    End If
    'over
    s = Mid(s, lps + 2)
   Else
    tArg.nCount = 0 'not a rule!
   End If
   m_sExp = s
   pCompile_ArgTwoL_Rule .d(.nCount).d, .d(.nCount).nCount, tConst, tArg, .fConst
  End If
 Next nIdx
End With
Exit Sub
a:
ls.nCount = 0
Erase ls.d, ls.nTable
End Sub

Private Sub pCompile_ArgTwoL_Rule(d() As typeArgZeroL_OperationD, nCount As Long, tConst As typeSymbolTable, tArg As typeSymbolTables, fConst() As Single)
Dim i As Long, j As Long
m_lps = 1
m_lpm = Len(m_sExp)
Do Until m_lps > m_lpm
 i = AscB(Mid(m_sExp, m_lps, 1))
 If Mid(m_sExp, m_lps + 1, 1) = "_" And m_lps + 2 <= m_lpm Then
  j = AscB(Mid(m_sExp, m_lps + 2, 1))
  m_lps = m_lps + 3
 Else
  j = 0
  m_lps = m_lps + 1
 End If
 nCount = nCount + 1
 ReDim Preserve d(1 To nCount)
 With d(nCount)
  .nType = i
  .nIndex = j
  .nArgCount = 0
  Erase .tExp
  'process arguments and expression
  If Mid(m_sExp, m_lps, 1) = "(" Then
   m_lps = m_lps + 1
   Do
    ReDim Preserve .tExp(.nArgCount)
    pCompile_ArgTwoL_Expression .tExp(.nArgCount), tConst, tArg, fConst
    .nArgCount = .nArgCount + 1
    Select Case Mid(m_sExp, m_lps, 1)
    Case ")"
     m_lps = m_lps + 1
     Exit Do
    Case ","
     m_lps = m_lps + 1
    Case Else
     Err.Raise 5
    End Select
   Loop
  End If
 End With
Loop
End Sub

'exp -> C | exp&C | exp|C
'C -> E | E>E | E<E | E=E | E>=E | E<=E | E<>E | !C
'E -> T | E+T | E-T
'T -> F | T*F | T/F | -T
'F -> G | F^G
'G -> <num> | <var> | (exp)

Private Sub pCompile_ArgTwoL_Expression(d As typeArgZeroL_Expression, tConst As typeSymbolTable, tArg As typeSymbolTables, fConst() As Single)
pCompile_ArgTwoL_ExpressionC d, tConst, tArg, fConst
Do Until m_lps > m_lpm
 Select Case Mid(m_sExp, m_lps, 1)
 Case "&"
  m_lps = m_lps + 1
  pCompile_ArgTwoL_ExpressionC d, tConst, tArg, fConst
  With d
   .nCount = .nCount + 1
   ReDim Preserve .d(1 To .nCount)
   .d(.nCount) = 10
  End With
 Case "|"
  m_lps = m_lps + 1
  pCompile_ArgTwoL_ExpressionC d, tConst, tArg, fConst
  With d
   .nCount = .nCount + 1
   ReDim Preserve .d(1 To .nCount)
   .d(.nCount) = 11
  End With
 Case Else
  Exit Do
 End Select
Loop
End Sub

Private Sub pCompile_ArgTwoL_ExpressionC(d As typeArgZeroL_Expression, tConst As typeSymbolTable, tArg As typeSymbolTables, fConst() As Single)
Dim b As Boolean
Do While Mid(m_sExp, m_lps, 1) = "!"
 b = Not b
 m_lps = m_lps + 1
Loop
pCompile_ArgTwoL_ExpressionE d, tConst, tArg, fConst
Select Case Mid(m_sExp, m_lps, 1)
Case "<"
 m_lps = m_lps + 1
 Select Case Mid(m_sExp, m_lps, 1)
 Case "=" '<=
  m_lps = m_lps + 1
  pCompile_ArgTwoL_ExpressionE d, tConst, tArg, fConst
  With d
   .nCount = .nCount + 1
   ReDim Preserve .d(1 To .nCount)
   .d(.nCount) = 14
  End With
 Case ">" '<>
  m_lps = m_lps + 1
  pCompile_ArgTwoL_ExpressionE d, tConst, tArg, fConst
  With d
   .nCount = .nCount + 1
   ReDim Preserve .d(1 To .nCount)
   .d(.nCount) = 16
  End With
 Case Else '<
  pCompile_ArgTwoL_ExpressionE d, tConst, tArg, fConst
  With d
   .nCount = .nCount + 1
   ReDim Preserve .d(1 To .nCount)
   .d(.nCount) = 7
  End With
 End Select
Case ">"
 m_lps = m_lps + 1
 If Mid(m_sExp, m_lps, 1) = "=" Then '>=
  m_lps = m_lps + 1
  pCompile_ArgTwoL_ExpressionE d, tConst, tArg, fConst
  With d
   .nCount = .nCount + 1
   ReDim Preserve .d(1 To .nCount)
   .d(.nCount) = 15
  End With
 Else
  pCompile_ArgTwoL_ExpressionE d, tConst, tArg, fConst
  With d
   .nCount = .nCount + 1
   ReDim Preserve .d(1 To .nCount)
   .d(.nCount) = 8
  End With
 End If
Case "="
 m_lps = m_lps + 1
 pCompile_ArgTwoL_ExpressionE d, tConst, tArg, fConst
 With d
  .nCount = .nCount + 1
  ReDim Preserve .d(1 To .nCount)
  .d(.nCount) = 9
 End With
End Select
If b Then
 With d
  .nCount = .nCount + 1
  ReDim Preserve .d(1 To .nCount)
  .d(.nCount) = &HD&
 End With
End If
End Sub

Private Sub pCompile_ArgTwoL_ExpressionE(d As typeArgZeroL_Expression, tConst As typeSymbolTable, tArg As typeSymbolTables, fConst() As Single)
pCompile_ArgTwoL_ExpressionT d, tConst, tArg, fConst
Do Until m_lps > m_lpm
 Select Case Mid(m_sExp, m_lps, 1)
 Case "+"
  m_lps = m_lps + 1
  pCompile_ArgTwoL_ExpressionT d, tConst, tArg, fConst
  With d
   .nCount = .nCount + 1
   ReDim Preserve .d(1 To .nCount)
   .d(.nCount) = 2
  End With
 Case "-"
  m_lps = m_lps + 1
  pCompile_ArgTwoL_ExpressionT d, tConst, tArg, fConst
  With d
   .nCount = .nCount + 1
   ReDim Preserve .d(1 To .nCount)
   .d(.nCount) = 3
  End With
 Case Else
  Exit Do
 End Select
Loop
End Sub

Private Sub pCompile_ArgTwoL_ExpressionT(d As typeArgZeroL_Expression, tConst As typeSymbolTable, tArg As typeSymbolTables, fConst() As Single)
Dim b As Boolean
Do While Mid(m_sExp, m_lps, 1) = "-"
 b = Not b
 m_lps = m_lps + 1
Loop
pCompile_ArgTwoL_ExpressionF d, tConst, tArg, fConst
Do Until m_lps > m_lpm
 Select Case Mid(m_sExp, m_lps, 1)
 Case "*"
  m_lps = m_lps + 1
  pCompile_ArgTwoL_ExpressionF d, tConst, tArg, fConst
  With d
   .nCount = .nCount + 1
   ReDim Preserve .d(1 To .nCount)
   .d(.nCount) = 4
  End With
 Case "/"
  m_lps = m_lps + 1
  pCompile_ArgTwoL_ExpressionF d, tConst, tArg, fConst
  With d
   .nCount = .nCount + 1
   ReDim Preserve .d(1 To .nCount)
   .d(.nCount) = 5
  End With
 Case Else
  Exit Do
 End Select
Loop
If b Then
 With d
  .nCount = .nCount + 1
  ReDim Preserve .d(1 To .nCount)
  .d(.nCount) = &HC&
 End With
End If
End Sub

Private Sub pCompile_ArgTwoL_ExpressionF(d As typeArgZeroL_Expression, tConst As typeSymbolTable, tArg As typeSymbolTables, fConst() As Single)
pCompile_ArgTwoL_ExpressionG d, tConst, tArg, fConst
Do Until m_lps > m_lpm
 Select Case Mid(m_sExp, m_lps, 1)
 Case "^"
  m_lps = m_lps + 1
  pCompile_ArgTwoL_ExpressionG d, tConst, tArg, fConst
  With d
   .nCount = .nCount + 1
   ReDim Preserve .d(1 To .nCount)
   .d(.nCount) = 6
  End With
 Case Else
  Exit Do
 End Select
Loop
End Sub

Private Sub pCompile_ArgTwoL_ExpressionG(d As typeArgZeroL_Expression, tConst As typeSymbolTable, tArg As typeSymbolTables, fConst() As Single)
Dim i As Long, s As String
Dim bHasExp As Boolean
Dim bHasPoint As Boolean
Dim bSign As Boolean
Dim f As Single
i = Asc(Mid(m_sExp, m_lps, 1))
Select Case i
Case 40 '"("
 m_lps = m_lps + 1
 pCompile_ArgTwoL_Expression d, tConst, tArg, fConst
 If Mid(m_sExp, m_lps, 1) <> ")" Then Err.Raise 5
 m_lps = m_lps + 1
Case 48 To 57, 46, 43, 45 '"0"-"9",".","+","-" : number
 If i = 43 Or i = 45 Then
  m_lps = m_lps + 1
  bSign = True
 End If
 i = m_lps
 Do Until i > m_lpm
  Select Case Asc(Mid(m_sExp, i, 1))
  Case 48 To 57 'number
  Case 46 '"."
   If bHasPoint Then Exit Do Else bHasPoint = True
  Case 69, 101
   If bHasExp Or i = m_lps Then
    Exit Do
   Else
    Select Case Mid(m_sExp, i + 1, 1)
    Case "+", "-"
     i = i + 1
    End Select
    bHasExp = True
   End If
  Case Else
   Exit Do
  End Select
  i = i + 1
 Loop
 If i > m_lps Then
  If bSign Then m_lps = m_lps - 1
  f = Val(Mid(m_sExp, m_lps, i - m_lps))
  m_lps = i
  'find const
  With tConst
   For i = 0 To .nCount - 1
    If f = fConst(i) Then Exit For
   Next i
   Debug.Assert i < 256
   If i >= .nCount Then
    ReDim Preserve .s(.nCount)
    .nCount = .nCount + 1
    'TODO:resize fConst ???
    fConst(i) = f
   End If
  End With
  With d
   .nCount = .nCount + 2
   ReDim Preserve .d(1 To .nCount)
   .d(.nCount - 1) = 1
   .d(.nCount) = i
  End With
 Else
  Err.Raise 5
 End If
Case 95, 65 To 90, 97 To 122 '"_","A"-"Z","a"-"z" : var
 i = m_lps
 Do
  m_lps = m_lps + 1
  If m_lps > m_lpm Then Exit Do
  Select Case Asc(Mid(m_sExp, m_lps, 1))
  Case 95, 65 To 90, 97 To 122, 48 To 57
  Case Else
   Exit Do
  End Select
 Loop
 s = Mid(m_sExp, i, m_lps - i)
 i = pSearchSymTable(s, tConst)
 If i < 0 Then
  i = pSearchSymTables(s, tArg)
  If i < 0 Then Err.Raise 5
  Debug.Assert i < 256
  With d
   .nCount = .nCount + 2
   ReDim Preserve .d(1 To .nCount)
   .d(.nCount - 1) = 0
   .d(.nCount) = i
  End With
 Else 'const
  Debug.Assert i < 256
  With d
   .nCount = .nCount + 2
   ReDim Preserve .d(1 To .nCount)
   .d(.nCount - 1) = 1
   .d(.nCount) = i
  End With
 End If
Case Else
 Err.Raise 5
End Select
End Sub

'case sensitive
Private Function pSearchSymTable(ByRef s As String, t As typeSymbolTable) As Long
Dim i As Long
For i = 0 To t.nCount - 1
 If t.s(i) = s Then
  pSearchSymTable = i
  Exit Function
 End If
Next i
pSearchSymTable = -1
End Function

Private Function pSearchSymTables(ByRef s As String, t As typeSymbolTables) As Long
Dim i As Long, j As Long
For i = 0 To t.nCount - 1
 For j = 0 To t.d(i).nCount - 1
  If t.d(i).s(j) = s Then
   pSearchSymTables = i * 8& + j
   Exit Function
  End If
 Next j
Next i
pSearchSymTables = -1
End Function

Private Sub pArgZeroL_Apply(tDest() As typeArgZeroL_Operation, ByVal nStart As Long, ByVal nCount As Long, d() As typeArgZeroL_OperationD, fConst() As Single, fArg() As Single)
Dim i As Long, j As Long
For i = 1 To nCount
 With tDest(nStart + i - 1)
  .nType = d(i).nType
  .nIndex = d(i).nIndex
  .nArgCount = d(i).nArgCount
  For j = 0 To .nArgCount - 1
   .fArg(j) = pCalc(d(i).tExp(j), fConst, fArg)
  Next j
 End With
Next i
End Sub

Private Function pCalc(d As typeArgZeroL_Expression, fConst() As Single, fArg() As Single) As Single
Dim i As Long
Dim esp As Long
TheCalcStack(0) = 0 '???
esp = -1
For i = 1 To d.nCount
 Select Case d.d(i)
 Case 0 'arg
  esp = esp + 1
  i = i + 1
  TheCalcStack(esp) = fArg(d.d(i))
 Case 1 'const
  esp = esp + 1
  i = i + 1
  TheCalcStack(esp) = fConst(d.d(i))
 Case 2 '+
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) + TheCalcStack(esp + 1)
 Case 3 '-
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) - TheCalcStack(esp + 1)
 Case 4 '*
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) * TheCalcStack(esp + 1)
 Case 5 '/
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) / TheCalcStack(esp + 1)
 Case 6 '^
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) ^ TheCalcStack(esp + 1)
 Case 7 '<
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) < TheCalcStack(esp + 1)
 Case 8 '>
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) > TheCalcStack(esp + 1)
 Case 9 '=
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) = TheCalcStack(esp + 1)
 Case 10 'and
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) And TheCalcStack(esp + 1)
 Case 11 'or
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) Or TheCalcStack(esp + 1)
 Case 12 'negative
  TheCalcStack(esp) = -TheCalcStack(esp)
 Case 13 'not
  TheCalcStack(esp) = TheCalcStack(esp) = 0
 Case 14 '<=
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) <= TheCalcStack(esp + 1)
 Case 15 '>=
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) >= TheCalcStack(esp + 1)
 Case 16 '<>
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) <> TheCalcStack(esp + 1)
 Case Else
  Debug.Assert False
 End Select
Next i
pCalc = TheCalcStack(0)
End Function

Private Sub pArgTwoL_Apply(tDest() As typeArgZeroL_Operation, ByVal nStart As Long, ByVal nCount As Long, d() As typeArgZeroL_OperationD, fConst() As Single, tSrc() As typeArgZeroL_Operation, nIndex() As Long)
Dim i As Long, j As Long
For i = 1 To nCount
 With tDest(nStart + i - 1)
  .nType = d(i).nType
  .nIndex = d(i).nIndex
  .nArgCount = d(i).nArgCount
  For j = 0 To .nArgCount - 1
   .fArg(j) = pCalc2(d(i).tExp(j), fConst, tSrc, nIndex)
  Next j
 End With
Next i
End Sub

'TODO:
Private Function pCalc2(d As typeArgZeroL_Expression, fConst() As Single, tSrc() As typeArgZeroL_Operation, nIndex() As Long) As Single
Dim i As Long, j As Long
Dim esp As Long
TheCalcStack(0) = 0 '???
esp = -1
For i = 1 To d.nCount
 Select Case d.d(i)
 Case 0 'arg
  esp = esp + 1
  i = i + 1
  j = d.d(i)
  TheCalcStack(esp) = tSrc(nIndex(j \ 8&)).fArg(j And 7&)
 Case 1 'const
  esp = esp + 1
  i = i + 1
  TheCalcStack(esp) = fConst(d.d(i))
 Case 2 '+
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) + TheCalcStack(esp + 1)
 Case 3 '-
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) - TheCalcStack(esp + 1)
 Case 4 '*
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) * TheCalcStack(esp + 1)
 Case 5 '/
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) / TheCalcStack(esp + 1)
 Case 6 '^
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) ^ TheCalcStack(esp + 1)
 Case 7 '<
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) < TheCalcStack(esp + 1)
 Case 8 '>
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) > TheCalcStack(esp + 1)
 Case 9 '=
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) = TheCalcStack(esp + 1)
 Case 10 'and
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) And TheCalcStack(esp + 1)
 Case 11 'or
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) Or TheCalcStack(esp + 1)
 Case 12 'negative
  TheCalcStack(esp) = -TheCalcStack(esp)
 Case 13 'not
  TheCalcStack(esp) = TheCalcStack(esp) = 0
 Case 14 '<=
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) <= TheCalcStack(esp + 1)
 Case 15 '>=
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) >= TheCalcStack(esp + 1)
 Case 16 '<>
  esp = esp - 1
  TheCalcStack(esp) = TheCalcStack(esp) <> TheCalcStack(esp + 1)
 Case Else
  Debug.Assert False
 End Select
Next i
pCalc2 = TheCalcStack(0)
End Function

Private Function pInterprete_ArgZeroL_Expression(tConst As typeSymbolTable, fConst() As Single) As Single
Dim f As Single
f = pInterprete_ArgZeroL_ExpressionC(tConst, fConst)
Do Until m_lps > m_lpm
 Select Case Mid(m_sExp, m_lps, 1)
 Case "&"
  m_lps = m_lps + 1
  f = pInterprete_ArgZeroL_ExpressionC(tConst, fConst) <> 0 And f <> 0
 Case "|"
  m_lps = m_lps + 1
  f = pInterprete_ArgZeroL_ExpressionC(tConst, fConst) <> 0 Or f <> 0
 Case Else
  Exit Do
 End Select
Loop
pInterprete_ArgZeroL_Expression = f
End Function

Private Function pInterprete_ArgZeroL_ExpressionC(tConst As typeSymbolTable, fConst() As Single) As Single
Dim b As Boolean, f As Single
Do While Mid(m_sExp, m_lps, 1) = "!"
 b = Not b
 m_lps = m_lps + 1
Loop
f = pInterprete_ArgZeroL_ExpressionE(tConst, fConst)
Select Case Mid(m_sExp, m_lps, 1)
Case "<"
 m_lps = m_lps + 1
 Select Case Mid(m_sExp, m_lps, 1)
 Case "="
  m_lps = m_lps + 1
  f = f <= pInterprete_ArgZeroL_ExpressionE(tConst, fConst)
 Case ">"
  m_lps = m_lps + 1
  f = f <> pInterprete_ArgZeroL_ExpressionE(tConst, fConst)
 Case Else
  f = f < pInterprete_ArgZeroL_ExpressionE(tConst, fConst)
 End Select
Case ">"
 m_lps = m_lps + 1
 If Mid(m_sExp, m_lps, 1) = "=" Then
  m_lps = m_lps + 1
  f = f >= pInterprete_ArgZeroL_ExpressionE(tConst, fConst)
 Else
  f = f > pInterprete_ArgZeroL_ExpressionE(tConst, fConst)
 End If
Case "="
 m_lps = m_lps + 1
 f = f = pInterprete_ArgZeroL_ExpressionE(tConst, fConst)
End Select
If b Then f = f = 0
pInterprete_ArgZeroL_ExpressionC = f
End Function

Private Function pInterprete_ArgZeroL_ExpressionE(tConst As typeSymbolTable, fConst() As Single) As Single
Dim f As Single
f = pInterprete_ArgZeroL_ExpressionT(tConst, fConst)
Do Until m_lps > m_lpm
 Select Case Mid(m_sExp, m_lps, 1)
 Case "+"
  m_lps = m_lps + 1
  f = f + pInterprete_ArgZeroL_ExpressionT(tConst, fConst)
 Case "-"
  m_lps = m_lps + 1
  f = f - pInterprete_ArgZeroL_ExpressionT(tConst, fConst)
 Case Else
  Exit Do
 End Select
Loop
pInterprete_ArgZeroL_ExpressionE = f
End Function

Private Function pInterprete_ArgZeroL_ExpressionT(tConst As typeSymbolTable, fConst() As Single) As Single
Dim b As Boolean, f As Single
Do While Mid(m_sExp, m_lps, 1) = "-"
 b = Not b
 m_lps = m_lps + 1
Loop
f = pInterprete_ArgZeroL_ExpressionF(tConst, fConst)
Do Until m_lps > m_lpm
 Select Case Mid(m_sExp, m_lps, 1)
 Case "*"
  m_lps = m_lps + 1
  f = f * pInterprete_ArgZeroL_ExpressionF(tConst, fConst)
 Case "/"
  m_lps = m_lps + 1
  f = f / pInterprete_ArgZeroL_ExpressionF(tConst, fConst)
 Case Else
  Exit Do
 End Select
Loop
If b Then f = -f
pInterprete_ArgZeroL_ExpressionT = f
End Function

Private Function pInterprete_ArgZeroL_ExpressionF(tConst As typeSymbolTable, fConst() As Single) As Single
Dim f As Single
f = pInterprete_ArgZeroL_ExpressionG(tConst, fConst)
Do Until m_lps > m_lpm
 Select Case Mid(m_sExp, m_lps, 1)
 Case "^"
  m_lps = m_lps + 1
  f = f ^ pInterprete_ArgZeroL_ExpressionG(tConst, fConst)
 Case Else
  Exit Do
 End Select
Loop
pInterprete_ArgZeroL_ExpressionF = f
End Function

Private Function pInterprete_ArgZeroL_ExpressionG(tConst As typeSymbolTable, fConst() As Single) As Single
Dim i As Long, s As String
Dim bHasExp As Boolean
Dim bHasPoint As Boolean
Dim bSign As Boolean
Dim f As Single
i = Asc(Mid(m_sExp, m_lps, 1))
Select Case i
Case 40 '"("
 m_lps = m_lps + 1
 f = pInterprete_ArgZeroL_Expression(tConst, fConst)
 If Mid(m_sExp, m_lps, 1) <> ")" Then Err.Raise 5
 m_lps = m_lps + 1
Case 48 To 57, 46, 43, 45 '"0"-"9",".","+","-" : number
 If i = 43 Or i = 45 Then
  m_lps = m_lps + 1
  bSign = True
 End If
 i = m_lps
 Do Until i > m_lpm
  Select Case Asc(Mid(m_sExp, i, 1))
  Case 48 To 57 'number
  Case 46 '"."
   If bHasPoint Then Exit Do Else bHasPoint = True
  Case 69, 101
   If bHasExp Or i = m_lps Then
    Exit Do
   Else
    Select Case Mid(m_sExp, i + 1, 1)
    Case "+", "-"
     i = i + 1
    End Select
    bHasExp = True
   End If
  Case Else
   Exit Do
  End Select
  i = i + 1
 Loop
 If i > m_lps Then
  If bSign Then m_lps = m_lps - 1
  f = Val(Mid(m_sExp, m_lps, i - m_lps))
  m_lps = i
 Else
  Err.Raise 5
 End If
Case 95, 65 To 90, 97 To 122 '"_","A"-"Z","a"-"z" : const
 i = m_lps
 Do
  m_lps = m_lps + 1
  If m_lps > m_lpm Then Exit Do
  Select Case Asc(Mid(m_sExp, m_lps, 1))
  Case 95, 65 To 90, 97 To 122, 48 To 57
  Case Else
   Exit Do
  End Select
 Loop
 s = Mid(m_sExp, i, m_lps - i)
 i = pSearchSymTable(s, tConst)
 If i < 0 Then Err.Raise 5
 Debug.Assert i < 256
 f = fConst(i)
Case Else
 Err.Raise 5
End Select
pInterprete_ArgZeroL_ExpressionG = f
End Function

Private Sub pCompile_TwoL(v As Variant, ls As typeTwoL)
On Error GoTo a
Dim s As String, m As Long
Dim nIdx As Long
Dim i As Long, j As Long, lps As Long
Dim f As Long
'init
With ls
 .nCount = 0
 Erase .d
 ReDim .nTable(255, 255)
 ReDim .bIsIgnored(255, 255)
 'get
 m = UBound(v)
 For nIdx = 0 To m
  s = Trim(v(nIdx))
  If s = "" Then
   'comment only
  ElseIf Left(s, 1) = "#" Then
   'find space
   i = InStr(1, s, " ")
   If i > 0 Then
    Select Case Mid(s, 2, i - 2)
    Case "ignore"
     s = Replace(Mid(s, i + 1), " ", "")
     f = Len(s)
     lps = 1
     Do While lps <= f
      i = AscB(Mid(s, lps, 1))
      If Mid(s, lps + 1, 1) = "_" And lps + 2 <= f Then
       j = AscB(Mid(s, lps + 2, 1))
       lps = lps + 3
      Else
       j = 0
       lps = lps + 1
      End If
      .bIsIgnored(i, j) = 1
     Loop
    Case Else
     'TODO:#include
    End Select
   End If
  Else
   .nCount = .nCount + 1
   ReDim Preserve .d(1 To .nCount)
   'is rule?
   s = Replace(s, " ", "")
   lps = InStr(1, s, "=>")
   If lps > 0 Then
    'get left context
    f = InStr(1, s, "<")
    If f > 0 And f < lps Then
     pCompile_ZeroL_Rule Left(s, f - 1), .d(.nCount).LeftContext, .d(.nCount).LeftContextCount
     f = f + 1
    Else
     f = 1
    End If
    'get type
    i = AscB(Mid(s, f, 1))
    If Mid(s, f + 1, 1) = "_" Then
     j = AscB(Mid(s, f + 2, 1))
    Else
     j = 0
    End If
    'process linked-list
    .d(.nCount).idxNext = .nTable(i, j)
    .nTable(i, j) = .nCount
    'get right context
    f = InStr(f, s, ">")
    If f > 0 And f < lps Then
     pCompile_ZeroL_Rule Mid(s, f + 1, lps - f - 1), .d(.nCount).RightContext, .d(.nCount).RightContextCount
    End If
    'over
    s = Mid(s, lps + 2)
   End If
   pCompile_ZeroL_Rule s, .d(.nCount).d, .d(.nCount).nCount
  End If
 Next nIdx
End With
Exit Sub
a:
ls.nCount = 0
Erase ls.d, ls.nTable, ls.bIsIgnored
End Sub

'A context-sensitive extension of tree L-systems requires neighbor
'edges of the replaced edge to be tested for context matching. A predecessor
'of a context-sensitive production p consists of three components:
'a **PATH** l forming the left context, an edge S called the strict predecessor,
'and an **AXIAL TREE** r constituting the right context (Figure 1.29). THE
'ASYMMETRY BETWEEN THE LEFT CONTEXT AND THE RIGHT CONTEXT reflects the
'fact that there is only one path from the root of a tree to a given edge,
'while there can be many paths from this edge to various terminal nodes.
'Production p matches a given occurrence of the edge S in a tree T if l
'is a path in T terminating at the starting node of S, and r is a subtree
'of T originating at the ending node of S. The production can then be
'applied by replacing S with the axial tree specified as the production
'successor.

Private Sub pCalcTwoLTest(bmOut As typeAlphaDibSectionDescriptor, bProps() As Byte, sProps As String, v As Variant)
Dim TheClrTable() As RGBQUAD, TheClrCount As Long
Dim nCount As Long
Dim fX As Single, fY As Single, FS As Single
Dim d As Single, l As Single
Dim bHQ As Boolean, b As Boolean
Dim nXCount As Long, nYCount As Long
Dim nXNow As Long, nYNow As Long
Dim TheSeed As Long
'Dim TheSeed2(255) As Long
Dim ls As typeTwoL
'///string replace algorithm
Dim TheString(1) As typeTwoL_String
'///
Dim nCur As Long 'current generation
Dim nBranch As Long 'current branch
Dim nDeletedBranch As Long
Dim tState() As typeLSystemState 'state stack
Dim curState As typeLSystemState
Dim tPolyPt() As typeLSystemPoint 'polygon point stack
Dim tPoly() As Long 'polygon point count stack
Dim nPolyPt As Long 'polygon point
Dim nPoly As Long 'polygon index
'///
Dim i As Long, j As Long
Dim op As typeZeroL_Operation
Dim idx As Long 'current string index = 0 or 1
Dim lp As Long ', lp2 As Long
'Dim f As Long
Dim hbr As Long, hpn As Long
'///
'get data
nXCount = bProps(0)
nYCount = (nXCount And 224&) \ 32&
bHQ = (nXCount And 1&)
nXCount = (nXCount And 28&) \ 4&
pGetSeed bProps, 1, TheSeed
CopyMemory fX, bProps(3), 4&
CopyMemory fY, bProps(7), 4&
CopyMemory FS, bProps(11), 4&
CopyMemory d, bProps(15), 4&
CopyMemory l, bProps(19), 4&
nCount = bProps(23)
'convert data
d = d * 二π
FS = FS * 二π
l = l * bmOut.Width  '???
'get color
pGetColor sProps, TheClrTable, TheClrCount
'//////////////////////string replace algorithm
'compile?
pCompile_TwoL v, ls
If ls.nCount = 0 Or nXCount = 0 Or nYCount = 0 Then Exit Sub
'init array
Dim ox As Long, oy As Long
TheBitmap = bmOut
With TheSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = bmOut.Height
 .Bounds(1).cElements = bmOut.Width
 .pvData = bmOut.lpbm
End With
CopyMemory ByVal VarPtrArray(TheArray()), VarPtr(TheSA), 4&
'init array
ReDim tStack(255), tState(255), tPolyPt(4095), tPoly(255)
With TheString(0)
 .nMax = ls.d(1).nCount + 65536
 ReDim .d(.nMax - 1)
End With
With TheString(1)
 .nMax = TheString(0).nMax
 ReDim .d(.nMax - 1)
End With
'init color cache
TheClrIndex = &H80000000
i = -1
CopyMemory TheClr, i, 4&
'start draw
For nXNow = 0 To nXCount - 1
 For nYNow = 0 To nYCount - 1
  nBranch = 0
  nDeletedBranch = &H7FFFFFFF
  nPolyPt = 0
  nPoly = -1
  With curState
   .x = fX + nXNow / nXCount
   If .x > 1 Then .x = .x - 1
   .x = .x * bmOut.Width
   .y = fY + nYNow / nYCount
   If .y > 1 Then .y = .y - 1
   .y = .y * bmOut.Height
   .FS = FS
   .idxClr = 0
   .idxWidth = 0
  End With
  idx = nCount And 1&
  'init string
  With TheString(idx)
   .nCount = ls.d(1).nCount
   If .nCount > 0 Then CopyMemory .d(0), ls.d(1).d(1), .nCount * 2&
  End With
  'start replace
  For nCur = 1 To nCount
   TheString(1 - idx).nCount = 0
   With TheString(idx)
    For lp = 0 To .nCount - 1
     'find expand
     j = 0
     op = .d(lp)
     With ls
      i = .nTable(op.nType, op.nIndex)
      Do Until i = 0
       With .d(i)
        'match left context
        If .LeftContextCount > 0 Then
         b = pMatch_TwoL_LeftContext(.LeftContext, .LeftContextCount, ls.bIsIgnored, TheString(idx), lp - 1)
        Else
         b = True
        End If
        If b And .RightContextCount > 0 Then
         'match right context
         b = pMatch_TwoL_RightContext(.RightContext, .RightContextCount, ls.bIsIgnored, TheString(idx), lp + 1)
        End If
       End With
       If b Then
        j = i
        'TODO:probability
        Exit Do
       End If
       i = .d(i).idxNext
      Loop
     End With
     If j > 0 Then 'expand it
      i = ls.d(j).nCount
      If i > 0 Then
       With TheString(1 - idx)
        .nCount = .nCount + i
        If .nCount >= .nMax Then
         .nMax = .nMax + i + 65536
         ReDim Preserve .d(.nMax - 1)
        End If
        CopyMemory .d(.nCount - i), ls.d(j).d(1), i * 2&
       End With
      End If
     Else 'just keep unchanged
      With TheString(1 - idx)
       .nCount = .nCount + 1
       If .nCount >= .nMax Then
        .nMax = .nMax + 65536
        ReDim Preserve .d(.nMax - 1)
       End If
       .d(.nCount - 1) = op
      End With
     End If
    Next lp
   End With
   idx = 1 - idx
  Next nCur
  'start draw
  With TheString(0)
   For nCur = 0 To .nCount - 1
    op = .d(nCur)
    Select Case op.nType
    Case 70, 71, 102, 103 'F,G,f,g
     If nBranch < nDeletedBranch Then
      With curState
       If op.nType < 80 Then
        '////////get color!!!
        If .idxClr <> TheClrIndex And TheClrCount > 0 Then
         TheClrIndex = .idxClr
         If TheClrIndex < 0 Then
          TheClr = TheClrTable(0)
         ElseIf TheClrIndex >= TheClrCount Then
          TheClr = TheClrTable(TheClrCount - 1)
         Else
          TheClr = TheClrTable(TheClrIndex)
         End If
        End If
        #If UseLineDDA Then
        #Else
        If TheClrCount > 0 Then
         With TheClr
          i = .rgbRed + .rgbGreen * &H100& + .rgbBlue * &H10000
         End With
        Else
         i = vbWhite
        End If
        #End If
        '////////
        #If UseLineDDA Then
        ox = .x
        oy = .y
        #Else
        MoveToEx bmOut.hdc, .x, .y, ByVal 0
        #End If
       End If
       .x = .x + l * Cos(.FS)
       .y = .y + l * Sin(.FS)
       If op.nType < 80 Then
        #If UseLineDDA Then
        If TheClr.rgbReserved > 0 Then
         If TheClr.rgbReserved = 255 Then
          LineDDA ox, oy, .x, .y, AddressOf pDrawLine, 0
         Else
          LineDDA ox, oy, .x, .y, AddressOf pDrawLineAlpha, 0
         End If
        End If
        #Else
        'ERROR!!! API didn't support 32-bit bitmap
        hpn = SelectObject(bmOut.hdc, CreatePen(0, 1, i))
        LineTo bmOut.hdc, .x, .y
        DeleteObject SelectObject(bmOut.hdc, hpn)
        #End If
       End If
       'record vertex
       If nPoly >= 0 And (op.nType And 1) = 0 Then
        tPoly(nPoly) = tPoly(nPoly) + 1
        With tPolyPt(nPolyPt)
         .x = curState.x
         .y = curState.y
        End With
        nPolyPt = nPolyPt + 1
       End If
      End With
     End If
    Case 46 '.
     'record vertex
     If nBranch < nDeletedBranch And nPoly >= 0 Then
      tPoly(nPoly) = tPoly(nPoly) + 1
      With tPolyPt(nPolyPt)
       .x = curState.x
       .y = curState.y
      End With
      nPolyPt = nPolyPt + 1
     End If
    Case 91 '[
     tState(nBranch) = curState
     nBranch = nBranch + 1
    Case 93 ']
     If nBranch > 0 Then nBranch = nBranch - 1
     If nBranch < nDeletedBranch Then nDeletedBranch = &H7FFFFFFF
     curState = tState(nBranch)
    Case 123 '{
     nPoly = nPoly + 1
     tPoly(nPoly) = 0
    Case 125 '}
     If nPoly >= 0 Then
      j = tPoly(nPoly)
      nPoly = nPoly - 1
      If j > 0 Then
       With curState
        '////////get color!!!
        If .idxClr <> TheClrIndex And TheClrCount > 0 Then
         TheClrIndex = .idxClr
         If TheClrIndex < 0 Then
          TheClr = TheClrTable(0)
         ElseIf TheClrIndex >= TheClrCount Then
          TheClr = TheClrTable(TheClrCount - 1)
         Else
          TheClr = TheClrTable(TheClrIndex)
         End If
        End If
        #If UseLineDDA Then
        #Else
        If TheClrCount > 0 Then
         With TheClr
          i = .rgbRed + .rgbGreen * &H100& + .rgbBlue * &H10000
         End With
        Else
         i = vbWhite
        End If
        #End If
        '////////
       End With
       nPolyPt = nPolyPt - j
       If nBranch < nDeletedBranch Then
        #If UseLineDDA Then
        pDrawPolygon tPolyPt, nPolyPt, j
        #Else
        pDrawPolygonTest tPolyPt, nPolyPt, j, i
        #End If
       End If
      End If
     End If
    Case 43 '+
     If nBranch < nDeletedBranch Then
      With curState
       .FS = .FS - d
       If .FS < 0 Then .FS = .FS + 二π
      End With
     End If
    Case 45 '-
     If nBranch < nDeletedBranch Then
      With curState
       .FS = .FS + d
       If .FS > 二π Then .FS = .FS - 二π
      End With
     End If
    Case 124 '|
     If nBranch < nDeletedBranch Then
      With curState
       If .FS < π Then .FS = .FS + π Else .FS = .FS - π
      End With
     End If
    Case 33 '!
     If nBranch < nDeletedBranch Then
      With curState
       .idxWidth = .idxWidth + 1 '??? TODO:
      End With
     End If
    Case 39 ''
     If nBranch < nDeletedBranch Then
      With curState
       .idxClr = .idxClr + 1 '??? TODO:
      End With
     End If
    Case 37 '%
     nDeletedBranch = nBranch
    End Select
   Next nCur
  End With
 Next nYNow
Next nXNow
'destroy
ZeroMemory ByVal VarPtrArray(TheArray()), 4&
End Sub

Private Function pMatch_TwoL_LeftContext(LeftContext() As typeZeroL_Operation, ByVal LeftContextCount As Long, bIsIgnored() As Byte, s As typeTwoL_String, ByVal nStart As Long) As Boolean
Dim i As Long, j As Long, k As Long
j = LeftContextCount
For i = nStart To 0 Step -1
 With s.d(i)
  If .nType = 91 Then '"["
   If k > 0 Then k = k - 1
  ElseIf .nType = 93 Then '"]"
   k = k + 1
  ElseIf k = 0 Then
   If bIsIgnored(.nType, .nIndex) = 0 Then
    If .nType = LeftContext(j).nType And .nIndex = LeftContext(j).nIndex Then
     j = j - 1
     If j = 0 Then Exit For
    Else
     Exit For
    End If
   End If
  End If
 End With
Next i
pMatch_TwoL_LeftContext = j = 0
End Function

'path only!!! TODO:add axial tree support
Private Function pMatch_TwoL_RightContext(RightContext() As typeZeroL_Operation, ByVal RightContextCount As Long, bIsIgnored() As Byte, s As typeTwoL_String, ByVal nStart As Long, Optional ByVal nContextStart As Long = 1) As Boolean
Dim i As Long, j As Long, k As Long
j = nContextStart
For i = nStart To s.nCount - 1
 With s.d(i)
  If .nType = 91 Then '"["
   If k = 0 Then
    'recursive
    If pMatch_TwoL_RightContext(RightContext, RightContextCount, bIsIgnored, s, i + 1, j) Then
     pMatch_TwoL_RightContext = True
     Exit For
    End If
   End If
   k = k + 1
  ElseIf .nType = 93 Then '"]"
   k = k - 1
   If k < 0 Then Exit For
  ElseIf k = 0 Then
   If bIsIgnored(.nType, .nIndex) = 0 Then
    If .nType = RightContext(j).nType And .nIndex = RightContext(j).nIndex Then
     j = j + 1
     If j > RightContextCount Then
      pMatch_TwoL_RightContext = True
      Exit For
     End If
    Else
     Exit For
    End If
   End If
  End If
 End With
Next i
End Function

Private Function pMatch_ArgTwoL_LeftContext(LeftContext() As typeArgTwoL_OperationContext, ByVal LeftContextCount As Long, bIsIgnored() As Byte, s As typeArgTwoL_String, ByVal nStart As Long, nIndex() As Long) As Boolean
Dim i As Long, j As Long, k As Long, kk As Long
j = LeftContextCount
For i = nStart To 0 Step -1
 With s.d(i)
  If .nType = 91 Then '"["
   If k > 0 Then k = k - 1
  ElseIf .nType = 93 Then '"]"
   k = k + 1
  ElseIf k = 0 Then
   If bIsIgnored(.nType, .nIndex) = 0 Then
    If .nType = LeftContext(j).nType And .nIndex = LeftContext(j).nIndex And _
    .nArgCount = LeftContext(j).nArgCount Then
     kk = LeftContext(j).nReserved
     If kk > 0 Then nIndex(kk) = i
     j = j - 1
     If j = 0 Then Exit For
    Else
     Exit For
    End If
   End If
  End If
 End With
Next i
pMatch_ArgTwoL_LeftContext = j = 0
End Function

'path only!!! TODO:add axial tree support
Private Function pMatch_ArgTwoL_RightContext(RightContext() As typeArgTwoL_OperationContext, ByVal RightContextCount As Long, bIsIgnored() As Byte, s As typeArgTwoL_String, ByVal nStart As Long, nIndex() As Long, Optional ByVal nContextStart As Long = 1) As Boolean
Dim i As Long, j As Long, k As Long, kk As Long
j = nContextStart
For i = nStart To s.nCount - 1
 With s.d(i)
  If .nType = 91 Then '"["
   If k = 0 Then
    'recursive
    If pMatch_ArgTwoL_RightContext(RightContext, RightContextCount, bIsIgnored, s, i + 1, nIndex, j) Then
     pMatch_ArgTwoL_RightContext = True
     Exit For
    End If
   End If
   k = k + 1
  ElseIf .nType = 93 Then '"]"
   k = k - 1
   If k < 0 Then Exit For
  ElseIf k = 0 Then
   If bIsIgnored(.nType, .nIndex) = 0 Then
    If .nType = RightContext(j).nType And .nIndex = RightContext(j).nIndex And _
    .nArgCount = RightContext(j).nArgCount Then
     kk = RightContext(j).nReserved
     If kk > 0 Then nIndex(kk) = i
     j = j + 1
     If j > RightContextCount Then
      pMatch_ArgTwoL_RightContext = True
      Exit For
     End If
    Else
     Exit For
    End If
   End If
  End If
 End With
Next i
End Function

Private Sub pCompile_ArgTwoL(v As Variant, ls As typeArgTwoL)
On Error GoTo a
Dim s As String, m As Long
Dim nIdx As Long
Dim i As Long, j As Long, lps As Long
Dim lp As Long
Dim f As Long
Dim tConst As typeSymbolTable
Dim tArg As typeSymbolTables
'init
With tConst
 .nCount = 3 'stupid??
 ReDim .s(2)
 .s(0) = "_a"
 .s(1) = "_l"
 .s(2) = "pi" 'stupid??
End With
With ls
 .nCount = 0
 Erase .d
 ReDim .nTable(255, 255)
 ReDim .bIsIgnored(255, 255)
 'get
 m = UBound(v)
 For nIdx = 0 To m
  s = Trim(v(nIdx))
  If s = "" Then
   'comment only
  ElseIf Left(s, 1) = "#" Then
   'delete space :-3
   i = Len(s)
   Do
    s = Replace(s, "  ", " ")
    j = Len(s)
    If i = j Then Exit Do
    i = j
   Loop
   'find space
   i = InStr(1, s, " ")
   If i > 0 Then
    Select Case Mid(s, 2, i - 2)
    Case "define"
     'find space again
     j = InStr(i + 1, s, " ")
     If j > 0 Then
      'parse expression
      m_sExp = Replace(Mid(s, j + 1), " ", "")
      m_lps = 1
      m_lpm = Len(m_sExp)
      .fConst(tConst.nCount) = pInterprete_ArgZeroL_Expression(tConst, .fConst)
      With tConst
       ReDim Preserve .s(.nCount)
       .s(.nCount) = Mid(s, i + 1, j - i - 1)
       .nCount = .nCount + 1
      End With
     End If
    Case "ignore"
     s = Replace(Mid(s, i + 1), " ", "")
     f = Len(s)
     lps = 1
     Do While lps <= f
      i = AscB(Mid(s, lps, 1))
      If Mid(s, lps + 1, 1) = "_" And lps + 2 <= f Then
       j = AscB(Mid(s, lps + 2, 1))
       lps = lps + 3
      Else
       j = 0
       lps = lps + 1
      End If
      .bIsIgnored(i, j) = 1
     Loop
    Case Else
     'TODO:#include
    End Select
   End If
  Else
   .nCount = .nCount + 1
   ReDim Preserve .d(1 To .nCount)
   'is rule?
   s = Replace(s, " ", "")
   lps = InStr(1, s, "=>")
   If lps > 0 Then
    'init
    With tArg
     .nCount = 1
     ReDim .d(0)
    End With
    lp = InStr(1, s, ":")
    'get left context
    f = InStr(1, s, "<")
    If f > 0 And f < lps And (lp = 0 Or f < lp) Then
     pCompile_ArgTwoL_Context Left(s, f - 1), .d(.nCount).LeftContext, .d(.nCount).LeftContextCount, tArg
     f = f + 1
    Else
     f = 1
    End If
    'get type
    i = AscB(Mid(s, f, 1))
    If Mid(s, f + 1, 1) = "_" Then
     j = AscB(Mid(s, f + 2, 1))
     f = f + 3
    Else
     j = 0
     f = f + 1
    End If
    'process linked-list
    .d(.nCount).idxNext = .nTable(i, j)
    .nTable(i, j) = .nCount
    'process arguments
    If Mid(s, f, 1) = "(" Then
     i = InStr(f, s, ")")
     Do
      j = InStr(f + 1, s, ",")
      If j = 0 Or j > i Then j = i
      With tArg.d(0)
       ReDim Preserve .s(.nCount)
       .s(.nCount) = Mid(s, f + 1, j - f - 1)
       .nCount = .nCount + 1
      End With
      f = j
     Loop While j < i
    End If
    With .d(.nCount)
     .nArgCount = tArg.d(0).nCount
     .tExp.nCount = 0
     Erase .tExp.d
     If .nArgCount > 7 Then Err.Raise 9
    End With
    'get right context
    f = InStr(f, s, ">")
    If f > 0 And f < lps And (lp = 0 Or f < lp) Then
     i = lps
     If i > lp And lp > 0 Then i = lp
     pCompile_ArgTwoL_Context Mid(s, f + 1, i - f - 1), .d(.nCount).RightContext, .d(.nCount).RightContextCount, tArg
    End If
    'process expression
    If lp > 0 Then
     m_sExp = Mid(s, lp + 1, lps - lp - 1)
     m_lps = 1
     m_lpm = Len(m_sExp)
     pCompile_ArgTwoL_Expression .d(.nCount).tExp, tConst, tArg, .fConst
    End If
    'over
    s = Mid(s, lps + 2)
   Else
    tArg.nCount = 0 'not a rule!
   End If
   m_sExp = s
   pCompile_ArgTwoL_Rule .d(.nCount).d, .d(.nCount).nCount, tConst, tArg, .fConst
  End If
 Next nIdx
End With
Exit Sub
a:
ls.nCount = 0
Erase ls.d, ls.nTable, ls.bIsIgnored
End Sub

Private Sub pCompile_ArgTwoL_Context(ByVal s As String, d() As typeArgTwoL_OperationContext, nCount As Long, tArg As typeSymbolTables)
Dim lps As Long, m As Long
Dim i As Long, j As Long
m = Len(s)
lps = 1
Do While lps <= m
 i = AscB(Mid(s, lps, 1))
 If Mid(s, lps + 1, 1) = "_" And lps + 2 <= m Then
  j = AscB(Mid(s, lps + 2, 1))
  lps = lps + 3
 Else
  j = 0
  lps = lps + 1
 End If
 nCount = nCount + 1
 ReDim Preserve d(1 To nCount)
 With d(nCount)
  .nType = i
  .nIndex = j
 End With
 'process arguments
 If Mid(s, lps, 1) = "(" Then
  With tArg
   ReDim Preserve .d(.nCount)
   .nCount = .nCount + 1
  End With
  i = InStr(lps, s, ")")
  Do
   j = InStr(lps + 1, s, ",")
   If j = 0 Or j > i Then j = i
   With tArg.d(tArg.nCount - 1)
    ReDim Preserve .s(.nCount)
    .s(.nCount) = Mid(s, lps + 1, j - lps - 1)
    .nCount = .nCount + 1
   End With
   lps = j
  Loop While j < i
  lps = lps + 1
  With d(nCount)
   .nReserved = tArg.nCount - 1
   .nArgCount = tArg.d(.nReserved).nCount
   If .nArgCount > 7 Then Err.Raise 9
  End With
 End If
Loop
End Sub

Private Sub pCalcArgTwoLTest(bmOut As typeAlphaDibSectionDescriptor, bProps() As Byte, sProps As String, v As Variant)
Dim TheClrTable() As RGBQUAD, TheClrCount As Long
Dim nCount As Long
Dim fX As Single, fY As Single, FS As Single
Dim d As Single, l As Single
Dim bHQ As Boolean, b As Boolean
Dim nXCount As Long, nYCount As Long
Dim nXNow As Long, nYNow As Long
Dim TheSeed As Long
'Dim TheSeed2(255) As Long
Dim ls As typeArgTwoL
'///string replace algorithm
Dim TheString(1) As typeArgTwoL_String
Dim nIndex(31) As Long
'///
Dim nCur As Long 'current generation
Dim nBranch As Long 'current branch
Dim nDeletedBranch As Long
Dim tState() As typeLSystemState 'state stack
Dim curState As typeLSystemState
Dim tPolyPt() As typeLSystemPoint 'polygon point stack
Dim tPoly() As Long 'polygon point count stack
Dim nPolyPt As Long 'polygon point
Dim nPoly As Long 'polygon index
'///
Dim i As Long, j As Long
Dim op As typeArgZeroL_Operation
Dim idx As Long 'current string index = 0 or 1
Dim lp As Long ', lp2 As Long
'Dim f As Long
Dim hbr As Long, hpn As Long
'///
'get data
nXCount = bProps(0)
nYCount = (nXCount And 224&) \ 32&
bHQ = (nXCount And 1&)
nXCount = (nXCount And 28&) \ 4&
pGetSeed bProps, 1, TheSeed
CopyMemory fX, bProps(3), 4&
CopyMemory fY, bProps(7), 4&
CopyMemory FS, bProps(11), 4&
CopyMemory ls.fConst(0), bProps(15), 8&
nCount = bProps(23)
'convert data
FS = FS * 二π
d = d * 二π
ls.fConst(0) = ls.fConst(0) * 二π
ls.fConst(1) = ls.fConst(1) * bmOut.Width
ls.fConst(2) = π
'get color
pGetColor sProps, TheClrTable, TheClrCount
'//////////////////////string replace algorithm
'compile?
pCompile_ArgTwoL v, ls
If ls.nCount = 0 Or nXCount = 0 Or nYCount = 0 Then Exit Sub
'init array
Dim ox As Long, oy As Long
TheBitmap = bmOut
With TheSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = bmOut.Height
 .Bounds(1).cElements = bmOut.Width
 .pvData = bmOut.lpbm
End With
CopyMemory ByVal VarPtrArray(TheArray()), VarPtr(TheSA), 4&
'init array
ReDim tStack(255), tState(255), tPolyPt(4095), tPoly(255)
With TheString(0)
 .nMax = ls.d(1).nCount + 65536
 ReDim .d(.nMax - 1)
End With
With TheString(1)
 .nMax = TheString(0).nMax
 ReDim .d(.nMax - 1)
End With
'init color cache
TheClrIndex = &H80000000
i = -1
CopyMemory TheClr, i, 4&
'start draw
For nXNow = 0 To nXCount - 1
 For nYNow = 0 To nYCount - 1
  nBranch = 0
  nDeletedBranch = &H7FFFFFFF
  nPolyPt = 0
  nPoly = -1
  With curState
   .x = fX + nXNow / nXCount
   If .x > 1 Then .x = .x - 1
   .x = .x * bmOut.Width
   .y = fY + nYNow / nYCount
   If .y > 1 Then .y = .y - 1
   .y = .y * bmOut.Height
   .FS = FS
   .idxClr = 0
   .idxWidth = 0
  End With
  idx = nCount And 1&
  'init string
  With TheString(idx)
   .nCount = ls.d(1).nCount
   If .nCount > 0 Then pArgTwoL_Apply .d, 0, .nCount, ls.d(1).d, ls.fConst, TheString(1 - idx).d, nIndex
  End With
  'start replace
  For nCur = 1 To nCount
   TheString(1 - idx).nCount = 0
   With TheString(idx)
    For lp = 0 To .nCount - 1
     'find expand
     j = 0
     op = .d(lp)
     With ls
      i = .nTable(op.nType, op.nIndex)
      Do Until i = 0
       With .d(i)
        If op.nArgCount = .nArgCount Then
         'match left context
         If .LeftContextCount > 0 Then
          b = pMatch_ArgTwoL_LeftContext(.LeftContext, .LeftContextCount, ls.bIsIgnored, TheString(idx), lp - 1, nIndex)
         Else
          b = True
         End If
         If b And .RightContextCount > 0 Then
          'match right context
          b = pMatch_ArgTwoL_RightContext(.RightContext, .RightContextCount, ls.bIsIgnored, TheString(idx), lp + 1, nIndex)
         End If
         nIndex(0) = lp
         If b And .tExp.nCount > 0 Then
          If pCalc2(.tExp, ls.fConst, TheString(idx).d, nIndex) = 0 Then b = False
         End If
        Else
         b = False
        End If
       End With
       If b Then
        j = i
        'TODO:probability
        Exit Do
       End If
       i = .d(i).idxNext
      Loop
     End With
     If j > 0 Then 'expand it
      i = ls.d(j).nCount
      If i > 0 Then
       With TheString(1 - idx)
        .nCount = .nCount + i
        If .nCount >= .nMax Then
         .nMax = .nMax + i + 65536
         ReDim Preserve .d(.nMax - 1)
        End If
        pArgTwoL_Apply .d, .nCount - i, i, ls.d(j).d, ls.fConst, TheString(idx).d, nIndex
       End With
      End If
     Else 'just keep unchanged
      With TheString(1 - idx)
       .nCount = .nCount + 1
       If .nCount >= .nMax Then
        .nMax = .nMax + 65536
        ReDim Preserve .d(.nMax - 1)
       End If
       .d(.nCount - 1) = op
      End With
     End If
    Next lp
   End With
   idx = 1 - idx
  Next nCur
  'start draw
  With TheString(0)
   For nCur = 0 To .nCount - 1
    op = .d(nCur)
    Select Case op.nType
    Case 70, 71, 102, 103 'F,G,f,g
     If nBranch < nDeletedBranch Then
      With curState
       If op.nType < 80 Then
        '////////get color!!!
        If .idxClr <> TheClrIndex And TheClrCount > 0 Then
         TheClrIndex = .idxClr
         If TheClrIndex < 0 Then
          TheClr = TheClrTable(0)
         ElseIf TheClrIndex >= TheClrCount Then
          TheClr = TheClrTable(TheClrCount - 1)
         Else
          TheClr = TheClrTable(TheClrIndex)
         End If
        End If
        #If UseLineDDA Then
        #Else
        If TheClrCount > 0 Then
         With TheClr
          i = .rgbRed + .rgbGreen * &H100& + .rgbBlue * &H10000
         End With
        Else
         i = vbWhite
        End If
        #End If
        '////////
        #If UseLineDDA Then
        ox = .x
        oy = .y
        #Else
        MoveToEx bmOut.hdc, .x, .y, ByVal 0
        #End If
       End If
       If op.nArgCount > 0 Then
        .x = .x + op.fArg(0) * Cos(.FS)
        .y = .y + op.fArg(0) * Sin(.FS)
       Else
        .x = .x + ls.fConst(1) * Cos(.FS)
        .y = .y + ls.fConst(1) * Sin(.FS)
       End If
       If op.nType < 80 Then
        #If UseLineDDA Then
        If TheClr.rgbReserved > 0 Then
         If TheClr.rgbReserved = 255 Then
          LineDDA ox, oy, .x, .y, AddressOf pDrawLine, 0
         Else
          LineDDA ox, oy, .x, .y, AddressOf pDrawLineAlpha, 0
         End If
        End If
        #Else
        'ERROR!!! API didn't support 32-bit bitmap
        hpn = SelectObject(bmOut.hdc, CreatePen(0, 1, i))
        LineTo bmOut.hdc, .x, .y
        DeleteObject SelectObject(bmOut.hdc, hpn)
        #End If
       End If
       'record vertex
       If nPoly >= 0 And (op.nType And 1) = 0 Then
        tPoly(nPoly) = tPoly(nPoly) + 1
        With tPolyPt(nPolyPt)
         .x = curState.x
         .y = curState.y
        End With
        nPolyPt = nPolyPt + 1
       End If
      End With
     End If
    Case 46 '.
     'record vertex
     If nBranch < nDeletedBranch And nPoly >= 0 Then
      tPoly(nPoly) = tPoly(nPoly) + 1
      With tPolyPt(nPolyPt)
       .x = curState.x
       .y = curState.y
      End With
      nPolyPt = nPolyPt + 1
     End If
    Case 91 '[
     tState(nBranch) = curState
     nBranch = nBranch + 1
    Case 93 ']
     If nBranch > 0 Then nBranch = nBranch - 1
     If nBranch < nDeletedBranch Then nDeletedBranch = &H7FFFFFFF
     curState = tState(nBranch)
    Case 123 '{
     nPoly = nPoly + 1
     tPoly(nPoly) = 0
    Case 125 '}
     If nPoly >= 0 Then
      j = tPoly(nPoly)
      nPoly = nPoly - 1
      If j > 0 Then
       With curState
        '////////get color!!!
        If .idxClr <> TheClrIndex And TheClrCount > 0 Then
         TheClrIndex = .idxClr
         If TheClrIndex < 0 Then
          TheClr = TheClrTable(0)
         ElseIf TheClrIndex >= TheClrCount Then
          TheClr = TheClrTable(TheClrCount - 1)
         Else
          TheClr = TheClrTable(TheClrIndex)
         End If
        End If
        #If UseLineDDA Then
        #Else
        If TheClrCount > 0 Then
         With TheClr
          i = .rgbRed + .rgbGreen * &H100& + .rgbBlue * &H10000
         End With
        Else
         i = vbWhite
        End If
        #End If
        '////////
       End With
       nPolyPt = nPolyPt - j
       If nBranch < nDeletedBranch Then
        #If UseLineDDA Then
        pDrawPolygon tPolyPt, nPolyPt, j
        #Else
        pDrawPolygonTest tPolyPt, nPolyPt, j, i
        #End If
       End If
      End If
     End If
    Case 43 '+
     If nBranch < nDeletedBranch Then
      With curState
       If op.nArgCount > 0 Then
        .FS = .FS - op.fArg(0) '??? rad?? deg??
       Else
        .FS = .FS - ls.fConst(0)
       End If
       If .FS < 0 Then .FS = .FS + 二π
      End With
     End If
    Case 45 '-
     If nBranch < nDeletedBranch Then
      With curState
       If op.nArgCount > 0 Then
        .FS = .FS + op.fArg(0) '??? rad?? deg??
       Else
        .FS = .FS + ls.fConst(0)
       End If
       If .FS > 二π Then .FS = .FS - 二π
      End With
     End If
    Case 124 '|
     If nBranch < nDeletedBranch Then
      With curState
       If .FS < π Then .FS = .FS + π Else .FS = .FS - π
      End With
     End If
    Case 33 '!
     If nBranch < nDeletedBranch Then
      With curState
       If op.nArgCount > 0 Then
        .idxWidth = .idxWidth + op.fArg(0)
       Else
        .idxWidth = .idxWidth + 1 '??? TODO:
       End If
      End With
     End If
    Case 39 ''
     If nBranch < nDeletedBranch Then
      With curState
       If op.nArgCount > 0 Then
        .idxClr = .idxClr + op.fArg(0)
       Else
        .idxClr = .idxClr + 1 '??? TODO:
       End If
      End With
     End If
    Case 37 '%
     nDeletedBranch = nBranch
    End Select
   Next nCur
  End With
 Next nYNow
Next nXNow
'destroy
ZeroMemory ByVal VarPtrArray(TheArray()), 4&
End Sub


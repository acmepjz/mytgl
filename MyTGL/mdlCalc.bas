Attribute VB_Name = "mdlCalc"
Option Explicit

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

Private Type RGBQUAD
    rgbBlue As Byte
    rgbGreen As Byte
    rgbRed As Byte
    rgbReserved As Byte
End Type

'/////////////////////some type

Public Type typeAlphaDibSectionDescriptor
 hdc As Long
 Width As Long
 Height As Long
 lpbm As Long
End Type

Public Type typeIFSPTransform '32 bytes
 f(5) As Single
 nClr As Long
 f2 As Single 'and IsSolid !!!
End Type

Private Type typeCellPoint '16 bytes
 x As Long '0-4095
 y As Long '0-4095
 nRnd As Integer   '0-255
 bNoColor As Boolean
 idxNext As Long 'linked-list
End Type

'四元数
Private Type typeQuat
 w As Single
 x As Single
 y As Single
 z As Single
End Type

Private Const π As Double = 3.14159265358979
Private Const 二π As Double = π * 2
Private Const 二分之π As Double = π / 2
Private Const 二分之三π As Double = π * 3 / 2
Private Const 四分之π As Double = π / 4

'standard lighthess = 0.071B+0.707G+0.222R

'must be 1-instance!!! or the ASM thunk may be erased then ERROR
Public cUnk As New clsUnknown

'even more stupid...
Private m_bIsInIDE As Boolean

Private Sub pQuatAdd2(nRet As typeQuat, n As typeQuat)
With nRet
 .w = .w + n.w
 .x = .x + n.x
 .y = .y + n.y
 .z = .z + n.z
End With
End Sub

Private Sub pQuatMul3(nRet As typeQuat, n1 As typeQuat, n2 As typeQuat)
With n1
 nRet.w = .w * n2.w - .x * n2.x - .y * n2.y - .z * n2.z
 nRet.x = .w * n2.x + .x * n2.w + .y * n2.z - .z * n2.y '?
 nRet.y = .w * n2.y + .y * n2.w + .z * n2.x - .x * n2.z '?
 nRet.z = .w * n2.z + .z * n2.w + .x * n2.y - .y * n2.x '?
End With
End Sub

Private Sub pQuatMulNum(nRet As typeQuat, ByVal n As Single)
With nRet
 .w = .w * n
 .x = .x * n
 .y = .y * n
 .z = .z * n
End With
End Sub

'fw=phi/2
Private Sub pQuatFromRotation(nRet As typeQuat, ByVal fw As Single, ByVal fX As Single, ByVal fY As Single, ByVal fz As Single)
With nRet
 .w = Cos(fw)
 .z = fX * fX + fY * fY + fz * fz
 If .z > 0 Then
  .z = Sin(fw) / Sqr(.z)
  .x = fX * .z
  .y = fY * .z
  .z = fz * .z
 Else
  .x = 0
  .y = 0
  .z = 0
 End If
End With
End Sub

'only normalize x,y,z
Private Sub pQuatNormalize3(nRet As typeQuat)
Dim f As Single
With nRet
 f = .w * .w + .x * .x + .y * .y + .z * .z
 If f > 0 Then
  f = 1 / Sqr(f)
  '.w = .w * f
  .x = .x * f
  .y = .y * f
  .z = .z * f
 End If
End With
End Sub

''test
'Public Sub TestRnd()
'Dim i As Long, j As Long
'Dim n(3) As Long
'cUnk.InitASM
'For i = 1 To 10000
' j = cUnk.fRnd2(i, 0, 0) And &H3&
' n(j) = n(j) + 1
'Next i
'For i = 0 To 3
' Debug.Print n(i),
'Next i
'End Sub

Private Function pCheckIDE() As Boolean
m_bIsInIDE = True
pCheckIDE = True
End Function

'0-based
Public Function CalcOperator(bmOut As cAlphaDibSection, ByVal nCount As Long, bmIn() As typeAlphaDibSectionDescriptor, ByVal nOpType As Long, bProps() As Byte, sProps() As String) As Boolean
Dim w As Long, h As Long
Dim d As typeAlphaDibSectionDescriptor
'stupid operation
Debug.Assert pCheckIDE
'////////////////
If Not ValidateOperator(nCount, bmIn, nOpType, bProps, sProps, w, h) Then Exit Function
If w <> bmOut.Width Or h <> bmOut.Height Then
 bmOut.Create w, -h
End If
Select Case nOpType
'////////////////////////////////
'
'  Generators
'
'////////////////////////////////
Case 1 'flat
 pCalcFlat bmOut.DIBSectionBitsPtr, w * h, bProps
Case 2 'cloud
 pCalcCloud bmOut.DIBSectionBitsPtr, w, h, bProps
Case 3 'gradient
 pCalcGradient bmOut.DIBSectionBitsPtr, w, h, bProps
Case 4 'gradient2
 pCalcGradient2 bmOut.DIBSectionBitsPtr, w, h, bProps
Case 5 'cell
 pCalcCell bmOut.DIBSectionBitsPtr, w, h, bProps
Case 6 'noise
 pCalcNoise bmOut.DIBSectionBitsPtr, w * h, bProps
Case 7 'brick
 pCalcBrick bmOut.DIBSectionBitsPtr, w, h, bProps
Case 8 'perlin
 pCalcPerlin bmOut.DIBSectionBitsPtr, w, h, bProps
Case 9 'import
 'TODO:
'////////////////////////////////
'
'  End of generators
'
'////////////////////////////////
Case 11 'SlowGrow
 pCalcSlowGrow bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w, h, bProps
Case 12 'L-system
 'test only
 d.hdc = bmOut.hdc
 d.Width = w
 d.Height = h
 d.lpbm = bmOut.DIBSectionBitsPtr
 CopyMemory ByVal d.lpbm, ByVal bmIn(0).lpbm, 4& * w * h
 CalcLSystemTest d, bProps, sProps
Case 13 'IFSP
 pCalcIFSP bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w, h, bProps, sProps(0)
Case 14 'rect
 pCalcRect bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w, h, bProps
Case 15 'pixels
 pCalcPixels bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w * h, bProps
Case 16 'glowrect
 pCalcGlowRect bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w, h, bProps
Case 17 'crack
 pCalcCrack bmOut.DIBSectionBitsPtr, nCount, bmIn, w, h, bProps
Case 20 'blur
 pCalcBlur bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w, h, bProps
Case 21 'color
 pCalcAddColor bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w * h, bProps
Case 22 'range
 pCalcRangeColor bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w * h, bProps
Case 23 'HSCB
 pCalcHSCB bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w * h, bProps
Case 24 'normals
 pCalcNormals bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w, h, bProps
Case 25 'colorbalance
 pCalcColorBalance bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w * h, bProps
Case 26 'rotzoom
 pCalcRotZoom bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w, h, bmIn(0).Width, bmIn(0).Height, bProps
Case 27 'rotatemul
 pCalcRotMul bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w, h, bProps
Case 28 'sharpen
 pCalcSharpen bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w, h, bProps
Case 29 'dialect
 pCalcDialect bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w, h, bProps
Case 34 'distort
 pCalcDistort bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, bmIn(1).lpbm, w, h, bProps
Case 35 'bump
 pCalcBump bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, bmIn(1).lpbm, w, h, bProps
Case 36 'add
 pCalcAddBitmap bmOut.DIBSectionBitsPtr, w * h, nCount, bmIn, bProps
Case 37 'mask
 pCalcMaskBitmap bmOut.DIBSectionBitsPtr, w * h, bmIn, bProps
Case 38 'particle
 If nCount > 2 Then
  pCalcParticle bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, bmIn(1).lpbm, bmIn(2).lpbm, w, h, bProps
 Else
  pCalcParticle bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, bmIn(1).lpbm, 0, w, h, bProps
 End If
Case 39 'segment
 pCalcSegment bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, bmIn(1).lpbm, w, h, bProps
Case 40 'bulge
 pCalcBulge bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w, h, bProps
Case 41 'twirl
 pCalcTwirl bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w, h, bProps
Case 42 'unwrap
 pCalcUnwrap bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, w, h, bProps
Case 43 'abnormals
 If nCount > 1 Then
  pCalcAbnormals bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, bmIn(1).lpbm, w * h, bProps
 Else
  pCalcAbnormals bmOut.DIBSectionBitsPtr, bmIn(0).lpbm, 0, w * h, bProps
 End If
Case Else
 'TODO:the other...
End Select
CalcOperator = True
End Function

'0-based
Public Function ValidateOperator(ByVal nCount As Long, bmIn() As typeAlphaDibSectionDescriptor, ByVal nOpType As Long, bProps() As Byte, sProps() As String, NewWidth As Long, NewHeight As Long) As Boolean
Dim i As Long
cUnk.InitASM
'validate test
Debug.Assert pTestBitmap(nCount, bmIn)
Select Case nOpType
Case 1 To int_Generator_Max 'generators?
 If nCount > 0 Then Exit Function
 NewWidth = cUnk.fShl(1, bProps(0) And &HF&)
 NewHeight = cUnk.fShl(1, (bProps(0) And &HF0&) \ &H10&)
Case 26 'rotzoom
 If nCount <> 1 Then Exit Function
 i = bProps(25)
 If i And &HF& Then
  NewWidth = cUnk.fShl(1, (i And &HF&) - 1)
 Else
  NewWidth = bmIn(0).Width 'default
 End If
 If i And &HF0& Then
  NewHeight = cUnk.fShl(1, ((i And &HF0&) \ &H10&) - 1)
 Else
  NewHeight = bmIn(0).Height 'default
 End If
 'TODO:L-system allows multiple inputs and SlowGrow????
Case 11 To 16, 20 To 25, 27 To 29, 40 To 42
 'SlowGrow,L-system,IFSP,rect,pixels,glowrect
 'blur,color,range,HSCB,normals,color balance
 'rotatemul,sharpen,dialect
 'bulge,twirl,unwrap
 If nCount <> 1 Then Exit Function
 NewWidth = bmIn(0).Width
 NewHeight = bmIn(0).Height
Case 38 'particle
 If nCount < 2 Or nCount > 3 Then Exit Function
 NewWidth = bmIn(0).Width
 NewHeight = bmIn(0).Height
 For i = 1 To nCount - 1
  If bmIn(i).Width <> NewWidth Or bmIn(i).Height <> NewHeight Then Exit Function
 Next i
Case 17, 43 'crack,abnormals
 If nCount < 1 Or nCount > 2 Then Exit Function
 NewWidth = bmIn(0).Width
 NewHeight = bmIn(0).Height
 For i = 1 To nCount - 1
  If bmIn(i).Width <> NewWidth Or bmIn(i).Height <> NewHeight Then Exit Function
 Next i
Case 34, 35, 39 'distort,bump,segment
 If nCount <> 2 Then Exit Function
 NewWidth = bmIn(0).Width
 NewHeight = bmIn(0).Height
 If bmIn(1).Width <> NewWidth Or bmIn(1).Height <> NewHeight Then Exit Function
Case 36 'add
 If nCount < 1 Then Exit Function
 NewWidth = bmIn(0).Width
 NewHeight = bmIn(0).Height
 For i = 1 To nCount - 1
  If bmIn(i).Width <> NewWidth Or bmIn(i).Height <> NewHeight Then Exit Function
 Next i
Case 37 'mask
 If nCount <> 3 Then Exit Function
 NewWidth = bmIn(0).Width
 NewHeight = bmIn(0).Height
 For i = 1 To 2
  If bmIn(i).Width <> NewWidth Or bmIn(i).Height <> NewHeight Then Exit Function
 Next i
Case Else
 'TODO:the other...
 Exit Function
End Select
ValidateOperator = True
End Function

'///////////////////////////////test function

Private Function pTestBitmap(ByVal nCount As Long, bmIn() As typeAlphaDibSectionDescriptor) As Boolean
Dim i As Long
pTestBitmap = True
For i = 0 To nCount - 1
 With bmIn(i)
  pTestBitmap = pTestBitmap And .hdc <> 0 And .Width > 0 And .Height > 0 And .lpbm <> 0
 End With
Next i
End Function

'///////////////////////////////private functions

Private Sub pCalcFlat(ByVal lpbm As Long, ByVal m As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim i As Long, clr As RGBQUAD
'init array
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = m
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
'get color
CopyMemory clr, bProps(1), 4&
For i = 0 To m - 1
 bDib(i) = clr
Next i
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcRect(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim i As Long, j As Long
Dim nLeft As Long, nTop As Long, nRight As Long, nBottom As Long
Dim clr As RGBQUAD
Dim nMode As Long
Dim f(3) As Single
'init array
CopyMemory ByVal lpbm, ByVal lpbmIn, 4& * w * h
With tSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
'get color
CopyMemory clr, bProps(0), 4&
'draw rect
CopyMemory f(0), bProps(4), 16&
If f(0) <= 0 Then nLeft = 0 Else nLeft = w * f(0)
If f(1) <= 0 Then nTop = 0 Else nTop = h * f(1)
If f(2) >= 1 Then nRight = w Else nRight = w * f(2)
If f(3) >= 1 Then nBottom = h Else nBottom = h * f(3)
nMode = bProps(20) And &H3&
'0-blend
'1-mix
If nMode = 0 Then
 If clr.rgbReserved = 255& Then
  For j = nTop To nBottom - 1
   For i = nLeft To nRight - 1
    bDib(i, j) = clr
   Next i
  Next j
 Else
  For j = nTop To nBottom - 1
   For i = nLeft To nRight - 1
    pBlendColor bDib(i, j), clr
   Next i
  Next j
 End If
ElseIf nMode = 1 And clr.rgbReserved > 0 Then
 For j = nTop To nBottom - 1
  For i = nLeft To nRight - 1
   pMixColor2 bDib(i, j), clr
  Next i
 Next j
End If
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcGlowRect(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim i As Long, j As Long, k As Long, lp As Long
Dim x As Long, y As Long
Dim nLeft As Long, nTop As Long, nRight As Long, nBottom As Long
Dim nXSize As Long, nYSize As Long
Dim clr As RGBQUAD
Dim bWrap As Boolean
Dim f(5) As Single
Dim TheTable(255) As Long
Dim TheClrTable(255) As RGBQUAD
Dim nBlend As Long
Dim nTemp() As Long
'init array
CopyMemory ByVal lpbm, ByVal lpbmIn, 4& * w * h
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = w * h
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
ReDim nTemp(w - 1, h - 1)
'get color
CopyMemory clr, bProps(24), 4&
'get size
CopyMemory f(0), bProps(0), 24&
nLeft = (f(0) - f(4)) * w
nTop = (f(1) - f(5)) * h
nRight = (f(0) + f(4)) * w
nBottom = (f(1) + f(5)) * h
nXSize = f(2) * w
nYSize = f(3) * h
'init table
CopyMemory f(0), bProps(28), 8&
f(2) = f(0) * 256#
nBlend = f(2)
For i = 0 To 255
 If bProps(36) And 2& Then
  '???
  TheTable(i) = f(2) * (1 - (1 - i / 255#) ^ f(1))
 Else
  TheTable(i) = f(2) * ((i / 255#) ^ f(1))
 End If
 With TheClrTable(i)
  .rgbBlue = (clr.rgbBlue * i) \ 255&
  .rgbGreen = (clr.rgbGreen * i) \ 255&
  .rgbRed = (clr.rgbRed * i) \ 255&
  .rgbReserved = (clr.rgbReserved * i) \ 255&
 End With
Next i
'calc rect
bWrap = bProps(36) And 1&
'inside the rect
If bWrap Then
 For j = nTop To nBottom - 1
  For i = nLeft To nRight - 1
   x = i And (w - 1)
   y = j And (h - 1)
   nTemp(x, y) = nTemp(x, y) + nBlend
  Next i
 Next j
Else
 If nLeft < 0 Then nLeft = 0
 If nTop < 0 Then nTop = 0
 If nRight > w Then nRight = w
 If nBottom > h Then nBottom = h
 For j = nTop To nBottom - 1
  For i = nLeft To nRight - 1
   nTemp(i, j) = nBlend
  Next i
 Next j
End If
'horizontal glow
For j = 0 To nYSize - 1
 k = (j * 256&) \ nYSize
 If bWrap Then
  y = (nTop - nYSize + j) And (h - 1)
  For i = nLeft To nRight - 1
   x = i And (w - 1)
   nTemp(x, y) = nTemp(x, y) + TheTable(k)
  Next i
  y = (nBottom - 1 + nYSize - j) And (h - 1)
  For i = nLeft To nRight - 1
   x = i And (w - 1)
   nTemp(x, y) = nTemp(x, y) + TheTable(k)
  Next i
 Else
  y = (nTop - nYSize + j)
  If y >= 0 Then
   For i = nLeft To nRight - 1
    nTemp(i, y) = TheTable(k)
   Next i
  End If
  y = (nBottom - 1 + nYSize - j)
  If y < h Then
   For i = nLeft To nRight - 1
    nTemp(i, y) = TheTable(k)
   Next i
  End If
 End If
Next j
'vertical glow
For i = 0 To nXSize - 1
 k = (i * 256&) \ nXSize
 If bWrap Then
  x = (nLeft - nXSize + i) And (w - 1)
  For j = nTop To nBottom - 1
   y = j And (h - 1)
   nTemp(x, y) = nTemp(x, y) + TheTable(k)
  Next j
  x = (nRight - 1 + nXSize - i) And (w - 1)
  For j = nTop To nBottom - 1
   y = j And (h - 1)
   nTemp(x, y) = nTemp(x, y) + TheTable(k)
  Next j
 Else
  x = (nLeft - nXSize + i)
  If x >= 0 Then
   For j = nTop To nBottom - 1
    nTemp(x, j) = TheTable(k)
   Next j
  End If
  x = (nRight - 1 + nXSize - i)
  If x < w Then
   For j = nTop To nBottom - 1
    nTemp(x, j) = TheTable(k)
   Next j
  End If
 End If
Next i
'corner glow
'////////////STUPID CALC SQRT
Dim nSqrt0 As Long, nSqrt1 As Long
Dim nNum0 As Long
'actual number=nSqrt^2+nNum
Dim nNumDelta0 As Long
Dim ε As Long, ε…2 As Long
Dim nYSize…2 As Long, nXYSize As Long
''////////////
nXYSize = nXSize * nYSize
nYSize…2 = nYSize * nYSize
ε = nXYSize \ 512& 'if nXSize and nYSize >1000 ... error
If ε <= 0 Then ε = 1
ε…2 = ε * ε
For j = 1 To nYSize
 'actual number=(j*nXSize)^2
 'nSqrt0=j*nXSize
 nSqrt0 = nSqrt0 + nXSize
 nNum0 = 0
 nSqrt1 = nSqrt0
 nNumDelta0 = 0
 lp = ε * (nSqrt0 + nSqrt0 + ε) '??
 For i = 1 To nXSize
  'delta= nYSize^2 *(i^2- (i-1)^2 )
  '     = ........ *((i-1) + i)
  nNum0 = nNum0 + nNumDelta0
  nNumDelta0 = nNumDelta0 + nYSize…2
  nNum0 = nNum0 + nNumDelta0
  '   if actualNumber>=(nSqrt+ε)^2
  '=> if nSqrt^2+nNum>=nSqrt^2+2*nSqrt*ε+ε^2
  '=> if         nNum>=ε(2*nSqrt+ε)
  Do Until nNum0 < lp _
  Or lp < 0 'error!!! to prevent a dead loop
   nNum0 = nNum0 - lp
   lp = lp + ε…2 + ε…2
   nSqrt1 = nSqrt1 + ε
  Loop
  'we got a result!!
  If nSqrt1 >= nXYSize Then Exit For
  If nSqrt1 >= 8388608 Then
   k = nSqrt1 \ (nXYSize \ 256&)
   If k > 255 Then k = 255
  Else
   k = (nSqrt1 * 256&) \ nXYSize
  End If
  k = TheTable(255 - k)
  'change data
  If bWrap Then
   y = (nTop - j) And (h - 1)
   nBlend = (nBottom - 1 + j) And (h - 1)
   x = (nLeft - i) And (w - 1)
   nTemp(x, y) = nTemp(x, y) + k
   nTemp(x, nBlend) = nTemp(x, nBlend) + k
   x = (nRight - 1 + i) And (w - 1)
   nTemp(x, y) = nTemp(x, y) + k
   nTemp(x, nBlend) = nTemp(x, nBlend) + k
  Else
   y = (nTop - j)
   nBlend = (nBottom - 1 + j)
   x = (nLeft - i)
   If x >= 0 Then
    If y >= 0 Then nTemp(x, y) = k
    If nBlend < h Then nTemp(x, nBlend) = k
   End If
   x = (nRight - 1 + i)
   If x < w Then
    If y >= 0 Then nTemp(x, y) = k
    If nBlend < h Then nTemp(x, nBlend) = k
   End If
  End If
 Next i
Next j
'apply color
lp = 0
For y = 0 To h - 1
 For x = 0 To w - 1
  'blend color
  i = 256& - nTemp(x, y)
  If i <> 256 Then
   If i < 256 Then
    If i < 0 Then i = 0
    With TheClrTable(255 - i)
     nLeft = .rgbBlue
     nTop = .rgbGreen
     nRight = .rgbRed
     nBottom = .rgbReserved
    End With
   Else
    If i > 512 Then i = 512
    With TheClrTable(i - 257) '????
     nLeft = -.rgbBlue
     nTop = -.rgbGreen
     nRight = -.rgbRed
     nBottom = .rgbReserved
    End With
   End If
   With bDib(lp)
    j = nLeft + (i * .rgbBlue) \ 256&
    If j < 0 Then j = 0 Else If j > 255 Then j = 255
    .rgbBlue = j
    j = nTop + (i * .rgbGreen) \ 256&
    If j < 0 Then j = 0 Else If j > 255 Then j = 255
    .rgbGreen = j
    j = nRight + (i * .rgbRed) \ 256&
    If j < 0 Then j = 0 Else If j > 255 Then j = 255
    .rgbRed = j
    j = nBottom + (i * .rgbReserved) \ 256&
    If j < 0 Then j = 0 Else If j > 255 Then j = 255
    .rgbReserved = j
   End With
  End If
  lp = lp + 1
 Next x
Next y
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcCrack(ByVal lpbm As Long, ByVal nCount As Long, bmIn() As typeAlphaDibSectionDescriptor, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Const TheConst As Double = π / 512
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim bDibSrc() As RGBQUAD
Dim tSASrc As SAFEARRAY2D
Dim i As Long, j As Long, k As Long
Dim ii As Long, jj As Long
Dim x As Long, y As Long '1 -> 256
Dim nAngle As Long '2π -> 1024
Dim nClr As Long, clr As RGBQUAD
Dim nXMax As Long, nYMax As Long
Dim nIndex As Long, nVar As Long, nClrVar As Long
Dim nMode As Long, nMode2 As Long, bHQ As Boolean, bDraw As Boolean
Dim clr1 As RGBQUAD, clr2 As RGBQUAD
Dim TheClrTable(255) As RGBQUAD
Dim TheSin(1023) As Long '1 -> 256
Dim TheFuncTable(255) As Byte
Dim TheSeed As Long
'init array
CopyMemory ByVal lpbm, ByVal bmIn(0).lpbm, 4& * w * h
With tSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
If nCount > 1 Then
 With tSASrc
  .cbElements = 4
  .cDims = 2
  .Bounds(0).cElements = h
  .Bounds(1).cElements = w
  .pvData = bmIn(1).lpbm
 End With
 CopyMemory ByVal VarPtrArray(bDibSrc()), VarPtr(tSASrc), 4&
End If
'get color
CopyMemory clr1, bProps(0), 4&
CopyMemory clr2, bProps(4), 4&
For i = 0 To 255
 pMixColor TheClrTable(i), clr1, clr2, i
Next i
'init table
For i = 0 To 1023
 TheSin(i) = 256# * Sin(TheConst * i)
Next i
For i = 0 To 255
 TheFuncTable(i) = Atn(i / (256# - i)) / TheConst
Next i
'draw bitmap
nXMax = w * 256& - 1
nYMax = h * 256& - 1
pGetSeed bProps, 11, TheSeed
nIndex = bProps(8)
nVar = bProps(9)
nClrVar = bProps(14)
nMode2 = bProps(13)
nMode = (nMode2 And 12&) \ 4& 'mode
bHQ = nMode2 And 128&
nMode2 = nMode2 And 3& 'length decision
Do While nIndex > 0
 nIndex = nIndex - 1
 x = cUnk.fRnd2(nIndex, &HBEE&, TheSeed) And nXMax
 y = cUnk.fRnd2(nIndex, &HF00D&, TheSeed) And nYMax
 nAngle = cUnk.fRnd2(nIndex, &HFADE&, TheSeed) And 1023&
 nClr = cUnk.fRnd2(nIndex, &HFADE0FF, TheSeed) And 255&
 i = bProps(10) 'length
 If nMode2 = 2 And nCount > 1 Then 'normal based
  With bDibSrc(x \ 256&, y \ 256&)
   j = .rgbBlue - 128&
   k = .rgbGreen - 128&
   j = (i * (j * j + k * k)) \ 32& '??
   If j > i + i Then i = i + i Else i = j
  End With
 ElseIf nMode2 = 0 Then 'random
  i = (i * (cUnk.fRnd2(nIndex, &HBE4&, TheSeed) And &H1FF&)) \ 256&
 End If
 For i = 1 To i ':-3
  'calc direction
  nAngle = (nAngle + (nVar * ((cUnk.fRnd2(nIndex, i, TheSeed) And 2047&) - 1024&) + 256&) \ 512&) And 1023&
  'calc color
  nClr = nClr + (nClrVar * ((cUnk.fRnd2(nIndex, i + 10220&, TheSeed) And &H1FF&) - 256&) + 128&) \ 256&
  If nClr < 0 Then nClr = 0 Else If nClr > 255 Then nClr = 255
  'calc position
  x = (x + TheSin((nAngle + 256&) And 1023&)) And nXMax 'cos
  y = (y + TheSin(nAngle)) And nYMax 'sin
  ii = x \ 256& 'actual X
  jj = y \ 256& 'actual Y
  bDraw = True
  If nCount > 1 Then 'apply normal map
   With bDibSrc(ii, jj)
    'get vector
    'j = 128& - .rgbBlue 'x
    j = .rgbBlue - 128& 'x
    k = .rgbGreen - 128& 'y
    'calc ATAN2
    'we use y/(x+y) (<=255) instead of y/x=tan(a)
    If j > 0 Then
     If k = 0 Then
      nAngle = 0
     ElseIf k > 0 Then 'I
      nAngle = TheFuncTable((k * 256&) \ (j + k))
     Else 'IV
      nAngle = -TheFuncTable((k * 256&) \ (k - j))
     End If
    ElseIf j < 0 Then
     If k = 0 Then
      nAngle = 512&
     ElseIf k > 0 Then 'II
      nAngle = 512& - TheFuncTable((k * 256&) \ (k - j))
     Else 'III
      nAngle = 512& + TheFuncTable((k * 256&) \ (j + k))
     End If
    Else
     If k < 0 Then nAngle = 768& Else nAngle = 256&
    End If
    '// rotate 90 degrees ccw
    nAngle = (nAngle + 256&) And 1023&
    '// alpha-based placement decision
    j = j * j + k * k
    k = (255 - .rgbReserved) \ 4
    If j < k * k Then bDraw = False
   End With
  End If
  If bDraw Then
   j = x And 255&
   k = y And 255&
   If bHQ And (j > 0 Or k > 0) Then 'high quality
    clr = TheClrTable(nClr)
    j = j * k
    pMixColor65536 bDib(ii, jj), clr, nMode, 65536 + j - ((x And 255&) + k) * 256&
    If x And 255& Then
     pMixColor65536 bDib((ii + 1) And (w - 1), jj), clr, nMode, (x And 255&) * 256& - j
     If k > 0 Then
      pMixColor65536 bDib((ii + 1) And (w - 1), (jj + 1) And (h - 1)), clr, nMode, j
     End If
    End If
    If k > 0 Then
     pMixColor65536 bDib(ii, (jj + 1) And (h - 1)), clr, nMode, k * 256& - j
    End If
   Else
    Select Case nMode
    Case 0 'normal
     bDib(ii, jj) = TheClrTable(nClr)
    Case 1 'blend
     clr = TheClrTable(nClr)
     If clr.rgbReserved = 255 Then
      bDib(ii, jj) = clr
     Else
      pBlendColor bDib(ii, jj), clr
     End If
    Case 2 'mix
     clr = TheClrTable(nClr)
     If clr.rgbReserved > 0 Then
      pMixColor2 bDib(ii, jj), clr
     End If
    End Select
   End If
  End If
 Next i
Loop
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
ZeroMemory ByVal VarPtrArray(bDibSrc()), 4&
End Sub

Private Sub pCalcIFSP(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte, s1 As String)
On Error Resume Next
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim i As Long, j As Long, k As Long
Dim ii As Long, jj As Long, x As Long
Dim m As Long, mm As Long
Dim nClr As Long, clr As RGBQUAD
Dim nMode As Long, bHQ As Boolean
Dim clr1 As RGBQUAD, clr2 As RGBQUAD
Dim TheSeed As Long
Dim TheFuncTable() As Single 'probability
Dim f As Single, f2 As Single
Dim xFloat As Single, yFloat As Single
Dim TheTable() As typeIFSPTransform
Dim TheClrTable() As RGBQUAD
Dim TheClr() As Long 'alpha
Dim nIsSolid() As Byte
'init array
CopyMemory ByVal lpbm, ByVal lpbmIn, 4& * w * h
With tSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
'get data
m = LenB(s1) \ 32&
pGetSeed bProps, 4, TheSeed
nMode = bProps(6)
bHQ = nMode And 128&
nMode = nMode And 3&
CopyMemory mm, bProps(0), 4&
If m > 0 Then
 'get transform data
 ReDim TheFuncTable(m - 1), TheTable(m - 1), _
 TheClrTable(m - 1), TheClr(m - 1), nIsSolid(m - 1)
 CopyMemory TheTable(0), ByVal StrPtr(s1), m * 32&
 For i = 0 To m - 1
  With TheTable(i)
   'S=x(0)y(1)-x(1)y(0)
   .f(2) = .f(2) - .f(0)
   .f(3) = .f(3) - .f(1)
   .f(4) = .f(4) - .f(0)
   .f(5) = .f(5) - .f(1)
   f2 = Abs(.f(2) * .f(5) - .f(3) * .f(4))
   If f2 < 0.001 Then f2 = 0.001 'too small?
   CopyMemory TheClrTable(i), .nClr, 4&
   '////process floating-point
   CopyMemory k, .f2, 4&
   j = (k And &H7F800000) \ &H800000
   If j = 0 Then
    If k And &H7FFFFF Then
     nIsSolid(i) = 1
     .f2 = 0
    End If
   ElseIf j >= 1 And j <= 84 Then
    nIsSolid(i) = 1
    j = j + 84
    k = (k And &H807FFFFF) Or (j * &H800000)
    CopyMemory .f2, k, 4&
   End If
   '////
   TheClr(i) = 1024# * .f2 'alpha
  End With
  f = f + f2
  TheFuncTable(i) = f
 Next i
 'normalize
 For i = 0 To m - 1
  TheFuncTable(i) = 2# * TheFuncTable(i) / f - 1
 Next i
 xFloat = 0.01
 yFloat = 0.01
 clr1 = TheClrTable(0)
 'draw bitmap
 For i = 1 To mm + 20&
  f = cUnk.fRnd2Float(i, &H10245, TheSeed) '-1 -> 1
  For ii = 0 To m - 2
   If f <= TheFuncTable(ii) Then Exit For
  Next ii
  Debug.Assert ii < m
  'calc transform
  With TheTable(ii)
   If nIsSolid(ii) Then
    f = (cUnk.fRnd2(i, &H12455, TheSeed) And &H7FFF&) / 32768#
    yFloat = (cUnk.fRnd2(i, &H14501, TheSeed) And &H7FFF&) / 32768#
   Else
    f = xFloat
   End If
   xFloat = .f(0) + .f(2) * f + .f(4) * yFloat
   yFloat = .f(1) + .f(3) * f + .f(5) * yFloat
  End With
  'calc color
  j = TheClr(ii)
  With TheClrTable(ii)
   k = clr1.rgbBlue + ((-clr1.rgbBlue + .rgbBlue) * j) \ 1024&
   If k < 0 Then k = 0 Else If k > 255 Then k = 255
   clr1.rgbBlue = k
   k = clr1.rgbGreen + ((-clr1.rgbGreen + .rgbGreen) * j) \ 1024&
   If k < 0 Then k = 0 Else If k > 255 Then k = 255
   clr1.rgbGreen = k
   k = clr1.rgbRed + ((-clr1.rgbRed + .rgbRed) * j) \ 1024&
   If k < 0 Then k = 0 Else If k > 255 Then k = 255
   clr1.rgbRed = k
   k = clr1.rgbReserved + ((-clr1.rgbReserved + .rgbReserved) * j) \ 1024&
   If k < 0 Then k = 0 Else If k > 255 Then k = 255
   clr1.rgbReserved = k
  End With
  'draw point
  If i > 20 Then
   If bHQ Then
    ii = w * 256&
    ii = CLng(xFloat * ii) And (ii - 1)
    jj = h * 256&
    jj = CLng(yFloat * jj) And (jj - 1)
    j = ii And 255&
    k = jj And 255&
    ii = ii \ 256&
    jj = jj \ 256&
   Else
    ii = CLng(xFloat * w) And (w - 1)
    jj = CLng(yFloat * h) And (h - 1)
   End If
   If bHQ And (j > 0 Or k > 0) Then 'high quality
    clr = TheClrTable(nClr)
    x = j
    j = j * k
    pMixColor65536 bDib(ii, jj), clr1, nMode, 65536 + j - (x + k) * 256&
    If x And 255& Then
     pMixColor65536 bDib((ii + 1) And (w - 1), jj), clr1, nMode, x * 256& - j
     If k > 0 Then
      pMixColor65536 bDib((ii + 1) And (w - 1), (jj + 1) And (h - 1)), clr1, nMode, j
     End If
    End If
    If k > 0 Then
     pMixColor65536 bDib(ii, (jj + 1) And (h - 1)), clr1, nMode, k * 256& - j
    End If
   Else
    Select Case nMode
    Case 0 'normal
     bDib(ii, jj) = clr1
    Case 1 'blend
     If clr.rgbReserved = 255 Then
      bDib(ii, jj) = clr1
     Else
      pBlendColor bDib(ii, jj), clr1
     End If
    Case 2 'mix
     If clr.rgbReserved > 0 Then
      pMixColor2 bDib(ii, jj), clr1
     End If
    End Select
   End If
  End If
 Next i
End If
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcPixels(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal m As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim i As Long, j As Long, k As Long, nCount As Long
Dim TheSeed As Long
Dim nMode As Long
Dim clrs(1) As RGBQUAD, clr As RGBQUAD
Dim TheClrTable(255) As RGBQUAD
'init array
CopyMemory ByVal lpbm, ByVal lpbmIn, 4& * m
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = m
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
'get color
CopyMemory clrs(0), bProps(0), 8&
CopyMemory nCount, bProps(8), 4&
pGetSeed bProps, 12, TheSeed
nMode = bProps(14) And &H3&
'0-normal
'1-blend
'2-mix
'init table
For i = 0 To 255
 pMixColor TheClrTable(i), clrs(0), clrs(1), i
Next i
If nMode = 0 Then
 For i = 1 To nCount
  j = cUnk.fRnd2(i, 71&, TheSeed) And (m - 1)
  k = cUnk.fRnd2(i, 45&, TheSeed) And &HFF&
  bDib(j) = TheClrTable(k)
 Next i
ElseIf nMode = 1 Then
 For i = 1 To nCount
  j = cUnk.fRnd2(i, 71&, TheSeed) And (m - 1)
  k = cUnk.fRnd2(i, 45&, TheSeed) And &HFF&
  clr = TheClrTable(k)
  If clr.rgbReserved = 255& Then
   bDib(j) = clr
  Else
   pBlendColor bDib(j), clr
  End If
 Next i
ElseIf nMode = 2 Then
 For i = 1 To nCount
  j = cUnk.fRnd2(i, 71&, TheSeed) And (m - 1)
  k = cUnk.fRnd2(i, 45&, TheSeed) And &HFF&
  clr = TheClrTable(k)
  If clr.rgbReserved > 0 Then
   pMixColor2 bDib(j), clr
  End If
 Next i
End If
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcCloud(ByVal lpbm As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim clr1 As RGBQUAD, clr2 As RGBQUAD
Dim TheSeed As Long
Dim i As Long, j As Long, lp As Long
Dim nAmount As Long
Dim ww As Long, hh As Long
Dim TheClrTable(255) As RGBQUAD
Dim nTemp() As Long
Dim jj As Long, t1 As Long, t2 As Long, t3 As Long, t4 As Long
'///new
Dim b As Boolean
'///
'init array
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = w * h
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
ReDim nTemp(w - 1, h - 1)
'get color
CopyMemory clr1, bProps(1), 4&
CopyMemory clr2, bProps(5), 4&
'get prop
b = bProps(12) And &H10&
'init table
For i = 0 To 255
 pMixColor TheClrTable(i), clr1, clr2, i
Next i
'randomize
pGetSeed bProps, 10, TheSeed
'create random map
lp = cUnk.fShl(1, bProps(12) And &HF&) 'get level
If lp < w And lp < h Then
 If w > h Then
  ww = (w * lp) \ h
  hh = lp
 Else
  ww = lp
  hh = (h * lp) \ w
 End If
Else
 ww = w
 hh = h
End If
Debug.Assert (ww And (ww - 1)) = 0 And (hh And (hh - 1)) = 0 'is 2^n??
nAmount = &H200&
lp = 0
Do
 'add random
 If nAmount > 0 Then
  For j = 0 To hh - 1
   For i = 0 To ww - 1
    nTemp(i, j) = nTemp(i, j) + ((cUnk.fRnd3(i, j, lp, TheSeed) And &H7FF&) - &H400&) * nAmount
   Next i
  Next j
  nAmount = (nAmount * bProps(9)) \ &H100&
 End If
 'resize and blur
 If ww >= w Or hh >= h Then Exit Do
 If b Then 'NEW:bicubic x=(9/16)(x2+x3)-(1/16)(x1+x4) bug????
  For j = hh - 1 To 0 Step -1
   For i = ww - 1 To 0 Step -1
    nTemp(i + i, j + j) = nTemp(i, j)
   Next i
  Next j
  ww = ww + ww
  hh = hh + hh
  For j = 0 To hh - 2 Step 2
   t1 = nTemp(0, j) 'right
   t2 = nTemp(2 And (ww - 1), j) 'right-right
   jj = (j + 2) And (hh - 1)
   For i = ww - 2 To 0 Step -2
    t3 = nTemp(i, j) 'this
    t4 = t1 + t3
    nTemp(i + 1, j) = (t4 * 8& + t4 - t2 - nTemp((i - 2) And (ww - 1), j)) \ 16&
    t4 = t3 + nTemp(i, jj) 'down
    nTemp(i, j + 1) = (t4 * 8& + t4 - nTemp(i, (j + 4) And (hh - 1)) - nTemp(i, (j - 2) And (hh - 1))) \ 16&
    'move it
    t2 = t1
    t1 = t3
   Next i
  Next j
  For j = 1 To hh - 1 Step 2
   t1 = nTemp(0, j) 'right
   t2 = nTemp(2 And (ww - 1), j) 'right-right
   For i = ww - 2 To 0 Step -2
    t3 = nTemp(i, j) 'this
    t4 = t1 + t3
    nTemp(i + 1, j) = (t4 * 8& + t4 - t2 - nTemp((i - 2) And (ww - 1), j)) \ 16&
    'move it
    t2 = t1
    t1 = t3
   Next i
  Next j
 Else
  For j = hh - 1 To 0 Step -1
   t1 = nTemp(0, j) 'right
   jj = (j + 1) And (hh - 1)
   t2 = nTemp(0, jj) 'right-down
   For i = ww - 1 To 0 Step -1
    t3 = nTemp(i, j) 'this
    t4 = nTemp(i, jj) 'down
    t1 = t1 + t3
    nTemp(i + i, j + j) = t3
    nTemp(i + i + 1, j + j) = t1 \ 2
    nTemp(i + i, j + j + 1) = (t3 + t4) \ 2
    nTemp(i + i + 1, j + j + 1) = (t1 + t2 + t4) \ 4
    'move it
    t1 = t3
    t2 = t4
   Next i
  Next j
  ww = ww + ww
  hh = hh + hh
 End If
 lp = lp + 1
Loop
'normalize
t1 = &H80000000 'max
t2 = &H7FFFFFFF 'min
For j = 0 To h - 1
 For i = 0 To w - 1
  t3 = nTemp(i, j)
  If t1 < t3 Then t1 = t3
  If t2 > t3 Then t2 = t3
 Next i
Next j
t1 = t1 - t2 + 1
'set value
lp = 0
For j = 0 To h - 1
 For i = 0 To w - 1
  t3 = ((nTemp(i, j) - t2) * 256&) \ t1
  bDib(lp) = TheClrTable(t3)
  lp = lp + 1
 Next i
Next j
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcPerlin(ByVal lpbm As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Const TheConst As Double = π / 512
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim clr1 As RGBQUAD, clr2 As RGBQUAD
Dim TheSeed As Long
Dim i As Long, j As Long, lp As Long
Dim TheClrTable(255) As RGBQUAD
Dim CurveTable(255) As Byte 'max=256 when index=256 XXX
Dim FuncTable(-255 To 256) As Long
Dim fAmp As Single, fGamma As Single
Dim x As Long, y As Long
Dim nXDelta As Long, nYDelta As Long '1 -> 16384
Dim nAmount As Long, nFadeOff As Long
'noise table
Dim nTemp() As Long     'max=256
'noise cache
Dim ww As Long
Dim ii As Long, jj As Long
Dim n00 As Long, n01 As Long, n10 As Long, n11 As Long
Dim nn0 As Long, nn1 As Long 'interpolated
Dim nnn As Long 'result
'init array
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = w * h
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
ReDim nTemp(w - 1, h - 1)
'init color table
CopyMemory fAmp, bProps(9), 4&
CopyMemory fGamma, bProps(13), 4&
CopyMemory clr1, bProps(17), 4&
CopyMemory clr2, bProps(21), 4&
fAmp = fAmp * 256#
For i = 0 To 255 '??
 j = fAmp * (i / 256#) ^ fGamma '??
 If j < 0 Then j = 0 Else If j > 255 Then j = 255
 pMixColor TheClrTable(i), clr1, clr2, j
Next i
'init curve
For i = 0 To 255
 'return (a * a * (3.0 - 2.0 * a));
 CurveTable(i) = (i * i * (768& - i - i)) \ &H10000
'    double a3 = a * a * a;
'    double a4 = a3 * a;
'    double a5 = a4 * a;
'    return (6.0 * a5) - (15.0 * a4) + (10.0 * a3);
Next i
'init function
Select Case bProps(2) And 3&
Case 0 'normal
 For i = -255 To 256
  FuncTable(i) = i
 Next i
Case 1 'abs
 FuncTable(0) = -256
 For i = 1 To 255
  FuncTable(i) = i + i - 256
  FuncTable(-i) = FuncTable(i)
 Next i
 FuncTable(256) = 256
Case 2 'sin
 For i = 1 To 255
  FuncTable(i) = 256# * Sin(i * TheConst)
  FuncTable(-i) = -FuncTable(i)
 Next i
 FuncTable(256) = 256
Case 3 'abs(sin)
 FuncTable(0) = -256
 For i = 1 To 255
  FuncTable(i) = CLng(512# * Sin(i * TheConst)) - 256
  FuncTable(-i) = FuncTable(i)
 Next i
 FuncTable(256) = 256
End Select
'calc noise
pGetSeed bProps, 7, TheSeed
CopyMemory fAmp, bProps(3), 4& 'fadeoff
nFadeOff = fAmp * 256#
ww = cUnk.fShl(1, bProps(1) And &HF&)
nXDelta = (ww * 16384&) \ w
nYDelta = (ww * 16384&) \ h
nAmount = 256&
For lp = 1 To ((bProps(1) And &HF0&) \ &H10&)
 If nXDelta > 16384& And nYDelta > 16384& Then Exit For
 x = 0
 y = 0
 jj = 0
 'init cache
 For j = 0 To h - 1
  If y >= 16384& Then
   y = y And 16383&
   jj = jj + 1
  End If
  i = 1 And (ww - 1)
  ii = (jj + 1) And (ww - 1)
  n00 = cUnk.fRnd3(0, jj, lp, TheSeed) And &H1FF&
  n01 = cUnk.fRnd3(i, jj, lp, TheSeed) And &H1FF&
  'y-interpolate
  If y >= 64& Then
   n10 = cUnk.fRnd3(0, ii, lp, TheSeed) And &H1FF&
   n11 = cUnk.fRnd3(i, ii, lp, TheSeed) And &H1FF&
   i = CurveTable(y \ 64&)
   nn0 = n00 + ((n10 - n00) * i) \ 256&
   nn1 = n01 + ((n11 - n01) * i) \ 256&
  Else
   nn0 = n00
   nn1 = n01
  End If
  x = 0
  ii = 0
  For i = 0 To w - 1
   If x >= 16384& Then
    x = x And 16383&
    ii = ii + 1
    nn0 = nn1
    n01 = cUnk.fRnd3((ii + 1) And (ww - 1), jj, lp, TheSeed) And &H1FF&
    If y >= 64& Then
     n11 = cUnk.fRnd3((ii + 1) And (ww - 1), _
     (jj + 1) And (ww - 1), lp, TheSeed) And &H1FF&
     'y-interpolate
     nn1 = n01 + ((n11 - n01) * CurveTable(y \ 64&)) \ 256&
    Else
     nn1 = n01
    End If
   End If
   If x >= 64& Then
    'x-interpolate
    nnn = nn0 + ((nn1 - nn0) * CurveTable(x \ 64&)) \ 256&
   Else
    nnn = nn0
   End If
   'calc function
   nnn = FuncTable(nnn - 255&)
   'add it!!
   nTemp(i, j) = nTemp(i, j) + (nnn * nAmount) \ 256&
   'next
   x = x + nXDelta
  Next i
  y = y + nYDelta
 Next j
 If nAmount > -1048576 And nAmount < 1048576 Then
  nAmount = (nAmount * nFadeOff) \ 256&
 ElseIf nFadeOff < 0 Then
  nAmount = -nAmount
 End If
 nXDelta = nXDelta + nXDelta
 nYDelta = nYDelta + nYDelta
 ww = ww + ww
Next lp
'draw picture
lp = 0
For j = 0 To h - 1
 For i = 0 To w - 1
  nnn = (nTemp(i, j) + 512&) \ 4&
  If nnn < 0 Then nnn = 0 Else If nnn > 255 Then nnn = 255
  bDib(lp) = TheClrTable(nnn)
  lp = lp + 1
 Next i
Next j
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pMixColor(clr As RGBQUAD, clr1 As RGBQUAD, clr2 As RGBQUAD, ByVal i As Long)
'TODO: x/255 = (x+(x<<8)+&H8000&)>>16
With clr1
 clr.rgbBlue = .rgbBlue + (i * (-.rgbBlue + clr2.rgbBlue)) \ 255&
 clr.rgbGreen = .rgbGreen + (i * (-.rgbGreen + clr2.rgbGreen)) \ 255&
 clr.rgbRed = .rgbRed + (i * (-.rgbRed + clr2.rgbRed)) \ 255&
 clr.rgbReserved = .rgbReserved + (i * (-.rgbReserved + clr2.rgbReserved)) \ 255&
End With
End Sub

'note:i=0-1024
Private Sub pMixColor1024(clrRet As RGBQUAD, clr As RGBQUAD, ByVal i As Long)
If i > 1023 Then
 clrRet = clr
ElseIf i > 0 Then
 With clrRet
  .rgbBlue = .rgbBlue + (i * (-.rgbBlue + clr.rgbBlue)) \ 1024&
  .rgbGreen = .rgbGreen + (i * (-.rgbGreen + clr.rgbGreen)) \ 1024&
  .rgbRed = .rgbRed + (i * (-.rgbRed + clr.rgbRed)) \ 1024&
  .rgbReserved = .rgbReserved + (i * (-.rgbReserved + clr.rgbReserved)) \ 1024&
 End With
End If
End Sub

Private Sub pBlendColor(clrRet As RGBQUAD, clr As RGBQUAD)
'TODO:use MMX instructions to optimize it
Dim i As Long, j As Long
i = 255& - clr.rgbReserved
With clrRet
 j = clr.rgbBlue + (i * .rgbBlue) \ 255&
 If j > 255& Then j = 255&
 .rgbBlue = j
 j = clr.rgbGreen + (i * .rgbGreen) \ 255&
 If j > 255& Then j = 255&
 .rgbGreen = j
 j = clr.rgbRed + (i * .rgbRed) \ 255&
 If j > 255& Then j = 255&
 .rgbRed = j
 j = clr.rgbReserved + (i * .rgbReserved) \ 255&
 If j > 255& Then j = 255&
 .rgbReserved = j
End With
End Sub

'stupid
Private Sub pMixColor2(clrRet As RGBQUAD, clr As RGBQUAD)
Dim i As Long
i = clr.rgbReserved
With clrRet
 .rgbBlue = .rgbBlue + (i * (-.rgbBlue + clr.rgbBlue)) \ 255&
 .rgbGreen = .rgbGreen + (i * (-.rgbGreen + clr.rgbGreen)) \ 255&
 .rgbRed = .rgbRed + (i * (-.rgbRed + clr.rgbRed)) \ 255&
End With
End Sub

'even more stupid
Private Sub pMixColor65536(clrRet As RGBQUAD, clr As RGBQUAD, ByVal nMode As Long, ByVal nAlpha As Long)
Dim k As Long
    If nMode > 0 Then nAlpha = (nAlpha * clr.rgbReserved) \ 255&
    With clrRet
     If nMode = 1 Then 'blend
      nAlpha = 65536 - nAlpha
      k = clr.rgbBlue + ((-clr.rgbBlue + .rgbBlue) * nAlpha) \ 65536
      If k > 255 Then k = 255
      .rgbBlue = k
      k = clr.rgbGreen + ((-clr.rgbGreen + .rgbGreen) * nAlpha) \ 65536
      If k > 255 Then k = 255
      .rgbGreen = k
      k = clr.rgbRed + ((-clr.rgbRed + .rgbRed) * nAlpha) \ 65536
      If k > 255 Then k = 255
      .rgbRed = k
      k = clr.rgbReserved + ((-clr.rgbReserved + .rgbReserved) * nAlpha) \ 65536
      If k > 255 Then k = 255
      .rgbReserved = k
     Else
      .rgbBlue = .rgbBlue + ((-.rgbBlue + clr.rgbBlue) * nAlpha) \ 65536
      .rgbGreen = .rgbGreen + ((-.rgbGreen + clr.rgbGreen) * nAlpha) \ 65536
      .rgbRed = .rgbRed + ((-.rgbRed + clr.rgbRed) * nAlpha) \ 65536
      If nMode = 0 Then .rgbReserved = .rgbReserved + ((-.rgbReserved + clr.rgbReserved) * nAlpha) \ 65536
     End If
    End With
End Sub

Public Sub pGetSeed(bProps() As Byte, ByVal nOffset As Long, ByRef TheSeed As Long)
TheSeed = 0
CopyMemory TheSeed, bProps(nOffset), 2&
TheSeed = TheSeed + 10220&
End Sub

Private Sub pCalcCell(ByVal lpbm As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim i As Long, j As Long, k As Long, lp As Long
Dim ii1 As Long, ii2 As Long, jj1 As Long, jj2 As Long
Dim ii As Long, jj As Long, iii As Long, jjj As Long
Dim ptGrid() As Long, nGridCount As Long
Dim ptNode() As typeCellPoint, m As Long
Dim x As Long, y As Long, dX As Long, dy As Long
Dim min1 As Long, min2 As Long, min1i As Long ', min2i As Long
Dim clr1 As RGBQUAD, clr2 As RGBQUAD, clr3 As RGBQUAD
Dim clr As RGBQUAD
Dim nMode1 As Integer, nMode2 As Integer
Dim bCellColor As Boolean, bInvert As Boolean
Dim TheSeed As Long
Dim fAmp As Single, fGamma As Single
Dim nAspect As Long '1024+
Dim nRFactor As Long
Dim TheTable(255) As Byte
'init array
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = w * h
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
'load data
nMode1 = bProps(28) 'color mode
nMode2 = (nMode1 And 12) \ 4 'cell mode
bCellColor = nMode1 And 64
bInvert = nMode1 And 128
nMode1 = nMode1 And 3
'get aspect
CopyMemory fAmp, bProps(24), 4&
If Abs(fAmp) > 0.0005 Then
 If fAmp > 0 Then
  nAspect = 1024 * (fAmp + 1)
 Else
  nAspect = 1024 * (fAmp - 1)
 End If
End If
'init table
CopyMemory fAmp, bProps(16), 4&
CopyMemory fGamma, bProps(20), 4&
fAmp = fAmp * 256#
For i = 1 To 256 '??
 j = 256& - CLng(fAmp * (i / 256#) ^ fGamma) '??
 If j < 0 Then j = 0 Else If j > 255 Then j = 255
 TheTable(i - 1) = j
Next i
'init cell
pGetSeed bProps, 14, TheSeed
If nMode2 = 0 Then 'random
 nMode2 = bProps(9) 'max
 CopyMemory fAmp, bProps(10), 4& 'min distance
 k = fAmp * 4096&
 k = k * k
 ReDim ptNode(1 To nMode2)
 For i = 1 To nMode2
  lp = 0
  'find a position
  Do
   If lp >= 16 Then Exit For Else lp = lp + 1
   With ptNode(i)
    .x = cUnk.fRnd(TheSeed) And 4095&
    .y = cUnk.fRnd(TheSeed) And 4095&
    For j = 1 To i - 1
     dX = .x - ptNode(j).x
     dy = .y - ptNode(j).y
     'wrap it
     If dX > 2048& Then dX = 4096& - dX Else If dX < -2048& Then dX = dX + 4096&
     If dy > 2048& Then dy = 4096& - dy Else If dy < -2048& Then dy = dy + 4096&
     'check distance
     If dX * dX + dy * dy < k Then Exit For
    Next j
   End With
  Loop While j < i
  'found it!
  With ptNode(i)
   .nRnd = cUnk.fRnd(TheSeed) And 255&
   If (cUnk.fRnd(TheSeed) And 255&) < bProps(33) Then .bNoColor = True
  End With
 Next i
 m = i - 1
 nMode2 = 0
 'get count
 If m > 128 Then
  nGridCount = 20
 ElseIf m > 64 Then
  nGridCount = 15
 ElseIf m > 32 Then
  nGridCount = 10
 ElseIf m > 16 Then
  nGridCount = 6
 End If
 nRFactor = Sqr(m) * 1024#
Else
 m = bProps(35)
 If m = 0 Then m = 1 'ERR :-/
 nGridCount = m
 nRFactor = m * 1024&
 k = m * m
 ReDim ptNode(1 To k)
 lp = 1
 For j = 0 To m - 1
  For i = 0 To m - 1
   With ptNode(lp)
    .x = (i * 4096& + 2048& + ((cUnk.fRnd(TheSeed) - &H4000&) * (255& - bProps(34))) \ 2048&) \ m
    .y = (j * 4096& + 2048& + ((cUnk.fRnd(TheSeed) - &H4000&) * (255& - bProps(34))) \ 2048&) \ m
    .nRnd = cUnk.fRnd(TheSeed) And 255&
    If (cUnk.fRnd(TheSeed) And 255&) < bProps(33) Then .bNoColor = True
   End With
   lp = lp + 1
  Next i
 Next j
 m = k
End If
If nGridCount > 5 Then
 'init linked-list
 ReDim ptGrid(nGridCount - 1, nGridCount - 1)
 For k = 1 To m
  With ptNode(k)
   i = (.x * nGridCount) \ 4096&
   j = (.y * nGridCount) \ 4096&
   lp = ptGrid(i, j)
   ptGrid(i, j) = k
   .idxNext = lp
  End With
 Next k
Else
 nGridCount = 1
 ReDim ptGrid(0, 0)
 ptGrid(0, 0) = 1
 For k = 1 To m - 1
  ptNode(k).idxNext = k + 1
 Next k
End If
'calc data
CopyMemory clr1, bProps(1), 4&
CopyMemory clr2, bProps(5), 4&
CopyMemory clr3, bProps(29), 4&
lp = 0
min1i = 1
For j = 0 To h - 1
 y = (j * 4096&) \ h
 If nGridCount > 5 Then
  jj1 = (y * nGridCount) \ 4096& - 2&
  If jj1 < 0 Then jj1 = jj1 + nGridCount
  jj2 = jj1 + 4&
 End If
 For i = 0 To w - 1
  x = (i * 4096&) \ w
  'calc min distance
  If nGridCount > 5 Then
   ii1 = (x * nGridCount) \ 4096& - 2&
   If ii1 < 0 Then ii1 = ii1 + nGridCount
   ii2 = ii1 + 4&
  End If
  min1 = &HFFFFFF
  min2 = &HFFFFFF
  For jj = jj1 To jj2
   If jj >= nGridCount Then jjj = jj - nGridCount Else jjj = jj
   For ii = ii1 To ii2
    If ii >= nGridCount Then iii = ii - nGridCount Else iii = ii
    k = ptGrid(iii, jjj)
    Do While k > 0
     dX = x - ptNode(k).x
     dy = y - ptNode(k).y
     'wrap it
     If dX > 2048& Then dX = 4096& - dX Else If dX < -2048& Then dX = dX + 4096&
     If dy > 2048& Then dy = 4096& - dy Else If dy < -2048& Then dy = dy + 4096&
     'calc distance
     If nAspect = 0 Then
      dX = dX * dX + dy * dy
     ElseIf nAspect > 0 Then
      dX = ((dX * dX * 256&) \ (nAspect)) * 4& + dy * dy
     Else
      dX = dX * dX + ((dy * dy * 256&) \ (-nAspect)) * 4&
     End If
     If dX < min1 Then
      min2 = min1
      min1 = dX
      min1i = k
     ElseIf dX < min2 Then
      min2 = dX
     End If
     'next
     k = ptNode(k).idxNext
    Loop
   Next ii
  Next jj
  'calc color
  If bCellColor And ptNode(min1i).bNoColor Then
   bDib(lp) = clr3
  Else
   'get color index
   'a=min b=not min r=predefined radius??
   '                   border  center
   'outer (b-a)/(b+a)  0       1
   'inner a/r          x       0
   'cross b/r
   '///////////////////Fast SQRT test!!!
   Dim nLastMin1 As Long
   Dim nLastMin2 As Long
   '       t      a
   ' t' = --- + -----
   '       2     2*t
   '///////////////////
   Select Case nMode1
   Case 0
    If nLastMin1 <= 0 Then
     nLastMin1 = Sqr(min1)
    Else
     nLastMin1 = (nLastMin1 + min1 \ nLastMin1) \ 2
    End If
    If nLastMin2 <= 0 Then
     nLastMin2 = Sqr(min2)
    Else
     nLastMin2 = (nLastMin2 + min2 \ nLastMin2) \ 2
    End If
    k = ((nLastMin2 - nLastMin1) * 255&) \ (nLastMin2 + nLastMin1)
   Case 1
    If nLastMin1 <= 0 Then
     nLastMin1 = Sqr(min1)
    Else
     nLastMin1 = (nLastMin1 + min1 \ nLastMin1) \ 2
    End If
    k = (nLastMin1 * nRFactor) \ 16384&
   Case 2
    If nLastMin2 <= 0 Then
     nLastMin2 = Sqr(min2)
    Else
     nLastMin2 = (nLastMin2 + min2 \ nLastMin2) \ 2
    End If
    k = (nLastMin2 * nRFactor) \ 16384&
   End Select
   If nMode2 = 2 Then 'chessboard
    k = k \ 2
    If ((x * nGridCount) Xor (y * nGridCount)) And 4096& Then
     k = 255 - k
    End If
   End If
   If k < 0 Then k = 0 Else If k > 255 Then k = 255
   k = TheTable(k)
   If bCellColor Then
    pMixColor clr, clr1, clr2, ptNode(min1i).nRnd
    If bInvert Then k = 255 - k
    pMixColor bDib(lp), clr3, clr, k
   Else
    If bInvert Then k = 255 - k
    pMixColor bDib(lp), clr1, clr2, k
   End If
  End If
  ''''''''''''''''''''''
  lp = lp + 1
 Next i
Next j
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcBrick(ByVal lpbm As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim i As Long, j As Long, lp As Long
Dim ii As Long, jj As Long
Dim iii As Long, jjj As Long
Dim clr1 As RGBQUAD, clr2 As RGBQUAD, clr3 As RGBQUAD
Dim TheSeed As Long
Dim TheClrTable(1023) As RGBQUAD
Dim nXCount As Long, nYCount As Long
Dim x As Long, xx As Long, y As Long
Dim nXDelta As Long, nYDelta As Long '1 -> 16384 ??
Dim nXDelta2 As Long
Dim nXSize As Long, nYSize As Long 'size of joint
Dim nMultiply As Long
Dim f1 As Single
Dim bFirstSingle As Boolean, bLastSingle As Boolean
Dim bNoAdjacentSin As Boolean, bSolidColor As Boolean
Dim TheBrick() As Byte '0=normal 1=single-width stone
Dim TheBrickClr() As RGBQUAD
'init array
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = w * h
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
'init color table
CopyMemory clr1, bProps(1), 4&
CopyMemory clr2, bProps(5), 4&
CopyMemory f1, bProps(31), 4&
For i = 0 To 1023
 j = 255# * ((i / 1023#) ^ f1) '0^0=1 !? stupid...
 pMixColor TheClrTable(i), clr1, clr2, j
Next i
CopyMemory clr1, bProps(9), 4& 'color joint
'load params
nXCount = bProps(21)
If nXCount = 0 Then nXCount = 1
nYCount = bProps(22)
If nYCount = 0 Then nYCount = 1
nMultiply = cUnk.fShl(1, (bProps(26) And 28&) \ 4&)
bNoAdjacentSin = bProps(26) And 1&
bSolidColor = bProps(26) And 2&
'init bricks
pGetSeed bProps, 23, TheSeed
ReDim TheBrick(nXCount - 1, nYCount - 1)
For j = 0 To nYCount - 1
 i = 0
 lp = 0
 If nXCount = 1 Then
  lp = 1
 ElseIf nXCount > 2 Or Not bNoAdjacentSin Then
  If (cUnk.fRnd2(i, j + 25981&, TheSeed) And &HFF&) < bProps(25) Then lp = 1
 End If
 bFirstSingle = lp
 Do
  TheBrick(i, j) = 1
  bLastSingle = lp
  i = i + 2 - lp
  nXSize = nXCount - i
  If nXSize <= 0 Then Exit Do
  If nXSize = 1 Or (nXSize = 3 And bFirstSingle And bNoAdjacentSin) Then
   lp = 1
  ElseIf (nXSize = 2 Or bLastSingle Or (nXSize = 4 And bFirstSingle)) And bNoAdjacentSin Then
   lp = 0
  ElseIf (cUnk.fRnd2(i, j + 25981&, TheSeed) And &HFF&) < bProps(25) Then
   lp = 1
  Else
   lp = 0
  End If
 Loop
Next j
'get width
nYDelta = nMultiply * 16384&
nXDelta = (nXCount * nYDelta) \ w
nYDelta = (nYCount * nYDelta) \ h
CopyMemory f1, bProps(27), 4&
nXDelta2 = 16384# * f1
CopyMemory f1, bProps(13), 4&
nXSize = 8192# * f1
CopyMemory f1, bProps(17), 4&
nYSize = 8192# * f1
'draw bricks
nMultiply = nMultiply * nXCount
ReDim TheBrickClr(nMultiply - 1)
lp = 0
xx = 16384& ':-3
y = 16384& ':-3
jj = -1 ':-3
jjj = -1 ':-3
For j = 0 To h - 1
 If y >= 16384& Then
  'next row
  Do
   y = y - 16384&
   jj = jj + 1 'y-index
   jjj = jjj + 1 'brick index
   If jjj >= nYCount Then jjj = 0
   xx = xx + nXDelta2
  Loop While y >= 16384&
  'calc next color
  iii = 0
  For ii = 0 To nMultiply - 1
   If TheBrick(iii, jjj) Then
    TheBrickClr(ii) = TheClrTable(cUnk.fRnd2(ii, jj, TheSeed) And 1023&)
   Else
    TheBrickClr(ii) = TheBrickClr(ii - 1)
   End If
   iii = iii + 1
   If iii >= nXCount Then iii = 0
  Next ii
 End If
 x = xx
 If cUnk.fRnd2(30087&, jj, TheSeed) And 1& Then x = x + 16384& 'test
 ii = -1 ':-3
 iii = -1 ':-3
 For i = 0 To w - 1
  If x >= 16384& Then
   Do
    x = x - 16384&
    ii = ii + 1
    If ii >= nMultiply Then ii = 0
    iii = iii + 1
    If iii >= nXCount Then iii = 0
   Loop While x >= 16384&
   'get brick color
   clr3 = TheBrickClr(ii)
   'calc joint color
   If bSolidColor Then
    If y < nYSize Or y > 16384& - nYSize Then
     clr2 = clr1
    Else
     clr2 = clr3
    End If
   ElseIf y < nYSize Then 'top
    pMixColor clr2, clr1, clr3, (y * 255&) \ nYSize
   ElseIf y > 16384& - nYSize Then 'bottom
    pMixColor clr2, clr1, clr3, ((16384& - y) * 255&) \ nYSize
   Else
    clr2 = clr3
   End If
  End If
  'calc color
  bFirstSingle = False
  If x < nXSize Then 'left
   If TheBrick(iii, jjj) Then
    If bSolidColor Then
     bDib(lp) = clr1
     bFirstSingle = True
    Else
     If y < nYSize Then
      bFirstSingle = x * nYSize < y * nXSize
     ElseIf y > 16384& - nYSize Then
      bFirstSingle = x * nYSize < (16384& - y) * nXSize
     Else
      bFirstSingle = True
     End If
     If bFirstSingle Then
      pMixColor bDib(lp), clr1, clr3, (x * 255&) \ nXSize
     End If
    End If
   End If
  ElseIf x > 16384& - nXSize Then 'right
   If iii + 1 < nXCount Then
    bFirstSingle = TheBrick(iii + 1, jjj)
   Else
    bFirstSingle = TheBrick(0, jjj)
   End If
   If bFirstSingle Then
    bFirstSingle = False
    If bSolidColor Then
     bDib(lp) = clr1
     bFirstSingle = True
    Else
     If y < nYSize Then
      bFirstSingle = (16384& - x) * nYSize < y * nXSize
     ElseIf y > 16384& - nYSize Then
      bFirstSingle = (16384& - x) * nYSize < (16384& - y) * nXSize
     Else
      bFirstSingle = True
     End If
     If bFirstSingle Then
      pMixColor bDib(lp), clr1, clr3, ((16384& - x) * 255&) \ nXSize
     End If
    End If
   End If
  End If
  If Not bFirstSingle Then bDib(lp) = clr2
  'next pixel
  lp = lp + 1
  x = x + nXDelta
 Next i
 'add delta
 y = y + nYDelta
Next j
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcGradient(ByVal lpbm As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Const TheConst As Double = π / 1024
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim i As Long, j As Long, lp As Long
Dim k As Long, n As Long, idx As Long
Dim TheTable(-511 To 511) As Byte
Dim TheClrTable(255) As RGBQUAD
Dim clr1 As RGBQUAD, clr2 As RGBQUAD
Dim fPosition As Single
Dim fAngle As Single
Dim fWidth As Single
Dim nXDelta As Long, nYDelta As Long 'fake-float operation!!! {x}=7 bit
'init array
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = w * h
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
'get color
CopyMemory fPosition, bProps(9), 4&
CopyMemory fAngle, bProps(13), 4&
CopyMemory fWidth, bProps(17), 4&
If fWidth > 0 Then
 CopyMemory clr1, bProps(1), 4&
 CopyMemory clr2, bProps(5), 4&
Else
 fWidth = -fWidth
 CopyMemory clr1, bProps(5), 4&
 CopyMemory clr2, bProps(1), 4&
End If
'init table
For i = 0 To 255
 pMixColor TheClrTable(i), clr1, clr2, i
Next i
n = bProps(21) And 3&
Select Case n
Case 1 'gaussian
 clr2 = clr1
 TheTable(0) = 255
 For i = 1 To 511
  TheTable(i) = Int(256 * Exp((-i * i) * 0.00002))
  TheTable(-i) = TheTable(i)
 Next i
Case 2 'sin
 clr2 = clr1
 TheTable(0) = 255
 For i = 1 To 511
  TheTable(i) = Int(256 * Cos(i * TheConst))
  TheTable(-i) = TheTable(i)
 Next i
Case Else 'linear
 For i = -511 To 511
  TheTable(i) = (512& + i) \ 4&
 Next i
End Select
'start calc
If fWidth < 0.0001 Then
 For j = 0 To h - 1
  For i = 0 To w - 1
   bDib(lp) = clr1
   lp = lp + 1
  Next i
 Next j
Else
 'init angle
 fAngle = (fAngle - Int(fAngle)) * 二π
 nXDelta = Cos(fAngle) / fWidth * (131072 \ w)
 nYDelta = Sin(fAngle) / fWidth * (131072 \ h)
 k = -((1 + fPosition) * Cos(fAngle) + Sin(fAngle)) / fWidth * 65536#
 For j = 0 To h - 1
  n = k
  For i = 0 To w - 1
   idx = n \ 128&
   If idx <= -512 Then
    bDib(lp) = clr1
   ElseIf idx >= 512 Then
    bDib(lp) = clr2
   Else
    idx = TheTable(idx)
    bDib(lp) = TheClrTable(idx)
   End If
   lp = lp + 1
   n = n + nXDelta
  Next i
  k = k + nYDelta
 Next j
End If
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcGradient2(ByVal lpbm As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim i As Long, j As Long, lp As Long
Dim k As Long, TheTable() As Byte
Dim clrs(3) As RGBQUAD
Dim clr1 As RGBQUAD, clr2 As RGBQUAD
'init array
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = w * h
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
'init table
ReDim TheTable(w - 1)
For i = 0 To w - 1
 TheTable(i) = (i * 256&) \ w
Next i
'get color
CopyMemory clrs(0), bProps(1), 16&
For j = 0 To h - 1
 k = (j * 256&) \ h
 For i = 0 To w - 1
  pMixColor clr1, clrs(0), clrs(2), k
  pMixColor clr2, clrs(1), clrs(3), k
  pMixColor bDib(lp), clr1, clr2, TheTable(i)
  lp = lp + 1
 Next i
Next j
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcNoise(ByVal lpbm As Long, ByVal m As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim i As Long
Dim TheSeed As Long
Dim clrs(1) As RGBQUAD
Dim TheClrTable(255) As RGBQUAD
'init array
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = m
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
'get color
CopyMemory clrs(0), bProps(1), 8&
'init table
For i = 0 To 255
 pMixColor TheClrTable(i), clrs(0), clrs(1), i
Next i
'randomize
pGetSeed bProps, 9, TheSeed
For i = 0 To m - 1
 bDib(i) = TheClrTable(cUnk.fRnd(TheSeed) And &HFF&)
Next i
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcAddBitmap(ByVal lpbm As Long, ByVal m As Long, ByVal nCount As Long, bmIn() As typeAlphaDibSectionDescriptor, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim bDib2() As RGBQUAD, clr As RGBQUAD
Dim tSA2 As SAFEARRAY2D
Dim i As Long, nIndex As Long
Dim f(3) As Single
Dim nScaleBlue As Long
Dim nScaleGreen As Long
Dim nScaleRed As Long
Dim nScaleReserved As Long 'alpha
Dim nNewBlue As Long
Dim nNewGreen As Long
Dim nNewRed As Long
Dim nMode As Long, k As Long
'init array
CopyMemory ByVal lpbm, ByVal bmIn(0).lpbm, 4& * m
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = m
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
'get scale
CopyMemory f(0), bProps(1), 16&
nScaleReserved = f(0) * 256#
nScaleRed = f(1) * 256#
nScaleGreen = f(2) * 256#
nScaleBlue = f(3) * 256#
'get mode
nMode = bProps(0)
Select Case nMode
Case 2, 3 'sub clamp,sub wrap
 nScaleReserved = -nScaleReserved
 nScaleRed = -nScaleRed
 nScaleGreen = -nScaleGreen
 nScaleBlue = -nScaleBlue
 nMode = nMode - 2
Case 4, 10, 11   'mul,min,max
 If nScaleReserved < 0 Then nScaleReserved = 0
 If nScaleRed < 0 Then nScaleRed = 0
 If nScaleGreen < 0 Then nScaleGreen = 0
 If nScaleBlue < 0 Then nScaleBlue = 0
End Select
For nIndex = 1 To nCount - 1
 'get array
 With tSA2
  .cbElements = 4
  .cDims = 1
  .Bounds(0).cElements = m
  .pvData = bmIn(nIndex).lpbm
 End With
 CopyMemory ByVal VarPtrArray(bDib2()), VarPtr(tSA2), 4&
 'mode?
 Select Case nMode
 Case 0 'add clamp,sub clamp
  For i = 0 To m - 1
   clr = bDib2(i)
   With bDib(i)
    k = .rgbBlue + (nScaleBlue * clr.rgbBlue) \ 256&
    If k < 0 Then k = 0 Else If k > 255 Then k = 255
    .rgbBlue = k
    k = .rgbGreen + (nScaleGreen * clr.rgbGreen) \ 256&
    If k < 0 Then k = 0 Else If k > 255 Then k = 255
    .rgbGreen = k
    k = .rgbRed + (nScaleRed * clr.rgbRed) \ 256&
    If k < 0 Then k = 0 Else If k > 255 Then k = 255
    .rgbRed = k
    k = .rgbReserved + (nScaleReserved * clr.rgbReserved) \ 256&
    If k < 0 Then k = 0 Else If k > 255 Then k = 255
    .rgbReserved = k
   End With
  Next i
 Case 1 'add wrap,sub wrap
  'IDE??
  If m_bIsInIDE Then
   For i = 0 To m - 1
    clr = bDib2(i)
    With bDib(i)
     .rgbBlue = (.rgbBlue + (nScaleBlue * clr.rgbBlue) \ 256&) And 255
     .rgbGreen = (.rgbGreen + (nScaleGreen * clr.rgbGreen) \ 256&) And 255
     .rgbRed = (.rgbRed + (nScaleRed * clr.rgbRed) \ 256&) And 255
     .rgbReserved = (.rgbReserved + (nScaleReserved * clr.rgbReserved) \ 256&) And 255
    End With
   Next i
  Else 'remove overflow check :-3
   For i = 0 To m - 1
    clr = bDib2(i)
    With bDib(i)
     .rgbBlue = (.rgbBlue + (nScaleBlue * clr.rgbBlue) \ 256&)
     .rgbGreen = (.rgbGreen + (nScaleGreen * clr.rgbGreen) \ 256&)
     .rgbRed = (.rgbRed + (nScaleRed * clr.rgbRed) \ 256&)
     .rgbReserved = (.rgbReserved + (nScaleReserved * clr.rgbReserved) \ 256&)
    End With
   Next i
  End If
 Case 4 'mul
  For i = 0 To m - 1
   clr = bDib2(i)
   With bDib(i)
    k = (.rgbBlue * nScaleBlue * clr.rgbBlue) \ &HFF00&
    If k > 255 Then k = 255
    .rgbBlue = k
    k = (.rgbGreen * nScaleGreen * clr.rgbGreen) \ &HFF00&
    If k > 255 Then k = 255
    .rgbGreen = k
    k = (.rgbRed * nScaleRed * clr.rgbRed) \ &HFF00&
    If k > 255 Then k = 255
    .rgbRed = k
    k = (.rgbReserved * nScaleReserved * clr.rgbReserved) \ &HFF00&
    If k > 255 Then k = 255
    .rgbReserved = k
   End With
  Next i
 Case 5 'diff
  For i = 0 To m - 1
   clr = bDib2(i)
   With bDib(i)
    k = .rgbBlue - (nScaleBlue * clr.rgbBlue) \ 256&
    If k < 0 Then k = -k
    If k > 255 Then k = 255
    .rgbBlue = k
    k = .rgbGreen - (nScaleGreen * clr.rgbGreen) \ 256&
    If k < 0 Then k = -k
    If k > 255 Then k = 255
    .rgbGreen = k
    k = .rgbRed - (nScaleRed * clr.rgbRed) \ 256&
    If k < 0 Then k = -k
    If k > 255 Then k = 255
    .rgbRed = k
    k = .rgbReserved - (nScaleReserved * clr.rgbReserved) \ 256&
    If k < 0 Then k = -k
    If k > 255 Then k = 255
    .rgbReserved = k
   End With
  Next i
 Case 6 'alpha
  For i = 0 To m - 1
   With bDib2(i)
    nNewBlue = (nScaleBlue * .rgbBlue) \ 256&
    nNewGreen = (nScaleGreen * .rgbGreen) \ 256&
    nNewRed = (nScaleRed * .rgbRed) \ 256&
   End With
   nNewRed = (nNewBlue * 146& + nNewGreen * 1454& + nNewRed * 456& _
   + 512&) \ 1024&
   With bDib(i)
    If nNewRed < 0 Then nNewRed = 0 Else If nNewRed > 512 Then nNewRed = 512
    .rgbBlue = (.rgbBlue * nNewRed) \ 512&
    .rgbGreen = (.rgbGreen * nNewRed) \ 512&
    .rgbRed = (.rgbRed * nNewRed) \ 512&
    .rgbReserved = (.rgbReserved * nNewRed) \ 512&
   End With
  Next i
 Case 7 'brightness
  For i = 0 To m - 1
   With bDib2(i)
    nNewBlue = (nScaleBlue * .rgbBlue) \ 256&
    nNewGreen = (nScaleGreen * .rgbGreen) \ 256&
    nNewRed = (nScaleRed * .rgbRed) \ 256&
   End With
   '0-524280
   nNewRed = (nNewBlue * 146& + nNewGreen * 1454& + nNewRed * 456& _
   + 512&) \ 1024&
   With bDib(i)
    If nNewRed > 256 Then 'to white
     If nNewRed > 512 Then nNewRed = 512
     nNewRed = nNewRed - 256
     .rgbBlue = .rgbBlue + ((255 - .rgbBlue) * nNewRed) \ 256&
     .rgbGreen = .rgbGreen + ((255 - .rgbGreen) * nNewRed) \ 256&
     .rgbRed = .rgbRed + ((255 - .rgbRed) * nNewRed) \ 256&
    ElseIf nNewRed < 256 Then 'to black
     If nNewRed < 0 Then nNewRed = 0
     .rgbBlue = (.rgbBlue * nNewRed) \ 256&
     .rgbGreen = (.rgbGreen * nNewRed) \ 256&
     .rgbRed = (.rgbRed * nNewRed) \ 256&
    End If
   End With
  Next i
 Case 8 'AlphaBlend
  For i = 0 To m - 1
   clr = bDib2(i)
   nNewRed = 256& - (clr.rgbReserved * nScaleReserved) \ 255& '255 -> 256
   With bDib(i)
    k = clr.rgbBlue + (.rgbBlue * nScaleBlue * nNewRed) \ 65536
    If k < 0 Then k = 0 Else If k > 255 Then k = 255
    .rgbBlue = k
    k = clr.rgbGreen + (.rgbGreen * nScaleGreen * nNewRed) \ 65536
    If k < 0 Then k = 0 Else If k > 255 Then k = 255
    .rgbGreen = k
    k = clr.rgbRed + (.rgbRed * nScaleRed * nNewRed) \ 65536
    If k < 0 Then k = 0 Else If k > 255 Then k = 255
    .rgbRed = k
    k = clr.rgbReserved + (.rgbReserved * nScaleReserved * nNewRed) \ 65536
    If k < 0 Then k = 0 Else If k > 255 Then k = 255
    .rgbReserved = k
   End With
  Next i
 Case 9 'add smooth
  'y=a+b-ab ??
  ' =a+(1-a)b
  For i = 0 To m - 1
   clr = bDib2(i)
   With bDib(i)
    k = .rgbBlue + ((255& - .rgbBlue) * nScaleBlue * clr.rgbBlue) \ &HFF00&
    If k < 0 Then k = 0 Else If k > 255 Then k = 255
    .rgbBlue = k
    k = .rgbGreen + ((255& - .rgbGreen) * nScaleGreen * clr.rgbGreen) \ &HFF00&
    If k < 0 Then k = 0 Else If k > 255 Then k = 255
    .rgbGreen = k
    k = .rgbRed + ((255& - .rgbRed) * nScaleRed * clr.rgbRed) \ &HFF00&
    If k < 0 Then k = 0 Else If k > 255 Then k = 255
    .rgbRed = k
    k = .rgbReserved + ((255& - .rgbReserved) * nScaleReserved * clr.rgbReserved) \ &HFF00&
    If k < 0 Then k = 0 Else If k > 255 Then k = 255
    .rgbReserved = k
   End With
  Next i
 Case 10 'min
  For i = 0 To m - 1
   With bDib2(i)
    nNewBlue = (nScaleBlue * .rgbBlue) \ 256&
    nNewGreen = (nScaleGreen * .rgbGreen) \ 256&
    nNewRed = (nScaleRed * .rgbRed) \ 256&
    k = (nScaleReserved * .rgbReserved) \ 256&
   End With
   With bDib(i)
    If .rgbBlue > nNewBlue Then .rgbBlue = nNewBlue
    If .rgbGreen > nNewGreen Then .rgbGreen = nNewGreen
    If .rgbRed > nNewRed Then .rgbRed = nNewRed
    If .rgbReserved > k Then .rgbReserved = k
   End With
  Next i
 Case 11 'max
  For i = 0 To m - 1
   With bDib2(i)
    nNewBlue = (nScaleBlue * .rgbBlue) \ 256&
    nNewGreen = (nScaleGreen * .rgbGreen) \ 256&
    nNewRed = (nScaleRed * .rgbRed) \ 256&
    k = (nScaleReserved * .rgbReserved) \ 256&
   End With
   With bDib(i)
    If .rgbBlue < nNewBlue Then
     If nNewBlue > 255 Then nNewBlue = 255
     .rgbBlue = nNewBlue
    End If
    If .rgbGreen < nNewGreen Then
     If nNewGreen > 255 Then nNewGreen = 255
     .rgbGreen = nNewGreen
    End If
    If .rgbRed < nNewRed Then
     If nNewRed > 255 Then nNewRed = 255
     .rgbRed = nNewRed
    End If
    If .rgbReserved < k Then
     If k > 255 Then k = 255
     .rgbReserved = k
    End If
   End With
  Next i
 End Select
 'destroy array
 ZeroMemory ByVal VarPtrArray(bDib2()), 4&
Next nIndex
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcAddColor(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal m As Long, bProps() As Byte)
Dim bDib() As RGBQUAD, lDib() As Long
Dim tSA As SAFEARRAY2D
Dim i As Long
Dim nClrBlue As Long
Dim nClrGreen As Long
Dim nClrRed As Long
Dim nClrReserved As Long 'alpha
Dim nMode As Long, k As Long
'init array
CopyMemory ByVal lpbm, ByVal lpbmIn, 4& * m
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = m
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
'get color
nClrBlue = bProps(1)
nClrGreen = bProps(2)
nClrRed = bProps(3)
nClrReserved = bProps(4)
'get mode
nMode = bProps(0)
Select Case nMode
Case 2, 3 'sub clamp,sub wrap
 nClrReserved = -nClrReserved
 nClrRed = -nClrRed
 nClrGreen = -nClrGreen
 nClrBlue = -nClrBlue
 nMode = nMode - 2
Case 4 'multiply
 '255 -> 1024
 nClrReserved = (nClrReserved * 1024&) \ 255&
 nClrRed = (nClrRed * 1024&) \ 255&
 nClrGreen = (nClrGreen * 1024&) \ 255&
 nClrBlue = (nClrBlue * 1024&) \ 255&
Case 10 'scale
 nClrReserved = nClrReserved * 64&
 nClrRed = nClrRed * 64&
 nClrGreen = nClrGreen * 64&
 nClrBlue = nClrBlue * 64&
 nMode = 4
End Select
'calc bitmap
Select Case nMode
Case 0 'add clamp,sub clamp
 For i = 0 To m - 1
  With bDib(i)
   k = .rgbBlue + nClrBlue
   If k < 0 Then k = 0 Else If k > 255 Then k = 255
   .rgbBlue = k
   k = .rgbGreen + nClrGreen
   If k < 0 Then k = 0 Else If k > 255 Then k = 255
   .rgbGreen = k
   k = .rgbRed + nClrRed
   If k < 0 Then k = 0 Else If k > 255 Then k = 255
   .rgbRed = k
   k = .rgbReserved + nClrReserved
   If k < 0 Then k = 0 Else If k > 255 Then k = 255
   .rgbReserved = k
  End With
 Next i
Case 1 'add wrap,sub wrap
 'IDE??
 If m_bIsInIDE Then
  For i = 0 To m - 1
   With bDib(i)
    .rgbBlue = (.rgbBlue + nClrBlue) And 255
    .rgbGreen = (.rgbGreen + nClrGreen) And 255
    .rgbRed = (.rgbRed + nClrRed) And 255
    .rgbReserved = (.rgbReserved + nClrReserved) And 255
   End With
  Next i
 Else 'remove overflow check :-3
  For i = 0 To m - 1
   With bDib(i)
    .rgbBlue = (.rgbBlue + nClrBlue)
    .rgbGreen = (.rgbGreen + nClrGreen)
    .rgbRed = (.rgbRed + nClrRed)
    .rgbReserved = (.rgbReserved + nClrReserved)
   End With
  Next i
 End If
Case 4 'mul,scale
 For i = 0 To m - 1
  With bDib(i)
   k = (.rgbBlue * nClrBlue) \ 1024&
   If k > 255 Then k = 255
   .rgbBlue = k
   k = (.rgbGreen * nClrGreen) \ 1024&
   If k > 255 Then k = 255
   .rgbGreen = k
   k = (.rgbRed * nClrRed) \ 1024&
   If k > 255 Then k = 255
   .rgbRed = k
   k = (.rgbReserved * nClrReserved) \ 1024&
   If k > 255 Then k = 255
   .rgbReserved = k
  End With
 Next i
Case 5 'diff
 For i = 0 To m - 1
  With bDib(i)
   k = .rgbBlue - nClrBlue
   If k < 0 Then k = -k
   .rgbBlue = k
   k = .rgbGreen - nClrGreen
   If k < 0 Then k = -k
   .rgbGreen = k
   k = .rgbRed - nClrRed
   If k < 0 Then k = -k
   .rgbRed = k
   k = .rgbReserved - nClrReserved
   If k < 0 Then k = -k
   .rgbReserved = k
  End With
 Next i
Case 6 'min
 For i = 0 To m - 1
  With bDib(i)
   If .rgbBlue > nClrBlue Then .rgbBlue = nClrBlue
   If .rgbGreen > nClrGreen Then .rgbGreen = nClrGreen
   If .rgbRed > nClrRed Then .rgbRed = nClrRed
   If .rgbReserved > nClrReserved Then .rgbReserved = nClrReserved
  End With
 Next i
Case 7 'max
 For i = 0 To m - 1
  With bDib(i)
   If .rgbBlue < nClrBlue Then .rgbBlue = nClrBlue
   If .rgbGreen < nClrGreen Then .rgbGreen = nClrGreen
   If .rgbRed < nClrRed Then .rgbRed = nClrRed
   If .rgbReserved < nClrReserved Then .rgbReserved = nClrReserved
  End With
 Next i
Case 8 'grayscale
 For i = 0 To m - 1
  With bDib(i)
   k = (.rgbBlue * 146& + .rgbGreen * 1454& + .rgbRed * 456& _
   + 512&)
   .rgbBlue = (nClrBlue * k) \ 524288
   .rgbGreen = (nClrGreen * k) \ 524288
   .rgbRed = (nClrRed * k) \ 524288
  End With
 Next i
Case 9 'invert
 CopyMemory ByVal VarPtrArray(lDib()), VarPtr(tSA), 4& ':-3
 For i = 0 To m - 1
  lDib(i) = lDib(i) Xor &HFFFFFF
 Next i
 ZeroMemory ByVal VarPtrArray(lDib()), 4& ':-3
Case 11 'pre-multiply
 For i = 0 To m - 1
  With bDib(i) 'TODO:nClrRed,etc ??
   nClrReserved = .rgbReserved * 257&
   .rgbBlue = (nClrReserved * .rgbBlue + 32768) \ 65536
   .rgbGreen = (nClrReserved * .rgbGreen + 32768) \ 65536
   .rgbRed = (nClrReserved * .rgbRed + 32768) \ 65536
  End With
 Next i
End Select
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcMaskBitmap(ByVal lpbm As Long, ByVal m As Long, bmIn() As typeAlphaDibSectionDescriptor, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim bDibMask() As RGBQUAD
Dim tSAMask As SAFEARRAY2D
Dim bDibSrc1() As RGBQUAD, clr1 As RGBQUAD
Dim tSASrc1 As SAFEARRAY2D
Dim bDibSrc2() As RGBQUAD, clr2 As RGBQUAD
Dim tSASrc2 As SAFEARRAY2D
Dim nNewBlue As Long
Dim nNewGreen As Long
Dim nNewRed As Long
Dim nNewReserved As Long
Dim i As Long, k As Long
'init array
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = m
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
With tSAMask
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = m
 .pvData = bmIn(0).lpbm
End With
CopyMemory ByVal VarPtrArray(bDibMask()), VarPtr(tSAMask), 4&
With tSASrc1
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = m
 .pvData = bmIn(1).lpbm
End With
CopyMemory ByVal VarPtrArray(bDibSrc1()), VarPtr(tSASrc1), 4&
With tSASrc2
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = m
 .pvData = bmIn(2).lpbm
End With
CopyMemory ByVal VarPtrArray(bDibSrc2()), VarPtr(tSASrc2), 4&
'get mode then calc
Select Case bProps(0)
Case 0 'mix
 For i = 0 To m - 1
  With bDibMask(i)
   nNewReserved = (.rgbBlue * 146& + .rgbGreen * 1454& + .rgbRed * 456& _
   + 512&) \ 1024&
  End With
  clr1 = bDibSrc1(i)
  clr2 = bDibSrc2(i)
  With bDib(i)
   .rgbBlue = clr1.rgbBlue + (nNewReserved * (-clr1.rgbBlue + clr2.rgbBlue)) \ 512&
   .rgbGreen = clr1.rgbGreen + (nNewReserved * (-clr1.rgbGreen + clr2.rgbGreen)) \ 512&
   .rgbRed = clr1.rgbRed + (nNewReserved * (-clr1.rgbRed + clr2.rgbRed)) \ 512&
   .rgbReserved = clr1.rgbReserved + (nNewReserved * (-clr1.rgbReserved + clr2.rgbReserved)) \ 512&
  End With
 Next i
Case 1 'add
 For i = 0 To m - 1
  With bDibMask(i)
   nNewReserved = (.rgbBlue * 146& + .rgbGreen * 1454& + .rgbRed * 456& _
   + 512&) \ 1024&
  End With
  clr1 = bDibSrc1(i)
  With bDibSrc2(i)
   nNewBlue = clr1.rgbBlue + (.rgbBlue * nNewReserved) \ 512&
   If nNewBlue > 255 Then nNewBlue = 255
   nNewGreen = clr1.rgbGreen + (.rgbGreen * nNewReserved) \ 512&
   If nNewGreen > 255 Then nNewGreen = 255
   nNewRed = clr1.rgbRed + (.rgbRed * nNewReserved) \ 512&
   If nNewRed > 255 Then nNewRed = 255
   nNewReserved = clr1.rgbReserved + (.rgbReserved * nNewReserved) \ 512&
   If nNewReserved > 255 Then nNewReserved = 255
  End With
  With bDib(i)
   .rgbBlue = nNewBlue
   .rgbGreen = nNewGreen
   .rgbRed = nNewRed
   .rgbReserved = nNewReserved
  End With
 Next i
Case 2 'sub
 For i = 0 To m - 1
  With bDibMask(i)
   nNewReserved = (.rgbBlue * 146& + .rgbGreen * 1454& + .rgbRed * 456& _
   + 512&) \ 1024&
  End With
  clr1 = bDibSrc1(i)
  With bDibSrc2(i)
   nNewBlue = clr1.rgbBlue - (.rgbBlue * nNewReserved) \ 512&
   If nNewBlue < 0 Then nNewBlue = 0
   nNewGreen = clr1.rgbGreen - (.rgbGreen * nNewReserved) \ 512&
   If nNewGreen < 0 Then nNewGreen = 0
   nNewRed = clr1.rgbRed - (.rgbRed * nNewReserved) \ 512&
   If nNewRed < 0 Then nNewRed = 0
   nNewReserved = clr1.rgbReserved - (.rgbReserved * nNewReserved) \ 512&
   If nNewReserved < 0 Then nNewReserved = 0
  End With
  With bDib(i)
   .rgbBlue = nNewBlue
   .rgbGreen = nNewGreen
   .rgbRed = nNewRed
   .rgbReserved = nNewReserved
  End With
 Next i
Case 3 'mul
 For i = 0 To m - 1
  With bDibMask(i)
   'NOT STANDRAD fix the range!!
   nNewReserved = (.rgbBlue * 294& + .rgbGreen * 2920& + .rgbRed * 917& _
   + 512&) \ 1024&
  End With
  clr1 = bDibSrc1(i)
  clr2 = bDibSrc2(i)
  With bDib(i)
   .rgbBlue = (clr1.rgbBlue * nNewReserved * clr2.rgbBlue) \ &H40000
   .rgbGreen = (clr1.rgbGreen * nNewReserved * clr2.rgbGreen) \ &H40000
   .rgbRed = (clr1.rgbRed * nNewReserved * clr2.rgbRed) \ &H40000
   .rgbReserved = (clr1.rgbReserved * nNewReserved * clr2.rgbReserved) \ &H40000
  End With
 Next i
Case 4 'dissolve
 For i = 0 To m - 1
  With bDibMask(i)
   nNewReserved = (.rgbBlue * 146& + .rgbGreen * 1454& + .rgbRed * 456& _
   + 512&) \ 1024&
  End With
  If (cUnk.fRnd2(i, &HBEE0F00D, &HDEADBEEF) And 511&) < nNewReserved Then
   bDib(i) = bDibSrc2(i)
  Else
   bDib(i) = bDibSrc1(i)
  End If
 Next i
End Select
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
ZeroMemory ByVal VarPtrArray(bDibMask()), 4&
ZeroMemory ByVal VarPtrArray(bDibSrc1()), 4&
ZeroMemory ByVal VarPtrArray(bDibSrc2()), 4&
End Sub

Private Sub pCalcRangeColor(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal m As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim clr1 As RGBQUAD, clr2 As RGBQUAD
Dim TheClrTable(255) As RGBQUAD
Dim i As Long, k As Long
Dim nMin As Long, nMax As Long
'init array
CopyMemory ByVal lpbm, ByVal lpbmIn, 4& * m
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = m
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
'init color table
If bProps(0) And 2& Then
 CopyMemory clr1, bProps(5), 4&
 CopyMemory clr2, bProps(1), 4&
Else
 CopyMemory clr1, bProps(1), 4&
 CopyMemory clr2, bProps(5), 4&
End If
nMin = bProps(9)
nMax = bProps(10)
For i = 0 To 255
 If i <= nMin Then
  TheClrTable(i) = clr1
 ElseIf i >= nMax Then
  TheClrTable(i) = clr2
 Else
  pMixColor TheClrTable(i), clr1, clr2, ((i - nMin) * 255&) \ (nMax - nMin)
 End If
Next i
'calc
If bProps(0) And 1& Then 'range
 For i = 0 To m - 1
  With bDib(i)
   k = (.rgbBlue * 146& + .rgbGreen * 1454& + .rgbRed * 456& _
   ) \ 2048&
  End With
  bDib(i) = TheClrTable(k)
 Next i
Else 'adjust
 For i = 0 To m - 1
  With bDib(i)
   .rgbBlue = TheClrTable(.rgbBlue).rgbBlue
   .rgbGreen = TheClrTable(.rgbGreen).rgbGreen
   .rgbRed = TheClrTable(.rgbRed).rgbRed
   .rgbReserved = TheClrTable(.rgbReserved).rgbReserved
  End With
 Next i
End If
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcHSCB(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal m As Long, bProps() As Byte)
Dim bDib() As RGBQUAD, lDib() As Long
Dim tSA As SAFEARRAY2D
Dim i As Long
Dim nHue As Long
Dim nSaturation As Long
Dim TheTable(255) As Byte
Dim f1 As Single, f2 As Single
'////
Dim nMax As Long, nMin As Long, nDelta As Long
Dim nH As Long, nL As Long, nS As Long
'////
'init array
CopyMemory ByVal lpbm, ByVal lpbmIn, 4& * m
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = m
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
'get props
CopyMemory f1, bProps(0), 4&
nHue = f1 * 1536#
CopyMemory f1, bProps(4), 4&
nSaturation = f1 * 256#
'gamma??
CopyMemory f1, bProps(8), 4&
CopyMemory f2, bProps(12), 4&
For i = 0 To 255
 nH = f2 * 255# * ((i / 255#) ^ f1)
 If nH > 255 Then nH = 255
 TheTable(i) = nH
Next i
'calc
For i = 0 To m - 1
 With bDib(i)
  'RGB2HLS
  If .rgbBlue > .rgbGreen Then
   If .rgbBlue > .rgbRed Then
    nMax = .rgbBlue
    If .rgbGreen > .rgbRed Then nMin = .rgbRed Else nMin = .rgbGreen
   Else
    nMax = .rgbRed
    nMin = .rgbGreen
   End If
  ElseIf .rgbRed > .rgbGreen Then
   nMax = .rgbRed
   nMin = .rgbBlue
  Else
   nMax = .rgbGreen
   If .rgbBlue > .rgbRed Then nMin = .rgbRed Else nMin = .rgbBlue
  End If
  If nMax = nMin Then
   nH = 0 '-256 - 1280
   nS = 0 '0-256
   nL = nMax '0-255
  Else
   nL = nMax + nMin 'max=510
   nDelta = nMax - nMin
   If nL < 256 Then
    nS = (nDelta * 256&) \ nL
   Else
    nS = (nDelta * 256&) \ (510 - nL)
   End If
   nL = nL \ 2
   If nMax = .rgbRed Then
    nH = ((-.rgbBlue + .rgbGreen) * 256&) \ nDelta
   ElseIf nMax = .rgbGreen Then
    nH = 512& + ((-.rgbRed + .rgbBlue) * 256&) \ nDelta
   Else
    nH = 1024& + ((-.rgbGreen + .rgbRed) * 256&) \ nDelta
   End If
  End If
  'calc
  nH = nH + nHue
  If nH > 1280 Then nH = nH - 1536
  nS = (nS * nSaturation) \ 256&
  If nS > 256 Then nS = 256
  nL = TheTable(nL)
  'HLS2RGB
  If nS = 0 Then
   .rgbBlue = nL
   .rgbGreen = nL
   .rgbRed = nL
  Else
   If nL < 128 Then
    nMin = (nL * (256 - nS)) \ 256&
   Else
    nMin = nL - (nS * (255 - nL)) \ 256&
   End If
   nMax = nL + nL - nMin
   nDelta = nMax - nMin
   If nH < 256 Then
    .rgbRed = nMax
    If nH < 0 Then
     .rgbGreen = nMin
     .rgbBlue = nMin - (nH * nDelta) \ 256&
    Else
     .rgbBlue = nMin
     .rgbGreen = nMin + (nH * nDelta) \ 256&
    End If
   ElseIf nH < 768 Then
    .rgbGreen = nMax
    nH = nH - 512
    If nH < 0 Then
     .rgbBlue = nMin
     .rgbRed = nMin - (nH * nDelta) \ 256&
    Else
     .rgbRed = nMin
     .rgbBlue = nMin + (nH * nDelta) \ 256&
    End If
   Else
    .rgbBlue = nMax
    nH = nH - 1024
    If nH < 0 Then
     .rgbRed = nMin
     .rgbGreen = nMin - (nH * nDelta) \ 256&
    Else
     .rgbGreen = nMin
     .rgbRed = nMin + (nH * nDelta) \ 256&
    End If
   End If
  End If
 End With
Next i
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

'box blur only :-3
Private Sub pCalcBlur(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Dim bDib() As RGBQUAD, clr As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim i As Long, j As Long
Dim nCount As Long
Dim nXSize As Long, nYSize As Long
Dim nXMax As Long, nYMax As Long 'sum / max
Dim nXClampMode As Long, nYClampMode As Long
Dim f As Single
Dim nAmplify As Long
'////
Dim nSumRed() As Long, nSumGreen() As Long, nSumBlue() As Long, nSumReserved() As Long
Dim nRed As Long, nGreen As Long, nBlue As Long, nReserved As Long
'////
'init array
CopyMemory ByVal lpbm, ByVal lpbmIn, 4& * w * h
With tSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
If w > h Then i = w - 1 Else i = h - 1
ReDim nSumRed(i), nSumGreen(i), nSumBlue(i), nSumReserved(i)
'get data
nCount = bProps(12)
nXClampMode = (nCount And 24&) \ 8&
nYClampMode = (nCount And 96&) \ 32&
nCount = nCount And 7&
'get size
CopyMemory f, bProps(0), 4&
nXSize = f * w
CopyMemory f, bProps(4), 4&
nYSize = f * h
CopyMemory f, bProps(8), 4&
nAmplify = f * 1024#
nXMax = nXSize + nXSize + 1
nYMax = nYSize + nYSize + 1
'start calc
Do While nCount > 0
 nCount = nCount - 1
 'blur X
 If nXSize > 0 Then
  For j = 0 To h - 1
   'calc first sum
   nRed = 0
   nGreen = 0
   nBlue = 0
   nReserved = 0
   For i = -nXSize To nXSize
    pGetColor bDib, w, h, i, j, nXClampMode, clr
    With clr
     nRed = nRed + .rgbRed
     nGreen = nGreen + .rgbGreen
     nBlue = nBlue + .rgbBlue
     nReserved = nReserved + .rgbReserved
    End With
   Next i
   nSumRed(0) = nRed
   nSumGreen(0) = nGreen
   nSumBlue(0) = nBlue
   nSumReserved(0) = nReserved
   'calc sum
   For i = 1 To w - 1
    pGetColor bDib, w, h, i + nXSize, j, nXClampMode, clr
    With clr
     nRed = nRed + .rgbRed
     nGreen = nGreen + .rgbGreen
     nBlue = nBlue + .rgbBlue
     nReserved = nReserved + .rgbReserved
    End With
    pGetColor bDib, w, h, i - nXSize - 1, j, nXClampMode, clr
    With clr
     nRed = nRed - .rgbRed
     nGreen = nGreen - .rgbGreen
     nBlue = nBlue - .rgbBlue
     nReserved = nReserved - .rgbReserved
    End With
    nSumRed(i) = nRed
    nSumGreen(i) = nGreen
    nSumBlue(i) = nBlue
    nSumReserved(i) = nReserved
   Next i
   'divide it
   For i = 0 To w - 1
    With bDib(i, j)
     .rgbRed = nSumRed(i) \ nXMax
     .rgbGreen = nSumGreen(i) \ nXMax
     .rgbBlue = nSumBlue(i) \ nXMax
     .rgbReserved = nSumReserved(i) \ nXMax
    End With
   Next i
  Next j
 End If
 'blur Y
 If nYSize > 0 Then
  For i = 0 To w - 1
   'calc first sum
   nRed = 0
   nGreen = 0
   nBlue = 0
   nReserved = 0
   For j = -nYSize To nYSize
    pGetColor bDib, w, h, i, j, nYClampMode, clr
    With clr
     nRed = nRed + .rgbRed
     nGreen = nGreen + .rgbGreen
     nBlue = nBlue + .rgbBlue
     nReserved = nReserved + .rgbReserved
    End With
   Next j
   nSumRed(0) = nRed
   nSumGreen(0) = nGreen
   nSumBlue(0) = nBlue
   nSumReserved(0) = nReserved
   'calc sum
   For j = 1 To h - 1
    pGetColor bDib, w, h, i, j + nYSize, nYClampMode, clr
    With clr
     nRed = nRed + .rgbRed
     nGreen = nGreen + .rgbGreen
     nBlue = nBlue + .rgbBlue
     nReserved = nReserved + .rgbReserved
    End With
    pGetColor bDib, w, h, i, j - nYSize - 1, nYClampMode, clr
    With clr
     nRed = nRed - .rgbRed
     nGreen = nGreen - .rgbGreen
     nBlue = nBlue - .rgbBlue
     nReserved = nReserved - .rgbReserved
    End With
    nSumRed(j) = nRed
    nSumGreen(j) = nGreen
    nSumBlue(j) = nBlue
    nSumReserved(j) = nReserved
   Next j
   'divide it
   For j = 0 To h - 1
    With bDib(i, j)
     .rgbRed = nSumRed(j) \ nYMax
     .rgbGreen = nSumGreen(j) \ nYMax
     .rgbBlue = nSumBlue(j) \ nYMax
     .rgbReserved = nSumReserved(j) \ nYMax
    End With
   Next j
  Next i
 End If
 'amplify
 If nAmplify <> 1024 Then
  For j = 0 To h - 1
   For i = 0 To w - 1
    With bDib(i, j)
     nRed = (.rgbBlue * nAmplify) \ 1024
     If nRed > 255 Then nRed = 255
     .rgbBlue = nRed
     nRed = (.rgbGreen * nAmplify) \ 1024
     If nRed > 255 Then nRed = 255
     .rgbGreen = nRed
     nRed = (.rgbRed * nAmplify) \ 1024
     If nRed > 255 Then nRed = 255
     .rgbRed = nRed
     nRed = (.rgbReserved * nAmplify) \ 1024
     If nRed > 255 Then nRed = 255
     .rgbReserved = nRed
    End With
   Next i
  Next j
 End If
Loop
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

'mode=??? xmode?? ymode??
Private Sub pGetColor(bDib() As RGBQUAD, ByVal w As Long, ByVal h As Long, ByVal x As Long, ByVal y As Long, ByVal nMode As Long, clr As RGBQUAD)
Select Case nMode
Case 0
 If x >= 0 And y >= 0 And x < w And y < h Then
  clr = bDib(x, y)
 Else
  With clr
   .rgbBlue = 0
   .rgbGreen = 0
   .rgbRed = 0
   .rgbReserved = 255 '?????????
  End With
 End If
Case 1
 clr = bDib(x And (w - 1), y And (h - 1))
Case 2
 If x < 0 Then x = 0 Else If x >= w Then x = w - 1
 If y < 0 Then y = 0 Else If y >= h Then y = h - 1
 clr = bDib(x, y)
Case 3
 If x And w Then x = (-1 - x) And (w - 1) Else x = x And (w - 1)
 If y And h Then y = (-1 - y) And (h - 1) Else y = y And (h - 1)
 clr = bDib(x, y)
End Select
End Sub

Private Sub pGetColorEx(bDib() As RGBQUAD, ByVal w As Long, ByVal h As Long, ByVal x As Long, ByVal y As Long, ByVal xx As Long, ByVal yy As Long, ByVal nXMode As Long, ByVal nYMode As Long, clr As RGBQUAD)
Dim clr00 As RGBQUAD, clr01 As RGBQUAD
Dim clr10 As RGBQUAD, clr11 As RGBQUAD
Dim x2 As Long, y2 As Long
'get x pos
Select Case nXMode
Case 0 'none
 If x < -1 Or x >= w Then x = -2
 If xx Then
  x2 = x + 1
  If x2 >= w Then x2 = -2
 End If
Case 1 'wrap
 If xx Then x2 = (x + 1) And (w - 1)
 x = x And (w - 1)
Case 2 'clamp
 If xx Then
  x2 = x + 1
  If x2 < 0 Then x2 = 0 Else If x2 >= w Then x2 = w - 1
 End If
 If x < 0 Then x = 0 Else If x >= w Then x = w - 1
Case 3 'mirror
 If xx Then
  x2 = x + 1
  If x2 And w Then x2 = (-1 - x2) And (w - 1) Else x2 = x2 And (w - 1)
 End If
 If x And w Then x = (-1 - x) And (w - 1) Else x = x And (w - 1)
End Select
'get y pos
Select Case nYMode
Case 0 'none
 If y < -1 Or y >= h Then y = -2
 If yy Then
  y2 = y + 1
  If y2 >= h Then y2 = -2
 End If
Case 1 'wrap
 If yy Then y2 = (y + 1) And (h - 1)
 y = y And (h - 1)
Case 2 'clamp
 If yy Then
  y2 = y + 1
  If y2 < 0 Then y2 = 0 Else If y2 >= h Then y2 = h - 1
 End If
 If y < 0 Then y = 0 Else If y >= h Then y = h - 1
Case 3 'mirror
 If yy Then
  y2 = y + 1
  If y2 And h Then y2 = (-1 - y2) And (h - 1) Else y2 = y2 And (h - 1)
 End If
 If y And h Then y = (-1 - y) And (h - 1) Else y = y And (h - 1)
End Select
'get color
If x >= 0 And y >= 0 Then clr00 = bDib(x, y) Else clr00.rgbReserved = 255
If yy Then
 If x >= 0 And y2 >= 0 Then clr10 = bDib(x, y2) Else clr10.rgbReserved = 255
End If
If xx Then
 If x2 >= 0 And y >= 0 Then clr01 = bDib(x2, y) Else clr01.rgbReserved = 255
 'mix
 With clr00
  .rgbBlue = .rgbBlue + ((-.rgbBlue + clr01.rgbBlue) * xx) \ 256&
  .rgbGreen = .rgbGreen + ((-.rgbGreen + clr01.rgbGreen) * xx) \ 256&
  .rgbRed = .rgbRed + ((-.rgbRed + clr01.rgbRed) * xx) \ 256&
  .rgbReserved = .rgbReserved + ((-.rgbReserved + clr01.rgbReserved) * xx) \ 256&
 End With
 If yy Then
  If x2 >= 0 And y2 >= 0 Then clr11 = bDib(x2, y2) Else clr11.rgbReserved = 255
  'mix
  With clr10
   .rgbBlue = .rgbBlue + ((-.rgbBlue + clr11.rgbBlue) * xx) \ 256&
   .rgbGreen = .rgbGreen + ((-.rgbGreen + clr11.rgbGreen) * xx) \ 256&
   .rgbRed = .rgbRed + ((-.rgbRed + clr11.rgbRed) * xx) \ 256&
   .rgbReserved = .rgbReserved + ((-.rgbReserved + clr11.rgbReserved) * xx) \ 256&
  End With
 End If
End If
If yy Then
 'mix
 With clr00
  .rgbBlue = .rgbBlue + ((-.rgbBlue + clr10.rgbBlue) * yy) \ 256&
  .rgbGreen = .rgbGreen + ((-.rgbGreen + clr10.rgbGreen) * yy) \ 256&
  .rgbRed = .rgbRed + ((-.rgbRed + clr10.rgbRed) * yy) \ 256&
  .rgbReserved = .rgbReserved + ((-.rgbReserved + clr10.rgbReserved) * yy) \ 256&
 End With
End If
clr = clr00
End Sub

Private Sub pCalcNormals(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim bDibSrc() As RGBQUAD
Dim tSASrc As SAFEARRAY2D
Dim i As Long, j As Long, k As Long
Dim nMode As Long, nScale As Long
Dim f As Single
'////
Dim ii As Long
Dim j0 As Long, j2 As Long
'all 0-512
Dim n00 As Long, n01 As Long, n02 As Long
Dim n10 As Long, n11 As Long, n12 As Long
Dim n20 As Long, n21 As Long, n22 As Long
'vector
Dim x As Long, y As Long, z As Long
'////
Dim TheTable(5121) As Long 'sqr table
'////
'init array
With tSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
With tSASrc
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbmIn
End With
CopyMemory ByVal VarPtrArray(bDibSrc()), VarPtr(tSASrc), 4&
'init table
For i = 0 To 5121
 TheTable(i) = Sqr(&H40000 * (i + 1&))
Next i
'get data
CopyMemory f, bProps(0), 4&
nScale = f * 1024# 'overflow?
nMode = bProps(4)
'calc
For j = 0 To h - 1
 j0 = (j - 1) And (h - 1)
 j2 = (j + 1) And (h - 1)
 'init value
 With bDibSrc(w - 1, j0)
  n01 = (.rgbBlue * 146& + .rgbGreen * 1454& + .rgbRed * 456& _
  + 512&) \ 1024&
 End With
 With bDibSrc(w - 1, j)
  n11 = (.rgbBlue * 146& + .rgbGreen * 1454& + .rgbRed * 456& _
  + 512&) \ 1024&
 End With
 With bDibSrc(w - 1, j2)
  n21 = (.rgbBlue * 146& + .rgbGreen * 1454& + .rgbRed * 456& _
  + 512&) \ 1024&
 End With
 With bDibSrc(0, j0)
  n02 = (.rgbBlue * 146& + .rgbGreen * 1454& + .rgbRed * 456& _
  + 512&) \ 1024&
 End With
 With bDibSrc(0, j)
  n12 = (.rgbBlue * 146& + .rgbGreen * 1454& + .rgbRed * 456& _
  + 512&) \ 1024&
 End With
 With bDibSrc(0, j2)
  n22 = (.rgbBlue * 146& + .rgbGreen * 1454& + .rgbRed * 456& _
  + 512&) \ 1024&
 End With
 For i = 0 To w - 1
  n00 = n01
  n10 = n11
  n20 = n21
  n01 = n02
  n11 = n12
  n21 = n22
  ii = (i + 1) And (w - 1)
  With bDibSrc(ii, j0)
   n02 = (.rgbBlue * 146& + .rgbGreen * 1454& + .rgbRed * 456& _
   + 512&) \ 1024&
  End With
  With bDibSrc(ii, j)
   n12 = (.rgbBlue * 146& + .rgbGreen * 1454& + .rgbRed * 456& _
   + 512&) \ 1024&
  End With
  With bDibSrc(ii, j2)
   n22 = (.rgbBlue * 146& + .rgbGreen * 1454& + .rgbRed * 456& _
   + 512&) \ 1024&
  End With
  'calc vector
  x = ((n00 + n10 + n10 + n20 - n02 - n12 - n12 - n22) * nScale) \ 1024&
  y = ((n00 + n01 + n01 + n02 - n20 - n21 - n21 - n22) * nScale) \ 1024&
  z = x * x + y * y '<= 1342177280
  'calc SQRT
  ii = z \ &H40000
  ii = TheTable(ii) + ((TheTable(ii + 1) - TheTable(ii)) * (z And &H3FFFF)) \ &H40000  'linear interp
  ii = (ii + (z + &H40000) \ ii) \ 2
  'normalize
  x = 128& + ((x * 128&) \ ii)
  If x < 0 Then x = 0 Else If x > 255 Then x = 255
  y = 128& + ((y * 128&) \ ii)
  If y < 0 Then y = 0 Else If y > 255 Then y = 255
  With bDib(i, j)
   If nMode And 2& Then 'tangent map
    .rgbBlue = y
    .rgbGreen = 255 - x
   Else 'normal map
    .rgbBlue = x
    .rgbGreen = y
   End If
   If nMode And 1& Then '3d?
    z = 128& + &H10000 \ ii
    If z > 255 Then z = 255
    .rgbRed = z
   Else
    .rgbRed = 128
   End If
   .rgbReserved = bDibSrc(i, j).rgbReserved
  End With
 Next i
Next j
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
ZeroMemory ByVal VarPtrArray(bDibSrc()), 4&
End Sub

'from internet
Private Sub pCalcColorBalance(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal m As Long, bProps() As Byte)
Dim bDib() As Byte
Dim tSA As SAFEARRAY2D
Dim i As Long, j As Long, k As Long
Dim f1_u8(255) As Long '1->65536
Dim f2_u8(255) As Long
Dim nClrShadows As Long
Dim nClrMidtones As Long
Dim nClrHighlights As Long
'init array
m = m * 4&
CopyMemory ByVal lpbm, ByVal lpbmIn, m
With tSA
 .cbElements = 1
 .cDims = 1
 .Bounds(0).cElements = m
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
'init table
For i = 0 To 255
 'f0_u8(x) = 1.075 - 1/((255.0-x)/16.0 + 1); /* Should this be used somewhere? --jcohen */
 'f1_u8(x) = 1.075 - 1/(x/16.0 + 1);
 f1_u8(i) = 70451 - 1048576 \ (i + 16&)
 'f2_u8(x) = 0.667 * (1 - SQR ((x - 127.0) / 127.0)); // SQR(x) means x*x :-3
 k = i - 127&
 f2_u8(i) = 43691 - (k * k * 1387) \ 512&
Next i
'calc
'      highlights_add_ptr[i] = shadows_sub_ptr[255 - i] = f1_u8 ((double)i);
'      midtones_add_ptr[i] = midtones_sub_ptr[i] = f2_u8 ((double)i);
'      shadows_add_ptr[i] = highlights_sub_ptr[i] = f2_u8 ((double)i);
'////////////////////////////////////////////////////////////////////////////////
'      r_n += cr[SHADOWS] * cyan_red_transfer[SHADOWS][r_n];
'      r_n = BOUNDS (r_n, 0, 255);
'      r_n += cr[MIDTONES] * cyan_red_transfer[MIDTONES][r_n];
'      r_n = BOUNDS (r_n, 0, 255);
'      r_n += cr[HIGHLIGHTS] * cyan_red_transfer[HIGHLIGHTS][r_n];
'      r_n = BOUNDS (r_n, 0, 255);
For j = 0 To 2 'blue->green->red
 nClrShadows = bProps(j) - 128
 nClrMidtones = bProps(j + 4) - 128
 nClrHighlights = bProps(j + 8) - 128
 For i = j To m - 1 Step 4
  k = bDib(i)
  If nClrShadows <> 0 Then
   If nClrShadows > 0 Then
    k = k + (nClrShadows * f2_u8(k)) \ 32768
   Else
    k = k + (nClrShadows * f1_u8(255 - k)) \ 32768
   End If
   If k < 0 Then k = 0 Else If k > 255 Then k = 255
  End If
  If nClrMidtones <> 0 Then
   k = k + (nClrMidtones * f2_u8(k)) \ 32768
   If k < 0 Then k = 0 Else If k > 255 Then k = 255
  End If
  If nClrHighlights <> 0 Then
   If nClrHighlights > 0 Then
    k = k + (nClrHighlights * f1_u8(k)) \ 32768
   Else
    k = k + (nClrHighlights * f2_u8(k)) \ 32768
   End If
   If k < 0 Then k = 0 Else If k > 255 Then k = 255
  End If
  bDib(i) = k
 Next i
Next j
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcDistort(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal lpbmIn2 As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim bDibSrc() As RGBQUAD
Dim tSASrc As SAFEARRAY2D
Dim bDibMap() As RGBQUAD
Dim tSAMap As SAFEARRAY2D
Dim i As Long, j As Long
Dim ii As Long, jj As Long 'actual x,y
Dim x As Long, y As Long '1 -> 256
Dim nXAmount As Long, nYAmount As Long
Dim nXClampMode As Long, nYClampMode As Long
Dim f As Single
'init array
With tSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
With tSASrc
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbmIn
End With
CopyMemory ByVal VarPtrArray(bDibSrc()), VarPtr(tSASrc), 4&
With tSAMap
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbmIn2
End With
CopyMemory ByVal VarPtrArray(bDibMap()), VarPtr(tSAMap), 4&
'get data
nXClampMode = bProps(4)
nYClampMode = (nXClampMode And 12&) \ 4&
nXClampMode = nXClampMode And 3&
CopyMemory f, bProps(0), 4&
'get size
nXAmount = f * (w * 256&)
nYAmount = f * (h * 256&)
'start calc
For j = 0 To h - 1
 For i = 0 To w - 1
  With bDibMap(i, j)
   x = ((.rgbBlue - 128&) * nXAmount + 64&) \ 128&
   y = ((.rgbGreen - 128&) * nYAmount + 64&) \ 128&
  End With
  ii = i + (x And &HFFFFFF00) \ 256&
  jj = j + (y And &HFFFFFF00) \ 256&
  pGetColorEx bDibSrc, w, h, ii, jj, x And 255&, y And 255&, nXClampMode, nYClampMode, bDib(i, j)
 Next i
Next j
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
ZeroMemory ByVal VarPtrArray(bDibSrc()), 4&
ZeroMemory ByVal VarPtrArray(bDibMap()), 4&
End Sub

'///////////////////////////////////////Phong reflection model
'We first define, for each light source in the scene to be rendered, the components i_s and i_d, where
'these are the intensities (often as RGB values) of the specular and diffuse components of the light
'sources respectively. A single i_a term controls the ambient lighting; it is sometimes computed as a
'sum of contributions from the light sources.
'
'k_s: specular reflection constant, the ratio of reflection of the specular term of incoming light
'
'k_d: diffuse reflection constant, the ratio of reflection of the diffuse term of incoming light
'
'k_a: ambient reflection constant, the ratio of reflection of the ambient term present in all
'points in the scene rendered
'
'α: is a shininess constant for this material, which decides how "evenly" light is
'reflected from a shiny spot, and is very large for most surfaces, on the order of
'50, getting larger the more mirror-like they are.
'
'We further define lights as the set of all light sources, Ｌ is the direction vector from the point on the
'surface toward each light source, Ｎ is the normal at this point of the surface, Ｒ is the direction that a
'perfectly reflected ray of light (represented as a vector) would take from this point of the surface, and
'Ｖ is the direction towards the viewer (such as a virtual camera).
'Then the shade value for each surface point Ip is calculated using this equation, which is the Phong
'reflection model:
'
'I_p=K_a*I_a+Σ(k_d*(Ｌ·Ｎ)*i_d + k_s*((Ｒ·Ｖ)^α)*i_s)
'///////////////////////////////////////Blinn–Phong shading model
'If we instead calculate a halfway vector between the viewer and light-source vectors
'     Ｌ+Ｖ
'Ｈ= -------
'    |Ｌ+Ｖ|
'we can replace Ｒ·Ｖ with Ｎ·Ｈ, where Ｎ is the normalized surface normal.

Private Sub pCalcBump(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal lpbmIn2 As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim bDibMap() As RGBQUAD
Dim tSAMap As SAFEARRAY2D
Dim i As Long, j As Long, k As Long, lp As Long
Dim f1 As Single, f2 As Single
Dim nMode As Long
Dim nNewBlue As Long, nNewGreen As Long, nNewRed As Long
Dim nAmbientBlue As Long, nAmbientGreen As Long, nAmbientRed As Long
Dim nDiffuseBlue As Long, nDiffuseGreen As Long, nDiffuseRed As Long
'vectors '1 -> 128 ??
Dim L_x As Long, L_y As Long, L_z As Long
Dim N_x As Long, N_y As Long, N_z As Long
Dim H_x As Long, H_y As Long, H_z As Long
'SQRT Table
Dim TheSqrTable(2049) As Long '0,64,128,...
'specular table
Dim TheClrTable(512) As RGBQUAD   '1 -> 512
'init array
lp = w * h
CopyMemory ByVal lpbm, ByVal lpbmIn, lp * 4&
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = lp
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
With tSAMap
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = lp
 .pvData = lpbmIn2
End With
CopyMemory ByVal VarPtrArray(bDibMap()), VarPtr(tSAMap), 4&
'init SQRT table (??)
For i = 1 To 2049
 TheSqrTable(i) = Sqr(i * 64&)
Next i
'init specular table
CopyMemory f1, bProps(45), 4&
CopyMemory f2, bProps(49), 4&
For i = 1 To 512
 k = 512# * f2 * (i / 512) ^ f1
 If k > 512 Then k = 512
 With TheClrTable(i)
  .rgbBlue = (bProps(41) * k + 256) \ 512
  .rgbGreen = (bProps(42) * k + 256) \ 512
  .rgbRed = (bProps(43) * k + 256) \ 512
 End With
Next i
'get color
nAmbientBlue = (bProps(25) * 1024&) \ 255&
nAmbientGreen = (bProps(26) * 1024&) \ 255&
nAmbientRed = (bProps(27) * 1024&) \ 255&
CopyMemory f1, bProps(37), 4&
k = 512# * f1
nDiffuseBlue = (bProps(21) * k) \ 512&
nDiffuseGreen = (bProps(22) * k) \ 512&
nDiffuseRed = (bProps(23) * k) \ 512&
'get data
nMode = bProps(0)
If nMode = 2 Then 'directional
 CopyMemory f1, bProps(13), 4&
 CopyMemory f2, bProps(17), 4&
 f1 = f1 * 二π
 f2 = f2 * π
 L_z = 128# * Sin(f2)
 f2 = 128# * Cos(f2)
 L_x = Sin(f1) * f2
 L_y = Cos(f1) * f2
 'calc H
 H_z = L_z + 128&
 H_x = L_x * L_x + L_y * L_y + H_z * H_z
 'calc SQRT
 If H_x > 256 Then
  k = H_x \ 64&
  H_y = TheSqrTable(k)
  H_y = H_y + ((TheSqrTable(k + 1) - H_y) * (H_x And 63&)) \ 64& 'linear interp
  H_y = (H_y + H_x \ H_y) \ 2
 Else
  H_y = (TheSqrTable(H_x) + 4) \ 8
 End If
 If H_y > 0 Then
  H_x = (L_x * 128&) \ H_y
  H_z = (H_z * 128&) \ H_y
  H_y = (L_y * 128&) \ H_y
 Else 'error
  H_x = 0
  H_y = 0
  H_z = 0
 End If
Else
 'TODO:
End If
'start calc
lp = 0
For j = 0 To h - 1
 For i = 0 To w - 1
  If nMode < 2 Then
   'TODO:point
   If nMode = 0 Then
    'TODO:spot
   End If
  End If
  With bDibMap(lp)
   N_x = .rgbBlue - 128&
   N_y = .rgbGreen - 128&
   N_z = .rgbRed - 128&
  End With
  'calc ambient
  With bDib(lp) '255!! :-3
   nNewBlue = (.rgbBlue * nAmbientBlue) \ 1024&
   nNewGreen = (.rgbGreen * nAmbientGreen) \ 1024&
   nNewRed = (.rgbRed * nAmbientRed) \ 1024&
  End With
  'calc diffuse
  k = L_x * N_x + L_y * N_y + L_z * N_z
  nNewBlue = nNewBlue + (nDiffuseBlue * k) \ 16384&
  nNewGreen = nNewGreen + (nDiffuseGreen * k) \ 16384&
  nNewRed = nNewRed + (nDiffuseRed * k) \ 16384&
  'calc specular
  k = (H_x * N_x + H_y * N_y + H_z * N_z) \ 32&
  If k > 0 Then
   If k > 512 Then k = 512
   With TheClrTable(k)
    nNewBlue = nNewBlue + .rgbBlue
    nNewGreen = nNewGreen + .rgbGreen
    nNewRed = nNewRed + .rgbRed
   End With
  End If
  'over
  If nNewBlue < 0 Then nNewBlue = 0 Else If nNewBlue > 255 Then nNewBlue = 255
  If nNewGreen < 0 Then nNewGreen = 0 Else If nNewGreen > 255 Then nNewGreen = 255
  If nNewRed < 0 Then nNewRed = 0 Else If nNewRed > 255 Then nNewRed = 255
  With bDib(lp)
   .rgbBlue = nNewBlue
   .rgbGreen = nNewGreen
   .rgbRed = nNewRed
  End With
  lp = lp + 1
 Next i
Next j
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
ZeroMemory ByVal VarPtrArray(bDibMap()), 4&
End Sub

Private Sub pCalcRotZoom(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal w As Long, ByVal h As Long, ByVal nSrcWidth As Long, ByVal nSrcHeight As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim bDibSrc() As RGBQUAD
Dim tSASrc As SAFEARRAY2D
Dim i As Long, j As Long
Dim x As Long, y As Long, xx As Long, yy As Long '1 -> 1024
Dim xi As Long, xj As Long, yi As Long, yj As Long
Dim nXClampMode As Long, nYClampMode As Long
Dim f(5) As Single
'init array
With tSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
With tSASrc
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = nSrcHeight
 .Bounds(1).cElements = nSrcWidth
 .pvData = lpbmIn
End With
CopyMemory ByVal VarPtrArray(bDibSrc()), VarPtr(tSASrc), 4&
'get data
nXClampMode = bProps(24)
nYClampMode = (nXClampMode And 12&) \ 4&
nXClampMode = nXClampMode And 3&
'calc transform 'TODO:Resize!!! change code
On Error Resume Next
CopyMemory f(0), bProps(0), 24&
f(0) = f(0) * 二π
'///resize?
f(1) = f(1) / nSrcWidth
f(2) = f(2) / nSrcHeight
'///
xi = (Cos(f(0)) / (f(1) * w)) * 1024#
xj = (Sin(f(0)) / (f(1) * h)) * 1024#
yi = (-Sin(f(0)) / (f(2) * w)) * 1024#
yj = (Cos(f(0)) / (f(2) * h)) * 1024#
x = f(5) * 1024#
xi = xi + (yi * x) \ 1024& '??
xj = xj + (yj * x) \ 1024& '??
x = (-w) \ 2&
y = (-h) \ 2&
xx = x * xi + y * xj
yy = x * yi + y * yj
x = f(3) * nSrcWidth
y = f(4) * nSrcHeight
xx = xx + x * 1024&
yy = yy + y * 1024&
On Error GoTo 0
'calc
For j = 0 To h - 1
 x = xx
 y = yy
 For i = 0 To w - 1
  pGetColorEx bDibSrc, nSrcWidth, nSrcHeight, (x And &HFFFFFC00) \ 1024, (y And &HFFFFFC00) \ 1024, (x And 1023&) \ 4, (y And 1023&) \ 4, _
  nXClampMode, nYClampMode, bDib(i, j)
  x = x + xi
  y = y + yi
 Next i
 xx = xx + xj
 yy = yy + yj
Next j
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
ZeroMemory ByVal VarPtrArray(bDibSrc()), 4&
End Sub

Private Sub pCalcRotMul(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
On Error Resume Next
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim bDibSrc() As RGBQUAD
Dim i As Long, j As Long, k As Long, m As Long
Dim x As Long, y As Long, xx As Long, yy As Long '1 -> 1024
Dim xi As Long, xj As Long, yi As Long, yj As Long
Dim nXClampMode As Long, nYClampMode As Long
Dim nMode As Long
Dim bRecursive As Boolean, nCount As Byte, nIndex As Byte
Dim clr As RGBQUAD
Dim nScaleBlue As Long, nScaleGreen As Long, nScaleRed As Long, nScaleReserved As Long
Dim f(6) As Single
'init array
With tSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
ReDim bDibSrc(w - 1, h - 1)
m = w * h * 4&
CopyMemory bDibSrc(0, 0), ByVal lpbmIn, m
'get data
nXClampMode = bProps(24)
nYClampMode = (nXClampMode And 12&) \ 4&
nMode = (nXClampMode And &HF0&) \ 16&
nXClampMode = nXClampMode And 3&
nCount = bProps(29)
bRecursive = nCount And 128&
nCount = nCount And 127&
'calc pre-adjust
nScaleBlue = (bProps(25) * 1024&) \ 255
nScaleGreen = (bProps(26) * 1024&) \ 255
nScaleRed = (bProps(27) * 1024&) \ 255
nScaleReserved = (bProps(28) * 1024&) \ 255
For j = 0 To w - 1
 For i = 0 To h - 1
  With bDibSrc(i, j)
   .rgbBlue = (.rgbBlue * nScaleBlue) \ 1024&
   .rgbGreen = (.rgbGreen * nScaleGreen) \ 1024&
   .rgbRed = (.rgbRed * nScaleRed) \ 1024&
   .rgbReserved = (.rgbReserved * nScaleReserved) \ 1024&
  End With
 Next i
Next j
CopyMemory ByVal lpbm, bDibSrc(0, 0), m
'calc rotate mul
If nMode = 6 Then 'alpha
 nScaleBlue = (bProps(30) * 149504) \ 255
 nScaleGreen = (bProps(31) * 1488896) \ 255
 nScaleRed = (bProps(32) * 466944) \ 255
Else
 If nMode = 4 Or nMode = 7 Or nMode = 8 Then 'mul,AlphaBlend,add smooth
  nScaleReserved = 1030
 Else
  nScaleReserved = 1024
 End If
 nScaleBlue = (bProps(30) * nScaleReserved) \ 255
 nScaleGreen = (bProps(31) * nScaleReserved) \ 255
 nScaleRed = (bProps(32) * nScaleReserved) \ 255
 nScaleReserved = (bProps(33) * nScaleReserved) \ 255
 If nMode = 2 Or nMode = 3 Then 'sub clamp,sub wrap
  nScaleBlue = -nScaleBlue
  nScaleGreen = -nScaleGreen
  nScaleRed = -nScaleRed
  nScaleReserved = -nScaleReserved
  nMode = nMode - 2
 End If
End If
For nIndex = 1 To nCount
 'calc transform
 CopyMemory f(0), bProps(0), 24&
 If bRecursive Then
  f(6) = 2 ^ (-nIndex)
  If f(1) < 0 Then f(1) = -(-f(1)) ^ f(6) Else f(1) = f(1) ^ f(6)
  If f(2) < 0 Then f(2) = -(-f(2)) ^ f(6) Else f(2) = f(2) ^ f(6)
 Else
  f(6) = nIndex / nCount
  f(1) = 1 + (f(1) - 1) * f(6)
  f(2) = 1 + (f(2) - 1) * f(6)
 End If
 f(0) = f(0) * f(6) * 二π
 xi = (Cos(f(0)) / f(1)) * 1024#
 xj = (Sin(f(0)) / f(2)) * 1024#
 yi = (-Sin(f(0)) / f(1)) * 1024#
 yj = (Cos(f(0)) / f(2)) * 1024#
 x = f(5) * f(6) * 1024#
 xi = xi + (yi * x) \ 1024& '??
 xj = xj + (yj * x) \ 1024& '??
 x = (-w) \ 2&
 y = (-h) \ 2&
 xx = x * xi + y * xj
 yy = x * yi + y * yj
 f(3) = 0.5 + (f(3) - 0.5) * f(6)
 f(4) = 0.5 + (f(4) - 0.5) * f(6)
 x = f(3) * w
 y = f(4) * h
 xx = xx + x * 1024&
 yy = yy + y * 1024&
 'calc
 For j = 0 To w - 1
  x = xx
  y = yy
  For i = 0 To h - 1
   pGetColorEx bDibSrc, w, h, (x And &HFFFFFC00) \ 1024, (y And &HFFFFFC00) \ 1024, (x And 1023&) \ 4, (y And 1023&) \ 4, _
   nXClampMode, nYClampMode, clr
   'combine color
   With bDib(i, j)
    Select Case nMode
    Case 0 'add clamp,sub clamp
     k = .rgbBlue + (nScaleBlue * clr.rgbBlue) \ 1024&
     If k < 0 Then k = 0 Else If k > 255 Then k = 255
     .rgbBlue = k
     k = .rgbGreen + (nScaleGreen * clr.rgbGreen) \ 1024&
     If k < 0 Then k = 0 Else If k > 255 Then k = 255
     .rgbGreen = k
     k = .rgbRed + (nScaleRed * clr.rgbRed) \ 1024&
     If k < 0 Then k = 0 Else If k > 255 Then k = 255
     .rgbRed = k
     k = .rgbReserved + (nScaleReserved * clr.rgbReserved) \ 1024&
     If k < 0 Then k = 0 Else If k > 255 Then k = 255
     .rgbReserved = k
    Case 1 'add wrap,sub wrap
     If m_bIsInIDE Then
      .rgbBlue = (.rgbBlue + (nScaleBlue * clr.rgbBlue) \ 1024&) And 255
      .rgbGreen = (.rgbGreen + (nScaleGreen * clr.rgbGreen) \ 1024&) And 255
      .rgbRed = (.rgbRed + (nScaleRed * clr.rgbRed) \ 1024&) And 255
      .rgbReserved = (.rgbReserved + (nScaleReserved * clr.rgbReserved) \ 1024&) And 255
     Else
      .rgbBlue = (.rgbBlue + (nScaleBlue * clr.rgbBlue) \ 1024&)
      .rgbGreen = (.rgbGreen + (nScaleGreen * clr.rgbGreen) \ 1024&)
      .rgbRed = (.rgbRed + (nScaleRed * clr.rgbRed) \ 1024&)
      .rgbReserved = (.rgbReserved + (nScaleReserved * clr.rgbReserved) \ 1024&)
     End If
    Case 4 'mul
     .rgbBlue = (.rgbBlue * nScaleBlue * clr.rgbBlue) \ &H40000
     .rgbGreen = (.rgbGreen * nScaleGreen * clr.rgbGreen) \ &H40000
     .rgbRed = (.rgbRed * nScaleRed * clr.rgbRed) \ &H40000
     .rgbReserved = (.rgbReserved * nScaleReserved * clr.rgbReserved) \ &H40000
    Case 5 'diff
     k = .rgbBlue - (nScaleBlue * clr.rgbBlue) \ 1024&
     If k < 0 Then k = -k
     .rgbBlue = k
     k = .rgbGreen - (nScaleGreen * clr.rgbGreen) \ 1024&
     If k < 0 Then k = -k
     .rgbGreen = k
     k = .rgbRed - (nScaleRed * clr.rgbRed) \ 1024&
     If k < 0 Then k = -k
     .rgbRed = k
     k = .rgbReserved - (nScaleReserved * clr.rgbReserved) \ 1024&
     If k < 0 Then k = -k
     .rgbReserved = k
    Case 6 'alpha
     k = (clr.rgbBlue * nScaleBlue + clr.rgbGreen * nScaleGreen + clr.rgbRed * nScaleRed _
     + 524288) \ 1048576 '0-512
     .rgbBlue = (.rgbBlue * k) \ 512&
     .rgbGreen = (.rgbGreen * k) \ 512&
     .rgbRed = (.rgbRed * k) \ 512&
     .rgbReserved = (.rgbReserved * k) \ 512&
    Case 7 'AlphaBlend
     nScaleRed = 256& - (clr.rgbReserved * nScaleReserved) \ 1024&
     k = clr.rgbBlue + (.rgbBlue * nScaleRed) \ 256&
     If k > 255 Then k = 255
     .rgbBlue = k
     k = clr.rgbGreen + (.rgbGreen * nScaleRed) \ 256&
     If k > 255 Then k = 255
     .rgbGreen = k
     k = clr.rgbRed + (.rgbRed * nScaleRed) \ 256&
     If k > 255 Then k = 255
     .rgbRed = k
     k = clr.rgbReserved + (.rgbReserved * nScaleReserved * nScaleRed) \ &H40000
     If k > 255 Then k = 255
     .rgbReserved = k
    Case 8 'add smooth
     .rgbBlue = .rgbBlue + ((255& - .rgbBlue) * nScaleBlue * clr.rgbBlue) \ &H40000
     .rgbGreen = .rgbGreen + ((255& - .rgbGreen) * nScaleGreen * clr.rgbGreen) \ &H40000
     .rgbRed = .rgbRed + ((255& - .rgbRed) * nScaleRed * clr.rgbRed) \ &H40000
     .rgbReserved = .rgbReserved + ((255& - .rgbReserved) * nScaleReserved * clr.rgbReserved) \ &H40000
    Case 9 'min
     If .rgbBlue > clr.rgbBlue Then .rgbBlue = clr.rgbBlue
     If .rgbGreen > clr.rgbGreen Then .rgbGreen = clr.rgbGreen
     If .rgbRed > clr.rgbRed Then .rgbRed = clr.rgbRed
     If .rgbReserved > clr.rgbReserved Then .rgbReserved = clr.rgbReserved
    Case 10 'max
     If .rgbBlue < clr.rgbBlue Then .rgbBlue = clr.rgbBlue
     If .rgbGreen < clr.rgbGreen Then .rgbGreen = clr.rgbGreen
     If .rgbRed < clr.rgbRed Then .rgbRed = clr.rgbRed
     If .rgbReserved < clr.rgbReserved Then .rgbReserved = clr.rgbReserved
    End Select
   End With
   'get next
   x = x + xi
   y = y + yi
  Next i
  xx = xx + xj
  yy = yy + yj
 Next j
 'recursive?
 If bRecursive Then
  CopyMemory bDibSrc(0, 0), ByVal lpbm, m
 End If
Next nIndex
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcParticle(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal lpbmIn2 As Long, ByVal lpbmIn3 As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
On Error Resume Next
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim bDibSrc() As RGBQUAD
Dim tSASrc As SAFEARRAY2D
Dim bDibSrc2() As RGBQUAD
Dim tSASrc2 As SAFEARRAY2D
Dim i As Long, j As Long, k As Long
Dim nLeft As Long, nTop As Long, nRight As Long, nBottom As Long
Dim x As Long, y As Long, xx As Long, yy As Long '1 -> 1024
Dim xi As Long, xj As Long ', yi As Long, yj As Long
Dim nMode As Long, nIndex As Long, nCount As Long
Dim nMode2 As Long
Dim TheSeed As Long
Dim clr As RGBQUAD, clr1 As RGBQUAD, clr2 As RGBQUAD
Dim clr3 As RGBQUAD
Dim nScaleBlue As Long, nScaleGreen As Long, nScaleRed As Long, nScaleReserved As Long
Dim TheTable(255) As Byte
Dim TheTable2(255) As Byte, nPercent As Long
Dim f(9) As Single, xFloat As Single, yFloat As Single
'init array
CopyMemory ByVal lpbm, ByVal lpbmIn, w * h * 4&
With tSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
With tSASrc
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbmIn2
End With
CopyMemory ByVal VarPtrArray(bDibSrc()), VarPtr(tSASrc), 4&
If lpbmIn3 Then
 With tSASrc2
  .cbElements = 4
  .cDims = 2
  .Bounds(0).cElements = h
  .Bounds(1).cElements = w
  .pvData = lpbmIn3
 End With
 CopyMemory ByVal VarPtrArray(bDibSrc2()), VarPtr(tSASrc2), 4&
End If
'get data
nMode = bProps(7)
nMode2 = (nMode And &HF0&) \ &H10&
nMode = nMode And &HF&
nCount = bProps(0)
pGetSeed bProps, 20, TheSeed
CopyMemory f(3), bProps(1), 4& 'size
f(4) = bProps(5) / 255# 'size variation
f(5) = bProps(6) / 255# 'spin variation
CopyMemory f(6), bProps(22), 16& 'center,radius
CopyMemory clr1, bProps(8), 4&
CopyMemory clr2, bProps(12), 4&
'init table
CopyMemory f(0), bProps(16), 4&
For i = 0 To 255
 TheTable(i) = 255# * (i / 255#) ^ f(0)
Next i
'calc transform
For nIndex = 1 To nCount
 'get spin
 f(0) = π * cUnk.fRnd2Float(&HF00D&, nIndex, TheSeed) * f(5)
 'get size
 f(1) = f(3) * (1# + f(4) * cUnk.fRnd2Float(&HABCD&, nIndex, TheSeed))
 'calc size
 x = w * f(1)
 y = h * f(1)
 'multi-particle?
 If lpbmIn3 Then
  nPercent = (cUnk.fRnd2(&H1423&, nIndex, TheSeed) And &H7F&) - 127& + bProps(38) '127??
  If nPercent > 128 Then nPercent = 128 Else If nPercent < 0 Then nPercent = 0
 End If
 'calc position
 Select Case nMode2 And 7&
 Case 0 'box
  '-1 to 1
  nLeft = Int(w * (f(6) + f(8) * cUnk.fRnd2Float(&HA1288, nIndex, TheSeed)))
  nTop = Int(h * (f(7) + f(9) * cUnk.fRnd2Float(&HA2008, nIndex, TheSeed)))
 Case 1 'circle
  i = 0
  Do
   xFloat = cUnk.fRnd2Float(&HA1288 + i, nIndex, TheSeed)
   yFloat = cUnk.fRnd2Float(&HA2008 + i, nIndex, TheSeed)
   i = i + 1
  Loop Until xFloat * xFloat + yFloat * yFloat < 1
  nLeft = Int(w * (f(6) + f(8) * xFloat))
  nTop = Int(h * (f(7) + f(9) * yFloat))
 Case 2 'gauss
  'a stupid algorithm
  xFloat = 0
  yFloat = 0
  For i = 1 To 16
   xFloat = xFloat + cUnk.fRnd2Float(&HA1288 + i, nIndex, TheSeed)
   yFloat = yFloat + cUnk.fRnd2Float(&HA2008 + i, nIndex, TheSeed)
  Next i
  nLeft = Int(w * (f(6) + f(8) * xFloat / 8#)) '8?? 16??
  nTop = Int(h * (f(7) + f(9) * yFloat / 8#))
 End Select
 'calc bound
 nRight = nLeft + x - 1 '??
 nBottom = nTop + y - 1 '??
 nLeft = nLeft - x
 nTop = nTop - y
 xi = (Cos(f(0)) / f(1)) * 512#
 xj = (Sin(f(0)) / f(1)) * 512#
 'yi = -xj, yj = xi
 xx = w * 512& - x * xi - y * xj
 yy = h * 512& + x * xj - y * xi
 'get scale
 If nMode = 4 Or nMode = 6 Or nMode = 7 Then 'mul,AlphaBlend,add smooth
  nScaleReserved = 1030
 Else
  nScaleReserved = 1024
 End If
 i = TheTable(cUnk.fRnd2(&HA1800, nIndex, TheSeed) And &HFF&)
 With clr1
  k = .rgbBlue + (i * (-.rgbBlue + clr2.rgbBlue)) \ 255
  nScaleBlue = (k * nScaleReserved) \ 255
  k = .rgbGreen + (i * (-.rgbGreen + clr2.rgbGreen)) \ 255
  nScaleGreen = (k * nScaleReserved) \ 255
  k = .rgbRed + (i * (-.rgbRed + clr2.rgbRed)) \ 255
  nScaleRed = (k * nScaleReserved) \ 255
  k = .rgbReserved + (i * (-.rgbReserved + clr2.rgbReserved)) \ 255
  nScaleReserved = (k * nScaleReserved) \ 255
 End With
 If nMode = 2 Or nMode = 3 Then 'sub clamp,sub wrap
  nScaleBlue = -nScaleBlue
  nScaleGreen = -nScaleGreen
  nScaleRed = -nScaleRed
  nScaleReserved = -nScaleReserved
 End If
 'calc
 For j = nTop To nBottom
  x = xx
  y = yy
  For i = nLeft To nRight
   'multi-particle?
   If lpbmIn3 Then
    If nMode2 And 8& Then 'mix
     pGetColorEx bDibSrc, w, h, (x And &HFFFFFC00) \ 1024, (y And &HFFFFFC00) \ 1024, (x And 1023&) \ 4, (y And 1023&) \ 4, _
     2, 2, clr
     pGetColorEx bDibSrc2, w, h, (x And &HFFFFFC00) \ 1024, (y And &HFFFFFC00) \ 1024, (x And 1023&) \ 4, (y And 1023&) \ 4, _
     2, 2, clr3
     With clr
      .rgbBlue = .rgbBlue + ((-.rgbBlue + clr3.rgbBlue) * nPercent) \ 128&
      .rgbGreen = .rgbGreen + ((-.rgbGreen + clr3.rgbGreen) * nPercent) \ 128&
      .rgbRed = .rgbRed + ((-.rgbRed + clr3.rgbRed) * nPercent) \ 128&
      .rgbReserved = .rgbReserved + ((-.rgbReserved + clr3.rgbReserved) * nPercent) \ 128&
     End With
    ElseIf nPercent > 64 Then
     pGetColorEx bDibSrc2, w, h, (x And &HFFFFFC00) \ 1024, (y And &HFFFFFC00) \ 1024, (x And 1023&) \ 4, (y And 1023&) \ 4, _
     2, 2, clr
    Else
     pGetColorEx bDibSrc, w, h, (x And &HFFFFFC00) \ 1024, (y And &HFFFFFC00) \ 1024, (x And 1023&) \ 4, (y And 1023&) \ 4, _
     2, 2, clr
    End If
   Else
    pGetColorEx bDibSrc, w, h, (x And &HFFFFFC00) \ 1024, (y And &HFFFFFC00) \ 1024, (x And 1023&) \ 4, (y And 1023&) \ 4, _
    2, 2, clr
   End If
   'combine color
   With bDib(i And (w - 1), j And (h - 1))
    Select Case nMode
    Case 0, 2 'add clamp,sub clamp
     k = .rgbBlue + (nScaleBlue * clr.rgbBlue) \ 1024&
     If k < 0 Then k = 0 Else If k > 255 Then k = 255
     .rgbBlue = k
     k = .rgbGreen + (nScaleGreen * clr.rgbGreen) \ 1024&
     If k < 0 Then k = 0 Else If k > 255 Then k = 255
     .rgbGreen = k
     k = .rgbRed + (nScaleRed * clr.rgbRed) \ 1024&
     If k < 0 Then k = 0 Else If k > 255 Then k = 255
     .rgbRed = k
     k = .rgbReserved + (nScaleReserved * clr.rgbReserved) \ 1024&
     If k < 0 Then k = 0 Else If k > 255 Then k = 255
     .rgbReserved = k
    Case 1, 3 'add wrap,sub wrap
     If m_bIsInIDE Then
      .rgbBlue = (.rgbBlue + (nScaleBlue * clr.rgbBlue) \ 1024&) And 255
      .rgbGreen = (.rgbGreen + (nScaleGreen * clr.rgbGreen) \ 1024&) And 255
      .rgbRed = (.rgbRed + (nScaleRed * clr.rgbRed) \ 1024&) And 255
      .rgbReserved = (.rgbReserved + (nScaleReserved * clr.rgbReserved) \ 1024&) And 255
     Else
      .rgbBlue = (.rgbBlue + (nScaleBlue * clr.rgbBlue) \ 1024&)
      .rgbGreen = (.rgbGreen + (nScaleGreen * clr.rgbGreen) \ 1024&)
      .rgbRed = (.rgbRed + (nScaleRed * clr.rgbRed) \ 1024&)
      .rgbReserved = (.rgbReserved + (nScaleReserved * clr.rgbReserved) \ 1024&)
     End If
    Case 4 'mul
     .rgbBlue = (.rgbBlue * nScaleBlue * clr.rgbBlue) \ &H40000
     .rgbGreen = (.rgbGreen * nScaleGreen * clr.rgbGreen) \ &H40000
     .rgbRed = (.rgbRed * nScaleRed * clr.rgbRed) \ &H40000
     .rgbReserved = (.rgbReserved * nScaleReserved * clr.rgbReserved) \ &H40000
    Case 5 'diff
     k = .rgbBlue - (nScaleBlue * clr.rgbBlue) \ 1024&
     If k < 0 Then k = -k
     .rgbBlue = k
     k = .rgbGreen - (nScaleGreen * clr.rgbGreen) \ 1024&
     If k < 0 Then k = -k
     .rgbGreen = k
     k = .rgbRed - (nScaleRed * clr.rgbRed) \ 1024&
     If k < 0 Then k = -k
     .rgbRed = k
     k = .rgbReserved - (nScaleReserved * clr.rgbReserved) \ 1024&
     If k < 0 Then k = -k
     .rgbReserved = k
    Case 6 'AlphaBlend
     nScaleRed = 256& - (clr.rgbReserved * nScaleReserved) \ 1024&
     k = clr.rgbBlue + (.rgbBlue * nScaleRed) \ 256&
     If k > 255 Then k = 255
     .rgbBlue = k
     k = clr.rgbGreen + (.rgbGreen * nScaleRed) \ 256&
     If k > 255 Then k = 255
     .rgbGreen = k
     k = clr.rgbRed + (.rgbRed * nScaleRed) \ 256&
     If k > 255 Then k = 255
     .rgbRed = k
     k = clr.rgbReserved + (.rgbReserved * nScaleReserved * nScaleRed) \ &H40000
     If k > 255 Then k = 255
     .rgbReserved = k
    Case 7 'add smooth
     .rgbBlue = .rgbBlue + ((255& - .rgbBlue) * nScaleBlue * clr.rgbBlue) \ &H40000
     .rgbGreen = .rgbGreen + ((255& - .rgbGreen) * nScaleGreen * clr.rgbGreen) \ &H40000
     .rgbRed = .rgbRed + ((255& - .rgbRed) * nScaleRed * clr.rgbRed) \ &H40000
     .rgbReserved = .rgbReserved + ((255& - .rgbReserved) * nScaleReserved * clr.rgbReserved) \ &H40000
    Case 8 'min
     If .rgbBlue > clr.rgbBlue Then .rgbBlue = clr.rgbBlue
     If .rgbGreen > clr.rgbGreen Then .rgbGreen = clr.rgbGreen
     If .rgbRed > clr.rgbRed Then .rgbRed = clr.rgbRed
     If .rgbReserved > clr.rgbReserved Then .rgbReserved = clr.rgbReserved
    Case 9 'max
     If .rgbBlue < clr.rgbBlue Then .rgbBlue = clr.rgbBlue
     If .rgbGreen < clr.rgbGreen Then .rgbGreen = clr.rgbGreen
     If .rgbRed < clr.rgbRed Then .rgbRed = clr.rgbRed
     If .rgbReserved < clr.rgbReserved Then .rgbReserved = clr.rgbReserved
    End Select
   End With
   'get next
   x = x + xi
   y = y - xj
  Next i
  xx = xx + xj
  yy = yy + xi
 Next j
Next nIndex
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
ZeroMemory ByVal VarPtrArray(bDibSrc()), 4&
If lpbmIn3 Then ZeroMemory ByVal VarPtrArray(bDibSrc2()), 4&
End Sub

'just a FLOOD-FILL algorithm...
Private Sub pCalcSegment(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal lpbmIn2 As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim bDibSrc() As RGBQUAD, clr As RGBQUAD
Dim tSASrc As SAFEARRAY2D
Dim i As Long, j As Long, k As Long
Dim x As Long, y As Long
Dim bFilled() As Byte
Dim TheStack() As Long
Dim nValue As Long, nTemp As Long
'init array
CopyMemory ByVal lpbm, ByVal lpbmIn, w * h * 4&
With tSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
With tSASrc
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbmIn2
End With
CopyMemory ByVal VarPtrArray(bDibSrc()), VarPtr(tSASrc), 4&
ReDim bFilled(w - 1, h - 1)
ReDim TheStack(w * h - 1)
'get data
nValue = bProps(0)
nValue = nValue + nValue + nValue
'calc
For j = 0 To h - 1
 For i = 0 To w - 1
  If bFilled(i, j) = 0 Then
   bFilled(i, j) = 1
   With bDib(i, j)
    If CLng(.rgbBlue) + .rgbGreen + .rgbRed >= nValue Then
     clr = bDibSrc(i, j)
     k = 0
     TheStack(0) = i + j * &H10000
     Do Until k < 0
      x = TheStack(k)
      y = x \ &H10000
      x = x And &HFFFF&
      With bDib(x, y)
       If CLng(.rgbBlue) + .rgbGreen + .rgbRed >= nValue Then
        bDib(x, y) = clr
        nTemp = (y - 1) And (h - 1)
        If bFilled(x, nTemp) = 0 Then
         bFilled(x, nTemp) = 1
         TheStack(k) = x + nTemp * &H10000
         k = k + 1
        End If
        nTemp = (y + 1) And (h - 1)
        If bFilled(x, nTemp) = 0 Then
         bFilled(x, nTemp) = 1
         TheStack(k) = x + nTemp * &H10000
         k = k + 1
        End If
        nTemp = (x - 1) And (w - 1)
        If bFilled(nTemp, y) = 0 Then
         bFilled(nTemp, y) = 1
         TheStack(k) = nTemp + y * &H10000
         k = k + 1
        End If
        nTemp = (x + 1) And (w - 1)
        If bFilled(nTemp, y) = 0 Then
         bFilled(nTemp, y) = 1
         TheStack(k) = nTemp + y * &H10000
         k = k + 1
        End If
       Else
        .rgbBlue = 0
        .rgbGreen = 0
        .rgbRed = 0
        .rgbReserved = 255
       End If
      End With
      k = k - 1
     Loop
    Else
     .rgbBlue = 0
     .rgbGreen = 0
     .rgbRed = 0
     .rgbReserved = 255
    End If
   End With
  End If
 Next i
Next j
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
ZeroMemory ByVal VarPtrArray(bDibSrc()), 4&
End Sub

Private Sub pCalcDialect(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim i As Long, j As Long, k As Long
Dim x As Long, y As Long
Dim xx As Long, yy As Long
Dim clr As RGBQUAD
Dim nIndex As Long, nCount As Long
Dim nMode As Long '0=max 1=min 2=mid
Dim nValue As Long '0-1024
'///max
Dim c0 As RGBQUAD, m0 As Long 'first
Dim c1 As RGBQUAD, m1 As Long 'last
Dim c2 As RGBQUAD, m2 As Long 'this
'///
'init array
CopyMemory ByVal lpbm, ByVal lpbmIn, 4& * w * h
With tSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
'calc
nCount = bProps(0)
nValue = (bProps(1) * 1024& + 128&) \ 255
nMode = bProps(2) And 3&
If nValue = 0 Then nCount = 0
For nIndex = 1 To nCount
 'horizontal
 For j = 0 To h - 1
  c0 = bDib(0, j)
  m0 = CLng(c0.rgbBlue) + c0.rgbGreen + c0.rgbRed
  c1 = bDib(w - 1, j)
  m1 = CLng(c1.rgbBlue) + c1.rgbGreen + c1.rgbRed
  c2 = c0
  m2 = m0
  Select Case nMode
  Case 0 'max
   For i = 1 To w - 1
    With bDib(i, j)
     k = CLng(.rgbBlue) + .rgbGreen + .rgbRed
    End With
    If m2 > k Then
     If m1 > m2 Then pMixColor1024 bDib(i - 1, j), c1, nValue
     c1 = c2
     c2 = bDib(i, j)
    Else
     If m1 > k Then
      pMixColor1024 bDib(i - 1, j), c1, nValue
      c1 = c2
      c2 = bDib(i, j)
     Else
      c1 = c2
      c2 = bDib(i, j)
      pMixColor1024 bDib(i - 1, j), c2, nValue
     End If
    End If
    m1 = m2
    m2 = k
   Next i
   'last
   If m2 > m0 Then
    If m1 > m2 Then pMixColor1024 bDib(w - 1, j), c1, nValue
   Else
    If m1 > m0 Then pMixColor1024 bDib(w - 1, j), c1, nValue Else pMixColor1024 bDib(w - 1, j), c0, nValue
   End If
  Case 1 'min
   For i = 1 To w - 1
    With bDib(i, j)
     k = CLng(.rgbBlue) + .rgbGreen + .rgbRed
    End With
    If m2 < k Then
     If m1 < m2 Then pMixColor1024 bDib(i - 1, j), c1, nValue
     c1 = c2
     c2 = bDib(i, j)
    Else
     If m1 < k Then
      pMixColor1024 bDib(i - 1, j), c1, nValue
      c1 = c2
      c2 = bDib(i, j)
     Else
      c1 = c2
      c2 = bDib(i, j)
      pMixColor1024 bDib(i - 1, j), c2, nValue
     End If
    End If
    m1 = m2
    m2 = k
   Next i
   'last
   If m2 < m0 Then
    If m1 < m2 Then pMixColor1024 bDib(w - 1, j), c1, nValue
   Else
    If m1 < m0 Then pMixColor1024 bDib(w - 1, j), c1, nValue Else pMixColor1024 bDib(w - 1, j), c0, nValue
   End If
  Case 2 'middle??
   For i = 1 To w - 1
    With bDib(i, j)
     k = CLng(.rgbBlue) + .rgbGreen + .rgbRed
    End With
    If m2 > k Then
     If m2 > m1 Then
      If m1 > k Then
       pMixColor1024 bDib(i - 1, j), c1, nValue
       c1 = c2
       c2 = bDib(i, j)
      Else
       c1 = c2
       c2 = bDib(i, j)
       pMixColor1024 bDib(i - 1, j), c2, nValue
      End If
     Else
      c1 = c2
      c2 = bDib(i, j)
     End If
    Else
     If m1 > k Then
      c1 = c2
      c2 = bDib(i, j)
      pMixColor1024 bDib(i - 1, j), c2, nValue
     Else
      If m1 > m2 Then pMixColor1024 bDib(i - 1, j), c1, nValue
      c1 = c2
      c2 = bDib(i, j)
     End If
    End If
    m1 = m2
    m2 = k
   Next i
   'last
   If m2 > m0 Then
    If m2 > m1 Then
     If m1 > m0 Then pMixColor1024 bDib(w - 1, j), c1, nValue Else pMixColor1024 bDib(w - 1, j), c0, nValue
    End If
   Else
    If m1 > m0 Then pMixColor1024 bDib(w - 1, j), c0, nValue Else _
    If m1 > m2 Then pMixColor1024 bDib(w - 1, j), c1, nValue
   End If
  End Select
 Next j
 'vertical
 For i = 0 To w - 1
  c0 = bDib(i, 0)
  m0 = CLng(c0.rgbBlue) + c0.rgbGreen + c0.rgbRed
  c1 = bDib(i, h - 1)
  m1 = CLng(c1.rgbBlue) + c1.rgbGreen + c1.rgbRed
  c2 = c0
  m2 = m0
  Select Case nMode
  Case 0 'max
   For j = 1 To h - 1
    With bDib(i, j)
     k = CLng(.rgbBlue) + .rgbGreen + .rgbRed
    End With
    If m2 > k Then
     If m1 > m2 Then pMixColor1024 bDib(i, j - 1), c1, nValue
     c1 = c2
     c2 = bDib(i, j)
    Else
     If m1 > k Then
      pMixColor1024 bDib(i, j - 1), c1, nValue
      c1 = c2
      c2 = bDib(i, j)
     Else
      c1 = c2
      c2 = bDib(i, j)
      pMixColor1024 bDib(i, j - 1), c2, nValue
     End If
    End If
    m1 = m2
    m2 = k
   Next j
   'last
   If m2 > m0 Then
    If m1 > m2 Then pMixColor1024 bDib(i, h - 1), c1, nValue
   Else
    If m1 > m0 Then pMixColor1024 bDib(i, h - 1), c1, nValue Else pMixColor1024 bDib(i, h - 1), c0, nValue
   End If
  Case 1 'min
   For j = 1 To h - 1
    With bDib(i, j)
     k = CLng(.rgbBlue) + .rgbGreen + .rgbRed
    End With
    If m2 < k Then
     If m1 < m2 Then pMixColor1024 bDib(i, j - 1), c1, nValue
     c1 = c2
     c2 = bDib(i, j)
    Else
     If m1 < k Then
      pMixColor1024 bDib(i, j - 1), c1, nValue
      c1 = c2
      c2 = bDib(i, j)
     Else
      c1 = c2
      c2 = bDib(i, j)
      pMixColor1024 bDib(i, j - 1), c2, nValue
     End If
    End If
    m1 = m2
    m2 = k
   Next j
   'last
   If m2 < m0 Then
    If m1 < m2 Then pMixColor1024 bDib(i, h - 1), c1, nValue
   Else
    If m1 < m0 Then pMixColor1024 bDib(i, h - 1), c1, nValue Else pMixColor1024 bDib(i, h - 1), c0, nValue
   End If
  Case 2 'middle??
   For j = 1 To h - 1
    With bDib(i, j)
     k = CLng(.rgbBlue) + .rgbGreen + .rgbRed
    End With
    If m2 > k Then
     If m2 > m1 Then
      If m1 > k Then
       pMixColor1024 bDib(i, j - 1), c1, nValue
       c1 = c2
       c2 = bDib(i, j)
      Else
       c1 = c2
       c2 = bDib(i, j)
       pMixColor1024 bDib(i, j - 1), c2, nValue
      End If
     Else
      c1 = c2
      c2 = bDib(i, j)
     End If
    Else
     If m1 > k Then
      c1 = c2
      c2 = bDib(i, j)
      pMixColor1024 bDib(i, j - 1), c2, nValue
     Else
      If m1 > m2 Then pMixColor1024 bDib(i, j - 1), c1, nValue
      c1 = c2
      c2 = bDib(i, j)
     End If
    End If
    m1 = m2
    m2 = k
   Next j
   'last
   If m2 > m0 Then
    If m2 > m1 Then
     If m1 > m0 Then pMixColor1024 bDib(i, h - 1), c1, nValue Else pMixColor1024 bDib(i, h - 1), c0, nValue
    End If
   Else
    If m1 > m0 Then pMixColor1024 bDib(i, h - 1), c0, nValue Else _
    If m1 > m2 Then pMixColor1024 bDib(i, h - 1), c1, nValue
   End If
  End Select
 Next i
Next nIndex
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcTwirl(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim bDibSrc() As RGBQUAD
Dim tSASrc As SAFEARRAY2D
Dim i As Long, j As Long, k As Long
Dim ii As Long, jj As Long 'twirl x,y = -16777216 - 16777216
Dim x As Long, y As Long
Dim xxx As Long, yyy As Long
Dim nXClampMode As Long, nYClampMode As Long
Dim f(5) As Single, fAngle As Single
Dim nLeft As Long, nTop As Long, nRight As Long, nBottom As Long
Dim xx As Long, yy As Long, ww As Long, hh As Long
'table
Dim TheSin(4095) As Long, TheCos(4095) As Long '1 -> 4096
'SQRT table
Dim TheTable(2049) As Long
'init array
CopyMemory ByVal lpbm, ByVal lpbmIn, w * h * 4&
With tSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
With tSASrc
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbmIn
End With
CopyMemory ByVal VarPtrArray(bDibSrc()), VarPtr(tSASrc), 4&
'get data
nXClampMode = bProps(24)
nYClampMode = (nXClampMode And 12&) \ 4&
nXClampMode = nXClampMode And 3&
CopyMemory f(0), bProps(0), 24&
'calc table
For i = 0 To 4095
 fAngle = f(0) * 二π * ((4095& - i) / 4096#) ^ f(1) '??????
 TheSin(i) = Sin(fAngle) * 4096#
 TheCos(i) = Cos(fAngle) * 4096#
Next i
'sqrt
For i = 1 To 2049
 TheTable(i) = Sqr(i * 16384&)
Next i
'center,radius
ww = f(2) * w
hh = f(3) * h
xx = f(4) * w
yy = f(5) * h
nLeft = xx - ww
nRight = xx + ww
nTop = yy - hh
nBottom = yy + hh
If ww > 0 Then ww = 16777216 \ ww
If hh > 0 Then hh = 16777216 \ hh
'start calc
y = -16777216
For j = nTop To nBottom
 If (j >= 0 And j < h) Or nYClampMode Then '<>0
  x = -16777216
  For i = nLeft To nRight
   If (i >= 0 And i < w) Or nXClampMode Then '<>0
    'calc SQRT
    ii = x \ 4096& '-4096 - 4096
    jj = y \ 4096&
    ii = ii * ii + jj * jj '0 - 33554432
    If ii < 2050 Then
     jj = (TheTable(ii) + 64&) \ 128&
    Else
     jj = ii \ 16384&
     k = TheTable(jj)
     jj = k + ((TheTable(jj + 1) - k) * (ii And 16383&)) \ 16384&
     jj = (jj + ii \ jj) \ 2
    End If
    'calc new pos
    If jj < 4096 Then
     k = i - xx
     yyy = j - yy
     xxx = (k * TheCos(jj) + yyy * TheSin(jj)) \ 16&
     yyy = (yyy * TheCos(jj) - k * TheSin(jj)) \ 16&
     ii = xx + (xxx And &HFFFFFF00) \ 256&
     jj = yy + (yyy And &HFFFFFF00) \ 256&
     pGetColorEx bDibSrc, w, h, ii, jj, xxx And 255&, yyy And 255&, nXClampMode, nYClampMode, bDib(i And (w - 1), j And (h - 1))
    End If
   End If
   x = x + ww
  Next i
 End If
 y = y + hh
Next j
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
ZeroMemory ByVal VarPtrArray(bDibSrc()), 4&
End Sub

'bug!!! look strange
Private Sub pCalcBulge(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim bDibSrc() As RGBQUAD
Dim tSASrc As SAFEARRAY2D
Dim i As Long, j As Long, k As Long
Dim ii As Long, jj As Long 'twirl x,y = -16777216 - 16777216
Dim x As Long, y As Long
Dim xxx As Long, yyy As Long
Dim nXClampMode As Long, nYClampMode As Long
Dim f(5) As Single, fAngle As Single
Dim nLeft As Long, nTop As Long, nRight As Long, nBottom As Long
Dim xx As Long, yy As Long, ww As Long, hh As Long
Dim nValue As Long
'SQRT table
Dim TheTable(2049) As Long
'init array
CopyMemory ByVal lpbm, ByVal lpbmIn, w * h * 4&
With tSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
With tSASrc
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbmIn
End With
CopyMemory ByVal VarPtrArray(bDibSrc()), VarPtr(tSASrc), 4&
'get data
nXClampMode = bProps(20)
nYClampMode = (nXClampMode And 12&) \ 4&
nXClampMode = nXClampMode And 3&
CopyMemory f(0), bProps(0), 20&
nValue = f(0) * 256#
'sqrt
For i = 1 To 2049
 TheTable(i) = Sqr(i * 16384&)
Next i
'center,radius
ww = f(1) * w
hh = f(2) * h
xx = f(3) * w
yy = f(4) * h
nLeft = xx - ww
nRight = xx + ww
nTop = yy - hh
nBottom = yy + hh
If ww > 0 Then ww = 16777216 \ ww
If hh > 0 Then hh = 16777216 \ hh
'start calc
y = -16777216
For j = nTop To nBottom
 If (j >= 0 And j < h) Or nYClampMode Then '<>0
  x = -16777216
  For i = nLeft To nRight
   If (i >= 0 And i < w) Or nXClampMode Then '<>0
    'calc SQRT
    ii = x \ 4096& '-4096 - 4096
    jj = y \ 4096&
    ii = ii * ii + jj * jj '0 - 33554432
    If ii < 2050 Then
     jj = (TheTable(ii) + 64&) \ 128&
    Else
     jj = ii \ 16384&
     k = TheTable(jj)
     jj = k + ((TheTable(jj + 1) - k) * (ii And 16383&)) \ 16384&
     jj = (jj + ii \ jj) \ 2
    End If
    'calc new pos
    If jj < 4096 Then
     jj = 4095& - jj
     If jj > 512 Then
      jj = 4096 - (((jj * jj) \ 64&) * nValue) \ 4096&
     Else
      jj = (1073741824 - nValue * jj * jj) \ 262144 '1 -> 4096
     End If
     xxx = ((i - xx) * jj) \ 16&
     yyy = ((j - yy) * jj) \ 16&
     ii = xx + (xxx And &HFFFFFF00) \ 256&
     jj = yy + (yyy And &HFFFFFF00) \ 256&
     pGetColorEx bDibSrc, w, h, ii, jj, xxx And 255&, yyy And 255&, nXClampMode, nYClampMode, bDib(i And (w - 1), j And (h - 1))
    End If
   End If
   x = x + ww
  Next i
 End If
 y = y + hh
Next j
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
ZeroMemory ByVal VarPtrArray(bDibSrc()), 4&
End Sub

Private Sub pCalcUnwrap(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim bDibSrc() As RGBQUAD
Dim tSASrc As SAFEARRAY2D
Dim i As Long, j As Long, k As Long
Dim ii As Long, jj As Long
Dim x As Long, y As Long
Dim xx As Long, yy As Long
Dim xxx As Long, yyy As Long
Dim nXClampMode As Long, nYClampMode As Long, nMode As Long
Dim fAngle As Single
'table
Dim TheSin(4095) As Long, TheCos(4095) As Long '1 -> 4096
'init array
With tSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
With tSASrc
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbmIn
End With
CopyMemory ByVal VarPtrArray(bDibSrc()), VarPtr(tSASrc), 4&
'get data
nXClampMode = bProps(0)
nYClampMode = (nXClampMode And 12&) \ 4&
nMode = (nXClampMode And &HF0&) \ 16&
nXClampMode = nXClampMode And 3&
''sqrt
'For i = 1 To 2049
' TheTable(i) = Sqr(i * 16384&)
'Next i
Select Case nMode
Case 0 'polar->normal
 'calc table
 For i = 0 To 4095
  fAngle = 二π * (i / 4096#)
  TheSin(i) = Sin(fAngle) * 4096#
  TheCos(i) = Cos(fAngle) * 4096#
 Next i
 k = 4096& \ w
 xx = w * 128&
 yy = h * 128&
 'calc
 For j = 0 To h - 1
  x = 0
  For i = 0 To w - 1
   'calc new pos
   xxx = xx + (j * TheCos(x)) \ 32&
   yyy = yy + (-j * TheSin(x)) \ 32&
   ii = (xxx And &HFFFFFF00) \ 256&
   jj = (yyy And &HFFFFFF00) \ 256&
   pGetColorEx bDibSrc, w, h, ii, jj, xxx And 255&, yyy And 255&, nXClampMode, nYClampMode, bDib(i, j)
   'get next
   x = x + k
  Next i
 Next j
Case 1 'normal->polar
 'init ATAN2
 Const TheConst As Double = π / 32768
 For i = 0 To 4095
  TheSin(i) = Atn(i / (4096# - i)) / TheConst '2π -> 65536
 Next i
 'calc
 y = h - 1 'y*2
 yy = y * y
 For j = 0 To h - 1
  x = 1 - w 'x*2
  xx = x * x
  k = Sqr(xx + yy) * 8
  For i = 0 To w - 1
   'calc ATAN2
   If x > 0 Then
    If y = 0 Then
     xxx = 0
    ElseIf y > 0 Then 'I
     xxx = TheSin((y * 4096&) \ (x + y))
    Else 'IV
     xxx = 65536 - TheSin((y * 4096&) \ (y - x))
    End If
   ElseIf x < 0 Then
    If y = 0 Then
     xxx = 32768
    ElseIf y > 0 Then 'II
     xxx = 32768 - TheSin((y * 4096&) \ (y - x))
    Else 'III
     xxx = 32768 + TheSin((y * 4096&) \ (x + y))
    End If
   Else
    If y < 0 Then xxx = 49152 Else xxx = 16384&
   End If
   'calc SQRT
   k = (k + ((xx + yy) * 64&) \ k) \ 2&
   'calc new pos
   xxx = (xxx * w) \ 256&
   yyy = k * 32&
   ii = (xxx And &HFFFFFF00) \ 256&
   jj = (yyy And &HFFFFFF00) \ 256&
   pGetColorEx bDibSrc, w, h, ii, jj, xxx And 255&, yyy And 255&, nXClampMode, nYClampMode, bDib(i, j)
   'get next
   xx = xx + (x + 1) * 4&
   x = x + 2
  Next i
  yy = yy - (y - 1) * 4&
  y = y - 2
 Next j
End Select
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
ZeroMemory ByVal VarPtrArray(bDibSrc()), 4&
End Sub

Private Sub pCalcAbnormals(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal lpbmIn2 As Long, ByVal m As Long, bProps() As Byte)
Const TheConst As Single = π / 512#
Const TheConst2 As Single = π / 255#
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim bDibMap() As RGBQUAD
Dim tSAMap As SAFEARRAY2D
Dim i As Long, k As Long
Dim nMode As Long, nMode2 As Long
'quat
Dim f As Single 'sensitivity
Dim r As typeQuat, r0 As typeQuat
Dim v As typeQuat
Dim t1 As typeQuat, t2 As typeQuat
'init array
CopyMemory ByVal lpbm, ByVal lpbmIn, m * 4&
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = m
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
If lpbmIn2 <> 0 Then
 With tSAMap
  .cbElements = 4
  .cDims = 1
  .Bounds(0).cElements = m
  .pvData = lpbmIn2
 End With
 CopyMemory ByVal VarPtrArray(bDibMap()), VarPtr(tSAMap), 4&
End If
'get data
nMode = bProps(17)
nMode2 = (nMode And 12) \ 4&
nMode = nMode And 3&
f = bProps(16) / 255#
'    quat rotation(2 * nv_pi * w, vec3(x, y, z));
'    quat rotation0 = rotation;
CopyMemory r, bProps(0), 16&
pQuatFromRotation r, r.w * -π, r.x, r.y, r.z
r0 = r
'calc
For i = 0 To m - 1
 With bDib(i)
  v.w = 0
  v.x = (.rgbBlue - 128&) / 128#
  v.y = (.rgbGreen - 128&) / 128#
  v.z = (.rgbRed - 128&) / 128#
 End With
 If lpbmIn2 <> 0 Then
  With bDibMap(i)
   If nMode < 2 Then
    If nMode = 0 Then 'normals
     v.x = v.x + ((.rgbBlue - 128&) / 128# - v.x) * f
     v.y = v.y + ((.rgbGreen - 128&) / 128# - v.y) * f
     v.z = v.z + ((.rgbRed - 128&) / 128# - v.z) * f
    Else 'height
     k = (.rgbBlue * 146& + .rgbGreen * 1454& + .rgbRed * 456& + 512&) \ 1024& '0-512
     pQuatFromRotation t1, (TheConst * f) * k, 0, 0, 1
     pQuatMul3 r, t1, r0
    End If
   Else 'quaternions
    pQuatFromRotation t1, (TheConst2 * f) * .rgbReserved, _
    (.rgbBlue - 128&) / 128#, (.rgbGreen - 128&) / 128#, (.rgbRed - 128&) / 128#
     pQuatMul3 r, t1, r0
   End If
  End With
 End If
 'v = rotation * v * rotation.Inverse();
 pQuatMul3 t2, r, v
 With v
  .w = r.w
  .x = -r.x
  .y = -r.y
  .z = -r.z
 End With
 pQuatMul3 t1, t2, v
 'v.Normalize();
 pQuatNormalize3 t1
 With bDib(i)
  k = CLng(t1.x * 128#) + 128&
  If k < 0 Then k = 0 Else If k > 255 Then k = 255
  'mirroring - for broken normal maps
  If nMode2 And 1& Then k = 255 - k
  .rgbBlue = k
  k = CLng(t1.y * 128#) + 128&
  If k < 0 Then k = 0 Else If k > 255 Then k = 255
  'mirroring - for broken normal maps
  If nMode2 And 2& Then k = 255 - k
  .rgbGreen = k
  k = CLng(t1.z * 128#) + 128&
  If k < 0 Then k = 0 Else If k > 255 Then k = 255
  .rgbRed = k
 End With
Next i
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
If lpbmIn2 <> 0 Then ZeroMemory ByVal VarPtrArray(bDibMap()), 4&
End Sub

'just anotner blur :-3
Private Sub pCalcSharpen(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Dim bDib() As RGBQUAD, clr As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim i As Long, j As Long, k As Long
Dim nCount As Long
Dim nXSize As Long, nYSize As Long
Dim nXMax As Long, nYMax As Long 'sum / max
Dim nXClampMode As Long, nYClampMode As Long
Dim f As Single
Dim nAmplify As Long
'////
Dim nSumRed() As Long, nSumGreen() As Long, nSumBlue() As Long, nSumReserved() As Long
Dim bDibTemp() As RGBQUAD
Dim nRed As Long, nGreen As Long, nBlue As Long, nReserved As Long
'////
CopyMemory ByVal lpbm, ByVal lpbmIn, 4& * w * h
'get data
nCount = bProps(0)
nXClampMode = (nCount And 24&) \ 8&
nYClampMode = (nCount And 96&) \ 32&
nCount = nCount And 7&
'get size
CopyMemory f, bProps(1), 4&
nXSize = f * w
CopyMemory f, bProps(5), 4&
nYSize = f * h
'////
If (nXSize = 0 And nYSize = 0) Or nCount = 0 Then Exit Sub
'init array
With tSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
If w > h Then i = w - 1 Else i = h - 1
ReDim nSumRed(i), nSumGreen(i), nSumBlue(i), nSumReserved(i)
If nXSize > 0 And nYSize > 0 Then ReDim bDibTemp(w - 1, h - 1)
CopyMemory f, bProps(9), 4&
nAmplify = f * 1024#
nXMax = nXSize + nXSize + 1
nYMax = nYSize + nYSize + 1
'start calc
Do While nCount > 0
 nCount = nCount - 1
 'X
 If nXSize > 0 Then
  For j = 0 To h - 1
   'calc first sum
   nRed = 0
   nGreen = 0
   nBlue = 0
   nReserved = 0
   For i = -nXSize To nXSize
    pGetColor bDib, w, h, i, j, nXClampMode, clr
    With clr
     nRed = nRed + .rgbRed
     nGreen = nGreen + .rgbGreen
     nBlue = nBlue + .rgbBlue
     nReserved = nReserved + .rgbReserved
    End With
   Next i
   nSumRed(0) = nRed
   nSumGreen(0) = nGreen
   nSumBlue(0) = nBlue
   nSumReserved(0) = nReserved
   'calc sum
   For i = 1 To w - 1
    pGetColor bDib, w, h, i + nXSize, j, nXClampMode, clr
    With clr
     nRed = nRed + .rgbRed
     nGreen = nGreen + .rgbGreen
     nBlue = nBlue + .rgbBlue
     nReserved = nReserved + .rgbReserved
    End With
    pGetColor bDib, w, h, i - nXSize - 1, j, nXClampMode, clr
    With clr
     nRed = nRed - .rgbRed
     nGreen = nGreen - .rgbGreen
     nBlue = nBlue - .rgbBlue
     nReserved = nReserved - .rgbReserved
    End With
    nSumRed(i) = nRed
    nSumGreen(i) = nGreen
    nSumBlue(i) = nBlue
    nSumReserved(i) = nReserved
   Next i
   '???????? bug!!! backup original
   If nYSize > 0 Then
    For i = 0 To w - 1
     With bDibTemp(i, j)
      .rgbRed = nSumRed(i) \ nXMax
      .rgbGreen = nSumGreen(i) \ nXMax
      .rgbBlue = nSumBlue(i) \ nXMax
      .rgbReserved = nSumReserved(i) \ nXMax
     End With
    Next i
   Else
    For i = 0 To w - 1
     With bDib(i, j)
      k = nSumBlue(i) \ nXMax
      k = .rgbBlue + ((.rgbBlue - k) * nAmplify) \ 1024&
      If k < 0 Then k = 0 Else If k > 255 Then k = 255
      .rgbBlue = k
      k = nSumGreen(i) \ nXMax
      k = .rgbGreen + ((.rgbGreen - k) * nAmplify) \ 1024&
      If k < 0 Then k = 0 Else If k > 255 Then k = 255
      .rgbGreen = k
      k = nSumRed(i) \ nXMax
      k = .rgbRed + ((.rgbRed - k) * nAmplify) \ 1024&
      If k < 0 Then k = 0 Else If k > 255 Then k = 255
      .rgbRed = k
      k = nSumReserved(i) \ nXMax
      k = .rgbReserved + ((.rgbReserved - k) * nAmplify) \ 1024&
      If k < 0 Then k = 0 Else If k > 255 Then k = 255
      .rgbReserved = k
     End With
    Next i
   End If
  Next j
 End If
 'Y
 If nYSize > 0 Then
  For i = 0 To w - 1
   'calc first sum
   nRed = 0
   nGreen = 0
   nBlue = 0
   nReserved = 0
   If nXSize > 0 Then
    For j = -nYSize To nYSize
     pGetColor bDibTemp, w, h, i, j, nYClampMode, clr
     With clr
      nRed = nRed + .rgbRed
      nGreen = nGreen + .rgbGreen
      nBlue = nBlue + .rgbBlue
      nReserved = nReserved + .rgbReserved
     End With
    Next j
   Else
    For j = -nYSize To nYSize
     pGetColor bDib, w, h, i, j, nYClampMode, clr
     With clr
      nRed = nRed + .rgbRed
      nGreen = nGreen + .rgbGreen
      nBlue = nBlue + .rgbBlue
      nReserved = nReserved + .rgbReserved
     End With
    Next j
   End If
   nSumRed(0) = nRed
   nSumGreen(0) = nGreen
   nSumBlue(0) = nBlue
   nSumReserved(0) = nReserved
   'calc sum
   If nXSize > 0 Then
    For j = 1 To h - 1
     pGetColor bDibTemp, w, h, i, j + nYSize, nYClampMode, clr
     With clr
      nRed = nRed + .rgbRed
      nGreen = nGreen + .rgbGreen
      nBlue = nBlue + .rgbBlue
      nReserved = nReserved + .rgbReserved
     End With
     pGetColor bDibTemp, w, h, i, j - nYSize - 1, nYClampMode, clr
     With clr
      nRed = nRed - .rgbRed
      nGreen = nGreen - .rgbGreen
      nBlue = nBlue - .rgbBlue
      nReserved = nReserved - .rgbReserved
     End With
     nSumRed(j) = nRed
     nSumGreen(j) = nGreen
     nSumBlue(j) = nBlue
     nSumReserved(j) = nReserved
    Next j
   Else
    For j = 1 To h - 1
     pGetColor bDib, w, h, i, j + nYSize, nYClampMode, clr
     With clr
      nRed = nRed + .rgbRed
      nGreen = nGreen + .rgbGreen
      nBlue = nBlue + .rgbBlue
      nReserved = nReserved + .rgbReserved
     End With
     pGetColor bDib, w, h, i, j - nYSize - 1, nYClampMode, clr
     With clr
      nRed = nRed - .rgbRed
      nGreen = nGreen - .rgbGreen
      nBlue = nBlue - .rgbBlue
      nReserved = nReserved - .rgbReserved
     End With
     nSumRed(j) = nRed
     nSumGreen(j) = nGreen
     nSumBlue(j) = nBlue
     nSumReserved(j) = nReserved
    Next j
   End If
   '???????? bug!!!
   For j = 0 To h - 1
    With bDib(i, j)
     k = nSumBlue(j) \ nYMax
     k = .rgbBlue + ((.rgbBlue - k) * nAmplify) \ 1024&
     If k < 0 Then k = 0 Else If k > 255 Then k = 255
     .rgbBlue = k
     k = nSumGreen(j) \ nYMax
     k = .rgbGreen + ((.rgbGreen - k) * nAmplify) \ 1024&
     If k < 0 Then k = 0 Else If k > 255 Then k = 255
     .rgbGreen = k
     k = nSumRed(j) \ nYMax
     k = .rgbRed + ((.rgbRed - k) * nAmplify) \ 1024&
     If k < 0 Then k = 0 Else If k > 255 Then k = 255
     .rgbRed = k
     k = nSumReserved(j) \ nYMax
     k = .rgbReserved + ((.rgbReserved - k) * nAmplify) \ 1024&
     If k < 0 Then k = 0 Else If k > 255 Then k = 255
     .rgbReserved = k
    End With
   Next j
  Next i
 End If
Loop
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub

Private Sub pCalcSlowGrow(ByVal lpbm As Long, ByVal lpbmIn As Long, ByVal w As Long, ByVal h As Long, bProps() As Byte)
Const TheConst As Double = π / 512
Dim bDib() As RGBQUAD
Dim tSA As SAFEARRAY2D
Dim i As Long, j As Long, k As Long
Dim ii As Long, jj As Long, kk As Long
Dim x As Long, y As Long '1 -> 256
Dim nAngle As Long '2π -> 1024
Dim nClr As Long, clr As RGBQUAD
Dim nXMax As Long, nYMax As Long
Dim nIndex As Long, nCount As Long
Dim nVar As Long, nLength As Long
Dim nMode As Long, nXMode As Long, nYMode As Long
Dim bHQ As Boolean, bIsEllipse As Boolean
Dim clr1 As RGBQUAD, clr2 As RGBQUAD
Dim TheClrTable(255) As RGBQUAD
Dim TheSin(1023) As Long '1 -> 256
Dim TheFuncTable(511) As Long 'probability!!
Dim TheSeed As Long
Dim fX As Single, fY As Single
Dim f(3) As Single
'init array
CopyMemory ByVal lpbm, ByVal lpbmIn, 4& * w * h
With tSA
 .cbElements = 4
 .cDims = 2
 .Bounds(0).cElements = h
 .Bounds(1).cElements = w
 .pvData = lpbm
End With
CopyMemory ByVal VarPtrArray(bDib()), VarPtr(tSA), 4&
'get color
CopyMemory clr1, bProps(8), 4&
CopyMemory clr2, bProps(12), 4&
For i = 0 To 255
 pMixColor TheClrTable(i), clr1, clr2, i
Next i
'init table
For i = 0 To 1023
 TheSin(i) = 256# * Sin(TheConst * i) '272 'a little bigger than 1 ??
Next i
CopyMemory f(0), bProps(34), 8&
For i = 0 To 511
 fX = Sin(TheConst * i)
 TheFuncTable(i) = (f(0) + (f(1) - f(0)) * fX * fX) * 5482.75 '=4194304/765
Next i
'get data TODO:
nXMax = w * 256& - 1
nYMax = h * 256& - 1
pGetSeed bProps, 4, TheSeed
CopyMemory nCount, bProps(0), 4&
nCount = nCount And 2097151
nLength = bProps(2) \ 32& + bProps(3) * 8&
nVar = bProps(6)
nXMode = bProps(7)
nYMode = (nXMode And 12&) \ 4&
nMode = (nXMode And 96&) \ 32& 'mode
bHQ = nXMode And 128&
bIsEllipse = nXMode And 16&
nXMode = nXMode And 3&
CopyMemory f(0), bProps(16), 16&
f(0) = f(0) * 256# * w
f(1) = f(1) * 256# * h
f(2) = f(2) * 256# * w
f(3) = f(3) * 256# * h
'start calc
For nIndex = 1 To nCount
 'calc start area
 i = 0
 Do
  fX = cUnk.fRnd2Float(nIndex, &HBEE& + i, TheSeed)
  fY = cUnk.fRnd2Float(nIndex, &HF00D& + i, TheSeed)
  If Not bIsEllipse Then Exit Do
  i = i + 1
 Loop While fX * fX + fY * fY > 1
 x = CLng(f(0) + fX * f(2)) And nXMax
 y = CLng(f(1) + fY * f(3)) And nYMax
 nAngle = (bProps(32) * 4& + (bProps(33) * ((cUnk.fRnd2(nIndex, &HFADE&, TheSeed) And 2047&) - 1024&) + 256&) \ 512&) And 1023&
 For i = 1 To nLength
  'calc direction
  nAngle = (nAngle + (nVar * ((cUnk.fRnd2(nIndex, i, TheSeed) And 2047&) - 1024&) + 256&) \ 512&) And 1023&
  'calc next position TODO:clamp mode?
  j = x + TheSin((nAngle + 256&) And 1023&)
  k = y + TheSin(nAngle)
  kk = 0
  If j < 0 Or j > nXMax Then
   If nXMode And 2 Then
    If nXMode And 1 Then 'mirror
     j = (-j - 1) And nXMax
     nAngle = (256& - nAngle) And 1023&
    Else 'clamp
     kk = &H7FFFFFFF
    End If
   Else
    If nXMode And 1 Then j = j And nXMax Else Exit For
   End If
  End If
  If k < 0 Or k > nYMax Then
   If nYMode And 2 Then
    If nYMode And 1 Then 'mirror
     k = (-k - 1) And nYMax
     nAngle = (-nAngle) And 1023&
    Else 'clamp
     kk = &H7FFFFFFF
    End If
   Else
    If nYMode And 1 Then k = k And nYMax Else Exit For
   End If
  End If
  'check probability
  ii = j \ 256&
  jj = k \ 256&
  If kk = 0 Then
   With bDib(ii, jj)
    kk = (CLng(.rgbBlue) + .rgbGreen + .rgbRed) * TheFuncTable(nAngle And 511&)
   End With
  End If
  If (cUnk.fRnd2(nIndex, i, TheSeed) And &H3FFFFF) < kk Then
   If bHQ Then
    x = (x - 128&) And nXMax
    y = (y - 128&) And nYMax
   End If
   ii = x \ 256&
   jj = y \ 256&
   j = x And 255&
   k = y And 255&
   nClr = cUnk.fRnd2(nIndex, &HFADE0FF, TheSeed) And 255&
   If bHQ And (j > 0 Or k > 0) Then 'high quality
    clr = TheClrTable(nClr)
    j = j * k
    pMixColor65536 bDib(ii, jj), clr, nMode, 65536 + j - ((x And 255&) + k) * 256&
    If x And 255& Then
     pMixColor65536 bDib((ii + 1) And (w - 1), jj), clr, nMode, (x And 255&) * 256& - j
     If k > 0 Then
      pMixColor65536 bDib((ii + 1) And (w - 1), (jj + 1) And (h - 1)), clr, nMode, j
     End If
    End If
    If k > 0 Then
     pMixColor65536 bDib(ii, (jj + 1) And (h - 1)), clr, nMode, k * 256& - j
    End If
   Else
    Select Case nMode
    Case 0 'normal
     bDib(ii, jj) = TheClrTable(nClr)
    Case 1 'blend
     clr = TheClrTable(nClr)
     If clr.rgbReserved = 255 Then
      bDib(ii, jj) = clr
     Else
      pBlendColor bDib(ii, jj), clr
     End If
    Case 2 'mix
     clr = TheClrTable(nClr)
     If clr.rgbReserved > 0 Then
      pMixColor2 bDib(ii, jj), clr
     End If
    End Select
   End If
   Exit For
  Else
   x = j
   y = k
  End If
 Next i
Next nIndex
'destroy array
ZeroMemory ByVal VarPtrArray(bDib()), 4&
End Sub


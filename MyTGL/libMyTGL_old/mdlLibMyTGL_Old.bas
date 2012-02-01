Attribute VB_Name = "mdlLibMyTGL_Old"
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

Public Type typeOperatorExport
 nType As Long
 nInputCount As Long
 nInputIndex() As Long '0-based
 nDeleteCount As Long
 nDeleteIndex() As Long '1-based
 bPropsSize As Long
 bProps() As Byte '0-based
 sPropsSize As Long
 sProps() As String 'string? 0-based
End Type

Public Type typeProjectExport
 nOperatorCount As Long
 tBitmap() As cAlphaDibSection '1-based
 nExportCount As Long
 nExportIndex() As Long '1-based
 sExportName() As String
 '///internal
 nCurrentIndex As Long '-1=error
 tOperators() As typeOperatorExport
End Type

Public Function LibMyTGLLoadFile(p As typeProjectExport, ByVal FileName As String) As Boolean
Dim b() As Byte, m As Long
'///
On Error Resume Next
Err.Clear
If GetAttr(FileName) And vbDirectory Then Exit Function
If Err.Number Then Exit Function
'///
Open FileName For Binary Access Read As #1
m = LOF(1)
If Err.Number Then
 Close
 Exit Function
End If
On Error GoTo 0
'///
If m > 0 Then
 ReDim b(m - 1)
 Get #1, 1, b
End If
Close
'///
If m > 0 Then LibMyTGLLoadFile = LibMyTGLLoadFileFromMemory(p, VarPtr(b(0)), m)
End Function

Public Function LibMyTGLLoadFileFromMemory(p As typeProjectExport, ByVal lp As Long, ByVal nSize As Long) As Boolean
Dim b As Boolean
Dim lps As Long, m As Long, nSize2 As Long
Dim i As Long, j As Long, k As Long
'///
Dim d() As Long
Dim tSA As SAFEARRAY2D
'///
p.nOperatorCount = 0
p.nExportCount = 0
p.nCurrentIndex = 0
Erase p.tBitmap, p.nExportIndex, p.sExportName, p.tOperators
'///
If nSize < 12 Then Exit Function
'///
m = nSize \ 4
With tSA
 .cbElements = 4
 .cDims = 1
 .Bounds(0).cElements = m
 .pvData = lp
End With
CopyMemory ByVal VarPtrArray(d()), VarPtr(tSA), 4&
'///
If d(0) = &H4754794D And d(1) = &H3078454C Then
 p.nOperatorCount = d(2)
 If p.nOperatorCount > 0 Then
  ReDim p.tBitmap(1 To p.nOperatorCount)
  ReDim p.tOperators(1 To p.nOperatorCount)
 End If
 lps = 3
 b = True
 '///load operators
 For i = 1 To p.nOperatorCount
  If i > 1 Then
   j = 0
   k = 0
   Do While lps < m
    If d(lps) < 0 Then
     j = j + 1
     If j > k Then
      k = k + 256&
      ReDim Preserve p.tOperators(i - 1).nDeleteIndex(1 To k)
     End If
     p.tOperators(i - 1).nDeleteIndex(j) = d(lps) And &H7FFFFFFF
     lps = lps + 1
    Else
     Exit Do
    End If
   Loop
   p.tOperators(i - 1).nDeleteCount = j
  End If
  '///
  If lps + 2 > m Then
   b = False
   Exit For
  End If
  p.tOperators(i).nType = d(lps)
  k = d(lps + 1)
  p.tOperators(i).nInputCount = k
  lps = lps + 2
  '///
  If k > 0 Then
   If lps + k > m Then
    b = False
    Exit For
   End If
   ReDim p.tOperators(i).nInputIndex(k - 1)
   CopyMemory p.tOperators(i).nInputIndex(0), d(lps), k * 4&
   lps = lps + k
  End If
 Next i
 '///load exports
 If b And lps < m Then
  p.nExportCount = d(lps)
  lps = lps + 1
  '///
  If lps + p.nExportCount * 2& <= m Then
   If p.nExportCount > 0 Then
    ReDim p.nExportIndex(1 To p.nExportCount)
    ReDim p.sExportName(1 To p.nExportCount)
   End If
   For i = 1 To p.nExportCount
    p.nExportIndex(i) = d(lps)
    j = d(lps + 1)
    nSize2 = nSize2 + j
    p.sExportName(i) = LeftB(Space(j \ 2 + 1), j)
    lps = lps + 2
   Next i
   '///operator data size
   For i = 1 To p.nOperatorCount
    If lps + 2 > m Then
     b = False
     Exit For
    End If
    j = d(lps)
    nSize2 = nSize2 + j
    p.tOperators(i).bPropsSize = j
    If j > 0 Then ReDim p.tOperators(i).bProps(j - 1)
    j = d(lps + 1)
    p.tOperators(i).sPropsSize = j
    If j > 0 Then ReDim p.tOperators(i).sProps(j - 1)
    lps = lps + 2
    '///
    If lps + j > m Then
     b = False
     Exit For
    End If
    For k = 0 To j - 1
     j = d(lps)
     nSize2 = nSize2 + j
     p.tOperators(i).sProps(k) = LeftB(Space(j \ 2 + 1), j)
     lps = lps + 1
    Next k
   Next i
   '///
   lps = lps * 4&
   If b And lps + nSize2 <= nSize Then
    '///export name
    For i = 1 To p.nExportCount
     j = LenB(p.sExportName(i))
     If j > 0 Then
      CopyMemory ByVal StrPtr(p.sExportName(i)), ByVal lp + lps, j
      lps = lps + j
     End If
    Next i
    '///binary data
    For i = 1 To p.nOperatorCount
     j = p.tOperators(i).bPropsSize
     If j > 0 Then
      CopyMemory p.tOperators(i).bProps(0), ByVal lp + lps, j
      lps = lps + j
     End If
     For k = 0 To p.tOperators(i).sPropsSize - 1
      j = LenB(p.tOperators(i).sProps(k))
      If j > 0 Then
       CopyMemory ByVal StrPtr(p.tOperators(i).sProps(k)), ByVal lp + lps, j
       lps = lps + j
      End If
     Next k
    Next i
    '///over
    LibMyTGLLoadFileFromMemory = True
   End If
  End If
 End If
End If
'///
ZeroMemory ByVal VarPtrArray(d()), 4&
'///
End Function

'return value: 1=calc current operator OK, 0=all operators are done, -1=error
Public Function LibMyTGLCalc(p As typeProjectExport, ByVal bDeleteUnusedOperator As Boolean) As Long
Dim idx As Long, i As Long, j As Long, m As Long
Dim bmIn() As typeAlphaDibSectionDescriptor
'///
idx = p.nCurrentIndex
If idx < 0 Then
 LibMyTGLCalc = -1
 Exit Function
ElseIf idx >= p.nOperatorCount Then
 Exit Function
End If
'///
idx = idx + 1
m = p.tOperators(idx).nInputCount
If m > 0 Then
 ReDim bmIn(m - 1)
 For i = 0 To m - 1
  j = p.tOperators(idx).nInputIndex(i)
  bmIn(i).Width = p.tBitmap(j).Width
  bmIn(i).Height = p.tBitmap(j).Height
  bmIn(i).lpbm = p.tBitmap(j).DIBSectionBitsPtr
 Next i
End If
'///
If Not CalcOperator(p.tBitmap(idx), m, bmIn, p.tOperators(idx).nType, p.tOperators(idx).bProps, p.tOperators(idx).sProps) Then
 p.nCurrentIndex = -1
 LibMyTGLCalc = -1
 Exit Function
End If
'///
p.nCurrentIndex = idx
If idx < p.nOperatorCount Then
 LibMyTGLCalc = 1
 If bDeleteUnusedOperator Then
  For i = 1 To p.tOperators(idx).nDeleteCount
   j = p.tOperators(idx).nDeleteIndex(i)
   Erase p.tBitmap(j).b
   p.tBitmap(j).Width = 0
   p.tBitmap(j).Height = 0
   p.tBitmap(j).DIBSectionBitsPtr = 0
  Next i
 End If
'Else
' If bDeleteUnusedOperator Then
'  'TODO:
' End If
End If
End Function

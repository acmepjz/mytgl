Attribute VB_Name = "mdlFile"
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

'/////////////////////file format
'£ªfile
'©À£ªheader
'©¸£ªdata
'  ©À£ªpage
'  ©¦©À£ªoperators
'  ©¦©¸£ªcomments,etc.
'  ©¸£ªoperators
'    ©À£ªstring props (byte array)
'    ©¸£ªbyte props
'/////////////////////

Private Type typePrjFileHeader '24 bytes
 Signature0 As Long '&H4754794D ("MyTG")
 Signature1 As Integer '&H4C ("L")
 nPageCount As Integer
 nReserved As Long 'must be 0 --- if compressed it's &H414D5A4C ("LZMA")
 nDataSize As Long 'can be 0
 nDecompressedSize As Long 'can be 0 --- not compressed
 nOpCount As Long 'max
End Type

Public Enum enumPrjCompressMode
 prjUncompressed = 0
 prjCompressLZSS = 1
 prjCompressLZMA = 2
 prjCompressZLib = 3
End Enum

Public Function LoadPrjFile(p As typeProject, ByVal FileName As String) As Boolean
On Error GoTo a
Dim i As Long, j As Long, k As Long
Dim bErr As Boolean
Dim h As typePrjFileHeader
Dim b() As Byte, b2() As Byte, lp As Long
'Dim cObj As New clsOperators
Dim cLZSS As New clsLZSS
'objects
Dim objs() As Long, objc As Long, objm As Long
'check file
If Dir(FileName, vbHidden Or vbSystem) = "" Then Err.Raise 53
'load file
Open FileName For Binary As #1
Get #1, 1, h
If h.Signature0 = &H4754794D And h.Signature1 = &H4C Then
 With p
  .nOpCount = h.nOpCount
  .nPageCount = h.nPageCount
  If .nOpCount > 0 Then ReDim .Operators(1 To .nOpCount) Else Erase .Operators
  If .nPageCount > 0 Then ReDim .Pages(1 To .nPageCount) Else Erase .Pages
 End With
 If h.nDataSize > 0 Then
  ReDim b(h.nDataSize - 1)
  Get #1, 25, b
  Select Case h.nReserved
  Case &H53535A4C 'LZSS
   h.nDataSize = cLZSS.DecompressData(b, b2, h.nDecompressedSize)
   ReDim b(h.nDataSize - 1)
   CopyMemory b(0), b2(LBound(b2)), h.nDataSize
   Erase b2
  Case &H414D5A4C 'LZMA!!
   LZMADecompress_Simple b, b2, h.nDecompressedSize
   h.nDataSize = h.nDecompressedSize
   ReDim b(h.nDataSize - 1)
   CopyMemory b(0), b2(LBound(b2)), h.nDataSize
   Erase b2
  Case &H62696C7A 'zlib
   ZLibDecompressByteArray b, b2, h.nDecompressedSize
   h.nDataSize = ZLibValueDecompressedSize
   ReDim b(h.nDataSize - 1)
   CopyMemory b(0), b2(LBound(b2)), h.nDataSize
   Erase b2
  Case Else 'uncompressed
   'do nothing
  End Select
 End If
Else
 bErr = True
End If
Close
If bErr Then Exit Function
'get data
lp = 0
With p
 'pages
 For i = 1 To .nPageCount
  With .Pages(i)
   CopyMemory k, b(lp), 4
   lp = lp + 4
   If k > 0 Then
    ReDim b2(k - 1)
    CopyMemory b2(0), b(lp), k
    .Name = b2
    lp = lp + k
   End If
   'rows
   For j = 0 To int_Page_Height - 1
    With .Rows(j)
     CopyMemory .nOpCount, b(lp), 4
     lp = lp + 4
     If .nOpCount > 0 Then
      ReDim .idxOp(1 To .nOpCount)
      k = .nOpCount * 4&
      CopyMemory .idxOp(1), b(lp), k
      lp = lp + k
      '//////Add!!
      For k = 1 To .nOpCount
       With p.Operators(.idxOp(k))
        .nPage = i
        .Top = j
       End With
      Next k
      '//////
     End If
    End With
   Next j
   'comments
   CopyMemory .nCommentCount, b(lp), 4
   lp = lp + 4
   If .nCommentCount > 0 Then
    ReDim .Comments(1 To .nCommentCount)
    For j = 1 To .nCommentCount
     With .Comments(j)
      CopyMemory .Left, b(lp), 4
      lp = lp + 4
      CopyMemory .Top, b(lp), 4
      lp = lp + 4
      CopyMemory .Width, b(lp), 4
      lp = lp + 4
      CopyMemory .Height, b(lp), 4
      lp = lp + 4
      CopyMemory .Color, b(lp), 4
      lp = lp + 4
      CopyMemory k, b(lp), 4
      lp = lp + 4
      If k > 0 Then
       ReDim b2(k - 1)
       CopyMemory b2(0), b(lp), k
       .Name = b2
       lp = lp + k
      End If
      CopyMemory k, b(lp), 4
      lp = lp + 4
      If k > 0 Then
       ReDim b2(k - 1)
       CopyMemory b2(0), b(lp), k
       .Value = b2
       lp = lp + k
      End If
     End With
    Next j
   End If
  End With
 Next i
 'operators
 For i = 1 To .nOpCount
  With .Operators(i)
   CopyMemory .nType, b(lp), 4
   lp = lp + 4
   If .nType > 0 Then
    .Flags = int_OpFlags_Error 'error=defaule ?? :-3
    CopyMemory k, b(lp), 4
    lp = lp + 4
    If k > 0 Then
     ReDim b2(k - 1)
     CopyMemory b2(0), b(lp), k
     .Name = b2
     lp = lp + k
    End If
'    CopyMemory .nPage, b(lp), 4
'    lp = lp + 4
    CopyMemory .Left, b(lp), 4
    lp = lp + 4
'    CopyMemory .Top, b(lp), 4
'    lp = lp + 4
    CopyMemory .Width, b(lp), 4
    lp = lp + 4
    CopyMemory k, b(lp), 4
    lp = lp + 4
    Debug.Assert k = tDef(.nType).StringCount 'check
    If k > 0 Then
     ReDim .sProps(k - 1)
     For j = 0 To k - 1
      CopyMemory k, b(lp), 4
      lp = lp + 4
      If k > 0 Then
       ReDim b2(k - 1)
       CopyMemory b2(0), b(lp), k
       .sProps(j) = b2
       lp = lp + k
      End If
     Next j
    End If
    CopyMemory k, b(lp), 4
    lp = lp + 4
    j = tDef(.nType).PropSize
    Debug.Assert k <= j 'check ??
    If k > 0 Then
     If k < j Then
      ReDim .bProps(j - 1)
     Else
      ReDim .bProps(k - 1)
     End If
     CopyMemory .bProps(0), b(lp), k
     lp = lp + k
    ElseIf j > 0 Then
     ReDim .bProps(j - 1)
    End If
   Else
    .Flags = -1 'unused!
   End If
  End With
 Next i
End With
'validate operator
objm = 256&
ReDim objs(1 To objm)
With p
 'find leaf
 For i = 1 To .nOpCount
  With .Operators(i)
   If .Flags >= 0 And .nType > 0 And .nType <= int_Generator_Max Then
    objc = objc + 1
    If objc > objm Then
     objm = objm + 256&
     ReDim Preserve objs(1 To objm)
    End If
    objs(objc) = i
   End If
  End With
 Next i
 'find load
 For i = 1 To .nOpCount
  With .Operators(i)
   If .Flags >= 0 And .nType > 0 And .nType = int_OpType_Load Then
    objc = objc + 1
    If objc > objm Then
     objm = objm + 256&
     ReDim Preserve objs(1 To objm)
    End If
    objs(objc) = i
   End If
  End With
 Next i
End With
'///disabled. TODO:
'For i = 1 To objc
' cObj.ValidateOps p, objs(i)
'Next i
'over
LoadPrjFile = True
Exit Function
a:
Close
End Function

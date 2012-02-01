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

Public Function SavePrjFile(p As typeProject, ByVal FileName As String, Optional ByVal nMode As enumPrjCompressMode) As Boolean
On Error GoTo a
Dim i As Long, j As Long, k As Long
Dim h As typePrjFileHeader
Dim b() As Byte, lp As Long
Dim b2() As Byte
Dim cLZSS As New clsLZSS
'calc size
With p
 'pages
 lp = .nPageCount * 8& 'string size,comment count
 For i = 1 To .nPageCount
  With .Pages(i)
   lp = lp + LenB(.Name)  'string
   'rows
   For j = 0 To int_Page_Height - 1
    lp = lp + (.Rows(j).nOpCount + 1) * 4& 'page+top
   Next j
   'new rows
   For j = 0 To int_Page_Height - 1
    If .Rows(j).nOpCount > 0 Then
     lp = lp + (.Rows(j).nOpCount + 1) * 4&
    End If
   Next j
   lp = lp + 4&
   'comments
   For j = 1 To .nCommentCount
    With .Comments(j)
     lp = lp + 28& 'left,top,width,height,color,stringX2
     lp = lp + LenB(.Name) + LenB(.Value)
    End With
   Next j
  End With
 Next i
 'operators
 lp = lp + .nOpCount * 4& 'type
 For i = 1 To .nOpCount
  With .Operators(i)
   If .Flags >= 0 Then
    lp = lp + LenB(.Name) 'string
    lp = lp + 20& 'string size,left,width,string count,byte data size
    k = tDef(.nType).StringCount
    lp = lp + k * 4& 'string prop size
    For j = 0 To k - 1
     lp = lp + LenB(.sProps(j)) 'string
    Next j
    lp = lp + tDef(.nType).PropSize
   End If
  End With
 Next i
End With
If lp > 0 Then
 h.nDecompressedSize = lp
 ReDim b(lp - 1)
 'put data
 lp = 0
 With p
  'pages
  For i = 1 To .nPageCount
   With .Pages(i)
    k = LenB(.Name)
    CopyMemory b(lp), k, 4
    lp = lp + 4
    If k > 0 Then
     CopyMemory b(lp), ByVal StrPtr(.Name), k
     lp = lp + k
    End If
    'rows
    For j = 0 To int_Page_Height - 1
     With .Rows(j)
      k = .nOpCount
      CopyMemory b(lp), k, 4
      lp = lp + 4
      If k > 0 Then
       k = k * 4&
       CopyMemory b(lp), .idxOp(1), k
       lp = lp + k
      End If
     End With
    Next j
    'comments,etc.
    CopyMemory b(lp), .nCommentCount, 4 'and indent,XX!
    lp = lp + 4
    For j = 1 To .nCommentCount
     With .Comments(j)
      CopyMemory b(lp), .Left, 4
      lp = lp + 4
      CopyMemory b(lp), .Top, 4
      lp = lp + 4
      CopyMemory b(lp), .Width, 4
      lp = lp + 4
      CopyMemory b(lp), .Height, 4
      lp = lp + 4
      CopyMemory b(lp), .Color, 4
      lp = lp + 4
      k = LenB(.Name)
      CopyMemory b(lp), k, 4
      lp = lp + 4
      If k > 0 Then
       CopyMemory b(lp), ByVal StrPtr(.Name), k
       lp = lp + k
      End If
      k = LenB(.Value)
      CopyMemory b(lp), k, 4
      lp = lp + 4
      If k > 0 Then
       CopyMemory b(lp), ByVal StrPtr(.Value), k
       lp = lp + k
      End If
     End With
    Next j
   End With
  Next i
  'operators
  For i = 1 To .nOpCount
   With .Operators(i)
    If .Flags >= 0 Then
     CopyMemory b(lp), .nType, 4
     lp = lp + 4
     k = LenB(.Name)
     CopyMemory b(lp), k, 4
     lp = lp + 4
     If k > 0 Then
      CopyMemory b(lp), ByVal StrPtr(.Name), k
      lp = lp + k
     End If
'     CopyMemory b(lp), .nPage, 4
'     lp = lp + 4
     CopyMemory b(lp), .Left, 4
     lp = lp + 4
'     CopyMemory b(lp), .Top, 4
'     lp = lp + 4
     CopyMemory b(lp), .Width, 4
     lp = lp + 4
     k = tDef(.nType).StringCount
     CopyMemory b(lp), k, 4
     lp = lp + 4
     For j = 0 To k - 1
      k = LenB(.sProps(j))
      CopyMemory b(lp), k, 4
      lp = lp + 4
      If k > 0 Then
       CopyMemory b(lp), ByVal StrPtr(.sProps(j)), k
       lp = lp + k
      End If
     Next j
     k = tDef(.nType).PropSize
     CopyMemory b(lp), k, 4
     lp = lp + 4
     If k > 0 Then
      CopyMemory b(lp), .bProps(0), k
      lp = lp + k
     End If
    Else 'empty
     lp = lp + 4
    End If
   End With
  Next i
 End With
End If
'compress?
If h.nDecompressedSize > 0 Then
 Select Case nMode
 Case 1 'LZSS
  h.nReserved = &H53535A4C
  h.nDataSize = cLZSS.CompressData(b, b2)
 Case 2 'LZMA!!
  h.nReserved = &H414D5A4C
  LZMACompress_Simple b, b2, i
  h.nDataSize = i
 Case 3 'zlib
  h.nReserved = &H62696C7A
  ZLibCompressByteArray b, b2, 9
  h.nDataSize = ZLibValueCompressedSize
 Case Else
  h.nDataSize = h.nDecompressedSize
  ReDim b2(h.nDataSize - 1)
  CopyMemory b2(0), b(0), h.nDataSize
 End Select
 Erase b
End If
'save file
Open FileName For Output As #1
Close
Open FileName For Binary As #1
With h
 .Signature0 = &H4754794D
 .Signature1 = &H4C
 .nPageCount = p.nPageCount
 .nOpCount = p.nOpCount
End With
Put #1, 1, h
If h.nDataSize > 0 Then
 Put #1, 25, b2
End If
Close
SavePrjFile = True
Exit Function
a:
Close
End Function

Public Function LoadPrjFile(p As typeProject, ByVal FileName As String) As Boolean
On Error GoTo a
Dim i As Long, j As Long, k As Long
Dim bErr As Boolean
Dim h As typePrjFileHeader
Dim b() As Byte, b2() As Byte, lp As Long
Dim cObj As New clsOperators
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
For i = 1 To objc
 cObj.ValidateOps p, objs(i)
Next i
'over
LoadPrjFile = True
Exit Function
a:
Close
End Function


'////////////////////////////////////////////////temp - file format convert

Private Sub CheckSize_1(b() As Byte, m As Long, ByVal lp As Long)
If lp >= m Then
 m = lp + 65536
 ReDim Preserve b(m - 1)
End If
End Sub

Public Function SavePrjFile_1(p As typeProject, ByVal FileName As String, Optional ByVal nMode As enumPrjCompressMode) As Boolean
On Error GoTo a
Dim s As String
Dim i As Long, j As Long, k As Long
Dim h As typePrjFileHeader
Dim pp As typeOperatorProp_DesignTime
Dim b() As Byte, lp As Long, lp2 As Long, m As Long
Dim m2 As Long 'header size
Dim m3 As Long 'total size
Dim sKey As String, nCount As Long
Dim b2() As Byte
Dim cLZSS As New clsLZSS
'calc size & put data
lp = 0
With p
 'header
 CheckSize_1 b, m, lp + 12
 CopyMemory b(lp), 12&, 4&
 lp = lp + 4
 CopyMemory b(lp), .nPageCount, 4&
 lp = lp + 4
 CopyMemory b(lp), .nOpCount, 4&
 lp = lp + 4
 'pages
 For i = 1 To .nPageCount
  With .Pages(i)
   lp2 = lp
   'resize
   k = LenB(.Name)
   m2 = k + 28&
   m3 = m2
   CheckSize_1 b, m, lp + m2
   'header size
   CopyMemory b(lp), m2, 4&
   'name
   CopyMemory b(lp + 8), k, 4&
   If k > 0 Then CopyMemory b(lp + 12), ByVal StrPtr(.Name), k
   'indent
   CopyMemory b(lp + k + 12), .nIndent, 1&
   'comments
   CopyMemory b(lp + k + 16), .nCommentCount, 2&
   'size
   CopyMemory b(lp + k + 20), 256&, 4&
   CopyMemory b(lp + k + 24), 128&, 4&
   '///
   lp = lp + m2
   For j = 1 To .nCommentCount
    With .Comments(j)
     k = LenB(.Name)
     m2 = 28& + k + LenB(.Value)
     m3 = m3 + m2
     CheckSize_1 b, m, lp + m2
     'save
     CopyMemory b(lp), .Left, 4
     lp = lp + 4
     CopyMemory b(lp), .Top, 4
     lp = lp + 4
     CopyMemory b(lp), .Width, 4
     lp = lp + 4
     CopyMemory b(lp), .Height, 4
     lp = lp + 4
     CopyMemory b(lp), .Color, 4
     lp = lp + 4
     CopyMemory b(lp), k, 4
     lp = lp + 4
     If k > 0 Then
      CopyMemory b(lp), ByVal StrPtr(.Name), k
      lp = lp + k
     End If
     k = LenB(.Value)
     CopyMemory b(lp), k, 4
     lp = lp + 4
     If k > 0 Then
      CopyMemory b(lp), ByVal StrPtr(.Value), k
      lp = lp + k
     End If
    End With
   Next j
   'total size
   m3 = lp - lp2
   CopyMemory b(lp2 + 4), m3, 4&
  End With
 Next i
 'operators
 For i = 1 To .nOpCount
  If .Operators(i).Flags >= 0 Then
   With .Operators(i)
    lp2 = lp
    'resize
    nCount = 0
    With tDef(.nType)
     sKey = .Name
     For j = 1 To .PropCount
      m3 = .props(j).nType
      If m3 <> eOPT_Group And m3 <> eOPT_Name Then nCount = nCount + 1
     Next j
    End With
    k = LenB(.Name)
    m2 = k + LenB(sKey) + 56&
    CheckSize_1 b, m, lp + m2
    'save
    CopyMemory b(lp), .nType, 4&
    lp = lp + 4
    CopyMemory b(lp), m2, 4&
    lp = lp + 8
    Select Case .nType
    Case int_OpType_Store, int_OpType_Load, int_OpType_Nop, int_OpType_Export
     j = 0
    Case Else
     j = 1
    End Select
    CopyMemory b(lp), j, 4&
    lp = lp + 4
    CopyMemory b(lp), k, 4&
    lp = lp + 4
    If k > 0 Then
     CopyMemory b(lp), ByVal StrPtr(.Name), k
     lp = lp + k
    End If
    CopyMemory b(lp), .nPage, 4
    lp = lp + 4
    CopyMemory b(lp), .Left, 4
    lp = lp + 4
    CopyMemory b(lp), .Top, 4
    lp = lp + 4
    CopyMemory b(lp), .Width, 4
    lp = lp + 4
    CopyMemory b(lp), 1&, 4&
    lp = lp + 4
    'TODO:animation
    lp = lp + 4
    'TODO:class prop
    lp = lp + 4
    CopyMemory b(lp), nCount, 4&
    lp = lp + 4
    '////////
    k = LenB(sKey)
    CopyMemory b(lp), k, 4&
    lp = lp + 4
    If k > 0 Then
     CopyMemory b(lp), ByVal StrPtr(sKey), k
     lp = lp + k
    End If
    '////////////////////////////////////////////////////////////////
    'properties
    For j = 1 To tDef(.nType).PropCount
     m3 = tDef(.nType).props(j).nType
     If m3 <> eOPT_Group And m3 <> eOPT_Name Then
      sKey = tDef(.nType).props(j).Name
      k = LenB(sKey)
      CheckSize_1 b, m, lp + k + 16&
      CopyMemory b(lp), k, 4&
      lp = lp + 4
      If k > 0 Then
       CopyMemory b(lp), ByVal StrPtr(sKey), k
       lp = lp + k
      End If
      Erase pp.iValue, pp.fValue
      PropRead p.Operators(i), tDef(.nType).props(j), pp
      Select Case tDef(.nType).props(j).nType
      Case eOPT_String, eOPT_Custom
       CopyMemory b(lp), 0&, 4&
       lp = lp + 4
       CopyMemory b(lp), 1&, 4&
       lp = lp + 4
       k = LenB(pp.sValue)
       CopyMemory b(lp), k, 4&
       lp = lp + 4
       If k > 0 Then
        CheckSize_1 b, m, lp + k
        CopyMemory b(lp), ByVal StrPtr(pp.sValue), k
        lp = lp + k
       End If
      Case eOPT_Color
       CopyMemory b(lp), 1&, 4&
       lp = lp + 4
       CopyMemory b(lp), 4&, 4&
       lp = lp + 4
       CopyMemory b(lp), 4&, 4&
       lp = lp + 4
       CheckSize_1 b, m, lp + 4
       CopyMemory b(lp), pp.iValue(0), 4&
       lp = lp + 4
      Case eOPT_Single, eOPT_PtFloat, eOPT_RectFloat
       CopyMemory b(lp), 2&, 4&
       lp = lp + 4
       CopyMemory b(lp), 4&, 4&
       lp = lp + 4
       CopyMemory b(lp), 16&, 4&
       lp = lp + 4
       CheckSize_1 b, m, lp + 16
       '???
       CopyMemory b(lp), pp.fValue(0), 4&
       CopyMemory b(lp + 4), pp.fValue(1), 4&
       CopyMemory b(lp + 8), pp.fValue(2), 4&
       CopyMemory b(lp + 12), pp.fValue(3), 4&
       '///
       lp = lp + 16
      Case Else
       CopyMemory b(lp), 1&, 4&
       lp = lp + 4
       CopyMemory b(lp), 4&, 4&
       lp = lp + 4
       CopyMemory b(lp), 16&, 4&
       lp = lp + 4
       CheckSize_1 b, m, lp + 16
       '???
       CopyMemory b(lp), pp.iValue(0), 4&
       CopyMemory b(lp + 4), pp.iValue(1), 4&
       CopyMemory b(lp + 8), pp.iValue(2), 4&
       CopyMemory b(lp + 12), pp.iValue(3), 4&
       '///
       lp = lp + 16
'      CopyMemory b(lp), .BasicDataType, 4&
'      lp = lp + 4
'      CopyMemory b(lp), .nElementCount, 4&
'      lp = lp + 4
'      k = .nSize
'      CopyMemory b(lp), k, 4&
'      lp = lp + 4
'      If k > 0 Then
'       CopyMemory b(lp), .d(0), k
'       lp = lp + k
'      End If
      End Select
     End If
    Next j
    '////////////////////////////////////////////////////////////////
    m3 = lp - lp2
    CopyMemory b(lp2 + 8), m3, 4&
   End With
  Else 'RLE optimize
   k = 0
   Do While i < p.nOpCount
    If p.Operators(i + 1).Flags >= 0 Then Exit Do
    i = i + 1
    k = k - 1
   Loop
   CheckSize_1 b, m, lp + 4
   CopyMemory b(lp), k, 4&
   lp = lp + 4
  End If
 Next i
End With
If lp > 0 Then
 h.nDecompressedSize = lp
 ReDim Preserve b(lp - 1)
End If
'compress?
If h.nDecompressedSize > 0 Then
 Select Case nMode
 Case 1 'LZSS
  h.nDataSize = cLZSS.CompressData(b, b2)
 Case 2 'LZMA!!
  LZMACompress_Simple b, b2, i
  h.nDataSize = i
 Case 3 'zlib
  ZLibCompressByteArray b, b2, 9
  h.nDataSize = ZLibValueCompressedSize
 Case Else
  h.nDataSize = h.nDecompressedSize
  ReDim b2(h.nDataSize - 1)
  CopyMemory b2(0), b(0), h.nDataSize
 End Select
 Erase b
End If
'save file
Open FileName + ".new.myt" For Output As #1
Close
Open FileName + ".new.myt" For Binary As #1
With h
 .Signature0 = &H4754794D
 .Signature1 = &H14C&
 .nPageCount = nMode
End With
Put #1, 1, h.Signature0
Put #1, 5, h.Signature1
Put #1, 7, h.nPageCount
Put #1, 9, h.nDataSize
Put #1, 13, h.nDecompressedSize
If h.nDataSize > 0 Then
 Put #1, 17, b2
End If
Close
SavePrjFile_1 = True
Exit Function
a:
Close
End Function


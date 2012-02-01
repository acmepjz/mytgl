VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsOperators"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub ZeroMemory Lib "kernel32.dll" Alias "RtlZeroMemory" (ByRef Destination As Any, ByVal Length As Long)
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private Declare Function QueryPerformanceCounter Lib "kernel32.dll" (ByRef lpPerformanceCount As Currency) As Long
Private Declare Function QueryPerformanceFrequency Lib "kernel32.dll" (ByRef lpFrequency As Currency) As Long

Private Declare Sub DebugBreak Lib "kernel32.dll" ()

Private bm() As cAlphaDibSection '1-based
Private bmCount As Long

Friend Property Get BitmapCount() As Long
BitmapCount = bmCount
End Property

Friend Property Get TheBitmap(ByVal Index As Long) As cAlphaDibSection
If Index > 0 And Index <= bmCount Then
 Set TheBitmap = bm(Index)
End If
End Property

Friend Sub Clear(p As typeProject)
pClear
Erase p.Operators, p.Pages
p.nOpCount = 0
p.nPageCount = 0
End Sub

Friend Sub ClearMemory(p As typeProject)
pClear
ClearFlags p, int_OpFlags_InMemory
End Sub

Private Sub pClear()
Erase bm
bmCount = 0
End Sub

Friend Function AddPage(p As typeProject, ByVal Name As String)
With p
 .nPageCount = .nPageCount + 1
 ReDim Preserve .Pages(1 To .nPageCount)
 .Pages(.nPageCount).Name = Name
 AddPage = .nPageCount
End With
End Function

Friend Property Get IsNothing(ByVal Index As Long) As Boolean
If Index > 0 And Index <= bmCount Then
 IsNothing = bm(Index) Is Nothing
Else
 IsNothing = True
End If
End Property

Friend Sub GetDescriptor(ByVal Index As Long, t As typeAlphaDibSectionDescriptor)
If IsNothing(Index) Then
 ZeroMemory t, 16&
Else
 With bm(Index)
  t.hdc = .hdc
  t.Width = .Width
  t.Height = .Height
  t.lpbm = .DIBSectionBitsPtr
 End With
End If
End Sub

Friend Function Create(ByVal Index As Long) As cAlphaDibSection
If Index <= 0 Then Exit Function
If Index > bmCount Then
 bmCount = Index
 ReDim Preserve bm(1 To bmCount)
End If
Set bm(Index) = New cAlphaDibSection
Set Create = bm(Index)
End Function

Friend Sub Delete(ByVal Index As Long)
If Index > 0 And Index <= bmCount Then
 Set bm(Index) = Nothing
End If
End Sub

Friend Function FindEmptyBitmap() As Long
Dim i As Long
For i = 1 To bmCount
 If bm(i) Is Nothing Then
  FindEmptyBitmap = i
  Exit Function
 End If
Next i
FindEmptyBitmap = bmCount + 1
End Function

Friend Function FindEmptyBitmapEx(p As typeProject) As Long
Dim i As Long
'find an operator that is not loaded into memory
For i = 1 To p.nOpCount
 With p.Operators(i)
  If .Flags >= 0 And (.Flags And int_OpFlags_InMemory) = 0 Then
   If .nBmIndex > 0 And .nBmIndex <= bmCount Then
    FindEmptyBitmapEx = .nBmIndex
    .nBmIndex = 0
    Exit Function
   End If
  End If
 End With
Next i
FindEmptyBitmapEx = FindEmptyBitmap
End Function

Private Sub Class_Terminate()
pClear
End Sub

'binary search
'return value=0-count
Friend Function PageHitTest(p As typeProject, ByVal nPage As Long, ByVal Left As Long, ByVal Top As Long) As Long
Dim i As Long, j As Long, k As Long
If nPage <= 0 Then Exit Function
With p
 If nPage <= .nPageCount And Top >= 0 And Top < int_Page_Height Then 'fix an unknown bug
  With .Pages(nPage).Rows(Top)
   If .nOpCount > 0 Then
    If Left < p.Operators(.idxOp(1)).Left Then
     PageHitTest = 0
    ElseIf Left >= p.Operators(.idxOp(.nOpCount)).Left Then
     PageHitTest = .nOpCount
    Else
     i = 1
     j = .nOpCount - 1
     Do
      Debug.Assert i <= j
      k = (i + j) \ 2
      If Left < p.Operators(.idxOp(k)).Left Then 'move left
       j = k - 1
      ElseIf Left >= p.Operators(.idxOp(k + 1)).Left Then 'move right
       i = k + 1
      Else
       PageHitTest = k
       Exit Do
      End If
     Loop
    End If
   End If
  End With
 End If
End With
End Function

'select an area
Friend Function PageHitTestEx(p As typeProject, ByVal nPage As Long, ByVal Left As Long, ByVal Top As Long, ByVal Right As Long, ByVal Bottom As Long, nSelected() As Long, Optional ByVal bAutoSelect As Boolean) As Long
Dim i As Long, j As Long, k As Long, m As Long
m = 0
Erase nSelected
With p
 If nPage > 0 And nPage <= .nPageCount Then
  If Top < 0 Then Top = 0
  If Bottom >= int_Page_Height Then Bottom = int_Page_Height - 1
  For i = Top To Bottom
   With .Pages(nPage).Rows(i)
    For j = 1 To .nOpCount
     k = .idxOp(j)
     With p.Operators(k)
      If .Left <= Right And .Left + .Width > Left And .Flags >= 0 Then
       If bAutoSelect Then .Flags = .Flags Or int_OpFlags_Selected
       m = m + 1
       ReDim Preserve nSelected(1 To m)
       nSelected(m) = k
      End If
     End With
    Next j
   End With
  Next i
 End If
End With
PageHitTestEx = m
End Function

'doesn't recalculate bitmap
Friend Function AddOperator(p As typeProject, ByVal nType As Long, ByVal nPage As Long, ByVal Left As Long, ByVal Top As Long, Optional ByVal Width As Long = 4, Optional ByVal Height As Long = 1) As Long
Dim i As Long, j As Long, m As Long
Dim s As String
Dim pp As typeOperatorProp_DesignTime
With p
 If nPage > 0 And nPage <= .nPageCount Then
  If Left >= 0 And Left < int_Page_Width And Top >= 0 And Top < int_Page_Height Then
   i = PageHitTest(p, nPage, Left, Top)
   With .Pages(nPage).Rows(Top)
    If i > 0 Then
     With p.Operators(.idxOp(i))
      If Left < .Left + .Width Then Width = 0
     End With
    End If
    If i < .nOpCount Then
     j = p.Operators(.idxOp(i + 1)).Left - Left
     If Width > j Then Width = j
    End If
    j = int_Page_Width - Left
    If Width > j Then Width = j
    If Width > 0 Then
     '////////////////////////////////valid!!! about to add operator
     For j = 1 To p.nOpCount
      If p.Operators(j).Flags < 0 Then Exit For
     Next j
     If j > p.nOpCount Then
      p.nOpCount = j
      '/////fix an unknown bug Error '10'
      'get address of SafeArray
      CopyMemory m, ByVal VarPtr(p) + 12&, 4&
      If m <> 0 Then
       'get address of cLocks
       m = m + 8&
       'get old cLocks
       CopyMemory pp, ByVal m, 4&
       If pp.iValue(0) <> 0 Then
        'error!!!
        MsgBox "cLocks=" + CStr(pp.iValue(0)) + "!?", vbCritical, "Fatal Error"
       End If
       ''set cLocks to zero (!!!)
       'ZeroMemory ByVal m, 4&
      End If
      ReDim Preserve p.Operators(1 To j)
     End If
     'insert
     i = i + 1
     .nOpCount = .nOpCount + 1
     ReDim Preserve .idxOp(1 To .nOpCount)
     If i < .nOpCount Then CopyMemory .idxOp(i + 1), .idxOp(i), (.nOpCount - i) * 4&
     .idxOp(i) = j
     'write properties
     With p.Operators(j)
      .Name = ""
      .nPage = nPage
      .Left = Left
      .Top = Top
      .Width = Width
      .nType = nType
      .Flags = 0
      .nBmIndex = 0
      m = tDef(nType).PropSize
      If m > 0 Then
       ReDim .bProps(m - 1)
      Else
       Erase .bProps
      End If
      m = tDef(nType).StringCount
      If m > 0 Then
       ReDim .sProps(m - 1)
      Else
       Erase .sProps
      End If
     End With
     m = tDef(nType).PropCount
     For i = 1 To m
      s = tDef(nType).props(i).sDefault
      If s <> "" Then
       PropFromString s, tDef(nType).props(i), pp
       PropWrite p.Operators(j), tDef(nType).props(i), pp
      End If
     Next i
     '////////////////////////////////
     AddOperator = j
    End If
   End With
  End If
 End If
End With
End Function

'doesn't recalculate bitmap
Friend Function MoveOperatorByIndex(p As typeProject, ByVal Index As Long, ByVal NewLeft As Long, ByVal NewTop As Long, Optional ByVal NewWidth As Long, Optional ByVal NewHeight As Long) As Boolean
Dim nPage As Long
Dim i As Long, j As Long
With p
 If .Operators(Index).Flags >= 0 Then
  nPage = .Operators(Index).nPage
  If NewWidth <= 0 Then NewWidth = .Operators(Index).Width
  'If NewHeight <= 0 Then
  If NewLeft >= 0 And NewLeft + NewWidth <= int_Page_Width And NewTop >= 0 And NewTop < int_Page_Height Then
   i = PageHitTest(p, nPage, NewLeft, NewTop)
   With .Pages(nPage).Rows(NewTop)
    If i > 0 Then
     j = .idxOp(i)
     If j <> Index Then
      With p.Operators(j)
       If NewLeft < .Left + .Width Then NewWidth = 0
      End With
     End If
    End If
    If i < .nOpCount Then
     j = .idxOp(i + 1)
     If j <> Index Then
      If NewWidth > p.Operators(j).Left - NewLeft Then NewWidth = 0
     End If
    End If
    If NewWidth > 0 Then
     '////////////////////////////////valid!!! about to move
     j = p.Operators(Index).Top
     If j = NewTop Then
      'same line
      For j = 1 To .nOpCount
       If .idxOp(j) = Index Then
        If i = j Or i = j - 1 Then j = -1       'do nothing!!! set flags
        Exit For
       End If
      Next j
      Debug.Assert j <= .nOpCount
      If j > 0 Then
       If i > j Then 'move right
        CopyMemory .idxOp(j), .idxOp(j + 1), (i - j) * 4&
        .idxOp(i) = Index
       Else 'move left
        i = i + 1
        CopyMemory .idxOp(i + 1), .idxOp(i), (j - i) * 4&
        .idxOp(i) = Index
       End If
      End If
     Else
      'insert new
      i = i + 1
      .nOpCount = .nOpCount + 1
      ReDim Preserve .idxOp(1 To .nOpCount)
      If i < .nOpCount Then CopyMemory .idxOp(i + 1), .idxOp(i), (.nOpCount - i) * 4&
      .idxOp(i) = Index
      'delete old
      With p.Pages(nPage).Rows(j)
       For i = 1 To .nOpCount
        If .idxOp(i) = Index Then Exit For
       Next i
       Debug.Assert i <= .nOpCount
       If .nOpCount <= 1 Then
        .nOpCount = 0
        Erase .idxOp
       Else
        If i < .nOpCount Then CopyMemory .idxOp(i), .idxOp(i + 1), (.nOpCount - i) * 4&
        .nOpCount = .nOpCount - 1
        ReDim Preserve .idxOp(1 To .nOpCount)
       End If
      End With
     End If
     'write properties
     With p.Operators(Index)
      .Left = NewLeft
      .Top = NewTop
      .Width = NewWidth
     End With
     '////////////////////////////////
     MoveOperatorByIndex = True
    End If
   End With
  End If
 End If
End With
End Function

Friend Function MoveOperator(p As typeProject, ByVal nPage As Long, ByVal Left As Long, ByVal Top As Long, ByVal NewLeft As Long, ByVal NewTop As Long, Optional ByVal NewWidth As Long, Optional ByVal NewHeight As Long) As Boolean
Dim i As Long
With p
 If nPage > 0 And nPage <= .nPageCount Then
  If Left >= 0 And Left < int_Page_Width And Top >= 0 And Top < int_Page_Height Then
   i = PageHitTest(p, nPage, Left, Top)
   If i > 0 Then
    i = .Pages(nPage).Rows(Top).idxOp(i)
    With .Operators(i)
     If Left >= .Left And Left < .Left + .Width Then
      MoveOperator = MoveOperatorByIndex(p, i, NewLeft, NewTop, NewWidth, NewHeight)
     End If
    End With
   End If
  End If
 End If
End With
End Function

'doesn't recalculate bitmap
Friend Function DeleteOperatorByIndex(p As typeProject, ByVal Index As Long) As Boolean
Dim i As Long
With p.Operators(Index)
 If .Flags >= 0 Then
  'delete row
  With p.Pages(.nPage).Rows(.Top)
   For i = 1 To .nOpCount
    If .idxOp(i) = Index Then Exit For
   Next i
   Debug.Assert i <= .nOpCount
   If .nOpCount <= 1 Then
    .nOpCount = 0
    Erase .idxOp
   Else
    If i < .nOpCount Then CopyMemory .idxOp(i), .idxOp(i + 1), (.nOpCount - i) * 4&
    .nOpCount = .nOpCount - 1
    ReDim Preserve .idxOp(1 To .nOpCount)
   End If
  End With
  'delete bitmap
  Delete .nBmIndex 'if index<0 then this is a reference only
  'delete data
  .Name = ""
  .nPage = 0
  .Left = 0
  .Top = 0
  .Width = 0
  .nType = 0
  .Flags = -1
  .nBmIndex = 0
  Erase .bProps, .sProps
  'over
  DeleteOperatorByIndex = True
 End If
End With
End Function

Friend Function DeleteOperator(p As typeProject, ByVal nPage As Long, ByVal Left As Long, ByVal Top As Long) As Boolean
Dim i As Long
With p
 If nPage > 0 And nPage <= .nPageCount Then
  If Left >= 0 And Left < int_Page_Width And Top >= 0 And Top < int_Page_Height Then
   i = PageHitTest(p, nPage, Left, Top)
   If i > 0 Then
    i = .Pages(nPage).Rows(Top).idxOp(i)
    With .Operators(i)
     If Left >= .Left And Left < .Left + .Width Then
      DeleteOperator = DeleteOperatorByIndex(p, i)
     End If
    End With
   End If
  End If
 End If
End With
End Function

'doesn't recalculate bitmap
Friend Function DeletePage(p As typeProject, ByVal nPage As Long) As Boolean
Dim i As Long
With p
 If nPage > 0 And nPage <= .nPageCount Then
  If .nPageCount <= 1 Then
   'delete all!!!
   Clear p
  Else
   'delete operator
   For i = 1 To p.nOpCount
    With p.Operators(i)
     If .Flags >= 0 Then
      If .nPage = nPage Then
       'delete bitmap
       Delete .nBmIndex 'if index<0 then this is a reference only
       'delete data
       .Name = ""
       .nPage = 0
       .Left = 0
       .Top = 0
       .Width = 0
       .nType = 0
       .Flags = -1
       .nBmIndex = 0
       Erase .bProps, .sProps
      ElseIf .nPage > nPage Then
       .nPage = .nPage - 1
      End If
     End If
    End With
   Next i
   'delete page
   .nPageCount = .nPageCount - 1
   For i = nPage To .nPageCount
    .Pages(i) = .Pages(i + 1)
   Next i
   ReDim Preserve .Pages(1 To .nPageCount)
  End If
  'over
  DeletePage = True
 End If
End With
End Function

Friend Function DeletePageEx(p As typeProject, ByVal nStart As Long, ByVal nEnd As Long) As Boolean
Dim i As Long
With p
 If nStart > 0 And nStart <= nEnd And nEnd <= .nPageCount Then
  If nStart = 1 And nEnd = .nPageCount Then
   'delete all!!!
   Clear p
  Else
   'delete operator
   For i = 1 To p.nOpCount
    With p.Operators(i)
     If .Flags >= 0 Then
      If .nPage >= nStart And .nPage <= nEnd Then
       'delete bitmap
       Delete .nBmIndex 'if index<0 then this is a reference only
       'delete data
       .Name = ""
       .nPage = 0
       .Left = 0
       .Top = 0
       .Width = 0
       .nType = 0
       .Flags = -1
       .nBmIndex = 0
       Erase .bProps, .sProps
      ElseIf .nPage > nEnd Then
       .nPage = .nPage - nEnd + nStart - 1
      End If
     End If
    End With
   Next i
   'delete page
   .nPageCount = .nPageCount - nEnd + nStart - 1
   For i = nStart To .nPageCount
    .Pages(i) = .Pages(i + nEnd - nStart + 1)
   Next i
   ReDim Preserve .Pages(1 To .nPageCount)
  End If
  'over
  DeletePageEx = True
 End If
End With
End Function

Friend Function MovePageEx(p As typeProject, ByVal nStart As Long, ByVal nEnd As Long, ByVal nNewStart As Long, Optional ByVal lpThePageState As Long) As Boolean
Dim i As Long
Dim tmp() As typePage, m As Long
Dim tmp2() As Long
With p
 'validate
 If nStart > 0 And nStart <= nEnd And nEnd <= .nPageCount And _
 nNewStart > 0 And nNewStart <= .nPageCount And _
 (nNewStart < nStart Or nNewStart > nEnd) Then
  'backup
  m = nEnd - nStart + 1
  ReDim tmp(1 To m)
  For i = nStart To nEnd
   tmp(i - nStart + 1) = p.Pages(i)
  Next i
  If nNewStart < nStart Then 'move left
   For i = nEnd To nNewStart + m Step -1
    p.Pages(i) = p.Pages(i - m)
   Next i
   For i = 1 To m
    p.Pages(i + nNewStart - 1) = tmp(i)
   Next i
   'change state
   If lpThePageState <> 0 Then
    ReDim tmp2(1 To m)
    CopyMemory tmp2(1), ByVal (lpThePageState + (nStart - 1) * 4&), m * 4&
    CopyMemory ByVal (lpThePageState + (nNewStart + m - 1) * 4&), _
    ByVal (lpThePageState + (nNewStart - 1) * 4&), (nStart - nNewStart) * 4&
    CopyMemory ByVal (lpThePageState + (nNewStart - 1) * 4&), tmp2(1), m * 4&
   End If
   'change operator
   For i = 1 To p.nOpCount
    With p.Operators(i)
     If .Flags >= 0 And .nPage <= nEnd Then
      If .nPage >= nStart Then
       .nPage = nNewStart + .nPage - nStart
      ElseIf .nPage >= nNewStart Then
       .nPage = .nPage + m
      End If
     End If
    End With
   Next i
  Else 'move right
   For i = nStart To nNewStart - 1
    p.Pages(i) = p.Pages(i + m)
   Next i
   For i = 1 To m
    p.Pages(i + nNewStart - 1) = tmp(i)
   Next i
   'change state
   If lpThePageState <> 0 Then
    ReDim tmp2(1 To m)
    CopyMemory tmp2(1), ByVal (lpThePageState + (nStart - 1) * 4&), m * 4&
    CopyMemory ByVal (lpThePageState + (nStart - 1) * 4&), _
    ByVal (lpThePageState + nEnd * 4&), (nNewStart - nStart) * 4&
    CopyMemory ByVal (lpThePageState + (nNewStart - 1) * 4&), tmp2(1), m * 4&
   End If
   'change operator
   For i = 1 To p.nOpCount
    With p.Operators(i)
     If .Flags >= 0 And .nPage >= nStart Then
      If .nPage <= nEnd Then
       .nPage = nNewStart + .nPage - nStart
      ElseIf .nPage < nNewStart + m Then
       .nPage = .nPage - m
      End If
     End If
    End With
   Next i
  End If
  MovePageEx = True
 End If
End With
End Function

Friend Sub ClearFlags(p As typeProject, Optional ByVal Flags As Long = int_OpFlags_Selected)
Dim i As Long
Flags = Not Flags
For i = 1 To p.nOpCount
 With p.Operators(i)
  If .Flags >= 0 Then .Flags = .Flags And Flags
 End With
Next i
End Sub

Friend Sub SetFlags(p As typeProject, Optional ByVal Flags As Long = int_OpFlags_Selected, Optional ByVal nPage As Long)
Dim i As Long, j As Long
With p
 If nPage > 0 And nPage <= .nPageCount Then
  For i = 0 To int_Page_Height - 1
   For j = 1 To .Pages(nPage).Rows(i).nOpCount
    With p.Operators(.Pages(nPage).Rows(i).idxOp(j))
     If .Flags >= 0 Then .Flags = .Flags Or Flags
    End With
   Next j
  Next i
 Else
  For i = 1 To .nOpCount
   With .Operators(i)
    If .Flags >= 0 Then .Flags = .Flags Or Flags
   End With
  Next i
 End If
End With
End Sub

Friend Function GetStoreObjects(p As typeProject, sto() As typeStoreOp_DesignTime) As Long
Dim i As Long, m As Long
Erase sto
With p
 For i = 1 To .nOpCount
  With .Operators(i)
   If .Flags >= 0 And .nType = int_OpType_Store Then
    m = m + 1
    ReDim Preserve sto(1 To m)
    sto(m).Name = .Name
    sto(m).Index = i
   End If
  End With
 Next i
End With
GetStoreObjects = m
End Function

Friend Sub ValidateAllLoadOps(p As typeProject, ByVal s As String)
Dim i As Long
For i = 1 To p.nOpCount
 With p.Operators(i)
  If .nType = int_OpType_Load Then
   If .sProps(0) = s Then
    ValidateOps p, i
   End If
  End If
 End With
Next i
End Sub

'bug!!! BFS-algorithm
'    [Flat]
'[Store][Nop]
'       [Nop]
'       .....
'       [Nop]
' [Load][Nop]
' [   Add   ]
'
' change the size? -> bug!!!

'///////////////////////////////validate operators - won't auto check the name change of store operator
Friend Sub ValidateOps(p As typeProject, ByVal Index As Long)
'store object
Dim sto() As typeStoreOp_DesignTime, stoc As Long
'objects queue
Dim objs() As Long, objc As Long, objm As Long
Dim objNow As Long
'last-row cache
Dim nLastPage As Long, nLastLeft As Long, nLastTop As Long
Dim nLastInputIndex As Long, nLastOutputIndex As Long
'input operator
Dim nOpCount As Long
'Dim nInputOps(int_Page_Width - 1) As Long
Dim tInputBm(int_Page_Width - 1) As typeAlphaDibSectionDescriptor '0-based
'misc
Dim i As Long
Dim idx As Long, w As Long, h As Long
Dim tmp As Long
Dim bErr As Boolean
'start calcuate!!
With p
 If Index > 0 And Index <= .nOpCount Then
  ClearFlags p, int_OpFlags_Dirty
  'init obj
  objm = 256&
  ReDim objs(1 To objm)
  objs(1) = Index
  objc = 1&
  objNow = 1&
  stoc = -1 'store object is uninitalized
  Do Until objNow > objc
   With .Operators(objs(objNow))
    Debug.Assert .Flags >= 0
    'clear flags
    .Flags = .Flags And Not (int_OpFlags_Error Or int_OpFlags_InMemory)
    bErr = False
    'update cache
    If nLastPage <> .nPage Or nLastTop <> .Top Or nLastLeft > .Left Then
     nLastInputIndex = 1
     nLastOutputIndex = 1
     nLastPage = .nPage
     nLastTop = .Top
    End If
    nLastLeft = .Left
    'check input
    nOpCount = 0
    If .Top > 0 And Not bErr Then
     'find left of input operator
     tmp = .Left
     With p.Pages(.nPage).Rows(.Top - 1)
      For i = nLastInputIndex To .nOpCount
       With p.Operators(.idxOp(i))
        Debug.Assert .Flags >= 0
        If .Left + .Width > tmp Then Exit For
       End With
      Next i
     End With
     nLastInputIndex = i
     'find right of input operator
     tmp = .Left + .Width
     With p.Pages(.nPage).Rows(.Top - 1)
      For i = nLastInputIndex To .nOpCount
       With p.Operators(.idxOp(i))
        Debug.Assert .Flags >= 0
        If .Left >= tmp Then Exit For
        'check error
        If .Flags And int_OpFlags_Error Then
         bErr = True
         Exit For
        End If
        'add operator
        w = .nBmWidth
        h = .nBmHeight
        Debug.Assert w > 0 And h > 0
        With tInputBm(nOpCount) '0-based
         .hdc = 1 'dummy
         .Width = w
         .Height = h
         .lpbm = 1 'dummy
        End With
        nOpCount = nOpCount + 1
       End With
      Next i
     End With
    End If
    'check output
    If .Top < int_Page_Height - 1 Then
     'find left of output operator
     tmp = .Left
     With p.Pages(.nPage).Rows(.Top + 1)
      For i = nLastOutputIndex To .nOpCount
       With p.Operators(.idxOp(i))
        Debug.Assert .Flags >= 0
        If .Left + .Width > tmp Then Exit For
       End With
      Next i
     End With
     nLastOutputIndex = i
     'find right of output operator
     tmp = .Left + .Width
     With p.Pages(.nPage).Rows(.Top + 1)
      For i = nLastOutputIndex To .nOpCount
       idx = .idxOp(i)
       If p.Operators(idx).Left >= tmp Then Exit For
       'add queue
       pOperatorQueueAddItem p, idx, Index, objs, objc, objm, bErr, objNow
      Next i
     End With
     'find store-load operator
     If .nType = int_OpType_Store Then
      For idx = 1 To p.nOpCount
       If p.Operators(idx).Flags >= 0 And p.Operators(idx).nType = int_OpType_Load Then
        If p.Operators(idx).sProps(0) = .Name Then
         'add queue
         pOperatorQueueAddItem p, idx, Index, objs, objc, objm, bErr, objNow
        End If
       End If
      Next idx
     End If
    End If
    'validate it
    If Not bErr Then
     Select Case .nType
     Case int_OpType_Load
      'process load operator
      If nOpCount = 0 Then
       'initalize list
       If stoc < 0 Then
        stoc = GetStoreObjects(p, sto)
       End If
       'search the list
       For i = 1 To stoc
        If sto(i).Name = .sProps(0) Then
         With p.Operators(sto(i).Index)
          If .Flags >= 0 Then
           'found it
           If .Flags And int_OpFlags_Error Then
            bErr = True
           Else
            w = .nBmWidth
            h = .nBmHeight
            Debug.Assert w > 0 And h > 0
           End If
           Exit For
          End If
         End With
        End If
       Next i
       If i > stoc Then 'not found, error
        bErr = True
       End If
      Else 'input error!!!
       bErr = True
      End If
     Case int_OpType_Store, int_OpType_Nop, int_OpType_Export
      'simply copy
      If nOpCount = 1 Then
       With tInputBm(0)
        w = .Width
        h = .Height
       End With
      Else 'input error!!!
       bErr = True
      End If
     Case Else
      If Not ValidateOperator(nOpCount, tInputBm, .nType, .bProps, .sProps, w, h) Then bErr = True
     End Select
    End If
    'set size and flags
    If bErr Then
     .Flags = .Flags Or int_OpFlags_Error
    Else
     .nBmWidth = w '??
     .nBmHeight = h '??
    End If
   End With
   'next
   objNow = objNow + 1
  Loop
 End If
End With
End Sub

Private Sub pOperatorQueueAddItem(p As typeProject, ByVal idx As Long, ByVal idxOld As Long, objs() As Long, objc As Long, objm As Long, bErr As Boolean, ByVal objNow As Long)
Dim j As Long
       With p.Operators(idx)
        Debug.Assert .Flags >= 0
        'check recursive reference
        If idx = idxOld Then
         'it errors!!!
         For j = 1 To objNow - 1
          With p.Operators(objs(j))
           .Flags = .Flags Or int_OpFlags_Error
          End With
         Next j
         bErr = True
        ElseIf (.Flags And int_OpFlags_Dirty) = 0 Then
         'add queue
         .Flags = .Flags Or int_OpFlags_Dirty
         objc = objc + 1
         If objc > objm Then
          objm = objm + 256&
          ReDim Preserve objs(1 To objm)
         End If
         objs(objc) = idx
        End If
       End With
End Sub

Friend Sub SetNotInMemoryFlags(p As typeProject, ByVal Index As Long)
'objects queue
Dim objs() As Long, objc As Long, objm As Long
Dim objNow As Long
'last-row cache
Dim nLastPage As Long, nLastLeft As Long, nLastTop As Long
Dim nLastOutputIndex As Long
'misc
Dim i As Long
Dim idx As Long, w As Long, h As Long
Dim tmp As Long
Dim bErr As Boolean
'start calcuate!!
With p
 If Index > 0 And Index <= .nOpCount Then
  'init obj
  objm = 256&
  ReDim objs(1 To objm)
  objs(1) = Index
  objc = 1&
  objNow = 1&
  Do Until objNow > objc
   With .Operators(objs(objNow))
    Debug.Assert .Flags >= 0
    'clear flags
    .Flags = .Flags And Not int_OpFlags_InMemory
    'update cache
    If nLastPage <> .nPage Or nLastTop <> .Top Or nLastLeft > .Left Then
     nLastOutputIndex = 1
     nLastPage = .nPage
     nLastTop = .Top
    End If
    nLastLeft = .Left
    'check output
    If .Top < int_Page_Height - 1 Then
     'find left of output operator
     tmp = .Left
     With p.Pages(.nPage).Rows(.Top + 1)
      For i = nLastOutputIndex To .nOpCount
       With p.Operators(.idxOp(i))
        Debug.Assert .Flags >= 0
        If .Left + .Width > tmp Then Exit For
       End With
      Next i
     End With
     nLastOutputIndex = i
     'find right of output operator
     tmp = .Left + .Width
     With p.Pages(.nPage).Rows(.Top + 1)
      For i = nLastOutputIndex To .nOpCount
       idx = .idxOp(i)
       With p.Operators(idx)
        If .Left >= tmp Then Exit For
        'add queue?
        If .Flags And int_OpFlags_InMemory Then
         objc = objc + 1
         If objc > objm Then
          objm = objm + 256&
          ReDim Preserve objs(1 To objm)
         End If
         objs(objc) = idx
        End If
       End With
      Next i
     End With
     'find store-load operator
     If .nType = int_OpType_Store Then
      For idx = 1 To p.nOpCount
       If p.Operators(idx).Flags >= 0 And p.Operators(idx).nType = int_OpType_Load Then
        If p.Operators(idx).sProps(0) = .Name Then
         With p.Operators(idx)
          'add queue?
          If .Flags And int_OpFlags_InMemory Then
           objc = objc + 1
           If objc > objm Then
            objm = objm + 256&
            ReDim Preserve objs(1 To objm)
           End If
           objs(objc) = idx
          End If
         End With
        End If
       End If
      Next idx
     End If
    End If
   End With
   'next
   objNow = objNow + 1
  Loop
 End If
End With
End Sub

'TODO:progress bar
Friend Function ShowOperator(p As typeProject, ByVal Index As Long, Optional obj As IOperatorCalc) As Double
'store object
Dim sto() As typeStoreOp_DesignTime, stoc As Long
'objects queue
Dim objs() As Long, objc As Long, objm As Long
Dim objNow As Long
Dim obj2() As typeOperatorCalc_DesignTime
'last-row cache
Dim nLastPage As Long, nLastLeft As Long, nLastTop As Long
Dim nLastInputIndex As Long
'input operator
Dim nOpCount As Long
Dim nInputOps(int_Page_Width - 1) As Long   '0-based
Dim tInputBm(int_Page_Width - 1) As typeAlphaDibSectionDescriptor '0-based
'misc
Dim i As Long
Dim idx As Long
Dim tmp As Long
Dim bAbort As Boolean
'calc time
Dim t1 As Currency, t2 As Currency
'start calcuate!!
With p
 If Index > 0 And Index <= .nOpCount Then
  'doesn't need to redraw?
  If .Operators(Index).Flags And (int_OpFlags_Error Or int_OpFlags_InMemory) Then
   Exit Function
  End If
  'get time
  QueryPerformanceCounter t1
  'clear flags
  ClearFlags p, int_OpFlags_Dirty
  'init obj
  objm = 256&
  ReDim objs(1 To objm), obj2(1 To objm)
  objs(1) = Index
  objc = 1&
  objNow = 1&
  stoc = -1 'store object is uninitalized
  Do Until objNow > objc
   With .Operators(objs(objNow))
    Debug.Assert .Flags >= 0
    'error!?
    Debug.Assert (.Flags And int_OpFlags_Error) = 0
    'update cache
    If nLastPage <> .nPage Or nLastTop <> .Top Or nLastLeft > .Left Then
     nLastInputIndex = 1
     nLastPage = .nPage
     nLastTop = .Top
    End If
    nLastLeft = .Left
    'check input
    nOpCount = 0
    If .Top > 0 Then
     'find left of input operator
     tmp = .Left
     With p.Pages(.nPage).Rows(.Top - 1)
      For i = nLastInputIndex To .nOpCount
       With p.Operators(.idxOp(i))
        Debug.Assert .Flags >= 0
        If .Left + .Width > tmp Then Exit For
       End With
      Next i
     End With
     nLastInputIndex = i
     'find right of input operator
     tmp = .Left + .Width
     With p.Pages(.nPage).Rows(.Top - 1)
      For i = nLastInputIndex To .nOpCount
       idx = .idxOp(i)
       With p.Operators(idx)
        Debug.Assert .Flags >= 0
        If .Left >= tmp Then Exit For
        'error!?
        Debug.Assert (.Flags And int_OpFlags_Error) = 0
        'add operator
        nInputOps(nOpCount) = idx
        nOpCount = nOpCount + 1
        'add queue?
        If (.Flags And (int_OpFlags_Dirty Or int_OpFlags_InMemory)) = 0 Then
         .Flags = .Flags Or int_OpFlags_Dirty
         objc = objc + 1
         If objc > objm Then
          objm = objm + 256&
          ReDim Preserve objs(1 To objm), obj2(1 To objm)
         End If
         objs(objc) = idx
        End If
       End With
      Next i
     End With
    End If
    'load-object add queue?
    If .nType = int_OpType_Load Then
     If stoc < 0 Then
      stoc = GetStoreObjects(p, sto)
     End If
     For i = 1 To stoc
      If .sProps(0) = sto(i).Name Then Exit For
     Next i
     Debug.Assert i <= stoc
     idx = sto(i).Index
     'add operator
     Debug.Assert nOpCount = 0
     nInputOps(0) = idx 'dummy
     nOpCount = 1 'dummy
     With p.Operators(idx)
      Debug.Assert .Flags >= 0
      'error!?
      Debug.Assert (.Flags And int_OpFlags_Error) = 0
      'add queue?
      If (.Flags And (int_OpFlags_Dirty Or int_OpFlags_InMemory)) = 0 Then
       .Flags = .Flags Or int_OpFlags_Dirty
       objc = objc + 1
       If objc > objm Then
        objm = objm + 256&
        ReDim Preserve objs(1 To objm), obj2(1 To objm)
       End If
       objs(objc) = idx
      End If
     End With
    End If
    'add item
    obj2(objNow).nCount = nOpCount
    If nOpCount > 0 Then
     ReDim obj2(objNow).idxOp(nOpCount - 1)
     CopyMemory obj2(objNow).idxOp(0), nInputOps(0), nOpCount * 4&
    End If
   End With
   'next
   objNow = objNow + 1
  Loop
  'draw it!!!
  'fix the BFS-algorithm bug
  objm = objc
  Do While objm > 0
   idx = 0
   '/////////////////////////////////////
   For objNow = objc To 1 Step -1
    With .Operators(objs(objNow))
     If (.Flags And int_OpFlags_InMemory) = 0 Then 'not in memory
      nLastPage = 1 'bValid
      Select Case .nType
      Case int_OpType_Load, int_OpType_Store, int_OpType_Nop, int_OpType_Export
       Debug.Assert .nBmIndex <= 0
       With p.Operators(obj2(objNow).idxOp(0))
        If (.Flags And int_OpFlags_InMemory) = 0 Then nLastPage = 0 'bValid
        tmp = .nBmIndex
       End With
       If nLastPage Then
        If tmp > 0 Then tmp = -tmp '<0 : this is a reference only
        Debug.Assert tmp < 0 And tmp + bmCount >= 0
        .nBmIndex = tmp
       End If
      Case Else
       'load input
       nOpCount = obj2(objNow).nCount
       For i = 0 To nOpCount - 1
        With p.Operators(obj2(objNow).idxOp(i))
         If (.Flags And int_OpFlags_InMemory) = 0 Then nLastPage = 0 'bValid
         If nLastPage Then 'bValid
          tmp = .nBmIndex
          If tmp < 0 Then tmp = -tmp
          Debug.Assert tmp > 0 And tmp <= bmCount
          Debug.Assert Not bm(tmp) Is Nothing
          With bm(tmp)
           Debug.Assert .Width > 0 And .Height > 0
           tInputBm(i).hdc = .hdc
           tInputBm(i).Width = .Width
           tInputBm(i).Height = .Height
           tInputBm(i).lpbm = .DIBSectionBitsPtr
          End With
         End If
        End With
       Next i
       If nLastPage Then 'bValid
        'allocate memory
        If .nBmIndex <= 0 Or .nBmIndex > bmCount Then
         '.nBmIndex = FindEmptyBitmapEx(p)
         .nBmIndex = FindEmptyBitmap
        End If
        If IsNothing(.nBmIndex) Then Create .nBmIndex
        'show progress
        If Not obj Is Nothing Then
         obj.OnProgress objc - objm + idx + 1, objc, 0, 0, bAbort
         If bAbort Then Exit Do
        End If
        'calc!
        CalcOperator bm(.nBmIndex), nOpCount, tInputBm, .nType, .bProps, .sProps
       End If
      End Select
      'check valid
      If nLastPage Then 'bValid
       .Flags = .Flags Or int_OpFlags_InMemory
       idx = idx + 1
      End If
     End If
    End With
   Next objNow
   '/////////////////////////////////////
   If idx = 0 Then
    'error!!!!!!!
    Debug.Assert False
    Exit Do
   End If
   objm = objm - idx
  Loop
  'get time
  QueryPerformanceCounter t2
  t2 = t2 - t1
  QueryPerformanceFrequency t1
  ShowOperator = t2 * 1000 / t1
 End If
End With
End Function

''///////////////////////////////test only
'Friend Sub CheckTest(p As typeProject)
'Debug.Assert pCheckTest(p)
'End Sub
'
'Private Function pCheckTest(p As typeProject) As Boolean
'Dim i As Long, j As Long, k As Long
'For i = 1 To p.nPageCount
' With p.Pages(i)
'  For j = 0 To int_Page_Height - 1
'   With .Rows(j)
'    For k = 1 To .nOpCount - 1
'     Debug.Assert p.Operators(.idxOp(k)).Left + p.Operators(.idxOp(k)).Width <= p.Operators(.idxOp(k + 1)).Left
'    Next k
'   End With
'  Next j
' End With
'Next i
'pCheckTest = True
'End Function


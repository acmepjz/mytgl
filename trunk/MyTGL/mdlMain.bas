Attribute VB_Name = "mdlMain"
Option Explicit

#Const IsConvert = 0

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
'Private Declare Sub DebugBreak Lib "kernel32.dll" ()

Public Type typeOperatorPropDef
 Name As String
 nType As enumOperatorPropType
 nOffset As Long
 nBitStart As Long
 nBitEnd As Long
 sDefault As String
 sMin As String
 sMax As String
 ListCount As Long
 List() As String '0-based
End Type

Public Type typeOperatorDef
 Name As String
 PropSize As Long
 StringCount As Long
 PropCount As Long
 props() As typeOperatorPropDef '1-based
End Type

'bit0-6: just is an index
'bit7: can't be bitfield
'bit8-15: size
Public Enum enumOperatorPropType
 '--NULL
 eOPT_Name = 0& 'stupid
 eOPT_String = 1&
 eOPT_Group = 2&
 eOPT_Custom = &HFF& 'add!! :-3
 '1-byte
 eOPT_Byte = &H101&
 eOPT_Bool = &H102&
 eOPT_Size = &H103&
 eOPT_ChangeSize = &H104& 'add!!
 '2-byte
 eOPT_Integer = &H201&
 eOPT_Half = &H282&
 eOPT_PtByte = &H283&
 '4-byte
 eOPT_Long = &H401&
 eOPT_Color = &H482&
 eOPT_Single = &H483&
 eOPT_PtInt = &H484&
 eOPT_PtHalf = &H485&
 eOPT_RectByte = &H486&
 '8-byte
 eOPT_Pt = &H881&
 eOPT_PtFloat = &H882&
 eOPT_RectInt = &H883&
 eOPT_RectHalf = &H884&
 '16-byte
 eOPT_Rect = &H1081&
 eOPT_RectFloat = &H1082&
End Enum

#Const Default_Pic_Size = 256

#If Default_Pic_Size = 512 Then
Public Const int_Size_Default As Long = 9 '512
#ElseIf Default_Pic_Size = 256 Then
Public Const int_Size_Default As Long = 8 '256
#Else
Public Const int_Size_Default As Long = 0 / 0 'Unsupported!
#End If
Public Const int_Size_Min As Long = 0
Public Const int_Size_Max As Long = 12

Public tDef() As typeOperatorDef, nOpTypeCount As Long

Public Type typeOperatorProp_DesignTime
 iValue(3) As Long '0-based
 fValue(3) As Single   '0-based
 sValue As String
End Type

Public Type typeOperator_String
 nCount As Long
 bData() As Byte '0-based
End Type

Public Const int_OpFlags_Error As Long = &H1&
Public Const int_OpFlags_InMemory As Long = &H2&
Public Const int_OpFlags_Selected As Long = &H4&
Public Const int_OpFlags_Dirty As Long = &H40000000
Public Const int_OpFlags_Deleted As Long = &H80000000

'Psuedo-Operators
Public Const int_OpType_Load As Long = 251
Public Const int_OpType_Store As Long = 252
Public Const int_OpType_Nop As Long = 253
Public Const int_OpType_Export As Long = 254

Public Type typeOperator_DesignTime
 Name As String
 'nIndex As Long
 nPage As Long
 Left As Long '0-based
 Top As Long '0-based
 Width As Long
 'Height As Long
 'idxNextOp As Long 'linked-list method
 nType As Long
 nBmIndex As Long '1-based bitmap index
 '////temp???
 nBmWidth As Long
 nBmHeight As Long
 '////
 Flags As Long
 bProps() As Byte '0-based
 sProps() As String 'string? 0-based
 'sProps() As typeOperator_String 'byte array?? 0-based
End Type

Public Const int_Generator_Max As Long = 9 'max generator - import

Public Type typeOperatorCalc_DesignTime
' Index As Long
' Flags As Long '??
 nCount As Long
 idxOp() As Long '0-based
End Type

Public Const int_Page_Width As Long = 256&
Public Const int_Page_Height As Long = 128&
Public Const int_Page_WidthPixels As Long = int_Page_Width * 16&
Public Const int_Page_HeightPixels As Long = int_Page_Height * 16&

Public Type typePageRow
 nOpCount As Long
 idxOp() As Long 'array method 1-based
 'idxFirstOp As Long 'linked-list method
End Type

Public Type typeComment
 Left As Long '0-based
 Top As Long '0-based
 Width As Long
 Height As Long
 Color As Long
 Name As String
 Value As String
End Type

Public Type typePage
 Name As String
 'nIndex As Long
 'nOpCount As Long
 nCommentCount As Integer
 nIndent As Byte
 nReserved As Byte
 Rows(int_Page_Height - 1) As typePageRow '0-based
 Comments() As typeComment '1-based
End Type

Public Type typeStoreOp_DesignTime
 Name As String
 Index As Long
End Type

Public Type typeProject
 nPageCount As Long
 nOpCount As Long
 Pages() As typePage '1-based
 Operators() As typeOperator_DesignTime '1-based
End Type

'/////////////////////////////// "HALF" float
'IEEE754

'浮点数表示
'尾数符号位(1bit): 0表示正数；1表示负数。
'所表示的数值=(-1)^尾数符号位*2^(阶码—移码)*((1.尾数)二进制)

'Single(32bit)
'表示: 尾数符号位(1bit)  阶码(8bit)  尾数(23bit)
'阶码范围为1~254=2^8-2。当阶码为0表示实数0；当阶码为255表示越界。
'移码=2^7-1=127。

'Double(64bit):1-11-52

'///////////////////////////////bit-field mask
Private bfMask1(31) As Long '1,2,4,8,...,&H80000000
Private bfMask2(31) As Long '1,3,7,15,...,&HFFFFFFFF

'///////////////////////////////global object
Public cd As New cCommonDialog
Public cSet As New clsSettings

'///////////////////////////////menu :-3
Public Const IDM_ADDCOMMENT As Long = 30001
Public Const IDM_SHOWOP As Long = 30002
Public Const IDM_BRINGTOFRONT As Long = 30003
Public Const IDM_SENDTOBACK As Long = 30004

Public mnu As typeFakeCommandBars

'///////////////////////////////float toolbar
Public m_sAddOpKey() As String, m_sAddOpCaption() As String
Public m_nAddOpTabCount As Long

Private Sub pAddTab(ByVal sKey As String, ByVal sCaption As String)
m_nAddOpTabCount = m_nAddOpTabCount + 1
ReDim Preserve m_sAddOpKey(1 To m_nAddOpTabCount)
ReDim Preserve m_sAddOpCaption(1 To m_nAddOpTabCount)
m_sAddOpKey(m_nAddOpTabCount) = sKey
m_sAddOpCaption(m_nAddOpTabCount) = sCaption
End Sub

'///////////////////////////////entry point

Private Sub Main()
'load manifest
NewLoadManifest
'init menu
FakeCommandBarAddCommandBar mnu, "[P1]"
FakeCommandBarAddCommandBar mnu, "\", "Add Operator", fcbfDragToMakeThisMenuFloat
FakeCommandBarAddButton mnu.d(1), "idx:" + CStr(IDM_ADDCOMMENT), "Add C&omment" + vbTab + "Shift+A"
FakeCommandBarAddButton mnu.d(1), , "&Add Operator" + vbTab + "A", , , , , , , , "\"
pAddTab "\", "Add Operator"
'load
LoadOperationDef
'load settings
cSet.LoadFile CStr(App.Path) + "\MyTGL.cfg"
'init global vars
Dim i As Long, j As Long
bfMask1(0) = 1
bfMask1(31) = &H80000000
bfMask2(30) = &H7FFFFFFF
bfMask2(31) = -1
j = 1
For i = 1 To 30
 j = j + j
 bfMask1(i) = j
 bfMask2(i - 1) = j - 1
Next i
'init color :-3
With cd
 .CustomColor(0) = d_Title1
 .CustomColor(1) = d_Title2
 .CustomColor(2) = d_Bar1
 .CustomColor(3) = d_Bar2
 .CustomColor(4) = d_Hl1
 .CustomColor(5) = d_Hl2
 .CustomColor(6) = d_Checked1
 .CustomColor(7) = d_Checked2
 .CustomColor(8) = d_Pressed1
 .CustomColor(9) = d_Pressed2
End With
'show
Form1.Show
End Sub

'///////////////////////////////
'new:init menu!!
Public Sub LoadOperationDef()
#If IsConvert Then
Dim sConvert As String
Dim sConvert2 As String
Dim sConvertDef As String
Dim sConvertMin As String
Dim sConvertMax As String
Dim sConvertEnum As String
Dim nConvertOffset As Long
#End If
Dim s As String, s2 As String
Dim i As Long, j As Long, k As Long, m As Long
Dim nOffset As Long
Dim bMissing As Boolean
Dim v As Variant
Open CStr(App.Path) + "\Prop.def" For Input As #1
Erase tDef
nOpTypeCount = 0
i = 0
Do Until EOF(1)
 Line Input #1, s
 s = Replace(Trim(s), vbTab, "")
 If Left(s, 2) = "//" Or Left(s, 1) = "'" Or s = "" Then
  'empty
 ElseIf Left(s, 1) = ">" Then
  'menu!
  s = Trim(Mid(s, 2))
  s = Replace(s, "/", "\")
  If Left(s, 1) <> "\" Then s = "\" + s
  pAddMenu s
  #If IsConvert Then
  sConvert = sConvert + "AddOperatorDef , , , , """ + s + """" + vbCrLf
  #End If
 ElseIf Left(s, 1) = "[" Then
  i = Val(Mid(s, 2))
  If i > nOpTypeCount Then
   nOpTypeCount = i
   ReDim Preserve tDef(1 To nOpTypeCount)
  End If
  s = Mid(s, InStr(1, s, "]") + 1)
  v = Split(s, ",")
  m = UBound(v)
  With tDef(i) 'TODO:menu
   'name
   .Name = Trim(v(0))
   'menu
   If m >= 1 Then
    s2 = Trim(v(1))
   Else
    s2 = "\"
   End If
   s2 = Replace(s2, "/", "\")
   #If IsConvert Then
   sConvert = sConvert + "i = AddOperatorDef(""" + .Name + _
   """, , " + CStr(i) + ", 1, """ + s2 + """)" + vbCrLf
   nConvertOffset = 0
   #End If
   If Left(s2, 1) <> "\" Then s2 = "\" + s2
   If Right(s2, 1) <> "\" Then s2 = s2 + "\"
   s2 = s2 + .Name
   If m >= 4 Then
    'TODO:
   End If
   pAddMenu s2, i
   'property size
   If m >= 2 Then
    .PropSize = Val(v(2))
   Else
    .PropSize = 0
   End If
   'string count
   If m >= 3 Then
    .StringCount = Val(v(3))
   Else
    .StringCount = 0
   End If
   '///stupid
   .PropCount = 1
   ReDim .props(1 To 1)
   .props(1).Name = "Name"
  End With
  nOffset = 0
 Else
  v = Split(s, ",")
  m = UBound(v)
  With tDef(i)
   .PropCount = .PropCount + 1
   ReDim Preserve .props(1 To .PropCount)
   With .props(.PropCount)
    #If IsConvert Then
    sConvertDef = ""
    sConvertMin = ""
    sConvertMax = ""
    sConvertEnum = ""
    #End If
    'type
    s = LCase(Trim(v(0)))
    Select Case s
    #If IsConvert Then
    Case "byte", "uchar": .nType = eOPT_Byte: sConvert2 = "AddPropDef i, ""int"", ""%name"", , , 1, , , %def, %min, %max, %enum, %offset"
    Case "bool", "boolean": .nType = eOPT_Bool: sConvert2 = "AddPropDef i, ""boolean"", ""%name"", , , 1, , , %def, , , , %offset"
    Case "integer", "short": .nType = eOPT_Integer: sConvert2 = "AddPropDef i, ""int"", ""%name"", , , 1, , , %def, %min, %max, %enum, %offset"
    Case "half": .nType = eOPT_Half
    Case "long": .nType = eOPT_Long: sConvert2 = "AddPropDef i, ""int"", ""%name"", , , 1, , , %def, %min, %max, %enum, %offset"
    Case "color": .nType = eOPT_Color: sConvert2 = "AddPropDef i, ""color"", ""%name"", , , 4, , , %def, , , , %offset"
    Case "single", "float": .nType = eOPT_Single: sConvert2 = "AddPropDef i, ""float"", ""%name"", , , 1, , , %def, %min, %max, , %offset"
    Case "pointbyte": .nType = eOPT_PtByte: sConvert2 = "AddPropDef i, ""int"", ""%name"", , , 2, , , %def, %min, %max, , %offset"
    Case "pointint": .nType = eOPT_PtInt: sConvert2 = "AddPropDef i, ""int"", ""%name"", , , 2, , , %def, %min, %max, , %offset"
    Case "pointhalf": .nType = eOPT_PtHalf
    Case "point", "pointapi": .nType = eOPT_Pt: sConvert2 = "AddPropDef i, ""int"", ""%name"", , , 2, , , %def, %min, %max, , %offset"
    Case "pointfloat": .nType = eOPT_PtFloat: sConvert2 = "AddPropDef i, ""float"", ""%name"", , , 2, , , %def, %min, %max, , %offset"
    Case "rectbyte": .nType = eOPT_RectByte: sConvert2 = "AddPropDef i, ""int"", ""%name"", , , 4, , , %def, %min, %max, , %offset"
    Case "rectint": .nType = eOPT_RectInt: sConvert2 = "AddPropDef i, ""int"", ""%name"", , , 4, , , %def, %min, %max, , %offset"
    Case "recthalf": .nType = eOPT_RectHalf
    Case "rect": .nType = eOPT_Rect: sConvert2 = "AddPropDef i, ""int"", ""%name"", , , 4, , , %def, %min, %max, , %offset"
    Case "rectfloat": .nType = eOPT_RectFloat: sConvert2 = "AddPropDef i, ""float"", ""%name"", , , 4, , , %def, %min, %max, , %offset"
    Case "size": .nType = eOPT_Size: sConvert2 = "AddPropDef i, ""size"", ""%name"", , , , , , int_Size_Default, , , , %offset"
    Case "size2", "sizeex", "changesize": .nType = eOPT_ChangeSize: sConvert2 = "AddPropDef i, ""resize"", ""%name"", , , , , , , , , , %offset"
    Case "string": .nType = eOPT_String: sConvert2 = "AddPropDef i, ""string"", ""%name"""
    Case "group": .nType = eOPT_Group: sConvert2 = "AddPropDef i, , ""%name"""
    Case "custom": .nType = eOPT_Custom: sConvert2 = "AddPropDef i, ???, ""%name"""
    #Else
    Case "byte", "uchar": .nType = eOPT_Byte
    Case "bool", "boolean": .nType = eOPT_Bool
    Case "integer", "short": .nType = eOPT_Integer
    Case "half": .nType = eOPT_Half
    Case "long": .nType = eOPT_Long
    Case "color": .nType = eOPT_Color
    Case "single", "float": .nType = eOPT_Single
    Case "pointbyte": .nType = eOPT_PtByte
    Case "pointint": .nType = eOPT_PtInt
    Case "pointhalf": .nType = eOPT_PtHalf
    Case "point", "pointapi": .nType = eOPT_Pt
    Case "pointfloat": .nType = eOPT_PtFloat
    Case "rectbyte": .nType = eOPT_RectByte
    Case "rectint": .nType = eOPT_RectInt
    Case "recthalf": .nType = eOPT_RectHalf
    Case "rect": .nType = eOPT_Rect
    Case "rectfloat": .nType = eOPT_RectFloat
    Case "size": .nType = eOPT_Size
    Case "size2", "sizeex", "changesize": .nType = eOPT_ChangeSize
    Case "string": .nType = eOPT_String
    Case "group": .nType = eOPT_Group
    Case "custom": .nType = eOPT_Custom
    #End If
    Case Else
     Debug.Assert False
    End Select
    'name
    .Name = Trim(v(1))
    If .nType < &H100& Then
     'NULL
     If m < 2 Then s = "" Else s = Trim(v(2))
     .nOffset = Val(s)
    Else
     'offset
     bMissing = False
     If m < 2 Then bMissing = True Else s = Trim(v(2)): If s = "" Then bMissing = True
     If Not bMissing Then nOffset = Val(s)
     .nOffset = nOffset
     nOffset = nOffset + (.nType And &HFF00&) \ &H100&
    End If
    'bitfield
    If .nType And &H80& Then
     bMissing = True
    Else
     bMissing = False
     If m < 3 Then bMissing = True Else s = Trim(v(3)): If s = "" Then bMissing = True
    End If
    If bMissing Then
     .nBitStart = -1
     .nBitEnd = -1
    Else
     If m < 4 Then s2 = s Else s2 = Trim(v(4)): If s2 = "" Then s2 = s
     .nBitStart = Val(s)
     .nBitEnd = Val(s2)
    End If
    'check predefined type
    If .nType < &H100& Then
     'NULL
    ElseIf .nType = eOPT_Size Then
     'image size
     .sDefault = CStr(int_Size_Default)
     .sMin = CStr(int_Size_Min)
     .sMax = CStr(int_Size_Max)
     .ListCount = int_Size_Max - int_Size_Min + 1
     ReDim .List(.ListCount - 1)
     k = 1
     For j = 0 To .ListCount - 1
      .List(j) = CStr(k)
      k = k + k
     Next j
    ElseIf .nType = eOPT_ChangeSize Then
     'change size
     .sDefault = CStr(int_Size_Min)
     .sMin = .sDefault
     .sMax = CStr(int_Size_Max + 1)
     .ListCount = int_Size_Max - int_Size_Min + 2
     ReDim .List(.ListCount - 1)
     .List(0) = "(Current)"
     k = 1
     For j = 1 To .ListCount - 1
      .List(j) = CStr(k)
      k = k + k
     Next j
    ElseIf .nType = eOPT_Color Then
     'color
     If m < 5 Then s = "&HFF000000" Else s = Trim(v(5))
     .sDefault = s
     #If IsConvert Then
     sConvertDef = Format(Hex(CLng(Val(s))), "00000000")
     sConvertDef = "Array(&H" + Mid(sConvertDef, 3, 2) + ",&H" + Mid(sConvertDef, 5, 2) + ",&H" + Mid(sConvertDef, 7, 2) + ",&H" + Mid(sConvertDef, 1, 2) + ")"
     #End If
    ElseIf .nType = eOPT_Bool Then
     'default
     If m < 5 Then s = "" Else s = Trim(v(5))
     .sDefault = s
     #If IsConvert Then
     sConvertDef = s
     #End If
     .ListCount = 2
     ReDim .List(1)
     .List(0) = "False"
     .List(1) = "True"
    Else
     'default
     If m < 5 Then s = "" Else s = Trim(v(5))
     .sDefault = s
     #If IsConvert Then
     If InStr(1, s, ";") > 0 Then
      sConvertDef = "Array(" + Replace(s, ";", ",") + ")"
     Else
      sConvertDef = s
     End If
     #End If
     'min
     If m < 6 Then s = "" Else s = Trim(v(6))
     .sMin = s
     #If IsConvert Then
     If InStr(1, s, ";") > 0 Then
      sConvertMin = "Array(" + Replace(s, ";", ",") + ")"
     Else
      sConvertMin = s
     End If
     #End If
     'max
     If m < 7 Then s = "" Else s = Trim(v(7))
     .sMax = s
     #If IsConvert Then
     If InStr(1, s, ";") > 0 Then
      sConvertMax = "Array(" + Replace(s, ";", ",") + ")"
     Else
      sConvertMax = s
     End If
     #End If
     'list
     If m < 8 Then s = "" Else s = Trim(v(8))
     If s <> "" Then
      #If IsConvert Then
      sConvertEnum = "Array(Array("
      #End If
      v = Split(s, ";")
      .ListCount = UBound(v) + 1
      ReDim .List(.ListCount - 1)
      For j = 0 To .ListCount - 1
       .List(j) = Trim(v(j))
       #If IsConvert Then
       If j > 0 Then sConvertEnum = sConvertEnum + ","
       sConvertEnum = sConvertEnum + """" + .List(j) + """"
       #End If
      Next j
      #If IsConvert Then
      sConvertEnum = sConvertEnum + "))"
      #End If
     End If
    End If
    #If IsConvert Then
    sConvert2 = Replace(sConvert2, "%offset", CStr(nConvertOffset))
    Select Case .nType
    Case eOPT_Byte, eOPT_PtByte, eOPT_RectByte
     If sConvertMin = "" Then sConvertMin = "0"
     If sConvertMax = "" Then sConvertMax = "255"
    Case eOPT_Integer, eOPT_PtInt, eOPT_RectInt
     If sConvertMin = "" Then sConvertMin = "-32768"
     If sConvertMax = "" Then sConvertMax = "32767"
    End Select
    Select Case .nType
    Case eOPT_Byte, eOPT_Bool, eOPT_Integer, eOPT_Long, eOPT_Size, eOPT_ChangeSize, eOPT_Single
     nConvertOffset = nConvertOffset + 4
    Case eOPT_PtByte, eOPT_PtInt, eOPT_Pt, eOPT_PtFloat
     nConvertOffset = nConvertOffset + 8
    Case eOPT_RectByte, eOPT_RectInt, eOPT_Rect, eOPT_RectFloat, eOPT_Color
     nConvertOffset = nConvertOffset + 16
    End Select
    sConvert2 = Replace(sConvert2, "%name", .Name)
    sConvert2 = Replace(sConvert2, "%def", sConvertDef)
    sConvert2 = Replace(sConvert2, "%min", sConvertMin)
    sConvert2 = Replace(sConvert2, "%max", sConvertMax)
    sConvert2 = Replace(sConvert2, "%enum", sConvertEnum)
'    Do
'     sConvert2 = Trim(sConvert2)
'     If Right(sConvert2, 1) <> "," Then Exit Do
'     sConvert2 = Left(sConvert2, Len(sConvert2) - 1)
'    Loop
    sConvert = sConvert + sConvert2 + vbCrLf
    #End If
   End With
  End With
 End If
Loop
#If IsConvert Then
Open CStr(App.Path) + "\PropNew.def" For Output As #45
Print #45, sConvert
#End If
Close
End Sub

'menu
Private Sub pAddMenu(ByVal s As String, Optional ByVal wID As Long)
Dim i As Long, idx As Long
If Right(s, 1) = "\" Then
 pAddMenuInternal s
Else
 i = InStrRev(s, "\")
 If i = 0 Then Exit Sub 'error!!
 idx = pAddMenuInternal(Left(s, i))
 If idx = 0 Then Exit Sub 'error!!
 s = Mid(s, i + 1)
 If s = "-" Then
  FakeCommandBarAddButton mnu.d(idx), , , , fbttSeparator
 ElseIf s = "|" Then
  FakeCommandBarAddButton mnu.d(idx), , , , fbttColumnSeparator
 Else
  FakeCommandBarAddButton mnu.d(idx), "idx:" + CStr(wID), s
 End If
End If
End Sub

'menu internal
Private Function pAddMenuInternal(ByVal s As String) As Long
Dim i As Long, m As Long
Dim idx As Long, idx2 As Long
Dim s2 As String
i = FakeCommandBarGetMenuIndex(mnu, LCase(s)) 'mnu.IndexFromKey(LCase(s))
If i = 0 Then
 m = Len(s)
 i = InStrRev(s, "\", m - 1)
 If i = 0 Then Exit Function 'error!!
 'recursive :-3
 idx = pAddMenuInternal(Left(s, i))
 If idx = 0 Then Exit Function 'error!!
 'idx2 = mnu.AddMenu(LCase(s))
 idx2 = FakeCommandBarAddCommandBar(mnu, LCase(s), , fcbfDragToMakeThisMenuFloat)
 s2 = Mid(s, i + 1, m - i - 1)
 pAddTab LCase(s), s2
 'mnu.AddItem idx, , s2, , mnu.hMenu(idx2)
 FakeCommandBarAddButton mnu.d(idx), , s2, , , fbtfShowDropdown, , , , , mnu.d(idx2).sKey
 pAddMenuInternal = idx2
Else
 pAddMenuInternal = i
End If
End Function

Private Function pReadBitFieldUnsigned(ByVal nValue As Long, ByVal nBitStart As Long, ByVal nBitEnd As Long) As Long
Dim i As Long
If 0 <= nBitStart And nBitStart <= nBitEnd And nBitEnd < 32 Then
 If nBitStart = 31 Then
  If nValue And &H80000000 Then i = 1
 ElseIf nBitEnd = 31 Then
  If nBitStart = 0 Then
   i = nValue
  Else
   i = (nValue And (bfMask2(30 - nBitStart) * bfMask1(nBitStart))) \ bfMask1(nBitStart)
   If nValue And &H80000000 Then i = i Or bfMask1(nBitEnd - nBitStart)
  End If
 Else
  i = (nValue And (bfMask2(nBitEnd - nBitStart) * bfMask1(nBitStart))) \ bfMask1(nBitStart)
 End If
End If
pReadBitFieldUnsigned = i
End Function

'Private Function pReadBitFieldSigned(ByVal nValue As Long, ByVal nBitStart As Long, ByVal nBitEnd As Long) As Long
'If 0 <= nBitStart And nBitStart <= nBitEnd And nBitEnd < 32 Then
' 'TODO:
'End If
'End Function

Private Function pWriteBitField(ByVal nOldValue As Long, ByVal nBitStart As Long, ByVal nBitEnd As Long, ByVal nNewValue As Long) As Long
If 0 <= nBitStart And nBitStart <= nBitEnd And nBitEnd < 32 Then
 If nBitStart = 31 Then
  If nNewValue And 1& Then
   nOldValue = nOldValue Or &H80000000
  Else
   nOldValue = nOldValue And &H7FFFFFFF
  End If
 ElseIf nBitEnd = 31 Then
  If nBitStart = 0 Then
   nOldValue = nNewValue
  Else
   nOldValue = (nOldValue And bfMask2(nBitStart - 1)) Or ((nNewValue And bfMask2(30 - nBitStart)) * bfMask1(nBitStart))
   If nNewValue And bfMask1(31 - nBitStart) Then nOldValue = nOldValue Or &H80000000
  End If
 Else
  nOldValue = nOldValue And Not (bfMask2(nBitEnd - nBitStart) * bfMask1(nBitStart))
  nOldValue = nOldValue Or ((nNewValue And bfMask2(nBitEnd - nBitStart)) * bfMask1(nBitStart))
 End If
End If
pWriteBitField = nOldValue
End Function

Private Sub pMovSX(i As Long)
If i And &H8000& Then i = i Or &HFFFF0000
End Sub

Public Sub PropRead(op As typeOperator_DesignTime, d As typeOperatorPropDef, p As typeOperatorProp_DesignTime)
Dim i As Long
With d
 Select Case .nType
 Case eOPT_Name
  p.sValue = op.Name
 Case eOPT_String, eOPT_Custom
  p.sValue = op.sProps(.nOffset)
 Case eOPT_Group
  'do nothing
 Case eOPT_Byte, eOPT_Bool, eOPT_Size, eOPT_ChangeSize
  i = op.bProps(.nOffset)
  If .nBitStart >= 0 And .nBitEnd >= 0 Then
   i = pReadBitFieldUnsigned(i, .nBitStart, .nBitEnd)
  End If
  p.iValue(0) = i
 Case eOPT_Integer
  CopyMemory i, op.bProps(.nOffset), 2&
  pMovSX i
  If .nBitStart >= 0 And .nBitEnd >= 0 Then
   i = pReadBitFieldUnsigned(i, .nBitStart, .nBitEnd)
  End If
  p.iValue(0) = i
 Case eOPT_Long, eOPT_Color
  CopyMemory i, op.bProps(.nOffset), 4&
  If .nBitStart >= 0 And .nBitEnd >= 0 Then
   i = pReadBitFieldUnsigned(i, .nBitStart, .nBitEnd)
  End If
  p.iValue(0) = i
 Case eOPT_Half
  'TODO:
 Case eOPT_PtByte, eOPT_RectByte
  p.iValue(0) = op.bProps(.nOffset)
  p.iValue(1) = op.bProps(.nOffset + 1)
  If .nType = eOPT_RectByte Then
   p.iValue(2) = op.bProps(.nOffset + 2)
   p.iValue(3) = op.bProps(.nOffset + 3)
  End If
 Case eOPT_PtInt, eOPT_RectInt
  CopyMemory i, op.bProps(.nOffset), 2&
  pMovSX i
  p.iValue(0) = i
  CopyMemory i, op.bProps(.nOffset + 2), 2&
  pMovSX i
  p.iValue(1) = i
  If .nType = eOPT_RectInt Then
   CopyMemory i, op.bProps(.nOffset + 4), 2&
   pMovSX i
   p.iValue(2) = i
   CopyMemory i, op.bProps(.nOffset + 6), 2&
   pMovSX i
   p.iValue(3) = i
  End If
 Case eOPT_PtHalf, eOPT_RectHalf
  'TODO:
 Case eOPT_Pt, eOPT_Rect
  CopyMemory p.iValue(0), op.bProps(.nOffset), .nType \ &H100&
 Case eOPT_Single, eOPT_PtFloat, eOPT_RectFloat
  CopyMemory p.fValue(0), op.bProps(.nOffset), .nType \ &H100&
 Case Else
  Debug.Assert False
 End Select
End With
End Sub

'Public Function PropToString(op As typeOperator_DesignTime, p As typeOperatorProp_DesignTime) As String
''TODO:
'End Function

Public Sub PropFromString(ByVal s As String, d As typeOperatorPropDef, p As typeOperatorProp_DesignTime)
Dim i As Long
With d
 Select Case .nType
 Case eOPT_Name, eOPT_String
  p.sValue = s
 Case eOPT_Group
  'do nothing
 Case eOPT_Byte, eOPT_Bool, eOPT_Size, eOPT_ChangeSize, eOPT_Integer, eOPT_Long, eOPT_Color
  p.iValue(0) = Val(s)
 Case eOPT_Half
  'TODO:
 Case eOPT_PtByte, eOPT_PtInt, eOPT_Pt ', eOPT_RectByte
  p.iValue(0) = Val(s)
  i = InStr(1, s, ";") + 1
  p.iValue(1) = Val(Mid(s, i))
 Case eOPT_Single
  p.fValue(0) = Val(s)
 Case eOPT_RectByte, eOPT_RectInt, eOPT_Rect
  p.iValue(0) = Val(s)
  i = InStr(1, s, ";") + 1
  p.iValue(1) = Val(Mid(s, i))
  i = InStr(i, s, ";") + 1
  p.iValue(2) = Val(Mid(s, i))
  i = InStr(i, s, ";") + 1
  p.iValue(3) = Val(Mid(s, i))
 Case eOPT_PtHalf
  'TODO:
 Case eOPT_RectHalf
  'TODO:
 Case eOPT_PtFloat, eOPT_RectFloat
  p.fValue(0) = Val(s)
  i = InStr(1, s, ";") + 1
  p.fValue(1) = Val(Mid(s, i))
  If .nType = eOPT_RectFloat Then
   i = InStr(i, s, ";") + 1
   p.fValue(2) = Val(Mid(s, i))
   i = InStr(i, s, ";") + 1
   p.fValue(3) = Val(Mid(s, i))
  End If
 Case Else
  Debug.Assert False
 End Select
End With
End Sub

Public Sub PropWrite(op As typeOperator_DesignTime, d As typeOperatorPropDef, p As typeOperatorProp_DesignTime)
Dim i As Long
Dim x As Double
With d
 'check range
 If d.sMin <> "" Then
  x = Val(d.sMin)
  For i = 0 To 3
   If p.iValue(i) < x Then p.iValue(i) = x
   If p.fValue(i) < x Then p.fValue(i) = x
  Next i
 End If
 If d.sMax <> "" Then
  x = Val(d.sMax)
  For i = 0 To 3
   If p.iValue(i) > x Then p.iValue(i) = x
   If p.fValue(i) > x Then p.fValue(i) = x
  Next i
 End If
 Select Case .nType
 Case eOPT_Byte, eOPT_PtByte, eOPT_RectByte
  For i = 0 To 3
   If p.iValue(i) < 0 Then p.iValue(i) = 0
   If p.iValue(i) > &HFF& Then p.iValue(i) = &HFF&
  Next i
 Case eOPT_Integer, eOPT_PtInt, eOPT_RectInt
  For i = 0 To 3
   If p.iValue(i) < &HFFFF8000 Then p.iValue(i) = &HFFFF8000
   If p.iValue(i) > &H7FFF& Then p.iValue(i) = &H7FFF&
  Next i
 Case Else
 End Select
 'check list
 If d.ListCount > 0 Then
  If p.iValue(0) < 0 Then p.iValue(0) = 0
  If p.iValue(0) >= d.ListCount Then p.iValue(0) = d.ListCount - 1
 End If
 'check type
 Select Case .nType
 Case eOPT_Name
  op.Name = p.sValue
 Case eOPT_String, eOPT_Custom
  op.sProps(.nOffset) = p.sValue
 Case eOPT_Group
  'do nothing
 Case eOPT_Byte, eOPT_Bool, eOPT_Size, eOPT_ChangeSize
  If .nBitStart >= 0 And .nBitEnd >= 0 Then
   i = pWriteBitField(op.bProps(.nOffset), .nBitStart, .nBitEnd, p.iValue(0))
  Else
   i = p.iValue(0)
  End If
  op.bProps(.nOffset) = i 'And &HFF&
 Case eOPT_Integer
  If .nBitStart >= 0 And .nBitEnd >= 0 Then
   CopyMemory i, op.bProps(.nOffset), 2&
   i = pWriteBitField(i, .nBitStart, .nBitEnd, p.iValue(0))
  Else
   i = p.iValue(0)
  End If
  'p.iValue(0) = i '??
  CopyMemory op.bProps(.nOffset), i, 2&
 Case eOPT_Long, eOPT_Color
  If .nBitStart >= 0 And .nBitEnd >= 0 Then
   CopyMemory i, op.bProps(.nOffset), 4&
   i = pWriteBitField(i, .nBitStart, .nBitEnd, p.iValue(0))
  Else
   i = p.iValue(0)
  End If
  CopyMemory op.bProps(.nOffset), i, 4&
 Case eOPT_Half
  'TODO:
 Case eOPT_PtByte, eOPT_RectByte
  op.bProps(.nOffset) = p.iValue(0) 'And &HFF&
  op.bProps(.nOffset + 1) = p.iValue(1) 'And &HFF&
  If .nType = eOPT_RectByte Then
   op.bProps(.nOffset + 2) = p.iValue(2) 'And &HFF&
   op.bProps(.nOffset + 3) = p.iValue(3) 'And &HFF&
  End If
 Case eOPT_PtInt, eOPT_RectInt
  CopyMemory op.bProps(.nOffset), p.iValue(0), 2&
  CopyMemory op.bProps(.nOffset + 2), p.iValue(1), 2&
  If .nType = eOPT_RectInt Then
   CopyMemory op.bProps(.nOffset + 4), p.iValue(2), 2&
   CopyMemory op.bProps(.nOffset + 6), p.iValue(3), 2&
  End If
 Case eOPT_PtHalf, eOPT_RectHalf
  'TODO:
 Case eOPT_Pt, eOPT_Rect
  CopyMemory op.bProps(.nOffset), p.iValue(0), .nType \ &H100&
 Case eOPT_Single, eOPT_PtFloat, eOPT_RectFloat
  CopyMemory op.bProps(.nOffset), p.fValue(0), .nType \ &H100&
 Case Else
  Debug.Assert False
 End Select
End With
End Sub

Public Function ColorPicker(Optional ByVal clr As Long = &HFF000000) As Long
Dim frm As frmColorPicker
Set frm = New frmColorPicker
frm.ClampBorder = True
frm.UseInteger = True
frm.SetColorData (clr And &HFF00FF00) Or ((clr And &HFF&) * &H10000) Or ((clr And &HFF0000) \ &H10000)
Load frm
frm.Show 1
clr = frm.GetColorData
ColorPicker = (clr And &HFF00FF00) Or ((clr And &HFF&) * &H10000) Or ((clr And &HFF0000) \ &H10000)
Set frm = Nothing
End Function

Public Function StringPicker(s As String, p As typeProject, cObj As clsOperators, Optional ByVal IsList As Boolean) As Boolean
Dim sto() As typeStoreOp_DesignTime, m As Long
Dim frm As frmStr
Dim i As Long
Set frm = New frmStr
Load frm
If IsList Then
 m = cObj.GetStoreObjects(p, sto)
 With frm.lstStr
  .Visible = True
  For i = 1 To m
   .AddItem sto(i).Name
  Next i
  For i = 1 To m
   If sto(i).Name = s Then
    .ListIndex = i - 1
    Exit For
   End If
  Next i
 End With
Else
 With frm.txtStr
  .Visible = True
  .Text = s
 End With
End If
frm.oString = s
frm.Show 1
If s <> frm.oString Then
 s = frm.oString
 StringPicker = True
End If
Set frm = Nothing
End Function

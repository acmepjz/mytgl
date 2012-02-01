Attribute VB_Name = "mdlCompiler"
Option Explicit

Public AppPath As String
Public SrcFile As String

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
Public Const int_Size_Default As Long = &H80000000 Mod &HFFFFFFFF 'Unsupported!
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

Public Sub LoadOperationDef()
Dim s As String, s2 As String
Dim i As Long, j As Long, k As Long, m As Long
Dim nOffset As Long
Dim bMissing As Boolean
Dim v As Variant
Open AppPath + "Prop.def" For Input As #1
Erase tDef
nOpTypeCount = 0
i = 0
Do Until EOF(1)
 Line Input #1, s
 s = Replace(Trim(s), vbTab, "")
 If Left(s, 2) = "//" Or Left(s, 1) = "'" Or s = "" Then
  'empty
 ElseIf Left(s, 1) = ">" Then
  'menu! do nothing
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
   If Left(s2, 1) <> "\" Then s2 = "\" + s2
   If Right(s2, 1) <> "\" Then s2 = s2 + "\"
   s2 = s2 + .Name
   If m >= 4 Then
    'TODO:
   End If
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
    'type
    s = LCase(Trim(v(0)))
    Select Case s
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
    ElseIf .nType = eOPT_Bool Then
     'default
     If m < 5 Then s = "" Else s = Trim(v(5))
     .sDefault = s
     .ListCount = 2
     ReDim .List(1)
     .List(0) = "False"
     .List(1) = "True"
    Else
     'default
     If m < 5 Then s = "" Else s = Trim(v(5))
     .sDefault = s
     'min
     If m < 6 Then s = "" Else s = Trim(v(6))
     .sMin = s
     'max
     If m < 7 Then s = "" Else s = Trim(v(7))
     .sMax = s
     'list
     If m < 8 Then s = "" Else s = Trim(v(8))
     If s <> "" Then
      v = Split(s, ";")
      .ListCount = UBound(v) + 1
      ReDim .List(.ListCount - 1)
      For j = 0 To .ListCount - 1
       .List(j) = Trim(v(j))
      Next j
     End If
    End If
   End With
  End With
 End If
Loop
Close
End Sub

'workaround for stupid VB collection :-3
Public Function StringToHex(ByVal s As String) As String
Dim i As Long
For i = 1 To Len(s)
 StringToHex = StringToHex + Right("000" + Hex(AscW(Mid(s, i, 1)) And &HFFFF&), 4)
Next i
End Function

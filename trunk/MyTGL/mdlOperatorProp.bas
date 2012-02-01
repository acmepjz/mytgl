Attribute VB_Name = "mdlOperatorProp"
Option Explicit

'////////////////////////////////
'This file is part of MyTGL, an opensource procedural media creation tool and library.
'Copyright (C) 2008,2009  acme_pjz
'
'This program is free software: you can redistribute it and/or modify
'it under the terms of the GNU General Public License as published by
'the Free Software Foundation, either version 3 of the License, or
'(at your option) any later version.
'
'This program is distributed in the hope that it will be useful,
'but WITHOUT ANY WARRANTY; without even the implied warranty of
'MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the
'GNU General Public License for more details.
'
'You should have received a copy of the GNU General Public License
'along with this program.  If not, see <http://www.gnu.org/licenses/>.
'////////////////////////////////

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

'////////////////////////////////////////////////////////////////data definitions

'color format=A32G32B32R32F

Public Type typeMyTGL11OperatorPropEnumItem
 sKey As String '???
 sCaption As String
 nValue As Long
End Type

Public Type typeMyTGL11OperatorPropEnum
 nCount As Long
 d() As typeMyTGL11OperatorPropEnumItem '1-based
End Type

Public Type typeMyTGL11OperatorPropDef
 sKey As String 'key(name)
 sCaption As String 'caption(display name)
 nType As Long 'index to OperatorPropTypeDef
 '&H80000000 - &H800000FE = group
 nFlags As Long
 nElementCount As Long 'count, float4 ...
 nSize As Long 'default size (0=sizable? >0=fixed?)
 nMinSize As Long '??? min size ??? unused
 datDefault() As Byte '0-based
 datMin() As Byte '0-based
 datMax() As Byte '0-based
 nEnumCount As Long
 datEnum() As typeMyTGL11OperatorPropEnum '0-based
 '////////calcuate-time properties
 nOffset As Long 'valid when default size>0
 '////////???????? TODO
End Type

Public Type typeMyTGL11OperatorPropDefs
 nCount As Long
 d() As typeMyTGL11OperatorPropDef '1-based
End Type

Public Type typeMyTGL11OperatorPropTypeDef
 sKey As String 'key(name)????
 nSize As Long 'default size per element
 BasicDataType As Long
 '0-other
 '1-integer
 '2-float
 nFlags As Long
 '1-customizable
 '2-can't reset
 '////////????????Callback????????
 objCallback As IOperatorPropCallback
End Type

Public Type typeMyTGL11OperatorDef
 Key As String
 Name As String
 nType As Long
 nClass As Long
 nFlags As Long
 '1-have custom prev node
 '2-have custom next node
 nPackedDataSize As Long '?
 props As typeMyTGL11OperatorPropDefs
End Type

Public g_PropTypeDefs() As typeMyTGL11OperatorPropTypeDef
Public g_PropTypeDefCount As Long
Public g_OpDefs() As typeMyTGL11OperatorDef
Public g_OpDefCount As Long

'////////////////////////////////////////////////////////////////documents data

Public Type typeMyTGL11OperatorProperty
 sKey As String
 BasicDataType As Long
 nElementCount As Long 'stored element count
 nSize As Long 'stored size in bytes
 '///new!!!
 nIndex As Long 'index to OperatorDef.props (don't save to file)
 '///
 d() As Byte '0-based
End Type

Public Type typeMyTGL11OperatorProperties
 nCount As Long
 d() As typeMyTGL11OperatorProperty '1-based
End Type

Public Sub AddDefaultProp(props As typeMyTGL11OperatorProperties, tDef As typeMyTGL11OperatorPropDef, ByVal nPropDefIndex As Long)
Dim j As Long, k As Long
With tDef
 If .nType > 0 Then
  j = props.nCount + 1
  props.nCount = j
  ReDim Preserve props.d(1 To j)
  props.d(j).sKey = .sKey
  props.d(j).BasicDataType = g_PropTypeDefs(.nType).BasicDataType
  props.d(j).nElementCount = .nElementCount
  props.d(j).nSize = .nSize
  props.d(j).nIndex = nPropDefIndex
  If .nSize > 0 Then
   ReDim props.d(j).d(.nSize - 1)
   On Error Resume Next
   Err.Clear
   k = 0
   k = UBound(.datDefault) + 1
   If Err.Number Then k = 0
   On Error GoTo 0
   If k > .nSize Then k = .nSize
   If k > 0 Then
    CopyMemory props.d(j).d(0), .datDefault(0), k
   End If
  Else
   On Error Resume Next
   Err.Clear
   k = 0
   k = UBound(.datDefault) + 1
   If Err.Number Then k = 0
   On Error GoTo 0
   If k > 0 Then
    props.d(j).nSize = k
    ReDim props.d(j).d(k - 1)
    CopyMemory props.d(j).d(0), .datDefault(0), k
   End If
  End If
 End If
End With
End Sub

'convert data load from file to default size
Public Sub ConvertDataType(d As typeMyTGL11OperatorProperties, tDef As typeMyTGL11OperatorPropDefs)
Dim i As Long, j As Long, k As Long
Dim lp As Long, lp2 As Long, x As Long, f As Single
Dim nSize2 As Long 'default size per element
Dim nDestType As Long
Dim idx As Long
Dim b() As Byte
'for each definition
For j = 1 To tDef.nCount
 'find prop
 For i = 1 To d.nCount
  If d.d(i).sKey = tDef.d(j).sKey Then
   With d.d(i)
    '///????????
    .nIndex = j
    '///
    idx = tDef.d(j).nType
    If idx > 0 Then
     nDestType = g_PropTypeDefs(idx).BasicDataType
     If nDestType = 0 Then
      'TODO:convert using callback
     Else
      nSize2 = tDef.d(j).nSize
      ReDim b(nSize2 - 1)
      nSize2 = nSize2 \ tDef.d(j).nElementCount
      'copy default data
      On Error Resume Next
      Err.Clear
      k = UBound(tDef.d(j).datDefault) + 1
      If Err.Number Then k = 0
      On Error GoTo 0
      If k > 0 Then CopyMemory b(0), tDef.d(j).datDefault(0), k
      'if source type can convert
      If .BasicDataType > 0 Then
       .nSize = .nSize \ .nElementCount
       If .nSize = 0 Then .nElementCount = 0
       lp = 0
       lp2 = 0
       For k = 0 To tDef.d(j).nElementCount - 1
        If k >= .nElementCount Then Exit For
        'read
        Select Case .BasicDataType
        Case 1 'int
         x = 0
         CopyMemory x, .d(lp), .nSize
         If .nSize = 2 Then
          If x And &H8000& Then x = x Or &HFFFF0000
         ElseIf .nSize = 3 Then
          If x And &H800000 Then x = x Or &HFF000000
         End If
         f = x
         lp = lp + .nSize
        Case 2 'float
         f = 0
         CopyMemory f, .d(lp), .nSize
         On Error Resume Next
         x = f
         On Error GoTo 0
         lp = lp + .nSize
        End Select
        'write
        Select Case nDestType
        Case 1 'int
         CopyMemory b(lp2), x, nSize2
         lp2 = lp2 + nSize2
        Case 2 'float
         CopyMemory b(lp2), f, nSize2
         lp2 = lp2 + nSize2
        End Select
       Next k
      End If
      'save data
      .BasicDataType = nDestType
      .nElementCount = tDef.d(j).nElementCount
      .nSize = tDef.d(j).nSize
      .d = b
     End If
    End If
   End With
   Exit For
  End If
 Next i
 'if not found
 If i > d.nCount Then
  'create an item with default value
  AddDefaultProp d, tDef.d(j), j
 End If
Next j
End Sub

Public Function FindProperty(props As typeMyTGL11OperatorProperties, ByVal sKey As String) As Long
Dim i As Long
For i = 1 To props.nCount
 If props.d(i).sKey = sKey Then
  FindProperty = i
  Exit Function
 End If
Next i
End Function

Public Sub PackProperty(props As typeMyTGL11OperatorProperties, propRet As typeMyTGL11OperatorProperties, tDef As typeMyTGL11OperatorDef)
Dim i As Long, m As Long
Dim nIndex As Long
With propRet
 .nCount = 1
 ReDim .d(1 To 1)
 If tDef.nPackedDataSize > 0 Then 'main(anonymous) data area
  .d(1).nSize = tDef.nPackedDataSize
  ReDim .d(1).d(tDef.nPackedDataSize - 1)
 End If
End With
With props
 For i = 1 To .nCount
  nIndex = .d(i).nIndex
  If nIndex > 0 Then
   m = tDef.props.d(nIndex).nSize
   If m > 0 Then 'fixed length
    If .d(i).nSize < m Then m = .d(i).nSize
    If m > 0 Then
     CopyMemory propRet.d(1).d(tDef.props.d(nIndex).nOffset), .d(i).d(0), m
    End If
   Else
    propRet.nCount = propRet.nCount + 1
    ReDim Preserve propRet.d(1 To propRet.nCount)
    propRet.d(propRet.nCount) = .d(i)
   End If
  End If
 Next i
End With
End Sub

Attribute VB_Name = "mdlFakeButtonCommon"
Option Explicit

'////////////////////////////////
'This file is public domain.
'////////////////////////////////

Public Type typeFakeButton
 nType As Byte
 '0=normal
 '1=separator
 '2=check
 '3=option
 '4=optionnull
 '5=split
 '6=v-separator (column)
 Value As Byte 'checked
 nFlags As Integer
 '1=hidden
 '2=disabled
 '4=show dropdown
 '8=owner mesaure
 '16=owner draw(before)
 '32=owner draw(after)
 '64=default item
 '128=hide caption
 '256=start new row
 '512=a full row in toolbar menu mode
 '1024=PicLeft is color!
 '2048=don't change text color if disabled
 'etc.
 GroupIndex As Integer
 '//////////////////////internal
 Width As Integer
 Left As Long
 'Top As Long
 'Height As Long
 mnuLeft As Long
 mnuTop As Long
 mnuWidth As Long '&H80000000=invisible -1=sizable
 mnuWidth2 As Long
 mnuHeight As Long
 '//////////////////////
 PicLeft As Long 'or color
 s As String 'caption
 s2 As String 'tooltiptext
 sTab As String 'shortcut
 sDesc As String 'description
 sSubMenu As String 'sub menu name
 sKey As String
End Type

Public Type typeFakeCommandBar
 nFlags As Long
 'TODO:
 '1=drag to make this menu float
 '2=only icons (?)
 '4=...
 sKey As String
 sCaption As String 'caption
 nCount As Long
 d() As typeFakeButton
 '//////////////////////internal
 nFlags2 As Long
 '1=dirty
 '2=empty menu
 w As Long
 h As Long
End Type

Public Type typeFakeCommandBars
 nCount As Long
 d() As typeFakeCommandBar
End Type

'fix the bug:two control can popup menu simultaneously
Public g_FakeMenuUserData As Long

Public Sub FakeCommandBarApplyToolBar(obj As FakeToolBar, d As typeFakeCommandBar)
obj.CreateIndirect d
End Sub

Public Sub FakeCommandBarApplyMenu(obj As FakeMenu, d As typeFakeCommandBars)
obj.CreateIndirect d
End Sub

Public Function FakeCommandBarAddCommandBar(d As typeFakeCommandBars, ByVal Key As String, Optional ByVal Caption As String, Optional ByVal nFlags As enumFakeCommandBarFlags) As Long
With d
 .nCount = .nCount + 1
 ReDim Preserve .d(1 To .nCount)
 With .d(.nCount)
  .sKey = Key
  .sCaption = Caption
  .nFlags = nFlags
  .nFlags2 = 1
 End With
 FakeCommandBarAddCommandBar = .nCount
End With
End Function

Public Function FakeCommandBarGetMenuIndex(d As typeFakeCommandBars, ByVal Key As String) As Long
Dim i As Long
For i = 1 To d.nCount
 If Key = d.d(i).sKey Then
  FakeCommandBarGetMenuIndex = i
  Exit For
 End If
Next i
End Function

Public Function FakeCommandBarAddCommandBarIndirect(d As typeFakeCommandBars, dSrc As typeFakeCommandBar) As Long
With d
 .nCount = .nCount + 1
 ReDim Preserve .d(1 To .nCount)
 .d(.nCount) = dSrc
 .d(.nCount).nFlags2 = 1
 FakeCommandBarAddCommandBarIndirect = .nCount
End With
End Function

Public Function FakeCommandBarAddButton(d As typeFakeCommandBar, Optional ByVal Key As String, Optional ByVal Caption As String, Optional ByVal ToolTipText As String, Optional ByVal nType As enumFakeButtonType, Optional ByVal nFlags As enumFakeButtonFlags, _
Optional ByVal GroupIndex As Long, Optional ByVal PicLeft As Long = -1, Optional ByVal Caption2 As String, Optional ByVal Description As String, Optional ByVal SubMenuKey As String, Optional ByVal Checked As Boolean) As Long
Dim i As Long
i = InStr(1, Caption, vbTab)
If i > 0 Then
 If Caption2 = "" Then
  Caption2 = Mid(Caption, i + 1)
  Caption = Left(Caption, i - 1)
 End If
End If
With d
 .nFlags2 = 1
 .nCount = .nCount + 1
 ReDim Preserve .d(1 To .nCount)
 With d.d(.nCount)
  .sKey = Key
  .s = Caption
  .s2 = ToolTipText
  .sTab = Caption2
  .sDesc = Description
  .sSubMenu = SubMenuKey
  .nType = nType
  .nFlags = nFlags
  .Value = Checked And 1&
  .GroupIndex = GroupIndex
  .PicLeft = PicLeft
 End With
 FakeCommandBarAddButton = .nCount
End With
End Function

Public Function FakeCommandBarAddButtonIndirect(d As typeFakeCommandBar, dSrc As typeFakeButton) As Long
With d
 .nCount = .nCount + 1
 ReDim Preserve .d(1 To .nCount)
 .d(.nCount) = dSrc
 FakeCommandBarAddButtonIndirect = .nCount
End With
End Function

Public Sub FakeCommandBarFromString(theStr As String, d As typeFakeCommandBar, Optional ByVal PicSize As Long = 16)
Dim v1 As Variant, v2 As Variant, v3 As Variant
Dim s As String
Dim m As Long, m2 As Long
Dim i As Long, j As Long
v1 = Split(theStr + " ", ",")
d.nFlags2 = 1
If theStr = "" Then d.nCount = 0 Else d.nCount = UBound(v1) + 1
If d.nCount <= 0 Then
 Erase d.d
 Exit Sub
End If
ReDim d.d(1 To d.nCount)
For i = 1 To d.nCount
 s = v1(i - 1)
 v2 = Split(s + " ", ";")
 m = UBound(v2)
 With d.d(i)
  'get picture
  If Trim(s) = "" Then 'separator
   .nType = 1
  Else
   s = Trim(v2(0))
   .PicLeft = (Val(s) - 1) * PicSize
   If m >= 1 Then .s = Trim(v2(1)) 'caption
   If m >= 2 Then .s2 = Trim(v2(2)) 'tooltiptext
   If m >= 3 Then
    v3 = Split(LCase(v2(3)) + " ", ":")
    m2 = UBound(v3)
    If m2 >= 1 Then j = Val(v3(1)) Else j = 0
    Select Case Trim(v3(0))
    Case "separator", "seperator"
     .nType = 1
    Case "check", "checkbox"
     .nType = 2
    Case "option", "optionbutton", "radio", "radiobutton"
     .nType = 3
     .GroupIndex = j
    Case "optionnullable", "optionnull", "optnull"
     .nType = 4
     .GroupIndex = j
    Case "split", "spliter", "splitter"
     .nType = 5
    Case "vseparator", "v-separator", "vseperator", "v-seperator", "column", "col"
     .nType = 6
    End Select
    For j = 0 To m2
     Select Case Trim(v3(j))
     Case "checked", "selected", "true"
      .Value = 1
     Case "hide", "hidden", "invisible"
      .nFlags = .nFlags Or 1&
     Case "disable", "disabled"
      .nFlags = .nFlags Or 2&
     Case "drop", "dropdown"
      .nFlags = .nFlags Or 4&
     Case "def", "default"
      .nFlags = .nFlags Or 64&
     Case "hidec", "hidecap", "hidecaption"
      .nFlags = .nFlags Or 128&
     Case "newrow", "startnewrow"
      .nFlags = .nFlags Or 256&
     Case "fullrow"
      .nFlags = .nFlags Or 512&
     Case "color"
      .nFlags = .nFlags Or 1024&
     End Select
    Next j
   End If
   If m >= 4 Then .sTab = Trim(v2(4))
   If m >= 5 Then .sDesc = Trim(v2(5))
   If m >= 6 Then .sSubMenu = Trim(v2(6))
   If m >= 7 Then .sKey = Trim(v2(7))
  End If
 End With
Next i
End Sub

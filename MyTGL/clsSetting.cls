VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSettings"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////
'This file is public domain.
'////////////////////////////////

Private fn As String

Private Type typeSettingData
 Key As String
 Value As String
 en As Boolean
End Type

Private dat() As typeSettingData
Private datc As Long

Public Property Get Key(ByVal Index As Long) As String
Key = dat(Index).Key
End Property

Public Property Let Key(ByVal Index As Long, ByVal s As String)
dat(Index).Key = s
End Property

Public Property Get Value(ByVal Index As Long) As String
Value = dat(Index).Value
End Property

Public Property Let Value(ByVal Index As Long, ByVal s As String)
dat(Index).Value = s
End Property

Public Property Get IsExist(ByVal Index As Long) As Boolean
IsExist = dat(Index).en
End Property

Public Property Let IsExist(ByVal Index As Long, ByVal b As Boolean)
dat(Index).en = b
End Property

Public Property Get FileName() As String
FileName = fn
End Property

Public Property Let FileName(ByVal s As String)
fn = s
End Property

Public Property Get Count() As Long
Count = datc
End Property

Public Sub Clear()
Erase dat
datc = 0
End Sub

Public Sub Remove(ByVal Key As String)
Dim i As Long
For i = 1 To datc
 If dat(i).Key = Key Then
  dat(i).en = False
  Exit Sub
 End If
Next i
End Sub

Public Function GetSettings(ByVal Key As String, Optional ByVal sDefault As String) As String
Dim i As Long
For i = 1 To datc
 If dat(i).Key = Key And dat(i).en Then
  GetSettings = dat(i).Value
  Exit Function
 End If
Next i
GetSettings = sDefault
End Function

Public Sub SetSettings(ByVal Key As String, ByVal Value As String)
Dim i As Long
For i = 1 To datc
 If dat(i).Key = Key And dat(i).en Then
  dat(i).Value = Value
  Exit Sub
 End If
Next i
For i = 1 To datc
 If Not dat(i).en Then
  dat(i).Key = Key
  dat(i).Value = Value
  dat(i).en = True
  Exit Sub
 End If
Next i
datc = datc + 1
ReDim Preserve dat(1 To datc)
With dat(datc)
 .Key = Key
 .Value = Value
 .en = True
End With
End Sub

Public Sub LoadFile(Optional ByVal FileName As String)
On Error GoTo a
Dim i As Long, m As Integer
Dim lp As Long
If FileName <> "" Then fn = FileName
Open fn For Binary As #1
Get #1, 1, m
datc = m
ReDim dat(1 To m)
lp = 3
For i = 1 To m
 lp = lp + LoadStringA(lp, dat(i).Key)
 lp = lp + LoadStringA(lp, dat(i).Value)
 dat(i).en = True
Next i
Close 1
Exit Sub
a:
Close 1
Erase dat
datc = 0
End Sub

Public Sub SaveFile(Optional ByVal FileName As String)
On Error GoTo a
Dim i As Long, m As Integer
Dim lp As Long
If FileName <> "" Then fn = FileName
Open fn For Output As #1
Close 1
Open fn For Binary As #1
lp = 3
For i = 1 To datc
 If dat(i).en Then
  lp = lp + SaveStringA(lp, dat(i).Key)
  lp = lp + SaveStringA(lp, dat(i).Value)
  m = m + 1
 End If
Next i
Put #1, 1, m
a:
Close 1
End Sub

Private Function SaveStringA(ByVal lp As Long, ByVal s As String) As Long
Dim i As Long, m As Integer, x As Integer
m = Len(s)
SaveStringA = m * 2 + 2
Put #1, lp, m
For i = 1 To m
 x = Asc(Mid(s, i, 1))
 m = m Xor x
 Put #1, lp + 2 * i, m
Next i
End Function

Private Function LoadStringA(ByVal lp As Long, s As String) As Long
Dim i As Long, m As Integer, x As Integer
s = ""
Get #1, lp, m
LoadStringA = m * 2 + 2
For i = 1 To m
 Get #1, lp + 2 * i, x
 m = m Xor x
 s = s + Chr(m)
 m = x
Next i
End Function

VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsUnknown"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)

Private bInit As Boolean
Private bCode() As Byte
Private nEntry() As Long, nEntryCount As Long

'/////////////////////////illegal method !!!!!

Private Function Sub0001(ByVal a As Long, ByVal b As Long) As Long
'shl
End Function

Private Function Sub0002(ByVal a As Long, ByVal b As Long) As Long
'shr
End Function

Private Function Sub0003(ByRef a As Long) As Long
'rand
End Function

Private Function Sub0004(ByVal x As Long, ByVal y As Long, ByVal TheSeed As Long) As Long
'rand2
End Function

Private Function Sub0005(ByVal x As Long, ByVal y As Long, ByVal z As Long, ByVal TheSeed As Long) As Long
'rand3
End Function

'/////////////////////////friend methods
Friend Function fShl(ByVal a As Long, ByVal b As Long) As Long
fShl = Sub0001(a, b)
End Function

Friend Function fShr(ByVal a As Long, ByVal b As Long) As Long
fShr = Sub0002(a, b)
End Function

Friend Function fRnd(ByRef a As Long) As Long
fRnd = Sub0003(a)
End Function

Friend Function fRnd2(ByVal x As Long, ByVal y As Long, ByVal TheSeed As Long) As Long
fRnd2 = Sub0004(x, y, TheSeed)
End Function

Friend Function fRnd3(ByVal x As Long, ByVal y As Long, ByVal z As Long, ByVal TheSeed As Long) As Long
fRnd3 = Sub0005(x, y, z, TheSeed)
End Function

'-1 to 1
Friend Function fRndFloat(ByRef a As Long) As Double
fRndFloat = 1# - Sub0003(a) / 16384#
End Function

'-1 to 1
'not so random!!!!
Friend Function fRnd2Float(ByVal x As Long, ByVal y As Long, ByVal TheSeed As Long) As Double
fRnd2Float = 1# - (Sub0004(x, y, TheSeed) And &H7FFF&) / 16384#
End Function

'-1 to 1
Friend Function fRnd3Float(ByVal x As Long, ByVal y As Long, ByVal z As Long, ByVal TheSeed As Long) As Double
fRnd3Float = 1# - (Sub0005(x, y, z, TheSeed) And &H7FFF&) / 16384#
End Function

'/////////////////////////initalize
Friend Sub InitASM()
Dim s As String
Dim lps As Long, lpe As Long, lpm As Long
Dim m As Long, i As Long, j As Long
If bInit Then Exit Sub
'init code
s _
= "$1$8B4424088B4C240CD3E08B4C24108901C21000" _
+ "$2$8B4424088B4C240CD3E88B4C24108901C21000" _
+ "$3$8B4C24088B0169C0354E5A014089018B4C240CC1F81025FF7F00008901C20C008BC1C3" _
+ "$4$558BEC8B4D108B45148B550C69C943AAB06869C0B977575A69D2D166E3712BC12BC2051F0731018D0CC5000000002BC88D0C49C1E1032BC88D0C498D14488BC80FAFC8C1FA078D0C49D1E133D133C2894510C14D10078B55108B451881E2FFFFFF7F89105DC21400" _
+ "$5$558BEC8B4D148B451069C9C962F02969C0BD554F978B55182BC18B4D0C69D2B977575A69C9D166E37103C22BC1051F0731018D0C808D0C898D14898D0C908D14488BC80FAFC8C1FA078D0C49D1E133D133C2894514C14D14078B55148B451C81E2FFFFFF7F89105DC21800"
'/////////////////////////////read code
Erase nEntry
nEntryCount = 0
ReDim bCode(63)
m = 64
s = Replace(s, " ", "")
lps = 1
lpm = Len(s)
Do
 If lps = 0 Or lps > lpm Then Exit Do
 If Mid(s, lps, 1) = "$" Then 'define entry
  lpe = InStr(lps + 1, s, "$")
  If lpe = 0 Then Exit Do
  j = Val(Mid(s, lps + 1, lpe - lps - 1))
  If j > nEntryCount Then
   nEntryCount = j
   ReDim Preserve nEntry(1 To nEntryCount)
  End If
  nEntry(j) = i
  lps = lpe + 1
 Else
  If i >= m Then
   m = m + 64
   ReDim Preserve bCode(m - 1)
  End If
  bCode(i) = Val("&H" + Mid(s, lps, 2))
  lps = lps + 2
  i = i + 1
 End If
Loop
'/////////////////////////////patch it !!!!
CopyMemory lps, ByVal ObjPtr(Me), 4& 'get VTable address
lpe = VarPtr(bCode(0))
For i = 1 To nEntryCount
 nEntry(i) = lpe + nEntry(i)
Next i
'there are 7 hidden functions
CopyMemory ByVal (lps + 28&), nEntry(1), nEntryCount * 4&
'/////////////////////////////over!!!!
bInit = True
End Sub

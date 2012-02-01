VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCPU"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////
'This file is public domain.
'////////////////////////////////

Private Declare Function GetProcessTimes Lib "kernel32.dll" (ByVal hProcess As Long, ByRef lpCreationTime As Currency, ByRef lpExitTime As Currency, ByRef lpKernelTime As Currency, ByRef lpUserTime As Currency) As Long
Private Declare Function GetCurrentProcess Lib "kernel32.dll" () As Long

Private Declare Function GetTickCount Lib "kernel32.dll" () As Long

Private t1 As Currency, t2 As Currency, t3 As Currency, t4 As Currency
Private t0 As Long

Public Function GetCPUUsage() As Single
Dim tt3 As Currency, tt4 As Currency
Dim tt0 As Long
tt0 = t0
tt3 = t3
tt4 = t4
t0 = GetTickCount
GetProcessTimes GetCurrentProcess, t1, t2, t3, t4
tt0 = t0 - tt0
tt3 = t3 - tt3
tt4 = t4 - tt4
If t0 > 0 Then GetCPUUsage = (tt3 + tt4) / tt0
End Function

Private Sub Class_Initialize()
t0 = GetTickCount
GetProcessTimes GetCurrentProcess, t1, t2, t3, t4
End Sub

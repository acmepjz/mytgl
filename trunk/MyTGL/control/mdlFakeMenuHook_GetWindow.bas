Attribute VB_Name = "mdlFakeMenuHook_GetWindow"
Option Explicit


'////////////////////////////////API Hook code

'code

'$ ==>    >    837C24 08 04  CMP DWORD PTR SS:[ESP+8],4
'$+5      >    75 1A         JNZ SHORT 工程Hook.0040147F
'$+7      >    B8 EFBEADDE   MOV EAX,DEADBEEF                         ;  array pointer
'$+C      >    85C0          TEST EAX,EAX
'$+E      >    74 11         JE SHORT 工程Hook.0040147F
'$+10     >    8B5424 04     MOV EDX,DWORD PTR SS:[ESP+4]
'$+14     >    8B08          MOV ECX,DWORD PTR DS:[EAX]
'$+16     >    E3 09         JECXZ SHORT 工程Hook.0040147F
'$+18     >    3BCA          CMP ECX,EDX
'$+1A     >    74 0F         JE SHORT 工程Hook.00401489
'$+1C     >    83C0 08       ADD EAX,8
'$+1F     >  ^ EB F3         JMP SHORT 工程Hook.00401472
'$+21     >    8BFF          MOV EDI,EDI
'$+23     >    55            PUSH EBP
'$+24     >    8BEC          MOV EBP,ESP
'$+26     >  - E9 66AA6DDE   JMP DEADBEEF                             ;  old func
'$+2B     >    8B40 04       MOV EAX,DWORD PTR DS:[EAX+4]
'$+2E     >    C2 0800       RETN 8

Private Declare Function GetProcAddress Lib "kernel32.dll" (ByVal hModule As Long, ByVal lpProcName As String) As Long
Private Declare Function GetModuleHandle Lib "kernel32.dll" Alias "GetModuleHandleA" (ByVal lpModuleName As String) As Long
Private Declare Function VirtualProtect Lib "kernel32.dll" (ByRef lpAddress As Any, ByVal dwSize As Long, ByVal flNewProtect As Long, ByRef lpflOldProtect As Long) As Long
Private Declare Function VirtualQuery Lib "kernel32.dll" (ByRef lpAddress As Any, ByRef lpBuffer As MEMORY_BASIC_INFORMATION, ByVal dwLength As Long) As Long
Private Declare Sub CopyMemory Lib "kernel32.dll" Alias "RtlMoveMemory" (ByRef Destination As Any, ByRef Source As Any, ByVal Length As Long)
Private Type MEMORY_BASIC_INFORMATION
    BaseAddress As Long
    AllocationBase As Long
    AllocationProtect As Long
    RegionSize As Long
    State As Long
    Protect As Long
    lType As Long
End Type

Private Const PAGE_EXECUTE_READ As Long = &H20
Private Const PAGE_EXECUTE_READWRITE As Long = &H40

Private ASMThunk_GetWindow(63) As Byte

Private hwds(16383) As Long, hwdm As Long

'////////////////////////////////

Public Function Hook_GetWindow(ByVal bHook As Boolean) As Boolean
Static m_bHook As Boolean
Dim b As Boolean, i As Long
Dim lpFunc As Long, lpNewFunc As Long
Dim tMBI As MEMORY_BASIC_INFORMATION
If m_bHook <> bHook Then
 lpFunc = GetProcAddress(GetModuleHandle("user32.dll"), "GetWindow")
 lpNewFunc = VarPtr(ASMThunk_GetWindow(0))
 If bHook Then
  'backup old code
  CopyMemory i, ByVal lpFunc, 4&
  If i <> &H8B55FF8B Then Exit Function
  'create code
  CopyMemory ByVal lpNewFunc, &H8247C83, 4&
  CopyMemory ByVal lpNewFunc + &H4&, &HB81A7504, 4&
  CopyMemory ByVal lpNewFunc + &H8&, VarPtr(hwds(0)), 4&
  CopyMemory ByVal lpNewFunc + &HC&, &H1174C085, 4&
  CopyMemory ByVal lpNewFunc + &H10&, &H424548B, 4&
  CopyMemory ByVal lpNewFunc + &H14&, &H9E3088B, 4&
  CopyMemory ByVal lpNewFunc + &H18&, &HF74CA3B, 4&
  CopyMemory ByVal lpNewFunc + &H1C&, &HEB08C083, 4&
  CopyMemory ByVal lpNewFunc + &H20&, &H55FF8BF3, 4&
  CopyMemory ByVal lpNewFunc + &H24&, &HE9EC8B, 3&
  i = (lpFunc + 5) - (lpNewFunc + &H2B&)
  CopyMemory ByVal lpNewFunc + &H27&, i, 4&
  CopyMemory ByVal lpNewFunc + &H2B&, &HC204408B, 4&
  CopyMemory ByVal lpNewFunc + &H2F&, &H8&, 2&
  'hook
  CopyMemory ByVal lpNewFunc + &H38&, &HE9&, 1&
  i = lpNewFunc - (lpFunc + 5)
  CopyMemory ByVal lpNewFunc + &H39&, i, 4&
  VirtualQuery ByVal lpFunc, tMBI, Len(tMBI)
  VirtualProtect ByVal tMBI.BaseAddress, tMBI.RegionSize, PAGE_EXECUTE_READWRITE, i
  CopyMemory ByVal lpFunc, ByVal lpNewFunc + &H38&, 5&
  VirtualProtect ByVal tMBI.BaseAddress, tMBI.RegionSize, PAGE_EXECUTE_READ, i
  m_bHook = True
  Hook_GetWindow = True
 Else
  'unhook
  VirtualQuery ByVal lpFunc, tMBI, Len(tMBI)
  VirtualProtect ByVal tMBI.BaseAddress, tMBI.RegionSize, PAGE_EXECUTE_READWRITE, i
  CopyMemory ByVal lpFunc, ByVal lpNewFunc + &H21&, 5&
  VirtualProtect ByVal tMBI.BaseAddress, tMBI.RegionSize, PAGE_EXECUTE_READ, i
  m_bHook = False
  Hook_GetWindow = True
 End If
End If
End Function

Public Sub Hook_GetWindow_AddWindow(ByVal hWnd As Long, ByVal hWndOwner As Long)
Dim i As Long
For i = 0 To hwdm - 1 Step 2
 If hwds(i) = 0 Or hwds(i) = -1 Or hwds(i) = hWnd Then
  hwds(i) = hWnd
  hwds(i + 1) = hWndOwner
  Exit Sub
 End If
Next i
If i >= hwdm Then
 hwds(hwdm) = hWnd
 hwds(hwdm + 1) = hWndOwner
 hwdm = hwdm + 2
End If
End Sub

Public Sub Hook_GetWindow_RemoveWindow(ByVal hWnd As Long)
Dim i As Long
For i = 0 To hwdm - 1 Step 2
 If hwds(i) = hWnd Then
  hwds(i) = -1
  hwds(i + 1) = -1
  Exit Sub
 End If
Next i
End Sub

Public Sub Hook_GetWindow_ClearWindow()
Erase hwds
hwdm = 0
End Sub


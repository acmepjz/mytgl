Attribute VB_Name = "mdlFakeMenuHook"
Option Explicit

'////////////////////////////////
'This file is public domain.
'////////////////////////////////

'////////////////////////////////API Hook code

'; BOOL __stdcall TrackPopupMenuEx(HMENU hMenu, UINT fuFlags, int x, int y, HWND hWnd, LPTPMPARAMS lpTPMParams)
'77D6CF62 >  B8 35120000     MOV EAX,1235
'77D6CF67    BA 0003FE7F     MOV EDX,7FFE0300
'77D6CF6C    FF12            CALL DWORD PTR DS:[EDX]
'77D6CF6E    C2 1800         RETN 18

'XP only
'7FFE0300h = ntdll.KiFastSystemCall !!!!!!!!

'hook TrackPopupMenuEx code

'$ ==>    >    8D4424 04     LEA EAX,DWORD PTR SS:[ESP+4]             ;  entry
'$+4      >    50            PUSH EAX
'$+5      >    E8 82AA6DDE   CALL DEADBEEF                            ;  my func
'$+A      >    83F8 FF       CMP EAX,-1                               ;  if eax=-1 then default
'$+D      >    75 0A         JNZ SHORT ¹¤³ÌHook.0040147C
'$+F      >    B8 EFBEADDE   MOV EAX,DEADBEEF                         ;  old header
'$+14     >  - E9 73AA6DDE   JMP DEADBEEF                             ;  old func
'$+19     >    C2 1800       RETN 18

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

Private ASMThunk_TrackPopupMenuEx(63) As Byte

'////////////////////////////////

Private Declare Function GetCurrentThreadId Lib "kernel32.dll" () As Long
Private Declare Function SetWindowsHookEx Lib "user32.dll" Alias "SetWindowsHookExA" (ByVal idHook As Long, ByVal lpfn As Long, ByVal hmod As Long, ByVal dwThreadId As Long) As Long
Private Declare Function UnhookWindowsHookEx Lib "user32.dll" (ByVal hHook As Long) As Long
Private Declare Function CallNextHookEx Lib "user32.dll" (ByVal hHook As Long, ByVal nCode As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Declare Sub keybd_event Lib "user32.dll" (ByVal bVk As Byte, ByVal bScan As Byte, ByVal dwFlags As Long, ByVal dwExtraInfo As Long)
Private Const KEYEVENTF_KEYUP As Long = &H2

Private Const WH_CALLWNDPROC As Long = 4
Private Const WH_CALLWNDPROCRET As Long = 12
Private Const WH_CBT As Long = 5
Private Const WH_MSGFILTER As Long = -1
Private Const WH_GETMESSAGE As Long = 3

Private Const TPM_NONOTIFY As Long = &H80&
Private Const TPM_RETURNCMD As Long = &H100&

Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32.dll" Alias "GetMenuItemInfoW" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, ByRef lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetClassName Lib "user32.dll" Alias "GetClassNameA" (ByVal hwnd As Long, ByVal lpClassName As String, ByVal nMaxCount As Long) As Long

Private Declare Function SendMessage Lib "user32.dll" Alias "SendMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByRef lParam As Any) As Long
Private Const WM_INITMENUPOPUP As Long = &H117
Private Const WM_CONTEXTMENU As Long = &H7B
Private Const WM_SYSCOMMAND As Long = &H112
Private Const SC_MOUSEMENU As Long = &HF090&
Private Const SC_KEYMENU As Long = &HF100&
Private Const SC_MINIMIZE As Long = &HF020&
Private Const SC_MAXIMIZE As Long = &HF030&
Private Const SC_RESTORE As Long = &HF120&
Private Const SC_CLOSE As Long = &HF060&

Private Const MFT_STRING = 0
Private Const MFT_BITMAP = &H4&
Private Const MFT_MENUBARBREAK = &H20&
Private Const MFT_MENUBREAK = &H40&
Private Const MFT_OWNERDRAW = &H100&
Private Const MFT_RADIOCHECK = &H200&
Private Const MFT_RIGHTJUSTIFY = &H4000&
Private Const MFT_RIGHTORDER = &H2000&
Private Const MFT_SEPARATOR = &H800&

Private Const MIIM_BITMAP As Long = &H80
Private Const MIIM_CHECKMARKS As Long = &H8
Private Const MIIM_DATA As Long = &H20
Private Const MIIM_FTYPE As Long = &H100
Private Const MIIM_ID As Long = &H2
Private Const MIIM_STATE As Long = &H1
Private Const MIIM_STRING As Long = &H40
Private Const MIIM_SUBMENU As Long = &H4
Private Const MIIM_TYPE As Long = &H10

Private Const MFS_CHECKED As Long = &H8&
Private Const MFS_GRAYED As Long = &H3&
Private Const MFS_DEFAULT As Long = &H1000&

Private Const HBMMENU_CALLBACK As Long = -1
Private Const HBMMENU_MBAR_CLOSE As Long = 5
Private Const HBMMENU_MBAR_CLOSE_D As Long = 6
Private Const HBMMENU_MBAR_MINIMIZE As Long = 3
Private Const HBMMENU_MBAR_MINIMIZE_D As Long = 7
Private Const HBMMENU_MBAR_RESTORE As Long = 2
Private Const HBMMENU_POPUP_CLOSE As Long = 8
Private Const HBMMENU_POPUP_MAXIMIZE As Long = 10
Private Const HBMMENU_POPUP_MINIMIZE As Long = 11
Private Const HBMMENU_POPUP_RESTORE As Long = 9
Private Const HBMMENU_SYSTEM As Long = 1

Private Type MENUITEMINFO
    cbSize As Long
    fMask As Long
    fType As Long
    fState As Long
    wID As Long
    hSubMenu As Long
    hbmpChecked As Long
    hbmpUnchecked As Long
    dwItemData As Long
    lpszTypeData As Long
    cch As Long
    hbmpItem As Long
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Type RECT
    Left As Long
    Top As Long
    Right As Long
    Bottom As Long
End Type

Private Type TPMPARAMS
    cbSize As Long
    rcExclude As RECT
End Type

Private Type MENUBARINFO
 cbSize As Long
 rcBar As RECT
 hMenu As Long
 hwndMenu As Long
 dwFlags As Long
End Type

Private Type MSG
    hwnd As Long
    message As Long
    wParam As Long
    lParam As Long
    time As Long
    pt As POINTAPI
End Type

Private Type CWPSTRUCT
    lParam As Long
    wParam As Long
    message As Long
    hwnd As Long
End Type

Private Declare Function ClientToScreen Lib "user32.dll" (ByVal hwnd As Long, ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long
Private Declare Function GetMenuBarInfo Lib "user32.dll" (ByVal hwnd As Long, ByVal idObject As Long, ByVal idItem As Long, ByRef pmbi As MENUBARINFO) As Long
Private Const OBJID_MENU As Long = &HFFFFFFFD
Private Declare Function GetSystemMenu Lib "user32.dll" (ByVal hwnd As Long, ByVal bRevert As Long) As Long
Private Declare Function EndMenu Lib "user32.dll" () As Long
Private Declare Function PostMessage Lib "user32.dll" Alias "PostMessageA" (ByVal hwnd As Long, ByVal wMsg As Long, ByVal wParam As Long, ByVal lParam As Long) As Long
Private Declare Function TrackPopupMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByRef lprc As Any) As Long

Private m_objMenu As FakeMenu, m_hWnd As Long, m_hHook As Long, m_hHook2 As Long, m_MenuPos As POINTAPI

'////////////////////////////////API Hook function

Private Function Dummy(ByVal n As Long) As Long
Dummy = n
End Function

Public Function Hook_TrackPopupMenuEx(ByVal bHook As Boolean) As Boolean
Static m_bHook As Boolean
Dim b As Boolean, i As Long
Dim lpFunc As Long, lpNewFunc As Long, lpMyFunc As Long
Dim tMBI As MEMORY_BASIC_INFORMATION
If m_bHook <> bHook Then
 lpFunc = GetProcAddress(GetModuleHandle("user32.dll"), "TrackPopupMenuEx")
 lpNewFunc = VarPtr(ASMThunk_TrackPopupMenuEx(0))
 lpMyFunc = Dummy(AddressOf My_TrackPopupMenuEx)
 If bHook Then
  'backup old code
  CopyMemory ByVal lpNewFunc + 48, ByVal lpFunc, 5&
  If ASMThunk_TrackPopupMenuEx(48) <> &HB8& Then Exit Function
  'create code
  CopyMemory ByVal lpNewFunc, &H424448D, 4&
  CopyMemory ByVal lpNewFunc + 4, &HE850&, 2&
  i = lpMyFunc - (lpNewFunc + 10)
  CopyMemory ByVal lpNewFunc + 6, i, 4&
  CopyMemory ByVal lpNewFunc + 10, &H75FFF883, 4&
  CopyMemory ByVal lpNewFunc + 14, &HB80A&, 2&
  CopyMemory ByVal lpNewFunc + 16, ByVal lpNewFunc + 49, 4&
  CopyMemory ByVal lpNewFunc + 20, &HE9&, 1&
  i = (lpFunc + 5) - (lpNewFunc + 25)
  CopyMemory ByVal lpNewFunc + 21, i, 4&
  CopyMemory ByVal lpNewFunc + 25, &H18C2&, 3&
  'create hook code
  CopyMemory ByVal lpNewFunc + 56, &HE9&, 1&
  i = lpNewFunc - (lpFunc + 5)
  CopyMemory ByVal lpNewFunc + 57, i, 4&
  'hook
  VirtualQuery ByVal lpFunc, tMBI, Len(tMBI)
  VirtualProtect ByVal tMBI.BaseAddress, tMBI.RegionSize, PAGE_EXECUTE_READWRITE, i
  CopyMemory ByVal lpFunc, ByVal lpNewFunc + 56, 5&
  VirtualProtect ByVal tMBI.BaseAddress, tMBI.RegionSize, PAGE_EXECUTE_READ, i
  m_bHook = True
  Hook_TrackPopupMenuEx = True
 Else
  'unhook
  VirtualQuery ByVal lpFunc, tMBI, Len(tMBI)
  VirtualProtect ByVal tMBI.BaseAddress, tMBI.RegionSize, PAGE_EXECUTE_READWRITE, i
  CopyMemory ByVal lpFunc, ByVal lpNewFunc + 48, 5&
  VirtualProtect ByVal tMBI.BaseAddress, tMBI.RegionSize, PAGE_EXECUTE_READ, i
  m_bHook = False
  Hook_TrackPopupMenuEx = True
 End If
End If
End Function

'////////////////////////////////menu function

Private Function My_MessageProc(ByVal nCode As Long, ByVal wParam As Long, ByRef lpMsg As CWPSTRUCT) As Long
Static b As Boolean
Dim p As POINTAPI
Dim tInfo As MENUBARINFO
Dim n As Long
If nCode >= 0 Then
 Select Case lpMsg.message
 Case WM_SYSCOMMAND
  n = lpMsg.wParam And &HFFFFFFF0
  If n = SC_MOUSEMENU Or n = SC_KEYMENU Then
   tInfo.cbSize = Len(tInfo)
   GetMenuBarInfo lpMsg.hwnd, OBJID_MENU, 0, tInfo
   If tInfo.hMenu <> 0 Then p.y = tInfo.rcBar.Top - tInfo.rcBar.Bottom - 1
   ClientToScreen lpMsg.hwnd, p
   m_MenuPos = p
   '????????
   b = True
   If n = SC_KEYMENU And (lpMsg.lParam And &HFF&) = &H20& Then 'Alt+Space
    'Debug.Print 2
    keybd_event vbKeyMenu, 0, KEYEVENTF_KEYUP, 0 '???
'    n = GetSystemMenu(lpMsg.hwnd, 0)
'    If n <> 0 Then
'     'PostMessage lpMsg.hwnd, WM_CONTEXTMENU, &HDEADBEEF, ByVal n
'     lpMsg.hwnd = &HDEADBEEF
'    End If
   End If
  End If
 Case WM_INITMENUPOPUP
  If lpMsg.lParam And &H10000 Then
   'Debug.Print 1
   If b = False Then EndMenu
   b = False
   PostMessage lpMsg.hwnd, WM_CONTEXTMENU, &HDEADBEEF, ByVal lpMsg.wParam
  End If
 Case WM_CONTEXTMENU, &H313&
  b = False
'  If lpMsg.wParam = &HDEADBEEF Then
'   Debug.Print 3
'   '///hehe!!! call TrackPopupMenu and hooked
'   n = TrackPopupMenu(lpMsg.lParam, TPM_NONOTIFY Or TPM_RETURNCMD, m_MenuPos.x, m_MenuPos.y, 0, lpMsg.hwnd, ByVal 0)
'   If n <> 0 Then
'    GetCursorPos p
'    PostMessage lpMsg.hwnd, WM_SYSCOMMAND, n, p.x Or (p.y * &H10000)
'   End If
'   '///
'   lpMsg.message = 0
'  Else
   m_MenuPos.x = lpMsg.lParam And &HFFFF&
   m_MenuPos.y = lpMsg.lParam \ &H10000
'  End If
 End Select
End If
My_MessageProc = CallNextHookEx(m_hHook, nCode, wParam, lpMsg)
End Function

Private Function My_MessageProc2(ByVal nCode As Long, ByVal wParam As Long, ByRef lpMsg As MSG) As Long
Static b As Boolean
Dim p As POINTAPI
Dim n As Long
If nCode >= 0 Then
 Select Case lpMsg.message
 Case WM_CONTEXTMENU
  If lpMsg.wParam = &HDEADBEEF Then
   EndMenu
   If b = False And lpMsg.lParam <> 0 Then
    'Debug.Print -3
    b = True
    '///hehe!!! call TrackPopupMenu and hooked
    n = TrackPopupMenu(lpMsg.lParam, TPM_NONOTIFY Or TPM_RETURNCMD, m_MenuPos.x, m_MenuPos.y, 0, lpMsg.hwnd, ByVal 0)
    If n <> 0 Then
     GetCursorPos p
     PostMessage lpMsg.hwnd, WM_SYSCOMMAND, n, p.x Or (p.y * &H10000)
    End If
   End If
   b = False
   '///
   lpMsg.message = 0
   lpMsg.wParam = 0
  End If
 End Select
End If
My_MessageProc2 = CallNextHookEx(m_hHook2, nCode, wParam, lpMsg)
End Function

Public Sub FakeMenuPopupHook(objMenu As FakeMenu)
If objMenu Is Nothing Then Exit Sub
Set m_objMenu = objMenu
Hook_TrackPopupMenuEx True
If m_hHook = 0 Then m_hHook = SetWindowsHookEx(WH_CALLWNDPROC, AddressOf My_MessageProc, 0, GetCurrentThreadId)
If m_hHook2 = 0 Then m_hHook2 = SetWindowsHookEx(WH_GETMESSAGE, AddressOf My_MessageProc2, 0, GetCurrentThreadId)
End Sub

Public Sub FakeMenuPopupUnhook()
Hook_TrackPopupMenuEx False
UnhookWindowsHookEx m_hHook
UnhookWindowsHookEx m_hHook2
m_hHook = 0
m_hHook2 = 0
End Sub

'fatal bug: check current thread id !!!
Private Function My_TrackPopupMenuEx(ByVal lpStack As Long) As Long
Dim s As String
Dim d(5) As Long
CopyMemory d(0), ByVal lpStack, &H18&
If d(0) <> 0 And Not m_objMenu Is Nothing Then
 m_hWnd = d(4)
 m_objMenu.Destroy
 pMenu d(0)
 m_objMenu.Hook m_hWnd, (d(1) And TPM_NONOTIFY) = 0
 m_objMenu.PopupMenuEx CStr(d(0)), d(2), d(3), , , , True, , , , s, True
 m_objMenu.Unhook
 If d(1) And TPM_RETURNCMD Then
  My_TrackPopupMenuEx = Val(s)
 Else
  My_TrackPopupMenuEx = (s <> "") And 1&
 End If
Else
 My_TrackPopupMenuEx = -1
End If
End Function

Private Sub pMenu(ByVal hMenu As Long)
Dim i As Long, j As Long, k As Long, m As Long
Dim tInfo As MENUITEMINFO
Dim s As String, sKey As String, sSubMenu As String
Dim idxMenu As Long
'add new menu
s = CStr(hMenu)
If m_objMenu.HasMenu(s) Then Exit Sub
idxMenu = m_objMenu.AddMenu(s)
'enum old menu item
tInfo.cbSize = Len(tInfo)
m = GetMenuItemCount(hMenu)
For i = 0 To m - 1
 s = Space(1024&)
 tInfo.lpszTypeData = StrPtr(s)
 tInfo.cch = 1024&
 tInfo.fMask = MIIM_FTYPE Or MIIM_ID Or MIIM_STRING Or MIIM_SUBMENU Or MIIM_STATE Or MIIM_CHECKMARKS Or MIIM_DATA
 GetMenuItemInfo hMenu, i, 1, tInfo
 j = InStr(1, s, vbNullChar)
 If j > 0 Then s = Left(s, j - 1)
 'check type
 If tInfo.fType And (MFT_MENUBREAK Or MFT_MENUBARBREAK) Then
  m_objMenu.AddButtonByIndex idxMenu, , , , , fbttColumnSeparator
 ElseIf tInfo.fType And MFT_SEPARATOR Then
  m_objMenu.AddButtonByIndex idxMenu, , , , , fbttSeparator
 Else
  sKey = CStr(tInfo.wID)
  If tInfo.fState And MFS_GRAYED Then k = 2 Else k = 0
  If tInfo.fState And MFS_DEFAULT Then k = k Or 64&
  If tInfo.fType And MFT_OWNERDRAW Then k = k Or 24&
  If tInfo.hSubMenu <> 0 Then
   sSubMenu = CStr(tInfo.hSubMenu)
   'recursive
   If k = 0 Then
    '////////????????
    SendMessage m_hWnd, WM_INITMENUPOPUP, tInfo.hSubMenu, ByVal 0
    '////////
    pMenu tInfo.hSubMenu
   End If
  Else
   sSubMenu = ""
  End If
  If tInfo.fType And MFT_RADIOCHECK Then j = 3 Else j = 0
  'bitmap?
  Select Case tInfo.wID
  Case SC_MINIMIZE: tInfo.hbmpUnchecked = HBMMENU_POPUP_MINIMIZE
  Case SC_MAXIMIZE: tInfo.hbmpUnchecked = HBMMENU_POPUP_MAXIMIZE
  Case SC_RESTORE: tInfo.hbmpUnchecked = HBMMENU_POPUP_RESTORE
  Case SC_CLOSE: tInfo.hbmpUnchecked = HBMMENU_POPUP_CLOSE
  End Select
  If tInfo.hbmpUnchecked <> 0 Then
   sKey = sKey + ";" + CStr(tInfo.hbmpUnchecked)
  End If
  If tInfo.dwItemData <> 0 Then
   sKey = sKey + "?" + CStr(tInfo.dwItemData)
  End If
  'over
  k = m_objMenu.AddButtonByIndex(idxMenu, , sKey, s, , j, k, , , , , sSubMenu, tInfo.fState And MFS_CHECKED)
 End If
Next i
End Sub

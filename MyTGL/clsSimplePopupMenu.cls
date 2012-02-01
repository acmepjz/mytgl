VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsSimplePopupMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private Declare Function CheckMenuItem Lib "user32.dll" (ByVal hMenu As Long, ByVal wIDCheckItem As Long, ByVal wCheck As Long) As Long
Private Declare Function CheckMenuRadioItem Lib "user32.dll" (ByVal hMenu As Long, ByVal un1 As Long, ByVal un2 As Long, ByVal un3 As Long, ByVal un4 As Long) As Long
Private Declare Function CreateMenu Lib "user32.dll" () As Long
Private Declare Function CreatePopupMenu Lib "user32.dll" () As Long
Private Declare Function DeleteMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function DestroyMenu Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function EnableMenuItem Lib "user32.dll" (ByVal hMenu As Long, ByVal wIDEnableItem As Long, ByVal wEnable As Long) As Long
Private Declare Function EndMenu Lib "user32.dll" () As Long
Private Declare Function GetMenuCheckMarkDimensions Lib "user32.dll" () As Long
Private Declare Function GetMenuContextHelpId Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuDefaultItem Lib "user32.dll" (ByVal hMenu As Long, ByVal fByPos As Long, ByVal gmdiFlags As Long) As Long
Private Declare Function GetMenuItemCount Lib "user32.dll" (ByVal hMenu As Long) As Long
Private Declare Function GetMenuItemID Lib "user32.dll" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function GetMenuItemInfo Lib "user32.dll" Alias "GetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal b As Long, ByRef lpMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function GetMenuState Lib "user32.dll" (ByVal hMenu As Long, ByVal wID As Long, ByVal wFlags As Long) As Long
Private Declare Function GetMenuString Lib "user32.dll" Alias "GetMenuStringA" (ByVal hMenu As Long, ByVal wIDItem As Long, ByVal lpString As String, ByVal nMaxCount As Long, ByVal wFlag As Long) As Long
Private Declare Function GetSubMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPos As Long) As Long
Private Declare Function HiliteMenuItem Lib "user32.dll" (ByVal hwnd As Long, ByVal hMenu As Long, ByVal wIDHiliteItem As Long, ByVal wHilite As Long) As Long
Private Declare Function InsertMenuItem Lib "user32.dll" Alias "InsertMenuItemA" (ByVal hMenu As Long, ByVal un As Long, ByVal BOOL As Long, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function RemoveMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long) As Long
Private Declare Function SetMenuContextHelpId Lib "user32.dll" (ByVal hMenu As Long, ByVal dw As Long) As Long
Private Declare Function SetMenuDefaultItem Lib "user32.dll" (ByVal hMenu As Long, ByVal uItem As Long, ByVal fByPos As Long) As Long
Private Declare Function SetMenuItemBitmaps Lib "user32.dll" (ByVal hMenu As Long, ByVal nPosition As Long, ByVal wFlags As Long, ByVal hBitmapUnchecked As Long, ByVal hBitmapChecked As Long) As Long
Private Declare Function SetMenuItemInfo Lib "user32.dll" Alias "SetMenuItemInfoA" (ByVal hMenu As Long, ByVal un As Long, ByVal BOOL As Long, ByRef lpcMenuItemInfo As MENUITEMINFO) As Long
Private Declare Function TrackPopupMenu Lib "user32.dll" (ByVal hMenu As Long, ByVal wFlags As Long, ByVal x As Long, ByVal y As Long, ByVal nReserved As Long, ByVal hwnd As Long, ByRef lprc As Any) As Long

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
End Type

Private Type POINTAPI
    x As Long
    y As Long
End Type

Private Const MF_BYCOMMAND As Long = &H0&
Private Const MF_BYPOSITION As Long = &H400&

Public Enum enumMenuItemType
 MFT_STRING = 0
 MFT_BITMAP = &H4&
 MFT_MENUBARBREAK = &H20&
 MFT_MENUBREAK = &H40&
 MFT_OWNERDRAW = &H100&
 MFT_RADIOCHECK = &H200&
 MFT_RIGHTJUSTIFY = &H4000&
 MFT_RIGHTORDER = &H2000&
 MFT_SEPARATOR = &H800&
End Enum

Private Const MIIM_BITMAP As Long = &H80
Private Const MIIM_CHECKMARKS As Long = &H8
Private Const MIIM_DATA As Long = &H20
Private Const MIIM_FTYPE As Long = &H100
Private Const MIIM_ID As Long = &H2
Private Const MIIM_STATE As Long = &H1
Private Const MIIM_STRING As Long = &H40
Private Const MIIM_SUBMENU As Long = &H4
Private Const MIIM_TYPE As Long = &H10

Private Const TPM_RETURNCMD As Long = &H100&
Private Const TPM_RECURSE As Long = &H1&
Private Const TPM_NONOTIFY As Long = &H80&

Private Declare Function GetCursorPos Lib "user32.dll" (ByRef lpPoint As POINTAPI) As Long

Private Type typeMenu
 Key As String
 hMenu As Long
End Type

Private mnus() As typeMenu, mnuc As Long

Public Property Get Count() As Long
Count = mnuc
End Property

Private Sub Class_Terminate()
Destroy
End Sub

Public Sub Destroy()
Dim i As Long
For i = 1 To mnuc
 DestroyMenu mnus(i).hMenu
Next i
Erase mnus
mnuc = 0
End Sub

Public Property Get MenuItemCount(ByVal Index As Long) As Long
If Index > 0 And Index <= mnuc Then
 MenuItemCount = GetMenuItemCount(mnus(Index).hMenu)
End If
End Property

Public Sub ClearMenu(ByVal Index As Long)
Dim i As Long, m As Long
If Index > 0 And Index <= mnuc Then
 With mnus(Index)
  m = GetMenuItemCount(.hMenu)
  For i = m - 1 To 0 Step -1
   RemoveMenu .hMenu, i, MF_BYPOSITION
  Next i
 End With
End If
End Sub

Public Property Get Key(ByVal Index As Long) As String
If Index > 0 And Index <= mnuc Then
 Key = mnus(Index).Key
End If
End Property

Public Property Get hMenu(ByVal Index As Long) As Long
If Index > 0 And Index <= mnuc Then
 hMenu = mnus(Index).hMenu
End If
End Property

Public Property Let Key(ByVal Index As Long, s As String)
If Index > 0 And Index <= mnuc Then
 mnus(Index).Key = s
End If
End Property

Public Function AddMenu(Optional ByVal Key As String, Optional ByVal IsPopup As Boolean) As Long
mnuc = mnuc + 1
ReDim Preserve mnus(1 To mnuc)
With mnus(mnuc)
 .Key = Key
 If IsPopup Then
  .hMenu = CreatePopupMenu
 Else
  .hMenu = CreateMenu
 End If
End With
AddMenu = mnuc
End Function

Public Function PopupMenu(ByVal hwnd As Long, ByVal Index As Long, Optional ByVal wFlags As Long, Optional ByVal x As Long = -1, Optional ByVal y As Long = -1, Optional ByVal DefaultMenu As Long = -1) As Long
Dim p As POINTAPI
If Index > 0 And Index <= mnuc Then
 If x < 0 Or y < 0 Then
  GetCursorPos p
 Else
  p.x = x
  p.y = y
 End If
 PopupMenu = TrackPopupMenu(mnus(Index).hMenu, wFlags Or TPM_RETURNCMD Or TPM_NONOTIFY, p.x, p.y, 0, hwnd, ByVal 0)
End If
End Function

Public Function AddItem(ByVal Index As Long, Optional ByVal nType As enumMenuItemType, Optional ByVal Caption As String, Optional ByVal wID As Long, Optional ByVal hSubMenu As Long, Optional ByVal idxInsert As Long = -1) As Boolean
Dim d As MENUITEMINFO
Dim s As String
If Index > 0 And Index <= mnuc Then
 With d
  .cbSize = Len(d)
  If nType And (MFT_SEPARATOR Or MFT_MENUBREAK Or MFT_MENUBARBREAK) Then
   .fMask = 0
  Else
   .cch = -1
   s = StrConv(Caption, vbFromUnicode) + vbNullChar
   .lpszTypeData = StrPtr(s)
   .fMask = MIIM_STRING
  End If
  If nType Then
   .fMask = .fMask Or MIIM_TYPE
   .fType = nType
  End If
  If wID Then
   .fMask = .fMask Or MIIM_ID
   .wID = wID
  End If
  If hSubMenu Then
   .fMask = .fMask Or MIIM_SUBMENU
   .hSubMenu = hSubMenu
  End If
 End With
 AddItem = InsertMenuItem(mnus(Index).hMenu, idxInsert, 1, d) <> 0
End If
End Function

Public Function RemoveItem(ByVal Index As Long, ByVal idxItem As Long) As Boolean
If Index > 0 And Index <= mnuc Then
 RemoveItem = RemoveMenu(mnus(Index).hMenu, idxItem, MF_BYPOSITION) <> 0
End If
End Function

Public Sub Remove(ByVal Index As Long)
Dim i As Long
If Index > 0 And Index <= mnuc Then
 DestroyMenu mnus(Index).hMenu
 If mnuc <= 1 Then
  Erase mnus
  mnuc = 0
 Else
  mnuc = mnuc - 1
  For i = Index To mnuc
   mnus(i) = mnus(i + 1)
  Next i
  ReDim Preserve mnus(1 To mnuc)
 End If
End If
End Sub

Public Function IndexFromKey(ByVal Key As String) As Long
Dim i As Long
For i = 1 To mnuc
 If mnus(i).Key = Key Then
  IndexFromKey = i
  Exit Function
 End If
Next i
End Function

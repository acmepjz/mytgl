VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "GFakeToolbarGlobal"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'////////////////////////////////
'This file is public domain.
'////////////////////////////////

'////////////////////////public constants

Public Enum enumFakeCommandBarFlags
 fcbfDragToMakeThisMenuFloat = 1
 fcbfToolBarMode = 2
End Enum

Public Enum enumFakeButtonType
 fbttNormal = 0
 fbttSeparator = 1
 fbttCheck = 2
 fbttOption = 3
 fbttOptionNullable = 4
 fbttSplit = 5
 fbttColumnSeparator = 6
End Enum

Public Enum enumFakeButtonFlags
 fbtfHidden = 1&
 fbtfDisabled = 2&
 fbtfShowDropdown = 4&
 fbtfOwnerMeasure = 8&
 fbtfOwnerDraw = 16&
 fbtfOwnerDrawAfter = 32&
 fbtfDefaultItem = 64&
 fbtfHideCationInToolBar = 128&
 fbtfStartNewRow = 256&
 fbtfFullRow = 512&
 fbtfColor = 1024&
 fbtfNoDisabledColor = 2048&
End Enum

Public Enum enumFakeButtonOwnerDrawType
 fbtoBefore = 0
 fbtoAfter = 1
End Enum

Public Enum enumFakeButtonOwnerDrawFlags
 fbtoDoDefaultHighlight = 1&
 fbtoDoDefaultIcon = 2&
 fbtoDoDefaultCaption = 4&
 fbtoDoDefaultDropdown = 8&
 fbtoDoDefaultChecked = 16&
 fbtoHasIcon = 65536
 fbtoCheckIconFromPicLeft = 131072
End Enum

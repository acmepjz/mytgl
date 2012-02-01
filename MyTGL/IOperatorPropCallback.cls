VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "IOperatorPropCallback"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
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

'callback interface

'nSubPropIndex=0-based
'nState 1=highlight 2=selected
'return value true=ownerdraw false=do default
Public Function Draw(ByRef objCaller As Object, ByVal hdc As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByRef bData() As Byte, ByVal lpPropDef As Long, ByVal lpPropTypeDef As Long, ByVal nSubPropIndex As Long, ByVal nState As Long) As Boolean
'
End Function

'nSubPropIndex=0-based &HFFFFFFFF=Custom
'return value true=value changed (eg. custom dialog) false=value unchanged (eg. edit box)
Public Function EditBegin(ByRef objCaller As Object, ByVal hwnd As Long, ByVal Left As Long, ByVal Top As Long, ByVal Width As Long, ByVal Height As Long, ByRef bData() As Byte, ByVal lpPropDef As Long, ByVal lpPropTypeDef As Long, ByVal nSubPropIndex As Long, ByRef objTextBox As TextBox, ByRef objComboBox As FakeComboBox, ByRef objLeftRight As LeftRight, ByRef objReserved As Object, ByVal bDblClick As Boolean) As Boolean
'
End Function

Public Function EditEnd(ByRef objCaller As Object) As Boolean
'
End Function

'return value=changed
Public Function OnChange(ByRef objCaller As Object, ByRef bData() As Byte, ByVal lpPropDef As Long, ByVal lpPropTypeDef As Long, ByVal nSubPropIndex As Long, ByVal sText As String) As Boolean
'
End Function

'nSubPropIndex=0-based &HFFFFFFFF=all
'return value=changed
Public Function OnDelta(ByRef objCaller As Object, ByRef bData() As Byte, ByVal lpPropDef As Long, ByVal lpPropTypeDef As Long, ByVal nSubPropIndex As Long, ByVal iDelta As Long, ByRef bCancel As Boolean) As Boolean
'
End Function

'nSubPropIndex=0-based &HFFFFFFFF=all
'return value=changed
Public Function Reset(ByRef objCaller As Object, ByRef bData() As Byte, ByVal lpPropDef As Long, ByVal lpPropTypeDef As Long, ByVal nSubPropIndex As Long) As Boolean
'
End Function

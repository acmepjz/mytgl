VERSION 5.00
Begin VB.Form frmStr 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Properties"
   ClientHeight    =   4800
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   6000
   BeginProperty Font 
      Name            =   "Tahoma"
      Size            =   8.25
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   320
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.ListBox lstStr 
      Height          =   4155
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.TextBox txtStr 
      Height          =   4215
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   5775
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   4320
      TabIndex        =   2
      Top             =   4380
      Width           =   735
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   3
      Top             =   4380
      Width           =   735
   End
End
Attribute VB_Name = "frmStr"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
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

Public oString As String

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
If txtStr.Visible Then
 oString = txtStr.Text
Else
 If lstStr.ListIndex >= 0 Then oString = lstStr.List(lstStr.ListIndex)
End If
Unload Me
End Sub

Private Sub lstStr_DblClick()
If lstStr.ListIndex >= 0 Then Command2_Click
End Sub

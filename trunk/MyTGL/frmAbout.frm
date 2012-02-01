VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "About"
   ClientHeight    =   4200
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
   ScaleHeight     =   280
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   4800
      TabIndex        =   0
      Top             =   3720
      Width           =   1095
   End
   Begin VB.Label Label1 
      Caption         =   $"frmAbout.frx":0000
      Height          =   3135
      Index           =   1
      Left            =   120
      TabIndex        =   2
      Top             =   480
      Width           =   5775
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   5775
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Form_Load()
Label1(0).Caption = "MyTGL alpha " + CStr(App.Major) + "." + CStr(App.Minor) + "." + CStr(App.Revision) + " by acme_pjz"
Label1(1).Caption = Label1(1).Caption + vbCr + vbCr + _
"This program is free software: you can redistribute it and/or modify it under the terms of the GNU General Public License as published by the Free Software Foundation, either version 3 of the License, or (at your option) any later version." + vbCr + vbCr + _
"This program is distributed in the hope that it will be useful, but WITHOUT ANY WARRANTY; without even the implied warranty of MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU General Public License for more details." + vbCr + vbCr + _
"You should have received a copy of the GNU General Public License along with this program.  If not, see <http://www.gnu.org/licenses/>."
End Sub

VERSION 5.00
Begin VB.Form frmOpt 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Options"
   ClientHeight    =   3600
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   4800
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
   ScaleHeight     =   240
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   320
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.CommandButton Command3 
      Caption         =   "Clear"
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1080
      Width           =   855
   End
   Begin VB.Frame Frame1 
      Caption         =   "Recent files"
      Height          =   615
      Index           =   1
      Left            =   120
      TabIndex        =   5
      Top             =   840
      Width           =   4575
   End
   Begin MyTGL.FakeComboBox cmbComp 
      Height          =   255
      Left            =   1560
      TabIndex        =   2
      Top             =   360
      Width           =   1695
      _ExtentX        =   2990
      _ExtentY        =   450
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BackColor       =   -2147483643
      ForeColor       =   -2147483640
   End
   Begin VB.CommandButton Command2 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   3600
      TabIndex        =   1
      Top             =   3120
      Width           =   1095
   End
   Begin VB.CommandButton Command1 
      Caption         =   "OK"
      Default         =   -1  'True
      Height          =   375
      Left            =   2400
      TabIndex        =   0
      Top             =   3120
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      Caption         =   "Compression"
      Height          =   615
      Index           =   0
      Left            =   120
      TabIndex        =   3
      Top             =   120
      Width           =   4575
      Begin VB.Label Label1 
         Caption         =   "File Compression"
         Height          =   255
         Left            =   120
         TabIndex        =   4
         Top             =   240
         Width           =   1335
      End
   End
End
Attribute VB_Name = "frmOpt"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private Sub Command1_Click()
cSet.SetSettings "CompressMode", CStr(cmbComp.ListIndex)
'save setting
cSet.SaveFile
'exit
Unload Me
End Sub

Private Sub Command2_Click()
Unload Me
End Sub

Private Sub Command3_Click()
Dim i As Long
For i = 1 To 8
 cSet.Remove "MRU" + CStr(i)
Next i
Form1.pMRU
End Sub

Private Sub Form_Load()
Dim i As Long
With cmbComp
 .AddItem "No compression"
 .AddItem "LZSS"
 .AddItem "LZMA"
 .AddItem "ZLib"
End With
'show options
i = Val(cSet.GetSettings("CompressMode", "1"))
If i < 0 Or i >= cmbComp.ListCount Then i = 1
cmbComp.ListIndex = i
End Sub

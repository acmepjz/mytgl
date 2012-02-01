VERSION 5.00
Begin VB.Form frmComment 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Edit Comment"
   ClientHeight    =   4500
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
   ScaleHeight     =   300
   ScaleMode       =   3  'Pixel
   ScaleWidth      =   400
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'ÆÁÄ»ÖÐÐÄ
   Begin VB.TextBox Text1 
      Height          =   3165
      Index           =   1
      Left            =   120
      MultiLine       =   -1  'True
      ScrollBars      =   3  'Both
      TabIndex        =   7
      Top             =   840
      Width           =   5775
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Index           =   0
      Left            =   600
      TabIndex        =   5
      Top             =   120
      Width           =   3735
   End
   Begin VB.CommandButton Command1 
      Cancel          =   -1  'True
      Caption         =   "Cancel"
      Height          =   375
      Left            =   5160
      TabIndex        =   1
      Top             =   4080
      Width           =   735
   End
   Begin VB.CommandButton Command2 
      Caption         =   "OK"
      Height          =   375
      Left            =   4320
      TabIndex        =   0
      Top             =   4080
      Width           =   735
   End
   Begin VB.Label lblColor 
      BorderStyle     =   1  'Fixed Single
      Height          =   255
      Left            =   4920
      MousePointer    =   10  'Up Arrow
      TabIndex        =   6
      Top             =   120
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Description"
      Height          =   255
      Index           =   2
      Left            =   120
      TabIndex        =   4
      Top             =   480
      Width           =   975
   End
   Begin VB.Label Label1 
      Caption         =   "Color"
      Height          =   255
      Index           =   1
      Left            =   4440
      TabIndex        =   3
      Top             =   120
      Width           =   420
   End
   Begin VB.Label Label1 
      Caption         =   "Title"
      Height          =   255
      Index           =   0
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   615
   End
End
Attribute VB_Name = "frmComment"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private c As typeComment

Friend Sub fGetComment(d As typeComment)
d = c
End Sub

Friend Sub fSetComment(d As typeComment)
c = d
End Sub

Private Sub Command1_Click()
Unload Me
End Sub

Private Sub Command2_Click()
c.Name = Text1(0).Text
c.Value = Text1(1).Text
c.Color = lblColor.BackColor
Unload Me
End Sub

Private Sub Form_Load()
Text1(0).Text = c.Name
Text1(1).Text = c.Value
lblColor.BackColor = c.Color
End Sub

Private Sub lblColor_Click()
Dim clr As Long
clr = lblColor.BackColor
If cd.VBChooseColor(clr, , , , Me.hWnd) Then
 lblColor.BackColor = clr
End If
End Sub

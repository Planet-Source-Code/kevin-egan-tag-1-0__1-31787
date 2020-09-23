VERSION 5.00
Begin VB.Form aboutfrm 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "About TAG 1.0"
   ClientHeight    =   2100
   ClientLeft      =   8490
   ClientTop       =   5715
   ClientWidth     =   3300
   LinkTopic       =   "Form2"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2100
   ScaleWidth      =   3300
   Begin VB.CommandButton cmdok 
      Caption         =   "OK"
      Height          =   375
      Left            =   1320
      TabIndex        =   4
      Top             =   1680
      Width           =   615
   End
   Begin VB.Label Label4 
      Caption         =   "AIM: kevinsit"
      Height          =   255
      Left            =   1200
      TabIndex        =   3
      Top             =   1320
      Width           =   975
   End
   Begin VB.Label Label3 
      Caption         =   "Mail: kevinegan@cbs.ie"
      Height          =   255
      Left            =   840
      TabIndex        =   2
      Top             =   960
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "By Kevin Egan ."
      Height          =   255
      Left            =   1080
      TabIndex        =   1
      Top             =   600
      Width           =   1215
   End
   Begin VB.Label Label1 
      Caption         =   "TAG 1.0"
      BeginProperty Font 
         Name            =   "Comic Sans MS"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00800080&
      Height          =   375
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1335
   End
End
Attribute VB_Name = "aboutfrm"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False




Private Sub Command1_Click()

End Sub

Private Sub cmdok_Click()
aboutfrm.Hide
End Sub

Private Sub Form_QueryUnload(Cancel As Integer, UnloadMode As Integer)
Cancel = True
End Sub


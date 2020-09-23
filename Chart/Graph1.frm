VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Example of graphic"
   ClientHeight    =   8505
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   8895
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   8895
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text6 
      Height          =   375
      Left            =   8760
      TabIndex        =   33
      Text            =   "Text6"
      Top             =   8400
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox Text5 
      Height          =   285
      Left            =   8760
      TabIndex        =   32
      Text            =   "Text5"
      Top             =   9000
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text4 
      Height          =   285
      Left            =   8880
      TabIndex        =   31
      Text            =   "Text4"
      Top             =   8520
      Visible         =   0   'False
      Width           =   255
   End
   Begin VB.TextBox Text3 
      Height          =   375
      Left            =   9240
      TabIndex        =   30
      Text            =   "Text3"
      Top             =   8880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   9600
      TabIndex        =   29
      Text            =   "Text1"
      Top             =   8880
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox TextE 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4200
      MaxLength       =   5
      TabIndex        =   5
      Top             =   7680
      Width           =   1455
   End
   Begin VB.TextBox TextD 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      MaxLength       =   5
      TabIndex        =   4
      Top             =   6840
      Width           =   1455
   End
   Begin VB.TextBox TextC 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   5280
      MaxLength       =   5
      TabIndex        =   3
      Top             =   6240
      Width           =   1455
   End
   Begin VB.TextBox TextB 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   2
      Top             =   6840
      Width           =   1455
   End
   Begin VB.Timer Timer1 
      Enabled         =   0   'False
      Interval        =   1
      Left            =   8640
      Top             =   8160
   End
   Begin VB.CommandButton Command3 
      Caption         =   "Make Chart"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   6720
      TabIndex        =   6
      Top             =   7680
      Width           =   1815
   End
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   8880
      TabIndex        =   18
      Text            =   "Text2"
      Top             =   8400
      Visible         =   0   'False
      Width           =   150
   End
   Begin VB.TextBox TextA 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   1560
      MaxLength       =   5
      TabIndex        =   1
      Top             =   6240
      Width           =   1455
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H80000009&
      Height          =   5000
      Left            =   960
      ScaleHeight     =   5000
      ScaleMode       =   0  'User
      ScaleWidth      =   7440
      TabIndex        =   0
      Top             =   240
      Width           =   7500
   End
   Begin VB.Label Label21 
      AutoSize        =   -1  'True
      Caption         =   "Value of E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   2640
      TabIndex        =   28
      Top             =   7800
      Width           =   1305
   End
   Begin VB.Label Label20 
      AutoSize        =   -1  'True
      Caption         =   "Value of D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3840
      TabIndex        =   27
      Top             =   6960
      Width           =   1305
   End
   Begin VB.Label Label19 
      AutoSize        =   -1  'True
      Caption         =   "Value of C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   3840
      TabIndex        =   26
      Top             =   6360
      Width           =   1305
   End
   Begin VB.Label Label18 
      AutoSize        =   -1  'True
      Caption         =   "Value of B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   120
      TabIndex        =   25
      Top             =   6960
      Width           =   1290
   End
   Begin VB.Label Label17 
      AutoSize        =   -1  'True
      Caption         =   "Value of A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   360
      Left            =   120
      TabIndex        =   24
      Top             =   6360
      Width           =   1305
   End
   Begin VB.Label Label16 
      AutoSize        =   -1  'True
      Caption         =   "E"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   7800
      TabIndex        =   23
      Top             =   5280
      Width           =   195
   End
   Begin VB.Label Label15 
      AutoSize        =   -1  'True
      Caption         =   "D"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   6240
      TabIndex        =   22
      Top             =   5280
      Width           =   195
   End
   Begin VB.Label Label14 
      AutoSize        =   -1  'True
      Caption         =   "C"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   4680
      TabIndex        =   21
      Top             =   5280
      Width           =   195
   End
   Begin VB.Label Label13 
      AutoSize        =   -1  'True
      Caption         =   "B"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   3240
      TabIndex        =   20
      Top             =   5280
      Width           =   180
   End
   Begin VB.Label Label12 
      AutoSize        =   -1  'True
      Caption         =   "A"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000080&
      Height          =   360
      Left            =   1680
      TabIndex        =   19
      Top             =   5280
      Width           =   195
   End
   Begin VB.Label Label11 
      AutoSize        =   -1  'True
      Caption         =   "500"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   360
      TabIndex        =   17
      Top             =   4560
      Width           =   450
   End
   Begin VB.Label Label10 
      AutoSize        =   -1  'True
      Caption         =   "0"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   600
      TabIndex        =   16
      Top             =   4920
      Width           =   150
   End
   Begin VB.Label Label9 
      AutoSize        =   -1  'True
      Caption         =   "1000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   240
      TabIndex        =   15
      Top             =   4080
      Width           =   600
   End
   Begin VB.Label Label8 
      AutoSize        =   -1  'True
      Caption         =   "2000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   240
      TabIndex        =   14
      Top             =   3045
      Width           =   600
   End
   Begin VB.Label Label7 
      AutoSize        =   -1  'True
      Caption         =   "1500"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   240
      TabIndex        =   13
      Top             =   3540
      Width           =   600
   End
   Begin VB.Label Label6 
      AutoSize        =   -1  'True
      Caption         =   "3000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   240
      TabIndex        =   12
      Top             =   2085
      Width           =   600
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "2500"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   240
      TabIndex        =   11
      Top             =   2565
      Width           =   600
   End
   Begin VB.Label Label4 
      AutoSize        =   -1  'True
      Caption         =   "4000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   240
      TabIndex        =   10
      Top             =   1080
      Width           =   600
   End
   Begin VB.Label Label3 
      AutoSize        =   -1  'True
      Caption         =   "3500"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   240
      TabIndex        =   9
      Top             =   1560
      Width           =   600
   End
   Begin VB.Label Label2 
      AutoSize        =   -1  'True
      Caption         =   "5000"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   240
      TabIndex        =   8
      Top             =   120
      Width           =   600
   End
   Begin VB.Label Label1 
      AutoSize        =   -1  'True
      Caption         =   "4500"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00008000&
      Height          =   360
      Left            =   240
      TabIndex        =   7
      Top             =   600
      Width           =   600
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command2_Click()

End Sub

Private Sub Command3_Click()

Picture1.Cls
Text2.Text = 5000 - Val(TextA)
Text3.Text = 5000 - Val(TextB)
Text4.Text = 5000 - Val(TextC)
Text5.Text = 5000 - Val(TextD)
Text6.Text = 5000 - Val(TextE)
Timer1.Enabled = True

End Sub

Private Sub Timer1_Timer()
For i = 1 To 1500
Picture1.Line (0 + i, Text2.Text)-(0 + i, Picture1.Height), vbRed
Picture1.Line (1500 + i, Text3.Text)-(1500 + i, Picture1.Height), vbGreen
Picture1.Line (3000 + i, Text4.Text)-(3000 + i, Picture1.Height), vbBlue
Picture1.Line (4500 + i, Text5.Text)-(4500 + i, Picture1.Height), vbYellow
Picture1.Line (6000 + i, Text6.Text)-(6000 + i, Picture1.Height), vbMagenta
Next i
Timer1.Enabled = False
End Sub

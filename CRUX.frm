VERSION 5.00
Begin VB.Form Form23 
   Caption         =   "Form23"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form23"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command8 
      Caption         =   "Main menu"
      Height          =   735
      Left            =   11880
      TabIndex        =   22
      Top             =   9840
      Width           =   1815
   End
   Begin VB.PictureBox Picture1 
      Height          =   9375
      Left            =   960
      Picture         =   "CRUX.frx":0000
      ScaleHeight     =   9315
      ScaleWidth      =   12675
      TabIndex        =   0
      Top             =   480
      Width           =   12735
   End
   Begin VB.PictureBox Picture9 
      Height          =   4695
      Left            =   12480
      Picture         =   "CRUX.frx":1F0A8
      ScaleHeight     =   4635
      ScaleWidth      =   6195
      TabIndex        =   17
      Top             =   5040
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture8 
      Height          =   4695
      Left            =   6240
      Picture         =   "CRUX.frx":29923
      ScaleHeight     =   4635
      ScaleWidth      =   6195
      TabIndex        =   16
      Top             =   5040
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture7 
      Height          =   4695
      Left            =   0
      Picture         =   "CRUX.frx":3338E
      ScaleHeight     =   4635
      ScaleWidth      =   6195
      TabIndex        =   15
      Top             =   5040
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture6 
      Height          =   5055
      Left            =   12480
      Picture         =   "CRUX.frx":3F5BD
      ScaleHeight     =   4995
      ScaleWidth      =   6195
      TabIndex        =   14
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture5 
      Height          =   5055
      Left            =   6240
      Picture         =   "CRUX.frx":4A568
      ScaleHeight     =   4995
      ScaleWidth      =   6195
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture4 
      Height          =   5055
      Left            =   0
      Picture         =   "CRUX.frx":53332
      ScaleHeight     =   4995
      ScaleWidth      =   6195
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture3 
      Height          =   3015
      Left            =   16200
      Picture         =   "CRUX.frx":5D8B3
      ScaleHeight     =   2955
      ScaleWidth      =   3915
      TabIndex        =   9
      Top             =   5760
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox Picture2 
      Height          =   3015
      Left            =   16200
      Picture         =   "CRUX.frx":621F6
      ScaleHeight     =   2955
      ScaleWidth      =   3915
      TabIndex        =   8
      Top             =   1920
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Book"
      Height          =   735
      Left            =   9720
      TabIndex        =   5
      Top             =   9840
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Height          =   735
      Left            =   7680
      Picture         =   "CRUX.frx":66D86
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9840
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Height          =   735
      Left            =   5520
      Picture         =   "CRUX.frx":674A7
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9840
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Height          =   735
      Left            =   3240
      Picture         =   "CRUX.frx":67BEC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9840
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   960
      Picture         =   "CRUX.frx":682F0
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9840
      Width           =   2295
   End
   Begin VB.PictureBox Picture10 
      Height          =   9735
      Left            =   3960
      Picture         =   "CRUX.frx":68CE4
      ScaleHeight     =   9675
      ScaleWidth      =   12795
      TabIndex        =   18
      Top             =   360
      Visible         =   0   'False
      Width           =   12855
      Begin VB.CommandButton Command6 
         Caption         =   "RETURN"
         Height          =   735
         Left            =   5280
         TabIndex        =   19
         Top             =   9000
         Width           =   2175
      End
   End
   Begin VB.PictureBox Picture11 
      Height          =   9735
      Left            =   3960
      Picture         =   "CRUX.frx":87E1C
      ScaleHeight     =   9675
      ScaleWidth      =   12795
      TabIndex        =   20
      Top             =   360
      Visible         =   0   'False
      Width           =   12855
      Begin VB.CommandButton Command7 
         Caption         =   "RETURN"
         Height          =   735
         Left            =   5160
         TabIndex        =   21
         Top             =   9000
         Width           =   2175
      End
   End
   Begin VB.Label Label4 
      Caption         =   "MAROON"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13800
      TabIndex        =   11
      Top             =   5760
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label3 
      Caption         =   "BLACK"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   13800
      TabIndex        =   10
      Top             =   1920
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   $"CRUX.frx":A6EC4
      BeginProperty Font 
         Name            =   "Century"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9735
      Left            =   13800
      TabIndex        =   7
      Top             =   480
      Visible         =   0   'False
      Width           =   6375
   End
   Begin VB.Label Label1 
      Caption         =   $"CRUX.frx":A75E4
      BeginProperty Font 
         Name            =   "Century"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9375
      Left            =   13800
      TabIndex        =   6
      Top             =   480
      Visible         =   0   'False
      Width           =   6255
   End
End
Attribute VB_Name = "Form23"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Picture1.Visible = True
Label1.Visible = True
Label2.Visible = False
Label3.Visible = False
Label4.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
Picture9.Visible = False
End Sub

Private Sub Command2_Click()
Picture1.Visible = True
Label2.Visible = True
Label3.Visible = False
Label4.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
Picture9.Visible = False
End Sub

Private Sub Command3_Click()
Picture1.Visible = False
Label1.Visible = False
Label2.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = True
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = True
Picture8.Visible = True
Picture9.Visible = True
End Sub

Private Sub Command4_Click()
Picture1.Visible = True
Label1.Visible = False
Label2.Visible = False
Label3.Visible = True
Label4.Visible = True
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
Picture9.Visible = False
End Sub

Private Sub Command5_Click()
Form17.Show
End Sub

Private Sub Command6_Click()
Command1.Visible = True
Command2.Visible = True
command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Picture1.Visible = True
Picture10.Visible = False
Label3.Visible = False
Label4.Visible = False
End Sub

Private Sub Command7_Click()
Command1.Visible = True
Command2.Visible = True
command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Picture1.Visible = True
picture11.Visible = False
Label3.Visible = False
Label4.Visible = False
End Sub

Private Sub Command8_Click()
Form28.Show
End Sub

Private Sub Picture2_Click()
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Command1.Visible = False
Command2.Visible = False
command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Picture10.Visible = True
End Sub

Private Sub Picture3_Click()
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Command1.Visible = False
Command2.Visible = False
command3.Visible = False
Command4.Visible = False
Command5.Visible = False
picture11.Visible = True
End Sub

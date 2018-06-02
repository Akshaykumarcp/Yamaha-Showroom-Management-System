VERSION 5.00
Begin VB.Form Form28 
   BackColor       =   &H8000000E&
   Caption         =   "Form28"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   BeginProperty Font 
      Name            =   "MS Sans Serif"
      Size            =   9.75
      Charset         =   0
      Weight          =   400
      Underline       =   0   'False
      Italic          =   0   'False
      Strikethrough   =   0   'False
   EndProperty
   LinkTopic       =   "Form28"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "Select.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2535
      TabIndex        =   9
      Top             =   0
      Width           =   2535
   End
   Begin VB.CommandButton Command8 
      Caption         =   "BACK"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   18
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   17880
      TabIndex        =   8
      Top             =   10320
      Width           =   2295
   End
   Begin VB.CommandButton command3 
      BackColor       =   &H8000000E&
      Caption         =   "FZ1"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   2775
      Left            =   960
      Picture         =   "Select.frx":0CD2
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   3000
      Width           =   3855
   End
   Begin VB.Timer Timer1 
      Interval        =   300
      Left            =   1800
      Top             =   960
   End
   Begin VB.CommandButton Command6 
      Caption         =   "VMAX"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   10440
      Picture         =   "Select.frx":4BA7
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   6240
      Width           =   3855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "YZF - R15"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   15240
      Picture         =   "Select.frx":852E
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   2760
      Width           =   3975
   End
   Begin VB.CommandButton Command4 
      Caption         =   "FZ-S FI"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3375
      Left            =   5760
      Picture         =   "Select.frx":FBAC
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   6240
      Width           =   3855
   End
   Begin VB.CommandButton Command2 
      Caption         =   "Crux"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   5640
      Picture         =   "Select.frx":1692D
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   2760
      Width           =   3975
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Fazer"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   3255
      Left            =   10440
      Picture         =   "Select.frx":1B270
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   2760
      Width           =   3855
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "SELECT            BIKE            FOR             MORE            INFORMATION"
      BeginProperty Font 
         Name            =   "Old English Text MT"
         Size            =   21.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000C0&
      Height          =   495
      Left            =   2280
      TabIndex        =   7
      Top             =   9720
      Width           =   15975
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "WELCOME TO YAMAHA BIKE"
      BeginProperty Font 
         Name            =   "Lucida Calligraphy"
         Size            =   48
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00404040&
      Height          =   1335
      Left            =   1560
      TabIndex        =   5
      Top             =   960
      Width           =   17055
   End
End
Attribute VB_Name = "Form28"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Form24.Show
End Sub

Private Sub Command2_Click()
Form23.Show
End Sub

Private Sub Command3_Click()
Form21.Show
End Sub

Private Sub Command4_Click()
Form25.Show
End Sub

Private Sub Command5_Click()
Form22.Show
End Sub

Private Sub Command6_Click()
Form26.Show
End Sub

Private Sub Command8_Click()
Form16.Show
End Sub

Private Sub Timer1_Timer()
If Label1.Visible = True Then
Label1.Visible = False
Else
Label1.Visible = True
End If
End Sub

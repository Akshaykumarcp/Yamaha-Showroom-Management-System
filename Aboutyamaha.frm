VERSION 5.00
Begin VB.Form Form14 
   BackColor       =   &H8000000E&
   Caption         =   "Form14"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form14"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command1 
      BackColor       =   &H00C0C0C0&
      Caption         =   "<=&BACK"
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
      Left            =   18240
      Style           =   1  'Graphical
      TabIndex        =   3
      ToolTipText     =   "CLICK TO MAIN SCREEN"
      Top             =   10320
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      Picture         =   "Aboutyamaha.frx":0000
      ScaleHeight     =   855
      ScaleWidth      =   2535
      TabIndex        =   1
      Top             =   0
      Width           =   2535
   End
   Begin VB.Label Label1 
      BackColor       =   &H80000009&
      BackStyle       =   0  'Transparent
      Caption         =   "  ABOUT    YAMAHA   MOTORS   "
      BeginProperty Font 
         Name            =   "Trebuchet MS"
         Size            =   27.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   735
      Left            =   5040
      TabIndex        =   2
      Top             =   240
      Width           =   8295
   End
   Begin VB.Label LABEL2 
      BackStyle       =   0  'Transparent
      Caption         =   $"Aboutyamaha.frx":0CD2
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   8415
      Left            =   2760
      TabIndex        =   0
      Top             =   1200
      Width           =   12735
   End
End
Attribute VB_Name = "Form14"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
MDIForm1.Show
End Sub

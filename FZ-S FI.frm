VERSION 5.00
Begin VB.Form Form25 
   Caption         =   "Form25"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form25"
   ScaleHeight     =   10950
   ScaleWidth      =   20250
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command6 
      Caption         =   "Main menu"
      Height          =   735
      Left            =   11640
      TabIndex        =   15
      Top             =   9840
      Width           =   1935
   End
   Begin VB.PictureBox Picture5 
      Height          =   2895
      Left            =   13920
      Picture         =   "FZ-S FI.frx":0000
      ScaleHeight     =   2835
      ScaleWidth      =   3915
      TabIndex        =   10
      Top             =   5400
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox Picture4 
      Height          =   2895
      Left            =   16320
      Picture         =   "FZ-S FI.frx":7DE9
      ScaleHeight     =   2835
      ScaleWidth      =   3915
      TabIndex        =   9
      Top             =   8040
      Visible         =   0   'False
      Width           =   3975
   End
   Begin VB.PictureBox Picture3 
      Height          =   2895
      Left            =   16440
      Picture         =   "FZ-S FI.frx":FAE0
      ScaleHeight     =   2835
      ScaleWidth      =   3795
      TabIndex        =   8
      Top             =   2760
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.PictureBox Picture2 
      Height          =   2895
      Left            =   13920
      Picture         =   "FZ-S FI.frx":16861
      ScaleHeight     =   2835
      ScaleWidth      =   3795
      TabIndex        =   7
      Top             =   -120
      Visible         =   0   'False
      Width           =   3855
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Book"
      Height          =   855
      Left            =   9120
      TabIndex        =   4
      Top             =   9720
      Width           =   2175
   End
   Begin VB.CommandButton Command4 
      Height          =   855
      Left            =   6240
      Picture         =   "FZ-S FI.frx":1E89C
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9720
      Width           =   2175
   End
   Begin VB.CommandButton Command2 
      Height          =   855
      Left            =   3360
      Picture         =   "FZ-S FI.frx":1EFBD
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9720
      Width           =   2295
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   720
      Picture         =   "FZ-S FI.frx":1F6C1
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9720
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   9255
      Left            =   720
      Picture         =   "FZ-S FI.frx":200B5
      ScaleHeight     =   9195
      ScaleWidth      =   12795
      TabIndex        =   0
      Top             =   600
      Width           =   12855
   End
   Begin VB.Label Label6 
      Caption         =   "MOLTEN ORANGE"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000040C0&
      Height          =   1095
      Left            =   14280
      TabIndex        =   14
      Top             =   9360
      Visible         =   0   'False
      Width           =   1695
   End
   Begin VB.Label Label5 
      Caption         =   "MOONWALK      WHITE"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   855
      Left            =   18000
      TabIndex        =   13
      Top             =   6480
      Visible         =   0   'False
      Width           =   2295
   End
   Begin VB.Label Label4 
      Caption         =   "CYBER GREEN"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FF80&
      Height          =   855
      Left            =   14400
      TabIndex        =   12
      Top             =   3600
      Visible         =   0   'False
      Width           =   1935
   End
   Begin VB.Label Label3 
      Caption         =   "ASPRAL BLUE"
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF8080&
      Height          =   855
      Left            =   18240
      TabIndex        =   11
      Top             =   840
      Visible         =   0   'False
      Width           =   2415
   End
   Begin VB.Label Label2 
      Caption         =   $"FZ-S FI.frx":4E008
      BeginProperty Font 
         Name            =   "Century"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10695
      Left            =   14040
      TabIndex        =   6
      Top             =   120
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.Label Label1 
      Caption         =   $"FZ-S FI.frx":4E7E8
      BeginProperty Font 
         Name            =   "Century"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   9615
      Left            =   13920
      TabIndex        =   5
      Top             =   240
      Visible         =   0   'False
      Width           =   6135
   End
End
Attribute VB_Name = "Form25"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Label1.Visible = True
Label2.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
End Sub

Private Sub Command2_Click()
Label2.Visible = True
Label1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Label3.Visible = False
Label4.Visible = False
Label5.Visible = False
Label6.Visible = False
End Sub

Private Sub Command4_Click()
Label1.Visible = False
Label2.Visible = False
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Label3.Visible = True
Label4.Visible = True
Label5.Visible = True
Label6.Visible = True
End Sub

Private Sub Command5_Click()
Form17.Show
End Sub

Private Sub Command6_Click()
Form28.Show
End Sub

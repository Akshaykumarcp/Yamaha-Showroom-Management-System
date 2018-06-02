VERSION 5.00
Begin VB.Form Form21 
   BackColor       =   &H80000008&
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form21"
   Picture         =   "FZ1.frx":0000
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.CommandButton Command9 
      Caption         =   "Main menu"
      Height          =   735
      Left            =   12000
      TabIndex        =   21
      Top             =   9600
      Width           =   2295
   End
   Begin VB.PictureBox Picture9 
      Height          =   8415
      Left            =   1920
      Picture         =   "FZ1.frx":2FEA
      ScaleHeight     =   8355
      ScaleWidth      =   12675
      TabIndex        =   16
      Top             =   1200
      Visible         =   0   'False
      Width           =   12735
      Begin VB.CommandButton Command10 
         Caption         =   "RETURN"
         Height          =   615
         Left            =   5280
         TabIndex        =   19
         Top             =   7800
         Width           =   2295
      End
   End
   Begin VB.CommandButton command6 
      BackColor       =   &H8000000E&
      Caption         =   "&FZ1-BLACK"
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
      Left            =   14880
      Picture         =   "FZ1.frx":1CC6F
      Style           =   1  'Graphical
      TabIndex        =   15
      Top             =   1560
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton command7 
      BackColor       =   &H8000000E&
      Caption         =   "&FZ1-WHITE"
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
      Left            =   14880
      Picture         =   "FZ1.frx":2112D
      Style           =   1  'Graphical
      TabIndex        =   14
      Top             =   6480
      Visible         =   0   'False
      Width           =   4455
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Book"
      Height          =   735
      Left            =   9960
      TabIndex        =   7
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Height          =   735
      Left            =   7920
      Picture         =   "FZ1.frx":25002
      Style           =   1  'Graphical
      TabIndex        =   6
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton Command3 
      Height          =   735
      Left            =   5880
      Picture         =   "FZ1.frx":25723
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Height          =   735
      Left            =   3840
      Picture         =   "FZ1.frx":25E68
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Height          =   735
      Left            =   1680
      Picture         =   "FZ1.frx":2656C
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   9600
      Width           =   2175
   End
   Begin VB.PictureBox Picture1 
      Height          =   9015
      Left            =   1680
      Picture         =   "FZ1.frx":26F60
      ScaleHeight     =   8955
      ScaleWidth      =   11955
      TabIndex        =   1
      Top             =   600
      Width           =   12015
   End
   Begin VB.PictureBox Picture6 
      Height          =   4695
      Left            =   1200
      Picture         =   "FZ1.frx":55F6F
      ScaleHeight     =   4635
      ScaleWidth      =   5835
      TabIndex        =   12
      Top             =   4680
      Visible         =   0   'False
      Width           =   5895
   End
   Begin VB.PictureBox Picture8 
      Height          =   4695
      Left            =   7080
      Picture         =   "FZ1.frx":5A5C9
      ScaleHeight     =   4635
      ScaleWidth      =   6195
      TabIndex        =   13
      Top             =   4680
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture7 
      Height          =   4695
      Left            =   13320
      Picture         =   "FZ1.frx":613C2
      ScaleHeight     =   4635
      ScaleWidth      =   6075
      TabIndex        =   20
      Top             =   4680
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.PictureBox Picture3 
      Height          =   4815
      Left            =   13320
      Picture         =   "FZ1.frx":673B7
      ScaleHeight     =   4755
      ScaleWidth      =   6075
      TabIndex        =   9
      Top             =   -120
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.PictureBox Picture4 
      Height          =   4815
      Left            =   7080
      Picture         =   "FZ1.frx":6D5C8
      ScaleHeight     =   4755
      ScaleWidth      =   6195
      TabIndex        =   10
      Top             =   -120
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture2 
      Height          =   4815
      Left            =   1200
      Picture         =   "FZ1.frx":70B40
      ScaleHeight     =   4755
      ScaleWidth      =   5835
      TabIndex        =   8
      Top             =   -120
      Visible         =   0   'False
      Width           =   5895
      Begin VB.PictureBox Picture5 
         Height          =   15
         Left            =   720
         ScaleHeight     =   15
         ScaleWidth      =   1575
         TabIndex        =   11
         Top             =   4800
         Width           =   1575
      End
   End
   Begin VB.PictureBox Picture10 
      Height          =   8415
      Left            =   1800
      Picture         =   "FZ1.frx":770F3
      ScaleHeight     =   8355
      ScaleWidth      =   12675
      TabIndex        =   17
      Top             =   1200
      Visible         =   0   'False
      Width           =   12735
      Begin VB.CommandButton Command8 
         Caption         =   "RETURN"
         Height          =   615
         Left            =   5280
         TabIndex        =   18
         Top             =   7800
         Width           =   2295
      End
   End
   Begin VB.Image Image1 
      Height          =   2520
      Left            =   -1320
      Picture         =   "FZ1.frx":8ED60
      Top             =   -960
      Width           =   4500
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"FZ1.frx":8F6D2
      BeginProperty Font 
         Name            =   "Arial Narrow"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000B&
      Height          =   10695
      Left            =   13920
      TabIndex        =   4
      Top             =   240
      Visible         =   0   'False
      Width           =   6015
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"FZ1.frx":900D4
      BeginProperty Font 
         Name            =   "Algerian"
         Size            =   20.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000E&
      Height          =   11415
      Left            =   13800
      TabIndex        =   3
      Top             =   0
      Visible         =   0   'False
      Width           =   5655
   End
End
Attribute VB_Name = "Form21"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Command6.Visible = False
Command7.Visible = False
Label1.Visible = True
LABEL2.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture1.Visible = True
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
End Sub

Private Sub Command10_Click()
Picture1.Visible = True
Picture9.Visible = False
Form21.Show
End Sub

Private Sub Command2_Click()
Picture1.Visible = True
Command6.Visible = False
LABEL2.Visible = True
Command7.Visible = False
Label1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
End Sub

Private Sub Command4_Click()
Picture1.Visible = True
Command6.Visible = True
Command7.Visible = True
Label1.Visible = False
LABEL2.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
End Sub

Private Sub Command6_Click()
Picture10.Visible = True
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Label1.Visible = False
LABEL2.Visible = False
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
Picture9.Visible = False
End Sub

Private Sub Command7_Click()
Picture1.Visible = False
Picture9.Visible = True
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Label1.Visible = False
LABEL2.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
Picture8.Visible = False
Picture10.Visible = False
End Sub

Private Sub Command8_Click()
Picture10.Visible = False
Form21.Show
Picture1.Visible = True
End Sub

Private Sub Command9_Click()
Form28.Show
End Sub

Private Sub Command3_Click()
Picture1.Visible = False
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = True
Picture8.Visible = True
LABEL2.Visible = False
Label1.Visible = False
Command6.Visible = False
Command7.Visible = False
End Sub

Private Sub Command5_Click()
Form17.Show
End Sub

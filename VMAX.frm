VERSION 5.00
Begin VB.Form Form26 
   Caption         =   "Form26"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form26"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.PictureBox Picture7 
      Height          =   4335
      Left            =   12360
      Picture         =   "VMAX.frx":0000
      ScaleHeight     =   4275
      ScaleWidth      =   6195
      TabIndex        =   16
      Top             =   4920
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture6 
      Height          =   4335
      Left            =   6120
      Picture         =   "VMAX.frx":8D39
      ScaleHeight     =   4275
      ScaleWidth      =   6195
      TabIndex        =   15
      Top             =   4920
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture5 
      Height          =   4335
      Left            =   0
      Picture         =   "VMAX.frx":127A6
      ScaleHeight     =   4275
      ScaleWidth      =   6075
      TabIndex        =   14
      Top             =   4920
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.PictureBox Picture4 
      Height          =   5055
      Left            =   12360
      Picture         =   "VMAX.frx":1711A
      ScaleHeight     =   4995
      ScaleWidth      =   6195
      TabIndex        =   13
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture3 
      Height          =   5055
      Left            =   6120
      Picture         =   "VMAX.frx":1F5E8
      ScaleHeight     =   4995
      ScaleWidth      =   6195
      TabIndex        =   12
      Top             =   0
      Visible         =   0   'False
      Width           =   6255
   End
   Begin VB.PictureBox Picture2 
      Height          =   5055
      Left            =   0
      Picture         =   "VMAX.frx":27DC1
      ScaleHeight     =   4995
      ScaleWidth      =   6075
      TabIndex        =   11
      Top             =   0
      Visible         =   0   'False
      Width           =   6135
   End
   Begin VB.CommandButton Command7 
      Caption         =   "RETURN"
      Height          =   735
      Left            =   9960
      TabIndex        =   10
      Top             =   7440
      Visible         =   0   'False
      Width           =   1815
   End
   Begin VB.CommandButton Command6 
      Caption         =   "Main Menu"
      Height          =   855
      Left            =   11280
      TabIndex        =   6
      Top             =   9600
      Width           =   1695
   End
   Begin VB.CommandButton Command5 
      Caption         =   "Book"
      Height          =   855
      Left            =   9240
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton Command4 
      Height          =   855
      Left            =   7080
      Picture         =   "VMAX.frx":31644
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   9600
      Width           =   2175
   End
   Begin VB.CommandButton Command3 
      Height          =   855
      Left            =   5040
      Picture         =   "VMAX.frx":31D65
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton Command2 
      Height          =   855
      Left            =   3120
      Picture         =   "VMAX.frx":324AA
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   9600
      Width           =   2055
   End
   Begin VB.CommandButton Command1 
      Height          =   855
      Left            =   960
      Picture         =   "VMAX.frx":32BAE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   9600
      Width           =   2055
   End
   Begin VB.PictureBox Picture1 
      Height          =   9015
      Left            =   960
      Picture         =   "VMAX.frx":335A2
      ScaleHeight     =   8955
      ScaleWidth      =   11955
      TabIndex        =   0
      Top             =   600
      Width           =   12015
   End
   Begin VB.Image Image1 
      Height          =   495
      Left            =   720
      Top             =   10440
      Width           =   11895
   End
   Begin VB.Label Label3 
      Caption         =   "NO MORE COLOURS"
      BeginProperty Font 
         Name            =   "Kristen ITC"
         Size            =   72
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0000FF00&
      Height          =   5535
      Left            =   5400
      TabIndex        =   9
      Top             =   2640
      Visible         =   0   'False
      Width           =   10575
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   $"VMAX.frx":7B8E7
      BeginProperty Font 
         Name            =   "Century"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10935
      Left            =   12960
      TabIndex        =   8
      Top             =   0
      Visible         =   0   'False
      Width           =   7215
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   $"VMAX.frx":7C1F8
      BeginProperty Font 
         Name            =   "Century"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      Height          =   10575
      Left            =   12960
      TabIndex        =   7
      Top             =   240
      Visible         =   0   'False
      Width           =   7215
   End
End
Attribute VB_Name = "Form26"
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
Picture6.Visible = False
Picture7.Visible = False
End Sub

Private Sub Command2_Click()
Label1.Visible = False
Label2.Visible = True
Picture1.Visible = True
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
End Sub

Private Sub Command3_Click()
Picture2.Visible = True
Picture3.Visible = True
Picture4.Visible = True
Picture5.Visible = True
Picture6.Visible = True
Picture7.Visible = True
Label1.Visible = False
Label2.Visible = False
Picture1.Visible = False
End Sub

Private Sub Command4_Click()
Label1.Visible = False
Label2.Visible = False
Label3.Visible = True
Command1.Visible = False
Command2.Visible = False
Command3.Visible = False
Command4.Visible = False
Command5.Visible = False
Command7.Visible = True
Command6.Visible = False
Picture1.Visible = False
Picture2.Visible = False
Picture3.Visible = False
Picture4.Visible = False
Picture5.Visible = False
Picture6.Visible = False
Picture7.Visible = False
End Sub

Private Sub Command5_Click()
Form17.Show
End Sub

Private Sub Command6_Click()
Form28.Show
End Sub

Private Sub Command7_Click()
Picture1.Visible = False
Command1.Visible = True
Command2.Visible = True
Command3.Visible = True
Command4.Visible = True
Command5.Visible = True
Command6.Visible = True
Picture1.Visible = True
Command7.Visible = False
Label3.Visible = False
End Sub
